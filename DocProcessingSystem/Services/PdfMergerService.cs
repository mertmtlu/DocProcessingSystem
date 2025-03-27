using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using iTextSharp.text;
using iTextSharp.text.pdf;
using DocProcessingSystem.Core;
using iTextSharp.text.pdf.parser;

namespace DocProcessingSystem.Services
{
    /// <summary>
    /// Implementation of IPdfMerger using iTextSharp
    /// </summary>
    public class PdfMergerService : IPdfMerger
    {
        private bool _disposed = false;

        /// <summary>
        /// Merges multiple PDF files into one output file
        /// </summary>
        /// <param name="mainPdf">Path to the main PDF file</param>
        public void MergePdf(MergeSequence mergeSequence)
        {
            MergePdf(mergeSequence.MainDocument, mergeSequence.AdditionalDocuments, mergeSequence.OutputPath, mergeSequence.Options);
        }

        /// <summary>
        /// Merges multiple PDF files into one output file
        /// </summary>
        /// <param name="mainPdf">Path to the main PDF file</param>
        /// <param name="additionalPdfs">List of paths to additional PDF files to merge</param>
        /// <param name="outputPath">Path where the merged PDF will be saved</param>
        /// <param name="options">Merge options</param>
        /// <summary>
        /// Merges a main PDF with additional PDFs and writes to the output path
        /// </summary>
        /// <param name="mainPdf">Path to the main PDF document</param>
        /// <param name="additionalPdfs">List of paths to additional PDFs to append</param>
        /// <param name="outputPath">Output path for the merged PDF</param>
        /// <param name="options">Options for the merge operation</param>
        public void MergePdf(string mainPdf, List<string> additionalPdfs, string outputPath, MergeOptions options)
        {
            if (string.IsNullOrEmpty(mainPdf))
                throw new ArgumentNullException(nameof(mainPdf), "Main PDF path cannot be null or empty");
            if (additionalPdfs == null)
                throw new ArgumentNullException(nameof(additionalPdfs), "Additional PDFs list cannot be null");
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentNullException(nameof(outputPath), "Output path cannot be null or empty");
            if (options == null)
                throw new ArgumentNullException(nameof(options), "Merge options cannot be null");

            // Filter out any PDFs that match exclude patterns
            var filteredAdditionalPdfs = FilterPdfsByExcludePatterns(additionalPdfs, options.ExcludePatterns);

            // Handle the case where main PDF and output path are the same
            string tempMainPdf = mainPdf;
            bool usingTempFile = false;

            try
            {
                // Create a temporary file if mainPdf and outputPath are the same
                if (string.Equals(System.IO.Path.GetFullPath(mainPdf), System.IO.Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
                {
                    // Create a temporary copy of the main PDF
                    tempMainPdf = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        $"temp_{Guid.NewGuid()}_{System.IO.Path.GetFileName(mainPdf)}"
                    );

                    // Copy the main PDF to the temp location
                    File.Copy(mainPdf, tempMainPdf, true);
                    usingTempFile = true;

                    Console.WriteLine($"Created temporary copy of main PDF: {tempMainPdf}");
                }

                using (var outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    // Create a document object
                    var document = new Document();

                    // Create a PdfCopy object for the document
                    var pdfCopy = new PdfCopy(document, outputStream);

                    // Create a bookmark processor - always create it as we'll add our own bookmarks
                    var bookmarkProcessor = new BookmarkProcessor();

                    // Open the document for writing
                    document.Open();

                    // Process the main PDF first (using temp file if needed)
                    MergeSinglePdf(tempMainPdf, pdfCopy, bookmarkProcessor, 0);

                    // Process additional PDFs
                    int pageOffset = GetPageCount(tempMainPdf);
                    foreach (var pdfPath in filteredAdditionalPdfs)
                    {
                        // Create a bookmark for this additional PDF
                        string pdfFileName = System.IO.Path.GetFileNameWithoutExtension(pdfPath);
                        if (options.CreateBookmarksForAdditionalPdf)
                            bookmarkProcessor.AddFileBookmark(pdfFileName, pageOffset + 1);

                        MergeSinglePdf(pdfPath, pdfCopy, bookmarkProcessor, pageOffset);
                        pageOffset += GetPageCount(pdfPath);
                    }

                    // Add all bookmarks to the merged document
                    bookmarkProcessor.AddBookmarksToDocument(pdfCopy);

                    // Close the document
                    document.Close();
                }

                Console.WriteLine($"Successfully merged PDFs to: {outputPath}");
            }
            finally
            {
                // Clean up the temporary file if we created one
                if (usingTempFile && File.Exists(tempMainPdf))
                {
                    try
                    {
                        File.Delete(tempMainPdf);
                        Console.WriteLine($"Deleted temporary file: {tempMainPdf}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Failed to delete temporary file {tempMainPdf}: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Filters PDFs by exclude patterns
        /// </summary>
        private List<string> FilterPdfsByExcludePatterns(List<string> pdfs, List<string> excludePatterns)
        {
            if (excludePatterns == null || excludePatterns.Count == 0)
                return pdfs;

            return pdfs.Where(pdf =>
            {
                string fileName = System.IO.Path.GetFileName(pdf);
                return !excludePatterns.Any(pattern =>
                    Regex.IsMatch(fileName, WildcardToRegex(pattern), RegexOptions.IgnoreCase));
            }).ToList();
        }

        /// <summary>
        /// Converts wildcard pattern to regex
        /// </summary>
        private string WildcardToRegex(string pattern)
        {
            return "^" + Regex.Escape(pattern)
                .Replace("\\*", ".*")
                .Replace("\\?", ".") + "$";
        }

        /// <summary>
        /// Ensures all required sections are included in the PDFs
        /// </summary>
        private void EnsureRequiredSectionsIncluded(string mainPdf, List<string> additionalPdfs, string[] requiredSections)
        {
            if (requiredSections == null || requiredSections.Length == 0)
                return;

            var allPdfs = new List<string> { mainPdf };
            allPdfs.AddRange(additionalPdfs);

            List<string> missingSections = new List<string>();

            foreach (var section in requiredSections)
            {
                bool sectionFound = false;

                foreach (var pdf in allPdfs)
                {
                    if (PdfContainsSection(pdf, section))
                    {
                        sectionFound = true;
                        break;
                    }
                }

                if (!sectionFound)
                {
                    missingSections.Add(section);
                }
            }

            if (missingSections.Count > 0)
            {
                throw new InvalidOperationException($"Required sections not found: {string.Join(", ", missingSections)}");
            }
        }

        /// <summary>
        /// Checks if a PDF contains a specific section (looks in bookmarks and text)
        /// </summary>
        private bool PdfContainsSection(string pdfPath, string section)
        {
            using (var reader = new PdfReader(pdfPath))
            {
                // Check bookmarks for section name
                var bookmarks = SimpleBookmark.GetBookmark(reader);
                if (bookmarks != null && ContainsSectionInBookmarks(bookmarks, section))
                {
                    return true;
                }

                // Fallback to simple text search in the PDF
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    string pageText = PdfTextExtractor.GetTextFromPage(reader, i);
                    if (pageText.Contains(section, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Recursively checks if bookmarks contain a specific section
        /// </summary>
        private bool ContainsSectionInBookmarks(IList<Dictionary<string, object>> bookmarks, string section)
        {
            foreach (var bookmark in bookmarks)
            {
                if (bookmark.TryGetValue("Title", out object title) &&
                    title.ToString().Contains(section, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (bookmark.TryGetValue("Kids", out object kids) &&
                    kids is IList<Dictionary<string, object>> childBookmarks)
                {
                    if (ContainsSectionInBookmarks(childBookmarks, section))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Merges a single PDF into the output document
        /// </summary>
        private void MergeSinglePdf(string pdfPath, PdfCopy pdfCopy, BookmarkProcessor bookmarkProcessor, int pageOffset)
        {
            using (var reader = new PdfReader(pdfPath))
            {
                // Save bookmarks if needed (only if preserveBookmarks is true)
                var bookmarks = SimpleBookmark.GetBookmark(reader);
                if (bookmarks != null)
                {
                    // Adjust page numbers in bookmarks to account for offset in the merged document
                    SimpleBookmark.ShiftPageNumbers(bookmarks, pageOffset, null);
                    bookmarkProcessor.AddBookmarks(bookmarks);
                }

                // Add all pages from this PDF to the output
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    pdfCopy.AddPage(pdfCopy.GetImportedPage(reader, i));
                }

                // Copy any interactive form fields if present
                PdfReader.unethicalreading = true; // Allow reading of protected PDFs

                // Note: PdfCopy automatically includes form fields when importing pages
                // No explicit CopyAcroForm call is needed with recent iTextSharp versions
            }
        }

        /// <summary>
        /// Gets the number of pages in a PDF
        /// </summary>
        private int GetPageCount(string pdfPath)
        {
            using (var reader = new PdfReader(pdfPath))
            {
                return reader.NumberOfPages;
            }
        }

        /// <summary>
        /// Helper class to process and merge bookmarks
        /// </summary>
        private class BookmarkProcessor
        {
            private readonly List<Dictionary<string, object>> _allBookmarks = new List<Dictionary<string, object>>();

            /// <summary>
            /// Adds bookmarks from a PDF to the collection
            /// </summary>
            public void AddBookmarks(IList<Dictionary<string, object>> bookmarks)
            {
                if (bookmarks != null && bookmarks.Count > 0)
                {
                    _allBookmarks.AddRange(bookmarks);
                }
            }

            /// <summary>
            /// Creates and adds a new bookmark for a merged file
            /// </summary>
            /// <param name="title">Bookmark title (PDF file name)</param>
            /// <param name="pageNumber">Page number where the PDF starts</param>
            public void AddFileBookmark(string title, int pageNumber)
            {
                var bookmark = new Dictionary<string, object>
                {
                    { "Title", title },
                    { "Action", "GoTo" },
                    { "Page", $"{pageNumber} Fit" } // "Fit" makes the page fit in the viewer
                };

                _allBookmarks.Add(bookmark);
            }

            /// <summary>
            /// Adds all collected bookmarks to the final document
            /// </summary>
            public void AddBookmarksToDocument(PdfCopy pdfCopy)
            {
                if (_allBookmarks.Count > 0)
                {
                    pdfCopy.Outlines = _allBookmarks;
                }
            }
        }

        /// <summary>
        /// Disposes resources
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Disposes resources
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources if any
                }

                _disposed = true;
            }
        }
    }
}