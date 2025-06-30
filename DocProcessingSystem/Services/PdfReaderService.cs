using System.Text.RegularExpressions;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using DocProcessingSystem.Core;
using Microsoft.Office.Interop.Word;

namespace DocProcessingSystem.Services
{
    /// <summary>
    /// Implementation of IPdfReader using iTextSharp
    /// </summary>
    public class PdfReaderService : IPdfReader
    {
        private bool _disposed = false;

        /// <summary>
        /// Checks if a specific page exists in the PDF
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="pageNumber">Page number to check</param>
        /// <returns>True if the page exists, false otherwise</returns>
        public bool PageExists(string pdfPath, int pageNumber)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath), "PDF path cannot be null or empty");

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            using (var reader = new PdfReader(pdfPath))
            {
                return pageNumber > 0 && pageNumber <= reader.NumberOfPages;
            }
        }

        public int GetPageCount(string pdfPath)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath), "PDF path cannot be null or empty");

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            using (var reader = new PdfReader(pdfPath))
            {
                return reader.NumberOfPages;
            }
        }

        /// <summary>
        /// Checks if a specific string exists in the PDF
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="searchText">Text to search for</param>
        /// <param name="isCaseSensitive">Whether the search should be case sensitive</param>
        /// <returns>True if the string exists, false otherwise</returns>
        public bool ContainsText(string pdfPath, string searchText, bool isCaseSensitive = false)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath), "PDF path cannot be null or empty");

            if (string.IsNullOrEmpty(searchText))
                throw new ArgumentNullException(nameof(searchText), "Search text cannot be null or empty");

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            using (var reader = new PdfReader(pdfPath))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    string pageText = PdfTextExtractor.GetTextFromPage(reader, i);

                    if (isCaseSensitive)
                    {
                        if (pageText.Contains(searchText))
                            return true;
                    }
                    else
                    {
                        if (pageText.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Checks if a specific string exists on a specific page of the PDF
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="searchText">Text to search for</param>
        /// <param name="pageNumber">Page number to search in</param>
        /// <param name="isCaseSensitive">Whether the search should be case sensitive</param>
        /// <returns>True if the string exists on the specified page, false otherwise</returns>
        public bool ContainsTextOnPage(string pdfPath, string searchText, int pageNumber, bool isCaseSensitive = false)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath), "PDF path cannot be null or empty");

            if (string.IsNullOrEmpty(searchText))
                throw new ArgumentNullException(nameof(searchText), "Search text cannot be null or empty");

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            using (var reader = new PdfReader(pdfPath))
            {
                if (pageNumber <= 0 || pageNumber > reader.NumberOfPages)
                    throw new ArgumentOutOfRangeException(nameof(pageNumber), $"Page number {pageNumber} is out of range. PDF has {reader.NumberOfPages} pages.");

                // Extract text from the specific page
                string pageText = PdfTextExtractor.GetTextFromPage(reader, pageNumber);

                // Check if the text exists on the page
                return isCaseSensitive
                    ? pageText.Contains(searchText)
                    : pageText.Contains(searchText, StringComparison.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// Searches for text in a PDF and returns the page numbers where it was found
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="searchText">Text to search for</param>
        /// <param name="isCaseSensitive">Whether the search should be case sensitive</param>
        /// <returns>List of page numbers where the text was found</returns>
        public List<int> FindTextInPages(string pdfPath, string searchText, bool isCaseSensitive = false)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath), "PDF path cannot be null or empty");

            if (string.IsNullOrEmpty(searchText))
                throw new ArgumentNullException(nameof(searchText), "Search text cannot be null or empty");

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            var pagesWithText = new List<int>();

            using (var reader = new PdfReader(pdfPath))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    string pageText = PdfTextExtractor.GetTextFromPage(reader, i);

                    bool found = isCaseSensitive
                        ? pageText.Contains(searchText)
                        : pageText.Contains(searchText, StringComparison.OrdinalIgnoreCase);

                    if (found)
                    {
                        pagesWithText.Add(i);
                    }
                }
            }

            return pagesWithText;
        }

        /// <summary>
        /// Extracts text from a specific page in the PDF
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="pageNumber">Page number to extract text from</param>
        /// <returns>Extracted text from the specified page</returns>
        public string ExtractTextFromPage(string pdfPath, int pageNumber)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath), "PDF path cannot be null or empty");

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            using (var reader = new PdfReader(pdfPath))
            {
                if (pageNumber <= 0 || pageNumber > reader.NumberOfPages)
                    throw new ArgumentOutOfRangeException(nameof(pageNumber), $"Page number {pageNumber} is out of range. PDF has {reader.NumberOfPages} pages.");

                return PdfTextExtractor.GetTextFromPage(reader, pageNumber);
            }
        }

        /// <summary>
        /// Extracts text from all pages in the PDF
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <returns>Dictionary with page numbers as keys and extracted text as values</returns>
        public Dictionary<int, string> ExtractTextFromAllPages(string pdfPath)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath), "PDF path cannot be null or empty");

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            var pageTextDictionary = new Dictionary<int, string>();

            using (var reader = new PdfReader(pdfPath))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    string pageText = PdfTextExtractor.GetTextFromPage(reader, i);
                    pageTextDictionary.Add(i, pageText);
                }
            }

            return pageTextDictionary;
        }

        /// <summary>
        /// Advanced search using regular expressions
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="pattern">Regular expression pattern to search for</param>
        /// <param name="options">Regular expression options</param>
        /// <returns>Dictionary with page numbers as keys and lists of matches as values</returns>
        public Dictionary<int, List<string>> SearchWithRegex(string pdfPath, string pattern, RegexOptions options = RegexOptions.None)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath), "PDF path cannot be null or empty");

            if (string.IsNullOrEmpty(pattern))
                throw new ArgumentNullException(nameof(pattern), "Regex pattern cannot be null or empty");

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            var results = new Dictionary<int, List<string>>();
            var regex = new Regex(pattern, options);

            using (var reader = new PdfReader(pdfPath))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    string pageText = PdfTextExtractor.GetTextFromPage(reader, i);

                    var matches = regex.Matches(pageText);

                    if (matches.Count > 0)
                    {
                        var matchList = new List<string>();
                        foreach (Match match in matches)
                        {
                            matchList.Add(match.Value);
                        }
                        results.Add(i, matchList);
                    }
                }
            }

            return results;
        }

        /// <summary>
        /// Extracts all bookmarks (outline) from a PDF
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <returns>Hierarchical list of bookmarks</returns>
        public List<Dictionary<string, object>> ExtractBookmarks(string pdfPath)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath), "PDF path cannot be null or empty");

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            using (var reader = new PdfReader(pdfPath))
            {
                var bookmarks = SimpleBookmark.GetBookmark(reader);
                return bookmarks?.Cast<Dictionary<string, object>>().ToList() ?? new List<Dictionary<string, object>>();
            }
        }

        /// <summary>
        /// Checks if a PDF has a specific section by looking for it in bookmarks
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="sectionName">Name of the section to look for</param>
        /// <param name="searchInContent">Whether to also search in the PDF content if not found in bookmarks</param>
        /// <returns>True if the section exists, false otherwise</returns>
        public bool HasSection(string pdfPath, string sectionName, bool searchInContent = true)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath), "PDF path cannot be null or empty");

            if (string.IsNullOrEmpty(sectionName))
                throw new ArgumentNullException(nameof(sectionName), "Section name cannot be null or empty");

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            using (var reader = new PdfReader(pdfPath))
            {
                // First try to find in bookmarks
                var bookmarks = SimpleBookmark.GetBookmark(reader);
                if (bookmarks != null && ContainsSectionInBookmarks(bookmarks, sectionName))
                {
                    return true;
                }

                // If not found in bookmarks and content search is enabled, search in content
                if (searchInContent)
                {
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        string pageText = PdfTextExtractor.GetTextFromPage(reader, i);
                        if (pageText.Contains(sectionName, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Gets metadata from the PDF document
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <returns>Dictionary containing metadata key-value pairs</returns>
        public Dictionary<string, string> GetMetadata(string pdfPath)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath), "PDF path cannot be null or empty");

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"PDF file not found: {pdfPath}");

            var metadata = new Dictionary<string, string>();

            using (var reader = new PdfReader(pdfPath))
            {
                var info = reader.Info;

                if (info != null)
                {
                    foreach (var key in info.Keys)
                    {
                        metadata[key] = info[key];
                    }
                }

                // Add some additional metadata
                metadata["PageCount"] = reader.NumberOfPages.ToString();
                metadata["FileSize"] = new FileInfo(pdfPath).Length.ToString();
                metadata["PDFVersion"] = reader.PdfVersion.ToString();

                // Check if the document is encrypted
                metadata["IsEncrypted"] = reader.IsEncrypted().ToString();
            }

            return metadata;
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