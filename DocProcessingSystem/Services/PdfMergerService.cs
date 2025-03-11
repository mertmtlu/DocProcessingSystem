using DocProcessingSystem.Core;
using iText.Kernel.Pdf.Navigation;
using iText.Kernel.Pdf;

namespace DocProcessingSystem.Services
{
    /// <summary>
    /// Service for merging PDF documents
    /// </summary>
    public class PdfMergerService : IPdfMerger
    {
        private bool _disposed;

        /// <summary>
        /// Merges multiple PDF files into one output file
        /// </summary>
        public void MergePdf(string mainPdf, List<string> additionalPdfs, string outputPath, MergeOptions options)
        {
            if (!File.Exists(mainPdf))
                throw new FileNotFoundException($"Main PDF file not found: {mainPdf}");

            var mainFolder = Path.GetDirectoryName(mainPdf);

            // Check if the combined EK-B,C.pdf exists
            string combinedPath = Path.Combine(mainFolder, "EK-B,C.pdf");
            if (File.Exists(combinedPath))
            {
                Console.WriteLine("Found combined EK-B,C.pdf file");

                // Find the positions of individual EK-B and EK-C files if they exist
                int ekbIndex = additionalPdfs.FindIndex(path =>
                    Path.GetFileNameWithoutExtension(path).Equals("EK-B", StringComparison.OrdinalIgnoreCase));
                int ekcIndex = additionalPdfs.FindIndex(path =>
                    Path.GetFileNameWithoutExtension(path).Equals("EK-C", StringComparison.OrdinalIgnoreCase));

                // Determine insertion point (use the first occurrence of either file)
                int insertPosition = -1;
                if (ekbIndex >= 0 && ekcIndex >= 0)
                    insertPosition = Math.Min(ekbIndex, ekcIndex);
                else if (ekbIndex >= 0)
                    insertPosition = ekbIndex;
                else if (ekcIndex >= 0)
                    insertPosition = ekcIndex;

                // Remove individual files if they exist
                if (ekbIndex >= 0)
                    additionalPdfs.RemoveAt(ekbIndex);
                // If we removed EK-B and it came before EK-C, adjust EK-C's index
                if (ekcIndex >= 0 && ekbIndex >= 0 && ekbIndex < ekcIndex)
                    ekcIndex--;
                if (ekcIndex >= 0)
                    additionalPdfs.RemoveAt(ekcIndex);

                // Insert the combined file at the appropriate position or add at the end
                if (insertPosition >= 0 && !additionalPdfs.Contains(combinedPath)) additionalPdfs.Insert(insertPosition, combinedPath);
                else if (!additionalPdfs.Contains(combinedPath)) throw new Exception("EK-B,C cannot be managed.");
            }

            // Check for required sections
            foreach (var required in options.RequiredSections)
            {
                // Check if any path in additionalPdfs contains the required filename
                bool foundFile = additionalPdfs.Any(path =>
                    Path.GetFileNameWithoutExtension(path).Equals(required, StringComparison.OrdinalIgnoreCase));

                if (!foundFile)
                    throw new FileNotFoundException($"Additional PDF file not found (required): {required}");
            }

            foreach (var filePath in additionalPdfs)
                // Check if any path in additionalPdfs contains the required filename
                if (!File.Exists(filePath))
                    throw new FileNotFoundException($"Additional PDF file not found: {filePath}");

            // Ensure the output directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            PdfWriter writer = null;
            PdfDocument resultDoc = null;
            PdfDocument mainDoc = null;

            try
            {
                // Initialize the documents
                writer = new PdfWriter(outputPath);
                resultDoc = new PdfDocument(writer);
                mainDoc = new PdfDocument(new PdfReader(mainPdf));

                // Copy pages from main document
                Console.WriteLine("Copying main document pages...");
                mainDoc.CopyPagesTo(1, mainDoc.GetNumberOfPages(), resultDoc);

                // Safely copy outlines if requested
                if (options.PreserveBookmarks)
                {
                    Console.WriteLine("Copying document structure...");
                    try
                    {
                        PdfOutline sourceOutline = mainDoc.GetOutlines(false);
                        if (sourceOutline != null)
                        {
                            PdfOutline destOutline = resultDoc.GetOutlines(false);
                            SafelyCopyOutlines(sourceOutline, destOutline, resultDoc);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not copy outlines: {ex.Message}");
                    }
                }

                // Process and append additional PDFs
                foreach (string pdfFile in additionalPdfs.Where(File.Exists))
                {
                    // Skip excluded files
                    if (options.ExcludePatterns?.Any(pattern =>
                        Path.GetFileName(pdfFile).Contains(pattern, StringComparison.OrdinalIgnoreCase)) == true)
                    {
                        Console.WriteLine($"Skipping excluded file: {Path.GetFileName(pdfFile)}");
                        continue;
                    }

                    try
                    {
                        Console.WriteLine($"Adding: {Path.GetFileName(pdfFile)}");
                        using (var additionalDoc = new PdfDocument(new PdfReader(pdfFile)))
                            additionalDoc.CopyPagesTo(1, additionalDoc.GetNumberOfPages(), resultDoc);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not copy {Path.GetFileName(pdfFile)}: {ex.Message}");
                    }
                }

                // Close documents in reverse order
                Console.WriteLine("Finalizing document...");
                if (mainDoc != null) mainDoc.Close();
                if (resultDoc != null) resultDoc.Close();
                if (writer != null) writer.Close();

                Console.WriteLine("PDF merge completed successfully.\n");
            }
            catch (Exception ex)
            {
                throw new Exception($"Error merging PDFs: {ex.Message}", ex);
            }
            finally
            {
                try { if (mainDoc != null && !mainDoc.IsClosed()) mainDoc.Close(); } catch { }
                try { if (resultDoc != null && !resultDoc.IsClosed()) resultDoc.Close(); } catch { }
                try { if (writer != null) writer.Close(); } catch { }
            }

        }

        /// <summary>
        /// Safely copies PDF outlines (bookmarks) from source to target
        /// </summary>
        public void SafelyCopyOutlines(PdfOutline source, PdfOutline target, PdfDocument targetDoc)
        {
            if (source == null || target == null) return;

            foreach (PdfOutline child in source.GetAllChildren())
            {
                if (child == null) continue;

                try
                {
                    string title = child.GetTitle();
                    if (string.IsNullOrEmpty(title)) title = "Untitled Bookmark";

                    PdfOutline newChild = target.AddOutline(title);
                    if (newChild == null) continue;

                    try
                    {
                        if (child.GetDestination() != null)
                        {
                            PdfExplicitDestination newDest = PdfExplicitDestination.CreateFit(
                                targetDoc.GetPage(1)
                            );
                            newChild.AddDestination(newDest);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not copy destination for outline {title}: {ex.Message}");
                    }

                    SafelyCopyOutlines(child, newChild, targetDoc);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not copy outline structure: {ex.Message}");
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
                _disposed = true;
            }
        }
    }
}
