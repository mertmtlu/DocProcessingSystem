using DocProcessingSystem.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Xobject;
using IOPath = System.IO.Path;
using PdfRectangle = iText.Kernel.Geom.Rectangle;

namespace DocProcessingSystem.Services
{
    /// <summary>
    /// Converts Word documents to PDF format
    /// </summary>
    public class WordToPdfConverter : IDocumentProcessor
    {
        private Application _wordApp;
        private bool _disposed;

        private EventHandler _processExitHandler;
        private UnhandledExceptionEventHandler _unhandledExceptionHandler;

        public WordToPdfConverter()
        {
            _wordApp = new Application();
            _wordApp.Visible = false;
            _wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            _wordApp.Options.PrintDraft = false;

            _processExitHandler = (object? sender, EventArgs e) => Dispose();
            _unhandledExceptionHandler = (object sender, UnhandledExceptionEventArgs e) => Dispose();

            AppDomain.CurrentDomain.ProcessExit += _processExitHandler;
            AppDomain.CurrentDomain.UnhandledException += _unhandledExceptionHandler;
        }

        /// <summary>
        /// Converts a Word document to PDF format and copies the original file to the output location
        /// </summary>
        public void Convert(string inputPath, string outputPath, bool saveWordChanges, bool copyWord = true)
        {
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Ensure the output directory exists
            string outputDirectory = IOPath.GetDirectoryName(outputPath);
            Directory.CreateDirectory(outputDirectory);

            Document doc = null;
            try
            {
                if (IOPath.GetFileName(inputPath).Contains("~$"))
                {
                    Console.WriteLine($"Warning: Passed: {IOPath.GetFileName(inputPath)}");
                    return;
                }

                //Console.WriteLine($"Converting {IOPath.GetFileName(inputPath)} to PDF");
                doc = _wordApp.Documents.Open(inputPath);

                // Remove background from all pages
                RemoveBackgrounds(doc);

                doc.ExportAsFixedFormat(
                    OutputFileName: outputPath,
                    ExportFormat: WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport: false,
                    OptimizeFor: WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Range: WdExportRange.wdExportAllDocument,
                    From: 0,
                    To: 0,
                    Item: WdExportItem.wdExportDocumentContent,
                    IncludeDocProps: true,
                    KeepIRM: true,
                    CreateBookmarks: WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                    DocStructureTags: true,
                    BitmapMissingFonts: true,
                    UseISO19005_1: false
                );

                if (saveWordChanges) doc.Save();

                // Copy the original file to the output location
                string originalFileName = IOPath.GetFileName(inputPath);
                string destinationPath = IOPath.Combine(outputDirectory, originalFileName);

                // Don't copy if source and destination are the same
                if (!string.Equals(inputPath, destinationPath, StringComparison.OrdinalIgnoreCase) && copyWord)
                {
                    File.Copy(inputPath, destinationPath.Replace(".docx", "_nt.docx"), true);
                    //Console.WriteLine($"Original file copied to: {destinationPath}");
                }

                //Console.WriteLine($"Successfully converted: {IOPath.GetFileName(inputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting {IOPath.GetFileName(inputPath)}: {ex.Message}");
                throw;
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.ReleaseComObject(doc);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            // Scale the exported PDF to A4
            if (outputPath.Contains("ön_kapak"))
            {
                ScalePdfToA4(outputPath);
            }
        }

        /// <summary>
        /// Scales a PDF to A4 size (stretch to fill)
        /// </summary>
        private void ScalePdfToA4(string pdfPath)
        {
            try
            {
                string tempPath = pdfPath + ".tmp";
                using (PdfDocument sourcePdf = new PdfDocument(new PdfReader(pdfPath)))
                using (PdfDocument destPdf = new PdfDocument(new PdfWriter(tempPath)))
                {
                    // A4 dimensions in points (210mm x 297mm)
                    float a4Width = 595.276f;   // 595.276
                    float a4Height = 841.89f; // 841.89
                    PageSize a4 = new PageSize(a4Width, a4Height);
                    Console.WriteLine($"Page Width: {a4Width}, Height: {a4Height}");

                    for (int i = 1; i <= sourcePdf.GetNumberOfPages(); i++)
                    {
                        PdfPage sourcePage = sourcePdf.GetPage(i);
                        float a5Width = 419.3158552381f;
                        float a5Height = 594.9752f;
                        PdfRectangle sourceRect = new(a5Width, a5Height);
                        Console.WriteLine($"Rect Width: {sourceRect.GetWidth()}, Height: {sourceRect.GetHeight()}");

                        // Create new page with explicit A4 dimensions
                        PdfPage destPage = destPdf.AddNewPage(a4);

                        // EXPLICITLY set the MediaBox to A4
                        destPage.SetMediaBox(new PdfRectangle(0, 0, a4Width, a4Height));
                        destPage.SetCropBox(new PdfRectangle(0, 0, a4Width, a4Height));

                        // Calculate scale factors
                        float scaleX = a4Width / sourceRect.GetWidth();
                        float scaleY = a4Height / sourceRect.GetHeight();

                        // Copy as XObject
                        PdfFormXObject xObject = sourcePage.CopyAsFormXObject(destPdf);

                        // Draw on A4 page with scaling
                        var canvas = new PdfCanvas(destPage);
                        canvas.SaveState();
                        canvas.ConcatMatrix(scaleX, 0, 0, scaleY, 0, 0);
                        canvas.AddXObjectAt(xObject, 0, 0);
                        canvas.RestoreState();
                    }
                }

                File.Delete(pdfPath);
                File.Move(tempPath, pdfPath);

                Console.WriteLine($"PDF converted to A4 format");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not scale PDF to A4: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes backgrounds and highlighting from Word document
        /// </summary>
        public void RemoveBackgrounds(Document doc)
        {
            try
            {
                // Remove highlights from all text
                foreach (Microsoft.Office.Interop.Word.Range range in doc.StoryRanges)
                {
                    try
                    {
                        range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                    }
                    catch { } // Skip if range can't be modified

                    // If there are tables, remove highlighting from table cells
                    foreach (Table table in range.Tables)
                    {
                        try
                        {
                            foreach (Row row in table.Rows)
                            {
                                foreach (Cell cell in row.Cells)
                                {
                                    cell.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                                }
                            }
                        }
                        catch { } // Skip if table can't be modified
                    }
                }

                // Remove any shape fills and backgrounds
                foreach (Microsoft.Office.Interop.Word.Shape shape in doc.Shapes)
                {
                    try
                    {
                        if (shape.Type == MsoShapeType.msoTextBox ||
                            shape.Type == MsoShapeType.msoPicture)
                        {
                            shape.Fill.Visible = MsoTriState.msoFalse;
                        }
                    }
                    catch { } // Skip if shape can't be modified
                }

                // Try to remove document background
                try
                {
                    if (doc.Background != null)
                    {
                        doc.Background.Fill.Visible = MsoTriState.msoFalse;
                    }
                }
                catch { } // Skip if background can't be modified
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not remove all backgrounds: {ex.Message}");
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
                    // Unsubscribe from events
                    AppDomain.CurrentDomain.ProcessExit -= _processExitHandler;
                    AppDomain.CurrentDomain.UnhandledException -= _unhandledExceptionHandler;

                    // Dispose Word application
                    if (_wordApp != null)
                    {
                        try
                        {
                            _wordApp.Quit(WdSaveOptions.wdDoNotSaveChanges);
                            Marshal.ReleaseComObject(_wordApp);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error disposing Word application: {ex.Message}");
                        }
                        finally
                        {
                            _wordApp = null;
                        }
                    }
                }
                _disposed = true;
            }
        }
    }
}