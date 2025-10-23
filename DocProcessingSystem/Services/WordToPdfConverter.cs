using DocProcessingSystem.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

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
            string outputDirectory = Path.GetDirectoryName(outputPath);
            Directory.CreateDirectory(outputDirectory);

            Document doc = null;
            try
            {
                if (Path.GetFileName(inputPath).Contains("~$"))
                {
                    Console.WriteLine($"Warning: Passed: {Path.GetFileName(inputPath)}");
                    return;
                }

                //Console.WriteLine($"Converting {Path.GetFileName(inputPath)} to PDF");
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
                string originalFileName = Path.GetFileName(inputPath);
                string destinationPath = Path.Combine(outputDirectory, originalFileName);

                // Don't copy if source and destination are the same
                if (!string.Equals(inputPath, destinationPath, StringComparison.OrdinalIgnoreCase) && copyWord)
                {
                    File.Copy(inputPath, destinationPath.Replace(".docx", "_nt.docx"), true); // 'true' to overwrite if file exists
                    //Console.WriteLine($"Original file copied to: {destinationPath}");
                }

                //Console.WriteLine($"Successfully converted: {Path.GetFileName(inputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting {Path.GetFileName(inputPath)}: {ex.Message}");
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
