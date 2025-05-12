using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Utils;
using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Pdf.Canvas;
using iText.IO.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Colors;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Filter;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;

namespace DocProcessingSystem.Services
{
    /// <summary>
    /// Options for text replacement in PDF documents
    /// </summary>
    public class TextReplacementOptions
    {
        /// <summary>
        /// Whether the search should be case sensitive
        /// </summary>
        public bool IsCaseSensitive { get; set; } = false;

        /// <summary>
        /// Whether to replace all occurrences or just the first one
        /// </summary>
        public bool ReplaceAllOccurrences { get; set; } = true;

        /// <summary>
        /// Specific pages to perform replacement on (empty means all pages)
        /// </summary>
        public List<int> SpecificPages { get; set; } = new List<int>();

        /// <summary>
        /// Whether to apply a visible highlight to replaced text (for debugging)
        /// </summary>
        public bool HighlightReplacedText { get; set; } = false;

        /// <summary>
        /// Color to use for highlighting (if enabled)
        /// </summary>
        public Color HighlightColor { get; set; } = ColorConstants.YELLOW;

        /// <summary>
        /// Path to the Calibri-Bold font file
        /// </summary>
        public string CalibriBoldFontPath { get; set; } = "C:\\Windows\\Fonts\\calibri.ttf";

        /// <summary>
        /// Font size to use for replaced text
        /// </summary>
        public float ReplacementFontSize { get; set; } = 7.25f;
    }

    /// <summary>
    /// Service for replacing text in PDF documents
    /// </summary>
    public class PdfTextReplacerService
    {
        private bool _disposed = false;

        /// <summary>
        /// Replaces text in a PDF document
        /// </summary>
        /// <param name="inputPdfPath">Path to the input PDF</param>
        /// <param name="outputPdfPath">Path where the modified PDF should be saved</param>
        /// <param name="oldText">Text to be replaced</param>
        /// <param name="newText">Text to replace with</param>
        /// <param name="options">Options for controlling the replacement</param>
        public void ReplaceText(string inputPdfPath, string outputPdfPath, string oldText, string newText, TextReplacementOptions options = null)
        {
            if (string.IsNullOrEmpty(inputPdfPath))
                throw new ArgumentNullException(nameof(inputPdfPath), "Input PDF path cannot be null or empty");

            if (string.IsNullOrEmpty(outputPdfPath))
                throw new ArgumentNullException(nameof(outputPdfPath), "Output PDF path cannot be null or empty");

            if (string.IsNullOrEmpty(oldText))
                throw new ArgumentNullException(nameof(oldText), "Text to replace cannot be null or empty");

            if (newText == null) // Allow empty replacement string
                newText = string.Empty;

            if (!File.Exists(inputPdfPath))
                throw new FileNotFoundException($"PDF file not found: {inputPdfPath}");

            // Create default options if none provided
            options ??= new TextReplacementOptions();

            // Ensure output directory exists
            string outputDir = System.IO.Path.GetDirectoryName(outputPdfPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Use a direct replacement approach that's better for tables
            try
            {
                using (var reader = new PdfReader(inputPdfPath))
                using (var writer = new PdfWriter(outputPdfPath))
                using (var pdfDocument = new PdfDocument(reader, writer))
                {
                    int replacementCount = ProcessTextReplacement(pdfDocument, oldText, newText, options);
                    if (replacementCount == 0) Console.WriteLine($"PDF: {inputPdfPath} completed with {replacementCount} replacements");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during PDF text replacement: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Performs specific text replacement from "Çap/Yükseklik" to "Yükseklik/Çap"
        /// </summary>
        /// <param name="inputPdfPath">Path to the input PDF</param>
        /// <param name="outputPdfPath">Path where the modified PDF should be saved</param>
        /// <param name="options">Options for controlling the replacement</param>
        public void ReplaceCapYukseklik(string inputPdfPath, string outputPdfPath, TextReplacementOptions options = null)
        {
            options ??= new TextReplacementOptions();

            var reader = new PdfReaderService();

            if (reader.ContainsText(inputPdfPath, "Numuneni"))
            {
                options.CalibriBoldFontPath = "C:\\Windows\\Fonts\\calibrib.ttf";
                ReplaceText(inputPdfPath, outputPdfPath, "Numuneni Çap/Yükseklik", "Numune Yükseklik/Çap", options);
            }
        }

        /// <summary>
        /// Process text replacement throughout the document
        /// </summary>
        private int ProcessTextReplacement(PdfDocument pdfDocument, string oldText, string newText, TextReplacementOptions options)
        {
            int totalReplacements = 0;
            int totalPages = pdfDocument.GetNumberOfPages();

            StringComparison comparisonType = options.IsCaseSensitive
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;

            // Process each page
            for (int pageNum = 1; pageNum <= totalPages; pageNum++)
            {
                // Skip this page if not in specific pages list (when the list is not empty)
                if (options.SpecificPages != null && options.SpecificPages.Count > 0 && !options.SpecificPages.Contains(pageNum))
                    continue;

                PdfPage page = pdfDocument.GetPage(pageNum);

                // Find text locations on this page
                LocationTextReplacer textReplacer = new LocationTextReplacer(oldText, newText, comparisonType);
                PdfCanvasProcessor processor = new PdfCanvasProcessor(textReplacer);
                processor.ProcessPageContent(page);

                // Get the replacements that were identified
                List<TextReplacement> replacements = textReplacer.GetReplacements();

                if (replacements.Count > 0)
                {
                    //Console.WriteLine($"Found {replacements.Count} occurrences on page {pageNum}");

                    // Apply the replacements with a new content stream
                    PdfCanvas canvas = new PdfCanvas(page);

                    foreach (var replacement in replacements)
                    {
                        // Apply a single replacement
                        ApplyReplacement(canvas, replacement, options);
                        totalReplacements++;

                        // Stop after first replacement if not replacing all
                        if (!options.ReplaceAllOccurrences)
                            break;
                    }

                    // Stop after first page with replacements if not replacing all
                    if (!options.ReplaceAllOccurrences && totalReplacements > 0)
                        break;
                }
            }

            return totalReplacements;
        }

        /// <summary>
        /// Apply a single text replacement
        /// </summary>
        private void ApplyReplacement(PdfCanvas canvas, TextReplacement replacement, TextReplacementOptions options)
        {
            // Save state before modifying
            canvas.SaveState();

            try
            {
                // Optional: Highlight the replacement area for debugging
                if (options.HighlightReplacedText)
                {
                    canvas.SetFillColor(options.HighlightColor);
                    canvas.Rectangle(
                        replacement.BoundingBox.GetLeft(),
                        replacement.BoundingBox.GetBottom(),
                        replacement.BoundingBox.GetWidth(),
                        replacement.BoundingBox.GetHeight());
                    canvas.Fill();
                }

                // Erase the original text with a white rectangle
                canvas.SetFillColor(ColorConstants.WHITE);
                canvas.Rectangle(
                    replacement.BoundingBox.GetLeft() - 1,
                    replacement.BoundingBox.GetBottom() - 1,
                    replacement.BoundingBox.GetWidth() + 2,
                    replacement.BoundingBox.GetHeight() + 0.8);
                canvas.Fill();

                // Draw the new text
                canvas.BeginText();

                // Use Calibri-Bold font with exact size of 7.66
                PdfFont font;
                try
                {
                    // Check if font file exists
                    if (!File.Exists(options.CalibriBoldFontPath))
                    {
                        Console.WriteLine($"Warning: Calibri Bold font file not found at {options.CalibriBoldFontPath}. Falling back to Helvetica-Bold.");
                        font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    }
                    else
                    {
                        // Load Calibri-Bold font from file - using the correct overload
                        // This is compatible with iText7
                        FontProgram fontProgram = FontProgramFactory.CreateFont(options.CalibriBoldFontPath);
                        font = PdfFontFactory.CreateFont(fontProgram, PdfEncodings.IDENTITY_H);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading Calibri-Bold font: {ex.Message}. Falling back to Helvetica-Bold.");
                    font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                }

                // Set font and size to exactly 7.66
                canvas.SetFontAndSize(font, options.ReplacementFontSize);

                // Set text rendering mode and colors
                canvas.SetFillColor(replacement.FontDetails.TextColor);

                // Position text at the same baseline
                canvas.MoveText(
                    replacement.BoundingBox.GetLeft(),
                    replacement.BoundingBox.GetBottom() + 2); // Small adjustment for baseline

                // Write the new text
                canvas.ShowText(replacement.NewText);
                canvas.EndText();
            }
            finally
            {
                // Restore state after modification
                canvas.RestoreState();
            }
        }

        /// <summary>
        /// Custom text extraction and replacement strategy
        /// </summary>
        private class LocationTextReplacer : IEventListener
        {
            private readonly string _searchText;
            private readonly string _replacementText;
            private readonly StringComparison _comparisonType;
            private readonly List<TextReplacement> _replacements = new List<TextReplacement>();

            // List to store text chunks for analysis
            private readonly List<TextChunkLocation> _textChunks = new List<TextChunkLocation>();

            public LocationTextReplacer(string searchText, string replacementText, StringComparison comparisonType)
            {
                _searchText = searchText;
                _replacementText = replacementText;
                _comparisonType = comparisonType;
            }

            public void EventOccurred(IEventData data, EventType type)
            {
                if (type == EventType.RENDER_TEXT)
                {
                    var renderInfo = (TextRenderInfo)data;
                    string text = renderInfo.GetText();

                    if (!string.IsNullOrEmpty(text))
                    {
                        // Create a new text chunk for this render event
                        var chunk = new TextChunkLocation(
                            renderInfo.GetText(),
                            renderInfo.GetDescentLine().GetStartPoint(),
                            renderInfo.GetAscentLine().GetEndPoint(),
                            renderInfo.GetBaseline().GetStartPoint(),
                            renderInfo.GetBaseline().GetEndPoint()
                        );

                        // Store the font details
                        chunk.FontDetails = new FontDetails
                        {
                            FontName = renderInfo.GetFont().GetFontProgram().ToString(),
                            FontSize = renderInfo.GetFontSize(),
                            TextColor = renderInfo.GetFillColor()
                        };

                        // Add to our chunks collection
                        _textChunks.Add(chunk);
                    }
                }
            }

            public ICollection<EventType> GetSupportedEvents()
            {
                return new HashSet<EventType>() { EventType.RENDER_TEXT };
            }

            /// <summary>
            /// After processing the page, analyze text chunks to find replacements
            /// </summary>
            public List<TextReplacement> GetReplacements()
            {
                // Try to group chunks by lines first
                var lines = GroupChunksByLines(_textChunks);

                // Process each line to look for the search text
                foreach (var line in lines)
                {
                    // Sort chunks from left to right
                    var sortedChunks = line.OrderBy(chunk => chunk.StartLocation.Get(0)).ToList();

                    // Combine text from all chunks in this line
                    StringBuilder lineText = new StringBuilder();
                    foreach (var chunk in sortedChunks)
                    {
                        lineText.Append(chunk.Text);
                    }

                    // Look for search text in the combined text
                    string fullLineText = lineText.ToString();
                    int pos = 0;
                    while ((pos = fullLineText.IndexOf(_searchText, pos, _comparisonType)) != -1)
                    {
                        // Found a match - now calculate which chunks are involved

                        // Find the chunks that contain the text
                        List<TextChunkLocation> matchChunks = new List<TextChunkLocation>();
                        int currentPos = 0;
                        int matchEndPos = pos + _searchText.Length;

                        for (int i = 0; i < sortedChunks.Count; i++)
                        {
                            int chunkStart = currentPos;
                            int chunkEnd = chunkStart + sortedChunks[i].Text.Length;

                            // Check if this chunk overlaps with our match
                            if (chunkEnd > pos && chunkStart < matchEndPos)
                            {
                                matchChunks.Add(sortedChunks[i]);
                            }

                            currentPos = chunkEnd;

                            // If we've passed the end of the match, we can stop
                            if (currentPos >= matchEndPos)
                                break;
                        }

                        // If we have chunks that match, create a replacement
                        if (matchChunks.Count > 0)
                        {
                            // Calculate a bounding box that covers all matched chunks
                            Rectangle boundingBox = GetBoundingBox(matchChunks);

                            // Create a replacement using information from the first chunk
                            TextReplacement replacement = new TextReplacement
                            {
                                OriginalText = _searchText,
                                NewText = _replacementText,
                                BoundingBox = boundingBox,
                                FontDetails = matchChunks[0].FontDetails
                            };

                            _replacements.Add(replacement);
                        }

                        // Move past this match
                        pos += _searchText.Length;
                    }
                }

                return _replacements;
            }

            /// <summary>
            /// Group text chunks into lines based on vertical position
            /// </summary>
            private List<List<TextChunkLocation>> GroupChunksByLines(List<TextChunkLocation> chunks)
            {
                // Sort chunks by their vertical position (top to bottom)
                var sortedChunks = chunks.OrderByDescending(c => c.StartLocation.Get(1)).ToList();

                List<List<TextChunkLocation>> lines = new List<List<TextChunkLocation>>();
                List<TextChunkLocation> currentLine = new List<TextChunkLocation>();

                if (sortedChunks.Count == 0)
                    return lines;

                // Use the first chunk to start the first line
                float lastY = sortedChunks[0].StartLocation.Get(1);
                currentLine.Add(sortedChunks[0]);

                // Group remaining chunks into lines based on vertical position
                for (int i = 1; i < sortedChunks.Count; i++)
                {
                    float currentY = sortedChunks[i].StartLocation.Get(1);

                    // If this chunk is on the same line (within a small tolerance)
                    if (Math.Abs(currentY - lastY) < 5)
                    {
                        currentLine.Add(sortedChunks[i]);
                    }
                    else
                    {
                        // This chunk is on a new line
                        lines.Add(currentLine);
                        currentLine = new List<TextChunkLocation> { sortedChunks[i] };
                        lastY = currentY;
                    }
                }

                // Add the last line if it has chunks
                if (currentLine.Count > 0)
                {
                    lines.Add(currentLine);
                }

                return lines;
            }

            /// <summary>
            /// Calculate a bounding box that covers all chunks
            /// </summary>
            private Rectangle GetBoundingBox(List<TextChunkLocation> chunks)
            {
                if (chunks.Count == 0)
                    return new Rectangle(0, 0, 0, 0);

                float minX = float.MaxValue;
                float minY = float.MaxValue;
                float maxX = float.MinValue;
                float maxY = float.MinValue;

                foreach (var chunk in chunks)
                {
                    // Get the chunk's coordinates
                    float chunkMinX = Math.Min(chunk.StartLocation.Get(0), chunk.EndLocation.Get(0));
                    float chunkMaxX = Math.Max(chunk.StartLocation.Get(0), chunk.EndLocation.Get(0));
                    float chunkMinY = Math.Min(chunk.DescendLocation.Get(1), chunk.AscentLocation.Get(1));
                    float chunkMaxY = Math.Max(chunk.DescendLocation.Get(1), chunk.AscentLocation.Get(1));

                    // Update the overall bounding box
                    minX = Math.Min(minX, chunkMinX);
                    minY = Math.Min(minY, chunkMinY);
                    maxX = Math.Max(maxX, chunkMaxX);
                    maxY = Math.Max(maxY, chunkMaxY);
                }

                // Create a Rectangle with (x, y, width, height)
                return new Rectangle(minX, minY, maxX - minX, maxY - minY);
            }
        }

        /// <summary>
        /// Class to store information about a text chunk's location
        /// </summary>
        private class TextChunkLocation
        {
            public string Text { get; }
            public Vector StartLocation { get; }
            public Vector EndLocation { get; }
            public Vector AscentLocation { get; }
            public Vector DescendLocation { get; }
            public FontDetails FontDetails { get; set; }

            public TextChunkLocation(string text, Vector descentLocation, Vector ascentLocation,
                Vector startLocation, Vector endLocation)
            {
                Text = text;
                StartLocation = startLocation;
                EndLocation = endLocation;
                AscentLocation = ascentLocation;
                DescendLocation = descentLocation;
            }
        }

        /// <summary>
        /// Class to store information about a text replacement
        /// </summary>
        private class TextReplacement
        {
            public string OriginalText { get; set; }
            public string NewText { get; set; }
            public Rectangle BoundingBox { get; set; }
            public FontDetails FontDetails { get; set; }
        }

        /// <summary>
        /// Class to store font details
        /// </summary>
        private class FontDetails
        {
            public string FontName { get; set; }
            public float FontSize { get; set; }
            public Color TextColor { get; set; }
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