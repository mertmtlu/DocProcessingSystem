using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace DocProcessingSystem.Services
{
    /// <summary>
    /// Options class for defining keyword search parameters for PDF extraction
    /// </summary>
    public class KeywordOptions
    {
        /// <summary>
        /// The keyword to search for in the PDF
        /// </summary>
        public string Keyword { get; set; }

        /// <summary>
        /// Defines which occurrence of the keyword to use
        /// </summary>
        public KeywordOccurrence Occurrence { get; set; } = KeywordOccurrence.First;

        /// <summary>
        /// The specific occurrence index (used when Occurrence is set to Specific)
        /// </summary>
        public int OccurrenceIndex { get; set; } = 1;

        /// <summary>
        /// Determines if the page containing the keyword is included in the output
        /// </summary>
        public bool IncludeMatchingPage { get; set; } = true;
    }

    /// <summary>
    /// Enumeration for keyword occurrence options
    /// </summary>
    public enum KeywordOccurrence
    {
        First,
        Last,
        Specific
    }

    /// <summary>
    /// Options for determining the start or end page
    /// </summary>
    public enum PageSelectionType
    {
        /// <summary>
        /// Use a keyword to determine the page
        /// </summary>
        Keyword,

        /// <summary>
        /// Always use the first page of the document
        /// </summary>
        FirstPage,

        /// <summary>
        /// Always use the last page of the document
        /// </summary>
        LastPage,

        /// <summary>
        /// Use a specific page number
        /// </summary>
        SpecificPage
    }

    /// <summary>
    /// Options class for PDF extraction
    /// </summary>
    public class PdfExtractionOptions
    {
        /// <summary>
        /// Type of selection for the start page
        /// </summary>
        public PageSelectionType StartPageSelectionType { get; set; } = PageSelectionType.Keyword;

        /// <summary>
        /// Options for the start point keyword (used when StartPageSelectionType is Keyword)
        /// </summary>
        public KeywordOptions StartKeyword { get; set; }

        /// <summary>
        /// Specific page number to start from (used when StartPageSelectionType is SpecificPage)
        /// </summary>
        public int StartPageNumber { get; set; } = 1;

        /// <summary>
        /// Type of selection for the end page
        /// </summary>
        public PageSelectionType EndPageSelectionType { get; set; } = PageSelectionType.Keyword;

        /// <summary>
        /// Options for the end point keyword (used when EndPageSelectionType is Keyword)
        /// </summary>
        public KeywordOptions EndKeyword { get; set; }

        /// <summary>
        /// Specific page number to end at (used when EndPageSelectionType is SpecificPage)
        /// </summary>
        public int EndPageNumber { get; set; } = 1;

        /// <summary>
        /// If keywords are used and not found, determines whether to throw an exception or continue
        /// </summary>
        public bool ThrowIfKeywordNotFound { get; set; } = true;
    }

    /// <summary>
    /// Service for extracting a range from a PDF document based on keywords or page numbers
    /// </summary>
    public class PdfRangeExtractorService
    {
        /// <summary>
        /// Extracts a range of pages from a PDF based on keywords or page selection options and saves to a new file
        /// </summary>
        /// <param name="inputPdfPath">The path to the input PDF</param>
        /// <param name="outputPdfPath">The path where the output PDF should be saved</param>
        /// <param name="options">Options for controlling the extraction</param>
        public void ExtractRange(string inputPdfPath, string outputPdfPath, PdfExtractionOptions options)
        {
            if (!File.Exists(inputPdfPath))
                throw new FileNotFoundException($"Input PDF file not found: {inputPdfPath}");

            // Validate at least one selection option is properly configured
            ValidateOptions(options);

            // Ensure the output directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(outputPdfPath));

            Console.WriteLine("Determining page range for extraction...");

            // Find the page numbers based on options
            var pageInfo = DeterminePageRange(inputPdfPath, options);

            // Extract the pages
            ExtractPages(inputPdfPath, outputPdfPath, pageInfo.startPage, pageInfo.endPage);

            Console.WriteLine($"PDF extraction completed successfully. Output saved to: {outputPdfPath}");
        }

        /// <summary>
        /// Validates that the options are properly configured
        /// </summary>
        private void ValidateOptions(PdfExtractionOptions options)
        {
            // Validate start page selection
            if (options.StartPageSelectionType == PageSelectionType.Keyword && options.StartKeyword == null)
                throw new ArgumentException("StartKeyword must be specified when StartPageSelectionType is set to Keyword");

            if (options.StartPageSelectionType == PageSelectionType.SpecificPage && options.StartPageNumber < 1)
                throw new ArgumentException("StartPageNumber must be at least 1 when StartPageSelectionType is set to SpecificPage");

            // Validate end page selection
            if (options.EndPageSelectionType == PageSelectionType.Keyword && options.EndKeyword == null)
                throw new ArgumentException("EndKeyword must be specified when EndPageSelectionType is set to Keyword");

            if (options.EndPageSelectionType == PageSelectionType.SpecificPage && options.EndPageNumber < 1)
                throw new ArgumentException("EndPageNumber must be at least 1 when EndPageSelectionType is set to SpecificPage");
        }

        /// <summary>
        /// Determines the start and end page numbers based on the configured options
        /// </summary>
        private (int startPage, int endPage) DeterminePageRange(string pdfPath, PdfExtractionOptions options)
        {
            int startPage = 1;
            int endPage = -1;

            using (var pdfReader = new PdfReader(pdfPath))
            using (var pdfDocument = new PdfDocument(pdfReader))
            {
                int totalPages = pdfDocument.GetNumberOfPages();
                endPage = totalPages; // Default to last page

                // Determine start page based on selection type
                switch (options.StartPageSelectionType)
                {
                    case PageSelectionType.FirstPage:
                        startPage = 1;
                        Console.WriteLine("Using first page of document as start page");
                        break;

                    case PageSelectionType.LastPage:
                        startPage = totalPages;
                        Console.WriteLine("Using last page of document as start page");
                        break;

                    case PageSelectionType.SpecificPage:
                        startPage = Math.Min(options.StartPageNumber, totalPages);
                        Console.WriteLine($"Using specific page {startPage} as start page");
                        break;

                    case PageSelectionType.Keyword:
                        // Use keyword to find start page
                        if (options.StartKeyword != null)
                        {
                            var page = FindKeywordPage(pdfDocument, options.StartKeyword);
                            if (page > 0)
                            {
                                startPage = options.StartKeyword.IncludeMatchingPage ? page : page + 1;
                                Console.WriteLine($"Start keyword '{options.StartKeyword.Keyword}' found on page {page}");
                            }
                            else if (options.ThrowIfKeywordNotFound)
                            {
                                throw new InvalidOperationException($"Start keyword '{options.StartKeyword.Keyword}' not found in the document");
                            }
                        }
                        break;
                }

                // Determine end page based on selection type
                switch (options.EndPageSelectionType)
                {
                    case PageSelectionType.FirstPage:
                        endPage = 1;
                        Console.WriteLine("Using first page of document as end page");
                        break;

                    case PageSelectionType.LastPage:
                        endPage = totalPages;
                        Console.WriteLine("Using last page of document as end page");
                        break;

                    case PageSelectionType.SpecificPage:
                        endPage = Math.Min(options.EndPageNumber, totalPages);
                        Console.WriteLine($"Using specific page {endPage} as end page");
                        break;

                    case PageSelectionType.Keyword:
                        // Use keyword to find end page
                        if (options.EndKeyword != null)
                        {
                            var page = FindKeywordPage(pdfDocument, options.EndKeyword);
                            if (page > 0)
                            {
                                endPage = options.EndKeyword.IncludeMatchingPage ? page : page - 1;
                                Console.WriteLine($"End keyword '{options.EndKeyword.Keyword}' found on page {page}");
                            }
                            else if (options.ThrowIfKeywordNotFound)
                            {
                                throw new InvalidOperationException($"End keyword '{options.EndKeyword.Keyword}' not found in the document");
                            }
                        }
                        break;
                }

                // Validate the range
                if (startPage > endPage)
                {
                    throw new InvalidOperationException($"Invalid page range: Start page ({startPage}) is after end page ({endPage})");
                }

                return (startPage, endPage);
            }
        }

        /// <summary>
        /// Finds the page number where a keyword occurs based on the specified options
        /// </summary>
        private int FindKeywordPage(PdfDocument pdfDocument, KeywordOptions options)
        {
            List<int> occurrences = new List<int>();

            // Search for the keyword in each page
            for (int i = 1; i <= pdfDocument.GetNumberOfPages(); i++)
            {
                string pageText = PdfTextExtractor.GetTextFromPage(pdfDocument.GetPage(i), new SimpleTextExtractionStrategy());

                // Check if the keyword is on this page
                if (pageText.Contains(options.Keyword, StringComparison.OrdinalIgnoreCase))
                {
                    occurrences.Add(i);
                }
            }

            if (occurrences.Count == 0)
                return 0;

            // Return the appropriate occurrence based on the options
            switch (options.Occurrence)
            {
                case KeywordOccurrence.First:
                    return occurrences[0];

                case KeywordOccurrence.Last:
                    return occurrences[occurrences.Count - 1];

                case KeywordOccurrence.Specific:
                    int index = options.OccurrenceIndex - 1; // Convert from 1-based to 0-based
                    if (index >= 0 && index < occurrences.Count)
                        return occurrences[index];
                    else
                        throw new InvalidOperationException($"Specific occurrence {options.OccurrenceIndex} of keyword '{options.Keyword}' not found. Only {occurrences.Count} occurrences exist.");

                default:
                    return occurrences[0];
            }
        }

        /// <summary>
        /// Extracts pages from the source PDF to a new PDF
        /// </summary>
        private void ExtractPages(string sourcePdfPath, string targetPdfPath, int startPage, int endPage)
        {
            Console.WriteLine($"Extracting pages {startPage} to {endPage}...");

            using (var reader = new PdfReader(sourcePdfPath))
            using (var writer = new PdfWriter(targetPdfPath))
            using (var sourceDoc = new PdfDocument(reader))
            using (var targetDoc = new PdfDocument(writer))
            {
                // Calculate the actual pages to copy
                int totalPages = sourceDoc.GetNumberOfPages();
                startPage = Math.Max(1, Math.Min(startPage, totalPages));
                endPage = Math.Max(startPage, Math.Min(endPage, totalPages));

                // Copy the pages
                sourceDoc.CopyPagesTo(startPage, endPage, targetDoc);
            }
        }
    }
}