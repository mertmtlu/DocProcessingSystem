using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem.Core
{
    /// <summary>
    /// Interface for PDF reading operations
    /// </summary>
    public interface IPdfReader : IDisposable
    {
        /// <summary>
        /// Checks if a specific page exists in the PDF
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="pageNumber">Page number to check</param>
        /// <returns>True if the page exists, false otherwise</returns>
        bool PageExists(string pdfPath, int pageNumber);

        /// <summary>
        /// Checks if a specific string exists in the PDF
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="searchText">Text to search for</param>
        /// <param name="isCaseSensitive">Whether the search should be case sensitive</param>
        /// <returns>True if the string exists, false otherwise</returns>
        bool ContainsText(string pdfPath, string searchText, bool isCaseSensitive = false);

        /// <summary>
        /// Searches for text in a PDF and returns the page numbers where it was found
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="searchText">Text to search for</param>
        /// <param name="isCaseSensitive">Whether the search should be case sensitive</param>
        /// <returns>List of page numbers where the text was found</returns>
        List<int> FindTextInPages(string pdfPath, string searchText, bool isCaseSensitive = false);

        /// <summary>
        /// Extracts text from a specific page in the PDF
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <param name="pageNumber">Page number to extract text from</param>
        /// <returns>Extracted text from the specified page</returns>
        string ExtractTextFromPage(string pdfPath, int pageNumber);

        /// <summary>
        /// Extracts text from all pages in the PDF
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file</param>
        /// <returns>Dictionary with page numbers as keys and extracted text as values</returns>
        Dictionary<int, string> ExtractTextFromAllPages(string pdfPath);
    }
}
