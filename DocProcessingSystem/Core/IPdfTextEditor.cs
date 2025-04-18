using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DocProcessingSystem.Core
{
    /// <summary>
    /// Interface for PDF text editing operations
    /// </summary>
    public interface IPdfTextEditor : IDisposable
    {
        /// <summary>
        /// Replaces text in a PDF document
        /// </summary>
        /// <param name="inputPath">Path to the input PDF</param>
        /// <param name="outputPath">Path where the modified PDF will be saved</param>
        /// <param name="textToReplace">Text to be replaced</param>
        /// <param name="replacementText">Text to replace with</param>
        /// <param name="isCaseSensitive">Whether the search should be case sensitive</param>
        /// <returns>Number of replacements made</returns>
        int ReplaceText(string inputPath, string outputPath, string textToReplace, string replacementText, bool isCaseSensitive = false);

        /// <summary>
        /// Replaces text in a PDF document on specific pages
        /// </summary>
        /// <param name="inputPath">Path to the input PDF</param>
        /// <param name="outputPath">Path where the modified PDF will be saved</param>
        /// <param name="textToReplace">Text to be replaced</param>
        /// <param name="replacementText">Text to replace with</param>
        /// <param name="pageNumbers">Specific pages to perform replacement on</param>
        /// <param name="isCaseSensitive">Whether the search should be case sensitive</param>
        /// <returns>Number of replacements made</returns>
        int ReplaceText(string inputPath, string outputPath, string textToReplace, string replacementText, IEnumerable<int> pageNumbers, bool isCaseSensitive = false);

        /// <summary>
        /// Replaces text in a PDF document using regular expressions
        /// </summary>
        /// <param name="inputPath">Path to the input PDF</param>
        /// <param name="outputPath">Path where the modified PDF will be saved</param>
        /// <param name="pattern">Regular expression pattern to match text</param>
        /// <param name="replacement">Replacement text (can include regex groups)</param>
        /// <param name="options">Regular expression options</param>
        /// <returns>Number of replacements made</returns>
        int ReplaceTextWithRegex(string inputPath, string outputPath, string pattern, string replacement, RegexOptions options = RegexOptions.None);

        /// <summary>
        /// Adds editable text field to a specific position in a PDF
        /// </summary>
        /// <param name="inputPath">Path to the input PDF</param>
        /// <param name="outputPath">Path where the modified PDF will be saved</param>
        /// <param name="pageNumber">Page number to add the text field to</param>
        /// <param name="fieldName">Name of the form field</param>
        /// <param name="x">X coordinate (left position)</param>
        /// <param name="y">Y coordinate (bottom position)</param>
        /// <param name="width">Width of the text field</param>
        /// <param name="height">Height of the text field</param>
        /// <param name="initialValue">Initial text value</param>
        void AddEditableTextField(string inputPath, string outputPath, int pageNumber, string fieldName,
                                 float x, float y, float width, float height, string initialValue = "");
    }
}