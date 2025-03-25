namespace DocProcessingSystem.Core
{
    /// <summary>
    /// Options for PDF merging
    /// </summary>
    public class MergeOptions
    {
        /// <summary>
        /// Patterns for files to exclude from merging
        /// </summary>
        public List<string> ExcludePatterns { get; set; } = new List<string>();

        /// <summary>
        /// Required sections that must be included in the merge
        /// </summary>
        public string[] RequiredSections { get; set; } = Array.Empty<string>();

        /// <summary>
        /// Whether to preserve bookmarks from source documents
        /// </summary>
        public bool PreserveBookmarks { get; set; } = true;

        /// <summary>
        /// Whether to create a new bookmark for each additional PDF
        /// </summary>
        public bool CreateBookmarksForAdditionalPdfs { get; set; } = false;
    }
}
