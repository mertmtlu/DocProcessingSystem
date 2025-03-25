namespace DocProcessingSystem.Core
{
    /// <summary>
    /// Represents a document processing operation like conversion
    /// </summary>
    public interface IDocumentProcessor : IDisposable
    {
        /// <summary>
        /// Converts a document from one format to another
        /// </summary>
        void Convert(string inputPath, string outputPath, bool saveWordChanges, bool copyWord = true);
    }
}
