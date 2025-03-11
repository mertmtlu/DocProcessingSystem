namespace DocProcessingSystem.Core
{
    /// <summary>
    /// Handles merging of PDF documents
    /// </summary>
    public interface IPdfMerger : IDisposable
    {
        /// <summary>
        /// Merges multiple PDF files into one output file
        /// </summary>
        void MergePdf(string mainPdf, List<string> additionalPdfs, string outputPath, MergeOptions options);
    }
}
