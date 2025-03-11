using DocProcessingSystem.Models;

namespace DocProcessingSystem.Core
{
    /// <summary>
    /// Interface for handling specific document types (parametric or deterministic)
    /// </summary>
    public interface IDocumentTypeHandler
    {
        /// <summary>
        /// Determines if this handler can process the specified document type
        /// </summary>
        bool CanHandle(string documentType);

        /// <summary>
        /// Processes documents from source folder against analysis groups
        /// </summary>
        void ProcessDocuments(string sourceFolder, ref List<FolderGroup> analysisGroups, in IPdfMerger merger); // TODO: Check ref keyword is correctly used

        /// <summary>
        /// Gets the merge sequence for a specific folder group
        /// </summary>
        MergeSequence GetMergeSequence(FolderGroup group);
    }
}
