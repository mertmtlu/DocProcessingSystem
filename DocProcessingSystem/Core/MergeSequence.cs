namespace DocProcessingSystem.Core
{
    /// <summary>
    /// Defines the sequence for merging PDFs
    /// </summary>
    public class MergeSequence
    {
        /// <summary>
        /// Main document to use as base
        /// </summary>
        public string MainDocument { get; set; }

        /// <summary>
        /// Additional documents to merge in order
        /// </summary>
        public List<string> AdditionalDocuments { get; set; } = new List<string>();

        /// <summary>
        /// Output path for the merged document
        /// </summary>
        public string OutputPath { get; set; }

        /// <summary>
        /// Custom merge options
        /// </summary>
        public MergeOptions Options { get; set; } = new MergeOptions();
    }
}
