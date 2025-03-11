using DocProcessingSystem.Core;
using DocProcessingSystem.Models;

namespace DocProcessingSystem.Services
{
    /// <summary>
    /// Manages the document processing workflow
    /// </summary>
    public class DocumentProcessingManager : IDisposable
    {
        private readonly IDocumentProcessor _converter;
        private readonly IPdfMerger _merger;
        private readonly IEnumerable<IDocumentTypeHandler> _handlers;
        private bool _disposed = false;

        /// <summary>
        /// Initializes a new document processing manager
        /// </summary>
        public DocumentProcessingManager(
            IDocumentProcessor converter,
            IPdfMerger merger,
            IEnumerable<IDocumentTypeHandler> handlers)
        {
            _converter = converter;
            _merger = merger;
            _handlers = handlers;
        }

        /// <summary>
        /// Processes all documents in the specified folders
        /// </summary>
        public void ProcessDocuments(Dictionary<string, string> keyValuesPairs, string analysisFolder)
        {
            Console.WriteLine("Document Processing System");
            Console.WriteLine("=========================\n");

            foreach (var item in keyValuesPairs)
            {
                var key = item.Key;
                var value = item.Value;

                if (!Directory.Exists(key))
                {
                    Console.WriteLine($"Warning: Parametrics folder not found: {value}");
                }
            }

            if (!Directory.Exists(analysisFolder))
                Console.WriteLine($"Warning: Analysis folder not found: {analysisFolder}");

            // Group analysis folders
            Console.WriteLine("Grouping analysis folders...");
            var analysisGroups = FolderHelper.GroupFolders(analysisFolder);
            Console.WriteLine($"Found {analysisGroups.Count} folder groups");

            // Process each document type
            foreach (var handler in _handlers)
            {
                foreach (var item in keyValuesPairs)
                {
                    var key = item.Key;
                    var value = item.Value;

                    if (handler.CanHandle(key) && Directory.Exists(value))
                    {
                        handler.ProcessDocuments(value, ref analysisGroups, in _merger);
                    }
                }
            }
            if (analysisGroups.Count > 0)
            {
                Console.WriteLine("Cannot Process Below Buildings");
                Console.WriteLine("=========================\n");
                foreach (var group in analysisGroups)
                {
                    Console.WriteLine($"TM No: {group.TmNo}, Building Code: {group.BuildingCode}, Building TM ID: {group.BuildingTmId}");
                }
            }

            Console.WriteLine("\nDocument processing completed.");
        }

        public void ProcessMasonry(string analysisFolder)
        {
            var analysisGroups = FolderHelper.GroupFolders(analysisFolder);

            foreach (var handler in _handlers)
            {
                handler.ProcessDocuments("", ref analysisGroups, _merger);
            }
        }

        /// <summary>
        /// Processes PDF merges for a specific group
        /// </summary>
        private void ProcessGroupMerges(FolderGroup group)
        {
            Console.WriteLine($"\nProcessing merges for group: TM {group.TmNo}, Building {group.BuildingCode}-{group.BuildingTmId}");

            // Get merge sequences from handlers
            var mergeSequences = new List<MergeSequence>();

            foreach (var handler in _handlers)
            {
                try
                {
                    var sequence = handler.GetMergeSequence(group);
                    if (sequence != null && File.Exists(sequence.MainDocument) &&
                        sequence.AdditionalDocuments.Any(File.Exists))
                    {
                        mergeSequences.Add(sequence);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not get merge sequence from handler {handler.GetType().Name}: {ex.Message}");
                }
            }

            // Execute merges
            foreach (var sequence in mergeSequences)
            {
                try
                {
                    Console.WriteLine($"Merging: {Path.GetFileName(sequence.OutputPath)}");
                    _merger.MergePdf(
                        sequence.MainDocument,
                        sequence.AdditionalDocuments,
                        sequence.OutputPath,
                        sequence.Options);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during merge: {ex.Message}");
                }
            }

            // Handle special case for multiple paths in a group
            if (group.PathCount > 1)
            {
                Console.WriteLine($"Group has {group.PathCount} paths - processing additional folders");

                // Implement specific logic for multi-path groups
                // This could include copying files between folders or creating additional merges
                for (int i = 1; i < group.Paths.Count; i++)
                {
                    string additionalPath = group.Paths[i];
                    string mainPath = group.MainFolder;

                    // Create HAKEDIS folder in additional path if it doesn't exist
                    string targetFolder = Path.Combine(additionalPath, "HAKEDIS");
                    Directory.CreateDirectory(targetFolder);

                    // Copy merged PDFs from main folder to additional folder
                    string sourceFolder = Path.Combine(mainPath, "HAKEDIS");
                    if (Directory.Exists(sourceFolder))
                    {
                        foreach (var pdf in Directory.GetFiles(sourceFolder, "*.pdf"))
                        {
                            string targetFile = Path.Combine(targetFolder, Path.GetFileName(pdf));
                            Console.WriteLine($"Copying {Path.GetFileName(pdf)} to {additionalPath}");
                            File.Copy(pdf, targetFile, true);
                        }
                    }
                }
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
                    // Dispose handlers that are disposable
                    foreach (var handler in _handlers)
                    {
                        if (handler is IDisposable disposable)
                        {
                            disposable.Dispose();
                        }
                    }

                    // Dispose converter and merger
                    _converter?.Dispose();
                    _merger?.Dispose();
                }

                _disposed = true;
            }
        }
    }
}