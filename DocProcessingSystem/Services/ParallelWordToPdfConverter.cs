using DocProcessingSystem.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace DocProcessingSystem.Services
{
    /// <summary>
    /// Converts Word documents to PDF format with parallel processing support
    /// </summary>
    public class ParallelWordToPdfConverter : IDisposable
    {
        private readonly int _maxDegreeOfParallelism;
        private readonly SemaphoreSlim _semaphore;
        private bool _disposed;

        public ParallelWordToPdfConverter(int maxDegreeOfParallelism = 5)
        {
            // Limit the number of concurrent Word instances to avoid resource issues
            _maxDegreeOfParallelism = maxDegreeOfParallelism > 0 ? maxDegreeOfParallelism : Environment.ProcessorCount;
            _semaphore = new SemaphoreSlim(_maxDegreeOfParallelism);
        }

        /// <summary>
        /// Asynchronously converts multiple Word documents to PDF format
        /// </summary>
        public async System.Threading.Tasks.Task ConvertMultipleAsync(string[] inputPaths, string outputFolderPath, bool saveWordChanges,
            IProgress<(string FileName, int Completed, int Total)> progress = null,
            CancellationToken cancellationToken = default)
        {
            if (inputPaths == null || inputPaths.Length == 0)
                throw new ArgumentException("No input files provided", nameof(inputPaths));

            Directory.CreateDirectory(outputFolderPath);

            var tasks = new List<System.Threading.Tasks.Task>();
            var counter = 0;
            var total = inputPaths.Length;

            foreach (var inputPath in inputPaths)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // Wait until a semaphore slot is available
                await _semaphore.WaitAsync(cancellationToken);

                var task = System.Threading.Tasks.Task.Run(async () =>
                {
                    try
                    {
                        var baseName = Path.GetFileNameWithoutExtension(inputPath);
                        var outputPath = Path.Combine(outputFolderPath, baseName + ".pdf");

                        // Create a new instance of Word for each conversion job
                        using (var converter = new WordToPdfConverter())
                        {
                            converter.Convert(inputPath, outputPath, saveWordChanges);
                        }

                        var completed = Interlocked.Increment(ref counter);
                        progress?.Report((Path.GetFileName(inputPath), completed, total));
                    }
                    finally
                    {
                        // Release the semaphore slot
                        _semaphore.Release();
                    }
                }, cancellationToken);

                tasks.Add(task);
            }

            // Wait for all conversion tasks to complete
            await System.Threading.Tasks.Task.WhenAll(tasks);
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
                    _semaphore?.Dispose();
                }
                _disposed = true;
            }
        }
    }

    // Extension method for Directory class
    public static class DirectoryExtensions
    {
        /// <summary>
        /// Gets all files with specified pattern asynchronously
        /// </summary>
        public static async Task<string[]> GetFilesAsync(string path, string pattern,
            SearchOption searchOption = SearchOption.TopDirectoryOnly)
        {
            return await System.Threading.Tasks.Task.Run(() => Directory.GetFiles(path, pattern, searchOption));
        }
    }
}