using DocProcessingSystem.Core;
using DocProcessingSystem.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem.Models
{
    public class PdfOperationsHelper
    {
        #region Public Methods


        public static async Task ConvertWordToPdfAsync(string inputFolderPath, string outputFolderPath, bool saveChanges, bool useRelativePath = false, int maxParallel = 5)
        {
            // Get all Word files asynchronously
            var wordFiles = await DirectoryExtensions.GetFilesAsync(inputFolderPath, "*.docx", SearchOption.AllDirectories);

            // Create progress reporting
            var progress = new Progress<(string FileName, int Completed, int Total)>(update =>
            {
                Console.WriteLine($"Converted {update.FileName} - {update.Completed} of {update.Total} completed");
            });

            // If using relative paths, we need to handle each file individually
            if (useRelativePath)
            {
                Console.WriteLine($"Starting conversion of {wordFiles.Length} files with relative paths");

                // Create a semaphore to limit concurrency
                using var semaphore = new SemaphoreSlim(maxParallel, maxParallel);
                using var cts = new CancellationTokenSource();

                // Create a list of tasks
                var tasks = new List<Task>();
                int completedCount = 0;

                foreach (var wordFile in wordFiles)
                {
                    // Wait for a slot to be available
                    await semaphore.WaitAsync(cts.Token);

                    // Create a task for each file
                    var task = Task.Run(async () =>
                    {
                        try
                        {
                            // Generate output path in the same directory as the input file
                            string outputPath = wordFile.Replace(".docx", ".pdf");

                            // Create a single-file converter for this task
                            using (var converter = new WordToPdfConverter())
                            {
                                converter.Convert(wordFile, outputPath, saveChanges, false);
                            }

                            // Report progress
                            int currentCount = Interlocked.Increment(ref completedCount);
                            ((IProgress<(string, int, int)>)progress).Report((Path.GetFileName(wordFile), currentCount, wordFiles.Length));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error converting {wordFile}: {ex.Message}");
                        }
                        finally
                        {
                            // Release the semaphore slot
                            semaphore.Release();
                        }
                    }, cts.Token);

                    tasks.Add(task);
                }

                try
                {
                    // Wait for all tasks to complete
                    await Task.WhenAll(tasks);
                    Console.WriteLine("All files converted successfully");
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine("Operation was cancelled");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during conversion: {ex.Message}");
                }
            }
            else
            {
                // When not using relative paths, use the original implementation
                using (var cts = new CancellationTokenSource())
                using (var converter = new ParallelWordToPdfConverter(maxParallel))
                {
                    Console.WriteLine($"Starting conversion of {wordFiles.Length} files with {maxParallel} parallel workers");

                    try
                    {
                        // Convert all files in parallel with controlled concurrency
                        await converter.ConvertMultipleAsync(
                            wordFiles,
                            outputFolderPath,
                            saveChanges,
                            progress,
                            cts.Token);

                        Console.WriteLine("All files converted successfully");
                    }
                    catch (OperationCanceledException)
                    {
                        Console.WriteLine("Operation was cancelled");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error during conversion: {ex.Message}");
                    }
                }
            }
        }

        #endregion

    }
}
