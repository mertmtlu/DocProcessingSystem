using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem.Models
{
    public static class FileCounter
    {
        public static Dictionary<string, Dictionary<string, int>> CountFilesBySubfolder(string rootFolder)
        {
            var result = new Dictionary<string, Dictionary<string, int>>();
            var targetExtensions = new[] { ".jpg", ".jpeg", ".JPG", ".JPEG", ".pdf" };

            try
            {
                // Validate root folder exists
                if (!Directory.Exists(rootFolder))
                {
                    throw new DirectoryNotFoundException($"Root folder not found: {rootFolder}");
                }

                // Get all subdirectories (including nested ones)
                var subdirectories = Directory.GetDirectories(rootFolder, "*", SearchOption.AllDirectories);

                // Also include the root folder itself
                var allDirectories = new List<string> { rootFolder };
                allDirectories.AddRange(subdirectories);

                foreach (var directory in allDirectories)
                {
                    try
                    {
                        var fileCounts = new Dictionary<string, int>();

                        // Initialize counts for each extension
                        foreach (var ext in targetExtensions)
                        {
                            fileCounts[ext] = 0;
                        }

                        // Get all files in current directory (not recursive)
                        var files = Directory.GetFiles(directory, "*", SearchOption.TopDirectoryOnly);

                        // Count files by extension
                        foreach (var file in files)
                        {
                            var extension = Path.GetExtension(file);
                            if (targetExtensions.Contains(extension))
                            {
                                fileCounts[extension]++;
                            }
                        }

                        // Only add to result if there are any files of interest
                        if (fileCounts.Values.Any(count => count > 0))
                        {
                            result[directory] = fileCounts;
                        }
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // Skip directories we don't have permission to access
                        Console.WriteLine($"Access denied to directory: {directory}");
                    }
                    catch (DirectoryNotFoundException)
                    {
                        // Skip directories that may have been deleted during enumeration
                        Console.WriteLine($"Directory not found (may have been deleted): {directory}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing root folder: {ex.Message}");
                throw;
            }

            return result;
        }

        // Alternative method that returns results in a more readable format
        public static void PrintFileCountsBySubfolder(string rootFolder)
        {
            var results = CountFilesBySubfolder(rootFolder);

            Console.WriteLine($"File counts for root folder: {rootFolder}");
            Console.WriteLine(new string('=', 60));

            if (!results.Any())
            {
                Console.WriteLine("No files with specified extensions found in any subdirectory.");
                return;
            }

            foreach (var kvp in results.OrderBy(x => x.Key))
            {
                Console.WriteLine($"\nDirectory: {kvp.Key}");

                var totalFiles = kvp.Value.Values.Sum();
                Console.WriteLine($"Total files: {totalFiles}");

                foreach (var fileCount in kvp.Value.Where(fc => fc.Value > 0))
                {
                    Console.WriteLine($"  {fileCount.Key}: {fileCount.Value} files");
                }
            }
        }
    }

}
