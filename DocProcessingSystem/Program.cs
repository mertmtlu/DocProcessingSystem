using DocProcessingSystem.Core;
using DocProcessingSystem.Models;
using DocProcessingSystem.Services;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace DocProcessingSystem
{
    public class Program
    {
        /// <summary>
        /// Main application entry method
        /// </summary>
        static void Main(string[] args)
        {
            ProcessSubFolder(@"C:\Users\Mert\Desktop\Mert EK-A - Kopya\TM Folders", @"C:\Users\Mert\Desktop\TM Folders");
        }

        static void HandleMasonry()
        {
            string inputFolderLocation = @"C:\Users\Mert\Desktop\Selin Report Revision\v7\Automatic";
            var grouppedFolders = FolderHelper.GroupFolders(inputFolderLocation);

            using (var converter = new WordToPdfConverter())
            using (var merger = new PdfMergerService())
            {
                // Create folder matcher
                var matcher = new FolderNameMatcher();

                // Create document handlers
                var handlers = new IDocumentTypeHandler[]
                {
                    new MasonryDocumentHandler(converter, matcher),
                };

                // Create processing manager
                using (var manager = new DocumentProcessingManager(converter, merger, handlers))
                {
                    // Process all documents
                    manager.ProcessMasonry(inputFolderLocation);
                }
            }

        }

        static void ConvertWordsToPdf()
        {
            var inputFolderLocation = @"C:\Users\Mert\Desktop\testInput";
            var outputFolderLocation = @"C:\Users\Mert\Desktop\testOutput";
            var wordFiles = Directory.GetFiles(inputFolderLocation, "*.docx", SearchOption.AllDirectories);

            using (var converter = new WordToPdfConverter())
            {
                foreach (var word in wordFiles)
                {
                    var outputLocation = Path.Combine(outputFolderLocation, Path.GetFileNameWithoutExtension(word));
                    converter.Convert(word, outputLocation + ".pdf");
                }
            }
        }

        static void ProcessDocuments()
        {
            // Get folder paths from arguments or use defaults
            string parametricsFolder = @"C:\Users\Mert\Desktop\testing\Parametric";
            string deterministicsFolder = @"C:\Users\Mert\Desktop\testing\Deterministic";
            string post2008 = @"C:\Users\Mert\Desktop\testing\Post2008";
            string analysisFolder = @"C:\Users\Mert\Desktop\testing\Analysis";

            using (var converter = new WordToPdfConverter())
            using (var merger = new PdfMergerService())
            {
                // Create folder matcher
                var matcher = new FolderNameMatcher();

                // Create document handlers
                var handlers = new IDocumentTypeHandler[]
                {
                    new Post2008DocumentHandler(converter, matcher),
                    new ParametricDocumentHandler(converter, matcher),
                    new DeterministicDocumentHandler(converter, matcher)
                };

                Dictionary<string, string> pathDictionary = new()
                {
                    {"Parametric", parametricsFolder},
                    {"Deterministic", deterministicsFolder},
                    {"Post2008", post2008},
                };

                // Create processing manager
                using (var manager = new DocumentProcessingManager(converter, merger, handlers))
                {
                    // Process all documents
                    manager.ProcessDocuments(pathDictionary, analysisFolder);
                }
            }

            Console.WriteLine("\nProcess completed. Press any key to exit.");
            Console.ReadKey();
        }

        // Example
        static void ExtractPagesFromPdf()
        {
            var options = new PdfExtractionOptions
            {
                StartPageSelectionType = PageSelectionType.Keyword,
                StartKeyword = new KeywordOptions
                {
                    Keyword = "EK-A",
                    Occurrence = KeywordOccurrence.Last,
                    IncludeMatchingPage = true
                },
                EndPageSelectionType = PageSelectionType.Keyword,
                EndKeyword = new KeywordOptions
                {
                    Keyword = "EK-B",
                    Occurrence = KeywordOccurrence.Last,
                    IncludeMatchingPage = false
                }
            };

            var service = new PdfRangeExtractorService();
            service.ExtractRange(
                @"C:\Users\Mert\Desktop\TEI-B01-TM-24-DIR-M10-01_nt.pdf",   // TODO: Change Input Pdf Location 
                @"C:\Users\Mert\Desktop\EK-A.pdf",                          // TODO: Change Output Pdf Location
                options
            );
        }

        /// <summary>
        /// Processes all subfolders in the input directory, merging PDFs and outputting to the corresponding output directory
        /// </summary>
        /// <param name="inputFolderPath">Path to the input folder containing subfolders with PDFs</param>
        /// <param name="outputFolderPath">Path to the output folder where processed PDFs will be stored</param>
        public static void ProcessSubFolder(string inputFolderPath, string outputFolderPath)
        {
            // Validate input parameters
            if (string.IsNullOrEmpty(inputFolderPath))
                throw new ArgumentNullException(nameof(inputFolderPath), "Input folder path cannot be null or empty");

            if (string.IsNullOrEmpty(outputFolderPath))
                throw new ArgumentNullException(nameof(outputFolderPath), "Output folder path cannot be null or empty");

            if (!Directory.Exists(inputFolderPath))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolderPath}");

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolderPath);

            // Get the cover page path
            string projectRootPath = AppDomain.CurrentDomain.BaseDirectory;
            string deterministicCoverPath = Path.Combine(projectRootPath, "CoverPages", "Deterministic");
            string mainPdf = Path.Combine(deterministicCoverPath, "EK-A_Kapak.pdf");

            // Handle absent ocver page
            if (!File.Exists(mainPdf))
                throw new FileNotFoundException($"Cover page not found: {mainPdf}");

            // Log
            Console.WriteLine($"Starting processing of {inputFolderPath}");
            Console.WriteLine($"Output will be saved to {outputFolderPath}");
            Console.WriteLine($"Using cover page: {mainPdf}");

            // Get all subfolders
            string[] subfolders = Directory.GetDirectories(inputFolderPath);
            Console.WriteLine($"Found {subfolders.Length} subfolders to process");

            // Process each subfolder
            foreach (string subfolder in subfolders)
            {
                string subfolderName = Path.GetFileName(subfolder);
                string outputSubfolderPath = Path.Combine(outputFolderPath, subfolderName);

                try
                {
                    ProcessSingleSubfolder(subfolder, outputSubfolderPath, mainPdf);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing subfolder {subfolderName}: {ex.Message}");
                }
            }

            Console.WriteLine("Subfolder processing completed.");
        }

        /// <summary>
        /// Processes a single subfolder, merging all PDFs with the cover page
        /// </summary>
        private static void ProcessSingleSubfolder(string subfolderPath, string outputSubfolderPath, string mainPdf)
        {
            string subfolderName = Path.GetFileName(subfolderPath);
            Console.WriteLine($"Processing subfolder: {subfolderName}");

            // Ensure output subfolder exists
            Directory.CreateDirectory(outputSubfolderPath);

            // Get all PDF files in the subfolder
            string[] pdfFiles = Directory.GetFiles(subfolderPath, "*.pdf", SearchOption.TopDirectoryOnly)
                                        .OrderBy(f => Path.GetFileName(f))
                                        .ToArray();

            if (pdfFiles.Length == 0)
            {
                Console.WriteLine($"No PDF files found in {subfolderName}, skipping");
                return;
            }

            Console.WriteLine($"Found {pdfFiles.Length} PDF files in {subfolderName}");

            // Prepare the merger options
            var mergeOptions = new MergeOptions
            {
                PreserveBookmarks = true
            };

            // Define output path for the merged file
            string outputFilePath = Path.Combine(outputSubfolderPath, $"EK-A.pdf");

            // Create list of additional PDFs
            List<string> additionalPdfs = new List<string>(pdfFiles);

            // Use the PdfMergerService to merge the files
            using (var pdfMerger = new PdfMergerService())
            {
                Console.WriteLine($"Merging PDFs for {subfolderName}");
                pdfMerger.MergePdf(mainPdf, additionalPdfs, outputFilePath, mergeOptions);
                Console.WriteLine($"Merged PDF saved to: {outputFilePath}");
            }
        }

        /// <summary>
        /// Gets a folder path from arguments or user input
        /// </summary>
        static string GetFolderPath(string[] args, int index, string prompt, string defaultPath)
        {
            if (args.Length > index && Directory.Exists(args[index]))
                return args[index];

            Console.Write(prompt + " ");
            string? input = Console.ReadLine();

            // If input is empty or invalid, use default
            if (string.IsNullOrWhiteSpace(input) || !Directory.Exists(input))
            {
                Console.WriteLine($"Using default path: {defaultPath}");
                return defaultPath;
            }

            return input;
        }
    }
}