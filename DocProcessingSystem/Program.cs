using DocProcessingSystem.Core;
using DocProcessingSystem.Models;
using DocProcessingSystem.Services;
using iText.Kernel.Numbering;
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
            ProcessDocuments();
        }
        
        public static void HandleCrisis()
        {
            var inputFolder = @"C:\Users\Mert\Desktop\HK15";
            var outputFolder = @"C:\Users\Mert\Desktop\HK15 TM FOLDERS";

            var ekCFiles = Directory.GetFiles(inputFolder, "EK-C.pdf", SearchOption.AllDirectories);

            foreach (var file in ekCFiles)
            {
                try
                {
                    Console.WriteLine("Processing: " + file);
                    var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(file);

                    if (tmNo == null || buildingCode == null || buildingTmId == null) throw new ArgumentNullException(file);
                    var outputFolderName = $"{tmNo}_M{buildingCode}_{buildingTmId}";

                    var outputDir = Path.Combine(outputFolder, outputFolderName);

                    if (!Directory.Exists(outputFolderName))
                    {
                        Directory.CreateDirectory(outputFolderName);
                    }

                    ExtractPagesFromPdf(file, outputDir);

                    string directoryPath = Path.GetDirectoryName(file);

                    // Files to copy
                    var ekA = Path.Combine(directoryPath, "EK-A.pdf");
                    var ekB = Path.Combine(directoryPath, "EK-B.pdf");
                    var ekC = Path.Combine(directoryPath, "EK-C.pdf");

                    // folders to copy
                    var buildingImages = Path.Combine(directoryPath, "Analysis", "BuildingImages");
                    var planFromCad = Path.Combine(directoryPath, "Analysis", "PlanFromCAD");

                    // Copy files and folders to outputDir
                    if (File.Exists(ekA)) File.Copy(ekA, Path.Combine(outputDir, "EK-A.pdf"), true);
                    if (File.Exists(ekB)) File.Copy(ekB, Path.Combine(outputDir, "EK-B.pdf"), true);
                    if (File.Exists(ekC)) File.Copy(ekC, Path.Combine(outputDir, "EK-C.pdf"), true);

                    if (Directory.Exists(buildingImages))
                    {
                        var targetBuildingImages = Path.Combine(outputDir, "BuildingImages");
                        CopyDirectory(buildingImages, targetBuildingImages);
                    }

                    if (Directory.Exists(planFromCad))
                    {
                        var targetPlanFromCad = Path.Combine(outputDir, "PlanFromCAD");
                        CopyDirectory(planFromCad, targetPlanFromCad);
                    }

                    Console.WriteLine("");
                }
                catch
                {
                    Console.WriteLine($"ERROR: Cannot process file, Name: {Path.GetFileName(file)}, Path: {file}");
                }

            }
        }
        public static void CopyDirectory(string sourceDir, string destinationDir)
        {
            if (!Directory.Exists(destinationDir))
            {
                Directory.CreateDirectory(destinationDir);
            }

            // Copy all files
            foreach (var file in Directory.GetFiles(sourceDir))
            {
                string fileName = Path.GetFileName(file);
                string destFile = Path.Combine(destinationDir, fileName);
                File.Copy(file, destFile, true); // true to overwrite if exists
            }

            foreach (var directory in Directory.GetDirectories(sourceDir))
            {
                string dirName = Path.GetFileName(directory);
                string destDir = Path.Combine(destinationDir, dirName);
                CopyDirectory(directory, destDir);
            }
        }

        static void ConvertWordToPdf(string inputFolderPath, string outputFolderPath, bool saveChanges)
        {
            var wordFiles = Directory.GetFiles(inputFolderPath, "*.docx", SearchOption.AllDirectories);

            using (var converter = new WordToPdfConverter())
            {
                foreach (string file in wordFiles)
                {
                    var baseName = Path.GetFileNameWithoutExtension(file);

                    var outputPath = Path.Combine(outputFolderPath, baseName + ".pdf");

                    converter.Convert(file, outputPath, saveChanges);
                }
            }
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

        static void ProcessDocuments()
        {
            // Get folder paths from arguments or use defaults
            string parametricsFolder = @"C:\Users\Mert\Desktop\testing\Parametric";
            string deterministicsFolder = @"C:\Users\Mert\Desktop\testing\Deterministic";
            string post2008 = @"C:\Users\Mert\Desktop\fırat\fırat\WORD";
            string analysisFolder = @"C:\Users\Mert\Desktop\fırat\fırat\NİHAİ_TESLİM";

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
        static void ExtractPagesFromPdf(string file, string outputFolder)
        {
            var service = new PdfRangeExtractorService();

            var summaryOptions = new PdfExtractionOptions
            {
                StartPageSelectionType = PageSelectionType.Keyword,
                StartKeyword = new KeywordOptions
                {
                    Keyword = "HASARSIZ TESPİT EDİLEN DONATILAR İÇİN",
                    Occurrence = KeywordOccurrence.Last,
                    IncludeMatchingPage = true
                },
                EndPageSelectionType = PageSelectionType.Keyword,
                EndKeyword = new KeywordOptions
                {
                    Keyword = "PASPAYI SIYIRMA YÖNTEMİ İLE TESPİT EDİLEN DONATILAR İÇİN",
                    Occurrence = KeywordOccurrence.Last,
                    IncludeMatchingPage = true
                }
            };

            service.ExtractRange(
                file, // TODO: Change Input Pdf Location 
                Path.Combine(outputFolder, Path.GetFileName(outputFolder) + "_Özet_Rapor.pdf"),
                summaryOptions
            );

            var sıyırmaOptions = new PdfExtractionOptions
            {
                StartPageSelectionType = PageSelectionType.Keyword,
                StartKeyword = new KeywordOptions
                {
                    Keyword = "EK-C",
                    Occurrence = KeywordOccurrence.Last,
                    IncludeMatchingPage = false
                },
                EndPageSelectionType = PageSelectionType.Keyword,
                EndKeyword = new KeywordOptions
                {
                    Keyword = "HASARSIZ TESPİT EDİLEN DONATILAR İÇİN",
                    Occurrence = KeywordOccurrence.Last,
                    IncludeMatchingPage = false
                }
            };

            service.ExtractRange(
                file,
                Path.Combine(outputFolder, "SIYIRMA FOTO.pdf"),
                sıyırmaOptions
                );

            var rontgenOptions = new PdfExtractionOptions
            {
                StartPageSelectionType = PageSelectionType.Keyword,
                StartKeyword = new KeywordOptions
                {
                    Keyword = "PASPAYI SIYIRMA YÖNTEMİ İLE TESPİT EDİLEN DONATILAR İÇİN",
                    Occurrence = KeywordOccurrence.Last,
                    IncludeMatchingPage = false
                },
                EndPageSelectionType = PageSelectionType.LastPage,
            };

            service.ExtractRange(
                file,
                Path.Combine(outputFolder, "RONTGEN.pdf"),
                rontgenOptions
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