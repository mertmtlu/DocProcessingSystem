using DocProcessingSystem.Core;
using DocProcessingSystem.Models;
using DocProcessingSystem.Services;
using Microsoft.VisualBasic;
using OfficeOpenXml;

namespace DocProcessingSystem
{
    public class Program
    {
        ///// <summary>
        ///// Main application entry method
        ///// </summary>
        //static async Task Main(string[] args) // Async
        //{
        //     await ConvertWordToPdfAsync(@"C:\Users\Mert\Desktop\iklim raporu düzenleme\final\reports", @"C:\Users\Mert\Desktop\REPORTS\MERT IKLİM", false);
        //}

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

        public static void RenameWordFiles()
        {
            var inputFolder = @"C:\Users\Mert\Desktop\iklim raporu düzenleme\final\reports";
            var excelFile = @"C:\Users\Mert\Desktop\SZL-2_TM_KISA_TR_ISIM_LISTE_20250319.xlsx";

            var tmNameJson = ConvertExcelToDictionary(excelFile);

            var wordDocuments = Directory.GetFiles(inputFolder, "*.docx", SearchOption.AllDirectories);

            foreach (var wordDocument in wordDocuments)
            {
                if (wordDocument.Contains("M00"))
                {
                    List<string> preferences = new()
                    {
                        "IKL",
                        "GEO",
                        "FAY",
                        "ZEV"
                    };
                    Dictionary<string, string> mapper = new()
                    {
                        { "ZEV", "ZEMIN ETUT-VERI"},
                        { "GEO", "ZEMIN ETUT-GEOTEKNIK"},
                        { "FAY", "DIRIFAY"},
                        { "IKL", "IKLIM DEGISIKLIGI"},
                    };

                    foreach (var preference in preferences)
                    {
                        if (wordDocument.Contains($"-{preference}-"))
                        {
                            try
                            {
                                var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(wordDocument, preference);

                                // Get the shortened name for this TM number
                                var shortenedName = FindShortenedName(tmNo, tmNameJson)?.ToString();

                                if (shortenedName == null) throw new ArgumentNullException("Shortened Name Not Found.");

                                // Split the TM number to get area ID and TM ID
                                var areaId = tmNo.Split("-")[0];
                                var tmId = tmNo.Split("-")[1];

                                var newName = $"TEI-B{areaId}-TM-{tmId}-{preference}-M00-00_NT ({shortenedName}-{mapper[preference]}).docx";

                                // Get the directory path from the original document
                                string directoryPath = Path.GetDirectoryName(wordDocument);

                                // Combine directory path with new filename
                                string newFilePath = Path.Combine(directoryPath, newName);

                                // Rename the file
                                if (File.Exists(newFilePath))
                                {
                                    Console.WriteLine($"Warning: A file with the name '{newName}' already exists. Skipping rename operation for {wordDocument}");
                                }
                                else
                                {
                                    File.Move(wordDocument, newFilePath);
                                    Console.WriteLine($"Successfully renamed: {Path.GetFileName(wordDocument)} -> {newName}");
                                }
                            }
                            catch (KeyNotFoundException)
                            {
                                Console.WriteLine($"Error: Could not find building code in dictionary for document: {wordDocument}");
                            }
                            catch (IndexOutOfRangeException)
                            {
                                Console.WriteLine($"Error: Invalid TM number format in document: {wordDocument}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error processing document {wordDocument}: {ex.Message}");
                            }
                        }
                    }
                }
                else
                {
                    try
                    {
                        // Extract information from the filename
                        var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(wordDocument);

                        // Get the shortened name for this TM number
                        var shortenedName = FindShortenedName(tmNo, tmNameJson)?.ToString();

                        if (shortenedName == null) throw new ArgumentNullException("Shortened Name Not Found.");

                        // Get the building name from our constants dictionary
                        var buildingName = Constants.CodeToName[Convert.ToInt32(buildingCode)];

                        // Split the TM number to get area ID and TM ID
                        var areaId = tmNo.Split("-")[0];
                        var tmId = tmNo.Split("-")[1];

                        // Create the new filename with the required format
                        var newName = $"TEI-B{areaId}-TM-{tmId}-DIR-M{buildingCode}-{buildingTmId}_NT ({shortenedName}-{buildingName}).docx";

                        // Get the directory path from the original document
                        string directoryPath = Path.GetDirectoryName(wordDocument);

                        // Combine directory path with new filename
                        string newFilePath = Path.Combine(directoryPath, newName);

                        // Rename the file
                        if (File.Exists(newFilePath))
                        {
                            Console.WriteLine($"Warning: A file with the name '{newName}' already exists. Skipping rename operation for {wordDocument}");
                        }
                        else
                        {
                            File.Move(wordDocument, newFilePath);
                            Console.WriteLine($"Successfully renamed: {Path.GetFileName(wordDocument)} -> {newName}");
                        }
                    }
                    catch (KeyNotFoundException)
                    {
                        Console.WriteLine($"Error: Could not find building code in dictionary for document: {wordDocument}");
                    }
                    catch (IndexOutOfRangeException)
                    {
                        Console.WriteLine($"Error: Invalid TM number format in document: {wordDocument}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing document {wordDocument}: {ex.Message}");
                    }
                }
            }
        }

        public static object FindShortenedName(string tmNo, List<Dictionary<string, object>> e)
        {
            foreach (var item in e)
            {
                var tmKey = item.Keys.FirstOrDefault(a => a.Contains("TM"));

                var value = item[tmKey]?.ToString();

                if (tmKey != null && value == tmNo)
                {
                    var shortenedNameKey = item.Keys.FirstOrDefault(a => a.Contains("KISA TR İSİM (DOSYA ADLANDIRMA İÇİN)"));

                    if (shortenedNameKey != null)
                    {
                        return item[shortenedNameKey];
                    }
                }
            }

            // If no match is found, return null
            return null;
        }

        public static List<Dictionary<string, object>> ConvertExcelToDictionary(string filePath, string sheetName = null)
        {
            // Set the license context (required for EPPlus 5+)
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            var result = new List<Dictionary<string, object>>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Get the specified worksheet, or the first one if not specified
                ExcelWorksheet worksheet = sheetName != null
                    ? package.Workbook.Worksheets[sheetName]
                    : package.Workbook.Worksheets[0];

                // Check if worksheet exists
                if (worksheet == null)
                {
                    throw new ArgumentException($"Worksheet '{sheetName}' not found.");
                }

                // Determine the dimensions of the worksheet
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                // Get the header row (property names)
                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    string header = worksheet.Cells[1, col].Value?.ToString();
                    if (!string.IsNullOrEmpty(header))
                    {
                        headers.Add(header);
                    }
                }

                // Process each row
                for (int row = 2; row < rowCount; row++) // Start from row 2 (after header)
                {
                    var rowDict = new Dictionary<string, object>();

                    for (int col = 1; col <= headers.Count; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;
                        if (col <= headers.Count) // Ensure we don't go out of bounds
                        {
                            rowDict[headers[col - 1]] = cellValue;
                        }
                    }

                    result.Add(rowDict);
                }
            }

            return result;
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

                    converter.Convert(file, outputPath, saveChanges, false);
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
        
        public static async Task ConvertWordToPdfAsync(string inputFolderPath, string outputFolderPath, bool saveChanges, int maxParallel = 8)
        {
            // Get all Word files asynchronously
            var wordFiles = await DirectoryExtensions.GetFilesAsync(inputFolderPath, "*.docx", SearchOption.AllDirectories);

            // Create progress reporting
            var progress = new Progress<(string FileName, int Completed, int Total)>(update =>
            {
                Console.WriteLine($"Converted {update.FileName} - {update.Completed} of {update.Total} completed");
            });

            // Use cancellation token to support cancellation
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


        static void ProcessDocuments()
        {
            // Get folder paths from arguments or use defaults
            string parametricsFolder = @"C:\Users\Mert\Desktop\testing\Parametric";
            string deterministicsFolder = @"C:\Users\Mert\Desktop\testing\Deterministic";
            string post2008 = @"C:\Users\Mert\Desktop\Fırat Report Revision\MM_RAPOR\WORD"; // TODO: FIRAT
            string analysisFolder = @"C:\Users\Mert\Desktop\Fırat Report Revision\MM_RAPOR\ANALİZ"; // TODO: FIRAT

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