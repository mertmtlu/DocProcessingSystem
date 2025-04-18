using DocProcessingSystem.Core;
using DocProcessingSystem.Models;
using DocProcessingSystem.Services;
using iText.Kernel.Pdf.Filters;
using iTextSharp.text.pdf;
using Microsoft.VisualBasic;
using OfficeOpenXml;
using Org.BouncyCastle.Asn1.Cmp;
using System;
using System.Diagnostics.CodeAnalysis;

namespace DocProcessingSystem
{
    public class Program
    {
        #region Application Entry Point

        /// <summary>
        /// Main application entry method
        /// </summary>
        static void Main(string[] args)
        {
            //PdfOperationsHelper.ConvertWordToPdfAsync(@"C:\Users\Mert\Downloads\KAROT ÇALIŞMALARI\5. Bölge Karot Ekleri", @"C:\Users\Mert\Downloads\KAROT ÇALIŞMALARI\5. Bölge Karot Ekleri", false, true).Wait();
            //PdfOperationsHelper.ProcessPdfDocuments();
            //GetTotalPageCount(@"C:\Users\Mert\Desktop\REPORTS");
            //CreateEkA(@"C:\Users\Mert\Desktop\fırat eka\TM FOLDERS - Kopya");
            //ProcessDocuments();
            //RenameDocumentFiles();
            //HandleCrisis();
            //SortPdfFiles();
            //CheckAndFixHakedisFolder();
            TryChangeText();
        }

        #endregion

        #region Document Processing Functions

        static void TryChangeText()
        {
            string inputFolder = @"C:\Users\Mert\Desktop\Analysis - Kopya";
            string outputFolder = @"C:\Users\Mert\Desktop\Analysis";

            var ekKarot = Directory.GetFiles(inputFolder, "EK_KAROT.pdf", SearchOption.AllDirectories);
            var ekB = Directory.GetFiles(inputFolder, "EK-B.pdf", SearchOption.AllDirectories);
            var textReplacer = new PdfTextReplacerService();

            foreach (var file in ekKarot)
            {
                textReplacer.ReplaceCapYukseklik(
                    file,
                    file.Replace(inputFolder, outputFolder)
                );
            }

            foreach (var file in ekB)
            {
                textReplacer.ReplaceCapYukseklik(
                    file,
                    file.Replace(inputFolder, outputFolder)
                );
            }

            //foreach (var file in allWordFiles)
            //{
            //    var dest = file.Replace(inputFolder, outputFolder).Replace(".docx", "_yukseklik_cap_hatali.docx");

            //    // Get the directory path from the destination file path
            //    string destDirectory = Path.GetDirectoryName(dest);

            //    // Create the directory if it doesn't exist
            //    if (!Directory.Exists(destDirectory))
            //    {
            //        Directory.CreateDirectory(destDirectory);
            //    }

            //    File.Copy(file, dest);
            //}
        }

        static void CheckHakedisFolder()
        {
            string rootFolder = @"C:\Users\Mert\Desktop\HK18 (FINAL)";

            var groups = FolderHelper.GroupFolders(rootFolder);

            foreach (var group in groups)
            {
                // Check required files exists first.
                var requiredFiles = Constants.requiredFiles[group.PathCount];

                foreach (var requiredFile in requiredFiles)
                {
                    if (!File.Exists(Path.Combine(group.MainFolder, "TBDYResults", requiredFile)))
                    {
                        Console.WriteLine($"{requiredFile} not found for: TM No: {group.TmNo}, Building Code: {group.BuildingCode}, Building TM ID: {group.BuildingTmId}");
                    }
                }

                // Check if EK-C.pdf and main pdf file has wrong cover page
                using (var reader = new PdfReaderService())
                {
                    try
                    {
                        // Check EK-C.pdf
                        var ekCPath = Path.Combine(group.MainFolder, "TBDYResults", "EK-C.pdf");
                        if (File.Exists(ekCPath))
                        {
                            var pagesForEkC = reader.ExtractTextFromAllPages(ekCPath);
                            var firstPageForEkC = pagesForEkC[1];

                            // Check for the proper EK-C cover page format (Image 2)
                            // This recognizes both potential formats with some flexibility for whitespace/formatting
                            bool hasCorrectCover = firstPageForEkC.Contains("EK-C") &&
                                                  (firstPageForEkC.Contains("Donatı Tespiti için Ferroscan") ||
                                                   firstPageForEkC.Contains("DONATI TESPİTİ İÇİN FERROSCAN")) &&
                                                  (firstPageForEkC.Contains("ve Sıyırma Sonuçları") ||
                                                   firstPageForEkC.Contains("VE SIYIRMA SONUÇLARI"));

                            // Make sure we're not just finding the table of contents (Image 1)
                            bool isTableOfContents = firstPageForEkC.Contains("EKLER") &&
                                                    firstPageForEkC.Contains("EK-A") &&
                                                    firstPageForEkC.Contains("EK-B") &&
                                                    firstPageForEkC.Contains("EK-D");

                            if (!hasCorrectCover || isTableOfContents)
                            {
                                Console.WriteLine($"Required revision - EK-C file has wrong cover page: TM No: {group.TmNo}, Building Code: {group.BuildingCode}, Building TM ID: {group.BuildingTmId}");
                            }
                            else
                            {
                                Console.WriteLine($"EK-C file has correct cover page on page 1");
                            }

                            // Check main PDF
                            var tmNoSplited = group.TmNo.Split("-");
                            var mainPdfName = $"TEI-B{tmNoSplited[0]}-TM-{tmNoSplited[1]}-DIR-M{group.BuildingCode}-{group.BuildingTmId}.pdf";
                            var mainPdfPath = Path.Combine(group.MainFolder, "TBDYResults", mainPdfName);

                            if (File.Exists(mainPdfPath))
                            {
                                var pagesForMainPdf = reader.ExtractTextFromAllPages(mainPdfPath);
                                bool foundCorrectCoverInMain = false;
                                int correctCoverPageNumber = -1;
                                bool foundWrongCover = false;
                                int wrongCoverPageNumber = -1;

                                // Check each page of the main PDF for EK-C content
                                foreach (var pageEntry in pagesForMainPdf)
                                {
                                    int pageNumber = pageEntry.Key;
                                    string pageText = pageEntry.Value;

                                    // If this looks like an EK-C cover page (not just a TOC reference)
                                    if (pageText.Contains("EK-C") &&
                                        !(pageText.Contains("EKLER") && pageText.Contains("EK-A") && pageText.Contains("EK-B") && pageText.Contains("EK-D")))
                                    {
                                        // Check if it's the correct format (should look like Image 2 not Image 1)
                                        bool hasCorrectCoverOnMain = (pageText.Contains("Donatı Tespiti için Ferroscan") ||
                                                               pageText.Contains("DONATI TESPİTİ İÇİN FERROSCAN")) &&
                                                              (pageText.Contains("ve Sıyırma Sonuçları") ||
                                                               pageText.Contains("VE SIYIRMA SONUÇLARI"));

                                        if (hasCorrectCoverOnMain)
                                        {
                                            foundCorrectCoverInMain = true;
                                            correctCoverPageNumber = pageNumber;
                                        }
                                        else
                                        {
                                            foundWrongCover = true;
                                            wrongCoverPageNumber = pageNumber;
                                        }
                                    }
                                }

                                // Report findings with color coding
                                ConsoleColor defaultColor = Console.ForegroundColor;

                                if (foundCorrectCoverInMain)
                                {
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine($"Main PDF contains correct EK-C cover page on page {correctCoverPageNumber}");
                                    Console.ForegroundColor = defaultColor;
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("Main PDF does not contain the correct EK-C cover page");
                                    Console.ForegroundColor = defaultColor;
                                }

                                if (foundWrongCover)
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine($"Main PDF contains wrong EK-C cover page on page {wrongCoverPageNumber}: {mainPdfName}");
                                    Console.ForegroundColor = defaultColor;
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Warning: Main PDF file not found: {mainPdfPath}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Warning: EK-C file not found: {ekCPath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error checking PDF files: {ex.Message}");
                    }
                }
            }
        }

        static void CheckAndFixHakedisFolder()
        {
            string rootFolder = @"C:\Users\Mert\Desktop\HK18 (FINAL) - Kopya";
            string projectRootPath = AppDomain.CurrentDomain.BaseDirectory;
            string deterministicCoverPath = Path.Combine(projectRootPath, "CoverPages", "Deterministic");

            // Get the correct EK-C cover page template
            string correctEkCCoverTemplate = Path.Combine(deterministicCoverPath, "KAPAK_EK-C.pdf");
            if (!File.Exists(correctEkCCoverTemplate))
            {
                Console.WriteLine($"ERROR: Correct EK-C cover template not found at: {correctEkCCoverTemplate}");
                return;
            }

            var groups = FolderHelper.GroupFolders(rootFolder);

            foreach (var group in groups)
            {
                // Check required files exists first
                var requiredFiles = Constants.requiredFiles[group.PathCount];

                foreach (var requiredFile in requiredFiles)
                {
                    if (!File.Exists(Path.Combine(group.MainFolder, "TBDYResults", requiredFile)))
                    {
                        Console.WriteLine($"{requiredFile} not found for: TM No: {group.TmNo}, Building Code: {group.BuildingCode}, Building TM ID: {group.BuildingTmId}");
                    }
                }

                // Check if EK-C.pdf and main pdf file has wrong cover page
                using (var reader = new PdfReaderService())
                {
                    try
                    {
                        // Check EK-C.pdf
                        var ekCPath = Path.Combine(group.MainFolder, "TBDYResults", "EK-C.pdf");
                        if (File.Exists(ekCPath))
                        {
                            var pagesForEkC = reader.ExtractTextFromAllPages(ekCPath);
                            var firstPageForEkC = pagesForEkC[1];

                            // Check for the proper EK-C cover page format
                            bool hasCorrectCover = firstPageForEkC.Contains("EK-C") &&
                                                  (firstPageForEkC.Contains("Donatı Tespiti için Ferroscan") ||
                                                   firstPageForEkC.Contains("DONATI TESPİTİ İÇİN FERROSCAN")) &&
                                                  (firstPageForEkC.Contains("ve Sıyırma Sonuçları") ||
                                                   firstPageForEkC.Contains("VE SIYIRMA SONUÇLARI"));

                            // Make sure we're not just finding the table of contents
                            bool isTableOfContents = firstPageForEkC.Contains("EKLER") &&
                                                    firstPageForEkC.Contains("EK-A") &&
                                                    firstPageForEkC.Contains("EK-B") &&
                                                    firstPageForEkC.Contains("EK-D");

                            bool needToFixEkC = (!hasCorrectCover || isTableOfContents);

                            if (needToFixEkC)
                            {
                                ConsoleColor defaultColor = Console.ForegroundColor;
                                Console.ForegroundColor = ConsoleColor.Yellow;
                                Console.WriteLine($"Fixing wrong cover page in EK-C file: TM No: {group.TmNo}, Building Code: {group.BuildingCode}, Building TM ID: {group.BuildingTmId}");
                                Console.ForegroundColor = defaultColor;

                                // Create a backup of the original file
                                string backupPath = ekCPath + ".backup";
                                if (File.Exists(backupPath))
                                    File.Delete(backupPath);
                                File.Copy(ekCPath, backupPath);

                                // Fix the EK-C file
                                FixPdfCoverPage(ekCPath, correctEkCCoverTemplate);

                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine($"Successfully fixed EK-C cover page: {ekCPath}");
                                Console.ForegroundColor = defaultColor;
                            }
                            else
                            {
                                Console.WriteLine($"EK-C file has correct cover page on page 1");
                            }

                            // Check main PDF
                            var tmNoSplited = group.TmNo.Split("-");
                            var mainPdfName = $"TEI-B{tmNoSplited[0]}-TM-{tmNoSplited[1]}-DIR-M{group.BuildingCode}-{group.BuildingTmId}.pdf";
                            var mainPdfPath = Path.Combine(group.MainFolder, "TBDYResults", mainPdfName);

                            if (File.Exists(mainPdfPath))
                            {
                                var pagesForMainPdf = reader.ExtractTextFromAllPages(mainPdfPath);
                                bool foundCorrectCoverInMain = false;
                                int correctCoverPageNumber = -1;
                                bool foundWrongCover = false;
                                List<int> wrongCoverPageNumbers = new List<int>();

                                // Check each page of the main PDF for EK-C content
                                foreach (var pageEntry in pagesForMainPdf)
                                {
                                    int pageNumber = pageEntry.Key;
                                    string pageText = pageEntry.Value;

                                    // If this looks like an EK-C cover page (not just a TOC reference)
                                    if (pageText.Contains("EK-C") &&
                                        !(pageText.Contains("EKLER") && pageText.Contains("EK-A") && pageText.Contains("EK-B") && pageText.Contains("EK-D")))
                                    {
                                        // Check if it's the correct format
                                        bool hasCorrectCoverOnMain = (pageText.Contains("Donatı Tespiti için Ferroscan") ||
                                                               pageText.Contains("DONATI TESPİTİ İÇİN FERROSCAN")) &&
                                                              (pageText.Contains("ve Sıyırma Sonuçları") ||
                                                               pageText.Contains("VE SIYIRMA SONUÇLARI"));

                                        if (hasCorrectCoverOnMain)
                                        {
                                            foundCorrectCoverInMain = true;
                                            correctCoverPageNumber = pageNumber;
                                        }
                                        else
                                        {
                                            foundWrongCover = true;
                                            wrongCoverPageNumbers.Add(pageNumber);
                                        }
                                    }
                                }

                                // Report findings with color coding
                                ConsoleColor defaultColor = Console.ForegroundColor;

                                Console.WriteLine($"Checking for: TM No: {group.TmNo}, Building Code: {group.BuildingCode}, Building TM ID: {group.BuildingTmId}");

                                if (foundCorrectCoverInMain)
                                {
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine($"Main PDF contains correct EK-C cover page on page {correctCoverPageNumber}");
                                    Console.ForegroundColor = defaultColor;
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("Main PDF does not contain the correct EK-C cover page");
                                    Console.ForegroundColor = defaultColor;
                                }

                                if (foundWrongCover)
                                {
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.WriteLine($"Main PDF contains wrong EK-C cover page on pages: {string.Join(", ", wrongCoverPageNumbers)}");
                                    Console.WriteLine($"To fix the main PDF, regenerate it using the fixed EK-C.pdf file");
                                    Console.ForegroundColor = defaultColor;
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Warning: Main PDF file not found: {mainPdfPath}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Warning: EK-C file not found: {ekCPath}");
                        }

                        Console.WriteLine("");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error checking PDF files: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Fixes a PDF by replacing its cover page with the correct template
        /// </summary>
        /// <param name="pdfToFix">Path to the PDF that needs its cover page fixed</param>
        /// <param name="coverTemplatePath">Path to the correct cover page template</param>
        static void FixPdfCoverPage(string pdfToFix, string coverTemplatePath)
        {
            try
            {
                // Create a temporary directory for intermediate files
                string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDir);

                string contentPdfPath = Path.Combine(tempDir, "content.pdf");

                try
                {
                    // 1. Extract all content from the PDF except the first page
                    using (var reader = new PdfReaderService())
                    {
                        int pageCount = reader.ExtractTextFromAllPages(pdfToFix).Count;

                        if (pageCount <= 1)
                        {
                            // If there's only one page, we need a different approach
                            // We'll just replace the entire file with the cover template
                            File.Copy(coverTemplatePath, pdfToFix, true);
                            return;
                        }

                        // Use PdfRangeExtractorService to extract all pages except the first one
                        var extractionOptions = new PdfExtractionOptions
                        {
                            StartPageSelectionType = PageSelectionType.SpecificPage,
                            StartPageNumber = 2, // Start from second page
                            EndPageSelectionType = PageSelectionType.LastPage
                        };

                        var extractor = new PdfRangeExtractorService();
                        extractor.ExtractRange(pdfToFix, contentPdfPath, extractionOptions);
                    }

                    // 2. Merge the cover template with the content
                    var mergeOptions = new MergeOptions
                    {
                        PreserveBookmarks = true,
                        CreateBookmarksForAdditionalPdf = false
                    };

                    using (var merger = new PdfMergerService())
                    {
                        var additionalDocs = new List<string> { contentPdfPath };
                        merger.MergePdf(coverTemplatePath, additionalDocs, pdfToFix, mergeOptions);
                    }
                }
                finally
                {
                    // Clean up temporary files
                    try
                    {
                        if (Directory.Exists(tempDir))
                            Directory.Delete(tempDir, true);
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fixing PDF cover page: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Regenerates the main PDF file by merging all component PDFs including the fixed EK-C
        /// </summary>
        static void RegenerateMainPdf(string mainPdfPath, List<string> componentPdfs, MergeOptions options)
        {
            try
            {
                if (!componentPdfs.Any())
                {
                    Console.WriteLine("No component PDFs provided for regeneration.");
                    return;
                }

                string mainDoc = componentPdfs.First();
                var additionalDocs = componentPdfs.Skip(1).ToList();

                using (var merger = new PdfMergerService())
                {
                    merger.MergePdf(mainDoc, additionalDocs, mainPdfPath, options);
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Successfully regenerated main PDF: {mainPdfPath}");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error regenerating main PDF: {ex.Message}");
                Console.ForegroundColor = ConsoleColor.Gray;
                throw;
            }
        }

        static void ProcessDocuments()
        {
            // Get folder paths from arguments or use defaults
            string parametricsFolder = @"C:\Users\Mert\Desktop\SZL2\Anıl Report Revision\archive\Anıl Final Report Merge\Parametric";
            string deterministicsFolder = @"C:\Users\Mert\Desktop\SZL2\Anıl Report Revision\archive\Anıl Final Report Merge\Deterministic";
            string post2008 = @"C:\Users\Mert\Desktop\Fırat Report Revision\MM_RAPOR\WORDasd"; // TODO: FIRAT
            string analysisFolder = @"C:\Users\Mert\Desktop\last revise folder\Analysis"; // TODO: FIRAT

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

        static void HandleCrisis()
        {
            var inputFolder = @"C:\Users\Mert\Desktop\HK-13";
            var outputFolder = @"C:\Users\Mert\Desktop\HK13 TM FOLDERS";

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
                catch (Exception e)
                {
                    Console.WriteLine($"ERROR: Cannot process file, Name: {Path.GetFileName(file)}, Path: {file}, Error: {e.Message}");
                }
            }
        }

        #endregion

        #region PDF Processing Functions

        /// <summary>
        /// Processes all subfolders in a root directory, merging PDFs in each subfolder
        /// </summary>
        /// <param name="rootFolderPath">The root folder containing subfolders with PDFs to merge</param>
        static void CreateEkA(string rootFolderPath)
        {
            if (string.IsNullOrEmpty(rootFolderPath))
                throw new ArgumentNullException(nameof(rootFolderPath), "Root folder path cannot be null or empty");

            if (!Directory.Exists(rootFolderPath))
                throw new DirectoryNotFoundException($"Root folder not found: {rootFolderPath}");

            // Get the main PDF path
            string projectRootPath = AppDomain.CurrentDomain.BaseDirectory;
            string deterministicCoverPath = Path.Combine(projectRootPath, "CoverPages", "Deterministic");
            string mainPdf = Path.Combine(deterministicCoverPath, "EK-A_Kapak.pdf");

            if (!File.Exists(mainPdf))
                throw new FileNotFoundException($"Main PDF file not found: {mainPdf}");

            // Set up merge options
            var options = new MergeOptions
            {
                PreserveBookmarks = false,
                CreateBookmarksForAdditionalPdf = false
            };

            // Get all subfolders in the root folder
            string[] subfolders = Directory.GetDirectories(rootFolderPath);
            Console.WriteLine($"Found {subfolders.Length} subfolders to process\n");

            using (var pdfMerger = new PdfMergerService())
            {
                foreach (string subfolder in subfolders)
                {
                    try
                    {
                        // Get all PDF files in the subfolder and sort them by name
                        string[] pdfFiles = Directory.GetFiles(subfolder, "*.pdf")
                                                     .OrderBy(f => Path.GetFileName(f))
                                                     .Where(f => !Path.GetFileNameWithoutExtension(f).Contains("EK-A"))
                                                     .ToArray();

                        if (pdfFiles.Length == 0)
                        {
                            Console.WriteLine($"No PDF files found in subfolder: {subfolder}");
                            continue;
                        }

                        Console.WriteLine($"Processing subfolder: {subfolder}");
                        Console.WriteLine($"Found {pdfFiles.Length} PDF files");

                        // Set the output path to be in the subfolder
                        string outputPath = Path.Combine(subfolder, "EK-A.pdf");

                        // Create the merge sequence
                        var mergeSequence = new MergeSequence
                        {
                            MainDocument = mainPdf,
                            AdditionalDocuments = pdfFiles.ToList(),
                            OutputPath = outputPath,
                            Options = options
                        };

                        // Perform the merge
                        Console.WriteLine($"Merging PDFs. Output path: {outputPath}");
                        pdfMerger.MergePdf(mergeSequence);

                        Console.WriteLine($"Successfully processed subfolder: {subfolder}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing subfolder {subfolder}: {ex.Message}");
                        // Continue with next subfolder instead of stopping the entire process
                    }

                    Console.WriteLine();
                }
            }

            Console.WriteLine("PDF merging process completed");
        }

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
                file,
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
        static void ProcessSubFolder(string inputFolderPath, string outputFolderPath)
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
        static void ProcessSingleSubfolder(string subfolderPath, string outputSubfolderPath, string mainPdf)
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


        #endregion

        #region File Renaming and Naming Functions

        static void SortPdfFiles()
        {
            string rootFolder = @"C:\Users\Mert\Desktop\REPORTS";
            string destinationFolder = @"C:\Users\Mert\Desktop\new pdfs";
            var files = Directory.GetFiles(rootFolder, "*.pdf", SearchOption.AllDirectories);
            List<string> preferences = new()
            {
                "TSU",
                "FAY", "GEO", "IKL", "ZEV",
                "SEL", "HEY", "CIG", "SES", "GUV", "SLT"
            };

            foreach (var file in files)
            {
                string tmNo = null;
                string buildingCode = null;
                string buildingTmId = null;
                bool processed = false;

                foreach (var pref in preferences)
                {
                    var result = FolderHelper.ExtractParts(file, pref);
                    tmNo = result.tmNo;
                    buildingCode = result.buildingCode;
                    buildingTmId = result.buildingTmId;

                    if (!string.IsNullOrEmpty(tmNo))
                    {
                        processed = true;
                        break; // Break out of preferences loop once we find a valid one
                    }
                }

                if (!processed)
                {
                    Console.WriteLine($"{file} cannot be processed");
                    continue; // Skip to next file
                }

                var areaID = Convert.ToInt32(tmNo.Split("-").First());
                var folder = Path.Combine(destinationFolder, $"BOLGE-{areaID}");

                // Create folder if it doesn't exist
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                // Copy file to destination folder
                string destinationFilePath = Path.Combine(folder, Path.GetFileName(file).Replace(".docx", ""));
                File.Copy(file, destinationFilePath, true); // true allows overwriting existing files
            }

            Console.WriteLine("Done.");
        }

        static void RenameDocumentFiles()
        {
            var inputFolder = @"C:\Users\Mert\Desktop\DIR";
            var excelFile = @"C:\Users\Mert\Desktop\SZL-2_TM_KISA_TR_ISIM_LISTE_20250319.xlsx";

            var tmNameJson = ConvertExcelToDictionary(excelFile);

            // Get both Word and PDF documents
            var wordDocuments = Directory.GetFiles(inputFolder, "*.docx", SearchOption.AllDirectories);
            var pdfDocuments = Directory.GetFiles(inputFolder, "*.pdf", SearchOption.AllDirectories);
            var allDocuments = wordDocuments.Concat(pdfDocuments).ToArray();

            foreach (var document in allDocuments)
            {
                // Get the file extension to preserve it in the renamed file
                string fileExtension = Path.GetExtension(document);

                if (document.Contains("M00"))
                {
                    foreach (var preference in Constants.preferences)
                    {
                        if (document.Contains($"-{preference}-"))
                        {
                            try
                            {
                                var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(document, preference);

                                // Get the shortened name for this TM number
                                var shortenedName = FindShortenedName(tmNo, tmNameJson)?.ToString();

                                if (shortenedName == null) throw new ArgumentNullException("Shortened Name Not Found.");

                                // Split the TM number to get area ID and TM ID
                                var areaId = tmNo.Split("-")[0];
                                var tmId = tmNo.Split("-")[1];

                                var newName = $"TEI-B{areaId}-TM-{tmId}-{preference}-M00-00_NT ({shortenedName}-{Constants.ReportType[preference]}){fileExtension}";

                                // Get the directory path from the original document
                                string directoryPath = Path.GetDirectoryName(document);

                                // Combine directory path with new filename
                                string newFilePath = Path.Combine(directoryPath, newName);

                                // Rename the file
                                if (File.Exists(newFilePath))
                                {
                                    Console.WriteLine($"Warning: A file with the name '{newName}' already exists. Skipping rename operation for {document}");
                                }
                                else
                                {
                                    File.Move(document, newFilePath);
                                    Console.WriteLine($"Successfully renamed: {Path.GetFileName(document)} -> {newName}");
                                }
                            }
                            catch (KeyNotFoundException)
                            {
                                Console.WriteLine($"Error: Could not find building code in dictionary for document: {document}");
                            }
                            catch (IndexOutOfRangeException)
                            {
                                Console.WriteLine($"Error: Invalid TM number format in document: {document}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error processing document {document}: {ex.Message}");
                            }
                        }
                    }
                }
                else
                {
                    try
                    {
                        // Extract information from the filename
                        var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(document);

                        // Get the shortened name for this TM number
                        var shortenedName = FindShortenedName(tmNo, tmNameJson)?.ToString();

                        if (shortenedName == null) throw new ArgumentNullException("Shortened Name Not Found.");

                        // Get the building name from our constants dictionary
                        var buildingCodeInt = Convert.ToInt32(buildingCode);
                        var buildingName = Constants.CodeToName[buildingCodeInt];

                        // Split the TM number to get area ID and TM ID
                        var areaId = tmNo.Split("-")[0];
                        var tmId = tmNo.Split("-")[1];

                        // Create the new filename with the required format
                        var newName = $"TEI-B{areaId}-TM-{tmId}-DIR-M{buildingCode}-{buildingTmId}_NT ({shortenedName}-{buildingName}){fileExtension}";

                        // Get the directory path from the original document
                        string directoryPath = Path.GetDirectoryName(document);

                        // Combine directory path with new filename
                        string newFilePath = Path.Combine(directoryPath, newName);

                        // Rename the file
                        if (File.Exists(newFilePath))
                        {
                            Console.WriteLine($"Warning: A file with the name '{newName}' already exists. Skipping rename operation for {document}");
                        }
                        else
                        {
                            File.Move(document, newFilePath);
                            Console.WriteLine($"Successfully renamed: {Path.GetFileName(document)} -> {newName}");
                        }
                    }
                    catch (KeyNotFoundException)
                    {
                        Console.WriteLine($"Error: Could not find building code in dictionary for document: {document}");
                    }
                    catch (IndexOutOfRangeException)
                    {
                        Console.WriteLine($"Error: Invalid TM number format in document: {document}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing document {document}: {ex.Message}");
                    }
                }
            }
        }

        static object FindShortenedName(string tmNo, List<Dictionary<string, object>> e)
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

        #endregion

        #region Data Processing and Helper Functions

        static List<Dictionary<string, object>> ConvertExcelToDictionary(string filePath, string sheetName = null)
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

        static void CopyDirectory(string sourceDir, string destinationDir)
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

        static void GetTotalPageCount(string folderPath)
        {
            int totalPages = 0;
            var pdfFiles = Directory.GetFiles(folderPath, "*.pdf", SearchOption.AllDirectories);
            Console.WriteLine($"Found {pdfFiles.Length} PDF files.");

            foreach (var pdfFile in pdfFiles)
            {
                totalPages += GetPdfPageCount(pdfFile);
            }

            Console.WriteLine($"TOTAL COUNT OF PAGES: {totalPages}");
        }

        static int GetPdfPageCount(string pathToPdf)
        {
            int totalPages = 0;
            try
            {
                using (PdfReader reader = new PdfReader(pathToPdf))
                {
                    int pageCount = reader.NumberOfPages;
                    totalPages += pageCount;
                    Console.WriteLine($"File: {Path.GetFileName(pathToPdf)}, Pages: {pageCount}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading {Path.GetFileName(pathToPdf)}: {ex.Message}");
            }

            return totalPages;
        }

        #endregion
    }
}