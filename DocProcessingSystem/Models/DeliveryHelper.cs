using DocProcessingSystem.Core;
using DocProcessingSystem.Services;
using OfficeOpenXml.Table.PivotTable;
using Org.BouncyCastle.Asn1.Cmp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace DocProcessingSystem.Models
{
    public static class DeliveryHelper
    {
        public static void CheckRedFolderStructure(string root)
        {
            var subfolders = Directory.GetDirectories(root, "*", SearchOption.TopDirectoryOnly);

            foreach (var subfolder in subfolders)
            {
                var folderName = Path.GetFileName(subfolder);

                var areaID = folderName.Split('-')[0];
                var tmID = folderName.Split("-")[1];

                var pdfs = Directory.GetFiles(subfolder, "TEI*.pdf", SearchOption.TopDirectoryOnly);

                foreach (var pdf in pdfs)
                {
                    var pdfFileName = Path.GetFileNameWithoutExtension(pdf);
                    var fileStartPatternt = $"TEI-B{areaID}-TM-{tmID}-";

                    if (!pdfFileName.StartsWith(fileStartPatternt))
                    {
                        Console.WriteLine($"Mismatch found in folder '{folderName}': PDF file '{pdfFileName}' does not start with '{fileStartPatternt}'");
                    }
                }
            }
        }

        public static List<string> GetReports(string root)
        {
            return Directory.GetFiles(root, "TEI*.pdf", SearchOption.AllDirectories).ToList();
        }

        public static void RearrangeRed(string root, string dest, string main)
        {
            var pdfs = Directory.GetFiles(root, "*.pdf", SearchOption.AllDirectories);

            foreach (var pdf in pdfs)
            {
                (string tmNo, string buildingCode, string buildingTmId) = ExtractParts(Path.GetFileNameWithoutExtension(pdf), "RED-M");

                var mainPdf = Path.Combine(main, tmNo, "main.pdf");

                int mainPdfPages = GetPageCount(mainPdf) + 2;
                int totalPages = GetPageCount(pdf);

                var (startPageColoredFirstHalf, endPageColoredFirstHalf) = (1, mainPdfPages);
                var (startPageColoredSecondHalf, endPageColoredSecondHalf) = (totalPages - mainPdfPages + 1, totalPages);

                var (startPageBlackAndWhite, endPageBlackAndWhite) = (endPageColoredFirstHalf + 1, startPageColoredSecondHalf - 1);

                var pdfExtractor = new PdfRangeExtractorService();
                var destFolder = Path.Combine(dest, tmNo);
                Directory.CreateDirectory(destFolder);
                var tempFolder = Path.Combine(destFolder, "temp");
                Directory.CreateDirectory(tempFolder);
                var coloredFirstHalfPath = Path.Combine(tempFolder, "colored_first_half.pdf");
                var coloredSecondHalfPath = Path.Combine(tempFolder, "colored_second_half.pdf");
                var blackAndWhitePath = Path.Combine(destFolder, "black_and_white.pdf");


                var coloredFirstHalfOptions = new PdfExtractionOptions
                {
                    StartPageSelectionType = PageSelectionType.SpecificPage,
                    StartPageNumber = startPageColoredFirstHalf,
                    EndPageSelectionType = PageSelectionType.SpecificPage,
                    EndPageNumber = endPageColoredFirstHalf,
                };

                pdfExtractor.ExtractRange(pdf, coloredFirstHalfPath, coloredFirstHalfOptions);

                var blackAndWhiteOptions = new PdfExtractionOptions
                {
                    StartPageSelectionType = PageSelectionType.SpecificPage,
                    StartPageNumber = startPageBlackAndWhite,
                    EndPageSelectionType = PageSelectionType.SpecificPage,
                    EndPageNumber = endPageBlackAndWhite,
                };

                pdfExtractor.ExtractRange(pdf, blackAndWhitePath, blackAndWhiteOptions);

                var coloredSecondHalfOptions = new PdfExtractionOptions
                {
                    StartPageSelectionType = PageSelectionType.SpecificPage,
                    StartPageNumber = startPageColoredSecondHalf,
                    EndPageSelectionType = PageSelectionType.SpecificPage,
                    EndPageNumber = endPageColoredSecondHalf,
                };

                pdfExtractor.ExtractRange(pdf, coloredSecondHalfPath, coloredSecondHalfOptions);

                PdfMergerService merger = new PdfMergerService();

                var mergeOption = new MergeOptions
                {
                    PreserveBookmarks = true,
                    CreateBookmarksForAdditionalPdf = false,
                };

                var coloredPath = Path.Combine(destFolder, "colored.pdf");

                var firstMergeSequence = new MergeSequence
                {
                    MainDocument = coloredFirstHalfPath,
                    AdditionalDocuments = new List<string> { coloredSecondHalfPath },
                    OutputPath = coloredPath,
                    Options = mergeOption
                };

                merger.MergePdf(firstMergeSequence);

                Directory.Delete(tempFolder, true);
            }
        }

        public static int GetPageCount(string pdfFile)
        {
            using (var reader = new PdfReader(pdfFile))
            using (var sourceDoc = new PdfDocument(reader))
            {
                return sourceDoc.GetNumberOfPages();
            }
        }

        public static (string tmNo, string buildingCode, string buildingTmId) ExtractParts(string folderName, string preferance)
        {
            // Standard pattern: digits-digits-M+digits(-digits)
            string patternStandard = @"^(\d{1,2}-\d{2})\s*-?M(\d{2})(?:-(\d{2}|\d{1}))?(?:-([A-Za-z0-9]+))?$";

            // TEI pattern: TEI-B+digits-TM-digits-DIR-M+digits(-digits)
            string patternTei = $@"TEI-B(\d{{2}})-TM-(\d{{2}})-{preferance}(\d{{2}})(?:-(\d{{2}}|\d{{1}}))?";

            // Try the standard pattern first
            Match match = Regex.Match(folderName, patternStandard);
            if (match.Success)
            {
                string tmNo = match.Groups[1].Value;                   // e.g., "18-10"
                string buildingCode = match.Groups[2].Value;           // e.g., "02"
                string buildingTmId = match.Groups[3].Success
                    ? match.Groups[3].Value
                    : "01";                                            // Default to 01 if not specified

                return (tmNo, buildingCode, buildingTmId);
            }

            // Try the TEI pattern
            match = Regex.Match(folderName, patternTei);
            if (match.Success)
            {
                string buildingCode = match.Groups[3].Value;           // e.g., "02"
                string tmNo = $"{match.Groups[1].Value}-{match.Groups[2].Value}"; // e.g., "05-13"
                string buildingTmId = match.Groups[4].Success
                    ? match.Groups[4].Value
                    : "01";                                            // Default to 01 if not specified

                return (tmNo, buildingCode, buildingTmId);
            }

            return (null, null, null);
        }

        public static void GetEkParts(string root, string dest)
        {
            var collection = GroupReports(root);

            var pdfExtractor = new PdfRangeExtractorService();

            List<string> deleted = new() { "01-05", "01-36", "01-39", "05-04", "05-21", "12-21", "12-30", "12-58", "13-02", "13-03", "18-03", "18-07", "18-41" };


            List<ReportEnum> directCopy = new()
            {
                //ReportEnum.FAYM,
                //ReportEnum.FOYM,
                //ReportEnum.ALTA,
                //ReportEnum.IKLM,
                //ReportEnum.FOYA,
            };

            Dictionary<ReportEnum, string> endKeywordExcluded = new()
            {
                //{ReportEnum.CIGM, "EK-A TESİS"},
                {ReportEnum.GUVM, "EK-A TESİS"},
                //{ReportEnum.HEYM, "EK-A TESİS"},
                //{ReportEnum.SELM, "EK-A TESİS"},
                //{ReportEnum.SESM, "EK-A TESİS"},
                //{ReportEnum.YANM, "EK-A TESİS"},
                //{ReportEnum.TSUM, "EK-A TESİS"},

                //{ReportEnum.SLTM, "EK-A" }
            };

            Dictionary<ReportEnum, string> endKeywordIncluded = new()
            {
                //{ReportEnum.DIRM, "SONUÇ VE ÖNERİLER"},
                //{ReportEnum.DGRM, "SONUÇ VE ÖNERİLER"},

            };

            foreach (var type in endKeywordExcluded)
            {
                var reportType = type.Key;
                var keyword = type.Value;

                var mainDocumentOptions = new PdfExtractionOptions
                {
                    StartPageSelectionType = PageSelectionType.FirstPage,
                    EndPageSelectionType = PageSelectionType.Keyword,
                    EndKeyword = new KeywordOptions
                    {
                        Keyword = keyword,
                        Occurrence = KeywordOccurrence.Last,
                        IncludeMatchingPage = false,
                    },
                };

                foreach (var item in collection.Groups)
                {
                    if (deleted.Contains(item.Identifier)) continue;

                    var destinationFolder = Path.Combine(dest, item.Identifier);

                    foreach (var report in item.GetReportsByType(reportType))
                    {
                        var destinationFile = Path.Combine(destinationFolder, report.FileName);
                        pdfExtractor.ExtractRange(report.FilePath, destinationFile, mainDocumentOptions);
                    }
                }
            }

            foreach (var type in endKeywordIncluded)
            {
                var reportType = type.Key;
                var keyword = type.Value;

                var mainDocumentOptions = new PdfExtractionOptions
                {
                    StartPageSelectionType = PageSelectionType.FirstPage,
                    EndPageSelectionType = PageSelectionType.Keyword,
                    EndKeyword = new KeywordOptions
                    {
                        Keyword = keyword,
                        Occurrence = KeywordOccurrence.Last,
                        IncludeMatchingPage = true,
                    },
                };

                foreach (var item in collection.Groups)
                {
                    if (deleted.Contains(item.Identifier)) continue;

                    var destinationFolder = Path.Combine(dest, item.Identifier);

                    foreach (var report in item.GetReportsByType(reportType))
                    {
                        var destinationFile = Path.Combine(destinationFolder, report.FileName);
                        pdfExtractor.ExtractRange(report.FilePath, destinationFile, mainDocumentOptions);
                    }
                }
            }

            foreach (var type in directCopy)
            {
                foreach (var item in collection.Groups)
                {
                    if (deleted.Contains(item.Identifier)) continue;

                    var destinationFolder = Path.Combine(dest, item.Identifier);

                    foreach (var report in item.GetReportsByType(type))
                    {
                        var destinationFile = Path.Combine(destinationFolder, report.FileName);
                        File.Copy(report.FilePath, destinationFile, true);
                    }
                }
            }


        }

        public static void RunPython(string pythonFile)
        {
            try
            {
                // Check if the Python file exists
                if (!File.Exists(pythonFile))
                {
                    Console.WriteLine($"Error: Python file '{pythonFile}' not found.");
                    return;
                }

                // Create process start info
                ProcessStartInfo startInfo = new ProcessStartInfo()
                {
                    FileName = "python",              // or "python3" on some systems
                    Arguments = $"\"{pythonFile}\"",  // Wrap in quotes to handle spaces in path
                    UseShellExecute = false,          // Required for redirecting output
                    RedirectStandardOutput = true,    // Capture standard output
                    RedirectStandardError = true,     // Capture error output
                    CreateNoWindow = true             // Don't create a console window
                };

                // Start the process
                using (Process process = Process.Start(startInfo))
                {
                    if (process != null)
                    {
                        // Read output and error streams
                        string output = process.StandardOutput.ReadToEnd();
                        string error = process.StandardError.ReadToEnd();

                        // Wait for the process to exit
                        process.WaitForExit();

                        Console.WriteLine($"Python script exited with code: {process.ExitCode}");
                    }
                    else
                    {
                        Console.WriteLine("Failed to start Python process.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error running Python script: {ex.Message}");
            }
        }

        public static void CreateMainPdf(string folder)
        {
            var mainPdfs = Directory.GetFiles(folder, "main.pdf", SearchOption.AllDirectories);
        }

        public static void CreateRedReport(string root, string dest)
        {
            var collection = GroupReports(root);

            //using (var converter = new WordToPdfConverter())
            //{
            //    foreach (var item in collection.Groups)
            //    {
            //        var mainFile = Path.Combine(collection.RootDir, item.Identifier, "main.docx");
            //        var outputFile = Path.Combine(collection.RootDir, item.Identifier, "main.pdf");

            //        converter.Convert(mainFile, outputFile, true, false);
            //    }
            //}


            List<ReportEnum> mergeOrder = new()
    {
        ReportEnum.DIRM,
        ReportEnum.DGRM,
        ReportEnum.FAYM,
        ReportEnum.SLTM,
        ReportEnum.SELM,
        ReportEnum.CIGM,
        ReportEnum.HEYM,
        ReportEnum.YANM,
        ReportEnum.SESM,
        ReportEnum.GUVM,
        ReportEnum.TSUM,
        ReportEnum.IKLM,
        ReportEnum.ALTA,
        ReportEnum.FOYG,
        ReportEnum.FOYM,
        ReportEnum.FOYA,
    };

            using (var merger = new PdfMergerService())
            {
                var mergeOption = new MergeOptions
                {
                    PreserveBookmarks = false,
                    CreateBookmarksForAdditionalPdf = true,
                };

                foreach (var item in collection.Groups)
                {
                    var mainFile = Path.Combine(collection.RootDir, item.Identifier, "main.pdf");

                    List<string> mergeOrderPaths = new();

                    foreach (var type in mergeOrder)
                    {
                        foreach (var value in item.Reports)
                        {
                            if (value.Type == type)
                            {
                                mergeOrderPaths.Add(value.FilePath);
                            }
                        }
                    }

                    var areaId = item.Identifier.Split('-')[0];
                    var centerId = item.Identifier.Split('-')[1];
                    var outputFileDest = Path.Combine(dest, $"TEI-B{areaId}-TM-{centerId}-RED-M00-00.pdf");

                    // Create a temporary file path for the intermediate merge result
                    var tempOutputFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.pdf");

                    MakeEvenPage(mainFile, mergeOrderPaths);

                    var firstOptionMergeSequence = new MergeSequence
                    {
                        MainDocument = mainFile,
                        AdditionalDocuments = mergeOrderPaths,
                        OutputPath = tempOutputFile, // Merge to the temporary file first
                        Options = mergeOption
                    };

                    merger.MergePdf(firstOptionMergeSequence);

                    MakeDividableByFour(tempOutputFile);

                    var coverPagePath = Path.Combine(collection.RootDir, item.Identifier, "ön_kapak.pdf");
                    var blankPagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CoverPages", "BlankPage.pdf");
                    var additionalDocs = new List<string> { blankPagePath, tempOutputFile, @"C:\Users\Mert\Desktop\Risk Raporları ile ilgili her şey\Rapor Kapakları\RED_KAPAK_ARKA.pdf" };

                    //MakeEvenPage(coverPagePath, additionalDocs);

                    var coverPageMergeSequence = new MergeSequence
                    {
                        MainDocument = coverPagePath,
                        AdditionalDocuments = additionalDocs,
                        OutputPath = outputFileDest, // Final output path
                        Options = new MergeOptions
                        {
                            PreserveBookmarks = true,
                            CreateBookmarksForAdditionalPdf = false,
                        }
                    };

                    merger.MergePdf(coverPageMergeSequence);

                    // Clean up the temporary file
                    if (File.Exists(tempOutputFile))
                    {
                        File.Delete(tempOutputFile);
                    }
                    //04-17
                }
            }
        }

        public static void MakeEvenPage(string mainPdf, List<string> additionalPdfs)
        {
            string projectRootPath = AppDomain.CurrentDomain.BaseDirectory;
            string blankPage = Path.Combine(projectRootPath, "CoverPages", "BlankPage.pdf");

            using (var reader = new PdfReaderService())
            using (var merger = new PdfMergerService())
            {
                var mainPageCount = reader.GetPageCount(mainPdf);

                var mergeOption = new MergeOptions
                {
                    PreserveBookmarks = true,
                    CreateBookmarksForAdditionalPdf = false,
                };

                if (mainPageCount % 2 != 0)
                {
                    var MergeSequence = new MergeSequence
                    {
                        MainDocument = mainPdf,
                        AdditionalDocuments = new List<string>() { blankPage },
                        OutputPath = mainPdf,
                        Options = mergeOption
                    };

                    merger.MergePdf(MergeSequence);
                }

                foreach (var item in additionalPdfs)
                {
                    var pageCount = reader.GetPageCount(item);

                    if (pageCount % 2 != 0)
                    {
                        var MergeSequence = new MergeSequence
                        {
                            MainDocument = item,
                            AdditionalDocuments = new List<string>() { blankPage },
                            OutputPath = item,
                            Options = mergeOption
                        };

                        merger.MergePdf(MergeSequence);
                    }
                }

            }
        }

        public static void MakeDividableByFour(string pdfFile)
        {
            string projectRootPath = AppDomain.CurrentDomain.BaseDirectory;
            string blankPage = Path.Combine(projectRootPath, "CoverPages", "BlankPage.pdf");
            using (var reader = new PdfReaderService())
            using (var merger = new PdfMergerService())
            {
                var pageCount = reader.GetPageCount(pdfFile);
                var mergeOption = new MergeOptions
                {
                    PreserveBookmarks = true,
                    CreateBookmarksForAdditionalPdf = false,
                };
                int pagesToAdd = (4 - (pageCount % 4)) % 4;
                List<string> additionalPages = new List<string>();
                for (int i = 0; i < pagesToAdd; i++)
                {
                    additionalPages.Add(blankPage);
                }
                if (additionalPages.Count > 0)
                {
                    var MergeSequence = new MergeSequence
                    {
                        MainDocument = pdfFile,
                        AdditionalDocuments = additionalPages,
                        OutputPath = pdfFile,
                        Options = mergeOption
                    };
                    merger.MergePdf(MergeSequence);
                }
            }
        }

        public static ReportCollection GroupReports(string root)
        {
            var reportFiles = GetReports(root);
            var collection = new ReportCollection(root);

            foreach (var reportFile in reportFiles)
            {
                var baseFileName = Path.GetFileNameWithoutExtension(reportFile);

                foreach (var reportType in Constants.ReportTypes)
                {
                    var (tmNo, buildingCode, buildingTmId) = ExtractParts(baseFileName, reportType.Pattern);
                    if (tmNo == null && buildingCode == null && buildingTmId == null)
                        continue;

                    var report = new Report
                    {
                        FilePath = reportFile,
                        Type = reportType.Type,
                        TmNo = tmNo,
                        BuildingCode = buildingCode,
                        BuildingTmId = buildingTmId
                    };

                    var group = collection.GetOrCreateGroup(tmNo);
                    group.AddReport(report);
                }
            }

            return collection;
        }
    }
}
