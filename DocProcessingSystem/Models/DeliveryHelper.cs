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

namespace DocProcessingSystem.Models
{
    public static class DeliveryHelper
    {
        public static List<string> GetReports(string root)
        {
            return Directory.GetFiles(root, "TEI*.pdf", SearchOption.AllDirectories).ToList();
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
                //{ReportEnum.GUVM, "EK-A TESİS"},
                //{ReportEnum.HEYM, "EK-A TESİS"},
                //{ReportEnum.SELM, "EK-A TESİS"},
                //{ReportEnum.SESM, "EK-A TESİS"},
                //{ReportEnum.YANM, "EK-A TESİS"},
                //{ReportEnum.TSUM, "EK-A TESİS"},

                {ReportEnum.SLTM, "EK-A" }
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
                ReportEnum.FOYG,
                ReportEnum.FOYM,
                ReportEnum.FOYA,
                ReportEnum.ALTA,
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

                    MakeEvenPage(mainFile, mergeOrderPaths);

                    var firstOptionMergeSequence = new MergeSequence
                    {
                        MainDocument = mainFile,
                        AdditionalDocuments = mergeOrderPaths,
                        OutputPath = outputFileDest, 
                        Options = mergeOption
                    };

                    merger.MergePdf(firstOptionMergeSequence);


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
                    PreserveBookmarks = true
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
