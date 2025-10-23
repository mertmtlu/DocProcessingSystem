using DocProcessingSystem.Core;
using DocProcessingSystem.Models;
using DocProcessingSystem.Services;
using iText.Kernel.Pdf.Filters;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using OfficeOpenXml;
using OfficeOpenXml.Style.Dxf;
using Org.BouncyCastle.Asn1.Cmp;
using System;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Net.WebSockets;
using System.Reflection.Metadata;
using System.Security.Cryptography.X509Certificates;
using static iTextSharp.text.pdf.PdfDocument;

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
            Fırat.Run();

            //var inputFolder = @"C:\Users\Mert\Desktop\ekstra 4 tm\raporlar";
            //var outputFolder = @"C:\Users\Mert\Desktop\ekstra 4 tm\raporlar";

            //PdfOperationsHelper.ConvertWordToPdfAsync(inputFolder, outputFolder, false, true).Wait();
            //PdfOperationsHelper.ProcessPdfDocuments();
            //GetTotalPageCount(@"C:\Users\Mert\Desktop\REPORTS");
            //CreateEkA(@"C:\Users\Mert\Desktop\Anıl EK-A\TM Folders");
            //ProcessDocuments();
            //RenameDocumentFiles();
            //CopyUpperDirectory();
            //HandleCrisis();
            //SortPdfFiles();
            //CheckAndFixHakedisFolder();
            //TryChangeText();
            //HakedisHelper.Run();
            //DeleteOnePageEkC();
            //RenameEkAFiles();
            //CheckEkCExists();
            //FixEkBFiles(@"C:\Users\Mert\Desktop\Yukseklik Cap Hatali", @"C:\Users\Mert\Desktop\Yukseklik Cap Hatali new");
            //CheckEkBFiles(@"C:\Users\Mert\Desktop\SZL-3\SZL-3 SAHA", @"C:\Users\Mert\Desktop\SZL-3\SZL-3 SAHA EK-B Corrected");
            //CheckEkBFiles(@"C:\Users\Mert\Desktop\SZL-3\SZL-3 SAHA");
            //TPHelper.Merge(@"Path/to/document/TP");
            //FireReportHelper.Check(@"C:\Users\Mert\Desktop\REPORTS\KK\YAN");
            //RegenMainPdf(@"C:\Users\Mert\Desktop\Yeni klasör", @"C:\Users\Mert\Desktop\dest");
            //Help(@"C:\Users\Mert\Desktop\Yeni klasör");
            //CheckEkFiles();
            //CheckEkCMissingFiles();

            //GenerateDonatiReports();
            //DeliveryHelper.GetEkParts(@"C:\users\Mert\Desktop\REPORTS", @"C:\Users\Mert\Desktop\REPORTSTM");
            //DeliveryHelper.CreateRedReport(@"C:\Users\Mert\Desktop\ekstra 4 tm\raporlar", @"C:\Users\Mert\Desktop\ekstra 4 tm\RED raporları");
            //CountFiles(@"C:\Users\Mert\Desktop\Yeni klasör (2) - Kopya");
            //RenameDocumentFiles();
            //CopyFoyReports();

            //ReArrangeCoverPages();
            //RegenMainPdf(@"C:\Users\Mert\Desktop\Yeni Klasör (2)", @"C:\Users\Mert\Desktop\Kapak");
            //MergeRedReports();
            //PlaceSignaturesOnFoyDocument();
            //PlaceSignaturesOnEachDocument();

            //FindEkCFiles();
            //GetTeiDocuments();
            //MergePdfs();
            //CopyPdfs();
            //SaveMeFromRontgenWork(@"C:\Users\Mert\Desktop\SZL-3\SZL-3 SAHA");
        }

        #endregion

        #region Document Processing Functions

        static string FindKeyValue(string key, Dictionary<string, string> dict)
        {
            var searchTerm = key
                    .Replace(" BİNASI", "")
                    .Replace(" EK BİNA", "")
                    .Replace(" GÜÇLENDİRME", "")
                    .Replace("ESKİ ", "")
                    .Replace(" ZEMİN KAT", "")
                    .Replace(" 1.KAT", "")
                    .Replace(" 1. KAT", "")
                    .Replace(" 2.KAT", "")
                    .Replace(" 2. KAT", "")
                    .Replace(" 3.KAT", "")
                    .Replace(" 3. KAT", "")
                    .Replace(" BODRUM KAT", "")
                    .Replace("380 KV ", "")
                    .Replace(" A BLOK", "")
                    .Replace("+A BLOK", "")
                    .Replace(" B BLOK", "")
                    .Replace("+B BLOK", "")
                    .Replace(" C BLOK", "")
                    .Replace("+C BLOK", "");

            foreach (var kvp in dict)
            {
                if (kvp.Key.Contains(searchTerm)) return kvp.Value;
            }
            try
            {
                throw new Exception();
            }
            catch (Exception ex)
            {
                return "";
            }
        }
        static bool Check(string folderName, string[] ekFiles)
        {
            foreach (var ek in ekFiles)
            {
                if (ek.Contains(folderName)) return true;
            }
            return false;
        }
        static void SaveMeFromRontgenWork(string root)
        {
            var subDirectories = Directory.GetDirectories(root);

            var ekFiles = Directory.GetFiles(root, "EK-C.pdf", SearchOption.AllDirectories);

            Dictionary<string, string> keyValuePairs = new()
            {
                { "KUMANDA", "M01_01" },
                { "KUMANDA A BLOK 1. KAT", "M01_01" },
                { "KUMANDA A BLOK 2. KAT", "M01_01" },
                { "KUMANDA A BLOK BODRUM KAT", "M01_01" },
                { "KUMANDA A BLOK ZEMİN KAT", "M01_01" },
                { "KUMANDA B BLOK 1. KAT", "M0_011" },
                { "KUMANDA B BLOK 2. KAT", "M01_01" },
                { "KUMANDA B BLOK BODRUM KAT", "M01_01" },
                { "KUMANDA B BLOK ZEMİN KAT", "M01_01" },
                { "KUMANDA BİNASI", "M01_01" },
                { "KUMANDA İLAVE", "M01_01" },
                { "380 KUMANDA 1. KAT", "M01_01" },
                { "380 KUMANDA ZEMİN KAT", "M01_01" },
                { "380 KV KUMANDA BİNASI 1.KAT", "M01_01" },
                { "380 KV KUMANDA BİNASI ZEMİN KAT", "M01_01" },
                { "KAPALI ŞALT-1", "M02_01" },
                { "KAPALI ŞALT-2", "M02_02" },
                { "KAPALI SALT", "M02_01" },
                { "KAPALI ŞALT", "M02_01" },
                { "KAPALI ŞALT A BLOK", "M02_01" },
                { "KAPALI ŞALT B BLOK", "M02_01" },
                { "KAPALI ŞALT C BLOK", "M02_01" },
                { "KAPALI ŞALT D BLOK", "M02_01" },
                { "GÜVENLIK", "M10_01" },
                { "GÜVENLİK", "M10_01" },
                { "GÜVENLİK BODRUM", "M10_01" },
                { "GÜVENLİK ZEMİN", "M10_01" },
                { "GÜVENLİK 1. KAT", "M10_01" },
                { "GÜVENLIK 2. KAT", "M10_01" },
                { "GÜVENLİK BODRUM KAT", "M10_01" },
                { "RÖLE", "M05_01" },
                { "GIS", "M07_01" },
                { "TRAFO", "M11_01" },
                { "MC VE KUMANDA", "M04_01" },
                { "MC KUMANDA 1. KAT", "M04_01" },
                { "MC KUMANDA BODRUM KAT", "M04_01" },
                { "MC KUMANDA ZEMİN KAT", "M04_01" },
                { "MC VE KUMANDA 1. KAT", "M04_01" },
                { "MC VE KUMANDA BODRUM KAT", "M04_01" },
                { "MC VE KUMANDA ZEMİN KAT", "M04_01" },
                { "MC BODRUM KAT", "M03_01" },
                { "MC ZEMİN KAT", "M03_01" },
                { "TELEKOM", "M06_01" },
                { "TELEKOM-", "M06_01" },
                { "TELEKOM RÖLE", "TELEKOM RÖLE" },
                { "YARDIMCI SERVİS TR BİNASI", "YARDIMCI SERVİS TR BİNASI" },
                { "ELEKTRONİK", "ELEKTRONİK" },
                { "PANO", "PANO" },

            };

            foreach (var sub in subDirectories)
            {
                var ltFolder = Directory.GetDirectories(sub).FirstOrDefault(f => f.Contains("LT"));
                var ltSubfolders = ltFolder is not null ? Directory.GetDirectories(ltFolder) : null;

                if (ltSubfolders != null)
                {
                    foreach (var ltSubfolder in ltSubfolders)
                    {
                        var rontgenFolder = Directory.GetDirectories(ltSubfolder).FirstOrDefault(f => f.Contains("RÖNTGEN"));

                        if (rontgenFolder != null)
                        {
                            var buildings = Directory.GetDirectories(rontgenFolder);

                            foreach (var building in buildings)
                            {
                                var checker = Directory.GetFiles(building,"*", SearchOption.AllDirectories).Length > 0;

                                var folderName = $"{Path.GetFileName(ltSubfolder).Substring(0, 5)}: {Path.GetFileName(building)}";

                                //var folderName = $"{Path.GetFileName(ltSubfolder).Substring(0, 5)}_{FindKeyValue(Path.GetFileName(building), keyValuePairs)}";

                                //if (Check(folderName, ekFiles))
                                //{
                                //    folderName += ": true";
                                //}

                                if (!checker)
                                {
                                    folderName += " sıkıntı";
                                }

                                Console.WriteLine(folderName);
                            }
                        }
                        else
                        {
                            Console.WriteLine($"{Path.GetFileName(ltSubfolder).Substring(0, 5)}: None");
                        }

                    }
                }
                else
                {
                    Console.WriteLine($"Subfolder not found in: {ltFolder}");
                }
            }
        }
        static void CopyPdfs()
        {
            string root = @"C:\Users\Mert\Downloads\SZL-3_Hakedis-20_Hazirlik (31.07.2025)";
            string dest = @"C:\Users\Mert\Desktop\hk20";

            var pdfs = Directory.GetFiles(root, "*.pdf", SearchOption.AllDirectories);

            foreach (var pdf in pdfs)
            {
                var pdfName = Path.GetFileName(pdf);

                var destination = Path.Combine(dest, pdfName);

                if (File.Exists(destination)) Console.WriteLine($"{pdfName} exists!!");

                File.Copy(pdf, destination, true);
            }
        }
        static void MergePdfs()
        {
            string root = @"C:\Users\Mert\Desktop\HK20_1";
            string dest = @"C:\Users\Mert\Desktop\merged_1";


            //PdfOperationsHelper.ConvertWordToPdfAsync(root, root, false, true, 8).Wait();

            var pdfs = Directory.GetFiles(root, "TEI*.pdf", SearchOption.AllDirectories);
            using (var merger = new PdfMergerService())
            {
                var mergeOption = new MergeOptions
                {
                    PreserveBookmarks = true
                };

                foreach (var pdf in pdfs)
                {
                    string folderPath = Path.GetDirectoryName(pdf);

                    var ekFiles = Directory.GetFiles(folderPath, "EK-*.pdf", SearchOption.TopDirectoryOnly)
                        .Order()
                        .ToList();

                    var MergeSequence = new MergeSequence
                    {
                        MainDocument = pdf,
                        AdditionalDocuments = ekFiles,
                        OutputPath = Path.Combine(dest, Path.GetFileName(pdf)),
                        Options = mergeOption
                    };

                    merger.MergePdf(MergeSequence);
                }
            }
        }
        static void GetTeiDocuments()
        {
            string root = @"C:\Users\Mert\Downloads\SZL-3\SZL-3";
            string dest = @"C:\Users\Mert\Downloads\SZL-3\out";

            var pdfs = Directory.GetFiles(root, "TEI*.pdf", SearchOption.AllDirectories);

            foreach (var pdf in pdfs)
            {
                var type = Constants.ReportTypes.Find(item => pdf.Contains(item.Pattern));

                if (type == null)
                {
                    Console.WriteLine($"{Path.GetFileNameWithoutExtension(pdf)} does not contain any of patterns: {pdf}");
                    continue;
                }

                var (tmNo, buildingCode, buildingTmID) = FolderHelper.ExtractParts(pdf, type.Pattern);
                var areaId = tmNo.Split("-")[0];
                var destination = Path.Combine(dest, $"BOLGE-{areaId}");

                if (!Directory.Exists(destination))
                {
                    Directory.CreateDirectory(destination);
                }

                File.Copy(pdf, Path.Combine(destination, Path.GetFileName(pdf)), true);
            }


        }
        static void FindEkCFiles()
        {
            string root = @"D:\SZL-3 2008 BA (AI, MM, SO)\SZL-3 SAHA";

            var pdfs = Directory.GetFiles(root, "EK-C.pdf", SearchOption.AllDirectories);

            foreach (var pdf in pdfs)
            {
                var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(pdf);

                Console.WriteLine($"{Path.GetFileName(Path.GetDirectoryName(pdf))}");
            }
        }
        static void PlaceSignaturesOnEachDocument()
        {
            var root = @"C:\Users\Mert\Desktop\ekstra 4 tm\raporlar";
            var dest = @"C:\Users\Mert\Desktop\ekstra 4 tm\raporlar2";
            var pdfs = Directory.GetFiles(root, "*.pdf", SearchOption.AllDirectories);

            foreach (var pdf in pdfs)
            {
                var destination = pdf.Replace(root, dest);

                var directoryPath = Path.GetDirectoryName(destination);

                if (!string.IsNullOrEmpty(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                if (Path.GetFileNameWithoutExtension(pdf).Contains("-FOY-M") || Path.GetFileNameWithoutExtension(pdf).Contains("-FOY-A"))
                {
                    File.Copy(pdf, destination, true);
                }
                else
                {
                    PlaceSignaturesOnTeiDocument(pdf, destination);
                }
            }
        }
        /// <summary>
        /// Places signature images on the TEI PDF document
        /// </summary>
        /// <param name="inputPdfPath">Path to the input PDF file</param>
        /// <param name="outputPdfPath">Path for the output PDF file with signatures</param>
        static void PlaceSignaturesOnFoyDocument(string inputPdfPath = null, string outputPdfPath = null)
        {
            // Default paths if not provided
            string defaultInputPath = @"C:\Users\Mert\Desktop\TEI-B01-TM-10-FOY-M00-00.pdf";
            string defaultOutputPath = @"C:\Users\Mert\Desktop\TEI-B01-TM-10-FOY-M00-00-SINGED.pdf";

            string finalInputPath = inputPdfPath ?? defaultInputPath;
            string finalOutputPath = outputPdfPath ?? defaultOutputPath;

            // Signature image paths
            string ahmetYakutSignature = @"C:\Users\Mert\Desktop\Ahmet Yakut.png";
            string barisBiniciSignature = @"C:\Users\Mert\Desktop\Barış Binici.png";

            try
            {
                Console.WriteLine("Starting signature placement process...");
                Console.WriteLine($"Input PDF: {finalInputPath}");
                Console.WriteLine($"Output PDF: {finalOutputPath}");

                // Get page dimensions to calculate positions
                var (pageWidth, pageHeight) = PdfImagePlacementService.GetPageDimensions(finalInputPath, 1);
                Console.WriteLine($"Page dimensions: {pageWidth:F1} x {pageHeight:F1} points");

                // Configuration for signature images
                var imageConfigs = new[]
                {
                    // Ahmet Yakut signature (upper right signature area)
                    (ahmetYakutSignature, new ImagePlacementOptions
                    {
                        PlaceOnAllPages = true,
                        X = pageWidth - PdfImagePlacementService.InchesToPoints(10f),
                        Y = pageHeight - PdfImagePlacementService.InchesToPoints(16f),
                        Height = PdfImagePlacementService.MillimetersToPoints(20f * 1.41428571429f),
                        MaintainAspectRatio = true,
                        PlaceInBackground = false, // Place over text
                        Opacity = 1.0f
                    }),

                    // Barış Binici signature (lower right signature area)
                    (barisBiniciSignature, new ImagePlacementOptions
                    {
                        PlaceOnAllPages = true,
                        X = pageWidth - PdfImagePlacementService.InchesToPoints(5.4f),
                        Y = pageHeight - PdfImagePlacementService.InchesToPoints(16f),
                        Height = PdfImagePlacementService.MillimetersToPoints(23.3f * 1.41428571429f),
                        MaintainAspectRatio = true,
                        PlaceInBackground = false, // Place over text
                        Opacity = 1.0f
                    })
                };

                // Place both signatures on the PDF
                PdfImagePlacementService.PlaceMultipleImagesOnPdf(
                    finalInputPath,
                    finalOutputPath,
                    imageConfigs
                );

                Console.WriteLine("✅ Signature placement completed successfully!");
                Console.WriteLine($"Signed document saved to: {finalOutputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error during signature placement: {ex.Message}");
                Console.WriteLine($"Full error details: {ex}");
                throw;
            }
        }
        /// <summary>
        /// Places signature images on the TEI PDF document
        /// </summary>
        /// <param name="inputPdfPath">Path to the input PDF file</param>
        /// <param name="outputPdfPath">Path for the output PDF file with signatures</param>
        static void PlaceSignaturesOnTeiDocument(string inputPdfPath = null, string outputPdfPath = null)
        {
            // Default paths if not provided
            string defaultInputPath = @"C:\Users\Mert\Desktop\TEI-B01-TM-10-SLT-M00-00_NT (AMBARLI-SALT INCELEME).pdf";
            string defaultOutputPath = @"C:\Users\Mert\Desktop\TEI-B01-TM-10-SLT-M00-00_NT (AMBARLI-SALT INCELEME)_SIGNED.pdf";

            string finalInputPath = inputPdfPath ?? defaultInputPath;
            string finalOutputPath = outputPdfPath ?? defaultOutputPath;

            // Signature image paths
            string ahmetYakutSignature = @"C:\Users\Mert\Desktop\Risk Raporları ile ilgili her şey\Ahmet Yakut.png";
            string barisBiniciSignature = @"C:\Users\Mert\Desktop\Risk Raporları ile ilgili her şey\Barış Binici.png";

            try
            {
                Console.WriteLine("Starting signature placement process...");
                Console.WriteLine($"Input PDF: {finalInputPath}");
                Console.WriteLine($"Output PDF: {finalOutputPath}");

                // Get page dimensions to calculate positions
                var (pageWidth, pageHeight) = PdfImagePlacementService.GetPageDimensions(finalInputPath, 1);
                Console.WriteLine($"Page dimensions: {pageWidth:F1} x {pageHeight:F1} points");

                // Configuration for signature images
                var imageConfigs = new[]
                {
                    // Ahmet Yakut signature (upper right signature area)
                    (ahmetYakutSignature, new ImagePlacementOptions
                    {
                        PageNumber = 1,
                        X = pageWidth - PdfImagePlacementService.InchesToPoints(7f), // ~2.8 inches from right edge
                        Y = pageHeight - PdfImagePlacementService.InchesToPoints(11f), // ~2.2 inches from top
                        Height = PdfImagePlacementService.MillimetersToPoints(20f), // 0.8 inch height
                        MaintainAspectRatio = true,
                        PlaceInBackground = false, // Place over text
                        Opacity = 1.0f
                    }),

                    // Barış Binici signature (lower right signature area)
                    (barisBiniciSignature, new ImagePlacementOptions
                    {
                        PageNumber = 1,
                        X = pageWidth - PdfImagePlacementService.InchesToPoints(3.2f), // ~2.8 inches from right edge
                        Y = pageHeight - PdfImagePlacementService.InchesToPoints(11f), // ~3.5 inches from top
                        Height = PdfImagePlacementService.MillimetersToPoints(23.3f), // 0.8 inch height
                        MaintainAspectRatio = true,
                        PlaceInBackground = false, // Place over text
                        Opacity = 1.0f
                    })
                };

                // Place both signatures on the PDF
                PdfImagePlacementService.PlaceMultipleImagesOnPdf(
                    finalInputPath,
                    finalOutputPath,
                    imageConfigs
                );

                Console.WriteLine("✅ Signature placement completed successfully!");
                Console.WriteLine($"Signed document saved to: {finalOutputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error during signature placement: {ex.Message}");
                Console.WriteLine($"Full error details: {ex}");
                throw;
            }
        }
        static void ReArrangeCoverPages()
        {
            string root = @"C:\Users\Mert\Desktop\Kapak";
            var pdfs = Directory.GetFiles(root, "kapak.pdf", SearchOption.AllDirectories);

            var pdfExtractor = new PdfRangeExtractorService();

            var option = new PdfExtractionOptions
            {
                StartPageSelectionType = PageSelectionType.FirstPage,
                EndPageSelectionType = PageSelectionType.SpecificPage,
                EndPageNumber = 6
            };

            foreach (var pdf in pdfs)
            {
                pdfExtractor.ExtractRange(pdf, pdf.Replace("kapak", "rearrenged"), option);
            }
        }
        static void MergeRedReports()
        {
            string root = @"C:\Users\Mert\Desktop\Kapak";
            string dest = @"C:\Users\Mert\Desktop\Yeni Klasör (2)";

            var subfolders = Directory.GetDirectories(root);

            using (var merger = new PdfMergerService())
            {
                foreach (var subfolder in subfolders)
                {
                    var coverPage = Path.Combine(subfolder, "rearrenged.pdf");
                    var mainPdf = Path.Combine(subfolder, "main.pdf");

                    var folderName = Path.GetFileName(subfolder);

                    var outputPath = Path.Combine(dest, folderName, "main.pdf");

                    MergeOptions option = new()
                    {
                        PreserveBookmarks = true,
                        CreateBookmarksForAdditionalPdf = false
                    };

                    merger.MergePdf(
                        coverPage,
                        new() { mainPdf },
                        outputPath,
                        option
                    );

                }
            }
        }
        static void CopyFoyReports()
        {
            string root = @"C:\Users\Mert\Desktop\Yeni Klasör (2)";
            string dest = @"C:\Users\Mert\Desktop\FOY Reports";

            var pdfs = Directory.GetFiles(root, "*.pdf", SearchOption.AllDirectories)
                .Where(f => Path.GetFileName(f).Contains("FOY"))
                .ToList();

            foreach (var pdf in pdfs)
            {
                var destinationFile = Path.Combine(dest, Path.GetFileName(pdf));

                File.Copy(pdf, destinationFile);
            }
        }
        static void CountFiles(string rootPath)
        {
            try
            {
                // Method 1: Get results as dictionary
                var results = FileCounter.CountFilesBySubfolder(rootPath);

                // Method 2: Print results directly
                FileCounter.PrintFileCountsBySubfolder(rootPath);

                // Example of using the dictionary results
                foreach (var folder in results)
                {
                    Console.WriteLine($"Folder: {folder.Key}");
                    var totalCount = folder.Value.Values.Sum();
                    Console.WriteLine($"Total image/PDF files: {totalCount}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        static void GenerateDonatiReports()
        {
            Dictionary<int, string> d = new()
            {
                { 2, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR2.xlsm" },
                { 3, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR3.xlsm" },
                { 4, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR4.xlsm" },
                { 5, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR5.xlsm" },
                { 6, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR6.xlsm" },
                { 7, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR7.xlsm" },
                { 8, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR8.xlsm" },
                { 9, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR9.xlsm" },
                { 10, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR10.xlsm" },
                { 11, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR11.xlsm" },
                { 12, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR12.xlsm" },
                { 13, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR13.xlsm" },
                { 14, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR14.xlsm" },
                { 15, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR15.xlsm" },
                { 16, @"C:\Users\Mert\Desktop\input donati rapor\_DONATI RAPOR16e.xlsm" },
            };

            EkCHelper.GenerateExcelReport(@"C:\Users\Mert\Desktop\input donati rapor\gen.xlsm", d, @"C:\Users\Mert\Desktop\outfiles");
        }
        static void CheckEkCMissingFiles()
        {
            var root = @"C:\Users\Mert\Desktop\SZL-3\SZL-3 SAHA\16.BÖLGE\RÖNTGEN";
            var pdfs = Directory.GetFiles(root, "*.pdf", SearchOption.AllDirectories);
            Dictionary<string, string> seen = new();
            List<string> pdfsToBeRemoved = new();
            foreach (var item in Constants.EkCPdfCheck)
            {
                bool found = false;
                foreach (var pdf in pdfs)
                {
                    if (pdf.Contains(item))
                    {
                        if (seen.Keys.Contains(item))
                        {
                            Console.WriteLine($"{item} allready exists");
                        }
                        seen.Add(item, pdf);
                        found = true;
                        pdfsToBeRemoved.Add(pdf);
                        break;
                    }
                }

                if (!found)
                {
                    Console.WriteLine($"Rontgen does not exists: {item}");
                }
            }

            foreach (var pdf in pdfs)
            {
                if (!pdfsToBeRemoved.Contains(pdf))
                {
                    Console.WriteLine($"Extra pdf detected: {pdf}");
                }
            }
        }
        static void Help(string root)
        {
            var pdfs = Directory.GetDirectories(root);

            foreach (var dict in pdfs)
            {
                var file = Path.Combine(dict, "main.pdf");

                if (!File.Exists(file))
                {
                    Console.WriteLine($"{file} does not exists.");
                }
            }
        }
        static void RegenMainPdf(string root, string dest)
        {
            var pdfs = Directory.GetFiles(root, "main.pdf", SearchOption.AllDirectories);

            var pdfExtractor = new PdfRangeExtractorService();

            // Options for extracting the main document (before EK-D)
            //var mainDocumentOptions = new PdfExtractionOptions
            //{
            //    StartPageSelectionType = PageSelectionType.FirstPage,
            //    EndPageSelectionType = PageSelectionType.Keyword,
            //    EndKeyword = new KeywordOptions
            //    {
            //        Keyword = "SONUÇ VE ÖNERİLER",
            //        Occurrence = KeywordOccurrence.Last,
            //        IncludeMatchingPage = true,
            //    }
            //};

            var mainDocumentOptions = new PdfExtractionOptions
            {
                StartPageSelectionType = PageSelectionType.Keyword,
                StartKeyword = new KeywordOptions
                {
                    Keyword = "Fizibilite Çalışmalarıyla",
                    Occurrence = KeywordOccurrence.First,
                    IncludeMatchingPage = true,
                },
                EndPageSelectionType = PageSelectionType.LastPage
            };

            foreach (var pdf in pdfs)
            {
                pdfExtractor.ExtractRange(pdf, pdf.Replace(root, dest), mainDocumentOptions);
            }

        }
        static void CheckEkCExists()
        {
            string root = @"C:\Users\Mert\Desktop\SZL-3\SZL-3 SAHA\9.BÖLGE\TM_Folders";

            var folders = Directory.GetDirectories(root);

            foreach (string folder in folders)
            {
                var ekCPath = Path.Combine(folder, "EK-C.pdf");

                if (!File.Exists(ekCPath))
                {
                    Console.WriteLine($"EK-C missing for path: {ekCPath}");
                }
            }
        }
        static void CheckEkBFiles(string root)
        {
            var ekB = Directory.GetFiles(root, "EK-B.pdf", SearchOption.AllDirectories);
            var textReplacer = new PdfTextReplacerService();

            List<(string, string, string)> checker = new();

            foreach (var building in Constants.HK19)
            {
                (string tmNo, string buildingCode, string buildingTmId) = FolderHelper.ExtractParts(building);
                if (tmNo == null || buildingCode == null || buildingTmId == null) throw new ArgumentNullException();

                if (!checker.Contains((tmNo, buildingCode, buildingTmId))) checker.Add((tmNo, buildingCode, buildingTmId));
            }


            foreach ((string tmNo, string buildingCode, string buildingTmId) building in checker)
            {
                var areaID = Convert.ToInt32(building.tmNo.Split('-')[0]);
                var areaPath = Path.Combine(root, $"{areaID}.BÖLGE", "TM_Folders");
                var folders = Directory.GetDirectories(areaPath);

                foreach (var folder in folders)
                {
                    (string tmNo, string buildingCode, string buildingTmId) = FolderHelper.ExtractParts(folder);
                    if (tmNo == null || buildingCode == null || buildingTmId == null) throw new ArgumentNullException();

                    if (tmNo == building.tmNo && buildingCode == building.buildingCode && buildingTmId == building.buildingTmId)
                    {
                        var ekBPath = Path.Combine(folder, "EK-B.pdf");

                        if (!File.Exists(ekBPath))
                        {
                            Console.WriteLine($"Check EK-B for {building.tmNo}-M{building.buildingCode}-{building.buildingTmId} for path: {ekBPath}");
                        }

                        break;
                    }
                }
            }
        }
        static void CheckEkBFiles(string root, string dest)
        {
            var ekB = Directory.GetFiles(root, "EK-B.pdf", SearchOption.AllDirectories);

            foreach (var ekBPath in ekB)
            {
                var destinationPath = ekBPath.Replace(root, dest);

                string destinationDirectory = Path.GetDirectoryName(destinationPath);
                if (!Directory.Exists(destinationDirectory))
                {
                    Directory.CreateDirectory(destinationDirectory);
                }

                File.Copy(ekBPath, destinationPath, true);
            }
        }
        static void FixEkBFiles(string root, string dest)
        {
            var ekB = Directory.GetFiles(root, "*.pdf", SearchOption.AllDirectories);
            var textReplacer = new PdfTextReplacerService();

            foreach (var ekBPath in ekB)
            {
                textReplacer.ReplaceCapYukseklik(ekBPath, ekBPath.Replace(root, dest));
            }
        }
        static void TryChangeText()
        {
            string inputFolder = @"C:\Users\Mert\Desktop\Yeni klasör (3)";
            string outputFolder = @"C:\Users\Mert\Desktop\Yeni klasör (4)";

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
            string parametricsFolder = @"C:\Users\Mert\Desktop\ANIL_REVISE 27.05.2025\words\Parametricasdasd";
            string deterministicsFolder = @"C:\Users\Mert\Desktop\ANIL_REVISE 27.05.2025\words\Deterministic";
            string post2008 = @"C:\Users\Mert\Desktop\Fırat Report Revision\MM_RAPOR\WORDasdasdasd"; // TODO: FIRAT
            string analysisFolder = @"C:\Users\Mert\Desktop\ANIL_REVISE 27.05.2025\Analysis"; // TODO: FIRAT

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
        static void CheckEkFiles()
        {
            var root = @"C:\Users\Mert\Desktop\SZL2\Anıl Report Revision\Analysis";

            var groups = FolderHelper.GroupFolders(root);

            using (var reader = new PdfReaderService())
            {

                foreach (var group in groups)
                {
                    if (File.Exists(Path.Combine(group.MainFolder, "EK_KAROT.pdf")) && File.Exists(Path.Combine(group.MainFolder, "EK_DONATI.pdf")))
                    {
                        var readRontgen = Path.Combine(group.MainFolder, "EK_DONATI.pdf");
                        var readBasınc = Path.Combine(group.MainFolder, "EK_KAROT.pdf");

                        if (!reader.ContainsText(readRontgen, "Röntgen"))
                        {
                            Console.WriteLine($"{readRontgen} does not contains Röntgen keyword!");
                        }

                        if (!reader.ContainsText(readBasınc, "Basınç"))
                        {
                            Console.WriteLine($"{readBasınc} does not contains Basınç keyword!");
                        }
                    }
                    else if (File.Exists(Path.Combine(group.MainFolder, "EK-B.pdf")) && File.Exists(Path.Combine(group.MainFolder, "EK-C.pdf")))
                    {

                        var readRontgen = Path.Combine(group.MainFolder, "EK-C.pdf");
                        var readBasınc = Path.Combine(group.MainFolder, "EK-B.pdf");

                        if (!reader.ContainsText(readRontgen, "Röntgen"))
                        {
                            Console.WriteLine($"{readRontgen} does not contains Röntgen keyword!");
                        }

                        if (!reader.ContainsText(readBasınc, "Basınç"))
                        {
                            Console.WriteLine($"{readBasınc} does not contains Basınç keyword!");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"{group.MainFolder} does not contain necessary files");
                    }
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
        static void DeleteOnePageEkC()
        {
            var root = @"C:\Users\Mert\Desktop\SZL-3\SZL-3 SAHA";

            var ekFiles = Directory.GetFiles(root, "EK-C.pdf", SearchOption.AllDirectories);

            PdfReaderService pdfReaderService = new();

            foreach (var ekFile in ekFiles)
            {
                if (!pdfReaderService.PageExists(ekFile, 2))
                {
                    Console.WriteLine($"Deleting file: {ekFile}");

                    File.Delete(ekFile);
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
        static void RenameEkAFiles()
        {
            var root = @"C:\Users\Mert\Desktop\fırat ek a";

            var files = Directory.GetFiles(root, "EK-A.pdf", SearchOption.AllDirectories);

            foreach (var file in files)
            {
                try
                {
                    var (tmNo, buildingCode, buildingTmID) = FolderHelper.ExtractParts(file);

                    if (tmNo == null || buildingCode == null || buildingTmID == null) { Console.WriteLine($"Cannot process: {file}"); ; continue; }

                    var newName = $"{tmNo}-{buildingCode}-{buildingTmID}.pdf";

                    var destination = Path.Combine(root, newName);

                    File.Copy(file, destination);

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {file}, Exception: {ex.Message}");
                }

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
            var inputFolder = @"C:\Users\Mert\Desktop\hk20";
            var excelFile = @"C:\Users\Mert\Desktop\SZL-2_TM_KISA_TR_ISIM_LISTE_20250319.xlsx";

            var tmNameJson = ConvertExcelToDictionary(excelFile);

            // Get both Word and PDF documents
            //var wordDocuments = Directory.GetFiles(inputFolder, "*.docx", SearchOption.AllDirectories);
            var allDocuments = Directory.GetFiles(inputFolder, "TEI*.pdf", SearchOption.AllDirectories);
            //var allDocuments = wordDocuments.Concat(pdfDocuments).ToArray();

            foreach (var document in allDocuments)
            {
                // Get the file extension to preserve it in the renamed file
                string fileExtension = Path.GetExtension(document);

                if (document.Contains("FOY-A0"))
                {
                    List<int> ints = new() { 1, 2, 3 };

                    foreach (var i in ints)
                    {
                        string preference = $"FOY-A0{i}";

                        try
                        {
                            var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(document, "FOY");

                            // Get the shortened name for this TM number
                            var shortenedName = FindShortenedName(tmNo, tmNameJson)?.ToString();

                            if (shortenedName == null) throw new ArgumentNullException("Shortened Name Not Found.");

                            // Split the TM number to get area ID and TM ID
                            var areaId = tmNo.Split("-")[0];
                            var tmId = tmNo.Split("-")[1];

                            var newName = $"TEI-B{areaId}-TM-{tmId}-{preference}-00_NT ({shortenedName}-ALTERNATIF {i} {Constants.ReportType["FOY"]}){fileExtension}";

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
                else if (document.Contains("M00"))
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
                else if (document.Contains("DGR"))
                {
                    try
                    {
                        // Extract information from the filename
                        var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(document, "DGR");

                        if (buildingCode == "19") buildingCode = "11";

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
                        var newName = $"TEI-B{areaId}-TM-{tmId}-DGR-M{buildingCode}-{buildingTmId}_NT ({shortenedName}-{buildingName} ACILGUCPAK){fileExtension}";

                        // Get the directory path from the original document
                        string directoryPath = Path.GetDirectoryName(document);
                        //string directoryPath = @"C:\Users\Mert\Desktop\SZL2\KK_INCELEDI_DEGISTIRDI\output";

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
                else if (document.Contains("DIR"))
                {
                    try
                    {
                        // Extract information from the filename
                        var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(document, "DIR");

                        if (buildingCode == "19") buildingCode = "11";

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
                        var newName = $"TEI-B{areaId}-TM-{tmId}-DIR-M{buildingCode}-{buildingTmId}_NT ({shortenedName}-{buildingName} ACILGUCPAK){fileExtension}";

                        // Get the directory path from the original document
                        string directoryPath = Path.GetDirectoryName(document);
                        //string directoryPath = @"C:\Users\Mert\Desktop\SZL2\KK_INCELEDI_DEGISTIRDI\output";

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
        static void CopyUpperDirectory()
        {
            var root = @"C:\Users\Mert\Desktop\KK\PUSHOVER";
            var allDocuments = Directory.GetFiles(root, "TEI*.pdf", SearchOption.AllDirectories);

            foreach (var file in allDocuments)
            {
                // Skip files that are already in the root directory
                if (Path.GetDirectoryName(file) == root)
                    continue;

                // Get the parent directory
                var parentDir = Directory.GetParent(Path.GetDirectoryName(file)).FullName;

                // Create the destination path by combining parent directory and filename
                var destPath = Path.Combine(parentDir, Path.GetFileName(file));

                // Only copy if the file doesn't already exist at the destination
                if (!File.Exists(destPath))
                {
                    File.Copy(file, destPath, false);
                    Console.WriteLine($"Copied {file} to {destPath}");
                }
                else
                {
                    Console.WriteLine($"Skipped {file} - file already exists at {destPath}");
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