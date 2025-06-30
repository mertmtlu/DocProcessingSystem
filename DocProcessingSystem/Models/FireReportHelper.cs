using DocProcessingSystem.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem.Models
{
    public static class FireReportHelper
    {
        public static void Check(string root)
        {
            using (var pdfReader = new PdfReaderService())
            {
                var pdfFiles = Directory.GetFiles(root, "*.pdf", SearchOption.AllDirectories);

                // Define checks with page specifications: <description, (text to find, list of page numbers)>
                // If page list is empty, check all pages
                var checks = new Dictionary<string, (string text, List<int> pages)>
                {
                    {"date", ("27.08.2024", new List<int> { 1, 3 }) },
                    {"deliverance type", ("NİHAİ TESLİM", new List<int> { 1 }) },
                    {"report name", ("YANGIN RİSKİ DEĞERLENDİRME RAPORU", new List<int> { 1 }) },
                    {"sonuç ve öneri part", ("SONUÇ VE ÖNERİLER", new List<int> { 12 }) },
                };

                foreach (var pdfFile in pdfFiles)
                {
                    foreach (var check in checks)
                    {
                        var description = check.Key;
                        var searchText = check.Value.text;
                        var pages = check.Value.pages;

                        // If page list is empty, check entire document
                        if (pages.Count == 0)
                        {
                            if (!pdfReader.ContainsText(pdfFile, searchText, true))
                            {
                                Console.WriteLine($"Incorrect {description}: {pdfFile}");
                            }
                        }
                        // Check specific pages
                        else
                        {
                            bool foundOnAnyPage = false;
                            foreach (var pageNum in pages)
                            {
                                try
                                {
                                    if (pdfReader.ContainsTextOnPage(pdfFile, searchText, pageNum, true))
                                    {
                                        foundOnAnyPage = true;
                                        break;
                                    }
                                }
                                catch (ArgumentOutOfRangeException)
                                {
                                    // Handle case where PDF doesn't have the specified page
                                    Console.WriteLine($"PDF doesn't have page {pageNum}: {pdfFile}");
                                }
                            }

                            if (!foundOnAnyPage)
                            {
                                if (pages.Count == 1)
                                    Console.WriteLine($"Incorrect {description} on page {pages[0]}: {pdfFile}");
                                else
                                    Console.WriteLine($"Incorrect {description} on pages {string.Join(", ", pages)}: {pdfFile}");
                            }
                        }
                    }

                    // Check extra report title exists in the pdf
                    var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(pdfFile, "YAN");
                    if (tmNo == null || buildingCode == null || buildingTmId == null)
                    {
                        Console.WriteLine($"Cannot process: {pdfFile}");
                        continue;
                    }

                    var areaId = tmNo.Split("-")[0];
                    var tmId = tmNo.Split("-")[1];
                    var reportTitle = $"TEI-B{areaId}-TM-{tmId}-YAN-M00-00";

                    // Specify which pages should contain the report title (e.g., pages 1-3)
                    var reportTitlePages = new List<int> { 1, 3 };
                    bool reportTitleFound = false;

                    foreach (var pageNum in reportTitlePages)
                    {
                        try
                        {
                            if (pdfReader.ContainsTextOnPage(pdfFile, reportTitle, pageNum, true))
                            {
                                reportTitleFound = true;
                                break;
                            }
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            // Handle case where PDF doesn't have the specified page
                            Console.WriteLine($"Incorrect pages: {pdfFile}");
                        }
                    }

                    if (!reportTitleFound)
                    {
                        Console.WriteLine($"Incorrect report title on pages {string.Join(", ", reportTitlePages)}: {pdfFile}");
                    }
                }
            }

            Console.WriteLine("DONE!!");
        }
    }
}
