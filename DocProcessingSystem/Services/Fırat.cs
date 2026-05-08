using DocProcessingSystem.Core;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem.Services
{
    public static class Fırat
    {
        private const string ANALYSISFOLDER = @"C:\Users\Mert\Desktop\testing";
        private const string TEMPORARYFOLDER = @"C:\Users\Mert\Desktop\temp";
        private const string FIRAT_FOLDER = @"C:\Users\Mert\Desktop\Fırat";

        public static void SlicePdfs()
        {
            var ekACoverPage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CoverPages", "Parametric", "OneBlock", "EK-A_Kapak.pdf");
            var inputFolder = Path.Combine(FIRAT_FOLDER, "input");
            var outputFolder = Path.Combine(FIRAT_FOLDER, "output");
            var excelPath = Path.Combine(FIRAT_FOLDER, "data.xlsx");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(excelPath));
            var sheet = package.Workbook.Worksheets[0];

            var slicer = new PdfHorizontalSlicerService();
            using var merger = new PdfMergerService();

            for (int row = 2; row <= sheet.Dimension.End.Row; row++)
            {
                var pdfName = sheet.Cells[row, 1].Text.Trim();
                var cutsRaw = sheet.Cells[row, 2].Text.Trim().Trim('[', ']');

                if (string.IsNullOrEmpty(pdfName) || string.IsNullOrEmpty(cutsRaw))
                    continue;

                var cuts = cutsRaw.Split(',').Select(s => float.Parse(s.Trim())).ToList();
                var inputPath = Path.Combine(inputFolder, pdfName);

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Skipping '{pdfName}': file not found.");
                    continue;
                }

                var outputName = Path.GetFileNameWithoutExtension(pdfName) + "_sliced.pdf";
                var outputPath = Path.Combine(outputFolder, outputName);
                var tempSlicedPath = Path.Combine(Path.GetTempPath(), $"temp_sliced_{Guid.NewGuid()}.pdf");

                try
                {
                    Console.WriteLine($"Slicing '{pdfName}' into {cuts.Count - 1} pages...");
                    slicer.Slice(inputPath, tempSlicedPath, cuts);

                    merger.MergePdf(ekACoverPage, [tempSlicedPath], outputPath, new Core.MergeOptions
                    {
                        PreserveBookmarks = false,
                        CreateBookmarksForAdditionalPdf = false,
                    });

                    Console.WriteLine($"  -> Saved to '{outputPath}'");
                }
                finally
                {
                    if (File.Exists(tempSlicedPath))
                        File.Delete(tempSlicedPath);
                }
            }
        }

        public static void Run()
        {
            //check if temp folder exists, if not create it
            if (!Directory.Exists(TEMPORARYFOLDER))
            {
                Directory.CreateDirectory(TEMPORARYFOLDER);
            }

            // Copy all Word documents from analysis folder to temporary folder
            foreach (var word in Directory.GetFiles(ANALYSISFOLDER, "*.docx", SearchOption.AllDirectories))
            {
                File.Copy(word, Path.Combine(TEMPORARYFOLDER, Path.GetFileName(word)), true);
            }

            // Create services
            using (var converter = new WordToPdfConverter())
            using (var merger = new PdfMergerService())
            {
                // Create folder matcher
                var matcher = new FolderNameMatcher();

                // Create document handlers
                var handlers = new IDocumentTypeHandler[]
                {
                    new Post2008DocumentHandler(converter, matcher),
                };

                // Create path dictionary
                Dictionary<string, string> pathDictionary = new()
                {
                    {"Post2008", TEMPORARYFOLDER},
                };

                // Create processing manager
                using (var manager = new DocumentProcessingManager(converter, merger, handlers))
                {
                    // Process all documents
                    manager.ProcessDocuments(pathDictionary, ANALYSISFOLDER);
                }
            }
        }
    }
}
