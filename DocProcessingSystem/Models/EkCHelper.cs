using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem.Models
{
    public static class EkCHelper
    {
        /// <summary>
        /// Converts an Excel file to List<Dictionary<string, string>>
        /// First row is treated as headers/keys
        /// Empty cells are skipped
        /// Throws exception if duplicate keys are found in header row
        /// </summary>
        /// <param name="filePath">Path to the Excel file</param>
        /// <param name="worksheetIndex">Index of worksheet to read (default: 0 for first sheet)</param>
        /// <returns>List of dictionaries representing the Excel data</returns>
        public static List<Dictionary<string, string>> ConvertExcelToListOfDictionaries(string filePath, int worksheetIndex = 0)
        {
            // Set the license context for EPPlus (required for newer versions)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var result = new List<Dictionary<string, string>>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count <= worksheetIndex)
                {
                    throw new ArgumentException($"Worksheet at index {worksheetIndex} does not exist.");
                }

                var worksheet = package.Workbook.Worksheets[worksheetIndex];

                if (worksheet.Dimension == null)
                {
                    return result; // Empty worksheet
                }

                int rowCount = worksheet.Dimension.End.Row;
                int colCount = worksheet.Dimension.End.Column;

                if (rowCount < 1)
                {
                    return result; // No data
                }

                // Read headers from first row and validate for duplicates
                var headers = new Dictionary<int, string>(); // column index -> header name
                var headerSet = new HashSet<string>();

                for (int col = 1; col <= colCount; col++)
                {
                    var headerValue = worksheet.Cells[1, col].Value?.ToString()?.Trim();

                    // Skip empty headers
                    if (string.IsNullOrWhiteSpace(headerValue))
                    {
                        continue;
                    }

                    // Check for duplicate headers
                    if (headerSet.Contains(headerValue))
                    {
                        throw new InvalidOperationException($"Duplicate header found: '{headerValue}' at column {col}");
                    }

                    headers[col] = headerValue;
                    headerSet.Add(headerValue);
                }

                // Process data rows (starting from row 2)
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new Dictionary<string, string>();
                    bool hasData = false;

                    foreach (var headerEntry in headers)
                    {
                        int col = headerEntry.Key;
                        string headerName = headerEntry.Value;

                        var cellValue = worksheet.Cells[row, col].Value?.ToString()?.Trim();

                        //// Skip empty cells
                        //if (string.IsNullOrWhiteSpace(cellValue))
                        //{
                        //    continue;
                        //}

                        rowData[headerName] = cellValue;
                        hasData = true;
                    }

                    // Only add row if it contains at least one non-empty cell
                    if (hasData)
                    {
                        result.Add(rowData);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Alternative method that accepts a byte array instead of file path
        /// Useful when working with uploaded files or streams
        /// </summary>
        /// <param name="excelData">Excel file as byte array</param>
        /// <param name="worksheetIndex">Index of worksheet to read (default: 0 for first sheet)</param>
        /// <returns>List of dictionaries representing the Excel data</returns>
        public static List<Dictionary<string, string>> ConvertExcelToListOfDictionaries(byte[] excelData, int worksheetIndex = 0)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var result = new List<Dictionary<string, string>>();

            using (var stream = new MemoryStream(excelData))
            using (var package = new ExcelPackage(stream))
            {
                if (package.Workbook.Worksheets.Count <= worksheetIndex)
                {
                    throw new ArgumentException($"Worksheet at index {worksheetIndex} does not exist.");
                }

                var worksheet = package.Workbook.Worksheets[worksheetIndex];

                if (worksheet.Dimension == null)
                {
                    return result;
                }

                int rowCount = worksheet.Dimension.End.Row;
                int colCount = worksheet.Dimension.End.Column;

                if (rowCount < 1)
                {
                    return result;
                }

                var headers = new Dictionary<int, string>();
                var headerSet = new HashSet<string>();

                for (int col = 1; col <= colCount; col++)
                {
                    var headerValue = worksheet.Cells[1, col].Value?.ToString()?.Trim();

                    if (string.IsNullOrWhiteSpace(headerValue))
                    {
                        continue;
                    }

                    if (headerSet.Contains(headerValue))
                    {
                        throw new InvalidOperationException($"Duplicate header found: '{headerValue}' at column {col}");
                    }

                    headers[col] = headerValue;
                    headerSet.Add(headerValue);
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new Dictionary<string, string>();
                    bool hasData = false;

                    foreach (var headerEntry in headers)
                    {
                        int col = headerEntry.Key;
                        string headerName = headerEntry.Value;

                        var cellValue = worksheet.Cells[row, col].Value?.ToString()?.Trim();

                        if (string.IsNullOrWhiteSpace(cellValue))
                        {
                            continue;
                        }

                        rowData[headerName] = cellValue;
                        hasData = true;
                    }

                    if (hasData)
                    {
                        result.Add(rowData);
                    }
                }
            }

            return result;
        }

        public static List<BuildingData> GroupByBuildings(List<Dictionary<string, string>> data)
        {
            List<BuildingData> building = new();

            foreach (var row in data)
            {
                var existing = building.FirstOrDefault(b => b.TmNo == row["TM No"] && b.TmName == row["TM Adı"] && b.BuildingName == row["Bina"]);

                if (existing == null)
                {
                    if (row["Donatı1"] == null) continue;
                    existing = new(row["TM No"], row["TM Adı"], row["Bina"]);
                    building.Add(existing);
                }
                existing.Data.Add(row);
            }

            return building;

        }

        public static List<BuildingData> GroupByBuildings(string filePath)
        {
            return GroupByBuildings(ConvertExcelToListOfDictionaries(filePath));
        }

        public static void GenerateExcelReport(string filePath, Dictionary<int, string> baseFileDict, string outputDest)
        {
            GenerateExcelReport(GroupByBuildings(filePath), baseFileDict, outputDest);
        }

        private static (string, string) GetConditions(BuildingData buildingData)
        {
            string rontgen = "(*) Saha çalışmalarından elde edilen tutanaklar kullanılarak özet rapor olarak hazırlanmıştır.";
            string sıyırma = "";
            
            var rontgenDatas = buildingData.Data;

            var columnDimensions = rontgenDatas
                .Where(dict => dict.ContainsKey("Kısa Kenar R") && dict.ContainsKey("Uzun Kenar R"))
                .Select(dict => dict["Kısa Kenar R"] + "/" + dict["Uzun Kenar R"])
                .Where(value => !string.IsNullOrEmpty(value))
                .Distinct()
                .ToList();

            var lateralReBarSpacings = rontgenDatas 
                .Where(dict => dict.ContainsKey("Etriye2"))
                .Select(dict => dict["Etriye2"])
                .Where(value => !string.IsNullOrEmpty(value) && value != "-")
                .Select(int.Parse)
                .Distinct()
                .ToList();

            var lateralReBarDiameters = rontgenDatas
                .Where(dict => dict.ContainsKey("Etriye1"))
                .Select(dict => dict["Etriye1"])
                .Where(value => !string.IsNullOrEmpty(value))
                .Select(int.Parse)
                .Distinct()
                .ToList();

            var longitudinalReBarDiameters = rontgenDatas
                .Where(dict => dict.ContainsKey("Donatı2"))
                .Select(dict => dict["Donatı2"])
                .Where(value => !string.IsNullOrEmpty(value))
                .Select(int.Parse)
                .Distinct()
                .ToList();

            var longitudinalReBarDiametersSıyırma = rontgenDatas
                .Where(dict => dict.ContainsKey("Düşey Donatı S"))
                .Select(dict => dict["Düşey Donatı S"])
                .Where(value => !string.IsNullOrEmpty(value))
                .Select(int.Parse)
                .Distinct()
                .ToList();

            var longitudinalReBarFy = rontgenDatas
                .Where(dict => dict.ContainsKey("Fy Enine"))
                .Select(dict => dict["Fy Enine"])
                .Where(value => !string.IsNullOrEmpty(value))
                .Select(int.Parse)
                .Distinct()
                .ToList();

            var longitudinalReBarCount = rontgenDatas
                .Where(dict => dict.ContainsKey("Donatı1"))
                .Select(dict => dict["Donatı1"])
                .Where(value => !string.IsNullOrEmpty(value) && value != "Ø")
                .Select(int.Parse)
                .Distinct()
                .ToList();

            if (lateralReBarSpacings.Count > 1)
            {
                rontgen += $"\r\n\r\nEtriye aralıklarında farklılıklar gözlemlenmiş olup, " +
                    $"güvenli tarafta kalmak adına {lateralReBarSpacings.Max()} cm " +
                    $"olarak alınmıştır.";
            }

            if (longitudinalReBarCount.Count > 1 && columnDimensions.Count == 1)
            {
                rontgen += $"\r\n\r\nBoyuna donatı adetlerinde farklılıklar gözlemlenmiş olup, " +
                    $"güvenli tarafta kalmak adına {longitudinalReBarCount.Min()}Ø{longitudinalReBarDiameters.Min()} " +
                    $"olarak alınmıştır.";
            }
            else if (longitudinalReBarCount.Count != 1 && columnDimensions.Count != 1) throw new Exception();

            if (longitudinalReBarFy.Count > 1)
            {
                var a = longitudinalReBarFy.Min() == 220 ? "DÜZ" : "NERVÜLÜ";

                sıyırma += $"Boyuna donatı tiplerinde farklılıklar gözlemlenmiş olup, " +
                    $"güvenli tarafta kalmak adına hepsi " +
                    $"{a} donatı olarak kabul edilmiştir.\r\n\r\n" +
                    $"(*) Saha çalışmalarından elde edilen tutanaklar kullanılarak özet rapor olarak hazırlanmıştır. ";
            }
            else if (longitudinalReBarFy.Count == 1)
            {
                var a = longitudinalReBarFy.Min() == 220 ? "DÜZ" : "NERVÜLÜ";
                sıyırma += $"Sıyırma işlemlerinde tespit edilen bütün donatılar {a} donatıdır. ";
            }
            else throw new Exception();

            if (lateralReBarDiameters.Count > 1)
            {
                sıyırma += $"\r\n\r\nEtriye çaplarında farklılıklar gözlemlenmiş olup, " +
                    $"güvenli tarafta kalmak adına hepsi Ø{lateralReBarDiameters.Min()} " +
                    $"olarak kabul edilmiştir.";
            }

            if (longitudinalReBarDiametersSıyırma.Count > 1)
            {
                sıyırma += $"\r\n\r\nDüşey donatı çaplarında farklılıklar gözlemlenmiş olup, " +
                    $"güvenli tarafta kalmak adına hepsi Ø{longitudinalReBarDiametersSıyırma.Min()} " +
                    $"olarak kabul edilmiştir.";
            }

            sıyırma += "\r\n\r\n(*) Saha çalışmalarından elde edilen tutanaklar kullanılarak özet rapor olarak hazırlanmıştır. ";




            return (rontgen, sıyırma);
        }

        /// <summary>
        /// Enhanced version of GenerateExcelReport using the SimpleExcelWrapper
        /// </summary>
        public static void GenerateExcelReport(List<BuildingData> datas, Dictionary<int, string> baseFileDict, string outputDest)
        {
            foreach (var buildingData in datas)
            {
                var baseFile = baseFileDict.TryGetValue(buildingData.Data.Count, out var value) ? value : "";
                if (baseFile == "")
                {
                    Console.WriteLine($"Skipping due to data count: {buildingData.Data.Count}, TM NO: {buildingData.TmNo}, Building: {buildingData.BuildingName}");
                    continue;
                }

                var folderName = $"{buildingData.TmNo}_M{Constants.NameToCode[buildingData.BuildingName]}";
                var excelName = "_DONATI_RAPOR.xlsm";
                var outputFolder = Path.Combine(outputDest, folderName);
                var outputFile = Path.Combine(outputFolder, excelName);

                // Create output directory if it doesn't exist
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // Use the simple wrapper for easy cell access
                using (var excel = new SimpleExcelWrapper(baseFile))
                {
                    // Example: Fill in building information
                    excel["A6"] = buildingData.TmNo;
                    excel["D6"] = buildingData.TmName + " TM";
                    excel["M6"] = buildingData.BuildingName;

                    // Process each row of data
                    int currentRow = 11; // Start from row 11

                    var (rontgen, sıyırma) = GetConditions(buildingData);

                    foreach (var dataRow in buildingData.Data)
                    {
                        // Adjust these mappings based on your actual data structure and target layout
                        foreach (var kvp in dataRow)
                        {
                            switch (kvp.Key)
                            {
                                case "Röntgen  No":
                                    excel[$"B{currentRow}"] = kvp.Value;
                                    break;
                                case "Aks1R":
                                    excel[$"C{currentRow}"] = kvp.Value;
                                    break;
                                case "AksRSlash":
                                    excel[$"D{currentRow}"] = kvp.Value;
                                    break;
                                case "Aks2R":
                                    excel[$"E{currentRow}"] = kvp.Value;
                                    break;
                                case "Kat":
                                    excel[$"G{currentRow}"] = kvp.Value switch
                                    {
                                        "Zemin" => "Zemin Kat",
                                        "" => "Zemin Kat",
                                        null => "Zemin Kat",
                                        _ => kvp.Value  // default case
                                    };
                                    break;
                                case "Yapı Elemanı R":
                                    excel[$"H{currentRow}"] = kvp.Value;
                                    break;
                                case "Kısa Kenar R":
                                    excel[$"I{currentRow}"] = kvp.Value;
                                    break;
                                case "KesitSlash":
                                    excel[$"J{currentRow}"] = kvp.Value;
                                    break;
                                case "Uzun Kenar R":
                                    excel[$"K{currentRow}"] = kvp.Value;
                                    break;
                                case "Donatı1":
                                    excel[$"L{currentRow}"] = kvp.Value;
                                    break;
                                case "DonatıFi":
                                    excel[$"M{currentRow}"] = kvp.Value;
                                    break;
                                case "Donatı2":
                                    excel[$"N{currentRow}"] = kvp.Value;
                                    break;
                                case "Uzun Kenar":
                                    excel[$"O{currentRow}"] = kvp.Value;
                                    break;
                                case "Kısa Kenar":
                                    excel[$"P{currentRow}"] = kvp.Value;
                                    break;
                                case "EtriyeFi":
                                    excel[$"Q{currentRow}"] = kvp.Value;
                                    break;
                                case "Etriye1":
                                    excel[$"R{currentRow}"] = kvp.Value;
                                    break;
                                case "EtriyeSlash":
                                    excel[$"S{currentRow}"] = kvp.Value;
                                    break;
                                case "Etriye2":
                                    excel[$"T{currentRow}"] = kvp.Value;
                                    break;


                                default:
                                    // Handle other columns dynamically or skip
                                    break;
                            }
                        }

                        currentRow++;
                    }

                    excel[$"A{currentRow + 5}"] = rontgen;
                    excel.UseWorksheet(1);

                    currentRow = 11; // Start from row 11 again
                    int deleteRow = 11 + buildingData.Data.Count - 1;
                    foreach (var dataRow in buildingData.Data)
                    {
                        if (dataRow["Sıyırma No"] == null)
                        {
                            excel.DeleteRow(deleteRow);
                            deleteRow--;
                            continue;
                        }

                        // Adjust these mappings based on your actual data structure and target layout
                        foreach (var kvp in dataRow)
                        {
                            switch (kvp.Key)
                            {
                                case "Sıyırma No":
                                    excel[$"B{currentRow}"] = kvp.Value;
                                    break;
                                case "Aks1S":
                                    excel[$"C{currentRow}"] = kvp.Value;
                                    break;
                                case "Aks2S":
                                    excel[$"E{currentRow}"] = kvp.Value;
                                    break;
                                case "Sıyırma Yapı Elemanı":
                                    excel[$"H{currentRow}"] = kvp.Value;
                                    break;
                                case "Kısa Kenar S":
                                    excel[$"I{currentRow}"] = kvp.Value;
                                    break;
                                case "Uzun Kenar S":
                                    excel[$"K{currentRow}"] = kvp.Value;
                                    break;
                                case "Düşey Donatı S":
                                    excel[$"M{currentRow}"] = kvp.Value;
                                    break;
                                case "Yatay Donatı S":
                                    excel[$"O{currentRow}"] = kvp.Value;
                                    break;
                                case "Kat":
                                    excel[$"G{currentRow}"] = kvp.Value switch
                                    {
                                        "Zemin" => "Zemin Kat",
                                        "" => "Zemin Kat",
                                        null => "Zemin Kat",
                                        _ => kvp.Value  // default case
                                    };
                                    break;
                                default:
                                    // Handle other columns dynamically or skip
                                    break;
                            }
                        }

                        currentRow++;
                    }

                    excel[$"A{currentRow + 5}"] = sıyırma;
                    excel.SaveAs(outputFile);
                }
            }
        }

        public class BuildingData
        {
            public string TmNo { get; set; }
            public string BuildingName { get; set; }
            public string TmName { get; set; }
            public List<Dictionary<string, string>> Data { get; set; } = new();

            public BuildingData(string tmNo, string tmName,string buildingName)
            {
                TmNo = tmNo;
                TmName = tmName;
                BuildingName = buildingName;
            }
        }

        /// <summary>
        /// Simple wrapper for Excel operations with easy cell access
        /// Usage: excelWrapper["A1"] = "some value"
        /// </summary>
        public class SimpleExcelWrapper : IDisposable
        {
            private ExcelPackage _package;
            private ExcelWorksheet _worksheet;
            private bool _disposed = false;

            public SimpleExcelWrapper(string templateFilePath = null)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                if (!string.IsNullOrEmpty(templateFilePath) && File.Exists(templateFilePath))
                {
                    // Load existing template
                    _package = new ExcelPackage(new FileInfo(templateFilePath));
                    _worksheet = _package.Workbook.Worksheets[0]; // Use first worksheet
                }
                else
                {
                    // Create new workbook
                    _package = new ExcelPackage();
                    _worksheet = _package.Workbook.Worksheets.Add("Sheet1");
                }
            }

            /// <summary>
            /// Simple indexer for cell access
            /// Usage: wrapper["A1"] = "value" or var value = wrapper["A1"]
            /// </summary>
            public object this[string cellAddress]
            {
                get => _worksheet.Cells[cellAddress].Value;
                set => _worksheet.Cells[cellAddress].Value = value;
            }

            /// <summary>
            /// Delete a row from the current worksheet
            /// </summary>
            public void DeleteRow(int rowNumber)
            {
                _worksheet.DeleteRow(rowNumber);
            }

            /// <summary>
            /// Access specific worksheet by index
            /// </summary>
            public SimpleExcelWrapper UseWorksheet(int index)
            {
                if (index < _package.Workbook.Worksheets.Count)
                {
                    _worksheet = _package.Workbook.Worksheets[index];
                }
                return this;
            }

            /// <summary>
            /// Access specific worksheet by name
            /// </summary>
            public SimpleExcelWrapper UseWorksheet(string name)
            {
                var ws = _package.Workbook.Worksheets[name];
                if (ws != null)
                {
                    _worksheet = ws;
                }
                return this;
            }

            /// <summary>
            /// Create a new worksheet
            /// </summary>
            public SimpleExcelWrapper AddWorksheet(string name)
            {
                _worksheet = _package.Workbook.Worksheets.Add(name);
                return this;
            }

            /// <summary>
            /// Set cell value with row/column indices (1-based)
            /// </summary>
            public void SetCell(int row, int col, object value)
            {
                _worksheet.Cells[row, col].Value = value;
            }

            /// <summary>
            /// Get cell value with row/column indices (1-based)
            /// </summary>
            public object GetCell(int row, int col)
            {
                return _worksheet.Cells[row, col].Value;
            }

            /// <summary>
            /// Set a range of values from a dictionary
            /// </summary>
            public void SetCellsFromDictionary(Dictionary<string, object> cellMappings)
            {
                foreach (var mapping in cellMappings)
                {
                    this[mapping.Key] = mapping.Value;
                }
            }

            /// <summary>
            /// Apply basic formatting to a cell
            /// </summary>
            public SimpleExcelWrapper FormatCell(string cellAddress, Action<ExcelRange> formatAction)
            {
                formatAction(_worksheet.Cells[cellAddress]);
                return this;
            }

            /// <summary>
            /// Auto-fit columns
            /// </summary>
            public SimpleExcelWrapper AutoFitColumns()
            {
                _worksheet.Cells[_worksheet.Dimension.Address].AutoFitColumns();
                return this;
            }

            /// <summary>
            /// Save to file
            /// </summary>
            public void SaveAs(string filePath)
            {
                // Ensure directory exists
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var fileInfo = new FileInfo(filePath);
                _package.SaveAs(fileInfo);
            }

            /// <summary>
            /// Save as byte array
            /// </summary>
            public byte[] SaveAsBytes()
            {
                return _package.GetAsByteArray();
            }

            public void Dispose()
            {
                if (!_disposed)
                {
                    _package?.Dispose();
                    _disposed = true;
                }
            }
        }
    }
}
