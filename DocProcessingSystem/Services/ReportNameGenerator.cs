using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace DocProcessingSystem.Services
{
    public class ReportNameGenerator
    {
        // Dictionary #1: Code to Building Name
        private static readonly Dictionary<string, string> CodeToName = new Dictionary<string, string>
        {
            { "#1", "Kumanda Binası" },
            { "#2", "Kapalı Şalt Binası" },
            { "#3", "Metal Clad Binası" },
            { "#4", "Kumanda + Metal Clad Binası" },
            { "#5", "Röle Binası" },
            { "#6", "Telekom Binası" },
            { "#7", "GİS-154 Binası" },
            { "#8", "GİS-400 Binası" },
            { "#9", "Kompresör Binası" },
            { "#10", "Güvenlik Binası" },
            { "#11", "Trafo Binası" },
            { "#13", "Hizmet Binası" },
            { "#19", "İdari Bina" }
        };

        // Dictionary #2: Report Types
        private static readonly Dictionary<string, string> ReportTypes = new Dictionary<string, string>
        {
            { "SEL", "SEL-TAŞKIN RİSKİ DEĞERLENDİRME RAPORU" },
            { "CIG", "ÇIĞ RİSKİ DEĞERLENDİRME RAPORU" },
            { "HEY", "HEYELAN RİSKİ DEĞERLENDİRME RAPORU" },
            { "SES", "SES ÖLÇÜM DEĞERLENDİRME RAPORU" },
            { "YAN", "YANGIN RİSKİ DEĞERLENDİRME RAPORU" },
            { "GUV", "GÜVENLİK RİSKİ DEĞERLENDİRME RAPORU" },
            { "SLT", "ŞALT SAHASI TEKNİK İNCELEME RAPORU" },
            { "IKL", "İKLİM DEĞİŞİKLİĞİ DEĞERLENDİRME RAPORU" },
            { "FAY", "DİRİFAY RİSKİ DEĞERLENDİRME RAPORU" },
            { "ALT", "SAHA DEĞERLENDİRME RAPORU" },
            { "TSU", "TSUNAMİ RİSKİ DEĞERLENDİRME RAPORU" },
            { "FOY", "YER SEÇİM FÖYÜ" },
            { "ZEV", "ZEMİN VE TEMEL ETÜDÜ VERİ RAPORU" },
            { "GEO", "ZEMİN VE TEMEL ETÜDÜ GEOTEKNİK RAPORU" },
            { "RED", "AFET RİSK ENVANTERİ DEĞERLENDİRME RAPORU" }

        };

        /// <summary>
        /// Generates the formatted report name based on the filename and TM Name.
        /// </summary>
        /// <param name="fileName">The input filename (e.g., TEI-B1-TM-2-DIR-M1-0)</param>
        /// <param name="tmName">The raw name of the TM (e.g., "Sincan")</param>
        /// <returns>Formatted string or null if regex doesn't match</returns>
        public static string GenerateReportName(string fileName)
        {
            // 1. Regex Definition
            // Corresponds to: r"TEI-B(\d+)-TM-(\d+)-([A-Z]{3})-([MA])(\d+)-(\d+)*"
            string pattern = @"TEI-B(\d+)-TM-(\d+)-([A-Z]{3})-([MA])(\d+)-(\d+)*";
            Match match = Regex.Match(fileName, pattern);

            if (match.Success)
            {
                // 2. Extract Groups
                // Note: Groups[0] is the whole match, so we start at 1
                string area = match.Groups[1].Value; // Not used in final 'name' string logic, only report_name
                string center = match.Groups[2].Value; // Not used in final 'name' string logic
                string dirCode = match.Groups[3].Value;
                string mOrA = match.Groups[4].Value;
                string mNumStr = match.Groups[5].Value;
                // string lastNum = match.Groups[6].Value; // Not used in final 'name' string logic

                string tmNo = $"{area}-{center}";
                string tmName = Constants.TmNoToName[tmNo];

                int mNum = int.Parse(mNumStr); // Convert to int to remove leading zeros (e.g., "01" -> 1)

                // 3. Prepare Base Name (Upper case with Turkish culture support)
                string upperTmName = ToTurkishUpperCase(tmName);
                string name = $"{upperTmName} TM";

                // 4. Logic Branching
                if (dirCode == "DIR")
                {
                    // Logic: name = f'{name} {self.uppercase_words(constants.CODE_TO_NAME[f"#{int(m_num)}"])} DPA RAPORU'
                    string codeKey = $"#{mNum}";

                    if (CodeToName.TryGetValue(codeKey, out string buildingName))
                    {
                        name = $"{name} {ToTurkishUpperCase(buildingName)} DEPREM PERFORMANS ANALİZİ RAPORU";
                    }
                    else throw new Exception($"");
                }
                else if (dirCode == "SRL")
                {
                    // Logic: name = f'{name} {self.uppercase_words(constants.CODE_TO_NAME[f"#{int(m_num)}"])} DPA RAPORU'
                    string codeKey = $"#{mNum}";

                    if (CodeToName.TryGetValue(codeKey, out string buildingName))
                    {
                        name = $"{name} {ToTurkishUpperCase(buildingName)} STATİK RÖLÖVE PROJESİ";
                    }
                    else throw new Exception($"");
                }
                else if (mOrA == "M")
                {
                    // Logic: name = f'{name} {report_types[dir_code]}'
                    if (ReportTypes.TryGetValue(dirCode, out string reportType))
                    {
                        name = $"{name} {reportType}";
                    }
                    else throw new Exception($"");
                }
                else
                {
                    // Logic: name = f'{name} ALTERNATİF {int(m_num)} {report_types[dir_code]}'
                    if (ReportTypes.TryGetValue(dirCode, out string reportType))
                    {
                        name = $"{name} ALTERNATİF {mNum} {reportType}";
                    }
                    else throw new Exception($"");
                }

                return name;
            }

            throw new Exception($"");
        }

        // Helper to handle Turkish Uppercasing (i -> İ, ı -> I)
        private static string ToTurkishUpperCase(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;
            return input.ToUpper(new CultureInfo("tr-TR"));
        }
    }
}
