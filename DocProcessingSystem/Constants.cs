using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem
{
    public enum ReportEnum
    {
        FAYM,
        SELM,
        HEYM,
        CIGM,
        SESM,
        YANM,
        GUVM,
        SLTM,
        IKLM,
        GEOM,
        ZEVM,
        TSUM,
        FOYG,
        FOYM,
        FOYA,
        ALTA,
        DIRM,
        DGRM,
    }

    public class ReportType
    {
        public ReportEnum Type { get; }
        public string Pattern { get; }

        public ReportType(ReportEnum type, string pattern)
        {
            Type = type;
            Pattern = pattern;
        }
    }

    public class Report
    {
        public string FilePath { get; set; }
        public ReportEnum Type { get; set; }
        public string TmNo { get; set; }
        public string BuildingCode { get; set; }
        public string BuildingTmId { get; set; }

        public string FileName => Path.GetFileName(FilePath);
        public string BaseFileName => Path.GetFileNameWithoutExtension(FilePath);
    }

    public class ReportGroup
    {
        public string Identifier { get; set; }
        public List<Report> Reports { get; set; } = new();

        public List<Report> GetReportsByType(ReportEnum type)
            => Reports.Where(r => r.Type == type).ToList();

        public Report GetFirstReportOfType(ReportEnum type)
            => Reports.FirstOrDefault(r => r.Type == type);

        public bool HasReportType(ReportEnum type)
            => Reports.Any(r => r.Type == type);

        public void AddReport(Report report)
        {
            Reports.Add(report);
        }
    }

    public class ReportCollection
    {
        public ReportCollection(string root)
        {
            RootDir = root;
        }

        public List<ReportGroup> Groups { get; set; } = new();

        public string RootDir { get; set; } = string.Empty;

        public List<string> GetIdentifiers() => Groups.Select(g => g.Identifier).ToList();

        public ReportGroup GetGroup(string identifier)
            => Groups.FirstOrDefault(g => g.Identifier == identifier);

        public ReportGroup GetOrCreateGroup(string identifier)
        {
            var group = GetGroup(identifier);
            if (group == null)
            {
                group = new ReportGroup { Identifier = identifier };
                Groups.Add(group);
            }
            return group;
        }

        public Dictionary<string, Dictionary<ReportEnum, string>> RegroupByTm()
        {
            Dictionary<string, Dictionary<ReportEnum, string>> result = new();

            var identifiers = GetIdentifiers();

            foreach (var ident in identifiers)
            {
                var group = GetGroup(ident);

                Dictionary<ReportEnum, string> keyValuePairs = new();

                foreach (var r in group.Reports)
                {
                    keyValuePairs.Add(r.Type, r.FilePath);
                }

                result.Add(ident, keyValuePairs);
            }

            return result;
        }

        public List<Report> GetAllReportsOfType(ReportEnum type)
            => Groups.SelectMany(g => g.GetReportsByType(type)).ToList();
    }

    public static class Constants
    {
        public static readonly Dictionary<int, string> CodeToName = new()
        {
            { 1, "KUMANDA" },
            { 2, "KAPALI SALT" },
            { 3, "METALCLAD" },
            { 4, "KUMANDA+MC" },
            { 5, "ROLE" },
            { 6, "TELEKOM" },
            { 7, "GIS-154" },
            { 8, "GIS-400" },
            { 9, "KOMPRESOR" },
            { 10, "GUVENLIK" },
            { 13, "HIZMET" },
            //{ 19, "TRAFO" }, // This is changed
            { 11, "TRAFO" }
        };

        public static readonly List<ReportType> ReportTypes = new()
        {
            new(ReportEnum.FAYM, "FAY-M"),
            new(ReportEnum.SELM, "SEL-M"),
            new(ReportEnum.HEYM, "HEY-M"),
            new(ReportEnum.CIGM, "CIG-M"),
            new(ReportEnum.SESM, "SES-M"),
            new(ReportEnum.YANM, "YAN-M"),
            new(ReportEnum.GUVM, "GUV-M"),
            new(ReportEnum.SLTM, "SLT-M"),
            new(ReportEnum.IKLM, "IKL-M"),
            new(ReportEnum.GEOM, "GEO-M"),
            new(ReportEnum.ZEVM, "ZEV-M"),
            new(ReportEnum.TSUM, "TSU-M"),
            new(ReportEnum.FOYG, "FOY-G"),
            new(ReportEnum.FOYM, "FOY-M"),
            new(ReportEnum.FOYA, "FOY-A"),
            new(ReportEnum.ALTA, "ALT-A"),
            new(ReportEnum.DIRM, "DIR-M"),
            new(ReportEnum.DGRM, "DGR-M")
        };

        public static readonly Dictionary<string, string> ReportType = new()
        {
            { "ZEV", "ZEMIN ETUT-VERI"},
            { "GEO", "ZEMIN ETUT-GEOTEKNIK"},
            { "FAY", "DIRIFAY"},
            { "IKL", "IKLIM DEGISIKLIGI"},
            { "RED", "AFET RISK ENVANTERI DEGERLENDIRME"},
            { "FOY", "YER SECIM FOYU"},
            { "SLT", "SALT INCELEME"},
        };

        public static readonly List<string> preferences = new()
        {
            "IKL",
            "GEO",
            "FAY",
            "ZEV",
            "RED",
            "FOY",
            "SLT",
        };

        public static readonly Dictionary<int, List<string>> requiredFiles = new()
        {
            { 1, new List<string>
              {
                  "EK-A.pdf",
                  "EK-B.pdf",
                  "EK-C.pdf",
                  "EK-D.pdf"
              }
            },
            { 2, new List<string>
              {
                  "EK-A.pdf",
                  "EK-B.pdf",
                  "EK-C.pdf",
                  "EK-D.pdf",
                  "EK-E.pdf"
              }
            }
        };

        public static readonly Dictionary<string, string> NameToCode = new()
        {
            { "Kumanda Binası", "01_01"},
            { "Kapalı Şalt Binası", "02_01"},
            { "Kapalı Şalt Binası-1", "02_01"},
            { "Kapalı Şalt Binası-2", "02_02"},
            { "Metalclad Binası", "03_01"},
            { "Kumanda+MC Binası", "04_01"},
            { "Röle Binası", "05_01"},
            { "Güvenlik Binası", "10_01"},
            { "Güvenlik Binası-2", "10_02"},
        };

        public static readonly List<string> EkCPdfCheck = new()
        {
            "_002127_",
            "_002125_",
            "_002126_",
            "_001971_",
            "_001972_",
            "_001973_",
            "_001974_",
            "_001975_",
            "_001979_",
            "_001976_",
            "_001977_",
            "_001978_",
            "_002061_",
            "_002062_",
            "_002063_",
            "_002064_",
            "_002065_",
            "_002066_",
            "_002067_",
            "_002039_",
            "_002040_",
            "_002044_",
            "_002041_",
            "_002042_",
            "_002043_",
            "_001968_",
            "_001969_",
            "_001970_",
            "_001967_",
            "_002070_",
            "_002071_",
            "_002072_",
            "_002068_",
            "_002073_",
            "_002074_",
            "_002075_",
            "_002076_",
            "_002077_",
            "_002078_",
            "_002045_",
            "_002046_",
            "_002047_",
            "_002137_",
            "_002138_",
            "_002139_",
            "_002141_",
            "_002140_",
            "_001964_",
            "_001965_",
            "_001966_",
            "_001994_",
            "_001995_",
            "_001996_",
            "_001992_",
            "_001993_",
            "_001997_",
            "_001998_",
            "_001999_",
            "_002000_",
            "_002001_",
            "_002002_",
            "_002131_",
            "_002132_",
            "_002133_",
            "_002128_",
            "_002129_",
            "_002130_",
            "_002135_",
            "_002136_",
            "_002134_",
            "_002031_",
            "_002032_",
            "_002033_",
            "_002034_",
            "_002025_",
            "_002026_",
            "_002027_",
            "_002024_",
            "_002029_",
            "_002030_",
            "_002021_",
            "_002022_",
            "_002023_",
            "_002013_",
            "_002014_",
            "_002019_",
            "_002020_",
            "_002015_",
            "_002016_",
            "_002017_",
            "_002018_",
            "_001985_",
            "_001991_",
            "_001983_",
            "_001984_",
            "_001986_",
            "_001987_",
            "_001980_",
            "_001981_",
            "_001982_",
            "_002006_",
            "_002007_",
            "_002008_",
            "_002003_",
            "_002004_",
            "_002005_",
            "_002035_",
            "_002036_",
            "_002037_",
            "_002038_",
            "_002010_",
            "_002011_",
            "_002012_",
            "_002098_",
            "_002099_",
            "_002100_",
            "_002096_",
            "_002097_",
            "_002101_",
            "_002095_",
            "_002102_",
            "_002103_",
            "_002104_",
            "_002089_",
            "_002090_",
            "_002091_",
            "_002085_",
            "_002086_",
            "_002087_",
            "_002088_",
            "_002092_",
            "_002093_",
            "_002094_",
            "_002155_",
            "_002156_",
            "_002153_",
            "_002154_",
            "_002119_",
            "_002120_",
            "_002121_",
            "_002116_",
            "_002117_",
        };

        public static readonly List<string> HK19 = new()
        {
            "TEI-B02-TM-16-DIR-M02-01",
            "TEI-B02-TM-18-DIR-M02-01",
            "TEI-B02-TM-25-DIR-M02-01",
            "TEI-B03-TM-01-DIR-M02-01",
            "TEI-B03-TM-01-DIR-M02-01",
            "TEI-B03-TM-19-DIR-M02-01",
            "TEI-B03-TM-20-DIR-M01-01",
            "TEI-B06-TM-07-DIR-M02-01",
            "TEI-B07-TM-01-DIR-M02-02",
            "TEI-B07-TM-02-DIR-M02-01",
            "TEI-B07-TM-02-DIR-M02-02",
            "TEI-B07-TM-03-DIR-M06-01",
            "TEI-B07-TM-05-DIR-M01-02",
            "TEI-B07-TM-06-DIR-M02-01",
            "TEI-B07-TM-10-DIR-M02-01",
            "TEI-B07-TM-11-DIR-M01-01",
            "TEI-B07-TM-11-DIR-M02-01",
            "TEI-B07-TM-12-DIR-M01-01",
            "TEI-B07-TM-12-DIR-M02-01",
            "TEI-B09-TM-02-DIR-M02-01",
            "TEI-B09-TM-02-DIR-M02-02",
            "TEI-B09-TM-05-DIR-M02-01",
            "TEI-B09-TM-06-DIR-M02-02",
            "TEI-B09-TM-07-DIR-M02-01",
            "TEI-B09-TM-07-DIR-M02-02",
            "TEI-B09-TM-10-DIR-M02-01",
            "TEI-B09-TM-13-DIR-M02-01",
            "TEI-B09-TM-14-DIR-M02-01",
            "TEI-B09-TM-15-DIR-M01-01",
            "TEI-B09-TM-15-DIR-M02-01",
            "TEI-B09-TM-16-DIR-M02-01",
            "TEI-B10-TM-07-DIR-M10-01",
            "TEI-B10-TM-09-DIR-M10-01",
            "TEI-B10-TM-20-DIR-M04-01",
            "TEI-B19-TM-10-DIR-M01-01",

            //"TEI-B02-TM-07-DIR-M01-01",
            //"TEI-B02-TM-07-DIR-M02-01",
            //"TEI-B02-TM-16-DIR-M02-01",
            //"TEI-B02-TM-18-DIR-M02-01",
            //"TEI-B02-TM-25-DIR-M02-01",
            //"TEI-B02-TM-33-DIR-M02-01",
            //"TEI-B02-TM-34-DIR-M03-01",
            //"TEI-B03-TM-01-DIR-M02-01",
            //"TEI-B03-TM-01-DIR-M02-01",
            //"TEI-B03-TM-19-DIR-M02-01",
            //"TEI-B03-TM-20-DIR-M01-01",
            //"TEI-B06-TM-07-DIR-M02-01",
            //"TEI-B06-TM-15-DIR-M02-01",
            //"TEI-B07-TM-01-DIR-M02-02",
            //"TEI-B07-TM-02-DIR-M02-01",
            //"TEI-B07-TM-02-DIR-M02-02",
            //"TEI-B07-TM-03-DIR-M06-01",
            //"TEI-B07-TM-05-DIR-M01-02",
            //"TEI-B07-TM-06-DIR-M02-01",
            //"TEI-B07-TM-10-DIR-M02-01",
            //"TEI-B07-TM-11-DIR-M01-01",
            //"TEI-B07-TM-11-DIR-M02-01",
            //"TEI-B07-TM-12-DIR-M01-01",
            //"TEI-B07-TM-12-DIR-M02-01",
            //"TEI-B09-TM-02-DIR-M02-01",
            //"TEI-B09-TM-02-DIR-M02-02",
            //"TEI-B09-TM-05-DIR-M02-01",
            //"TEI-B09-TM-06-DIR-M02-02",
            //"TEI-B09-TM-07-DIR-M02-01",
            //"TEI-B09-TM-07-DIR-M02-02",
            //"TEI-B09-TM-10-DIR-M02-01",
            //"TEI-B09-TM-13-DIR-M02-01",
            //"TEI-B09-TM-14-DIR-M02-01",
            //"TEI-B09-TM-15-DIR-M01-01",
            //"TEI-B09-TM-15-DIR-M02-01",
            //"TEI-B09-TM-16-DIR-M02-01",
            //"TEI-B10-TM-07-DIR-M10-01",
            //"TEI-B10-TM-09-DIR-M10-01",
            //"TEI-B10-TM-20-DIR-M04-01",
            //"TEI-B19-TM-10-DIR-M01-01",
            //"TEI-B19-TM-12-DIR-M04-01",
            //"TEI-B21-TM-10-DIR-M02-01",
        };
    }
}
