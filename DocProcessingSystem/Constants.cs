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

        public static readonly Dictionary<string, string> TmNoToName = new()
        {
            {"01-01", "Hadımköy"},
            {"01-02", "Küçükköy"},
            {"01-03", "İkitelli"},
            {"01-04", "Alibeyköy"},
            {"01-06", "Etiler"},
            {"01-07", "Topkapı GİS"},
            {"01-08", "Veliefendi GİS"},
            {"01-09", "Büyükçekmece"},
            {"01-10", "Ambarlı"},
            {"01-11", "Şişli GİS"},
            {"01-12", "Kasımpaşa GİS"},
            {"01-13", "Beylikdüzü"},
            {"01-14", "Maslak"},
            {"01-15", "Silivri"},
            {"01-16", "Bağcılar GİS"},
            {"01-17", "Sultanmurat GİS"},
            {"01-18", "Levent GİS"},
            {"01-19", "Habibler"},
            {"01-20", "Atışalanı"},
            {"01-21", "Yıldıztepe GİS"},
            {"01-22", "Bahçelievler GİS"},
            {"01-23", "Taşoluk"},
            {"01-24", "Bahçeşehir"},
            {"01-25", "Davutpaşa GİS"},
            {"01-26", "Zekeriyaköy GİS"},
            {"01-27", "Altıntepe GİS"},
            {"01-28", "Yenibosna GİS"},
            {"01-29", "Yenikapı GİS"},
            {"01-30", "Kayabaşı TOKİ GİS"},
            {"01-31", "Esenyurt GİS"},
            {"01-32", "Çağlayan GİS"},
            {"01-33", "Deliklikaya GİS"},
            {"01-34", "Yakuplu GİS"},
            {"01-35", "Sağmalcılar GİS"},
            {"01-37", "Aksaray İstanbul"},
            {"01-38", "Çatalca"},
            {"01-40", "Beşyüzevler GİS"},
            {"01-41", "Taşoluk GİS"},
            {"04-01", "Kartal GİS"},
            {"04-02", "Ümraniye"},
            {"04-03", "K.Bakkalköy GİS"},
            {"04-04", "Tepeören"},
            {"04-05", "Dudullu"},
            {"04-06", "Tuzla"},
            {"04-07", "Şile"},
            {"04-08", "Kurtköy"},
            {"04-09", "Soğanlık GİS"},
            {"04-10", "Paşaköy"},
            {"04-11", "Vaniköy GİS"},
            {"04-12", "Selimiye GİS"},
            {"04-13", "Göztepe GİS"},
            {"04-14", "Gebze OSB"},
            {"04-15", "İsaköy"},
            {"04-16", "İçmeler"},
            {"04-17", "Büyükbakkalköy"},
            {"04-18", "Beykoz GİS"},
            {"04-19", "Cumhuriyet"},
            {"04-20", "Çolakoğlu 380"},
            {"04-21", "Maltepe GİS"},
            {"04-22", "Kadıköy GİS"},
            {"04-23", "Makina OSB"},
            {"04-24", "Dudullu Metro GİS"},
            {"04-25", "Ataşehir GİS"},
            {"04-26", "Kavacık GİS"},
            {"04-27", "Diliskelesi"},
            {"04-28", "Dilovası"},
            {"04-29", "Bostancı M. GİS"},
            {"05-01", "Karadeniz Ereğli"},
            {"05-02", "Karamürsel"},
            {"05-03", "Köseköy (İzmit 2)"},
            {"05-05", "Bölücek (Ereğli 2)"},
            {"05-06", "Adapazarı"},
            {"05-07", "Gerkonsan"},
            {"05-08", "Bartın"},
            {"05-09", "Hendek"},
            {"05-10", "Kaynarca"},
            {"05-11", "Mudurnu"},
            {"05-12", "Çates M. Demirel"},
            {"05-13", "Osmanca"},
            {"05-14", "Zonguldak"},
            {"05-15", "Akçakoca"},
            {"05-16", "Kaynaşlı"},
            {"05-17", "İzmit GİS"},
            {"05-18", "Hereke"},
            {"05-19", "Yarımca 2"},
            {"05-20", "Bolu 2"},
            {"05-22", "Çaycuma"},
            {"05-23", "Karasu"},
            {"05-24", "Pamukova"},
            {"05-25", "Sakarya"},
            {"05-26", "Göynük"},
            {"05-27", "Kuzuluk"},
            {"05-28", "Kozlu"},
            {"05-29", "Gölcük"},
            {"05-30", "Arslanbey"},
            {"05-31", "Yarımca 1"},
            {"05-32", "Kocaeli (İzmit 380)"},
            {"05-33", "Bolu 1"},
            {"05-34", "Bartın OSB"},
            {"05-35", "Alaplı"},
            {"05-36", "Düzce OSB"},
            {"05-37", "Cumayeri"},
            {"05-38", "Ferizli OSB"},
            {"05-39", "Sapanca"},
            {"05-40", "Devrek"},
            {"12-01", "Gaziantep"},
            {"12-02", "Çırçıp"},
            {"12-03", "Şanlıurfa"},
            {"12-04", "Telhamut"},
            {"12-05", "Elbistan TES"},
            {"12-06", "Kahramanmaraş"},
            {"12-07", "Sarıl PS (PS-5)"},
            {"12-08", "Viranşehir"},
            {"12-09", "Narlı"},
            {"12-10", "Yardımcı"},
            {"12-11", "Siverek"},
            {"12-12", "Akçakale"},
            {"12-13", "Suruç"},
            {"12-14", "Atatürk (Karababa)"},
            {"12-15", "Karacadağ PS"},
            {"12-16", "Fevzipaşa"},
            {"12-17", "Araban PS (PS-4b)"},
            {"12-18", "Şanlıurfa Çimento"},
            {"12-19", "Atatürk HES"},
            {"12-20", "Göksun"},
            {"12-22", "Başpınar OSB"},
            {"12-23", "Çobanbeyli"},
            {"12-24", "İbrahimli"},
            {"12-25", "Andırın"},
            {"12-26", "Göksun SKM"},
            {"12-27", "Kilis"},
            {"12-28", "Kılavuzlu"},
            {"12-29", "Yeşilvadi"},
            {"12-31", "Şanlıurfa OSB"},
            {"12-32", "Kılılı"},
            {"12-33", "Sırrın (Şanlıurfa-2)"},
            {"12-34", "Doğanköy"},
            {"12-35", "Tatarhöyük"},
            {"12-36", "Karaca"},
            {"12-37", "Çöçelli (Gaski P3)"},
            {"12-38", "Beykent OSB"},
            {"12-39", "Pekmezli"},
            {"12-40", "Hilvan"},
            {"12-41", "Karakeçili"},
            {"12-42", "Belkıs"},
            {"12-43", "Ödülalan"},
            {"12-44", "Elgün (Viranşehir-2)"},
            {"12-45", "Çağlayan Havza"},
            {"12-46", "Yazgüneşi"},
            {"12-47", "Şehitkamil OSB"},
            {"12-48", "Zeugma"},
            {"12-49", "Kırlık"},
            {"12-50", "Abdülhamit Han"},
            {"12-51", "Kapıçam"},
            {"12-52", "Karaköprü"},
            {"12-53", "Birecik"},
            {"12-54", "Elbeyli"},
            {"12-55", "K.Maraş OSB"},
            {"12-56", "Taşlıca"},
            {"12-57", "Kuzeyşehir"},
            {"13-01", "Maden 2"},
            {"13-04", "Malatya (Malatya 1)"},
            {"13-05", "Adıyaman"},
            {"13-06", "Bingöl"},
            {"13-07", "Adıyaman Çim."},
            {"13-08", "Darende"},
            {"13-09", "Karakaya HES"},
            {"13-10", "Hasan Çelebi"},
            {"13-11", "Malorsa"},
            {"13-12", "Adıyaman Gölbaşı"},
            {"13-13", "Özal (Malatya 2)"},
            {"13-14", "Tunceli"},
            {"13-15", "Hankendi (Elazığ 3)"},
            {"13-16", "Elazığ (Elazığ 2)"},
            {"13-17", "Bizna Havza"},
            {"13-18", "Sincik Havza"},
            {"13-19", "Göynük Havza"},
            {"13-20", "Ferrokrom Elazığ"},
            {"13-21", "Arapgir Canpolat"},
            {"13-22", "Kiğı Havza"},
            {"13-23", "Kömürhan"},
            {"13-24", "Karakoçan"},
            {"13-25", "Kahta"},
            {"13-26", "Yazıhan Havza"},
            {"13-27", "Hazar"},
            {"18-01", "Tarsus"},
            {"18-02", "Seyhan HES"},
            {"18-04", "İncirlik"},
            {"18-05", "Ceyhan 1"},
            {"18-06", "Osmaniye"},
            {"18-08", "Yakaköy"},
            {"18-09", "Payas"},
            {"18-10", "Toroslar"},
            {"18-11", "Kozan"},
            {"18-12", "Taşucu"},
            {"18-13", "Akbelen"},
            {"18-14", "Zeytinli"},
            {"18-15", "Anamur"},
            {"18-16", "Erzin"},
            {"18-17", "Yumurtalık"},
            {"18-18", "Karaisalı"},
            {"18-19", "İskenderun 2"},
            {"18-20", "İkizler"},
            {"18-21", "Misis"},
            {"18-22", "Cihadiye"},
            {"18-23", "Karahan"},
            {"18-24", "Bahçe"},
            {"18-25", "Nacarlı"},
            {"18-26", "Antakya 2"},
            {"18-27", "Adana"},
            {"18-28", "Erdemli"},
            {"18-29", "Mersin 2"},
            {"18-30", "Berke"},
            {"18-31", "Güney Adana"},
            {"18-32", "Kadirli"},
            {"18-33", "Ceyhan 2"},
            {"18-34", "İskenderun 3"},
            {"18-35", "Osmaniye OSB"},
            {"18-36", "Yeni Şehitlik"},
            {"18-37", "Yüreğir"},
            {"18-38", "Feke Havza"},
            {"18-39", "Otluca HES"},
            {"18-40", "Kozan Havza"},
            {"18-42", "Reyhanlı"},
            {"18-43", "Hatay"},
            {"18-44", "Mersin Termik"},
            {"18-45", "Mersin 380"},
            {"18-46", "Kuzeytepe"},
            {"18-47", "Kuzey Adana GİS"},
            {"18-48", "Kızkalesi"},
            {"18-49", "İskenderun 1"},
            {"18-50", "Gülnar Havza"},
            {"18-51", "Çamlıyayla Havza"},
            {"18-52", "Samandağ"},
            {"18-53", "Aladağ"},
            {"18-54", "Mersin 3"},
            {"18-55", "Osmaniye 2"},
            {"18-56", "Mut"},
            {"18-57", "Misis OSB"},
            {"18-58", "Mihmandar GİS"},
            {"20-01", "Çorlu"},
            {"20-02", "Babaeski"},
            {"20-03", "Çerkezköy"},
            {"20-04", "Tekirdağ"},
            {"20-05", "Malkara"},
            {"20-06", "Büyükkarıştıran"},
            {"20-07", "Kırklareli"},
            {"20-08", "Hamitabat DGKÇS"},
            {"20-09", "Uzunköprü"},
            {"20-10", "Tegesan"},
            {"20-11", "Edirne Çimento"},
            {"20-12", "Lüleburgaz"},
            {"20-13", "Havsa"},
            {"20-14", "Botaş"},
            {"20-15", "Kıyıköy"},
            {"20-16", "Keşan"},
            {"20-17", "Ulaş"},
            {"20-18", "Enez"},
            {"20-19", "Çerkezköy OSB"},
            {"20-20", "Estergon (Edirne 2)"},
            {"20-21", "Türkgücü (Çorlu 2)"},
            {"20-22", "Şarköy"},
            {"20-23", "Velimeşe"},
            {"20-24", "Demirköy"},
            {"20-25", "Pınarhisar"},
            {"20-26", "Edirne GİS"},
            {"20-27", "Vize Havza"},
            {"20-28", "Hayrabolu"},
            {"20-29", "Pagder OSB" },
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
