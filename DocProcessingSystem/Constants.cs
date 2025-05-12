using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem
{
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
            { 19, "TRAFO" }
        };

        public static readonly Dictionary<string, string> ReportType = new()
        {
            { "ZEV", "ZEMIN ETUT-VERI"},
            { "GEO", "ZEMIN ETUT-GEOTEKNIK"},
            { "FAY", "DIRIFAY"},
            { "IKL", "IKLIM DEGISIKLIGI"},
        };

        public static readonly List<string> preferences = new()
        {
            "IKL",
            "GEO",
            "FAY",
            "ZEV"
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
