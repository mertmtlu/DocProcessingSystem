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
                { 9, "KOMPRESSOR" },
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
    }
}
