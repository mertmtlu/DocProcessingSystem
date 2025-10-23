using DocProcessingSystem.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem.Services
{
    public static class Fırat
    {
        private const string ANALYSISFOLDER = @"C:\Users\Mert\Documents\TPProjects";
        private const string TEMPORARYFOLDER = @"C:\Users\Mert\Desktop\temporary";

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
