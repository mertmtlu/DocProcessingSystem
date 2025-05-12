using DocProcessingSystem.Core;
using DocProcessingSystem.Services;
using Org.BouncyCastle.Asn1.Cmp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem.Models
{
    public static class TPHelper
    {
        public static void Merge(string root)
        {
            var groupped = FolderHelper.GroupFolders(root);
            
            var options = new MergeOptions
            {
                PreserveBookmarks = true
            };

            using (var converter = new WordToPdfConverter())
            using (var merger = new PdfMergerService())
            {
                foreach (var folder in groupped)
                {
                    var mainWord = Directory.GetFiles(folder.MainFolder, "TEI*.docx", SearchOption.AllDirectories).FirstOrDefault();
                    var ekA = Directory.GetFiles(folder.MainFolder, "EK-A.pdf", SearchOption.AllDirectories).FirstOrDefault();
                    var ekB = Directory.GetFiles(folder.MainFolder, "EK-B.pdf", SearchOption.AllDirectories).FirstOrDefault();
                    
                    if (mainWord == null || ekA == null || ekB == null)
                    {
                        Console.WriteLine($"Cannot process: {folder.MainFolder}");
                        continue;
                    }
                    
                    var wordName = Path.GetFileNameWithoutExtension(mainWord);

                    converter.Convert(mainWord, mainWord.Replace(wordName, "main"), true);
                    merger.MergePdf(mainWord.Replace(wordName, "main"), new List<string>() { ekA, ekB }, mainWord.Replace(".docx", ".pdf"), options);
                }
            }
        }
    }
}
