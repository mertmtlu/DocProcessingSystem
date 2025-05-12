using Microsoft.Office.Interop.Word;
using OfficeOpenXml.Style.Dxf;
using Org.BouncyCastle.Asn1.Cmp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem.Models
{
    public static class HakedisHelper
    {
        public static void Run()
        {
            var rootFolder = @"C:\Users\Mert\Desktop\SZL-3\SZL-3 HAKEDİŞ\HK18";
            var destFolder = @"C:\Users\Mert\Desktop\SZL-3\SZL-3 SAHA";


            var hakedisFolders = Directory.GetDirectories(rootFolder).Where(folder => !(folder.Contains("-B") || folder.Contains("-C") || folder.Contains("Kontrol")));
            var sahaFolders = Directory.GetDirectories(destFolder);

            foreach (var sahafolder in sahaFolders)
            {
                var tmFolders = Directory.GetDirectories(Path.Combine(sahafolder, "TM_Folders"));

                foreach (var tmFolder in tmFolders)
                {
                    bool found = false;

                    if (found) continue;

                    var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(tmFolder);

                    if (tmNo == null || buildingCode == null || buildingTmId == null)
                    {
                        Console.WriteLine($"Cannot process for {tmFolder}");
                        continue;
                    }

                    foreach (var hakedisFolder in hakedisFolders)
                    {
                        var (tmNoHakedis, buildingCodeHakedis, buildingTmIdHakedis) = FolderHelper.ExtractParts(hakedisFolder);

                        if (tmNoHakedis == null || buildingCodeHakedis == null || buildingTmIdHakedis == null)
                        {
                            Console.WriteLine($"Cannot process for {hakedisFolder}");
                            continue;
                        }

                        if (tmNoHakedis == tmNo && buildingCodeHakedis == buildingCode && buildingTmIdHakedis == buildingTmId)
                        {
                            found = true;

                            //var resultFolder = Path.Combine(hakedisFolder, "TBDYResults");
                            //if (!Directory.Exists(resultFolder))
                            //{
                            //    throw new ArgumentOutOfRangeException("Result folder is not found");
                            //}

                            //List<string> filesToCopy = new()
                            //{
                            //    "EK-A.pdf",
                            //    "EK-B.pdf",
                            //    "EK-C.pdf"
                            //};

                            //foreach (var fileName in filesToCopy)
                            //{
                            //    var file = Path.Combine(resultFolder, fileName);

                            //    if (!File.Exists(file))
                            //    {
                            //        Console.WriteLine($"{file} do not exists to copy!!!");
                            //        continue;
                            //    }

                            //    File.Copy(file, Path.Combine(tmFolder, fileName), true);
                            //}

                            break;
                        }
                    }

                    List<string> filesToCheck = new()
                    {
                        "EK-A.pdf",
                        "EK-C.pdf"
                    };

                    bool allExists = true;

                    foreach (var fileName in filesToCheck)
                    {
                        var file = Path.Combine(tmFolder, fileName);

                        if (!File.Exists(file))
                        {
                            //Console.WriteLine($"{file} do not exists!!");
                            allExists = false;
                        }
                    }

                    string folderName = Path.GetFileName(tmFolder);
                    string parentFolder = Path.GetDirectoryName(tmFolder);
                    string newFolderName = folderName;

                    // Remove any existing (OK), (OK?), or (OK!!) suffixes
                    if (folderName.Contains("(OK)"))
                    {
                        newFolderName = folderName.Replace("(OK)", "").Trim();
                    }
                    else if (folderName.Contains("(OK?)"))
                    {
                        newFolderName = folderName.Replace("(OK?)", "").Trim();
                    }
                    else if (folderName.Contains("(OK!!)"))
                    {
                        newFolderName = folderName.Replace("(OK!!)", "").Trim();
                    }
                    else if (folderName.Contains("(OK_CHECK)"))
                    {
                        newFolderName = folderName.Replace("(OK_CHECK)", "").Trim();
                    }
                    else continue;

                    // Add the appropriate suffix based on conditions
                    if (allExists && found)
                    {
                        newFolderName += " (OK)";
                        //Console.WriteLine($"Renaming folder to have (OK): {folderName} -> {newFolderName}");
                    }
                    else if (allExists && !found)
                    {
                        newFolderName += " (OK_CHECK)";
                        //Console.WriteLine($"Renaming folder to have (OK_CHECK): {folderName} -> {newFolderName}");
                    }
                    else if (!allExists && !found)
                    {
                        newFolderName += " (OK!!)";
                        //Console.WriteLine($"Renaming folder to have (OK!!): {folderName} -> {newFolderName}");
                    }

                    // Only rename if the folder name has changed
                    if (newFolderName != folderName)
                    {
                        string newPath = Path.Combine(parentFolder, newFolderName);
                        try
                        {
                            Directory.Move(tmFolder, newPath);
                            //Console.WriteLine($"Successfully renamed: {tmFolder} -> {newPath}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error renaming folder {tmFolder}: {ex.Message}");
                        }
                    }
                }
            }
        }

    }
}
