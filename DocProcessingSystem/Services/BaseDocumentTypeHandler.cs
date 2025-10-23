using DocProcessingSystem.Core;
using DocProcessingSystem.Models;

namespace DocProcessingSystem.Services
{
    #region Base Handler Class

    /// <summary>
    /// Base implementation for document type handlers
    /// </summary>
    public abstract class BaseDocumentTypeHandler : IDocumentTypeHandler
    {
        protected readonly IDocumentProcessor _processor;
        protected readonly IFolderMatcher _matcher;
        protected readonly string _documentType;


        /// <summary>
        /// Initializes a new instance of the base document handler
        /// </summary>
        protected BaseDocumentTypeHandler(IDocumentProcessor processor, IFolderMatcher matcher, string documentType)
        {
            _processor = processor;
            _matcher = matcher;
            _documentType = documentType;
        }

        /// <summary>
        /// Determines if this handler can process the specified document type
        /// </summary>
        public bool CanHandle(string documentType)
        {
            return string.Equals(documentType, _documentType, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Processes documents from source folder against analysis groups
        /// </summary>
        public virtual void ProcessDocuments(string sourceFolder, ref List<FolderGroup> analysisGroups, in IPdfMerger merger)
        {
            var toBeRemoved = new List<FolderGroup>();
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            Console.WriteLine($"Processing {_documentType} documents from {sourceFolder}");

            foreach (var group in analysisGroups)
            {
                //Console.WriteLine($"Processing group: TM {group.TmNo}, Building {group.BuildingCode}-{group.BuildingTmId}");

                // Find relevant documents for this group
                var relevantDocs = FindRelevantDocuments(sourceFolder, group);

                if (!relevantDocs.Any())
                {
                    Console.WriteLine($"No relevant {_documentType} documents found for this group");
                    continue;
                }

                //Console.WriteLine($"Found {relevantDocs.Count} relevant documents");

                if (relevantDocs.Count != 1) throw new Exception($"Found duplicate words TM No: {group.TmNo}, Building Code: {group.BuildingCode}, Building TM ID: {group.BuildingTmId}");

                // Process each document
                foreach (var docPath in relevantDocs)
                {
                    string outputPath = GetOutputPath(group, docPath);

                    try
                    {
                        //if (File.Exists(outputPath)) continue;
                        _processor.Convert(docPath, outputPath, false, false);
                        var sequence = GetMergeSequence(group);

                        Console.WriteLine($"Merging: {Path.GetFileNameWithoutExtension(sequence.OutputPath)}");
                        merger.MergePdf(
                            sequence.MainDocument,
                            sequence.AdditionalDocuments,
                            sequence.OutputPath,
                            sequence.Options
                            );

                        //if (this is Post2008DocumentHandler post2008Document)
                        //{
                        //    var ekBSquence = post2008Document.GetMergeSequenceForEkB(group);

                        //    Console.WriteLine($"Merging EK-B for extra");

                        //    merger.MergePdf(
                        //        ekBSquence.MainDocument,
                        //        ekBSquence.AdditionalDocuments,
                        //        ekBSquence.OutputPath,
                        //        ekBSquence.Options
                        //        );
                        //}

                        //var options = new PdfExtractionOptions // TODO: May need a few adjustments
                        //{
                        //    StartPageSelectionType = PageSelectionType.Keyword,
                        //    StartKeyword = new KeywordOptions
                        //    {
                        //        Keyword = "EK-B",
                        //        Occurrence = KeywordOccurrence.Last,
                        //        IncludeMatchingPage = true
                        //    },
                        //    EndPageSelectionType = PageSelectionType.LastPage
                        //};

                        //var service = new PdfRangeExtractorService();
                        //service.ExtractRange(
                        //    sequence.OutputPath,
                        //    Path.Combine(Path.GetDirectoryName(sequence.OutputPath), "EK-B.pdf"),
                        //    options
                        //    );


                        toBeRemoved.Add(group);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing document {Path.GetFileName(docPath)}: {ex.Message}");
                    }
                }

            }

            analysisGroups = analysisGroups
                .Except(toBeRemoved)
                .ToList();
        }

        /// <summary>
        /// Gets the merge sequence for a specific folder group
        /// </summary>
        public abstract MergeSequence GetMergeSequence(FolderGroup group);

        /// <summary>
        /// Finds relevant documents for a specific group
        /// </summary>
        protected abstract List<string> FindRelevantDocuments(string sourceFolder, FolderGroup group);

        /// <summary>
        /// Gets the output path for a converted document
        /// </summary>
        protected abstract string GetOutputPath(FolderGroup group, string documentPath);
    }

    #endregion

    #region Sealed Handlers

    /// <summary>
    /// Handler for post2008 documents
    /// </summary>
    public sealed class Post2008DocumentHandler : BaseDocumentTypeHandler
    {
        /// <summary>
        /// Initializes a new parametric document handler
        /// </summary>
        public Post2008DocumentHandler(IDocumentProcessor processor, IFolderMatcher matcher)
            : base(processor, matcher, "Post2008")
        {
        }

        private void CreateEkFiles(string sourceFolder, ref List<FolderGroup> analysisGroups, in IPdfMerger merger)
        {
            Console.WriteLine("Processing additional EK-B documents for Post2008 handler");

            // Process the remaining groups for EK-B documents
            foreach (var group in analysisGroups)
            {
                var ekBSequence = GetMergeSequenceForEkB(group);
                Console.WriteLine($"Merging EK-B for TM No: {group.TmNo}, Building Code: {group.BuildingCode}, Building TM ID: {group.BuildingTmId}");

                try
                {
                    merger.MergePdf(
                        ekBSequence.MainDocument,
                        ekBSequence.AdditionalDocuments,
                        ekBSequence.OutputPath,
                        ekBSequence.Options
                    );
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing EK-B document: {ex.Message}");
                }

                if (group.PathCount > 1)
                {
                    var ekCSequence = GetMergeSequenceForEkC(group);
                    Console.WriteLine($"Merging EK-C for TM No: {group.TmNo}, Building Code: {group.BuildingCode}, Building TM ID: {group.BuildingTmId}");

                    try
                    {
                        merger.MergePdf(
                            ekCSequence.MainDocument,
                            ekCSequence.AdditionalDocuments,
                            ekCSequence.OutputPath,
                            ekCSequence.Options
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing EK-C document: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Processes documents from source folder against analysis groups
        /// </summary>
        public override void ProcessDocuments(string sourceFolder, ref List<FolderGroup> analysisGroups, in IPdfMerger merger)
        {
            // Process the specific code first
            //CreateEkFiles(sourceFolder, ref analysisGroups, in merger);

            // Then call the base implementation
            base.ProcessDocuments(sourceFolder, ref analysisGroups, in merger);
        }

        /// <summary>
        /// Gets the merge sequence for parametric documents
        /// </summary>
        public override MergeSequence GetMergeSequence(FolderGroup folderGroup)
        {
            //CopyAdditionalFiles(folderGroup);

            //throw new NotImplementedException();
            string mainPdf = Path.Combine(folderGroup.MainFolder, "main.pdf");

            // Get additional PDFs
            var additionalPdfs = new List<string>();

            string projectRootPath = AppDomain.CurrentDomain.BaseDirectory;
            string post2008CoverPath = Path.Combine(projectRootPath, "CoverPages", "Post2008");

            if (folderGroup.MainFolder.Contains("M10"))
            {
                switch (folderGroup.PathCount) 
                {
                    case 1:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "TBDYResults", "EK-A.pdf"));
                        //additionalPdfs.Add(Path.Combine(post2008CoverPath, "EK-B.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "TBDYResults", "EK-B.pdf"));
                        break;
                    case 2:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "TBDYResults", "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "TBDYResults", "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "TBDYResults", "EK-C.pdf"));
                        //additionalPdfs.Add(Path.Combine(post2008CoverPath, "EK-B_A-Blok.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        //additionalPdfs.Add(Path.Combine(post2008CoverPath, "EK-B_B-Blok.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD2.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD2.pdf"));
                        break;
                    // TODO: Add cases for more blocks

                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }
            else
            {
                switch (folderGroup.PathCount)
                {
                    case 1:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "TBDYResults", "EK-A.pdf"));
                        //additionalPdfs.Add(Path.Combine(post2008CoverPath, "EK-B.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD1.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD1.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD3.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "TBDYResults", "EK-B.pdf"));
                        break;
                    case 2:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "TBDYResults", "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "TBDYResults", "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "TBDYResults", "EK-C.pdf"));
                        //additionalPdfs.Add(Path.Combine(post2008CoverPath, "EK-B_A-Blok.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD1.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD1.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD3.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD3.pdf"));
                        //additionalPdfs.Add(Path.Combine(post2008CoverPath, "EK-B_B-Blok.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD1.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD1.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD2.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD2.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD3.pdf"));
                        //additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD3.pdf"));
                        break;
                    // TODO: Add cases for more blocks

                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }

            var options = new MergeOptions
            {
                PreserveBookmarks = true,
                RequiredSections = new[] { "EK-A" }
            };

            var areaId = folderGroup.TmNo.Split("-")[0];
            var tmId = folderGroup.TmNo.Split("-")[1];

            return new MergeSequence
            {
                MainDocument = mainPdf,
                AdditionalDocuments = additionalPdfs,
                OutputPath = Path.Combine(folderGroup.MainFolder, "TBDYResults", $"TEI-B{areaId}-TM-{tmId}-DIR-M{folderGroup.BuildingCode}-{folderGroup.BuildingTmId}.pdf"),
                Options = options
            };
        }

        private string GetEkAFile(FolderGroup folderGroup)
        {
            var location = @"C:\Users\Mert\Desktop\Fırat Report Revision\MM_RAPOR\EK-A"; // TODO: FIRAT

            var possibleFolderNames = new List<string>()
            {
                @$"{folderGroup.TmNo}_M{folderGroup.BuildingCode}-{folderGroup.BuildingTmId}",
                @$"{folderGroup.TmNo}_M{folderGroup.BuildingCode}-{folderGroup.BuildingTmId}-A"
            };

            string ekFile = null;

            foreach (var name in possibleFolderNames)
            {
                var path = Path.Combine(location, name, "EK-A.pdf");

                if (File.Exists(path)) ekFile = path;
            }

            if (ekFile == null) throw new ArgumentNullException("Could not find EK-A file");

            string destinationPath = Path.Combine(folderGroup.MainFolder, "EK-A.pdf");
            File.Copy(ekFile, destinationPath, true);

            return ekFile;
        }

        /// <summary>
        /// Gets the merge sequence for additional EK-B document
        /// </summary>
        public MergeSequence GetMergeSequenceForEkB(FolderGroup folderGroup)
        {
            //throw new NotImplementedException();

            // Get additional PDFs
            var additionalPdfs = new List<string>();

            string projectRootPath = AppDomain.CurrentDomain.BaseDirectory;
            string post2008CoverPath = Path.Combine(projectRootPath, "CoverPages", "Post2008");
            string mainPdf = folderGroup.PathCount == 1 ? Path.Combine(post2008CoverPath, "EK-B.pdf")
                : Path.Combine(post2008CoverPath, "EK-B_A-Blok.pdf");

            if (folderGroup.MainFolder.Contains("M10"))
            {
                switch (folderGroup.PathCount)
                {
                    case 1:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        break;
                    case 2:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(post2008CoverPath, "EK-B_B-Blok.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD2.pdf"));
                        break;
                    // TODO: Add cases for more blocks

                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }
            else
            {
                switch (folderGroup.PathCount)
                {
                    case 1:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD3.pdf"));
                        break;
                    case 2:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(post2008CoverPath, "EK-B_B-Blok.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD3.pdf"));
                        break;
                    // TODO: Add cases for more blocks

                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }

            var options = new MergeOptions
            {
                PreserveBookmarks = true
            };

            var areaId = folderGroup.TmNo.Split("-")[0];
            var tmId = folderGroup.TmNo.Split("-")[1];

            var outputFolder = folderGroup.MainFolder.Replace("NİHAİ_TESLİM", "EK_B");

            if (!Path.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            return new MergeSequence
            {
                MainDocument = mainPdf,
                AdditionalDocuments = additionalPdfs,
                OutputPath = Path.Combine(outputFolder, "EK-B.pdf"),
                Options = options
            };
        }

        /// <summary>
        /// Gets the merge sequence for additional EK-C document
        /// </summary>
        public MergeSequence GetMergeSequenceForEkC(FolderGroup folderGroup)
        {
            //throw new NotImplementedException();

            // Get additional PDFs
            var additionalPdfs = new List<string>();

            string projectRootPath = AppDomain.CurrentDomain.BaseDirectory;
            string post2008CoverPath = Path.Combine(projectRootPath, "CoverPages", "Post2008");
            string mainPdf = Path.Combine(post2008CoverPath, "EK-B_B-Blok.pdf");

            if (folderGroup.MainFolder.Contains("M10"))
            {
                switch (folderGroup.PathCount)
                {
                    case 2:
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD2.pdf"));
                        break;
                    // TODO: Add cases for more blocks

                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }
            else
            {
                switch (folderGroup.PathCount)
                {
                    case 2:
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD3.pdf"));
                        break;
                    // TODO: Add cases for more blocks

                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }

            var options = new MergeOptions
            {
                PreserveBookmarks = true
            };

            var areaId = folderGroup.TmNo.Split("-")[0];
            var tmId = folderGroup.TmNo.Split("-")[1];

            var outputFolder = folderGroup.Paths[1].Replace("NİHAİ_TESLİM", "EK_B");

            if (!Path.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            return new MergeSequence
            {
                MainDocument = mainPdf,
                AdditionalDocuments = additionalPdfs,
                OutputPath = Path.Combine(outputFolder, "EK-C.pdf"),
                Options = options
            };
        }

        /// <summary>
        /// Finds relevant parametric documents for a specific group
        /// </summary>
        protected override List<string> FindRelevantDocuments(string sourceFolder, FolderGroup group)
        {
            var relevantDocs = new List<string>();

            // Get all Word documents in the source folder and subfolders
            var allDocs = Directory.GetFiles(sourceFolder, "*.docx", SearchOption.AllDirectories);

            foreach (var docPath in allDocs)
            {
                string folderName = Path.GetFileName(docPath);

                // Check if this document's folder matches the group
                if (_matcher.IsMatch(folderName, group))
                {
                    relevantDocs.Add(docPath);
                }
            }

            return relevantDocs;
        }

        /// <summary>
        /// Gets the output path for a parametric document
        /// </summary>
        protected override string GetOutputPath(FolderGroup group, string documentPath)
        {
            string mainFolder = group.MainFolder;
            string mainPdfPath = Path.Combine(mainFolder, "main.pdf");

            return mainPdfPath;
        }

        private void CopyAdditionalFiles(FolderGroup folderGroup)
        {
            var location = @"C:\Users\Mert\Desktop\MM_RAPOR\ALTLIK"; // TODO: FIRAT
            foreach (var group in folderGroup.Paths)
            {
                string folderName;
                // Extract the folder name from the group path
                string groupFolderName = Path.GetFileName(group);
                // Check if the group ends with "-{char}"
                if (groupFolderName.Length > 0 &&
                    groupFolderName.Contains("-") &&
                    groupFolderName.LastIndexOf("-") == groupFolderName.Length - 2)
                {
                    // If it ends with "-{char}", include the block in the folder name
                    var block = groupFolderName.Split("-").Last();
                    folderName = $"{folderGroup.TmNo}_M{folderGroup.BuildingCode}-{folderGroup.BuildingTmId}-{block}";
                }
                else
                {
                    // Otherwise, use the name without a block
                    folderName = $"{folderGroup.TmNo}_M{folderGroup.BuildingCode}-{folderGroup.BuildingTmId}";
                }
                var contentsToCopy = Path.Combine(location, folderName);
                // Check if source directory exists
                if (Directory.Exists(contentsToCopy))
                {
                    // Create destination directory if it doesn't exist
                    if (!Directory.Exists(group))
                    {
                        Directory.CreateDirectory(group);
                    }

                    // Copy all files from the root of the source directory
                    CopyFilesInDirectory(contentsToCopy, group);

                    // Copy all subdirectories and their contents
                    foreach (string dirPath in Directory.GetDirectories(contentsToCopy, "*", SearchOption.AllDirectories))
                    {
                        string dirName = dirPath.Substring(contentsToCopy.Length + 1);
                        string destDirPath = Path.Combine(group, dirName);

                        // Create the destination subdirectory
                        if (!Directory.Exists(destDirPath))
                        {
                            Directory.CreateDirectory(destDirPath);
                        }

                        // Copy files from the current subdirectory
                        CopyFilesInDirectory(dirPath, destDirPath);
                    }

                    Console.WriteLine($"Copied all files and folders to {group}");
                }
                else
                {
                    Console.WriteLine($"Source directory not found: {contentsToCopy}");
                }
            }
        }

        private void CopyFilesInDirectory(string sourceDir, string destDir)
        {
            int newFiles = 0;
            int overwrittenFiles = 0;

            // Get all files from the source directory
            string[] filesToCopy = Directory.GetFiles(sourceDir);

            // Copy each file to the destination
            foreach (string file in filesToCopy)
            {
                string fileName = Path.GetFileName(file);
                string destFile = Path.Combine(destDir, fileName);
                bool fileExists = File.Exists(destFile);

                // Copy the file, overwrite if exists
                File.Copy(file, destFile, true);

                if (fileExists)
                {
                    overwrittenFiles++;
                }
                else
                {
                    newFiles++;
                }
            }

            if (filesToCopy.Length > 0)
            {
                if (overwrittenFiles > 0)
                {
                    Console.WriteLine($"Copied {newFiles} new files and overwrote {overwrittenFiles} existing files in {destDir}");
                }
                else
                {
                    Console.WriteLine($"Copied {newFiles} files to {destDir}");
                }
            }
        }
    }

    /// <summary>
    /// Handler for parametric documents
    /// </summary>
    public sealed class ParametricDocumentHandler : BaseDocumentTypeHandler
    {
        /// <summary>
        /// Initializes a new parametric document handler
        /// </summary>
        public ParametricDocumentHandler(IDocumentProcessor processor, IFolderMatcher matcher)
            : base(processor, matcher, "Parametric")
        {
        }

        /// <summary>
        /// Finds relevant parametric documents for a specific group
        /// </summary>
        protected override List<string> FindRelevantDocuments(string sourceFolder, FolderGroup group)
        {
            var relevantDocs = new List<string>();

            // Get all Word documents in the source folder and subfolders
            var allDocs = Directory.GetFiles(sourceFolder, "*.docx", SearchOption.AllDirectories);

            foreach (var docPath in allDocs)
            {
                string folderName = Path.GetFileName(docPath);

                // Check if this document's folder matches the group
                if (_matcher.IsMatch(folderName, group))
                {
                    relevantDocs.Add(docPath);
                }
            }

            return relevantDocs;
        }

        /// <summary>
        /// Gets the output path for a parametric document
        /// </summary>
        protected override string GetOutputPath(FolderGroup group, string documentPath)
        {
            string mainFolder = group.MainFolder;
            string mainPdfPath = Path.Combine(mainFolder, "main.pdf");

            return mainPdfPath;
        }

        /// <summary>
        /// Gets the merge sequence for parametric documents
        /// </summary>
        public override MergeSequence GetMergeSequence(FolderGroup folderGroup)
        {
            string mainPdf = Path.Combine(folderGroup.MainFolder, "main.pdf");

            // Get additional PDFs
            var additionalPdfs = new List<string>();

            string projectRootPath = AppDomain.CurrentDomain.BaseDirectory;
            string parametricCoverPath = Path.Combine(projectRootPath, "CoverPages", "Parametric");

            if (folderGroup.MainFolder.Contains("M10"))
            {
                switch (folderGroup.PathCount)
                {
                    case 1:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "OneBlock", "KAPAK_EK-C.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_KAROT.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "OneBlock", "KAPAK_EK-D.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_DONATI.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "OneBlock", "KAPAK_EK-E.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        break;
                    case 2:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-C.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "TwoBlocks", "KAPAK_EK-D.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_KAROT.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "TwoBlocks", "KAPAK_EK-E.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_DONATI.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "TwoBlocks", "KAPAK_EK-F (A Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "TwoBlocks", "KAPAK_EK-G (B Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD2.pdf"));
                        break;
                    case 3:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-C.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-D.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "ThreeBlocks", "KAPAK_EK-E.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_KAROT.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "ThreeBlocks", "KAPAK_EK-F.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_DONATI.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "ThreeBlocks", "KAPAK_EK-G (A Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "ThreeBlocks", "KAPAK_EK-H (B Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "ThreeBlocks", "KAPAK_EK-I (C Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[2], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[2], "2_Tespit_DD2.pdf"));
                        break;
                    // TODO: Add cases for more blocks

                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }
            else
            {
                switch (folderGroup.PathCount)
                {
                    case 1:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "OneBlock", "KAPAK_EK-C.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_KAROT.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "OneBlock", "KAPAK_EK-D.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_DONATI.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "OneBlock", "KAPAK_EK-E.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD3.pdf"));
                        break;
                    case 2:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-C.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "TwoBlocks", "KAPAK_EK-D.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_KAROT.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "TwoBlocks", "KAPAK_EK-E.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_DONATI.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "TwoBlocks", "KAPAK_EK-F (A Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "TwoBlocks", "KAPAK_EK-G (B Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD3.pdf"));
                        break;
                    case 3:
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-C.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK-D.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "ThreeBlocks", "KAPAK_EK-E.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_KAROT.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "ThreeBlocks", "KAPAK_EK-F.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "EK_DONATI.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "ThreeBlocks", "KAPAK_EK-G (A Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.MainFolder, "2_Tespit_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "ThreeBlocks", "KAPAK_EK-H (B Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[1], "2_Tespit_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(parametricCoverPath, "ThreeBlocks", "KAPAK_EK-I (C Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[2], "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[2], "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[2], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[2], "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[2], "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(folderGroup.Paths[2], "2_Tespit_DD3.pdf"));
                        break;
                    // TODO: Add cases for more blocks

                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }

            var options = new MergeOptions
            {
                PreserveBookmarks = true,
                RequiredSections = new[] { "EK-A", "EK_KAROT", "EK_DONATI" }
            };

            var areaId = folderGroup.TmNo.Split("-")[0];
            var tmId = folderGroup.TmNo.Split("-")[1];

            return new MergeSequence
            {
                MainDocument = mainPdf,
                AdditionalDocuments = additionalPdfs,
                OutputPath = Path.Combine(folderGroup.MainFolder, $"TEI-B{areaId}-TM-{tmId}-DIR-M{folderGroup.BuildingCode}-{folderGroup.BuildingTmId}_nt.pdf"),
                Options = options
            };
        }
    }

    /// <summary>
    /// Handler for deterministic documents
    /// </summary>
    public sealed class DeterministicDocumentHandler : BaseDocumentTypeHandler
    {
        /// <summary>
        /// Initializes a new deterministic document handler
        /// </summary>
        public DeterministicDocumentHandler(IDocumentProcessor processor, IFolderMatcher matcher)
            : base(processor, matcher, "Deterministic")
        {
        }

        /// <summary>
        /// Finds relevant deterministic documents for a specific group
        /// </summary>
        protected override List<string> FindRelevantDocuments(string sourceFolder, FolderGroup group)
        {
            var relevantDocs = new List<string>();

            // Get all Word documents in the source folder and subfolders
            var allDocs = Directory.GetFiles(sourceFolder, "*.docx", SearchOption.AllDirectories);

            foreach (var docPath in allDocs)
            {
                string folderName = Path.GetFileName(docPath);

                // Check if this document's folder matches the group
                if (_matcher.IsMatch(folderName, group))
                {
                    relevantDocs.Add(docPath);
                }
            }

            return relevantDocs;
        }

        /// <summary>
        /// Gets the output path for a deterministic document
        /// </summary>
        protected override string GetOutputPath(FolderGroup group, string documentPath)
        {
            string mainFolder = group.MainFolder;
            string mainPdfPath = Path.Combine(mainFolder, "main.pdf");

            return mainPdfPath;
        }

        /// <summary>
        /// Gets the merge sequence for deterministic documents
        /// </summary>
        public override MergeSequence GetMergeSequence(FolderGroup group)
        {
            string mainPdf = Path.Combine(group.MainFolder, "main.pdf");


            // Get additional PDFs
            var additionalPdfs = new List<string>();

            var requiredPdfs = new { };

            string projectRootPath = AppDomain.CurrentDomain.BaseDirectory;
            string deterministicCoverPath = Path.Combine(projectRootPath, "CoverPages", "Deterministic");

            if (group.MainFolder.Contains("M10"))
            {
                switch (group.PathCount)
                {
                    case 1:
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-C.pdf"));
                        additionalPdfs.Add(Path.Combine(deterministicCoverPath, "KAPAK_EK-D (Single Block).pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD2.pdf"));
                        break;
                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }
            else
            {
                switch (group.PathCount)
                {
                    case 1:
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-C.pdf"));
                        additionalPdfs.Add(Path.Combine(deterministicCoverPath, "KAPAK_EK-D (Single Block).pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD3.pdf"));
                        break;
                    case 2:
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-C.pdf"));
                        additionalPdfs.Add(Path.Combine(deterministicCoverPath, "KAPAK_EK-D (A Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(deterministicCoverPath, "KAPAK_EK-E (B Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "2_Tespit_DD3.pdf"));
                        break;
                    case 3:
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-A.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-B.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "EK-C.pdf"));
                        additionalPdfs.Add(Path.Combine(deterministicCoverPath, "KAPAK_EK-D (A Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(deterministicCoverPath, "KAPAK_EK-E (B Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "2_Tespit_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(deterministicCoverPath, "KAPAK_EK-F (C Blok).pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[2], "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[2], "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[2], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[2], "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[2], "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[2], "2_Tespit_DD3.pdf"));
                        break;
                    // TODO: Add cases

                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }
            var options = new MergeOptions
            {
                PreserveBookmarks = true,
                RequiredSections = new[] { "EK-A", "Kapak_DD2", "2_Tespit_DD2" }
            };

            var areaId = group.TmNo.Split("-")[0];
            var tmId = group.TmNo.Split("-")[1];

            return new MergeSequence
            {
                MainDocument = mainPdf,
                AdditionalDocuments = additionalPdfs,
                OutputPath = Path.Combine(group.MainFolder, $"TEI-B{areaId}-TM-{tmId}-DIR-M{group.BuildingCode}-{group.BuildingTmId}_nt.pdf"),
                Options = options
            };
        }
    }

    /// <summary>
    /// Handler for masonry documents
    /// </summary>
    public sealed class MasonryDocumentHandler : BaseDocumentTypeHandler
    {
        public MasonryDocumentHandler(IDocumentProcessor processor, IFolderMatcher matcher)
            : base(processor, matcher, "Deterministic")
        {
        }

        public override void ProcessDocuments(string sourceFolder, ref List<FolderGroup> analysisGroups, in IPdfMerger merger)
        {
            foreach (var group in analysisGroups)
            {
                //Console.WriteLine($"Processing group: TM {group.TmNo}, Building {group.BuildingCode}-{group.BuildingTmId}");

                try
                {
                    var sequence = GetMergeSequence(group);

                    Console.WriteLine($"Merging: {Path.GetFileNameWithoutExtension(sequence.OutputPath)}");
                    merger.MergePdf(
                        sequence.MainDocument,
                        sequence.AdditionalDocuments,
                        sequence.OutputPath,
                        sequence.Options
                        );

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing document {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Gets the merge sequence for deterministic documents
        /// </summary>
        public override MergeSequence GetMergeSequence(FolderGroup group)
        {
            var additionalPdfs = new List<string>();
            var requiredPdfs = new { };

            string projectRootPath = AppDomain.CurrentDomain.BaseDirectory;
            string masonryCoverPath = Path.Combine(projectRootPath, "CoverPages", "Masonry");
            string mainPdf = group.PathCount == 1 ? Path.Combine(masonryCoverPath, "EK-C_Kapak.pdf")
                : Path.Combine(masonryCoverPath, "EK-C_Kapak-A_Blok.pdf");

            if (group.MainFolder.Contains("M10"))
            {
                switch (group.PathCount)
                {
                    case 1:
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD2.pdf"));
                        break;
                    case 2:
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(masonryCoverPath, "EK-C_Kapak-B_Blok.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.Paths[1], "2_Tespit_DD2.pdf"));
                        break;
                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }
            else
            {
                switch (group.PathCount)
                {
                    case 1:
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD3.pdf"));
                        break;
                    case 2:
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(masonryCoverPath, "EK-C_Kapak-B_Blok.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD1.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD2.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "Kapak_DD3.pdf"));
                        additionalPdfs.Add(Path.Combine(group.MainFolder, "2_Tespit_DD3.pdf"));
                        break;
                    // TODO: Add cases

                    default:
                        throw new Exception("Unhandled number of blocks!!");
                }
            }
            var options = new MergeOptions
            {
                PreserveBookmarks = true,
                RequiredSections = new[] { "Kapak_DD2", "2_Tespit_DD2" }
            };

            var areaId = group.TmNo.Split("-")[0];
            var tmId = group.TmNo.Split("-")[1];

            return new MergeSequence
            {
                MainDocument = mainPdf,
                AdditionalDocuments = additionalPdfs,
                OutputPath = GetOutputPath(group, ""),
                Options = options
            };
        }

        protected override List<string> FindRelevantDocuments(string sourceFolder, FolderGroup group)
        {
            return new List<string>();
        }

        protected override string GetOutputPath(FolderGroup group, string documentPath)
        {
            return Path.Combine(group.MainFolder, "EK-C.pdf");
        }
    }

    #endregion

    #region Folder Name Matcher

    /// <summary>
    /// Matches folders based on name pattern
    /// </summary>
    public class FolderNameMatcher : IFolderMatcher
    {
        /// <summary>
        /// Determines if a folder matches with an analysis group
        /// </summary>
        public bool IsMatch(string fileName, FolderGroup group)
        {
            var (tmNo, buildingCode, buildingTmId) = FolderHelper.ExtractParts(fileName);

            if (string.IsNullOrEmpty(tmNo) || string.IsNullOrEmpty(buildingCode))
                return false;

            // Match by TM number, building code, and building TM ID
            return string.Equals(tmNo, group.TmNo, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(buildingCode, group.BuildingCode, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(buildingTmId, group.BuildingTmId, StringComparison.OrdinalIgnoreCase);
        }
    }

    #endregion
}