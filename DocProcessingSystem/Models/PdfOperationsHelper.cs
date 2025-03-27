using DocProcessingSystem.Core;
using DocProcessingSystem.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessingSystem.Models
{
    /// <summary>
    /// Differentiate folder types 
    /// </summary>
    enum FolderType
    {
        SPLIT_EK_D,
        SPLIT_EK_B,
        MERGE_THREE_ATTACHMENTS,
        MERGE_ONE_ATTACHMENT
    }

    public class PdfOperationsHelper
    {
        #region Public Methods

        /// <summary>
        /// Processes PDF documents based on their location in the predefined folder structure
        /// </summary>
        public static void ProcessPdfDocuments()
        {
            // Root folder where all operations will take place
            var rootFolder = @"C:\Users\Mert\Desktop\total";              // TODO: FIRAT Create a main folder for all documents

            // Define folder paths using the existing structure
            var splitPdfRootFolder = Path.Combine(rootFolder, "MM");  // TODO: FIRAT For words that should be seperated from keyword "EK-X"
            var mergePdfRootFolder = Path.Combine(rootFolder, "asd"); // TODO: FIRAT For words that should be just merged with additional pdfs

            var complexAttachmentFolder = "specific";                 // TODO: FIRAT For each folder above create a new folder to hold folders which ends with "--"
            var simpleAttachmentFolder = "nonspecific";               // TODO: FIRAT For each folder above create a new folder to hold folders which does not end with "--"

            // First convert Word documents to PDF
            ConvertWordToPdfAsync(rootFolder, rootFolder, false, true, 8).Wait();

            // Dictionary mapping processing types to their folder paths
            var processingFolderPaths = new Dictionary<FolderType, string>
            {
                {FolderType.SPLIT_EK_D, Path.Combine(splitPdfRootFolder, complexAttachmentFolder)},
                {FolderType.SPLIT_EK_B, Path.Combine(splitPdfRootFolder, simpleAttachmentFolder)},
                {FolderType.MERGE_THREE_ATTACHMENTS, Path.Combine(mergePdfRootFolder, complexAttachmentFolder)},
                {FolderType.MERGE_ONE_ATTACHMENT, Path.Combine(mergePdfRootFolder, simpleAttachmentFolder)}
            };

            // Process PDFs according to their type
            using (var pdfMerger = new PdfMergerService())
            {
                var pdfExtractor = new PdfRangeExtractorService();

                // Process all PDF types
                ProcessSplitEkD(processingFolderPaths[FolderType.SPLIT_EK_D], pdfExtractor, pdfMerger);
                ProcessSplitEkB(processingFolderPaths[FolderType.SPLIT_EK_B], pdfExtractor, pdfMerger);
                ProcessWithThreeAttachments(processingFolderPaths[FolderType.MERGE_THREE_ATTACHMENTS], pdfMerger);
                ProcessWithOneAttachment(processingFolderPaths[FolderType.MERGE_ONE_ATTACHMENT], pdfMerger);
            }
        }

        public static void ConvertWordToPdf(string inputFolderPath, string outputFolderPath, bool saveChanges, bool useRelativePath = false)
        {
            var wordFiles = Directory.GetFiles(inputFolderPath, "*.docx", SearchOption.AllDirectories);

            using (var converter = new WordToPdfConverter())
            {
                foreach (string file in wordFiles)
                {
                    var baseName = Path.GetFileNameWithoutExtension(file);

                    var outputPath = useRelativePath ? file.Replace(".docx", ".pdf") : Path.Combine(outputFolderPath, baseName + ".pdf");

                    converter.Convert(file, outputPath, saveChanges, false);
                }
            }
        }

        public static async Task ConvertWordToPdfAsync(string inputFolderPath, string outputFolderPath, bool saveChanges, bool useRelativePath = false, int maxParallel = 5)
        {
            // Get all Word files asynchronously
            var wordFiles = await DirectoryExtensions.GetFilesAsync(inputFolderPath, "*.docx", SearchOption.AllDirectories);

            // Create progress reporting
            var progress = new Progress<(string FileName, int Completed, int Total)>(update =>
            {
                Console.WriteLine($"Converted {update.FileName} - {update.Completed} of {update.Total} completed");
            });

            // If using relative paths, we need to handle each file individually
            if (useRelativePath)
            {
                Console.WriteLine($"Starting conversion of {wordFiles.Length} files with relative paths");

                // Create a semaphore to limit concurrency
                using var semaphore = new SemaphoreSlim(maxParallel, maxParallel);
                using var cts = new CancellationTokenSource();

                // Create a list of tasks
                var tasks = new List<Task>();
                int completedCount = 0;

                foreach (var wordFile in wordFiles)
                {
                    // Wait for a slot to be available
                    await semaphore.WaitAsync(cts.Token);

                    // Create a task for each file
                    var task = Task.Run(async () =>
                    {
                        try
                        {
                            // Generate output path in the same directory as the input file
                            string outputPath = wordFile.Replace(".docx", ".pdf");

                            // Create a single-file converter for this task
                            using (var converter = new WordToPdfConverter())
                            {
                                converter.Convert(wordFile, outputPath, saveChanges, false);
                            }

                            // Report progress
                            int currentCount = Interlocked.Increment(ref completedCount);
                            ((IProgress<(string, int, int)>)progress).Report((Path.GetFileName(wordFile), currentCount, wordFiles.Length));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error converting {wordFile}: {ex.Message}");
                        }
                        finally
                        {
                            // Release the semaphore slot
                            semaphore.Release();
                        }
                    }, cts.Token);

                    tasks.Add(task);
                }

                try
                {
                    // Wait for all tasks to complete
                    await Task.WhenAll(tasks);
                    Console.WriteLine("All files converted successfully");
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine("Operation was cancelled");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during conversion: {ex.Message}");
                }
            }
            else
            {
                // When not using relative paths, use the original implementation
                using (var cts = new CancellationTokenSource())
                using (var converter = new ParallelWordToPdfConverter(maxParallel))
                {
                    Console.WriteLine($"Starting conversion of {wordFiles.Length} files with {maxParallel} parallel workers");

                    try
                    {
                        // Convert all files in parallel with controlled concurrency
                        await converter.ConvertMultipleAsync(
                            wordFiles,
                            outputFolderPath,
                            saveChanges,
                            progress,
                            cts.Token);

                        Console.WriteLine("All files converted successfully");
                    }
                    catch (OperationCanceledException)
                    {
                        Console.WriteLine("Operation was cancelled");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error during conversion: {ex.Message}");
                    }
                }
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Processes PDFs by splitting at "EK-D" and merging with attachments A through D
        /// </summary>
        static void ProcessSplitEkD(string folderPath, PdfRangeExtractorService extractor, PdfMergerService merger)
        {
            // Find all TEI PDF files in the directory
            var pdfFiles = Directory.GetFiles(folderPath, "TEI*.pdf", SearchOption.AllDirectories);

            // Options for extracting the main document (before EK-D)
            var mainDocumentOptions = new PdfExtractionOptions
            {
                StartPageSelectionType = PageSelectionType.FirstPage,
                EndPageSelectionType = PageSelectionType.Keyword,
                EndKeyword = new KeywordOptions
                {
                    Keyword = "EK-D",
                    Occurrence = KeywordOccurrence.Last,
                    IncludeMatchingPage = false,
                }
            };

            // Options for extracting the EK-D attachment
            var ekDAttachmentOptions = new PdfExtractionOptions
            {
                StartPageSelectionType = PageSelectionType.Keyword,
                StartKeyword = new KeywordOptions
                {
                    Keyword = "EK-D",
                    Occurrence = KeywordOccurrence.Last,
                    IncludeMatchingPage = true,
                },
                EndPageSelectionType = PageSelectionType.LastPage,
            };

            // Merge options preserving bookmarks
            var mergeOptions = new MergeOptions
            {
                PreserveBookmarks = true
            };

            foreach (var pdfFile in pdfFiles)
            {
                try
                {
                    string fileDirectory = Path.GetDirectoryName(pdfFile);
                    string fileNameNoExt = Path.GetFileNameWithoutExtension(pdfFile);

                    // Extract main document (before EK-D)
                    string mainOutputPath = Path.Combine(fileDirectory, "main.pdf");
                    extractor.ExtractRange(pdfFile, mainOutputPath, mainDocumentOptions);

                    // Extract EK-D attachment
                    string ekDOutputPath = Path.Combine(fileDirectory, "EK-D.pdf");
                    extractor.ExtractRange(pdfFile, ekDOutputPath, ekDAttachmentOptions);

                    // List of attachments to merge
                    var attachments = new List<string>
                {
                    Path.Combine(fileDirectory, "EK-A.pdf"),
                    Path.Combine(fileDirectory, "EK-B.pdf"),
                    Path.Combine(fileDirectory, "EK-C.pdf"),
                    Path.Combine(fileDirectory, "EK-D.pdf")
                };

                    // Merge the main document with attachments
                    var mergeSequence = new MergeSequence
                    {
                        MainDocument = mainOutputPath,
                        AdditionalDocuments = attachments,
                        OutputPath = pdfFile,
                        Options = mergeOptions
                    };

                    merger.MergePdf(mergeSequence);

                    Console.WriteLine($"Processed {Path.GetFileName(pdfFile)} with EK-D splitting");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: Cannot process file: {pdfFile}, Exception: {ex}");
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Processes PDFs by splitting at "EK-B" and merging with attachments A and B
        /// </summary>
        static void ProcessSplitEkB(string folderPath, PdfRangeExtractorService extractor, PdfMergerService merger)
        {
            // Find all TEI PDF files in the directory
            var pdfFiles = Directory.GetFiles(folderPath, "TEI*.pdf", SearchOption.AllDirectories);

            // Options for extracting the main document (before EK-B)
            var mainDocumentOptions = new PdfExtractionOptions
            {
                StartPageSelectionType = PageSelectionType.FirstPage,
                EndPageSelectionType = PageSelectionType.Keyword,
                EndKeyword = new KeywordOptions
                {
                    Keyword = "EK-B",
                    Occurrence = KeywordOccurrence.Last,
                    IncludeMatchingPage = false,
                }
            };

            // Options for extracting the EK-B attachment
            var ekBAttachmentOptions = new PdfExtractionOptions
            {
                StartPageSelectionType = PageSelectionType.Keyword,
                StartKeyword = new KeywordOptions
                {
                    Keyword = "EK-B",
                    Occurrence = KeywordOccurrence.Last,
                    IncludeMatchingPage = true,
                },
                EndPageSelectionType = PageSelectionType.LastPage,
            };

            // Merge options preserving bookmarks
            var mergeOptions = new MergeOptions
            {
                PreserveBookmarks = true
            };

            foreach (var pdfFile in pdfFiles)
            {
                try
                {
                    string fileDirectory = Path.GetDirectoryName(pdfFile);
                    string fileNameNoExt = Path.GetFileNameWithoutExtension(pdfFile);

                    // Extract main document (before EK-B)
                    string mainOutputPath = Path.Combine(fileDirectory, "main.pdf");
                    extractor.ExtractRange(pdfFile, mainOutputPath, mainDocumentOptions);

                    // Extract EK-B attachment
                    string ekBOutputPath = Path.Combine(fileDirectory, "EK-B.pdf");
                    extractor.ExtractRange(pdfFile, ekBOutputPath, ekBAttachmentOptions);

                    // List of attachments to merge
                    var attachments = new List<string>
                {
                    Path.Combine(fileDirectory, "EK-A.pdf"),
                    Path.Combine(fileDirectory, "EK-B.pdf")
                };

                    // Merge the main document with attachments
                    var mergeSequence = new MergeSequence
                    {
                        MainDocument = mainOutputPath,
                        AdditionalDocuments = attachments,
                        OutputPath = pdfFile,
                        Options = mergeOptions
                    };

                    merger.MergePdf(mergeSequence);

                    Console.WriteLine($"Processed {Path.GetFileName(pdfFile)} with EK-B splitting");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: Cannot process file: {pdfFile}, Exception: {ex}");
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Processes PDFs by directly merging with three attachments (EK-A, EK-B, EK-C)
        /// </summary>
        static void ProcessWithThreeAttachments(string folderPath, PdfMergerService merger)
        {
            // Find all TEI PDF files in the directory
            var pdfFiles = Directory.GetFiles(folderPath, "TEI*.pdf", SearchOption.AllDirectories);

            // Merge options preserving bookmarks
            var mergeOptions = new MergeOptions
            {
                PreserveBookmarks = true
            };

            foreach (var pdfFile in pdfFiles)
            {
                try
                {
                    string fileDirectory = Path.GetDirectoryName(pdfFile);

                    // List of attachments to merge
                    var attachments = new List<string>
                {
                    Path.Combine(fileDirectory, "EK-A.pdf"),
                    Path.Combine(fileDirectory, "EK-B.pdf"),
                    Path.Combine(fileDirectory, "EK-C.pdf")
                };

                    // Merge the main document with attachments
                    var mergeSequence = new MergeSequence
                    {
                        MainDocument = pdfFile,
                        AdditionalDocuments = attachments,
                        OutputPath = pdfFile,
                        Options = mergeOptions
                    };

                    merger.MergePdf(mergeSequence);

                    Console.WriteLine($"Processed {Path.GetFileName(pdfFile)} with three attachments");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: Cannot process file: {pdfFile}, Exception: {ex}");
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Processes PDFs by directly merging with one attachment (EK-A)
        /// </summary>
        static void ProcessWithOneAttachment(string folderPath, PdfMergerService merger)
        {
            // Find all TEI PDF files in the directory
            var pdfFiles = Directory.GetFiles(folderPath, "TEI*.pdf", SearchOption.AllDirectories);

            // Merge options preserving bookmarks
            var mergeOptions = new MergeOptions
            {
                PreserveBookmarks = true
            };

            foreach (var pdfFile in pdfFiles)
            {
                try
                {
                    string fileDirectory = Path.GetDirectoryName(pdfFile);

                    // List of attachments to merge
                    var attachments = new List<string>
                {
                    Path.Combine(fileDirectory, "EK-A.pdf")
                };

                    // Merge the main document with attachments
                    var mergeSequence = new MergeSequence
                    {
                        MainDocument = pdfFile,
                        AdditionalDocuments = attachments,
                        OutputPath = pdfFile,
                        Options = mergeOptions // Fix: was using thirdOptionMergeOption
                    };

                    merger.MergePdf(mergeSequence);

                    Console.WriteLine($"Processed {Path.GetFileName(pdfFile)} with one attachment");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: Cannot process file: {pdfFile}, Exception: {ex}");
                }
                Console.WriteLine();
            }
        }

        [Obsolete("Use ProcessPdfDocuments() for clear format")]
        public static void ProcessFiles()
        {
            var mainFolder = @"C:\Users\Mert\Desktop\total";
            ConvertWordToPdf(mainFolder, mainFolder, false, true);

            var deconstructorFolderName = "MM";
            var nonDeconstructorFolderName = "asd";

            var specificCaseFolderName = "specific";
            var nonSpecificFolderName = "nonspecific";

            Dictionary<string, string> dic = new()
            {
                {"SPLIT-EK-D", Path.Combine(mainFolder, deconstructorFolderName, specificCaseFolderName) },
                {"SPLIT-EK-B", Path.Combine(mainFolder, deconstructorFolderName, nonSpecificFolderName) },
                {"DON'T_SPLIT-3-PDF", Path.Combine(mainFolder, nonDeconstructorFolderName, specificCaseFolderName) },
                {"DON'T_SPLIT-1-PDF", Path.Combine(mainFolder, nonDeconstructorFolderName, nonSpecificFolderName) },
            };

            using (var merger = new PdfMergerService())
            {
                var service = new PdfRangeExtractorService();

                #region First Option
                var pdfFilesForFirstOption = Directory.GetFiles(dic["SPLIT-EK-D"], "TEI*.pdf", SearchOption.AllDirectories);
                var pdfFilesForSecondOption = Directory.GetFiles(dic["SPLIT-EK-B"], "TEI*.pdf", SearchOption.AllDirectories);
                var pdfFilesForThirdOption = Directory.GetFiles(dic["DON'T_SPLIT-3-PDF"], "TEI*.pdf", SearchOption.AllDirectories);
                var pdfFilesForFourthOption = Directory.GetFiles(dic["DON'T_SPLIT-1-PDF"], "TEI*.pdf", SearchOption.AllDirectories);

                var FirstOptionForMain = new PdfExtractionOptions
                {
                    StartPageSelectionType = PageSelectionType.FirstPage,
                    EndPageSelectionType = PageSelectionType.Keyword,
                    EndKeyword = new KeywordOptions
                    {
                        Keyword = "EK-D",
                        Occurrence = KeywordOccurrence.Last,
                        IncludeMatchingPage = false,
                    }
                };

                var FirstOptionForRemains = new PdfExtractionOptions
                {
                    StartPageSelectionType = PageSelectionType.Keyword,
                    StartKeyword = new KeywordOptions
                    {
                        Keyword = "EK-D",
                        Occurrence = KeywordOccurrence.Last,
                        IncludeMatchingPage = true,
                    },
                    EndPageSelectionType = PageSelectionType.LastPage,
                };

                var firstOptionMergeOption = new MergeOptions
                {
                    PreserveBookmarks = true
                };

                foreach (var file in pdfFilesForFirstOption)
                {
                    service.ExtractRange(
                        file, // TODO: Change Input Pdf Location 
                        file.Replace(Path.GetFileNameWithoutExtension(file), "main"),
                        FirstOptionForMain
                    );

                    service.ExtractRange(
                        file, // TODO: Change Input Pdf Location 
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-D"),
                        FirstOptionForRemains
                    );

                    var additionalPdfs = new List<string>()
                    {
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-A"),
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-B"),
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-C"),
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-D"),
                    };

                    var firstOptionMergeSequence = new MergeSequence
                    {
                        MainDocument = file.Replace(Path.GetFileNameWithoutExtension(file), "main"),
                        AdditionalDocuments = additionalPdfs,
                        OutputPath = file,
                        Options = firstOptionMergeOption
                    };

                    merger.MergePdf(firstOptionMergeSequence);

                    Console.WriteLine();
                }

                #endregion

                #region Second Option


                var secondOptionForMain = new PdfExtractionOptions
                {
                    StartPageSelectionType = PageSelectionType.FirstPage,
                    EndPageSelectionType = PageSelectionType.Keyword,
                    EndKeyword = new KeywordOptions
                    {
                        Keyword = "EK-B",
                        Occurrence = KeywordOccurrence.Last,
                        IncludeMatchingPage = false,
                    }
                };

                var secondOptionForRemains = new PdfExtractionOptions
                {
                    StartPageSelectionType = PageSelectionType.Keyword,
                    StartKeyword = new KeywordOptions
                    {
                        Keyword = "EK-B",
                        Occurrence = KeywordOccurrence.Last,
                        IncludeMatchingPage = true,
                    },
                    EndPageSelectionType = PageSelectionType.LastPage,
                };

                var secondOptionMergeOption = new MergeOptions
                {
                    PreserveBookmarks = true
                };

                foreach (var file in pdfFilesForSecondOption)
                {
                    service.ExtractRange(
                        file, // TODO: Change Input Pdf Location 
                        file.Replace(Path.GetFileNameWithoutExtension(file), "main"),
                        secondOptionForMain
                    );

                    service.ExtractRange(
                        file, // TODO: Change Input Pdf Location 
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-B"),
                        secondOptionForRemains
                    );

                    var additionalPdfs = new List<string>()
                    {
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-A"),
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-B"),
                    };

                    var secondOptionMergeSequence = new MergeSequence
                    {
                        MainDocument = file.Replace(Path.GetFileNameWithoutExtension(file), "main"),
                        AdditionalDocuments = additionalPdfs,
                        OutputPath = file,
                        Options = secondOptionMergeOption
                    };

                    merger.MergePdf(secondOptionMergeSequence);

                    Console.WriteLine();
                }

                #endregion

                #region Third Option

                var thirdOptionMergeOption = new MergeOptions
                {
                    PreserveBookmarks = true
                };

                foreach (var file in pdfFilesForThirdOption)
                {

                    var additionalPdfs = new List<string>()
                    {
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-A"),
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-B"),
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-C"),
                    };

                    var thirdOptionMergeSequence = new MergeSequence
                    {
                        MainDocument = file,
                        AdditionalDocuments = additionalPdfs,
                        OutputPath = file,
                        Options = thirdOptionMergeOption
                    };

                    merger.MergePdf(thirdOptionMergeSequence);

                    Console.WriteLine();
                }

                #endregion

                #region Fourth Option

                var fourthOptionMergeOption = new MergeOptions
                {
                    PreserveBookmarks = true
                };

                foreach (var file in pdfFilesForFourthOption)
                {

                    var additionalPdfs = new List<string>()
                    {
                        file.Replace(Path.GetFileNameWithoutExtension(file), "EK-A"),
                    };

                    var fourthOptionMergeSequence = new MergeSequence
                    {
                        MainDocument = file,
                        AdditionalDocuments = additionalPdfs,
                        OutputPath = file,
                        Options = thirdOptionMergeOption
                    };

                    merger.MergePdf(fourthOptionMergeSequence);

                    Console.WriteLine();
                }

                #endregion
            }
        }

        #endregion
    }
}
