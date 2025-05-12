using System.Text.RegularExpressions;

namespace DocProcessingSystem.Models
{
    /// <summary>
    /// Helper class for folder information operations
    /// </summary>
    public static class FolderHelper
    {
        /// <summary>
        /// Groups related folders based on naming patterns
        /// </summary>
        public static List<FolderGroup> GroupFolders(string path)
        {
            var result = new List<FolderGroup>();
            var folders = Directory.GetDirectories(path)
                .Select(f => Path.GetFullPath(f))
                .ToList();

            // Create a dictionary to group folders
            var folderDict = new Dictionary<string, List<FolderInfo>>();

            foreach (var folder in folders)
            {
                string baseName = Path.GetFileName(folder);

                // Extract parts using pattern matching
                var (tmNo, buildingCode, buildingTmId) = ExtractParts(baseName);
                if (string.IsNullOrEmpty(tmNo) || string.IsNullOrEmpty(buildingCode))
                {
                    continue; // Skip folders that don't match the pattern
                }

                // Try to extract block if exists, otherwise use an empty string
                string block = ExtractBlock(baseName);

                // Create a grouping key that strips out block and variations
                // This handles different naming patterns
                var groupKeys = GenerateGroupKeys(tmNo, buildingCode, buildingTmId);

                // Try each group key
                bool matched = false;
                foreach (var groupKey in groupKeys)
                {
                    if (!folderDict.ContainsKey(groupKey))
                    {
                        folderDict[groupKey] = new List<FolderInfo>();
                    }

                    // Add the folder to this group
                    folderDict[groupKey].Add(new FolderInfo
                    {
                        Path = folder,
                        TmNo = tmNo,
                        BuildingCode = buildingCode,
                        BuildingTmId = buildingTmId,
                        Block = block
                    });

                    matched = true;
                    break; // Match found, exit loop
                }

                if (!matched)
                {
                    // Create a single-folder group as fallback
                    result.Add(new FolderGroup
                    {
                        Paths = new List<string> { folder },
                        TmNo = tmNo ?? string.Empty,
                        BuildingCode = buildingCode ?? string.Empty,
                        BuildingTmId = buildingTmId ?? "01"
                    });
                }
            }

            // Process the grouped folders
            foreach (var group in folderDict)
            {
                var foldersInfo = group.Value;

                // Create a list of paths
                var paths = foldersInfo.Select(f => f.Path).ToList();

                // Use first folder's details
                var firstFolder = foldersInfo[0];
                result.Add(new FolderGroup
                {
                    Paths = paths,
                    TmNo = firstFolder.TmNo,
                    BuildingCode = firstFolder.BuildingCode,
                    BuildingTmId = firstFolder.BuildingTmId
                });
            }

            return result;
        }

        /// <summary>
        /// Helper class to store folder information
        /// </summary>
        private class FolderInfo
        {
            public string Path { get; set; }
            public string TmNo { get; set; }
            public string BuildingCode { get; set; }
            public string BuildingTmId { get; set; }
            public string Block { get; set; }
        }

        /// <summary>
        /// Extracts TM number, building code, and building TM ID from a folder name or path
        /// </summary>
        public static (string tmNo, string buildingCode, string buildingTmId) ExtractParts(string input)
        {
            // Check if input looks like a file path
            if (input.Contains("\\") || input.Contains("/"))
            {
                // Split the path into components (works with both / and \ separators)
                string[] pathComponents = input.Split(new char[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);

                // Try to match each path component
                foreach (string component in pathComponents)
                {
                    var result = ExtractPartsFromSingleFolder(component);
                    if (result.tmNo != null)
                    {
                        return result; // Return the first valid match
                    }
                }
                return (null, null, null); // No match found in any path component
            }

            // If not a path, process as a single folder name
            return ExtractPartsFromSingleFolder(input);
        }

        /// <summary>
        /// Internal helper method that applies regex patterns to a single folder name
        /// </summary>
        private static (string tmNo, string buildingCode, string buildingTmId) ExtractPartsFromSingleFolder(string folderName)
        {
            // Standard patterns with consistent separator handling
            var standardPatterns = new List<string>
            {
                @"(\d{1,2}-\d{2})[_\s-]*M(\d{2})[_*\s-]*(\d{2}|\d{1})?(?:[_-]([A-Za-z0-9]+))?",
                @"(\d{1,2}-\d{2})[_\s-]*M(\d{2})[_*\s-]*(\d{2}|\d{1})?(?:[_-]([A-Za-z0-9]+))?",
                @"(\d{1,2}-\d{2})[_*\s-]*M(\d{2})[_*\s-]*(\d{2}|\d{1})?(?:[_-]([A-Za-z0-9]+))?(?:\s*\([A-Za-z0-9\s]+\))?"
            };
            // Try standard patterns first
            foreach (string pattern in standardPatterns)
            {
                Match match = Regex.Match(folderName, pattern);
                if (match.Success)
                {
                    string tmNo = match.Groups[1].Value;
                    string buildingCode = match.Groups[2].Value;
                    string buildingTmId = match.Groups[3].Success ? match.Groups[3].Value : "01";
                    return (tmNo, buildingCode, buildingTmId);
                }
            }
            // Try TEI pattern separately due to different group structure
            string teiPattern = @"TEI-B(\d{2})-TM-(\d{2})-DIR-M(\d{2})(?:-(\d{2}|\d{1}))?";
            Match teiMatch = Regex.Match(folderName, teiPattern);
            if (teiMatch.Success)
            {
                string tmNo = $"{teiMatch.Groups[1].Value}-{teiMatch.Groups[2].Value}";
                string buildingCode = teiMatch.Groups[3].Value;
                string buildingTmId = teiMatch.Groups[4].Success ? teiMatch.Groups[4].Value : "01";
                return (tmNo, buildingCode, buildingTmId);
            }
            return (null, null, null);
        }

        public static (string tmNo, string buildingCode, string buildingTmId) ExtractParts(string folderName, string preferance)
        {
            // Standard pattern: digits-digits-M+digits(-digits)
            string patternStandard = @"^(\d{1,2}-\d{2})\s*-?M(\d{2})(?:-(\d{2}|\d{1}))?(?:-([A-Za-z0-9]+))?$";

            // TEI pattern: TEI-B+digits-TM-digits-DIR-M+digits(-digits)
            string patternTei = $@"TEI-B(\d{{2}})-TM-(\d{{2}})-{preferance}-M(\d{{2}})(?:-(\d{{2}}|\d{{1}}))?";

            // Try the standard pattern first
            Match match = Regex.Match(folderName, patternStandard);
            if (match.Success)
            {
                string tmNo = match.Groups[1].Value;                   // e.g., "18-10"
                string buildingCode = match.Groups[2].Value;           // e.g., "02"
                string buildingTmId = match.Groups[3].Success
                    ? match.Groups[3].Value
                    : "01";                                            // Default to 01 if not specified

                return (tmNo, buildingCode, buildingTmId);
            }

            // Try the TEI pattern
            match = Regex.Match(folderName, patternTei);
            if (match.Success)
            {
                string buildingCode = match.Groups[3].Value;           // e.g., "02"
                string tmNo = $"{match.Groups[1].Value}-{match.Groups[2].Value}"; // e.g., "05-13"
                string buildingTmId = match.Groups[4].Success
                    ? match.Groups[4].Value
                    : "01";                                            // Default to 01 if not specified

                return (tmNo, buildingCode, buildingTmId);
            }

            return (null, null, null);
        }


        /// <summary>
        /// Extracts the block identifier from a folder name, if present
        /// </summary>
        private static string ExtractBlock(string folderName)
        {
            string block = "";
            try
            {
                var parts = folderName.Split('-');
                if (parts.Length > 0)
                {
                    var lastPart = parts[parts.Length - 1];
                    // Check if block is a single letter
                    if (lastPart.Length == 1 && char.IsLetter(lastPart[0]))
                    {
                        block = lastPart;
                    }
                }
            }
            catch
            {
                block = "";
            }
            return block;
        }

        /// <summary>
        /// Generates possible grouping keys for a folder
        /// </summary>
        private static List<string> GenerateGroupKeys(string tmNo, string buildingCode, string buildingTmId)
        {
            return new List<string>
            {
                $"{tmNo}-M{buildingCode}-{buildingTmId}",   // Most specific
                $"{tmNo}-M{buildingCode}",                  // Less specific
                $"{tmNo}_M{buildingCode}-{buildingTmId}",   // Most specific with underscore
                $"{tmNo}_M{buildingCode}"                   // Less specific with underscore
            };
        }
    }
}
