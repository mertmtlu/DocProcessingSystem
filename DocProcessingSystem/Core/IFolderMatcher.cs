using DocProcessingSystem.Models;

namespace DocProcessingSystem.Core
{
    /// <summary>
    /// Interface for matching folders with analysis groups
    /// </summary>
    public interface IFolderMatcher
    {
        /// <summary>
        /// Determines if a folder matches with an analysis group
        /// </summary>
        bool IsMatch(string folderPath, FolderGroup group);
    }
}
