namespace DocProcessingSystem.Models
{
    /// <summary>
    /// Represents a group of related folders
    /// </summary>
    public class FolderGroup
    {
        /// <summary>
        /// Paths to the folders in this group
        /// </summary>
        public List<string> Paths { get; set; } = new List<string>();

        /// <summary>
        /// TM number identifier
        /// </summary>
        public string TmNo { get; set; }

        /// <summary>
        /// Building code identifier
        /// </summary>
        public string BuildingCode { get; set; }

        /// <summary>
        /// Building TM identifier
        /// </summary>
        public string BuildingTmId { get; set; }

        /// <summary>
        /// Main folder in the group (typically first one)
        /// </summary>
        public string MainFolder => Paths.FirstOrDefault() ?? string.Empty;

        /// <summary>
        /// Number of paths in this group, represents number of blocks
        /// </summary>
        public int PathCount => Paths.Count;
    }
}
