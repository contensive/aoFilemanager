
using System.Collections.Generic;

namespace Contensive.Addons.aoFileManager2.Models.View {
    //
    // ====================================================================================================
    /// <summary>
    /// View model for the file manager Mustache template
    /// </summary>
    public class FileManagerViewModel {
        public string currentPath { get; set; }
        public string currentTreeItemId { get; set; }
        public string selectedFileSystem { get; set; }
        public string parentUrl { get; set; }
        public List<FileSystemOptionViewModel> fileSystemOptions { get; set; } = new List<FileSystemOptionViewModel>();
        public List<TreeItemViewModel> folderTree { get; set; } = new List<TreeItemViewModel>();
        public List<FolderViewModel> folders { get; set; } = new List<FolderViewModel>();
        public List<FileViewModel> files { get; set; } = new List<FileViewModel>();
    }
    //
    // ====================================================================================================
    /// <summary>
    /// File system selector dropdown option
    /// </summary>
    public class FileSystemOptionViewModel {
        public string value { get; set; }
        public string label { get; set; }
        public bool selected { get; set; }
    }
    //
    // ====================================================================================================
    /// <summary>
    /// Tree item for the left-panel folder navigation
    /// </summary>
    public class TreeItemViewModel {
        public string treeItemClass { get; set; }
        public string treeItemId { get; set; }
        public bool hasChildren { get; set; }
        public string folderName { get; set; }
        public string folderUrl { get; set; }
        public TreeChildrenViewModel children { get; set; }
    }
    //
    // ====================================================================================================
    /// <summary>
    /// Container for child tree items (Mustache section rendering)
    /// </summary>
    public class TreeChildrenViewModel {
        public List<TreeItemViewModel> items { get; set; } = new List<TreeItemViewModel>();
    }
    //
    // ====================================================================================================
    /// <summary>
    /// Folder row in the right-panel listing
    /// </summary>
    public class FolderViewModel {
        public string rowClass { get; set; }
        public string folderName { get; set; }
        public string folderUrl { get; set; }
    }
    //
    // ====================================================================================================
    /// <summary>
    /// File row in the right-panel listing
    /// </summary>
    public class FileViewModel {
        public string rowClass { get; set; }
        public string fileName { get; set; }
        public string fileUrl { get; set; }
        public bool hasLink { get; set; }
        public string editUrl { get; set; }
        public string fileSize { get; set; }
        public string modifiedDate { get; set; }
    }
}
