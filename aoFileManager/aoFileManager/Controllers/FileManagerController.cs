
using System;
using System.Collections.Generic;
using Contensive.BaseClasses;
using Contensive.Addons.aoFileManager2.Models.View;

namespace Contensive.Addons.aoFileManager2.Controllers {
    //
    // ====================================================================================================
    /// <summary>
    /// Business logic for the file manager addon
    /// </summary>
    public static class FileManagerController {
        //
        // ====================================================================================================
        /// <summary>
        /// Returns the file system object based on the dropdown selection
        /// </summary>
        public static CPFileSystemBaseClass GetFileSystem(CPBaseClass cp, string fileSystemName) {
            switch ((fileSystemName ?? "").ToLowerInvariant()) {
                case "wwwfiles":
                    return cp.WwwFiles;
                case "tempfiles":
                    return cp.TempFiles;
                case "privatefiles":
                    return cp.PrivateFiles;
                default:
                    return cp.CdnFiles;
            }
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Returns display label for a file system name
        /// </summary>
        public static string GetFileSystemLabel(string fileSystemName) {
            switch ((fileSystemName ?? "").ToLowerInvariant()) {
                case "wwwfiles":
                    return "Website Files (wwwFiles)";
                case "tempfiles":
                    return "Temp Files (tmpFiles)";
                case "privatefiles":
                    return "Private Files (privateFiles)";
                default:
                    return "CDN Files (cdnFiles)";
            }
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Build the file system dropdown options
        /// </summary>
        public static List<FileSystemOptionViewModel> BuildFileSystemOptions(string selectedFileSystem) {
            string selected = (selectedFileSystem ?? "cdnFiles").ToLowerInvariant();
            return new List<FileSystemOptionViewModel> {
                new FileSystemOptionViewModel { value = "cdnFiles", label = "CDN Files (cdnFiles)", selected = selected == "cdnfiles" },
                new FileSystemOptionViewModel { value = "wwwFiles", label = "Website Files (wwwFiles)", selected = selected == "wwwfiles" },
                new FileSystemOptionViewModel { value = "tempFiles", label = "Temp Files (tmpFiles)", selected = selected == "tempfiles" },
                new FileSystemOptionViewModel { value = "privateFiles", label = "Private Files (privateFiles)", selected = selected == "privatefiles" }
            };
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Sanitize a path to prevent directory traversal. Returns empty string if invalid.
        /// </summary>
        public static string SanitizePath(string path) {
            if (string.IsNullOrWhiteSpace(path)) {
                return "\\";
            }
            // -- reject any path with .. to prevent traversal
            if (path.Contains("..")) {
                return "\\";
            }
            // -- normalize separators to backslash
            string result = path.Replace("/", "\\");
            // -- ensure it starts with backslash
            if (!result.StartsWith("\\")) {
                result = "\\" + result;
            }
            // -- ensure it ends with backslash
            if (!result.EndsWith("\\")) {
                result += "\\";
            }
            return result;
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Convert the UI path (backslash-based) to the Contensive file system path (forward slash)
        /// </summary>
        public static string ToFileSystemPath(string uiPath) {
            string result = (uiPath ?? "").Replace("\\", "/");
            // -- remove leading slash for Contensive API (it uses relative paths)
            if (result.StartsWith("/")) {
                result = result.Substring(1);
            }
            return result;
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Build the tree item ID from a path
        /// </summary>
        public static string GetTreeItemId(string path) {
            if (string.IsNullOrEmpty(path) || path == "\\") {
                return "-";
            }
            string trimmed = path.Trim('\\').Replace("\\", "-");
            return $"-{trimmed.ToUpperInvariant()}-";
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Build the folder list for the left navigation panel.
        /// Shows only the subfolders of the current path (flat list, not recursive).
        /// </summary>
        public static List<TreeItemViewModel> BuildFolderTree(CPFileSystemBaseClass fs, string currentPath, string queryBase) {
            var result = new List<TreeItemViewModel>();
            string fsPath = ToFileSystemPath(currentPath);
            var subFolders = GetSubFolders(fs, fsPath);
            foreach (var folder in subFolders) {
                string folderUiPath = currentPath == "\\"
                    ? $"\\{folder.Name}\\"
                    : $"{currentPath.TrimEnd('\\')}{"\\"}{folder.Name}\\";
                result.Add(new TreeItemViewModel {
                    folderName = folder.Name,
                    folderUrl = $"{queryBase}&CurrentPath={folderUiPath}"
                });
            }
            return result;
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Build the parent URL for navigating up one level from the current path.
        /// Returns empty string if already at root.
        /// </summary>
        public static string BuildParentUrl(string currentPath, string queryBase) {
            if (string.IsNullOrEmpty(currentPath) || currentPath == "\\") {
                return "";
            }
            // -- remove trailing backslash, then find the last backslash
            string trimmed = currentPath.TrimEnd('\\');
            int lastSlash = trimmed.LastIndexOf('\\');
            string parentPath = lastSlash <= 0 ? "\\" : trimmed.Substring(0, lastSlash) + "\\";
            return $"{queryBase}&CurrentPath={parentPath}";
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Safely get subfolders, returning empty list on error
        /// </summary>
        private static List<CPFileSystemBaseClass.FolderDetail> GetSubFolders(CPFileSystemBaseClass fs, string path) {
            try {
                return fs.FolderList(path) ?? new List<CPFileSystemBaseClass.FolderDetail>();
            } catch {
                return new List<CPFileSystemBaseClass.FolderDetail>();
            }
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Build the folder listing for the right panel
        /// </summary>
        public static List<FolderViewModel> BuildFolderList(CPFileSystemBaseClass fs, string currentPath, string queryBase, ref int rowIndex) {
            var result = new List<FolderViewModel>();
            string fsPath = ToFileSystemPath(currentPath);
            var folders = GetSubFolders(fs, fsPath);
            foreach (var folder in folders) {
                string folderUiPath = $"{currentPath.TrimEnd('\\')}{(currentPath == "\\" ? "" : "\\")}{folder.Name}\\";
                if (currentPath == "\\") {
                    folderUiPath = $"\\{folder.Name}\\";
                }
                result.Add(new FolderViewModel {
                    rowClass = (rowIndex % 2 == 0) ? "fm-row-odd" : "fm-row-even",
                    folderName = folder.Name,
                    folderUrl = $"{queryBase}&CurrentPath={folderUiPath}"
                });
                rowIndex++;
            }
            return result;
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Build the file listing for the right panel.
        /// cdnFiles link to cp.Http.CdnFilePathPrefix + path, wwwFiles link to "/" + path,
        /// tempFiles and privateFiles have no link.
        /// </summary>
        public static List<FileViewModel> BuildFileList(CPBaseClass cp, CPFileSystemBaseClass fs, string fileSystemName, string currentPath, string queryBase, ref int rowIndex) {
            var result = new List<FileViewModel>();
            string fsPath = ToFileSystemPath(currentPath);
            string fsNameLower = (fileSystemName ?? "").ToLowerInvariant();
            bool hasLink = (fsNameLower == "cdnfiles" || fsNameLower == "wwwfiles");
            string urlPrefix = "";
            if (fsNameLower == "cdnfiles") {
                urlPrefix = cp.Http.CdnFilePathPrefix;
            } else if (fsNameLower == "wwwfiles") {
                urlPrefix = "/";
            }
            try {
                var files = fs.FileList(fsPath) ?? new List<CPFileSystemBaseClass.FileDetail>();
                foreach (var file in files) {
                    string fileFsPath = string.IsNullOrEmpty(fsPath) ? file.Name : $"{fsPath}{file.Name}";
                    result.Add(new FileViewModel {
                        rowClass = (rowIndex % 2 == 0) ? "fm-row-odd" : "fm-row-even",
                        fileName = file.Name,
                        fileUrl = hasLink ? $"{urlPrefix}{fileFsPath}" : "",
                        hasLink = hasLink,
                        editUrl = $"{queryBase}&CurrentPath={currentPath}&EditFilename={file.Name}",
                        fileSize = FormatFileSize(file.Size),
                        modifiedDate = file.DateLastModified.HasValue ? file.DateLastModified.Value.ToString("M/d/yyyy h:mm tt") : ""
                    });
                    rowIndex++;
                }
            } catch {
                // -- folder may not exist yet, return empty list
            }
            return result;
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Format file size in human-readable form
        /// </summary>
        public static string FormatFileSize(long bytes) {
            if (bytes < 1024) {
                return $"{bytes} B";
            }
            if (bytes < 1024 * 1024) {
                return $"{bytes / 1024.0:F1} KB";
            }
            if (bytes < 1024 * 1024 * 1024) {
                return $"{bytes / (1024.0 * 1024.0):F1} MB";
            }
            return $"{bytes / (1024.0 * 1024.0 * 1024.0):F1} GB";
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Process delete requests for files
        /// </summary>
        public static void ProcessDeleteFiles(CPBaseClass cp, CPFileSystemBaseClass fs, string currentPath, string deleteFileList) {
            if (string.IsNullOrWhiteSpace(deleteFileList)) { return; }
            string fsPath = ToFileSystemPath(currentPath);
            string[] fileNames = deleteFileList.Split(',');
            foreach (string fileName in fileNames) {
                string trimmed = fileName.Trim();
                if (string.IsNullOrEmpty(trimmed)) { continue; }
                try {
                    string fullPath = string.IsNullOrEmpty(fsPath) ? trimmed : $"{fsPath}{trimmed}";
                    fs.DeleteFile(fullPath);
                } catch (Exception ex) {
                    cp.Site.ErrorReport(ex);
                }
            }
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Process delete requests for folders
        /// </summary>
        public static void ProcessDeleteFolders(CPBaseClass cp, CPFileSystemBaseClass fs, string currentPath, string deleteFolderList) {
            if (string.IsNullOrWhiteSpace(deleteFolderList)) { return; }
            string fsPath = ToFileSystemPath(currentPath);
            string[] folderNames = deleteFolderList.Split(',');
            foreach (string folderName in folderNames) {
                string trimmed = folderName.Trim();
                if (string.IsNullOrEmpty(trimmed)) { continue; }
                try {
                    string fullPath = string.IsNullOrEmpty(fsPath) ? trimmed : $"{fsPath}{trimmed}";
                    fs.DeleteFolder(fullPath);
                } catch (Exception ex) {
                    cp.Site.ErrorReport(ex);
                }
            }
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Process new folder creation
        /// </summary>
        public static void ProcessNewFolder(CPBaseClass cp, CPFileSystemBaseClass fs, string currentPath, string newFolderName) {
            if (string.IsNullOrWhiteSpace(newFolderName)) { return; }
            if (newFolderName.Contains("..") || newFolderName.Contains("/") || newFolderName.Contains("\\")) { return; }
            string fsPath = ToFileSystemPath(currentPath);
            try {
                string fullPath = string.IsNullOrEmpty(fsPath) ? newFolderName : $"{fsPath}{newFolderName}";
                fs.CreateFolder(fullPath);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
        }
        //
        // ====================================================================================================
        /// <summary>
        /// Process file upload
        /// </summary>
        public static void ProcessUpload(CPBaseClass cp, CPFileSystemBaseClass fs, string currentPath) {
            string fsPath = ToFileSystemPath(currentPath);
            try {
                string returnFilename = "";
                fs.SaveUpload("NewFile", fsPath, ref returnFilename);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
            }
        }
    }
}
