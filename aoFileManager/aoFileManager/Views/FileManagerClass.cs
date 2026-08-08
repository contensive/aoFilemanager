
using System;
using Contensive.Addons.aoFileManager2.Controllers;
using Contensive.Addons.aoFileManager2.Models.View;
using Contensive.BaseClasses;

namespace Contensive.Addons.aoFileManager2.Views {
    //
    // ====================================================================================================
    /// <summary>
    /// File Manager addon - provides a two-panel file browser for Contensive file systems.
    /// Admin-only access. Supports browsing, deleting, uploading, and creating folders.
    /// </summary>
    public class FileManagerClass : AddonBaseClass {
        //
        // ====================================================================================================
        //
        public override object Execute(CPBaseClass cp) {
            try {
                //
                // -- admin-only access
                if (!cp.User.IsAdmin) {
                    return "<p>Access denied. Administrator login required.</p>";
                }
                //
                // -- read inputs
                // -- read fmFileSystem from form POST (avoids conflict with addon ArgumentList "FileSystem" default)
                string fileSystemName = cp.Doc.GetText("fmFileSystem");
                if (string.IsNullOrEmpty(fileSystemName)) {
                    fileSystemName = cp.Doc.GetText("FileSystem");
                }
                if (string.IsNullOrEmpty(fileSystemName)) {
                    fileSystemName = "cdnFiles";
                }
                string currentPath = FileManagerController.SanitizePath(cp.Doc.GetText("CurrentPath"));
                string button = cp.Doc.GetText("Button");
                //
                // -- get the selected file system
                var fs = FileManagerController.GetFileSystem(cp, fileSystemName);
                //
                // -- process form submission
                if (!string.IsNullOrEmpty(button)) {
                    //
                    // -- process deletes
                    string deleteFileList = cp.Doc.GetText("DeleteFileList");
                    FileManagerController.ProcessDeleteFiles(cp, fs, currentPath, deleteFileList);
                    //
                    string deleteFolderList = cp.Doc.GetText("DeleteFolderList");
                    FileManagerController.ProcessDeleteFolders(cp, fs, currentPath, deleteFolderList);
                    //
                    // -- process new folder
                    string newFolder = cp.Doc.GetText("NewFolder");
                    FileManagerController.ProcessNewFolder(cp, fs, currentPath, newFolder);
                    //
                    // -- process file upload
                    FileManagerController.ProcessUpload(cp, fs, currentPath);
                }
                //
                // -- build query string base for links (preserve file system selection and addon context)
                string addonId = cp.Doc.GetText("addonid");
                string addonGuid = cp.Doc.GetText("addonguid");
                string queryBase = "?";
                if (!string.IsNullOrEmpty(addonGuid)) {
                    queryBase += $"addonguid={addonGuid}";
                } else if (!string.IsNullOrEmpty(addonId)) {
                    queryBase += $"addonid={addonId}";
                }
                //
                // -- linkQs includes FileSystem for folder tree links (GET navigation)
                // -- formAction does NOT include FileSystem so the dropdown input is the source of truth on POST
                string linkQs = $"{queryBase}&fmFileSystem={fileSystemName}";
                //
                // -- build view model
                int rowIndex = 0;
                var viewModel = new FileManagerViewModel {
                    currentPath = currentPath,
                    currentTreeItemId = FileManagerController.GetTreeItemId(currentPath),
                    selectedFileSystem = fileSystemName,
                    parentUrl = FileManagerController.BuildParentUrl(currentPath, linkQs),
                    fileSystemOptions = FileManagerController.BuildFileSystemOptions(fileSystemName),
                    folderTree = FileManagerController.BuildFolderTree(fs, currentPath, linkQs),
                    folders = FileManagerController.BuildFolderList(fs, currentPath, linkQs, ref rowIndex),
                    files = FileManagerController.BuildFileList(cp, fs, fileSystemName, currentPath, linkQs, ref rowIndex)
                };
                //
                // -- render the mustache body content
                string layoutHtml = Contensive.Addon.FileManager.Properties.Resources.FileManagerLayout;
                string bodyHtml = cp.Mustache.Render(layoutHtml, viewModel);
                //
                // -- add script to set enctype for file upload support
                bodyHtml += "<script>document.addEventListener('DOMContentLoaded',function(){var f=document.querySelector('form');if(f){f.setAttribute('enctype','multipart/form-data');}});</script>";
                //
                // -- wrap in layout builder for headline and buttons
                var layoutBuilder = cp.AdminUI.CreateLayoutBuilder();
                layoutBuilder.body = bodyHtml;
                layoutBuilder.title = "File Manager";
                layoutBuilder.description = $"{FileManagerController.GetFileSystemLabel(fileSystemName)} - {currentPath}";
                layoutBuilder.isOuterContainer = true;
                layoutBuilder.formActionQueryString = queryBase.TrimStart('?');
                layoutBuilder.addFormHidden("CurrentPath", currentPath);
                layoutBuilder.addFormHidden("fmFileSystem", fileSystemName);
                layoutBuilder.addFormButton("Apply");
                return layoutBuilder.getHtml(cp);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return "<!-- File Manager, Unexpected Exception -->";
            }
        }
    }
}
