# File Manager Mustache Template - Data Model

## Overview

The `filemanager.html` Mustache template requires a specific data structure to render the file manager interface. This document describes the expected data model.

## Usage in Contensive

```csharp
// Load the template
string template = cp.Layout.GetLayout("filemanager.html");

// Create your data model
var model = new FileManagerViewModel {
    // ... populate as described below
};

// Render the template
string html = cp.Mustache.Render(template, model);
```

## Data Model Structure

### Root Level Properties

| Property | Type | Description |
|----------|------|-------------|
| `formAction` | string | URL for the form submission (handles file uploads, folder creation, deletions) |
| `currentPath` | string | The currently displayed path (e.g., `\`, `\folder1\folder2\`) |
| `currentTreeItemId` | string | ID of the currently selected tree item for expansion |
| `folderTree` | array | Array of folder tree items for left navigation panel |
| `folders` | array | Array of folder objects to display in the current directory |
| `files` | array | Array of file objects to display in the current directory |

### Folder Tree Structure (`folderTree`)

Represents the hierarchical folder structure in the left navigation panel.

```json
{
    "folderTree": [
        {
            "treeItemClass": "mklc",
            "treeItemId": "-",
            "hasChildren": true,
            "folderName": "\\",
            "folderUrl": "?CurrentPath=\\",
            "children": {
                "items": [
                    {
                        "treeItemClass": "mklb",
                        "treeItemId": "-SUBFOLDER-",
                        "folderName": "subfolder",
                        "folderUrl": "?CurrentPath=\\subfolder\\"
                    }
                ]
            }
        }
    ]
}
```

**Tree Item Properties:**
- `treeItemClass`: CSS class for the tree item (`mklc` for containers with children, `mklb` for leaf nodes)
- `treeItemId`: Unique ID for the tree item (typically `-FOLDERNAME-` format)
- `hasChildren`: Boolean indicating if this folder has subfolders
- `folderName`: Display name of the folder
- `folderUrl`: URL to navigate to this folder (should include query parameters like `CurrentPath`, `addonguid`, `addonid`)
- `children`: Optional object containing nested folder structure
  - `items`: Array of child folder objects

### Folders Array (`folders`)

Represents folders in the current directory listing (right panel).

```json
{
    "folders": [
        {
            "rowClass": "ccPanelRowOdd",
            "folderName": "Documents",
            "folderUrl": "?CurrentPath=\\Documents\\"
        },
        {
            "rowClass": "ccPanelRowEven",
            "folderName": "Images",
            "folderUrl": "?CurrentPath=\\Images\\"
        }
    ]
}
```

**Folder Properties:**
- `rowClass`: CSS class for alternating row colors (`ccPanelRowOdd` or `ccPanelRowEven`)
- `folderName`: Name of the folder (used for display and delete checkbox value)
- `folderUrl`: URL to navigate into this folder

**Note:** Alternate `rowClass` values between `ccPanelRowOdd` and `ccPanelRowEven` for visual distinction.

### Files Array (`files`)

Represents files in the current directory listing (right panel).

```json
{
    "files": [
        {
            "rowClass": "ccPanelRowOdd",
            "fileName": "document.pdf",
            "fileUrl": "/files/document.pdf",
            "editUrl": "?EditFilename=C:\\path\\to\\document.pdf",
            "fileSize": "2048576",
            "modifiedDate": "2/15/2026 10:30:00 AM"
        },
        {
            "rowClass": "ccPanelRowEven",
            "fileName": "image.jpg",
            "fileUrl": "/files/image.jpg",
            "editUrl": "?EditFilename=C:\\path\\to\\image.jpg",
            "fileSize": "1024000",
            "modifiedDate": "2/14/2026 3:45:12 PM"
        }
    ]
}
```

**File Properties:**
- `rowClass`: CSS class for alternating row colors (`ccPanelRowOdd` or `ccPanelRowEven`)
- `fileName`: Name of the file (used for display and delete checkbox value)
- `fileUrl`: Public URL to download/view the file (opens in new tab)
- `editUrl`: URL to edit the file (should include EditFilename parameter with full file path)
- `fileSize`: File size in bytes (displayed as-is, format as needed before passing to template)
- `modifiedDate`: Last modified timestamp (formatted string)

**Note:** Continue alternating `rowClass` values from the folders array to maintain consistent striping across folders and files.

## Example Complete Model

```json
{
    "formAction": "https://example.com/filemanager?action=submit",
    "currentPath": "\\",
    "currentTreeItemId": "-",
    "folderTree": [
        {
            "treeItemClass": "mklc",
            "treeItemId": "-",
            "hasChildren": true,
            "folderName": "\\",
            "folderUrl": "?CurrentPath=\\",
            "children": {
                "items": [
                    {
                        "treeItemClass": "mklb",
                        "treeItemId": "-DOCUMENTS-",
                        "folderName": "Documents",
                        "folderUrl": "?CurrentPath=\\Documents\\"
                    },
                    {
                        "treeItemClass": "mklb",
                        "treeItemId": "-IMAGES-",
                        "folderName": "Images",
                        "folderUrl": "?CurrentPath=\\Images\\"
                    }
                ]
            }
        }
    ],
    "folders": [
        {
            "rowClass": "ccPanelRowOdd",
            "folderName": "Documents",
            "folderUrl": "?CurrentPath=\\Documents\\"
        },
        {
            "rowClass": "ccPanelRowEven",
            "folderName": "Images",
            "folderUrl": "?CurrentPath=\\Images\\"
        }
    ],
    "files": [
        {
            "rowClass": "ccPanelRowOdd",
            "fileName": "readme.txt",
            "fileUrl": "/files/readme.txt",
            "editUrl": "?EditFilename=C:\\files\\readme.txt",
            "fileSize": "1024",
            "modifiedDate": "2/15/2026 10:00:00 AM"
        }
    ]
}
```

## Implementation Notes

1. **Row Striping**: Ensure folders and files alternate between `ccPanelRowOdd` and `ccPanelRowEven` classes. Start with folders, then continue the pattern with files.

2. **URL Construction**: All URLs should include necessary query parameters:
   - `CurrentPath`: The path being navigated to
   - `addonguid`: Your addon GUID (if required)
   - `addonid`: Your addon ID (if required)
   - `EditFilename`: Full server path for file editing

3. **Tree Expansion**: The `currentTreeItemId` should match the `treeItemId` of the folder currently being viewed to ensure proper tree expansion via JavaScript.

4. **File Sizes**: Format file sizes appropriately before passing to the template (bytes, KB, MB, etc.).

5. **Date Formatting**: Ensure dates are formatted consistently (e.g., "M/d/yyyy h:mm:ss tt").

6. **Tree JavaScript**: The template assumes `convertTrees()` and `expandToItem()` JavaScript functions are available globally. These should be included in your page layout or referenced via script tags.

## CSS Classes Used

The template defines these custom CSS classes:
- `.fm-container`: Main container styling
- `.fm-tree-panel`: Left navigation panel
- `.fm-content-panel`: Right content panel
- `.fm-header-row`: Table header row
- `.fm-row-odd` / `.fm-row-even`: Alternating row styles
- `.fm-cell-center` / `.fm-cell-left` / `.fm-cell-right`: Cell alignment
- `.fm-footer-row`: Footer input rows
- `.fm-submit-row`: Submit button row

It also uses existing Contensive classes:
- `.ccpanel3dreverse`: Panel styling
- `.ccPanelRowOdd` / `.ccPanelRowEven`: Standard Contensive row classes
- `.ccButtonCon`: Button container
- `.mktree`, `.mklc`, `.mklb`, `.mkd`, `.mkb`: Tree navigation classes

## Form Inputs

The template includes these form inputs that will be submitted:
- `DeleteFolderList[]`: Checkboxes for selected folders to delete
- `DeleteFileList[]`: Checkboxes for selected files to delete
- `NewFolder`: Text input for new folder name
- `NewFile`: File upload input
- `Button`: Submit button (value="Apply")
