VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileViewClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private Const FileSystemLegacyContentFiles = "files"
Private Const FileSystemLegacyWebsiteFiles = "root"
Private Const FileSystemContentFiles = "content files"
Private Const FileSystemWebsiteFiles = "website files"

Private ReplaceString As String
Private PageLink As String
'
Private Main As Object
Private Csv As Object
'Private Main As ccWeb3.MainClass
'
'========================================================================
'   v3.3 Add-on Compatibility
'       To make an Add-on that works the same in v3.3 and v3.4, use this adapter instead of the execute above
'========================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set Csv = CsvObject
    Call Init(MainObject)
    Execute = GetContent(OptionString)
End Function
'
'
'
Public Sub Init(MainObject As Object)
    '
    Set Main = MainObject
    '
    Exit Sub
ErrorTrap:
    Call HandleError("ResourceLibraryClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim BaseFolder As String
    Dim AllowEdit As Boolean
    Dim AllowNavigation As Boolean
    Dim FileSystem As String
    Dim AdminLayout As Boolean
    Dim IncludeForm As Boolean
    Dim AllowFileRadioSelect As Boolean
    '
    If Not (Main Is Nothing) Then
        AdminLayout = kmaEncodeBoolean(Main.GetAggrOption("AdminLayout", OptionString))
        FileSystem = LCase(Main.GetAggrOption("FileSystem", OptionString))
        If AdminLayout Then
            Stream = GetContentFileView3(Main, "", True, True, True, True, True, False, FileSystem, False)
        Else
            BaseFolder = Main.GetAggrOption("BaseFolder", OptionString)
            AllowEdit = kmaEncodeBoolean(Main.GetAggrOption("AllowEdit", OptionString))
            AllowNavigation = kmaEncodeBoolean(Main.GetAggrOption("AllowNavigation", OptionString))
            IncludeForm = kmaEncodeBoolean(Main.GetAggrOption("IncludeForm", OptionString))
            AllowFileRadioSelect = kmaEncodeBoolean(Main.GetAggrOption("AllowFileRadioSelect", OptionString))
            Stream = GetContentFileView3(Main, BaseFolder, AllowEdit, AllowEdit, True, AllowEdit, AllowNavigation, IncludeForm, FileSystem, AllowFileRadioSelect)
        End If
    End If
    '
    GetContent = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("ResourceLibraryClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'==============================================================================================
'   Display a path in the Content Files with links to download and change folders
'==============================================================================================
'
Public Function GetContentFileView(Main As Object, BasePathInput As String, AllowAdd As Boolean, AllowDelete As Boolean, AllowFileLink As Boolean, AllowFileRadios As Boolean, AllowNav As Boolean) As String
    GetContentFileView = GetContentFileView3(Main, BasePathInput, AllowAdd, AllowDelete, AllowFileLink, AllowFileRadios, AllowNav, True, FileSystemContentFiles, False)
End Function
'
'==============================================================================================
'   Display a path in the Content Files with links to download and change folders
'==============================================================================================
'
Public Function GetContentFileView2(Main As Object, BasePathInput As String, AllowAdd As Boolean, AllowDelete As Boolean, AllowFileLink As Boolean, AllowFileRadios As Boolean, AllowNav As Boolean, IncludeForm As Boolean) As String
    GetContentFileView2 = GetContentFileView3(Main, BasePathInput, AllowAdd, AllowDelete, AllowFileLink, AllowFileRadios, AllowNav, IncludeForm, FileSystemContentFiles, False)
End Function
'
'
'
Private Function GetContentFileView3(Main As Object, BasePathInput As String, AllowAdd As Boolean, AllowDelete As Boolean, AllowFileLink As Boolean, AllowFileEditing As Boolean, AllowNav As Boolean, IncludeForm As Boolean, FileSystem As String, AllowFileRadioSelect As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim srcPathFilename As String
    Dim dstPathFilename As String
    Dim EditFilename As String
    Dim PathFilename As String
    Dim TextEditExtensionList As String
    Dim Pos As Long
    Dim QS As String
    'Dim Remote As ccRemote.RemoteClass
    Dim Column2Ptr As Long
    Dim NamePtr As Long
    Dim ServerFilePath As String
    Dim FS As New kmaFileSystem3.FileSystemClass
    Dim FolderNav As String
    Dim Tree As New menuTreeClass
    Dim TreeBasepath As String
    Dim RQS As String
    Dim CurrentFolder As String
    Dim SourceFolders As String
    Dim FolderSplit() As String
    Dim FolderCount As Long
    Dim FolderPointer As Long
    Dim LineSplit() As String
    Dim FolderLine As String
    Dim FolderName As String
    Dim ParentPath As String
    Dim Position As Long
    Dim CurrentPath As String
    Dim Filename As String
    Dim RowEven As Boolean
    Dim FileURL As String
    Dim PathTest As String
    'Dim IsContentManager As Boolean
    Dim NewFile As String
    Dim NewFolder As String
    Dim InlineStyle As String
    Dim NameCell As String
    Dim IconCell As String
    Dim DateCell As String
    Dim SizeCell As String
    Dim DeleteList As String
    Dim DeleteNames() As String
    Dim Ptr As Long
    Dim DelCell As String
    Dim BasePath As String
    Dim PhysicalFilepath As String
    Dim FilenameExt As String
    '
    Const FileShareColumnCnt = 6
    Const GetTableStart = "<table border=""1"" cellpadding=""0"" cellspacing=""0"" width=""100%""><TR><TD><table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><TR><TD Width=""23""><img src=""/ccLib/images/spacer.gif"" width=""23"" height=""1""></TD><TD Width=""100%""><img src=""/ccLib/images/spacer.gif"" width=""100%"" height=""1""></TD></TR>"
    Const GetTableEnd = "</table></td></tr></table>"
    '
    Const SpacerImage = "<img src=""/ccLib/Images/spacer.gif"" width=""23"" height=""22"" border=""0"">"
    Const FolderOpenImage = "<img src=""/ccLib/Images/iconfolderopen.gif"" width=""23"" height=""22"" border=""0"">"
    Const FolderClosedImage = "<img src=""/ccLib/Images/iconfolderclosed.gif"" width=""23"" height=""22"" border=""0"">"
    Const FileNew = "<img src=""/cclib/images/IconContentAdd.gif"" width=18 height=22 border=0>"
    Const FileEdit = "<img src=""/cclib/images/IconContentEdit.gif"" width=18 height=22 border=0>"
    Const FolderNew = "<img src=""/cclib/images/IconFolderAdd2.gif"" width=23 height=22 border=0>"
    '
    FileSystem = LCase(FileSystem)
    If (FileSystem = FileSystemLegacyWebsiteFiles) Then
        FileSystem = FileSystemWebsiteFiles
    ElseIf (FileSystem = FileSystemLegacyContentFiles) Then
        FileSystem = FileSystemContentFiles
    End If
    If (FileSystem = FileSystemWebsiteFiles) Then
        ServerFilePath = "\"
        PhysicalFilepath = Main.PhysicalWWWPath
        If Right(PhysicalFilepath, 1) <> "\" Then
            PhysicalFilepath = PhysicalFilepath & "\"
        End If
    Else
        ServerFilePath = Main.ServerFilePath
        PhysicalFilepath = Main.PhysicalFilepath
    End If
    'PhysicalFilepath = "c:\"
    'IsContentManager = main.IsContentManager()
    '
    If InStr(1, BasePathInput, "..") <> 0 Then
        Call Main.AddUserError("Illegal characters in base path.")
        BasePathInput = "\"
    End If
    BasePath = Trim(BasePathInput)
    BasePath = Replace(BasePath, "/", "\")
    If BasePath = "" Then
        BasePath = "\"
    End If
    If Mid(BasePath, 1, 1) <> "\" Then
        BasePath = "\" & BasePath
    End If
    If Mid(BasePath, Len(BasePath), 1) <> "\" Then
        BasePath = BasePath & "\"
    End If
    If BasePath <> "\" Then
        Call Main.CreateFileFolder(PhysicalFilepath & Mid(BasePath, 2))
    End If
    '
    EditFilename = Main.GetStreamText("EditFilename")
    '
    ' Determine current folder
    '
    CurrentPath = Main.GetStreamText("CurrentPath")
    If CurrentPath = "" Then
        CurrentPath = BasePath
    ElseIf InStr(1, CurrentPath, "..") <> 0 Then
        Call Main.AddUserError("Illegal characters in current path.")
        CurrentPath = BasePath
    End If
    If Mid(CurrentPath, 1, 1) <> "\" Then
        CurrentPath = "\" & CurrentPath
    End If
    If Mid(CurrentPath, Len(CurrentPath), 1) <> "\" Then
        CurrentPath = CurrentPath & "\"
    End If
    '
    ' Block CurrentPath to BasePath
    '
    If InStr(1, CurrentPath, BasePath, vbTextCompare) <> 1 Then
        CurrentPath = BasePath
    End If
    '
    ' Calculate CurrentFolder
    '
    If CurrentPath = "\" Then
        CurrentFolder = "\"
    ElseIf Mid(CurrentPath, Len(CurrentPath), 1) = "\" Then
        CurrentFolder = Mid(CurrentPath, 1, Len(CurrentPath) - 1)
    Else
        CurrentFolder = CurrentPath
    End If
    '
    ' Determine Parent Path
    '
    If CurrentFolder = "\" Then
        ParentPath = "\"
    Else
        Position = InStrRev(CurrentFolder, "\")
        ParentPath = Mid(CurrentPath, 1, Position)
        If ParentPath = "" Then
            ParentPath = "\"
        End If
    End If
    '
    ' Upload new file
    '
    NewFile = Main.GetStreamText("NewFile")
    If NewFile <> "" Then
        Call Main.TestPoint("Uploading " & Filename)
        If FileSystem = FileSystemWebsiteFiles Then
            Call Main.ProcessFormInputFile("NewFile", "Upload")
            '
            ' copy this file to destination
            '
            If CurrentFolder = "\" Then
                dstPathFilename = PhysicalFilepath & NewFile
            Else
                dstPathFilename = PhysicalFilepath & Mid(CurrentPath, 2) & NewFile
            End If
            srcPathFilename = Main.PhysicalFilepath & "Upload\" & NewFile
            Call Main.CopyFile(srcPathFilename, dstPathFilename)
            'Set Remote = New RemoteClass
            'Remote.IPAddress = "127.0.0.1"
            'Remote.port = "4531"
            'QS = "SrcFile=" & kmaEncodeRequestVariable(Main.PhysicalFilepath & "Upload\" & NewFile) & "&DstFile=" & kmaEncodeRequestVariable(dstPathFilename)
            'Call Remote.ExecuteCmd("CopyFile", QS)
            'Set Remote = Nothing
        Else
            Call Main.ProcessFormInputFile("NewFile", CurrentFolder)
        End If
    End If
    '
    ' Delete Files
    '
    DeleteList = Main.GetStreamText("DeleteFileList")
    If DeleteList <> "" Then
        DeleteNames = Split(DeleteList, ",")
        For Ptr = 0 To UBound(DeleteNames)
            If CurrentFolder = "\" Then
                Filename = PhysicalFilepath & DeleteNames(Ptr)
            Else
                Filename = PhysicalFilepath & Mid(CurrentPath, 2) & DeleteNames(Ptr)
            End If
            Call Main.TestPoint("deleting " & Filename)
            Call Main.DeleteFile(Filename)
'            If FileSystem = FileSystemWebsiteFiles Then
'                Set Remote = New RemoteClass
'                Remote.IPAddress = "127.0.0.1"
'                Remote.port = "4531"
'                QS = "File=" & kmaEncodeRequestVariable(Filename)
'                Call Remote.ExecuteCmd("DeleteFile", QS)
'                Set Remote = Nothing
'            Else
'                Call Main.TestPoint("deleting " & Filename)
'                Call Main.DeleteFile(Filename)
'            End If
        Next
    End If
    '
    ' Create new folder
    '
    NewFolder = Main.GetStreamText("NewFolder")
    If NewFolder <> "" Then
        If CurrentFolder = "\" Then
            PathTest = PhysicalFilepath & NewFolder & "\"
        Else
            PathTest = PhysicalFilepath & Mid(CurrentPath, 2) & NewFolder & "\"
        End If
        Call Main.TestPoint("Create Folder " & PathTest)
        PathTest = PathTest & "test.txt"
        Call Main.SaveFile(PathTest, "content")
        Call Main.DeleteFile(PathTest)
'        If FileSystem = FileSystemWebsiteFiles Then
'            Set Remote = New RemoteClass
'            Remote.IPAddress = "127.0.0.1"
'            Remote.port = "4531"
'            QS = "Folder=" & kmaEncodeRequestVariable(PathTest)
'            Call Remote.ExecuteCmd("CreateFolder", QS)
'            Set Remote = Nothing
'        Else
'            PathTest = PathTest & "test.txt"
'            Call Main.SaveFile(PathTest, "content")
'            Call Main.DeleteFile(PathTest)
'        End If
    End If
    '
    ' Delete Folders
    '
    DeleteList = Main.GetStreamText("DeleteFolderList")
    If DeleteList <> "" Then
        DeleteNames = Split(DeleteList, ",")
        For Ptr = 0 To UBound(DeleteNames)
            If DeleteNames(Ptr) <> "" Then
                If CurrentFolder = "\" Then
                    PathTest = PhysicalFilepath & DeleteNames(Ptr)
                Else
                    PathTest = PhysicalFilepath & Mid(CurrentPath, 2) & DeleteNames(Ptr)
                End If
                Call Main.TestPoint("delete file folder " & PathTest)
                Set FS = New FileSystemClass
                Call FS.DeleteFileFolder(PathTest)
    '            If FileSystem = FileSystemWebsiteFiles Then
    '                Set Remote = New RemoteClass
    '                Remote.IPAddress = "127.0.0.1"
    '                Remote.port = "4531"
    '                QS = "Folder=" & kmaEncodeRequestVariable(PathTest)
    '                Call Remote.ExecuteCmd("DeleteFolder", QS)
    '                Set Remote = Nothing
    '            Else
    '                Set FS = New FileSystemClass
    '                Call FS.DeleteFileFolder(PathTest)
    '                'Call Main.DeleteFile(Filename)
    '            End If
            End If
        Next
    End If
    '
    ' Save the editor file
    '
Dim EditorFile As String
Dim EditorContent As String
Dim TempPathFilename As String

    If AllowFileEditing Then
        EditorFile = Main.GetStreamText("editorfile")
        If EditorFile <> "" Then
            EditorContent = Main.GetStreamText("editorcontent")
            Call Main.SaveFile(EditorFile, EditorContent)
            'TempPathFilename = Main.PhysicalFilepath & "Upload\EditorContent.tmp"
            'Call Main.SaveFile(TempPathFilename, EditorContent)
            'Set FS = New FileSystemClass
            'Set Remote = New RemoteClass
            'Remote.IPAddress = "127.0.0.1"
            'Remote.port = "4531"
            'QS = "SrcFile=" & kmaEncodeRequestVariable(TempPathFilename) & "&DstFile=" & kmaEncodeRequestVariable(EditorFile)
            'Call Remote.ExecuteCmd("CopyFile", QS)
            'Call Main.DeleteFile(TempPathFilename)
        End If
    End If
    '
    ' Display the Form
    '
    If AllowFileEditing And (EditFilename <> "") Then
        '
        ' Display Text Editor
        '
        InlineStyle = " style=""" _
            & "background-color: #d0d0d0; " _
            & "border-top: 1px solid #f0f0f0; " _
            & "border-bottom: 1px solid #a0a0a0; " _
            & "border-left: 1px solid #f0f0f0 ; " _
            & "border-right: 1px solid #a0a0a0 ; " _
            & "padding: 4px; " _
            & "text-align:left;" _
            & " """
        GetContentFileView3 = GetContentFileView3 _
            & "<div " & InlineStyle & ">Editing " & EditFilename & "</div>" _
            & Main.GetFormInputTextExpandable("editorcontent", Main.ReadFile(EditFilename), 20) _
            & Main.GetFormInputHidden("editorfile", EditFilename) _
            & Main.GetFormInputHidden("currentpath", CurrentPath)
    Else
        '
        ' Display File List
        '
        '
        ' Sub-Folders
        '
        'GetContentFileView3 = GetContentFileView3 & Main.GetContentCopy2("Custom Page - File Share Resource", , ContentFileViewDefaultCopy)
        GetContentFileView3 = GetContentFileView3 _
            & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" _
            & ""
        'GetContentFileView3 = GetContentFileView3 _
        '    & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" _
        '    & "<TR>" _
        '    & "<TD Width=""23""><img src=""/ccLib/images/spacer.gif"" width=""23"" height=""1""></TD>" _
        '    & "<TD Width=""23""><img src=""/ccLib/images/spacer.gif"" width=""23"" height=""1""></TD>" _
        '    & "<TD Width=""100%""><img src=""/ccLib/images/spacer.gif"" width=""100%"" height=""1""></TD>" _
        '    & "<TD colspan=" & (FileShareColumnCnt - 3) & "<img src=/ccLib/images/spacer.gif width=1 height=1></TD>" _
        '    & "</TR>"
        '
        ' folder path header
        '
        InlineStyle = " style="" " _
            & "background-color: #d0d0d0; " _
            & "border-top: 1px solid #f0f0f0; " _
            & "border-bottom: 1px solid #a0a0a0; " _
            & "padding: 10px; " _
            & " """
        GetContentFileView3 = GetContentFileView3 & "<TR>" _
            & "<TD colspan=" & FileShareColumnCnt & " " & InlineStyle & ">" & CurrentFolder & "</TD>" _
            & "</TR>"
        '
        ' column headers
        '
        InlineStyle = " style=""" _
            & "background-color: #d0d0d0;" _
            & "border-top: 1px solid #f0f0f0;" _
            & "border-bottom: 1px solid #a0a0a0;" _
            & "border-left: 1px solid #f0f0f0;" _
            & "border-right: 1px solid #a0a0a0;" _
            & "padding:4px; " _
            & "" _
            & "font-size:90%;" _
            & "color:#444;" _
            & ""
        If AllowDelete Then
        'If IsContentManager Then
            DelCell = "Del"
        Else
            DelCell = "&nbsp;"
        End If
        GetContentFileView3 = GetContentFileView3 & vbCrLf _
            & vbCrLf & "<TR>" _
            & vbCrLf & vbTab & "<TD colspan=1 " & InlineStyle & "width:20px;text-align:center;"">" & DelCell & "</TD>" _
            & vbCrLf & vbTab & "<TD colspan=1 " & InlineStyle & "width:20px;text-align:center;"">&nbsp;</TD>" _
            & vbCrLf & vbTab & "<TD colspan=1 " & InlineStyle & "text-align:left;"">Name</TD>" _
            & vbCrLf & vbTab & "<TD colspan=1 " & InlineStyle & "width:50px;text-align:right;"">Size</TD>" _
            & vbCrLf & vbTab & "<TD colspan=1 " & InlineStyle & "width:100px;text-align:left;"">Modified</TD>" _
            & vbCrLf & vbTab & "<TD colspan=" & FileShareColumnCnt - 5 & " " & InlineStyle & "width:20px;text-align:center;"">&nbsp;</TD>" _
            & vbCrLf & "</TR>"
    
        If CurrentPath <> BasePath Then
            DelCell = "&nbsp;"
            IconCell = "<A href=""" & Main.ServerPage & "?CurrentPath=" & ParentPath & "&" & Main.RefreshQueryString & """>" & FolderOpenImage & "</A>"
            NameCell = "<A href=""" & Main.ServerPage & "?CurrentPath=" & ParentPath & "&" & Main.RefreshQueryString & """>..</A>"
            SizeCell = ""
            DateCell = ""
            GetContentFileView3 = GetContentFileView3 & GetContentFileView_GetRow(Main, RowEven, DelCell, IconCell, NameCell, SizeCell, DateCell, "&nbsp", FileShareColumnCnt)
        End If
        '
Dim Source As String
        Source = PhysicalFilepath & CurrentFolder
        'Source = "c:" & CurrentFolder
        SourceFolders = Main.GetFolderList(Source)
        'SourceFolders = Main.GetVirtualFolderList(CurrentFolder)
        If SourceFolders <> "" Then
            FolderSplit = Split(SourceFolders, vbCrLf)
            FolderCount = UBound(FolderSplit) + 1
            For FolderPointer = 0 To FolderCount - 1
                FolderLine = FolderSplit(FolderPointer)
                If FolderLine <> "" Then
                    LineSplit = Split(FolderLine, ",")
                    '
                    ' account for commas in filenames, delimited by commas
                    '
                    Column2Ptr = 2
                    Do While Not IsDate(LineSplit(Column2Ptr)) And Column2Ptr < UBound(LineSplit)
                        Column2Ptr = Column2Ptr + 1
                    Loop
                    FolderName = ""
                    For NamePtr = 0 To Column2Ptr - 2
                        FolderName = FolderName & "," & LineSplit(NamePtr)
                    Next
                    If FolderName <> "" Then
                        FolderName = Mid(FolderName, 2)
                    End If
                    If AllowDelete Then
                    'If IsContentManager Then
                        DelCell = "<INPUT TYPE=CheckBox NAME=DeleteFolderList VALUE=""" & Main.encodeHTML(FolderName) & """>"
                    Else
                        DelCell = "&nbsp;"
                    End If
                    IconCell = "<A href=""" & Main.ServerPage & "?CurrentPath=" & kmaEncodeRequestVariable(CurrentPath & FolderName) & "\" & "&" & Main.RefreshQueryString & """>" & FolderClosedImage & "</A>"
                    NameCell = "<A href=""" & Main.ServerPage & "?CurrentPath=" & kmaEncodeRequestVariable(CurrentPath & FolderName) & "\" & "&" & Main.RefreshQueryString & """>" & FolderName & "</A>"
                    SizeCell = LineSplit(Column2Ptr + 3)
                    DateCell = LineSplit(Column2Ptr)
                    GetContentFileView3 = GetContentFileView3 & GetContentFileView_GetRow(Main, RowEven, DelCell, IconCell, NameCell, SizeCell, DateCell, "&nbsp", FileShareColumnCnt)
                End If
            Next
        End If
        '
        ' Files
        '
        SourceFolders = Main.GetFileList(PhysicalFilepath & CurrentFolder)
        'SourceFolders = Main.GetVirtualFileList(CurrentFolder)
        If SourceFolders = "" Then
            DelCell = "&nbsp;"
            IconCell = SpacerImage
            NameCell = "no files were found in this folder"
            SizeCell = ""
            DateCell = ""
            GetContentFileView3 = GetContentFileView3 & GetContentFileView_GetRow(Main, RowEven, DelCell, IconCell, NameCell, SizeCell, DateCell, "", FileShareColumnCnt)
        Else
            TextEditExtensionList = Main.GetSiteProperty("FileManagerTextEditExtensionList", "css,js,txt,html,htm,asp,aspx,php,log")
            FolderSplit = Split(SourceFolders, vbCrLf)
            FolderCount = UBound(FolderSplit) + 1
            For FolderPointer = 0 To FolderCount - 1
                FolderLine = FolderSplit(FolderPointer)
                If FolderLine <> "" Then
                    LineSplit = Split(FolderLine, ",")
                    '
                    ' account for commas in filenames, delimited by commas
                    '
                    Column2Ptr = 2
                    Do While Not IsDate(LineSplit(Column2Ptr)) And Column2Ptr < UBound(LineSplit)
                        Column2Ptr = Column2Ptr + 1
                    Loop
                    Filename = ""
                    For NamePtr = 0 To Column2Ptr - 2
                        Filename = Filename & "," & LineSplit(NamePtr)
                    Next
                    If Filename <> "" Then
                        Filename = Mid(Filename, 2)
                    End If
                    FilenameExt = ""
                    Pos = InStrRev(Filename, ".")
                    If Pos > 0 Then
                        FilenameExt = Mid(Filename, Pos + 1)
                    End If
                    'Filename = LineSplit(0)
                    FileURL = Replace(ServerFilePath & CurrentPath & Filename, "\", "/")
                    FileURL = Replace(FileURL, "//", "/")
                    FileURL = kmaEncodeURL(FileURL)
                    If AllowDelete Then
                    'If IsContentManager Then
                        DelCell = "<INPUT TYPE=CheckBox NAME=DeleteFileList VALUE=""" & Main.encodeHTML(Filename) & """>"
                    ElseIf AllowFileRadioSelect Then
                        DelCell = "<INPUT TYPE=Radio NAME=SelectFile VALUE=""" & Main.encodeHTML(CurrentFolder & "\" & Filename) & """>"
                    Else
                        DelCell = "&nbsp;"
                    End If
                    IconCell = SpacerImage
                    If AllowFileEditing And (InStr(1, TextEditExtensionList, FilenameExt, vbTextCompare) > 0) Then
                        If FileSystem = FileSystemWebsiteFiles Then
                            PathFilename = Main.PhysicalWWWPath & CurrentPath & Filename
                        Else
                            PathFilename = Main.PhysicalFilepath & CurrentPath & Filename
                        End If
                        PathFilename = Replace(PathFilename, "/", "\")
                        PathFilename = Replace(PathFilename, "\\", "\")
                        QS = Main.RefreshQueryString
                        QS = ModifyQueryString(QS, "currentpath", CurrentPath, True)
                        QS = ModifyQueryString(QS, "EditFilename", PathFilename, True)
                        IconCell = "<A href=""" & Main.ServerPage & "?" & QS & """>" & FileEdit & "</A>"
                    End If
                    If AllowFileLink Then
                        NameCell = "<A href=""" & FileURL & """ target=""_blank"">" & Filename & "</A>"
                    Else
                        NameCell = Filename
                    End If
                    SizeCell = LineSplit(Column2Ptr + 3)
                    DateCell = LineSplit(Column2Ptr + 2)
                    GetContentFileView3 = GetContentFileView3 & GetContentFileView_GetRow(Main, RowEven, DelCell, IconCell, NameCell, SizeCell, DateCell, "&nbsp", FileShareColumnCnt)
                End If
            Next
        End If
        '
        ' Create new folder
        '
        If IncludeForm Then
            GetContentFileView3 = Main.GetUploadFormStart() & GetContentFileView3
        End If
        '
        ' Add the currentpath hidden (must be in form)
        '
        GetContentFileView3 = GetContentFileView3 & Main.GetFormInputHidden("CurrentPath", CurrentFolder)
        '
        ' Upload text boxes
        '
        If AllowAdd Then
            GetContentFileView3 = GetContentFileView3 _
                & "<TR><TD colspan=" & FileShareColumnCnt & " style=""background-color:#eee;color:#444;font-size:90%;border-top:2px solid #444;padding:5px;"">As an administrator, you can upload files and create new folders. Others do not see this section.</TD></TR>" _
                & "<TR><TD colspan=1 style=""border-top:1px solid #ccc;padding:5px;background-color:#eee;color:#888;"">" & FolderNew & "</TD><TD colspan=" & FileShareColumnCnt - 1 & " style=""border-top:1px solid #ccc;background-color:#eee;color:#444;padding:5px;font-size:90%;"">New Folder&nbsp;" & Main.GetFormInputText("NewFolder", "", 1, 20) & "</TD></TR>" _
                & "<TR><TD colspan=1 style=""border-top:1px solid #ccc;padding:5px;background-color:#eee;color:#888;"">" & FileNew & "</TD><TD colspan=" & FileShareColumnCnt - 1 & " style=""border-top:1px solid #ccc;background-color:#eee;color:#444;padding:5px;font-size:90%;"">Upload File&nbsp;" & Main.GetFormInputFile("NewFile") & "</TD></TR>"
        End If
        '
        ' Close form
        '
        If IncludeForm Then
            GetContentFileView3 = GetContentFileView3 _
                & "<TR><TD colspan=" & FileShareColumnCnt & " style=""border-top: 1px solid #000000 ;"">" & Main.GetPanelButtons("Apply", "Button") & "</TD></TR>" _
                & Main.GetUploadFormEnd()
        End If
        '
        ' Close outer table
        '
        GetContentFileView3 = GetContentFileView3 & "</table>"
    End If
    '
    ' Build FolderNav
    '
    'If False Then
    If Not AllowNav Then
        GetContentFileView3 = "" _
            & "<table border=1 cellpadding=0 cellspacing=0 width=100%><TR>" _
            & "<TD width=100% valign=top>" & GetContentFileView3 & "</TD>" _
            & "</TR></Table>" _
            & ""
    Else
        RQS = Main.RefreshQueryString
        TreeBasepath = BasePath
        
        Call Tree.AddEntry(Replace(TreeBasepath, "\", "-"), "FileView", , , "?" & RQS & "&CurrentPath=" & BasePath, BasePath)
        Call AddTreeFolders(FS, Main, Tree, TreeBasepath, "", RQS, PhysicalFilepath, 1, 2, CurrentPath)
        FolderNav = Tree.GetTree("FileView", Replace(CurrentPath, "\", "-"))
        'FolderNav = Tree.GetTree(Replace(TreeBasepath, "\", "-"), Replace(CurrentPath, "\", "-"))
        'FolderNav = "<div>" & Tree.GetTree(Replace(TreeBasepath, "\", "-"), Replace(CurrentPath, "\", "-")) & "</div>"
        GetContentFileView3 = "" _
            & "<table border=1 cellpadding=0 cellspacing=0 width=100% style=""Background-color:white;""><TR>" _
            & "<TD width=120 valign=top class=ccpanel3dreverse style=""padding:5px;Background-color:white;"">" & FolderNav & "<BR><img src=/cclib/images/spacer.gif width=120 height=1></TD>" _
            & "<TD width=100% valign=top>" & GetContentFileView3 & "</TD>" _
            & "</TR></Table>" _
            & ""
    End If
    '
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleLocalTrapError("GetContentFileView3", "ErrorTrap")
End Function
'
'=============================================================================
'   Table Rows
'=============================================================================
'

Private Function GetContentFileView_GetRow(Main As Object, RowEven As Boolean, DelCell As String, IconCell As String, NameCell As String, SizeCell As String, DateCell As String, NullCell As String, FileShareColumnCnt As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim ClassString As String
    '
    If Main.EncodeBoolean(RowEven) Then
        RowEven = False
        ClassString = " class=""ccPanelRowEven"" "
    Else
        ClassString = " class=""ccPanelRowOdd"" "
        RowEven = True
    End If
    '
    DelCell = Main.EncodeText(DelCell)
    If DelCell = "" Then
        DelCell = "&nbsp;"
    End If
    '
    IconCell = Main.EncodeText(IconCell)
    If IconCell = "" Then
        IconCell = "&nbsp;"
    End If
    '
    NameCell = Main.EncodeText(NameCell)
    '
    If NameCell = "" Then
        GetContentFileView_GetRow = "" _
            & vbCrLf & "<TR>" _
            & vbCrLf & vbTab & "<TD" & ClassString & " style=""vertical-align:middle;padding:4px;text-align:center;"">" & DelCell & "</TD>" _
            & vbCrLf & vbTab & "<TD" & ClassString & " style=""vertical-align:middle;padding:4px;text-align:center;"" colspan=" & FileShareColumnCnt - 1 & ">" & IconCell & "</TD>" _
            & vbCrLf & "</TR>"
    ElseIf SizeCell = "" Then
        GetContentFileView_GetRow = "" _
            & vbCrLf & "<TR>" _
            & vbCrLf & vbTab & "<TD" & ClassString & " style=""vertical-align:middle;padding:4px;text-align:center;"">" & DelCell & "</TD>" _
            & vbCrLf & vbTab & "<TD" & ClassString & " style=""vertical-align:middle;padding:4px;text-align:center;"">" & IconCell & "</TD>" _
            & vbCrLf & vbTab & "<TD" & ClassString & " style=""vertical-align:middle;padding:4px;text-align:left;"" colspan=" & FileShareColumnCnt - 2 & ">" & NameCell & "</TD>" _
            & vbCrLf & "</TR>"
    Else
        GetContentFileView_GetRow = "" _
            & vbCrLf & "<TR>" _
            & vbCrLf & vbTab & "<TD" & ClassString & " style=""vertical-align:middle;padding:4px;text-align:center;"">" & DelCell & "</TD>" _
            & vbCrLf & vbTab & "<TD" & ClassString & " style=""vertical-align:middle;padding:4px;text-align:center;"">" & IconCell & "</TD>" _
            & vbCrLf & vbTab & "<TD" & ClassString & " style=""vertical-align:middle;padding:4px;text-align:left;"">&nbsp;" & NameCell & "</TD>" _
            & vbCrLf & vbTab & "<TD" & ClassString & " style=""vertical-align:middle;padding:4px;text-align:right;"">&nbsp;" & SizeCell & "</TD>" _
            & vbCrLf & vbTab & "<TD" & ClassString & " style=""vertical-align:middle;padding:4px;text-align:right;""><NOBR>&nbsp;" & DateCell & "</NOBR></TD>" _
            & vbCrLf & vbTab & "<TD colspan=" & FileShareColumnCnt - 5 & "" & ClassString & ">" & NullCell & "</TD>" _
            & vbCrLf & "</TR>"
        End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleLocalTrapError("GetContentFileView_GetRow", "ErrorTrap")
End Function
'
'
'
Private Sub AddTreeFolders(kmafs As FileSystemClass, Main As Object, Tree As menuTreeClass, FullPath As String, FolderName As String, RQS As String, PhysicalFilepath As String, Depth As Long, MaxDepthPastCurrentPath As Long, CurrentPath As String)
    On Error GoTo ErrorTrap
    '
    Dim Column2Ptr As Long
    Dim NamePtr As Long
    Dim FolderList As String
    Dim Folders() As String
    Dim Ptr As Long
    Dim ParentMenuName As String
    Dim FolderMenuName As String
    Dim ChildFolderName As String
    Dim ChildFolderCaption As String
    Dim ChildFolder() As String
    Dim FullChildPath As String
    Dim WorkingDepth As Long
    Dim WorkingMaxDepth As Long

    '
    WorkingDepth = Depth
    WorkingMaxDepth = MaxDepthPastCurrentPath
    If (FullPath <> "\") And (InStr(1, CurrentPath, FullPath, vbTextCompare) = 1) Then
        '
        ' path \ does not count as a level match.
        ' The current path being displayed on the right is in the fullpath beling deisplay on the left
        '
        '
        If Len(CurrentPath) = Len(FullPath) Then
            '
            ' The left side has navigated to the path opened on the right
            ' Allow MaxDepth levels past this point
            '
            WorkingMaxDepth = WorkingMaxDepth + 1
        ElseIf Len(CurrentPath) < Len(FullPath) Then
            '
            ' The left side has navigated into the path opened on the right and is past or equal it
            ' let the natural MaxDepth stop the tree
            '
            WorkingMaxDepth = WorkingMaxDepth
        Else
            '
            ' The left is navigating toward the path on the right, but not there yet
            ' Allow this level and one more
            '
            WorkingMaxDepth = WorkingDepth + 1
        End If
    End If
    If WorkingDepth < WorkingMaxDepth Then
        On Error Resume Next
        FolderList = kmafs.GetSubFolders(PhysicalFilepath & Mid(FullPath, 2))
        If Err.Number <> 0 Then
            FolderList = ""
        Else
            On Error GoTo ErrorTrap
            If FolderList <> "" Then
                Folders = Split(FolderList, vbCrLf)
                For Ptr = 0 To UBound(Folders)
                    If Folders(Ptr) <> "" Then
                        ChildFolder = Split(Folders(Ptr), ",")
                        '
                        ' account for commas in filenames, delimited by commas
                        '
                        Column2Ptr = 2
                        Do While Not IsDate(ChildFolder(Column2Ptr)) And Column2Ptr < UBound(ChildFolder)
                            Column2Ptr = Column2Ptr + 1
                        Loop
                        FolderName = ""
                        For NamePtr = 0 To Column2Ptr - 2
                            FolderName = FolderName & "," & ChildFolder(NamePtr)
                        Next
                        If FolderName <> "" Then
                            FolderName = Mid(FolderName, 2)
                        End If
                        ChildFolderName = FolderName
                        ChildFolderCaption = Replace(ChildFolderName, " ", "&nbsp;")
                        FullChildPath = FullPath & ChildFolderName & "\"
                        Call Tree.AddEntry(Replace(FullChildPath, "\", "-"), Replace(FullPath, "\", "-"), , , "?" & RQS & "&CurrentPath=" & FullChildPath, ChildFolderCaption)
                        Call AddTreeFolders(kmafs, Main, Tree, FullChildPath, ChildFolderName, RQS, PhysicalFilepath, WorkingDepth + 1, WorkingMaxDepth, CurrentPath)
                    End If
                Next
            End If
        End If
    End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleLocalTrapError("AddTreeFolders", "ErrorTrap")
End Sub
'
'
'
Private Sub HandleLocalTrapError(MethodName As String, Optional ignore0 As String)
    Call HandleError("FileViewClass", MethodName, Err.Number, Err.Source, Err.Description, True, False)
    End Sub



