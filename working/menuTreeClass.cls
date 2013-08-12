VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "menuTreeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'==============================================================================
'
'   Creates custom menus
'   Stores caches of the menus
'   Stores the menu data, and can generate different kind
'
'==============================================================================
'
Const MenuStyleRollOverFlyout = 1
' Const MenuStyleDropDown = 1
Const MenuStyleTree = 2
Const MenuStyleTreeList = 3
Const MenuStyleFlyoutDown = 4
Const MenuStyleFlyoutRight = 5
Const MenuStyleFlyoutUp = 6
Const MenuStyleFlyoutLeft = 7
Const MenuStyleHoverDown = 8
Const MenuStyleHoverRight = 9
Const MenuStyleHoverUp = 10
Const MenuStyleHoverLeft = 11
'
' ----- Each menu item has an MenuEntry
'
Private Type MenuEntryType
    Caption As String           ' What is displayed for this entry (does not need to be unique)
    Name As String              ' Unique name for this entry
    ParentName As String        ' Unique name of the parent entry
    Link As String              ' URL
    Image As String             ' Image
    ImageOver As String         ' Image Over
    ImageOpen As String         ' Image when menu is open
    'StyleSheet As String        ' Stylesheet
    'StyleSheetHover As String   ' Hover Stylesheet
    NewWindow As Boolean        ' True opens link in a new window
    OnClick As String           ' Holds action for onClick
    End Type
'
' ----- A collection of menu entries that have a single function (unique name)
'
'Private Type MenuType
'    Name As String              ' Unique name for this menu
'    Link As String              ' The linked text at the top of this menu
'    LinkLabel As String         ' The linked text at the top of this menu
'    CreateDate As Date          ' DateTime when this panel was created
'    PositionX As Long           ' pixel position on the screen, (default -1)
'    PositionY As Long           ' pixel position on the screen, (default -1)
'    StyleSheet As String        ' Stylesheet to put on the whole menu
'    StyleSheetHover As String   ' Hover Stylesheet to put on the whole menu
'    EntryCount As Long          ' Number of Entries in this panel
'    EntrySize As Long           ' Number of Entries in this panel
'    Entries() As MenuEntryType  ' The menu entries
'    End Type
'
' ----- Local storage
'
'Private iMenuCount As Long          ' Count of Menus in the object
'Private iMenuSize As Long
'Private iMenu() As MenuType
Private iMenuFilePath As String
'
' ----- Menu Entry storage
'
Private iEntryCount As Long          ' Count of Menus in the object
Private iEntrySize As Long
Private iEntry() As MenuEntryType
'
' Private iDQMCount As Long           ' Count of Default Menus for this instance
' Private iDQMCLosed As Boolean       ' true if the menu has been closed
'
Private iTreeCount As Long          ' Count of Tree Menus for this instance
Private iMenuCloseString As String  ' String returned for closing menus
'
Private UsedEntries As String       ' String of EntryNames that have been used (for unique test)
Private EntryIndexName As fastIndex6Class
Private EntryIndexID As fastIndex6Class
'
' ----- RollOverFlyout storage
'
'Private MenuFlyoutCount As Long           ' Count of Default Menus for this instance
Private MenuFlyoutNamePrefix As String    ' Random prefix added to element IDs to avoid namespace collision
Private MenuFlyoutIcon_Local As String      ' string used to mark a button that has a non-hover flyout
' Private RollOverFlyoutClosed As Boolean       ' true if the menu has been closed
Const newmode = True
'
'===============================================================================
'   Create a new Menu Entry
'===============================================================================
'
Public Sub AddEntry(EntryName As Variant, Optional ParentiEntryName As Variant, Optional ImageLink As Variant, Optional ImageOverLink As Variant, Optional Link As Variant, Optional Caption As Variant, Optional OnClickJavascript As Variant, Optional Ignore1 As Variant, Optional ImageOpenLink As Variant, Optional NewWindow As Variant)
    On Error GoTo ErrorTrap
    '
    Dim MenuEntrySize As Long
    Dim iEntryName As String
    Dim UcaseEntryName As String
    Dim iNewWindow As Boolean
    '
    iEntryName = Replace(KmaEncodeMissingText(EntryName, ""), ",", " ")
    UcaseEntryName = UCase(iEntryName)
    '
    If (iEntryName <> "") And (InStr(1, UsedEntries & ",", "," & UcaseEntryName & ",", vbBinaryCompare) = 0) Then
        UsedEntries = UsedEntries & "," & UcaseEntryName
        If iEntryCount >= iEntrySize Then
            iEntrySize = iEntrySize + 10
            ReDim Preserve iEntry(iEntrySize)
        End If
        With iEntry(iEntryCount)
            .Link = KmaEncodeMissingText(Link, "")
            .Image = KmaEncodeMissingText(ImageLink, "")
            .OnClick = KmaEncodeMissingText(OnClickJavascript, "")
            If .Image = "" Then
                '
                ' No image, must have a caption
                '
                .Caption = KmaEncodeMissingText(Caption, iEntryName)
            Else
                '
                ' Image present, caption is extra
                '
                .Caption = KmaEncodeMissingText(Caption, "")
            End If
            .Name = UcaseEntryName
            .ParentName = UCase(KmaEncodeMissingText(ParentiEntryName, ""))
            .ImageOver = KmaEncodeMissingText(ImageOverLink, "")
            .ImageOpen = KmaEncodeMissingText(ImageOpenLink, "")
            .NewWindow = KmaEncodeMissingBoolean(NewWindow, False)
            End With
        Call EntryIndexName.SetPointer(UcaseEntryName, iEntryCount)
        iEntryCount = iEntryCount + 1
        End If
    '
    Exit Sub
    '
ErrorTrap:
    Call HandleClassError("AddEntry", Err.Number, Err.Source, Err.Description)
End Sub
'
'===============================================================================
'   Returns the menu specified, if it is in local storage
'
'   It also creates the menu data in a close string that is returned in GetMenuClose.
'   It must be done there so the link buttons height can be calculated.
'===============================================================================
'
Public Function GetMenu(MenuName As String, Optional StyleSheetPrefix As Variant) As String
    On Error GoTo ErrorTrap
    GetMenu = GetTree(MenuName, "", KmaEncodeMissingText(StyleSheetPrefix, "ccTree"))
    Exit Function
    '
    Dim Link As String
    Dim EntryPointer As Long
    Dim UcaseMenuName As String
    Dim LocalStyleSheetPrefix As String
    '
    ' ----- Get the menu pointer
    '
    If iEntryCount > 0 Then
        UcaseMenuName = MenuName
        LocalStyleSheetPrefix = KmaEncodeMissingText(StyleSheetPrefix, "ccTree")
        For EntryPointer = 0 To iEntryCount - 1
            If iEntry(EntryPointer).Name = UcaseMenuName Then
                Exit For
            End If
        Next
        If EntryPointer < iEntryCount Then
            '
            ' ----- Build the linked -button-
            '
            Link = iEntry(EntryPointer).Link
            If Link = "" Then
                Link = "javascript: ;"
            End If
            '
            GetMenu = vbCrLf _
                & "<DIV id=""tree"" class=""" & LocalStyleSheetPrefix & "Root"" ></DIV>" & vbCrLf
            '
            '   Find the Menu Entry, and create the top element here
            '
            For EntryPointer = 0 To iEntryCount - 1
                With iEntry(EntryPointer)
                    If .Name = UcaseMenuName Then
                        'iMenuCloseString = iMenuCloseString
                        GetMenu = GetMenu _
                            & "<SCRIPT Language=""JavaScript"" type=""text/javascript"">" & vbCrLf _
                            & "var DivLeft,DivTop,ElementObject; " & vbCrLf _
                            & "DivTop = -18; " & vbCrLf _
                            & "DivLeft = 0; " & vbCrLf _
                            & "for (ElementObject=tree;  ElementObject.tagName!='BODY'; ElementObject = ElementObject.offsetParent) { " & vbCrLf _
                            & "    DivTop = DivTop+ElementObject.offsetTop; " & vbCrLf _
                            & "    DivLeft = DivLeft+ElementObject.offsetLeft; " & vbCrLf _
                            & "    } " & vbCrLf _
                            & "var menuBase = new  menuObject(DivTop,DivLeft); " & vbCrLf _
                            & "menuBase.s[0] = new so(0,'" & .Caption & "','" & .Link & "','_blank',''); " & vbCrLf _
                            & GetMenuTreeBranch(.Name, "menuBase.s[0]", "," & EntryPointer) _
                            & "</SCRIPT>" & vbCrLf
                            ' & "<SCRIPT LANGUAGE=""JavaScript"" src=""/cclib/ClientSide/tree30.js""></SCRIPT>" & vbCrLf
                        Exit For
                        End If
                    End With
                Next
            '
            ' ----- Add what is needed to the close string, be carefull of the order
            '
            '
            ' increment the menu count
            '
            iTreeCount = iTreeCount + 1
            End If
        End If
    Exit Function
    '
ErrorTrap:
    Call HandleClassError("GetMenu", Err.Number, Err.Source, Err.Description)
End Function
'
'===============================================================================
'   Gets the Menu Branch for the Tree Menu
'===============================================================================
'
Private Function GetMenuTreeBranch(ParentName As String, JSObject As String, UsedEntries As String) As String
    On Error GoTo ErrorTrap
    '
    Dim EntryPointer As Long
    Dim iUsedEntries As String
    Dim JSChildObject As String
    Dim SubMenuCount As Long
    '
    iUsedEntries = UsedEntries
    SubMenuCount = 0
    For EntryPointer = 0 To iEntryCount - 1
        With iEntry(EntryPointer)
            If .ParentName = ParentName Then
                If (InStr(1, iUsedEntries & ",", "," & EntryPointer & ",") = 0) Then
                    JSChildObject = JSObject & ".s[" & SubMenuCount & "]"
                    iUsedEntries = iUsedEntries & "," & EntryPointer
                    GetMenuTreeBranch = GetMenuTreeBranch _
                        & JSChildObject & " = new so(0,'" & .Caption & "','" & .Link & "','_blank',''); " & vbCrLf _
                        & GetMenuTreeBranch(.Name, JSChildObject, iUsedEntries)
                    SubMenuCount = SubMenuCount + 1
                    End If
                End If
            End With
        Next
    '
    Exit Function
    '
ErrorTrap:
    Call HandleClassError("GetMenuTreeBranch", Err.Number, Err.Source, Err.Description)
End Function

Private Sub Class_Initialize()
    ' iDQMCount = 0
    Set EntryIndexName = New fastIndex6Class
    Randomize
    MenuFlyoutNamePrefix = "id" & CStr(Int(9999 * Rnd()))
End Sub
'
' Temp log file solution
'
Private Sub AppendLogFile(Message As String)
    Dim FS As New FileSystemClass
    '
    Call FS.AppendFile(GetProgramPath() & "\logs\MenuLog" & Int(Now) & ".log", Now & " " & Message)
End Sub
'
'===============================================================================
'   Returns the menu specified, if it is in local storage
'
'   It also creates the menu data in a close string that is returned in GetMenuClose.
'   It must be done there so the link buttons height can be calculated.
'   Uses a simple UL/Stylesheet method, returning to the server with every click
'===============================================================================
'
Private Function GetMenuTreeList(MenuName As String, OpenNodesList As String) As String
    On Error GoTo ErrorTrap
    '
    Dim EntryPointer As Long
    Dim UcaseMenuName As String
    '
    ' ----- Get the menu pointer
    '
    If iEntryCount > 0 Then
        UcaseMenuName = UCase(MenuName)
        EntryPointer = EntryIndexName.GetPointer(UcaseMenuName)
        GetMenuTreeList = GetMenuTreeListBranch2(EntryPointer, "", OpenNodesList)
        Exit Function
    End If
    Exit Function
    '
ErrorTrap:
    Call HandleClassError("GetMenuTreeList", Err.Number, Err.Source, Err.Description)
End Function
'
'===============================================================================
'   Gets the Menu Branch for the Tree Menu
'===============================================================================
'
Private Function GetMenuTreeListBranch2(NodePointer As Long, UsedEntriesList As String, OpenNodesList As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Link As String
    Dim EntryPointer As Long
    Dim UcaseNodeName As String
    Dim Image As String
    Dim Caption As String
    '
    If iEntryCount > 0 Then
        '
        ' Output this node
        '
        If InStr(1, "," & CStr(NodePointer) & ",", "," & UsedEntriesList & ",") = 0 Then
            GetMenuTreeListBranch2 = GetMenuTreeListBranch2 & "<ul Style=""list-style-type: none; margin-left: 20px"">"
            '
            ' The Node has not already been used in this branch
            '
            Caption = iEntry(NodePointer).Caption
            Link = kmaEncodeHTML(iEntry(NodePointer).Link)
            If Link <> "" Then
                Caption = "<A TARGET=""_blank"" HREF=""" & Link & """>" & Caption & "</A>"
                End If
            '
            If InStr(1, "," & OpenNodesList & ",", "," & CStr(NodePointer) & ",") = 0 Then
                '
                ' The branch is closed
                '
                Image = iEntry(NodePointer).Image
                GetMenuTreeListBranch2 = GetMenuTreeListBranch2 & "<li><A HREF=""?OpenNodesList=" & OpenNodesList & "&OpenNode=" & NodePointer & """><IMG SRC=""" & Image & """ HEIGHT=""18"" WIDTH=""18"" BORDER=0 ALT=""Open Folder""></A>&nbsp;" & Caption & "</li>"
            Else
                '
                ' The branch is open
                '
                Image = iEntry(NodePointer).ImageOpen
                If Image = "" Then
                    Image = iEntry(NodePointer).Image
                    End If
                GetMenuTreeListBranch2 = GetMenuTreeListBranch2 _
                    & "<li>" _
                    & "<A HREF=""?OpenNodesList=" & OpenNodesList & "&CloseNode=" & NodePointer & """>" _
                    & "<IMG SRC=""" & Image & """ HEIGHT=""18"" WIDTH=""18"" BORDER=0 ALT=""Close Folder"">" _
                    & "</A>&nbsp;" & Caption & "</li>"
                '
                ' Now output any child branches of this node
                '
                UcaseNodeName = UCase(iEntry(NodePointer).Name)
                For EntryPointer = 0 To iEntryCount - 1
                    If (iEntry(EntryPointer).ParentName = UcaseNodeName) Then
                        GetMenuTreeListBranch2 = GetMenuTreeListBranch2 & GetMenuTreeListBranch2(EntryPointer, UsedEntriesList & "," & NodePointer, OpenNodesList)
                        End If
                    Next
                ' GetMenuTreeListBranch2 = GetMenuTreeListBranch2 & GetMenuTreeListBranch2(iEntry(NodePointer).Name, UsedEntriesList & "," & CStr(NodePointer), OpenNodesList)
                End If
            GetMenuTreeListBranch2 = GetMenuTreeListBranch2 & "</ul>" & vbCrLf
            End If
        End If
    '
    Exit Function
    '
ErrorTrap:
    Call HandleClassError("GetMenutreeListBranch2", Err.Number, Err.Source, Err.Description)
End Function
'
'===============================================================================
'   Returns the menu specified, if it is in local storage
'
'   It also creates the menu data in a close string that is returned in GetTreeClose.
'   It must be done there so the link buttons height can be calculated.
'===============================================================================
'
Public Function GetTree(MenuName As String, OpenMenuName As String, Optional StyleSheetPrefix As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Link As String
    Dim EntryPointer As Long
    Dim UcaseMenuName As String
    Dim RootFound As Boolean
    Dim UsedEntries As String
    Dim Caption As String
    Dim JSString As String
    '
    ' ----- Get the menu pointer
    '
    If iEntryCount > 0 Then
        UcaseMenuName = UCase(MenuName)
        If StyleSheetPrefix = "" Then
            StyleSheetPrefix = "ccTree"
        End If
        If True Then
            '
            ' ----- Build the linked -button-
            '
            Link = iEntry(EntryPointer).Link
            If Link = "" Then
                Link = "javascript: ;"
            End If
            '
            '   Find the Menu Entry, and create the top element here
            '
            UsedEntries = ""
            For EntryPointer = 0 To iEntryCount - 1
                With iEntry(EntryPointer)
                    If .Name = UcaseMenuName Then
                        Caption = .Caption
                        If .Link <> "" Then
                            Caption = "<a href=""" & kmaEncodeHTML(.Link) & """>" & Caption & "</a>"
                        End If
                        UsedEntries = UsedEntries & "," & CStr(EntryPointer)
                        GetTree = "" _
                            & vbCrLf & "<ul class=mktree id=tree" & iTreeCount & ">" & vbCrLf _
                            & vbCrLf & " <li id=""" & .Name & """><span class=mkc>" & Caption & "</span>" _
                            & vbCrLf & " <ul>" & vbCrLf _
                            & GetMKTreeBranch(UcaseMenuName, UsedEntries, 2) _
                            & vbCrLf & " </ul>" & vbCrLf _
                            & vbCrLf & "</li></ul>" & vbCrLf
                        Exit For
                    End If
                End With
            Next
            If UsedEntries = "" Then
                GetTree = "" _
                    & vbCrLf & "<ul class=mktree id=tree" & iTreeCount & ">" _
                    & GetMKTreeBranch(UcaseMenuName, UsedEntries, 1) _
                    & vbCrLf & "</ul>" & vbCrLf
            End If
            '
            'Call cmc.main_AddStylesheetLink("/cclib/mktree/mktree.css")
            'Call cmc.main_AddHeadScriptLink("/cclib/mktree/mktree.js", "mktree")
            'Call cmc.main_AddOnLoadJavascript("convertTrees();")
            GetTree = "" _
                & vbCrLf & "<link rel=stylesheet href=/cclib/mktree/mktree.css type=text/css>" _
                & vbCrLf & "<script type=""text/javascript"" src=/cclib/mktree/mktree.js></script>" _
                & GetTree
            GetTree = GetTree & "<script type=""text/javascript"">convertTrees();"
            If OpenMenuName <> "" Then
                JSString = UCase(OpenMenuName)
                JSString = Replace(JSString, "\", "\\")
                JSString = Replace(JSString, vbCrLf, "\n")
                JSString = Replace(JSString, "'", "\'")
                'Call cmc.main_AddOnLoadJavascript("expandToItem('tree" & iTreeCount & "','" & JSString & "');")
                GetTree = GetTree & "expandToItem('tree" & iTreeCount & "','" & JSString & "');"
            End If
            GetTree = GetTree & "</script>"
            '
            ' increment the menu count
            '
            iTreeCount = iTreeCount + 1
            End If
        End If
    
    
    Exit Function
    '
ErrorTrap:
    Call HandleClassError("GetTree", Err.Number, Err.Source, Err.Description)
End Function
'
'===============================================================================
'   Gets the Menu Branch for the Tree Menu
'===============================================================================
'
Private Function GetMKTreeBranch(ParentName As String, UsedEntries As String, Depth As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim EntryPointer As Long
    Dim iUsedEntries As String
    Dim SubMenuCount As Long
    Dim ChildMenu As String
    Dim Caption As String
    '
    iUsedEntries = UsedEntries
    SubMenuCount = 0
    For EntryPointer = 0 To iEntryCount - 1
        With iEntry(EntryPointer)
            If .ParentName = ParentName Then
                If (InStr(1, iUsedEntries & ",", "," & EntryPointer & ",") = 0) Then
                    Caption = .Caption
                    If .OnClick <> "" And .Link <> "" Then
                        Caption = "<a href=""" & kmaEncodeHTML(.Link) & """ onClick=""" & .OnClick & """>" & Caption & "</a>"
                    ElseIf .OnClick <> "" Then
                        Caption = "<a href=""#"" onClick=""" & .OnClick & """>" & Caption & "</a>"
                    ElseIf .Link <> "" Then
                        Caption = "<a href=""" & kmaEncodeHTML(.Link) & """>" & Caption & "</a>"
                    Else
                        Caption = Caption
                    End If
                    iUsedEntries = iUsedEntries & "," & EntryPointer
                    
                    ChildMenu = GetMKTreeBranch(.Name, iUsedEntries, Depth + 1)
                    If newmode Then
                        If ChildMenu = "" Then
                            GetMKTreeBranch = GetMKTreeBranch _
                                & vbCrLf & String(Depth, " ") & "<li class=mklb id=""" & .Name & """ >" _
                                & "<div class=""mkd"">" _
                                & "<span class=mkb>&nbsp;</span>" _
                                & "</div>" _
                                & Caption _
                                & "</li>"
                        Else
                            '
                            ' 3/18/2010 changes to keep firefox from blocking clicks
                            '
                            GetMKTreeBranch = GetMKTreeBranch _
                                & vbCrLf & String(Depth, " ") & "<li class=""mklc"" id=""" & .Name & """ >" _
                                & "<div class=""mkd"" >" _
                                & "<span class=mkb onclick=""mkClick(this)"">&nbsp;</span>" _
                                & "</div>" _
                                & Caption _
                                & vbCrLf & String(Depth + 1, " ") & "<ul>" _
                                & ChildMenu _
                                & vbCrLf & String(Depth + 1, " ") & "</ul>" _
                                & "</li>"
                        End If
                    Else
                        If ChildMenu <> "" Then
                            ChildMenu = "" _
                                & vbCrLf & String(Depth + 1, " ") & "<ul>" _
                                & ChildMenu _
                                & vbCrLf & String(Depth + 1, " ") & "</ul>" _
                                & ""
                        End If
                        GetMKTreeBranch = GetMKTreeBranch _
                            & vbCrLf & String(Depth, " ") & "<li class=mklc id=""" & .Name & """>" _
                            & Caption _
                            & ChildMenu _
                            & vbCrLf & String(Depth, " ") & "</li>"
                    End If
                    SubMenuCount = SubMenuCount + 1
                    End If
                End If
            End With
        Next
    '
    Exit Function
    '
ErrorTrap:
    Call HandleClassError("GetMKTreeBranch", Err.Number, Err.Source, Err.Description)
End Function
'
'========================================================================
'   HandleError
'========================================================================
'
Private Sub HandleClassError(MethodName As String, ErrNumber As Long, ErrSource As String, ErrDescription As String)
    '
    On Error GoTo 0
    Call Err.Raise(ErrNumber, ErrSource, App.EXEName & ".Menu." & MethodName & " encountered and error: " & ErrDescription)
    '
End Sub



