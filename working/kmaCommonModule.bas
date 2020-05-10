Attribute VB_Name = "kmaCommonModule"
 
Option Explicit
'
'========================================================================
'   kma defined errors
'       1000-1999 Contensive
'       2000-2999 Datatree
'
'   see kmaErrorDescription() for transations
'========================================================================
'
Const Error_DataTree_RootNodeNext = 2000
Const Error_DataTree_NoGoNext = 2001
'
'========================================================================
'
'========================================================================
'
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
'========================================================================
'   Declarations for SetPiorityClass
'========================================================================
'
Const THREAD_BASE_PRIORITY_IDLE = -15
Const THREAD_BASE_PRIORITY_LOWRT = 15
Const THREAD_BASE_PRIORITY_MIN = -2
Const THREAD_BASE_PRIORITY_MAX = 2
Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
Const THREAD_PRIORITY_NORMAL = 0
Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Const HIGH_PRIORITY_CLASS = &H80
Const IDLE_PRIORITY_CLASS = &H40
Const NORMAL_PRIORITY_CLASS = &H20
Const REALTIME_PRIORITY_CLASS = &H100
'
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
'

'
'========================================================================
'Converts unsafe characters,
'such as spaces, into their
'corresponding escape sequences.
'========================================================================
'
Declare Function UrlEscape Lib "shlwapi" _
   Alias "UrlEscapeA" _
  (ByVal pszURL As String, _
   ByVal pszEscaped As String, _
   pcchEscaped As Long, _
   ByVal dwFlags As Long) As Long
'
'Converts escape sequences back into
'ordinary characters.
'
Declare Function UrlUnescape Lib "shlwapi" _
   Alias "UrlUnescapeA" _
  (ByVal pszURL As String, _
   ByVal pszUnescaped As String, _
   pcchUnescaped As Long, _
   ByVal dwFlags As Long) As Long

'
'   Error reporting strategy
'       Popups are pop-up boxes that tell the user what to do
'       Logs are files with error details for developers to use debugging
'
'       Attended Programs
'           - errors that do not effect the operation, resume next
'           - all errors trickle up to the user interface level
'           - User Errors, like file not found, return "UserError" code and a description
'           - Internal Errors, like div-by-0, User should see no detail, log gets everything
'           - Dependant Object Errors, codes that return from objects:
'               - If UserError, translate ErrSource for raise, but log all original info
'               - If InternalError, log info and raise InternalError
'               - If you can not tell, call it InternalError
'
'       UnattendedMode
'           The same, except each routine decides when
'
'       When an error happens in-line (bad condition without a raise)
'           Log the error
'           Raise the appropriate Code/Description in the current Source
'
'       When an ErrorTrap occurs
'           If ErrSource is not AppTitle, it is a dependantObjectError, log and translate code
'           If ErrNumber is not an ObjectError, call it internal error, log and translate code
'           Error must be either "InternalError" or "UserError", just raise it again
'
' old - If an error is raised that is not a KmaCode, it is logged and translated
' old - If an error is raised and the soure is not he current App.EXEname, it is logged and translated
'
Public Const KmaErrorBase = vbObjectError                 ' Base on which Internal errors should start
'
'Public Const KmaError_UnderlyingObject = vbObjectError + 1     ' An error occurec in an underlying object
'Public Const KmaccErrorServiceStopped = vbObjectError + 2       ' The service is not running
'Public Const KmaError_BadObject = vbObjectError + 3            ' The Server Pointer is not valid
'Public Const KmaError_UpgradeInProgress = vbObjectError + 4    ' page is blocked because an upgrade is in progress
'Public Const KmaError_InvalidArgument = vbObjectError + 5      ' and input argument is not valid. Put details at end of description
'
Public Const KmaErrorUser = KmaErrorBase + 16                   ' Generic Error code that passes the description back to the user
Public Const KmaErrorInternal = KmaErrorBase + 17               ' Internal error which the user should not see
Public Const KmaErrorPage = KmaErrorBase + 18                   ' Error from the page which called Contensive
'
Public Const KmaObjectError = KmaErrorBase + 256                ' Internal error which the user should not see
'
'==========================================================================
'       NTSvc.ocx, LogEvent Constants
'==========================================================================
'
Public Const NTServiceEventError = 1                    ' Error event.
Public Const NTServiceEventWarning = 2                  ' Warning event.
Public Const NTServiceEventInformation = 4              ' Information event.
Public Const NTServiceEventAuditSuccess = 8             ' Audit success event.
Public Const NTServiceEventAuditFailure = 10            ' Audit failure event.

Public Const NTServiceIDDebug = 108                     ' Debugging message
Public Const NTServiceIDError = 109                     ' Error message
Public Const NTServiceIDInfo = 110                      ' Information message
'
'==========================================================================
'       NTSvc.ocx, LogEvent Constants
'==========================================================================
'
Public Const SQLTrue = "1"
Public Const SQLFalse = "0"
'
'
'
Public Const kmaEndTable = "</table >"
Public Const kmaEndTableCell = "</td>"
Public Const kmaEndTableRow = "</tr>"
'
'==========================================================================
' kmaByteArrayToString / kmaStringToByteArray
'==========================================================================
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'The WideCharToMultiByte function maps a wide-character string to a new character string.
'The function is faster when both lpDefaultChar and lpUsedDefaultChar are NULL.
'CodePage
Private Const CP_ACP = 0 'ANSI
Private Const CP_MACCP = 2 'Mac
Private Const CP_OEMCP = 1 'OEM
Private Const CP_UTF7 = 65000
Private Const CP_UTF8 = 65001
'dwFlags
Private Const WC_NO_BEST_FIT_CHARS = &H400
Private Const WC_COMPOSITECHECK = &H200
Private Const WC_DISCARDNS = &H10
Private Const WC_SEPCHARS = &H20 'Default
Private Const WC_DEFAULTCHAR = &H40

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long
'
'==========================================================================
'   Convert a variant to an long (long)
'   returns 0 if the input is not an integer
'   if float, rounds to integer
'==========================================================================
'
Public Function kmaEncodeInteger(ExpressionVariant As Variant) As Long
    ' 7/14/2009 - cover the overflow case, return 0
    On Error Resume Next
    '
    If Not IsArray(ExpressionVariant) Then
        If Not IsMissing(ExpressionVariant) Then
            If Not IsNull(ExpressionVariant) Then
                If ExpressionVariant <> "" Then
                    If IsNumeric(ExpressionVariant) Then
                        kmaEncodeInteger = CLng(ExpressionVariant)
                    End If
                End If
            End If
        End If
    End If
    '
    Exit Function
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    Err.Clear
    kmaEncodeInteger = 0
End Function
'
'==========================================================================
'   Convert a variant to a number (double)
'   returns 0 if the input is not a number
'==========================================================================
'
Public Function KmaEncodeNumber(ExpressionVariant As Variant) As Double
    On Error GoTo ErrorTrap
    '
    'KmaEncodeNumber = 0
    If Not IsMissing(ExpressionVariant) Then
        If Not IsNull(ExpressionVariant) Then
            If ExpressionVariant <> "" Then
                If IsNumeric(ExpressionVariant) Then
                    KmaEncodeNumber = ExpressionVariant
                    End If
                End If
            End If
        End If
    '
    Exit Function
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    Err.Clear
    KmaEncodeNumber = 0
End Function
'
'==========================================================================
'   Convert a variant to a date
'   returns 0 if the input is not a number
'==========================================================================
'
Public Function KmaEncodeDate(ExpressionVariant As Variant) As Date
    On Error GoTo ErrorTrap
    '
'    KmaEncodeDate = CDate(ExpressionVariant)
'    KmaEncodeDate = CDate("1/1/1980")
    'KmaEncodeDate = CDate(0)
    If Not IsMissing(ExpressionVariant) Then
        If Not IsNull(ExpressionVariant) Then
            If ExpressionVariant <> "" Then
                If IsDate(ExpressionVariant) Then
                    KmaEncodeDate = ExpressionVariant
                    End If
                End If
            End If
        End If
    '
    Exit Function
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    Err.Clear
    KmaEncodeDate = CDate(0)
End Function
'
'==========================================================================
'   Convert a variant to a boolean
'   Returns true if input is not false, else false
'==========================================================================
'
Public Function kmaEncodeBoolean(ExpressionVariant As Variant) As Boolean
    On Error GoTo ErrorTrap
    '
    'KmaEncodeBoolean = False
    If Not IsMissing(ExpressionVariant) Then
        If Not IsNull(ExpressionVariant) Then
            If ExpressionVariant <> "" Then
                If IsNumeric(ExpressionVariant) Then
                    If ExpressionVariant <> "0" Then
                        If ExpressionVariant <> 0 Then
                            kmaEncodeBoolean = True
                            End If
                        End If
                ElseIf UCase(ExpressionVariant) = "ON" Then
                    kmaEncodeBoolean = True
                ElseIf UCase(ExpressionVariant) = "YES" Then
                    kmaEncodeBoolean = True
                ElseIf UCase(ExpressionVariant) = "TRUE" Then
                    kmaEncodeBoolean = True
                Else
                    kmaEncodeBoolean = False
                    End If
                End If
            End If
        End If
    Exit Function
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    Err.Clear
    kmaEncodeBoolean = False
End Function
'
'==========================================================================
'   Convert a variant into 0 or 1
'   Returns 1 if input is not false, else 0
'==========================================================================
'
Public Function KmaEncodeBit(ExpressionVariant As Variant) As Long
    On Error GoTo ErrorTrap
    '
    'KmaEncodeBit = 0
    If kmaEncodeBoolean(ExpressionVariant) Then
        KmaEncodeBit = 1
    End If
    '
    Exit Function
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    Err.Clear
    KmaEncodeBit = 0
End Function
'
'==========================================================================
'   Convert a variant to a string
'   returns emptystring if the input is not a string
'==========================================================================
'
Public Function kmaEncodeText(ExpressionVariant As Variant) As String
    On Error GoTo ErrorTrap
    '
    'KmaEncodeText = ""
    If Not IsMissing(ExpressionVariant) Then
        If Not IsNull(ExpressionVariant) Then
            kmaEncodeText = CStr(ExpressionVariant)
        End If
    End If
    '
    Exit Function
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    Err.Clear
    kmaEncodeText = ""
End Function
'
'==========================================================================
'   Converts a possibly missing value to variant
'==========================================================================
'
Public Function KmaEncodeMissing(ExpressionVariant As Variant, DefaultVariant As Variant) As Variant
    'On Error GoTo ErrorTrap
    '
    If IsMissing(ExpressionVariant) Then
        KmaEncodeMissing = DefaultVariant
    Else
        KmaEncodeMissing = ExpressionVariant
    End If
    '
    Exit Function
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    Err.Clear
End Function
'
'
'
Public Function KmaEncodeMissingText(ExpressionVariant As Variant, DefaultText As String) As String
    KmaEncodeMissingText = kmaEncodeText(KmaEncodeMissing(ExpressionVariant, DefaultText))
End Function
'
'
'
Public Function KmaEncodeMissingInteger(ExpressionVariant As Variant, DefaultInteger As Long) As Long
    KmaEncodeMissingInteger = kmaEncodeInteger(KmaEncodeMissing(ExpressionVariant, DefaultInteger))
End Function
'
'
'
Public Function KmaEncodeMissingDate(ExpressionVariant As Variant, DefaultDate As Date) As Date
    KmaEncodeMissingDate = KmaEncodeDate(KmaEncodeMissing(ExpressionVariant, DefaultDate))
End Function
'
'
'
Public Function KmaEncodeMissingNumber(ExpressionVariant As Variant, DefaultNumber As Double) As Double
    KmaEncodeMissingNumber = KmaEncodeNumber(KmaEncodeMissing(ExpressionVariant, DefaultNumber))
    End Function
'
'
'
Public Function KmaEncodeMissingBoolean(ExpressionVariant As Variant, DefaultState As Boolean) As Boolean
    KmaEncodeMissingBoolean = kmaEncodeBoolean(KmaEncodeMissing(ExpressionVariant, DefaultState))
    End Function
'
'================================================================================================================
'   Separate a URL into its host, path, page parts
'================================================================================================================
'
Public Sub SeparateURL(ByVal SourceURL As String, ByRef Protocol As String, ByRef Host As String, ByRef Path As String, ByRef Page As String, ByRef QueryString As String)
    'On Error GoTo ErrorTrap
    '
    '   Divide the URL into URLHost, URLPath, and URLPage
    '
    Dim WorkingURL As String
    Dim Position As Long
    '
    ' Get Protocol (before the first :)
    '
    WorkingURL = SourceURL
    Position = InStr(1, WorkingURL, ":")
    'Position = InStr(1, WorkingURL, "://")
    If Position <> 0 Then
        Protocol = Mid(WorkingURL, 1, Position + 2)
        WorkingURL = Mid(WorkingURL, Position + 3)
        End If
    '
    ' compatibility fix
    '
    If InStr(1, WorkingURL, "//") = 1 Then
        If Protocol = "" Then
            Protocol = "http:"
            End If
        Protocol = Protocol & "//"
        WorkingURL = Mid(WorkingURL, 3)
        End If
    '
    ' Get QueryString
    '
    Position = InStr(1, WorkingURL, "?")
    If Position > 0 Then
        QueryString = Mid(WorkingURL, Position)
        WorkingURL = Mid(WorkingURL, 1, Position - 1)
        End If
    '
    ' separate host from pathpage
    '
    'iURLHost = WorkingURL
    Position = InStr(WorkingURL, "/")
    If (Position = 0) And (Protocol = "") Then
        '
        ' Page without path or host
        '
        Page = WorkingURL
        Path = ""
        Host = ""
    ElseIf (Position = 0) Then
        '
        ' host, without path or page
        '
        Page = ""
        Path = "/"
        Host = WorkingURL
    Else
        '
        ' host with a path (at least)
        '
        Path = Mid(WorkingURL, Position)
        Host = Mid(WorkingURL, 1, Position - 1)
        '
        ' separate page from path
        '
        Position = InStrRev(Path, "/")
        If Position = 0 Then
            '
            ' no path, just a page
            '
            Page = Path
            Path = "/"
        Else
            Page = Mid(Path, Position + 1)
            Path = Mid(Path, 1, Position)
            End If
        End If
    Exit Sub
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    Err.Clear
    End Sub
'
'================================================================================================================
'   Separate a URL into its host, path, page parts
'================================================================================================================
'
Public Sub ParseURL(ByVal SourceURL As String, ByRef Protocol As String, ByRef Host As String, ByRef Port As String, ByRef Path As String, ByRef Page As String, ByRef QueryString As String)
    'On Error GoTo ErrorTrap
    '
    '   Divide the URL into URLHost, URLPath, and URLPage
    '
    Dim iURLWorking As String               ' internal storage for GetURL functions
    Dim iURLProtocol As String
    Dim iURLHost As String
    Dim iURLPort As String
    Dim iURLPath As String
    Dim iURLPage As String
    Dim iURLQueryString As String
    Dim Position As Long
    '
    iURLWorking = SourceURL
    Position = InStr(1, iURLWorking, "://")
    If Position <> 0 Then
        iURLProtocol = Mid(iURLWorking, 1, Position + 2)
        iURLWorking = Mid(iURLWorking, Position + 3)
        End If
    '
    ' separate Host:Port from pathpage
    '
    iURLHost = iURLWorking
    Position = InStr(iURLHost, "/")
    If Position = 0 Then
        '
        ' just host, no path or page
        '
        iURLPath = "/"
        iURLPage = ""
    Else
        iURLPath = Mid(iURLHost, Position)
        iURLHost = Mid(iURLHost, 1, Position - 1)
        '
        ' separate page from path
        '
        Position = InStrRev(iURLPath, "/")
        If Position = 0 Then
            '
            ' no path, just a page
            '
            iURLPage = iURLPath
            iURLPath = "/"
        Else
            iURLPage = Mid(iURLPath, Position + 1)
            iURLPath = Mid(iURLPath, 1, Position)
            End If
        End If
    '
    ' Divide Host from Port
    '
    Position = InStr(iURLHost, ":")
    If Position = 0 Then
        '
        ' host not given, take a guess
        '
        Select Case UCase(iURLProtocol)
            Case "FTP://"
                iURLPort = "21"
            Case "HTTP://", "HTTPS://"
                iURLPort = "80"
            Case Else
                iURLPort = "80"
            End Select
    Else
        iURLPort = Mid(iURLHost, Position + 1)
        iURLHost = Mid(iURLHost, 1, Position - 1)
        End If
    Position = InStr(1, iURLPage, "?")
    If Position > 0 Then
        iURLQueryString = Mid(iURLPage, Position)
        iURLPage = Mid(iURLPage, 1, Position - 1)
        End If
    Protocol = iURLProtocol
    Host = iURLHost
    Port = iURLPort
    Path = iURLPath
    Page = iURLPage
    QueryString = iURLQueryString
    Exit Sub
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    Err.Clear
    End Sub
'
'
'
Function DecodeGMTDate(GMTDate As String) As Date
    'On Error GoTo ErrorTrap
    '
    Dim WorkString As String
    DecodeGMTDate = 0
    If GMTDate <> "" Then
        WorkString = Mid(GMTDate, 6, 11)
        If IsDate(WorkString) Then
            DecodeGMTDate = CDate(WorkString)
            WorkString = Mid(GMTDate, 18, 8)
            If IsDate(WorkString) Then
                DecodeGMTDate = DecodeGMTDate + CDate(WorkString) + 4 / 24
                End If
            End If
        End If
    Exit Function
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    End Function
'
'
'
Function EncodeGMTDate(MSDate As Date) As String
    'On Error GoTo ErrorTrap
    '
    Dim WorkString As String
    Exit Function
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    End Function
'
'=================================================================================
'   Renamed to catch all the cases that used it in addons
'
'   Do not use this routine in Addons to get the addon option string value
'   to get the value in an option string, use csv.getAddonOption("name")
'
' Get the value of a name in a string of name value pairs parsed with vrlf and =
'   the legacy line delimiter was a '&' -> name1=value1&name2=value2"
'   new format is "name1=value1 crlf name2=value2 crlf ..."
'   There can be no extra spaces between the delimiter, the name and the "="
'=================================================================================
'
Function getSimpleNameValue(Name As String, ArgumentString As String, DefaultValue As String, Delimiter As String) As String
'Function getArgument(Name As String, ArgumentString As String, Optional DefaultValue As Variant, Optional Delimiter As String) As String
    '
    Dim WorkingString As String
    Dim iDefaultValue As String
    Dim NameLength As Long
    Dim ValueStart As Long
    Dim ValueEnd As Long
    Dim IsQuoted As Boolean
    '
    ' determine delimiter
    '
    If Delimiter = "" Then
        '
        ' If not explicit
        '
        If InStr(1, ArgumentString, vbCrLf) <> 0 Then
            '
            ' crlf can only be here if it is the delimiter
            '
            Delimiter = vbCrLf
        Else
            '
            ' either only one option, or it is the legacy '&' delimit
            '
            Delimiter = "&"
        End If
    End If
    iDefaultValue = KmaEncodeMissing(DefaultValue, "")
    WorkingString = ArgumentString
    getSimpleNameValue = iDefaultValue
    If WorkingString <> "" Then
        WorkingString = Delimiter & WorkingString & Delimiter
        ValueStart = InStr(1, WorkingString, Delimiter & Name & "=", vbTextCompare)
        If ValueStart <> 0 Then
            NameLength = Len(Name)
            ValueStart = ValueStart + Len(Delimiter) + NameLength + 1
            If Mid(WorkingString, ValueStart, 1) = """" Then
                IsQuoted = True
                ValueStart = ValueStart + 1
            End If
            If IsQuoted Then
                ValueEnd = InStr(ValueStart, WorkingString, """" & Delimiter)
            Else
                ValueEnd = InStr(ValueStart, WorkingString, Delimiter)
            End If
            If ValueEnd = 0 Then
                getSimpleNameValue = Mid(WorkingString, ValueStart)
            Else
                getSimpleNameValue = Mid(WorkingString, ValueStart, ValueEnd - ValueStart)
            End If
        End If
    End If
    '
    
    Exit Function
    '
    ' ----- ErrorTrap
    '
ErrorTrap:
    End Function
'
'=================================================================================
'   Do not use this code
'
'   To retrieve a value from an option string, use csv.getAddonOption("name")
'
'   This was left here to work through any code issues that might arrise during
'   the conversion.
'
'   Return the value from a name value pair, parsed with =,&[|].
'   For example:
'       name=Jay[Jay|Josh|Dwayne]
'       the answer is Jay. If a select box is displayed, it is a dropdown of all three
'=================================================================================
'
Public Function GetAggrOption_old(Name As String, SegmentCMDArgs As String) As String
    '
    Dim Pos As Long
    '
    GetAggrOption_old = getSimpleNameValue(Name, SegmentCMDArgs, "", vbCrLf)
    '
    ' remove the manual select list syntax "answer[choice1|choice2]"
    '
    Pos = InStr(1, GetAggrOption_old, "[")
    If Pos <> 0 Then
        GetAggrOption_old = Left(GetAggrOption_old, Pos - 1)
    End If
    '
    ' remove any function syntax "answer{selectcontentname RSS Feeds}"
    '
    Pos = InStr(1, GetAggrOption_old, "{")
    If Pos <> 0 Then
        GetAggrOption_old = Left(GetAggrOption_old, Pos - 1)
    End If
    '
End Function
'
'=================================================================================
'   Do not use this code
'
'   To retrieve a value from an option string, use csv.getAddonOption("name")
'
'   This was left here to work through any code issues that might arrise during
'   Compatibility for GetArgument
'=================================================================================
'
Function kmaGetNameValue_old(Name As String, ArgumentString As String, Optional DefaultValue As String) As String
    kmaGetNameValue_old = getSimpleNameValue(Name, ArgumentString, DefaultValue, vbCrLf)
End Function
'
'========================================================================
'   KmaEncodeSQLText
'========================================================================
'
Public Function KmaEncodeSQLText(ExpressionVariant As Variant) As String
    'On Error GoTo ErrorTrap
    '
    'Dim MethodName As String
    '
    'MethodName = "KmaEncodeSQLText"
    '
    If IsNull(ExpressionVariant) Then
        KmaEncodeSQLText = "null"
    ElseIf IsMissing(ExpressionVariant) Then
        KmaEncodeSQLText = "null"
    ElseIf ExpressionVariant = "" Then
        KmaEncodeSQLText = "null"
    Else
        KmaEncodeSQLText = CStr(ExpressionVariant)
        ' ??? this should not be here -- to correct a field used in a CDef, truncate in SaveCS by fieldtype
        'KmaEncodeSQLText = Left(ExpressionVariant, 255)
        'remove-can not find a case where | is not allowed to be saved.
        'KmaEncodeSQLText = Replace(KmaEncodeSQLText, "|", "_")
        KmaEncodeSQLText = "'" & Replace(KmaEncodeSQLText, "'", "''") & "'"
        End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    End Function
'
'========================================================================
'   KmaEncodeSQLLongText
'========================================================================
'
Public Function KmaEncodeSQLLongText(ExpressionVariant As Variant) As String
    'On Error GoTo ErrorTrap
    '
    'Dim MethodName As String
    '
    'MethodName = "KmaEncodeSQLLongText"
    '
    If IsNull(ExpressionVariant) Then
        KmaEncodeSQLLongText = "null"
    ElseIf IsMissing(ExpressionVariant) Then
        KmaEncodeSQLLongText = "null"
    ElseIf ExpressionVariant = "" Then
        KmaEncodeSQLLongText = "null"
    Else
        KmaEncodeSQLLongText = ExpressionVariant
        'KmaEncodeSQLLongText = Replace(ExpressionVariant, "|", "_")
        KmaEncodeSQLLongText = "'" & Replace(KmaEncodeSQLLongText, "'", "''") & "'"
        End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    End Function
'
'========================================================================
'   KmaEncodeSQLDate
'       encode a date variable to go in an sql expression
'========================================================================
'
Public Function KmaEncodeSQLDate(ExpressionVariant As Variant) As String
    'On Error GoTo ErrorTrap
    '
    Dim TimeVar As Date
    Dim TimeValuething As Single
    Dim TimeHours As Long
    Dim TimeMinutes As Long
    Dim TimeSeconds As Long
    'Dim MethodName As String
    ''
    'MethodName = "KmaEncodeSQLDate"
    '
    If IsNull(ExpressionVariant) Then
        KmaEncodeSQLDate = "null"
    ElseIf IsMissing(ExpressionVariant) Then
        KmaEncodeSQLDate = "null"
    ElseIf ExpressionVariant = "" Then
        KmaEncodeSQLDate = "null"
    ElseIf IsDate(ExpressionVariant) Then
        TimeVar = CDate(ExpressionVariant)
        If TimeVar = 0 Then
            KmaEncodeSQLDate = "null"
        Else
            TimeValuething = 86400! * (TimeVar - Int(TimeVar + 0.000011!))
            TimeHours = Int(TimeValuething / 3600!)
            If TimeHours >= 24 Then
                TimeHours = 23
            End If
            TimeMinutes = Int(TimeValuething / 60!) - (TimeHours * 60)
            If TimeMinutes >= 60 Then
                TimeMinutes = 59
            End If
            TimeSeconds = TimeValuething - (TimeHours * 3600!) - (TimeMinutes * 60!)
            If TimeSeconds >= 60 Then
                TimeSeconds = 59
            End If
            KmaEncodeSQLDate = "{ts '" & Year(ExpressionVariant) & "-" & Right("0" & Month(ExpressionVariant), 2) & "-" & Right("0" & Day(ExpressionVariant), 2) & " " & Right("0" & TimeHours, 2) & ":" & Right("0" & TimeMinutes, 2) & ":" & Right("0" & TimeSeconds, 2) & "'}"
            End If
    Else
        KmaEncodeSQLDate = "null"
        End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    End Function
'
'========================================================================
'   KmaEncodeSQLNumber
'       encode a number variable to go in an sql expression
'========================================================================
'
Function KmaEncodeSQLNumber(ExpressionVariant As Variant) As String
    'On Error GoTo ErrorTrap
    '
    'Dim MethodName As String
    ''
    'MethodName = "KmaEncodeSQLNumber"
    '
    If IsNull(ExpressionVariant) Then
        KmaEncodeSQLNumber = "null"
    ElseIf IsMissing(ExpressionVariant) Then
        KmaEncodeSQLNumber = "null"
    ElseIf ExpressionVariant = "" Then
        KmaEncodeSQLNumber = "null"
    ElseIf IsNumeric(ExpressionVariant) Then
        Select Case VarType(ExpressionVariant)
            Case vbBoolean
                If ExpressionVariant Then
                    KmaEncodeSQLNumber = SQLTrue
                Else
                    KmaEncodeSQLNumber = SQLFalse
                    End If
            Case Else
                KmaEncodeSQLNumber = ExpressionVariant
            End Select
    Else
        KmaEncodeSQLNumber = "null"
        End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    End Function
'
'========================================================================
'   KmaEncodeSQLBoolean
'       encode a boolean variable to go in an sql expression
'========================================================================
'
Public Function KmaEncodeSQLBoolean(ExpressionVariant As Variant) As String
    '
    Dim src As String
    '
    KmaEncodeSQLBoolean = SQLFalse
    If kmaEncodeBoolean(ExpressionVariant) Then
        KmaEncodeSQLBoolean = SQLTrue
    End If
'    If Not IsNull(ExpressionVariant) Then
'        If Not IsMissing(ExpressionVariant) Then
'            If ExpressionVariant <> False Then
'                    KmaEncodeSQLBoolean = SQLTrue
'                End If
'            End If
'        End If
'    End If
    '
    End Function
'
'========================================================================
'   Gets the next line from a string, and removes the line
'========================================================================
'
Public Function KmaGetLine(Body As String) As String
    Dim EOL As String
    Dim NextCR As Long
    Dim NextLF As Long
    Dim BOL As Long
    '
    NextCR = InStr(1, Body, vbCr)
    NextLF = InStr(1, Body, vbLf)
    
    If NextCR <> 0 Or NextLF <> 0 Then
        If NextCR <> 0 Then
            If NextLF <> 0 Then
                If NextCR < NextLF Then
                    EOL = NextCR - 1
                    If NextLF = NextCR + 1 Then
                        BOL = NextLF + 1
                    Else
                        BOL = NextCR + 1
                        End If
                    
                Else
                    EOL = NextLF - 1
                    BOL = NextLF + 1
                    End If
            Else
                EOL = NextCR - 1
                BOL = NextCR + 1
                End If
        Else
            EOL = NextLF - 1
            BOL = NextLF + 1
            End If
        KmaGetLine = Mid(Body, 1, EOL)
        Body = Mid(Body, BOL)
    Else
        KmaGetLine = Body
        Body = ""
        End If
    
    'EOL = InStr(1, Body, vbCrLf)
    
    'If EOL <> 0 Then
    '    KmaGetLine = Mid(Body, 1, EOL - 1)
    '    Body = Mid(Body, EOL + 2)
    '    End If
    '
    End Function
'
'=================================================================================
'   Get a Random Long Value
'=================================================================================
'
Public Function GetRandomInteger() As Long
    '
    Dim RandomBase As Long
    Dim RandomLimit As Long
    '
    RandomBase = App.ThreadID
    RandomBase = RandomBase And ((2 ^ 30) - 1)
    RandomLimit = (2 ^ 31) - RandomBase - 1
    Randomize
    GetRandomInteger = RandomBase + (Rnd * RandomLimit)
    '
    End Function
'
'=================================================================================
'
'=================================================================================
'
Public Function IsRSOK(RS As Object) As Boolean
    IsRSOK = False
    If Not (RS Is Nothing) Then
        If RS.State <> 0 Then
            If Not RS.EOF Then
                IsRSOK = True
                End If
            End If
        End If
    End Function
'
'=================================================================================
'
'=================================================================================
'
Public Sub CloseRS(RS As Object)
    If Not (RS Is Nothing) Then
        If RS.State <> 0 Then
            Call RS.Close
            End If
        End If
    End Sub
'
'=============================================================================
' Create the part of the sql where clause that is modified by the user
'   WorkingQuery is the original querystring to change
'   QueryName is the name part of the name pair to change
'   If the QueryName is not found in the string, ignore call
'=============================================================================
'
Public Function ModifyQueryString(WorkingQuery As String, QueryName As String, QueryValue As String, Optional AddIfMissing As Variant) As String
    '
    If InStr(1, WorkingQuery, "?") Then
        ModifyQueryString = kmaModifyLinkQuery(WorkingQuery, QueryName, QueryValue, AddIfMissing)
    Else
        ModifyQueryString = Mid(kmaModifyLinkQuery("?" & WorkingQuery, QueryName, QueryValue, AddIfMissing), 2)
        End If
    End Function
'
'=============================================================================
'   Modify a querystring name/value pair in a Link
'=============================================================================
'
Public Function kmaModifyLinkQuery(Link As String, QueryName As String, QueryValue As String, Optional AddIfMissing As Variant) As String
    '
    Dim PositionName As Long
    Dim PositionEqual As Long
    Dim PositionValueStart As Long
    Dim PositionValueEnd As Long
    Dim Element() As String
    Dim ElementCount As Long
    Dim ElementPointer As Long
    Dim NameValue() As String
    Dim UcaseQueryName As String
    Dim ElementFound As Boolean
    Dim iAddIfMissing As Boolean
    Dim LeftPart As String
    Dim QueryString As String
    '
    iAddIfMissing = KmaEncodeMissingBoolean(AddIfMissing, True)
    If InStr(1, Link, "?") <> 0 Then
        kmaModifyLinkQuery = Mid(Link, 1, InStr(1, Link, "?") - 1)
        QueryString = Mid(Link, Len(kmaModifyLinkQuery) + 2)
    Else
        kmaModifyLinkQuery = Link
        QueryString = ""
        End If
    UcaseQueryName = UCase(kmaEncodeRequestVariable(QueryName))
    If QueryString <> "" Then
        Element = Split(QueryString, "&")
        ElementCount = UBound(Element) + 1
        For ElementPointer = 0 To ElementCount - 1
            NameValue = Split(Element(ElementPointer), "=")
            If UBound(NameValue) = 1 Then
                If UCase(NameValue(0)) = UcaseQueryName Then
                    If QueryValue = "" Then
                        Element(ElementPointer) = ""
                    Else
                        Element(ElementPointer) = QueryName & "=" & QueryValue
                        End If
                    ElementFound = True
                    Exit For
                    End If
                End If
            Next
        End If
    If Not ElementFound And (QueryValue <> "") Then
        '
        ' element not found, it needs to be added
        '
        If iAddIfMissing Then
            If QueryString = "" Then
                QueryString = kmaEncodeRequestVariable(QueryName) & "=" & kmaEncodeRequestVariable(QueryValue)
            Else
                QueryString = QueryString & "&" & kmaEncodeRequestVariable(QueryName) & "=" & kmaEncodeRequestVariable(QueryValue)
                End If
            End If
    Else
        '
        ' element found
        '
        QueryString = Join(Element, "&")
        If (QueryString <> "") And (QueryValue = "") Then
            '
            ' element found and needs to be removed
            '
            QueryString = Replace(QueryString, "&&", "&")
            If Left(QueryString, 1) = "&" Then
                QueryString = Mid(QueryString, 2)
                End If
            If Right(QueryString, 1) = "&" Then
                QueryString = Mid(QueryString, 1, Len(QueryString) - 1)
                End If
            End If
        End If
    If (QueryString <> "") Then
        kmaModifyLinkQuery = kmaModifyLinkQuery & "?" & QueryString
        End If
    End Function
'
'=================================================================================
'
'=================================================================================
'
Public Function GetIntegerString(Value As Long, DigitCount As Long) As String
    If Len(Value) <= DigitCount Then
        GetIntegerString = String(DigitCount - Len(CStr(Value)), "0") & CStr(Value)
    Else
        GetIntegerString = CStr(Value)
        End If
    End Function
'
'==========================================================================================
'   Set the current process to a high priority
'       Should be called once from the objects parent when it is first created.
'
'   taken from an example labeled
'       KPD-Team 2000
'       URL: http://www.allapi.net/
'       Email: KPDTeam@Allapi.net
'==========================================================================================
'
Public Sub SetProcessHighPriority()
    Dim hProcess As Long
    '
    'set the new priority class
    '
    hProcess = GetCurrentProcess
    Call SetPriorityClass(hProcess, HIGH_PRIORITY_CLASS)
    '
    End Sub
'
'==========================================================================================
'   Format the current error object into a standard string
'==========================================================================================
'
Public Function GetErrString(Optional ErrorObject As Object) As String
    Dim Copy As String
    If ErrorObject Is Nothing Then
        If Err.Number = 0 Then
            GetErrString = "[no error]"
        Else
            Copy = Err.Description
            Copy = Replace(Copy, vbCrLf, "-")
            Copy = Replace(Copy, vbLf, "-")
            Copy = Replace(Copy, vbCrLf, "")
            GetErrString = "[" & Err.Source & " #" & Err.Number & ", " & Copy & "]"
        End If
    Else
        If ErrorObject.Number = 0 Then
            GetErrString = "[no error]"
        Else
            Copy = ErrorObject.Description
            Copy = Replace(Copy, vbCrLf, "-")
            Copy = Replace(Copy, vbLf, "-")
            Copy = Replace(Copy, vbCrLf, "")
            GetErrString = "[" & ErrorObject.Source & " #" & ErrorObject.Number & ", " & Copy & "]"
        End If
    End If
    '
    End Function
'
'==========================================================================================
'   Format the current error object into a standard string
'==========================================================================================
'
Public Function GetProcessID() As Long
    GetProcessID = GetCurrentProcessId
    End Function
'
'==========================================================================================
'   Test if a test string is in a delimited string
'==========================================================================================
'
Public Function IsInDelimitedString(DelimitedString As String, TestString As String, Delimiter As String) As Boolean
    IsInDelimitedString = (0 <> InStr(1, Delimiter & DelimitedString & Delimiter, Delimiter & TestString & Delimiter, vbTextCompare))
    End Function
'
'========================================================================
' kmaEncodeURL
'
'   Encodes only what is to the left of the first ?
'   All URL path characters are assumed to be correct (/:#)
'========================================================================
'
Function kmaEncodeURL(Source As String) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim URLSplit() As String
    Dim LeftSide As String
    Dim RightSide As String
    '
    kmaEncodeURL = Source
    If Source <> "" Then
        URLSplit = Split(Source, "?")
        kmaEncodeURL = URLSplit(0)
        kmaEncodeURL = Replace(kmaEncodeURL, "%", "%25")
        '
        kmaEncodeURL = Replace(kmaEncodeURL, """", "%22")
        kmaEncodeURL = Replace(kmaEncodeURL, " ", "%20")
        kmaEncodeURL = Replace(kmaEncodeURL, "$", "%24")
        kmaEncodeURL = Replace(kmaEncodeURL, "+", "%2B")
        kmaEncodeURL = Replace(kmaEncodeURL, ",", "%2C")
        kmaEncodeURL = Replace(kmaEncodeURL, ";", "%3B")
        kmaEncodeURL = Replace(kmaEncodeURL, "<", "%3C")
        kmaEncodeURL = Replace(kmaEncodeURL, "=", "%3D")
        kmaEncodeURL = Replace(kmaEncodeURL, ">", "%3E")
        kmaEncodeURL = Replace(kmaEncodeURL, "@", "%40")
        If UBound(URLSplit) > 0 Then
            kmaEncodeURL = kmaEncodeURL & "?" & kmaEncodeQueryString(URLSplit(1))
            End If
        End If
    '
    End Function
'
'========================================================================
' kmaEncodeQueryString
'
'   This routine encodes the URL QueryString to conform to rules
'========================================================================
'
Function kmaEncodeQueryString(Source As String) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim QSSplit() As String
    Dim QSPointer As Long
    Dim NVSplit() As String
    Dim NV As String
    '
    kmaEncodeQueryString = ""
    If Source <> "" Then
        QSSplit = Split(Source, "&")
        For QSPointer = 0 To UBound(QSSplit)
            NV = QSSplit(QSPointer)
            If NV <> "" Then
                NVSplit = Split(NV, "=")
                If UBound(NVSplit) = 0 Then
                    NVSplit(0) = kmaEncodeRequestVariable(NVSplit(0))
                    kmaEncodeQueryString = kmaEncodeQueryString & "&" & NVSplit(0)
                Else
                    NVSplit(0) = kmaEncodeRequestVariable(NVSplit(0))
                    NVSplit(1) = kmaEncodeRequestVariable(NVSplit(1))
                    kmaEncodeQueryString = kmaEncodeQueryString & "&" & NVSplit(0) & "=" & NVSplit(1)
                End If
            End If
        Next
        If kmaEncodeQueryString <> "" Then
            kmaEncodeQueryString = Mid(kmaEncodeQueryString, 2)
        End If
    End If
    '
End Function
'
'========================================================================
' kmaEncodeRequestVariable
'
'   This routine encodes a request variable for a URL Query String
'       ...can be the requestname or the requestvalue
'========================================================================
'
Function kmaEncodeRequestVariable(Source As String) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim SourcePointer As Long
    Dim Character As String
    Dim LocalSource As String
    '
    If Source <> "" Then
        LocalSource = Source
        ' "+" is an allowed character for filenames. If you add it, the wrong file will be looked up
        'LocalSource = Replace(LocalSource, " ", "+")
        For SourcePointer = 1 To Len(LocalSource)
            Character = Mid(LocalSource, SourcePointer, 1)
            ' "%" added so if this is called twice, it will not destroy "%20" values
            'If Character = " " Then
            '    kmaEncodeRequestVariable = kmaEncodeRequestVariable & "+"
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-_!*()", Character, vbTextCompare) <> 0 Then
             'ElseIf InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789./:-_!*()", Character, vbTextCompare) <> 0 Then
            'ElseIf InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789./:?#-_!~*'()%", Character, vbTextCompare) <> 0 Then
                kmaEncodeRequestVariable = kmaEncodeRequestVariable & Character
            Else
                kmaEncodeRequestVariable = kmaEncodeRequestVariable & "%" & Hex(Asc(Character))
                End If
            Next
        End If
    '
    End Function
'
'========================================================================
' kmaEncodeHTML
'
'   Convert all characters that are not allowed in HTML to their Text equivalent
'   in preperation for use on an HTML page
'========================================================================
'
Function kmaEncodeHTML(Source As String) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    kmaEncodeHTML = Source
    kmaEncodeHTML = Replace(kmaEncodeHTML, "&", "&amp;")
    kmaEncodeHTML = Replace(kmaEncodeHTML, "<", "&lt;")
    kmaEncodeHTML = Replace(kmaEncodeHTML, ">", "&gt;")
    kmaEncodeHTML = Replace(kmaEncodeHTML, """", "&quot;")
    kmaEncodeHTML = Replace(kmaEncodeHTML, "'", "&#39;")
    'kmaEncodeHTML = Replace(kmaEncodeHTML, "'", "&apos;")
    '
    End Function
'
'========================================================================
' kmaDecodeHTML
'
'   Convert HTML equivalent characters to their equivalents
'========================================================================
'
Function kmaDecodeHTML(Source As String) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim Pos As Long
    Dim s As String
    Dim CharCodeString As String
    Dim CharCode As Long
    Dim posEnd As Long
    '
    ' 11/26/2009 - basically re-wrote it, I commented the old one out below
    '
    kmaDecodeHTML = ""
    If Source <> "" Then
        s = Source
        '
        Pos = Len(s)
        Pos = InStrRev(s, "&#", Pos)
        Do While Pos <> 0
            CharCodeString = ""
            If Mid(s, Pos + 3, 1) = ";" Then
                CharCodeString = Mid(s, Pos + 2, 1)
                posEnd = Pos + 4
            ElseIf Mid(s, Pos + 4, 1) = ";" Then
                CharCodeString = Mid(s, Pos + 2, 2)
                posEnd = Pos + 5
            ElseIf Mid(s, Pos + 5, 1) = ";" Then
                CharCodeString = Mid(s, Pos + 2, 3)
                posEnd = Pos + 6
            End If
            If CharCodeString <> "" Then
                If IsNumeric(CharCodeString) Then
                    CharCode = CLng(CharCodeString)
                    s = Mid(s, 1, Pos - 1) & Chr(CharCode) & Mid(s, posEnd)
                End If
            End If
            '
            Pos = InStrRev(s, "&#", Pos)
        Loop
        '
        ' replace out all common names (at least the most common for now)
        '
        s = Replace(s, "&lt;", "<")
        s = Replace(s, "&gt;", ">")
        s = Replace(s, "&quot;", """")
        s = Replace(s, "&apos;", "'")
        '
        ' Always replace the amp last
        '
        s = Replace(s, "&amp;", "&")
        '
        kmaDecodeHTML = s
    End If
    ' pre-11/26/2009
    'kmaDecodeHTML = Source
    'kmaDecodeHTML = Replace(kmaDecodeHTML, "&amp;", "&")
    'kmaDecodeHTML = Replace(kmaDecodeHTML, "&lt;", "<")
    'kmaDecodeHTML = Replace(kmaDecodeHTML, "&gt;", ">")
    'kmaDecodeHTML = Replace(kmaDecodeHTML, "&quot;", """")
    'kmaDecodeHTML = Replace(kmaDecodeHTML, "&nbsp;", " ")
    '
End Function
'
'========================================================================
' kmaAddSpanClass
'
'   Adds a span around the copy with the class name provided
'========================================================================
'
Function kmaAddSpan(Copy As String, ClassName As String) As String
    '
    kmaAddSpan = "<SPAN Class=""" & ClassName & """>" & Copy & "</SPAN>"
    '
End Function
'
'========================================================================
' kmaDecodeResponseVariable
'
'   Converts a querystring name or value back into the characters it represents
'   This is the same code as the decodeurl
'========================================================================
'
Function kmaDecodeResponseVariable(Source As String) As String
    '
    Dim Position As Long
    Dim ESCString As String
    Dim ESCValue As Long
    Dim Digit0 As String
    Dim Digit1 As String
    Dim iURL As String
    '
    iURL = kmaEncodeText(Source)
    kmaDecodeResponseVariable = Replace(iURL, "+", " ")
    Position = InStr(1, kmaDecodeResponseVariable, "%")
    Do While Position <> 0
        ESCString = Mid(kmaDecodeResponseVariable, Position, 3)
        Digit0 = UCase(Mid(ESCString, 2, 1))
        Digit1 = UCase(Mid(ESCString, 3, 1))
        If ((Digit0 >= "0") And (Digit0 <= "9")) Or ((Digit0 >= "A") And (Digit0 <= "F")) Then
            If ((Digit1 >= "0") And (Digit1 <= "9")) Or ((Digit1 >= "A") And (Digit1 <= "F")) Then
                ESCValue = CLng("&H" & Mid(ESCString, 2))
                kmaDecodeResponseVariable = Mid(kmaDecodeResponseVariable, 1, Position - 1) & Chr(ESCValue) & Mid(kmaDecodeResponseVariable, Position + 3)
                '  & Replace(kmaDecodeResponseVariable, ESCString, Chr(ESCValue), Position, 1)
            End If
        End If
        Position = InStr(Position + 1, kmaDecodeResponseVariable, "%")
    Loop
    '
End Function
'
'========================================================================
' kmadecodeURL
'   Converts a querystring from an Encoded URL (with %20 and +), to non incoded (with spaced)
'========================================================================
'
Function kmaDecodeURL(Source As String) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim Position As Long
    Dim ESCString As String
    Dim ESCValue As Long
    Dim Digit0 As String
    Dim Digit1 As String
    Dim iURL As String
    '
    iURL = kmaEncodeText(Source)
    kmaDecodeURL = Replace(iURL, "+", " ")
    Position = InStr(1, kmaDecodeURL, "%")
    Do While Position <> 0
        ESCString = Mid(kmaDecodeURL, Position, 3)
        Digit0 = UCase(Mid(ESCString, 2, 1))
        Digit1 = UCase(Mid(ESCString, 3, 1))
        If ((Digit0 >= "0") And (Digit0 <= "9")) Or ((Digit0 >= "A") And (Digit0 <= "F")) Then
            If ((Digit1 >= "0") And (Digit1 <= "9")) Or ((Digit1 >= "A") And (Digit1 <= "F")) Then
                ESCValue = CLng("&H" & Mid(ESCString, 2))
                kmaDecodeURL = Replace(kmaDecodeURL, ESCString, Chr(ESCValue))
            End If
        End If
        Position = InStr(Position + 1, kmaDecodeURL, "%")
    Loop
    '
End Function
'
'========================================================================
' kmaGetFirstNonZeroDate
'
'   Converts a querystring name or value back into the characters it represents
'========================================================================
'
Function kmaGetFirstNonZeroDate(Date0 As Date, Date1 As Date) As Date
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim NullDate As Date
    '
    NullDate = CDate(0)
    If Date0 = NullDate Then
        If Date1 = NullDate Then
            '
            ' Both 0, return 0
            '
            kmaGetFirstNonZeroDate = NullDate
        Else
            '
            ' Date0 is NullDate, return Date1
            '
            kmaGetFirstNonZeroDate = Date1
        End If
    Else
        If Date1 = NullDate Then
            '
            ' Date1 is nulldate, return Date0
            '
            kmaGetFirstNonZeroDate = Date0
        ElseIf Date0 < Date1 Then
            '
            ' Date0 is first
            '
            kmaGetFirstNonZeroDate = Date0
        Else
            '
            ' Date1 is first
            '
            kmaGetFirstNonZeroDate = Date1
        End If
    End If
    '
End Function
'
'========================================================================
' kmaGetFirstposition
'
'   returns 0 if both are zero
'   returns 1 if the first integer is non-zero and less then the second
'   returns 2 if the second integer is non-zero and less then the first
'========================================================================
'
Function kmaGetFirstNonZeroLong(Integer1 As Long, Integer2 As Long) As Long
    ' ##### removed to catch err<>0 problem on error resume next
    '
    If Integer1 = 0 Then
        If Integer2 = 0 Then
            '
            ' Both 0, return 0
            '
            kmaGetFirstNonZeroLong = 0
        Else
            '
            ' Integer1 is 0, return Integer2
            '
            kmaGetFirstNonZeroLong = 2
        End If
    Else
        If Integer2 = 0 Then
            '
            ' Integer2 is 0, return Integer1
            '
            kmaGetFirstNonZeroLong = 1
        ElseIf Integer1 < Integer2 Then
            '
            ' Integer1 is first
            '
            kmaGetFirstNonZeroLong = 1
        Else
            '
            ' Integer2 is first
            '
            kmaGetFirstNonZeroLong = 2
        End If
    End If
    '
End Function
'
'========================================================================
' kmaSplit
'   returns the result of a Split, except it honors quoted text
'   if a quote is found, it is assumed to also be a delimiter ( 'this"that"theother' = 'this "that" theother' )
'========================================================================
'
Function kmaSplit(WordList As String, Delimiter As String) As Variant
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim QuoteSplit() As String
    Dim QuoteSplitCount As Long
    Dim QuoteSplitPointer As Long
    Dim InQuote As Boolean
    Dim Out() As String
    Dim OutPointer As Long
    Dim OutSize As Long
    Dim SpaceSplit() As String
    Dim SpaceSplitCount As Long
    Dim SpaceSplitPointer As Long
    Dim Fragment As String
    '
    OutPointer = 0
    ReDim Out(0)
    OutSize = 1
    If WordList <> "" Then
        QuoteSplit = Split(WordList, """")
        QuoteSplitCount = UBound(QuoteSplit) + 1
        InQuote = (Mid(WordList, 1, 1) = "")
        For QuoteSplitPointer = 0 To QuoteSplitCount - 1
            Fragment = QuoteSplit(QuoteSplitPointer)
            If Fragment = "" Then
                '
                ' empty fragment
                ' this is a quote at the end, or two quotes together
                ' do not skip to the next out pointer
                '
                        If OutPointer >= OutSize Then
                            OutSize = OutSize + 10
                            ReDim Preserve Out(OutSize)
                            End If
                'OutPointer = OutPointer + 1
            Else
                If Not InQuote Then
                    SpaceSplit = Split(Fragment, Delimiter)
                    SpaceSplitCount = UBound(SpaceSplit) + 1
                    For SpaceSplitPointer = 0 To SpaceSplitCount - 1
                        If OutPointer >= OutSize Then
                            OutSize = OutSize + 10
                            ReDim Preserve Out(OutSize)
                            End If
                        Out(OutPointer) = Out(OutPointer) & SpaceSplit(SpaceSplitPointer)
                        If (SpaceSplitPointer <> (SpaceSplitCount - 1)) Then
                            '
                            ' divide output between splits
                            '
                            OutPointer = OutPointer + 1
                            If OutPointer >= OutSize Then
                                OutSize = OutSize + 10
                            ReDim Preserve Out(OutSize)
                                End If
                            End If
                        Next
                Else
                    Out(OutPointer) = Out(OutPointer) & """" & Fragment & """"
                    End If
                End If
            InQuote = Not InQuote
            Next
        End If
    ReDim Preserve Out(OutPointer)
    '
    '
    kmaSplit = Out
    '
    End Function
'
'========================================================================
' kmaSplit_Old
'   returns the result of a Split, except it honors quoted text
'   if a quote is found, it is assumed to also be a delimiter ( 'this"that"theother' = 'this "that" theother' )
'========================================================================
'
Function kmaSplit_Old(WordList As String, Delimiter As String) As Variant
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim CurrentPosition As Long
    Dim NextDelimiterPosition As Long
    Dim NextQuotePosition As Long
    Dim ResultCount As Long
    Dim Result() As String
    Dim WorkingWordList As String
    Dim LenWorkingWordList As Long
    '
    WorkingWordList = Trim(WordList)
    If WorkingWordList <> "" Then
        LenWorkingWordList = Len(WorkingWordList)
        CurrentPosition = 1
        ResultCount = 0
        Do While CurrentPosition <> 0
            NextDelimiterPosition = InStr(CurrentPosition, WorkingWordList, Delimiter, vbTextCompare)
            NextQuotePosition = InStr(CurrentPosition, WorkingWordList, """", vbTextCompare)
            Select Case kmaGetFirstNonZeroLong(NextDelimiterPosition, NextQuotePosition)
                Case 0
                    '
                    ' no more left
                    '
                    If CurrentPosition <= LenWorkingWordList Then
                    ReDim Preserve Result(ResultCount)
                        Result(ResultCount) = Mid(WorkingWordList, CurrentPosition)
                        ResultCount = ResultCount + 1
                        End If
                    CurrentPosition = 0
                Case 1
                    '
                    ' Delimiter found before quote
                    '
                    ReDim Preserve Result(ResultCount)
                    Result(ResultCount) = Mid(WorkingWordList, CurrentPosition, NextDelimiterPosition - CurrentPosition)
                    ResultCount = ResultCount + 1
                    If NextDelimiterPosition >= LenWorkingWordList Then
                        CurrentPosition = 0
                    Else
                        CurrentPosition = NextDelimiterPosition + 1
                        End If
                Case 2
                    '
                    ' Quote Found before delimiter
                    '
                    CurrentPosition = NextQuotePosition + 1
                    NextQuotePosition = InStr(CurrentPosition, WorkingWordList, """", vbTextCompare)
                    If NextQuotePosition = 0 Then
                        '
                        ' Problem, as single quote. Just end the phrase here
                        '
                        NextQuotePosition = LenWorkingWordList + 1
                        End If
                    ReDim Preserve Result(ResultCount)
                    Result(ResultCount) = Mid(WorkingWordList, CurrentPosition, NextQuotePosition - CurrentPosition)
                    ResultCount = ResultCount + 1
                    If NextQuotePosition >= LenWorkingWordList Then
                        CurrentPosition = 0
                    Else
                        CurrentPosition = NextQuotePosition + 1
                        End If
                End Select
                '
                ' pass any delimiters
                '
                If CurrentPosition <> 0 Then
                    Do While Mid(WorkingWordList, CurrentPosition, 1) = Delimiter
                        CurrentPosition = CurrentPosition + 1
                        If CurrentPosition >= LenWorkingWordList Then
                            CurrentPosition = 0
                            Exit Do
                            End If
                        Loop
                    End If
            Loop
        End If
    kmaSplit_Old = Result
    '
    
    '
    'kmaSplit_Old = Split(WorkingWordList, Delimiter, , vbTextCompare)
    '
    End Function
'
'
'
Public Function kmaGetYesNo(Key As Boolean) As String
    If Key Then
        kmaGetYesNo = "Yes"
    Else
        kmaGetYesNo = "No"
        End If
    End Function
'
'
'
Public Function kmaGetFilename(PathFilename As String) As String
    Dim Position As Long
    '
    kmaGetFilename = PathFilename
    Position = InStrRev(kmaGetFilename, "/")
    If Position <> 0 Then
        kmaGetFilename = Mid(kmaGetFilename, Position + 1)
        End If
    End Function
'
'
'
Public Function kmaStartTable(Padding As Long, Spacing As Long, Border As Long, Optional ClassStyle As String) As String
    kmaStartTable = "<table border=""" & Border & """ cellpadding=""" & Padding & """ cellspacing=""" & Spacing & """ class=""" & ClassStyle & """ width=""100%"">"
    End Function
'
'
'
Public Function kmaStartTableRow() As String
    kmaStartTableRow = "<tr>"
    End Function
'
'
'
Public Function kmaStartTableCell(Optional Width As String, Optional ColSpan As Long, Optional EvenRow As Boolean, Optional Align As String, Optional BGColor As String) As String
    If Width <> "" Then
        kmaStartTableCell = " width=""" & Width & """"
        End If
    If BGColor <> "" Then
        kmaStartTableCell = kmaStartTableCell & " bgcolor=""" & BGColor & """"
    ElseIf EvenRow Then
        kmaStartTableCell = kmaStartTableCell & " class=""ccPanelRowEven"""
    Else
        kmaStartTableCell = kmaStartTableCell & " class=""ccPanelRowOdd"""
        End If
    If ColSpan <> 0 Then
        kmaStartTableCell = kmaStartTableCell & " colspan=""" & ColSpan & """"
        End If
    If Align <> "" Then
        kmaStartTableCell = kmaStartTableCell & " align=""" & Align & """"
        End If
    kmaStartTableCell = "<TD" & kmaStartTableCell & ">"
    End Function
'
'
'
Public Function kmaGetTableCell(Copy As String, Optional Width As String, Optional ColSpan As Long, Optional EvenRow As Boolean, Optional Align As String, Optional BGColor As String) As String
    kmaGetTableCell = kmaStartTableCell(Width, ColSpan, EvenRow, Align, BGColor) & Copy & kmaEndTableCell
    End Function
'
'
'
Public Function kmaGetTableRow(Cell As String, Optional ColSpan As Long, Optional EvenRow As Boolean) As String
    kmaGetTableRow = kmaStartTableRow & kmaGetTableCell(Cell, "100%", ColSpan, EvenRow) & kmaEndTableRow
    End Function
'
' remove the host and approotpath, leaving the "active" path and all else
'
Public Function kmaConvertShortLinkToLink(URL As String, PathPagePrefix As String) As String
    kmaConvertShortLinkToLink = URL
    If URL <> "" And PathPagePrefix <> "" Then
        If InStr(1, kmaConvertShortLinkToLink, PathPagePrefix, vbTextCompare) = 1 Then
            kmaConvertShortLinkToLink = Mid(kmaConvertShortLinkToLink, Len(PathPagePrefix) + 1)
            End If
        End If
    End Function
'
' ------------------------------------------------------------------------------------------------------
'   Preserve URLs that do not start HTTP or HTTPS
'   Preserve URLs from other sites (offsite)
'   Preserve HTTP://ServerHost/ServerVirtualPath/Files/ in all cases
'   Convert HTTP://ServerHost/ServerVirtualPath/folder/page -> /folder/page
'   Convert HTTP://ServerHost/folder/page -> /folder/page
' ------------------------------------------------------------------------------------------------------
'
Public Function kmaConvertLinkToShortLink(URL As String, ServerHost As String, ServerVirtualPath As String) As String
    '
    Dim BadString As String
    Dim GoodString As String
    Dim Protocol As String
    Dim WorkingLink As String
    '
    WorkingLink = URL
    '
    ' ----- Determine Protocol
    '
    If InStr(1, WorkingLink, "HTTP://", vbTextCompare) = 1 Then
        '
        ' HTTP
        '
        Protocol = Mid(WorkingLink, 1, 7)
    ElseIf InStr(1, WorkingLink, "HTTPS://", vbTextCompare) = 1 Then
        '
        ' HTTPS
        '
        ' try this -- a ssl link can not be shortened
        kmaConvertLinkToShortLink = WorkingLink
        Exit Function
        Protocol = Mid(WorkingLink, 1, 8)
        End If
    If Protocol <> "" Then
        '
        ' ----- Protcol found, determine if is local
        '
        GoodString = Protocol & ServerHost
        If (InStr(1, WorkingLink, GoodString, vbTextCompare) <> 0) Then
            '
            ' URL starts with Protocol ServerHost
            '
            GoodString = Protocol & ServerHost & ServerVirtualPath & "/files/"
            If (InStr(1, WorkingLink, GoodString, vbTextCompare) <> 0) Then
                '
                ' URL is in the virtual files directory
                '
                BadString = GoodString
                GoodString = ServerVirtualPath & "/files/"
                WorkingLink = Replace(WorkingLink, BadString, GoodString, , , vbTextCompare)
            Else
                '
                ' URL is not in files virtual directory
                '
                BadString = Protocol & ServerHost & ServerVirtualPath & "/"
                GoodString = "/"
                WorkingLink = Replace(WorkingLink, BadString, GoodString, , , vbTextCompare)
                '
                BadString = Protocol & ServerHost & "/"
                GoodString = "/"
                WorkingLink = Replace(WorkingLink, BadString, GoodString, , , vbTextCompare)
                End If
            End If
        End If
    kmaConvertLinkToShortLink = WorkingLink
    End Function
'
' Correct the link for the virtual path, either add it or remove it
'
Public Function kmaEncodeAppRootPath(Link As String, VirtualPath As String, AppRootPath As String, ServerHost As String) As String
    '
    Dim Protocol As String
    Dim Host As String
    Dim Path As String
    Dim Page As String
    Dim QueryString As String
    Dim VirtualHosted As Boolean
    '
    kmaEncodeAppRootPath = Link
    If (InStr(1, kmaEncodeAppRootPath, ServerHost, vbTextCompare) <> 0) Or (InStr(1, Link, "/") = 1) Then
    'If (InStr(1, kmaEncodeAppRootPath, ServerHost, vbTextCompare) <> 0) And (InStr(1, Link, "/") <> 0) Then
        '
        ' This link is onsite and has a path
        '
        VirtualHosted = (InStr(1, AppRootPath, VirtualPath, vbTextCompare) <> 0)
        If VirtualHosted And (InStr(1, Link, AppRootPath, vbTextCompare) = 1) Then
            '
            ' quick - virtual hosted and link starts at AppRootPath
            '
        ElseIf (Not VirtualHosted) And (Mid(Link, 1, 1) = "/") And (InStr(1, Link, AppRootPath, vbTextCompare) = 1) Then
            '
            ' quick - not virtual hosted and link starts at Root
            '
        Else
            Call SeparateURL(Link, Protocol, Host, Path, Page, QueryString)
            If VirtualHosted Then
                '
                ' Virtual hosted site, add VirualPath if it is not there
                '
                If InStr(1, Path, AppRootPath, vbTextCompare) = 0 Then
                    If Path = "/" Then
                        Path = AppRootPath
                    Else
                        Path = AppRootPath & Mid(Path, 2)
                        End If
                    End If
            Else
                '
                ' Root hosted site, remove virtual path if it is there
                '
                If InStr(1, Path, AppRootPath, vbTextCompare) <> 0 Then
                    Path = Replace(Path, AppRootPath, "/")
                    End If
                End If
            kmaEncodeAppRootPath = Protocol & Host & Path & Page & QueryString
            End If
        End If
    End Function
'
' Return just the tablename from a tablename reference (database.object.tablename->tablename)
'
Function GetDbObjectTableName(DbObject As String) As String
    Dim Position As Long
    '
    GetDbObjectTableName = DbObject
    Position = InStrRev(GetDbObjectTableName, ".")
    If Position > 0 Then
        GetDbObjectTableName = Mid(GetDbObjectTableName, Position + 1)
        End If
    End Function
'
'
'
Function kmaGetLinkedText(AnchorTag As Variant, AnchorText As Variant) As String
    '
    Dim UcaseAnchorText As String
    Dim LinkPosition As Long
    Dim MethodName As String
    Dim iAnchorTag As String
    Dim iAnchorText As String
    '
    MethodName = "kmaGetLinkedText"
    '
    kmaGetLinkedText = ""
    iAnchorTag = kmaEncodeText(AnchorTag)
    iAnchorText = kmaEncodeText(AnchorText)
    UcaseAnchorText = UCase(iAnchorText)
    If (iAnchorTag <> "") And (iAnchorText <> "") Then
        LinkPosition = InStrRev(UcaseAnchorText, "<LINK>", -1)
        If LinkPosition = 0 Then
            kmaGetLinkedText = iAnchorTag & iAnchorText & "</A>"
        Else
            kmaGetLinkedText = iAnchorText
            LinkPosition = InStrRev(UcaseAnchorText, "</LINK>", -1)
            Do While LinkPosition > 1
                kmaGetLinkedText = Mid(kmaGetLinkedText, 1, LinkPosition - 1) & "</A>" & Mid(kmaGetLinkedText, LinkPosition + 7)
                LinkPosition = InStrRev(UcaseAnchorText, "<LINK>", LinkPosition - 1)
                If LinkPosition <> 0 Then
                    kmaGetLinkedText = Mid(kmaGetLinkedText, 1, LinkPosition - 1) & iAnchorTag & Mid(kmaGetLinkedText, LinkPosition + 6)
                    End If
                LinkPosition = InStrRev(UcaseAnchorText, "</LINK>", LinkPosition)
                Loop
            End If
        End If
    '
    End Function
'
'========================================================================
'   HandleError
'       Logs the error and either resumes next, or raises it to the next level
'========================================================================
'
Public Function HandleError(ClassName As String, MethodName As String, ErrNumber As Long, ErrSource As String, ErrDescription As String, ErrorTrap As Boolean, ResumeNext As Boolean, Optional URL As String) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim ErrorMessage As String
    '
    If ErrorTrap Then
        ErrorMessage = ErrorMessage & " Unexpected ErrorTrap"
    Else
        ErrorMessage = ErrorMessage & " Error"
        End If
    '
    If URL <> "" Then
        ErrorMessage = ErrorMessage & " on page [" & URL & "]"
        End If
    '
    If ErrorTrap Then
        If ResumeNext Then
            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will resume after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
        Else
            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will abort after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
            On Error GoTo 0
            Call Err.Raise(ErrNumber, ErrSource, ErrDescription)
            End If
    Else
        If ResumeNext Then
            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will resume after logging  [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
        Else
            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will abort after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
            On Error GoTo 0
            Call Err.Raise(ErrNumber, ErrSource, ErrDescription, , -1)
            End If
        End If
    '
    End Function
'
'==========================================================================================
' handle error and resume next
'==========================================================================================
'
Public Sub HandleErrorAndResumeNext(ClassName As String, MethodName As String, Optional Description As String, Optional ErrorNumber As Long)
    Dim ErrDescription As String
    Dim ErrSource As String
    Dim ErrNumber As Long
    '
    ErrNumber = Err.Number
    ErrSource = Err.Source
    ErrDescription = Err.Description
    '
    If ErrNumber = 0 Then
        '
        ' internal error, no VB error
        '
        If Description = "" Then
            ErrDescription = "Unknown error"
        Else
            ErrDescription = Description
        End If
        If ErrorNumber = 0 Then
            Call HandleError(ClassName, MethodName, KmaErrorInternal, App.EXEName, Description, False, True)
        Else
            Call HandleError(ClassName, MethodName, ErrNumber, App.EXEName, Description, False, True)
        End If
    Else
        '
        ' VB Error
        '
        If Description <> "" Then
            ErrDescription = Description & ",VB Error [" & Err.Description & "]"
        Else
            ErrDescription = "Unexpected VB Error [" & Err.Description & "]"
        End If
        Call HandleError(ClassName, MethodName, ErrNumber, ErrSource, ErrDescription, True, True)
    End If
End Sub
'
'
'
Public Sub AppendLogFile(Text)
    On Error GoTo 0
    '
    Dim MonthNumber As Long
    Dim DayNumber As Long
    Dim Filename As String
    '
    DayNumber = Day(Now)
    MonthNumber = Month(Now)
    Filename = Year(Now)
    If MonthNumber < 10 Then
        Filename = Filename & "0"
    End If
    Filename = Filename & MonthNumber
    If DayNumber < 10 Then
        Filename = Filename & "0"
    End If
    Filename = Filename & DayNumber
    '
    Call AppendLog("Trace" & Filename & ".log", kmaEncodeText(Text))
    '
End Sub
'
'
'
Public Sub AppendLog(Filename As String, Text As String)
    On Error GoTo 0
    Dim kmafs As Object
    '
    If (Filename <> "") Then
        Set kmafs = CreateObject("kmaFileSystem3.FileSystemClass")
        Call kmafs.AppendFile(App.Path & "\logs\" & Filename, """" & FormatDateTime(Now(), vbGeneralDate) & """,""" & Text & """" & vbCrLf)
        Set kmafs = Nothing
    End If
    '
    End Sub
'
'
'
Public Function kmaErrorMsg(ErrorNumber As Long) As String
    kmaErrorMsg = ""
End Function
'
'
'
Public Function kmaEncodeInitialCaps(Source As String) As String
    Dim SegSplit() As String
    Dim SegPtr As Long
    Dim SegMax As Long
    '
    If Source <> "" Then
        SegSplit = Split(Source, " ")
        SegMax = UBound(SegSplit)
        If SegMax >= 0 Then
            For SegPtr = 0 To SegMax
                SegSplit(SegPtr) = UCase(Left(SegSplit(SegPtr), 1)) & LCase(Mid(SegSplit(SegPtr), 2))
            Next
        End If
        kmaEncodeInitialCaps = Join(SegSplit, " ")
    End If
End Function
'
'
'
Public Function kmaRemoveTag(Source As String, TagName As String) As String
    Dim Pos As Long
    Dim posEnd As Long
    kmaRemoveTag = Source
    Pos = InStr(1, Source, "<" & TagName, vbTextCompare)
    If Pos <> 0 Then
        posEnd = InStr(Pos, Source, ">")
        If posEnd > 0 Then
            kmaRemoveTag = Left(Source, Pos - 1) & Mid(Source, posEnd + 1)
        End If
    End If
End Function
'
'
'
Public Function kmaRemoveStyleTags(Source As String) As String
    kmaRemoveStyleTags = Source
    Do While InStr(1, kmaRemoveStyleTags, "<style", vbTextCompare) <> 0
        kmaRemoveStyleTags = kmaRemoveTag(kmaRemoveStyleTags, "style")
    Loop
    Do While InStr(1, kmaRemoveStyleTags, "</style", vbTextCompare) <> 0
        kmaRemoveStyleTags = kmaRemoveTag(kmaRemoveStyleTags, "/style")
    Loop
End Function
'
'
'
Public Function kmaGetSingular(PluralSource As String) As String
    '
    Dim UpperCase As Boolean
    Dim LastCharacter As String
    '
    kmaGetSingular = PluralSource
    If Len(kmaGetSingular) > 1 Then
        LastCharacter = Right(kmaGetSingular, 1)
        If LastCharacter <> UCase(LastCharacter) Then
            UpperCase = True
        End If
        If UCase(Right(kmaGetSingular, 3)) = "IES" Then
            If UpperCase Then
                kmaGetSingular = Mid(kmaGetSingular, 1, Len(kmaGetSingular) - 3) & "Y"
            Else
                kmaGetSingular = Mid(kmaGetSingular, 1, Len(kmaGetSingular) - 3) & "y"
            End If
        ElseIf UCase(Right(kmaGetSingular, 2)) = "SS" Then
            ' nothing
        ElseIf UCase(Right(kmaGetSingular, 1)) = "S" Then
            kmaGetSingular = Mid(kmaGetSingular, 1, Len(kmaGetSingular) - 1)
        Else
            ' nothing
        End If
    End If
End Function
'
'
'
Public Function kmaEncodeJavascript(Source As String) As String
    '
    kmaEncodeJavascript = Source
    kmaEncodeJavascript = Replace(kmaEncodeJavascript, "'", "\'")
    'kmaEncodeJavascript = Replace(kmaEncodeJavascript, "'", "'+""'""+'")
    kmaEncodeJavascript = Replace(kmaEncodeJavascript, vbCrLf, "\n")
    kmaEncodeJavascript = Replace(kmaEncodeJavascript, vbCr, "\n")
    kmaEncodeJavascript = Replace(kmaEncodeJavascript, vbLf, "\n")
    kmaEncodeJavascript = Replace(kmaEncodeJavascript, "</script", "</scr'+'ipt", 1, 99, vbTextCompare)
    '
End Function
'
'   Indent every line by 1 tab
'
Public Function kmaIndent(Source As String) As String
    Dim posStart As Long
    Dim posEnd As Long
    Dim pre As String
    Dim post As String
    Dim target As String
    '
    posStart = InStr(1, Source, "<![CDATA[", 1)
    If posStart = 0 Then
        '
        ' no cdata
        '
        posStart = InStr(1, Source, "<textarea", 1)
        If posStart = 0 Then
            '
            ' no textarea
            '
            kmaIndent = Replace(Source, vbCrLf & vbTab, vbCrLf & vbTab & vbTab)
        Else
            '
            ' text area found, isolate it and indent before and after
            '
            posEnd = InStr(posStart, Source, "</textarea>", 1)
            pre = Mid(Source, 1, posStart - 1)
            If posEnd = 0 Then
                target = Mid(Source, posStart)
                post = ""
            Else
                target = Mid(Source, posStart, posEnd - posStart + Len("</textarea>"))
                post = Mid(Source, posEnd + Len("</textarea>"))
            End If
            kmaIndent = kmaIndent(pre) & target & kmaIndent(post)
        End If
    Else
        '
        ' cdata found, isolate it and indent before and after
        '
        posEnd = InStr(posStart, Source, "]]>", 1)
        pre = Mid(Source, 1, posStart - 1)
        If posEnd = 0 Then
            target = Mid(Source, posStart)
            post = ""
        Else
            target = Mid(Source, posStart, posEnd - posStart + Len("]]>"))
            post = Mid(Source, posEnd + 3)
        End If
        kmaIndent = kmaIndent(pre) & target & kmaIndent(post)
    End If
'    kmaIndent = Source
'    If InStr(1, kmaIndent, "<textarea", vbTextCompare) = 0 Then
'        kmaIndent = Replace(Source, vbCrLf & vbTab, vbCrLf & vbTab & vbTab)
'    End If
End Function
'
'
'
Public Function kmaGetListIndex(Item As String, ListOfItems As String) As Long
    '
    Dim Items() As String
    Dim LcaseItem As String
    Dim LcaseList As String
    Dim Ptr As Long
    '
    If ListOfItems <> "" Then
        LcaseItem = LCase(Item)
        LcaseList = LCase(ListOfItems)
        Items = Split(LcaseList, ",")
        For Ptr = 0 To UBound(Items)
            If Items(Ptr) = LcaseItem Then
                kmaGetListIndex = Ptr + 1
                Exit For
            End If
        Next
    End If
    '
End Function
'
'========================================================================================================
'
' Finds all tags matching the input, and concatinates them into the output
' does NOT account for nested tags, use for body, script, style
'
' ReturnAll - if true, it returns all the occurances, back-to-back
'
'========================================================================================================
'
Public Function GetTagInnerHTML(PageSource As String, Tag As String, ReturnAll As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim TagStart As Long
    Dim TagEnd As Long
    Dim LoopCnt As Long
    Dim WB As String
    Dim Pos As Long
    Dim posEnd As Long
    Dim CommentPos As Long
    Dim ScriptPos As Long
    '
    Pos = 1
    Do While (Pos > 0) And (LoopCnt < 100)
        TagStart = InStr(Pos, PageSource, "<" & Tag, vbTextCompare)
        If TagStart = 0 Then
            Pos = 0
        Else
            '
            ' tag found, skip any comments that start between current position and the tag
            '
            CommentPos = InStr(Pos, PageSource, "<!--")
            If (CommentPos <> 0) And (CommentPos < TagStart) Then
                '
                ' skip comment and start again
                '
                Pos = InStr(CommentPos, PageSource, "-->")
            Else
                ScriptPos = InStr(Pos, PageSource, "<script")
                If (ScriptPos <> 0) And (ScriptPos < TagStart) Then
                    '
                    ' skip comment and start again
                    '
                    Pos = InStr(ScriptPos, PageSource, "</script")
                Else
                    '
                    ' Get the tags innerHTML
                    '
                    TagStart = InStr(TagStart, PageSource, ">", vbTextCompare)
                    Pos = TagStart
                    If TagStart <> 0 Then
                        TagStart = TagStart + 1
                        TagEnd = InStr(TagStart, PageSource, "</" & Tag, vbTextCompare)
                        If TagEnd <> 0 Then
                            GetTagInnerHTML = GetTagInnerHTML & Mid(PageSource, TagStart, TagEnd - TagStart)
                        End If
                    End If
                End If
            End If
            LoopCnt = LoopCnt + 1
            If ReturnAll Then
                TagStart = InStr(TagEnd, PageSource, "<" & Tag, vbTextCompare)
            Else
                TagStart = 0
            End If
        End If
    Loop
    '
    Exit Function
    '
ErrorTrap:
    'Call HandleError("EncodePage_SplitBody")
End Function
'
'========================================================================================================
'Place code in a form module
'Add a Command button.
'========================================================================================================
'
Public Function kmaByteArrayToString(Bytes() As Byte) As String
    Dim iUnicode As Long, i As Long, j As Long
    
    On Error Resume Next
    i = UBound(Bytes)
    
    If (i < 1) Then
        'ANSI, just convert to unicode and return
        kmaByteArrayToString = StrConv(Bytes, vbUnicode)
        Exit Function
    End If
    i = i + 1
    
    'Examine the first two bytes
    CopyMemory iUnicode, Bytes(0), 2
    
    If iUnicode = Bytes(0) Then 'Unicode
        'Account for terminating null
        If (i Mod 2) Then i = i - 1
        'Set up a buffer to recieve the string
        kmaByteArrayToString = String$(i / 2, 0)
        'Copy to string
        CopyMemory ByVal StrPtr(kmaByteArrayToString), Bytes(0), i
    Else 'ANSI
        kmaByteArrayToString = StrConv(Bytes, vbUnicode)
    End If
                    
End Function
'
'========================================================================================================
'
'========================================================================================================
'
Public Function kmaStringToByteArray(strInput As String, _
                                Optional bReturnAsUnicode As Boolean = True, _
                                Optional bAddNullTerminator As Boolean = False) As Byte()
    
    Dim lRet As Long
    Dim bytBuffer() As Byte
    Dim lLenB As Long
    
    If bReturnAsUnicode Then
        'Number of bytes
        lLenB = LenB(strInput)
        'Resize buffer, do we want terminating null?
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        'Copy characters from string to byte array
        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
    Else
        'METHOD ONE
'        'Get rid of embedded nulls
'        strRet = StrConv(strInput, vbFromUnicode)
'        lLenB = LenB(strRet)
'        If bAddNullTerminator Then
'            ReDim bytBuffer(lLenB)
'        Else
'            ReDim bytBuffer(lLenB - 1)
'        End If
'        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
        
        'METHOD TWO
        'Num of characters
        lLenB = Len(strInput)
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        lRet = WideCharToMultiByte(CP_ACP, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(bytBuffer(0)), lLenB, 0&, 0&)
    End If
    
    kmaStringToByteArray = bytBuffer
    
End Function
'
'========================================================================================================
'   Sample kmaStringToByteArray
'========================================================================================================
'
Private Sub SampleStringToByteArray()
    Dim bAnsi() As Byte
    Dim bUni() As Byte
    Dim str As String
    Dim i As Long
    '
    str = "Convert"
    bAnsi = kmaStringToByteArray(str, False)
    bUni = kmaStringToByteArray(str)
    '
    For i = 0 To UBound(bAnsi)
        Debug.Print "=" & bAnsi(i)
    Next
    '
    Debug.Print "========"
    '
    For i = 0 To UBound(bUni)
        Debug.Print "=" & bUni(i)
    Next
    '
    Debug.Print "ANSI= " & kmaByteArrayToString(bAnsi)
    Debug.Print "UNICODE= " & kmaByteArrayToString(bUni)
    'Using StrConv to convert a Unicode character array directly
    'will cause the resultant string to have extra embedded nulls
    'reason, StrConv does not know the difference between Unicode and ANSI
    Debug.Print "Resull= " & StrConv(bUni, vbUnicode)
End Sub

