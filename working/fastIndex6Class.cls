VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fastIndex6Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
' ----- Index Type - This structure is the basis for Element Indexing
'       Records are read into thier data structure, and keys(Key,ID,etc.) and pointers
'       are put in the KeyPointerArrays.
'           AddIndex( Key, value )
'           BubbleSort( Index ) - sorts the index by the key field
'           GetIndexValue( index, Key ) - retrieves the pointer
'
Private Const KeyPointerArrayChunk = 1000
'
Private Type KeyPointerArrayType
    UcaseKey As String                      ' key value
    Pointer As Long                     ' Pointer into the data being indexed
    End Type
'
Private ArraySize As Long
Private ArrayCount As Long
Private ArrayDirty As Boolean
Private UcaseKeyArray() As String
Private PointerArray() As String
'Private PointerArray() As Long
Private ArrayPointer As Long
'
'========================================================================
'   Returns a pointer into the index for this Key
'   Used only by GetIndexValue and setIndexValue
'   Returns -1 if there is no match
'========================================================================
'
Private Function GetArrayPointer(Key As String) As Long
    On Error GoTo ErrorTrap
    '
    Dim PointerStep As Long
    Dim UcaseTargetKey As String
    Dim ElementKey As String
    Dim HighGuess As Long
    Dim LowGuess As Long
    Dim PointerGuess As Long
    Dim MethodName As String
    Dim test As String
    test = ""
    '
    MethodName = "ContentServer.GetArrayPointer"
    '
    If ArrayDirty Then
        Call Sort
    End If
    '
    GetArrayPointer = -1
    
    If ArrayCount > 0 Then
        UcaseTargetKey = Replace(UCase(Key), vbCrLf, "")
        LowGuess = -1
        HighGuess = ArrayCount - 1
        Do While (HighGuess - LowGuess) > 1
            PointerGuess = (HighGuess + LowGuess) / 2
            If UcaseTargetKey = UcaseKeyArray(PointerGuess) Then
                HighGuess = PointerGuess
                Exit Do
            ElseIf UcaseTargetKey < UcaseKeyArray(PointerGuess) Then
                HighGuess = PointerGuess
            Else
                LowGuess = PointerGuess
            End If
        Loop
        If UcaseTargetKey = UcaseKeyArray(HighGuess) Then
            GetArrayPointer = HighGuess
        End If
    End If
    Exit Function
    '
ErrorTrap:
    Call HandleError("GetArrayPointer", Err.Number, Err.Source, Err.Description)
End Function
'
'========================================================================
'   Returns the matching pointer from a ContentIndex
'   Returns -1 if there is no match
'========================================================================
'
Public Function GetPointer(Key As String) As Long
    On Error GoTo ThisAndThat
    '
    Dim test As String
    Dim MatchFound As Boolean
    Dim UcaseKey As String
    test = ""
    '
    UcaseKey = Replace(UCase(Key), vbCrLf, "")
    'UcaseKey = UCase(Key)
    ArrayPointer = GetArrayPointer(Key)
    If ArrayPointer > -1 Then
        ' Make sure this is the first match
        MatchFound = True
        Do While MatchFound
            ArrayPointer = ArrayPointer - 1
            If ArrayPointer < 0 Then
                MatchFound = False
            Else
                MatchFound = (UcaseKeyArray(ArrayPointer) = UcaseKey)
            End If
        Loop
        ArrayPointer = ArrayPointer + 1
        GetPointer = kmaEncodeInteger(PointerArray(ArrayPointer))
    Else
        GetPointer = -1
    End If
    Exit Function
    '
ThisAndThat:
    Call HandleError("GetPointer", Err.Number, Err.Source, Err.Description)
End Function
'
'========================================================================
'   Add an element to an ContentIndex
'
'   if the entry is a duplicate, it is added anyway
'========================================================================
'
Public Sub SetPointer(Key As String, Pointer As Long)
    On Error GoTo ErrorTrap
    '
    Dim test As String
    Dim keyToSave As String
    Dim testArrayPtr As Long
    '
    keyToSave = Replace(UCase(Key), vbCrLf, "")
    '
    ' test if this key/Pointer pair is already saved
    '
'    testArrayPtr = GetArrayPointer(keyToSave)
'    If testArrayPtr <> -1 Then
'        If (keyToSave <> UcaseKeyArray(testArrayPtr)) And (Pointer = PointerArray(testArrayPtr)) Then
'            testArrayPtr = -1
'        End If
'    End If
'    If testArrayPtr = -1 Then
        '
        ' save the pair if they are not already here
        '
        If ArrayCount >= ArraySize Then
            ArraySize = ArraySize + KeyPointerArrayChunk
            ReDim Preserve PointerArray(ArraySize)
            ReDim Preserve UcaseKeyArray(ArraySize)
        End If
        ArrayPointer = ArrayCount
        ArrayCount = ArrayCount + 1
        UcaseKeyArray(ArrayPointer) = keyToSave
        PointerArray(ArrayPointer) = CStr(Pointer)
        ArrayDirty = True
'    End If
    Exit Sub
    '
ErrorTrap:
    Call HandleError("SetPointerAllowDups", Err.Number, Err.Source, Err.Description)
End Sub
'
'========================================================================
'   Returns the next matching pointer from a ContentIndex
'   Returns -1 if there is no match
'========================================================================
'
Public Function GetNextPointerMatch(Key As String) As Long
    On Error GoTo ThisAndThat
    '
    Dim MatchFound As Boolean
    Dim UcaseKey As String
    '
    GetNextPointerMatch = -1
    If ArrayPointer < (ArrayCount - 1) Then
        ArrayPointer = ArrayPointer + 1
        UcaseKey = UCase(Key)
        If (UcaseKeyArray(ArrayPointer) = UcaseKey) Then
            GetNextPointerMatch = kmaEncodeInteger(PointerArray(ArrayPointer))
        Else
            ArrayPointer = ArrayPointer - 1
        End If
    End If
    Exit Function
    '
ThisAndThat:
    Call HandleError("GetNextPointerMatch", Err.Number, Err.Source, Err.Description)
End Function
'
'========================================================================
'   Returns the first Pointer in the current index
'   returns empty if there are no Pointers indexed
'========================================================================
'
Public Function GetFirstPointer() As Long
    On Error GoTo ThisAndThat
    '
    If ArrayDirty Then
        Call Sort
    End If
    '
    GetFirstPointer = -1
    If ArrayCount > 0 Then
        ArrayPointer = 0
        GetFirstPointer = kmaEncodeInteger(PointerArray(ArrayPointer))
    End If
    '
    Exit Function
    '
ThisAndThat:
    Call HandleError("GetFirstPointer", Err.Number, Err.Source, Err.Description)
End Function
'
'========================================================================
'   Returns the next Pointer, past the last one returned
'   Returns empty if the index is at the end
'========================================================================
'
Public Function GetNextPointer() As Long
    On Error GoTo ThisAndThat
    '
    If ArrayDirty Then
        Call Sort
    End If
    '
    GetNextPointer = -1
    If (ArrayPointer + 1) < ArrayCount Then
        ArrayPointer = ArrayPointer + 1
        GetNextPointer = kmaEncodeInteger(PointerArray(ArrayPointer))
    End If
    '
    Exit Function
    '
ThisAndThat:
    Call HandleError("GetFirstPointer", Err.Number, Err.Source, Err.Description)
End Function
'
'========================================================================
'   HandleError
'========================================================================
'
Private Sub HandleError(MethodName As String, ErrNumber As Long, ErrSource As String, ErrDescription As String)
    '
    On Error GoTo 0
    Call Err.Raise(ErrNumber, ErrSource, App.EXEName & ".FastIndex." & MethodName & " encountered and error: " & ErrDescription)
    '
End Sub
'
'
'
Public Function ExportPropertyBag() As Variant
    On Error GoTo ErrorTrap
    '
    Dim PB As New PropertyBag
    '
    Call PB.WriteProperty("UcaseKeyArray", Join(UcaseKeyArray, vbCrLf))
    Call PB.WriteProperty("PointerArray", Join(PointerArray, vbCrLf))
    Call PB.WriteProperty("ArrayCount", ArrayCount)
    Call PB.WriteProperty("ArrayPointer", ArrayPointer)
    Call PB.WriteProperty("ArrayDirty", ArrayDirty)
    ExportPropertyBag = PB.Contents
    '
    Exit Function
    '
ErrorTrap:
    Call HandleError("ExportPropertyBag", Err.Number, Err.Source, Err.Description)
End Function
'
'
'
Public Sub ImportPropertyBag(bag As Variant)
    On Error GoTo ErrorTrap
    '
    Dim PB As New PropertyBag
    '
    PB.Contents = bag
    UcaseKeyArray = Split(PB.ReadProperty("UcaseKeyArray"), vbCrLf)
    PointerArray = Split(PB.ReadProperty("PointerArray"), vbCrLf)
    ArrayCount = PB.ReadProperty("ArrayCount")
    ArrayPointer = PB.ReadProperty("ArrayPointer")
    ArrayDirty = PB.ReadProperty("ArrayDirty")
    '
    Exit Sub
    '
ErrorTrap:
    Call HandleError("ImportPropertyBag", Err.Number, Err.Source, Err.Description)
End Sub
'
'========================================================================
'
'========================================================================
'
Private Sub BubbleSort()
    On Error GoTo ErrorTrap
    '
    Dim TempUcaseKey As String
    Dim tempPtrString As String
    'Dim TempPointer As Long
    Dim CleanPass As Boolean
    Dim MaxPointer As Long
    Dim SlowPointer As Long
    Dim FastPointer As Long
    Dim test As String
    Dim PointerDelta As Long
    test = ""
    '
    If ArrayCount > 1 Then
        PointerDelta = 1
        MaxPointer = ArrayCount - 2
        For SlowPointer = MaxPointer To 0 Step -1
            CleanPass = True
            For FastPointer = MaxPointer To (MaxPointer - SlowPointer) Step -1
                If UcaseKeyArray(FastPointer) > UcaseKeyArray(FastPointer + PointerDelta) Then
                    TempUcaseKey = UcaseKeyArray(FastPointer + PointerDelta)
                    tempPtrString = PointerArray(FastPointer + PointerDelta)
                    UcaseKeyArray(FastPointer + PointerDelta) = UcaseKeyArray(FastPointer)
                    PointerArray(FastPointer + PointerDelta) = PointerArray(FastPointer)
                    UcaseKeyArray(FastPointer) = TempUcaseKey
                    PointerArray(FastPointer) = tempPtrString
                    CleanPass = False
                End If
            Next
            If CleanPass Then
                Exit For
            End If
        Next
    End If
    ArrayDirty = False
    Exit Sub
    '
ErrorTrap:
    Call HandleError("BubbleSort", Err.Number, Err.Source, Err.Description)
End Sub
'
'========================================================================
'
' Made by Michael Ciurescu (CVMichael from vbforums.com)
' Original thread: http://www.vbforums.com/showthread.php?t=231925
'
'========================================================================
'
Private Sub QuickSort()
    On Error GoTo ErrorTrap
    '
    If ArrayCount >= 2 Then
        Call QuickSort_Segment(UcaseKeyArray, PointerArray, 0, ArrayCount - 1)
    End If
    '
    Exit Sub
    '
ErrorTrap:
    Call HandleError("QuickSort", Err.Number, Err.Source, Err.Description)
End Sub
    '
'
'========================================================================
'
' Made by Michael Ciurescu (CVMichael from vbforums.com)
' Original thread: http://www.vbforums.com/showthread.php?t=231925
'
'========================================================================
'
Private Sub QuickSort_Segment(C() As String, P() As String, ByVal First As Long, ByVal Last As Long)
    On Error GoTo ErrorTrap
    '
    '
    Dim Low As Long, High As Long
    Dim MidValue As String
    Dim TC As String
    Dim TP As String
    '
    Low = First
    High = Last
    MidValue = C((First + Last) \ 2)
    '
    Do
        While C(Low) < MidValue
            Low = Low + 1
        Wend
        While C(High) > MidValue
            High = High - 1
        Wend
        If Low <= High Then
            TC = C(Low)
            TP = P(Low)
            C(Low) = C(High)
            P(Low) = P(High)
            C(High) = TC
            P(High) = TP
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    If First < High Then
        QuickSort_Segment C, P, First, High
    End If
    If Low < Last Then
        QuickSort_Segment C, P, Low, Last
    End If
    '
    Exit Sub
    '
ErrorTrap:
    Call HandleError("QuickSort_Segment", Err.Number, Err.Source, Err.Description)
End Sub
'
'
'
Private Sub Sort()
    On Error GoTo ErrorTrap
    '
    Call QuickSort
    ArrayDirty = False
    '
    Exit Sub
    '
ErrorTrap:
    Call HandleError("Sort", Err.Number, Err.Source, Err.Description)
End Sub



