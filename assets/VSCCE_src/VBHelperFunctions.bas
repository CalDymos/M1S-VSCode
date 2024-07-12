Attribute VB_Name = "VBHelp"
'Option Explicit
'

'
'Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function StrTrim Lib "shlwapi" Alias "StrTrimW" (ByVal pszSource As Long, ByVal pszTrimChars As Long) As Long
'
'Private Type SafeArray
'    cDims As Integer
'    fFeatures As Integer
'    cbElements As Long
'    cLocks As Long
'    pvData As Long
'End Type
'
'
'Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (arr() As Any) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFilename As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Function MakePath Lib "imagehlp.dll" _
        Alias "MakeSureDirectoryPathExists" (ByVal _
        lpPath As String) As Long
Public Declare Function GetLastError Lib "kernel32.dll" () As Long
        
' Flag for debug mode
Private mbDebugMode As Boolean


' Set mbDebugMode to true. This happens only
' if the Debug.Assert call happens. It only
' happens in the IDE.
Private Function InDebugMode() As Boolean
    mbDebugMode = True
    InDebugMode = True
End Function
' Set the mbDebugMode flag.
Public Function DebugMode() As Boolean
    ' This will only be done if in the IDE
    Debug.Assert InDebugMode
    If mbDebugMode Then
        DebugMode = True
    End If
End Function

Public Function FolderExists(ByVal Path As String) As Boolean
  ' If not available, add backslash to the path
  If Right$(Path, 1) <> "\" Then Path = Path & "\"
 
  ' Check whether directory exists
  FolderExists = (UCase$(dir$(Path & "\nul")) = "NUL")
End Function

Public Function FileExists(ByVal fName As String) As Boolean

    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT

    lRetVal = OpenFile(fName, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        FileExists = True
    Else
        FileExists = False
    End If

End Function


Function IsArrayInitialized(arr As Variant) As Boolean

    If IsArray(arr) Then

        If VarType(arr) = vbArray + vbString Then
            Dim aStr() As String

            aStr = arr

            If ((Not aStr) = -1) Then
                IsArrayInitialized = False
            Else
                IsArrayInitialized = True
            End If
            Exit Function

        ElseIf VarType(arr) = vbArray + vbInteger Then
            Dim aInt() As Integer

            aInt = arr

            If ((Not aInt) = -1) Then
                IsArrayInitialized = False
            Else
                IsArrayInitialized = True
            End If
            Exit Function

        ElseIf VarType(arr) = vbArray + vbByte Then
            Dim aByte() As Byte

            aByte = arr

            If ((Not aByte) = -1) Then
                IsArrayInitialized = False
            Else
                IsArrayInitialized = True
            End If
            Exit Function

        ElseIf VarType(arr) = vbArray + vbLong Then
            Dim aLng() As Long

            aLng = arr

            If ((Not aLng) = -1) Then
                IsArrayInitialized = False
            Else
                IsArrayInitialized = True
            End If
            Exit Function

        End If


    End If
End Function

Public Sub DeleteArrayItem(ItemArray As Variant, ByVal ItemElement As Long)
    Dim i As Long

    If Not IsArray(ItemArray) Then
        Err.Raise 13, , "Type Mismatch"
        Exit Sub
    End If

    If ItemElement < LBound(ItemArray) Or ItemElement > UBound(ItemArray) Then
        Err.Raise 9, , "Subscript out of Range"
        Exit Sub
    End If

    For i = ItemElement To UBound(ItemArray) - 1
        ItemArray(i) = ItemArray(i + 1)
    Next
    On Error GoTo ErrorHandler:

    ReDim Preserve ItemArray(LBound(ItemArray) To UBound(ItemArray) - 1)

    Exit Sub
ErrorHandler:
'    ~~> An error will occur if array is fixed
    Err.Raise Err.Number, , "Array not resizable."
End Sub

Public Function GetDirectoryFromPath(FullPath As String) As String
    If InStr(FullPath, "\") <> 0 Then
        GetDirectoryFromPath = Left(FullPath, InStrRev(FullPath, "\"))
    End If
End Function

Public Function GetFilenameFromPath(FullPath As String, Optional bWithoutExt As Boolean = False) As String
    Dim strTmp As String

    If InStr(FullPath, "\") <> 0 Then
        strTmp = Right(FullPath, Len(FullPath) - InStrRev(FullPath, "\"))
    Else
        strTmp = FullPath
    End If

    If InStr(strTmp, ".") <> 0 Then
        If bWithoutExt Then
            strTmp = Left(strTmp, InStrRev(strTmp, ".") - 1)
        End If
    End If

    GetFilenameFromPath = strTmp
End Function

'Public Sub BSTR2LPSTR(ByVal BSTR As String, ByVal LPSTR As Long)
'    Dim arrSTR() As Byte
'    arrSTR() = StrConv(BSTR & vbNullChar, vbFromUnicode)
'    CopyMemory ByVal LPSTR, arrSTR(0), UBound(arrSTR) + 1
'End Sub
'
'Public Function LPSTR2BSTR(ByVal LPSTR As Long, Optional ByVal NullSearchDist As Long = 256) As String
'    Dim arrSTR() As Byte
'    Dim NullOffset As Long
'    Dim StrLen As Long
'
'    ReDim arrSTR(NullSearchDist - 1)
'    CopyMemory arrSTR(0), ByVal LPSTR, NullSearchDist
'
'    NullOffset = InStrB(1, arrSTR, StrConv(vbNullChar, vbFromUnicode)) - 1
'    Select Case NullOffset
'    Case -1
'        StrLen = NullSearchDist
'    Case 0
'        StrLen = 1
'    Case Else
'        StrLen = NullOffset
'    End Select
'    ReDim Preserve arrSTR(StrLen - 1)
'    LPSTR2BSTR = StrConv(arrSTR, vbUnicode)
'End Function

' Trim at first null Char
Public Function Trim0(sName As String) As String
'    Right trim string at first null.
    Dim x As Integer
    x = InStr(sName, vbNullChar)
    If x > 0 Then Trim0 = Left$(sName, x - 1) Else Trim0 = sName
End Function

Public Function IDEDebugMode() As Boolean
    On Error Resume Next
    Debug.Print 0 / 0
    IDEDebugMode = Err.Number <> 0
End Function

Public Sub DeleteStrArrayElemnt(ItemArray() As String, ByVal ItemElement As Long)
    Dim i As Long

    If ItemElement < LBound(ItemArray) Or ItemElement > UBound(ItemArray) Then
        Err.Raise 9, , "Subscript out of Range"
        Exit Sub
    End If

    For i = ItemElement To UBound(ItemArray) - 1
        ItemArray(i) = ItemArray(i + 1)
    Next
    On Error GoTo ErrorHandler

    If (UBound(ItemArray) - 1) < 0 Then
        Erase ItemArray
    Else
        ReDim Preserve ItemArray(LBound(ItemArray) To UBound(ItemArray) - 1)
    End If
    Exit Sub
ErrorHandler:
'    ~~> An error will occur if array is fixed
    Err.Raise Err.Number, , "Array not resizable."
End Sub

Public Sub InsertStrArrayElemnt(ItemArray() As String, ByVal ItemElement As Long, ByVal sValue As String)
    Dim i As Long

    If Not IsArrayInitialized(ItemArray) Then
        ReDim ItemArray(0)
    End If

    If ItemElement < LBound(ItemArray) Then
        Err.Raise 9, , "Subscript out of Range"
        Exit Sub
    End If

    If ItemElement > UBound(ItemArray) Then ItemElement = UBound(ItemArray)


    On Error GoTo ErrorHandler
    ReDim Preserve ItemArray(LBound(ItemArray) To UBound(ItemArray) + 1)
    On Error GoTo 0

    For i = UBound(ItemArray) To ItemElement + 1 Step -1
        ItemArray(i) = ItemArray(i - 1)
    Next

    ItemArray(ItemElement) = sValue

    Exit Sub
ErrorHandler:
'    ~~> An error will occur if array is fixed
    Err.Raise Err.Number, , "Array not resizable."
End Sub

Public Function TrimTab(ByVal Text As String) As String
'    Unicode-safe.
    Const WHITE_SPACE As String = vbTab

    If StrTrim(StrPtr(Text), StrPtr(WHITE_SPACE)) Then
        TrimTab = Left$(Text, lstrlen(StrPtr(Text)))
    Else
        TrimTab = Text
    End If
End Function

Public Sub Alert(ByVal s As String)
    MsgBox s, vbCritical, "Alert!"
End Sub
'
'Public Function Peek(ByVal lPtr As Long) As Long
'    Call CopyMemory(Peek, ByVal lPtr, 4)
'End Function
'
'Public Function Poke(address As Long, Value As Long)
'    CopyMemory ByVal address, Value, LenB(Value)
'End Function
'
'Public Function PeekB(address As Long) As Byte
'    Call CopyMemory(PeekB, ByVal address, Len(PeekB))
'End Function
'
'Public Function PokeB(address As Long, Value As Byte)
'    CopyMemory ByVal address, Value, LenB(Value)
'End Function
'
'Public Function ByteArrayToString(bytArray() As Byte, Optional iPos As Long = -1, Optional iLen As Long = -1) As String
'    Dim sAns As String
'    Dim iNullPos As String
'    Dim bTmp() As Byte
'
'    If iPos <> -1 And iLen <> -1 Then
'        If iPos < UBound(bytArray) And (iPos + iLen) < UBound(bytArray) Then
'            ReDim bTmp(iLen + 1)
''            copy the array
'            CopyMemory bTmp(0), bytArray(iPos), iLen
'            sAns = bTmp
'        End If
'    Else
'        sAns = bytArray
'    End If
'    iNullPos = InStr(sAns, Chr(0))
'    If iNullPos > 0 Then sAns = Left(sAns, iNullPos - 1)
'
'    ByteArrayToString = sAns
'
'End Function
'
'Public Function GetULongFromByteArr(s() As Byte, Pos As Long) As Variant
'    Dim Buffer As Long
'    CopyMemory Buffer, s(Pos), 4
'    If Buffer < 0 Then
'        GetULongFromByteArr = (Buffer And &H7FFFFFFF) + 2147483648#
'    Else
'        GetULongFromByteArr = Buffer
'    End If
'End Function
'
'Public Function GetLongFromByteArr(s() As Byte, Pos As Long) As Long
'    Dim Buffer As Long
'    CopyMemory Buffer, s(Pos), 4
'    GetLongFromByteArr = Buffer
'    Exit Function
'End Function
'Public Function GetUIntFromByteArr(s() As Byte, Pos As Long) As Long
'    Dim Buffer As Integer
'    CopyMemory Buffer, s(Pos), 2
'    If Buffer < 0 Then
'        GetUIntFromByteArr = (Buffer And &H7FFF) + &H8000
'    Else
'        GetUIntFromByteArr = Buffer
'    End If
'End Function
'Public Function GetIntFromByteArr(s() As Byte, Pos As Long) As Integer
'    Dim Buffer As Integer
'    CopyMemory Buffer, s(Pos), 2
'    GetIntFromByteArr = Buffer
'End Function
'
'Public Function GetDblFromByteArr(s() As Byte, Pos As Long) As Double
'    Dim Buffer As Double
'    CopyMemory Buffer, s(Pos), 8
'    GetDblFromByteArr = Buffer
'End Function
'
'Public Function Replace(ByRef Text As String, _
'  ByRef sOld As String, ByRef sNew As String, _
'  Optional ByVal Start As Long = 1, _
'  Optional ByVal Count As Long = 2147483647, _
'  Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
'  ) As String
'
'  ' (c) Jost Schwider, VB-Tec.de
'  If LenB(sOld) Then
'
'    If Compare = vbBinaryCompare Then
'      ReplaceBin Replace, Text, Text, _
'        sOld, sNew, Start, Count
'    Else
'      ReplaceBin Replace, Text, LCase$(Text), _
'        LCase$(sOld), sNew, Start, Count
'    End If
'
'  Else ' Suchstring ist leer:
'    Replace = Text
'  End If
'End Function
'
'Private Static Sub ReplaceBin(ByRef Result As String, _
'  ByRef Text As String, ByRef Search As String, _
'  ByRef sOld As String, ByRef sNew As String, _
'  ByVal Start As Long, ByVal Count As Long _
'  )
'
'  Dim TextLen As Long
'  Dim OldLen As Long
'  Dim NewLen As Long
'  Dim ReadPos As Long
'  Dim WritePos As Long
'  Dim CopyLen As Long
'  Dim Buffer As String
'  Dim BufferLen As Long
'  Dim BufferPosNew As Long
'  Dim BufferPosNext As Long
'
'  ' Ersten Treffer bestimmen:
'  If Start < 2 Then
'    Start = InStrB(Search, sOld)
'  Else
'    Start = InStrB(Start + Start - 1, Search, sOld)
'  End If
'  If Start Then
'
'    OldLen = LenB(sOld)
'    NewLen = LenB(sNew)
'    Select Case NewLen
'      Case OldLen ' einfaches Überschreiben:
'
'        Result = Text
'        For Count = 1 To Count
'          ' String "patchen":
'          MidB$(Result, Start) = sNew
'
'          ' Position aktualisieren:
'          Start = InStrB(Start + OldLen, Search, sOld)
'          If Start = 0 Then Exit Sub
'        Next Count
'        Exit Sub
'
'      Case Is < OldLen ' Ergebnis wird kürzer:
'
'        ' Buffer initialisieren:
'        TextLen = LenB(Text)
'        If TextLen > BufferLen Then
'          Buffer = Text
'          BufferLen = TextLen
'        End If
'
'        ' Ersetzen:
'        ReadPos = 1
'        WritePos = 1
'        For Count = 1 To Count
'          ' String "patchen":
'          CopyLen = Start - ReadPos
'          BufferPosNew = WritePos + CopyLen
'          MidB$(Buffer, WritePos) _
'            = MidB$(Text, ReadPos, CopyLen)
'          MidB$(Buffer, BufferPosNew) = sNew
'
'          ' Positionen aktualisieren:
'          WritePos = BufferPosNew + NewLen
'          ReadPos = Start + OldLen
'          Start = InStrB(ReadPos, Search, sOld)
'          If Start = 0 Then Exit For
'        Next Count
'
'        ' Ergebnis zusammenbauen:
'        If ReadPos > TextLen Then
'          Result = LeftB$(Buffer, WritePos - 1)
'        Else
'          MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
'          Result = LeftB$(Buffer, WritePos + _
'            LenB(Text) - ReadPos)
'        End If
'        Exit Sub
'
'     Case Else ' Ergebnis wird länger:
'
'        ' Buffer initialisieren:
'        TextLen = LenB(Text)
'        BufferPosNew = TextLen + NewLen
'        If BufferPosNew > BufferLen Then
'          Buffer = Space$(BufferPosNew)
'          BufferLen = LenB(Buffer)
'        End If
'
'        ' Ersetzung:
'        ReadPos = 1
'        WritePos = 1
'        For Count = 1 To Count
'          ' Positionen berechnen:
'          CopyLen = Start - ReadPos
'          BufferPosNew = WritePos + CopyLen
'          BufferPosNext = BufferPosNew + NewLen
'
'          ' Ggf. Buffer vergrößern:
'          If BufferPosNext > BufferLen Then
'            Buffer = Buffer & Space$(BufferPosNext)
'            BufferLen = LenB(Buffer)
'          End If
'
'          ' String "patchen":
'          MidB$(Buffer, WritePos _
'            ) = MidB$(Text, ReadPos, CopyLen)
'          MidB$(Buffer, BufferPosNew) = sNew
'
'          ' Positionen aktualisieren:
'          WritePos = BufferPosNext
'          ReadPos = Start + OldLen
'          Start = InStrB(ReadPos, Search, sOld)
'          If Start = 0 Then Exit For
'        Next Count
'
'        ' Ergebnis zusammenbauen:
'        If ReadPos > TextLen Then
'          Result = LeftB$(Buffer, WritePos - 1)
'        Else
'          BufferPosNext = WritePos + TextLen - ReadPos
'          If BufferPosNext < BufferLen Then
'            MidB$(Buffer, WritePos _
'              ) = MidB$(Text, ReadPos)
'            Result = LeftB$(Buffer, BufferPosNext)
'          Else
'            Result = LeftB$(Buffer, WritePos - 1) & _
'              MidB$(Text, ReadPos)
'          End If
'        End If
'        Exit Sub
'
'    End Select
'
'  Else ' Kein Treffer:
'    Result = Text
'  End If
'End Sub

'Helper function, to assign a function address (AddressOf) to a variable
Public Function AssignFuncAddrToVar(ByVal faddr As Long) As Long
  AssignFuncAddrToVar = faddr
End Function

Function FileToString(strFilename As String) As String
  iFile = FreeFile
  Open strFilename For Input As #iFile
    FileToString = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
End Function

Public Sub ExitApp(Optional uExitCode As Long = 0)

If IDEDebugMode Or uExitCode = 0 Then
    End
Else
    ExitProcess uExitCode
End If

End Sub
