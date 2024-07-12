Attribute VB_Name = "modMain"
Option Explicit

' FormatMessage Konstanten
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H900
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

Private Const ERROR_PROCESS_ABORTED = &H42B
Private Const ERROR_NOINTERFACE = &H278
Private Const ERROR_BAD_COMMAND = &H16
Private Const ERROR_BAD_ARGUMENTS = &HA0

Private Type tCmdLineOption
  Option As String
  Parameter As String
End Type

Private Type tExpand
    nCurIndex As Integer
    nCurLevel As Integer
    asFiles() As String
End Type

Private Type tLines
    lLine As Long
    sFile As String
End Type

Private m_ExpandScript As tExpand
Private m_ctx As Long

Private m_objMach4 As IMach4Emu
Private m_objMScript As Object

Private m_yieldCounter As Long
Private m_maxYield As Long

Private m_Command As String ' command fro VSC => "RUN", "COMPILE", "CHECK"
Private m_ScriptFile As String ' Mach3 Script Filename
Private m_OutputFolder As String ' Output Folder for compiling

Private m_ErrLine As Long ' Error Line num on compiling
Private m_errMessage As String ' Error message on compiling
Private m_errType As String ' Error Type on Compiling
Private m_errFile As String ' file in which error occurred

Private g_asBuffer() As String 'Global buffer that contains the source code lines during compiling and syntax checking

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Function ParseCommandLine(ret_aCmdLineOptions() As tCmdLineOption, Optional MaxArgs) As Boolean
    'Declare variables.
    Dim C, CmdLine, CmdLnLen, InArg, i, NumArgs, bOpt, bParam, QParam, bNot1stChr
    'See if MaxArgs was provided.
    If IsMissing(MaxArgs) Then MaxArgs = 10
    'Make array of the correct size.
    ReDim ret_aCmdLineOptions(1 To MaxArgs)
    NumArgs = 0: InArg = False
    'Get command line arguments.
    CmdLine = Command()
    CmdLnLen = Len(CmdLine)
    'Go thru command line one character
    'at a time.
                
    'Pipe.stdout CmdLine
    If CmdLnLen > 0 Then
    For i = 1 To CmdLnLen
        C = Mid(CmdLine, i, 1)
        'Ignore Quote
        If C = """" Then GoTo Continue_I
        'Test for space or tab
        If (C <> " " And C <> vbTab) Then
            'Neither space nor tab.
            'Test if already in argument.
            If Not InArg Then
                If C = "-" Then
                    'New argument begins.
                    'Test for too many arguments.
                    If NumArgs = MaxArgs Then Exit For
                    NumArgs = NumArgs + 1
                    InArg = True
                    bOpt = True
                    bParam = False
                    bNot1stChr = False
                    GoTo Continue_I
                End If
            End If
            If bOpt And C = ":" Then
                bParam = True
                bOpt = False
                bNot1stChr = False
                GoTo Continue_I
            End If
            'Concatenate character to current argument.
            If bOpt Then
                ret_aCmdLineOptions(NumArgs).Option = ret_aCmdLineOptions(NumArgs).Option & UCase(C)
            End If
            If bParam Then
                If Not bNot1stChr Then
                    bNot1stChr = True
                    If C = "'" Then
                        QParam = True
                        GoTo Continue_I
                    End If
                Else
                    If C = "'" Then
                        QParam = False
                        GoTo Continue_I
                    End If
                End If
                ret_aCmdLineOptions(NumArgs).Parameter = ret_aCmdLineOptions(NumArgs).Parameter & C
            End If
        Else
            'Found a space or tab.
            'Set InArg flag to False if not in quoted Param
            If Not QParam Then
                InArg = False
                bOpt = False
                bParam = False
                bNot1stChr = False
            Else
                If bParam Then
                    ret_aCmdLineOptions(NumArgs).Parameter = ret_aCmdLineOptions(NumArgs).Parameter & C
                End If
            End If
        End If
Continue_I:
        
    Next i
    'Resize array just enough to hold arguments.
    ReDim Preserve ret_aCmdLineOptions(1 To NumArgs)
    ParseCommandLine = True
    Else
    ParseCommandLine = False
    End If
End Function

Private Function LoadMacroFile(sFileName As String, sBuffer As String) As Boolean
    Dim iFile As Integer
    
    iFile = FreeFile
    If dir$(sFileName, vbNormal) <> "" Then
        Open sFileName For Binary As #iFile
        If LOF(iFile) <> 0 Then
            sBuffer = Space((LOF(iFile)))
            Get #iFile, , sBuffer
            Close #iFile
            LoadMacroFile = True
        End If
    End If
End Function

Private Sub RemoveExpandBlocks(sBuffer As String)
    Dim asBuffer() As String
    Dim i As Long
    Dim bIgnore As Boolean
    Dim sPos As Integer
    Dim ePos As Integer
    
    If Len(sBuffer) Then
        asBuffer = GetArrayFromText(sBuffer)
        sBuffer = ""
        
        For i = 0 To UBound(asBuffer)
            If Not bIgnore Then
                If InStr(1, Trim(TrimTab(asBuffer(i))), "'#Expand Start of", vbTextCompare) = 1 Then
                    bIgnore = 1
                Else
                    sBuffer = sBuffer + asBuffer(i) + vbCrLf
                End If
                
            ElseIf bIgnore And InStr(1, Trim(TrimTab(asBuffer(i))), "'#Expand End of", vbTextCompare) = 1 Then
                sPos = InStr(asBuffer(i), "'#Expand End of") + 15
                ePos = InStr(asBuffer(i), "##", sPos) - 2
                sBuffer = sBuffer + Trim(Mid(asBuffer(i), sPos, ePos - sPos)) + vbCrLf
                bIgnore = 0
            End If
            
        Next
    End If
End Sub

Private Function ExpandScript(asScript() As String, Optional recCall As Boolean = False)
    
    Dim bInfinityLoop As Boolean
    Dim sBuffer As String
    Dim asBuffer() As String
    Dim v71 As Byte
    Dim i As Integer
    Dim j As Integer
    Dim nSPos As Integer
    Dim nEPos As Integer
    Dim sFileName As String
    Dim sExpandBlock As String
    Dim iFile As Integer
    Dim bFileLoaded  As Boolean
    Dim bQFN As Boolean
    Dim bExpand As Boolean
            
    ExpandScript = ERROR_PROCESS_ABORTED
    
    If Not recCall Then
        Erase m_ExpandScript.asFiles
        m_ExpandScript.nCurIndex = 0
    End If
    
    If Not IsArrayInitialized(asScript) Then Exit Function
    
    If (m_ExpandScript.nCurLevel = 0) Then
        ReDim m_ExpandScript.asFiles(0)
    End If
    m_ExpandScript.nCurLevel = m_ExpandScript.nCurLevel + 1
    
    i = 0
    
    Do
        If (InStr(1, Trim(TrimTab(asScript(i))), "#expand """, vbTextCompare) = 1) Then
            nSPos = InStr(1, asScript(i), "#expand """, vbTextCompare)
            nEPos = InStr(nSPos + 9, asScript(i), """", vbTextCompare)
            bQFN = True
            bExpand = True
        ElseIf (InStr(1, Trim(TrimTab(asScript(i))), "#expand <", vbTextCompare) = 1) Then
            nSPos = InStr(1, asScript(i), "#expand <", vbTextCompare)
            nEPos = InStr(nSPos + 9, asScript(i), ">", vbTextCompare)
            bQFN = False
            bExpand = True
        Else
            bExpand = False
        End If
        If bExpand Then
            sExpandBlock = Mid(asScript(i), nSPos, (nEPos + 1) - nSPos)
            
            If bQFN Then
                sFileName = Mid(asScript(i), nSPos + 9, nEPos - (nSPos + 9))
            Else
                sFileName = Mach3.g_MainFolder + "\ScreenSetMacros\" + Mach3.g_CurrentScreenSet + "\" + Mid(asScript(i), nSPos + 9, nEPos - (nSPos + 9)) + ".m1s"
            End If
            
            
            If Not FileExists(sFileName) Then
                Pipe.stderr App.EXEName & " Error: " & asScript(i) & vbCrLf
                Pipe.stderr App.EXEName & " File '" & sFileName & "' not found !" & vbCrLf
                ExpandScript = &H2
                Exit Function
            End If
            
            bFileLoaded = LoadMacroFile(sFileName, sBuffer)
            
            If (bFileLoaded) Then
                
                asBuffer = GetArrayFromText(sBuffer)
                
                If IsArrayInitialized(asBuffer) Then
                    If Len(asBuffer(0)) Then
                        v71 = 1
                    End If
                End If
                If (v71) Then
                    
                    ReDim Preserve m_ExpandScript.asFiles(m_ExpandScript.nCurIndex) '
                    m_ExpandScript.asFiles(m_ExpandScript.nCurIndex) = sFileName
                    
                    bInfinityLoop = False
                    For j = 0 To m_ExpandScript.nCurIndex - 1 '
                        
                        If (m_ExpandScript.asFiles(j) = sFileName) Then
                            bInfinityLoop = True
                            Exit For
                        End If
                    Next
                    m_ExpandScript.nCurIndex = m_ExpandScript.nCurIndex + 1
                    If (bInfinityLoop) Then
                        
                        m_ExpandScript.nCurIndex = m_ExpandScript.nCurIndex - 1
                        m_ExpandScript.nCurLevel = m_ExpandScript.nCurLevel - 1
                        Pipe.stderr ("#Expand is calling itself (infinity loop)" & vbCrLf)
                        ExpandScript = False
                        Exit Function
                    End If
                    DeleteStrArrayElemnt asScript, i
                    InsertStrArrayElemnt asScript, i, "'#Expand Start of " & sExpandBlock & vbTab & "file:'" & sFileName & "' ##"
                    For j = 0 To UBound(asBuffer)
                        InsertStrArrayElemnt asScript(), i + j + 1, asBuffer(j)
                    Next j
                    
                    If (Not ExpandScript(asScript, True)) Then
                        ExpandScript = False
                        Exit Function
                    End If
                    InsertStrArrayElemnt asScript, i + j + 1, "'#Expand End of " & sExpandBlock & vbTab & " ##"
                    m_ExpandScript.nCurIndex = m_ExpandScript.nCurIndex - 1
                End If
                
            Else
                Pipe.stderr ("#Expand File Not found at: " & sFileName & vbCrLf)
            End If
            
        End If
        i = i + 1
        
    Loop While (i <= UBound(asScript))
    
    m_ExpandScript.nCurLevel = m_ExpandScript.nCurLevel - 1
    If Not m_ExpandScript.nCurLevel Then
        
        i = i - 1
        If Len(asScript(i)) Then
            If Right(asScript(i), 1) = vbCr Then
                
                asScript(i) = asScript(i) & vbLf
            ElseIf Right(asScript(i), 1) <> vbLf Then
                asScript(i) = asScript(i) & vbCrLf
            End If
        End If

    End If
    
    ExpandScript = True
    
End Function

Private Function CompileScript() As Long
    Dim rc As Long
    Dim hwnd As Long
    Dim sBuffer As String
    Dim i As Integer
    Dim RetVal As Long
    Erase g_asBuffer() ' clear Buffer
    
    CompileScript = ERROR_PROCESS_ABORTED
    
    If m_ctx Then
        EnableAPI.enaFreeContext (m_ctx)
        m_ctx = 0
    End If
    
    rc = EnableAPI.enaCreateContext("VSCCE CompileScript", App.hInstance, 0&, m_ctx)
    'Pipe.stdout "enaCreateContext:" & CStr(rc) & vbCrLf
    If rc = 0 Then
        
        m_ErrLine = 0
        m_errType = 0
        m_errMessage = ""
        
        If VBHelp.FileExists(m_ScriptFile) Then
            sBuffer = VBHelp.FileToString(m_ScriptFile)
        Else
            Pipe.stderr ("Error: file '" & m_ScriptFile & "' does not exist" & vbCrLf)
            CompileScript = &H2
            EnableAPI.enaFreeContext (m_ctx)
            Exit Function
        End If
        
        g_asBuffer = GetArrayFromText(sBuffer)
        
        If IsArrayInitialized(g_asBuffer) Then
            If ExpandScript(g_asBuffer()) = ERROR_PROCESS_ABORTED Then
                EnableAPI.enaFreeContext m_ctx
                Exit Function
            End If
                
            sBuffer = Join(g_asBuffer(), vbLf)
            Sleep (100)
                                      
        End If
        
        rc = EnableAPI.enaSetOption(m_ctx, ENOPT_CALL_APP_WITH_CTX, 1)   ' Include implicit context parameter in App calls

        rc = EnableAPI.enaAddTypeLibFileRef(m_ctx, Mach3.g_MainFolder & "\Mach3.exe", 0)

        For i = 0 To UBound(Mach3.g_asDefinitions())
            rc = EnableAPI.enaAppendText(m_ctx, Mach3.g_asDefinitions(i))
        Next i
        
        rc = EnableAPI.enaSetDefaultDispatch(m_ctx, 0, m_objMScript)

        
        rc = EnableAPI.enaAppendText(m_ctx, sBuffer)

    End If
    If Not rc Then
        rc = EnableAPI.enaRegisterOutputCallback(m_ctx, AddressOf CompileOutputCallback)

        rc = EnableAPI.enaRegisterGetProcAddrCallback(m_ctx, AddressOf MyGetProcAddr)

        rc = EnableAPI.enaCompile(m_ctx)

    Else
        sBuffer = Space(512)
        RetVal = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, rc, &H400, sBuffer, Len(sBuffer), 0&)
        If RetVal Then
            Pipe.stderr (sBuffer)
        End If
        
    End If
    If rc <> 0 Then
        EnableAPI.enaFreeContext (m_ctx)
    End If
            
    
    CompileScript = rc
    
End Function

Private Function GetArrayFromText(ByVal sText As String) As String()
    Dim asBuffer() As String
    If Len(sText) Then
        If InStr(1, sText, vbCrLf) Then
            asBuffer = Split(sText, vbCrLf)
        ElseIf InStr(1, sText, vbCr) Then
            asBuffer = Split(sText, vbCr)
        ElseIf InStr(1, sText, vbLf) Then
            asBuffer = Split(sText, vbLf)
        Else
            ReDim asBuffer(0)
            asBuffer(0) = sText
        End If
        
        If asBuffer(UBound(asBuffer)) = "" Then ReDim Preserve asBuffer(UBound(asBuffer) - 1)
        GetArrayFromText = asBuffer
    End If
    
End Function

Private Sub AssignCmdLineOptionsToVars(aOptions() As tCmdLineOption)
    Dim i As Integer

    For i = 1 To UBound(aOptions())

        Select Case aOptions(i).Option
            Case "CMD"
                m_Command = aOptions(i).Parameter
            Case "MACH3DIR"
                Mach3.g_MainFolder = aOptions(i).Parameter
            Case "PROFILE"
                Mach3.g_XMLProfile = aOptions(i).Parameter
            Case "SCREENSET"
                Mach3.g_CurrentScreenSet = aOptions(i).Parameter
            Case "SRC"
                m_ScriptFile = aOptions(i).Parameter
            Case "OUTPUTFOLDER"
                m_OutputFolder = aOptions(i).Parameter
        End Select

    Next
End Sub
Private Sub ResetYieldCounter(): m_yieldCounter = 0: End Sub

Private Function MyOutputCallback(pOutputInfo As OutputInfo) As Long
'    suppress all errors
    Pipe.stderr "Error" & vbCrLf & "Buff:" & CStr(pOutputInfo.bufNum) & vbCrLf & "Line:" & CStr(pOutputInfo.lineNum) & vbCrLf & "offset:" & CStr(pOutputInfo.offset) & vbCrLf
    MyOutputCallback = 0
End Function

Private Function CompileOutputCallback(pOutputInfo As OutputInfo) As Long
    Dim k As Long
    Dim lLine As Long
    Dim iIndex As Integer
    Dim lLines() As tLines
    Dim sFile As String
    Dim nSPos As Integer
    Dim nEPos As Integer
    Dim sMsg As String
    
    If pOutputInfo.scode <> 0 Then
        m_ErrLine = pOutputInfo.lineNum
        m_errMessage = pOutputInfo.errMessage
        m_errType = pOutputInfo.scode
        m_errFile = m_ScriptFile
        
        ' check in which file the error occurred (main / or include file)
        ' and adjust the line number if necessary
        lLine = CLng(m_ErrLine - 1)
        iIndex = -1
        If lLine > -1 And lLine <= UBound(g_asBuffer()) Then
            ReDim lLines(0)
            iIndex = 0
            lLines(iIndex).lLine = 0
            lLines(iIndex).sFile = m_ScriptFile
            For k = 0 To lLine
                If InStr(g_asBuffer(k), "'#Expand Start of") <> 0 Then
                    lLines(iIndex).lLine = lLines(iIndex).lLine + 1
                    iIndex = iIndex + 1
                    ReDim Preserve lLines(iIndex)
                    lLines(iIndex).lLine = 0
                    nSPos = InStr(1, g_asBuffer(k), vbTab & "file:'", vbTextCompare)
                    nEPos = InStr(nSPos + 7, g_asBuffer(k), "'", vbTextCompare)
                    lLines(iIndex).sFile = Mid(g_asBuffer(k), nSPos + 7, nEPos - (nSPos + 7))
                ElseIf InStr(g_asBuffer(k), "'#Expand End of") <> 0 Then
                    iIndex = iIndex - 1
                    '      : g_asBuffer(6) : "'#Expand End of #expand <test>  ##" : String
                Else
                    lLines(iIndex).lLine = lLines(iIndex).lLine + 1
                End If
            Next k
            m_errFile = lLines(iIndex).sFile
            m_ErrLine = lLines(iIndex).lLine
            If UBound(lLines()) > 0 Then
                If InStr(1, m_errMessage, " - ") = 0 Then
                    sMsg = Replace(m_errMessage, vbLf, " ") & vbLf
                Else
                    sMsg = Mid(m_errMessage, InStr(1, m_errMessage, " - ") + 3)
                End If
                m_errMessage = " Error on line: " & CStr(m_ErrLine) & " - " & sMsg
            End If
            Erase lLines()
        Else
            m_errFile = m_ScriptFile
        End If
               
    Else
        m_ErrLine = 0
        m_errType = 0
        m_errMessage = ""
        m_errFile = ""
    End If
    CompileOutputCallback = 0
End Function

Private Function YieldCallback(ByVal Ctx As Long) As Long
    
   If m_yieldCounter = m_maxYield Then
    YieldCallback = -1
   End If
   m_yieldCounter = m_yieldCounter + 1
End Function


Private Function MyGetProcAddr(ByVal Ctx As Long, ByVal funcName As String) As Long

        MyGetProcAddr = 0

End Function

Private Sub RegServeEXE(ByVal Path As String, mode As Boolean)
    On Error GoTo RegServeEXE_Error
    
    If mode Then
        ShellExecute 0, "RunAs", Path, " /regserver", "", 0
    Else
        ShellExecute 0, "RunAs", Path, " /unregserver", "", 0
    End If
    
    On Error GoTo 0
    Exit Sub
RegServeEXE_Error:
    MsgBox ("Error! ActiveX registration." & vbCrLf & Err.Number & ": " & Err.Description)
End Sub


Sub Main()
    On Error GoTo Main_Error
    
    Dim aCmdLineOptions() As tCmdLineOption
    Dim ErrCode As Long
    Dim dstFile As String
    Dim AppPath As String
    
    If Right(App.Path, 1) <> "\" Then
        AppPath = App.Path & "\"
    Else
        AppPath = App.Path
    End If

    Call RegServeEXE(AppPath & "vscce.exe", True)
     
    If Not ParseCommandLine(aCmdLineOptions()) Then
        MsgBox "No Arguments"
        VBHelp.ExitApp ERROR_BAD_ARGUMENTS
    End If
    
    AssignCmdLineOptionsToVars aCmdLineOptions()
    'Init Mach3/Script Interface
    Set m_objMach4 = New CMach4DocEmu
    If Not m_objMach4 Is Nothing Then
        Set m_objMScript = m_objMach4.GetScriptDispatch()
    Else
        Pipe.stderr (App.EXEName & " Error: can't init IMach4")
        VBHelp.ExitApp ERROR_NOINTERFACE
    End If
    
    Mach3.InitDefinitions
    
    Select Case UCase(m_Command)
        Case "COMPILE"
            If Len(m_ScriptFile) Then
                ChDir Mach3.g_MainFolder
                Pipe.stdout (App.EXEName & " Info: start compiling Script '" & m_ScriptFile & "'" & vbCrLf)
                ErrCode = CompileScript()
                If ErrCode = 0 Then
                    ErrCode = VBHelp.MakePath(m_OutputFolder & "\")
                    If VBHelp.FolderExists(m_OutputFolder) Then
                        dstFile = m_OutputFolder + "\" & VBHelp.GetFilenameFromPath(m_ScriptFile, True) & ".mcc"
                        ErrCode = EnableAPI.enaSaveCompiledCode(m_ctx, dstFile)
                        If ErrCode = 0 Then
                            Pipe.stdout (App.EXEName & " Info: compiling complete" & vbCrLf)
                            Pipe.stdout (App.EXEName & " Info: file saved to: " & dstFile & vbCrLf)
                        End If
                    Else
                        Pipe.stderr (App.EXEName & " Error: could not access output folder '" & m_OutputFolder & "'" & vbCrLf)
                    End If
                    EnableAPI.enaFreeContext (m_ctx)
                    m_ctx = 0
                Else
                    Pipe.stderr (App.EXEName & " Error: compiling Script " & m_ScriptFile & "'" & vbCrLf)
                    If Len(m_errMessage) Or m_ErrLine <> 0 Then
                        If Len(m_errMessage) Or m_ErrLine <> 0 Then
                            'Pipe.stderr (App.EXEName & " Error on line: " & m_ErrLine & " - syntax error (" & m_errFile & ")")
                            Pipe.stderr (App.EXEName & " Error: syntax error found " & vbCrLf)
                        Else
                            Pipe.stderr (App.EXEName & " " & m_errMessage & "(" & m_errFile & ")" & vbCrLf)
                        End If
                    End If
                End If
            Else
                ErrCode = &H2
            End If
        Case "RUN"
            Pipe.stdout (App.EXEName & " Info: Running the script is not yet supported" & vbCrLf)
        Case "CHECK"
            If Len(m_ScriptFile) Then
                ChDir Mach3.g_MainFolder
                Pipe.stdout (App.EXEName & " Info: check Script '" & m_ScriptFile & "' for Syntax Errors" & vbCrLf)
                ErrCode = CompileScript()
                If ErrCode = 0 Then
                    Pipe.stdout (App.EXEName & " Info: check complete, no errors found !" & vbCrLf)
                Else
                    If Len(m_errMessage) Or m_ErrLine <> 0 Then
                        
                        If Len(m_errMessage) = 0 Or InStr(m_errMessage, "Compile error:") = 1 Then
                            'Pipe.stderr (App.EXEName & " Error on line: " & m_ErrLine & " - syntax error (" & m_errFile & ")")
                            Pipe.stderr (App.EXEName & " Error: syntax error found " & vbCrLf)
                        Else
                            Pipe.stderr (App.EXEName & " " & m_errMessage & "(" & m_errFile & ")" & vbCrLf)
                        End If
                    End If
                End If
            Else
                ErrCode = &H2
            End If
        Case Else
            Pipe.stderr (App.EXEName & " Error: unknown option (cmd)" & vbCrLf)
            ErrCode = ERROR_BAD_COMMAND
    End Select
    
    Set m_objMScript = Nothing
    Set m_objMach4 = Nothing
    Erase aCmdLineOptions()
    
    VBHelp.ExitApp ErrCode
    
Main_Error:
    Pipe.stderr (App.EXEName & "(Main) Error :" & Err.Number & " (" & Err.Description & ")" & vbCrLf)
    VBHelp.ExitApp Err.Number
End Sub
