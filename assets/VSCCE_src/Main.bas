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

Private Type tCmdLineOption
  Option As String
  Parameter As String
End Type

Private Type tExpand
    nCurIndex As Integer
    nCurLevel As Integer
    asFiles() As String
End Type

Private tExpandScript As tExpand
Private m_ctx As Long

Private m_objMach4 As IMach4Emu
Private m_objMScript As IMyScriptObjectEmu

Private m_yieldCounter As Long
Private m_maxYield As Long

Private m_Command As String ' command fro VSC => "RUN", "COMPILE", "CHECK"
Private m_ScriptFile As String ' Mach3 Script Filename
Private m_OutputFolder As String ' Output Folder for compiling

Private m_ErrLine As Long ' Error Line num on compiling
Private m_errMessage As String ' Error message on compiling
Private m_errType As String ' Error Type on Compiling

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long




Private Sub ParseCommandLine(ret_aCmdLineOptions() As tCmdLineOption, Optional MaxArgs)
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
End Sub

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

Private Function ExpandScript(asScript() As String)
    
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
    
    If Not IsArrayInitialized(asScript) Then Exit Function
    
    If (tExpandScript.nCurLevel = 0) Then
        ReDim tExpandScript.asFiles(0)
        
    End If
    tExpandScript.nCurLevel = tExpandScript.nCurLevel + 1
    
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
                Pipe.stderr App.EXEName & " Error: " & asScript(i)
                Pipe.stderr App.EXEName & " File '" & sFileName & "' not found !"
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
                    
                    ReDim Preserve tExpandScript.asFiles(tExpandScript.nCurIndex) ';sub_40B132((void *)(4 * nCurIndex + 6986664), &v72);
                    tExpandScript.asFiles(tExpandScript.nCurIndex) = sFileName
                    
                    bInfinityLoop = False
                    For j = 0 To tExpandScript.nCurIndex - 1 ';;For ( j = 0; j < nCurIndex - 1; ++j )
                        
                        If (tExpandScript.asFiles(j) = sFileName) Then
                            bInfinityLoop = True
                            Exit For
                        End If
                    Next
                    tExpandScript.nCurIndex = tExpandScript.nCurIndex + 1
                    If (bInfinityLoop) Then
                        
                        tExpandScript.nCurIndex = tExpandScript.nCurIndex - 1
                        tExpandScript.nCurLevel = tExpandScript.nCurLevel - 1
                        Alert ("#Expand is calling itself (infinity loop)")
                        ExpandScript = False
                        Exit Function
                    End If
                    DeleteStrArrayElemnt asScript, i
                    InsertStrArrayElemnt asScript, i, "'#Expand Start of " & sExpandBlock & vbTab & " ##"
                    For j = 0 To UBound(asBuffer)
                        InsertStrArrayElemnt asScript(), i + j + 1, asBuffer(j)
                    Next j
                    
                    If (Not ExpandScript(asScript)) Then
                        ExpandScript = False
                        Exit Function
                    End If
                    InsertStrArrayElemnt asScript, i + j + 2, "'#Expand End of " & sExpandBlock & vbTab & " ##"
                    tExpandScript.nCurIndex = tExpandScript.nCurIndex - 1
                End If
                
            Else
                Alert ("#Expand File Not found at:" & vbCrLf & vbCrLf & sFileName)
            End If
            
        End If
        i = i + 1
        
    Loop While (i <= UBound(asScript))
    
    tExpandScript.nCurLevel = tExpandScript.nCurLevel - 1
    If Not tExpandScript.nCurLevel Then
        
        i = i - 1
        If Len(asScript(i)) Then
            If Right(asScript(i), 1) = vbCr Then
                
                asScript(i) = asScript(i) & vbLf
            ElseIf Right(asScript(i), 1) <> vbLf Then
                asScript(i) = asScript(i) & vbCrLf
            End If
        End If
        Erase tExpandScript.asFiles
        tExpandScript.nCurIndex = 0
    End If
    
    ExpandScript = True
    
End Function

Private Function CompileScript() As Long
    Dim rc As Long
    Dim hwnd As Long
    Dim asBuffer() As String
    Dim sBuffer As String
    Dim i As Integer
    Dim RetVal As Long
    
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
        
        asBuffer = GetArrayFromText(sBuffer)
        
        If IsArrayInitialized(asBuffer) Then
            If ExpandScript(asBuffer()) = ERROR_PROCESS_ABORTED Then
                EnableAPI.enaFreeContext m_ctx
                Exit Function
            End If
            'Pipe.stdout "Script Expanded:" & CStr(rc) & vbCrLf
                
            sBuffer = Join(asBuffer(), vbCrLf)
            Sleep (100)
              
            
        End If
        
        rc = EnableAPI.enaSetOption(m_ctx, ENOPT_CALL_APP_WITH_CTX, True)   ' Include implicit context parameter in App calls
        'Pipe.stdout "enaSetOption:" & CStr(rc) & vbCrLf
        rc = EnableAPI.enaAddTypeLibFileRef(m_ctx, Mach3.g_MainFolder & "\Mach3.exe", 0)
        'Pipe.stdout "enaAddTypeLibFileRef:" & CStr(rc) & vbCrLf
        For i = 0 To UBound(Mach3.g_asDefinitions())
            rc = EnableAPI.enaAppendText(m_ctx, Mach3.g_asDefinitions(i))
        Next i
        'Pipe.stdout "enaAppendText:" & CStr(rc) & vbCrLf
        
        rc = EnableAPI.enaSetDefaultDispatch(m_ctx, 0, m_objMScript)
        'Pipe.stdout "enaSetDefaultDispatch:" & CStr(rc) & vbCrLf
        
        rc = EnableAPI.enaAppendText(m_ctx, sBuffer)
        'Pipe.stdout "enaAppendText:" & CStr(rc) & vbCrLf
    End If
    If Not rc Then
        rc = EnableAPI.enaRegisterOutputCallback(m_ctx, AddressOf CompileOutputCallback)
        'Pipe.stdout "enaRegisterOutputCallback:" & CStr(rc) & vbCrLf
        rc = EnableAPI.enaRegisterGetProcAddrCallback(m_ctx, AddressOf MyGetProcAddr)
        'Pipe.stdout "enaRegisterGetProcAddrCallback:" & CStr(rc) & vbCrLf
        rc = EnableAPI.enaCompile(m_ctx)
        'Pipe.stdout "enaCompile:" & CStr(rc) & vbCrLf
    Else
        sBuffer = Space(512)
        RetVal = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, rc, &H400, sBuffer, Len(sBuffer), 0&)
        If RetVal Then
            Pipe.stderr (sBuffer)
        Else
            Pipe.stderr ("Error running script")
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
                'MsgBox m_Command
            Case "MACH3DIR"
                Mach3.g_MainFolder = aOptions(i).Parameter
                'MsgBox Mach3.g_MainFolder
            Case "PROFILE"
                Mach3.g_XMLProfile = aOptions(i).Parameter
            Case "SCREENSET"
                Mach3.g_CurrentScreenSet = aOptions(i).Parameter
            Case "SRC"
                m_ScriptFile = aOptions(i).Parameter
                'MsgBox m_ScriptFile
            Case "OUTPUTFOLDER"
                m_OutputFolder = aOptions(i).Parameter
        End Select

    Next
End Sub
Private Sub ResetYieldCounter(): m_yieldCounter = 0: End Sub

Private Function MyOutputCallback(pOutputInfo As OutputInfo) As Long
'    suppress all errors
    MsgBox "Error" & vbCrLf & "Buff:" & CStr(pOutputInfo.bufNum) & vbCrLf & "Line:" & CStr(pOutputInfo.lineNum) & vbCrLf & "offset:" & CStr(pOutputInfo.offset)
    MyOutputCallback = 0
End Function

Private Function CompileOutputCallback(pOutputInfo As OutputInfo) As Long
    'Dim Result As VbMsgBoxResult
    'Result = MsgBox(pOutputInfo.errMessage, IIf(pOutputInfo.outputType <> 0, &H30, 0), App.Title)
    If pOutputInfo.scode <> 0 Then
        m_ErrLine = pOutputInfo.lineNum
        m_errMessage = pOutputInfo.errMessage
        m_errType = pOutputInfo.scode
    Else
        m_ErrLine = 0
        m_errType = 0
        m_errMessage = ""
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


Sub Main()
    On Error GoTo Main_Error
    
    Dim aCmdLineOptions() As tCmdLineOption
    Dim ErrCode As Long
    Dim dstFile As String
    
    ParseCommandLine aCmdLineOptions()
    
    AssignCmdLineOptionsToVars aCmdLineOptions()
    
    'Init Mach3/Script Interface
    Set m_objMach4 = New IMach4Emu
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
                    Pipe.stdout (App.EXEName & " CompileScript '" & ErrCode & "'" & vbCrLf)
                    ErrCode = VBHelp.MakePath(m_OutputFolder & "\")
                    If VBHelp.FolderExists(m_OutputFolder) Then
                        dstFile = m_OutputFolder + "\" & VBHelp.GetFilenameFromPath(m_OutputFolder, True) & ".mcc"
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
                        Pipe.stderr (App.EXEName & " " & m_errMessage)
                    End If
                End If
            Else
                ErrCode = &H2
            End If
        Case "RUN"
            Pipe.stdout (App.EXEName & " Info: Running the script is not yet supported" & vbCrLf)
        Case "CHECK"
            Pipe.stdout (App.EXEName & " Info: Checking the script is not yet supported" & vbCrLf)
        Case Else
            Pipe.stderr (App.EXEName & " Error: unknown option (cmd)" & vbCrLf)
            ErrCode = ERROR_BAD_COMMAND
    End Select
    
    Set m_objMScript = Nothing
    Set m_objMach4 = Nothing
    Erase aCmdLineOptions()
    
    VBHelp.ExitApp ErrCode
    
Main_Error:
    Pipe.stderr (App.EXEName & " Error :" & Err.Number & " (" & Err.Description & ")" & vbCrLf)
    VBHelp.ExitApp Err.Number
End Sub
