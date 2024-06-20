VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMach4DocEmu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IMach4Emu
Implements IMyScriptObjectEmu

Private Sub IMach4Emu_CloseFile()

End Sub

Private Sub IMach4Emu_CycleStart()

End Sub

Private Sub IMach4Emu_FeedHold()

End Sub

Private Sub IMach4Emu_GetAxisPosition(Xpos As Double, Ypos As Double, Zpos As Double, Apos As Double, Bpos As Double, Cpos As Double)

End Sub

Private Sub IMach4Emu_GetLineNumber()

End Sub

Private Function IMach4Emu_GetRev() As String

End Function

Private Function IMach4Emu_GetScriptDispatch() As Object
        Dim II As IMyScriptObjectEmu
        Set II = Me

        Set IMach4Emu_GetScriptDispatch = II
End Function

Private Sub IMach4Emu_LoadGCodeFile(ByVal Filename As String)

End Sub

Private Sub IMach4Emu_RewindGcode()

End Sub

Private Sub IMach4Emu_SetFRO(ByVal FeedOverride As Double)

End Sub

Private Sub IMach4Emu_ShutDown()

End Sub



Private Sub Class_Initialize()
        If Mach3.g_bLoadProfile Then
            Dim xmlDoc As DOMDocument60
            
            If OpenProfile(xmlDoc, Mach3.g_MainFolder & Mach3.g_XMLProfile) Then
                
                
                
                
                Set xmlDoc = Nothing
            End If
        End If
End Sub

' Erweiterungen
Private Function OpenProfile(xmlDoc As DOMDocument60, sFile As String) As Boolean
    
    Set xmlDoc = New DOMDocument60
    xmlDoc.async = False
    If xmlDoc.Load(sFile) Then
        OpenProfile = True
    Else
        Set xmlDoc = Nothing
        OpenProfile = False
    End If
    
End Function

Private Function GetProfileParam(xmlDoc As DOMDocument60, sRootNode As String, sChildNode As String) As String
    Dim oNodeList As Object
    Dim oRootNode As Object
    Dim oChildNode As Object
    
    If Len(sRootNode) = 0 Then sRootNode = "Preferences"
    
    If Not xmlDoc Is Nothing Then
        If sRootNode = "" Then
            Set oNodeList = xmlDoc.getElementsByTagName(sChildNode)
            If oNodeList.length > 0 Then
                GetProfileParam = oNodeList.Item(0).Text
            End If
            Set oNodeList = Nothing
        Else
            Set oRootNode = xmlDoc.selectNodes("//" & sRootNode)
            
            If oRootNode.length > 0 Then
                
                Set oChildNode = oRootNode.Item(0).selectSingleNode(sChildNode)
                If Not oChildNode Is Nothing Then
                    GetProfileParam = oChildNode.Text
                End If
            End If
        End If
    End If
    
End Function

Private Sub IMyScriptObjectEmu_ActivateSignal(ByVal Signal As Integer)

End Sub

Private Function IMyScriptObjectEmu_AppendTeachFile(ByVal name As String) As Integer

End Function

Private Function IMyScriptObjectEmu_AskTextQuestion(ByVal Question As String) As String

End Function

Private Sub IMyScriptObjectEmu_ClearFiFo()

End Sub

Private Sub IMyScriptObjectEmu_CloseDigFile()

End Sub

Private Sub IMyScriptObjectEmu_CloseTeachFile()

End Sub

Private Sub IMyScriptObjectEmu_Code(ByVal Command As String)

End Sub

Private Function IMyScriptObjectEmu_CommandedFeed() As Double

End Function

Private Sub IMyScriptObjectEmu_CoupleSlave(ByVal state As Integer)

End Sub

Private Sub IMyScriptObjectEmu_DeActivateSignal(ByVal Signal As Integer)

End Sub

Private Sub IMyScriptObjectEmu_DoButton(ByVal ButtonNum As Integer)

End Sub

Private Sub IMyScriptObjectEmu_DoMenu(ByVal MenuIndex As Integer, ByVal MenuItem As Integer)

End Sub

Private Sub IMyScriptObjectEmu_DoOEMButton(ByVal Button As Integer)

End Sub

Private Sub IMyScriptObjectEmu_DoSpinCCW()

End Sub

Private Sub IMyScriptObjectEmu_DoSpinCW()

End Sub

Private Sub IMyScriptObjectEmu_DoSpinStop()

End Sub

Private Sub IMyScriptObjectEmu_EndTHC()

End Sub

Private Function IMyScriptObjectEmu_FeedRate() As Double

End Function

Private Function IMyScriptObjectEmu_Filename() As String

End Function

Private Function IMyScriptObjectEmu_FillFromCoil(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_FillFromHolding(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_FillFromInput(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_FillFromStatus(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_GetABSPostion(Axis As Variant) As Double

End Function

Private Function IMyScriptObjectEmu_GetACoor() As Double

End Function

Private Function IMyScriptObjectEmu_GetActiveProfileDir() As String

End Function

Private Function IMyScriptObjectEmu_GetActiveProfileName() As String

End Function

Private Function IMyScriptObjectEmu_GetActiveScreenSetName() As String

End Function

Private Sub IMyScriptObjectEmu_GetCoord(ByVal Question As String)

End Sub

Private Function IMyScriptObjectEmu_GetCurrentTool() As Integer

End Function

Private Function IMyScriptObjectEmu_GetDRO(ByVal dro As Integer) As Double

End Function

Private Function IMyScriptObjectEmu_GetDROString(ByVal nDRO As Integer) As Variant

End Function

Private Function IMyScriptObjectEmu_GetFiFoEntry() As String

End Function

Private Function IMyScriptObjectEmu_GetIJMode() As Integer

End Function

Private Function IMyScriptObjectEmu_GetInBit(ByVal index As Integer, ByVal bit As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_GetInput(ByVal word As Integer) As Variant

End Function

Private Function IMyScriptObjectEmu_GetIODevInput(ByVal DeviceID As Integer, ByVal InputNum As Integer) As Double

End Function

Private Function IMyScriptObjectEmu_GetIODevIOName(ByVal DeviceID As Integer, ByVal InputNumber As Integer) As String

End Function

Private Function IMyScriptObjectEmu_GetIODevName(ByVal DeviceNumber As Integer) As String

End Function

Private Function IMyScriptObjectEmu_GetLED(ByVal Led As Integer) As Boolean

End Function

Private Function IMyScriptObjectEmu_GetLoadedGCodeDir() As String

End Function

Private Function IMyScriptObjectEmu_GetLoadedGCodeFileName() As String

End Function

Private Function IMyScriptObjectEmu_GetMachVersion(major As Integer, minor As Integer, Build As Integer) As Boolean

End Function

Private Function IMyScriptObjectEmu_GetMainFolder() As String

End Function

Private Function IMyScriptObjectEmu_GetMasterInput(ByVal Param As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_GetMasterOutput(ByVal Param As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_GetMessage() As String

End Function

Private Function IMyScriptObjectEmu_GetMinPass() As Double

End Function

Private Function IMyScriptObjectEmu_GetModWord(ByVal index As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_GetMyWindowsHandle() As Long

End Function

Private Function IMyScriptObjectEmu_GetOEMDRO(ByVal dro As Integer) As Double

End Function

Private Function IMyScriptObjectEmu_GetOEMLed(ByVal Led As Integer) As Boolean

End Function

Private Function IMyScriptObjectEmu_GetPage() As Integer

End Function

Private Function IMyScriptObjectEmu_GetParam(ByVal Param As String) As Double

End Function

Private Function IMyScriptObjectEmu_GetPortByte(ByVal PortAddress As Integer) As Byte

End Function

Private Function IMyScriptObjectEmu_GetRPM() As Double

End Function

Private Function IMyScriptObjectEmu_GetSafeZ() As Double

End Function

Private Function IMyScriptObjectEmu_GetScale(ByVal Axis As Integer) As Double

End Function

Private Function IMyScriptObjectEmu_GetSelectedTool() As Integer

End Function

Private Function IMyScriptObjectEmu_GetSetupUnits() As Integer

End Function

Private Function IMyScriptObjectEmu_GetTimer(ByVal iTimer As Integer) As Double

End Function

Private Function IMyScriptObjectEmu_GetToolChangeStart(ByVal Axis As Integer) As Double

End Function

Private Function IMyScriptObjectEmu_GetToolDesc(ByVal Tool As Integer) As String

End Function

Private Function IMyScriptObjectEmu_GetToolParam(ByVal toolnum As Integer, ByVal Param As Integer) As Double

End Function

Private Function IMyScriptObjectEmu_GetTurretAng() As Double

End Function

Private Function IMyScriptObjectEmu_GetUserDRO(ByVal dro As Integer) As Double

End Function

Private Function IMyScriptObjectEmu_GetUserLabel(ByVal nLabel As Integer) As String

End Function

Private Function IMyScriptObjectEmu_GetUserLED(ByVal Led As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_GetVar(ByVal var As Integer) As Double

End Function

Private Function IMyScriptObjectEmu_GetXCoor() As Double

End Function

Private Function IMyScriptObjectEmu_GetYCoor() As Double

End Function

Private Function IMyScriptObjectEmu_GetZCoor() As Double

End Function

Private Sub IMyScriptObjectEmu_GotoSafeZ()

End Sub

Private Sub IMyScriptObjectEmu_HelpAbout()

End Sub

Private Function IMyScriptObjectEmu_IncludeTLOinZFromG31() As Boolean

End Function

Private Sub IMyScriptObjectEmu_InFeeds(ByVal iteration As Integer)

End Sub

Private Function IMyScriptObjectEmu_IsActive(ByVal Signal As Integer) As Byte

End Function

Private Function IMyScriptObjectEmu_IsDiameter() As Integer

End Function

Private Function IMyScriptObjectEmu_IsEstop() As Integer

End Function

Private Function IMyScriptObjectEmu_IsFirst() As Boolean

End Function

Private Function IMyScriptObjectEmu_IsLoading() As Integer

End Function

Private Function IMyScriptObjectEmu_IsMoving() As Integer

End Function

Private Function IMyScriptObjectEmu_IsOutputActive(ByVal Signal As Integer) As Boolean

End Function

Private Function IMyScriptObjectEmu_IsPeriodicScriptRunning(ByVal PathString As String) As Boolean

End Function

Private Sub IMyScriptObjectEmu_IsProbing()

End Sub

Private Function IMyScriptObjectEmu_IsSafeZ() As Integer

End Function

Private Function IMyScriptObjectEmu_IsStopped() As Integer

End Function

Private Sub IMyScriptObjectEmu_IsSuchSignal(ByVal Signal As Integer)

End Sub

Private Function IMyScriptObjectEmu_IsWiz() As Integer

End Function

Private Sub IMyScriptObjectEmu_JogOff(ByVal Axis As Integer)

End Sub

Private Sub IMyScriptObjectEmu_JogOn(ByVal Axis As Integer, ByVal dir As Integer)

End Sub

Private Sub IMyScriptObjectEmu_LoadFile(ByVal Filename As String)

End Sub

Private Sub IMyScriptObjectEmu_LoadLinTable()

End Sub

Private Sub IMyScriptObjectEmu_LoadRun(ByVal Filename As String)

End Sub

Private Sub IMyScriptObjectEmu_LoadStandardLayout()

End Sub

Private Sub IMyScriptObjectEmu_LoadTeachFile()

End Sub

Private Sub IMyScriptObjectEmu_LoadWizard(ByVal name As String)

End Sub

Private Sub IMyScriptObjectEmu_MachMSG(ByVal Msg As String, ByVal Title As String, ByVal iType As Integer)

End Sub

Private Function IMyScriptObjectEmu_MaxX() As Double

End Function

Private Function IMyScriptObjectEmu_MaxY() As Double

End Function

Private Sub IMyScriptObjectEmu_Message(ByVal Text As String)

End Sub

Private Function IMyScriptObjectEmu_MinX() As Double

End Function

Private Function IMyScriptObjectEmu_MinY() As Double

End Function

Private Function IMyScriptObjectEmu_ModGetInputWord(ByVal bit As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_nFmt(ByVal val As Double, ByVal DEC As Integer) As Double

End Function

Private Sub IMyScriptObjectEmu_NoDelaySpindle(ByVal StopForSpeedChange As Integer)

End Sub

Private Sub IMyScriptObjectEmu_NotifyPlugins(ByVal Message As Integer)

End Sub

Private Function IMyScriptObjectEmu_NumberPad(ByVal MessageString As String) As Boolean

End Function

Private Sub IMyScriptObjectEmu_OpenDigFile()

End Sub

Private Sub IMyScriptObjectEmu_OpenSubFile(ByVal Filename As String)

End Sub

Private Function IMyScriptObjectEmu_OpenSubroutineFile(ByVal name As String) As Integer

End Function

Private Sub IMyScriptObjectEmu_OpenTeachFile(ByVal name As String)

End Sub

Private Function IMyScriptObjectEmu_Param1() As Double

End Function

Private Function IMyScriptObjectEmu_Param2() As Double

End Function

Private Function IMyScriptObjectEmu_Param3() As Double

End Function

Private Sub IMyScriptObjectEmu_PlayWave(ByVal wavename As String)

End Sub

Private Function IMyScriptObjectEmu_ProgramSafetyLockout() As Boolean

End Function

Private Function IMyScriptObjectEmu_PutPortByte(ByVal PortAddress As Integer, ByVal data As Byte) As Integer

End Function

Private Function IMyScriptObjectEmu_Random() As Double

End Function

Private Sub IMyScriptObjectEmu_RefCombination(ByVal combo As Integer)

End Sub

Private Sub IMyScriptObjectEmu_ResetAxisSwap()

End Sub

Private Sub IMyScriptObjectEmu_ResetOutBit(ByVal index As Integer, ByVal bit As Integer)

End Sub

Private Sub IMyScriptObjectEmu_ResetTHC()

End Sub

Private Function IMyScriptObjectEmu_RetractMode() As Integer

End Function

Private Function IMyScriptObjectEmu_roun(val As Variant) As Double

End Function

Private Function IMyScriptObjectEmu_Round(val As Variant) As Double

End Function

Private Sub IMyScriptObjectEmu_RunFile()

End Sub

Private Sub IMyScriptObjectEmu_RunProgram(ByVal Program As String)

End Sub

Private Function IMyScriptObjectEmu_RunScript(ByVal macro As String) As Boolean

End Function

Private Sub IMyScriptObjectEmu_SaveWizard()

End Sub

Private Sub IMyScriptObjectEmu_SendFIFO(ByVal com As String)

End Sub

Private Sub IMyScriptObjectEmu_SendFIFObyte(ByVal data As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SendFIFOcr(ByVal data As String)

End Sub

Private Sub IMyScriptObjectEmu_SendFIFOword(ByVal data As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SendSerial(ByVal sData As String)

End Sub

Private Sub IMyScriptObjectEmu_SetButtonText(ByVal Text As String)

End Sub

Private Sub IMyScriptObjectEmu_SetCurrentTool(ByVal Tool As Integer)

End Sub

Private Function IMyScriptObjectEmu_SetDevOutput(ByVal DeviceID As Integer, ByVal OutputNum As Integer, ByVal Value As Integer) As Integer

End Function

Private Sub IMyScriptObjectEmu_SetDRO(ByVal dro As Integer, ByVal Value As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetFeedRate(ByVal Rate As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetFormula(ByVal Formula As String, ByVal Axis As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetHomannString(ByVal x As Integer, ByVal Y As Integer, ByVal Text As String)

End Sub

Private Sub IMyScriptObjectEmu_SetIJMode(ByVal mode As Integer)

End Sub

Private Function IMyScriptObjectEmu_SetInputData(ByVal InputNumber As Integer, ByVal PinNumber As Integer, ByVal ActiveLow As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_SetIODevDataPointers(ByVal DeviceID As Integer, ByVal PinNumber As Integer, ByVal SetValueDRO As Integer, ByVal SetOffsetDRO As Integer) As Integer

End Function

Private Function IMyScriptObjectEmu_SetMachPos(ByVal Axis As Integer, ByVal Value As Double) As Integer

End Function

Private Sub IMyScriptObjectEmu_SetMachZero(ByVal Axis As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetMasterInput(ByVal Param As Integer, ByVal Value As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetMasterOutput(ByVal Param As Integer, ByVal Value As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetModIOString(ByVal slave As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal OutText As String)

End Sub

Private Sub IMyScriptObjectEmu_SetModOutput(ByVal reg As Integer, ByVal Value As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetModPlugString(ByVal Message As String, ByVal cfg As Integer, ByVal address As Integer)

End Sub

Private Sub IMyScriptObjectEmu_setobj(ByVal thevar As String, ByVal thevalue As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetOEMDRO(ByVal dro As Integer, ByVal Value As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetOEMLED(ByVal Led As Integer, ByVal state As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetOutBit(ByVal index As Integer, ByVal bit As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetOutput(ByVal word As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetPage(ByVal page As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetParam(ByVal Param As String, ByVal Value As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetProbeActive(ByVal level As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetProbeState(ByVal state As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetPulley(ByVal pulley As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetSafeZ(ByVal SafeZ As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetScale(ByVal Axis As Integer, ByVal dScale As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetSpinSpeed(ByVal rpm As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetTicker(ByVal TickerNum As Integer, ByVal Message As String)

End Sub

Private Sub IMyScriptObjectEmu_SetTimer(ByVal iTimer As Integer)

End Sub

Private Function IMyScriptObjectEmu_SetToolDesc(ByVal ToolNumber As Integer, ByVal Description As String) As Boolean

End Function

Private Sub IMyScriptObjectEmu_SetToolParam(ByVal toolnum As Integer, ByVal Param As Integer, ByVal data As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetToolX(ByVal Value As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetToolZ(ByVal Value As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetTriggerMacro(ByVal macro As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetUserDRO(ByVal dro As Integer, ByVal Value As Double)

End Sub

Private Sub IMyScriptObjectEmu_SetUserLabel(ByVal LabelNum As Integer, ByVal LabelText As String)

End Sub

Private Sub IMyScriptObjectEmu_SetUserLED(ByVal Led As Integer, ByVal state As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SetVar(ByVal var As Integer, ByVal Value As Double)

End Sub

Private Sub IMyScriptObjectEmu_SingleVerify()

End Sub

Private Sub IMyScriptObjectEmu_SingleVerifyReport(ByVal Axis As Integer)

End Sub

Private Sub IMyScriptObjectEmu_Sleep(ByVal iTime As Integer)

End Sub

Private Sub IMyScriptObjectEmu_Speak(ByVal Text As String)

End Sub

Private Function IMyScriptObjectEmu_StartPeriodicScript(ByVal ScriptPath As String, ByVal UpdateTime As Double) As Boolean

End Function

Private Sub IMyScriptObjectEmu_StartTHC()

End Sub

Private Sub IMyScriptObjectEmu_StartVideo()

End Sub

Private Function IMyScriptObjectEmu_StopPeriodicScript(ByVal ScriptPath As String) As Boolean

End Function

Private Sub IMyScriptObjectEmu_StraightFeed(ByVal x As Double, ByVal Y As Double, ByVal Z As Double, ByVal a As Double, ByVal B As Double, ByVal C As Double)

End Sub

Private Sub IMyScriptObjectEmu_StraightTraverse(ByVal x As Double, ByVal Y As Double, ByVal Z As Double, ByVal a As Double, ByVal B As Double, ByVal C As Double)

End Sub

Private Sub IMyScriptObjectEmu_SwapAxis(ByVal Primary As Integer, ByVal Secondary As Integer)

End Sub

Private Sub IMyScriptObjectEmu_SystemWaitFor(ByVal Signal As Integer)

End Sub

Private Function IMyScriptObjectEmu_tChamfer() As Double

End Function

Private Function IMyScriptObjectEmu_tClearX() As Double

End Function

Private Function IMyScriptObjectEmu_tCutDepth() As Double

End Function

Private Function IMyScriptObjectEmu_tDepthLastPass() As Double

End Function

Private Function IMyScriptObjectEmu_tEndX() As Double

End Function

Private Function IMyScriptObjectEmu_tEndZ() As Double

End Function

Private Function IMyScriptObjectEmu_tGetCutType() As Integer

End Function

Private Function IMyScriptObjectEmu_tGetInfeedType() As Integer

End Function

Private Sub IMyScriptObjectEmu_THCOff()

End Sub

Private Sub IMyScriptObjectEmu_THCOn()

End Sub

Private Sub IMyScriptObjectEmu_ThreadDepth(ByVal iteration As Integer)

End Sub

Private Function IMyScriptObjectEmu_tInFeed() As Double

End Function

Private Function IMyScriptObjectEmu_tLead() As Double

End Function

Private Function IMyScriptObjectEmu_tMinDepth() As Double

End Function

Private Sub IMyScriptObjectEmu_ToggleScreens()

End Sub

Private Function IMyScriptObjectEmu_ToolLengthOffset() As Double

End Function

Private Function IMyScriptObjectEmu_tPasses() As Integer

End Function

Private Sub IMyScriptObjectEmu_tSetCutType(ByVal iType As Integer)

End Sub

Private Sub IMyScriptObjectEmu_tSetInFeedType(ByVal iType As Integer)

End Sub

Private Function IMyScriptObjectEmu_tSpring() As Integer

End Function

Private Function IMyScriptObjectEmu_tTaper() As Double

End Function

Private Sub IMyScriptObjectEmu_tTapers(ByVal iteration As Integer)

End Sub

Private Function IMyScriptObjectEmu_tXStart() As Double

End Function

Private Function IMyScriptObjectEmu_tZClear() As Double

End Function

Private Function IMyScriptObjectEmu_tZStart() As Double

End Function

Private Sub IMyScriptObjectEmu_VBTest()

End Sub

Private Sub IMyScriptObjectEmu_VerifyAxis(ByVal Silent As Long)

End Sub

Private Sub IMyScriptObjectEmu_WaitForPoll()

End Sub

Private Sub IMyScriptObjectEmu_ZeroTHC()

End Sub
