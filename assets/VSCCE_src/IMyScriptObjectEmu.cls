VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMyScriptObjectEmu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub StraightFeed(ByVal x As Double, ByVal Y As Double, ByVal Z As Double, ByVal a As Double, ByVal B As Double, ByVal C As Double)
        
End Sub
Public Sub HelpAbout()
    
End Sub
Public Function GetSafeZ() As Double

End Function
Public Sub SetSafeZ(ByVal SafeZ As Double)

End Sub
Public Sub SetCurrentTool(ByVal Tool As Integer)
End Sub
Public Function GetSelectedTool() As Integer

End Function
Public Function GetToolChangeStart(ByVal Axis As Integer) As Double

End Function
Public Sub StraightTraverse(ByVal x As Double, ByVal Y As Double, ByVal Z As Double, ByVal a As Double, ByVal B As Double, ByVal C As Double)

End Sub
Public Function ToolLengthOffset() As Double

End Function
Public Function CommandedFeed() As Double

End Function
Public Sub SetFeedRate(ByVal Rate As Double)

End Sub
Public Sub ActivateSignal(ByVal Signal As Integer)

End Sub
Public Function IsActive(ByVal Signal As Integer) As Byte
    

End Function
Public Sub DeActivateSignal(ByVal Signal As Integer)

End Sub
Public Sub SystemWaitFor(ByVal Signal As Integer)
End Sub
Public Function Param1() As Double

End Function
Public Function Param2() As Double

End Function
Public Function Param3() As Double

End Function
Public Sub VerifyAxis(ByVal Silent As Long)

End Sub
Public Function GetVar(ByVal var As Integer) As Double

    
End Function
Public Sub SetVar(ByVal var As Integer, ByVal Value As Double)
    
End Sub
Public Sub OpenDigFile()

End Sub
Public Sub CloseDigFile()

End Sub
Public Sub THCOn()

End Sub
Public Sub THCOff()

End Sub
Public Sub Code(ByVal Command As String)
    
End Sub
Public Function GetScale(ByVal Axis As Integer) As Double

End Function
Public Sub SetScale(ByVal Axis As Integer, ByVal dScale As Double)

End Sub
Public Sub SendSerial(ByVal sData As String)

End Sub
Public Function GetPortByte(ByVal PortAddress As Integer) As Byte

End Function
Public Function PutPortByte(ByVal PortAddress As Integer, ByVal data As Byte) As Integer

End Function
Public Function IsMoving() As Integer

End Function
Public Sub ResetTHC()
End Sub
Public Function GetParam(ByVal Param As String) As Double

End Function
Public Sub SetParam(ByVal Param As String, ByVal Value As Double)
        
End Sub
Public Sub setobj(ByVal thevar As String, ByVal thevalue As Double)

End Sub
Public Function GetCurrentTool() As Integer

End Function
Public Sub DoOEMButton(ByVal Button As Integer)

End Sub
Public Sub LoadRun(ByVal Filename As String)

End Sub
Public Sub DoButton(ByVal ButtonNum As Integer)

End Sub
Public Function GetLED(ByVal Led As Integer) As Boolean

End Function
Public Sub SetOEMLED(ByVal Led As Integer, ByVal state As Integer)

End Sub
Public Function GetOEMLed(ByVal Led As Integer) As Boolean

End Function
Public Function GetOEMDRO(ByVal dro As Integer) As Double

End Function
Public Function GetDRO(ByVal dro As Integer) As Double

End Function
Public Sub SetDRO(ByVal dro As Integer, ByVal Value As Double)

End Sub
Public Sub SetOEMDRO(ByVal dro As Integer, ByVal Value As Double)
    
End Sub
Public Sub GetCoord(ByVal Question As String)
End Sub
Public Function GetXCoor() As Double
End Function
Public Function GetYCoor() As Double
End Function
Public Function GetZCoor() As Double
End Function
Public Function GetACoor() As Double
End Function
Public Function IsFirst() As Boolean
End Function
Public Sub SetMachZero(ByVal Axis As Integer)
End Sub
Public Sub RunFile()
End Sub
Public Sub SetUserDRO(ByVal dro As Integer, ByVal Value As Double)
        
End Sub
Public Sub SetUserLED(ByVal Led As Integer, ByVal state As Integer)
End Sub
Public Function GetUserDRO(ByVal dro As Integer) As Double
   
End Function
Public Function GetUserLED(ByVal Led As Integer) As Integer

End Function
Public Function tXStart() As Double
End Function
Public Function tZStart() As Double
End Function
Public Function tEndX() As Double
End Function
Public Function tEndZ() As Double
End Function
Public Function tClearX() As Double
End Function
Public Function tLead() As Double
End Function
Public Function tSpring() As Integer
End Function
Public Function tPasses() As Integer
End Function
Public Function tChamfer() As Double
End Function
Public Function tTaper() As Double
End Function
Public Function tInFeed() As Double
End Function
Public Function tDepthLastPass() As Double
End Function
Public Function IsLoading() As Integer
End Function
Public Function GetABSPostion(Axis) As Double
End Function
Public Function GetRPM() As Double

End Function
Public Sub DoSpinCW()

End Sub
Public Sub DoSpinCCW()

End Sub
Public Sub DoSpinStop()

End Sub
Public Sub SetSpinSpeed(ByVal rpm As Double)

End Sub
Public Sub InFeeds(ByVal iteration As Integer)
End Sub
Public Sub ThreadDepth(ByVal iteration As Integer)
End Sub
Public Sub tTapers(ByVal iteration As Integer)
End Sub
Public Sub SetTicker(ByVal TickerNum As Integer, ByVal Message As String)

End Sub
Public Sub SetUserLabel(ByVal LabelNum As Integer, ByVal LabelText As String)

End Sub
Public Sub RefCombination(ByVal combo As Integer)

End Sub
Public Sub IsSuchSignal(ByVal Signal As Integer)

End Sub
Public Sub OpenTeachFile(ByVal name As String)

End Sub
Public Sub CloseTeachFile()

End Sub
Public Sub SetButtonText(ByVal Text As String)
End Sub
Public Sub LoadStandardLayout()
End Sub
Public Sub LoadTeachFile()
End Sub
Public Function IsOutputActive(ByVal Signal As Integer) As Boolean

End Function
Public Sub SingleVerify()

End Sub
Public Function GetPage() As Integer

End Function
Public Sub SetPage(ByVal page As Integer)

End Sub
Public Sub ToggleScreens()
End Sub
Public Sub PlayWave(ByVal wavename As String)
End Sub
Public Sub Speak(ByVal Text As String)
End Sub
Public Sub Message(ByVal Text As String)
End Sub
Public Function IsStopped() As Integer
End Function
Public Sub SaveWizard()
End Sub
Public Sub SingleVerifyReport(ByVal Axis As Integer)
End Sub
Public Sub SetIJMode(ByVal mode As Integer)
End Sub
Public Function GetIJMode() As Integer
End Function
Public Function GetMinPass() As Double
End Function
Public Function AppendTeachFile(ByVal name As String) As Integer
End Function
Public Function Random() As Double
End Function
Public Function IsDiameter() As Integer
End Function
Public Function ModGetInputWord(ByVal bit As Integer) As Integer
End Function
Public Function FillFromCoil(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer
End Function
Public Function FillFromStatus(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer
End Function
Public Function FillFromHolding(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer
End Function
Public Function FillFromInput(ByVal slave As Integer, ByVal startAddress As Integer, ByVal nBytes As Integer) As Integer
End Function
Public Function GetModWord(ByVal index As Integer) As Integer
End Function
Public Function tMinDepth() As Double
End Function
Public Function tGetCutType() As Integer
End Function
Public Function tGetInfeedType() As Integer
End Function
Public Sub tSetCutType(ByVal iType As Integer)
End Sub
Public Sub tSetInFeedType(ByVal iType As Integer)
End Sub
Public Function OpenSubroutineFile(ByVal name As String) As Integer
End Function
Public Function tZClear() As Double
End Function
Public Function FeedRate() As Double
End Function
Public Function tCutDepth() As Double
End Function
Public Function Filename() As String
End Function
Public Function MinX() As Double
End Function
Public Function MaxX() As Double
End Function
Public Function MinY() As Double
End Function
Public Function MaxY() As Double
End Function
Public Function RetractMode() As Integer
End Function
Public Sub GotoSafeZ()
End Sub
Public Sub SetPulley(ByVal pulley As Integer)
End Sub
Public Sub LoadWizard(ByVal name As String)
End Sub
Public Sub SetOutput(ByVal word As Integer)
End Sub
Public Function GetInput(ByVal word As Integer)
End Function
Public Sub SetModOutput(ByVal reg As Integer, ByVal Value As Integer)
End Sub
Public Sub LoadFile(ByVal Filename As String)

End Sub
Public Sub SetTriggerMacro(ByVal macro As Integer)
End Sub
Public Sub SetHomannString(ByVal x As Integer, ByVal Y As Integer, ByVal Text As String)
End Sub
Public Sub JogOn(ByVal Axis As Integer, ByVal dir As Integer)
End Sub
Public Sub JogOff(ByVal Axis As Integer)

End Sub
Public Sub OpenSubFile(ByVal Filename As String)

End Sub
Public Sub SetToolX(ByVal Value As Double)

End Sub
Public Sub SetToolZ(ByVal Value As Double)

End Sub
Public Function GetTurretAng() As Double
End Function
Public Function GetFiFoEntry() As String

End Function
Public Sub RunProgram(ByVal Program As String)

End Sub
Public Sub SendFIFO(ByVal com As String)

End Sub
Public Function Round(val) As Double

End Function
Public Function roun(val) As Double

End Function
Public Function nFmt(ByVal val As Double, ByVal DEC As Integer) As Double

End Function
Public Sub ZeroTHC()
End Sub
Public Function IsWiz() As Integer

End Function
Public Sub WaitForPoll()

End Sub
Public Sub SetModIOString(ByVal slave As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal OutText As String)
End Sub
Public Sub Sleep(ByVal iTime As Integer)
End Sub
Public Sub SendFIFOcr(ByVal data As String)
End Sub
Public Sub SendFIFObyte(ByVal data As Integer)
End Sub
Public Sub SendFIFOword(ByVal data As Integer)
End Sub
Public Sub SetTimer(ByVal iTimer As Integer)
End Sub
Public Function GetTimer(ByVal iTimer As Integer) As Double
End Function
Public Function AskTextQuestion(ByVal Question As String) As String
End Function
Public Function GetDROString(ByVal nDRO As Integer)
End Function
Public Sub LoadLinTable()
End Sub
Public Sub SwapAxis(ByVal Primary As Integer, ByVal Secondary As Integer)
End Sub
Public Sub ResetAxisSwap()
End Sub
Public Sub SetFormula(ByVal Formula As String, ByVal Axis As Integer)
End Sub
Public Sub IsProbing()
End Sub
Public Function GetInBit(ByVal index As Integer, ByVal bit As Integer) As Integer
End Function
Public Sub SetOutBit(ByVal index As Integer, ByVal bit As Integer)
End Sub
Public Sub ResetOutBit(ByVal index As Integer, ByVal bit As Integer)
End Sub
Public Sub SetProbeActive(ByVal level As Integer)
End Sub
Public Sub StartVideo()
End Sub
Public Sub VBTest()
End Sub
Public Sub StartTHC()
End Sub
Public Sub EndTHC()
End Sub
Public Sub CoupleSlave(ByVal state As Integer)
End Sub
Public Function GetToolParam(ByVal toolnum As Integer, ByVal Param As Integer) As Double
End Function
Public Sub SetToolParam(ByVal toolnum As Integer, ByVal Param As Integer, ByVal data As Double)
End Sub
Public Sub SetProbeState(ByVal state As Integer)
End Sub
Public Function GetMainFolder() As String

End Function
Public Sub DoMenu(ByVal MenuIndex As Integer, ByVal MenuItem As Integer)
End Sub
Public Function GetMasterInput(ByVal Param As Integer) As Integer
End Function
Public Function GetMasterOutput(ByVal Param As Integer) As Integer
End Function
Public Sub SetMasterOutput(ByVal Param As Integer, ByVal Value As Integer)
End Sub
Public Sub SetMasterInput(ByVal Param As Integer, ByVal Value As Integer)
End Sub
Public Function GetUserLabel(ByVal nLabel As Integer) As String

End Function
Public Function GetToolDesc(ByVal Tool As Integer) As String
End Function
Public Sub NotifyPlugins(ByVal Message As Integer)
End Sub
Public Function IsEstop() As Integer
End Function
Public Function IsSafeZ() As Integer
End Function
Public Sub MachMSG(ByVal Msg As String, ByVal Title As String, ByVal iType As Integer)
End Sub
Public Function GetMessage() As String
End Function
Public Sub SetModPlugString(ByVal Message As String, ByVal cfg As Integer, ByVal address As Integer)
End Sub
Public Sub ClearFiFo()
End Sub
Public Function GetMachVersion(ByRef major As Integer, ByRef minor As Integer, ByRef Build As Integer) As Boolean

End Function
Public Function SetToolDesc(ByVal ToolNumber As Integer, ByVal Description As String) As Boolean
End Function
Public Function RunScript(ByVal macro As String) As Boolean
  
End Function
Public Function GetActiveProfileName() As String

End Function
Public Function NumberPad(ByVal MessageString As String) As Boolean
End Function
Public Function GetMyWindowsHandle() As Long

End Function
Public Function GetLoadedGCodeDir() As String
End Function
Public Function GetLoadedGCodeFileName() As String
End Function
Public Function GetActiveProfileDir() As String
End Function
Public Function GetActiveScreenSetName() As String
End Function
Public Function IncludeTLOinZFromG31() As Boolean
End Function
Public Function ProgramSafetyLockout() As Boolean
End Function
Public Function GetSetupUnits() As Integer
End Function
Public Function StartPeriodicScript(ByVal ScriptPath As String, ByVal UpdateTime As Double) As Boolean
End Function
Public Function StopPeriodicScript(ByVal ScriptPath As String) As Boolean
End Function
Public Function IsPeriodicScriptRunning(ByVal PathString As String) As Boolean
End Function
Public Function GetIODevName(ByVal DeviceNumber As Integer) As String

End Function
Public Function SetDevOutput(ByVal DeviceID As Integer, ByVal OutputNum As Integer, ByVal Value As Integer) As Integer

End Function
Public Function GetIODevInput(ByVal DeviceID As Integer, ByVal InputNum As Integer) As Double

End Function
Public Function GetIODevIOName(ByVal DeviceID As Integer, ByVal InputNumber As Integer) As String

End Function
Public Function SetMachPos(ByVal Axis As Integer, ByVal Value As Double) As Integer

End Function
Public Function SetInputData(ByVal InputNumber As Integer, ByVal PinNumber As Integer, ByVal ActiveLow As Integer) As Integer

End Function
Public Function SetIODevDataPointers(ByVal DeviceID As Integer, ByVal PinNumber As Integer, ByVal SetValueDRO As Integer, ByVal SetOffsetDRO As Integer) As Integer

End Function
Public Sub NoDelaySpindle(ByVal StopForSpeedChange As Integer)

End Sub

