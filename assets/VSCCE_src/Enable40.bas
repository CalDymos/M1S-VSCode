Attribute VB_Name = "EnableAPI"

Option Explicit
' Constants and Declare statements for API calls to Cypress Enable
' Error return codes
Public Const ENS_FUNCNOTDEFINED = &H8007007F    'Sub or function not defined
Public Const ENS_RUNTIMEERROR = &H8007065B      ' Runtime error
' Informational return codes
Public Const ENS_END_INSTRUCTION = &H40201      ' Encountered End instruction
Public Const ENS_HIT_BREAKPOINT = &H40202       ' Encountered breakpoint


' The following structure is DWORD aligned (sizeof is 44)
Type OutputInfo
    context As Long                 ' Context in which error or print statment occurred
    outputType As Long              ' Type 0 indicates Print statement, 1 indicates
                                    ' general error, 2 compile error, 3 runtime error
    scode As Long                   ' Bit encoded HRESULT (see codes above)
    sStr As String                  ' Extra info (eg. function name), optional
    bufNum As Long                  ' Buffer number for syntax error
    lineNum As Long                 ' Line number for syntax error and for runtime error
    offset As Integer               ' Offset of token for syntax error
    tokenLen As Integer             ' Token length for syntax error
    runtimeErr As Long              ' Err.Number for runtime errors (see errors.txt)
    errMessage As String            ' Display error message (error type, line #, and text)
    bOnErrorStmt As Long            ' True if there is an OnError statement in effect for this scope
    pExcepInfo As Long              ' Pointer to EXCEPINFO structure returned from IDispatch::Invoke
End Type

' Return codes for OutputCallback function
Const EN_DEFAULTACTION = 0   ' Continue run using default handling
Const EN_ENDRUN = 1          ' End script run immediately regardless of handler in script
Const EN_RESUMENEXT = 2      ' Simulate OnError Resume Next, regardless of OnError status
Const EN_DEBUGBREAK = 3      ' Debug break
Const EN_DISPLAYOUTPUT = 4   ' Enable should display output

Declare Sub enaSetSymbolTableSize Lib "enable40.dll" (ByVal nSize As Long)
Declare Sub enaVersion Lib "enable40.dll" (ByRef major As Long, ByRef minor As Long)

Declare Function enaCreateContext Lib "enable40.dll" Alias "enaCreateContextA" (ByVal contextName As String, ByVal hInst As Long, ByVal hwnd As Long, pEnaCtx As Long) As Long

Declare Function enaRegisterGetProcAddrCallback Lib "enable40.dll" Alias "enaRegisterGetProcAddrCallbackW" (ByVal context As Long, ByVal pFcn As Long) As Long
Declare Function enaRegisterOutputCallback Lib "enable40.dll" Alias "enaRegisterOutputCallbackW" (ByVal context As Long, ByVal pFcn As Long) As Long
Declare Function enaRegisterUndefFuncCallback Lib "enable40.dll" Alias "enaRegisterUndefFuncCallbackW" (ByVal context As Long, ByVal pFcn As Long) As Long
Declare Function enaRegisterUndefVarCallback Lib "enable40.dll" Alias "enaRegisterUndefVarCallbackW" (ByVal context As Long, ByVal pFcn As Long) As Long
Declare Function enaRegisterParseFuncCallback Lib "enable40.dll" Alias "enaRegisterParseFuncCallbackW" (ByVal context As Long, ByVal pFcn As Long) As Long
Declare Function enaRegisterDimVarCallback Lib "enable40.dll" Alias "enaRegisterDimVarCallbackW" (ByVal context As Long, ByVal pFcn As Long) As Long
Declare Function enaRegisterYieldCallback Lib "enable40.dll" (ByVal context As Long, ByVal pFcn As Long) As Long

Public Const ENOPT_CALL_APP_WITH_CTX = 1
Public Const ENOPT_ALLOW_IMMEDIATE_CODE = 2
Public Const ENOPT_CACHE_DISPIDS = 4
Public Const ENOPT_ONERROR_CALLBACK = 8
Public Const ENOPT_RELAX_STRING_CONVERSION = 16   ' Not recommended
Public Const ENOPT_COMPCODE_SAVE_LOCAL_VARS = 32 '
Public Const ENOPT_NO_IMPLICIT_END_STATEMENT = 64

Declare Sub enaSetUserParam Lib "enable40.dll" (ByVal context As Long, ByVal lParam As Long)
Declare Sub enaSetDialogFont Lib "enable40.dll" Alias "enaSetDialogFontA" (ByVal context As Long, ByVal nFontSize As Integer, ByVal fontName As String)

Declare Function enaGetUserParam Lib "enable40.dll" (ByVal context As Long) As Long

Declare Function enaSetOption Lib "enable40.dll" (ByVal context As Long, ByVal lOption As Long, ByVal lParam As Long) As Long

Declare Function enaLoadCompiledCode Lib "enable40.dll" Alias "enaLoadCompiledCodeA" (ByVal context As Long, ByVal name As String) As Long

Declare Function enaAddTypeLibFileRef Lib "enable40.dll" Alias "enaAddTypeLibFileRefA" (ByVal context As Long, ByVal szFile As String, ByVal dwFlags As Long) As Long

Declare Function enaAddFile Lib "enable40.dll" Alias "enaAddFileA" (ByVal context As Long, ByVal Filename As String) As Long
Declare Function enaAppendText Lib "enable40.dll" Alias "enaAppendTextA" (ByVal context As Long, ByVal sText As String) As Long
Declare Function enaAppendTextLP Lib "enable40.dll" Alias "enaAppendTextA" (ByVal context As Long, ByVal lpsText As Long) As Long
Declare Function enaAppendEditWindow Lib "enable40.dll" (ByVal context As Long, ByVal hwnd As Long) As Long

Declare Function enaCompile Lib "enable40.dll" (ByVal context As Long) As Long
Declare Function enaSaveCompiledCode Lib "enable40.dll" Alias "enaSaveCompiledCodeA" (ByVal context As Long, ByVal name As String) As Long
Declare Function enaGetCompiledCodeSize Lib "enable40.dll" (ByVal context As Long, lpSize As Long)
Declare Function enaGetCompiledCode Lib "enable40.dll" (ByVal context As Long, ByVal lpcodeBuf As Long, ByVal bufSize As Long)

Declare Function enaEnumValidTokens Lib "enable40.dll" Alias "enaEnumValidTokensW" (ByVal context As Long, ByVal pFcn As Long, ByVal lParam As Long, dwFlags As Long) As Long

Declare Function enaSetGlobalIntegerVar Lib "enable40.dll" Alias "enaSetGlobalIntegerVarA" (ByVal context As Long, ByVal varName As String, ByVal intVar As Integer) As Long
Declare Function enaSetGlobalLongVar Lib "enable40.dll" Alias "enaSetGlobalLongVarA" (ByVal context As Long, ByVal varName As String, ByVal longVar As Long) As Long
Declare Function enaSetGlobalDoubleVar Lib "enable40.dll" Alias "enaSetGlobalDoubleVarA" (ByVal context As Long, ByVal varName As String, ByVal dblVar As Double) As Long
Declare Function enaSetGlobalStringVar Lib "enable40.dll" Alias "enaSetGlobalStringVarA" (ByVal context As Long, ByVal varName As String, ByVal strVar As String) As Long
Declare Function enaSetGlobalObjectVar Lib "enable40.dll" Alias "enaSetGlobalObjectVarA" (ByVal context As Long, ByVal varName As String, ByVal objVar As Object) As Long
Declare Function enaSetGlobalVariantVar Lib "enable40.dll" Alias "enaSetGlobalVariantVarA" (ByVal context As Long, ByVal varName As String, ByVal varVAr As Variant) As Long
Declare Function enaSetGlobalVariantArrayVar Lib "enable40.dll" Alias "enaSetGlobalArrayVarA" (ByVal context As Long, ByVal varName As String, ByRef arrVar() As Variant) As Long
Declare Function TestArray Lib "enable40.dll" Alias "enaSetGlobalArrayVarA" (ByVal context As Long, ByVal varName As String, x() As Any) As Long

Declare Function enaSetGlobalArrayRef Lib "enable40.dll" Alias "enaSetGlobalArrayRefA" (ByVal context As Long, ByVal varName As String, ByRef arrName() As Integer) As Long
Declare Function enaSetGlobalVariantArrayRef Lib "enable40.dll" Alias "enaSetGlobalArrayRefA" (ByVal context As Long, ByVal varName As String, ByRef arrVar() As Variant) As Long
Declare Function enaSetGlobalVariantRef Lib "enable40.dll" Alias "enaSetGlobalVariantRefA" (ByVal context As Long, ByVal varName As String, ByRef varVAr As Object) As Long
Declare Function enaSetErrorCode Lib "enable40.dll" Alias "enaSetErrorCodeW" (ByVal context As Long, ByVal scode As Long, ByVal runtimeErr As Integer, ByVal ErrStr As String) As Long

Declare Function enaSetDefaultDispatch Lib "enable40.dll " (ByVal context As Long, ByVal uModul As Long, ByVal pCtx As Object) As Long

Declare Function enaGetScriptDispatch Lib "enable40.dll" (ByVal Ctx As Long, obj As Object) As Long
Declare Function enaExecute Lib "enable40.dll" (ByVal context As Long) As Long
Declare Function enaCallFunction Lib "enable40.dll" Alias "enaCallFunctionA" (ByVal context As Long, ByVal funcName As String, Optional args As Any = vbNull, Optional RetVal As Any = vbNull) As Long
Declare Function enaDebugFunction Lib "enable40.dll" Alias "enaDebugFunctionA" (ByVal context As Long, ByVal funcName As String, Optional args As Any = vbNull, Optional RetVal As Any = vbNull, Optional pBufNum As Long = vbNull, Optional pLineNum As Long = vbNull)

Declare Function enaDebugGo Lib "enable40.dll" (ByVal context As Long, bufNum As Long, lineNum As Long) As Long

Declare Function enaDebugStep Lib "enable40.dll" (ByVal context As Long, bufNum As Long, lineNum As Long) As Long
Declare Function enaDebugStepInto Lib "enable40.dll" (ByVal context As Long, bufNum As Long, lineNum As Long) As Long

Type tRuntimeInfo
    cbSize As Long ' Set equal to sizeof( ENRUNTIMEINFO )
    nInst As Long '
    nBuf As Long '
    nLine As Long '
End Type

Declare Function enaGetRuntimeInfo Lib "enable40.dll" (ByVal context As Long, pRuntimeInfo As tRuntimeInfo, dwFlags As Long) As Long

Declare Function enaGetVarValueString Lib "enable40.dll" Alias "enaGetVarValueStringA" (ByVal context As Long, ByVal varName As String, ByVal strVar As Long, ByVal max As Integer) As Long

Declare Function enaGetGlobalIntegerVar Lib "enable40.dll" Alias "enaGetGlobalIntegerVarA" (ByVal context As Long, ByVal funcName As String, intVar As Integer) As Long
Declare Function enaGetGlobalLongVar Lib "enable40.dll" Alias "enaGetGlobalLongVarA" (ByVal context As Long, ByVal varName As String, longVar As Long) As Long
Declare Function enaGetGlobalSingleVar Lib "enable40.dll" Alias "enaGetGlobalSingleVarA" (ByVal context As Long, ByVal varName As String, ByRef snglVar As Single) As Long
Declare Function enaGetGlobalDoubleVar Lib "enable40.dll" Alias "enaGetGlobalDoubleVarA" (ByVal context As Long, ByVal varName As String, dblVarVar As Double) As Long
Declare Function enaGetGlobalStringVar Lib "enable40.dll" Alias "enaGetGlobalStringVarA" (ByVal context As Long, ByVal varName As String, ByVal strVar As String, ByVal max As Long) As Long
Declare Function enaGetGlobalBooleanVar Lib "enable40.dll" Alias "enaGetGlobalBooleanVarA" (ByVal context As Long, ByVal varName As String, ByRef boolVar As Boolean) As Long
Declare Function enaGetGlobalVariantVar Lib "enable40.dll" Alias "enaGetGlobalVariantVarA" (ByVal context As Long, ByVal varName As String, ByRef varVAr As Variant) As Long
Declare Function enaGetGlobalArrayElement Lib "enable40.dll" Alias "enaGetGlobalArrayElementA" (ByVal context As Long, ByVal arrName As String, ByVal arrIndices As Long, ByRef pvData As Any) As Long

Declare Function enaGetVarRef Lib "enable40.dll" (ByVal context As Long, ByVal varName As String, ByRef pVar As Variant, ByVal dwFlags As Long) As Long

Declare Function enaFreeContext Lib "enable40.dll" (ByVal context As Long) As Long

' The following structure is DWORD aligned (sizeof is 52)
Type CustomEdit
    cbSize As Long                  ' Set equal to sizeof( ENCUSTOMEDIT )
    pszContextName As String        ' Parameter to enaCreateContext
    lUserParam As Long              ' Parameter to enaSetUserParam
    dwOptions As Long               ' Parameter to enaSetOption( ctx, dwOptions, TRUE )
    pfnGetProcAddr As Long          ' Parameter to enaRegisterGetProcAddrCallback
    pfnOutput As Long               ' Parameter to enaRegisterOutputCallback
    pfnUndefFunc As Long            ' Parameter to enaUndefFuncCallback
    pDefaultDispatch As Object      ' Parameter to enaSetDefaultDispatch
    pszDeclarations As String       ' Script fragment containing constant and functions declarations
    pfnSetGlobals As Long           ' This function gets called after enaCompile, but before any
'    script code is executed to allow for calls to enaSetGlobal*
    pfnGetGlobals As Long           ' This function gets called before enaFreeContext to allow
'    for retrieval of global variables (enaGetGlobal*)
    pfnAbout As Long                ' This function gets called for About menu processing
    szFileName As String * 260      ' Current open file name in EnableEdit control, Specify initial file
    pszReferences As String         ' Type library references
End Type ' customEdit

Const ENEDIT_DISPLAY_INIT_ERROR = &H1      ' Display a message on a script editor initialization error
Const ENEDIT_CALLING_FROM_VB = &H2         ' The caller of this function is a VB application or control

Declare Function enaScriptEditor Lib "enable40.dll" (pCustomEdit As Any, ByVal szCmdLine As String, ByVal iCmdShow As Long, ByVal flags As Long) As Long

' Script Edit Messages
Public Const SEM_SHOWOPTIONS = &HDD               ' Display the options dialog box
Public Const SEM_ADDTYPELIBREF = &HEC               ' Add a type library reference

Public Const SEM_SETCOLORTABLE = &HFC               ' Set the color table for syntax coloring
Public Const SEM_SETEDITOPTIONS = &HFB              ' Set editing options

Public Const SEM_CRLF = &HED

Public Const SEM_SETCONTEXT = &HFA                  ' Context used to update debug cursor, etc.

Public Const SEM_BREAKPOINT = &HEF                  ' Toggle, set, or clear a breakpoint
Public Const SEM_CLEARBREAKPOINTS = &HEE            ' Remove all breakpoints
Public Const SEM_ENUMBREAKPOINTS = &HF9             ' Enumerate all breakpoints

Public Const SEM_GETTEXTSELECTION = &HF7            ' Return an ITextSelection * to the current selection

Public Const SEM_OPEN = &HFE                        ' Open a file or stream in the edit control
Public Const SEM_SAVE = &HFD                        ' Save the contents of the edit control

Public Type enaEditColorTable
    Text As Long
    KeyWord As Long
    comment As Long
    Function_ As Long
    Identifier As Long
    Number As Long
    String_ As Long
    Constant As Long
End Type





