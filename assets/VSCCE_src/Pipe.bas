Attribute VB_Name = "Pipe"
Option Explicit

' Constants that will be used in the API functions
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&

' Declare the needed API functions
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal bsName As String, ByVal buff As String, ByVal ch As Long) As Long


Public Sub stdout(OutStr As String)
    Dim lResult As Long
    If Not VBHelp.DebugMode Then
        WriteFile GetStdHandle(STD_OUTPUT_HANDLE), OutStr, Len(OutStr), lResult, ByVal 0&
    Else
        Debug.Print OutStr
    End If
End Sub

Public Sub stderr(ErrStr As String)
    Dim lResult As Long
    If Not VBHelp.DebugMode Then
        WriteFile GetStdHandle(STD_ERROR_HANDLE), ErrStr, Len(ErrStr), lResult, ByVal 0&
    Else
        Debug.Print ErrStr
    End If
End Sub

Public Function stdin() As String
   Dim postData As String
   Dim llStdIn As Long
   Dim lsBuff As String
   Dim llBytesRead As Long
   
   ' Read the standard input handle
   llStdIn = GetStdHandle(STD_INPUT_HANDLE)
   ' Get POSTed data from STDIN
   Do
      lsBuff = String(1024, 0)    ' Create a buffer big enough to hold the 1024 bytes
      llBytesRead = 1024          ' Tell it we want at least 1024 bytes
      If ReadFile(llStdIn, ByVal lsBuff, 1024, llBytesRead, ByVal 0&) Then
         ' Read the data
         ' Add the data to our string
         postData = postData & Left(lsBuff, llBytesRead)
         If llBytesRead < 1024 Then Exit Do
      Else
         Exit Do
      End If
   Loop
stdin = postData

End Function
