VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMach4Emu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_bIsInit As Boolean

Public Sub LoadGCodeFile(ByVal Filename As String)
End Sub
Public Sub CloseFile()
End Sub
Public Function GetScriptDispatch() As Object
    If m_bIsInit Then
        Set GetScriptDispatch = New IMyScriptObjectEmu
    Else
        MsgBox "Faild to get script dispatch!" & vbCrLf & "IMach4Emu is not initialized!"
    End If
End Function
Public Sub ShutDown()
End Sub
Public Sub GetAxisPosition(ByRef Xpos As Double, ByRef Ypos As Double, ByRef Zpos As Double, ByRef Apos As Double, ByRef Bpos As Double, ByRef Cpos As Double)
End Sub
Public Sub CycleStart()
End Sub
Public Sub FeedHold()
End Sub
Public Sub RewindGcode()
End Sub
Public Sub GetLineNumber()
End Sub
Public Sub SetFRO(ByVal FeedOverride As Double)
End Sub
Public Function GetRev() As String
End Function

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

Private Sub Class_Initialize()
    If Not m_bIsInit Then
        
        If Mach3.g_bLoadProfile Then
            Dim xmlDoc As DOMDocument60
            
            If OpenProfile(xmlDoc, Mach3.g_MainFolder & Mach3.g_XMLProfile) Then
                
                
                
                
                Set xmlDoc = Nothing
            End If
        End If
        m_bIsInit = True
    End If
End Sub

