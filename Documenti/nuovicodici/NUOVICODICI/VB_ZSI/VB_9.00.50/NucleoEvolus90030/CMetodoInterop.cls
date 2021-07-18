VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMetodoInterop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements MXInterfacce.IMetodoInterop

Public Sub Dispose()
    Call IMetodoInterop_Dispose
End Sub

Public Function Init(ByVal sPrefPath As String, ByVal sFirmId As String, Optional ByVal sLanguage As String = "IT", Optional ByVal sUserId As String, Optional ByVal sPassword As String, Optional ByVal TypeLog As Integer = 0) As String
    Init = IMetodoInterop_Init(sPrefPath, sFirmId, sLanguage, sUserId, sPassword, TypeLog)
End Function

Public Function GetMetodoObject(ByVal sObjectName As String) As Object
   Set GetMetodoObject = IMetodoInterop_GetMetodoObject(sObjectName)
End Function

Private Sub IMetodoInterop_Dispose()
    
End Sub

Private Function IMetodoInterop_GetMetodoObject(ByVal sObjectName As String) As Object
    Dim objResult As Object

   On Error GoTo GetMetodoObject_Error

    Select Case UCase$(sObjectName)
        Case "MXNU"
            Set objResult = MXNU
        Case "MXDB"
            Set objResult = MXDB
        Case "MXCREP"
            Set objResult = MXCREP
        Case "MXFT"
            Set objResult = MXFT
        Case "MXVI"
            Set objResult = MXVI
        Case "MXAA"
            Set objResult = MXAA
        Case "MXVA"
            Set objResult = MXVA
        Case "MXCT"
            Set objResult = MXCT
        Case "MXSC"
            Set objResult = MXSC
        Case "MXART"
            Set objResult = MXART
        Case "MXGD"
            Set objResult = MXGD
        Case "MXVALID"
            Set objResult = New MXKit.ControlliCampo
        Case "MXUTIL"
            Set objResult = New MXUTil.XUtil
        Case "MXSPREAD"
            Set objResult = New MXSpread.CSpread
        Case "MXALL"
            Set objResult = MXALL
        Case "MXVI"
            Set objResult = MXVI
        Case "HNDDBARCHIVI"
            Set objResult = hndDBArchivi
        Case "COLLECTION"
            Set objResult = New NetCollection
        Case "MXBROWSER"
            Set objResult = mMetodoBrowser
        Case "NETFX"
            Set objResult = NETFX
    End Select
    
    If (Not objResult Is Nothing) Then
        Set IMetodoInterop_GetMetodoObject = objResult
    End If

   On Error GoTo 0
   Exit Function

GetMetodoObject_Error:
    Call MsgBox("Errore nella funzione GetMetodoObject del modulo cMetodoInterop" & UCase$(sObjectName) & "[" & Err.Number & "]: " & Err.Description, vbCritical)
End Function

Private Function IMetodoInterop_Init(ByVal sPrefPath As String, ByVal sFirmId As String, Optional ByVal sLanguage As String = "IT", Optional ByVal sUserId As String, Optional ByVal sPassword As String, Optional ByVal TypeLog As Integer = 0) As String
    Err.Raise 425, "CMetodoInterop.Init", "Init not Supported"
End Function
