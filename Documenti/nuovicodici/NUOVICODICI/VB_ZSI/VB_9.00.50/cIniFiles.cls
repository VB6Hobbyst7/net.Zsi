VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIniFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function SetValue(IniFileName As String, SectionName As String, KeyInSection As String, DefaultValue As String) As Variant
   Dim lngRis As Long
   Dim StrRit As String * 256
   '
   On Error GoTo ErrHand
   lngRis = WritePrivateProfileString(SectionName, KeyInSection, DefaultValue, IniFileName)
   If lngRis = 0 Then
      MsgBox "Problemi di scrittura valori in file INI", vbCritical
   End If
   Exit Function
   
ErrHand:
   MsgBox Err.Number & " " & Err.Description & "," & Err.Source
End Function

Public Function GetValue(IniFileName As String, SectionName As String, KeyInSection As String, DefaultValue As Variant) As Variant
   Dim lngRis As Long
   Dim StrRit As String * 256
   '
   On Error GoTo ErrHand
   lngRis = GetPrivateProfileString(SectionName, KeyInSection, DefaultValue, StrRit, Len(StrRit) - 1, IniFileName)
   GetValue = ConvertCToVBString(StrRit)
   Exit Function
   
ErrHand:
   MsgBox Err.Number & " " & Err.Description & "," & Err.Source
End Function

Public Function ConvertCToVBString(InString As String) As String
   ' trim string to first chr(0)
   If InStr(1, InString, Chr(0)) > 0 Then
      ConvertCToVBString = Left(InString, InStr(1, InString, Chr(0), vbBinaryCompare) - 1)
   Else
      ConvertCToVBString = InString
   End If
End Function



