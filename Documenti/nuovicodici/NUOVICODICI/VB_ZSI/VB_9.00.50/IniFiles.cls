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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Public Function SetValue(IniFileName As String, SectionName As String, KeyInSection As String, DefaultValue As String) As Variant
   Dim lngRis As Long
   Dim StrRit As String * 256
   '
   On Error GoTo errhand
   lngRis = WritePrivateProfileString(SectionName, KeyInSection, DefaultValue, IniFileName)
   If lngRis = 0 Then
      MsgBox "Problemi di scrittura valori in file INI", vbCritical
   End If
   Exit Function
   
errhand:
   MsgBox Err.Number & " " & Err.Description & "," & Err.Source
End Function

Public Function GetValue(IniFileName As String, SectionName As String, KeyInSection As String, DefaultValue As Variant) As Variant
   Dim lngRis As Long
   Dim StrRit As String * 256
   '
   On Error GoTo errhand
   lngRis = GetPrivateProfileString(SectionName, KeyInSection, DefaultValue, StrRit, Len(StrRit) - 1, IniFileName)
   GetValue = ConvertCToVBString(StrRit)
   Exit Function
   
errhand:
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


