VERSION 5.00
Begin VB.UserControl ExtWrapper 
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   ControlContainer=   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   3990
End
Attribute VB_Name = "ExtWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
DefLng A-Z

Private vbcext As VBControlExtender


Public Function CaricaEstensione(strNomeEstensione As String) As Object
    
    On Local Error Resume Next
    Licenses.Add strNomeEstensione
    Err.Clear
    On Local Error GoTo 0
    Set vbcext = UserControl.Controls.Add(strNomeEstensione, OGGETTO_ESTENSIONE, Me)
    If Err.Number = 0 Then
        UserControl.Height = vbcext.Height
        UserControl.Width = vbcext.Width
    Else
        MsgBox Err.Description
        'Err.Clear
    End If
    Set CaricaEstensione = vbcext

End Function

Public Function ScaricaEstensione() As Boolean
    
    On Local Error Resume Next
    If Not vbcext Is Nothing Then
        Set vbcext = Nothing
        UserControl.Controls.Remove OGGETTO_ESTENSIONE
    End If
    ScaricaEstensione = Err = 0
End Function

Public Property Get Controls(Optional key As Variant) As Object
    If IsMissing(key) Then
        Set Controls = UserControl.Controls
    Else
        Set Controls = UserControl.Controls(key)
    End If
End Property
