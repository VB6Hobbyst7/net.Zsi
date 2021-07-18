VERSION 5.00
Begin VB.Form EmptyForm 
   BorderStyle     =   0  'None
   ClientHeight    =   1545
   ClientLeft      =   11010
   ClientTop       =   4350
   ClientWidth     =   1725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "EmptyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Activate()
    Call MXNU.Attiva_Toolbar(hwnd, 0)
    Call MXNU.ImpostaFormAttiva(Me)
    On Local Error Resume Next
    metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_FRMORIGINALSIZE).Enabled = False
    On Local Error GoTo 0
End Sub

Private Sub Form_Deactivate()
    On Local Error Resume Next
    metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_FRMORIGINALSIZE).Enabled = True
    On Local Error GoTo 0
End Sub


Private Sub Form_Load()
    Me.Width = 1
    Me.Height = 1
    Me.Show
End Sub


