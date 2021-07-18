VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParametri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public IDPar As String
Public NomePar  As String
Public Valore1  As String
Public Valore2  As String
Public Valore3  As String
Public Valore4  As String
Public Valore5  As String

Public Tipo1  As String
Public Tipo2  As String
Public Tipo3  As String
Public Tipo4  As String
Public Tipo5  As String

Public Function Parametro2Variant() As Variant
    Parametro2Variant = Array(Valore1, Valore2, Valore3, Valore4, Valore5)
End Function

Public Property Let Tipo(ByVal intInd As Integer, ByVal vData As String)
    Select Case intInd
        Case 1: Tipo1 = vData
        Case 2: Tipo2 = vData
        Case 3: Tipo3 = vData
        Case 4: Tipo4 = vData
        Case 5: Tipo5 = vData
    End Select
End Property

Public Property Get Tipo(ByVal intInd As Integer) As String
    Select Case intInd
        Case 1: Tipo = Tipo1
        Case 2: Tipo = Tipo2
        Case 3: Tipo = Tipo3
        Case 4: Tipo = Tipo4
        Case 5: Tipo = Tipo5
    End Select
End Property

Public Property Let Valore(ByVal intInd As Integer, ByVal vData As String)
    Select Case intInd
        Case 1: Valore1 = vData
        Case 2: Valore2 = vData
        Case 3: Valore3 = vData
        Case 4: Valore4 = vData
        Case 5: Valore5 = vData
    End Select
End Property

Public Property Get Valore(ByVal intInd As Integer) As String
    Select Case intInd
        Case 1: Valore = Valore1
        Case 2: Valore = Valore2
        Case 3: Valore = Valore3
        Case 4: Valore = Valore4
        Case 5: Valore = Valore5
    End Select
End Property


