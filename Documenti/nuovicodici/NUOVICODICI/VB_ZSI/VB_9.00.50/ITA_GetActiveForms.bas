Attribute VB_Name = "ITA_GetActiveForms"
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal LPARAM As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    
Global bolFoundForm As Boolean
Dim hwndself As Long
Dim captionself As String

Global ArrayForm(4) As String


Public Enum EnProcedure
    
    EnRETTIFICASTORICOMAGR = 1 '"RETTIFICASTORICOMAGR"
    EnVISIONESTORICOMAGR = 2 '"VISIONESTORICOMAGR"
    EnMODIFICASTORICOMAGR = 3 '"MODIFICASTORICOMAGR"
    EnRETTIFICAINVENTARIO = 4 '"RETTIFICAINVENTARIO"
    

End Enum

' Return the class name of the specified window
' Example: MsgBox GetWindowClass(Me.hWnd)
Function GetWindowClass(ByVal hwnd As Long) As String
    Dim sClass As String

    sClass = Space$(256)
    GetClassName hwnd, sClass, 255
    GetWindowClass = Left$(sClass, InStr(sClass, vbNullChar) - 1)
End Function

Function EnumChildProc(ByVal hwnd As Long, ByVal LPARAM As Long) As Long
    ' ricerca il controllo che si sta aprendo
    Dim ss As String
    Dim strCaption As String
    
    ss = Space(2000)
    
    EnumChildProc = 1
   
    If hwnd <> hwndself Then
        'If UCase(Left(GetWindowClass(hwnd), 11)) = UCase("ThunderForm") Then
            Call GetWindowText(hwnd, ss, Len(ss))
    
            strCaption = Left(ss, InStr(ss, vbNullChar) - 1)
            If (strCaption > "") Then
                If strCaption = captionself Then
                    ' il controllo che si sta aprendo è già aperto
                    EnumChildProc = 0
                    bolFoundForm = True
                
                Else
                
                End If
            End If
        'End If
    End If
    
    
End Function
 
Function HasFoundForm(ByVal f As Form) As Boolean
    hwndself = f.hwnd
    captionself = f.Caption
    
'    bolFoundForm = False
'    HasFoundForm = False
'
'    Call EnumChildWindows(MXNU.FrmMetodo.hwnd, AddressOf EnumChildProc, 0)
    
    HasFoundForm = True

End Function

Public Function FormAccess(ByVal UserID As String, ByVal proc As Long)
    Dim strSQL As String
    Dim rSQL As CRecordSet
    Dim ArrayForm(4) As String
    
    ArrayForm(1) = "RETTIFICASTORICOMAGR"
    ArrayForm(2) = "VISIONESTORICOMAGR"
    ArrayForm(3) = "MODIFICASTORICOMAGR"
    ArrayForm(4) = "RETTIFICAINVENTARIO"
    
    strSQL = "SELECT * FROM ITA_PERMESSIUTENTI WHERE :CAMPO = 1 AND USERID = :USERID"
    strSQL = Replace(strSQL, ":CAMPO", ArrayForm(proc))
    strSQL = Replace(strSQL, ":USERID", hndDBArchivi.FormatoSQL(UserID, DB_TEXT))
    With MXDB
        Set rSQL = .dbCreaSS(hndDBArchivi, strSQL, TIPO_SNAPSHOT)
        
        FormAccess = Not (.dbFineTab(rSQL))
        
        Call .dbChiudiSS(rSQL)
        
        Set rSQL = Nothing
    
    
    End With
    
End Function

Public Sub SalvaLog(ByVal User As String, ByVal IdSessione As Long, ByVal strMSG As String, Optional ByVal strPROC As String = "")
        Dim oDLG As Object
10      Dim cNFile As String
20      Dim cTxt As String

' ObjInserter
'---------------------------------------------
30      On Error GoTo SalvaLog_ErrHandler
'---------------------------------------------

        Set oDLG = CreateObject("MSComDlg.CommonDialog")
        
        If (Not (oDLG Is Nothing)) And (strMSG <> "") Then
    
40          With oDLG
            
50              .CancelError = True
60              .Filter = "File di Testo(*.txt)|*.txt"
                If strPROC <> "" Then
                    .FileName = strPROC & " - " & CStr(IdSessione) & " .txt"
                Else
70                  .FileName = "LogFile - " & CStr(IdSessione) & " .txt"
                End If
80              .ShowSave
90              cNFile = .FileName
            
100          End With
        
110          cTxt = strMSG
        
120          Open cNFile For Output As #1
    
130          Print #1, cTxt
        
140          Close #1
        
        End If
        
        Set oDLG = Nothing
        
        
' ObjInserter
'---------------------------------------------
    Exit Sub

150 SalvaLog_ErrHandler:

160      Dim cLogString As String
         cLogString = "Error in ImpContabilita: Sub SalvaLog (" & Erl & ")" & vbCrLf & Err.Description & vbCr & vbCr & "(Error #" & Err.Number & ")"
170
    If Err.Number = 5 Or Err.Number = 32755 Then
        
        Exit Sub
    Else
        Call MXNU.MsgBoxEX(cLogString, vbCritical, "Salvatawggio file di Log")
    End If
'---------------------------------------------

End Sub

