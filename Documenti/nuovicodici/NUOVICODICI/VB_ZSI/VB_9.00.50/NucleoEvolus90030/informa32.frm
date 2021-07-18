VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInformazioni 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Project1"
   ClientHeight    =   6600
   ClientLeft      =   6870
   ClientTop       =   2985
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "informa32.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Tag             =   "About Project1"
   Begin VB.PictureBox PictCandidate 
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   1095
      Left            =   5160
      Picture         =   "informa32.frx":0CCA
      ScaleHeight     =   1095
      ScaleWidth      =   1095
      TabIndex        =   17
      Top             =   1620
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Dll Version"
      Height          =   345
      Left            =   4920
      TabIndex        =   15
      Tag             =   "&System Info..."
      Top             =   4680
      Width           =   1245
   End
   Begin VB.CommandButton cmdShowactiveModules 
      Caption         =   "Moduli Attivi"
      Height          =   345
      Left            =   4920
      TabIndex        =   12
      Tag             =   "&System Info..."
      Top             =   5940
      Width           =   1245
   End
   Begin MSComctlLib.ImageList imgListInformazioni 
      Left            =   4350
      Top             =   2475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   114
      ImageHeight     =   114
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "informa32.frx":18E5
            Key             =   "info98"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "informa32.frx":9A37
            Key             =   "infoxp"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "informa32.frx":133B9
            Key             =   "infoEvolus"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   1500
      Left            =   240
      ScaleHeight     =   1440
      ScaleMode       =   0  'User
      ScaleWidth      =   1710
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   300
      Width           =   1770
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4905
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   5100
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4920
      TabIndex        =   1
      Tag             =   "&System Info..."
      Top             =   5520
      Width           =   1245
   End
   Begin VB.Label lblDL 
      Caption         =   "In esecuzione D.L. 518 del 29/12/92."
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   5760
      Width           =   4575
   End
   Begin VB.Image ImgWin7Logo 
      Height          =   1080
      Left            =   5220
      Picture         =   "informa32.frx":14413
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label lblNomeComputer 
      Caption         =   "Nome Computer: "
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Tag             =   "Version"
      Top             =   1980
      Width           =   3885
   End
   Begin VB.Label lblPatch 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4140
      TabIndex        =   14
      Tag             =   "Version"
      Top             =   900
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label lblPatch 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   13
      Tag             =   "Version"
      Top             =   900
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label lblSessione 
      Caption         =   "Sessione: "
      Height          =   195
      Left            =   2160
      TabIndex        =   10
      Tag             =   "Version"
      Top             =   1740
      Width           =   3885
   End
   Begin VB.Label lblTerminale 
      Caption         =   "Terminale: "
      Height          =   195
      Left            =   2160
      TabIndex        =   9
      Tag             =   "Version"
      Top             =   1500
      Width           =   3885
   End
   Begin VB.Label lblUtente 
      Caption         =   "Utente: "
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Tag             =   "Version"
      Top             =   1260
      Width           =   3885
   End
   Begin VB.Label lblLicenseTo 
      Caption         =   "Questo prodotto è concesso in licenza a:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Tag             =   "LicenseTo"
      Top             =   3750
      Width           =   5835
   End
   Begin VB.Label lblLicenseTo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LicenseTo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Tag             =   "LicenseTo"
      Top             =   4080
      Width           =   5895
   End
   Begin VB.Label lblDescription 
      Caption         =   $"informa32.frx":14FA6
      Height          =   1290
      Left            =   240
      TabIndex        =   5
      Tag             =   "App Description"
      Top             =   2340
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   2160
      TabIndex        =   4
      Tag             =   "Application Title"
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"informa32.frx":150EE
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   240
      TabIndex        =   2
      Tag             =   "Warning: ..."
      Top             =   4740
      Width           =   4545
   End
End
Attribute VB_Name = "frmInformazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Reg Key Security Options...
Const KEY_ALL_ACCESS = &H2003F
                                          

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Command1_Click()
'carico la form con un metodo inizializza in quanto e' una form modale dentro una form modale
'se non faccio cosi' la form andrebbe in errore alla chiusura della stessa
    frmControllaVersioneDllEOcx.Inizializza
End Sub

Private Sub Form_Load()
    Dim strSP As String

    Me.Caption = "Informazioni su " & App.Title
    lblTitle.Caption = App.Title

    lblVersion.Caption = "Versione " & MXNU.VersioneMetodo
    'modifica del 19/05/2003 - caricamento versione patch client/server
    Call LoadPatchVersion

'    strSP = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\SP.ini", "Version", "Version", "")
'    If strSP <> "" Then
'        lblVersion.Caption = lblVersion.Caption & " " & strSP
'    End If

    lblUtente.Caption = lblUtente.Caption & MXNU.UtenteAttivo
    lblTerminale.Caption = lblTerminale.Caption & MXNU.NTerminale
    lblSessione.Caption = lblSessione.Caption & MXNU.IDSessione
    lblNomeComputer.Caption = lblNomeComputer.Caption & MXNU.NomeComputer
    Dim strLic As String
    strLic = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "Registrazione", "Nome Cliente", "") & Chr$(10)
    strLic = strLic & MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "Registrazione", "Nome Ditta", "")
    lblLicenseTo(1).Caption = strLic
    CentraFinestra Me.hwnd
    
    'Modifica per MetodoXP
    Dim strImageInfo As String
    #If ISMETODOXP = 1 Then
        If MXNU.MetodoXP Then
            strImageInfo = "infoxp"
        Else
            strImageInfo = "info98"
        End If
    #Else
        strImageInfo = "info98"
    #End If
    #If ISMETODO2005 = 1 Then
        strImageInfo = "infoEvolus"
        Call CambiaColoriControlli(Me)
        lblTitle.BackStyle = vbTransparent
        lblVersion.BackStyle = vbTransparent
        lblUtente.BackStyle = vbTransparent
        lblTerminale.BackStyle = vbTransparent
        lblSessione.BackStyle = vbTransparent
        lblNomeComputer.BackStyle = vbTransparent
        lblDescription.BackStyle = vbTransparent
        lblLicenseTo(0).BackStyle = vbTransparent
        lblDisclaimer.BackStyle = vbTransparent
        If TemaConGradiente Then
            lblLicenseTo(1).BackColor = SysGradientColor1
            lblPatch(0).BackColor = SysGradientColor1
            lblPatch(1).BackColor = SysGradientColor1
        End If
        If PictCandidate.Visible Then
            PictCandidate.BackColor = SysGradientColor1
        End If
    #End If
    Set picIcon.Picture = imgListInformazioni.ListImages(strImageInfo).Picture
End Sub

Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub


Private Sub cmdOK_Click()
        Unload Me
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' Try To Get System Info Program Path\Name From Registry...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Validate Existance Of Known 32 Bit File Version
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Error - File Can Not Be Found...
                Else
                        GoTo SysInfoErr
                End If
        ' Error - Registry Entry Can Not Be Found...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
        

        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
        KeyValSize = 1024                                       ' Mark Variable Size
        

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        

        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
                tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
        Else                                                    ' WinNT Does NOT Null Terminate String...
                tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
        End If
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
                For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
        

        GetKeyValue = True                                      ' Return Success
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
        Exit Function                                           ' Exit
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmInformazioni = Nothing
End Sub

Private Sub cmdShowactiveModules_Click()
    FrmShowActiveModules.Show vbModal
End Sub

Private Sub LoadPatchVersion()
Const PATCH_SERVER = 0
Const PATCH_CLIENT = 1
Dim oXMLDoc As MSXML2.DOMDocument
Dim strVersion As String
Dim strBuild As String

    Exit Sub
    
    On Local Error GoTo ERR_LoadPatchVersion
    strAppliedPatch = ""
    'caricamento del file XML
    Set oXMLDoc = New MSXML2.DOMDocument
    If (oXMLDoc.Load(MXNU.PercorsoPreferenze & "\SP.xml")) Then
        'patch server
        '   carico le informazioni dell'ultima versione installata
        If (LoadPatchInformations(oXMLDoc, "server", strVersion, strBuild)) Then
            If (strVersion = MXNU.VersioneMetodo) Then
                lblPatch(PATCH_SERVER).Caption = "Patch Server: " & strBuild
            End If
        End If
        '   patch client
        If (LoadPatchInformations(oXMLDoc, MXNU.NomeComputer, strVersion, strBuild)) Then
            If (strVersion = MXNU.VersioneMetodo) Then
                lblPatch(PATCH_CLIENT).Caption = "Patch Client: " & strBuild
            End If
        End If
    End If

END_LoadPatchVersion:
    If (Len(lblPatch(PATCH_SERVER).Caption) = 0) Then lblPatch(PATCH_SERVER).Caption = "Patch Server: Nessuna"
    If (Len(lblPatch(PATCH_CLIENT).Caption) = 0) Then lblPatch(PATCH_CLIENT).Caption = "Patch Client: Nessuna"
    Exit Sub
    
ERR_LoadPatchVersion:
    lblPatch(PATCH_SERVER).Caption = ""
    lblPatch(PATCH_CLIENT).Caption = ""
    Resume END_LoadPatchVersion
End Sub

Private Function LoadPatchInformations(oXMLDoc As MSXML2.DOMDocument, ByVal strNodeName As String, ByRef strVersion As String, ByRef strBuild As String) As Boolean
Dim bolRes As Boolean
Dim oInfoNodeList As MSXML2.IXMLDOMNodeList
Dim oInfoNode As MSXML2.IXMLDOMNode
Dim xPathQuery As String

    bolRes = True
    On Local Error GoTo ERR_LoadPatchInformations
    xPathQuery = "patchinfo/" & strNodeName & "/release"
    Set oInfoNodeList = oXMLDoc.selectNodes(xPathQuery)
    bolRes = (Not oInfoNodeList Is Nothing)
    If (bolRes) Then
        Set oInfoNode = oInfoNodeList.Item(oInfoNodeList.length - 1)
        strVersion = oInfoNode.Attributes.getNamedItem("version").nodeValue
        strBuild = oInfoNode.Attributes.getNamedItem("build").nodeValue
    End If
    
END_LoadPatchInformations:
    LoadPatchInformations = bolRes
    Exit Function
    
ERR_LoadPatchInformations:
    bolRes = False
    Resume END_LoadPatchInformations
End Function



