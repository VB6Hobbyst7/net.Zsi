VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.1#0"; "CODEJOCK.COMMANDBARS.V12.1.1.OCX"
Begin VB.Form frmShortcuts 
   BorderStyle     =   0  'None
   Caption         =   "Shortcuts"
   ClientHeight    =   6915
   ClientLeft      =   6435
   ClientTop       =   2070
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImgListMenu 
      Left            =   1320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":0000
            Key             =   "menunew"
            Object.Tag             =   "2101"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":059A
            Key             =   "rename"
            Object.Tag             =   "2102"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":0B34
            Key             =   "DelGrp"
            Object.Tag             =   "2103"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":10CE
            Key             =   "Del"
            Object.Tag             =   "2104"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":1668
            Key             =   "AddProg"
            Object.Tag             =   "2105"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":1F42
            Key             =   "RunProg"
            Object.Tag             =   "2106"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":24DC
            Key             =   "Edit"
            Object.Tag             =   "2107"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListShortCuts 
      Left            =   0
      Top             =   -60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":2A76
            Key             =   "CartellaCH"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":3010
            Key             =   "CartellaAP"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":35AA
            Key             =   "UNREG"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TrVShortcuts 
      Height          =   6015
      Left            =   480
      TabIndex        =   0
      Top             =   420
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10610
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImgListShortCuts"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListShortCutsBig 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":3B44
            Key             =   "CartellaCH"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":423E
            Key             =   "CartellaAP"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":4938
            Key             =   "UNREG"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   1980
      Top             =   0
      _Version        =   786433
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ID_POPUP_ADDGROUP = 2101
Private Const ID_POPUP_RENAME = 2102
Private Const ID_POPUP_DELGROUP = 2103
Private Const ID_POPUP_DELITEM = 2104
Private Const ID_POPUP_ADDPROG = 2105
Private Const ID_POPUP_RUNPROG = 2106
Private Const ID_POPUP_EDITPROG = 2107

Private Const MAX_PATH = 260
Private Const SHGFI_SYSICONINDEX = &H4000  ' System icon index
Private Const SHGFI_LARGEICON = &H0        ' Large icon
Private Const SHGFI_SMALLICON = &H1        ' Small icon
Private Const ILD_TRANSPARENT = &H1        ' Display transparent
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Const VK_F2 = &H71

Private Const ICON_UNREG = "UNREG"

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, ipic As IPicture) As Long

Private MColShortcuts As New Collection

Dim MbolIconeGrandi As Boolean
Dim MLngKeyIndex As Long
Public Property Get IconeGrandi() As Boolean
    IconeGrandi = MbolIconeGrandi
End Property
Public Sub CaricaShortcuts()
End Sub

Private Sub AddGroupItemsTw(objParent As MSXML2.IXMLDOMNode, objTwGroupParent As MSComctlLib.Node)

End Sub
'Restituisce un riferimento (key) all'icona del file selezionato e lo aggiunge all'imagelist dell'albero se non presente
Private Function GetFileIcon(ByVal strNomeFile As String) As String
    

End Function

Private Function IconToPicture(ByVal hIcon As Long) As Picture
End Function
Public Sub SalvaShortcuts()
End Sub

Private Sub AddGroupItems(oDocument As MSXML2.DOMDocument, objGroupNode As MSXML2.IXMLDOMNode, objParent As MSComctlLib.Node)
End Sub


Private Sub Shortcuts_EditProgram()

End Sub

Private Sub Shortcuts_AddGroup()
End Sub


Private Sub Shortcuts_AddProgram()
End Sub

Private Sub Shortcuts_DelGroup()
End Sub

Private Sub Shortcuts_DelItem()
End Sub

Private Sub Shortcuts_Rename()

End Sub


Private Sub Form_Load()
    Me.Hide
End Sub

Public Sub ImpostaImageList(Optional ByVal TipoIcone As setTipoIconeAlbero = enmLeggiDaProfilo)
    
End Sub

