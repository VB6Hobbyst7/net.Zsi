Attribute VB_Name = "MMDIChild"
Option Explicit



Public Function SetMDIChild(ByRef f As Form, ByVal bChild As Boolean, Optional frmParent As Form) As Boolean

Dim oStyle As Long, Ret As Long, lNewStyle As Long

   'oStyle = GetWindowLong(f.hwnd, GWL_STYLE)

   '/ cambio lo stile
   If bChild Then
      'oStyle = oStyle Or WS_CHILD Or (Not WS_POPUP)
      oStyle = WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or _
                WS_CLIPSIBLINGS Or WS_DLGFRAME Or WS_GROUP Or _
                WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or _
                WS_SYSMENU Or WS_TABSTOP Or _
                WS_THICKFRAME Or WS_SIZEBOX Or WS_VISIBLE Or _
                WS_CHILD Or (Not WS_POPUP)
      
   Else
      'oStyle = oStyle Or (Not WS_CHILD) Or WS_POPUP
      oStyle = WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or _
                WS_CLIPSIBLINGS Or WS_DLGFRAME Or WS_GROUP Or _
                WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or _
                 WS_SYSMENU Or WS_TABSTOP Or _
                WS_THICKFRAME Or WS_SIZEBOX Or WS_VISIBLE Or _
                WS_POPUP
   End If
   
   lNewStyle = oStyle
   Ret = SetWindowLong(f.hwnd, GWL_STYLE, lNewStyle)
   
   
   If Not frmParent Is Nothing Then
      Ret = SetParent(f.hwnd, frmParent.hwnd)
      SendMessage f.hwnd, WM_NCPAINT, 1, 0&
   Else
      Ret = SetParent(f.hwnd, 0&)
   End If
End Function

Public Sub MakeChild2(Child As Form, Optional Maximazebox As Boolean = True, Optional MDI As VB.Form = Nothing, Optional bolShowChild As Boolean, Optional ByVal sngResizeOffsetX As Single = 0, Optional ByVal sngResizeOffsetY As Single = 0)
    Dim max As Boolean
    Dim mdichild As Form
    If Not MDI Is Nothing Then
        Set mdichild = MDI
    Else
        Set mdichild = MXNU.FrmMetodo.GetMdiChild
    End If
    
    ShowHideTitleBar Child, False
    
    SetParent Child.hwnd, mdichild.hwnd
    
    'RIF.A#11747 - passo gli offset di ridimensionamento
    If (sngResizeOffsetX <> 0) Then
        mdichild.mSngResizeOffsetX = sngResizeOffsetX
    End If
    If (sngResizeOffsetY <> 0) Then
        mdichild.mSngResizeOffsetY = sngResizeOffsetY
    End If
    
    'copio la caption e l'icona
    mdichild.Caption = Child.Tag
    Set mdichild.Icon = Child.Icon
    
    Call mdichild.SetChild(Child)
    
    'resize la form padre come la figlia solo se non sono massimizzata
    If mdichild.WindowState <> 2 Then
        mdichild.Move Child.Left, Child.Top, Child.Width, Child.Height
    Else
        max = True
    End If
    
    'rif.GIGI 20/05/2014 - problema Tecnica3 visualizzazione su Terminal Server
    If mdichild.WindowState <> 2 Then
        CentraFinestra mdichild.hwnd
    End If
    mdichild.Show
    
    If max Then
       mdichild.WindowState = 2
    End If
    
    If Child.WindowState <> 2 Then Child.Move 0, 0
    
    If (bolShowChild And Not Child.Visible) Then
        Child.Show
    End If
    
    If Not Maximazebox Then
        NoMAXIMIZE mdichild.hwnd
    End If
    
End Sub

Private Function ShowHideTitleBar(myform As Form, ByVal bState As Boolean)
Dim lStyle As Long
Dim tR As RECT

   ' Get the window's position:
   GetWindowRect myform.hwnd, tR

   ' Modify whether title bar will be visible:
   lStyle = GetWindowLong(myform.hwnd, GWL_STYLE)
   If (bState) Then
      myform.Caption = myform.Tag
      If myform.ControlBox Then
         lStyle = lStyle Or WS_SYSMENU
      End If
      If myform.MaxButton Then
         lStyle = lStyle Or WS_MAXIMIZEBOX
      End If
      If myform.MinButton Then
         lStyle = lStyle Or WS_MINIMIZEBOX
      End If
      If myform.Caption <> "" Then
         lStyle = lStyle Or WS_CAPTION
      End If
   Else
      myform.Tag = myform.Caption
      myform.Caption = ""
      lStyle = lStyle And Not WS_SYSMENU
      lStyle = lStyle And Not WS_MAXIMIZEBOX
      lStyle = lStyle And Not WS_MINIMIZEBOX
      lStyle = lStyle And Not WS_CAPTION
   End If
   SetWindowLong myform.hwnd, GWL_STYLE, lStyle

   ' Ensure the style takes and make the window the
   ' same size, regardless that the title bar etc
   ' is now a different size:
   SetWindowPos myform.hwnd, _
       0, tR.Left, tR.Top, _
       tR.Right - tR.Left, tR.Bottom - tR.Top, _
       SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED

   RemoveBorder myform.hwnd

   myform.Refresh

End Function

Public Function GetMDIClient(hFrame As Long) As Long
  Dim hwnd As Long
  hwnd = GetWindow(hFrame, GW_CHILD)
  While hwnd <> 0
    If Right(LCase(WindowClass(hwnd)), 9) = "mdiclient" Then
      GetMDIClient = hwnd
      Exit Function
    End If
    hwnd = GetWindow(hwnd, GW_HWNDNEXT)
  Wend
  GetMDIClient = 0
End Function

Public Function WindowClass(ByVal hwnd As Long) As String
    Dim TextLen As Long
    Dim Ret As String
    Dim i As Long

    Ret = String$(255, 0)
    TextLen = GetClassName(hwnd, Ret, 256)
    WindowClass = Left$(Ret, TextLen)

End Function


Private Sub RemoveBorder(hwnd As Long)
Const WS_THICKFRAME = &H40000
Dim l As Long
    
    l = GetWindowLong(hwnd, GWL_STYLE)
    l = l And Not (WS_THICKFRAME)
    l = SetWindowLong(hwnd, GWL_STYLE, l)
End Sub


Public Sub NoMAXIMIZE(hwnd As Long)
Dim lngStyle As Long
    
    lngStyle = GetWindowLong(hwnd, GWL_STYLE)
    lngStyle = lngStyle And Not (WS_MAXIMIZEBOX)
    SetWindowLong hwnd, GWL_STYLE, lngStyle
    
End Sub




