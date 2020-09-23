Attribute VB_Name = "modGraphics"
'Create font
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

'used with fnWeight
Const FW_NORMAL = 400
'used with fdwCharSet
Const DEFAULT_CHARSET = 1
'used with fdwOutputPrecision
Const OUT_DEFAULT_PRECIS = 0
'used with fdwClipPrecision
Const CLIP_DEFAULT_PRECIS = 0
'used with fdwQuality
Const PROOF_QUALITY = 2
'used with fdwPitchAndFamily
Const DEFAULT_PITCH = 0

Const LOGPIXELSY = 90
'create font

Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As SHGFI_FLAGS) As Long
    
    Public Enum SHGFI_FLAGS
        SHGFI_LARGEICON = &H0           'sfi.hIcon is large icon
        SHGFI_SMALLICON = &H1           'sfi.hIcon is small icon
        SHGFI_OPENICON = &H2            'sfi.hIcon is open icon
        SHGFI_SHELLICONSIZE = &H4       'sfi.hIcon is shell size (not system size), rtns BOOL
        SHGFI_PIDL = &H8                'pszPath is pidl, rtns BOOL
        SHGFI_USEFILEATTRIBUTES = &H10  'parent pszPath exists, rtns BOOL
        SHGFI_ICON = &H100              'fills sfi.hIcon, rtns BOOL, use DestroyIcon
        SHGFI_DISPLAYNAME = &H200       'isf.szDisplayName is filled, rtns BOOL
        SHGFI_TYPENAME = &H400          'isf.szTypeName is filled, rtns BOOL
        SHGFI_ATTRIBUTES = &H800        'rtns IShellFolder::GetAttributesOf  SFGAO_* flags
        SHGFI_ICONLOCATION = &H1000     'fills sfi.szDisplayName with filename containing the icon, rtns BOOL
        SHGFI_EXETYPE = &H2000          'rtns two ASCII chars of exe type
        SHGFI_SYSICONINDEX = &H4000     'sfi.iIcon is sys il icon index, rtns hImagelist
        SHGFI_LINKOVERLAY = &H8000      'add shortcut overlay to sfi.hIcon
        SHGFI_SELECTED = &H10000        'sfi.hIcon is selected icon
        SHGFI_ATTR_SPECIFIED = &H20000  'get only attributes specified in sfi.dwAttributes
    End Enum
    
    Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
       Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
       Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
    
    Public Type SHFILEINFO
      hIcon As Long
      iIcon As Long
      dwAttributes As Long
      szDisplayName As String * MAX_PATH
      szTypeName As String * 80
    End Type

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As SM_Constants) As Long
    Public Enum SM_Constants
      SM_CXICON = 11
      SM_CYICON = 12
      SM_CXFULLSCREEN = 16
      SM_CYFULLSCREEN = 17
      SM_CXSMICON = 49
      SM_CYSMICON = 50
    End Enum

Declare Function ImageList_Draw Lib "comctl32.dll" _
   (ByVal hIml&, ByVal i&, ByVal hDCDest&, _
    ByVal X&, ByVal Y&, ByVal flags&) As Long

    Public Const ILD_BLEND25 = &H2
    Public Const ILD_BLEND50 = &H4
    Public Const ILD_NORMAL = &H0
    Public Const ILD_TRANSPARENT = &H1

'extract icon
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal _
  hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
    ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, _
    ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, _
    ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
      Public Const DI_NORMAL = &H3

'make translucent windows
Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct _
    As Long) As Boolean
    
    Type BLENDFUNCTION
      BlendOp As Byte
      BlendFlags As Byte
      SourceConstantAlpha As Byte
      AlphaFormat As Byte
    End Type

'GENERAL GRAPHICAL API
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As BK_Mode) As Long
    ' Background Modes
    Public Enum BK_Mode
      TRANSPARENT = 1
      OPAQUE = 2
    End Enum

Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
    Const IDANI_OPEN = &H1
    Const IDANI_CLOSE = &H2
    Const IDANI_CAPTION = &H3

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    Const PS_DASH = 1           ' -------
    Const PS_DASHDOT = 3        ' _._._._
    Const PS_DASHDOTDOT = 4     ' _.._.._
    Const PS_DOT = 2            ' .......
    Const PS_SOLID = 0
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As FloodFill) As Long
    Enum FloodFill
      FLOODFILLBORDER = 0
      FLOODFILLSURFACE = 1
    End Enum
Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

' DRAWING API'S
Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'Font API's
Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" _
    (ByVal lpFileName As String) As Long
Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" _
    (ByVal lpFileName As String) As Long

' TEXT API'S
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'regions
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'Constants used by the CombineRgn() API function.
Public Const RGN_AND = 1&
Public Const RGN_OR = 2&
Public Const RGN_XOR = 3&
Public Const RGN_DIFF = 4&
Public Const RGN_COPY = 5&

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'used in sub positioniconandlabel instead of giving it each time as a parameter
Private lIconwidth As Long

Function CreateMyFont(hDC As Long, ByVal sFace As String, Optional nSize As Integer = 8, _
                                       Optional nDegrees As Long = 0) As Long
    'Create a specified font
    CreateMyFont = CreateFont(-MulDiv(nSize, GetDeviceCaps(hDC, LOGPIXELSY), _
                              72), 0, nDegrees * 10, 0, FW_NORMAL, False, False, _
                              False, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
                              CLIP_DEFAULT_PRECIS, PROOF_QUALITY, _
                              DEFAULT_PITCH, sFace)
End Function

' be sure the form's scalemode is set to Pixels
Public Sub MakeTransLucent(Frm As Form, X As Long, Y As Long)
  If CBool(GetSetting("Translucency", "0")) Then
      Frm.Cls
      X = X / Screen.TwipsPerPixelX: Y = Y / Screen.TwipsPerPixelY
      AlphaBlending Frm.hDC, 0, 0, Frm.ScaleWidth, Frm.ScaleHeight, _
                    frmMain.hDC, X, Y, Frm.ScaleWidth, Frm.ScaleHeight, _
                    GetLong("TranslucencyLevel", 100)
  End If
End Sub

Public Sub AlphaBlending(ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest _
  As Long, ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, _
  ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight _
  As Long, ByVal AlphaSource As Long)

  Dim Blend As BLENDFUNCTION, BlendLng As Long

  Blend.SourceConstantAlpha = AlphaSource
  MoveMemory BlendLng, Blend, 4
    
  AlphaBlend destHDC, XDest, YDest, destWidth, destHeight, _
    srcHDC, xSrc, ySrc, srcWidth, srcHeight, BlendLng
End Sub

Public Function GetTextWidth(hDC As Long, ByVal sText As String) As Integer
  Dim Size As POINTAPI
  GetTextExtentPoint32 hDC, sText, Len(sText), Size
  GetTextWidth = Size.X
End Function
Public Function GetTextHeight(hDC As Long, ByVal sText As String) As Integer
  Dim Size As POINTAPI
  GetTextExtentPoint32 hDC, sText, Len(sText), Size
  GetTextHeight = Size.Y
End Function

' DESKTOP ICON MANAGEMENT
' ***********************
'extract file-icons, and puts them in an Image-control
Public Sub DrawIcon(sFilePath$, img As Image, Optional DrawFlags As Long)
  With frmMain.pictemp
    Dim shfi As SHFILEINFO, hImgLarge As Long, uFlags As Long
    
    BitBlt .hDC, 0, 0, img.Width, img.Height, img.Container.hDC, img.Left, img.Top, vbSrcCopy
    uFlags = SHGFI_LARGEICON Or SHGFI_SYSICONINDEX
    
    hImgLarge& = SHGetFileInfo(sFilePath, 0&, shfi, Len(shfi), uFlags)
    ImageList_Draw hImgLarge&, shfi.iIcon, .hDC, 0, 0, ILD_NORMAL Or ILD_TRANSPARENT Or DrawFlags
    
    img.Picture = .Image
    img.Refresh
  End With
End Sub

Public Sub DrawDrives()
    
    Dim Drives() As String, i As Integer
    
    Drives = DrivesPresent(True)
    'get the width of a desktop icon from windows
    lIconwidth = GetSystemMetrics(SM_CXICON)
    
    With frmMain
        'so background pic can be copied
        .AutoRedraw = True
        
        For i = 0 To UBound(Drives)
            If i > .imgDrives.UBound Then Load .imgDrives(i): Load .lblDrive(i)
            With .imgDrives(i)
                'size of icon
                .Width = lIconwidth: .Height = lIconwidth
                'position
                If i = 0 Then
                    .Left = frmMain.ScaleWidth / 2.5
                Else
                    .Left = frmMain.imgDrives(i - 1).Left + lIconwidth + 20
                End If
                .Top = 25
                .Tag = Drives(i)
                .Visible = False 'True
            End With
            .lblDrive(i) = Drives(i)
            .lblDrive(i).Move .imgDrives(i).Left + .imgDrives(i).Width / 2 - .lblDrive(i).Width / 2, .imgDrives(i).Top + .imgDrives(i).Height + 2
            .lblDrive(i).Visible = False 'True
            DrawIcon Drives(i), .imgDrives(i)
        Next
        
        .AutoRedraw = False
    End With

End Sub
Public Sub FillIcons(Optional FirstLoad As Boolean = False, _
                     Optional ChangeScreenRes As Boolean = False)
    
    Dim i As Integer, lForeColor As Long
    Dim lUbound As Integer, bShowDesktop As Boolean
       
    On Error Resume Next
    
    With frmMain
     'get the width of a desktop icon from windows
     lIconwidth = GetSystemMetrics(SM_CXICON)
     
     'turn off selection, if any
     .SelectIcon -1
     'so the background picture can be copied
     .AutoRedraw = True
     
     'set variabeles that are used in desktop arrow navigation
     IconsPerColumn = ((.ScaleHeight - 7) / (lIconwidth + 40)) - 1
     
     'if firstload then refresh main icons to
     If FirstLoad Then
          'load two icons if neccesary and position them
          If .imgIcon.UBound = 0 Then Load .imgIcon(1): Load .lblName(1): PositionIconAndLabel 1
       
          'these are for the main icons
          'You can change these icons from the Screen properties window
          .imgIcon(0).Tag = AppResourcePath & "Computer.lnk"
          .imgIcon(1).Tag = AppResourcePath & "Recycle.lnk"
       
                 
          For i = 0 To 1
           .imgIcon(i).Width = lIconwidth: .imgIcon(i).Height = lIconwidth + 2
           .lblName(i) = GetSetting(ColorNames(i + 13), Format(Colors(i + 13)), General)
           DrawIcon .imgIcon(i).Tag, .imgIcon(i)
          Next
     End If
     
     'only continue if other desktop items are to be shown
     bShowDesktop = CBool(GetSetting("ShowDesktopIcons", "1"))
     If bShowDesktop Then
           Dim Files() As String, DesktopDir As String
           DesktopDir = GetSpecialfolder(CSIDL_DESKTOP)
           Files = GetFilesFolders(DesktopDir, True, lUbound, 0)
           
           'set variabeles that are used in desktop arrow navigation
           IconsPerColumn = Int((.ScaleHeight - 7) / (lIconwidth + 40))
           Rows = Int((lUbound + 3) / IconsPerColumn) '0-based
     
           For i = 2 To lUbound + 2
            'load icons if neccesary
            If i > .imgIcon.UBound Then Load .imgIcon(i): Load .lblName(i)
            'if other file than currently displayed then
            If .imgIcon(i).Tag <> DesktopDir & Files(i - 2) Or ChangeScreenRes Then
              'set label and put file in tag of image
              .lblName(i) = CheckExtension(Files(i - 2))
              .imgIcon(i).Tag = DesktopDir & Files(i - 2)
              PositionIconAndLabel i
            End If
            'redraw the icons
            DrawIcon DesktopDir & Files(i - 2), .imgIcon(i)
           Next
           'erase array from memory
           Erase Files
     End If
     
     'set forecolor of all labels
     lForeColor = GetLong(ColorNames(11), Colors(11))
     For i = 0 To .lblName.UBound
       .lblName(i).ForeColor = lForeColor
       'unload icons that are no longer needed
       If i >= lUbound + IIf(bShowDesktop, 3, 2) Then Unload .imgIcon(i): Unload .lblName(i)
     Next
     
     'the current image is made background (cls won't delete it)
     .AutoRedraw = False
    End With
End Sub

'Not used for i=0 and i=1
Sub PositionIconAndLabel(i As Integer)
    
    'position icon in row, in next row if neccesary
    With frmMain.imgIcon(i)
      Dim imgPrev As Image
      Set imgPrev = frmMain.imgIcon(i - 1)
      
      .Width = lIconwidth: .Height = lIconwidth + 2
      If (i Mod IconsPerColumn) = 0 Then
        .Move imgPrev.Left + 90, 7
      Else
        .Move imgPrev.Left, imgPrev.Top + lIconwidth + 40
      End If
    End With
    
    'how is the text in the label to be displayed
    With frmMain.lblName(i)
    
        lTextWidth = GetTextWidth(frmMain.hDC, .Caption)
        If lTextWidth > 85 Then
          .Width = 85
          lHeight = Int(lTextWidth / 85)
          If lHeight > 2 Then lHeight = 2
          .WordWrap = True
          .Height = (lHeight + 1) * 13
          .ToolTipText = .Caption
        End If
        
    End With
    
    'move label under imgicon and center it
    With frmMain
      .lblName(i).Move .imgIcon(i).Left + (lIconwidth - .lblName(i).Width) * 0.5, _
                       .imgIcon(i).Top + .imgIcon(i).Height
      .imgIcon(i).Visible = True: .lblName(i).Visible = True
    End With

End Sub

Sub MakeMenuItems(hDC As Long, sPath As String, r As RECT, ByVal Active _
    As Boolean, ByVal DrawArrow As Boolean, Optional bSeperator As Boolean, _
    Optional lIconHandle As Long = -1, Optional obj As Object = Nothing)
    
    On Error Resume Next

    Dim hImgLarge&, shfi As SHFILEINFO
    Dim hPen As Long, hBrush As Long, hFont As Long
    Dim hOldBrush As Long, hOldFont As Long, hOldPen As Long
    Dim sPrint As String

    'Fill Background
    hBrush = CreateSolidBrush(IIf(Active, Colors(2), Colors(0)))
    hOldBrush = SelectObject(hDC, hBrush)
    FillRect hDC, r, hBrush
    DeleteObject SelectObject(hDC, hOldBrush)

    If bSeperator Then '"-"
        hPen = CreatePen(PS_SOLID, 1, &HFFFFFF)
        hOldPen = SelectObject(hDC, hPen)

        MoveToEx hDC, r.Left, r.Top + 2, 0&
        LineTo hDC, r.Right, r.Top + 2
        DeleteObject SelectObject(hDC, hOldPen)
    Else
        'DrawIcon
        If lIconHandle = -1 Then
            If InStr(1, sPath, "\") <> 0 Then
                hImgLarge& = SHGetFileInfo(sPath, 0&, shfi, Len(shfi), _
                        SHGFI_SMALLICON Or BASIC_SHGFI_FLAGS)
                ImageList_Draw hImgLarge&, shfi.iIcon, hDC, r.Left + 2, 2 + r.Top, ILD_TRANSPARENT
                DeleteObject hImgLarge
            End If
        ElseIf IconHandle > 0 Then
            DrawIconEx hDC, r.Left + 2, 2 + r.Top, lIconHandle, 16, 16, 0, 0, DI_NORMAL
        End If

        'if our startmenu, then check length
        If Not (obj Is Nothing) Then
            lTextW = GetTextWidth(hDC, sPrint)
            If lTextW > obj.ScaleWidth Then sPrint = Left(sPrint, 45) & "..."
        End If
        'DrawText
        SetBkMode hDC, TRANSPARENT
        SetTextColor hDC, IIf(Active, Colors(3), Colors(1))
        sPrint = ExtractFilename(sPath, True)
        'select our font into hdc
        hFont = CreateMyFont(hDC, Colors(12))
        hOldFont = SelectObject(hDC, hFont)
        TextOut hDC, r.Left + 22, r.Bottom - 10 - (GetTextHeight(hDC, "S")) * 0.5, sPrint, Len(sPrint)
        DeleteObject SelectObject(hDC, hOldFont)
        
        'DrawArrow
        If DrawArrow Then
            'outline
            hPen = CreatePen(PS_SOLID, 1, IIf(Active, Colors(4), Colors(5)))
            hOldPen = SelectObject(hDC, hPen)
            X = r.Right - 12
            Y = r.Bottom - 10
            DrawPolygon hDC, 3, X, Y - 4, X + 4, Y, X, Y + 4
            DeleteObject SelectObject(hDC, hOldPen)
            'fill arrow
            If bFillArrow Then
                hBrush = CreateSolidBrush(IIf(Active, Colors(6), Colors(7)))
                hOldBrush = SelectObject(hDC, hBrush)
                ExtFloodFill hDC, X + 1, Y, IIf(Active, Colors(4), Colors(5)), FLOODFILLBORDER
                DeleteObject SelectObject(hDC, hOldBrush)
            End If
        End If
    End If
    If Not (obj Is Nothing) Then
        obj.Refresh
    Else
        Call ExcludeClipRect(hDC, r.Left, r.Top, r.Right, r.Bottom)
    End If
End Sub

Public Sub MakeFormRounded(Frm As Form, Radius As Long)
    Dim hRgn As Long
    hRgn = CreateRoundRectRgn(0, 0, Frm.ScaleWidth, Frm.ScaleHeight, Radius, Radius)
    SetWindowRgn Frm.hwnd, hRgn, True
    Call DeleteObject(hRgn)
End Sub

'I USE THIS TO DRAW TRIANGLES
Public Sub DrawPolygon(hDC As Long, nCount As Integer, ParamArray XYCouples() As Variant)
    
    MoveToEx hDC, XYCouples(0), XYCouples(1), 0&
    For i = 2 To 2 * nCount - 1 Step 2
        LineTo hDC, XYCouples(i), XYCouples(i + 1)
    Next
    LineTo hDC, XYCouples(0), XYCouples(1)
End Sub
