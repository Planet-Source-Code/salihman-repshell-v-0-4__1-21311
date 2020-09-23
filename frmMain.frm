VERSION 5.00
Object = "{6A27B64A-0A70-11D5-A06F-86B0E384F25B}#1.0#0"; "REPCONTROLS.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404000&
   BorderStyle     =   0  'None
   Caption         =   "0"
   ClientHeight    =   7005
   ClientLeft      =   1380
   ClientTop       =   465
   ClientWidth     =   7320
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picWallPaper 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4800
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picBackTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4680
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin RepControls.ctlAnimatedButton anmShutdown 
      Height          =   750
      Index           =   0
      Left            =   6120
      TabIndex        =   5
      Top             =   6120
      Width           =   1050
      _ExtentX        =   106
      _ExtentY        =   450
      FramesPerSecond =   50
      FrameCount      =   5
   End
   Begin VB.PictureBox picQuickIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   6240
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox picQuickIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   6600
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1920
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtRename 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   615
      Visible         =   0   'False
      Width           =   1215
   End
   Begin RepControls.ctlAnimatedButton anmShutdown 
      Height          =   750
      Index           =   1
      Left            =   5040
      TabIndex        =   6
      Top             =   6120
      Width           =   1050
      _ExtentX        =   106
      _ExtentY        =   450
      FramesPerSecond =   50
      FrameCount      =   5
   End
   Begin RepControls.ctlAnimatedButton anmShutdown 
      Height          =   750
      Index           =   2
      Left            =   3960
      TabIndex        =   7
      Top             =   6120
      Width           =   1050
      _ExtentX        =   106
      _ExtentY        =   450
      FramesPerSecond =   50
      FrameCount      =   5
   End
   Begin VB.Label lblDrive 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4275
      TabIndex        =   10
      Top             =   615
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgDrives 
      Height          =   480
      Index           =   0
      Left            =   4080
      Top             =   105
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIcon 
      Height          =   510
      Index           =   0
      Left            =   480
      Stretch         =   -1  'True
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400040&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   690
      TabIndex        =   0
      Top             =   615
      Width           =   60
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'                        RepShell v 0.4 - made in VB6
'                      Copyright(c) 2001 Salih Gunaydin
'                        Bug - Hunter: Wouter Tollet
'                    Email: wippo@antwerp.crosswinds.net
'*****************************************************************************

Private PrevIcon As Integer  'the index of the icon we are on
Private CurRow As Integer    'the row in which the previcon is located
Private IChNameNr As Integer 'index of icon which we're changing the name of


Private Sub anmShutdown_Click(Index As Integer)
    Select Case Index
    Case 2: ShowRunDialog
    Case 1
        If MsgBox("This really works, so if you continue, " & vbCrLf _
                  & "RepShell will logoff." & vbCrLf & vbCrLf & _
                  "Are you sure?", vbInformation + vbOKCancel) = vbOK Then
            ExitApp
            ExitWindowsEx EWX_LOGOFF, 0&
        End If
    Case 0
        If MsgBox("This really works, so if you continue, " & vbCrLf _
                  & "the computer will shutdown." & vbCrLf & vbCrLf & _
                  "Are you sure?", vbInformation + vbOKCancel) = vbOK Then
            ExitApp
            ExitWindowsEx EWX_SHUTDOWN, 0&
        End If
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Tooltip.DeleteTool picQuickIcon(0).hwnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmMain = Nothing
End Sub
Private Sub Form_Click()
    SelectIcon -1
End Sub

' arrow navigation, Enter to execute, escape to unload, F5 to refresh
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
        'if coming from lower icon, go one up
        If PrevIcon > 0 Then SelectIcon PrevIcon - 1
    Case vbKeyDown
        'if coming from up, go one down, if no down then goto prev row if exist
        If PrevIcon < imgIcon.UBound Then
            SelectIcon PrevIcon + 1
        ElseIf CurRow > 0 Then
            SelectIcon PrevIcon - IconsPerColumn + 1
        End If
    Case vbKeyLeft
        'jumpt to prev row

        If IconsPerColumn < imgIcon.UBound + 1 And CurRow > 0 Then
            SelectIcon PrevIcon - IconsPerColumn
        End If
    Case vbKeyRight
        'jump to next row
        If IconsPerColumn < imgIcon.UBound + 1 And CurRow < Rows Then
            SelectIcon PrevIcon + IconsPerColumn
        End If
    Case vbKeyReturn
        'execute icon
        Call imgIcon_DblClick(PrevIcon)
    Case vbKeyF5 'refresh desktop
        Init
    Case 93
        'show context menu
        'if not on this computer then show shell context menu
        If PrevIcon > 0 Then
          Dim pt As POINTAPI
          With imgIcon(PrevIcon)
            pt.X = .Left + .Width * 0.5: pt.Y = .Top + .Height * 0.5
          End With
          Call ShellContextMenu(imgIcon(PrevIcon), pt, PrevIcon = 1)
        'start menu
        Else
          frmStart.GetMenu "", , True
        End If
  End Select
End Sub

Private Sub Form_Load()
    
    PrevIcon = -1
    Init True, True
    Show
    
    'init tooltip
    With Tooltip
        .Create hwnd, ttfBalloon Or ttfAlwaysTip

        .SetDelayTime sdtInitial, 100
        .SetDelayTime sdtAutoPop, 500
        .SetDelayTime sdtReshow, 100

        .Icon = itInfoIcon
        .Title = "Connection Info"
        .FontFace = "Courier New"
        'the objects you want to add a tooltip to
        sData = GetConnectionData
        .AddTool picQuickIcon(1).hwnd, 0, _
                "Connected to     : " & sData(0) & vbCrLf & _
                "Connection speed : " & sData(1) & vbCrLf & _
                "Connection time  : " & sData(2) & vbCrLf & _
                vbCrLf & _
                "Local IP         : " & sData(3) & vbCrLf & _
                "Remote IP        : " & sData(4) & vbCrLf & _
                "Remote Host      : " & sData(5)
        .Enabled = True
    End With

    Hook                'start subclass
    SHNotify_Register   'setup for desktop refreshing
  
End Sub

Sub Init(Optional RefreshDesktop As Boolean = True, _
         Optional SizeAndPosChange As Boolean = False)
  
  On Error Resume Next

    'to be able to clear background
    AutoRedraw = True
    
    'because startmenu checks main form for textwidths
    Font.Name = Colors(12): Font.Size = 8
    
    'retrieve handle OF OUR SYSTRAY again, in case of ...
    'ThunderRTForm6Dc is the class name for VB6 forms
    Do
        lSystrayHwnd = FindWindow("ThunderRT6FormDc", "RepShell_Tray_Wnd")
        DoEvents
    Loop Until lSystrayHwnd <> 0
    'Set Systraycolor
    SendMessage lSystrayHwnd, WM_CHANGEBACKCOLOR, Colors(9), 0
    'make the window topmost
    MakeTopMost lSystrayHwnd, True
    
    'if the screen resolution has changed, move items if neccesary
    If SizeAndPosChange Then
        'set new size and pos
        Move 0, 0, Screen.Width, Screen.Height

        'SetWindowPos hwnd, HWND_BOTTOM, 0, 0, Screen.Width / _
                     Screen.TwipsPerPixelX, Screen.Height / _
                     Screen.TwipsPerPixelY, 0
        
        picBackTemp.Width = ScaleWidth
        picBackTemp.Height = ScaleHeight
        
    End If
  
    'does desktop need refreshing
    If RefreshDesktop Then
        
        Dim sWallPaper$, sWallPaperStyle$, sTile$
        
        picBackTemp.Cls: Cls
        BackColor = GetLong(ColorNames(8), Colors(8))
        picBackTemp.BackColor = BackColor
        
        sWallPaper = ReadString(HKEY_CURRENT_USER, "Control Panel\Desktop\", "WallPaper", "")
        sWallPaperStyle = ReadString(HKEY_CURRENT_USER, "Control Panel\Desktop\", "WallPaperStyle", "0")
        sTile = ReadString(HKEY_CURRENT_USER, "Control Panel\Desktop\", "TileWallPaper", "0")
        
        'if there's wallpaper
        If Len(sWallPaper) > 0 Then
            
            'load the bitmap in a picbox, which autosizes to picture
            picWallPaper.Picture = LoadPicture(sWallPaper)
            
            'Which wallpaper style
            If sWallPaperStyle = "2" Then
                'Stretch picture in background
                StretchBlt picBackTemp.hDC, 0, 0, ScaleWidth, ScaleHeight, _
                           picWallPaper.hDC, 0, 0, picWallPaper.ScaleWidth, _
                           picWallPaper.ScaleHeight, vbSrcCopy
            Else
                'tile the bitmap
                If sTile = "1" Then
                    'tile the bitmap
                    For i = 0 To Int(ScaleWidth / picWallPaper.ScaleWidth)
                        For j = 0 To Int(ScaleHeight / picWallPaper.ScaleHeight)
                            BitBlt picBackTemp.hDC, i * picWallPaper.ScaleWidth, _
                            j * picWallPaper.ScaleHeight, picWallPaper.ScaleWidth, _
                            picWallPaper.ScaleHeight, picWallPaper.hDC, 0, 0, vbSrcCopy
                        Next
                    Next
                Else
                    'just blt the bitmap in the center
                    BitBlt picBackTemp.hDC, (ScaleWidth - picWallPaper.ScaleWidth) / 2, _
                          (ScaleHeight - picWallPaper.ScaleHeight) / 2, _
                          picWallPaper.ScaleWidth, picWallPaper.ScaleHeight, _
                          picWallPaper.hDC, 0, 0, vbSrcCopy
                End If
            End If
            
            'if translucency is enabled
            If CBool(GetSetting("Translucency", "1")) Then

                AlphaBlending hDC, 0, 0, ScaleWidth, ScaleHeight, _
                        picBackTemp.hDC, 0, 0, ScaleWidth, ScaleHeight, _
                        GetLong("TranslucencyLevel", 100)
            Else
                'blt the result on our form
                BitBlt hDC, 0, 0, ScaleWidth, ScaleHeight, picBackTemp.hDC, 0, 0, vbSrcCopy
            End If
            Refresh
        End If
        
    End If 'End If DesktopColorChange
  
    'start value for the number of quickicons to show
    iCounter = -1
    'position QuickIcons
    For i = 0 To 1
      picQuickIcon(i).Visible = False
      bShow = GetLong("QuickIconVisible" & i, 1)
      'show the quickicon
      If bShow Then
        iCounter = iCounter + 1
        'Move control into position
        If iCounter = 0 Then
            picQuickIcon(i).Move ScaleWidth - 50, 17
        Else
            picQuickIcon(i).Move picQuickIcon(i - 1).Left - 22, 17
        End If
        'Draw Icon in quickicon control
        DrawQuickIcon GetSetting(ColorNames(15 + i), Colors(15 + i)), i
        picQuickIcon(i).Visible = True
      End If
    Next
    
    'position Wanimated buttons
    For i = 0 To 2
      With anmShutdown(i)
        'sets the width and height
        .AnimFileLocation = AppResourcePath & "Bitmaps\" & IIf(i = 0, _
                            "Shutdown.bmp", IIf(i = 1, "LogOff.bmp", "Run.bmp"))
        .FrameCount = 5
        .FramesPerSecond = 25
        'Move control into position
        If i = 0 Then
            .Move ScaleWidth - .Width - 50, ScaleHeight - .Height - 50
        Else
            .Move anmShutdown(i - 1).Left - .Width, ScaleHeight - .Height - 50
        End If
        'blt part of background in anim buttons hdc
        BitBlt .BackGroundHdc, 0, 0, .FrameWidth, .FrameHeight, hDC, .Left, .Top, vbSrcCopy
        'redraw the image
        .paintImage
      End With
    Next
    
    'to make current picture the background
    AutoRedraw = False
    
    'we send the size sizeandposchange var to reposition the desktopicons
    'if neccesary
    FillIcons True, SizeAndPosChange
    frmTask.Init

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  If Button = vbRightButton Then
    Dim MenuItems()
    MenuItems = Array("Log off", "Shut Down", "Restart", "-", "Exit RepShell", _
                "-", "Paste", "-", "Options", "RepShell Options", "-", _
                "Control Panel", "Printers", "Screen", "Background", _
                "Screensaver", "Options", "Settings")
    ReDim SubMenuNo(UBound(MenuItems)), MemberOfSubNo(UBound(MenuItems))
    SubMenuNo(8) = 1: SubMenuNo(13) = 2
    For i = 9 To 13
      MemberOfSubNo(i) = 1
    Next
    For i = 14 To 17
      MemberOfSubNo(i) = 2
    Next
    MakeAPIMenu MenuItems, SubMenuNo, MemberOfSubNo, 2, "Shut Down"
  End If
End Sub


'-----------------------------------
'DESKTOP ICON MANAGMENT
'-----------------------------------
Private Sub imgIcon_DblClick(Index As Integer)
    If Index <> 0 Then
        ExecuteFile imgIcon(Index).Tag
    Else
        Shell "explorer.exe ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", vbNormalFocus
    End If
End Sub
Private Sub ImgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pressed the same icon
    If PrevIcon = Index Then Exit Sub
    SelectIcon Index
End Sub
Private Sub imgIcon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
     If Index <> 0 Then
         Dim pt As POINTAPI
         GetCursorPos pt
         Call ShellContextMenu(imgIcon(Index), pt, Index = 1)
     Else
         frmStart.GetMenu "", , True
     End If
    End If
End Sub

Private Sub lblName_DblClick(Index As Integer)
    Call imgIcon_DblClick(Index)
End Sub
Private Sub lblName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PrevIcon = Index Then
        If Button = vbLeftButton Then DesktopRenameShow
    Else
        Call ImgIcon_MouseDown(Index, Button, Shift, X, Y)
    End If
End Sub
Private Sub lblName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgIcon_MouseUp(Index, Button, Shift, X, Y)
End Sub
'Function that draws an icon selected an other deselects an other one
Public Sub SelectIcon(ByVal nIndex As Integer)
                      
    On Error Resume Next
    'restore previous icon
    If PrevIcon <> -1 Then
        DrawIcon imgIcon(PrevIcon).Tag, imgIcon(PrevIcon)
        lblName(PrevIcon).BackStyle = 0
    End If
    If nIndex <> -1 Then
        If nIndex > imgIcon.UBound Then nIndex = imgIcon.UBound
        If nIndex < 0 Then nIndex = 0
        'draw new icon selected
        DrawIcon imgIcon(nIndex).Tag, imgIcon(nIndex), ILD_BLEND50
        lblName(nIndex).BackStyle = 1
        lblName(nIndex).BackColor = GetLong(ColorNames(10), Colors(10))
    End If
    PrevIcon = nIndex
    CurRow = Int(nIndex / IconsPerColumn)
End Sub

'show textbox where the user can input a new name
Public Sub DesktopRenameShow()
    'On Error Resume Next
    
    IChNameNr = PrevIcon
    SelectIcon -1
    KeyPreview = False 'if we use cursors within textbox
                       'the desktop won't be affected
    With txtRename
     .Tag = imgIcon(IChNameNr).Tag
     
     .Move lblName(IChNameNr).Left - 2, lblName(IChNameNr).Top - 1, _
           lblName(IChNameNr).Width + 4, lblName(IChNameNr).Height + 2
     
     If IChNameNr > 1 Then
        .Text = ExtractFilename(.Tag, True)
     Else
        .Text = Colors(13 + IChNameNr)
     End If
     
     .SelStart = 0: .SelLength = Len(txtRename)
     .Visible = True
     
     SetCapture .hwnd
     .SetFocus
     
    End With
End Sub

'actual renaming of file
Private Sub txtRename_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
      
      'If we renamed a file and not the computer or recycle icon
      If IChNameNr > 1 Then
        Dim CurExt As String
        'if changed extension by mistake
        CurExt = GetExtension(txtRename.Tag)
        
        If (GetExtension(txtRename) <> CurExt And Not (CurExt = ".lnk" Or _
            CurExt = ".pif" Or CurExt = ".url") And KeyAscii = vbKeyReturn) Then
          
            'ask if you want to proceed with changing extension
            If MsgBox("By changing the extension the file can become useless." _
                      & vbCrLf & vbCrLf & _
                      "Do you want to continue?", vbYesNo + vbExclamation, "Rename") _
                      = vbNo Then
                      
                SetCapture txtRename.hwnd
                Exit Sub
            End If
              
        End If
      End If
      
      ReleaseCapture
      txtRename.Visible = False
      KeyPreview = True
      
      If KeyAscii = vbKeyReturn And txtRename <> "" Then
        'if newly entered name is null jsut exit sub
        
        'if we renamed a file on desktop
        If IChNameNr > 1 Then
            Dim sDesktop As String, lRet As Long
            
            'Put name in variable
            sDesktop = GetSpecialfolder(CSIDL_DESKTOP) & Trim(txtRename) & _
                       IIf(CurExt = ".lnk" Or CurExt = ".pif" Or CurExt = _
                       ".url", CurExt, "")
                       
            If txtRename.Tag <> sDesktop Then
                'if we did change name, rename the file
                Dim sh As SHFILEOPSTRUCT
                With sh
                    .wFunc = FO_RENAME
                    .fFlags = FOF_SILENT
                    .pFrom = txtRename.Tag
                    .pTo = sDesktop
                End With
                lRet = SHFileOperation(sh)
                If lRet Then txtRename = ExtractFilename(txtRename.Tag)
            End If
            'put filename in tag
            imgIcon(IChNameNr).Tag = sDesktop
        Else
            'if we renamed a main icon
            Colors(13 + IChNameNr) = Trim(txtRename)
            SaveSetting ColorNames(13 + IChNameNr), Trim(txtRename)
        End If
        'refresh the label on desktop
        lblName(IChNameNr) = txtRename
        
        'select the icon
        SelectIcon IChNameNr
        
      End If
    End If
End Sub
Private Sub txtRename_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCapture txtRename.hwnd
End Sub
Private Sub txtRename_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call txtRename_KeyPress(vbKeyReturn)
End Sub
'-----------------------------------
'END DESKTOP ICON MANAGMENT
'-----------------------------------

'-----------------------------------
'QUICKICON MANAGMENT
'-----------------------------------
Private Sub picQuickIcon_DblClick(Index As Integer)
    Dim sTask As String
    
    If Index = 0 Then
        sTask = "SNDVOL32.exe"
    Else
        Exit Sub
    End If
    ExecuteFile sTask
End Sub
Private Sub picQuickIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If Button = vbRightButton Then
        Dim Entries
        Select Case Index
        Case 1
            Entries = GetEntries
            MakeAPIMenu Entries, , , , "Ras"
        End Select
    End If
End Sub
Sub DrawQuickIcon(ByVal sFilePath$, ByVal IndexOfPic%)
    Dim hIco As Long
    hIco = ExtractIcon(0, sFilePath, 0)
    With picQuickIcon(IndexOfPic)
        BitBlt .hDC, 0, 0, 16, 16, frmMain.hDC, .Left, .Top, vbSrcCopy
        DrawIconEx .hDC, 0, 0, hIco, 16, 16, 0, 0, DI_NORMAL
    End With
    DestroyIcon hIco
End Sub
'-----------------------------------
'END QUICKICON MANAGMENT
'-----------------------------------
