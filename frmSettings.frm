VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404000&
   BorderStyle     =   0  'None
   Caption         =   "RepShell Settings"
   ClientHeight    =   6285
   ClientLeft      =   5040
   ClientTop       =   3915
   ClientWidth     =   7020
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSideBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4950
      Left            =   120
      ScaleHeight     =   330
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   39
      Top             =   600
      Width           =   1380
      Begin VB.Image imgIcon 
         Height          =   510
         Index           =   4
         Left            =   450
         Top             =   3840
         Width           =   480
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   44
         Top             =   4440
         Width           =   420
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Behaviour"
         Height          =   195
         Index           =   3
         Left            =   330
         TabIndex        =   43
         Top             =   3540
         Width           =   720
      End
      Begin VB.Image imgIcon 
         Height          =   510
         Index           =   3
         Left            =   450
         Top             =   3000
         Width           =   480
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desktop"
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   42
         Top             =   2580
         Width           =   600
      End
      Begin VB.Image imgIcon 
         Height          =   510
         Index           =   2
         Left            =   450
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menu Settings"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   41
         Top             =   1620
         Width           =   1020
      End
      Begin VB.Image imgIcon 
         Height          =   510
         Index           =   1
         Left            =   450
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desktop Items"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   40
         Top             =   660
         Width           =   1020
      End
      Begin VB.Image imgIcon 
         Height          =   510
         Index           =   0
         Left            =   450
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   340
      Left            =   5610
      TabIndex        =   29
      Top             =   5760
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   4305
      TabIndex        =   1
      Top             =   5760
      Width           =   1212
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   3000
      TabIndex        =   0
      Top             =   5760
      Width           =   1212
   End
   Begin VB.PictureBox picSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4950
      Index           =   4
      Left            =   1500
      ScaleHeight     =   330
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   45
      Top             =   600
      Visible         =   0   'False
      Width           =   5400
      Begin VB.Label lblComments 
         AutoSize        =   -1  'True
         Caption         =   "Click to open Comments.txt"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2880
         TabIndex        =   52
         Top             =   4560
         Width           =   1920
      End
      Begin VB.Label lblReadme 
         AutoSize        =   -1  'True
         Caption         =   "Click to open ReadMe.txt"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   600
         TabIndex        =   51
         Top             =   4560
         Width           =   1800
      End
      Begin VB.Label lblLink 
         AutoSize        =   -1  'True
         Caption         =   "salih@belgacom.net"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   600
         TabIndex        =   47
         Top             =   4320
         Width           =   1440
      End
      Begin VB.Label lblInfo 
         Height          =   4215
         Left            =   600
         TabIndex        =   46
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.PictureBox picSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4950
      Index           =   2
      Left            =   1500
      ScaleHeight     =   330
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   5400
      Begin VB.Frame Frame2 
         Caption         =   "Colors"
         Height          =   1935
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   5055
         Begin VB.ListBox lstColor 
            Height          =   1425
            Index           =   1
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "&Import Standard Settings file"
            Height          =   495
            Left            =   2760
            TabIndex        =   36
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton cmdExport 
            Caption         =   "&Export Current Settings"
            Height          =   495
            Left            =   2760
            TabIndex        =   30
            ToolTipText     =   "Exports current color settings and several other settings to the Standard Settings File"
            Top             =   720
            Width           =   2175
         End
         Begin VB.CommandButton cmdColor 
            BackColor       =   &H0000C000&
            Height          =   255
            Index           =   1
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   380
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Translucency"
         Height          =   1815
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   5055
         Begin VB.CheckBox chkTranslucency 
            Caption         =   "Enable Translucency"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   3195
         End
         Begin VB.PictureBox picTrans 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            Height          =   480
            Left            =   4200
            ScaleHeight     =   28
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   28
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   480
            Width           =   480
         End
         Begin VB.PictureBox pictemp 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            Height          =   480
            Left            =   4200
            ScaleHeight     =   420
            ScaleWidth      =   420
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   480
         End
         Begin MSComctlLib.Slider sliTranslucency 
            Height          =   375
            Left            =   360
            TabIndex        =   22
            Top             =   1200
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   661
            _Version        =   393216
            LargeChange     =   51
            SmallChange     =   5
            Max             =   255
            TickFrequency   =   26
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Opaque"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   26
            Top             =   960
            Width           =   585
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Transparent"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   2760
            TabIndex        =   25
            Top             =   960
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Preview:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   4140
            TabIndex        =   24
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Translucency Level:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   1515
         End
      End
   End
   Begin VB.PictureBox picSettings 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4950
      Index           =   0
      Left            =   1500
      ScaleHeight     =   330
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   600
      Width           =   5400
      Begin VB.Frame Frame3 
         Caption         =   "QuickIcons"
         Height          =   1335
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   4935
         Begin VB.CheckBox chkQuickIcon 
            Caption         =   "NetConnect Icon"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkQuickIcon 
            Caption         =   "Sound Icon"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   285
            Index           =   1
            Left            =   4545
            TabIndex        =   35
            Top             =   720
            Width           =   285
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   285
            Index           =   0
            Left            =   4545
            TabIndex        =   34
            Top             =   360
            Width           =   285
         End
         Begin VB.TextBox txtIconLoc 
            Height          =   285
            Index           =   1
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtIconLoc 
            Height          =   285
            Index           =   0
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.CheckBox chkShowDesktopIcons 
         Caption         =   "Show Desktop Icons"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   3855
      End
   End
   Begin VB.PictureBox picSettings 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4950
      Index           =   3
      Left            =   1500
      ScaleHeight     =   330
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   5400
      Begin VB.CheckBox chkDefaultShell 
         Caption         =   "Make RepShell default Shell"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
      Begin VB.OptionButton optTime 
         Caption         =   "24 Hour"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   8
         Top             =   840
         Width           =   1155
      End
      Begin VB.OptionButton optTime 
         Caption         =   "12 Hour"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   7
         Top             =   840
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Clock Format:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   600
         TabIndex        =   9
         Top             =   840
         Width           =   1035
      End
   End
   Begin VB.PictureBox picSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4950
      Index           =   1
      Left            =   1500
      ScaleHeight     =   4950
      ScaleWidth      =   5400
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   5400
      Begin VB.ComboBox cmbFont 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2760
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   1440
         Width           =   2055
      End
      Begin VB.ListBox lstColor 
         Height          =   1815
         Index           =   0
         Left            =   360
         TabIndex        =   38
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox chkShowHiddenFiles 
         Caption         =   "Show hidden folders in the menus"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   3240
         Width           =   3015
      End
      Begin VB.CheckBox chkFillArrow 
         Caption         =   "Fill Arrow"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H000040C0&
         Height          =   345
         Index           =   0
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   2055
      End
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1560
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   145
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   315
         Width           =   2175
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Menu Example"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Set Menu Colors"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1170
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RepShell Settings"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   7
      Left            =   240
      TabIndex        =   2
      Top             =   105
      Width           =   2190
   End
   Begin VB.Line lin1 
      BorderColor     =   &H00C0C0C0&
      Index           =   4
      X1              =   8
      X2              =   464
      Y1              =   31
      Y2              =   31
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OldTab As Integer
'Temp settings
Private TempColors(11) As Long
Private TempFont As String
Private bTempFillArrow As Boolean
Private bDraw As Boolean


'this enables/disables the matching textbox and command button
Private Sub chkQuickIcon_Click(Index As Integer)
    txtIconLoc(Index).Enabled = chkQuickIcon(Index).Value
    cmdBrowse(Index).Enabled = chkQuickIcon(Index).Value
End Sub


'Move borderless form
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
       ReleaseCapture
       SendMessage hwnd, WM_SYSCOMMAND, WM_MOVE, 0
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'clear region
    SetWindowRgn hwnd, 0&, True
    Set frmSettings = Nothing
End Sub
'If you press escape the form unloads, just like pressing cancel
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call cmdCancel_Click
End Sub



Private Sub cmdOk_Click()
    ApplyChanges True
End Sub
Private Sub cmdApply_Click()
    ApplyChanges False
End Sub
Private Sub cmdCancel_Click()
    'return all settings
    For i = 0 To 11
        Colors(i) = TempColors(i)
    Next
    bFillArrow = bTempFillArrow
    Colors(12) = TempFont
    
    Unload Me
End Sub
'Function that saves all changes and unloads the form if neccesary
Sub ApplyChanges(Unloadme As Boolean)

  Dim bTemp As Boolean, i As Integer
  
  '9:32 PM - 21:32
  sClockFormat = IIf(optTime(0).Value, "h:mm AMPM", "hh:mm")

  'conditions to check if anything changed in the desktop
  If Val(GetSetting("Translucency", "0")) <> chkTranslucency.Value Then bTemp = True
  If GetLong("TranslucencyLevel", 100) <> sliTranslucency.Value Then bTemp = True
  
  'save checkboxes
  SaveSetting "ShowDesktopIcons", chkShowDesktopIcons.Value
  SaveSetting "ShowHiddenFiles", chkShowHiddenFiles.Value
  SaveSetting "Translucency", chkTranslucency.Value
  SaveLong "TranslucencyLevel", sliTranslucency.Value
  SaveSetting "ClockFormat", IIf(optTime(0).Value, "12", "24")
  SaveSetting "FillArrow", chkFillArrow.Value

  For i = 0 To 1
      SaveLong "QuickIconVisible" & i, chkQuickIcon(i).Value
  Next
    
  'Save New Colors
  For i = 0 To 11
    Call SaveLong(ColorNames(i), Colors(i))
  Next
  'save font,computername, recyclebinname
  For i = 12 To 14
    SaveSetting ColorNames(i), Colors(i)
  Next
  For i = 0 To 1
    Call SaveSetting(ColorNames(15 + i), IIf(Dir(txtIconLoc(i)) = "", _
         Colors(i + 15), txtIconLoc(i)))
  Next
  
  'if repshell is set to be deafult shell
  If chkDefaultShell.Value Then
   'set the change in the system.ini file
    WritePrivateProfileString "boot", "shell", AppPath & App.EXEName & ".exe", "system.ini"
  Else
    WritePrivateProfileString "boot", "shell", "explorer.exe", "system.ini"
  End If
  
  'Apply Changes
  'change desktop background if backcolor is changed or if translucency is changed
  frmMain.Init (Colors(8) <> frmMain.BackColor) Or bTemp
  If Unloadme Then
    Unload Me
  Else
    Show
  End If
End Sub



Private Sub chkFillArrow_Click()
' This variable is changed on the spot to show it in the example
    bFillArrow = CBool(chkFillArrow.Value)
End Sub

Private Sub chkTranslucency_Click()
    sliTranslucency.Enabled = chkTranslucency.Value
End Sub

'start the default mail program and open new mail with my address filled in
Private Sub lblLink_Click()
    Shell "start mailto:salih@belgacom.net?Subject=RepShell", vbHide
End Sub
Private Sub lblReadme_Click()
    ExecuteFile AppPath & "Readme.txt"
End Sub
Private Sub lblComments_Click()
    ExecuteFile AppPath & "Comments.txt"
End Sub


Private Sub cmbFont_Click()
    Colors(12) = cmbFont.Text
End Sub
Private Sub lstColor_Click(Index As Integer)
    'index=0 = menucolors; index=1 = desktopcolors
    If Index = 0 And lstColor(0).ListIndex = 8 Then
        cmbFont.Enabled = True
        cmdColor(0).Enabled = False
    Else
        cmbFont.Enabled = False
        cmdColor(0).Enabled = True
        cmdColor(Index).BackColor = Colors(lstColor(Index).ListIndex + IIf(Index, 8, 0))
    End If
End Sub
Private Sub cmdColor_Click(Index As Integer)
  'index=0 = menucolors; index=1 = desktopcolors
  Dim lTempColor As Long
  lTempColor = cmdColor(Index).BackColor
  RetVal = ShowColor(lTempColor, hwnd)
  If RetVal Then
    cmdColor(Index).BackColor = lTempColor
    Colors(lstColor(Index).ListIndex + IIf(Index, 8, 0)) = lTempColor
  End If
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Dim lRet As Boolean, sFile As String
        
    On Error Resume Next
    lRet = ShowOpen(sFile, , , , , True, "Icon (*.ico)|*.ico", , , , , hwnd)
    If lRet Then txtIconLoc(Index) = sFile
End Sub

Private Sub cmdExport_Click()
On Error GoTo 1
    Open AppResourcePath & "options.dat" For Output As #1
        For i = 0 To UBound(Colors) - 2
            Print #1, ColorNames(i) & " " & Colors(i)
        Next

        Print #1, ColorNames(15) & " " & IIf(ExtractPath(Colors(15)) = _
                  AppResourcePath, ExtractFilename(Colors(15)), Colors(15))
        Print #1, ColorNames(16) & " " & IIf(ExtractPath(Colors(16)) = _
                  AppResourcePath, ExtractFilename(Colors(16)), Colors(16))
    Close #1
    Exit Sub
1: Close #1
   MsgBox "Error exporting current settings to file.", vbExclamation, "Error while exporting"
End Sub
Private Sub cmdImport_Click()
    Dim iNum As Integer
    
    iNum = FreeFile
    Open AppResourcePath & "options.dat" For Input As #iNum
      'read file line per line
      While Not EOF(iNum)
        Line Input #iNum, sTemp
        
        'read first part, the name of the variable
        iSpace = InStr(1, sTemp, " ")
        sValue = Mid(sTemp, iSpace + 1)
        If i < 12 Then Colors(i) = CLng(sValue)
        
        'MenuFont
        If i = 12 Then Colors(i) = sValue
        
        'The quickicons
        If i > 14 Then
          ' no path, then path is resource path
          ' path doesn't exist then resource path
          If ExtractPath(sValue) = "" Or Dir(sValue) = "" Then
            sPath = AppResourcePath: sValue = ExtractFilename(sValue)
          End If
        End If
        'ComputerName...
        If i > 12 Then
            Colors(i) = sPath & sValue
            'if value from registry is crap, then reset value
            If i > 14 And Dir(Colors(i)) = "" Then
                Colors(i) = sPath & sValue
                SaveSetting ColorNames(i), Colors(i)
            End If
        End If
        
        sPath = ""
        i = i + 1
      Wend
    Close #iNum

End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    CenterForm Me
    MakeFormRounded Me, 30
    
    'set font of each control
    For Each Control In Controls
      Control.Font.Name = "Technical"
      If Control <> lbl(7) Then Control.Font.Size = 10
    Next
    
    'fill the available fonts in the combobox
    For i = 0 To Screen.FontCount - 1
        cmbFont.AddItem Screen.Fonts(i)
    Next
    'set index to currently used font
    cmbFont.Text = Colors(12)
    
    'draw the sidebar
    FillIcons
    'draw the example menu
    DrawDemoMenu False
    
    'read which quickicon is shown
    For i = 0 To 1
        chkQuickIcon(i).Value = GetLong("QuickIconVisible" & i, 1)
    Next
    
    lblInfo = "RepShell" & vbCrLf & vbCrLf & _
              "v. " & App.Major & "." & App.Minor & " build " & App.Revision & _
              vbCrLf & vbCrLf & _
              "RepShell is an attempt to a Rep(lacement)Shell for" & _
              " the standard Explorer." & vbCrLf & vbCrLf & _
              "Although it has relatively lot of bugs compared to commercial " & _
              "products, it is well on it's way to form some real competetion." & _
              vbCrLf & vbCrLf & _
              "If you're an experencied programmer, please take a look at the " & _
              "source code and let me know what you think of it. " & _
              "All improvements, suggestions and ideas are welcome." & vbCrLf & vbCrLf & _
              "Please mail your comments to."
              'i put another label control under this one with my email address
              
    'load icon locations
    For i = 0 To 1
      txtIconLoc(i) = GetSetting(ColorNames(15 + i), Colors(15 + i))
    Next
    
    'save current settings in temp variables
    bTempFillArrow = bFillArrow
    TempFont = Colors(12)
    'Add Colors to list
    For i = 0 To 11
        lstColor(IIf(i < 8, 0, 1)).AddItem ColorNames(i)
        TempColors(i) = Colors(i)
    Next
    lstColor(0).AddItem ColorNames(12)
    lstColor(0).ListIndex = 0: lstColor(1).ListIndex = 0
    
    'set checkboxes
    chkFillArrow.Value = Abs(bFillArrow)
    chkShowDesktopIcons.Value = Val(GetSetting("ShowDesktopIcons", "1"))
    chkShowHiddenFiles.Value = Val(GetSetting("ShowHiddenFiles", "0"))
    chkDefaultShell.Value = Abs(CInt((GetKeyVal("system.ini", "boot", "shell") = AppPath & App.EXEName & ".exe")))
    
    chkTranslucency.Value = Val(GetSetting("Translucency", "1"))
    sliTranslucency.Enabled = chkTranslucency.Value
    sliTranslucency.Value = GetLong("TranslucencyLevel", 100)

    If sClockFormat = "hh:mm" Then optTime(1).Value = True
       
End Sub




Private Sub picMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bDraw Then DrawDemoMenu True
End Sub
Private Sub picSettings_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bDraw Then DrawDemoMenu False
End Sub
Sub DrawDemoMenu(bActive As Boolean)
    Dim r As RECT
    SetRect r, 0, 0, picMenu.ScaleWidth, picMenu.ScaleHeight
    MakeMenuItems picMenu.hDC, "c:\command.com", r, bActive, True, , , picMenu
    bDraw = Not bDraw
End Sub
Private Sub sliTranslucency_Change()
    picTrans.Cls
    DrawIcon2 pictemp.hDC, "Config.ico", 2, 2
    AlphaBlending picTrans.hDC, 0, 0, 32, 32, pictemp.hDC, 0, 0, 32, 32, sliTranslucency.Value
End Sub


'Extracts Icon files and draws them on hdc
Sub DrawIcon2(hDC As Long, sFile$, X%, Y%)
  Dim hIco As Long
  hIco = ExtractIcon(0, AppResourcePath & sFile, 0)
  DrawIconEx hDC, X, Y, hIco, 24, 24, 0, 0, DI_NORMAL
  DestroyIcon hIco
End Sub
'draw the sidebar icons
Sub FillIcons()
    'position the icons
    For i = 0 To 4
        ImgIcon(i).Move 30, 10 + i * ((picSideBar.Height - 20) / 5)
        lblName(i).Move (92 - lblName(i).Width) / 2, ImgIcon(i).Top + 34
    Next
    'put the icon location in the tag
    ImgIcon(0).Tag = AppResourcePath & "Icons\Menu.ico"
    ImgIcon(1).Tag = AppResourcePath & "Icons\Colors.ico"
    ImgIcon(2).Tag = AppResourcePath & "Icons\Desktop.ico"
    ImgIcon(3).Tag = AppResourcePath & "Icons\Behaviour.ico"
    ImgIcon(4).Tag = AppResourcePath & "Icons\About.ico"
    'draw their respective icons
    For i = 0 To 4
        DrawIcon ImgIcon(i).Tag, ImgIcon(i)
    Next
    
    SelectIcon 0 'select first icon
End Sub
Private Sub lblName_Click(Index As Integer)
    Call ImgIcon_Click(Index)
End Sub
Private Sub ImgIcon_Click(Index As Integer)
    If Index = OldTab Then Exit Sub     'if pressed same icon, do nothing
    picSettings(Index).Visible = True   'show new tab
    picSettings(OldTab).Visible = False 'hide old tab
    SelectIcon Index                    'highlight icon
End Sub
Sub SelectIcon(nIndex As Integer)
    
    On Error Resume Next
    'restore previous icon
    DrawIcon ImgIcon(OldTab).Tag, ImgIcon(OldTab)
    lblName(OldTab).BackStyle = 0

    'draw new icon selected
    DrawIcon ImgIcon(nIndex).Tag, ImgIcon(nIndex), ILD_BLEND50
    lblName(nIndex).BackStyle = 1
    lblName(nIndex).BackColor = GetLong(ColorNames(10), Colors(10))

    OldTab = nIndex
    
End Sub
