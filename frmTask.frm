VERSION 5.00
Object = "{713F0067-08FD-11D5-A06F-DFD761FF1C08}#1.0#0"; "REPCONTROLS.OCX"
Begin VB.Form frmTask 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RepControls.ctlTaskButton Task 
      Height          =   300
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextHoverColor  =   32768
      BackHoverColor  =   14737632
      Caption         =   ""
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   225
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lstApp 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblMinimize 
      BackColor       =   &H0000FF00&
      Height          =   1650
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   105
   End
End
Attribute VB_Name = "frmTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldIndex As Integer         'the previous selected button
Dim bRightClickMenu As Boolean  'flag, are we showing right click menu
Dim bMinimized As Boolean       'toggle, minimized/maximized

Private Sub Form_Load()
    'initial settings
    Task(0).Font.Name = "Technical"
    'get taskbox size from registry
    bMinimized = GetLong("TaskBoxSize", 0, General)
    'set color accordingly
    lblMinimize.BackColor = IIf(bMinimized, vbRed, vbGreen)
    
    'startup position
    Dim mLeft, mTop As Long
    mLeft = GetLong("TaskBoxLeft", Screen.Width - Width - 100, General)
    If mLeft > Screen.Width - Width - 50 Then mLeft = Screen.Width - Width - 50
    mTop = GetLong("TaskBoxTop", 50, General)
    If mTop > Screen.Height - Height - 50 Then mTop = Screen.Height - Height - 50
    Move mLeft, mTop
    
    'set appearance
    Init
    
    SetTimer hwnd, 1, 100, 0&       'applister
    AppListing
    'after primary init make visible
    Visible = True
End Sub

'this function is called from the Form_Load and from the Init function of
'the main form
Sub Init()
    'if off limits then repos
    If Left + Width > Screen.Width Then Left = Screen.Width - Width - 50
    If Top + Height > Screen.Height Then Top = Screen.Height - Height - 50
    If Left < 0 Then Left = 10
    If Top < 0 Then Top = 10
    
    MakeTopMost hwnd, True
    
    BackColor = GetLong(ColorNames(9), Colors(9))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'save startup position
    Call SaveLong("TaskBoxLeft", Left, General)
    Call SaveLong("TaskBoxTop", Top, General)
    Call SaveLong("TaskBoxSize", bMinimized, General)
    'clear region
    SetWindowRgn hwnd, 0&, True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmTask = Nothing
End Sub
'if click form, then deselect items
Private Sub Form_Click()
    SelectButton -1
End Sub
'--------------------
'MOVE BORDERLESS FORM
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
       ReleaseCapture
       SendMessage hwnd, WM_SYSCOMMAND, WM_MOVE, 0
    End If
End Sub
Private Sub lblMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(Button, Shift, X, Y)
End Sub
'--------------------

'toggle minimized/maximized form
Private Sub lblMinimize_DblClick()
    Dim iCounter As Integer
    'toggle variable
    bMinimized = Not bMinimized
    'set color accordingly
    lblMinimize.BackColor = IIf(bMinimized, vbRed, vbGreen)
    
    If bMinimized Then
        'make all buttons small and reposition them
        For iCounter = 1 To Task.UBound
            Task(iCounter).Width = 20
            Task(iCounter).Move 12 + (iCounter - 1) * 21, 5
        Next
        SetTaskBoxForm
    Else
        SetTaskBoxForm
        'make buttons big and repos
        For iCounter = 1 To Task.UBound
            Task(iCounter).Width = 153
            Task(iCounter).Move 12, (iCounter - 1) * 21 + 5
        Next
    End If
    Refresh
End Sub

'when button is leftclicked
Private Sub Task_LeftMouseDown(Index As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bCondition As Boolean
    
    'set condition, if wasn't activebutton then show
    '               if was active button then minimize
    bCondition = Not (Task(Index).Value)
    'make the button look selected and deselect old
    SelectButton IIf(bCondition, Index, -1)
    SetFGWindow Task(Index).Task, bCondition
End Sub
Private Sub Task_RightMouseDown(Index As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    Dim hWindow As Long         'the handle of the window from taskbutton
    Dim RetVal As Long          'general returnvalue
    Dim cPos As POINTAPI        'pos where to show menu (mousepos)
    Dim hMenu As Long           'the handle of the menu
    Dim bTop As Boolean         'flag, is window topmost
    
    'which window are we dealing with, in terms of hwnd
    hWindow = Task(Index).Task
    
    'set flag
    'if we have right-click menu, stop applisting sub
    bRightClickMenu = True
    SelectButton Index
    'see if the topmost flag of that window is set
    bTop = IsWindowTopMost(hWindow)
    
    'Get handle of the menu
    hMenu = GetSystemMenu(hWindow, False)
    'Append a seporator line
    RetVal = AppendMenu(hMenu, MFT_SEPARATOR, 500, "")
    'Append Always on Top Item, and set it's state
    RetVal = AppendMenu(hMenu, IIf(bTop, MFS_CHECKED, MFS_UNCHECKED), 501, "Always On Top")

    'Show system menu at current position
    GetCursorPos cPos
    'show menu and set it to return idCmd
    RetVal = TrackPopupMenu(GetSystemMenu(hWindow, False), _
        TPM_TOPALIGN Or TPM_LEFTALIGN Or TPM_RETURNCMD Or _
        TPM_NONOTIFY Or TPM_RIGHTBUTTON, cPos.X, cPos.Y, 0, hwnd, ByVal 0&)
    
    '501 is our own menuitem
    If RetVal = 501 Then
        'make window topmost according to flag
        MakeTopMost hWindow, Not bTop
    Else
        'if an other item is clicked, send to window to handle it
        RetVal = PostMessage(hWindow, WM_SYSCOMMAND, RetVal, ByVal 0&)
    End If
    'delete items, because outside RepShell they will not respond
    RetVal = DeleteMenu(hMenu, 500, MF_BYCOMMAND)
    RetVal = DeleteMenu(hMenu, 501, MF_BYCOMMAND)
    
    'set flag back
    bRightClickMenu = False

End Sub


Sub AppListing()
    On Error Resume Next
    
    Dim nTasks As Long, i As Long
    Dim iFind As Integer, iTaskPrevUbound As Integer
    Dim sHwnd As String, bTemp As Boolean
    
    nTasks = fEnumWindows(lstApp)
    
    '  check if taskitems are still present, unload if neccesary
    '
    '  this is done to keep the order, if I should unload all
    '  of them and then reload, their positions would change
    '  according to screen zorder. This way their order is held
    
    For i = 1 To Task.UBound
      'if unloaded then no need to go on
      If i > Task.UBound Then Exit For
      ' If Hwnd is still in the list
      sHwnd = Format(Task(i).Task)
      iFind = SendMessage(lstApp.hwnd, LB_FINDSTRINGEXACT, -1, ByVal sHwnd)
      'if not in list, unload item
      If iFind = LB_ERR Then
        UnloadItem i
      Else
        Task(i).Caption = lstNames.List(iFind)
      End If
    Next
    
    'fill remaining tasks
    For i = 0 To nTasks - 1
      If FindhWndInTask(CLng(lstApp.List(i))) = -1 Then
        iTaskPrevUbound = Task.UBound
        Load Task(Task.UBound + 1)
        With Task(Task.UBound)
          .Task = lstApp.List(i)
          .Caption = lstNames.List(i)
          If bMinimized Then .Width = 20
          'Put first button at top
          If Task.UBound = 1 Then
            'this is left/top of task(0) which isn't visible
            .Move 12, 5
          Else
            If bMinimized Then
                .Move Task(iTaskPrevUbound).Left + 20 + 1, 5
            Else
                '20 = Task(iTaskPrevUbound).Height
                .Move 12, Task(iTaskPrevUbound).Top + 20 + 1
            End If
          End If
          SetTaskBoxForm
          .Visible = True
          bTemp = True
        End With
      End If
    Next
    
    'if not already set form (added new item(s)), but removed then set form
    If Not bTemp Then SetTaskBoxForm
    
    'if button is not cur active window then set it to be so, but check flag first
    If Not bRightClickMenu Then
        If GetActiveWindow <> Task(OldIndex).Task Then
            SelectButton FindhWndInTask(GetActiveWindow)
        End If
    End If
End Sub
'Function that set the dimension of the form and makes it rounded
Function SetTaskBoxForm()
    Dim hRgn As Long, lUbound As Long
    
    lUbound = Task.UBound
    If bMinimized Then
        'if less than 3 buttons then set standard width
        ' 17 = 12 on Left side of button and 5 on right side
        Width = IIf(lUbound < 3, 80, lUbound * 21 + 17) * Screen.TwipsPerPixelY
        Height = 450
    Else
        'if no buttons then set one button height
        ' 10 = 5 on top , 5 on bottom
        Height = IIf(lUbound = 0, 30, lUbound * 21 + 9) * Screen.TwipsPerPixelY
        Width = 2550
    End If
    lblMinimize.Height = ScaleHeight
    
    'Set the special form of this window, decl. on form level
    hRgn = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, 15, 15)
    'combine the two regions and set the window to that region
    SetWindowRgn hwnd, hRgn, True
    'delete the regions to free resources
    DeleteObject hRgn
End Function

'Look into all task button until you find one with the correct hwnd
'and then return its index
Function FindhWndInTask(ByVal sHwnd As Long) As Integer
    Dim i As Integer
    For i = 1 To Task.UBound
      If Task(i).Task = sHwnd Then
         FindhWndInTask = i
         Exit Function
      End If
    Next
    FindhWndInTask = -1
End Function

' Unload Item
' Passes the item info on to the previous and unloads the last taskbutton
Function UnloadItem(ByVal i As Integer)
    Dim j As Integer
    For j = i To Task.UBound - 1
      Task(j).Task = Task(j + 1).Task
    Next
    If OldIndex = Task.UBound Then SelectButton -1
    Unload Task(Task.UBound)
End Function
'Function to make clicked button selected and deselect an other one
Function SelectButton(nIndex As Integer)
    On Error Resume Next
    'restore previous icon
    If OldIndex <> -1 Then Task(OldIndex).Value = False
    If nIndex <> -1 Then Task(nIndex).Value = True
    OldIndex = nIndex
End Function
