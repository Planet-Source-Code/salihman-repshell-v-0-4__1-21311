Attribute VB_Name = "modGeneral"
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'TODO
'REMEBER THAT OU CAN USE WINDOWPLACEMENT TO SET  THE MIN POS OF WINDOWS
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------

Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
    ByVal fuWinIni As Long) As Long

    Public Const SPI_GETWORKAREA = 48
    Public Const SPI_SETWORKAREA = 47

'Undocumented shell function
Declare Function SHRunDialog Lib "shell32.dll" Alias "#61" _
    (ByVal hwndOwner As Long, ByVal hIcon As Long, _
    ByVal lpstrDirectory As String, ByVal szTitle As String, _
    ByVal szPrompt As String, ByVal uFlags As Browse) As Long

    Enum Browse
      SHRD_NOBROWSE = &H1
      SHRD_NOSTRING = &H2
    End Enum
    
'General
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (pDest As Any, pSource As Any, ByVal dwLength As Long)
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'timer functions
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'exit windows
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As ExitWindowsConst, ByVal dwReserved As Long) As Long
    Public Enum ExitWindowsConst
      EWX_LOGOFF = 0
      EWX_SHUTDOWN = 1
      EWX_REBOOT = 2
      EWX_FORCE = 4
      EWX_POWEROFF = 8
    End Enum
    
'MOUSE AND KEYB API'S
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long

Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    'listbox messages
    Public Const LB_ADDSTRING = &H180
    Public Const LB_FINDSTRINGEXACT = &H1A2
    Public Const LB_ERR = (-1)

'WINDOW FUNCTIONS
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As HWNDFlags, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As SWPFlags) As Long
    Enum HWNDFlags
        HWND_NOTOPMOST = -2
        HWND_TOPMOST = -1
        HWND_BOTTOM = 1
    End Enum
    Enum SWPFlags
        SWP_FRAMECHANGED = &H20
        SWP_DRAWFRAME = SWP_FRAMECHANGED
        SWP_HIDEWINDOW = &H80
        SWP_NOACTIVATE = &H10
        SWP_NOCOPYBITS = &H100
        SWP_NOMOVE = &H2
        SWP_NOOWNERZORDER = &H200
        SWP_NOREDRAW = &H8
        SWP_NOREPOSITION = SWP_NOOWNERZORDER
        SWP_NOSIZE = &H1
        SWP_NOZORDER = &H4
        SWP_SHOWWINDOW = &H40
    End Enum

Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Boolean
Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Boolean
Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Public Const SW_HIDE = 0
    Public Const SW_NORMAL = 1
    Public Const SW_SHOWMINIMIZED = 2
    Public Const SW_SHOWMAXIMIZED = 3
    Public Const SW_SHOWNOACTIVATE = 4
    Public Const SW_SHOW = 5
    Public Const SW_MINIMIZE = 6
    Public Const SW_SHOWMINNOACTIVE = 7
    Public Const SW_SHOWNA = 8
    Public Const SW_RESTORE = 9
    Public Const SW_SHOWDEFAULT = 10

'WINDOW API's
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'file operation
Declare Function SHFileOperation Lib "shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
    
    Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As FO_Flags
        pFrom As String
        pTo As String
        fFlags As Integer
        fAborted As Boolean
        hNameMaps As Long
        sProgress As String
    End Type
    
    Public Enum FO_Flags
        FO_DELETE = &H3
        FOF_ALLOWUNDO = &H40
        FO_RENAME = &H4
        FO_COPY = &H2
    End Enum
    Public Const FOF_SILENT = &H4

'HOTKEY API's
Declare Function RegisterHotkey Lib "user32" Alias "RegisterHotKey" (ByVal hwnd As Long, ByVal ID As Long, ByVal fsModifiers As MODKeys, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long) As Long
Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
    Public Enum MODKeys
        MOD_ALT = &H1&
        MOD_CONTROL = &H2&
        MOD_SHIFT = &H4&
        MOD_WIN = &H8&
    End Enum


'Set Foreground Window
Public Sub SetFGWindow(ByVal hwnd As Long, Show As Boolean)
  If Show Then
    If IsIconic(hwnd) Then
        ShowWindow hwnd, SW_RESTORE
    Else
        BringWindowToTop hwnd
    End If
  Else
    ShowWindow hwnd, SW_MINIMIZE
  End If
End Sub

'LIST ALL WINDOWS, Return the number of tasks
Public Function fEnumWindows(lst As ListBox) As Long
    With lst
      .Clear
      frmTask.lstNames.Clear
      Call EnumWindows(AddressOf fEnumWindowsCallBack, .hwnd)
      fEnumWindows = .ListCount
    End With
End Function

'FILTER WINDOWS, CALLBACK FUNCTION
Private Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lParam As Long) As Long
    
    Dim lExStyle As Long, bHasNoOwner As Boolean, sAdd As String, sCaption As String

    ' THE FILTER
    '  (* Check to see that it isnt this App) No longer used
    '  * Is it visible
    '  * has no owner and isn't Tool window OR
    '  * has an owner and is App window
    
    If IsWindowVisible(hwnd) Then
        bHasNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)
        lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
        
        If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bHasNoOwner) Or _
            ((lExStyle And WS_EX_APPWINDOW) And Not bHasNoOwner) Then
            sAdd = hwnd: sCaption = GetCaption(hwnd)
            Call SendMessage(lParam, LB_ADDSTRING, 0, ByVal sAdd)
            Call SendMessage(frmTask.lstNames.hwnd, LB_ADDSTRING, 0, ByVal sCaption)
        End If
    End If

    fEnumWindowsCallBack = True
End Function


'ADD AND DELETE HOTKEYS
Public Sub RegisterHotkeys()
    SetHotkey "KeyUnload", iUnload, MOD_CONTROL + MOD_ALT, vbKeyA
    SetHotkey "KeyStart", iStart, MOD_WIN, vbKeyS
    SetHotkey "KeyFavorites", iFavorites, MOD_WIN, vbKeyF
    SetHotkey "KeyRun", iRun, MOD_WIN, vbKeyR
End Sub
Public Sub UnregisterHotKeys()
    DeleteHotkey iUnload
    DeleteHotkey iStart
    DeleteHotkey iFavorites
    DeleteHotkey iRun
End Sub

Sub SetHotkey(ByVal sAtomName$, ByRef iAtom, fModifier As MODKeys, Key As Long)
    iAtom = GlobalAddAtom(sAtomName)
    If (iAtom <> 0) Then
       lR = RegisterHotkey(frmMain.hwnd, iAtom, fModifier, Key)
       If (lR = 0) Then GlobalDeleteAtom iAtom
    End If
End Sub
Sub DeleteHotkey(iAtom)
    UnregisterHotKey frmMain.hwnd, iAtom
    GlobalDeleteAtom iAtom
End Sub


Public Sub CenterForm(Frm As Form)
    Frm.Move (Screen.Width - Frm.Width) * 0.5, (Screen.Height - Frm.Height) * 0.5
End Sub

Public Function ShowRunDialog()
    Dim sTitle As String, sPrompt As String, hIco As Long
    On Error Resume Next
    sTitle = "RepShell Run Dialog"
    sPrompt = "Enter the name of a program, a directory, a document " _
              & "or an Internet resource and RepShell will open it for" _
              & "you." & vbCrLf & "(with a 'little bit' help from Windows)"
    hIco = ExtractIcon(0, AppResourcePath & "prog2.ico", 0)
    
    ShowRunDialog = SHRunDialog(0&, hIco, "c:\", sTitle, sPrompt, 0&)
    
    DestroyIcon hIco
End Function


Public Function GetActiveWindow() As Long
   Dim i As Long, j As Long
   i = GetForegroundWindow
   Do While i
     j = i   'store temp var if getparent returns 0
     i = GetParent(i)
   Loop
   GetActiveWindow = j
End Function

Public Function IsBounded(vntArray As Variant) As Boolean
    On Error Resume Next
    IsBounded = IsNumeric(UBound(vntArray))
End Function

'THIS FUNCTION IS USED FOR THE CUSTOM ADDITION TO THE SYSTEM MENU OF EACH
'PROGRAM, IT CHECKS THE STYLE TO SEE IF IT IS TOPMOST
Public Function IsWindowTopMost(hwnd As Long) As Boolean
    Dim lExStyle As Long
    lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (lExStyle And WS_EX_TOPMOST) = WS_EX_TOPMOST Then IsWindowTopMost = True
End Function
Public Function MakeTopMost(hwnd As Long, bTop As Long) As Long
    MakeTopMost = SetWindowPos(hwnd, IIf(bTop, HWND_TOPMOST, HWND_NOTOPMOST), _
    0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
End Function

'HIDE AND SHOW WINDOWS DESKTOP
Sub DesktopIcons(ByVal HideIcons As Boolean)
    Dim hDesktop As Long, hTaskBar As Long
    
    hDesktop = FindWindow("Progman", vbNullString)
    hTaskBar = FindWindow("Shell_TrayWnd", vbNullString)
    If HideIcons Then
        ShowWindow hDesktop, SW_HIDE 'hide desktopicons
        ShowWindow hTaskBar, SW_HIDE 'hide taskbar
    Else
        'show desktopicons and taskbar
        ShowWindow hDesktop, SW_SHOW
        ShowWindow hTaskBar, SW_SHOW
    End If

End Sub
