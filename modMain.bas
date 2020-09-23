Attribute VB_Name = "modMain"
DefInt I

'Colors
Public ColorNames(16) As String
Public Colors(16) As Variant

' 0 = BackColor, 1 = ForeColor,2 = Active BackColor,3 = Active ForeColor
' 4 = Active ArrowColor,5 = InActive ArrowColor,6 = Active ArrowFillColor
' 7 = InActive ArrowFillColor, 8 = DesktopBackColor, 9 = TaskBoxBackColor
' 10 = LabelBackColor, 11 = LabelForeColor, 12 = MenuFont
' 13 = ComputerName, 14 = RecycleBinName, 15 = QuickSound
' 16 = QuickNetConnect

Public bFillArrow As Boolean            'fill arrows in menus
Public sClockFormat As String           'clock style
Public AppPath As String                'apllication path (to reduce typing)
Public AppResourcePath As String        'app res path      """"""""""""""""
Public CurActiveMenu As frmStart        'still used in modSubclassing to unload menu

Public Rows As Integer                  'how many rows are there on desktop
Public IconsPerColumn As Integer        'how many icons fit in one column
Public MenuDirection As Boolean         'which dir is the menu currently expanding to
                                        'false=right, true=left

'each hotkey has to have its own public var
Public iUnload%, iStart%, iFavorites%, iRun%
'collection that holds the hwnds of the loaded menus
Public frmStartHwnd As New Collection
Public Tooltip As New pToolTip.CToolTip
'the handle of the systray
Public lSystrayHwnd As Long
Sub Main()
    On Error Resume Next
    
    Dim sPath As String
    Dim sTemp As String, iNum As Integer, sValue As String
    Dim i As Integer, iSpace As Integer
    
    'Set Paths
    AppPath = ProperPath(App.Path): AppResourcePath = AppPath & "Resource\"

    ' Load Colors
    iNum = FreeFile
    Open AppResourcePath & "options.dat" For Input As #iNum
      'read file line per line
      While Not EOF(iNum)
        Line Input #iNum, sTemp
        
        'read first part, the name of the variable
        iSpace = InStr(1, sTemp, " ")
        ColorNames(i) = Left(sTemp, iSpace - 1)
        'second part, value
        sValue = Mid(sTemp, iSpace + 1)
        If i < 12 Then Colors(i) = GetLong(ColorNames(i), CLng(sValue))
        
        'MenuFont
        If i = 12 Then Colors(i) = GetSetting(ColorNames(i), sValue)
        
        'The quickicons
        If i > 14 Then
          ' no path, then path is resource path
          'path doesn't exist then resource path
          If ExtractPath(sValue) = "" Or Dir(sValue) = "" Then
            sPath = AppResourcePath: sValue = ExtractFilename(sValue)
          End If
        End If
        'ComputerName...
        If i > 12 Then
            Colors(i) = GetSetting(ColorNames(i), sPath & sValue)
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
    'read some settings from the registry
    bFillArrow = CBool(GetSetting("FillArrow", "1"))
    sClockFormat = IIf(GetSetting("ClockFormat", "24") = "12", "h:mm AMPM", "hh:mm")
    
    'add the fonts as a resource, so we can use them in the app
    AddFontResource AppResourcePath & "Fonts\" & "Presdntn.ttf"
    AddFontResource AppResourcePath & "Fonts\" & "Techncln.ttf"
    
    'Hide the desktop icons and the taskbar
    DesktopIcons True
    
    'set workarea to fullscreen
    'SystemParametersInfo SPI_SETWORKAREA, 0&, 0&, 0

    'start the systray and retrieve it's handle
    ExecuteFile AppPath & "Systray.exe"
    
    Load frmMain: Load frmTask: Load frmStart
    RegisterHotkeys
    
    'if RepShell is default shell then run the startup programs
    'in the startup folder and in the registry
    If (GetKeyVal("system.ini", "boot", "shell") = _
        AppPath & App.EXEName & ".exe") Then

        RunStartUpPrograms

    End If
End Sub

Public Sub ExitApp()
  SHNotify_Unregister
  UnregisterHotKeys
  
  SendMessage lSystrayHwnd, WM_UNLOAD, 0, 0
      
  KillTimer frmTask.hwnd, 1

  UnHook
  
  DesktopIcons False 'show desktopicons
  
  RemoveFontResource AppResourcePath & "Presdntn.ttf"
  RemoveFontResource AppResourcePath & "Techncln.ttf"
    
  For Each Form In Forms
    Unload Form
  Next
End Sub
