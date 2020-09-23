VERSION 5.00
Begin VB.Form frmStart 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1620
   ClientLeft      =   795
   ClientTop       =   1425
   ClientWidth     =   3450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   230
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUnload 
      Interval        =   250
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OldIndex As Long        'the index of the current item (0-based)
Private sPath As String         'the path which contents we are showing
Private fChild As frmStart      'a reference to the submenu
Private mParent As frmStart     'a reference to the parent menu

Private Items() As String       'an array with the items to be shown
Private lUbound As Integer      'the Ubound of the array, if no items -1
Private lNumFolders As Integer  'the number of folders in the menu
Private bArrowMoved As Boolean  'help var to tell us that we opened sub with arrow nav
Private iTemp As Integer        'help var when pressing arrow up

' arrow navigation, Enter to start, escape to unload
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
      Case vbKeyUp
        'go on item down, or if at end of list go one up
        iTemp = IIf(lUbound = -1, 0, lUbound)
        MoveSel IIf(OldIndex <= 0, iTemp, OldIndex - 1), False
      Case vbKeyDown
        'go one item up, or if at top of list go to bottom
        MoveSel IIf(OldIndex >= lUbound, 0, OldIndex + 1), False
      Case vbKeyLeft
        'go to parent menu and unload sub
        If Not (mParent Is Nothing) Then mParent.SetFocus: Unload Me
      Case vbKeyRight
        'if there's a sub
        If IIf(Tag <> "Drives", OldIndex <= lNumFolders, _
           IIf(OldIndex <> 3, True, False)) Then
            'open sub and select first item
            bArrowMoved = True
            Timer1.Interval = 1
        End If
    
      'execute current item
      Case vbKeyReturn: Call Form_MouseDown(vbLeftButton, 0, 1, 1)
      'unload entire menu
      Case vbKeyEscape: UnloadAll
    End Select
End Sub

'WITH THESE 2 FUNCTIONS THE MENUSYSTEM UNLOADS ITSELF
'WHEN IT LOSES FOCUS
Private Sub Form_LostFocus()
    'check if current window, is a startmenu window
    'if it isn't so then unload entire menusystem
    For Each Item In frmStartHwnd
        If Format(Item) = Format(GetActiveWindow) Then Exit Sub
    Next
    UnloadAll
End Sub
Private Sub tmrUnload_Timer()
    Call Form_LostFocus
End Sub


Private Sub Form_Load()
    'don't highlight items
    OldIndex = -1: ScaleMode = vbPixels
    'add this menus hwnd to the collection
    'so we can check in the other function for the hwnds
    frmStartHwnd.Add hwnd, Format(hwnd)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'free mem
    SetWindowRgn hwnd, 0&, True
    'remove it from the collection
    frmStartHwnd.Remove Format(hwnd)
    Set CurActiveMenu = Nothing
    Set frmStart = Nothing
End Sub
'just unload submenus
Public Sub UnloadChildren()
    If Not (fChild Is Nothing) Then fChild.UnloadChildren
    Unload Me
End Sub
'unload entire menu
Public Sub UnloadAll()
    If Not (fChild Is Nothing) Then fChild.UnloadChildren
    If Not (mParent Is Nothing) Then mParent.UnloadAll
    MenuDirection = False
    Unload Me
End Sub
'this is used to hide the menu when executing a command
'because sometimes it takes longer to process a command
'and then the menu would stay visible, if we would unload
'first the command is not executed
Public Sub HideAll()
  If Not (mParent Is Nothing) Then mParent.Hide
  Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static OldX As Long, OldY As Long
  'check if mouse moved
  If (X <> OldX Or Y <> OldY) Then MoveSel Int(Y * 0.05), True
  OldX = X: OldY = Y
End Sub


Public Sub GetMenu(Path As String, Optional Parent As frmStart = Nothing, _
                   Optional Drv As Boolean, Optional DoFiles As Boolean, _
                   Optional sTag As String)
  
  Dim i As Integer, sTemp As String, Maxlen As Long, r As RECT
  
  On Error Resume Next
 
  'setup start
  r.Top = -20
  Set CurActiveMenu = Me 'so subclassing can unload menu
      
  If Drv Then
    'the first menu
    Tag = "Drives"
    'set font again to check widths  just in case got lost somewhere
    frmMain.Font.Name = Colors(12)
    
    'put items in array
    Items = DrivesPresent(True)
    ReDim Preserve Items(UBound(Items) + 4)
    lUbound = UBound(Items)
    
    For i = lUbound To 4 Step -1
        Items(i) = Items(i - 4)
    Next
    Items(0) = "Start Menu": Items(1) = "Favorites"
    Items(2) = "Documents": Items(3) = "Run"
    
    'check width
    For i = 0 To lUbound
        CheckMaxLen Items(i), Maxlen
    Next
    'set dim and pos
    Width = (Maxlen + 40) * Screen.TwipsPerPixelX
    Height = (lUbound + 1) * 300
    Left = 255: Top = 810  'under My computer Icon

    'Draw Menu
    For i = 0 To lUbound
        SetRect r, 0, r.Top + 20, ScaleWidth, r.Bottom + 20
        MakeMenuItems hDC, Items(i), r, False, IIf(i = 3, False, True), , , Me
    Next
    
  Else
  
    'set reference to parent
    Set mParent = Parent
    sPath = ProperPath(Path)
    Tag = sTag
    
    'get files and folders in path
    Items = GetFilesFolders(sPath, DoFiles, lUbound, lNumFolders)
    
    If lUbound = -1 Then
      'empty
      Width = (GetTextWidth(frmMain.hDC, "[No SubFolders]") + 30) * Screen.TwipsPerPixelX
      Height = 300
      SetRect r, 0, 0, ScaleWidth, 20
      MakeMenuItems hDC, "[No SubFolders]", r, False, False, , , Me
      
    Else
      'check widest item
      For i = 0 To lUbound
        CheckMaxLen Items(i), Maxlen
      Next

      'because, if there are folders an arrow has to be drawn and the form
      'must be wider, otherwise just icon
      Width = (Maxlen + IIf(lNumFolders > -1, 40, 20)) * Screen.TwipsPerPixelX
      Height = (lUbound + 1) * 300
      
      'drawfolders
      For i = 0 To lUbound
        SetRect r, 0, r.Top + 20, ScaleWidth, r.Bottom + 20
        MakeMenuItems hDC, sPath & Items(i), r, False, i <= lNumFolders, , , Me
      Next
      
    End If
    
  End If
  'jsut what the functions says
  MakeFormRounded Me, 20
  
  'if menu larger then screen or beyond screen it repositions it
  If mParent Is Nothing Then
    mLeft = IIf(Left + Width > Screen.Width, Left - Width, Left)
  Else
    
    With mParent
      If (.Left + .Width + Width > Screen.Width And Not MenuDirection) Or _
         (.Left - Width < 0 And MenuDirection) Then
         MenuDirection = Not MenuDirection
      End If
    End With
    mLeft = mParent.Left + IIf(MenuDirection, -Width, mParent.Width)
  
  End If
  mTop = IIf(Top + Height > Screen.Height, IIf(Screen.Height - Height < 0, 0, Screen.Height - Height), Top)
  Move mLeft, mTop

  Visible = True: MakeTopMost hwnd, True
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  
    If Button = vbLeftButton Or Button = vbRightButton Then
      
      Dim bExplore As Boolean
      
      Playsound "success"
      HideAll 'hide menusystem, otherwise if it takes long to start program
              'the menu would still be visible
              
      'if right clicked explore the folder
      bExplore = (Button = vbRightButton)
      
      If Tag <> "Drives" Then
          
          'If OldIndex <= lNumFolders Then
            ExecuteFile sPath & Items(OldIndex) ', , , bExplore
          'Else
          '  ExecuteFile sPath & Items(OldIndex)
          'End If
          
      Else
      
          Select Case OldIndex
            Case 0: ExecuteFile GetSpecialfolder(CSIDL_PROGRAMS) ', , , bExplore
            Case 1: ExecuteFile GetSpecialfolder(CSIDL_FAVORITES) ', , , bExplore
            Case 2: ExecuteFile GetSpecialfolder(CSIDL_RECENT) ', , , bExplore
            Case 3: ShowRunDialog
            Case Else: ExecuteFile Items(OldIndex) ', , , bExplore
          End Select
          
      End If
      UnloadAll 'unload menu system
      
    End If
  
End Sub

'Timer to show menu
Private Sub Timer1_Timer()
    On Error Resume Next
    'disable timer
    Timer1.Interval = 0
    'check, maybe we moved to another non-folder item during delay interval
    If Tag <> "Drives" Then
        If OldIndex > lNumFolders Then Exit Sub
    End If
        
    Set fChild = New frmStart
    With fChild
        'move sub
        .Move Left + Width, Top + (OldIndex * 20) * Screen.TwipsPerPixelY
        'we're on the first menu
        If Tag = "Drives" Then
            'load special folders in new sub
            Select Case OldIndex
            Case 0: .GetMenu GetSpecialfolder(CSIDL_PROGRAMS), Me, , True, "Start Menu"
            Case 1: .GetMenu GetSpecialfolder(CSIDL_FAVORITES), Me, , True, "Favorites"
            Case 2: .GetMenu GetSpecialfolder(CSIDL_RECENT), Me, , True, "Documents"
            
            ' TO DO : Option To DELETE RECENT FILES
            
            Case Is <> 3: .GetMenu Items(OldIndex), Me, , False
            End Select
        Else
            If Tag = "Start Menu" Or Tag = "Favorites" Then
                'if a special folder, set flag that the files
                'have to be shown to
                .GetMenu sPath & Items(OldIndex), Me, , True, Tag
            Else
                'the files don't have to be shown in other folders
                'but you can specify that they have to be, so the
                'GetFilesAndFolders returns files also regardless of
                'the DoFiles param
                .GetMenu sPath & Items(OldIndex), Me, , False
            End If
        End If
        'if the sub is opened with right arrow then move to first item
        If bArrowMoved Then .MoveSel 0, False: bArrowMoved = False
    End With
    Playsound "open"

End Sub

Sub MoveSel(Index As Integer, ShowSub As Boolean)
  
  Static sOldName As String, bDrawArrow As Boolean
  Dim sNewName As String, bDrawNewArrow As Boolean, r As RECT
  
  On Error Resume Next
  
  If Index <> OldIndex Then                'are we on another item
    If OldIndex = -1 Then OldIndex = 0     'otherwise error
    
    If Tag <> "Drives" Then
        If lUbound = -1 Then
            sNewName = "[No SubFolders]"   'empty foldout
        Else
            sNewName = sPath & Items(Index)
            bDrawNewArrow = Index <= lNumFolders
        End If
    Else 'drives menu
        If Index <> 3 Then bDrawNewArrow = True 'dont draw arrow if 'Run'
        sNewName = Items(Index)
    End If
    
    If Not (fChild Is Nothing) Then fChild.UnloadChildren
    
    'RESET OLD
    If sOldName <> "" Then
        SetRect r, 0, OldIndex * 20, ScaleWidth, (OldIndex + 1) * 20
        MakeMenuItems hDC, sOldName, r, False, bDrawArrow, , , Me
    End If
    'SELECT NEW
    SetRect r, 0, Index * 20, ScaleWidth, (Index + 1) * 20
    MakeMenuItems hDC, sNewName, r, True, bDrawNewArrow, , , Me

    '(if on first menu but NOT on Run (3rd) item OR
    'on all other menus and a folder item) AND
    'ShowSub parameter is True
    If IIf(Tag <> "Drives", Index <= lNumFolders, IIf(Index <> 3, True, False)) _
        And ShowSub Then
        Timer1.Interval = 200
    Else
        Playsound "hover"
    End If

    OldIndex = Index
    sOldName = sNewName
    bDrawArrow = bDrawNewArrow
  End If

End Sub

Sub CheckMaxLen(ByVal strItem As String, Maxlen As Long)
    Dim lTextW As Long
    
    lTextW = GetTextWidth(frmMain.hDC, strItem)
    If lTextW > Maxlen Then
        
        If lTextW > Screen.Width / Screen.TwipsPerPixelX / 3 Then
          strItem = Left(strItem, 45) & "..."
          lTextW = GetTextWidth(frmMain.hDC, strItem)
        End If
        Maxlen = lTextW

    End If
    
End Sub
