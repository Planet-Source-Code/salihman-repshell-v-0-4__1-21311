Attribute VB_Name = "modMenuDefs"
Option Explicit

Public Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
End Type

Public Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As OD_ACTIONS
    itemState As OD_STATE
    hwndItem As Long
    hDC As Long
    rcItem As RECT
    itemData As Long
End Type

' Owner draw actions
Public Enum OD_ACTIONS
    ODA_DRAWENTIRE = &H1
    ODA_SELECT = &H2
    ODA_FOCUS = &H4
End Enum

' Owner draw state
Public Enum OD_STATE
    ODS_SELECTED = &H1
    ODS_GRAYED = &H2
    ODS_DISABLED = &H4
    ODS_CHECKED = &H8
    ODS_FOCUS = &H10
    ODS_DEFAULT = &H20
    ODS_COMBOBOXEDIT = &H1000
End Enum

Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function TrackPopupMenu Lib "user32" _
                              (ByVal hMenu As Long, _
                              ByVal wFlags As TPM_wFlags, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal nReserved As Long, _
                              ByVal hwnd As Long, _
                              ByVal lprc As Any) As Long   ' lprc As RECT

Public Enum TPM_wFlags
  TPM_LEFTBUTTON = &H0
  TPM_RIGHTBUTTON = &H2
  TPM_LEFTALIGN = &H0
  TPM_CENTERALIGN = &H4
  TPM_RIGHTALIGN = &H8
  TPM_TOPALIGN = &H0
  TPM_VCENTERALIGN = &H10
  TPM_BOTTOMALIGN = &H20

  TPM_HORIZONTAL = &H0         ' Horz alignment matters more
  TPM_VERTICAL = &H40          ' Vert alignment matters more
  TPM_NONOTIFY = &H80          ' Don't send any notification msgs
  TPM_RETURNCMD = &H100
End Enum

Public Type MENUITEMINFO
  cbSize As Long
  fMask As MII_Mask
  fType As MF_Type            ' MIIM_TYPE
  fState As MF_State          ' MIIM_STATE
  wID As Long                 ' MIIM_ID
  hSubMenu As Long            ' MIIM_SUBMENU
  hbmpChecked As Long         ' MIIM_CHECKMARKS
  hbmpUnchecked As Long       ' MIIM_CHECKMARKS
  dwItemData As Long          ' MIIM_DATA
  dwTypeData As String        ' MIIM_TYPE
  cch As Long                 ' MIIM_TYPE
End Type

Public Enum MII_Mask
  MIIM_STATE = &H1
  MIIM_ID = &H2
  MIIM_SUBMENU = &H4
  MIIM_CHECKMARKS = &H8
  MIIM_TYPE = &H10
  MIIM_DATA = &H20
End Enum

Public Enum MF_Type
  MFT_STRING = &H0
  MFT_BITMAP = &H4
  MFT_MENUBARBREAK = &H20
  MFT_MENUBREAK = &H40
  MFT_OWNERDRAW = &H100
  MFT_RADIOCHECK = &H200
  MFT_SEPARATOR = &H800
  MFT_RIGHTORDER = &H2000
  MFT_RIGHTJUSTIFY = &H4000
End Enum

Public Enum MF_State
  MFS_GRAYED = &H3
  MFS_DISABLED = &H1
  MFS_CHECKED = &H8
  MFS_HILITE = &H80
  MFS_ENABLED = &H0
  MFS_UNCHECKED = &H0
  MFS_UNHILITE = &H0
  MFS_DEFAULT = &H1000
End Enum

Public Const MF_BYCOMMAND = &H0
Public Const MF_BYPOSITION = &H400

Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" _
                              (ByVal hMenu As Long, _
                              ByVal uItem As Long, _
                              ByVal fByPosition As Boolean, _
                              lpmii As MENUITEMINFO) As Boolean

Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" _
                              (ByVal hMenu As Long, _
                              ByVal uItem As Long, _
                              ByVal fByPosition As Boolean, _
                              lpmii As MENUITEMINFO) As Boolean

Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, _
                              ByVal bRevert As Long) As Long

Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" _
    (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, _
    ByVal lpNewItem As Any) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, _
    ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, _
    ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

'This function is called from frmMain when right-clicking an item
'This function reserves memory and retrieves the pidls of the items
'Then it calls the menu below "ShowShellContextMenu" which actually
'shows the context menu
Public Sub ShellContextMenu(obj As Control, pt As POINTAPI, Optional RecycleBin As Boolean)
  
  Dim cItems As Integer            ' count of selected items
  Dim i As Integer                 ' counter
  Dim asPaths() As String          ' array of selected items' paths (zero based)
  Dim apidlFQs() As Long           ' array of selected items' fully qualified pidls (zero based)
  Dim isfParent As IShellFolder    ' selected items' parent shell folder
  Dim apidlRels() As Long          ' array of selected items' relative pidls (zero based)
  
  ' This only works for one file, if you want to select multiple files
  ' Put the Items in the array
  cItems = 1
  ReDim asPaths(0)
  asPaths(0) = obj.Tag
  
  ' ==================================================
  ' Finally, get the IShellFolder of the selected directory, load the relative
  ' pidl(s) of the selected items into the array, and show the menu.
  
  If Len(asPaths(0)) Then
    
    ' Get a copy of each selected item's fully qualified pidl from it's path.
    For i = 0 To cItems - 1
      ReDim Preserve apidlFQs(i)
      apidlFQs(i) = GetPIDLFromPath(frmMain.hwnd, asPaths(i))
    Next
    
    If RecycleBin Then
        Dim pidl As Long
        Call SHGetSpecialFolderLocation(0&, CSIDL_BITBUCKET, pidl)
        apidlFQs(0) = pidl
    End If
    
    If apidlFQs(0) Then
    
      ' Get the selected item's parent IShellFolder.
      Set isfParent = GetParentIShellFolder(apidlFQs(0))
      If (isfParent Is Nothing) = False Then
        
        ' Get a copy of each selected item's relative pidl (the last item ID)
        ' from each respective item's fully qualified pidl.
        For i = 0 To cItems - 1
          ReDim Preserve apidlRels(i)
          apidlRels(i) = GetItemID(apidlFQs(i), GIID_LAST)
        Next
        
        If apidlRels(0) Then
          ' Show the shell context menu for the selected items.
          Call ShowShellContextMenu(frmMain.hwnd, isfParent, cItems, apidlRels(0), pt)
        End If   ' apidlRels(0)
        
        ' Free each item's relative pidl.
        For i = 0 To cItems - 1
          Call MemAllocator.Free(ByVal apidlRels(i))
        Next
        
      End If   ' (isfParent Is Nothing) = False

      ' Free each item's fully qualified pidl.
      For i = 0 To cItems - 1
        Call MemAllocator.Free(ByVal apidlFQs(i))
      Next
      
    End If   ' apidlFQs(0)
  End If   ' Len(asPaths(0))
  
End Sub

' Displays the specified items' shell context menu.
'
'    hwndOwner  - window handle that owns context menu and any err msgboxes
'    isfParent  - pointer to the items' parent shell folder
'    cPidls     - count of pidls at, and after, pidlRel
'    pidlRel    - the first item's pidl, relative to isfParent
'    pt         - location of the context menu, in screen coords
'    fPrompt    - flag specifying whether to prompt before executing any selected
'                 context menu command
'
' Returns True if a context menu command was selected, False otherwise.

Public Function ShowShellContextMenu(hwndOwner As Long, isfParent As IShellFolder, _
                                    cPidls As Integer, pidlRel As Long, _
                                    pt As POINTAPI) As Boolean
                                    
  Dim IID_IContextMenu As IShellFolderEx_TLB.Guid
  Dim IID_IContextMenu2 As IShellFolderEx_TLB.Guid
  Dim icm As IContextMenu
  Dim hr As Long   ' HRESULT
  Dim hMenu As Long
  Dim idCmd As Long
  Dim cmi As IShellFolderEx_TLB.CMINVOKECOMMANDINFO
  Dim mii As MENUITEMINFO
  Const idOurCmd = 100
  Const sOurCmd = "&Rename"
  
  ' Fill the IContextMenu interface ID, {000214E4-000-000-C000-000000046}
  With IID_IContextMenu
    .Data1 = &H214E4
    .Data4(0) = &HC0
    .Data4(7) = &H46
  End With
    
  ' Get a refernce to the item's IContextMenu interface.
  hr = isfParent.GetUIObjectOf(hwndOwner, cPidls, pidlRel, IID_IContextMenu, 0, icm)
  If hr >= NOERROR Then
    
    ' Fill the IContextMenu2 interface ID, {000214F4-000-000-C000-000000046}
    ' and get the folder's IContextMenu2. Is needed so the "Send To" and "Open
    ' With" submenus get filled from the HandleMenuMsg call in WndProc.
    With IID_IContextMenu2
      .Data1 = &H214F4
      .Data4(0) = &HC0
      .Data4(7) = &H46
    End With
    Call icm.QueryInterface(IID_IContextMenu2, ICtxMenu2)
    
    ' Create a new popup menu...
    hMenu = CreatePopupMenu()
    If hMenu Then

      ' Add the item's shell commands to the popup menu.
      If (ICtxMenu2 Is Nothing) = False Then
        hr = ICtxMenu2.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_EXPLORE)
      Else
        hr = icm.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_EXPLORE)
      End If
     
      ' Now add our own menu item
      With mii
       .cbSize = Len(mii)
       .fMask = MIIM_ID Or MIIM_TYPE
       .wID = idOurCmd
       .fType = MFT_STRING
       .dwTypeData = sOurCmd
       .cch = Len(sOurCmd)
      End With
      
      Call InsertMenuItem(hMenu, GetMenuItemCount(hMenu) - 2, MF_BYPOSITION, mii)
          
      If hr >= NOERROR Then
        ' Show the item's context menu
        idCmd = TrackPopupMenu(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or _
                                TPM_RIGHTBUTTON, pt.X, pt.Y, 0, hwndOwner, ByVal 0&)
        
        ' If a menu command is selected...
        If (idCmd = idOurCmd) Then
          frmMain.DesktopRenameShow
        ElseIf idCmd Then
          ' Fill the struct with the selected command's information.
          With cmi
            .cbSize = Len(cmi)
            .hwnd = hwndOwner
            .lpVerb = idCmd - 1 ' MAKEINTRESOURCE(idCmd-1);
            .nShow = SW_NORMAL
          End With

          ' Invoke the shell's context menu command. The call itself does
          ' not err if the pidlRel item is invalid, but depending on the selected
          ' command, Explorer *may* raise an err. We don't need the return
          ' val, which should always be NOERROR anyway...
          If Not (ICtxMenu2 Is Nothing) Then
            Call ICtxMenu2.InvokeCommand(cmi)
          Else
            Call icm.InvokeCommand(cmi)
          End If
        End If   ' idCmd
      
      End If   ' hr >= NOERROR (QueryContextMenu)

      Call DestroyMenu(hMenu)
    
    End If   ' hMenu
  End If   ' hr >= NOERROR (GetUIObjectOf)

  ' Release the folder's IContextMenu2 from the global variable.
  Set ICtxMenu2 = Nothing
  
  ' Return True if a menu command was selected
  ' (letting us know to react accordingly...)
  ShowShellContextMenu = CBool(idCmd)

End Function

' Returns the string of the specified menu command ID in the specified menu.
Public Function GetMenuCmdStr(hMenu As Long, ByVal idCmd As Long) As String
    
    '--------------------------------------
    'Doesn't work and I don't know WHY????
    '--------------------------------------
    
    Dim m As MENUITEMINFO
    m.dwTypeData = Space(64)
    m.cbSize = Len(m)
    m.cch = 64
    m.fMask = MIIM_DATA Or MIIM_TYPE
    m.fType = MFT_STRING
    If GetMenuItemInfo(hMenu, idCmd, MF_BYCOMMAND, m) Then _
        GetMenuCmdStr = Left(m.dwTypeData, InStr(1, m.dwTypeData, " ") - 1)
    If Len(GetMenuCmdStr) Then frmMain.Print GetMenuCmdStr 'just to see if it's working
    
End Function

'Does the item have a submenu
Public Function GetMenuArrow(hMenu As Long, ByVal idCmd As Integer) As Boolean
    Dim info As MENUITEMINFO
    'information to retreive with GetMenuItemInfo
    info.fMask = MIIM_SUBMENU
    info.cbSize = LenB(info) 'size in byte of structure
    
    Call GetMenuItemInfo(hMenu, idCmd, MF_BYCOMMAND, info)
    
    GetMenuArrow = False
    If info.hSubMenu <> 0 Then GetMenuArrow = True
End Function
