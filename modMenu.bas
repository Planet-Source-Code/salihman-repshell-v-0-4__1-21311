Attribute VB_Name = "modMenu"
Public sMenuItems() As String   'array of menu strings, due to lack of better method
Public lPopUp As Long           'handle of the menu

' This creates the Menu
Public Sub MakeAPIMenu(Items, Optional SubMenu, Optional MemberOfSubNo, _
    Optional ByVal NumSubMenus As Long = 0, Optional sTag As String)
    'Items()         : Array containing items
    'SubMenu()       : Array , which submenu does belong to this item
    'SubMenuOfItem() : Array , to which submenu does it belong
    'NumSubMenus     : Number of submenus
    
    Dim lPopUpMenu() As Long  ' handle to the popup menu to display
    Dim lPlace() As Long
    Dim MI As MENUITEMINFO    ' describes menu items to add
    
    Dim cPos As POINTAPI  ' holds the current mouse coordinates
    Dim MenuSel As Long  ' ID of what the user selected in the popup menu
    Dim RetVal As Long
    Dim sMenuItem As String
    
    On Error Resume Next
    
    ReDim lPlace(NumSubMenus)
    ReDim lPopUpMenu(NumSubMenus)
    ReDim sMenuItems(UBound(Items))
        
    ' Create all menus
    lPopUpMenu(0) = CreatePopupMenu()
    lPopUp = lPopUpMenu(0)
    For i = 1 To NumSubMenus
        lPopUpMenu(i) = CreatePopupMenu()
    Next
    'Get mouse coordinates
    RetVal = GetCursorPos(cPos)
    
    'Create the structure which is the base for all menus:
    With MI
        .cbSize = Len(MI)
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU
    End With
    
    For i = 1000 To UBound(Items) + 1000
        sMenuItems(i - 1000) = Items(i - 1000)
        ' Add each item to its submenu
        With MI
            Select Case Items(i - 1000)
                Case "-": .fType = MFT_SEPARATOR Or MFT_OWNERDRAW
                Case "_": .fType = MFT_MENUBARBREAK Or MFT_OWNERDRAW
                Case Else: .fType = MFT_STRING Or MFT_OWNERDRAW
            End Select
            .fState = MFS_ENABLED
            
            .wID = i ' Assign this item an item identifier.
            .dwTypeData = Items(i - 1000)
            .cch = Len(Items(i - 1000))
            
            Sub1 = IIf(IsArray(SubMenu), Val(SubMenu(i - 1000)), 0)
            If Sub1 Then  '<> 0
                .hSubMenu = lPopUpMenu(Sub1)
            Else
                .hSubMenu = 0
            End If
        End With
        Sub2 = IIf(IsArray(MemberOfSubNo), Val(MemberOfSubNo(i - 1000)), 0)
        lPlace(Sub2) = lPlace(Sub2) + 1
        RetVal = InsertMenuItem(lPopUpMenu(Sub2), lPlace(Sub2), 1, MI)
    Next
    
    'returns wID
    MenuSel = TrackPopupMenu(lPopUpMenu(0), TPM_TOPALIGN Or TPM_RETURNCMD Or TPM_RIGHTALIGN Or TPM_RIGHTBUTTON, cPos.X, cPos.Y, 0, frmMain.hwnd, ByVal 0&)
    
    If MenuSel = 0 Then GoTo 1
    sMenuItem = Items(MenuSel - 1000)
    Playsound "succes"
    
    Select Case sTag
      Case "Ras": StartDialDialog sMenuItem
      Case "Shut Down"
        Select Case sMenuItem
          Case "RepShell Options": frmSettings.Show , frmMain
          Case "Control Panel"
            Shell "explorer.exe ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}", vbNormalFocus
          Case "Printers"
            Shell "explorer.exe ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{2227A280-3AEA-1069-A2DE-08002B30309D}", vbNormalFocus
          Case "Screen"
            Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0")
          Case "Background"
            Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0")
          Case "Screensaver"
            Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,1")
          Case "Options"
            Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2")
          Case "Settings"
            Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3")
          Case "Exit RepShell": ExitApp
          Case "Paste": MsgBox "Not yet implemented"
          Case Else
            If MsgBox("Are you sure you want to " & UCase(sMenuItem) & " ?", _
              vbOKCancel + vbExclamation) = vbOK Then ExitWindowsEx MenuSel - 1000, 0
        End Select
      Case "Multi"
        Select Case sMenuItem
        
        End Select
    End Select

1:  For i = 0 To NumSubMenus
        RetVal = DestroyMenu(lPopUpMenu(i))
    Next
    Erase sMenuItems
    Erase lPopUpMenu
    Erase lPlace
    lPopUp = 0
End Sub
