Attribute VB_Name = "modSHChange"
Option Explicit
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'This is from PSC, tell me if it's yours
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Private lSHNotify As Long       'the one and only shell change notification handle for the desktop folder

Public Type PIDLSTRUCT
     pidl As Long               'Fully qualified pidl (relative to the desktop folder)
                                '0 can also be specified for the desktop folder.
     bWatchSubFolders As Long   '(it's actually a Boolean, but we'll go Long because
                                'of VB's DWORD struct alignment).
End Type

Declare Function SHChangeNotifyRegister Lib "shell32" Alias "#2" (ByVal hwnd As Long, ByVal uFlags As SHCN_ItemFlags, ByVal dwEventID As SHCN_EventIDs, ByVal uMsg As Long, ByVal cItems As Long, lpps As PIDLSTRUCT) As Long
Declare Function SHChangeNotifyDeregister Lib "shell32" Alias "#4" (ByVal hNotify As Long) As Boolean

Type SHNOTIFYSTRUCT
    dwItem1 As Long
    dwItem2 As Long
End Type

Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As SHCN_EventIDs, ByVal uFlags As SHCN_ItemFlags, ByVal dwItem1 As Long, ByVal dwItem2 As Long)

'Shell notification event IDs
Public Enum SHCN_EventIDs
    SHCNE_RENAMEITEM = &H1          '(D) A non-folder item has been renamed.
    SHCNE_CREATE = &H2              '(D) A non-folder item has been created.
    SHCNE_DELETE = &H4              '(D) A non-folder item has been deleted.
    SHCNE_MKDIR = &H8               '(D) A folder item has been created.
    SHCNE_RMDIR = &H10              '(D) A folder item has been removed.
    SHCNE_MEDIAINSERTED = &H20      '(G) Storage media has been inserted into a drive.
    SHCNE_MEDIAREMOVED = &H40       '(G) Storage media has been removed from a drive.
    SHCNE_DRIVEREMOVED = &H80       '(G) A drive has been removed.
    SHCNE_DRIVEADD = &H100          '(G) A drive has been added.
    SHCNE_NETSHARE = &H200          '    A folder on the local computer is being shared via the network.
    SHCNE_NETUNSHARE = &H400        '    A folder on the local computer is no longer being shared via the network.
    SHCNE_ATTRIBUTES = &H800        '(D) The attributes of an item or folder have changed.
    SHCNE_UPDATEDIR = &H1000        '(D) The contents of an existing folder have changed, but the folder still exists and has not been renamed.
    SHCNE_UPDATEITEM = &H2000       '(D) An existing non-folder item has changed, but the item still exists and has not been renamed.
    SHCNE_SERVERDISCONNECT = &H4000 '    The computer has disconnected from a server.
    SHCNE_UPDATEIMAGE = &H8000&     '(G) An image in the system image list has changed.
    SHCNE_DRIVEADDGUI = &H10000     '(G) A drive has been added and the shell should create a new window for the drive.
    SHCNE_RENAMEFOLDER = &H20000    '(D) The name of a folder has changed.
    SHCNE_FREESPACE = &H40000       '(G) The amount of free space on a drive has changed.
    
    #If (WIN32_IE >= &H400) Then
        SHCNE_EXTENDED_EVENT = &H4000000 '(G) Not currently used.
    #End If
    
    SHCNE_ASSOCCHANGED = &H8000000  '(G) A file type association has changed.
    SHCNE_DISKEVENTS = &H2381F      '(D) Specifies a combination of all of the disk event identifiers.
    SHCNE_GLOBALEVENTS = &HC0581E0  '(G) Specifies a combination of all of the global event identifiers.
    SHCNE_ALLEVENTS = &H7FFFFFFF
    SHCNE_INTERRUPT = &H80000000    '    The specified event occurred as a result of a system interrupt. It is stripped out before the clients of SHCNNotify_ see it.
End Enum

#If (WIN32_IE >= &H400) Then
    Public Const SHCNEE_ORDERCHANGED = &H2 'dwItem2 is the pidl of the changed folder
#End If

'Notification flags
'uFlags & SHCNF_TYPE is an ID which indicates what dwItem1 and dwItem2 mean
Public Enum SHCN_ItemFlags
    SHCNF_IDLIST = &H0          'LPITEMIDLIST
    SHCNF_PATHA = &H1           'path name
    SHCNF_PRINTERA = &H2        'printer friendly name
    SHCNF_DWORD = &H3           'DWORD
    SHCNF_PATHW = &H5           'path name
    SHCNF_PRINTERW = &H6        'printer friendly name
    SHCNF_TYPE = &HFF
    SHCNF_FLUSH = &H1000        'Flushes the system event buffer. The function does not return until the system is finished processing the given event.
    SHCNF_FLUSHNOWAIT = &H2000  'Flushes the system event buffer. The function returns immediately regardless of whether the system is finished processing the given event.
    
    #If UNICODE Then
        SHCNF_PATH = SHCNF_PATHW
        SHCNF_PRINTER = SHCNF_PRINTERW
    #Else
        SHCNF_PATH = SHCNF_PATHA
        SHCNF_PRINTER = SHCNF_PRINTERA
    #End If

End Enum


'Registers shell change notification.
Public Function SHNotify_Register() As Boolean
    Dim ps As PIDLSTRUCT, pidl As Long
    
    ps.pidl = 0
    ps.bWatchSubFolders = True
    'Register the notification, specifying that we want the dwItem1
    'and dwItem2 members of the SHNOTIFYSTRUCT to be pidls. We're
    'watching all events.
    lSHNotify = SHChangeNotifyRegister(frmMain.hwnd, SHCNF_TYPE Or SHCNF_IDLIST, SHCNE_ALLEVENTS Or SHCNE_INTERRUPT, WM_SHNOTIFY, 1, ps)
    SHNotify_Register = CBool(lSHNotify)

End Function

'Unregisters shell change notification.
Public Function SHNotify_Unregister() As Boolean
  'If we have a registered notification handle.
  If lSHNotify Then
   'Unregister it. If the call is successful, zero the handle's variable
   If SHChangeNotifyDeregister(lSHNotify) Then lSHNotify = 0: SHNotify_Unregister = True
  End If
End Function


Public Sub NotificationReceipt(wParam As Long, lParam As Long)
    On Error Resume Next
    Dim SHNS As SHNOTIFYSTRUCT
       
    'Fill the SHNOTIFYSTRUCT from it's pointer.
    MoveMemory SHNS, ByVal wParam, Len(SHNS)

    'lParam is the ID of the notification event,
    'one of the SHCN_EventIDs.
    Select Case lParam
      Case SHCNE_UPDATEIMAGE
        Dim iImage As Long
        MoveMemory iImage, ByVal SHNS.dwItem1 + 2, 4
        'recycle bin
        If iImage = 31 Or iImage = 32 Then _
          DrawIcon frmMain.ImgIcon(1).Tag, frmMain.ImgIcon(1)
      Case SHCNE_CREATE: FillIcons
      Case SHCNE_DELETE: FillIcons
      Case SHCNE_UPDATEDIR: FillIcons
      Case SHCNE_RENAMEFOLDER: FillIcons
      Case SHCNE_MEDIAINSERTED: DrawDrives
      Case SHCNE_MEDIAREMOVED: DrawDrives
      
    End Select
End Sub
