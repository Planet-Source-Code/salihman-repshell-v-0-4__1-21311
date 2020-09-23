Attribute VB_Name = "modFileSys"
'Get a string of all avalaible drives seperated by nulls
Declare Function GetLogicalDriveStrings Lib "kernel32" Alias _
    "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'get special folder locations(e.g. "frmStart Menu","Temp","Recent documents")
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As SpecialFolders, pidl As Long) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    ' an item id
    Public Type SHITEMID
        cb As Long
        abID(0) As Byte
    End Type
    ' an item id list, packed in SHITEMID.abID
    Public Type ITEMIDLIST
        mkid As SHITEMID
    End Type
    'constants that represent special folders
    Enum SpecialFolders
      CSIDL_DESKTOP = &H0
      CSIDL_INTERNET = &H1
      CSIDL_PROGRAMS = &H2
      CSIDL_CONTROLS = &H3
      CSIDL_PRINTERS = &H4
      CSIDL_PERSONAL = &H5
      CSIDL_FAVORITES = &H6
      CSIDL_STARTUP = &H7
      CSIDL_RECENT = &H8
      CSIDL_SENDTO = &H9
      CSIDL_BITBUCKET = &HA
      CSIDL_STARTMENU = &HB
      CSIDL_DESKTOPDIRECTORY = &H10
      CSIDL_DRIVES = &H11
      CSIDL_NETWORK = &H12
      CSIDL_NETHOOD = &H13
      CSIDL_FONTS = &H14
      CSIDL_TEMPLATES = &H15
      CSIDL_COMMON_STARTMENU = &H16
      CSIDL_COMMON_PROGRAMS = &H17
      CSIDL_COMMON_STARTUP = &H18
      CSIDL_COMMON_DESKTOPDIRECTORY = &H19
      CSIDL_APPDATA = &H1A
      CSIDL_PRINTHOOD = &H1B
      CSIDL_ALTSTARTUP = &H1D           ' // DBCS
      CSIDL_COMMON_ALTSTARTUP = &H1E    ' // DBCS
      CSIDL_COMMON_FAVORITES = &H1F
      CSIDL_INTERNET_CACHE = &H20
      CSIDL_COOKIES = &H21
      CSIDL_HISTORY = &H22
    End Enum

'APIs to access INI files and retrieve data
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal filename$)
    
'Find Files
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
    Const FILE_ATTRIBUTE_NORMAL = &H80
    Const FILE_ATTRIBUTE_HIDDEN = &H2
    Const FHIDDEN = FILE_ATTRIBUTE_HIDDEN
    Const FILE_ATTRIBUTE_DIRECTORY = &H10
    Const FDIRECTORY = FILE_ATTRIBUTE_DIRECTORY
    
    Public Type FILETIME
      dwLowDateTime As Long
      dwHighDateTime As Long
    End Type

    Public Type WIN32_FIND_DATA
      dwFileAttributes As Long
      ftCreationTime As FILETIME
      ftLastAccessTime As FILETIME
      ftLastWriteTime As FILETIME
      nFileSizeHigh As Long
      nFileSizeLow As Long
      dwReserved0 As Long
      dwReserved1 As Long
      cFileName As String * MAX_PATH
      cAlternate As String * 14
    End Type
    
'execute any kind of command
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'INI FUNCTIONS
Function GetKeyVal(ByVal filename As String, ByVal Section As String, ByVal Key As String)
    Dim RetVal As String, Worked As Integer
    RetVal = String$(255, 0)
    Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), filename)
    If Worked Then GetKeyVal = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
End Function
Function AddToINI(ByVal filename As String, ByVal Section As String, ByVal Key As String, ByVal KeyValue As String) As Integer
    WritePrivateProfileString Section, Key, KeyValue, filename
End Function

'CLEAR NULLS FROM API RETURNS
Public Function ClearNulls(ByVal strSource As String) As String
    Dim iPos As Integer
    iPos = InStr(strSource, Chr$(0))
    If iPos <> 0 Then ClearNulls = Trim$(Left$(strSource, iPos - 1))
End Function
'ADD A '\' IF NECASSARY TO THE PATH
Public Function ProperPath(ByVal Path As String)
    ProperPath = IIf(Right(Path, 1) = "\", Path, Path & "\")
End Function

'IF c:\dir\test.htm THEN test.htm
Public Function ExtractFilename(ByVal sPath As String, Optional CheckExt As Boolean) As String
    On Error Resume Next
    'this line is used if its a drive 'c:\'
    'otherwise a null string would be returned
    If Len(sPath) = 3 Or InStr(1, sPath, "\") = 0 Then ExtractFilename = sPath: Exit Function
    sPath = StrReverse(sPath)
    sPath = Left(sPath, InStr(sPath, "\") - 1)
    sPath = StrReverse(sPath)
    If CheckExt Then sPath = CheckExtension(sPath)
    ExtractFilename = sPath
End Function
'IF c:\dir\test.htm THEN c:\dir\
Public Function ExtractPath(ByVal sPath As String)
    ExtractPath = Left(sPath, InStrRev(sPath, "\"))
End Function

'Get the extension of the curr file
Function GetExtension(sPath As String) As String
    Dim sFile As String
    sFile = ExtractFilename(sPath)
    sFile = StrReverse(sFile)
    sFile = Left(sFile, InStr(1, sFile, "."))
    GetExtension = StrReverse(sFile)
End Function

'When displaying files, files with extensions lnk, pif, url
'don't show the extension
Function CheckExtension(ByVal sFile As String) As String
  Dim sTemp As String
  sTemp = LCase(Right(sFile, 4))
  If sTemp = ".lnk" Or sTemp = ".pif" Or sTemp = ".url" Then sFile = Left(sFile, Len(sFile) - 4)
  CheckExtension = sFile
End Function

'Get the location of special folders
Public Function GetSpecialfolder(ByVal CSIDL As SpecialFolders) As String
    Dim sPath As String, pidl As Long
    If SHGetSpecialFolderLocation(0&, CSIDL, pidl) = ERROR_SUCCESS Then
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then _
            GetSpecialfolder = ProperPath(ClearNulls(sPath))
    End If
End Function

'THIS IS USED TO SORT THE ARRAY OF FOLDERITEMS
Public Sub QuickSort(sArray() As String, inLow As Integer, inHi As Integer)
  
   Dim pivot As String, tmpSwap As String, tmpLow As Integer, tmpHi As Integer
   
   tmpLow = inLow: tmpHi = inHi
   pivot = sArray((inLow + inHi) * 0.5)
  
   While (tmpLow <= tmpHi)
      While (LCase(sArray(tmpLow)) < LCase(pivot) And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      While (LCase(pivot) < LCase(sArray(tmpHi)) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = sArray(tmpLow)
         sArray(tmpLow) = sArray(tmpHi)
         sArray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   Wend
  
   If (inLow < tmpHi) Then QuickSort sArray(), inLow, tmpHi
   If (tmpLow < inHi) Then QuickSort sArray(), tmpLow, inHi
  
End Sub
'returns files and/or folders in an sorted array
Public Function GetFilesFolders(ByVal sPath As String, bFiles As Boolean, _
                iUbound As Integer, NumFolders As Integer) As String()
    
  Dim Items() As String, FFind As WIN32_FIND_DATA
  Dim FindHnd As Long, FNext As Long, bShowCur As Boolean
  Dim bShowHiddenFiles As Boolean, sFile As String
  Dim Folders() As String, Files() As String, lFolders As Integer, lFiles As Integer
 
  
  FindHnd = FindFirstFile(ProperPath(sPath) & "*.*", FFind)
  'Init vars
  iUbound = -1: FNext = 1: lFolders = -1: lFiles = -1
  bShowHiddenFiles = CBool(GetSetting("ShowHiddenFiles", "0"))

  Do While FNext
   bShowCur = IIf((FFind.dwFileAttributes And FHIDDEN) = FHIDDEN, bShowHiddenFiles, True)
   sFile = ClearNulls(FFind.cFileName)
   
   If bShowCur And Len(sFile) And Left(sFile, 1) <> "." Then
      If ((FFind.dwFileAttributes And FDIRECTORY) = FDIRECTORY) Then
        lFolders = lFolders + 1
        ReDim Preserve Folders(lFolders)
        Folders(lFolders) = sFile
        iUbound = iUbound + 1
      ElseIf bFiles Then
        lFiles = lFiles + 1
        ReDim Preserve Files(lFiles)
        Files(lFiles) = sFile
        iUbound = iUbound + 1
      End If
   End If
   
   FNext = FindNextFile(FindHnd, FFind)
  Loop
  FindClose FindHnd
  NumFolders = lFolders
  
  If lFolders > 0 Then QuickSort Folders, 0, lFolders
  If lFiles > 0 Then QuickSort Files, 0, lFiles
  
  If iUbound > -1 Then
    ReDim Items(iUbound)
    For i = 0 To lFolders
      Items(i) = Folders(i)
    Next
    For i = 0 To lFiles
      Items(lFolders + 1 + i) = Files(i)
    Next
  End If
  GetFilesFolders = Items
End Function
'Get the windows directory
Public Function GetWinDir() As String
    Dim WD As Long, Windir As String
    
    Windir = Space(144)
    WD = GetWindowsDirectory(Windir, 144)
    GetWinDir = ProperPath(Trim(Windir))
End Function
'Get the caption of a window
Public Function GetCaption(hwnd As Long) As String
    Dim mCaption As String, lReturn As Long
    'get caption
    mCaption = Space(255)
    lReturn = GetWindowText(hwnd, mCaption, 255)
    GetCaption = Left(mCaption, lReturn)
End Function
'WHICH DRIVES ARE AVAILABLE
Function DrivesPresent(Optional UpperCase As Boolean) As String()
    Dim Drives() As String, strDrives As String
    
    strDrives = String$(255, Chr$(0))
    ret& = GetLogicalDriveStrings(255, strDrives)
    Drives = Split(Left(UCase(strDrives), InStr(1, strDrives, _
             Chr(0) & Chr(0)) - 1), Chr(0))
    DrivesPresent = Drives
End Function

Public Function ExecuteFile( _
            ByVal sFile As String, _
            Optional ByVal sParam As String = "", _
            Optional sDirectory As String = "", _
            Optional bExplore As Boolean = False) As Long
                
    ExecuteFile = ShellExecute(0&, IIf(bExplore, "explore", "open"), _
                               sFile, sParam, sDirectory, SW_SHOWDEFAULT)

End Function

'This function runs the links that are in the StartUp folder
'but it also reads HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\
'Currentversion\Run\"
'to see what other programs need to be run, it makes sure not to
'include specific Explorer programs, such as Systray.exe and taskmon.exe
'If you know about any other key that i should check or any programs
'of Microsoft to avoid please let me know
Public Function RunStartUpPrograms()
    Dim lUbound As Integer, lNumFolders As Integer
    Dim iCounter As Integer, sStartUpFolder As String
    Dim sStartup() As String
    
    sStartUpFolder = GetSpecialfolder(CSIDL_STARTUP)
    Items = GetFilesFolders(sStartUpFolder, True, lUbound, lNumFolders)
    
    'execute files in startup folder
    For iCounter = 0 To lUbound 'lNumFolders + 1 To lUbound
        ExecuteFile sStartUpFolder & Items(iCounter)
    Next
    
'    EnumRunValues sStartup
    For i = 0 To UBound(sStartup)
        ExecuteFile sStartup(i)
    Next
    
End Function
