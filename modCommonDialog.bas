Attribute VB_Name = "modCommonDialog"
'This module is taken from Steve McMahon's gCommonDialog class and stripped from
'any unnecesarry checks and functions


Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Const MAX_FILE = 260
Public Const MAX_PATH = 260

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hwndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" _
    Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" _
    Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" _
    Alias "GetFileTitleA" (ByVal szFile As String, _
    ByVal szTitle As String, ByVal cbBuf As Long) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum
'Common Dialog ShowColor
Public Type TCHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Declare Function ChooseColor Lib "COMDLG32.DLL" _
    Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long

Public Enum EChooseColor
    CC_RGBInit = &H1
    CC_FullOpen = &H2
    CC_PreventFullOpen = &H4
    CC_ColorShowHelp = &H8
  ' Win95 only
    CC_SolidColor = &H80
    CC_AnyColor = &H100
  ' End Win95 only
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_EnableTemplateHandle = &H40
End Enum
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'Array of cust colors, last for life of app
Private alCustom(0 To 15) As Long, fNotFirst As Boolean

' ChooseColor wrapper
Public Function ShowColor(Color As Long, _
                   Optional Owner As Long = -1 _
                   ) As Boolean

    Dim chclr As TCHOOSECOLOR, i As Integer
    chclr.lStructSize = Len(chclr)
    
    ' Color must get reference variable to receive result
    ' Flags can get reference variable or constant with bit flags
    ' Owner can take handle of owning window
    If Owner <> -1 Then chclr.hwndOwner = Owner

    ' Assign color (default uninitialized value of zero is good default)
    chclr.rgbResult = Color

    ' Mask out unwanted bits
    Dim afMask As Long
    afMask = CLng(Not (CC_ENABLEHOOK Or CC_ENABLETEMPLATE))
    ' Pass in flags
    chclr.flags = afMask And (CC_RGBInit Or CC_FullOpen)
    
    ' If first time, initialize to white
    If Not fNotFirst Then
        ' Initialize with first 16 system interface colors
        For i = 0 To 15
            alCustom(i) = GetSysColor(i)
        Next
        fNotFirst = True
    End If
    
    chclr.lpCustColors = VarPtr(alCustom(0))
    ' All other fields zero

    m_lApiReturn = ChooseColor(chclr)
    Select Case m_lApiReturn
        ' Success
        Case 1: ShowColor = True: Color = chclr.rgbResult
        ' Cancelled or Error (ELSE) no error handling here
    Case Else: ShowColor = False: Color = -1
    End Select

End Function

Public Function ShowOpen(filename As String, _
                       Optional FileTitle As String, _
                       Optional FileMustExist As Boolean = True, _
                       Optional MultiSelect As Boolean = False, _
                       Optional ReadOnly As Boolean = False, _
                       Optional HideReadOnly As Boolean = False, _
                       Optional Filter As String = "All (*.*)| *.*", _
                       Optional FilterIndex As Long = 1, _
                       Optional InitDir As String, _
                       Optional DlgTitle As String, _
                       Optional DefaultExt As String, _
                       Optional Owner As Long = -1, _
                       Optional flags As Long = 0 _
                    ) As Boolean

    Dim opfile As OPENFILENAME, s As String, afFlags As Long
    
With opfile
    .lStructSize = Len(opfile)
    ' Add in specific flags and strip out non-VB flags
    .flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
            (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-ReadOnly * OFN_READONLY) Or _
             (-HideReadOnly * OFN_HIDEREADONLY) Or _
             (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hwndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' To make Windows-style filter, replace | and : with nulls
    s = Replace(Filter, "|", vbNullChar)
    s = Replace(s, ":", vbNullChar)
    
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

   ' Pad file and file title buffers to maximum path
    If (.flags And OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT Then
      s = filename & String$(8192 - Len(filename), 0)
      .lpstrFile = s
      .nMaxFile = 8192
      s = FileTitle & String$(8192 - Len(FileTitle), 0)
      .lpstrFileTitle = s
      .nMaxFileTitle = 8192
    Else
       s = filename & String$(MAX_PATH - Len(filename), 0)
      .lpstrFile = s
      .nMaxFile = MAX_PATH
      s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
      .lpstrFileTitle = s
      .nMaxFileTitle = MAX_FILE
    End If
    ' All other fields set to zero
    
   m_lApiReturn = GetOpenFileName(opfile)
   Select Case m_lApiReturn
   Case 1
      ' Success
      ShowOpen = True

      If (.flags And OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT Then
         filename = .lpstrFile
      Else
         filename = Left(.lpstrFile, lstrlen(.lpstrFile))
         FileTitle = Left(.lpstrFileTitle, lstrlen(.lpstrFileTitle))
      End If
      flags = .flags
      ' Return the filter index
      FilterIndex = .nFilterIndex
      ' Look up the filter the user selected and return that
      Filter = FilterLookup(.lpstrFilter, FilterIndex)
      If (.flags And OFN_READONLY) Then ReadOnly = True
      
   Case Else     'cancelled or extended error
      ' Cancelled
      ShowOpen = False
      filename = ""
      FileTitle = ""
      flags = 0
      FilterIndex = -1
      Filter = ""
   End Select
End With
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long, iEnd As Long, s As String
    iStart = 1
    If sFilters = "" Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function
