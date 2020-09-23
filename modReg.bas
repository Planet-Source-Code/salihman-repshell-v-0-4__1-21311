Attribute VB_Name = "modReg"
'constants
Public Enum RootKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_DYN_DATA = &H80000006
End Enum

Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const KEY_ALL_ACCESS = &H3F
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2

Public Const ERROR_NO_MORE_ITEMS = 259&

'Registry functions
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As RootKey, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As RootKey, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As RootKey, ByVal lpSubKey As String) As Long

' Note that if you declare the lpData parameter as String, you must pass it By Value.
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

'Used in RegPaths
Public Const Options = "Options\Appearance\"
Public Const General = "General\"

Public Function EnumRunValues(sRetArray As Variant) As Long

    Dim hKey&, lResult&, iCounter%
    Dim sValueName As String * 255
    Dim sValue As String * 255
    Dim sValues() As String
    Dim sTemp As String
    
    lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
        "Software\Microsoft\Windows\Currentversion\Run\", 0&, _
        KEY_ALL_ACCESS, hKey)
    
    'if succesfull opened
    If lResult = ERROR_SUCCESS Then
        'enumerate all values until no more
        Do
            lResult = RegEnumValue(hKey, iCounter, sValueName, 255, 0&, REG_SZ, sValue, 255)
            sTemp = LCase(ClearNulls(sValue))
            'redim array and save new value but check if specific explorer
            'program first
            If sTemp <> "systray.exe" And (InStr(1, sTemp, "taskmon.exe") = 0) Then
                ReDim Preserve sValues(iCounter)
                sValues(iCounter) = ClearNulls(sValue)
                'increment counter
                iCounter = iCounter + 1
            End If
        Loop Until lResult = ERROR_NO_MORE_ITEMS
        sRetArray = sValues
        
        Call RegCloseKey(hKey)
        EnumRunValues = 1
        Exit Function
    End If
    Call RegCloseKey(hKey)
    EnumRunValues = 0
End Function

Public Function CreateKey(ByVal lRootKey As RootKey, ByVal sKeyName As String) As Long
Dim hKey As Long
    lResult = RegCreateKey(lRootKey, sKeyName, hKey)
    Call RegCloseKey(hKey)
    CreateKey = lResult
End Function

Public Sub WriteString(ByVal lRootKey As RootKey, ByVal sPath As String, ByVal sValueName As String, sValueData As String)
Dim hKey As Long, lResult As Long
    lResult = RegOpenKeyEx(lRootKey, sPath, vbNull, KEY_SET_VALUE, hKey)
    If lResult <> ERROR_SUCCESS Then lResult = CreateKey(lRootKey, sPath)
    If lResult = ERROR_SUCCESS Then
        sValueData = sValueData & Chr(0)
        Call RegSetValueEx(hKey, sValueName, vbNull, REG_SZ, ByVal sValueData, Len(sValueData))
        Call RegCloseKey(hKey)
    End If
End Sub

Public Function ReadString(ByVal lRootKey As RootKey, ByVal sPath As String, ByVal sValueName As String, ByRef sDefault As String) As String
Dim hKey As Long, lResult As Long, lValueType As Long, lDataBufSize As Long, strBuf As String
    lResult = RegOpenKeyEx(lRootKey, sPath, 0, KEY_QUERY_VALUE, hKey)
    If lResult = ERROR_SUCCESS Then
        lResult = RegQueryValueEx(hKey, sValueName, vbNull, lValueType, ByVal 0&, lDataBufSize)
        If lValueType = REG_SZ Then
            strBuf = Space$(lDataBufSize)
            lResult = RegQueryValueEx(hKey, sValueName, vbNull, REG_SZ, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                ReadString = ClearNulls(strBuf)
            Else
                ReadString = sDefault
            End If
        Else
            ReadString = sDefault
        End If
    Else
        lResult = CreateKey(lRootKey, sPath)
        ReadString = sDefault
    End If
    Call RegCloseKey(hKey)
End Function

Public Sub WriteLong(ByVal iKey As RootKey, ByVal sPath As String, ByVal sValueName As String, lValue As Long)
Dim hKey As Long, lResult As Long
    lResult = RegOpenKeyEx(iKey, sPath, 0, KEY_SET_VALUE, hKey)
    If lResult = ERROR_SUCCESS Then
        Call RegSetValueEx(hKey, sValueName, vbNull, REG_DWORD, lValue, LenB(lValue))
        Call RegCloseKey(hKey)
    Else
        lResult = CreateKey(lRootKey, sPath)
    End If
End Sub

Public Function ReadLong(ByVal lRootKey As RootKey, sPath As String, sValueName As String, lDefault As Long) As Long
Dim hKey As Long, lResult As Long, lData As Long
    lResult = RegOpenKeyEx(lRootKey, sPath, 0, KEY_QUERY_VALUE, hKey)
    If lResult = ERROR_SUCCESS Then
        lResult = RegQueryValueEx(hKey, sValueName, 0&, REG_DWORD, lData, LenB(lData))
        If lResult = ERROR_SUCCESS Then
            ReadLong = lData
        Else
            ReadLong = lDefault
        End If
        Call RegCloseKey(hKey)
    Else
        lResult = CreateKey(lRootKey, sPath)
        ReadLong = lDefault
    End If
End Function

Public Sub DeleteKey(ByVal lRootKey As RootKey, ByVal sSubKey As String)
    Call RegDeleteKey(lRootKey, sSubKey)
End Sub

Public Sub DeleteValue(ByVal lRootKey As RootKey, ByVal strPath As String, ByVal sValueName As String)
Dim hKey As Long, lResult As Long
    lResult = RegOpenKeyEx(lRootKey, strPath, 0&, KEY_SET_VALUE, hKey)
    If lResult = ERROR_SUCCESS Then
        Call RegDeleteValue(hKey, sValueName)
        Call RegCloseKey(hKey)
    End If
End Sub

Public Function KeyExists(ByVal lRootKey As RootKey, ByVal strKeyName As String) As Boolean
Dim hKey As Long, lResult As Long
    lResult = RegOpenKeyEx(lRootKey, strKeyName, 0, 0&, hKey)
    Call RegCloseKey(hKey)
    KeyExists = (lResult = ERROR_SUCCESS)
End Function


'Functions to shorten RepShell Registry calls
Public Function SaveSetting(ByVal sKey As String, ByVal sValue As String, Optional sPath As String = Options)
    WriteString HKEY_CURRENT_USER, "Software\RepShell\" & sPath, sKey, sValue
End Function

Public Function GetSetting(ByVal sKey As String, Optional ByVal sDefault As String = "", Optional sPath As String = Options) As String
    GetSetting = ReadString(HKEY_CURRENT_USER, "Software\RepShell\" & sPath, sKey, sDefault)
End Function

Public Function SaveLong(ByVal sKey As String, ByVal sValue As Long, Optional sPath As String = Options)
    WriteLong HKEY_CURRENT_USER, "Software\RepShell\" & sPath, sKey, sValue
End Function

Public Function GetLong(ByVal sKey As String, Optional ByVal sDefault As Long = 0, Optional sPath As String = Options) As Long
    GetLong = ReadLong(HKEY_CURRENT_USER, "Software\RepShell\" & sPath, sKey, sDefault)
End Function
