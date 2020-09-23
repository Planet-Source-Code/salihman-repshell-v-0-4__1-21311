Attribute VB_Name = "modRas"
Private Const RAS_MaxEntryName = 256
Private Const RAS_MaxDeviceType = 16
Private Const RAS_MaxDeviceName = 128

Private Type RASENTRYNAME
    dwSize As Long
    szEntryName(RAS_MaxEntryName) As Byte
End Type

Private Type RASCONN
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MaxEntryName) As Byte
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Private Type RASCONNSTATUS
    dwSize As Long
    RASCONNSTATE As Long
    dwError As Long
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Private Type RASPPPIP
    dwSize As Long
    dwError As Long
    szIpAddress(15) As Byte
    szServerIpAddress(15) As Byte
End Type

Private Type HOSTENT
    hName As Long           'Official name of the host (PC).
    hAliases As Long        'Null-terminated array of alternate names.
    hAddrType As Integer    'Type of address being returned.
    hLength As Integer      'Length of each address, in bytes.
    hAddrList As Long       'Null-terminated list of addresses for the host.
End Type

'****************************
'  RAS Functions - Declares
'****************************
Declare Function RasGetConnectStatus Lib "rasapi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpCb As Long, lpcConnections As Long) As Long
Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" (ByVal lpStrNull As String, ByVal lpszPhonebook As String, lprasentryname As RASENTRYNAME, lpCb As Long, lpCEntries As Long) As Long
Declare Function RasGetProjectionInfo Lib "rasapi32.dll" Alias _
    "RasGetProjectionInfoA" (ByVal hRasConn As Long, ByVal rasProjection As Long, _
    ByRef lpProjection As Any, ByRef lpCb As Long) As Long
    
Declare Function gethostbyaddr Lib "ws2_32.dll" (Addr As Long, _
    addrLen As Long, addrType As Long) As Long
Declare Function inet_addr Lib "ws2_32.dll" (ByVal sIPAddress As String) As Long

'Get conn speed
Private Function RasGetConnectionSpeed() As Long
    RasGetConnectionSpeed = ReadLong(HKEY_DYN_DATA, "PerfStats\StatData", "Dial-Up Adapter\ConnectSpeed", 0)
End Function

'name of connected entry
Private Function RasGetConnectedEntry() As String
Dim TRasCon(255) As RASCONN, lg As Long, lpcon As Long
Dim Tstatus As RASCONNSTATUS
    On Error Resume Next
    TRasCon(0).dwSize = LenB(TRasCon(0))
    lg = 256 * TRasCon(0).dwSize
    Call RasEnumConnections(TRasCon(0), lg, lpcon)
    Tstatus.dwSize = 160
    Call RasGetConnectStatus(TRasCon(0).hRasConn, Tstatus)
    If Tstatus.RASCONNSTATE <> RASCS_Disconnected And Tstatus.RASCONNSTATE <> 0 Then
        RasGetConnectedEntry = ClearNulls(StrConv(TRasCon(0).szEntryName, vbUnicode))
    Else
        RasGetConnectedEntry = "Disconnected"
    End If
End Function

'Start Dial-Up dialog
Public Sub StartDialDialog(dialupname As String)
    Call Shell("rundll32.exe rnaui.dll,RnaDial " & dialupname, 1)
End Sub

Public Function GetEntries() As String()
    Dim lResult, lConns, lSize As Long, i As Integer
    ReDim rasentry(64) As RASENTRYNAME
    Dim Entries() As String
    
    rasentry(0).dwSize = LenB(rasentry(0))
    lSize = rasentry(0).dwSize * 64
    lResult = RasEnumEntries(0&, 0&, rasentry(0), lSize, lConns)
    ReDim Entries(lConns - 1)
    For i = 0 To lConns - 1
        Entries(i) = ClearNulls(StrConv(rasentry(i).szEntryName, vbUnicode))
    Next
    'return an array of strings with the entries
    GetEntries = Entries
End Function

'Data 0 : Connected entry
'     1 : Connection Speed
'     2 : Connection Time
'     3 : Local IP
'     4 : Remote IP
Public Function GetConnectionData() As String()
    On Error Resume Next
    Dim lSpeed As Long, lpProjection As RASPPPIP, lpCb As Long
    Dim Data(5) As String, RemoteIP$, LocalIP$
    
    'Get connection info
    Data(0) = RasGetConnectedEntry
    
    lSpeed = RasGetConnectionSpeed
    Data(1) = IIf(Data(0) = "Disconnected", "Not available", _
                            CStr(lSpeed) & " bps")
    Data(2) = "Not yet implemented."
    
    lpProjection.dwSize = LenB(lpProjection)
    Call RasGetProjectionInfo(hRasConn, &H8021&, lpProjection, lpCb)
    
    LocalIP = Byte2String(lpProjection.szIpAddress)
    RemoteIP = Byte2String(lpProjection.szServerIpAddress)
    If Len(LocalIP) = 0 Then LocalIP = "0.0.0.0"
    If Len(RemoteIP) = 0 Then RemoteIP = "0.0.0.0"
    
    Data(3) = LocalIP
    Data(4) = RemoteIP
    Data(5) = GetHostByIP(RemoteIP)
        
    GetConnectionData = Data
End Function

Private Function MakeIP(ByVal lData As Long) As String
Dim s, sResult As String
    If Len(CStr(lData)) < 4 Then
        MakeIP = CStr(lData)
        Exit Function
    End If
    s = CStr(lData)
    Do While Len(s) > 3
        sResult = "." & Right(s, 3) & sResult
        s = Left(s, Len(s) - 3)
    Loop
    sResult = s & sResult
    MakeIP = sResult
End Function

Private Function GetHostByIP(ByVal strIP As String) As String
    If Len(strIP) < 1 Then GoTo 1
    
    Dim Host As HOSTENT
    Dim lngIP As Long
    Dim strHost As String * 255
    
    lngIP = inet_addr(strIP & Chr(0))
    
    apiError = gethostbyaddr(lngIP, Len(lngIP), AF_INET)
    If apiError = 0 Then GoTo 1
    
    MoveMemory Host, apiError, Len(Host)
    MoveMemory ByVal strHost, Host.hName, 255

    GetHostByIP = ClearNulls(strHost)
    Exit Function
1:  GetHostByIP = "No host detected."
End Function

Private Function Byte2String(bString() As Byte) As String
Dim i As Integer
    While bString(i) <> 0&
        Byte2String = Byte2String & Chr(bString(i))
        i = i + 1
    Wend
End Function
