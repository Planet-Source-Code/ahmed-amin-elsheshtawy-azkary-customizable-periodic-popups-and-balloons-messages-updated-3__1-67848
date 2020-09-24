Attribute VB_Name = "mRasAPI"

'C:\VBmyProjects\Docs\DialUpRAS

Public Enum RASCONNSTATE
   RASCS_OpenPort = 0
   RASCS_PortOpened = 1
   RASCS_ConnectDevice = 2
   RASCS_DeviceConnected = 3
   RASCS_AllDevicesConnected = 4
   RASCS_Authenticate = 5
   RASCS_AuthNotify = 6
   RASCS_AuthRetry = 7
   RASCS_AuthCallback = 8
   RASCS_AuthChangePassword = 9
   RASCS_AuthProject = 10
   RASCS_AuthLinkSpeed = 11
   RASCS_AuthAck = 12
   RASCS_ReAuthenticate = 13
   RASCS_Authenticated = 14
   RASCS_PrepareForCallback = 15
   RASCS_WaitForModemReset = 16
   RASCS_WaitForCallback = 17
   RASCS_Projected = 18
   RASCS_StartAuthentication = 19
   RASCS_CallbackComplete = 20
   RASCS_LogonNetwork = 21
   RASCS_SubEntryConnected = 22
   RASCS_SubEntryDisconnected = 23
   RASCS_Interactive = &H1000
   RASCS_RetryAuthentication = &H1001
   RASCS_CallbackSetByCaller = &H1002
   RASCS_PasswordExpired = &H1003
   RASCS_InvokeEapUI = &H1004
   RASCS_Connected = &H2000
   RASCS_Disconnected = &H2001
End Enum

Public Type VBRASCONNSTATUS
      lRasConnState As RASCONNSTATE
      dwError As Long
      sDeviceType As String
      sDeviceName As String
      sNTPhoneNumber As String
End Type

Public Type VBRASCTRYINFO
    CountryCode As Long
    CountryID As Long
    CountryName As String
    NextCountryID As Long
End Type

Public Type VBRASDEVINFO
   DeviceType As String
   DeviceName As String
End Type

Public Type RASIPADDR
    a As Byte
    b As Byte
    c As Byte
    d As Byte
End Type

Public Enum RasEntryOptions
   RASEO_UseCountryAndAreaCodes = &H1
   RASEO_SpecificIpAddr = &H2
   RASEO_SpecificNameServers = &H4
   RASEO_IpHeaderCompression = &H8
   RASEO_RemoteDefaultGateway = &H10
   RASEO_DisableLcpExtensions = &H20
   RASEO_TerminalBeforeDial = &H40
   RASEO_TerminalAfterDial = &H80
   RASEO_ModemLights = &H100
   RASEO_SwCompression = &H200
   RASEO_RequireEncryptedPw = &H400
   RASEO_RequireMsEncryptedPw = &H800
   RASEO_RequireDataEncryption = &H1000
   RASEO_NetworkLogon = &H2000
   RASEO_UseLogonCredentials = &H4000
   RASEO_PromoteAlternates = &H8000
   RASEO_SecureLocalFiles = &H10000
   RASEO_RequireEAP = &H20000
   RASEO_RequirePAP = &H40000
   RASEO_RequireSPAP = &H80000
   RASEO_Custom = &H100000
   RASEO_PreviewPhoneNumber = &H200000
   RASEO_SharedPhoneNumbers = &H800000
   RASEO_PreviewUserPw = &H1000000
   RASEO_PreviewDomain = &H2000000
   RASEO_ShowDialingProgress = &H4000000
   RASEO_RequireCHAP = &H8000000
   RASEO_RequireMsCHAP = &H10000000
   RASEO_RequireMsCHAP2 = &H20000000
   RASEO_RequireW95MSCHAP = &H40000000
   RASEO_CustomScript = &H80000000
End Enum

Public Enum RASNetProtocols
   RASNP_NetBEUI = &H1
   RASNP_Ipx = &H2
   RASNP_Ip = &H4
End Enum

Public Enum RasFramingProtocols
   RASFP_Ppp = &H1
   RASFP_Slip = &H2
   RASFP_Ras = &H4
End Enum
Public Type VBRasEntry
   Options As RasEntryOptions
   CountryID As Long
   CountryCode As Long
   AreaCode As String
   LocalPhoneNumber As String
   AlternateNumbers As String
   ipAddr As RASIPADDR
   ipAddrDns As RASIPADDR
   ipAddrDnsAlt As RASIPADDR
   ipAddrWins As RASIPADDR
   ipAddrWinsAlt As RASIPADDR
   FrameSize As Long
   fNetProtocols As RASNetProtocols
   FramingProtocol As RasFramingProtocols
   ScriptName As String
   AutodialDll As String
   AutodialFunc As String
   DeviceType As String
   DeviceName As String
   X25PadType As String
   X25Address As String
   X25Facilities As String
   X25UserData As String
   Channels As Long
   NT4En_SubEntries As Long
   NT4En_DialMode As Long
   NT4En_DialExtraPercent As Long
   NT4En_DialExtraSampleSeconds As Long
   NT4En_HangUpExtraPercent As Long
   NT4En_HangUpExtraSampleSeconds As Long
   NT4En_IdleDisconnectSeconds As Long
   Win2000_Type As Long
   Win2000_EncryptionType As Long
   Win2000_CustomAuthKey As Long
   Win2000_guidId(0 To 15) As Byte
   Win2000_CustomDialDll As String
   Win2000_VpnStrategy As Long
End Type

Type VBRASCONN
   hRasConn As Long
   sEntryName As String
   sDeviceType As String
   sDeviceName As String
   sPhonebook  As String
   lngSubEntry As Long
   guidEntry(15) As Byte
End Type

Public Declare Function RasEnumDevices _
   Lib "rasapi32.dll" Alias "RasEnumDevicesA" ( _
        lpRasDevInfo As Any, _
        lpcB As Long, _
        lpCDevices As Long _
) As Long

Declare Function RasGetCountryInfo _
   Lib "rasapi32.dll" Alias "RasGetCountryInfoA" _
   (lpRasCtryInfo As Any, lpdwSize As Long) As Long

Public Declare Function RasGetErrorString _
     Lib "rasapi32.dll" Alias "RasGetErrorStringA" _
      (ByVal uErrorValue As Long, ByVal lpszErrorString As String, _
       cBufSize As Long) As Long

Public Declare Function FormatMessage _
     Lib "kernel32" Alias "FormatMessageA" _
      (ByVal dwFlags As Long, lpSource As Any, _
       ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
       ByVal lpBuffer As String, ByVal nSize As Long, _
       Arguments As Long) As Long

Public Type VBRasEntryName
   entryName As String
   Win2000_SystemPhonebook As Boolean
   PhonebookPath As String
End Type

Public Declare Function RasEnumEntries _
    Lib "rasapi32.dll" Alias "RasEnumEntriesA" _
    (ByVal lpStrNull As String, ByVal lpszPhonebook As String, _
    lpRasEntryName As Any, lpcB As Long, lpCEntries As Long) As Long


Declare Function RasHangUp _
         Lib "rasapi32.dll" Alias "RasHangUpA" _
        (ByVal hRasConn As Long) As Long

Public Declare Function RasGetEntryDialParams _
      Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" _
        (ByVal lpszPhonebook As String, _
        lpRasDialParams As Any, _
        blnPasswordRetrieved As Long) As Long

Public Declare Function RasSetEntryDialParams _
      Lib "rasapi32.dll" Alias "RasSetEntryDialParamsA" _
        (ByVal lpszPhonebook As String, _
        lpRasDialParams As Any, _
        ByVal blnRemovePassword As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)



Private Declare Function RegisterWindowMessage Lib "user32" Alias _
   "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Const RASDIALEVENT = "RasDialEvent"
Private Const WM_RASDIALEVENT = &HCCCD&
Private m_RasMessage As Long

Public Declare Function RasDial _
      Lib "rasapi32.dll" Alias "RasDialA" _
      (lpRasDialExtensions As Any, _
       ByVal lpszPhonebook As String, _
       lpRasDialParams As Any, _
       ByVal dwNotifierType As Long, _
       ByVal hwndNotifier As Long, _
       lphRasConn As Long) _
As Long


Public Type VBRasDialParams
    entryName As String
    PhoneNumber As String
    CallbackNumber As String
    UserName As String
    Password As String
    Domain As String
    SubEntryIndex As Long
    RasDialFunc2CallbackId As Long
End Type


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
   ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
   ByVal lParam As Any) As Long

Private Declare Function CallWindowProc Lib "user32" Alias _
   "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
   ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias _
   "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
   ByVal wNewWord As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias _
   "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function RasGetConnectStatus Lib "rasapi32.dll" _
Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, _
lpStatus As Any) As Long

Public Declare Function RasRenameEntry _
      Lib "rasapi32.dll" Alias "RasRenameEntryA" _
        (ByVal lpszPhonebook As String, _
        ByVal lpszOldEntry As String, _
        ByVal lpszNewEntry As String) As Long

Public Declare Function RasDeleteEntry _
      Lib "rasapi32.dll" Alias "RasDeleteEntryA" _
        (ByVal lpszPhonebook As String, _
    ByVal lpszEntry As String) As Long

Public Declare Function RasValidateEntryName _
      Lib "rasapi32.dll" Alias "RasValidateEntryNameA" _
        (ByVal lpszPhonebook As String, _
        ByVal lpszEntry As String) As Long

Declare Function RasCreatePhonebookEntry _
      Lib "rasapi32.dll" Alias "RasCreatePhonebookEntryA" _
        (ByVal hwnd As Long, _
        ByVal lpszPhonebook As String) As Long

Declare Function RasEditPhonebookEntry _
      Lib "rasapi32.dll" Alias "RasEditPhonebookEntryA" _
        (ByVal hwnd As Long, _
        ByVal lpszPhonebook As String, _
        ByVal lpszEntryName As String) As Long

Public Declare Function RasGetEntryProperties _
      Lib "rasapi32.dll" Alias "RasGetEntryPropertiesA" _
       (ByVal lpszPhonebook As String, _
        ByVal lpszEntry As String, _
        lpRasEntry As Any, _
        lpdwEntryInfoSize As Long, _
        lpbDeviceInfo As Any, _
        lpdwDeviceInfoSize As Long) _
As Long

Public Declare Function RasSetEntryProperties _
      Lib "rasapi32.dll" Alias "RasSetEntryPropertiesA" _
        (ByVal lpszPhonebook As String, _
        ByVal lpszEntry As String, _
        lpRasEntry As Any, _
        ByVal dwEntryInfoSize As Long, _
        lpbDeviceInfo As Any, _
        ByVal dwDeviceInfoSize As Long) _
As Long

Public Declare Function RasEnumConnections _
      Lib "rasapi32.dll" Alias "RasEnumConnectionsA" _
         (lpRasconn As Any, _
          lpcB As Long, _
          lpcConnections As Long) As Long

Public Declare Sub SleepAPI Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const ERROR_INVALID_HANDLE = 6&

Private Const GWL_WNDPROC = (-4)

Private m_hWnd As Long
Private m_lpPrev As Long

Public Sub Hook(ByVal hwnd As Long)
   If m_hWnd Then Call UnHook
   m_hWnd = hwnd
   m_lpPrev = SetWindowLong(m_hWnd, GWL_WNDPROC, AddressOf WindowProc)
   
   'register RasMEvent Message
   If m_RasMessage = 0 Then
      m_RasMessage = RegisterWindowMessage(RASDIALEVENT)
      If m_RasMessage = 0 Then m_RasMessage = WM_RASDIALEVENT
   End If
   
End Sub

Public Sub UnHook()
   If m_hWnd Then
      Call SetWindowLong(m_hWnd, GWL_WNDPROC, m_lpPrev)
      m_hWnd = 0
   End If
End Sub

Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
   If uMsg = m_RasMessage Then
      Call RasDialFunc(uMsg, wParam, lParam)
   End If
   WindowProc = CallWindowProc(m_lpPrev, hwnd, uMsg, wParam, lParam)
End Function

Sub RasDialFunc(ByVal unMsg As Long, _
       ByVal rasConnectionState As Long, ByVal dwError As Long)
   
   ' do stuff here, such as check for errors
   ' or check the rasConnectionState
   'Debug.Print hRasConn, Hex$(rasConnectionState), dwError
   If rasConnectionState = RASCS_Connected Then
      '  Debug.Print "connected"
   ElseIf rasConnectionState = RASCS_Disconnected Then
      '  Debug.Print "disconnected"
   End If
   
   If dwError <> 0 Then
      'Debug.Print (VBRASErrorHandler(dwError))
      '  Debug.Print "disconnected"
   End If
   '  Warning !!  if an error occurs or you get disconnected
   ' you should still call RasHangUp
End Sub

Function VBRasGetEntryProperties(strEntryName As String, _
         clsRasEntry As VBRasEntry, _
         Optional strPhoneBook As String) As Long
   
   Dim rtn As Long, lngCb As Long, lngBuffLen As Long
   Dim b() As Byte
   Dim lngPos As Long, lngStrLen As Long

   rtn = RasGetEntryProperties(vbNullString, vbNullString, _
                           ByVal 0&, lngCb, ByVal 0&, ByVal 0&)
   
   rtn = RasGetEntryProperties(strPhoneBook, strEntryName, _
                        ByVal 0&, lngBuffLen, ByVal 0&, ByVal 0&)
   
   If rtn <> 603 Then VBRasGetEntryProperties = rtn: Exit Function
   
   ReDim b(lngBuffLen - 1)
   CopyMemory b(0), lngCb, 4
   
   rtn = RasGetEntryProperties(strPhoneBook, strEntryName, _
                           b(0), lngBuffLen, ByVal 0&, ByVal 0&)
   
   VBRasGetEntryProperties = rtn
   If rtn <> 0 Then Exit Function
   
   CopyMemory clsRasEntry.Options, b(4), 4
   CopyMemory clsRasEntry.CountryID, b(8), 4
   CopyMemory clsRasEntry.CountryCode, b(12), 4
   CopyByteToTrimmedString clsRasEntry.AreaCode, b(16), 11
   CopyByteToTrimmedString clsRasEntry.LocalPhoneNumber, b(27), 129
   
   CopyMemory lngPos, b(156), 4
   If lngPos <> 0 Then
     lngStrLen = lngBuffLen - lngPos
     clsRasEntry.AlternateNumbers = String(lngStrLen, 0)
     CopyMemory ByVal clsRasEntry.AlternateNumbers, _
               b(lngPos), lngStrLen
   End If
   
   CopyMemory clsRasEntry.ipAddr, b(160), 4
   CopyMemory clsRasEntry.ipAddrDns, b(164), 4
   CopyMemory clsRasEntry.ipAddrDnsAlt, b(168), 4
   CopyMemory clsRasEntry.ipAddrWins, b(172), 4
   CopyMemory clsRasEntry.ipAddrWinsAlt, b(176), 4
   CopyMemory clsRasEntry.FrameSize, b(180), 4
   CopyMemory clsRasEntry.fNetProtocols, b(184), 4
   CopyMemory clsRasEntry.FramingProtocol, b(188), 4
   CopyByteToTrimmedString clsRasEntry.ScriptName, b(192), 260
   CopyByteToTrimmedString clsRasEntry.AutodialDll, b(452), 260
   CopyByteToTrimmedString clsRasEntry.AutodialFunc, b(712), 260
   CopyByteToTrimmedString clsRasEntry.DeviceType, b(972), 17
      If lngCb = 1672& Then lngStrLen = 33 Else lngStrLen = 129
   CopyByteToTrimmedString clsRasEntry.DeviceName, b(989), lngStrLen
      lngPos = 989 + lngStrLen
   CopyByteToTrimmedString clsRasEntry.X25PadType, b(lngPos), 33
      lngPos = lngPos + 33
   CopyByteToTrimmedString clsRasEntry.X25Address, b(lngPos), 201
      lngPos = lngPos + 201
   CopyByteToTrimmedString clsRasEntry.X25Facilities, b(lngPos), 201
      lngPos = lngPos + 201
   CopyByteToTrimmedString clsRasEntry.X25UserData, b(lngPos), 201
      lngPos = lngPos + 203
   CopyMemory clsRasEntry.Channels, b(lngPos), 4
   
   If lngCb > 1768 Then 'NT4 Enhancements & Win2000
      CopyMemory clsRasEntry.NT4En_SubEntries, b(1768), 4
      CopyMemory clsRasEntry.NT4En_DialMode, b(1772), 4
      CopyMemory clsRasEntry.NT4En_DialExtraPercent, b(1776), 4
      CopyMemory clsRasEntry.NT4En_DialExtraSampleSeconds, b(1780), 4
      CopyMemory clsRasEntry.NT4En_HangUpExtraPercent, b(1784), 4
      CopyMemory clsRasEntry.NT4En_HangUpExtraSampleSeconds, b(1788), 4
      CopyMemory clsRasEntry.NT4En_IdleDisconnectSeconds, b(1792), 4
      
      If lngCb > 1796 Then ' Win2000
         CopyMemory clsRasEntry.Win2000_Type, b(1796), 4
         CopyMemory clsRasEntry.Win2000_EncryptionType, b(1800), 4
         CopyMemory clsRasEntry.Win2000_CustomAuthKey, b(1804), 4
         CopyMemory clsRasEntry.Win2000_guidId(0), b(1808), 16
         CopyByteToTrimmedString _
                  clsRasEntry.Win2000_CustomDialDll, b(1824), 260
         CopyMemory clsRasEntry.Win2000_VpnStrategy, b(2084), 4
      End If
      
   End If
   
End Function

Function VBRasSetEntryProperties(strEntryName As String, _
         clsRasEntry As VBRasEntry, _
         Optional strPhoneBook As String) As Long
   
   Dim rtn As Long, lngCb As Long, lngBuffLen As Long
   Dim b() As Byte
   Dim lngPos As Long, lngStrLen As Long
   
   rtn = RasGetEntryProperties(vbNullString, vbNullString, _
                           ByVal 0&, lngCb, ByVal 0&, ByVal 0&)

   If rtn <> 603 Then VBRasSetEntryProperties = rtn: Exit Function
   
   lngStrLen = Len(clsRasEntry.AlternateNumbers)
   lngBuffLen = lngCb + lngStrLen + 1
   ReDim b(lngBuffLen)
   
   CopyMemory b(0), lngCb, 4
   CopyMemory b(4), clsRasEntry.Options, 4
   CopyMemory b(8), clsRasEntry.CountryID, 4
   CopyMemory b(12), clsRasEntry.CountryCode, 4
   CopyStringToByte b(16), clsRasEntry.AreaCode, 11
   CopyStringToByte b(27), clsRasEntry.LocalPhoneNumber, 129
   
   If lngStrLen > 0 Then
     CopyMemory b(lngCb), _
               ByVal clsRasEntry.AlternateNumbers, lngStrLen
     CopyMemory b(156), lngCb, 4
   End If

   CopyMemory b(160), clsRasEntry.ipAddr, 4
   CopyMemory b(164), clsRasEntry.ipAddrDns, 4
   CopyMemory b(168), clsRasEntry.ipAddrDnsAlt, 4
   CopyMemory b(172), clsRasEntry.ipAddrWins, 4
   CopyMemory b(176), clsRasEntry.ipAddrWinsAlt, 4
   CopyMemory b(180), clsRasEntry.FrameSize, 4
   CopyMemory b(184), clsRasEntry.fNetProtocols, 4
   CopyMemory b(188), clsRasEntry.FramingProtocol, 4
   CopyStringToByte b(192), clsRasEntry.ScriptName, 260
   CopyStringToByte b(452), clsRasEntry.AutodialDll, 260
   CopyStringToByte b(712), clsRasEntry.AutodialFunc, 260
   CopyStringToByte b(972), clsRasEntry.DeviceType, 17
      If lngCb = 1672& Then lngStrLen = 33 Else lngStrLen = 129
   CopyStringToByte b(989), clsRasEntry.DeviceName, lngStrLen
      lngPos = 989 + lngStrLen
   CopyStringToByte b(lngPos), clsRasEntry.X25PadType, 33
      lngPos = lngPos + 33
   CopyStringToByte b(lngPos), clsRasEntry.X25Address, 201
      lngPos = lngPos + 201
   CopyStringToByte b(lngPos), clsRasEntry.X25Facilities, 201
      lngPos = lngPos + 201
   CopyStringToByte b(lngPos), clsRasEntry.X25UserData, 201
      lngPos = lngPos + 203
   CopyMemory b(lngPos), clsRasEntry.Channels, 4
   
   If lngCb > 1768 Then 'NT4 Enhancements & Win2000
      CopyMemory b(1768), clsRasEntry.NT4En_SubEntries, 4
      CopyMemory b(1772), clsRasEntry.NT4En_DialMode, 4
      CopyMemory b(1776), clsRasEntry.NT4En_DialExtraPercent, 4
      CopyMemory b(1780), clsRasEntry.NT4En_DialExtraSampleSeconds, 4
      CopyMemory b(1784), clsRasEntry.NT4En_HangUpExtraPercent, 4
      CopyMemory b(1788), clsRasEntry.NT4En_HangUpExtraSampleSeconds, 4
      CopyMemory b(1792), clsRasEntry.NT4En_IdleDisconnectSeconds, 4
      
      If lngCb > 1796 Then ' Win2000
         CopyMemory b(1796), clsRasEntry.Win2000_Type, 4
         CopyMemory b(1800), clsRasEntry.Win2000_EncryptionType, 4
         CopyMemory b(1804), clsRasEntry.Win2000_CustomAuthKey, 4
         CopyMemory b(1808), clsRasEntry.Win2000_guidId(0), 16
         CopyStringToByte b(1824), clsRasEntry.Win2000_CustomDialDll, 260
         CopyMemory b(2084), clsRasEntry.Win2000_VpnStrategy, 4
      End If
      
   End If
   
   rtn = RasSetEntryProperties(strPhoneBook, strEntryName, _
                              b(0), lngCb, ByVal 0&, ByVal 0&)
   
   VBRasSetEntryProperties = rtn

End Function

Public Function BytesToVBRasDialParams(bytesIn() As Byte, _
            udtVBRasDialParamsOUT As VBRasDialParams) As Boolean
   
   Dim iPos As Long, lngLen As Long
   Dim dwSize As Long
   On Error GoTo badBytes
   
   CopyMemory dwSize, bytesIn(0), 4
   
   If dwSize = 816& Then
      lngLen = 21&
   ElseIf dwSize = 1060& Or dwSize = 1052& Then
      lngLen = 257&
   Else
      'unkown size
      Exit Function
   End If
   
   iPos = 4
   
   With udtVBRasDialParamsOUT
      CopyByteToTrimmedString .entryName, bytesIn(iPos), lngLen
      
      iPos = iPos + lngLen: lngLen = 129
      CopyByteToTrimmedString .PhoneNumber, bytesIn(iPos), lngLen
      
      iPos = iPos + lngLen: lngLen = 129
      CopyByteToTrimmedString .CallbackNumber, bytesIn(iPos), lngLen
      
      iPos = iPos + lngLen: lngLen = 257
      CopyByteToTrimmedString .UserName, bytesIn(iPos), lngLen
      
      iPos = iPos + lngLen: lngLen = 257
      CopyByteToTrimmedString .Password, bytesIn(iPos), lngLen
      
      iPos = iPos + lngLen: lngLen = 16
      CopyByteToTrimmedString .Domain, bytesIn(iPos), lngLen
      
      If dwSize > 1052& Then
         CopyMemory .SubEntryIndex, bytesIn(1052), 4&
         CopyMemory .RasDialFunc2CallbackId, bytesIn(1056), 4&
      End If
   End With
   
   BytesToVBRasDialParams = True
   
   Exit Function
   
badBytes:
   'error handling goes here ??
   BytesToVBRasDialParams = False
End Function

Function VBRasDialParamsToBytes( _
            udtVBRasDialParamsIN As VBRasDialParams, _
            bytesOut() As Byte) As Boolean
   
   Dim rtn As Long
   Dim blnPsswrd As Long
   Dim b() As Byte
   Dim bLens As Variant
   Dim dwSize As Long, i As Long
   Dim iPos As Long, lngLen As Long
   
   bLens = Array(1060&, 1052&, 816&)
   For i = 0 To 2
      dwSize = bLens(i)
      ReDim b(dwSize - 1)
      CopyMemory b(0), dwSize, 4
      rtn = RasGetEntryDialParams(vbNullString, b(0), blnPsswrd)
      If rtn = 623& Then Exit For
   Next i
   
   If rtn <> 623& Then Exit Function
   
   On Error GoTo badBytes
   ReDim bytesOut(dwSize - 1)
   CopyMemory bytesOut(0), dwSize, 4
   
   If dwSize = 816& Then
      lngLen = 21&
   ElseIf dwSize = 1060& Or dwSize = 1052& Then
      lngLen = 257&
   Else
      'unkown size
      Exit Function
   End If
   iPos = 4
   With udtVBRasDialParamsIN
      CopyStringToByte bytesOut(iPos), .entryName, lngLen
      iPos = iPos + lngLen: lngLen = 129
      CopyStringToByte bytesOut(iPos), .PhoneNumber, lngLen
      iPos = iPos + lngLen: lngLen = 129
      CopyStringToByte bytesOut(iPos), .CallbackNumber, lngLen
      iPos = iPos + lngLen: lngLen = 257
      CopyStringToByte bytesOut(iPos), .UserName, lngLen
      iPos = iPos + lngLen: lngLen = 257
      CopyStringToByte bytesOut(iPos), .Password, lngLen
      iPos = iPos + lngLen: lngLen = 16
      CopyStringToByte bytesOut(iPos), .Domain, lngLen
      
      If dwSize > 1052& Then
         CopyMemory bytesOut(1052), .SubEntryIndex, 4&
         CopyMemory bytesOut(1056), .RasDialFunc2CallbackId, 4&
      End If
   End With
   VBRasDialParamsToBytes = True
   Exit Function
badBytes:
   'error handling goes here ??
   VBRasDialParamsToBytes = False
End Function

Public Sub CopyByteToTrimmedString(strToCopyTo As String, _
                              bPos As Byte, lngMaxLen As Long)
   Dim strTemp As String, lngLen As Long
   strTemp = String(lngMaxLen + 1, 0)
   CopyMemory ByVal strTemp, bPos, lngMaxLen
   lngLen = InStr(strTemp, Chr$(0)) - 1
   strToCopyTo = Left$(strTemp, lngLen)
End Sub

Public Sub CopyStringToByte(bPos As Byte, _
                        strToCopy As String, lngMaxLen As Long)
   Dim lngLen As Long
   lngLen = Len(strToCopy)
   If lngLen = 0 Then
      Exit Sub
   ElseIf lngLen > lngMaxLen Then
      lngLen = lngMaxLen
   End If
   CopyMemory bPos, ByVal strToCopy, lngLen
End Sub

Public Function VBRasGetEntryDialParams _
              (bytesOut() As Byte, _
          strPhoneBook As String, strEntryName As String, _
               Optional blnPasswordRetrieved As Boolean) As Long
   
   Dim rtn As Long
   Dim blnPsswrd As Long
   Dim bLens As Variant
   Dim lngLen As Long, i As Long
   
   bLens = Array(1060&, 1052&, 816&)
   
   'try our three different sizes for RasDialParams
   For i = 0 To 2
      lngLen = bLens(i)
      ReDim bytesOut(lngLen - 1)
      CopyMemory bytesOut(0), lngLen, 4
      If lngLen = 816& Then
         CopyStringToByte bytesOut(4), strEntryName, 20
      Else
         CopyStringToByte bytesOut(4), strEntryName, 256
      End If
      rtn = RasGetEntryDialParams(strPhoneBook, bytesOut(0), blnPsswrd)
      If rtn = 0 Then Exit For
   Next i
   
   blnPasswordRetrieved = blnPsswrd
   VBRasGetEntryDialParams = rtn
End Function


Public Function GetRasEntries() As String()

   Dim strPhoneBook As String
   Dim rtn As Long, i As Long
   Dim lpcB As Long                             'count of bytes
   Dim lpCEntries As Long                       'count of entries
   Dim b() As Byte
   Dim strTemp As String
   Dim dwSize As Long                           'size of each entry
   Dim lngLen As Long
   Dim lngBLen As Variant
   Dim clsRasEntryName() As VBRasEntryName
   Dim Ret() As String
   
   ReDim b(3)
   
   strPhoneBook = vbNullString
   'determine appropiate size for b()
   
   lngBLen = Array(532&, 264&, 28&)
   For i = 0 To 2
      CopyMemory b(0), CLng(lngBLen(i)), 4
      rtn = RasEnumEntries(vbNullString, strPhoneBook, _
             b(0), lpcB, lpCEntries)
      If rtn <> 632 Then Exit For
   Next i

   If lpCEntries = 0 Then Exit Function
   
   dwSize = lpcB \ lpCEntries
   
   ReDim b(lpcB - 1)
   
   CopyMemory b(0), dwSize, 4
   
   rtn = RasEnumEntries(vbNullString, strPhoneBook, _
             b(0), lpcB, lpCEntries)
   
   If rtn <> 0 Then Enum_List = Ret() 'MsgBox VBRASERRorHandler(rtn)
   
   strTemp = String(dwSize - 4, 0)
   
   ReDim clsRasEntryName(lpCEntries - 1)
   ReDim Ret(lpCEntries - 1)
   
   If dwSize = 28 Then lngLen = 21 Else lngLen = 257
   For i = 0 To lpCEntries - 1
     CopyMemory ByVal strTemp, b((i * dwSize) + 4), lngLen
     clsRasEntryName(i).entryName = _
            Left(strTemp, InStr(strTemp, Chr$(0)) - 1)
    Ret(i) = clsRasEntryName(i).entryName
   Next i
   
   GetRasEntries = Ret
End Function

Public Function GetEntryDialParams(entryName As String, RasDialParams As VBRasDialParams) As Long

    Dim Ret As Long, b() As Byte
    Dim Ret1 As Boolean
    Dim blnPasswordRetrieved As Boolean
    
    Ret = VBRasGetEntryDialParams(b, vbNullString, entryName, blnPasswordRetrieved)
    
    If Not Ret = 0 Then
        GetEntryDialParams = Ret
        Exit Function
    End If

    'Debug.Print "Password Retrieved: "; blnPasswordRetrieved
    'Note: the phone number of the connection is only returned on Windows NT4 enhanced and windows 2000.
    Ret1 = BytesToVBRasDialParams(b(), RasDialParams)

End Function

'Function VBRasGetEntryProperties(strEntryName As String, _
         clsRasEntry As VBRasEntry, _
         Optional strPhoneBook As String) As Long
            
Public Function GetRasEntryProperties(strEntryName As String, _
         RasEntry As VBRasEntry, _
         Optional strPhoneBook As String) As Long
    GetRasEntryProperties = VBRasGetEntryProperties(strEntryName, RasEntry, strPhoneBook)
    
End Function

'The following sample does a simple synchronous dial.
'The function returns the handle to the connection.
'Remember you have to eventually Hangup the connection if the connection's
'handle value is non-zero .  This function also use the VBRasGetEntryDialParams
'function (from the RasDialParams page).
'You can call it like so:
'Dim hConn As Long
'hConn = VBSyncronousDial(vbNullString, "My Connection")
Public Function VBSyncronousDial(strPhoneBook As String, _
                              strEntryName As String) As Long
   Dim Ret As Long
   Dim b() As Byte
   Dim lngHConn As Long
   
   Ret = VBRasGetEntryDialParams(b, strPhoneBook, strEntryName)
   
   'todo: check if rtn = 0 else handle error
   
   ' note how code pauses on next line
   ' until connection established or fails
    If Not Ret = 0 Then
        VBSyncronousDial = Ret
        Exit Function
    End If

   Ret = RasDial(ByVal 0&, strPhoneBook, b(0), 0&, 0&, lngHConn)
   
   'todo: check if rtn = 0 else handle error
   
   VBSyncronousDial = lngHConn
End Function

'Simple Asyncronous Dial
'The following sample does a simple asynchronous dial.  The function returns
'the handle to the connection.  Remember you have to eventually Hangup the
'connection if the connection's handle value is non-zero .  This function also
'use the VBRasGetEntryDialParams function (from the RasDialParams page).
'You can call it like so:
'Dim hConn As Long
'hConn = VBAsyncronousDial(vbNullString, "My Connection")


Public Function VBAsyncronousDial(strPhoneBook As String, _
                              strEntryName As String) As Long
   Dim rtn As Long
   Dim b() As Byte
   Dim lngHConn As Long
   
   rtn = VBRasGetEntryDialParams(b, strPhoneBook, strEntryName)
   
   'todo: check if rtn = 0 else handle error
   
   rtn = RasDial(ByVal 0&, strPhoneBook, b(0), _
                     1&, AddressOf RasDialFunc1, lngHConn)
   
   'todo: check if rtn = 0 else handle error
   
   VBAsyncronousDial = lngHConn
End Function

Private Sub RasDialFunc1(ByVal hRasConn As Long, ByVal unMsg As Long, _
       ByVal rasConnectionState As Long, ByVal dwError As Long, _
       ByVal dwExtendedError As Long)
   
    ' do stuff here, such as check for errors
    ' or check the rasConnectionState
    'Debug.Print hRasConn, Hex$(rasConnectionState), dwError
    If rasConnectionState = RASCS_Connected Then
    '  Debug.Print "connected"
    ElseIf rasConnectionState = RASCS_Disconnected Then
    '  Debug.Print "disconnected"
    End If
    
    If dwError <> 0 Then
        Debug.Print VBRASERRorHandler(dwError)
    End If

End Sub

'You can call this function from VB like so: (note: you need to already
'have the handle to an existing connection, hRasConn)
'Dim rtn As Long
'Dim myConnStatus As VBRASCONNSTATUS
'rtn = VBRasGetConnectStatus(hRasConn, myConnStatus)
'If rtn <> 0 Then MsgBox "failed"
Public Function VBRasGetConnectStatus _
               (hRasConn As Long, _
               udtVBRasConnStatus As VBRASCONNSTATUS) As Long

   Dim i As Long, dwSize As Long
   Dim aVarLens As Variant
   Dim b() As Byte

   aVarLens = Array(288&, 160&, 64&)
   
   For i = 0 To 2
      dwSize = aVarLens(i)
      ReDim b(dwSize - 1)
      CopyMemory b(0), dwSize, 4
      rtn = RasGetConnectStatus(hRasConn, b(0))
      If rtn <> 632 Then Exit For
   Next i
   
   VBRasGetConnectStatus = rtn
   If rtn <> 0 Then Exit Function
      
   With udtVBRasConnStatus
      CopyMemory .lRasConnState, b(4), 4
      CopyMemory .dwError, b(8), 4
      CopyByteToTrimmedString .sDeviceType, b(12), 17&
      If dwSize = 64& Then
         CopyByteToTrimmedString .sDeviceName, b(29), 33&
      ElseIf dwSize = 160& Then
         CopyByteToTrimmedString .sDeviceName, b(29), 129&
      Else
         CopyByteToTrimmedString .sDeviceName, b(29), 129&
         CopyByteToTrimmedString .sNTPhoneNumber, b(158), 129&
      End If
   End With
   
End Function

'The RASGetCountryInfo function retrieves the country code and name
'of a country.  You need to specify the countryID .  The function also returns
'the countryID of the next country as stored in the registry.
'Dim MyCountry As VBRASCTRYINFO
'MyCountry.CountryID = 61
'rtn = VBRasGetCountryInfo(MyCountry)
'You can also enumerate all the countries, by starting with CountryID =1, then calling the VBRasGetCountryInfo again, substituting the CountryID with the previously returned NextCountryID.
'This sample shows how to call on the above function to enumerate all countries and print their code and name to the Debug window in VB.
'Dim lngID As Long, rtn As Long
'Dim MyCountry As VBRASCTRYINFO
'lngID = 1
'Do
'   MyCountry.CountryID = lngID
'   rtn = VBRasGetCountryInfo(MyCountry)
'
'   With MyCountry
'        Debug.Print .CountryCode, .CountryName
'    End With
'
'    If rtn <> 0 Then Exit Do
'    lngID = MyCountry.NextCountryID
'Loop While lngID <> 0
Public Function VBRasGetCountryInfo(clsCountryInfo As VBRASCTRYINFO) As Long
   
   Dim b(511) As Byte, lpSize As Long, rtn As Long
   Dim lPos As Long, strTemp As String, lngLen As Long
   b(0) = 20
   CopyMemory b(4), clsCountryInfo.CountryID, 4
   lpSize = 512
   
   rtn = RasGetCountryInfo(b(0), lpSize)
   
   VBRasGetCountryInfo = rtn
   If rtn <> 0 Then Exit Function
   
   CopyMemory clsCountryInfo.NextCountryID, b(8), 4
   CopyMemory clsCountryInfo.CountryCode, b(12), 4
   
   CopyMemory lPos, b(16), 4
   lngLen = lpSize - lPos - 2
   
   If lngLen > 0 Then
      strTemp = String(lngLen, 0)
      CopyMemory ByVal strTemp, b(lPos), lngLen
   End If
   
   clsCountryInfo.CountryName = strTemp
End Function

Public Function VBRasEnumDevices(clsVBRasDevInfo() As VBRASDEVINFO) As Long
   
   Dim rtn As Long, i As Long
   Dim lpcB As Long, lpCDevices As Long
   Dim b() As Byte
   Dim dwSize As Long
   
   rtn = RasEnumDevices(ByVal 0&, lpcB, lpCDevices)

   If lpCDevices = 0 Then Exit Function
   
   dwSize = lpcB \ lpCDevices
   
   ReDim b(lpcB - 1)
   
   CopyMemory b(0), dwSize, 4
   
   rtn = RasEnumDevices(b(0), lpcB, lpCDevices)
   
   If lpCDevices = 0 Then Exit Function
   
   ReDim clsVBRasDevInfo(lpCDevices - 1)
   
   For i = 0 To lpCDevices - 1
     CopyByteToTrimmedString clsVBRasDevInfo(i).DeviceType, _
                                    b((i * dwSize) + 4), 17
     CopyByteToTrimmedString clsVBRasDevInfo(i).DeviceName, _
                           b((i * dwSize) + 21), dwSize - 21
   Next i
   
   VBRasEnumDevices = lpCDevices

End Function

Public Function VBRASERRorHandler(rtn As Long) As String
   Dim strp_sERRor As String, i As Long
   strp_sERRor = String(512, 0)
   If rtn > 600 Then
      RasGetErrorString rtn, strp_sERRor, 512&
   Else
      FormatMessage &H1000, ByVal 0&, rtn, 0&, strp_sERRor, 512, ByVal 0&
   End If
   i = InStr(strp_sERRor, Chr$(0))
   If i > 1 Then VBRASERRorHandler = Left$(strError, i - 1)
End Function

Public Function RenameRasEntry(oldName As String, newName As String) As Long
    Dim rtn As Long
    rtn = RasRenameEntry(vbNullString, oldName, newName)
    'If successful it will return 0.
    RenameRasEntry = rtn
    If rtn <> 0 Then
       'MsgBox VBRASERRorHandler(rtn)
    End If

End Function

Public Function DeleteRasEntry(entryName As String) As Long

    Dim rtn As Long
    rtn = RasDeleteEntry(vbNullString, entryName)
    'If successful it will return 0.
    DeleteRasEntry = rtn
    If rtn <> 0 Then
       'MsgBox VBRASERRorHandler(rtn)
    End If

End Function

'If the name is valid and does not already exist it will return 0.
'If the entry name already exists it will return 183.  You could use this to check for an existing entry rather than using the RASEnumEntries.
'If the name syntax is invalid it will return 123
'The following is an example:
Public Function ValidateRasEntryName(entryName As String) As Long
    Dim rtn As Long
    rtn = RasValidateEntryName(vbNullString, entryName)
    ValidateRasEntryName = rtn
    If rtn <> 0 Then
       'Debug.Print rtn, VBRASERRorHandler(rtn)
    End If
End Function


'The following code shows how to enumerate connections and populate an
'array of VBRASCONN structures.  The function returns a long equal to
'the number of connections.
'You could call this function like so:
'Dim nConnections As Long
'Dim myConnections() As VBRASCONN
'nConnections = VBRasEnumConnections(myConnections)
Function VBRasEnumConnections(aVBRasConns() As VBRASCONN) As Long

   Dim rtn As Long
   Dim b() As Byte
   Dim aLens As Variant, dwSize As Long
   Dim lpcB As Long, lpConns As Long
   Dim i As Long
   
   ReDim b(3)
   aLens = Array(692&, 676&, 412&, 32&)

   For i = 0 To 3
      dwSize = aLens(i)
      CopyMemory b(0), dwSize, 4
      lpcB = 4
      rtn = RasEnumConnections(b(0), lpcB, lpConns)
      If rtn <> 632 And rtn <> 610 Then Exit For
   Next i
   
   VBRasEnumConnections = lpConns
   If lpConns = 0 Then Exit Function
   
   lpcB = dwSize * lpConns
   ReDim b(lpcB - 1)
   CopyMemory b(0), dwSize, 4
   rtn = RasEnumConnections(b(0), lpcB, lpConns)
   
   If lpConns = 0 Then
        VBRasEnumConnections = 0
        Exit Function
   End If
   
   ' now copy the bytes to the aVBRasConns array
   ReDim aVBRasConns(lpConns - 1)
   For i = 0 To lpConns - 1
      With aVBRasConns(i)
         CopyMemory .hRasConn, b(i * dwSize + 4), 4
         If dwSize = 32& Then
            CopyByteToTrimmedString .sEntryName, b(i * dwSize + 8), 21&
         Else
            CopyByteToTrimmedString .sEntryName, b(i * dwSize + 8), 257&
            CopyByteToTrimmedString .sDeviceType, b(i * dwSize + 265), 17&
            CopyByteToTrimmedString .sDeviceName, b(i * dwSize + 282), 129&
            If dwSize > 412& Then
              CopyByteToTrimmedString .sPhonebook, b(i * dwSize + 411), 260&
              CopyMemory .lngSubEntry, b(i * dwSize + 672), 4
              If dwSize > 676& Then
                CopyMemory .guidEntry(0), b(i * dwSize + 676), 16
              End If
            End If
         End If
      End With
   Next i
End Function

'Checks online status: Returns True if Online and False if Offline
Public Function GetConnectionStatus() As Boolean
    
    Dim nConnections As Long, Ret As Long
    Dim myConnections() As VBRASCONN
    
    nConnections = VBRasEnumConnections(myConnections)
    
    If nConnections = 0 Then
        GetConnectionStatus = False
        Exit Function
    End If
    
    Dim ConnStatus As VBRASCONNSTATUS
    
    Dim X As Long
    
    X = 0
    
    Ret = VBRasGetConnectStatus(myConnections(0).hRasConn, ConnStatus)
    
    If ConnStatus.lRasConnState = RASCS_Connected Then ' = &H2000
        'Online
        GetConnectionStatus = True
    Else
        'Offline
        GetConnectionStatus = False
    End If
    
End Function

'Note:  The RasHangUp function returns immediately, however it may take some
'time for Ras to actually disconnect and release resources.  It's up to yourself
'as to how you decide to handle this, but you should wait for Ras to complete
'it's hangup before closing your app or calling the RasDial or RasHangUp again.  Eg:
Public Function HangUpRas(hRasConn As Long) As Long
    
    Dim rtn As Long, lngError As Long
    Dim myConnStatus As VBRASCONNSTATUS
    
    rtn = RasHangUp(hRasConn)
    
    'If rtn <> 0 Then Debug.Print VBRasErrorHandler(rtn)
    
    Do
        Sleep 0&
        lngError = VBRasGetConnectStatus(hRasConn, myConnStatus)
    Loop While lngError <> ERROR_INVALID_HANDLE

End Function

'The HangUpRasAsync function returns immediately
Public Function HangUpRasAsync(hRasConn As Long) As Long
    HangUpRasAsync = RasHangUp(hRasConn)
    'If rtn <> 0 Then Debug.Print VBRasErrorHandler(rtn)
End Function

'Hand up the first connected RAS found no need for RAS handle
Public Function HangUpConnectedRas() As Boolean
    
    Dim nConnections As Long, Ret As Long
    Dim myConnections() As VBRASCONN
    
    nConnections = VBRasEnumConnections(myConnections)
    
    If nConnections = 0 Then
        HangUpConnectedRas = False
        Exit Function
    End If
    
    Dim ConnStatus As VBRASCONNSTATUS
    
    Dim X As Long
    
    X = 0
    
    Ret = VBRasGetConnectStatus(myConnections(0).hRasConn, ConnStatus)
    
    If ConnStatus.lRasConnState = RASCS_Connected Then ' = &H2000
        'Online
        HangUpConnectedRas = RasHangUp(myConnections(0).hRasConn)
    Else
        'Offline
        HangUpConnectedRas = False
    End If
    
End Function


