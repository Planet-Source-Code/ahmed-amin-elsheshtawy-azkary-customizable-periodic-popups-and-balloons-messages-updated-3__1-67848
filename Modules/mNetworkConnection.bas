Attribute VB_Name = "mNetworkConnection"
Option Explicit
'-----------------------------------------------------------------------------------------
' Copyright ©1996-2006 VBnet, Randy Birch. All Rights Reserved Worldwide.
'        Terms of use http://vbnet.mvps.org/terms/pages/terms.htm
'-----------------------------------------------------------------------------------------
Private Const NETWORK_ALIVE_LAN = &H1  'net card connection
Private Const NETWORK_ALIVE_WAN = &H2  'RAS connection
Private Const NETWORK_ALIVE_AOL = &H4  'AOL
       
Private Type QOCINFO
    dwSize As Long
    dwFlags As Long
    dwInSpeed As Long 'in bytes/second
    dwOutSpeed As Long 'in bytes/second
End Type
       
Private Declare Function IsNetworkAlive Lib "SENSAPI.DLL" (lpdwFlags As Long) As Long
Private Declare Function IsDestinationReachable Lib "SENSAPI.DLL" Alias _
    "IsDestinationReachableA" (ByVal lpszDestination As String, _
    ByRef lpQOCInfo As QOCINFO) As Long

Public Function IsNetConnectionAlive() As Boolean

   Dim lpdwFlags As Long
   IsNetConnectionAlive = IsNetworkAlive(lpdwFlags) = 1
   
End Function

Public Function IsNetConnectionLAN() As Boolean

   Dim lpdwFlags As Long
   
   If IsNetworkAlive(lpdwFlags) = 1 Then
      IsNetConnectionLAN = lpdwFlags = NETWORK_ALIVE_LAN
   End If

End Function

Public Function IsNetConnectionRAS() As Boolean

   Dim lpdwFlags As Long
   
   If IsNetworkAlive(lpdwFlags) = 1 Then
      IsNetConnectionRAS = lpdwFlags = NETWORK_ALIVE_WAN
   End If
   
End Function

Public Function IsNetConnectionAOL() As Boolean

   Dim lpdwFlags As Long
   
   If IsNetworkAlive(lpdwFlags) = 1 Then
      IsNetConnectionAOL = lpdwFlags = NETWORK_ALIVE_AOL
   End If

End Function

Public Function GetNetConnectionType() As String

   Dim lpdwFlags As Long

   If IsNetworkAlive(lpdwFlags) = 1 Then
      Select Case lpdwFlags
         Case NETWORK_ALIVE_LAN:
            GetNetConnectionType = _
             "The system has one or more active LAN cards"
         Case NETWORK_ALIVE_WAN:
            GetNetConnectionType = _
             "The system has one or more active RAS connections"
         Case NETWORK_ALIVE_AOL:
            GetNetConnectionType = _
             "The system is connected to America Online"
         Case Else
      End Select
      
   Else
      GetNetConnectionType = _
       "The system has no connection or an error occurred"
   End If
   
End Function

Public Function DestinationReachable(ByVal Destination As String) As Boolean

    Dim Ret As QOCINFO
    Ret.dwSize = Len(Ret)
    If IsDestinationReachable(Destination, Ret) = 0 Then
        DestinationReachable = False
    Else
        DestinationReachable = True
    End If
    
End Function



