VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&

Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Declare Function WSAStartup _
                Lib "WSOCK32.DLL" (ByVal wVersionRequired&, _
                                   lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname _
                Lib "WSOCK32.DLL" (ByVal hostname$, _
                                   ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory _
                Lib "kernel32" (hpvDest As Any, _
                                ByVal hpvSource&, _
                                ByVal cbCopy&)

Private Function hibyte(ByVal wParam As Integer)
100     hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Integer)
100     lobyte = wParam And &HFF&
End Function

Private Sub SocketsInitialize()
        Dim WSAD     As WSADATA
        Dim iReturn  As Integer
        Dim sLowByte As String, sHighByte As String, sMsg As String

100     iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

102     If iReturn = 0 Then
104         If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
106             sHighByte = Trim$(str$(hibyte(WSAD.wversion)))
108             sLowByte = Trim$(str$(lobyte(WSAD.wversion)))
110             sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
                'Debug.Print sMsg
                'sMsg = sMsg & " winsock.dll tarafindan desteklenmiyor. "
                'MsgBox sMsg
                'End
            End If

        Else
            'Debug.Print "Winsock.dll Error."
        End If

End Sub

Public Function GetCurrentIP(ByVal blnExternalIP As Boolean) As String

        Dim hostname          As String * 256
        Dim hostent_addr      As Long
        Dim host              As HOSTENT
        Dim hostip_addr       As Long
        Dim temp_ip_address() As Byte
        Dim i                 As Integer
        Dim ip_address        As String
        Dim IP                As String
        Dim Internal          As String
        Dim EXTERNAL          As String

100     If gethostname(hostname, 256) <> SOCKET_ERROR Then
102         hostname = Trim$(hostname)

104         hostent_addr = gethostbyname(hostname)

106         If hostent_addr <> 0 Then
108             RtlMoveMemory host, hostent_addr, LenB(host)
110             RtlMoveMemory hostip_addr, host.hAddrList, 4

                Do
112                 ReDim temp_ip_address(1 To host.hLength)
114                 RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

116                 For i = 1 To host.hLength
118                     ip_address = ip_address & temp_ip_address(i) & "."
                    Next

120                 ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

                    ' Return Both LAN and External IP Fix
                    ' Master Yoda 30-05-2000
                    ' ##########################################
                    ' HERE'S THE PROBLEM!!!
                    'TheIP = TheIP + ip_address
                    ' ##########################################
                    ' HERE'S THE FIX!!!
122                 Internal = IP ' Send ONLY the External IP to the CurrentIP Function
124                 EXTERNAL = ip_address ' Send the External IP to the function parameter External
126                 IP = ip_address ' Send LAN IP to the function para Internal

                    ' You don't really need to return parameters,
                    ' it just allows you to get both IPs :)
                    ' ##########################################

128                 ip_address = ""
130                 host.hAddrList = host.hAddrList + LenB(host.hAddrList)
132                 RtlMoveMemory hostip_addr, host.hAddrList, 4
134             Loop While (hostip_addr <> 0)

136             If blnExternalIP = True Then
138                 GetCurrentIP = EXTERNAL
                Else
140                 GetCurrentIP = Internal
                End If

            Else
                'Debug.Print "Winsock.dll error."

142             GetCurrentIP = ""
            End If

        Else
            'Debug.Print "Windows Socket Error " & Str(WSAGetLastError())

144         GetCurrentIP = ""
        End If

End Function

Private Sub SocketsCleanup()

        Dim lReturn As Long

100     lReturn = WSACleanup()

102     If lReturn <> 0 Then
            'MsgBox "Socket Error " & Trim$(Str$(lReturn)) & " occurred In Cleanup "
        End If

End Sub

Private Sub Class_Initialize()

100     SocketsInitialize

End Sub

Private Sub Class_Terminate()

    SocketsCleanup

End Sub

