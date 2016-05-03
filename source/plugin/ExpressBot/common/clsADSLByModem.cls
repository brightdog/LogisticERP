VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADSLByModem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Dim mstrCurrentIP As String
Dim dicConfig As Scripting.Dictionary
'枚举网络连接

'=================================================
Private Const RAS_MaxDeviceType = 16
Private Const RAS95_MaxDeviceName = 128
Private Const RAS95_MaxEntryName = 256
Private Type RASCONN95
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
Private Type RASENTRYNAME95
    dwSize As Long
    szEntryName(RAS95_MaxEntryName) As Byte
End Type
Private Declare Function RasEnumConnections Lib "RasApi32.DLL" Alias "RasEnumConnectionsA" (lprasconn As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasEnumEntries Lib "RasApi32.DLL" Alias "RasEnumEntriesA" (ByVal reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long
Private Declare Function RasHangUp Lib "RasApi32.DLL" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
  
 
'枚举网络连接名称s
Private Function GetAdslName() As String
    Dim s As Long, l As Long, ln As Long, a As String
    ReDim r(255) As RASENTRYNAME95
     
    r(0).dwSize = 264
    s = 256 * r(0).dwSize
    l = RasEnumEntries(vbNullString, vbNullString, r(0), s, ln)
    For l = 0 To ln - 1
        a = StrConv(r(l).szEntryName(), vbUnicode)
        GetAdslName = Left$(a$, InStr(a$, Chr$(0)) - 1)
        Exit Function
    Next
End Function

Public Function Connect() As Boolean

        'If Not modCheckisConnected.IsConnected() Then
100     Call modDosPrint.DosPrint("rasdial " & GetAdslName() & " " & dicConfig.Item("txtADSL_UserName") & " " & dicConfig.Item("txtADSL_Password"), False, False)
        WriteLog "$$rasdial " & GetAdslName() & " " & dicConfig.Item("txtADSL_UserName") & " " & dicConfig.Item("txtADSL_Password")
        Dim sngNow As Single
102     sngNow = VBA.Timer
            
        Dim dtDate As Date
104     dtDate = VBA.Date
            
        Do
            
106         MySleep 2
                
108         If checkADSLConnected Then
110             Connect = True

                Exit Function
            End If
            
112     Loop While Timer - sngNow < 25 And VBA.Date = dtDate

114     Connect = False

End Function

Private Function checkADSLConnected() As Boolean

        Dim obj As clsGetIP
100     Set obj = New clsGetIP
    
102     'WriteLog "Checking ADSL is Connected..."
    
        Dim strInternalIP As String
        Dim strOuternalIP As String
        
104     strInternalIP = obj.GetCurrentIP(False)
106     strOuternalIP = obj.GetCurrentIP(True)
        WriteLog "*NEW IP: " & strOuternalIP
108     If strInternalIP <> strOuternalIP And strInternalIP <> "" Then

110         checkADSLConnected = True
        Else
112         checkADSLConnected = False
        End If


116     Set obj = Nothing
End Function

Public Function Disconnect() As Boolean
        'rasdial /disconnect
        Dim sngNow As Single
102     sngNow = VBA.Timer
            
        Dim dtDate As Date
104     dtDate = VBA.Date
            
        Do
100         Call modDosPrint.DosPrint("rasdial " & GetAdslName() & " " & "/DISCONNECT", True, True)
            WriteLog "$$rasdial " & GetAdslName() & " " & "/DISCONNECT"
106         MySleep 1
                
108         If Not checkADSLConnected Then
110             Disconnect = True

                Exit Function
            Else
                WriteLog "断线失败，重新断一次！"
            End If
            
112     Loop While Timer - sngNow < 25 And VBA.Date = dtDate

114     Disconnect = False
End Function


Public Function CheckCurrentState() As Boolean


    CheckCurrentState = True

End Function





Private Sub ReadConfig()
        '<EhHeader>
        On Error GoTo ReadConfig_Err
        '</EhHeader>

        Dim iFile As Integer
    
100     iFile = FreeFile()

102     Open App.path & "\ADSLConfig.Cfg" For Input As #iFile
    
        'MsgBox App.Path
    
        Dim i As Integer
104     i = 1
    
        Dim clsEcypt As clsEncrypt
106     Set clsEcypt = New clsEncrypt
    
108     Do While Not EOF(1)
            Dim strTmp As String
            Dim arr() As String
110         Line Input #iFile, strTmp
112         arr = Split(clsEcypt.Decode(strTmp), "|-|", 2, vbBinaryCompare)

114         If UBound(arr) = 1 Then
            
                dicConfig.Add arr(0), restoreCRLF(arr(1))
        
118         Else
        
            End If
        
120         i = i + 1
        Loop

        Close #iFile

        '<EhFooter>
        Exit Sub

ReadConfig_Err:
        
        '</EhFooter>
End Sub




Private Sub Class_Initialize()
    Set dicConfig = New Scripting.Dictionary
    Call ReadConfig
End Sub




