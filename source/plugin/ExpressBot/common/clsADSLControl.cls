VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADSLByRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mbolConnectionState As Boolean
Dim mbolExecuteFinish As Boolean
Dim mstrCurrentIP As String
Dim dicConfig As Scripting.Dictionary

Public Function Connect() As Boolean

    Dim strMethod As String
    Dim strPostData As String
    Dim strHeader As String
    Dim strUrl As String
    
    'Call SaveConfig
    
    mbolExecuteFinish = False
    Connect = False
    dicConfig.Item("txtLog") = ""
    
    Dim iWeb As clsXMLHTTPGetHtml
    Set iWeb = New clsXMLHTTPGetHtml
    
    strUrl = MakeURL

    If strUrl <> "" Then
    
        strUrl = strUrl & Trim(dicConfig.Item("txtURL_Connect"))
        strPostData = dicConfig.Item("txtPost_Connect")
        
        If strPostData <> "" Then
            iWeb.PostData = strPostData

        End If
        
        iWeb.URL = strUrl
        Call iWeb.Send

        dicConfig.Item("txtLog") = iWeb.ReturnData
        
        mbolConnectionState = CheckConnectState

        If mbolConnectionState Then
            Connect = True
            'MsgBox "连接成功"
        End If

    Else
    
    End If

    '        MyInet.AccessType = icDirect
    '        dicConfig.item("txtLog") = MyInet.OpenURL(dicConfig.item("txtIP_Connect"))
        
End Function

Public Function Disconnect() As Boolean
    
    Dim strMethod As String
    Dim strPostData As String
    Dim strHeader As String
    Dim strUrl As String
    
    'Call SaveConfig
    mbolExecuteFinish = False
    Disconnect = False
    dicConfig.Item("txtLog") = ""
    
    Dim iWeb As clsXMLHTTPGetHtml
    Set iWeb = New clsXMLHTTPGetHtml
    
    strUrl = MakeURL

    If strUrl <> "" Then
        
        strUrl = strUrl & Trim(dicConfig.Item("txtURL_DisConnect"))
        
        strPostData = dicConfig.Item("txtPost_DisConnect")
        
        If strPostData <> "" Then
            iWeb.PostData = strPostData

        End If
        
        iWeb.URL = strUrl
        Call iWeb.Send

        dicConfig.Item("txtLog") = iWeb.ReturnData
        
        mbolConnectionState = CheckConnectState

        If Not mbolConnectionState Then
            Disconnect = True
            'MsgBox "断开成功"
        End If
        
    Else
    
    End If
    
    'dicConfig.item("txtLog") = MyInet.OpenURL(dicConfig.item("txtIP_DisConnect"))

End Function

Public Function CheckCurrentState() As Boolean

    dicConfig.Item("txtLog") = ""
    
    Dim iWeb As clsXMLHTTPGetHtml
    Set iWeb = New clsXMLHTTPGetHtml
    Dim strUrl As String
    Dim strMethod As String
    Dim strHeader As String
    
    strUrl = MakeURL

    If strUrl <> "" Then
        
        strUrl = strUrl & Trim(dicConfig.Item("txtURL_CheckState"))
        
        iWeb.URL = strUrl
        Call iWeb.Send

        dicConfig.Item("txtLog") = iWeb.ReturnData
        mstrCurrentIP = ""
        Call CheckConnectState
        'MsgBox "当前IP地址为:" & mstrCurrentIP

    Else
    
    End If

End Function

Private Function CheckConnectState() As Boolean
    'TRUE表示连接上
    'FALSE表示断开了连接
    
    Dim iMaxWaitSecond As Integer
    iMaxWaitSecond = 15 '最多检测15秒！超时就判失败！以免程序陷入死循环！2014-04-17
    Dim StartTime As String
    StartTime = Now '纪录检测开始的时间，作为当次比对的依据
    
    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    
    Reg.IgnoreCase = True
    Reg.MultiLine = False
    Reg.Global = False
    Reg.Pattern = "var wanPara = new Array\((.*?);"
    
    Dim Mc        As MatchCollection
    Dim m         As Match
    Dim strResult As String
    strResult = Split(dicConfig.Item("txtLog"), vbCrLf, -1, vbBinaryCompare)(0)
    Set Mc = Reg.Execute(convertCRLF(strResult))

    If Mc.Count = 1 Then
    
        Set m = Mc.Item(0)
        
        Dim arr() As String
        
        arr = Split(restoreCRLF(m.SubMatches(0)), ",", -1, vbBinaryCompare)
        
        If arr(14) = 0 Then
            
            CheckConnectState = False

        Else
            
            If arr(14) = 2 Then
                If DateDiff("s", StartTime, Now) < iMaxWaitSecond Then
                    CheckConnectState = True
                    MySleep 3
                    Call CheckCurrentState
                Else
                    CheckConnectState = False
                End If

            ElseIf arr(14) = 1 Then
                
                mstrCurrentIP = arr(2)
            Else
                
            End If
        End If
    
    End If

End Function



Private Function MakeURL() As String

    Dim strUrl As String
    
    If dicConfig.Item("txtUserName") = "" Then
    
        Exit Function
    End If
    
    If dicConfig.Item("txtPassWord") = "" Then
    
        Exit Function
        
    End If
    
    strUrl = "http://" & Trim(dicConfig.Item("txtUserName")) & ":" & Trim(dicConfig.Item("txtPassWord")) & "@" & dicConfig.Item("txtIP")
    
    If dicConfig.Item("txtPort") <> "" Then
    
        If IsNumeric(dicConfig.Item("txtPort")) Then
        
            If Val(dicConfig.Item("txtPort")) > 0 Then
                
                strUrl = strUrl & ":" & dicConfig.Item("txtPort")
            
            Else
            
                
                Exit Function
            End If
        
        Else
            
            Exit Function
        End If
    
    End If
    
    'strURL = strURL & "/"
    
    MakeURL = strUrl
    
End Function



Private Sub ReadConfig()
        '<EhHeader>
        On Error GoTo ReadConfig_Err
        '</EhHeader>

        Dim iFile As Integer
    
100     iFile = FreeFile()

102     Open App.Path & "\ADSLConfig.Cfg" For Input As #iFile
    
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

Public Property Get CurrentIP() As String
CurrentIP = mstrCurrentIP
End Property


