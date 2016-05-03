VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'100     Call SendMailCDOCacheConf("wulf@ctrip.com", "", "” º˛Ã·abc.,;:–—≤‚ ‘123", "” º˛Ã·–—≤‚ ‘” asdfasdº˛Ã·–—,:;.'""≤‚ ‘123", "Monitor_test@ctrip.com")
'        Dim objMail As CDONTS.NewMail
'100     Set objMail = New CDONTS.NewMail
'102     objMail.AttachFile App.Path & "\frmMain.frm"
'104     objMail.send "BOT_Hotel@Ctrip_BI.com   ", "wulf@ctrip.com", "Title", "Hello"
'106     Set objMail = Nothing
'
'Public Function SendMailCDOCacheConf(ByVal aFrom As String, ByVal aTo As String, ByVal Subject As String, ByVal TextBody As String) As Boolean
'
'        On Error Resume Next
'        Dim objMail As CDONTS.NewMail
'100     Set objMail = New CDONTS.NewMail
'        '102     objMail.AttachFile App.Path & "\frmMain.frm"
'104     objMail.send aFrom, aTo, Subject, TextBody
'106     Set objMail = Nothing
'
'        '        'cached configuration
'        '        '<EhHeader>
'        '        On Error GoTo SendMailCDOCacheConf_Err
'        '        '</EhHeader>
'        '        Dim DI As DefaultInfo
'        '100     Call ReadDIConfig(DI)
'        '102     SendMailCDOCacheConf = True
'        '
'        '        Dim Conf As CDO.Configuration
'        '104     Set Conf = New CDO.Configuration
'        '
'        '        'If IsEmpty(Conf) Then
'        '        Const cdoSendUsingPort = 2
'        '
'        '        'Dim Conf As New CDO.Configuration
'        '
'        '106     With Conf.Fields
'        '108         .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'        '
'        '110         .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = DI.Smtp
'        '112                     .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = DI.Port
'        '114                     .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'        '116                     .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = DI.UserName
'        '118                     .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = DI.UserPwd
'        '120                     .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = DI.SSL
'        '122         .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
'        '124         .Update
'        '        End With
'        '
'        '        'End If
'        '
'        '        Dim Message As CDO.Message
'        '126     Set Message = New CDO.Message
'        '
'        '        'Create CDO message object
'        '        'Set Message = CreateObject("CDO.Message")
'        '
'        '128     With Message
'        '            'Set cached configuration
'        '130         Set .Configuration = Conf
'        '132         .BodyPart.CharSet = "UTF-8"
'        '            'Set email adress, subject And body
'        '134         .To = aTo
'        '136         .CC = aCC
'        '138         .Subject = Subject
'        '140         .HTMLBody = TextBody
'        '
'        '            'Set sender address If specified.
'        '142         If Len(aFrom) > 0 Then
'        '
'        '144             .From = aFrom
'        '
'        '            Else
'        '
'        '146             aFrom = DI.SenderName
'        '
'        '148             If DI.SenderEmail & "" <> "" Then
'        '
'        '150                 aFrom = aFrom & "<" & DI.SenderEmail & ">"
'        '
'        '                End If
'        '152             .From = aFrom
'        '            End If
'        '
'        '            'Send the message
'        '154         .send
'        '        End With
'        '
'        '        'WriteLog x.UTF2GB(Err.Description)
'        '156     Set Conf = Nothing
'        '158     Set Message = Nothing
'        '        '<EhFooter>
'        '        Exit Function
'        '
'        'SendMailCDOCacheConf_Err:
'        '
'        '        SendMailCDOCacheConf = False
'        '        'MsgBox x.UTF2GB(Err.Description)
'        '        Dim i As Integer
'        '        Dim strERR As String
'        '
'        '        For i = 1 To Len(Err.Description)
'        '
'        '            strERR = strERR & Chr(AscW(Mid(Err.Description, i, 1)))
'        '
'        '        Next
'        '
'        '        WriteLog "SendMailCDOCacheConf ERL:" & Erl & ":" & strERR
'        '        Resume Next
'        '        '</EhFooter>
'End Function


'
'
'Private Function ReadDIConfig(ByRef DI As DefaultInfo) As Boolean
'    On Error GoTo Err
'    Dim strConstConnString As String
'    Dim strTmp As String
'
'    WriteLog "Reading DI File..."
'
'    Open App.path & "\DI.dat" For Input As #1
'
'    'MsgBox App.Path
'
'    Dim i As Integer
'    i = 1
'
'    Dim clsEcypt As clsEncrypt
'    Set clsEcypt = New clsEncrypt
'
'    Do While Not EOF(1)
'
'        Line Input #1, strTmp
'
'        Select Case i
'
'            Case 1
'                DI.SenderName = clsEcypt.Decode(strTmp)
'
'            Case 2
'                DI.SenderEmail = clsEcypt.Decode(strTmp)
'
'            Case 3
'                DI.Smtp = clsEcypt.Decode(strTmp)
'
'            Case 4
'                DI.Port = clsEcypt.Decode(strTmp)
'
'            Case 5
'
'                DI.UserName = clsEcypt.Decode(strTmp)
'
'            Case 6
'                DI.UserPwd = clsEcypt.Decode(strTmp)
'
'            Case 7
'                DI.SSL = clsEcypt.Decode(strTmp)
'
'        End Select
'
'        i = i + 1
'
'    Loop
'
'    Close #1
'
'    If i <> 8 Then
'
'        ReadDIConfig = False
'
'        WriteLog "Reading DI Config Fail!"
'
'        Exit Function
'
'    End If
'
'    'MsgBox "a"
'    WriteLog "Reading DI Config Success!"
'
'    ReadDIConfig = True
'
'    Set clsEcypt = Nothing
'    Exit Function
'Err:
'    ReadDIConfig = False
'
'End Function

