VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLHTTPGetHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim strUrl As String
Dim strMyPostData As String
Dim strCharSet As String
Dim iRetry As Integer
Dim intTimeOut As Integer

Private strReferer As String

Dim dicCookies As Scripting.Dictionary

Private dicRequestHeader As Scripting.Dictionary
Private Const iSleepTime = 2
Private strReturnData As String
Private intReturnCode As Integer

Public Function Send()
        '<EhHeader>
        On Error GoTo Send_Err
        '</EhHeader>

100     If strUrl = "" Then

102
            Exit Function

        End If
    
104     iRetry = 0

        Dim oReq As New MSXML2.XMLHTTP60
    
Retry:

106     If strMyPostData = "" Then
            
            
108         oReq.Open "GET", strUrl, True
            oReq.SetRequestHeader "Cookie", ""
112         oReq.Send Null

        Else

114         oReq.Open "POST", strUrl, True
116         oReq.SetRequestHeader "Content-Length", Len(strMyPostData)
118         oReq.SetRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
120         oReq.SetRequestHeader "Accept", "text/html"
122         oReq.SetRequestHeader "Accept-Encoding", "deflate"

126         oReq.Send strMyPostData

        End If

        'oreq.
        Dim i As Long
        
128     i = 1

130     Do While oReq.readyState <> 4

132         DoEvents

134         Sleep 10
136         i = i + 1

138         If i > intTimeOut * 100 Then

140             i = 1
142             WriteLog "***本内容获取超时***" & strUrl

144             If iRetry < 3 Then
                    Dim SleepTime As Single
146                 SleepTime = 5 * iRetry
148                 iRetry = iRetry + 1
150                 WriteLog "第 " & iRetry & " 次重试，在" & SleepTime & "秒之后 " & strUrl
152                 MySleep SleepTime
154                 GoTo Retry

                Else

156                 WriteLog "3次重试均告失败。。。" & strUrl
158                 iRetry = 0

                End If

                Exit Do

            End If

160         DoEvents

        Loop

162     If oReq.readyState = 4 Then

            intReturnCode = oReq.Status

164         If oReq.Status = 0 Then
                strReturnData = ""
                Exit Function
            End If
            
168         If oReq.Status <> 200 Then
                
170             If iRetry >= 3 Then

172                 strReturnData = "Status: " & oReq.Status & vbCrLf & oReq.responseText
174                 WriteLog strUrl & "@@" & strMyPostData & "::" & strReturnData
176                 GoTo ExitFunc

                End If

                If oReq.Status = 500 Or oReq.Status = 503 Or oReq.Status = 404 Then
180                 strReturnData = "Status: " & oReq.Status & vbCrLf & oReq.responseText
                    WriteLog strUrl & "@@" & strMyPostData & "::" & strReturnData
182                 GoTo ExitFunc
                End If

178

190             iRetry = iRetry + 1
192             WriteLog "Status: " & oReq.Status & ", 第 " & iRetry & " 次重试，在" & iSleepTime * iRetry & "秒之后 " & strUrl
194             MySleep iSleepTime * iRetry
196             GoTo Retry

            End If
        
198         If oReq.responseText <> "" Then

                Dim strCookies As String
200             strCookies = ""
            
202             strCookies = oReq.getAllResponseHeaders
            
204             If strCookies <> "" Then
            
                    Dim Reg As VBScript_RegExp_55.RegExp
206                 Set Reg = New VBScript_RegExp_55.RegExp
                    Dim Mc As VBScript_RegExp_55.MatchCollection
                    Dim m As VBScript_RegExp_55.Match
208                 Reg.Global = True
210                 Reg.IgnoreCase = True
212                 Reg.MultiLine = False
214                 Reg.Pattern = "Set-Cookie\: (.*?)=(.*?) "
216                 Set Mc = Reg.Execute(strCookies)
                
218                 If Mc.Count > 0 Then

220                     dicCookies.RemoveAll

222                     For Each m In Mc
                        
224                         dicCookies.Add m.SubMatches(0), m.SubMatches(1)
                    
                        Next
                
                    End If

                End If

226             If strCharSet = "" Then
228                 strCharSet = "UTF-8"
                End If

230             strReturnData = BytesToBstr(oReq.responseBody, strCharSet)

232             strReturnData = strReturnData

            Else

234             strReturnData = ""

            End If

        Else
            intReturnCode = -99
236         strReturnData = ""

        End If

ExitFunc:

238     Set oReq = Nothing
    
        '<EhFooter>
        Exit Function

Send_Err:
        Resume Next
        'Err.Raise vbObjectError + 100, "clsXMLHTTPGetHtml.StartGetHtml at line " & Erl, Err.Description
        '</EhFooter>
End Function

Public Sub ClearAllCookies()
    
    dicCookies.RemoveAll

End Sub

Private Sub SetCookies(ByRef obj As Variant)

100     If dicCookies.Count > 0 Then
        
            Dim v As Variant
            Dim strResult As String
102         strResult = ""

104         For Each v In dicCookies.keys
            
106             strResult = strResult & v & "=" & dicCookies.Item(v) & " "
            
            Next

108         If strResult <> "" Then
110             obj.SetRequestHeader "Cookie", Left(strResult, Len(strResult) - 2)
112             obj.SetRequestHeader "Cookie", Left(strResult, Len(strResult) - 2)      'One More Time to Fix the Microsoft's Bug
            End If
        End If

End Sub

Private Function BytesToBstr(ByVal body As Variant, _
   ByVal Cset As String) As String
        '<EhHeader>
        On Error GoTo BytesToBstr_Err
        '</EhHeader>

        If UBound(body) > 0 Then
            Dim objStream As ADODB.Stream
            Set objStream = New ADODB.Stream
100         objStream.Type = 1
110         objStream.Mode = 3
120         objStream.Open
130         objStream.Write body
140         objStream.Position = 0
150         objStream.Type = 2
160         objStream.CharSet = Cset
170         BytesToBstr = objStream.ReadText
180         objStream.Close
190         Set objStream = Nothing
            '<EhFooter>
        Else
            BytesToBstr = ""
        End If

        Exit Function

BytesToBstr_Err:
        Err.Raise vbObjectError + 100, "in modHTTP.BytesToBstr at line " & Erl, Err.Description
        Resume Next
        '</EhFooter>
End Function

Public Property Get ReturnCode() As Integer
    ReturnCode = intReturnCode
End Property

Public Property Get ReturnData() As String
    ReturnData = strReturnData
End Property

Public Property Get URL() As String
    URL = strUrl
End Property

Public Property Let URL(ByVal URL As String)
    URL = Replace(URL, "&amp;", "&", 1, -1, vbBinaryCompare)
    strUrl = URL
    strMyPostData = ""
End Property

Public Property Get PostData() As String
    PostData = strMyPostData
End Property

Public Property Let PostData(ByVal PostData As String)
    strMyPostData = PostData
End Property

Public Property Let CharSet(ByVal MyCharSet As String)
    strCharSet = MyCharSet
End Property

Public Property Get Cookies() As Scripting.Dictionary
    Set Cookies = dicCookies
End Property

Public Property Let TimeOut(ByVal TimeOutSecond As Integer)
    intTimeOut = TimeOutSecond
End Property

Public Property Get TimeOut() As Integer
    TimeOut = intTimeOut
End Property

Public Property Let Cookies(ByVal MydicCookies As Scripting.Dictionary)
    Set dicCookies = MydicCookies
End Property

Private Sub Class_Initialize()
    Set dicCookies = New Scripting.Dictionary
    Set dicRequestHeader = New Scripting.Dictionary
    intTimeOut = 10 '默认10秒超时~
End Sub

Public Property Let Referer(ByVal Referer As String)
    strReferer = Replace(Referer, "&amp;", "&", 1, -1, vbBinaryCompare)
End Property

Public Function SetRequestHeader(ByRef strName As String, ByRef strVal As String, Optional ForceOverWrite As Boolean = False) As Boolean

    If Not dicRequestHeader.Exists(strName) Then
    
        dicRequestHeader.Add strName, strVal
        SetRequestHeader = True
    Else

        If ForceOverWrite Then
            dicRequestHeader.Item(strName) = strVal
            SetRequestHeader = True
        Else
            SetRequestHeader = False
        End If
    End If

End Function

Private Sub SetRequestHeaders(ByRef obj As Variant)

    If dicRequestHeader.Count > 0 Then
        
        Dim v As Variant

        For Each v In dicRequestHeader.keys
            
            obj.SetRequestHeader CStr(v), dicRequestHeader.Item(v)
            
        Next
      
    End If

End Sub

