VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinHTTPGetHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep _
                Lib "kernel32" (ByVal dwMilliseconds As Long)
Private strUrl As String
Private strMyPostData As String
Private strCharSet As String
Private strReferer As String
Private strLastURL As String
Private iRetry As Integer
Private iRedirectTimes As Integer

Private dicCookies As Scripting.Dictionary
Private dicRequestHeader As Scripting.Dictionary
Private WithEvents oReq As WinHttp.WinHttpRequest
Attribute oReq.VB_VarHelpID = -1
Private bolDataFinished As Boolean
Private Const iSleepTime = 2
Private intTimeOut As Integer
Private strReturnData As String
Private intReturnCode As Integer

Public Function Send() As String
        '<EhHeader>
        On Error GoTo Send_Err
        '</EhHeader>

100     If strUrl = "" Then

            '102         StartGetHtml = ""
            Exit Function

        End If
    
102     iRetry = 0
        Dim oReq As WinHttp.WinHttpRequest
104     Set oReq = New WinHttp.WinHttpRequest
106     oReq.Option(WinHttpRequestOption_EnableHttp1_1) = True
108     oReq.Option(WinHttpRequestOption_EnableRedirects) = True
110     oReq.Option(WinHttpRequestOption_MaxAutomaticRedirects) = iRedirectTimes
112     oReq.setTimeouts 30000, 30000, 30000, 60000
        ' 第一个数值：解析DNS名字的超时时间
        ' 第二个数值：建立Winsock连接的超时时间
        ' 第三个数值：发送数据的超时时间
        ' 第四个数值：接收response的超时时间
        
114     oReq.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 13056
        
Retry:

        'oReq.setProxy 2, "127.0.0.1:8888"
116     oReq.setProxy 0
        'HTTPREQUEST_PROXYSETTING_DEFAULT   = 0;
        'HTTPREQUEST_PROXYSETTING_PRECONFIG = 0;
        'HTTPREQUEST_PROXYSETTING_DIRECT    = 1;
        'HTTPREQUEST_PROXYSETTING_PROXY     = 2;
118     bolDataFinished = False
        
120     If strMyPostData = "" Then

122         oReq.Open "GET", strUrl, True

124         If Not dicRequestHeader.Exists("User-Agent") Then
126             oReq.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.46 Safari/536.5"
            End If

128         oReq.SetRequestHeader "Accept", "text/html"
            'oReq.SetRequestHeader "Proxy-Connection", "keep-alive"
130         oReq.SetRequestHeader "Pragma", "no-cache"

132         If strCharSet <> "" Then
134             oReq.SetRequestHeader "Accept-Charset", strCharSet
            End If

136         Call SetCookies(oReq)
138         Call SetRequestHeaders(oReq)

140         If strReferer <> "" Then
142             oReq.SetRequestHeader "Referer", strReferer
            Else

144             If strLastURL <> "" Then
146                 oReq.SetRequestHeader "Referer", strLastURL
                End If
            End If

148         MySleep 0.2
150         oReq.Send Null

        Else

152         oReq.Open "POST", strUrl, True
154         oReq.SetRequestHeader "Content-Length", Len(strMyPostData)
156         oReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            'oReq.SetRequestHeader "Proxy-Connection", "keep-alive"

158         If Not dicRequestHeader.Exists("User-Agent") Then
160             oReq.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.46 Safari/536.5"
            End If

162         oReq.SetRequestHeader "Accept", "text/json, text/javascript, text/html"
164         oReq.SetRequestHeader "Pragma", "no-cache"

166         If strCharSet <> "" Then
168             oReq.SetRequestHeader "Accept-Charset", strCharSet
            End If

170         Call SetCookies(oReq)
172         Call SetRequestHeaders(oReq)

174         If strReferer <> "" Then
176             oReq.SetRequestHeader "Referer", strReferer
            Else

178             If strLastURL <> "" Then
180                 oReq.SetRequestHeader "Referer", strLastURL
                End If
            End If

182         oReq.Send strMyPostData
            
        End If
        
184     strLastURL = strUrl
        'oreq.
        Dim i As Long
        
186     i = 1
        
188     Do While Not bolDataFinished

190         Debug.Print "111"
            
192         MySleep 0.2

194         If oReq.Status = 200 Then
196             Debug.Print "222"
                Exit Do
            End If

198         If iRetry >= 3 Then

200             strReturnData = "Status: " & oReq.Status & vbCrLf & oReq.responseText
202             WriteLog strReturnData
204             GoTo ExitFunc

            End If

206         If oReq.Status = 500 Or oReq.Status = 503 Or oReq.Status = 404 Then
208             strReturnData = "Status: " & oReq.Status & vbCrLf & oReq.responseText
210             GoTo ExitFunc
            End If

            'Debug.Print TypeName(oReq.Status)
212         DoEvents

214         Sleep 10
216         i = i + 1

218         If i > intTimeOut * 100 Then

220             i = 1
222             WriteLog "***本内容获取超时***" & strUrl

224             If iRetry < 3 Then
                    Dim SleepTime As Single
226                 SleepTime = 5 * iRetry
228                 iRetry = iRetry + 1
230                 WriteLog "第 " & iRetry & " 次重试，在" & SleepTime & "秒之后 " & strUrl
232                 MySleep SleepTime
234                 GoTo Retry

                Else

236                 WriteLog "3次重试均告失败。。。" & strUrl
238                 iRetry = 0

                End If

                Exit Do

            End If

240         DoEvents

        Loop

242     If oReq.Status = 200 Then
            
            '            If oReq.Status = 503 Then
            '                StartGetHtml = ""
            '                Exit Function
            '            End If

244         If oReq.statusText <> "OK" Then

WaitRetry:

246             If iRetry >= 10 Then

248                 strReturnData = "Status: " & oReq.Status & vbCrLf & oReq.responseText
250                 WriteLog strReturnData
252                 GoTo ExitFunc

                End If

254             iRetry = iRetry + 1
                'WriteLog "Status: " & oReq.Status & ":" & oReq.statusText & ", 第 " & iRetry & " 次重试，在" & iSleepTime * iRetry / 10 & "秒之后 " & strUrl

256             WriteLog "Wait: " & iRetry & ":" & iSleepTime * iRetry / 10
258             MySleep iSleepTime * iRetry / 10

260             If oReq.statusText <> "OK" Then
262                 Debug.Print "-999"
264                 GoTo WaitRetry
                End If

            End If
            
266         If oReq.statusText <> "OK" Then
268             Debug.Print "-999"
270             GoTo Retry
            End If
        
272         If oReq.responseText <> "" Then
274             Debug.Print "333"
                Dim strHeaders As String
276             strHeaders = ""
            
278             strHeaders = oReq.getAllResponseHeaders
            
280             If strHeaders <> "" Then
                    Dim arrHeader() As String
                    
282                 arrHeader = Split(strHeaders, vbCrLf, -1, vbBinaryCompare)
                    
284                 If UBound(arrHeader) >= 0 Then
                        
                        Dim iHeader As Integer
                        
                        Dim Reg As VBScript_RegExp_55.RegExp
286                     Set Reg = New VBScript_RegExp_55.RegExp
                        Dim Mc As VBScript_RegExp_55.MatchCollection
                        Dim m As VBScript_RegExp_55.Match
288                     Reg.Global = True
290                     Reg.IgnoreCase = True
292                     Reg.MultiLine = False
294                     Reg.Pattern = "Set-Cookie\:\s*([^=]+=.*?)$"

                        'dicCookies.RemoveAll
296                     For iHeader = 0 To UBound(arrHeader)

298                         Set Mc = Reg.Execute(arrHeader(iHeader))
                
300                         If Mc.Count > 0 Then



302                             For Each m In Mc
                                    Dim arr(1) As String
                        
304                                 arr(0) = Left(m.SubMatches(0), InStr(1, m.SubMatches(0), "=", vbBinaryCompare) - 1)
306                                 arr(1) = Right(m.SubMatches(0), Len(m.SubMatches(0)) - Len(arr(0)) - 1)

308                                 If Not dicCookies.Exists(arr(0)) Then
310                                     dicCookies.Add arr(0), arr(1)
                                    End If
                    
                                Next
                
                            End If

                        Next

312                     Set Reg = Nothing
314                     Set Mc = Nothing
316                     Set m = Nothing
                    End If
                End If
                
318             If strCharSet = "" Then
320                 strCharSet = "GB2312"
                End If

322             Debug.Print "444"
324             iRetry = 0
RetryResponseBody:
326             strReturnData = BytesToBstr(oReq.responseBody, strCharSet)
                
328             If strReturnData = "" Then
                
330                 If iRetry < 10 Then
332                     WriteLog iRetry
334                     iRetry = iRetry + 1
336                     MySleep 0.1 * iRetry
338                     GoTo RetryResponseBody

                    End If
                End If
                
340             Debug.Print "555" & "|" & Len(strReturnData)
342             strReturnData = strReturnData

            Else
344             Debug.Print "-444"
346             strReturnData = oReq.responseText

            End If

        Else
348         Debug.Print "-666"
350         strReturnData = oReq.responseText

        End If

ExitFunc:
352     oReq.abort
354     Set oReq = Nothing
    
        '<EhFooter>
        Exit Function

Send_Err:

        If Err.Number = -2147483638 Then
            Resume Next
        Else
        
            WriteLog Err.Number & "|" & Err.Description
            
        End If

        Debug.Print TypeName(oReq.Status) & "*"
        Resume Next
        'Err.Raise vbObjectError + 100, "clsXMLHTTPGetHtml.StartGetHtml at line " & Erl, Err.Description
        '</EhFooter>
End Function


Private Sub SetCookies(ByRef obj As WinHttp.WinHttpRequest)

    If dicCookies.Count > 0 Then
        
        Dim v As Variant
        Dim strResult As String
        strResult = ""
        For Each v In dicCookies.keys
            
            strResult = strResult & v & "=" & dicCookies.Item(v) & "; "
            
        Next
        If strResult <> "" Then
            obj.SetRequestHeader "Cookie", Left(strResult, Len(strResult) - 2)
            'obj.AddRequestHeader "Cookie", Left(strResult, Len(strResult) - 2)      'One More Time to Fix the Microsoft's Bug
        End If
    End If

End Sub

Private Sub SetRequestHeaders(ByRef obj As WinHttp.WinHttpRequest)

    If dicRequestHeader.Count > 0 Then
        
        Dim v As Variant

        For Each v In dicRequestHeader.keys
            
            obj.SetRequestHeader CStr(v), dicRequestHeader.Item(v)
            
        Next
      
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

Public Property Get URL() As String
    URL = strUrl
End Property

Public Property Let URL(ByVal URL As String)
    URL = Replace(URL, "&amp;", "&", 1, -1, vbBinaryCompare)
    strUrl = URL
    strMyPostData = ""
    strReferer = ""
End Property

Public Property Let Referer(ByVal Referer As String)
    strReferer = Replace(Referer, "&amp;", "&", 1, -1, vbBinaryCompare)
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

Public Property Let Cookies(ByVal MydicCookies As Scripting.Dictionary)
    Set dicCookies = MydicCookies
End Property
Public Property Let RedirectTimes(ByVal RedirectTimes As Integer)
    iRedirectTimes = RedirectTimes
End Property

Public Property Let TimeOut(ByVal TimeOutSecond As Integer)
    intTimeOut = TimeOutSecond
End Property
Public Property Get TimeOut() As Integer
   TimeOut = intTimeOut
End Property



Private Sub Class_Initialize()
    Set dicCookies = New Scripting.Dictionary
    Set dicRequestHeader = New Scripting.Dictionary
    iRedirectTimes = 5  '默认跳转5次
End Sub

Public Sub ClearAllCookies()
    
    dicCookies.RemoveAll

End Sub

Public Sub ClearAllRequestHeaders()

    dicRequestHeader.RemoveAll
End Sub

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

Public Function SetRequestCookie(ByRef strName As String, ByRef strVal As String, Optional ForceOverWrite As Boolean = False) As Boolean

    If Not dicCookies.Exists(strName) Then
    
        dicCookies.Add strName, strVal
        SetRequestCookie = True
    Else
        If ForceOverWrite Then
            dicCookies.Item(strName) = strVal
            SetRequestCookie = True
        Else
            SetRequestCookie = False
        End If
    End If

End Function

Private Sub oReq_OnResponseDataAvailable(data() As Byte)
    bolDataFinished = True
End Sub

Private Sub oReq_OnResponseFinished()
    bolDataFinished = True
End Sub




Public Property Get ReturnCode() As Integer
    ReturnCode = intReturnCode
End Property

Public Property Get ReturnData() As String
    ReturnData = strReturnData
End Property

