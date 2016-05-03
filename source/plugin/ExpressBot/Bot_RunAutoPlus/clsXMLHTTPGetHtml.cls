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
Private Declare Sub Sleep _
                Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim strURL As String
Dim strMyPostData As String
Dim strCharSet As String
Dim iRetry As Integer
Dim iRetryTimes As Integer
Dim intTimeOut As Integer

Private strReferer As String

Private RawResponseText As String

Dim dicCookies As Scripting.Dictionary

Private dicRequestHeader As Scripting.Dictionary
Private Const iSleepTime = 2

Public Function StartGetHtml() As String
        '<EhHeader>
        On Error GoTo StartGetHtml_Err
        '</EhHeader>

100     If strURL = "" Then

102         StartGetHtml = ""
            Exit Function

        End If
    
104     iRetry = 0

        Dim oReq As MSXML2.XMLHTTP
        Set oReq = New MSXML2.XMLHTTP
    
Retry:

106     If strMyPostData = "" Then

108         oReq.Open "GET", Url, True
110         oReq.SetRequestHeader "Accept", "Text/Html"
112         oReq.SetRequestHeader "Accept-Encoding", "deflate"

114         If Not dicRequestHeader.Exists("User-Agent") Then
116             oReq.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 5.1; rv:35.0) Gecko/20100101 Firefox/35.0"
                'oReq.SetRequ estHeader "User-Agent", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.46 Safari/536.5"
            Else
                oReq.SetRequestHeader "User-Agent", dicRequestHeader.Item("User-Agent")
            End If

            '        oReq.setRequestHeader "Accept-Charset", "GB2312"
            'oReq.Open "Get", "http://" & Url, True
118         Call SetCookies(oReq)
120         oReq.send Null

        Else

122         oReq.Open "POST", Url, True
124         oReq.SetRequestHeader "Content-Length", Len(strMyPostData)
126         oReq.SetRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
128         oReq.SetRequestHeader "Accept", "Text/Html"
130         oReq.SetRequestHeader "Accept-Encoding", "deflate"
132         Call SetCookies(oReq)
134         oReq.send strMyPostData

        End If

        'oreq.
        Dim i As Long
        
136     i = 1

138     Do While oReq.readyState <> 4

140         DoEvents

142         MySleep 0.1
144         i = i + 1

146         If i > CInt(intTimeOut * 10) Then

148             i = 1
150             WriteLog "***本内容获取超时***" & strURL
                Debug.Print VBA.Now()
152             If iRetry < iRetryTimes Then
                    Dim SleepTime As Single
154                 SleepTime = 5 * iRetry
156                 iRetry = iRetry + 1
158                 WriteLog "第 " & iRetry & " 次重试，在" & SleepTime & "秒之后 " & strURL
160                 MySleep SleepTime
oReq.abort
162                 GoTo Retry

                Else

164                 WriteLog "3次重试均告失败。。。" & strURL
166                 iRetry = 0

                End If

                Exit Do

            End If

168         DoEvents

        Loop

170     If oReq.readyState = 4 Then
            
172         If oReq.status = 0 Then
174             StartGetHtml = "oReq.Status = 0"
                Exit Function
            End If
            
176         If oReq.status <> 200 Then

178             If iRetry >= 3 Then
180                 RawResponseText = oReq.status & oReq.responseText
182                 StartGetHtml = "Status: " & oReq.status & ", 第 " & iRetry & " 次重试失败！"
184                 WriteLog StartGetHtml
186                 GoTo ExitFunc

                End If

188             If oReq.status = 503 Then
190                 StartGetHtml = "Status: " & oReq.status & ", 第 " & iRetry & " 次重试失败！"
192                 GoTo ExitFunc
                End If

194             If oReq.status = 404 Then
196                 StartGetHtml = "Status: " & oReq.status & ", 页面不存在"
198                 GoTo ExitFunc
                End If

200             iRetry = iRetry + 1
202             WriteLog "Status: " & oReq.status & ", 第 " & iRetry & " 次重试，在" & iSleepTime * iRetry * 10 & "秒之后 " & strURL
204             MySleep iSleepTime * iRetry * 10
206             GoTo Retry

            End If
        
208         If oReq.responseText <> "" Then

                Dim strCookies As String
210             strCookies = ""
            
212             strCookies = oReq.getAllResponseHeaders
            
214             If strCookies <> "" Then
            
                    Dim Reg As VBScript_RegExp_55.RegExp
216                 Set Reg = New VBScript_RegExp_55.RegExp
                    Dim Mc As VBScript_RegExp_55.MatchCollection
                    Dim m As VBScript_RegExp_55.Match
218                 Reg.Global = True
220                 Reg.IgnoreCase = True
222                 Reg.MultiLine = False
224                 Reg.Pattern = "Set-Cookie\: (.*?)=(.*?) path=/; expires="
226                 Set Mc = Reg.Execute(strCookies)
                
228                 If Mc.Count > 0 Then

230                     dicCookies.RemoveAll

232                     For Each m In Mc
                        
234                         dicCookies.Add m.SubMatches(0), m.SubMatches(1)
                    
                        Next
                
                    End If

                End If

236             If strCharSet = "" Then
238                 strCharSet = "GB2312"
                End If

240             StartGetHtml = BytesToBstr(oReq.responseBody, strCharSet)

242             StartGetHtml = convertCRLF(StartGetHtml)

            Else

244             StartGetHtml = ""

            End If

        Else

246         StartGetHtml = ""

        End If

ExitFunc:

248     Set oReq = Nothing
    
        '<EhFooter>
        Exit Function

StartGetHtml_Err:
        Resume Next
        'Err.Raise vbObjectError + 100, "clsXMLHTTPGetHtml.StartGetHtml at line " & Erl, Err.Description
        '</EhFooter>
End Function

Public Sub ClearAllCookies()
    
    dicCookies.RemoveAll

End Sub

Private Sub SetCookies(ByRef obj As MSXML2.XMLHTTP60)

100     If dicCookies.Count > 0 Then
        
            Dim v As Variant
            Dim strResult As String
102         strResult = ""
104         For Each v In dicCookies.Keys
            
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

Public Property Get Url() As String
    Url = strURL
End Property

Public Property Let Url(ByVal Url As String)
    Url = Replace(Url, "&amp;", "&", 1, -1, vbBinaryCompare)
    strURL = Url
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
Public Property Let RetryTimes(ByVal RetryTimes As Integer)
    iRetryTimes = RetryTimes
End Property
Public Property Get RetryTimes() As Integer
   RetryTimes = iRetryTimes
End Property
Public Property Let Cookies(ByVal MydicCookies As Scripting.Dictionary)
    Set dicCookies = MydicCookies
End Property
Private Sub Class_Initialize()
    Set dicCookies = New Scripting.Dictionary
    Set dicRequestHeader = New Scripting.Dictionary
    intTimeOut = 10 '默认10秒超时~
    iRetryTimes = 3
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
Private Sub SetRequestHeaders(ByRef obj As MSXML2.XMLHTTP60)

    If dicRequestHeader.Count > 0 Then
        
        Dim v As Variant

        For Each v In dicRequestHeader.Keys
            
            obj.SetRequestHeader CStr(v), dicRequestHeader.Item(v)
            
        Next
      
    End If

End Sub

