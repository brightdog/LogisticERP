VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDownFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'StartDownFile 文件地址, 保存目录

Option Explicit
Private Declare Function InternetOpen _
                Lib "wininet.dll" _
                Alias "InternetOpenA" (ByVal sAgent As String, ByVal LAccessType As Long, ByVal sProxyName As String, ByVal SProxyBypass As String, ByVal lFlags As Long) As Long

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0

Private Declare Function InternetConnect _
                Lib "wininet.dll" _
                Alias "InternetConnectA" (ByVal HinterNetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Const INTERNET_DEFAULT_HTTP_PORT = 80

Private Const INTERNET_SERVICE_HTTP = 3
Private Declare Function httpopenrequest _
                Lib "wininet.dll" _
                Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function HttpSendRequest _
                Lib "wininet.dll" _
                Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Integer

Private Declare Function InternetReadImg _
                Lib "wininet.dll" _
                Alias "InternetReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal dwNumberOfBytesToRead As Long, lNumberOfBytesRead As Long) As Integer

Private Declare Function HttpQueryInfo _
                Lib "wininet.dll" _
                Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

Private Const HTTP_QUERY_CONTENT_LENGTH = 5

Private Declare Function InternetCloseHandle _
                Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Declare Function HttpAddRequestHeaders _
                Lib "wininet.dll" _
                Alias "HttpAddRequestHeadersA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lModifiers As Long) As Integer

Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000
Private Const HTTP_QUERY_STATUS_CODE = 19
Dim TempUrl As String
Dim FileName As String
Dim Revice As Long
Dim Total As Long
Dim isend As Boolean

Public Event DownLoadBlock(ByVal iBlock As Long)

'开始下载文件
Public Function StartDownFile(URL As String, _
                              path As String) As Boolean
        On Error GoTo errh
        Dim Session As Long
        Dim Connect As Long
        Dim host As String
        Dim request As Long
        Dim Getstr As String
        Dim length As String
        Dim lret As Long
        Dim code As String
        Dim currpath As String
        Dim name As String
100     name = GetName(URL)

102     DoEvents

        Dim intDownFileRetry As Integer
104     intDownFileRetry = 0

106     If URL <> "" And path <> "" Then

DownFileRetry:

108         Session = InternetOpen(URL, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0) '打开连接

110         If Session Then

112             host = GetIP(URL)
114             Connect = InternetConnect(Session, host, 80, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)

116             If Connect Then

118                 currpath = GetPath(URL, host)
120                 request = httpopenrequest(Connect, "GET", currpath, "HTTP/1.1", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

122                 If request Then

124                     DoEvents

126                     Getstr = "Accept: */*"
128                     HttpAddRequestHeaders request, Getstr, Len(Getstr), HTTP_ADDREQ_FLAG_ADD
130                     Getstr = "User-Agent: DownJet1.0"
132                     HttpAddRequestHeaders request, Getstr, Len(Getstr), HTTP_ADDREQ_FLAG_ADD
134                     Getstr = "Host: " & host
136                     HttpAddRequestHeaders request, Getstr, Len(Getstr), HTTP_ADDREQ_FLAG_ADD
138                     Getstr = "Connection: close"
140                     HttpAddRequestHeaders request, Getstr, Len(Getstr), HTTP_ADDREQ_FLAG_ADD
142                     lret = HttpSendRequest(request, vbNullString, 0, 0, 0)

144                     If lret Then

146                         DoEvents

148                         GetQueryInfo request, length, HTTP_QUERY_CONTENT_LENGTH '得到文件的长度
150                         GetQueryInfo request, code, HTTP_QUERY_STATUS_CODE
152                         Total = Val(length)
154                         Revice = 0

156                         If code = "200" Or code = "206" Then

158                             DoEvents
                            
160                             If Not SaveData(request, path) Then
                                
162                                 If intDownFileRetry > 2 Then
                                
164                                     StartDownFile = False

                                    Else
                                
166                                     intDownFileRetry = intDownFileRetry + 1
168                                     GoTo DownFileRetry

                                    End If

                                Else

170                                 StartDownFile = True

                                End If
                            
                            Else
                        
172                             WriteLog "Download ERR :" & code
                            
                            End If

                        End If

                    End If

                End If

            End If

        End If

174     InternetCloseHandle Session
176     InternetCloseHandle Connect
178     InternetCloseHandle request
        Exit Function
errh:
180     StartDownFile = False
182     WriteLog "StartDownFile = False:" & Err.Description & "ERL=" & Erl, True
End Function
'是否下载完成
Property Get IsOver() As Boolean
    IsOver = isend
End Property

Private Function GetName(URL As String) As String
    Dim stemp() As String
    stemp = Split(URL, "/")

    If UBound(stemp) > 0 Then

        GetName = stemp(UBound(stemp))

    Else

        GetName = ""

    End If

End Function

'得到相对路径
Private Function GetPath(URL As String, host As String) As String
    Dim stemp() As String
    stemp = Split(URL, host)

    If UBound(stemp) > 0 Then

        GetPath = stemp(1)

    Else

        GetPath = ""

    End If

End Function
'得到总长度
Property Get CountLength() As Long
    CountLength = Total
End Property
'接收的长度
Property Get ReviceCount() As Long
    ReviceCount = Revice
End Property

'保存数据
Private Function SaveData(request As Long, path As String) As Boolean
        On Error GoTo errh
        Dim buffer(20480) As Byte '为InternetReadFile函数提供字符串作为数据接收缓冲区
        Dim currbyte As Long '调用InternetReadFile函数读取的字节数
        Dim i As Long
        Dim FreeNumber As Integer
        Dim bDoLoop As Boolean
100     bDoLoop = True
102     FreeNumber = FreeFile

104     DoEvents
    
        Dim strStartTime As String
106     strStartTime = Now()
    
108     Open path For Binary As #FreeNumber
110     While bDoLoop

112         DoEvents
        
114         If currbyte = 0 And DateDiff("s", strStartTime, Now()) > 120 Then
        
116             SaveData = False
118             Close #FreeNumber
                Exit Function

            End If
        
120         bDoLoop = InternetReadImg(request, buffer(0), 20480, currbyte)

122         If Not CBool(bDoLoop) Then

124             Close #FreeNumber
                Exit Function

            End If

126         Seek #FreeNumber, Revice + 1

128         For i = 0 To currbyte - 1

130             DoEvents

132             Put #FreeNumber, , buffer(i)

134         Next i

136         Revice = Revice + currbyte

138         RaiseEvent DownLoadBlock(Revice)

140         If Revice = Total Then
            
142             Close #FreeNumber
144             SaveData = True
                Exit Function

            End If

        Wend
146     Close #FreeNumber
        Exit Function
errh:
148     SaveData = False
150     WriteLog "SaveData = False:" & Err.Description & "ERL=" & Erl, True
End Function

'得到IP地址
Private Function GetIP(temp As String) As String
    On Error Resume Next
    Dim stemp() As String

    If temp <> "" Then

        stemp = Split(temp, "http://")

        If UBound(stemp) > 0 Then

            stemp = Split(stemp(1), "/")
            GetIP = stemp(0)

        End If

    End If

End Function

'得到文件的长度
Private Function GetQueryInfo(ByVal hHttpRequest As Long, sReContent As String, ByVal iInfoLevel As Long) As String
        '<EhHeader>
        On Error GoTo GetQueryInfo_Err
        '</EhHeader>

        Dim sBuffer As String * 1024
        Dim lBufferLength As Long
100     lBufferLength = Len(sBuffer)
102     HttpQueryInfo hHttpRequest, iInfoLevel, ByVal sBuffer, lBufferLength, 0
104     sReContent = Mid(sBuffer, 1, lBufferLength)

        '<EhFooter>
        Exit Function

GetQueryInfo_Err:
        WriteLog "GetQueryInfo:" & Err.Description & "ERL=" & Erl, True
        '</EhFooter>
End Function

