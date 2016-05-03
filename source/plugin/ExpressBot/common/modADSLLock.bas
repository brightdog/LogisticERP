Attribute VB_Name = "modADSLLock"
Option Explicit


Public Function NoDialADSL() As Boolean
        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
        Dim iFile As Integer

102     If FSo.FileExists("D:\ServiceApp\GlobalADSLControl\NOREDIAL") Then
104         NoDialADSL = True
        Else

106         iFile = FreeFile()
108         Open "D:\ServiceApp\GlobalADSLControl\NOREDIAL" For Output As #iFile
110         Close #iFile
        End If

112     Set FSo = Nothing
End Function

Public Function CanDialADSL() As Boolean
        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
        Dim iFile As Integer

102     If FSo.FileExists("D:\ServiceApp\GlobalADSLControl\NOREDIAL") Then
104         Call FSo.DeleteFile("D:\ServiceApp\GlobalADSLControl\NOREDIAL")
106         CanDialADSL = True
        End If

108     Set FSo = Nothing
End Function

Public Function checkADSLCanReconnect(Optional ByVal RedialInterval As Integer = 10) As Boolean
        '<EhHeader>
        On Error GoTo checkADSLCanReconnect_Err
        '</EhHeader>

        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
        Dim iFile As Integer

102     If FSo.FileExists("D:\ServiceApp\GlobalADSLControl\NOREDIAL") Then
104         checkADSLCanReconnect = False
106         WriteLog "用户强制不重拨"
        
        Else

108         If FSo.FileExists("D:\ServiceApp\GlobalADSLControl\ADSLRECONNECTING") Then
                
110             If DateDiff("n", FSo.GetFile("D:\ServiceApp\GlobalADSLControl\ADSLRECONNECTING").DateLastModified, Now()) > 2 Then
112                 WriteLog "发现上次重拨的残留文件ADSLRECONNECTING，并且已经超过2分钟了，强制删除！"
114                 Call UnlockADSL
116                 checkADSLCanReconnect = True
                Else
                
118                 checkADSLCanReconnect = False
                End If

            Else
                If RedialInterval < 0 Then  '确保那些没有写过配置文件的爬虫，可以使用默认值 2014-04-16
                    RedialInterval = 10
                End If
                Dim strLastTime As String

120             iFile = FreeFile()
122             Open "D:\ServiceApp\GlobalADSLControl\LastTime" For Input As #iFile

124             If Not EOF(iFile) Then
126                 Line Input #iFile, strLastTime
                Else
128                 WriteLog "LastTime文件内容为空"
130                 checkADSLCanReconnect = True
132                 strLastTime = ""
                End If

134             Close #iFile

136             If strLastTime <> "" Then
138                 If DateDiff("s", Now, strLastTime) > 0 Then
140                     WriteLog "系统时间设置可能调值更早的时间了！"
142                     checkADSLCanReconnect = True
                    Else

144                     If DateDiff("s", strLastTime, Now) >= RedialInterval Then
146                         checkADSLCanReconnect = True
                        Else
148                         WriteLog "距离上次重拨不足" & RedialInterval & "秒，等！"
150                         checkADSLCanReconnect = False
                        End If
                    End If

                Else

152                 checkADSLCanReconnect = True
                End If
            End If
        End If

154     Set FSo = Nothing
        '<EhFooter>
        Exit Function

checkADSLCanReconnect_Err:
        checkADSLCanReconnect = True
        
        If Err.Number = 53 Then

            iFile = FreeFile()
            Open "D:\ServiceApp\GlobalADSLControl\LastTime" For Output As #iFile
            Print #iFile, Now
            Close #iFile
        End If
        
        '        If Err.Number = 62 Then
        '
        '            FSo.DeleteFile "D:\ServiceApp\GlobalADSLControl\LastTime", True
        '            checkADSLCanReconnect = True
        '        End If
        
        '</EhFooter>
End Function

Public Sub LockADSL()

        Dim iFile As Integer

100     iFile = FreeFile()
102     Open "D:\ServiceApp\GlobalADSLControl\ADSLRECONNECTING" For Output As #iFile
104     Print #iFile, ""
106     Close #iFile

108     iFile = FreeFile()
110     Open "D:\ServiceApp\GlobalADSLControl\LastTime" For Output As #iFile
112     Print #iFile, Now
114     Close #iFile

End Sub

Public Sub UnlockADSL()
        On Error Resume Next
        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
    
102     If FSo.FileExists("D:\ServiceApp\GlobalADSLControl\ADSLRECONNECTING") Then
104         FSo.DeleteFile ("D:\ServiceApp\GlobalADSLControl\ADSLRECONNECTING")
        End If
    
116     Set FSo = Nothing
        Dim iFile As Integer

106     iFile = FreeFile()
108     Open "D:\ServiceApp\GlobalADSLControl\LastTime" For Output As #iFile
110     Print #iFile, Now
112     Close #iFile

End Sub

