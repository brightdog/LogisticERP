Attribute VB_Name = "modCatchStateControl"
Option Explicit

Public Function CheckCanDoCatch(ByVal SiteName As String) As Boolean
        '<EhHeader>
        On Error GoTo checkWBExecuteCanExecute_Err
        '</EhHeader>
        
        '\StopDoCatch" & SiteName
        '\CanDoCatch" & SiteName

        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
        Dim iFile As Integer

108     If FSo.FileExists("D:\ServiceApp\GlobalWBExecuteControl\StopDoCatch_" & SiteName) Then
                
110         If DateDiff("s", FSo.GetFile("D:\ServiceApp\GlobalWBExecuteControl\StopDoCatch_" & SiteName).DateLastModified, Now()) > 30 Then
112             WriteLog "发现残留文件StopDoCatch，并且已经超过30秒钟了，强制删除！"
114             Call CanDoCatch(SiteName)
116             CheckCanDoCatch = True
            Else
                
118             CheckCanDoCatch = False
            End If

        Else
            CheckCanDoCatch = True
        End If

154     Set FSo = Nothing
        '<EhFooter>
        Exit Function

checkWBExecuteCanExecute_Err:
        CheckCanDoCatch = True

        '</EhFooter>
End Function

Public Sub CanDoCatch(ByVal SiteName As String)

        On Error Resume Next
        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
    
102     If FSo.FileExists("D:\ServiceApp\GlobalWBExecuteControl\StopDoCatch_" & SiteName) Then
104         FSo.DeleteFile ("D:\ServiceApp\GlobalWBExecuteControl\StopDoCatch_" & SiteName)
        End If

        If Not FSo.FileExists("D:\ServiceApp\GlobalWBExecuteControl\CanDoCatch_" & SiteName) Then
106         FSo.CreateTextFile "D:\ServiceApp\GlobalWBExecuteControl\CanDoCatch_" & SiteName, True, False
        End If

108     Set FSo = Nothing

End Sub

Public Sub StopDoCatch(ByVal SiteName As String)
        On Error Resume Next
        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
    
102     If FSo.FileExists("D:\ServiceApp\GlobalWBExecuteControl\CanDoCatch_" & SiteName) Then
104         FSo.DeleteFile ("D:\ServiceApp\GlobalWBExecuteControl\CanDoCatch_" & SiteName)
        End If

        'If Not FSo.FileExists("D:\ServiceApp\GlobalWBExecuteControl\StopDoCatch_" & SiteName) Then
106         FSo.CreateTextFile "D:\ServiceApp\GlobalWBExecuteControl\StopDoCatch_" & SiteName, True, False
'        End If

108     Set FSo = Nothing

End Sub
