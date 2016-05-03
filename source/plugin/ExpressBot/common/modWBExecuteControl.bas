Attribute VB_Name = "modWBExecuteControl"
Option Explicit


Public Function checkWBExecuteCanExecute(Optional ByVal RedialInterval As Integer = 10) As Boolean
        '<EhHeader>
        On Error GoTo checkWBExecuteCanExecute_Err
        '</EhHeader>

        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
        Dim iFile As Integer



108         If FSo.FileExists("D:\ServiceApp\GlobalWBExecuteControl\WBExecuting") Then
                
110             If DateDiff("s", FSo.GetFile("D:\ServiceApp\GlobalWBExecuteControl\WBExecuting").DateLastModified, Now()) > 5 Then
112                 WriteLog "发现上次重拨的残留文件WBExecuting，并且已经超过5秒钟了，强制删除！"
114                 Call UnlockWBExecute
116                 checkWBExecuteCanExecute = True
                Else
                
118                 checkWBExecuteCanExecute = False
                End If
            Else
                checkWBExecuteCanExecute = True
            End If


154     Set FSo = Nothing
        '<EhFooter>
        Exit Function

checkWBExecuteCanExecute_Err:
        checkWBExecuteCanExecute = True
        

        '</EhFooter>
End Function

Public Sub LockWBExecute()

        Dim iFile As Integer

100     iFile = FreeFile()
102     Open "D:\ServiceApp\GlobalWBExecuteControl\WBExecuting" For Output As #iFile
104     Print #iFile, ""
106     Close #iFile


End Sub

Public Sub UnlockWBExecute()
        On Error Resume Next
        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
    
102     If FSo.FileExists("D:\ServiceApp\GlobalWBExecuteControl\WBExecuting") Then
104         FSo.DeleteFile ("D:\ServiceApp\GlobalWBExecuteControl\WBExecuting")
        End If
    
116     Set FSo = Nothing


End Sub
