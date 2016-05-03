Attribute VB_Name = "modZipnUpload"
Option Explicit

Dim mstrZippedFolder As String
Dim mstrRAWFolder As String

Public Sub ZipnUploadResultFile()
100     mstrZippedFolder = App.Path & "\ZippedResultFile"
102     mstrRAWFolder = App.Path & "\Result"
104     Call ZipResultFile
    
106     Call UploadResultFile
        'Call ZipSqlFile(App.Path & "\SQL\" & gstrSqlFileName, "")

End Sub

Private Sub ZipResultFile()

        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject

102     If Not FSo.FolderExists(mstrRAWFolder) Then
104         Call FSo.CreateFolder(mstrRAWFolder)
        End If
        
106     If Not FSo.FolderExists(mstrZippedFolder) Then
108         Call FSo.CreateFolder(mstrZippedFolder)
        End If
    
        Dim Fles As Scripting.Files
110     Set Fles = FSo.GetFolder(mstrRAWFolder).Files
    
112     If Fles.Count > 0 Then
    
            Dim FLe As Scripting.File
        
114         For Each FLe In Fles
                Dim strGUID As String
            
                Dim objGUID As clsRowGuid
116             Set objGUID = New clsRowGuid
            
118             strGUID = objGUID.Guid
120             strGUID = Replace(strGUID, "{", "")
122             strGUID = Replace(strGUID, "}", "")
                Dim strZipName As String
            
124             strZipName = mstrZippedFolder & "\" & GetClientName & "_" & strGUID & ".zip"
            
126             Call ZipSqlFile(FLe.Path, strZipName)
128             Call MoveFiletoBkFolder(FLe.Path)
130             Set objGUID = Nothing
            Next
    
        End If
    
132     Set Fles = Nothing
134     Set FSo = Nothing

End Sub

Private Sub MoveFiletoBkFolder(ByVal strFilePath As String)
    
        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
    
        Dim strFolderPath As String
        Dim strFileName As String
    
102     strFolderPath = GetFolderPathFromFilePath(strFilePath)
        Dim arrfileInfo() As String
104     arrfileInfo = Split(strFilePath, "\", -1, vbBinaryCompare)
106     strFileName = arrfileInfo(UBound(arrfileInfo))
108     If Not FSo.FolderExists(strFolderPath & "\bk") Then
110         Call FSo.CreateFolder(strFolderPath & "\bk")
        End If
112     If Not FSo.FileExists(strFolderPath & "\bk\" & strFileName) Then
114         Call FSo.MoveFile(strFilePath, strFolderPath & "\bk\")
        Else
116         If FSo.FileExists(strFilePath) Then
118             Call FSo.DeleteFile(strFilePath)
            Else
            End If
        End If
120     Set FSo = Nothing

End Sub

Private Function GetFolderPathFromFilePath(ByVal strFilePath As String) As String

        Dim arr() As String
100     arr = Split(strFilePath, "\", -1, vbBinaryCompare)
        Dim strResult As String
102     strResult = ""

104     If UBound(arr) > -1 Then
        
            Dim i As Integer
        
106         For i = 0 To UBound(arr) - 1
        
108             strResult = strResult & arr(i)

110             If i < UBound(arr) - 1 Then
            
112                 strResult = strResult & "\"
            
                End If

            Next
        
        Else
    
        End If

114     GetFolderPathFromFilePath = strResult
End Function

Private Sub UploadResultFile()

        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
    
102     If Not FSo.FolderExists(mstrZippedFolder) Then
104         Call FSo.CreateFolder(mstrZippedFolder)
        End If
    
        Dim Fles As Scripting.Files
106     Set Fles = FSo.GetFolder(mstrZippedFolder).Files
    
108     If Fles.Count > 0 Then
    
            Dim FLe As Scripting.File
        
110         For Each FLe In Fles
                Dim fBytes() As Byte
            
                Dim iFile As Integer
112             iFile = VBA.FreeFile()
114             Open FLe.Path For Binary As #iFile
116             ReDim fBytes(LOF(iFile) - 1)
118             Get #iFile, , fBytes
120             Close #iFile
            
            
                Dim objUp As clsws_SFMainService
122             Set objUp = New clsws_SFMainService
            
124             Call objUp.SOAPUploadFile(fBytes, 0, UBound(fBytes), FLe.Name, "zxb@fjw")
            
126             Set objUp = Nothing
            
128             If Not FSo.FolderExists(mstrZippedFolder & "\bk") Then
130                 Call FSo.CreateFolder(mstrZippedFolder & "\bk")
                End If
132             If Not FSo.FileExists(mstrZippedFolder & "\bk" & "\" & FLe.Name) Then
134                 Call FSo.MoveFile(FLe.Path, mstrZippedFolder & "\bk" & "\" & FLe.Name)
                Else
136                 Call FSo.DeleteFile(FLe.Path)
                End If
            
138             DoEvents
            Next
    
        End If
    
140     Set Fles = Nothing
142     Set FSo = Nothing

End Sub
