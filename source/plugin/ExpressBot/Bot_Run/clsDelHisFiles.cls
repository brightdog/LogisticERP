VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDelHisFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub DeleteHistoryFiles(ByVal strPath As String, ByVal intDays As Integer)

    Dim FSo As Scripting.FileSystemObject
    Set FSo = New Scripting.FileSystemObject
    'Dim strPath As String

    'strPath = App.Path & "\SQL"
    If FSo.FolderExists(strPath) Then
    
        Call DelHistoryFile_Core(FSo.GetFolder(strPath), intDays)
    
    End If
    
    'strPath = App.Path & "\log"
    'If FSo.FolderExists(strPath) Then
    
    '    Call DelHistoryFile_Core(FSo.GetFolder(strPath), intDays)
    '
    'End If
End Sub

Private Sub DelHistoryFile_Core(ByRef Fld As Scripting.Folder, ByVal Days As Integer)
        '<EhHeader>
        On Error GoTo DelHistoryFile_Core_Err
        '</EhHeader>

        Dim dtNowDate As Date
100     dtNowDate = VBA.Date()
    
        Dim iFile As Scripting.File
    
102     For Each iFile In Fld.Files
    
            Dim dtFileDate As Date
104         dtFileDate = iFile.DateCreated
    
106         If DateDiff("d", dtFileDate, dtNowDate) > Days Then
        
108             Call iFile.Delete(True)
            
            End If
            DoEvents
        Next
    

        '<EhFooter>
        Exit Sub

DelHistoryFile_Core_Err:
        WriteLog Err.Description & vbCrLf & _
               "in DelHistoryFile_Core " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

