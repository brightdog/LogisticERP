Attribute VB_Name = "modWriteLog"
Option Explicit
    
Public Sub CheckDirectory()

    Dim FSo As Scripting.FileSystemObject

    Set FSo = New Scripting.FileSystemObject

    With FSo

        If Not .FolderExists(App.Path & "\log") Then

            .CreateFolder App.Path & "\log"

        End If

        '108         If Not .FolderExists(App.Path & "\SQL") Then
        '
        '110             .CreateFolder App.Path & "\SQL"
        '
        '            End If

    End With
        
    Set FSo = Nothing
        
End Sub

Public Sub WriteLog(ByRef strLog As String, Optional ByRef bolLogTime As Boolean = True, Optional ByRef FileName As String = "")
    
    On Error Resume Next
    
    Dim strLogFileName As String
    
    If FileName <> "" Then
    
        strLogFileName = FileName

    End If
    
    If strLogFileName = "" Then
    
        strLogFileName = "Log_" & Format(Date, "YYYY-MM-DD") & ".txt"
        
    End If

    Dim i As Integer
    
    Dim iFile As Integer

    iFile = FreeFile()
    
    Open App.Path & "\log\" & strLogFileName For Append As #iFile

    If bolLogTime Then

        Print #iFile, strLog & "<-- " & Now()

    Else

        Print #iFile, strLog

    End If

    Close #iFile
    'frmMain.lblLog.Caption = strLog
    '.Visible = True

End Sub



