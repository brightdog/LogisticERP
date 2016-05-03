Attribute VB_Name = "modWriteFile"
Option Explicit

Public Sub WriteFile(ByRef strContent As String, Optional ByRef FileName As String = "")
    
    'On Error Resume Next
    
    Dim strLogFileName As String
    
    If FileName <> "" Then
    
        strLogFileName = FileName

    End If

    Dim i As Integer

    Dim iFile As Integer

    iFile = FreeFile()
    
    Open App.path & "\" & strLogFileName For Append As #iFile

    If strContent <> "" Then

        Print #iFile, strContent

    End If

    Close #iFile

    '.Visible = True
    
End Sub

