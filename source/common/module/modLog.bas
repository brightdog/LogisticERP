Attribute VB_Name = "modLog"
Option Explicit

Public Sub WriteLog(ByVal str As String, Optional ByVal logTime As Boolean = True) ', Optional ByRef frm As VB.Form = Null)
    On Error Resume Next
    Dim strLogFileName As String
    If strLogFileName = "" Then
    
        strLogFileName = "Log_" & Format(Date, "YYYY-MM-DD") & ".txt"
        
    End If

    Dim i As Integer

'    If frm Is Not Null Then
'
'        With frmMain.lstLog
'
'            If .ListCount > 100 Then
'                .Visible = False
'
'                For i = .ListCount To 30 Step -1
'
'                    .RemoveItem i - 1
'
'                    DoEvents
'
'                Next
'
'                .Visible = True
'            End If
'
'            If logTime Then
'
'                frmMain.lstLog.AddItem str & "<-- " & Now(), 0
'
'            Else
'
'                frmMain.lstLog.AddItem str, 0
'
'            End If
'
'            DoEvents
'            '.Refresh
'        End With
'
'    End If

    Open App.path & "\log\" & strLogFileName For Append As #1

    If logTime Then

        Print #1, str & "<-- " & Now()

    Else

        Print #1, str

    End If

    Close #1
    '.Visible = True

End Sub

Public Sub WriteCaption(ByRef frm As VB.Form, ByVal str As String)
    On Error Resume Next

    frm.Caption = str

End Sub

