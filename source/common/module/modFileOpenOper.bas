Attribute VB_Name = "modFileOpenOper"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Public Sub OpenFileWithSysProgram(ByVal strFilePath As String)

    Dim result
    result = ShellExecute(0, vbNullString, strFilePath, vbNullString, vbNullString, SW_SHOWNORMAL)

    If result <= 32 Then
        MsgBox "��ʧ�ܣ�" & strFilePath, vbOKOnly, "ʧ����ʾ"
    End If

End Sub
