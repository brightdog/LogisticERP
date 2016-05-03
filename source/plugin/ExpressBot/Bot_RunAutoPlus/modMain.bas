Attribute VB_Name = "modMain"
Option Explicit
Private Declare Function GetComputerName _
                Lib "kernel32" _
                Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function URLDownloadToFile _
                Lib "urlmon" _
                Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Const SERVERREMOVEFAILEDLISTNAME As String = "RemoveFailed.list"

Public Function GetClientName() As String

    Dim tmpstrName As String * 255
    GetComputerName tmpstrName, 255
    GetClientName = Left(tmpstrName, InStr(1, tmpstrName, Chr(0)) - 1)

End Function


Private Sub Main()

        Dim Fso As Scripting.FileSystemObject

100     Set Fso = New Scripting.FileSystemObject

102     With Fso

104         If Not .FolderExists(App.Path & "\log") Then

106             .CreateFolder App.Path & "\log"

            End If

            

        End With

120     Set Fso = Nothing
        
122     frmMain.Show
End Sub



Public Sub ZipSqlFile(ByVal strSourcePath As String, ByVal strDestPath As String)
    Dim Fso As Scripting.FileSystemObject
    Set Fso = New Scripting.FileSystemObject

    If Fso.FileExists("C:\Program Files\7-Zip\7z.exe") Then
        Call DosPrint("""C:\Program Files\7-Zip\7z.exe"" a " & strDestPath & " " & strSourcePath, False)
    Else
        Call DosPrint("""C:\Program Files (x86)\7-Zip\7z.exe"" a " & strDestPath & " " & strSourcePath, False)
    End If
    
    Set Fso = Nothing
End Sub


Public Function convertCRLF(ByRef value As String) As String
    convertCRLF = Replace(value, vbCr, Chr$(3))
    convertCRLF = Replace(convertCRLF, vbLf, Chr$(4))
End Function

Public Function restoreCRLF(ByRef value As String) As String
    restoreCRLF = Replace(value, Chr$(3), vbCr)
    restoreCRLF = Replace(restoreCRLF, Chr$(4), vbLf)
End Function


