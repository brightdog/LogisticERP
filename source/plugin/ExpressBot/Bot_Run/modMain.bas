Attribute VB_Name = "modMain"
Option Explicit
Private Declare Function URLDownloadToFile _
                Lib "urlmon" _
                Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Const SERVERREMOVEFAILEDLISTNAME As String = "RemoveFailed.list"

Private Sub Main()

        Dim FSo As Scripting.FileSystemObject

100     Set FSo = New Scripting.FileSystemObject

102     With FSo

104         If Not .FolderExists(App.Path & "\log") Then

106             .CreateFolder App.Path & "\log"

            End If

            

        End With

120     Set FSo = Nothing
        
122     frmMain.Show
End Sub

Public Function DownLoadFile(ByVal url As String, ByVal LocalFilename As String) As Boolean
    
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, url, LocalFilename, 0, 0)

    If lngRetVal = 0 Then

        DownLoadFile = True

    Else

        DownLoadFile = False

    End If

End Function

Public Function convertCRLF(ByRef value As String) As String
    convertCRLF = Replace(value, vbCr, Chr$(3))
    convertCRLF = Replace(convertCRLF, vbLf, Chr$(4))
End Function

Public Function restoreCRLF(ByRef value As String) As String
    restoreCRLF = Replace(value, Chr$(3), vbCr)
    restoreCRLF = Replace(restoreCRLF, Chr$(4), vbLf)
End Function
