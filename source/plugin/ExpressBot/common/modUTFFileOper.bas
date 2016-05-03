Attribute VB_Name = "modUTFFileOper"
Option Explicit
'1.VB写入utf-8文本文件
'
'复制内容到剪贴板程序代码 程序代码

Public Function UEFSaveTextFile(ByVal strPath As String, ByVal strContent As String, Optional ByVal bolCreate As Boolean = False) As Boolean
        '<EhHeader>
        On Error GoTo UEFSaveTextFile_Err
        '</EhHeader>
        On Error Resume Next
        Dim adoStream As ADODB.Stream
100     Set adoStream = New ADODB.Stream

102     With adoStream
104         .Type = adTypeText
106         .Mode = adModeReadWrite
108         .CharSet = "utf-8"
110         .Open
112         .Position = 0
114         .WriteText strContent

116         If bolCreate Then
118             .SaveToFile strPath, adSaveCreateNotExist
            Else
120             .SaveToFile strPath, adSaveCreateOverWrite
            End If

122         .Close
        End With

124     Set adoStream = Nothing
        '<EhFooter>
        Exit Function

UEFSaveTextFile_Err:
        WriteLog Err.Description & vbCrLf & _
           "in HotelPrice_Bot.modUTFFileOper.UEFSaveTextFile " & _
           "at line " & Erl
        Err.Clear
        Resume Next
        '</EhFooter>
End Function

'2.VB读取utf-8文本文件
'
'复制内容到剪贴板程序代码 程序代码

Public Function UEFLoadTextFile(ByVal strPath As String) As String
        '<EhHeader>
        On Error GoTo UEFLoadTextFile_Err
        '</EhHeader>

        Dim adoStream As ADODB.Stream
100     Set adoStream = New ADODB.Stream

102     With adoStream
104         .Type = adTypeText
106         .Mode = adModeReadWrite
108         .CharSet = "utf-8"
110         .Open
112         .LoadFromFile strPath
114         UEFLoadTextFile = .ReadText
116         .Close
        End With

118     Set adoStream = Nothing

        '<EhFooter>
        Exit Function

UEFLoadTextFile_Err:
        WriteLog Err.Description & vbCrLf & _
           "in HotelPrice_Bot.modUTFFileOper.UEFLoadTextFile " & _
           "at line " & Erl
        Err.Clear
        Resume Next
        '</EhFooter>
End Function
