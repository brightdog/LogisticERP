VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbsFileUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''
'Desc:VBS文件上传类
'Author:shadu@foxmail.com + http://www.motobit.com/tips/detpg_uploadvbsie/ (2001 Antonin Foller, PSTRUH Software)
'Version:1.0
'Update:2010-12-28 19:00
'
'''''''''''''''''''''''''''''''

''' VBS文件上传类，二进制方式上传

Public c_strDestURL      ' 文件上传URL http://127.0.0./upload_file.php
Public c_strFileName     ' 要上传的本地文件名
Public c_strFieldName    ' 字段名，类似HTML表单Form中的input name
Public c_strBoundary     ' 文件上传Post数据包中的分隔符
Public c_strContentType  ' text/plain or image/pjpeg and so on "application/upload"
Public c_strResponseText ' 文件上传后，服务器返回的信息
Public c_boolPrepared    '
Public c_strErrMsg       ' 可能的错误信息

Public Sub Class_Initialize()
    c_strDestURL = "http://127.0.0.1/test/upload.php"
    c_strFileName = "c:\desktop.jpg"
    c_strContentType = "application/upload"
    c_strFieldName = "file"
    c_strBoundary = "---------------------------7da1c52160186"
    c_boolPrepared = False
End Sub
    
Public Sub Class_Terminate()
End Sub
    
''' 公共调用函数,文件上传
Public Function vbsUpload()
    Call CheckRequirements

    If c_boolPrepared Then
        UploadFile c_strDestURL, c_strFileName, c_strFieldName
    Else
        'WScript.Echo c_strErrMsg
    End If

End Function
    
''' 检查程序工作环境
Private Function CheckRequirements()
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If Not objFSO.FileExists(c_strFileName) Then
        c_strErrMsg = c_strErrMsg & vbCrLf & "文件不存在.."
    Else
        On Error Resume Next
        CreateObject "MSXML2.XMLHTTP"

        If Not Err = 0 Then
            c_strErrMsg = c_strErrMsg & vbCrLf & Err.Descriptiof
        Else
            c_boolPrepared = True
        End If
    End If

End Function
    
''' 文件上传
Private Function UploadFile(DestURL, FileName, FieldName)
    Dim FileContents, FormData, Boundary
    Boundary = c_strBoundary
    FileContents = GetFile(FileName) ' 二进制文件内容

    FormData = BuildFormData(FileContents, Boundary, FileName, FieldName)
    WinHTTPPostRequest DestURL, FormData, Boundary
End Function
    
''' WinHTTPPostRequest
Private Function WinHTTPPostRequest(URL, FormData, Boundary)
    Dim xmlhttp
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    On Error Resume Next
    xmlhttp.Open "POST", URL, False
    xmlhttp.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" + Boundary
    xmlhttp.Send FormData
    c_strResponseText = xmlhttp.responseText ' 服务端返回信息
    Set xmlhttp = Nothing
End Function

'''组合上传数据包 multipart/form-data document Header + Content
Private Function BuildFormData(FileContents, Boundary, FileName, FieldName)
    Dim FormData, Pre, Po, ContentType
    ContentType = c_strContentType
    FileName = Split(FileName, "\")(UBound(Split(FileName, "\")))
    'The two parts around file contents In the multipart-form data.
    Pre = "--" + Boundary + vbCrLf + mpFields(FieldName, FileName, ContentType)
    Po = vbCrLf + "--" + Boundary + "--" + vbCrLf
      
    'Build form data using recordset binary field
    Const adLongVarBinary = 205
    Dim RS: Set RS = CreateObject("ADODB.Recordset")
    RS.fields.Append "b", adLongVarBinary, Len(Pre) + LenB(FileContents) + Len(Po)
    RS.Open
    RS.AddNew
    Dim LenData
    'Convert Pre string value To a binary data
    LenData = Len(Pre)
    RS("b").AppendChunk (StringToMB(Pre) & ChrB(0))
    Pre = RS("b").GetChunk(LenData)
    RS("b") = ""
        
    'Convert Po string value To a binary data
    LenData = Len(Po)
    RS("b").AppendChunk (StringToMB(Po) & ChrB(0))
    Po = RS("b").GetChunk(LenData)
    RS("b") = ""
        
    'Join Pre + FileContents + Po binary data
    RS("b").AppendChunk (Pre)
    RS("b").AppendChunk (FileContents)
    RS("b").AppendChunk (Po)
    RS.Update

    FormData = RS("b")
    RS.Close
    BuildFormData = FormData
End Function
    
''Converts OLE string To multibyte string
'Private Function StringToMB(S)
'    Dim i, B
'
'    For i = 1 To Len(S)
'
'        If Asc(Mid(S, i, 1)) < 0 Then '汉字是双字节，得两个字符一起处理
'
'            B = B & ChrB(Asc(Mid(S, i, 1)) + 65536)
'
'        Else
'            B = B & ChrB(Asc(Mid(S, i, 1)))
'        End If
'
'    Next
'
'    StringToMB = B
'End Function

'把普通字符串转成二进制字符串
Function StringToMB(varstr)
    StringToMB = ""
    Dim i As Long
    Dim VarChar As String
    Dim varAsc As Long
    Dim varlow As String
    Dim varhigh As String
    
    For i = 1 To Len(varstr)
        VarChar = Mid(varstr, i, 1)
        varAsc = Asc(VarChar)

        If varAsc < 0 Then
            varAsc = varAsc + 65535
        End If

        If varAsc > 255 Then
            varlow = Left(HEX(Asc(VarChar)), 2)
            varhigh = Right(HEX(Asc(VarChar)), 2)
            StringToMB = StringToMB & ChrB("&H" & varlow) & ChrB("&H" & varhigh)
        Else
            StringToMB = StringToMB & ChrB(Asc(VarChar))
        End If

    Next

End Function

''' 组织HTTP头
Private Function mpFields(FieldName, FileName, ContentType)
    Dim MPTemplate 'template For multipart header
    MPTemplate = "Content-Disposition: form-data; name=""{field}"";" + _
       " filename=""{file}""" + vbCrLf + _
       "Content-Type: {ct}" + vbCrLf + vbCrLf
    Dim Out
    Out = Replace(MPTemplate, "{field}", FieldName)
    Out = Replace(Out, "{file}", FileName)
    mpFields = Replace(Out, "{ct}", ContentType)
End Function
    
''' 二进制载入文件内容
Private Function GetFile(FileName)
    Dim Stream: Set Stream = CreateObject("ADODB.Stream")
    Stream.Type = 1 'Binary
    Stream.Open
    Stream.LoadFromFile FileName
    GetFile = Stream.Read
    Stream.Close
End Function

''' ============================== 示例 =======================================
'Dim myUpload
'Set myUpload = New vbsFileUpload
'myUpload.c_strDestURL = "http://127.0.0.1/test/upload.php"      ' 必选
'myUpload.c_strFileName = "C:\test.jpg"                                   ' 必选
'myUpload.c_strFieldName = "file1"                                        ' 必选
'myUpload.c_strContentType = "image/pjpeg"                               ' 可选
'myUpload.vbsUpload()
'WScript.Echo myUpload.c_strResponseText
'WScript.Echo myUpload.c_strErrMsg
'Set myUpload = Nothing
'

''' ============================== 抓包 =======================================
'POST /test/upload_file.php HTTP/1.1
'Accept: image/gif, image/jpeg, image/pjpeg, image/pjpeg, application/x-shockwave-flash, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*
'Referer: http://127.0.0.1/test/upload.php
'Accept-Language: zh-cn
'User-Agent: Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; Trident/4.0; QQDownload 667; TencentTraveler 4.0)
'Content-Type: multipart/form-data; boundary=---------------------------7da2424be0662
'Accept-Encoding: gzip, deflate
'Host: 127.0.0.1
'Content-Length: 30837
'Connection: Keep-Alive
'Cache-Control: no-cache

'-----------------------------7da2424be0662
'Content-Disposition: form-data; name="file"; filename="kvm.JPG"
'Content-Type: image/pjpeg

'.............
'..
'-----------------------------7da2424be0662--

