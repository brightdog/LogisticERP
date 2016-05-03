<%

If request.totalbytes > 0 Then

    formsize = request.totalbytes           ' 取二进制流字节长度
    formdata = request.binaryread(formsize)      ' 读取二进制流内容
    bncrlf = ChrB(13) & ChrB(10)
    datastart = InStrB(formdata, bncrlf & bncrlf) + 3 ' 取二进制流文件开始位置 (两个回车换行符)
    divider = LeftB(formdata, InStrB(formdata, bncrlf) - 1) ' 定义取二进制流 Field 分隔标记 (内容为二进制)
    dataend = InStrB(datastart, formdata, divider) - datastart ' 取二进制流文件部分结束位置
    '将文件信息保存到数据库
    'Call ImgToDb()        '将上传的图片以二进制保存到数据库中
    Call SaveTofile         '将上传的文件保存到服务器
End If

'

Sub SaveTofile() '将上传的文件保存到服务器
    '2.将获取的信息以二进制流文件存放 --- stm
    savepath = Server.mappath("images") & "\" '根据情况自己要先建立相应目录 或者开启fso自动建立
    Set strm = CreateObject("adodb.str" & "eam")

    With strm
        .Type = 1        ' 二进制模式
        .mode = 3        ' 指定打开模式为读写
        .open
        .write formdata        '写入二进制流内容
        '以文本模式读取数据,用于获得提交上来的文件路径及名称等信息
        .position = 0       '将游标指向数据首部
        .Type = 2        '以文本模式读取
        .Charset = "gb2312"     '设置中文编码

        formhead = .ReadText(datastart - 1) '读取表单头部内容
    End With

    '2.1获取上传的文件名称filename
    fullname = fRegExpSgl(formhead, True, True, True, "[\s\S]*filename\=""(.*?)""[\s\S]*", "$1")
    fname = Split(fullname, "\")
    FileName = fname(UBound(fname)) '获取到文件名
    Set fso = Server.CreateObject("Scripting.File" & "System" & "Object") '判断是否与本地盘文件重名,否则重命名 XXX(1).xxx

    If fso.FileExists(savepath & FileName) Then

        For i = 1 To 999
            fxname = Split(FileName, ".")
            Fn = Left(FileName, InStrRev(FileName, ".") - 1)
            Fnx = fxname(UBound(fxname))

            If Not fso.FileExists(savepath & Fn & "(" & i & ")." & Fnx) Then
                FileName = Fn & "(" & i & ")." & Fnx
                Exit For
            End If

        Next

    End If

    '3.从stm二进制流文件中获取有效信息 及 保存文件
    Set formstrm = CreateObject("adodb.str" & "eam")

    With formstrm
        .Type = 1       ' 二进制模式
        .mode = 3
        .open
        strm.position = datastart      ' 指定 stm 对象的起始位置, 以变量 bStart 的值为起始位置
        strm.copyTo formstrm, dataend   ' 拷贝 stm 二进制流至 fromStm 对象, 长度为 bEnd 变量的长度
        .SaveTofile (savepath & FileName), 2 ' 将信息保存到文件, 如果存在相同名称, 则覆盖
        .Close
    End With

    Set strm = Nothing
    Set formstrm = Nothing
    response.write "name=" & FileName
End Sub

Function fRegExpSgl(str, glb, igc, mtl, pt, rpt)
    Dim re
    Set re = New RegExp
    re.Global = glb
    re.ignoreCase = igc
    re.MultiLine = mtl
    re.Pattern = pt
    fRegExpSgl = re.Replace(str, rpt)
    Set re = Nothing
End Function

' Set objStream = Server.CreateObject("ADODB.Stream")
' objStream.Type = 1 ' adTypeBinary
' objStream.Open
' objStream.Write mydata
' objStream.SaveToFile Server.MapPath(GetFileName(strFileName)),2
' objStream.Close

%>

