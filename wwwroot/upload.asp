<%

If request.totalbytes > 0 Then

    formsize = request.totalbytes           ' ȡ���������ֽڳ���
    formdata = request.binaryread(formsize)      ' ��ȡ������������
    bncrlf = ChrB(13) & ChrB(10)
    datastart = InStrB(formdata, bncrlf & bncrlf) + 3 ' ȡ���������ļ���ʼλ�� (�����س����з�)
    divider = LeftB(formdata, InStrB(formdata, bncrlf) - 1) ' ����ȡ�������� Field �ָ���� (����Ϊ������)
    dataend = InStrB(datastart, formdata, divider) - datastart ' ȡ���������ļ����ֽ���λ��
    '���ļ���Ϣ���浽���ݿ�
    'Call ImgToDb()        '���ϴ���ͼƬ�Զ����Ʊ��浽���ݿ���
    Call SaveTofile         '���ϴ����ļ����浽������
End If

'

Sub SaveTofile() '���ϴ����ļ����浽������
    '2.����ȡ����Ϣ�Զ��������ļ���� --- stm
    savepath = Server.mappath("images") & "\" '��������Լ�Ҫ�Ƚ�����ӦĿ¼ ���߿���fso�Զ�����
    Set strm = CreateObject("adodb.str" & "eam")

    With strm
        .Type = 1        ' ������ģʽ
        .mode = 3        ' ָ����ģʽΪ��д
        .open
        .write formdata        'д�������������
        '���ı�ģʽ��ȡ����,���ڻ���ύ�������ļ�·�������Ƶ���Ϣ
        .position = 0       '���α�ָ�������ײ�
        .Type = 2        '���ı�ģʽ��ȡ
        .Charset = "gb2312"     '�������ı���

        formhead = .ReadText(datastart - 1) '��ȡ��ͷ������
    End With

    '2.1��ȡ�ϴ����ļ�����filename
    fullname = fRegExpSgl(formhead, True, True, True, "[\s\S]*filename\=""(.*?)""[\s\S]*", "$1")
    fname = Split(fullname, "\")
    FileName = fname(UBound(fname)) '��ȡ���ļ���
    Set fso = Server.CreateObject("Scripting.File" & "System" & "Object") '�ж��Ƿ��뱾�����ļ�����,���������� XXX(1).xxx

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

    '3.��stm���������ļ��л�ȡ��Ч��Ϣ �� �����ļ�
    Set formstrm = CreateObject("adodb.str" & "eam")

    With formstrm
        .Type = 1       ' ������ģʽ
        .mode = 3
        .open
        strm.position = datastart      ' ָ�� stm �������ʼλ��, �Ա��� bStart ��ֵΪ��ʼλ��
        strm.copyTo formstrm, dataend   ' ���� stm ���������� fromStm ����, ����Ϊ bEnd �����ĳ���
        .SaveTofile (savepath & FileName), 2 ' ����Ϣ���浽�ļ�, ���������ͬ����, �򸲸�
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

