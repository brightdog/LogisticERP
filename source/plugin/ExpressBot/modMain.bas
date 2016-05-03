Attribute VB_Name = "modMain"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public gstrLogFileName As String
Public gstrSqlFileName As String
Public gstrLogDateTime As String
Public gbolisGetDataFromWeb As Boolean

Private Declare Function GetComputerName _
                Lib "kernel32" _
                Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public gstrComputerName As String



Public gbolExitJob As Boolean        '2013-12-30 用于将跨日的任务取消掉。当这个值为真时，直接退出当前Cls的执行，如果已经下载到数据，则将其打包，其余还来不及下载的列表，则丢弃掉。



Public gstrSite As String
Public gstrExpressNO As String

Public gUSERNAME As String
Public gHTTPURL As String
Public gSERVERIP As String
Public gSERVERPORT As String
'<a href="http\://(([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?)" title="(.*?)">


Public Function convertCRLF(ByRef Value As String) As String
    convertCRLF = Replace(Value, vbCr, Chr$(3))
    convertCRLF = Replace(convertCRLF, vbLf, Chr$(4))
End Function

Public Function restoreCRLF(ByRef Value As String) As String
    restoreCRLF = Replace(Value, Chr$(3), vbCr)
    restoreCRLF = Replace(restoreCRLF, Chr$(4), vbLf)
End Function

Sub Main()
        '<EhHeader>
        On Error GoTo Main_Err
        '</EhHeader>
        gUSERNAME = "BOT"
100     gbolExitJob = False
        Dim FSo As Scripting.FileSystemObject

102     Set FSo = New Scripting.FileSystemObject

104     With FSo

106         If Not .FolderExists(App.Path & "\log") Then

108             .CreateFolder App.Path & "\log"

            End If

110         If Not .FolderExists(App.Path & "\Result") Then

112             .CreateFolder App.Path & "\Result"

            End If



        End With

134     Set FSo = Nothing

        Call Init
142     frmMain.Show

        'End If

        '<EhFooter>


Main_Err:
        WriteLog Err.Number & "--" & Err.Description
        '136     frmODBCLogon.Show
        '</EhFooter>
End Sub
Private Function Init() As String

    Dim objConfig As clsWebConfig
    Set objConfig = New clsWebConfig
    
    gSERVERIP = objConfig.ReadProperty("ServerIP")
    gSERVERPORT = objConfig.ReadProperty("ServerPort")
    gHTTPURL = "http://" & gSERVERIP & ":" & gSERVERPORT & "/inc/"
    

    Set objConfig = Nothing

End Function

Public Function ConvertHTML(ByVal Content As String, Optional ByVal iLen As Integer = 0)
    Content = restoreCRLF(Content)
    Content = Replace(Content, vbTab, " ", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "&nbsp;", " ", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "&yen;", "￥", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "'", "`", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "&lt;", "<", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "&gt;", ">", 1, -1, vbBinaryCompare)
    Content = Replace(Content, Chr$(10), "", 1, -1, vbBinaryCompare)
    Content = Replace(Content, Chr$(9), "", 1, -1, vbBinaryCompare)
    Content = Replace(Content, Chr$(13), "", 1, -1, vbBinaryCompare)
    Content = Replace(Content, "<br>", vbCrLf, 1, -1, vbBinaryCompare)
    Content = Replace(Content, "<BR>", vbCrLf, 1, -1, vbBinaryCompare)
    Content = Replace(Content, "<br />", vbCrLf, 1, -1, vbBinaryCompare)
    Content = Replace(Content, "<br/>", vbCrLf, 1, -1, vbBinaryCompare)
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf, 1, -1, vbBinaryCompare)
    Content = Replace(Content, "&#20803;", "元", 1, -1, vbBinaryCompare)
    '以上的顺序是有讲究的，不可乱动！
    Dim i As Integer

    For i = 0 To 4
        Content = Replace(Content, "  ", "", 1, -1, vbBinaryCompare)
    Next

    Dim regTmp As VBScript_RegExp_55.RegExp

    Set regTmp = New VBScript_RegExp_55.RegExp
    regTmp.Global = True
    regTmp.MultiLine = True
    regTmp.IgnoreCase = True
    '======================= add by brightdog 去除页面中的干扰码
    regTmp.Pattern = "(<span[^>]*?display\s*?:\s*?none[^>]*?>[\w\W]*?<\/span>)"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "<font([^>]+)(0px|0pt)+([^>]*)>([\w\W]*?)<\/font>"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "<span[^>]*?font\s*?-\s*?size\s*?:\s*(0px|0pt)[^>]*?>([\w\W]*?)<\/span>"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "<script[^>]*?>([\w\W]*?)<\/script>"            '<span style="display:none">/ u6 i* t4 {1 Z. f5 m$ B. H" P1 u</span><br />
    Content = regTmp.Replace(Content, "")
    '=======================
    regTmp.Pattern = "(width\s*>\s*\d+)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "(height\s*>\s*\d+)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    Content = regTmp.Replace(Content, "")
     regTmp.Pattern = "(<em>.*?</em>)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    Content = regTmp.Replace(Content, "")
    
    regTmp.Pattern = "(<.*?[^>]>)"            '没办法，好多论坛中的图片带有JS脚本，只能这样试试看了。
    Content = regTmp.Replace(Content, "")
    regTmp.Pattern = "(<.*?>)"
    Content = regTmp.Replace(Content, "")
    
    If iLen > 0 Then
        Content = Left(Content, iLen)
    End If
    
    ConvertHTML = Trim$(Content)
    Set regTmp = Nothing
End Function

Public Sub WriteLog(ByVal str As String, Optional ByVal logTime As Boolean = True)
    On Error Resume Next
    
'    If gstrLogFileName = "" Then
'
'        gstrLogFileName = "Log_" & Format(Date, "YYYY-MM-DD") & ".txt"
'
'    End If
'
'    Dim i As Integer
'
'    With frmMain.lstLog
'
'        If .ListCount > 100 Then
'            .Visible = False
'
'            For i = .ListCount To 30 Step -1
'
'                .RemoveItem i - 1
'
'                DoEvents
'
'            Next
'
'            .Visible = True
'        End If
'
'        If logTime Then
'
'            frmMain.lstLog.AddItem str & "<-- " & Now(), 0
'
'        Else
'
'            frmMain.lstLog.AddItem str, 0
'
'        End If
'
'        DoEvents
'        '.Refresh
'        Dim iFile As Integer
'
'        'modUTFFileOper.UEFSaveTextFile App.path & "\log\" & gstrLogFileName & "Current.txt", str, True
'
'        iFile = VBA.FreeFile()
'        Open App.Path & "\log\" & gstrLogFileName For Append As #iFile
'
'        If logTime Then
'
'            Print #iFile, str & "<-- " & Now()
'
'        Else
'
'            Print #iFile, str
'
'        End If
'
'        Close #iFile
'        '.Visible = True
'    End With

End Sub

Public Sub WriteCaption(ByVal str As String)
    On Error Resume Next

    frmMain.Caption = str

End Sub

Public Sub WriteSQL(ByRef strSql As String, _
                    Optional ByRef SqlFileName As String = "")
    On Error Resume Next
    Dim strFileName As String
    Dim iFile       As Integer

    If SqlFileName <> "" Then
        strFileName = SqlFileName
    Else
        strFileName = gstrSqlFileName
    End If

    Dim i As Integer
    
    iFile = FreeFile()

    Open App.Path & "\Result\" & strFileName For Append As #iFile

    Print #iFile, strSql

    Close #iFile

End Sub


Public Sub WriteFile(ByRef strFileContent As String, ByRef strFileName As String, Optional ByRef gstrLogDateTime As String)
    On Error Resume Next

    Dim iFile As Integer
    

    Dim i As Integer
    
    iFile = FreeFile()

    Open App.Path & "\log\" & strFileName For Append As #iFile

    Print #iFile, strFileContent

    Close #iFile

End Sub



Public Function isNeedLogSource(ByRef strSiteName As String) As Boolean
        '<EhHeader>
        On Error GoTo isNeedLogSource_Err
        '</EhHeader>

        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
    
102     If FSo.FileExists(App.Path & "\NeedLogSource_" & strSiteName) Then
104         isNeedLogSource = True
        Else

106         isNeedLogSource = False
        End If
    
108     Set FSo = Nothing
        '<EhFooter>
        Exit Function

isNeedLogSource_Err:
        WriteLog Err.Description & vbCrLf & _
           "in HotelPrice_Bot_ADSL.modMain.isNeedLogSource " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function GetRndNum(Optional ByVal iStart As Single = 0, Optional ByVal iEnd As Single = 1) As Single

    Call Randomize
    
    GetRndNum = (Rnd * (iEnd - iStart)) + iStart + 0#

End Function

Public Function ReadDebugFiletoVar(ByVal FilePath As String) As String


    Dim FSo As Scripting.FileSystemObject
    Set FSo = New Scripting.FileSystemObject
    Dim TS As Scripting.TextStream
    
    Set TS = FSo.OpenTextFile(FilePath, ForReading, False)
    
    ReadDebugFiletoVar = TS.ReadAll

    Set FSo = Nothing
End Function

Public Function ReadRedailIntervalFile(ByVal strPath As String) As Integer
        '<EhHeader>
        On Error GoTo ReadRedailIntervalFile_Err
        '</EhHeader>

        Dim iFile  As Integer
        Dim strTmp As String
100     iFile = FreeFile()
102     Open App.Path & "\Config\" & strPath For Input As #iFile
104     Line Input #iFile, strTmp
106     Close #iFile
    
108     If IsNumeric(strTmp) Then
110         ReadRedailIntervalFile = strTmp
        Else
112         ReadRedailIntervalFile = -1
        End If

        '<EhFooter>
        Exit Function

ReadRedailIntervalFile_Err:
        Err.Clear
        ReadRedailIntervalFile = -1
        '</EhFooter>
End Function


Public Function CollectionToString(ByRef col As VBA.Collection, Optional ByVal strSplit As String = vbTab) As String

    Dim i As Long
    Dim strResult As String
    
    For i = 1 To col.Count
    
        strResult = strResult & col.Item(i) & strSplit
    
    
    Next

    CollectionToString = strResult
End Function



Public Function CheckisOtheruse() As String
        '<EhHeader>
        On Error GoTo CheckisOtheruse_Err
        '</EhHeader>


        Dim FSo As Scripting.FileSystemObject
100     Set FSo = New Scripting.FileSystemObject
    
102     If FSo.FileExists(App.Path & "\Otheruse.state") Then
            CheckisOtheruse = FSo.OpenTextFile(App.Path & "\Otheruse.state").ReadLine
        Else

106         CheckisOtheruse = ""
        End If
    
108     Set FSo = Nothing


        '<EhFooter>
        Exit Function

CheckisOtheruse_Err:
        WriteLog "HotelPrice_Bot_ADSL.CheckisOtheruse at line " & Erl
        CheckisOtheruse = ""
        '</EhFooter>
End Function

Public Function FormatMoneySymb(ByVal strSymb As String) As String

    Select Case strSymb
    
        Case "$"
            strSymb = "USD"

        Case "&yen;"
            strSymb = "CNY"

        Case "CNY"

        Case "HK$"

        Case "MOP"

        Case "SGD"
    
        Case Else
            strSymb = "ERR"
    
    End Select


    FormatMoneySymb = strSymb
End Function

Public Function MappingBreakFast(ByVal strDesc As String, ByVal strWebSite As String) As String
        '<EhHeader>
        On Error GoTo MappingBreakFast_Err
        '</EhHeader>

        Dim FSo As Scripting.FileSystemObject
    
100     Set FSo = New Scripting.FileSystemObject
    
        Dim strResult As String
102     strResult = strDesc
    
        Dim strFileName As String
    
104     strFileName = strWebSite & "_BreakFast_Mapping.config"
        Dim strFileContent  As String
    
106     strFileContent = FSo.OpenTextFile(App.Path & "\" & strFileName).ReadAll
        Dim arrLine() As String
    
108     arrLine = Split(strFileContent, vbCrLf, -1, vbBinaryCompare)
    
        Dim i As Integer
    
110     For i = 0 To UBound(arrLine)
    
112         If Trim(arrLine(i)) <> "" Then
        
                Dim arrTmp() As String
            
114             arrTmp = Split(arrLine(i), vbTab, -1, vbBinaryCompare)
            
116             If UBound(arrTmp) = 1 Then
            
118                 If arrTmp(0) = strDesc Then
                
120                     strResult = arrTmp(1)
                        Exit For
                    End If
            
                Else
            
122                 WriteLog "参数错误被丢弃：" & strWebSite & ":" & arrLine(i)
            
                End If
        
            End If
    
        Next

124     MappingBreakFast = strResult

        '<EhFooter>
        Exit Function

MappingBreakFast_Err:
        WriteLog Err.Description & vbCrLf & _
           "in HotelPrice_Bot_ADSL.modMain.MappingBreakFast " & _
           "at line " & Erl
        MappingBreakFast = strResult

        '</EhFooter>
End Function

Public Sub ZipSqlFile(ByVal strSourcePath As String, ByVal strDestPath As String)
    Dim FSo As Scripting.FileSystemObject
    Set FSo = New Scripting.FileSystemObject

    If FSo.FileExists("C:\Program Files\7-Zip\7z.exe") Then
        Call DosPrint("""C:\Program Files\7-Zip\7z.exe"" a " & strDestPath & " -ppassword -mhe " & strSourcePath)
    Else
        Call DosPrint("""C:\Program Files (x86)\7-Zip\7z.exe"" a " & strDestPath & " -ppassword -mhe " & strSourcePath)
    End If
    
    Set FSo = Nothing
End Sub

Public Function ReadSleepTime(ByVal strPath As String) As String

        On Error GoTo Err:
        Dim iFile As Integer

100     iFile = VBA.FreeFile()
        Dim strTmp As String
102     Open App.Path & "\Config\" & strPath For Input As #iFile
104     Line Input #iFile, strTmp
106     Close #iFile
    
108     If IsNumeric(strTmp) Then
110         ReadSleepTime = strTmp
        Else
112         ReadSleepTime = 0
        End If
        Exit Function
Err:
114     ReadSleepTime = 0
End Function

Public Function ReadJSContent(ByVal strFileName As String) As String

        '"Qunar_Decode_Core.txt"

        On Error GoTo Err:
        
        ReadJSContent = modUTFFileOper.UEFLoadTextFile(App.Path & "\AdditionalParam\" & strFileName)
        
'        Dim iFile As Integer
'
'100     iFile = VBA.FreeFile()
'        Dim strTmp As String
'102     Open App.path & "\AdditionalParam\" & strFileName For Input As #iFile
'
'        Do While Not EOF(iFile)
'104         Line Input #iFile, strTmp
'            ReadJSContent = ReadJSContent & strTmp & vbCrLf
'        Loop
'
'106     Close #iFile

110

        Exit Function
Err:
114     ReadJSContent = ""

End Function

Public Function convertResult(ByVal CompCode As String, ByVal ExpressNO As String, ByVal strResult As String) As String

    Dim arrLine() As String
    arrLine = Split(strResult, vbCrLf, -1, vbBinaryCompare)
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    SB.Value = ""
    Dim i As Integer

    For i = 0 To UBound(arrLine)
    
        If arrLine(i) <> "" Then
        
            Dim arrCell() As String
            
            arrCell = Split(arrLine(i), vbTab, 2, vbBinaryCompare)

            If UBound(arrCell) = 1 Then
            
                SB.Append CompCode & vbTab & ExpressNO & vbTab & arrCell(1) & vbTab & arrCell(0) & vbCrLf
            
            End If
        
        End If
    
    Next
    
    convertResult = SB.toString
    Set SB = Nothing
End Function
