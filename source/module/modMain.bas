Attribute VB_Name = "modMain"
'2014-11-15 出库和入库逻辑判断还有问题。双击运单号列表框的时候，可能还有BUG，需要重新整理一下思路。派送模块还没做。

Option Explicit

Const btnOK As Integer = 1
Const btnCancel As Integer = 2
Const btnSave As Integer = 4
Public gSERVERIP As String
Public gSERVERPORT As String
Public gHTTPURL As String
Public gUSERNAME As String

Public gHeight As Long
Public gWidth As Long
Public gLeft As Long    '不好用，关键在窗体被拖动之后，这个值找不到刷新的办法。。。。
Public gTop As Long     '理由同上

Public gdicDBConfig As Scripting.Dictionary
Public gdicLocation As Scripting.Dictionary


Public gdicTitleMapping As Scripting.Dictionary

'==== 全局常量，只能放在这里====

Public Sub Main()
    Call modAppPath.Init_Path
    Call Init
    
    If InStr(1, VBA.Command, "[TEST]") > 0 Then
        'IN TEST MODE
        frmTest.Show
    Else
        'IN PRODUCT MODE
        WriteLog "login_show"
        frmLogin.Show
        'frmSplash.Show
        'frmMain.Show
    End If

End Sub

Private Function Init() As String

    Dim objConfig As clsWebConfig
    Set objConfig = New clsWebConfig
    
    gSERVERIP = objConfig.ReadProperty("ServerIP")
    gSERVERPORT = objConfig.ReadProperty("ServerPort")
    gHTTPURL = "http://" & gSERVERIP & ":" & gSERVERPORT & "/inc/"
    
    Dim obj As clsXMLHTTPGetHtml
    Set obj = New clsXMLHTTPGetHtml
    
    Dim strHtml As String
    
    obj.URL = gHTTPURL & "getdbfieldsinfo.asp"
    
    Call obj.Send
    
    strHtml = obj.ReturnData
    
    Dim iFile As Integer
    Dim objDE As clsDE
    Set objDE = New clsDE

    iFile = VBA.FreeFile()
    '==============2014-08-18==================================
    '目前为不判断版本，每次都直接从服务器取一份数据库表结构信息。
    '为兼容今后的版本控制，及载入速度，先加密存放到本地。
    '但目前的本地文件，并没有起作用，仅存储而已。
    '==========================================================
    Open App.path & "\Config\DB.Config" For Output As #iFile
    Print #iFile, objDE.EnCode(strHtml)
    Close #iFile
    
    Set gdicDBConfig = JSON.Parse(strHtml)
    'gdicDBConfig这个全局变量，在程序中的需要访问数据库的部分，会用做输入验证
    '以及确定字段的类型是否为文本：SQL拼接过程中，是否需要加单引号!!
    '配合另一个字段MAPPING文件，做为查询逻辑用。
    '比如日期部门的范围，由于输入框是2个，但是字段是一个，并且是>=...and <=...的结构。
    '既包含 a 又包含 b !!
    
    'Debug.Print objDE.Decode(objDE.EnCode(strHtml))
    
    Dim Fso As Scripting.FileSystemObject
    Set Fso = New Scripting.FileSystemObject
    Dim strFileContent As String

    If Not Fso.FileExists(App.path & "\Config\Location.Config") Then
    
        obj.URL = gHTTPURL & "location.json"
        Call obj.Send

        strHtml = obj.ReturnData
        strFileContent = strHtml
        Call Fso.CreateTextFile(App.path & "\Config\Location.Config").Write(objDE.EnCode(strFileContent))
    Else
        
        strFileContent = Fso.OpenTextFile(App.path & "\Config\Location.Config").ReadAll
    End If
    
    Set gdicLocation = JSON.Parse(objDE.Decode(strFileContent))
    
    
    
    '
    '    iFile = VBA.FreeFile()
    '    '==============2014-08-27 ==================================
    '    '目前为不判断版本，每次都直接从服务器取一份3级城市数据信息。
    '    '为兼容今后的版本控制，及载入速度，先加密存放到本地。
    '    '客户以及订单输入的时候，需要用到。
    '    '解密比加密快哈~~所以服务器端存放的是加密的版本，然后下载到本地存储密文后，解密直接使用。
    '    '加密慢，是因为拼接字符串引起的！！！换成StringBuilder之后秒开啊！！
    '    '不过服务器上为了安全考虑，存放密文比较好。
    '    '为了速度考虑，还是存放明文吧，一个90KB，一个200KB。。。
    '    '总结:加密也是要有代价的！
    '    '==========================================================
    '    On Error Resume Next
    '    Debug.Print Timer
    '    Dim strFileContent As String
    '    strFileContent = strHtml
    '    Debug.Print strFileContent
    '    Open App.path & "\Config\Location.Config" For Output As #iFile
    '    Print #iFile, objDE.EnCode(strFileContent)
    '    Close #iFile
    '
    '    Debug.Print Timer

    '    Set gdicLocation = JSON.Parse(strFileContent)
    '    Debug.Print ":" & Timer
    'Debug.Print objDE.Decode(objDE.Encode(strHtml))
    obj.URL = gHTTPURL & "titlemapping\title.json"
    
    Call obj.Send
    
    strHtml = obj.ReturnData
    
    Set gdicTitleMapping = modCommon.GetTitleMapping(strHtml)




    Set objDE = Nothing
    Set obj = Nothing
End Function
