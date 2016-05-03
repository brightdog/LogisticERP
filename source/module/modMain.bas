Attribute VB_Name = "modMain"
'2014-11-15 ���������߼��жϻ������⡣˫���˵����б���ʱ�򣬿��ܻ���BUG����Ҫ��������һ��˼·������ģ�黹û����

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
Public gLeft As Long    '�����ã��ؼ��ڴ��屻�϶�֮�����ֵ�Ҳ���ˢ�µİ취��������
Public gTop As Long     '����ͬ��

Public gdicDBConfig As Scripting.Dictionary
Public gdicLocation As Scripting.Dictionary


Public gdicTitleMapping As Scripting.Dictionary

'==== ȫ�ֳ�����ֻ�ܷ�������====

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
    'ĿǰΪ���жϰ汾��ÿ�ζ�ֱ�Ӵӷ�����ȡһ�����ݿ��ṹ��Ϣ��
    'Ϊ���ݽ��İ汾���ƣ��������ٶȣ��ȼ��ܴ�ŵ����ء�
    '��Ŀǰ�ı����ļ�����û�������ã����洢���ѡ�
    '==========================================================
    Open App.path & "\Config\DB.Config" For Output As #iFile
    Print #iFile, objDE.EnCode(strHtml)
    Close #iFile
    
    Set gdicDBConfig = JSON.Parse(strHtml)
    'gdicDBConfig���ȫ�ֱ������ڳ����е���Ҫ�������ݿ�Ĳ��֣�������������֤
    '�Լ�ȷ���ֶε������Ƿ�Ϊ�ı���SQLƴ�ӹ����У��Ƿ���Ҫ�ӵ�����!!
    '�����һ���ֶ�MAPPING�ļ�����Ϊ��ѯ�߼��á�
    '�������ڲ��ŵķ�Χ�������������2���������ֶ���һ����������>=...and <=...�Ľṹ��
    '�Ȱ��� a �ְ��� b !!
    
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
    '    'ĿǰΪ���жϰ汾��ÿ�ζ�ֱ�Ӵӷ�����ȡһ��3������������Ϣ��
    '    'Ϊ���ݽ��İ汾���ƣ��������ٶȣ��ȼ��ܴ�ŵ����ء�
    '    '�ͻ��Լ����������ʱ����Ҫ�õ���
    '    '���ܱȼ��ܿ��~~���Է������˴�ŵ��Ǽ��ܵİ汾��Ȼ�����ص����ش洢���ĺ󣬽���ֱ��ʹ�á�
    '    '������������Ϊƴ���ַ�������ģ���������StringBuilder֮���뿪������
    '    '������������Ϊ�˰�ȫ���ǣ�������ıȽϺá�
    '    'Ϊ���ٶȿ��ǣ����Ǵ�����İɣ�һ��90KB��һ��200KB������
    '    '�ܽ�:����Ҳ��Ҫ�д��۵ģ�
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
