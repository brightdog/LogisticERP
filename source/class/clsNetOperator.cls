VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function SendData(ByRef dicData As Scripting.Dictionary) As Scripting.Dictionary

    '==================================================
    '2014-07-21����BY   wulf
    '�����ύ�����࣡
    '��Ҫע����ǣ��������ݴ��䣬����UTF-8��ʽ�����ݽ��ķ�չ����
    '��ǰ����Ҫ�ύ�����ݷ�װ��Dictionary���󣬱������ݴ���
    '��ʱ��û��÷��ص�����Ҫ��ʲô���ʹ��ݣ���ʱ��һ�㣬����String�ɣ�������ʱ��Ҳ�ܶ��󻯳�Dictionary
    '��������ҪԼ��һЩ���ݽṹ����������GET����POST��ɵɵ�ֲ������������
    '���������ֵ���󣬱���Ҫ�������漸��KEY��
    '   URL         GET��POST ����
    '   PostData    �����GET��ʽ���������ַ���
    '   Referer     ǿ����IIS�ˣ�������֤����ҳ�Ƿ�Ϊ�ض��ı��������У����Ƿ����ʡ�
    '��Ҫ��������ϢParse����Ӧ���ı�
    'Ȼ���ύ����������IISҳ����
    '�ȷ��������ؽ����Ϣ
    
    '�����������ص��ı���Ϣ����ͬ��������Dictionary��һ�𷵻س�ȥ��
    '�ֵ��������KEY��
    '   ReturnCode  ���������ص�״̬��
    '   ReturnData  ���������ص��ı�
    '   ERRDesc     ���ع����д������Ϣ���Լ�д�Ĵ���������
    '==================================================
    Dim dicResult As Scripting.Dictionary
    Set dicResult = New Scripting.Dictionary
    
    Set dicResult = dicData
    '���������Ĳ�����ֵ���ڲ������������еĲ������Ǹ����ڲ��ֵ������еġ�
    '�����ⲿ�����Ĵ�����
    
    If dicResult.Exists("ReturnCode") Then
    
        dicResult.Item("ReturnCode") = ""
    Else
        dicResult.Add "ReturnCode", ""
    End If
    
    If dicResult.Exists("ReturnData") Then
    
        dicResult.Item("ReturnData") = ""
    Else
        dicResult.Add "ReturnData", ""
    End If

    If dicResult.Exists("ERRDesc") Then
    
        dicResult.Item("ERRDesc") = ""
    Else
        dicResult.Add "ERRDesc", ""
    End If
    
    '����������һ���ֵ�������ֵ
    
    If checkisValid(dicResult) Then
        
        Dim iWeb As clsXMLHTTPGetHtml
        Set iWeb = New clsXMLHTTPGetHtml
        
        iWeb.URL = dicResult.Item("URL")
        iWeb.PostData = dicResult.Item("PostData")
        iWeb.Referer = dicResult.Item("Referer")
        Dim strResult As String
        
        Call iWeb.Send
        
        dicResult.Item("ReturnData") = iWeb.ReturnData
        dicResult.Item("ReturnCode") = iWeb.ReturnCode
        
        Set iWeb = Nothing
    Else
    
    End If
    Set SendData = dicResult
End Function

Private Function checkisValid(ByRef dicData As Scripting.Dictionary) As Boolean
    '������Ƚϴֲڣ���������Ƿ��ж�Ӧ��KEY����û�м��KEY�����ݣ��Լ����ݵĺϷ��ԣ�
    '�����ڲ�ʹ�õķ��������������ˣ�ûʱ�������ˣ�
    '�Ժ���ʱ���ٸ�ϸ�е㡣
    checkisValid = True
    Dim strKeys As String
    strKeys = "URL|PostData|Referer"
    
    Dim arrKeys() As String
    
    arrKeys = Split(strKeys, "|", -1, vbBinaryCompare)
    
    Dim i As Integer
    
    For i = 0 To UBound(arrKeys)
    
        If dicData.Exists(arrKeys(i)) Then
        
        Else
        
            checkisValid = False
            Exit For
        
        End If
    
    Next

End Function