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
    '2014-07-21创建BY   wulf
    '数据提交核心类！
    '需要注意的是：所有数据传输，均以UTF-8方式！兼容今后的发展方向
    '提前将需要提交的数据封装成Dictionary对象，便于数据传输
    '暂时还没想好返回的数据要用什么类型传递，暂时简单一点，就用String吧，反正到时候也能对象化成Dictionary
    '在这里需要约定一些数据结构，否则将来连GET还是POST都傻傻分不清楚。。。。
    '传进来的字典对象，必须要包括下面几个KEY：
    '   URL         GET和POST 共用
    '   PostData    如果是GET方式，则留空字符串
    '   Referer     强制在IIS端，必须验证引用页是否为特定文本，否则判：“非法访问”
    '先要将以上信息Parse成相应的文本
    '然后提交到服务器的IIS页面上
    '等服务器返回结果信息
    
    '将服务器返回的文本信息，连同传进来的Dictionary，一起返回出去！
    '字典对象添加KEY：
    '   ReturnCode  服务器返回的状态码
    '   ReturnData  服务器返回的文本
    '   ERRDesc     返回过程中错误的信息，自己写的错误描述！
    '==================================================
    Dim dicResult As Scripting.Dictionary
    Set dicResult = New Scripting.Dictionary
    
    Set dicResult = dicData
    '将传进来的参数赋值给内部对象，往后所有的操作都是根据内部字典对象进行的。
    '保持外部参数的纯净性
    
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
    
    '以上先清理一下字典对象里的值
    
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
    '这里检测比较粗糙，仅检测了是否有对应的KEY，并没有检查KEY的内容，以及内容的合法性！
    '反正内部使用的方法，就先这样了，没时间完善了！
    '以后有时间再搞细有点。
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
