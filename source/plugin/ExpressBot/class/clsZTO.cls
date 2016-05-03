VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetInfo(ByVal strExpressNO As String) As String

    Dim iWeb As clsXMLHTTPGetHtml
    Set iWeb = New clsXMLHTTPGetHtml

    iWeb.CharSet = "UTF-8"

    iWeb.URL = "http://www.zto.cn/GuestService/Bill?txtBill=" & strExpressNO

    'iWeb.PostData = "txtBill=" & strExpressNO 'txtBill=728499564190

    Dim strHtml As String

    Call iWeb.Send

    strHtml = iWeb.ReturnData
    strHtml = convertCRLF(strHtml)
    Debug.Print strHtml
    
    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    
    Dim Mc As VBScript_RegExp_55.MatchCollection
    Dim m As VBScript_RegExp_55.Match
    
    
    
    Reg.Global = True
    Reg.IgnoreCase = True
    Reg.Pattern = "<div class=""state"">(.*?)</ul>"
    
    Set Mc = Reg.Execute(strHtml)
    
    If Mc.Count = 1 Then
        

        
        'Reg.Pattern = "<div class=""\w*?"">(.*?)<a[^>]+>(.*?)</a>(.*?)<a[^>]+>(.*?)</a>(.*?)</div>.*?<div class=""time"">(\d\d\d\d-\d\d-\d\d \d\d\:\d\d\:\d\d)</div>"
        '不用这么详细，只纪录一个笼统的状态就可以了吧。
        Reg.Pattern = "<div class=""\w*?"">(.*?)</div>.*?<div class=""time"">(\d\d\d\d-\d\d-\d\d \d\d\:\d\d\:\d\d)</div>"
        Set Mc = Reg.Execute(Mc.Item(0).SubMatches(0))
        
        If Mc.Count > 0 Then
            Dim SB As clsStringBuilder
        
            Set SB = New clsStringBuilder
            
            Dim i As Integer
            
            For Each m In Mc
                
'                For i = 0 To 5
'
'                    SB.Append modMain.ConvertHTML(m.SubMatches(i)) & vbTab
'                Next
                
                
                '在这里设置CSV格式的原始数据，后续再根据模板拼接成写库的SQL
                SB.Append modMain.ConvertHTML(m.SubMatches(0)) & "|Tab|" & modMain.ConvertHTML(m.SubMatches(1)) & "|CrLf|"
                
                'SB.Append vbCrLf
            
            Next
            
            
            '这里需要将当次爬到的数据往服务器的数据接收页面提交，其余的事情交给服务器去做了。
            '包括删除旧数据，写入新数据等。2015-03-04
            
            'WriteSQL SB.toString
            Dim ObjUpload As clsUpLoadData
            Set ObjUpload = New clsUpLoadData
            Call ObjUpload.PostExpressData(gstrSite, gstrExpressNO, SB.toString)
            Set ObjUpload = Nothing
            
        Else
            WriteLog "*NOT MATCH DETAIL LIST"
        End If
    
    Else
        WriteLog "*NOT MATCH CONTENT"
    End If
    
End Function
