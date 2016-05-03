VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetInfo(ByVal strExpressNO As String) As String

    Dim iWeb As clsXMLHTTPGetHtml
    Set iWeb = New clsXMLHTTPGetHtml

    iWeb.CharSet = "UTF-8"

    iWeb.URL = "http://q.sto.cn/?wen=" & strExpressNO & "&btnQuery=%u67e5%u8be2"
    Dim strHTML As String

    Call iWeb.Send

    strHTML = iWeb.ReturnData

    'Debug.Print strHTML
    
    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    
    Dim Mc As VBScript_RegExp_55.MatchCollection
    Dim m As VBScript_RegExp_55.Match
    
    Reg.Global = True
    Reg.IgnoreCase = True
    Reg.Pattern = "<input type=""hidden""[^>]+id=""(__\w+)"" value=""([^""]+)"""
    
    Set Mc = Reg.Execute(strHTML)
    
    If Mc.Count > 0 Then
    
        Dim strPostData As String
        
        strPostData = ""
        
        For Each m In Mc
        
            strPostData = strPostData & m.SubMatches(0) & "=" & Replace(Replace(Replace(m.SubMatches(1), "/", "%2F"), "=", "%3D"), "+", "%2B") & "&"
        
        Next
        
        iWeb.URL = "http://q.sto.cn"
        
        iWeb.PostData = strPostData & "wen=" & strExpressNO & "&btnQuery=%E6%9F%A5%E8%AF%A2" '&x=" & modGetRandomNum.GetRandomNum(2, 1, 47) & "&y=" & modGetRandomNum.GetRandomNum(2, 1, 23)
        
        Call iWeb.Send

        strHTML = iWeb.ReturnData
        
        Debug.Print strHTML
        
        Reg.Pattern = "<td width=""25%"">(\d\d\d\d.\d\d.\d\d \d\d\:\d\d\:\d\d)</td>[^<]+<td width=""75%"">(.*?)</td>"
        Set Mc = Reg.Execute(strHTML)
        
        If Mc.Count > 0 Then
            Dim SB As clsStringBuilder
        
            Set SB = New clsStringBuilder
            
            Dim i As Integer
            
            For Each m In Mc
                
                For i = 0 To 1
                
                    SB.Append modMain.ConvertHTML(m.SubMatches(i)) & vbTab
                Next
                
                SB.Append vbCrLf
            
            Next
            
            WriteSQL SB.toString
            
        Else
            WriteLog "*NOT MATCH DETAIL LIST"
        End If
    
    Else
        WriteLog "*NOT MATCH CONTENT"
    End If
    
End Function

