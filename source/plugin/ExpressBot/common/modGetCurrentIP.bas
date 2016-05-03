Attribute VB_Name = "modGetCurrentIP"
Option Explicit

Public Function GetCurrentIP() As String

    Dim iWeb As clsXMLHTTPGetHtml
    Set iWeb = New clsXMLHTTPGetHtml
    iWeb.CharSet = "GB2312"
    iWeb.URL = "http://1111.ip138.com/ic.asp"
    
    Dim Reg As VBScript_RegExp_55.RegExp
    Set Reg = New VBScript_RegExp_55.RegExp
    
    Reg.Pattern = "(\d+\.\d+\.\d+\.\d+)"
    Dim Mc As VBScript_RegExp_55.MatchCollection
    Call iWeb.Send

    Set Mc = Reg.Execute(iWeb.ReturnData)
    
    If Mc.Count > 0 Then
    
        GetCurrentIP = Mc.Item(0).SubMatches(0)
    
    Else
    
        GetCurrentIP = ""
    
    End If
    
    Set iWeb = Nothing
    Set Mc = Nothing
    Set Reg = Nothing

End Function
