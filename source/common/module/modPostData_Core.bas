Attribute VB_Name = "modPostData_Core"
Option Explicit

Public Function PostData(ByVal strURL As String, Optional ByVal strPostData As String = "", Optional ByVal TimeOut As Integer = 10) As String

    Dim iWeb As clsXMLHTTPGetHtml
    Set iWeb = New clsXMLHTTPGetHtml
'    Dim objTransPRD As clsTransformPWD
'    Set objTransPRD = New clsTransformPWD

    If Left(strPostData, 1) = "{" Then
        strPostData = Left(strPostData, Len(strPostData) - 1)
        strPostData = strPostData & ",""CreateEmp"":""" & gUSERNAME & """}"
    Else
    
        strPostData = strPostData & "&CreateEmp=" & gUSERNAME
    End If

    Debug.Print "modPostData_Core.PostData_URL:" & gHTTPURL & strURL

    If strPostData <> "" Then
        Debug.Print "modPostData_Core.PostData_POST:" & strPostData
    End If

    iWeb.URL = gHTTPURL & strURL
    iWeb.PostData = strPostData
    iWeb.CharSet = "UTF-8"
    iWeb.TimeOut = TimeOut
    Call iWeb.Send
    
    PostData = iWeb.ReturnData
    Debug.Print "modPostData_Core.PostData_ReturnData:" & PostData
    'Set objTransPRD = Nothing
    Set iWeb = Nothing
End Function

