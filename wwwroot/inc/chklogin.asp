<!--#include file="inc-connection.asp"-->
<!--#include file="common.asp"-->
<%


Dim aname, apass
uid = Request.Form("u")
pwd = Request.Form("p")

aname = uid
apass = pwd

If aname = "" or apass = "" Then
	Response.Write("iERR")
    Response.End()
End If

Dim sql, rs
Set cmd = Server.CreateObject("Adodb.Command")
Set rs = Server.CreateObject("Adodb.RecordSet")
cmd.ActiveConnection = adoConn
cmd.Commandtext = "Select * From UserInfo Where EngName = ? And PWD = ?"
cmd.CreateParameter
'cmd.CreateParameter

cmd.Parameters(0) = aname
cmd.Parameters(1) = apass
'response.write aname & ":" & apass
Set rs = cmd.Execute
If Not (rs.bof Or rs.EOF) Then

	
	Session("EmpID") = rs.Fields.Item("EmpID").Value
	Session("EngName") = rs.Fields.Item("EngName").Value & ""
	Session("ChsName") = rs.Fields.Item("ChsName").Value & ""


	Response.Cookies("EmpID") = rs.Fields.Item("EmpID").Value
	Response.Cookies("EngName") = rs.Fields.Item("EngName").Value & ""
	Response.Cookies("ChsName") = rs.Fields.Item("ChsName").Value & ""

	
	Response.Cookies("EmpID").Expires=Date+365
	Response.Cookies("EngName").Expires=Date+365
	Response.Cookies("ChsName").Expires=Date+365

	Session("chkLogin") = 1
	Response.Write("OK")
Else
	Session("chkLogin") = 0
	Response.write "ERR"
End If

Response.End

rs.Close
Set rs = Nothing



Public Function TransFormPWD(strPWD)
    Dim i
    i = 0
    Dim sb

    For i = 0 To Len(strPWD) - 1

        Select Case i Mod 6

            Case 0:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 8)

            Case 1:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 3)
     
            Case 2:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 9)

            Case 3:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 6)
         
            Case 4:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 2)

            Case 5:
                sb = sb & Chr(Asc(Mid(strPWD, i + 1, 1)) - 3)
                      
        End Select

    Next

    TransFormPWD = sb
End Function
%>