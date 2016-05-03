<!--#include file="inc-connection.asp"-->
<!--#include file="common.asp"-->
<!--#include file="json.inc"-->
<!--#include file="clsstringbuilder.inc"-->
<%

Dim strData
strData = Request.Form
'Response.write strdata
'Response.write "<br>===<br>"

Dim dic
Set dic = toObject(strData)

Dim id

id = dic.id


If id = ""  Then
	Response.Write("{""ERR"":""IDERR""}")
    Response.End()
End If
If Not isNumeric(id)  Then
	Response.Write("{""ERR"":""IDERRNO""}")
    Response.End()
End If


Dim sql, rs
Set cmd = Server.CreateObject("Adodb.Command")
Set rs = Server.CreateObject("Adodb.RecordSet")
cmd.ActiveConnection = adoConn
cmd.Commandtext = "Select ExpressNO From tblOrder Where PickupReceiptID = ?"
cmd.CreateParameter


cmd.Parameters(0) = id


Set rs = cmd.Execute
If Not (rs.bof Or rs.EOF) Then

	Response.Write MakeJsonFromRst(rs, 0, 0, 0, 0)
Else

	Response.write "{""ERR"":""NOT FOUND""}"
End If

Response.End

rs.Close
Set rs = Nothing

%>