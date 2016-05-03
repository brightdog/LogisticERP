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
	Response.Write("iERR")
    Response.End()
End If
If Not isNumeric(id)  Then
	Response.Write("iERRNO")
    Response.End()
End If


Dim sql, rst
Set cmd = Server.CreateObject("Adodb.Command")
Set rst = Server.CreateObject("Adodb.RecordSet")
cmd.ActiveConnection = adoConn
cmd.Commandtext = "Select * From PickupReceipt Where PickupReceiptID = ?"
cmd.CreateParameter


cmd.Parameters(0) = id


Set rst = cmd.Execute
If Not (rst.bof Or rst.EOF) Then

	Response.Write MakeJsonFromRst(Rst, 0, 0, 0, 0)
Else

	Response.write("{""ERR"":""NOT EXIST""}")
End If

Response.End

rst.Close
Set rs = Nothing

%>