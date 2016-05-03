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


Dim strSql, Rst
Set Rst = Server.CreateObject("Adodb.RecordSet")

strSql = "Select * From tblOutWarehouseReceipt Where OutWarehouseReceiptID = " & SqlSafe(id)

Rst.Open strSql, adoConn, 0, 1

If Rst.Eof Then
	Response.Write("{""ERR"":""OWRIDERR""}")
    Response.End()
End If

Dim strToWarehouseID

strToWarehouseID = Rst.Fields.Item("WarehouseID").Value
Rst.Close

Set cmd = Server.CreateObject("Adodb.Command")
cmd.ActiveConnection = adoConn
cmd.Commandtext = "Select ExpressNO From tblInventory Where isValued = 1 And ToWarehouseID = " & strToWarehouseID & " And InventoryState='ÒÑÈë¿â' And  OutWarehouseReceiptID = ?"
cmd.CreateParameter


cmd.Parameters(0) = id


Set Rst = cmd.Execute
If Not (Rst.bof Or Rst.EOF) Then

	Response.Write MakeJsonFromRst(Rst, 0, 0, 0, 0)
Else

	Response.write "{""ERR"":""NOT FOUND""}"
End If

Response.End

Rst.Close
Set Rst = Nothing

%>