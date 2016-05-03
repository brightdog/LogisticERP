<!--#include file="inc-connection.asp"-->

<%


Dim id
id = Request("id")


If id = ""  Then
	Response.Write("iERR")
    Response.End()
End If
If Not isNumeric(id)  Then
	Response.Write("iERRNO")
    Response.End()
End If


Dim sql, rs
Set cmd = Server.CreateObject("Adodb.Command")
Set rs = Server.CreateObject("Adodb.RecordSet")
cmd.ActiveConnection = adoConn
cmd.Commandtext = "Select * From PickupReceipt Where PickupReceiptID = ?"
cmd.CreateParameter


cmd.Parameters(0) = id


Set rs = cmd.Execute
If Not (rs.bof Or rs.EOF) Then

	Response.Write("OK")
Else

	Response.write("ERR")
End If

Response.End

rs.Close
Set rs = Nothing

%>