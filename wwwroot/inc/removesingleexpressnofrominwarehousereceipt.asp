<!--#include file="inc-connection.asp"-->
<!--#include file="common.asp"-->
<!--#include file="json.inc"-->
<!--#include file="clsstringbuilder.inc"-->
<%


Dim strData
strData = CStr(Request.Form)

Dim dic
Set dic = toObject(strData)
If strData = ""  Then
	Response.Write "{""ERR"":""DATA is NULL""}"
	Response.End
End If
'Response.end

Dim strExpressNO
Dim strOutWarehouseReceiptID
strExpressNO = dic.ExpressNO
strOutWarehouseReceiptID = dic.OutWarehouseReceiptNO

'Response.write strExpressNO & "::" & strOutWarehouseReceiptID

If strOutWarehouseReceiptID= ""  Then
	Response.Write("{""ERR"":""PRIDERR""}")
    Response.End()
End If
If Not isNumeric(strOutWarehouseReceiptID)  Then
	Response.Write("{""ERR"":""PRIDERRNO""}")
    Response.End()
End If

If strExpressNO= ""  Then
	Response.Write("{""ERR"":""EIDERR""}")
    Response.End()
End If


Dim sql, rs

Dim Rst
Set Rst = Server.CreateObject("Adodb.RecordSet")

	strSql = "Select * From tblOutWarehouseReceipt Where OutWarehouseReceiptID = " & strOutWarehouseReceiptID 

Rst.Open strSql, adoConn, 0, 1

If Rst.Eof Then
	Response.Write("{""ERR"":""ID NOT FOUND""}")
    Response.End()
Else
	'Rst.Fields.Item("ExpressNO") = strExpressNO
	'Rst.Update

		adoConn.Execute "update tblInventory Set isValued = 0 Where ExpressNO = '" & SqlSafe(strExpressNO) & "' And OutWarehouseReceiptID = " & strOutWarehouseReceiptID 
		adoConn.Execute "Insert Into tblOrderStateLog(ExpressNO,StateDesc,CreateEmp,ModelName)Values('" & strExpressNO & "','Remove OutWarehouseReceipt:" & strOutWarehouseReceiptID & "','" & dic.CreateEmp & "','frmOrderOutWarehouse')"

	Response.Write "{""State"":""SUCCESS""}"

End If

Response.End

rs.Close
Set rs = Nothing

%>