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
Dim strPickupReceiptID
Dim strWarehouseID
strExpressNO = dic.ExpressNO
strPickupReceiptID = dic.PickupReceiptNO

strWarehouseID = GetWareHouseIDfromName(dic.WareHouseName)
'Response.write strExpressNO & "::" & strPickupReceiptID
'response.write strWarehouseID & "***"
If strPickupReceiptID = ""  Then
	Response.Write("{""ERR"":""PRIDERR""}")
    Response.End()
End If
If Not isNumeric(strPickupReceiptID)  Then
	Response.Write("{""ERR"":""PRIDERRNO""}")
    Response.End()
End If

If strExpressNO = ""  Then
	Response.Write("{""ERR"":""EIDERR""}")
    Response.End()
End If

If strWarehouseID < 0  Then
	Response.Write("{""ERR"":""WHNAMENOTEXIST""}")
    Response.End()
End If

Dim sql, rs

Dim Rst
Set Rst = Server.CreateObject("Adodb.RecordSet")

strSql = "Select * From tblOrder Where ExpressNO = '" & SqlSafe(strExpressNO) & "'"
Rst.Open strSql, adoConn, 0, 1
If Not Rst.Eof Then

	Response.Write "{""ERR"":""EIDDUPLICATE""}"
	Response.End()
End If
Rst.Close

	strSql = "Select * From tblOrder Where PickupReceiptID = " & strPickupReceiptID & " And ExpressNO = '" & SqlSafe(strExpressNO) & "'"
'Response.write strsql 
Rst.Open strSql, adoConn, 1, 3

If Rst.Eof Then
	Rst.AddNew
	Rst.Fields.Item("ExpressNO") = strExpressNO
	Rst.Fields.Item("PickupReceiptID") = strPickupReceiptID
	Rst.Fields.Item("CreateEmp") = dic.CreateEmp

	Rst.Update
	Rst.Close


	Rst.Open "Select * From tblInventory Where 1=2", adoConn, 1,3

	Rst.AddNew
	Rst.Fields.Item("ToWarehouseID").Value = strWarehouseID
	Rst.Fields.Item("ExpressNO").Value = strExpressNO
	Rst.Fields.Item("InventoryState").Value = "已收件"
	Rst.Fields.Item("CreateEmp").Value = dic.CreateEmp
	Rst.Update
	Rst.Close

	adoConn.Execute "Update PickupReceipt Set PickupState = '已取件' Where PickupReceiptID = " & strPickupReceiptID
	
	

	adoConn.Execute "Insert Into tblOrderStateLog(ExpressNO,StateDesc,CreateEmp,ModelName)Values('" & strExpressNO & "','NEW to PickupReceipt:" & strPickupReceiptID & "','" & dic.CreateEmp & "','frmOrderInWarehouse')"

	Response.Write "{""State"":""NEW""}"
Else
	'Rst.Fields.Item("ExpressNO") = strExpressNO
	'Rst.Update
		adoConn.Execute "Insert Into tblOrderStateLog(ExpressNO,StateDesc,CreateEmp,ModelName)Values('" & strExpressNO & "','UPDATE to PickupReceipt:" & strPickupReceiptID & "','" & dic.CreateEmp & "','frmOrderInWarehouse')"

	Response.Write "{""State"":""UPDATE""}"

End If

Response.End

rs.Close
Set rs = Nothing

%>