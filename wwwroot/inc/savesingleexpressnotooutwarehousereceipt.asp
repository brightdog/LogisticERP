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
strOutWarehouseReceiptID = dic.OutWarehouseExpressID

'Response.write strExpressNO & "::" & strPickupReceiptID

If strOutWarehouseReceiptID= ""  Then
	Response.Write("{""ERR"":""OWRIDERR""}")
    Response.End()
End If
If Not isNumeric(strOutWarehouseReceiptID)  Then
	Response.Write("{""ERR"":""OWRIDERRNO""}")
    Response.End()
End If

If strExpressNO= ""  Then
	Response.Write("{""ERR"":""EIDERR""}")
    Response.End()
End If


Dim sql, rs, strToWarehouseID

Dim Rst
Set Rst = Server.CreateObject("Adodb.RecordSet")
'检查出库单号是否有效
strSql = "Select * From tblOutWarehouseReceipt Where OutWarehouseReceiptID = " & strOutWarehouseReceiptID 
Set Rst = adoConn.Execute(strSql)
If Rst.Eof Then
	Response.Write("{""ERR"":""OWRIDNOTEXIST""}")
	Response.End()
Else

	strToWarehouseID = Rst.Fields.Item("WarehouseID").Value

End If
Rst.Close
'检查运单号是否有效
strSql = "Select Top 1 * From tblOrder Where ExpressNO = '" & SqlSafe(strExpressNO) & "'"
'Response.Write strsql 
Rst.Open strSql, adoConn, 0, 1
If Rst.Eof Then
	Response.Write("{""ERR"":""EIDNOTEXIST""}")
	Response.End()
End If
Rst.Close


strSql = "Select Top 1 * From tblInventory Where isValued = 1 And  ExpressNO = '" & SqlSafe(strExpressNO) & "'"
strSql = strSql & " Order By CreateDT Desc"
'Response.Write strsql 
Rst.Open strSql, adoConn, 0, 1

If Not Rst.Eof Then
	If Rst.Fields.Item("InventoryState").Value <> "已出库" Then
	Dim strFromWarehouseID
	strFromWarehouseID = Rst.Fields.Item("ToWarehouseID").Value

	
	adoConn.Execute "Insert Into tblInventory(OutWarehouseReceiptID,FromWarehouseID,ToWarehouseID,ExpressNO,InventoryState,CreateEmp)Values(" & strOutWarehouseReceiptID & "," & strFromWarehouseID & "," & strToWarehouseID & ",'" & strExpressNO & "','已出库','" & dic.CreateEmp & "')"

	Response.Write "{""State"":""OK""}"
	Else
		Response.Write "{""ERR"":""InventorySTATEERR""}"

	End If
Else

	Response.Write "{""ERR"":""STATEERR""}"

End If



Rst.Close
Set Rst = Nothing
Response.End()
%>