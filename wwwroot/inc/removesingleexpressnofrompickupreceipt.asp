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
strExpressNO = dic.ExpressNO
strPickupReceiptID = dic.PickupReceiptNO

'Response.write strExpressNO & "::" & strPickupReceiptID

If strPickupReceiptID= ""  Then
	Response.Write("{""ERR"":""PRIDERR""}")
    Response.End()
End If
If Not isNumeric(strPickupReceiptID)  Then
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

	strSql = "Select * From tblOrder Where PickupReceiptID = " & strPickupReceiptID & " And ExpressNO = '" & SqlSafe(strExpressNO) & "'"

Rst.Open strSql, adoConn, 0, 1

If Rst.Eof Then
	Response.Write("{""ERR"":""ID NOT FOUND""}")
    Response.End()
Else
	'Rst.Fields.Item("ExpressNO") = strExpressNO
	'Rst.Update
		adoConn.Execute "Update tblOrder Set PickupReceiptID = 0 Where ExpressNO = '" & SqlSafe(strExpressNO) & "'"
		adoConn.Execute "Insert Into tblOrderStateLog(ExpressNO,StateDesc,CreateEmp,ModelName)Values('" & strExpressNO & "','Remove PickupReceipt:" & strPickupReceiptID & "','" & dic.CreateEmp & "','frmOrderInWarehouse')"

	Response.Write "{""State"":""SUCCESS""}"

End If

Response.End

rs.Close
Set rs = Nothing

%>