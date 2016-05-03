<!--#include file="inc-connection.asp"-->
<!--#include file="common.asp"-->
<!--#include file="json.inc"-->
<!--#include file="clsstringbuilder.inc"-->
<%


Dim strData
strData = Request("data")
'Response.write strdata
'Response.write "<br>===<br>"

Dim dic
Set dic = toObject(strData)
'Response.write typename(dic) & "**"
'Response.write CStr(dic.fields)
'Response.write "<br>===<br>"
'Response.write dic.values
'Response.write "<br>===<br>"
'Response.write typename(dic.fields)
'Response.write "###"

'Response.end

If strData = ""  Then
	Response.Write "{""ERR"":""DATA is NULL""}"
	Response.End
End If
	



	Dim dicFieldValue
	Set dicFieldValue = Server.CreateObject("Scripting.Dictionary")

	Dim arrField, arrValue


	arrField = Split(CStr(dic.fields), ",")
	arrValue = Split(CStr(dic.values), ",")
	
	Dim i
	If UBound(arrField) <> UBound(arrValue) Then
	
		Response.Write "{""ERR"":""Data ERR""}"
		Response.End
	End If
	For i = 0 To UBound(arrField)

		dicFieldValue.Add arrField(i), arrValue(i)

	Next

	Dim strSql
	Dim SqlWhere
	Dim SqlOrder
	Dim Rst
	Dim dicResult
	Set dicResult = Server.CreateObject("Scripting.Dictionary")

	strSql = "SELECT * FROM vwPickupReceipt_List Where PickupReceiptID =" & dicFieldValue.Item("ID")
	'Response.write strSql & vbcrlf
	Set Rst = adoConn.execute(strSql)
	
	'If Not Rst.EOF Then
		Response.Write MakeJsonFromRst(Rst, 0, 0, 0, 0)
	'Else
	'End If
	
	Rst.Close
	Set Rst = Nothing

Response.End




%>