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
'Response.write typename(dic) & "**"
'Response.write CStr(dic.fields)
'Response.write "<br>===<br>"
'Response.write dic.values
'Response.write "<br>===<br>"
'Response.write typename(dic.fields)
'Response.write "###"

'Response.end


Dim strSql
Dim Rst
Set Rst = Server.CreateObject("Adodb.RecordSet")

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

		dicFieldValue.Add MappingField(arrField(i)), arrValue(i)
	'Response.write MappingField(arrField(i)) & "," &  arrValue(i) & vbcrlf
Next
'Response.end
If dic.Type = "OutWarehouseReceipt_Detail" Then


'	Dim arrOrderList
'	arrOrderList = Split(dicFieldValue.Item("SelectedOrders"), "|")


	If dicFieldValue.Item("OutWarehouseReceiptID") <> "" Then
		strSql = "Select * from tblOutWarehouseReceipt Where OutWarehouseReceiptID = " & dicFieldValue.Item("OutWarehouseReceiptID")
	Else
		strSql = "Select * from tblOutWarehouseReceipt Where 1 = 2"
	End If
'Response.write strsql
	Rst.Open strSql, adoConn, 1, 3

	If Rst.Eof Then
		Rst.AddNew
	End If

	Dim v

	For Each v In dicFieldValue.Keys

		If CStr(v) <> "OutWarehouseReceiptID" Then
			'对于当前取简单所包含的子订单，需要存到另外一个明细表中。
			'Response.write CStr(v) & "+" & dicFieldValue.Item(v) & vbcrlf
			Rst.Fields.Item(CStr(v)).Value = dicFieldValue.Item(v)

		End If

	Next

	Rst.Update
	
	Rst.Close

	
	'Rst.Open "Select * From vwPickupReceipt_Detail Where PickupReceiptID = " & dicFieldValue.Item("PickupReceiptID")

	'Response.Write MakeJsonFromRst(Rst, 0, 0, 0, 0)
	Response.Write "{""Save"":""OK""}"
	



ElseIf dic.Type = "LoadOutWarehouseReceipt_Detail" Then

	strSql = "Select * from tblOutWarehouseReceipt Where OutWarehouseReceiptID = " & dicFieldValue.Item("OutWarehouseReceiptID")

	Rst.Open strSql, adoConn, 0, 1
	Response.Write MakeJsonFromRst(Rst, 0, 0, 0, 0)

End If
Response.End


Set rs = Nothing


'Function UpdatePickupReceipt_Detail(PickupReceiptID, arrOrderList)
	'Dim Rst
	'Set Rst = Server.CreateObject("Adodb.RecordSet")

	'If UBound(arrOrderList) > -1 Then
		
''		Dim i
''		For i = 0 To UBound(arrOrderList)
''			If arrOrderList(i) <> "" Then
''
''
''				Rst.Open "Select * From PickupReceipt_Detail Where PickupReceiptID = " & PickupReceiptID & " And OrderID = " & arrOrderList(i), adoConn, 1, 3
''			
''				If Rst.Eof Then
''					
''					Rst.AddNew
''
''				End If
''				'Response.write PickupReceiptID & "::" & arrOrderList(i) & vbcrlf
''				Rst.Fields.Item("PickupReceiptID") = PickupReceiptID
''				Rst.Fields.Item("OrderID").Value = arrOrderList(i)
''				Rst.Fields.Item("OrderID").Value = arrOrderList(i)
''				Rst.Update
''				Rst.Close
''			Else
''				'万一列表为空该怎么办呢？
''			End If
''
''		Next
	'End If

'End Function

Function MappingField(strFieldName)

	Select Case strFieldName

		Case "TransferList"

			MappingField = "TransferID"

		Case "WarehouseList"

			MappingField = "WarehouseID"

		Case Else

			MappingField = strFieldName

	End Select


End Function

%>