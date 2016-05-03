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
	Dim Rst
	Set Rst = Server.CreateObject("Adodb.RecordSet")
	If dicFieldValue.Item("OrderID") <> "" Then
		strSql = "Select * from tblOrder Where OrderID = " & dicFieldValue.Item("OrderID")
	Else
		strSql = "Select * from tblOrder Where 1 = 2"
	End If

	Rst.Open strSql, adoConn, 1, 3

	If Rst.Eof Then
		Rst.AddNew

		Dim v
		For Each v In dicFieldValue.Keys

			If CStr(v) <> "OrderID" And dicFieldValue.Item(v) <> "" Then
				'Response.write CStr(v) & "+" & dicFieldValue.Item(v) & vbcrlf
				Rst.Fields.Item(CStr(v)).Value = dicFieldValue.Item(v)

			End If

		Next
	
		Rst.Update
		
		Rst.Close

		Rst.Open "Select Top 1 * From tblOrder Order By OrderID Desc", adoConn, 0, 1


	Else

		For Each v In dicFieldValue.Keys

			If CStr(v) <> "OrderID" And dicFieldValue.Item(v) <> "" Then
				'Response.write CStr(v) & "+" & dicFieldValue.Item(v) & vbcrlf
				Rst.Fields.Item(CStr(v)).Value = dicFieldValue.Item(v)

			End If

		Next
	
		Rst.Update
		
		Rst.Close

		Rst.Open "Select Top 1 * From vwOrder_Detail Where OrderID = " & dicFieldValue.Item("OrderID"), adoConn, 0, 1


	End If

	Response.Write MakeJsonFromRst(Rst, 0, 0, 0, 0)

	Rst.Close
	Set Rst = Nothing

Response.End

rs.Close
Set rs = Nothing


%>