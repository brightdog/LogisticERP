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
	
Dim PageSize, TableName, strField, Oper, Value

PageSize = dic.PageSize
TableName = dic.TableName
strField = dic.strField
Oper = dic.Oper
Value = dic.Value


	Dim SqlWhere
	Dim SqlOrder
	Dim Rst
	Dim dicResult
	Set dicResult = Server.CreateObject("Scripting.Dictionary")

	SqlOrder = " Order By " & strField & " Asc"

	PageSize = 20
	
	SqlWhere = " And " & strField & " " & Oper & " '" & Value & "' "

	
	
	SqlSelect = " SELECT TOP " & PageSize & " " & strField & " From " & TableName & " where 1=1 " & SqlWhere  & SqlOrder
	
'		Response.write SqlSelect
'		Response.end
	Set Rst = adoConn.execute(SqlSelect)

	'If Not Rst.EOF Then
		Response.Write MakeJsonFromRst(Rst, 0, 0, 0, 0)
	'Else
	'End If
	
	Rst.Close
	Set Rst = Nothing

Response.End




%>