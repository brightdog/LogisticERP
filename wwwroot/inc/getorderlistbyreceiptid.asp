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
	

PageSize = dic.PageSize

If PageSize <= 0 Then
	PageSize = 20'防止忘记传这个参数的时候，也确保可以至少有数据输出
End if

	Dim SqlWhere
	Dim SqlOrder
	Dim Rst
	Dim dicResult
	Set dicResult = Server.CreateObject("Scripting.Dictionary")

	SqlOrder = " Order By CreateDT Desc"

	SqlWhere = MakeSqlWhere(dic.Fields,dic.Opers, dic.Values)
	
	SqlSelect = " SELECT TOP " & PageSize & " *  From tblOrder where 1 = 1 " & SqlWhere  & SqlOrder
	
'	Response.write SqlSelect
'       Response.end
	Set Rst = adoConn.execute(SqlSelect)

	'If Not Rst.EOF Then
		Response.Write MakeJsonFromRst(Rst, 0, 0, 0, 0)
	'Else
	'End If
	
	Rst.Close
	Set Rst = Nothing

Response.End




%>