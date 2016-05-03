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
	
Dim CurrentPage, PageNum

CurrentPage = dic.PageNum
PageSize = dic.PageSize

If CurrentPage = "" or Not IsNumeric(CurrentPage) Or CurrentPage = 0  Then
	CurrentPage = 1
End If

	Dim SqlWhere
	Dim SqlOrder
	Dim Rst
	Dim dicResult
	Set dicResult = Server.CreateObject("Scripting.Dictionary")

	SqlOrder = " Order By CreateDT Desc"

	PageSize = 20
	
	SqlWhere = MakeSqlWhere(dic.Fields,dic.Opers, dic.Values)

	SqlCount = "SELECT count(*) FROM tblTransfer Where 1=1 " & SqlWhere

	Set Rst = adoConn.execute(SqlCount)
	
	RsCount = 0
	
	If Not Rst.EOF Then
		RsCount = Rst(0)
	End If
	
	Set Rst = Nothing
	
	If RsCount Mod PageSize = 0 Then
		PageCount = Int(RsCount / PageSize)
	Else
		PageCount = Int(RsCount / PageSize) + 1
	End If
	
	If Pages <> "" Then
		If Not IsNumeric(Pages) Then Pages = 0
		Pages = Pages - 1

		If CLng(Pages) >= CLng(PageCount) Then Pages = PageCount - 1
		If Pages < 0 Then Pages = 0
	End If
	
	If Pages = "" Then Pages = 0
	
	SqlSelect = " SELECT TOP " & PageSize & " *  From tblTransfer " & " where 1=1 " & SqlWhere & " and TransferID Not In(SELECT TOP " & (CurrentPage - 1) * PageSize & " TransferID From tblTransfer where 1=1 " & SqlWhere & SqlOrder & ")"  & SqlOrder
	
	 'Response.write SqlSelect
'       Response.end
	Set Rst = adoConn.execute(SqlSelect)

	'If Not Rst.EOF Then
		Response.Write MakeJsonFromRst(Rst, PageCount, CurrentPage, RsCount, PageSize)
	'Else
	'End If
	
	Rst.Close
	Set Rst = Nothing

Response.End




%>