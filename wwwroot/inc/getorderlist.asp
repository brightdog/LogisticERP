<!--#include file="inc-connection.asp"-->
<!--#include file="common.asp"-->
<%


Dim CurrentPage
CurrentPage = Request("p")


If CurrentPage = "" or Not IsNumeric(CurrentPage) Then
	CurrentPage = 1
End If

	Dim SqlWhere
	Dim SqlOrder
	Dim Rst
	Dim dicResult
	Set dicResult = Server.CreateObject("Scripting.Dictionary")

	SqlOrder = ""'" Order By CreateDT Desc"

	PageSize = 20
	
	

	SqlCount = "SELECT count(*) FROM tblOrder" & SqlWhere

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
	
	SqlSelect = " SELECT TOP " & PageSize & " OrderID  From tblOrder " & " where 1=1 " & SqlWhere & " and OrderID Not In(SELECT TOP " & CurrentPage * PageSize & " OrderID From tblOrder where 1=1 " & SqlWhere & SqlOrder & ")"  & SqlOrder
	
'	 Response.write SqlSelect
'       Response.end
	Set Rst = adoConn.execute(SqlSelect)

	'If Not Rst.EOF Then
		Response.Write MakeJsonFromRst(Rst, PageCount, CurrentPage, RsCount, PageSize)
	'Else
	'End If
	
	Rst.Close
	Set Rst = Nothing

Response.End

rs.Close
Set rs = Nothing


%>