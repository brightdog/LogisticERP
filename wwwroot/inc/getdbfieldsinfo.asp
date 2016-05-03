<!--#include file="inc-connection.asp"-->
<!--#include file="common.asp"-->
<!--#include file="json.inc"-->
<!--#include file="clsstringbuilder.inc"-->
<%



	Dim Rst
	Dim strSql
	Dim dicResult
	Set dicResult = Server.CreateObject("Scripting.Dictionary")




	strSql= "select Name from sysobjects where xtype='u' and status>=0"
	
	Set Rst = adoConn.Execute(strSql)

	If Not Rst.EOF Then
		
		Do While Not Rst.EOF
		

			Dim tmpRs
			
			strSql = "select column_name AS ColName ,data_type AS ColType, "
			strSql = strSql & "CASE WHEN CHARACTER_MAXIMUM_LENGTH IS NULL THEN 0 ELSE CHARACTER_MAXIMUM_LENGTH END as MaxLen "
			strSql = strSql & "from information_schema.columns "
			strSql = strSql & "where table_name ='" & Rst.Fields.Item(0).Value & "'"

			Set tmpRs = adoConn.Execute(strSql)
			
			If Not tmpRs.EOF Then
				
				dicResult.Add Rst.Fields.Item(0).Value, MakeJsonFromRst(tmpRs, 0, 0, 0, 0)

			Else
			
			End If

			

			Rst.MoveNext

		Loop
		

		Dim v

		For Each v In dicResult.Keys

		
			strResult = strResult & """" & CStr(v) & """:" & dicResult.Item(v) & ","


		Next
		
		Response.Write "{" & Left(strResult, Len(strResult) - 1)  & "}"

	Else
	End If
	
	Rst.Close
	Set Rst = Nothing

Response.End




%>