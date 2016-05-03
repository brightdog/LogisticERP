<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc-detransform.asp"-->
<%
	Response.Charset="utf-8"
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "cache-ctrol","no-cache"
	Session.CodePage=65001


	'set time session and disable caches
	session.lcid = 2057
	response.buffer = true
	response.expires = 60
	response.expiresabsolute = now() - 1
	response.addheader "pragma","no-cache"
	response.addheader "cache-control","private"
	response.cachecontrol = "no-cache"
	
%>

<%
'dimension variables
dim adoConn,iserver,loginuid,loginpwd,database

iserver=")/.(.+(+(&3211+"
loginuid="k^"
loginpwd="Ka]a00,2"
database="Dl^cqqa`<LN\<?"

set adoConn = Server.CreateObject("ADODB.Connection")

adoConn.ConnectionString = "provider=sqloledb; data source=" & DeTransform(iserver) & "; User ID=" & DeTransform(loginuid) & "; pwd=" & DeTransform(loginpwd) & "; Initial Catalog=" & DeTransform(DATABASE) & ";"
'response.write adoConn.ConnectionString
'Response.End()
adoConn.open



Function SqlSafe(strSql)

	Dim sql_injdata
	SQL_injdata = "'--|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare"
	SQL_inj = Split(SQL_Injdata,"|")
	If strSql <> "" Then
	
		For SQL_Data=0 To Ubound(SQL_inj)
			strSql = Replace(strSql, Sql_inj(SQL_Data), "")
		Next
	End If 
	
	SqlSafe = strSql
End Function

%>