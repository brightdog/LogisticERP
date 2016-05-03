<!--#include file="inc-connection.asp"-->
<!--#include file="common.asp"-->
<!--#include file="json.inc"-->
<!--#include file="clsstringbuilder.inc"-->
<%


 
Dim json
json = "{'field':['uid','username','email'], 'value':['1','abc','123@163.com']}"
Set json = toObject(json)

	Response.write json

Response.end
Response.Write json.uid & "<br/>"
Response.Write json.username & "<br/>"
Response.Write json.email & "<br/>"
 
Set json = Nothing
%>
