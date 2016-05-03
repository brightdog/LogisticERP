<%
For Each x In Request.Form
	Response.Write "name:" & x & "," & "value:" & Request.form(x) & "<br />"
Next
%>