
<%

	Session("EmpID") = ""
	Session("EngName") = ""
	Session("ChsName") = ""

	Response.Cookies("EmpID") = ""
	Response.Cookies("EngName") = ""
	Response.Cookies("ChsName") = ""
	
	Response.Write("OK")

%>