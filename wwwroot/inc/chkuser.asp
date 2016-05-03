
<%
Function ChkUser()

	If Request.Cookies("EmpID") & "" <> "" And Request.Cookies("EngName") & "" <> "" Then
		
		Session("EmpID")  = Request.Cookies("EmpID")
		Session("EngName") = Request.Cookies("EngName")
		Session("ChsName") = Request.Cookies("ChsName") & ""
		
		ChkUser = True
	
	Else
		
		Session("EmpID") = ""
		Session("EngName") = ""
		Session("ChsName") = ""
		

		Response.Cookies("EmpID") = ""
		Response.Cookies("EngName") = ""
		Response.Cookies("ChsName") = ""
		
		ChkUser = False	
	
	End If

End Function

%>