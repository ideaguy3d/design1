<%on error resume next%>
<%

	pwd_to_use = "CALLCENTER"
	pwd_entered = trim(request("callcenterpwd"))
	
	if pwd_to_use = ucase(pwd_entered) then
		session("callcenter_logged")=true
		response.Redirect("callcenter.asp")
	else
		session("callcenter_logged")=false
		response.Redirect("call_center_login.asp")
	end if 

%>