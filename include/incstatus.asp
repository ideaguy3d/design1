
<% if session("osloginverified")<>"YES" then
	response.redirect("statuslogin.asp")
end if
%>

