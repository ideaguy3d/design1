
<%
	if session("registeredshopper")="NO" then
		session("user_want_multiship")=true	
		session("Clicked_Multishipto")=true
		response.redirect("login.asp")	
	end if
	
	if session("registeredshopper")="YES" then
		session("user_want_multiship")=true		
		response.redirect("basket.asp")
	end if


%>