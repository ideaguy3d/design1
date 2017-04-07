
<%

if session("GoodforCookies") = false then
		Response.cookies("checking") = "test"
		Response.Redirect("Results.asp")	
end if 

  
 %>
