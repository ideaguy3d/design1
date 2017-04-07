
<%

If Request.cookies("checking") = "" then 
  Response.redirect("cookieerror.asp")
end if 

session("GoodforCookies") = true

if len(trim(session("destpage"))) > 0 then
	response.redirect(session("destpage"))
else
	response.redirect("default.asp")
end if 


%>