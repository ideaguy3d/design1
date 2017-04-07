<%on error resume next%>
<!--#INCLUDE FILE = "include/momappNoClibrary.asp" -->
<% 
	session("viewfullsite") = 1 
	set sitelink = nothing
    response.Redirect(insecureurl)
    
%>
