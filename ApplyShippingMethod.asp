<%on error resume next%>

<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 
<%

	session("shippingmethod") = cstr(REQUEST.form("txtshiplist"))

	
 %>
