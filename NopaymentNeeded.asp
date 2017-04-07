<%on error resume next%>
<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 


<%
    if request.Form("SITELINK")="SITELINK" then
        response.Redirect("processorder.asp")
    else
        response.Redirect("checkout.asp")
    end if 

 %>