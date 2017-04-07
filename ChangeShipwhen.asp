<%on error resume next%>
<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 
<!--#INCLUDE FILE = "include/momapp.asp" -->

<%

item_id = request.QueryString("itemid")
item_id = item_id + 1 - 1
shipwhendate = cstr(request.QueryString("shipwhendate"))

call sitelink.WRITESHIP_WHEN_BASKET(session("ordernumber"),item_id , cstr(shipwhendate))

set sitelink=nothing


 %>