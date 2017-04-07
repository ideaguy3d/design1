<%on error resume next%>

<%
	item_id = REQUEST.FORM("item_id")
	item_id = item_id + 1 -1 

	quanto = REQUEST.FORM("quanto")
	quanto = quanto + 1 - 1
	
	if isnumeric(item_id)=false or isnumeric(quanto)=false then
			Response.Status="301 Moved Permanantly"
			Response.AddHeader "Location", insecureurl&"default.asp"
			Response.End
	end if
	

%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<%


'basketrecord = REQUEST.FORM("basketrecord")
'basketrecord = basketrecord + 1 - 1


call sitelink.ITEMMODIFY(cstr(REQUEST.FORM("item")), cstr(quanto), cstr(item_id), cstr(REQUEST.FORM("variant")), session("shopperid"),cstr(REQUEST.FORM("txtcustominfo")), session("ordernumber"))

set sitelink= nothing

Response.Redirect("basket.asp")


%>