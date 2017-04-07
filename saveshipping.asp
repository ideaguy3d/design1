<%on error resume next%>

<%

	session("shippingmethod") = cstr(REQUEST.FORM("txtshiplist"))

	'on checkout page total is called anyway.. No neeed to call here..
	'call sitelink.TOTALORDER(session("shopperid"),session("ordernumber"),session("billtocopy"),session("shippingmethod"))
	
	'set sitelink=nothing
	response.redirect("checkout.asp")

 %>
