<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<%

rowcount = REQUEST.FORM("rowcount")

if rowcount = 0 then
	set sitelink=nothing 
  Response.Redirect("basket.asp")
end if 

'Response.Write(rowcount)


for z=1 to rowcount

stockitem = REQUEST.FORM("stockitem")(z) 
qty = REQUEST.FORM("txtquanto")(z) 
variantstring = "txtvariant" & z
TXTVARIANT = cstr(REQUEST.FORM(variantstring))

custominfofield = REQUEST.FORM("txtcustominfo")(z) 

qty = fix_xss_Chars(qty)

'Response.Write(stockitem&"---"&qty&"--"&TXTVARIANT&"<br>")

	'if MULTI_SHIP_TO=1 and session("user_want_multiship")= true then
	'	qty_to_add = cstr(qty)
		'response.write("hhh")
	'	for y = 1 to cint(qty_to_add)
	'		CALL sitelink.ITEMADD(cstr(stockitem),cstr("1"),cstr(TXTVARIANT),session("shopperid"),session("ordernumber"),cstr(REQUEST.FORM("TXTcustominfo")),cint(0),cint(0))
	'	next
	'else
	'	CALL sitelink.ITEMADD(cstr(stockitem),cstr(qty),cstr(TXTVARIANT),session("shopperid"),session("ordernumber"),cstr(REQUEST.FORM("TXTcustominfo")),cint(0),cint(0))
	'end if 
		
'CALL sitelink.ITEMADD(cstr(stockitem),CSTR(qty),cstr(TXTVARIANT),session("shopperid"),session("ordernumber"),cstr(""),cint(0),cint(0))
CALL sitelink.ITEMADD(cstr(stockitem),cstr(qty),cstr(TXTVARIANT),session("shopperid"),session("ordernumber"),cstr(custominfofield),cint(0),cint(0),cstr(""),cstr("0"))	

next

set sitelink=nothing 
Response.Redirect("basket.asp")
%>