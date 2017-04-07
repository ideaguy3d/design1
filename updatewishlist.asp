<%on error resume next%>


<%
  if session("registeredshopper")="NO" then
  	response.redirect("login.asp")
  end if
 %>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<%

	y=REQUEST.FORM("qty").count
	
	'if (request.form("updatewishlist")="updatewishlist") then
	'	action="UPDATE"
	'else
	'	'multiple itemadd from wishlist
	'	action="ADDTOCART"
	'end if
	
	action="UPDATE"
	
	for z=y  to 1 step-1
		remove_flag = REQUEST.FORM("remove"&cstr(z)).count
		qty = REQUEST.FORM("qty")(z)
		qty = qty + 1 -1
		itemnum = cstr(REQUEST.FORM("item")(z))	
		extra=""
		if remove_flag= 1 or qty= 0 then
			call sitelink.WISHLIST("DELETE",itemnum,qty,0,"", cstr(extra),session("addressbookcustnum"))
		else
			if action="UPDATE" then
				'response.write(itemnum &"---"&qty&"<br>")
				call sitelink.WISHLIST("UPDATE",itemnum,(qty),0,"", cstr(extra),session("addressbookcustnum"))
			end if
			if action="ADDTOCART" then
				item_id=-1
				SL_WishlistCinfo=""
				SL_StAttibNonSkuprice=""
				call sitelink.ITEMADD(cstr(itemnum),cstr(qty),cstr(""),session("shopperid"),session("ordernumber"),cstr(SL_WishlistCinfo),cint(0),item_id,cstr(""),cstr(SL_StAttibNonSkuprice))
			
			end if 
		end if 
	
	next 
	
	set sitelink=nothing
	
	'response.write(action)
	
	'if action="UPDATE" then
	'	response.redirect("wishlist.asp")
	'else
	'	response.redirect("basket.asp")
	'end if
		'for z=y  to 1 step-1
		'	remove_flag = REQUEST.FORM("remove"&cstr(z)).count
		'	qty = REQUEST.FORM("qty")(z)
		'	qty = qty + 1 -1
		'	itemnum = cstr(REQUEST.FORM("item")(z))
		'	extra = ""
		'	if remove_flag= 1 or qty= 0 then
		 '		call sitelink.WISHLIST("DELETE",itemnum,qty,0,"", cstr(extra),session("addressbookcustnum"))
			'else
			'	call sitelink.WISHLIST("UPDATE",itemnum,qty,0,"", cstr(extra),session("addressbookcustnum"))
			'end if
		'next
		
		
	set sitelink=nothing
	response.redirect("wishlist.asp")
	
%>

