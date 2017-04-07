<%on error resume next%>

<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 
<%
if session("SL_BasketCount")= 0 then
	response.redirect("basket.asp")
end if
	
%>
<!--#INCLUDE FILE = "include/momapp.asp" -->


<%

	'orderconfirm = sitelink.ORDER_CONFIRMED(session("ordernumber"))
	'if orderconfirm = true then
	 ' 	set sitelink=nothing	
	 ' 	set objDoc=nothing  	
	'	response.redirect("receipt.asp")
	'end if

	
	y=REQUEST.FORM("qty").count
	
	ship_via_count = REQUEST.FORM("Ship_via").count
	for z=ship_via_count  to 1 step -1
	  itemid  = REQUEST.FORM("item_id")(z)
	  item_id = itemid + 1 - 1
	  Shipping_Method  =""	  
	  'get ship when dates.
	   inputstring = "shipwhen"&z
	   shipwhendate = REQUEST.FORM(inputstring)

	   shipwhendate=fix_xss_Chars(shipwhendate)

	    if ( shipwhendate= "") or isdate(shipwhendate)=true then	   
	     ' write ship via / ship when dates to basket
	      call sitelink.WRITESHIP_WHEN_BASKET(session("ordernumber"),item_id , cstr(shipwhendate))
	   end if 
	next
	'end if
		
	
	'now update/delete line item
	basket_changed=false	
	for z=y  to 1 step-1
		previousqty = REQUEST.FORM("prevqty")(z)
		quantity = REQUEST.FORM("qty")(z)
		quantity=fix_xss_Chars(quantity)
		
		if isnumeric(quantity)=false then
			quantity=previousqty
		end if
		if isnumeric(quantity)=true then
			if quantity < 0  then 
				quantity = 0		
			end if
		end if 
		SL_itemid = REQUEST.FORM("item_id")(z)
		SL_itemid= SL_itemid + 1 -1 
		
		if REQUEST.FORM("remove"&cstr(z)).count = 1 then
			removeitem = true
		else
			removeitem = false
		end if
		
		if ALLOW_POINTS=1 then
			if REQUEST.FORM("Points"&cstr(z)).count = 1 then
				'response.write(REQUEST.FORM("Points"&cstr(z)).count)
				'ItemPointvalue = REQUEST.FORM("ItemPointvalue")(z)
				'response.write(SL_itemid)
				'response.write("----")
				'response.write(ItemPointvalue)					
				'response.write("<br>")
				call sitelink.APPLY_POINTS_ITEMID(SL_itemid)
				'session("pointavail")= session("pointavail") - ItemPointvalue
				basket_changed=true
			end if
		end if
		
		if quantity=0 or removeitem=true then
			call sitelink.ITEMDELETE_ITEMID(SL_itemid,session("ordernumber"))
			basket_changed=true
		else
		 	if quantity <> previousqty then
				'update the qty				
		 	    call sitelink.MODIFYBASKET(SL_itemid,quantity,session("ordernumber"))				
				basket_changed=true
			end if
		end if 
		
	next

	 
	 if REQUEST.FORM("checkout") ="checkout" then
	 	if basket_changed=true then
		 	call SITELINK.quantitypricing(session("shopperid"),session("ordernumber"))
			session("SL_BasketCount") = sitelink.GetData_Basket(cstr("BASKETCOUNT"),session("ordernumber"),session("SL_VatInc"))
			session("SL_BasketSubTotal") = sitelink.GetData_Basket(cstr("BASKETSUBTOTAL"),session("ordernumber"),session("SL_VatInc"))	

		end if
		set sitelink=nothing
	 	securepage = secureurl + "custinfo.asp"
	 	response.redirect(securepage)
	 else	 
	 	set sitelink=nothing
	 	response.redirect("basket.asp")
	 end if 
	 

%>


