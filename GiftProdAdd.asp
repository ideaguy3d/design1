<%on error resume next%>
<!--#INCLUDE FILE = "include/momapp.asp" -->

<%

  ' first check whether the stock number entered is valid before adding to the basket
      
   allitems_valid =true
   hassizecolor =false   
   num_of_quick_items = REQUEST.FORM("num_of_quick_items")

   
   for z=1 to num_of_quick_items
	   dim vStockNumber 
	   'dim vQuantity
	   inputstring = "Stock"&z
	    quantitystring = "txtquantm"&z
		vStockNumber = trim(REQUEST.FORM(inputstring))
		vQuantity = REQUEST.FORM(quantitystring)
		'vQuantity    = REQUEST.FORM("TXTQUANTM")(z)
		session("Stock"&z) = vStockNumber
		session("txtquantm"&z) = vQuantity
		session("invaliditem"&z) = false
		

		session("Stock"&z)=fix_xss_Chars(vStockNumber)
		session("txtquantm"&z)=fix_xss_Chars(vQuantity)		

		
   next
   
   'now check whether any of the items entered are size/color.
   'if no then add to cart and launch basket page
   'if no then dispplay a page, just like dancing deer..
   
   for z=1 to num_of_quick_items
    if len(trim(session("Stock"&z))) > 0 then
		returnval = sitelink.validstocknumber2(cstr(session("Stock"&z)))
		'returnval = -1 not found
		'returnval = 0 found .. no size./color
		'returnval = 1 found and has size/color		
		if returnval = -1 then  'found atleast one that is invalid
			allitems_valid= false  
			session("invaliditem"&z) = true
		end if
		
		if returnval = 1 then 'found atleast one which is size/color
			hassizecolor = true
		end if	
	end if
	next
   
   if allitems_valid =false then    
    session("errmsg") ="All or some stock items are invalid.. <br><br>"
		for z=1 to num_of_quick_items
			if session("invaliditem"&z) = true then
				session("errmsg") = session("errmsg") + session("Stock"&z) &"---->Invalid Stock Item<br>"
			end if		
		next 
		
	session("errmsg") = session("errmsg") + "<br><br>Please try again..<br><br><a  class=""allpage"" href=""quick_order.asp"">Click here</a>&nbsp; to go to quick order page"
	set sitelink=nothing 
    Response.Redirect("quick_ordererror.asp")
  end if 
   
   if hassizecolor = true then
	   	set sitelink=nothing 
		session("num_of_quick_items") = num_of_quick_items
		Response.Redirect("display_results.asp")   
   end if 
   
   
   'all items are not size/color..
if len(trim(session("giftregcustnum"))) > 0  then
   for z=1 to num_of_quick_items
		if len(trim(session("Stock"&z))) > 0 then			
			
			'CALL sitelink.ITEMADD(cstr(session("Stock"&z)),CSTR(session("txtquantm"&z)),cstr(""),session("shopperid"),session("ordernumber"),cstr(""),cint(0),cint(0))
					 
						'now add item to gift registry
						aa = sitelink.ADDTOGIFTLIST(cstr(session("Stock"&z)),"",cint(session("txtquantm"&z)))
						'set sitelink=nothing 
						'response.redirect("giftlist.asp")
								 
		end if 
	next		
else
						set sitelink=nothing 
						gotopage = "giftreglogin.asp" 
						session("item_to_add") = cstr(current_stocknumber)
						response.redirect(gotopage)		 
 end if 

		
	
	set sitelink=nothing 
	Response.Redirect("GIFTLIST.asp")
	
	

%>