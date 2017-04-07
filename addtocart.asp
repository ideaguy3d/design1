<%on error resume next%>
<%
  if session("registeredshopper")<>"YES" then
  	session("destpage")=""
	destpage = secureurl&"login.asp"
  	response.redirect(destpage)
  end if 
%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%

	'just add  item number and qty and item_id
	item_number = request.querystring("item")
	'item_id = request.querystring("item_id")
	'item_id = item_id + 1 -1 
	
	qty_to_add = request.querystring("qty")
	
	if len(trim(qty_to_add)) = 0 then
	    qty_to_add = 1
	end if
	
	if isnumeric(qty_to_add) = false then
	    qty_to_add = 1
	end if  
	
	'qty_to_add = 1
	item_id=-1
	
	
	'response.write(item_id)
	'Response.Write("item Num" &item_number)
	'Response.Write("<br>")
	'Response.Write("Shopper id "& session("shopperid"))
	'Response.Write("<br>")
	'Response.Write("Order Num "& session("ordernumber"))
	'Response.Write("<br>")
	'response.write("Item id "& item_id)
	
	if isnumeric(item_id)=true then
	    has_attribs = false
	    SET LOSTOCK =  sitelink.setupproduct(item_number,cstr(""),0,session("ordernumber"),cstr("")) 
	        has_attribs=LOSTOCK.attribs
	    SET LOSTOCK = nothing
	    
	    variant_part =""
	    if has_attribs=true then
			'item_number= left(item_number,10)
	        variant_part = mid(item_number,11,20)
	    end if 
		
		SL_StAttibNonSkuprice=0
		SL_WishlistCinfo=""
		
		'get custominfo and addtional non sku prices
		xmlstring =sitelink.WISHLIST("LIST","0",0,0,"","",session("addressbookcustnum"))
		objDoc.loadxml(xmlstring)
		tnode = "//wls[item='"+cstr(item_number)+"']"
		set SL_Wishlist = objDoc.selectNodes(tnode)
		
		'response.write(SL_Wishlist.length)
		if SL_Wishlist.length > 0 then
			SL_StAttibNonSkuprice  = SL_Wishlist.item(0).selectSingleNode("nskuprice").text
			SL_WishlistCinfo  = SL_Wishlist.item(0).selectSingleNode("custominfo").text
		end if
		
		set SL_Wishlist = nothing

		'response.write(SL_StAttibNonSkuprice)	
	    
		call sitelink.ITEMADD(cstr(item_number),cstr(qty_to_add),cstr(variant_part),session("shopperid"),session("ordernumber"),cstr(SL_WishlistCinfo),cint(0),item_id,cstr(""),cstr(SL_StAttibNonSkuprice))
	end if

  
  set sitelink=nothing
  Set ObjDoc=nothing
response.redirect("basket.asp")


%>