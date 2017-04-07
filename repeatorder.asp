
<%on error resume next%>


<%
  if session("registeredshopper")="NO" then
  	response.redirect("login.asp")
  end if
 %>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%
  ' Now copy the entire basket info from old order number to new order number
  dim ThisOrderNumber
  ThisOrderNumber = request.querystring("ordernum")  
  ThisOrderNumber = ThisOrderNumber + 1 -1

  
  MOMOrd  = Request.QueryString("ismom")
  
  if MOMOrd = 1 then
	IsMOMOrdNum = true
  else
    IsMOMOrdNum = false
  end if 
  
  'empty the cart
  CALL sitelink.EMPTYBASKET(SESSION("SHOPPERID"),SESSION("ORDERNUMBER")) 	
  
  SL_Promotional_item=0
  set ORDER_RECORD = sitelink.ORDER_RECORD(ThisOrderNumber,IsMOMOrdNum)
  	SL_Promotional_item = ORDER_RECORD.item_id  
  set ORDER_RECORD = nothing
  
    
  xmlstring = sitelink.getbasketinfo(session("shopperid"),ThisOrderNumber,"",IsMOMOrdNum,true)
  objDoc.loadxml(xmlstring)
    
  has_multishipto= false
  
  if SL_Promotional_item > 0 then
  	t_node = "//gbi[certid=0 and record !="+cstr(SL_Promotional_item)+"]"
  '	t_node= "//gbi[basketdesc!='PROMO']"
  else
  	t_node = "//gbi[certid=0]"
  end if
  
 
  
  set SL_Basket = objDoc.selectNodes(t_node)    
  for x=0 to SL_Basket.length-1 	  
	  SL_Number	= SL_Basket.item(x).selectSingleNode("number2").text
	  SL_Variant= SL_Basket.item(x).selectSingleNode("variant").text
  	  SL_BasketQty	= SL_Basket.item(x).selectSingleNode("quanto").text
	  SL_Basket_Shipto =SL_Basket.item(x).selectSingleNode("ship_to").text
	  SL_Basket_custominfo =SL_Basket.item(x).selectSingleNode("custominfo").text
	  
	  new_addrress_id =0
	  
	  if SL_Basket_Shipto > 0 and MULTI_SHIP_TO=1 then
	  	 has_multishipto=  true
	 	 sfirstname	= SL_Basket.item(x).selectSingleNode("firstname").text
		 slastname 	= SL_Basket.item(x).selectSingleNode("lastname").text
		 scompany	= SL_Basket.item(x).selectSingleNode("company").text
		 saddress1	= SL_Basket.item(x).selectSingleNode("addr").text
		 saddress2	= SL_Basket.item(x).selectSingleNode("addr2").text
		 saddress3 = SL_Basket.item(x).selectSingleNode("addr3").text
		 scity		= SL_Basket.item(x).selectSingleNode("city").text
		 sstate		= SL_Basket.item(x).selectSingleNode("state").text
		 szipcode	= SL_Basket.item(x).selectSingleNode("zipcode").text
		 scountry	= SL_Basket.item(x).selectSingleNode("country").text
		 sphone 	= SL_Basket.item(x).selectSingleNode("phone").text
		 semail		= SL_Basket.item(x).selectSingleNode("email").text		 
		 shopperid = 0
		 cintaddress_id =0
		 shopcustid =-1
		 
		 session("user_want_multiship")= true 
		 session("items_spaned") = true
		 
		 'check whether address exists
		 shopcustid= sitelink.CHECK_ADDRESSBOOK(sfirstname,slastname,scompany,saddress1,saddress2,scity,sstate,szipcode,semail,session("addressbookcustnum"))
		 				 
		 if shopcustid=-1 and (session("addressbookcustnum") > 0 or len(session("addressbookcustnum")) > 0) then
		 	 'this will make entry in address and shopper
		 	 
		 	 set SLShopper = New cShopperInfo
		                SLShopper.shopperid=session("shopperid")
		                SLShopper.bfirstname=sfirstname
		                SLShopper.blastname=slastname
		                SLShopper.bcompany=scompany
		                SLShopper.baddr1=saddress1
		                SLShopper.baddr2=saddress2
		                SLShopper.baddr3=saddress3
		                SLShopper.bcity=scity
		                SLShopper.bstate=sstate
		                SLShopper.bzipcode=szipcode
		                SLShopper.bcountry=scountry
		                SLShopper.bcounty=scounty
		                SLShopper.bphone=sphone	
		                SLShopper.bemail=semail	
		                SLShopper.bAddbookCust=session("addressbookcustnum")		                
		                new_addrress_id=sitelink.UpdateAddressBook(SLShopper)
		                set SLShopper = nothing

			 'new_addrress_id= sitelink.UpdateAddressBook(sfirstname,slastname,scompany,saddress1,saddress2,scity,sstate,szipcode,sphone,semail,shopperid,scountry,cintaddress_id,session("ordernumber"),saddress3,session("addressbookcustnum"))
			 'new_addrress_id = session("lastshopper")	
		 else
		 	 new_addrress_id =shopcustid
		 end if
	 end if
	  CALL sitelink.ITEMADD(cstr(SL_Number),cstr(SL_BasketQty),cstr(SL_Variant),session("shopperid"),session("ordernumber"),cstr(SL_Basket_custominfo),cint(new_addrress_id),cint(0))
  
  next
  
  
  if has_multishipto=false then
  	set shopperRecord = sitelink.SHIPTONAME(ThisOrderNumber,IsMOMOrdNum)
	'populate ship to
	session("billtocopy")=0
	session("sfirstname")=cstr(trim(shopperRecord.firstname))
	session("slastname")=cstr(trim(shopperRecord.lastname))
	session("scompany")=cstr(trim(shopperRecord.company))
	session("saddress1")=cstr(trim(shopperRecord.addr))
	session("saddress2")=cstr(trim(shopperRecord.addr2))
	session("saddress3")=cstr(trim(shopperRecord.addr3))
	session("scity")=cstr(trim(shopperRecord.city))
	session("sstate")=UCASE(cstr(trim(shopperRecord.state)))
	session("szipcode")=cstr(trim(shopperRecord.zipcode))
	session("scountry")=cstr(shopperRecord.country)
	session("scounty")=cstr(shopperRecord.county)
	session("sphone")=cstr(trim(shopperRecord.phone))
	session("semail")=cstr(trim(shopperRecord.email))	
	set shopperRecord = nothing
  end if
  
  
  set SL_Basket = nothing
  Set objDoc = nothing
  Set sitelink = nothing
 
 response.redirect("basket.asp")


%>