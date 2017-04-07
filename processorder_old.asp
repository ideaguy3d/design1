
<%on error resume next%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 

<%

    'paypal will send this.. It is the SL order number that was passed from SL to PayPal.
	invoice_pp = cstr(REQUEST.FORM("invoice"))
	
	if len(trim(invoice_pp)) > 0 then
	    session("ordernumber") = cstr(invoice_pp)
	    session("SL_BasketCount") = sitelink.GetData_Basket(cstr("BASKETCOUNT"),session("ordernumber"),session("SL_VatInc"))
	    
	     'get billto
	        billtocust = sitelink.GetBilltoCust(session("ordernumber"))
	        set shopperRecord= sitelink.SHOPPERINFO(cstr(billtocust),false)
	        session("shopperid") = shopperRecord.custnum
	        session("firstname")= trim(shopperRecord.FirstName)
	        session("lastname")= trim(shopperRecord.Lastname)
	        session("address1")= trim(shopperRecord.addr)
	        session("address2")= trim(shopperRecord.addr2)
	        session("city")= trim(shopperRecord.city)
	        session("state")= trim(shopperRecord.state)
	        session("zipcode")= trim(shopperRecord.zipcode)
	        session("email")= trim(shopperRecord.email)
	        session("country")= trim(shopperRecord.country)
	        set shopperRecord= nothing
	        first_name = session("firstname")
	        last_name=  session("lastname")
	        address1 =session("address1")
	        address2 =session("address2")
	        city = session("city")
	        state = session("state")
	        zip= session("zipcode")
			session("paymentmethod")="PP"
	  
	end if 
	
    if session("SL_BasketCount")= 0 then
		set sitelink=nothing 
		set objDoc=nothing
		response.redirect("basket.asp")
	end if

    if len(trim(session("firstname")))= 0 or len(trim(session("lastname")))= 0  then
		set sitelink=nothing 
		set objDoc=nothing
		response.redirect("custinfo.asp")
	end if


	
	'orderconfirm = sitelink.ORDER_CONFIRMED(session("ordernumber"))
	'if orderconfirm = true then
	 ' 	set sitelink=nothing	
	  '	set objDoc=nothing  	
	'	response.redirect("receipt.asp")
	'end if
		
	
	
	
	'payment notification from PapPal.
	'Do not rely on session variables. They could expire by the time they come to this page.
	'Generate all info from this order to write to DB Sent and Receive
	if len(trim(invoice_pp)) > 0 then
		session("ordernumber") = cstr(invoice_pp)
		
	  shiptotal = FORMATCURRENCY(sitelink.ORD_SHIPPING(session("ordernumber"),false))
	  taxtotal  = FORMATCURRENCY(sitelink.ORD_TAX(session("ordernumber"),false)) 
	  paypalaccountid = sitelink.Get_PayPalLogin()
	  paypalcurrency  =sitelink.Get_PayPalCurrencyCode()
	  

	  shipping_for_Paypal = mid(shiptotal,2,len(shiptotal))
	  taxtotal_for_Paypal =  mid(taxtotal,2,len(taxtotal))
	  
	  successurl = session("secureurl") + cstr("processorder.asp")
	  cancelurl = session("secureurl") + cstr("checkout.asp")
	  notifyurl = session("secureurl") + cstr("processorder.asp")	 
	  imageurl  = session("secureurl") + cstr("images/TopR1-C1.gif")
	  imageurl  = session("secureurl") + "/images/"+ cstr(COMPANY_LOGO_IMG)
	  
	  'get billto
	  'billtocust = sitelink.GetBilltoCust(session("ordernumber"))
	  'set shopperRecord= sitelink.SHOPPERINFO(cstr(billtocust),false)
	  'session("shopperid") = shopperRecord.custnum
	  'session("firstname")= trim(shopperRecord.FirstName)
	  'session("lastname")= trim(shopperRecord.Lastname)
	  'session("address1")= trim(shopperRecord.addr)
	  'session("address2")= trim(shopperRecord.addr2)
	  'session("city")= trim(shopperRecord.city)
	  'session("state")= trim(shopperRecord.state)
	  'session("zipcode")= trim(shopperRecord.zipcode)
	  'session("email")= trim(shopperRecord.email)
	  'session("country")= trim(shopperRecord.country)
	  'set shopperRecord= nothing
	  'first_name = session("firstname")
	  'last_name=  session("lastname")
	  'address1 =session("address1")
	  'address2 =session("address2")
	  'city = session("city")
	  'state = session("state")
	  'zip= session("zipcode")
	  
	    extrafield =""
	    xmlstring = sitelink.getbasketinfo(session("shopperid"),session("ordernumber"),extrafield,false,session("user_want_multiship"))
		objDoc.loadxml(xmlstring)
		set SL_Basket = objDoc.selectNodes("//gbi")
		tempstr = ""
		for x=0 to SL_Basket.length-1		
			 SL_Number	= SL_Basket.item(x).selectSingleNode("item").text
			 SL_BasketDesc		= SL_Basket.item(x).selectSingleNode("inetsdesc").text
			 SL_BasketDesc      = replace(SL_BasketDesc,"""","")
			 if SL_Number = "" then
				SL_BasketDesc ="Order Promotion Discount" 
			 end if 
			 SL_BasketUnitPrice	= formatcurrency(SL_Basket.item(x).selectSingleNode("it_unlist").text)
			 SL_BasketUnitPrice = mid(SL_BasketUnitPrice,2,len(SL_BasketUnitPrice))
			 SL_BasketQauntity	= SL_Basket.item(x).selectSingleNode("quanto").text
			 SL_BasketQauntity = replace(SL_BasketQauntity,".00","")			
			 SL_Basket_Shipto	= SL_Basket.item(x).selectSingleNode("ship_to").text
			 SL_Basket_Shipto = SL_Basket_Shipto + 1 - 1
		 
			 tempstr = tempstr + + VbCrLf + "<input type=hidden name=item_name_"+ cstr(x+1)+" value="+cstr(SL_BasketDesc)+">" + VbCrLf
			 tempstr = tempstr + "<input type=hidden name=amount_"+cstr(x+1)+" value="+cstr(SL_BasketUnitPrice)+">" + VbCrLf
			 tempstr = tempstr + "<input type=hidden name=quantity_"+cstr(x+1)+" value="+cstr(SL_BasketQauntity)+">" + VbCrLf
			 tempstr = tempstr + "<input type=hidden name=item_number_"+ cstr(x+1)+" value="+cstr(SL_Number)+">" + VbCrLf
			 
			 if SL_Basket_Shipto > 0 then 
		 		 fname = SL_Basket.item(x).selectSingleNode("firstname").text
				 lname = SL_Basket.item(x).selectSingleNode("lastname").text
				 fullname = (cstr(fname) + " " + cstr(lname))
				 tempstr = tempstr + "<input type=hidden name=on0_"+cstr(x+1)+" value=Shipping to "+fullname+">" + VbCrLf
			 end if  
		next
		set SL_Basket=nothing
		
		
		
		PayPalStr="<input type=hidden name=cmd value=_cart>" + VbCrLf
		PayPalStr= PayPalStr + "<input type=hidden name=mrb value=R-1V748593L7668510W>" + VbCrLf
		PayPalStr= PayPalStr + "<input type=hidden name=pal value=R-NCY5RMVCQZZDG>" + VbCrLf
		PayPalStr= PayPalStr + "<input type=hidden name=bn value=Dydacomp>" + VbCrLf
		PayPalStr= PayPalStr + "<input type=hidden name=business value="+ cstr(paypalaccountid) +">" + VbCrLf
		PayPalStr= PayPalStr + "<input type=hidden name=upload value=1>" 
		PayPalStr= PayPalStr +  + VbCrLf + tempstr + VbCrLf
		PayPalStr=  PayPalStr + "<input type=hidden name=shipping_1 value="+ cstr(shipping_for_Paypal)+">" + VbCrLf
		PayPalStr = PayPalStr + VbCrLf + "<input type=hidden name=currency_code value="+cstr(paypalcurrency)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=invoice value="+cstr(session("ordernumber"))+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=return value="+cstr(successurl)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=cancel value="+cstr(cancelurl)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=cancel value="+cstr(cancelurl)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=notify_url value="+cstr(notifyurl)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=image_url  value="+cstr(imageurl)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=first_name value="+cstr(first_name)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=last_name value="+cstr(last_name)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=address1 value="+cstr(address1)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=address2 value="+cstr(address2)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=city value="+cstr(city)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=state value="+cstr(state)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=zip value="+cstr(zip)+">" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=no_note value=1>" + VbCrLf
		PayPalStr = PayPalStr +  "<input type=hidden name=rm value=2>" + VbCrLf
	
		'this will be the sent data string to PP.
		session("PayPalStr") = PayPalStr
		
		'now capture the return from PP.
		PayPal_ret_Str =""
	 	For Each name In Request.Form
			formName = name
			formValue = Request.Form(formName)
			formValue = replace(formValue,"""","")
			strContent = strContent&formName&":"&formValue&"<br>"
			PayPal_ret_Str = PayPal_ret_Str + + VbCrLf + "<input type=hidden id="""" name="+cstr(formName) + " value="+cstr(formValue)+">"
		Next
		
		session("PayPal_ret_Str") = PayPal_ret_Str
		
		session("PayPal_AuthAmt") = cstr(REQUEST.FORM("mc_gross"))
		session("PayPal_Trans_id") = cstr(REQUEST.FORM("txn_id"))
		if session("PayPal_Trans_id") = "0" then
			session("PayPal_Trans_id")=""
		end if
			
		session("PayPalBuyerID") = cstr(REQUEST.FORM("payer_email"))
		session("PayPalPayerID") = cstr(REQUEST.FORM("payer_id"))
		call sitelink.savedata_PP(session("ordernumber"))
		session("paymentmethod") ="PP"
		'put order hold date when Paypal
		if PAYPAL_HOLD_DAYS > 0 then
			paypal_holdate = cstr(DateAdd("D",PAYPAL_HOLD_DAYS,date()))
			if isdate(session("orderholdDate")) = true then
				diff_days = cdate(paypal_holdate)-cdate(session("orderholdDate"))
				if diff_days > 0 then
					session("orderholdDate")= paypal_holdate
				end if
			else
				session("orderholdDate")= paypal_holdate
			end if	
			call sitelink.ORDER_HOLD_DATE(session("ordernumber"),"WRITE",cstr(session("orderholdDate")))
		end if
		
		'session("ordererrormessage")=sitelink.checkorderinfo(session("shopperid"),session("ordernumber"),session("cc_name"),session("cc_number"),session("cc_type"),session("cc_expmonth"),session("cc_expyear"),session("paymentmethod"),session("giftmsg1"),session("giftmsg2"),session("giftmsg3"),session("giftmsg4"),session("giftmsg5"),session("giftmsg6"),session("ponum"),"","",session("memo_field1"),session("memo_field2"),session("memo_field3"), cstr(session("orderholdDate")),cstr(""),cstr(""))	

	end if	
	
	
	'send email to seller	
	if session("paymentmethod") ="PP" and (session("PayPal_Trans_id")="" or session("PayPal_Trans_id") = "0") then		
			if len(trim(AUTHNET_ERR_NOTIFY))>0 then
				messsg ="Weborder number #" + cstr(session("ordernumber")) +" " +_
						 "does not have a transaction id returned by Paypal." +_
						 VbCrLf + VbCrLf +_
						 "You need to go into your Paypal account and manually accept the payment." +_
						 VbCrLf + VbCrLf +_
						 "Thank you."		
				DataToSend ="msg="+ cstr(server.urlencode(messsg))
				DataToSend = DataToSend & "&email="  & cstr(AUTHNET_ERR_NOTIFY)
				DataToSend = DataToSend & "&sitename="  & cstr(session("insecureurl"))
				DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
				set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
				xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink600/authnetnotify.asp",false
				xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				xmlhttp.setRequestHeader "Content-Type", "charset=iso-8859-1"		
				xmlhttp.send DataToSend
				'Response.Write xmlhttp.ResponseText
				set xmlHTTP=nothing
			end if 
	end if
	
	

	extrafield =""
	xmlstring = sitelink.getbasketinfo(session("shopperid"),session("ordernumber"),extrafield,false,session("user_want_multiship"))	
	
	objDoc.loadxml(xmlstring)
	basket_has_multiship = false
	filter_condition = "//gbi[ship_to>0]"
	set SL_has_multishipto_basket = objDoc.selectNodes(filter_condition)
		if SL_has_multishipto_basket.length > 0 then
			basket_has_multiship= true
		end if		
	set SL_has_multishipto_basket = nothing

	basket_has_giftcert = false
	giftcertitem=""
	giftcertAmount=0
      filter_condition = "//gbi[certid>0]"
    	set SL_Basket_giftcert = objDoc.selectNodes(filter_condition)
    		if SL_Basket_giftcert.length > 0 then
				basket_has_giftcert = true
				giftcertitem = SL_Basket_giftcert.item(0).selectSingleNode("item").text
				giftcertAmount = SL_Basket_giftcert.item(0).selectSingleNode("it_unlist").text
				giftcertAmount = abs(giftcertAmount)
			end if 
		set SL_Basket_giftcert = nothing
		
	'set SL_Basket = objDoc.selectNodes("//gbi")
    if basket_has_giftcert=true then
		    filter_condition = "//gbi[item !='"+cstr(giftcertitem)+"']"					    
	else
	        filter_condition = "//gbi"
	end if
	
	set SL_Basket = objDoc.selectNodes(filter_condition)
	

	' now update the wishlist database
	if session("registeredshopper")="YES" and WANT_WISHLIST =1 then
		'desc will have WISHLIST in basket db
		filter_condition = "//gbi[desc='WISHLIST' and certid=0]"
		has_wishlistitem= false
		set SL_has_wishlist_item = objDoc.selectNodes(filter_condition)
		if SL_has_wishlist_item.length > 0 then
			for x= 0 to SL_has_wishlist_item.length-1
				SL_Number = SL_Basket.item(x).selectSingleNode("number2").text
			    SL_Variant	= SL_Basket.item(x).selectSingleNode("variant").text
				SL_BasketQuantity	= SL_Basket.item(x).selectSingleNode("quanto").text
				SL_BasketQuantity = SL_BasketQuantity + 1 -1
				SL_item_id=0
				extra = ""
				call sitelink.WISHLIST("QTYUPDATE",cstr(SL_Number),SL_BasketQuantity,SL_item_id,cstr(SL_Variant), cstr(extra),session("addressbookcustnum"))
			next		
		end if
		set SL_has_wishlist_item=nothing
	
	end if



'  if session("registeredshopper")="YES" and WANT_WISHLIST =1 then
'	xmlstringwls =sitelink.WISHLIST("LIST","0",0,0,"",session("addressbookcustnum"))
'	 objDoc.loadxml(xmlstringwls)
	
'	for x=0 to SL_Basket.length-1
'		SL_item_id          = SL_Basket.item(x).selectSingleNode("record").text
'		SL_BasketQauntity	= SL_Basket.item(x).selectSingleNode("quanto").text
'		SL_Number			= SL_Basket.item(x).selectSingleNode("number2").text
'		SL_Variant			= SL_Basket.item(x).selectSingleNode("variant").text
'
'		targetlistNode = "//wls/item_id[.="+SL_item_id+"]"
'		
'		set SL_Wishlist = objDoc.selectNodes(targetlistNode)	  
'			if SL_Wishlist.length > 0 then
'				SL_item_id = SL_item_id + 1 - 1
'				SL_BasketQauntity = SL_BasketQauntity + 1 -1
'				extra = ""
'				call sitelink.WISHLIST("QTYUPDATE",cstr(SL_Number),SL_BasketQauntity,SL_item_id,cstr(SL_Variant), cstr(extra))
'			end if
'		set SL_Wishlist = nothing	
'		
'	next

 ' end if
  
	  'update best seller  
		for x=0 to SL_Basket.length-1
		'SL_item_id          = SL_Basket.item(x).selectSingleNode("record").text
		SL_BasketQuantity	= SL_Basket.item(x).selectSingleNode("quanto").text
		'SL_Number			= SL_Basket.item(x).selectSingleNode("number2").text
		SL_Number			= SL_Basket.item(x).selectSingleNode("item").text
		SL_Variant			= SL_Basket.item(x).selectSingleNode("variant").text
		SL_desc			    = SL_Basket.item(x).selectSingleNode("basketdesc").text
		SL_certid           =0
		SL_certid           =SL_Basket.item(x).selectSingleNode("certid").text
			
		if SL_Variant<>"" then
			SL_Variant_flag = true
		else
			SL_Variant_flag = false
		end if

		if SL_Number <> "" and SL_desc<>"PROMO" and SL_certid=0 then
			call sitelink.UPDATEBESTSELLER(session("ordernumber"),cstr(SL_Number),cstr(SL_BasketQuantity),SL_Variant_flag)	 
		end if
		
		'update live inventory.
		if LIVE_INVENT= 1 and SL_Number <> "" then
			call sitelink.UPDATE_STOCK_UNITS(cstr(SL_Number),cstr(SL_BasketQuantity))
		end if
	next 
  
 	

' write important session variables to file		
StrFileName = "text\sess_txt.txt"
StrPhysicalPath = Server.MapPath(StrFileName)
Set fs=Server.CreateObject("Scripting.FileSystemObject")
set ts = fs.OpenTextFile(StrPhysicalPath , 8)
ts.write(session("ordernumber") & "," & session("shopperid") & "," & session("sfirstname") & "," & session("slastname") & ",")
ts.write(session("scompany") & "," & session("saddress1") & "," & session("saddress2") & "," & session("scity") & "," & "," )
ts.write(session("sstate") & session("szipcode") & "," & session("scountry") & "," & session("sphone") & "," & session("semail") & ",")
ts.write(session("firstname") & "," & session("lastname") & "," & session("company") & "," & session("address1") & "," & session("address2") & ",")
ts.write(session("city") & "," & session("state") & "," & session("zipcode") & "," & session("country") & "," & session("phone") & "," & session("email") & ",")
ts.write(session("billtocopy"))
ts.writeline()
ts.close
set ts=nothing
set fs=nothing

'REM CLEAR THE cookie for this order
cookiename=session("shortstorename")+"order"
RESPONSE.COOKIES(cookiename).Expires="January 1, 1997"

if isnumeric(session("restoreoldordernumber"))=false then
	session("restoreoldordernumber")=0
end if

call sitelink.processorder(session("shopperid"),session("ordernumber"),session("restoreoldordernumber"))

 
		points_used = 0 
		points_amt = 0 	 
		set ORDER_RECORD = sitelink.ORDER_RECORD(session("ordernumber"),false)
			points_used = ORDER_RECORD.points_usd
			points_amt = ORDER_RECORD.checkamoun
			ord_total  = ORDER_RECORD.ord_total			  
			remaining_balance = ord_total-points_amt	
		set ORDER_RECORD=nothing

'now fire email confirmation with order detail
	
	dim DataToSend, xmlHTTP
		
		DataToSend ="OrderNum="& session("ordernumber")
		DataToSend = DataToSend & "&basketcount="  & SL_Basket.length
		DataToSend = DataToSend & "&email="  & session("email")		
		DataToSend = DataToSend & "&Subtotal="  & formatcurrency(sitelink.ORD_SUBTOTAL(session("ordernumber"),false))
		DataToSend = DataToSend & "&tax="  & formatcurrency(sitelink.ORD_TAX(session("ordernumber"),false))
		DataToSend = DataToSend & "&shipping="  & formatcurrency(sitelink.ORD_SHIPPING(session("ordernumber"),false))	
		'DataToSend = DataToSend & "&OrdTotal="  & formatcurrency(sitelink.ORD_TOTAL(session("ordernumber"),false)) 
		DataToSend = DataToSend & "&OrdTotal="  & formatcurrency(ord_total) 
		DataToSend = DataToSend & "&points_amt="  & formatcurrency(points_amt)
		DataToSend = DataToSend & "&points_used="  & points_used
		DataToSend = DataToSend & "&remaining_balance="  & formatcurrency(remaining_balance)
		DataToSend = DataToSend & "&basket_has_giftcert="  & basket_has_giftcert
		DataToSend = DataToSend & "&giftcertAmount="  & formatcurrency(giftcertAmount)
		DataToSend = DataToSend & "&remaining_GCBalance="  & formatcurrency(0)

		default_ship_title = trim(sitelink.dataforstatusnew("SHIPPINGMETHOD",session("ordernumber"),session("shopperid"),"EMAIL",false))
	  
		'billing
		DataToSend = DataToSend & "&bFname="  & server.urlencode(session("firstname"))
		DataToSend = DataToSend & "&bLname="  & server.urlencode(session("lastname"))
		baddress =   session("address1") &" "& session("address2")
		DataToSend = DataToSend & "&baddress="  & server.urlencode(baddress)
		DataToSend = DataToSend & "&bcity="  & server.urlencode(session("city"))
		DataToSend = DataToSend & "&bstate="  & session("state")
		DataToSend = DataToSend & "&bzipcode="  & session("zipcode")
		DataToSend = DataToSend & "&bcountry="  & server.urlencode(sitelink.countryname(session("country")))

        

	
	if basket_has_multiship = false then
		has_multiship = 0
				
		DataToSend = DataToSend & "&multiship=0" 
		DataToSend = DataToSend & "&billtocopy="  & session("billtocopy")
		'shipping 
		DataToSend = DataToSend & "&sFname="  & server.urlencode(session("sfirstname"))
		DataToSend = DataToSend & "&sLname="  & server.urlencode(session("slastname"))
		saddress =   session("saddress1") &" "& session("saddress2")
		DataToSend = DataToSend & "&saddress="  & server.urlencode(saddress)
		DataToSend = DataToSend & "&scity="  & server.urlencode(session("scity"))
		DataToSend = DataToSend & "&sstate="  & session("sstate")
		DataToSend = DataToSend & "&szipcode="  & session("szipcode")
		DataToSend = DataToSend & "&scountry="  & server.urlencode(sitelink.countryname(session("scountry")))
		
		for x=0 to SL_Basket.length-1	
	 			SL_Number			= SL_Basket.item(x).selectSingleNode("number2").text
				SL_Variant			= SL_Basket.item(x).selectSingleNode("variant").text
				item_number 		= SL_Number +" "+ SL_Variant								
				SL_BasketQuantity	= SL_Basket.item(x).selectSingleNode("quanto").text
				SL_BasketQuantity   = replace(SL_BasketQuantity,".00","")
				SL_BasketDesc		= SL_Basket.item(x).selectSingleNode("inetsdesc").text
				SL_BasketUnitPrice	= SL_Basket.item(x).selectSingleNode("it_unlist").text
				SL_BasketExtPrice	= SL_Basket.item(x).selectSingleNode("extendp").text
				SL_ShipVia 			= SL_Basket.item(x).selectSingleNode("ship_via").text
				SL_ShippingTitle  = ""
				if SL_Number="" then
					SL_BasketDesc="Order Promotion Discount"
				end if

				if len(trim(SL_ShipVia)) > 0 then
					SL_ShippingTitle = SL_Basket.item(x).selectSingleNode("ca_title").text
				else
					SL_ShippingTitle = default_ship_title
				end if 
				
				DataToSend = DataToSend & "&item="  & item_number				
				DataToSend = DataToSend & "&desc="  & server.urlencode(SL_BasketDesc)
				DataToSend = DataToSend & "&qty="  & SL_BasketQuantity
				DataToSend = DataToSend & "&unitprice="  & formatcurrency(SL_BasketUnitPrice)
				DataToSend = DataToSend & "&price="  & formatcurrency(SL_BasketExtPrice)
				DataToSend = DataToSend & "&ship_via="  & server.urlencode(SL_ShippingTitle)

				
		next 
	else
		has_multiship= 1
		DataToSend = DataToSend & "&multiship=1" 
		DataToSend = DataToSend & "&billtocopy="  & session("billtocopy")
		for x=0 to SL_Basket.length-1
				SL_Number			= SL_Basket.item(x).selectSingleNode("number2").text
				SL_Variant			= SL_Basket.item(x).selectSingleNode("variant").text
				item_number 		= SL_Number +" "+ SL_Variant								
				SL_BasketQuantity	= SL_Basket.item(x).selectSingleNode("quanto").text
				SL_BasketQuantity   = replace(SL_BasketQuantity,".00","")
				SL_BasketDesc		= SL_Basket.item(x).selectSingleNode("inetsdesc").text
				SL_BasketUnitPrice	= SL_Basket.item(x).selectSingleNode("it_unlist").text
				SL_BasketExtPrice	= SL_Basket.item(x).selectSingleNode("extendp").text
				SL_ShipVia 			=SL_Basket.item(x).selectSingleNode("ship_via").text
				SL_ShippingTitle  = ""

				if SL_Number="" then
					SL_BasketDesc="Order Promotion Discount"
				end if
				if len(trim(SL_ShipVia)) > 0 then
					SL_ShippingTitle = SL_Basket.item(x).selectSingleNode("ca_title").text
				else
					SL_ShippingTitle = default_ship_title
				end if 
				
				DataToSend = DataToSend & "&item="  & item_number				
				DataToSend = DataToSend & "&desc="  & server.urlencode(SL_BasketDesc)
				DataToSend = DataToSend & "&qty="  & SL_BasketQuantity
				DataToSend = DataToSend & "&unitprice="  & formatcurrency(SL_BasketUnitPrice)
				DataToSend = DataToSend & "&price="  & formatcurrency(SL_BasketExtPrice)
				DataToSend = DataToSend & "&ship_via="  & server.urlencode(SL_ShippingTitle)
				
				SL_Basket_Shipto	= SL_Basket.item(x).selectSingleNode("ship_to").text
				if SL_Basket_Shipto > 0 then
					
				 	fname				= SL_Basket.item(x).selectSingleNode("firstname").text
				    lname				= SL_Basket.item(x).selectSingleNode("lastname").text
					company			    = SL_Basket.item(x).selectSingleNode("company").text
					addr				= SL_Basket.item(x).selectSingleNode("addr").text
					addr2				= SL_Basket.item(x).selectSingleNode("addr2").text
					city				= SL_Basket.item(x).selectSingleNode("city").text
					vstate				= SL_Basket.item(x).selectSingleNode("state").text
					zipcode			    = SL_Basket.item(x).selectSingleNode("zipcode").text
					country			    = SL_Basket.item(x).selectSingleNode("country").text
					
					DataToSend = DataToSend & "&fname="  & server.urlencode(fname)
					DataToSend = DataToSend & "&lname="  & server.urlencode(lname)
					DataToSend = DataToSend & "&company="  & server.urlencode(company)
					DataToSend = DataToSend & "&addr="  & server.urlencode(addr)
					DataToSend = DataToSend & "&addr2="  & server.urlencode(addr2)
					DataToSend = DataToSend & "&city="  & server.urlencode(city)
					DataToSend = DataToSend & "&vstate="  & server.urlencode(vstate)
					DataToSend = DataToSend & "&zipcode="  & server.urlencode(zipcode)
					DataToSend = DataToSend & "&country="  & server.urlencode(sitelink.countryname(country))
					DataToSend = DataToSend & "&ship_to=" & SL_Basket_Shipto
				else
					DataToSend = DataToSend & "&fname="  
					DataToSend = DataToSend & "&lname="
					DataToSend = DataToSend & "&company=" 
					DataToSend = DataToSend & "&addr="  
					DataToSend = DataToSend & "&addr2="  
					DataToSend = DataToSend & "&city="  
					DataToSend = DataToSend & "&vstate=" 
					DataToSend = DataToSend & "&zipcode=" 
					DataToSend = DataToSend & "&country=" 
					DataToSend = DataToSend & "&ship_to=0" 
				end if
		
		
		next
	
	end if
	
	

	set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink600/orderconfirmation.asp",false
	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlhttp.setRequestHeader "Content-Type", "charset=iso-8859-1"
	xmlhttp.send DataToSend
	'Response.Write xmlhttp.ResponseText
	'Response.Write DataToSend
	Set xmlhttp = nothing

	

set SL_Basket = nothing
Set objDoc= nothing
%>

<%
	
	'NOW NEW ORDER
	session("previousorder") = session("ordernumber")
	session("previousbilltocopy") = session("billtocopy")
	session("previoususer_want_multiship") = session("user_want_multiship") 
	
	'now fire neworder
	session("shopperid")=sitelink.initshopper("","") 
	'session("clshopper")="S"+cstr(session("shopperid"))
	session("ordernumber")=sitelink.initorder("") 
	call sitelink.createcustomer(session("shopperid"),"")
	'session("needtocheckorderinfo")="YES"
	SAVEDSHOPPERID=""
	session("returnvisitor")=1
	
	session("cardwasauthed")=false
	
	if session("registeredshopper")="YES" then	
		'session("regemail")= regemail
		'session("password")=regpassword
		'response.write("New shopper -->"&session("shopperid"))	
		set SLShopper = New cShopperInfo
		call InitcShopperInfo(SLShopper)
			SLShopper.shopperid=session("shopperid")
			SLShopper.bemail=session("regemail")
			SLShopper.bpassword=session("password")	
			SAVEDSHOPPERID=sitelink.verifyshopper3(SLShopper)	
			
			if SAVEDSHOPPERID<>"-1" then
				session("firstname")=SLShopper.bfirstname
				session("lastname")=SLShopper.blastname
				session("company")=SLShopper.bcompany
				session("address1")=SLShopper.baddr1
				session("address2")=SLShopper.baddr2
				session("address3")=SLShopper.baddr3
				session("city")=SLShopper.bcity
				session("state")=SLShopper.bstate
				session("zipcode")=SLShopper.bzipcode
				session("phone")=SLShopper.bphone
				session("email")=SLShopper.bemail
				session("country")=SLShopper.bcountry
				session("pointavail")=SLShopper.bpoints
				if SLShopper.bNopoints=true then
					session("pointavail")=0
				end if

			end if
		set SLShopper=nothing
	end if 
	set sitelink= nothing
	
	
	
	'clear this information
	session("SL_BasketCount") = 0
	session("SL_BasketSubTotal") = 0
	
	Session.Contents.Remove("showdept")
	Session.Contents.Remove("department")
	Session.Contents.Remove("topdept")
	Session.Contents.Remove("items_spaned")
	Session.Contents.Remove("shippingmethod")
	
	Session.Contents.Remove("sfirstname")
	Session.Contents.Remove("slastname")
	Session.Contents.Remove("scompany")
	Session.Contents.Remove("saddress1")
	Session.Contents.Remove("saddress2")
	Session.Contents.Remove("saddress3")	
	Session.Contents.Remove("scity")
	Session.Contents.Remove("sstate")
	Session.Contents.Remove("szipcode")
	Session.Contents.Remove("scountry")
	Session.Contents.Remove("sphone")
	Session.Contents.Remove("sphone2")
	Session.Contents.Remove("semail")
	
	Session.Contents.Remove("custerrormessage")
	Session.Contents.Remove("ordererrormessage")
	Session.Contents.Remove("Title_String")
	Session.Contents.Remove("prospecterror")
	Session.Contents.Remove("catalog")
	Session.Contents.Remove("billtocopy")
	Session.Contents.Remove("receiveemail")
	Session.Contents.Remove("cardwasauthed")

	
	Session.Contents.Remove("paymentmethod")
	Session.Contents.Remove("cc_name")
	Session.Contents.Remove("cc_number")
	Session.Contents.Remove("cc_id")
	Session.Contents.Remove("cc_type")
	Session.Contents.Remove("cc_expmonth")
	Session.Contents.Remove("cc_expyear")
	Session.Contents.Remove("accttype")
	Session.Contents.Remove("bankname")
	Session.Contents.Remove("bankrountingnum")
	Session.Contents.Remove("bankacctnum")	
	Session.Contents.Remove("PayPalStr")
	Session.Contents.Remove("PayPalBuyerID")
	Session.Contents.Remove("ordermemo")
	
	Session.Contents.Remove("from_date")
	Session.Contents.Remove("issue_num")

	Session.Contents.Remove("giftmsg1")
	Session.Contents.Remove("giftmsg2")
	Session.Contents.Remove("giftmsg3")
	Session.Contents.Remove("giftmsg4")
	Session.Contents.Remove("giftmsg5")
	Session.Contents.Remove("giftmsg6")

	
	Session.Contents.Remove("memo_field1")
	Session.Contents.Remove("memo_field2")
	Session.Contents.Remove("memo_field3")
	Session.Contents.Remove("ordermemo")
	
	Session.Contents.Remove("orderholdDate")
	
	Session.Contents.Remove("ponum")	
	Session.Contents.Remove("CurrentSourceKey")
	
	Session.Contents.Remove("Points_need")
	Session.Contents.Remove("Redeem_Amount")
	Session.Contents.Remove("DoNotShow_PointsTab")
	Session.Contents.Remove("PaybyPoints")
	Session.Contents.Remove("Remaining_balance")
	
	Session.Contents.Remove("DoNotShow_GCTab")
	Session.Contents.Remove("PaybyGC")
	Session.Contents.Remove("Remaining_BalanceAfterGC")
	
	
	
	session("cookiewritten") = false 
	session("user_want_multiship") = false
	
	session("askedforsourcekey")=-1
	session("sourcekeyisgood")=-1
	

set sitelink=nothing 

	

Response.Redirect("receipt.asp")
	
%>