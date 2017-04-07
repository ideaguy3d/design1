
<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->


<%	
		
	'response.write("current shopper -->"&session("shopperid"))	
	session("shopperid")=sitelink.initshopper(cstr(request.servervariables("REMOTE_HOST")),session("storeclosed")) 
	'session("clshopper")="S"+cstr(session("shopperid"))
	session("ordernumber")=sitelink.initorder(cstr(request.servervariables("REMOTE_HOST"))) 
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
			end if
		set SLShopper=nothing
	end if 
	set sitelink= nothing
	
	
	
	'clear this information
	session.Contents.Remove("SL_BasketCount")
	session.Contents.Remove("SL_BasketSubTotal")

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
	
	Session.Contents.Remove("cc_id")

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
	
	session("cookiewritten") = false 
	session("user_want_multiship") = false
	
	session("askedforsourcekey")=-1
	session("sourcekeyisgood")=-1
	
	if len(trim(request.QueryString("dest"))) > 0 then
		usethisurl = trim(request.QueryString("dest"))
	else
		usethisurl= insecureurl
	end if
	
	
	response.redirect(usethisurl)
	'server.transfer(usethisurl)
%>

