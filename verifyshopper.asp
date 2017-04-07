<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->

<%	
    
    if REQUEST.servervariables("server_port_secure")<> 1 and SLsitelive=true then 
        set sitelink=nothing
        set ObjDoc=nothing
    	response.redirect(secureurl+"login.asp") 
    end if
	
	session("regemail") = cstr(trim(REQUEST.FORM("regtxtemail")))
	session("regpassword")=cstr(trim(REQUEST.FORM("txtregpassword")))
	

	session("regemail")=fix_xss_Chars(session("regemail"))
	session("regpassword")=fix_xss_Chars(session("regpassword"))
	set SLShopper = New cShopperInfo
	SLShopper.shopperid=session("shopperid")
	SLShopper.bemail=session("regemail")
	SLShopper.bpassword=session("regpassword")

	'SAVEDSHOPPERID=sitelink.verifyshopper3(session("shopperid"),session("regemail"),session("regpassword"))
	SAVEDSHOPPERID=sitelink.verifyshopper3(SLShopper)
	
if SAVEDSHOPPERID<>"-1" then
	session("registeredshopper")="YES"
	session("failedtofindshopper")="NO"
	' variables are populated by verify shopper3.. Just trim it
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
	session("bcounty")=SLShopper.bcounty
	session("password")=SLShopper.bpassword
	session("pointavail")=SLShopper.bpoints
	if SLShopper.bNopoints=true then
		session("pointavail")=0
	end if
	
	
	
	session("loginverified")="YES"	
	'call to address book in basket page
	session("addressbookcustnum") = sitelink.makeaddressbook(session("LASTNAME"),session("ZIPCODE"),SESSION("PASSWORD"),session("shopperid"))

	'need to trigger order promotion if custom logs in and goes to custinfo page.
	'i.e login or register before checkout	
	if ucase(session("destpage"))<>"BASKET.ASP" then
	    call SITELINK.quantitypricing(session("shopperid"),session("ordernumber"))
	    session("SL_BasketCount") = sitelink.GetData_Basket(cstr("BASKETCOUNT"),session("ordernumber"),session("SL_VatInc"))
	    session("SL_BasketSubTotal") = sitelink.GetData_Basket(cstr("BASKETSUBTOTAL"),session("ordernumber"),session("SL_VatInc"))
    end if 	
	
	set SLShopper =nothing
	'response.write(SAVEDSHOPPERID)
	'response.redirect("custinfo.asp")
	set sitelink= nothing
	'if session("Clicked_Multishipto")=true then
	'	session("user_want_multiship")=true
	'end if 		
	'response.redirect("loggedIn.asp")
	if len(session("destpage")) > 0 then
		response.redirect(session("destpage"))
	else
	    urltoredirect = insecureurl & "loggedIn.asp"
		response.redirect(urltoredirect)
	end if
else
	SAVEDSHOPPERID=""
	session("registeredshopper")="NO"
	session("failedtofindshopper")="YES"
	set sitelink= nothing
	set SLShopper =nothing	
	response.redirect("login.asp")
end if

%>

