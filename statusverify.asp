<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->

<%	

    'if REQUEST.servervariables("server_port_secure")<> 1 and SLsitelive=true then     
    if Request.ServerVariables("HTTPS") = "off" and SLsitelive=true then     
        set sitelink=nothing
        set ObjDoc=nothing
    	response.redirect(secureurl+"statuslogin.asp") 
    end if
 
	
	session("regemail") = cstr(trim(REQUEST.FORM("regtxtemail")))
	session("regpassword")=cstr(trim(REQUEST.FORM("txtregpassword")))
	

	session("regemail")=fix_xss_Chars(session("regemail"))
	session("regpassword")=fix_xss_Chars(session("regpassword"))

	set SLShopper = New cShopperInfo
	SLShopper.shopperid=session("shopperid")
	SLShopper.bemail=session("regemail")
	SLShopper.bpassword=session("regpassword")

	SAVEDSHOPPERID=sitelink.verifyshopper3(SLShopper)

	'SAVEDSHOPPERID=sitelink.verifyshopper3(session("shopperid"),session("regemail"),session("regpassword"))


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
	'response.redirect("custinfo.asp")
	set SLShopper = nothing
	session("addressbookcustnum") = sitelink.makeaddressbook(session("LASTNAME"),session("ZIPCODE"),SESSION("PASSWORD"),session("shopperid"))
	
    if ucase(session("destpage"))<>"BASKET.ASP" then
	    call SITELINK.quantitypricing(session("shopperid"),session("ordernumber"))
	    session("SL_BasketCount") = sitelink.GetData_Basket(cstr("BASKETCOUNT"),session("ordernumber"),session("SL_VatInc"))
	    session("SL_BasketSubTotal") = sitelink.GetData_Basket(cstr("BASKETSUBTOTAL"),session("ordernumber"),session("SL_VatInc"))
    end if 
    
	set sitelink=nothing
	response.redirect("accountinfo.asp")
else
	SAVEDSHOPPERID=""
	session("registeredshopper")="NO"
	session("failedtofindshopper")="YES"
	set sitelink=nothing
	set SLShopper = nothing	
	response.redirect("statuslogin.asp")
end if

%>

