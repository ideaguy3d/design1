
<%on error resume next%>

<%
  if session("registeredshopper")="NO" then
  	response.redirect("login.asp")
  end if
 
    
 %>
 
<!--#INCLUDE FILE = "include/momapp.asp" -->

<%

	if REQUEST.FORM("deleteAddress").count = 1 then
		address_id = cstr(trim(REQUEST.FORM("address_id")))
		cintaddress_id = (address_id) + 1 -1
		call sitelink.DELETEADDRESSBOOK(cintaddress_id)
		
		'change the basket shipto to zero
		'whatbastnum = REQUEST.FORM("basketrecord")	
		'intlastshipto = cstr(whatbastnum)
		'aa=sitelink.CHANGEBASKETSHIPTO(session("ordernumber"),intlastshipto,session("shopperid"),cint(0))

		set sitelink = nothing
		if len(trim(session("destpage"))) > 0 then
			response.redirect(session("destpage"))
		else		
			response.redirect("address.asp")
		end if
	end if 
	
		sfirstname=cstr(REQUEST.FORM("txtsfirstname"))
		slastname=cstr(REQUEST.FORM("txtslastname"))
		scompany=cstr(REQUEST.FORM("txtscompany"))
		saddress1=cstr(REQUEST.FORM("txtsaddress1"))
		saddress2=cstr(REQUEST.FORM("txtsaddress2"))
		saddress3=cstr(REQUEST.FORM("txtsaddress3"))
		scity=cstr(REQUEST.FORM("txtscity"))
		sstate=UCASE(cstr(REQUEST.FORM("txtsstate")))
		szipcode=cstr(REQUEST.FORM("txtszipcode"))
		scountry=cstr(REQUEST.FORM("shiptocountry"))
		scounty=cstr(REQUEST.FORM("txtscounty"))
		sphone=cstr(REQUEST.FORM("txtsphone"))
		semail=cstr(REQUEST.FORM("txtsemail"))
		
		address_id = REQUEST.FORM("address_id")
		cintaddress_id = (address_id) + 1 -1
		
		sfirstname=fix_xss_Chars(sfirstname)
		slastname=fix_xss_Chars(slastname)
		scompany=fix_xss_Chars(scompany)
		saddress1=fix_xss_Chars(saddress1)
		saddress2=fix_xss_Chars(saddress2)
		saddress3=fix_xss_Chars(saddress3)
		scity=fix_xss_Chars(scity)
		sstate=fix_xss_Chars(sstate)
		szipcode=fix_xss_Chars(szipcode)
		scountry=fix_xss_Chars(scountry)
		sphone=fix_xss_Chars(sphone)
		semail=fix_xss_Chars(semail)
		
		
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
		SLShopper.bAddressId=cintaddress_id	
					
		
		session("prospecterror")=sitelink.UpdateAddressBook(SLShopper)
		set SLShopper=nothing
		
		if isnumeric(session("prospecterror"))=false then
			has_error = true
		else
			has_error = false
		end if
		
		
		if has_error = true then
			set sitelink=nothing
			response.redirect("proserror.asp")
		end if
		set sitelink = nothing
		
		basketrecord = REQUEST.FORM("basketrecord")
		
		if isnumeric(basketrecord) =true  then
			response.redirect("FromAddress.asp?basketrecord=" & basketrecord )
		end if
		if len(trim(session("destpage"))) > 0 then
			response.redirect(session("destpage"))
		end if
		
		'response.write(session("destpage"))
		response.redirect("address.asp")


%>