<%on error resume next%>
<%
  if session("registeredshopper")="NO" then
  	response.redirect("login.asp")
  end if
 %>
 
<!--#INCLUDE FILE = "include/momapp.asp" -->

<%


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
		
		
		giftmsg1=cstr(REQUEST.FORM("txtgiftmsg1"))
		giftmsg2=cstr(REQUEST.FORM("txtgiftmsg2"))
		giftmsg3=cstr(REQUEST.FORM("txtgiftmsg3"))
		giftmsg4=cstr(REQUEST.FORM("txtgiftmsg4"))
		giftmsg5=cstr(REQUEST.FORM("txtgiftmsg5"))
		giftmsg6=cstr(REQUEST.FORM("txtgiftmsg6"))
		
	if REQUEST.FORM("deleteAddress").count = 1 then
		address_id = cstr(trim(REQUEST.FORM("address_id")))
		cintaddress_id = (address_id) + 1 -1
		call sitelink.DELETEADDRESSBOOK(cintaddress_id)
		
		set sitelink = nothing
		response.redirect("address.asp")
	end if 
		
	if REQUEST.FORM("addtobook") = 1 then
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
		session("prospecterror")=sitelink.ADDTOADDRESSBOOK(SLShopper)
		set SLShopper = nothing

		if isnumeric(session("prospecterror"))=false then
			has_error = true
		else
			has_error = false
		end if
		
		if has_error = true then
			set sitelink=nothing
			response.redirect("proserror.asp")
		end if
		
		whatbastnum = REQUEST.FORM("basketrecord")	
		intlastshipto = cstr(whatbastnum)
		'lastadress_id = session("lastshopper")
		lastadress_id = session("prospecterror")
		aa=sitelink.CHANGEBASKETSHIPTO(session("ordernumber"),intlastshipto,session("shopperid"),lastadress_id,giftmsg1,giftmsg2,giftmsg3,giftmsg4,giftmsg5,giftmsg6)
	  end if
	  
	  set sitelink=nothing
	response.redirect("basket.asp")
%>