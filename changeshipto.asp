<%on error resume next%>
<%
  if session("registeredshopper")="NO" then
  	response.redirect("login.asp")
  end if

  whatbastnum = Request.QueryString("basketrecord")	
  intlastshipto = cstr(whatbastnum)
	if isnumeric(whatbastnum)=false then
		response.redirect("login.asp")
	end if 
	

    intlastshipto = intlastshipto + 1 - 1
    shopcust_id = 0
	'if intlastshipto > 0 then
	    shopcust_id = request.querystring("shopcust")	    
	    if isnumeric(shopcust_id)=false then
	        shopcust_id = 0 
	    end if	    
	'end if	

 %>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%Response.ExpiresAbsolute = #Feb 18,1998 13:26:26# 


   'if cstr(Request.QueryString("shipdrop"))="0" then
   if intlastshipto = 0 then 'just populate the ship to on custinfo
		'shopcust_id = request.querystring("shopcust")
		shopcust_id = shopcust_id + 1 -1
		'response.write(address_id)
		xmlAddressBookstring =sitelink.getaddressbook(session("addressbookcustnum"),shopcust_id)
		objDoc.loadxml(xmlAddressBookstring)
	  set SL_AddressBook = objDoc.selectNodes("//gab")  
	
	  for z=0 to SL_AddressBook.length-1 
	  
	  	 'custnum		= trim(SL_AddressBook.item(z).selectSingleNode("custnum").text)
		 session("slastname")	= SL_AddressBook.item(z).selectSingleNode("lastname").text
		 session("sfirstname")	= SL_AddressBook.item(z).selectSingleNode("firstname").text
		 session("scompany")	= SL_AddressBook.item(z).selectSingleNode("company").text
		 session("saddress1") 	= SL_AddressBook.item(z).selectSingleNode("addr").text
		 session("saddress2") 	= SL_AddressBook.item(z).selectSingleNode("addr2").text
		 session("saddress3") 	= SL_AddressBook.item(z).selectSingleNode("addr3").text
		 session("scity") 		= SL_AddressBook.item(z).selectSingleNode("city").text
		 session("sstate") 		= SL_AddressBook.item(z).selectSingleNode("state").text
		 session("szipcode") 	= SL_AddressBook.item(z).selectSingleNode("zipcode").text
		 session("sphone") 		= SL_AddressBook.item(z).selectSingleNode("phone").text
		 session("semail") 		= SL_AddressBook.item(z).selectSingleNode("email").text
		 session("address_id")	= SL_AddressBook.item(z).selectSingleNode("address_id").text
		 session("scountry")    = SL_AddressBook.item(z).selectSingleNode("country").text
		 session("scounty")    = SL_AddressBook.item(z).selectSingleNode("county").text
		 'shopcust	    = trim(SL_AddressBook.item(z).selectSingleNode("shopcust").text)	
		 
		 if session("sstate") ="" then
		 	session("sstate")="INT"
		 end if	
		   
	  next	  
	  
	  session("billtocopy")= 0 
	  set SL_AddressBook=nothing
	  set objDoc = nothing

	set sitelink=nothing
	Response.Redirect("custinfo.asp")
   
   		'aa=sitelink.changebasketshipto(session("ordernumber"),intlastshipto,session("shopperid"),Cint(0))
   else
   	'address_id = cstr(trim(request.querystring("shopcust")))
	'cintaddress_id = (address_id) + 1 -1
	'cintaddress_id = request.querystring("shopcust")
	cintaddress_id = shopcust_id
	cintaddress_id = cintaddress_id + 1 - 1
	'whatbastnum = Request.QueryString("basketrecord")	
	'intlastshipto = cstr(whatbastnum)
	aa=sitelink.CHANGEBASKETSHIPTO(session("ordernumber"),cstr(intlastshipto),session("shopperid"),cintaddress_id)
    'response.write(intlastshipto)
'	response.write("hhh")
	set sitelink=nothing
	Response.Redirect("basket.asp")
   end if 
	
	
 %>
