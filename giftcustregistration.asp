<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->

<%
   ' Information about the shopper
	session("gfirstname")=cstr(REQUEST("txtfirstname"))
	session("glastname")=cstr(REQUEST("txtlastname"))
	session("giftregemail") = cstr(trim(REQUEST("txtemail")))	
	
	tempError = ""
	if len(trim(REQUEST("txtemail"))) = 0 then 
		tempError = "<LI>Email address cannot be blank."
	else
		pos1 = InStr(REQUEST("txtemail"),"@")
		pos2 = InStr(pos1,REQUEST("txtemail"),".")
		totLen = len(trim(REQUEST("txtemail")))
		
		if (InStr(REQUEST("txtemail"),"@") = 0) or (InStr(REQUEST("txtemail"),".") = 0) or ((pos2-pos1) < 1) or ((totLen-pos2) < 3) then 
			tempError = "<LI>Email address is incorrect."
		end if
		
	end if
'	if len(trim(Request("txtregpassword")))=0 then
'		tempError = tempError & "<LI>Password cannot be blank"
'	end if 
'	if len(trim(Request("txtregpassword2")))=0 then
'		tempError = tempError & "<LI>Re-Enter Password cannot be blank"
'	end if 
	if UCase(trim(Request("txtregpassword"))) <> UCase(trim(Request("txtregpassword2"))) then
		tempError = tempError & "<LI>Password and Re-Enter Password does not match"
	end if 
	
	if tempError <> "" then tempError = tempError & "</UL>"

' 	session("prospecterror")=sitelink.savecustomerinfo(session("sfirstname"),session("slastname"),session("scompany"),session("saddress1"),session("saddress2"),session("scity"),session("sstate"),session("szipcode"),session("sphone"),session("semail"),session("firstname"),session("lastname"),session("company"),session("address1"),session("address2"),session("city"),session("state"),session("zipcode"),session("phone"),session("email"),1,"   ",session("shopperid"),session("referral"),session("scountry"),session("country"))
 '	
'	if len(trim(tempError)) > 0 then
'		if len(trim(session("prospecterror"))) = 0 then 
'			session("prospecterror") = "<UL type=round>" & tempError
'		else
'			session("prospecterror") = replace(session("prospecterror"),"</ul>",tempError)
'		end if
'	end if
'	session("prospecterror") = replace(session("prospecterror"),"Bill-to","")
	
'	aa=session("prospecterror")	
'	if len(trim(aa))>0 then
'		response.redirect("proserror.asp")
'	end if
	
'	session("prospecterror")=sitelink.saveprospectinfo(session("firstname"),session("lastname"),session("company"),session("address1"),session("address2"),session("city"),session("state"),session("zipcode"),session("phone"),session("email"),session("catalog"),session("shopperid"),session("referral"),session("country"))
'	aa=session("prospecterror")	
'	if len(trim(aa))>0 then
'		response.redirect("proserror.asp")
'	end if
	
	
	
	' Now Regsiter the shopper. capture these values and use in savecollge.asp
'	session("reglastname")=session("lastname")
'	session("regzipcode")=session("zipcode")
'	session("regpassword")=cstr(request("txtregpassword"))
'	session("regpassword2")=cstr(request("txtregpassword2"))
'	session("registererror")=sitelink.registershopper(session("clshopper"),session("reglastname"),session("regzipcode"),SESSION("regpassword"),SESSION("regpassword2"),session("ordernumber"))
'	if len(trim(session("registererror"))) >0 then
'	  session("prospecterror")= "<ul type=""ROUND""><LI>"&session("registererror")&"</ul>"
 ' 	  session("registeredshopper")="NO"
'	  session("failedtoregistershopper")="YES"
'	  response.redirect("proserror.asp")
'	end if 
	
	' GOOD TIME TO WRITE OUT A COOKIE
'	session("cookiename")=ShortStoreName+"SiteLINK"
'	session("key_name")=sitelink.writecookie("shopper","lastname",session("reglastname"),1)
'	session("key_word")=sitelink.writecookie("shopper","password",session("regpassword"),1)
'	session("key_zip")=sitelink.writecookie("shopper","zipcode",session("regzipcode"),1)
'	session("key_expires")=sitelink.writecookie("shopper","expire","",0)
'	RESPONSE.COOKIES(session("cookiename")).Expires=session("key_expires")
'	RESPONSE.COOKIES(session("cookiename"))("WORD")=session("key_word")
'	RESPONSE.COOKIES(session("cookiename"))("NAME")=session("key_name")
'	RESPONSE.COOKIES(session("cookiename"))("ZIP")=session("key_zip")

'	session("cookiename")=""
'	session("key_name")=""
'	session("key_word")=""
'	session("key_zip")=""
'	session("key_expires")=""


	session("registeredshopper")="YES"
	'session("wishlistID") = sitelink.WISHCUST(session("shopperid"),false,cstr(""))
	eventdate=""
	if len(trim(REQUEST("txteventdate"))) > 0 and IsDate(REQUEST("txteventdate")) = true then
		eventday=Day(REQUEST("txteventdate"))
		if len(eventday) = 1 then eventday ="0" & eventday
		eventmonth=Month(REQUEST("txteventdate"))
		if len(eventmonth) = 1 then eventmonth ="0" & eventmonth
		eventyear=Year(REQUEST("txteventdate"))
		eventdate=eventmonth&"/"&eventday&"/"&eventyear
	end if
'Response.Write(session("glastname"))
'Response.Write(cstr(eventdate))
	session("giftcustID") = sitelink.MAKEGIFTCUST(session("gfirstname"),session("glastname"),cstr(REQUEST("txtcompany")),cstr(REQUEST("txtaddress1")),cstr(REQUEST("txtaddress2")),cstr(REQUEST("txtcity")),cstr(REQUEST("txtstate")),cstr(REQUEST("txtzipcode")),cstr(REQUEST("txtphone")),cstr(REQUEST("txtemail")),cstr(REQUEST("txtcountry")),cstr(REQUEST("txtreg2fname")),cstr(REQUEST("txtreg2lname")),"","",cstr(REQUEST("txtregistry")),cstr(eventdate),cstr(REQUEST("txtregpassword")),session("shopperid"),cstr(REQUEST("txtregdesc")))

'Response.Write(session("giftcustid")) & "," & cstr(REQUEST("txteventdate"))


	if left(ucase(session("giftcustID")),3)<>"<UL" then
	
	'Now send an email to the client when gift registration is complete.
	 if cint(session("giftcustID"))=0 then
		session("prospecterror")="<ul>This information is already registerd.</ul><ul>Please go to <a href=""giftreglogin.asp?dest=giftlist.asp"">Edit Registry</a> link to locate your registry.<br><br>OR</ul>"
		
		set sitelink=nothing 		
		response.redirect("proserror.asp")
	 else
		 session("giftregcustnum")=session("giftcustID")
		 session("giftregcustnum")=session("giftregcustnum")+ 1 - 1
		 
		
			set sitelink=nothing 
			session("strmessage") = "Thank you for registering with " & althomepage
			session("Title_String")="<a href=""giftregistry.asp""><font class=""TopNavRow2Text"">Gift Registry</font></a>&nbsp;>&nbsp;<font class=""TopNavRow2Text"">Registration</font>"

			response.redirect("thankyou.asp")
		end if 	
	else
		session("prospecterror")=session("giftcustID")
		set sitelink=nothing 		
		response.redirect("proserror.asp")
	end if
 %>

