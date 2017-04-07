<%on error resume next%>


<!--#INCLUDE FILE = "include/momapp.asp" -->
<%	

	
	session("giftregemail") = cstr(trim(request("regtxtemail")))
	session("giftregpassword")=cstr(trim(request("txtregpassword")))
	
	session("giftregemail") = fix_xss_Chars(session("giftregemail"))
	session("giftregpassword")= fix_xss_Chars(session("giftregpassword"))
	
	session("savegiftregid")=sitelink.verifygiftreg(session("giftregemail"),session("giftregpassword"))
	session("savegiftregid")=UCase(Trim(session("savegiftregid")))

'if one record found for entered login then redirect customer to giftlist page
if session("savegiftregid")<>"-1" and session("savegiftregid")<>"0" and cint(session("savegiftregid")) > 1 then

	session("giftregverified")="YES"
	session("giftregcustnum")= session("savegiftregid")
	session("giftregcustnum")= session("giftregcustnum")+1-1
	set sitelink= nothing
	
	if len(trim(session("destpage"))) = 0 then
		response.redirect("giftlist.asp")
	else
		response.redirect(session("destpage"))
	end if

'if more than one record found for entered login then redirect customer to select gift registry page
'where they can select proper registry
elseif session("savegiftregid")="0" then
	session("giftregverified")="YES"
	set sitelink= nothing

	if len(trim(session("destpage"))) = 0 then
		response.redirect("viewRegistries.asp")
	else
		response.redirect(session("destpage"))
	end if
else
	'if no records found go back to login page & show error
	session("giftregverified")="NO"
	set sitelink= nothing	
	response.redirect("giftreglogin.asp")
end if

%>
