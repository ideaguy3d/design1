
<!--#INCLUDE FILE = "commonvar.asp" -->
<% 
'This page does not include any site related preferences from text folder so do not include this file in any page that uses site preferences.

    if session("storeclosed")="1" then		
		response.redirect(insecureurl+"closed.asp")
	end if
	
	
	set sitelink=createobject("sitel700.sitelink")	
	call sitelink.setdata(datapath,cstr(""),ShortStoreName)

	'if session("referral")="0" then
	'	session("referral")=request.servervariables("HTTP_REFERER")
	'end if
	
	
	'response.Write("ref site-->"&session("reff"))
		
	function fix_xss_Chars(thisvar)
		'for scan alert
		thisvar = replace(thisvar,"<","&lt;")
		thisvar = replace(thisvar,">","&gt;")
		thisvar = replace(thisvar,"'","")
		thisvar = replace(thisvar,"(","")
		thisvar = replace(thisvar,")","")
		thisvar = replace(thisvar,"""","")		
		fix_xss_Chars = thisvar
	end function
	
	'function url_cleanse(thatvar)
		'remove spaces and slashes		
	'	thatvar = replace(thatvar," ","-")
	'	thatvar = replace(thatvar,"/","_",1,10,1)		
	'	thatvar = replace(thatvar,">","")
	'	thatvar = replace(thatvar,"<","")
	'	thatvar = replace(thatvar,"%","")
	'	thatvar = replace(thatvar,"""","")
	'	url_cleanse = thatvar	
	'end function

	
%>
