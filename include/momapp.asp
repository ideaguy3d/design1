

<!--#INCLUDE FILE = "commonvar.asp" -->

<!--#INCLUDE FILE = "browsercheck.asp" -->

<% 

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


<!--#INCLUDE FILE = "cLibrary.asp" -->

<!--#INCLUDE FILE = "../text/adminPayPref.asp" -->
<!--#INCLUDE FILE = "../text/admincartpref.asp" -->
<!--#INCLUDE FILE = "../text/adminPagePref.asp" -->
<!--#INCLUDE FILE = "../text/adminCustMagmtpref.asp" -->
<!--#INCLUDE FILE = "../text/adminProdDispPref.asp" -->
<!--#INCLUDE FILE = "../text/adminProdPref.asp" -->
<!--#INCLUDE FILE = "../text/adminDeptPref.asp" -->
<!--#INCLUDE FILE = "../text/adminSeoPref.asp" -->
<!--#INCLUDE FILE = "../text/LivePersonPref.asp" -->
<!--#INCLUDE FILE = "../text/analyticspref.asp" -->
<!--#INCLUDE FILE = "../text/storetext.asp" -->
<!--#INCLUDE FILE = "../text/BuySafeAct.asp" -->
<!--#INCLUDE FILE = "../text/adminFeadProdPref.asp" -->
<!--#INCLUDE FILE = "../text/adminbuySafePref.asp" -->
<!--#INCLUDE FILE = "url_cleanse.asp" -->


