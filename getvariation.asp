

<%on error resume next%>

<!--#INCLUDE FILE = "text/adminSeoPref.asp" -->
<!--#INCLUDE FILE = "include/url_cleanse.asp" -->

<%

 numberpart = cstr(trim(request.querystring("number")))
 variationpart =cstr(trim(request.querystring("whatvar")))
 addthisurl =cstr(trim(request.querystring("addthis")))
 
 session("Current_selectedVariant") = variationpart
 
 if WANT_REWRITE = 1 then
 	addthisurl = url_cleanse(addthisurl)
 	 if request.ServerVariables("SERVER_PORT_SECURE") = 1 then
	 	usethisurl=secureurl+ cstr(addthisurl) + "/productinfo/"+ cstr(numberpart)+"/"+ cstr(variationpart)+ "/"
	 else
	 	usethisurl=insecureurl+ cstr(addthisurl) + "/productinfo/"+ cstr(numberpart)+ "/"+ cstr(variationpart)+ "/"
	 end if
 else
	 usethisurl = "prodinfo.asp?number="&numberpart&"&variation="&variationpart
 end if
 
' response.write(usethisurl)
 
 response.redirect(usethisurl)


%>