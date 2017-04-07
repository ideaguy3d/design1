<%on error resume next%>
<%

'set the advanced field

 search_on = Request.QueryString("search")
 searchstring = Request.QueryString("what")
 
 if search_on =1 then
	session("search_on")="STOCK.advanced1"
	
	usethisurl = "searchprods.asp?searchstring=" & searchstring &"&pagenumber=1"
	Response.redirect(usethisurl) 
 end if

 if search_on =2 then
	session("search_on")="STOCK.advanced2"
	
	usethisurl = "searchprods.asp?searchstring=" & searchstring &"&pagenumber=1"
	Response.redirect(usethisurl) 
 end if

 if search_on =3 then
	session("search_on")="STOCK.advanced3"
	
	usethisurl = "searchprods.asp?searchstring=" & searchstring &"&pagenumber=1"
	Response.redirect(usethisurl) 
 end if

 if search_on =4 then
	session("search_on")="STOCK.advanced4"
	
	usethisurl = "searchprods.asp?searchstring=" & searchstring &"&pagenumber=1"
	Response.redirect(usethisurl) 
 end if


%>