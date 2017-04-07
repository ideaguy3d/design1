<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->


<% 
		
	
	txtebasketitem = REQUEST.FORM("txtebasketitem")
	
	if isnumeric(txtebasketitem)=true then
		if txtebasketitem> 0 then
			call sitelink.itemmodify(cstr(request.form("EDITITEM")),cstr(REQUEST.FORM("txtequanto")),cstr(txtebasketitem),cstr(REQUEST.FORM("TXTeVARIANT")),session("shopperid"),cstr(REQUEST.FORM("txtcustominfo")),session("ordernumber")) 
		end if
	end if
	
	
	set sitelink=nothing 
	response.redirect("basket.asp")
	
	%>





