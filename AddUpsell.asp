<%on error resume next%>


<%
current_stocknumber =  REQUEST.FORM("OptionalSku")
Selltype = request.form("TypeofSell")

baccept = request.form("UserAccept")
breject = request.form("UserReject")

sellmand = request.form("sellmand")
 
 if baccept="Accept" and breject="" then
	 if len(trim(current_stocknumber)) = 0 then
		response.redirect("default.asp")
	 end if
 end if
%>
<!--#INCLUDE FILE = "include/momapp.asp" -->


<%

  if Selltype="U" then
  	if baccept="Accept" then
		AllNonSkuAddPrice="0"
		custominfofield=session("ParentskuCustomInfo")
  		CALL sitelink.ITEMADD(cstr(current_stocknumber),cstr(session("ParentSkuQty")),cstr(""),session("shopperid"),session("ordernumber"),cstr(custominfofield),cint(0),cint(0),cstr(""),cstr(AllNonSkuAddPrice))	  
  	end if
	
	if breject="Reject" then
  		CALL sitelink.ITEMADD(cstr(session("Parentsku")),cstr(session("ParentSkuQty")),cstr(session("ParentskuVariant")),session("shopperid"),session("ordernumber"),cstr(session("ParentskuCustomInfo")),cint(0),cint(0),cstr(""),cstr(session("ParentskuNonAttribPrice")))	
  	end if	
  end if
  
  if Selltype="S" then
  	if baccept="Accept" then
		AllNonSkuAddPrice="0"
		custominfofield=session("ParentskuCustomInfo")
  		CALL sitelink.ITEMADD(cstr(current_stocknumber),cstr(session("ParentSkuQty")),cstr(""),session("shopperid"),session("ordernumber"),cstr(custominfofield),cint(0),cint(0),cstr(""),cstr(AllNonSkuAddPrice))	  
  	end if
	
	if sellmand="false" and breject="Reject"  then		
  		CALL sitelink.ITEMADD(cstr(session("Parentsku")),cstr(session("ParentSkuQty")),cstr(session("ParentskuVariant")),session("shopperid"),session("ordernumber"),cstr(session("ParentskuCustomInfo")),cint(0),cint(0),cstr(""),cstr(session("ParentskuNonAttribPrice")))	
	end if
  
  end if 
  
  

set sitelink=nothing

'response.write(session("ParentSkuQty"))

response.redirect("basket.asp")
		
	
	
%>

