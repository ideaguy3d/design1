<%on error resume next%>


<%

	   if session("registeredshopper")="NO" then
  			response.redirect("statuslogin.asp")
	    end if 
	  
		whatbastnum = REQUEST.FORM("basketrecord")	
		intlastshipto = cstr(whatbastnum)
		
		lastadress_id = REQUEST.FORM("shopcust")		
		lastadress_id = lastadress_id + 1 -1
		
		if isnumeric(intlastshipto)=false or isnumeric(lastadress_id)=false then	
			Response.Status="301 Moved Permanantly"
			Response.AddHeader "Location", insecureurl&"default.asp"
			Response.End
		end if 

		giftmsg1=cstr(REQUEST.FORM("txtgiftmsg1"))
		giftmsg2=cstr(REQUEST.FORM("txtgiftmsg2"))
		giftmsg3=cstr(REQUEST.FORM("txtgiftmsg3"))
		giftmsg4=cstr(REQUEST.FORM("txtgiftmsg4"))
		giftmsg5=cstr(REQUEST.FORM("txtgiftmsg5"))
		giftmsg6=cstr(REQUEST.FORM("txtgiftmsg6"))
%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<%		
		aa=sitelink.CHANGEBASKETSHIPTO(session("ordernumber"),intlastshipto,session("shopperid"),lastadress_id,giftmsg1,giftmsg2,giftmsg3,giftmsg4,giftmsg5,giftmsg6)
		
		'response.write(aa)
		set sitelink=nothing
		response.redirect("basket.asp")


%>