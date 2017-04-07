<%on error resume next%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--INCLUDE FILE = "CreateXmlObject.asp" -->


<% 

	'orderconfirm = sitelink.ORDER_CONFIRMED(session("ordernumber"))
	'if orderconfirm = true then
	'  	set sitelink=nothing
	 ' 	set ObjDoc = nothing
	'	response.redirect("receipt.asp")
	'end if
	
	sl_discount = trim(cstr(REQUEST.FORM("discount")))
	sl_discount=fix_xss_Chars(sl_discount)
	
	
  if len(sl_discount) = 0 then
	   session("sourcekeyisgood") = -1 
	   session("askedforsourcekey")=1
	   set sitelink = nothing
	   set objDoc = nothing
	   response.redirect("basket.asp")
  end if
    
  
  
   '  check whether the source key is for item level or prder promo
  	 session("sourcekeyisgood")=sitelink.verifysourcekey(session("shopperid"),cstr(sl_discount),session("ordernumber"))
  	 
  	 if session("sourcekeyisgood") <> -1 then
  		session("askedforsourcekey")=1
	 else
		session("sourcekeyisgood") = -1 
		session("askedforsourcekey")=1
	 end if 	
	 
		Set objDoc =nothing
		set sitelink=nothing
		response.redirect("basket.asp")
  	 
  	   		

 ' Now check whether the source is for global discounts	
'	xmlstring =  sitelink.globaldiscounts("READ", cstr(REQUEST.FORM("discount")), cint(0),cint(0), cint(0))	
'	objDoc.loadxml(xmlstring)
	
	
'	vsourcekey = trim(ucase(REQUEST.FORM("discount")))
	'targetnode = "//discount/sourcekey[.='"+vsourcekey+"']"	
	'set SL_Discount = objDoc.selectNodes(targetnode)

'	set SL_Discount = objDoc.selectNodes("//gdc")	

	
'	if SL_Discount.length > 0 then 'match found in discount database
'	   'check the dates.
'	   venddate = trim(SL_Discount.item(0).selectSingleNode("enddate").text)
	   
'	   if len(trim(venddate)) = 0 then
'			session("askedforsourcekey")=1
'			session("sourcekeyisgood") = 1	
'			session("CurrentSourceKey") = vsourcekey			   
'	   else
'	   	   	sourceKeyEnddate = formatdatetime(venddate,2)
'			SL_system_date = formatdatetime(date(),2)
'			diff_days = cdate(sourceKeyEnddate)-cdate(SL_system_date)
'	   		if (diff_days < 0 ) then
'	   			'coupon code has expired.			
'				session("sourcekeyisgood") = -1 
'				session("askedforsourcekey")=1
'			else
'				session("askedforsourcekey")=1
'				session("sourcekeyisgood") = 1	
'				session("CurrentSourceKey") = vsourcekey			
'	   		end if
'		end if
'	   
'		set SL_Discount = nothing
'		Set objDoc =nothing
'		set sitelink=nothing
'
'		response.redirect("basket.asp")
'	end if 
'	
'	set SL_Discount = nothing
'
 ' '  check whether the source key is for item level
  '	 session("sourcekeyisgood")=sitelink.verifysourcekey(session("shopperid"),cstr(trim(REQUEST.FORM("discount"))),session("ordernumber"))
  '	 
  '	 if session("sourcekeyisgood") <> -1 then
  '		session("askedforsourcekey")=1	
  '		session("CurrentSourceKey") = cstr(trim(REQUEST.FORM("discount")))
'		'session("CurrentSourceKey") = ""  'verify sourcekey will write to the receipt.. Do not alter this.		
'		
'		set SL_Discount = nothing
'		Set objDoc =nothing
'		set sitelink=nothing
'		response.redirect("basket.asp")
 ' 	 end if 
  	   		
		
	
	' If non of the above is full filled then sourcekey is invalid	
	'Response.Write(session("sourcekeyisgood"))
'	session("sourcekeyisgood") = -1 
'	session("askedforsourcekey")=1
	
'	set SL_Discount = nothing
'	Set objDoc =nothing
'	set sitelink=nothing
'
'	response.redirect("basket.asp")

%>

