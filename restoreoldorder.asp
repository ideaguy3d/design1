<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<%
	'clear the basket.
	'CALL sitelink.EMPTYBASKET(SESSION("SHOPPERID"),SESSION("ORDERNUMBER")) 		

		extrafield =""
		'extrafield =""
		xmlstring = sitelink.getbasketinfo(session("shopperid"),session("restoreoldordernumber"),extrafield,false,false)
		
		objDoc.loadxml(xmlstring)
		set SL_Basket = objDoc.selectNodes("//gbi[basketdesc!='PROMO' and certid=0]")
		for x=0 to SL_Basket.length-1
			SL_Number			= SL_Basket.item(x).selectSingleNode("number2").text
			SL_Variant			= SL_Basket.item(x).selectSingleNode("variant").text
			SL_BasketQauntity	= SL_Basket.item(x).selectSingleNode("quanto").text
			SL_BasketCustInfo	= SL_Basket.item(x).selectSingleNode("custominfo").text
			if SL_Number<>"" then
				CALL sitelink.ITEMADD(cstr(SL_Number),cstr(SL_BasketQauntity),cstr(SL_Variant),session("shopperid"),session("ordernumber"),cstr(SL_BasketCustInfo),cint(0),cint(0))	
			end if
			
		next
		
		set sitelink=nothing
		set ObjDoc=nothing
		
		session("restoreoldorder")=0
		'session("restoreoldordernumber")=""
		response.redirect("basket.asp")
		

%>