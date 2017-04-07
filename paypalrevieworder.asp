
<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<%
    TokenID = cstr(session("TokenID"))
    
    if len(trim(TokenID)) = 0 then
        set sitelink = nothing
        ErrorMsg = "Invalid Token"
		session("custerrormessage") = 	ErrorMsg +"<br>"
		response.Redirect("PayPalError.asp")		
    end if

    'TokenID = cstr(request.QueryString("TokenID"))
    'PayerID = cstr(request.QueryString("PayerID"))
    
    set SLCustomer = New cCustomer
    set PayPalExpressObj = New cPayPalExpressObj
        PayPalExpressObj.Action= "VALIDATE_PAYPALREVIEWORDER"
        PayPalExpressObj.OrderNum = 0
        PayPalExpressObj.TokenID = TokenID
        
     call sitelink.Getpostdata_PayPalExpress(PayPalExpressObj,SLCustomer)
        OrderNumber =PayPalExpressObj.OrderNum
         ErrorMsg = PayPalExpressObj.ErrorMsg
        HasError = PayPalExpressObj.HasError
        
    set PayPalExpressObj = nothing
    set SLCustomer = nothing
    
    'response.Write(OrderNumber)
    'newordernumber = session("ordernumber") 
    'response.Write("<br>")
    'response.Write("express order -->" & OrderNumber)
    'response.Write("<br>")
    'response.Write("current session order -->" & newordernumber)
    
    if OrderNumber <> session("ordernumber") then
        'regenarate session
        session("ordernumber") = OrderNumber
        
        billtocust = sitelink.GetBilltoCust(session("ordernumber"))
	        set shopperRecord= sitelink.SHOPPERINFO(cstr(billtocust),false)
	        session("shopperid") = shopperRecord.custnum
	        session("firstname")= trim(shopperRecord.FirstName)
	        session("lastname")= trim(shopperRecord.Lastname)
	        session("address1")= trim(shopperRecord.addr)
	        session("address2")= trim(shopperRecord.addr2)
	        session("address3")= trim(shopperRecord.addr3)
	        session("company")= trim(shopperRecord.company)
	        session("city")= trim(shopperRecord.city)
	        session("state")= trim(shopperRecord.state)
	        session("zipcode")= trim(shopperRecord.zipcode)
	        session("email")= trim(shopperRecord.email)
	        session("country")= trim(shopperRecord.country)
	        session("bcounty")= trim(shopperRecord.county)
	        set shopperRecord= nothing
	        
	        set shopperRecord= sitelink.SHIPTONAME(cstr(session("ordernumber")),false)
				session("sfirstname")= trim(shopperRecord.FirstName)
				session("slastname")= trim(shopperRecord.Lastname)
				session("saddress1")= trim(shopperRecord.addr)
				session("saddress2")= trim(shopperRecord.addr2)
				session("saddress3")= trim(shopperRecord.addr3)
				session("scompany")= trim(shopperRecord.company)
				session("scity")= trim(shopperRecord.city)
				session("sstate")= trim(shopperRecord.state)
				session("szipcode")= trim(shopperRecord.zipcode)
				session("semail")= trim(shopperRecord.email)
				session("scountry")= trim(shopperRecord.country)	
				session("scounty")= trim(shopperRecord.county)							 
			 	set shopperRecord= nothing
			 	
	        session("billtocopy") = 0
	        set ORDER_RECORD = sitelink.ORDER_RECORD(session("ordernumber"),false)	 			 
			 session("shippingmethod") = ORDER_RECORD.shiplist
			 set ORDER_RECORD = nothing
			 
			 session("SL_BasketCount") = sitelink.GetData_Basket(cstr("BASKETCOUNT"),session("ordernumber"),session("SL_VatInc"))
	         session("SL_BasketSubTotal") = sitelink.GetData_Basket(cstr("BASKETSUBTOTAL"),session("ordernumber"),session("SL_VatInc"))
	
        'response.Write("<br>")
        'response.Write(session("ordernumber"))
    end if
    
    
    'response.Write("<br>")
    'response.Write("reverted session order -->" & session("ordernumber"))
    'response.Write("<br>")
    'response.Write("reverted session shopper -->" & session("shopperid"))
    
    
    'response.Write(HasError)
    
    if HasError=true then		
		    set sitelink = nothing
			session("custerrormessage") = 	ErrorMsg +"<br>"
			response.Redirect("PayPalError.asp")			 
	end if	
	
	DoNot_Pass_CartContents = false
    has_backOrder_item = false

		
	        
		nodes_remaining=0
		basket_has_multiship = false		
		'nodes_remaining=sitelink.CHECK_FOR_ALL_SHIPVIA(session("ordernumber"))
		
		set SL_BasketShipViaShiptoInfo = New cBasketShipViaShiptoObj

		SL_BasketShipViaShiptoInfo.lordernumber=session("ordernumber")		
		call sitelink.CHECK_FOR_ALL_SHIPVIA(SL_BasketShipViaShiptoInfo)
		    nodes_remaining=SL_BasketShipViaShiptoInfo.lshipvianodes
		    if SL_BasketShipViaShiptoInfo.lshiptonodes > 0 then
		        basket_has_multiship=true
		    end if
		set SL_BasketShipViaShiptoInfo = nothing
		
   
    'must pass basket_has_multiship correctly
	    retval = sitelink.TOTALORDER(session("shopperid"),session("ordernumber"),session("billtocopy"),session("shippingmethod"),false)
	    if (retval=false) then
		    session("ordererrormessage")  = "There was an error completing your order at this time. Please try again"
		    set sitelink=nothing 
		    set ObjDoc = nothing
		    response.redirect("ordererror.asp")	
		end if
		
			'response.Write(retval)
            if session("billtocopy")=1 then
				 vtxtzipCode = session("zipcode")
				 vtextshipcountry = session("country")
				 vtxtstate = session("state")
			else
				 vtxtzipCode = session("szipcode")
				 vtextshipcountry = session("scountry")
				 vtxtstate = session("sstate")
			end if	
			               
        if nodes_remaining > 0 then
            show_shipping_option=true  
        else
            show_shipping_option=false
        end if
        
        
		if nodes_remaining > 0 then
			show_shipping_option=true
			 shipping_order="3"				
				shipping_order = cstr(SHIPPING_SORTBY) 
				if SHIPPING_SORTORDER=1 then
					shipping_order = shipping_order +" desc"
				end if 
											
				xmlstring=sitelink.previewshipping(session("shopperid"),session("ordernumber"),cstr(vtextshipcountry),UCase(cstr(vtxtstate)),cstr(vtxtzipCode),shipping_order,session("shippingmethod"))				
				objDoc.loadxml(xmlstring)	
				set SL_Ship = objDoc.selectNodes("//preship" + CUSTOM_SHIPPING_FILTER)				
		else
			show_shipping_option=false			
		end if 
		        

		extrafield =""
		xmlstring = sitelink.getbasketinfo(session("shopperid"),session("ordernumber"),extrafield,false,session("user_want_multiship"))
	
		objDoc.loadxml(xmlstring)
		
		'set SL_Basket = objDoc.selectNodes("//gbi")
				
		'basket_has_multiship = false
		'filter_condition = "//gbi[ship_to>0]"
		'set SL_has_multishipto_basket = objDoc.selectNodes(filter_condition)
		'	if SL_has_multishipto_basket.length > 0 then
		'		basket_has_multiship= true
		'	end if		
		'set SL_has_multishipto_basket = nothing
		
		if ALLOW_SHIPAHEAD=1 then
		    filter_condition ="//gbi[number2!='' and units<= 0 and dropship='false' and  construct='false']"			
			    set SL_Basket_backordered_item = objDoc.selectNodes(filter_condition)
				    if SL_Basket_backordered_item.length >0 then
					    has_backOrder_item = true
				    end if 
		    set SL_Basket_backordered_item = nothing
		end if
			
	  if session("hasProdRest") = true then
	    has_backOrder_item = true
	  end if


	basket_has_giftcert = false
	basket_has_NegativePrice = false
	giftcertitem=""
	giftcertAmount=0
	filter_condition= "//gbi[it_unlist < 0]"
	set SL_Basket_NegativePrice = objDoc.selectNodes(filter_condition)
	if SL_Basket_NegativePrice.length > 0 then
		basket_has_NegativePrice=true
	end if 
	set SL_Basket_NegativePrice=nothing
	
	 
    filter_condition = "//gbi[certid>0]"
    	set SL_Basket_giftcert = objDoc.selectNodes(filter_condition)
    		if SL_Basket_giftcert.length > 0 then
				basket_has_giftcert = true
				giftcertitem = SL_Basket_giftcert.item(0).selectSingleNode("item").text
				giftcertAmount = SL_Basket_giftcert.item(0).selectSingleNode("it_unlist").text
				giftcertAmount = abs(giftcertAmount)
				DoNot_Pass_CartContents = true
			end if 
		set SL_Basket_giftcert = nothing
		
	    'set SL_Basket = objDoc.selectNodes("//gbi")
        'if basket_has_giftcert=true then
		    'filter_condition = "//gbi[item !='"+cstr(giftcertitem)+"']"			
	    'else
	    '    filter_condition = "//gbi"
	    'end if
	
		filter_condition = "//gbi[certid=0]"					    
		'response.write(filter_condition)
	    set SL_Basket = objDoc.selectNodes(filter_condition)
		
		
		set SL_OrderRecInfo = New cOrderRecObj

		SL_OrderRecInfo.lordernumber=session("ordernumber")
		SL_OrderRecInfo.lshopperid=session("shopperid")
		call sitelink.Ord_Details(SL_OrderRecInfo)
			SL_ORD_TOTAL		= SL_OrderRecInfo.lordertotal
			SL_PrintOrderAmount = SL_OrderRecInfo.lordertotal
			SL_ORD_DATE 		= SL_OrderRecInfo.lodrdate
			default_ship_title	= SL_OrderRecInfo.lshipmethod
			SL_PrintTaxAmount 	= SL_OrderRecInfo.ltax
			SL_PrintShipAmount 	= SL_OrderRecInfo.lshipping
			points_amt 			= SL_OrderRecInfo.lcheckamoun
			points_used 		= SL_OrderRecInfo.lpoints_usd
			
			GiftCardAmount 		= SL_OrderRecInfo.lGCardAmount
			
			'used on basket to restrict from using Google Checkout
			session("OrderLevelPoints_Amt") = points_amt
			
			SL_ALLOWCOD				= SL_OrderRecInfo.lstartslc
			SL_ALLOWTERMS			= SL_OrderRecInfo.lstartsli
			SL_ALLOWPOINTS 			= SL_OrderRecInfo.lstartprk			
			SL_POINT_REDEEEM_VAL	= SL_OrderRecInfo.lstartrtp
			SL_POINT_REDEEEM_VAL_AMT= SL_OrderRecInfo.lstartram
			SL_ALLGIFTCARD          = SL_OrderRecInfo.lstarttid
			
			SL_SourcekeyXmlstream   = SL_OrderRecInfo.lsourcekeylist
			terms_avail				= SL_OrderRecInfo.lshopperHasTerms
			PayPal_LoginID			= SL_OrderRecInfo.lPaypalLoginID
			SL_CardListXmlstream    = SL_OrderRecInfo.lCardList
			
		set SL_OrderRecInfo=nothing
		
        
		
		
		'SL_ORD_TOTAL=0
	 	'if SL_ALLOWPOINTS=true then
		'set ORDER_RECORD = sitelink.ORDER_RECORD(session("ordernumber"),false)
		'  points_used = ORDER_RECORD.points_usd
		'  points_amt = ORDER_RECORD.checkamoun
		'  SL_ORD_TOTAL  = ORDER_RECORD.ord_total
		'set ORDER_RECORD=nothing
		'end if
		
		if points_amt > 0 and points_used > 0 then			
			SL_ORD_TOTAL = cdbl(SL_ORD_TOTAL - points_amt)
		end if
		
		'giftcertAmount is absolute
		SL_ORD_TOTAL = SL_ORD_TOTAL - giftcertAmount
		SL_ORD_TOTAL = round(SL_ORD_TOTAL,2)
		
		if GiftCardAmount > 0 then
		    SL_ORD_TOTAL = SL_ORD_TOTAL - GiftCardAmount		
		end if
		
		'don't want to trigger persist cart
		session("returnvisitor")=true 
%>
<html>
<head>
<title><%=althomepage%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
<link rel="stylesheet" href="text/footernav.css" type="text/css">
<link rel="stylesheet" href="text/design.css" type="text/css">
<script type="text/javascript" language="JavaScript" src="checkout.js"></script>
</head>

<!--#INCLUDE FILE = "text/Background5.asp" -->


<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">



<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top"  >
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	<td valign="top" class="pagenavbg">        
        <!-- sl code goes here -->
        <form name="checkoutform" method="post" action="<%=secureurl%>GetPayPalPayment.asp">
        <div id="page-content" class="static">
		<table width="100%"  border="0" cellspacing="0" cellpadding="3">
			<tr>
             <td align="left"><h2>Your purchase will cost 			  
			  <%if SL_ORD_TOTAL > 0  then %>
			  	<%=FORMATCURRENCY(SL_ORD_TOTAL)%>				
			  <%else%>
				  <%=FORMATCURRENCY(0)%>			 
			  <%end if%>.</h2>
             </td>
	         </tr>			
	    </table>
	    
	    <!--#INCLUDE FILE = "print-no-multiship.asp" -->
	    
	     <%
	 
		'SL_PrintOrderAmount is populated in print-no-multiship.asp/print-multiship.asp
		SL_PrintGrandTotal=SL_PrintOrderAmount
	 	if SL_ALLOWPOINTS=true then
			if points_amt > 0 and points_used > 0 then
				SL_PrintGrandTotal = cdbl(SL_PrintGrandTotal - points_amt)
				SL_PrintGrandTotal = round(SL_PrintGrandTotal,2)
				DoNot_Pass_CartContents = true
	 %>	
		 <table width="98%">
		<tr>
			<td align="right" class="plaintextbold">POINTS REDEEMED:</td>
			<td align="right" width="11%" class="plaintextbold"><%=formatcurrency(points_amt)%></td>
		
		</tr>
		
				  
		</table>
	<% 
	end if	
	end if%>
	
	<%
	'session("Remaining_BalanceAfterGC")=0
	if basket_has_giftcert=true then 
		SL_PrintGrandTotal = cdbl(SL_PrintGrandTotal-giftcertAmount)
	    SL_PrintGrandTotal = round(SL_PrintGrandTotal,2)

	    'session("Remaining_BalanceAfterGC") = SL_PrintGrandTotal
	%>
        <table width="98%">
		<tr>
			<td align="right" class="plaintextbold">GIFT CERTIFICATE REDEEMED:</td>
			<td align="right" width="11%" class="plaintextbold"><%=formatcurrency(giftcertAmount)%></td>
		
		</tr>	
		</table>		
        <%end if%>
        
        <% if GiftCardAmount > 0 then 
            SL_PrintGrandTotal = cdbl(SL_PrintGrandTotal-GiftCardAmount)
	        SL_PrintGrandTotal = round(SL_PrintGrandTotal,2)
        %>
            
         <table width="98%">
		    <tr>
			    <td align="right" class="plaintextbold">GIFT CARD AMOUNT:</td>
			    <td align="right" width="11%" class="plaintextbold"><%=formatcurrency(GiftCardAmount)%></td>		
		    </tr>	
		</table>
        <%end if%>
        
        <%if (points_amt > 0 and points_used > 0) or basket_has_giftcert=true or GiftCardAmount > 0 then %>
	    <table width="98%">
	    <tr><td colspan="2"><hr size="1"></td></tr>
		<tr>
			<td align="right" class="plaintextbold">GRAND TOTAL:</td>
			<td align="right" width="11%" class="plaintextbold"><%=formatcurrency(SL_PrintGrandTotal)%></td>
		
		</tr>	
		</table>
		<%end if%>
	 <br>
	    
	    
	    <center>	    
	    
	    <table>
	        <tr>
	            <td><input type="image" src="images/btn-completeorder.gif" border="0" /></td>
	            <td><input type="hidden" name="TokenID" id="TokenID" value="<%=TokenID%>" /></td>
	        </tr>
	    
	    </table>
	    </center>
	    
		</div> 
		</form>
        <!-- end sl_code here -->	</td>
</tr>

</table>

    </div> <!-- Closes main  -->
    <div id="footer" class="footerbgcolor">
    <!--#INCLUDE FILE = "include/bottomlinks.asp" -->
    <!--#INCLUDE FILE = "googletracking.asp" -->
    <!--#INCLUDE FILE = "RemoveXmlObject.asp" -->
    <!--#INCLUDE FILE = "text/footer.asp" -->
    </div>
</div> <!-- Closes container  -->




</body>
</html>
