<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">


<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 
<%
if session("SL_BasketCount")= 0 then
	response.redirect("basket.asp")
end if

if session("CheckedForProdRest") = false then
    response.redirect("custinfo.asp")
end if

%>

<!--#INCLUDE FILE = "include/CommonVar.asp" -->

<%
 if REQUEST.servervariables("server_port_secure")<> 1 and SLsitelive=true then 
 	response.redirect(secureurl+"checkout.asp") 
  end if
%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<% 
	'orderconfirm = sitelink.ORDER_CONFIRMED(session("ordernumber"))
	'if orderconfirm = true then
	 ' 	set sitelink=nothing	
	 ' 	set objDoc=nothing  	
	'	response.redirect("receipt.asp")
	'end if
		
%>


<%


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
	    call sitelink.TOTALORDER(session("shopperid"),session("ordernumber"),session("billtocopy"),session("shippingmethod"),basket_has_multiship)
		
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
        
        
'		if nodes_remaining > 0 then
'			show_shipping_option=true
'			 shipping_order="3"				
'				shipping_order = cstr(SHIPPING_SORTBY) 
'				if SHIPPING_SORTORDER=1 then
'					shipping_order = shipping_order +" desc"
'				end if 
'				
'				if session("billtocopy")=1 then
'					 vtxtzipCode = session("zipcode")
'					 vtextshipcountry = session("country")
'					 vtxtstate = session("state")
'				else
'					 vtxtzipCode = session("szipcode")
'					 vtextshipcountry = session("scountry")
'					 vtxtstate = session("sstate")
'				end if				
'				xmlstring=sitelink.previewshipping(session("shopperid"),session("ordernumber"),cstr(vtextshipcountry),UCase(cstr(vtxtstate)),cstr(vtxtzipCode),shipping_order,session("shippingmethod"))				
'				objDoc.loadxml(xmlstring)	
'				set SL_Ship = objDoc.selectNodes("//preship" + CUSTOM_SHIPPING_FILTER)				
'		else
'			show_shipping_option=false			
'		end if 
		        

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
	
%>


<html>
<head>
<title>Aussie Products.com | Checkout Process | Australian Products Co.</title>
<meta name="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls" />
	<meta name="keywords" content="Aussie Products, Australian, Aussie, Australian Products, Australian Food, Aussie Foods, Australian Desserts, tim tams, vegemite, violet crumble, cherry ripe, koala, kangaroo, down under, Australian Groceries, foods, lollies, candy, gourmet food" />
	<meta name="robots" content="index, follow" />
    <meta name="robots" content="ALL">
	<meta name="distribution" content="Global">
	<meta name="revisit-after" content="5">
	<meta name="classification" content="Food">	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
<link rel="stylesheet" href="text/footernav.css" type="text/css">
<link rel="stylesheet" href="text/design.css" type="text/css">
<script type="text/javascript" language="JavaScript" src="calender.js"></script>
<script type="text/javascript" language="JavaScript" src="checkout.js"></script>
</head>

<% 


	defaultpan_id = "tab1"
	
	if session("paymentmethod")="Credit Card" then
			defaultpan_id = "tab1"
	end if
	
	if session("paymentmethod")="EC" then
			defaultpan_id = "tab2"
	end if

	if session("paymentmethod")="COD" then
			defaultpan_id = "tab3"
	end if
	if session("paymentmethod")="IN" then
			defaultpan_id = "tab4"
	end if

	if session("paymentmethod")="PP" then
			defaultpan_id = "tab5"
	end if

	if session("paymentmethod")="CK" then
			defaultpan_id = "tab6"
	end if
	if session("paymentmethod")="CK" and session("PaybyPoints")=1 then
			defaultpan_id = "tab6"
	end if
	if session("paymentmethod")="CK" and session("PaybyGC")=1 then
			defaultpan_id = "tab7"
	end if
	
	if session("paymentmethod")="GiftCard" then
			defaultpan_id = "tab8"
	end if
	
	

%>
<%if SL_ORD_TOTAL > 0 then%>
<!--#INCLUDE FILE = "text/Backgroundcheckout.asp" -->
<%else%>
<!--#INCLUDE FILE = "text/Background5.asp" -->
<%end if%>


<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">

<%'response.write(shipfilter)%>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
	<%if WANT_SIDENAV=1 then%>
	<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
	<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>		
	<%end if%>
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	
	<table width="97%"  cellpadding="0" cellspacing="0" border="0">
		<tr>
            <td align="left" width="97%"><img src="images/step2.gif" border="0" width="403" height="29"></td>
          </tr>
	</table>
	<br>
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
	<br>
	
	
	<%if SL_ORD_TOTAL > 0 then	%>
	<form name="checkoutform"  action="<%=secureurl%>paymentInfo.asp" onSubmit="return Process()" method="POST">
	<%else%>
	<form name="checkoutform"  action="<%=secureurl%>NopaymentNeeded.asp"  method="POST">
	<%end if%>
			
	<% 

		if basket_has_multiship= false then
	 %>
		<!--#INCLUDE FILE = "print-no-multiship.asp" -->
	<%else%>	 	
	 	<!--#INCLUDE FILE = "print-multiship.asp" -->	 
	 <%end if%>
	 <% 'set SL_Basket = nothing
	 	set SL_Ship =nothing
	 %>
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
	 
	 <%
	 'response.write(SL_ORD_TOTAL)
	 if SL_ORD_TOTAL > 0 then	 
	 objDoc.loadxml(SL_SourcekeyXmlstream)
	 set SL_SourceKeyList = objDoc.selectNodes("//skeylist")
	 if SL_SourceKeyList.length > 0 then
	 %>
	 <table width="100%" cellpadding="0" cellspacing="0" border="0">
	 	<tr><td align="left" class="tab-style-header" width="30%">How did you hear about us ?</td><td valign="middle" align="left">
			<table width="0" cellpadding="0" cellspacing="0" border="0"> 
						<tr><td class="plaintextbold" align="center">&nbsp;&nbsp;Source:&nbsp;</td>
						<td valign="top" class="nopadding">
							<select name="whatsourcekey" id="whatsourcekey" class="plaintext">
								<option value="">
								<% for x=0 to SL_SourceKeyList.length-1
									SL_KeyValue = SL_SourceKeyList.item(x).selectSingleNode("adkey").text
									SL_KeyDesc = SL_SourceKeyList.item(x).selectSingleNode("adin").text
								%>
									<option value="<%=SL_KeyValue%>"><%=SL_KeyDesc%>
								<%next%>			
							</select>
						</td>
					</tr>
					</table>
			</td></tr>
			 
	 </table>
		<%end if
		set SL_SourceKeyList=nothing
		%>

		<br>
		<table width="100%" cellpadding="0" cellspacing="0" border="0">
        	<tr>
        		<td width="50%" valign="top">
                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                        <tr>
                            <td valign="top" class="tab-style-header">Special Instructions&nbsp;</td>
                        </tr>
                        <tr>	
                            <td>
                                <table width="100%" border="0" cellpadding="2" cellspacing="0">
                                <% if basket_has_multiship = false  then %>
                                    <tr>
                                        <td  align="left" class="plaintextbold"><br>Order Hold Date&nbsp;
                                    <input name="orderhold" maxlength="10" size="15" id="orderhold" value="<%=session("orderholdDate")%>">
                                    &nbsp;<a href="javascript:show_calendar('checkoutform.orderhold');"><img src="images/calendarSm.gif" border="0" align="center" WIDTH="25" HEIGHT="19"></a>
                                        </td>
                                    </tr>
                                <%end if%>   
                                    <tr><td align="left"><input type="text" name="memo_field1" id="memo_field1" value="<%=session("memo_field1")%>" maxlength="70" size="50" class="txtboxlong"></td></tr>
                                    <tr><td align="left"><input type="text" name="memo_field2" id="memo_field2" value="<%=session("memo_field2")%>" maxlength="70" size="50" class="txtboxlong"></td></tr>
                                    <tr><td align="left"><input type="text" name="memo_field3" id="memo_field3" value="<%=session("memo_field3")%>" maxlength="70" size="50" class="txtboxlong"><br><span class="smalltextblk">(70 max characters per line)</span></td></tr>
                                <% if has_backOrder_item = true then %>
                                    <tr>
                                        <td  class="plaintext">
                                        <b>Some of the items on this order are not available for immediate shipment.<br>
                                        Please choose one of the following shipment options:
                                        </b>
                                        <br>
                                        <input type="radio" name="ship_ahead" id="ship_ahead" value="1" checked>Ship immediately those items that are available.
                                        <br>
                                        <input type="radio" name="ship_ahead" id="ship_ahead" value="0">Ship all items at once when available.
                                        
                                        </td>
                                    </tr>												
                                <%end if%>
            
                                    <tr><td  colspan="2"><img src="images/clear.gif" width="1" height="5" border="0"></td></tr>
                                
                                </table>
                            </td>					
                            
                        </tr>
    
                    </table>
			
                </td>
                <td width="20"><img src="images/clear.gif" width="20" height="1" border="0"></td>
                <td width="50%" valign="top">
                    <table width="100%" cellpadding="0" cellspacing="0" border="0">
                            <tr>
                                <td valign="top" class="tab-style-header">Shipping Instructions&nbsp;<span id="chleft" class="smalltextblk">(1200 max characters)</span></td>
                            </tr>
                            <tr>
                                <td align="left">
                                <br>
                                    <textarea name="ordermemo" id="ordermemo" class="txtboxlong" cols="40" rows="7" onKeyUp="Contchar('ordermemo','chleft','{CHAR} characters left',1200);"><%=session("ordermemo")%></textarea>							
                                </td>						
                            </tr>
                            <tr><td><img src="images/clear.gif" alt="" border="0" width="1" height="5"></td></tr>
                        
                    </table>
				</td>
            </tr>
		</table>
		
		<br>
	<%
		show_cod = false
		show_terms = false
		show_echeck=false
		show_paypal=false
		show_points =false
		show_giftcard =false
		
				
		if SL_ALLOWCOD = true then
			show_cod=true
		end if 
		
		if SL_ALLOWTERMS = true and session("registeredshopper")="YES" then
			'check if terms is available
			if terms_avail = true then
				show_terms=true
			end if
		end if
		
		if WANT_ECHECK=1 then
			show_echeck=true
		end if
		
		if WANT_PAYPAL=1 and PayPal_LoginID <> "" then
			show_paypal=true
		end if 
		
		session("Remaining_balance")= 0
		session("Redeem_Amount")=0 
		session("Points_need") = 0
		
		'response.write("--->"&session("pointavail"))
		
		if SL_ALLOWPOINTS=false then
			ALLOW_POINTS=0
		end if
		
		'response.write(SL_POINT_REDEEEM_VAL_AMT)
		
		'if ALLOW_POINTS=1 
		'Has_itemsWithPoints
		'response.write(SL_PrintGrandTotal)
		if ALLOW_POINTS=1 and session("pointavail")> 0 then 
		

			if SL_POINT_REDEEEM_VAL=1 then											
				session("Points_need") = cdbl((SL_PrintGrandTotal)/(SL_POINT_REDEEEM_VAL_AMT) * (1.00))					
				
				session("Points_need")  = round(session("Points_need"),2)
				point_mod = cdbl(session("Points_need") - fix(session("Points_need")) )	
				
				'session("points_need") = round(session("points_need"),0)
				'session("Points_need") = fix(session("Points_need")) + 1
				'if point_mod > 0 then
					'session("Points_need") = cint(session("Points_need")) + 1
				'	session("Points_need") = cdbl(session("Points_need")) + 1
				'end if				
				if point_mod > 0 then
				    session("Points_need") = fix(session("Points_need")) + 1				
				end if 
				session("Redeem_Amount") = cdbl(session("pointavail")*(SL_POINT_REDEEEM_VAL_AMT))
				
				if SL_PrintGrandTotal <= session("Redeem_Amount") then
					session("Redeem_Amount") = SL_PrintGrandTotal
				else
					session("Points_need")= session("pointavail")
				end if 
				
				session("Remaining_balance") =(SL_PrintGrandTotal)-session("Redeem_Amount")
				
				if session("Remaining_balance") < 0 then
					session("Remaining_balance")= 0 
				end if
			end if
			
			if SL_POINT_REDEEEM_VAL=2 then  
			' item level point redemption is done on basket page
				session("DoNotShow_PointsTab")=true
			end if 
					
			if session("DoNotShow_PointsTab")=false then
				show_points=true
			end if
			
		end if
		'response.write(show_points)
		
		if ALLOW_GC=1 and session("DoNotShow_GCTab")=false then
			show_gcredemption=true			
		end if
		
		if len(SL_ALLGIFTCARD) > 0 and session("DoNotShow_GCardTab")=false and ALLOW_GCARD=1 then
		       show_giftcard=true
		end if
		
        'response.Write("Balance -->"&session("Remaining_BalanceAfterGC"))
        'response.Write("<br>")
        'response.Write(formatcurrency(0))
        'response.Write("R Balance -->"&session("Remaining_balance"))
        'response.Write(session("DoNotShow_PointsTab"))
                

	%>

	
	
					<table width="100%" cellpadding="3" cellspacing="1" border="0"> 
						<tr>
						<td valign="top" align="center" class="nopadding">

						<div class="tab-container" id="container1">
						<ul class="tabs">						
						
                            <li><a href="#" onClick="return showPane('pane1', this)" id="tab1">Credit Card</a></li>
						<%if show_echeck=true then %>						
                            <li><a href="#" onClick="return showPane('pane2', this)" id="tab2">E-Check</a></li>
						<%end if%>
						<%if show_cod=true then%>						
                            <li><a href="#" onClick="return showPane('pane3', this)" id="tab3">COD</a></li>
						<%end if%>
						<%if show_terms=true then%>						
                            <li><a href="#" onClick="return showPane('pane4', this)" id="tab4">Terms</a></li>
						<%end if%>
						<%if show_paypal=true then%>						
                            <li><a href="#" onClick="return showPane('pane5', this)" id="tab5">PayPal</a></li>
						<%end if%>
						<%if show_points=true then%>						
                            <li><a href="#" onClick="return showPane('pane6', this)" id="tab6">Redeem Points</a></li>
						<%end if%>
						<%if show_gcredemption=true then%>												
                            <li><a href="#" onClick="return showPane('pane7', this)" id="tab7">Redeem Gift Certificate</a></li>
						<%end if%>	
						<%if show_giftcard=true then%>												
                            <li><a href="#" onClick="return showPane('pane8', this)" id="tab8">Redeem Gift Card</a></li>
						<%end if%>	
											
						</ul>


						<div class="tab-panes">
						
						<div id="pane1">
						<%
								
								objDoc.loadxml(SL_CardListXmlstream)			
								set SL_CCInfo = objDoc.selectNodes("//cclist")
							
								%>
								<table width="75%" border="0">
									<tr>
									   <td class="plaintextbold"  align="right">Name On Card:&nbsp;</td>
									   <td align="left"><input type="text" name="txtcc_name" value="<%=session("cc_name")%>" size="40,1">
									   <input type="hidden" name="payment_type" id="payment_type" value="Credit Card">
									   <input type="hidden" name="Applypoints" id="applypoints" value="0">
									   <input type="hidden" name="ApplyGC" id="applygc" value="0">
									   </td></tr>
									<tr>
									  <td class="plaintextbold"  valign="top" align="right">Card Number:&nbsp;</td>
									  <td align="left"><input type="text" id="txtcc_number" name="txtcc_number" value="<%=session("cc_number")%>" maxlength="20" size="40,1" autocomplete="off">&nbsp;<span class="plaintextbold">
											<br>(No dashes in between)</span></td>
										</tr>
									<tr>
										<td align="right" class="plaintextbold">Type:&nbsp;</td>
										<td align="left" class="plaintext"  valign="top">
								
										<select NAME="cc_type" id="cc_type" class="plaintext">
										
										<option value="">Please select card type
										
										<%
									  
										for x=0 to SL_CCInfo.length-1 
										cardcode = SL_CCInfo.item(x).selectSingleNode("cc_code").text
										cardname = SL_CCInfo.item(x).selectSingleNode("cc_name").text
										ccextra  = SL_CCInfo.item(x).selectSingleNode("ccextra").text
										%>
										<option value="<%=cardcode%>/<%=ccextra%>" <% if session("cc_type") = cstr(cardcode) then Response.Write(" selected") end if %>><%=cardname%>
										<%next%>
									</select>
									<%
										for x=0 to SL_CCInfo.length-1 
									ccextra  = SL_CCInfo.item(x).selectSingleNode("ccextra").text									
									val1 = ccextra
								%>
								<% if val1 = "V" then %>
									 <img src="images/visa.gif" width="51" height="33" border="0" align="absmiddle"> 
								<%end if%>
								<% if val1 = "M" then %>
									 <img src="images/MasterCard.gif" width="50" height="33" border="0" align="absmiddle"> 
								 <%end if%>
								 <% if val1 = "D" then %>
									 <img src="images/DiscoverCard.gif" width="50" height="33" border="0" align="absmiddle"> 
								 <%end if%>
								 <% if val1 = "A" then %>
									<img src="images/Amex.gif" width="50" height="33" border="0" align="absmiddle"> 
								 <%end if%>
								<%next%>				
									<% 
									Set SL_CCInfo = nothing
									%>
									</td>
								</tr>
								<tr>
									<td align="right" class="plaintextbold">Expiration:&nbsp;</td>
									<td align="left">
									<%dim lamonths(12)
										 lamonths(1)="January"
										 lamonths(2)="February"
										 lamonths(3)="March"
										 lamonths(4)="April"
										 lamonths(5)="May"
										 lamonths(6)="June"
										 lamonths(7)="July"
										 lamonths(8)="August"
										 lamonths(9)="September"
										 lamonths(10)="October"
										 lamonths(11)="November"
										 lamonths(12)="December"
										 
										 IMonth =Month(now) 
										%>
										<select NAME="cc_expmonth" id="cc_expmonth" class="plaintext">
										<%for x=1 to 12
											if x <10 then
												showthis = "0"+cstr(x)+" "
											else
												showthis= cstr(x)+" "
											end if				
										%>
											<option value="<%=x%>" 
											<% if session("cc_expmonth") = cstr(x) then Response.Write(" selected") end if %>
											<% if len(trim(session("cc_expmonth"))) = 0 and (x= IMonth) then Response.Write(" selected") end if %>
											><%=showthis%> <%=lamonths(x)%>
										<%next%>
										</select>
										&nbsp;&nbsp;&nbsp;
										<select NAME="cc_expyear" id="cc_expyear" class="plaintext">
										  <%for x=0 to 9%>
											  <option value="<%=(year(date())+x)%>" <% if session("cc_expyear") = cstr(year(date())+x) then response.write(" selected") end if%>> <%=year(date())+x%> 
										  <%next%>
									 </select>						
									</td>	
								</tr>
								<%if WANT_AUTHNET=1 then %>
								<tr>
									<td class="plaintextbold" align="right" valign="top">Credit Card ID:&nbsp;</td>
									<td align="left" valign="top"><input type="text" name="txtcc_id" value="<%=session("cc_id")%>" maxlength="10" size="10" autocomplete="off">
										&nbsp;&nbsp;&nbsp;&nbsp;<a class="allpage" href="javascript:openlearnmore();">Learn More</a>
									</td>
								</tr>
								<%end if%>		
								</table></div>
						
						<div id="pane2">
						<table width="75%" border="0">
									<tr><td>&nbsp;</td>
										<td><img src="images/eCheckNet.gif" alt="Safe and Secure" border="0" align="middle" width="249" height="46"></td>
									</tr>						 
									<tr>
									   <td class="plaintextbold" align="right">Bank Name:&nbsp;</td>
									   <td align="left"><input type="text" name="bankname" maxlength="30" value="<%=session("bankname")%>" size="40,1">
									   
									</tr>
									<tr>
								   <td class="plaintextbold"  align="right">Type of Account:&nbsp;</td>
								   <td align="left">
								   <select name="accttype" class="plaintext">				
										<option value="Savings" <%if session("accttype")="Savings" then response.write(" selected") end if%>>Savings
										<option value="Checking" <%if session("accttype")="Checking" then response.write(" selected") end if%>>Checking				
									</select>
								   </td>
									</tr>
									<tr>
								   <td class="plaintextbold"  align="right">Routing Number:&nbsp;</td>
								   <td align="left"><input type="text" name="bankrountingnum" maxlength="9" value="<%=session("bankrountingnum")%>" size="40,1" autocomplete="off"></td>
								</tr>
								<tr>
								   <td class="plaintextbold"  align="right">Account Number:&nbsp;</td>
								   <td align="left"><input type="text" name="bankacctnum" maxlength="30" value="<%=session("bankacctnum")%>" size="40,1" autocomplete="off">
								   &nbsp;<a class="allpage" href="javascript:openlearnmore1();">Learn More</a>
								   </td>
								</tr>
									
									
									</table>
						</div>
						
						<div id="pane3">
						<br><br>
						<table width="100%" border="0">
						<tr><td align="center" class="plaintext">
						(Additional shipping charges may apply)
						</td></tr>
						</table>
						</div>
						
						<div id="pane4">
						<br><br>
						<table width="100%" border="0">
						<tr><td align="center" class="plaintextbold">PO Number <input type="text" name="ponum" maxlength="15" size="20" class="plaintext">
						</td></tr>
						</table>
						</div>
						
						<div id="pane5">
						<table border="0" cellpadding="0">
												<tr>
													<td valign="top">
													<!-- PayPal Logo -->
													<table border="0" cellpadding="0" cellspacing="0" align="center">
														<tr>
															<td align="center"><a href="#" onClick="javascript:window.open('https://www.paypal.com/us/cgi-bin/webscr?cmd=xpt/cps/popup/OLCWhatIsPayPal-outside','olcwhatispaypal','toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes, width=400, height=350');"><img  src="images/horizontal_solution_PP.gif" border="0" alt="PayPal"></a></td>
														</tr>
													</table>													 
													<!-- PayPal Logo -->
													<!--<img border="0" src="images/Paypal_logo2.gif" width="294" alt="Save time.  Checkout securely.Pay without sharing your financial information" height="25" align="middle">-->
													</td>
												</tr>
												<tr><td><span style="font-size:11px; font-family: Arial, Verdana;">Save time.  Checkout securely.  
												<br>
													Pay without sharing your financial information.
												<br>
												<!--<b>You must have a valid Pay Pal Account.</b>-->
												</span></td></tr>
										
												</table>
						</div>
						<div id="pane6">
						<br><br>
				
						<table width="100%" border="0">
						<tr><td align="center" class="plaintext">
						
						<table cellpadding="5">
						
						<tr><td class="plaintextbold">
								You have accumulated total points of <%=session("pointavail")%> for this redemption.															
								</td>
							</tr>
							<% if SL_POINT_REDEEEM_VAL=1 then %>
								<tr><td class="plaintextbold">You can apply <%=session("Points_need")%> leaving a balance of 
								<% if session("Remaining_balance") > 0 then %>
									<%=formatcurrency(session("Remaining_balance"))%> for this order.
								<%end if%>
								<% if session("Remaining_balance") <= 0 then %>
								  $0 for this order.
								  <br>
								  No additional payment will be needed.
								<%end if%>
							</td></tr>
							
							<%end if%>
							
							
							
							<tr><td class="plaintextbold">
								You will have a total of <%=(session("pointavail")-session("Points_need"))%> points remaining. 
							</td>
							</tr>
						</table>
						<table width="90%" cellpadding="5">
						<tr><td class="plaintext" align="center">
						<font color="red">If you do not wish to redeem points please select a different type of payment otherwise please continue.</font>
						</td></tr>						
						</table>
					
					
						</td></tr>
						</table>
						</div>
						
						<div id="pane7">
						<br><br><br>
						<table width="100%" border="0">
						<tr><td align="center" class="plaintextbold">Enter Gift Certificate Number <input type="text" name="gcnumber" value="" maxlength="15" size="20" class="plaintext" autocomplete="off" onblur="extractNumber(this,0,true);" onKeyUp="extractNumber(this,0,true);" onKeyPress="return blockNonNumbers(this, event, false, true,15);">&nbsp;(Numeric Values Only)
						</td></tr>
						</table>
						</div>
						
						
						<div id="pane8">
						<br><br><br>
						<table width="100%" border="0">
						<tr><td align="center" class="plaintextbold">Enter Gift Card Number <input type="text" name="gcardnumber" maxlength="50" size="50" class="plaintext" value="" autocomplete="off" onblur="extractNumber(this,0,true);" onkeyup="extractNumber(this,0,true);" onkeypress="return blockNonNumbers(this, event, false, true,40);">&nbsp;(Numeric Values Only)
						</td></tr>
						</table>
						</div>
						
						</div> 
						
						
						<!-- panes -->
						</div>
			
						</td>
					</tr>
					</table>
	
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
	
	<tr><td colspan="2">&nbsp;</td></tr>
	
	<tr><td width="40%" class="smalltextblk" align="center">&nbsp;By clicking "SUBMIT YOUR ORDER" your order will be complete and payment(s) will be applied.</td>
		<td style="padding-left:10px;"><input type="Image" value="Submit Order" src="images/btn-submitorder.gif" border="0" alt="Submit Order"></td>
		<td><span id="buySAFE_Kicker" name="buySAFE_Kicker" type="Kicker Guaranteed Ribbon 200x90"></span></td>
	</tr>
	
	</table>
	
	<%else%>
	
	<table width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr><td valign="top" colspan="3" class="THHeader">&nbsp;<input type="hidden" id="needsourcekey" value="1"></td></tr>
	<tr>
		<td valign="top" class="sidenavTxt" width="1"><img src="images/clear.gif" alt="" border="0" width="1" height="1"></td>
		<td valign="top" align="center">
		<table width="0" cellpadding="5" cellspacing="1" border="0"> 
		<tr><td class="plaintextbold" align="center" colspan="2">
    		
		    <font color="red">No additional payment is required on this order
		    <br />
		    <!--  Click the 'Continue' button to finish this order.-->
		    </font>
		    <br /><br />
		    <input type="hidden" name="SITELINK" id="SITELINK" value="SITELINK">						
		    </td></tr>
    	<tr><td width="40%" class="smalltextblk" align="center">&nbsp;By clicking "SUBMIT YOUR ORDER" your order will be complete and payment(s) will be applied.</td>
    	    <td><input type="Image" value="Submit Order" src="images/btn-submitorder.gif" border="0" alt="Submit Order"></td>
		</table>
		
		</td>
		<td valign="top" class="sidenavTxt" width="1"><img src="images/clear.gif" alt="" border="0" width="1" height="1"></td>
	</tr>
	
	<tr><td valign="top" class="sidenavTxt" colspan="3"><img src="images/clear.gif" alt="" border="0" width="1" height="1"></td></tr>
	
	</table>
	
	<br>
	<%end if  'if SL_ORD_TOTAL %>
	
	
	
	<% 
	'response.write(SL_PrintGrandTotal)
	 'points_used = ORDER_RECORD.points_usd
	 ' points_amt = ORDER_RECORD.checkamoun
	 'response.write(DoNot_Pass_CartContents)
	 TotalLineItems = SL_Basket.length	 
	 if (points_used > 0 and points_amt > 0) or basket_has_giftcert=true or basket_has_NegativePrice=true or TotalLineItems > 99 or GiftCardAmount > 0 then
		  	PayPal_Order_total = FORMATCURRENCY(SL_PrintGrandTotal)
			PayPal_Order_total =  mid(PayPal_Order_total,2,len(PayPal_Order_total))	 
			DoNot_Pass_CartContents = true 
	    else
	  		PayPal_Order_total = FORMATCURRENCY(SL_ORD_TOTAL)
			PayPal_Order_total =  mid(PayPal_Order_total,2,len(PayPal_Order_total))
	    end if
		
	 'response.write(SL_PrintGrandTotal)
	 
	  'shiptotal = FORMATCURRENCY(sitelink.ORD_SHIPPING(session("ordernumber"),false))
	  'taxtotal  = FORMATCURRENCY(sitelink.ORD_TAX(session("ordernumber"),false)) 
			
	  shiptotal = FORMATCURRENCY(SL_PrintShipAmount)
	  taxtotal  = FORMATCURRENCY(SL_PrintTaxAmount) 
	  paypalaccountid = sitelink.Get_PayPalLogin()	 
	  paypalcurrency  =sitelink.Get_PayPalCurrencyCode()
	  

	  shipping_for_Paypal = mid(shiptotal,2,len(shiptotal))
	  taxtotal_for_Paypal =  mid(taxtotal,2,len(taxtotal))
	  
	  successurl = secureurl + cstr("processorder.asp")
	  cancelurl = secureurl + cstr("checkout.asp")
	  notifyurl = secureurl + cstr("processorder.asp")
	  imageurl  = secureurl + "images/"+ cstr(COMPANY_LOGO_IMG)
	  first_name = session("firstname")
	  last_name=  session("lastname")
	  address1 =session("address1")
	   address2 =session("address2")
	   city = session("city")
	   state = session("state")
	   zip= session("zipcode")	
	   
	   if DoNot_Pass_CartContents = false then %>
						
	<input type="hidden" name="cmd" value="_cart">			
	
	<%

	for x=0 to SL_Basket.length-1		 
		 SL_BasketDesc		= SL_Basket.item(x).selectSingleNode("inetsdesc").text
		 SL_BasketDesc      = replace(SL_BasketDesc,"""","")
		 SL_BasketUnitPrice	= formatcurrency(SL_Basket.item(x).selectSingleNode("extendp").text)
		 SL_BasketUnitPrice = mid(SL_BasketUnitPrice,2,len(SL_BasketUnitPrice))		 		 
		 SL_BasketQuantity	= 1
		 
		 actual_Quantity = SL_Basket.item(x).selectSingleNode("quanto").text
		 actual_Quantity = replace(actual_Quantity,".00","")
		 actual_Unitprice = formatcurrency(SL_Basket.item(x).selectSingleNode("it_unlist").text)
		 actual_Unitprice = mid(actual_Unitprice,2,len(actual_Unitprice))	
		 		 	 
		 SL_Number	= SL_Basket.item(x).selectSingleNode("item").text
		 SL_Basket_Shipto	= SL_Basket.item(x).selectSingleNode("ship_to").text
		 SL_Basket_Shipto = SL_Basket_Shipto + 1 - 1
		 
		 if len(trim(SL_BasketDesc)) = 0 then
		    SL_BasketDesc = SL_Number
		 end if
	%>
	<input type="hidden" name="item_name_<%=x+1%>" value="<%=SL_BasketDesc%>">
	<input type="hidden" name="amount_<%=x+1%>" value="<%=SL_BasketUnitPrice%>">
	<input type="hidden" name="quantity_<%=x+1%>" value="<%=SL_BasketQuantity%>">
	<input type="hidden" name="item_number_<%=x+1%>" value="<%=SL_Number%>">
	<%
		option_val = actual_Quantity + "@" + actual_Unitprice
		if SL_Basket_Shipto > 0 then 
		 	fname = SL_Basket.item(x).selectSingleNode("firstname").text
		 	lname = SL_Basket.item(x).selectSingleNode("lastname").text
		 	fullname = (cstr(fname) + " " + cstr(lname))	
		 	fullname = "  Shipping to "+ cstr(fullname)
			option_val = option_val +  fullname
		end if
	%>
	
	<input type="hidden" name="on0_<%=x+1%>" value="<%=option_val%>">
	
	<%next%>
		
	<input type="hidden" name="shipping_1" value="<%=shipping_for_Paypal%>">
	<input type="hidden" name="tax_cart" value="<%=taxtotal_for_Paypal%>">
	<%else%>
		<input type="hidden" name="cmd" value="_xclick">
		<input type="hidden" name="amount" value="<%=PayPal_Order_total%>">
	<%end if %>
	
	<input type="hidden" name="invoice" value="<%=session("ordernumber")%>">
	<input type="hidden" name="item_name" value="Web Order: <%=session("ordernumber")%>">
	<input type="hidden" name="currency_code" value="<%=paypalcurrency%>">	
	<input type="hidden" name="return" value="<%=successurl%>">
	<input type="hidden" name="cancel_return" value="<%=cancelurl%>">
	<input type="hidden" name="notify_url" value="<%=notifyurl%>">
	<input type="hidden" name="image_url" value="<%=imageurl%>">
	<input type="hidden" name="first_name" value="<%=first_name%>">
	<input type="hidden" name="last_name" value="<%=last_name%>">
	<input type="hidden" name="address1" value="<%=address1%>">
	<input type="hidden" name="address2" value="<%=address2%>">
	<input type="hidden" name="city" value="<%=city%>">
	<input type="hidden" name="state" value="<%=state%>">
	<input type="hidden" name="zip" value="<%=zip%>">
	<input type="hidden" name="no_note" value="1">
	<input type="hidden" name="rm" value="2">
	<input type="hidden" name="mrb" value="R-1V748593L7668510W">
	<input type="hidden" name="pal" value="R-NCY5RMVCQZZDG">
	<input type="hidden" name="bn" value="Dydacomp">
	<input type="hidden" name="business" value="<%=paypalaccountid%>">
	<input type="hidden" name="upload" value="1">
	<input type="hidden" name="cbt" value="Press this button to complete your order">
	  
	
	
	</form>
	

	
	
	
	
	<%set SL_Basket = nothing %>
	</div>
<!-- end sl_code here -->
	</td>
	
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
