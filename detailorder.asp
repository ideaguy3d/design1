<%on error resume next%>



<% Response.Expires = 0 %>
<% 

if session("callcenter_logged")=true then
	session("registeredshopper")="YES"
end if


  if session("registeredshopper")="NO" then
  	response.redirect("statuslogin.asp")
  end if 


dim ThisOrderNumber 
	 ThisOrderNumber= REQUEST.FORM("OrdNum")
	 
	 if isnumeric(ThisOrderNumber) = false then
	    ThisOrderNumber= 0 
	 end if
	 
	 if(len(trim(ThisOrderNumber))) = 0 then
	    response.redirect("statuslogin.asp")
	 end if
	 
	  if (ThisOrderNumber <= 0) then
	    response.redirect("statuslogin.asp")
	 end if
	 
	 
	 ThisOrderNumber = ThisOrderNumber+1-1
	 
	 MOMOrd  = REQUEST.FORM("MOMOrdNum")
	 	 
	 if isnumeric(MOMOrd) = false then
	    MOMOrd= 0 
	 end if
	 
	 
	 if MOMOrd = 1 then
	 	IsMOMOrdNum = true
		SL_orderNum = REQUEST.FORM("SL_OrderNum")
	 else
	   IsMOMOrdNum = false
	   SL_orderNum = ThisOrderNumber
	 end if
	
%>
<!--#INCLUDE FILE = "include/momapp.asp" -->

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<%
     'if IsMOMOrdNum = true then
	 multishipto = 0 
	 points_amt=0
	 Points_onOrderTotal = false	 
		 set ORDER_RECORD = sitelink.ORDER_RECORD(ThisOrderNumber,IsMOMOrdNum)	 
			 SL_paymentType = ORDER_RECORD.paymethod
			 SL_Ship_type = ORDER_RECORD.shiplist
			 SL_Po_Number = ORDER_RECORD.ponumber
			 SL_Ship_Number = ORDER_RECORD.shipnum
			 
			 if SL_Ship_Number > 0 then
			 	multishipto = 0 
			 end if 

			points_used = ORDER_RECORD.points_usd
			points_amt = ORDER_RECORD.checkamoun
			ord_total  = ORDER_RECORD.ord_total
			if points_used > 0 and points_amt > 0 then			  
				remaining_balance = ord_total-points_amt
				Points_onOrderTotal= true	
			end if
			
			if IsMOMOrdNum=false then
			    GiftCardAmount = ORDER_RECORD.gcardamt
			 else
			    GiftCardAmount = 0
			 end if 
			 
			 
		 set ORDER_RECORD = nothing
	 'end if 

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=althomepage%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
<link rel="stylesheet" href="text/footernav.css" type="text/css">
<link rel="stylesheet" href="text/design.css" type="text/css">
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
<div id="page-content">
	
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%" >Order History</td></tr>
	</table>
	<br>
	
	<p align="center"><font class="plaintextbold"><FONT size="4" face="arial" ><b>Order Number: &nbsp;<%=ThisOrderNumber%></b></font></font><br>
					<font class="plaintextbold">Order Placed:&nbsp;					
					<%=sitelink.dataforstatusnew("ORDERDATE",ThisOrderNumber,session("shopperid"),"EMAIL",IsMOMOrdNum)%>
					
					<br><br>
					<a class="allpage" href="repeatorder.asp?ordernum=<%=ThisOrderNumber%>&ismom=<%=MOMOrd%>">Click here to Repeat this order</a>
					
					
					</font>
				</p>
				<TABLE bordercolor="#000000" cellSpacing="0" cellPadding="1" width="98%" border="1">
		        	<TR><TD colspan="2"  class="THHeader" align="center" height="40" ><b>Order and Shipping Information</b></TD></TR>
					<TR>
			         <TD colspan="2">
					 <TABLE cellSpacing="0" cellPadding="2" width="100%" border="0" >
						 <tr><td width="100%" colspan="2">
							 <table width="100%" border="0" cellspacing="1" cellpadding="2" >
							 <%
							 	'multishipto= 0
								default_carrier =  sitelink.dataforstatusnew("SHIPPINGMETHOD",ThisOrderNumber,session("shopperid"),"EMAIL",IsMOMOrdNum)
								
								xmlCarrier = sitelink.SHIPPINGMETHODS() 
								objDoc.loadxml(xmlCarrier)
								set SL_Ship = objDoc.selectNodes("//shipmeths") 
								
								if len(trim(default_carrier)) = 0 then
								  default_carrier = SL_Ship_type 
								end if 
																
							 	xmlstring = sitelink.getbasketinfo(session("shopperid"),ThisOrderNumber,"",IsMOMOrdNum,true)
				
								objDoc.loadxml(xmlstring)	
				
								filter_condition  ="//gbi/ship_to[. > 0]"
								set SL_Basket_shipto = objDoc.selectNodes(filter_condition)
									'response.write("ttttt---> "& SL_Basket_shipto.length)
									if SL_Basket_shipto.length > 0 then
									 	multishipto = 1
									end if 
								set SL_Basket_shipto = nothing

					            basket_has_giftcert = false
					            giftcertitem=""
					            giftcertAmount=0
		                        filter_condition = "//gbi[certid>0]"
					            set SL_Basket_giftcert = objDoc.selectNodes(filter_condition)
					            if SL_Basket_giftcert.length > 0 then
						            basket_has_giftcert = true
						            giftcertitem = SL_Basket_giftcert.item(0).selectSingleNode("item").text
						            giftcertAmount = SL_Basket_giftcert.item(0).selectSingleNode("it_unlist").text
						            giftcertAmount = abs(giftcertAmount)
					            end if 
					            set SL_Basket_giftcert = nothing


								basket_total = 0 
								GlobalDiscount = 0
		
					            if basket_has_giftcert=true then
					                filter_condition = "//gbi[item !='"+cstr(giftcertitem)+"']"					    
					            else
					                filter_condition = "//gbi"
					            end if
					            set SL_Basket = objDoc.selectNodes(filter_condition)
								
							 %>
							 <tr>
							 	<th align="CENTER"  class="THHeader">Qty</th>
								<th align="CENTER"  class="THHeader">Description</th>
								<th align="CENTER"  class="THHeader">Status</th>
								<th align="CENTER"  class="THHeader">Total</th>
							 </tr>
							 <% for x=0 to SL_Basket.length-1 
							 SL_Number			= SL_Basket.item(x).selectSingleNode("number2").text
							 SL_Variant			= SL_Basket.item(x).selectSingleNode("variant").text
							 SL_BasketNumber 	= SL_Number +" "+ SL_Variant
							 SL_basket_item	    = SL_Basket.item(x).selectSingleNode("item").text
							 SL_BasketQty		= SL_Basket.item(x).selectSingleNode("quanto").text							 
			  				 SL_BasketDesc		= SL_Basket.item(x).selectSingleNode("inetsdesc").text
				             SL_BasketDesc2		= SL_Basket.item(x).selectSingleNode("desc2").text
			 				 SL_BasketUnitPrice	= SL_Basket.item(x).selectSingleNode("it_unlist").text
							 SL_BasketExtPrice	= SL_Basket.item(x).selectSingleNode("extendp").text
							 SL_BasketCustInfo	= SL_Basket.item(x).selectSingleNode("custominfo").text
							 SL_Itemstate	    = SL_Basket.item(x).selectSingleNode("item_state").text
							 SL_BasketExtPrice	= SL_Basket.item(x).selectSingleNode("extendp").text
							 SL_Basket_Shipto    =SL_Basket.item(x).selectSingleNode("ship_to").text
							 'basket_total		= basket_total + SL_BasketExtPrice
							 SL_Record			=SL_Basket.item(x).selectSingleNode("record").text
							 SL_Record			=SL_Record + 1 - 1
							 SL_BasketQty = replace(SL_BasketQty,".00","")
							 
			 				SL_Points_ned       = SL_Basket.item(x).selectSingleNode("points_ned").text
							SL_PointsRdm		=SL_Basket.item(x).selectSingleNode("pts_rdeemd").text

							 SL_Points_ned	 	= SL_Points_ned * SL_BasketQty
							
							 
							 if SL_Basket_Shipto > 0 then
							 	 fname				= SL_Basket.item(x).selectSingleNode("firstname").text
								 lname				= SL_Basket.item(x).selectSingleNode("lastname").text
								 company			= SL_Basket.item(x).selectSingleNode("company").text
								 addr				= SL_Basket.item(x).selectSingleNode("addr").text
								 addr2				= SL_Basket.item(x).selectSingleNode("addr2").text
								 city				= SL_Basket.item(x).selectSingleNode("city").text
								 vstate				= SL_Basket.item(x).selectSingleNode("state").text
								 zipcode			= SL_Basket.item(x).selectSingleNode("zipcode").text
								 country			= SL_Basket.item(x).selectSingleNode("country").text
							 end if
							 SL_ship_when       = SL_Basket.item(x).selectSingleNode("ship_when").text
							 	
							 SL_ship_via        = SL_Basket.item(x).selectSingleNode("ship_via").text						 
							 SL_ShippingTitle = ""
							 if len(trim(SL_ship_via)) > 0 then							
								'SL_ShippingTitle = SL_Basket.item(x).selectSingleNode("ca_title").text
								'find what is ca_title.. it may or may not be available on web
								for y = 0 to SL_Ship.length-1								
									if ucase(SL_ship_via) = SL_Ship.item(y).selectSingleNode("ca_code").text then
										SL_ShippingTitle = SL_Ship.item(y).selectSingleNode("ca_title").text 
									end if								
								next
								
								if len(trim(SL_ShippingTitle)) = 0 then
									SL_ShippingTitle = SL_ship_via
								end if 
							 else
							 	SL_ShippingTitle = default_carrier
							 end if
							 							 
								basket_total		= basket_total + SL_BasketExtPrice
							
							 if (x mod 2) = 0 then
							   class_to_use = "tdRow1Color"
							else
								class_to_use = "tdRow2Color"
							end if 
							 %>
							 <tr>
							
							 <td class="<%=class_to_use%>" align="center" ><span class="plaintext">
							 <%if SL_Number<>"" then %>
							<%=SL_BasketQty%>
							<%else%>
							&nbsp;
							<%end if%>
							 </span></td>
							 <td class="<%=class_to_use%>">
							 <span class="plaintext">
							
							 <%=SL_BasketDesc%>			
								<% if SL_BasketDesc2<>"" then response.write("<br>" +SL_BasketDesc2) end if %>
								<br>
								<b>Item :</b><%=SL_basket_item%> &nbsp; <b>Price : </b><%=formatcurrency(SL_BasketUnitPrice)%> &nbsp;				
								<%if SL_BasketDiscount > 0 then   response.write("<b>Discount :</b>"+ SL_BasketDiscount +"%") end if %>
								<%if SL_BasketCustInfo <> "" then response.write("<br>" + SL_BasketCustInfo) end if %>
								<%if SL_PointsRdm="true" then%>
									<br><%=SL_Points_ned%> points were redeemed to purchase this product
								<%end if%>
								<%if SL_Number<>"" then 
								if SL_Basket_Shipto > 0 then
								'if multishipto = 1 then																								
								%>
								<br><br>
								<b>Ship To:</b>
								<%
								
								stringto_write ="&nbsp;" + (lname) +" " + (fname)
								if len(company) > 0 then
									stringto_write = stringto_write + "<br>&nbsp;" + (company)
								end if 
								if len(addr) > 0 then
									stringto_write = stringto_write + ", " + (addr)
								end if 
								if len(addr2) > 0 then
									stringto_write = stringto_write + ", " + (addr2)
								end if 
			
								stringto_write = stringto_write + "<br>&nbsp;" + (city)
								stringto_write = stringto_write + ", " + (vstate)
								stringto_write = stringto_write + ", " + (zipcode)
								stringto_write = stringto_write + ", " + (sitelink.countryname(country))
			
								response.write(stringto_write)
								
								%>
								<%else%>
									<% if multishipto = 1 then %>
										<br><br>
										<b>Ship To:</b>Same as Bill to
									<%end if%>
								<%end if%>								
								<br><br>								
								<b>Ship Via :</b><%=SL_ShippingTitle%>
								
								<%end if%>

							 </span>
							 </td>
							 <td class="<%=class_to_use%>" align="center"><span class="plaintext">
							 <% line_item_status = sitelink.item_status(SL_Itemstate,SL_Record,IsMOMOrdNum)
							    line_item_status = replace(line_item_status,"Committed", "In Progress")
							 %>
							 
							 <%=line_item_status%></span></td>							 
							 <td class="<%=class_to_use%>" align="right"><span class="plaintext"><%=formatcurrency(SL_BasketExtPrice)%></span>&nbsp;&nbsp;</td>
							 </tr>
							 
							 <%next%>							 
							 <%
							 	set SL_Basket = nothing
								set SL_Ship = nothing
							 %>
							 <tr>
							 <td colspan="3"  align="right" class="THHeader">Subtotal:</td>
							 <td class="THHeader"  align="right"><%=formatcurrency(cdbl(basket_total))%>&nbsp;</td>
							 </tr>
							
							 
							 </table>
					 		 
					 
					</table>		 
					 </tr>
				</table>
				
		<P>&nbsp;

				<TABLE bordercolor="#000000" cellSpacing="0" cellPadding="5" width="98%" border="1">
			        <TR><TD  align="center" height="40"  class="THHeader" ><b>Billing Information</b></TD></TR>
					<TR>
			         <TD width="100%">
						 <TABLE cellSpacing="0" cellPadding="3" width="100%" border="0" >
						 <tr><td width="100%">
							 <table width="100%" border="0" cellspacing="0" cellpadding="3">
				                <tr> 
									<td width="50%" valign="top" class="tdRow1Color">
				    				  <table width="100%" border="0" cellpadding="3" cellspacing="1" >
										  <tr><td class="plaintextbold" width="40%" valign="top">Payment Type: </td>
											  <td class="plaintext" valign="top">
											  <%
											  pay_type = SL_paymentType
											  
											  select case SL_paymentType
											  	case "CC"
													pay_type = sitelink.dataforstatusnew("CARDTYPE",ThisOrderNumber,session("SHOPPERID"),"EMAIL",IsMOMOrdNum)
												case "IN"	
													pay_type = "Terms"													
												case "CO"	
													pay_type = "COD"
					                            case "CK"
					                                pay_type = "Check"
					                                if points_used > 0 then
						                                pay_type = "Points"
						                            end if 
						                            if basket_has_giftcert=true then
						                                pay_type = "Gift Certificate"
						                            end if
												case "EC"
													pay_type = "E-Check"
												case "PP"
													pay_type = "PayPal"
													
											  END SELECT
											%>
											<%=pay_type%> <% if len(trim(SL_Po_Number)) > 0 then Response.Write("<br>(PO Number :"+SL_Po_Number +")") end if%>
											  </td>
										  </tr>
										  
										  
										  
										  <%if SL_paymentType = "CC" then %>
										  <tr>
                                        <td class="plaintextbold">Last 4 digits:</td>
										      <td class="plaintext">											  
											  <%=sitelink.dataforstatusnew("CARDNUMBER",ThisOrderNumber,4,"EMAIL",IsMOMOrdNum)%>
											  </td>
										  </tr>
										 <%end if%>
										 <%if SL_paymentType = "EC" then %>
										 <tr>
                                        	<td class="plaintextbold" valign="top">E-Check Info:</td>
										      <td class="plaintext" valign="top">											  
											  <%
											  xmlstring=sitelink.dataforstatusnew("ECHECK",ThisOrderNumber,4,"EMAIL",IsMOMOrdNum)
											  objDoc.loadxml(xmlstring)
											  set SL_EInfo = objDoc.selectNodes("//echeckinfo")
											  for x=0 to  SL_EInfo.length-1
											   	Sl_routingnum = SL_EInfo.item(x).selectSingleNode("routingnum").text
					 							SL_accountnum = SL_EInfo.item(x).selectSingleNode("accountnum").text
												SL_accttype = SL_EInfo.item(x).selectSingleNode("accttype").text
												SL_bankname = SL_EInfo.item(x).selectSingleNode("bankname").text
											  %>
											  <table border="0" width="100%" cellpadding="3" cellspacing="0">
											  <tr><td class="plaintext"><b>Routing #</b><%=Sl_routingnum%></td></tr>
											  <tr><td class="plaintext"><b>Account #</b> <%=SL_accountnum%></td></tr>
											  <tr><td class="plaintext"><b>Type # </b><%=SL_accttype%></td></tr>
											  <tr><td class="plaintext"><b>Name #</b><%=SL_bankname%></td></tr>
											  </table>
											  
											  <% 
											  next
											  set SL_EInfo = nothing%>
											  </td>
										  </tr>
										 <%end if%>
										  <tr><td class="plaintextbold"><b>Billing Address: </b></td><td class="plaintext"><%=session("firstname")%>&nbsp;&nbsp;<%=session("lastname")%></td> </tr>
										  <% if len(trim(session("company"))) > 0 then%>
											  <tr><td>&nbsp;</td><td class="plaintext"><%=session("company")%></td></tr>
										  <%end if%>
										  <% if len(trim(session("address1"))) > 0 then%>
											  <tr><td>&nbsp;</td><td class="plaintext"><%=session("address1")%></td></tr>
										  <%end if%>
					  					  <% if len(trim(session("address2"))) > 0 then%>
										  <tr><td>&nbsp;</td><td class="plaintext"><%=session("address2")%></td></tr>
										  <%end if%>
										  <tr><td>&nbsp;</td><td class="plaintext"><%=session("city")%></td></tr>
										  <tr><td>&nbsp;</td><td class="plaintext"><%=session("state")%>&nbsp;&nbsp;<%=session("zipcode")%>&nbsp;&nbsp;<%=sitelink.countryname(session("country"))%></td></tr>					  
										  <% 'if multishipto = 0 and SL_orderNum > 0 then 											   
										  if multishipto = 0 then
										    'if SL_orderNum > 0 then
											 '   set shopperRecord = sitelink.SHIPTONAME(SL_orderNum,false)
											'else
											'	set shopperRecord = sitelink.SHIPTONAME(ThisOrderNumber,true)
											'end if
										   set shopperRecord = sitelink.SHIPTONAME(ThisOrderNumber,IsMOMOrdNum)
										  %>
										  <tr>
										  	<td class="plaintextbold"><b>Shipping Address: </b></td>
										  	<td class="plaintext"><%=trim(shopperRecord.firstname)%>&nbsp;&nbsp;<%=trim(shopperRecord.lastname)%></td> 
										 </tr>
										 <% if len(trim(shopperRecord.company)) > 0 then%>
											  <tr><td>&nbsp;</td><td class="plaintext"><%=trim(shopperRecord.company)%></td></tr>
										  <%end if%>
										 <% if len(trim(shopperRecord.addr)) > 0 then%>
											  <tr><td>&nbsp;</td><td class="plaintext"><%=trim(shopperRecord.addr)%></td></tr>
										  <%end if%>
					  					  <% if len(trim(shopperRecord.addr2)) > 0 then%>
										  <tr><td>&nbsp;</td><td class="plaintext"><%=trim(shopperRecord.addr2)%></td></tr>
										  <%end if%>
										  <tr><td>&nbsp;</td><td class="plaintext"><%=trim(shopperRecord.city)%></td></tr>
										  <tr><td>&nbsp;</td><td class="plaintext"><%=trim(shopperRecord.state)%>&nbsp;&nbsp;<%=trim(shopperRecord.zipcode)%>&nbsp;&nbsp;<%=sitelink.countryname(shopperRecord.country)%></td></tr>					  
										  
										  <%
										  set shopperRecord = nothing
										  end if%>
								  </table>
								 				
								</td>
								
								<td width="50%" valign="top">
								   <table width="100%" border="0" cellpadding="3" cellspacing="0" >
								        <%
	                                        SL_PrintTaxAmount = sitelink.ORD_TAX(ThisOrderNumber,IsMOMOrdNum)
	                                        SL_PrintShipAmount = sitelink.ORD_SHIPPING(ThisOrderNumber,IsMOMOrdNum)
                                    	    
                                    	    
	                                        if giftcertAmount > 0 then
	                                            SL_PrintOrderAmount= basket_total + SL_PrintTaxAmount + SL_PrintShipAmount
	                                        else
	                                            SL_PrintOrderAmount = sitelink.ORD_TOTAL(ThisOrderNumber,IsMOMOrdNum)
	                                        end if	
	                                        
	                                        SL_PrintGrandTotal= SL_PrintOrderAmount
	                                        
	                                        if Points_onOrderTotal= true then
	                                            SL_PrintGrandTotal = cdbl(SL_PrintOrderAmount-points_amt)
	                                            SL_PrintGrandTotal = round(SL_PrintGrandTotal,4)	 
	                                        end if
	                                        
	                                        
	                                        if basket_has_giftcert=true then 
	                                            SL_PrintGrandTotal = cdbl(SL_PrintGrandTotal-giftcertAmount)
	                                            SL_PrintGrandTotal = round(SL_PrintGrandTotal,4)	                                       
	                                        end if 

                                            if GiftCardAmount > 0 then
                                                SL_PrintGrandTotal = cdbl(SL_PrintGrandTotal-GiftCardAmount)
                                                SL_PrintGrandTotal = round(SL_PrintGrandTotal,4)	
                                            end if
	                                        
	                                    %>
									   <tr><td class="plaintextbold" >Subtotal: </td><td class="plaintext"><%=formatcurrency(cdbl(basket_total))%><%'=FORMATCURRENCY(sitelink.ORD_SUBTOTAL(ThisOrderNumber,IsMOMOrdNum)) %></td></tr>
									   <tr><td class="plaintextbold">Tax: </td><td class="plaintext"><%= FORMATCURRENCY(sitelink.ORD_TAX(ThisOrderNumber,IsMOMOrdNum)) %></td></tr>
									   <tr><td class="plaintextbold">Shipping & handling</td><td class="plaintext"><%= FORMATCURRENCY(sitelink.ORD_SHIPPING(ThisOrderNumber,IsMOMOrdNum)) %></td></tr>									   
										<%if Points_onOrderTotal= true then %>
										<tr><td class="plaintextbold"><b>Points Redeemed</b></td><td class="plaintext"><%= FORMATCURRENCY(points_amt) %></td></tr>										
										<%end if%>
										<%if basket_has_giftcert=true then %>
										<tr><td class="plaintextbold"><b>Gift Certificate Redeemed</b></td><td class="plaintext"><%= FORMATCURRENCY(giftcertAmount) %></td></tr>										
										<%end if %>
										
										<%if GiftCardAmount > 0 then  %>
										        <tr><td class="plaintextbold"><b>Gift Card Amount</b></td><td class="plaintext"><%= FORMATCURRENCY(GiftCardAmount) %></td></tr>		                                        
		                                <%end if %>
		
										<%'if Points_onOrderTotal= false and basket_has_giftcert=false  then %>
									   <tr><td class="plaintextbold"><b>Grand Total</b></td><td class="plaintext"><b><%=formatcurrency(SL_PrintGrandTotal)%><%'= FORMATCURRENCY(sitelink.ORD_TOTAL(ThisOrderNumber,IsMOMOrdNum)) %></b></td></tr>
									   <%'end if%>
									   <tr><td colspan="2">&nbsp;</td></tr>
									   
									 <!-- tracking # -->
									   <% if IsMOMOrdNum = true then %>
									   <TR><TD  align="center" colspan="2"  class="THHeader" ><b>Tracking Information</b></TD></TR>
									   <%
										
									   extrafld ="(cust_bchrg+cust_ochrg) as charges"
									   xmlstring = sitelink.GET_TRACKINGINFO(ThisOrderNumber,extrafld)
									   objDoc.loadxml(xmlstring)
									   
										
									   set SL_Tracking = objDoc.selectNodes("//gti")								
							 			for x=0 to SL_Tracking.length-1 
							 				SL_shiplist= SL_Tracking.item(x).selectSingleNode("shiplist").text
											SL_track   = SL_Tracking.item(x).selectSingleNode("trackingno").text
											SL_charges   = SL_Tracking.item(x).selectSingleNode("charges").text
											
											'SL_track ="aa"
											'if len(trim(SL_track)) = 0 then
											'	SL_track = "No Tracking Number Available"
											'end if 
											
											if (x mod 2) = 0 then
											   class_to_use = "tdRow1Color"
											else
												class_to_use = "tdRow2Color"
											end if 
											
											extrafld1 = ""
											xmlstring = sitelink.GET_SHIPLISTINFO(SL_shiplist,extrafld1)
											objDoc.loadxml(xmlstring)
											
											ca_type = ""
											set SL_ShipMethods = objDoc.selectNodes("//gsi")
											if SL_ShipMethods.length > 0 then
												ca_type = SL_ShipMethods.item(0).selectSingleNode("ca_type").text
											end if 
											set SL_ShipMethods = nothing											
											 
											 												
										%>										
											<tr><td class="<%=class_to_use%>" colspan="2">												
																							
											<%if ca_type = "UPS" then %>												
												<a href= "http://wwwapps.ups.com/etracking/tracking.cgi?TypeOfInquiryNumber=T&InquiryNumber1=<%=SL_track%>" target="blank">
												<%=SL_track%>
												</a>
											<%elseif  ca_type = "FEX" then %>
												<!--a href= "http://www.fedex.com/cgi-bin/tracking?action=track&language=english&cntry_code=us&initial=x&tracknumbers=<%=SL_track%>" target="blank"-->
												<a href= "http://www.fedex.com/Tracking?tracknumbers=<%=SL_track%>" target="blank">
												<%=SL_track%>
												</a>
											<%elseif  ca_type = "USP" then %>
											<a href= "https://tools.usps.com/go/TrackConfirmAction.action?tLabels=<%=SL_track%>" target="blank">
												<%=SL_track%>
												</a>
												
											<%else%>
												<span class="plaintext"><%=SL_track%></span>
											<%end if%>
											</td></tr>
											<tr><td class="<%=class_to_use%>" colspan="2"><span class="plaintext">Box <%=x+1%>&nbsp;Shipping Charges : <%=formatcurrency(SL_charges)%></span></td></tr>											
									   <%next
									   set SL_Tracking = nothing									   
									   %>
									   <%end if%>
									  
									   <!-- tracking # -->
								   </table>				
								</td>
							</tr>
						 </table>
				  		 </td></tr>
						  <%
							if SL_orderNum > 0 then
							 ordercomments = sitelink.DISPLAYNOTES(SL_orderNum)
							  if len(trim(ordercomments)) > 0 then %>
								  <tr><th colspan="2" align="left"  class="THHeader" >Special Instructions</th></tr>
								  <tr>
									  <td class="plaintext" colspan="2" ><%=ordercomments%></td>
								  </tr>  
							  <%end if 
								end if	  
							%>		 
						 </table>
					 </td>
					 </tr>
					

				</table>
	
	
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