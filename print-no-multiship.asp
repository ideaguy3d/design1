
<%

	basket_total = 0 	
	
%>


<!-- billing & shipping info -->
<script type="text/javascript">
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
</script>



<table width="100%" cellpadding="0" cellspacing="0" border="0">
   <tr><td width="50%" valign="top" class="split-col">
   <table width="100%">
   <tr><td align="LEFT"  class="THHeader"  colspan="2">Billing Address</td></tr>
   <tr><td width="15">&nbsp;</td>
   		<td class="plaintext" align="left">
			<%=session("firstname") %>&nbsp;<%=session("lastname") %>
			<%if len(trim(session("company"))) >0 then response.write("<br>"&session("company")) end if %>
			<%if len(trim(session("address1"))) >0 then response.write("<br>"&session("address1")) end if %>
			<%if len(trim(session("address2"))) >0 then response.write(", "&session("address2")) end if %>
			<%if len(trim(session("address3"))) >0 then response.write(", "&session("address3")) end if %>
			<br>
			<%=session("city")%>,&nbsp;
			<%if session("country")="073" and len(trim(session("bcounty"))) > 0 then %>
			    <%=sitelink.Get_CountyName(session("bcounty")) %>
			<%else %>
			    <%=session("state")%>
			<%end if %>
			&nbsp;<%=session("zipcode")%>
			&nbsp;
			<% =sitelink.countryname(session("country")) %>  
			<br>
			<%=session("email")%> 
   	    </td>
	</tr>
	</table>
	</td>
	
	<td width="50%" valign="top" class="split-col">
	<table width="100%">
		<tr><td align="LEFT"  class="THHeader"  colspan="2">Shipping Address</td></tr>
		<%if session("billtocopy")=1 then%>
			<tr><td width="15">&nbsp;</td>
				<td class="plaintext" align="left">Same as Billing</td></tr>
		<%else%>
		<tr><td width="15">&nbsp;</td>
   		<td class="plaintext" align="left">
			<%=session("sfirstname") %>&nbsp;<%=session("slastname") %>
			<%if len(trim(session("scompany"))) >0 then response.write("<br>"&session("scompany")) end if %>
			<%if len(trim(session("saddress1"))) >0 then response.write("<br>"&session("saddress1")) end if %>
			<%if len(trim(session("saddress2"))) >0 then response.write(", "&session("saddress2")) end if %>
			<%if len(trim(session("saddress3"))) >0 then response.write(", "&session("saddress3")) end if %>
			<br>
			<%=session("scity")%>,&nbsp;
			<%if session("scountry")="073" and len(trim(session("scounty"))) > 0 then %>
			    <%=sitelink.Get_CountyName(session("scounty")) %>
			<%else %>
			    <%=session("sstate")%>
			<%end if %>			
			&nbsp;<%=session("szipcode")%>
			&nbsp;
			<% =sitelink.countryname(session("scountry")) %>  
			<%if len(trim(session("semail"))) >0 then response.write("<br>"&session("semail")) end if %>
			
   	    </td>
	</tr>
		
		<%end if%>
	</table>
	</td>
	
</table>
<!-- basket info -->


<%
		
		default_ship_title = sitelink.dataforstatusnew("SHIPPINGMETHOD",session("ordernumber"),session("shopperid"),"EMAIL",false)
%>
	<br>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	
	<tr> 
       <td align="left" width="5%" class="THHeader">Qty</td>
       
       <td align="left" width="80%" class="THHeader" style="padding-left:0;">Description</td>
       <td align="right" width="10%" class="THHeader">Total&nbsp;</td>
       
     </tr>
	 <% 
	  Show_prodRestMsg = false
	  if len(trim(session("ProdRestHoldList"))) > 0 then	   
	    objDoc.loadxml(session("ProdRestHoldList"))
	    set ProdRestrict = objDoc.selectNodes("//results") 
	    Show_prodRestMsg = true
	  end if
			 
	 for x=0 to SL_Basket.length-1
	 			SL_Number			= SL_Basket.item(x).selectSingleNode("number2").text
				SL_Variant			= SL_Basket.item(x).selectSingleNode("variant").text
				SL_BasketNumber 	= SL_Number +" "+ SL_Variant
				SL_BasketQuantity	= SL_Basket.item(x).selectSingleNode("quanto").text
				SL_BasketDesc		= SL_Basket.item(x).selectSingleNode("inetsdesc").text
				SL_BasketDesc2		= SL_Basket.item(x).selectSingleNode("desc2").text
				SL_BasketUnitPrice	= SL_Basket.item(x).selectSingleNode("it_unlist").text
				SL_BasketExtPrice	= SL_Basket.item(x).selectSingleNode("extendp").text
				'SL_BasketCmpPrice	= SL_Basket.item(x).selectSingleNode("inetcprice").text
				SL_BasketDiscount	= SL_Basket.item(x).selectSingleNode("discount").text
				SL_BasketCustInfo	= SL_Basket.item(x).selectSingleNode("custominfo").text
				
				SL_Points_ned       = SL_Basket.item(x).selectSingleNode("points_ned").text
				SL_PointsRdm		=SL_Basket.item(x).selectSingleNode("pts_rdeemd").text

				
				SL_units 			=SL_Basket.item(x).selectSingleNode("units").text
				SL_drop 			=SL_Basket.item(x).selectSingleNode("dropship").text
				SL_construct 		=SL_Basket.item(x).selectSingleNode("construct").text
				
				'SL_FullNumber 		=SL_Basket.item(x).selectSingleNode("item").text
				SL_item_id   		=SL_Basket.item(x).selectSingleNode("record").text
				
				SL_Promo_Ind  =SL_Basket.item(x).selectSingleNode("basketdesc").text
				 SL_PROMO_ITEM= false
				 if SL_Promo_Ind ="PROMO" then
				 	SL_PROMO_ITEM = true
				 end if
				
				SL_BasketQty = replace(SL_BasketQuantity,".00","")
				
				SL_ShipVia 			=SL_Basket.item(x).selectSingleNode("ship_via").text
				SL_ShippingTitle  = ""

				if SL_ShipVia<>"" then
					SL_ShippingTitle = SL_Basket.item(x).selectSingleNode("ca_title").text
				else
					SL_ShippingTitle = ""
				end if 
				
				SL_ShipWhen 		=SL_Basket.item(x).selectSingleNode("ship_when").text
				
				if len(trim(SL_ShipWhen)) > 0 then
					SL_ShipWhen=replace(SL_ShipWhen,"T"," ")
					SL_ShipWhen_Date = FormatDateTime(SL_ShipWhen, 0)
				else
					SL_ShipWhen_Date = "" 
				end if 
			
				basket_total		= basket_total + SL_BasketExtPrice
				SL_Points_ned	 	= SL_Points_ned * SL_BasketQty

		%>
		<tr>
			<td class="plaintext" align="center" valign="top">
			<%if SL_PROMO_ITEM=false then %>
			<%=SL_BasketQty%>
			<%else%>
			&nbsp;
			<%end if%>
			</td>
           
			<td valign="top" class="plaintextbold" align="left">
				
				<%=SL_BasketDesc%>
				<%	
				if Show_prodRestMsg = true then				    
				    for z=0 to ProdRestrict.length-1
				        ProdRestrictNumber = ProdRestrict.item(z).selectSingleNode("number").text
				        ProdRestrictDays   = ProdRestrict.item(z).selectSingleNode("ldays").text
				        ProdRestrictEndDate  = ProdRestrict.item(z).selectSingleNode("lenddate").text
				        ProdRestictItemId  = ProdRestrict.item(z).selectSingleNode("item_id").text
				        ProdRestictShipWhen  = ProdRestrict.item(z).selectSingleNode("ship_when").text
				        'response.Write(ProdRestictShipWhen & "--" &ProdRestrictEndDate &"--"&isdate(ProdRestictShipWhen) &"--->"&ProdRestrictDays)
				        if SL_item_id=ProdRestictItemId then				            
				            if isdate(ProdRestictShipWhen)= true then				                
				                diff_days = cdate(ProdRestrictEndDate)-cdate(ProdRestictShipWhen)					               
				                'if diff_days < cint(ProdRestrictDays) and diff_days > 0 then
				                if diff_days => 0 then
				                    'ProdRestrictEndDate= ProdRestrictEndDate
				                    ProdRestrictEndDate = cstr(DateAdd("D",ProdRestrictDays,ProdRestrictEndDate)) 				                    
				                else
				                    'ProdRestrictEndDate = cstr(DateAdd("D",ProdRestrictDays,ProdRestrictEndDate)) 	
				                    ProdRestrictEndDate= ProdRestrictEndDate				                    
				                end if
				             else				               				        		            
    				            ProdRestrictEndDate = cstr(DateAdd("D",ProdRestrictDays,ProdRestrictEndDate)) 
    				         end if
    				         
    				         
				            SL_ShipWhen_Date = cdate(ProdRestrictEndDate)				            
				
				end if
				next
				end if %>
			  <%
				if SHOW_IN_STOCK = 1 and SHOW_DUE_DATE = 0  then
					if  SL_units >0 or  SL_drop="true" or SL_construct= "true" then
						Response.Write("&nbsp;&nbsp;(In Stock)")
					else
						Response.Write("&nbsp;&nbsp;(Out of Stock)")
					end if
				end if 
				
				%>			
						
				<% if len(trim(SL_BasketDesc2)) > 0 then response.write("<br>" +SL_BasketDesc2) end if %>
				<br>
			  <%if SL_PROMO_ITEM=false then %><strong>Item Number: </strong><%=SL_BasketNumber%> <br /><br /> <strong>Price: </strong><%=formatcurrency(SL_BasketUnitPrice)%> &nbsp;<%end if%>				
				<%if SL_BasketDiscount > 0 then   response.write("<strong>Discount: </strong>"+ SL_BasketDiscount + "%") end if %>
				<%if len(trim(SL_BasketCustInfo)) > 0 then response.write("<br>" + SL_BasketCustInfo) end if %>
				<%if SL_PointsRdm="true" then%>
					<br><%=SL_Points_ned%> points were redeemed to purchase this product
				<%end if%>
				
			  <%if SL_Number<>"" then %>						
				
				<% if len(trim(SL_ShipWhen_Date)) > 0 then %>
					<br>
					<font color="red"><strong>Ship On: </strong><%=SL_ShipWhen_Date%></font>
				<%end if%>
				<%end if%>
		  <% if SL_ShipVia<>"" then %>
					<br>
					<strong>Ship Via: </strong><%=SL_ShippingTitle%>
			  <%end if%>
				<br /><br />
			
			</td>
			<td class="plaintext" ALIGN="right" valign="top"><%=formatcurrency(SL_BasketExtPrice)%>&nbsp;&nbsp;</td>
            
		</tr>

	 <%next%>
	<% set ProdRestrict = nothing %>
</table>
	
	<%
	    'SL_PrintTaxAmount = sitelink.ORD_TAX(session("ordernumber"),false)
	    'SL_PrintShipAmount = sitelink.ORD_SHIPPING(session("ordernumber"),false)
	    
	    if giftcertAmount > 0 then
	        SL_PrintOrderAmount= basket_total + SL_PrintTaxAmount + SL_PrintShipAmount
	    else
	        'SL_PrintOrderAmount = sitelink.ORD_TOTAL(session("ordernumber"),false)
			'SL_PrintOrderAmount = SL_ORD_TOTAL
	    end if
	    
	    
	
	 %>

         
<!-- sub total/shipping/ order total info-->
            <table width="100%" border="0">
            	<tr>
                	<td colspan="4" class="THHeader">&nbsp;</td>
                </tr>
              <tr> 
                <td align="right" colspan="2" class="plaintext"> 
                  <strong>Subtotal:</strong> </td>
                <td align="right"  width="11%" class="plaintext">                   
				  <strong><%=formatcurrency(cdbl(basket_total))%>
				  <%' = FORMATCURRENCY(sitelink.ORD_SUBTOTAL(session("ordernumber"),false)) %>
				  </strong>
                  </td>
              </tr>
			  
			 
			 
              <tr> 
                <td align="right" colspan="2"  class="plaintext"> 
                  (National,State and Local taxes)&nbsp;&nbsp;<strong>Tax:</strong> </td>
                <td align="right"  width="11%" class="plaintext">  
                  <strong>
                  <%=formatcurrency(SL_PrintTaxAmount) %>
                  <%' = FORMATCURRENCY(sitelink.ORD_TAX(session("ordernumber"),false)) %></strong>
                  </td>
              </tr>
              <tr> 
                <td class="plaintextbold" align="right">
								
					<%if show_shipping_option=true then %>
					<table width="100%" cellpadding="3" cellspacing="0" border="0">
					<tr><td align="right" class="plaintext">(<%=default_ship_title%>)</td></tr>
					
					<tr><td align="right"><a href="#" onclick="MM_openBrWindow('PreviewShipping.asp','','width=450,height=400')">Select Another Shipping Method</a></td></tr>
					</table>
					<%end if %>
					</td>
					
					<td width="10" align="right" valign="top" class="plaintext">&nbsp;<strong>Shipping:</strong></td>					   					    
					
                	<td align="RIGHT"  width="11%"  valign="top" class="plaintext">  <strong>
                    <%=formatcurrency(SL_PrintShipAmount) %>
                  	<%' = FORMATCURRENCY(sitelink.ORD_SHIPPING(session("ordernumber"),false)) %>
				  
				  	</strong>
                  	</td>
              </tr>
              <tr> 
                <td colspan="3"> 
                  <hr size="1">
                </td>
              </tr>
              <tr> 
                <td align="right" colspan="2" class="plaintext"> <strong>TOTAL: </strong>
                  </td>
                <td align="RIGHT"  width="11%" class="plaintext">  <strong>
                    <%=formatcurrency(SL_PrintOrderAmount) %>
                <%' = FORMATCURRENCY(sitelink.ORD_TOTAL(session("ordernumber"),false)) %>
				  </strong>
                  </td>
              </tr>

 </table>
<%
	'Set objDoc = nothing
	'set SL_Ship = nothing	
	'set SL_Basket = nothing	
%>


