
<%

	basket_total = 0.0
	
%>


<!-- billing & shipping info -->


<table width="99%" cellpadding="2" cellspacing="0" border="1" bordercolor="#000000">
   <tr><td width="50%" valign="top">
   <table width="100%">
   <tr><td align="LEFT"  class="plaintext"  colspan="2"><b> Billing Address</b></td></tr>
   <tr><td width="15">&nbsp;</td>
   		<td class="plaintext">
			<%=session("firstname") %>&nbsp;<%=session("lastname") %>
			<%if len(trim(session("company"))) >0 then response.write("<br>"&session("company")) end if %>
			<%if len(trim(session("address1"))) >0 then response.write("<br>"&session("address1")) end if %>
			<%if len(trim(session("address2"))) >0 then response.write(", "&session("address2")) end if %>
			<%if len(trim(session("address3"))) >0 then response.write("<br>"&session("address3")) end if %>
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
	
	<td width="50%" valign="top">
	<table width="100%">
		<tr><td align="LEFT"  class="plaintext"  colspan="2"><b> Shipping Address</b></td></tr>
		<%if session("previousbilltocopy")=1 then%>
			<tr><td width="15">&nbsp;</td>
				<td class="plaintext">Same as Billing</td></tr>
		<%else%>
		<tr><td width="15">&nbsp;</td>
   		<td class="plaintext">
			<%=session("sfirstname") %>&nbsp;<%=session("slastname") %>
			<%if len(trim(session("scompany"))) >0 then response.write("<br>"&session("scompany")) end if %>
			<%if len(trim(session("saddress1"))) >0 then response.write("<br>"&session("saddress1")) end if %>
			<%if len(trim(session("saddress2"))) >0 then response.write(", "&session("saddress2")) end if %>
			<%if len(trim(session("saddress3"))) >0 then response.write("<br>"&session("saddress3")) end if %>
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
		
		default_ship_title = trim(sitelink.dataforstatusnew("SHIPPINGMETHOD",session("previousorder"),session("previousshopperid"),"EMAIL",false))
%>
	<br>
	<table width="99%" border="1" cellspacing="0" cellpadding="3" bordercolor="#000000">
	
	<tr> 
       <td align="CENTER" width="10%" class="plaintext"><b>Qty</b></td>
       <td align="CENTER" width="80%" class="plaintext"><b>Description</b></td>
       <td align="CENTER" width="10%" class="plaintext"><b>Total</b></td>
     </tr>
	 <% for x=0 to SL_Basket.length-1
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
				SL_PointsRdm		= SL_Basket.item(x).selectSingleNode("pts_rdeemd").text			
				
				SL_units 			=SL_Basket.item(x).selectSingleNode("units").text
				SL_drop 			=SL_Basket.item(x).selectSingleNode("dropship").text
				SL_construct 		=SL_Basket.item(x).selectSingleNode("construct").text
				
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
			<td class="plaintext" align="center">
			<%if SL_PROMO_ITEM=false then %>
			<%=SL_BasketQty%>
			<%else%>
			&nbsp;
			<%end if%>
			</td>
			<td valign="top">
				<table width="100%" cellpadding="3" cellspacing="0" border="0">
				<tr>
				<td valign="top" class="smalltextblk">
				<%=SL_BasketDesc%>
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
				<%if SL_PROMO_ITEM=false then %><b>Item: </b><%=SL_BasketNumber%> &nbsp; <b>Price: </b><%=formatcurrency(SL_BasketUnitPrice)%> &nbsp;<%end if%>				
				<%if SL_BasketDiscount > 0 then   response.write("<b>Discount: </b>"+ SL_BasketDiscount + "%") end if %>
				<%if len(trim(SL_BasketCustInfo)) > 0 then response.write("<br>" + SL_BasketCustInfo) end if %>
				<%if SL_PointsRdm="true" then%>
					<br><%=SL_Points_ned%> points were redeemed to purchase this product
				<%end if%>
				<%if SL_Number<>"" then %>						
				
				<% if len(trim(SL_ShipWhen_Date)) > 0 then %>
					<br>
					<b>Ship On: </b><%=SL_ShipWhen_Date%>
				<%end if%>
				<%end if%>
				<% if SL_ShipVia<>"" then %>
					<br>
					<b>Ship Via: </b><%=SL_ShippingTitle%>
				<%end if%>
				</td></tr>
				</table>
			
			</td>
			<td class="plaintext" ALIGN="right"><%=formatcurrency(SL_BasketExtPrice)%>&nbsp;&nbsp;</td>
		</tr>

	 <%next%>
	
	</table>

         
     <%
        'SL_PrintTaxAmount=0.0
        'SL_PrintShipAmount=0.0
	    'SL_PrintTaxAmount = sitelink.ORD_TAX(session("previousorder"),false)
	    'SL_PrintShipAmount = sitelink.ORD_SHIPPING(session("previousorder"),false)
	    
	    if giftcertAmount > 0 then
	        SL_PrintOrderAmount= cdbl(basket_total + SL_PrintTaxAmount + SL_PrintShipAmount)	        
	    else
	        'SL_PrintOrderAmount = sitelink.ORD_TOTAL(session("previousorder"),false)
			'SL_PrintOrderAmount = SL_ORD_TOTAL
	    end if
	
	 %>
<!-- sub total/shipping/ order total info-->
            <table width="98%"  border="0">
              <tr> 
                <td align="right" colspan="2" class="plaintext"> 
                  <b>Subtotal:</b> </td>
                <td align="right"  width="11%" class="plaintext">                   
				  <b><%=formatcurrency(cdbl(basket_total))%>
				  <%' = FORMATCURRENCY(sitelink.ORD_SUBTOTAL(session("previousorder"),false)) %>
				  </b>
                  </td>
              </tr>
			  
			 
			 
              <tr> 
                <td align="right" colspan="2"  class="plaintext"> 
                  (National,State and Local taxes)&nbsp;&nbsp;<b>Tax:</b> </td>
                <td align="right"  width="11%" class="plaintext">  
                  <b><%=formatcurrency(SL_PrintTaxAmount) %></b>
                  </td>
              </tr>
              <tr> 
                <td  colspan="2" class="plaintext">
					<table width="100%" cellpadding="3" cellspacing="0" border="0">
					<tr><td align="right" class="plaintext">(<%=default_ship_title%>)&nbsp;<b>Shipping:</b></td></tr>
					
					</table>                  
				</td>
                <td align="RIGHT"  width="11%" class="plaintext">  <b>
                  <%=formatcurrency(SL_PrintShipAmount) %>
				  
				  </b>
                  </td>
              </tr>
              <tr> 
                <td colspan="3"> 
                  <hr size="1">
                </td>
              </tr>
              <tr> 
                <td align="right" colspan="2" class="plaintext"> <b>TOTAL: </b>
                  </td>
                <td align="RIGHT"  width="11%" class="plaintext">  <b>
                  <%=formatcurrency(SL_PrintOrderAmount) %>
				  </b>
                  </td>
              </tr>

 </table>
<%
	'Set objDoc = nothing
	'set SL_Ship = nothing	
	'set SL_Basket = nothing	
%>


