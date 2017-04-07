
<%

	basket_total = 0 
	

%>


<!-- billing & shipping info -->


<table width="99%" cellpadding="2" cellspacing="0" border="0" bordercolor="#000000">
   
   <tr><td align="LEFT"  class="plaintext"  colspan="2"><b> Billing Address</b></td></tr>
   <tr><td width="15">&nbsp;</td>
   		<td class="plaintext">
			<%=session("firstname") %>&nbsp;<%=session("lastname") %>
			<%if len(trim(session("company"))) >0 then response.write("<br>"&session("company")) end if %>
			<%if len(trim(session("address1"))) >0 then response.write("<br>"&session("address1")) end if %>
			<%if len(trim(session("address2"))) >0 then response.write(", "&session("address2")) end if %>
			<%if len(trim(session("address3"))) >0 then response.write(", "&session("address3")) end if %>
			<br>
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
<!-- basket info -->

<%
		
		default_ship_title = trim(sitelink.dataforstatusnew("SHIPPINGMETHOD",session("previousorder"),session("previousshopperid"),"EMAIL",false))
				
	
%>
	<br>
	<table width="99%" border="1" cellspacing="0" cellpadding="3" bordercolor="#000000">
	
	<tr> 
       <td align="CENTER" width="9%" class="plaintext"><b>Qty</b></td>
       <td align="CENTER" width="39%" class="plaintext"><b>Description</b></td>
       <td align="CENTER" width="4%" class="plaintext"><b>Total</b></td>
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
				SL_Basket_Shipto	= SL_Basket.item(x).selectSingleNode("ship_to").text
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
				
				 if SL_Basket_Shipto > 0 then
				 	fname				= SL_Basket.item(x).selectSingleNode("firstname").text
				    lname				= SL_Basket.item(x).selectSingleNode("lastname").text
					company			    = SL_Basket.item(x).selectSingleNode("company").text
					addr				= SL_Basket.item(x).selectSingleNode("addr").text
					addr2				= SL_Basket.item(x).selectSingleNode("addr2").text
					addr3				= SL_Basket.item(x).selectSingleNode("addr3").text
					city				= SL_Basket.item(x).selectSingleNode("city").text
					vstate				= SL_Basket.item(x).selectSingleNode("state").text
					zipcode			    = SL_Basket.item(x).selectSingleNode("zipcode").text
					country			    = SL_Basket.item(x).selectSingleNode("country").text
					county			    = SL_Basket.item(x).selectSingleNode("county").text
				end if


				
				SL_ShipVia 			=SL_Basket.item(x).selectSingleNode("ship_via").text
				SL_ShippingTitle  = ""
				if len(trim(SL_ShipVia)) > 0 then
					SL_ShippingTitle = SL_Basket.item(x).selectSingleNode("ca_title").text
				else
					SL_ShippingTitle = default_ship_title
				end if 
				
				
				SL_ShipWhen =SL_Basket.item(x).selectSingleNode("ship_when").text
				
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
				<td valign="top" width="60%" class="plaintext">
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
				<b>Item: </b><%=SL_BasketNumber%> &nbsp; <b>Price: </b><%=formatcurrency(SL_BasketUnitPrice)%> &nbsp;
				<%if SL_PointsRdm="true" then%>
					<br><%=SL_Points_ned%> points were redeemed to purchase this product
				<%end if%>
				<%if SL_BasketDiscount > 0 then   response.write("<b>Discount: </b>"+ SL_BasketDiscount + "%") end if %>
				<%if len(trim(SL_BasketCustInfo)) > 0 then response.write("<br>" + SL_BasketCustInfo) end if %>
				
				<%if SL_Number<>"" then %>
				<table width="0" cellpadding="0" cellspacing="0" border="0">
				  <tr><td valign="top" width="25%" class="plaintextbold">Ship To: </td>
					<td valign="top" class="plaintext">
					<% if SL_Basket_Shipto > 0 then 
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

                        if len(addr3) > 0 then
							stringto_write = stringto_write + ", " + (addr3)
						end if
						
						stringto_write = stringto_write + "<br>&nbsp;" + (city)
						if country="073" and county<>"" then
						    stringto_write = stringto_write + ", " + (sitelink.Get_CountyName(county))
						else
						    stringto_write = stringto_write + ", " + (vstate)
						end if
						stringto_write = stringto_write + ", " + (zipcode)
						stringto_write = stringto_write + ", " + (sitelink.countryname(country))

						response.write(stringto_write)
						
					%>
					<%else%>
						Same as Bill to
					<%end if%>
					</td>
					
				</tr>
				<tr><td colspan="2" class="smalltextblk">
					
						<b>Ship Via: </b><%=SL_ShippingTitle%>
					<% if len(trim(SL_ShipWhen_Date)) > 0 then %>
						<br>
						<b>Ship On: </b><%=SL_ShipWhen_Date%>
					<%end if%>
					
				</td>
				</table>
				<%end if%>						
						
				</td>
				<td valign="top" width="1">&nbsp;</td>
				<td valign="top" class="plaintext">
					<%
						usethisship = SL_Basket_Shipto + 1 - 1
						set logiftmsg2=sitelink.editgiftmsg(session("previousorder"),usethisship) 
							giftmsg1=trim(logiftmsg2.note1)
							giftmsg2=trim(logiftmsg2.note2)
							giftmsg3=trim(logiftmsg2.note3)
							giftmsg4=trim(logiftmsg2.note4)
							giftmsg5=trim(logiftmsg2.note5)
							giftmsg6=trim(logiftmsg2.note6)
									
							if len(giftmsg1+giftmsg2+giftmsg3+giftmsg4+giftmsg5+giftmsg6) > 0 then

							response.write("<b>Gift Message</b>")
							if len(giftmsg1) > 0 then
								response.write("<br>")
								response.write(giftmsg1)
							end if
							if len(giftmsg2) > 0 then
								response.write("<br>")
								response.write(giftmsg2)
							end if
							if len(giftmsg3) > 0 then
								response.write("<br>")
								response.write(giftmsg3)
							end if
							if len(giftmsg4) > 0 then
								response.write("<br>")
								response.write(giftmsg4)
							end if
							if len(giftmsg5) > 0 then
								response.write("<br>")
								response.write(giftmsg5)
							end if
							if len(giftmsg6) > 0 then
								response.write("<br>")
								response.write(giftmsg6)
							end if
							
							end if			
									
							  set logiftmsg2 = nothing 

						%>
					
					</td>
				
				
				</tr>
				</table>
			
			</td>
			<td class="plaintext"><%=formatcurrency(SL_BasketExtPrice)%></td>
		</tr>

	 <%next%>
	
	</table>
	
	<%
	    'SL_PrintTaxAmount = sitelink.ORD_TAX(session("previousorder"),false)
	    'SL_PrintShipAmount = sitelink.ORD_SHIPPING(session("previousorder"),false)
	    
	    if giftcertAmount > 0 then
	        SL_PrintOrderAmount= basket_total + SL_PrintTaxAmount + SL_PrintShipAmount
	    else
	        'SL_PrintOrderAmount = sitelink.ORD_TOTAL(session("previousorder"),false)
			'SL_PrintOrderAmount = SL_ORD_TOTAL
	    end if
	    
	    
	
	 %>


          
<!-- sub total/shipping/ order total info-->
            <table width="98%">
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
                  <b>(National,State and Local taxes)&nbsp;&nbsp;Tax:</b> </td>
                <td align="right"  width="11%" class="plaintext">  
                  <b><%=formatcurrency(SL_PrintTaxAmount)%><%' = FORMATCURRENCY(sitelink.ORD_TAX(session("previousorder"),false)) %></b>
                  </td>
              </tr>
			  
			                <tr> 
                <td  colspan="2" class="plaintext">
					<table width="100%" cellpadding="3" cellspacing="0" border="0">
					<tr><td align="right" class="plaintext">&nbsp;<b>Shipping:</b></td></tr>
					
					</table>                  
				</td>
                <td align="RIGHT"  width="11%" class="plaintext">  <b>
                   <%=formatcurrency(SL_PrintShipAmount) %>
                  <%' = FORMATCURRENCY(sitelink.ORD_SHIPPING(session("previousorder"),false)) %>
				  
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
                  <%'= FORMATCURRENCY(sitelink.ORD_TOTAL(session("previousorder"),false)) %>
				  </b>
                  </td>
              </tr>

 </table>


