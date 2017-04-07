

<%if SHOW_BESTSELLER=1 then %>

<%
            NumberOfBestSellers = BEST_SELLERS_COUNT
			extrafield="LEFT(Justfname(inetthumb)+Space(50),50) As inetthumb, LEFT(Justfname(inetimage)+Space(50),50) As inetimage, size_color, giftcard, price1"
			xmlstring =sitelink.READBESTSELLER(NumberOfBestSellers,extrafield)
			objDoc.loadxml(xmlstring)
			

			set SL_BestSeller = objDoc.selectNodes("//bestsell")
			if SL_BestSeller.length > 0 then							  

		        all_prods=""
		        for x=0 to SL_BestSeller.length-1
			        SLSLProdNumberber = SL_BestSeller.item(x).selectSingleNode("number").text
			        all_prods =all_prods+","+cstr(SLSLProdNumberber)
		        next
		        all_prods =all_prods+","
        	
		        xmlstring = sitelink.GET_ALL_SPECIALPRICES(cstr(all_prods),session("ordernumber"),session("shopperid"),cstr(""))
		        
		        objDoc.loadxml(xmlstring)
			%>
			
<div id="rightsidebar" style="width:<%=SIDE_NAV_WIDTH%>px;"> 
	<div class="sidenavheader sidenavTxt">Best Sellers</div>
    <div class="best-seller-wrap">			
			 <table width="100%"  cellpadding="0" cellspacing="0" border="0">
			 <tr><td valign="top">
							<table width="100%" cellpadding="3" cellspacing="0" border="0">
								<%
								  for x = 0 to SL_BestSeller.length-1
									SLProdNumber    = SL_BestSeller.item(x).selectSingleNode("number").text
									SLProductTitle = SL_BestSeller.item(x).selectSingleNode("inetsdesc").text
									SL_thumbnail  = SL_BestSeller.item(x).selectSingleNode("inetthumb").text
									SL_Fullimage  = SL_BestSeller.item(x).selectSingleNode("inetimage").text
									
									SL_Size_color  = SL_BestSeller.item(x).selectSingleNode("size_color").text
									SL_MinPrice    = 0 'SL_BestSeller.item(x).selectSingleNode("minprice").text
									SL_MaxPrice    = 0 'SL_BestSeller.item(x).selectSingleNode("maxprice").text
									SL_GiftCard    = SL_BestSeller.item(x).selectSingleNode("giftcard").text
									SL_price       = SL_BestSeller.item(x).selectSingleNode("price1").text
									
									if SL_thumbnail="" then
			                            SL_thumbnail = SL_Fullimage
		                            end if
		                            
		                            imagename="images/noimage.gif"
		                            if len(SL_thumbnail) > 0 then
			                             StrFileName = "images/"+ SL_thumbnail  
			                             StrPhysicalPath = Server.MapPath(StrFileName)
		                                 set objFileName = CreateObject("Scripting.FileSystemObject")	
				                            if objFileName.FileExists(StrPhysicalPath) then
		  			                            imagename=StrFileName
				                            end if 
			                             set objFileName = nothing
	                                end if 

		                             target_in_SpecialPrice = "//gsp[number='"+cstr(SLProdNumber)+"']"
		                             has_special_Price= false
		                             has_Onespecial_Price=false
		                             set SL_prodSpecial = objDoc.selectNodes(target_in_SpecialPrice) 
		 	                            if SL_prodSpecial.length > 0 then
			 	                            has_special_Price=true
                            				
				                            SL_SP_qty = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("qty").text
				                            SL_SP_ordertotal = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("total_dol").text
				                            SL_SP_qty = replace(SL_SP_qty,".00","")
				                            if SL_SP_qty= 1 and (cdbl(session("SL_BasketSubTotal")) >= cdbl(SL_SP_ordertotal) )then
					                            SL_SP_price = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("price").text
					                            SL_SP_discount = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("discount").text

					                            SL_SP_CostMeth = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("costmethod").text
					                            SL_SP_CostPlusDiscount = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("costplus").text
					                            SL_SP_StockReOrdPrice = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("reordprice").text
					                            SL_SP_UnCostprice = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("uncost").text
                            					
					                            if SL_SP_CostMeth="" then
						                            unitprice = SL_price
						                            unitprice2 = SL_SP_price
						                            if unitprice2>0 then
							                            unitprice = unitprice2
						                            end if
						                            percentoff = (SL_SP_discount/100.0)
						                            discountedUnitPrice = cdbl(unitprice * (1- percentoff))									
					                            end if
                            									
					                            if SL_SP_CostMeth="P" then
							                            discountedUnitPrice= cdbl(SL_SP_StockReOrdPrice*(1+SL_SP_CostPlusDiscount/100.0))
					                            end if
                            							
					                            if SL_SP_CostMeth="U" then										
							                            discountedUnitPrice= cdbl(SL_SP_UnCostprice*(1+SL_SP_CostPlusDiscount/100.0))									
					                            end if
                            						
					                            has_Onespecial_Price=true
				                            end if
			                            end if
		                             set SL_prodSpecial = nothing
									
									if WANT_REWRITE = 1 then
										 SL_urltitle = SLProductTitle
										 SL_urltitle = url_cleanse(SL_urltitle)
										prodlink = SL_urltitle + "/productinfo/" + server.URLEncode(SLProdNumber) + "/"  
									else 
										prodlink = "prodinfo.asp?number=" + SLProdNumber
									end if

					                SHOW_MINMAXPRICE=0 '''''''''''''''''''''''''''''' this should be removed after min, max is added to the dll. Uncomment SL_MinPrice & SL_MaxPrice above. Rem extra fields & add it to dll.
								%>
								<tr><td><a href="<%=prodlink%>"><img src="<%=imagename%>" alt="<%=SL_urltitle%>" border="0" width="69"></a></td>
									<td valign="top" class="ProductPrice"><a href="<%=prodlink%>" class="producttitlelink"><%=SLProductTitle%></a><br>
			                            <%if SL_Size_color="true" and SHOW_MINMAXPRICE=1 then 
			                                ShowRangePrice = false
    		                                if cdbl(SL_MinPrice) <> cdbl(SL_MaxPrice) then
	    	                                    ShowRangePrice= true
		                                    end if

			                            %>
			                                    <% if ShowRangePrice= true then %>
                                                        <%=formatcurrency(SL_MinPrice)%> - <%=formatcurrency(SL_MaxPrice)%>
                                                <%else %>
                                                        <%=formatcurrency(SL_MinPrice)%>
                                                <%end if %>			        
	                                            <%if has_special_Price=true then %>
	                                                <img src="images/RedSaleTag.gif" border="0" alt="" align="absmiddle" hspace="6">
	                                            <%end if %>
			                            <%end if %>
                            			
			                            <%if SL_GiftCard="false" and SL_Size_color="false" then %>
			                                <%if has_special_Price=true or has_Onespecial_Price=true then %>
			                                    <%if has_Onespecial_Price=true then %>
				                                    <span class="plaintext"><s><b><%=formatcurrency(SL_price)%></b></s></span>&nbsp;<%=formatcurrency(discountedUnitPrice)%>
			                                    <%else%>
					                                <a href="<%=prodlink%>">Special Pricing Available</a>
			                                    <%end if%>
		                                    <%else%>
			                                    <%=formatcurrency(SL_price)%>
		                                    <%end if%>
		                                <%end if 'SL_GiftCard="false"%>
									</td>
								</tr>
							<% 
								next %>
								</table>

			</td></tr>
			
		</table>
    </div>
</div>
		<%
			end if
			set SL_BestSeller =nothing
		 %>
<%end if 'if SHOW_BESTSELLER %>