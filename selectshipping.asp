<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<%
'response.Write(session("ordernumber"))

Response.ExpiresAbsolute = #Feb 18,1998 13:26:26#  



if session("billtocopy")=1 then
	 vtxtzipCode = session("zipcode")
	 vtextshipcountry = session("country")
	 vtxtstate = session("state")
else
	 vtxtzipCode = session("szipcode")
	 vtextshipcountry = session("scountry")
	 vtxtstate = session("sstate")
end if

'call sitelink.TOTALORDER(session("shopperid"),session("ordernumber"),session("billtocopy"),session("shippingmethod"))
'call sitelink.WRITE_SHIPMODIFY_FLAG(session("ordernumber"),false)
				 
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
<style type="text/css">
<!--
.style7 {font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: 14px;
	color: #FFFFFF;}
.style15 {color: #FF0000}
.style3 {	font-size: 22px;
	font-weight: bold;
	color: #FF0000;
}
-->
</style>
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">



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
	<center><br><p align="center"></p>
        <div align="center">
    <table width="98%"  border="0" cellspacing="1">
              <tr>
                <td align="center" valign="middle" bgcolor="#29417A"><table width="99%" height="339"  border="0" cellspacing="2">
                    <tr>
                      <td height="43" bgcolor="#29417A"><p class="style7">PLEASE NOTE WHEN SHIPPING... </p></td>
                    </tr>
					<tr>
                      <td height="40" align="left" valign="middle" bgcolor="#CCCCCC" class="Plaintextbold"><ul>
                        <li>IF YOU ARE PLACING AN ORDER WITH <span class="style15">FROZEN FOODS AND GROCERY ITEMS</span>, PLEASE CREATE TWO SEPARATE ORDERS. FROZEN FOODS <span class="style15">MUST BE SHIPPED VIA FEDEX OVERNIGHT</span> AND TRAVEL SEPARATELY FROM ANY GROCERY ITEMS.</li>
                      </ul><br></td>
                    </tr>
                    <tr>
                      <td height="40" align="left" valign="middle" bgcolor="#FFFFFF" class="Plaintextbold"><ul>
                        <li>IF YOU ARE USING A <span class="style15">COMPANY FEDEX ACCOUNT</span>, PLEASE INCLUDE YOUR FEDEX ACCOUNT NUMBER IN THE SPECIAL INSTRUCTIONS BOX ON THE NEXT PAGE. PACKING AND HANDLING CHARGES WILL APPLY.</li>
                      </ul><br></td>
                    </tr>
                    <tr>
                      <td height="40" align="left" valign="middle" bgcolor="#CCCCCC" class="Plaintextbold"><ul>
                        <li>PRIOR TO  SHIPPING TO AN <span class="style15">INTERNATIONAL ADDRESS,</span> WE MAY REQUIRE FURTHER VERIFICATION FROM YOU ABOUT YOUR ORDER.  ALL INTERNATIONAL ORDERS INCUR A <span class="style15">$7.95 CUSTOMS FORMS CHARGE</SPAN>. COSMETIC, MEDICINAL, AND GROCERY ITEMS <span class="style15">CANNOT</span> BE SHIPPED INTERNATIONALLY (EXCEPT TO CANADA).</li>
                      </ul><br></td>
                    </tr>
					<tr>
                      <td height="40" align="left" valign="middle" bgcolor="#FFFFFF" class="Plaintextbold"><ul>
                        <li>SHIPPING TO A U.S. ADDRESS <span class="style15">OUTSIDE THE CONTINENTAL UNITED STATES</span>  MAY INCUR ADDITIONAL SHIPPING CHARGES AND ARE SHIPPED VIA THE US POSTAL SERVICE ONLY.</li>
                      </ul><br></td>
                    </tr>
                    <tr>
                      <td height="40" align="left" valign="middle" bgcolor="#CCCCCC" class="Plaintextbold"><ul>
                        <li><span class="style15">HEAVY ORDERS </span> (SUCH AS CORDIAL, GINGER BEER, NAPISAN, AND CANNED GOODS) WILL REQUIRE <span class="style15">ADDITIONAL</span> SHIPPING CHARGES DUE TO THE WEIGHT. WE WILL CONTACT YOU IF ADDITIONAL CHARGES ARE REQUIRED. </li>
                      </ul><br></td>
                    </tr>
                    <tr>
                      <td height="40" align="left" valign="middle" bgcolor="#FFFFFF" class="Plaintextbold"><ul>
                        <li>PLEASE CALL FOR SHIPPING RATES IF YOUR ORDER IS <span class="style15">OVER $300.00</span>.</li>
                      </ul><br></td>
                    </tr>
                    <tr>
                      <td height="36" align="left" valign="middle" bgcolor="#CCCCCC" class="Plaintextbold"><ul>
                        <li><span class="style15">CHOCOLATES DURING SUMMER</span> GENERALLY TRAVEL WELL. HOWEVER, TO AVOID HEAT DAMAGE, YOU MAY WISH TO CONSIDER EXPEDITED DELIVERY.</li>
                      </ul><br></td>
                    </tr>
					<tr>
                      <td height="40" align="left" valign="middle" bgcolor="#FFFFFF" class="Plaintextbold"><ul><li>
                        <p>RATES LISTED ARE BASED ON THE <span class="style15">RETAIL VALUE</span> OF THE PRODUCTS. SHIPPING CHARGES MAY BE ADJUSTED TO COMPENSATE FOR <span class="style15">SIZE AND/OR WEIGHT</span> AND FINAL SHIPPING CHARGES WILL BE ADJUSTED ACCORDINGLY.</p>
                      </li></ul><br></td>
                    </tr>
					<tr>
                      <td height="40" align="left" valign="middle" bgcolor="#CCCCCC" class="Plaintextbold"><ul>
                        <li>GROCERY ITEMS <span class="style15">CANNOT BE RETURNED</span> WITHOUT PRIOR AUTHORIZATION FROM A CUSTOMER SERVICE REPRESENTATIVE. PLEASE CALL 1-888-422-9259 TO SPEAK WITH A CUSTOMER SERVICE REPRESENTATIVE.</li>
                      </ul><br></td>
                    </tr>
					<tr>
                      <td height="40" align="left" valign="middle" bgcolor="#FFFFFF" class="Plaintextbold"><ul>
					  <li>IF YOU DO NOT SEE THE SHIPPING OPTION YOU REQUIRE LISTED BELOW, CHOOSE ANY OPTION AND PLEASE <span class="style15">INDICATE YOUR DESIRED SHIPPING METHOD (via FEDEX or USPS)</span> IN THE SPECIAL INSTRUCTIONS FIELD ON THE NEXT PAGE. WE WILL MAKE ANY AND ALL ADJUSTMENTS TO YOUR ORDER AFTER WE DOWNLOAD IT AND BEFORE WE CHARGE YOUR CARD.
                      </li></ul><br>
                      <p><h1><span class="style3">All orders with residential ground addresses will ship via FedEx Home Delivery.</span></p><br></td>
                    </tr>
                </table>
				</td>
              </tr>
            </table>
            <br> </div>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0"  >
		<tr><td class="TopNavRow2Text" width="100%" >Select Shipping Method<%'=session("shopperid")%></td></tr>
	</table>
	<br>
	
	
	<br>
	 <form method="post" action="saveshipping.asp">
			<table width="100%"   border="0" cellpadding="0" cellspacing="0">
			<tr>
			    <td class="plaintextbold">
				<%
				shipping_order="3"
				
				shipping_order = cstr(SHIPPING_SORTBY) 
				if SHIPPING_SORTORDER=1 then
					shipping_order = shipping_order +" desc"
				end if 
				
								
				xmlstring=sitelink.previewshipping(session("shopperid"),session("ordernumber"),cstr(vtextshipcountry),UCase(cstr(vtxtstate)),cstr(vtxtzipCode),shipping_order,session("shippingmethod"))				
				
				objDoc.loadxml(xmlstring)	
				set SL_Ship = objDoc.selectNodes("//preship" + shipfilter) 
				if SL_Ship.length = 0 then
				%>
				<center>
				<table >
				<tr><td class="plaintextbold">
				No Shipping Methods are currently available. <br>
				Click the continue button to complete order. <br>
				You will be contacted at a future time about shipping. Or <br>
				Click back button to change shipping address.
				
				
				<input type="hidden" value="" name="txtshiplist">						
				</td></tr>
				</table>
				</center>
				<%else%>
					<table width="100%" cellpadding="3" cellspacing="1" border="0">
					<tr><td class="plaintextbold" colspan="2">
					<% if session("user_want_multiship")= true then %>
							Some or all line items does not have a shipping method.
							<br>
							Please select a shipping method.
					  <%else%>					 
						<br>
						Please select a shipping method.						
					<%end if%>
					<br><br>
					</td></tr>
					
					 <tr><td  class="THHeader">Shipping Method</td>
					 	<td  class="THHeader">Cost</td>
					  </tr>
					 <%
					  for x=0 to SL_Ship.length-1
					  SL_ShippingCode =  SL_Ship.item(x).selectSingleNode("shiplist").text
					  SL_ShippingTitle = trim(SL_Ship.item(x).selectSingleNode("ca_title").text)
					  SL_Shippingcost = SL_Ship.item(x).selectSingleNode("cost").text
					  
					 
					  
					  if (x mod 2) = 0 then
					   class_to_use = "tdRow1Color"
					  else
						class_to_use = "tdRow2Color"
					  end if 
					  
					  %>
					  <tr>
						<td class="<%=class_to_use%>">
						<span class="plaintext">
							<input name="txtshiplist" type="radio" value="<%=SL_ShippingCode%>" <%if x=0 then response.write("  checked") end if %>><%=SL_ShippingTitle%>
							</span>
							</td>
					  	  <td  class="<%=class_to_use%>"><span class="plaintext"><%=formatcurrency(SL_Shippingcost)%></span></td>
					</tr>
					  					
					<%next%>
					</table>
				<%
				set SL_Ship= nothing
				end if%>
				
				</td>
			</tr>
			
			<tr>
				<td align="center">
				
				<br><br><input type="Image" value="continue" src="images/btn-continue.gif" border="0" alt="continue">
				</td>
			</tr>
			</table>
		</form>
		</center>
			<P>

	
	</div>
<!-- end sl_code here -->
	<img alt="" src="images/clear.gif" width="1" height="100" border="0">
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
