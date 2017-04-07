<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!--#INCLUDE FILE = "include/momapp.asp" -->

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->





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
	
	<%if WANT_SIDENAV=1 then%>
	<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
	<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	
	<%end if%>
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<center>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0"  >
		<tr><td class="TopNavRow2Text" width="100%" >Customer Information Missing</td></tr>
	</table>
	<br>
	<br>
			<table width="100%"  border="0" cellpadding="0" cellspacing="0">
			<tr>
			    <td class="plaintextbold">
				We're sorry, but we cannot process your order because of the following:<br><br>
				<%=session("custerrormessage")%>
				<%if len(trim(session("ProdRestList"))) > 0 then %>
				    <table width="100%"  border="0" cellspacing="1" cellpadding="3"
				    style="BORDER-RIGHT: #CCCCCC 1px solid  ; 
	    	            BORDER-TOP: #CCCCCC 1px solid  ; 
			            BORDER-LEFT: #CCCCCC 1px solid  ; 
			            BORDER-BOTTOM: #CCCCCC 1px solid "
				    >
				    <tr>
			            <th align="CENTER" class="THHeader" width="15%">Item</th>			
			            <th align="CENTER" class="THHeader">Description</th>
			            <th align="CENTER" class="THHeader" width="70%">Cannot Ship To</th>
			        </tr>
				<%
				    set shopperRecord = sitelink.SHIPTONAME(session("ordernumber"),false)
				        sfname = shopperRecord.firstname
					    slname = shopperRecord.lastname
					    slAddr = shopperRecord.addr
					    slAddr2 = shopperRecord.addr2
					    slCity = shopperRecord.city
					    slState = shopperRecord.state
					    slZipCode = shopperRecord.zipcode
					    slCountry = shopperRecord.country
					set shopperRecord = nothing	
				    				   
				    xmlstring=session("ProdRestList")
				    objDoc.loadxml(xmlstring)				    				  


				    set ProdRestrict = objDoc.selectNodes("//results[lrest_sale=1 or lrest_sale=0]") 
				    
				    
				    extrafield =""
		            xmlstring = sitelink.getbasketinfo(session("shopperid"),session("ordernumber"),extrafield,false,session("user_want_multiship"))
		            objDoc.loadxml(xmlstring)
		            
		          		            
		            for x= 0 to ProdRestrict.length-1
				        SLnumber = ProdRestrict.item(x).selectSingleNode("number").text
				        SLBasketShipto	= ProdRestrict.item(x).selectSingleNode("lshipto").text
				        filter_condition = "//gbi[item='"+cstr(SLnumber)+"' and ship_to="+cstr(SLBasketShipto)+"]"				        
				        set SL_Basket = objDoc.selectNodes(filter_condition)
				            if SL_Basket.length > 0 then
				                SL_BasketDesc		= SL_Basket.item(0).selectSingleNode("inetsdesc").text
				                SL_Basket_Shipto	= SL_Basket.item(0).selectSingleNode("ship_to").text
				                if SL_Basket_Shipto > 0 then
				                    fname = SL_Basket.item(0).selectSingleNode("firstname").text
						            lname = SL_Basket.item(0).selectSingleNode("lastname").text
						            lAddr = SL_Basket.item(0).selectSingleNode("addr").text
						            lAddr2 = SL_Basket.item(0).selectSingleNode("addr2").text
						            lCity = SL_Basket.item(0).selectSingleNode("city").text
						            lState = SL_Basket.item(0).selectSingleNode("state").text
						            lZipCode = SL_Basket.item(0).selectSingleNode("zipcode").text
						            lCountry = SL_Basket.item(0).selectSingleNode("Country").text	
						        else
						        	fname = sfname
						            lname = slname
						            lAddr = slAddr
						            lAddr2 = slAddr2
						            lCity = slCity
						            lState = slState
						            lZipCode = slZipCode
						            lCountry = slCountry
				                end if 
				        
				            end if
				        set SL_Basket =nothing
				        
				        if (x mod 2) = 0 then
					        class_to_use = "tdRow1Color"
				        else
					        class_to_use = "tdRow2Color"
				        end if 
				%>
				    
				    <tr class="plaintext">
			            <td align="CENTER" class="<%=class_to_use%>"><%=SLnumber%></td>			
			            <td align="CENTER" class="<%=class_to_use%>"><%=SL_BasketDesc%></td>
			            <td class="<%=class_to_use%>">
			            <img src="images/clear.gif" width="20" height="1" border="0">
			            <%=fname%>&nbsp;<%=lname%>&nbsp;<%=lAddr%>&nbsp;<%=lAddr2%>&nbsp;<%=lCity%>&nbsp;<%=lState%>&nbsp;<%=lZipCode%>&nbsp;<%=sitelink.countryname(lCountry)%></td>
			        </tr>
				    
				    <%next %>
				    </table>
				    <br><br>
				    
				        Please change the shipping location for these items or remove it from the cart.
				    
				<%set ProdRestrict = nothing
				else
				%>
				
				<br>
				Please go back to the <A HREF="Javascript:history.go(-1)">Previous Page</A> and complete the required information
				<%end if %>
				
				
				</td>
			</tr>
			</table>
		<br>

	</center>
	</div>
<!-- end sl_code here -->
	<img alt="" src="images/clear.gif" width="1" height="160" border="0">
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
