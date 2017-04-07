<%on error resume next%>
<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<html>
<head>
<title>Aussie Products.com | Customer Order Reciept @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
	<meta http-equiv="Content-Type" content="application/xhtml+xml; charset=utf-8" />
	<meta name="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls" />
	<meta name="keywords" content="Aussie Products, Australian, Aussie, Australian Products, Australian Food, Aussie Foods, Australian Desserts, tim tams, vegemite, violet crumble, cherry ripe, koala, kangaroo, down under, Australian Groceries, foods, lollies, candy, gourmet food" />
	<meta name="robots" content="index, follow" />
    <meta name="robots" content="ALL">
    <!-- Update your html tag to include the itemscope and itemtype attributes -->
<html itemscope itemtype="http://schema.org/LocalBusiness">


<meta itemprop="name" content="Aussie Products.com | Australian Products Co. Bringing Australia to You - Aussie Foods">
<meta itemprop="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls">
<meta itemprop="image" content="http://www.aussieproducts.com/">
	<meta name="distribution" content="Global">
	<meta name="revisit-after" content="5">
	<meta name="classification" content="Food">
    <meta name="google-site-verification" content="K2A2-bw3DiKf9q0_Ul-o-hJLTo3YkFNb_JHtsYJ0LJ0"
	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />

<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />


<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
</head>
<body >

<%
	
set ORDER_RECORD = sitelink.ORDER_RECORD(session("previousorder"),false)
	points_used = ORDER_RECORD.points_usd
	points_amt = ORDER_RECORD.checkamoun
	SL_PrintOrderAmount  = ORDER_RECORD.ord_total
	ord_total            = ORDER_RECORD.ord_total
	SL_PrintTaxAmount 	= ORDER_RECORD.tax
	SL_PrintShipAmount 	= ORDER_RECORD.shipping
	GiftCardAmount 		= ORDER_RECORD.gcardamt
			  
set ORDER_RECORD=nothing

default_ship_title = trim(sitelink.dataforstatusnew("SHIPPINGMETHOD",session("previousorder"),session("previousshopperid"),"EMAIL",false))
 %>


<br><br>

<span class="plaintextbold" >Thank you for your 
                        order.<br>
						Your web confirmation number is <%=session("previousorder")%>.</span>
						<br><br>

<% 

		
		extrafield=""
		xmlstring = sitelink.getbasketinfo(session("previousshopperid"),session("previousorder"),extrafield,false,session("previoususer_want_multiship"))
		objDoc.loadxml(xmlstring)
		
		basket_has_multiship = false
		filter_condition = "//gbi[ship_to>0]"
		set SL_has_multishipto_basket = objDoc.selectNodes(filter_condition)
			if SL_has_multishipto_basket.length > 0 then
				basket_has_multiship= true
			end if		
		set SL_has_multishipto_basket = nothing
		
		
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
		
	    'if basket_has_giftcert=true then
		'	    filter_condition = "//gbi[item !='"+cstr(giftcertitem)+"']"					    
		'else
		'       filter_condition = "//gbi"
		'end if
		
		filter_condition = "//gbi[certid=0]"
		set SL_Basket = objDoc.selectNodes(filter_condition)


		show_shipping_option=false
		'if session("registeredshopper")<> "YES" then
		
		if basket_has_multiship= false then
		 %>
			<!--#INCLUDE FILE = "print-no-multishipReceipt.asp" -->
		<%else%>	 
		 	<!--#INCLUDE FILE = "print-multishipReceipt.asp" -->	 
		 <%end if%>

		<%
		set SL_Basket = nothing
		%>
		 <%
	 	points_used = 0 
		points_amt = 0 	 
		SL_PrintGrandTotal=SL_PrintOrderAmount
			
			  
			  remaining_balance = ord_total-points_amt
			  
			  if points_amt > 0 and points_amt > 0 then
  				'SL_PrintGrandTotal = cdbl(ord_total - points_amt)
				SL_PrintGrandTotal = cdbl(SL_PrintGrandTotal - points_amt)
				SL_PrintGrandTotal = round(SL_PrintGrandTotal,2)
			  
	 %>	
		 <table width="98%">
		<tr>
			<td align="right" class="plaintextbold">POINTS REDEEMED:</td>
			<td align="right" width="11%" class="plaintextbold"><%=formatcurrency(points_amt)%></td>
		
		</tr>		
		</table>
		
		<% 
		end if
		
		%>
		<%if basket_has_giftcert=true then 
		    'SL_PrintGrandTotal = (SL_PrintOrderAmount-giftcertAmount)
			SL_PrintGrandTotal = cdbl(SL_PrintGrandTotal-giftcertAmount)
		    SL_PrintGrandTotal = round(SL_PrintGrandTotal,4)

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
<%ordercomments = sitelink.DISPLAYNOTES(session("previousorder"))
  if len(trim(ordercomments)) > 0 then %>

 <table width="100%" border="0" cellpadding="0" cellspacing="0">
 <tr><td class="plaintextbold" width="150" valign="top">
	 Special Instructions :
	 </td>
	 <td valign="top" class="plaintext"><%=ordercomments%></td>
</tr>
 </table>

<%end if%>
<br><br>

</body>

</html>
<!--#INCLUDE FILE = "googletracking.asp" -->
<!--#INCLUDE FILE = "RemoveXmlObject.asp" -->

<%set sitelink=nothing%>