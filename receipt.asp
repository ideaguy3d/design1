<%on error resume next%>

<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Order Receipt | Australian Products Co.</title>
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
<link rel="alternate" media="print" href="popup-receipt.asp">
<script>
function printWindow(){
bV = parseInt(navigator.appVersion)
if (bV >= 4) window.print()
}
function clearorder() {
		
		window.location="neworder2.asp"	
 }

</script>
</head>
<!--#INCLUDE FILE = "text/Backgroundreceipt5.asp" -->
<div id="container">
    <div id="header" class="topnav1bgcolor">    	
        	<div class="divlogo">
            	<div class="logo-wrap">
                    <div class="logo-img">
                        <a href="<%=insecureurl%>" title="<%=althomepage%>">
                        <%if COMPANY_LOGO_IMG<>"" then %>
                        <img alt="<%=althomepage%>" src="images/<%=COMPANY_LOGO_IMG%>" border="0" >
                        <%else %>
                        <img alt="<%=althomepage%>" src="images/clear.gif" width="1" height="1" border="0" >
                        <%end if %>
                        </a>
                    </div>
                </div>
            </div>
            <div class="restore-order"></div>        
	</div>
    <div id="main">


<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
	<div id="page-content">
	<br>
	<table width="100%"  cellpadding="0" cellspacing="0" border="0">
		<tr>
            <td align="left" width="100%"><img src="images/step3.gif" border="0" width="403" height="29"></td>
          </tr>
	</table>
	<br>
	<table width="98%" cellpadding="2" cellspacing="0" border="0">
	<tr><td class="plaintextbold">
        <div style="float: left; width: 50%;">Thank you for your order.<br>
            Your web confirmation number is <%=session("previousorder")%>.
            <br><br>
        </div>
        <div align="right" style="float: right; width: 50%;">
			<br>
            <a class="allpage" href="javascript:printWindow()"><strong>Click Here to Print this Page</strong><br></a>			
		</div>
	</td>
	</tr>
	</table>
		<% 'if session("multishipto")= 0 then

			'get all billto and ship to info from session("previousorder")	
			billtocust = sitelink.GetBilltoCust(session("previousorder"))
	        set shopperRecord= sitelink.SHOPPERINFO(cstr(billtocust),false)
	        session("previousshopperid") = shopperRecord.custnum
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
						 
			 if session("previousbilltocopy")=0 then
			 	set shopperRecord= sitelink.SHIPTONAME(cstr(session("previousorder")),false)
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
			 end if
			 
	
		set SL_OrderRecInfo = New cOrderRecObj
		SL_OrderRecInfo.lordernumber=session("previousorder")
		SL_OrderRecInfo.lshopperid=session("previousshopperid")
		SL_OrderRecInfo.lPagename=""
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
		set SL_OrderRecInfo=nothing
		
		'response.write(points_amt)
		
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
		'        filter_condition = "//gbi"
		'end if
		
		filter_condition = "//gbi[certid=0]"
		set SL_Basket = objDoc.selectNodes(filter_condition)
	

		'if session("registeredshopper")<> "YES" then
		if basket_has_multiship= false then
	 %>
		<!--#INCLUDE FILE = "print-no-multishipReceipt.asp" -->
	<%else%>	 
	 	<!--#INCLUDE FILE = "print-multishipReceipt.asp" -->	 
	 <%end if%>
	 <%' set SL_Basket = nothing%>
	 
	 

 <%
		'SL_PrintOrderAmount is populated in print-no-multiship.asp/print-multiship.asp
		SL_PrintGrandTotal=SL_PrintOrderAmount
	 
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
	end if%>
	
	<%if basket_has_giftcert=true then 
		SL_PrintGrandTotal = cdbl(SL_PrintGrandTotal-giftcertAmount)
	    SL_PrintGrandTotal = round(SL_PrintGrandTotal,4)

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
			 
	 
	
	
	<table width="98%" cellpadding="2" cellspacing="0" border="0">
	<tr><td>&nbsp;</td>						
        <td align="center"> 
		<a href="<%=insecureurl%>neworder2.asp"><img src="images/btn_PlaceNewOrder.gif" border="0" alt="Place New Order"></a>
		</td>
	</tr>
	</table>	
	
	
	</div>
<!-- end sl_code here -->
	</td>	
</tr>

</table>



<!--#INCLUDE FILE = "googletracking.asp" -->

<%    
	
	orderTotal=formatcurrency(SL_ORD_TOTAL)
	orderTotal=replace(orderTotal,"$","")
	orderTotal=replace(orderTotal,",","")
			
	tax=formatcurrency(SL_PrintTaxAmount)
	tax=replace(tax,"$","")
	tax=replace(tax,",","")
			
	shippingCost=formatcurrency(SL_PrintShipAmount)
	shippingCost=replace(shippingCost,"$","")
	shippingCost=replace(shippingCost,",","")
 %>

<%if len(trim(GA_ACCTNUM)) > 0 then
		
	
	
		
	'extrafield =""
	'xmlstring = sitelink.getbasketinfo(session("previousshopperid"),session("previousorder"),extrafield,false,session("user_want_multiship"))
		
	'objDoc.loadxml(xmlstring)
	set SL_Basket = objDoc.selectNodes("//gbi[certid=0]")
			
	%>
		
			
	<script type="text/javascript">
	   var pageTracker = _gat._getTracker("<%=GA_ACCTNUM%>");
	       pageTracker._initData();
		   pageTracker._trackPageview();
			
		  pageTracker._addTrans(
			"<%=session("previousorder")%>",
			"<%=COMPANY_NAME%>",
			"<%=orderTotal%>",
			"<%=tax%>",
			"<%=shippingCost%>",
			"<%=session("city")%>",
			"<%=session("state")%>",
			"<%=trim(sitelink.countryname(session("country")))%>"
		  );
			
		<%		
		   for x=0 to SL_Basket.length-1
			SL_Number			= SL_Basket.item(x).selectSingleNode("number2").text
			SL_Variant			= SL_Basket.item(x).selectSingleNode("variant").text
			SL_BasketNumber 	= SL_Number +" "+ SL_Variant
			SL_BasketQauntity	= SL_Basket.item(x).selectSingleNode("quanto").text
			SL_BasketDesc		= SL_Basket.item(x).selectSingleNode("inetsdesc").text
			SL_BasketDesc2		= SL_Basket.item(x).selectSingleNode("desc2").text
			SL_BasketUnitPrice	= SL_Basket.item(x).selectSingleNode("it_unlist").text
			SL_BasketExtPrice	= SL_Basket.item(x).selectSingleNode("extendp").text				
			SL_units 			=SL_Basket.item(x).selectSingleNode("units").text							
			
			
			SL_units 			=replace(SL_units,".00","")
			SL_units 			=replace(SL_units,",","")

			SL_BasketQty = replace(SL_BasketQauntity,".00","")
			SL_BasketQty = replace(SL_BasketQty,",","")				

			SL_BasketUnitPrice = formatcurrency(SL_BasketUnitPrice)
			SL_BasketUnitPrice = replace(SL_BasketUnitPrice,"$","")
			SL_BasketUnitPrice = replace(SL_BasketUnitPrice,",","")
			
			SL_BasketDesc = url_cleanse(SL_BasketDesc)
			SL_BasketDesc2 = url_cleanse(SL_BasketDesc2)
			
			SL_BasketDesc = replace(SL_BasketDesc,"-"," ")
			SL_BasketDesc = replace(SL_BasketDesc,"_"," ")
			SL_BasketDesc2 = replace(SL_BasketDesc2,"-"," ")
			SL_BasketDesc2 = replace(SL_BasketDesc2,"_"," ")
			
						
		%>					
					
		  pageTracker._addItem(
			"<%=session("previousorder")%>",
			"<%=SL_Number%>",
			"<%=SL_BasketDesc%>",
			"<%=SL_BasketDesc2%>",
			"<%=SL_BasketUnitPrice%>",
			"<%=SL_BasketQty%>"
		  );
		<%next		  
		%>
		  pageTracker._trackTrans();
</script>
<%end if%>

<%if SHOW_BUYSAFE =1 then %>
    <!-- BEGIN: buySAFE Guarantee-->
    <script src="https://sb.buysafe.com/private/rollover/rollover.js"></script>
    <span id="BuySafeGuaranteeSpan"></span> 
    <script type="text/javascript"> 
        buySAFE.Hash = '<%=BUYSAFEHASHVALUE%>';
        buySAFE.Guarantee.order = "<%=session("previousorder")%>"; 
        buySAFE.Guarantee.total = "<%=orderTotal%>"; 
        buySAFE.Guarantee.email = "<%=session("email")%>"; 
        WriteBuySafeGuarantee("JavaScript"); 
    </script> 
    <!-- END: buySAFE Guarantee-->
<%end if %>

<% set SL_Basket = nothing %>
    </div> <!-- Closes main  -->
    <div id="footer" class="footerbgcolor">    
    <div id="bottomlinks">
    <div class="powered-by">
        <a href="http://www.dydacomp.com/sitelink-ecommerce.asp" target="_blank"><img src="images/logo-sitelink7.png" border="0" alt="Order Management System powered by MOM and SiteLINK7"></a>
    </div>   
    </div> 

    <%if ACTIVATE_LIVE_PERSON=1 then %>
    <center>
    <!--#INCLUDE FILE = "text/LivepersonMonitorscript.asp" -->
    </center>
    <%end if %>
    
    <!--#INCLUDE FILE = "RemoveXmlObject.asp" -->
    <!--#INCLUDE FILE = "text/footer.asp" -->
    </div>
</div> <!-- Closes container  -->


</body>
</html>
