<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<%
Response.ExpiresAbsolute = #Feb 18,1998 13:26:26#
%>



<%

'orderconfirm = sitelink.ORDER_CONFIRMED(session("ordernumber"))
'	if orderconfirm = true then
'	  	set sitelink=nothing
'	  	set ObjDoc = nothing
'		response.redirect("receipt.asp")'
'	end if
CALL sitelink.EMPTYBASKET(SESSION("SHOPPERID"),SESSION("ORDERNUMBER")) 		
session("SL_BasketCount") = 0
session("SL_BasketSubTotal") = 0
session("DoNotShow_PointsTab")=false
session("DoNotShow_GCTab")=false
session("DoNotShow_GCardTab")=false

'CLEAR THE cookie for this order
session("cookiename")=ShortStoreName+"order"
RESPONSE.COOKIES(session("cookiename")).Expires="January 1, 1997"



%>

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Empty Shopping Cart | Australian Products Co.</title>
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
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">



<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<center>
    <br>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%">Cart Empty</td></tr>
	</table>
	</center>
	<br>
	 <p>&nbsp;
	<p align="center" class="plaintextbold">Your cart is empty.<br><br>
		  <a HREF="default.asp"><img src="images/btn-continueshopping.gif" border="0" alt="Continue Shopping"></a>
	 

	
	</div>
<!-- end sl_code here -->
	<br>
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
