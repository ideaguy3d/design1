<%on error resume next%>
<%
  if session("registeredshopper")="NO" then
  	response.redirect("statuslogin.asp")
  end if 

%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Customer's Account Infomation  | Australian Products Co.</title>
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
<STYLE>
<!--
	INPUT  { font-size: 10px;  color: #000000 ; }

-->
</STYLE>



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
<br>
<!-- sl code goes here -->
<div id="page-content">
	
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%" >Account Information</td></tr>
	</table>
	<br>
	<table width="100%"  border="0"  cellpadding="2" cellspacing="1">	
		<TR><td   class="THHeader"   width="33%" align="center"><b>Customer Name</b></Td>
			 <td  class="THHeader"  width="33%" align="center"><b>Zip Code</b></Td>					 
		</TR>
		 <TR><Td  class="tdRow1Color"   width="33%" align="center"><span class="plaintext"><br><%=session("firstname")%>&nbsp;<%=session("lastname")%><br><br>
		 <%if ALLOW_POINTS=1 and session("pointavail")> 0 then %>
		 Available Points: <%=session("pointavail") %>
		 <%end if%>
		 </span></Td>
			 <Td  class="tdRow1Color"   width="33%" align="center"><span class="plaintext"><%=session("zipcode")%></span></Td>					
		 </TR>
	</TABLE>
	
	
	
				<br>
				<br>

	<table width="100%"  border="0"  cellpadding="3" cellspacing="1">
	<%if WANT_ADDRESS_CHANGE=1 then %>
	<tr>
         <td><a class="allpage" href="updatecustinfo.asp">Update your profile</a></td>
	<tr>
	<%end if%>
	<tr>
         <td><a class="allpage" href="statusoforders.asp">View your order history</a></td>
	<tr>
	<tr>
         <td><a class="allpage" href="address.asp">View your Address Book</a></td>
	<tr>
	<%if WANT_WISHLIST=1 then %>
	<tr>
         <td><a class="allpage" href="wishlist.asp">View your Wish List</a></td>
	<tr>
	<%end if%>
	<tr>
         <td>&nbsp;</td>
	<tr>
	<tr>
         <td>&nbsp;</td>
	<tr>
	<tr>
         <td><a class="allpage" href="logout.asp">Logout</a></td>
	<tr>
	
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
<script type="text/javascript">
  (function() {
    window._pa = window._pa || {};
    // _pa.orderId = "myOrderId"; // OPTIONAL: attach unique conversion identifier to conversions
    // _pa.revenue = "19.99"; // OPTIONAL: attach dynamic purchase values to conversions
    // _pa.productId = "myProductId"; // OPTIONAL: Include product ID for use with dynamic ads
    var pa = document.createElement('script'); pa.type = 'text/javascript'; pa.async = true;
    pa.src = ('https:' == document.location.protocol ? 'https:' : 'http:') + "//tag.marinsm.com/serve/55c94a54db45c07a90000003.js";
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(pa, s);
  })();
</script>
</body>
</html>
