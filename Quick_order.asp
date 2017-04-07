<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!--#INCLUDE FILE = "include/momapp.asp" -->

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->



<html>
<head>
<title>Aussie Products.com | Quick Order Page | Australian Products Co.</title>
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
<% if request.servervariables("server_port_secure") = 1 then %>
	<base href="<%=secureurl%>">
<% else %>
	<base href="<%=insecureurl%>">
<% end if %>
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->




<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">

<% session("destpage")="" %>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top"  >
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<center>
    <br>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0"  >
		<tr><td class="TopNavRow2Text" width="100%" >Quick Order Form</td></tr>
	</table>
	<br>

		 <table  border="0" cellspacing="0" cellpadding="0"  width="90%">
				  <tr><td >
				 <br>
 				 <span class="plaintext">
					 <b>Order By Product Number:</b> This order form will allow you to select
					products by Product Number as displayed in our catalog. Please enter the
					product numbers you wish to purchase in the boxes below, and then press the
					"Add to Cart" button.
				</span>
				
				  </td></tr>
				  
				  <tr>
				  <td>
				  <br>
				  <form method="post" action="ProdAdd.asp" id="form1">
		  <TABLE  BORDER="0" cellpadding="2"  cellspacing="0" >			  
		 <tr><th class="plaintextbold"><font color="#93816b">Enter Item #/s</font></th>
		    <th class="plaintextbold"><font color="#93816b">Quantity</font></th>
		 	</tr>
		<% 
		num_of_quick_items = 10
		for i = 1 to num_of_quick_items
		 session("Stock"&i) = ""
		 session("txtquantm"&i) ="1"
		next%>
		<% for i = 1 to num_of_quick_items %>		
		<tr>	
		<td align="center"  ><span class="plaintextbold">Item #</span>&nbsp;<input name="Stock<%=i%>" class="InputTextSmall" maxlength="20"  value="<%=session("Stock"&i)%>" size="25"></td>
		<td>&nbsp;&nbsp;&nbsp;<input name="txtquantm<%=i%>" class="InputTextSmall" maxlength="5" size="5" value="<%=session("txtquantm"&i)%>"></td> 
		</tr>
		<%next%>
		
		<tr><td align="center">
		<br><br>
		<input type="hidden" name="num_of_quick_items" value="<%=num_of_quick_items%>">
		<input type="submit" value="Add to cart"  id=submit1 name=submit1>
		</td></tr>
		 </table>
		</form>
		</td></tr>
		  </table>

	</center>
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



</body>
</html>
