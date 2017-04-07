<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!--#INCLUDE FILE = "include/momapp.asp" -->

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->



<html>
<head>
<title>Aussie Products.com | Quick Order Search Page | Australian Products Co.</title>
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
	
    <br>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0"  >
		<tr><td class="TopNavRow2Text" width="100%" >Quick Order Search</td></tr>
	</table>
	<br>
	<form method="POST" id="searchprodform" action="searchprods.asp">
        	<input name="ProductSearchBy" value="2" type="hidden">
		 <table  border="0" cellspacing="0" cellpadding="0"  width="90%">
				  <tr><td >
				 <br>
 				 <span class="plaintext">
					 <b>Order By Product Number or Product Name.</b> 
                     <br>
                     <br>
                     This product order search will allow you to select products by Item Number and Product Name as displayed in our catalog. Please enter the Item number or product name you wish to purchase in the boxes below.
				</span>
				
				  </td></tr>
				  
				  <tr>
				  <td>
				  <br>
                  <br>
                  <center>
				            <div class="divsearch">
                    <ul class="search-wrap">
                        <li class="searchbox"><input class="plaintext" type="text" size="18" maxlength="256" name="txtsearch" value="Quick Order Search" onfocus="if (this.value=='Product Search') this.value='';" onblur="if (this.value=='') this.value='Product Search';"></li>                
                        <li class="btn-go"><input type="image" src="images/btn_go.gif" style="height:24px;width:24px;border:0"></li>
                    
                    </ul>
                </div>
		</td></tr>
		  </table>
 </form>
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
