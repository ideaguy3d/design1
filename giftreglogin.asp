<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Registered Customer Gift Registry @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
<link rel="stylesheet" href="text/footernav.css" type="text/css">
<link rel="stylesheet" href="text/design.css" type="text/css">
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->

<center>

<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">



<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top"  >
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<table width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr><td width="60%" align="center" valign="top">	
		<table width="99%" cellpadding="3" cellspacing="0" border="0">	
		<tr><td width="18" align="left" valign="top"></td>
			<td width="297" align="left" valign="top">
				<font class="TopNavRow2Text"><b>Registered Gift Registry Customers</b></font><br><br></td>
		</tr>
			  
		<tr><td width="18" align="left" valign="top"></td>
		  <td width="297" align="left" valign="top" class="plaintext">
			<%if session("savegiftregid")="-1" then%>
				<font color="red"><b>Unable to locate the registry.<br>Please try again.</b></font>
			<%else%>
				Please enter the following information in order to locate your registry.
			<%end if%>				
			<br><br></td>
		</tr>
		<tr>
		  <td width="18" align="left" valign="top"></td>
		  <form name="registeredshopper" method="POST" action="verifygiftreg.asp">
		  <td width="297" valign="top" class="plaintext">
			<center>
			<table width="297" cellpadding="0" cellspacing="0" border="0">
			<tr><td class="plaintextbold">Email Address</td></tr>
			<tr><td><input type="text" name="regtxtemail" size="40,1" maxlength="60" value="<%=session("giftregemail")%>"></td></tr>
			<tr><td class="plaintextbold">
			<br>
			Password</td></tr>
			<tr><td><input type="password" name="txtregpassword" size="15" maxlength="10" value="<%=session("giftregpassword")%>"></td></tr>
			<tr><td><br><input type="Image" value="continue" src="images/btn-continue.gif" border="0" alt="continue" ></td></tr>
			</table>
			</center>
			<br>			
			Forget your password ?<br>
			Click <a class="allpage" href="emailgiftregpass.asp">here</a> to have password emailed to you.
			</td>
		  </form>
		</tr>
		</table>
		<td width="1" class="THHeader"><img src="images/clear.gif" width="1" border="0"></td>
		<td width="35%" align="center" valign="top">
		<table width="210" cellpadding="0" cellspacing="0" border="0">
		<tr><td width="18" align="left" valign="top"></td>
			<td width="182" align="left" valign="top">
				<font class="TopNavRow2Text"><b>New Gift Registry<br> Customers</b></font><br><br></td>
		</tr>
		<tr><td width="18" align="left" valign="top"></td>
		  <td width="182" align="left" valign="top" class="plaintext">
				Click Register button to create your Gift Registry.<br><br>&nbsp;<a href="giftcustregister.asp"><img src="images/btn-register.gif" border="0" value="Register" ></a>
			<br><br></td>
		</tr>
		</table>
		</td>
	</tr>
	</table>
	</div>
<!-- end sl_code here -->
	</td>
</tr>

</table>

</div> <!-- Closes main  -->
<div id="footer">
<!--#INCLUDE FILE = "text/footer.asp" -->
</div>
</div> <!-- Closes container  -->


</center>

</body>
</html>
