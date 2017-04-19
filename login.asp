<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<%

 if REQUEST.servervariables("server_port_secure")<> 1 and SLsitelive=true then 
    set sitelink = nothing
 	response.redirect(secureurl+"login.asp") 
  end if

	
 if ucase(session("destpage")) = "BASKET.ASP" and session("registeredshopper")="YES" and MULTI_SHIP_TO=1 then
 	set sitelink = nothing
 	session("user_want_multiship") = true	
	response.redirect("basket.asp") 
 end if 

%>


<!--#INCLUDE FILE = "CreateXmlObject.asp" -->



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Account Log-in Page- Australian Products Co.</title>
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

<!-- Latest compiled and minified CSS -->
<link rel="stylesheet" href="design/styles/julius-css.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
<link rel="stylesheet" href="design/jstyles.css">
<link rel="stylesheet" href="design/styles/jstyles.css">
<link rel="stylesheet" href="design/bower_components/font-awesome/css/font-awesome.css">

<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />

</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav-modern.asp" -->
    <div id="main">



<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
	<%
	  if session("destpage")<>"custinfo.asp" then
	    WANT_SIDENAV=1
	  end if 
	%>
	<%if WANT_SIDENAV=1 then%>
	<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
	<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	
	<%end if%>
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<center>
    <br>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%">Login</td></tr>
	</table>
	<br>
	<span class="plaintextbold">
	You must Login or Register to gain access to this feature.
	</span>
	<br>
	<br>
	
	
	<table width="100%"  cellpadding="0" cellspacing="0" border="0">
	<tr><td width="50%" valign="top">
			<table width="100%"  cellpadding="0" cellspacing="0" border="0">
			<tr><td align="center"><font class="TopNavRow2Text"><b>&nbsp;New Customers&nbsp;</b></font><br><br></td></tr>
			<tr><td align="center" class="plaintextbold">Registration is fast and easy.</td></tr>
			<tr><td align="center">
			<br><br>
			<a href="firsttimeregister.asp"><img src="images/btn-register.gif" border="0" alt="Register"></a></td></tr>
			</table>
		</td>
		<td width="2" class="THHeader nopadding"><img alt="" src="images/clear.gif" width="1" border="0"></td>
		<td width="50%" valign="top">
			<form action="<%=secureurl%>verifyshopper.asp" method="post">
			<table width="100%"  cellpadding="0" cellspacing="0" border="0">
			<tr><td align="center"><font class="TopNavRow2Text"><b>&nbsp;Registered Customers&nbsp;</b></font><br><br></td></tr>
			<tr><td align="center" class="plaintext">
				<%if session("failedtofindshopper")="YES" then%>
					<font color="red"><b>The email address or password you provided is not valid.<br>Please try again.</b></font>
				<%else%>
					Please enter the following information in order to sign in.
				<%end if%>				
				<br><br></td></tr>
				<tr><td align="center">
				<table width="85%" cellpadding="0" cellspacing="0" border="0">
				<tr><td class="plaintextbold" align="left">Email Address</td></tr>
				<tr><td align="left"><input type="text" name="regtxtemail" size="40,1" maxlength="60" value="<%=session("regemail")%>"></td></tr>
				<tr><td class="plaintextbold" align="left">
				<br>
				Password</td></tr>
				<tr><td align="left"><input type="password" name="txtregpassword" size="15" maxlength="10" value="<%=session("regpassword")%>"></td></tr>
				<tr><td align="left"><br><input type="Image" value="continue" src="images/btn-continue.gif" style="border:0" alt="continue"></td></tr>
				</table>
				</td>
				</tr>
				<tr><td class="plaintext" align="center">
			
				<br><br>
				
				Forget your password ?<br>
				<a class="allpage btn btn-sm btn-info" href="emailpassword.asp">Click here</a> to have password emailed to you.
			
				</td>
				</tr>
			</table>
			</form>
		</td>
		
	
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
