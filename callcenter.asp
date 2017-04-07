<%on error resume next%>


<%
	if session("callcenter_logged") = false then
		response.redirect("call_center_login.asp")
	end if

%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
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
	<table width="95%" cellpadding="3" cellspacing="0" border="0"  >
		<tr><td class="TopNavRow2Text" width="100%" >Order Information</td></tr>
	</table>
	<br>
	<br>
	<center>
	<table width="95%" border="0" cellspacing="1" cellpadding="3">
	<tr><td width="50%" valign="top">
		<form action="get_custinfo.asp" method="post">
		<table width="100%" border="0" cellspacing="1" cellpadding="1">
	 		<tr><td bgcolor="#333333">			
			 <table width="100%" border="0" cellspacing="0" cellpadding="3" >
			 <tr><td class="pagenavbg" colspan="2">&nbsp;</td></tr>
			 <tr><td class="pagenavbg" align="right"><span class="plaintextbold">FirstName:</span></td>
			 	 <td class="pagenavbg"><input class="plaintext" type="text" name="fname" size="20" maxlength="15"></td>
			 </tr>
			 <tr><td class="pagenavbg" align="right"><span class="plaintextbold">LastName:</span></td>
			 	 <td class="pagenavbg"><input class="plaintext" type="text" name="lname" size="20" maxlength="20"></td>
			 </tr>
			  <tr><td class="pagenavbg" align="right"><span class="plaintextbold">Zipcode:</span></td>
			 	 <td class="pagenavbg"><input class="plaintext" type="text" name="lzipcode" size="20" maxlength="10"></td>
			 </tr>
			 <tr><td class="pagenavbg" colspan="2">&nbsp;<br>
			 <input type="hidden" name="getwhat" value="byfname">
			 </td></tr>
			 <tr><td class="pagenavbg" colspan="2" align="center">
			 <input class="FormButtons" type="submit" value="Get History">
			 </td></tr>
			 <tr><td class="pagenavbg" colspan="2">&nbsp;<br></td></tr>
			
			 </table>
		</table>
		</form>
	
		</td>
		<td width="50%" valign="top">
			<form action="get_custinfo.asp" method="post">
			<table width="100%" border="0" cellspacing="1" cellpadding="1">
	 		<tr><td bgcolor="#333333">			
			 <table width="100%" border="0" cellspacing="0" cellpadding="3" >
			 <tr><td class="pagenavbg" colspan="2">&nbsp;</td></tr>
			  <tr><td class="pagenavbg" align="right"><span class="plaintextbold">Customer Number:</span></td>
			 	 <td class="pagenavbg"><input class="plaintext" type="text" name="bcustnum" size="20" maxlength="15"></td>
			 </tr>
			 <tr><td class="pagenavbg" colspan="2">&nbsp;<br>
			 <input type="hidden" name="getwhat" value="bycustnum">
			 </td></tr>


			 <tr><td class="pagenavbg" colspan="2" align="center">
			 <input class="FormButtons" type="submit" value="Get History">
			 </td></tr>
			 
			 </table>
		</table>
		</form>
		<form action="get_custinfo.asp" method="post">
			<table width="100%" border="0" cellspacing="1" cellpadding="1">
	 		<tr><td bgcolor="#333333">			
			 <table width="100%" border="0" cellspacing="0" cellpadding="3" >
			 <tr><td class="pagenavbg" colspan="2">&nbsp;</td></tr>
			  <tr>
                          <td class="pagenavbg" align="right" valign="top" width="40%"><span class="plaintextbold">Phone 
                            Number:</span></td>
			 	          <td class="pagenavbg" width="60%"><input class="plaintext" type="text" name="bphone" size="20" maxlength="15">
				 <br>
				 <span class="plaintext">eg.8001234567
				 <br>No dashes
				 <br>No spaces
				 <br>No parenthesis
				 </span>
				 </td>
			 </tr>
			 <tr><td class="pagenavbg" colspan="2"><br>
			 <input type="hidden" name="getwhat" value="ByPhone">
			 </td></tr>


			 <tr><td class="pagenavbg" colspan="2" align="center">
			 <input class="FormButtons" type="submit" value="Get History">
			 </td></tr>
			 
			 </table>
		</table>
		</form>
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
