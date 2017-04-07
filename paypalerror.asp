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
		<tr><td class="TopNavRow2Text" width="100%" >Information Missing</td></tr>
	</table>
	<br>
	<br>
			<table width="100%"  border="0" cellpadding="0" cellspacing="0">
			<tr>
			    <td class="plaintextbold">
				We're sorry, but we cannot process your order because of the following:
				<br><br>
				<span style="color:#ff0000;">
				<%=session("custerrormessage")%>
				</span>
				<br>
				<br />
				
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
