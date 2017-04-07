<%on error resume next%>

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
	<center>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0"  >
		<tr><td class="TopNavRow2Text" width="100%" >Cookies Disabled</td></tr>
	</table>
	<br>
	</center>
	
	 <P>	
		<p>&nbsp;
		<p >
		<ul><font class="plaintextbold"><font size="+1" color="#FF0000">
		Cookies are currently disabled on your browser </font>
		<br><br><br>
		
		</font> 
		</ul>
		<br>
		<ul>
		<font class="plaintextbold">
		Please enable this in order to complete this process
		<br><br><br>
		If you have tried to add an item to your basket and the basket is still empty, you may not have "cookies" enabled on your browser. 
		<br><br><br>
		<a class="allpage" href="moreinfo.asp">More information</a> &nbsp;on cookies.</font>
		</ul>
	
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
