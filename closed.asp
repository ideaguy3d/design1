<%on error resume next%>



<%
	
	
	set sitelink=createobject("sitel700.sitelink")
	call sitelink.setdata(datapath,cstr(""),ShortStoreName)
	
	
    function url_cleanse(thatvar)
		'remove spaces and slashes		
		thatvar = replace(thatvar," ","-")
		thatvar = replace(thatvar,"/","_",1,10,1)
		'tharvar = replace(thatvar,"/","_")
		thatvar = replace(thatvar,">","")
		thatvar = replace(thatvar,"<","")
		url_cleanse = thatvar	
	end function



%>

<!--#INCLUDE FILE = "text/adminPayPref.asp" -->
<!--#INCLUDE FILE = "text/admincartpref.asp" -->
<!--#INCLUDE FILE = "text/adminPagePref.asp" -->
<!--#INCLUDE FILE = "text/adminCustMagmtpref.asp" -->
<!--#INCLUDE FILE = "text/adminProdDispPref.asp" -->
<!--#INCLUDE FILE = "text/adminProdPref.asp" -->
<!--#INCLUDE FILE = "text/adminDeptPref.asp" -->
<!--#INCLUDE FILE = "text/adminSeoPref.asp" -->
<!--#INCLUDE FILE = "text/LivePersonPref.asp" -->
<!--#INCLUDE FILE = "text/storetext.asp" -->
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
<table width="630" cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%" >Store is Closed</td></tr>
	</table>
	<br><br>
	
	<span class="plaintext">
	<ul>
	<font size="+2"  color="#FF0000">
	We are currently upgrading our purchasing software and www.aussieproducts.com's online store will be unavailable from 12:30 pm to 1:00 PM (PST) Monday, April 18, 2011 to 1:00 pm.
	<br>
	</br>
	For all of your online orders please call us toll free at:
	1-888-422-9259
	</font>
	</ul>
	</span>

	</center>
	<!-- end sl_code here -->
	</td>
	<td width="1" class="THHeader"><img src="images/clear.gif" width="1" border="0"></td>
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
