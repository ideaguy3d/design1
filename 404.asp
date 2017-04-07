<%on error resume next%>


<!--#INCLUDE FILE = "include/cLibraryExtra.asp" -->
<!--#INCLUDE FILE = "include/momapp.asp" -->
<%
strUrlCopy = Trim(Mid(Request.QueryString, 5))

siteurl=insecureurl
siteurl80=replace(insecureurl,".com/",".com:80/")
'siteurl80=mid(insecureurl,1,len(insecureurl)-1) &":80/" 
strUrlCopy = replace(strUrlCopy,siteurl,"",1,1,1)
strUrlCopy = replace(strUrlCopy,siteurl80,"",1,1,1)
strUrlCopy = lcase(trim(strUrlCopy))

set SLSeoObj = New cSEOUrl
SLSeoObj.action = "READ"
SLSeoObj.oldurl = strUrlCopy
SLSeoObj.sl_url = "1"
SL_Url = sitelink.Maintain_SEO_Urls(SLSeoObj)
SL_Url= trim(SL_Url)
set SLSeoObj =nothing

if len(SL_Url) > 0 and SL_Url<> "-1" then
	'SL_Url_lcase= lcase(SL_Url)
	'insecurl_lcase= lcase(insecureurl)
	'insecurl_lcase= replace(insecurl_lcase,"/","")
	'insecurlwohttp_lcase= replace(insecurl_lcase,"http://","")
	'if InStr(1,insecurl_lcase,SL_Url,1)=0 and InStr(1,insecurlwohttp_lcase,SL_Url,1)=0 then
		if SL_Url_lcase <>"default.asp" then 
		    SL_Url = insecureurl & SL_Url
	    else 
	        SL_Url = insecureurl
		end if 
	'end if
	'Response.Clear()
	Response.Status="301 Moved Permanently" 
	Response.AddHeader "Location", SL_Url
	Response.End
	'Response.Redirect urltoredirect
	'Server.Transfer urltoredirect
else
	response.status = "404 Not Found"
end if
%>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<html>
<head>
<%=usethismeta%>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<% if request.servervariables("server_port_secure") = 1 then %>
	<base href="<%=secureurl%>">
<% else %>
	<base href="<%=insecureurl%>">
<% end if %>
<title>Aussie Products.com | 404 - Page Not Found | AussieProducts.com @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td>
		<tr><td class="TopNavRow2Text">The page you are looking for cannot be found.</td></tr>
		<tr><td class="plaintext"><br>Try searching for a product below or <a href="sitemap.asp" class="allpage">visit our sitemap</a>.</td></tr>
		<tr><td valign="top"><br />
		<form method="POST" action="searchprods.asp" >
			<table width="0"  border="0" cellpadding="4" cellspacing="5">
			<tr>
			    <td width="10%" class="plaintextbold">				
				Search :&nbsp;
				 <td class="plaintextbold">
				 <input type="text" size="50" maxlength="256" class="plaintext" name="txtsearch" value=""><input name="ProductSearchBy" value="2" type="hidden">
				 &nbsp;<input type="submit" value="Search">
				</td>
			</tr>
			
			</table>
			</form>
	</td></tr>
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