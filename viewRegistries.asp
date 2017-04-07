<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%
  email = cstr(trim(session("giftregemail")))
  password = cstr(trim(session("giftregpassword")))
  xmlstring = ""
  xmlstring = sitelink.GIFTREGSEARCH("","",True,session("giftregemail"),session("giftregpassword"))
  
  objDoc.loadxml(xmlstring)
  set SL_GiftReg = objDoc.selectNodes("//giftreg")	  				 
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | View Customer Gift Registry @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
		<table width="590" cellpadding="3" cellspacing="0" border="0">
			<tr><td class="TopNavRow2Text" width="630">
				<a href="giftregistry.asp"><font class="TopNavRow2Text">Gift Registry</font></a>
				&nbsp;>&nbsp;<a href="parentsearch.asp"><font class="TopNavRow2Text">Search</font></a></td></tr>
		</table>

		<BR><BR><BR>

		<TABLE align="center" width="95%" cellSpacing=0 cellPadding=0 border=0>
	  <% if SL_GiftReg.length = 0 then
		    response.write("<tr><td class=""plaintextbold"">No Gift Registry list found </td></tr>")
		    response.write("</TABLE>")
		else
		%>
		 <tr><td colspan=7 class="plaintextbold"><b>We found following registries. Please click "View" to view the gift registry.</b><br></td></tr>
  		 <tr><td colspan=7 class="plaintext">&nbsp;</td></tr>
  		 <tr><td colspan=7 class="plaintext">&nbsp;</td></tr>
  		 </TABLE>
  		 <TABLE align="center" width="95%" border="1" bordercolor="#f0e68c" cellSpacing=0 cellPadding=0>
		 <tr><th class="secondary">Registrant's First Name</th>
			 <th class="secondary">Registrant's Last Name</th>
			 <th class="secondary">City</th>
			 <th class="secondary">State</th>
			 <th class="secondary">Zip</th>
			 <th class="secondary">Date Of Event</th>
			 <th class="secondary">Type of Registry</th>
			 <th class="secondary">View Registry</th>
		 </tr>
		 <%
		 for x=0 to SL_GiftReg.length-1 
		 
			SL_GiftCustnum = SL_GiftReg.item(x).selectSingleNode("custnum").text
			SL_RegLastName = SL_GiftReg.item(x).selectSingleNode("lastname").text
			SL_RegFirstName = SL_GiftReg.item(x).selectSingleNode("firstname").text
			SL_City = SL_GiftReg.item(x).selectSingleNode("city").text
			SL_State = SL_GiftReg.item(x).selectSingleNode("state").text
			SL_Zip = SL_GiftReg.item(x).selectSingleNode("zipcode").text
			SL_EventDate = SL_GiftReg.item(x).selectSingleNode("eventdate").text
			SL_RegistryType = SL_GiftReg.item(x).selectSingleNode("reg_type").text
		 
			if len(trim(SL_EventDate)) > 0 then SL_EventDate = CDate(SL_EventDate)
			
			'url = "giftlist.asp?id=" & SL_GiftCustnum '& "&index=" & x %>
		 
		 <tr><td align="center"class="plaintext"><%=SL_RegFirstName%></td>
		 <td align="center"class="plaintext"><%=SL_RegLastName%></td>
		 <td align="center"class="plaintext"><%=SL_City%></td>
		 <td align="center"class="plaintext"><%=SL_State%></td>
		 <td align="center"class="plaintext"><%=SL_Zip%></td>
		 <td align="center"class="plaintext">&nbsp;<%=SL_EventDate%></td>
		 <td align="center"class="plaintext">&nbsp;<%=SL_RegistryType%>
			<%'=GetRegDesc(UCase(trim(sitelink.getdata(session("clshopper"),"gawishsearch",cint(x),21,0))))%>
		 </td>
		 <form action="giftlist.asp" method="post">
		 <input type="hidden" name="id" value="<%=SL_GiftCustnum%>">
		 <td align="center"class="plaintext"><input class="FormButtons" type="submit"  value="View"></td>
		 </form>
		 </tr>
		 <%next%>
		<%end if%>
	
        </TABLE>
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
