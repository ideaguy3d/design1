<%on error resume next%>


<%
	if session("callcenter_logged") = false then
		response.redirect("call_center_login.asp")
	end if
	
	
	
	
	thiscust = "0"
	lcname = ""
	thisphone =""
	
	search_by = request("getwhat")

	select case search_by
	case "byfname"
		fname= trim(request("fname"))
		lcname= trim(request("lname"))
		lzipcode= trim(request("lzipcode"))		
	case "bycustnum"
		thiscust = cstr(trim(request("bcustnum")))		
	case "ByPhone"
		thisphone = cstr(trim(request("bphone")))
	END SELECT
	
	'if search_by = "byfname" then
	'	fname= trim(request("fname"))
	'	lcname= trim(request("lname"))
	'	lzipcode= trim(request("lzipcode"))	
	'else
	'	thiscust = cstr(trim(request("bcustnum")))
	'end if
	
 



%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Customer Information @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
<STYLE>
<!--
	INPUT  { font-size: 10px;  color: #000000 ; }

-->
</STYLE>
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
	<table width="100%" cellpadding="3" cellspacing="0" border="0"  >
		<tr><td class="TopNavRow2Text" width="100%" >Customer Information
		
		<%'response.write("<br>===="&session("ordernumber")&"---"&SL_order_confirmed)%>
		</td></tr>
	</table>
	<br>
	
	<%	
	'xmlstring=sitelink.Get_Cust_Information(thiscust,cstr(lcname))
	xmlstring = sitelink.GET_CUST_RECORD(cstr(search_by),cstr(thiscust),cstr(thisphone), cstr(fname),cstr(lcname),cstr(lzipcode))
		objDoc.loadxml(xmlstring)
					
		targetnode = "//cust_rec[lastname[string-length(.)> 0]]"			
		'set SL_Custlist = objDoc.selectNodes("//cust_rec")
		set SL_Custlist = objDoc.selectNodes(targetnode)
		
	%>
	<table width="100%" border="0" cellspacing="1" cellpadding="1">
	
			<tr><td bgcolor="#333333"> 		
			 <table width="100%" border="0" cellspacing="1" cellpadding="3" >
			 
			 <tr><td class="THHeader" width="80">Customer Number</td>
				 <td class="THHeader">&nbsp;</td>
				<td class="THHeader">Customer Information</td>
			</tr>
			<%	if SL_Custlist.length= 0 then %>
				<tr><td colspan="3" class="pagenavbg"  align="center"><span class="plaintextbold">No Customers found</span></td></tr>
			<%else%>
				<%
				for x=0 to SL_Custlist.length-1
				
				
					SL_custnum = SL_Custlist.item(x).selectSingleNode("custnum").text
					SL_lastname = SL_Custlist.item(x).selectSingleNode("lastname").text
					SL_firstname = SL_Custlist.item(x).selectSingleNode("firstname").text
					SL_lcompany = SL_Custlist.item(x).selectSingleNode("company").text
					SL_addr = SL_Custlist.item(x).selectSingleNode("addr").text
					SL_addr2 = SL_Custlist.item(x).selectSingleNode("addr2").text
					SL_addr3 = SL_Custlist.item(x).selectSingleNode("addr3").text
					SL_city = SL_Custlist.item(x).selectSingleNode("city").text
					SL_state = SL_Custlist.item(x).selectSingleNode("state").text
					'SL_county = SL_Custlist.item(x).selectSingleNode("county").text
					SL_zipcode = SL_Custlist.item(x).selectSingleNode("zipcode").text
					SL_phone = SL_Custlist.item(x).selectSingleNode("phone").text
					SL_email = SL_Custlist.item(x).selectSingleNode("email").text
					IS_SLWeborder = SL_Custlist.item(x).selectSingleNode("weborder").text
					SL_country = SL_Custlist.item(x).selectSingleNode("country").text
					
					'SL_county=sitelink.Get_CountyName(SL_county)
					'SL_country=sitelink.countryname(SL_country)
				%>	
			<form action="get_cust_history.asp" method="post">
			<tr>
				<td class="pagenavbg"><span class="plaintext"><%=SL_custnum%></span></td>
				<td class="pagenavbg" align="center"><input class="FormButtons" type="submit"  value="Get History" >
				<input type="hidden" name="bcustnum" value="<%=SL_custnum%>">
				<input type="hidden" name="isweborder" value="<%=IS_SLWeborder%>">
				</td>
				<td class="pagenavbg"><span class="plaintext"><%=SL_lastname%>
				&nbsp;
				<%=SL_firstname%>&nbsp;<%=SL_addr%>&nbsp;<%=SL_addr2%>&nbsp;<%=SL_addr3%>&nbsp;<%=SL_city%>
				&nbsp;
				<%=SL_state%>&nbsp;<%=SL_county%>&nbsp;<%=SL_zipcode%>
				<br>
				<%=SL_phone%>&nbsp;<%=SL_email%>
				
				</span>
				
				
				</td>
			</tr>
			</form>
			<%next %>
			<%end if
			
			set SL_Custlist =nothing%>
			 </table>
			</td>
		</tr>
	</table>
	<form action="default.asp" method="post">
	<input type="submit" class="FormButtons" name="continueshopping" value="Continue Shopping">
	</form>
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
