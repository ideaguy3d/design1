<%on error resume next%>
<% 
	  session("destpage")="selectaddress.asp"
	if session("registeredshopper")<>"YES" then 
	  response.redirect("login.asp")
	end if
%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Select Customer Address @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">

<%session("destpage")="selectaddress.asp" %>

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
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%" >Address Book</td></tr>
	</table>
	<br>
		<P align="left">&nbsp;&nbsp;<a href="AddNewAddressBook.asp"><span class="TopNavRow2Text">&nbsp;<u>Click here to Add New Addresses</U>&nbsp;</span></a>
		<br><br>
	<%if session("registeredshopper")="YES" then 
	  if session("addressbookcustnum") = 0 then
			session("addressbookcustnum") = sitelink.makeaddressbook(session("LASTNAME"),session("ZIPCODE"),SESSION("PASSWORD"),session("shopperid"))
	  end if 
	  xmlAddressBookstring =sitelink.getaddressbook(session("addressbookcustnum"),0)
	    
		objDoc.loadxml(xmlAddressBookstring)
		set SL_AddressBook = objDoc.selectNodes("//gab")  

	%>
	<table width="100%"  border="0"  cellpadding="2" cellspacing="1">
	<tr>
			<th align="left"  width="30%" class="THHeader" >Customer Name</th>
			<th align="CENTER"  class="THHeader" >Address</th>
			</tr>
	<% for z=0 to SL_AddressBook.length-1 
		  SL_AddBookCustnum			= (SL_AddressBook.item(z).selectSingleNode("custnum").text)
		  SL_AddBookLname			= (SL_AddressBook.item(z).selectSingleNode("lastname").text)
		  SL_AddBookFname			= (SL_AddressBook.item(z).selectSingleNode("firstname").text)
		  
		  SL_AddBookCompany			= (SL_AddressBook.item(z).selectSingleNode("company").text)
		  SL_AddBookAddr			= (SL_AddressBook.item(z).selectSingleNode("addr").text)
		  SL_AddBookAddr2			= (SL_AddressBook.item(z).selectSingleNode("addr2").text)
		  SL_AddBookCity			= (SL_AddressBook.item(z).selectSingleNode("city").text)
		  SL_AddBookSate			= (SL_AddressBook.item(z).selectSingleNode("state").text)
		  SL_AddBookZipcode			= (SL_AddressBook.item(z).selectSingleNode("zipcode").text)
		  SL_AddBookcountry			= (SL_AddressBook.item(z).selectSingleNode("country").text)
		   SL_AddBookcounty			= SL_AddressBook.item(z).selectSingleNode("county").text
		  
		  SL_AddBookAddr_id			= (SL_AddressBook.item(z).selectSingleNode("address_id").text)					  
		  SL_AddBookshopcust_id	    = (SL_AddressBook.item(z).selectSingleNode("shopcust").text)
				if (z mod 2) = 0 then
					class_to_use = "tdRow1Color"
				else
					class_to_use = "tdRow2Color"
				end if 


	%>
	<tr>
		<td  class="<%=class_to_use%>">
		<span class="plaintextbold">
		&nbsp;<a href="changeshipto.asp?shopcust=<%=SL_AddBookshopcust_id%>&basketrecord=0"><font color="red">(Select)</font></a>
		&nbsp;<%=SL_AddBookLname%>&nbsp;&nbsp;<%=SL_AddBookFname%>
		<a class="allpage" href="EditAddressBook.asp?shopcust=<%=SL_AddBookshopcust_id%>&basketrecord=0"><font color="red">(Edit)</font></a>
		</span>
		
		</td>
		<td valign="top" class="<%=class_to_use%>">
		<span class="plaintext">
		<% if len(SL_AddBookCompany) > 0 then
			response.write(SL_AddBookCompany + ",&nbsp;")
		end if
		if len(SL_AddBookAddr) > 0 then
			response.write(SL_AddBookAddr + ",&nbsp;")
		end if
		if len(SL_AddBookAddr2) > 0 then
			response.write(SL_AddBookAddr2 + ",&nbsp;")
		end if
		
		if len(SL_AddBookAddr3) > 0 then
			response.write(SL_AddBookAddr3 + ",&nbsp;")
		end if
		
			response.write(SL_AddBookcity + ",&nbsp;")
		    if SL_AddBookcountry="073" then
		        if SL_AddBookcounty <>"" then
			        response.write(sitelink.Get_CountyName(SL_AddBookcounty) + ",&nbsp;")
		        end if
		    else
			    response.write(SL_AddBookSate + ",&nbsp;")
			end if 
			response.write(SL_AddBookZipcode + ",&nbsp;")
			response.write(sitelink.countryname(SL_AddBookcountry) )
		%>
		</span>
		</td>
	</tr>
	
	<%next%>
	</table>
	
	
	<%
		set SL_AddressBook = nothing
		'Set objDocAddBook = nothing		
		
	%>
	<%else%>
	
	
	<%end if%>
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
