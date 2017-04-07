<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">



<%
  if session("registeredshopper")="NO" then
  	response.redirect("statuslogin.asp")
  end if 

%>
<!--#INCLUDE FILE = "include/momapp.asp" -->

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->



<html>
<head>
<title>Aussie Products.com | Address Book | Australian Products Co.</title>
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

</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">

<%session("destpage")="address.asp" %>

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
		<tr><td class="TopNavRow2Text" width="100%" >Address Book 
		</td></tr>
	</table>
    <br>
	
	<%if session("registeredshopper")="YES" then 
	  if session("addressbookcustnum") = 0 then						
			session("addressbookcustnum") = sitelink.makeaddressbook(session("LASTNAME"),session("ZIPCODE"),SESSION("PASSWORD"),session("shopperid"))
	  end if 
	  xmlAddressBookstring =sitelink.getaddressbook(session("addressbookcustnum"),0)
	
		objDoc.loadxml(xmlAddressBookstring)
		set SL_AddressBook = objDoc.selectNodes("//gab")  
	%>
	<P align="left">&nbsp;&nbsp;<a href="AddNewAddressBook.asp">Click here to Add New Addresses</a>
		&nbsp;&nbsp;
	<span class="plaintext">&nbsp;Click on the <font color="red"><strong>(Edit)</strong></font> to Edit this shipping address.</span>

	</p>
	<%'=session("addressbookcustnum")%>
	<table width="100%"  border="0"  cellpadding="4" cellspacing="1">
	<tr>
			<th align="left"  width="30%" class="THHeader" >Customer Name</th>
			<th align="CENTER"  class="THHeader" >Address</th>
			</tr>
			
	<% if SL_AddressBook.length = 0 then %>
	 <tr><td class="plaintextbold" colspan="2" >
	 <br>
	 <ul>No addresses assigned..
	 <br><br>
	 Click "<a class="allpage" href="AddNewAddressBook.asp">Add New Addresses</a>" to add New Addresses to your address book</ul>
	 </td></tr>
	<%end if%>
	<% for z=0 to SL_AddressBook.length-1 
		  SL_AddBookCustnum			= SL_AddressBook.item(z).selectSingleNode("custnum").text
		  SL_AddBookLname			= SL_AddressBook.item(z).selectSingleNode("lastname").text
		  SL_AddBookFname			= SL_AddressBook.item(z).selectSingleNode("firstname").text
		  
		  SL_AddBookCompany			= SL_AddressBook.item(z).selectSingleNode("company").text
		  SL_AddBookAddr			= SL_AddressBook.item(z).selectSingleNode("addr").text
		  SL_AddBookAddr2			= SL_AddressBook.item(z).selectSingleNode("addr2").text
		  SL_AddBookAddr3			= SL_AddressBook.item(z).selectSingleNode("addr3").text
		  SL_AddBookCity			= SL_AddressBook.item(z).selectSingleNode("city").text
		  SL_AddBookSate			= SL_AddressBook.item(z).selectSingleNode("state").text
		  SL_AddBookZipcode			= SL_AddressBook.item(z).selectSingleNode("zipcode").text
		  SL_AddBookcountry			= SL_AddressBook.item(z).selectSingleNode("country").text
		  SL_AddBookcounty			= SL_AddressBook.item(z).selectSingleNode("county").text
		  
		  SL_AddBookAddr_id			= SL_AddressBook.item(z).selectSingleNode("address_id").text				  
		  SL_AddBookshopcust_id	    = SL_AddressBook.item(z).selectSingleNode("shopcust").text
		  
		 if (z mod 2) = 0 then
		   class_to_use = "tdRow1Color"
		else
			class_to_use = "tdRow2Color"
end if 


	%>
	<tr>
		<td class="<%=class_to_use%>">
		<span class="plaintextbold">
		&nbsp;<a class="allpage" href="EditAddressBook.asp?shopcust=<%=SL_AddBookshopcust_id%>"><font color="red">(Edit)</font></a>
		&nbsp;<%=SL_AddBookLname%>&nbsp;&nbsp;<%=SL_AddBookFname%>
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
	
	<table width="100%"  cellpadding="0" cellspacing="0" border="0">
	<tr><td class="plaintextbold" align="center">
	<br><br>
	You must Login to gain access to your address Book
	<br><br>
	<a class="allpage" href="login.asp">Click to Login</a>
	<br><br>
	<img alt="" src="images/clear.gif" width="1" height="200" border="0">
	</td></tr>
	</table>
	
	
	
	<%end if%>
	
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
