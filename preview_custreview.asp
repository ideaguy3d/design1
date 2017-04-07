<%on error resume next%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<% 
 
 if session("rating_val")  = 0 or len(trim(session("item_number"))) = 0 or len(trim(session("ReviewTitle"))) = 0 or len(trim(session("ReviewComment"))) = 0 then
  set sitelink= nothing
  response.redirect("writeReview.asp")
 end if
 
SET LOSTOCK =  sitelink.setupproduct(cstr(session("item_number")),cstr(""),0,session("ordernumber"),"") 

producttitle = trim(LOSTOCK.inetsdesc)
Popuptitle = fix_xss_Chars(producttitle)
	

%>




<html>
<head>
<title>Aussie Products.com | Customer Order Review @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
<% session("destpage")="" 
	session("viewpage")=""
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	   
		<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
			<!--#INCLUDE FILE = "include/side_nav.asp" -->			
		</td>	
		

	    
		<td valign="top" class="pagenavbg">

		<!-- sl code goes here -->
<div id="page-content">		
		<table width="95%" cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%">Your review of&nbsp;<%=trim(LOSTOCK.inetsdesc)%></td></tr>
	</table>
	<br>
	
	<hr size="1" color="#10568D">
	<table width="100%" BORDER="0" cellpadding="3" cellspacing="0">

	
	<%
	
	  usethis_rating_img = "images/clear.gif"
	  if session("rating_val") > 0 then
			usethis_rating_img = "images/"&session("rating_val")&+"note.gif"
	  end if
	 
	%>
	
	<tr><td class="plaintextbold"><img src="<%=usethis_rating_img%>" width="70" height="14" border="0">&nbsp;&nbsp;&nbsp;<%=session("ReviewTitle")%></td></tr>
	<tr><td class="plaintext"><!--Review:  
		<br>-->
		<%=session("ReviewComment")%>
		</td>
	</tr>
		<%
		footer_txt=""
		if len(trim(session("firstname"))) > 0 then
			footer_txt = trim(session("firstname"))
		end if
		
		if len(trim(session("lastname"))) > 0 then
			if len(trim(footer_txt)) > 0 then
				footer_txt = footer_txt & "&nbsp;"
			end if
			footer_txt = footer_txt & trim(session("lastname"))
		end if
		
		if len(trim(session("state"))) > 0 then
			if len(trim(footer_txt)) > 0 then
				footer_txt = "-&nbsp;" & footer_txt & ",&nbsp;"
			end if
			footer_txt = footer_txt & trim(session("state"))
		end if
		
		if len(trim(footer_txt)) > 0 then
		%>
		<tr><td class="plaintext" style="font-size:11px"><i><%=footer_txt%></i></td></tr>
		<%end if%>
	</table>
	<hr size="1" color="#10568D">
	<table width="15%" BORDER="0" cellpadding="3" cellspacing="0">
	<tr><td><form action="submit_review.asp?number=<%=session("item_number")%>" method="post">
	<input type="submit" value="Submit" name="submit"
	onclick="Javascript:alert('Thank you for submitting your review for the <%=Popuptitle %>.\n\nComments are subject to moderation - your comment will appear on the site shortly.');"
	></form></td>
	<td class="plaintextbold" valign="top">OR</td>
	<td width="40%" valign="top"><a href="javascript:history.go(-1)" method="post"><input type="button" value="Edit" ID="Submit1" NAME="Edit" onClick="javascript:history.go(-1)"></a></td>
	</tr>
	</table>
	<span class="plaintext">Product reviews are updated/posted every weekday morning, thank you.</span>
	
	<%set LOSTOCK=nothing %>
	

		<br><br>
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
