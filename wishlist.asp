<%on error resume next%>

<%
 session("destpage")="wishlist.asp" 
  if session("registeredshopper")="NO" then
  	response.redirect("login.asp")
  end if 

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<html>
<head>
<title>Aussie Products.com | Wish list | Australian Products Co.</title>
<meta name="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls" />
	<meta name="keywords" content="Aussie Products, Australian, Aussie, Australian Products, Australian Food, Aussie Foods, Australian Desserts, tim tams, vegemite, violet crumble, cherry ripe, koala, kangaroo, down under, Australian Groceries, foods, lollies, candy, gourmet food" />
	<meta name="robots" content="index, follow" />
    <meta name="robots" content="ALL">
	<meta name="distribution" content="Global">
	<meta name="revisit-after" content="5">
	<meta name="classification" content="Food">	
    <link rel="shortcut icon" href="favicon.ico" type="image/x-icon" /><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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
	<br>
    <center>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%" >Wish List</td></tr>
	</table>
	
	<%if session("registeredshopper")="YES" then 
	  
	 extra_field = "LEFT(JUSTFNAME(stock.inetimage)+SPACE(50),50) AS inetimage "
	  xmlstring =sitelink.WISHLIST("LIST","0",0,0,"",extra_field,session("addressbookcustnum"))

		  objDoc.loadxml(xmlstring)
		set SL_Wishlist = objDoc.selectNodes("//wls")	  				 

	%>  
	<br>
	<form action="updatewishlist.asp" method="post">
	<table width="100%"  border="0"  cellpadding="3" cellspacing="1">
		<tr>
			<th align="CENTER" class="THHeader" width="15%"><span class="THHeader">Item #</span></th>
			<th align="CENTER" class="THHeader" width="15%">Image</th>
			<th align="CENTER" class="THHeader" width="60%"><span class="THHeader">Description</span></th>
			<th align="CENTER" class="THHeader" width="10%"><span class="THHeader">Price</span></th>
            <th align="CENTER" class="THHeader"><span class="THHeader">Remove</span></th>
		</tr>
	<% if SL_Wishlist.length = 0 then %>
	<tr><td class="plaintextbold" colspan="5" >
	 <br>
	 <ul>Your Wishlist is Empty...
	 <br><br>
	 </ul>
	 </td></tr>
	
	<%else%>
		<% for x=0 to SL_Wishlist.length-1 
		 SL_Number	= SL_Wishlist.item(x).selectSingleNode("item").text
		 SL_item_id	= SL_Wishlist.item(x).selectSingleNode("item_id").text
		 SL_qty		= SL_Wishlist.item(x).selectSingleNode("qty_want").text
		 SL_ord		= SL_Wishlist.item(x).selectSingleNode("qty_ord").text
		 SL_ItemDesc= SL_Wishlist.item(x).selectSingleNode("inetsdesc").text
		 SL_ItemDesc2= SL_Wishlist.item(x).selectSingleNode("desc2").text		 
		 SL_ItemPrice= SL_Wishlist.item(x).selectSingleNode("price1").text
		 SL_thumbnail= SL_Wishlist.item(x).selectSingleNode("inetthumb").text
		 SL_Fullimage  = SL_Wishlist.item(x).selectSingleNode("inetimage").text
		 
		 SL_StAttibprice  = cdbl(SL_Wishlist.item(x).selectSingleNode("price_wopt").text)
		 SL_StAttibNonSkuprice  = cdbl(SL_Wishlist.item(x).selectSingleNode("nskuprice").text)
		 SL_WishlistCinfo  = SL_Wishlist.item(x).selectSingleNode("custominfo").text
		 

		 SL_ItemPrice = cdbl(SL_ItemPrice) + SL_StAttibprice + SL_StAttibNonSkuprice
		 
		 'MyArray = Split(SL_qty, ".", -1, 1)
		 'numeric_part = MyArray(0)
		 'decimal_part = MyArray(1)
				
		 'if decimal_part > 0 then
		'	SL_qty = numeric_part +"." + decimal_part
		' else
		'	SL_qty = numeric_part				
		'end if
		
		SL_qty = replace(SL_qty,".00","")
		 
		 if (x mod 2) = 0 then
			 class_to_use = "tdRow1Color"
		else
			 class_to_use = "tdRow2Color"
		end if 
		
		if len(trim(SL_thumbnail)) = 0 then
			SL_thumbnail = SL_Fullimage
		end if


		StrFileName = "images/" + trim(SL_thumbnail)
		StrPhysicalPath = Server.MapPath(StrFileName)
		set objFileName = CreateObject("Scripting.FileSystemObject")	
		if objFileName.FileExists(StrPhysicalPath) then
			inethumbImage=StrFileName
		else
			inethumbImage="images/noimage.gif"
	    end if
		 set objFileName = nothing
		 
		 altag =trim(SL_ItemDesc)
				
		%>
		<tr>
        <td class="<%=class_to_use%> plaintext" valign="middle" align="center"><%=SL_Number%>
		</td>
        <td class="<%=class_to_use%>" valign="middle" align="center"><div class="prodthumb"><div class="prodthumbcell"><img src="<%=inethumbImage%>" alt="<%=altag%>" class="prodlistimg" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %>border="0" class="prodlistimg"></div></div></td>
        <td class="<%=class_to_use%> plaintext" valign="top"><br><%=SL_ItemDesc%>
		<% if len(trim(SL_ItemDesc2)) > 0 then response.write("<br>" +SL_ItemDesc2) end if %>
		<br>
		<%=SL_WishlistCinfo%>
        </td>
        <td class="<%=class_to_use%> ProductPrice" align="center" valign="top">
            <table width="100%" cellpadding="3" cellspacing="0">
                <tr><td><br><br><%=formatcurrency(SL_ItemPrice)%><br><br></td></tr>
                <tr><td><span class="plaintext">	
                <input type="text" name="qty" size="3" class="smalltextblk" maxlength="4" value="<%=SL_qty%>">
            </span><br></td></tr>
                <tr><td><a  class="allpage" href="addtocart.asp?item=<%=SL_Number%>&qty=<%=SL_qty%>"><img src="images/btn-buy.gif"></a></td></tr>
            </table>
        </td>
		<td valign="middle" align="center" class="<%=class_to_use%>">			  
					<%remname  ="remove" & cstr(x+1)%>
					<input type="checkbox" name="<%=remname%>" value="0">
					<input type="hidden" name="item" value="<%=SL_Number%>">
					<input type="hidden" name="item_id" value="<%=SL_item_id%>">
		</td>
		
		</tr>
		<%next%>

		<tr><td colspan="5" align="center">
		<br>
		<table width="100%" border="0">
		<tr>
		    <td align="center"><a href="<%=insecureurl%>"><img SRC="images/btn-continueshopping.gif" style="border:0" ALT="Continue Shopping"></a></td>
			<td align="center">
			<input type="image" src="images/btn_UpdateWishlist.gif" border="0" alt="Update Wishlist">
			</td>
		</tr>
		</table>		
		</td></tr>
		
	
	<%end if%>
	</table>
	</form>
	 <% set SL_Wishlist = nothing
	 	'Set objDoc = nothing
	 %>
	 
	<%else%>
	
	<table width="100%"  cellpadding="0" cellspacing="0" border="0">
	<tr><td class="plaintextbold" align="center">
	<br><br>
	You must Log in to gain access to your Wishlist
	<br><br>
	<a class="allpage" href="login.asp">Click to Login</a>
	<br><br>
	</td></tr>
	</table>
	
	
	
	<%end if%>
	</center>
	</div>
<!-- end sl_code here -->
	<img alt="" src="images/clear.gif" width="1" height="200" border="0">
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
