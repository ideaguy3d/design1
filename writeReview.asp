<%on error resume next%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<%
    current_stocknumber = cstr(REQUEST.QUERYSTRING("number"))
    current_stocknumber =fix_xss_Chars(current_stocknumber)
    
    current_variation = cstr(REQUEST.QUERYSTRING("variation"))
    current_variation =fix_xss_Chars(current_variation)
	
	isvalidstock = sitelink.validstocknumber2(current_stocknumber)
	
	if isvalidstock=-1 then
		set sitelink=nothing
		set ObjDoc=nothing
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl&"default.asp"
		Response.End
	end if

	SET LOSTOCK =  sitelink.setupproduct(cstr(current_stocknumber),cstr(current_variation),0,session("ordernumber"),"")
	
	producttitle = trim(LOSTOCK.inetsdesc)
	Popuptitle = fix_xss_Chars(producttitle)
	
	SET LOSTOCK =nothing
	

	If Request.TotalBytes > 0 Then
	   session("rating_val") = cint(request.Form("ReviewRating"))
	   session("ReviewTitle") = trim(cstr(request.Form("ReviewTitle")))
	   session("ReviewComment") = trim(cstr(request.Form("ReviewComment")))
	   session("ReviewerEmail") = trim(cstr(request.Form("ReviewerEmail")))
	   session("NoShowEmail") = request.Form("NoShowEmail").count
	   session("item_number") = request.Form("itemnumber")
	   session("firstname") = trim(cstr(request.Form("firstname")))
	   session("lastname") = trim(cstr(request.Form("lastname")))
	   session("state") = trim(cstr(request.Form("state")))
	   
	   session("rating_val") =  fix_xss_Chars(session("rating_val"))
	   session("ReviewTitle") =  fix_xss_Chars(session("ReviewTitle"))
	   session("ReviewComment") =  fix_xss_Chars(session("ReviewComment"))
	   session("ReviewerEmail") =  fix_xss_Chars(session("ReviewerEmail"))
	   session("NoShowEmail") =  fix_xss_Chars(session("NoShowEmail"))
	   session("item_number") =  fix_xss_Chars(session("item_number"))
	   session("firstname") =  fix_xss_Chars(session("firstname"))
	   session("lastname") =  fix_xss_Chars(session("lastname"))
	   session("state") =  fix_xss_Chars(session("state"))

	   'response.write(session("NoShowEmail"))
	   'response.write(session("ReviewTitle"))
	   'response.write(session("ReviewComment"))
	   'response.write(session("ReviewerEmail"))	   
	   if session("rating_val")  > 0 and len(trim(session("ReviewTitle"))) > 0 and len(trim(session("ReviewComment"))) > 0 then
	   	set sitelink=nothing
		SET LOSTOCK = nothing
	   	response.redirect("preview_custreview.asp")
	   end if
	 end if


%>







<html>
<head>
<title>Aussie Products.com | Write Customer Review @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
<script language="javascript" type="text/javascript">
function validateForm()
{
  var ProdRating = document.getElementById('ReviewRating').selectedIndex ;
  
  if (ProdRating==0)
  {
    alert("Please select rating from 'How do you rate this selection?'") ;
  }
  
  var ProdReviewTitle = document.getElementById('ReviewTitle').value ;
  
  if (ProdReviewTitle.length==0)
  {
    alert("Please enter review title") ;
  }
  
  var ProdReviewComment = document.getElementById('ReviewComment').value ;
  
  if (ProdReviewComment.length==0)
  {
    alert("Please enter review") ;
  }
    
  //alertmsg = "Thank you for submitting your review for " + prodtitle + ".\n\nComments are subject to moderation - your comment will appear on the site shortly." ;
  //alert(alertmsg);  
  return true ;

}
</script>

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
	<br>
	<table width="100%" cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%">Write a review for <%=producttitle%></td></tr>
	</table>
	<br>	
	<form action="<%=secureurl%>writeReview.asp?number=<%=current_stocknumber%>" method="post" onsubmit="return validateForm();">
	<table width="100%" BORDER="0" cellpadding="4" cellspacing="4">
	<tr><td class="plaintextbold" colspan="2">
	<%if session("rating_val")= "0" then %>
	<font color="red">How do you rate this item? &nbsp;</font>
	<%else%>
	How do you rate this selection? &nbsp;
	<%end if%>
	<select name="ReviewRating" id="ReviewRating" size="1">
		<option value="0" <%if session("rating_val")= "0" then response.write(" Selected") end if%>>&nbsp;-&nbsp;</option>
        <option value="5" <%if session("rating_val")= "5" then response.write(" Selected") end if%>>5 Stars</option>
        <option value="4" <%if session("rating_val")= "4" then response.write(" Selected") end if%>>4 Stars</option>
        <option value="3" <%if session("rating_val")= "3" then response.write(" Selected") end if%>>3 Stars</option>
        <option value="2" <%if session("rating_val")= "2" then response.write(" Selected") end if%>>2 Stars</option>
        <option value="1" <%if session("rating_val")= "1" then response.write(" Selected") end if%>>1 Star</option>
	</select>
	</td></tr>
	
	<tr><td class="plaintextbold" colspan="2">
	<%if len(trim(session("ReviewTitle"))) = 0 then %>
	<font color="red">Please enter a title for your review: </font>
	<%else%>
	Please enter a title for your review:
	<%end if%>
	</td></tr>
	<tr><td colspan="2"><input type="text" name="ReviewTitle" id="ReviewTitle" maxlength="70" size="60" value="<%=session("ReviewTitle")%>"></td></tr>
	<tr><td class="plaintextbold" colspan="2">
	<%if len(trim(session("ReviewComment"))) = 0 then %>
	<font color="red">Type your review in the space below: </font>
	<%else%>
	Type your review in the space below:
	<%end if%>
	
	</td></tr>
	<tr><td colspan="2">
		<textarea cols=43 rows=5 name="ReviewComment" id="ReviewComment" style="width:405px;"><%=session("ReviewComment")%></textarea>
	</td></tr>
	<tr><td width="175" class="plaintextbold"><strong>Your First Name (optional):</strong></td><td><input type="text" name="firstname" maxlength="15" size="15" value="<%=session("firstname")%>"></td></tr>
	<tr><td class="plaintextbold"><strong>Your Last Name (optional):</strong></td><td><input type="text" name="lastname" maxlength="20" size="20" value="<%=session("lastname")%>"></td></tr>
	<tr><td class="plaintextbold"><strong>Your email (optional):</strong></td><td><input type="text" name="ReviewerEmail" maxlength="50" size="39" value="<%=session("ReviewerEmail")%>"></td></tr>
		<tr><td class="plaintextbold"><strong>State (optional):</strong></td><td>
			<select name="state" class="smalltextblk">
            <option value="">
			<option value="INT">INTERNATIONAL
			<%
			xmlstring =sitelink.Get_state_list(session("ordernumber"))
			objDoc.loadxml(xmlstring)
			set SL_Ship_state = objDoc.selectNodes("//state_list")
				for x=0 to SL_Ship_state.length-1
			  	SL_state_name= SL_Ship_state.item(x).selectSingleNode("statename").text
			  	SL_state_code= SL_Ship_state.item(x).selectSingleNode("sstate").text
			%>
			<option value="<%=SL_state_code%>"
			<%'if session("state") = SL_state_code then response.write(" selected") end if%>
			><%=SL_state_name%>
			<%next						  
			set SL_Ship_state=nothing
			%>
			</select></td></tr>


	<tr><td colspan="2" style="padding-left:185px"><br> 
             <input type="hidden" name="itemnumber" value="<%=trim(current_stocknumber)%>">
             <input type="submit" name="PreviewBtn" value="Preview Your Review" >
    </td></tr>
	</table>
	</form>

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
