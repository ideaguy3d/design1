<%on error resume next%>
<!--#INCLUDE FILE = "include/AdminCreateSL.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<!--#INCLUDE FILE = "text/storetext.asp" -->


<html>
<head>
<title><%=althomepage%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->

<center>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td class="topnav1bgcolor">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top"><a href="<%=insecureurl%>"><img alt="<%=session("storename")%>" src="images/<%=COMPANY_LOGO_IMG%>" border="0" ></a></td>		
		<td align="right" valign="top">
			<table width="223" border="0" cellpadding="0" cellspacing="0">
			<tr><td height="5" colspan="5"><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td></tr>
			<tr><td colspan="5">&nbsp;</td></tr>
			<tr><td height="25" colspan="5"><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td></tr>
			<tr>
				<td class="plaintext"><a title="View Basket" href="basket.asp"><img src="images/newcart_WhtImg.gif" alt="View Basket" style="width:23;height:12;border:0"></a></td>
				<td><img src="images/clear.gif" alt="" style="width:5;height:1;border:0"></td>
				<td style="background-image:url(images/bkg_cartinfo.gif);height:21">
				<table width="200" border="0" cellpadding="0" cellspacing="0" >
				<tr>
				<td class="plaintext" align="center"><a title="View Basket" href="basket.asp"><span class="plaintext"><u>Items</u></span></a>&nbsp;<%=session("SL_BasketCount")%>
				&nbsp;|&nbsp;
				Sub Total&nbsp;<%=formatcurrency(session("SL_BasketSubTotal"))%></td>					
				</tr>
				</table>
			</td>
			<td>&nbsp;</td>
			<td><a class="topnav1" title="Checkout" href="<%=secureurl%>custinfo.asp" ><u>Checkout</u></a></td>
			</tr>			
			</table>		
		</td>
		<td style="width:10"><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
	</tr>
	</table>
</td>
</tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td bgcolor="#A0A0A0" height="1"><img src="images/clear.gif" alt="" width="1" height="1" border="0"></td></tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr><td  class="topnav2bgColor" height="23">
	<center>
	<table width="0" cellpadding="2" cellspacing="0" border="0">
	<tr>
		<td><a href="<%=insecureurl%>" title="<%=ALT_HOME_TXT%>" class="topnav2"><%=ALT_HOME_TXT%></a></td>
		<td><img src="images/clear.gif" alt="" style="width:5;height:1;border:0"></td>
		<td><img src="images/Sep.jpg" alt="" style="width:5;height:23;border:0"></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td><a href="<%=insecureurl%>Aboutus.asp" title="<%=ALT_ABOUT_TXT%>" class="topnav2"><%=ALT_ABOUT_TXT%></a></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td><img src="images/Sep.jpg" alt="" style="width:5;height:23;border:0"></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td ><a href="<%=insecureurl%>sitemap.asp" title="Site Map" class="topnav2">Site Map</a></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td><img src="images/Sep.jpg" alt="" style="width:5;height:23;border:0"></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td ><% if session("registeredshopper")="NO" then %><a href="<%=secureurl%>login.asp" title="Login" class="topnav2">Account Login</a><%else%><a href="<%=secureurl%>logout.asp" title="Logout" class="topnav2">Logout</a><%end if%></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td><img src="images/Sep.jpg" alt="" style="width:5;height:23;border:0"></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td ><a href="<%=insecureurl%>testimonials.asp" title="<%=ALT_TESTIMONIAL_TXT%>" class="topnav2"><%=ALT_TESTIMONIAL_TXT%></a></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td><img src="images/Sep.jpg" alt="" style="width:5;height:23;border:0"></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td ><a href="<%=insecureurl%>Quick_order.asp" title="Quick Order" class="topnav2">Quick Order</a></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td><img src="images/Sep.jpg" alt="" style="width:5;height:23;border:0"></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<%if WANT_WISHLIST = 1 then %>
	   <td><a href="<%=insecureurl%>wishlist.asp" title="Wish List" class="topnav2">Wish List</a></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td><img src="images/Sep.jpg" alt="" style="width:5;height:23;border:0"></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
	   <%end if%>
	    <td ><a href="<%=insecureurl%>privacy.asp" title="<%=ALT_PRIVACY_TXT%>" class="topnav2"><%=ALT_PRIVACY_TXT%></a></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td><img src="images/Sep.jpg" alt="" style="width:5;height:23;border:0"></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>		
		<td ><a href="<%=insecureurl%>contactus.asp"  title="<%=ALT_CONTACT_TXT%>" class="topnav2"><%=ALT_CONTACT_TXT%></a></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
		<td><img src="images/Sep.jpg" alt="" style="width:5;height:23;border:0"></td>
		<td><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td>
	   <td ><a href="<%=insecureurl%>faq.asp" title="<%=ALT_FAQ_TXT%>" class="topnav2"><%=ALT_FAQ_TXT%></a></td>
	</tr>	
	</table>
	</center>
	</td>
</tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td bgcolor="#A0A0A0" height="1"><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td></tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td class="topnav3bgcolor" height="15" align="center">
	
    <img src="images/clear.gif" alt="" style="width:15;height:1;border:0"><a class="topnav3" href="restoreoldorder.asp">Restore Last Order</a>
 
</td>
<td valign="top" align="right" class="topnav3bgcolor">
		<table border="0">
			<tr>
				<td class="TopNav3Text"><b>Search</b></td>
				<td><select name="ProductSearchBy" class="smalltextblk">
					<option value="2" <%if session("searchcriteria") ="DESC" then response.write(" selected") end if%> >Description</option>
					<option value="1" <%if session("searchcriteria") ="PARTNUMBER" then response.write(" selected") end if%> >Item Number</option>
					<% if USE_ADVANCED_SEARCH= 1 then %>
						<% if len(trim(session("SL_Advanced1"))) > 0  then %>
							<option value="3" <%if ucase(session("searchcriteria")) =ucase(session("SL_Advanced1")) then response.write(" selected") end if%> ><%=session("SL_Advanced1")%></option>
						<%end if%>
						<% if len(trim(session("SL_Advanced2"))) > 0 then %>
							<option value="4" <%if ucase(session("searchcriteria")) =ucase(session("SL_Advanced2")) then response.write(" selected") end if%> ><%=session("SL_Advanced2")%></option>
						<%end if%>
						<% if len(trim(session("SL_Advanced3"))) > 0 then %>
							<option value="5" <%if ucase(session("searchcriteria")) =ucase(session("SL_Advanced3")) then response.write(" selected") end if%> ><%=session("SL_Advanced3")%></option>
						<%end if%>
						<% if len(trim(session("SL_Advanced4"))) > 0 then %>
							<option value="6" <%if ucase(session("searchcriteria")) =ucase(session("SL_Advanced4")) then response.write(" selected") end if%> ><%=session("SL_Advanced4")%></option>
						<%end if%>
					<%end if%>
					</select>
				</td>
				<td><input type="text" size="18" maxlength="256" class="plaintext" name="txtsearch" value="<%=replace(session("searchstring"),"""","")%>"></td>
				<td><input type="image" src="images/btn_go.gif" style="height:19;width:25;border:0"></td>
			</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td bgcolor="#A0A0A0" height="1"><img src="images/clear.gif" alt="" style="width:1;height:1;border:0"></td></tr>
</table>

<br><br>

<table width="0" cellpadding="10" cellspacing="10">
<tr><td  valign="top">
	<font color="black" face="Verdana" size="2">
	<b>Top Nav1 Links</b>
	<br><br>
	<li>Checkout <br>	
	
	
	</font>

	</td>
	<td  valign="top">
	<font color="black" face="Verdana" size="2">
	<b>Top Nav2 Links</b>
	<br><br>
	<li>Home<br>
	<li>About us <br>
	<li>Account Login <br>
	<li>Testimonials <br>
	<li>Quick Order <br>
	<li>Wish List <br>
	<li>Privacy Policy <br>
	<li>Contact Us <br>
	<li>Faq
	</font>

	</td>
	
	<td  valign="top">
	<font color="black" face="Verdana" size="2">
	<b>Top Nav3 Links</b>
	<br><br>
	<li>Restore Last Order <br>
	</font>
	
	</td>
	
	<td  valign="top">
	<font color="black" face="Verdana" size="2">
	<b>Top Nav3 Text</b>
	<br><br>
	<li>Search <br>
	</font>
	
	</td>
	
	
</tr>
</table>

<br><br>
<a href="Javascript:window.close()"><span class="CompPrice">Close Window</span></a>

</center>





<!--#INCLUDE FILE = "include/AdminReleaseSL.asp" -->
