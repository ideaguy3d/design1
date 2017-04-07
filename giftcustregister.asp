<%on error resume next%>


<%
   ' already logged in then go directly to display wishlist
	if len(trim(session("wishlistID"))) > 0 and session("registeredgiftreg")="YES" then
      response.redirect("Display_wishlist.asp")
	end if 

%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Customer Gift Registry @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
	<table width="95%" cellpadding="3" cellspacing="0" border="0">
		<tr><td class="breadcrumbs" width="100%">
				<a href="/">Home</a><span class="breadcrumb-divide">&nbsp;&nbsp;&#47;&nbsp;&nbsp;</span>Create a Registry
			</td>
		</tr>
	</table>
<table width="100%" border="0">
	 <tr> 
      <td align="left" width="" valign="TOP"> 
	  <form name="custinfoform" onSubmit="return validateFrm();" method="POST" action="giftcustregistration.asp">
	  <input type="hidden" name="cdate" value="<%=Date()%>">
      <table width="95%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td class="plaintextbold" align="center">
			<p align="left"><br>
			</td>
        </tr>
        <tr>
          <td width="100%" align="left" valign="top"> 
            <table width="100%" cellpadding="0" cellspacing="1" border="0">
			 <tr><td colspan="3" align="center" class="plaintext">
				<p align="left"></td>
			</tr>
			 <tr><td colspan="3" align="right" class="plaintext">
			  <p align="left">
			  <strong><font color="#ff0000">*</font></strong><b>
				<i>Required Fields</i></b></td>
			</tr>
			<tr>
			  <td align="center" colspan="2" class="THHeader">
				<p align="left"><b>Registrant's Information</b></td>
			</tr>
			<tr>
			  <td align="right" width="150" class="plaintext">
				<p align="left">First Name<b><font color="red">*</font></b>  </td>
			  <td width="374" align="left" class="plaintext"> 
			    <font size="1" face="Arial"> </font>
			    <input type="text" name="txtfirstname" size="23,1" maxlength="15" value="">
			  	
			  </td>
			</tr>
			<tr> 
			  <td align="right" width="150" class="plaintext">
				<p align="left">Last Name<b><font color="red">*</font>
			  </b> </td>
			  <td width="374" align="left" class="plaintext"> 
			    <font face="Arial"></font><font size="1"> </font>
			    <input type="text" name="txtlastname" size="33,1" maxlength="20" value=""></td>
			</tr>
			<tr><td colspan="4"><hr color=#a9a9a9 SIZE=1 length="430" align="left"></td></tr>
			<tr>
			  <td align="center" colspan="2" class="THHeader">
				<p align="left"><b>Co-Registrant's Information</b></td>
			</tr>
			<tr>
			  <td align="right" width="150" class="plaintext">
				<p align="left">First Name  </td>
			  <td width="374" align="left" class="plaintext"> 
			    <font size="1" face="Arial"> </font>
			    <input type="text" name="txtreg2fname" size="23,1" maxlength="15" value="">
			  	
			  </td>
			</tr>
			<tr> 
			  <td align="right" width="150" class="plaintext">
				<p align="left">Last Name  </td>
			  <td width="374" align="left" class="plaintext"> 
			    <font size="1" face="Arial"> </font>
			    <input type="text" name="txtreg2lname" size="33,1" maxlength="20" value="">
			  	
			  </td>
			</tr>
			<tr><td colspan="4"><hr color=#a9a9a9 SIZE=1 align="left"></td></tr>
			<tr> 
			  <td align="center" colspan="2" class="THHeader">
				<p align="left"><b>Event Information</b></td>
			</tr>
			<tr>
			  <td align="right" width="150" class="plaintext">
				<p align="left">Type of Registry<b><font color="red">*</font>
			  </b></td>
			  <td width="374" align="left" class="plaintext"> 
			    <font size="1" face="Arial"> </font>
			    <select name="txtregistry">
			        <option value="birth">birth
					<option value="birthday">birthday
			        <option value="shower">shower
			        <option value="adoption">adoption
					<option value="back to School">back to School
					<option value="christening">christening
					<option value="christmas">christmas
					<option value="hanukkah">hanukkah
					<option value="other">other
			    </select>
			    
			  </td>
			</tr>
			<tr> 
			  <td align="right" width="150" class="plaintext">
				<p align="left">Date of Event<b><font color="red">*</font>
			  </b></td>
			  <td width="374" align="left" class="plaintext"> 
				<font face="Arial"><font size="1"> </font></font>
				<input type="text" name="txteventdate" size="12" maxlength="10" value=""></td>
			</tr>
			<tr> 
			  <td align="right" valign="top" width="150" class="plaintext">
				<p align="left">Registry Description<b><font color="red">*</font>
			  </b></td>
			  <td width="374" align="left" class="plaintext"> 
			    <font size="1" face="Arial"> </font>
			    <input type="text" name="txtregdesc" size="40,1" maxlength="70" value=""><br>
				(A short description of your registry/ a message for your registry viewers)
			  	
			  </td>
			</tr>
			<tr><td colspan="4"><hr color=#a9a9a9 SIZE=1 align="left"></td></tr>
			<tr>
			  <td align="center" colspan="2" class="THHeader">
				<p align="left"><b>Shipping Address</b></td> 
			</tr>
			<tr>
			  <td align="right" width="150" class="plaintext"> 
				<p align="left">Company  </td>
			  <td width="374" align="left" class="plaintext"> 
			    <font size="1" face="Arial"> </font>
			    <input type="text" name="txtcompany" size="50,1" maxlength="40" value="">
			  	
			  </td>
			</tr>
			<tr> 
			  <td align="right" width="150" class="plaintext"> 
				<p align="left">Address&nbsp;<b><font color="red">*</font>
			  </b> </td>
			  <td width="374" align="left" class="plaintext"> 
			    <font face="Arial"><font size="1"> </font></font>
			    <input type="text" name="txtaddress1" size="50,1" maxlength="40" value=""><font size="1"></font>
				
			  </td>
			</tr>
			<tr> 
			  <td align="right" width="150" class="plaintext"> 
				<p align="left"> </td>
			  <td width="374" align="left" class="plaintext"> 
			    <font size="1" face="Arial"> </font>
			    <input type="text" name="txtaddress2" size="50,1" maxlength="40" value="">
			  	
			  </td>
			</tr>
			<tr> 
			  <td align="right" width="150" class="plaintext" valign="top">
				<p align="left">
				City 
				<b>
			    <font color="red">*</font></b></p>
				</td>
			  <td width="374" align="left" class="plaintext"> 
			    <font face="Arial"><font size="1"> </font></font>
			    <input type="text" name="txtcity" size="25,1" maxlength="20" value=""> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			    State<b><font color="red">* </font>
			  </b> <font face="Arial"><font size="1"></font> </font>
			    &nbsp;<input type="text" name="txtstate" size="5,1" maxlength="3" value=""><br>
				Zip Code<b><font color="red">* </font>
			  </b> &nbsp; <nobr>
				<font face="Arial"><font size="1"></font></font><input type="text" name="txtzipcode" size="10,1" maxlength="10" value=""></td>
			</tr>
			<tr> 
			  <td align="right" width="150" class="plaintext"> 
				<p align="left">Country  </td>
			  <td width="374" align="left" class="plaintext"> 
			    <font size="1" face="Arial"> </font>
			    
			    <select NAME="txtcountry" class="smalltextblk">
			<%=sitelink.listcountries(session("ordernumber"),session("shopperid"),"B",0)%>
			</select>
			  
			  </td>
			</tr>
			<tr> 
			  <td align="right" width="150" class="plaintext"> 
				<p align="left">Phone<b><font color="red">*</font></b>  </td>
			  <td width="374" align="left" class="plaintext"> 
			    <font face="Arial"><font size="1"> </font></font>
			    <input type="text" name="txtphone" size="20,1" maxlength="14" value=""></td>
			</tr>
			<tr><td colspan="4"><hr color=#a9a9a9 SIZE=1 align="left"></td></tr>
			<tr> 
			  <td align="center" colspan="2" class="THHeader">
				<p align="left"><b>Login Information</b></td>
			</tr>
			<tr>
			  <td align="right" width="150" class="plaintext"> 
				<p align="left">E-Mail<b><font color="red">*</font></b>  </td>
			  <td width="374" align="left" class="plaintext"> 
			    <font face="Arial"><font size="1"> </font></font>
			    <input type="text" name="txtemail" size="50,1" maxlength="50" value=""><font size="1"></font>
				
			  </td>
			</tr>
			<tr> 
				<td align="right" width="150" class="plaintext">
				<p align="left">Registry Password<b><font color="red">*</font></b></td>
				<td width="374" align="left" class="plaintext">
					<font face="Arial"><font size="1"></font></font>
					<input type="password"  onFocus="style.background='#FFFFFF';" name="txtregpassword" size="15" maxlength="10" value="">
					</td>
			</tr>
			<tr>
				<td align="right" width="150" class="plaintext">
				<p align="left">Re-Enter Registry Password<b><font color="red">*</font></b></td>
				<td width="374" align="left" class="plaintext">
					<font size="1" face="Arial"></font>
					<input type="password"  onFocus="style.background='#FFFFFF';" name="txtregpassword2" size="15" maxlength="10" value="">
					 </td>
			</tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td align="center" colspan="3">
				<p><br><input type="Image" value="Register" src="images/btn-register.gif" border="0" alt="Register" id=Image1 name=Image1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			</tr>	  
			  
			  
            </table>
			
      </table>
	  </form>
  </tr>
</table>
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
