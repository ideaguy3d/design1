<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">



<%
    session("destpage")="custinfo.asp" 
	if session("SL_BasketCount")= 0 then		
		response.redirect("basket.asp")
	end if
	
%>

<!--#INCLUDE FILE = "include/momapp.asp" -->

<%
	
	if REQUEST.servervariables("server_port_secure")<> 1 and SLsitelive=true then 
	    set sitelink=nothing
        set ObjDoc=nothing		
    	response.redirect(secureurl+"custinfo.asp") 
    end if
  
	if REGISTER_BEFORE_CHECKOUT =1 and  session("registeredshopper") = "NO"  then		
			set sitelink=nothing 			
			response.redirect("login.asp")		
	end if
	
	if session("BILLTOCOPY")=1 then	
		session("sfirstname")=session("firstname")	
		session("slastname")=session("lastname")
		session("scompany")=session("company")
		session("saddress1")=session("address1")	
		session("saddress2")=session("address2")	
		session("saddress3")=session("address3")
		session("scity")=session("city")
		session("sstate")=session("state")
		session("szipcode")=session("zipcode")
		session("scounty")=session("bcounty")
		session("scountry")=session("country")
		session("sphone")=session("phone")
		session("semail")=session("email")
	end if 
	
	
	'response.write("---"&session("scounty"))
	'response.write("<br>---"&session("bcounty"))
	
	show_baddr3=false
	show_saddr3=false
	
	 billcountry = session("country")
	 if len(trim(billcountry)) = 0 then
	 	billcountry= session("SL_START_COO")
	 end if
	 shipcountry = session("scountry")
	 if len(trim(shipcountry)) = 0 then
	 	shipcountry= session("SL_START_COO")
	 end if
	 
	if billcountry="073" then
	    show_baddr3=true	    
	end if
    if shipcountry="073" then
        show_saddr3=true
    end if 
    
    txtstatelabelTxt="State"
    txtsstatelabelTxt = "State"
    if billcountry ="034" then
        txtstatelabelTxt = "Province"
    end if 
	
	if shipcountry="034" then
        txtsstatelabelTxt = "Province"
    end if 
    	
    txtbZiplabelTxt="Zip Code"
    txtsZiplabelTxt="Zip Code"
	if billcountry="073" or billcountry="034" then
	    txtbZiplabelTxt="Postal Code"
	end if 
	if shipcountry="073" or shipcountry="034" then
	    txtsZiplabelTxt="Postal Code"
	end if
	    
	
	
%>


<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<html>
<head>
<title>Aussie Products.com | Customer Information | Australian Products Co.</title>
<meta name="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls" />
	<meta name="keywords" content="Aussie Products, Australian, Aussie, Australian Products, Australian Food, Aussie Foods, Australian Desserts, tim tams, vegemite, violet crumble, cherry ripe, koala, kangaroo, down under, Australian Groceries, foods, lollies, candy, gourmet food" />
	<meta name="robots" content="index, follow" />
    <meta name="robots" content="ALL">
	<meta name="distribution" content="Global">
	<meta name="revisit-after" content="5">
	<meta name="classification" content="Food">	
    <link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
<link rel="stylesheet" href="text/footernav.css" type="text/css">
<link rel="stylesheet" href="text/design.css" type="text/css">
<script type="text/javascript" language="JavaScript" src="custinfo.js"></script>
</head>

<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">

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

<br>
	
	
	<table width="100%"  cellpadding="0" cellspacing="0" border="0">
		<tr>
            <td align="left" width="100%"><img src="images/step1.gif" border="0" width="403" height="29"></td>
          </tr>
	</table>
	<br>	
	<table width="100%"  border="0" cellspacing="0" cellpadding="1">
		  <tr><td class="plaintext">		  
			<p>We need some information before we can ship your order.  Please enter your address information below and then press the 'Continue' button.</p>
		  </td></tr>
		  
		  <%if  session("registeredshopper")<>"YES" then%>
		  <tr><td class="plaintext">If you are already a registered customer, please <a href="login.asp">login</a>.</td></tr>
		  
		  <%end if%>
		  <tr><td class="plaintext">
			If you want to change your order before proceeding, go to the <a class="allpage" href="basket.asp">shopping basket.</a>
		  </td></tr>
		 </table>
		<br>
		<form name="custinfoform" method="POST" onSubmit="return validateForm(<%=BILL_TOPHONE_REQUIRED%>,<%=SHIP_TOEMAIL_REQUIRED%>)" action="<%=secureurl%>initorder.asp">
		<table width="100%" cellpadding="0" cellspacing="0" border="0">
			<tr><td width="50%" valign="top">
				 <table width="100%" cellpadding="0" cellspacing="0" border="0">
					 <tr>
						 <td class="THHeader">&nbsp;&nbsp;Billing Address &nbsp;<span class="smalltextblk">(* Required)&nbsp;&nbsp; </span><a name="billtab"></a></td>						 
					 </tr>
					 <tr>
					 	<td valign="top">
						<table width="100%" cellpadding="0" cellspacing="0" border="0">	
			 			<tr>
							
							<td valign="top" width="5"><img src="images/clear.gif" alt="" border="0" width="5" height="1"></td>
						    <td width="100%" valign="top">
							<%if  session("registeredshopper")="YES" and WANT_ADDRESS_CHANGE=1 then%>
								<table width="100%" border="0" cellspacing="2" cellpadding="3">
								<tr><td colspan="2"><img src="images/clear.gif" height="30" width="1" border="0"></td></tr>
								<tr><td class="plaintextbold" align="right"> &nbsp;First Name&nbsp; </td>
									<td class="plaintext"><b><%=session("firstname")%></b>&nbsp;
									<input type="hidden" name="txtfirstname" id="txtfirstname" value="<%=session("firstname")%>">			
									&nbsp;&nbsp;&nbsp;<a class="allpage" href="updatecustinfo.asp">(Change Billing Information)</a>
									</td>
								</tr>
								<tr><td class="plaintextbold" align="right" > &nbsp;Last Name&nbsp; </td>
									<td class="plaintextbold"><%=session("lastname")%>&nbsp;<input type="hidden" name="txtlastname" id="txtlastname" value="<%=session("lastname")%>"></td>
								</tr>
								<tr><td class="plaintextbold" align="right" > &nbsp;Company&nbsp; </td>
									<td class="plaintextbold"><%=session("company")%>&nbsp;<input type="hidden" name="txtcompany" id="txtcompany"  value="<%=session("company")%>"></td>
								</tr>
								<tr><td class="plaintextbold" align="right" > &nbsp;Address&nbsp; </td>
									<td class="plaintextbold"><%=session("address1")%>&nbsp;<input type="hidden" name="txtaddress1" id="txtaddress1"  value="<%=session("address1")%>"></td>
								</tr>
								<tr><td class="plaintextbold" align="right" > &nbsp;Address 2&nbsp; </td>
									<td class="plaintextbold"><%=session("address2")%>&nbsp;<input type="hidden" name="txtaddress2" id="txtaddress2"  value="<%=session("address2")%>"></td>
								</tr>
								
								<tr id="addr3" <%if show_baddr3=true then %>style="display:''"<%else%>style="display:none"<%end if%> >
								<td class="plaintextbold" align="right"> Address 3 </td>
								<td class="plaintextbold"><%=session("address3")%>&nbsp;<input type="hidden" name="txtaddress3" id="txtaddress3" value="<% =session("address3") %>"></td>
							    </tr>
							    
								<tr><td class="plaintextbold" align="right" > &nbsp;City&nbsp; </td>
									<td class="plaintextbold"><%=session("city")%>&nbsp;<input type="hidden" name="txtcity" id="txtcity" value="<%=session("city")%>"></td>
								</tr>
								<%'if show_baddr3=false then %>
								<tr id="txtstatelabel" <%if show_baddr3=false then %>style="display:''"<%else%>style="display:none"<%end if%> >
								<td class="plaintextbold" align="right" > &nbsp;<%=txtstatelabelTxt %>&nbsp; </td>
									<td class="plaintextbold"><%=session("State")%>&nbsp;<input type="hidden" name="txtstate" id="txtstate" value="<%=session("State")%>"></td>
								</tr>
								<%'end if %>
								
								<tr id="bukcounty" <%if show_baddr3=true then %>style="display:''"<%else%>style="display:none"<%end if%>><td class="plaintextbold" align="right">County</td>
								<td class="plaintextbold"><input type="hidden" name="txtcounty" id="txtcounty" value="<%=session("bcounty")%>">
								<%=sitelink.Get_CountyName(session("bcounty"))%>
							</td>
							</tr>
								<tr><td class="plaintextbold" align="right" > &nbsp;<span id="bzipcodelabel"><%=txtbZiplabelTxt%></span>&nbsp; </td>
									<td class="plaintextbold"><%=session("zipcode")%>&nbsp;<input type="hidden" name="txtzipcode" id="txtzipcode"  value="<%=session("zipcode")%>"></td>
								</tr>
								<tr><td class="plaintextbold" align="right" > &nbsp;Country&nbsp; </td>
									<td class="plaintextbold"><%=sitelink.countryname(session("country"))%>  &nbsp;<input type="hidden" name="billtocountry" id="billtocountry" value="<%=session("country")%>"></td>
								</tr>
								<tr><td class="plaintextbold" align="right" > &nbsp;Phone&nbsp; </td>
									<td class="plaintextbold"><%=session("phone")%>&nbsp;<input type="hidden" name="txtphone" id="txtphone" value="<%=session("phone")%>">
									<span id="bphone" class="smalltextred"></span>
									</td>
								</tr>
								<tr><td class="plaintextbold" align="right" > &nbsp;Email&nbsp; </td>
									<td class="plaintextbold"><%=session("email")%>&nbsp;<input type="hidden" name="txtemail" id="txtemail"  value="<%=session("email")%>">
									<input type="hidden" name="txtemail1" id="txtemail1"  value="<%=session("email")%>">
									<input type="hidden" name="receiveemail"  value="0">
									<span id="bfname"></span>
									<span id="blname"></span>
									<span id="baddr1"></span>
									<span id="bcity"></span>
									<span id="bstate"></span>
									<span id="bzip"></span>									
									<span id="bemail"></span>
									<span id="bemailrepeat"></span>
									<span id="bpwd"></span>
									<input type="hidden" name="registered" id="registered" value="1">												
									</td>
								</tr>
								<SCRIPT language="JavaScript">
								<!--
								var browserName=navigator.appName;
								if (browserName=="Netscape")
								{
								 document.write('<tr><td class="plaintext" colspan="2"><img src="images/clear.gif" width="1" height="35" border="0"></td></tr>');
								}
								else
								{
								 document.write('<tr><td class="plaintext" colspan="2"><img src="images/clear.gif" width="1" height="55" border="0"></td></tr>');
								}
								//-->
								</script>
								
								</table>
						<%else%>
							<table width="100%" border="0" cellspacing="0" cellpadding="1">
							<tr><td colspan="2"><img src="images/clear.gif" height="30" width="1" border="0"></td></tr>
							<tr><td colspan="2" height="5"></td></tr>
							<tr><td class="plaintext" align="right"> First Name </td>
								<td class="smalltextred"><input type="text" name="txtfirstname" id="txtfirstname" size="20,1" maxlength="15" value="<% =session("firstname") %>">&nbsp;*
								<span id="bfname"></span>
								<input type="hidden" name="registered" id="registered" value="0">
								</td>
							</tr>
							<tr>
								<td class="plaintext" align="right"> Last Name </td>
								<td class="smalltextred"><input type="text" name="txtlastname" id="txtlastname" size="30,1" maxlength="20" value="<% =session("lastname") %>">&nbsp;*
								<span id="blname"></span>
								</td>
							</tr>
							<tr><td class="plaintext" align="right"> Company </td>
								<td><input type="text" name="txtcompany" id="txtcompany" size="40,1" maxlength="40" value="<% =session("company") %>"></td>
							</tr>
							<tr><td class="plaintext" align="right"> Address </td>
								<td class="smalltextred"><input type="text" name="txtaddress1" id="txtaddress1" size="40,1" maxlength="40" value="<% =session("address1") %>">&nbsp;*
								<span id="baddr1"></span>
								</td>
							</tr>
							<tr><td class="plaintext" align="right"> Apt. or Suite </td>
								<td><input type="text" name="txtaddress2" id="txtaddress2" size="40,1" maxlength="40" value="<% =session("address2") %>"></td>
							</tr>
							
							
							
							<tr id="addr3" <%if show_baddr3=true then %>style="display:''"<%else%>style="display:none"<%end if%> >
								<td class="plaintext" align="right"> Address 3 </td>
								<td ><input type="text" name="txtaddress3" id="txtaddress3" size="40,1" maxlength="40" value="<% =session("address3") %>"></td>
							</tr>
							
							<tr><td class="plaintext" align="right"> City </td>
								<td class="smalltextred"><input type="text" name="txtcity" id="txtcity" size="25,1" maxlength="20" value="<% =session("city") %>">&nbsp;*
								<span id="bcity"></span>
								</td>
								</tr>
								<tr id="txtstatelabel" <%if show_baddr3=true then %>style="display:none"<%else%>style="display:''"<%end if%>>
								<td class="plaintext" align="right"><span id="bstatelabel"><%=txtstatelabelTxt%></span> </td>
											 <td class="smalltextred">
											  <select name="txtstate" id="txtstate" class="smalltextblk">
												  <option value="">
												  <option value="INT" <%if session("state") = "INT" then response.write(" selected") end if%>>INTERNATIONAL																		  
												  <%
												  xmlstring =sitelink.Get_state_list(session("ordernumber"))
												  objDoc.loadxml(xmlstring)
												  tnode = "//state_list[ccode='" + cstr(billcountry) +"']"
												  set SL_Ship_state = objDoc.selectNodes(tnode)
													for x=0 to SL_Ship_state.length-1
														SL_state_name= SL_Ship_state.item(x).selectSingleNode("statename").text
														SL_state_code= SL_Ship_state.item(x).selectSingleNode("sstate").text														
												  %>
												  <option value="<%=SL_state_code%>"
												  <%if session("state") = SL_state_code then response.write(" selected") end if%>
												  ><%=SL_state_name%>
												  <%next						  
												  set SL_Ship_state=nothing
												  %>
												  </select>&nbsp;*
												  <span id="bstate"></span>
								 </td>
							</tr>
							
							<tr id="bukcounty" <%if show_baddr3=true then %>style="display:''"<%else%>style="display:none"<%end if%>><td class="plaintext" align="right">County</td>
								<td class="plaintext">
								
									 <select name="txtcounty" id="txtcounty" class="smalltextblk">												 
												  <%
												  xmlstring =sitelink.Get_County_list(session("ordernumber"))
												  objDoc.loadxml(xmlstring)
												  set SL_County_List = objDoc.selectNodes("//county_list")
													for x=0 to SL_County_List.length-1
														SL_county_name= SL_County_List.item(x).selectSingleNode("countyname").text
														SL_county_Code= SL_County_List.item(x).selectSingleNode("ccode").text													
												  %>
												  <option value="<%=SL_county_Code%>"
												  <%if session("bcounty") = SL_county_Code then response.write(" selected") end if%>
												   ><%=SL_county_name%>												   
												  <%next						  
												    set SL_County_List =nothing
												  %>
												  </select>
							</td>
							</tr>
							
							<tr><td class="plaintext" align="right"><span id="bzipcodelabel"><%=txtbZiplabelTxt%></span></td>
								<td class="smalltextred">
									<input type="text" name="txtzipcode" id="txtzipcode" size="10,1" maxlength="10" value="<% =session("zipcode") %>">&nbsp;*
									<span id="bzip"></span>
							</td>
							</tr>
							<tr><td class="plaintext" align="right">Country </td>
								<td><input type="hidden" id="whatbcountrytype" name="whatbcountrytype" value="1">				
								<select NAME="billtocountry" id="billtocountry" class="smalltextblk" onChange="Fixbcountry(this.options[this.selectedIndex].value,'B')">
									<%'=sitelink.listcountries(session("ordernumber"),session("shopperid"),"B",0)%>
									<%=sitelink.GET_COUNTRYLIST(billcountry)%>
								</select>
								</td>
							</tr>
							<tr><td class="plaintext" align="right"> Phone</td>
								<td class="smalltextred"><input type="text" name="txtphone" id="txtphone" size="20,1" maxlength="20" value="<% =session("phone") %>">
								<%if BILL_TOPHONE_REQUIRED=1 then response.write("&nbsp;*") end if%>
								<span id="bphone"></span>
								</td>
							</tr>
							<tr><td class="plaintext" align="right">E-Mail</td>
								<td class="smalltextred"><input type="text" name="txtemail" id="txtemail" size="40,1" maxlength="50" value="<% =session("email") %>">&nbsp;*
								<span id="bemail"></span>							
								</td>
							</tr>					
							<tr><td class="plaintext" align="right">Re-enter E-Mail</td>
								<td class="smalltextred"><input type="text" name="txtemail1" id="txtemail1" size="40,1" maxlength="50" value="<% =session("email1") %>">&nbsp;*
								<span id="bemailrepeat"></span>							
								</td>
							</tr>
						<tr><td class="plaintext" colspan="2"><img src="images/clear.gif" width="1" height="20" border="0"></td></tr>
						</table>
						<%end if%>
							
							</td>
							
						</tr>
						</table>						
						</td>					 
					 </tr>
					 
			 </table>
			 <br>
			  <%if session("registeredshopper") = "NO" and REGISTER_BEFORE_CHECKOUT=0 then %>
			 <table width="100%" border="0" cellspacing="0" cellpadding="1">			 
					<tr><td colspan="2" class="plaintext">
					If you would like to register, please enter a password to check your order status. Leave it blank if you decide not to register.
					
					</td></tr>
					
					<tr><td class="plaintext" align="right">Password</td>
						<td><input type="password" name="txtregpassword" id="txtregpassword" size="15" maxlength="10" value="<% =session("regpassword") %>">
							&nbsp;
							</td>														
					</tr>
						<tr><td class="plaintext" align="right" valign="top">Re-Enter Password</td>
							<td><input type="password" name="txtregpassword2" id="txtregpassword2" size="15" maxlength="10" value="<%=session("regpassword2")%>">
							&nbsp;
							<span class="smalltextred"><span id="bpwd"></span></span>
							</td>
					</tr>
					
				<tr><td colspan="2">&nbsp;</td></tr>
				<tr><td class="plaintext">&nbsp;  </td>
					<td class="plaintext">
					<input type="CHECKBOX" name="receiveemail" id="receiveemail" value="<%=session("receiveemail")%>" 
					<%if session("receiveemail")= 1 then
							response.write(" checked") 
						end if					
					 %>
					>
					Receive Email
					
					</td>
				</tr>
			 
			 </table>
			 <%end if%>
		
				</td>
				<td width="1"><img src="images/clear.gif" width="1" height="1" border="0"></td>
				<td width="50%" valign="top">
				<% if session("user_want_multiship")= true then	 %>
				<input type="hidden" name="billtocopy" id="billtocopy" value="1">
				<input type="hidden" name="chkboxbilltocopy" id="chkboxbilltocopy" value="0">
				<%else%>
				
				<table width="100%" cellpadding="0" cellspacing="0" border="0">
					 <tr>
						 <td class="THHeader">&nbsp;&nbsp;Shipping Address &nbsp;<span class="smalltextblk">(* Required)&nbsp;&nbsp; </span><a name="billtab"></a></td>						 
					 </tr>
					 <tr>
					 <tr>
					 	<td valign="top">
						<table width="100%" cellpadding="0" cellspacing="0" border="0">	
			 			<tr>
							
							<td valign="top" width="5"><img src="images/clear.gif" alt="" border="0" width="5" height="1"></td>
						    <td width="100%" valign="top">
							<table width="100%" border="0" cellspacing="0" cellpadding="1">													
							<tr><td colspan="2" height="5"></td></tr>
							<tr><td colspan="2" align="center" class="plaintext">
									<input type="hidden" name="chkboxbilltocopy" value="1">
									<input type="CHECKBOX" name="billtocopy" id="billtocopy" value="0" <%if session("BILLTOCOPY")=1 then response.write(" checked") end if%> onClick="ConvertAllFields(true,'<%=session("SL_START_COO")%>')">&nbsp;Same as Billing Address					
								&nbsp;&nbsp;&nbsp;
								<a class="allpage" href="SelectAddress.asp">Select From Address Book</a>
								<br><br>
								
								</td>
							</tr>
							<tr><td class="plaintext" align="right"> First Name </td>
								<td class="smalltextred"><input type="text" name="txtsfirstname" id="txtsfirstname" size="20,1" maxlength="15" value="<% =session("sfirstname") %>">
								&nbsp;*
								<span id="sfname"></span>
								</td>
							</tr>
							<tr>
								<td class="plaintext" align="right"> Last Name </td>
								<td class="smalltextred"><input type="text" name="txtslastname" id="txtslastname" size="30,1" maxlength="20" value="<% =session("slastname") %>">&nbsp;*
								<span id="slname"></span>
								</td>
							</tr>
							<tr><td class="plaintext" align="right"> Company</td>
								<td><input type="text" name="txtscompany" id="txtscompany" size="40,1" maxlength="40" value="<% =session("scompany") %>"></td>
							</tr>
							<tr><td class="plaintext" align="right"> Address </td>
								<td class="smalltextred"><input type="text" name="txtsaddress1" id="txtsaddress1" size="40,1" maxlength="40" value="<% =session("saddress1") %>">&nbsp;*
								<span id="saddr1"></span>
								</td>
							</tr>
							<tr><td class="plaintext" align="right"> Apt. or Suite </td>
								<td><input type="text" name="txtsaddress2" id="txtsaddress2" size="40,1" maxlength="40" value="<% =session("saddress2") %>"></td>
							</tr>
							<tr id="saddr3"  <%if show_saddr3=true then %>style="display:''"<%else%>style="display:none"<%end if%> ><td class="plaintext" align="right"> Address 3 </td>
								<td ><input type="text" name="txtsaddress3" id="txtsaddress3" size="40,1" maxlength="40" value="<% =session("saddress3") %>"></td>
							</tr>
							<tr><td class="plaintext" align="right">City</td>
								<td class="smalltextred">
								<input type="text" name="txtscity" id="txtscity" size="25,1" maxlength="20" value="<% =session("scity") %>">&nbsp;*
								<span id="scity"></span>
								</td>
							</tr>
							<tr id="txtsstatelabel" <%if show_saddr3=true then %>style="display:none"<%else%>style="display:''"<%end if%>>
							<td class="plaintext" align="right"><span id="sstatelabel"><%=txtsstatelabelTxt%></span></td>
							<td class="smalltextred">
								<select name="txtsstate" id="txtsstate"  class="smalltextblk">
									  
									  <option value="">
									  <option value="INT" <%if session("sstate") = "INT" then response.write(" selected") end if%>>INTERNATIONAL
									  <%
									  xmlstring =sitelink.Get_state_list(session("ordernumber"))
									  objDoc.loadxml(xmlstring)
									  tnode = "//state_list[ccode='" + cstr(shipcountry) +"']"
									  set SL_Ship_state = objDoc.selectNodes(tnode)
										for x=0 to SL_Ship_state.length-1
											SL_state_name= SL_Ship_state.item(x).selectSingleNode("statename").text
											SL_state_code= SL_Ship_state.item(x).selectSingleNode("sstate").text
									  %>
									  <option value="<%=SL_state_code%>"
									  <%if session("sstate") = SL_state_code then response.write(" selected") end if%>
									  ><%=SL_state_name%>
									  <%next						  
									  set SL_Ship_state=nothing
									  %>
									  </select>&nbsp;*
									  <span id="sstate"></span>
							 </td>
							</tr>
							<tr id="sukcounty" <%if show_saddr3=true then %>style="display:''"<%else%>style="display:none"<%end if%>><td class="plaintext" align="right">County</td>
								<td class="plaintext">
									
									
									 <select name="txtscounty" id="txtscounty" class="smalltextblk">												 
												  <%
												     xmlstring =sitelink.Get_County_list(session("ordernumber"))
												    objDoc.loadxml(xmlstring)
												    set SL_County_List = objDoc.selectNodes("//county_list")
													for x=0 to SL_County_List.length-1
														SL_county_name= SL_County_List.item(x).selectSingleNode("countyname").text
														SL_county_Code= SL_County_List.item(x).selectSingleNode("ccode").text													
												  %>
												  
												  <option value="<%=SL_county_Code%>"
												  <%if session("scounty") = SL_county_Code then response.write(" selected") end if%>
												  ><%=SL_county_name%>
												  <%next						  
												    set SL_County_List = nothing
												  %>
												  </select>
							</td>
							</tr>
							<tr><td class="plaintext" align="right"><span id="szipcodelabel"><%=txtsZiplabelTxt%></span>
							  <td class="smalltextred">
							  <input type="text" name="txtszipcode" id="txtszipcode" size="10,1" maxlength="10" value="<% =session("szipcode") %>">&nbsp;*
							  <span id="szip"></span>
							</td>
						</tr>
						<tr><td class="plaintext" align="right">Country </td>
							<td>
								<select NAME="shiptocountry" id="shiptocountry" class="smalltextblk" onChange="Fixbcountry(this.options[this.selectedIndex].value,'S')">
									<%'=sitelink.listcountries(session("ordernumber"),session("shopperid"),"S",session("address_id"))%>
									<%=sitelink.GET_COUNTRYLIST(shipcountry)%>
								</select>
							</td>
						</tr>
						<tr><td class="plaintext" align="right"> Phone</td>
							<td><input type="text" name="txtsphone" id="txtsphone" size="20,1" maxlength="20" value="<% =session("sphone") %>"></td>
						</tr>
						<tr><td class="plaintext" align="right"> E-Mail </td>
							<td class="smalltextred"><input type="text" name="txtsemail" size="40,1" id="txtsemail" maxlength="50" value="<% =session("semail") %>">
							<%if SHIP_TOEMAIL_REQUIRED=1 then response.write("&nbsp;*") end if%>
			                <span id="semail"></span>
							</td>
						</tr>
		                <tr><td class="plaintext" align="right"> &nbsp;Re-Enter E-Mail</td>
			                <td colspan="3" class="smalltextred"><input type="text" name="txtsemail1" id="txtsemail1"  size="40,1" maxlength="50" value="<%=session("semail")%>">
			                <span id="semailrepeat"></span>
			                </td>
		                </tr>
					</table>
					<br>
					<div id="GiftMsg" style="display:<%if session("billtocopy")=0 then Response.Write "visible" else Response.Write "none" end if%>">
					<table width="98%" cellpadding="0" cellspacing="0" border="0">
						 <tr>
							 <td class="THHeader">&nbsp;&nbsp;Gift Message for Recipient&nbsp;<span class="smalltextblk">(Optional)</span></td>
							 
						 </tr>
						 <tr>
							 <td  valign="top">
								<table width="100%" border="0" cellspacing="0" cellpadding="1">
								<tr><td align="center"><input type="text" name="txtgiftmsg1" value="<%=session("giftmsg1")%>" maxlength="35" size="50"></td></tr>
								<tr><td align="center"><input type="text" name="txtgiftmsg2" value="<%=session("giftmsg2")%>" maxlength="35" size="50"></td></tr>
								<tr><td align="center"><input type="text" name="txtgiftmsg3" value="<%=session("giftmsg3")%>" maxlength="35" size="50"></td></tr>
								<tr><td align="center"><input type="text" name="txtgiftmsg4" value="<%=session("giftmsg4")%>" maxlength="35" size="50"></td></tr>
								<tr><td align="center"><input type="text" name="txtgiftmsg5" value="<%=session("giftmsg5")%>" maxlength="35" size="50"></td></tr>
								<tr><td align="center"><input type="text" name="txtgiftmsg6" value="<%=session("giftmsg6")%>" maxlength="35" size="50"></td></tr>
								
								</table>				
							 </td>
							 
							</tr>
							
						 </table>
					</div>
							
						</td>
							
						</tr>
						</table>
						</td>
					 
					 </tr>
					 
				</table>
				
				
				<%end if%>
				</td>
			</tr>
			
		
		</table>
		
		
		
		
		<!-- old -->
		
		
		<div class="button-group">
        	<div class="button">
            	&nbsp;
				<input type="hidden" name="sitestore" value="SITELINK">
            </div>
            <div class="button">
            	<input type="image" name="Submitform" value="Submit" src="images/btn-continue.gif" border="0" alt="continue">
            </div>
            &nbsp;
            <span id="buySAFE_Kicker" name="buySAFE_Kicker" type="Kicker Guaranteed Ribbon 200x90"></span>
        </div>
	</form>
	
	
	</div>
<!-- end sl_code here -->
	</td>
	
</tr>

</table>
<% 
set SL_Ship_state=nothing 
set SL_County_List=nothing
%>
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
