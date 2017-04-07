<%on error resume next%>

<%
	if session("callcenter_logged") = false then
		response.redirect("call_center_login.asp")
	end if

%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->



<%
	
		lcustnumber = request("bcustnum")
		isweborder = request("isweborder")
		
		if isweborder = 1 then
			'get shopper record			
			Set shopperInfo = sitelink.SHOPPERINFO(lcustnumber)
			
			    SL_fname = trim(shopperInfo.firstname)
				SL_lname = trim(shopperInfo.lastname)	
				SL_company = trim(shopperInfo.company)	
				SL_addr = trim(shopperInfo.addr)	
				SL_addr2 = trim(shopperInfo.addr2)
				SL_addr3 = trim(shopperInfo.addr3)
				SL_city = trim(shopperInfo.city)	
				SL_state = trim(shopperInfo.state)				
				SL_zipcode = trim(shopperInfo.zipcode)	
				SL_email = trim(shopperInfo.email)
				SL_Phone = trim(shopperInfo.phone)
				SL_county = trim(shopperInfo.county)
				SL_country = shopperInfo.country
			
				session("firstname")=SL_fname
				session("lastname")=SL_lname
				session("company")=SL_company
				session("address1")=SL_addr
				session("address2")=SL_addr2
				session("address3")=SL_addr3
				session("city")=SL_city
				session("state")=SL_state
				session("zipcode")=SL_zipcode
				session("email")=SL_email	
				session("phone")=SL_Phone
				session("bcounty") = SL_county
				session("country") = SL_country
			
				if len(trim(session("email"))) = 0 then
					session("email") = "XX"
				end if
				
				
			Set shopperInfo= nothing
		
		else
			xmlstring=sitelink.GET_CUST_RECORD(cstr("BYCUSTNUM"),lcustnumber,cstr(""),cstr(""),cstr(""),cstr(""))
			objDoc.loadxml(xmlstring)
			set SL_order_rec = objDoc.selectNodes("//cust_rec")
			'response.write(SL_order_rec.length)
		
			if SL_order_rec.length > 0 then
				SL_fname = SL_order_rec.item(0).selectSingleNode("firstname").text
				SL_lname = SL_order_rec.item(0).selectSingleNode("lastname").text	
				SL_company = SL_order_rec.item(0).selectSingleNode("company").text	
				SL_addr = SL_order_rec.item(0).selectSingleNode("addr").text	
				SL_addr2 = SL_order_rec.item(0).selectSingleNode("addr2").text
				SL_addr3 = SL_order_rec.item(0).selectSingleNode("addr3").text
				SL_city = SL_order_rec.item(0).selectSingleNode("city").text	
				SL_state = SL_order_rec.item(0).selectSingleNode("state").text				
				SL_zipcode = SL_order_rec.item(0).selectSingleNode("zipcode").text	
				SL_email = SL_order_rec.item(0).selectSingleNode("email").text
				SL_Phone = SL_order_rec.item(0).selectSingleNode("phone").text
				SL_county = SL_order_rec.item(0).selectSingleNode("county").text
				SL_country = SL_order_rec.item(0).selectSingleNode("country").text
				
				session("firstname")=SL_fname
				session("lastname")=SL_lname
				session("company")=SL_company
				session("address1")=SL_addr
				session("address2")=SL_addr2
				session("address3")=SL_addr3
				session("city")=SL_city
				session("state")=SL_state
				session("zipcode")=SL_zipcode
				session("phone")=SL_Phone
				session("email")=SL_email
				session("bcounty") = SL_county
				session("country") = SL_country
			
				if len(trim(session("email"))) = 0 then
					session("email") = "XX"
				end if
								
		end if 
		
		set SL_order_rec = nothing
		end if	
	
		'response.write(session("bcounty"))	


   if session("callcenter_logged")=true then
	    session("registeredshopper")="YES"
   end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Customer Order History Information @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
		<tr><td class="TopNavRow2Text" width="100%">Order History</td></tr>
	</table>
	<br>
	<%
	'xmlstring = sitelink.CUST_ORDER_HISTORY_HOCHUNK(cstr(""),cstr(""),cstr(""),lcustnumber)
		xmlstring = sitelink.CUST_ORDER_HISTORY_HOCHUNK(cstr(SL_fname),cstr(SL_lname),cstr(SL_zipcode),lcustnumber)
		
		objDoc.loadxml(xmlstring)
	
		%>
	<table width="100%" border="0" cellspacing="1" cellpadding="1">
	
			<tr><td bgcolor="#333333"> 		
			 <table width="100%" border="0" cellspacing="0" cellpadding="3" >
			 <TR><td  class="THHeader"   width="33%" align="center">Customer Information</Td>

		</TR>
		<% if len(trim(SL_fname)) = 0 and len(trim(SL_lname)) = 0 then %>
		 <TR><Td  class="pagenavbg"   colspan="2" align="center"><span class="plaintext"><br>No Customer Found<br><br></span></Td>

		 </TR>
		
		<%else%>
		 <TR><Td  class="pagenavbg"   width="33%" align="center"><span class="plaintext">
		 <br><%=SL_fname%>&nbsp;<%=SL_lname%>
		 &nbsp;&nbsp;
		 <%=sl_addr%>&nbsp;<%=sl_addr2%>&nbsp;<%=sl_city%>&nbsp;<%=sl_state%>&nbsp;<%=SL_zipcode%>
		 </span>
		 </Td>
			 
		 </TR>
		 <%end if%>
			 
			 </table>
			 </td></tr>
	</table>
	
	
	
	<%   
		
		set SL_Orderstatus = objDoc.selectNodes("//dfs")
		'response.write(SL_Orderstatus.length)
	
	%>
	<p align="center">
				<font class="plaintextbold"><b>Click on the specific order number for details
				
				<%'=sitelink.getdata("startlic")%>
				
				</b></font>
				<br>
				<br>
	<p>
	<table width="100%" border="0" bordercolor="#617F9B" bgcolor="#333333" cellpadding="1" cellspacing="1">
	<tr>
         <td  class="THHeader" bgcolor="black" align="center">Web Confirmation Number</td>
         <td  class="THHeader" bgcolor="black" align="center">Order Number</td>
         <td  class="THHeader" bgcolor="black" align="center">Ordered date</td>
         <td  class="THHeader" bgcolor="black" align="center">Status</td>
	<tr>
	<% if SL_Orderstatus.length = 0 then %>
		<tr><td colspan="4" align="center" class="plaintextbold">
		<br>
		<ul>
		<b>No Orders</b>
		</ul>
		 </td></tr>
	<%else %>
	<% for x=0 to SL_Orderstatus.length-1 
		SL_orderNum = SL_Orderstatus.item(x).selectSingleNode("internetid").text
		'SL_orderNum = SL_orderNum + 1 -1 
		MOM_OrderNum = SL_Orderstatus.item(x).selectSingleNode("order").text
		Order_date = SL_Orderstatus.item(x).selectSingleNode("order_date").text
		'Order_status = SL_Orderstatus.item(x).selectSingleNode("status").text
		
		
		MyArray = Split(SL_orderNum, ".", -1, 1)
			numeric_part = MyArray(0)
			decimal_part = MyArray(1)
			SL_orderNum = numeric_part
			
		if (x mod 2) = 0 then
		    class_to_use = "tdRow2Color"
		else
			 class_to_use = "tdRow1Color"
		end if 
		
		if MOM_OrderNum = 0 then
			Order_status= sitelink.dataforstatusnew("ORDERSTATUS",SL_orderNum,SL_orderNum,"EMAIL",false)
		else
			Order_status= sitelink.dataforstatusnew("ORDERSTATUS",MOM_OrderNum,MOM_OrderNum,"EMAIL",true)
		end if 
	

	%>	
	<form action="DetailOrder.asp" method="post" id=form1 name=form1>
	<%if MOM_OrderNum = 0 then %>
	<tr>	
		<td class="<%=class_to_use%>" align="center" >
		<input type="hidden" name="MOMOrdNum" value="0">
		<input type="hidden" name="OrdNum" value="<%=SL_orderNum%>">
		<input type="submit"  value="<%=SL_orderNum%>" >
		<input type="hidden" name="oktocontinue" value="1">
		</td>
		<td class="<%=class_to_use%>" align="center" ><span class="plaintext">N/A</span></td>
		<td class="<%=class_to_use%>" align="center" ><span class="plaintext"><%=formatdatetime(Order_date,2)%></span></td>
		<td class="<%=class_to_use%>" align="center" ><span class="plaintext"><%=Order_status%></span></td>
	</tr>
	<%else%>
		<tr>
		<td class="<%=class_to_use%>" align="center" ><span class="plaintext"><%=SL_orderNum%></span></td>
		<td class="<%=class_to_use%>" align="center" >
		<input type="hidden" name="MOMOrdNum" value="1">								
		<input type="hidden" name="OrdNum" value="<%=MOM_OrderNum%>">
		<input type="hidden" name="SL_OrderNum" value="<%=SL_orderNum%>">
		<input type="submit" value="<%=MOM_OrderNum%> " >
		<input type="hidden" name="oktocontinue" value="1">
		</td>		
		<td class="<%=class_to_use%>" align="center" ><span class="plaintext"><%=formatdatetime(Order_date,2)%></span></td>
		<td class="<%=class_to_use%>" align="center" ><span class="plaintext"><%=Order_status%></span></td>
	</tr>
	
	<%end if%>
	</form>
	
	<%next%>
	<%end if%>
	</table>
	
	
	<% 
	set SL_Orderstatus = nothing
	'Set objDoc = nothing %>
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
