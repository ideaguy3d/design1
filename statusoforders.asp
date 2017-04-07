<%on error resume next%>

<%
  if session("registeredshopper")="NO" then
  	response.redirect("statuslogin.asp")
  end if 
  
 ' response.write(session("shopperid"))

%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Customer Order History | Australian Products Co.</title>
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
<STYLE>
<!--
	INPUT  { font-size: 10px;  color: #000000 ; }

-->
</STYLE>

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
		<tr><td class="TopNavRow2Text" width="100%" >Order History</td></tr>
	</table>
	<br>
	<table width="100%"  border="0"  cellpadding="2" cellspacing="1">	
		<TR><td   class="THHeader"   width="33%" align="center"><b>Customer Name</b></Td>
			 <td  class="THHeader"  width="33%" align="center"><b>Zip Code</b></Td>					 
		</TR>
		 <TR><Td  class="tdRow1Color"   width="33%" align="center"><span class="plaintext"><br><%=session("firstname")%>&nbsp;<%=session("lastname")%><br><br></span></Td>
			 <Td  class="tdRow1Color"   width="33%" align="center"><span class="plaintext"><%=session("zipcode")%></span></Td>					
		 </TR>
	</TABLE>
	
	
	<%   		
		xmlstring = sitelink.dataforstatusnew("ORDERLIST",session("SHOPPERID"),0,"EMAIL") 
		 objDoc.loadxml(xmlstring)
		 
		 		 
		 set SL_Orderstatus = objDoc.selectNodes("//dfs")

	
	%>
	<p align="center">
				<font class="plaintextbold"><b>Click on the specific order number for details
				</b></font>
				<br>
				<br>
	<p>
	<table width="100%"  border="0" bordercolor="#617F9B" cellpadding="2" cellspacing="1">
	<tr>
         <td    class="THHeader" align="center"><b>Web Confirmation Number</b></td>
         <td    class="THHeader" align="center"><b>Order Number</b></td>
         <td   class="THHeader" align="center"><b>Ordered date</b></td>
         <td    class="THHeader" align="center"><b>Status</b></td>
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
		
		SL_orderNum = replace(SL_orderNum,".00","")
		MOM_OrderNum = replace(MOM_OrderNum,".00","")

		'MyArray = Split(SL_orderNum, ".", -1, 1)
		'	numeric_part = MyArray(0)
		'	decimal_part = MyArray(1)
		'	SL_orderNum = numeric_part
			
		if (x mod 2) = 0 then
		    class_to_use = "tdRow1Color"
		else
			 class_to_use = "tdRow2Color"
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
		<input type="submit"  value="<%=SL_orderNum%> " >
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
