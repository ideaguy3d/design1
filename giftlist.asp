<%on error resume next%>

<%
giftcustloginId= 0
if len(trim(session("giftregcustnum"))) > 0  then
	giftcustloginId = session("giftregcustnum")
end if
if isnumeric(giftcustloginId)=false then
	response.redirect("giftreglogin.asp")
end if
%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%
if (Request("id") > 0) then
	session("giftregcustnum")=cint(Request("id"))
	giftregistrycustnum=cint(Request("id"))
elseif (session("giftregcustnum") > 0) then
	giftregistrycustnum=cint(session("giftregcustnum"))
end if
'giftregistrycustnum=session("savegiftregid")
xmlstring = ""
xmlstring = sitelink.GIFTREGSEARCH("","CUSTNUM",false,"","",giftregistrycustnum)

objDoc.loadxml(xmlstring)

set SL_GiftReg = objDoc.selectNodes("//giftreg")	  				 
		
if SL_GiftReg.length = 1 then
	SL_Regfname = SL_GiftReg.item(0).selectSingleNode("firstname").text
	SL_Reglname = SL_GiftReg.item(0).selectSingleNode("lastname").text
	SL_EventDate = SL_GiftReg.item(0).selectSingleNode("eventdate").text
	SL_RegDesc = SL_GiftReg.item(0).selectSingleNode("eventdesc").text
	SL_CoRegfname = SL_GiftReg.item(0).selectSingleNode("reg2fname").text
	SL_CoReglname = SL_GiftReg.item(0).selectSingleNode("reg2lname").text
	SL_RegType = SL_GiftReg.item(0).selectSingleNode("reg_type").text
end if

'response.write(giftregistrycustnum)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Customer Gift List Registry @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
	<table width="630" cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="630">
			<a href="giftregistry.asp"><font class="TopNavRow2Text">Gift Registry</font></a>
			&nbsp;>&nbsp;<font class="TopNavRow2Text">Gift List</font></td></tr>
	</table>

	
	<%
	if (giftregistrycustnum > 0) then
	  
		extra_field = "LEFT(JUSTFNAME(stock.inetimage)+SPACE(50),50) AS inetimage "
		xmlstring =sitelink.GIFTLIST("LIST",giftregistrycustnum,"0",0,0,"",extra_field)

		objDoc.loadxml(xmlstring)
		set SL_GIFTLIST = objDoc.selectNodes("//wls")	  				 

	%>  
	<br>
	<table width="745" border="0"  cellpadding="2" cellspacing="1">
	<tr><td class="plaintext"><%=SL_Regfname%>&nbsp;<%=SL_Reglname%> <%if len(SL_CoRegfname) > 0 or len(SL_CoReglname) > 0 then Response.Write(" and " & SL_CoRegfname & "&nbsp;" & SL_CoReglname) end if%>
		<br>Event: <%=SL_RegType%>
		<br><%=FormatDateTime(CDate(SL_EventDate),1)%>
		<%if len(trim(SL_RegDesc)) > 0 then Response.Write "<br>" & SL_RegDesc%><br><br>
		<!-- <%if SL_Giftlist.length = 1 then %>
		There is <%=SL_Giftlist.length%> item in your gift registry
		<%elseif SL_Giftlist.length > 1 then %>
		There are <%=SL_Giftlist.length%> items in your gift registry
		<%end if%> -->
		Please make any changes to your registry below. Click on "add additional items" to add more items to your registry.
		</td>
	</tr>
	</table>
	<%if SL_Giftlist.length > 0 then %>
	<font class="plaintext">There are <%=SL_Giftlist.length%> items in Gift Registry</font>
	<%end if%>
	<form name="giftlist" id="giftlist" action="updategiftlist.asp" method="post">
	<table width="590" border="0"  cellpadding="2" cellspacing="1">
	<tr>
			<th align="left" width="80" class="THHeader" >Remove</th>
			<th align="CENTER" width="150" class="THHeader">Qty Requested</th>
			<th align="CENTER" width="360" class="THHeader" >Description
			
			</th>
			</tr>
	<% if SL_Giftlist.length = 0 then %>
	<tr><td class="plaintextbold" colspan="3" >
	 <br>
	 <ul>Gift Registry is Empty...
	 <br><br>
	 </ul>
	 </td></tr>
	 <tr><td colspan="2">&nbsp;</td>
	 <td valign="top"><a HREF="default.asp"><img src="images/btn-continueshopping.gif" border="0" alt="Continue Shopping"></a>
	 </td></tr>
	
	<%else%>
		<% for x=0 to SL_Giftlist.length-1 
		 SL_Number	= SL_Giftlist.item(x).selectSingleNode("item").text
		 SL_item_id	= SL_Giftlist.item(x).selectSingleNode("item_id").text
		 SL_qty		= SL_Giftlist.item(x).selectSingleNode("qty_want").text
		 SL_ord		= SL_Giftlist.item(x).selectSingleNode("qty_ord").text
		 SL_ItemDesc= SL_Giftlist.item(x).selectSingleNode("inetsdesc").text
		 SL_ItemDesc2= SL_Giftlist.item(x).selectSingleNode("desc2").text		 
		 SL_ItemPrice= SL_Giftlist.item(x).selectSingleNode("price1").text
		 SL_thumbnail= SL_Giftlist.item(x).selectSingleNode("inetthumb").text
		 SL_Fullimage  = SL_Giftlist.item(x).selectSingleNode("inetimage").text
		 
		 SL_StAttibprice  = cdbl(SL_Giftlist.item(x).selectSingleNode("price_wopt").text)
		 SL_StAttibNonSkuprice  = cdbl(SL_Giftlist.item(x).selectSingleNode("nskuprice").text)
		 SL_WishlistCinfo  = SL_Giftlist.item(x).selectSingleNode("custominfo").text
		 
		 SL_ItemPrice = cdbl(SL_ItemPrice) + SL_StAttibprice + SL_StAttibNonSkuprice
		 
		 if (len(trim(SL_qty)) = 0) or (IsNumeric(SL_qty) = false) then 
			SL_qty = 0
		 else
			SL_qty = SL_qty + 1 - 1
		 end if
		 if (len(trim(SL_ord)) = 0) or (IsNumeric(SL_ord) = false) then 
			SL_ord = 0
		 else
			SL_ord = SL_ord + 1 - 1
		 end if

		
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
		<td valign="middle" align="center" class="<%=class_to_use%>">
		<%'if (SL_qty > SL_ord) then%>
		<%if (SL_ord <= 0) then%>
					<%remname  ="remove" & cstr(x+1)%>
					<input type="checkbox" name="<%=remname%>" value="0">
		<%else
			Response.Write SL_ord & " item/s Fulfilled"%>
		<%end if%>
			<input type="hidden" name="item" value="<%=SL_Number%>">
			<input type="hidden" name="item_id" value="<%=SL_item_id%>">
		</td>
		<td valign="middle" align="center" class="<%=class_to_use%>">			  
			<span class="plaintext">	
			<%if (SL_qty-SL_ord) > 0 then%>
					<input type="text" name="qty" size="3" class="smalltextblk" maxlength="4" value="<%=SL_qty%>">
			<%else%>
				<%=SL_qty%>
				<input type="hidden" name="qty" value="<%=SL_qty%>">
				<input type="hidden" name="qtyord" value="<%=SL_qty%>">
			<%end if%>
			</span>
		</td>
		<td  valign="top" class="<%=class_to_use%>">			
				<table width="501" cellpadding="3" cellspacing="0" border="0">
				<tr>
					<td valign="top" width="60">
						<img src="<%=inethumbImage%>" alt="<%=altag%>" width="60" height="60" border="0">
					</td>
					<td valign="top" class="smalltextblk">
					<%=SL_ItemDesc%>			
					<% if len(trim(SL_ItemDesc2)) > 0 then response.write("<br>" +SL_ItemDesc2) end if %>
					<br>
					<b>Item :</b><%=SL_Number%> &nbsp; 
					<br>
					<b>Unit Price : </b><%=formatcurrency(SL_ItemPrice)%> &nbsp;
					<!--&nbsp;&nbsp;&nbsp;<a  class="allpage" href="addtocart.asp?item=<%=SL_Number%>&item_id=<%=SL_item_id%>"><font color="red"><b>Add to Cart</b></font></a>-->
					
					</td>
				</table>
		</td>
		</tr>
		<%next%>

		<tr><td colspan="2" align="center">
			<input type="image" value="Update Gift List" src="images/btn_updategiftlist.gif" border="0" alt="Update Gift List" id=image1 name=image1> 
			</td>
			<td align="center"></td>
		</tr>
		<tr><td colspan="3" align="center">
		<br>
		<table width="80%" border="0">
		<tr><td colspan="3" align="right">
			<!--<form action="getnoemails.asp" method="post">-->
				<a href="getnoemails.asp"><img src="images/btn-completegiftreg.gif" border="0" alt="Complete Gift Registry" value="Complete Gift Registry"></a>
			<!--</form>-->
		</td>
		<td colspan="2">&nbsp;</td>
		<td valign="top"><a HREF="default.asp"><img src="images/btn-continueshopping.gif" border="0" alt="Continue Shopping"></a>
		</td>
		</table>		
		</td></tr>
		
	
	<%end if%>
	</table>
	</form>

	 <% set SL_Giftlist = nothing
	 	'Set objDoc = nothing
	 %>
	 
	<%else%>
	
	<table width="99%" cellpadding="3" cellspacing="0" border="0">	
	<tr><td class="plaintextbold" align="center">
	<br><br>
	You must Log in to gain access to your Gift Registry
	<br><br>
	<a class="allpage" href="giftreglogin.asp?dest=Giftlist.asp">Click to Login</a>
	<br><br>
	</td></tr>
	</table>
	<%end if%>
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
