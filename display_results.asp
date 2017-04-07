<%on error resume next%>


<!--#INCLUDE FILE = "include/momapp.asp" -->

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=althomepage%></title>
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



<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top"  >
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<center>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0"  >
		<tr><td class="TopNavRow2Text" width="100%" >Quick Order Form</td></tr>
	</table>
	<br>
	
	<form action="multiaddprod.asp" method="post" id=form1 name=form1>
	<table width="100%"  border="0"   cellspacing="1" cellpadding="5">
	<%
	rowcount = 0 
	for z=1 to session("num_of_quick_items") %>
	<%if len(trim(session("Stock"&z))) > 0 then 
		  rowcount = rowcount + 1
	
		 SET locstock = sitelink.setupproduct(cstr(session("Stock"&z)),cstr(""),0,session("ordernumber"),"") 
		 
		if (rowcount mod 2) = 0 then
			class_to_use = "tdRow1Color"
		else
			class_to_use = "tdRow2Color"
		end if 

	%>
	<tr>
	<% 
		if len(trim(locstock.inetthumb)) > 0 then
			StrFileName = "images/"+trim(locstock.inetthumb)
		else
			StrFileName = "images/"+trim(locstock.inetimage)
		end if
		StrPhysicalPath = Server.MapPath(StrFileName)
		set objFileName = CreateObject("Scripting.FileSystemObject")	
			if objFileName.FileExists(StrPhysicalPath) then
    	  		Thumbnail_imagename=StrFileName
		    else
			    Thumbnail_imagename="images/noimage.gif"
			end if
		set objFileName = nothing
		%>		
		<td width="140" class="<%=class_to_use%>"  valign="middle" align="center"><div class="prodthumb"><div class="prodthumbcell"><img SRC="<%=Thumbnail_imagename%>"  alt="<%=trim(locstock.inetsdesc)%>" class="prodlistimg" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %> border="0"></div></div></td>
		<td valign="top" class="<%=class_to_use%>" >
		<table width="100%" cellpadding="0" cellspacing="0" border="0">
		<tr><td class="plaintextbold"><% =trim(locstock.inetsdesc) %></td></tr>
		<tr><td class="plaintext">
		<% if locstock.size_color = false then %>
			<b>Item Number : <%=locstock.number%></b>
			<br>
			<%if locstock.giftcard = false then %>
			Price : <%=formatcurrency(locstock.price1)%>
			<%end if %>
		<%else%>
			<b>Item Number : <%=left(trim(session("Stock"&z)),10)%></b>
		<%end if%>		
		</td></tr>
		<% if locstock.attribs = true or locstock.giftcard = true then 
			if WANT_REWRITE = 1 then
				 SL_urltitle = trim(locstock.inetsdesc)
				 SL_urltitle = url_cleanse(SL_urltitle)
				 prod_num =left(session("Stock"&z),10)
				 prodlink = insecureurl  + SL_urltitle + "/productinfo/" + prod_num + "/" 
			else 
			    prod_num =left(session("Stock"&z),10)
				 prodlink = insecureurl +  "prodinfo.asp?number=" + prod_num
			end if
		%>
			<tr><td class="plaintext">
			<br>
			<b>This item cannot be added through Quick order.</b>
			<br>
			Click <a class="allpage" href="<%=prodlink%>">here</a> to add this item
			<input type="hidden" name="txtcustominfo" value="">
			</td></tr>
		<%else%>
		
		<% if locstock.size_color = true  then %>
			<tr><td>
			<select NAME="txtvariant<%=z%>" class="smalltextblk">
				<%
					xmlstring =sitelink.BUILDSIZECOLORLIST(cstr(trim(session("Stock"&z))),cstr(""),"STOCK.DESC2")
					objDoc.loadxml(xmlstring)						
					set SL_sbcl = objDoc.selectNodes("//sbcl") 
					for x=0 to SL_sbcl.length-1
						SL_Variation= SL_sbcl.item(x).selectSingleNode("scolor").text
						SL_VarationDesc= SL_sbcl.item(x).selectSingleNode("desc2").text
						SL_VarationPrice= SL_sbcl.item(x).selectSingleNode("price1").text
						'SL_VarationUnits= SL_sbcl.item(x).selectSingleNode("units").text
				%>
				<option value="<%=SL_Variation%>"><%=SL_VarationDesc%>&nbsp;<%=FORMATCURRENCY(SL_VarationPrice)%>
				<%next%>
		</select>
		</td></tr>
		<%
		set SL_sbcl = nothing
		end if%>
		<%if locstock.inetcustom=true then %>
			<tr><td>
			<table>
			<tr><td class="THHeader"><%=trim(locstock.inetcprmpt)%></td></tr>
			<tr><td><textarea name="txtcustominfo" cols="30" rows="2" value></textarea></td></tr>
			</table>
			</td></tr>
		<%else%>
			<input type="hidden" name="txtcustominfo" value="">
		<%end if%>
		<%end if%>
		
		</table>
		
		
		</td>
		<td valign="middle" align="center" width="80" class="<%=class_to_use%>" >
			<%if locstock.attribs=true or locstock.giftcard = true then%>
				<span class="plaintextbold">
				<input type="hidden" class="plaintext" name="txtquanto" value="0">
				<input type="hidden" name="stockitem" value="<%=session("Stock"&z)%>" >
				</span>
			<%else%>
				<span class="plaintextbold">
				Qty <input type="text" class="plaintext" name="txtquanto" value="1" size="3" valign="bottom" align="absmiddle">
				<input type="hidden" name="stockitem" value="<%=session("Stock"&z)%>" >
				</span>
			<%end if%>
		</td>
	</tr>
	  
	
	<% set locstock = nothing %>
	<%end if%>
	<%next%>
	<tr><td colspan="10" align="center">
		  <input type="hidden" name="rowcount" value="<%=rowcount%>">
		  <input type="submit" value="Add to cart"  id=submit1 name=submit1></td></tr>
	</table>
	</form>
	
	

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
