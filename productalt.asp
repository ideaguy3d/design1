<%on error resume next%>
<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 


<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<% 

    current_stocknumber = cstr(REQUEST.QUERYSTRING("altitem"))
    current_stocknumber =fix_xss_Chars(current_stocknumber)
    
    current_variation = cstr(REQUEST.QUERYSTRING("altvariation"))
    current_variation =fix_xss_Chars(current_variation)
    
	'this is the itemid
	txtebasketitem = REQUEST.QUERYSTRING("altbasketnum")	
    
	isvalidstock = sitelink.validstocknumber2(current_stocknumber)
	
	if isvalidstock=-1 or isnumeric(txtebasketitem)=false then
		set sitelink=nothing
		set ObjDoc=nothing	
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl&"default.asp"
		Response.End	
	end if
  	
	SET LOSTOCK =  sitelink.setupproduct(current_stocknumber,cstr(current_variation),2,session("ordernumber"),cstr(txtebasketitem))
	
  
	bookmarktitle = replace(LOSTOCK.inetsdesc,"'","")
	
	session("department") = sitelink.GET_DEPT_FROM_PROD(current_stocknumber)
	
	froogle_desc = trim(LOSTOCK.froogle)
	froogle_desc = replace(froogle_desc,"""","")
	froogle_desc = replace(froogle_desc,"<","")
	froogle_desc = replace(froogle_desc,">","")
	
	keylist = trim(LOSTOCK.inetkeywrd)
	keylist = replace(keylist,"""","")
	keylist = replace(keylist,"<","")
	keylist = replace(keylist,">","")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | <%=trim(LOSTOCK.inetsdesc)%>-<%=althomepage%> | Australian Products Co. Bringing Australia to You - Aussie Foods </title>
<meta NAME="KEYWORDS" CONTENT="<%=keylist%>">
<meta NAME="description" CONTENT="<%=froogle_desc%>">
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

<script language="JavaScript">
function openWindow() {  popupWin = window.open('http://mail.mailordercentral.com/sitelink700/default.asp?item_url=<%=insecureurl%>prodinfo.asp?number=<%=cstr(current_stocknumber)%>&desc=<%=trim(bookmarktitle)%>', '', 'scrollbars,resizable,width=550,height=450')
}

</script>



</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">

<% session("destpage")="productalt.asp?altitem="+cstr(current_stocknumber) %>

<% if request.servervariables("server_port_secure") = 1 then %>
	<base href="<%=secureurl%>">
<% else %>
	<base href="<%=insecureurl%>">
<% end if %>


<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<center>
	<% if session("department") > 0 then %>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
				<tr><td class="breadcrumbs" width="100%">
						<a class="allpage" href="default.asp"><font class="breadcrumb-link">Home</font></a><span class="breadcrumb-divide">&nbsp;&nbsp;&#47;&nbsp;&nbsp;</span>
						<%
						deptstr = sitelink.deptstring(session("department"),true,"<span class='breadcrumb-divide'>&nbsp;&nbsp;&#47;&nbsp;&nbsp;</span>","breadcrumb-link","",WANT_REWRITE)
						if WANT_REWRITE=1 then
						    targetstr = "departments/"+cstr(session("department"))
						    str_replace = "products/"+cstr(session("department"))
						    deptstr = replace(deptstr,targetstr,str_replace)
						else
						    targetstr = "departments.asp?dept="+cstr(session("department"))
						    str_replace = "products.asp?dept="+cstr(session("department"))
						    deptstr = replace(deptstr,targetstr,str_replace)
						end if
						%>
						<%=deptstr%>
					</td>
				</tr>
	</table>
	<%end if%>
	<br>
	
	<table width="100%"  BORDER="0" cellpadding="3" cellspacing="0">
		<tr><td class="ProductTitle"><%=trim(LOSTOCK.inetsdesc)%></td>				
		</tr>
				
	</table>
	<br>
	<table width="100%"  cellpadding="0" cellspacing="0" border="0">

	        <tr> 
              <td width="315" valign="top"> 
			  <center>
                <% StrFileName = "images/"+trim(LOSTOCK.inetimage)
				StrPhysicalPath = Server.MapPath(StrFileName)
			  set objFileName = CreateObject("Scripting.FileSystemObject")	
			  if objFileName.FileExists(StrPhysicalPath) then
		       		imagename=StrFileName
			  else
					imagename="images/noimage.gif"
  			  end if
			  set objFileName = nothing
			%>
                <img SRC="<%=imagename%>" alt="<%=trim(LOSTOCK.inetsdesc)%>" title="<%=trim(LOSTOCK.inetsdesc)%>" hspace="5" border="0" align="middle"> 
               <%
				'view large image
				found = false
				'check for invalid character.
				
				Has_invalid_char= false 
				firstchar = left(current_stocknumber,1)
				
				select case firstchar
					case "\","/",":","?","*","""","<",">","|",";"
						Has_invalid_char=true					
				end select
											
				if Has_invalid_char = true then
					 found=false
				else				
					StrFileName = "images/"+cstr(current_stocknumber) +"-Large.jpg"
					StrPhysicalPath = Server.MapPath(StrFileName)
					set objFileName = CreateObject("Scripting.FileSystemObject")	
					if objFileName.FileExists(StrPhysicalPath) then
			     			largeimagename=StrFileName
							found = true
					else
						largeimagename="images/noimage.gif"
						found = false
	  				end if
					set objFileName = nothing		
				end if
				  
				 	
			%>
                <script>
			function ShowlargeImage(whatimage)
				{				 
				 popupWin = window.open(whatimage, '', 'scrollbars,resizable,width=500,height=450');
				}//-->
			</script>
                <% if found = true then %>
                <br>
               
                <center>
                  <a class="allpage" href="javascript:ShowlargeImage('<%=largeimagename%>')"><font color="red"><u><b>View 
                  Large Image</b></u></font></a> 
                </center>
                <%end if%>

					</center>


              </td>
			  
              <form method="POST" name="prodinfo" action="updatebasket.asp">			  
                <td width="315" valign="top">
					<input type="hidden" name="EDITITEM" value="<%=current_stocknumber%>">
					<table width="100%" border="1" bordercolor="#CCCCCC" style="border-collapse: collapse" cellpadding="8" cellspacing="0">
						<tr><td class="tdRow1Color" width="100%">
							<table width="100%" cellpadding="0" cellspacing="0" border="0">
							
							
								<tr><td class="plaintextbold">Item Number:&nbsp;<%=current_stocknumber%></td></tr>
								
							<tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0"></td></tr>
							<%if LOSTOCK.size_color = false then %>
							<tr><td class="ProductPrice">Unit Price: <%=FORMATCURRENCY(LOSTOCK.price1)%></td></tr>
							<%end if%>
							
							<%if SHOW_COMP_PRICE="TRUE" and LOSTOCK.inetcprice > 0 then %>
								<tr><td class="CompPrice">Compare At: <%=FORMATCURRENCY(LOSTOCK.inetcprice)%></td></tr>
							<%end if%>

							<%
							  extra =""
							  xmlstring = sitelink.getspecialprices(current_stocknumber,cstr(""),SESSION("ordernumber"),session("shopperid"),true,cstr(extra)) 
							  objDoc.loadxml(xmlstring)
					  				  					  
							  set SL_gsp = objDoc.selectNodes("//gsp") 
							  if SL_gsp.length > 0 then %>
							  <tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0"></td></tr>
							  <tr><td valign="top">

							  <table width="100%" border="0" cellspacing="1" cellpadding="1">
							   <tr><td class="THHeader" colspan="3">Special Pricing</td></tr>
							   <%
							   
									for x=0 to SL_gsp.length-1
									
									'SL_SP_Number = SL_gsp.item(x).selectSingleNode("number").text
									SL_SP_qty = SL_gsp.item(x).selectSingleNode("qty").text
									SL_SP_price = SL_gsp.item(x).selectSingleNode("price").text
									SL_SP_discount = SL_gsp.item(x).selectSingleNode("discount").text
									'SL_SP_desc2 = SL_gsp.item(x).selectSingleNode("desc2").text

									unitprice = LOSTOCK.price1
									unitprice2 = SL_SP_price
									if unitprice2>0 then
										unitprice = unitprice2
									end if
									percentoff = (SL_SP_discount/100.0)
									discountedUnitPrice = cdbl(unitprice * (1- percentoff))
							
									'MyArray = Split(SL_SP_qty, ".", -1, 1)
									'numeric_part = MyArray(0)
									'decimal_part = MyArray(1)
				
									'if decimal_part > 0 then
									'	SL_SP_qty = numeric_part +"." + decimal_part
									'else
									'	SL_SP_qty = numeric_part				
									'end if
									
									SL_SP_qty = replace(SL_SP_qty,".00","")
							
									if (x mod 2) = 0 then
										   class_to_use = "tdRow1Color"
									else
										class_to_use = "tdRow2Color"
									end if 

									
													
								%>
								<tr>
								<td class="<%=class_to_use%>"><span class="plaintext">&nbsp;Buy&nbsp;<%=SL_SP_qty%>&nbsp;for&nbsp;<%=formatcurrency(discountedUnitPrice)%>&nbsp;each</span></td>
								</tr>
								
							<%next
							set SL_gsp =  nothing
							
							%>
							  </table>
							  </td>
							  </tr>
							  
							  <%end if%>
							  
							  <%if SHOW_IN_STOCK =1 and LOSTOCK.size_color = false then 
								avail_status = "In Stock"
					  			if (LOSTOCK.units > 0) or (LOSTOCK.dropship = true) or (LOSTOCK.construct = true) then
									avail_status="In Stock"
							  	else
									if SHOW_DUE_DATE = 1 then
									    exp_date =sitelink.expected_date(current_stocknumber,current_variation)
											'if exp_date= FALSE or exp_date = TRUE then
											if IsDate(exp_date) = false then
												avail_status="Out of Stock"
											else
											     avail_status="Expected on : "+ exp_date
											end if
									else
										avail_status="Out of Stock"
									end if
								end if
								%>
								<tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0"></td></tr>
								<tr><td class="plaintextbold"><%=avail_status%></td></tr>
								<%
								end if
								%>

							  							
							<tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0"></td></tr>
							<%
							SL_prefShip_Title = ""
							SL_prefShip_Method = trim(LOSTOCK.prefship)
							if len(SL_prefShip_Method) > 0 then							
								xmlCarrier = sitelink.SHIPPINGMETHODS() 
							  	objDoc.loadxml(xmlCarrier)
								tnode = "//shipmeths[ca_code='"+SL_prefShip_Method+"']"
								set SL_Ship = objDoc.selectNodes(tnode) 
								if SL_Ship.length > 0 then
									SL_prefShip_Title = SL_Ship.item(0).selectSingleNode("ca_title").text									
								end if									
								set SL_Ship = nothing				
							end if							
							
							%>
							<% if len(SL_prefShip_Title) > 0 then %>
							<tr><td class="plaintext"><br>Will be shipped via: <%=SL_prefShip_Title%></td></tr>
							<%end if%>
							
							<% if LOSTOCK.size_color = true then %>
							<tr><td class="THHeader">Size/Color</td></tr>
							<tr><td>
							<select NAME="txtvariant" class="smalltextblk">
							<%
								extra_field = ""
					 			xmlstring =sitelink.BUILDSIZECOLORLIST(current_stocknumber,cstr(""),"STOCK.DESC2",extra_field)		
								objDoc.loadxml(xmlstring)	
								
								set SL_sbcl = objDoc.selectNodes("//sbcl") 

								 for x=0 to SL_sbcl.length-1
									SL_Variation= SL_sbcl.item(x).selectSingleNode("scolor").text
									SL_VarationDesc= SL_sbcl.item(x).selectSingleNode("desc2").text
									SL_VarationPrice= SL_sbcl.item(x).selectSingleNode("price1").text
									SL_VarationUnits= SL_sbcl.item(x).selectSingleNode("units").text
									SL_VarationDrop= SL_sbcl.item(x).selectSingleNode("dropship").text
									SL_VarationConst= SL_sbcl.item(x).selectSingleNode("construct").text
										
								 %>
							<option value="<%=trim(SL_Variation)%>"><%=SL_VarationDesc%>&nbsp;<%=formatcurrency(SL_VarationPrice)%>
							<%if SHOW_IN_STOCK =1 then
								if (SL_VarationUnits > 0) or (SL_VarationDrop = "true") or (SL_VarationConst = "true") then								
						      		Response.Write(" - In Stock")
								 else
								       Response.Write(" - Out of Stock")
								  end if 	
							  end if
							 %>
							 <%
							 next
							 set SL_sbcl = nothing
							 %>		
							</select>
							</td></tr>
							<%end if%>
							
							<% if LOSTOCK.inetcustom= true or LOSTOCK.giftcert= true  then %>
							<tr><td><img alt="" src="images/clear.gif" width="1" height="15" border="0"></td></tr>							
								<%if LOSTOCK.giftcert= true then %>
								<tr><td class="THHeader">&nbsp;Please enter the following information</td></tr>
								<tr><td class="plaintextbold">Recipient's Name:</td></tr>
								<tr><td><input class="plaintext" type="text" name="Recipientname" size="30" maxlength="40"></td></tr>
								<tr><td><img alt="" src="images/clear.gif" width="1" height="15" border="0"></td></tr>
								<tr><td class="plaintextbold">Notation:</td></tr>
								<tr><td><textarea name="giftcertnotation" cols="30" rows="2" value></textarea></td></tr>
								<tr><td><img alt="" src="images/clear.gif" width="1" height="15" border="0"></td></tr>
								

							
								<%else%>
									<tr><td class="THHeader"><%=trim(LOSTOCK.inetcprmpt)%></td></tr>
									<tr><td><textarea name="txtcustominfo" cols="30" rows="2" value></textarea></td></tr>
								<%end if%>							
							<%end if%>
							
							<tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0"></td></tr>
							
							<tr><td><img alt="" src="images/clear.gif" width="1" height="10" border="0"></td></tr>
							<tr><td class="plaintextbold" valign="bottom">Quantity&nbsp;<input type="text" class="plaintext" name="TXTEQUANTO" value="<%=request.querystring("altquanto")%>" size="4" MAXLENGTH="5" valign="bottom">
							<input type="hidden" name="txtebasketitem" value="<%=txtebasketitem%>">
							</td></tr>
							
							<tr><td><input type="image" value="cart"  src="images/btn-buynow.gif" border="0" align="bottom">
							
							</td></tr>

							
							</table>
						</td>
					</tr>
					</table>
					
									<script language="JavaScript">
					<!--
					var txt = "Bookmark This Page!"
					//var url = this.location;
					var url = '<%=insecureurl%>prodinfo.asp?number=<%=current_stocknumber%>';
					//var who = document.title;
					var who = '<%=althomepage%>-<%=(trim(bookmarktitle))%>';
					var ver = navigator.appName
					var num = parseInt(navigator.appVersion)
					if ((ver == "Microsoft Internet Explorer")&&(num >= 4)) {
					   document.write('<A class="allpage" HREF="javascript:window.external.AddFavorite(url,who);" ');
					   document.write('onMouseOver=" window.status=')
					   document.write("txt; return true ")
					   document.write('"onMouseOut=" window.status=')
					   document.write("' '; return true ")
					   document.write('">'+ txt + '</a>')
					}

					//-->
					</script>
					<br>
					<a class="allpage" href="javascript:openWindow();">Refer this page to a friend</a>
					
				</td>
              </form>
	</tr>	
	</table>
	
	<table width="100%" >
		<tr>
			<% 
				SLAdvanced1Val = trim(LOSTOCK.advanced1)
				if len(trim(session("SL_Advanced1"))) > 0 and len(trim(SLAdvanced1Val)) > 0 then
				%>
					<td class="plaintext"><br><b><%=session("SL_Advanced1")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=1&amp;what=<%=server.urlencode(SLAdvanced1Val)%>"><%=SLAdvanced1Val%></a></td>
				<%end if%>									
					<% 
					SLAdvanced2Val = trim(LOSTOCK.advanced2)
						if len(trim(session("SL_Advanced2"))) > 0 and len(trim(SLAdvanced2Val)) > 0 then
					%>
						<td class="plaintext"><br><b><%=session("SL_Advanced2")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=2&amp;what=<%=server.urlencode(SLAdvanced2Val)%>"><%=SLAdvanced2Val%></a></td>
					<%end if%>			
					<% 
					SLAdvanced3Val = trim(LOSTOCK.advanced3)
					if len(trim(session("SL_Advanced3"))) > 0 and len(trim(SLAdvanced3Val)) > 0 then
					%>
					<td class="plaintext"><br><b><%=session("SL_Advanced3")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=3&amp;what=<%=server.urlencode(SLAdvanced3Val)%>"><%=SLAdvanced3Val%></a></td>
					<%end if%>			
					<% 
					SLAdvanced4Val = trim(LOSTOCK.advanced4)
					if len(trim(session("SL_Advanced4"))) > 0 and len(trim(SLAdvanced4Val)) > 0 then
					%>
					<td class="plaintext"><br><b><%=session("SL_Advanced4")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=4&amp;what=<%=server.urlencode(SLAdvanced4Val)%>"><%=SLAdvanced4Val%></a></td>
					<%end if%>		
				</tr>
	  </table>
	 <br>
	 <%  if len(trim(LOSTOCK.inetfdesc)) > 0 then %>
	<table width="100%"  border="0" cellpadding="3" cellspacing="0">
			<tr><td class="THHeader">&nbsp;Detailed Description</td></tr>
			<tr><td class="plaintext">
						<p style="text-align:justify">
						<%=LOSTOCK.inetfdesc%> </p>				
						</td>
					</tr>
			</table>
	<%end if%>
	
	
	</center>
	<% SET LOSTOCK = nothing%>
	</div>
<!-- end sl_code here -->
	<br><br><br><br>
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





