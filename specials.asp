<%on error resume next%>
<!--#INCLUDE FILE = "include/momapp.asp" -->

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%	'session("department")=cstr(REQUEST.QUERYSTRING("dept"))

	session("department") = "-1"
	showthispage=request.querystring("pagenumber")
	
	
		if showthispage=0 then
		showthispage = 1
		'sort_on = "title"
		'sortorderby = "STOCK.inetsdesc"
		sort_on=""
		sortorder = ""
		sortorderby = ""

	 else
	 
	 	sort_on = request.querystring("sort_on")
		sort_by = request.querystring("sort_by")
		if sort_on="price" then
			sortorderby= "STOCK.PRICE1"
		end if 
		if sort_on="title" then
			sortorderby= "STOCK.inetsdesc"
		end if 
		if sort_on="number" then
			sortorderby= "STOCK.NUMBER"
		end if 
	
	 end if
	 
		'if len(trim(sort_on)) = 0 then
		'	sort_on= "title"
		'	sortorderby = "STOCK.inetsdesc"			
		'end if
		'if len(trim(sort_by)) = 0 then
		'	sort_by = "ASC"			
		'end if
	
	
	

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
	<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" />
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
<br>
	 <center>
	<table width="630" cellpadding="3" cellspacing="0" border="0">
		        <tr>
      <br>
  <td class="TopNavRow2Text">Specials</td>

		</tr>
	</table>
        <!-- sl code goes here -->
        <div id="page-content" class="static">
		&nbsp;
		<%
			
			xmlstring =sitelink.PRODUCTSINDEPT(session("department"),"",cstr(sortorderby),cint(PRODUCTS_PER_PAGE),cint(showthispage),sort_by )	
		%>
	<% if session("pcount") > PRODUCTS_PER_PAGE then %>
		<table width="630" border="0">
	<tr><td ALIGN="right" colspan="4">
					<br>
					<%if session("pcount")>PRODUCTS_PER_PAGE and PRODUCTS_PER_PAGE>0 then 
									totpages=int(session("pcount")/PRODUCTS_PER_PAGE) 
	    							totpages2=session("pcount")/PRODUCTS_PER_PAGE
									if totpages<>totpages2 then
										totpages=totpages+1
									end if %>
									<p class="plaintextbold">
									
									<!--start -->
									<%
									'response.write(session("pcount"))	
									dim productsy
									dim startpage
									dim limitpages
									dim nextprev
									dim nextnext
									numberofpages = PRODUCTS_PER_PAGE 'just change this value.. to number of pages to be displayed on one page i.e 1 to 15 on one page, 16-30 on next page

									if int(showthispage)>numberofpages then%>
										<%
										    nextprev=int((showthispage-numberofpages)/numberofpages)*numberofpages+1
										%>
									
										<a HREF="specials.asp?pagenumber=<%=nextprev%>&amp;sort_on=<%=sort_on%>&amp;sort_by=<%=sort_by%>">&nbsp;Previous&nbsp;</a><<&nbsp;&nbsp; <%	startpage=nextprev+numberofpages  'int(session("showthispage"))
										if int(showthispage)+ (numberofpages-1) > totpages then
											limitpages=totpages
											'startpage = nextprev
											if startpage => totpages then
												startpage=nextprev												
											end if 
											
										else
											nextnext=startpage+ numberofpages
											limitpages=startpage+ (numberofpages-1)
										end if		
									else
										nextprev=2
										startpage=1
										nextnext= numberofpages+1
										if totpages<=numberofpages then
											limitpages=totpages
										else
											limitpages=numberofpages
										end if		
									end if%> <b>Page</b>&nbsp;&nbsp;
									<%
									for productsy=startpage to limitpages%>
										<%if productsy=int(showthispage) then%>
											<font class="plaintextbold"><b><%=productsy%></b></font>&nbsp;
											 <% if productsy < limitpages then %>
												|
											 <%end if%>
										<%else%>
											<a HREF="specials.asp?pagenumber=<%=productsy%>&amp;sort_on=<%=sort_on%>&amp;sort_by=<%=sort_by%>"><%=productsy%></a>&nbsp;
											<% if (productsy < limitpages) then %>
											| 
											<%end if%>
										<%end if%>
									<%next%>
									<%if limitpages<totpages then%>
										&nbsp;&nbsp;&gt;&gt;<a HREF="specials.asp?pagenumber=<%=nextnext%>&amp;sort_on=<%=sort_on%>&amp;sort_by=<%=sort_by%>"><font class="plaintextbold">&nbsp;More</font></a> 									
									<%end if%>
									
								<%end if%>
					</td></tr>
	</table>		
	<%end if%>
	<div align="center">
	<table border="0" cellspacing="1" cellpadding="3">
	<tr><td valign="middle" class="plaintext">Sort By:</td>
		<td align="left" valign="top" class="plaintext">
			<table width="100%" cellpadding="1" cellspacing="2" border="0">
			<tr align="center" valign="middle">
			  <td >
				<% if sort_on="number" then %>
					<span class="plaintext"><b>Item #</b></span>
					<%
					 if sort_by = "ASC"  then %>
					<a HREF="specials.asp?pagenumber=1&amp;sort_on=number&amp;sort_by=DESC">
					<img src="images/down1.gif" width="11" height="7" border="0" alt="list by item Number descending">
					</a>
				<%else%>
				<a HREF="specials.asp?pagenumber=1&amp;sort_on=number&amp;sort_by=ASC">
					<img src="images/up1.gif" width="11" height="7" border="0" alt="list by item Number ascending">
				</a>
				<%end if
				else
				%>
				<a HREF="specials.asp?pagenumber=1&amp;sort_on=number&amp;sort_by=ASC">
				<span class="plaintext">Item #</span>
				</a>
				<%end if %>
				</td>
				<td >&nbsp;&nbsp;&nbsp;</td>
				<td >
			
				<% if sort_on="title" then %>				
				  <span class="plaintext"><b>Product Name</b></span>
				  <%
					if sort_by = "ASC" then %>
					<a HREF="specials.asp?pagenumber=1&amp;sort_on=title&amp;sort_by=DESC">
					<img src="images/down1.gif" width="11" height="7" border="0" alt="list by title ascending">
					</a>
				<%else%>
					<a HREF="specials.asp?pagenumber=1&amp;sort_on=title&amp;sort_by=ASC">
					<img src="images/up1.gif" width="11" height="7" border="0" alt="list by title descending">
				   </a>
				<%end if 
				else
				%>
				<a HREF="specials.asp?pagenumber=1&amp;sort_on=title&amp;sort_by=ASC">
				<span class="plaintext">Product Name</span>
				</a>
				<%end if%>

			</td>
			    <td >&nbsp;&nbsp;&nbsp;</td>
			    <td >
			
			<% if sort_on="price" then %>
				<span class="plaintext"><b>Price</b></span>
				<%
				if sort_by = "ASC" then %>				
				<a HREF="specials.asp?pagenumber=1&amp;sort_on=price&amp;sort_by=DESC">
					<img src="images/down1.gif" width="11" height="7" border="0" alt="list by price ascending">
				</a>
			<%else%>
				<a HREF="specials.asp?pagenumber=1&amp;sort_on=price&amp;sort_by=ASC">
					<img src="images/up1.gif" width="11" height="7" border="0" alt="list by price descending">
				 </a>
			<%end if 
			else
			%>
			<a HREF="specials.asp?pagenumber=1&amp;sort_on=price&amp;sort_by=ASC">
			<span class="plaintext">Price</span>
			</a>
			<%end if%>

			</td>
	</tr>
	</table>
        </div> 
<table width="630" border="0" cellspacing="1" cellpadding="3">
			
	<%
	
	
	objDoc.loadxml(xmlstring)
	
	If objDoc.parseError.errorCode <> 0 Then
	%>	
	<tr><td colspan="3" class="plaintextbold" align="center"><br>Error Processing<br><br></td></tr>
	<%else%>
	
	
	<%
	set SL_prod = objDoc.selectNodes("//prodindept") 
	
	if SL_prod.length = 0 then
	%>	
	<tr><td colspan="3" class="plaintextbold" align="center">	
	<br>
	<br>
	No Products found<br><br></td></tr>
	<%end if%>
	</table>

	<!--new code -->
	<%If objDoc.parseError.errorCode = 0 Then %>
	<%set SL_prod = objDoc.selectNodes("//prodindept") 
	NUM_OF_PROD_PER_LINE = 3
	rowcount = 0
	%>
	<table width="550" border="0" align="center" cellpadding="3" cellspacing="1">
      <%
	for x=0 to SL_prod.length-1		
		 SLProdNumber = SL_prod.item(x).selectSingleNode("number").text
		 'SLVariation  = SL_prod.item(x).selectSingleNode("variant").text
		 SLProductTitle  = SL_prod.item(x).selectSingleNode("inetsdesc").text
		 SL_Size_color  = SL_prod.item(x).selectSingleNode("size_color").text
		 SL_thumbnail  = SL_prod.item(x).selectSingleNode("inetthumb").text
		 SL_Fullimage  = SL_prod.item(x).selectSingleNode("inetimage").text
		 SL_price  = SL_prod.item(x).selectSingleNode("price1").text		 
		 SL_ComparePrice  = SL_prod.item(x).selectSingleNode("inetcprice").text
		 SL_Units  = SL_prod.item(x).selectSingleNode("units").text
		 SL_Drop  = SL_prod.item(x).selectSingleNode("drop").text
		 SL_Contruct  = SL_prod.item(x).selectSingleNode("construct").text
		 SL_Shortdesc  = SL_prod.item(x).selectSingleNode("inetshortd").text
		 SL_Pref_Ship = SL_prod.item(x).selectSingleNode("prefship").text
		 
		 if rowcount=NUM_OF_PROD_PER_LINE then
			rowcount=0
		%>
      <p>
      <tr>
        <%end if
		rowcount= rowcount + 1
		
		if len(trim(SL_thumbnail)) = 0 then
			SL_thumbnail = SL_Fullimage
		end if
		
		 
			 StrFileName = "images/"+ SL_thumbnail  
			 StrPhysicalPath = Server.MapPath(StrFileName)
			     set objFileName = CreateObject("Scripting.FileSystemObject")	
					 if objFileName.FileExists(StrPhysicalPath) then
			  			imagename=StrFileName
				    else
						imagename="images/noimage.gif"
  					end if 
				set objFileName = nothing
		%>
        <td width="<%=(100/NUM_OF_PROD_PER_LINE)%>%" valign="bottom" bordercolor="0"><table width="100%" cellpadding="2" cellspacing="0" border="0">
            <tr>
              <td align="center"></td>
            </tr>
            <tr>
              <td align="left"><div align="center"><a onMouseover="ddrivetip('<%=SL_Shortdesc%>','yellow', 300)"; onMouseout="hideddrivetip()" HREF="prodinfo.asp?number=<%=SLProdNumber%>"> <img SRC="<%=imagename%>" alt="<%if len(trim(SL_Shortdesc))<>0 then response.write(SL_Shortdesc) else response.write(SLProductTitle) end if%>" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %> border="0"><br>
                        <span class="Plaintextbold">
                        <%if SHOW_COMP_PRICE="TRUE" and ( discountedUnitPrice > 0) then %>
                        <%end if%>
                        <%if SHOW_DUE_DATE = 0 and SHOW_IN_STOCK =1 and SL_Size_color = "false" then 	
			if (SL_Units > 0) or (SL_Drop = "true") or (SL_Contruct = "true") then
	    		avail_status="In Stock"				
			else
				avail_status="Out of Stock"		
			end if
		%>
                        <%=avail_status%>
                        <%end if%>
                    </span></a></div></td>
            </tr>
            <tr>
              <td align="left"><a class="allpage" onMouseover="ddrivetip('<%=SL_Shortdesc%>','yellow', 315)"; onMouseout="hideddrivetip()" HREF="prodinfo.asp?number=<%=SLProdNumber%>"><%=SLProductTitle%></a></td>
            </tr>
            <tr>
              <td align="left" class="plaintext"><table width="100%"  border="0" cellspacing="0">
                  <%if SHOW_COMP_PRICE="TRUE" and ( SL_ComparePrice > 0) then %>
                  <tr>
                    <td class="plaintext">List Price :<s><%=FORMATCURRENCY( SL_ComparePrice)%></s></td>
                  </tr>
                  <%end if%>
                  <tr>
                    <td class="CompPrice">Our Price: <%=formatcurrency(SL_price)%></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td align="left"><!--- <a HREF="prodinfo.asp?number=<%=SLProdNumber %>"><img src="images/more_info.gif" width="62" height="15" border="0"></a>---></td>
            </tr>
            <tr>
              <td align="left" ></td>
            </tr>
        </table></td>
        <%next%>
      </tr>
    </table>
	<%set SL_prod =  nothing%>
	<!--end new code -->
	
	<%end if 'If objDoc.parseError.errorCode =0 %>
<%end if%>
	</table>

	<!-- end sl_code here -->
	</td>
	<!-- <td width="1" class="THHeader"><img src="images/clear.gif" width="1" border="0"></td> -->
</tr>

</table></td>
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
