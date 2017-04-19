<%on error resume next%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<% 
Response.ExpiresAbsolute = Now() - 1 
'Response.AddHeader "Cache-Control", "must-revalidate" 
'Response.AddHeader "Cache-Control", "no-store" 
%> 

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->



<%	

	session("department") = 0
	showthispage=request.querystring("pagenumber")
	
	if isnumeric(showthispage)= false then
	    showthispage = 0 
	end if
	
	session("val")=request.querystring("val")
	
	if isnumeric(session("val"))= false then		
			session("val")=0
	end if

	if len(trim(session("val")))= 0 then		
		session("val")=0
	end if

	sortorder = ""
	sortorderby = ""
	
	    'session("searchcriteria") = REQUEST.FORM("ProductSearchBy")
		if showthispage=0 then
			showthispage = 1
		
			session("searchstring")=cstr(trim(REQUEST.FORM("txtsearch")))

            ProductSearchBy = REQUEST.FORM("ProductSearchBy")
			
			if isnumeric(ProductSearchBy) = false then
			    ProductSearchBy = 1
			end if
			
			
			'if REQUEST.FORM("ProductSearchBy") =2 then
				session("searchcriteria")= "DESC"
				
                build_searchcriteria=""			
				if SEARCH_ON_TITLE=1 then
				    build_searchcriteria = "stock.number+STOCK.inetsdesc+"
				end if

				if SEARCH_ON_SHORTDESC=1 then
				    build_searchcriteria = build_searchcriteria + "STOCK.inetshortd+"
				end if

				if SEARCH_ON_FULLDESC=1 then
				    build_searchcriteria = build_searchcriteria + "STOCK.inetfdesc+"
				end if

				if SEARCH_ON_KEYWORD=1 then
				    build_searchcriteria = build_searchcriteria + "STOCK.inetkeywrd+"
				end if
				
				build_searchcriteria = build_searchcriteria + "STOCK.number+"
								
				Lensearch_on = len(trim(build_searchcriteria))
				build_searchcriteria = mid(build_searchcriteria,1,Lensearch_on-1)
				'session("search_on") = "STOCK.inetsdesc+STOCK.inetfdesc+STOCK.inetkeywrd+stock.inetshortd"
				session("search_on") = build_searchcriteria
			'end if
		
			
			'***DeActivate this code when using search.asp*******************************			
			if ProductSearchBy =3 then
				session("searchcriteria")= session("SL_Advanced1")
				session("search_on") = "STOCK.Advanced1"						
			end if
			if ProductSearchBy =4 then
				session("searchcriteria")= session("SL_Advanced2")
				session("search_on") = "STOCK.Advanced2"						
			end if
			if ProductSearchBy = 5 then
				session("searchcriteria")= session("SL_Advanced3")
				session("search_on") = "STOCK.Advanced3"						
			end if
		
			if ProductSearchBy =6 then
				session("searchcriteria")= session("SL_Advanced4")
				session("search_on") = "STOCK.Advanced4"						
			end if				

			'**********************************************************
		else		
			session("searchstring")=request.querystring("searchstring")
		end if
				
		if session("val")=1 then
			sortorderby = "STOCK.PRICE1"
			sort_by = "ASC"
		end if
		if session("val")=2 then
			sortorderby = "STOCK.PRICE1"
			sort_by = "DESC"
		end if
		if session("val")=3 then
			sortorderby = "STOCK.INETSDESC"
			sort_by = "ASC"
		end if
		if session("val")=4 then
			sortorderby = "STOCK.INETSDESC"
			sort_by = "DESC"
		end if
		if session("val")=5 then
			sortorderby = "STOCK.NUMBER"
			sort_by = "ASC"
		end if
		if session("val")=6 then
			sortorderby = "STOCK.NUMBER"
			sort_by = "DESC"
		end if
		if session("val")=7 then
			sortorderby = "bestseller.qty_ord"
			sort_by = "DESC"
		end if
		if session("val")=8 then
			sortorderby = "rating"
			sort_by = "DESC"
		end if
		

		
		session("Current_selectedVariant")=""
		
		
		session("searchstring")=fix_xss_Chars(session("searchstring")) 
		
		if len(trim(session("searchstring"))) = 0 then
		    set sitelink=nothing
		    set ObjDoc=nothing
		    Response.Status="301 Moved Permanantly"
		    Response.AddHeader "Location", insecureurl&"default.asp"
		    Response.End
		end if
		
		
		 IsValidUrl = true
         CorrectProductUrl = ""
         UrlWriteStr = request.ServerVariables("http_x_rewrite_url")
         IsValidUrl = sitelink.ValidateProdInfoUrl("",UrlWriteStr,CorrectProductUrl) 
         
         if IsValidUrl=false then    
            set sitelink=nothing
		    set ObjDoc=nothing
		    Response.Status="301 Moved Permanantly"
		    Response.AddHeader "Location", insecureurl&"default.asp"
		    Response.End
        end if
		
	%>
	



<html>
<head>
<title>Aussie Products.com | Product Search @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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

<!-- Latest compiled and minified CSS -->
<link rel="stylesheet" href="design/styles/julius-css.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
<link rel="stylesheet" href="design/jstyles.css">
<link rel="stylesheet" href="design/styles/jstyles.css">
<link rel="stylesheet" href="design/bower_components/font-awesome/css/font-awesome.css">

<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />

<script language="JavaScript">
<!-- hide from JavaScript-challenged browsers

function RedirectToAddADDress (page)
   {
   var new_url=page;	  
   if (  (new_url != "")  &&  (new_url != null)  )
   //document.basketform.submit();
    window.location=new_url;
   }
//-->hiding
</script>
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->

<div id="container">
    <!--#INCLUDE FILE = "include/top_nav-copy.asp" -->
    <div id="main">

<% session("destpage")="" 
  session("viewpage")=session("destpage")
%>

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
	
	<br>
	<table width="100%"  border="0" cellpadding="0" cellspacing="0">
				<tr><td class="TopNavRow2Text">
						Search Results for : &quot;<%=session("searchstring")%>&quot;
									
				</tr>
	</table>

  <%
  	if showthispage=0 then
		showthispage = 1
	end if
		

	

	extrafield=""
	set SLProdObj = New cProductObj
	SLProdObj.ldeptcode    		= ""
	SLProdObj.lSORTBY		 	= cstr(sortorderby)
	SLProdObj.lProd_Per_Page 	= cint(PRODUCTS_PER_PAGE)
	SLProdObj.llnpage		 	= showthispage
	SLProdObj.lsortorder	 	= cstr(sort_by)
	SLProdObj.llextrafield	 	= extrafield
	SLProdObj.lsearchstr	 	= cstr(session("searchstring"))
	SLProdObj.lsearch_on	 	= cstr(session("search_on"))
	
			
	xmlstring =sitelink.SEARCHPRODUCTS(SLProdObj)
	pcount=SLProdObj.lpcount
	set SLProdObj = nothing
	
    objDoc.loadxml(xmlstring)
			
	If objDoc.parseError.errorCode = 0 Then	
	     'get special pricing.
	       all_prods=""
	        set SL_prod = objDoc.selectNodes("//srchprod") 
	        for x=0 to SL_prod.length-1
		        SLProdNumber = SL_prod.item(x).selectSingleNode("number").text
	            all_prods =all_prods+","+cstr(SLProdNumber)
		     next
		     all_prods =all_prods+","
        	
		     xmlstring = sitelink.GET_ALL_SPECIALPRICES(cstr(all_prods),session("ordernumber"),session("shopperid"),cstr(""))
		        
		    objDoc.loadxml(xmlstring)
		        
	end if	
			
	%>
	<br>
	<table width="100%"  border="0" cellpadding="0" cellspacing="0">
	<tr><td class="plaintextbold">
			<%if pcount > 0 then %>
			<select name="listby" class="smalltextblk" onChange="RedirectToAddADDress(this.options[this.selectedIndex].value)">
				<option value="">Sort by
				<option value="searchprods.asp?val=1&pagenumber=1&searchstring=<%=server.urlencode(session("searchstring"))%>" <%if session("val")=1 then response.write(" selected") end if%>>Price Low->High
				<option value="searchprods.asp?val=2&pagenumber=1&searchstring=<%=server.urlencode(session("searchstring"))%>" <%if session("val")=2 then response.write(" selected") end if%>>Price High->Low
				<option value="searchprods.asp?val=3&pagenumber=1&searchstring=<%=server.urlencode(session("searchstring"))%>" <%if session("val")=3 then response.write(" selected") end if%>>Title A->Z
				<option value="searchprods.asp?val=4&pagenumber=1&searchstring=<%=server.urlencode(session("searchstring"))%>" <%if session("val")=4 then response.write(" selected") end if%>>Title Z->A
				<option value="searchprods.asp?val=5&pagenumber=1&searchstring=<%=server.urlencode(session("searchstring"))%>" <%if session("val")=5 then response.write(" selected") end if%>>Item Code A->Z
				<option value="searchprods.asp?val=6&pagenumber=1&searchstring=<%=server.urlencode(session("searchstring"))%>" <%if session("val")=6 then response.write(" selected") end if%>>Item Code Z->A
				<option value="searchprods.asp?val=7&pagenumber=1&searchstring=<%=server.urlencode(session("searchstring"))%>" <%if session("val")=7 then response.write(" selected") end if%>>Popularity		
				<option value="searchprods.asp?val=8&pagenumber=1&searchstring=<%=server.urlencode(session("searchstring"))%>" <%if session("val")=8 then response.write(" selected") end if%>>Top Rated
			</select>
		</td>
		<td align="right" id="pagelinks" class="plaintextbold">		
					<%
					if pcount>PRODUCTS_PER_PAGE and PRODUCTS_PER_PAGE>0 then 
						totpages=int(pcount/PRODUCTS_PER_PAGE) 
						totpages2=pcount/PRODUCTS_PER_PAGE
						if totpages<>totpages2 then
							totpages=totpages+1
						end if 
                        
                        lastpagelink = "searchprods.asp?searchstring=" & server.urlencode(session("searchstring")) & "&amp;pagenumber=" & totpages & "&val=" & session("val")
									%>
									
									<!--start -->
									<%if int(showthispage) > 1 then 
                                        firstpagelink = "searchprods.asp?searchstring=" & server.urlencode(session("searchstring")) & "&amp;pagenumber=1" & "&val=" & session("val")
									%>
									<a HREF="<%=firstpagelink%>" class="previous">FIRST</a>&nbsp;&nbsp; <a href="searchprods.asp?searchstring=<%=server.urlencode(session("searchstring"))%>&amp;pagenumber=<%=showthispage-1%>&val=<%=session("val")%>" class="arrowleft"></a>
									<%end if %>
									<%
									'response.write(pcount)	
									dim productsy
									dim startpage
									dim limitpages
									dim nextprev
									dim nextnext
									numberofpages = 6 'just change this value.. to number of pages to be displayed on one page i.e 1 to 15 on one page, 16-30 on next page
									'response.write("numberofpages" & numberofpages)
									if int(showthispage)>numberofpages then%>
										<%
								 			modval = showthispage mod numberofpages
			  								cintval = int((showthispage-numberofpages)/numberofpages)
											if modval = 0 then
												cintval = cintval-1
											end if
										    'nextprev=int((showthispage-numberofpages)/numberofpages)*numberofpages+1
											nextprev=cintval*numberofpages+1
										%>
									
										<%	startpage=nextprev+numberofpages  'int(session("showthispage"))
										if int(showthispage)+ (numberofpages-1) > totpages then
											if modval = 0 then
												nextnext=startpage+ numberofpages
												limitpages=startpage+ (numberofpages-1)
											else
												limitpages=totpages
											end if 
																						
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
									end if
									
									%> <!--<b>Page</b>&nbsp;&nbsp;-->
									<%
									
									for productsy=startpage to limitpages%>									
										<%if productsy=int(showthispage) then%>
											<%=productsy%>&nbsp;
										<%else%>
											<a HREF="searchprods.asp?searchstring=<%=server.urlencode(session("searchstring"))%>&amp;pagenumber=<%=productsy%>&val=<%=session("val")%>"><%=productsy%></a>&nbsp;
										<%end if%>
									<%next%>
									
									<%if int(showthispage) < totpages then %>
									    <%if limitpages < totpages then %>
									        ... <a href="<%=lastpagelink%>"><%=totpages%></a>&nbsp;
									    <% end if %>
									<a HREF="searchprods.asp?searchstring=<%=server.urlencode(session("searchstring"))%>&amp;pagenumber=<%=showthispage+1%>&val=<%=session("val")%>" class="arrowright"></a>&nbsp;
									<a href="<%=lastpagelink%>" class="next">LAST</a>&nbsp;
									<%end if %>
									
								<%end if%>			
			
			
			<%end if%>
		</td>
		<td width="10">&nbsp;</td>
	</tr>
	
	</table>
	
	<%If objDoc.parseError.errorCode <> 0 Then %>
	    <table>
	        <tr><td colspan="4" class="plaintextbold" align="center"><br>Error Processing<br><br></td></tr>
	    </table>
	
	<%end if %>
	
	 <%
	 if objDoc.parseError.errorCode = 0 then
	 if SHOW_PRODUCTSLISTVIEW=1 then %>
    			
	<table width="100%"  border="0" cellspacing="1" cellpadding="3">
		<tr>
			<th align="CENTER" class="THHeader" width="15%"><span class="THHeader">Item #</span></th>
			<th align="CENTER" class="THHeader" width="10%">Image</th>
			<th align="CENTER" class="THHeader" width="60%"><span class="THHeader">Description</span></th>
			<th align="CENTER" class="THHeader" width="15%"><span class="THHeader">Price</span></th>
		</tr>
			
	<%
		
	if SL_prod.length = 0 then
	%>	
	<tr><td colspan="4" class="plaintextbold" align="center">	
	<br>
	<br>
	No Products found<br><br></td></tr>
	<%end if%>
	<%
	
		for x=0 to SL_prod.length-1		
		 SLProdNumber = SL_prod.item(x).selectSingleNode("number").text
		 SLURLCodeNumber= server.URLEncode(SLProdNumber)
		 'SLVariation  = SL_prod.item(x).selectSingleNode("variant").text
		 SLProductTitle  = SL_prod.item(x).selectSingleNode("inetsdesc").text
		 SL_Size_color  = SL_prod.item(x).selectSingleNode("size_color").text
		 SL_thumbnail  = SL_prod.item(x).selectSingleNode("inetthumb").text
		 SL_Fullimage  = SL_prod.item(x).selectSingleNode("inetimage").text
		 SL_price  = SL_prod.item(x).selectSingleNode("price1").text		 
		 SL_ComparePrice  = SL_prod.item(x).selectSingleNode("inetcprice").text
		 SL_Units  = SL_prod.item(x).selectSingleNode("units").text
		 SL_Drop  = SL_prod.item(x).selectSingleNode("dropship").text
		 SL_Contruct  = SL_prod.item(x).selectSingleNode("construct").text
		 SL_Shortdesc  = SL_prod.item(x).selectSingleNode("inetshortd").text
		 SL_Pref_Ship = SL_prod.item(x).selectSingleNode("prefship").text
		 SL_Discont = SL_prod.item(x).selectSingleNode("discont").text
		 
		 SL_GIFTCERT = SL_prod.item(x).selectSingleNode("giftcert").text
		 
		 SL_GiftCard       =SL_prod.item(x).selectSingleNode("giftcard").text
		 
		 SL_MinPrice       =SL_prod.item(x).selectSingleNode("minprice").text
		 SL_MaxPrice       =SL_prod.item(x).selectSingleNode("maxprice").text

		 
		 
		 if SL_GIFTCERT ="true"  then
		 	is_giftcert = true
		 else
		 	is_giftcert = false
		 end if
		 
		 
		 
		 target_in_SpecialPrice = "//gsp[number='"+cstr(SLProdNumber)+"']"
		 has_special_Price= false
		 has_Onespecial_Price=false
		 set SL_prodSpecial = objDoc.selectNodes(target_in_SpecialPrice) 
		 	if SL_prodSpecial.length > 0 then
			 	has_special_Price=true
				
				SL_SP_qty = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("qty").text
				SL_SP_ordertotal = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("total_dol").text
				SL_SP_qty = replace(SL_SP_qty,".00","")
				if SL_SP_qty= 1 and (cdbl(session("SL_BasketSubTotal")) >= cdbl(SL_SP_ordertotal) )then
					SL_SP_price = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("price").text
					SL_SP_discount = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("discount").text

					SL_SP_CostMeth = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("costmethod").text
					SL_SP_CostPlusDiscount = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("costplus").text
					SL_SP_StockReOrdPrice = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("reordprice").text
					SL_SP_UnCostprice = SL_prodSpecial.item(SL_prodSpecial.length-1).selectSingleNode("uncost").text
					
					if SL_SP_CostMeth="" then
						unitprice = SL_price
						unitprice2 = SL_SP_price
						if unitprice2>0 then
							unitprice = unitprice2
						end if
						percentoff = (SL_SP_discount/100.0)
						discountedUnitPrice = cdbl(unitprice * (1- percentoff))									
					end if
									
					if SL_SP_CostMeth="P" then
							discountedUnitPrice= cdbl(SL_SP_StockReOrdPrice*(1+SL_SP_CostPlusDiscount/100.0))
					end if
							
					if SL_SP_CostMeth="U" then										
							discountedUnitPrice= cdbl(SL_SP_UnCostprice*(1+SL_SP_CostPlusDiscount/100.0))									
					end if
						
					has_Onespecial_Price=true
				end if
			end if
		 set SL_prodSpecial = nothing
		 
		 
		 if SL_Pref_Ship<>"" then								
			SL_prefShip_Title	=trim(SL_prod.item(x).selectSingleNode("ca_title").text)
  		end if
		
		if SL_thumbnail="" then
			SL_thumbnail = SL_Fullimage
		end if
		
		 
		if len(SL_thumbnail) > 0 then
			 StrFileName = "images/"+ SL_thumbnail  
			 StrPhysicalPath = Server.MapPath(StrFileName)
			     set objFileName = CreateObject("Scripting.FileSystemObject")	
					 if objFileName.FileExists(StrPhysicalPath) then
			  			imagename=StrFileName
				    else
						imagename="images/noimage.gif"
  					end if 
				set objFileName = nothing
	    else
	            imagename="images/noimage.gif"
	    end if 
		
				if (x mod 2) = 0 then
					class_to_use = "tdRow1Color"
				else
					class_to_use = "tdRow2Color"
				end if 
		

		if WANT_REWRITE = 1 then
			 SL_urltitle = SLProductTitle
			 SL_urltitle = url_cleanse(SL_urltitle)
			 prodlink = insecureurl  + SL_urltitle + "/productinfo/" + SLURLCodeNumber + "/" 
		else 
			 prodlink = insecureurl +  "prodinfo.asp?number=" + SLProdNumber
		end if
		
		usethis_rating_img = "images/clear.gif"
        if SHOW_PRODUCTREVIEWS=1 then
         SL_rating = SL_prod.item(x).selectSingleNode("rating").text
         SL_rating =replace(SL_rating,".00","")
         'response.Write(SL_rating)
         'SL_rating = replace(SL_rating,".00","")
            if cdbl(SL_rating) > 0 then              
                pos=InStr(SL_rating,".")                
                if pos > 0 then                    
                    decimal_part = mid(SL_rating,pos)
                    avg_rating = mid(SL_rating,1,pos-1)
                    if len(decimal_part) > 3 then 
                        decimal_part = mid(decimal_part,1,3)
                     end if
                    'Response.Write avg_rating & "," & decimal_part & ","
                    if decimal_part > 0.01 and decimal_part < 0.5 then 
                    
    					avg_rating=avg_rating&"-5"
					else
						avg_rating=avg_rating+1
					end if
					
					usethis_rating_img = "images/"&avg_rating&+"note.gif"      
			    else
			        avg_rating = SL_rating
			        usethis_rating_img = "images/"&avg_rating&+"note.gif"      			        
			        
                end if 'if pos > 0 then
            end if   
       end if 
		
	%>	
	<tr>	
	<td class="<%=class_to_use%>" valign="middle" align="center"><a title="<%=SLProductTitle%>" class="allpage" href="<%=prodlink%>"><%=SLProdNumber%></a>	
	</td>
	<td class="<%=class_to_use%>" valign="middle" align="center"><div class="prodthumb"><div class="prodthumbcell"><a href="<%=prodlink%>"><img SRC="<%=imagename%>" alt="<%=SLProductTitle%>" title="<%=SLProductTitle%>" class="prodlistimg" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %> border="0"></a></div></div></td>
	<td class="<%=class_to_use%>" valign="top">
	<table width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr><td align="left"><img src="<%=usethis_rating_img %>" /></td></tr>
	<tr><td><a title="<%=SLProductTitle%>" class="producttitlelink" href="<%=prodlink%>"><%=SLProductTitle%></a></td></tr>
	<tr><td class="plaintext"><%=SL_Shortdesc%></td></tr>
	<%if SHOW_COMP_PRICE="TRUE" and ( SL_ComparePrice > 0) then %>
	<tr><td class="CompPrice">Compare At: <%=FORMATCURRENCY( SL_ComparePrice)%></td></tr>
	<%end if%>
	
	<%if has_special_Price=true and has_Onespecial_Price=false then %>
	<tr><td><a class="allpage" href="<%=prodlink%>">Special Pricing Available</a></td></tr>
	<%end if%>
	
	<%if  SHOW_IN_STOCK =1 and SL_Size_color = "false" then 	
			if (SL_Units > 0) or (SL_Drop = "true") or (SL_Contruct = "true") then
	    		avail_status="In Stock"				
			else
				if SHOW_DUE_DATE=1 then
					avail_status="Check for Availability"
				else				
					avail_status="Out of Stock"						
				end if
			end if
		%>
		<%if SHOW_DUE_DATE=1 and SL_Units=0 then%>
			<tr><td class="plaintext"><a title="<%=SLProductTitle%>" class="allpage" href="<%=prodlink%>"><%=avail_status%></a></td></tr>
		<%else%>
			<tr><td class="plaintext"><%=avail_status%></td></tr>
		<%end if%>
	<%end if%>
	
	
	
	<% if SL_Pref_Ship<>"" then %>				
	<tr><td class="plaintext">Will be shipped via: <%=SL_prefShip_Title%></td></tr>
	<%end if%>
	
	<%if USE_ADVANCED_SEARCH=1 then %>
	<tr><td class="plaintext">
		
		<% 
		SLAdvanced1Val = SL_prod.item(x).selectSingleNode("advanced1").text
		if len(trim(session("SL_Advanced1"))) > 0 and SLAdvanced1Val<>"" then
		%>
		<b><%=session("SL_Advanced1")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=1&amp;what=<%=server.urlencode(SLAdvanced1Val)%>"><%=SLAdvanced1Val%></a>
		<%end if%>		
		
	</td></tr>
	<tr><td class="plaintext">
		<% 
		SLAdvanced2Val = SL_prod.item(x).selectSingleNode("advanced2").text
		if len(trim(session("SL_Advanced2"))) > 0 and SLAdvanced2Val<>"" then
		%>
		<b><%=session("SL_Advanced2")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=2&amp;what=<%=server.urlencode(SLAdvanced2Val)%>"><%=SLAdvanced2Val%></a>
		<%end if%>		
	
	</td></tr>
	<tr><td class="plaintext">
		<% 
		SLAdvanced3Val = SL_prod.item(x).selectSingleNode("advanced3").text
		if len(trim(session("SL_Advanced3"))) > 0 and SLAdvanced3Val<>"" then
		%>
		<b><%=session("SL_Advanced3")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=3&amp;what=<%=server.urlencode(SLAdvanced3Val)%>"><%=SLAdvanced3Val%></a>
		<%end if%>		
	
	</td></tr>
	<tr><td class="plaintext">
		<% 
		SLAdvanced4Val = SL_prod.item(x).selectSingleNode("advanced4").text
		if len(trim(session("SL_Advanced4"))) > 0 and SLAdvanced4Val<>"" then
		%>
		<b><%=session("SL_Advanced4")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=4&amp;what=<%=server.urlencode(SLAdvanced4Val)%>"><%=SLAdvanced4Val%></a>
		<%end if%>		
	</td></tr>
	<%end if%>

	</table>
	
	</td>
	<td class="<%=class_to_use%>" align="center" valign="middle">
	<form enctype="application/x-www-form-urlencoded" method="post" action="itemadd.asp">
	<%if SL_Discont="true" and SL_Units=0 then %>
		<span class="plaintext">This item is no longer available</span>
	<%else%>
	<% if SL_Size_color  = "true" or is_giftcert=true THEN%>
	
	    <%if SHOW_MINMAXPRICE=1 then 
		    ShowRangePrice = false
		    if cdbl(SL_MinPrice) <> cdbl(SL_MaxPrice) then
		        ShowRangePrice= true
		    end if
	     %>
	        <% if ShowRangePrice= true then %>
	            <%=formatcurrency(SL_MinPrice)%> - <%=formatcurrency(SL_MaxPrice)%>
	        <%else %>
	            <%=formatcurrency(SL_MinPrice)%>
	        <%end if %>
	    <%else %>
	        <a class="allpage" HREF="<%=prodlink%>"><u>Options</u></a>
	    <%end if %>
        <%if has_special_Price=true then %>
            <img src="images/RedSaleTag.gif" border="0" align="absmiddle" hspace="6" />
        <%end if %>
		
	<%else%>	
		<%if has_Onespecial_Price=true then %>
			<span class="ProductPrice"><s><%=formatcurrency(SL_price)%></s></span>
			<br>
			<table width="0" cellpadding="0" cellspacing="0" border="0">
			<tr><td class="nopadding"><img src="images/RedSaleTag.gif" style="border:0; width:53; height:23;" alt="sale"></td>
				<td class="plaintext nopadding" align="center" style="background-image:url(images/RedSaleTag_bkg.gif);height:23">
				<span style="font-weight:bold;color:White;"><%=formatcurrency(discountedUnitPrice)%></span>&nbsp;
				</td>
			</tr>
			</table>			
			
		<%else%>
			<span class="ProductPrice"><%=formatcurrency(SL_price)%></span>
            <%if has_special_Price=true then %>
                <img src="images/RedSaleTag.gif" border="0" align="absmiddle" hspace="6" />
            <%end if %>
		<%end if %>

		  <%if SHOW_QTY=1 then %>
		  		<br><br>
				<table width="0" cellpadding="3" cellspacing="0" border="0">
				<tr><td align="center">
				<input type="hidden" name="additem" value="<%=SLProdNumber%>">	  								  		
				<%if SL_GiftCard= "true" then %>
				    <input type="hidden" class="plaintext" name="txtquanto" value="1" size="4"  align="middle">
				<%else %>				
			  	<input type="text" class="plaintext" name="txtquanto" value="1" size="4"  align="middle">				
				<%end if %>
				</td></tr>
				<tr><td align="center"><input TYPE="image" SRC="images/btn-buy.gif" style="border:0;" ALT="Add to Basket"  align="middle"></td></tr>
				</table>			  		 	       
	  	  <%end if%>	  	
	<%end if%>
	<%end if%>
	</form>
	</td>
	
	
	</tr>		
	<%next%>
	</table>
	<%
	set SL_prod = nothing
	%>	
	
	<%else 'SHOW_PRODUCTSLISTVIEW %>
	
	<table width="100%" border="0" cellspacing="1" cellpadding="3" class="grid">
	
	<%
	
	
	if SL_prod.length = 0 then
	%>	
	<tr><td colspan="4" class="plaintextbold" align="center" >	
	<br>
	<br>
	No Products found<br><br></td></tr>
	<%end if%>
	<br>
	<%
		colcount=0
		
		
		
		for x=0 to SL_prod.length-1		
		 SLProdNumber = SL_prod.item(x).selectSingleNode("number").text
		 SLURLCodeNumber= server.URLEncode(SLProdNumber)
		 'SLVariation  = SL_prod.item(x).selectSingleNode("variant").text
		 SLProductTitle  = SL_prod.item(x).selectSingleNode("inetsdesc").text
		 SL_Size_color  = SL_prod.item(x).selectSingleNode("size_color").text
		 SL_thumbnail  = SL_prod.item(x).selectSingleNode("inetthumb").text
		 SL_Fullimage  = SL_prod.item(x).selectSingleNode("inetimage").text
		 SL_price  = SL_prod.item(x).selectSingleNode("price1").text		 
		 SL_ComparePrice  = SL_prod.item(x).selectSingleNode("inetcprice").text
		 SL_Units  = SL_prod.item(x).selectSingleNode("units").text
		 SL_Drop  = SL_prod.item(x).selectSingleNode("dropship").text
		 SL_Contruct  = SL_prod.item(x).selectSingleNode("construct").text
		 SL_Shortdesc  = SL_prod.item(x).selectSingleNode("inetshortd").text
		 
		 
		 SL_GIFTCERT = SL_prod.item(x).selectSingleNode("giftcert").text
		 
		 SL_GiftCard       =SL_prod.item(x).selectSingleNode("giftcard").text
		 
		 SL_MinPrice       =SL_prod.item(x).selectSingleNode("minprice").text
		 SL_MaxPrice       =SL_prod.item(x).selectSingleNode("maxprice").text
		 
		 
		 if SL_GIFTCERT ="true"  then
		 	is_giftcert = true
		 else
		 	is_giftcert = false
		 end if
		
		 
		 target_in_SpecialPrice = "//gsp[number='"+cstr(SLProdNumber)+"']"
		 has_special_Price= false
		 has_Onespecial_Price=false
		 set SL_prodSpecial = objDoc.selectNodes(target_in_SpecialPrice) 
		 	if SL_prodSpecial.length > 0 then
			 	has_special_Price=true
				
				SL_SP_qty = SL_prodSpecial.item(0).selectSingleNode("qty").text
				SL_SP_ordertotal = SL_prodSpecial.item(0).selectSingleNode("total_dol").text
				SL_SP_qty = replace(SL_SP_qty,".00","")
				if SL_SP_qty= 1 and (cdbl(session("SL_BasketSubTotal")) >= cdbl(SL_SP_ordertotal) )then
					SL_SP_price = SL_prodSpecial.item(0).selectSingleNode("price").text
					SL_SP_discount = SL_prodSpecial.item(0).selectSingleNode("discount").text

					SL_SP_CostMeth = SL_prodSpecial.item(0).selectSingleNode("costmethod").text
					SL_SP_CostPlusDiscount = SL_prodSpecial.item(0).selectSingleNode("costplus").text
					SL_SP_StockReOrdPrice = SL_prodSpecial.item(0).selectSingleNode("reordprice").text
					SL_SP_UnCostprice = SL_prodSpecial.item(0).selectSingleNode("uncost").text
					
					if SL_SP_CostMeth="" then
						unitprice = SL_price
						unitprice2 = SL_SP_price
						if unitprice2>0 then
							unitprice = unitprice2
						end if
						percentoff = (SL_SP_discount/100.0)
						discountedUnitPrice = cdbl(unitprice * (1- percentoff))									
					end if
									
					if SL_SP_CostMeth="P" then
							discountedUnitPrice= cdbl(SL_SP_StockReOrdPrice*(1+SL_SP_CostPlusDiscount/100.0))
					end if
							
					if SL_SP_CostMeth="U" then										
							discountedUnitPrice= cdbl(SL_SP_UnCostprice*(1+SL_SP_CostPlusDiscount/100.0))									
					end if
						
					has_Onespecial_Price=true
				end if
			end if
		 set SL_prodSpecial = nothing
		 
		 'if SL_Pref_Ship<> "" then								
		'	SL_prefShip_Title	=(SL_prod.item(x).selectSingleNode("ca_title").text)
  		'end if
		
		
		
		if SL_thumbnail="" then
			SL_thumbnail = SL_Fullimage
		end if
		
		 
		 if len(SL_thumbnail) > 0 then
			 StrFileName = "images/"+ SL_thumbnail  
			 StrPhysicalPath = Server.MapPath(StrFileName)
			     set objFileName = CreateObject("Scripting.FileSystemObject")	
					 if objFileName.FileExists(StrPhysicalPath) then
			  			imagename=StrFileName
				    else
						imagename="images/noimage.gif"
  					end if 
				set objFileName = nothing
	    else
	            imagename="images/noimage.gif"
	    end if 
		
		Products_Per_Line=PRODUCTS_PER_GRID
		colcount=colcount+1
		
		if WANT_REWRITE = 1 then
			 SL_urltitle = SLProductTitle
			 SL_urltitle = url_cleanse(SL_urltitle)
			 prodlink = insecureurl  + SL_urltitle + "/productinfo/" + SLURLCodeNumber + "/" 
		else 
			 prodlink = insecureurl +  "prodinfo.asp?number=" + SLProdNumber
		end if
		
		usethis_rating_img = "images/clear.gif"
        if SHOW_PRODUCTREVIEWS=1 then
         SL_rating = SL_prod.item(x).selectSingleNode("rating").text
         SL_rating =replace(SL_rating,".00","")
         'response.Write(SL_rating)
         'SL_rating = replace(SL_rating,".00","")
            if cdbl(SL_rating) > 0 then              
                pos=InStr(SL_rating,".")                
                if pos > 0 then                    
                    decimal_part = mid(SL_rating,pos)
                    avg_rating = mid(SL_rating,1,pos-1)
                    if len(decimal_part) > 3 then 
                        decimal_part = mid(decimal_part,1,3)
                     end if
                    'Response.Write avg_rating & "," & decimal_part & ","
                    if decimal_part > 0.01 and decimal_part < 0.5 then 
                    
    					avg_rating=avg_rating&"-5"
					else
						avg_rating=avg_rating+1
					end if
					
					usethis_rating_img = "images/"&avg_rating&+"note.gif"      
			    else
			        avg_rating = SL_rating
			        usethis_rating_img = "images/"&avg_rating&+"note.gif"      			        
			        
                end if 'if pos > 0 then
            end if   
       end if 
	
	if colcount=PRODUCTS_PER_GRID+1 then
	colcount=1
		%>
	<tr>
	<%end if%>
	<td valign="top" width="140" >
		
		<table width="100%" cellpadding="0" cellspacing="0" border="0">
		<tr valign="top">
		<td valign="top" align="center" style="width:140px; padding-left: 15px;">
		
		 <table width="100%" cellpadding="2" cellspacing="0" border="0" style="height:130px;" class="table-layout-fixed">
			<tr><td align="left" valign="top"><a HREF="<%=prodlink%>"><img SRC="<%=imagename%>" alt="<%=SLProductTitle%>" title="<%=SLProductTitle%>" class="prodlistimg" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %> class="prodlistimg" align="top"></a></td></tr>
            <tr><td align="left"><img src="<%=usethis_rating_img%>" /></td></tr>						
			<tr valign="top"><td valign="top" align="left" style="height:35;"><a title="<%=SLProductTitle%>" class="producttitlelink" HREF="<%=prodlink%>"><%=SLProductTitle%></a></td></tr>	
			<tr><td valign="middle" align="left" class="plaintext">Item #: <%=SLProdNumber%></td></tr>
			
			<%if SL_Size_color="true"  then 
			
			    if SHOW_MINMAXPRICE=1 then
			        ShowRangePrice = false
    		        if cdbl(SL_MinPrice) <> cdbl(SL_MaxPrice) then
	    	            ShowRangePrice= true
		            end if

			%>
			
			    <tr>
			        <td valign="middle" align="left" class="ProductPrice">
			        <% if ShowRangePrice= true then %>
                            <%=formatcurrency(SL_MinPrice)%> - <%=formatcurrency(SL_MaxPrice)%>
                    <%else %>
                            <%=formatcurrency(SL_MinPrice)%>
                    <%end if %>			        
	                <%if has_special_Price=true then %>
	                    <img src="images/RedSaleTag.gif" border="0" align="absmiddle" hspace="6" />
	                <%end if %>
			        </td>
			    </tr>
			    
			    <%end if %>
			<%end if %>
			
			<%if SL_Size_color="false" then %>
			    <%if has_special_Price=true or has_Onespecial_Price=true then %>
			        <%if has_Onespecial_Price=true then %>
				        <tr><td align="left" class="plaintext">
			                <span class="ProductPrice"><s><%=formatcurrency(SL_price)%></s></span>
			                <br>
			                <table width="0" cellpadding="0" cellspacing="0" border="0">
			                <tr><td class="nopadding"><img src="images/RedSaleTag.gif" style="border:0; width:53; height:23;" alt="sale"></td>
				                <td class="plaintext nopadding" align="center" style="background-image:url(images/RedSaleTag_bkg.gif);height:23">
				                <span style="font-weight:bold;color:White;"><%=formatcurrency(discountedUnitPrice)%></span>&nbsp;
				                </td>
			                </tr>
			                </table>			
    					    
				        </td></tr>
			        <%else%>
			            <tr><td align="left" class="ProductPrice"><%=formatcurrency(SL_price)%></td></tr>
				        <tr><td valign="middle" align="left" style="padding-top:3px"><img src="images/RedSaleTag.gif" border="0" align="absmiddle"  /></td></tr>
			        <%end if%>
		        <%else%>
			        <tr><td valign="middle" align="left" class="ProductPrice"><%=formatcurrency(SL_MinPrice)%></td></tr>
		        <%end if%>
		    <%end if 'SL_GiftCard="false"%>
		    		   
			
		</table>
		
		</td></tr>
		
		<tr><td style="height:30px;"></td></tr>
		</table>
		
	</td>
	<%next%>
	
	<% diff = Products_Per_Line - colcount 
			if diff > 0 then
				for y=1 to diff 
				%>
				<td width="140">&nbsp;</td>
				<%next%>
		<%end if%>
		
	
	<%
	set SL_prod = nothing
	%>	
	
	</table>
	
	<%end if 'SHOW_PRODUCTSLISTVIEW %>
	
	<%end if 'objDoc.parseError.errorCode <> 0%>
	
	
	<table width="100%"  border="0" cellpadding="0" cellspacing="0">
	<tr><td ALIGN="right" id="pagelinks">
					<%if pcount>PRODUCTS_PER_PAGE and PRODUCTS_PER_PAGE>0 then 
									totpages=int(pcount/PRODUCTS_PER_PAGE) 
	    							totpages2=pcount/PRODUCTS_PER_PAGE
									if totpages<>totpages2 then
										totpages=totpages+1
									end if %>
								
									<!--start -->
									<%if int(showthispage) > 1 then %>
									<a HREF="<%=firstpagelink%>" class="previous">FIRST</a>&nbsp;&nbsp; <a href="searchprods.asp?searchstring=<%=server.urlencode(session("searchstring"))%>&amp;pagenumber=<%=showthispage-1%>&val=<%=session("val")%>" class="arrowleft"></a>
									<%end if %>
									<%
									numberofpages = 6 'just change this value.. to number of pages to be displayed on one page i.e 1 to 15 on one page, 16-30 on next page

									if int(showthispage)>numberofpages then%>
										<%
								 			modval = showthispage mod numberofpages
			  								cintval = int((showthispage-numberofpages)/numberofpages)
											if modval = 0 then
												cintval = cintval-1
											end if
										    'nextprev=int((showthispage-numberofpages)/numberofpages)*numberofpages+1
											nextprev=cintval*numberofpages+1
										%>
									
										<%	startpage=nextprev+numberofpages  'int(session("showthispage"))
										if int(showthispage)+ (numberofpages-1) > totpages then
											if modval = 0 then
												nextnext=startpage+ numberofpages
												limitpages=startpage+ (numberofpages-1)
											else
												limitpages=totpages
											end if 
																						
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
									end if
									
									%> <!--<b>Page</b>&nbsp;&nbsp;-->
									<%
									for productsy=startpage to limitpages%>
										<%if productsy=int(showthispage) then%>
											<%=productsy%>&nbsp;
										<%else%>
											<a HREF="searchprods.asp?searchstring=<%=server.urlencode(session("searchstring"))%>&amp;pagenumber=<%=productsy%>&val=<%=session("val")%>"><%=productsy%></a>&nbsp;
										<%end if%>
									<%next%>

									<%if int(showthispage) < totpages then %>
									    <%if limitpages < totpages then %>
									        ... <a href="<%=lastpagelink%>"><%=totpages%></a>&nbsp;
									    <% end if %>
    									<a HREF="searchprods.asp?searchstring=<%=server.urlencode(session("searchstring"))%>&amp;pagenumber=<%=showthispage+1%>&val=<%=session("val")%>" class="arrowright"></a>&nbsp;
    									<a href="<%=lastpagelink%>" class="next">LAST</a>&nbsp;
									<%end if %>
									
								<%end if%>
								
					</td></tr>				  
	</table>	

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
