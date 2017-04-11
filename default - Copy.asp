<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%
xmlstring =sitelink.READ_METATAG(cstr("0"),cstr(""))
objDoc.loadxml(xmlstring)
set SL_Dept_Meta = objDoc.selectNodes("//results")
if SL_Dept_Meta.length > 0 then
	metatagtitle = SL_Dept_Meta.item(0).selectSingleNode("ptitle").text
	metatagdesc = SL_Dept_Meta.item(0).selectSingleNode("pdesc").text
	metatagkeywrd = SL_Dept_Meta.item(0).selectSingleNode("pkeywrd").text

	usethismeta="<title>"+ cstr(metatagtitle) +"</title>"
	usethismeta =usethismeta + VbCrLf
	usethismeta =usethismeta + "<META NAME=""KEYWORDS"" CONTENT="""+ cstr(metatagkeywrd) + """>"
	usethismeta =usethismeta + VbCrLf
	usethismeta =usethismeta + "<META NAME=""DESCRIPTION"" CONTENT="""+ cstr(metatagdesc) + """>"
else
	usethismeta = "<title>"+ cstr(althomepage) +"</title>"
end if
set SL_Dept_Meta = nothing

session("department")=0
%>



<html>
<title>Aussie Products.com | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
	<meta http-equiv="Content-Type" content="application/xhtml+xml; charset=utf-8" />
	<meta name="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls" />
	<meta name="keywords" content="Aussie Products, Australian, Aussie, Australian Products, Australian Food, Aussie Foods, Australian Desserts, tim tams, vegemite, violet crumble, cherry ripe, koala, kangaroo, down under, Australian Groceries, foods, lollies, candy, gourmet food" />
	<meta name="robots" content="index, follow"/>
    <meta name="robots" content="ALL">
  
<html itemscope itemtype="http://schema.org/LocalBusiness">



<meta itemprop="name" content="Australian Products Co.">
<meta itemprop="image" content="http://www.aussieproducts.com/images/Green_Gold_Banner.jpg">

<meta itemprop="name" content="Aussie Products.com | Australian Products Co. Bringing Australia to You - Aussie Foods">
<meta itemprop="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls">
<meta itemprop="image" content="http://www.aussieproducts.com/">
	<meta name="distribution" content="Global">
	<meta name="revisit-after" content="5">
	<meta name="classification" content="Food">
    <meta name="google-site-verification" content="K2A2-bw3DiKf9q0_Ul-o-hJLTo3YkFNb_JHtsYJ0LJ0">
	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
<link rel="canonical" href="<%=insecureurl%>">
<link rel="stylesheet" href="text/topnav.css" type="text/css">

<link rel="stylesheet" href="text/storestyle.css" type="text/css">
<link rel="stylesheet" href="text/footernav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/design.css" type="text/css">

<!--<link rel="stylesheet" href="design/bootstrap/bootstrap.min.css">-->
<!-- Latest compiled and minified CSS -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
<link rel="stylesheet" href="design/jstyles.css">
<!-- this file is still on the server for now -->
<!--<link rel="stylesheet" href="design/styles/jstyles.css">-->

<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->

<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">
        <%
            session("destpage")=""
            session("viewpage") = session("destpage")
        %>
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
                <!-- Side Navigation: -->
                 <td width="210" class="sidenavbg" valign="top">
                    <!--#INCLUDE FILE = "include/side_nav.asp" -->
                 </td>
                 <td valign="top" class="pagenavbg">
                 <!-- sl code goes here -->
                 <div id="page-content" class="plaintext">
                 <br>
                 <!--#INCLUDE FILE = "text/HomePage5.htm" -->
                 <%
                 if FEATURED_DPT > 0 then

                  'response.write(sortorderby)
                     extrafield=""
                     set SLProdObj = New cProductObj
                     SLProdObj.ldeptcode    		= cstr(FEATURED_DPT)
                     SLProdObj.lSORTBY		 	= cstr("")
                     SLProdObj.lProd_Per_Page 	= TOTAL_PRODUCTS
                     SLProdObj.llnpage		 	= 1
                     SLProdObj.lsortorder	 	= cstr("")
                     SLProdObj.llextrafield	 	= extrafield
                     xmlstring =sitelink.PRODUCTSINDEPT(SLProdObj)
                     pcount=SLProdObj.lpcount
                     set SLProdObj = nothing

                     objDoc.loadxml(xmlstring)
                     set SL_prod = objDoc.selectNodes("//prodindept")

                     if SL_prod.length > 0 then
                         TotalProds = SL_prod.length
                         if TOTAL_PRODUCTS > 0 and TOTAL_PRODUCTS <= 50 and TotalProds > TOTAL_PRODUCTS then
                             TotalProds = TOTAL_PRODUCTS
                         end if

                         'get special pricing.
                         all_prods=""
                         for x=0 to TotalProds-1
                             SLProdNumber = SL_prod.item(x).selectSingleNode("number").text
                             all_prods =all_prods+","+cstr(SLProdNumber)
                         next
                         all_prods =all_prods+","

                         xmlstring = sitelink.GET_ALL_SPECIALPRICES(cstr(all_prods),session("ordernumber"),session("shopperid"),cstr(""))

                         objDoc.loadxml(xmlstring)
                 %>

                     <br /><br /><h1>Featured Products</h1>

                     <% if FEA_SHOW_PRODUCTSLISTVIEW=1 then %>
                     <table width="100%"  border="0" cellspacing="1" cellpadding="3" class="list">
                     <tr>
                         <th align="CENTER" class="THHeader" width="15%"><span class="THHeader">Item #</span></th>
                         <th align="CENTER" class="THHeader" width="10%">Image</th>
                         <th align="CENTER" class="THHeader" width="60%"><span class="THHeader">Description</span></th>
                         <th align="CENTER" class="THHeader" width="15%"><span class="THHeader">Price</span></th>
                     </tr>
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
                         <td class="<%=class_to_use%>" valign="middle" align="center"><div class="prodthumb"><div class="prodthumbcell"><a href="<%=prodlink%>">
                         <img SRC="<%=imagename%>" alt="<%=SLProductTitle%>" title="<%=SLProductTitle%>" class="prodlistimg" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %> border="0"></a></div></div></td>
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
                                     <a class="allpage" HREF="<%=prodlink%>"><u><%=formatcurrency(SL_MinPrice)%> - <%=formatcurrency(SL_MaxPrice)%></u></a>
                                 <%else %>
                                     <a class="allpage" HREF="<%=prodlink%>"><u><%=formatcurrency(SL_MinPrice)%></u></a>
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

                     <%else 'FEA_SHOW_PRODUCTSLISTVIEW %>

                     <table width="100%" border="0" cellspacing="1" cellpadding="3" class="grid">
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

                         Products_Per_Line=FEA_PRODUCTS_PER_GRID
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


                     if colcount=FEA_PRODUCTS_PER_GRID+1 then
                     colcount=1
                         %>
                     <tr>
                     <%end if%>
                     <td valign="top" width="140">

                         <table width="100%" cellpadding="0" cellspacing="0" border="0">
                         <tr valign="top">
                         <td valign="top" align="center" style="width:140px; padding-left: 15px;">

                          <table width="100%" cellpadding="2" cellspacing="0" border="0" style="height:130px;" class="table-layout-fixed">
                             <tr><td align="left" valign="top"><a HREF="<%=prodlink%>"><img SRC="<%=imagename%>" alt="<%=SLProductTitle%>" class="prodlistimg" title="<%=SLProductTitle%>" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %> class="prodlistimg" align="top"></a></td></tr>
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
                         </table>

                         <%end if 'FEA_SHOW_PRODUCTSLISTVIEW %>
                     <%end if 'if SL_prod.length > 0%>


                     <% set SL_prod = nothing %>
                     <%end if 'if FEATURED_DPT > 0 %>

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
<script type="text/javascript">
  (function() {
    window._pa = window._pa || {};
    // _pa.orderId = "myOrderId"; // OPTIONAL: attach unique conversion identifier to conversions
    // _pa.revenue = "19.99"; // OPTIONAL: attach dynamic purchase values to conversions
    // _pa.productId = "myProductId"; // OPTIONAL: Include product ID for use with dynamic ads
    var pa = document.createElement('script'); pa.type = 'text/javascript'; pa.async = true;
    pa.src = ('https:' == document.location.protocol ? 'https:' : 'http:') + "//tag.marinsm.com/serve/55c94a54db45c07a90000003.js";
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(pa, s);
  })();
</script>
</body>
</html>
