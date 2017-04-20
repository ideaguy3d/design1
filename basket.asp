<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<% 
	session("destpage")="basket.asp" 
	
	if session("GoodforCookies") = false then			
		response.redirect("cookiecheck.asp")
	end if 
	
	
%>

<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<!-- #include file="google/googleglobal.asp" -->

<% 
Response.ExpiresAbsolute = #Feb 18,1998 13:26:26#
'Response.ExpiresAbsolute = 0

'	if MULTI_SHIP_TO=1 and session("user_want_multiship")= true and session("items_spaned") = false then
'		session("items_spaned") = true
'		xmlstring = sitelink.GETBASKETINFO_Group(session("ordernumber"))
'		objDoc.loadxml(xmlstring)
'		set SL_basket_group = objDoc.selectNodes("//basket_group")
'		
'		'now call basketinfo to get custominfo.
'		extrafield="basket.pricechang"
'		xmlstring = sitelink.getbasketinfo(session("shopperid"),session("ordernumber"),extrafield,false,false)
'		objDoc.loadxml(xmlstring)
'		
'		'now empty the cart
'		CALL sitelink.EMPTYBASKET(SESSION("SHOPPERID"),SESSION("ORDERNUMBER"))
'		
'		
'		 for x=0 to SL_basket_group.length-1
'			SL_Number	= SL_basket_group.item(x).selectSingleNode("item").text
'			SL_qty	= SL_basket_group.item(x).selectSingleNode("qty").text
'			
'			targetnode = "//gbi[item='"+(SL_Number)+"' and custominfo[string-length(.)> 0]]"
'			custominfo=""
'			slpricechange="false"
'			slunlistprice=0
'			set SL_Basketcinfo  = objDoc.selectNodes(targetnode) 
'			if SL_Basketcinfo.length > 0 then
'				custominfo=SL_Basketcinfo.item(0).selectSingleNode("custominfo").text
'				slpricechange = SL_Basketcinfo.item(0).selectSingleNode("pricechang").text
'				slunlistprice = SL_Basketcinfo.item(0).selectSingleNode("it_unlist").text
'			end if
'			set SL_Basketcinfo =nothing
'			
'			MyArray = Split(SL_qty, ".", -1, 1)
'				numeric_part = MyArray(0)
'				decimal_part = MyArray(1)
'				
'			 for y = 1 to cint(numeric_part)
'			 	'call item itemdd
'			 	if slpricechange="true" then
'			 	    call sitelink.dydasuppadd("",cstr(SL_Number),cstr("1"),cdbl(slunlistprice),session("shopperid"),session("ordernumber"),false,cint(0),cstr(custominfo))
'			 	else
'				    CALL sitelink.ITEMADD(cstr(SL_Number),cstr("1"),cstr(""),session("shopperid"),session("ordernumber"),cstr(custominfo),cint(0),cint(0))			 
'			    end if
'			 next 
'			 if decimal_part > 0 then
'			    if slpricechange="true" then
'			        call sitelink.dydasuppadd("",cstr(SL_Number),cstr(decimal_part),cdbl(slunlistprice),session("shopperid"),session("ordernumber"),false,cint(0),cstr(custominfo))
'			    else
'			 	    CALL sitelink.ITEMADD(cstr(SL_Number),cstr(decimal_part),cstr(""),session("shopperid"),session("ordernumber"),cstr(custominfo),cint(0),cint(0))			 
'			    end if
'			 end if 
'			
'		next 
'		set SL_basket_group = nothing 
'
'	end if
	
	
	call SITELINK.quantitypricing(session("shopperid"),session("ordernumber"))
	
	session("SL_BasketCount") = sitelink.GetData_Basket(cstr("BASKETCOUNT"),session("ordernumber"),session("SL_VatInc"))
	session("SL_BasketSubTotal") = sitelink.GetData_Basket(cstr("BASKETSUBTOTAL"),session("ordernumber"),session("SL_VatInc"))
	
	 basket_total = 0
	 

		if session("checkedforpoints") = false then
			session("ENABLE_POINTS")=true
			SL_ALLOWPOINTS 	=  sitelink.getdata("STARTPRK",session("ordernumber"))
			SL_POINT_REDEEEM_VAL= sitelink.getdata("STARTRTP",session("ordernumber"))
			
			if SL_ALLOWPOINTS=false or SL_POINT_REDEEEM_VAL=1 then
				session("ENABLE_POINTS")=false
			end if
		    session("checkedforpoints")= true
		end if
		
		if ALLOW_POINTS = 1 then
			if session("ENABLE_POINTS")=false then
				ALLOW_POINTS=0
			end if			
		end if


 %>


<html>
<head>
<title>Aussie Products.com | Shopping Cart | Australian Products Co.</title>
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

<!-- Latest compiled and minified CSS -->
<link rel="stylesheet" href="design/styles/julius-css.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
<link rel="stylesheet" href="design/jstyles.css">
<link rel="stylesheet" href="design/styles/jstyles.css">
<link rel="stylesheet" href="design/bower_components/font-awesome/css/font-awesome.css">

<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />

<script type="text/javascript">
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
</script>
<script type="text/javascript" language="JavaScript" src="calender.js"></script>
<script type="text/javascript" language="JavaScript" src="basket.js"></script>
<script type="text/javascript" language="javascript">
<!-- hide from JavaScript-challenged browsers
function displaythis(what)    
		{
		var bcheckout = document.getElementById("checkout");
		var bmodify   = document.getElementById("modifybasket");
		if (what.value =="checkout")
		    {
		        bcheckout.value = "checkout"
		        bmodify.value = ""		  
		    }
		if (what.value =="modify")
		    {
		        bcheckout.value = ""
		        bmodify.value = "modifybasket"	
		    }		 
		 }
	//-->hiding
	
function dosubmit()    
{
    var bmodify   = document.getElementById("modifybasket")	;    
    var bcheckout = document.getElementById("checkout");
    bcheckout.value = ""
    bmodify.value = "modifybasket";
    document.basketform.submit();	
}
</script>


</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav-modern.asp" -->
    <div id="main">

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<%if WANT_SIDENAV=1 then%>
        <td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
        <!--#INCLUDE FILE = "include/side_nav.asp" -->
        </td>	
		<%end if%>
		<td valign="top" class="pagenavbg">
<br>
<!-- sl code goes here -->
<div id="page-content">	
	
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%">Shopping Cart <%'=session("ordernumber")%></td></tr>
	</table>
	<br>		
	
	<%		
		extrafield ="shipcharge"
		xmlstring = sitelink.getbasketinfo(session("shopperid"),session("ordernumber"),extrafield,false,session("user_want_multiship"))
		
		objDoc.loadxml(xmlstring)
			
		basket_total = 0 
		
		GC_Redeem = false
		set SL_Basket_GC = objDoc.selectNodes("//gbi[certid>0]")
		    if SL_Basket_GC.length > 0 then
		        GC_Redeem=true
		    end if
		set SL_Basket_GC = nothing
		
		basket_has_fractions = false
		
		if WANT_GOOGLECHECKOUT=1 then
		    NO_fract_qty =0
		    tnode = "//gbi[quanto[contains(.,'.00')] and certid=0 and basketdesc!='PROMO']"
		    set SL_Basket_NoFrac =  objDoc.selectNodes(tnode)
		       for x=0 to SL_Basket_NoFrac.length-1
		         NO_fract_qty = NO_fract_qty + SL_Basket_NoFrac.item(x).selectSingleNode("quanto").text
		       next			    
		    set SL_Basket_NoFrac = nothing
    				    
    		
		    if cdbl(session("SL_BasketCount"))= cdbl(NO_fract_qty) then
			    basket_has_fractions = false
		    else
			    basket_has_fractions = true
		    end if	
		    
		    tnode= "//gbi[ship_via[string-length(.)> 0]]"
		    set SL_Basket_HasShipVia =  objDoc.selectNodes(tnode)
		    if SL_Basket_HasShipVia.length > 0 then
		        basket_has_fractions = true
		    end if		
		    set SL_Basket_HasShipVia=nothing    	
		end if
		
		set SL_Basket = objDoc.selectNodes("//gbi[certid=0]")
		
		points_left = session("pointavail")
		LineItempoints_redeemed=false
			
		if ALLOW_POINTS=1 then
				filter_condition = "//gbi[pts_rdeemd='true']"
								
				set SL_Basket_Points_item = objDoc.selectNodes(filter_condition)
		        if SL_Basket_Points_item.length > 0 then
		            LineItempoints_redeemed=true
					for x= 0 to SL_Basket_Points_item.length-1
		                SL_Points_ned       = SL_Basket_Points_item.item(x).selectSingleNode("points_ned").text				
					    SL_BasketQuantity	= SL_Basket_Points_item.item(x).selectSingleNode("quanto").text
					    SL_BasketQuantity   = replace(SL_BasketQuantity,".00","")					
					    points_redeemed	    = SL_Points_ned * SL_BasketQuantity					
					    points_left			= points_left - points_redeemed					
					next 
		        end if
		    	set SL_Basket_Points_item =nothing
			
		end if
		
		
		if session("user_want_multiship")=true then
			basketfilter = replace(CUSTOM_SHIPPING_FILTER,"shiplist","ca_code")
			'load all shipping method
			 xmlstring = sitelink.SHIPPINGMETHODS() 
			 objDoc.loadxml(xmlstring)
			 set SL_Carrier = objDoc.selectNodes("//shipmeths" + basketfilter)								
		 end if
			
		
		if SL_Basket.length = 0 then
	%>
		 <p align="center" class="plaintextbold">Your Cart is empty.<br><br>
		  <a HREF="default.asp"><img  src="images/btn-continueshopping.gif" border="0" alt="Continue Shopping"></a>
		  <br>
		  
		  <img alt="" src="images/clear.gif" width="1" height="160" border="0">
	 
		<%else%>
		
		<%
			'write cookie
	        if session("cookiewritten") = false then
		        'make a cookie for this order
		        cookiename=ShortStoreName+"order"
		        key_expires=sitelink.writecookie("order","expire","",0)
		        key_order=sitelink.writecookie("order","order",cstr(session("ordernumber")),1)
		        RESPONSE.COOKIES(cookiename).Expires=key_expires
		        RESPONSE.COOKIES(cookiename)("ordernumber")=(key_order)
		        session("cookiewritten")= true
          END IF
		'curr_skey =sitelink.GET_PROMOCODE(session("ordernumber"))		
		'if curr_skey = "" then
		 %>
		
		<table width="100%" border="0" cellpadding="4">
		<form id="checksourcekeyform" method="post" action="checksourcekey.asp">
		<tr><td width="100%">
			<span class="plaintextbold">
				<%if session("askedforsourcekey")=-1 or (session("askedforsourcekey")=1 and session("sourcekeyisgood")=-1) then%>					
            	      <strong>&nbsp;&nbsp;&nbsp;&nbsp;If you have a 
                	  key code, please enter it.</strong>
				  		&nbsp;
	                    <input type="text" name="discount" size="10" maxlength="9" value="<%=session("sourcekeyentered")%>">
    	                <input type="submit" name="Submit" value="Submit"><br>
    	         <%end if%>	
				<%if session("sourcekeyisgood")=1 then%>
					key code has been accepted.			
				<%end if%>
				<%if session("askedforsourcekey")=1 and session("sourcekeyisgood")=-1 then%>
        		    The key code you entered is not valid.&nbsp;
		            Please try again.
				<%end if%>
				</span>
		</td></tr>
		</form>		
		</table>
		

		
		<%'end if 'if len(curr_skey) %>				
		<form method="post" id="basketform" name="basketform" action="modifybasket.asp">
		<table width="100%"  cellpadding="2" cellspacing="0" border="0">
		
		<tr>
		<td>
		<% if session("user_want_multiship")= false then %>
		&nbsp;&nbsp;&nbsp;<a href="#" class="allpage btn btn-sm btn-info" onClick="MM_openBrWindow('PreviewShipping.asp','','width=450,height=400')">click here to estimate shipping cost </a><br>
		</td>
		<%end if%>
		
		</tr>
		</table>
		

		 
		<table width="100%"  border="0" cellspacing="0" cellpadding="3">
		<tr>
			<th align="CENTER" class="THHeader" width="6%">Remove</th>
			<th align="CENTER" class="THHeader" width="6%">Qty</th>
			<th align="CENTER" class="THHeader" width="15%">Description</th>
            <th align="center" class="THHeader" width="57%">&nbsp;</th>
			<th align="CENTER" class="THHeader" style="text-align:right;" width="16%">Total&nbsp;</th>
		</tr>
        <tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0"></td></tr>
			
			<% 
				lastitemadded = ""
			   for x=0 to SL_Basket.length-1
				SL_Number			= SL_Basket.item(x).selectSingleNode("number2").text
				SL_Variant			= SL_Basket.item(x).selectSingleNode("variant").text
				SL_BasketNumber 	= SL_Number +" "+ SL_Variant
				SL_BasketQauntity	= SL_Basket.item(x).selectSingleNode("quanto").text
				SL_BasketDesc		= SL_Basket.item(x).selectSingleNode("inetsdesc").text
				SL_BasketDesc2		= SL_Basket.item(x).selectSingleNode("desc2").text
				SL_BasketUnitPrice	= SL_Basket.item(x).selectSingleNode("it_unlist").text
				SL_BasketExtPrice	= SL_Basket.item(x).selectSingleNode("extendp").text
				SL_BasketCmpPrice	= SL_Basket.item(x).selectSingleNode("inetcprice").text
				SL_BasketDiscount	= SL_Basket.item(x).selectSingleNode("discount").text
				'SL_BasketUS			= SL_Basket.item(x).selectSingleNode("inetusprod").text				
				'SL_BasketUS_Msg		= SL_Basket.item(x).selectSingleNode("inetusmsg").text
				'SL_BasketCS			= SL_Basket.item(x).selectSingleNode("inetcsprod").text
				'SL_BasketCS_Msg		= SL_Basket.item(x).selectSingleNode("inetcsmsg").text								
				SL_BasketCustInfo	= SL_Basket.item(x).selectSingleNode("custominfo").text
				SL_Basket_Shipto	= SL_Basket.item(x).selectSingleNode("ship_to").text
				SL_inetthumb		=SL_Basket.item(x).selectSingleNode("inetthumb").text
				SL_Fullimage		=SL_Basket.item(x).selectSingleNode("inetimage").text
				SL_item_id   		=SL_Basket.item(x).selectSingleNode("record").text
				SL_prefShip   		=SL_Basket.item(x).selectSingleNode("prefship").text
				SL_ShipVia 			=SL_Basket.item(x).selectSingleNode("ship_via").text
				
				SL_Points_ned       = SL_Basket.item(x).selectSingleNode("points_ned").text				
				SL_PointsRdm		= SL_Basket.item(x).selectSingleNode("pts_rdeemd").text
				
				'SL_PointsRdm        ="false"
				'if SL_PointsRdm="true" then
				'	points_left = points_left - SL_Points_ned
				'end if
				
				SL_ShipWhen 		=SL_Basket.item(x).selectSingleNode("ship_when").text
				
				SL_units 			=SL_Basket.item(x).selectSingleNode("units").text
				SL_drop          	=SL_Basket.item(x).selectSingleNode("dropship").text
				SL_construct 	    =SL_Basket.item(x).selectSingleNode("construct").text
				
				SL_certid           =0
				SL_certid           =SL_Basket.item(x).selectSingleNode("certid").text
				SL_hasAttribs       =SL_Basket.item(x).selectSingleNode("attribs").text
				
				SL_GiftCard       =SL_Basket.item(x).selectSingleNode("giftcard").text
			    SL_ProdShipCharge = trim(SL_Basket.item(x).selectSingleNode("shipcharge").text)
				
				IsGiftCard = false
				
				if SL_GiftCard="true" then
				    IsGiftCard=true
				end if
				
				has_attrib          = false
				
				if SL_hasAttribs="true" then
				    has_attrib=true
				end if
				
				 SL_GIFTCERT = SL_Basket.item(x).selectSingleNode("giftcert").text
				 SL_Promo_Ind  =SL_Basket.item(x).selectSingleNode("basketdesc").text
				 SL_PROMO_ITEM= false
				 if SL_Promo_Ind ="PROMO"  or SL_certid > 0 then
				 	SL_PROMO_ITEM = true
				 end if
		 
				 if SL_GIFTCERT ="true" then
		 			is_giftcert = true
					SL_BasketCustInfo= sitelink.FIX_STRING(server.urlencode(SL_BasketCustInfo))
					SL_BasketCustInfo = replace(SL_BasketCustInfo,"Gift","")
		 		else
				 	is_giftcert = false
				end if
				
				if SL_PROMO_ITEM=false then
					lastitemadded = SL_Number
				end if
				
				if len(trim(SL_ShipWhen)) > 0 then
					SL_ShipWhen  = replace(SL_ShipWhen,"T"," ")
					SL_ShipWhen_Date = FormatDateTime(SL_ShipWhen, 0)
				else
					SL_ShipWhen_Date = "" 
				end if 
				
				if len(trim(SL_prefShip)) > 0 then								
					SL_prefShip_Title	=SL_Basket.item(x).selectSingleNode("ca_title").text
				end if
									
				
				SL_BasketQty = replace(SL_BasketQauntity,".00","")
				
				SL_Points_ned   = SL_Points_ned * SL_BasketQty
								
				basket_total = basket_total + SL_BasketExtPrice

				if (x mod 2) = 0 then
					class_to_use = "tdRow1Color"
				else
					class_to_use = "tdRow2Color"
				end if 
				
				if SL_inetthumb = "" then
					SL_inetthumb = SL_Fullimage
				end if

				if SL_Number="" then
					SL_inetthumb = "clear.gif"
				end if 
				StrFileName = "images/"+cstr(SL_inetthumb)
				StrPhysicalPath = Server.MapPath(StrFileName)
				set objFileName = CreateObject("Scripting.FileSystemObject")	
				  if objFileName.FileExists(StrPhysicalPath) then
			       		inethumbImage=StrFileName
				  else
						inethumbImage="images/noimage.gif"
	  			  end if
				  set objFileName = nothing
				altag =trim(SL_BasketDesc)
				
				
								
			%>
			<tr class="<%=class_to_use%>">
			<td valign="middle" align="center">			
					<%
						remname  ="remove" & cstr(x+1)
						applypoints = "Points"& cstr(x+1)
					
					%>					
					<% if SL_PROMO_ITEM=true  then %>					
						
					<%else%>
						<input type="checkbox" name="<%=remname%>" value="0">
					<%end if%>
					<input type="hidden" name="item_id" value="<%=SL_item_id%>">
					
			</td>
			
			<td align="center">			
			<% if SL_PROMO_ITEM=true or is_giftcert=true or SL_PointsRdm="true" or IsGiftCard=true then %>
				<input type="hidden" name="qty" value="<%=SL_BasketQty%>">
				<%if is_giftcert=true or IsGiftCard=true then %>
				<span class="plaintext">				
				<%=SL_BasketQty%></span>
				<%end if%>
			<%else%>
				<input type="text" name="qty" size="2" class="plaintext" maxlength="5" value="<%=SL_BasketQty%>">
			<%end if%>
			<input type="hidden" name="prevqty" value="<%=SL_BasketQty%>">
			</td>
			<td valign="top" width="120"><div class="prodthumb"><div class="prodthumbcell"><img src="<%=inethumbImage%>" alt="<%=altag%>" class="prodlistimg" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %> border="0"></div></div></td>
			<td valign="top" class="plaintextbold">
				<%=SL_BasketDesc%>
				<%
				if SHOW_IN_STOCK = 1 and SHOW_DUE_DATE = 0  then
					if  SL_units >0 or  SL_drop="true" or SL_construct= "true" then
						Response.Write("&nbsp;&nbsp;(In Stock)")
					else
						Response.Write("&nbsp;&nbsp;(Out of Stock)")
					end if
				end if 
				
				%>			

				<% if len(trim(SL_BasketDesc2)) > 0 then response.write("<br>" +SL_BasketDesc2) end if %>
				<br><br>
				<strong>Item: </strong><%=SL_BasketNumber%> &nbsp; <br><strong>Price: </strong><%=formatcurrency(SL_BasketUnitPrice)%> &nbsp;
				<% if SHOW_PRODSHIPCHRG=1 then 
				    if len(SL_ProdShipCharge)>0 and IsNumeric(SL_ProdShipCharge)=true and SL_ProdShipCharge>0 then %>
				    <br>Product specific shipping charge: <%=formatcurrency(SL_ProdShipCharge)%>
				    <%end if%>
				<%end if%>
				
				<%if SL_PointsRdm="true" then%>
					<br><%=SL_Points_ned%> points were redeemed to purchase this product
				<%end if%>
				
				<%if SHOW_COMP_PRICE="TRUE" and SL_BasketCmpPrice > 0 then  response.write("<br><strong>Compare at :</strong>"+ formatcurrency(SL_BasketCmpPrice)) end if %>
				<%if SL_BasketDiscount > 0 then   response.write("<strong>&nbsp;&nbsp;Discount: </strong>"+ (SL_BasketDiscount)+"% ") end if %>
				<%if len(trim(SL_BasketCustInfo)) > 0  then %>
				    <br>
				    <%if is_giftcert=false and SL_PROMO_ITEM=false and has_attrib=false then %>
					    <a class="sidenav2" href="editcustomInfo.asp?item_id=<%=SL_item_id%>&amp;quant=<%=SL_BasketQty%>&amp;item=<%=SL_Number%>&amp;variant=<%=SL_Variant%>&amp;basketrecord=<%=x+1%>"><%=SL_BasketCustInfo%></a>&nbsp;<a class="sidenav2" href="editcustomInfo.asp?item_id=<%=SL_item_id%>&amp;quant=<%=SL_BasketQty%>&amp;item=<%=SL_Number%>&amp;variant=<%=SL_Variant%>&amp;basketrecord=<%=x+1%>">(Edit)</a>
				    <%else%>
					    <%=SL_BasketCustInfo%>
				    <%end if%>
				<%end if%>
				
				
				<%
				'start ship to multiple address code 
				if SL_PROMO_ITEM=false then
				    if MULTI_SHIP_TO=1 and session("user_want_multiship")= true and session("registeredshopper")="YES" then
				%>
				    <table width="100%" border="0" cellpadding="1" cellspacing="0">
                    <tr><td height="9"></td></tr>
                    <tr><td class="plaintext">
				            <strong>Ship To: </strong>&nbsp;
				        </td>
				        <td>
				            <%'=SL_Basket_Shipto%>
				            <%if SL_Basket_Shipto > 0 then
					                hasgiftmsg=true
						            usethisship = SL_Basket_Shipto + 1 - 1
            												
						             fname = SL_Basket.item(x).selectSingleNode("firstname").text
						             lname = SL_Basket.item(x).selectSingleNode("lastname").text
            												
				            end if
				            %>									
            				
				            <select name="shipdrop" class="smalltextblk" onChange="RedirectToAddADDress(this.options[this.selectedIndex].value)">
            				
				            <option value="changeshipto.asp?basketrecord=<%=x+1%>&amp;shipdrop=0" <%if SL_Basket_Shipto= 0 then response.write(" selected") end if%>>Same As Bill to
				            <option value="AddToAddressBook.asp?basketrecord=<%=x+1%>">Someone else
				            <% if SL_Basket_Shipto > 0 then %>
					            <option value selected><%=lname%>&nbsp;<%=fname%>
				            <%end if%>
        				
					        </select>
					        &nbsp;
					        <% 'if SL_Basket_Shipto > 0 and hasgiftmsg= false then %>
					        <% if SL_Basket_Shipto > 0 then %>
					        <a class="sidenav2" href="EditGiftMessage.asp?basketrecord=<%=x+1%>&amp;shopcust=<%=SL_Basket_Shipto%>&amp;numberpart=<%=SL_Number%>&amp;variation=<%=SL_Variant%>">Enter Gift Message(optional)</a>
					        <%end if%>
				        </td> 
                    </tr>
                    <tr><td height="3"></td></tr>
                    <tr><td class="plaintext">
                	    <b>Ship Via: </b>&nbsp;
                        </td>
                        <td>
				        <% if SL_prefShip<>"" then %>
					        <%=SL_prefShip_Title%>
					        <input type="hidden" name="Ship_via" value="<%=SL_prefShip%>">
				        <%else%>
					        <% if session("registeredshopper")="YES" and MULTI_SHIP_TO= 1 then %>
					        <select name="Ship_via" class="smalltextblk" onChange="ChangeCarrier(this.options[this.selectedIndex].value)">
					         <option value="changecarrier.asp?item_id=<%=SL_item_id%>&amp;shipmethod=">Select Shipping Method					 
        					 
					        <%
						        for z=0 to SL_Carrier.length-1
							        SL_ShippingCode = SL_Carrier.item(z).selectSingleNode("ca_code").text
							        SL_ShippingTitle = SL_Carrier.item(z).selectSingleNode("ca_title").text					
					        %>
						        <option value="changecarrier.asp?item_id=<%=SL_item_id%>&amp;shipmethod=<%=SL_ShippingCode%>" <% if SL_ShipVia = SL_ShippingCode then response.write(" selected") end if %>><%=SL_ShippingTitle%>
						        <%next%>					
					        </select>
					        <%end if  'if SL_prefShip<>"" %>

				        <%end if%>
					    </td>
                    </tr>
                    <tr><td height="3"></td></tr>
                    <tr><td class="plaintext">
                        <b>Ship When: </b>
                        </td>
                        <td><input type="text" name="shipwhen<%=(x+1)%>" value="<%=SL_ShipWhen_Date%>" maxlength="10" size="12" class="smalltextblk">
					        &nbsp;<a href="javascript:show_calendar('basketform.shipwhen<%=(x+1)%>','<%=SL_item_id%>');"><img src="images/calendarSm.gif" border="0" align="center" WIDTH="25" HEIGHT="19"></a>
                        </td>
                    </tr>
                    <tr><td height="6"></td></tr>
                    </table>
                    <%else%>
					    <% if SL_prefShip<>"" then %>
                            <br><b>Ship Via: </b>&nbsp;<%=SL_prefShip_Title%>
                            <input type="hidden" name="Ship_via" value="<%=SL_prefShip%>">
                        <%end if%>
				<%  end if
				end if 
				%>
			</td>
						
			
			<td ALIGN="right" valign="top"><span class="plaintext"><%=formatcurrency(SL_BasketExtPrice)%>&nbsp;&nbsp;
			<% if ALLOW_POINTS=1 then %>
			<% if session("registeredshopper")<>"YES" then %>
			<br>
			 <a href="<%=secureurl%>login.asp">Apply points</a>
			 <%end if%>
			 
			 <% if session("registeredshopper")="YES" and points_left < 0 and SL_PointsRdm="false" then %>
			 	<br>
				You do not have enough points
			 <%end if%>
			 
			<% 			
			if (points_left => SL_Points_ned) and (SL_Points_ned) > 0  then %>
				<%if SL_PointsRdm="false" and points_left > 0  then%>
					<br>
					<input type="checkbox" name="<%=applypoints%>" onClick="Javascript:dosubmit()" value="0">Apply <%=SL_Points_ned%> Points				
				<%end if%>
				
			<%else%>
				<%if SL_PointsRdm="false" and session("registeredshopper")="YES" and SL_Points_ned > 0 then%>
				<br>
				You do not have enough points
				<%end if%>
			<%end if%>
			<%end if 'ALLOW_POINTS=1 %>
			
			</span>
			</td>
			</tr>

			<%next%>
				<tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0"></td></tr>
                <tr>
					<th align="CENTER" class="THHeader" colspan="5">&nbsp;</th>
				</tr>
				<tr>
				<td colspan="4" class="plaintext">
				<input type="hidden" value="modify">&nbsp;
					<a class="allpage btn btn-sm btn-primary" href="Javascript:dosubmit()" value="modify" onClick="Javascript:dosubmit()">Update Cart</a>
					&nbsp;|&nbsp;&nbsp;<a href="emptybasket.asp" class="allpage btn btn-sm btn-primary">Empty Cart</a>
				
				</td>
				<td ALIGN="RIGHT" class="plaintextbold">Subtotal:&nbsp;<%=formatcurrency(cdbl(basket_total))%>&nbsp;&nbsp;</td>
			</tr>		
		</table>
		<div class="button-group">
        	<div class="button">
				<% 'if session("viewpage")="" then
					'destpage="default.asp"
					'else
					'	destpage=session("viewpage")
					'end if
					destpage="Javascript:history.go(-1)"
				%>
				<a href="<%=destpage%>"><img SRC="images/btn-continueshopping.gif" style="border:0" ALT="Continue Shopping"></a>
            </div>
            <div class="button">
            	<% 'if session("registeredshopper")="NO" and MULTI_SHIP_TO=1 then
					  if MULTI_SHIP_TO=1 and session("user_want_multiship")= false  then
				%>
                <a href="multi_ship.asp"><img SRC="images/btn-shiptomultiple.gif" style="border:0;"   ALT="Multiple Ship to"></a>
                <%end if%>
            </div>
            <div class="button">
            	<input type="image" onClick="displaythis(this)" value="checkout" SRC="images/btn-proceedtocheckout.gif" BORDER="0" ALT="Purchase Now" />
            </div>      
        </div>
			<input type="hidden" name = "checkout" id="checkout" value=""/>&nbsp;<input type="hidden" name = "modifybasket" id="modifybasket" value = ""/>
				
		</form>
		
		<%		
		    if session("user_want_multiship")=true or LineItempoints_redeemed=true or GC_Redeem=true or session("OrderLevelPoints_Amt") > 0 or basket_has_fractions=true then
		        WANT_GOOGLECHECKOUT=0
		    end if		   
		%>
		
		<%if WANT_GOOGLECHECKOUT=1 or SHOW_BUYSAFE=1 then %>
		<div style="padding-left:350px;">
		<table width="0" border="0" cellspacing="0" cellpadding="3">
		    <tr><td colspan="2" style="height:20px;"></td></tr>			
             <tr>
             <%if SHOW_BUYSAFE=1 then %>
                <td align="right"><span id="buySAFE_Kicker" name="buySAFE_Kicker" type="Kicker Guaranteed Ribbon 200x90"></span></td>
                <td style="width:20px;"></td>
             <%end if %>
             
             <td align="right">
            <%
            if WANT_GOOGLECHECKOUT=1 then
                Dim ButtonSize, Enabled, Analytics

                ' Display a large, enabled Google Checkout button without analytics code.
                ButtonSize = "MEDIUM"
                Enabled = True
                Analytics = False
                DisplayButton ButtonSize, Enabled, Analytics
            end if 
            %>
            </td></tr>
			</table>
			</div>
	    <%end if%>
			
		<%end if%>
		
		
	<% 
	set SL_Basket = nothing
	set SL_Carrier=nothing	
	%>

	<%	if lastitemadded<>"" then 
		HOW_MANY_SITECROSS = 3
		xmlstring = sitelink.Get_From_SiteCross(lastitemadded,"",HOW_MANY_SITECROSS)
		objDoc.loadxml(xmlstring)
		set SL_SiteCross = objDoc.selectNodes("//sitecross") 
		if SL_SiteCross.length > 0 then
	%>
	<br><br>
		<table width="100%"  border="0" cellpadding="3" cellspacing="0">
			<tr>
            <td class="THHeader">&nbsp;Customers also bought these items</td>
          </tr>			
		</table>
		<br>
		<table width="100%"  border="0" cellspacing="0" cellpadding="0">

		  <% 

			rowcount=0  
			CSPROD_PER_LINE = 3			    					
			for x=0 to SL_SiteCross.length-1
				SLsellNumber = SL_SiteCross.item(x).selectSingleNode("sellwhatsku").text
				SLURLCodeNumber = server.urlencode(SLsellNumber)
				SLsellPrice  = SL_SiteCross.item(x).selectSingleNode("price1").text
				SLsellTitle  = SL_SiteCross.item(x).selectSingleNode("inetsdesc").text
				SLsellThumb  = SL_SiteCross.item(x).selectSingleNode("inetthumb").text
				SLsellFull   = SL_SiteCross.item(x).selectSingleNode("inetimage").text
				
				if WANT_REWRITE = 1 then
					 SL_urltitle = SLsellTitle
					 SL_urltitle = url_cleanse(SL_urltitle)					  
					 prodlink = insecureurl + SL_urltitle + "/productinfo/" + SLURLCodeNumber + "/"
				else 
					 prodlink = "prodinfo.asp?number=" + cstr(SLsellNumber)
				end if
				
						

			if rowcount=CSPROD_PER_LINE  then
			rowcount=0
			%>
			<tr><td></td></tr>		     					
			<%end if%>
			
			<%if rowcount=0 then %>
		        <tr>
		    <%end if %>
		   
		   <td valign="top" width="<%=(100/CSPROD_PER_LINE)%>%"> 				
		   <table  width="100%" border="0" cellspacing="0" cellpadding="0" class="table-layout-fixed">
			<tr>      
									
				 <% rowcount=rowcount+1
					 if SLsellThumb <>"" then
						StrFileName = "images/"+cstr(SLsellThumb)
					else
						StrFileName = "images/"+cstr(SLsellFull)
					end if
									 
					 found = false 
					 StrPhysicalPath = Server.MapPath(StrFileName)
					 set objFileName = CreateObject("Scripting.FileSystemObject")	
					 if objFileName.FileExists(StrPhysicalPath) then
						imagename=StrFileName
					else
						imagename="images/noimage.gif"
					end if 
					set objFileName = nothing
					
					%>
				   
					<td align="center"><a href="<%=prodlink%>">
					<img SRC="<%=imagename%>" alt="<%=SLsellTitle%>" title="<%=SLsellTitle%>"
					   class="prodlistimg" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %> border="0"></a></td>
					</tr>
					<tr><td align="center">
					<a class="allpage" href="<%=prodlink%>"><%=SLsellTitle%></a>
					</td></tr>
					<tr><td align="center" class="plaintext">
					<%=formatcurrency(SLsellPrice)%>
					</td></tr>
					<tr><td>&nbsp;</td></tr>
		</table>
		</td>											
		<%next%>
		<% diff = CSPROD_PER_LINE - rowcount 
			if diff > 0 then
				for y=1 to diff 
				%>
				<td>&nbsp;</td>
				<%next%>
		<%end if%>
		
		
		</tr>
				  
	</table>
	
	
	<%
	end if 'if SL_SiteCross.length > 0
	set SL_SiteCross=nothing
	end if%>
	
	
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
