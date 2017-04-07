<%on error resume next%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 


<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->



<% 

    current_stocknumber = cstr(REQUEST.QUERYSTRING("number"))    
    current_stocknumber =fix_xss_Chars(current_stocknumber)
    
    'current_variation = cstr(REQUEST.QUERYSTRING("variation"))
    'current_variation =fix_xss_Chars(current_variation)
    current_variation=session("Current_selectedVariant")
    
    'youselected = request.QueryString("opt")
    
    'response.Write("===youselected -->"&youselected)
    
    
	isvalidstock = sitelink.validstocknumber2(current_stocknumber)
	
	if isvalidstock=-1 then
		set sitelink=nothing
		set ObjDoc=nothing
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl&"default.asp"
		Response.End
	end if
	
	IsValidUrl = true
    CorrectProductUrl = ""
    
    
    UrlWriteStr = request.ServerVariables("http_x_rewrite_url")
    IsValidUrl = sitelink.ValidateProdInfoUrl(current_stocknumber,UrlWriteStr,CorrectProductUrl)        
    
        
    
    if IsValidUrl=false then    
        set sitelink=nothing
		set ObjDoc=nothing
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl& CorrectProductUrl
		Response.End
    end if
    
	
	attribstring = request.QueryString("opt")  
	attribstring =fix_xss_Chars(attribstring)	
	attribstring = replace(attribstring,"/","")	
    if len(trim(attribstring)) > 0 then
		selected_variant = sitelink.GET_VARIANT_FROM_ATTRIBUTE(attribstring)
		if len(trim(selected_variant)) > 0 then
			current_variation=selected_variant	
		end if	
	end if
	
  
	SET LOSTOCK =  sitelink.setupproduct(current_stocknumber,cstr(current_variation),0,session("ordernumber"),cstr("")) 
	
	'for url rewrite
	if len(trim(LOSTOCK.inetsdesc)) = 0 then 
		usethis_for_urlwrite = URL_CLEANSE(trim(LOSTOCK.NUMBER))
	else
		usethis_for_urlwrite = URL_CLEANSE(trim(LOSTOCK.inetsdesc))
	end if
	
	'http://www.site.com/Product-Title/productinfo/AAA/?additem=&lt;script>%20blah%20blah%20blah</script>&maxoption=1
	if WANT_REWRITE = 1 and instr(request.ServerVariables("http_x_rewrite_url"),"?") > 0 then 
        splitlink = split(request.ServerVariables("http_x_rewrite_url"),"?")

        prodlink = insecureurl  + usethis_for_urlwrite + "/productinfo/" + current_stocknumber + "/" 
        if left(insecureurl,len(insecureurl)-1) & splitlink(0) <> prodlink then 
            response.Status = "301 Moved Permanently"
            response.AddHeader "location",prodlink
        end if 
	
	end if
	
	'if len(trim(LOSTOCK.number))= 0 then
	'SET LOSTOCK =  sitelink.setupproduct(cstr(REQUEST.QUERYSTRING("number")),cstr(""),0,session("ordernumber"),cstr(""))
	'end if
  
	bookmarktitle = replace(trim(LOSTOCK.inetsdesc),"'","")
	
	if len(trim(LOSTOCK.META_TITLE)) = 0 then
	    bookmarktitle=bookmarktitle
	else
	    bookmarktitle= trim(LOSTOCK.META_TITLE)
	end if 
    
	has_stock_attribs = LOSTOCK.attribs
	
	if len(trim(session("department")))=0 or session("department")=0 then
    	session("department") = sitelink.GET_DEPT_FROM_PROD(current_stocknumber)
    end if
	
	'session("department") = sitelink.GET_DEPT_FROM_PROD(cstr(current_stocknumber))
	'response.Write(session("department") )
	
	froogle_desc = trim(LOSTOCK.froogle)
	froogle_desc = replace(froogle_desc,"""","")
	froogle_desc = replace(froogle_desc,"<","")
	froogle_desc = replace(froogle_desc,">","")
	
	keylist = trim(LOSTOCK.inetkeywrd)
	keylist = replace(keylist,"""","")
	keylist = replace(keylist,"<","")
	keylist = replace(keylist,">","")
	
	set SL_ProdInfoDetails = New cProductInfoDetailsObj
	
	SL_ProdInfoDetails.lnumber=cstr(current_stocknumber)
	SL_ProdInfoDetails.lVariant=cstr(current_variation)
	SL_ProdInfoDetails.lOrderNumber=session("ordernumber")
	SL_ProdInfoDetails.lShopperID=session("shopperid")
	SL_ProdInfoDetails.lShowExpectedDate=SHOW_DUE_DATE
	
	call sitelink.GET_SLSTOCK_DETAILS(SL_ProdInfoDetails)
	
	SL_SizeColorXmlStream = SL_ProdInfoDetails.lSizeColorXmlStream
	SL_SpecialPriceXmlStream = SL_ProdInfoDetails.lSpecialPriceXmlStream
	exp_date = SL_ProdInfoDetails.lExpectedDate
	
	SL_CrossSellXmlStream = SL_ProdInfoDetails.lCrossSellXmlStream
	optcodexmlstring = SL_ProdInfoDetails.lst_optcode
	optvalxmlstring = SL_ProdInfoDetails.lst_optval
	optskuxmlstring = SL_ProdInfoDetails.lst_optsku
	optakuFirstCodestring = SL_ProdInfoDetails.lst_optfirstsku
	StockAttrib_MinPrice = SL_ProdInfoDetails.lStockMinPrice
	StockAttrib_MaxPrice = SL_ProdInfoDetails.lStockMaxPrice
	Has_Subs_ForDiscont  = SL_ProdInfoDetails.lHas_Subs_ForDiscont	
	
	IsGiftCard = false
	
	if LOSTOCK.GiftCard = true  then
	    IsGiftCard = true	    
	end if	
	
		set SL_cProdReviewObj = New cProdReviewObj
	
					SL_cProdReviewObj.laction="READ_REVIEW"
					SL_cProdReviewObj.llitemnumber=cstr(current_stocknumber)
					xmlstring= sitelink.Customer_Review_Deergear(SL_cProdReviewObj)
					set SL_cProdReviewObj=nothing
					
	
					objDoc.loadxml(xmlstring)
						
					targetnode= "//get_review"
						
					set SL_cust_review = objDoc.selectNodes(targetnode) 
						
					SL_all_rating = SL_cust_review.length
							
					rating_sum = 0
					'get sum of rating
					for x =0 to SL_cust_review.length-1			
						rating_sum =  rating_sum + SL_cust_review.item(x).selectSingleNode("rating").text
					next 
							
					'rating_sum = 26
					'response.write(rating_sum)
					'response.write("<br>")
					'response.write(round(rating_sum/SL_all_rating,0))
					'avg_rating = 0
					'rating_sum=3
					'SL_all_rating=3
							
					if SL_all_rating > 0 then
						'avg_rating = round(rating_sum/SL_all_rating,0)
						avg_rating = rating_sum/SL_all_rating
						half_star=0
						pos=InStr(avg_rating,".")
						if pos > 0 then
							decimal_part = mid(avg_rating,pos)
							avg_rating = mid(avg_rating,1,pos-1)
							if len(decimal_part) > 3 then decimal_part = mid(decimal_part,1,3)
							'Response.Write avg_rating & "," & decimal_part & ","
								
							if decimal_part > 0.01 and decimal_part < 0.5 then 
								avg_rating=avg_rating&"-5"
							else
								avg_rating=avg_rating+1
							end if
						end if
					end if
						
					usethis_rating_img = "images/clear.gif"
					if avg_rating > 0 then
						usethis_rating_img = "images/"&avg_rating&+"note.gif"
					end if
			
	
	
    tmpstocknumber = current_stocknumber
    if LOSTOCK.size_color=true then 
        tmpstocknumber = trim(left(current_stocknumber,10))
    end if 
    if WANT_REWRITE = 1 then
        producturl = insecureurl & usethis_for_urlwrite & "/productinfo/" & tmpstocknumber & "/"
    else
        producturl = insecureurl & "prodinfo.asp?number=" & tmpstocknumber
    end if 
    
    'for recently view products                
     ProductFullSizeImage = "images/noimage.gif"
                
     if len(trim(LOSTOCK.inetimage)) > 0 then
           StrFileName = "images/"+trim(LOSTOCK.inetimage)
	       StrPhysicalPath = Server.MapPath(StrFileName)
		   set objFileName = CreateObject("Scripting.FileSystemObject")	
		     if objFileName.FileExists(StrPhysicalPath) then
		        imagename=StrFileName		        
		        ProductFullSizeImage = StrFileName
			 else
			    imagename="images/noimage.gif"
  			  end if
			set objFileName = nothing
	else
	   imagename="images/noimage.gif"			   
	end if
%>

<html>
<head>
<title>Aussie Products.com | <%=session("althomepage")%> <%=trim(LOSTOCK.inetsdesc)%> | Aussie Foods</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="robots" content="index, follow" />
    <meta name="robots" content="ALL">
	<meta name="distribution" content="Global">
	<meta name="revisit-after" content="5">
	<meta name="classification" content="Food">
<meta NAME="KEYWORDS" CONTENT="<%=keylist%>">
<meta NAME="description" CONTENT="<%=froogle_desc%>">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta property="og:image" content="<%=insecureurl+imagename%>" />
<link rel="canonical" href="<%=producturl%>">

<% if request.servervariables("server_port_secure") = 1 then %>
	<base href="<%=secureurl%>">
<% else %>
	<base href="<%=insecureurl%>">
<% end if %>

<link rel="stylesheet" href="text/topnav.css" type="text/css">

<link rel="stylesheet" href="text/storestyle.css" type="text/css">
<link rel="stylesheet" href="text/footernav.css" type="text/css">
<link rel="stylesheet" href="text/design.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="include/prettyPhoto.css" type="text/css" media="screen" title="prettyPhoto main stylesheet" charset="utf-8" />
<style>
	ul.gallery {list-style: none; margin: 0; padding: 0 0 0 5px; text-align: center;}
	ul.gallery li {float: none; margin: 0; padding: 0; }
	ul.gallery li a {display: block; padding: 1px;}
</style>

<script src="include/jquery-1.6.1.min.js" type="text/javascript"></script>
<script src="include/jquery.prettyPhoto.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript" language="JavaScript" src="calender_old.js"></script>
<script type="text/javascript" language="JavaScript" src="prodinfo.js"></script>
<script type="text/javascript" language="javascript">

function RedirectToAddADDress (lvalue,stockid,dropdownnum,needurlrewrite,rewritedesc){
     var new_url="" ;
     var addparam ="" ;
     for( var i = 1 ; i <= dropdownnum ; i++){
      var attribtype = "slopttype"+i ;
      attribtype = "document.getElementById('"+attribtype+"').value" ;
      attribtypeval = eval(attribtype)
      if (attribtypeval=="D"){
        var Prodattribname = "Attrib"+i ;
        var Prodattribevalname = "document.getElementById('"+Prodattribname+"').selectedIndex" ;
        selIndex= eval(Prodattribevalname) ;
        Prodattribevalname = "document.getElementById('"+Prodattribname+"').options["+selIndex+"].value" ;
        var selIndexvalue = eval(Prodattribevalname) 
      }//if (attribtypeval=="D")
      if (attribtypeval=="R"){
                
        attribradio = eval("document.prodinfoform.Attrib"+i+".length") ;
        j = attribradio
        for (k=0; k<j; k++){
            radiochecked = eval("document.prodinfoform.Attrib"+i+"["+k+"].checked") ;
             if (radiochecked){
                selIndexvalue = eval("document.prodinfoform.Attrib"+i+"["+k+"].value") ;               
             }
        }
      } //if (attribtypeval=="R")
      if (i==1){
        addparam = selIndexvalue
        }
        else{
        addparam = addparam + "|"+selIndexvalue
        }
     }
    if (needurlrewrite==1){
		new_url = "<%=insecureurl%>"+ rewritedesc + "/productinfo/" + stockid  + "/" + addparam + "/"
		}
	else{	
	    new_url = "prodinfo.asp?number="+stockid+"&opt="+addparam
		}
    window.location=new_url;
}
function GoToAddADDress (page)
   {
	var new_url = '<%=session("insecureurl")%>'+page
   //if (  (new_url != "")  &&  (new_url != null)  )
    window.location=new_url;
	//alert(new_url)
   }
</script>
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">

<% 
	session("destpage")="prodinfo.asp?number="+cstr(current_stocknumber) 
	session("viewpage")=session("destpage")

%>



<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content" class="prodinfopage">
	
	<%if session("department")= "-1" then 
        if WANT_REWRITE = 1 then
			uselink= insecureurl & "Specials/0/0"
		else
			uselink= insecureurl & "Specials.asp"
		end if
    %>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
    <br>
		<tr><td class="breadcrumbrow"><a href="<%=insecureurl%>default.asp" class="breadcrumb">Home</a>&nbsp;&nbsp;&#47;&nbsp;&nbsp;<a href="<%=uselink%>" title="Specials" class="breadcrumb">Specials</a></td></tr>
	</table>
	<%end if %>
	
	
	<%if session("department") > 0 then %>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
				<tr><br><td class="breadcrumbrow" width="100%">
						<a href="default.asp" class="breadcrumb">Home</a>&nbsp;&nbsp;&#47;&nbsp;&nbsp;
						<%
						deptstr = sitelink.deptstring(session("department"),true,"&nbsp;&nbsp;&#47;&nbsp;&nbsp;","breadcrumb","",WANT_REWRITE)
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
		<tr><td><H1><%=trim(LOSTOCK.inetsdesc)%></H1></td></tr>				
	</table>
	
		
	<form enctype="application/x-www-form-urlencoded" method="post" name="prodinfoform" id="prodinfoform" onSubmit="return Process()" action="itemadd.asp">
	<table width="100%" class="table-layout-fixed" cellpadding="0" cellspacing="0" border="0">
	
	        <tr> 
              <td width="auto" valign="top"> 
			  <center>
                
                <img SRC="<%=imagename%>" alt="<%=trim(LOSTOCK.inetsdesc)%>" title="<%=trim(LOSTOCK.inetsdesc)%>" hspace="5" border="0" align="middle" class="ProdInfoImage"> 
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
			     			largeimagename=insecureurl + StrFileName
							found = true
					else
						largeimagename="images/noimage.gif"
						found = false
	  				end if
					set objFileName = nothing		
				end if
				  
				 	
			%>
                <script type="text/javascript" language="javascript">

				PositionX = 100;
				PositionY = 100;
				
				// Set these value approximately 20 pixels greater than the
				// size of the largest image to be used (needed for Netscape)
				
				defaultWidth  = 500;
				defaultHeight = 500;
				
				// Set autoclose true to have the window close automatically
				// Set autoclose false to allow multiple popup windows
				
				var AutoClose = true;
				
				// Do not edit below this line...
				// ================================
				if (parseInt(navigator.appVersion.charAt(0))>=4){
				var isNN=(navigator.appName=="Netscape")?1:0;
				var isIE=(navigator.appName.indexOf("Microsoft")!=-1)?1:0;}
				var optNN='scrollbars=no,width='+defaultWidth+',height='+defaultHeight+',left='+PositionX+',top='+PositionY;
				var optIE='scrollbars=no,width=150,height=100,left='+PositionX+',top='+PositionY;
				function ShowlargeImage(imageURL,imageTitle){
				if (isNN){imgWin=window.open('about:blank','',optNN);}
				if (isIE){imgWin=window.open('about:blank','',optIE);}
				with (imgWin.document){
				writeln('<html><head><title>Loading...</title><style>body{margin:0px;}</style>');writeln('<sc'+'ript>');
				writeln('var isNN,isIE;');writeln('if (parseInt(navigator.appVersion.charAt(0))>=4){');
				writeln('isNN=(navigator.appName=="Netscape")?1:0;');writeln('isIE=(navigator.appName.indexOf("Microsoft")!=-1)?1:0;}');
				writeln('function reSizeToImage(){');writeln('if (isIE){');writeln('window.resizeTo(300,300);');
				writeln('width=300-(document.body.clientWidth-document.images[0].width);');
				writeln('height=300-(document.body.clientHeight-document.images[0].height);');
				writeln('window.resizeTo(width,height);}');writeln('if (isNN){');       
				writeln('window.innerWidth=document.images["SiteLink"].width;');writeln('window.innerHeight=document.images["SiteLink"].height;}}');
				writeln('function doTitle(){document.title="'+imageTitle+'";}');writeln('</sc'+'ript>');
				if (!AutoClose) writeln('</head><body bgcolor=000000 scroll="no" onload="reSizeToImage();doTitle();self.focus()">')
				else writeln('</head><body bgcolor=000000 scroll="no" onload="reSizeToImage();doTitle();self.focus()" onblur="self.close()">');
				writeln('<img name="SiteLink" src='+imageURL+' style="display:block"></body></html>');
				close();		
				}}
				
				</script>
                <% if found = true then %>
                <br>
               
                <center>
                  <a class="allpage" href="javascript:ShowlargeImage('<%=largeimagename%>','<%=trim(LOSTOCK.inetsdesc)%>')"><font color="red"><u><b>View 
                  Large Image</b></u></font></a> 
                </center>
                <%end if%>

				<%
				 if len(trim(LOSTOCK.sl_image1)) > 0 or len(trim(LOSTOCK.sl_image2)) > 0 or len(trim(LOSTOCK.sl_image3)) > 0 or len(trim(LOSTOCK.sl_image4)) > 0 or len(trim(LOSTOCK.sl_image5)) > 0 or len(trim(LOSTOCK.sl_image6)) > 0 or len(trim(LOSTOCK.sl_image7)) > 0 or len(trim(LOSTOCK.sl_image8)) > 0 then
	                
	                Dim SL_AddImages(8)
	                count=0


	                sl_image1 = trim(LOSTOCK.sl_image1)
	                StrFileName = "images/"+cstr(sl_image1)
	                StrPhysicalPath = Server.MapPath(StrFileName)
	                set objFileName = CreateObject("Scripting.FileSystemObject")	
	                if objFileName.FileExists(StrPhysicalPath) then
   		                count = count + 1
   		                SL_AddImages(count)=sl_image1
	                end if
	                set objFileName = nothing


	                sl_image2 = trim(LOSTOCK.sl_image2)
	                StrFileName = "images/"+cstr(sl_image2)
	                StrPhysicalPath = Server.MapPath(StrFileName)
	                set objFileName = CreateObject("Scripting.FileSystemObject")	
	                if objFileName.FileExists(StrPhysicalPath) then
   		                count = count + 1
   		                SL_AddImages(count)=sl_image2
	                end if
	                set objFileName = nothing
                	
	                sl_image3 = trim(LOSTOCK.sl_image3)
	                StrFileName = "images/"+cstr(sl_image3)
	                StrPhysicalPath = Server.MapPath(StrFileName)
	                set objFileName = CreateObject("Scripting.FileSystemObject")	
	                if objFileName.FileExists(StrPhysicalPath) then
   		                count = count + 1
   		                SL_AddImages(count)=sl_image3
	                end if
	                set objFileName = nothing
                	

	                sl_image4 = trim(LOSTOCK.sl_image4)
	                StrFileName = "images/"+cstr(sl_image4)
	                StrPhysicalPath = Server.MapPath(StrFileName)
	                set objFileName = CreateObject("Scripting.FileSystemObject")	
	                if objFileName.FileExists(StrPhysicalPath) then
   		                count = count + 1
   		                SL_AddImages(count)=sl_image4
	                end if
	                set objFileName = nothing
                	
                	
	                sl_image5 = trim(LOSTOCK.sl_image5)
	                StrFileName = "images/"+cstr(sl_image5)
	                StrPhysicalPath = Server.MapPath(StrFileName)
	                set objFileName = CreateObject("Scripting.FileSystemObject")	
	                if objFileName.FileExists(StrPhysicalPath) then
   		                count = count + 1
   		                SL_AddImages(count)=sl_image5
	                end if
	                set objFileName = nothing

	                sl_image6 = trim(LOSTOCK.sl_image6)
	                StrFileName = "images/"+cstr(sl_image6)
	                StrPhysicalPath = Server.MapPath(StrFileName)
	                set objFileName = CreateObject("Scripting.FileSystemObject")	
	                if objFileName.FileExists(StrPhysicalPath) then
   		                count = count + 1
   		                SL_AddImages(count)=sl_image6
	                end if
	                set objFileName = nothing
                	
	                sl_image7 = trim(LOSTOCK.sl_image7)
	                StrFileName = "images/"+cstr(sl_image7)
	                StrPhysicalPath = Server.MapPath(StrFileName)
	                set objFileName = CreateObject("Scripting.FileSystemObject")	
	                if objFileName.FileExists(StrPhysicalPath) then
   		                count = count + 1
   		                SL_AddImages(count)=sl_image7
	                end if
	                set objFileName = nothing
                	
	                sl_image8 = trim(LOSTOCK.sl_image8)
	                StrFileName = "images/"+cstr(sl_image8)
	                StrPhysicalPath = Server.MapPath(StrFileName)
	                set objFileName = CreateObject("Scripting.FileSystemObject")	
	                if objFileName.FileExists(StrPhysicalPath) then
   		                count = count + 1
   		                SL_AddImages(count)=sl_image8
	                end if
	                set objFileName = nothing
				 
				%>
				<br><br>
                <ul class="gallery">
                <%
                if count > 0 then 
                    for x=1 to count
                %>
                    
					<%if x=1 then%>
                    <li><a href="images/<%=SL_AddImages(x)%>" rel="prettyPhoto[gallery] nofollow"><img src="images/btn-addtnl-images.gif" border="0" /></a></li>
                    <%else %>
                    <li style="display: none;"><a href="images/<%=SL_AddImages(x)%>" rel="prettyPhoto[gallery] nofollow"><img src="images/clear.gif" height="1" width="1" /></a></li>
                    <%end if%>
                <%
                    next
                end if%>
                </ul>
                
				<%end if%>
					</center>


              </td>
			                
                <td width="50%" valign="top" class="prodinfocell">
					<input type="hidden" name="addtocart" id="addtocart"><input type="hidden" name="addtowishlist" id="addtowishlist">
					<input type="hidden" name="additem" value="<%=current_stocknumber%>">					

					<table width="100%" border="0"  cellpadding="8" cellspacing="0">					
						<tr><td class="THHeader" width="100%">
							<div class="innerprodcell tdRow1Color">
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" class="tdRow1Color">
							
							
								<tr><td class="plaintextbold">Item Number:&nbsp;<%=current_stocknumber%></td></tr>
								
							    <tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0"></td></tr>
							
							    <%
							    if SHOW_PRODUCTREVIEWS = 1 then
							    if SL_cust_review.length > 0 then%>
							        <tr><td style="height:5px;"></td></tr>
							        <tr><td class="plaintext">&nbsp;<span class="reviewtext">Review Average:</span> &nbsp;<%if avg_rating > 0 then%><img src="<%=usethis_rating_img%>" border="0" width="70" height="14"><%end if%></td></tr>
							        <tr><td class="plaintext">&nbsp;<span class="reviewtext">Number of Reviews:</span> <%=SL_all_rating%></td></tr>
							        <tr><td height="3"></td></tr>
								    <%
								        if WANT_REWRITE=1 then
								            newbookmarktitle = url_cleanse(bookmarktitle)
									        new_url = insecureurl+ newbookmarktitle + "/productinfo/" + current_stocknumber  + "/"
									    else
									        new_url = insecureurl+ "prodinfo.asp?number=" + current_stocknumber  
									    end if 
								    %>
    							    <tr><td class="plaintext">&nbsp;<a href="<%=new_url%>#reviews" class="prodlink">View Reviews</a> |&nbsp;<a class="prodlink" href="writeReview.asp?number=<%=current_stocknumber%>">Review this item</a></td></tr>
	    						<%else%>
		    					  <tr>
					                      <td class="plaintext" valign="middle">&nbsp;<img src="images/0note.gif" border="0" align="absmiddle">&nbsp;<a class="prodlink" href="writeReview.asp?number=<%=current_stocknumber%>">Be the first to review this item</a></td>
			                      </tr>
							<%end if
							 end if  'if SHOW_PRODUCTREVIEWS
							%>
							
							<tr><td style="height:10px;"></td></tr>
							<%
							
								'SL_prefShip_Title = ""
							    'SL_prefShip_Method = trim(LOSTOCK.prefship)
								'if len(SL_prefShip_Method) > 0 then
									SL_prefShip_Title=SL_ProdInfoDetails.lPrefShipMethod
								'end if
								
							  if LOSTOCK.size_color = true then
							  	 extra_field = ""					 			
								 objDoc.loadxml(SL_SizeColorXmlStream)									
								set SL_sbcl = objDoc.selectNodes("//sbcl") 
							  end if 
															
							  has_special_price=false
							  has_Onespecial_Price=false
							  extra =""
							  AllSpecialXmlstring = SL_SpecialPriceXmlStream							  
							  objDoc.loadxml(SL_SpecialPriceXmlStream)
							  							  
							  set SL_gsp = objDoc.selectNodes("//gsp")
							  if SL_gsp.length > 0 then has_special_price=true
							  
							  tnode = "//gsp[qty=1 and total_dol <="+cstr(cdbl(session("SL_BasketSubTotal")))+"]"
							  Set SL_QtyOnePricing = objDoc.selectNodes(tnode)
							  
							  
							  if SL_QtyOnePricing.length > 0 then
									SL_SP_price = SL_gsp.item(SL_QtyOnePricing.length-1).selectSingleNode("price").text
									SL_SP_discount = SL_gsp.item(SL_QtyOnePricing.length-1).selectSingleNode("discount").text
									
									SL_SP_CostMeth = SL_gsp.item(SL_QtyOnePricing.length-1).selectSingleNode("costmethod").text
									SL_SP_CostPlusDiscount = SL_gsp.item(SL_QtyOnePricing.length-1).selectSingleNode("costplus").text
									SL_SP_StockReOrdPrice = SL_gsp.item(SL_QtyOnePricing.length-1).selectSingleNode("reordprice").text
									SL_SP_UnCostprice = SL_gsp.item(SL_QtyOnePricing.length-1).selectSingleNode("uncost").text
									
									if SL_SP_CostMeth="" then
										unitprice = LOSTOCK.price1
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
								
								set SL_QtyOnePricing=nothing
								
								set SL_gsp = nothing
								
							objDoc.loadxml(AllSpecialXmlstring)
							set SL_gsp = objDoc.selectNodes("//gsp[qty>1]")
							
							%>
							<%if LOSTOCK.size_color = false then %>
							<tr><td class="ProductPrice">
								<%if has_Onespecial_Price=true then %>
								<table>
								 <tr><td>
								 		Unit Price: <s><%=FORMATCURRENCY(LOSTOCK.price1)%></s>
								 	</td>
								  </tr>
								  <tr>
									<td>
									<table cellpadding="0" cellspacing="0" border="0">
									<tr><td class="nopadding"><img src="images/RedSaleTag.gif" style="border:0;width=53;height:23;"></td>
										<td class="plaintext nopadding" align="center" style="background-image:url(images/RedSaleTag_bkg.gif);height:23">
										<span style="font-weight:bold;color:White;"><%=formatcurrency(discountedUnitPrice)%></span>&nbsp;
										</td>
									</tr>
									</table>
									
									</td>
								</tr>
								</table>
							   <%else%>			
							        <%'if IsGiftCard=false then %>				        
							   	    Unit Price: <%=FORMATCURRENCY(LOSTOCK.price1)%>
							   	    <%'end if %>
							   <%end if%>
							</td></tr>
							<%end if%>
							
							
							<%if SHOW_COMP_PRICE="TRUE" and LOSTOCK.inetcprice > 0 then %>
								<tr><td class="CompPrice">Compare At: <%=FORMATCURRENCY(LOSTOCK.inetcprice)%></td></tr>
							<%end if%>
							
							<%if (has_special_price=true) and ( (LOSTOCK.size_color = true) or (LOSTOCK.size_color = false and has_Onespecial_Price=false) ) then %>
								<tr><td class="smalltextblk" style="color:#999999;"><img src="images/RedSaleTag.gif" border="0" align="absmiddle"  /> <%if has_stock_attribs=true then response.Write " Sale price will be displayed in the cart." end if %></td></tr>
							<%end if%>

							<%
							

							  'if SL_gsp.length > 0 and has_Onespecial_Price=false then 
							  if SL_gsp.length > 0 and has_stock_attribs=false then %>
							  <tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0"></td></tr>
							  <tr><td valign="top">

							  <table width="100%" border="0" cellspacing="1" cellpadding="1">
							   <tr><td class="THHeader" colspan="3">Special Pricing</td></tr>
							   <%

								for x=0 to SL_gsp.length-1
									'SL_SP_Number = SL_gsp.item(x).selectSingleNode("number").text
									SL_SP_qty = SL_gsp.item(x).selectSingleNode("qty").text
									'if SL_SP_qty > 1 then
									SL_SP_price = SL_gsp.item(x).selectSingleNode("price").text
									SL_SP_discount = SL_gsp.item(x).selectSingleNode("discount").text
									SL_SP_ordertotal = SL_gsp.item(x).selectSingleNode("total_dol").text
									SL_SP_desc2 = SL_gsp.item(x).selectSingleNode("desc2").text
									
									if LOSTOCK.size_color=false then
										SL_SP_desc2=""
									end if 

					                SL_SP_CostMeth = SL_gsp.item(0).selectSingleNode("costmethod").text
					                SL_SP_CostPlusDiscount = SL_gsp.item(0).selectSingleNode("costplus").text
										
					                if SL_SP_CostMeth="" then
									    unitprice = LOSTOCK.price1
									    unitprice2 = SL_SP_price
									    if unitprice2>0 then
										    unitprice = unitprice2
									    end if
									    percentoff = (SL_SP_discount/100.0)
									    discountedUnitPrice = cdbl(unitprice * (1- percentoff))							
					                end if

					                if SL_SP_CostMeth="P" then
    					                SL_SP_StockReOrdPrice = SL_gsp.item(0).selectSingleNode("reordprice").text
						                discountedUnitPrice= cdbl(SL_SP_StockReOrdPrice*(1+SL_SP_CostPlusDiscount/100.0))
					                end if
                							
					                if SL_SP_CostMeth="U" then										
    					                SL_SP_UnCostprice = SL_gsp.item(0).selectSingleNode("uncost").text
						                discountedUnitPrice= cdbl(SL_SP_UnCostprice*(1+SL_SP_CostPlusDiscount/100.0))									
					                end if
									
									SL_SP_qty = replace(SL_SP_qty,".00","")
							
									if (x mod 2) = 0 then
										 class_to_use = "tdRow1Color"
									else
										class_to_use = "tdRow2Color"
									end if 
									
									threshholdTxt = ""
									if SL_SP_ordertotal > 0 then
										threshholdTxt= "&nbsp;for order over "+formatcurrency(SL_SP_ordertotal)
									end if

								%>
								<tr>
								<td class="<%=class_to_use%>"><span class="plaintext">&nbsp;Buy&nbsp;<%=SL_SP_qty%>&nbsp;<%=SL_SP_desc2%>&nbsp;for&nbsp;<%=formatcurrency(discountedUnitPrice)%>&nbsp;each<%=threshholdTxt%></span></td>
								</tr>
								
							<%
							'end if
							next %>
							  </table>
							  </td>
							  </tr>
							  
							  <%end if %>
							  
							  <%
							 
							  if SHOW_IN_STOCK =1 and LOSTOCK.size_color = false then 
								avail_status = "In Stock"
					  			if (LOSTOCK.units > 0) or (LOSTOCK.dropship = true) or (LOSTOCK.construct = true) then
									avail_status="In Stock"
							  	else
									if SHOW_DUE_DATE = 1 then									    
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

							 <%'if len(trim(LOSTOCK.prodavail)) > 0 then %>
							  <tr><td class="plaintext"><%=LOSTOCK.prodavail%></td></tr>
							 <%'end if%>					
							<tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0"></td></tr>
							<% if SL_prefShip_Title<>"" then %>
							<tr><td class="plaintext"><br>Will be shipped via: <%=SL_prefShip_Title%></td></tr>
							<%end if%>
							<% if SHOW_PRODSHIPCHRG=1 then 
							    ProdShipCharge = trim(LOSTOCK.shipcharge)
							    if len(ProdShipCharge)>0 and IsNumeric(ProdShipCharge)=true and ProdShipCharge>0 then %>
							    <tr><td class="plaintext"><br>Product specific shipping charge: <%=formatcurrency(ProdShipCharge)%></td></tr>
							    <tr><td height="10"></td></tr>
							    <%end if%>
							<%end if%>
							
							<% 							 
							if LOSTOCK.size_color = true and has_stock_attribs=false then %>
							<tr><td class="THHeader">Size/Color</td></tr>
							<tr><td>
							
							<select NAME="txtvariantVar" class="smalltextblk" onchange="GoToAddADDress(this.options[this.selectedIndex].value)">
							
							<%
								 for x=0 to SL_sbcl.length-1
									SL_Variation= SL_sbcl.item(x).selectSingleNode("scolor").text
									SL_VarationDesc= SL_sbcl.item(x).selectSingleNode("desc2").text
									SL_VarationPrice= SL_sbcl.item(x).selectSingleNode("price1").text
									SL_VarationUnits= SL_sbcl.item(x).selectSingleNode("units").text
									SL_VarationDrop= SL_sbcl.item(x).selectSingleNode("dropship").text
									SL_VarationConst= SL_sbcl.item(x).selectSingleNode("construct").text
									
									if SL_VarationUnits=".00" then
									    SL_VarationUnits = 0 
									end if
									
									target_in_SpecialPrice = "//gsp[number2='"+cstr(SL_Variation)+"' and qty=1 and total_dol <="+cstr(cdbl(session("SL_BasketSubTotal")))+"]"
									set SP_PriceObj = ObjDoc.selectNodes(target_in_SpecialPrice)
									if SP_PriceObj.length > 0 then										
											SL_SP_price = SP_PriceObj.item(0).selectSingleNode("price").text
											SL_SP_discount = SP_PriceObj.item(0).selectSingleNode("discount").text

											unitprice = SL_VarationPrice
											unitprice2 = SL_SP_price
											if unitprice2>0 then
												unitprice = unitprice2
											end if
											percentoff = (SL_SP_discount/100.0)
											discountedUnitPrice = cdbl(unitprice * (1- percentoff))
											SL_VarationPrice=discountedUnitPrice
									end if

									set SP_PriceObj =nothing
									
									if WANT_REWRITE =1 then
										addthisurl = "&addthis="+ usethis_for_urlwrite
									else
										addthisurl=""
									end if	
									if x= 0 then
							           usethisvariation = SL_Variation							        
							        end if			
							        
																	
										
								 %>
							<option value="getvariation.asp?number=<%=current_stocknumber%>&whatvar=<%=SL_Variation%><%=addthisurl%>"
								<% if current_variation = cstr(SL_Variation) then response.write(" selected") end if%>							
							>
							<%=SL_VarationDesc%>&nbsp;<%=formatcurrency(SL_VarationPrice)%>
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
							<%
							 if len(trim(current_variation)) > 0  then
							          usethisvariation = current_variation
 					        end if
							
							 %>
							<input type="hidden" name="txtvariant" value="<%=usethisvariation%>">
							</td></tr>
							<%end if%>
							
							
							<%if has_stock_attribs=true then								
								'set SLProdAttribObj = New cProductAttribObj
								'SLProdAttribObj.lnumber= cstr(current_stocknumber)
								
								
								'call sitelink.GET_SLSTOCK_ATTRIBUTE(SLProdAttribObj)
								
								'optcodexmlstring = SLProdAttribObj.lst_optcode
								'optvalxmlstring = SLProdAttribObj.lst_optval
								'optskuxmlstring = SLProdAttribObj.lst_optsku
								'optakuFirstCodestring = SLProdAttribObj.lst_optfirstsku
								'StockAttrib_MinPrice = SLProdAttribObj.lStockMinPrice
								'StockAttrib_MaxPrice = SLProdAttribObj.lStockMaxPrice
								
								'set SLProdAttribObj =nothing
								
								maxoption = 0 
								maxAttriboption = 0
								objDoc.loadxml(optvalxmlstring)								
								set SL_optVal = objDoc.selectNodes("//loptval") 
								maxoption = SL_optVal.length-1
								set SL_optVal = nothing
								
								
								objDoc.loadxml(optcodexmlstring)
								target_optcodenode ="//loptcode[opttype!='T' and skulink='true']"
								set SL_optcode = objDoc.selectNodes(target_optcodenode) 								
								maxoption = maxoption + SL_optcode.length-1
								maxAttriboption = SL_optcode.length-1
								'set SL_optcode = nothing
								
								'target_optcodenodeNONSku ="//loptcode[opttype='T']"
								target_optcodenodeNONSku ="//loptcode[skulink='false']"
								set SL_optcodeNONSku = objDoc.selectNodes(target_optcodenodeNONSku)
								maxTextoption = SL_optcodeNONSku.length-1
								'set SL_optcodeNONSkue = nothing
								
								'response.Write(maxoption)
								
								'target_optcodenode ="//loptcode"								
								'set SL_optcode = objDoc.selectNodes(target_optcodenode) 
								
								
								'init all request.quertstring.
								
								'current_selection = request.QueryString("opt")								
								current_selection= attribstring
								
								for x=1 to maxoption								    
								    session("Current_AttribSelected"&x)=""							
								next 
								
								MyArray = Split(current_selection, "|", -1, 1)
								dim initselect	
								x=1								
								for each initselect in MyArray								
								    session("Current_AttribSelected"&x)=initselect								    
								    x=x+1
								  
								next
								
								'for x=1 to maxoption								    
								'    response.write("prodoptionval -->"&session("Current_AttribSelected"&x) &"<br>")					
								'next 
								'response.Write("<br>=====================<br>")
								
								for x=0 to SL_optcode.length-1
									sloptcode = SL_optcode.item(x).selectSingleNode("optcode").text									
									slopttitle = SL_optcode.item(x).selectSingleNode("opttitle").text
									slopttype = SL_optcode.item(x).selectSingleNode("opttype").text
									sloptlen = SL_optcode.item(x).selectSingleNode("vallen").text
									
								 
							%>
								<tr><td class="plaintextbold">
								<table border="0">
								<tr><td class="plaintextbold" valign="top"><%=slopttitle%>
								<input type="hidden" name="slopttype<%=(x+1)%>" id="slopttype<%=(x+1)%>" value="<%=slopttype%>">
								<input type="hidden" name="SkuAttribTitle<%=(x+1)%>" id="SkuAttribTitle<%=(x+1)%>" value="<%=slopttitle%>">
								
								</td>
								<td valign="top">																
								
								<%
								
								current_selectionlen = len(trim(current_selection))
								
								valfilter =""
								if current_selectionlen > 0 then	
								
								    matchonlythis = ""
								    
								    for m=1 to x
								      matchonlythis = matchonlythis + session("Current_AttribSelected"&m) + "|"								    
								    next
								    matchonlythislen =len(matchonlythis)
								    		
								    newcurrent_selection = cstr(current_selection) + "|"					    								    
								    target_node = "//loptsku[optskukey[substring(.,1,"+cstr(matchonlythislen)+")='"+cstr(matchonlythis)+"']]"								    
    								objDoc.loadxml(optskuxmlstring)
								
								
								set SL_optsku = objDoc.selectNodes(target_node) 
								    for z= 0 to SL_optsku.length-1
								        sl_opskukey = SL_optsku.item(z).selectSingleNode("optskukey").text
								        MyArray = Split(sl_opskukey, "|", -1, 1)
								        sl_opskukey = "optvalkey = " + cstr(MyArray(x)) + " and opttitle='"+slopttitle+"'"
								        sl_opskukey = "(" + sl_opskukey + ")"
								        valfilter = valfilter + sl_opskukey + " or "
								    next   
								set SL_optsku= nothing
								
								end if
								
								lenvalfilter = len(valfilter)
								
								if lenvalfilter > 0 then								
								    valfilter = mid(valfilter,1, len(valfilter)-3)
								    'valfilter = valfilter + " and opttitle='"+slopttitle+"'"
								end if
								 %>
								 
								 <%if slopttype="D" then%>
								 <select name="Attrib<%=(x+1)%>" id="Attrib<%=(x+1)%>"  onChange="RedirectToAddADDress(this.options[this.selectedIndex].value,'<%=current_stocknumber%>',<%=x+1%>,<%=WANT_REWRITE%>,'<%=usethis_for_urlwrite%>')" >
								 <option value="">Select 
								 <%
								 
								 if x=0 then
								    objDoc.loadxml(optakuFirstCodestring)
								    target_node = "//firstoptsku[opttitle='"+cstr(slopttitle)+"']"
								    set SL_optvalsku = objDoc.selectNodes(target_node) 
								    
								 else	
								    if lenvalfilter > 0 then	
								        objDoc.loadxml(optvalxmlstring)
								        target_node = "//loptval["+valfilter+"]"	
								        'response.Write(target_node)							        
								        set SL_optvalsku = objDoc.selectNodes(target_node)
								    end if 
								 
								 end if 
								 
								 if lenvalfilter > 0 or x=0 then
								    for y= 0 to SL_optvalsku.length-1 
								    sloptvalue = SL_optvalsku.item(y).selectSingleNode("optvalue").text
								    sloptvalkey = SL_optvalsku.item(y).selectSingleNode("optvalkey").text
								    sloptprice = SL_optvalsku.item(y).selectSingleNode("optprice").text
								    
								    
								    
								    additionalPrice= ""
								    if sloptprice > 0 then
								        sloptprice =formatcurrency(sloptprice)
								        additionalPrice = "(+" + cstr(replace(sloptprice,".00","")) + ")"								  
								    end if 
								    if sloptprice < 0 then
								        sloptprice =formatcurrency(sloptprice)
								        additionalPrice =  cstr(replace(sloptprice,".00","")) 
								        additionalPrice =  replace(additionalPrice,"(","(-")
								    end if
								    
								    
								  %>
								  <option value="<%=sloptvalkey%>" <%if session("Current_AttribSelected"&(x+1)) = sloptvalkey then response.write(" selected") end if%>>
								  <%=sloptvalue%> <%=additionalPrice%>
								  <%next%>
								 
								 <%end if 'if lenvalfilter > 0 or x=0 then%>
								
								 </select>
								<%
								end if 'if slopttype="D"  %>
								
								<%if slopttype="R" then%>
								<% 
								 
								 if x=0 then
								    
								    objDoc.loadxml(optakuFirstCodestring)
								    target_node = "//firstoptsku[opttitle='"+cstr(slopttitle)+"']"
								    set SL_optvalsku = objDoc.selectNodes(target_node) 								    
								 else	
								    if lenvalfilter > 0 then	
								        objDoc.loadxml(optvalxmlstring)
								        target_node = "//loptval["+valfilter+"]"		
								        'response.Write(target_node)						        
								        set SL_optvalsku = objDoc.selectNodes(target_node)
								    end if 
								 
								 end if 
								 
								 
								 
								 if lenvalfilter > 0 or x=0 then
								    for y= 0 to SL_optvalsku.length-1 
								    sloptvalue =  SL_optvalsku.item(y).selectSingleNode("optvalue").text
								    sloptvalkey = SL_optvalsku.item(y).selectSingleNode("optvalkey").text
								    sloptprice = SL_optvalsku.item(y).selectSingleNode("optprice").text
								    
								    additionalPrice= ""
								    if sloptprice > 0 then
								        sloptprice =formatcurrency(sloptprice)
								        additionalPrice = "(+" + cstr(replace(sloptprice,".00","")) + ")"								  
								    end if 
								    if sloptprice < 0 then
								        sloptprice =formatcurrency(sloptprice)
								        additionalPrice =  cstr(replace(sloptprice,".00","")) 
								        additionalPrice =  replace(additionalPrice,"(","(-")
								    end if
								    
								%>
								    
								    <input type="hidden" name="noradio<%=(x+1)%>" id="noradio<%=(x+1)%>" value="1">
								    <input type="radio" name="Attrib<%=(x+1)%>" id="Attrib<%=(x+1)%>" value="<%=sloptvalkey%>" <%if session("Current_AttribSelected"&(x+1)) = sloptvalkey then response.write(" checked") end if%>  onClick="RedirectToAddADDress(<%=sloptvalkey%>,'<%=current_stocknumber%>',<%=x+1%>,<%=WANT_REWRITE%>,'<%=usethis_for_urlwrite%>')" >
								    <span class="plaintext"><%=sloptvalue%>
								    <%=additionalPrice %>
								    </span><br>
								    <%next%>
								<%end if %>
								    <input type="hidden" name="noradio<%=(x+1)%>"  id="noradio<%=(x+1)%>" value="0">
								<%'end if 'if lenvalfilter > 0 or x=0 then%>
								
								<%end if 'if slopttype="R" %>
								
								
								<%
								set SL_optvalsku=nothing
								set SL_optval = nothing
								
								%>
								
								</td></tr>
								</table>
																	
								</td></tr>
							
							<%next
							set SL_optcode=nothing 
							%>
							<% 
								for x=0 to SL_optcodeNONSku.length-1 
									sloptcode = SL_optcodeNONSku.item(x).selectSingleNode("optcode").text									
									slopttitle = SL_optcodeNONSku.item(x).selectSingleNode("opttitle").text
									slopttype = SL_optcodeNONSku.item(x).selectSingleNode("opttype").text
									sloptlen = SL_optcodeNONSku.item(x).selectSingleNode("vallen").text
									sloptvaltype = SL_optcodeNONSku.item(x).selectSingleNode("valtype").text
									sloptvalDes = SL_optcodeNONSku.item(x).selectSingleNode("valdec").text
									
									if sloptvaltype="D" then
										sloptlen=10										
									end if
									
									if sloptvaltype<>"N" then
										sloptvalDes=0
									end if
							%>
							<tr><td class="plaintextbold">
								<table border="0">
								<tr><td class="plaintextbold" valign="top"><%=slopttitle%>
								<input type="hidden" name="NonSkuslopttype<%=(x+1)%>" id="NonSkuslopttype<%=(x+1)%>" value="<%=slopttype%>">
								<input type="hidden" name="NonSkuAttribTitle<%=(x+1)%>" id="NonSkuAttribTitle<%=(x+1)%>" value="<%=slopttitle%>">
								<input type="hidden" name="NonSkusvaltype<%=(x+1)%>" id="NonSkusvaltype<%=(x+1)%>" value="<%=sloptvaltype%>">
								<input type="hidden" name="NonSkuNumericBeforeDec<%=(x+1)%>" id="NonSkuNumericBeforeDec<%=(x+1)%>" value="<%=sloptlen%>">
								<input type="hidden" name="NonSkuNumericAfterDec<%=(x+1)%>" id="NonSkuNumericAfterDec<%=(x+1)%>" value="<%=sloptvalDes%>">
								</td>
								<td valign="top" class="plaintext">
								
								 <%if sloptvaltype="N" then %>
								 	<input type="text" name="NonSkuAttribtext<%=(x+1)%>" id="NonSkuAttribtext<%=(x+1)%>" maxlength="<%=sloptlen+1%>" size="10" class="plaintext" onBlur="extractNumber(this,0,true);" onKeyUp="extractNumber(this,0,true);" onKeyPress="return blockNonNumbers(this, event, false, true,<%=sloptlen%>);">									
								 	<%if sloptvalDes > 0 then%>
										 &nbsp;.&nbsp;<input type="text" name="NonSkuAttribtextDec<%=(x+1)%>" id="NonSkuAttribtextDec<%=(x+1)%>" maxlength="<%=sloptvalDes+1%>" size="10" class="plaintext" onBlur="extractNumber(this,0,false);" onKeyUp="extractNumber(this,0,false);" onKeyPress="return blockNonNumbers(this, event, false, false,<%=sloptvalDes%>);">			
										<%end if%>
									&nbsp;(Numeric Values Only)
								 <%end if%>
								 
								 
								 <%if sloptvaltype="D" then %>
								 	<input type="text" name="NonSkuAttribtext<%=(x+1)%>" id="NonSkuAttribtext<%=(x+1)%>" maxlength="<%=sloptlen%>" size="10" class="plaintext">
								   <a href="javascript:show_calendar('prodinfoform.NonSkuAttribtext<%=(x+1)%>');"><img src="images/calendarSm.gif" border="0" align="center" WIDTH="25" HEIGHT="19"></a>
								 <%end if%>
								 
								 <%if sloptvaltype="C" then %>
								    (<span id="chleft<%=(x+1)%>"><%=sloptlen%> max characters</span>)
								    <br />
								    <textarea name="NonSkuAttribtext<%=(x+1)%>" id="NonSkuAttribtext<%=(x+1)%>" cols="20" rows="4" onkeyup="Contchar('NonSkuAttribtext<%=(x+1)%>','chleft<%=(x+1)%>','{CHAR} characters left','<%=sloptlen%>');"></textarea>								   								 
								 <%end if%>
								 								
								
									<%if slopttype="D" then%>
										<select name="NonSkuAttribtext<%=(x+1)%>">
										    <option value="">Select</option>
										<%
										    objDoc.loadxml(optvalxmlstring)										   
										    target_node = "//loptval[optcode='"+cstr(sloptcode)+"' and opttitle='"+cstr(slopttitle)+"']"
										    set sl_optvalnode = objDoc.selectNodes(target_node)
										    for y= 0 to sl_optvalnode.length-1 
										        optvaluenode = sl_optvalnode.item(y).selectSingleNode("optvalue").text
												sloptprice   = sl_optvalnode.item(y).selectSingleNode("optprice").text
												sloptpriceval = sloptprice
												additionalPrice= ""
												if sloptprice > 0 then
													sloptprice =formatcurrency(sloptprice)
													additionalPrice = "(+" + cstr(replace(sloptprice,".00","")) + ")"								  
												end if 
												if sloptprice < 0 then
													sloptprice =formatcurrency(sloptprice)
													additionalPrice = cstr(replace(sloptprice,".00","")) 
													additionalPrice = replace(additionalPrice,"(","(-")
												end if

										 %>
										    <option value="<%=optvaluenode%>|<%=sloptpriceval%>"><%=optvaluenode %><%=additionalPrice%></option>
										    
										    <%next
										    set sl_optvalnode=nothing
										    %>
										</select>								
									<%end if%>
									<%if slopttype="R" then%>
									    <%
										    objDoc.loadxml(optvalxmlstring)
										    'target_node = "//loptval[optcode='"+cstr(sloptcode)+"']"
											target_node = "//loptval[optcode='"+cstr(sloptcode)+"' and opttitle='"+cstr(slopttitle)+"']"
										    set sl_optvalnode = objDoc.selectNodes(target_node)
										    for y= 0 to sl_optvalnode.length-1 
										        optvaluenode = sl_optvalnode.item(y).selectSingleNode("optvalue").text
												sloptprice   = sl_optvalnode.item(y).selectSingleNode("optprice").text
												sloptpriceval = sloptprice
												 additionalPrice= ""
												if sloptprice > 0 then
													sloptprice =formatcurrency(sloptprice)
													additionalPrice = "(+" + cstr(replace(sloptprice,".00","")) + ")"								  
												end if 
												if sloptprice < 0 then
													sloptprice =formatcurrency(sloptprice)
													additionalPrice = cstr(replace(sloptprice,".00","")) 
													additionalPrice = replace(additionalPrice,"(","(-")
												end if
										 %>
										<input type="radio" name="NonSkuAttribtext<%=(x+1)%>" id="NonSkuAttribtext<%=(x+1)%>" value="<%=optvaluenode%>|<%=sloptpriceval%>"><%=optvaluenode %>&nbsp;<%=additionalPrice %>
										<%next
								        set sl_optvalnode=nothing
								        %>
									<%end if%>
								</td>
								</table>
							</td>
							</tr>
							
							<%next
							
							set SL_optcodeNONSkue=nothing 
							%>
							<tr><td><img alt="" src="images/clear.gif" width="1" height="10" border="0"></td></tr>
							<tr><td class="plaintextbold">
    							<% if len(trim(current_variation)) =0 then %>
	    						    Price: Starting at  <%=formatcurrency(StockAttrib_MinPrice)%>
		    					    <%else %>
			    				    Price: <%=formatcurrency(LOSTOCK.price1)%>
			    				    <%end if %>
							</td></tr>
							<%end if  'if has attribs%>
							
							
							
							
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
							
							<tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0">
							<%if has_stock_attribs=true then %>
							    <input type="hidden" name="maxoption" id="MaxAttriboption" value="<%=maxAttriboption+1%>">
							    <input type="hidden" name="maxNONSkuOption" id="maxNONSkuOption" value="<%=maxTextoption+1%>">
							<%else %>
							    <input type="hidden" name="maxoption" id="MaxAttriboption" value="0">
							    <input type="hidden" name="maxNONSkuOption" id="maxNONSkuOption" value="0">
							<%end if %>
							
							</td></tr>
							<%if LOSTOCK.discont= true and LOSTOCK.units=0  then %>
								<tr><td class="plaintextbold" valign="bottom">
									This item is no longer available
									<br><br>
									<%if Has_Subs_ForDiscont=true then %>									
									<a href="<%=insecureurl%>itemadd.asp?ADDITEM=<%=current_stocknumber%>">Check for Substitute items</a>
									<%end if%>
								</td></tr>
							<%else%>
							<%if LOSTOCK.giftcert= true or IsGiftCard=true then %>
							<input type="hidden" class="plaintext" name="txtquanto" value="1" size="4" MAXLENGTH="5" align="middle">
							    <% if LOSTOCK.giftcert= true then %>
							        <input type="hidden" class="plaintext" name="is_giftcert" value="1">
							        <input type="hidden" class="plaintext" name="is_giftcard" value="0">
							    <%else %>
							        <input type="hidden" class="plaintext" name="is_giftcert" value="0">
							        <input type="hidden" class="plaintext" name="is_giftcard" value="1">
							    <%end if %>
							    
							<%else%>
								<tr><td class="plaintextbold" valign="bottom">Quantity&nbsp;
								<input type="text" class="plaintext" name="txtquanto" value="1" size="4" MAXLENGTH="5" align="middle">
								<input type="hidden" class="plaintext" name="is_giftcert" value="0">
								</td></tr>
								
							<%end if%>
							<tr><td><img alt="" src="images/clear.gif" width="1" height="10" border="0"></td></tr>
							<tr><td><input type="image" value="cart" onClick="displaythis(this)" src="images/btn-addtocart.gif" style="border:0;" align="middle">							
							<%'if WANT_WISHLIST = 1 and session("registeredshopper")= "YES" then 
							  if WANT_WISHLIST = 1 and  LOSTOCK.giftcert= false and IsGiftCard=false  then
							%>					
			                  <br><br>
							  <input type="image" value="wishlist" onClick="displaythis(this)" src="images/btn_Add_Wishlist.gif" style="border:0;"> 
						   <%end if%>
                           <!--#INCLUDE FILE = "text/Addthis.htm" -->
						   <%end if%>
						  
							</td></tr>
							<%if SHOW_BUYSAFE =1 then %>
							    <tr><td><span id="buySAFE_Kicker" name="buySAFE_Kicker" type="Kicker Guaranteed Ribbon 200x90"></span></td></tr>
							<%end if %>
							</table>
                            </div>
						</td>
					</tr>
					</table>
					
					
					
					
					
				</td>
              
	</tr>	
	</table>
	</form>
	
	<%
	    useadv1 = false
	    useadv2 = false
	    useadv3 = false
	    useadv4 = false
	    
	    SLAdvanced1Val = trim(LOSTOCK.advanced1)
	    SLAdvanced2Val = trim(LOSTOCK.advanced2)
	    SLAdvanced3Val = trim(LOSTOCK.advanced3)
	    SLAdvanced4Val = trim(LOSTOCK.advanced4)
	    
	    
	    if len(trim(session("SL_Advanced1"))) > 0 and len(trim(SLAdvanced1Val)) > 0 then	        
	        useadv1 = true
	    end if
	
	    if len(trim(session("SL_Advanced2"))) > 0 and len(trim(SLAdvanced2Val)) > 0 then	        
	        useadv2 = true
	    end if
	    if len(trim(session("SL_Advanced3"))) > 0 and len(trim(SLAdvanced3Val)) > 0 then	        
	        useadv3 = true
	    end if	
	    if len(trim(session("SL_Advanced4"))) > 0 and len(trim(SLAdvanced4Val)) > 0 then	        
	        useadv4 = true
	    end if	
	    
	
	 %>
	 
	 <%if useadv1=true or useadv2=true or useadv3=true or useadv4=true then %>
	    <table width="100%">
	        <tr>
	            <%if useadv1=true then %>
	                <td class="plaintext"><br><b><%=session("SL_Advanced1")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=1&amp;what=<%=server.urlencode(SLAdvanced1Val)%>"><%=SLAdvanced1Val%></a></td>
	            <%end if%>
	            <%if useadv2=true then %>
	                <td class="plaintext"><br><b><%=session("SL_Advanced2")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=2&amp;what=<%=server.urlencode(SLAdvanced2Val)%>"><%=SLAdvanced2Val%></a></td>
	            <%end if%>
	            <%if useadv3=true then %>
	                <td class="plaintext"><br><b><%=session("SL_Advanced3")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=3&amp;what=<%=server.urlencode(SLAdvanced3Val)%>"><%=SLAdvanced3Val%></a></td>
	            <%end if%>
	            <%if useadv4=true then %>
	                <td class="plaintext"><br><b><%=session("SL_Advanced4")%></b><a class="allpage" href="SetAdvancedSearch.asp?search=4&amp;what=<%=server.urlencode(SLAdvanced4Val)%>"><%=SLAdvanced4Val%></a></td>
	            <%end if%>
	        
	        </tr>
		</table>
	 <%end if %>
	
	 
	 <%  if len(trim(LOSTOCK.inetfdesc)) > 0 then %>
    <br>
	<table width="100%"  border="0" cellpadding="3" cellspacing="0">
			<tr><td class="THHeader">&nbsp;Detailed Description</td></tr>
            <tr><td><img alt="" src="images/clear.gif" width="1" height="5" border="0"></td></tr>
			<tr><td class="plaintext" style="text-align:justify;">			           	           
						<%=LOSTOCK.inetfdesc%>
						</td>
					</tr>
			</table>
	<%end if%>
	
	
	<%
	if SHOW_PRODUCTREVIEWS = 1 then
	if SL_cust_review.length > 0 then%>
		<a name="reviews"></a>
		<br />				
		<table width="100%" BORDER="0" cellpadding="3" cellspacing="0" ID="Table5">
		<tr><td class="THHeader" colspan="3">&nbsp;Product Reviews<img src="images/clear.gif" border="0" width="100" height="1" /><a class="allpage" href="<%=insecureurl%>writeReview.asp?number=<%=current_stocknumber%>">Click here to review this item</a></td></tr>	
        <tr><td><img alt="" src="images/clear.gif" width="1" height="10" border="0"></td></tr>
		<%
		for x =0 to SL_cust_review.length-1			
			SL_rating =  SL_cust_review.item(x).selectSingleNode("rating").text
			SL_RTitle =  SL_cust_review.item(x).selectSingleNode("title").text
			SL_RDesc =  SL_cust_review.item(x).selectSingleNode("desc").text
			SL_R_FirstName = SL_cust_review.item(x).selectSingleNode("firstname").text
			SL_R_LastName = SL_cust_review.item(x).selectSingleNode("lastname").text
			SL_R_State = SL_cust_review.item(x).selectSingleNode("state").text
			
			usethis_rating_img = "images/"&SL_rating&+"note.gif"

			footer_txt=""
			if len(trim(SL_R_FirstName)) > 0 then
				footer_txt = trim(SL_R_FirstName)
			end if
			
			if len(trim(SL_R_LastName)) > 0 then
				if len(trim(footer_txt)) > 0 then
					footer_txt = footer_txt & "&nbsp;"
				end if
				footer_txt = footer_txt & trim(SL_R_LastName)
			end if
			
			if len(trim(SL_R_State)) > 0 then
				if len(trim(footer_txt)) > 0 then
					footer_txt = "-&nbsp;" & footer_txt & ",&nbsp;"
				end if
				footer_txt = footer_txt & trim(SL_R_State)
			end if


		%>
		<tr><td width="100"><img src="<%=usethis_rating_img%>" border="0" width="70" height="14"></td>
			<td class="plaintext" width="300"><b><%=SL_RTitle%></b></td>
		</tr>
		<tr><td colspan="2" class="plaintext" style="padding-top:3px"><%=SL_RDesc%></td></tr>	
		<%if len(trim(footer_txt)) > 0 then%>
		<tr><td colspan="2" class="plaintext" style="font-size:10px"><em><%=footer_txt%></em></td></tr>
		<%end if%>
		<tr><td colspan="2" class="plaintext">&nbsp;</td></tr>
		<% next %>
		</table>
		<%
		end if 'if SL_cust_review.length = 0
		set SL_cust_review = nothing 
		
		end if 'if SHOW_PRODUCTREVIEWS
		%>		
		
		<%
    
        if SHOW_RECENTLYVIEWEDPRODUCTS=1 then

        if len(trim(LOSTOCK.inetthumb)) > 0 then            
            ProductThumbNail = "images/"+trim(LOSTOCK.inetthumb)            
            set objFileName = CreateObject("Scripting.FileSystemObject")
            StrPhysicalPath = Server.MapPath(ProductThumbNail)
            if objFileName.FileExists(StrPhysicalPath) then
                RecentlyViewProductThumbNail=ProductThumbNail
            else
                RecentlyViewProductThumbNail = ProductFullSizeImage            
            end if        
            set objFileName = nothing
        else
            RecentlyViewProductThumbNail = ProductFullSizeImage  
        end if		
        
   %>
	<table width="100%"  border="0" cellpadding="3" cellspacing="0">		
							<%if len(trim(session("item1viewednumber"))) > 0 or len(trim(session("item2viewednumber"))) > 0  or len(trim(session("item3viewednumber"))) > 0 or len(trim(session("item4viewednumber"))) > 0  then %>
									<tr><td colspan="5" style="height:5px;"></td></tr>									
								    <% nowviewitemstext = "Recently Viewed Items"%>								    
								    <tr><td class="THHeader" >&nbsp;Recently Viewed Items</td></tr>
								    <tr><td><img alt="" src="images/clear.gif" width="1" height="10" border="0"></td></tr>
									<tr><td colspan="2">
									<table width="0" border="0">									
									<tr>
									
									<%if (len(trim(session("item1viewednumber"))) > 0) and (current_stocknumber <> session("item1viewednumber")) then %>
									    <td valign="top" width="140">
                            		    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="height:130px;" class="table-layout-fixed">
									    <tr><td><a href="<%=session("item1viewedlink")%>"><img src="<%=session("item1viewedimgsrc")%>" border="0" class="prodlistimg" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_WIDTH > 0 and PROD_IMAGE_WIDTH > 125 then  response.write(" width=""125""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %>></a></td>
    									    <td style="width:20px;"></td>
									        </tr>
									    </table>
									    </td>
							        <%end if%>
							        
							        <%if (len(trim(session("item2viewednumber"))) > 0) and (current_stocknumber <> session("item2viewednumber")) then %>
							            <td valign="top" width="140">
                            		    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="height:130px;" class="table-layout-fixed">
									    <tr><td><a href="<%=session("item2viewedlink")%>"><img src="<%=session("item2viewedimgsrc")%>" border="0" class="prodlistimg" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_WIDTH > 0 and PROD_IMAGE_WIDTH > 125 then  response.write(" width=""125""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %>></a></td>						
                                            <td style="width:20px;"></td>
									        </tr>
									    </table>
									    </td>
							        <%end if%>
							        
							        <%if (len(trim(session("item3viewednumber"))) > 0) and (current_stocknumber <> session("item3viewednumber")) then %>
							            <td valign="top" width="140">
                            		    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="height:130px;" class="table-layout-fixed">
									    <tr><td><a href="<%=session("item3viewedlink")%>"><img src="<%=session("item3viewedimgsrc")%>" border="0" class="prodlistimg"  <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_WIDTH > 0 and PROD_IMAGE_WIDTH > 125 then  response.write(" width=""125""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %>></a></td>
    									    <td style="width:20px;"></td>
									        </tr>
									    </table>
									    </td>
							        <%end if%>
							        
							        <%if (len(trim(session("item4viewednumber"))) > 0) and (current_stocknumber <> session("item4viewednumber")) then %>
							            <td valign="top" width="140">
                            		    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="height:130px;" class="table-layout-fixed">
									    <tr><td><a href="<%=session("item4viewedlink")%>"><img src="<%=session("item4viewedimgsrc")%>" border="0" class="prodlistimg" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_WIDTH > 0 and PROD_IMAGE_WIDTH > 125 then  response.write(" width=""125""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %>></a></td>	
                                            <td style="width:20px;"></td>
									        </tr>
									    </table>
									    </td>
							        <%end if%>
							   
							       
							        </tr>
									
									</table>
							</td></tr>							
							<% end if %>
							
							<%
							    
									if len(trim(current_stocknumber)) > 0 then									  
										if current_stocknumber <> session("item1viewednumber") then										    
										  if current_stocknumber <> session("item2viewednumber") then										  
											if current_stocknumber <> session("item3viewednumber") then
											  if current_stocknumber <> session("item4viewednumber") then
											    if current_stocknumber <> session("item5viewednumber") then
											      'if current_stocknumber <> session("item6viewednumber") then											        
     						           			     'ViewStrFileName = RecentlyViewProductThumbNail
												     'ViewStrPhysicalPath = Server.MapPath(ViewStrFileName)
			  									     'set ViewobjFileName = CreateObject("Scripting.FileSystemObject")	
			  									     'if ViewobjFileName.FileExists(ViewStrPhysicalPath) then
													     'session("item6viewednumber") = session("item5viewednumber") 
													     session("item5viewednumber") = session("item4viewednumber") 
													     session("item4viewednumber") = session("item3viewednumber") 
													     session("item3viewednumber") = session("item2viewednumber") 
													     session("item2viewednumber") = session("item1viewednumber") 
													     session("item1viewednumber") = current_stocknumber
													     'session("item6viewedlink") = session("item5viewedlink")
													     session("item5viewedlink") = session("item4viewedlink")
													     session("item4viewedlink") = session("item3viewedlink")
													     session("item3viewedlink") = session("item2viewedlink")
													     session("item2viewedlink") = session("item1viewedlink")
													     if WANT_REWRITE=1 then
													        session("item1viewedlink") = insecureurl + usethis_for_urlwrite + "/productinfo/" + current_stocknumber+"/"
													     else
													        session("item1viewedlink") = insecureurl + "prodinfo.asp?number=" + current_stocknumber
													     end if
													     'session("item6viewedimgsrc") = session("item5viewedimgsrc") 
													     session("item5viewedimgsrc") = session("item4viewedimgsrc") 
													     session("item4viewedimgsrc") = session("item3viewedimgsrc") 
													     session("item3viewedimgsrc") = session("item2viewedimgsrc") 
													     session("item2viewedimgsrc") = session("item1viewedimgsrc") 
													     'session("item1viewedimgsrc") = insecureurl + RecentlyViewProductThumbNail
													     session("item1viewedimgsrc") = RecentlyViewProductThumbNail
									  			     'end if
											         'set ViewobjFileName = nothing
											      'end if
											    end if
											  end if
										    end if
										  end if
										end if
									end if
								%>
								
								
            </table>
    <%end if 'if SHOW_RECENTLYVIEWEDPRODUCTS %>
	
	</div>
<!-- end sl_code here -->
	<br><br><br><br>
	</td>
	<%
		objDoc.loadxml(SL_CrossSellXmlStream)
		
		set SL_SellTool = objDoc.selectNodes("//seltool") 
		if SL_SellTool.length > 0 then
	%>
	<td width="150" valign="top" class="pagenavbg">
	
	<table width="100%" cellpadding="3" cellspacing="0">
			<tr><td align="center"><h3>You may also like:</h3> </td></tr>			

		  <% 
		  all_prods=""
		  for x=0 to SL_SellTool.length-1
			SLProdNumber = SL_SellTool.item(x).selectSingleNode("sellwhatsku").text
			all_prods =all_prods+","+cstr(SLProdNumber)
		  next
		  
		  all_prods =all_prods+","
		  
		    SPxmlstring = sitelink.GET_ALL_SPECIALPRICES(cstr(all_prods),session("ordernumber"),session("shopperid"),cstr(""))
		    objDoc.loadxml(SPxmlstring)

			rowcount=0   
			CSPROD_PER_LINE = 1			    					
			for x=0 to SL_SellTool.length-1
				SLsellNumber = SL_SellTool.item(x).selectSingleNode("sellwhatsku").text
				SLsellPrice  = SL_SellTool.item(x).selectSingleNode("price1").text
				SLsellTitle  = SL_SellTool.item(x).selectSingleNode("inetsdesc").text
				SLsellThumb  = SL_SellTool.item(x).selectSingleNode("inetthumb").text
				SLsellFull   = SL_SellTool.item(x).selectSingleNode("inetimage").text
				
				if WANT_REWRITE = 1 then
					 SL_urltitle = SLsellTitle
					 SL_urltitle = url_cleanse(SL_urltitle)
					 SLURLCodeNumber = server.URLEncode(trim(SLsellNumber))
					 'prodlink = SL_urltitle + "/productinfo/" + SLURLCodeNumber 
					 prodlink = insecureurl + SL_urltitle + "/productinfo/" + SLURLCodeNumber + "/"
				else 
					 prodlink = "prodinfo.asp?number=" + cstr(SLsellNumber)
				end if
				
				
				target_in_SpecialPrice = "//gsp[number='"+cstr(SLsellNumber)+"']"
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
						unitprice = SLsellPrice
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
						

			if rowcount=CSPROD_PER_LINE then
			rowcount=0
			%>
			<tr><td></td></tr>		     					
			<%end if%>
			
			<%if rowcount=0 then %>
		        <tr>
		    <%end if %>
		  
		   <td valign="top" width="<%=(100/CSPROD_PER_LINE)%>%" class="cross-sell-bg">				
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
				   
					<td align="center"> 
					 <a href="<%=prodlink%>"><img SRC="<%=imagename%>" alt="<%=SLsellTitle%>" title="<%=SLsellTitle%>" class="cross-sell-img" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %> border="0"></a>
					</td>		                         
					</tr>
					<tr><td align="center">
					<a class="producttitlelink" href="<%=prodlink%>"><%=SLsellTitle%></a>
					</td></tr>
					<tr><td align="center" class="plaintext">
					<%if has_Onespecial_Price=true then %>
					    
					    
					    <span class="plaintext"><s><%=formatcurrency(SLsellPrice)%></s></span>
			            <br>
			            <table width="0" cellpadding="0" cellspacing="0" border="0">
			            <tr><td><img src="images/RedSaleTag.gif" style="border:0; width:53; height:23;" alt="sale"></td>
				            <td class="plaintext" align="center" style="background-image:url(images/RedSaleTag_bkg.gif);height:23">
				            <span style="font-weight:bold;color:White;"><%=formatcurrency(discountedUnitPrice)%></span>&nbsp;
				            </td>
			            </tr>
			            </table>
					
					<%else %>
					    <%=formatcurrency(SLsellPrice)%>
					<%end if %>
					</td></tr>
					<tr><td style="height:10px;"></td></tr>
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
	</td>
	<%end if%>
	
</tr>

</table>

<%     
	SET LOSTOCK = nothing
	set SL_sbcl = nothing
	set SL_gsp =  nothing
	set SL_SellTool=nothing
	set SL_ProdInfoDetails=nothing
%>

    </div> <!-- Closes main  -->
    <div id="footer" class="footerbgcolor">
    <!--#INCLUDE FILE = "include/bottomlinks.asp" -->
    <!--#INCLUDE FILE = "googletracking.asp" -->
    <!--#INCLUDE FILE = "RemoveXmlObject.asp" -->
    <!--#INCLUDE FILE = "text/footer.asp" -->
    </div>
</div> <!-- Closes container  -->

<script src="include/initialize_prettyPhoto.js" type="text/javascript" charset="utf-8"></script>
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