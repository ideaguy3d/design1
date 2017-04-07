<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%

 
 isvalidstock = sitelink.validstocknumber2(cstr(session("Parentsku")))
 
 if len(trim(session("Parentsku"))) = 0 then
 	isvalidstock=-1
 end if
 
 if isvalidstock=-1 then
		set sitelink=nothing
		set ObjDoc=nothing
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl&"default.asp"
		Response.End
	end if

%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Product Offerings @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
<script type="text/javascript" language="javascript">
function displaythis(what)
{
	var baccepted = document.getElementById('UserAccept') ;
	var brejected  = document.getElementById('UserReject') ;
	
	if (what.value=="Accept")
	{
		 baccepted.value="Accept" ;
		 brejected.value="" ; 
	}
	else
	{
	 	baccepted.value="" ;
		brejected.value="Reject" ;
	}
	
}

function Process()
{
	var acceptval = document.getElementById('UserAccept').value ;
	
	if (acceptval=="Accept")
	{	
	 notallradioselected=true ;
	  j = UpsellInfoForm.OptionalSku.length?UpsellInfoForm.OptionalSku.length:0 ;
	 //j = document.getElementById('OptionalSku').length ;
	 //alert(j);
	 if (j == 0 ) 
	 {
	 	if (document.getElementById('OptionalSku').checked)
			{
				notallradioselected=false ;
			}
	 }
	 else
	 {
	  for (k=0; k<j; k++)
	  {
	  	radiochecked = UpsellInfoForm.OptionalSku[k].checked ;
		//alert(radiochecked);
		 if (radiochecked)
		 {
		 	notallradioselected=false ;
		 }		
	  }
	  }
	  
	  if (notallradioselected==true)
	  {
	  	alert("Please select one option") ;
		return false ;	  
	  }	 
	}
	
	return true ;
}
</script>
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">



<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
	<%if WANT_SIDENAV=1 then%>
	<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
	<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	
	<%end if%>
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<center>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0"  >
		<tr><td class="TopNavRow2Text" width="100%" >&nbsp;</td></tr>
	</table>
	<br>
	<br>
	<%
		extrafield=""
		xmlstring = sitelink.Get_From_SellTool("GETALL",cstr(session("Parentsku")),cstr(session("ParentskuVariant")),extrafield)
		objDoc.loadxml(xmlstring)
		
		
		set SL_selltool = objDoc.selectNodes("//seltool")
		
		if SL_selltool.length > 0 then
		  selltype =  SL_selltool.item(0).selectSingleNode("sell_type").text
		  sellmand = SL_selltool.item(0).selectSingleNode("sell_mand").text
		
		
		
	%>
	<%if selltype="U" then%>	
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<!-- <tr><td class="plaintextbold">This item <%=session("Parentsku")%> is Upsell to </td></tr> -->
		<tr><td class="plaintextbold">This item <%=session("Parentsku")%> has alternate product recommendation, may we suggest the following:&nbsp;</td></tr>
	</table>	
	<%end if%>

	<%if selltype="S" then%>	
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<!-- <tr><td class="plaintextbold">This item <%=session("Parentsku")%> is Substituted with</td></tr> -->
		<tr><td class="plaintextbold">This item <%=session("Parentsku")%> is no longer available, may we recommend the following substitutes:&nbsp;</td></tr>
	</table>	
	<%end if%>

	<form action="AddUpsell.asp" name="UpsellInfoForm" onSubmit="return Process()" method="post">
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
	<tr>
		<td valign="top">
		
		<table width="100%"  cellpadding="3" cellspacing="0" border="0">			
			<%for x= 0 to SL_selltool.length-1 
				SLProdNumber =  SL_selltool.item(x).selectSingleNode("sell_what").text
				SLProductTitle = SL_selltool.item(x).selectSingleNode("inetsdesc").text
				SL_thumbnail  = SL_selltool.item(x).selectSingleNode("inetthumb").text
				SL_Fullimage  = SL_selltool.item(x).selectSingleNode("inetimage").text
				SL_price  	  = SL_selltool.item(x).selectSingleNode("price1").text
				SL_Units      = SL_selltool.item(x).selectSingleNode("units").text
				SL_Drop  	  = SL_selltool.item(x).selectSingleNode("dropship").text
				SL_Contruct   = SL_selltool.item(x).selectSingleNode("construct").text
				
				if SL_thumbnail="" then
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
				
				if WANT_REWRITE = 1 then
				 SL_urltitle = SLProductTitle
				 SL_urltitle = url_cleanse(SL_urltitle)
				 SLURLCodeNumber = server.URLEncode(trim(SLProdNumber))
				 prodlink = insecureurl  + SL_urltitle + "/productinfo/" + SLURLCodeNumber + "/" 
				else 
					 prodlink = "prodinfo.asp?number=" + SLProdNumber
				end if
				
			%>
			<tr><td class="plaintextbold">
			
			<table width="100%"  cellpadding="3" cellspacing="0" border="0"
			style="BORDER-RIGHT: #CCCCCC 1px solid  ; 
	    	            BORDER-TOP: #CCCCCC 1px solid  ; 
			            BORDER-LEFT: #CCCCCC 1px solid  ; 
			            BORDER-BOTTOM: #CCCCCC 1px solid ">
			
			<tr>
				<td  width="30"><input type="radio" name="OptionalSku" id="OptionalSku" value="<%=SLProdNumber%>" <% if x=0 and SL_selltool.length=1 then response.write " checked" end if %>></td>
				<td valign="top" width="140"><div class="prodthumb"><div class="prodthumbcell"><a href="<%=prodlink%>"><img SRC="<%=imagename%>" alt="<%=SLProductTitle%>" title="<%=SLProductTitle%>" class="prodlistimg" <% if PROD_IMAGE_WIDTH > 0 then  response.write(" width="""& PROD_IMAGE_WIDTH&"""") end if %> <%if PROD_IMAGE_HEIGHT > 0 then response.write(" height="""& PROD_IMAGE_HEIGHT&"""") end if %> border="0"></a></div></div></td>
				<td valign="top" align="left">
					<table width="100%"  cellpadding="1" cellspacing="0" border="0">
						<tr><td class="plaintextbold">Item Number:&nbsp;<%=SLProdNumber%></td></tr>
						<tr><td class="plaintext"><a title="<%=SLProductTitle%>" class="allpage" href="<%=prodlink%>"><%=SLProductTitle%></a></td></tr>
						<%if SHOW_DUE_DATE = 0 and SHOW_IN_STOCK =1 then 	
							if (SL_Units > 0) or (SL_Drop = "true") or (SL_Contruct = "true") then
								avail_status="In Stock"				
							else
								avail_status="Out of Stock"		
							end if
							%>
							<tr><td class="plaintextbold"><%=avail_status%></td></tr>
						<%end if%>
					</table>
				</td>				
			</tr>
			
			</table>			
			
			</td></tr>
			<%next%>
			</table>
		
		</td>
	</tr>
	
	
	</table>
	<br><br>
	<table width="0"  cellpadding="3" cellspacing="0" border="0">
	<tr><td colspan="2" class="plaintextbold" align="center">
		<!-- Do you want to accept it? -->
		&nbsp;
		Please select any of the options and click "Accept" if any
		<br><br>
		<input type="hidden" name="UserAccept" id="UserAccept">
		<input type="hidden" name="UserReject" id="UserReject">
		<input type="hidden" name="TypeofSell" value="<%=selltype%>" id="<%=TypeofSell%>">
		<input type="hidden" name="sellmand" value="<%=sellmand%>" id="<%=sellmand%>">
		
	</td></tr>
	<tr><td align="center"><input type="image" value="Accept" onClick="displaythis(this)" src="images/btn-accept.gif" style="border:0;" align="middle"></td>
		<td align="center"><input type="image" value="Reject" onClick="displaythis(this)" src="images/btn-decline.gif" style="border:0;" align="middle"></td>
	</tr>
	</table>
	</form>
	<%else%>
	
	
	<table width="0"  cellpadding="3" cellspacing="0" border="0">
	<tr><td colspan="2" class="plaintextbold" align="center">

		&nbsp;
		There are no Substitutes available at this time.
		<br><br>
		
	</td></tr>
	
	</table>
		
	<% 

	end if
	set SL_selltool=nothing %>
	<br>
	


	</center>
	</div>
<!-- end sl_code here -->
	<img alt="" src="images/clear.gif" width="1" height="160" border="0">
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
