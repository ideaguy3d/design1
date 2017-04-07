<%on error resume next%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<%
	current_stocknumber = cstr(REQUEST.QUERYSTRING("number"))
    current_stocknumber =fix_xss_Chars(current_stocknumber)
	
	
	SET LOSTOCK =  sitelink.setupproduct(current_stocknumber,cstr(current_variation),0,session("ordernumber"),cstr(""))
	
	showimageThumb = false
	showimageFull = false
	showimage1 = false
	showimage2 = false
	showimage3 = false
	showimage4 = false
	showimage5 = false
	showimage6 = false
	showimage7 = false
	showimage8 = false


	sl_image1 = trim(LOSTOCK.sl_image1)
	StrFileName = "images/"+cstr(sl_image1)
	StrPhysicalPath = Server.MapPath(StrFileName)
	set objFileName = CreateObject("Scripting.FileSystemObject")	
	if objFileName.FileExists(StrPhysicalPath) then
   		showimage1=true
	end if
	set objFileName = nothing


	sl_image2 = trim(LOSTOCK.sl_image2)
	StrFileName = "images/"+cstr(sl_image2)
	StrPhysicalPath = Server.MapPath(StrFileName)
	set objFileName = CreateObject("Scripting.FileSystemObject")	
	if objFileName.FileExists(StrPhysicalPath) then
   		showimage2=true
	end if
	set objFileName = nothing
	
	sl_image3 = trim(LOSTOCK.sl_image3)
	StrFileName = "images/"+cstr(sl_image3)
	StrPhysicalPath = Server.MapPath(StrFileName)
	set objFileName = CreateObject("Scripting.FileSystemObject")	
	if objFileName.FileExists(StrPhysicalPath) then
   		showimage3=true
	end if
	set objFileName = nothing
	

	sl_image4 = trim(LOSTOCK.sl_image4)
	StrFileName = "images/"+cstr(sl_image4)
	StrPhysicalPath = Server.MapPath(StrFileName)
	set objFileName = CreateObject("Scripting.FileSystemObject")	
	if objFileName.FileExists(StrPhysicalPath) then
   		showimage4=true
	end if
	set objFileName = nothing
	
	sl_image4 = trim(LOSTOCK.sl_image4)
	StrFileName = "images/"+cstr(sl_image4)
	StrPhysicalPath = Server.MapPath(StrFileName)
	set objFileName = CreateObject("Scripting.FileSystemObject")	
	if objFileName.FileExists(StrPhysicalPath) then
   		showimage4=true
	end if
	set objFileName = nothing
	

	sl_image5 = trim(LOSTOCK.sl_image5)
	StrFileName = "images/"+cstr(sl_image5)
	StrPhysicalPath = Server.MapPath(StrFileName)
	set objFileName = CreateObject("Scripting.FileSystemObject")	
	if objFileName.FileExists(StrPhysicalPath) then
   		showimage5=true
	end if
	set objFileName = nothing

	sl_image6 = trim(LOSTOCK.sl_image6)
	StrFileName = "images/"+cstr(sl_image6)
	StrPhysicalPath = Server.MapPath(StrFileName)
	set objFileName = CreateObject("Scripting.FileSystemObject")	
	if objFileName.FileExists(StrPhysicalPath) then
   		showimage6=true
	end if
	set objFileName = nothing
	
	sl_image7 = trim(LOSTOCK.sl_image7)
	StrFileName = "images/"+cstr(sl_image7)
	StrPhysicalPath = Server.MapPath(StrFileName)
	set objFileName = CreateObject("Scripting.FileSystemObject")	
	if objFileName.FileExists(StrPhysicalPath) then
   		showimage7=true
	end if
	set objFileName = nothing
	
	sl_image8 = trim(LOSTOCK.sl_image8)
	StrFileName = "images/"+cstr(sl_image8)
	StrPhysicalPath = Server.MapPath(StrFileName)
	set objFileName = CreateObject("Scripting.FileSystemObject")	
	if objFileName.FileExists(StrPhysicalPath) then
   		showimage8=true
	end if
	set objFileName = nothing
	
	
	
	showfirst="images/noimage.gif"
	if showimage8=true then
		showfirst="images/"+cstr(sl_image8)
	end if
	if showimage7=true then
		showfirst="images/"+cstr(sl_image7)
	end if
	if showimage6=true then
		showfirst="images/"+cstr(sl_image6)
	end if
	if showimage5=true then
		showfirst="images/"+cstr(sl_image5)
	end if

	if showimage4=true then
		showfirst="images/"+cstr(sl_image4)
	end if
	if showimage3=true then
		showfirst="images/"+cstr(sl_image3)
	end if

	if showimage2=true then
		showfirst="images/"+cstr(sl_image2)
	end if

	if showimage1=true then
		showfirst="images/"+cstr(sl_image1)
	end if
	
	SET LOSTOCK =nothing
	

%>
<html>
<head>

<title>Aussie Products.com | Addtional Product Images @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
<script type="text/javascript" language="javascript">
function setimage(imgname)
{
	document.Sitelink.src="images/"+imgname ;
	//alert(imgname)
}

</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">

<table width="100" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <th scope="row"><table border="0">
      <tr>
        <td height="500" rowspan="2" valign="top" scope="row"><%if showimage1=true then%>
            <a href="#" onClick="Javascript:setimage('<%=sl_image1%>')"><img src="images/<%=sl_image1%>" alt="<%=sl_image1%>" width="70" height="70" border="0" style="BORDER-RIGHT: #000000 1px solid  ; 
	    	BORDER-TOP: #000000 1px solid  ; 
			BORDER-LEFT: #000000 1px solid  ; 
			BORDER-BOTTOM: #000000 1px solid "></a>
            <%end if%>
            <p>
              <%if showimage2=true then%>
              <a href="#" onClick="Javascript:setimage('<%=sl_image2%>')"><img src="images/<%=sl_image2%>" alt="<%=sl_image2%>" width="70" height="70" border="0" style="BORDER-RIGHT: #000000 1px solid  ; 
	    	BORDER-TOP: #000000 1px solid  ; 
			BORDER-LEFT: #000000 1px solid  ; 
			BORDER-BOTTOM: #000000 1px solid "></a>
              <%end if%>
            <p>
              <%if showimage3=true then%>
              <a href="#" onClick="Javascript:setimage('<%=sl_image3%>')"><img src="images/<%=sl_image3%>" alt="<%=sl_image3%>" width="70" height="70" border="0" style="BORDER-RIGHT: #000000 1px solid  ; 
	    	BORDER-TOP: #000000 1px solid  ; 
			BORDER-LEFT: #000000 1px solid  ; 
			BORDER-BOTTOM: #000000 1px solid "></a>
            <p>
              <%end if%>
              <%if showimage4=true then%>
              <a href="#" onClick="Javascript:setimage('<%=sl_image4%>')"><img src="images/<%=sl_image4%>" alt="<%=sl_image4%>" width="70" height="70" border="0" style="BORDER-RIGHT: #000000 1px solid  ; 
	    	BORDER-TOP: #000000 1px solid  ; 
			BORDER-LEFT: #000000 1px solid  ; 
			BORDER-BOTTOM: #000000 1px solid "></a>
            <p>
              <%end if%>
            </td>
        <td valign="top"><%if showimage5=true then%>
            <a href="#" onClick="Javascript:setimage('<%=sl_image5%>')"><img src="images/<%=sl_image5%>" alt="<%=sl_image5%>" width="70" height="70" border="0" align="top" style="BORDER-RIGHT: #000000 1px solid  ; 
	    	BORDER-TOP: #000000 1px solid  ; 
			BORDER-LEFT: #000000 1px solid  ; 
			BORDER-BOTTOM: #000000 1px solid "></a>
          <p>
              <%end if%>
              <%if showimage6=true then%>
              <a href="#" onClick="Javascript:setimage('<%=sl_image6%>')"><img src="images/<%=sl_image6%>" alt="<%=sl_image6%>" width="70" height="70" border="0" align="top" style="BORDER-RIGHT: #000000 1px solid  ; 
	    	BORDER-TOP: #000000 1px solid  ; 
			BORDER-LEFT: #000000 1px solid  ; 
			BORDER-BOTTOM: #000000 1px solid "></a>
          <p>
              <%end if%>
              <%if showimage7=true then%>
              <a href="#" onClick="Javascript:setimage('<%=sl_image7%>')"><img src="images/<%=sl_image7%>" alt="<%=sl_image7%>" width="70" height="70" border="0" align="top" style="BORDER-RIGHT: #000000 1px solid  ; 
	    	BORDER-TOP: #000000 1px solid  ; 
			BORDER-LEFT: #000000 1px solid  ; 
			BORDER-BOTTOM: #000000 1px solid "></a>
          <p>
              <%end if%>
              <%if showimage8=true then%>
              <a href="#" onClick="Javascript:setimage('<%=sl_image8%>')"><img src="images/<%=sl_image8%>" alt="<%=sl_image8%>" width="70" height="70" border="0" align="top" style="BORDER-RIGHT: #000000 1px solid  ; 
	    	BORDER-TOP: #000000 1px solid  ; 
			BORDER-LEFT: #000000 1px solid  ; 
			BORDER-BOTTOM: #000000 1px solid "></a>
          <p>
              <%end if%>
            </td>
        <td height="500" rowspan="2"><img src="<%=showfirst%>" alt="<%=showfirst%>" name=Sitelink border="0" align="top" ></td>
      </tr>
    </table></th>
  </tr>
  <tr>
    <th scope="row" height=1><table width="100%">
<tr>
      <td height="2" align="left" > 
      
         <a href="#" onClick="JavaScript:window.close()"><font face="Arial" size="2" color="#CC0000">Close Window</font></a>              </td>
  </tr>
</table></th>
  </tr>
</table>

<!--#INCLUDE FILE = "RemoveXmlObject.asp" -->
</body>
</html>

