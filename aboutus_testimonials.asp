<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!--#INCLUDE FILE = "include/momapp.asp" -->

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<%
xmlstring =sitelink.READ_METATAG(cstr("-2"),cstr(""))
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
	usethismeta = "<title>"+ cstr(dept_title)+ "-" + cstr(althomepage) +"</title>"
end if
set SL_Dept_Meta = nothing
%>





<html>
<title>Aussie Products.com | About US Page- Aussie Foods</title>
<meta name="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls" />
	<meta name="keywords" content="Aussie Products, Australian, Aussie, Australian Products, Australian Food, Aussie Foods, Australian Desserts, tim tams, vegemite, violet crumble, cherry ripe, koala, kangaroo, down under, Australian Groceries, foods, lollies, candy, gourmet food" />
	<meta name="robots" content="index, follow" />
    <meta name="robots" content="ALL">
	<meta name="distribution" content="Global">
	<meta name="revisit-after" content="5">
	<meta name="classification" content="Food">
 	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />   
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

<% session("destpage")="" %>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
		<!--#INCLUDE FILE = "include/side_nav.asp" -->
		</td>	
		<td valign="top" class="pagenavbg">
        
        <!-- sl code goes here -->
        <div id="page-content" class="plaintext">	
        <!--#INCLUDE FILE = "text/AboutUs5.htm" -->
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
