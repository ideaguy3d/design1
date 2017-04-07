<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!--#INCLUDE FILE = "include/momapp.asp" -->

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<%
xmlstring =sitelink.READ_METATAG(cstr("-3"),cstr(""))
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
<head>
<title>Aussie Products.com | SiteMap Page- Aussie Foods</title>
<meta name="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls" />
	<meta name="keywords" content="Aussie Products, Australian, Aussie, Australian Products, Australian Food, Aussie Foods, Australian Desserts, tim tams, vegemite, violet crumble, cherry ripe, koala, kangaroo, down under, Australian Groceries, foods, lollies, candy, gourmet food" />
	<meta name="robots" content="index, follow" />
    <meta name="robots" content="ALL">
	<meta name="distribution" content="Global">
	<meta name="revisit-after" content="5">
	<meta name="classification" content="Food">
 	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" /><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="canonical" href="<%=insecureurl%>sitemap.asp">
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
	
<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top"  >
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<center>
    <br>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%">Site Map</td></tr>
	</table>
	<br>
	
	<%
			xmlstring =sitelink.DEPARTMENTDIRECTORY(0,true)
			objDoc.loadxml(xmlstring)			
			set SL_Dept = objDoc.selectNodes("//thesedepts")
			
			objDoc.loadxml(xmlstring)
		%>
		<table  width="95%" cellpadding="2" cellspacing="0" border="0">
			  <% 			
			  
				for x=0 to SL_Dept.length-1
					 deptname = SL_Dept.item(x).selectSingleNode("name").text
					 deptcode = SL_Dept.item(x).selectSingleNode("deptcode").text
					 deptunder = SL_Dept.item(x).selectSingleNode("under").text
					 deptlevel = SL_Dept.item(x).selectSingleNode("nlevel").text
		 
 					 url_deptname = deptname
					 url_deptname = url_cleanse(url_deptname)

					 
					 hassubdept= false					 
					 target_subnode = "//thesedepts[under="+deptcode+"]"
					 set SL_deptsub = objDoc.selectNodes(target_subnode) 
					 	if SL_deptsub.length > 0 then
							hassubdept=true
							if WANT_REWRITE = 1 then
							uselink = insecureurl  + url_deptname + "/departments/" + cstr(deptcode) + "/"
							else 
							uselink = insecureurl + "departments.asp?dept=" + cstr(deptcode)
							end if
						else 
							if want_rewrite = 1 then
							uselink = insecureurl + url_deptname + "/products/" + cstr(deptcode) + "/" 
							else 
							uselink = insecureurl + "products.asp?dept=" + cstr(deptcode)
							end if
						end if
					 set SL_deptsub= nothing					 
					 
					 
					 altdeptname = server.URLEncode(deptname)
    
			  %>
				<tr>
					<td>
					<% if deptunder = 0 then %>
						<a title="<%=altdeptname%>" class="allpage" href="<%=uselink%>"><b><%=deptname%></b></a>
					<%else%>
						<img alt="" src="images/clear.gif" width="<%=(deptlevel * 15)%>" height="1" border="0"><a title="<%=altdeptname%>" class="allpage" href="<%=uselink%>"><%=deptname%></a>
					<%end if %>					
					</td>
				</tr>
			<%next%>
									
				</table>

	<%
			set SL_Dept = nothing
			'Set objDoc =nothing
	%>
	
	
	
	</center>
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
