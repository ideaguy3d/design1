<%on error resume next%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">



<%
session("department")=cstr(Request.QueryString("dept"))

'response.write(len(trim(session("department"))))
if len(trim(session("department"))) = 0 then
	session("department") = 0 
end if

if isnumeric(session("department"))=false then	
	Response.Status="301 Moved Permanantly"
	Response.AddHeader "Location", insecureurl&"default.asp"
	Response.End
end if

if session("department") < 0  then	
	Response.Status="301 Moved Permanantly"
	Response.AddHeader "Location", insecureurl&"default.asp"
	Response.End
end if
%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<%

    IsValidUrl = true
    CorrectProductUrl = ""
    

    UrlWriteStr = request.ServerVariables("http_x_rewrite_url")
    IsValidUrl = sitelink.ValidateProdInfoUrl(session("department"),UrlWriteStr,CorrectProductUrl)        

         
    if IsValidUrl=false then    
        set sitelink=nothing
		set ObjDoc=nothing
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl & CorrectProductUrl
		Response.End
    end if
    
    if  session("department") > 0 then
		if sitelink.hassubdepts(session("department"))=0 then
			set sitelink=nothing 
			'Response.Redirect("products.asp?dept="+cstr(session("department")))		
			Response.Status="301 Moved Permanantly"
			Response.AddHeader "Location", insecureurl&"products.asp?dept="&session("department")
			Response.End
		end if	
end if
%>

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<%
set SL_DepartmentDetails = New cDepartmentObj
	SL_DepartmentDetails.ldeptcode=session("department")
	call sitelink.GET_SLDEPT_DETAILS(SL_DepartmentDetails)
	desc = SL_DepartmentDetails.ldeptDesc
	dept_title= SL_DepartmentDetails.ldeptTitle
	Sl_DeptMetaXmlStream = SL_DepartmentDetails.ldeptMetagXmlStream

set SL_DepartmentDetails = nothing


if WANT_REWRITE=1 then
	URL_dept_title = dept_title
	url_dept_title = url_cleanse(URL_dept_title)
end if



objDoc.loadxml(Sl_DeptMetaXmlStream)
set SL_Dept_Meta = objDoc.selectNodes("//results")
if SL_Dept_Meta.length > 0 then
	metatagtitle = trim(SL_Dept_Meta.item(0).selectSingleNode("ptitle").text)
	metatagdesc = trim(SL_Dept_Meta.item(0).selectSingleNode("pdesc").text)
	metatagkeywrd = trim(SL_Dept_Meta.item(0).selectSingleNode("pkeywrd").text)

	usethismeta="<title>"+ cstr(metatagtitle) +"</title>"
	usethismeta =usethismeta + VbCrLf
	usethismeta =usethismeta + "<META NAME=""KEYWORDS"" CONTENT="""+ cstr(metatagkeywrd) + """>"
	usethismeta =usethismeta + VbCrLf
	usethismeta =usethismeta + "<META NAME=""DESCRIPTION"" CONTENT="""+ cstr(metatagdesc) + """>"
else
	usethismeta = "<title>"+ cstr(dept_title)+ "-" + cstr(althomepage) +"</title>"
end if
set SL_Dept_Meta = nothing

if WANT_REWRITE=1 then
	dptlink= insecureurl & URL_dept_title & "/departments/" & session("department") & "/" 
else
	dptlink = insecureurl & "departments.asp?dept=" & session("department")
end if
%>
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
	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
<link rel="canonical" href="<%=dptlink%>">
<%if REQUEST.servervariables("server_port_secure")<> 1 then%>
<base href = "<%=insecureurl%>">
<%else%>
<base href = "<%=secureurl%>">
<%end if%>
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
<% session("destpage")="departments.asp?dept="+cstr(session("department")) %>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
	<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top"  >
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>
	<td valign="top" class="pagenavbg">

	<!-- sl code goes here -->
<div id="page-content">
<br>
	<% if session("department") > 0 then %>
	<table width="100%"  border="0" cellpadding="3" cellspacing="0">
	<tr><td class="breadcrumbrow">
				<a href="default.asp" class="breadcrumb">Home</a>&nbsp;&nbsp;&#47;&nbsp;&nbsp;
				<%
				    deptstr = sitelink.deptstring(session("department"),true,"&nbsp;&nbsp;&#47;&nbsp;&nbsp;","breadcrumb","",WANT_REWRITE)
				    deptstr = replace(deptstr,"https://","http://")
				 %>
				<%=deptstr%>
		</td></tr>	
	</table>
	<%end if%>
	<center>
	<% if len(trim(desc)) > 0 then %>
	<table width="100%"  border="0">
		<tr><td>
			<p align="left"><span class="plaintext"><%=desc%></span></p>
		</td></tr>
	 </table>
	<%end if%>
	<br>
	
	 <%
	
	xmlstring =sitelink.DEPARTMENTSNEW(cint(session("department")),true)
	if session("department") > 0 then
		mainnode ="//thesedepts[under="+session("department")+"]"
	else
		mainnode ="//thesedepts[under=0]"
	end if

	objDoc.loadxml(xmlstring)
	
	set SL_Dept = objDoc.selectNodes(mainnode)
	
	objDoc.loadxml(xmlstring)
	%>
	
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">

		  <% 

			rowcount=0   				    					
			for x=0 to SL_Dept.length-1
			 deptname = SL_Dept.item(x).selectSingleNode("name").text
			 deptcode = SL_Dept.item(x).selectSingleNode("deptcode").text
			 deptimage = SL_Dept.item(x).selectSingleNode("deptimage").text
			 
			 if WANT_REWRITE=1 then
	 			 url_deptname = trim(deptname)
				 url_deptname = url_cleanse(url_deptname)
			 end if

			 
			 hassubdept= false					 
			 target_subnode = "//thesedepts[under="+deptcode+"]"
			 set SL_deptsub = objDoc.selectNodes(target_subnode) 
				if SL_deptsub.length > 0 then
					hassubdept=true
				end if
			 set SL_deptsub= nothing
			 
			 if hassubdept=true then
				if WANT_REWRITE=1 then
					uselink= url_deptname + "/departments/" + deptcode + "/"
				else
					uselink = "departments.asp?dept=" + deptcode
				end if	
			 else
				if WANT_REWRITE=1 then
					uselink= url_deptname + "/products/" + deptcode + "/" 
				else
					uselink = "products.asp?dept=" + deptcode
				end if
			 end if
			 uselink=insecureurl+uselink
						

			if rowcount=NUM_OF_DEPT_PER_LINE then
			rowcount=0
			%>
			<tr><td></td></tr>		     					
			<%end if%>
			
			<%if rowcount=0 then %>
		        <tr>
		    <%end if %>
		   <% rowcount=rowcount+1%>
		   <td valign="top" align="center" width="<%=(100/NUM_OF_DEPT_PER_LINE)%>%" > 				
		   <table  width="100%" border="0" cellspacing="0" cellpadding="0" >
			<tr>      
									
				 <% 
					 StrFileName = "images/"+deptimage
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
					  <a href="<%=uselink%>">
							 <img src="<%=imagename%>" alt="<%=deptname%>" title="<%=deptname%>" <% if DEPT_IMAGE_WIDTH > 0 then response.write(" width="""& DEPT_IMAGE_WIDTH&"""") end if%> <%if DEPT_IMAGE_HEIGHT > 0 then response.write(" height="""& DEPT_IMAGE_HEIGHT&"""") end if %>	border="0" >
							</a> 
					</td>		                         
					</tr>
					<tr><td align="center"><a title="<%=deptname%>" class="allpage" href="<%=uselink%>"><%=deptname%></a></td></tr>
					<tr><td>&nbsp;</td></tr>
		</table>
		</td>											
		<%next%>
		</tr>
				  
	</table>
	
	<% 
    set SL_Dept = nothing
	'Set objDoc= nothing 
	%>
	</center>
	<br><br><br><br>
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