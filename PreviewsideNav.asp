<%on error resume next%>
<!--#INCLUDE FILE = "include/AdminCreateSL.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<html>
<head>
<title><%=althomepage%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->

<center>

<br><br>

<table width="600" border="0" cellpadding="0" cellspacing="0">
<tr><td valign="top">
<table width="0" class="SideNavbordercolor" cellpadding="0" cellspacing="0" border="0">
			 <tr>
			 	 <td align="center" width="<%=SIDE_NAV_WIDTH%>" class="sidenavTxt" height="29">DEPARTMENTS</td>			 	 
			 </tr>
			 
			 <tr><td valign="top" >
			 		<table width="100%" cellpadding="0" cellspacing="0" border="0">	
			 			<tr>
							
							<td valign="top" width="5"><img src="images/clear.gif" alt="" border="0" width="5" height="1"></td>
						    <td width="100%" valign="top">
							 <table  width="100%" cellpadding="3" cellspacing="0" border="0">				
							
								<%					
									xmlstring =sitelink.DEPARTMENTS(0,true)
			
									objDoc.loadxml(xmlstring)
									If objDoc.parseError.errorCode <> 0 Then 
									  response.write("<tr><td class=""plaintext""><br>Error Processing<br><br></td></tr>")
								   else		
										set SL_Dept = objDoc.selectNodes("//thesedepts[under=0]")				
										objDoc.loadxml(xmlstring)			   
								  %>
								<% if session("hasspecials") = 1 then 
									
								%>
								<tr>
									<td><a title="Specials" class="sidenav2" href="#">Specials</a></td>
								</tr>
								<%end if%>
								<% 			 
							 
								 for x=0 to SL_Dept.length-1
									 deptname = SL_Dept.item(x).selectSingleNode("name").text
									 deptcode = SL_Dept.item(x).selectSingleNode("deptcode").text
							  %>								
									<tr><td><a title="<%=deptname%>" class="sidenav2" href="#"><%=deptname%></a></td></tr>												
							<%next%>
			
								<%
											   
								 end if
								%>
								
								</table>
										</td>
										
								</table>
						 
							</td>
						</tr>
			
			</table>			
			
			<table width="0" cellpadding="0" cellspacing="0" border="0">
			<tr><td height="10"><img  alt="" src="images/clear.gif" width="1" height="1" border="0"></td></tr>
			<tr><td height="1" class="sidenavTxt"><img  alt="" src="images/clear.gif" width="1" height="1" border="0"></td></tr>
			<tr><td height="10"><img  alt="" src="images/clear.gif" width="1" height="1" border="0"></td></tr>
			</table>

			<table width="0" class="SideNavbordercolor" cellpadding="0" cellspacing="0" border="0">
			 <tr>
			 	 <td width="<%=SIDE_NAV_WIDTH%>" align="center" class="sidenavTxt" height="29">ACCOUNT INFO</td>			 	 
			 </tr>
			 
			 <tr><td valign="top">
			 		<table width="100%" cellpadding="0" cellspacing="0" border="0">	
			 			<tr>							
							<td valign="top" width="5"><img src="images/clear.gif" alt="" border="0" width="5" height="1"></td>
						    <td width="100%" valign="top">
							<table width="100%" cellpadding="3" cellspacing="0" border="0">					
								<tr><td><a class="sidenav2" title="View Basket" href="#">VIEW BASKET</a></td></tr>
								<tr><td><a class="sidenav2" title="Account Info" href="#">ACCOUNT INFO</a></td></tr>
								<tr><td><a class="sidenav2" title="Request Catalog" href="#">REQUEST CATALOG</a></td></tr>
								<tr><td><a class="sidenav2" title="Address Book" href="#">ADDRESS BOOK</a></td></tr>
								</table>
							</td>							
			 	    </table>
			 
			 	</td>
			</tr>
			
			</table>
	</td>
	
	<td valign="top">
<table width="100%" cellpadding="10">
		<tr><td>
		<font color="black" face="Verdana" size="2">
		<b>Side Nav Text </b>
			<ul>
				DEPARTMENTS <br>
				ACCOUNT INFO <br>
			
			</ul>
		</font>
		
			</td>
		</tr>
		<tr><td>
		<font color="black" face="Verdana" size="2">
		<b>Side Nav Links </b>
			<ul>
				All Other Links <br>
				<br><br>
			
			</ul>
		</font>
		
			</td>
		</tr>
		
		</table>
	</td>
	</tr>
</table>	
			
			


<br><br>
<br><br>
<a href="Javascript:window.close()"><span class="CompPrice">Close Window</span></a>

</center>





<!--#INCLUDE FILE = "include/AdminReleaseSL.asp" -->
