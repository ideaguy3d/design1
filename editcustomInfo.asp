<%on error resume next%>


<%
	item_id = request.querystring("item_id")
		
	item_id = item_id + 1 -1 
	
	quanto = request.querystring("quant")
	quanto = quanto + 1 - 1
	
	basketrecord = request.querystring("basketrecord")
	
	if isnumeric(item_id)=false or isnumeric(quanto)=false or isnumeric(basketrecord)=false then	
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl&"default.asp"
		Response.End
	end if
	
	
	
	
%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<%

	current_stocknumber = cstr(REQUEST.QUERYSTRING("item"))
    current_stocknumber =fix_xss_Chars(current_stocknumber)
    
    current_variation = cstr(REQUEST.QUERYSTRING("variant"))
    current_variation =fix_xss_Chars(current_variation)
	
	isvalidstock = sitelink.validstocknumber2(current_stocknumber)
	
	if isvalidstock=-1 then
		set sitelink=nothing
		set ObjDoc=nothing	
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl&"default.asp"
		Response.End
	end if
	
	'get custom info field
	extrafield = ""
	xmlstring = sitelink.getbasketinfo(session("shopperid"),session("ordernumber"),extrafield,false,session("user_want_multiship"))		
	objDoc.loadxml(xmlstring)
	set SL_Basket = objDoc.selectNodes("//gbi")
	
	'we know the basket record
	'for x=0 to SL_Basket.length-1
	SL_BasketCustInfo	= SL_Basket.item(basketrecord-1).selectSingleNode("custominfo").text
	'next
	set SL_Basket = nothing
	
	
	'response.write(SL_BasketCustInfo)
	
	
			 
	 
	 SET LOSTOCK =  sitelink.setupproduct(current_stocknumber,cstr(current_variation),0,session("ordernumber"),cstr(""))
		StrFileName = "images/"+trim(LOSTOCK.inetimage)
		StrPhysicalPath = Server.MapPath(StrFileName)
	
		  set objFileName = CreateObject("Scripting.FileSystemObject")	
			  if objFileName.FileExists(StrPhysicalPath) then
		       		imagename=StrFileName
			  else
					imagename="images/noimage.gif"
  			  end if
		set objFileName = nothing
		alttag =trim(LOSTOCK.inetsdesc)
		
		'response.write(LOSTOCK.custominfo)
	  

%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%=althomepage%></title>
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



<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top"  >
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<center>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%" >Edit Information &nbsp;&nbsp; </td></tr>
	</table>
		
	<br>
	<form name="registerform" method="POST" action="EditCustomMsg.asp">
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
	  <tr><td><input type="hidden" name="item_id" value="<%=item_id%>">
	  		  <input type="hidden" name="quanto" value="<%=quanto%>">
	  		  <input type="hidden" name="item" value ="<%=current_stocknumber%>">
	  		  <input type="hidden" name="variant" value ="<%=current_variation%>">	  		  
	   </td></tr>
	  
	  <tr><td width="315" valign="top" align="center">
	  	 	<img SRC="<%=imagename%>" alt="<%=alttag%>" hspace="5" border="0">
	  	  </td>
		  <% if LOSTOCK.inetcustom= true or LOSTOCK.giftcert= true   then %>
		  <td width="315" valign="top">
		  <table width="100%">
		  
		  <tr><td ><img alt="" src="images/clear.gif" width="1" height="15" border="0"></td></tr>
		 <tr><td  class="THHeader" ><%=trim(LOSTOCK.inetcprmpt)%></td></tr>
		  <tr><td><textarea  name="txtcustominfo" cols="40" rows="4" value=""><%=SL_BasketCustInfo%></textarea></td></tr>
			<tr>
				<td valign="bottom">
				<br>
				<input type="image" value="modify" src="images/btn-updatebasket.gif"  border="0">
				
				</td>
			</tr>
		  
		  </table>
		  </td>			 
		
		<%end if%>
	<tr><td  >&nbsp;</td>
		<td  colspan =3 class="plaintextbold">&nbsp;	
	</td></tr>

		
		
	</table>
	</form>
	
	
	</center>
	</div>
<!-- end sl_code here -->
	</td>
	
</tr>

</table>

<% SET LOSTOCK = nothing%>

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
