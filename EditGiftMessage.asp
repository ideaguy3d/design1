<%on error resume next%>

<%
 if session("registeredshopper")<>"YES" then
  	session("destpage")=""
	destpage = secureurl&"login.asp"
  	response.redirect(destpage)
  end if 
  
	shopcust_id = request.querystring("shopcust")
	if isnumeric(shopcust_id)=false then	
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl&"default.asp"
		Response.End
	end if
	
	basketrecord=request.querystring("basketrecord")
	if isnumeric(basketrecord)=false then	
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl&"default.asp"
		Response.End
	end if

%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<%
		
	current_stocknumber = cstr(REQUEST.QUERYSTRING("numberpart"))
    current_stocknumber =fix_xss_Chars(current_stocknumber)
    
    current_variation = cstr(REQUEST.QUERYSTRING("variation"))
    current_variation =fix_xss_Chars(current_variation)
	
	isvalidstock = sitelink.validstocknumber2(current_stocknumber)
	
	if isvalidstock=-1 then
		set sitelink=nothing
		set ObjDoc=nothing	
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl&"default.asp"
		Response.End
	end if
	
	
	
	shopcust_id = shopcust_id + 1 -1
	'response.write(address_id)
	xmlAddressBookstring =sitelink.getaddressbook(session("addressbookcustnum"),shopcust_id)
	
	  objDoc.loadxml(xmlAddressBookstring)
	  set SL_AddressBook = objDoc.selectNodes("//gab")  
	
	  for z=0 to SL_AddressBook.length-1 
	  	 'custnum		= trim(SL_AddressBook.item(z).selectSingleNode("custnum").text)
		 lname			= trim(SL_AddressBook.item(z).selectSingleNode("lastname").text)
		 fname			= trim(SL_AddressBook.item(z).selectSingleNode("firstname").text)
		 'company		= trim(SL_AddressBook.item(z).selectSingleNode("company").text)
		 'addr 		    = trim(SL_AddressBook.item(z).selectSingleNode("addr").text)
		 'addr2 		    = trim(SL_AddressBook.item(z).selectSingleNode("addr2").text)
		 'city 		    = trim(SL_AddressBook.item(z).selectSingleNode("city").text)
		 'vstate 		= trim(SL_AddressBook.item(z).selectSingleNode("state").text)
		 'zipcode 		= trim(SL_AddressBook.item(z).selectSingleNode("zipcode").text)
		 'phone 		    = trim(SL_AddressBook.item(z).selectSingleNode("phone").text)
		 'email 		    = trim(SL_AddressBook.item(z).selectSingleNode("email").text)

		 address_id		= trim(SL_AddressBook.item(z).selectSingleNode("address_id").text)					  
		 shopcust	    = trim(SL_AddressBook.item(z).selectSingleNode("shopcust").text)
	  
	  next	  
	  
	  set SL_AddressBook=nothing
	  Set objDocAddBook=nothing
	  
	  set logiftmsg2=sitelink.editgiftmsg(session("ordernumber"),shopcust_id) 
			giftmsg1=trim(logiftmsg2.note1)
			giftmsg2=trim(logiftmsg2.note2)
			giftmsg3=trim(logiftmsg2.note3)
			giftmsg4=trim(logiftmsg2.note4)
			giftmsg5=trim(logiftmsg2.note5)
			giftmsg6=trim(logiftmsg2.note6)
			
	   set logiftmsg2 = nothing 

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
	  SET LOSTOCK = nothing

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
		<tr><td class="TopNavRow2Text" width="100%" >Add Gift Messages for <%=fname%>&nbsp;<%=lname%>&nbsp;&nbsp; </td></tr>
	</table>
		
	<br>
	<form name="registerform" method="POST" action="AddGiftMessage.asp">
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
	  <tr><td><input type="hidden" name="address_id" value="<%=address_id%>">
	  		  <input type="hidden" name="shopcust" value="<%=shopcust_id%>">
			  <input type="hidden" name="basketrecord" value="<%=basketrecord%>">
			  <input type="hidden" name="addtobook" value="0">
	   </td></tr>
	  
	  <tr><td width="315" valign="top" align="center">
	  	 	<img SRC="<%=imagename%>" alt="<%=alttag%>" hspace="5" border="0">
	  	  </td>
		  
		  <td width="315" valign="top">
		  <table width="100%">
		  	<tr><td >&nbsp;</td><td colspan="2" class="plaintextbold">Message for Recipient </td></tr>
		    <tr><td class="plaintextbold">&nbsp;</td>
	    		<td  colspan ="3"><input type="text" class="bodytext" name="txtgiftmsg1" value="<%=giftmsg1%>" maxlength="35" size="40,1"></td>
		    </tr>
		    <tr><td  >&nbsp;</td>
	    		<td  colspan ="3"><input type="text" class="bodytext" name="txtgiftmsg2" value="<%=giftmsg2%>" maxlength="35" size="40,1"></td>
		    </tr>
			<tr><td  >&nbsp;</td>
			    <td  colspan =3 ><input type="text"  class="bodytext" name="txtgiftmsg3" value="<%=giftmsg3%>" maxlength="35" size="40,1"></td>
			</tr>
		    <tr><td >&nbsp;</td>
			   <td  colspan =3 ><input type="text" class="bodytext" name="txtgiftmsg4" value="<%=giftmsg4%>" maxlength="35" size="40,1"></td>
		    </tr>
		    <tr><td >&nbsp;</td>
	    		<td  colspan =3 ><input type="text" class="bodytext" name="txtgiftmsg5" value="<%=giftmsg5%>" maxlength="35" size="40,1"></td>
		    </tr>
		    <tr><td  >&nbsp;</td>
	    		<td  colspan =3 ><input type="text" class="bodytext" name="txtgiftmsg6" value="<%=giftmsg6%>" maxlength="35" size="40,1"></td>
		    </tr>
			<tr><td>&nbsp;<br><br></td>
				<td valign="bottom">
				<br>
				<input type="image" src="images/btn_Add_GiftMsg.gif" border="0" alt="Add Gift Messsage">
				</td>
			</tr>
		  
		  </table>
		  </td>			 
		
	
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
