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
<% if request.querystring("id") = "pgtitle" then %>
<table width="600" cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%" >Basket</td></tr>
	</table>
<%end if%>

<% if request.querystring("id") = "tbheader" then %>
<table width="600" cellpadding="3" cellspacing="1" border="0">

		<tr>
			<th align="CENTER"  class="THHeader" width="6%">Remove</th>
			<th align="CENTER"  class="THHeader" width="6%">Qty</th>
			<th align="CENTER"  class="THHeader" width="70%">Description</th>
			<th align="CENTER"  class="THHeader" width="18%">Total</th>
		</tr>
	</table>
<%end if%>

<% if request.querystring("id") = "row1" then %>
<table width="600" cellpadding="3" cellspacing="1" border="0">

		<tr>
			<td align="CENTER"  class="tdRow1Color " width="15%"><span class="plaintext">Row 1 Color</span></td>			
			<td align="CENTER"  class="tdRow1Color " width="85%"><span class="plaintext">Row 1 Color</span></td>
		</tr>
		<tr>
			<td align="CENTER"  class="tdRow2Color " width="15%"><span class="plaintext">Row 2 Color</span></td>			
			<td align="CENTER"  class="tdRow2Color " width="85%"><span class="plaintext">Row 2 Color</span></td>
		</tr>
		<tr>
			<td align="CENTER"  class="tdRow1Color " width="15%"><span class="plaintext">Row 1 Color</span></td>			
			<td align="CENTER"  class="tdRow1Color " width="85%"><span class="plaintext">Row 1 Color</span></td>
		</tr>
		<tr>
			<td align="CENTER"  class="tdRow2Color " width="15%"><span class="plaintext">Row 2 Color</span></td>			
			<td align="CENTER"  class="tdRow2Color " width="85%"><span class="plaintext">Row 2 Color</span></td>
		</tr>

	</table>
<%end if%>
<% if request.querystring("id") = "active_link" then %>
		<table  width="600" cellpadding="5" cellspacing="0" border="0" >
			  
				<tr>
					<td >
					<a href="#">Test Link </a>
					<br><br>
					<a href="#">Test Link </a>
					<br><br>
					<a href="#">Test Link </a>
					
					</td>
				</tr>
			
						
				</table>

<%end if%>

<% if request.querystring("id") = "text" then %>
		<table  width="600" cellpadding="5" cellspacing="0" border="0" >
			  
				<tr>
					<td class="plaintext">
					Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;
					<br><br>
					Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;
					<br><br>
					
					</td>
				</tr>
			
						
				</table>

<%end if%>
<% if request.querystring("id") = "ptitle" then %>
		<table  width="600" cellpadding="5" cellspacing="0" border="0" >
			  
				<tr>
					<td class="ProductTitle">
					Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;
					<br><br>
					Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;
					<br><br>
					
					</td>
				</tr>
			
						
				</table>

<%end if%>
<% if request.querystring("id") = "pprice" then %>
		<table  width="600" cellpadding="5" cellspacing="0" border="0" >
			  
				<tr>
					<td class="ProductPrice">
					Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;
					<br><br>
					Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;
					<br><br>
					
					</td>
				</tr>
			
						
				</table>

<%end if%>
<% if request.querystring("id") = "pcompprice" then %>
		<table  width="600" cellpadding="5" cellspacing="0" border="0" >
			  
				<tr>
					<td class="CompPrice">
					Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;
					<br><br>
					Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;Test&nbsp;
					<br><br>
					
					</td>
				</tr>
			
						
				</table>

<%end if%>


<br><br>
<br><br>
<a href="Javascript:window.close()"><span class="CompPrice">Close Window</span></a>

</center>





<!--#INCLUDE FILE = "include/AdminReleaseSL.asp" -->
