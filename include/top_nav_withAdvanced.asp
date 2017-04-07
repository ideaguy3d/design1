
<table width="780" border="0" cellpadding="0" cellspacing="0">
<tr>
      <td width="780" height="100" background="images/TopR1-C1.gif" > 
	  <img src="images/clear.gif" width="20" height="1" border="0" align="right"> 
        <table width="30%" align="right" cellpadding="0" cellspacing="0" border="0">
		<tr><td colspan="3" ><img src="images/clear.gif"  border="0" height="30" width="1"></td></tr>
			<tr><td colspan="3" bgcolor="<%=TB_HEADER_COLOR%>"><img src="images/clear.gif" width="1" border="0"></td></tr>
			<tr><td width="1" bgcolor="<%=TB_HEADER_COLOR%>"><img src="images/clear.gif" width="1" border="0"></td>	
				<td bgcolor="white">
				<table width="100%" cellpadding="3" cellspacing="1" border="0">
					<tr><td class="plaintext">Subtotal : <%=formatcurrency(SL_BasketSubTotal)%></td>
						<td class="plaintext" align="right" valign="top"><a href="basket.asp">View Basket</a>
						&nbsp;&nbsp;
					</td>
					</tr>
					<tr><td class="plaintext">Qty in Basket : <%=SL_BasketCount%></td>
						<td class="plaintext" align="right" valign="top">
						<a  href="<%=secureurl%>statuslogin.asp">Order Status</a>
						&nbsp;&nbsp;
					</tr>
					
				</table>
				</td>
				<td width="1" bgcolor="<%=TB_HEADER_COLOR%>"><img src="images/clear.gif" width="1" border="0"></td>
			</tr>
			<tr><td colspan="3" bgcolor="<%=TB_HEADER_COLOR%>"><img src="images/clear.gif" width="1" border="0"></td></tr>
		</table>
</td></tr>
</table>
<table width="780" border="0" cellpadding="0" cellspacing="0">
<tr>
  <td valign="top" bgcolor="#08416C" height="26">
   <table width="100%" cellpadding="3" cellspacing="0" border="0">
	   <td ><a href="default.asp" class="topnav1">Home</a></td>
	   <td ><a href="Aboutus.asp" class="topnav1">About us</a></td>
	   <td >
	   <% if session("registeredshopper")="NO" then %>
	   <a href="login.asp" class="topnav1">Account Login</a>
	   <%else%>
			<a href="logout.asp" class="topnav1">Logout</a>
		<%end if%>
	   </td>
	   <td ><a href="testimonials.asp" class="topnav1">Testimonials</a></td>
	   <td ><a href="Quick_order.asp" class="topnav1">Quick Order</a></td>
	   <%if WANT_WISHLIST = 1 then %>
	   <td><a href="wishlist.asp" class="topnav1">Wish List</a></td>
	   <%end if%>
	   <td ><a href="privacy.asp" class="topnav1">Privacy Policy</a></td>
	   <td ><a href="contactus.asp" class="topnav1">Contact Us</a></td>
	   <td ><a href="faq.asp" class="topnav1">Faq</a></td>
   </table>
</td>
</tr>
<tr>
  <td valign="top" bgcolor="#08416C" height="26">
   <table width="100%" cellpadding="0" cellspacing="0" border="0">
  
	
	<td valign="top" align="right">
		<table>
		<form method="POST" action="searchprods.asp" >
			<tr><td><span class="TopNavRow1Text">Search</span></td>
				<td><select name="ProductSearchBy" class="smalltextblk" >
					<option value="1" <%if session("searchcriteria") ="PARTNUMBER" then response.write(" selected") end if%> >Part Number</option>
					<option value="2" <%if session("searchcriteria") ="DESC" then response.write(" selected") end if%> >Description</option>			
					<% if len(trim(session("SL_Advanced1"))) > 0 then %>
						<option value="3" <%if ucase(session("searchcriteria")) =ucase(session("SL_Advanced1")) then response.write(" selected") end if%> ><%=session("SL_Advanced1")%></option>
					<%end if%>
					<% if len(trim(session("SL_Advanced2"))) > 0 then %>
						<option value="4" <%if ucase(session("searchcriteria")) =ucase(session("SL_Advanced2")) then response.write(" selected") end if%> ><%=session("SL_Advanced2")%></option>
					<%end if%>
					<% if len(trim(session("SL_Advanced3"))) > 0 then %>
						<option value="5" <%if ucase(session("searchcriteria")) =ucase(session("SL_Advanced3")) then response.write(" selected") end if%> ><%=session("SL_Advanced3")%></option>
					<%end if%>
					<% if len(trim(session("SL_Advanced4"))) > 0 then %>
						<option value="6" <%if ucase(session("searchcriteria")) =ucase(session("SL_Advanced4")) then response.write(" selected") end if%> ><%=session("SL_Advanced4")%></option>
					<%end if%>
					
					</select>
				</td>
				<td><input type="text" size="18" maxlength="256" class="smalltextblk" name="txtsearch" value="<%=session("searchstring")%>"></td>
				<td><input type="image" src="images/btn_go.gif" border="0" width="25" height="19"></td>
			</tr>
		</form>
		</table>		
	</td>
	</tr>
   </table>
</td>
</tr>

</table>