
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td  background="images/top_navImg_Bkg.gif"><a href="<%=insecureurl%>"><img alt="<%=session("storename")%>" src="images/<%=COMPANY_LOGO_IMG%>" border="0" ></a></td></tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
  <td valign="top" class="topnav1bgColor"  height="26">
   <table width="100%" cellpadding="3" cellspacing="0" border="0">
	<tr>
	   <td ><a href="<%=insecureurl%>" class="topnav1"><%=ALT_HOME_TXT%></a></td>
	   <td ><a href="Aboutus.asp" class="topnav1"><%=ALT_ABOUT_TXT%></a></td>
	   <td ><a href="sitemap.asp" class="topnav1">Site Map</a></td>	  
	   <td ><% if session("registeredshopper")="NO" then %><a href="login.asp" class="topnav1">Account Login</a><%else%><a href="logout.asp" class="topnav1">Logout</a><%end if%></td>
	   <td ><a href="testimonials.asp" class="topnav1"><%=ALT_TESTIMONIAL_TXT%></a></td>
	   <td ><a href="Quick_order.asp" class="topnav1">Quick Order</a></td>
	   <%if WANT_WISHLIST = 1 then %>
	   <td><a href="wishlist.asp" class="topnav1">Wish List</a></td>
	   <%end if%>
	   <td ><a href="privacy.asp" class="topnav1"><%=ALT_PRIVACY_TXT%></a></td>
	   <td ><a href="contactus.asp" class="topnav1"><%=ALT_CONTACT_TXT%></a></td>
	   <td ><a href="faq.asp" class="topnav1"><%=ALT_FAQ_TXT%></a></td>
	  </tr>
   </table>
</td>
</tr>
<tr>
  <td valign="top" class="topnav1bgColor"  height="26">
   <table width="100%" cellpadding="1" cellspacing="0" border="0">
   <tr>
   	<td class="TopNavText">
	Subtotal: <%=formatcurrency(SL_BasketSubTotal)%>
	</td>
   <td class="TopNavText">Qty in Basket : <%=SL_BasketCount%></td>
   <td ><a href="basket.asp" ><span class="TopNavText"><u>View Basket</u></span></a></td>
	
	<td valign="top" align="right" class="TopNavText">				
		<table border="0">			
		<form method="POST" action="searchprods.asp">
			<tr>
				<td class="TopNavText">Search</td>
				<td><select name="ProductSearchBy" class="smalltextblk" >
					<option value="2" <%if session("searchcriteria") ="DESC" then response.write(" selected") end if%> >Description</option>
					<option value="1" <%if session("searchcriteria") ="PARTNUMBER" then response.write(" selected") end if%> >Item Number</option>					
					<% if USE_ADVANCED_SEARCH= 1 then %>
						<% if len(trim(session("SL_Advanced1"))) > 0  then %>
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
					<%end if%>

					</select>
				</td>
				<td><input type="text" size="18" maxlength="256" class="plaintext" name="txtsearch" value="<%=replace(session("searchstring"),"""","")%>"></td>
				<td><input type="image" src="images/btn_go.gif" style="height:19;width:25;border:0"></td>
			</tr>
		</form>
		</table>		
	</td>
	</tr>
   </table>
</td>
</tr>
</table>