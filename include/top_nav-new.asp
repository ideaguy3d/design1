<link rel="stylesheet" href="../text/topnav.css" type="text/css">
<link href="../css/style2.css" rel="stylesheet" type="text/css">
<link href="../css/sidenavstyle.css" rel="stylesheet" type="text/css"> 

<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-22883279-1']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>

<script type="text/javascript">
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
</script>

<div id="header">
	<h1>
		<a href="/">
			<span>Australian Products Co.</span>
		</a>
	</h1>
<div id="topfixed">
<a href="../default.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('home','','../images/home-over.jpg',1)"><img src="../images/home.jpg" width="39" height="22" border="0" id="home" /></a>

<a href="../login.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('login','','../images/login-over.jpg',1)"><img src="../images/login.jpg" width="82" height="22" border="0" id="login" /></a>

<a href="../Quick_order.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('order','','../images/order-over.jpg',1)"><img src="../images/order.jpg" width="70" height="22" border="0" id="order" /></a>

<a href="../shipping.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('shipping policy','','../images/shippingpolicy-over.jpg',1)"><img src="../images/shippingpolicy.jpg" width="84" height="22" border="0" id="shipping policy" /></a>

<a href="../Aboutus.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('about us','','../images/about-over.jpg',1)"><img src="../images/about.jpg" width="57" height="22" border="0" id="about us" /></a>

<a href="../meetthestaff.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('staff','','../images/staff-over.jpg',1)"><img src="../images/staff.jpg" width="80" height="22" border="0" id="staff" /></a>

<a href="../testimonials.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('testimonials','','../images/testimonials-over.jpg',1)"><img src="../images/testimonials.jpg" width="71" height="22" border="0" id="testimonials" /></a>

<a href="../sizechart.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('size chart','','../images/sizechart-over.jpg',1)"><img src="../images/sizechart.jpg" width="68" height="22" border="0" id="size chart" /></a>

<a href="../contactus.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('contact us','','../images/contact-over.jpg',1)"><img src="../images/contact.jpg" width="69" height="22" border="0" id="contact us" /></a>
</div></div>
<div id="space"></div>
<!--table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td>
			
		</td>
		<td>
			<table width="780" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<a href="default.asp"><img src="images/<%=COMPANY_LOGO_IMG%>" border="0" ></a>
					</td>
				</tr>
			</table>
			<table width="780" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="topnav1bgColor"  height="26">
						<table width="100%" border="0" cellpadding="3" cellspacing="0">
							<tr>
								<td>
									<a href="default.asp" class="active"><%=ALT_HOME_TXT%></a>
								</td>
								<td >
									<% if session("registeredshopper")="NO" then %>
									<a href="login.asp" class="topnav1">Account Login</a>
									<%else%>
									<a href="logout.asp" class="topnav1">Logout</a>
									<%end if%>
								</td>
								<td>
									<a href="Quick_order.asp" class="topnav1">Quick Order</a>
								</td>
								<td>
									<a href="shipping.asp" class="topnav1">Shipping Policy </a>
								</td>
								<td>
									<a href="Aboutus.asp" class="topnav1"><%=ALT_ABOUT_TXT%></a>
								</td>
								<%if WANT_WISHLIST = 1 then %>
								<td>
									<a href="faq.asp" class="topnav1"><%=ALT_FAQ_TXT%></a>
								</td>
									<%end if%>
								<td>
									<a href="testimonials.asp" class="topnav1"><%=ALT_TESTIMONIAL_TXT%></a>
								</td>
								<td>
									<a href="../sizechart.asp" class="topnav1">Size Charts</a>
								</td>
								<td>
									<a href="contactus.asp" class="topnav1"><%=ALT_CONTACT_TXT%></a>
								</td>
							</tr>
						</table>
					</td>
				</tr-->
				<tr>
					<td valign="top" class="topnav1bgColor"  height="26">
						<table width="50%" border="0" cellpadding="1" cellspacing="0">
							<tr>
								<td class="TopNavText">
									Subtotal: <%=formatcurrency(SL_BasketSubTotal)%>
								</td>
								<td class="TopNavText">Qty in Basket: <%=SL_BasketCount%>
								</td>
								<td>
									<a href="basket.asp" ><span class="TopNavText">
										<u>View Basket</u></span></a>
								</td>
								<td valign="top" align="right" class="TopNavText">				
									<table border="0" align="center">			
										<form method="POST" action="searchprods.asp">
											<tr>
												<td class="TopNavText">Search
												</td>
												<td>
													<select name="ProductSearchBy" class="smalltextblk" >
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
												<td>
													<input type="text" size="18" maxlength="256" class="plaintext" name="txtsearch" value="<%=replace(session("searchstring"),"""","")%>">
												</td>
												<td>
													<input type="image" src="images/btn_go_green.gif" border="0" width="25" height="19">
												</td>
											</tr>
										</form>
									</table>		
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<!--a id="top_nav_link" href="http://maps.google.com/maps/place?cid=16549058341025031223&q=Australian+Products+Co,+Stevens+Creek+Boulevard,+San+Jose,+CA&hl=en&dtab=0&sll=37.322578,-121.965685&sspn=0.006295,0.006295&ie=UTF8&ll=37.328826,-121.974463&spn=0,0&z=16" target= _blank" id="header_text">
						<div align="center">Hours: 10am - 6pm Pacific Time, Monday - Saturday. Location: 3680 Stevens Creek Blvd, San Jose, CA 95117 USA</a-->
				    </div></td>
				</tr>
			</table>
		</td>
		<td>
		</td>
	</tr>
</table>