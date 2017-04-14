<%

if session("returnvisitor")=false then
session("returnvisitor")=true
'session("restoreoldorder")=false
'session("restoreoldordernumber")=0
cookiename=cstr(ShortStoreName)+"order"
cookieval = (request.COOKIES(cookiename)("ordernumber"))
newcookieval = replace(cookieval," ","")
if isnumeric(newcookieval)= true then
SL_restoreoldorder = sitelink.READCOOKIE("order",cstr(cookieval))
if len(trim(SL_restoreoldorder)) > 0 then
if isnumeric(SL_restoreoldorder) = true then
'session("ordernumber") = cstr(SL_restoreoldorder)

xmlstring = sitelink.getbasketinfo(session("shopperid"),cstr(SL_restoreoldorder),"pricechang",false,false)
objDoc.loadxml(xmlstring)
set SL_Basket = objDoc.selectNodes("//gbi[certid=0]")

for x=0 to SL_Basket.length-1
SL_Number = SL_Basket.item(x).selectSingleNode("number2").text
SL_Variant= SL_Basket.item(x).selectSingleNode("variant").text
SL_BasketQty          = SL_Basket.item(x).selectSingleNode("quanto").text
SL_Basket_custominfo =SL_Basket.item(x).selectSingleNode("custominfo").text


SL_Basket_PriceChange =SL_Basket.item(x).selectSingleNode("pricechang").text
SL_Basket_Discount =SL_Basket.item(x).selectSingleNode("discount").text
SL_Basket_FullItem =SL_Basket.item(x).selectSingleNode("item").text
SL_Basket_UnitPrice =SL_Basket.item(x).selectSingleNode("it_unlist").text
SL_Basket_VatInc =SL_Basket.item(x).selectSingleNode("vatincl").text
SL_Basket_VatList =SL_Basket.item(x).selectSingleNode("vatlist").text

if SL_Basket_PriceChange="false" then
CALL sitelink.ITEMADD(cstr(SL_Number),cstr(SL_BasketQty),cstr(SL_Variant),session("shopperid"),session("ordernumber"),cstr(SL_Basket_custominfo),cint(0),cint(0))
else
if SL_Basket_VatInc="true" then
SL_Basket_UnitPrice = SL_Basket_VatList
end if
call sitelink.dydasuppadd(cstr(""),cstr(SL_Basket_FullItem),cstr(SL_BasketQty),cdbl(SL_Basket_UnitPrice),session("shopperid"),session("ordernumber"),true,cstr(SL_Basket_Discount),cstr(SL_Basket_custominfo))
end if
next

set SL_Basket = nothing

call SITELINK.quantitypricing(session("shopperid"),session("ordernumber"))

session("SL_BasketCount") = sitelink.GetData_Basket(cstr("BASKETCOUNT"),session("ordernumber"),session("SL_VatInc"))
session("SL_BasketSubTotal") = sitelink.GetData_Basket(cstr("BASKETSUBTOTAL"),session("ordernumber"),session("SL_VatInc"))
end if
end if
end if
end if

%>


<!-- SiteLink generated horizontal navigation -->
<div id="header" class="topnav1bgcolor">
	<!-- aussie logo area: -->
	<div class="divlogo">
		<div class="logo-wrap">
			<div class="logo-img">
				<a href="<%=insecureurl%>" title="<%=althomepage%>">
					<%if COMPANY_LOGO_IMG<>"" then %>
					<img alt="<%=althomepage%>" src="images/<%=COMPANY_LOGO_IMG%>" border="0">
					<%else %>
					<img alt="<%=althomepage%>" src="images/clear.gif" width="1" height="1" border="0">
					<%end if %>
				</a>
			</div>
		</div>
	</div>
	<!-- shopping cart nav: -->
	<ul class="cartnav">
		<!-- shopping cart link -->
		<li class="TopNav1Text">
			<a title="Shopping Cart" href="basket.asp" class="topnav1">Shopping Cart</a>
		</li>
		<!-- shopping cart img -->
		<li class="shopcart">
			<a title="Shopping Cart" href="basket.asp">
				<img src="images/cart-icon.png" alt="Shopping Cart">
			</a>
		</li>
		<li class="TopNav1Text"><%=session("SL_BasketCount")%>&nbsp;Items</li>
		<li class="TopNav1Text"><%=formatcurrency(session("SL_BasketSubTotal"))%></li>
	</ul>

	<ul class="welcome">
		<li class="TopNav1Text">
			<% if session("registeredshopper")="NO" then %>
			<a href="<%=secureurl%>login.asp" title="Login" class="topnav1 j-log-btn">Login</a>
			<% else %> Hello, <%= session("firstname") %>
		</li>
		<li class="TopNav1Text">
			<a href="<%=secureurl%>logout.asp" title="Logout" class="topnav1 j-log-btn">Logout</a>
			<%end if%>
		</li>
	</ul>

	<form method="POST" id="searchprodform" action="searchprods.asp">
		<input name="ProductSearchBy" value="2" type="hidden">
		<!-- Aussie search area:
		<div class="divsearch">
			<ul class="search-wrap">
				<li class="searchbox">
					<input class="plaintext" type="text" size="18" maxlength="256" name="txtsearch"
					       value="Product Search" onfocus="if (this.value=='Product Search') this.value='';"
					       onblur="if (this.value=='') this.value='Product Search';">
				</li>
				<li class="btn-go">
					<input type="image" src="images/btn_go.gif" style="height:24px;width:24px;border:0">
				</li>
			</ul>
		</div>
        -->
	</form>
</div>


<!-- START: top horizontal navigation -->
<nav id="j-nav1" class="navbar navbar-default">
	<div class="container-fluid">
		<!-- Brand and toggle get grouped for better mobile display -->
		<div class="navbar-header">
			<button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1" aria-expanded="false">
				<span class="sr-only">Toggle navigation</span>
				<span class="icon-bar"></span>
				<span class="icon-bar"></span>
				<span class="icon-bar"></span>
			</button>
			<a class="navbar-brand" href="http://www.aussieproducts.com">
				<img alt="Brand" src="images/kangaroo.png" height="24px">
			 </a>
		</div>


		<!-- Bootstrap nav links, forms, and other content for toggling -->
		<div id="j-top-nav-1st" class="collapse navbar-collapse">
			<!-- website page links: -->
			<ul class="nav navbar-nav navbar-left">
				<li><a href="<%=insecureurl%>" title="<%=ALT_HOME_TXT%>">HOME</a></li>
				<%if DISP_ABOUT_PAGE = true then %>
				<li><a href="<%=insecureurl%>Aboutus.asp" title="<%=ALT_ABOUT_TXT%>">ABOUT US</a></li>
				<% end if %>
				<%if DISP_HELP_PAGE = true then %>
				<li><a href="<%=insecureurl%>help.asp" title="<%=ALT_HELP_TXT%>">SHIPPING</a></li>
				<% end if %>
				<%if DISP_CONTACT_PAGE = true then %>
				<li><a href="<%=insecureurl%>contactus.asp" title="<%=ALT_CONTACT_TXT%>">CONTACT US</a></li>
				<% end if %>

				<li><a href="<%=insecureurl%>basket.asp">VIEW CART</a></li>
			</ul>

			<!-- search form area: -->
			<form method="POST" id="searchprodform" action="searchprods.asp" class="navbar-form navbar-right">
				<input name="ProductSearchBy" value="2" type="hidden">
				<div class="form-group search">
					<input class="form-control" type="text" size="18" maxlength="256" name="txtsearch"
					value="Product Search" onfocus="if (this.value=='Product Search') this.value='';"
					onblur="if (this.value=='') this.value='Product Search';"
					>
				</div>

				<button type="submit" class="btn btn-default btn-go">Search</button>
			</form>
			<!-- shopping cart nav:
			<ul class="nav navbar-nav">
				<li>
					<a title="Shopping Cart" href="basket.asp">Shopping Cart</a>
				</li>
				<li>
					 <i class="fa fa-shopping-cart fa-2x j-cart"></i>
				</li>
				<li>
					<a href="basket.asp"><%=session("SL_BasketCount")%>&nbsp;Items</a>
				</li>
				<li>
					<a href="basket.asp"><%=formatcurrency(session("SL_BasketSubTotal"))%></a>
				</li>
			</ul>
			-->
		</div><!-- /.navbar-collapse -->
	</div><!-- /.container-fluid -->
</nav>
<!-- END: top horizontal navigation -->