<style type="text/css">
<!--
.style1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-weight: bold;
}
-->
</style>

<link rel="stylesheet" type="text/css" href="../css/flexdropdown.css" />
<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js"></script>

<script type="text/javascript" src="../text/flexdropdown.js">
</script>
   <map name="Networks">
  <area shape="rect" coords="1,4,41,46" href="https://www.facebook.com/aussieproducts" alt="Facebook">
  <area shape="rect" coords="63,4,104,45" href="https://twitter.com/aussieproducts" alt="Twitter">
  <area shape="rect" coords="126,4,167,44" href="http://pinterest.com/aussieproducts/" alt="Pintrest">
</map>
<img src="/nav_spacer.gif" alt="" width="15" height="15" /><img src="../images/ARTWORK IMAGES/Website Images/summerroo.png" width="177" height="181" alt="WinterRoo" /><br /><p><a align="left" href="https://www.facebook.com/aussieproducts#" target="_blank">
<script>
  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
  })(window,document,'script','//www.google-analytics.com/analytics.js','ga');

  ga('create', 'UA-45483712-1', 'aussieproducts.com');
  ga('send', 'pageview');

</script>
<img src="http://www.niftybuttons.com/webicons2/facebook.png" width="41" height="41" /></a><a align="left" href="https://twitter.com/aussieproducts" target="_blank"><img src="http://www.niftybuttons.com/webicons2/twitter.png" width="40" height="40" /></a><a align="left" href="http://www.reddit.com/user/AussieProducts/" target="_blank"><img src="http://www.niftybuttons.com/webicons2/reddit.png" width="41" height="41" /></a><a align="left" href="http://www.stumbleupon.com/stumbler/AussieProducts" target="_blank"><img src="http://www.niftybuttons.com/webicons2/stumbleupon.png" width="41" height="41" /></a><a align="left" href="http://pinterest.com/aussieproducts/" target="_blank"><img src="http://i.imgur.com/sxwB0TO.png" width="41" height="41" /></a></p><a href="http://www.aussieproducts.com/newsletter/Theme.mp3" target="_blank"></br><img alt="Bringing Australia to You" src="/images/Theme-button.gif" hspace="20" style="width: 175px; height: 50px;" /></a><br /><hr><span class="sidenavTxt">
 <span class="THHeader"><span class="sidenavTxt">
 <table width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr>
	  
</tr>

		<%
			xmlstring =sitelink.DEPARTMENTS(0,true)
								
			objDoc.loadxml(xmlstring)
			If objDoc.parseError.errorCode <> 0 Then 
			response.write("<tr><td class=""plaintext""><br>Error Processing<br><br></td></tr>")
			else		
			set SL_Dept = objDoc.selectNodes("//thesedepts")
			   
		%>
	<tr>
		<td>
			<table  width="100%" cellpadding="3" cellspacing="0" border="0" >
				<% if session("hasspecials") = 1 then %>
				<tr>
					<td >
						<div align="left">
					 
					  </div>
					</td>
				</tr>
			
				<%end if%>

				<% 			 
			     
			    for x=0 to SL_Dept.length-1
				deptname = SL_Dept.item(x).selectSingleNode("name").text
				deptcode = SL_Dept.item(x).selectSingleNode("deptcode").text
					 
				%>
				<%next%>				       
<tr>
<td>
<div align="center">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=252" data-flexmenu="flexmenu28" data-dir="h" data-offsets="8,0"><b>WHAT'S ON SALE!</a>
<ul id="flexmenu28" class="flexdropdownmenu">
<li><a href="/products.asp?dept=252">Australia</a></li> 
<li><a href="/products.asp?dept=255">British and Irish</a></li>
<li><a href="/products.asp?dept=66">All Countries Two For One Sale</a></li>
<li><a href="/products.asp?dept=262">All Countries 25% Off</a></li>
<li><a href="/products.asp?dept=253">All Countries Clearance Items</a></li>
</ul>
</div>  
</td>
</tr>
</br>

<tr>
<td>
<div align="center">
<img src="/images/a_heading.gif" />
</div>
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=25" data-flexmenu="flexmenu" data-dir="h" data-offsets="8,0"><b>Aboriginal</a>

<ul id="flexmenu" class="flexdropdownmenu">
 <li><a href="/products.asp?dept=108">Art Panels</a></li> 
        <li><a href="/products.asp?dept=175">Books, DVD's & CD's</a></li> 
        <li><a href="/products.asp?dept=107">Boomerangs</a></li> 
        <li><a href="/products.asp?dept=176">Bull Roarers & Click Sticks</a></li> 
        <li><a href="/products.asp?dept=106">Didgeridoos & Protective Bags</a></li> 
        <li><a href="/products.asp?dept=200">Dot Art Wall Animals</a></li> 
        <li><a href="/products.asp?dept=204">Pottery</a></li>
</ul>
</div>
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=58" data-flexmenu="flexmenu2" data-dir="h" data-offsets="8,0"><b>Aussie Animals</a> 

<ul id="flexmenu2" class="flexdropdownmenu">
       <li><a href="/products.asp?dept=209">Canned Animals</a></li> 
        <li><a href="/products.asp?dept=211">Clip-on Animals</a></li> 
        <li><a href="/products.asp?dept=212">Inflatable Animals</a></li> 
        <li><a href="/products.asp?dept=113">Plush Animals</a></li>
         <li><a href="/products.asp?dept=243">Statuettes & Figuerines</a></li>
</ul>
</div>
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=90" data-flexmenu="flexmenu3" data-dir="h" data-offsets="8,0"><b>Aussie Foods</a>

<ul id="flexmenu3" class="flexdropdownmenu">
 <li><a href="/products.asp?dept=76">Arnott's Biscuits</a></li> 
        <li><a href="/products.asp?dept=26">Aussie Spreads</a></li> 
        <li><a href="/products.asp?dept=91">Cadbury Chocolates</a></li>
        <li><a href="/products.asp?dept=27">Cereal</a></li> 
        <li><a href="/products.asp?dept=88">Condiments & Sauces</a></li> 
        <li><a href="/products.asp?dept=60">Cordials & Soda</a></li> 
        <li><a href="/products.asp?dept=71">Crackers, Twisties & Chips</a></li> 
        <li><a href="/products.asp?dept=62">Desserts and Mixes</a></li> 
        <li><a href="/products.asp?dept=89">Fruit - Canned, Dried, Glace & Ginger</a></li>
        <li><a href="/products.asp?dept=73">Lollies & Licorice</a></li> 
        <li><a href="/products.asp?dept=31">Meat Pies, Pasties, Sausage Rolls & Lamingtons</a></li> 
        <li><a href="/products.asp?dept=28">Soups</a></li>  
        <li><a href="/products.asp?dept=29">Spices & Seasoning</a></li> 
        <li><a href="/products.asp?dept=61">Tea & Coffee</a></li> 
        <li><a href="/products.asp?dept=55">Tim Tams</a></li> 
</ul>
</div>
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=59" data-flexmenu="flexmenu4" data-dir="h" data-offsets="8,0"><b>Aussie Pharmacy</a>

<ul id="flexmenu4" class="flexdropdownmenu">
 <li><a href="/prodinfo.asp?number=DDAM10">Am-O-Lin</a></li> 
        <li><a href="/prodinfo.asp?number=DDAN10">Anticol</a></li> 
        <li><a href="/prodinfo.asp?number=DDGC10">Arthritis Cream</a></li>
        <li><a href="/prodinfo.asp?number=DDBERO2">Berocca</a></li> 
        <li><a href="/prodinfo.asp?number=DDBM10">Butter-Menthol</a></li> 
        <li><a href="/prodinfo.asp?number=DDSS3PK">Dettol</a></li>   
        <li><a href="/prodinfo.asp?number=DDDM24">Disprin</a></li>
        <li><a href="/prodinfo.asp?number=DDEN20  200G">Eno</a></li> 
        <li><a href="/prodinfo.asp?number=DDEO50">Eucalyptus Oil</a></li>
        <li><a href="/prodinfo.asp?number=DDMEXT">Malt Extract</a></li> 
        <li><a href="/prodinfo.asp?number=DDPT24">Panadol</a></li> 
        <li><a href="/prodinfo.asp?number=DDQEZ5">Quick-Eze</a></li>
        <li><a href="/prodinfo.asp?number=DDRB60">Ribena</a></li> 
        <li><a href="/prodinfo.asp?number=DDSV37">Salvital</a></li> 
        <li><a href="/prodinfo.asp?number=DDSC30">Savlon</a></li>
        <li><a href="/products.asp?dept=53">Tea Tree Products</a></li> 
        <li><a href="/prodinfo.asp?number=DDTR10">Teething Rusks</a></li> 
</ul>
</div>
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=78"><b>Baby Corner</a>
</div>
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=65" data-flexmenu="flexmenu6" data-dir="h" data-offsets="8,0"><b>Books, DVD's & CD's</a> 
<ul id="flexmenu6" class="flexdropdownmenu">
<li><a href="/products.asp?dept=94">Activity & Educational Books</a></li>  
        <li><a href="/products.asp?dept=35">Childrens Classics</a></li> 
        <li><a href="/products.asp?dept=41">Coffee Table/Steve Parish</a></li> 
        <li><a href="/products.asp?dept=98">Discover & Learn/Field Guides</a></li> 
        <li><a href="/products.asp?dept=116">DVD's & CD's</a></li>
		<li><a href="/products.asp?dept=96">Fact & Reference Books</a></li> 
        <li><a href="/products.asp?dept=95">Jigsaw & Puzzle Books</a></li>
        <li><a href="/products.asp?dept=160">Paperback</a></li> 
        <li><a href="/products.asp?dept=97">Toddler, 0-8 Reading & Picture Books</a></li>
</ul>
</div>
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=225" data-flexmenu="flexmenu7" data-dir="h" data-offsets="8,0"><b>Celebrate Australia - Flags, etc.</a>
<ul id="flexmenu7" class="flexdropdownmenu">
<li><a href="/products.asp?dept=63">Australian Flags/Green & Gold</a></li> 
<li><a href="/products.asp?dept=31">Meat Pies, Pasties, Sausage Rolls & Lamingtons</a></li> 
</ul>
</div>
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=64"><b>Celebrate New Zealand - Flags, etc.</a>
</div>
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=68" data-flexmenu="flexmenu9" data-dir="h" data-offsets="8,0"><b>Christmas Downunder</a>
<ul id="flexmenu9" class="flexdropdownmenu">
<li><a href="/products.asp?dept=210">Cakes, Puddings & Custard</a></li> 
<li><a href="/products.asp?dept=214">Christmas Crackers, Cards & Music</a></li> 
<li><a href="/products.asp?dept=218">Christmas Decor & Stocking Stuffers</a></li>   
<li><a href="/products.asp?dept=154">Gift Boxes, Crates, Mugs & Totes</a></li> 
<li><a href="/products.asp?dept=217">Glace Fruit</a></li>
</ul>
</div>
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=69" data-flexmenu="flexmenu10" data-dir="h" data-offsets="8,0"><b>Down Under Delights</a>
<ul id="flexmenu10" class="flexdropdownmenu">
<li><a href="/prodinfo.asp?number=APLA12">Lamingtons</a></li> 
<li><a href="/products.asp?dept=37">Meat Pies</a></li> 
<li><a href="/products.asp?dept=52">Pasties</a></li> 
<li><a href="/products.asp?dept=74">Sausage Rolls</a></li>
</ul>
</div>                      
</td>
</tr>
     
<tr>
<td>         
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=56" data-flexmenu="flexmenu11" data-dir="h" data-offsets="8,0"><b>Fine Art Prints</a>
<ul id="flexmenu11" class="flexdropdownmenu">
<li><a href="/products.asp?dept=56">Dr. John Xiong</a></li> 
</ul>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=70" data-flexmenu="flexmenu12" data-dir="h" data-offsets="8,0"><b>Footwear</a>
<ul id="flexmenu12" class="flexdropdownmenu">
<li><a href="/products.asp?dept=45">Blundstone Boots/Masseur Sandals</a></li> 
<li><a href="/products.asp?dept=203">Koala Slippers (adult and child)/Fuzzie Footies/Wool Socks</a></li>
<li><a href="/products.asp?dept=189">Sheepskin Boots, Slippers & Insoles</a></li>
</ul>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=154"><b>Gift Boxes, Crates, Mugs & Totes</a>
</div>  	
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=92"><b>Gift Certificates</a>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=51" data-flexmenu="flexmenu15" data-dir="h" data-offsets="8,0"><b>Hats & Caps</a>
<ul id="flexmenu15" class="flexdropdownmenu">
<li><a href="/products.asp?dept=196">Aussie Sun Hats/Slouch Hats/Umbrella Hats</a></li> 
<li><a href="/products.asp?dept=42">Akubra Hats</a></li> 
<li><a href="/products.asp?dept=123">Baseball Caps/Beanies/Visors</a></li>
</ul>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=202" data-flexmenu="flexmenu16" data-dir="h" data-offsets="8,0"><b>Health & Beauty</a>
<ul id="flexmenu16" class="flexdropdownmenu">
<li><a href="/products.asp?dept=202">Health & Beauty Products</a></li>  
<li><a href="/products.asp?dept=53">Tea Tree Products</a></li>
</ul>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=87" data-flexmenu="flexmenu17" data-dir="h" data-offsets="8,0"><b>Jewelry & Watches</a>
<ul id="flexmenu17" class="flexdropdownmenu">
<li><a href="/products.asp?dept=250">Beate's Hand Painted Porcelain</a></li> 
<li><a href="/products.asp?dept=128">Brooches/Tie Pins</a></li> 
<li><a href="/products.asp?dept=131">Charms, Phone Charms & Keychains</a></li> 
<li><a href="/products.asp?dept=129">Earrings</a></li>
<li><a href="/products.asp?dept=130">Pendants</a></li> 
<li><a href="/products.asp?dept=126">Pins</a></li> 
<li><a href="/products.asp?dept=132">Watches</a></li>
</ul>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=167"><b><b>Kangaroo Collectibles</a>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=79" data-flexmenu="flexmenu19" data-dir="h" data-offsets="8,0"><b>Kitchen & Home Decor</a>
<ul id="flexmenu19" class="flexdropdownmenu">
<li><a href="/products.asp?dept=122">Aprons, Oven Mitts & Tea Towels</a></li> 
<li><a href="/products.asp?dept=117">Children's Room</a></li>
<li><a href="/products.asp?dept=13">Gift Totes</a></li> 
<li><a href="/products.asp?dept=17">Kitchenware</a></li>
<li><a href="/products.asp?dept=157">Paper Napkins, Plates & Cups</a></li>
<li><a href="/products.asp?dept=169">Pottery & Pewter</a></li>
<li><a href="/products.asp?dept=162">Statuettes & Figurines</a></li> 
</ul>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=254"><b>Koala Collectibles</a>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=77" data-flexmenu="flexmenu21" data-dir="h" data-offsets="8,0"><b>Laundry</a>
<ul id="flexmenu21" class="flexdropdownmenu">
<li><a href="/products.asp?dept=77">Napisan, Sard & Eucalyptus</a></li>
</ul>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=80" data-flexmenu="flexmenu22" data-dir="h" data-offsets="8,0"><b>Outback Clothing & Accessories</a>
<ul id="flexmenu22" class="flexdropdownmenu">
<li><a href="/products.asp?dept=127">Aussie Fabrics & Embroidered Patches</a></li>
<li><a href="/products.asp?dept=42">Akubra Hats</a></li> 
<li><a href="/products.asp?dept=177">Beach Towels</a></li> 
<li><a href="/products.asp?dept=43">Driz-a-bone Coats</a></li> 
<li><a href="/products.asp?dept=166">Gloves, Scarves, Belts & Ties</a></li> 
<li><a href="/products.asp?dept=47">T-Shirts & Sweatshirts</a></li>
<li><a href="/products.asp?dept=224">Shirts & Shorts</a></li>
<li><a href="/products.asp?dept=105">Wallets & Pouches</a></li> 
</ul>
</div> 
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=317"><b>Pet Products</a>
</div> 
</td>
</tr>


<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=109"><b>Puzzles, Games 
& Sports</a>
</div> 
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=84" data-flexmenu="flexmenu24" data-dir="h" data-offsets="8,0"><b>Road Signs</a>
<ul id="flexmenu24" class="flexdropdownmenu">
<li><a href="/products.asp?dept=135">On Board</a></li> 
<li><a href="/products.asp?dept=134">Road</a></li>
</ul>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=83" data-flexmenu="flexmenu25" data-dir="h" data-offsets="8,0"><b>Sheepskin Products</a>
<ul id="flexmenu25" class="flexdropdownmenu">
<li><a href="/products.asp?dept=14">Car Seat Covers & Accessories</a></li> 
<li><a href="/products.asp?dept=197">Medical Sheepskin</a></li>
<li><a href="/products.asp?dept=184">Rugs & Cushions</a></li> 
<li><a href="/products.asp?dept=189">Sheepskin Boots, Slippers & Insoles</a></li>
</ul>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=82" data-flexmenu="flexmenu26" data-dir="h" data-offsets="8,0"><b>Stationery</a>
<ul id="flexmenu26" class="flexdropdownmenu">
<li><a href="/products.asp?dept=38">Bookmarks</a></li> 
<li><a href="/products.asp?dept=163">Calendars, Diaries & Address Books</a></li>
<li><a href="/products.asp?dept=158">Desktop & Stationery Accessories</a></li> 
<li><a href="/products.asp?dept=67">Gift Wrap, Cards & Postcards</a></li>
<li><a href="/products.asp?dept=118">Magnets</a></li> 
<li><a href="/products.asp?dept=57">Maps & Travel</a> 
<li><a href="/products.asp?dept=9">Pens, Pencils & Rulers</a></li><li><a href="/products.asp?dept=164">Posters & Prints</a></li> 
<li><a href="/products.asp?dept=85">Stickers & Tattoos</a></li> 
</ul>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=99" ><b>The Pub</a>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">
<a href="/timtam.asp"><img src="/images/T_heading.jpg" alt="Arnott's Tim Tam Biscuits - Buy Tim Tams - Aussie Products.com" /></a><br>
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/timtam.asp"><b>The World Famous Tim Tam</b></a>
</div>  
</td>
</tr>

<tr>
<td>
<div align="left">

<a href="/vegemite.asp"><img src="/images/V_heading.jpg" alt="Kraft Vegemite - Aussie Products.com Buy Vegemite!" /></a><br>
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/vegemite.asp"><b>Australia's Favourite-Vegemite</b></a>
</div>  
</td>
</tr>

<tr>
<td>
<div align="center">
<br>
<a href="/products.asp?dept=223"><img src="/images/B_heading.gif" /></a><BR>
<tr>
<td>
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=240"><b>British Christmas</b></a>
</td>
</tr>
<tr>
<td>
<div align="left">
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=241" data-flexmenu="flexmenu7" data-dir="h" data-offsets="8,0"><b>British Easter</a>
</div>
</td>
</tr>
<tr>
<td>
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=223" data-flexmenu="flexmenu31" data-dir="h" data-offsets="8,0"><b>British Foods</b></a>
<ul id="flexmenu31" class="flexdropdownmenu">
<li><a href="/products.asp?dept=231">British Beverages</a></li> 
<li><a href="/products.asp?dept=226">British Biscuits</a></li> 
<li><a href="/products.asp?dept=229">British Cereal</a></li> 
<li><a href="/products.asp?dept=228">British Chocolates </a></li> 
<li><a href="/products.asp?dept=240">British Christmas</a></li>
<li><a href="/products.asp?dept=230">British Condiments & Sauces</a></li>
<li><a href="/products.asp?dept=232">British Crisps & Snacks</a></li>
<li><a href="/products.asp?dept=233">British Desserts and Mixes</a></li>
<li><a href="/products.asp?dept=241">British Easter</a></li>
<li><a href="/products.asp?dept=222">British Frozen Foods</a></li> 
<li><a href="/products.asp?dept=234">British Fruit - Canned, Dried,Glace & Ginger</a></li>
<li><a href="/products.asp?dept=221">British Grocery</a></li> 
<li><a href="/products.asp?dept=236">British Meat Pies, Pasties, Bangers & Bacon</a></li> 
<li><a href="/products.asp?dept=237">British Soups</a></li>
<li><a href="/products.asp?dept=238">British Spices & Seasoning</a></li>
<li><a href="/products.asp?dept=227">British Spreads</a></li>
<li><a href="/products.asp?dept=235">British Sweets</a></li>
<li><a href="/products.asp?dept=239">British Tea</a></li>
</ul>
</div>  
</td>
</tr>
<tr>
<td>
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=248"><b>British Flags and Decals</b></a>

</td>
<tr>
<br>
<tr>
<td>
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=249"><b>British Health & Beauty Products</b></a>
</td>
</tr>
<tr>
<td>
<div align="center">
<a href="/products.asp?dept=242"><img src="/images/I_heading.gif" /></a><BR>
<img src="/spacer2.gif" width="6" height="9" /><a class="sidenav2" href="/products.asp?dept=242"><b>IRISH FOOD</b></a>
</div>
</tr>
</td>

								</li>
							</ul>
						</div>
					</td>
			    </tr>
				<tr>
					<td>
						
					</td>

				</tr>						
			</table>

		</td>
	</tr>
	<tr>
		<td align="center" class="sidenavTxt">     ACCOUNT INFO
	  </td>
	</tr>
	<tr>
		<td height="1" width="<%=SIDE_NAV_WIDTH%>" class="THHeader" ><img src="images/clear.gif" height="1" border="0">
		</td>
	</tr>
				
	<%
	set SL_Dept=nothing			   
	end if
	%>
<%
	'Set objDoc = nothing
	%>
	<tr>
		<td>
			<table  width="100%" cellpadding="3" cellspacing="0" border="0" >
				<tr>
					<td height="5">
					</td>
				</tr>
				<tr>
					<td>
						<div align="left">
							<a class="sidenav2" href="basket.asp">View Basket</a>
						</div>
					</td>
				</tr>
				<tr>
					<td>
						<div align="left">
							<a class="sidenav2" href="<%=session("secureurl")%>statuslogin.asp">Order Status</a>
						</div>
					</td>
				</tr>
				<tr>
					<td>
						<div align="left">
							<a class="sidenav2" href="address.asp">Address Book</a>
						</div>
					</td>
				</tr>
				<tr>
					<td>
						<div align="left">
							<a class="sidenav2" href="wishlist.asp">My Wish List </a>
						</div>
					</td>
				</tr>
                <tr>
                <td>
                <a class="twitter-timeline" href="https://twitter.com/aussieproducts" data-widget-id="304331663255670784">Tweets by @aussieproducts</a>
<script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>

                </td>
                </tr>
                <tr>
                <td><div id="fb-root"></div>
<script>(function(d, s, id) {
  var js, fjs = d.getElementsByTagName(s)[0];
  if (d.getElementById(id)) return;
  js = d.createElement(s); js.id = id;
  js.src = "//connect.facebook.net/en_GB/all.js#xfbml=1";
  fjs.parentNode.insertBefore(js, fjs);
}(document, 'script', 'facebook-jssdk'));</script>
                </td>
                </tr>
				<tr>
					<td>
                    <br /><img src="/nav_spacer.gif" width="200" height="15" /><div align="center">
     <div class="sidenav2">
     
     <strong>Companies that we work with past/present:</strong><br />
     <img src="/nav_spacer.gif" width="200" height="20" /><p>
			  Tourism Australia<br />
			  Seven Network Australia <br>
			  AusTrade<br />
			  MCI<br />
			  Avon America<br />
			  Avon Japan<br />
			  Qantas<br />
			  Maniscalco Stone<br />
			  LinkedIn<br />
			  JAG<br />
			  Survivor<br />
			  Harpo Productions<br />
			  Paramount Studios<br />
			  Google<br />
              Yahoo!<br />
              Somnamed<br />
		    VISA <br />
            The Ellen DeGeneres Show<br/>

            	  
		    <img src="/nav_spacer.gif" width="200" height="30" /><p></td>
				</tr>
			</table>
  </tr>
	</td>
    
			<!--<tr><td><img src="images/clear.gif" width="1" height="160" border="0"></td></tr>-->

<a href="https://plus.google.com/115506228003611122326" rel="publisher"></a>
</table>
<div align="left">