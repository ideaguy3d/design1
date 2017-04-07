<%on error resume next%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%

        if session("billtocopy")=1 then
    	    vtxtzipCode = session("zipcode")
	        vtextshipcountry = session("country")
	        vtxtstate = session("state")
        else
	        vtxtzipCode = session("szipcode")
	        vtextshipcountry = session("scountry")
	        vtxtstate = session("sstate")
        end if

        'vtextshipcountry = REQUEST.querystring("country")
        'vtxtstate  = REQUEST.querystring("vstatecode")
        'vtxtzipCode = ucase(REQUEST.querystring("vzipcode"))
        
        shipping_order="3"				
		shipping_order = cstr(SHIPPING_SORTBY) 
		if SHIPPING_SORTORDER=1 then
			shipping_order = shipping_order +" desc"
		end if 
		
		'default shipping method will be the ship_via that appears most on line item
		session("shippingmethod")=""
		
		
		xmlShiplist = sitelink.GET_SHIPLIST_ZONES(cstr(vtextshipcountry),cstr(vtxtzipCode))
		objDoc.loadxml(xmlShiplist)
			    
			    ShipMethod=""
			    if DEFAULT_SHIPPING_METHOD<>"" then
			        target_node = "//shiplist[ca_code='"+cstr(DEFAULT_SHIPPING_METHOD)+"']"
			        set SL_Ship = objDoc.selectNodes(target_node) 
			            if SL_Ship.length > 0 then
			                ShipMethod = DEFAULT_SHIPPING_METHOD
			            end if
			        set SL_Ship = nothing
			    end if
			    
			    if ShipMethod="" then
			        target_node = "//shiplist"
			        set SL_Ship = objDoc.selectNodes(target_node) 
			            if SL_Ship.length > 0 then
			                ShipMethod = SL_Ship.item(0).selectSingleNode("ca_code").text
			            end if
			        set SL_Ship = nothing 			    
			    end if
			    
			    session("shippingmethod")= ShipMethod
				'response.write(session("shippingmethod"))
		
		
		
		xmlstring=sitelink.previewshipping(session("shopperid"),session("ordernumber"),cstr(vtextshipcountry),UCase(cstr(vtxtstate)),cstr(vtxtzipCode),shipping_order,session("shippingmethod"))						
		
						
		objDoc.loadxml(xmlstring)	
	
	
	'response.Write("Shopper-->"& session("shopperid"))
	'response.Write("<br>")
	'response.Write("Order-->"& session("ordernumber"))

%>


<html>
<head>
<title>Aussie Products.com | Preview Customer Shipping @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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

<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
<SCRIPT language="JavaScript">
   
 function CheckForEmpty(inputstr)
{
  if(inputstr.length == 0) return true ;
  for( var i = 0 ; i <  inputstr.length ; i++)
  {
   var onechar = inputstr.charAt(i) ;
   if( onechar != " ") return false ;
  }
 return true ;
}


function Process() {
     var form = document.forms[0] ;
     for (var i = 0 ; i < form.elements.length ; i++)
      {	   	  
		if( form.elements[i].name =="Listcountry") 
			var countrycode = form.elements[i][form.elements[i].selectedIndex].value // only IE !! form.elements[i].value ;
		if( form.elements[i].name == "txtzipcode" && countrycode == "001")
		{
		  var inputstr =  form.elements[i].value ;  
		  if(CheckForEmpty(inputstr))
		   {         
			 window.alert("Please enter a Zip Code");
			 return false ;         
		   }               
		}
      }	 
	 return true ;
}

function makePOSTRequest(params)
{
    //alert(params)
    //var new_url="SaveCheckoutInfo.asp" ;
    if (window.XMLHttpRequest)
		  {
		  xmlhttp=new XMLHttpRequest()
		  }
		// code for IE
		else if (window.ActiveXObject)
		  {
		  xmlhttp=new ActiveXObject("Msxml2.XMLHTTP");
		  }
		
		if (xmlhttp!=null)
		{
		    xmlhttp.onreadystatechange=state_Change
		    xmlhttp.open("POST","ApplyShippingMethod.asp",false) ;
		    xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		    //xmlhttp.setRequestHeader("Content-type", "text/xml");
		    xmlhttp.setRequestHeader("Content-length", params.length);
		    xmlhttp.setRequestHeader("Connection", "close");
		    xmlhttp.send(params);
		    setTimeout('', 5000);
		}
		else
		  {
		  alert("Your browser does not support XMLHTTP.")
		  }

}
function state_Change()
{
// if xmlhttp shows "loaded"
if (xmlhttp.readyState==4)
  {
  //alert(PPxmlhttp.statusText)
  // if "OK"
  if (xmlhttp.status==200)
    {
    // ...some code here...
    }
  else
    {
    alert("Problem retrieving XML data")
    }
  }
}

function ApplyThisMethod(ShipMethod)
{

    //var pageurl = "ApplyShippingMethod.asp?txtshiplist=" + ShipMethod
    
    //window.close() ;
    //window.location=pageurl;    
    //window.opener.location.reload(true);
    
    var poststr ="txtshiplist=" + encodeURI(ShipMethod)
   	makePOSTRequest(poststr);
	
	//alert("hello " + window.opener.location.href);
    window.opener.location.href = window.opener.location.href;
    window.close(); if (window.opener && !window.opener.closed) { 
window.opener.location.reload(); 
} 

    
    
}

</script>

<center>



<table width="100%"  border="0" cellpadding="2" cellspacing="0">
			<tr><td width="10%">&nbsp;</td>
			    <td class="plaintextbold">
				<%
				set SL_Ship = objDoc.selectNodes("//preship" + CUSTOM_SHIPPING_FILTER) 
				if SL_Ship.length = 0 then
				%>
				
				No shipping methods are available at this time. 
				<br />
				You will be contacted shortly after your order is submitted to determine the best wasy for us to fulfill your order.				
				<%else%>
					<table width="100%" cellpadding="3" cellspacing="0" border="0">
					 <tr><td  class="THHeader">Shipping Method</td>
					 	<td  class="THHeader">Cost</td>
					  </tr>
					  
					 <tr><td colspan="3" class="plaintextbold" style="padding-top:10px;padding-bottom:15px;">
					 Please click a shipping method from the list below or click the "Close Window" link to return to the checkout page. 
					 </td></tr>
					 <%
					  for x=0 to SL_Ship.length-1
					  SL_ShippingCode =  SL_Ship.item(x).selectSingleNode("shiplist").text
					  SL_ShippingTitle = SL_Ship.item(x).selectSingleNode("ca_title").text
					  SL_Shippingcost = SL_Ship.item(x).selectSingleNode("cost").text
					  
					  %>
					  <tr>
						<td class="plaintext"><a href="Javascript:ApplyThisMethod('<%=SL_ShippingCode%>');"><%=SL_ShippingTitle%></a></td>
					  	  <td class="plaintext"><%=formatcurrency(SL_Shippingcost)%></td>
					</tr>
					  					
					<%next%>
					</table>
				<%
				set SL_Ship= nothing
				end if%>

</td>
</tr>
</table>


<table width="100%">
<tr>
      <td align="center" > 
        <p>&nbsp; </p>
		<p>
         <a href="#" onClick="JavaScript:window.close()"><font face="Arial" size="2" color="#CC0000">Close Window</font></a>
                   
        </p>
        </td>
	</tr>
</table>






<!--#INCLUDE FILE = "googletracking.asp" -->
<!--#INCLUDE FILE = "RemoveXmlObject.asp" -->

</center>

</body>
</html>