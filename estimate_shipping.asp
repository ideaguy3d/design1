<%on error resume next%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%

billcountry = session("country")
if len(trim(billcountry)) = 0 then
	 	billcountry= session("SL_START_COO")
end if
	 

If Request.TotalBytes > 0 Then
        vtextshipcountry = REQUEST.FORM("listcountry")
        vtxtstate  = ""
        vtxtzipCode = ucase(REQUEST.FORM("txtzipCode"))
        
        shipping_order="3"				
		shipping_order = cstr(SHIPPING_SORTBY) 
		if SHIPPING_SORTORDER=1 then
			shipping_order = shipping_order +" desc"
		end if 
		
		'default shipping method will be the ship_via that appears most on line item
		session("shippingmethod")=""
		'xmlstring =sitelink.GET_MOST_BASKETSHIP_VIA(session("ordernumber"))
		'objDoc.loadxml(xmlstring)
		'set ship_via = objDoc.selectNodes("//result") 
		'	ALL_Nodes_With_Ship_Via= ship_via.length
		'	if ship_via.length > 0 then
		'		    shippingmethod = ship_via.item(0).selectSingleNode("ship_via").text
		'		    if shippingmethod<>"" then
		'		        session("shippingmethod")=shippingmethod
		'		    end if					
		'	end if	
		'set ship_via =nothing

	
		'if multiship..
		'if ALL_Nodes_With_Ship_Via > 0 then
		'	xmlstring=sitelink.GET_PROMO(session("ordernumber"))
		'	objDoc.loadxml(xmlstring)
		'	set SL_Shippingcode = objDoc.selectNodes("//result")
		'		if SL_Shippingcode.length>0 then
		'			session("shippingmethod")=SL_Shippingcode.item(0).selectSingleNode("ship_via").text
		'		end if
		'	set SL_Shippingcode =nothing
		'
		'end if 

		'if session("shippingmethod") = "" then
		'get a default shipping method first
		'	xmlCarrier = sitelink.SHIPPINGMETHODS() 
		'	objDoc.loadxml(xmlCarrier)
		'				
		'	target_node = "//shipmeths[ca_code='"+cstr(DEFAULT_SHIPPING_METHOD)+"']"
		'	
		'	set SL_Ship = objDoc.selectNodes(target_node) 
		'		if SL_Ship.length > 0 then
		'			session("shippingmethod") =DEFAULT_SHIPPING_METHOD
		'		end if 
		'	set SL_Ship= nothing
			
		'	if session("shippingmethod")="" then			  
		'		set SL_Ship = objDoc.selectNodes("//shipmeths") 
		'		if SL_Ship.length > 0 then
		'			session("shippingmethod") = SL_Ship.item(0).selectSingleNode("ca_code").text		
		'		end if
		'		set SL_Ship= nothing
		'	end if 
		
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
		'end if
		
		call sitelink.setzipcode(session("shopperid"),cstr(vtextshipcountry),"",cstr(vtxtzipCode))
		xmlstring=sitelink.previewshipping(session("shopperid"),session("ordernumber"),cstr(vtextshipcountry),UCase(cstr(vtxtstate)),cstr(vtxtzipCode),shipping_order,session("shippingmethod"))						
		'call sitelink.TOTALORDER(session("shopperid"),session("ordernumber"),session("billtocopy"),session("shippingmethod"),false)		
						
		objDoc.loadxml(xmlstring)	
	end if
	
	'response.Write("Shopper-->"& session("shopperid"))
	'response.Write("<br>")
	'response.Write("Order-->"& session("ordernumber"))

%>


<html>
<head>
<title><%=althomepage%></title>
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

</script>

<center>

<form action name="est_shipping" onSubmit="return Process()" method="post">
<table width="80%" cellpadding="3" border="0" cellspacing="3">
<tr>
      <td class="plaintextbold">Country</td>
      <td> 
	  <select NAME="Listcountry" class="smalltextblk">
				<%=sitelink.GET_COUNTRYLIST(billcountry)%>
			</select>        
		</td> 
	</tr>
	<tr>
	  <td class="plaintextbold">Zip&nbsp;Code</td>
	  <td> 
        <input type="text" name="txtzipcode" size="10,1" maxlength="10" value="">
		&nbsp;&nbsp;&nbsp;<input type="submit" value="Submit">
		</td>
	</tr>	

</table>
</form>
<%if Request.TotalBytes > 0 then %>
<table width="100%"  border="0" cellpadding="2" cellspacing="0">
			<tr><td width="10%">&nbsp;</td>
			    <td class="plaintextbold">
				<%
				set SL_Ship = objDoc.selectNodes("//preship" + CUSTOM_SHIPPING_FILTER) 
				if SL_Ship.length = 0 then
				%>
				
				No Shipping Methods available..				
				<%else%>
					<table width="100%" cellpadding="3" cellspacing="0" border="0">
					 <tr><td  class="THHeader">Shipping Method</td>
					 	<td  class="THHeader">Cost</td>
					  </tr>
					 <%
					  for x=0 to SL_Ship.length-1
					  SL_ShippingCode =  SL_Ship.item(x).selectSingleNode("shiplist").text
					  SL_ShippingTitle = trim(SL_Ship.item(x).selectSingleNode("ca_title").text)
					  SL_Shippingcost = SL_Ship.item(x).selectSingleNode("cost").text
					  
					  %>
					  <tr>
						<td class="plaintext">
						<%=SL_ShippingTitle%>
						</td>
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
<%end if%>

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