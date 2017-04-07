<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->

<%

if REQUEST.servervariables("server_port_secure")<> 1 and SLsitelive=true then 
        set sitelink=nothing
        set ObjDoc=nothing
    	response.redirect(secureurl+"firsttimeregister.asp") 
  end if

 %>

<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%

 show_baddr3=false
 billcountry = session("country")
 if len(trim(billcountry)) = 0 then
 	billcountry= session("SL_START_COO")
 end if
 
 if billcountry="073" then
	    show_baddr3=true	    
 end if
 
    txtstatelabelTxt="State"    
    if billcountry ="034" then
        txtstatelabelTxt = "Province"
    end if 
	
    txtbZiplabelTxt="Zip Code"

	if billcountry="073" or billcountry="034" then
	    txtbZiplabelTxt="Postal Code"
	end if 



 %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | First Time Customer Registration @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
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
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
<link rel="stylesheet" href="text/footernav.css" type="text/css">
<link rel="stylesheet" href="text/design.css" type="text/css">
</head>
<script language="JavaScript">
function getObject(obj) {
  var theObj;
  if(document.all) {
    if(typeof obj=="string") {
      return document.all(obj);
    } else {
      return obj.style;
    }
  }
  if(document.getElementById) {
    if(typeof obj=="string") {
      return document.getElementById(obj);
    } else {
      return obj.style;
    }
  }
  return null;
}

function Contchar(exittxt,texto) {
  var exittxtObj=getObject(exittxt);
  exittxtObj.innerHTML = texto ;
}

function state_Change()
{
// if xmlhttp shows "loaded"
if (xmlhttp.readyState==4)
  {
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

function Get_allstates(countrycode)
{

xmlhttp=null ;
var new_url="Get_stateonly.asp?code=" + countrycode ;

// code for Mozilla, etc.
if (window.XMLHttpRequest)
  {
  xmlhttp=new XMLHttpRequest()
  }
// code for IE
else if (window.ActiveXObject)
  {
  //xmlhttp=new ActiveXObject("Microsoft.XMLHTTP")
  xmlhttp=new ActiveXObject("Msxml2.XMLHTTP");
  }
if (xmlhttp!=null)
  {
  xmlhttp.onreadystatechange=state_Change
  xmlhttp.open("GET",new_url,false)
  xmlhttp.send(null)
  }
else
  {
  alert("Your browser does not support XMLHTTP.")
  }

} 


function addOption(selectbox,text,value )
{
var optn = document.createElement("OPTION");
optn.text = text;
optn.value = value;
selectbox.options.add(optn);
}


function Fixbcountry(ccode,whichone)
{
    var bbstate = document.getElementById('txtstate')
    var btxtstatelabel = document.getElementById('txtstatelabel')
    var bstlabel = document.getElementById('bstatelabel')
    var bcountylabel =document.getElementById('bukcounty')
    var bcountydrop = document.getElementById('txtcounty')
    var bZiplabel = document.getElementById('bzipcodelabel')
    var bbstate1 = bbstate.length ;
    //var bbstateval = bbstate.value ;    
    // first remove all
    for(i=bbstate1-1;i>=0;i--)
    {
        //bbstateval = bbstate.options[i].value        
        bbstate.remove(i);
       
    }
    var bbaddr3 = document.getElementById('addr3')

    if (ccode=="073")
        {
            bbaddr3.style.display = '';
            btxtstatelabel.style.display='none' ;
            bcountydrop.style.display='' ;
            bcountylabel.style.display='' ;
            bZiplabel.innerHTML="Postal Code" ;
        }
        else
        {
        
         if ((ccode=="001") || (ccode=="034"))
            {//Now add only US/Canada States
            Get_allstates(ccode) ;        
            var answer = xmlhttp.responseText;
            allanswer = answer.split(";");
            listloaded = true 
            for (i=0 ; i < allanswer.length-1 ; i++)
            {
                btstr1 = allanswer[i] ;
                
                btstr2 = btstr1.split("-") ;
                btstcode = btstr2[0] ;
                btstcodeName = btstr2[1] ;
                addOption(bbstate,btstcodeName,btstcode) ;
            }           
            var printthis = ((ccode=="034")?"Province" : "State")
            bstlabel.innerHTML = printthis ;
            }
            else    
            {
                addOption(bbstate,'INTERNATIONAL','INT') ; 
                bstlabel.innerHTML = "State" ;   
            }   
            bbaddr3.style.display = 'none'; 
            btxtstatelabel.style.display='' ;     
            bcountydrop.style.display='none';
            bcountylabel.style.display='none' ;  
            
            var printthis1 = (((ccode=="034")||(ccode=="073"))?"Postal Code" : "Zip Code")
            bZiplabel.innerHTML = printthis1 ;
                               
        }    

}

function CheckForEmpty(inputstr)
{
  if(inputstr.length == 0)
    return true ;
    
  for( var i = 0 ; i <  inputstr.length ; i++)
  {
   var onechar = inputstr.charAt(i) ;
   if( onechar != " ")
     return false ;
 
  }
  
 return true ;
	
}

function validateForm(billtophonereq) {
	
	 //var checkforcountrydrop = document.getElementById('whatbcountrytype')
     
      //if (checkforcountrydrop !=null)
        {
            var bcountrysel = document.getElementById('billtocountry').selectedIndex ;   		
            var bcountryselval = document.getElementById('billtocountry').options[bcountrysel].value ;
        }
       // else
        {
       //     bcountryselval=document.getElementById('billtocountry').value ;
        }

  var haserror=false

	Contchar('bfname','')
	Contchar('blname','')
	Contchar('baddr1','')
	Contchar('bcity','')
	Contchar('bstate','')
	Contchar('bzip','')
	Contchar('bemail', '')
	Contchar('bemailrepeat', '')
	Contchar('bpwd','')
	Contchar('bpwd2','')
	
	
	var fname = document.getElementById('txtfirstname').value ;
	var lname = document.getElementById('txtlastname').value ;
	var addr1 = document.getElementById('txtaddress1').value ;
	var addr2 = document.getElementById('txtaddress2').value ;
	var scity = document.getElementById('txtcity').value ;
	var sstate = document.getElementById('txtstate').value ;
	var szipcode = document.getElementById('txtzipcode').value ;
	var bbphone = document.getElementById('txtphone').value ;
	var semail = document.getElementById('txtemail').value;
	var semailrepeat = document.getElementById('txtemail1').value;
	
	
// Begin Required Field Check
	if (CheckForEmpty(fname))
	{
		Contchar('bfname','(Required)')
		haserror=true		
	}
	if (CheckForEmpty(lname))
	{
		Contchar('blname','(Required)')
		haserror=true		
	}
	if ((CheckForEmpty(addr1)==true) && (CheckForEmpty(addr2)==true))
	{
		Contchar('baddr1','(Required)')
		haserror=true		
	}
	if (CheckForEmpty(scity))
	{
		Contchar('bcity','(Required)')
		haserror=true		
	}
	if ((CheckForEmpty(sstate)) && (bcountryselval !="073"))	
	{
		Contchar('bstate','(Required)')
		haserror=true		
	}
	if (CheckForEmpty(szipcode))
	{
		Contchar('bzip','(Required)')
		haserror=true		
	}
	if (CheckForEmpty(bbphone)&& billtophonereq==1)
	{
		Contchar('bphone','(Required)')
		haserror=true		
	}
	
	
	if (CheckForEmpty(semail))
	{
		Contchar('bemail','(Required)')		
		haserror=true
	}
	if (CheckForEmpty(semailrepeat)) {
	    Contchar('bemailrepeat', '(Required)')
	    haserror = true
	}
	
	//var bbphone = custinfoform.txtphone.value ;
	
	//var regex = /\s+/;
	//bbphone = bbphone.replace(regex,'')
	bbphone = bbphone.replace(/ /g,'')
	bbphone = bbphone.replace(/-/g,'')
	bbphone = bbphone.replace(/\(/g,'')
	bbphone = bbphone.replace(/\)/g,'')
	
	//bbphone = bbphone.replace(/^\s+|\s+$/g,"");
	if (bbphone.length > 0 )
	{
    	regex =/^[0-9]+$/ ;
    	testresult = regex.test(bbphone) ;    	    
    	if (testresult==false)
    	{
    	    Contchar('bphone','(Invalid Phone Number)')
		    haserror=true	
    	}
    	else
    	{
    	    Contchar('bphone','')
    	}
        
	}
	
	if (semail.length != 0 && !/^[a-zA-Z0-9._-]+@([a-zA-Z0-9.-]+\.)+[a-zA-Z0-9.-]{2,4}$/.test(semail))
	{
		Contchar('bemail','Invalid Email Address')		
		haserror=true
	}

	if ((CheckForEmpty(semail)==false) && (semail.toUpperCase() != semailrepeat.toUpperCase())) {
	    Contchar('bemailrepeat', '(E-Mail and Re-enter E-Mail do not match)')
	    haserror = true
	}
	
	var bbpwd1 = document.getElementById('txtregpassword').value ;
	var bbpwd2 = document.getElementById('txtregpassword2').value ;
		 
				
			if (bbpwd1.length < 4)
			{
				Contchar('bpwd','\(password must be at least 4 characters)')
				haserror=true
			}
			if (bbpwd1 !== bbpwd2)
			{
				Contchar('bpwd2','(Both passwords do not match; please retype your password in both fields)')
				haserror=true
			}	
		
		

	
	if (haserror)
		return false ;
	
	 return true ;
}
</script>
<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">



<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
<br>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%">Registration</td></tr>
	</table>
	<br>
	<form name="custinfoform" method="POST" onSubmit="return validateForm(<%=BILL_TOPHONE_REQUIRED%>)" action="<%=secureurl%>FirstTimeRegistration.asp">
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
		<tr><td colspan="4" align="center" class="THHeader">Customer Address &nbsp;(* Required )</td></tr>
		<tr><td class="plaintextbold" align="right"  width="160"> First Name </td>
								<td class="smalltextred"><input type="text" name="txtfirstname" id="txtfirstname" size="20,1" maxlength="15" value="<% =session("firstname") %>">&nbsp;*
								<span id="bfname"></span>
								<input type="hidden" name="registered"  value="0">
								</td>
							</tr>
							<tr>
								<td class="plaintextbold" align="right"> Last Name </td>
								<td class="smalltextred"><input type="text" name="txtlastname" id="txtlastname" size="30,1" maxlength="20" value="<% =session("lastname") %>">&nbsp;*
								<span id="blname"></span>
								</td>
							</tr>
							<tr><td class="plaintextbold" align="right"> Company </td>
								<td><input type="text" name="txtcompany" id="txtcompany" size="40,1" maxlength="40" value="<% =session("company") %>"></td>
							</tr>
							<tr><td class="plaintextbold" align="right"> Address </td>
								<td class="smalltextred"><input type="text" name="txtaddress1" id="txtaddress1" size="40,1" maxlength="40" value="<% =session("address1") %>">&nbsp;*
								<span id="baddr1"></span>
								</td>
							</tr>
							<tr><td class="plaintextbold" align="right"> Apt. or Suite </td>
								<td><input type="text" name="txtaddress2" id="txtaddress2" size="40,1" maxlength="40" value="<% =session("address2") %>"></td>
							</tr>
							
							
							
							<tr id="addr3" <%if show_baddr3=true then %>style="display:''"<%else%>style="display:none"<%end if%> >
								<td class="plaintextbold" align="right"> Address 3 </td>
								<td ><input type="text" name="txtaddress3" size="40,1" maxlength="40" value="<% =session("address3") %>"></td>
							</tr>
							
							<tr><td class="plaintextbold" align="right"> City </td>
								<td class="smalltextred"><input type="text" name="txtcity" id="txtcity" size="25,1" maxlength="20" value="<% =session("city") %>">&nbsp;*
								<span id="bcity"></span>
								</td>
								</tr>
								<tr id="txtstatelabel" <%if show_baddr3=true then %>style="display:none"<%else%>style="display:''"<%end if%>>
								<td class="plaintextbold" align="right"><span id="bstatelabel"><%=txtstatelabelTxt%></span> </td>
											 <td class="smalltextred">
											  <select name="txtstate" id="txtstate" class="smalltextblk">
												  <option value="">
												  <option value="INT" <%if session("state") = "INT" then response.write(" selected") end if%>>INTERNATIONAL																		  
												  <%
												  xmlstring =sitelink.Get_state_list(session("ordernumber"))
												  objDoc.loadxml(xmlstring)
												  tnode = "//state_list[ccode='" + cstr(billcountry) +"']"
												  set SL_Ship_state = objDoc.selectNodes(tnode)
													for x=0 to SL_Ship_state.length-1
														SL_state_name= SL_Ship_state.item(x).selectSingleNode("statename").text
														SL_state_code= SL_Ship_state.item(x).selectSingleNode("sstate").text														
												  %>
												  <option value="<%=SL_state_code%>"
												  <%if session("state") = SL_state_code then response.write(" selected") end if%>
												  ><%=SL_state_name%>
												  <%next						  
												  set SL_Ship_state=nothing
												  %>
												  </select>&nbsp;*
												  <span id="bstate"></span>
								 </td>
							</tr>
							
							<tr id="bukcounty" <%if show_baddr3=true then %>style="display:''"<%else%>style="display:none"<%end if%>><td class="plaintextbold" align="right">County</td>
								<td class="plaintextbold">
								
									 <select name="txtcounty" id="txtcounty" class="smalltextblk">												 
												  <%
												  xmlstring =sitelink.Get_County_list(session("ordernumber"))
												  objDoc.loadxml(xmlstring)
												  set SL_County_List = objDoc.selectNodes("//county_list")
													for x=0 to SL_County_List.length-1
														SL_county_name= SL_County_List.item(x).selectSingleNode("countyname").text
														SL_county_Code= SL_County_List.item(x).selectSingleNode("ccode").text													
												  %>
												  <option value="<%=SL_county_Code%>"
												  <%if session("bcounty") = SL_county_Code then response.write(" selected") end if%>
												   ><%=SL_county_name%>												   
												  <%next						  
												    set SL_County_List =nothing
												  %>
												  </select>
							</td>
							</tr>
							
							<tr><td class="plaintextbold" align="right"><span id="bzipcodelabel"><%=txtbZiplabelTxt%></span></td>
								<td class="smalltextred">
									<input type="text" name="txtzipcode" id="txtzipcode" size="10,1" maxlength="10" value="<% =session("zipcode") %>">&nbsp;*
									<span id="bzip"></span>
							</td>
							</tr>
							<tr><td class="plaintextbold" align="right">Country </td>
								<td><input type="hidden" id="whatbcountrytype" name="whatbcountrytype" value="1">				
								<select NAME="billtocountry" id="billtocountry" class="smalltextblk" onChange="Fixbcountry(this.options[this.selectedIndex].value,'B')">
									<%'=sitelink.listcountries(session("ordernumber"),session("shopperid"),"B",0)%>
									<%=sitelink.GET_COUNTRYLIST(billcountry)%>
								</select>
								</td>
							</tr>
							<tr><td class="plaintextbold" align="right"> Phone</td>
								<td class="smalltextred"><input type="text" name="txtphone" id="txtphone" size="20,1" maxlength="20" value="<% =session("phone") %>">
								<%if BILL_TOPHONE_REQUIRED=1 then response.write("&nbsp;*") end if%>
								<span id="bphone"></span>
								</td>
							</tr>
							
		<tr><td class="plaintextbold" align="right"> &nbsp;E-Mail&nbsp; </td>
			<td colspan="3" class="smalltextred"><input type="text" name="txtemail" id="txtemail" size="50,1" maxlength="50" value="<% =session("email") %>">&nbsp;*
			<span id="bemail"></span>
			<input type="hidden" name="sitestore" value="SITELINK">
		</td>
		</tr>
		<tr><td class="plaintextbold" align="right"> &nbsp;Re-enter E-Mail&nbsp; </td>
			<td colspan="3" class="smalltextred"><input type="text" name="txtemail1" id="txtemail1" size="50,1" maxlength="50" value="<% =session("email") %>">&nbsp;*
			<span id="bemailrepeat"></span>
		</td>
		</tr>
		<tr><td class="plaintextbold" align="right">&nbsp;Password&nbsp;</td>
			<td colspan="3" class="smalltextred"><input type="password" name="txtregpassword" id="txtregpassword" size="15" maxlength="10" value="<% =session("regpassword") %>">
                *
				<span id="bpwd"></span>
				</td>
		</tr>
			<tr><td class="plaintextbold" align="right">&nbsp;Re-Enter Password&nbsp;</td>
				<td colspan="3" class="smalltextred"><input type="password" name="txtregpassword2" id="txtregpassword2" size="15" maxlength="10" value="<%=session("regpassword2")%>">
                *
				<span id="bpwd2"></span>				
				</td>
		</tr>
		
		<tr><td>&nbsp;<br><br></td>
			<td valign="bottom"><input type="Image" value="Register" src="images/btn-register.gif" border="0" alt="Register"></td>
		</tr>

	</table>
	</form>
	
	
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

<script type="text/javascript">
  (function() {
    window._pa = window._pa || {};
    // _pa.orderId = "myOrderId"; // OPTIONAL: attach unique conversion identifier to conversions
    // _pa.revenue = "19.99"; // OPTIONAL: attach dynamic purchase values to conversions
    // _pa.productId = "myProductId"; // OPTIONAL: Include product ID for use with dynamic ads
    var pa = document.createElement('script'); pa.type = 'text/javascript'; pa.async = true;
    pa.src = ('https:' == document.location.protocol ? 'https:' : 'http:') + "//tag.marinsm.com/serve/55c94a54db45c07a90000003.js";
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(pa, s);
  })();
</script>
</body>
</html>
