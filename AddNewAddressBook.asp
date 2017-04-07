<%on error resume next%>

<%
  if session("registeredshopper")="NO" then
  	session("destpage")="AddNewAddressBook.asp"
  	response.redirect("login.asp")
  end if 
  
  	basketrecord = request.querystring("basketrecord")
  	if len(trim(basketrecord))= 0 then
  		basketrecord=0	
	 end if
  
	 if isnumeric(basketrecord)=false then	
	 	Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl&"default.asp"
		Response.End
	  end if

  show_baddr3=false
 shipcountry = session("country")
 if len(trim(shipcountry)) = 0 then
 	shipcountry= session("SL_START_COO")
 end if
 
 if shipcountry="073" then
	   show_saddr3=true	    
 end if
 
    txtsstatelabelTxt="State"    
    if shipcountry ="034" then
       txtsstatelabelTxt = "Province"
    end if 
	
    txtsZiplabelTxt="Zip Code"

	if shipcountry="073" or shipcountry="034" then
	   txtsZiplabelTxt="Postal Code"
	end if 

%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->






<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Add New Address to Address Book | Australian Products Co.</title>
<meta name="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls" />
	<meta name="keywords" content="Aussie Products, Australian, Aussie, Australian Products, Australian Food, Aussie Foods, Australian Desserts, tim tams, vegemite, violet crumble, cherry ripe, koala, kangaroo, down under, Australian Groceries, foods, lollies, candy, gourmet food" />
	<meta name="robots" content="index, follow" />
    <meta name="robots" content="ALL">
	<meta name="distribution" content="Global">
	<meta name="revisit-after" content="5">
	<meta name="classification" content="Food">	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
<link rel="stylesheet" href="text/footernav.css" type="text/css">
<link rel="stylesheet" href="text/design.css" type="text/css">
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
    var bbstate = document.getElementById('txtsstate')    
    var stxtstatelabel = document.getElementById('txtsstatelabel')
        
    var sstlabel = document.getElementById('sstatelabel')
    
    var scountylabel =document.getElementById('sukcounty')
    var scountydrop = document.getElementById('txtscounty')
    var sZiplabel = document.getElementById('szipcodelabel')
    
     var bbstate1 = bbstate.length ;
    //var bbstateval = bbstate.value ;    
    // first remove all
    for(i=bbstate1-1;i>=0;i--)
    {
        //bbstateval = bbstate.options[i].value        
        bbstate.remove(i);
       
    }
    var sbaddr3 = document.getElementById('saddr3')    
    if (ccode=="073")
        {
            sbaddr3.style.display = '';
            stxtstatelabel.style.display='none' ;
            //bcountydrop.style.display='';
            //bcountylabel.style.display='' ;
            scountydrop.style.display='';
            scountylabel.style.display='' ;
            sZiplabel.innerHTML="Postal Code" ;
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
            sstlabel.innerHTML = printthis ;
            }
            else    
            {
                addOption(bbstate,'INTERNATIONAL','INT') ;   
                sstlabel.innerHTML = "State" ;   
            }   
            sbaddr3.style.display = 'none'; 
            stxtstatelabel.style.display='' ; 
            //bcountydrop.style.display='none';
            //bcountylabel.style.display='none';   
            scountydrop.style.display='none';
            scountylabel.style.display='none' ;   
            var printthis1 = (((ccode=="034")||(ccode=="073"))?"Postal Code" : "Zip Code")
            sZiplabel.innerHTML = printthis1 ;                   
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

function validateForm(ShiptoEmailreq) {
	

  var haserror=false
 // var regex = /^(?:\([2-9]\d{2}\)\ ?|[2-9]\d{2}(?:\-?|\ ?))[2-9]\d{2}[- ]?\d{4}$/
 // var regex2 = /^(?:\([2-9]\d{2}\)\ ?|[2-9]\d{2}(?:\-?|\ ?))[2-9]\d{2}[- ]?\d{4}$/
 // var reqfield = "The following fields are required: "

    // var checkforcountrydrop = document.getElementById('whatbcountrytype')
     
   //   if (checkforcountrydrop !=null)
   //     {
   //         var bcountrysel = document.getElementById('billtocountry').selectedIndex ;   		
   //         var bcountryselval = document.getElementById('billtocountry').options[bcountrysel].value ;
   //     }
   //     else
   //     {
   //         bcountryselval=document.getElementById('billtocountry').value ;
   //     }
  
	
     scountrysel = document.getElementById('shiptocountry').selectedIndex ;
     scountryselval = document.getElementById('shiptocountry').options[scountrysel].value ;   


	Contchar('sfname','')
	Contchar('slname','')
	Contchar('saddr1','')
	Contchar('scity','')
	Contchar('sstate','')
	Contchar('szip','')
	Contchar('semail', '')
	Contchar('semailrepeat', '')

	// Begin Required Field Check  
	var fname = document.getElementById('txtsfirstname').value ;
	var lname = document.getElementById('txtslastname').value ;
	var addr1 = document.getElementById('txtsaddress1').value ;
	var addr2 = document.getElementById('txtsaddress2').value ;
	var scity = document.getElementById('txtscity').value ;
	var sstate = document.getElementById('txtsstate').value ;
	var szipcode = document.getElementById('txtszipcode').value ;
	var semail = document.getElementById('txtsemail').value;
	var semailrep = document.getElementById('txtsemail1').value; 
    
		if (CheckForEmpty(fname))
		{
			Contchar('sfname','(Required)')
			haserror=true		
		}
		if (CheckForEmpty(lname))
		{
			Contchar('slname','(Required)')
			haserror=true		
		}
		
		if ((CheckForEmpty(addr1)==true) && (CheckForEmpty(addr2)==true))
		{
			Contchar('saddr1','(Required)')
			haserror=true		
		}
		if (CheckForEmpty(scity))
		{
			Contchar('scity','(Required)')
			haserror=true		
		}
		
		if ((CheckForEmpty(sstate)) && (scountryselval !="073"))
		{
			Contchar('sstate','(Required)')
			haserror=true		
		}
		if (CheckForEmpty(szipcode))
		{
			Contchar('szip','(Required)')
			haserror=true		
		}
		
		
		if ((CheckForEmpty(semail)) && (ShiptoEmailreq==1))
		{
			Contchar('semail','(Required)')
			haserror=true		
		}
		
		
		/* if (ShiptoEmailreq==1)
		{ */
		    if (semail.length != 0 && !/^[a-zA-Z0-9._-]+@([a-zA-Z0-9.-]+\.)+[a-zA-Z0-9.-]{2,4}$/.test(semail))
	        {
		        Contchar('semail','Invalid Email Address')		
		        haserror=true
	        }
	    ///}

	    if ((CheckForEmpty(semail) == false) && (semail.toUpperCase() != semailrep.toUpperCase())) {
	        Contchar('semailrepeat', '(E-Mail and Re-enter E-Mail do not match)')
	        haserror = true
	    } 
	    
	
	if (haserror)
		return false ;
	
	 return true ;
	 
}
</script>
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">



<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
<%if WANT_SIDENAV=1 then%>
	<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top">
	<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	
	<%end if%>
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<br>
	<table width="100%"  cellpadding="3" cellspacing="0" border="0">
		<tr><td class="TopNavRow2Text" width="100%" >Add New Address&nbsp;&nbsp; </td></tr>
	</table>
	<br>
	<form name="custinfoform" method="POST" onSubmit="return validateForm(<%=SHIP_TOEMAIL_REQUIRED%>)" action="Updateaddress.asp">
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
		<tr><td>
			  <input type="hidden" name="addtobook" value="1">
			  <input type="hidden" name="basketrecord" value="<%=basketrecord%>">
		</td></tr>
		<tr><td colspan="4" align="center"  class="THHeader">Customer Address &nbsp;(* Required )</td></tr>
		<tr><td class="plaintextbold" align="right" width="160"> &nbsp;First Name&nbsp; </td>
			<td class="smalltextred"><input type="text" name="txtsfirstname" id="txtsfirstname" size="20,1" maxlength="15" value="">&nbsp;*
			<span id="sfname"></span>
			</td>
		</tr>
		<tr>
            <td class="plaintextbold" align="right"> &nbsp;Last Name&nbsp; </td>
			<td class="smalltextred"><input type="text" name="txtslastname" id="txtslastname" size="30,1" maxlength="20" value="">&nbsp;*
			<span id="slname"></span>
			</td>
		</tr>
		<tr><td class="plaintextbold" align="right"> &nbsp;Company&nbsp; </td>
			<td colspan="3"><input type="text" name="txtscompany" id="txtscompany" size="50,1" maxlength="40" value=""></td>
		</tr>
		<tr><td class="plaintextbold" align="right"> &nbsp;Address&nbsp; </td>
			<td colspan="3" class="smalltextred"><input type="text" name="txtsaddress1" id="txtsaddress1" size="50,1" maxlength="40" value="">&nbsp;*
			<span id="saddr1"></span>
			</td>
		</tr>
		<tr><td class="plaintextbold" align="right"> &nbsp;Address 2&nbsp; </td>
			<td colspan="3"><input type="text" name="txtsaddress2" id="txtsaddress2" size="50,1" maxlength="40" value=""></td>
		</tr>
			<tr id="saddr3"  <%if show_saddr3=true then %>style="display:''"<%else%>style="display:none"<%end if%> >
		<td class="plaintextbold" align="right"> Address 3 </td>
				<td ><input type="text" name="txtsaddress3" size="40,1" maxlength="40" value=""></td>
		</tr>
		<tr><td class="plaintextbold" align="right"> &nbsp;City&nbsp; </td>
			<td colspan="3" class="smalltextred">
		    <input type="text" name="txtscity" id="txtscity" size="25,1" maxlength="20" value="">&nbsp;*
			<span id="scity"></span>
			</td>
		</tr>
		<tr id="txtsstatelabel" <%if show_saddr3=true then %>style="display:none"<%else%>style="display:''"<%end if%>>
							<td class="plaintextbold" align="right"><span id="sstatelabel"><%=txtsstatelabelTxt%></span></td>
							<td class="smalltextred">
								<select name="txtsstate" id="txtsstate"  class="smalltextblk">
									  
									  <option value="">
									  <option value="INT" <%if session("state") = "INT" then response.write(" selected") end if%>>INTERNATIONAL
									  <%
									  xmlstring =sitelink.Get_state_list(session("ordernumber"))
									  objDoc.loadxml(xmlstring)
									  tnode = "//state_list[ccode='" + cstr(shipcountry) +"']"
									  set SL_Ship_state = objDoc.selectNodes(tnode)
										for x=0 to SL_Ship_state.length-1
											SL_state_name= SL_Ship_state.item(x).selectSingleNode("statename").text
											SL_state_code= SL_Ship_state.item(x).selectSingleNode("sstate").text
									  %>
									  <option value="<%=SL_state_code%>"
									  <%if session("sstate") = SL_state_code then response.write(" selected") end if%>
									  ><%=SL_state_name%>
									  <%next						  
									  set SL_Ship_state=nothing
									  %>
									  </select>&nbsp;*
									  <span id="sstate"></span>
							 </td>
							</tr>
							<tr id="sukcounty" <%if show_saddr3=true then %>style="display:''"<%else%>style="display:none"<%end if%>>
							<td class="plaintextbold" align="right">County</td>
								<td class="plaintext">
									
									
									 <select name="txtscounty" id="txtscounty" class="smalltextblk">												 
												  <%
												     xmlstring =sitelink.Get_County_list(session("ordernumber"))
												    objDoc.loadxml(xmlstring)
												    set SL_County_List = objDoc.selectNodes("//county_list")
													for x=0 to SL_County_List.length-1
														SL_county_name= SL_County_List.item(x).selectSingleNode("countyname").text
														SL_county_Code= SL_County_List.item(x).selectSingleNode("ccode").text													
												  %>
												  
												  <option value="<%=SL_county_Code%>"
												  <%if session("scounty") = SL_county_Code then response.write(" selected") end if%>
												  ><%=SL_county_name%>
												  <%next						  
												    set SL_County_List = nothing
												  %>
												  </select>
							</td>
							</tr>
							<tr><td class="plaintextbold" align="right"><span id="szipcodelabel"><%=txtsZiplabelTxt%></span>
							  <td class="smalltextred">
							  <input type="text" name="txtszipcode" id="txtszipcode" size="10,1" maxlength="10" value="<% =session("szipcode") %>">&nbsp;*
							  <span id="szip"></span>
							</td>
						</tr>
						<tr><td class="plaintextbold" align="right">Country </td>
							<td>
								<select NAME="shiptocountry" id="shiptocountry" class="smalltextblk" onChange="Fixbcountry(this.options[this.selectedIndex].value,'S')">
									<%'=sitelink.listcountries(session("ordernumber"),session("shopperid"),"S",session("address_id"))%>
									<%=sitelink.GET_COUNTRYLIST(shipcountry)%>
								</select>
							</td>
						</tr>
    	<tr><td class="plaintextbold" align="right"> &nbsp;Phone&nbsp; </td>
			<td colspan="3"><input type="text" name="txtsphone" id="txtsphone" size="20,1" maxlength="14" value=""></td>
		</tr>
		<tr><td class="plaintextbold" align="right"> &nbsp;E-Mail&nbsp; </td>
			<td colspan="3" class="smalltextred"><input type="text" name="txtsemail" id="txtsemail"  size="50,1" maxlength="50" value="">
			<%if SHIP_TOEMAIL_REQUIRED=1 then response.write("&nbsp;*") end if%>
			<span id="semail"></span>
			</td>
		</tr>
		<tr><td class="plaintextbold" align="right"> &nbsp;Re-Enter E-Mail&nbsp; </td>
			<td colspan="3" class="smalltextred"><input type="text" name="txtsemail1" id="txtsemail1"  size="50,1" maxlength="50" value="">
			<span id="semailrepeat"></span>
			</td>
		</tr>

		
		<tr><td>&nbsp;<br><br></td>
			<td valign="bottom">
			<input type="image" src="images/btn_Add_address.gif" border="0" Alt="Add to Address Book">
			</td>
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

</body>
</html>
