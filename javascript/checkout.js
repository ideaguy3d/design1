function RedirectToAddADDress (page)
   {
   var new_url=page;	  
   if (  (new_url != "")  &&  (new_url != null)  )
   //document.basketform.submit();
    window.location=new_url;
   }
function openlearnmore() 
{  
  popupWin = window.open('learnmore.asp', '', 'scrollbars,resizable,width=500,height=450');
	
}
function openlearnmore1() 
{  
  popupWin = window.open('learnmore1.asp', '', 'scrollbars,resizable,width=400,height=350');
	
}
function changeship()
{
	{
	for (var i=0; i < document.checkoutform.txtshiplist.length; i++)
	   {
	   if (document.checkoutform.txtshiplist[i].checked)
    	  {
	      var rad_val = document.checkoutform.txtshiplist[i].value;
		  //alert(rad_val);
		  window.location="changeshipping.asp?shipmethod="+rad_val;
    	  }
	   }
}

}

function MM_openBrWindow() 
{ //v2.0
  
  //var pageurl = 'PreviewShipping.asp?country='+countrycode + "&vstatecode="+statecode + "&vzipcode="+zipcode  ;
  
  var pageurl = 'PreviewShipping.asp' ;
  
  //window.open(pageurl,'estship','location=no,toolbar=no,scrollbars=no,width=550,height=400,top=20,left=20');  
  
     var width  = 500;
     var height = 400;
     var left   = (screen.width  - width)/2;
     var top    = (screen.height - height)/2;
     var params = 'width='+width+', height='+height;
     params += ', top='+top+', left='+left;
     params += ', directories=no';
     params += ', location=no';
     params += ', menubar=no';
     params += ', resizable=no';
     params += ', scrollbars=yes';
     params += ', status=no';
     params += ', toolbar=no';
     newwin=window.open(pageurl,'estship', params);
     //if (window.focus) {newwin.focus()}
     //return false;
  
}

function Contchar(entrytxt,exittxt,texto,maxchars) {
  var entrytxtObj=getObject(entrytxt);
  var exittxtObj=getObject(exittxt);
  var longitud=maxchars - entrytxtObj.value.length;
  if(longitud <= 0) {
    longitud=0;
    texto='<span class="disable"> '+texto+' </span>';
    entrytxtObj.value=entrytxtObj.value.substr(0,maxchars);
  }
  exittxtObj.innerHTML = texto.replace("{CHAR}",longitud);
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
		    xmlhttp.open("POST","SaveCheckoutInfo.asp",false) ;
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
function Process() {

      var p_type = document.getElementById("payment_type").value;	
      if (p_type=="Credit Card")
      {
          var SelectedCardType = document.getElementById('cc_type').selectedIndex ;
          
          
          if (SelectedCardType== 0)
          {
            alert ("Please select Card Type")
		    return false;
          }
          
          
          var myCardNo = document.getElementById("txtcc_number").value;
	      var myCardType = document.getElementById("cc_type").value;
	      var cctypesplit = myCardType.split("/",2)[1]
	      
	      if (myCardNo.length ==0)
	      {
	  	    alert ("Please enter Card Number")
		    return false;    	  
	      }
  		
	      if (cctypesplit.length > 0)
	      {
    	  
	       retval = checkCreditCard (myCardNo,cctypesplit)
	       if (retval==false)
	       {
	        alert ("Invalid Card Number")
		    return false;
	       }	
	      }
      
      }
      
      if (p_type=="PP")
      {
        //save all these values..
	      poststr=""
	      var hasskey=document.getElementById("whatsourcekey");
		  
		  //source key
		  if (hasskey !=null)
		  {
		  	var skeysel = document.getElementById('whatsourcekey').selectedIndex ;   		
		    var skeyselvalue = document.getElementById('whatsourcekey').options[skeysel].value ;
		    poststr ="skeysel=" + encodeURI(skeyselvalue)
		  }
		  else
		  {		  
		   poststr ="skeysel="
		  }
    	 
		 //ship ahead
		 var isRadioPresent = document.getElementById('ship_ahead') ;
		 
		 if (isRadioPresent !=null)
		 {		 
			var ship_aheadRadioLength = checkoutform.ship_ahead.length ;
			j = ship_aheadRadioLength ;			
			 for (k=0; k<j; k++)
			   {
				 radiochecked = checkoutform.ship_ahead[k].checked ;
				 
				 if (radiochecked)
				   {
					radiocheckedValue = checkoutform.ship_ahead[k].value ;
				   }		                                                           
			}
		 }
		 else
		{
			radiocheckedValue=0 ;
		}
		
		//shipping instr
		var hasordermemo = document.getElementById("ordermemo");		
		if (hasordermemo !=null)
		{
			var ordermemoval = hasordermemo.value;
		}
		else
		{
			var ordermemoval=""
		}
		 
		 
		 // order hold date		
		var hasorderhold = document.getElementById("orderhold");
		if (hasorderhold !=null)
		{
			var orderholdval = hasorderhold.value;
		}
		else
		{
			var orderholdval ="" ;
		}
		 
		 
		 //Order comments
		var hasmemofield1 = document.getElementById("memo_field1");
		if (hasmemofield1 !=null)
		{
			var memo_field1Val = hasmemofield1.value ;
		}
		else
		{
			var memo_field1Val ="" ;
		}
		
		var hasmemofield2 = document.getElementById("memo_field2");		
		if (hasmemofield2 !=null)
		{
			var memo_field2Val = hasmemofield2.value ;
		}
		else
		{
			var memo_field2Val ="" ;
		}
		
		var hasmemofield3 = document.getElementById("memo_field3");		
		if (hasmemofield3 !=null)
		{
			var memo_field3Val = hasmemofield3.value ;
		}
		else
		{
			var memo_field3Val ="" ;
		}
		
		var poststr = poststr +
					"&Ship_aheadflag="+ radiocheckedValue +
                    "&orderholddt=" + encodeURI(orderholdval) +
					"&ordermemo1=" + encodeURI(memo_field1Val) +
					"&ordermemo2=" + encodeURI(memo_field2Val) +
					"&ordermemo3=" + encodeURI(memo_field3Val) +
					"&orderfullfill=" + encodeURI(ordermemoval) 		
		
		 //alert(poststr);
		 makePOSTRequest(poststr);
      
	  
      //sandbox
	   //document.checkoutform.action='https://www.sandbox.paypal.com/cgi-bin/webscr'		
		
	   //live
	  document.checkoutform.action='https://www.paypal.com/cgi-bin/webscr'
	  
	  return true ;
	  
      }
      
      return true ;
   }

var panes = new Array();

function setupPanes(containerId, defaultTabId) {
    // go through the DOM, find each tab-container
    // set up the panes array with named panes
    // find the max height, set tab-panes to that height
    panes[containerId] = new Array();
    var maxHeight = 0; var maxWidth = 0;
    var container = document.getElementById(containerId);
    var paneContainer = container.getElementsByTagName("div")[0];
    var paneList = paneContainer.childNodes;
    for (var i=0; i < paneList.length; i++ ) {
        var pane = paneList[i];
        if (pane.nodeType != 1) continue;
        if (pane.offsetHeight > maxHeight) maxHeight = pane.offsetHeight;
        if (pane.offsetWidth  > maxWidth ) maxWidth  = pane.offsetWidth;
        panes[containerId][pane.id] = pane;
        pane.style.display = "none";
    }
    paneContainer.style.height = maxHeight + "px";
    paneContainer.style.width  = maxWidth + "px";
    document.getElementById(defaultTabId).onclick();
}

function showPane(paneId, activeTab) {
    // make tab active class
    // hide other panes (siblings)
    // make pane visible

	var paytype = document.getElementById('payment_type')
	var PaybyPoints = document.getElementById('applypoints')
	var PaybyGC = document.getElementById('applygc')
	
	if (paneId=="pane1")
		{
			paytype.value="Credit Card"
			PaybyPoints.value=0
			PaybyGC.value = 0
		}
			
	if (paneId=="pane2")
		{
			paytype.value="EC"			
			PaybyPoints.value=0
			PaybyGC.value = 0			
		}
	if (paneId=="pane3")
		{
			paytype.value="COD"			
			PaybyPoints.value=0
			PaybyGC.value = 0			
		}		
	if (paneId=="pane4")
		{
			paytype.value="Terms"			
			PaybyPoints.value=0
			PaybyGC.value = 0			
		}
	if (paneId=="pane5")
		{
			paytype.value="PP"			
			PaybyPoints.value=0
			PaybyGC.value = 0			
		}
	if (paneId=="pane6")
		{
			paytype.value="CK"			
			PaybyPoints.value=1
			PaybyGC.value = 0			
		}
	if (paneId=="pane7")
		{
			paytype.value="CK"			
			PaybyPoints.value=0
			PaybyGC.value = 1
		}
		
    if (paneId=="pane8")
		{
			paytype.value="GiftCard"			
			PaybyPoints.value=0
			PaybyGC.value = 0
	}
		

	
    // work out which container this tab is part of. could also bounce on parent node
    for (var con in panes) {
        activeTab.blur();
        activeTab.className = "tab-active";
        if (panes[con][paneId] != null) { // tab and pane are members of this container
            var pane = document.getElementById(paneId);
            pane.style.display = "block";
            var container = document.getElementById(con);
            var tabs = container.getElementsByTagName("ul")[0];
            var tabList = tabs.getElementsByTagName("a")
            for (var i=0; i<tabList.length; i++ ) {
                var tab = tabList[i];
                if (tab != activeTab) tab.className = "tab-disabled";
            }
            for (var i in panes[con]) {
                var pane = panes[con][i];
                if (pane == undefined) continue;
                if (pane.id == paneId) continue;
                pane.style.display = "none"
            }
        }
    }
    return false;    
}
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
function extractNumber(obj, decimalPlaces, allowNegative,beforedecimal)
{
	var temp = obj.value;
	
	// avoid changing things if already formatted correctly
	var reg0Str = '[0-9]*';
	if (decimalPlaces > 0) {
		reg0Str += '\\.?[0-9]{0,' + decimalPlaces + '}';
	} else if (decimalPlaces < 0) {
		reg0Str += '\\.?[0-9]*';
	}
	reg0Str = allowNegative ? '^-?' + reg0Str : '^' + reg0Str;
	reg0Str = reg0Str + '$';
	var reg0 = new RegExp(reg0Str);
	if (reg0.test(temp)) return true;

	// first replace all non numbers
	var reg1Str = '[^0-9' + (decimalPlaces != 0 ? '.' : '') + (allowNegative ? '-' : '') + ']';
	var reg1 = new RegExp(reg1Str, 'g');
	temp = temp.replace(reg1, '');

	if (allowNegative) {
		// replace extra negative
		var hasNegative = temp.length > 0 && temp.charAt(0) == '-';
		var reg2 = /-/g;
		temp = temp.replace(reg2, '');
		if (hasNegative) temp = '-' + temp;
	}
	
	if (decimalPlaces != 0) {
		var reg3 = /\./g;
		var reg3Array = reg3.exec(temp);
		if (reg3Array != null) {
			// keep only first occurrence of .
			//  and the number of places specified by decimalPlaces or the entire string if decimalPlaces < 0
			var reg3Right = temp.substring(reg3Array.index + reg3Array[0].length);
			reg3Right = reg3Right.replace(reg3, '');
			reg3Right = decimalPlaces > 0 ? reg3Right.substring(0, decimalPlaces) : reg3Right;
			temp = temp.substring(0,reg3Array.index) + '.' + reg3Right;
		}
	}
	
	//temp = temp.substring(0,3) ;
	obj.value = temp ;
}
function blockNonNumbers(obj, e, allowDecimal, allowNegative,maxbeforedecimal)
{
	var key;
	var isCtrl = false;
	var keychar;
	var reg;
		
	if(window.event) {
		key = e.keyCode;
		isCtrl = window.event.ctrlKey
	}
	else if(e.which) {
		key = e.which;
		isCtrl = e.ctrlKey;
	}
	
	if (isNaN(key)) return true;
	
	keychar = String.fromCharCode(key);
	
	// check for backspace or delete, or if Ctrl was pressed
	if (key == 8 || isCtrl)
	{
		return true;
	}

	
	reg = /\d/;
	var isFirstN = allowNegative ? keychar == '-' && obj.value.indexOf('-') == -1 : false;
	var isFirstD = allowDecimal ? keychar == '.' && obj.value.indexOf('.') == -1 : false;
	
	retval = true ;
	if (maxbeforedecimal > 0)
	{
		temp = obj.value ;
		
		var hasNegative = temp.length > 0 && temp.charAt(0) == '-';
		var reg2 = /-/g;
		temp = temp.replace(reg2, '');
		
		if (temp.length >= maxbeforedecimal)
		{
			retval=false ;
			return retval ;				
		}
	}

	return isFirstN || isFirstD || reg.test(keychar);

}

/*============================================================================*/



function checkCreditCard (cardnumber, cardname) {
     
  // Array to hold the permitted card characteristics
  var cards = new Array();

  // Define the cards we support. You may add addtional card types.
  
  //  Name:      As in the selection box of the form - must be same as user's
  //  Length:    List of possible valid lengths of the card number for the card
  //  prefixes:  List of possible prefixes for the card
  //  checkdigit Boolean to say whether there is a check digit
  
  cards [0] = {name: "V", 
               length: "13,16", 
               prefixes: "4",
               checkdigit: true};
  cards [1] = {name: "M", 
               length: "16", 
               prefixes: "51,52,53,54,55,22",
               checkdigit: true};
  cards [2] = {name: "C", 
               length: "14,16", 
               prefixes: "300,301,302,303,304,305,36,38,55",
               checkdigit: true};
  cards [3] = {name: "CarteBlanche", 
               length: "14", 
               prefixes: "300,301,302,303,304,305,36,38",
               checkdigit: true};
  cards [4] = {name: "A", 
               length: "15", 
               prefixes: "34,37",
               checkdigit: true};
  cards [5] = {name: "D", 
               length: "16", 
               prefixes: "6011,650,622,624,628,644",
               checkdigit: true};
  cards [6] = {name: "J", 
               length: "15,16", 
               prefixes: "3,1800,2131",
               checkdigit: true};
  cards [7] = {name: "enRoute", 
               length: "15", 
               prefixes: "2014,2149",
               checkdigit: true};
  cards [8] = {name: "Solo", 
               length: "16,18,19", 
               prefixes: "6334, 6767",
               checkdigit: true};
  cards [9] = {name: "Switch", 
               length: "16,18,19", 
               prefixes: "4903,4905,4911,4936,564182,633110,6333,6759",
               checkdigit: true};
  cards [10] = {name: "Maestro", 
               length: "16,18", 
               prefixes: "5020,6",
               checkdigit: true};
  cards [11] = {name: "VisaElectron", 
               length: "16", 
               prefixes: "417500,4917,4913",
               checkdigit: true};
               
  // Establish card type
  var cardType = -1;
  for (var i=0; i<cards.length; i++) {

    // See if it is this card (ignoring the case of the string)
    if (cardname.toLowerCase () == cards[i].name.toLowerCase()) {
      cardType = i;
      break;
    }
  }
  
  // If card type not found, report an error
  if (cardType == -1) {
     ccErrorNo = 0;
     return false; 
  }
   
  // Ensure that the user has provided a credit card number
  if (cardnumber.length == 0)  {
     ccErrorNo = 1;
     return false; 
  }
    
  // Now remove any spaces from the credit card number
  cardnumber = cardnumber.replace (/\s/g, "");
  
  // Check that the number is numeric
  var cardNo = cardnumber
  var cardexp = /^[0-9]{13,19}$/;
  if (!cardexp.exec(cardNo))  {
     ccErrorNo = 2;
     return false; 
  }
       
  // Now check the modulus 10 check digit - if required
  if (cards[cardType].checkdigit) {
    var checksum = 0;                                  // running checksum total
    var mychar = "";                                   // next char to process
    var j = 1;                                         // takes value of 1 or 2
  
    // Process each digit one by one starting at the right
    var calc;
    for (i = cardNo.length - 1; i >= 0; i--) {
    
      // Extract the next digit and multiply by 1 or 2 on alternative digits.
      calc = Number(cardNo.charAt(i)) * j;
    
      // If the result is in two digits add 1 to the checksum total
      if (calc > 9) {
        checksum = checksum + 1;
        calc = calc - 10;
      }
    
      // Add the units element to the checksum total
      checksum = checksum + calc;
    
      // Switch the value of j
      if (j ==1) {j = 2} else {j = 1};
    } 
  
    // All done - if checksum is divisible by 10, it is a valid modulus 10.
    // If not, report an error.
    if (checksum % 10 != 0)  {
     ccErrorNo = 3;
     return false; 
    }
  }  

  // The following are the card-specific checks we undertake.
  var LengthValid = false;
  var PrefixValid = false; 
  var undefined; 

  // We use these for holding the valid lengths and prefixes of a card type
  var prefix = new Array ();
  var lengths = new Array ();
    
  // Load an array with the valid prefixes for this card
  prefix = cards[cardType].prefixes.split(",");
      
  // Now see if any of them match what we have in the card number
  for (i=0; i<prefix.length; i++) {
    var exp = new RegExp ("^" + prefix[i]);
    if (exp.test (cardNo)) PrefixValid = true;
  }
      
  // If it isn't a valid prefix there's no point at looking at the length
  if (!PrefixValid) {
     ccErrorNo = 3;
     return false; 
  }
    
  // See if the length is valid for this card
  lengths = cards[cardType].length.split(",");
  for (j=0; j<lengths.length; j++) {
    if (cardNo.length == lengths[j]) LengthValid = true;
  }
  
  // See if all is OK by seeing if the length was valid. We only check the 
  // length if all else was hunky dory.
  if (!LengthValid) {
     ccErrorNo = 4;
     return false; 
  };   
  
  // The credit card is in the required format.
  return true;
}

/*============================================================================*/