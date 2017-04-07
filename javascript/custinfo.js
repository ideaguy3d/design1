function modifyTextField(field, bDisabled) {
	field.disabled = bDisabled
	if (bDisabled) {
		field.onfocus = function() {this.blur(); };	
	}
	else {
		field.onfocus = function() {};			
	}
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


function ConvertAllFields(bClearFormIfNecessary, lstartcoo) {



 var checkforcountrydrop = document.getElementById('whatbcountrytype')
 
 //when user is logged in there is not drop down
 
 if (checkforcountrydrop !=null)
  {
      var billcountrysel = document.getElementById('billtocountry').selectedIndex ;   		
      var billcountryvalue = document.getElementById('billtocountry').options[billcountrysel].value ;
  }
  else
  {
   billcountryvalue=document.getElementById('billtocountry').value ;
  }
  
   var ccode = billcountryvalue
   var btxtstatelabel = document.getElementById('txtstatelabel')
   var stxtstatelabel = document.getElementById('txtsstatelabel')
    
   var bstlabel = document.getElementById('bstatelabel')
   var sstlabel = document.getElementById('sstatelabel')
    
   var bcountylabel =document.getElementById('bukcounty')
   var bcountydrop = document.getElementById('txtcounty')
    
   var scountylabel =document.getElementById('sukcounty')
   var scountydrop = document.getElementById('txtscounty')
    
   var bbaddr3 = document.getElementById('addr3')
   var sbaddr3 = document.getElementById('saddr3')  
   
    var bZiplabel = document.getElementById('bzipcodelabel')
    var sZiplabel = document.getElementById('szipcodelabel')
    
            
  if (document.custinfoform.billtocopy.checked) {

    document.custinfoform.billtocopy.value = "1"
	document.custinfoform.txtsfirstname.value = document.custinfoform.txtfirstname.value
	document.custinfoform.txtslastname.value = document.custinfoform.txtlastname.value
	document.custinfoform.txtscompany.value = document.custinfoform.txtcompany.value
	document.custinfoform.txtsaddress1.value = document.custinfoform.txtaddress1.value
	document.custinfoform.txtsaddress2.value = document.custinfoform.txtaddress2.value
	document.custinfoform.txtsaddress3.value = document.custinfoform.txtaddress3.value
	document.custinfoform.txtscity.value = document.custinfoform.txtcity.value
	document.custinfoform.txtscounty.value = document.custinfoform.txtcounty.value
	document.custinfoform.txtszipcode.value = document.custinfoform.txtzipcode.value	
	document.custinfoform.shiptocountry.value = document.custinfoform.billtocountry.value
	document.custinfoform.txtsphone.value = document.custinfoform.txtphone.value
	document.custinfoform.txtsemail.value = document.custinfoform.txtemail.value
	document.custinfoform.txtsemail1.value = document.custinfoform.txtemail.value
	
	if (ccode=="073")
        {
            sbaddr3.style.display = '';
            stxtstatelabel.style.display='none' ;
            scountydrop.style.display='';
            scountylabel.style.display='' ;
        }
    else
    {
        sbaddr3.style.display = 'none';
        stxtstatelabel.style.display='' ;
        scountydrop.style.display='none';
        scountylabel.style.display='none' ;
    
    }

    var printthis = ((ccode=="034")?"Province" : "State")
    sstlabel.innerHTML = printthis ;
    
    //copy all fields from state/county
    if (checkforcountrydrop !=null)
    {
        
        var sbstate = document.getElementById('txtsstate')
        var sbstate1 = sbstate.length ;
        for(i=sbstate1-1;i>=0;i--)
        {        
            sbstate.remove(i);       
        }
        var bbstate = document.getElementById('txtstate')
        var bbstate1 = bbstate.length ;
        
        for(i=0; i <= bbstate1-1 ; i++)
        {            
           btstcodeName= bbstate.options[i].text;
           btstcode= bbstate.options[i].value;
           addOption(sbstate,btstcodeName,btstcode) ; 
            
        }
        //alert("hhh");
        document.custinfoform.txtsstate.value = document.custinfoform.txtstate.value	
	}
	else
	{
	    if ((ccode=="001") || (ccode=="034"))
            {//Now add only US/Canada States
            var sbstate = document.getElementById('txtsstate')
            var sbstate1 = sbstate.length ;
             for(i=sbstate1-1;i>=0;i--)
               {        
                    sbstate.remove(i);       
                }
            
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
                addOption(sbstate,btstcodeName,btstcode) ;
            }           
            var printthis = ((ccode=="034")?"Province" : "State")
            sstlabel.innerHTML = printthis ;            
            document.custinfoform.txtsstate.value = document.custinfoform.txtstate.value
            }
           else    
            {
            
                if (ccode!="073")
                {
                var sbstate = document.getElementById('txtsstate');
                var sbstate1 = sbstate.length ;
                for(i=sbstate1-1;i>=0;i--)
               {        
                    sbstate.remove(i);       
                }
                addOption(sbstate,'INTERNATIONAL','INT') ; 
                sstlabel.innerHTML = "State" ;                  
                }
            }   
            //bbaddr3.style.display = 'none'; 
            //btxtstatelabel.style.display='' ;     
            //bcountydrop.style.display='none';
            //bcountylabel.style.display='none' ;  
            
            var printthis1 = (((ccode=="034")||(ccode=="073"))?"Postal Code" : "Zip Code")
            sZiplabel.innerHTML = printthis1 ;
	}

	EnableAllFields(true)
  }
  else if (bClearFormIfNecessary) {
	document.custinfoform.txtsfirstname.value = ""
	document.custinfoform.txtslastname.value = ""
	document.custinfoform.txtscompany.value = ""
	document.custinfoform.txtsaddress1.value = ""
	document.custinfoform.txtsaddress2.value = ""
	document.custinfoform.txtsaddress3.value = ""
	document.custinfoform.txtscity.value = ""
	document.custinfoform.txtszipcode.value = ""
	document.custinfoform.txtsstate.value=""
	document.custinfoform.txtsstate.selectedIndex = ""
	//document.custinfoform.shiptocountry.selectedIndex = 247
	document.custinfoform.shiptocountry.value = lstartcoo
	document.custinfoform.txtsphone.value = ""
	document.custinfoform.txtsemail.value = ""
	document.custinfoform.txtsemail1.value = ""

    if (lstartcoo=="073")
        {
            sbaddr3.style.display = '';
            stxtstatelabel.style.display='none' ;
            scountydrop.style.display='';
            scountylabel.style.display='' ;
        }
    else
    {
        sbaddr3.style.display = 'none';
        stxtstatelabel.style.display='' ;
        scountydrop.style.display='none';
        scountylabel.style.display='none' ;    
    }
    
    if ((lstartcoo=="001") || (lstartcoo=="034"))
    {
    var sbstate = document.getElementById('txtsstate')
            var sbstate1 = sbstate.length ;
             for(i=sbstate1-1;i>=0;i--)
               {        
                    sbstate.remove(i);       
                }
            
            Get_allstates(lstartcoo) ;        
            var answer = xmlhttp.responseText;
            allanswer = answer.split(";");
            listloaded = true 
            for (i=0 ; i < allanswer.length-1 ; i++)
            {
                btstr1 = allanswer[i] ;
                
                btstr2 = btstr1.split("-") ;
                btstcode = btstr2[0] ;
                btstcodeName = btstr2[1] ;
                addOption(sbstate,btstcodeName,btstcode) ;
            }           
            var printthis = ((lstartcoo=="034")?"Province" : "State")
            sstlabel.innerHTML = printthis ; 
             var printthis1 = ((lstartcoo=="034")?"Postal Code" : "Zip Code")
            sZiplabel.innerHTML = printthis1 ;
    
    }
	EnableAllFields(false)
  }
  
  var a = document.getElementById('GiftMsg');
	if (a.style.display =='') {
		a.style.display = 'none';
		//alert("none") ;
	}
	else {
		a.style.display='';
		//alert("blank") ;
	}
}

function EnableAllFields(bDisabled) {
	modifyTextField(document.custinfoform.txtsfirstname, bDisabled)
	modifyTextField(document.custinfoform.txtslastname, bDisabled)
	modifyTextField(document.custinfoform.txtscompany, bDisabled)
	modifyTextField(document.custinfoform.txtsaddress1, bDisabled)
	modifyTextField(document.custinfoform.txtsaddress2, bDisabled)
	modifyTextField(document.custinfoform.txtsaddress3, bDisabled)
	modifyTextField(document.custinfoform.txtscity, bDisabled)
	modifyTextField(document.custinfoform.txtsstate, bDisabled)
	modifyTextField(document.custinfoform.txtscounty, bDisabled)
	modifyTextField(document.custinfoform.txtszipcode, bDisabled)
	modifyTextField(document.custinfoform.shiptocountry, bDisabled)
	modifyTextField(document.custinfoform.txtsphone, bDisabled)
	modifyTextField(document.custinfoform.txtsemail, bDisabled)
	modifyTextField(document.custinfoform.txtsemail1, bDisabled)
	modifyTextField(document.custinfoform.txtgiftmsg1, bDisabled)
	modifyTextField(document.custinfoform.txtgiftmsg2, bDisabled)
	modifyTextField(document.custinfoform.txtgiftmsg3, bDisabled)
	modifyTextField(document.custinfoform.txtgiftmsg4, bDisabled)
	modifyTextField(document.custinfoform.txtgiftmsg5, bDisabled)
	modifyTextField(document.custinfoform.txtgiftmsg6, bDisabled)

	if (bDisabled)
	{
		document.custinfoform.txtgiftmsg1.style.background='#E7E7E7'
		document.custinfoform.txtgiftmsg2.style.background='#E7E7E7'
		document.custinfoform.txtgiftmsg3.style.background='#E7E7E7'
		document.custinfoform.txtgiftmsg4.style.background='#E7E7E7'
		document.custinfoform.txtgiftmsg5.style.background='#E7E7E7'
		document.custinfoform.txtgiftmsg6.style.background='#E7E7E7'
	}
	else
	{
		
		document.custinfoform.txtgiftmsg1.style.background='#FFFFFF'
		document.custinfoform.txtgiftmsg2.style.background='#FFFFFF'
		document.custinfoform.txtgiftmsg3.style.background='#FFFFFF'
		document.custinfoform.txtgiftmsg4.style.background='#FFFFFF'
		document.custinfoform.txtgiftmsg5.style.background='#FFFFFF'
		document.custinfoform.txtgiftmsg6.style.background='#FFFFFF'		
	}
	return true
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

function Contchar(exittxt,texto) {
  var exittxtObj=getObject(exittxt);
  exittxtObj.innerHTML = texto ;
}

function validateForm(billtophonereq,ShiptoEmailreq) {
	

  var haserror=false
 // var regex = /^(?:\([2-9]\d{2}\)\ ?|[2-9]\d{2}(?:\-?|\ ?))[2-9]\d{2}[- ]?\d{4}$/
 // var regex2 = /^(?:\([2-9]\d{2}\)\ ?|[2-9]\d{2}(?:\-?|\ ?))[2-9]\d{2}[- ]?\d{4}$/
 // var reqfield = "The following fields are required: "
  
    // bcountrysel = document.getElementById('billtocountry').selectedIndex ;
    // bcountryselval = document.getElementById('billtocountry').options[bcountrysel].value ;
     
     var checkforcountrydrop = document.getElementById('whatbcountrytype')
     
      if (checkforcountrydrop !=null)
        {
            var bcountrysel = document.getElementById('billtocountry').selectedIndex ;   		
            var bcountryselval = document.getElementById('billtocountry').options[bcountrysel].value ;
        }
        else
        {
            bcountryselval=document.getElementById('billtocountry').value ;
        }
  
	
     //scountrysel = document.getElementById('shiptocountry').selectedIndex ;
     scountrysel = document.getElementById('shiptocountry') ;
     if(scountrysel != null)
        {
            scountryselval = document.getElementById('shiptocountry').options[scountrysel.selectedIndex].value ;   
        }
     //alert(bcountryselval) ;

	Contchar('bfname','')
	Contchar('blname','')
	Contchar('baddr1','')
	Contchar('bcity','')
	//if (bcountryselval !="073")
	{
	Contchar('bstate','')
	}
	Contchar('bzip','')
	Contchar('bemail', '')
	Contchar('bemailrepeat', '')
	Contchar('bpwd','')
	
	var bchkboxbilltocopy = document.getElementById('shiptocountry').value ;
	if (bchkboxbilltocopy == 1)
	{
	//if (custinfoform.billtocopy.value == 0)
	{	
	Contchar('sfname','')
	Contchar('slname','')
	Contchar('saddr1','')
	Contchar('scity','')
	Contchar('sstate','')
	Contchar('szip','')
	Contchar('semail', '')
	Contchar('semailrepeat', '')
	}
	}
	
// Begin Required Field Check

    var bbfname = document.getElementById('txtfirstname').value ;
	var bblname = document.getElementById('txtlastname').value ;
	var bbaddr1 = document.getElementById('txtaddress1').value ;
	var bbaddr2 = document.getElementById('txtaddress2').value ;
	var bbcity = document.getElementById('txtcity').value ;
	var bbstate = document.getElementById('txtstate').value ;
	var bbzipcode = document.getElementById('txtzipcode').value ;
	var bbphone = document.getElementById('txtphone').value ;
	var bbemail = document.getElementById('txtemail').value;
	
	var bbregistered = document.getElementById('registered').value ;
	
	
	
	if (CheckForEmpty(bbfname))
	{
		Contchar('bfname','(Required)')
		haserror=true		
	}
	if (CheckForEmpty(bblname))
	{
		Contchar('blname','(Required)')
		haserror=true		
	}
	if ((CheckForEmpty(bbaddr1)==true) && (CheckForEmpty(bbaddr2)==true))
	{
		Contchar('baddr1','(Required)')
		haserror=true		
	}
	if (CheckForEmpty(bbcity))
	{
		Contchar('bcity','(Required)')
		haserror=true		
	}
	
	//if ((custinfoform.txtstate.value == "") && (bcountryselval !="073"))
	if (bcountryselval =="001")
	{
	    if (CheckForEmpty(bbstate))
	    {
		Contchar('bstate','(Required)')
		haserror=true		
		}
	}
	if (CheckForEmpty(bbzipcode))
	{
		Contchar('bzip','(Required)')
		haserror=true		
	}
	if ((CheckForEmpty(bbphone)==true) && billtophonereq==1)
	{
		Contchar('bphone','(Required)')
		haserror=true		
	}
	if (CheckForEmpty(bbemail))
	{
		Contchar('bemail','(Required)')		
		haserror=true
	}
	if (bbemail.length != 0 && !/^[a-zA-Z0-9._-]+@([a-zA-Z0-9.-]+\.)+[a-zA-Z0-9.-]{2,4}$/.test(bbemail))
	{
		Contchar('bemail','Invalid Email Address')		
		haserror=true
	}
	
		
	 if (bbregistered==0)
	 {

	     var bbemailrepeat = document.getElementById('txtemail1').value;
	     if (CheckForEmpty(bbemailrepeat)) {
	         Contchar('bemailrepeat', '(Required)')
	         haserror = true
	     }
	     
	     if ((CheckForEmpty(bbemail) == false) && (bbemail.toUpperCase() != bbemailrepeat.toUpperCase())) {
	         Contchar('bemailrepeat', '(E-Mail and Re-enter E-Mail do not match)')
	         haserror = true
	     }
	    
	     var bbpwd1 = document.getElementById('txtregpassword').value ;
		 var bbpwd2 = document.getElementById('txtregpassword2').value ;
		    
		if (bbpwd1.length > 0)
		{		
			if (bbpwd1.length < 4)
			{
				Contchar('bpwd','\<br>(password must be at least 4 characters)')
				haserror=true
			}
			if (bbpwd1 !== bbpwd2)
			{
				Contchar('bpwd','\<br>(Your passwords do not match; please retype your password in both fields)')
				haserror=true
			}	
		}
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
	
	
	
    var bbilltocopy = document.getElementById('billtocopy') ;
    
	if (bbilltocopy.checked == false)
	{
	 var sbfname = document.getElementById('txtsfirstname').value ;
	 var sblname = document.getElementById('txtslastname').value ;
	 var sbaddr1 = document.getElementById('txtsaddress1').value ;
	 var sbaddr2 = document.getElementById('txtsaddress2').value ;
	 var sbcity = document.getElementById('txtscity').value ;
	 var sbstate = document.getElementById('txtsstate').value ;
	 var sbzipcode = document.getElementById('txtszipcode').value ;
	 var sbphone = document.getElementById('txtsphone').value ;
	 var sbemail = document.getElementById('txtsemail').value;
	 var semailrep = document.getElementById('txtsemail1').value; 
	
	
		if (CheckForEmpty(sbfname))
		{
			Contchar('sfname','(Required)')
			haserror=true		
		}
		if (CheckForEmpty(sblname))
		{
			Contchar('slname','(Required)')
			haserror=true		
		}
		if ((CheckForEmpty(sbaddr1)==true) && (CheckForEmpty(sbaddr2)==true))
		{
			Contchar('saddr1','(Required)')
			haserror=true		
		}
		if (CheckForEmpty(sbcity))
		{
			Contchar('scity','(Required)')
			haserror=true		
		}
		
		//if ((custinfoform.txtsstate.value == "") && (scountryselval !="073"))
		if (scountryselval =="001")
		{
		    if (CheckForEmpty(sbstate))
		    {
			Contchar('sstate','(Required)')
			haserror=true		
			}
		}
		if (CheckForEmpty(sbzipcode))
		{
			Contchar('szip','(Required)')
			haserror=true		
		}
		
		if ((CheckForEmpty(sbemail)) && (ShiptoEmailreq==1))
		{
			Contchar('semail','(Required)')
			haserror=true		
		}
		
		
		/* if (ShiptoEmailreq==1)
		{ */
		    if (sbemail.length != 0 && !/^[a-zA-Z0-9._-]+@([a-zA-Z0-9.-]+\.)+[a-zA-Z0-9.-]{2,4}$/.test(sbemail))
	        {
		        Contchar('semail','Invalid Email Address')		
		        haserror=true
	        }
	    ///}

	        if ((CheckForEmpty(sbemail) == false) && (sbemail.toUpperCase() != semailrep.toUpperCase())) {
            Contchar('semailrepeat', '(E-Mail and Re-enter E-Mail do not match)')
            haserror = true
        } 

		
	}
	
	
	
	if (haserror)
		{
		window.location="custinfo.asp#billtab";
		return false ;
		}
	
     return true ;
	//return false ;
	 
}



function Fixbcountry(ccode,whichone)
{

//alert(whichone)
//if ((ccode=="001") || (ccode=="034"))
{
    if (whichone=="B")
    {
    var bbstate = document.getElementById('txtstate')
    }
    else
    {
    var bbstate = document.getElementById('txtsstate')
    }
    
    var btxtstatelabel = document.getElementById('txtstatelabel')
    var stxtstatelabel = document.getElementById('txtsstatelabel')
    
    var bstlabel = document.getElementById('bstatelabel')
    var sstlabel = document.getElementById('sstatelabel')
    
    var bcountylabel =document.getElementById('bukcounty')
    var bcountydrop = document.getElementById('txtcounty')
    
    var scountylabel =document.getElementById('sukcounty')
    var scountydrop = document.getElementById('txtscounty')
    
    var bZiplabel = document.getElementById('bzipcodelabel')
    var sZiplabel = document.getElementById('szipcodelabel')
    
    var bbstate1 = bbstate.length ;
    //var bbstateval = bbstate.value ;    
    // first remove all
    for(i=bbstate1-1;i>=0;i--)
    {
        //bbstateval = bbstate.options[i].value        
        bbstate.remove(i);
       
    }
    
    var bbaddr3 = document.getElementById('addr3')
    var sbaddr3 = document.getElementById('saddr3')    
    
    if (whichone=="B")
    {
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
    else
    {
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
    
   
}

}

function frmreset() {

    document.custinfoform.reset();
    
    EnableAllFields(false);
    
    var bbilltocopy = document.getElementById('billtocopy') ;
    
    if (bbilltocopy !=null)
    {
    if (bbilltocopy.checked==false)
    {
    var a = document.getElementById('GiftMsg');
    a.style.display = '' ;
    }
    }
      
    /*
    var a = document.getElementById('GiftMsg');
   
    if (document.custinfoform.billtocopy.checked && a.style.display == '') {
        a.style.display = 'none';
        modifyTextField(document.custinfoform.txtgiftmsg1, true)
        modifyTextField(document.custinfoform.txtgiftmsg2, true)
        modifyTextField(document.custinfoform.txtgiftmsg3, true)
        modifyTextField(document.custinfoform.txtgiftmsg4, true)
        modifyTextField(document.custinfoform.txtgiftmsg5, true)
        modifyTextField(document.custinfoform.txtgiftmsg6, true)

        document.custinfoform.txtgiftmsg1.style.background = '#E7E7E7'
        document.custinfoform.txtgiftmsg2.style.background = '#E7E7E7'
        document.custinfoform.txtgiftmsg3.style.background = '#E7E7E7'
        document.custinfoform.txtgiftmsg4.style.background = '#E7E7E7'
        document.custinfoform.txtgiftmsg5.style.background = '#E7E7E7'
        document.custinfoform.txtgiftmsg6.style.background = '#E7E7E7'
    }
    if (document.custinfoform.billtocopy.checked == false && a.style.display == 'none') {
        a.style.display = '';
        modifyTextField(document.custinfoform.txtgiftmsg1, false)
        modifyTextField(document.custinfoform.txtgiftmsg2, false)
        modifyTextField(document.custinfoform.txtgiftmsg3, false)
        modifyTextField(document.custinfoform.txtgiftmsg4, false)
        modifyTextField(document.custinfoform.txtgiftmsg5, false)
        modifyTextField(document.custinfoform.txtgiftmsg6, false)

        document.custinfoform.txtgiftmsg1.style.background = '#FFFFFF'
        document.custinfoform.txtgiftmsg2.style.background = '#FFFFFF'
        document.custinfoform.txtgiftmsg3.style.background = '#FFFFFF'
        document.custinfoform.txtgiftmsg4.style.background = '#FFFFFF'
        document.custinfoform.txtgiftmsg5.style.background = '#FFFFFF'
        document.custinfoform.txtgiftmsg6.style.background = '#FFFFFF'
    }
    */
    

}
