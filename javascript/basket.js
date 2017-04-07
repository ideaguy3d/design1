function ChangeCarrier(page)
{
xmlhttp=null
var new_url=page;
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
function RedirectToAddADDress (page)
   {
   var new_url=page;	  
   //if (  (new_url != "")  &&  (new_url != null)  )
   //document.basketform.submit();
    window.location=new_url;
   }

function MM_openBrWindow() 
{ //v2.0
  window.open('estimate_shipping.asp','estship','scrollbars=yes,width=550,height=400');
}