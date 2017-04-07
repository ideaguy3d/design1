function Changeshipdate(item_id,stdate)
{
xmlhttp=null

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
  new_url ="ChangeShipwhen.asp?itemid="+item_id+"&shipwhendate="+ stdate
 
  xmlhttp.onreadystatechange=state_Change
  xmlhttp.open("GET",new_url,false)
  xmlhttp.send(null)
  }
else
  {
  alert("Your browser does not support XMLHTTP.")
  }
}function state_Change()
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