<%on error resume next%>
<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 


<!--#INCLUDE FILE = "include/momapp.asp" -->




<%

  itemid  = Request.QueryString("item_id")
  shipmethod = Request.QueryString("shipmethod")
  itemid = itemid + 1 - 1
  
  'itemid  = REQUEST.FORM("item_id")
  'itemid = itemid + 1 - 1
  'shipmethod = REQUEST.FORM("shipmethod")
  
  if len(trim(shipmethod)) = 0 then
  	shipmethod = ""
  end if 
  
  'response.write(shipmethod)
  
  if isnumeric(itemid)=true then
	  call sitelink.WRITESHIP_VIA_BASKET(session("ordernumber"),itemid , cstr(shipmethod))
  end if
    
  set sitelink=nothing 
	

'response.write(shipmethod)
' Response.Redirect("basket.asp")

%>