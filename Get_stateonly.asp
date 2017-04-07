<%on error resume next%>
<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 

<%

countrycode = request.QueryString("code")
'countrycode="073"
codeisgood = false
if countrycode="001" or countrycode="034" or countrycode="073" then
    codeisgood=true
end if

if codeisgood=true then

 %>
 
<!--#INCLUDE FILE = "include/momappNoClibrary.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->



<%

if countrycode="001" or countrycode="034" then

    xmlstring =sitelink.Get_state_list(session("ordernumber"))
    objDoc.loadxml(xmlstring)
    tnode = "//state_list[ccode='"+countrycode+"']"
    'response.Write(tnode)
    set SL_Ship_state = objDoc.selectNodes(tnode)
        for x=0 to SL_Ship_state.length-1
		    SL_state_name= SL_Ship_state.item(x).selectSingleNode("statename").text
		    SL_state_code= SL_Ship_state.item(x).selectSingleNode("sstate").text
	        response.Write(SL_state_code +  "-" + SL_state_name + ";")
       next
    set SL_Ship_state = nothing  
 else
  '' get UK county
  xmlstring =sitelink.Get_County_list(session("ordernumber"))
  objDoc.loadxml(xmlstring)
  set SL_County_List = objDoc.selectNodes("//county_list")
        for x=0 to SL_County_List.length-1
               SL_county_name= SL_County_List.item(x).selectSingleNode("countyname").text
               SL_county_Code= SL_County_List.item(x).selectSingleNode("ccode").text
               response.Write(SL_county_Code +  "-" + SL_county_name + ";")
        next
  set SL_County_List = nothing
 end if 
 
 set ObjDoc = nothing
 set sitelink=nothing
 end if  
 
 
%>
		
