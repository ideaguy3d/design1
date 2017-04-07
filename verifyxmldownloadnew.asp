<% server.ScriptTimeout=300 %>

<%
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
 %>

<!--#INCLUDE FILE = "include/commonvar.asp" -->



<%


    set sitelink=createobject("sitel700.sitelink")	
	call sitelink.setdata(datapath,cstr(""),ShortStoreName)
	
    xmlresult= sitelink.verifyxmldownloadordersnew(cstr(Request.QUERYSTRING("userid")),cstr(Request.QUERYSTRING("userpassword")),session,ShortStoreName)
    Response.Write(xmlresult)
    
    set sitelink=nothing
    session.Abandon()
    
 
%>
