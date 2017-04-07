<%on error resume next%>


<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 

<!--#INCLUDE FILE = "include/momapp.asp" -->

<%

    'usethisskey = request.QueryString("key")
    usethisskey = fix_xss_Chars(request.QueryString("key"))
    
    'desturl     = request.QueryString("dest")
    desturl     = fix_xss_Chars_emailpromo(request.QueryString("dest"))
    
    if len(trim(usethisskey)) > 0 then
        call sitelink.WRITE_SLSKEY(cstr(usethisskey), session("shopperid"))
    end if
    
    
    DestUrlIsGood = sitelink.Check_DestUrl_EmailPromotion(insecureurl,desturl)
    
    set sitelink=nothing
    
    
    if (DestUrlIsGood=false) then
        response.Redirect(insecureurl)
    end if 
    
    if len(trim(desturl)) > 0 then
        response.Redirect(desturl)
    else
        response.Redirect("default.asp")    
    end if

	function fix_xss_Chars_emailpromo(thisvar)
		'for scan alert
		thisvar = replace(thisvar,"<","")
		thisvar = replace(thisvar,">","")
		thisvar = replace(thisvar,"'","")
		thisvar = replace(thisvar,"(","")
		thisvar = replace(thisvar,")","")
		'thisvar = replace(thisvar,"/","")
		thisvar = replace(thisvar,"\","")
		thisvar = replace(thisvar,"""","")		
		fix_xss_Chars_emailpromo = thisvar
	end function
 %>
