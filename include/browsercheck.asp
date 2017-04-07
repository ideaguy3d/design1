<!-- #include file="../text/adminMobileSiteRedirect.asp" -->
<% 
if session("viewfullsite") <> 1 and MobileStoreRedirect=1 then 
    
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Code for mobile site redirect. Mobile site supports & tested for iPhone, iPod, Android User Agents ONLY. Redirects Windows 7 phone to mobile template but the site is not tested for this UA.
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    IsMobileUA = false 
    user_agent = Request.ServerVariables("HTTP_USER_AGENT")

    Set Regex = New RegExp
    With Regex
        '.Pattern = "(android|iphone|windows (ce|phone))"
       .Pattern = "(android.+mobile|android.+nexus 7|ip(hone|od)|kindle|windows (ce|phone))"
       .IgnoreCase = True
       .Global = True
    End With
    
    match = Regex.Test(user_agent)
    if match then IsMobileUA = true 
    
    'check for kindle fire
    if IsMobileUA = false then 
        With Regex
           .Pattern = "(Silk|Silk-Accelerated)"
           .IgnoreCase = True
           .Global = True
        End With
        match = Regex.Test(user_agent)
        if match then IsMobileUA = true 
    end if
    
    if IsMobileUA = true then 
        
        'patch for iPad OS 4.3.1
        IsTablet=false
        With Regex
           .Pattern = "(ipad)"
           .IgnoreCase = True
           .Global = True
        End With
        
        match = Regex.Test(user_agent)
        if match then IsTablet = true 
        if IsTablet=false then 
	        rewriteurl = request.ServerVariables("http_x_rewrite_url")
	        if len(rewriteurl) > 0 then 
	            page_url = request.ServerVariables("http_host") & request.ServerVariables("http_x_rewrite_url") 'mid(insecureurl,1,len(insecureurl)-1)
	        else 
	            page = Request.ServerVariables("URL")
	            querystring = Request.Querystring
	            page_url = Request.ServerVariables("SERVER_NAME")
	            if ( (lcase(page) ="default.asp") and (len(querystring) > 0) ) or (len(querystring) > 0) or (lcase(page) <> "default.asp") then 
	                page_url = page_url & page 
	                if len(querystring) > 0 then page_url = page_url & "?" & querystring
	            end if 
	        end if 
        	
	        tmpmobileinsurl = replace(mobileinsecureurl,"http://","")
	        tmpinsurl = replace(insecureurl,"http://","")
	        page_url = replace(page_url,tmpinsurl,tmpmobileinsurl,1,1,1)
	        if REQUEST.servervariables("server_port_secure")=1 then 
		        page_url = "https://"& page_url
	        else 
		        page_url = "http://"& page_url
	        end if 

            set objDoc = nothing 
            set sitelink = nothing
            response.Redirect (page_url)
        end if 
    end if 
    Set Regex = nothing
    
end if
%>
