<%on error resume next%>


<%
	sitestore	=request.form("sitestore")
	if sitestore<>"SITELINK" then	
		Response.Status="301 Moved Permanantly"
		Response.AddHeader "Location", insecureurl&"default.asp"
		Response.End
	end if
%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<%

    if REQUEST.servervariables("server_port_secure")<> 1 and SLsitelive=true then 
        set sitelink=nothing
        set ObjDoc=nothing
    	response.redirect(secureurl+"firsttimeregister.asp") 
    end if
    
   ' Information about the shopper
	session("firstname")=cstr(REQUEST.FORM("txtfirstname"))
	session("lastname")=cstr(REQUEST.FORM("txtlastname"))
	session("company")=cstr(REQUEST.FORM("txtcompany"))
	session("address1")=cstr(REQUEST.FORM("txtaddress1"))
	session("address2")=cstr(REQUEST.FORM("txtaddress2"))
	session("address3")=cstr(REQUEST.FORM("txtaddress3"))	
	session("city")=cstr(REQUEST.FORM("txtcity"))
	session("state")=UCASE(cstr(REQUEST.FORM("txtstate")))
	session("zipcode")=cstr(REQUEST.FORM("txtzipcode"))
	session("country")=cstr(REQUEST.FORM("billtocountry"))
	session("phone")=cstr(REQUEST.FORM("txtphone"))
	session("email")=cstr(REQUEST.FORM("txtemail"))
	session("billtocopy")=1
	session("bcounty")=cstr(REQUEST.FORM("txtcounty"))
	
	
	session("reglastname")=session("lastname")
	session("regzipcode")=session("zipcode")
	session("regpassword")=cstr(REQUEST.FORM("txtregpassword"))
	session("regpassword2")=cstr(REQUEST.FORM("txtregpassword2"))

	
	
		session("firstname")=fix_xss_Chars(session("firstname"))
		session("lastname")=fix_xss_Chars(session("lastname"))
		session("company")=fix_xss_Chars(session("company"))
		session("address1")=fix_xss_Chars(session("address1"))
		session("address2")=fix_xss_Chars(session("address2"))
		session("address3")=fix_xss_Chars(session("address3"))
		session("city")=fix_xss_Chars(session("city"))
		session("state")=fix_xss_Chars(session("state"))
		session("zipcode")=fix_xss_Chars(session("zipcode"))
		session("country")=fix_xss_Chars(session("country"))
		session("phone")=fix_xss_Chars(session("phone"))
		session("email")=fix_xss_Chars(session("email"))	
		session("regpassword")=fix_xss_Chars(session("regpassword"))
		session("regpassword2")=fix_xss_Chars(session("regpassword2"))
	
		invalid_entry = false
		
		'now check billto country.
'	 	retval = sitelink.COUNTRYNAME(session("country"))
'	 	if retval="" then
'	 	    invalid_entry = true
'	 	end if
'		
		if session("lastname")="0" or session("state")="0" then
			invalid_entry = true	 
		end if

		
'		xmlstring =sitelink.Get_state_list(session("ordernumber"))
'	 	objDoc.loadxml(xmlstring)
	 	
	 	'check billto state..
'	 	if len(session("state")) > 0 then
'	 	    if session("state")<>"INT" then	 	        
'	 	        tnode = "//state_list[sstate='"+session("state")+"']"
'	 	        set SL_Ship_state = objDoc.selectNodes(tnode)
'	 	        if SL_Ship_state.length = 0 then
'	 	            invalid_entry = true
'	 	        end if
'	 	        
'	 	    end if 	 	
'	 	end if
'	 	set SL_Ship_state=nothing
	 
		if invalid_entry = true then
			set sitelink=nothing
			set ObjDoc=nothing	
			Response.Status="301 Moved Permanantly"
			Response.AddHeader "Location", insecureurl&"default.asp"
			Response.End
		end if
		
	if BILL_TOPHONE_REQUIRED=1 and len(trim(session("phone")))=0 then
		session("prospecterror") = "<ul><li>Bill to Phone Required"
		session("prospecterror") = session("prospecterror") + "</li></ul><br><br>"  
		set sitelink=nothing 
	    Set objDoc = nothing
		response.redirect("proserror.asp")
	end if
	
	'validate phone
	   bphone = session("phone")
     if len(trim(bphone)) > 0 then
	     bphone = replace(bphone," ","")
	     bphone = replace(bphone,"-","")
	     bphone = replace(bphone,"(","")
	     bphone = replace(bphone,")","")
    	 
	     blenphone = len(bphone)
	     invalidphone = false
    	 
	     For intCounter = 1 To blenphone
           strChar = Mid(bphone, intCounter, 1)
           if isnumeric(strChar)=false then
             invalidphone=true
           end if 
         next
         
         if invalidphone=true  then
		    session("prospecterror") = "<ul><li>Invalid Phone"
		    session("prospecterror") = session("prospecterror") + "</li></ul><br><br>"  
		    set sitelink=nothing 
	        Set objDoc = nothing
		    response.redirect("proserror.asp")
	    end if
	end if
	
	aa = sitelink.Check_existing_email("EMAIL",session("lastname"),session("zipcode"),session("regpassword"),session("regpassword2"),session("email"))	
			if len(trim(aa))>0 then
			   session("prospecterror") = "<ul><li>"+ aa 
			   session("prospecterror") = session("prospecterror") + "<br><br>" 
			   session("prospecterror") = session("prospecterror") + "<a href=""emailpassword.asp"">Click here</a> to retrieve  your password if you already registered."
			   session("prospecterror") = session("prospecterror") + "</li></ul><br><br>"  
			   set sitelink=nothing 
			   Set objDoc = nothing
				response.redirect("proserror.asp")
			end if
			
	aa= sitelink.Check_existing_password("EMAIL",session("lastname"),session("zipcode"),session("regpassword"),session("regpassword2"),session("email"))	
	if len(trim(aa))>0 then
	   session("prospecterror") = "<ul><li>"+ aa +"</li></ul>"
	   set sitelink=nothing 
	   Set objDoc = nothing
		response.redirect("proserror.asp")
	end if
	
	
	set SLCustomer = New cCustomer	
		SLCustomer.shopperid    = session("shopperid")
		SLCustomer.thisreferral = session("referral")
		SLCustomer.bordernumber	= 0
		SLCustomer.docopy       = session("billtocopy")
		SLCustomer.bfirstname   =session("firstname")
        SLCustomer.blastname    =session("lastname")
        'SLCustomer.bsalu        =""
        SLCustomer.bcompany     =session("company")
        'SLCustomer.bTitle       =""
        SLCustomer.baddr1       =session("address1")
        SLCustomer.baddr2       =session("address2")
        SLCustomer.baddr3       =session("address3")
        SLCustomer.bcity        =session("city")
        SLCustomer.bstate       =session("state")
        SLCustomer.bzipcode     =session("zipcode")
        SLCustomer.bphone       =session("phone")
        'SLCustomer.bphone2      =""
        SLCustomer.bemail       =session("email")
        SLCustomer.bcountry     =session("country")
        SLCustomer.bcounty      =session("bcounty")
        SLCustomer.Receive_Email=0
		'SLCustomer.bCatalogcode =""
        'SLCustomer.sfirstname   =""
        'SLCustomer.slastname    =""
        'SLCustomer.scompany     =""
        'SLCustomer.saddr1       =""
        'SLCustomer.saddr2       =""
        'SLCustomer.saddr3       =""
        'SLCustomer.scity        =""
        'SLCustomer.sstate       =""
        'SLCustomer.szipcode     =""
        'SLCustomer.sphone       =""
        'SLCustomer.sphone2      =""
        'SLCustomer.scountry     =""
        'SLCustomer.semail       =""

 	session("prospecterror") =sitelink.SAVECUSTOMERINFO(SLCustomer)
	
	set SLCustomer = nothing
	
	aa=session("prospecterror")
	if len(trim(aa))>0 then	    
	    set sitelink=nothing 
		Set objDoc = nothing		
		response.redirect("proserror.asp")
	end if
	

	if CREATE_PROSPECT =1 then
		'save as prospect so that it atleast goes to to MOM
		'session("prospecterror")=sitelink.saveprospectinfo(session("firstname"),session("lastname"),session("company"),session("address1"),session("address2"),session("city"),session("state"),session("zipcode"),session("phone"),session("email"),session("catalog"),session("shopperid"),session("referral"),session("country"),session("address3"))
		'session("prospecterror") =sitelink.saveprospectinfo(SLCustomer)
		call sitelink.MARK_AS_PROSPECT(session("shopperid"))
	end if
	
	
	
	
	' Now Regsiter the shopper
	
	session("registererror")=sitelink.registershopper(session("shopperid"),session("reglastname"),session("regzipcode"),SESSION("regpassword"),SESSION("regpassword2"),session("ordernumber"))
	
	
if len(trim(session("registererror")))=0 then
	session("registeredshopper")="YES"
	session("failedtoregistershopper")="NO"

	session("PASSWORD")= session("regpassword")

	session("addressbookcustnum") = sitelink.makeaddressbook(session("LASTNAME"),session("ZIPCODE"),SESSION("PASSWORD"),session("shopperid"))
	

	session("password")=session("regpassword")
	session("Title_String") = "Thank you for Registering"
	set sitelink=nothing 
	'if session("Clicked_Multishipto")=true then
	'	session("user_want_multiship")=true
	'end if 
	response.redirect("thankyou.asp")

else
	session("registeredshopper")="NO"
	session("failedtoregistershopper")="YES"
	set sitelink=nothing 
	Set objDoc = nothing
	session("prospecterror")="<ul type=ROUND><LI>"&session("registererror")&"</ul>"
	response.redirect("proserror.asp")	

end if

 %>

