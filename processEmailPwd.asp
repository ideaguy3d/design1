
<%on error resume next%>
<%
	 dim emailtouse,pwdreturn     
	 emailtouse = trim(REQUEST.FORM("emailtouse"))	 
	 txtzipcode = cstr(request.form("txtzipcode"))
	 
	  if len(emailtouse) = 0 or len(txtzipcode) = 0 then	 	
		 response.redirect("emailpassword.asp?error=1") 
	 end if 

%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->
<!--#INCLUDE FILE = "include/cLibraryExtra.asp" -->

<%
   	
	 
	 emailtouse=fix_xss_Chars(emailtouse)
	 txtzipcode=fix_xss_Chars(txtzipcode)
	 
	 ' first check whether a  password is available
	 pwdreturn = ""
	 pwdreturn =  sitelink.SENDPASSWORD_DPS(cstr(emailtouse),cstr(txtzipcode))
	 

	 	 
	 if len(trim(pwdreturn)) > 0 then
	 
	     set objOrdConfrmEmail = New cAdminEmail
           objOrdConfrmEmail.action="READ"
            objOrdConfrmEmail.store = ShortStoreName
            xmlstring = sitelink.Maintain_OrdConfrmEmailText(objOrdConfrmEmail)
        set objOrdConfrmEmail = nothing
    
        objDoc.loadxml(xmlstring)
        set SL_OrdConfrmEmail = objDoc.selectNodes("//results")
        
        SL_FromEmail=""
        SL_StoreName= ""
        if SL_OrdConfrmEmail.length > 0 then
            SL_StoreName = SL_OrdConfrmEmail.item(0).selectSingleNode("storename").text
            SL_FromEmail = SL_OrdConfrmEmail.item(0).selectSingleNode("fromemail").text
        end if 
    
        set SL_OrdConfrmEmail = nothing
     
     
     
		dim DataToSend, xmlHTTP
		    DataToSend ="pwd="& server.URLEncode(pwdreturn)
			DataToSend = DataToSend & "&email="  & emailtouse
			DataToSend = DataToSend & "&fromemail="& SL_FromEmail
			DataToSend = DataToSend & "&storename="& SL_StoreName
			
			set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
			xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink700/emailpasswordSl7.asp",false
            xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            'xmlhttp.setRequestHeader "Content-Type", "charset=iso-8859-1"
			xmlhttp.send DataToSend
			'Response.Write xmlhttp.ResponseText
			Set xmlhttp = nothing 
			session("Title_String") = "Thank you for your Request"
			session("strmessage") = "<ul>Your password has been emailed to you. </ul>"
			
			
			set ObjDoc = nothing
	        set sitelink = nothing
	    
			session("destpage")=""
			response.redirect("thankyou.asp")
	 else
	    set ObjDoc = nothing
	    set sitelink = nothing
	    response.redirect("emailpassword.asp?error=1") 		 
	 end if 


%>
