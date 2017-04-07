
<%on error resume next%>

<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%
    'token id and payerid are in the querystring.. It is the same as what was passed in SetExpressCheckout
    'token=EC-7MY86553DU322535A&PayerID=PCEUUXECN34H2
    
    TokenID = cstr(request.QueryString("token"))
    PayerID = request.QueryString("PayerID")
    
    if len(trim(TokenID)) = 0 then
        set sitelink = nothing
        ErrorMsg = "Invalid Token"
		session("custerrormessage") = 	ErrorMsg +"<br>"
		response.Redirect("PayPalError.asp")		
    end if
    
    PayPalUrl = ""
    DataToSend=""
    
    set SLCustomer = New cCustomer
    set PayPalExpressObj = New cPayPalExpressObj
        PayPalExpressObj.Action= "GETEXPRESSCHECKOUTDETAILS"
        PayPalExpressObj.OrderNum = 0
        PayPalExpressObj.TokenID = TokenID        
        
        call sitelink.Getpostdata_PayPalExpress(PayPalExpressObj,SLCustomer)
        DataToSend =PayPalExpressObj.PayPalPostData
        PayPalUrl = PayPalExpressObj.PayPalPostUrl
        
    set PayPalExpressObj = nothing
    set SLCustomer = nothing

    
   set xmlhttp = server.Createobject("Msxml2.ServerXMLHTTP.6.0")
	        xmlhttp.Open "POST",PayPalUrl,false
            xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			xmlhttp.send DataToSend
			ReturnText = xmlhttp.responseText
			set xmlHTTP=nothing
			
			'response.Write(DataToSend)
			'response.Write("<br>")
			'response.Write("<br>")
			'response.Write(ReturnText)
			
			'MyArray = Split(ReturnText, "&", -1, 1)
			
			'For Each name In MyArray			    
			'    response.Write(name)
			'    response.Write("<br>")			
			'next
		set xmlhttp =nothing	
			
		set SLCustomer = New cCustomer
        set PayPalExpressObj = New cPayPalExpressObj		
		
        PayPalExpressObj.Action= "VALIDATE_GETEXPRESSCHECKOUTDETAILS"
        PayPalExpressObj.OrderNum = session("ordernumber")
        PayPalExpressObj.PayPalRetMsg = ReturnText
        PayPalExpressObj.PayerID=PayerID
        SLCustomer.shopperid = session("shopperid")
        PayPalExpressObj.PayPalPostData =DataToSend
        
        call sitelink.Getpostdata_PayPalExpress(PayPalExpressObj,SLCustomer)
        ErrorMsg = PayPalExpressObj.ErrorMsg
        HasError = PayPalExpressObj.HasError       
        set PayPalExpressObj = nothing    
        
        if HasError=true then		
		    set sitelink = nothing
			session("custerrormessage") = 	ErrorMsg +"<br>"
			response.Redirect("PayPalError.asp")			 
		end if	
		
		
'		response.Write("<Br>")
'		response.Write("HasError -->" & HasError)
'		response.Write("<Br>")
'		response.Write("<Br>")
'		response.Write("ErrorMsg -->" & ErrorMsg)			
'		response.Write("<br>")
'		
		SLCustomer.shopperid   = session("shopperid")
		session("ordernumber") = SLCustomer.bordernumber
		session("billtocopy")  = SLCustomer.docopy		
       session("firstname") = SLCustomer.bfirstname
       session("lastname")  = SLCustomer.blastname   
       session("company") =   SLCustomer.bcompany 
       session("address1") =SLCustomer.baddr1  
       session("address2")=SLCustomer.baddr2   
       session("address3")=SLCustomer.baddr3   
       session("city") = SLCustomer.bcity      
       session("state") = SLCustomer.bstate    
        session("zipcode") = SLCustomer.bzipcode
        session("phone") = SLCustomer.bphone    
        session("email") = SLCustomer.bemail    
        session("country") = SLCustomer.bcountry 
        session("receiveemail") = SLCustomer.Receive_Email
        session("sfirstname") = SLCustomer.sfirstname           
        session("slastname") = SLCustomer.slastname           
        session("scompany") = SLCustomer.scompany     
        session("saddress1") = SLCustomer.saddr1      
        session("saddress2") = SLCustomer.saddr2      
        session("saddress3") = SLCustomer.saddr3      
        session("scity") = SLCustomer.scity        
        session("sstate") = SLCustomer.sstate      
        session("szipcode") = SLCustomer.szipcode  
        session("sphone") = SLCustomer.sphone      
        session("scountry") = SLCustomer.scountry  
        session("semail") = SLCustomer.semail      
        session("bcounty") = SLCustomer.bcounty    
        session("scounty") = SLCustomer.scounty    
        session("billtocopy") = SLCustomer.docopy        
'        
'		
'		response.Write("<br>CustDetails<Br>")
'		response.Write("======Billing====<br>")	
'		response.Write("Order Number->"& SLCustomer.bordernumber)
'		response.Write("<br>")
'		response.Write("Shopper Number->"& SLCustomer.shopperid)
'		response.Write("<br>")			
'		response.Write("First Name->"& session("firstname"))
'		response.Write("<br>")			
'        response.Write("last Name->"& session("lastname"))
'        response.Write("<br>")			
'        response.Write("addr1->"& session("address1"))
'        response.Write("<br>")			
'        response.Write("addr2 Name->"& session("address2"))
'        response.Write("<br>")			
'        response.Write("city->"& session("city"))
'        response.Write("<br>")			
'        response.Write("state->"& session("state"))
'        response.Write("<br>")			
'        response.Write("zip->"& session("zipcode"))
'         response.Write("<br>")			
'        response.Write("country->"& session("country"))
'         response.Write("<br>")			
'        response.Write("email->"& session("email"))
'        response.Write("<br>")	
'        response.Write("======Shipping====<br>")        
'        response.Write("Do Copy -->"& SLCustomer.docopy)
'        response.Write("<br>")	
'        response.Write("first name->"& session("sfirstname"))
'        response.Write("<br>")			
'        response.Write("last name->"& session("slastname"))
'        response.Write("<br>")			
'        response.Write("addr1->"& session("saddress1"))
'        response.Write("<br>")	
'        response.Write("addr2->"& session("saddress2"))
'        response.Write("<br>")			
'        response.Write("city->"& session("scity"))
'        response.Write("<br>")			
'        response.Write("state->"& session("sstate"))
'        response.Write("<br>")			
'        response.Write("zip->"& session("szipcode"))
'        response.Write("<br>")			
'        response.Write("country->"& session("scountry"))
'        response.Write("<br>")	
        
        
        
        session("custerrormessage") =sitelink.SAVECUSTOMERINFO(SLCustomer)
        aa=session("custerrormessage")
        if len(trim(aa))>0 then
		    Set objDoc = nothing
		    set sitelink=nothing 		
		    response.Redirect("PayPalError.asp")		
	    end if
        
        
        if session("billtocopy")=1 then			 
			 vtextshipcountry = session("country")		
			 vtxtzipCode = session("zipcode")	 
		else		
			 vtextshipcountry = session("scountry")	
			 vtxtzipCode = session("szipcode")	
		end if	
		
		
		'Now check for Shipping restriction.
		session("ProdRestList")=""
		hasProdRest = false
		session("CheckedForProdRest") = false 
		xmlstring = sitelink.ProductRestriction(session("ordernumber"),session("shopperid"),session("user_want_multiship"))
		objDoc.loadxml(xmlstring)		
			    
		set ProdRestrict = objDoc.selectNodes("//results[lrest_sale=1 or lrest_sale=0]") 
		if ProdRestrict.length > 0 then		    
		     session("ProdRestList") = xmlstring  
		    hasProdRest = true
		end if
		
		set ProdRestrict=nothing
		
		if hasProdRest=true then
		     set Sitelink=nothing
		     set ObjDoc = nothing
		     session("custerrormessage") =" These items are restricted from being shipped to its location"
		     session("custerrormessage") = "<ul type=ROUND><LI>"&session("custerrormessage")&"</ul>" 
		     response.redirect("custerror.asp")	
		end if
		        
        session("ProdRestHoldList")=""
        session("hasProdRest") = false
		set ProdRestrict = objDoc.selectNodes("//results[lrest_sale=2]") 
		if ProdRestrict.length > 0 then		    
		     session("ProdRestHoldList") = xmlstring  
		     session("hasProdRest") = true
		end if
		
		
		set ProdRestrict=nothing
				        
		session("CheckedForProdRest") = true
		
		

		if session("shippingmethod") = "" then
			'get a default shipping method first			    
			    xmlShiplist = sitelink.GET_SHIPLIST_ZONES(cstr(vtextshipcountry),cstr(vtxtzipCode))
			    objDoc.loadxml(xmlShiplist)
			    
			    ShipMethod=""
			    if DEFAULT_SHIPPING_METHOD<>"" then
			        target_node = "//shiplist[ca_code='"+cstr(DEFAULT_SHIPPING_METHOD)+"']"
			        set SL_Ship = objDoc.selectNodes(target_node) 
			            if SL_Ship.length > 0 then
			                ShipMethod = DEFAULT_SHIPPING_METHOD
			            end if
			        set SL_Ship = nothing
			    end if
			    
			    if ShipMethod="" then
			        target_node = "//shiplist"
			        set SL_Ship = objDoc.selectNodes(target_node) 
			            if SL_Ship.length > 0 then
			                ShipMethod = SL_Ship.item(0).selectSingleNode("ca_code").text
			            end if
			        set SL_Ship = nothing 			    
			    end if
			    
			    session("shippingmethod")= ShipMethod
		end if
		
		
		
		'shipping method must exists
		xmlCarrier = sitelink.SHIPPINGMETHODS() 
		objDoc.loadxml(xmlCarrier)
		
		ShipMethodValid = false
		target_node = "//shipmeths[ca_code='"+cstr(session("shippingmethod"))+"']"
		set SL_Ship = objDoc.selectNodes(target_node) 
		    if SL_Ship.length > 0 then
		        ShipMethodValid=true
		    end if
		set SL_Ship = nothing
		
		if ShipMethodValid = false then
		    set SL_Ship = objDoc.selectNodes("//shipmeths") 
		    if SL_Ship.length > 0 then
		        session("shippingmethod") = SL_Ship.item(0).selectSingleNode("ca_code").text	
		    end if 
		end if 
		
		
		session("TokenID") = cstr(TokenID) 
		
		'securepage = secureurl + "PayPalReviewOrder.asp?tokenid=" + cstr(TokenID) + "&PayerID=" + cstr(PayerID)
		securepage = secureurl + "PayPalReviewOrder.asp"
        
        set sitelink= nothing
    
        response.Redirect(securepage)
			
 %>
