<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%
    TokenID = cstr(request.form("TokenID"))
    PayerID =""
    
    if len(trim(TokenID)) = 0 then
        set sitelink = nothing
        ErrorMsg = "Invalid Token"
		session("custerrormessage") = 	ErrorMsg +"<br>"
		response.Redirect("PayPalError.asp")		
    end if
    
    set SLCustomer = New cCustomer
    set PayPalExpressObj = New cPayPalExpressObj
        PayPalExpressObj.Action= "VALIDATE_PAYPALREVIEWORDER"
        PayPalExpressObj.OrderNum = 0
        PayPalExpressObj.TokenID = TokenID
     call sitelink.Getpostdata_PayPalExpress(PayPalExpressObj,SLCustomer)
        OrderNumber =PayPalExpressObj.OrderNum
        ErrorMsg = PayPalExpressObj.ErrorMsg
        HasError = PayPalExpressObj.HasError
        PayerID = PayPalExpressObj.PayerID
        DataToSend =PayPalExpressObj.PayPalPostData
        PayPalUrl = PayPalExpressObj.PayPalPostUrl
        
        'if PayPalExpressObj.OrderNum > 0 then        
        '    session("ordernumber") = PayPalExpressObj.OrderNum
        'end if
    set PayPalExpressObj = nothing
    set SLCustomer = nothing
    
    
    'response.Write(HasError)
    'response.Write("<br>")
    'response.Write(session("ordernumber"))
    'response.Write("<br>")
    'response.Write(TokenID)
    
    
    if HasError=true then		
	    set sitelink = nothing
		session("custerrormessage") = 	ErrorMsg +"<br>"
		response.Redirect("PayPalError.asp")			 
	end if	
	
    
    if OrderNumber <> session("ordernumber") then
        'regenarate session
        session("ordernumber") = OrderNumber
        
        billtocust = sitelink.GetBilltoCust(session("ordernumber"))
	        set shopperRecord= sitelink.SHOPPERINFO(cstr(billtocust),false)
	        session("shopperid") = shopperRecord.custnum
	        session("firstname")= trim(shopperRecord.FirstName)
	        session("lastname")= trim(shopperRecord.Lastname)
	        session("address1")= trim(shopperRecord.addr)
	        session("address2")= trim(shopperRecord.addr2)
	        session("address3")= trim(shopperRecord.addr3)
	        session("company")= trim(shopperRecord.company)
	        session("city")= trim(shopperRecord.city)
	        session("state")= trim(shopperRecord.state)
	        session("zipcode")= trim(shopperRecord.zipcode)
	        session("email")= trim(shopperRecord.email)
	        session("country")= trim(shopperRecord.country)
	        'session("bcounty")= trim(shopperRecord.county)
	        set shopperRecord= nothing
	        
	        set shopperRecord= sitelink.SHIPTONAME(cstr(session("ordernumber")),false)
				session("sfirstname")= trim(shopperRecord.FirstName)
				session("slastname")= trim(shopperRecord.Lastname)
				session("saddress1")= trim(shopperRecord.addr)
				session("saddress2")= trim(shopperRecord.addr2)
				session("saddress3")= trim(shopperRecord.addr3)
				session("scompany")= trim(shopperRecord.company)
				session("scity")= trim(shopperRecord.city)
				session("sstate")= trim(shopperRecord.state)
				session("szipcode")= trim(shopperRecord.zipcode)
				session("semail")= trim(shopperRecord.email)
				session("scountry")= trim(shopperRecord.country)	
				session("scounty")= trim(shopperRecord.county)							 
			 	set shopperRecord= nothing
			 	
	        session("billtocopy") = 0
	        set ORDER_RECORD = sitelink.ORDER_RECORD(session("ordernumber"),false)	 			 
			 session("shippingmethod") = ORDER_RECORD.shiplist
			 set ORDER_RECORD = nothing
			 
			 session("SL_BasketCount") = sitelink.GetData_Basket(cstr("BASKETCOUNT"),session("ordernumber"),session("SL_VatInc"))
	         session("SL_BasketSubTotal") = sitelink.GetData_Basket(cstr("BASKETSUBTOTAL"),session("ordernumber"),session("SL_VatInc"))
	         
        'response.Write("<br>")
        'response.Write(session("ordernumber"))
    end if
    
   
    
 
	
	set xmlhttp = server.Createobject("Msxml2.ServerXMLHTTP.6.0")
	        xmlhttp.Open "POST",PayPalUrl,false
            xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			xmlhttp.send DataToSend
			ReturnText = xmlhttp.responseText
			set xmlHTTP=nothing
			
			'response.Write(DataToSend)
			'response.Write("<br>")
			'response.Write("<br>")
			
			'response.Write(PayPalUrl)
			'response.Write("<br>")
			'response.Write("<br>")
			
			'response.Write(ReturnText)
			'response.Write("<br>")
			'response.Write("<br>")
    set xmlhttp = nothing			
			
    set SLCustomer = New cCustomer
    set PayPalExpressObj = New cPayPalExpressObj
        PayPalExpressObj.Action= "VALIDATE_PAYPALCHECKOUTPAYMENT"
        PayPalExpressObj.PayPalPostData = DataToSend
        PayPalExpressObj.PayPalRetMsg = ReturnText
     call sitelink.Getpostdata_PayPalExpress(PayPalExpressObj,SLCustomer)
        OrderNumber =PayPalExpressObj.OrderNum
        ErrorMsg = PayPalExpressObj.ErrorMsg
        HasError = PayPalExpressObj.HasError
        PayerID = PayPalExpressObj.PayerID      
        TokenID = PayPalExpressObj.TokenID
    set PayPalExpressObj = nothing
    set SLCustomer = nothing

    'response.Write(ErrorMsg)
	'response.Write("<br>")
	'response.Write("<br>")
    
    if HasError=true then		
	    set sitelink = nothing
		session("custerrormessage") = 	ErrorMsg +"<br>"
		response.Redirect("PayPalError.asp")			 
	end if	
	
	
	Session.Contents.Remove("TokenID")
	
	if len(trim(TokenID)) > 0 then
	    'ok to confirm the order and redirect to processorder.asp
	    call sitelink.ProcessOrderPayPalExpress(session("ordernumber"), cstr(TokenID))
	    set sitelink= nothing
	    securepage = secureurl + "processorder.asp" 
        response.Redirect(securepage)    
	end if
	
	set sitelink= nothing

    

 %>