<%on error resume next%>

<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 
<%
if session("SL_BasketCount")= 0 then
	response.redirect("basket.asp")
end if
	
%>
<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%

    WhichCheckout= cstr(request.Form("WhichCheckout"))
    
    'WhichCheckout = BillMeLater or PayPalExpress

        PayPalUrl = ""
        
    if WhichCheckout = "PayPalExpress" then
        set SLCustomer = New cCustomer
        set PayPalExpressObj = New cPayPalExpressObj
            PayPalExpressObj.Action= "SETEEXPRESSCHECKOUT"
            PayPalExpressObj.OrderNum=session("ordernumber")
            PayPalExpressObj.OrdAmount=session("SL_BasketSubTotal")
            PayPalExpressObj.InSecureUrl=insecureurl
            PayPalExpressObj.SecureUrl=secureurl
            call sitelink.Getpostdata_PayPalExpress(PayPalExpressObj,SLCustomer)
            
            DataToSend =PayPalExpressObj.PayPalPostData        
                
            DataToSend = DataToSend + "&HDRIMG=" + secureurl + "images/" + COMPANY_LOGO_IMG            
            extrafield =""
		    xmlstring = sitelink.getbasketinfo(session("shopperid"),session("ordernumber"),extrafield,false,false)
		    objDoc.loadxml(xmlstring)
		    set SL_Basket = objDoc.selectNodes("//gbi[certid=0 and it_unlist>0]")
		    
		    TotalLineItems = SL_Basket.length-1
		    
		    if TotalLineItems > 99 then
		        DataToSend = DataToSend + "&AMT=" & round(session("SL_BasketSubTotal"),2)
		    else
		        basket_total = 0 
		        for x=0 to SL_Basket.length-1
		        
		            SL_Number	= SL_Basket.item(x).selectSingleNode("item").text
		            SL_BasketDesc = SL_Basket.item(x).selectSingleNode("inetsdesc").text
		            SL_BasketUnitPrice	= SL_Basket.item(x).selectSingleNode("extendp").text
		            SL_BasketUnitPrice = round(SL_BasketUnitPrice,2)
		            		            
		            actual_Quantity = SL_Basket.item(x).selectSingleNode("quanto").text
		            actual_Quantity = replace(actual_Quantity,".00","")
		            actual_Unitprice = SL_Basket.item(x).selectSingleNode("it_unlist").text
		            
		            actual_Unitprice = round(actual_Unitprice,2)
		            option_val = actual_Quantity & "@" & actual_Unitprice
		 
		            if SL_Number="" then
		                SL_Number="Order Promotion Discount"
		                SL_BasketDesc = "Order Promotion Discount"
		            end if
                    
                    if PayPalExpressObj.IsPayFlowAccount = true then
                        L_NUMBER = "L_SKU" & x & "="
		                L_NAME = "L_NAME"& x & "="
		                L_COST = "L_COST" & x & "="
		                'L_AMT = "L_AMT" & x & "="
		                L_QTY = "L_QTY" & x & "=1"		            
		                L_DESC= "L_DESC"& x & "=" & option_val

		                DataToSend = DataToSend + "&" & L_NUMBER & SL_Number
		                DataToSend = DataToSend + "&" & L_NAME & (SL_BasketDesc)
		                'DataToSend = DataToSend + "&" & L_AMT & SL_BasketUnitPrice
		                DataToSend = DataToSend + "&" & L_COST & SL_BasketUnitPrice
		                DataToSend = DataToSend + "&" & L_QTY 
		                DataToSend = DataToSend + "&" & L_DESC 
                    else
                        L_NUMBER = "L_PAYMENTREQUEST_0_NUMBER" & x & "="
		                L_NAME = "L_PAYMENTREQUEST_0_NAME"& x & "="		                
		                L_AMT = "L_PAYMENTREQUEST_0_AMT" & x & "="
		                L_QTY = "L_PAYMENTREQUEST_0_QTY" & x & "=1"		            
		                L_DESC= "L_PAYMENTREQUEST_0_DESC"& x & "=" & option_val

		                DataToSend = DataToSend + "&" & L_NUMBER & SL_Number
		                DataToSend = DataToSend + "&" & L_NAME & (SL_BasketDesc)
		                DataToSend = DataToSend + "&" & L_AMT & SL_BasketUnitPrice		                
		                DataToSend = DataToSend + "&" & L_QTY 
		                DataToSend = DataToSend + "&" & L_DESC 
                    end if

		            basket_total = basket_total + SL_BasketUnitPrice
		        next
		        set SL_Basket = nothing
		        
		        basket_total = round(basket_total,2)
		        
		        if PayPalExpressObj.IsPayFlowAccount = true then
		            DataToSend = DataToSend + "&AMT=" & basket_total &"&ITEMAMT=" & basket_total		        
		        else
		            DataToSend = DataToSend + "&PAYMENTREQUEST_0_AMT=" & basket_total &"&PAYMENTREQUEST_0_ITEMAMT=" & basket_total
		            
		        end if
		        

		    end if
		    
            PayPalUrl = PayPalExpressObj.PayPalPostUrl
        
        set PayPalExpressObj = nothing
        set SLCustomer = nothing
        
        
        set xmlhttp = server.Createobject("Msxml2.ServerXMLHTTP.6.0")
	        xmlhttp.Open "POST",PayPalUrl,false
            xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			xmlhttp.send DataToSend
			ReturnText = xmlhttp.responseText
			set xmlHTTP=nothing     
			
			
			HasError = false
			PayPalUrl = ""
			
			set SLCustomer = New cCustomer
			set PayPalExpressObj = New cPayPalExpressObj
			    PayPalExpressObj.Action= "VALIDATE_SETEEXPRESSCHECKOUT"
			    PayPalExpressObj.WhatMethodUsed= "PPE"
                PayPalExpressObj.OrderNum=session("ordernumber")
                PayPalExpressObj.PayPalRetMsg = ReturnText
                PayPalExpressObj.PayPalPostData =DataToSend
                
                call sitelink.Getpostdata_PayPalExpress(PayPalExpressObj,SLCustomer)
                ErrorMsg = PayPalExpressObj.ErrorMsg
                HasError = PayPalExpressObj.HasError
                PayPalUrl = PayPalExpressObj.PayPalRedirectUrl
			set PayPalExpressObj = nothing
			set SLCustomer = nothing
			
			
			set sitelink = nothing
			
			'response.Write(DataToSend)
			'response.Write("<Br>")
			'response.Write("<Br>")
			'response.Write(ReturnText)
			'response.Write("<Br>")
			'response.Write("<Br>")
			'response.Write("PayPalUrl -->" & PayPalUrl)
			'response.Write("<Br>")
			'response.Write("<Br>")
			'response.Write("HasError -->" & HasError)
			'response.Write("<Br>")
			'response.Write("<Br>")
			'response.Write("ErrorMsg -->" & ErrorMsg)
			
			if HasError=false then			    
			    response.Redirect(PayPalUrl)	
			else			    
			    session("custerrormessage") = 	ErrorMsg +"<br>"
			    response.Redirect("PayPalError.asp")			 
			end if		
	end if 'if WhichCheckout = "PayPalExpress" 
	
	if WhichCheckout = "BillMeLater" then
	
	    set SLCustomer = New cCustomer
        set PayPalExpressObj = New cPayPalExpressObj
            PayPalExpressObj.Action= "SETEEXPRESSCHECKOUT_BILLMELATER"
            PayPalExpressObj.OrderNum=session("ordernumber")
            PayPalExpressObj.OrdAmount=session("SL_BasketSubTotal")
            PayPalExpressObj.InSecureUrl=insecureurl
            PayPalExpressObj.SecureUrl=secureurl
            call sitelink.Getpostdata_PayPalExpress(PayPalExpressObj,SLCustomer)
            
            DataToSend =PayPalExpressObj.PayPalPostData
            PayPalUrl = PayPalExpressObj.PayPalPostUrl
        
        
        set SLCustomer = nothing
	
	    DataToSend = DataToSend + "&HDRIMG=" + secureurl + "images/" + COMPANY_LOGO_IMG
	   
	    extrafield =""
		    xmlstring = sitelink.getbasketinfo(session("shopperid"),session("ordernumber"),extrafield,false,false)
		    objDoc.loadxml(xmlstring)
		    set SL_Basket = objDoc.selectNodes("//gbi[certid=0 and it_unlist>0]")
		    
		    TotalLineItems = SL_Basket.length-1
		    
		    if TotalLineItems > 99 then
		        DataToSend = DataToSend + "&AMT=" & round(session("SL_BasketSubTotal"),2)
		    else
		        basket_total = 0 
		        for x=0 to SL_Basket.length-1
		        
		            SL_Number	= SL_Basket.item(x).selectSingleNode("item").text
		            SL_BasketDesc = SL_Basket.item(x).selectSingleNode("inetsdesc").text
		            SL_BasketUnitPrice	= SL_Basket.item(x).selectSingleNode("extendp").text
		            SL_BasketUnitPrice = round(SL_BasketUnitPrice,2)
		            		            
		            actual_Quantity = SL_Basket.item(x).selectSingleNode("quanto").text
		            actual_Quantity = replace(actual_Quantity,".00","")
		            actual_Unitprice = SL_Basket.item(x).selectSingleNode("it_unlist").text
		            
		            actual_Unitprice = round(actual_Unitprice,2)
		            option_val = actual_Quantity & "@" & actual_Unitprice
		 
		            if SL_Number="" then
		                SL_Number="Order Promotion Discount"
		                SL_BasketDesc = "Order Promotion Discount"
		            end if
		        
		           if PayPalExpressObj.IsPayFlowAccount = true then
                        L_NUMBER = "L_SKU" & x & "="
		                L_NAME = "L_NAME"& x & "="
		                L_COST = "L_COST" & x & "="
		                'L_AMT = "L_AMT" & x & "="
		                L_QTY = "L_QTY" & x & "=1"		            
		                L_DESC= "L_DESC"& x & "=" & option_val

		                DataToSend = DataToSend + "&" & L_NUMBER & SL_Number
		                DataToSend = DataToSend + "&" & L_NAME & (SL_BasketDesc)
		                'DataToSend = DataToSend + "&" & L_AMT & SL_BasketUnitPrice
		                DataToSend = DataToSend + "&" & L_COST & SL_BasketUnitPrice
		                DataToSend = DataToSend + "&" & L_QTY 
		                DataToSend = DataToSend + "&" & L_DESC 
                    else
                        L_NUMBER = "L_PAYMENTREQUEST_0_NUMBER" & x & "="
		                L_NAME = "L_PAYMENTREQUEST_0_NAME"& x & "="		                
		                L_AMT = "L_PAYMENTREQUEST_0_AMT" & x & "="
		                L_QTY = "L_PAYMENTREQUEST_0_QTY" & x & "=1"		            
		                L_DESC= "L_PAYMENTREQUEST_0_DESC"& x & "=" & option_val

		                DataToSend = DataToSend + "&" & L_NUMBER & SL_Number
		                DataToSend = DataToSend + "&" & L_NAME & (SL_BasketDesc)
		                DataToSend = DataToSend + "&" & L_AMT & SL_BasketUnitPrice		                
		                DataToSend = DataToSend + "&" & L_QTY 
		                DataToSend = DataToSend + "&" & L_DESC 
                    end if

		            basket_total = basket_total + SL_BasketUnitPrice
		        next
		        set SL_Basket = nothing
		        
		        basket_total = round(basket_total,2)
		        
		        if PayPalExpressObj.IsPayFlowAccount = true then
		            DataToSend = DataToSend + "&AMT=" & basket_total &"&ITEMAMT=" & basket_total		        
		        else
		            DataToSend = DataToSend + "&PAYMENTREQUEST_0_AMT=" & basket_total &"&PAYMENTREQUEST_0_ITEMAMT=" & basket_total		            
		        end if
		        
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
		'response.Write(ReturnText)
		
		HasError = false
			PayPalUrl = ""
			
			set SLCustomer = New cCustomer
			set PayPalExpressObj = New cPayPalExpressObj
			    PayPalExpressObj.Action= "VALIDATE_SETEEXPRESSCHECKOUT"
			    PayPalExpressObj.WhatMethodUsed= "BML"
                PayPalExpressObj.OrderNum=session("ordernumber")
                PayPalExpressObj.PayPalRetMsg = ReturnText
                PayPalExpressObj.PayPalPostData =DataToSend
                
                call sitelink.Getpostdata_PayPalExpress(PayPalExpressObj,SLCustomer)
                ErrorMsg = PayPalExpressObj.ErrorMsg
                HasError = PayPalExpressObj.HasError
                PayPalUrl = PayPalExpressObj.PayPalRedirectUrl
			set PayPalExpressObj = nothing
			set SLCustomer = nothing
			
			
			set sitelink = nothing
			
			
			if HasError=false then			    
			    response.Redirect(PayPalUrl)	
			else			    
			    session("custerrormessage") = 	ErrorMsg +"<br>"
			    response.Redirect("PayPalError.asp")			 
			end if		
			
		
	end if 

 %>