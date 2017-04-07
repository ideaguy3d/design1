<%on error resume next%>
<%

if session("cardwasauthed")=true then
	response.redirect("receipt.asp")
end if
%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->


<%
dim DataToSend, xmlHTTP

session("need_to_notify")=false

if session("cardwasauthed")=false and (session("paymentmethod") = "Credit Card" or session("paymentmethod") = "EC") then

    SL_ORD_TOTAL = round(sitelink.ORD_TOTAL(session("ordernumber"),false),2)		
	Orderbalance=round(sitelink.ORDER_BALANCE_REMAINING(session("ordernumber")),2)
	
	sendOrdTotalOnly = false
	if SL_ORD_TOTAL <> Orderbalance then
	       sendOrdTotalOnly=true
	end if
	
	CurrentGateway = sitelink.CHECK_FOR_GATEWAY(session("ordernumber"),session("paymentmethod"))
	
	if ( (CurrentGateway="F" or CurrentGateway="P" or CurrentGateway="U")  and session("paymentmethod") = "EC") then
	    'E check is only for authnet
	    set Sitelink=nothing
	    set ObjDoc=nothing
	    session("cardwasauthed")=true
        response.redirect("processorder.asp")		    
	end if
	
	
	if CurrentGateway="F" then  'First Data
	
	     set SL_PostDataInfo = New cPostDataInfo		
			SL_PostDataInfo.shopperid=session("shopperid")
			SL_PostDataInfo.lordernumber=session("ordernumber")
			SL_PostDataInfo.paymeth=session("paymentmethod")
			SL_PostDataInfo.cc_name=session("cc_name")
			SL_PostDataInfo.cc_number=session("cc_number")
			SL_PostDataInfo.cc_type=session("cc_type")
			SL_PostDataInfo.cc_expmonth=session("cc_expmonth")
			SL_PostDataInfo.cc_expyear  =session("cc_expyear")
			SL_PostDataInfo.cc_id =session("cc_id")
			SL_PostDataInfo.ponum =session("ponum")
			'echeck info
			SL_PostDataInfo.laccttype=session("accttype")
			SL_PostDataInfo.lbankname=session("bankname") 
			SL_PostDataInfo.lbankrountingnum=session("bankrountingnum")
			SL_PostDataInfo.lbankacctnum=session("bankacctnum")		
			
			if sendOrdTotalOnly=true then
				 SL_PostDataInfo.lAuthamount =cdbl(Orderbalance)
			else
				SL_PostDataInfo.lAuthamount=0
			end if 
			
			'Bill to info
			SL_PostDataInfo.lfirstname=session("firstname")
			SL_PostDataInfo.llastname=session("lastname")
			SL_PostDataInfo.lcompany=session("company")
			SL_PostDataInfo.laddress1=session("address1")
			SL_PostDataInfo.lcity=session("city")
			SL_PostDataInfo.lstate=session("state")
			SL_PostDataInfo.lzip=session("zipcode")
			SL_PostDataInfo.lcountry=session("country")
			SL_PostDataInfo.lemail=session("email")
			SL_PostDataInfo.lbilltocopy=session("billtocopy")
			'Ship to 
			SL_PostDataInfo.lsfirstname=session("sfirstname")
			SL_PostDataInfo.lslastname=session("slastname")
			SL_PostDataInfo.lscompany=session("scompany")
			SL_PostDataInfo.lsaddress1=session("saddress1")
			SL_PostDataInfo.lscity=session("scity")
			SL_PostDataInfo.lsstate=session("sstate")
			SL_PostDataInfo.lszip=session("szipcode")
			SL_PostDataInfo.lscountry=session("scountry")
			postdata=""
			postdata = sitelink.getpostdata_FirstData(SL_PostDataInfo)
			
			LoginId = SL_PostDataInfo.lfrom_date
			LoginPwd = SL_PostDataInfo.lissue_number
			
		    set SL_PostDataInfo=nothing    	
		    
		    if len(postdata)=0 then
			        if len(trim(AUTHNET_ERR_NOTIFY))>0 then
			            DataToSend=""			
			            messsg = "Global Gateway data is not published to sitelink." +_
					              VbCrLf + VbCrLf +_
					              "Thank you."					  
			            DataToSend ="msg="+ cstr(server.urlencode(messsg))
			            DataToSend = DataToSend & "&email="  & cstr(AUTHNET_ERR_NOTIFY)
			            DataToSend = DataToSend & "&sitename="  & cstr(insecureurl)
			            DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
			            set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
			            xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink700/authnetnotify.asp",false
                        xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			            xmlhttp.send DataToSend
			            set xmlHTTP=nothing
			        end if		
		        set sitelink=nothing
		        Set ObjDoc=nothing    
		        response.redirect("processorder.asp")	
		   end if	'if len(session("postdata"))=0 
		   
		   
		   'now do the post
		   'test mode..
		   ' StrURL = "https://ws.merchanttest.firstdataglobalgateway.com/fdggwsapi/services/order.wsdl"
		    
		    'Live mode..
		   StrURL = "https://ws.firstdataglobalgateway.com/fdggwsapi/services/order.wsdl"
		   
		   'LoginId = "WS1001236157._.1"
		   'LoginPwd = "f1JZ9RBw"
		    
		    'Set xmlhttp = server.CreateObject("WinHttp.WinHttpRequest.5.1")
		    'xmlhttp.Open "POST",StrURL,false
		    'xmlhttp.setRequestHeader "Content-Type", "text/xml"
		    'xmlhttp.setCredentials LoginId,LoginPwd,0
		    'ClientCertLocation = "LOCAL_MACHINE\My\" + cstr(LoginId)		
		    'xmlhttp.SetClientCertificate ClientCertLocation
		    'xmlhttp.send postdata
		    
		     
		    
		     Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")				     
		     xmlhttp.Open "POST",StrURL,false, LoginId,LoginPwd		     		     
		     xmlhttp.setRequestHeader "Content-Type", "text/xml"			     
		     ClientCertLocation = "LOCAL_MACHINE\My\" + cstr(LoginId)		     
		     xmlhttp.setOption 3,ClientCertLocation		  		    
		     xmlhttp.send postdata
		     

		    attempt=1		      
		    While (xmlhttp.readyState <> 4) and (attempt < 1000000)
			     RetResponse=xmlhttp.waitForResponse(30)			       
			     attempt = attempt + 1			
		    Wend 
            
		    returnreponse = xmlhttp.ResponseText	
		    
		    Set xmlhttp = nothing

		    'Save Postdata		    
		    call sitelink.savepostdataFirstData(session("ordernumber"),postdata,returnreponse)
		    
		    returnreponse = replace(returnreponse,":FDGGWSApiOrderResponse","")
		    returnreponse = replace(returnreponse,"fdggwsapi:","")
		    
		   		   
		    
		    objDoc.loadxml(returnreponse)
		    
		    Rstatus =""
		    RErrMsg=""
		    
		    
		    If objDoc.parseError.errorCode <> 0 Then
		           'Cannot read xml
		           session("cardwasauthed")=true
		           set sitelink=nothing
		           set ObjDoc=nothing
                   response.redirect("processorder.asp")                   
		    else
		     		        
		        set SL_FdataResponse = objDoc.selectNodes("//fdggwsapi") 
		            if SL_FdataResponse.length > 0 then
		                RErrMsg =  SL_FdataResponse.item(0).selectSingleNode("ErrorMessage").text
		                Rstatus =  SL_FdataResponse.item(0).selectSingleNode("TransactionResult").text
		                RApprovalCode = SL_FdataResponse.item(0).selectSingleNode("TransactionID").text
		                RAvsResponse = SL_FdataResponse.item(0).selectSingleNode("AVSResponse").text
		                RTransID = SL_FdataResponse.item(0).selectSingleNode("OrderId").text		                
		                RTDate = SL_FdataResponse.item(0).selectSingleNode("TDate").text		                
		                
		                if sendOrdTotalOnly=true then
                            auth_amt = 	Orderbalance
    	                else
	                        auth_amt = SL_ORD_TOTAL
	                    end if 
	                    
		                call sitelink.WriteAPPRFirstData_New(session("ordernumber"), Rstatus,RApprovalCode, RAvsResponse, RTDate,RTransID, cdbl(auth_amt))	
		                
		            end if
		        
		        set SL_FdataResponse = nothing
		    end if
		    

	        set sitelink=nothing
	        set ObjDoc= nothing
		    		    
		    if (Rstatus<>"APPROVED") then
		        session("ordererrormessage") =""	
		        session("ordererrormessage") ="<ul type=ROUND><li>"& RErrMsg &"</ul>"
		        Response.Redirect("ordererror.asp")    
		    end if 
		    
		    if (Rstatus="APPROVED") then
		        session("cardwasauthed")=true
                response.redirect("processorder.asp")   
		    end if 
    
    end if ' if CurrentGateway=F

        if CurrentGateway="P" then  'Plug and Pay
            set SL_PostDataInfo = New cPostDataInfo		
			SL_PostDataInfo.shopperid=session("shopperid")
			SL_PostDataInfo.lordernumber=session("ordernumber")
			SL_PostDataInfo.paymeth=session("paymentmethod")
			SL_PostDataInfo.cc_name=session("cc_name")
			SL_PostDataInfo.cc_number=session("cc_number")
			SL_PostDataInfo.cc_type=session("cc_type")
			SL_PostDataInfo.cc_expmonth=session("cc_expmonth")
			SL_PostDataInfo.cc_expyear  =session("cc_expyear")
			SL_PostDataInfo.cc_id =session("cc_id")
			SL_PostDataInfo.ponum =session("ponum")
			'echeck info
			SL_PostDataInfo.laccttype=session("accttype")
			SL_PostDataInfo.lbankname=session("bankname") 
			SL_PostDataInfo.lbankrountingnum=session("bankrountingnum")
			SL_PostDataInfo.lbankacctnum=session("bankacctnum")		
			
			if sendOrdTotalOnly=true then
				 SL_PostDataInfo.lAuthamount =cdbl(Orderbalance)
			else
				SL_PostDataInfo.lAuthamount=0
			end if 
			
			'Bill to info
			SL_PostDataInfo.lfirstname=session("firstname")
			SL_PostDataInfo.llastname=session("lastname")
			SL_PostDataInfo.lcompany=session("company")
			SL_PostDataInfo.laddress1=session("address1")
			SL_PostDataInfo.lcity=session("city")
			SL_PostDataInfo.lstate=session("state")
			SL_PostDataInfo.lzip=session("zipcode")
			SL_PostDataInfo.lcountry=session("country")
			SL_PostDataInfo.lemail=session("email")
			SL_PostDataInfo.lbilltocopy=session("billtocopy")
			'Ship to 
			SL_PostDataInfo.lsfirstname=session("sfirstname")
			SL_PostDataInfo.lslastname=session("slastname")
			SL_PostDataInfo.lscompany=session("scompany")
			SL_PostDataInfo.lsaddress1=session("saddress1")
			SL_PostDataInfo.lscity=session("scity")
			SL_PostDataInfo.lsstate=session("sstate")
			SL_PostDataInfo.lszip=session("szipcode")
			SL_PostDataInfo.lscountry=session("scountry")
			session("postdata")=""
			session("postdata") = sitelink.getpostdata_PLUGNPAY(SL_PostDataInfo)	  
		    set SL_PostDataInfo=nothing    			
		
		    if len(session("postdata"))=0 then			
			    if len(trim(AUTHNET_ERR_NOTIFY))>0 then
				    DataToSend=""			
				    messsg = "Plug and Pay merchant account data is not published to sitelink." +_
						      VbCrLf + VbCrLf +_
						      "Thank you."					  
				    DataToSend ="msg="+ cstr(server.urlencode(messsg))
				    DataToSend = DataToSend & "&email="  & cstr(AUTHNET_ERR_NOTIFY)
				    DataToSend = DataToSend & "&sitename="  & cstr(insecureurl)
				    DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
				    set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
					    xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink700/authnetnotify.asp",false
                        xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
					    xmlhttp.send DataToSend
				    set xmlHTTP=nothing
				    end if		
			    set sitelink=nothing
			    Set ObjDoc=nothing
			    response.redirect("processorder.asp")
		    end if 
        
            'now post the data for Plug and Pay.
		    set PNPObj = server.Createobject("pnpcom.main")
		    'Post to PNP  
		    pnp_response = PNPObj.DoTransaction("",session("postdata"),"","")
    		
		    'response.write("<br><br><br>")
		    'response.write(pnp_response)
		    'response.write("<br><br><br>")
		    'response.write("responsetext==>"&session("responsetext"))
		    'response.write("<br>")
		    'response.write("need_to_notify==>"&session("need_to_notify"))
		    set PNPObj =nothing
    		
		   
		    
		    set SLPostData = New cPostData
        		
                 SLPostData.SentData = session("postdata")
                 SLPostData.RecData = pnp_response
                 SLPostData.OrderNumber =  session("ordernumber")             
                 SLPostData.lcc_cid = session("cc_id")
                 SLPostData.lccNumber = session("cc_number")
		         call sitelink.savepostdataPNP(SLPostData)
		         		         		        
		         approval       = SLPostData.approval
		         rrcode         = SLPostData.rrcode
		         overall_code   = SLPostData.overall_code
		         rrtext         = SLPostData.rrtext
		         
		        Set xmlhttp = nothing 
		        set SLPostData = nothing
		        
            session("cardwasauthed")=true
		    set sitelink=nothing
            Set ObjDoc=nothing    
    end if '' if CurrentGateway=P
    
    
        if CurrentGateway="A" then  'Authorize.net
            set SL_PostDataInfo = New cPostDataInfo		
			SL_PostDataInfo.shopperid=session("shopperid")
			SL_PostDataInfo.lordernumber=session("ordernumber")
			SL_PostDataInfo.paymeth=session("paymentmethod")
			SL_PostDataInfo.cc_name=session("cc_name")
			SL_PostDataInfo.cc_number=session("cc_number")
			SL_PostDataInfo.cc_type=session("cc_type")
			SL_PostDataInfo.cc_expmonth=session("cc_expmonth")
			SL_PostDataInfo.cc_expyear  =session("cc_expyear")
			SL_PostDataInfo.cc_id =session("cc_id")
			SL_PostDataInfo.ponum =session("ponum")
			'echeck info
			SL_PostDataInfo.laccttype=session("accttype")
			SL_PostDataInfo.lbankname=session("bankname") 
			SL_PostDataInfo.lbankrountingnum=session("bankrountingnum")
			SL_PostDataInfo.lbankacctnum=session("bankacctnum")		
			
			if sendOrdTotalOnly=true then
				 SL_PostDataInfo.lAuthamount =cdbl(Orderbalance)
			else
				SL_PostDataInfo.lAuthamount=0
			end if 
			
			'Bill to info
			SL_PostDataInfo.lfirstname=session("firstname")
			SL_PostDataInfo.llastname=session("lastname")
			SL_PostDataInfo.lcompany=session("company")
			SL_PostDataInfo.laddress1=session("address1")
			SL_PostDataInfo.lcity=session("city")
			SL_PostDataInfo.lstate=session("state")
			SL_PostDataInfo.lzip=session("zipcode")
			SL_PostDataInfo.lcountry=session("country")
			SL_PostDataInfo.lemail=session("email")
			SL_PostDataInfo.lbilltocopy=session("billtocopy")
			'Ship to 
			SL_PostDataInfo.lsfirstname=session("sfirstname")
			SL_PostDataInfo.lslastname=session("slastname")
			SL_PostDataInfo.lscompany=session("scompany")
			SL_PostDataInfo.lsaddress1=session("saddress1")
			SL_PostDataInfo.lscity=session("scity")
			SL_PostDataInfo.lsstate=session("sstate")
			SL_PostDataInfo.lszip=session("szipcode")
			SL_PostDataInfo.lscountry=session("scountry")
			session("postdata")=""
			session("postdata") = sitelink.getpostdata(SL_PostDataInfo)	  
		    set SL_PostDataInfo=nothing      
				
	        if len(session("postdata"))=0 then
		        'session("rrtext")="Unable to authorize card at this time."
		        'session("ordererrormessage") =""
		        'session("ordererrormessage") ="<ul type=round><li>"& session("rrtext") &"</ul>" 
			        if len(trim(AUTHNET_ERR_NOTIFY))>0 then
			        DataToSend=""			
			        messsg = "Authorize.net data is not published to sitelink." +_
					          VbCrLf + VbCrLf +_
					          "Thank you."					  
			        DataToSend ="msg="+ cstr(server.urlencode(messsg))
			        DataToSend = DataToSend & "&email="  & cstr(AUTHNET_ERR_NOTIFY)
			        DataToSend = DataToSend & "&sitename="  & cstr(insecureurl)
			        DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
			        set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
			        xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink700/authnetnotify.asp",false
                    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			        xmlhttp.send DataToSend
			        set xmlHTTP=nothing
			        end if		
		        set sitelink=nothing
		        Set ObjDoc=nothing    
		        response.redirect("processorder.asp")	
            else
                'set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
		        Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")				
		        
		        'add Ip address of the customer
		        session("postdata") = "x_customer_ip=" + cstr(session("ShopperIPAddress")) +"&" + session("postdata")
		        
		        StrURL="https://secure.authorize.net/gateway/transact.dll"
		        xmlhttp.Open "POST",StrURL,false
		        xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        		
		        xmlhttp.send session("postdata")			
		        attempt=1
		       While (xmlhttp.readyState <> 4) and (attempt < 1000000)			
			        xmlhttp.waitForResponse(30)
			        attempt = attempt + 1			
		        Wend 

		        returnreponse = xmlhttp.ResponseText
		        
        		set SLPostData = New cPostData
        		
                 SLPostData.SentData = session("postdata")
                 SLPostData.RecData = xmlhttp.ResponseText
                 SLPostData.OrderNumber =  session("ordernumber")             
                 SLPostData.lcc_cid = session("cc_id")
                 SLPostData.lccNumber = session("cc_number")
		         call sitelink.savepostdata(SLPostData)
		         		         		        
		         approval       = SLPostData.approval
		         rrcode         = SLPostData.rrcode
		         overall_code   = SLPostData.overall_code
		         rrtext         = SLPostData.rrtext
		         responsetext   = SLPostData.responsetext
		         
		        Set xmlhttp = nothing 
		        set SLPostData = nothing
                
                'if authnet is down.
		        authnetdown=false
		        
		        need_to_notify=false
		        if len(trim(approval)) = 0 and len(trim(rrcode))=0 then
			        authnetdown=true
			        responsetext="BAD"
			        need_to_notify=true
		        end if
        		
		        if rrcode="103" then
		            responsetext="BAD"
			        need_to_notify=true
		        end if
        		
		        if overall_code="2" then
		            responsetext="BAD"
		            need_to_notify=false
		        end if
                
                if responsetext="BAD" and len(trim(AUTHNET_ERR_NOTIFY))>0 and need_to_notify=true then
		        'send email to the customer
			        DataToSend=""
			        'DataToSend ="ecode="& session("rrcode")
			        if authnetdown=true then
				        messsg ="Server Busy at Authorize.net." +_
						         VbCrLf + VbCrLf 				
			        else
				        messsg ="Please contact Dydacomp Technical Support and refer to Authorize.net error code " + cstr(session("rrcode")) +_
						         VbCrLf + VbCrLf +_
						         VbCrLf + "Thank you."
			        end if
			        DataToSend ="msg="+ cstr(server.urlencode(messsg))
			        DataToSend = DataToSend & "&email="  & cstr(AUTHNET_ERR_NOTIFY)
			        DataToSend = DataToSend & "&sitename="  & cstr(insecureurl)
			        DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
			        set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
			        xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink700/authnetnotify.asp",false
                    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			        xmlhttp.send DataToSend
			        'response.write(DataToSend)
			        'Response.Write xmlhttp.ResponseText
			        set xmlHTTP=nothing
		        end if
		        if responsetext="BAD" and need_to_notify=false then
			        session("ordererrormessage") =""			
			        session("ordererrormessage") ="<ul type=ROUND><li>"& rrtext &"</ul>" 
			        set sitelink=nothing
			        Set ObjDoc=nothing    
			        Response.Redirect("ordererror.asp")
		        End If
    		
		        session("cardwasauthed")=true
        	    set sitelink=nothing
        	    Set ObjDoc=nothing    
            
            end if  'postdata is good for authnet
    
    end if 'if CurrentGateway="A" 
    
    
       if CurrentGateway="U" then  'Protx
    
           	set SL_PostDataInfo = New cPostDataInfo		
			SL_PostDataInfo.shopperid=session("shopperid")
			SL_PostDataInfo.lordernumber=session("ordernumber")
			SL_PostDataInfo.paymeth=session("paymentmethod")
			SL_PostDataInfo.cc_name=session("cc_name")
			SL_PostDataInfo.cc_number=session("cc_number")
			SL_PostDataInfo.cc_type=session("cc_type")
			SL_PostDataInfo.cc_expmonth=session("cc_expmonth")
			SL_PostDataInfo.cc_expyear  =session("cc_expyear")
			SL_PostDataInfo.cc_id =session("cc_id")
			SL_PostDataInfo.ponum =session("ponum")
			'echeck info
			SL_PostDataInfo.laccttype=session("accttype")
			SL_PostDataInfo.lbankname=session("bankname") 
			SL_PostDataInfo.lbankrountingnum=session("bankrountingnum")
			SL_PostDataInfo.lbankacctnum=session("bankacctnum")	
			
			SL_PostDataInfo.lfrom_date=session("from_date")
			SL_PostDataInfo.lissue_number=session("issue_num")	
			
			if sendOrdTotalOnly=true then
				 SL_PostDataInfo.lAuthamount =cdbl(Orderbalance)
			else
				SL_PostDataInfo.lAuthamount=0
			end if 
			
			'Bill to info
			SL_PostDataInfo.lfirstname=session("firstname")
			SL_PostDataInfo.llastname=session("lastname")
			SL_PostDataInfo.lcompany=session("company")
			SL_PostDataInfo.laddress1=session("address1")
			SL_PostDataInfo.lcity=session("city")
			SL_PostDataInfo.lstate=session("state")
			SL_PostDataInfo.lzip=session("zipcode")
			SL_PostDataInfo.lcountry=session("country")
			SL_PostDataInfo.lemail=session("email")
			SL_PostDataInfo.lbilltocopy=session("billtocopy")
			'Ship to 
			SL_PostDataInfo.lsfirstname=session("sfirstname")
			SL_PostDataInfo.lslastname=session("slastname")
			SL_PostDataInfo.lscompany=session("scompany")
			SL_PostDataInfo.lsaddress1=session("saddress1")
			SL_PostDataInfo.lscity=session("scity")
			SL_PostDataInfo.lsstate=session("sstate")
			SL_PostDataInfo.lszip=session("szipcode")
			SL_PostDataInfo.lscountry=session("scountry")
			session("postdata")=""			
			session("postdata") = sitelink.getpostdataUK(SL_PostDataInfo)	  
			session("SLTx_code") = SL_PostDataInfo.lSLTx_code
		    set SL_PostDataInfo=nothing      
				
	        if len(session("postdata"))=0 then
		    'session("rrtext")="Unable to authorize card at this time."
		    'session("ordererrormessage") =""
		    'session("ordererrormessage") ="<ul type=round><li>"& session("rrtext") &"</ul>" 
			    if len(trim(AUTHNET_ERR_NOTIFY))>0 then
			        DataToSend=""			
			        messsg = "Authorize.net data is not published to sitelink." +_
					          VbCrLf + VbCrLf +_
					          "Thank you."					  
			        DataToSend ="msg="+ cstr(server.urlencode(messsg))
			        DataToSend = DataToSend & "&email="  & cstr(AUTHNET_ERR_NOTIFY)
			        DataToSend = DataToSend & "&sitename="  & cstr(insecureurl)
			        DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
			        set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
			        xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink700/authnetnotify.asp",false
                    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			        xmlhttp.send DataToSend
			        set xmlHTTP=nothing
			        end if		
		            set sitelink=nothing
		            Set ObjDoc=nothing
		            response.redirect("processorder.asp")	
                else
                    'set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
		            Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")				

        		    'for test account
		            'StrURL="https://ukvpstest.protx.com/VPSDirectAuth/PaymentGateway.asp"								
        		
		            'old url
		            'StrUrl ="https://ukvps.protx.com/VPSDirectAuth/PaymentGateway.asp"		
        		
        	        'New url 
	    	        StrUrl ="https://ukvps.protx.com/vspgateway/service/vspdirect-register.vsp"
		            xmlhttp.Open "POST",StrURL,false
		            'leave this 1 lines commented
		            'this header is needed for new url
		            xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        		
		            'this header is needed for old url
		            'xmlhttp.setRequestHeader "Content-Type", "charset=iso-8859-1" 
    		
		            xmlhttp.send session("postdata")			
		            attempt=1
		            While (xmlhttp.readyState <> 4) and (attempt < 1000000)			
			            xmlhttp.waitForResponse(30)
			            attempt = attempt + 1			
		            Wend 

        		    returnreponse = xmlhttp.ResponseText
    		
    		        call sitelink.savepostdataUK(xmlhttp.ResponseText,TRUE)
	    	        Set xmlhttp = nothing 
                       
		            responsecode = ucase(session("rrcode"))		
		            responsereason = ucase(session("rrtext"))		
        		
		            if responsecode <> "A" then
			            session("responsetext")="BAD"
		            end if
    		    
		            'VRTXInvalidIPAddress i.e invalid account
		            invalidID = false
		            set objRegExp = New Regexp
		            strstring =ucase(responsereason)		
		            objRegExp.Pattern = ucase("VRTXInvalidIPAddress")
		            invalidID = objRegExp.Test(strstring)
		            set objRegExp = nothing
		        
                    if invalidID=true or len(trim(session("approval"))) = 0 then
			           session("need_to_notify")=true			
		            else
			           session("need_to_notify")=false
		            end if
		        
                    if session("responsetext")="BAD" and len(trim(AUTHNET_ERR_NOTIFY))>0 and session("need_to_notify")=true then
		            'send email to the customer
			            DataToSend=""
			            'DataToSend ="ecode="& session("rrcode")
			           if invalidID=true then
				            messsg ="Account is Invalid or Server is not responding" +_
						                VbCrLf + VbCrLf
			           else
			               messsg ="Please contact Dydacomp Technical Support and refer to Authorize.net error code " + cstr(session("rrcode")) +_
				                    VbCrLf + VbCrLf +_
						            VbCrLf + "Thank you."			
			            end if
			            DataToSend ="msg="+ cstr(server.urlencode(messsg))
			            DataToSend = DataToSend & "&email="  & cstr(AUTHNET_ERR_NOTIFY)
			            DataToSend = DataToSend & "&sitename="  & cstr(insecureurl)
			            DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
			            set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
			            xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink700/authnetnotify.asp",false
                        xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			            xmlhttp.send DataToSend
			            'response.write(DataToSend)
			            'Response.Write xmlhttp.ResponseText
			            set xmlHTTP=nothing
		            end if
		            if session("responsetext")="BAD" and session("need_to_notify")=false then
			            session("ordererrormessage") =""			
			            session("ordererrormessage") ="<ul type=ROUND><li>"& session("rrtext") &"</ul>" 
			            set sitelink=nothing
			            Set ObjDoc=nothing
			            Response.Redirect("ordererror.asp")
		            End If
        		
		            session("cardwasauthed")=true
        	        set sitelink=nothing
        	        Set ObjDoc=nothing
        
        end if  'postdata is good for Protx
    
    end if 'if CurrentGateway="A" 
    
    
end if

set sitelink=nothing
set ObjDoc=nothing
'processorder will confirm the order, send email confirmation, update best seller, wish list
response.redirect("processorder.asp")

'response.write(session("postdata"))
'response.write("<br>")
'response.write(session("pnp_response"))


%>