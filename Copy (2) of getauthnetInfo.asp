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
	
            set gl_LPTxn = server.CreateObject("LpiCom_6_0.LinkPointTxn")
            
            set gl_LPOrder = server.CreateObject("LpiCom_6_0.LPOrderPart")
            gl_LPOrder.setPartName("order")
            set gl_LPPart = server.CreateObject("LpiCom_6_0.LPOrderPart")
            
            'Build merchantinfo
            res = gl_LPPart.clear()
            res = gl_LPPart.put("configfile", cstr(trim(sitelink.VERIFYCCACCOUNT(session("cc_type")))))
            
            ' Add merchantinfo to order
            res = gl_LPOrder.addPart("merchantinfo", gl_LPPart)
            
            'Build orderoptions
            res = gl_LPPart.clear()
            res = gl_LPPart.put("ordertype", "PREAUTH")
            'Add orderoptions to order
            res = gl_LPOrder.addPart("orderoptions", gl_LPPart)
            
           'Build transactiondetails
            res = gl_LPPart.clear()
                
            res = gl_LPPart.put("oid", cstr(sitelink.Get_TransIDFirstData(session("ordernumber"))) )
            
           
            
            'level2 ponumber, taxexempt    
            res = gl_LPPart.put("ponumber", cstr(session("ordernumber")))   
            if sitelink.ORD_TAX(session("ordernumber"),false)= 0 then
                res = gl_LPPart.put("taxexempt", "Y")
            else
                res = gl_LPPart.put("taxexempt", "N")
            end if
            
            'Add transactiondetails to order
            res = gl_LPOrder.addPart("transactiondetails", gl_LPPart)
            
            'Build payment
            res = gl_LPPart.clear()  
            if sendOrdTotalOnly=true then	   
	            res = gl_LPPart.put("chargetotal", cstr(Orderbalance))
	        else
		        res = gl_LPPart.put("chargetotal", cstr(SL_ORD_TOTAL))
	        end if 
        			 
            
            res = gl_LPPart.put("tax", cstr(sitelink.ORD_TAX(session("ordernumber"),false)))
            
            'Add payment to order
            res = gl_LPOrder.addPart("payment", gl_LPPart)
            
            'Build creditcard
            res = gl_LPPart.clear()	
            res = gl_LPPart.put("cardnumber", session("cc_number"))
            res = gl_LPPart.put("cardexpmonth", session("cc_expmonth"))
            res = gl_LPPart.put("cardexpyear", right(session("cc_expyear"),2))
            If len(trim(session("cc_id"))) > 0 then
	            res = gl_LPPart.put("cvmvalue", session("cc_id"))
	            res = gl_LPPart.put("cvmindicator", "provided")
            END IF

            'Add creditcard to order
            res = gl_LPOrder.addPart("creditcard", gl_LPPart)
            
            'Build billing
            res = gl_LPPart.clear()
            res = gl_LPPart.put("name", session("cc_name"))
            res = gl_LPPart.put("zip", session("zipcode"))
            res = gl_LPPart.put("addrnum", session("address1"))
            
            'Add billing to order
            res = gl_LPOrder.addPart("billing", gl_LPPart)
            
            'Build notes
            res = gl_LPPart.clear()
            res = gl_LPPart.put("comments", "Web Order")
            'Add notes to order
            res = gl_LPOrder.addPart("notes", gl_LPPart)
            
            'Generate transaction xml 
            postdata = gl_LPOrder.toXML
            
            lccertfile = sitelink.Get_FirstPEMData(session("cc_type"))
            
            StrFileName = "include/"+ cstr(lccertfile)
            PeMfound = false 
	         StrPhysicalPath = Server.MapPath(StrFileName)
		         set objFileName = CreateObject("Scripting.FileSystemObject")	
	            if objFileName.FileExists(StrPhysicalPath) then				
				        PeMfound = true
		        end if 
	        set objFileName = nothing

            lccertfile = StrPhysicalPath
            
            if PeMfound = false then
                lccertfile=""
                if len(trim(AUTHNET_ERR_NOTIFY))>0 then
                        DataToSend=""			
				        messsg = "First Data merchant account PEM is not published to sitelink." +_
						          VbCrLf + VbCrLf +_
						          "Thank you."					  
				        DataToSend ="msg="+ cstr(server.urlencode(messsg))
				        DataToSend = DataToSend & "&email="  & cstr(AUTHNET_ERR_NOTIFY)
				        DataToSend = DataToSend & "&sitename="  & cstr(session("insecureurl"))
				        DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
				        set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
					        xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink600/authnetnotify.asp",false
					        xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
					        xmlhttp.setRequestHeader "Content-Type", "charset=iso-8859-1"		
					        xmlhttp.send DataToSend
				        set xmlHTTP=nothing
		          end if
        		  
		           session("cardwasauthed")=true
		           set sitelink=nothing
		           set ObjDoc=nothing
		           Set gl_LPTxn = Nothing
                   Set gl_LPOrder = Nothing
                   Set gl_LPPart    = Nothing
                   response.redirect("processorder.asp")		   
               end if
           
            'response.Write(postdata)
            'response.Write("<br>")
            
            
            'test url
	        'lpresponse = gl_LPTxn.send(lccertfile, "staging.linkpt.net", 1129, postdata)
            
            'live url
	        lpresponse = gl_LPTxn.send(lccertfile, "secure.linkpt.net", 1129, postdata)
            
         '    response.Write(lpresponse)
         '   response.Write("<br>")
            
            
            Set gl_LPTxn = Nothing
            Set gl_LPOrder = Nothing
            Set gl_LPPart    = Nothing
            
            lpresponse = "<results>"&lpresponse&"</results>"
            
            RErrMsg=""
            Rstatus=""
            RTransID=""
            RApprovalCode=""
            RAvsResponse=""
            
            objDoc.loadxml(lpresponse)
            If objDoc.parseError.errorCode <> 0 Then
                'response.write("<br>Error Processing<br><br>")
                
                session("cardwasauthed")=true
		           set sitelink=nothing
		           set ObjDoc=nothing
		           Set gl_LPTxn = Nothing
                   Set gl_LPOrder = Nothing
                   Set gl_LPPart    = Nothing
                   response.redirect("processorder.asp")
            else
                set SL_FdataResponse = objDoc.selectNodes("//results") 
                    RErrMsg =  SL_FdataResponse.item(0).selectSingleNode("r_message").text
                    Rstatus =  SL_FdataResponse.item(0).selectSingleNode("r_approved").text
                    RTransID =  SL_FdataResponse.item(0).selectSingleNode("r_ordernum").text
                    RApprovalCode = SL_FdataResponse.item(0).selectSingleNode("r_code").text
                    RAvsResponse = SL_FdataResponse.item(0).selectSingleNode("r_avs").text
                set SL_FdataResponse = nothing  
            end if
            
            
            call sitelink.savepostdataFirstData(session("ordernumber"),postdata,lpresponse)
            
            if Rstatus="APPROVED" then
                'write to the database
                if sendOrdTotalOnly=true then
                    auth_amt = 	Orderbalance
    	        else
	                auth_amt = SL_ORD_TOTAL
	            end if 
	            session("cardwasauthed")=true
	            call sitelink.WriteAPPRFirstData(session("ordernumber"),cstr(RApprovalCode),cstr(RApprovalCode),cstr(RAvsResponse),cstr(RTransID),cdbl(auth_amt))
            end if
            
            if Rstatus="FRAUD" or Rstatus="DUPLICATE" or Rstatus="DECLINED"  then    
                    session("ordererrormessage") =""			
			        session("ordererrormessage") ="<ul type=ROUND><li>"& RErrMsg &"</ul>" 
			        set sitelink=nothing
			        set ObjDoc= nothing
			        Response.Redirect("ordererror.asp")    
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
				    DataToSend = DataToSend & "&sitename="  & cstr(session("insecureurl"))
				    DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
				    set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
					    xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink600/authnetnotify.asp",false
					    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
					    xmlhttp.setRequestHeader "Content-Type", "charset=iso-8859-1"		
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
    		
		    call sitelink.savepostdataPNP(pnp_response,TRUE)
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
			        DataToSend = DataToSend & "&sitename="  & cstr(session("insecureurl"))
			        DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
			        set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
			        xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink600/authnetnotify.asp",false
			        xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			        xmlhttp.setRequestHeader "Content-Type", "charset=iso-8859-1"		
			        xmlhttp.send DataToSend
			        set xmlHTTP=nothing
			        end if		
		        set sitelink=nothing
		        Set ObjDoc=nothing    
		        response.redirect("processorder.asp")	
            else
                'set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
		        Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")				

		        'StrURL="https://transact.authorize.net/gateway/transact.dll"				
		        StrURL="https://secure.authorize.net/gateway/transact.dll"
		        xmlhttp.Open "POST",StrURL,false
		        xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		        'xmlhttp.setRequestHeader "Content-Type", "charset=iso-8859-1" 
        		
		        xmlhttp.send session("postdata")			
		        attempt=1
		       While (xmlhttp.readyState <> 4) and (attempt < 1000000)			
			        xmlhttp.waitForResponse(30)
			        attempt = attempt + 1			
		        Wend 

		        returnreponse = xmlhttp.ResponseText
        		
		        call sitelink.savepostdata(xmlhttp.ResponseText,TRUE)
		        Set xmlhttp = nothing 
                
                'if authnet is down.
		        authnetdown=false
		        session("need_to_notify")=false
		        if len(trim(session("approval"))) = 0 and len(trim(session("rrcode")))=0 then
			        authnetdown=true
			        session("responsetext")="BAD"
			        session("need_to_notify")=true
		        end if
        		
		        if session("rrcode")="103" then
		            session("responsetext")="BAD"
			        session("need_to_notify")=true
		        end if
        		
		        if session("overall_code")="2" then
		            session("responsetext")="BAD"
		            session("need_to_notify")=false
		        end if
                
                if session("responsetext")="BAD" and len(trim(AUTHNET_ERR_NOTIFY))>0 and session("need_to_notify")=true then
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
			        DataToSend = DataToSend & "&sitename="  & cstr(session("insecureurl"))
			        DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
			        set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
			        xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink600/authnetnotify.asp",false
			        xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			        xmlhttp.setRequestHeader "Content-Type", "charset=iso-8859-1"		
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
			        DataToSend = DataToSend & "&sitename="  & cstr(session("insecureurl"))
			        DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
			        set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
			        xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink600/authnetnotify.asp",false
			        xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			        xmlhttp.setRequestHeader "Content-Type", "charset=iso-8859-1"		
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
			            DataToSend = DataToSend & "&sitename="  & cstr(session("insecureurl"))
			            DataToSend = DataToSend & "&Ordnum="  & cstr(session("ordernumber"))
			            set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
			            xmlhttp.Open "POST","http://mail.mailordercentral.com/sitelink600/authnetnotify.asp",false
			            xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			            xmlhttp.setRequestHeader "Content-Type", "charset=iso-8859-1"		
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
    
    end if
    
end if

set sitelink=nothing
set ObjDoc=nothing
'processorder will confirm the order, send email confirmation, update best seller, wish list
response.redirect("processorder.asp")

'response.write(session("postdata"))
'response.write("<br>")
'response.write(session("pnp_response"))


%>