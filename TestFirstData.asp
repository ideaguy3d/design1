<%'on error resume next%>


<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->



<%

session("paymentmethod") = "Credit Card"
session("cc_name") = "Test"
session("cc_number") = "4444333322221111"
session("cc_type") ="VI"
session("cc_expmonth") = "1"
session("cc_expyear") ="2016"
session("cc_id") ="1234"
session("ponum") = "1234"

session("accttype") =""
session("bankname") =""
session("bankrountingnum")=""
session("bankacctnum")=""



CurrentGateway = sitelink.CHECK_FOR_GATEWAY(session("ordernumber"),session("paymentmethod"))


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
			
			Orderbalance = "1.00"
		    SL_PostDataInfo.lAuthamount=0
			
			
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
			
		    'response.Write(postdata)
		    
		      'Live mode..
		   StrURL = "https://ws.firstdataglobalgateway.com/fdggwsapi/services/order.wsdl"
		   
		   
		    Set xmlhttp = server.CreateObject("Msxml2.ServerXMLHTTP.3.0")				     
		     xmlhttp.Open "POST",StrURL,false, LoginId,LoginPwd		     		     
		     xmlhttp.setRequestHeader "Content-Type", "text/xml"			     
		     ClientCertLocation = "LOCAL_MACHINE\My\" + cstr(LoginId)		     
		     xmlhttp.setOption 3,ClientCertLocation		  		    
		     xmlhttp.send postdata
		     
		     returnreponse = xmlhttp.ResponseText
		     
		     response.Write(returnreponse)
    else
        response.Write("site is not set up with firsdata")
	end if    

set SiteLink = nothing
set objDoc = nothing		    

 %>

