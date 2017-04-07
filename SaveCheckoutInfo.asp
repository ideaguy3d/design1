<%on error resume next%>
<% 
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%> 

<!--#INCLUDE FILE = "include/momapp.asp" -->


<%

    whatsourcekey            = cstr(request("skeysel"))
    session("orderholdDate") = request("orderholddt")
    session("memo_field1")   = request("ordermemo1")
    session("memo_field2")   = request("ordermemo2")
    session("memo_field3")   = request("ordermemo3")
    session("ordermemo")     = request("orderfullfill")
    
    ship_ahead               = cstr(REQUEST("Ship_aheadflag"))
        
    if len(trim(ship_ahead)) = 0 then    
        ship_ahead=0
    end if
    
    
    session("memo_field1")=fix_xss_Chars(session("memo_field1"))
	session("memo_field2")=fix_xss_Chars(session("memo_field2"))
	session("memo_field3")=fix_xss_Chars(session("memo_field3"))
	session("orderholdDate")=fix_xss_Chars(session("orderholdDate"))
	session("ordermemo")=fix_xss_Chars(session("ordermemo"))
	
	if session("ordermemo")<>"" then
		session("ordermemo") = left(session("ordermemo"),1200)
	end if
	
	session("paymentmethod")="PP"
	
	 set SL_PayInfo = New cPayment	 
	 SL_PayInfo.shopperid		= session("shopperid")
	 SL_PayInfo.lordernumber	=session("ordernumber")
     SL_PayInfo.paymeth 		=session("paymentmethod")
	 SL_PayInfo.cc_name 		=session("cc_name") 
	 SL_PayInfo.cc_number 		=session("cc_number")   
	 SL_PayInfo.cc_number 		=session("cc_number")  
	 SL_PayInfo.cc_type 		=session("cc_type")  
	 SL_PayInfo.cc_expmonth 	=session("cc_expmonth")
	 SL_PayInfo.cc_expyear	 	=session("cc_expyear")
	 SL_PayInfo.lponumber	 	=session("ponum")
	 SL_PayInfo.laccttype	 	=session("accttype")
	 SL_PayInfo.lbankname	 	=session("bankname")	 
	 SL_PayInfo.lbankrountingnum=session("bankrountingnum")
	 SL_PayInfo.lbankacctnum	=session("bankacctnum")
	 SL_PayInfo.lordercommt1	=session("memo_field1")
	 SL_PayInfo.lordercommt2	=session("memo_field2")
	 SL_PayInfo.lordercommt3	=session("memo_field3")
	 SL_PayInfo.lgiftmsg1		=session("giftmsg1")
	 SL_PayInfo.lgiftmsg2		=session("giftmsg2")
	 SL_PayInfo.lgiftmsg3		=session("giftmsg3")
	 SL_PayInfo.lgiftmsg4		=session("giftmsg4")
	 SL_PayInfo.lgiftmsg5		=session("giftmsg5")
	 SL_PayInfo.lgiftmsg6		=session("giftmsg6")
	 SL_PayInfo.lfullfill		=session("ordermemo")
	 'SL_PayInfo.lfrom_date		=""
	 'SL_PayInfo.lissue_number	=""
	 SL_PayInfo.lship_ahead		=ship_ahead
	 SL_PayInfo.lorder_holdDate	=session("orderholdDate")
	 SL_PayInfo.lcl_key			=cstr(whatsourcekey)
	 'points redemption
	 SL_PayInfo.lpointsused		=session("Points_need")
	 SL_PayInfo.lcheckamount	=session("Redeem_Amount")
	 SL_PayInfo.lpaybypoints	=session("PaybyPoints")
	 'GC redemption
	 SL_PayInfo.lPaybyGC		=session("PaybyGC")
	 SL_PayInfo.lGCNumber		=session("GCnumber")

	 if session("paymentmethod") <>"CK" then
	 	session("Redeem_Amount")=0
	 end if
		
	'if session("DoNotShow_PointsTab")=false then 
	' 	Previous_BasketSubTotal = sitelink.GetData_Basket(cstr("BASKETSUBTOTAL"),session("ordernumber"),session("SL_VatInc"))	  
	'end if
	
	session("ordererrormessage")=sitelink.checkorderinfo(SL_PayInfo)
	
	set SL_PayInfo = nothing
	
	set Sitelink=Nothing
	

	
	


 %>