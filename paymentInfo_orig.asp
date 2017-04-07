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


<%

	'orderconfirm = sitelink.ORDER_CONFIRMED(session("ordernumber"))
	'if orderconfirm = true then
	'  	set sitelink=nothing	
	'	response.redirect("receipt.asp")
	'end if	

	session("paymentmethod")=cstr(REQUEST.FORM("payment_type"))
	session("cc_name")=cstr(trim(REQUEST.FORM("txtcc_name")))
	session("cc_number")=cstr(trim(REQUEST.FORM("txtcc_number")))
	session("cc_type")=cstr(REQUEST.FORM("cc_type"))
	session("cc_expmonth")=cstr(REQUEST.FORM("cc_expmonth"))
	session("cc_expyear")=cstr(REQUEST.FORM("cc_expyear"))
	
	session("accttype") =cstr(trim(REQUEST.FORM("accttype")))
	session("bankname") =cstr(trim(REQUEST.FORM("bankname")))
	session("bankrountingnum") =cstr(trim(REQUEST.FORM("bankrountingnum")))
	session("bankacctnum") =cstr(trim(REQUEST.FORM("bankacctnum")))
	session("cc_id") = cstr(trim(REQUEST.FORM("txtcc_id")))
	
	
	session("giftmsg1")=cstr(REQUEST.FORM("txtgiftmsg1"))
	session("giftmsg2")=cstr(REQUEST.FORM("txtgiftmsg2"))
	session("giftmsg3")=cstr(REQUEST.FORM("txtgiftmsg3"))
	session("giftmsg4")=cstr(REQUEST.FORM("txtgiftmsg4"))
	session("giftmsg5")=cstr(REQUEST.FORM("txtgiftmsg5"))
	session("giftmsg6")=cstr(REQUEST.FORM("txtgiftmsg6"))
	
	session("memo_field1") = cstr(REQUEST.FORM("memo_field1"))
	session("memo_field2") = cstr(REQUEST.FORM("memo_field2"))
	session("memo_field3") = cstr(REQUEST.FORM("memo_field3"))
	
	session("orderholdDate") = cstr(trim(REQUEST.FORM("orderhold")))
	
	
	session("ponum") = cstr(REQUEST.FORM("ponum"))
	
	session("PaybyPoints") = REQUEST.FORM("Applypoints")
	session("PaybyGC")     = REQUEST.FORM("applygc")
	session("GCnumber")    = cstr(REQUEST.FORM("gcnumber"))

	
	
	
		session("cc_name")=fix_xss_Chars(session("cc_name"))
		session("cc_number")=fix_xss_Chars(session("cc_number"))
		session("giftmsg1")=fix_xss_Chars(session("giftmsg1"))
		session("giftmsg2")=fix_xss_Chars(session("giftmsg2"))
		session("giftmsg3")=fix_xss_Chars(session("giftmsg3"))
		session("giftmsg4")=fix_xss_Chars(session("giftmsg4"))
		session("giftmsg5")=fix_xss_Chars(session("giftmsg5"))
		session("giftmsg6")=fix_xss_Chars(session("giftmsg6"))
		session("memo_field1")=fix_xss_Chars(session("memo_field1"))
		session("memo_field2")=fix_xss_Chars(session("memo_field2"))
		session("memo_field3")=fix_xss_Chars(session("memo_field3"))
		session("orderholdDate")=fix_xss_Chars(session("orderholdDate"))
		session("ponum")=fix_xss_Chars(session("ponum"))
		session("cc_id")=fix_xss_Chars(session("cc_id"))	
		session("accttype") = fix_xss_Chars(session("accttype"))
		session("bankname") = fix_xss_Chars(session("bankname"))
		session("bankrountingnum") = fix_xss_Chars(session("bankrountingnum"))
		session("bankacctnum") = fix_xss_Chars(session("bankacctnum"))
		session("GCnumber") = fix_xss_Chars(session("GCnumber"))

	
	ordermemotxt = cstr(trim(REQUEST.FORM("ordermemo")))
	ship_ahead   = cstr(REQUEST.FORM("ship_ahead"))
	
	whatsourcekey = cstr(trim(REQUEST.FORM("whatsourcekey")))
	
	'ordermemolength = len(ordermemotxt)	
	session("ordermemo") = ordermemotxt
	
	session("ordermemo")=fix_xss_Chars(session("ordermemo"))
	
	
	if session("ordermemo")<>"" then
		session("ordermemo") = left(ordermemotxt,1200)
	end if
		
	if len(trim(session("paymentmethod")))= 0 then
		session("ordererrormessage")="<ul><li>Please select Payment Method</ul>"
		set sitelink=nothing 
		response.redirect("ordererror.asp")		
	end if
	
	if session("PaybyGC")=1 and isnumeric(session("GCnumber"))=false then
		session("ordererrormessage")="<ul><li>Invalid Gift Certificate Number</li></ul>"
		set sitelink=nothing 
		response.redirect("ordererror.asp")
	end if
	
	
	
	'check cctype v/s number
	if session("paymentmethod") = "Credit Card" then
		currenttype = left(session("cc_type"),1)
		if len(session("cc_number")) > 0 then
			cnumber = left(session("cc_number"),1)
		else
			cnumber="0"
		end if 

	
		correcttype =""
		select case cnumber
			  case 3
				correcttype = "A"
			  case 4
				correcttype = "V"
			  case 5
				correcttype = "M"
		      case 6
				correcttype = "D"		 
		END SELECT
	
		if len(correcttype)= 0  then
			set sitelink=nothing 
			session("ordererrormessage")="<ul><li>Please select Correct Credit Card Type</ul>"
			response.redirect("ordererror.asp")
		end if
		
		if len(correcttype)> 0 and correcttype <> currenttype then
			set sitelink=nothing 
			session("ordererrormessage")="<ul><li>Please select Correct Credit Card Type</ul>"
			response.redirect("ordererror.asp")
		end if
	
		if WANT_AUTHNET =1 and WANT_CID_MANDATORY =1 and len(trim(session("cc_id"))) = 0 then		
			
			session("ordererrormessage") = "<ul><li>Credit Card ID is Required </ul>" +_											
												"<a class=""allpage"" href=""javascript:openlearnmore();"">Learn More about Credit Card ID</a>" +_
											"<br><br>"
			set sitelink=nothing 
			response.redirect("ordererror.asp")
		end if	
	end if
	
	
	set SL_PayInfo = New cPayment
	 call InitcPayment(SL_PayInfo)
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
	 SL_PayInfo.lcl_key			=whatsourcekey
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
	
	session("ordererrormessage")=sitelink.checkorderinfo(SL_PayInfo)
	
	GCAmountLeft =SL_PayInfo.lGCAmountLeft
	GCItemNumber = SL_PayInfo.lGCItemNumber
	'OrderRemaingBalance = SL_PayInfo.lOrder_Balance  
	
	set SL_PayInfo = nothing
	
	'response.Write(GCAmountLeft)
	
	aa= session("ordererrormessage") 	
	
	'response.Write("This is error msg -->"&aa)
	if len(trim(aa))>0 then 
		set sitelink=nothing 
		response.redirect("ordererror.asp")	
	end if
	
	
	if session("paymentmethod")="Terms" then
		set sitelink=nothing
		response.redirect("processorder.asp")
	end if

	if session("PaybyPoints")=1 then
		'redeeming against order total
		Remaining_Orderbalance=sitelink.ORDER_BALANCE_REMAINING(session("ordernumber"))
		'if session("Remaining_balance") > 0 and session("paymentmethod")="CK" then
		if Remaining_Orderbalance > 0 and session("paymentmethod")="CK" then
			session("DoNotShow_PointsTab")=true
			session("paymentmethod")=""
			set sitelink=nothing		
			response.redirect("checkout.asp")
		end if
	end if
	
	if session("PaybyGC")=1 then
		Remaining_Orderbalance=sitelink.ORDER_BALANCE_REMAINING(session("ordernumber"))
		if GCAmountLeft <= 0 then
	        session("ordererrormessage") ="<ul>There is no balance left on this gift certificate</ul>"
	        set sitelink=nothing
	        response.redirect("ordererror.asp")	
	    end if
		
		'if OrderRemaingBalance > 0 then		   
		if Remaining_Orderbalance > 0 then
	        session("DoNotShow_GCTab")=true
	        session("paymentmethod")=""
	        set sitelink=nothing
			'response.redirect("checkout.asp")
	    end if
	    
	    
	end if 
	
	'response.Write(Remaining_Orderbalance)
	
'	response.write("Current_total-->"&Current_total)
'	response.write("<br>")
'	response.write("Remaining_balance-->"&session("Remaining_balance"))
'	response.write("<br>")
'	response.write("SL_BasketSubTotal -->"&session("SL_BasketSubTotal"))
'	response.write("<br>")




	if session("paymentmethod")="CK" then
		set sitelink=nothing
		response.redirect("processorder.asp")
	end if

	if session("paymentmethod")="COD" then
    	'to adjust cod charges if any
		call sitelink.TOTALORDER(session("shopperid"),session("ordernumber"),session("billtocopy"),session("shippingmethod"))
		set sitelink=nothing
		response.redirect("processorder.asp")
	end if


	set sitelink=nothing

	if session("paymentmethod") = "Credit Card" then
		if WANT_AUTHNET= 1 and session("cardwasauthed")=false then
			response.redirect("page_wait.asp")			
		else
			response.redirect("processorder.asp")
		end if	
	end if

	if session("paymentmethod") = "EC" then
		if WANT_ECHECK= 1 and session("cardwasauthed")=false then
			response.redirect("page_wait.asp")
		else
			response.redirect("processorder.asp")
		end if	
	end if

	
	
%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Aussie Products.com | Customer Payment Information @ 2012 | Australian Products Co. Bringing Australia to You - Aussie Foods</title>
	<meta http-equiv="Content-Type" content="application/xhtml+xml; charset=utf-8" />
	<meta name="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls" />
	<meta name="keywords" content="Aussie Products, Australian, Aussie, Australian Products, Australian Food, Aussie Foods, Australian Desserts, tim tams, vegemite, violet crumble, cherry ripe, koala, kangaroo, down under, Australian Groceries, foods, lollies, candy, gourmet food" />
	<meta name="robots" content="index, follow" />
    <meta name="robots" content="ALL">
    <!-- Update your html tag to include the itemscope and itemtype attributes -->
<html itemscope itemtype="http://schema.org/LocalBusiness">


<meta itemprop="name" content="Aussie Products.com | Australian Products Co. Bringing Australia to You - Aussie Foods">
<meta itemprop="description" content="Australian Products Co. Bringing Australia to You. Australian Products include food, books, home decor, souvenirs, clothing, Australian food, Akubra Hats, Driza-Bone coats, Educational Material, Australian Meat Pies and Sausage Rolls">
<meta itemprop="image" content="http://www.aussieproducts.com/">
	<meta name="distribution" content="Global">
	<meta name="revisit-after" content="5">
	<meta name="classification" content="Food">
    <meta name="google-site-verification" content="K2A2-bw3DiKf9q0_Ul-o-hJLTo3YkFNb_JHtsYJ0LJ0"
	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />

<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="text/topnav.css" type="text/css">
<link rel="stylesheet" href="text/sidenav.css" type="text/css">
<link rel="stylesheet" href="text/storestyle.css" type="text/css">
<link rel="stylesheet" href="text/footernav.css" type="text/css">
<link rel="stylesheet" href="text/design.css" type="text/css">
</head>
<!--#INCLUDE FILE = "text/Background5.asp" -->



<div id="container">
    <!--#INCLUDE FILE = "include/top_nav.asp" -->
    <div id="main">



<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	
<td width="<%=SIDE_NAV_WIDTH%>" class="sidenavbg" valign="top"  >
<!--#INCLUDE FILE = "include/side_nav.asp" -->
	</td>	
	<td valign="top" class="pagenavbg">

<!-- sl code goes here -->
<div id="page-content">
	<center>
	
	<% if session("paymentmethod") = "PP" then %>
	<table width="100%" border="0" cellspacing="0" cellpadding="2">
	<tr><td colspan="2" class="THHeader" align="center">Pay by PayPal&nbsp;&nbsp;</td></tr>
	<tr><td align="center" class="plaintextbold">
	<br><br>
	<Font color=red>Click the PayPal image to link to the PayPal site to submit your payment</Font>
	<br><br>
	
	<input type="hidden" name="payment_type" value="PP">

	<!-- test account -->
	
	<form action="https://www.sandbox.paypal.com/cgi-bin/webscr" method="post">	
	
	<!--
	<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
	-->

	<%
	
		points_used = 0 
		points_amt = 0 	 
		DoNot_Pass_CartContents = false
		set ORDER_RECORD = sitelink.ORDER_RECORD(session("ordernumber"),false)
			points_used = ORDER_RECORD.points_usd
			points_amt = ORDER_RECORD.checkamoun
			ord_total  = ORDER_RECORD.ord_total			  
			'remaining_balance = ord_total-points_amt	
		set ORDER_RECORD=nothing
		
	   remaining_balance=sitelink.ORDER_BALANCE_REMAINING(session("ordernumber")) 
		
	   extrafield =""
		xmlstring = sitelink.getbasketinfo(session("shopperid"),session("ordernumber"),extrafield,false)		
		objDoc.loadxml(xmlstring)
		
		set SL_Basket = objDoc.selectNodes("//gbi")
		
		tnode= "//gbi[it_unlist < 0]"
		
		DoNot_Pass_CartContents = false
		set SL_Basket_NegativePrice = objDoc.selectNodes(tnode)
		if SL_Basket_NegativePrice.length > 0 then
			DoNot_Pass_CartContents = true
			'NegativeAmount = SL_Basket_giftcert.item(0).selectSingleNode("it_unlist").text
			'remaining_balance = ord_total + NegativeAmount
		end if 	
		set SL_Basket_NegativePrice = nothing
		
		
			
		if (points_used > 0 and points_amt > 0) or basket_has_giftcert=true  then
		  	PayPal_Order_total = FORMATCURRENCY(remaining_balance)
			PayPal_Order_total =  mid(PayPal_Order_total,2,len(PayPal_Order_total))	 
			DoNot_Pass_CartContents = true 
	    else
	  		PayPal_Order_total = FORMATCURRENCY(sitelink.ORD_TOTAL(session("ordernumber"),false))
			PayPal_Order_total =  mid(PayPal_Order_total,2,len(PayPal_Order_total))
	    end if
		
		


	  shiptotal = FORMATCURRENCY(sitelink.ORD_SHIPPING(session("ordernumber"),false))
	  taxtotal  = FORMATCURRENCY(sitelink.ORD_TAX(session("ordernumber"),false)) 
	  paypalaccountid = sitelink.Get_PayPalLogin()	 
	  paypalcurrency  =sitelink.Get_PayPalCurrencyCode()
	  

	  shipping_for_Paypal = mid(shiptotal,2,len(shiptotal))
	  taxtotal_for_Paypal =  mid(taxtotal,2,len(taxtotal))
	  
	  successurl = secureurl + cstr("processorder.asp")
	  cancelurl = secureurl + cstr("checkout.asp")
	  notifyurl = secureurl + cstr("processorder.asp")
	  imageurl  = secureurl + "images/"+ cstr(COMPANY_LOGO_IMG)
	  first_name = session("firstname")
	  last_name=  session("lastname")
	  address1 =session("address1")
	   address2 =session("address2")
	   city = session("city")
	   state = session("state")
	   zip= session("zipcode")	
	
	%>
	<table>
				<tr>
                    <td valign="top"><input type="image" name="submit" border="0" src="images/horizontal_solution_PP.gif" alt="Save time. Checkout securely.Pay without sharing your financial information" align="middle"></td>
				<td><span style="font-size:11px; font-family: Arial, Verdana;">Save time.  Checkout securely.  
				<br>
				Pay without sharing your financial information.
				<br>
				<!--<b>You must have a valid Pay Pal Account.</b>-->
				</span></td></tr>
				
				</table>
				
	<% if DoNot_Pass_CartContents = false then %>
						
	<input type="hidden" name="cmd" value="_cart">			
	
	<%

	for x=0 to SL_Basket.length-1		 
		 SL_BasketDesc		= SL_Basket.item(x).selectSingleNode("inetsdesc").text
		 SL_BasketDesc      = replace(SL_BasketDesc,"""","")
		 SL_BasketUnitPrice	= formatcurrency(SL_Basket.item(x).selectSingleNode("extendp").text)
		 SL_BasketUnitPrice = mid(SL_BasketUnitPrice,2,len(SL_BasketUnitPrice))		 		 
		 SL_BasketQuantity	= 1
		 
		 actual_Quantity = SL_Basket.item(x).selectSingleNode("quanto").text
		 actual_Quantity = replace(actual_Quantity,".00","")
		 actual_Unitprice = formatcurrency(SL_Basket.item(x).selectSingleNode("it_unlist").text)
		 actual_Unitprice = mid(actual_Unitprice,2,len(actual_Unitprice))	
		 		 	 
		 SL_Number	= SL_Basket.item(x).selectSingleNode("item").text
		 SL_Basket_Shipto	= SL_Basket.item(x).selectSingleNode("ship_to").text
		 SL_Basket_Shipto = SL_Basket_Shipto + 1 - 1
	%>
	<input type="hidden" name="item_name_<%=x+1%>" value="<%=SL_BasketDesc%>">
	<input type="hidden" name="amount_<%=x+1%>" value="<%=SL_BasketUnitPrice%>">
	<input type="hidden" name="quantity_<%=x+1%>" value="<%=SL_BasketQuantity%>">
	<input type="hidden" name="item_number_<%=x+1%>" value="<%=SL_Number%>">
	<%
		option_val = actual_Quantity + "@" + actual_Unitprice
		if SL_Basket_Shipto > 0 then 
		 	fname = SL_Basket.item(x).selectSingleNode("firstname").text
		 	lname = SL_Basket.item(x).selectSingleNode("lastname").text
		 	fullname = (cstr(fname) + " " + cstr(lname))	
		 	fullname = "  Shipping to "+ cstr(fullname)
			option_val = option_val +  fullname
		end if
	%>
	
	<input type="hidden" name="on0_<%=x+1%>" value="<%=option_val%>">
	
	<%next%>
		
	<input type="hidden" name="shipping_1" value="<%=shipping_for_Paypal%>">
	<input type="hidden" name="tax_cart" value="<%=taxtotal_for_Paypal%>">
	<%else%>
		<input type="hidden" name="cmd" value="_xclick">
		<input type="hidden" name="amount" value="<%=PayPal_Order_total%>">
	<%end if %>
	
	<input type="hidden" name="invoice" value="<%=session("ordernumber")%>">
	<input type="hidden" name="item_name" value="Web Order: <%=session("ordernumber")%>">
	<input type="hidden" name="currency_code" value="<%=paypalcurrency%>">	
	<input type="hidden" name="return" value="<%=successurl%>">
	<input type="hidden" name="cancel_return" value="<%=cancelurl%>">
	<input type="hidden" name="notify_url" value="<%=notifyurl%>">
	<input type="hidden" name="image_url" value="<%=imageurl%>">
	<input type="hidden" name="first_name" value="<%=first_name%>">
	<input type="hidden" name="last_name" value="<%=last_name%>">
	<input type="hidden" name="address1" value="<%=address1%>">
	<input type="hidden" name="address2" value="<%=address2%>">
	<input type="hidden" name="city" value="<%=city%>">
	<input type="hidden" name="state" value="<%=state%>">
	<input type="hidden" name="zip" value="<%=zip%>">
	<input type="hidden" name="no_note" value="1">
	<input type="hidden" name="rm" value="2">
	<input type="hidden" name="mrb" value="R-1V748593L7668510W">
	<input type="hidden" name="pal" value="R-NCY5RMVCQZZDG">
	<input type="hidden" name="bn" value="Dydacomp">
	<input type="hidden" name="business" value="<%=paypalaccountid%>">
	<input type="hidden" name="upload" value="1">
	<input type="hidden" name="cbt" value="Press this button to complete your order">
	
	</form>
	</td></tr>
	</table>
	
	
	<%
	set SL_Basket = nothing
	end if%>
			
	<br>

	</center>
	</div>
<!-- end sl_code here -->
	<img alt="" src="images/clear.gif" width="1" height="160" border="0">
	</td>
	
</tr>

</table>



    </div> <!-- Closes main  -->
    <div id="footer" class="footerbgcolor">
    <!--#INCLUDE FILE = "include/bottomlinks.asp" -->
    <!--#INCLUDE FILE = "googletracking.asp" -->
    <!--#INCLUDE FILE = "RemoveXmlObject.asp" -->
    <!--#INCLUDE FILE = "text/footer.asp" -->
    </div>
</div> <!-- Closes container  -->


</body>
</html>
