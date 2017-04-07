
<%on error resume next%>


<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--#INCLUDE FILE = "CreateXmlObject.asp" -->

<%
Response.ExpiresAbsolute = Now() - 1 
Response.AddHeader "Cache-Control", "must-revalidate" 
Response.AddHeader "Cache-Control", "no-store" 
%>   


<%

' write important session variables to file		
StrFileName = "text\sess_txt.txt"
StrPhysicalPath = Server.MapPath(StrFileName)
Set fs=Server.CreateObject("Scripting.FileSystemObject")
set ts = fs.OpenTextFile(StrPhysicalPath , 8)
ts.write(session("ordernumber") & "," & session("shopperid") & "," & session("sfirstname") & "," & session("slastname") & ",")
ts.write(session("scompany") & "," & session("saddress1") & "," & session("saddress2") & "," & session("scity") & "," & "," )
ts.write(session("sstate") & session("szipcode") & "," & session("scountry") & "," & session("sphone") & "," & session("semail") & ",")
ts.write(session("firstname") & "," & session("lastname") & "," & session("company") & "," & session("address1") & "," & session("address2") & ",")
ts.write(session("city") & "," & session("state") & "," & session("zipcode") & "," & session("country") & "," & session("phone") & "," & session("email") & ",")
ts.write(session("billtocopy"))
ts.writeline()
ts.close
set ts=nothing
set fs=nothing

'REM CLEAR THE cookie for this order
cookiename=ShortStoreName+"order"
RESPONSE.COOKIES(cookiename).Expires="January 1, 1997"

	
	'NOW NEW ORDER
	session("previousorder") = session("ordernumber")
	session("previousbilltocopy") = session("billtocopy")
	session("previoususer_want_multiship") = session("user_want_multiship") 
	
	'now fire neworder
	session("shopperid")=sitelink.initshopper("","") 
	'session("clshopper")="S"+cstr(session("shopperid"))
	session("ordernumber")=sitelink.initorder("") 
	call sitelink.createcustomer(session("shopperid"),"")
	'session("needtocheckorderinfo")="YES"
	SAVEDSHOPPERID=""
	session("returnvisitor")=1
	
	session("cardwasauthed")=false
	
	if session("registeredshopper")="YES" then	
		'session("regemail")= regemail
		'session("password")=regpassword
		'response.write("New shopper -->"&session("shopperid"))	
		set SLShopper = New cShopperInfo
			SLShopper.shopperid=session("shopperid")
			SLShopper.bemail=session("regemail")
			SLShopper.bpassword=session("password")	
			SAVEDSHOPPERID=sitelink.verifyshopper3(SLShopper)	
			
			if SAVEDSHOPPERID<>"-1" then
				session("firstname")=SLShopper.bfirstname
				session("lastname")=SLShopper.blastname
				session("company")=SLShopper.bcompany
				session("address1")=SLShopper.baddr1
				session("address2")=SLShopper.baddr2
				session("address3")=SLShopper.baddr3
				session("city")=SLShopper.bcity
				session("state")=SLShopper.bstate
				session("zipcode")=SLShopper.bzipcode
				session("phone")=SLShopper.bphone
				session("email")=SLShopper.bemail
				session("country")=SLShopper.bcountry
				session("pointavail")=SLShopper.bpoints
				if SLShopper.bNopoints=true then
					session("pointavail")=0
				end if

			end if
		set SLShopper=nothing
	end if 
	set sitelink= nothing
	
	
	
	'clear this information
	session("SL_BasketCount") = 0
	session("SL_BasketSubTotal") = 0
	
	Session.Contents.Remove("showdept")
	Session.Contents.Remove("department")
	Session.Contents.Remove("topdept")
	Session.Contents.Remove("items_spaned")
	Session.Contents.Remove("shippingmethod")
	
	Session.Contents.Remove("sfirstname")
	Session.Contents.Remove("slastname")
	Session.Contents.Remove("scompany")
	Session.Contents.Remove("saddress1")
	Session.Contents.Remove("saddress2")
	Session.Contents.Remove("saddress3")	
	Session.Contents.Remove("scity")
	Session.Contents.Remove("sstate")
	Session.Contents.Remove("szipcode")
	Session.Contents.Remove("scountry")
	Session.Contents.Remove("sphone")
	Session.Contents.Remove("sphone2")
	Session.Contents.Remove("semail")
	
	Session.Contents.Remove("custerrormessage")
	Session.Contents.Remove("ordererrormessage")
	Session.Contents.Remove("Title_String")
	Session.Contents.Remove("prospecterror")
	Session.Contents.Remove("catalog")
	Session.Contents.Remove("billtocopy")
	Session.Contents.Remove("receiveemail")
	Session.Contents.Remove("cardwasauthed")

	
	Session.Contents.Remove("paymentmethod")
	Session.Contents.Remove("cc_name")
	Session.Contents.Remove("cc_number")
	Session.Contents.Remove("cc_id")
	Session.Contents.Remove("cc_type")
	Session.Contents.Remove("cc_expmonth")
	Session.Contents.Remove("cc_expyear")
	Session.Contents.Remove("accttype")
	Session.Contents.Remove("bankname")
	Session.Contents.Remove("bankrountingnum")
	Session.Contents.Remove("bankacctnum")	
	Session.Contents.Remove("PayPalStr")
	Session.Contents.Remove("PayPalBuyerID")
	Session.Contents.Remove("ordermemo")
	
	Session.Contents.Remove("from_date")
	Session.Contents.Remove("issue_num")

	Session.Contents.Remove("giftmsg1")
	Session.Contents.Remove("giftmsg2")
	Session.Contents.Remove("giftmsg3")
	Session.Contents.Remove("giftmsg4")
	Session.Contents.Remove("giftmsg5")
	Session.Contents.Remove("giftmsg6")

	
	Session.Contents.Remove("memo_field1")
	Session.Contents.Remove("memo_field2")
	Session.Contents.Remove("memo_field3")
	Session.Contents.Remove("ordermemo")
	
	Session.Contents.Remove("orderholdDate")
	
	Session.Contents.Remove("ponum")	
	Session.Contents.Remove("CurrentSourceKey")
	
	Session.Contents.Remove("Points_need")
	Session.Contents.Remove("Redeem_Amount")
	Session.Contents.Remove("DoNotShow_PointsTab")
	Session.Contents.Remove("PaybyPoints")
	Session.Contents.Remove("Remaining_balance")
	
	Session.Contents.Remove("DoNotShow_GCTab")
	Session.Contents.Remove("PaybyGC")
	Session.Contents.Remove("Remaining_BalanceAfterGC")
	
	
	
	session("cookiewritten") = false 
	session("user_want_multiship") = false
	
	session("askedforsourcekey")=-1
	session("sourcekeyisgood")=-1
	

set sitelink=nothing 


'default_page = insecureurl 

Response.Redirect(insecureurl)
	
%>