<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<!--INCLUDE FILE = "CreateXmlObject.asp" -->



<%

	'if session("NoShowEmail") = 1 then
		KeepEmailPrivate = true
	'else
	'	KeepEmailPrivate = false
	'end if
	active=false
	
	
	

    
	set SL_cProdReviewObj = New cProdReviewObj
	
	SL_cProdReviewObj.laction="ADDTO_REVIEW"
	SL_cProdReviewObj.llitemnumber=cstr(session("item_number"))
	SL_cProdReviewObj.llrating=session("rating_val")
    SL_cProdReviewObj.lltitle=session("ReviewTitle")
    SL_cProdReviewObj.lldesc=session("ReviewComment")
    SL_cProdReviewObj.llemail=session("ReviewerEmail")
    SL_cProdReviewObj.llprivate=KeepEmailPrivate
    SL_cProdReviewObj.llactivate=active
    SL_cProdReviewObj.llfirstname=cstr(session("firstname"))
    SL_cProdReviewObj.lllastname=cstr(session("lastname"))
    SL_cProdReviewObj.llstate=cstr(session("state"))
    

    
	xmlstring= sitelink.Customer_Review_Deergear(SL_cProdReviewObj)
	
	set SL_cProdReviewObj=nothing
	
	'call sitelink.Customer_Review_Deergear("ADDTO_REVIEW",cstr(session("item_number")),session("rating_val"),session("ReviewTitle"), session("ReviewComment"),session("ReviewerEmail"),KeepEmailPrivate,active,0,false,cstr(session("firstname")),cstr(session("lastname")),cstr(session("state")),false)	

	session("rating_val") = ""
	session("ReviewTitle") = ""
	session("ReviewComment") = ""
	session("ReviewerEmail") = ""
	session("NoShowEmail") = ""
	session("firstname") = ""
	session("lastname") = ""
	session("state") = ""

	set sitelink = nothing
		
	gotothispage = insecureurl & "prodinfo.asp?number="+session("item_number")
	response.redirect(gotothispage)

%>