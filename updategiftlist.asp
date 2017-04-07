<%on error resume next%>

<!--#INCLUDE FILE = "include/momapp.asp" -->
<%

	y=request("qty").count
	
		for z=y  to 1 step-1
			remove_flag = request("remove"&cstr(z)).count
			qty = request("qty")(z)
			qty = qty + 1 -1
			itemnum = cstr(request("item")(z))
			extra = ""
			if remove_flag= 1 or qty= 0 then
		 		call sitelink.GIFTLIST("DELETE",cint(session("giftregcustnum")),itemnum,qty,0,"", cstr(extra))
			else
				call sitelink.GIFTLIST("UPDATE",cint(session("giftregcustnum")),itemnum,qty,0,"", cstr(extra))
			end if
		next
		 set sitelink=nothing
		 response.redirect("giftlist.asp")
	
%>

