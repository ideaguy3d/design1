<%on error resume next%>
<!-- Do not include verifyincl.asp -->
<!--#INCLUDE FILE = "include/AdminCreateSL.asp" -->


<%

 accountid = request.QueryString("Site")
 
 
 username = cstr(request.QueryString("uname"))
 
 
 'response.Write(request.QueryString)
 
  'username="Rakesh?1234?rakesh@dydacomp.com?https://sl6test.mailordercentral.com/test2/&site=28694037"
  
  MyArray = Split(username, "?", -1, 1)
  
  session("SLusername") = MyArray(0)
  session("SLuserpwd") =  MyArray(1)
  session("SLuseremail") =MyArray(2) 
  session("SLusersiterul") =MyArray(3)
  
  
  'lstr = "username -->"&username&" --- pwd -->"&lpwd &"---email -->"&lemail &"---url-->"&lurl
  'lstr =" pwd -->"&username & "account id -->"&accountid
'  
'                StrFileName = "text\xml.txt"
'		 		StrPhysicalPath = Server.MapPath(StrFileName)
'				 set objFileName = CreateObject("Scripting.FileSystemObject")		 
'				 set objFileNameTS = objFileName.CreateTextFile(StrPhysicalPath)
'				objFileNameTS.Writeline lstr
'			    objFileNameTS.Close
'				set objFileNameTS = nothing
'			    set objFileName = nothing
			    
  
'        StrFileName = "text\LivepersonAccount.asp"	  
'        StrPhysicalPath = Server.MapPath(StrFileName)
'	     set objFileName = CreateObject("Scripting.FileSystemObject")	
'	     set objFileNameTS = objFileName.CreateTextFile(StrPhysicalPath)
    		  			    
        'starting script tag
'	    objFileNameTS.Writeline "<"&"%"
    	
'	    vtext ="LIVE_PERSON_ACCTID =""""" 		 
'        objFileNameTS.Writeline vtext
        
'	    'Ending string tag
'	    objFileNameTS.Writeline "%"&">"
    		
'	    objFileNameTS.Close
'	    set objFileName = nothing
  
  
    if len(trim(accountid)) > 0 then
        'StrFileName = "text\LivepersonAccount.asp"	  
        'StrPhysicalPath = Server.MapPath(StrFileName)
	    ' set objFileName = CreateObject("Scripting.FileSystemObject")	
	    ' set objFileNameTS = objFileName.CreateTextFile(StrPhysicalPath)
    		  			    
        'starting script tag
	    'objFileNameTS.Writeline "<"&"%"
    	
	    'vtext ="LIVE_PERSON_ACCTID ="""& accountid &""""		 		 
        'objFileNameTS.Writeline vtext
        
	    'Ending string tag
	    'objFileNameTS.Writeline "%"&">"
    		
	    'objFileNameTS.Close
	    'set objFileName = nothing
    
        call sitelink.CreateLiveAccount("WRITE",session("SLusername"),session("SLuserpwd"),session("SLuseremail"),session("SLusersiterul"),cstr(accountid))        

        'now generate the script
        
        StrFileName = "text\Livepersonscript.asp"	  
        StrPhysicalPath = Server.MapPath(StrFileName)
	     set objFileName = CreateObject("Scripting.FileSystemObject")	
	     set objFileNameTS = objFileName.CreateTextFile(StrPhysicalPath)
	     
	     'starting script tag
	   
	     
	     vtext="<table border=""0"" cellspacing=""2"" cellpadding=""2"">"
	     objFileNameTS.Writeline vtext
	     
	     vtext="<tr>"
	     objFileNameTS.Writeline vtext
	     
	     vtext="<td align=""center"">"
	     objFileNameTS.Writeline vtext
	     	     
	     vtext ="<a id='_lpChatBtn' href='https://server.iad.liveperson.net/hc/"+cstr(accountid)+"/?cmd=file&file=visitorWantsToChat&site="+cstr(accountid)+ "&byhref=1&imageUrl="+ cstr(secureurl) +"images/" + "' target='chat"+cstr(accountid)+"'  " + "onclick=""lpButtonCTTUrl = 'https://server.iad.liveperson.net/hc/" + cstr(accountid)+"/?cmd=file&file=visitorWantsToChat&site="+cstr(accountid)+"&imageUrl="+cstr(secureurl)+"images/&referrer='+escape(document.location); lpButtonCTTUrl = (typeof(lpAppendVisitorCookies) != 'undefined' ? lpAppendVisitorCookies(lpButtonCTTUrl):lpButtonCTTUrl);"
         
         
         vtext =vtext + " window.open(lpButtonCTTUrl,'chat"+cstr(accountid)+"','width=472,height=320,resizable=yes');return false;"" >"
                 
         
         vtext = vtext +"<img src='https://www.liveperson.com/gallery/images/f1/reponline.gif?site="+cstr(accountid)+"&channel=web&&ver=1&imageUrl="+cstr(secureurl)+"images/' name='hcIcon' border=0></a>"
         
          objFileNameTS.Writeline vtext
          
          vtext = "</td></tr><tr><td align=""center"">"          
          objFileNameTS.Writeline vtext
          
          vtext = "<a href='http://www.liveperson.com' target=""_blank"" title=""Go to www.liveperson.com"" style='text-decoration:none'>"
          objFileNameTS.Writeline vtext
          
          vtext ="<span style='font-size: 10px; font-family: Arial, Helvetica, sans-serif; color:#000000'>Live Chat by <span style='color:red;'>LivePerson</span></span>"
          objFileNameTS.Writeline vtext
          
          vtext ="</a></td></tr></table>"
          objFileNameTS.Writeline vtext
          
        
    		
	    objFileNameTS.Close
	    set objFileName = nothing
	else
	    StrFileName = "text\Livepersonscript.asp"	  
        StrPhysicalPath = Server.MapPath(StrFileName)
	     set objFileName = CreateObject("Scripting.FileSystemObject")	
	     set objFileNameTS = objFileName.CreateTextFile(StrPhysicalPath)
	     
	     'starting script tag
	    'objFileNameTS.Writeline "<"&"%"
	     
	     vtext="<table border=""0"" cellspacing=""2"" cellpadding=""2"">"
	     objFileNameTS.Writeline vtext
	     
	     vtext ="<tr><td></td></tr>"
         objFileNameTS.Writeline vtext
	     
	     vtext ="</table>"
	     objFileNameTS.Writeline vtext
	     
	     'Ending string tag
	    'objFileNameTS.Writeline "%"&">"
    		
	    objFileNameTS.Close
	    set objFileName = nothing     
	end if
	
	
set sitelink=nothing

'response.Redirect("addliveperson.asp")




 %>
