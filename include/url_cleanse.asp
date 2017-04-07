
<% 


	function url_cleanse(thatvar)
		'remove spaces and slashes
	
		thatvar = trim(thatvar)
		intLen = len(thatvar)
		if intLen > 0 then
		    newstrText=""
    		
		    For intCounter = 1 To intLen
                strChar = Mid(thatvar, intCounter, 1)
                ascval = ASC(strChar)
                
                strCharOK=""
                'A-Z asc65-90
                if ascval >64 and ascval <91 then
                    strCharOK=strChar
                end if

                'a-z asc 97-122
                if ascval >96 and ascval <123 then
                    strCharOK=strChar
                end if
                
                '0-9 asc 48-57
                if ascval >47 and ascval <58 then
                    strCharOK=strChar
                end if
                
                'space asc 32     -
                if ascval =32 or ascval=45 then
                    strCharOK="-"
                end if
                ' / asc val 47
                if ascval =47 then
                    strCharOK="_"
                end if
                newstrText= newstrText & strCharOK
            next
            
            newstrText = replace(newstrText,"--","-")
            newstrText = replace(newstrText,"--","-")	
        else
            newstrText = "PRODUCT-TITLE"
        end if 	
        
		url_cleanse = newstrText	
	end function

	
%>
