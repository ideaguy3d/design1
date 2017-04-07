<%if len(trim(GA_ACCTNUM)) > 0 then%> 
	<script type="text/javascript">			
	    var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");		
	    document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));	
	</script>
	<script type="text/javascript">	
	    var pageTracker = _gat._getTracker("<%=GA_ACCTNUM%>");	
	    pageTracker._initData();	
	    pageTracker._trackPageview();	
	</script>
<%end if%>

<%if SHOW_BUYSAFE =1 then %>
<!--#INCLUDE FILE = "text/BuySafeSeal.html" -->
<%end if %>