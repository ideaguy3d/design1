
<img src="vspacer.gif" width="5" height="8" /><script>
  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
  })(window,document,'script','//www.google-analytics.com/analytics.js','ga');

  ga('create', 'UA-45483712-1', 'aussieproducts.com');
  ga('send', 'pageview');

</script>
<div id="bottomlinks">
   <img src="spacer.gif" width="10" height="1" /> <%if DISP_ABOUT_PAGE = true then %>
         <a class="footerlink" href="<%=insecureurl%>Aboutus.asp" title="<%=ALT_ABOUT_TXT%>"><%=ALT_ABOUT_TXT%></a> <img src="spacer.gif" width="12" height="1" />
        <% end if %>
  <%if DISP_TERMS_PAGE = true then %>
  <a class="footerlink" href="<%=insecureurl%>sizechart.asp" title="<%=ALT_TERMS_TXT%>"><%=ALT_TERMS_TXT%></a> <img src="spacer.gif" width="12" height="1" />
         <% end if %>
        <%if DISP_PRIVACY_PAGE = true then %>
  <a class="footerlink" href="<%=insecureurl%>privacy.asp" title="<%=ALT_PRIVACY_TXT%>"><%=ALT_PRIVACY_TXT%></a> <img src="spacer.gif" width="12" height="1" />
        <% end if %>
        <%if DISP_TESTIMONIAL_PAGE = true then %>
  <a class="footerlink" href="<%=insecureurl%>Testimonials.asp" title="<%=ALT_TESTIMONIAL_TXT%>"><%=ALT_TESTIMONIAL_TXT%></a><img src="spacer.gif" width="12" height="1" />
<% end if %>
         <a class="footerlink" href="<%=insecureurl%>sitemap.asp" title="Site Map"> Site Map </a> <img src="spacer.gif" width="12" height="1" /> <a class="footerlink">  
        <% if session("registeredshopper")="NO" then %>
        <img src="spacer.gif" width="8" height="1" /><a href="<%=secureurl%>login.asp" title="Login" class="footerlink">  Login / Register  </a><%else%><a href="<%=secureurl%>logout.asp" title="Logout" class="footerlink">  Logout  </a><img src="spacer.gif" width="8" height="1" />
        <%end if%> 
         <img src="spacer.gif" width="8" height="1" /><a class="footerlink" href="<%=insecureurl%>basket.asp" title="View Cart">  View Cart  </a><img src="spacer.gif" width="8" height="1" />
<%if DISP_HELP_PAGE = true or DISP_RETURN_PAGE = true or DISP_CONTACT_PAGE = true then %>
      
        <%if DISP_HELP_PAGE = true then %>
       <img src="spacer.gif" width="12" height="1" /><a class="footerlink" href="<%=insecureurl%>shipping.asp" title="<%=ALT_HELP_TXT%>"><%=ALT_HELP_TXT%></a> <img src="spacer.gif" width="20" height="1" />
       <% end if %>
        <%if DISP_RETURN_PAGE = true then %>
         <img src="spacer.gif" width="12" height="1" /><a class="footerlink" href="<%=insecureurl%>meet_the_staff.asp" title="<%=ALT_RETURN_TXT%>"><%=ALT_RETURN_TXT%></a><img src="spacer.gif" width="20" height="1" />
<% end if %>
 <%if DISP_CONTACT_PAGE = true then %>
         <img src="spacer.gif" width="12" height="1" /><a class="footerlink" href="<%=insecureurl%>contactus.asp" title="<%=ALT_CONTACT_TXT%>"><%=ALT_CONTACT_TXT%></a> 
        <% end if %>
    <% end if %> 
</div> 


