<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
function checked()
{	
	f = basketform.j.length?basketform.j.length:0 ;
	alert(f);
	for(t=0; t < f; t++)
	{
		j = basketform.x[t].length?basketform.x[t].length:0 ;
		alert(j);
		for (k=0; k<j; k++)
		{
			radiochecked = basketform.x[t][k].checked ;
			if (radiochecked)
			 {
				alert(k);
			 }		
	
		}
	}
}
</script>
</head>

<body>
<form method="post" id="basketform" name="basketform" action="posttestjc.asp">
	<%for x=0 to 5%>
		<%for y=0 to 3%>
			<input type="hidden" name="x[<%=y%>]" id="x[<%=y%>]" value="<%=y%>">
			<br> <img src="images/clear.gif" width="<%=3*y%>">
			<input type="radio" name="choice" id="choice" value="<%=y%>"><br>
		<%next%>
		<input type="hidden" name="j" id="j" value="<%=x%>">
	<%next%>
	<input type="submit" value="Submit">
</form>
	<a href="javascript:checked();"><img src="images/btn-buynow.gif"></a>
</body>
</html>
