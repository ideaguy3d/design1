
<table border=1>
<% 
count = 1
for each key in Session.Contents
  %><tr><%
  %><td><%=count %> --> <b style="COLOR: black; BACKGROUND-COLOR: #ffff66">Session</b>("<%=lcase (key)%>")</td><td><%
  if IsObject(Session.Contents(key)) then
    if lcase(TypeName(Session.Contents(key))) = "dictionary" then
      dumpDictionary(Session.Contents(key))
    else
      Response.Write TypeName(Session.Contents(key))
    end if
  else
    if IsArray(Session.Contents(key)) then
      %><ul><%
      for each n in Session.Contents(key)
        Response.Write ("<li>" & Session.Contents(key)(n))
        count = count + 1
      next
      %></ul><%
    else
      Response.Write ("<pre>" & Session.Contents(key) & "</pre>")
    end if
  end if
  %></td><%
  %></tr><%
count = count + 1

next


response.Write("<br>")
response.Write("<br>")
response.Write(" count " & count-1)

%>
</table>