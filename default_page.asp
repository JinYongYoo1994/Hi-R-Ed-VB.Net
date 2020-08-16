<% @language="vbscript" %>
<html>
<head><title>Home Page</title></head>
<body>
<h3>Home Page</h3>
<p>You are logged on as: 
<%
  If Len(Session("UID")) = 0 Then
    Response.Write "<b>You are not logged on.</b>"
  Else
    Response.Write "<b>" & Session("UID") & "</b>"
  End If
%>
</p>
<ul>
 
<li><a href="passwordprotect.asp">Password-Protected Page</a></li>
<li><a href="nonsecure.asp">Nonsecure Page</a></li>

</ul>
</body>
</html>