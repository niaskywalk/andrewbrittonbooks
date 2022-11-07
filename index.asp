<%
Response.Buffer = True
%>
<HTML>
<BODY>
<!--Replace http://scripts.adcgroup.com with the URL that you want to redirect to-->
<%
If 1 = 1 Then
   Response.Clear
   Response.Redirect "content/index.asp"
End If
%>
<%
Response.End
%>
</BODY>
</HTML>