<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Buffer = True %>


<!-- 取得資料　!-->
<% STATUS = Request.FORM("STATUS") %>
<% PASSWORD = Request.FORM("PASSWORD") %>

<!-- #INCLUDE FILE="password.inc" -->

         <%If PASSWORD = adminpass THEN%>
                   <%Session("password") = True%>
                   <%RESPONSE.REDIRECT "book.asp"%>
         <%else%>
                 <%If PASSWORD <> adminpass and password <> "" THEN%>
                      <%=error%>
                 <%End If%>
         <%End If%>

<!----------輸入密碼------------------!-->
<HTML><HEAD><TITLE><%=TITLE2%></TITLE></HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<BODY background="石塊地板.gif">
 <%'=err%>
<form method="POST" action="admin.asp" name="admin">
  <div align="center"><font color=blue><b><%=TITLE%></font></b><br><br>管理者密碼<center><input type="password" name="PASSWORD" size="10"><br><br>
  <input type="hidden" value="CHECKEM" Name="STATUS">
  <input type="submit" value="確　定"></p>
  </center></div>
</form>

<%=RETU%>
<!-- #INCLUDE FILE="footer.inc" -->
</p>

</BODY>
</HTML>
<script language="vbscript">
<!-- 
   document.admin.elements(0).focus()
-->
</script>