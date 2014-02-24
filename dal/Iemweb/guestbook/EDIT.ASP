<%
' 與資料庫連線
  Set conn = Server.CreateObject("ADODB.Connection")
  DBPath = Server.MapPath("GBook.mdb")
  conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open "select * from guestbook where messageid=" & Request.Querystring("MESSAGEID") , conn,1,3
%>
<html>
<head><title>管理模式--修改</title></head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
 <STYLE TYPE="text/css">
        <!-- body,th,td,input,select,textarea,select,checkbox {font:9pt 新細明體}
         a:link {font: 9pt "新細明"; text-decoration: none;color:none}
         a:visited      {font: 9pt "新細明"; text-decoration: none; color: 000099}
         a:active       {font: 9pt "新細明"; text-decoration: none; color: 00CC66}
         a:hover        {font: 9pt 新細明; text-decoration: underline; color: ff0000}
        -->
  </STYLE>
<body background="石塊地板.gif">
<%
   ' 讀取資料
   messageid = Request.Querystring("messageid")
   posttime = rs("posttime")
   Name = rs("Name")
'   Sex = rs("Sex")
'   Job = rs("Job")
   Email = rs("Email")
   HomePage = rs("HomePage")
   messagebody = rs("messagebody")
'   icq = rs("icq")

   rs.close
   SET CONN = NOTHING

  ' 將兩個全形的字串改成空字串
  MESSAGEbody = Replace(MESSAGEbody,"　　","")

  ' 將 <br> 改成一個換行符號
  MESSAGEbody = Replace(MESSAGEbody,"<br>",vbCrLf)
%>
 <form method="POST" action="SaveEdit.asp" name="edit">
  <div align="center"><center><p><font  face="Arial"><big><strong>Marge Story Guest Book v2.0<BR><BR><FONT COLOR=RED>修改留言者資料</FONT></strong></big></font></p>
  </center></div><div align="center"><center>
<table border="0">
    <tr>
      <td bgcolor="#CCCCFF"><font color="blue">留言編號</font></td>
      <td bgcolor="#CCCCFF"><%=messageid%></td>
    </tr>
    <tr>
      <td bgcolor="#CCCCFF"><font color="blue">留言日期</font></td>
      <td bgcolor="#CCCCFF"><%=posttime%></td>
    </tr>
    <tr>
      <td bgcolor="#CCCCFF"><font color="blue">姓　　名</font></td>
      <td bgcolor="#CCCCFF"><input type="text" name="NAME" size="20"  value="<%=name%>"></td>
    </tr>
    <tr>
      <td bgcolor="#CCCCFF"><font color="blue" >電子郵件</font></td>
      <td bgcolor="#CCCCFF"><input type="text" name="EMAIL" size="40"  value="<%=email%>"></td>
    </tr>
    <tr>
      <td bgcolor="#CCCCFF"><font color="blue">網頁位址</font></td>
      <td bgcolor="#CCCCFF"><input type="text" name="URL" size="40" value="<%=homepage%>"></td>
    </tr>
    <tr>
      <td bgcolor="#CCCCFF"><font color="blue" >留言內容</font></td>
      <td bgcolor="#CCCCFF"><textarea rows="5" name="MESSAGEbody" cols="60"><%=Messagebody%></textarea></td>
    </tr>
    <tr>
      <td colspan="2" align="center">
        <input type="hidden" name="messageid" value="<%=Request.querystring("messageid")%>">
        <input type="hidden" name="replylevel" value="<%=Request.querystring("replylevel")%>">
        <input type="submit" value="修改完成">
        <input type="button" value="返回留言板" onclick="location.href='book.asp'">
      </td>
    </tr>
  </table>
  </center>
</form>
</body>
</html>
<script language="vbscript">
<!-- 
   document.edit.elements(3).focus()
-->
</script>