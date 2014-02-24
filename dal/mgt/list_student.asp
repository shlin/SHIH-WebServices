<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="../Connections/Dal.asp" -->
<!--#include file="fckeditor.asp" -->
<%
Dim studentDB
Dim studentDB_cmd
Dim studentDB_numRows

Set studentDB_cmd = Server.CreateObject ("ADODB.Command")
studentDB_cmd.ActiveConnection = MM_Dal_STRING
studentDB_cmd.CommandText = "SELECT * FROM dbo.Student ORDER BY [year] DESC " 
studentDB_cmd.Prepared = true

Set studentDB = studentDB_cmd.Execute
studentDB_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
studentDB_numRows = studentDB_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>新增</title>
<style type="text/css">
<!--
body {
	background-image: url(../img/body.gif);
}
-->
</style>
<link href="../dal.css" rel="stylesheet" type="text/css" />
</head>

<body>


<p align="center"><img src="../img/banner5.jpg" alt="Banner" width="758" height="115" /></p>
<table width="95%" border="0" cellpadding="0" cellspacing="0" class="border" >
  <tr>
    <td width="100"><div align="center"></div></td>
    <td width="50"><div align="center"><strong>年份</strong></div></td>
    <td width="80"><div align="center"><strong>級別</strong></div></td>
    <td width="80"><div align="center"><strong>類型</strong></div></td>
    <td><div align="center"><strong>主題</strong></div></td>
  </tr>    <% 
While ((Repeat1__numRows <> 0) AND (NOT studentDB.EOF)) 
%>
  <tr>

      <td width="100"><div align="center"><a href="del_student.asp?MemId=<%=(studentDB.Fields.Item("MemId").Value)%>">刪除</a>│<a href="mod_student.asp?memid=<%=(studentDB.Fields.Item("MemId").Value)%>">修改</a></div></td>
      <td width="50"><div align="center"><%=(studentDB.Fields.Item("year").Value)%></div></td>
      <td width="80"><div align="center"><%=(studentDB.Fields.Item("class").Value)%></div></td>
      <td width="80"><div align="center"><%=(studentDB.Fields.Item("type").Value)%>│</div></td>
      <td><%=(studentDB.Fields.Item("Topic").Value)%></td>
    </tr>
  <tr>
    <td colspan="5"><img src="../img/line.gif" alt="line" width="100%" height="8" /></td>
    </tr>      
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  studentDB.MoveNext()
Wend
%>
  <tr>
    <td width="100">&nbsp;</td>
    <td width="50">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
studentDB.Close()
Set studentDB = Nothing
%>
