<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="Config.asp"-->
<%
CheckLogin
Dim Msg
	If Request.QueryString("Action") = "Save" Then SaveData
	Sub SaveData()
	myConn.execute("update Config set OKAr='"&Request.Form("ftype")&"',OKsize="&Request.Form("fsize"))
	Msg = "成功修改了文件數據信息"
	End Sub

If msg <> "" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_Main.asp'>"&Msg&"<br>本頁將在3秒內返回<BR>如果你的瀏覽器沒有反應，請<a href=Admin_Main.asp>點擊此處返回</a>")
Response.End()
End If
%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<title>:::丁書記上傳程序 ::: 演示版本</title>
<link href="upstyle.css" rel="stylesheet" type="text/css">
<script src="Images/up.Js"></script>
</head>

<body background="../images/mainbg.gif" leftmargin="0" topmargin="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="75">&nbsp;</td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#FF9900" class="text">當前位置：:::丁書記上傳程序Ver1.2 ::: 演示版本 
      &gt;&gt;&gt; 後台管理</td>
  </tr>
  <tr>
    <td height="2" bgcolor="#cccccc"></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="1" bgcolor="#CCCCCC"></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
<form name="Edit" action="Admin_Main.asp?Action=Save" method="post">
  <tr> 
    <td height="25"> <%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from Config"
	frst.open sql,myconn,1,1
	If not frst.Eof then
	%>
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
          <tr align="center" class="text"> 
            <td height="1" colspan="2">後台管理-系統設置&nbsp;&nbsp;[<a href="Admin_List.asp">點此進入文件管理</a>]&nbsp;&nbsp;[<a href="Index.asp">退出管理</a>]</td>
          </tr>
          <tr class="text"> 
            <td width="200"> <div align="right">允許上傳的文件類型：</div></td>
            <td><input name="ftype" type="text" class="TextBoxT" value="<%=rs(0)%>" size="50">
              以,分隔後綴名,切記勿允許上傳asp/exe文件</td>
          </tr>
          <tr class="text"> 
            <td width="200"><div align="right">允許上傳的文件大小：</div></td>
            <td><input name="fsize" type="text" class="TextBoxT" value="<%=rs(1)%>" size="15">
              單位:Byte</td>
          </tr>
          <tr class="text"> 
            <td height="1" align="right" bgcolor="#0153A3">&nbsp;</td>
            <td height="1" bgcolor="#0153A3"> <input type="submit" name="Submit" value="修改"></td>
          </tr>
        </table>
      <%
	  else
	  %> 
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
          <td>沒有對應的數據！</td>
        </tr>
      </table>
      <%
	  end if
	  frst.close
	  set frst = nothing
	  myconn.close
	  set myconn = nothing
	  %>
	  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
            <td align="right" class="heading">&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table></td>
  </tr></form>
</table>

<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
  <tr> 
    <td height="25">&nbsp;</td>
  </tr>
</table>
</body>
</html>