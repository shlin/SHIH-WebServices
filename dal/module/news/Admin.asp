<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="Config.asp"-->
<%
Dim Msg,Uname,Upass
Uname = replace(trim(Request.Form("Uname")),"'","")
Upass = replace(trim(Request.Form("Upass")),"'","")

	If Request.QueryString("Action") = "Check" Then Check Uname,Upass
	Sub Check(Uname,Upass)
	Set rs = myConn.execute("Select Adminname,AdminPass From Config")
	If rs.Eof Then
	Msg = "系統數據庫錯誤，請仔細檢查數據庫結構及數據記錄"
	Exit Sub
	End If
	If Uname <> rs(0) or Upass <> rs(1) Then
	Msg = "帳戶或密碼錯誤"
	Exit Sub
	End If
	Session("Admin") = "OK"
	Msg = "身份驗證通過"
	End Sub

If msg <> "" then
If Session("Admin") = "OK" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_List.asp'>"&Msg&"<br>本頁將在3秒內返回<BR>如果你的瀏覽器沒有反應，請<a href=Admin_List.asp>點擊此處返回</a>")
Response.End()
Else
Response.Write("<meta http-equiv=refresh content='3;URL=Admin.asp'>"&Msg&"<br>本頁將在3秒內返回<BR>如果你的瀏覽器沒有反應，請<a href=Admin.asp>點擊此處返回</a>")
Response.End()
End If
End If
%>


<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="keywords" content="丁書記上傳程式 search engine marketing ">
<title>新聞管理模組</title>
</head>
<LINK href="master.css" type=text/css rel=stylesheet>
<body topmargin="0" leftmargin="0">

<table border="0" cellpadding="0" style="border-collapse: collapse" width="650" id="table1" height="500">
	<tr>
		<td valign="top">
		<table border="0" cellpadding="0" style="border-collapse: collapse" width="760" id="table2">
			<tr>
				<td width="449" height="30" bgcolor="#EEEEEE">　上傳程式</td>
				<td width="311" bgcolor="#EEEEEE">
　|　<a href="?page=1">首頁</a>
　|　<a href="fileupload.asp">上傳</a>
　|　<a href="admin.asp">管理</a>
　|　

				</td>
			</tr>
			<tr>
				<td width="760" colspan="2">　　　　　管理<hr color="#C0C0C0" size="1" width="90%"></td>
			</tr>
			<tr>
				<td width="760" colspan="2" height="350" valign="top">　
				<div align="center" class=pic>
<form name="Check" action="Admin.asp?Action=Check" method="post">
<table border="1" cellpadding="0" style="border-collapse: collapse" width="80%" id="table3" bordercolor="#EBEBEB">
						<tr>
							<td>管理員帳戶：<input name="uname" type="text" class="TextBoxT" id="uname" size="15"></td>
						</tr>
						<tr>
							<td valign="top">管理員密碼：<input name="upass" type="password" class="TextBoxT" id="upass" size="15"></td>
						</tr>
						<tr>
							<td valign="top">
							<p align="center"><input type="submit" name="Submit" value="登入"></td>
						</tr>
						<tr>
							<td bgcolor="#328CDD">
							<img border="0" src="jpg" height="1"></td>
						</tr>
</table>


				</td>
			</tr>
			<tr>
				<td width="760" colspan="2">
			
			</td>
			</tr>
			<tr>
				<td width="760" colspan="2" bgcolor="#877460">
							<img border="0" src="jpg" height="5"></td>
			</tr>
			<tr>
				<td width="760" colspan="2" height="13"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>

</body>

</html>