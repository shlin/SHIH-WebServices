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
	Msg = "�t�μƾڮw���~�A�ХJ���ˬd�ƾڮw���c�μƾڰO��"
	Exit Sub
	End If
	If Uname <> rs(0) or Upass <> rs(1) Then
	Msg = "�b��αK�X���~"
	Exit Sub
	End If
	Session("Admin") = "OK"
	Msg = "�������ҳq�L"
	End Sub

If msg <> "" then
If Session("Admin") = "OK" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_List.asp'>"&Msg&"<br>�����N�b3����^<BR>�p�G�A���s�����S�������A��<a href=Admin_List.asp>�I�����B��^</a>")
Response.End()
Else
Response.Write("<meta http-equiv=refresh content='3;URL=Admin.asp'>"&Msg&"<br>�����N�b3����^<BR>�p�G�A���s�����S�������A��<a href=Admin.asp>�I�����B��^</a>")
Response.End()
End If
End If
%>


<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="keywords" content="�B�ѰO�W�ǵ{�� search engine marketing ">
<title>�s�D�޲z�Ҳ�</title>
</head>
<LINK href="master.css" type=text/css rel=stylesheet>
<body topmargin="0" leftmargin="0">

<table border="0" cellpadding="0" style="border-collapse: collapse" width="650" id="table1" height="500">
	<tr>
		<td valign="top">
		<table border="0" cellpadding="0" style="border-collapse: collapse" width="760" id="table2">
			<tr>
				<td width="449" height="30" bgcolor="#EEEEEE">�@�W�ǵ{��</td>
				<td width="311" bgcolor="#EEEEEE">
�@|�@<a href="?page=1">����</a>
�@|�@<a href="fileupload.asp">�W��</a>
�@|�@<a href="admin.asp">�޲z</a>
�@|�@

				</td>
			</tr>
			<tr>
				<td width="760" colspan="2">�@�@�@�@�@�޲z<hr color="#C0C0C0" size="1" width="90%"></td>
			</tr>
			<tr>
				<td width="760" colspan="2" height="350" valign="top">�@
				<div align="center" class=pic>
<form name="Check" action="Admin.asp?Action=Check" method="post">
<table border="1" cellpadding="0" style="border-collapse: collapse" width="80%" id="table3" bordercolor="#EBEBEB">
						<tr>
							<td>�޲z���b��G<input name="uname" type="text" class="TextBoxT" id="uname" size="15"></td>
						</tr>
						<tr>
							<td valign="top">�޲z���K�X�G<input name="upass" type="password" class="TextBoxT" id="upass" size="15"></td>
						</tr>
						<tr>
							<td valign="top">
							<p align="center"><input type="submit" name="Submit" value="�n�J"></td>
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