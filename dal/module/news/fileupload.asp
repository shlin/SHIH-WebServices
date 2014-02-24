
<!--#include file="Config.asp"--><head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>上傳程式</title>
</head>


<link href="master.css" type="text/css" rel="stylesheet">
<html>
<body>

<table border="0" cellpadding="0" style="border-collapse: collapse" width="650" id="table1" height="500">
	<tr>
		<td valign="top">
		<table border="0" cellpadding="0" style="border-collapse: collapse" width="760" id="table2">
			<tr>
				<td width="1071" colspan="2">　　　　　上傳<hr color="#C0C0C0" size="1" width="90%"></td>
			</tr>
			<tr>
				<td width="760" colspan="2" height="350" valign="top">　
				
				<div align="center" class="pic">
<form action="SaveUpload.asp" method="post" enctype="multipart/form-data" name="form1">
<table border="1" cellpadding="0" style="border-collapse: collapse" width="80%" id="table3" bordercolor="#EBEBEB">
						<tr>
							<td width="400">
								新聞時間：<input name="NewsTime" class="TextBoxT" id="NewsTime">ex:2007/06/06 <br>
							    新聞標題：<input name="filetitle" type="text" class="TextBoxT" size="27"> <br>
				                作　　者：<input name="filetype" class="TextBoxT" id="filetype"><br>
								新聞來源：<input name="FILEFrom" class="TextBoxT" id="FILEFrom">
										
            </td>
							<td rowspan="3" valign="top">摘要：<br>
							<textarea name="filedesc" cols="25" rows="4" class="TextBoxT" id="filedesc"></textarea></td>
						</tr>
						<tr>
							<td width="400">上傳文件：<input name="filedata" type="file" class="TextBoxT" id="filedata" size="31"></td>
						</tr>
						<tr>
							<td width="400">允許文件：<%
		Set Fs = Server.CreateObject("Scripting.FileSystemObject")
		For Each str In OKAr
		If Fs.FileExists(Server.MapPath("Images\"& Str &".gif")) Then
		Response.Write "<img src='Images/" & str & ".gif' alt='" & str & "文件'> "
		Else
		Response.Write "<img src='Images/X.gif' alt='" & str & "文件'> "
		End If
		Next
		Set Fs = Nothing
		Response.Write "<br>  允許上傳的文件最大: "&Oksize / 1024  & " KB"
		%></td>
						</tr>
						<tr>
							<td valign="top" colspan="2">　　<font color="#FF0000">注意:請務必填寫<b>新聞標題</b>、<b>摘要</b>欄位！</font></td>
						</tr>
						<tr>
							<td valign="top" colspan="2">
							<p align="center"><input type="submit" name="Submit" value="上傳內容"></td>
						</tr>
						<tr>
							<td colspan="2" bgcolor="#328CDD">
							<img border="0" src="jpg" height="1"></td>
						</tr>
</table>
</form>

				
			
			<tr>
				<td width="760" colspan="2">			</td>
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