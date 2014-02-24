<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<%
' 如果編輯的內容很多，上傳速度太慢，請設置以下的時間，單位秒
'Server.ScriptTimeout = 600

Dim sContent1, i

For i = 1 To Request.Form("content1").Count 
	sContent1 = sContent1 & Request.Form("content1")(i) 
Next 

Response.Write "編輯內容如下：<br><br>" & sContent1
Response.Write "<br><br><p><input type=button value=' 退回 ' onclick='history.back()'></p>"

%>