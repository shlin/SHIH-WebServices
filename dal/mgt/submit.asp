<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<%
' �p�G�s�誺���e�ܦh�A�W�ǳt�פӺC�A�г]�m�H�U���ɶ��A����
'Server.ScriptTimeout = 600

Dim sContent1, i

For i = 1 To Request.Form("content1").Count 
	sContent1 = sContent1 & Request.Form("content1")(i) 
Next 

Response.Write "�s�褺�e�p�U�G<br><br>" & sContent1
Response.Write "<br><br><p><input type=button value=' �h�^ ' onclick='history.back()'></p>"

%>