<!--#include file="Config.asp"-->
<%
CheckLogin
Dim GetID,ID
GetID = split(Trim(Request.Form("DelID")),",")
Set Fs = Server.CreateObject("Scripting.FileSystemObject")
For each Str in GetID
VarStr = Split(Str,"|")
ID = VarStr(0)
Path = VarStr(1)
myconn.execute("Delete From Info Where ID="&ID)
If Fs.FileExists(server.mappath(Path)) Then
Set Os = Fs.GetFile(server.mappath(Path))
Os.Delete
Response.Write Path&"已被刪除！<br>"
Else
Response.Write Path&"此圖片不存在！數據已被刪除<br>"
End If
Next
%>
<meta http-equiv="refresh" content="3;URL=Admin_List.asp">
<link href="upstyle.css" rel="stylesheet" type="text/css">
<body>
<span class=smtext>本頁將在3秒後返回<br>
如果您的瀏覽器沒有反應，請<a href=Admin_List.asp><font color=000000><b>點擊此處返回</b></font></a></span>
</body>