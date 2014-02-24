<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="Config.asp"-->
<%
''主程序開始部分

dim formsize,formdata,Msg
formsize = Request.TotalBytes
formdata = Request.BinaryRead(formsize)
UploadSize=True

If formsize = 0 or Formsize > OKsize Then
UploadSize=False
Response.Write"你要上傳的文件大小超出程序限制,請<a href=index.asp>返回</a>修改重試"
Response.End
End If


dim sinfo_Stream
Set Sinfo_Stream = Server.CreateObject("adodb.stream")
Sinfo_Stream.Type = 1		''2進制流
Sinfo_Stream.Mode = 3		''讀寫模式
Sinfo_Stream.Open
Sinfo_Stream.Write formdata		''保存二進制內容到流對像
''分離數據變量
dim VbEnter
dim spStr,lenOfspStr,bpos
dim loopcnt,exitflag,ppoint,npoint
''保存數據變量		
dim FldData,fldHeadStr,infldpos
dim databpos,datalen
dim FldInfo(15,1)
''fldInfo(0)表單頭內容
''fldInfo(1)表單數據

VbEnter = chrb(13)&chrb(10)''讀取第一個VbEnter位置
bpos = Instrb(formdata,VbEnter)
SpStr = midb(formdata,1,bpos+1) ''包含了一個0d0a
LenOfspStr = lenb(Spstr) 
ppoint = LenOfspStr+1 ''位置指針,指向每一個表單域內容的開始位置
formdata = midb(formdata,ppoint)
loopcnt = 0   ''表單元素
do 
	bpos = instrb(formdata,spStr) ''分割位置
	npoint = (ppoint+bpos+lenofspstr-1)  ''指向下一表單開始位置
	if bpos < 1 then
		fldData = midb(formdata,1,instrb(formdata,leftb(spStr,lenOfspstr-2))-1)
		bpos = lenb(fldData)+1
		exitflag = true
	else
		FldData = leftb(formdata,bpos-1)		
		formdata = midb(formdata,bpos+LenOfspstr)
	end if
	infldpos = instrb(fldData,vbEnter&vbEnter)
	fldHeadStr = bytes2bstr(midb(fldData,1,infldpos-1))
	fldInfo(loopcnt,0) = fldHeadStr	''表單頭
	''Response.Write fldHeadStr&"<br>"
	databpos = (ppoint+infldpos-1+4)
	Sinfo_Stream.Position = databpos-1
	datalen = (bpos-infldpos-6)
	if datalen = 0 then
		fldInfo(loopcnt,1) = ""
	else
		fldInfo(loopcnt,1) = Sinfo_Stream.Read(datalen)
	end if
	ppoint = npoint
	loopcnt = loopcnt + 1
loop until exitflag = true
Sinfo_Stream.close
Set Sinfo_Stream = Nothing


''以上程序數據處理過程
''寫入數據庫並處理文件上傳開始
Sub SaveData()

ftitle = MyRequest("filetitle")
Msg = ""
		if ftitle = "" then 
			Msg = Msg & "新聞標題：空<br>"
		else
			Msg = Msg & "新聞標題："&ftitle&"<br>"
		end if

ftime= MyRequest("NewsTime")
	Msg = Msg & "新聞時間："&ftime&"<br><br>"
	
	
	
ftype = myrequest("fileType")		
		Msg = Msg & "文件類型："&ftype&"<br>"

filedata = myrequest("filedata")

filesize = lenb(filedata)
		if  filesize = 0 then 
			Msg = Msg & "上傳文件：沒有<br>"
		else 
			filename = GetFileName("filedata")
			''限制加入的類型 *.asp
			file_ctype = GetContentType("filedata")
			Msg = Msg & "上傳文件："&filename&"&nbsp;&nbsp;&nbsp;"
			Msg = Msg & "數據流："&file_ctype&"&nbsp;&nbsp;&nbsp;"
			Msg = Msg & "文件長度："&filesize&"<br>"
		end if

filedesc = myrequest("fileDesc")
		Msg = Msg & "摘要："&filedesc&"<br><br>"
		
FileTypeName = GetFileTypeName(FileName)
		If  IsvalidFile(FileTypeName)=False Then
		Msg = "文件類型非法，不允許上傳"&FileTypeName&"文件！"
		Exit Sub
		End If		

		if ftitle<>"" and fileSize > 0 and UploadSize=True then
			''保存數據到數據庫

			dim basepath,sql
			basepath = "./../../File/決策新聞/"
			sql = "insert into info (filetitle,filedesc,filetype,filecontenttype,filepath,filesize) values ('"
			sql = sql & ftitle &"','"&filedesc&"','"&ftype&"','"&file_ctype&"','"&basepath&filename&"',"&filesize&")"
			myconn.Execute(sql)
			Call SavetoFile(filedata,basepath,filename)
			Msg = Msg & "文件已經上傳<br>"
		else
			Msg = Msg & "上傳失敗！ "&ErrorMsg&"<br>"
		end if
		myconn.close()
		set myconn = nothing

End Sub	
''文件上傳已經寫入數據完畢，提示信息出口為變量msg
SaveData
%>

<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>上傳程式</title>
<meta http-equiv="refresh" content="3;URL=Index.asp">
</head>
<LINK href="master.css" type=text/css rel=stylesheet>
<body topmargin="0" leftmargin="0">

<table border="0" cellpadding="0" style="border-collapse: collapse" width="650" id="table1" height="500">
	<tr>
		<td valign="top">
		<table border="0" cellpadding="0" style="border-collapse: collapse" width="760" id="table2">
			<tr>
				<td width="449" height="30" bgcolor="#EEEEEE">上傳程式</td>
				<td width="311" bgcolor="#EEEEEE">
　|　<a href="?page=1">首頁</a>
　|　<a href="fileupload.asp">上傳</a>
　|　<a href="admin.asp">管理</a>
　|　

				</td>
			</tr>
			<tr>
				<td width="760" colspan="2">　　　　　上傳<hr color="#C0C0C0" size="1" width="90%"></td>
			</tr>
			<tr>
				<td width="760" colspan="2" height="350" valign="top">　
				<div align="center">
					<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="80%" id="table4">
						<tr>
							<td>
							提示信息：<br>　<%=msg%>
							
							<br><br><div align="center">本頁將在3秒後返回 如果您的瀏覽器沒有反應，請<a href=Index.asp>點擊此處返回</a></div>
							
							</td>
						</tr>
					</table>
					
					
				</div>

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