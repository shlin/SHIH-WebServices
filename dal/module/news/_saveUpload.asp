<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="Config.asp"-->
<%
''�D�{�Ƕ}�l����

dim formsize,formdata,Msg
formsize = Request.TotalBytes
formdata = Request.BinaryRead(formsize)
UploadSize=True

If formsize = 0 or Formsize > OKsize Then
UploadSize=False
Response.Write"�A�n�W�Ǫ����j�p�W�X�{�ǭ���,��<a href=index.asp>��^</a>�קﭫ��"
Response.End
End If


dim sinfo_Stream
Set Sinfo_Stream = Server.CreateObject("adodb.stream")
Sinfo_Stream.Type = 1		''2�i��y
Sinfo_Stream.Mode = 3		''Ū�g�Ҧ�
Sinfo_Stream.Open
Sinfo_Stream.Write formdata		''�O�s�G�i��e��y�ﹳ
''�����ƾ��ܶq
dim VbEnter
dim spStr,lenOfspStr,bpos
dim loopcnt,exitflag,ppoint,npoint
''�O�s�ƾ��ܶq		
dim FldData,fldHeadStr,infldpos
dim databpos,datalen
dim FldInfo(15,1)
''fldInfo(0)����Y���e
''fldInfo(1)���ƾ�

VbEnter = chrb(13)&chrb(10)''Ū���Ĥ@��VbEnter��m
bpos = Instrb(formdata,VbEnter)
SpStr = midb(formdata,1,bpos+1) ''�]�t�F�@��0d0a
LenOfspStr = lenb(Spstr) 
ppoint = LenOfspStr+1 ''��m���w,���V�C�@�Ӫ��줺�e���}�l��m
formdata = midb(formdata,ppoint)
loopcnt = 0   ''��椸��
do 
	bpos = instrb(formdata,spStr) ''���Φ�m
	npoint = (ppoint+bpos+lenofspstr-1)  ''���V�U�@���}�l��m
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
	fldInfo(loopcnt,0) = fldHeadStr	''����Y
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


''�H�W�{�ǼƾڳB�z�L�{
''�g�J�ƾڮw�óB�z���W�Ƕ}�l
Sub SaveData()

ftitle = MyRequest("filetitle")
Msg = ""
		if ftitle = "" then 
			Msg = Msg & "�s�D���D�G��<br>"
		else
			Msg = Msg & "�s�D���D�G"&ftitle&"<br>"
		end if

ftime= MyRequest("NewsTime")
	Msg = Msg & "�s�D�ɶ��G"&ftime&"<br><br>"
	
	
	
ftype = myrequest("fileType")		
		Msg = Msg & "��������G"&ftype&"<br>"

filedata = myrequest("filedata")

filesize = lenb(filedata)
		if  filesize = 0 then 
			Msg = Msg & "�W�Ǥ��G�S��<br>"
		else 
			filename = GetFileName("filedata")
			''����[�J������ *.asp
			file_ctype = GetContentType("filedata")
			Msg = Msg & "�W�Ǥ��G"&filename&"&nbsp;&nbsp;&nbsp;"
			Msg = Msg & "�ƾڬy�G"&file_ctype&"&nbsp;&nbsp;&nbsp;"
			Msg = Msg & "�����סG"&filesize&"<br>"
		end if

filedesc = myrequest("fileDesc")
		Msg = Msg & "�K�n�G"&filedesc&"<br><br>"
		
FileTypeName = GetFileTypeName(FileName)
		If  IsvalidFile(FileTypeName)=False Then
		Msg = "��������D�k�A�����\�W��"&FileTypeName&"���I"
		Exit Sub
		End If		

		if ftitle<>"" and fileSize > 0 and UploadSize=True then
			''�O�s�ƾڨ�ƾڮw

			dim basepath,sql
			basepath = "./../../File/�M���s�D/"
			sql = "insert into info (filetitle,filedesc,filetype,filecontenttype,filepath,filesize) values ('"
			sql = sql & ftitle &"','"&filedesc&"','"&ftype&"','"&file_ctype&"','"&basepath&filename&"',"&filesize&")"
			myconn.Execute(sql)
			Call SavetoFile(filedata,basepath,filename)
			Msg = Msg & "���w�g�W��<br>"
		else
			Msg = Msg & "�W�ǥ��ѡI "&ErrorMsg&"<br>"
		end if
		myconn.close()
		set myconn = nothing

End Sub	
''���W�Ǥw�g�g�J�ƾڧ����A���ܫH���X�f���ܶqmsg
SaveData
%>

<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�W�ǵ{��</title>
<meta http-equiv="refresh" content="3;URL=Index.asp">
</head>
<LINK href="master.css" type=text/css rel=stylesheet>
<body topmargin="0" leftmargin="0">

<table border="0" cellpadding="0" style="border-collapse: collapse" width="650" id="table1" height="500">
	<tr>
		<td valign="top">
		<table border="0" cellpadding="0" style="border-collapse: collapse" width="760" id="table2">
			<tr>
				<td width="449" height="30" bgcolor="#EEEEEE">�W�ǵ{��</td>
				<td width="311" bgcolor="#EEEEEE">
�@|�@<a href="?page=1">����</a>
�@|�@<a href="fileupload.asp">�W��</a>
�@|�@<a href="admin.asp">�޲z</a>
�@|�@

				</td>
			</tr>
			<tr>
				<td width="760" colspan="2">�@�@�@�@�@�W��<hr color="#C0C0C0" size="1" width="90%"></td>
			</tr>
			<tr>
				<td width="760" colspan="2" height="350" valign="top">�@
				<div align="center">
					<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="80%" id="table4">
						<tr>
							<td>
							���ܫH���G<br>�@<%=msg%>
							
							<br><br><div align="center">�����N�b3����^ �p�G�z���s�����S�������A��<a href=Index.asp>�I�����B��^</a></div>
							
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