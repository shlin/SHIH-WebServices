<%
''Option Explicit 
Server.ScriptTimeout = 999     '�n�W�Ǹ��j�����N�n�}�Ҧ���

Dim myconn,connstr
connstr = "Provider=Microsoft.Jet.Oledb.4.0;data source="&Server.MapPath("module/news/Upload.mdb")
set myconn = Server.CreateObject("adodb.connection")
myconn.open connstr

Dim OKAr
Dim OKsize

Set rs = myconn.execute("select * from Config")
'***************�]�m���\�W�Ǫ��������********************

OKAr = split(rs(0),",")

'***********************************

'***************�]�m���\�W�Ǫ����j�p********************

OKsize = Clng(rs(1))

'***********************************

Function Bytes2bStr(vin)
if lenb(vin) =0 then
	Bytes2bStr = ""
	exit function
end if
''�G�i���ഫ���r�Ŧ�
 Dim BytesStream,StringReturn
 Set BytesStream = Server.CreateObject("ADODB.Stream")
 BytesStream.Type = 2 
 BytesStream.Open
 BytesStream.WriteText vin
 BytesStream.Position = 0
 BytesStream.Charset = "big5"
 BytesStream.Position = 2
 StringReturn = BytesStream.ReadText
 BytesStream.close
 Set BytesStream = Nothing
 Bytes2bStr = StringReturn
End Function

Function Myrequest(fldname)
	''�����ƾ�
	''�����P�W���쪺Ū��
	dim i
	dim fldHead
	dim tmpvalue
	for i = 0 to loopcnt-1
		fldHead = fldInfo(i,0)
		if instr(lcase(fldHead),lcase(fldname))>0 then
			''���b�Ʋդ�
			''�P�_�Ӫ��줺�e
			tmpvalue = FldInfo(i,1)
			if instr(fldHead,"filename=""")<1 then
				Tmpvalue = Bytes2bStr(tmpvalue)
				if myrequest <> "" then 
					myrequest = myrequest & "," &tmpvalue
				else
					MyRequest = tmpvalue
				end if				
			else
				myrequest = tmpvalue
			end if				
		end if
	next
end function
function GetFileName(fldName)
	''������W�Ǥ����W
	dim i
	dim fldHead
	dim fnpos
	for i = 0 to loopcnt-1
		fldHead = lcase(fldInfo(i,0))
		if instr(fldHead,lcase(fldName)) > 0 then
			fnpos = instr(fldHead,"filename=""")
			if fnpos < 1 then exit for
			fldHead = mid(fldHead,fnpos+10)
			''��椺�e
			GetFileName = mid(fldHead,1,instr(fldHead,"""")-1)
			GetfileName = mid(GetFileName,instrRev(GetFileName,"\")+1)
		end if
	next
end function
function GetContentType(fldName)
	''Ū���W�Ǥ������
	''���wŪ�����쪺���e
	dim i
	dim fldHead,cpos
	for i = 0 to loopcnt - 1
		fldHead = lcase(fldInfo(i,0))
		if instr(fldHead,lcase(fldName)) > 0 and instr(fldHead,"filename=""") >0 then
			''Ū��contentType
			cpos = instr(fldHead,"content-type: ")
			GetContentType = mid(fldHead,cpos+14)
		end if
	next
end function
Sub SaveToFile(fd,path,fname)
''�O�s���''�Ѽƻ���:''fd:byte�]�^�����ƾڡA��󤺮e''path:�O�s���|�᭱�����a"/"''fname:���W
	dim Fstream
	Set FStream = Server.CreateObject("adodb.stream")
	fstream.mode = 3
	fstream.type = 1
	fstream.open
	fstream.position = 0
	fstream.Write fd
	fstream.savetofile Server.Mappath(path&fname),2
	fstream.close
	set fstream = nothing
end sub

Function GetFileTypeName(Fldname)
If instr(Fldname,".") > 0 Then
GetFileTypeName = right(Fldname,3)
Else
Response.Write "���W�D�k�A�Эק��A�W��"
Response.End()
End If
End Function

Function IsvalidFile(TypeName)   '����W�Ǥ������
IsvalidFile = False
Dim GName
For Each GName in OKAr
If TypeName = GName Then
IsvalidFile = True
Exit For
End If
Next
End Function

Sub CheckLogin()
If Session("Admin") <> "OK" Then
Logined = False
Response.Write ("���q�L�������ҡA��<a href=../../get_content_news.asp?id=10>��^</a>")
Response.Redirect ("../../get_content_news.asp?id=10")
Response.End()
End If
End Sub

%>