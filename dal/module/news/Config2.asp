<%
''Option Explicit 
Server.ScriptTimeout = 999     '要上傳較大的文件就要開啟此項

Dim myconn,connstr
connstr = "Provider=Microsoft.Jet.Oledb.4.0;data source="&Server.MapPath("module/news/Upload.mdb")
set myconn = Server.CreateObject("adodb.connection")
myconn.open connstr

Dim OKAr
Dim OKsize

Set rs = myconn.execute("select * from Config")
'***************設置允許上傳的文件類型********************

OKAr = split(rs(0),",")

'***********************************

'***************設置允許上傳的文件大小********************

OKsize = Clng(rs(1))

'***********************************

Function Bytes2bStr(vin)
if lenb(vin) =0 then
	Bytes2bStr = ""
	exit function
end if
''二進制轉換為字符串
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
	''取表單數據
	''支持對同名表單域的讀取
	dim i
	dim fldHead
	dim tmpvalue
	for i = 0 to loopcnt-1
		fldHead = fldInfo(i,0)
		if instr(lcase(fldHead),lcase(fldname))>0 then
			''表單在數組中
			''判斷該表單域內容
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
	''都取原上傳文件文件名
	dim i
	dim fldHead
	dim fnpos
	for i = 0 to loopcnt-1
		fldHead = lcase(fldInfo(i,0))
		if instr(fldHead,lcase(fldName)) > 0 then
			fnpos = instr(fldHead,"filename=""")
			if fnpos < 1 then exit for
			fldHead = mid(fldHead,fnpos+10)
			''表單內容
			GetFileName = mid(fldHead,1,instr(fldHead,"""")-1)
			GetfileName = mid(GetFileName,instrRev(GetFileName,"\")+1)
		end if
	next
end function
function GetContentType(fldName)
	''讀取上傳文件的類型
	''限定讀取文件域的內容
	dim i
	dim fldHead,cpos
	for i = 0 to loopcnt - 1
		fldHead = lcase(fldInfo(i,0))
		if instr(fldHead,lcase(fldName)) > 0 and instr(fldHead,"filename=""") >0 then
			''讀取contentType
			cpos = instr(fldHead,"content-type: ")
			GetContentType = mid(fldHead,cpos+14)
		end if
	next
end function
Sub SaveToFile(fd,path,fname)
''保存文件''參數說明:''fd:byte（）類型數據，文件內容''path:保存路徑後面必須帶"/"''fname:文件名
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
Response.Write "文件名非法，請修改後再上傳"
Response.End()
End If
End Function

Function IsvalidFile(TypeName)   '限制上傳文件類型
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
Response.Write ("未通過身份驗證，請<a href=../../get_content_news.asp?id=10>返回</a>")
Response.Redirect ("../../get_content_news.asp?id=10")
Response.End()
End If
End Sub

%>