<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="Config.asp"-->
<%
CheckLogin
Dim ID,Msg
	ID = Request.QueryString("ID")
	If Request.QueryString("Action") = "Save" Then SaveData ID
	Sub SaveData(ID)
	If ID < 1 Then 
	Response.Write("�Ѽƿ��~")
	Response.End()
	End If
	myConn.execute("update info set FILETITLE='"&Request.Form("fname")&"',FILEDESC='"&Request.Form("fdesc")&"',FILETYPE='"&Request.Form("ftype")&"',FILEPATH='"&Request.Form("fpath")&"',FILESIZE='"&Request.Form("fsize")&"' where ID="&ID)
	Msg = "���\�ק�F���ƾګH��"
	End Sub

If msg <> "" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_List.asp'>"&Msg&"<br>�����N�b3����^<BR>�p�G�A���s�����S�������A��<a href=Admin_List.asp>�I�����B��^</a>")
Response.End()
End If
%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<title>:::�B�ѰO�W�ǵ{�� ::: �t�ܪ���</title>
<link href="upstyle.css" rel="stylesheet" type="text/css">
<script src="Images/up.Js"></script>
</head>

<body background="../images/mainbg.gif" leftmargin="0" topmargin="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#003366">
  <tr> 
    <td height="75">�@</td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#FF9900" class="text">��x�޲z&nbsp;�X&nbsp;<a href="./">��^�W�ǭ���</a></td>
  </tr>
  <tr>
    <td height="2" bgcolor="#cccccc"></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="778" height="200">
        <param name="movie" value="FLASH_AD.swf">
        <param name="quality" value="high">
        <embed src="FLASH_AD.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="778" height="200"></embed></object></td>
  </tr>
  <tr>
    <td height="1" bgcolor="#CCCCCC"></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
<form name="Edit" action="Edit.asp?Action=Save&ID=<%=ID%>" method="post">
  <tr> 
    <td height="25"> <%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from info Where Id="&ID
	frst.open sql,myconn,1,1
	If not frst.Eof then
			fid = frst("id").Value
			ftitle = frst("fileTitle").Value
			fdesc = frst("fileDesc").Value
			ftype = frst("fileType").Value
			fpath = frst("filePath").Value
			fsize = frst("filesize").Value
			fhits = frst("hits").Value
			fuploadtime = frst("uploadTime").Value
'FileNameStr=Server.Mappath(fpath)
'Isize.GetImgSize Cstr(FileNameStr)
	%>
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
          <tr class="text"> 
            <td width="150"><div align="right">���W�١G</div></td>
            <td>
            <input type="text" name="fname" class="TextBoxT" value="<%=ftitle%>" size="20"> 
            </td>
            <td width="30%" rowspan="3">&nbsp; </td>
          </tr>
          <tr class="text"> 
            <td width="150"><div align="right">��������G</div></td>
            <td><select name="ftype" class="TextBoxT" id="filetype">
                <option value="�����Ϥ�"<%if ftype="�����Ϥ�" then%> selected<% end if %>>�����Ϥ�</option>
                <option value="�`�Τu��"<%if ftype="�`�Τu��" then%> selected<% end if %>>�`�Τu��</option>
                <option value="�{�Ƿ��X"<%if ftype="�{�Ƿ��X" then%> selected<% end if %>>�{�Ƿ��X</option>
              </select></td>
          </tr>
          <tr class="text"> 
            <td width="150"><div align="right">�����|�G</div></td>
            <td><input name="fpath" type="text" class="TextBoxT" value="<%=fpath%>" size="50"> 
              <%
			Set Fs = Server.CreateObject("Scripting.FileSystemObject")
			If Fs.FileExists(server.mappath(fPath)) Then
			Response.Write("<img src=Images/isexists.gif")
			End If
			%>
            </td>
          </tr>
          <tr class="text"> 
            <td width="150"><div align="right">�����G</div></td>
            <td colspan="2">
            <input type="text" name="fdesc" class="TextBoxT" value="<%=fdesc%>" size="20"></td>
          </tr>
          <tr class="text"> 
            <td height="1" align="right" bgcolor="#0153A3">���j�p�G</td>
            <td height="1" colspan="2" bgcolor="#0153A3"> 
            <input type="text" name="fsize" class="TextBoxT" value="<%=fsize%>" size="20">
              bytes</td>
          </tr>
          <tr class="text"> 
            <td height="1" align="right" bgcolor="#0153A3">�@</td>
            <td height="1" colspan="2" bgcolor="#0153A3"> <input type="submit" name="Submit" value="�ק�">&nbsp;<input type="button" name="Submit2" value="��^" onclick='javascript:window.location="Admin_List.asp"'></td>
          </tr>
          <tr class="text"> 
            <td height="1" align="right" bgcolor="#0153A3">�@</td>
            <td height="1" colspan="2" bgcolor="#0153A3"></td>
          </tr>
        </table>
      <%
	  else
	  %> 
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
          <td>�S���������ƾڡI</td>
        </tr>
      </table>
      <%
	  end if
	  frst.close
	  set frst = nothing
	  myconn.close
	  set myconn = nothing
	  %>
	  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
            <td align="right" class="heading">&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table></td>
  </tr></form>
</table>

<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
  <tr> 
    <td height="25">�@</td>
  </tr>
</table>
</body>
</html>