<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="Config.asp"-->
<%
CheckLogin
%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="keywords" content="�B�ѰO�W�ǵ{�� search engine marketing ">
<title></title>
<link href="upstyle.css" rel="stylesheet" type="text/css">
<script src="Images/up.Js"></script>
</head>

<body background="../images/mainbg.gif" leftmargin="0" topmargin="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#003366">
  <tr> 
    <td height="75">&nbsp;</td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#FF9900" class="text">��x�޲z</td>
  </tr>
  <tr>
    <td height="2" bgcolor="#cccccc"></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="1" bgcolor="#CCCCCC"></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0"><tr class="text"> 
          
    <td align="center">��x�޲z-���޲z&nbsp;&nbsp;[<a href="Admin_main.asp">�I���i�J�t�γ]�m</a>]&nbsp;&nbsp;[<a href="Index.asp">�h�X�޲z</a>]</td>
        </tr>
<form name="del" action="del.asp" method="post">
  <tr> 
    <td height="25"> <%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from info order by uploadtime desc"
	frst.open sql,myconn,1,1
	fcount = frst.recordcount
	if fcount > 0 then 	
		''��ܰѼ�
		dim tbbgcolor
		''�����Ѽ�
		dim maxperpage,pages,page
		maxperpage = 5
		frst.pagesize = maxperpage
		pages = frst.pagecount
		''�����ѼƳ]�m
		page = Request.QueryString("page")
		if not isnumeric(page) then page = 1 else page = cint(page)
		if page < 1 then page = 1
		if page > pages then page = pages
		frst.absolutepage = page
		''��ܤ��e
'Set Isize=Server.CreateObject("WinImg.ImgSize")
		for i = 1 to maxperpage
			if frst.eof then exit for
			if i mod 2 = 0 then tbbgcolor = "" else tbbgcolor = "#0066cc"
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
          <td><a href="<%=fpath%>" target="_NEWwIN"><%=ftitle%>( ���j�p�G<%=fsize%> bytes)</a> </td>
          <td width="20%">�R�� 
              <input type="checkbox" name="DelID" value="<%=fid&"|"&fpath%>">
          </td>
        </tr>
        <tr class="text"> 
          <td width="150"><div align="right">��������G</div></td>
          <td><%=ftype%></td>
            <td><a href="Edit.asp?ID=<%=Fid%>">�s��</a></td>
        </tr>
        <tr class="text"> 
          <td width="150"><div align="right">�W�Ǥ���G</div></td>
          <td><%=fuploadtime%></td>
            <td>
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
          <td colspan="2"><%=fdesc%></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="1"></td>
          <td height="1" colspan="2"></td>
        </tr>
      </table>
      <%
		  	frst.movenext
		next
	  else
	  %> 
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
          <td>�٨S���W�Ǫ����e�I</td>
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
            <td align="right" class="heading">��<img src="Images/isexists.gif" width="16" height="16">�аO�h�N����s�b 
              <%
		  If Page > 2 Then Response.Write ("<a href='?page=1'>����</a><a href='?page="& Page - 1 &"'>�W�@��</a>")
		  If page < pages Then Response.Write ("&nbsp;<a href='?page="& Page + 1 &"'>�U�@��</a>&nbsp;<a href='?page="& Pages &"'>����</a>")
		  %>
            �襤�Ҧ�<input name="chkall" type="checkbox" id="chkall" value="select" onclick=CheckAll(this.form)>
              <input type="submit" name="Submit" value="�R���ҿ�"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table></td>
  </tr></form>
</table>

<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
  <tr> 
    <td height="25">&nbsp;</td>
  </tr>
</table>
</body>
</html>
