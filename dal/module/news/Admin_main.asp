<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="Config.asp"-->
<%
CheckLogin
Dim Msg
	If Request.QueryString("Action") = "Save" Then SaveData
	Sub SaveData()
	myConn.execute("update Config set OKAr='"&Request.Form("ftype")&"',OKsize="&Request.Form("fsize"))
	Msg = "���\�ק�F���ƾګH��"
	End Sub

If msg <> "" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_Main.asp'>"&Msg&"<br>�����N�b3����^<BR>�p�G�A���s�����S�������A��<a href=Admin_Main.asp>�I�����B��^</a>")
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
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="75">&nbsp;</td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#FF9900" class="text">��e��m�G:::�B�ѰO�W�ǵ{��Ver1.2 ::: �t�ܪ��� 
      &gt;&gt;&gt; ��x�޲z</td>
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
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
<form name="Edit" action="Admin_Main.asp?Action=Save" method="post">
  <tr> 
    <td height="25"> <%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from Config"
	frst.open sql,myconn,1,1
	If not frst.Eof then
	%>
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
          <tr align="center" class="text"> 
            <td height="1" colspan="2">��x�޲z-�t�γ]�m&nbsp;&nbsp;[<a href="Admin_List.asp">�I���i�J���޲z</a>]&nbsp;&nbsp;[<a href="Index.asp">�h�X�޲z</a>]</td>
          </tr>
          <tr class="text"> 
            <td width="200"> <div align="right">���\�W�Ǫ���������G</div></td>
            <td><input name="ftype" type="text" class="TextBoxT" value="<%=rs(0)%>" size="50">
              �H,���j���W,���O�Ť��\�W��asp/exe���</td>
          </tr>
          <tr class="text"> 
            <td width="200"><div align="right">���\�W�Ǫ����j�p�G</div></td>
            <td><input name="fsize" type="text" class="TextBoxT" value="<%=rs(1)%>" size="15">
              ���:Byte</td>
          </tr>
          <tr class="text"> 
            <td height="1" align="right" bgcolor="#0153A3">&nbsp;</td>
            <td height="1" bgcolor="#0153A3"> <input type="submit" name="Submit" value="�ק�"></td>
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
    <td height="25">&nbsp;</td>
  </tr>
</table>
</body>
</html>