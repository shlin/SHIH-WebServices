<% ' �}��Ʈw
   Set conn = Server.CreateObject("ADODB.Connection")
   DBPath = Server.MapPath("download.mdb")
   conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
   SET rs = Server.CreateObject("ADODB.Recordset")
   rs.Open "Select * From download", conn,1,1
%>
<html>
<head>
<title>Asp �ɮפU��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<STYLE TYPE="text/css">
        <!-- body,th,td,input,select,textarea,select,checkbox {font:9pt �s�ө���}
         a:link {font: 9pt "�s�ө�"; text-decoration: none;color:none}
         a:visited      {font: 9pt "�s�ө�"; text-decoration: none; color: 000099}
         a:active       {font: 9pt "�s�ө�"; text-decoration: none; color: 00CC66}
         a:hover        {font: 9pt �s�ө�; text-decoration: underline; color: ff0000}
        -->
   </STYLE>
</head>
<body background="�۶��a�O.gif">
<table width="60%" border="0" align="center" height="170">
  <tr> 
    <td colspan="14"> 
      <div align="center"><font size="6" face="�з���">�ɮפU��</font></div>
    </td>
  </tr>
  <tr> 
    <td colspan="14"> 
      <hr size="1" noshade>
    </td>
  </tr>
  <tr> 
    <td height="50"> 
      <div align="left"> 
        <p>&nbsp;</p>
        </div>
    </td>
    <td colspan="13" height="50"> 
      <p><font color="#990000">�H�U���Ѥ��ɮק������U���ҵ{�оǡA�Ф�����γs���C</font> 
        </p>
      </td>
  </tr>
  <tr> 
    <td colspan="14" height="3"> 
      <div align="center"> 
        <hr size="1" noshade>
      </div>
    </td>
  </tr>
  <tr> 
    <td width="9%" nowrap bgcolor="#CCFFCC"> 
      <div align="center"><font color="#000099">���Ѥ��</font></div>
    </td>
    <td width="10%" nowrap bgcolor="#CCFFCC"> 
      <div align="center"><font color="#000099">�ɮצW��</font></div>
    </td>
    <td bgcolor="#CCFFCC" colspan="6">
      <div align="center"><font color="#000099">���@�@��</font></div>
    </td>
    <td width="5%" nowrap bgcolor="#CCFFCC"> 
      <div align="center"><font color="#000099">����</font>
    </td>
    <td width="5%" nowrap bgcolor="#CCFFCC"> 
      <div align="center"><font color="#000099">�s����</font></div>
    </td>
    <td width="5%" nowrap bgcolor="#CCFFCC"> 
      <div align="center"><font color="#000099">�U��</font></div>
    </td>
    <td width="10%" nowrap bgcolor="#CCFFCC"> 
      <div align="center"><font color="#000099">SIZE</font></div>
    </td>
    <td width="15%" nowrap bgcolor="#CCFFCC"><font color="#000098">�U������</font></div>
   </td>
    <td width="15%" nowrap bgcolor="#CCFFCC"> 
      <div align="center"><font color="#000099">���Ѫ�</font></div>
    </td>
  </tr>
  <tr> 
    <td nowrap colspan="14"> 
      <hr size="1" noshade>
    </td>
  </tr>
<%FOR I=1 TO RS.RECORDCOUNT %>
  <tr bgcolor="#DFFBFF"> 
    <td width="9%" nowrap rowspan="2"> 
      <div align="center"><%=rs("designdate")%></div>
    </td>
    <td width="10%" nowrap rowspan="2">
      <div align="center"><font color="#990000"><%=rs("programName")%></font></div>
    </td>
<%IF len(TRIM(rs("message"))) <> 0 Then%>
    <td nowrap width="30%" colspan="6" rowspan="2"> <font color="#980098"><%=rs("message")%></font></td>
<%ELSE%>
    <td nowrap width="5%">
      <div align="center"><font color="#990099">�s�W</font></div>
    </td>
    <td nowrap width="5%">
      <div align="center"><font color="#990099">�R��</font></div>
    </td>
    <td nowrap width="5%">
      <div align="center"><font color="#990099">�ק�</font></div>
    </td>
    <td nowrap width="5%">
      <div align="center"><font color="#990099">�^��</font></div>
    </td>
    <td nowrap width="5%">
      <div align="center"><font color="#990099">�޲z</font></div>
    </td>
    <td nowrap width="5%">
      <div align="center"><font color="#990099">����</font></div>
    </td>
<%End IF%>
    <td rowspan="2" width="5%"> 
      <div align="center"><%=rs("version")%></div>
    </td>
    <td width="11%" rowspan="2" nowrap> 
      <div align="center"><%=rs("browser")%></div>
    </td>
    <td width="5%" rowspan="2"> 
      <div align="center"><a href="redirect.asp?fileid=<%=rs("fileid")%>"><img src="dn_zip.gif" width="16" height="16" border="0"></a></div>
    </td>
    <td width="10%" rowspan="2"> 
      <div align="center"><%=rs("filesize")%></div>
    </td>
    <td width="20%" rowspan="2"> 
      <div align="center"><font color="#CC0000"><%=rs("hits")%></font></div>
    </td>
    <td width="11%" rowspan="2"> 
      <div align="center"><font color="#CC0000"><%=rs("author")%></font></div>
    </td>
  </tr>
<%IF len(TRIM(rs("message"))) <> 0 Then %>
  <tr>
  </tr>
<%ELSE%>
  <tr> 
    <td width="5%" bgcolor="#E8E8FF"> 
      <div align="center"><%IF rs("add") = "��" Then response.write "��" End IF%></div>
    </td>
    <td width="5%" bgcolor="#E8E8FF"> 
      <div align="center"><%IF rs("del") = "��" Then response.write "��" End IF%></div>
    </td>
    <td width="5%" bgcolor="#E8E8FF"> 
      <div align="center"><%IF rs("modify") = "��" Then response.write "��" End IF%></div>
    </td>
    <td width="5%" bgcolor="#E8E8FF"> 
      <div align="center"><%IF rs("reply") = "��" Then response.write "��" End IF%></div>
    </td>
    <td width="5%" bgcolor="#E8E8FF"> 
      <div align="center"><%IF rs("manager") = "��" Then response.write "��" End IF%></div>
    </td>
    <td width="5%" bgcolor="#E8E8FF"> 
      <div align="center"><%IF rs("page") = "��" Then response.write "��" End IF%></div>
    </td>
  </tr>
<%End IF%>
  <tr> 
    <td colspan="14" height="3"> 
      <div align="center"> 
        <hr size="1" noshade>
      </div>
    </td>
  </tr>
<%RS.MOVENEXT
  IF RS.EOF THEN EXIT FOR
NEXT%>
  <tr> 
    <td colspan="13" height="23"> 
      <div align="center">DownLoad Asp Design by junior 1999/8/28</div>
    </td>
  </tr>
</table>
</body>
</html>