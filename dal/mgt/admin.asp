<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="../Connections/Dal.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="../index.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
Dim GetContent
Dim GetContent_cmd
Dim GetContent_numRows

Set GetContent_cmd = Server.CreateObject ("ADODB.Command")
GetContent_cmd.ActiveConnection = MM_Dal_STRING
GetContent_cmd.CommandText = "SELECT * FROM dbo.Content ORDER BY id ASC" 
GetContent_cmd.Prepared = true

Set GetContent = GetContent_cmd.Execute
GetContent_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
GetContent_numRows = GetContent_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>::: DAL - 決策分析實驗室 :::</title>

<link href="../dal.css" rel="stylesheet" type="text/css" />
<script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script><style type="text/css">
<!--
body {
	background-image: url(../img/body.gif);
	background-attachment : fixed;
}
.style1 {font-size: 18px; color: #666666; text-transform: capitalize; font-variant: normal; font-family: Arial, Helvetica, sans-serif;}

-->
</style>

<script type="text/JavaScript">
<!--



function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>


</head>

<body>



<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
    <td><img src="../img/banner.gif" alt="banner" width="758" height="115" border="0" usemap="#Map2"  />
      <map name="Map2" id="Map22">
        <area shape="rect" coords="157,61,366,96" href="http://www.ms.tku.edu.tw" alt="淡江大學管理科學研究所" />
        <area shape="rect" coords="410,64,610,97" href="http://www.im.tku.edu.tw" alt="淡江大學資訊管理研究所" />
        <area shape="rect" coords="221,10,543,56" href="../index.asp" alt="決策分析研究室" />
        <area shape="circle" coords="65,57,50" href="../index.asp" alt="決策分析研究室" />
        <area shape="circle" coords="690,57,51" href="http://www.tku.edu.tw" alt="淡江大學" />
      </map></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><div align="right" class="smalltext"></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right"><img src="../img/left.gif" alt="left" width="5" height="400" align="right" /></td>
    <td valign="top"><table width="760" border="0" align="center" cellpadding="00" cellspacing="0" class="border"   id="leftlink">
      <tr>
        <td width="152" valign="middle" background="../img/MenuBg.jpg" class="MenuSmall"><span class="smalltext"><img src="../img/w.gif" alt="w" width="15" height="15" /></span></td>
        <td width="608" height="24" colspan="2" align="right" valign="middle" bgcolor="#1ca4c0"><table width="100%" border="0" cellpadding="0" cellspacing="0" >
          <tr>
            <td><div align="left" class="MenuSmall"><img src="../img/w.gif" alt="w" width="15" height="15" />::: Every decision you make is a mistake. -- Edward Dahlberg :::</div></td>
           
            <td><div align="right" class="MenuSmall" ><img src="../img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Search</div></td>
            <td><div align="right"><span class="MenuSmall"><img src="../img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Contact Us</span></div></td>
            <td><span class="MenuSmall" onclick="MM_goToURL('parent','../index.asp');return document.MM_returnValue"><img src="../img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Home<span class="smalltext"><img src="../img/w.gif" alt="w" width="15" height="15" /></span>　</span></td>
          </tr>
        </table></td>
        </tr>
      <tr>
        <td colspan="2" valign="top"><span class="smalltext"><img src="../img/w.gif" alt="w" width="15" height="10" /></span></td>
      </tr>
      <tr>
        <td colspan="6" valign="top"><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td colspan="4"><p><span class="InsideTitle">◎主要類別│</span><span class="text"><a href="add_Content.asp">新增</a></span></p>
              </td>
            </tr>
			 <tr>
			   	<td colspan="4" bgcolor="#1ca4c0"><img border="0" src="jpg" height="1"></td>
              </tr>
          <tr>
            <td width="14%"><strong>索引鍵</strong></td>
            <td><strong>英文類別</strong></td>
            <td width="35%"><strong>中文類別</strong></td>
            <td>&nbsp;</td>
            </tr>
			   <tr>
			   	<td colspan="4" bgcolor="#1ca4c0"><img border="0" src="jpg" height="1"></td>
              </tr>
            <% 
While ((Repeat1__numRows <> 0) AND (NOT GetContent.EOF)) 
%>
             
                <tr>

              <td><%=(GetContent.Fields.Item("id").Value)%></td>
              <td><%=(GetContent.Fields.Item("name").Value)%></td>
              <td><%=(GetContent.Fields.Item("cname").Value)%></td>
              <td><a href="Modify_Content.asp?id=<%=(GetContent.Fields.Item("id").Value)%>">修改</a>│
			 <a href="del_Content.asp?id=<%=(GetContent.Fields.Item("id").Value)%>"> 刪除</td>
</tr>

<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  GetContent.MoveNext()
Wend
%>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            </tr>
			
        </table>
          <p>&nbsp;</p>
          <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td class="InsideTitle">◎模組類別</td>
            </tr>
            <tr>
              <td><hr size="1" noshade="noshade" />
              <p>‧決策新聞│<a href="../module/news/fileupload.asp">新增</a>│<a href="../module/news/Admin.asp">刪除│修改</a></p>
                <p>‧留言板　│<a href="../module/guestbook/mod_del.asp">回覆│刪除│修改</a></p>
                <p>‧成果成員│<a href="add_student.asp">新增</a>│<a href="list_student.asp">刪除│修改</a></p>
                </td>
            </tr>
          </table>
          <p>&nbsp;</p>
        <table width="760" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><p><span class="smalltext"><img src="../img/w.gif" alt="w" width="15" height="10" /></span><br />
                <a href="../get_content.asp?id=1"></a></p>
				
				<table width="100%" border="0" cellpadding="0" cellspacing="0" id="link">
                <tr>
                  <td><div align="right"><img src="../img/copyright6.gif" alt="CopyRight" width="337" height="37" /><span class="smalltext">
				  <img src="../img/w.gif" alt="w" width="10" height="15" /></span></div></td>
                </tr>
                </table></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
    <td align="left"><img src="../img/right.gif" alt="right" width="5" height="400" /></td>
  </tr>


</table>


</body>
</html>
<%
GetContent.Close()
Set GetContent = Nothing
%>
