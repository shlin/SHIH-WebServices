<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="../../Connections/Dal.asp" -->
<% ClientIP = Request.ServerVariables("REMOTE_ADDR") %>
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_Dal_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.GuestBook (Name, Email, Subject, Memo, [Time], ClientIP) VALUES (?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 255, Request.Form("Name")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 255, Request.Form("Email")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 203, 1, 1073741823, Request.Form("Subject")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 203, 1, 1073741823, Request.Form("Memo")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 135, 1, -1, MM_IIF(Request.Form("Time"), Request.Form("Time"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 50, Request.Form("ClientIP")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "book.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim gb__MMColParam
gb__MMColParam = "1"
If (Request.QueryString("mid") <> "") Then 
  gb__MMColParam = Request.QueryString("mid")
End If
%>
<%
Dim gb
Dim gb_cmd
Dim gb_numRows

Set gb_cmd = Server.CreateObject ("ADODB.Command")
gb_cmd.ActiveConnection = MM_Dal_STRING
gb_cmd.CommandText = "SELECT * FROM dbo.GuestBook WHERE mid = ?" 
gb_cmd.Prepared = true
gb_cmd.Parameters.Append gb_cmd.CreateParameter("param1", 5, 1, -1, gb__MMColParam) ' adDouble

Set gb = gb_cmd.Execute
gb_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>::: DAL - 決策分析實驗室 :::</title>

<link href="../../dal.css" rel="stylesheet" type="text/css" />
<script src="../../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>

<style type="text/css">
<!--

body {
background-attachment : fixed;
background-image: url(../../img/body.gif);



}

-->
</style>

<script type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>


</head>

<body onload="MM_preloadImages('../../img/menu/1-g.gif','../../img/menu/2-g.gif','../../img/menu/3-g.gif','../../img/menu/4-g.gif','../../img/menu/5-g.gif','../../img/menu/6-g.gif','../../img/menu/7-g.gif','../../img/menu/8-g.gif','../../img/menu/9-g.gif','../../img/menu/10-g.gif','../../img/menu/11-g.gif')">



<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
    <td><img src="../../img/banner.gif" alt="banner" width="758" height="115" border="0" usemap="#Map2"  />
      <map name="Map2" id="Map22">
        <area shape="rect" coords="157,61,366,96" href="http://www.ms.tku.edu.tw" alt="淡江大學管理科學研究所" />
        <area shape="rect" coords="410,64,610,97" href="http://www.im.tku.edu.tw" alt="淡江大學資訊管理研究所" />
        <area shape="rect" coords="221,10,543,56" href="../../index.asp" alt="決策分析研究室" />
        <area shape="circle" coords="65,57,50" href="../../index.asp" alt="決策分析研究室" />
        <area shape="circle" coords="690,57,51" href="http://www.tku.edu.tw" alt="淡江大學" />
      </map>
    </td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><div align="right"></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><img src="../../img/left.gif" alt="left" width="4" height="250" align="right" /></td>
    <td valign="top"><table width="760" border="0" align="center" cellpadding="00" cellspacing="0" class="border"   id="leftlink">
      <tr>
        <td width="152" valign="middle" background="../../img/MenuBg.jpg" class="MenuSmall">&nbsp;</td>
        <td width="608" height="24" colspan="2" align="right" valign="middle" bgcolor="#1ca4c0"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div align="left" class="MenuSmall"><img src="../../img/w.gif" alt="w" width="15" height="15" />::: Every decision you make is a mistake. -- Edward Dahlberg :::</div></td>
            <td><div align="right"><span class="MenuSmall"><img src="../../img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Search<img src="../../img/w.gif" alt="w" width="15" height="15" /><img src="../../img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Contact Us<img src="../../img/w.gif" alt="w" 
			width="15" height="15" /></span>
			<span class="MenuSmall" onclick="MM_goToURL('parent','../../index.asp');return document.MM_returnValue"><img src="../../img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Home<img src="../../img/w.gif" alt="w" width="15" height="15" /></span></div></td>
          </tr>
        </table></td>
        </tr>
      <tr>
        <td colspan="2" valign="top"><span class="smalltext"><img src="../../img/w.gif" alt="w" width="15" height="10" /></span></td>
      </tr>
      <tr>
        <td colspan="6" valign="top"><table width="760" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td align="left" valign="top"><table width="100%" border="0" align="left" cellpadding="0" cellspacing="0">
              <tr>
                <td><a href="../../get_content.asp?id=1"> <img src="../../img/menu/1.gif" alt="成立宗旨" name="a" width="140" height="20" border="0" id="1" onmouseover="MM_swapImage('a','','../../img/menu/1-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
              </tr>
              <tr>
                <td><img src="../../img/line.gif" alt="line" width="126" height="8" /></td>
              </tr>
              <tr>
                <td><img src="../../img/menu/2.gif" alt="決策新聞" name="b" width="140" height="20" id="b" onmouseover="MM_swapImage('b','','../../img/menu/2-g.gif',1)" onmouseout="MM_swapImgRestore()" /></td>
              </tr>
              <tr>
                <td><img src="../../img/line.gif" alt="line" width="126" height="8" /></td>
              </tr>
              <tr>
                <td><a href="../../get_content.asp?id=2"><img src="../../img/menu/3.gif" alt="研究成員" name="c" width="140" height="20" border="0" id="c" onmouseover="MM_swapImage('c','','../../img/menu/3-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
              </tr>
              <tr>
                <td><img src="../../img/line.gif" alt="line" width="126" height="8" /></td>
              </tr>
              <tr>
                <td><a href="../../get_content.asp?id=3"><img src="../../img/menu/4.gif" alt="研究設備" name="d" width="140" height="20" border="0" id="d" onmouseover="MM_swapImage('d','','../../img/menu/4-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
              </tr>
              <tr>
                <td><img src="../../img/line.gif" alt="line" width="126" height="8" /></td>
              </tr>
              <tr>
                <td><a href="../../get_content.asp?id=4"><img src="../../img/menu/5.gif" alt="研究成果" name="e" width="140" height="20" border="0" id="e" onmouseover="MM_swapImage('e','','../../img/menu/5-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
              </tr>
              <tr>
                <td><img src="../../img/line.gif" alt="line" width="126" height="8" /></td>
              </tr>
              <tr>
                <td><a href="../../get_content.asp?id=5"><img src="../../img/menu/6.gif" alt="教學支援系統" name="f" width="140" height="20" border="0" id="f" onmouseover="MM_swapImage('f','','../../img/menu/6-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
              </tr>
              <tr>
                <td><img src="../../img/line.gif" alt="line" width="126" height="8" /></td>
              </tr>
              <tr>
                <td><a href="../../get_content.asp?id=6"><img src="../../img/menu/7.gif" alt="決策支援系統" name="g" width="140" height="20" border="0" id="g" onmouseover="MM_swapImage('g','','../../img/menu/7-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
              </tr>
              <tr>
                <td><img src="../../img/line.gif" alt="line" width="126" height="8" /></td>
              </tr>
              <tr>
                <td><a href="../../get_content.asp?id=7"><img src="../../img/menu/8.gif" alt="課程支援" name="h" width="140" height="20" border="0" id="h" onmouseover="MM_swapImage('h','','../../img/menu/8-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
              </tr>
              <tr>
                <td><img src="../../img/line.gif" alt="line" width="126" height="8" /></td>
              </tr>
              <tr>
                <td><a href="../../get_content.asp?id=8"><img src="../../img/menu/9.gif" alt="各項連結" name="i" width="140" height="20" border="0" id="i" onmouseover="MM_swapImage('i','','../../img/menu/9-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
              </tr>
              <tr>
                <td><img src="../../img/line.gif" alt="line" width="126" height="8" /></td>
              </tr>
              <tr>
                <td><img src="../../img/menu/10.gif" alt="網站導覽" name="j1" width="140" height="20" id="j1" onmouseover="MM_swapImage('j1','','../../img/menu/10-g.gif',1)" onmouseout="MM_swapImgRestore()" /></td>
              </tr>
              <tr>
                <td><img src="../../img/line.gif" alt="line" width="126" height="8" /></td>
              </tr>
              <tr>
                <td><a href="book.asp"><img src="../../img/menu/11.gif" alt="留言板" name="j" width="140" height="20" border="0" id="j" onmouseover="MM_swapImage('j','','../../img/menu/11-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
              </tr>
              <tr>
                <td><img src="../../img/line.gif" alt="line" width="126" height="8" /></td>
              </tr>
            </table>
            <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;                </p>
              <p>&nbsp;</p>
              <p>
                <img src="../../img/w.gif" alt="white" width="10" height="5" />
                
                
              </p>
              </td>
            <td align="center" valign="top"><table width="600" height="188" border="0" cellpadding="0" cellspacing="0" class="tb_border" id="link_content">
                <tr>
                  <td valign="bottom" background="../../img/MenuBg3.gif"><div align="center">
                    <p><img src="../../img/InsideIcon.gif" alt="Icon" width="14" height="14" align="absmiddle" /> <span class="InsideTitle">
					GuestBook</span><br />
                      <br />
                      </p>
                    </div>                      </td>
                </tr>
                <tr>
                  <td height="100%" align="left" valign="top" background="../../img/tb_bg6_0.gif"><div align="left"><form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
                    <table width="95%" align="center">
                      <tr valign="baseline">
                        <td width="118" bordercolor="#000000" bordercolorlight="#C0C0C0"><div align="left">姓名│Name</div></td>
                        <td><input name="Name" type="text" class="text" value="" size="32" />                        </td>
                      </tr>
                      <tr valign="baseline">
                        <td bordercolor="#000000" bordercolorlight="#C0C0C0"><div align="left">信箱│E-mail</div></td>
                        <td><input name="Email" type="text" class="text" value="" size="32" />                        </td>
                      </tr>
                      <tr valign="baseline">
                        <td bordercolor="#000000" bordercolorlight="#C0C0C0"><div align="left">主題│Subject</div></td>
                        <td bgcolor='#FFCCFF'><input name="Subject" type="text" class="text" value="Re:<%=(gb.Fields.Item("Subject").Value)%>" size="32" />                        </td>
                      </tr>
                      <tr>
                        <td bordercolor="#000000" bordercolorlight="#C0C0C0"><div align="left">內容│Memo</div></td>
                        <td valign="baseline"><textarea name="Memo" cols="50" rows="5" class="text"></textarea>                        </td>
                      </tr>
                      <tr valign="baseline">
                        <td bordercolor="#000000" bordercolorlight="#C0C0C0"><div align="left">時間│Time</div></td>
                        <td><input name="Time" type="text" class="text" value="<%Response.write DateValue(now) %>" size="32" />                        </td>
                      </tr>
                      <tr valign="baseline">
                        <td bordercolor="#000000" bordercolorlight="#C0C0C0"><div align="left">位址│IP add. </div></td>
                        <td class="smalltext"><%Response.write (ClientIP)%>
                            <input name="ClientIP" type="hidden" value="<%Response.write (ClientIP)%>" />                        </td>
                      </tr>
                      <tr valign="baseline">
                        <td colspan="2" align="right" nowrap="nowrap"><div align="center">
                          <input name="submit" type="submit"  value="留言" />                        
                        </div></td>
                        </tr>
                    </table>
                  
                    <input type="hidden" name="MM_insert" value="form1">
</form>

<%
gb.Close()
Set gb = Nothing
%>
</div>

</td>
                </tr>
                <tr>
                  <td height="43" background="../../img/tb_bg6_1.gif"><div align="right" class="smalltext"><img src="../../img/w.gif" alt="w" width="15" height="15" /></div></td>
                </tr>
                </table>              
              </td>
          </tr>
          <tr>
            <td valign="top">
			
<SCRIPT language="JavaScript"> 
<!--  
document.write('<div align="center"><font size="1" color="gray"><font face="Verdana, Arial, Helvetica, sans-serif">'); 
document.write("<span id='clock'></span>"); 
var now,hours,minutes,seconds,timeValue; 
function showtime(){ 
now = new Date(); 
hours = now.getHours(); 
minutes = now.getMinutes(); 
seconds = now.getSeconds(); 
timeValue = (hours >= 12) ? "PM " : "AM "; 
timeValue += ((hours > 12) ? hours - 12 : hours) + " 時"; 
timeValue += ((minutes < 10) ? " 0" : " ") + minutes + " 分"; 
timeValue += ((seconds < 10) ? " 0" : " ") + seconds + " 秒"; 
clock.innerHTML = timeValue; 
setTimeout("showtime()",100); 
} 
showtime(); 
document.write('</font></div>');
//--> 
</SCRIPT>
			
			<br /></td>
            <td align="right" valign="bottom"><div align="center"></div>
            <table width="100%" border="0" cellpadding="0" cellspacing="0" id="link">
              <tr>
                <td class="smalltext"><p>工業工程推廣影片：<a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext">五分鐘版</a> <a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext"></a> │<a href="http://www.ciie.org.tw/Video/IEM_10.wmv" class="smalltext">十分鐘版</a></p></td>
                <td><div align="right"><img src="../../img/copyright6.gif" alt="CopyRight" width="337" height="37" /><img src="../../img/w.gif" alt="white" width="15" height="15" /></div></td>
              </tr>
            </table></td>
          </tr>
          
        </table></td>
      </tr>
    </table></td>
    <td><img src="../../img/right.gif" alt="right" width="4" height="250" /></td>
  </tr>
  
</table>

</body>
</html>

