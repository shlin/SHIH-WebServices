<%@  language="VBSCRIPT" codepage="950" %>
<!--#include file="Connections/Dal.asp" -->
<%
Dim GetContent__MMColParam
'決定頁面 預設為1
GetContent__MMColParam = "1"
'當request不為空時 把id後面的數字接起來
If (Request.QueryString("id") <> "") Then 
  GetContent__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim GetContent
Dim GetContent_numRows

Set GetContent = Server.CreateObject("ADODB.Recordset")
GetContent.ActiveConnection = MM_Dal_STRING
GetContent.Source = "SELECT * FROM dbo.Content WHERE id = " + Replace(GetContent__MMColParam, "'", "'")
GetContent.CursorType = 0
GetContent.CursorLocation = 2
GetContent.LockType = 1
GetContent.Open()

GetContent_numRows = 0
%>
<%
Dim gb
Dim gb_cmd
Dim gb_numRows

Set gb_cmd = Server.CreateObject ("ADODB.Command")
gb_cmd.ActiveConnection = MM_Dal_STRING
gb_cmd.CommandText = "SELECT * FROM dbo.GuestBook ORDER BY [Time] DESC" 
gb_cmd.Prepared = true

Set gb = gb_cmd.Execute
gb_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
gb_numRows = gb_numRows + Repeat1__numRows
%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=big5" />
    <title>::: DAL - 決策分析實驗室 :::</title>
    <link href="dal.css" rel="stylesheet" type="text/css" />

    <style type="text/css">
<!--

body {
background-attachment : fixed;
background-image: url(img/body.gif);
}

-->
</style>


</head>
<body onload="MM_preloadImages('img/menu/1-g.gif','img/menu/2-g.gif','img/menu/4-g.gif','img/menu/6-g.gif','img/menu/7-g.gif','img/menu/8-g.gif','img/menu/9-g.gif','img/menu/10-g.gif','img/menu/11-g.gif','img/menu/12-g.gif')">
    <table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td>
                &nbsp;</td>
            <td>
                <img src="img/banner.gif" alt="banner" width="758" height="115" border="0" usemap="#Map2" />
                <map name="Map2">
                    <area shape="rect" coords="157,61,366,96" href="http://www.ms.tku.edu.tw" alt="淡江大學管理科學研究所" />
                    <area shape="rect" coords="410,64,610,97" href="http://www.im.tku.edu.tw" alt="淡江大學資訊管理研究所" />
                    <area shape="rect" coords="221,10,543,56" href="index.asp" alt="決策分析研究室" />
                    <area shape="circle" coords="65,57,50" href="index.asp" alt="決策分析研究室" />
                    <area shape="circle" coords="690,57,51" href="http://www.tku.edu.tw" alt="淡江大學" />
                </map>
            </td>
            <td>
                &nbsp;</td>
        </tr>
        <tr>
            <td>
                &nbsp;</td>
            <td>
                <div align="right">
                    <span class="smalltext">
                        │Today :
                        <%=date%>
                    </span>
                </div>
            </td>
            <td>
                &nbsp;</td>
        </tr>
        <tr>
            <td>
                <img src="img/left.gif" alt="left" width="4" height="250" align="right" /></td>
            <td valign="top">
                <table width="760" border="0" align="center" cellpadding="00" cellspacing="0" class="border"
                    id="leftlink">
                    <tr>
                        <td width="152" valign="middle" background="img/MenuBg.jpg" class="MenuSmall">
                            &nbsp;</td>
                        <td width="608" height="24" colspan="2" align="right" valign="middle" bgcolor="#1ca4c0">
                            <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td>
                                        <div align="left" class="MenuSmall">
                                            <img src="img/w.gif" alt="w" width="15" height="15" />::: Every decision you make
                                            is a mistake. -- Edward Dahlberg :::</div>
                                    </td>
                                    <td>
                                        <div align="right">
                                            <span class="MenuSmall">
                                                <img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Search<img
                                                    src="img/w.gif" alt="w" width="15" height="15" /><img src="img/MenuPic.gif" alt="Menu"
                                                        width="12" height="13" align="absmiddle" />Contact Us<img src="img/w.gif" alt="w"
                                                            width="15" height="15" /></span> <span class="MenuSmall" onclick="MM_goToURL('parent','index.asp');return document.MM_returnValue">
                                                                <img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Home<img
                                                                    src="img/w.gif" alt="w" width="15" height="15" /></span></div>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" valign="top">
                            <span class="smalltext">
                                <img src="img/w.gif" alt="w" width="15" height="10" /></span></td>
                    </tr>
                    <tr>
                        <td colspan="6" valign="top">
                            <table width="760" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td align="left" valign="top">
                                        <!--#include file="left_menu.asp" -->
                                    </td>
                                    <td align="center" valign="top">
                                        <table width="600" height="198" border="0" cellpadding="0" cellspacing="0" class="tb_border"
                                            id="link_content">
                                            <tr>
                                                <td height="34" align="center" valign="top" background="img/MenuBg3.gif">
                                                    <p>
                                                        <img src="img/InsideIcon.gif" alt="Icon" width="14" height="14" align="absmiddle" />
                                                        <span class="InsideTitle">
                                                            <%=(GetContent.Fields.Item("name").Value)%>
                                                        </span>
                                                        <br />
                                                        <br />
                                                    </p>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td height="100%" align="left" valign="top" background="img/tb_bg6_0.gif">
                                                    <div align="left">
                                                        <%=(GetContent("content"))%>
                                                    </div>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td height="43" background="img/tb_bg6_1.gif">
                                                    <div align="right" class="smalltext">
                                                        Update：<%=(GetContent.Fields.Item("date").Value)%><img src="img/w.gif" alt="w" width="15"
                                                            height="15" /></div>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td valign="top">

                                        <script language="JavaScript"> 
<!--  
document.write('<div align="center"><font size="1" color="gray"><font face="Verdana, Arial, Helvetica, sans-serif">'); 
document.write("<span id='clock'></span>"); 
var now,hours,minutes,seconds,timeValue,showtime;
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
                                        </script>

                                    </td>
                                    <td align="right" valign="bottom">
                                        <div align="center">
                                        </div>
                                        <table width="100%" border="0" cellpadding="0" cellspacing="0" id="link">
                                            <tr>
                                                <td class="smalltext">
                                                    <p>
                                                        工業工程推廣影片：<a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext">五分鐘版</a>
                                                        <a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext"></a>│<a href="http://www.ciie.org.tw/Video/IEM_10.wmv"
                                                            class="smalltext">十分鐘版</a></p>
                                                </td>
                                                <td>
                                                    <div align="right">
                                                        <img src="img/copyright6.gif" alt="CopyRight" width="337" height="37" /><img src="img/w.gif"
                                                            alt="white" width="15" height="15" /></div>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
            <td>
                <img src="img/right.gif" alt="right" width="4" height="250" /></td>
        </tr>
    </table>
</body>
</html>
<%
GetContent.Close()
Set GetContent = Nothing
%>
<%
gb.Close()
Set gb = Nothing
%>
