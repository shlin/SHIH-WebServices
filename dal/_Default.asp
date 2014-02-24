<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="Connections/Dal.asp" -->
<%
Dim GetContent__MMColParam
GetContent__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  GetContent__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim GetContent
Dim GetContent_cmd
Dim GetContent_numRows

GetContent_cmd = Server.CreateObject ("ADODB.Command")
GetContent_cmd.ActiveConnection = MM_Dal_STRING
GetContent_cmd.CommandText = "SELECT * FROM dbo.Content WHERE id = ?" 
GetContent_cmd.Prepared = true
GetContent_cmd.Parameters.Append GetContent_cmd.CreateParameter("param1", 5, 1, -1, GetContent__MMColParam) ' adDouble

GetContent = GetContent_cmd.Execute
GetContent_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>::: DAL - 決策分析實驗室 :::</title>

<link href="dal.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script><style type="text/css">
<!--
body {
	background-image: url(img/body.gif);
	background-attachment : fixed;
}
.style2 {color: #FFFF00}

-->
</style>

</head>

<body>



<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
    <td><img src="img/banner.gif" alt="banner" width="758" height="115" border="0" usemap="#Map2"  />
      <map name="Map2" id="Map22">
        <area shape="rect" coords="157,61,366,96" href="http://www.ms.tku.edu.tw" alt="淡江大學管理科學研究所" />
        <area shape="rect" coords="410,64,610,97" href="http://www.im.tku.edu.tw" alt="淡江大學資訊管理研究所" />
        <area shape="rect" coords="221,10,543,56" href="index.asp" alt="決策分析研究室" />
        <area shape="circle" coords="65,57,50" href="index.asp" alt="決策分析研究室" />
        <area shape="circle" coords="690,57,51" href="http://www.tku.edu.tw" alt="淡江大學" />
      </map></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="1"><font face="Verdana, Arial, Helvetica, sans-serif"><div align="right" class="smalltext">瀏覽人數 : 
       <!--#include file="module/counter/counter.asp" -->│
	   <%=date%></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right"><img src="img/left.gif" alt="left" width="5" height="400" align="right" /></td>
    <td><table width="760" border="0" align="center" cellpadding="00" cellspacing="0" class="border"   id="leftlink">
      <tr>
        <td width="152" valign="middle" background="img/MenuBg.jpg" class="MenuSmall"><span class="smalltext"><img src="img/w.gif" alt="w" width="15" height="15" /></span></td>
        <td width="608" height="24" colspan="2" align="right" valign="middle" bgcolor="#1ca4c0"><table width="100%" border="0" cellpadding="0" cellspacing="0" >
          <tr>
            <td><div align="left" class="MenuSmall"><img src="img/w.gif" alt="w" width="15" height="15" />::: Every decision you make is a mistake. -- Edward Dahlberg :::</div></td>
           
            <td><div align="right" class="MenuSmall" ><img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Search</div></td>
            <td><div align="right"><span class="MenuSmall"><img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Contact Us</span></div></td>
            <td><div align="right"><span class="MenuSmall" onclick="MM_goToURL('parent','index.asp');return document.MM_returnValue"><img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Home</span></div></td>
            </tr>
        </table></td>
        </tr>
      <tr>
        <td colspan="2" valign="top"><span class="smalltext"><img src="img/w.gif" alt="w" width="15" height="10" /></span></td>
      </tr>
      <tr>
        <td colspan="6" valign="top"><table width="760" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>
			<!--#include file="left_menu.asp" -->
			</td>
            <td><embed src="img/want.swf" width="596" height="288" align="middle" ></embed></td>
          </tr>
          <tr>
            <td rowspan="2" valign="top">
              <table width="150" border="0" cellpadding="0" cellspacing="0" id="link">
              <tr>
                <td style="height: 344px"><p align="left" class="smalltext"><span class="MenuSmall"><img src="img/w.gif" alt="w" width="5" height="20" /></span>:::::::::::::Quick Link:::::::::::::<br />
                    <br />
						
                    <img src="img/w.gif" alt="w" width="5" height="13" />元智工工<br />
                    <a class="smalltext"><img src="img/w.gif" alt="w" width="5" height="13" border="0" /></a><a href="http://decision.iem.yzu.edu.tw/" class="smalltext">決策與最佳化實驗室</a></p>
                    <p align="left" class="smalltext"><img src="img/w.gif" alt="w" width="5" height="15" />清華工工<br />
                        <img src="img/w.gif" alt="w" width="5" height="15" /><a href="http://dalab.ie.nthu.edu.tw/" class="smalltext">決策分析研究室</a><br />
                        <img src="img/w.gif" alt="w" width="5" height="15" /><a href="http://upwen.ie.nthu.edu.tw/internet.htm" class="smalltext">作業研究與管理科學研究室</a><br />
                    </p>
                    
                    <p align="left" class="smalltext"><img src="img/w.gif" alt="w" width="5" height="15" />華梵工管<br />
                        <img src="img/w.gif" alt="w" width="5" height="15" /><a href="http://cat.hfu.edu.tw/~m9823013/jyteng/newweb/myweb/iidex.htm" class="smalltext">運籌管理與決策分析研究室</a><br />
                          <br />
                          <a href="http://www.im.hfu.edu.tw/people/jyteng/newweb/myweb/index.htm" class="text"></a><img src="img/w.gif" alt="w" width="5" height="15" />義守工管<br />
                      <img src="img/w.gif" alt="w" width="5" height="15" /><a href="http://140.127.189.81/dalab/modules/news/" class="smalltext">決策與分析研究室</a></p>
                    <!--<p align="left" class="smalltext"><img src="img/w.gif" alt="w" width="5" height="15" /><a href="http://web.ypu.edu.tw/antai/cids/index.asp" class="smalltext">中華決策科學學會</a></p>-->
                    </td>
              </tr>
            </table></td>
            <td valign="top"><span class="smalltext"><img src="img/w.gif" alt="w" width="15" height="10" /></span><br />
              <table width="600" height="198" border="0" cellpadding="0" cellspacing="0" class="tb_border" id="link_content">
                <tr>
                  <td height="34" align="center" valign="top" background="img/MenuBg3.gif"><p><img src="img/InsideIcon.gif" alt="Icon" width="14" height="14" align="absmiddle" /> <span class="InsideTitle"><%=(GetContent.Fields.Item("name").Value)%></span><br />
                          <br />
                  </p></td>
                </tr>
                <tr>
                  <td height="100%" align="left" valign="middle" background="img/tb_bg6_0.gif"><div align="left"><%=(GetContent.Fields.Item("content").Value)%></div></td>
                </tr>
                <tr>
                  <td height="43" background="img/tb_bg6_1.gif"><div align="right" class="smalltext"><img src="img/w.gif" alt="w" width="15" height="15" /></div></td>
                </tr>
              </table>
              </td>
          </tr>
          <tr>
            <td align="right" valign="bottom"><table width="100%" border="0" cellpadding="0" cellspacing="0" id="link">
              <tr>
                <td class="smalltext"><p>工業工程推廣影片：<a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext">五分鐘版</a> <a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext"></a>
                  │<a href="http://www.ciie.org.tw/Video/IEM_10.wmv" class="smalltext">十分鐘版</a></p></td>
<td><div align="right"><img src="img/copyright6.gif" alt="CopyRight" width="337" height="37" /><span class="smalltext"><img src="img/w.gif" alt="w" width="10" height="15" /></span></div></td>
              </tr>
            </table>              </td>
          </tr>
        </table></td>
      </tr>
    </table></td>
    <td align="left"><img src="img/right.gif" alt="right" width="5" height="400" /></td>
  </tr>


</table>


<div align="left"><a href="lodal.asp"><img src="img/w.gif" width="300" height="10" border="0"></a>
</div>
</body>
</html>
<%
GetContent.Close()
GetContent = Nothing
%>