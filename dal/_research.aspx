<%@ Page Language="C#" AutoEventWireup="true" CodeFile="_research.aspx.cs" Inherits="research2" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>::: DAL - 決策分析實驗室 :::</title>
    <script type="text/javascript" src="http://code.jquery.com/jquery.min.js"></script>
    <script type="text/javascript">
        var specificCSS;
        var browser = $.browser;

        if (browser.msie) {
            if (browser.version < 10) {
                $('.jquery-shadow-raised').append('<img style="width:100%;height:8px;" src="img/line.gif" />');
            }
        }
    </script>
    <link href="dal.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        body {
            background-attachment: fixed;
            background-image: url(img/body.gif);
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td>&nbsp;</td>
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
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td>&nbsp;</td>
                <td>
                    <div align="right">
                        <span class="smalltext">│Today :
                        <script type="text/javascript">var date = new Date(); document.write(date.getFullYear() + "/" + (date.getMonth() + 1) + "/" + date.getDate())</script>
                        </span>
                    </div>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td>
                    <img src="img/left.gif" alt="left" width="4" height="250" align="right" /></td>
                <td valign="top">
                    <table width="760" border="0" align="center" cellpadding="00" cellspacing="0" class="border" id="leftlink">
                        <tr>
                            <td width="152" valign="middle" background="img/MenuBg.jpg" class="MenuSmall">&nbsp;</td>
                            <td width="608" height="24" colspan="2" align="right" valign="middle" bgcolor="#1ca4c0">
                                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                        <td>
                                            <div align="left" class="MenuSmall">
                                                <img src="img/w.gif" alt="w" width="15" height="15" />::: Every decision you make is a mistake. -- Edward Dahlberg :::
                                            </div>
                                        </td>
                                        <td>
                                            <div align="right">
                                                <span class="MenuSmall">
                                                    <img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Search<img src="img/w.gif" alt="w" width="15" height="15" /><img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Contact Us<img src="img/w.gif" alt="w"
                                                        width="15" height="15" /></span>
                                                <span class="MenuSmall">
                                                    <img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Home<img src="img/w.gif" alt="w" width="15" height="15" /></span>
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" valign="top"><span class="smalltext">
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
                                            <table width="600" height="198" border="0" cellpadding="0" cellspacing="0" class="tb_border" id="link_content">
                                                <tr>
                                                    <td height="100%" align="left" valign="top" background="img/tb_bg6_0.gif">
                                                        <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
                                                        <p align="center">
                                                            &lt;<a href="get_content.asp?id=2">上一頁</a>&gt;
                                                        </p>
                                                        </div></td>
                                                </tr>
                                                <tr>
                                                    <td height="43" background="img/tb_bg6_1.gif">
                                                        <div align="right" class="smalltext">
                                                            <asp:SqlDataSource ID="StudentDatasource" runat="server" ConnectionString="<%$ ConnectionStrings:2007DALConnectionString %>"
                                                                SelectCommand="SELECT * FROM [Student]"></asp:SqlDataSource>
                                                        </div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td valign="top">
                                            <script type="text/javascript"> 
<!--  
    document.write('<div align="center"><font size="1" color="gray"><font face="Verdana, Arial, Helvetica, sans-serif">');
    document.write("<span id='clock'></span>");
    var now, hours, minutes, seconds, timeValue, showtime;
    function showtime() {
        now = new Date();
        hours = now.getHours();
        minutes = now.getMinutes();
        seconds = now.getSeconds();
        timeValue = (hours >= 12) ? "PM " : "AM ";
        timeValue += ((hours > 12) ? hours - 12 : hours) + " 時";
        timeValue += ((minutes < 10) ? " 0" : " ") + minutes + " 分";
        timeValue += ((seconds < 10) ? " 0" : " ") + seconds + " 秒";
        clock.innerHTML = timeValue;
        setTimeout("showtime()", 100);
    }
    showtime();
    document.write('</font></div>');
    //--> 
                                            </script>
                                        </td>
                                        <td align="right" valign="bottom">
                                            <div align="center"></div>
                                            <table width="100%" border="0" cellpadding="0" cellspacing="0" id="link">
                                                <tr>
                                                    <td class="smalltext">
                                                        <p>工業工程推廣影片：<a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext">五分鐘版</a> <a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext"></a>│<a href="http://www.ciie.org.tw/Video/IEM_10.wmv" class="smalltext">十分鐘版</a></p>
                                                    </td>
                                                    <td>
                                                        <div align="right">
                                                            <img src="img/copyright6.gif" alt="CopyRight" width="337" height="37" /><img src="img/w.gif" alt="white" width="15" height="15" />
                                                        </div>
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
    </form>
</body>
</html>
