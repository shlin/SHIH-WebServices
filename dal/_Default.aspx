<%@ Page Language="C#" AutoEventWireup="true" CodeFile="_Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>::: DAL - 決策分析實驗室 :::</title>
    <link href="dal.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        body {
            background-image: url(img/body.gif);
            background-attachment: fixed;
        }

        .style2 {
            color: #FFFF00;
        }

        li {
            list-style-image: url(img/arrow2.gif);
        }
        #slash {
        }
        .auto-style1 {
            height: 28px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                    <td>&nbsp;</td>
                    <td>
                        <img src="img/banner.gif" alt="banner" width="758" height="115" border="0" usemap="#Map2" />
                        <map name="Map2" id="Map22">
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
                    <td><font size="1"><font face="Verdana, Arial, Helvetica, sans-serif"><div align="right" class="smalltext">瀏覽人數 : </div></td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td align="right">
                        <img src="img/left.gif" alt="left" width="5" height="400" align="right" /></td>
                    <td>
                        <table width="760" border="0" align="center" cellpadding="00" cellspacing="0" class="border" id="leftlink">
                            <tr>
                                <td width="152" valign="middle" background="img/MenuBg.jpg" class="MenuSmall"><span class="smalltext">
                                    <img src="img/w.gif" alt="w" width="15" height="15" /></span></td>
                                <td width="608" height="24" colspan="2" align="right" valign="middle" bgcolor="#1ca4c0">
                                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td>
                                                <div align="left" class="MenuSmall">
                                                    <img src="img/w.gif" alt="w" width="15" height="15" />::: Every decision you make is a mistake. -- Edward Dahlberg :::
                                                </div>
                                            </td>

                                            <td>
                                                <div align="right" class="MenuSmall">
                                                    <img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Search
                                                </div>
                                            </td>
                                            <td>
                                                <div align="right">
                                                    <span class="MenuSmall">
                                                        <img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Contact Us</span>
                                                </div>
                                            </td>
                                            <td>
                                                <div align="right">
                                                    <span class="MenuSmall" onclick="MM_goToURL('parent','index.asp');return document.MM_returnValue">
                                                        <img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Home</span>
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
                                    <table width="760" border="0" cellspacing="0" cellpadding="0">
                                        <tr>
                                            <td>
                                                <table>
                                                    <tr>
                                                        <td class="auto-style1"><a href="get_content.asp?id=1">
                                                            <img src="img/menu/1.gif" alt="成立宗旨" name="a" width="140" height="20" border="0" id="1" onmouseover="MM_swapImage('a','','img/menu/1-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <img src="img/line.gif" alt="line" width="126" height="8" /></td>
                                                    </tr>
                                                    <tr>
                                                        <td><a href="get_content.asp?id=22">
                                                            <img src="img/menu/2.gif" alt="決策藏經閣" name="b" width="140" height="20" border="0" id="b" onmouseover="MM_swapImage('b','','img/menu/2-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <img src="img/line.gif" alt="line" width="126" height="8" /></td>
                                                    </tr>
                                                    <tr>
                                                        <td><a href="get_content.asp?id=2">
                                                            <img src="img/menu/12.gif" alt="研究成員" name="c" width="140" height="20" border="0" id="c" onmouseover="MM_swapImage('c','','img/menu/12-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <img src="img/line.gif" alt="line" width="126" height="8" /></td>
                                                    </tr>
                                                    <tr>
                                                        <td><a href="get_content.asp?id=3">
                                                            <img src="img/menu/4.gif" alt="研究設備" name="d" width="140" height="20" border="0" id="d" onmouseover="MM_swapImage('d','','img/menu/4-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <img src="img/line.gif" alt="line" width="126" height="8" /></td>
                                                    </tr>
                                                    <tr>
                                                        <td><a href="get_content.asp?id=4"></a><a href="get_content.asp?id=5">
                                                            <img src="img/menu/6.gif" alt="教學支援系統" name="f" width="140" height="20" border="0" id="f" onmouseover="MM_swapImage('f','','img/menu/6-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
                                                    </tr>

                                                    <tr>
                                                        <td>
                                                            <img src="img/line.gif" alt="line" width="126" height="8" /></td>
                                                    </tr>
                                                    <tr>
                                                        <td><a href="get_content.asp?id=6">
                                                            <img src="img/menu/7.gif" alt="決策支援系統" name="g" width="140" height="20" border="0" id="g" onmouseover="MM_swapImage('g','','img/menu/7-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <img src="img/line.gif" alt="line" width="126" height="8" /></td>
                                                    </tr>
                                                    <tr>
                                                        <td><a href="get_content.asp?id=7">
                                                            <img src="img/menu/8.gif" alt="課程支援" name="h" width="140" height="20" border="0" id="h" onmouseover="MM_swapImage('h','','img/menu/8-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <img src="img/line.gif" alt="line" width="126" height="8" /></td>
                                                    </tr>
                                                    <tr>
                                                        <td><a href="get_content.asp?id=8">
                                                            <img src="img/menu/9.gif" alt="各項連結" name="i" width="140" height="20" border="0" id="i" onmouseover="MM_swapImage('i','','img/menu/9-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <img src="img/line.gif" alt="line" width="126" height="8" /></td>
                                                    </tr>
                                                    <tr>
                                                        <td><a href="get_content.asp?id=11">
                                                            <img src="img/menu/10.gif" alt="網站導覽" name="j1" width="140" height="20" border="0" id="j1" onmouseover="MM_swapImage('j1','','img/menu/10-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <img src="img/line.gif" alt="line" width="126" height="8" /></td>
                                                    </tr>
                                                    <tr>
                                                        <td><a href="module/guestbook/book.asp">
                                                            <img src="img/menu/11.gif" alt="留言板" name="j" width="140" height="20" border="0" id="j" onmouseover="MM_swapImage('j','','img/menu/11-g.gif',1)" onmouseout="MM_swapImgRestore()" /></a></td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <img src="img/line.gif" alt="line" width="126" height="8" /></td>
                                                    </tr>
                                                </table>
                                            </td>
                                            <td>
                                                <embed src="img/want.swf" width="596" height="288" align="middle"></embed></td>
                                        </tr>
                                        <tr>
                                            <td rowspan="2" valign="top">
                                                <table width="150" border="0" cellpadding="0" cellspacing="0" id="link">
                                                    <tr>
                                                        <td style="height: 344px">
                                                            <p align="left" class="smalltext">
                                                                <span class="MenuSmall">
                                                                    <img src="img/w.gif" alt="w" width="5" height="20" /></span>:::::::::::::Quick Link:::::::::::::<br />
                                                                <br />

                                                                <img src="img/w.gif" alt="w" width="5" height="13" />元智工工<br />
                                                                <a class="smalltext">
                                                                    <img src="img/w.gif" alt="w" width="5" height="13" border="0" /></a><a href="http://decision.iem.yzu.edu.tw/" class="smalltext">決策與最佳化實驗室</a>
                                                            </p>
                                                            <p align="left" class="smalltext">
                                                                <img src="img/w.gif" alt="w" width="5" height="15" />清華工工<br />
                                                                <img src="img/w.gif" alt="w" width="5" height="15" /><a href="http://dalab.ie.nthu.edu.tw/" class="smalltext">決策分析研究室</a><br />
                                                                <img src="img/w.gif" alt="w" width="5" height="15" /><a href="http://upwen.ie.nthu.edu.tw/internet.htm" class="smalltext">作業研究與管理科學研究室</a><br />
                                                            </p>

                                                            <p align="left" class="smalltext">
                                                                <img src="img/w.gif" alt="w" width="5" height="15" />華梵工管<br />
                                                                <img src="img/w.gif" alt="w" width="5" height="15" /><a href="http://cat.hfu.edu.tw/~m9823013/jyteng/newweb/myweb/iidex.htm" class="smalltext">運籌管理與決策分析研究室</a><br />
                                                                <br />
                                                                <a href="http://www.im.hfu.edu.tw/people/jyteng/newweb/myweb/index.htm" class="text"></a>
                                                                <img src="img/w.gif" alt="w" width="5" height="15" />義守工管<br />
                                                                <img src="img/w.gif" alt="w" width="5" height="15" /><a href="http://140.127.189.81/dalab/modules/news/" class="smalltext">決策與分析研究室</a>
                                                            </p>
                                                            <!--<p align="left" class="smalltext"><img src="img/w.gif" alt="w" width="5" height="15" /><a href="http://web.ypu.edu.tw/antai/cids/index.asp" class="smalltext">中華決策科學學會</a></p>-->
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                            <td valign="top"><span class="smalltext">
                                                <img src="img/w.gif" alt="w" width="15" height="10" /></span><br />
                                                <table width="600" height="198" border="0" cellpadding="0" cellspacing="0" class="tb_border" id="link_content">
                                                    <tr>
                                                        <td height="34" align="center" valign="top" background="img/MenuBg3.gif">

                                                            <div>
                                                                <img src="img/InsideIcon.gif" alt="Icon" width="14" height="14" align="absmiddle" />
                                                                Mission Statement
                                                            </div>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td height="100%" align="left" valign="middle" background="img/tb_bg6_0.gif">
                                                            <div align="left">
                                                                <ul>
                                                                    <li>提供決策分析相關課程學習場所
                                                                    </li>
                                                                    <li>支援決策分析有關研究之平台</li>
                                                                    <li>彙集決策分析相關文獻及軟體資料庫</li>
                                                                    <li>促進決策分析有關知識之推廣與交流</li>
                                                                </ul>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td height="43" background="img/tb_bg6_1.gif">
                                                            <div align="right" class="smalltext">
                                                                <img src="img/w.gif" alt="w" width="15" height="15" />
                                                            </div>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="right" valign="bottom">
                                                <table width="100%" border="0" cellpadding="0" cellspacing="0" id="Table1">
                                                    <tr>
                                                        <td class="smalltext">
                                                            <p>
                                                                工業工程推廣影片：<a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext">五分鐘版</a> <a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext"></a>
                                                                │<a href="http://www.ciie.org.tw/Video/IEM_10.wmv" class="smalltext">十分鐘版</a>
                                                            </p>
                                                        </td>
                                                        <td>
                                                            <div align="right">
                                                                <img src="img/copyright6.gif" alt="CopyRight" width="337" height="37" /><span class="smalltext"><img src="img/w.gif" alt="w" width="10" height="15" /></span>
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
                    <td align="left">
                        <img src="img/right.gif" alt="right" width="5" height="400" /></td>
                </tr>


            </table>
        </div>
    </form>
</body>
</html>
