<%@ Page Title="" Language="C#" MasterPageFile="~/DalMasterPage.master" AutoEventWireup="true" CodeFile="SiteMap.aspx.cs" Inherits="SiteMap" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
<link href="dal.css" type="text/css" rel="stylesheet" />
<script type="text/javascript">
function changeBgcolor(t) {
 t.style.background="#B0D8FF";
}
function reBgcolor(t) {
 t.style.background="";
}
</script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<table class="tableLine" cellspacing="0" cellpadding="0" width="95%" align="center" border="1">
    <tbody>
        <tr onmouseout="reBgcolor(this)" onmouseover="changeBgcolor(this)">
            <td valign="middle" align="center" width="8%" bgcolor="#cccccc" rowspan="20">
            <p class="style4"><a href="Default.aspx"><font color="#000000">首頁</font></a></p>
            </td>
        </tr>
        <tr>
        </tr>
        <tr onmouseout="reBgcolor(this)" onmouseover="changeBgcolor(this)">
            <td width="23%">
            <div class="style4" align="left">‧<a href="mission-statement.aspx"><font color="#808080">成立宗旨</font></a></div>
            </td>
            <td width="80%">
            <p align="left"><font color="#808080">&nbsp;</font></p>
            </td>
        </tr>
        <%--<tr onmouseout="reBgcolor(this)" onmouseover="changeBgcolor(this)">
            <td>
            <div class="style4" align="left">‧<a href="get_content_news.asp?id=10"><font color="#808080">決策新聞</font></a></div>
            </td>
            <td class="style4">
            <p align="left"><font color="#808080">&nbsp;</font></p>
            </td>
        </tr>--%>
        <tr onmouseout="reBgcolor(this)" onmouseover="changeBgcolor(this)">
            <td>
            <div class="style4" align="left">‧<a href="research.aspx"><font color="#808080">研究成員成果</font></a></div>
            </td>
            <td class="style4">
            <p align="left"><font color="#999999"><font color="#000000">‧</font>老師<br />
            <font color="#000000">‧</font>研究生<br />
            <font color="#000000">‧</font>大學生</font></p>
            </td>
        </tr>
        <tr onmouseout="reBgcolor(this)" onmouseover="changeBgcolor(this)">
            <td>
            <div class="style4" align="left">‧<a href="ResearchPlatforms.aspx"><font color="#808080">研究設備</font></a></div>
            </td>
            <td class="style4">
            <p align="left"><font color="#808080">&nbsp;</font></p>
            </td>
        </tr>
        <tr onmouseout="reBgcolor(this)" onmouseover="changeBgcolor(this)">
            <td>
            <div class="style4" align="left">‧<a href="TeachingSupportSystem.aspx"><font color="#808080">教學支援系統</font></a></div>
            </td>
            <td class="style4">
            <p align="left"><font color="#999999"><font color="#000000">‧</font>決策分析教育諮詢網站 <br />
            <font color="#000000">‧</font>工業工程與管理概論網路教學輔助系統 (郭幸民老師)</font></p>
            </td>
        </tr>
        <tr onmouseout="reBgcolor(this)" onmouseover="changeBgcolor(this)">
            <td>
            <div class="style4" align="left">‧<a href="DecisionSupportSystem.aspx"><font color="#808080">決策支援系統</font></a></div>
            </td>
            <td class="style4">
            <p align="left"><font color="#999999">
            <font color="#000000">‧</font>GDSS (1) 戰機選擇 (Excel)<br />
            <font color="#000000">‧</font>GDSS (2) 小汽車選擇 (Excel+Netmeeting) <br />
            <font color="#000000">‧</font>GDSS (3) 小汽車選擇 (Excel+ASP+Access) <br />
            <font color="#000000">‧</font>GDSS (4) 人力遴選 (ASP+SQL) <br />
            <font color="#000000">‧</font>GDSS (5) 供應商評估 (ASP+SQL) <br />
            <font color="#000000">‧</font>GDSS (6) 決策與磋商支援 (ASP+SQL) <br />
            <font color="#000000">‧</font>NSS (7) 採購協商 (ASP+SQL)</font>
            </p>
            </td>
        </tr>
        <tr onmouseout="reBgcolor(this)" onmouseover="changeBgcolor(this)">
            <td>
            <div class="style4" align="left">‧<a href="CourseSupports.aspx"><font color="#808080">課程支援</font></a></div>
            </td>
            <td class="style4">
            <p align="left"><font color="#999999"><font color="#000000">‧</font>服務管理<br />
            <font color="#000000">‧</font>供應鏈管理<br />
            <font color="#000000">‧</font>決策與系統分析<br />
            <font color="#000000">‧</font>資訊管理 <br />
            <font color="#000000">‧</font>生產與作業管理 <br />
            <font color="#000000">‧</font>多準則決策分析 <br />
            <font color="#000000">‧</font>其它</font></p>
            </td>
        </tr>
        <tr onmouseout="reBgcolor(this)" onmouseover="changeBgcolor(this)">
            <td>
            <div class="style4" align="left">‧<a href="Links.aspx"><font color="#808080">各項連結</font></a></div>
            </td>
            <td class="style4">
            <p align="left"><font color="#808080">&nbsp;</font></p>
            </td>
        </tr>
        <tr onmouseout="reBgcolor(this)" onmouseover="changeBgcolor(this)">
            <td>
            <div class="style4" align="left">‧<a href="SiteMap.aspx"><font color="#808080">網站導覽</font></a></div>
            </td>
            <td class="style4">
            <p align="left"><font color="#808080">&nbsp;</font></p>
            </td>
        </tr>
        <%--<tr onmouseout="reBgcolor(this)" onmouseover="changeBgcolor(this)">
            <td>
            <div class="style4" align="left">‧<a href="module/guestbook/book.asp"><font color="#808080">留言板</font></a></div>
            </td>
            <td class="style4">
            <p align="left"><font color="#808080">&nbsp;</font></p>
            </td>
        </tr>--%>
    </tbody>
</table>
</asp:Content>

