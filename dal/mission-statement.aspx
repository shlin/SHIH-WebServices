<%@ Page Title="" Language="C#" MasterPageFile="~/DalMasterPage.master" AutoEventWireup="true" CodeFile="mission-statement.aspx.cs" Inherits="mission_statement" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <style>
        #content {
            float:left;
            margin:0;
            margin-top:25px;
            margin-bottom:25px;
        }
        #content li {
            list-style-position:inside;
            list-style-image:url(img/arrow2.gif);
            margin-top:10px;
            margin-bottom:10px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<%--    <h2>Mission Statement</h2>
    <ul id="content">
        <li>提供決策分析相關課程學習場所</li>
        <li>支援決策分析有關研究之平台</li>
        <li>彙集決策分析相關文獻及軟體資料庫</li>
        <li>促進決策分析有關知識之推廣與交流</li>
    </ul>--%>
    <table cellspacing="0" cellpadding="0" width="98%" align="center" border="0">
    <tbody>
        <tr>
            <td valign="top" width="3%"><span class="smalltext"><img height="15" alt="w" width="15" src="img/w.gif" /></span></td>
            <td width="97%"><img height="14" alt="arrow" width="13" src="img/arrow2.gif" /><span class="smalltext"><img height="15" alt="w" width="15" src="img/w.gif" /></span>提供決策分析相關課程學習場所<img height="30" alt="w" width="15" src="img/w.gif" /></td>
        </tr>
        <tr>
            <td valign="bottom">&nbsp;</td>
            <td><img height="14" alt="arrow" width="13" src="img/arrow2.gif" /><span class="smalltext"><img height="15" alt="w" width="15" src="img/w.gif" /></span>支援決策分析有關研究之平台<span class="smalltext"><img height="30" alt="w" width="15" src="img/w.gif" /></span></td>
        </tr>
        <tr>
            <td valign="bottom">&nbsp;</td>
            <td><img height="14" alt="arrow" width="13" src="img/arrow2.gif" /><span class="smalltext"><img height="15" alt="w" width="15" src="img/w.gif" /></span>彙集決策分析相關文獻及軟體資料庫<span class="smalltext"><img height="30" alt="w" width="15" src="img/w.gif" /></span></td>
        </tr>
        <tr>
            <td valign="bottom">&nbsp;</td>
            <td><img height="14" alt="arrow" width="13" src="img/arrow2.gif" /><span class="smalltext"><img height="15" alt="w" width="15" src="img/w.gif" /></span>促進決策分析有關知識之推廣與交流 <span class="smalltext"><img height="15" alt="w" width="15" src="img/w.gif" /><img height="30" alt="w" width="15" src="img/w.gif" /></span></td>
        </tr>
        <tr>
            <td valign="bottom">&nbsp;</td>
            <td>&nbsp;</td>
        </tr>
    </tbody>
</table>
</asp:Content>

