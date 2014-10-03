<%@ Page Title="" Language="C#" MasterPageFile="MasterPage.master" AutoEventWireup="true"
    CodeFile="aboutme.aspx.cs" Inherits="aboutme" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <link href="css/frame.css" rel="stylesheet" type="text/css" />
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div id="nowsite">
        時光序曲 &raquo; 時光序曲</div>
    <center>
        <div id="showalbum-shih">
            <img src="images/Shih.jpg" alt="Shih" /></div>
    </center>
    <h3>
        個人興趣</h3>
    旅遊、慢跑、攝影、音樂<br />
    <br />
    <h3>
        主要學歷</h3>
    <div align="center">
        <table border="1" cellspacing="0" cellpadding="0" id="table2">
            <thead>
                <tr align="center">
                    <td>
                        畢／肄業學校
                    </td>
                    <td>
                        國別
                    </td>
                    <td width="200">
                        主修學門系所
                    </td>
                    <td width="45">
                        學位
                    </td>
                    <td width="185">
                        起訖年月(<u>西元年</u>/<u>月</u>)
                    </td>
                </tr>
            </thead>
            <tbody id="eduTable" runat="server">
            </tbody>
        </table>
    </div>
    <br />
    <h3>
        現職與專長相關之經歷</h3>
    <div align="center">
        <table cellspacing="0" cellpadding="0" border="1">
            <thead>
                <tr align="center">
                    <td width="200">
                        服務機關
                    </td>
                    <td width="200">
                        服務部門／系所
                    </td>
                    <td width="75">
                        職稱
                    </td>
                    <td width="185">
                        起訖年月(<u>西元年</u>/<u>月</u>)
                    </td>
                </tr>
            </thead>
            <tbody id="expTable" runat="server">
                <tr align="center" runat="server">
                    <td colspan="4">
                        現職
                    </td>
                </tr>
                <tr align="center" runat="server">
                    <td colspan="4">
                        經歷
                    </td>
                </tr>
            </tbody>
        </table>
    </div>
    <br />
    <h3>
        服務</h3>
    <ol type="1" id="serviceList" runat="server">
    </ol>
    <h3>
        榮譽</h3>
    <ol type="1" id="honorList" runat="server">
    </ol>
</asp:Content>
