<%@ Page Title="" Language="C#" MasterPageFile="~/DalMasterPage.master" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <style>
        embed {
            width:596px;
            height:288px;
            margin:0;
            padding:0;
            float:left;
        }
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
    <embed src="img/want.swf" />
    <h2>Mission Statement</h2>
    <ul id="content">
        <li>提供決策分析相關課程學習場所</li>
        <li>支援決策分析有關研究之平台</li>
        <li>彙集決策分析相關文獻及軟體資料庫</li>
        <li>促進決策分析有關知識之推廣與交流</li>
    </ul>
</asp:Content>

