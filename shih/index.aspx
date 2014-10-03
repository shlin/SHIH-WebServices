<%@ Page Title="" Language="C#" MasterPageFile="MasterPage.master" AutoEventWireup="true"
    CodeFile="index.aspx.cs" Inherits="index" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <link href="css/frame.css" rel="stylesheet" type="text/css" />
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div id="nowsite">
        時光序曲 &raquo; 首頁</div>
    <center>
        <div>
            <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"
                width="300" height="400" title="時光序曲">
                <param name="movie" value="media/index_show.swf">
                <param name="quality" value="high">
                <embed src="media/index_show.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer"
                    type="application/x-shockwave-flash" width="300" height="400"></embed>
            </object>
        </div>
        <br />
        <marquee behavior="alternate" height="30" scrollamount="3" style="font-size: 14pt;
            font-family: 標楷體; font-weight: bold">利之所在&nbsp; 弊即寓焉&nbsp; 福之所起&nbsp; 禍即伏焉</marquee>
    </center>
</asp:Content>
