<%@ Page Title="" Language="C#" MasterPageFile="~/DalMasterPage.master" AutoEventWireup="true" CodeFile="achievement.aspx.cs" Inherits="achievement" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <link href="dal.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="http://code.jquery.com/jquery-1.8.3.min.js"></script>
    <script type="text/javascript">
        var specificCSS;
        var browser = $.browser;

        if (browser.msie) {
            if (browser.version < 10) {
                $('.jquery-shadow-raised').append('<img style="width:100%;height:8px;" src="img/line.gif" />');
            }
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:Label ID="Label2" runat="server"></asp:Label>
</asp:Content>

