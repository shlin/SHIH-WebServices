<%@ Page Title="" Language="C#" MasterPageFile="~/DalMasterPage.master" AutoEventWireup="true" CodeFile="research.aspx.cs" Inherits="research" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <style type="text/css">
        .type {
            float:left;
            width: 150px;
            display: block;
            margin:50px 40px 50px 0;
        }

            .type ul {
                float: left;
                width: 100%;
                margin: 0;
                padding: 0;
                border-bottom: solid;
                border-bottom-width: 2px;
            }

                .type ul li {
                    list-style-position: inside;
                    list-style: none;
                }

                    .type ul li a {
                        padding-left: 25px;
                        padding-right: 25px;
                        padding-top: 5px;
                        padding-bottom: 5px;
                        font-size: large;
                        display: block;
                        text-align: center;
                    }

            .type div {
                width: 100%;
                display: block;
                border-left: solid;
                border-left-width: 10px;
                border-bottom: solid;
                border-bottom-width: 2px;
                text-indent: 5px;
            }

        #color1 * {
            border-color: #0072FF;
        }
        #color2 * {
            border-color: #0bf2f2;
        }
        #color3 * {
            border-color: #1762bf;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <h2>Research Members & Achievement</h2>
    <div class="type" id="color1">
        <div>教師</div><ul>
            <li><a href="http://arsp.most.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do?action=initRsm05&rsNo=5b9110c829e14f1182f68c3fb907ecb0" target="_blank">時序時</a></li>
            <li><a href="http://arsp.most.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do?action=initRsm05&rsNo=9a7fbdc653ac4a158d8b7364db803e0e" target="_blank">徐煥智</a></li>
            <li><a href="http://arsp.most.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do?action=initRsm05&rsNo=9770ce10e54d45f2be5ab7f54036a6f4" target="_blank">周清江</a></li>
            <li><a href="http://arsp.most.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do?action=initRsm05&rsNo=bfffc37e80394fb8ac7b1c07d92dd07b" target="_blank">黃良志</a></li>
            <li><a href="http://arsp.most.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do?action=initRsm05&rsNo=868b8ee27d3c428287a2ed01e4210ae8" target="_blank">鄭啟斌</a></li>
        </ul>
    </div>
    <div class="type" id="color2">
        <div>研究生</div><ul>
            
            <asp:Label ID="Label2" runat="server" Text=""></asp:Label>
            
        </ul>
    </div>
    <div class="type" id="color3">
        <div>大學生</div><ul>


        <asp:Label ID="Label3" runat="server" Text=""></asp:Label>
            

        </ul>
    </div>
</asp:Content>

