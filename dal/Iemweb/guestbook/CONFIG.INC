<%
' 設定 book.asp 內的變數
Const readme1 = "注意：<FONT COLOR=RED>目前為管理模式,如要離開管理模式,請按 離開管理 按鈕</FONT>　　　"
Const readme2 = "說明 : <font color=red>請勿填寫不雅文字 或 有關人身攻擊言論　<< <font color=green>HTML 關閉</font> >>　　　</font>"
Const background = "marge.jpg"

' 設定連結時的顏色 %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
 <STYLE TYPE="text/css">
        <!-- body,th,td,input,select,textarea,select,checkbox {font:9pt 新細明體}
         a:link {font: 9pt "新細明"; text-decoration: none;color:none}
         a:visited      {font: 9pt "新細明"; text-decoration: none; color: 000099}
         a:active       {font: 9pt "新細明"; text-decoration: none; color: 00CC66}
         a:hover        {font: 9pt 新細明; text-decoration: underline; color: ff0000}
        -->
   </STYLE>
</head>

<!------------------此處為本站的 logo & 走馬燈 -------------------
<div id="Layer23" style="position:absolute; left:298px; top:-20px; width:177px; height:49px; z-index:51; overflow: visible"> 
  <marquee height=100 behavior=loop direction=up scrolldelay=4 scrollamount=1> 
  <table width="100%" border="0" height="20">
    <tr> 
      <td height="30"> 
        <div align="center">
         <FONT COLOR=RED><B>本站 友言板 ASP 程式</B></FONT><br><br>
         免費提供網友索取<br><br>
         有興趣者,請留言或 E-mail<br><br>
         並註明<br><br>
         <font color=blue>姓名 E-mail 網址 索取原因</font>
      </td>
    </tr>
  </table>
  </marquee> </div>
-- ******************** 友言板 LOGO ******************************!-->
<div id="Layer1" style="position:absolute; width:147px; height:40px; z-index:52; left: 313px; top: 2px"><img src="./images/m-board.gif" width="100" height="40"> 
  <img src="./images/e9.gif" width="41" height="50"></div>

<table border="0" align="center" width=70%>
  <tr> 
    <td colspan="9" height="13">
      <p>&nbsp;</p>
      <p>&nbsp;</p>
    </td>
  </tr>
    <td colspan="9"> 
      <div align="center"><img src=./images/house.gif width="512" height="15"></div>
    </td>
  </tr>