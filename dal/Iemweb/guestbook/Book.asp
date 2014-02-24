<% ' 開啟資料檔
' 國外免費 WEBHOSTME 的開啟語法如下:
'Set Conn = Server.CreateObject ("ADODB.Connection")
'Conn.Provider = "Microsoft.Jet.OLEDB.4.0"
'Conn.ConnectionString = "Data Source=" & Server.MapPath ("/cgi-bin/BOARD/GBOOK.mdb")
'Conn.Open 

' 國內
 Set conn = Server.CreateObject("ADODB.Connection")
 DBPath = Server.MapPath("GBook.mdb")
 conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
SortSql = "Select *From guestbook order by sort"
rs.Open SortSql, conn, 1,1

   ' 取所要瀏覽的頁次
     Page = cint(Request("Page"))

     If Page = 0 Then
        Page = 1
     End If

     RS.PageSize = 20            ' 設定每頁顯示 10 筆

     If Not rs.eof Then          ' 有資料才執行 
        RS.AbsolutePage = PAGE   ' 將資料錄移至 PAGE 頁
     End if

   '判斷  admin 是否按下 "離開管理" 按鈕
   '如果有則清除 session("password") 內的密碼
    IF request("admin") = "exit" THEN
       Session("password") = False
    END IF
%>
<HTML>
<head>
<title><%=MBTitle%></title>
<body background="石塊地板.gif">
<%'讀取BOOK.ASP所需的設定值,請自行修改 CONFIG.INC%>
<!-- #INCLUDE FILE="config.inc" -->

  <tr> 
    <td height="17" width="60"> 
      <div align="left"></div>
    </td>
    <td height="17" width="55" nowrap> 
      <p align="center"><a href=add.asp><img src="./images/n11.gif"  border=0></a><br>
      <font size="2" color="#0000CC"><a href=add.asp>我要留言</a></font></p>
    </td>
    <td width="55" height="17" nowrap> 
               <p align="center">
              <%IF session("password") = TRUE Then%>
                     <a href=book.asp?admin=exit><img src="./images/admin.gif" border=0></a><br>
                     <font size="2" color="#0000CC"><a href=book.asp?admin=exit>離開管理</a></font></p>
              <%Else%>
                     <a href=admin.asp><img src="./images/admin.gif" width="32" height="28" border=0></a><br>
                     <font size="2" color="#0000CC"><a href=admin.asp>管理模式</a></font></p>
              <%End if%>
    </td>
    <td width="250" height="17" valign="bottom" nowrap> 
      <div align="right">
                 <!-- #INCLUDE FILE="COUNTER.ASP" -->
                 　留言數 : <font size="2" color="cc0000"><%=rs.recordcount%></font>　頁數: <font color="cc0000"><%=CINT(page)%></font> / <font  color="cc0000"><%=rs.PageCount%></font> 
        </font></div>
    </td>
    <td width="92" height="17" valign="bottom"> 
      <div align="right"> 
        <select name="jump" onChange="location.href=this.options[this.selectedIndex].value;" >
              <% ' 快速連結 的設定值,請自行修改 FastLink.inc %>
              <!-- #INCLUDE FILE="fastlink.inc" -->
        </select>
      </div>
    </td>
  </tr>
  <tr valign="top"> 
    <td colspan="9"> 
      <div align="center"><img src="./images/house.gif" width="512" height="15"></div>
    </td>
  </tr>
<tr>
  <td colspan=9 height=9>
     <div align=right>
         <%IF session("password") = FALSE THEN
              RESPONSE.WRITE README2
           ELSE
              RESPONSE.WRITE README1
           END IF%>
              <select name="jump" onChange="location.href=this.options[this.selectedIndex].value;" >
                  <option value=" ">快速跳頁</option>
                   <%FOR J=1 to rs.pagecount
                       IF J <> page THEN
                          IF J < 10  THEN 
                             J = "0" & J
                          END IF%>
                          <option value="book.asp?Page=<%=cint(J)%>">第 <%=J%> 頁
                     <%END IF
                     NEXT%>
           </select>
        </div>
    </td>
</tr>
<%
If rs.Eof Then
     RESPONSE.WRITE "  <td colspan=9 height=9><div align=center><hr size=1 align=center></div></td>"
     RESPONSE.WRITE "<tr>"
     RESPONSE.WRITE "  <td colspan=9 height=9>"
     RESPONSE.WRITE "      <center><FONT COLOR=maroon>Sorry !! 目前尚無留言...</FONT></center>"
     RESPONSE.WRITE "  </td>"
     RESPONSE.WRITE "</tr>"
Else

  FOR SH=1 to RS.PageSize       ' 顯示資料

    ' 取得資料庫內各欄位的資料,資料庫以全部欄名英文化
      MESSAGEID = rs("messageid")
      REPLYID = RS("REPLYID")
      REPLYPARENT = RS("REPLYPARENT")
      posttime = rs("posttime")
      shorttime = rs("shorttime")
      Email = rs("email")
      HomePage = rs("homepage")
      Name = rs("name")
      Sex = rs("sex")
      Job = rs("job")
      Memo = rs("messagebody")
      icq = rs("icq")
      mimi = rs("security")

      IF Sex = "女" Then
         Sexcolor = "red"
      ELSE
         Sexcolor = "BLUE"
      END IF

      IF len(TRIM(Email)) <> 0 Then
         EmailLink = " <A HREF=mailto:"& Email &"><IMG SRC=./images/e_mail.gif BORDER=0 ALT= 電子信箱 ></A>"
      ELSE
         EmailLink = ""
      End If

      If icq <> "n" and icq <> " " Then
         IcqLink = " <img src=./images/icq.gif width=15 height=16 border=0 alt=" & icq & " >"
      ELSE
         IcqLink = ""
      End If

      If HomePage <> "http://" Then
         HomePageLink = " <A HREF=" & HomePage & " TARGET=_BLANK><IMG SRC=./images/Home-21.gif BORDER=0 ALT=首頁></A>"
      ELSE
         HomePageLink = ""
      End If

      IF Session("password") = True AND MIMI = "Yes" THEN
         MiMiLink = " [<font color=green>悄悄話</font>]"
      ELSE
         MiMiLink = ""
      End If

      IF replyparent = 0 Then   ' replyparent = 0 代表這是第一次的留言,而不是回應的留言
         barcolor = "#CCCCFF"   ' 淡紫色
         RESPONSE.WRITE "<tr>"
         RESPONSE.WRITE "  <td colspan=9 height=9><div align=center><hr size=1 align=center></div></td>"
         RESPONSE.WRITE "</tr>"
      Else
         barcolor = "ECF5FF"   ' 如果是回應的留言,則背景為淡白色
      End If

         response.write "<tr bgcolor=" & barcolor & ">"
         RESPONSE.WRITE "<td colspan=9>"

      IF REPLYPARENT = 0 THEN        ' 第一次留言

         RESPONSE.WRITE "<img src=./images/point01.gif> "
         RESPONSE.WRITE "留言者: <font color=blue>" & name & "</font>"

         RESPONSE.WRITE EmailLink:RESPONSE.WRITE IcqLink:RESPONSE.WRITE HomePageLink:RESPONSE.WRITE MiMiLink

         RESPONSE.WRITE "　[<FONT cOLOR=" & sexcolor & " size=2>" & Sex & "</FONT>] [<font color=blue size=2>" & Job & "</font>]  <FONT FACE=Arial SIZE=1>[" & PostTime & "]</font>"

         ReplyText3 = "[<A HREF=reply.ASP?messageid=" & messageid & ">回應</A>]"

      ELSE         ' 回應的留言,第二階層

         IF REPLYID = REPLYPARENT THEN
            ReplyImage = "<img src=./images/file.gif> "
            ReplyText1 = ""
            ReplyText2 = " <font color=24567B>留下此言</font>"
            ReplyText3 = "[<A HREF=reply.ASP?replylevel=2&messageid=" & messageid & ">回應</A>]"
            ReplyLevel = 2
         ELSE    ' 回應的留言,第三階層以上
            ReplyImage = "　　<img src=./images/thread.gif> "
            ReplyText1 = "<font color=navy>留言</font>"
            ReplyText2 = ""
            ReplyText3 = ""
            ReplyEdit = 2       ' 用於修改時(在留言的字串內插入兩個全形的空白)
        '   ReplyText3 = "[<A HREF=reply.ASP?messageid=" & messageid & ">回應</A>]"
         END IF

        ' 顯示回應的留言者資料
          RESPONSE.WRITE ReplyImage & "<FONT COLOR=#740553>" & name & "</FONT> " & ReplyText1 & "於 " & posttime & ReplyText2 & IcqLink & EmailLink & HomePageLink & MiMiLink 

      END IF

    ' 如果是非管理模式(公開留言),則顯示 [回應] 的選項
      IF session("password") = False then
         IF MIMI = "No" THEN
            RESPONSE.WRITE "　" & ReplyText3
         END IF
      ELSE                     ' 如果是管理模式,則顯示 管理功能 的選項
         RESPONSE.WRITE " <FONT SIZE=2 COLOR=CF0336>[<A HREF=reply.ASP?ReplyLevel=" & ReplyLevel & "&messageid=" & messageid & ">回應</A>| <a href=edit.asp?replyLevel=" & replyedit & "&MESSAGEID=" & MESSAGEID & ">修改</A>| <a href=delete.asp?MESSAGEID=" & MESSAGEID & ">刪除</A>]</font>"
      END IF

      IF date() = shorttime Then
         RESPONSE.WRITE " <img src=./images/new2.gif width=28 height=11> "
      END IF

      RESPONSE.WRITE " </td>"
      RESPONSE.WRITE "</tr>"
      RESPONSE.WRITE "<tr>"
      RESPONSE.WRITE "<td colspan=9 height=26>"

    ' 顯示(公開)留言,管理模式時也顯示
      IF MIMI = "No" OR session("password") = True THEN
         RESPONSE.WRITE MEMO
      ELSE   ' 留言給站長
         RESPONSE.WRITE "<FONT COLOR=GREEN>站長</FONT><FONT COLOR=BLUE>這是留給您的悄悄話...  ^.<FONT COLOR=RED>*</FONT> ...記得回我哦!!</FONT>"
      END IF

      RESPONSE.WRITE " </td>"
      RESPONSE.WRITE "</tr>"

   rs.MoveNext
   IF RS.EOF THEN EXIT FOR
   Next
End If
rs.close
Set Conn = nothing
%>
     <tr>
        <td colspan=9 height=9><div align=center><hr size=1 align=center></div></td>
     </tr>
  <tr> 
    <td colspan="9"> 
      <div align="center">
         <!-- #INCLUDE FILE="footer.inc" -->
     </div>
    </td>
  </tr>
</table>
</body>
</html>