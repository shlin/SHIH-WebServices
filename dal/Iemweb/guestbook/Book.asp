<% ' �}�Ҹ����
' ��~�K�O WEBHOSTME ���}�һy�k�p�U:
'Set Conn = Server.CreateObject ("ADODB.Connection")
'Conn.Provider = "Microsoft.Jet.OLEDB.4.0"
'Conn.ConnectionString = "Data Source=" & Server.MapPath ("/cgi-bin/BOARD/GBOOK.mdb")
'Conn.Open 

' �ꤺ
 Set conn = Server.CreateObject("ADODB.Connection")
 DBPath = Server.MapPath("GBook.mdb")
 conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
SortSql = "Select *From guestbook order by sort"
rs.Open SortSql, conn, 1,1

   ' ���ҭn�s��������
     Page = cint(Request("Page"))

     If Page = 0 Then
        Page = 1
     End If

     RS.PageSize = 20            ' �]�w�C����� 10 ��

     If Not rs.eof Then          ' ����Ƥ~���� 
        RS.AbsolutePage = PAGE   ' �N��ƿ����� PAGE ��
     End if

   '�P�_  admin �O�_���U "���}�޲z" ���s
   '�p�G���h�M�� session("password") �����K�X
    IF request("admin") = "exit" THEN
       Session("password") = False
    END IF
%>
<HTML>
<head>
<title><%=MBTitle%></title>
<body background="�۶��a�O.gif">
<%'Ū��BOOK.ASP�һݪ��]�w��,�Цۦ�ק� CONFIG.INC%>
<!-- #INCLUDE FILE="config.inc" -->

  <tr> 
    <td height="17" width="60"> 
      <div align="left"></div>
    </td>
    <td height="17" width="55" nowrap> 
      <p align="center"><a href=add.asp><img src="./images/n11.gif"  border=0></a><br>
      <font size="2" color="#0000CC"><a href=add.asp>�ڭn�d��</a></font></p>
    </td>
    <td width="55" height="17" nowrap> 
               <p align="center">
              <%IF session("password") = TRUE Then%>
                     <a href=book.asp?admin=exit><img src="./images/admin.gif" border=0></a><br>
                     <font size="2" color="#0000CC"><a href=book.asp?admin=exit>���}�޲z</a></font></p>
              <%Else%>
                     <a href=admin.asp><img src="./images/admin.gif" width="32" height="28" border=0></a><br>
                     <font size="2" color="#0000CC"><a href=admin.asp>�޲z�Ҧ�</a></font></p>
              <%End if%>
    </td>
    <td width="250" height="17" valign="bottom" nowrap> 
      <div align="right">
                 <!-- #INCLUDE FILE="COUNTER.ASP" -->
                 �@�d���� : <font size="2" color="cc0000"><%=rs.recordcount%></font>�@����: <font color="cc0000"><%=CINT(page)%></font> / <font  color="cc0000"><%=rs.PageCount%></font> 
        </font></div>
    </td>
    <td width="92" height="17" valign="bottom"> 
      <div align="right"> 
        <select name="jump" onChange="location.href=this.options[this.selectedIndex].value;" >
              <% ' �ֳt�s�� ���]�w��,�Цۦ�ק� FastLink.inc %>
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
                  <option value=" ">�ֳt����</option>
                   <%FOR J=1 to rs.pagecount
                       IF J <> page THEN
                          IF J < 10  THEN 
                             J = "0" & J
                          END IF%>
                          <option value="book.asp?Page=<%=cint(J)%>">�� <%=J%> ��
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
     RESPONSE.WRITE "      <center><FONT COLOR=maroon>Sorry !! �ثe�|�L�d��...</FONT></center>"
     RESPONSE.WRITE "  </td>"
     RESPONSE.WRITE "</tr>"
Else

  FOR SH=1 to RS.PageSize       ' ��ܸ��

    ' ���o��Ʈw���U��쪺���,��Ʈw�H������W�^���
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

      IF Sex = "�k" Then
         Sexcolor = "red"
      ELSE
         Sexcolor = "BLUE"
      END IF

      IF len(TRIM(Email)) <> 0 Then
         EmailLink = " <A HREF=mailto:"& Email &"><IMG SRC=./images/e_mail.gif BORDER=0 ALT= �q�l�H�c ></A>"
      ELSE
         EmailLink = ""
      End If

      If icq <> "n" and icq <> " " Then
         IcqLink = " <img src=./images/icq.gif width=15 height=16 border=0 alt=" & icq & " >"
      ELSE
         IcqLink = ""
      End If

      If HomePage <> "http://" Then
         HomePageLink = " <A HREF=" & HomePage & " TARGET=_BLANK><IMG SRC=./images/Home-21.gif BORDER=0 ALT=����></A>"
      ELSE
         HomePageLink = ""
      End If

      IF Session("password") = True AND MIMI = "Yes" THEN
         MiMiLink = " [<font color=green>������</font>]"
      ELSE
         MiMiLink = ""
      End If

      IF replyparent = 0 Then   ' replyparent = 0 �N��o�O�Ĥ@�����d��,�Ӥ��O�^�����d��
         barcolor = "#CCCCFF"   ' �H����
         RESPONSE.WRITE "<tr>"
         RESPONSE.WRITE "  <td colspan=9 height=9><div align=center><hr size=1 align=center></div></td>"
         RESPONSE.WRITE "</tr>"
      Else
         barcolor = "ECF5FF"   ' �p�G�O�^�����d��,�h�I�����H�զ�
      End If

         response.write "<tr bgcolor=" & barcolor & ">"
         RESPONSE.WRITE "<td colspan=9>"

      IF REPLYPARENT = 0 THEN        ' �Ĥ@���d��

         RESPONSE.WRITE "<img src=./images/point01.gif> "
         RESPONSE.WRITE "�d����: <font color=blue>" & name & "</font>"

         RESPONSE.WRITE EmailLink:RESPONSE.WRITE IcqLink:RESPONSE.WRITE HomePageLink:RESPONSE.WRITE MiMiLink

         RESPONSE.WRITE "�@[<FONT cOLOR=" & sexcolor & " size=2>" & Sex & "</FONT>] [<font color=blue size=2>" & Job & "</font>]  <FONT FACE=Arial SIZE=1>[" & PostTime & "]</font>"

         ReplyText3 = "[<A HREF=reply.ASP?messageid=" & messageid & ">�^��</A>]"

      ELSE         ' �^�����d��,�ĤG���h

         IF REPLYID = REPLYPARENT THEN
            ReplyImage = "<img src=./images/file.gif> "
            ReplyText1 = ""
            ReplyText2 = " <font color=24567B>�d�U����</font>"
            ReplyText3 = "[<A HREF=reply.ASP?replylevel=2&messageid=" & messageid & ">�^��</A>]"
            ReplyLevel = 2
         ELSE    ' �^�����d��,�ĤT���h�H�W
            ReplyImage = "�@�@<img src=./images/thread.gif> "
            ReplyText1 = "<font color=navy>�d��</font>"
            ReplyText2 = ""
            ReplyText3 = ""
            ReplyEdit = 2       ' �Ω�ק��(�b�d�����r�ꤺ���J��ӥ��Ϊ��ť�)
        '   ReplyText3 = "[<A HREF=reply.ASP?messageid=" & messageid & ">�^��</A>]"
         END IF

        ' ��ܦ^�����d���̸��
          RESPONSE.WRITE ReplyImage & "<FONT COLOR=#740553>" & name & "</FONT> " & ReplyText1 & "�� " & posttime & ReplyText2 & IcqLink & EmailLink & HomePageLink & MiMiLink 

      END IF

    ' �p�G�O�D�޲z�Ҧ�(���}�d��),�h��� [�^��] ���ﶵ
      IF session("password") = False then
         IF MIMI = "No" THEN
            RESPONSE.WRITE "�@" & ReplyText3
         END IF
      ELSE                     ' �p�G�O�޲z�Ҧ�,�h��� �޲z�\�� ���ﶵ
         RESPONSE.WRITE " <FONT SIZE=2 COLOR=CF0336>[<A HREF=reply.ASP?ReplyLevel=" & ReplyLevel & "&messageid=" & messageid & ">�^��</A>| <a href=edit.asp?replyLevel=" & replyedit & "&MESSAGEID=" & MESSAGEID & ">�ק�</A>| <a href=delete.asp?MESSAGEID=" & MESSAGEID & ">�R��</A>]</font>"
      END IF

      IF date() = shorttime Then
         RESPONSE.WRITE " <img src=./images/new2.gif width=28 height=11> "
      END IF

      RESPONSE.WRITE " </td>"
      RESPONSE.WRITE "</tr>"
      RESPONSE.WRITE "<tr>"
      RESPONSE.WRITE "<td colspan=9 height=26>"

    ' ���(���})�d��,�޲z�Ҧ��ɤ]���
      IF MIMI = "No" OR session("password") = True THEN
         RESPONSE.WRITE MEMO
      ELSE   ' �d��������
         RESPONSE.WRITE "<FONT COLOR=GREEN>����</FONT><FONT COLOR=BLUE>�o�O�d���z��������...  ^.<FONT COLOR=RED>*</FONT> ...�O�o�^�ڮ@!!</FONT>"
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