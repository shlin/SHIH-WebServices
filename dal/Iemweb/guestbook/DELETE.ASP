<%
   MESSAGEID = Request("MESSAGEID")

' 國外免費 WEBHOSTME 的開啟語法如下:
'Set Conn = Server.CreateObject ("ADODB.Connection")
'Conn.Provider = "Microsoft.Jet.OLEDB.4.0"
'Conn.ConnectionString = "Data Source=" & Server.MapPath ("/cgi-bin/BOARD/GBOOK.mdb")
'Conn.Open 

' 國內中文 IIS
   Set conn = Server.CreateObject("ADODB.Connection")
   DBPath = Server.MapPath("GBOOK.mdb")
   conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

        ' 刪除資料

          SET rs = Server.CreateObject("ADODB.Recordset")
  '       SET rs2 = Server.CreateObject("ADODB.Recordset")
          rs.Open "Select * From GUESTBOOK WHERE (MESSAGEID= " & MESSAGEID & ")", conn,1,3

          SORT = RS("SORT")

          IF RS("REPLYPARENT") <> 0 THEN
             MESSAGEID = RS("REPLYPARENT")

             SET rs2 = Server.CreateObject("ADODB.Recordset")
             rs2.Open "Select * From GUESTBOOK WHERE (MESSAGEID= " & MESSAGEID & ")", conn,1,3
             rs2("REPLYTOTAL") = rs2("REPLYTOTAL") - 1
             rs2.update
             rs2.CLOSE
          END IF

          CONN.EXECUTE "DELETE From GUESTBOOK WHERE SORT LIKE '%" & SORT & "%'"
 
   SET CONN = NOTHING

Response.Redirect ("BOOK.asp")%>