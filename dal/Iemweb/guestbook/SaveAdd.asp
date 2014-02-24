<%
   ' 取得各欄位的資料
     Name = Request.Form("Name")
     Sex = Request.Form("Sex")
     Job = Request.Form("Job")
     Email = Request.Form("Email")
     HomePage = Request.Form("HomePage")
     messagebody = Request.Form("messagebody")
     icq = Request.Form("icq")
     security = Request.Form("security")

    ' 如果 HTML 語法關閉 則使用此敘述
      messagebody = Replace(messagebody,"<" , "&lt;")

      messagebody = Replace(messagebody,vbCrLf , "<br>")

    ' 於執行斷行時,<br>一樣會被加一,算入是否大於 80 字元,為了不加一,所以用`字,先代替<br>
      messagebody = Replace(messagebody, "<br>" ,"`")

    ' textarea 自動斷行的判斷仍未研究出來,所以只好用程式碼 Parse_text 來指行斷行處理,
      parse_text(messagebody)

      messagebody = Replace(messagebody ,"`", "<br>" )
      messagebody = Replace(messagebody,vbCrLf,"<br>")

    ' 如果沒填 e-mail
      If Email="" Then
         Email=" "
      End If

    ' 如果沒填 icq
      If icq="" Then
         icq=" "
      End If

      Homepage = trim(HomePage)
    ' 如果沒填 HTTp
      If len(trim(HomePage))  = 0 or mid(HomePage,1,4) <> "http" Then
         HomePage ="http://" & Homepage
      End If

    ' user data write to bowser...Cookies
      Domain = Request.ServerVariables("SERVER_NAME")

      Response.Cookies("Name") = Name
      Response.Cookies("Name").Domain = Domain
      Response.Cookies("Name").Expires = date()+30

      Response.Cookies("sex") = sex
      Response.Cookies("sex").Domain = Domain
      Response.Cookies("sex").Expires = date()+30

      Response.Cookies("Email") = Email
      Response.Cookies("Email").Domain = Domain
      Response.Cookies("Email").Expires = date()+30

      Response.Cookies("HomePage") = HomePage
      Response.Cookies("HomePage").Domain = Domain
      Response.Cookies("HomePage").Expires = date()+30

      Response.Cookies("icq") = icq
      Response.Cookies("icq").Domain = Domain
      Response.Cookies("icq").Expires = date()+30

' 國外免費 WEBHOSTME 的開啟語法如下:
'Set Conn = Server.CreateObject ("ADODB.Connection")
'Conn.Provider = "Microsoft.Jet.OLEDB.4.0"
'Conn.ConnectionString = "Data Source=" & Server.MapPath ("/cgi-bin/BOARD/GBOOK.mdb")
'Conn.Open 


    ' 與資料庫連線
      Set conn = Server.CreateObject("ADODB.Connection")
      DBPath = Server.MapPath("GBook.mdb")
      conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
     Set RS = Server.CreateObject("ADODB.Recordset")
     RS.Open "select * from guestbook ORDER BY MESSAGEID DESC", conn,1,3

    ' 因為想直接由 SORT 欄位來排序整個張貼者與回應者的順序,但又有許多問題,
    ' 像回應原張貼者的時候我想直接就將回應者的順序,排序在他的下方,這樣就不
    ' 必像 討論板 一樣須透過開啟二次資料庫使用迴圈來做資料的取得與排序,那種速度
    ' 會慢很多,所以乾脆就在資料寫入時就排序...且新張貼者的資料在最上方
    ' 但 COUNT 起始值我是設成從 999999 開啟算起,所以當你的站有那麼多人在留言時,
    ' 最好能將 999999 改大一點,多大都沒關係...
    ' 記得 SAVEREPLY.ASP 內的 999999 也要一起改哦!!
    ' 注意: 當 count = 0 時,排序問題將出現錯誤,因為我想不出其他方式來設定 count 值了...
    ' 如果有人想出來了,麻煩告知...thanks.

   IF RS.RECORDCOUNT <= 0 THEN    ' 因為須將新張貼者的資料排在最上方,所以我沒使用自動編號的功能
      count = 999999
   ELSE
      rs.movelast
      count = rs("messageid") - 1   ' 抓取messageid 最後的值來減 1
   END IF

   ' 新增一筆資料錄,並寫入資料庫
   rs.AddNew
   rs("messageid") = count
   rs("replyid") = count
   rs("Sort") = count
   rs("Name") = Name
   rs("Sex") = Sex
   rs("Job") = Job
   rs("Email") = Email
   rs("HomePage") = HomePage
   rs("messagebody") = messagebody
   rs("icq") = icq
   rs("security") = security
   rs.update
   rs.close
   SET CONN = NOTHING

   Response.Redirect "book.asp" ' 返回『友言板』


Function Parse_text(m_output)
  '上一次位置、目前累計長度、實際位置、總長度
  Dim old_pos, tmp_Len, i, tot_Len,temp_Memo,temp2_Memo
  old_pos = 0   '上一次位置
  tmp_Len = 0   '記錄邏輯上的長度
  TOT_LEN = Len(Trim(m_output))
  temp_Memo = ""
  temp2_Memo = "" 

 ' i 代表字串中實際的位置
  For i = 1 To tot_Len

      If Len(Hex(Asc(Mid(m_output, i, 1)))) > 2 and Mid(m_output,i,1) <> "`" Then
         '將長度加2
         tmp_Len = tmp_Len  + 2
      ElseIf Mid(m_output,i,1) = "`" then
         temp_Memo =  Mid(m_output,old_pos+1,i-old_pos)
         '還原計數值,因為 <BR> 也會被算入,所以將<BR>將換為"`",以避免被加一
         old_pos = i
         tmp_Len = 0
      ELSE
         '否則加1
         tmp_Len = tmp_Len + 1
      End If

      '如果大於等於 80 字元，準備輸出結果
      If tmp_Len >=  80  Then
         temp_Memo =  Mid(m_output,old_pos+1,i-old_pos) & "<br>"
        ' 還原計數值
         old_pos = i
         tmp_Len = 0
      End If
         temp2_Memo = temp2_Memo & temp_Memo
         temp_Memo = ""
  Next

  '將其餘的字元(如果還有的話)一併輸出
  If tmp_Len > 0 Then
     messagebody = temp2_Memo & Mid(m_output,old_pos+1,tot_Len-old_pos)
     messagebody = Replace(messagebody ,"`", "<br>" )
  End If
End Function
%>