<% ' 設定 admin.asp 的標題 & 管理者密碼 ,你可以自行修改
       ' 注意: 在此設定的密碼必須為小寫,但於執行 ADMIN.ASP 輸入管理者密碼時,可不分大小寫

         Const title="友緣故事 --- 友言板管理模式"
         Const title2="管理模式"
         ' 管理者密碼...
         Const adminpass = "test"
         Const error = "<center>ㄛ喔，密碼錯誤!! <br>請重新輸入!! 非管理者請離開...</center>" 
         Const retu = "<p align=center><a href=book.asp style=text-decoration: none><strong>返回留言板</font></strong></a><br><br>"

         ' 將密碼轉換為小寫
          PASSWORD =  LCASE(PASSWORD)
%>