<%@ LANGUAGE="VBSCRIPT"%>
<%Response.Buffer = True%>

<%'取得從 download.asp 按下所要下載的檔名
   fileid  = Request("fileid")

  ' 將 "人氣指數"  存入 "登入者資料" 資料表 內的 "人氣指數" 欄位
    Set conn = Server.CreateObject ("ADODB.Connection")
    DBPath = Server.MapPath("download.mdb")
    conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

    SET rs   = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From download Where (fileid=" & fileid & ")",conn,2,3

  ' 下載人氣指數 + 1 並存入資料庫
	RS("hits") = RS("hits") + 1
	RS.Update

  ' 轉址...從資料庫內取得 USER 按下所要下載的檔案
	Response.Redirect RS("filename")

	rs.Close
	Set rs = Nothing
	Conn.Close
	Set Conn = Nothing
%>