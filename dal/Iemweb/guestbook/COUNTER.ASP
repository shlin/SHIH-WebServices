<%
' 國外免費 WEBHOSTME 的開啟語法如下:
'Set ConnCOUNTER = Server.CreateObject ("ADODB.Connection")
'ConnCOUNTER.Provider = "Microsoft.Jet.OLEDB.4.0"
'ConnCOUNTER.ConnectionString = "Data Source=" & Server.MapPath ("/cgi-bin/BOARD/GBOOK.mdb")
'ConnCOUNTER.Open 


' 計算人氣指數...也可當做計數器使用( 使用 RECORDSET 方式來 UPDATE 人氣指數,因為資料少,所以用此一方式)
 Set ConnCOUNTER = Server.CreateObject("ADODB.Connection")
 DBPath = Server.MapPath("COUNTER.MDB")
 ConnCOUNTER.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

Set cmdTemp = Server.CreateObject("ADODB.Command")
Set CmdEditInfo = Server.CreateObject("ADODB.Recordset")
' 上面這一段可不動...
cmdTemp.CommandText = "SELECT * FROM counter"

cmdTemp.CommandType = 1
Set cmdTemp.ActiveConnection = ConnCOUNTER
CmdEditInfo.Open cmdTemp, , 1, 3

 CMDEDITINFO.MoveFirst

 LASTIP = CMDEDITINFO("LASTIP")
 NEWIP = REQUEST.servervariables("REMOTE_ADDR") 
 COUNTS = CMDEDITINFO("counts")

 IF LASTIP <>  NEWIP THEN
       COUNTS = COUNTS + 1
       CmdEditInfo.Fields("COUNTS") = COUNTS
       CmdEditInfo.Fields("LASTIP") = NEWIP
       CmdEditInfo.Update
   END IF

  CONNCOUNTER.CLOSE
%>

人氣指數：<font color="CC0000"><%=COUNTS%></font>
