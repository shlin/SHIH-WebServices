<%

CounterAction = "GetStatus/Update" 
'CounterId = The_Id_Of_Your_Counter(Number) 
<!--#include virtual="counter.asp" -->

Dim CounterFile
CounterFile = "module\counter\db\counter.mdb"

Dim CookieSwitch
<!--設定是否開啟 Cookies 記錄功能 on or off-->
CookieSwitch = "on"


Dim CounterDataConn
Dim CounterSQL
Dim CounterRs
Dim CounterHits

Id = Clng(CounterId)

Set CounterDataConn = Server.CreateObject("ADODB.Connection")
CounterDataConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & server.mappath(CounterFile)
CounterSQL = "SELECT * FROM Counter WHERE Id=" & Id
Set CounterRs = CounterDataConn.Execute(CounterSQL)
If CounterRs.EOF then
 CounterSQL = "INSERT INTO Counter(Id,Count) Values(" & Id & ",1)"
 Set CounterRs = CounterDataConn.Execute(CounterSQL)
 CounterHits = 1
Else
 If CounterAction = "GetStatus" then
  CounterHits = Clng(CounterRs("Count"))
 Else
  If NOT Request.Cookies("CounterHits") = "" and Id = 1 then
   CounterHits = Request.Cookies("CounterHits")
  Else
   CounterHits = Clng(CounterRs("Count"))
   CounterHits = CounterHits + 1
   CounterSQL = "UPDATE counter SET count = " & CounterHits & " where id = " & id
   CounterRs = CounterDataConn.Execute(CounterSQl)
   Response.Cookies("CounterHits") = CounterHits
  End If
 End If
End If

Set CounterRs = nothing
CounterDataConn.Close
Set CounterDataConn = nothing

<!--顯示計數器-->
Response.Write CounterHits
%>