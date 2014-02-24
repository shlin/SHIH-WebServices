<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="../Connections/Dal.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_Dal_STRING
    MM_editCmd.CommandText = "DELETE FROM dbo.Student WHERE MemId = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "admin.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If

End If
%>
<%
Dim studentDB__MMColParam
studentDB__MMColParam = "1"
If (Request.QueryString("MemId") <> "") Then 
  studentDB__MMColParam = Request.QueryString("MemId")
End If
%>
<!--#include file="fckeditor.asp" -->
<%
Dim studentDB
Dim studentDB_cmd
Dim studentDB_numRows

Set studentDB_cmd = Server.CreateObject ("ADODB.Command")
studentDB_cmd.ActiveConnection = MM_Dal_STRING
studentDB_cmd.CommandText = "SELECT * FROM dbo.Student WHERE MemId = ?" 
studentDB_cmd.Prepared = true
studentDB_cmd.Parameters.Append studentDB_cmd.CreateParameter("param1", 5, 1, -1, studentDB__MMColParam) ' adDouble

Set studentDB = studentDB_cmd.Execute
studentDB_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>刪除</title>
<style type="text/css">
<!--
body {
	background-image: url(../../img/body.gif);
}
-->
</style></head>

<body>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
  <table width="800" border="1" align="center">
    <tr valign="baseline">
      <td width="74" align="right" valign="middle" nowrap><div align="center">題目</div></td>
      <td><input type="text" name="Topic" value="<%=(studentDB.Fields.Item("Topic").Value)%>" size="80">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">年份</div></td>
      <td><input type="text" name="year" value="<%=(studentDB.Fields.Item("year").Value)%>" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">姓名</div></td>
      <td><input name="name" type="text" id="name" value="<%=(studentDB.Fields.Item("name").Value)%>" size="80" />
      請以&quot;、&quot;頓號作為分隔</td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">級別</div></td>
      <td><%=(studentDB.Fields.Item("class").Value)%></td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">類別</div></td>
      <td><%=(studentDB.Fields.Item("type").Value)%></td>
    </tr>
    <tr>
      <td nowrap align="right" valign="middle"><p align="center">內容</p>      </td>
      <td width="714" valign="baseline"><%=(studentDB.Fields.Item("content").Value)%></td>
    </tr>
    <tr valign="baseline">
      <td colspan="2" align="right" nowrap><div align="center">　</div>        <div align="center">
          <input name="Del" type="submit" id="Del" value="刪除記錄">      
        </div></td>
    </tr>
  </table>
  

  

  <input type="hidden" name="MM_delete" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= studentDB.Fields.Item("MemId").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
studentDB.Close()
Set studentDB = Nothing
%>
