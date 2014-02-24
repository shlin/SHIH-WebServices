
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
    MM_editCmd.CommandText = "DELETE FROM dbo.Content WHERE id = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 200, 1, 10, Request.Form("MM_recordId")) ' adVarChar
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
Dim modify__MMColParam
modify__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  modify__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim modify
Dim modify_cmd
Dim modify_numRows

Set modify_cmd = Server.CreateObject ("ADODB.Command")
modify_cmd.ActiveConnection = MM_Dal_STRING
modify_cmd.CommandText = "SELECT * FROM dbo.Content WHERE id = ?" 
modify_cmd.Prepared = true
modify_cmd.Parameters.Append modify_cmd.CreateParameter("param1", 200, 1, 10, modify__MMColParam) ' adVarChar

Set modify = modify_cmd.Execute
modify_numRows = 0
%><style type="text/css">
<!--
body {
	background-image: url(../img/body.gif);
}
-->
</style>

<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
  <table align="center" id="form">
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="center">索引鍵</div></td>
      <td><%=(modify.Fields.Item("id").Value)%></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="center">中文類別:</div></td>
      <td><%=(modify.Fields.Item("cname").Value)%></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="center">英文類別</div></td>
      <td><%=(modify.Fields.Item("name").Value)%></td>
    </tr>
    <tr valign="baseline">
      <td align="center" valign="middle" nowrap="nowrap"><div align="center">內容</div></td>
      <td><label>
        <textarea name="textfield" cols="100" rows="10"><%=(modify.Fields.Item("Content").Value)%></textarea>
      </label></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="center">修改日期</div></td>
      <td><%=(modify.Fields.Item("date").Value)%></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="center"></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr valign="baseline">
      <td align="center" valign="middle" nowrap="nowrap"><div align="center"></div></td>
      <td><input type="submit" value="Delete Record">      </td>
    </tr>
  </table>
  
  

  <input type="hidden" name="MM_delete" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= modify.Fields.Item("id").Value %>">
</form>
<p>&nbsp;</p>
<style type="text/css">
<!--
body {
	background-image: url(../img/body.gif);
}
-->
</style>
<%
modify.Close()
Set modify = Nothing
%>
