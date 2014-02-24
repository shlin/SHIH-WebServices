<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="../Connections/Dal.asp" -->
<%
Dim modify__MMColParam
modify__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  modify__MMColParam = Request.QueryString("id")
End If
%>
<!-- #include file="fckeditor.asp" -->

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
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>

<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_Dal_STRING
    MM_editCmd.CommandText = "UPDATE dbo.Content SET cname = ?, name = ?, Content = ?, [date] = ? WHERE id = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 50, Request.Form("cname")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 50, Request.Form("name")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 203, 1, 1073741823, Request.Form("Content1")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 10, Request.Form("date")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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
%>



<body background="../img/body.gif">
<form method="POST" action="<%=MM_editAction%>" name="form1">


  <table align="center" id="form">
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="center">中文類別:</div></td>
      <td><input type="text" name="cname" value="<%=(modify.Fields.Item("cname").Value)%>" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="center">英文類別</div></td>
      <td><input type="text" name="name" value="<%=(modify.Fields.Item("name").Value)%>" size="32">      </td>
    </tr>
    <tr valign="baseline">
     <td align="center" valign="middle" nowrap="nowrap"><div align="center">內容</div></td> 

	 <td width="800" height="300">	 	

	 	 <%
' Automatically calculates the editor base path based on the _samples directory.
' This is usefull only for these samples. A real application should use something like this:
'oFCKeditor.BasePath = "/fckeditor/" 	 '/fckeditor/' is the default value.
Dim sBasePath
sBasePath = Request.ServerVariables("PATH_INFO")
sBasePath = Left( sBasePath, InStrRev( sBasePath, "/" ) )



Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath	= sBasePath


oFCKeditor.Value	= modify.Fields.Item("Content").Value
oFCKeditor.Create "Content1"
		%>

	
		
<input name="Content" type="hidden" value="" size="32" /></td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="center">修改日期</div></td>
      <td><input type="text" name="date" value="<%=(modify.Fields.Item("date").Value)%>" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" nowrap="nowrap"><div align="center"></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr valign="baseline">
      <td align="center" valign="middle" nowrap="nowrap"><div align="center"></div></td>
      <td><input type="submit" value="Update Record">      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= modify.Fields.Item("id").Value %>">
</form>
<p>&nbsp;</p>
</body>
<%
modify.Close()
Set modify = Nothing
%>
