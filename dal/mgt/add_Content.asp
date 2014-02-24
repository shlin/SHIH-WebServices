
<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/Dal.asp" -->
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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_Dal_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Content (cname, name, Content, [date]) VALUES (?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 50, Request.Form("cname")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, 50, Request.Form("name")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 203, 1, 1073741823, Request.Form("Content")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, 10, Request.Form("date")) ' adLongVarChar
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
<style type="text/css">
<!--
body {
	background-image: url(../img/body.gif);
}
-->
</style>

  <form method="POST" action="<%=MM_editAction%>" name="form1">
    <table align="center">
      <tr valign="baseline">
        <td nowrap="nowrap" align="right"><div align="center">中文類別:</div></td>
        <td><input type="text" name="cname" value="" size="32" />        </td>
      </tr>

      <tr valign="baseline">
        <td nowrap align="right"><div align="center">英文類別</div></td>
        <td><input type="text" name="name" value="" size="32">        </td>
      </tr>
      <tr valign="baseline">
        <td align="center" valign="middle" nowrap><div align="center">內容</div></td>
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


oFCKeditor.Value	= Request.Form("Content")
oFCKeditor.Create "Content"
		%>
		
		<input type="hidden" name="Content" cols="100" rows="30"><!--</textarea>-->        </td>
	  </tr>
      <tr valign="baseline">
        <td nowrap align="right"><div align="center">修改日期</div></td>
        <td><input type="text" name="date" value="<%Response.write date() %>" size="32">        </td>
      </tr>

      <tr valign="baseline">
        <td nowrap align="right"><div align="center"></div></td>
        <td><input type="submit" value="Insert Record">        </td>
      </tr>
    </table>
    
<input type="hidden" name="MM_insert" value="form1">
  </form>
  <p>&nbsp;</p>
