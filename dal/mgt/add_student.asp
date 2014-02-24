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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_Dal_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.Student (Topic, [year], name, class, type, content) VALUES (?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 200, Request.Form("Topic")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 10, Request.Form("year")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, 100, Request.Form("name")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 10, Request.Form("class")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, 50, Request.Form("type")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 203, 1, 1073741823, Request.Form("content")) ' adLongVarWChar
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
<!--#include file="fckeditor.asp" -->
<%
Dim studentDB
Dim studentDB_cmd
Dim studentDB_numRows

Set studentDB_cmd = Server.CreateObject ("ADODB.Command")
studentDB_cmd.ActiveConnection = MM_Dal_STRING
studentDB_cmd.CommandText = "SELECT * FROM dbo.Student" 
studentDB_cmd.Prepared = true

Set studentDB = studentDB_cmd.Execute
studentDB_numRows = 0
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>新增</title>
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
      <td><input type="text" name="Topic" value="" size="80">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">年份</div></td>
      <td><input type="text" name="year" value="" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">姓名</div></td>
      <td><input name="name" type="text" id="name" size="80" />
      請以&quot;、&quot;頓號作為分隔</td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">級別</div></td>
      <td>
        <input name="class" type="radio" value="研究生" />
        研究生
        <input name="class" type="radio" value="大學生" />
        大學生 
        <input name="class" type="radio" value="其它"/>
        其它        </td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">類別</div></td>
      <td>
      <input name="type" type="radio" value="博士論文" />
      博士論文
	<input name="type" type="radio" value="碩士論文" />
碩士論文
  <input name="type" type="radio" value="專題研究" />
專題研究
<input name="type" type="radio" value="其它" />
其它      </td>
    </tr>
    <tr>
      <td nowrap align="right" valign="middle"><p align="center">內容</p>      </td>
      <td width="714" valign="baseline">
	  
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
	  <input type="hidden" name="content" cols="100" rows="20"><!--</textarea> -->     </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right"><div align="center">　</div></td>
      <td><input type="submit" value="插入記錄">      </td>
    </tr>
  </table>
  

  <input type="hidden" name="MM_insert" value="form1">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
studentDB.Close()
Set studentDB = Nothing
%>
