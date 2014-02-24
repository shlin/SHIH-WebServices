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
    MM_editCmd.CommandText = "UPDATE dbo.Student SET Topic = ?, [year] = ?, name = ?, [class] = ?, type = ?, content = ? WHERE MemId = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 200, Request.Form("Topic")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 10, Request.Form("year")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 203, 1, 1073741823, Request.Form("name")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 50, Request.Form("class")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, -1, Request.Form("type")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 203, 1, 1073741823, Request.Form("content1")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "list_student.asp"
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
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>�s�W</title>
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
      <td width="74" align="right" valign="middle" nowrap><div align="center">�D��</div></td>
      <td><input type="text" name="Topic" value="<%=(studentDB.Fields.Item("Topic").Value)%>" size="80">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">�~��</div></td>
      <td><input type="text" name="year" value="<%=(studentDB.Fields.Item("year").Value)%>" size="32">      </td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">�m�W</div></td>
      <td><input name="name" type="text" id="name" value="<%=(studentDB.Fields.Item("name").Value)%>" size="80" />
      �ХH&quot;�B&quot;�y���@�����j</td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">�ŧO</div></td>
      <td>
	  <input <%If (CStr((studentDB.Fields.Item("class").Value)) = CStr("��s��")) Then Response.Write("checked=""checked""") : Response.Write("")%> name="class" type="radio" value="��s��" />
        ��s��
        <input <%If (CStr((studentDB.Fields.Item("class").Value)) = CStr("�j�ǥ�")) Then Response.Write("checked=""checked""") : Response.Write("")%> name="class" type="radio" value="�j�ǥ�" />
        �j�ǥ� 
        <input <%If (CStr((studentDB.Fields.Item("class").Value)) = CStr("�䥦")) Then Response.Write("checked=""checked""") : Response.Write("")%> name="class" type="radio" value="�䥦"/>
        �䥦 </td>
    </tr>
    <tr valign="baseline">
      <td align="right" valign="middle" nowrap><div align="center">���O</div></td>
      <td>
      <input <%If (CStr((studentDB.Fields.Item("type").Value)) = CStr("�դh�פ�")) Then Response.Write("checked=""checked""") : Response.Write("")%> name="type" type="radio" value="�դh�פ�" />
�դh�פ�
<input <%If (CStr((studentDB.Fields.Item("type").Value)) = CStr("�Ӥh�פ�")) Then Response.Write("checked=""checked""") : Response.Write("")%> name="type" type="radio" value="�Ӥh�פ�" />
�Ӥh�פ�
  <input <%If (CStr((studentDB.Fields.Item("type").Value)) = CStr("�M�D��s")) Then Response.Write("checked=""checked""") : Response.Write("")%> name="type" type="radio" value="�M�D��s" />
�M�D��s
<input <%If (CStr((studentDB.Fields.Item("type").Value)) = CStr("�䥦")) Then Response.Write("checked=""checked""") : Response.Write("")%> name="type" type="radio" value="�䥦" />
�䥦      </td>
    </tr>
    <tr>
      <td nowrap align="right" valign="middle"><p align="center">���e</p>      </td>
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


oFCKeditor.Value	= (studentDB.Fields.Item("content").Value)
oFCKeditor.Create "Content1"
		%>
	  <input type="hidden" name="content" cols="100" rows="20"><!--</textarea>-->      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right"><div align="center">�@</div></td>
      <td><input name="Mod" type="submit" id="Mod" value="��s�O��">      </td>
    </tr>
  </table>
  

  

  
  

    <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= studentDB.Fields.Item("MemId").Value %>">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
studentDB.Close()
Set studentDB = Nothing
%>
