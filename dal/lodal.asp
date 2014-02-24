<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="Connections/Dal.asp" -->
<%
Dim userdb
Dim userdb_cmd
Dim userdb_numRows

Set userdb_cmd = Server.CreateObject ("ADODB.Command")
userdb_cmd.ActiveConnection = MM_Dal_STRING
userdb_cmd.CommandText = "SELECT * FROM dbo.[User]" 
userdb_cmd.Prepared = true

Set userdb = userdb_cmd.Execute
userdb_numRows = 0
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("nigol"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "mgt/admin.asp"
  MM_redirectLoginFailed = "index.asp"

  MM_loginSQL = "SELECT login, password"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM dbo.[User] WHERE login = ? AND password = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_Dal_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 255, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 255, Request.Form("drowssap")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>無標題文件</title>
<link href="dal.css" rel="stylesheet" type="text/css" />
</head>

<body>
<p>&nbsp;</p>
<form id="form1" name="form1" method="POST" action="<%=MM_LoginAction%>">
  <div align="center">
    <table border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="54" class="smalltext"><p align="center">Login </p>      </td>
        <td width="144"><input name="nigol" type="text" id="nigol" /></td>
      </tr>
      <tr>
        <td class="smalltext"><div align="center">Password</div></td>
        <td><input name="drowssap" type="password" id="drowssap" /></td>
      </tr>
    </table>
  </div>
  <p align="center">
    <input type="submit" name="Submit" value="送出" />
  </p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
</html>
<%
userdb.Close()
Set userdb = Nothing
%>

