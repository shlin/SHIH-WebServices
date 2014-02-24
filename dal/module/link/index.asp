<!--#include file="config.inc.asp"-->

<%
'response.write Server.Mappath(fr)
Set FSO = CreateObject("Scripting.FileSystemObject")
if Request("DirName") <>  "" then
	'DirNameMP=Request.ServerVariables("APPL_PHYSICAL_PATH")
	'DirNameMP=Request.ServerVariables("PATH_TRANSLATED")
	DirName=Request("DirName") 
	DirNameMP=Server.MapPath(config_path&DirName)
	PathTempS=split(DirName,"\")
	for i = 0 to Ubound(PathTempS) - 2
		DirNameFP=DirNameFP & PathTempS(i) & "\"
	next
	'response.write DirNameFP
else
	DirName=Request("DirName") 
	DirNameMP=Server.MapPath(config_path&DirName)
end if
'response.write DirName
'response.write DirNameMP
'response.end
showkind=Request("ShowKind")
if showkind ="" then
	showkind=2 '開啟時要的模式
end if
'response.write Request("ShowKind")
Set Folder = FSO.GetFolder(DirNameMP)
DirNameTemp=split(DirName,"/")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=config_header%></title>
<link href="main.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript">
	function openlink(strLink){
		window.open(strLink, '','toolbar=0,location=0,directories=0,status=1,menubar=0,scrollbars=yes,resizable=yes,width=500,height=400');
	}
</script>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {color: #FFFFFF}
a {
	font-size: 16px;
	color: #666666;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #666666;
}
a:hover {
	text-decoration: none;
	color: #FF66FF;
}
a:active {
	text-decoration: none;
	color: #FF33FF;
}
body,td,th {
	font-size: 16px;
	color: #666666;
}
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=big5" /></head>
<body>
<table width="100%" border="0" cellpadding="10" cellspacing="0">
  <tr>
    <td>
      <table width="100%" border="0">
        <tr align="left" valign="top">
					<%
					if config_is_SubFolders=true then
					%>
          <td width="21%">
            <table border="0" cellpadding="0" cellspacing="0" class="table1">
              <tr>
                <td width="145" class="td1">資料夾</td>
              </tr>
              <%
							if DirName <> "" then
							%>
              <tr>
                <td colspan="8"  valign="middle" class="forumRow"><a href="?showkind=<%=showkind%>&DirName=<%=DirNameFP%>">..回到上層</a></td>
              </tr>
              <%
							end if
							
							set SubFolder = Folder.SubFolders
							For Each f1 in SubFolder
							%>
              <tr>
                <td> <img border="0" src="icon/dir.gif" align="absmiddle"><a href="?showkind=<%=showkind%>&dirname=<%=DirName%><%=f1.name%>\"><%=f1.name%></a> </td>
              </tr>
              <%
							Next
							%>
            </table>
					</td>
          <td width="100">&nbsp;</td>
          <%
					end if
					%>
          <td width="79%">
            <form name="frmMain">
						<input name="dirname" type="hidden" value="<%=dirname%>" />
              <table width="601" border="0" cellpadding="0" cellspacing="0" class="table1">
                <tr>
                  <td width="82%" class="td1">檔案清單</td>
                  <td width="18%" align="right" class="td1">檢視依據：
                    <select name="showkind" id="showkind" onchange="frmMain.submit()">
                      <option value="1" <%if showkind=1 then response.write " selected"%>>縮圖</option>
                      <option value="2" <%if showkind=2 then response.write " selected"%>>清單</option>
                    </select>
                  </td>
                </tr>
                <tr>
                  <td colspan="2">
                   
										<%
										if showkind ="" or showkind=1 then
										%>
										 <table border="0" cellpadding="0" cellspacing="0">
										<%
											'檔案清單
											intIndex=0
											Set FileCollection = Folder.Files
											For Each file in FileCollection
												'response.write config_show_hidden
												'response.write file.Attributes
												if (file.Attributes <> 38 and file.Attributes <> 34 and file.Attributes <> 4 and file.Attributes <> 2) or config_show_hidden <>false then
												'ekstenzija = Right(file.name,3)
												intIndex=intIndex+1
											%>
														<%if intindex = 1 then%>
														<tr>
															<%end if%>
															<%'file.DateLastModified%>
															<td valign="middle" align="center" width="15%">&nbsp; <a href="<%=config_path%><%=dirname%><%=file.name%>" target="_blank"><img src="icon/file.gif" border="0"><br>
																<%=file.name%></a></td>
															<%if intindex > 7 then%>
														</tr>
														<%end if%>
														<%
													if intindex > 7 then
														intIndex=0
													end if
												end if
											Next
											if intindex <= 7 and IntIndex > 0 then 
												response.write "<td colspan="&8-intIndex&">&nbsp;</td>"
												response.write "</tr>"
											end if
										%>
										</table>
										<%
										end if
										%>
										
                  <%
									if showkind=2 then
									%>
									<table width="100%" border="0" cellpadding="0" cellspacing="0" class="td3">
													<tr>
														<td colspan="2" align="left" valign="middle">檔名</td>
														<td valign="middle" align="left" width="15%">檔案大小</td>
														<td valign="middle" align="left" width="15%">建立時間</td>
														<td valign="middle" align="left" width="15%">最後存取時間</td>
														<td valign="middle" align="left" width="15%">最後修改時間</td>
													</tr>
									<%
										'檔案清單
										intIndex=0
										Set FileCollection = Folder.Files
										
										For Each file in FileCollection
											if (file.Attributes <> 38 and file.Attributes <> 34 and file.Attributes <> 4 and file.Attributes <> 2) or config_show_hidden=true then
											usersize=file.size
											if usersize > 1024 then
												usersize=round(usersize/1024) & "KB"
											elseif usersize > 1024*1024 then
												usersize=round(usersize/(1024*1024),2) & "MB"
											else
												usersize=usersize/1024 & "Byte"
											end if
											'response.write "test"
											'response.end
										%>
													<tr>
														<td valign="middle" align="left">
														  <div align="center"><a href="file/<%=dirname%><%=file.name%>" target="_blank"><img src="icon/file.gif" border="0"></a></div></td>
														<td align="left" valign="middle" class="td3"><a href="<%=config_path%><%=dirname%><%=file.name%>" target="_blank"><%=file.name%></a></td>
														<td valign="middle" align="left"><%=usersize%></td>
														<td valign="middle" align="left"><%=file.DateCreated%></td>
														<td valign="middle" align="left"><%=file.DateLastAccessed%></td>
														<td valign="middle" align="left"><%=file.DateLastModified%></td>
													</tr>
										<%
											end if
										Next
									%>
					</table>
									<%
									end if
									%>										
                    
                  </td>
                </tr>
                <tr class="td2">
                  <td colspan="2" align="center">&nbsp;</td>
                </tr>
              </table>
            </form>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
