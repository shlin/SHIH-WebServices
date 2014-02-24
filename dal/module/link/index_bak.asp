<table class="tableborder">
  <tr>
	<td colspan="8" class="forumRowHeader"><font face="標楷體" size="4">
    請直接點選檔名觀看，或按滑鼠右鍵右存目標下載，<font color="#ff0000">檔案為PDF檔需用Adobe Reader觀看</font></font></td>
  </tr>
  
<%
'response.write Server.Mappath(fr)
'response.end
Set FSO = CreateObject("Scripting.FileSystemObject")
if Request("DirName") =  "" then
	'DirNameMP=Request.ServerVariables("APPL_PHYSICAL_PATH")
	'DirNameMP=Request.ServerVariables("PATH_TRANSLATED")
	PathTemp=Request.ServerVariables("PATH_TRANSLATED")
	PathTempS=split(PathTemp,"\")
	for i = 0 to Ubound(PathTempS) - 1
		DirNameMP=DirNameMP & PathTempS(i) & "\"
	next
	DirNameMP=DirNameMP & "..\..\File\"
else
	DirName=Request("DirName") 
	DirNameMP=Server.MapPath("File\"&DirName)
end if
'response.write DirNameMP
'response.end
Set Folder = FSO.GetFolder(DirNameMP)
DirNameTemp=split(DirName,"/")
IF UBound(DirNameTemp) > 0 then
FDirName=""
'response.write Lbound(DirNameTemp)
'response.write UBound(DirNameTemp)
%>
<%
for i = Lbound(DirNameTemp) to UBound(DirNameTemp) - 2
	'FDirName = FDirName & DirNameTemp(i) & "/"
	'response.write FDirName & "<br>"
next
%>
  <tr>
    <td colspan="8"  valign="middle" class="forumRow">
		<a href="?DirName=<%=FDirName%>&amp;FieldName=<%=request("FieldName")%>">..回到上層</a>	</td>
  </tr>
<%
end if
%>

<%
set SubFolder = Folder.SubFolders
For Each f1 in SubFolder
%>
<tr>
    <td width="100%" height="24">
		<img border="0" src="icon/dir.gif" align="absmiddle">
		<a href="?dirname=<%=f1.name%>"><%=f1.name%></a>
		</td>
</tr>
<%
Next
%>

<%
'檔案清單
intIndex=0
Set FileCollection = Folder.Files
For Each file in FileCollection
'ekstenzija = Right(file.name,3)
  intIndex=intIndex+1
%>
  <%if intindex = 1 then%><tr><%end if%>
    <td class="forumRow"  valign="middle" align="center">&nbsp;
	<a href="file/<%=file.name%>" target="_blank"><img src="icon/file.gif" border="0"><br><%=file.name%></a><%=file.DateLastModified%></td>
  <%if intindex > 7 then%></tr><%end if%>
<%
	if intindex > 7 then
		intIndex=0
	end if
Next
if intindex <= 7 and IntIndex > 0 then 
	response.write "<td colspan="&8-intIndex&">&nbsp;</td>"
	response.write "</tr>"
end if
%>
</table>