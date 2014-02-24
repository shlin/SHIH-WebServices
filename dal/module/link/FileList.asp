<!--#include file = "CSS.asp"-->
<%
	FileName=request("FileName")
	FieldName=request("FieldName")
	
	if FileName <> "" then
	%>
	<script language = "JavaScript">
	window.opener.form1.<%=FieldName%>.value = '<%=FileName%>';
	self.close()
	</Script>
	<%
	End if
	%>
<table class="tableborder">
  <tr>
	<td colspan="8" class="forumRowHeader">檔案視窗 v1.1 </td>
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
	DirNameMP=DirNameMP & "File\"
else
	DirName=Request("DirName") & ""
	DirNameMP=Server.MapPath(DirName)
end if
'response.write DriNameMP
'response.end
Set Folder = FSO.GetFolder(DirNameMP)
DirNameTemp=split(DirName,"/")
IF UBound(DirNameTemp) > 0 then
FDirName=""
'response.write Lbound(DirNameTemp)
'response.write UBound(DirNameTemp)
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
'檔案清單
intIndex=0
Set FileCollection = Folder.Files
For Each file in FileCollection
'ekstenzija = Right(file.name,3)
  intIndex=intIndex+1
%>
  <%if intindex = 1 then%><tr><%end if%>
    <td class="forumRow"  valign="middle" align="center">&nbsp;
	<a href="file/<%=file.name%>" target="_blank"><img src="icon/file.gif" border="0"><br><%=file.name%></a></td>
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
  <tr>
    <td colspan="8" align="middle" class="forumRowHighlight"><a href="#" onclick="javascript:self.close()">關閉視窗</a></td>
  </tr>
</table>
