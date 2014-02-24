<!--#include file="Config2.asp"-->
<link href="master.css" type="text/css" rel="stylesheet">


	
			
<%
	set frst = Server.CreateObject("adodb.recordset")
	sql = "select * from info order by uploadtime desc"
	frst.open sql,myconn,1,1
	fcount = frst.recordcount
	if fcount > 0 then 	
		''顯示參數
		dim tbbgcolor
		''分頁參數
		dim maxperpage,pages,page
		maxperpage = 10
		frst.pagesize = maxperpage
		pages = frst.pagecount
		''頁面參數設置
		page = Request.QueryString("page")
		if not isnumeric(page) then page = 1 else page = cint(page)
		if page < 1 then page = 1
		if page > pages then page = pages
		frst.absolutepage = page
		''顯示內容
'Set Isize=Server.CreateObject("WinImg.ImgSize")
		for i = 1 to maxperpage
			if frst.eof then exit for
			if i mod 2 = 0 then tbbgcolor = "" else tbbgcolor = "#0066cc"
			fid = frst("id").Value
			ftitle = frst("fileTitle").Value
			fdesc = frst("fileDesc").Value
			ftype = frst("fileType").Value
			fpath = frst("filePath").Value
			fsize = frst("filesize").Value
			fhits = frst("hits").Value
			fuploadtime = frst("uploadTime").Value
			ftime = frst("NewsTime").Value
			ffrom = frst("FILEFrom").Value
			
'FileNameStr=Server.Mappath(fpath)
'Isize.GetImgSize Cstr(FileNameStr)
	%>
	
	<table width="95%" border="1" cellpadding="0" bordercolor="#EBEBEB" id="table3" style="border-collapse: collapse">
  <tr>
    <td align="left"><img src="../../img/arrow2.gif" alt="arrow" /><%=ftime%></td>
    <td width="78%" align="left"><b><a href="<%=fpath%>" target="_NEWwIN"><%=ftitle%></a></td>
  </tr>
  
  <tr>
    <td align="left"><div align="center"><font class="one"><%=ftype%></font></div></td>
    <td rowspan="3" ><p align="left">　　<%=fdesc%><a href="<%=fpath%>" target="_NEWwIN">(詳全文)</a></p>    </td>
  </tr>
 
  <tr>
    <td align="left"><div align="center"><%=ffrom%></div></td>
  </tr>
  <tr>
    <td align="left"><div align="center"><%=fsize%>(bytes)</div></td>
  </tr>

  <tr>
    <td colspan="2"><p align="right" class="smalltext">上傳日期：<%=fuploadtime%></p></td>
  </tr>
  
  	<tr>
							<td colspan="2" bgcolor="#1ca4c0">
							<img border="0" src="jpg" height="1"></td>
	  </tr>
</table>　
<%
		  	frst.movenext
		next
	  else
	  %> 
	<table border="1" cellpadding="0" style="border-collapse: collapse" width="95%" id="table3" bordercolor="#EBEBEB">
						<tr>
							<td width="300"></td>
							<td rowspan="3" valign="top">目前還沒有資料喔！</td>
						</tr>
						<tr>
							<td width="300"></td>
						</tr>
						<tr>
							<td width="300"></td>
						</tr>
						<tr>
							<td colspan="2" bgcolor="#1ca4c0">
							<img border="0" src="jpg" height="1"></td>
						</tr>
</table>  
	  <%
	  end if
	  frst.close
	  set frst = nothing
	  myconn.close
	  set myconn = nothing
	  %>				</td>
			</tr>
			
				<td background="img/tb_bg6_1.gif">
						
			＜<a href='?page=1'>首頁</a>
				<%
		  If Page > 1 Then Response.Write ("<a href='?id=10&page="& Page - 1 &"'>│上一頁</a>")
		  If page < pages Then Response.Write ("&nbsp;<a href='?id=10&page="& Page + 1 &"'>│下一頁</a>&nbsp;<a href='?page="& Pages &"'>│末頁</a>")
		  %>
		  
		  ＞
 </td>

