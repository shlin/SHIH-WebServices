<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="../../Connections/Dal.asp" -->
<%
Dim GetContent__MMColParam
GetContent__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  GetContent__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim GetContent
Dim GetContent_numRows

Set GetContent = Server.CreateObject("ADODB.Recordset")
GetContent.ActiveConnection = MM_Dal_STRING
GetContent.Source = "SELECT * FROM dbo.Content WHERE id = '" + Replace(GetContent__MMColParam, "'", "''") + "'"
GetContent.CursorType = 0
GetContent.CursorLocation = 2
GetContent.LockType = 1
GetContent.Open()

GetContent_numRows = 0
%>
<%
Dim gb
Dim gb_cmd
Dim gb_numRows

Set gb_cmd = Server.CreateObject ("ADODB.Command")
gb_cmd.ActiveConnection = MM_Dal_STRING
gb_cmd.CommandText = "SELECT * FROM dbo.GuestBook ORDER BY mid DESC" 
gb_cmd.Prepared = true

Set gb = gb_cmd.Execute
gb_numRows = 0
%>

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
gb_numRows = gb_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim gb_total
Dim gb_first
Dim gb_last

' set the record count
gb_total = gb.RecordCount

' set the number of rows displayed on this page
If (gb_numRows < 0) Then
  gb_numRows = gb_total
Elseif (gb_numRows = 0) Then
  gb_numRows = 1
End If

' set the first and last displayed record
gb_first = 1
gb_last  = gb_first + gb_numRows - 1

' if we have the correct record count, check the other stats
If (gb_total <> -1) Then
  If (gb_first > gb_total) Then
    gb_first = gb_total
  End If
  If (gb_last > gb_total) Then
    gb_last = gb_total
  End If
  If (gb_numRows > gb_total) Then
    gb_numRows = gb_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (gb_total = -1) Then

  ' count the total records by iterating through the recordset
  gb_total=0
  While (Not gb.EOF)
    gb_total = gb_total + 1
    gb.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (gb.CursorType > 0) Then
    gb.MoveFirst
  Else
    gb.Requery
  End If

  ' set the number of rows displayed on this page
  If (gb_numRows < 0 Or gb_numRows > gb_total) Then
    gb_numRows = gb_total
  End If

  ' set the first and last displayed record
  gb_first = 1
  gb_last = gb_first + gb_numRows - 1
  
  If (gb_first > gb_total) Then
    gb_first = gb_total
  End If
  If (gb_last > gb_total) Then
    gb_last = gb_total
  End If

End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = gb
MM_rsCount   = gb_total
MM_size      = gb_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
gb_first = MM_offset + 1
gb_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (gb_first > MM_rsCount) Then
    gb_first = MM_rsCount
  End If
  If (gb_last > MM_rsCount) Then
    gb_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>::: DAL - 決策分析實驗室 :::</title>

<link href="../../dal.css" rel="stylesheet" type="text/css" />
<script src="../../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>

<style type="text/css">
<!--

body {
background-attachment : fixed;
background-image: url(../../img/body.gif);



}

-->
</style>

<script type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>


</head>

<body onload="MM_preloadImages('../../img/menu/1-g.gif','../../img/menu/2-g.gif','../../img/menu/3-g.gif','../../img/menu/4-g.gif','../../img/menu/5-g.gif','../../img/menu/6-g.gif','../../img/menu/7-g.gif','../../img/menu/8-g.gif','../../img/menu/9-g.gif','../../img/menu/10-g.gif','../../img/menu/11-g.gif')">



<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
    <td><img src="../../img/banner.gif" alt="banner" width="758" height="115" border="0" usemap="#Map2"  />
      <map name="Map2" id="Map22">
        <area shape="rect" coords="157,61,366,96" href="http://www.ms.tku.edu.tw" alt="淡江大學管理科學研究所" />
        <area shape="rect" coords="410,64,610,97" href="http://www.im.tku.edu.tw" alt="淡江大學資訊管理研究所" />
        <area shape="rect" coords="221,10,543,56" href="../../index.asp" alt="決策分析研究室" />
        <area shape="circle" coords="65,57,50" href="../../index.asp" alt="決策分析研究室" />
        <area shape="circle" coords="690,57,51" href="http://www.tku.edu.tw" alt="淡江大學" />
      </map>
    </td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><div align="right"></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><img src="../../img/left.gif" alt="left" width="4" height="250" align="right" /></td>
    <td valign="top"><table width="760" border="0" align="center" cellpadding="00" cellspacing="0" class="border"   id="leftlink">
      <tr>
        <td width="152" valign="middle" background="../../img/MenuBg.jpg" class="MenuSmall">&nbsp;</td>
        <td width="608" height="24" colspan="2" align="right" valign="middle" bgcolor="#1ca4c0"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div align="left" class="MenuSmall"><img src="../../img/w.gif" alt="w" width="15" height="15" />::: Every decision you make is a mistake. -- Edward Dahlberg :::</div></td>
            <td><div align="right"><span class="MenuSmall"><img src="../../img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Search<img src="../../img/w.gif" alt="w" width="15" height="15" /><img src="../../img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Contact Us<img src="../../img/w.gif" alt="w" 
			width="15" height="15" /></span>
			<span class="MenuSmall" onclick="MM_goToURL('parent','../../index.asp');return document.MM_returnValue"><img src="../../img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Home<img src="../../img/w.gif" alt="w" width="15" height="15" /></span></div></td>
          </tr>
        </table></td>
        </tr>
      <tr>
        <td colspan="2" valign="top"><span class="smalltext"><img src="../../img/w.gif" alt="w" width="15" height="10" /></span></td>
      </tr>
      <tr>
        <td colspan="6" valign="top"><table width="760" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td align="left" valign="top"><!--#include virtual="/left_menu_module.asp" -->

              </td>
            <td align="center" valign="top"><table width="600" height="188" border="0" cellpadding="0" cellspacing="0" class="tb_border" id="link_content">
                <tr>
                  <td valign="bottom" background="../../img/MenuBg3.gif"><div align="center">
                    <p><img src="../../img/InsideIcon.gif" alt="Icon" width="14" height="14" align="absmiddle" /> <span class="InsideTitle">GuestBook</span><br />
                      <br />
                      </p>
                    </div>                      </td>
                </tr>
                <tr>
                  <td height="100%" align="center" valign="middle" background="../../img/tb_bg6_0.gif"><div align="left">
                    <p align="left"><img src="../../img/w.gif" alt="white" width="15" height="15" /><a href="post_form.asp">留言│Post</a></p>
                      <table width="95%" border="0" align="center" class="smalltext">
                        <tr>
                          <td width="23%" align="center"><% If MM_offset <> 0 Then %>
                              <a href="<%=MM_moveFirst%>">第一頁</a>
                              <% End If ' end MM_offset <> 0 %>
</td>
                          <td width="31%" align="center"><% If MM_offset <> 0 Then %>
                              <a href="<%=MM_movePrev%>">上一頁</a>
                              <% End If ' end MM_offset <> 0 %>
</td>
                          <td width="23%" align="center"><% If Not MM_atTotal Then %>
                              <a href="<%=MM_moveNext%>">下一頁</a>
                              <% End If ' end Not MM_atTotal %>
                          </td>
                          <td width="23%" align="center"><% If Not MM_atTotal Then %>
                              <a href="<%=MM_moveLast%>">最後一頁</a>
                              <% End If ' end Not MM_atTotal %>
</td>
                        </tr>
                      </table>
                      
                          <div align="center" class="smalltext">留言<%=(gb_first)%> 到 <%=(gb_last)%> / 共 <%=(gb_total)%> 則</div>
                          <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000" bordercolorlight="#C0C0C0" id="table1">

	
	
   <% 
While ((Repeat1__numRows <> 0) AND (NOT gb.EOF)) 
%>



  <tr>
    <td width="60" height="46" background="../../img/tb/1-a.gif">&nbsp;</td>
    <td background="../../img/tb/1-b.gif">&nbsp;</td>
    <td width="56" background="../../img/tb/1-c.gif">&nbsp;</td>
  </tr>
  <tr>
    <td width="56" background="../../img/tb/2-a.gif"><img src="../../img/w.gif" alt="white" width="10" height="15" />姓名│</td>
    <td><%=(gb.Fields.Item("Name").Value)%></td>
    <td background="../../img/tb/2-b.gif">&nbsp;</td>
  </tr>
  <tr>
    <td background="../../img/tb/2-a.gif"><img src="../../img/w.gif" alt="white" width="10" height="15" />信箱│</td>
    <td><p><a href="mailto://<%=(gb.Fields.Item("Email").Value)%>"><%=(gb.Fields.Item("Email").Value)%></a></p>        </td>
    <td rowspan="5" background="../../img/tb/2-b.gif">&nbsp;</td>
  </tr>
  <tr>
    <td background="../../img/tb/2-a.gif"><img src="../../img/w.gif" alt="white" width="10" height="15" />主題│</td>
    
	
	 <% BString=(gb.Fields.Item("Subject").Value) %>
	  
	 <%if left(BString,3)="Re:" then%>

      <td  bgcolor="#FFCCFF"><%Response.Write(BString)%></td>
	<%else%>
	     <td ><%Response.Write(BString)%>
		 <%end if%>		 </td>
  </tr>
  <tr>
    <td valign="top" background="../../img/tb/2-a.gif"><img src="../../img/w.gif" alt="white" width="10" height="15" />內容│</td>
    <td>
	  
	  <table width="100%" border="0" align="left" cellpadding="0" cellspacing="0">
        <tr>
          <td bgcolor="#EAEAEA"><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td><% Bmemo = Replace(gb.Fields.Item("Memo").Value,vbCrlf,"<br>") %>
                <% Response.Write(Bmemo) %></td>
            </tr>
          </table>          </td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td background="../../img/tb/2-a.gif"><img src="../../img/w.gif" alt="white" width="10" height="15" />時間│</td>
    <td><%=(gb.Fields.Item("Time").Value)%></td>
  </tr>
  <tr>
    <td background="../../img/tb/2-a.gif"><img src="../../img/w.gif" alt="white" width="10" height="15" />位址│</td>
    <td><%=(gb.Fields.Item("ClientIP").Value)%></td>
  </tr>
  <tr>
    <td background="../../img/tb/2-a.gif">&nbsp;</td>
    <td><div align="right"></div></td>
    <td background="../../img/tb/2-b.gif"><span class="smalltext">No.<%=(gb.Fields.Item("mid").Value)%></span></td>
  </tr>
  <tr>
    <td background="../../img/tb/2-a.gif">&nbsp;</td>
    <td><div align="center">
	<a href="reply.asp?mid=<%=(gb.Fields.Item("mid").Value)%>">回覆</a>│
	<a href="del.asp?mid=<%=(gb.Fields.Item("mid").Value)%>">刪除</a>│
	<a href="mod.asp?mid=<%=(gb.Fields.Item("mid").Value)%>">修改</div></td>
    <td background="../../img/tb/2-b.gif"><a href="reply.asp?mid=<%=(gb.Fields.Item("mid").Value)%>"></a></td>
  </tr>
  <tr>
    <td height="16" background="../../img/tb/3-a.gif">&nbsp;</td>
    <td background="../../img/tb/3-b.gif">&nbsp;</td>
    <td height="16" background="../../img/tb/3-c.gif">&nbsp;</td>
  </tr>
  <tr>
    <td height="16"><img src="../../img/w.gif" alt="white" width="10" height="15" /></td>
    <td>&nbsp;</td>
    <td height="16">&nbsp;</td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  gb.MoveNext()
Wend
%>
  </table>
                  </div></td>
                </tr>
                <tr>
                  <td height="43" background="../../img/tb_bg6_1.gif"><div align="right" class="smalltext">頁面&amp;程式設計：鄭勝仁<img src="../../img/w.gif" alt="w" width="15" height="15" /></div></td>
                </tr>
                </table>              
              </td>
          </tr>
          <tr>
            <td valign="top">
			
<SCRIPT language="JavaScript"> 
<!--  
document.write('<div align="center"><font size="1" color="gray"><font face="Verdana, Arial, Helvetica, sans-serif">'); 
document.write("<span id='clock'></span>"); 
var now,hours,minutes,seconds,timeValue; 
function showtime(){ 
now = new Date(); 
hours = now.getHours(); 
minutes = now.getMinutes(); 
seconds = now.getSeconds(); 
timeValue = (hours >= 12) ? "PM " : "AM "; 
timeValue += ((hours > 12) ? hours - 12 : hours) + " 時"; 
timeValue += ((minutes < 10) ? " 0" : " ") + minutes + " 分"; 
timeValue += ((seconds < 10) ? " 0" : " ") + seconds + " 秒"; 
clock.innerHTML = timeValue; 
setTimeout("showtime()",100); 
} 
showtime(); 
document.write('</font></div>');
//--> 
</SCRIPT>
			
			<br /></td>
            <td align="right" valign="bottom"><div align="center"></div>
            <table width="100%" border="0" cellpadding="0" cellspacing="0" id="link">
              <tr>
                <td class="smalltext"><p>工業工程推廣影片：<a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext">五分鐘版</a> <a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext"></a> │<a href="http://www.ciie.org.tw/Video/IEM_10.wmv" class="smalltext">十分鐘版</a></p></td>
                <td><div align="right"><img src="../../img/copyright6.gif" alt="CopyRight" width="337" height="37" /><img src="../../img/w.gif" alt="white" width="15" height="15" /></div></td>
              </tr>
            </table></td>
          </tr>
          
        </table></td>
      </tr>
    </table></td>
    <td><img src="../../img/right.gif" alt="right" width="4" height="250" /></td>
  </tr>
  
</table>

</body>
</html>
<%
GetContent.Close()
Set GetContent = Nothing
%>
<%
gb.Close()
Set gb = Nothing
%>

