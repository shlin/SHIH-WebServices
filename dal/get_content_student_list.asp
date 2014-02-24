<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!--#include file="Connections/Dal.asp" -->
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
Dim studentDB__MMColParam
studentDB__MMColParam = "1"
If (Request.QueryString("year") <> "") Then 
  studentDB__MMColParam = Request.QueryString("year")
End If
%>
<%
Dim studentDB
Dim studentDB_cmd
Dim studentDB_numRows

Set studentDB_cmd = Server.CreateObject ("ADODB.Command")
studentDB_cmd.ActiveConnection = MM_Dal_STRING
studentDB_cmd.CommandText = "SELECT * FROM dbo.Student WHERE [year] = ? AND class LIKE '研究生'" 
studentDB_cmd.Prepared = true
studentDB_cmd.Parameters.Append studentDB_cmd.CreateParameter("param1", 200, 1, 10, studentDB__MMColParam) ' adVarChar

Set studentDB = studentDB_cmd.Execute
studentDB_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
studentDB_numRows = studentDB_numRows + Repeat1__numRows
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>::: DAL - 決策分析實驗室 :::</title>

<link href="dal.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>

<style type="text/css">
<!--

body {
background-attachment : fixed;
background-image: url(img/body.gif);



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

<body>



<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
    <td><img src="img/banner.gif" alt="banner" width="758" height="115" border="0" usemap="#Map2"  />
      <map name="Map2">
        <area shape="rect" coords="157,61,366,96" href="http://www.ms.tku.edu.tw" alt="淡江大學管理科學研究所" />
        <area shape="rect" coords="410,64,610,97" href="http://www.im.tku.edu.tw" alt="淡江大學資訊管理研究所" />
        <area shape="rect" coords="221,10,543,56" href="index.asp" alt="決策分析研究室" />
        <area shape="circle" coords="65,57,50" href="index.asp" alt="決策分析研究室" />
        <area shape="circle" coords="690,57,51" href="http://www.tku.edu.tw" alt="淡江大學" />
      </map>
    </td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><div align="right"><span class="smalltext">Hits :
          <!--#include file="module/counter/counter.asp" -->
    │Today : <script type="text/javascript">var date = new Date();document.write(date.getFullYear() + "/" + date.getMonth() + "/" + date.getDate())</script></span></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><img src="img/left.gif" alt="left" width="4" height="250" align="right" /></td>
    <td valign="top"><table width="760" border="0" align="center" cellpadding="00" cellspacing="0" class="border"   id="leftlink">
      <tr>
        <td width="152" valign="middle" background="img/MenuBg.jpg" class="MenuSmall">&nbsp;</td>
        <td width="608" height="24" colspan="2" align="right" valign="middle" bgcolor="#1ca4c0">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div align="left" class="MenuSmall"><img src="img/w.gif" alt="w" width="15" height="15" />::: Every decision you make is a mistake. -- Edward Dahlberg :::</div></td>
            <td><div align="right"><span class="MenuSmall"><img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Search<img src="img/w.gif" alt="w" width="15" height="15" /><img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Contact Us<img src="img/w.gif" alt="w" 
			width="15" height="15" /></span>
			<span class="MenuSmall" onclick="MM_goToURL('parent','index.asp');return document.MM_returnValue"><img src="img/MenuPic.gif" alt="Menu" width="12" height="13" align="absmiddle" />Home<img src="img/w.gif" alt="w" width="15" height="15" /></span></div></td>
          </tr>
        </table></td>
        </tr>
      <tr>
        <td colspan="2" valign="top"><span class="smalltext"><img src="img/w.gif" alt="w" width="15" height="10" /></span></td>
      </tr>
      <tr>
        <td colspan="6" valign="top"><table width="760" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td align="left" valign="top"><!--#include file="left_menu.asp" -->
          
              </td>
            <td align="center" valign="top"><table width="600" height="198" border="0" cellpadding="0" cellspacing="0" class="tb_border" id="link_content">
                <tr>
                  <td height="34" align="center" valign="top" background="img/MenuBg3.gif">
                    <p><img src="img/InsideIcon.gif" alt="Icon" width="14" height="14" align="absmiddle" /> <span class="InsideTitle"><%=(GetContent.Fields.Item("name").Value)%></span><br />
                      <br />
                      </p>
                      </td>
                </tr>
                <tr>
                  <td height="100%" align="left" valign="top" background="img/tb_bg6_0.gif">
				  <div align="left">
                    <% 
While ((Repeat1__numRows <> 0) AND (NOT studentDB.EOF)) 
%>
                      <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td><div align="center"><%=(studentDB.Fields.Item("type").Value)%><br />
                                  <br />
                          </div></td>
                        </tr>
                        <tr>
                          <td><div align="center"><span class="student_topic"><%=(studentDB.Fields.Item("Topic").Value)%></span><br />
                                    <br />
                          </div></td>
                        </tr>
                        <tr>
                          <td><div align="center"><%=(studentDB.Fields.Item("name").Value)%><br />
                                  <br />
                          </div></td>
                        </tr>
                        <tr>
                          <td><%=(studentDB.Fields.Item("content").Value)%></td>
                        </tr>
                        <tr>
                          <td><img src="img/line.gif" alt="LINE" width="100%" height="8" /><br />
                            <br /></td>
                        </tr>
                      </table>
                      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  studentDB.MoveNext()
Wend
%>
<p align="center">&lt;<a href="get_content.asp?id=2">上一頁</a>&gt;</p>
				  </div></td>
                </tr>
                <tr>
                  <td height="43" background="img/tb_bg6_1.gif"><div align="right" class="smalltext"></div></td>
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
var now,hours,minutes,seconds,timeValue,showtime;
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
		  </td>
            <td align="right" valign="bottom"><div align="center"></div>
            <table width="100%" border="0" cellpadding="0" cellspacing="0" id="link">
              <tr>
                <td class="smalltext"><p>工業工程推廣影片：<a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext">五分鐘版</a> <a href="http://www.ciie.org.tw/Video/IEM_05.wmv" class="smalltext"></a> │<a href="http://www.ciie.org.tw/Video/IEM_10.wmv" class="smalltext">十分鐘版</a></p></td>
                <td><div align="right"><img src="img/copyright6.gif" alt="CopyRight" width="337" height="37" /><img src="img/w.gif" alt="white" width="15" height="15" /></div></td>
              </tr>
            </table></td>
          </tr>
          
        </table></td>
      </tr>
    </table></td>
    <td><img src="img/right.gif" alt="right" width="4" height="250" /></td>
  </tr>
  
</table>

</body>
</html>
<%
GetContent.Close()
Set GetContent = Nothing
%>
<%
studentDB.Close()
Set studentDB = Nothing
%>