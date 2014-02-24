<%@LANGUAGE="VBSCRIPT"%>
<link href="dal.css" rel="stylesheet" type="text/css">
<!--#include file="Connections/Dal.asp" -->
<%
'找出研究生的所有年份
Dim studentDB
Dim studentDB_cmd
Dim studentDB_numRows

Set studentDB_cmd = Server.CreateObject ("ADODB.Command")
studentDB_cmd.ActiveConnection = MM_Dal_STRING
studentDB_cmd.CommandText = "SELECT Distinct [year] FROM dbo.Student WHERE class like '研究生' ORDER BY [year] ASC" 
studentDB_cmd.Prepared = true

Set studentDB = studentDB_cmd.Execute
studentDB_numRows = 0
%>
<%
'找出大學的所有年份
Dim StudentDB_uni
Dim StudentDB_uni_cmd
Dim StudentDB_uni_numRows

Set StudentDB_uni_cmd = Server.CreateObject ("ADODB.Command")
StudentDB_uni_cmd.ActiveConnection = MM_Dal_STRING
StudentDB_uni_cmd.CommandText = "SELECT Distinct [year] FROM dbo.Student WHERE class like '大學生' ORDER BY [year] ASC" 
StudentDB_uni_cmd.Prepared = true

Set StudentDB_uni = StudentDB_uni_cmd.Execute
StudentDB_uni_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
studentDB_numRows = studentDB_numRows + Repeat1__numRows
StudentDB_uni_numRows = StudentDB_uni_numRows + Repeat2__numRows
%>


<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
studentDB_numRows = studentDB_numRows + Repeat2__numRows
StudentDB_uni_numRows = StudentDB_uni_numRows + Repeat2__numRows
%>
<div style="height:350px;OVERFLOW-Y: scroll; OVERFLOW-X: hidden">
<table cellspacing="0" cellpadding="0" width="100%" border="0">
    <tbody>
        <tr>
            <td valign="top">
            <table cellspacing="0" cellpadding="0" width="25%" align="center" border="0">
                <tbody>
                    <tr>
                        <td valign="top" align="left" width="47%" height="205">
                        <table height="1" cellspacing="0" cellpadding="0" width="50%" border="0">
                            <tbody>
                                <tr>
                                    <td width="1%"><img height="1" width="4" alt="" src="img/space.gif" /></td>
                                    <td width="39%"><img height="1" width="100" alt="" src="img/space.gif" /></td>
                                    <td width="1%"><img height="1" width="4" alt="" src="img/space.gif" /></td>
                                    <td width="59%"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                                <tr>
                                    <td valign="baseline">
                                    <div align="right"><img height="1" width="1" alt="" src="img/space.gif" /></div>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td valign="baseline"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                                    <td valign="top"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                                <tr>
                                    <td height="1">&nbsp;</td>
                                    <td valign="top">
                                    <table cellspacing="0" cellpadding="0" width="100%" border="0">
                                        <tbody>
                                            <tr>
                                                <td class="Color_18" valign="top" width="7%" bgcolor="#2d669f"><img height="1" width="3" alt="" src="img/space.gif" /></td>
                                                <td class="Color_18" valign="baseline" width="93%"><strong class="Color_18"><strong class="Color_18"><img height="1" width="5" alt="" src="img/space.gif" /></strong>老師</strong></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                    <table cellspacing="0" cellpadding="0" width="100%" border="0">
                                        <tbody>
                                            <tr>
                                                <td valign="top" bgcolor="#2d669f"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                                            </tr>
                                            <tr>
                                                <td valign="top" width="100%"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                    <table cellspacing="0" cellpadding="0" width="100%" border="0">
                                        <tbody>
                                            <tr>
                                                <td valign="top">
                                                <table cellspacing="0" cellpadding="0" width="100%" border="0">
                                                    <tbody>
                                                        <tr>
                                                            <td valign="top" width="100%">
                                                            <span class="Color_18"><strong class="Color_18"><img height="1" width="15" alt="" src="img/space.gif" /><img height="7" width="17" alt="" src="img/freccetta.gif" /></strong></span><a target="_blank" href="http://shih.ms.tku.edu.tw/">時序時</a><a href="http://shih.ms.tku.edu.tw/" title="時光序曲" target="_blank"><img src="img/shih_clock_link.jpg" style="height: 13px; border:0;" /></a><br />
                                                            <br />
                                                            <span class="Color_18"><img height="1" width="15" alt="" src="img/space.gif" /><img height="7" width="17" alt="" src="img/freccetta.gif" /></span><a href="File/DALabInfo/hsu.pdf" target="_blank">徐煥智</a><br /><br />
                                                            <span class="Color_18"><img height="1" width="15" alt="" src="img/space.gif" /><img height="7" width="17" alt="" src="img/freccetta.gif" /></span><a href="File/DALabInfo/chou.pdf" target="_blank">周清江</a><br /><br />
                                                            <span class="Color_18"><img height="1" width="15" alt="" src="img/space.gif" /><img height="7" width="17" alt="" src="img/freccetta.gif" /></span><a href="File/DALabInfo/huang.pdf" target="_blank">黃良志</a><br /><br />
                                                            <span class="Color_18"><img height="1" width="15" alt="" src="img/space.gif" /><img height="7" width="17" alt="" src="img/freccetta.gif" /></span><a href="File/DALabInfo/Jheng.pdf" target="_blank">鄭啟斌</a><br />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td valign="top"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                </td>
                                                <td><img height="170" width="1" alt="" src="img/space.gif" /></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                    <table cellspacing="0" cellpadding="0" width="100%" border="0">
                                        <tbody>
                                            <tr>
                                                <td valign="top" bgcolor="#2d669f"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                            </tr>
                                            <tr>
                                                <td valign="top" width="100%"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                    </td>
                                    <td><img height="170" width="1" alt="" src="img/space.gif" /></td>
                                    <td><img height="1" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                                <tr>
                                    <td valign="top" height="4">
                                    <div align="right"><img height="1" width="1" alt="" src="img/space.gif" /></div>
                                    </td>
                                    <td valign="top">&nbsp;</td>
                                    <td valign="top"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                                    <td valign="top"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                            </tbody>
                        </table>
                        </td>
                        <td valign="top" align="left" width="53%">&nbsp;</td>
                    </tr>
                </tbody>
            </table>
            </td>
            <td valign="top">
            <table height="1" cellspacing="0" cellpadding="0" width="25%" align="center" border="0">
                <tbody>
                    <tr>
                        <td width="1%"><img height="1" width="4" alt="" src="img/space.gif" /></td>
                        <td width="39%"><img height="1" width="100" alt="" src="img/space.gif" /></td>
                        <td width="1%"><img height="1" width="4" alt="" src="img/space.gif" /></td>
                        <td width="59%"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                    </tr>
                    <tr>
                        <td valign="baseline">
                        <div align="right"><img height="1" width="1" alt="" src="img/space.gif" /></div>
                        </td>
                        <td>&nbsp;</td>
                        <td valign="baseline"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                        <td valign="top"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                    </tr>
                    <tr>
                        <td height="1">&nbsp;</td>
                        <td valign="top">
                        <table cellspacing="0" cellpadding="0" width="100%" border="0">
                            <tbody>
                                <tr>
                                    <td class="Color_18" valign="top" width="7%" bgcolor="#66ccff"><img height="1" width="3" alt="" src="img/space.gif" /></td>
                                    <td class="Color_18" valign="baseline" width="93%"><strong class="Color_18"><img height="1" width="5" alt="" src="img/space.gif" />研究生</strong></td>
                                </tr>
                            </tbody>
                        </table>
                        <table cellspacing="0" cellpadding="0" width="100%" border="0">
                            <tbody>
                                <tr>
                                    <td valign="top" bgcolor="#66ccff"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                                <tr>
                                    <td valign="top" width="100%"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                            </tbody>
                        </table>
                        <table cellspacing="0" cellpadding="0" width="100%" border="0">
                            <tbody>
                                <tr>
                                    <td valign="top">
                                    <table cellspacing="0" cellpadding="0" width="100%" border="0">
                                        <tbody>
                                            <tr>
                                                <td valign="top" width="100%">
                                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                  <tr>
                                                    <td>
<% 
While ((Repeat1__numRows <> 0) AND (NOT studentDB.EOF)) 
%>

<span class="Color_18"><img height="1" width="15" alt="space" src="img/space.gif" /><img height="7" width="17" alt="freccetta" src="img/freccetta.gif" />


<a href="research.aspx?class=1&year=<%=(studentDB.Fields.Item("year").Value)%>" target="_parent"><%=(studentDB.Fields.Item("year").Value)%></a></span><br>
</br>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  studentDB.MoveNext()
Wend
%>
</td>
                                                  </tr>
                                              
												  
                                                </table>
                                              
                                              </td>
                                            </tr>
                                            <tr>
                                                <td valign="top"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                    </td>
                                
                                </tr>
                            </tbody>
                        </table>
                        <table cellspacing="0" cellpadding="0" width="100%" border="0">
                            <tbody>
                                <tr>
                                    <td valign="top" bgcolor="#66ccff"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                                <tr>
                                    <td valign="top" width="100%"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                            </tbody>
                        </table>
                        </td>
                        
                     
                    </tr>
                    <tr>
                        <td valign="top" height="4">
                        <div align="right"><img height="1" width="1" alt="" src="img/space.gif" /></div>
                        </td>
                        <td valign="top">&nbsp;</td>
                        <td valign="top"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                        <td valign="top"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                    </tr>
                </tbody>
            </table>
            </td>
            <td valign="top">
            <table height="1" cellspacing="0" cellpadding="0" width="25%" align="center" border="0">
                <tbody>
                    <tr>
                        <td width="1%"><img height="1" width="4" alt="" src="img/space.gif" /></td>
                        <td width="39%"><img height="1" width="100" alt="" src="img/space.gif" /></td>
                        <td width="1%"><img height="1" width="4" alt="" src="img/space.gif" /></td>
                        <td width="59%"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                    </tr>
                    <tr>
                        <td valign="baseline">
                        <div align="right"><img height="1" width="1" alt="" src="img/space.gif" /></div>
                        </td>
                        <td>&nbsp;</td>
                        <td valign="baseline"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                        <td valign="top"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                    </tr>
                    <tr>
                        <td height="1">&nbsp;</td>
                        <td valign="top">
                        <table cellspacing="0" cellpadding="0" width="100%" border="0">
                            <tbody>
                                <tr>
                                    <td class="Color_18" valign="top" width="7%" bgcolor="#4b8ccd"><img height="1" width="3" alt="" src="img/space.gif" /></td>
                                    <td class="Color_18" valign="baseline" style="width: 93%"><strong class="Color_18"><img height="1" width="5" alt="" src="img/space.gif" />大學生</strong></td>
                                </tr>
                            </tbody>
                        </table>
                        <table cellspacing="0" cellpadding="0" width="100%" border="0">
                            <tbody>
                                <tr>
                                    <td valign="top" bgcolor="#4b8ccd"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                                <tr>
                                    <td valign="top" width="100%"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                            </tbody>
                        </table>
                        <table cellspacing="0" cellpadding="0" width="100%" border="0">
                            <tbody>
                                <tr>
                                    <td valign="top">
                                    <table cellspacing="0" cellpadding="0" width="100%" border="0">
                                        <tbody>
                                            <tr>
<td valign="top" width="100%">
<span class="Color_18">


<% 
While ((Repeat2__numRows <> 0) AND (NOT StudentDB_uni.EOF)) 
%>
  <img height="1" width="15" alt="" src="img/space.gif" />
  <img height="7" width="17" alt="" src="img/freccetta.gif" />
  <a href="research.aspx?class=0&year=<%=(StudentDB_uni.Fields.Item("year").Value)%>" target="_parent">
  <%=(StudentDB_uni.Fields.Item("year").Value)%></a><br><br>
  <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  StudentDB_uni.MoveNext()
Wend
%>
</span></td>
                                            </tr>
                                   
                                        </tbody>
                                    </table>
                                    </td>
                                    <td><img height="170" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                            </tbody>
                        </table>
                        <table cellspacing="0" cellpadding="0" width="100%" border="0">
                            <tbody>
                                <tr>
                                    <td valign="top" bgcolor="#4b8ccd"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                                <tr>
                                    <td valign="top" width="100%"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                                </tr>
                            </tbody>
                        </table>
                        </td>
                        <td><img height="170" width="1" alt="" src="img/space.gif" /></td>
                        <td>&nbsp;</td>
                    </tr>
                    <tr>
                        <td valign="top" height="4">
                        <div align="right"><img height="1" width="1" alt="" src="img/space.gif" /></div>
                        </td>
                        <td valign="top">&nbsp;</td>
                        <td valign="top"><img height="1" width="1" alt="" src="img/space.gif" /></td>
                        <td valign="top"><img height="4" width="1" alt="" src="img/space.gif" /></td>
                    </tr>
                </tbody>
            </table>
            </td>
        </tr>
    </tbody>
</table>
</div>
<%
studentDB.Close()
Set studentDB = Nothing
%>
<%
StudentDB_uni.Close()
Set StudentDB_uni = Nothing
%>
