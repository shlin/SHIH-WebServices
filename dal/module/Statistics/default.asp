<%
'#################################################################################
'## 	About Author : 	hANjAN STUDIO (Small Office & Home Office) - WEB DESIGNer.
'##			modified By jAnArA
'##			World Wide Web. http://www.hanjan.biz/
'#################################################################################

	' ��Ʈw�s�u
	strConnString =	"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Statistics.mdb")
	set my_Conn = 	 Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString

	' ����ܼƳ]�w
	strTimeAdjust = 0
	strHomeTimeAdjust = 	DateAdd("h", strTimeAdjust , Now())
	strOnlineUserIP = 		Request.ServerVariables("REMOTE_ADDR")
	strUserLanguage = 	Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
	strUserBorwser = 		Trim(Request.ServerVariables("HTTP_USER_AGENT"))
	strApprovalValue = 	&HDE000 & &HDC000
	strCookieMark = 		"n08m01t"	' YOUR SIDE MARK
	
'########################################################################################

Response.Write	"<HTML><HEAD>" & vbNewline & vbNewline & vbNewline & _
		"<TITLE>::: Statistics Demo :::</TITLE>"& vbNewline & vbNewline  & _
		"<META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""TEXT/HTML; CHARSET=BIG5"">" & vbNewline & _
		"<META NAME=""COPYRIGHT"" CONTENT="" (C) Designer in hANjAN STUDIO, Communications"">" & vbNewline & vbNewline 



	' ~~~~~ �o�̨Ҧp�O�z���������e ~~~~~
%>

<!--#INCLUDE FILE="inc_function.asp" -->

<STYLE type=text/css><!-- // ONLY for STATISTICS
	/* BASIC ACTIVE LINK DECORATION AND FONT COLOR */
	A:link	{Color:#EE0000;	TEXT-DECORATION: underline}
	A:visited 	{Color:#EE0000;	TEXT-DECORATION: underline}
	A:hover	{Color:#EE0000;	TEXT-DECORATION: none}
	A:active	{Color:#EE0000;	TEXT-DECORATION: none}
	
	/* BASIC TABLE FONT COLOR AND SIZE */
	 td			{FONT-SIZE: 11px; 
	 				line-HEIGHT: 12pt; 
	 				FONT-FAMILY: Verdana, Tahoma, Arial; 
	 				Color: #3F3F3F; line-height: 150%
	 				}
--></STYLE>

<BODY topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF">
<TABLE Border=0 Width="100%" CellSpacing=0 CellPadding=0>
<TR><TD style="Padding-right: 10px; Padding-left: 10px" align="CENTER"><BR><BR><BR>

		<TABLE BORDER="0" WIDTH="500" CELLPADDING=3 CELLSPACING="1" BGCOLOR="#BFBFBF">
		<TR>
		<TD BGCOLOR="#EEEEEE"><A HREF="default.asp">�����s���έp</A></TD>
		<TD BGCOLOR="#EEEEEE">Statistics V14</TD>
		</TR>
		<TR>
		<TD COLSPAN="2" BGCOLOR="#CFCFCF">
		<BR><BR><BR>
		</TD>
		</TR>
		<TR>
		<TD BGCOLOR="#EFEFEF">2004-08</TD>
		<TD BGCOLOR="#EFEFEF">
		<A HREF="VisitorsIP.asp">�d�� IP �C��</A>
		&nbsp;&nbsp;&nbsp;
		<A HREF="Statistics.asp">�d���s���έp</A>
		&nbsp;&nbsp;&nbsp;
		<A HREF="delete_statistics.asp">�R�������έp</A>
		</TD>
		</TR>
		</TABLE>

</TD></TR>
</TABLE>


<%	' �H�U�K�ܱz��������, �`�N�ݭn�O .asp ������
	' #################################
	' ### 	�p�G�X�Ȥ��Ѥw�g�X�ݹL,�h���p�� !!!
	' ### 	���q��k�i�ܩ󭺭���,���|���Ʋέp
	' #################################


	If mid(DateToStr(strHomeTimeAdjust),1,8) <> mid(Request.Cookies(strCookieURL & strCookieMark & "LAST_HERE"),1,8) then %>
	
		<!--#INCLUDE FILE="inc_counter.asp" -->
	
<% 	End if %>