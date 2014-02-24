<%
'#################################################################################
'## 	About Author : 	hANjAN STUDIO (Small Office & Home Office) - WEB DESIGNer.
'##			modified By jAnArA
'##			World Wide Web. http://www.hanjan.biz/
'##			Customer Relationship Management
'#################################################################################

	' 資料庫連線
	strConnString =	"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Statistics.mdb")
	set my_Conn = 	 Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString
	%>
	<!--#INCLUDE FILE="inc_function.asp" -->
	<%
	'///////////////////////// SUBTOTAL PAGE ///////////////////////////////
	Default_Of_Number = 100 '#### DEFAULT DISPLAY NUMBER
	
	If Request.QueryString("CursorsPage") = "" or Request.QueryString("CursorsPage") = 0 then
		CursorsPage = 1
	Else
		CursorsPage = CINT(Request.QueryString("CursorsPage"))
	End if 
	
	strSql = 	"SELECT COUNT(IP_NUMBER) FROM IP_RECORD " & _
		"WHERE IP_COUNTER > 0 "
	set rsCount = my_Conn.Execute (strSql)
	Sum_Of_Total_Counter = rsCount(0)
	rsCount.Close
	
	if  int(Sum_Of_Total_Counter / Default_Of_Number) = (Sum_Of_Total_Counter / Default_Of_Number) then
		theTotalPages = int(Sum_Of_Total_Counter / Default_Of_Number)
	Else
		theTotalPages = int(Sum_Of_Total_Counter / Default_Of_Number) +1
	End if
	
	If CursorsPage > theTotalPages Then
		CursorsPage = theTotalPages
	End if 

	'########################################################################################
	Response.Write	"<HTML><HEAD>" & vbNewline & vbNewline & vbNewline & _
			"<TITLE>::: Visitors Statistics IP LIST :::</TITLE>"& vbNewline & vbNewline  & _
			"<META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""TEXT/HTML; CHARSET=BIG5"">" & vbNewline & _
			"<META NAME=""COPYRIGHT"" CONTENT="" (C) Designer in hANjAN STUDIO, Communications"">" & vbNewline & vbNewline 
	'#################################################################################




%>
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
<TR><TD style="Padding-right: 10px; Padding-left: 10px" align="CENTER"><BR>

	<TABLE Border=0 Width="650px" CellSpacing=0 CellPadding=0 style="Border: 1px Solid #CCCCCC">
	<TR><TD>


<%




	Dim strQueryStringOrderValue, strOrderName
	strQueryStringOrderValue = Ucase(Request.QueryString("Order"))

	If strQueryStringOrderValue = "SINCE" then
	strOrderName = "IP_SINCE_DATE"
	ElseIf strQueryStringOrderValue = "LAST" then
	strOrderName = "IP_LAST_DATE"
	ElseIf strQueryStringOrderValue = "HITS" then
	strOrderName = "IP_COUNTER"
	Else
	strOrderName = "IP_NUMBER"
	End If

	strSql =	"SELECT * FROM IP_RECORD " & _
		"ORDER BY " & strOrderName & " DESC"
	set rsAdmin = my_Conn.Execute (strSql)
	
	Response.Write 	"<TABLE Width=99% Border=0 CellSpacing=0 CellPadding=0>"
	Response.Write 	"<TR BGCOLOR=#DFDFDF><TD colSpan=5 Class=ClsSmallTxt>"
	Response.Write 	"&nbsp; ORDER BY "
	Response.Write 	"</TD></TR>"
	Response.Write 	"<TR BGCOLOR=#EFEFEF>"
	Response.Write 	"<TD Class=ClsSmallTxt>&nbsp; <A HREF=""VisitorsIP.asp?Version=Demo&Order=IP""><B>IP</B></A></TD>"
	Response.Write 	"<TD Class=ClsSmallTxt><A HREF=""VisitorsIP.asp?Version=Demo&Order=Since""><B>SINCE_DATE</B></A></TD>"
	Response.Write 	"<TD Class=ClsSmallTxt><A HREF=""VisitorsIP.asp?Version=Demo&Order=Last""><B>LAST_DATE</B></A></TD>"
	Response.Write 	"<TD Class=ClsSmallTxt>LANGUAGE</TD>"
	Response.Write 	"<TD Class=ClsSmallTxt><A HREF=""VisitorsIP.asp?Version=Demo&Order=Hits""><B>HITS</B></A></TD>"
	Response.Write 	"</TR>"
	
	Lister_Of_Pages = 1
	Do while NOT rsAdmin.EOF
	
	if Lister_Of_Pages <= (Default_Of_Number * CursorsPage) AND Lister_Of_Pages > (Default_Of_Number * (CursorsPage-1)) then
	
		Response.Write 	"<TR vAlign=top>" & vbNewLine
		Response.Write 	"<TD Style=""Border-bottom: 1px Solid #DFDFDF"" Class=ClsMidTxt>&nbsp; <font color=#3E8BFA><b>" & rsAdmin("IP_NUMBER") & "</b></font></TD>" & vbNewLine
		Response.Write 	"<TD Style=""Border-bottom: 1px Solid #DFDFDF"" Class=ClsSmallTxt>" & ChkDate(rsAdmin("IP_SINCE_DATE"),"<br/>",true) & "</TD>" & vbNewLine
		Response.Write 	"<TD Style=""Border-bottom: 1px Solid #DFDFDF"" Class=ClsSmallTxt>" & ChkDate(rsAdmin("IP_LAST_DATE"),"<br/>",true) & "</TD>" & vbNewLine
		Response.Write 	"<TD Style=""Border-bottom: 1px Solid #DFDFDF"" Class=ClsSmallTxt>" & rsAdmin("IP_LANGUAGE") & "&nbsp;</TD>" & vbNewLine
		Response.Write 	"<TD Style=""Border-bottom: 1px Solid #DFDFDF"" Class=ClsSmallTxt Align=right><font color=#CC0000><b>" & rsAdmin("IP_COUNTER") & "</b></font> &nbsp;&nbsp;&nbsp;</TD>" & vbNewLine
		Response.Write 	"</TR>" & vbNewLine

	end if
	Lister_Of_Pages = Lister_Of_Pages + 1

	rsAdmin.MoveNext 
	loop

	Response.Write	"<TR BgColor=""#EEEEEE""><TD style=""Padding-top: 3px; Padding-bottom: 3px"" Colspan=""5"" Class=ClsMidTxt>&nbsp;&nbsp;" & vbNewLine

		if theTotalPages > 1 then
		
		strHrefUrl = "VisitorsIP.asp?Version=Demo&Order=" & Request.QueryString("Order") & "&CursorsPage="
		
		Response.Write	"	分頁 : &nbsp;"

			For theCursorsPage=1 to theTotalPages
				if theCursorsPage = CursorsPage then
				Response.write"<font color=#0000AA><b>["&theCursorsPage&"]" & "</b></font>&nbsp;"
				else
				Response.write"<a TITLE="" LIST : " & INT((theCursorsPage - 1) * Default_Of_Number) +1  &  " ~ " & theCursorsPage * Default_Of_Number & " "" href=""" & strHrefUrl & theCursorsPage & """>["&theCursorsPage&"]" & "</a>&nbsp;"
				end if
			Next
			
	
		else
		Response.Write	"&nbsp;&nbsp;"
		end if

	Response.Write	"&nbsp;&nbsp;&nbsp; 總筆數 : <B>" & Sum_Of_Total_Counter & "</B></TD></TR>" & vbNewLine

	Response.Write 	"</TABLE>" & vbNewLine
%>

		</TD></TR>
		</TABLE>
	</TD></TR>
	</TABLE>
</BODY></HTML>