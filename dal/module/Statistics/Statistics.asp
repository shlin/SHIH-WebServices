<%
'#################################################################################
'## 	About Author : 	hANjAN STUDIO (Small Office & Home Office) - WEB DESIGNer.
'##			modified By jAnArA
'##			World Wide Web. http://www.hanjan.biz/
'#################################################################################

Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1


	strSinceDate = "2004, May."	' # You Counter Since Date
	strImageURL = "ShellExt/"
	ImgAddress1 = strImageURL + "StatsBar1.GIF"
	ImgAddress2 = strImageURL + "StatsBar2.GIF"
	ImgAddress3 = strImageURL + "StatsBar3.GIF"
	
'########################################################################################

	strConnString =	"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Statistics.mdb")
	set my_Conn = 	 Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString
	
	strTimeAdjust = 0
	strHomeTimeAdjust = DateAdd("h", strTimeAdjust , Now())
	
'########################################################################################
Response.Write	"<HTML><HEAD>" & vbNewline & vbNewline & vbNewline & _
		"<TITLE>::: Visitors Statistics :::</TITLE>"& vbNewline & vbNewline  & _
		"<META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""TEXT/HTML; CHARSET=BIG5"">" & vbNewline & _
		"<META NAME=""COPYRIGHT"" CONTENT="" (C) Designer in hANjAN STUDIO, Communications"">" & vbNewline & vbNewline 


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
	
	.SF9Pt	{FONT-SIZE: 9pt ; Color: #0077EE ; FONT-FAMILY: Tahoma, Verdana, Arial}
	.SF8Pt	{FONT-SIZE: 8pt ; Color: #3F3F3F ; FONT-FAMILY: Tahoma, Verdana, Arial, metrixpixel}
	.SF7Pt	{FONT-SIZE: 7.5pt ; Color: #3F3F3F ; FONT-FAMILY: Tahoma, Verdana, Arial}
	.SF7Sv	{FONT-SIZE: 7.5pt ; Color: #3F3F3F ; FONT-FAMILY: Tahoma, Verdana, Arial}
	.SF7Bu	{FONT-SIZE: 7.5pt ; Color: #0077EE ; FONT-FAMILY: Tahoma, Verdana, Arial}
	.SF7Rds	{FONT-SIZE: 7.5pt ; Color: #CC0000 ; FONT-FAMILY: Tahoma, Verdana, Arial}
	.SF7Rd	{FONT-SIZE: 7.5pt ; Color: #CC0000 ; FONT-FAMILY: Tahoma, Verdana, Arial}
	.SFTab	{FONT-SIZE: 7.5pt ; Color: #6F6F6F ; Background-color: #F6F6F6; FONT-FAMILY: Arial, Verdana}

--></STYLE>

<BODY topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF">




<TABLE Border=0 Width="100%" CellSpacing=0 CellPadding=0>
<TR><TD style="Padding-right: 10px; Padding-left: 10px" align="CENTER"><BR>

	<TABLE Border=0 Width="650px" CellSpacing=0 CellPadding=0 style="Border: 1px Solid #CCCCCC">
	
	<%
	Dim array_Data_Month(2,6), array_Data_Day(4,30), array_Data_Hr(24), Browser_Data_Max(1,4), array_Language(3,150), array_Data_Total, array_All_Total, Item_Counter
	
	strSql =	"SELECT DISTINCT " & _
		"CALENDAR_LIST, VISITORS_TOTAL, VISITORS_MEMBER, VISITORS_RETURN, USERLOOKOVER, " & _
		"H00, H01, H02, H03, H04, H05, H06, H07, H08, H09, H10, H11, H12, H13, H14, H15, H16, H17, H18, H19, H20, H21, H22, H23 " & _
		"FROM STATISTICS " & _
		"ORDER BY CALENDAR_LIST ASC"
		
		set rsStst = my_Conn.Execute (strSql)
	
	for i = 0 to 5
	
		array_Data_Month(2,i) = mid(DateToStr(DateAdd("m",(i-5),now())),1,4) & " / " & mid(DateToStr(DateAdd("m",(i-5),now())),5,2)
		Do while NOT rsStst.EOF AND Item_Counter < 7
		if mid(rsStst("CALENDAR_LIST"),1,6) = MID(DateToStr(DateAdd("m",(i-5),strHomeTimeAdjust)),1,6) then
		array_Data_Month(1,i) =array_Data_Month(1,i) + rsStst("VISITORS_TOTAL")
		
		end if
		rsStst.MoveNext 
		loop
		
		rsStst.Movefirst
	next
	rsStst.Movefirst
	Item_Counter = 0
	Do while NOT rsStst.EOF
	
		if Item_Counter < 31 AND rsStst("CALENDAR_LIST") > MID(DateToStr(DateAdd("d",(-31),strHomeTimeAdjust)),1,8) then
			array_Data_Day(4,Item_Counter) = rsStst("USERLOOKOVER")
			array_Data_Day(3,Item_Counter) = rsStst("VISITORS_TOTAL")
			array_Data_Day(2,Item_Counter) = rsStst("VISITORS_RETURN")
			array_Data_Day(1,Item_Counter) = rsStst("VISITORS_MEMBER")
			array_Data_Day(0,Item_Counter) = mid(DateToStr(DateAdd("d",(Item_Counter-30),strHomeTimeAdjust)),5,2) & " / " & mid(DateToStr(DateAdd("d",(Item_Counter-30),strHomeTimeAdjust)),7,2)
			array_Data_Total = array_Data_Total + rsStst("VISITORS_TOTAL")
			array_All_Total = array_All_Total  + rsStst("VISITORS_TOTAL")
			
			if rsStst("CALENDAR_LIST") = MID(DateToStr(strHomeTimeAdjust),1,8) then
				for Item_Counter_Hr = 0 to 23
				DbNameTitle = "H" & doublenum(Item_Counter_Hr)
				array_Data_Hr(Item_Counter_Hr) = rsStst(DbNameTitle)
				
					if array_Data_Hr(Item_Counter_Hr) = 0 then 
					array_Data_Hr(Item_Counter_Hr) = 0 ' *** GET VALUE = 0
					End if
				Next
			End if
			Item_Counter = Item_Counter + 1
		Else
			array_All_Total = array_All_Total  + rsStst("VISITORS_TOTAL")
		End if
		
	rsStst.MoveNext 
	loop
	
	
	
	' *** GET MAX VALUE
	GetMonthMaxValue = 0
	GetDayMaxValue = 0
	GetHrMaxValue = 0
	
	for Check_max_Value = 1 to 30
		if array_Data_Day(3,Check_max_Value) > GetDayMaxValue then 
			GetDayMaxValue = array_Data_Day(3,Check_max_Value)
		End if
	Next
	
	for Check_max_Value = 0 to 23
		if array_Data_Hr(Check_max_Value) > GetHrMaxValue then 
			GetHrMaxValue = array_Data_Hr(Check_max_Value)
		End if
	Next
	
	for Check_max_Value = 0 to 5
		if array_Data_Month(1,Check_max_Value) > GetMonthMaxValue then
			GetMonthMaxValue = array_Data_Month(1,Check_max_Value)
		End if
	Next
	'########################################################################################
	%>
	<TR><TD style="Padding-left: 20px; Padding-top: 5px; Padding-bottom: 5px" bgcolor=#F5F5F5 Class=ClsSmallTxt>
		最近 30 瀏覽統計 
		日期範圍 (
		<B><%=array_Data_Day(0,1) %> ~ <%=array_Data_Day(0,30) %></B>
		)
		&nbsp;&nbsp; 小計 (Visitors) : 
		<font color=#EE0000><B><%=FormatNumber(array_Data_Total,0)%></B></font>
		&nbsp;&nbsp;&nbsp;&nbsp; 日平均 :  
		<font color=#EE0000><B><%=int(array_Data_Total/30)%></B></font>
	</TD></TR>
	
	<TR><TD background="ShellExt/Labeldot.gif"><IMG SRC="ShellExt/icon_blank.gif" width=1 height=1></TD></TR>
	
	<!-- ( ///////////////////////////////////////////////////////////// list Date Item I ///////////////////////////////////////////////////////////// ) -->
	<TR><TD><BR>
	<%
	
	
	' ////////// HEAD ///////////////
	
	Response.Write 	"<TABLE Width=100% Border=0 CellSpacing=0 CellPadding=0><TR>" & _
			"<TD style='Padding-left: 10px'>" & vbNewLine & vbNewLine
	
	
	ActiveStat = ""
	ActiveStat = ActiveStat & 	"<TD><TABLE Border=0 CellSpacing=1 CellPadding=1><TR>" & vbNewLine & _
				"<TD Align=Right	Class=SFTab Width=50>		<B>Date</B> 	</TD>" & _
				"<TD 		Class=SFTab Width=120 Height=15>			</TD>" & _
				"<TD Align=Center Class=SFTab Width=35>		Visitor 		</TD>" & _
				"<TD Align=Center Class=SFTab Width=35>		China 		</TD>" & _
				"<TD Align=Center Class=SFTab Width=35>		Return 		</TD>" & _
				"<TD Align=Center Class=SFTab Width=35>		(R) % 		</TD>" & _
				"<TD Align=Right 	Class=SFTab Width=35>		Pages &nbsp;	</TD></TR>" & vbNewLine
	
	For array_Lister=1 to 30
	
		
	
			' ***** LIST DATE TITLE (1)
			ActiveStat = ActiveStat & "<TR><TD Align=Right "
				if array_Lister = 30 then
					ActiveStat = ActiveStat & " Class=SF7Rd>" & vbNewLine
				else
					ActiveStat = ActiveStat & " Class=SFTab>" & vbNewLine
				End if
				if array_Data_Day(0,array_Lister) = "" then array_Data_Day(0,array_Lister) = "&nbsp;" 
			ActiveStat = ActiveStat & array_Data_Day(0,array_Lister) &"&nbsp; </TD>" & vbNewLine
			
			' ***** LIST IMAGE (2)
			ActiveStat = ActiveStat & "<TD Class=SFTab>"
				if array_Data_Day(3,array_Lister) = 0 then
					Get_Image_Width = 1 ' *** GET IMAGE HEIGHT = 1
				else
					Get_Image_Width = int((array_Data_Day(3,array_Lister) *100) / (GetDayMaxValue + 0.5))
					if Get_Image_Width < 1 then Get_Image_Width = 1
				end if	
	
			ActiveStat = ActiveStat & 	"<IMG Align=""absmiddle"" Src=" & ImgAddress1 & " Width=4 Height=16 Border=0>" & _
						"<IMG Align=""absmiddle"" Src=" & ImgAddress2 & " Width="& Get_Image_Width &" Height=16 Border=0>" & _
						"<IMG Align=""absmiddle"" Src=" & ImgAddress3 & " Width=4 Height=16 Border=0></TD>" & vbNewLine
	
			' ***** LIST VISITORS_TOTAL VALUE (3) - (Total)
			' *************************************************************************************************
			NumArrayDataDay = 3
			
			ActiveStat = ActiveStat & "<TD Align=RIGHT " & vbNewLine
				if array_Lister = 30 OR array_Data_Day(NumArrayDataDay,array_Lister) = GetDayMaxValue then
					ActiveStat = ActiveStat & "Class=SF7Rds>" & vbNewLine
				else
					ActiveStat = ActiveStat & "Class=SF7Pt>" & vbNewLine
				End if
				if array_Data_Day(NumArrayDataDay,array_Lister) = 0 then array_Data_Day(NumArrayDataDay,array_Lister) = 0 ' *** GET VALUE = 0
			ActiveStat = ActiveStat & array_Data_Day(NumArrayDataDay,array_Lister)
			ActiveStat = ActiveStat & "</TD>" & vbNewLine
	
	
			' ***** LIST VISITORS_MEMBER VALUE - (Member / China)
			' *************************************************************************************************
			NumArrayDataDay = 1
			ActiveStat = ActiveStat & "<TD Align=RIGHT" & vbNewLine
	
					ActiveStat = ActiveStat & "Class=SF7Bu>" & vbNewLine
	
				if array_Data_Day(NumArrayDataDay,array_Lister) = 0 then array_Data_Day(NumArrayDataDay,array_Lister) = 0 ' *** GET VALUE = 0
			ActiveStat = ActiveStat & array_Data_Day(NumArrayDataDay,array_Lister)
			ActiveStat = ActiveStat & "</TD>" & vbNewLine
	
			' ***** LIST VISITORS_RETURN VALUE (4) - (Return)
			' *************************************************************************************************
			NumArrayDataDay = 2
			ActiveStat = ActiveStat & "<TD Align=RIGHT" & vbNewLine
	
					ActiveStat = ActiveStat & "Class=SFTab>" & vbNewLine
	
				if array_Data_Day(NumArrayDataDay,array_Lister) = 0 then array_Data_Day(NumArrayDataDay,array_Lister) = 0 ' *** GET VALUE = 0
			ActiveStat = ActiveStat & array_Data_Day(NumArrayDataDay,array_Lister)
			ActiveStat = ActiveStat & "</TD>" & vbNewLine		
	
	
			' ***** LIST VISITORS_RETURN VALUE (5) - ( Create IMAGEs )
			' *************************************************************************************************
			ActiveStat = ActiveStat & "<TD Align=RIGHT" & vbNewLine
	
					ActiveStat = ActiveStat & "Class=SFTab>" & vbNewLine
	
				if array_Data_Day(3,array_Lister) = 0 then array_Data_Day(3,array_Lister) = 0.1 ' *** GET VALUE = 0
			ActiveStat = ActiveStat & FormatNumber((array_Data_Day(2,array_Lister) / array_Data_Day(3,array_Lister))*100,0) 
			ActiveStat = ActiveStat & "%</TD>" & vbNewLine	
			
			' ***** LIST VISITORS_RETURN VALUE %  (6)
			ActiveStat = ActiveStat & "<TD Align=RIGHT" & vbNewLine
	
					ActiveStat = ActiveStat & "Class=SF7Bu>" & vbNewLine
	
				if array_Data_Day(4,array_Lister) = 0 then array_Data_Day(4,array_Lister) = 0 ' *** GET VALUE = 0
			ActiveStat = ActiveStat & array_Data_Day(4,array_Lister)
			ActiveStat = ActiveStat & "</TD></TR>" & vbNewLine	
		
		
	
	Next 
	ActiveStat = ActiveStat & "</TABLE></TD>" & vbNewLine
	Response.Write ActiveStat & "</TR></TABLE><BR><BR>"
	%>
	</TD></TR><!-- (Item list End) -->
	
	<TR><TD background="ShellExt/Labeldot.gif"><IMG SRC="ShellExt/icon_blank.gif" width=1 height=1></TD></TR>
	<TR><TD style="Padding-left: 20px; Padding-top: 5px; Padding-bottom: 5px" bgcolor=#F5F5F5 Class=ClsSmallTxt>
		<font color=#EE0000><B><%=array_Data_Day(0,30)%></B></font>
		&nbsp;
		今日時段統計
		&nbsp;&nbsp;
		<FONT COLOR=#0077EE>AM:00~PM:23</FONT>
		&nbsp;&nbsp;
		/
		&nbsp;&nbsp;
		月份統計
	</TD></TR>
	<TR><TD background="ShellExt/Labeldot.gif"><IMG SRC="ShellExt/icon_blank.gif" width=1 height=1></TD></TR>

	<!-- ( ///////////////////////////////////////////////////////////// list Time Item II ///////////////////////////////////////////////////////////// ) -->
	<TR><TD style='Padding-left: 10px'><BR>
	<%
	Response.Write 	"<TABLE Border=0 CellSpacing=1 CellPadding=1><TR><TD>" & vbNewLine & vbNewLine
	
	
	ActiveStat = 		""
	ActiveStat = ActiveStat & 	"<TABLE Border=0 CellSpacing=1 CellPadding=1><TR>" & _
				"<TD Align=Center Class=SFTab Width=50>		<B>Time</B>	&nbsp;</TD>" & _
				"<TD Align=Center Class=SFTab Width=120  Height=15>	 		&nbsp;</TD>" & _
				"<TD Align=Center	Class=SFTab Width=50>		Visitor 		&nbsp;</TD>" & _
				"<TD Width=170 Bgcolor=#FFFFFF></TD></TR>" & vbNewLine
	
	For array_Lister=0 to 23
	
		
	
			' ***** LIST TIME TITLE	
			ActiveStat = ActiveStat & "<TR><TD Width=70 Align=CENTER Class=SFTab>" & vbNewLine
			if array_Lister < 12 then
			TheAMPM = "AM:"
			else
			TheAMPM = "PM:"
			End if
			ActiveStat = ActiveStat & TheAMPM & doublenum(array_Lister) &"</TD>" & vbNewLine
			
			' ***** LIST TIME IMAGE 
			ActiveStat = ActiveStat & "<TD Class=SFTab>"
				if array_Data_Hr(array_Lister) = 0 then
					Get_Image_Width = 1 ' *** GET IMAGE HEIGHT = 1
				else
					Get_Image_Width = int((array_Data_Hr(array_Lister) *100) / (GetHrMaxValue + 0.5))
					if Get_Image_Width < 1 then Get_Image_Width = 1
				end if	
	
			ActiveStat = ActiveStat & 	"<IMG Align=""absmiddle"" Src=" & ImgAddress1 & " Width=4 Height=16 Border=0>" & _
						"<IMG Align=""absmiddle"" Src=" & ImgAddress2 & " Width="& Get_Image_Width &" Height=16 Border=0>" & _
						"<IMG Align=""absmiddle"" Src=" & ImgAddress3 & " Width=4 Height=16 Border=0></TD>" & vbNewLine
	
	
			' ***** LIST TIME VISITORS_TOTAL VALUE
			ActiveStat = ActiveStat & "<TD Align=RIGHT " & vbNewLine
				if array_Data_Hr(array_Lister) = GetHrMaxValue then
					ActiveStat = ActiveStat & "Class=SF7Rd>" & vbNewLine
				else
					ActiveStat = ActiveStat & "Class=SF7Pt>" & vbNewLine
				End if
				
				
			ActiveStat = ActiveStat & array_Data_Hr(array_Lister)
			ActiveStat = ActiveStat & "</TD><TD>&nbsp;</TD></TR>" & vbNewLine
		
		
	
	Next 
	ActiveStat = ActiveStat & "</TABLE>" & vbNewLine
	Response.Write ActiveStat & vbNewLine
	Response.Write 		"<BR><TABLE Border=0 CellSpacing=1 CellPadding=1><TR>" & vbNewLine & _
			
				"<TD Align=Center Class=SFTab Width=50><B>Month</B>	&nbsp;</TD>" & _
				"<TD Align=Center Class=SFTab Width=120  Height=15>	 	&nbsp;</TD>" & _
				"<TD Align=Center	Class=SFTab Width=50>Visitor 		&nbsp;</TD>" & _
				"<TD Width=170 Bgcolor=#FFFFFF></TD></TR>" & vbNewLine



	For array_Lister=0 to 5
	
			' ***** LIST MONTH TITLE	
			Response.Write "<TR><TD Align=CENTER Width=70 Class=SF7Sv> &nbsp;&nbsp;&nbsp;" & vbNewLine
			Response.Write array_Data_Month(2,array_Lister) &"</TD>" & vbNewLine
			
			' ***** LIST IMAGE
			Response.Write"<TD Class=SFTab>"
				if array_Data_Month(1,array_Lister) = 0 then
					Get_Image_Width = 1 ' *** GET IMAGE HEIGHT = 1
				else
					Get_Image_Width = int((array_Data_Month(1,array_Lister) *100) / (GetMonthMaxValue + 0.5))
					if Get_Image_Width < 1 then Get_Image_Width = 1
				end if	
				
			Response.Write 	 	"<IMG Align=""absmiddle"" Src=" & ImgAddress1 & " Width=4 Height=16 Border=0>" & _
						"<IMG Align=""absmiddle"" Src=" & ImgAddress2 & " Width="& Get_Image_Width &" Height=16 Border=0>" & _
						"<IMG Align=""absmiddle"" Src=" & ImgAddress3 & " Width=4 Height=16 Border=0></TD>" & vbNewLine
	
			' ***** LIST VISITORS_TOTAL VALUE
			Response.Write"<TD Align=RIGHT Class=SF7Pt>" & vbNewLine
				if array_Data_Month(1,array_Lister) = 0 then array_Data_Month(1,array_Lister) = 0 ' *** GET VALUE = 0
			Response.Write array_Data_Month(1,array_Lister)
			Response.Write "</TD><TD>&nbsp;</TD></TR>" & vbNewLine
		
	
	
	Next 
		Response.Write "</TABLE>" & vbNewLine
	Response.Write "</TD></TR></TABLE><BR><BR>"
	%>
	</TD></TR>
	
	
	<!-- ( ///////////////////////////////////////////////////////////// list System / Browser Item III ///////////////////////////////////////////////////////////// ) -->
	
	<%
			strSql ="SELECT BROWSER_LIST.BROWSER_KEY FROM BROWSER_LIST WHERE AUTO_ID = 1"
			set rsStst = my_Conn.Execute (strSql)
			Last_Update_Time = rsStst("BROWSER_KEY")
			if Mid(DateToStr(strHomeTimeAdjust),1,8) > Mid(Last_Update_Time,1,8) or Request("update") = "now" then
				Call Update_Browser_Data
				Response.redirect "Statistics.asp"
			End if
			
	
			for i = 1 to 4
				Browser_temp_Max = 0
				Browser_temp_Total = 0
				strSql ="SELECT * FROM BROWSER_LIST WHERE BROWSER_TYPE = " & i & " AND BROWSER_SET = True"
				strSql = strSql & " ORDER BY BROWSER_LIST ASC"
				set rsStst = my_Conn.Execute (strSql)
				Do while NOT rsStst.EOF
				if rsStst("BROWSER_COUNT") > Browser_temp_Max then Browser_temp_Max = rsStst("BROWSER_COUNT")
				Browser_temp_Total = Browser_temp_Total + rsStst("BROWSER_COUNT")
				rsStst.MoveNext 
				loop
			Browser_Data_Max(1,i) = Browser_temp_Max + 0.1
			Browser_Data_Max(0,i) = Browser_temp_Total + 0.1
			Next
	%>
	
	<TR><TD background="ShellExt/Labeldot.gif"><IMG SRC="ShellExt/icon_blank.gif" width=1 height=1></TD></TR>
	<TR><TD style="Padding-left: 20px; Padding-top: 5px; Padding-bottom: 5px" bgcolor=#F5F5F5 Class=ClsSmallTxt>
		<FONT COLOR=#0077EE>Search WebRobot</FONT>
		&nbsp;&nbsp;
		搜尋引擎瀏覽分類數據 / 分析 (含最近訪問日期)
	</TD></TR>
	<TR><TD background="ShellExt/Labeldot.gif"><IMG SRC="ShellExt/icon_blank.gif" width=1 height=1></TD></TR>
	
	
	<TR><TD Align=CENTER><BR>
	
		
	
			<TABLE Border=0 CellSpacing=1 CellPadding=1>
			<%
			ActiveStat = ""
			for i = 1 to 4
			
				strSql ="SELECT * FROM BROWSER_LIST WHERE BROWSER_TYPE = " & i & " AND BROWSER_SET = True"
				strSql = strSql & " ORDER BY BROWSER_LIST ASC"
				set rsStst = my_Conn.Execute (strSql)
				Do while NOT rsStst.EOF
	
				if rsStst("BROWSER_COUNT") = 0 then
					Get_Image_Width = 1 ' *** GET IMAGE HEIGHT = 1
				else
					Get_Image_Width = int((rsStst("BROWSER_COUNT") *100) / (Browser_Data_Max(1,i) + 0.5))
					if Get_Image_Width < 1 then Get_Image_Width = 1
				end if	
				
				ActiveStat = ActiveStat & _
					 	"<TR><TD Align=RIGHT Class=SF7Pt>" & rsStst("BROWSER_NAME") & "&nbsp;&nbsp;&nbsp; </TD>" & vbNewLine & _
						"<TD Align=RIGHT Class=SF7Rd>" & FormatNumber(rsStst("BROWSER_COUNT"),0) & "&nbsp;&nbsp;&nbsp; </TD>" & _
						"<TD Class=SFTab>" & vbNewLine & _
						"<IMG Align=""absmiddle"" Src=" & ImgAddress1 & " Width=4 Height=16 Border=0>" & _
						"<IMG Align=""absmiddle"" Src=" & ImgAddress2 & " Width="& Get_Image_Width &" Height=16 Border=0>" & _
						"<IMG Align=""absmiddle"" Src=" & ImgAddress3 & " Width=4 Height=16 Border=0>" & _
						"&nbsp;" & FormatNumber((rsStst("BROWSER_COUNT") / Browser_Data_Max(0,i))*100,1) & " % &nbsp; </TD>" & vbNewLine & _	
						"<TD Align=RIGHT Class=SF7Bu>&nbsp; " & ChkDate(rsStst("BROWSER_DATE"),"",f) & "</TD>" & _
						"</TR>" & vbNewLine & vbNewLine
				rsStst.MoveNext 
				loop
	
				
				if i = 1 then 
					Response.Write ActiveStat
					ActiveStat = ""
					Response.Write 	"<TR><TD colSpan=""3""><BR>&nbsp;</TD></TR>" & vbNewLine & _
							"</TABLE>" & vbNewLine & _
							"</TD></TR>" & vbNewLine & vbNewLine & _
							"<TR><TD background=""ShellExt/Labeldot.gif""><IMG SRC=""ShellExt/icon_blank.gif"" width=1 height=1></TD></TR>" & vbNewLine & _
							"<TR><TD style=""Padding-left: 20px; Padding-top: 5px; Padding-bottom: 5px"" bgcolor=#F5F5F5 Class=ClsSmallTxt>" & vbNewLine & _
							"	<FONT COLOR=#0077EE>Options System</FONT>" & vbNewLine & _
							"	&nbsp;&nbsp;" & vbNewLine & _
							"	作業平台系統分類數據 / 分析 (含最近訪問日期)" & vbNewLine & _
							"</TD></TR>" & vbNewLine & vbNewLine & _
							"<TR><TD background=""ShellExt/Labeldot.gif""><IMG SRC=""ShellExt/icon_blank.gif"" width=1 height=1></TD></TR>" & vbNewLine & _
							"<TR><TD Align=CENTER><BR>" & vbNewLine & _
							"<TABLE Border=0 CellSpacing=1 CellPadding=1>" & vbNewLine

			
				elseif i = 3 then
				
					Response.Write ActiveStat
					ActiveStat = ""
					Response.Write 	"<TR><TD colSpan=""3""><BR>&nbsp;</TD></TR>" & vbNewLine & _
							"</TABLE>" & vbNewLine & _
							"</TD></TR>" & vbNewLine & vbNewLine & _
							"<TR><TD background=""ShellExt/Labeldot.gif""><IMG SRC=""ShellExt/icon_blank.gif"" width=1 height=1></TD></TR>" & vbNewLine & _
							"<TR><TD style=""Padding-left: 20px; Padding-top: 5px; Padding-bottom: 5px"" bgcolor=#F5F5F5 Class=ClsSmallTxt>" & vbNewLine & _
							"	<FONT COLOR=#0077EE>Browser system / Plugin</FONT>" & vbNewLine & _
							"	&nbsp;&nbsp;" & vbNewLine & _
							"	瀏覽器及外掛等分類數據 / 分析 (含最近訪問日期)" & vbNewLine & _
							"</TD></TR>" & vbNewLine & vbNewLine & _
							"<TR><TD background=""ShellExt/Labeldot.gif""><IMG SRC=""ShellExt/icon_blank.gif"" width=1 height=1></TD></TR>" & vbNewLine & _
							"<TR><TD Align=CENTER><BR>" & vbNewLine & _
							"<TABLE Border=0 CellSpacing=1 CellPadding=1>" & vbNewLine
				
				end if
			
			Next
			Response.Write ActiveStat
			%>
			</TABLE><BR><BR><BR>
	
		
		
	</TD></TR><!-- (Browser Item list End) -->

	<TR><TD background="ShellExt/Labeldot.gif"><IMG SRC="ShellExt/icon_blank.gif" width=1 height=1></TD></TR>
	<TR><TD style="Padding-left: 20px; Padding-top: 5px; Padding-bottom: 5px" bgcolor=#F5F5F5 Class=ClsSmallTxt>
		<FONT COLOR=#0077EE>Browser / Language</FONT>
		&nbsp;&nbsp;
		瀏覽器主要語系分類 / 統計
	</TD></TR>
	<TR><TD background="ShellExt/Labeldot.gif"><IMG SRC="ShellExt/icon_blank.gif" width=1 height=1></TD></TR>


	<TR><TD Align=CENTER><BR>
		<%
			array_Lister = 1
			total_language = 0
			strSql ="SELECT * FROM BROWSER_LANGUAGE WHERE L_COUNT > 0"
			strSql = strSql & " ORDER BY L_E_NAME ASC"
			set rsStst = my_Conn.Execute (strSql)
			Do while NOT rsStst.EOF
			array_Language(1,array_Lister) = rsStst("L_E_NAME")
			array_Language(2,array_Lister) = rsStst("L_C_NAME")
			array_Language(3,array_Lister) = rsStst("L_COUNT")
			total_language = total_language + rsStst("L_COUNT")
			array_Lister = array_Lister + 1
			rsStst.MoveNext 
			loop
		%>
	
	
			<TABLE Border=0 CellSpacing=0 CellPadding=1>
			<%
			ActiveStat = ""
			FOR i = 1 to array_Lister -1
			ActiveStat = ActiveStat & 	"	<TR>" & _
						"	<TD Align=RIGHT Class=SF8Pt> &nbsp;&nbsp; " & array_Language(1,i) & "&nbsp;</TD>" & _
						"	<TD Class=SF8Pt>&nbsp;" & array_Language(2,i) & "&nbsp;&nbsp;</TD>" & _
						"	<TD Align=RIGHT Class=SF7Rd>" & array_Language(3,i) & "&nbsp;</TD>" & _
						"	<TD Align=RIGHT Class=SF7Sv>" & FormatNumber((array_Language(3,i) / total_language)*100,0) & " %&nbsp;</TD></TR>" & vbNewLine
			NEXT
			Response.Write ActiveStat
			%>
			</TABLE><BR><BR>
	
	
	
	
	
	</TD></TR><!-- (Language Item list End) -->
	
	
	<TR><TD background="ShellExt/Labeldot.gif"><IMG SRC="ShellExt/icon_blank.gif" width=1 height=1></TD></TR>
	<TR><TD Align=CENTER Class=SF8Pt>
	<BR>
	總瀏覽計數統計 
	( <small>TOTAL</small> ) : <FONT Class=SF9Pt><Big><B><% =FormatNumber(array_All_Total,0) %></B></Big></FONT>
	<br><br>
	統計啟動日期
	<b><%=strSinceDate%></b>
	&nbsp; / &nbsp;
	最後更新 <%=ChkDate(Last_Update_Time,"&nbsp;",true)%>
	&nbsp; / &nbsp;
	<a Title=" 更新統計數據 " href="Statistics.asp?update=now">Update</a> Now !
	<BR><BR></TD></TR>
	</TABLE>
	<p>&nbsp;</p></TD>
</TR>
</TABLE>



<% Sub Update_Browser_Data()
'**************************************************************************************************
	for Get_Browser_Type = 1 to 4

		strSql =	"SELECT BROWSER_LIST.AUTO_ID, BROWSER_LIST.BROWSER_KEY FROM BROWSER_LIST " & _
			"WHERE BROWSER_TYPE = "& Get_Browser_Type &" AND BROWSER_SET = True " & _
			"ORDER BY BROWSER_LIST ASC"
		set rsStst = my_Conn.Execute (strSql)

			Do while NOT rsStst.EOF
				
				Total_Counter = 0
				Last_Record_Date = ""

				strSql =	"SELECT BROWSER_RECORD.REC_COUNT, BROWSER_RECORD.REC_DATE FROM BROWSER_RECORD " & _
					"WHERE Instr(BROWSER_NAME,'" &  rsStst("BROWSER_KEY") & "') > 0 " & _
					"AND REC_COUNT > 0 ORDER BY REC_DATE"
				set rsGetData = my_Conn.Execute (strSql)
				Do while NOT rsGetData.EOF
					Total_Counter = Total_Counter + rsGetData("REC_COUNT")
					Last_Record_Date = rsGetData("REC_DATE")
					rsGetData.MoveNext 
				loop
				
				strSql = "UPDATE BROWSER_LIST SET "
				strSql = strSql & " BROWSER_COUNT = " & Total_Counter & ", "
				strSql = strSql & " BROWSER_DATE = '" & Last_Record_Date & "'"
				strSql = strSql & " WHERE AUTO_ID = " & rsStst("AUTO_ID")
				my_Conn.Execute (strSql)
				
			rsStst.MoveNext 
			loop

	Next	

		strSql = "UPDATE BROWSER_LIST SET "
		strSql = strSql & " BROWSER_KEY = '" & DateToStr(strHomeTimeAdjust) & "'"
		strSql = strSql & " WHERE AUTO_ID = 1"
		my_Conn.Execute (strSql)
		Last_Update_Time = DateToStr(strHomeTimeAdjust)
end sub
'**************************************************************************************************
%>

<%
'*******************************************************************************************************************
		my_Conn.Close
		set my_Conn = nothing	
Response.Write	"</BODY></HTML>" & vbNewline
%>