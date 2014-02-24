<%
'#################################################################################
'## 	About Author : 	hANjAN STUDIO (Small Office & Home Office) - WEB DESIGNer.
'##			modified By jAnArA
'##			World Wide Web. http://www.hanjan.biz/
'#################################################################################

Response.Write	"<HTML><HEAD>" & vbNewline & vbNewline & vbNewline & _
		"<TITLE>Statistics Record Delete</TITLE>"& vbNewline & vbNewline  & _
		"<META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""TEXT/HTML; CHARSET=BIG5"">" & vbNewline & _
		"<META NAME=""COPYRIGHT"" CONTENT="" (C) Designer : hANjAN STUDIO""></HEAD><BORD>" & vbNewline & vbNewline
		
		
		
		%><!--#INCLUDE FILE="inc_function.asp" --><%
		'#########################################################################
		'# 	STATISTICS on USER OS SYSTEM COUNTER ## last Edit by JB 2004-08-01 #######
		'#########################################################################
				
		
		
		' ##### strConnString SET #####
		strConnString =	"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Statistics.mdb")
		set my_Conn = 	Server.CreateObject("ADODB.Connection")
		my_Conn.Open strConnString
		'#########################################################################
		strUserLanguage = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
		strUserBorwser = Trim(Request.Servervariables("HTTP_USER_AGENT"))
		strTimeAdjust = 0
		strHomeTimeAdjust = DateAdd("h", strTimeAdjust , Now())
		'#########################################################################



		strSql = "UPDATE BROWSER_RECORD SET "
		strSql = strSql & "  REC_COUNT = 0"
		strSql = strSql & " ,REC_DATE = '20030630000000'"
		strSql = strSql & " WHERE BROWSER_ID > 0"
		my_conn.Execute (strSql)

		strSql = "UPDATE BROWSER_LIST SET "
		strSql = strSql & "  BROWSER_COUNT = 0"
		strSql = strSql & " ,BROWSER_DATE = '20030630000000'"
		strSql = strSql & " WHERE AUTO_ID > 0"
		my_conn.Execute (strSql)

		' TYPE = 1
		strSql = "UPDATE BROWSER_LIST SET "
		strSql = strSql & "  BROWSER_COUNT = 1"
		strSql = strSql & " WHERE AUTO_ID = 23"
		my_conn.Execute (strSql)

		' TYPE = 2
		strSql = "UPDATE BROWSER_LIST SET "
		strSql = strSql & "  BROWSER_COUNT = 1"
		strSql = strSql & " WHERE AUTO_ID = 17"
		my_conn.Execute (strSql)
		
		' TYPE = 3
		strSql = "UPDATE BROWSER_LIST SET "
		strSql = strSql & "  BROWSER_COUNT = 1"
		strSql = strSql & " WHERE AUTO_ID = 19"
		my_conn.Execute (strSql)
		
		' TYPE = 4
		strSql = "UPDATE BROWSER_LIST SET "
		strSql = strSql & "  BROWSER_COUNT = 1"
		strSql = strSql & " WHERE AUTO_ID = 15"
		my_conn.Execute (strSql)						

		strSql = "UPDATE BROWSER_LANGUAGE SET "
		strSql = strSql & "  L_COUNT = 0"
		strSql = strSql & " WHERE L_ID > 0"
		my_conn.Execute (strSql)

		strSql ="DELETE FROM STATISTICS WHERE CALENDAR_LIST <> """""
		my_conn.Execute (strSql)
		
		
		FOR i = 0 to 31
		
			strSql = "INSERT INTO STATISTICS ("
			strSql = strSql & " CALENDAR_LIST"
			strSql = strSql & ", VISITORS_TOTAL"
			strSql = strSql & ", VISITORS_MEMBER"
			strSql = strSql & ", VISITORS_RETURN"
			strSql = strSql & ", H00"
			strSql = strSql & ") VALUES ("
			strSql = strSql & "'" & Mid(DateToStr(DateAdd("d", i - 31 , chkDate(DateToStr(strHomeTimeAdjust),"",f))),1,8) & "'"
			strSql = strSql & ", " & 0 & ", " & 0 & ", " & 0 & ", " & 0
			strSql = strSql & ")"
			my_Conn.Execute(strSql)
		NEXT

		strTodayValue = left(DateToStr(strHomeTimeAdjust),8)
		strHoursValue = mid(DateToStr(strHomeTimeAdjust),9,2)
		
		strSql = "UPDATE STATISTICS SET "
		strSql = strSql & " VISITORS_TOTAL = 1 "				
		strSql = strSql & ", H" & strHoursValue & " = 1 "
		strSql = strSql & " WHERE CALENDAR_LIST= '" & strTodayValue & "'"
		my_Conn.Execute(strSql)
		
		
		Response.redirect "Statistics.asp"
%>
