<%
'#################################################################################
'## 	About Author : 	hANjAN STUDIO - WEB DESIGNer.
'##			modified By jAnArA
'##			World Wide Web. http://www.hanjan.biz/
'#################################################################################

		Response.Cookies(strCookieURL & strCookieMark & "LAST_HERE") = DateToStr(strHomeTimeAdjust)
		Response.Cookies(strCookieURL & strCookieMark & "LAST_HERE").Expires = dateAdd("m", 24, strHomeTimeAdjust)
		
		if Request.Cookies(strCookieURL & strCookieMark & "SINCE_LOGIN") = "" then
		
			' ***** SINCE LOGIN DATE *****
			Response.Cookies(strCookieURL & strCookieMark & "SINCE_LOGIN") = DateToStr(strHomeTimeAdjust)
			Response.Cookies(strCookieURL & strCookieMark & "SINCE_LOGIN").Expires = dateAdd("m", 24, strHomeTimeAdjust)
		End if

		
		' ####### THE RECORD USER IP ###########################################
				
			strSql = "SELECT IP_RECORD.IP_NUMBER" 
			strSql = strSql & " FROM IP_RECORD "
			strSql = strSql & " WHERE IP_NUMBER = '" & strOnlineUserIP & "' AND IP_LANGUAGE = '" & strUserLanguage & "'"
			set rsAgent = my_Conn.Execute (strSql)
			if rsAgent.EOF then
				strSql = "INSERT INTO IP_RECORD ("
				strSql = strSql & " IP_NUMBER"
				strSql = strSql & ", IP_SINCE_DATE"
				strSql = strSql & ", IP_LAST_DATE"
				strSql = strSql & ", IP_LANGUAGE"
				strSql = strSql & ") VALUES ("
				strSql = strSql & "'" & strOnlineUserIP & "'"
				strSql = strSql & ", '" & DateToStr(strHomeTimeAdjust) & "'"
				strSql = strSql & ", '" & DateToStr(strHomeTimeAdjust) & "'"
				strSql = strSql & ", '" & strUserLanguage & "'"
				strSql = strSql & ")"
				my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
			else
				strSql = "UPDATE IP_RECORD SET "
				strSql = strSql & "  IP_COUNTER = IP_COUNTER + 1 "
				strSql = strSql & " ,IP_LAST_DATE = '" & DateToStr(strHomeTimeAdjust) & "'"
				strSql = strSql & " WHERE IP_NUMBER = '" & strOnlineUserIP & "' AND IP_LANGUAGE = '" & strUserLanguage & "'"
				my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
			end if

		' ####### THE RECORD USER OS SYSTEM IS UNKNOWN ###########################################
				
			strSql = "SELECT BROWSER_RECORD.BROWSER_NAME" 
			strSql = strSql & " FROM BROWSER_RECORD "
			strSql = strSql & " WHERE BROWSER_NAME = '" & strUserBorwser & "'"
			set rsAgent = my_Conn.Execute (strSql)
			if rsAgent.EOF then
				strSql = "INSERT INTO BROWSER_RECORD ("
				strSql = strSql & " BROWSER_NAME"
				strSql = strSql & ", REC_DATE"
				strSql = strSql & ", REC_COUNT"
				strSql = strSql & ") VALUES ("
				strSql = strSql & "'" & strUserBorwser & "'"
				strSql = strSql & ", '" & DateToStr(strHomeTimeAdjust) & "'"
				strSql = strSql & ", 1"
				strSql = strSql & ")"
				my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
			else
				strSql = "UPDATE BROWSER_RECORD SET "
				strSql = strSql & "  REC_COUNT = REC_COUNT + 1 "
				strSql = strSql & " ,REC_DATE = '" & DateToStr(strHomeTimeAdjust) & "'"
				strSql = strSql & " WHERE BROWSER_NAME= '" & strUserBorwser & "'"
				my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
			end if

		
		' ####### THE RECORD USER LANGUAGE IS UNKNOWN	###########################################
			
			strAcceptLanguage = strUserLanguage
			check_value1 = InStr(1, strAcceptLanguage, ",")
				If check_value1 > 0 Then
					strAcceptLanguage = Left(strAcceptLanguage, check_value1 - 1)
				End If
			check_value2 = InStr(1, strAcceptLanguage, ";")
				If check_value2 > 0 Then
					strAcceptLanguage = Left(strAcceptLanguage, check_value2 - 1)
				End If

			strSql = 	" SELECT BROWSER_LANGUAGE.L_CODE" & _
				" FROM BROWSER_LANGUAGE " & _
				" WHERE L_CODE = '" & strAcceptLanguage & "'" & _
				" ORDER BY L_CODE DESC"
			set rsLanguage = my_Conn.Execute (strSql)
			
				if NOT rsLanguage.EOF then
				
					strSql = 	" UPDATE BROWSER_LANGUAGE SET " & _
						" L_COUNT = L_COUNT + 1" & _
						" WHERE L_CODE = '" & strAcceptLanguage & "'"
					my_Conn.Execute (strSql)
				else
					' ***  RECORD * unknown   µLªk¿ëÃÑ
					strSql = 	" UPDATE BROWSER_LANGUAGE SET " & _
						" L_COUNT = L_COUNT + 1" & _
						" WHERE L_ID = 127"
					my_Conn.Execute (strSql)
	
					' *** RECORD to BROWSER_LIST
					strSql = 	" SELECT BROWSER_LIST.BROWSER_NAME" & _
						" FROM BROWSER_LIST " & _
						" WHERE BROWSER_NAME = '" & strUserLanguage & "'"
					set rsRecord = my_Conn.Execute (strSql)
					if rsRecord.EOF then
						strSql = "INSERT INTO BROWSER_LIST ("
						strSql = strSql & " BROWSER_NAME" & ", BROWSER_KEY" & ", BROWSER_TYPE" & ", BROWSER_COUNT"
						strSql = strSql & ") VALUES ("
						strSql = strSql & "'" & strUserLanguage & "'"
						strSql = strSql & ", '" & DateToStr(strHomeTimeAdjust) & "', 9, 1"
						strSql = strSql & ")"
						my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
					else
						strSql = 	" UPDATE BROWSER_LIST SET " & _
							" BROWSER_COUNT = BROWSER_COUNT + 1 ," & _
							" BROWSER_KEY = '" & DateToStr(strHomeTimeAdjust) & "'" & _
							" WHERE BROWSER_NAME= '" & strUserLanguage & "'"
						my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
					end if
	
				end if

		'########################################################################################
		
		strTodayValue = left(DateToStr(strHomeTimeAdjust),8)
		strHoursValue = mid(DateToStr(strHomeTimeAdjust),9,2)
		
		strSql =	" SELECT STATISTICS.CALENDAR_LIST FROM STATISTICS " & _
			" WHERE STATISTICS.CALENDAR_LIST= '" & strTodayValue & "'"
			set rsDay = my_Conn.Execute (strSql)
		
			if rsDay.EOF OR rsDay.BOF then
			
				strSql =	"SELECT STATISTICS.CALENDAR_LIST FROM STATISTICS " & _
					"ORDER BY CALENDAR_LIST DESC"
					set rsLast = my_Conn.Execute (strSql)			


				FOR i = 1 to DateDiff("d",cDate(chkDate(DateToStr(rsLast("CALENDAR_LIST") & "000000"),"",f)),cDate(chkDate(DateToStr(strHomeTimeAdjust),"",f))) 
				
					'## STATISTICS DATA INSERT INTO ## ADD NEW ##
					strSql = "INSERT INTO STATISTICS ("
					strSql = strSql & " CALENDAR_LIST"
					strSql = strSql & ", VISITORS_TOTAL" & ", VISITORS_MEMBER" & ", VISITORS_RETURN" & ", H"& strHoursValue 
					strSql = strSql & ") VALUES ("
					strSql = strSql & "'" & Mid(DateToStr(DateAdd("d", i , chkDate(DateToStr(rsLast("CALENDAR_LIST") & "000000"),"",f))),1,8) & "'"
					strSql = strSql & ", " & 0 & ", " & 0 & ", " & 0 & ", " & 0
					strSql = strSql & ")"
					my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
				NEXT

				strSql = 	" UPDATE STATISTICS SET " & _
					" VISITORS_TOTAL = 1 ," & _			
					" H" & strHoursValue & " = 1" & _
					" WHERE CALENDAR_LIST= '" & strTodayValue & "'"
				my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
				
			Else		
				
				strSql = 	" UPDATE STATISTICS SET " & _
					" VISITORS_TOTAL = VISITORS_TOTAL + 1 "				
					
					' ***** THE COOKIEs FOR OTHER COUNTER  IS NOT NEW USER *****************
					if strSinceLoginDate <> strNowDateMMDD then 	' ##### IS RETURN USER VISITORs
						strSql = strSql & ", VISITORS_RETURN = VISITORS_RETURN + 1 "
					End if
					' ****************************************************************
					
				
				strSql = strSql & 	", H" & strHoursValue & " = H" & strHoursValue & " + 1 " & _
						" WHERE CALENDAR_LIST= '" & strTodayValue & "'"
				my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
	
			End if

%>