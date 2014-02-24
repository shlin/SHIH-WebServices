<%
function doublenum(fNum)
	if fNum > 9 then 
		doublenum = fNum 
	else 
		doublenum = "0" & fNum
	end if
end function

function DateToStr(dtDateTime)
	if not isDate(dtDateTime) then
		dtDateTime = strToDate(dtDateTime)
	end if
	DateToStr = year(dtDateTime) & doublenum(Month(dtdateTime)) & doublenum(Day(dtdateTime)) & doublenum(Hour(dtdateTime)) & doublenum(Minute(dtdateTime)) & doublenum(Second(dtdateTime)) & ""
end function

function StrToDate(strDateTime)
	if ChkDateFormat(strDateTime) then
	
		'Testing for server format
		if strComp(Month("04/05/2002"),"4") = 0 then
			StrToDate = cdate("" & Mid(strDateTime, 5,2) & "/" & Mid(strDateTime, 7,2) & "/" & Mid(strDateTime, 1,4) & " " & Mid(strDateTime, 9,2) & ":" & Mid(strDateTime, 11,2) & ":" & Mid(strDateTime, 13,2) & "")
		else
			StrToDate = cdate("" & Mid(strDateTime, 7,2) & "/" & Mid(strDateTime, 5,2) & "/" & Mid(strDateTime, 1,4) & " " & Mid(strDateTime, 9,2) & ":" & Mid(strDateTime, 11,2) & ":" & Mid(strDateTime, 13,2) & "")
		end if
	else
		if strComp(Month("04/05/2002"),"4") = 0 then
			tmpDate = DatePart("m",strHomeTimeAdjust) & "/" & DatePart("d",strHomeTimeAdjust) & "/" & DatePart("yyyy",strHomeTimeAdjust) & " " & DatePart("h",strHomeTimeAdjust) & ":" & DatePart("n",strHomeTimeAdjust) & ":" & DatePart("s",strHomeTimeAdjust)
		else
			tmpDate = DatePart("d",strHomeTimeAdjust) & "/" & DatePart("m",strHomeTimeAdjust) & "/" & DatePart("yyyy",strHomeTimeAdjust) & " " & DatePart("h",strHomeTimeAdjust) & ":" & DatePart("n",strHomeTimeAdjust) & ":" & DatePart("s",strHomeTimeAdjust)
		end if
		StrToDate = tmpDate
	end if
end function

function chkDateFormat(strDateTime)
	chkDateFormat = isdate("" & Mid(strDateTime, 5,2) & "/" & Mid(strDateTime, 7,2) & "/" & Mid(strDateTime, 1,4) & " " & Mid(strDateTime, 9,2) & ":" & Mid(strDateTime, 11,2) & ":" & Mid(strDateTime, 13,2) & "")
end function

function chkDate(fDate,separator,fTime)
	if fDate = "" or isNull(fDate) then
		if fTime then
			chkTime(fDate)
		end if
		exit function
	end if
	select case strDateType
		case "dmy"
			chkDate = Mid(fDate,7,2) & "/" & _
			Mid(fDate,5,2) & "/" & _
			Mid(fDate,1,4)
		case "mdy"
			chkDate = Mid(fDate,5,2) & "/" & _
			Mid(fDate,7,2) & "/" & _
			Mid(fDate,1,4)
		case "ymd"
			chkDate = Mid(fDate,1,4) & "/" & _
			Mid(fDate,5,2) & "/" & _
			Mid(fDate,7,2)
		case "ydm"
			chkDate =Mid(fDate,1,4) & "/" & _
			Mid(fDate,7,2) & "/" & _
			Mid(fDate,5,2)
		case "dmmy"
			chkDate = Mid(fDate,7,2) & " " & _
			Monthname(Mid(fDate,5,2),1) & " " & _
			Mid(fDate,1,4)
		case "mmdy"
			chkDate = Monthname(Mid(fDate,5,2),1) & " " & _
			Mid(fDate,7,2) & " " & _
			Mid(fDate,1,4)
		case "ymmd"
			chkDate = Mid(fDate,1,4) & " " & _
			Monthname(Mid(fDate,5,2),1) & " " & _
			Mid(fDate,7,2)
		case "ydmm"
			chkDate = Mid(fDate,1,4) & " " & _
			Mid(fDate,7,2) & " " & _
			Monthname(Mid(fDate,5,2),1)
		case "dmmmy"
			chkDate = Mid(fDate,7,2) & " " & _
			Monthname(Mid(fDate,5,2),0) & " " & _
			Mid(fDate,1,4)
		case "mmmdy"
			chkDate = Monthname(Mid(fDate,5,2),0) & " " & _
			Mid(fDate,7,2) & " " & _
			Mid(fDate,1,4)
		case "ymmmd"
			chkDate = Mid(fDate,1,4) & " " & _
			Monthname(Mid(fDate,5,2),0) & " " & _
			Mid(fDate,7,2)
		case "ydmmm"
			chkDate = Mid(fDate,1,4) & " " & _
			Mid(fDate,7,2) & " " & _
			Monthname(Mid(fDate,5,2),0)
		case else
			chkDate = Mid(fDate,5,2) & "/" & _
			Mid(fDate,7,2) & "/" & _
			Mid(fDate,1,4)
	end select
	if fTime then
		chkDate = chkDate & separator & chkTime(fDate)
	end if
end function

function chkTime(fTime)
	if fTime = "" or isNull(fTime) then
		exit function
	end if
	if strTimeType = 12 then
		if cLng(Mid(fTime, 9,2)) > 12 then
			chkTime = ChkTime & " " & _
			(cLng(Mid(fTime, 9,2)) -12) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & "PM"
		elseif cLng(Mid(fTime, 9,2)) = 12 then
			chkTime = ChkTime & " " & _
			cLng(Mid(fTime, 9,2)) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & "PM"
		elseif cLng(Mid(fTime, 9,2)) = 0 then
			chkTime = ChkTime & " " & _
			(cLng(Mid(fTime, 9,2)) +12) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & "AM"
		else
			chkTime = ChkTime & " " & _
			Mid(fTime, 9,2) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & "AM"
		end if
		
	else
		ChkTime = ChkTime & " " & _
		Mid(fTime, 9,2) & ":" & _
		Mid(fTime, 11,2) & ":" & _
		Mid(fTime, 13,2) 
	end if
end function
%>