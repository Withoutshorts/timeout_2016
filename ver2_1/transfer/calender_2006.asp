
<!--#include file="../xml/calender_xml_inc.asp"-->

<%
'*** Se timer for andre medarb ****
if request("FM_andre") = "show" then%>
<!--#include file="../connection/conn_db_inc.asp"-->
<LINK rel="stylesheet" type="text/css" href="../style/timeout_style.css">
<%
varMed_id = request("FM_medarb_id")
leftDiv = 30
topDiv = 77
Response.write "<br>Timeregistreringer for:<br><b>" & request("FM_Mnavn") & "</b><br>"
else
'*** fjernet 6/5-2004 ***%>
<!--include file="../connection/conn_db_inc.asp"-->
<!--include file="header_inc.asp"-->
<%
varMed_id = usemrn
leftDiv = 0
topDiv = 10
end if


'******************************************* dato funktioner ***********************************
		
		'*** Sætter lokal dato/kr format. *****
		'Session.LCID = 1030
		
		If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		
		'cval = "her"
		
			if len(request("FM_start_dag")) <> 0 then 'Fra datovælger form
			strMrd = request("FM_start_mrd")
				select case strMrd
				case 2
					if request("FM_start_dag") > 28 then
					strDag = 28
					else
					strDag = request("FM_start_dag")
					end if
				case 1, 3, 5, 7, 8, 10, 12
					strDag = request("FM_start_dag")
				case else
					if request("FM_start_dag") > 30 then
					strDag = 30
					else
					strDag = request("FM_start_dag")
					end if
				end select
			strAar = request("FM_start_aar")
			else
				
				
				
				if len(request.cookies("timereg_2006")("dato")) <> 0 AND session("forste") = "n" then
					
					strDag = day(request.cookies("timereg_2006")("dato"))
					strMrd = month(request.cookies("timereg_2006")("dato"))
					strAar = year(request.cookies("timereg_2006")("dato"))
					
					'cval = "cook"
					
				else
					
					strDag = day(now)
					strMrd = month(now)
					strAar = year(now)
					
					'cval = "dagsdato"
				
				end if
			
			end if
		else
			if len(request("strdag")) <> 0 then 'Fra dato link
			strMrd = request("strmrd")
				select case strMrd
				case 2
					if request("strdag") > 28 then
					strDag = 28
					else
					strDag = request("strdag")
					end if
				case 1, 3, 5, 7, 8, 10, 12
					strDag = request("strdag")
				case else
					if request("strdag") > 30 then
					strDag = 30
					else
					strDag = request("strdag")
					end if
				end select
			strAar = request("straar")
			else
				
				if len(request.cookies("timereg_2006")("dato")) <> 0 AND session("forste") = "n" then
					
					strDag = day(request.cookies("timereg_2006")("dato"))
					strMrd = month(request.cookies("timereg_2006")("dato"))
					strAar = year(request.cookies("timereg_2006")("dato"))
					
					'cval = "cook"
					
				else
					
					strDag = day(now)
					strMrd = month(now)
					strAar = year(now)
					
					'cval = "dagsdato"
				
				end if
			
			end if
		end if
		
		daynow = formatdatetime(day(now) & "/" & month(now) & "/" & year(now), 0)
		useDate = formatdatetime(strDag & "/" & strMrd & "/" & strAar, 0)
		firstDayOfMonth = formatdatetime(1 & "/" & strMrd & "/" & strAar, 0)
		
		useDatePrevWeek = dateadd("d", -7, useDate)
		useDateNextWeek = dateadd("d", 7, useDate)
		
		Select case strMrd
		case 1, 3, 5, 7, 8, 10, 12
		lastDayOfMonth = formatdatetime("31/" & strMrd & "/" & strAar, 0)
		numberofdaysinmonth = 31
		case 2
			select case strAar
			case 2004, 2008, 2012, 2016, 2020
			lastDayOfMonth = formatdatetime("29/" & strMrd & "/" & strAar, 0)
			numberofdaysinmonth = 29
			case else
			lastDayOfMonth = formatdatetime("28/" & strMrd & "/" & strAar, 0)
			numberofdaysinmonth = 28
			end select
		case else
		lastDayOfMonth = formatdatetime("30/" & strMrd & "/" & strAar, 0)
		numberofdaysinmonth = 30
		end select
		
		firstWeekday = Weekday(firstDayOfMonth, 2) 
		lastWeekday = Weekday(lastDayOfMonth, 2) 
		
		prevMonth = datePart("m", DateAdd("m",-1, useDate))
		nextMonth = datePart("m", DateAdd("m",1, useDate))
		
		thisMonthName = monthname(strMrd)
		prevMonthName = left(monthname(prevMonth), 3)
		nextMonthName = left(monthname(nextMonth), 3)
		
		select case prevMonth 
		case 12
		prevYear = datePart("yyyy", DateAdd("yyyy",-1, useDate))
		case else
		prevYear = strAar
		end select
		
		select case nextMonth 
		case 1
		nextYear = datePart("yyyy", DateAdd("yyyy",1, useDate))
		case else
		nextYear = strAar
		end select
		
		countDay = 1
		andre = request("FM_andre")
		
		
			seldocument = "timereg_akt_2006"
			showother = "dontshow"
			kalenderlink = "searchstring="&searchstring&""
			illpath = "../"
			
		
	response.cookies("timereg_2006")("dato") = strDag&"/"&strMrd&"/"&strAar
	
'*****************************************************************************************

kalenderMnavn = ""


if request("FM_andre") <> "show" AND (thisfile = "timereg_akt_2006") then
strSQL = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = "& usemrn 'medarbedjer id
oRec.open strSQL, oConn, 3 
if not oRec.EOF then 
kalenderMnavn = "<font class=megetlillehvid>"&oRec("mnavn")&"&nbsp;("&oRec("mnr")&")</font>"
end if
oRec.close 
end if


if print <> "j" then
%>

<table cellpadding="0" cellspacing="0" border="0" width="100%">
<form action="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=1&fromsdsk=<%=fromsdsk%>" method="POST" name="pickdate" id="pickdate">
<tr bgcolor="#5C75AA">
	<td width="194" colspan="4" style="border-top : 1px; border-bottom : 0px; border-left : 1px; border-right : 1px; border-color : #003399; border-style : solid; padding-right:0; padding-bottom:3; padding-left:5; padding-top:2;">
	<font class="stor-hvid"><%=calender_txt_107%></font> <%=kalenderMnavn%></td>
</tr>
<tr bgcolor="#5C75AA">
	<td valign="top"><img src="<%=illpath%>ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td>
	<td align="left" style="padding-bottom:3; padding-left:5;">
	<!--#include file="../../timereg/inc/weekselector_cal.asp"--></td>
	<td valign="top" align="right"><img src="<%=illpath%>ill/tabel_top.gif" width="1" height="25" alt="" border="0"></td></tr>
	<tr><td colspan=4 bgcolor="#003399" height=1><img src="<%=illpath%>ill/blank.gif" width="1" height="1" alt="" border="0"></td></tr>
</form></table>

<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#FFFFFF">
	<tr bgcolor="ffffff" height="25">
		<td>
		<img src="<%=illpath%>ill/pile_tilbage.gif" alt="" vspace="1" hspace="2" border="0">
		<a class="vmenu" href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=<%=showakt%>&strdag=28&strmrd=<%=prevMonth%>&straar=<%=prevYear%>&<%=kalenderlink%>&fromsdsk=<%=fromsdsk%>"><%=prevMonthName%></a></td>
		<td align="center"><b><%=thisMonthName%>&nbsp;<%=strAar%></b></td>
		<td align="right"><a class="vmenu" href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=<%=showakt%>&strdag=1&strmrd=<%=nextMonth%>&straar=<%=nextYear%>&<%=kalenderlink%>&fromsdsk=<%=fromsdsk%>"><%=nextMonthName%></a><img src="<%=illpath%>ill/pile_selected.gif" alt="" vspace="1" hspace="2" border="0"></td>
		</tr>
</table>
<table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#FFFFFF">
<tr bgcolor="#ffffff">
	<td align=center width=15></td>
	<td align=center width=30><%=calender_txt_108 %></td>
    <td align=center width=30><%=calender_txt_109 %></td>
    <td align=center width=30><%=calender_txt_110 %></td>
    <td align=center width=30><%=calender_txt_111 %></td>
    <td align=center width=30><%=calender_txt_112 %></td>
    <td align=center width=30><%=calender_txt_113 %></td>
	<td align=center width=30><%=calender_txt_114 %></td>
	<td align=center width=30>Tot.</td>
</tr>
<tr>
	<td colspan="9" bgcolor="#000000"><img src="<%=illpath%>ill/blank.gif" alt="" border="0"></td>
</tr>

<%
'**************************************************************************************************'
'** Udskriver dage og timeforbrug i kalender.
'**************************************************************************************************'
WeekdayfirstW = firstWeekday 

fwdno = 1
acls = "class=vmenu"
acls2 = "class=vmenu"
'** ugenr **%>
<tr><td height=25 class=lillegray valign=top align=right style="padding:0px 3px 0px 3px;">
<%=datepart("ww", firstDayOfMonth,2,2)%></td>

<%
'** Mellemrum før første dag i første uge
for fwdno = 1 to firstWeekday - 1 %>
<td>&nbsp;</td>
<%
next

daysinFirstWeek = 1
cweekTot = 0
mTot = 0
'** Finder de typer der er med i det daglige timeregnskab ***'
call akttyper2009(2)

'*** Første uge datoer og timer/crm aktioner
for WeekdayfirstW = firstWeekday To 7 

if menu = "timereg" OR menu = "" then
	'** Henter timer ***
		strSQL = "SELECT sum(timer) AS timer_indtastet FROM timer WHERE Tmnr = "& varMed_id &" AND Tdato='"& strAar &"/" & strMrd & "/" & daysinFirstWeek &"' AND ("& aty_sql_realhours &")"
		'(tfaktim = 1 OR tfaktim = 2 OR tfaktim = 6 OR tfaktim = 13 OR tfaktim = 14 OR tfaktim = 20 OR tfaktim = 21)"
		
		'Response.Write strSQL
		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		intDayHoursVal = oRec("timer_indtastet")
	    end if
		strAktnavn = ""
		oRec.close
else
	'** Henter crmaktioner ***
			strSQL3 = "SELECT crmhistorik.id, aktionsid, medarbid, kkundenavn, crmhistorik.navn, crmemne.navn AS emnenavn FROM crmhistorik, aktionsrelationer LEFT JOIN crmemne ON(crmemne.id = kontaktemne) LEFT JOIN kunder ON (kunder.kid = kundeid) WHERE ((crmdato = '"& strAar &"/" & strMrd & "/" & daysinFirstWeek &"') OR (crmdato <= '"& strAar &"/" & strMrd & "/" & daysinFirstWeek &"' AND crmdato_slut >= '"& strAar &"/" & strMrd & "/" & daysinFirstWeek &"' AND crmdato <> crmdato_slut)) "& usemedarbKri &" ORDER BY crmhistorik.id"
			oRec3.open strSQL3, oConn, 3
			akt_today = "n"
			strAktnavn = "Aktioner idag:" &vbcrlf
				while not oRec3.EOF
				strAktnavn = strAktnavn & oRec3("Kkundenavn") & chr(032) & oRec3("emnenavn") & chr(032) & oRec3("navn") &vbcrlf
				akt_today = "y"
				oRec3.movenext
				wend
			oRec3.close
			
end if		
%>
<td align=center valign=top>
<%
	if cdate(day(now)&"/" & month(now)&"/" & year(now)) = cdate(daysinFirstWeek&"/" & strMrd &"/" & strAar) then
	acls = "class=red"
	else
	acls = "class=kalblue"
	end if
%>
<a <%=acls%> href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=1&strdag=<%=daysinFirstWeek%>&strmrd=<%=strMrd%>&straar=<%=strAar%>&<%=kalenderlink%>&fromsdsk=<%=fromsdsk%>"><font title="<%=strAktnavn%>"><%=daysinFirstWeek%></font></a>
<%

	Response.write "<br>"
	if intDayHoursVal <> 0 then%>
	<span class=cal><%=intDayHoursVal%></span>
	
	<%
	cweekTot = cweekTot + intDayHoursVal
	end if%>
</td>
<%		
lastDayinfirstWeek = daysinFirstWeek
daysinFirstWeek = daysinFirstWeek + 1


next
%>
<td valign=bottom align=right><span class=caltot><%=cweekTot %></span>
<%mTot = mTot +  cweekTot%>
		<%cweekTot = 0 %></td>
</tr>



<%
'*** De næste uger datoer og timer/ crm aktioner
startsecondWeek = lastDayinfirstWeek + 1
cweekTot = 0
for dayCount = startsecondWeek to numberofdaysinmonth
	
	
	
	
	if Weekday(formatdatetime(dayCount &"/" & strMrd & "/" & strAar, 0), 2) = 1 then 
		
		if dayCount <> startsecondWeek then%>
		<td valign=bottom align=right><span class=caltot><%=cweekTot %></span>
		<%mTot = mTot +  cweekTot%>
		<%cweekTot = 0 %></td>
		</tr>
		<%end if%>
	<tr>
		<td colspan="9" bgcolor="#D6DFF5"><img src="<%=illpath%>ill/blank.gif" alt="" border="0"></td>
	</tr>
	<tr>
		<td height=25 class=lillegray valign=top align=right style="padding:0px 3px 0px 3px;">
		<%=datepart("ww", dayCount &"/" & strMrd & "/" & strAar,2,2)%>
</td>
	<%end if%>
	
	<%
	
	strSQL = "SELECT sum(timer) AS timer_indtastet FROM timer WHERE Tmnr = "& varMed_id &" AND Tdato='"& strAar &"/" & strMrd & "/" & dayCount  &"' AND ("& aty_sql_realhours &")" 
	'AND (tfaktim = 1 OR tfaktim = 2 OR tfaktim = 6 OR tfaktim = 13 OR tfaktim = 14 OR tfaktim = 20 OR tfaktim = 21)"
		
		'Response.Write strSQL
		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		intDayHoursVal = oRec("timer_indtastet")
		end if
		strAktnavn = ""
		oRec.close
		
		
	
	 %>
	
<td align=center valign=top>
<%
if cdate(day(now)&"/" & month(now)&"/" & year(now)) = cdate(dayCount &"/" & strMrd &"/" & strAar) then
acls2 = "class=red"
else
acls2 = "class=kalblue"
end if%>
<a <%=acls2%> href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=1&strdag=<%=dayCount%>&strmrd=<%=strMrd%>&straar=<%=strAar%>&<%=kalenderlink%>&fromsdsk=<%=fromsdsk%>"><font title="<%=strAktnavn%>"><%=dayCount%></font></a>
<br>
<%

	if intDayHoursVal <> 0 then%>
	<span class=cal><%=intDayHoursVal %></span>
	<%
	cweekTot = cweekTot + intDayHoursVal
	end if%>
</td>
<%

lastDay = dayCount
   
next 
'**************************************************************************************************'
ld = Weekday(formatdatetime(lastDay &"/" & strMrd & "/" & strAar, 0), 2)

for d = ld + 1 to 7 %>
<td>&nbsp</td>
<%next %>

<td valign=bottom align=right>
<span class=caltot><%=cweekTot %>
<%mTot = mTot +  cweekTot%></span> <br />
<span class=caltot><b><%=mTot %></b></span></td>
</tr>
</table>
<%cweekTot = 0 
mTot = 0%>
<%end if %>


