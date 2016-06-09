<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	
	select case func 
	case "slet"
	oConn.execute("DELETE FROM ressourcer WHERE id = "&id&"")
	activeuge = request("activeuge")
	yearthis = request("yearthis")
	monththis = request("monththis")
	Response.redirect "jbpla_w.asp?activeuge="&activeuge&"&mdselect="&monththis&"&aarselect="&yearthis
	
	
	case "reddb"
	antalU = request("antalu")
	
	for u = 1 to antalU
	
		medid = Request("u_this_"&u&"")
		
		
				'*** Bestemmer dato **
				for xval = 1 to 7
					select case xval
					case 1
					bundval = 0
					topval = 90
					case 2
					bundval = 90
					topval = 134
					case 3
					bundval = 134
					topval = 178
					case 4
					bundval = 178
					topval = 222
					case 5
					bundval = 222
					topval = 266
					case 6
					bundval = 266
					topval = 310
					case 7
					bundval = 310
					topval = 800
					end select
					
					if cint(request("bookdiv_x_"& medid &"_"&u&"")) > cint(bundval) AND cint(request("bookdiv_x_"& medid &"_"&u&"")) <= cint(topval) then
						thisdato = request("datoweekday"&xval&"") 
						thisdatoSQL = year(thisdato)&"/"&month(thisdato)&"/"&day(thisdato)
					end if
				
				
				next
		
				'Response.write thisdato &"<br>"
				
				
				'** Betstemmer klokkeslet **
				sttid = thisdato&" 08:00:00"
				
				for yval = 1 to 36
				sttidthis = dateadd("n", (yval*15), sttid)
				bundval = 155 + (yval*15)
				topval = 170 + (yval*15)	
				
					if cint(request("bookdiv_y_"& medid &"_"&u&"")) > cint(bundval) AND cint(request("bookdiv_y_"& medid &"_"&u&"")) <= cint(topval) then
						thisstarttid = sttidthis 
					end if
					
					
				next
				
				'Response.write "<b>"&thisstarttid &"</b><br>"
		
		
				'*** Bestemmmer sluttid ***
				h = request("bookdiv_height_"& medid &"_"&u&"")
				instr_h = 0
				instr_h = instr(h, "p")
				if instr_h <> 0 then
				len_h = len(h)
				left_h = left(h, len_h - 2)
				else
				left_h = h
				end if
				
				thissluttid = dateadd("n", left_h, thisstarttid)
				
				'Response.write thissluttid & "<br>"
				
				divid = request("bookdiv_id_"& medid &"_"&u&"")
				
				strSQL = "UPDATE ressourcer SET dato = '"& thisdatoSQL &"', starttp = '"& formatdatetime(thisstarttid, 3) &"', sluttp = '"& formatdatetime(thissluttid, 3) &"' WHERE id = " & divid 
				oConn.execute(strSQL)
		%>
	
	<%next
	
	activeuge = request("activeuge")
	yearthis = request("yearthis")
	monththis = request("monththis")
	Response.redirect "jbpla_w.asp?activeuge="&activeuge&"&mdselect="&monththis&"&aarselect="&yearthis
	
	
	case else 
	
	
	
	if len(request("mdselect")) <> 0 then
		monththis = request("mdselect")
		
		select case monththis 
		case 0
		monththis = 12
		yearthis = request("aarselect")
		yearthis = yearthis - 1
		case 13
		monththis = 1
		yearthis = request("aarselect")
		yearthis = yearthis + 1
		case else
		yearthis = request("aarselect")
		end select
		
	else
	monththis = month(now)
	yearthis = year(now)
	end if
	
	
			'** Antal dage i md ***
			Select case monththis
			case "4", "6", "9", "11"
				mthDays = 30
			case 2
					select case yearthis
					case 2004, 2008, 2012, 2016, 2020, 2024, 2028
					mthDays = 29
					case else 
					mthDays = 28
					end select
			case else
				mthDays = 31
			end select
			
			
	
	
	
	
	
	%>
	<SCRIPT language=javascript src="inc/jbpla_w_func.js"></script>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
		
	
	<%leftPos = 20%>
	<%topPos = 20%>
	<!-------------------------------Sideindhold------------------------------------->
	
	<!-- Oversigts kalender -->
	<div id="overblik" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0">
	<form action="jbpla_w.asp" target="w" method="post" name="mdselect" id="mdselect">
	<tr><td>
	<!--Job:
	<select name="jobidselect" style="width:150;">
	<option value="0">Alle (ikke tilsluttet)</option>
	<%
	'strSQLJ = "SELECT jobnavn, id, jobnr FROM job WHERE jobstatus = 1 AND fakturerbart = 1 ORDER BY jobnavn"
	'oRec.open strSQLJ, oConn, 3 
	'while not oRec.EOF 
	%>
	<option value="<=oRec("id")%>"><=left(oRec("jobnavn"), 20)%> (<=oRec("jobnr")%>)</option>
	<%
	'oRec.movenext
	'wend
	'oRec.close 
	%>
	</select>-->
	
	Vælg periode:&nbsp;<select name="mdselect">
	<option value="<%=monththis%>" SELECTED><%=left(monthname(monththis),3)%></option>
	<option value="1">Jan</option>
	<option value="2">Feb</option>
	<option value="3">Mar</option>
	<option value="4">Apr</option>
	<option value="5">Maj</option>
	<option value="6">Jun</option>
	<option value="7">Jul</option>
	<option value="8">Aug</option>
	<option value="9">Sep</option>
	<option value="10">Okt</option>
	<option value="11">Nov</option>
	<option value="12">Dec</option>
	</select><select name="aarselect">
	<option value="<%=yearthis%>" SELECTED><%=yearthis%></option>
	<option value="2002">2002</option>
	<option value="2003">2003</option>
	<option value="2004">2004</option>
	<option value="2005">2005</option>
	<option value="2005">2006</option>
	<option value="2005">2007</option>
	<option value="2005">2008</option>
	</select></td><td valign="bottom"><input type="image" src="../ill/pilstorxp.gif"></td></tr>
	</form>
	</table>
	
	<a href="#" class=vmenu>Opret nyt job</a><br>
	
	<%
	
	mthDaysUse = mthDays
	monthUse = monththis
	yearUse = yearthis
	
	daynow = day(now)&"/"&month(now)&"/"&year(now)
	denneuge = datepart("ww", day(now)&"/"&month(now)&"/"&year(now))
	lastweek = 0
	
		'** Finder aktive uge **
		if len(request("activeuge")) <> 0 then
		activeuge = request("activeuge")
		else
		
		for d = 0 to mthDaysUse - 1
		ugeday = d + 1
		ugeweekdato = ugeday&"/"&monththis&"/"&yearUse
		
			if d = 0 then
			activeuge = datepart("ww", ugeweekdato)
			end if 
			
			if denneuge = datepart("ww", ugeweekdato) then
			activeuge = datepart("ww", ugeweekdato)
			else
			activeuge = activeuge
			end if
			
		next
		end if
	
			
		firstdayofyear = "1/1/2004"
		weekdayfirstday = dateadd("ww", activeuge-1, firstdayofyear)
		'Response.write weekdayfirstday
		
		select case weekday(weekdayfirstday)
		case 1
		perStartKri = weekdayfirstday
		perSlutKri = dateadd("d", 6, weekdayfirstday)
		case 2	
		perStartKri = dateadd("d", -1, weekdayfirstday)
		perSlutKri = dateadd("d", 5, weekdayfirstday)
		case 3
		perStartKri = dateadd("d", -2, weekdayfirstday)
		perSlutKri = dateadd("d", 4, weekdayfirstday)
		case 4
		perStartKri = dateadd("d", -3, weekdayfirstday)
		perSlutKri = dateadd("d", 3, weekdayfirstday)
		case 5
		perStartKri = dateadd("d", -4, weekdayfirstday)
		perSlutKri = dateadd("d", 2, weekdayfirstday)
		case 6
		perStartKri = dateadd("d", -5, weekdayfirstday)
		perSlutKri = dateadd("d", 1, weekdayfirstday)
		case 7
		perStartKri = dateadd("d", -6, weekdayfirstday)
		perSlutKri = weekdayfirstday
		end select
	
	
	jobStartKri = year(perStartKri)&"/"&month(perStartKri)&"/"&day(perStartKri)
	jobSlutKri = year(perSlutKri)&"/"&month(perSlutKri)&"/"&day(perSlutKri)						
		
	
	'Response.write 	"jobStartKri " & jobStartKri & "jobSlutKri " & jobSlutKri
	
	'****************************************************************
	'*** Ugekalender ***
	'****************************************************************
	
	
	%>
	<form action="jbpla_w.asp?func=reddb&activeuge=<%=activeuge%>&mdselect=<%=monththis%>&aarselect=<%=yearthis%>" target="w" method="post" name="opdater" id="opdater">
	
	<%
	'** Medarbejdere ***%>
	<div id="medarbejdere" name="medarbejdere" style="position:absolute; left:10; top:40; visibility:visible; display:; width:150; padding:2px;">
	<table cellspacing=0 cellpadding=0 border=0><tr><td width=120 valign=top>
	<img src="../ill/soeg-knap.gif" width="16" height="16" alt="Oversigt" border="0">&nbsp;<a href="#" onClick="showmidbook(0)" class=vmenu>Oversigt</a><br>
	<%
	dim medarbid, medarbnavn
	redim medarbid(0), medarbnavn(0)
	m = 1
			strSQLM = "SELECT mnavn, mid, mnr FROM medarbejdere WHERE mansat <> 2 ORDER BY mnavn"
			oRec.open strSQLM, oConn, 3 
			while not oRec.EOF 
			
				if right(m, 1) = 5 OR right(m, 1) = 0 then%>
				</td><td width=120 valign=top>
				<%end if
				
							select case oRec("mid")
							case 1, 47, 93
							medcol = "#66CC33"
							case 2, 48, 94
							medcol = "#333366"
							case 3, 49, 95
							medcol = "#0066CC"
							case 4, 50, 96
							medcol = "#0083D7"
							case 5, 51, 97
							medcol = "#0099FF"
							case 6, 52, 98
							medcol = "#006699"
							case 7, 53, 99
							medcol = "#34A4DE"
							case 8, 54, 100
							medcol = "#B4E2FF"
							case 9, 55, 101
							medcol = "#003399"
							case 10, 56, 102
							medcol = "#FFCCFF"
							case 11, 57, 103
							medcol = "#CCCCFF"
							case 12, 58, 104
							medcol = "#996600"
							case 13, 59, 105
							medcol = "#CC9900"
							case 14, 60, 106
							medcol = "#FFFF99"
							case 15, 61, 107
							medcol = "#FFCC00"
							case 16, 62, 108
							medcol = "#FFFF00"
							case 17, 63, 109
							medcol = "#FFDB9D"
							case 18, 64, 110
							medcol = "#FFCC66"
							case 19, 65, 111
							medcol = "#FF9933"
							case 20, 66, 112
							medcol = "#FF794B"
							case 21, 67, 113
							medcol = "#FF3300"
							case 22, 68, 114
							medcol = "#990000"
							case 23, 69, 115
							medcol = "#9999FF"
							case 24, 70, 116
							medcol = "#6666CC"
							case 25, 71, 117
							medcol = "#9999CC"
							case 26, 72, 118
							medcol = "#666699"
							case 27, 73, 119
							medcol = "#006600"
							case 28, 74, 120
							medcol = "#009900"
							case 29, 75, 121
							medcol = "#DEFFFF"
							case 30, 76, 122
							medcol = "Silver"
							case 31, 77, 123
							medcol = "LightGrey"
							case 32, 78, 125
							medcol = "DarkGray"
							case 33, 79
							medcol = "Olive"
							case 34, 80
							medcol = "Dimgray"
							case 35, 81
							medcol = "#99FF66"
							case 36, 82
							medcol = "#CCFFCC"
							case 37, 83
							medcol = "RosyBrown"
							case 38, 84
							medcol = "Honeydew"
							case 39, 85
							medcol = "CornflowerBlue"
							case 40, 86
							medcol = "DarkMagenta"
							case 41, 78
							medcol = "skyblue"
							case 42, 88
							medcol = "Thistle"
							case 43, 89
							medcol = "Tomato"
							case 44, 90
							medcol = "Gold"
							case 45, 91
							medcol = "yellow"
							case 46, 92
							medcol = "black"
							case else
							medcol = "#5582d2"
							end select
				
				
				
					thisfornavn = left(oRec("mnavn"), 1)
					navn_instr = instr(oRec("mnavn"), " ")
					navn_len = len(oRec("mnavn"))
					navn_right = right(oRec("mnavn"), (navn_len - (navn_instr)))
					if cint(navn_instr) > 0 then
					thisefternavn = left(navn_right, 1)
					else
					thisefternavn = ""
					end if
					
					initthis = thisfornavn & thisefternavn
					%>
				
				
				<table height=10 cellspacing=1 cellpadding=0 border=0><tr>
				<td width=10 height=8 bgcolor="<%=medcol%>" valign=top style="border:1px #000000 solid; padding:1px;">
				<font class=megetmegetlillehvid><%=initthis%></td>
				<td valign=top style="padding-left:2;"><a href="#" onClick="showmidbook(<%=oRec("mid")%>)" class=vmenu><%=oRec("mnavn")%> (<%=oRec("mnr")%>)</a></td>
				</tr></table> 
				<%
				redim preserve medarbid(m) 
				redim preserve medarbnavn(m)
				medarbid(m) = oRec("mid")
				medarbnavn(m) = oRec("mnavn")
				m = m + 1
							
							
			oRec.movenext
			wend
			oRec.close 
		
	%></td>
	</tr></table>
	</div>
	<input type="hidden" name="antalmedarb" id="antalmedarb" value="<%=m-1%>">
	
	
	<%'** Uger ***%>
	<div id="ugeoversigt" name="ugeoversigt" style="position:absolute; left:0; top:132; visibility:visible; display:; width:150; padding:2px;">
		
		<%
		antaluger = 1
		for d = 0 to mthDaysUse - 1
				
				if d = 0 then%>
				<b>Vælg uge:</b>&nbsp;
				<%end if
		
		ugeday = d + 1
		ugeweekdato = ugeday&"/"&monththis&"/"&yearthis
		
		if lastweek <> datepart("ww", ugeweekdato) then
		%>
		<a href="jbpla_w.asp?activeuge=<%=datepart("ww", ugeweekdato)%>&mdselect=<%=monththis%>&aarselect=<%=yearthis%>" target="w"><%=datepart("ww", ugeweekdato)%></a>
		<%
		end if	
		lastweek = datepart("ww", ugeweekdato)
		next	
		%>
	</div>
		
							
		<%
		'*** Ugens 7 dage ***
		%>
		<div id="klokke_<%=datepart("ww", ugeweekdato)%>" name="klokke_<%=datepart("ww", ugeweekdato)%>" style="position:absolute; z-index:<%=datepart("ww", ugeweekdato)%>; left:8; top:155; visibility:<%=vz%>; display:<%=disp%>;">
		<table cellspacing=1 cellpadding=0 border=0 bgcolor="#ffffff" width="350">
		<tr>
		<td width=50><font class=megetlillesort>Uge: <%=datepart("ww", ugeweekdato)%></td>
					
		<%'*** Datoer på de viste ugedage til brug for DB ***
		for d = 0 to 6
				
		ugeday = d 
		ugeweekdato = dateadd("d", d, perStartKri)
		
				select case weekday(ugeweekdato)
				case 1
				dayn = "S"
				bgdato = "#cccccc"
				ugeweekdato1 = ugeweekdato
				%>
				<input type="hidden" name="datoweekday1" id="datoweekday1" value="<%=ugeweekdato%>">
				<%
				case 2
				dayn = "M"
				bgdato = "#8caae6"
				ugeweekdato2 = ugeweekdato%>
				<input type="hidden" name="datoweekday2" id="datoweekday2" value="<%=ugeweekdato%>">
				<%
				case 3
				dayn = "T"
				bgdato = "#8caae6"
				ugeweekdato3 = ugeweekdato%>
				<input type="hidden" name="datoweekday3" id="datoweekday3" value="<%=ugeweekdato%>">
				<%
				case 4
				dayn = "O"
				bgdato = "#8caae6"
				ugeweekdato4 = ugeweekdato%>
				<input type="hidden" name="datoweekday4" id="datoweekday4" value="<%=ugeweekdato%>">
				<%
				case 5
				dayn = "T"
				bgdato = "#8caae6"
				ugeweekdato5 = ugeweekdato%>
				<input type="hidden" name="datoweekday5" id="datoweekday5" value="<%=ugeweekdato%>">
				<%
				case 6
				dayn = "F"
				bgdato = "#8caae6"
				ugeweekdato6 = ugeweekdato%>
				<input type="hidden" name="datoweekday6" id="datoweekday6" value="<%=ugeweekdato%>">
				<%
				case 7
				dayn = "L"
				bgdato = "#cccccc"
				ugeweekdato7 = ugeweekdato%>
				<input type="hidden" name="datoweekday7" id="datoweekday7" value="<%=ugeweekdato%>">
				<%							
				end select
				
				if cdate(daynow) = cdate(ugeweekdato) then
					bgdato = "#FFCC00"
				else
					bgdato = bgdato
				end if
						
				%>
				<td width=50 bgcolor="<%=bgdato%>" align="center"><font class=megetlillesort><%=dayn%>&nbsp;<%=left(ugeweekdato, 5)%></td>
				<%
				
		next
		%>
		</tr>
		</table>
		</div>	
		
	
	<%'** Grid **%>
	<div id="grid" name="grid" style="position:absolute; left:8; top:173; visibility:visible; display:; border:1px #003399 solid; background-color:#cccccc; padding:2px; z-index:-1000;">
	<map name="grid">
	<%
	r = 0
	x1 = 0
	x2 = 40
	
	for r = 1 to 7
		select case r
		case 1
		imgmapdato = ugeweekdato1
		case 2
		imgmapdato = ugeweekdato2
		case 3
		imgmapdato = ugeweekdato3
		case 4
		imgmapdato = ugeweekdato4
		case 5
		imgmapdato = ugeweekdato5
		case 6
		imgmapdato = ugeweekdato6
		case 7
		imgmapdato = ugeweekdato7
		end select
	x1 = x1 + 45
	x2 = x1 + 45
	
		c = 1
		y1 = 0
		y2 = 14
		imgmapsttid = "08:15:00"
		for c = 0 to 36
		imgmapsttid = dateadd("n", 15, imgmapsttid)
		y1 = y1 + 15
		y2 = y2 + 15%>
		<area alt="" coords="<%=x1%>,<%=y1%>,<%=x2%>,<%=y2%>" href="javascript:NewWin_opr('jbpla_w_opr.asp?func=opr&sttid=<%=imgmapsttid%>&dato=<%=imgmapdato%>&datostkri=<%=jobStartKri%>&datoslkri=<%=jobSlutKri%>&id=0');" target="_self">
		<%next%>
	<%next%>
	</map>
	<img src="../ill/jbpla_grid.gif" width="350" height="542" alt="" border="0" usemap="#grid">
	</div>	
		
		
	<%	
	'*** Udksriver div felter i kalender *** 
	dim ugemid, ugedato, ugesttid, ugesluttid, ugejobid, ugemnavn, ugemnr, divid, bgcol, ugejobnavn, ugejobnr, ugeaid, ugeanavn
	Redim ugemid(0), ugedato(0), ugesttid(0), ugesluttid(0), ugejobid(0), ugemnavn(0), ugemnr(0), divid(0), bgcol(0),ugejobnavn(0), ugejobnr(0), ugeaid(0), ugeanavn(0)
	
	
	strSQL = "SELECT ressourcer.id AS divid, timer, starttp, sluttp, ressourcer.dato, ressourcer.mid AS mid, jobid, aktid, mnavn, mnr, jobnavn, jobnr, aktiviteter.id AS aid, aktiviteter.navn AS anavn FROM ressourcer "_
	&" LEFT JOIN job ON (job.id = jobid) "_
	&" LEFT JOIN medarbejdere ON (medarbejdere.mid = ressourcer.mid) "_
	&" LEFT JOIN aktiviteter ON (aktiviteter.id = aktid) WHERE ressourcer.dato "_
	&" BETWEEN '"& jobStartKri &"' AND '"& jobSlutKri &"' ORDER BY dato"
	
	'Response.write strSQL
	'Response.flush
	
	
	oRec.open strSQL, oConn, 3 
	u = 1
	lastmid = 0
	while not oRec.EOF 
	
	Redim preserve ugemid(u)
	Redim preserve ugedato(u)
	Redim preserve ugesttid(u)
	Redim preserve ugesluttid(u)
	Redim preserve ugejobid(u)
	Redim preserve ugemnavn(u)
	Redim preserve ugemnr(u)
	Redim preserve divid(u)
	Redim preserve bgcol(u)
	Redim preserve ugejobnavn(u)
	Redim preserve ugejobnr(u)
	Redim preserve ugeaid(u)
	Redim preserve ugeanavn(u)
	ugemid(u) = oRec("mid")
	ugedato(u) = oRec("dato")
	ugesttid(u) = oRec("starttp")
	ugesluttid(u) = oRec("sluttp")
	ugejobid(u) = oRec("jobid")
	ugejobnavn(u) = oRec("jobnavn")
	ugejobnr(u) = oRec("jobnr")
	ugemnavn(u) = oRec("mnavn")
	ugemnr(u) = oRec("mnr")
	divid(u) = oRec("divid")
	ugeaid(u) = oRec("aid")
	ugeanavn(u) = oRec("anavn")
	
						
						
						
						select case ugemid(u)
							case 1, 47, 93
							bgcol(u) = "#66CC33"
							case 2, 48, 94
							bgcol(u) = "#333366"
							case 3, 49, 95
							bgcol(u) = "#0066CC"
							case 4, 50, 96
							bgcol(u) = "#0083D7"
							case 5, 51, 97
							bgcol(u) = "#0099FF"
							case 6, 52, 98
							bgcol(u) = "#006699"
							case 7, 53, 99
							bgcol(u) = "#34A4DE"
							case 8, 54, 100
							bgcol(u) = "#B4E2FF"
							case 9, 55, 101
							bgcol(u) = "#003399"
							case 10, 56, 102
							bgcol(u) = "#FFCCFF"
							case 11, 57, 103
							bgcol(u) = "#CCCCFF"
							case 12, 58, 104
							bgcol(u) = "#996600"
							case 13, 59, 105
							bgcol(u) = "#CC9900"
							case 14, 60, 106
							bgcol(u) = "#FFFF99"
							case 15, 61, 107
							bgcol(u) = "#FFCC00"
							case 16, 62, 108
							bgcol(u) = "#FFFF00"
							case 17, 63, 109
							bgcol(u) = "#FFDB9D"
							case 18, 64, 110
							bgcol(u) = "#FFCC66"
							case 19, 65, 111
							bgcol(u) = "#FF9933"
							case 20, 66, 112
							bgcol(u) = "#FF794B"
							case 21, 67, 113
							bgcol(u) = "#FF3300"
							case 22, 68, 114
							bgcol(u) = "#990000"
							case 23, 69, 115
							bgcol(u) = "#9999FF"
							case 24, 70, 116
							bgcol(u) = "#6666CC"
							case 25, 71, 117
							bgcol(u) = "#9999CC"
							case 26, 72, 118
							bgcol(u) = "#666699"
							case 27, 73, 119
							bgcol(u) = "#006600"
							case 28, 74, 120
							bgcol(u) = "#009900"
							case 29, 75, 121
							bgcol(u) = "#DEFFFF"
							case 30, 76, 122
							bgcol(u) = "Silver"
							case 31, 77, 123
							bgcol(u) = "LightGrey"
							case 32, 78, 125
							bgcol(u) = "DarkGray"
							case 33, 79
							bgcol(u) = "Olive"
							case 34, 80
							bgcol(u) = "Dimgray"
							case 35, 81
							bgcol(u) = "#99FF66"
							case 36, 82
							bgcol(u) = "#CCFFCC"
							case 37, 83
							bgcol(u) = "RosyBrown"
							case 38, 84
							bgcol(u) = "Honeydew"
							case 39, 85
							bgcol(u) = "CornflowerBlue"
							case 40, 86
							bgcol(u) = "DarkMagenta"
							case 41, 78
							bgcol(u) = "skyblue"
							case 42, 88
							bgcol(u) = "Thistle"
							case 43, 89
							bgcol(u) = "Tomato"
							case 44, 90
							bgcol(u) = "Gold"
							case 45, 91
							bgcol(u) = "yellow"
							case 46, 92
							bgcol(u) = "black"
							case else
							bgcol(u) = "#5582d2"
							end select
	
	%>
	<input type="hidden" name="u_this_<%=u%>" id="u_this_<%=u%>" value="<%=oRec("mid")%>">
	<%
	
	u = u + 1
	oRec.movenext
	wend
	oRec.close 
	
	antal1 = 0
	antal2 = 0
	antal3 = 0
	antal4 = 0
	antal5 = 0
	antal6 = 0
	antal7 = 0
	
	one_iswrt = "#0#,"
	two_iswrt = "#0#,"
	three_iswrt = "#0#,"
	four_iswrt = "#0#,"
	five_iswrt = "#0#,"
	six_iswrt = "#0#,"
	seven_iswrt = "#0#,"
	
	one_user_iswrt = "#0#,"
	two_user_iswrt = "#0#,"
	three_user_iswrt = "#0#,"
	four_user_iswrt = "#0#,"
	five_user_iswrt = "#0#,"
	six_user_iswrt = "#0#,"
	seven_user_iswrt = "#0#,"
	
	
	for u = 1 to u - 1 
	antalx = 0
	antalux = 0
			
		dtop = datediff("n", "08:00:00 AM", formatdatetime(ugesttid(u), 3))
		dheight = datediff("n", ugesttid(u), ugesluttid(u))	
			
					select case weekday(ugedato(u))
					case 1
					leftc = 53
					if cint(instr(one_iswrt, "#"&dtop&"#")) > 0 then
					antal1 = antal1 + 8
					antalx = antal1
					else
					antal1 = antal1
					antalx = 0
					end if
					one_iswrt = one_iswrt &"#"&dtop&"#," 
					
					if cint(instr(one_user_iswrt, "#"&ugemid(u)&"_"&dtop&"#")) > 0 then
					antalu1 = antalu1 + 8
					antalux = antalu1
					else
					antalu1 = antalu1
					antalux = 0
					end if
					one_user_iswrt = one_user_iswrt &"#"&ugemid(u)&"_"&dtop&"#," 
					
					
					
					case 2
					leftc = 98
					if cint(instr(two_iswrt, "#"&dtop&"#")) > 0 then
					antal2 = antal2 + 8
					antalx = antal2
					else
					antal2 = antal2
					antalx = 0
					end if
					two_iswrt = two_iswrt &"#"&dtop&"#," 
					
					if cint(instr(two_user_iswrt, "#"&ugemid(u)&"_"&dtop&"#")) > 0 then
					antalu2 = antalu2 + 8
					antalux = antalu2
					else
					antalu2 = antalu2
					antalux = 0
					end if
					two_user_iswrt = two_user_iswrt &"#"&ugemid(u)&"_"&dtop&"#," 
					
					case 3
					leftc = 142
					if cint(instr(three_iswrt, "#"&dtop&"#")) > 0 then
					antal3 = antal3 + 8
					antalx = antal3
					else
					antal3 = antal3
					antalx = 0
					end if
					three_iswrt = three_iswrt &"#"&dtop&"#," 
					
					if cint(instr(three_user_iswrt, "#"&ugemid(u)&"_"&dtop&"#")) > 0 then
					antalu3 = antalu3 + 8
					antalux = antalu3
					else
					antalu3 = antalu3
					antalux = 0
					end if
					three_user_iswrt = three_user_iswrt &"#"&ugemid(u)&"_"&dtop&"#," 
					
					case 4
					leftc = 186
					if cint(instr(four_iswrt, "#"&dtop&"#")) > 0 then
					antal4 = antal4 + 8
					antalx = antal4
					else
					antal4 = antal4
					antalx = 0
					end if
					four_iswrt = four_iswrt &"#"&dtop&"#," 
					
					if cint(instr(four_user_iswrt, "#"&ugemid(u)&"_"&dtop&"#")) > 0 then
					antalu4 = antalu4 + 8
					antalux = antalu4
					else
					antalu4 = antalu4
					antalux = 0
					end if
					four_user_iswrt = four_user_iswrt &"#"&ugemid(u)&"_"&dtop&"#," 
					
					case 5
					leftc = 230
					if cint(instr(five_iswrt, "#"&dtop&"#")) > 0 then
					antal5 = antal5 + 8
					antalx = antal5
					else
					antal5 = antal5
					antalx = 0
					end if
					five_iswrt = five_iswrt &"#"&dtop&"#," 
					
					if cint(instr(five_user_iswrt, "#"&ugemid(u)&"_"&dtop&"#")) > 0 then
					antalu5 = antalu5 + 8
					antalux = antalu5
					else
					antalu5 = antalu5
					antalux = 0
					end if
					five_user_iswrt = five_user_iswrt &"#"&ugemid(u)&"_"&dtop&"#," 
					
					case 6
					leftc = 274
					if cint(instr(six_iswrt, "#"&dtop&"#")) > 0 then
					antal6 = antal6 + 8
					antalx = antal6
					else
					antal6 = antal6
					antalx = 0
					end if
					six_iswrt = six_iswrt &"#"&dtop&"#," 
					
					if cint(instr(six_user_iswrt, "#"&ugemid(u)&"_"&dtop&"#")) > 0 then
					antalu6 = antalu6 + 8
					antalux = antalu6
					else
					antalu6 = antalu6
					antalux = 0
					end if
					six_user_iswrt = six_user_iswrt &"#"&ugemid(u)&"_"&dtop&"#," 
					
					case 7
					leftc = 318
					if cint(instr(seven_iswrt, "#"&dtop&"#")) > 0 then
					antal7 = antal7 + 8
					antalx = antal7
					else
					antal7 = antal7
					antalx = 0
					end if
					seven_iswrt = seven_iswrt &"#"&dtop&"#," 
					
					if cint(instr(seven_user_iswrt, "#"&ugemid(u)&"_"&dtop&"#")) > 0 then
					antalu7 = antalu7 + 8
					antalux = antalu7
					else
					antalu7 = antalu7
					antalux = 0
					end if
					seven_user_iswrt = seven_user_iswrt &"#"&ugemid(u)&"_"&dtop&"#," 
					
					end select
					
					
					
					%>
					<div name="bookdiv_<%=ugemid(u)%>_<%=u%>" id="bookdiv_<%=ugemid(u)%>_<%=u%>" style="position:absolute; display:none; visibility:hidden; left:<%=leftc+antalux%>; top:<%=dtop+162%>; height:<%=dheight%>; background-color:<%=bgcol(u)%>; width:35; z-index:4000; border:1px #000000 solid; padding:2px;" onMousedown="bookdivon()" onMousemove="bookdivmove()" onMouseup="bookdivoff('<%=ugemid(u)%>','<%=u%>')">
					<table cellspacing=0 cellpadding=0 border=0><tr>
					<td valign=top><a href="#" onClick="bookminus('<%=ugemid(u)%>','<%=u%>')"><img src="../ill/pilop2.gif" width="11" height="10" alt="Minimer 15 min." border="0"></a></td>
					<td style="padding-left:1px;"><a href="#" onClick="showinfo(<%=u%>)"><img src="../ill/info2.gif" width="11" height="10" alt="Info" border="0"></a></td>
					<td valign=top style="padding-left:1px;"><a href="#" onClick="bookplus('<%=ugemid(u)%>','<%=u%>')"><img src="../ill/pilned2.gif" width="11" height="10" alt="Tilføj 15 min." border="0"></a></td>
					</tr>
					</table>
					</div>
					<input type="hidden" id="bookdiv_x_<%=ugemid(u)%>_<%=u%>" name="bookdiv_x_<%=ugemid(u)%>_<%=u%>" value="<%=(leftc)%>">
					<input type="hidden" id="bookdiv_y_<%=ugemid(u)%>_<%=u%>" name="bookdiv_y_<%=ugemid(u)%>_<%=u%>" value="<%=(dtop+162)%>">
					<input type="hidden" id="bookdiv_id_<%=ugemid(u)%>_<%=u%>" name="bookdiv_id_<%=ugemid(u)%>_<%=u%>" value="<%=divid(u)%>">
					<input type="hidden" id="bookdiv_height_<%=ugemid(u)%>_<%=u%>" name="bookdiv_height_<%=ugemid(u)%>_<%=u%>" value="<%=dheight%>">
					<%
					dspthis = ""
					vzthis = "visible"
					%>
					
					<div name="bookdiv_<%=ugemid(u)%>_0" id="bookdiv_<%=datepart("ww", ugedato(u))%>_<%=ugemid(u)%>_0" style="position:absolute; display:<%=dspthis%>; visibility:<%=vzthis%>; left:<%=(leftc)%>; top:<%=dtop+162+antalx%>; height:15; width:35; z-index:1000; border:1px #000000 solid; background-color:<%=bgcol(u)%>; padding:2;" onMouseover="cursorhand()" onClick="showinfo(<%=u%>)">
					<%
					thisfornavn = left(ugemnavn(u), 1)
					navn_instr = instr(ugemnavn(u), " ")
					navn_len = len(ugemnavn(u))
					navn_right = right(ugemnavn(u), (navn_len - (navn_instr)))
					if cint(navn_instr) > 0 then
					thisefternavn = left(navn_right, 1)
					else
					thisefternavn = ""
					end if
					
					initthis = thisfornavn & thisefternavn
					Response.write "<font class=megetmegetlillehvid>"& initthis
					%>
					<a href="#" onClick="showinfo(<%=u%>)" calss=alt><img src="../ill/info2.gif" width="11" height="10" alt="Info" border="0"></a>
					</font>
					</div>
					
					<div name="info_<%=u%>" id="info_<%=u%>" style="position:absolute; display:none; visibility:hidden; left:<%=(leftc)+10%>; top:<%=dtop+162+antalx+10%>; height:100; width:160; z-index:8000; border:1px silver solid; background-color:#ffffe1; padding:2px;">
					<table cellspacing=0 cellpadding=1 border=0 width=160>
					<tr>
						<td valign=top colspan=2 class="lille" align=right><a href="jbpla_w.asp?func=slet&id=<%=divid(u)%>&activeuge=<%=activeuge%>&yearthis=<%=yearthis%>&monththis=<%=monththis%>" target=w><img src="../ill/slet.gif" width="20" height="20" alt="Slet" border="0"></a></td>
					</tr>
					<tr>
						<td valign=top colspan=2 class="lille"><%=ugemnavn(u)%>  (<%=ugemnr(u)%>)</td>
					</tr>
					<tr>
						<td valign=top class="lille">Dato:</td><td valign=top class="lille"><%=ugedato(u)%></td>
					</tr>
					<tr>
						<td valign=top class="lille">Start:</td><td valign=top class="lille"><%=formatdatetime(ugesttid(u), 3)%></td>
					</tr>
					<tr>
						<td valign=top class="lille">Slut:</td><td valign=top class="lille"> <%=formatdatetime(ugesluttid(u), 3)%></td>
					</tr>
					<tr>
						<td valign=top class="lille">Job:</td><td valign=top class="lille"><%=left(ugejobnavn(u), 22)%> (<%=ugejobnr(u)%>)</td>
					</tr>
					<tr>	
						<td valign=top class="lille">Akt.:</td><td valign=top class="lille"><%=ugeanavn(u)%></td>
					</tr>
					<tr>	
						<td><a href="#" onClick="hideinfo(<%=u%>)" class=vmenu>Luk</a></td>
						<td align=right><a href="javascript:NewWin_opr('jbpla_w_opr.asp?func=red&sttid=<%=formatdatetime(ugesttid(u), 3)%>&dato=<%=ugedato(u)%>&datostkri=<%=jobStartKri%>&datoslkri=<%=jobSlutKri%>&jobid=<%=ugejobid(u)%>&medid=<%=ugemid(u)%>&aid=<%=ugeaid(u)%>&id=<%=divid(u)%>');" target="_self" class=vmenu><img src="../ill/soeg-knap.gif" width="16" height="16" alt="Rediger" border="0"></a></td>
					</tr>
					</table>
					</div>
					<%
			
			
	next
	
	
	%>
	<input type="hidden" name="antalu" id="antalu" value="<%=u-1%>">
	<input type="hidden" name="showlastmid" id="showlastmid" value="0">
	<div name="redsubmit" id="redsubmit" style="position:absolute; display:; visibility:visible; left:160; top:130;">
	<input type="submit" name="" value="Opdater uge <%=activeuge%>!" style="font-size:9px; border:1px #66CC33 solid;">
	</div>
	</form>
	
	

	<%
	end select 'func

end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
