<%
function timerdage(jobnrher, intMnr)
'**** timer på job son - lor ****
		days = 7
		for intcounter = 1 to days
			select case intcounter
			case 1
			usedatoher = varTjDatoUS_man
			case 2
			usedatoher = varTjDatoUS_tir
			case 3
			usedatoher = varTjDatoUS_ons
			case 4
			usedatoher = varTjDatoUS_tor
			case 5
			usedatoher = varTjDatoUS_fre
			case 6
			usedatoher = varTjDatoUS_lor
			case 7
			usedatoher = varTjDatoUS_son
			end select
			
			strSQL2 = "SELECT sum(timer) AS sumtimer FROM timer WHERE tjobnr = "& jobnrher &" AND tfaktim <> 5 AND Tmnr = "& intMnr &" AND tdato = '"& usedatoher &"' ORDER BY timer"
			'Response.write strSQL2
			oRec2.open strSQL2, oConn, 3
			if not oRec2.EOF then 
			timerbrugtdato = oRec2("sumtimer")
			end if
			oRec2.close
			
			if len(timerbrugtdato) > 0 then
			timerbrugtdato = timerbrugtdato 
			bgthis = "#d6dff5"
				if len(timerbrugtdato) > 2 then
				fontthis = "megetmegetlillesort"
				else
				fontthis = "megetlillesort"
				end if
			else
			timerbrugtdato = "<img src='../ill/blank.gif' width='10' height='10' alt='' border='0'>"
			bgthis = "#8caae6"
			fontthis = "megetlillesort"
			end if
		
		if intcounter = 1 then%>
		<table border=0 cellspacing=1 cellpadding=0 width=100><tr>
		<%end if%>
		<td align=center width=15 bgcolor="<%=bgthis%>"><font class='<%=fontthis%>'><%=timerbrugtdato%></td>
		<%if intcounter = 7 then%>
		</tr></table>
		<%end if
		
		
		next
end function




'**** funktion til at vise kommentar ikoner ****
'public vis
public kommTrue
public function showcommentfunc(dag)
		
		dag = dag
		select case dag
		case 1
		useKomm = sonKomm
		case 2
		useKomm = manKomm
		case 3
		useKomm = tirKomm
		case 4
		useKomm = onsKomm
		case 5
		useKomm = torKomm
		case 6
		useKomm = freKomm
		case 7
		useKomm = lorKomm
		end select
		
		'Response.write "usek: "& useKomm & "<br>"
		
			if len(useKomm) <> 0 then
			kommtrue = "+"
			else
			kommtrue = "<font color='#999999'>+</font>"
			end if 
end function
'***************************************************


	
	
	
	
	
	Public fakbgcol_son
	Public fakbgcol_man	
	Public fakbgcol_tir
	Public fakbgcol_ons	
	Public fakbgcol_tor 
	Public fakbgcol_fre 
	Public fakbgcol_lor
	
	Public maxl_son 
	Public maxl_man 
	Public maxl_tir 
	Public maxl_ons 
	Public maxl_tor 
	Public maxl_fre 
	Public maxl_lor 
	
	Public fmbg_son 
	Public fmbg_man 
	Public fmbg_tir 
	Public fmbg_ons
	Public fmbg_tor
	Public fmbg_fre 
	Public fmbg_lor
	
	
	
	
	
function fakfarver()
if len(sonTimerVal(iRowLoop)) <> 0 then
dtimeTtidspkt_son_this = dtimeTtidspkt_son(sonRLoop)
else
dtimeTtidspkt_son_this = "11:00:01 AM"
end if

if len(manTimerVal(iRowLoop)) <> 0 then
dtimeTtidspkt_man_this = dtimeTtidspkt_man(manRLoop)
else
dtimeTtidspkt_man_this = "11:00:01 AM"
end if

if len(tirTimerVal(iRowLoop)) <> 0 then
dtimeTtidspkt_tir_this = dtimeTtidspkt_tir(tirRLoop)
else
dtimeTtidspkt_tir_this = "11:00:01 AM"
end if

if len(onsTimerVal(iRowLoop)) <> 0 then
dtimeTtidspkt_ons_this = dtimeTtidspkt_ons(onsRLoop)
else
dtimeTtidspkt_ons_this = "11:00:01 AM"
end if

if len(torTimerVal(iRowLoop)) <> 0 then
dtimeTtidspkt_tor_this = dtimeTtidspkt_tor(torRLoop)
else
dtimeTtidspkt_tor_this = "11:00:01 AM"
end if

if len(freTimerVal(iRowLoop)) <> 0 then
dtimeTtidspkt_fre_this = dtimeTtidspkt_fre(freRLoop)
else
dtimeTtidspkt_fre_this = "11:00:01 AM"
end if

if len(lorTimerVal(iRowLoop)) <> 0 then
dtimeTtidspkt_lor_this = dtimeTtidspkt_lor(lorRLoop)
else
dtimeTtidspkt_lor_this = "11:00:01 AM"
end if



'** Sætter farve på indtastningfelt efter om der er udskrevet en faktura eller ej ***
diff = dateDiff("d", lastfakdato, varTjDatoUS_start)
'Response.write lastfakdato &" # " & varTjDatoUS_start & "<br>"
'Response.write "diff: " & diff &"  last fak dato:"& datepart("ww", lastfakdato) &"  Den aktuelle uge: "& strWeek &"<br>"
if len(lastfakdato) <> 0 AND diff <= 0 then
		'** Hvis fakuge > den valgte uge ***
          call thisWeekNo53_fn(lastfakdato)
            thisWeekNo53_lastfakdato = thisWeekNo53

		if thisWeekNo53_lastfakdato > strWeek then

			fakbgcol_son = "limegreen"
			fakbgcol_man = "limegreen"	
			fakbgcol_tir = "limegreen"	
			fakbgcol_ons = "limegreen"	
			fakbgcol_tor = "limegreen"	
			fakbgcol_fre = "limegreen"	
			fakbgcol_lor = "limegreen"
			
			maxl_son = 0
			maxl_man = 0
			maxl_tir = 0
			maxl_ons = 0
			maxl_tor = 0
			maxl_fre = 0
			maxl_lor = 0
			
			fmbg_son = "#CCCCCC" 
			fmbg_man = "#CCCCCC" 
			fmbg_tir = "#CCCCCC" 
			fmbg_ons = "#CCCCCC" 
			fmbg_tor = "#CCCCCC" 
			fmbg_fre = "#CCCCCC" 
			fmbg_lor = "#CCCCCC"  
	
		else
			'** Hvis fakuge = den valgte uge ***
			'** Tidspunkt sættes altid til 23:59:59 på fakturaer, fra d. 8/9-2004
            'lastfakdato = "10-05-2017"
			select case DatePart("w", lastfakdato, 2)
			case 7
			fakbgcol_son = "limegreen"
			maxl_son = 0
			fmbg_son = "#cccccc"
			
			fakbgcol_man = "limegreen"	
			fakbgcol_tir = "limegreen"	
			fakbgcol_ons = "limegreen"	
			fakbgcol_tor = "limegreen"	
			fakbgcol_fre = "limegreen"	
			fakbgcol_lor = "limegreen"

			maxl_man = 0
			maxl_tir = 0
			maxl_ons = 0
			maxl_tor = 0
			maxl_fre = 0
			maxl_lor = 0
			
			fmbg_man = "#cccccc" 
			fmbg_tir = "#cccccc" 
			fmbg_ons = "#cccccc" 
			fmbg_tor = "#cccccc" 
			fmbg_fre = "#cccccc" 
			fmbg_lor = "#cccccc"  
			
			
			case 1
			fakbgcol_man = "limegreen"
			maxl_man = 0
			fmbg_man = "#cccccc"
			
			
			
			fakbgcol_tir = "#7F9DB9"	
			fakbgcol_ons = "#7F9DB9"	
			fakbgcol_tor = "#7F9DB9"	
			fakbgcol_fre = "#7F9DB9"	
			fakbgcol_lor = "#cd853f"
			fakbgcol_son = "#cd853f"	

			maxl_son = 5
			maxl_tir = 5
			maxl_ons = 5
			maxl_tor = 5
			maxl_fre = 5
			maxl_lor = 5
			
			
			fmbg_tir = "#FFFFFF" 
			fmbg_ons = "#FFFFFF" 
			fmbg_tor = "#FFFFFF" 
			fmbg_fre = "#FFFFFF" 
			fmbg_lor = "#FFDFDF"
			fmbg_son = "#FFDFDF"   
			 
			case 2
			fakbgcol_tir = "limegreen"
			maxl_tir = 0
			fmbg_tir = "#cccccc"
			
			
			fakbgcol_man = "limegreen"	
			fakbgcol_ons = "#7F9DB9"	
			fakbgcol_tor = "#7F9DB9"	
			fakbgcol_fre = "#7F9DB9"	
			fakbgcol_lor = "#cd853f"
			fakbgcol_son = "#cd853f"	

			maxl_son = 5
			maxl_man = 0
			maxl_ons = 5
			maxl_tor = 5
			maxl_fre = 5
			maxl_lor = 5
			
			
			fmbg_man = "#cccccc" 
			fmbg_ons = "#FFFFFF" 
			fmbg_tor = "#FFFFFF" 
			fmbg_fre = "#FFFFFF" 
			fmbg_lor = "#FFDFDF"
			fmbg_son = "#FFDFDF"   
			 
			
			
			case 3
			fakbgcol_ons = "limegreen"
			maxl_ons = 0
			fmbg_ons = "#cccccc"
			
			
			
			
			fakbgcol_man = "limegreen"	
			fakbgcol_tir = "limegreen"	
			fakbgcol_tor = "#7F9DB9"	
			fakbgcol_fre = "#7F9DB9"	
			fakbgcol_lor = "#cd853f"
			fakbgcol_son = "#cd853f"	

			maxl_son = 5
			maxl_man = 0
			maxl_tir = 0
			maxl_tor = 5
			maxl_fre = 5
			maxl_lor = 5
			
			 
			fmbg_man = "#cccccc" 
			fmbg_tir = "#cccccc" 
			fmbg_tor = "#FFFFFF" 
			fmbg_fre = "#FFFFFF" 
			fmbg_lor = "#FFDFDF"
			fmbg_son = "#FFDFDF"  
			
			 
			case 4
			fakbgcol_tor = "limegreen"
			maxl_tor = 0
			fmbg_tor = "#cccccc"
			
			
			fakbgcol_man = "limegreen"	
			fakbgcol_tir = "limegreen"	
			fakbgcol_ons = "limegreen"	
			fakbgcol_fre = "#7F9DB9"	
			fakbgcol_lor = "#cd853f"
			fakbgcol_son = "#cd853f"

			
			maxl_man = 0
			maxl_tir = 0
			maxl_ons = 0
			maxl_fre = 5
			maxl_lor = 5
			maxl_son = 5
			
			
			fmbg_man = "#cccccc"  
			fmbg_tir = "#cccccc" 
			fmbg_ons = "#cccccc" 
			fmbg_fre = "#FFFFFF" 
			fmbg_lor = "#FFDFDF"
			fmbg_son = "#FFDFDF"   
			
			
			case 5
			fakbgcol_fre = "limegreen"
			maxl_fre = 0
			fmbg_fre = "#cccccc"
			
			
			
			fakbgcol_man = "limegreen"	
			fakbgcol_tir = "limegreen"	
			fakbgcol_ons = "limegreen"	
			fakbgcol_tor = "limegreen"	
			fakbgcol_lor = "#cd853f"
			fakbgcol_son = "#cd853f"

			maxl_son = 5
			maxl_man = 0
			maxl_tir = 0
			maxl_ons = 0
			maxl_tor = 0
			maxl_lor = 5
			
			
			fmbg_man = "#cccccc"  
			fmbg_tir = "#cccccc" 
			fmbg_ons = "#cccccc" 
			fmbg_tor = "#cccccc" 
			fmbg_lor = "#FFDFDF"
			fmbg_son = "#FFDFDF" 
			
			
			case 6
			fakbgcol_lor = "limegreen"
			maxl_lor = 0
			fmbg_lor = "#cccccc"
			
		
			fakbgcol_son = "#cd853f"
			fakbgcol_man = "limegreen"	
			fakbgcol_tir = "limegreen"	
			fakbgcol_ons = "limegreen"	
			fakbgcol_tor = "limegreen"	
			fakbgcol_fre = "limegreen"

			maxl_son = 5
			maxl_man = 0
			maxl_tir = 0
			maxl_ons = 0
			maxl_tor = 0
			maxl_fre = 0
			
			fmbg_son = "#FFDFDF"
			fmbg_man = "#cccccc"  
			fmbg_tir = "#cccccc" 
			fmbg_ons = "#cccccc" 
			fmbg_tor = "#cccccc" 
			fmbg_fre = "#cccccc"  
			end select
		end if
else
	'** nofakfarver ***
	fakbgcol_son = "#cd853f"
	fakbgcol_man = "#7F9DB9"	
	fakbgcol_tir = "#7F9DB9"	
	fakbgcol_ons = "#7F9DB9"	
	fakbgcol_tor = "#7F9DB9"	
	fakbgcol_fre = "#7F9DB9"	
	fakbgcol_lor = "#cd853f"
	
	maxl_son = 5
	maxl_man = 5
	maxl_tir = 5
	maxl_ons = 5
	maxl_tor = 5
	maxl_fre = 5
	maxl_lor = 5
	
	fmbg_son = "#FFDFDF" 
	fmbg_man = "#FFFFFF" 
	fmbg_tir = "#FFFFFF" 
	fmbg_ons = "#FFFFFF" 
	fmbg_tor = "#FFFFFF" 
	fmbg_fre = "#FFFFFF" 
	fmbg_lor = "#FFDFDF"  
	
	
	'*** Markerer dagsdato *** 
	select case weekday(useDate)
	case 1
	fmbg_son = "#eff3ff" 
	case 2
	fmbg_man = "#eff3ff" 
	case 3
	fmbg_tir = "#eff3ff" 
	case 4
	fmbg_ons = "#eff3ff"
	case 5
	fmbg_tor = "#eff3ff" 
	case 6
	fmbg_fre = "#eff3ff" 
	case 7
	fmbg_lor = "#eff3ff"
	end select
	
	
end if	
	
end function
%>
