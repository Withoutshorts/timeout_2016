<%
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
'Response.write "diff:" & diff &"fakd:"& datepart("ww", lastfakdato) &"week "& strWeek &"<br>"
if len(lastfakdato) <> 0 AND diff <= 0 then
		'** Hvis fakuge > den valgte uge ***
		if datepart("ww", lastfakdato) > strWeek then
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
			select case DatePart("w", lastfakdato)
			case 1
			if formatdatetime(dtimeTtidspkt_son_this, 3) < lastFaktime AND len(sonTimerVal(sonRLoop)) <> 0 then
			fakbgcol_son = "limegreen"
			maxl_son = 0
			fmbg_son = "#cccccc"
			else
			fakbgcol_son = "#cd853f"
			maxl_son = 4
			fmbg_son = "#FFDFDF"
			end if 
			
			
			fakbgcol_man = "#7F9DB9"	
			fakbgcol_tir = "#7F9DB9"	
			fakbgcol_ons = "#7F9DB9"	
			fakbgcol_tor = "#7F9DB9"	
			fakbgcol_fre = "#7F9DB9"	
			fakbgcol_lor = "#cd853f"

			maxl_man = 5
			maxl_tir = 5
			maxl_ons = 5
			maxl_tor = 5
			maxl_fre = 5
			maxl_lor = 5
			
			fmbg_man = "#FFFFFF" 
			fmbg_tir = "#FFFFFF" 
			fmbg_ons = "#FFFFFF" 
			fmbg_tor = "#FFFFFF" 
			fmbg_fre = "#FFFFFF" 
			fmbg_lor = "#FFDFDF"  
			
			
			case 2
			if formatdatetime(dtimeTtidspkt_man_this, 3) < lastFaktime AND len(manTimerVal(manRLoop)) <> 0 then
			fakbgcol_man = "limegreen"
			maxl_man = 0
			fmbg_man = "#cccccc"
			else
			fakbgcol_man = "#7F9DB9"
			maxl_man = 4
			fmbg_man = "#FFFFFF"
			end if 
			
			
			fakbgcol_son = "limegreen"	
			fakbgcol_tir = "#7F9DB9"	
			fakbgcol_ons = "#7F9DB9"	
			fakbgcol_tor = "#7F9DB9"	
			fakbgcol_fre = "#7F9DB9"	
			fakbgcol_lor = "#cd853f"

			maxl_son = 0
			maxl_tir = 5
			maxl_ons = 5
			maxl_tor = 5
			maxl_fre = 5
			maxl_lor = 5
			
			fmbg_son = "#cccccc" 
			fmbg_tir = "#FFFFFF" 
			fmbg_ons = "#FFFFFF" 
			fmbg_tor = "#FFFFFF" 
			fmbg_fre = "#FFFFFF" 
			fmbg_lor = "#FFDFDF"  
			 
			case 3
			if formatdatetime(dtimeTtidspkt_tir_this, 3) < lastFaktime AND len(tirTimerVal(tirRLoop)) <> 0 then
			fakbgcol_tir = "limegreen"
			maxl_tir = 0
			fmbg_tir = "#cccccc"
			else
			fakbgcol_tir = "#7F9DB9"
			maxl_tir = 4
			fmbg_tir = "#FFFFFF"
			end if 
			
			
			fakbgcol_son = "limegreen"	
			fakbgcol_man = "limegreen"	
			fakbgcol_ons = "#7F9DB9"	
			fakbgcol_tor = "#7F9DB9"	
			fakbgcol_fre = "#7F9DB9"	
			fakbgcol_lor = "#cd853f"

			maxl_son = 0
			maxl_man = 0
			maxl_ons = 5
			maxl_tor = 5
			maxl_fre = 5
			maxl_lor = 5
			
			fmbg_son = "#cccccc" 
			fmbg_man = "#cccccc" 
			fmbg_ons = "#FFFFFF" 
			fmbg_tor = "#FFFFFF" 
			fmbg_fre = "#FFFFFF" 
			fmbg_lor = "#FFDFDF"  
			 
			case 4
			if formatdatetime(dtimeTtidspkt_ons_this, 3) < lastFaktime AND len(onsTimerVal(onsRLoop)) <> 0 then
			fakbgcol_ons = "limegreen"
			maxl_ons = 0
			fmbg_ons = "#cccccc"
			else
			fakbgcol_ons = "#7F9DB9"
			maxl_ons = 4
			fmbg_ons = "#FFFFFF"
			end if 
			
			
			fakbgcol_son = "limegreen"	
			fakbgcol_man = "limegreen"	
			fakbgcol_tir = "limegreen"	
			fakbgcol_tor = "#7F9DB9"	
			fakbgcol_fre = "#7F9DB9"	
			fakbgcol_lor = "#cd853f"

			maxl_son = 0
			maxl_man = 0
			maxl_tir = 0
			maxl_tor = 5
			maxl_fre = 5
			maxl_lor = 5
			
			fmbg_son = "#cccccc" 
			fmbg_man = "#cccccc" 
			fmbg_tir = "#cccccc" 
			fmbg_tor = "#FFFFFF" 
			fmbg_fre = "#FFFFFF" 
			fmbg_lor = "#FFDFDF"  
			
			 
			case 5
			if formatdatetime(dtimeTtidspkt_tor_this, 3) < lastFaktime AND len(torTimerVal(torRLoop)) <> 0 then
			fakbgcol_tor = "limegreen"
			maxl_tor = 0
			fmbg_tor = "#cccccc"
			else
			fakbgcol_tor = "#7F9DB9"
			maxl_tor = 4
			fmbg_tor = "#FFFFFF"
			end if
			
			
			fakbgcol_son = "limegreen"
			fakbgcol_man = "limegreen"	
			fakbgcol_tir = "limegreen"	
			fakbgcol_ons = "limegreen"	
			fakbgcol_fre = "#7F9DB9"	
			fakbgcol_lor = "#cd853f"

			maxl_son = 0
			maxl_man = 0
			maxl_tir = 0
			maxl_ons = 0
			maxl_fre = 5
			maxl_lor = 5
			
			fmbg_son = "#cccccc"
			fmbg_man = "#cccccc"  
			fmbg_tir = "#cccccc" 
			fmbg_ons = "#cccccc" 
			fmbg_fre = "#FFFFFF" 
			fmbg_lor = "#FFDFDF"   
			case 6
			if formatdatetime(dtimeTtidspkt_fre_this, 3) < lastFaktime AND len(freTimerVal(freRLoop)) <> 0 then
			fakbgcol_fre = "limegreen"
			maxl_fre = 0
			fmbg_fre = "#cccccc"
			else
			fakbgcol_fre = "#7F9DB9"
			maxl_fre = 4
			fmbg_fre = "#FFFFFF"
			end if
			
			
			fakbgcol_son = "limegreen"
			fakbgcol_man = "limegreen"	
			fakbgcol_tir = "limegreen"	
			fakbgcol_ons = "limegreen"	
			fakbgcol_tor = "limegreen"	
			fakbgcol_lor = "#cd853f"

			maxl_son = 0
			maxl_man = 0
			maxl_tir = 0
			maxl_ons = 0
			maxl_tor = 0
			maxl_lor = 5
			
			fmbg_son = "#cccccc"
			fmbg_man = "#cccccc"  
			fmbg_tir = "#cccccc" 
			fmbg_ons = "#cccccc" 
			fmbg_tor = "#cccccc" 
			fmbg_lor = "#FFDFDF" 
			
			
			case 7
			if formatdatetime(dtimeTtidspkt_lor_this, 3) < lastFaktime AND len(lorTimerVal(lorRLoop)) <> 0 then
			fakbgcol_lor = "limegreen"
			maxl_lor = 0
			fmbg_lor = "#cccccc"
			else
			fakbgcol_lor = "#7F9DB9"
			maxl_lor = 4
			fmbg_lor = "#FFFFFF"
			end if
		
			
			fakbgcol_son = "limegreen"
			fakbgcol_man = "limegreen"	
			fakbgcol_tir = "limegreen"	
			fakbgcol_ons = "limegreen"	
			fakbgcol_tor = "limegreen"	
			fakbgcol_fre = "limegreen"

			maxl_son = 0
			maxl_man = 0
			maxl_tir = 0
			maxl_ons = 0
			maxl_tor = 0
			maxl_fre = 0
			
			fmbg_son = "#cccccc"
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
end if	

%>
