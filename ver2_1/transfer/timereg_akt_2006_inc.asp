<%

function SQLBless(s)
dim tmp
tmp = s
tmp = replace(tmp, ",", ".")
SQLBless = tmp
end function

function SQLBless2(s2)
dim tmp2
tmp2 = s2
tmp2 = replace(tmp2, "'", "''")
SQLBless2 = tmp2
end function


function SQLBless3(s3)
dim tmp3
tmp3 = s3
tmp3 = replace(tmp3, ":", "")
SQLBless3 = tmp3
end function


'*** Finder evt. alternativ timepris på medarbejder ****
'*** Hvis der ikke findes alternative timepriser bruges default ****
public intTimepris, foundone, intValuta
function alttimepris(useaktid, intjobid, strMnr)
                            
                            intTimepris = 0
							timeprisAlt = 0
							foundone = "n"
							strSQL = "SELECT id AS tpid, jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = "& intjobid &" AND aktid = "& useaktid &" AND medarbid =  "& strMnr
								
						    oRec2.open strSQL, oConn, 3 
							if not oRec2.EOF then
							        
							        foundone = "y"
				                    intTimepris = oRec2("6timepris")
				                    intValuta = oRec2("6valuta")
									tpid = oRec2("tpid")
									timeprisAlt = oRec2("timeprisalt")
									
									if cdbl(intTimepris) = 0 AND cint(timeprisAlt) <> 6 then
									
							
							            Select case timeprisAlt
							            case 1
							            timeprisalernativ = "timepris_a1"
							            valutaAlt = "tp1_valuta"
							            case 2
							            timeprisalernativ = "timepris_a2"
							            valutaAlt = "tp2_valuta"
							            case 3
							            timeprisalernativ = "timepris_a3"
							            valutaAlt = "tp3_valuta"
							            case 4
							            timeprisalernativ = "timepris_a4"
							            valutaAlt = "tp4_valuta"
							            case 5
							            timeprisalernativ = "timepris_a5"
							            valutaAlt = "tp5_valuta"
							           
							            case else
							            timeprisalernativ = "timepris"
							            valutaAlt = "tp0_valuta"
							            end select
							            
							            
							        strSQL3 = "SELECT mid, "&timeprisalernativ&" AS useTimepris, "& valutaAlt &" AS useValuta, medarbejdertype FROM medarbejdere, medarbejdertyper WHERE mid =" & strMnr & " AND medarbejdertyper.id = medarbejdertype"
									
									'Response.Write strSQL3
									'Response.flush
									oRec3.open strSQL3, oConn, 3 
									if not oRec3.EOF then
									intTimepris = oRec3("useTimepris")
									intValuta = oRec3("useValuta")
									end if
									oRec3.close
									
									intTimepris = replace(intTimepris, ",", ".")
									
									
									strSQLupdtp = "UPDATE timepriser SET 6timepris = "& intTimepris&", 6valuta = "& intValuta &" WHERE id = " & tpid
									oConn.execute(strSQLupdtp)
									
									
									
									end if
									
						    
						    end if 
							oRec2.close 
						
						
						
end function



'****** Medarbejder timepris og kostpris ********************
public strMnavn, dblkostpris, tprisGen, valutaGen 
function mNavnogKostpris(strMnr)
'** Henter navn og kostpris ***'

SQLmedtpris = "SELECT medarbejdertype, timepris, tp0_valuta, kostpris, mnavn FROM medarbejdere, medarbejdertyper "_
&" WHERE Mid = "& strMnr &" AND medarbejdertyper.id = medarbejdertype"

'Response.Write SQLmedtpris
'Response.flush
oRec.Open SQLmedtpris, oConn, 3

		if Not oRec.EOF then
		 	
		 	if oRec("kostpris") <> "" then
			dblkostpris = oRec("kostpris")
			else
			dblkostpris = 0
			end if
		
		strMnavn = oRec("mnavn")
		tprisGen = oRec("timepris")
		valutaGen = oRec("tp0_valuta")
		
		end if

oRec.close
'** Slut timepris **


end function

'**** Opdaterer timer ****
function opdaterimer(aktid, aktnavn, tfaktimvalue, strFastpris, jobnr, strJobnavn, strJobknr, strJobknavn,_
medid, strMnavn, datothis, timerthis, kommthis, intTimepris,_
dblkostpris, offentlig, intServiceAft, strYear, sTtid, sLtid, visTimerelTid, stopur, intValuta)

    
    call valutaKurs(intValuta)
    
    '*** Timer på Afspad og ferie ***'
    if lto = "execon999" then 'OR lto = "intranet - local"
    select case tfaktimvalue
    case 11,12,13,14
    
    if timerthis >= 6.9 then
         '*** Normeret uge  til ferie beregning (gns. fuld dag) ***'
	    call normtimerPer(medid, datothis, 7)
	    
	    timerthis = formatnumber(ntimPer/5,2)
	    timerthis = replace(timerthis, ".", "")
	    timerthis = replace(timerthis, ",", ".")
    end if
    
    end select
    
    'Response.Write "tiemr:"& timerthis
    'Response.flush
    
    end if
    
			
	if cint(timerthis) <> -9001 then '*** <> Slet ***'
			
			datothis = ConvertDateYMD(datothis)
			
					if visTimerelTid <> 0 AND len(trim(sTtid)) <> 0 then
						
						'Response.write len(trim(sTtid)) &" - "& sTtid &"<br>"
						'Response.flush
						
						if len(trim(sTtid)) = 1 then '** slet
						
						timerthis = 0
						
						else
						idag = day(now)&"/"&month(now)&"/"&year(now)
						totalmin = datediff("n", idag &" "& sTtid, idag &" "& sLtid)
						call timerogminutberegning(totalmin)
						timerthis = thoursTot&"."&tminProcent 'tminTot
						
						sTtid = sTtid &":00"
						sLtid = sLtid &":00"
						end if
					
					else
					
					sTtid = ""
					sLtid = ""
					
					end if
			
			'Tjek om job/akt allerede er registreret for den valgte dato
			strSQLfindes = "SELECT timer, Tjobnr, TAktivitetId, timerkom FROM timer WHERE TAktivitetId = "& aktid &" AND Tjobnr = "& jobnr &" AND Tmnr = "& medid &" AND Tdato = '"& datothis &"'"
			oRec.Open strSQLfindes, oConn, 3  
			
			'son, man, tir, ons ,tor, fre, lor
			if oRec.EOF then
				
				
				if len(timerthis) > 0 AND cint(timerthis) > -9000 AND cint(timerthis) <> 0 then 
				
				'** AND timerthis > 0  then -minus er OK HUSK ret i faktrua opr. 15.05.2008 **'
				'** Rettet tilbage 20.05.2008 pga fejl i SponsorCar ***' 
				'** De bruger st. tid og slut tid regsistrering **'
				
				strSQLins = "INSERT INTO timer (TJobnr, TJobnavn, Tmnr, Tmnavn, Tdato, "_
				&" Timer, Timerkom, Tknavn, Tknr, TAktivitetId, TAktivitetNavn, "_
				&" Tfaktim, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
				&" editor, kostpris, offentlig, seraft, sttid, sltid, valuta, kurs) VALUES"_
				& "(" & jobnr & ", '"& strJobnavn &"', " & medid & ", '" & Cstr(strMnavn) & "', '"& datothis &"', "_
				&" "& timerthis &", '"& SQLBless2(kommthis) &"', "_
				&" '" & SQLBless2(strJobknavn) & "', " & strJobknr & ", "_
				&" "& aktid &", '"& SQLBless2(aktnavn) &"', "_
				&" "& tfaktimvalue &", "& strYear &", "& SQLBless(intTimepris) &", "_
				&" '"&year(now)&"/"&month(now)&"/"&day(now)&"', '"& strFastpris  &"', "_
				&" '" & time & "', '"& session("user") &"', "& dblkostpris &", "_
				&" "& offentlig &", "& intServiceAft &", '"&sTtid&"', '"&sLtid&"', "& intValuta &", "& dblKurs &")"
				end if
				
				
				if len(trim(strSQLins)) <> 0 then
				    'Response.Write "Ins: "& strSQLins & "<br>"
				    'Response.flush
					oConn.Execute(strSQLins)
				end if
			
			else
				
					'** Sletter ***
					if cint(timerthis) = 0 then
					
					oConn.execute("DELETE FROM timer"_
					&" WHERE Tjobnr = "& jobnr & ""_
					&" AND Tmnr = "& medid & ""_
					&" AND Tdato = '" & datothis & "' AND TAktivitetId = "& aktid & "")
					
					
					else
					
						'** Opdaterer ****
						if (cdbl(SQLBless(oRec("timer"))) <> cdbl(timerthis)) OR stopur = "1" then
						
						        if stopur = "1" then 
						        '** tilføjer timer til eksisterende timer istedet for at overskrive **'
						        eksiTimer = replace(oRec("timer"),".",",")
						        timerthis = replace(timerthis, ".", ",")
						        timerthis = ((timerthis * 1) + (eksiTimer * 1)/1)
						        timerthis = replace(timerthis, ",", ".")
						        kommentarthis = replace(oRec("timerkom"), "'", "''") &" "& kommthis
						        'Response.Write "timerthis:" & timerthis 
						        'Response.end
						        else
						        kommentarthis = kommthis 'replace(oRec("timerkom"), "'", "''") 
						        end if
						
						strSQLupd = "UPDATE timer SET"_
						&" Timer = "& timerthis &", "_
						&" Taar = "& strYear &", TimePris = "& SQLBless(intTimepris) &", "_
						&" seraft = "& intServiceAft &", kostpris = "& dblkostpris &", "_
						&" Timerkom = '"& kommentarthis &"', "_
						&" tastedato = '"&year(now)&"/"&month(now)&"/"&day(now)&"', "_
						&" editor = '"& session("user") &"', offentlig = "& offentlig &", "_
						&" sttid = '"&sTtid&"', sltid = '"&sLtid&"', valuta = "& intValuta &", kurs = "& dblKurs &""_
						&" WHERE Tjobnr = "& jobnr & ""_
						&" AND Tmnr = "& medid & ""_
						&" AND Tdato = '" & datothis & "' AND TAktivitetId = "& aktid
						
						else
						
						'*************************************************************************************************'
						'** Opdatering af reg. tidspunkt på eksisterende timeregistreringer                            **'
						'** Tidspukter bliver ikke opdateret her, da man ellers vil overskrive tidspunkter              **'
						'** på eksisterende records, ved opdatering af en vilkårlig anden timeregistrering (som timer)  **'
						'*************************************************************************************************'
						
						
						strSQLupd = "UPDATE timer SET"_
						&" Timer = "& timerthis &", Timerkom = '"& kommthis &"', "_
						&" editor = '"& session("user") &"', offentlig = "& offentlig &""_
						&" WHERE Tjobnr = "& jobnr & ""_
						&" AND Tmnr = "& medid & ""_
						&" AND Tdato = '" & datothis & "' AND TAktivitetId = "& aktid 
						
						end if
						
						'Response.write "Upd: "& strSQLupd & "<br>"
					    'Response.end
						oConn.Execute(strSQLupd)
						
					end if
			end if
			oRec.close
			
	end if '** timer -9001 (-1) ***'

end function




public maxl, fmbgcol, fmborcl 
function fakfarver(lastfakdato, varTjDato_man, varTjDato_ugedag)

'** Sætter farve på indtastningfelt efter om der er udskrevet en faktura eller ej ***
strWeek = datepart("ww", tjekdag(4), 2, 2)
diff = dateDiff("d", lastfakdato, tjekdag(1))
ugeafsluttet = 0


'Response.Write "UgadagDato: "& varTjDato_ugedag &" Uge afl: " &ugeNrAfsluttet& "# " & datepart("ww", ugeNrAfsluttet, 2, 2) &" = "& strWeek &" AND "& smilaktiv &" AND "& autogk &" AND "& ugeNrAfsluttet &" autolukvdatodato "& autolukvdatodato &"<br>"
'Response.Write "herder45"
'Response.Write day(now) &" > "& autolukvdatodato  &" AND "& DatePart("yyyy", varTjDato_ugedag) &" = "& year(now) & " AND "& DatePart("m", varTjDato_ugedag) &" < "& month(now) & "<br>"
'Response.Write "aktdata(iRowLoop, 16): " & aktdata(iRowLoop, 16) & "<br>"

    if (datepart("ww", ugeNrAfsluttet, 2, 2) = strWeek AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", varTjDato_ugedag) = year(now) AND DatePart("m", varTjDato_ugedag) < month(now)) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", varTjDato_ugedag) < year(now) AND DatePart("m", varTjDato_ugedag) = 12)) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", varTjDato_ugedag) < year(now) AND DatePart("m", varTjDato_ugedag) <> 12) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", varTjDato_ugedag) > 1))) then


         '*** Uge afsluttet via Smiley Ordning og autogokend slået til i kontrolpanel **'
        '*** hvis admin level 1 kan timer stadigvæk redigeres indtil der oprettes faktura ***'
	    if level <> 1 then
	    maxl = 0
	    fmbgcol = "#cccccc"
	    fmborcl = "1px #999999 dashed"
	    else
	    maxl = 5
	    fmbgcol = "#FFFFFF"
	    fmborcl = "1px #999999 dashed"
	    end if
    	
	    ugeafsluttet = 1
    	
     end if
            
            
                '** ER registrering godkendt af joabans eller admin ***'
                if len(trim(aktdata(iRowLoop, 16))) <> 0 then
                aktdata(iRowLoop, 16) = aktdata(iRowLoop, 16)
                else
                aktdata(iRowLoop, 16) = 0
                end if
                
                if cint(aktdata(iRowLoop, 16)) = 1 then
                
                    maxl = 0
	                fmbgcol = "#cccccc"
	                fmborcl = "1px #999999 dashed"
	                
	                ugeafsluttet = 1
    						        
                end if   
                
                
                
            
            '**** Findes der en faktura ****'
            '**** Overskriver Smiley luk / godkendte uger ****'
			if len(lastfakdato) <> 0 AND diff <= 0 AND cint(aktdata(iRowLoop, 16)) = 0 then
					        '** Hvis fakuge > den valgte uge ***
					        if datepart("ww", lastfakdato, 2, 2) > strWeek AND datepart("yyyy", lastfakdato) => datepart("yyyy", tjekdag(4)) then
        					
        					    maxl = 0
						        fmbgcol = "#cccccc"
						        fmborcl = "1px #999999 dashed"
        						
        				    else
        					
        					 
        					
						        '** Hvis fakuge = den valgte uge ***
						        '** Tidspunkt sættes altid til 23:59:59 på fakturaer, fra d. 8/9-2004
						        'Response.write DatePart("y", lastfakdato) &" >=  " & DatePart("y", varTjDato_ugedag) & "<br>" 
						        if DatePart("y", lastfakdato) >= DatePart("y", varTjDato_ugedag) AND DatePart("yyyy", lastfakdato) >= DatePart("yyyy", varTjDato_ugedag) then
        					
        							
							        maxl = 0
							        fmbgcol = "#cccccc"
							        fmborcl = "1px #999999 dashed"
        						
						        else
        							
        							
								        if ugeafsluttet = 1 then
				                        maxl = maxl 
				                        fmbgcol = fmbgcol
				                        fmborcl = fmborcl
				                        else	
				                        maxl = 5
				                        fmbgcol = "#ffffff"
				                        fmborcl = "1px #999999 solid"
				                        end if
        							
        							
        							
						        end if
        					
					        end if
			
			else
				
				if ugeafsluttet = 1 then
				maxl = maxl 
				fmbgcol = fmbgcol
				fmborcl = fmborcl
				else	
				maxl = 5
				fmbgcol = "#ffffff"
				fmborcl = "1px #999999 solid"
				end if
						
			end if	
			
			
			


	
end function



function logintimerUge()

	'**** login historik (denne uge) ****
	for intcounter = 1 to 7
	
					select case intcounter
					case 1
					useSQLd = varTjDatoUS_man
					case 2
					useSQLd = varTjDatoUS_tir
					case 3
					useSQLd = varTjDatoUS_ons
					case 4
					useSQLd = varTjDatoUS_tor
					case 5
					useSQLd = varTjDatoUS_fre
					case 6
					useSQLd = varTjDatoUS_lor
					case 7
					useSQLd = varTjDatoUS_son
					end select 
					
					
					strSQL = "SELECT l.id AS lid, l.mid AS lmid, l.login, l.logud, l.minutter, "_
					&" s.navn AS stempelurnavn, s.faktor, s.minimum, stempelurindstilling FROM login_historik l"_
					&" LEFT JOIN stempelur s ON (s.id = l.stempelurindstilling) WHERE "_
					&" l.dato = '"& useSQLd &"' AND l.mid = " & useMrn &""_
					&" ORDER BY l.login" 
					
					'Response.write strSQL
					'Response.flush
					
					x = 0
					oRec.open strSQL, oConn, 3 
					while not oRec.EOF 
					
						timerThis = 0
						timerThisDIFF = 0
						
						if len(oRec("login")) <> 0 AND len(oRec("logud")) <> 0 then
						
						'loginTidAfr = left(formatdatetime(oRec("login"), 3), 5)
						'logudTidAfr = left(formatdatetime(oRec("logud"), 3), 5)
						
						if cint(oRec("stempelurindstilling")) = -1 then
						    
						    timerThisDIFF = oRec("minutter")
						    useFaktor = 0
						    
						    timerThisPause = timerThisDIFF
						    timerThis = 0
						    
						else 
						    
						    timerThisDIFF = oRec("minutter") 'datediff("s", loginTidAfr, logudTidAfr)/60
						
						    if timerThisDIFF < oRec("minimum") then
							    timerThisDIFF = oRec("minimum")
						    end if
						
							
							if oRec("faktor") > 0 then
							useFaktor = oRec("faktor")
							else
							useFaktor = 0
							end if
							
							timerThisPause = 0
							timerThis = (timerThisDIFF * useFaktor)
							
						end if
						
						
						
						'totaltimer = totaltimer + timerThis
						'Response.write oRec("lid") & ": " &  timerThis &" - "
						end if
						
						
						timTemp = formatnumber(timerThis/60, 3)
						timTemp_komma = split(timTemp, ",")
						
						for x = 0 to UBOUND(timTemp_komma)
							
							if x = 0 then
							thours = timTemp_komma(x)
							end if
							
							if x = 1 then
							tmin = timerThis - (thours * 60)
							end if
							
						next
						
						'totalhours = totalhours + cint(thours)
						'totalmin = totalmin + tmin
					
						select case intcounter
						case 1
						manMin = manMin + timerThis 
						manMinPause = manMinPause + timerThisPause 
						case 2
						tirMin = tirMin + timerThis 
						tirMinPause = tirMinPause + timerThisPause  
						case 3
						onsMin = onsMin + timerThis
						onsMinPause = onsMinPause + timerThisPause   
						case 4
						torMin = tormin + timerThis
						torMinPause = torminPause + timerThisPause   
						case 5
						freMin = freMin + timerThis
						freMinPause = freMinPause + timerThisPause   
						case 6
						lorMin = lorMin + timerThis
						lorMinPause = lorMinPause + timerThisPause  
						case 7
						sonMin = sonMin + timerThis
						sonMinPause = sonMinPause + timerThisPause   
						end select 
						
						'thours = 0
						'tmin = 0
						
					oRec.movenext
					wend
					oRec.close 
	
	next
	%>


	<table cellspacing=1 cellpadding=2 border=0>
	<tr><td colspan=9><img src="../ill/ac0018-16.gif" width="16" height="16" alt="" border="0">&nbsp;
	<a href="stempelur.asp?func=stat&medarbSel=<%=useMrn%>&showonlyone=1&hidemenu=1" target="_blank"; class=vmenu><%=tsa_txt_123 %></a>&nbsp;- 
		<%=tsa_txt_005 %>: <%=datepart("ww", varTjDatoUS_man, 2, 2)%>
		<br />
		<%=tsa_txt_134 %>: 
		<%
		
		sLoginTid = "00:00:00"
		
		strSQL = "SELECT l.id AS lid, l.login "_
		&" FROM login_historik l WHERE "_
		&" l.mid = " & useMrn &" AND stempelurindstilling <> -1"_
		&" ORDER BY l.id DESC" 
					
		'Response.write strSQL
		'Response.flush
		
		
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then 
        
        sLoginTid = oRec("login") 
        
        end if
        oRec.close
		
		%>
		<b><%=formatdatetime(sLoginTid, 3) %></b>
		<% 
		logindiffSidste = datediff("n", sLoginTid, now, 2, 3)
		%>
		<br /><%=tsa_txt_135 %>: 
		<%call timerogminutberegning(logindiffSidste) %>
		<b><%=thoursTot &":"& left(tminTot, 2) %>&nbsp;<%=tsa_txt_141 %></b>
		
		
		</td></tr>
	<tr bgcolor="#d6dff5">
		<td>
            &nbsp;</td>
		
		<td width=40 valign=bottom align=center><b><%=tsa_txt_128 %></b></td>
		<td width=40 valign=bottom align=center><b><%=tsa_txt_129 %></b></td>
		<td width=40 valign=bottom align=center><b><%=tsa_txt_130 %></b></td>
		<td width=40 valign=bottom align=center><b><%=tsa_txt_131 %></b></td>
		<td width=40 valign=bottom align=center><b><%=tsa_txt_132 %></b></td>
	    <td width=40 valign=bottom align=center><b><%=tsa_txt_133 %></b></td>
		<td width=40 valign=bottom align=center><b><%=tsa_txt_127 %></b></td>
		<td width=40 bgcolor="#ffdfdf" valign=bottom align=right><b><%=tsa_txt_136 %></b></td>
	</tr>
	<tr bgcolor="#ffffff">
		<td><%=tsa_txt_137 %>:</td>
		<td valign=top align=right><%call timerogminutberegning(manMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  manMin%>
		<td valign=top align=right><%call timerogminutberegning(tirMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  tirMin%>
		<td valign=top align=right><%call timerogminutberegning(onsMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  onsMin%>
		<td valign=top align=right><%call timerogminutberegning(torMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  torMin%>
		<td valign=top align=right><%call timerogminutberegning(freMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  freMin%>
		<td valign=top align=right><%call timerogminutberegning(lorMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  lorMin%>
		<td valign=top align=right><%call timerogminutberegning(sonMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  sonMin%>
		<td valign=top align=right><%call timerogminutberegning(ugeIalt)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<tr bgcolor="lightgrey">
		<td><%=tsa_txt_138 %>:</td>
		<td valign=top align=right><%call timerogminutberegning(manMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause +  manMinPause%>
		<td valign=top align=right><%call timerogminutberegning(tirMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause +  tirMinPause%>
		<td valign=top align=right><%call timerogminutberegning(onsMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause +  onsMinPause%>
		<td valign=top align=right><%call timerogminutberegning(torMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause +  torMinPause%>
		<td valign=top align=right><%call timerogminutberegning(freMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause +  freMinPause%>
		<td valign=top align=right><%call timerogminutberegning(lorMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause +  lorMinPause%>
		<td valign=top align=right><%call timerogminutberegning(sonMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause +  sonMinPause%>
		<td valign=top align=right><%call timerogminutberegning(ugeIaltPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	<tr bgcolor="#ffdfdf">
		<td><b><%=tsa_txt_136 %>:</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(manMin - (manMinPause))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(tirMin - (tirMinPause))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(onsMin - (onsMinPause))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(torMin -(torMinPause))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
	    <td valign=top align=right><b><%call timerogminutberegning(freMin -(freMinPause))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(lorMin - (lorMinPause))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(sonMin - (sonMinPause))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(ugeIalt - (ugeIaltPause))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<%loginTimerTot = ugeIalt - (ugeIaltPause)  %>
	</tr>
	</table>
	<br /><%=tsa_txt_139 %>: <% 
	call timerogminutberegning(logindiffSidste+loginTimerTot)
	%>
	<b><%=thoursTot &":"& left(tminTot, 2) %>&nbsp;<%=tsa_txt_141 %></b>
	
	

<%
end function
%>




