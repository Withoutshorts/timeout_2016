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
dblkostpris, offentlig, intServiceAft, strYear, sTtid, sLtid, visTimerelTid, stopur, intValuta, bopal)

if len(timerthis) <> 0 then
timerthis = timerthis
else
timerthis = 0
end if
    
    call akttyper2009prop(tfaktimvalue)
    
    call valutaKurs(intValuta)
    
    '*** Timer på Ferie og Feriefridage ***'
    '*** Omregn fra hele dage til timer iflg. norm tid ***'
    if muDag = 0 then
    select case tfaktimvalue
    case 11,12,13,14,15,16,17 '** Ferie typer **'
        if cint(timerthis) <> -9001 then
        if cint(tildeliheledage) = 1 then
        
        
             '*** Normeret uge  til beregning (gns. fuld dag) **'
             '*** Brug altid første dag i uge  + 6 arb. dage  **'
             gnstimerprdag = 0
             
             '** Finder mType. ***'
             strSQLmtyp = "SELECT medarbejdertype FROM medarbejdere WHERE mid = " & medid
             'Response.Write 
             oRec4.open strSQLmtyp, oConn, 3
             if not oRec4.EOF then
    	    
    	        call nortimerStandardDag(oRec4("medarbejdertype"))
	            gnstimerprdag = formatnumber(normtimerStDag) 
	        
	        end if
	        oRec4.close
    	    
	        if cdbl(gnstimerprdag) > 0 then
    	    
    	    timercalc = (timerthis/1 * gnstimerprdag/1)
    	    
	        timercalc = replace(timercalc, ".", "")
	        timercalc = replace(timercalc, ",", ".")
    	   
	        timerthis = timercalc
    	    
	        'Response.Write "timerthis " & timerthis & ", timercalc: " & timercalc & ", gnstimerprdag "& gnstimerprdag & ", ntimPer" & ntimPer & "<br>"
	        'Response.flush
    	    
	        else
    	    
	        timerthis = timerthis
    	    
	        end if
    	    
        end if
        end if
    end select
    end if
    
    
   
			
	if cint(timerthis) <> -9001 then '*** <> Slet ***'
			
			
			datothis = ConvertDateYMD(datothis)
			        
			       
			
			'*************************************************************'
			'Tjek om job/akt allerede er registreret for den valgte dato**'
			'*************************************************************'
			
			strSQLfindes = "SELECT timer, Tjobnr, TAktivitetId, timerkom FROM timer WHERE TAktivitetId = "& aktid &" AND Tjobnr = '"& jobnr &"' AND Tmnr = "& medid &" AND Tdato = '"& datothis &"'"
			oRec.Open strSQLfindes, oConn, 3  
			
			'son, man, tir, ons ,tor, fre, lor
			if oRec.EOF then
				
				
				if len(timerthis) > 0 AND cint(timerthis) > -9000 AND timerthis <> 0 _
				OR (len(timerthis) > 0 AND cint(timerthis) > -9000 AND aty_pre <> 0) then 
				
				'** AND timerthis > 0  then -minus er OK HUSK ret i faktrua opr. 15.05.2008 **'
				'** Rettet tilbage 20.05.2008 pga fejl i SponsorCar ***' 
				'** De bruger st. tid og slut tid regsistrering **'
				
				strSQLins = "INSERT INTO timer (TJobnr, TJobnavn, Tmnr, Tmnavn, Tdato, "_
				&" Timer, Timerkom, Tknavn, Tknr, TAktivitetId, TAktivitetNavn, "_
				&" Tfaktim, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
				&" editor, kostpris, offentlig, seraft, sttid, sltid, valuta, kurs, bopal) VALUES"_
				& "(" & jobnr & ", '"& strJobnavn &"', " & medid & ", '" & cstr(strMnavn) & "', '"& datothis &"', "_
				&" "& timerthis &", '"& SQLBless2(kommthis) &"', "_
				&" '" & SQLBless2(strJobknavn) & "', " & strJobknr & ", "_
				&" "& aktid &", '"& SQLBless2(aktnavn) &"', "_
				&" "& tfaktimvalue &", "& strYear &", "& SQLBless(intTimepris) &", "_
				&" '"&year(now)&"/"&month(now)&"/"&day(now)&"', '"& strFastpris  &"', "_
				&" '" & time & "', '"& session("user") &"', "& dblkostpris &", "_
				&" "& offentlig &", "& intServiceAft &", '"&sTtid&"', '"&sLtid&"', "& intValuta &", "& dblKurs &", "& bopal &")"
				
				end if
				
				
				if len(trim(strSQLins)) <> 0 then
				    'Response.Write "Ins: "& strSQLins & "<br>"
				    'Response.flush
					oConn.Execute(strSQLins)
				end if
			
			else
				
					'** Sletter ***'
					'** En preudfyldr aktivtet kan ikke slettes (da den i så fald igen vil være preudfyldt)
					'** Men der kan til gengæld angives 0 ****'
					'Response.Write "<br>timerthis" & timerthis & " pre:"& aty_pre &"<br>"
					
					if (timerthis = 0 AND aty_pre = 0) then
					
					oConn.execute("DELETE FROM timer"_
					&" WHERE Tjobnr = '"& jobnr & "'"_
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
						&" Taar = "& strYear &", TimePris = "& SQLBless(intTimepris) &", tfaktim = "& tfaktimvalue &", "_
						&" seraft = "& intServiceAft &", kostpris = "& dblkostpris &", "_
						&" Timerkom = '"& kommentarthis &"', "_
						&" tastedato = '"&year(now)&"/"&month(now)&"/"&day(now)&"', "_
						&" editor = '"& session("user") &"', offentlig = "& offentlig &", "_
						&" sttid = '"&sTtid&"', sltid = '"&sLtid&"', valuta = "& intValuta &", kurs = "& dblKurs &", bopal = "& bopal &""_
						&" WHERE Tjobnr = '"& jobnr & "'"_
						&" AND Tmnr = "& medid & ""_
						&" AND Tdato = '" & datothis & "' AND TAktivitetId = "& aktid
						
						else
						
						'*************************************************************************************************'
						'** Opdatering af reg. tidspunkt på eksisterende timeregistreringer                            **'
						'** Tidspukter bliver ikke opdateret her, da man ellers vil overskrive tidspunkter              **'
						'** på eksisterende records, ved opdatering af en vilkårlig anden timeregistrering (som timer)  **'
						'*************************************************************************************************'
						
						
						strSQLupd = "UPDATE timer SET"_
						&" Timer = "& timerthis &", tfaktim = "& tfaktimvalue &", Timerkom = '"& kommthis &"', "_
						&" editor = '"& session("user") &"', offentlig = "& offentlig &", bopal = "& bopal &""_
						&" WHERE Tjobnr = '"& jobnr & "'"_
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
function fakfarver(lastfakdato, varTjDato_man, varTjDato_ugedag, erIndlast)




'** Sætter farve på indtastningfelt efter om der er udskrevet en faktura eller ej ***
strWeek = datepart("ww", tjekdag(4), 2, 2)
diff = dateDiff("d", lastfakdato, tjekdag(1))
ugeafsluttet = 0


'Response.Write "x = "& x &" UgadagDato: "& varTjDato_ugedag &" Uge afl: " &ugeNrAfsluttet& "# " & datepart("ww", ugeNrAfsluttet, 2, 2) &" = "& strWeek &" AND "& smilaktiv &" AND "& autogk &" AND "& ugeNrAfsluttet &" autolukvdatodato "& autolukvdatodato &"<br>"
'Response.Write "herder45"
'Response.Write day(now) &" > "& autolukvdatodato  &" AND "& DatePart("yyyy", varTjDato_ugedag) &" = "& year(now) & " AND "& DatePart("m", varTjDato_ugedag) &" < "& month(now) & "<br>"
'Response.Write "aktdata(iRowLoop, 16): " & aktdata(iRowLoop, 16) & "<br>"
 
    select case x
    case 1
    fieldVal = aktdata(iRowLoop, 20)
    case 2
    fieldVal = aktdata(iRowLoop, 21)
    case 3
    fieldVal = aktdata(iRowLoop, 22)
    case 4
    fieldVal = aktdata(iRowLoop, 23)
    case 5
    fieldVal = aktdata(iRowLoop, 24)
    case 6
    fieldVal = aktdata(iRowLoop, 25)
    case 7
    fieldVal = aktdata(iRowLoop, 26)
    end select
    
    
    'Response.Write " aktdata(iRowLoop, 17)  = 0  OR ( " & aktdata(iRowLoop, 17) &" = 1 AND "& fieldVal & " = 1 AND (( " & formatdatetime(aktdata(iRowLoop, 18), 3) & " <=" & timenow & " AND "& aktdata(iRowLoop, 19) &" >= "& timenow &" ) OR " _
	'&" ("& aktdata(iRowLoop, 18) &" <=  " & timenow &" AND "& aktdata(iRowLoop, 19) &" <= "&  aktdata(iRowLoop, 18) &" ) OR " _
	'&" ("& aktdata(iRowLoop, 19) &" >= "& timenow &" AND "& aktdata(iRowLoop, 19) &" <= "& aktdata(iRowLoop, 18) &")))<br><br>" 
    
    
    '*** Er aktivitet lå st pga tidslås ****'
    if len(trim(aktdata(iRowLoop, 18))) <> 0 AND len(trim(aktdata(iRowLoop, 19))) <> 0 then
    
        if cint(ignorertidslas) = 1 OR aktdata(iRowLoop, 17) = 0 OR (aktdata(iRowLoop, 17) = 1 AND fieldVal = 1 AND ((formatdatetime(aktdata(iRowLoop, 18), 3) <= timenow AND formatdatetime(aktdata(iRowLoop, 19), 3) >= timenow ) OR _
	    (formatdatetime(aktdata(iRowLoop, 18), 3) <= timenow AND formatdatetime(aktdata(iRowLoop, 19), 3) <= formatdatetime(aktdata(iRowLoop, 18), 3)) OR _
	    (formatdatetime(aktdata(iRowLoop, 19), 3) >= timenow AND formatdatetime(aktdata(iRowLoop, 19), 3) <= formatdatetime(aktdata(iRowLoop, 18), 3)))) then
        'Åben
        
            maxl = maxl
	        fmbgcol = fmbgcol
	        fmborcl =  fmborcl
    	    
	    else
	    'Lukket
    	
	        'Response.Write "Kaj 2 lukket"
            
            maxl = 0
	        fmbgcol = "#EfF3FF" 
	        fmborcl = "1px #cccccc dashed"
        
            ugeafsluttet = 1
        end if
    
    end if
    
    
    
    if (datepart("ww", ugeNrAfsluttet, 2, 2) = strWeek AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", varTjDato_ugedag) = year(now) AND DatePart("m", varTjDato_ugedag) < month(now)) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", varTjDato_ugedag) < year(now) AND DatePart("m", varTjDato_ugedag) = 12)) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", varTjDato_ugedag) < year(now) AND DatePart("m", varTjDato_ugedag) <> 12) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", varTjDato_ugedag) > 1))) OR cint(lonKorsel_lukketIO) = 1 then


        '*** Uge afsluttet via Smiley Ordning og autogodkend slået til i kontrolpanel **'
        '*** hvis admin level 1 kan timer stadigvæk redigeres indtil der oprettes faktura ***'
	    if level <> 1 then
	    maxl = 0
	    fmbgcol = "#cccccc"
	    fmborcl = "1px #999999 dashed"
	    else
	    maxl = maxl
	    fmbgcol = fmbgcol
	    fmborcl =  fmborcl
	    end if
    	
	    ugeafsluttet = 1
    	
     end if
            
        
            
            
                '** ER registrering godkendt af joabans eller admin ***'
                if len(trim(aktdata(iRowLoop, 16))) <> 0 then
                aktdata(iRowLoop, 16) = aktdata(iRowLoop, 16)
                else
                aktdata(iRowLoop, 16) = 0
                end if
                
                
                
                if cint(aktdata(iRowLoop, 16)) = 1 AND erIndlast = 1 then
                
                    maxl = 0
	                fmbgcol = "#cccccc"
	                fmborcl = "1px yellowgreen dashed"
	                
	                ugeafsluttet = 1
	            
    						        
                end if   
                
                
                
            
            '**** Findes der en faktura ****'
            '**** Overskriver Smiley luk / Godkendte uger / Tidslås ****'
			if len(lastfakdato) <> 0 AND diff <= 0 AND (cint(aktdata(iRowLoop, 16)) = 0 OR erIndlast = 0) then
					        '** Hvis fakuge > den valgte uge ***
					        if datepart("ww", lastfakdato, 2, 2) > strWeek AND datepart("yyyy", lastfakdato) => datepart("yyyy", tjekdag(4)) then
        					
        					    maxl = 0
						        fmbgcol = "#cccccc"
						        fmborcl = "1px #999999 dashed"
						        
						             'Response.Write "Fak her"
						             'Response.Write  aktdata(iRowLoop, 2) &" - "& cint(aktdata(iRowLoop, 16)) &"erIndlast: "& erIndlast &"<br>"
                
        						
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
				                        maxl = 6
				                        fmbgcol = "#ffffff"
				                            '*** Er timer indlæst id DB (specielt ved preval)
				                            if erIndlast = 0 then
				                            fmborcl = "1px #999999 solid"
				                            else
				                            fmborcl = "1px yellowgreen solid"
				                            end if
				                        end if
        							
        							
        							
						        end if
        					
					        end if
			
			else
				
				if ugeafsluttet = 1 then
				maxl = maxl 
				fmbgcol = fmbgcol
				fmborcl = fmborcl
				else	
				maxl = 6
				fmbgcol = "#ffffff"
				           
				            '*** Er timer indlæst id DB (specielt ved preval)
                            if erIndlast = 0 then
                            fmborcl = "1px #999999 solid"
                            else
                            fmborcl = "1px yellowgreen solid"
                            end if
				end if
						
			end if	
			
			
			
          

	
end function



%>





