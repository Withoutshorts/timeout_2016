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
dblkostpris, offentlig, intServiceAft, strYear, sTtid, sLtid, visTimerelTid, stopur, intValuta, bopal, destination)

if len(timerthis) <> 0 then
timerthis = timerthis
else
timerthis = 0
end if
    
    call akttyper2009prop(tfaktimvalue)
    
    call valutaKurs(intValuta)
    
    '*** Timer p� Ferie og Feriefridage ***'
    '*** Omregn fra hele dage til timer iflg. norm tid ***'
    'if muDag = 0 then '** Kun ved f�rste indtastning p� hver medarbejder ***'
    select case tfaktimvalue
    case 11,12,13,14,15,16,17,18,19,20,21,22 '** Ferie typer + syg og barn syg + barsel **'
        if cint(timerthis) <> -9001 then 'tom
       
        
        
           
    	    '** Finder der nomrtimer p� den valgte dag **'
    	    ntimPer = 0
    	    call normtimerPer(medid, datoThis, 0)
    	    
    	    
	        if cdbl(ntimPer) > 0 then
    	  
    	        if cint(tildeliheledage) = 1 then
                    'Response.Write replace(timerthis, ".", ",") & "<br>"
                    if replace(timerthis, ".", ",") <= 1 OR (tfaktimvalue = 12 AND replace(timerthis, ".", ",") <= 10) OR (tfaktimvalue = 15 AND replace(timerthis, ".", ",") <= 35) OR (tfaktimvalue = 22 AND replace(timerthis, ".", ",") <= 365) then 

                            '** Ved tilskrivning af optjent ferie / feriefridage, baseres indtastning p� ugenorm
                            if tfaktimvalue = 12 OR tfaktimvalue = 15 OR tfaktimvalue = 22 then
                            
                            gnstimerprdag = 0
		                    call normtimerPer(medid, datoThis, 6)
    	                    gnstimerprdag = ntimPer/antalDageMtimer '*** Gns. for ugen, hvis timeantal varierer for dag til dag.
    	                    timerthis = replace(replace(timerthis, ".", ",")*gnstimerprdag, ",",".")

                            else

                            gnstimerprdag = 0
		                    call normtimerPer(medid, datoThis, 0)
    	                    gnstimerprdag = ntimPer
    	                    timerthis = replace(replace(timerthis, ".", ",")*gnstimerprdag, ",",".")

                            end if

                    else
                    timerthis = timerthis
                    end if
                else
                timerthis = timerthis
                end if
	            'Response.Write "timerthis " & timerthis & ", timercalc: " & timercalc & ", gnstimerprdag "& gnstimerprdag & ", ntimPer" & ntimPer & "<br>"
	            'Response.flush
	            'timercalc = 0
    	    
	        else
    	    
	        timerthis = 0 'timerthis
    	    
	        end if
    	    
        end if
        'end if
    end select
    'end if '** muDag
    
   
    
   
			
	if cdbl(timerthis) <> -9001 then '*** <> Slet ***'
			
			
			datothis = ConvertDateYMD(datothis)
			        
			        
			
					if visTimerelTid <> 0 AND len(trim(sTtid)) <> 0 then
						
						'Response.write len(trim(sTtid)) &" : "& sTtid &" - "& sLtid &"<br>"
						'Response.flush
						
						if len(trim(sTtid)) = 1 then '** slet
						
						timerthis = 0
						
						else
						idag = day(now)&"/"&month(now)&"/"&year(now)
						totalmin = datediff("n", idag &" "& sTtid, idag &" "& sLtid)
						call timerogminutberegning(totalmin)
						timerthis = thoursTot&"."&tminProcent 'tminTot
						
						 
					    if timerthis < 0 then '** Hen over kl 24:00 **'
						timerthis = 24 + (replace(timerthis,".",","))
						timerthis = replace(timerthis,",",".")
						end if
						
						
						sTtid = sTtid &":00"
						sLtid = sLtid &":00"
						end if
					
					else
					
					sTtid = ""
					sLtid = ""
					
					end if
					
			
			
			strSQLfindes = "SELECT timer, Tjobnr, TAktivitetId, timerkom FROM timer WHERE TAktivitetId = "& aktid &" AND Tjobnr = "& jobnr &" AND Tmnr = "& medid &" AND Tdato = '"& datothis &"'"
			oRec.Open strSQLfindes, oConn, 3  
			
			'son, man, tir, ons ,tor, fre, lor
			if oRec.EOF then
				
                'Response.Write timerthis & "<br>"
				
				if len(timerthis) > 0 AND cdbl(timerthis) > -9000 AND timerthis <> 0 _
				OR (len(timerthis) > 0 AND cdbl(timerthis) > -9000 AND aty_pre <> 0) then 
				
				'** AND timerthis > 0  then -minus er OK HUSK ret i faktrua opr. 15.05.2008 **'
				'** Rettet tilbage 20.05.2008 pga fejl i SponsorCar ***' 
				'** De bruger st. tid og slut tid regsistrering **'
				
				strSQLins = "INSERT INTO timer (TJobnr, TJobnavn, Tmnr, Tmnavn, Tdato, "_
				&" Timer, Timerkom, Tknavn, Tknr, TAktivitetId, TAktivitetNavn, "_
				&" Tfaktim, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
				&" editor, kostpris, offentlig, seraft, sttid, sltid, valuta, kurs, bopal, destination) VALUES"_
				& "(" & jobnr & ", '"& strJobnavn &"', " & medid & ", '" & cstr(strMnavn) & "', '"& datothis &"', "_
				&" "& timerthis &", '"& SQLBless2(kommthis) &"', "_
				&" '" & SQLBless2(strJobknavn) & "', " & strJobknr & ", "_
				&" "& aktid &", '"& SQLBless2(aktnavn) &"', "_
				&" "& tfaktimvalue &", "& strYear &", "& SQLBless(intTimepris) &", "_
				&" '"&year(now)&"/"&month(now)&"/"&day(now)&"', '"& strFastpris  &"', "_
				&" '" & time & "', '"& session("user") &"', "& dblkostpris &", "_
				&" "& offentlig &", "& intServiceAft &", '"&sTtid&"', '"&sLtid&"', "_
				&" "& intValuta &", "& dblKurs &", "& bopal &", '"& destination &"')"
				
				end if
				
				
				if len(trim(strSQLins)) <> 0 then
				    'Response.Write "Ins: "& strSQLins & "<br>"
				    'Response.flush
					oConn.Execute(strSQLins)
				end if
			
			else
				
					'** Sletter ***'
					'** En preudfyldr aktivtet kan ikke slettes (da den i s� fald igen vil v�re preudfyldt)
					'** Men der kan til geng�ld angives 0 ****'
					'Response.Write "<br>timerthis" & timerthis & " pre:"& aty_pre &"<br>"
					
					if (timerthis = 0 AND aty_pre = 0) then
					
					strSQLdel = "DELETE FROM timer"_
					&" WHERE Tjobnr = "& jobnr & ""_
					&" AND Tmnr = "& medid & ""_
					&" AND Tdato = '" & datothis & "' AND TAktivitetId = "& aktid & ""
					
					oConn.execute(strSQLdel)
					
					'Response.Write strSQLdel & "<br>"
					
					
					else
					
						'** Opdaterer ****
						if (cdbl(SQLBless(oRec("timer"))) <> cdbl(timerthis)) OR stopur = "1" then
						
						        if stopur = "1" then 
						        '** tilf�jer timer til eksisterende timer istedet for at overskrive **'
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
						&" sttid = '"&sTtid&"', sltid = '"&sLtid&"', valuta = "& intValuta &", kurs = "& dblKurs &", bopal = "& bopal &", "_
						&" destination = '"& destination &"'"_
						&" WHERE Tjobnr = "& jobnr & ""_
						&" AND Tmnr = "& medid & ""_
						&" AND Tdato = '" & datothis & "' AND TAktivitetId = "& aktid
						
						'Response.write "Upd A: "& strSQLupd & "<br>"
						
						else
						
						'*************************************************************************************************'
						'** Opdatering af reg. tidspunkt p� eksisterende timeregistreringer                            **'
						'** Tidspukter bliver ikke opdateret her, da man ellers vil overskrive tidspunkter              **'
						'** p� eksisterende records, ved opdatering af en vilk�rlig anden timeregistrering (som timer)  **'
						'*************************************************************************************************'
						
						
						strSQLupd = "UPDATE timer SET"_
						&" Timer = "& timerthis &", tfaktim = "& tfaktimvalue &", Timerkom = '"& kommthis &"', "_
						&" editor = '"& session("user") &"', offentlig = "& offentlig &", bopal = "& bopal &", destination = '"& destination &"'"_
						&" WHERE Tjobnr = "& jobnr & ""_
						&" AND Tmnr = "& medid & ""_
						&" AND Tdato = '" & datothis & "' AND TAktivitetId = "& aktid 
						
						'Response.write "Upd B: "& strSQLupd & "<br>"
						
						end if
						
						
					    'Response.end
						oConn.Execute(strSQLupd)
						
					end if
			end if
			oRec.close
			
	end if '** timer -9001 (-1) ***'


end function




public maxl, fmbgcol, fmborcl 
function fakfarver(lastfakdato, varTjDato_man, varTjDato_ugedag, erIndlast, medid, akttype, tidslaasMan, tidslaasTir, tidslaasOns, tidslaasTor, tidslaasFre, tidslaasLor, tidslaasSon, job_tidslass, job_tidslassSt, job_tidslassSl, ignorertidslas, godkendtstatus)


'Response.Write "lastfakdato" & lastfakdato
'Response.end

timenow = formatdatetime(now, 3)

'call erugeAfslutte(year(tjekdag(4)), tjekdag(7), usemrn)
'** S�tter farve p� indtastningfelt efter om der er udskrevet en faktura eller ej ***
strWeek = datepart("ww", tjekdag(4), 2, 2)
strAar = datepart("yyyy", tjekdag(4), 2, 2)
diff = dateDiff("d", lastfakdato, tjekdag(1))
ugeafsluttet = 0




call erugeAfslutte(strAar, strWeek, usemrn)

'Response.Write "lastfakdato" & lastfakdato & "strAar: "& strAar & " strWeek: "& strWeek & " usemrn:" & usemrn
'Response.end

'** Quickfix, ugeNrAfsluttet bliver slet ikka kaldt og derfor ikke sat
'** og bliver af VB sat = 52, hvorfor man ikke kan registrere timer 
'if ugeNrAfsluttet <> "" then
'ugeNrAfsluttet = ugeNrAfsluttet
'else
'ugeNrAfsluttet = dateadd("d", -60, now)
'end if

'Response.Write "varTjDato_ugedag" & varTjDato_ugedag
'Response.Write " #"& ugeNrAfsluttet &"#"& datepart("ww", ugeNrAfsluttet, 2, 2) & "<br>"

'Response.Write "x = "& x &" UgadagDato: "& varTjDato_ugedag &" Uge afl: " &ugeNrAfsluttet& "# " & datepart("ww", ugeNrAfsluttet, 2, 2) &" = "& strWeek &" AND "& smilaktiv &" AND "& autogk &" AND "& ugeNrAfsluttet &" autolukvdatodato "& autolukvdatodato &"<br>"
'Response.Write "herder45"
'Response.Write day(now) &" > "& autolukvdatodato  &" AND "& DatePart("yyyy", varTjDato_ugedag) &" = "& year(now) & " AND "& DatePart("m", varTjDato_ugedag) &" < "& month(now) & "<br>"
'Response.Write "aktdata(iRowLoop, 16): " & aktdata(iRowLoop, 16) & "<br>"
 


    select case x
    case 1
    fieldVal = tidslaasMan 'aktdata(iRowLoop, 20)
    case 2
    fieldVal = tidslaasTir 'aktdata(iRowLoop, 21)
    case 3
    fieldVal = tidslaasOns 'aktdata(iRowLoop, 22)
    case 4
    fieldVal = tidslaasTor 'aktdata(iRowLoop, 23)
    case 5
    fieldVal = tidslaasFre 'aktdata(iRowLoop, 24)
    case 6
    fieldVal = tidslaasLor 'aktdata(iRowLoop, 25)
    case 7
    fieldVal = tidslaasSon 'aktdata(iRowLoop, 26)
    end select
    
    
    'Response.Write " job_tidslass  = 0  OR ( " & job_tidslass &" = 1 AND "& fieldVal & " = 1 AND (( " & formatdatetime(job_tidslassSt, 3) & " <=" & timenow & " AND "& job_tidslassSl &" >= "& timenow &" ) OR " _
	'&" ("& job_tidslassSt &" <=  " & timenow &" AND "& job_tidslassSl &" <= "&  job_tidslassSt &" ) OR " _
	'&" ("& job_tidslassSl &" >= "& timenow &" AND "& job_tidslassSl &" <= "& job_tidslassSt &")))<br><br>" 
    

    


    '*** Er aktivitet l� st pga tidsl�s ****'
    if len(trim(job_tidslassSt)) <> 0 AND len(trim(job_tidslassSl)) <> 0 then
    
        if cint(ignorertidslas) = 1 OR job_tidslass = 0 OR (job_tidslass = 1 AND fieldVal = 1 AND ((formatdatetime(job_tidslassSt, 3) <= timenow AND formatdatetime(job_tidslassSl, 3) >= timenow ) OR _
	    (formatdatetime(job_tidslassSt, 3) <= timenow AND formatdatetime(job_tidslassSl, 3) <= formatdatetime(job_tidslassSt, 3)) OR _
	    (formatdatetime(job_tidslassSl, 3) >= timenow AND formatdatetime(job_tidslassSl, 3) <= formatdatetime(job_tidslassSt, 3)))) then
        '�ben
        
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
    
    '**** Er periode lukket via l�nk�rsel **''
    call lonKorsel_lukketPer(varTjDato_ugedag)

    'Response.write "lonKorsel_lukketIO = " & lonKorsel_lukketIO 

    if (datepart("ww", ugeNrAfsluttet, 2, 2) = strWeek AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", varTjDato_ugedag, 2, 2) = year(now) AND DatePart("m", varTjDato_ugedag, 2, 2) < month(now)) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", varTjDato_ugedag, 2, 2) < year(now) AND DatePart("m", varTjDato_ugedag, 2, 2) = 12)) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", varTjDato_ugedag, 2, 2) < year(now) AND DatePart("m", varTjDato_ugedag, 2, 2) <> 12) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", varTjDato_ugedag, 2, 2) > 1))) OR _
    cint(lonKorsel_lukketIO) = 1 then

        
        '*** L�nperiode afsluttet ***'
        '*** Uge afsluttet via Smiley Ordning og autogodkend sl�et til i kontrolpanel **'
        '*** hvis admin level 1 kan timer stadigv�k redigeres indtil der oprettes faktura ***'
	    if level <> 1 then
	        maxl = 0
	        fmbgcol = "#cccccc"
	        fmborcl = "1px #999999 dashed"
	    else
	        maxl = maxl
	        fmbgcol = fmbgcol
	        fmborcl =  "1px #999999 dashed" 'fmborcl
	    end if
    	
	    ugeafsluttet = 1
    	
     end if
            
                

                 '*** Ferie og Sygdom skal ikke kunne indtastes p� dage uden normtid ***'
                 call normtimerPer(medid, varTjDato_ugedag, 0)

                 'Response.Write "nTimPer " & nTimPer & " akttype:" &akttype & "<br>"

                 if (nTimPer = 0) AND (akttype > 10 AND akttype < 22) then
                 

                maxl = 0
	            fmbgcol = "#EfF3FF" 
	            fmborcl = "1px #cccccc dashed"
        
                ugeafsluttet = 1
                 
                 end if  




            
            
                '** ER registrering godkendt af joabans eller admin ***'
                if len(trim(godkendtstatus)) <> 0 then
                godkendtstatus = godkendtstatus
                else
                godkendtstatus = 0
                end if
                
                
                
                if cint(godkendtstatus) = 1 AND erIndlast = 1 then
                
                    maxl = 0
	                fmbgcol = "#cccccc"
	                fmborcl = "1px yellowgreen dashed"
	                
	                ugeafsluttet = 1
	            
    						        
                end if   
                
                
                
            

            
            '**** Findes der en faktura ****'
            '**** Overskriver Smiley luk / Godkendte uger / Tidsl�s ****'
			if len(lastfakdato) <> 0 AND diff <= 0 AND (cint(godkendtstatus) = 0 OR erIndlast = 0) then
					        '** Hvis fakuge > den valgte uge ***
					        if datepart("ww", lastfakdato, 2, 2) > strWeek AND datepart("yyyy", lastfakdato) => datepart("yyyy", tjekdag(4)) then
        					
        					    maxl = 0
						        fmbgcol = "#cccccc"
						        fmborcl = "1px #999999 dashed"
						        
						             'Response.Write "Fak her"
						             'Response.Write  aktdata(iRowLoop, 2) &" - "& cint(aktdata(iRowLoop, 16)) &"erIndlast: "& erIndlast &"<br>"
                
        						
        				    else
        					
        					 
        					
						        '** Hvis fakuge = den valgte uge ***
						        '** Tidspunkt s�ttes altid til 23:59:59 p� fakturaer, fra d. 8/9-2004
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
				                            '*** Er timer indl�st id DB (specielt ved preval)
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
				           
				            '*** Er timer indl�st id DB (specielt ved preval)
                            if erIndlast = 0 then
                            fmborcl = "1px #999999 solid"
                            else
                            fmborcl = "1px yellowgreen solid"
                            end if
				end if
						
			end if	
			
			
			 
	
end function


public jLuft 
sub jobbeskrivelse_stdata

select case lto
case "dencker", "xintranet - local"
jbeskWzb = "visible"
jbeskDsp = ""
jLuft = "<br>&nbsp;"
case else
jbeskWzb = "hidden"
jbeskDsp = "none"
jLuft = ""
end select

%>

  <div id="div_det_<%=aktdata(iRowLoop, 4)%>"" style="position:relative; background-color:#ffffff; padding:0px 3px 3px 3px; width:745px; border:3px #CCCCCC solid; visibility:<%=jbeskWzb%>; display:<%=jbeskDsp%>; z-index:2000">
   
            <table cellspacing="0" cellpadding="0" border="0" width="100%" bgcolor="#8caae6">
            

            <tr>
                  <td colspan=5 style="background-color:#D6dff5; background-image:url('../ill/stripe_10_graa.png'); height:4px;">
                    <img src="../ill/blank.gif" width=1 height=4 /></td>  </tr>

                   

                                                 <tr bgcolor="#FFFFFF">
                                                 <td valign=top class=lille style="width:250px; padding-top:3px;">

                                                

                                                            <b>Timefordeling seneste 4 md.</b>
                                                            <table cellpadding=0 cellspacing=0 border=0>
                                                            <tr>
                                                                <%
                                                                'call akttyper2009(2)
            
                                                                for dt = 0 TO 3 
                                                                 if dt <> 0 then
                                                                 dtnowHigh = dateadd("m", 1, dtnowHigh)
                                                                 else
                                                                 dtnowHigh = dateadd("m", -3, now)
                                                                 end if

            

                                                                %>
                                                                <td align=center style="font-size:7px; width:20px;"><%=left(monthname(month(dtnowHigh)), 3) %></td>
                                                                <%next %>
                                                             </tr>

                                                             <tr bgcolor="#FFFFFF">
                                                                <%for dt = 0 TO 3
                                                                 if dt <> 0 then
                                                                 dtnowHigh = dateadd("m", 1, dtnowHigh)
                                                                 else
                                                                 dtnowHigh = dateadd("m", -3, now)
                                                                 end if
             
             
                                                                 select case month(dtnowHigh)
                                                                 case 1,3,5,7,8,10,12
                                                                    eday = 31
                                                                    case 2
                                                                        select case right(year(dtnowHigh), 2)
                                                                        case "00", "04", "08", "12", "16", "20", "24", "28", "30", "34", "38", "42", "46"
                                                                        eday = 29
                                                                        case else
                                                                        eday = 28
                                                                        end select
                                                                 case else
                                                                    eday = 30
                                                                 end select

                                                                 dtnowHighSQL = year(dtnowHigh) &"/"& month(dtnowHigh) &"/"& eday
                                                                 dtnowLowSQL = year(dtnowHigh) &"/"& month(dtnowHigh) &"/1"

                                                                   hgt = 0
                                                                   timerThis = 0

                                                                 strSQLjl = "SELECT sum(timer) AS timer FROM timer WHERE tjobnr = "& aktdata(iRowLoop, 31) &" AND tdato BETWEEN '"& dtnowLowSQL &"' AND '"& dtnowHighSQL &"'"_
                                                                 &" AND ("& aty_sql_realhours &") GROUP BY tjobnr"

                                                                 'Response.Write strSQLjl
                                                                 'Response.flush

                                                                 oRec2.open strSQLjl, oConn, 3
                                                                 if not oRec2.EOF then

                                                                 hgt = formatnumber(oRec2("timer"), 0)
                                                                 timerThis = oRec2("timer")

                                                                 end if
                                                                 oRec2.close

                                                                 if hgt > 100 then
                                                                 hgt = 20
                                                                 else
                                                                 hgt = hgt/5 
                                                                 end if
                                                                %>
                                                                <td class=lille align=right height=20 valign=bottom style="border-right:1px #CCCCCC solid; width:20px;">
                                                                <%if hgt <> 0 then 
                                                                bdThis = 1
                                                                timerThis = formatnumber(timerThis,2)
                                                                else
                                                                timerThis = ""
                                                                bdThis = 0
            
                                                                end if 
            
                                                                bgThis = "#D6dff5"

                                                                if hgt = 0 then
                                                                bgThis = "#ffffff"
                                                                end if

           
          
                                                                %>
                                                                <div style="height:<%=hgt%>px; width:20px; border-top:<%=bdThis%>px #CCCCCC solid; padding:0px; background-color:<%=bgThis%>;"><img src="../ill/blank.gif" width="20" height="<%=hgt%>" border="0" alt="<%=timerThis %>" /></div>
            
                                                                </td>
                                                                <%next %>
                                                             </tr>
	    
                                                            </table>


                                                            <%
                                                            call timer_fordeling_medarb_typer(aktdata(iRowLoop, 31), 4)
                                                            %>
                                                          



                                                <br />
                                                <b>Stamdata:</b><br />
                                                Start & slutdato: <%=formatdatetime(aktdata(iRowLoop, 45), 2) %> - <%=formatdatetime(aktdata(iRowLoop, 46), 2) %>	


                                                <br />

                                                

                                               <%call meStamdata(aktdata(iRowLoop, 43)) %>
                                               <%=tsa_txt_007 %>: <%=meTxt %>

                                               <br />
                                               <%call meStamdata(aktdata(iRowLoop, 44)) %>
                                               <%=tsa_txt_010 %>: <%=meTxt %>

                                                  <br />
                                               <%call meStamdata(aktdata(iRowLoop, 50)) %>
                                               <%=tsa_txt_011 %>: <%=meTxt %>

                                               <%if len(trim(aktdata(iRowLoop, 27))) <> 0 AND aktdata(iRowLoop, 27) <> "0" then %>
                                               <br />Kontaktperson hos kunde:<br />
                                               <%
                                               kkpers = ""
                                               strSQLkk = "SELECT navn, email FROM kontaktpers WHERE id = "& aktdata(iRowLoop, 27) 
                                               oRec6.open strSQLkk, oCOnn, 3
                                               if not oRec6.EOF then

                                               kkpers = oRec6("navn") & ", "& oRec6("email")

                                               end if
                                               oRec6.close

                                               Response.write kkpers
                                               end if
                                               %>


                                               

      
		
                                                <%if level <= 2 OR level = 6 then %>	

		                                                <%select case aktdata(iRowLoop, 47) '** Fastpris
		                                                case 1
		                                                jobtype = tsa_txt_013 '"Fastpris"
				                                                'if oRec("usejoborakt_tp") <> 0 then
				                                                'jobtype = jobtype & " ("& lcase(tsa_txt_125) &")"
				                                                'end if
		                                                case else
		                                                jobtype = tsa_txt_014 '"Lbn. timer"
		                                                end select
					
		
		                                        %>
		                                        <br /><%=tsa_txt_015 %>: <%=jobtype%>		
		                                         <%end if %>		
		 
         
                                                 <%
             
			

				

						
						                                        '*** Aftale ***
						                                        if aktdata(iRowLoop, 51) <> 0 then

                                                                %>
                                                                <br /><%=tsa_txt_018 %>:

						                                        <%
                                                                strSQLaft = "SELECT navn, aftalenr, status FROM serviceaft WHERE id = " & aktdata(iRowLoop, 51)
						                                        oRec5.open strSQLaft, oConn, 3
                                                                if not oRec5.EOF then
							
							                                        %>
                            
							                                        <%=oRec5("navn")%> (<%=oRec5("aftalenr")%>)
							
							                                        <%
							                                        select case oRec5("status")
							                                        case 1
							                                        %>
							                                         - <i><%=tsa_txt_019 %></i>
							                                        <%case else
							                                        %> 
							                                         - <i> <%=tsa_txt_020 %></i>
							                                        <%end select
							
						
						                                        end if
						                                        oRec5.close

                                                                end if
                        
                                                                %>
                  
				
				                                        <%if len(trim(rekvnr)) <> 0 then %>
				                                        <br /><%=tsa_txt_269 %>: <%=aktdata(iRowLoop, 49) %>
				                                        <%end if %>

                                              </td>
                                               <td style="border-right:1px #cccccc solid;">&nbsp;</td>
                                               <td valign=top class=lille style="padding:5px 5px 5px 10px; width:250px;">

                                                 <%if level <= 2 OR level = 6 then %>	
     
                                                <!--Budget -->
                                                            <b><%=tsa_txt_016 %></b>
                    
				                                            <br />Timer: <%=formatnumber((j_timerforkalk), 2)%> t.
                                                            <br />Bruttooms�tning: <%=formatnumber(aktdata(iRowLoop, 48), 2) %> DKK

                                                            <br />
                                                            <br /><b><%=tsa_txt_017 %>:</b><br />
					                                        <%
					                                        timerbrugtthis = 0
                                                            internKost = 0
                                                            oms = 0

                                                            if j_timerforkalk <> 0 then
                                                            j_timerforkalk = j_timerforkalk
                                                            else
                                                            j_timerforkalk = 1
                                                            end if 
							
							                                        '* finder antal real timer p� job ***
							                                        strSQL2 = "SELECT timer AS sumtimer, (kostpris * timer * kurs / 100) AS kostpris, timepris FROM timer WHERE tjobnr = "& aktdata(iRowLoop, 31) &" AND ("& aty_sql_realhours &") ORDER BY timer"
							                                        oRec2.open strSQL2, oConn, 3
							                                        while not oRec2.EOF 

							                                        timerbrugtthis = timerbrugtthis + oRec2("sumtimer")
                            
                                                                    if aktdata(iRowLoop, 47) = 1 then '** Fastpris
                                                                    oms = oms + (oRec2("sumtimer") * (aktdata(iRowLoop, 48)/j_timerforkalk))
                                                                    else
                                                                    oms = oms + (oRec2("sumtimer") * oRec2("timepris"))
                                                                    end if

                                                                    internKost = internKost + oRec2("kostpris")
							                                        oRec2.movenext
                                                                    wend 
							                                        oRec2.close
							
							                                        if len(timerbrugtthis) > 0 then
							                                        timerbrugtthis = timerbrugtthis 
							                                        else
							                                        timerbrugtthis = 0
							                                        end if

                                                                    if len(internKost) > 0 then
							                                        internKost = internKost 
							                                        else
							                                        internKost = 0
							                                        end if
							
							                                        %>
					
					                                        Timer: <%=formatnumber(timerbrugtthis, 2)%> t. 
                                                            <br />
                    
                                                            Oms�tning: <%=formatnumber(oms) %> DKK<br />
                                                            Intern kostpris: <%=formatnumber(internKost) %> DKK
					

                    
                                                            <%end if %>


                                                            <!-- stade --->

                                                            <br />
                                                            <span style="color:darkred;">
                                                            <%if cint(stade_tim_proc) = 0 then %>
                                                            <b>Restestimat: </b><%=restestimat &" timer"%>
                                                            <%else%>
                                                            <b>Afsluttet: </b><%=restestimat &" %"%>
                                                            <%end if %>
                                                            </span> (angivet) 
		
        
                                                </td>
                                                 <td style="border-right:1px #cccccc solid;"><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
                                                <td valign=top style="padding:5px 5px 5px 10px; width:245px;" class=lille>


                                                           
                

                	                                        <%if editok = 1 AND fromsdsk = 0 then  %>

                   
                                                            <a href="jobs.asp?menu=job&func=red&id=<%=aktdata(iRowLoop, 4)%>&int=1&rdir=treg" class=rmenu target="_top">Rediger job >></a> 
				                                        <br />
				                                        <a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&func=opret&jobid=<%=aktdata(iRowLoop, 4)%>&id=0&jobnavn=<%=aktdata(iRowLoop, 5)%>&fb=1&rdir=treg2&nomenu=1', '', 'width=850,height=600,resizable=yes,scrollbars=yes')" class=rmenu>Opret ny aktivitet >></a>
			
				                                        <br />
				                                        <a href="job_print.asp?menu=job&id=<%=aktdata(iRowLoop, 4)%>&kid=<%=aktdata(iRowLoop, 42)%>" target="_blank" class='rmenu'>Print / PDF >></a>
				
				                                        <br />
				                                        <a href="job_kopier.asp?func=kopier&id=<%=aktdata(iRowLoop, 4)%>&fm_kunde=0&filt=aaben&showaspopup=y&rdir=timereg&usemrn=<%=session("mid")%>&showakt=1" class='rmenu'>Kopier job >></a><br />
                                                        <a href="#" onclick="Javascript:window.open('../timereg/milepale.asp?menu=job&func=opr&jid=<%=aktdata(iRowLoop, 4)%>', '', 'width=650,height=400,resizable=yes,scrollbars=yes')" class=rmenu>Opret milep�l >></a><br />
                                                        <a href="#" onclick="Javascript:window.open('../timereg/milepale.asp?menu=job&func=opr&jid=<%=aktdata(iRowLoop, 4)%>&type=1', '', 'width=650,height=400,resizable=yes,scrollbars=yes')" class=rmenu>Opret bet. termin >></a>
				
				                                        <%end if %>





                
				
				                                        <%'if (level < 7) then %>
				
                                                        <br /><br />		
				                                        <a href="materialer_indtast.asp?id=<%=aktdata(iRowLoop, 4)%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aktdata(iRowLoop, 51)%>" target="_blank" class=rmenu><img src="../ill/ikon_materiale.png" border="0" /><%=tsa_txt_021 %> >> </a>
				                                        <br />
				                                        <%
                                                         '*** Dine **'
				                                         antalmatreg = 0
				                                         fordeling = 0
				                                         strSQLmr = "SELECT COUNT(id) AS fordeling, SUM(matantal) AS matantal FROM materiale_forbrug WHERE jobid = "& aktdata(iRowLoop, 4) & " AND usrid = "& usemrn & " GROUP BY jobid, usrid"
				                                         'Response.Write strSQLmr 
				                                         'Response.flush
				 
				 
				                                         oRec4.open strSQLmr, oConn, 3
				                                         if not oRec4.EOF then
				                                         fordeling = oRec4("fordeling")
				                                         antalmatreg = oRec4("matantal")
				                                         end if
				                                         oRec4.close
				                                         %>


                                                          <%'*** total **'
				                                         antalmatregTot = 0
				                                         fordelingTot = 0
				                                         strSQLmr = "SELECT COUNT(id) AS fordeling, SUM(matantal) AS matantal FROM materiale_forbrug WHERE jobid = "& aktdata(iRowLoop, 4) & " GROUP BY jobid, usrid"
				                                         'Response.Write strSQLmr 
				                                         'Response.flush
				 
				 
				                                         oRec4.open strSQLmr, oConn, 3
				                                         if not oRec4.EOF then
				                                         fordelingTot = oRec4("fordeling")
				                                         antalmatregTot = oRec4("matantal")
				                                         end if
				                                         oRec4.close
				                                         %>

				
				                                        <%=tsa_txt_326 %> <a href="materiale_stat.asp?id=0&hidemenu=1&FM_job=<%=aktdata(iRowLoop, 4)%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aktdata(iRowLoop, 51)%>&FM_visprjob_ell_sum=0" target="_blank" class=rmenu><%=antalmatregTot%> (<%=antalmatreg%>) </a><br /> <%=tsa_txt_327 %> <b><%=fordelingTot%> (<%=fordeling%>)</b> <%=tsa_txt_328 %>.
				
				                                        <%'end if %>
				
				
			                                             <br /><br />
			                                             <!--<input type="checkbox" value="<%=aktdata(iRowLoop, 4)%>" name="FM_lukjob" id="FM_lukjob" class="FM_lukjob" />--><%=tsa_txt_375 %><br />
                                                         <input type="button" value=" Send >> " id="FM_lukjob_<%=aktdata(iRowLoop, 4)%>" class="FM_lukjob" style="font-size:9px;" />
                                                         
                                                       

            
                                                     &nbsp;

                                                        </td>
                                                        </tr>

                 

                                                         <!--- job besk --->

                                                         <%if (len(trim(aktdata(iRowLoop, 41))) <> 0 OR len(trim(aktdata(iRowLoop, 40))) <> 0) then 
	 
	                                                     select case lto
	                                                     case "dencker", "intranet - local", "jttek"
	                                                     jobbeskdivSHOW = "visible"
	                                                     jobbeskdivDSP = ""
	                                                     case else
	                                                     jobbeskdivSHOW = "hidden"
	                                                     jobbeskdivDSP = "none"
	                                                     end select
	 
	                                                     %>
	                                                                 <!---job besk -->
	                    
                                                                    <tr bgcolor="#FFFFFF"><td colspan=5 style="padding-left:12px;">
                                                                    <br />
	                        
                          
                                                                        <a href="#" class="showjobbesk" id="showjobbesk_<%=aktdata(iRowLoop, 4) %>">(+ Se <%=lcase(tsa_txt_029) %>: <%=len(trim(aktdata(iRowLoop, 41))) &" "& tsa_txt_371 &"." %>)</a>
                             
	           
	          
				                                                                <div id="jobbeskdiv_<%=aktdata(iRowLoop, 4)%>" style="overflow:auto; height:150px; width:710px; padding:10px 0px 10px 5px; visibility:<%=jobbeskdivSHOW%>; display:<%=jobbeskdivDSP%>; border:1px #cccccc solid; background-color:snow;">
				                                                                <table border=0 cellspacing=0 cellpadding=5 width=100%>
                                                                                <tr><td valign=top style="width:350px;"><b>Job beskrivelse:</b><br />
				                                                                <%=trim(aktdata(iRowLoop, 41))%>
				                                                                </td>
				
				                                                                <%if len(trim(aktdata(iRowLoop, 40))) <> 0 then %>
				                                                                <td valign=top style="width:250px;">
				                                                                <b>Intern note:</b><br />
				                                                                <%=trim(aktdata(iRowLoop, 40)) %></td>
				                                                                <%end if %>
				
				                                                                </tr></table>
				                                                                </div>
	
	                                                                    <br />&nbsp;
	                                                             

                                                                 
	                                                                    </td></tr>
                    
                           





	                                                    <%end if 
                                                        '***** SLUT Job stamdata + beskrivesle ****'


                                                        %>

                                                      
                                                        </table>
                                                        </div>

                                                        <%

end sub



sub etsSub

                                                    '*** Sletter indtastninger p� den valgte uge p� et valge job,
                            		                '*** s� man ikke risikerer at have indtastninger st�ende i et
                            		                '*** tidsrum der ved redigering ikke l�ngere er i det indstastede tidsinterval ***'
                            		                '*** Kun nat, dat etc og E1 typen. IKKE p� grundregistreringen ***'
                                                    thisJobid = 0
                                                    strSQLjobid = "SELECT jobnr FROM job WHERE id = "& jobids(0)
                                                    'Response.Write strSQLjobid
                                                    'Response.flush
                                                    
                                                    oRec.open strSQLjobid, oConn, 3
                                                    if not oRec.EOF then
                                                    thisJobnr = oRec("jobnr")
                                                    end if
                                                    oRec.close
                                                    
                                                    vlgtDato = strDag &"/"& strMrd &"/"& strAar
                                                    select case weekday(vlgtDato,2)
                                                    case 1
                                                    stDeldato = dateadd("d",0, vlgtDato)
                                                    slDeldato = dateadd("d",6, vlgtDato)
                                                    case 2
                                                    stDeldato = dateadd("d",-1, vlgtDato)
                                                    slDeldato = dateadd("d",5, vlgtDato)
                                                    case 3
                                                    stDeldato = dateadd("d",-2, vlgtDato)
                                                    slDeldato = dateadd("d",4, vlgtDato)
                                                    case 4
                                                    stDeldato = dateadd("d",-3, vlgtDato)
                                                    slDeldato = dateadd("d",3, vlgtDato)
                                                    case 5
                                                    stDeldato = dateadd("d",-4, vlgtDato)
                                                    slDeldato = dateadd("d",2, vlgtDato)
                                                    case 6
                                                    stDeldato = dateadd("d",-5, vlgtDato)
                                                    slDeldato = dateadd("d",1, vlgtDato)
                                                    case 7
                                                    stDeldato = dateadd("d",-6, vlgtDato)
                                                    slDeldato = dateadd("d",0, vlgtDato)
                                                    end select
                                                    
                                                    stDeldato = year(stDeldato) &"/"& month(stDeldato) &"/"& day(stDeldato)
                                                    slDeldato = year(slDeldato) &"/"& month(slDeldato) &"/"& day(slDeldato)
                                                    
                                                    for m = 0 to UBOUND(tildelselmedarb)
	
	                                                'Response.Write tildelselmedarb(m) & "<br>"
                                                	
	                                                if cint(multitildel) = 1 then
	                                                medarbejderid = tildelselmedarb(m) 
	                                                else
	                                                medarbejderid = medarbejderid
	                                                end if
	                                                
                                                    strSQLdelweek = "DELETE FROM timer WHERE tjobnr = "& thisJobnr &" AND "_
                                                    &" tdato BETWEEN '"& stDeldato &"' AND '"& slDeldato &"' AND tmnr = " & medarbejderid & ""_
                                                    &" AND (tfaktim = 50 OR tfaktim = 51 OR tfaktim = 52 OR tfaktim = 53 OR tfaktim = 54 OR tfaktim = 55 OR tfaktim = 90 OR tfaktim = 91)"  
                            				
                            				        'Response.Write strSQLdelweek & "<br><br>"
                            				        oConn.execute(strSQLdelweek)
                            				        
                            				        next
                            				        'Response.end   
			                       
			                       
			                       
			                       
			                       
			                       
			                       
			                        
			                    
	                             'newArrSizeTim = tTimertildelt
	                             icc = 0
	                             cta = 0
			                     jarrforl = ""
			                     for j = 0 to UBOUND(jobids) 'antal akt linier
			             
			                        
			                        if j = 0 then
			                        ystart = 0
			                        else
			                        ystart = (j * 7) 
			                        end if
                        			
			                        yslut = (ystart + 6)
			                     
			                     
			                     'gnlob = 1
			                     'oprArrSize = ubound(tTimertildelt)
			                     lasty2 = 0
			                     for y = ystart to yslut    
			                         
			                         
			                                        '*** Nulstiller alle DAG / NAT / E1 typer s� der ikke kan indtastes timer p� disse typer
                            			            '*** Selvom de evt. er �bne og der tastes timer ind i felterne ***'
                            			            'if t = 9999 then
                            			            aktType = 0
                            			            strSQLtype = "SELECT fakturerbar "_
                                                    &" FROM aktiviteter WHERE "_
                                                    &" id = "& aktids(j) 
                                                    
                                                    'Response.Write strSQLtype & "<br>"
                            		                
                            		                oRec5.open strSQLtype, oConn, 3
                            		                if not oRec5.EOF then
                            		                
                            		                aktType = oRec5("fakturerbar")
                            		                
                            		                end if
                            		                oRec5.close
                            		                
                            		                if aktType <> 1 then
                            		                tSttid(y) = ""
                            		                tSltid(y) = ""
                            		                tTimertildelt(y) = ""
                            		                end if  
                            		                
                            		                'end if 't 
			        
			                         'response.Write "y: nyt arr genneml�b: " & y & "<br>"
			                         'Response.flush
			                         '***** ER klokkeslet Sttid og Sltid benyttet eller timeangivelse ****'
			                         '*** tjekker for slet tSttid(y) <> 0 ***'
                                     if cint(visTimerelTid) <> 0 AND len(trim(tSttid(y))) <> 0 then
                                     
                                     if len(trim(tSttid(y))) <> 1 then ' == Slet
                                     
                                     
                                     'Response.Write "<br><b>tSttid(y):</b>" & tSttid(y) & "<br>"
                            			
                            			           
                            		                
                            				    
                                                    '**** Tjekker for tidsl�s ***'
                                                    '**** Skal timer l�gges p� en efterf�lgende akt? ***' 
                                                    weekdayThis = weekday(datoer(y), 2)
                                                    'Response.write "weekdayThis: " & weekdayThis & "<br>"
                            				        
                                                    select case weekdayThis
                                                    case 1
                                                    flt = "tidslaas_man"
                                                    case 2
                                                    flt = "tidslaas_tir"
                                                    case 3
                                                    flt = "tidslaas_ons"
                                                    case 4
                                                    flt = "tidslaas_tor"
                                                    case 5
                                                    flt = "tidslaas_fre"
                                                    case 6
                                                    flt = "tidslaas_lor"
                                                    case 7
                                                    flt = "tidslaas_son"
                                                    end select 
                            				        
                                                    tidslaas = 0
				                                    tidslaasDayOn = 0
                            				        
                            				        '** Finder grund aktivitet. Skal v�re = 1 fakturerbar.
                                                    strSQLtidslaas = "SELECT tidslaas, tidslaas_st, "_
                                                    &" tidslaas_sl, "& flt &" AS tidslaasDayOn FROM aktiviteter WHERE "_
                                                    &" id = "& aktids(j) & " AND fakturerbar = 1"
                                                    
                                                    'Response.Write strSQLtidslaas
                                                    'Response.flush
                                                    
                                                    oRec5.open strSQLtidslaas, oConn, 3
                                                    if not oRec5.EOF then
                            				        
                                                    tidslaas = oRec5("tidslaas")
                            				        
                                                    tidslaas_st = formatdatetime(oRec5("tidslaas_st"), 3)
                                                    tidslaas_sl = formatdatetime(oRec5("tidslaas_sl"), 3)
                            				        
                                                    tidslaasDayOn = oRec5("tidslaasDayOn")
                            				        
                                                    end if
                                                    oRec5.close
                                                    
                                                    'Response.Write "<br>A tidslaas " & tidslaas
                                                    'Response.Write "<br>A tidslaas_st " & tidslaas_st
                                                    'Response.Write "<br>A tidslaas_sl" & tidslaas_sl
                                                    'Response.Write "<br>A tidslaasDayOn "& tidslaasDayOn & "<br><br>"
                            				        
				                                    tidslaasErr = 0
				                                    
				                                       
                                                                 
                            				        
                                                    if tidslaas <> 0 then
                                                        
                                                        
                                                        
                                                            if tidslaasDayOn <> 0 then
                                                            
                                                            
                                                                     '*** Forl�nger Array strings ***'
        						                                     '** Hvis der er fundet en aktivitet forl�nges j arr.
        						                                     '** En indtasting via tidsl�s g�r forud for en manuel indtastning direkte i feltet ***'
                                                                    if instr(jarrforl, ",#"& aktids(j) &"#") = 0 then
                                                                    
                                                                    
                                                                    
                                                                    '*** Skal finde op til 8 aktiviteter ***'
                                                                    '** 1+2 Man-Fre Dag/Man-Tor Nat aktiviteter p� medarbejder (type:Dag+Nat)
                                                                    '** 3+4 Dag/Nat aktiviteter p� kunde (type:E1) 
                                                                    '** 5+6 Mandag morgen akt. aktivitet p� kunde (type:Nat+E1)
                                                                    '** 7+8 L�rdag morgen akt. aktivitet p� kunde (type:Nat+E1)
                                                                    '** 9+10 Fredag aften + S�ndag aften (type:Nat+E1)
                                                                    '** 11+12 Weekend Dag medarb + kunde (type:Weekend+E1)
                                                                    '** 13+14 Weekend Nat medarb + kunde (type:Weekend Nat+E1)
                                                                    
                                                                   
                                                                    cta = 7
                                                                   
                                                                    
                                                                    newArrSizeAkt = ubound(aktids)+100 
                                                                    Redim preserve jobids(newArrSizeAkt) 
                                                                    Redim preserve aktids(newArrSizeAkt)
                                                                    
                                                                    
                                                                    
                                                                    jarrforl = jarrforl & ",#"& aktids(j) & "#"
                                                                    
                                                                    '** G�r plads til 14 akt. (se ovenfor) ***'
                                                                    'st_high_fetnr = high_fetnr
                                                                    for aa = 1 to 100
                                                                    
                                                                    jobids(newArrSizeAkt-100+aa) = jobids(j)
                                                                    aktids(newArrSizeAkt-100+aa) = 0 'aktids(j) '=0?
                                                                    
                                                                    Redim preserve tTimertildelt(ubound(tTimertildelt)+cta)
                                                                    Redim preserve datoer(ubound(tTimertildelt)+cta)
                                                                    Redim preserve tSttid(ubound(tTimertildelt)+cta)
                                                                    Redim preserve tSltid(ubound(tTimertildelt)+cta)
                                                                    Redim preserve feltnr(ubound(tTimertildelt)+cta)
                                                                    
                                                                            for y2 = (ubound(tTimertildelt) - 6) to ubound(tTimertildelt)
                                                                                
                                                                                if y2 = (ubound(tTimertildelt) - 6) then
                                                                                ic = 1
                                                                                else
                                                                                ic = ic + 1
                                                                                end if
                                                                            
                                                                            tTimertildelt(y2) = -9001
                                                                            datoer(y2) = datoer(y)
                                                                            tSttid(y2) = ""
                                                                            tSltid(y2) = ""
                                                                            feltnr(y2) = high_fetnr+ic+3 
                                                                            '** 3 for at f� 7 til at blive n�ste 10'er
                                                                            '** dvs 17 bliver til 20 for at f� f�rste felt i n�ste r�kke
                                                                            
                                                                            'Response.Write "high_fetnr+y2:" & high_fetnr &" - "& y2&"<br>"
                                                                            'Response.Write "feltnr(y2):" &  feltnr(y2) & " y2:"& y2 &"<br>"
                                                                            '*** ialtcounter **'
                                                                            'iic = iic + 1
                                                                            next
                                                                            
                                                                            high_fetnr = high_fetnr + 10
                                                                    
                                                                    next
                                                                    
                                                                end if 'jarrforl
                                                                
                                                                
                                                                
                                                                
                                                                '*** Kig efter ny akt. der er �ben iflg. tidsl�s **'
                                                                '*** Kun hvis der ikke eer angivet NUL for slet **'
                                                                
                                                                    
                                                                    strGrundAktnavn = ""
                                                                    strSQLgrundAkt = "SELECT navn FROM aktiviteter WHERE id = "& aktids(j) 
                                                                    oRec5.open strSQLgrundAkt, oConn ,3
                                                                    if not oRec5.EOF then
                                                                    
                                                                    strGrundAktnavn = oRec5("navn")
                                                                    
                                                                    end if
                                                                    oRec5.close
                                                                    
                                                                    'select case weekdayThis
                                                                    'case 5 'fre
                                                                    'eksFlt = " OR (tidslaas_lor = 1 AND fakturerbar = 53)" 'weekend nat
                                                                    'case else
                                                                    eksFlt = ""
                                                                    'end select
                                                                    
                                                                    eksFlt = ""
                                                                    
                                                                    aktfundet = 0
                                                                    strSQLfind = "SELECT a.id AS id, a.navn, a.job, tidslaas, tidslaas_st, "_
                                                                    &" tidslaas_sl, "& flt &" AS tidslaasDayOn, a.fakturerbar, "_
                                                                    &" tidslaas_man, tidslaas_tir, tidslaas_ons, tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son"_
                                                                    &" FROM job j "_
                                                                    &" LEFT JOIN aktiviteter a ON (("_
                                                                    &" a.id <> "& aktids(j) & " AND tidslaas = 1 AND ("& flt &" = 1 "& eksFlt &")) AND "_
                                                                    &" (a.fakturerbar = 91 OR a.fakturerbar = 90 OR a.fakturerbar = 50 OR a.fakturerbar = 51 OR a.fakturerbar = 52 OR a.fakturerbar = 53)"_
                                                                    &" AND a.aktstatus <> 0 AND a.job = j.id AND navn LIKE '"& left(strGrundAktnavn, 6) &"%')"_
                                                                    &" WHERE j.id = "& jobids(j) &" AND a.job <> ''"
                                                                    
                                                                    
                                                                   
                                                                    
                                                                    'Response.Write  "lasty2: "& lasty2 & "<br>"
                                                                    'Response.Write "<b>"& strSQLfind & "</b><br>"
                                                                    'Response.flush
                            				                        
                            				                        gnlob = 1
                            				                        'icc = 0
                                                                    oRec5.open strSQLfind, oConn, 3
                                                                    while not oRec5.EOF 
                                                                    
                                                                    
                                                                    
                                                                    'Response.Write "<br><u>Unders�ger: "& aktids(j) &" Finder:"& oRec5("id") & "</u><br><br>"
                                    						        'aktidprfelt(y) = oRec5("id")
                                    						        
                                    						                                 
                                    						        
                                    						        
                                    						        B_type = oRec5("fakturerbar")
                                    						        
                                                                    tidslaas_B = oRec5("tidslaas")
                                    						        
                                                                    tidslaas_st_B = formatdatetime(oRec5("tidslaas_st"), 3)
                                                                    tidslaas_sl_B = formatdatetime(oRec5("tidslaas_sl"), 3)
                                    						        
                                                                    tidslaasDayOn_B = oRec5("tidslaasDayOn")
                                                                    
                                                                    tidslaasManOn_B = oRec5("tidslaas_man")
                                                                    tidslaasTirOn_B = oRec5("tidslaas_tir")
                                                                    tidslaasOnsOn_B = oRec5("tidslaas_ons")
                                                                    tidslaasTorOn_B = oRec5("tidslaas_tor")
                                                                    tidslaasFreOn_B = oRec5("tidslaas_fre")
                                                                    tidslaasLorOn_B = oRec5("tidslaas_lor")
                                                                    tidslaasSonOn_B = oRec5("tidslaas_son")
                                                                    
                                    						        
        						                                    strAktNavn_B = replace(oRec5("navn"), "'", "")
                                    						        
        						                                    
        						                                    '*** Redigere tidspunkter efter tidsl�s ***'
        						                                    
        						                                    '* Hvis tidsl�s A = 1
        						                                    '* Hvis tidsl�s A for dagen = 1
        						                                    
        						                                    
        						                                   
        						                                    '*** Trimmer grundreg **'
        						                                    if cint(gnlob) = 1 then
        						                                    
        						                                    '*** Tjekker Dato format her ***'
        						                                    
        						                                    'KODE
        						                                    
        						                                    '****'
        						                                    gRegSt = tSttid(y)
        						                                    gRegSl = tSltid(y)
        						                                    
        						                                    
        						                                    
        						                                            
        						                                           
        						                                    end if 'gnlob   
        						                                    
        						                                    
        						                                    
        						                                    if request("minindtast_on") = "1" then
        						                                    '**** Min 7,4 timer p� medarb / 8 timer p� kunde ****'
					                                                '**** Det skal v�re samlet for dagen             ****'
					                                                '**** hvilken akt. skal den placere det p�?      ****'
					                                                '**** Denne beregning skal ligge i INC filen     ****'
    					                                            
    					                                            idag = formatdatetime(now, 2)
                                                                    minutberegnDiffmin = datediff("n", idag &" "& gRegSt, idag &" "& gRegSl, 2, 2) 
                                						                                           
    					                                            
					                                                select case B_type
					                                                case "50", "51", "52", "53", "54", "55"
					                                                    if minutberegnDiffmin < 444 AND minutberegnDiffmin > 0 then
					                                                    tillaeg = (444 - minutberegnDiffmin) 
					                                                    gRegSl = dateAdd("n", tillaeg, idag &" "& gRegSl)
					                                                    gRegSl = left(formatdatetime(gRegSl, 3), 5)
					                                                    
					                                                    end if
    					                                                
    					                                               
					                                                    if minutberegnDiffmin < 0 then 'AND minutberegnDiffmin > -444 then
					                                                    
					                                                        imorgen = formatdatetime(dateadd("d", 1, idag), 2)
					                                                        minutberegnDiffmin = datediff("n", idag &" "& gRegSt, imorgen &" "& gRegSl, 2, 2) 
                                    						            
					                                                        tillaeg = (444 - (minutberegnDiffmin)) 
					                                                        gRegSl = dateAdd("n", tillaeg, idag &" "& gRegSl)
					                                                        gRegSl = left(formatdatetime(gRegSl, 3), 5)
    					                                                    
					                                                    end if
    					                                            
					                                                case "90", "91"
    					                                                
    					                                                
					                                                    if minutberegnDiffmin < 480 AND minutberegnDiffmin > 0 then
					                                                        tillaeg = (480 - minutberegnDiffmin) 
					                                                        gRegSl = dateAdd("n", tillaeg, idag &" "& gRegSl)
					                                                        gRegSl = left(formatdatetime(gRegSl, 3), 5)
					                                                    end if
    					                                                
    					                                                
					                                                    if minutberegnDiffmin < 0 then 'AND minutberegnDiffmin > -480 then
    					                                                    
					                                                        imorgen = formatdatetime(dateadd("d", 1, idag), 2)
					                                                        minutberegnDiffmin = datediff("n", idag &" "& gRegSt, imorgen &" "& gRegSl, 2, 2) 
                                    						            
					                                                        tillaeg = (480 - (minutberegnDiffmin)) 
					                                                        gRegSl = dateAdd("n", tillaeg, idag &" "& gRegSl)
					                                                        gRegSl = left(formatdatetime(gRegSl, 3), 5)
    					                                                    
    					                                                    
    					                                                    'Response.Write "tillaeg"
    					                                                    'Response.end
					                                                    end if
    					                                            
					                                                'case else
					                                                'minutberegnDiffmin = minutberegnDiffmin
					                                                end select
        						                                    
        						                                    
        						                                    end if '*** force min indtast 7,4 / 8,0
        						                                    
        						                                   
        						                                   
        						                                     
        						                                    '*** Er tidsl�s sl�ettil p� dagen ***'
        						                                     if cint(tidslaas_B) = 1 AND cint(tidslaasDayOn_B) = 1 then
        						                                    
        						                                     aktfundet = 1
        						                                   
        						                                          
                                                                    '*** Finder de korrekte tidspunkter ***'
                                                                    
                                                                    'if cint(gnlob) = 1 then
                                                                    'if y = ystart then 
                                                                    if lasty2 = 0 then
                                                                    gnlob_st_high_fetnr = UBOUND(feltnr) - (707) '(105) '7*14 = 98 (+ 7)
                                                                    else
                                                                    gnlob_st_high_fetnr = gnlob_st_high_fetnr + 7 '** 7 dage
                                                                    end if
                                                                    
                                                                    'icc = 0
                                                                                    
                                                                                    for y2 = gnlob_st_high_fetnr + 1 to gnlob_st_high_fetnr + 7 '*** L�gges 7 til for 7 dage ***'
                                                                                    
                                                                                           
                                                                                            
                                                                                            aktids(newArrSizeAkt-100+icc+1) = oRec5("id") '+gnlob
                                                                                            
                                                                                            if right(feltnr(y2), 1) = right(feltnr(y),1) then
                                                                                            
                                                                                             
                                                                                                'y2 = y2 + (cint(right(feltnr(y),1)) * 7)
                                                                                            
                                                                                            
                                                                                                     '*** Trimmer de funde aktiviteter **'
        						                                                                     'Response.Write "<b>Aktid:</b>"& oRec5("id") &"/"& aktids(newArrSizeAkt-7+gnlob) &"y2 val: " & y2 &"/"& newArrSizeAkt-7+gnlob &" gRegSt: "& gRegSt & " til "& gRegSl &" - tidslaas_st_B: " & tidslaas_st_B & " tidslaas_sl_B: " & tidslaas_sl_B & "cint(right(feltnr(y2), 1)) >= cint(weekdayThis)"& cint(right(feltnr(y2), 1)) &" >= "& cint(weekdayThis) &"<br>"
                                        						                                     
                						                                                            '*** Er sttid st�rre end start tidsl�s og er sttid mindre end tidsl�s slut.
        						                                                                    if (cDate(gRegSt&":00") >= cDate(tidslaas_st_B) AND cDate(tidslaas_st_B) >= cDate(tidslaas_sl_B)) _ 
        						                                                                    OR (cDate(gRegSt&":00") >= cDate(tidslaas_st_B) AND cDate(gRegSt&":00") <= cDate(tidslaas_sl_B))then 
        						                                                                        'if cDate(gRegSt&":00") > cDate(tidslaas_sl_B) then
                                                                                                        'tSttid(y2) = ""
                                                                                                        'else
                                                                                                        tSttid(y2) = gRegSt
                                                                                                        'end if
                                                                                                        'Response.Write "Bruger st.tid: " & gRegSt 
                                                                                                    else 
                                                                                                    tSttid(y2) = left(tidslaas_st_B, 5)
        						                                                                        'Response.Write "Bruger tidsl�s start: " & left(tidslaas_st_B, 5)
        						                                                                    end if
                                						                                            
                                						                                            
        						                                                                    '* Hvis angivet SLUT tidspunkt er MINDRE end tidsl�s slut A
        						                                                                    '>>
        						                                                                    '** Er slut tid n�ste morgen, dvs efter 24:00 **'
        						                                                                    '_
        						                                                                    'OR (cDate(gRegSl&":00") <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) <= cDate(gRegSt&":00"))
        						                                                                    
        						                                                                    if (cDate(gRegSl&":00") <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) <= cDate(tidslaas_sl_B) AND cDate(gRegSl&":00") > cDate(tidslaas_st_B)) _ 
        						                                                                    OR (cDate(gRegSl&":00") > cDate(tidslaas_st_B) AND cDate(gRegSl&":00") > cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) >= cDate(tidslaas_sl_B)) _
        						                                                                    OR (cDate(gRegSl&":00") <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) >= cDate(tidslaas_sl_B) AND cDate(gRegSt&":00") >= cDate(gRegSl&":00")) _
        						                                                                    OR (cDate(gRegSl&":00") <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) <= cDate(gRegSt&":00") AND cDate(tidslaas_sl_B) <= cDate(gRegSt&":00")) then
        						                                                                    '===   TID   ======    DAG      =====   NAT   ========
        						                                                                    '1
        						                                                                    '2
        						                                                                    '3 21:00 - 07:00 => 06:00 - 07:00 => 21:00 - 06:00
        						                                                                    '5 10:45 - 04:30 => 10:45 - 18:00 => 18:00 - 04:30    
        						                                                                    '=====================================================  
        						                                                                    tSltid(y2) = gRegSl
        						                                                                    else
        						                                                                    tSltid(y2) = left(tidslaas_sl_B, 5)
        						                                                                    end if
        						                                                                    '* Hvis angivet SLUT tidspunkt er ST�RRE end tidsl�s slut A
        						                                                                    '>>
                                						                                    
        						                                                                    idag = formatdatetime(now, 2)
        						                                                                    minutberegning = datediff("n", idag &" "& tSttid(y2), idag &" "& tSltid(y2), 2, 2) 
                                						                                            
                                						                                            
                                						                                            
                                						                                            
        						                                                                    'Response.Write "minutberegning tidsl�s st: "& tidslaas_st_B &" tidsl�s sl: "& tidslaas_sl_B &" gregST: "& cDate(gRegSt) &" gRegSl: "& cDate(gRegSl) &" yST: "& cDate(tSttid(y2)) &" ySL "&tSltid(y2)&" typ: "& B_type &": " & minutberegning & "<br>"
        						                                                                    'Response.flush
                                						                                            
        						                                                                    '*** Tjekker om der skal indl�ses p� akt selvom minutbe. er negativt **'
        						                                                                    '*** AND cDate(tSltid(y2)) < cDate(gRegSl)
        						                                                                    if (minutberegning < 0 AND cDate(tSttid(y2)) > cDate(gRegSt) AND cDate(gRegSt) < cDate(gRegSl))_
        						                                                                    OR (minutberegning < 0 AND cDate(tidslaas_sl_B) < cDate(tidslaas_st_B) AND cDate(tidslaas_sl_B) <= cDate(gRegSl) AND cDate(tidslaas_st_B) <= cDate(tSltid(y2)))_
        						                                                                    OR (minutberegning < 0 AND cDate(tidslaas_st_B) < cDate(tidslaas_sl_B) AND cDate(gRegSt) >= cDate(tidslaas_sl_B) AND cDate(gRegSl) <= cDate(tidslaas_st_B))_
        						                                                                    
        						                                                                    then
        						                                                                    
        						                                                                    'Response.Write "minutbereg. negativt, indtastning IKKE godkendt<br><br>"
        						                                                                    
        						                                                                    
        						                                                                    tSttid(y2) = ""
        						                                                                    tSltid(y2) = ""
        						                                                                    end if
                                                                                                    
                                                                                                    
                                                                                                    '** Korrigere dato (l�r morgen / Mandag morgen mv.)
                                                                                                    datoer(y2) = datoer(y)
                                                                                                    'if cint(weekdayThis) = 5 AND cint(tidslaasFreOn_B) = 0 AND cint(tidslaasLorOn_B) = 1 then 
                                                                                                    'datoer(y2) = dateadd("d", 1, datoer(y))
                                                                                                    'feltnr(y2) = feltnr(y2) + 1
                                                                                                    'end if
                                                                                                    'feltnr(y2) = feltnr(y)
                                                                                                    tTimertildelt(y2) = ""
                                                                                                    
                                                                                                    'Response.Write "<b>Resultat af fundet akt:</b><br>"
                                                                                                    'Response.Write "FELTNR: "& feltnr(y2) &" y2 val2 /y: "& y2 &"/ "& y &" # "& tSttid(y2) &" - "& tSltid(y2) & " Dato" & datoer(y2) & "dserial: <br><br>"
                                                                                                    'lastUsedy2serie = y2
                                                                                                    
                                                                                                    
                                                                                                    'tSttid(y) = tSttid(y2)
        						                                                                    'tSltid(y) = tSttid(y2)
        						                                                                    'tTimertildelt(y2) = ""
                                                                                                    
                                                                                            else
                                                                                            
                                                                                            'Response.Write "Nulstiller felt<br>"
                                                                                            
                                                                                            'tSttid(y2) = ""
        						                                                            'tSltid(y2) = ""
        						                                                            'tTimertildelt(y2) = -9001
                                                                                            
                                                                                            end if
                                                                                            
                                                                                          
                                                                                    lasty2 = y2
                                                                                    next
                                                                                    
                                                                   
                                                                    end if 'tidsl�s sl�et til
                                                                    
                                                                    
                                                                    icc = icc + 1  
        						                                    gnlob = gnlob + 1
                                                                    oRec5.movenext
                                                                    wend 
                                                                    oRec5.close
                                                                    
                                                 
                                                                    
                                                                    
                                                                   
                                                                    
                                                            '*** err hvis der ikke findes en akt med en d�kkende tidsl�s ***'
                                                            if cint(aktfundet) = 0 then
                                                             tidslaasErr = 1
                                                            end if
                                                                        
                                                                        
                                                                        
                                                                
                                                        
                                                else
                                                
                                                tidslaasErr = 1
                                                
                                                end if 'tidslaasDayOn
                                                
                    				        
                                            end if 'tidsl�as
                                                    
                                                    
                                                    
                                                    
                                                    'if tidslaasErr <> 0 then
                                                    if tidslaasErr = 1110 then
                                                    %>
                                                    <!--#include file="../../inc/regular/header_lysblaa_inc.asp"-->
                                                    <% 
                                    			    
                                                    useleftdiv = "t"
                                                    errortype = 136
                                                    call showError(errortype)
                                    		        
                                                    Response.end
                                                    end if
                                                    
                                                    
                            			        
                            			
                                        end if 'slet
                            		
                                    else
                            		
                                    '** Hvis alm. timereg er benyttet, 
                                    '** dvs der ikke er angivet klokkeslet for start og slut ***'
                                    sTtid = ""
                                    sLtid = ""
                            		
                                    end if
                                   
                                   
                                   
                                   
                                   next 'y
                                   next 'j 
                                    
				                    
				                    
				                    
				                            
				                
				                'Response.Write "aktids(newArrSizeAkt)<br>" 
				                
				                for j = 0 to UBOUND(aktids) 
				                'Response.Write aktids(j) & " == "
				                
				                     if j = 0 then
			                        ystart = 0
			                        else
			                        ystart = (j * 7) 
			                        end if
                        			
			                        yslut = (ystart + 6)
			                     
			                     'oprArrSize = ubound(tTimertildelt)
			                     for y = ystart to yslut    
			                        
			                        'Response.Write  feltnr(y) & ":("& y &"):" & tSttid(y) &" - "& tSltid(y) & " | "
			                        
			                        
			                     next
			                     
			                     
			                     'Response.Write "<hr>"
				                
				                next
				                
				                'Response.end        

end sub
%>





