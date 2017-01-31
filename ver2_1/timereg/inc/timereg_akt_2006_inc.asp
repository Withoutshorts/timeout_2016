<%'GIT 20160811 - SK



 


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
public strMnavn, dblkostpris, tprisGen, valutaGen, mkostpristarif_A, mkostpristarif_B, mkostpristarif_C, mkostpristarif_D 
function mNavnogKostpris(strMnr)
'** Henter navn og kostpris ***'

SQLmedtpris = "SELECT medarbejdertype, timepris, tp0_valuta, kostpris, mnavn, "_
&" kostpristarif_A, kostpristarif_B, kostpristarif_C, kostpristarif_D FROM medarbejdere, medarbejdertyper "_
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

        mkostpristarif_A = oRec("kostpristarif_A")
        mkostpristarif_B = oRec("kostpristarif_B")
        mkostpristarif_C = oRec("kostpristarif_C") 
        mkostpristarif_D = oRec("kostpristarif_D")
		
		end if

oRec.close

'** Slut timepris **


end function



'*********************************************************************************************************************************
'**** Opdaterer timer ************************************************************************************************************
'*********************************************************************************************************************************
public EmailNotificerTxt, EmailNotificer, EmailNotificerTxt2, EmailNotificer2, EmailNotificerTxt3, EmailNotificer3 

function opdaterimer(aktid, aktnavn, tfaktimvalue, strFastpris, jobnr, strJobnavn, strJobknr, strJobknavn,_
medid, strMnavn, datothis, timerthis, kommthis, intTimepris,_
dblkostpris, offentlig, intServiceAft, strYear, sTtid, sLtid, visTimerelTid, stopur, intValuta, bopal, destination, dage, tildeliheledage, origin, extsysid, mtrx)



if len(timerthis) <> 0 then
timerthis = timerthis
else
timerthis = 0
end if


if cint(timerround) = 1 AND cint(intEasyreg) <> 1 then 'fra cookie
if tfaktimvalue = 1 OR tfaktimvalue = 2 then 'KUN fakturerbare / ikke fakbare timer 

 
    if instr(timerthis, ".") <> 0 then
    instr_timerthis = instr(timerthis, ".") 
    len_timerthis = len(timerthis)
    right_timerthis = mid(timerthis, instr_timerthis+1, len_timerthis)
    
    if len(right_timerthis) < 2 then
    right_timerthis = right_timerthis&"0"
    end if
    right_timerthis = left(right_timerthis,2)
    'right_timerthis = right(right_timerthis, 2)


    if cdbl(right_timerthis) > 0 AND cdbl(right_timerthis) <= 50 then
    left_timerthis = mid(timerthis, 1, instr_timerthis-1)
    timerthis = left_timerthis&".5"  
    end if

    if cdbl(right_timerthis) > 50 then
    timerthis = replace(timerthis, ".",",")
    timerthis = round(timerthis)
    end if

    else
    timerthis = timerthis
    end if
   
    'response.write " = timerthis:" & timerthis &" right_timerthis: "&  right_timerthis & "    <br>"

else
timerthis = timerthis
end if
end if


brug_fasttp = 0
brug_fastkp = 0
'*** Tjekker om aktiviteten er sat til ens timpris for alle medarbejdere (overskriver medarbejderens egen timepris)
strSQLtjkAktTp = "SELECT brug_fasttp, brug_fastkp, fasttp, fasttp_val, fastkp, fastkp_val FROM aktiviteter WHERE id = "& aktid
oRec5.open strSQLtjkAktTp, oConn, 3
if not oRec5.EOF then

if cint(oRec5("brug_fasttp")) = 1 then
brug_fasttp = 1
fasttp = oRec5("fasttp")
fasttp = replace(fasttp, ".", "")
fasttp = replace(fasttp, ",", ".")

fasttp_val = oRec5("fasttp_val")
end if


if cint(oRec5("brug_fastkp")) = 1 then
brug_fastkp = 1
fastkp = oRec5("fastkp")
fastkp = replace(fastkp, ".", "")
fastkp = replace(fastkp, ",", ".")
end if

end if
oRec5.close 


'** Salg og kostpris altid = 0 hvis Salgs timer **
dblkostprisUse = dblkostpris

if cint(brug_fasttp) = 1 then
intTimepris = fasttp
intValuta = fasttp_val
end if

if cint(brug_fastkp) = 1 then
dblkostprisUse = fastkp
end if


'** Sepcial setting Epinion ***'
'select case lto
'case "epi", "intranet - local"
'if tfaktimvalue = 6 then
'dblkostprisUse = 0
'intTimepris = 0
'end if
'end select

     

dblkostprisUse = replace(dblkostprisUse, ".", "")
dblkostprisUse = replace(dblkostprisUse, ",", ".")
    
    call akttyper2009prop(tfaktimvalue)
    
    call valutaKurs(intValuta)


         

    
             '** Kun hvis tildeliheledage er slået til
             '** DET SKAL DEN KUN VÆRE VED MULTI TILDEL
             if cdbl(timerthis) <> -9001 AND  cint(tildeliheledage) = 1 then 'tom

                '*** Timer på Ferie og Feriefridage ***'
                '*** Omregn fra hele dage til timer iflg. norm tid ***'
                'if muDag = 0 then '** Kun ved første indtastning på hver medarbejder ***'
                select case tfaktimvalue
                case "900"' = Dvs ingen, beregning er slået fra 
                case 11,12,13,14,15,16,17,18,19,20,21,22,23,91,111,112,120,121,122,125 '** Ferie typer + syg og barn syg + barsel + omsorgsdag + E2 **'
                  
                       
    	               
                            select case tfaktimvalue
               
                            case 11,12,15,18,120,121,122,125 '** Ferie Planlagt & Optjent, Rejsedage = BRUG UGE NORM **' (Så det kan tastes på alle dage)

                    
                                       '** Findes der nomrtimer på den valgte dag **'
    	                               ntimPer = 0
    	                               call normtimerPer(medid, datoThis, 6, 0)
                                       normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
            
    	    
    	                                if cdbl(normTimerGns5) > 0 then

                                        dage = replace(dage, ".", ",")

                                        if len(trim(dage)) <> 0 then
                                        dage = dage
                                        else
                                        dage = 0
                                        end if

                                        timerthis = formatnumber(dage * normTimerGns5, 2)
                                        timerthis = replace(timerthis, ".", "")
                                        timerthis = replace(timerthis, ",", ".")

                                        else

                                        timerthis = timerthis

                                        end if


                            case else '*** Afholdt: BRUG DAGSNORM (så det følger den enekelte dag)


                                        '** Findes der nomrtimer på den valgte dag **'
    	                                ntimPer = 0
    	                                call normtimerPer(medid, datoThis, 0, 0)
                                       
    	    
    	                                if cdbl(ntimPer) > 0 then

                                        dage = replace(dage, ".", ",")

                                        if len(trim(dage)) <> 0 then
                                        dage = dage
                                        else
                                        dage = 0
                                        end if

                                        timerthis = formatnumber(dage * ntimPer, 2)
                                        timerthis = replace(timerthis, ".", "")
                                        timerthis = replace(timerthis, ",", ".")

                                        else

                                        timerthis = timerthis

                                        end if        


                            end select
               
    	        
                    
                end select
               
    
          end if '9001
    


           '********* EMAIL notificering til leder eller hvis leder har ændret på medarbejders timer **********
           if cdbl(timerthis) <> -9001 AND cdbl(timerthis) <> -9002 then 'Tom / Opdater værdier der ikke er ændret
       

                    '*** Email notifikation 1 til TEAMLEDER på følgende typer ***'
                    select case tfaktimvalue
                    'case 11,13,14,18,19,20,21,23,26,120,121,122 '** Ferie typer planlagt og afholdt + syg og barn syg + omsorgsdag PL + Aldersreduktion PL + Omsorg 2,10, K PL **'
                    case "900"
                    case 1,2,5,6 'alle andre typer end fakturerbar og ikke fakturerbar + Km og SALG bliver adviseret til teamleder, der abonnerer.

                    case else

                            select case lto
                            case "epi2017"
                            EmailNotificer = 0
                            case else
                            EmailNotificer = 1
                            end select

                    timerthisENotiTxt = replace(timerthis, ".", ",")
                    EmailNotificerTxt = EmailNotificerTxt & strMnavn & " har registreret "& timerthisENotiTxt & " timer på "& aktnavn & " d. "& formatdatetime(datoThis, 2) &"</br>"
                    end select


            
    
                    '*** Email notifikation 2 til MEDARBEJDER HVIS DER BLIVER INDTASTET AF LEDER på følgende typer ***'
                    select case tfaktimvalue
                    case 20,21 '** Syg og barn syg  indtastet af leder **'
            
                    if medid <> session("mid") then
                    EmailNotificer2 = 1

                    timerthisENotiTxt = replace(timerthis, ".", ",")
                    EmailNotificerTxt2 = EmailNotificerTxt2 & "Der er registreret "& timerthisENotiTxt & " timer, for medarbejder "& strMnavn &", på "& aktnavn & " d. "& formatdatetime(datoThis, 2) &"</br>"
                    end if        


                    end select


                    select case lto 
                    case "esn"

                            '*** Email notifikation 3 til ENIGA ***'
                            select case tfaktimvalue
                            case 20,21 '** Syg og barn syg  ALLE sygsomstyper **'

                            '	Sygdom
                            '	Barns 1. sygedag
                            '	Barns 2. sygedag
                            '	Arbejdsskade
                            '	Sygdom efter særlig aftale (§56 eller ansat i fleksjob)
                            '	Indlagt med barn
                            '	Graviditetsbetinget sygdom
                            '	Dato for sygemeldingen

            
                   
                            EmailNotificer3 = 1

                            timerthisENotiTxt = replace(timerthis, ".", ",")
                            EmailNotificerTxt3 = EmailNotificerTxt3 & "Der er registreret "& timerthisENotiTxt & " timer, for medarbejder "& strMnavn &", på "& aktnavn & " d. "& formatdatetime(datoThis, 2) &"</br>"
                  

                            end select

                    end select

            end if '9001 9002


   
			
	if cdbl(timerthis) <> -9001 AND cdbl(timerthis) <> -9003  then '*** <> Slet: -9001 // FRA TT: -9003 ***'
			
			
			datothis = ConvertDateYMD(datothis)
			        
			        
                    	'Response.write "visTimerelTid: "& len(trim(sTtid)) &" : "& sTtid &" - "& sLtid &"<br>"
						'Response.flush

			
					if visTimerelTid <> 0 AND len(trim(sTtid)) <> 0 then
						
					
						
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



                                '**** NOSTOP MATRIX ******************************* 
                                'response.write "<br><br>ORG: "& mtrx &" timerthis: "& timerthis & " sTtid: "& sTtid & " sLtid: "& sLtid & " datepart: " & datepart("w", datoThis, 2,2)
                                
                               
                                 if cint(matrixAktFundet) = 1 then
                                
                                               call matrixtimespan(idag, mtrx, sTtid, sLtid, datoThis)
                                               timerthis = timerthis_mtx
        
                                    end if 'MatrsixFundet









                                '******************************************

						end if
					
					else
					
					sTtid = ""
					sLtid = ""
					
					end if
					
			
			if cint(origin) = 11 OR cint(origin) = 12 then 'ved indlæsning fra TimeTag_web / ugeseddel skal der altid indlæses aldrig opdateres 
			strSQLfindes = "SELECT timer, Tjobnr, TAktivitetId, timerkom FROM timer WHERE TAktivitetId = -2130 AND Tjobnr = 'WWvsfWWx445' AND Tmnr = -23423440"
            else
            strSQLfindes = "SELECT timer, Tjobnr, TAktivitetId, timerkom FROM timer WHERE TAktivitetId = "& aktid &" AND Tjobnr = '"& jobnr &"' AND Tmnr = "& medid &" AND Tdato = '"& datothis &"'"
            end if
			oRec.Open strSQLfindes, oConn, 3  
			
			'son, man, tir, ons ,tor, fre, lor
			if oRec.EOF then
				
                'Response.Write timerthis & "<br>"

                            'if session("mid") = 1 then
				            'Response.write "aty_pre: "& aty_pre &" timerthis: "& timerthis & "<br>"
    				        'response.end
                            'end if 

				
				if len(timerthis) > 0 AND cdbl(timerthis) > -9000 AND timerthis <> 0 _
				OR (len(timerthis) > 0 AND cdbl(timerthis) > -9000 AND aty_pre <> 0) then 
				
				'** AND timerthis > 0  then -minus er OK HUSK ret i faktrua opr. 15.05.2008 **'
				'** Rettet tilbage 20.05.2008 pga fejl i SponsorCar ***' 
				'** De bruger st. tid og slut tid regsistrering **'
				
				strSQLins = "INSERT INTO timer (Tjobnr, Tjobnavn, Tmnr, Tmnavn, Tdato, "_
				&" Timer, Timerkom, Tknavn, Tknr, TAktivitetId, TAktivitetNavn, "_
				&" Tfaktim, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
				&" editor, kostpris, offentlig, seraft, sttid, sltid, valuta, kurs, bopal, destination, origin, extsysid) VALUES"_
				& "('" & jobnr & "', '"& strJobnavn &"', " & medid & ", '" & cstr(strMnavn) & "', '"& datothis &"', "_
				&" "& timerthis &", '"& SQLBless2(kommthis) &"', "_
				&" '" & SQLBless2(strJobknavn) & "', " & strJobknr & ", "_
				&" "& aktid &", '"& SQLBless2(aktnavn) &"', "_
				&" "& tfaktimvalue &", "& strYear &", "& SQLBless(intTimepris) &", "_
				&" '"&year(now)&"/"&month(now)&"/"&day(now)&"', '"& strFastpris  &"', "_
				&" '" & time & "', '"& session("user") &"', "& dblkostprisUse &", "_
				&" "& offentlig &", "& intServiceAft &", '"&sTtid&"', '"&sLtid&"', "_
				&" "& intValuta &", "& dblKurs &", "& bopal &", '"& destination &"', "& origin &", '"& extsysid &"')"
				
				end if
				
				
				if len(trim(strSQLins)) <> 0 then
				    'if session("mid") =  1 then
                    'Response.Write "Ins: "& strSQLins & "<br>"
				    'Response.flush
                    'end if
					oConn.Execute(strSQLins)


                    '*** Hvis indlæsning kommer fra afvigelses håndtering ***'
                    if origin = 11 then 'timetag afvigelses håndtering
                    sqlInTimerImpTemp = "UPDATE timer_import_temp SET Overfort = 1 WHERE id = "& extsysid
                    oConn.Execute(sqlInTimerImpTemp)
                    end if
            

				end if
			
			else
				
					'** Sletter ***'
					'** En preudfyldr aktivtet kan ikke slettes (da den i så fald igen vil være preudfyldt)
					'** Men der kan til gengæld angives 0 ****'
                    
					'Response.Write "<br>timerthis" & timerthis & " pre:"& aty_pre &"<br>"
					
					if (timerthis = 0 AND aty_pre = 0) then
					
					strSQLdel = "DELETE FROM timer"_
					&" WHERE Tjobnr = '"& jobnr & "'"_
					&" AND Tmnr = "& medid & ""_
					&" AND Tdato = '" & datothis & "' AND TAktivitetId = "& aktid & ""
					
					oConn.execute(strSQLdel)
					
					'Response.Write strSQLdel & "<br>"
					
					
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
						&" seraft = "& intServiceAft &", kostpris = "& dblkostprisUse &", "_
						&" Timerkom = '"& kommentarthis &"', "_
						&" tastedato = '"&year(now)&"/"&month(now)&"/"&day(now)&"', "_
						&" editor = '"& session("user") &"', offentlig = "& offentlig &", "_
						&" sttid = '"&sTtid&"', sltid = '"&sLtid&"', valuta = "& intValuta &", kurs = "& dblKurs &", bopal = "& bopal &", "_
						&" destination = '"& destination &"'"_
						&" WHERE Tjobnr = '"& jobnr & "'"_
						&" AND Tmnr = "& medid & ""_
						&" AND Tdato = '" & datothis & "' AND TAktivitetId = "& aktid
						
						'Response.write "Upd A: "& strSQLupd & "<br>"
						
						else
						
						'*************************************************************************************************'
						'** Opdatering af reg. tidspunkt på eksisterende timeregistreringer                            **'
						'** Tidspukter bliver ikke opdateret her, da man ellers vil overskrive tidspunkter              **'
						'** på eksisterende records, ved opdatering af en vilkårlig anden timeregistrering (som timer)  **'
						'*************************************************************************************************'
						
						
						strSQLupd = "UPDATE timer SET"_
						&" Timer = "& timerthis &", tfaktim = "& tfaktimvalue &", Timerkom = '"& kommthis &"', "_
						&" editor = '"& session("user") &"', offentlig = "& offentlig &", bopal = "& bopal &", destination = '"& destination &"'"_
						&" WHERE Tjobnr = '"& jobnr & "'"_
						&" AND Tmnr = "& medid & ""_
						&" AND Tdato = '" & datothis & "' AND TAktivitetId = "& aktid 
						
						'Response.write "Upd B: "& strSQLupd & "<br>"
						
						end if
						
						'response.write "<br>strSQLupd: "& strSQLupd & "<br>"
					    'Response.end
						oConn.Execute(strSQLupd)
						
					end if
			end if
			oRec.close
			
	end if '** timer -9001 (-1) ***'


end function




public maxl, fmbgcol, fmborcl, ugeafsluttetTxt, tfeltwth
function fakfarver(lastfakdato, varTjDato_man, varTjDato_ugedag, erIndlast, medid, akttype, tidslaasMan, tidslaasTir, tidslaasOns, tidslaasTor, tidslaasFre, tidslaasLor, _
tidslaasSon, job_tidslass, job_tidslassSt, job_tidslassSl, ignorertidslas, godkendtstatus, job_internt, origin, resforecastMedOverskreddet)

'ugeafsluttetTxt = "" *MÅ ALDRIG VÆRE TOM, da indstilling fra tidligere ellers glemmes.

 


'Response.Write "lastfakdato" & lastfakdato
'Response.end
tfeltwth = 45
timenow = formatdatetime(now, 3)

'call erugeAfslutte(year(tjekdag(4)), tjekdag(7), usemrn)
'** Sætter farve på indtastningfelt efter om der er udskrevet en faktura eller ej ***

strMrd_sm = datepart("m", tjekdag(x), 2, 2)
strAar_sm = datepart("yyyy", tjekdag(x), 2, 2)
strWeek = datepart("ww", tjekdag(4), 2, 2)
strAar = datepart("yyyy", tjekdag(4), 2, 2)
diff = dateDiff("d", lastfakdato, tjekdag(1))
ugeafsluttet = 0

'response.write "SmiWeekOrMonth: "& SmiWeekOrMonth & "<br><br>"

    if cint(SmiWeekOrMonth) = 0 then
    usePeriod = strWeek
    useYear = strAar
    else
    usePeriod = strMrd_sm
    useYear = strAar_sm
    end if

call erugeAfslutte(useYear, usePeriod, usemrn, SmiWeekOrMonth, 0)


'Response.Write "lastfakdato" & lastfakdato & "strMrd_sm: "& strMrd_sm & " strWeek: "& strWeek & " usemrn:" & usemrn
'Response.end

'** Quickfix, ugeNrAfsluttet bliver slet ikka kaldt og derfor ikke sat
'** og bliver af VB sat = 52, hvorfor man ikke kan registrere timer 
'if ugeNrAfsluttet <> "" then
'ugeNrAfsluttet = ugeNrAfsluttet
'else
'ugeNrAfsluttet = dateadd("d", -60, now)
'end if

'if session("mid") = 1 then
'Response.Write "varTjDato_ugedag" & varTjDato_ugedag
'Response.Write "#"& ugeNrAfsluttet &"#"& datepart("ww", ugeNrAfsluttet, 2, 2) & "<br>"
'Response.Write "x = "& x &" UgadagDato: "& varTjDato_ugedag &" Uge afl: " &ugeNrAfsluttet& "# " & datepart("ww", ugeNrAfsluttet, 2, 2) &" = "& strWeek &" AND "& smilaktiv &" AND "& autogk &" AND "& ugeNrAfsluttet &" autolukvdatodato "& autolukvdatodato &"<br>"
'Response.Write "herder45"
'Response.Write day(now) &" > "& autolukvdatodato  &" AND "& DatePart("yyyy", varTjDato_ugedag) &" = "& year(now) & " AND "& DatePart("m", varTjDato_ugedag) &" < "& month(now) & "<br>"
'Response.Write "aktdata(iRowLoop, 16): " & aktdata(iRowLoop, 16) & "<br>"

'response.write "<br>erIndlast: " & erIndlast
'end if 


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
    

    


    '*** Er aktivitet lå st pga tidslås ****'
    if len(trim(job_tidslassSt)) <> 0 AND len(trim(job_tidslassSl)) <> 0 then
    
        if cint(ignorertidslas) = 1 OR job_tidslass = 0 OR (job_tidslass = 1 AND fieldVal = 1 AND ((formatdatetime(job_tidslassSt, 3) <= timenow AND formatdatetime(job_tidslassSl, 3) >= timenow ) OR _
	    (formatdatetime(job_tidslassSt, 3) <= timenow AND formatdatetime(job_tidslassSl, 3) <= formatdatetime(job_tidslassSt, 3)) OR _
	    (formatdatetime(job_tidslassSl, 3) >= timenow AND formatdatetime(job_tidslassSl, 3) <= formatdatetime(job_tidslassSt, 3)))) then
        'Åben
        
            maxl = maxl
	        fmbgcol = fmbgcol
	        fmborcl =  fmborcl
            'ugeafsluttetTxt = ""
    	    
	    else
	    'Lukket
    	
	        'Response.Write "Kaj 2 lukket"
            
            maxl = 0
	        fmbgcol = "#EfF3FF" 
	        fmborcl = "1px #cccccc dashed"
        
            ugeafsluttet = 1
            ugeafsluttetTxt = "En ell. flere aktiviteter er lukket pga. tidslås."
        end if
    
    end if


   'ugeafsluttetTxt = ugeafsluttetTxt & "<br>aktBudgettjkOn: " & aktBudgettjkOn & " resforecastMedOverskreddet: "& resforecastMedOverskreddet & " lto: " & lto & " level: "& level

    '''** Er forecast på aktivitet overskreddet *****'
    if (cint(aktBudgettjkOn) = 1 AND (cint(resforecastMedOverskreddet) = 1 OR cint(resforecastMedOverskreddet) = 2) AND (job_fakturerbar = 1 OR (job_fakturerbar = 90 AND lto = "mmmi"))) then  'level > 1 fjernet 20160104 == Vis amme indstillinger for alle
    '** Kun fakturerbare aktiviteter / WWF fravær / + E1 UNIK og MMMI
    '** Admin må gerne indtaste
            

            '*** Overksreddet og der må ikke tastes ***
            if cint(akt_maksforecast_treg) = 1 then

                    maxl = 0    'Kan ikke tastes
                    fmbgcol = "#CCCCCC" '"#F7F7F7" 'lysgraa
	                fmborcl = "1px #999999 solid"


            else

                '** Overskreddet Ok at taste. Uden budget. Der må ikke tastes.
                if cint(resforecastMedOverskreddet) = 1 then 
                    '**** 1 Forecast / timebudget overskreddet
                    '**** 2 Der er ikke angivet buget. Kun advarsel. Dvs gråfelter og lyserødmarkering
                    maxl = maxl '0

                    'select case lto
                    'case "wwf", "xintranet - local"
                    fmbgcol = "#ffdfdf" 'lyserød
	                'case else
                    'fmbgcol = "#FFFFFF" 
                    'end select

                    fmborcl = "1px #999999 solid"

                else 'resforecastMedOverskreddet kode 2
    
                    select case lto '** Lukket
                    case "mmmi", "sdutek"
                    maxl = maxl 'Må lleigevel taste i gråfelter (lav flueben i kontrolpanel)
                    case else
                    'case "wwf", "intranet - local"
                    maxl = 0    'Kan ikke tastes, DIV boks bruges --> Bruge denne hvis man gerne må gå over forecast på aktiviteter hvor der ER angivet forecast.
                    end select

                    fmbgcol = "#CCCCCC" '"#F7F7F7" 'lysgraa
	                fmborcl = "1px #999999 solid"
                
                end if

	       
           end if

        
            ugeafsluttet = 1
            ugeafsluttetTxt = "<b>Dit forecast</b> (kode: "& resforecastMedOverskreddet &") er overskreddet / ikke angivet, p&aring; en eller flere af de viste aktiviteter. <br>Kontakt den jobansvarlige for at &aelig;ndre forecast. (timebudget)"
            
            'resFc_skriveret_bgColor = "#FFC0CB"

    else
            

            'resFc_skriveret_bgColor = "#FFFFFF"
            maxl = maxl
	        fmbgcol = fmbgcol
	        fmborcl =  fmborcl
            'ugeafsluttetTxt = ""

    end if

    
    '**** Er periode lukket via lønkørsel **''
    call lonKorsel_lukketPer(varTjDato_ugedag, job_internt) 'Hr -2 job bliver kun lukket ifh.- lønperiode

    'Response.write "lonKorsel_lukketIO = " & lonKorsel_lukketIO 

    if ( ((datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 )) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", varTjDato_ugedag, 2, 2) = year(now) AND DatePart("m", varTjDato_ugedag, 2, 2) < month(now)) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", varTjDato_ugedag, 2, 2) < year(now) AND DatePart("m", varTjDato_ugedag, 2, 2) = 12)) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", varTjDato_ugedag, 2, 2) < year(now) AND DatePart("m", varTjDato_ugedag, 2, 2) <> 12) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", varTjDato_ugedag, 2, 2) > 1))) OR _
    cint(lonKorsel_lukketIO) = 1 then

        
        '*** Lønperiode afsluttet ***'
        '*** Uge afsluttet via Smiley Ordning og autogodkend slået til i kontrolpanel **'
        '*** hvis admin level 1 kan timer stadigvæk redigeres indtil der oprettes faktura ***'
	    if cint(level) <> 1 then
	        maxl = 0
	        fmbgcol = "#F7F7F7"
	        fmborcl = "1px #999999 dashed"
	    else
	        maxl = maxl
	        fmbgcol = fmbgcol
	        fmborcl = "1px #999999 dashed" 'fmborcl
	    end if
    	
	    ugeafsluttet = 1
        ugeafsluttetTxt = "Uge er afsluttet via smiley, periode er afsluttet via kontrolpanel ell. l&oslash;nk&oslash;rsel / godkendt af teamleder."
    	
     end if
            
                

                 '*** Ferie og Sygdom skal ikke kunne indtastes på dage uden normtid ***'
                 call normtimerPer(medid, varTjDato_ugedag, 0, 0)

                 'Response.Write "nTimPer " & nTimPer & " akttype:" &akttype & "<br>"

                 if (nTimPer = 0) AND (akttype >= 11 AND akttype <= 24) then
                 
               
                maxl = 0
	            fmbgcol = "#EfF3FF" 
	            fmborcl = "1px #cccccc dashed"
              
        
                ugeafsluttet = 1
                ugeafsluttetTxt = "Der kan ikke indtastes ferie og sygdom p&aring; en dag uden normtid."

                
                 
                 end if  




            
            
                '** ER registrering godkendt af joabans eller admin ***'
               
                if cint(ugegodkendt) = 1 AND erIndlast = 1 then
                
                    maxl = 0
                    fmbgcol = "#F7F7F7"
	                fmborcl = "1px yellowgreen dashed"
	                
	                ugeafsluttet = 1
	                ugeafsluttetTxt = "Uge er godkendt via smiley, og uge er godkendt/lukket af godkender"
    						        
                end if   
                
                
                '*** hvis afvist skal uge være åben igen ****
                if cint(ugegodkendt) = 2 then

                    maxl = 6
	                fmbgcol = "#F7F7F7"
	                fmborcl = "1px yellowgreen dashed"
	                
	                ugeafsluttet = 0
                    ugeafsluttetTxt = "Ugeseddel er afvist af godkender"

                end if
                
            
            '** Hvis Admin skal felt altid være åben, FØR tjk om der findes faktura ***''
            if cint(level) = 1 then
            maxl = 6
            end if 


            
            '**** Findes der en faktura ****'
            '**** Overskriver Smiley luk / Godkendte uger / Tidslås ****'
			if len(lastfakdato) <> 0 AND diff <= 0 AND (cint(ugegodkendt) <> 1 OR erIndlast = 0) then
					        '** Hvis fakuge > den valgte uge ***
					        if datepart("ww", lastfakdato, 2, 2) > strWeek AND datepart("yyyy", lastfakdato) => datepart("yyyy", tjekdag(4)) then
        					
        					    maxl = 0
						        fmbgcol = "#F7F7F7"
						        fmborcl = "1px #999999 dashed"
						        
                                ugeafsluttetTxt = "<b>Jobbet er lukket for timeregistrering</b><br>Der foreligger en faktura p&aring; jobbet i den p&aring;g&aelig;ldende periode. ("& lastFakDato &")"
						             'Response.Write "Fak her"
						             'Response.Write  aktdata(iRowLoop, 2) &" - "& cint(aktdata(iRowLoop, 16)) &"erIndlast: "& erIndlast &"<br>"
                
        						
        				    else
        					
        					 
        					
						        '** Hvis fakuge = den valgte uge ***
						        '** Tidspunkt sættes altid til 23:59:59 på fakturaer, fra d. 8/9-2004
						        'Response.write DatePart("y", lastfakdato) &" >=  " & DatePart("y", varTjDato_ugedag) & "<br>" 
						        if DatePart("y", lastfakdato) >= DatePart("y", varTjDato_ugedag) AND DatePart("yyyy", lastfakdato) >= DatePart("yyyy", varTjDato_ugedag) then
        					
        							
							        maxl = 0
							        fmbgcol = "#F7F7F7"
							        fmborcl = "1px #999999 dashed"
							        ugeafsluttetTxt = "<b>Jobbet er lukket for timeregistrering</b><br>Der foreligger en faktura p&aring; jobbet i den p&aring;g&aelig;ldende periode. ("& lastFakDato &")"
							         
        						
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
			

             '**** Kun se / IKKE rettigheder ***********************'
                if (jobid = 196 AND lto = "wwf" AND level > 1) OR origin <> 0 then 
            
                        '** Kun Se / Ikke opdatere == FRA TimeTag, Mobil og andre eksterne  **'
                        maxl = 0
	                    fmbgcol = "#FFFFFF" '"#F7F7F7" 
	                    fmborcl = "1px #999999 solid"
        
                       ugeafsluttetTxt = "<span style='background-color:#FFFFFF;'>Du har IKKE skrive rettigheder p&aring; en eller flere af aktiviteterne p&aring; dette job."_
                       &"<br>Timerne er indl&aelig;st via <b>TimeTag, TimeOut mobile eller Excel upload</b>, og kan v&aelig;re en sum af flere registreringer.<br>"_
                       &"Du kan <b>redigere indtastningerne via ugesedlen</b>.</span><br><br>"
            


                else

                        maxl = maxl
	                    fmbgcol = fmbgcol
	                    fmborcl =  fmborcl

                end if





            if maxl = 6 then
                select case lto
                case "mmmi", "unik", "xintranet - local"
                maxl = 9
                case else
                'maxl = maxl
                maxl = 9
                end select 
            else
            maxl = maxl
            end if

                'Feltbredde på timereg. felter
                select case lto
                case "mmmi", "unik", "xintranet - local"
                tfeltwth = 53
                case else
                tfeltwth = 45
                end select 
			
    'response.Write "fmbgcol: "& fmbgcol
			 
	
end function


public jLuft 
sub jobbeskrivelse_stdata

if cint(intHR) <> 1 then

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

else

jbeskWzb = "hidden"
jbeskDsp = "none"
jLuft = ""

end if

%>

  <div id="div_det_<%=aktdata(iRowLoop, 4)%>"" style="position:relative; background-color:#ffffff; padding:0px 3px 3px 3px; width:745px; border:0px #CCCCCC solid; visibility:<%=jbeskWzb%>; display:<%=jbeskDsp%>; z-index:2000">
   
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

                                                                 strSQLjl = "SELECT sum(timer) AS timer FROM timer WHERE tjobnr = '"& aktdata(iRowLoop, 31) &"' AND tdato BETWEEN '"& dtnowLowSQL &"' AND '"& dtnowHighSQL &"'"_
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
                                                     <br />


                                                            <%
                                                            mTypStDato = aktdata(iRowLoop, 45) ') & "/"& month(aktdata(iRowLoop, 45)) &"/"& day(aktdata(iRowLoop, 45))
                                                            
                                                            call bdgmtypon_fn()
                                                            if cint(bdgmtypon_val) = 1 then 
                                                            mTypStDato = aktdata(iRowLoop, 45)
                                                            mTypSlDato = dateadd("d", -1, now)
                                                            periodeFordeling = 0
                                                            call timer_fordeling_medarb_typer(aktdata(iRowLoop, 31), periodeFordeling, 100, mTypStDato, mTypSlDato, aktdata(iRowLoop, 4))
                                                                
                                                                else
                                                            mTypSlDato = now 'year(now) & "/"& month(now) &"/"& day(now)
                                                            periodeFordeling = 1
                                                            call timer_fordeling_medarb_typer(aktdata(iRowLoop, 31), periodeFordeling, 100, mTypStDato, mTypSlDato, aktdata(iRowLoop, 4))
                                                            end if   

                                                          
                                                            %>
                                                           
                                                          



                                                <br />
                                              
                                                <b>Start & slutdato:</b><br /> <%=formatdatetime(aktdata(iRowLoop, 45), 1) %> - <%=formatdatetime(aktdata(iRowLoop, 46), 1) %><br />	


                                                <br />

                                                
                                                <b>Job og kunde ansv.:</b><br />
                                               <%call meStamdata(aktdata(iRowLoop, 43)) %>
                                               <%=tsa_txt_007 %>: <%=meTxt %>

                                               <br />
                                               <%call meStamdata(aktdata(iRowLoop, 44)) %>
                                               <%=tsa_txt_010 %>: <%=meTxt %>

                                                  <br />
                                               <%call meStamdata(aktdata(iRowLoop, 50)) %>
                                               <%=tsa_txt_011 %>: <%=meTxt %>

                                               <%if len(trim(aktdata(iRowLoop, 27))) <> 0 AND aktdata(iRowLoop, 27) <> "0" then %>
                                               <br /><br /><b>Kontaktperson hos kunde:</b><br />
                                               <%
                                               kkpers = ""
                                               strSQLkk = "SELECT navn, email, dirtlf, mobiltlf FROM kontaktpers WHERE id = "& aktdata(iRowLoop, 27) 
                                               oRec6.open strSQLkk, oCOnn, 3
                                               if not oRec6.EOF then

                                               kkpers = oRec6("navn") 

                                               if len(trim(oRec6("email"))) <> 0 then
                                               kkpers = kkpers & ", <a href='mailto:"&oRec6("email")&"' class=rmenu>"& oRec6("email")&"</a>"
                                               end if

                                               if len(trim(oRec6("dirtlf"))) <> 0 then
                                               kkpers = kkpers & "<br>Dir.: "& oRec6("dirtlf")
                                               end if

                                               if len(trim(oRec6("dirtlf"))) <> 0 then
                                               kkpers = kkpers & "<br>M.: "& oRec6("mobiltlf")
                                               end if

                                               

                                               end if
                                               oRec6.close

                                               Response.write kkpers
                                               end if
                                               %>

                                                  </td>
                                               <td style="border-right:1px #cccccc solid;">&nbsp;</td>
                                               <td valign=top class=lille style="padding:5px 5px 5px 10px; width:250px;">
                                               

      
		
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
		                                        <b><%=tsa_txt_015 %>:</b><br /> <%=jobtype%><br />		
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

                                           

                                                 <%if level <= 2 OR level = 6 then %>	
     
                                                <!--Budget -->
                                                            <br /><br /><b><%=tsa_txt_016 %></b>
                    
				                                            <br />Timer: <%=formatnumber(aktdata(iRowLoop, 34), 2)%> t.
                                                            <br />Bruttoomsætning: <%=formatnumber(aktdata(iRowLoop, 48), 2) &" "& basisValISO%> 

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
							
							                                        '* finder antal real timer på job ***
							                                        strSQL2 = "SELECT timer AS sumtimer, (kostpris * timer * kurs / 100) AS kostpris, timepris FROM timer WHERE tjobnr = '"& aktdata(iRowLoop, 31) &"' AND ("& aty_sql_realhours &") ORDER BY timer"
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

                                                                    
                                                                    fakturerbareTimer = 0
                                                                   '* finder antal fakturerbare timer på job ***
							                                       strSQL2 = "SELECT sum(timer) AS realfakbare FROM timer WHERE Tjobnr = '" & aktdata(iRowLoop, 31) &"' AND tfaktim = 1"
	                                                               oRec2.open strSQL2, oConn, 3
							                                        if not oRec2.EOF then
                                                                    
                                                                    if len(oRec2("realfakbare")) <> 0 then
							                                        fakturerbareTimer = oRec2("realfakbare")
                                                                    else
                                                                fakturerbareTimer = 0
                                                                        end if
                            
                                                                    end if 
							                                        oRec2.close
							
							                                        %>
					
					                                        Timer: <%=formatnumber(timerbrugtthis, 2)%> t. <br />
                                                            <span style="color:#999999;">Heraf fakturerbare: <%=formatnumber(fakturerbareTimer, 2) %> t.</span> <br />
                                                            <br />
                    
                                                            Omsætning: <%=formatnumber(oms) &" "& basisValISO%> <br />
                                                            Intern kostpris: <%=formatnumber(internKost) &" "&basisValISO %> 
					

                    
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
                                                        <a href="#" onclick="Javascript:window.open('../timereg/milepale.asp?menu=job&func=opr&jid=<%=aktdata(iRowLoop, 4)%>', '', 'width=650,height=400,resizable=yes,scrollbars=yes')" class=rmenu>Opret milepæl >></a><br />
                                                        <a href="#" onclick="Javascript:window.open('../timereg/milepale.asp?menu=job&func=opr&jid=<%=aktdata(iRowLoop, 4)%>&type=1', '', 'width=650,height=400,resizable=yes,scrollbars=yes')" class=rmenu>Opret bet. termin >></a>
				                                        

                                                       
				                                        <%end if %>


                                                        <%if level <= 7 then %>
                                                           <br /><a href="#" onclick="Javascript:window.open('../timereg/tilknytprojektgrupper.asp?id=<%=aktdata(iRowLoop, 4)%>&medid=<%=usemrn %>', '', 'width=650,height=500,resizable=yes,scrollbars=yes')" class=rmenu>Tilføj projektgruppe til job >></a>
				                                        <%end if %>

                                                        <%if level <= 2 OR level = 6 then %>
                                                        <br /><a href="../timereg/erp_opr_faktura_fs.asp?visjobogaftaler=1&visminihistorik=1&FM_kunde=<%=aktdata(iRowLoop, 42)%>&FM_job=<%=aktdata(iRowLoop, 4)%>&FM_aftale=0&reset=1&FM_usedatokri=&1FM_start_dag=<%=day(tjekdag(1))%>&FM_start_mrd=<%=month(tjekdag(1))%>&FM_start_aar=<%=year(tjekdag(1))%>&FM_slut_dag=<%=day(tjekdag(7))%>&FM_slut_mrd=<%=month(tjekdag(7))%>&FM_slut_aar=<%=year(tjekdag(7))%>" target="_blank" class=rmenu>Opret faktura >> </a>
                                                        <%end if %>


                
				
				                                        <%if (level <= 7) then %>
				
                                                        <br /><br />		
				                                        <a href="materialer_indtast.asp?id=<%=aktdata(iRowLoop, 4)%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aktdata(iRowLoop, 51)%>" target="_blank" class=rmenu><img src="../ill/ikon_materiale.png" border="0" /><%=tsa_txt_021 %> >> </a>
				                                        <br />
				                                        <%
                                                         '*** Dine **'
				                                         'antalmatreg = 0
				                                         'fordeling = 0
				                                         'strSQLmr = "SELECT COUNT(id) AS fordeling, SUM(matantal) AS matantal FROM materiale_forbrug WHERE jobid = "& aktdata(iRowLoop, 4) & " AND usrid = "& usemrn & " GROUP BY jobid, usrid"
				                                         ''Response.Write strSQLmr 
				                                         ''Response.flush
				 
				 
				                                         'oRec4.open strSQLmr, oConn, 3
				                                         'if not oRec4.EOF then
				                                         'fordeling = oRec4("fordeling")
				                                         'antalmatreg = oRec4("matantal")
				                                         'end if
				                                         'oRec4.close
				                                         %>


                                                          <%'*** total **'
				                                         'antalmatregTot = 0
				                                         'fordelingTot = 0
				                                         'strSQLmr = "SELECT COUNT(id) AS fordeling, SUM(matantal) AS matantal FROM materiale_forbrug WHERE jobid = "& aktdata(iRowLoop, 4) & " GROUP BY jobid, usrid"
				                                         ''Response.Write strSQLmr 
				                                         ''Response.flush
				 
				 
				                                         'oRec4.open strSQLmr, oConn, 3
				                                         'if not oRec4.EOF then
				                                         'fordelingTot = oRec4("fordeling")
				                                         'antalmatregTot = oRec4("matantal")
				                                         'end if
				                                         'oRec4.close

                                                              %>
                                                              <div style="width:280px; height:200px; padding:5px; border:1px #CCCCCC solid; overflow-y:auto;">

                                                              <table cellpadding="2" cellspacing="0" border="0" border="0" width="100%" >
                                                                <%
                                                              
                                                                    'strIndLaestAllerede = "<b>Indlæste materialer / udlæg:</b><br>"
                                                                  
                                                               matregSum = 0
                                                               strSQLmatreg = "SELECT m.matnavn, m.matantal, m.matsalgspris, m.valuta, matenhed, v.valutakode, m.kurs, "_
                                                                &" forbrugsdato, m.editor, me.init, a.navn AS aktnavn, a.fase FROM materiale_forbrug AS m "_
                                                                &"LEFT JOIN aktiviteter AS a ON (a.id = m.aktid) "_
                                                                &"LEFT JOIN valutaer AS v ON (v.id = m.valuta) "_
                                                                &"LEFT JOIN medarbejdere AS me ON (me.mid = m.usrid) WHERE m.jobid = "& aktdata(iRowLoop, 4) & " ORDER BY a.fase, a.navn, forbrugsdato" 
                
                                                               'response.write strSQLmatreg
                                                               'response.flush
                                                                m = 0
                                                                lastFase = ""
                                                                latsAkt = ""
                                                                oRec5.open strSQLmatreg, oConn, 3
                                                                while not oRec5.EOF
                                                                    
                                                                    showluftbeforeAkt = 1

                                                                select case right(m, 1)
                                                                case 0,2,4,6,8
                                                                    bgmattr = "#EFF3FF" 
                                                                case else
                                                                    bgmattr = "#ffffff" 
                                                                end select


                                                                    if len(trim(oRec5("fase"))) <> 0 then
                                                                        
                                                                    if lcase(lastFase) <> lcase(oRec5("fase")) then
                                                                      strIndLaestAllerede = strIndLaestAllerede &"<tr style=""background-color:#ffffff;""><td colspan=6 class=lille>&nbsp;</td></tr>"
                                                                     strIndLaestAllerede = strIndLaestAllerede &"<tr style=""background-color:#D6Dff5;""><td colspan=6 class=lille>fase: <b>"& oRec5("fase") &"</b></td></tr>"
                                                                       showluftbeforeAkt = 0
                                                                    end if

                                                                     lastFase = oRec5("fase")

                                                                    end if

                                                                     if lastAkt <> oRec5("aktnavn") AND len(trim(oRec5("aktnavn"))) <> 0 then
                                                                        
                                                                        if cint(showluftbeforeAkt) = 1 then
                                                                        strIndLaestAllerede = strIndLaestAllerede &"<tr style=""background-color:#ffffff;""><td colspan=6 class=lille>&nbsp;</td></tr>"
                                                                        end if

                                                                     strIndLaestAllerede = strIndLaestAllerede &"<tr style=""background-color:#ffffff;""><td colspan=6 class=lille><b>"& oRec5("aktnavn") &"</b></td></tr>"
                                                                   
                                                                     lastAkt = oRec5("aktnavn")    

                                                                     end if


                                                                    call beregnValuta(oRec5("matsalgspris"),oRec5("kurs"),100)
                                                                    matsalgspris = formatnumber(valBelobBeregnet)
                                                                    matregSum = matregSum + (matsalgspris/1)

                                                                    strIndLaestAllerede = strIndLaestAllerede &"<tr style=""background-color:"& bgmattr &";"">"_
                                                                    &"<td class=lille>"& replace(left(formatdatetime(oRec5("forbrugsdato"), 2), 5), "-", ".") &"</td>"_
                                                                    &"<td class=lille>"& oRec5("matantal") &" "& oRec5("matenhed") &"</td>"_
                                                                    &"<td class=lille>"& left(oRec5("matnavn"), 20) &"</td>"_
                                                                    &"<td class=lille align=right>"& formatnumber(oRec5("matsalgspris"), 2) &" "& oRec5("valutakode") &"</td>"_
                                                                    &"<td align=right class=lille style=""color:#999999;"">["& oRec5("init") &"]</td></tr>"


                                                                   
                                                                   

                                                                m = m + 1
                                                                oRec5.movenext
                                                                wend
                                                                oRec5.close

                                                              '*** ÆØÅ **'
                                                              call jq_format(strIndLaestAllerede)
                                                              strIndLaestAllerede = jq_formatTxt
                                                              response.write strIndLaestAllerede


                                                            

				                                         %>
                                                        </table>
                                                    </div>
				
                                                        <%call basisValutaFN() %>

				                                        <%=tsa_txt_326 %> <b><%=m %></b>, beløb: <%=formatnumber(matregSum, 2) & " "& basisValISO %><br />
                                                        <!--<a href="materiale_stat.asp?id=0&hidemenu=1&FM_job=<%=aktdata(iRowLoop, 4)%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aktdata(iRowLoop, 51)%>&FM_visprjob_ell_sum=0" target="_blank" class=rmenu><%=m%> </a><br /> <%=tsa_txt_327 %> <b><%=fordelingTot%> (<%=fordeling%>)</b> <%=tsa_txt_328 %>.-->
				
				                                        <%end if %>
				
				
			                                             <br />
                                                         <b>Afslut og notificer</b><br />
                                                         <%'jobstatus IKKE tilbud
                                                         if aktdata(iRowLoop, 33) <> 3 then

                                                         js_SEL0 = ""
                                                         js_SEL1 = ""
                                                         js_SEL2 = ""
                                                         js_SEL4 = ""
                                                         jobstatus_val = 1

                                                         select case aktdata(iRowLoop, 33)
                                                         case 0
                                                         js_SEL0 = "SELECTED"
                                                            jobstatus_val = 0
                                                         case 1
                                                         js_SEL1 = "SELECTED"
                                                             jobstatus_val = 1
                                                         case 2
                                                         js_SEL2 = "SELECTED"
                                                             jobstatus_val = 2
                                                         case 4 
                                                         js_SEL4 = "SELECTED"
                                                             jobstatus_val = 4
                                                         end select%>
                                                                                                        
                                                         Job status: <select name="FM_lukjobstatus" id="FM_lukjobstatus_<%=aktdata(iRowLoop, 4)%>" style="font-size:9px; width:100px;">
                                                         <option value="1" <%=js_SEL1 %>>Aktiv</option>
                                                         <option value="2" <%=js_SEL2 %>>Passiv / Til fakturering</option>
                                                         <option value="0" <%=js_SEL0 %>>Lukket</option>
                                                         <option value="4" <%=js_SEL4 %>>Gennemsyn</option>
                                                         </select>
                                                         <br /><br />
                                                 
			                                             <!--<input type="checkbox" value="<%=aktdata(iRowLoop, 4)%>" name="FM_lukjob" id="FM_lukjob" class="FM_lukjob" />--><%=tsa_txt_375 %><br />
                                                         <input type="button" value=" Send / opdater >> " id="FM_lukjob_<%=aktdata(iRowLoop, 4)%>" class="FM_lukjob" style="font-size:9px;" />
                                                         
                                                          <%end if %>

            
                                                     &nbsp;

                                                        </td>
                                                        </tr>

                 

                                                         <!--- job besk --->

                                                         <%if (len(trim(aktdata(iRowLoop, 41))) <> 0 OR len(trim(aktdata(iRowLoop, 40))) <> 0) then 
	 
	                                                     select case lto
	                                                     case "dencker", "intranet - local", "jttek", "synergi1"
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

                                                    '*** Sletter indtastninger på den valgte uge på et valge job,
                            		                '*** så man ikke risikerer at have indtastninger stående i et
                            		                '*** tidsrum der ved redigering ikke længere er i det indstastede tidsinterval ***'
                            		                '*** Kun nat, dat etc og E1 typen. IKKE på grundregistreringen ***'
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
	                                                
                                                    strSQLdelweek = "DELETE FROM timer WHERE tjobnr = '"& thisJobnr &"' AND "_
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
			                         
			                         
			                                        '*** Nulstiller alle DAG / NAT / E1 typer så der ikke kan indtastes timer på disse typer
                            			            '*** Selvom de evt. er åbne og der tastes timer ind i felterne ***'
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
			        
			                         'response.Write "y: nyt arr gennemløb: " & y & "<br>"
			                         'Response.flush
			                         '***** ER klokkeslet Sttid og Sltid benyttet eller timeangivelse ****'
			                         '*** tjekker for slet tSttid(y) <> 0 ***'
                                     if cint(visTimerelTid) <> 0 AND len(trim(tSttid(y))) <> 0 then
                                     
                                     if len(trim(tSttid(y))) <> 1 then ' == Slet
                                     
                                     
                                     'Response.Write "<br><b>tSttid(y):</b>" & tSttid(y) & "<br>"
                            			
                            			           
                            		                
                            				    
                                                    '**** Tjekker for tidslås ***'
                                                    '**** Skal timer lægges på en efterfølgende akt? ***' 
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
                            				        
                            				        '** Finder grund aktivitet. Skal være = 1 fakturerbar.
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
                                                            
                                                            
                                                                     '*** Forlænger Array strings ***'
        						                                     '** Hvis der er fundet en aktivitet forlænges j arr.
        						                                     '** En indtasting via tidslås går forud for en manuel indtastning direkte i feltet ***'
                                                                    if instr(jarrforl, ",#"& aktids(j) &"#") = 0 then
                                                                    
                                                                    
                                                                    
                                                                    '*** Skal finde op til 8 aktiviteter ***'
                                                                    '** 1+2 Man-Fre Dag/Man-Tor Nat aktiviteter på medarbejder (type:Dag+Nat)
                                                                    '** 3+4 Dag/Nat aktiviteter på kunde (type:E1) 
                                                                    '** 5+6 Mandag morgen akt. aktivitet på kunde (type:Nat+E1)
                                                                    '** 7+8 Lørdag morgen akt. aktivitet på kunde (type:Nat+E1)
                                                                    '** 9+10 Fredag aften + Søndag aften (type:Nat+E1)
                                                                    '** 11+12 Weekend Dag medarb + kunde (type:Weekend+E1)
                                                                    '** 13+14 Weekend Nat medarb + kunde (type:Weekend Nat+E1)
                                                                    
                                                                   
                                                                    cta = 7
                                                                   
                                                                    
                                                                    newArrSizeAkt = ubound(aktids)+100 
                                                                    Redim preserve jobids(newArrSizeAkt) 
                                                                    Redim preserve aktids(newArrSizeAkt)
                                                                    
                                                                    
                                                                    
                                                                    jarrforl = jarrforl & ",#"& aktids(j) & "#"
                                                                    
                                                                    '** Gør plads til 14 akt. (se ovenfor) ***'
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
                                                                            '** 3 for at få 7 til at blive næste 10'er
                                                                            '** dvs 17 bliver til 20 for at få første felt i næste række
                                                                            
                                                                            'Response.Write "high_fetnr+y2:" & high_fetnr &" - "& y2&"<br>"
                                                                            'Response.Write "feltnr(y2):" &  feltnr(y2) & " y2:"& y2 &"<br>"
                                                                            '*** ialtcounter **'
                                                                            'iic = iic + 1
                                                                            next
                                                                            
                                                                            high_fetnr = high_fetnr + 10
                                                                    
                                                                    next
                                                                    
                                                                end if 'jarrforl
                                                                
                                                                
                                                                
                                                                
                                                                '*** Kig efter ny akt. der er åben iflg. tidslås **'
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
                                                                    
                                                                    
                                                                    
                                                                    'Response.Write "<br><u>Undersøger: "& aktids(j) &" Finder:"& oRec5("id") & "</u><br><br>"
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
                                    						        
        						                                    
        						                                    '*** Redigere tidspunkter efter tidslås ***'
        						                                    
        						                                    '* Hvis tidslås A = 1
        						                                    '* Hvis tidslås A for dagen = 1
        						                                    
        						                                    
        						                                   
        						                                    '*** Trimmer grundreg **'
        						                                    if cint(gnlob) = 1 then
        						                                    
        						                                    '*** Tjekker Dato format her ***'
        						                                    
        						                                    'KODE
        						                                    
        						                                    '****'
        						                                    gRegSt = tSttid(y)
        						                                    gRegSl = tSltid(y)
        						                                    
        						                                    
        						                                    
        						                                            
        						                                           
        						                                    end if 'gnlob   
        						                                    
        						                                    
        						                                    
        						                                    if request("minindtast_on") = "1" then
        						                                    '**** Min 7,4 timer på medarb / 8 timer på kunde ****'
					                                                '**** Det skal være samlet for dagen             ****'
					                                                '**** hvilken akt. skal den placere det på?      ****'
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
        						                                    
        						                                   
        						                                   
        						                                     
        						                                    '*** Er tidslås slåettil på dagen ***'
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
                                                                                    
                                                                                    for y2 = gnlob_st_high_fetnr + 1 to gnlob_st_high_fetnr + 7 '*** Lægges 7 til for 7 dage ***'
                                                                                    
                                                                                           
                                                                                            
                                                                                            aktids(newArrSizeAkt-100+icc+1) = oRec5("id") '+gnlob
                                                                                            
                                                                                            if right(feltnr(y2), 1) = right(feltnr(y),1) then
                                                                                            
                                                                                             
                                                                                                'y2 = y2 + (cint(right(feltnr(y),1)) * 7)
                                                                                            
                                                                                            
                                                                                                     '*** Trimmer de funde aktiviteter **'
        						                                                                     'Response.Write "<b>Aktid:</b>"& oRec5("id") &"/"& aktids(newArrSizeAkt-7+gnlob) &"y2 val: " & y2 &"/"& newArrSizeAkt-7+gnlob &" gRegSt: "& gRegSt & " til "& gRegSl &" - tidslaas_st_B: " & tidslaas_st_B & " tidslaas_sl_B: " & tidslaas_sl_B & "cint(right(feltnr(y2), 1)) >= cint(weekdayThis)"& cint(right(feltnr(y2), 1)) &" >= "& cint(weekdayThis) &"<br>"
                                        						                                     
                						                                                            '*** Er sttid større end start tidslås og er sttid mindre end tidslås slut.
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
        						                                                                        'Response.Write "Bruger tidslås start: " & left(tidslaas_st_B, 5)
        						                                                                    end if
                                						                                            
                                						                                            
        						                                                                    '* Hvis angivet SLUT tidspunkt er MINDRE end tidslås slut A
        						                                                                    '>>
        						                                                                    '** Er slut tid næste morgen, dvs efter 24:00 **'
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
        						                                                                    '* Hvis angivet SLUT tidspunkt er STØRRE end tidslås slut A
        						                                                                    '>>
                                						                                    
        						                                                                    idag = formatdatetime(now, 2)
        						                                                                    minutberegning = datediff("n", idag &" "& tSttid(y2), idag &" "& tSltid(y2), 2, 2) 
                                						                                            
                                						                                            
                                						                                            
                                						                                            
        						                                                                    'Response.Write "minutberegning tidslås st: "& tidslaas_st_B &" tidslås sl: "& tidslaas_sl_B &" gregST: "& cDate(gRegSt) &" gRegSl: "& cDate(gRegSl) &" yST: "& cDate(tSttid(y2)) &" ySL "&tSltid(y2)&" typ: "& B_type &": " & minutberegning & "<br>"
        						                                                                    'Response.flush
                                						                                            
        						                                                                    '*** Tjekker om der skal indlæses på akt selvom minutbe. er negativt **'
        						                                                                    '*** AND cDate(tSltid(y2)) < cDate(gRegSl)
        						                                                                    if (minutberegning < 0 AND cDate(tSttid(y2)) > cDate(gRegSt) AND cDate(gRegSt) < cDate(gRegSl))_
        						                                                                    OR (minutberegning < 0 AND cDate(tidslaas_sl_B) < cDate(tidslaas_st_B) AND cDate(tidslaas_sl_B) <= cDate(gRegSl) AND cDate(tidslaas_st_B) <= cDate(tSltid(y2)))_
        						                                                                    OR (minutberegning < 0 AND cDate(tidslaas_st_B) < cDate(tidslaas_sl_B) AND cDate(gRegSt) >= cDate(tidslaas_sl_B) AND cDate(gRegSl) <= cDate(tidslaas_st_B))_
        						                                                                    
        						                                                                    then
        						                                                                    
        						                                                                    'Response.Write "minutbereg. negativt, indtastning IKKE godkendt<br><br>"
        						                                                                    
        						                                                                    
        						                                                                    tSttid(y2) = ""
        						                                                                    tSltid(y2) = ""
        						                                                                    end if
                                                                                                    
                                                                                                    
                                                                                                    '** Korrigere dato (lør morgen / Mandag morgen mv.)
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
                                                                                    
                                                                   
                                                                    end if 'tidslås slået til
                                                                    
                                                                    
                                                                    icc = icc + 1  
        						                                    gnlob = gnlob + 1
                                                                    oRec5.movenext
                                                                    wend 
                                                                    oRec5.close
                                                                    
                                                 
                                                                    
                                                                    
                                                                   
                                                                    
                                                            '*** err hvis der ikke findes en akt med en dækkende tidslås ***'
                                                            if cint(aktfundet) = 0 then
                                                             tidslaasErr = 1
                                                            end if
                                                                        
                                                                        
                                                                        
                                                                
                                                        
                                                else
                                                
                                                tidslaasErr = 1
                                                
                                                end if 'tidslaasDayOn
                                                
                    				        
                                            end if 'tidslåas
                                                    
                                                    
                                                    
                                                    
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




function timerIndlaesPeriodeLukket(medarbejderid, regdato, intjobid)
       

        'Response.Write "medarbejderid: " & medarbejderid & " regdato: "& regdato &" intjobid: "& intjobid
        'response.end



            firstWeekDay = weekday(regdato, 2)
                
                select case firstWeekDay
                case 1
                stTil = 0
                slTil = 6
                case 2
                stTil = 1
                slTil = 5
                case 3
                stTil = 2
                slTil = 4
                case 4
                stTil = 3
                slTil = 3
                case 5
                stTil = 4
                slTil = 2
                case 6
                stTil = 5
                slTil = 1
                case 7
                stTil = 6
                slTil = 0
                end select
                
                stwDayDato = dateadd("d", -stTil, regDato)
                slwDayDato = dateadd("d", slTil, regDato)
                
                regdatoStSQL = year(stwDayDato) &"/"& month(stwDayDato) &"/"& day(stwDayDato)
                regdatoSlSQL = year(slwDayDato) &"/"& month(slwDayDato) &"/"& day(slwDayDato)

            
                'Response.Write "firstWeekDay "& firstWeekDay
                
                call afsluger(medarbejderid, regdatoStSQL, regdatoSlSQL)
                call smileyAfslutSettings()
                
                
                erugeafsluttet = instr(afslUgerMedab(medarbejderid), "#"&datepart("ww", regdato,2,2)&"_"& datepart("yyyy", regdato) &"#")
                
                'Response.Write useDatoer(j) & "<br>"
                'Response.Write "erugeafsluttet --" & erugeafsluttet  &"<br>"
                'Response.Write "autogk --" & autogk  &"<br>"
                'Response.Write "smilaktiv  --" & smilaktiv   &"<br>" 
                'Response.flush
                'Response.end

                strMrd_sm = datepart("m", regdato, 2, 2)
                strAar_sm = datepart("yyyy", regdato, 2, 2)
                strWeek = datepart("ww", regdato, 2, 2)
                strAar = datepart("yyyy", regdato, 2, 2)

                if cint(SmiWeekOrMonth) = 0 then
                usePeriod = strWeek
                useYear = strAar
                else
                usePeriod = strMrd_sm
                useYear = strAar_sm
                end if

                
                call erugeAfslutte(useYear, usePeriod, medarbejderid, SmiWeekOrMonth, 0)
                
                call lonKorsel_lukketPer(regdato, -2)
              
                 
                 'if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                 if ( (( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 )) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", regdato) = year(now) AND DatePart("m", regdato) < month(now)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", regdato) < year(now) AND DatePart("m", regdato) = 12)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", regdato) < year(now) AND DatePart("m", regdato) <> 12) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", regdato) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
              
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = 0
                end if 
		        
		                
		                
		        '*** Tjekker sidste fakdato ***'
		        'if len(trim(intjobid)) <> 0 then
		        'intjobid = intjobid
		        'else
		        'intjobid = 0
		        'end if
		        
		        
		        
		        lastFakdato = "01/01/2002"
		        strSQL = "SELECT fakdato FROM fakturaer WHERE jobid = "& intjobid &" AND faktype = 0 ORDER BY fakdato DESC LIMIT 0,1"
		        
		       
		        oRec.open strSQL, oConn, 3
		        if not oRec.EOF then
		        lastFakdato = oRec("fakdato")
		        end if
		        oRec.close
		        
		        if ugeerAfsl_og_autogk_smil = 1 OR cdate(lastFakdato) >= cdate(regdato) then 'AND level > 1
		        
		        %>
			    <!--#include file="../../inc/regular/header_lysblaa_inc.asp"-->
			    <% 
			    
			    if lastFakdato = "01/01/2002" then
			    lastFakdato = "(ingen)"
			    end if
			    
			    useleftdiv = "t"
			    errortype = 117
			    call showError(errortype)
		        
		        Response.end
		        end if
                                                        
                                                        
                                                                                                         
                                                        
end function





%>





