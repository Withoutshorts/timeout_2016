

<%
'**********************************'
'**** Ets tidslås beregning ***'
'**** HUSK timereg_2006_inc.asp
'**** kode er fjernet fra den fil.
'**********************************'    


       
         '***** ER klokkeslet Sttid og Sltid benyttet eller timeangivelse ****'
         if visTimerelTid <> 0 AND len(trim(tSttid(y))) <> 0 then
			
            'Response.write len(trim(sTtid)) &" - "& sTtid &" : "& sLtid &"<br>"
            'Response.end
			
			
            if len(trim(tSttid(y))) = 1 then '** slet
			
            useTimer = 0
			
            else
			
                '** Beregner timer og min ****'
                idag = day(now)&"/"&month(now)&"/"&year(now)
                totalmin = datediff("n", idag &" "& tSttid(y), idag &" "& tSltid(y))
                call timerogminutberegning(totalmin)
                useTimer = thoursTot&"."&tminProcent 'tminTot
				
				 
                if useTimer < 0 then '** Hen over kl 24:00 **'
                useTimer = 24 + (replace(useTimer,".",","))
                useTimer = replace(useTimer,",",".")
                end if
				
				
                tSttid(y) = tSttid(y) &":00"
                tSltid(y) = tSltid(y) &":00"
				    
                        '**** Tjekker for tidslås ***'
                        '**** Skal timer lægges på en efterfølgende akt? ***' 
                        weekdayThis = weekday(useDato, 2)
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
				        
                        strSQLtidslaas = "SELECT tidslaas, tidslaas_st, "_
                        &" tidslaas_sl, "& flt &" AS tidslaasDayOn FROM aktiviteter WHERE "_
                        &" id = "& aktids(j)
                        
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
				        
				        tidslaasErr = 0
				        
                        if tidslaas <> 0 then
				        
				            'Response.Write " y1: "& tSttid(y) &" > tidslaas_st " & tidslaas_st  &" AND y2: "& tSltid(y) &" < tidslaas_sl " & tidslaas_sl & "<br>"
                            'Response.flush
                            
                            if tidslaasDayOn <> 0 then
                                    
                                    if cDate(tSttid(y)) >= cDate(tidslaas_st) AND cDate(tSltid(y)) <= cDate(tidslaas_sl) then
                                    
                                    else
                                    '*** Kig efter ny akt. der er åben iflg. tidslås **'
                                        
                                        
                                        aktfundet = 0
                                        strSQLfind = "SELECT a.id AS id, a.navn, a.job, tidslaas, tidslaas_st, "_
                                        &" tidslaas_sl, "& flt &" AS tidslaasDayOn FROM job j "_
                                        &" LEFT JOIN aktiviteter a ON ("_
                                        &" a.id <> "& aktids(j) & " AND tidslaas = 1 AND "& flt &" = 1 AND "_
                                        &" tidslaas_sl >= '"& tSltid(y) &"'"_
                                        &" AND a.job = j.id AND navn LIKE '"& left(strAktNavn, 6) &"%')"_
                                        &" WHERE j.jobnr = "& intJobnr &" AND a.job <> ''"
                                         
                                        'tidslaas_st >= '"& tSttid(y) &"'
                                        ', job j
                                        'AND a.job = j.id AND j.jobnr = "& intJobnr &"
                                        'Response.Write strSQLfind
                                        'Response.end
				                        
                                        oRec5.open strSQLfind, oConn, 3
                                        if not oRec5.EOF then
        						        
                                        'tidslaas = oRec5("tidslaas")
        						        
                                        'tidslaas_st = formatdatetime(oRec5("tidslaas_st"), 3)
                                        'tidslaas_sl = formatdatetime(oRec5("tidslaas_sl"), 3)
        						        
                                        'tidslaasDayOn = oRec5("tidslaasDayOn")
        						        
        						        aktids(j) = oRec5("id")
        						        strAktNavn = replace(oRec5("navn"), "'", "")
        						        
        						        aktfundet = 1
                                        end if
                                        oRec5.close
                                        
                                        
                                        'Response.end
                                        
                                        
                                            '*** err hvis der ikke findes en akt ***'
                                            if cint(aktfundet) = 0 then
                                            tidslaasErr = 1
                                            end if
                                    end if
                            
                            else
                            
                            tidslaasErr = 1
                            
                            end if
                            
				        
                        end if
                        
                        if tidslaasErr <> 0 then
                        %>
                        <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
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
       '*****************************************************************'        
                    
				                  



 %>
