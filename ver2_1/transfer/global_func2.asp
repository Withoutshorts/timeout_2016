




<%

      
       

        public guidjobids, guideasyids
        function projktgrpPaaJobids(pid)
        
                            '*** Sætter guiden aktive job ready ***'
					        strSQLjob = "SELECT id AS jid FROM job WHERE "_
					        &" projektgruppe1 = "& pid &" OR " _
					        &" projektgruppe2 = "& pid &" OR " _
					        &" projektgruppe3 = "& pid &" OR " _
					        &" projektgruppe4 = "& pid &" OR " _
					        &" projektgruppe5 = "& pid &" OR " _
					        &" projektgruppe6 = "& pid &" OR " _
					        &" projektgruppe7 = "& pid &" OR " _
					        &" projektgruppe8 = "& pid &" OR " _
					        &" projektgruppe9 = "& pid &" OR " _
					        &" projektgruppe10 = "& pid 
					        
					        oRec4.open strSQLjob, oConn, 3
					        while not oRec4.EOF
					        
					        guidjobids = guidjobids & oRec4("jid") &", #, "
					        
					        oRec4.movenext
					        wend
					        oRec4.close
					        
					        
					        '*** Sætter guiden aktive job Easyreg ready ***'
					        strSQLeasy = "SELECT j.id AS jid, COUNT(a.id) AS antal_aeasy FROM job AS j "_
					        &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND a.easyreg = 1) WHERE "_
					        &" j.projektgruppe1 = "& pid &" OR " _
					        &" j.projektgruppe2 = "& pid &" OR " _
					        &" j.projektgruppe3 = "& pid &" OR " _
					        &" j.projektgruppe4 = "& pid &" OR " _
					        &" j.projektgruppe5 = "& pid &" OR " _
					        &" j.projektgruppe6 = "& pid &" OR " _
					        &" j.projektgruppe7 = "& pid &" OR " _
					        &" j.projektgruppe8 = "& pid &" OR " _
					        &" j.projektgruppe9 = "& pid &" OR " _
					        &" j.projektgruppe10 = "& pid & " GROUP BY j.id" 
					        
					        'Response.Write strSQLeasy
					        'Response.flush
					        
					        oRec4.open strSQLeasy, oConn, 3
					        while not oRec4.EOF
					        
					        if oRec4("antal_aeasy") <> "0" then
					        guideasyids = guideasyids & oRec4("jid") &", #, "
					        else
					        guideasyids = guideasyids &"0, #, "
					        end if
					        
					        oRec4.movenext
					        wend
					        oRec4.close
        
        
        end function

        function setGuidenUsejob(medid, useJob, del, useEasy)
		
		'Response.Write useJob & "<hr>"
		'Response.Write useEasy & "<hr>"
		
		if len(trim(useJob)) <> 0 then
		len_useJob = len(useJob)
		left_useJob = left(useJob, (len_useJob-3))
		useJob = left_useJob
		end if
		
		'Response.Write useEasy & "<br>"
		'Response.flush
		
		if len(trim(useEasy)) <> 0 then
		len_useEasy = len(useEasy)
		left_useEasy = left(useEasy, (len_useEasy-3))
		useEasy = left_useEasy
		end if
		
		'Response.flush
		if cint(del) = 0 then
		oConn.execute("DELETE FROM timereg_usejob WHERE medarb = "& medid &"")
		end if
		
		
		j = 0
		intuseJob = Split(useJob, "#, ")
		intUseEasy = split(useEasy, "#, ")
	   	For j = 0 to Ubound(intuseJob)
	   	    
	   	    'Response.Write len(trim(intuseJob(j))) & ": " & intuseJob(j) &" Easy: "& intUseEasy(j) &"<br>"
	   	    'Response.flush
	   	    
	   	    if len(trim(intuseJob(j))) > 1 then
	   	    intuseJob(j) = trim(left(intuseJob(j), len(intuseJob(j)) - 2))
	   	    else
	   	    intuseJob(j) = 0
	   	    end if
	   	    
	   	    if len(trim(intUseEasy(j))) > 1 then
	   	    intUseEasy(j) = trim(left(intUseEasy(j), len(intUseEasy(j)) - 2))
	   	    else
	   	    intUseEasy(j) = 0
	   	    end if
	   	
	   	'Response.Write " intusejob DB: " & intuseJob(j) & "<br>"
	   	
	   	if intuseJob(j) <> 0 then
		strSQL = "INSERT INTO timereg_usejob (medarb, jobid, easyreg) VALUES ("& medid &", "& intuseJob(j) &","& intUseEasy(j) &")"
		'Response.write strSQL & "<br>"
		'Response.flush
		oConn.execute(strSQL)
		end if
		next
	    
	    'Response.end
	    
		end function


function sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	%>
	<div id="sideoverskrift" style="position:relative; left:<%=oleft%>px; top:<%=otop%>px; width:<%=owdt%>px; visibility:visible; border:0px #000000 solid; z-index:1200;">
	<!-- sideoverskrift -->
	<table cellspacing="0" cellpadding="5" border="0" width=100%>
	<tr>
	    <td style="width:48px;"><img src="../ill/<%=oimg%>" alt="" border="0"></td>
	    <td align=left style="padding-top:30px; padding-left:20px;"><h3><%=oskrift %></h3></td>
	</tr>
	</table>
	</div>	
	<%
	end function


    
    
    public chkMedarblinier, chklog, chktlffax, chkemail, chkdeak, chklukjob
	function setFakPreInd()
	
	
	if len(request("FM_chkmed")) <> 0 then
	    if request("FM_chkmed") = 1 then
	    chkMedarblinier = 1
	    else
	    chkMedarblinier = 0
	    end if
	else
	    if len(trim(request.cookies("erp")("tvmedarb"))) <> 0 then
	    chkMedarblinier = request.cookies("erp")("tvmedarb")
	    else
	    chkMedarblinier = 0
	    end if
	end if
	
	Response.Cookies("erp")("tvmedarb") = chkMedarblinier
	
	
	if len(request("FM_chklog")) <> 0 then
	    if request("FM_chklog") = 1 then
	    chklog = 1
	    else
	    chklog = 0
	    end if
	else
	    chklog = request.cookies("erp")("tvlogs")
	end if
	
	if len(request("FM_chktlffax")) <> 0 then
	    if request("FM_chktlffax") = 1 then
	    chktlffax = 1
	    else
	    chktlffax = 0
	    end if
	else
	    chktlffax = request.cookies("erp")("tvtlffax")
	end if
	
	if len(request("FM_chkemail")) <> 0 then
	    if request("FM_chkemail") = 1 then
	    chkemail = 1
	    else
	    chkemail = 0
	    end if
	else
	    chkemail = request.cookies("erp")("tvemail")
	end if
	
	'if len(request("FM_chkdeak")) <> 0 then
	'    if cint(request("FM_chkdeak")) = 1 then
	'    chkdeak = 1
	'    else
	    chkdeak = 0
	'    end if
	'else
	'    if request.cookies("erp")("deak") <> "" then
	'    chkdeak = request.cookies("erp")("deak")
	'    else
	'    chkdeak = 0
	'    end if
	'end if
	
	if len(request("FM_chklukjob")) <> 0 then
	    if cint(request("FM_chklukjob")) = 1 then
	    chklukjob = 1
	    else
	    chklukjob = 0
	    end if
	else
	    if request.cookies("erp")("lukjob") <> "" then
	    chklukjob = request.cookies("erp")("lukjob")
	    else
	    chklukjob = 0
	    end if
	end if
	
	Response.cookies("erp")("lukjob") = chklukjob 
	Response.cookies("erp")("deak") = chkdeak 
	Response.Cookies("erp")("tvemail") = chkemail
	Response.Cookies("erp")("tvtlffax") = chktlffax
	Response.Cookies("erp")("tvlogs") = chklog
	
	end function
    
    

    public  betCHK8,  betCHK7,  betCHK6,  betCHK5,  betCHK4,  betCHK3,  betCHK2,  betCHK1, betCHKu1
	function betalingsbetDage(betbetint,disa)
	
	'if cint(disa) = 1 then
	'disTxt = "DISABLED"
	'else
	'disTxt = ""
	'end if
	
	    betCHK8 = ""
	    betCHK7 = ""
	    betCHK6 = ""
	    betCHK5 = ""
	    betCHK4 = ""
	    betCHK2 = ""
	    betCHK3 = ""
	    betCHK1 = ""
	    betCHKu1 = ""
	
	select case betbetint
	    case 1
	    betCHK1 = "SELECTED"
	    case 2
	    betCHK2 = "SELECTED"
	    case 3
	    betCHK3 = "SELECTED"
	    case 4
	    betCHK4 = "SELECTED"
	    case 5
	    betCHK5 = "SELECTED"
	    case 6
	    betCHK6 = "SELECTED"
	    case 7
	    betCHK7 = "SELECTED"
	    case "-1"
	    betCHKu1 = "SELECTED"
	    case 8
	    betCHK8 = "SELECTED"
	   
	    
	    end select %>
	
		
		
		<select id="FM_betbetint" name="FM_betbetint">
		
            <option value="0">Ikke angivet</option>
            <option value="-1" <%=betCHKu1 %>>Angiver selv</option>
                <option value="1" <%=betCHK1 %>>8 dage</option>
                <option value="2" <%=betCHK2 %>>14 dage</option>
                <option value="5" <%=betCHK5 %>>21 dage</option>
                <option value="3" <%=betCHK3 %>>30 dage</option>
                 <option value="6" <%=betCHK6 %>>45 dage</option>
                <option value="4" <%=betCHK4 %>>Lbn. månd + 15 dage</option>
                <option value="8" <%=betCHK8 %>>Lbn. månd + 30 dage</option>
                <option value="7" <%=betCHK7 %>>Lbn. månd + 45 dage</option>
       
            </select>
	
	<%
	end function

public antalFak
function antalFakturaerKid(kid)
'*** kunde må kun slettes hvis der ikke findes fakturaer **'
			antalfak = 0
			strSQLfak = "SELECT COUNT(fid) AS antalfak FROM fakturaer WHERE fakadr = " & kid & " GROUP BY fakadr"
			oRec4.open strSQLfak, oConn, 3
			if not oRec4.EOF then
			
			antalfak = oRec4("antalfak")
			
			end if
			oRec4.close 
end function


public totaltimerPer, totalpausePer, totalTimerPer100
function fLonTimerPer(stDato, periode, visning, medid)

'Response.Write "her " & stDato & " Periode: " & periode



totaltimerPer = 0
totalpausePer = 0
ugeIaltFraTilTimer = 0

slutDato = dateadd("d", periode, stDato)
weekDiff = datediff("ww", stDato, slutDato, 2, 2)

if visning = 0 then '** Stempelur faneblad på timereg siden
    call akttyper2009(2)
end if


'*** Finder navne på til/fra typer **'
		                
akttyperTFh = replace(aty_sql_tilwhours, "t.tfaktim = 0", "")
akttyperTF = split(akttyperTFh, " OR t.tfaktim = ")
for tf = 0 to UBOUND(akttyperTF)

'Response.Write akttyperTF(tf) & "<br>"

if len(trim(akttyperTF(tf))) <> 0 then
call akttyper(akttyperTF(tf),0)

if tf = 0 then
akttypenavnTil = akttypenavn
else
akttypenavnTil = akttypenavnTil &", "& akttypenavn
end if
'Response.Write akttyperTF(tf) & "<br>"
end if

next


		                
akttyperTFr = replace(aty_sql_frawhours, "t.tfaktim = 0", "")
akttyperTF = split(akttyperTFr, " OR t.tfaktim = ")
for tf = 0 to UBOUND(akttyperTF)

'Response.Write akttyperTF(tf) & "<br>"

if len(trim(akttyperTF(tf))) <> 0 then
call akttyper(akttyperTF(tf),0)

if tf = 0 then
akttypenavnFra = akttypenavn
else
akttypenavnFra = akttypenavnFra & ", "& akttypenavn
end if
'Response.Write akttyperTF(tf) & "<br>"
end if

next

'*******


'Response.Write stDato & ", "& periode & "meid: " & medid & " weekdiff: "& weekdiff

	'**** login historik (denne uge/ Periode) ****
	for intcounter = 0 to periode
	
					select case intcounter
					case 0
					useSQLd = stDato
					case else
					tDat = dateadd("d", 1, useSQLd)
					useSQLd = year(tDat) &"/"& month(tDat) & "/"& day(tDat)
					end select
					
					
					
					strSQL = "SELECT l.id AS lid, l.mid AS lmid, l.login, l.logud, l.minutter, "_
					&" s.navn AS stempelurnavn, s.faktor, s.minimum, stempelurindstilling FROM login_historik l"_
					&" LEFT JOIN stempelur s ON (s.id = l.stempelurindstilling) WHERE "_
					&" l.dato = '"& useSQLd &"' AND l.mid = " & medid &""_
					&" ORDER BY l.login" 
					
					'Response.Write strSQL & "<br><br>"
					
					
					f = 0
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
						
						
						
						totaltimerPer = totaltimerPer + timerThis
						totalpausePer = totalpausePer + timerThisPause
						'Response.write oRec("lid") & ": " &  timerThis &" - "
						else
						
						totaltimerPer = totaltimerPer
						totalpausePer = totalpausePer 
						
						end if
						
						
						'timTemp = formatnumber(timerThis/60, 3)
						'timTemp_komma = split(timTemp, ",")
						
						'for f = 0 to UBOUND(timTemp_komma)
							
						'	if f = 0 then
						'	thours = timTemp_komma(f)
						'	end if
							
						'	if f = 1 then
						'	tmin = timerThis - (thours * 60)
						'	end if
							
						'next
						
						if visning = 0 then
					    
						select case intcounter
						case 0
						manMin = manMin + timerThis 
						manMinPause = manMinPause + timerThisPause 
						case 1
						tirMin = tirMin + timerThis 
						tirMinPause = tirMinPause + timerThisPause  
						case 2
						onsMin = onsMin + timerThis
						onsMinPause = onsMinPause + timerThisPause   
						case 3
						torMin = tormin + timerThis
						torMinPause = torminPause + timerThisPause   
						case 4
						freMin = freMin + timerThis
						freMinPause = freMinPause + timerThisPause   
						case 5
						lorMin = lorMin + timerThis
						lorMinPause = lorMinPause + timerThisPause  
						case 6
						sonMin = sonMin + timerThis
						sonMinPause = sonMinPause + timerThisPause   
						end select 
						
						
						end if
					    
					    'Response.write "tot:"& intcounter &" - "& totaltimer &":"& totalpause
						
						
					oRec.movenext
					wend
					oRec.close 
					
					f = 0
					
					
					
					    '*** Tillæg / Fradrag via Realtimer **'
					    if visning = 0 then
                	
		                
                		
		                tiltimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS tiltimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND ("& aty_sql_tilwhours &")) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                tiltimer = oRec2("tiltimer") * 60
		                end if
		                oRec2.close 
                		
		                if len(trim(tiltimer)) <> 0 then
		                tiltimer = tiltimer
		                else
		                tiltimer = 0
		                end if
                		
                		 
                		
                		
                		fradtimer = 0
		                strSQL2 = "SELECT sum(timer) AS fratimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND ("& aty_sql_frawhours &")) GROUP BY tmnr "
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                fradtimer = oRec2("fratimer") * 60
		                end if
		                oRec2.close 
                		
		                if len(trim(fradtimer)) <> 0 then
		                fradtimer = fradtimer
		                else
		                fradtimer = 0
		                end if
		                
		                
		                select case intcounter
						case 0
						manFraTimer = (tiltimer - fradtimer)
						case 1
						tirFraTimer = (tiltimer - fradtimer)
						case 2
						onsFraTimer = (tiltimer - fradtimer)
						case 3
						torFraTimer = (tiltimer - fradtimer)
						case 4
						freFraTimer = (tiltimer - fradtimer)
						case 5
						lorFraTimer = (tiltimer - fradtimer)
						case 6
						sonFraTimer = (tiltimer - fradtimer)
						end select 
						
		                
		                
		                '*** Fleks ***'
                		call akttyper2009prop(7)
                		aty_fleks_on = aty_on
                		aty_fleks_tf = aty_tfval
                		if cint(aty_fleks_on) = 1 then
		                '*** Fleks ****'   
                	    flekstimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS flekstimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 7) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                flekstimer = oRec2("flekstimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(flekstimer)) <> 0 then
		                flekstimer = flekstimer
		                else
		                flekstimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manflekstimer = (flekstimer)
						case 1
						tirflekstimer = (flekstimer)
						case 2
						onsflekstimer = (flekstimer)
						case 3
						torflekstimer = (flekstimer)
						case 4
						freflekstimer = (flekstimer)
						case 5
						lorflekstimer = (flekstimer)
						case 6
						sonflekstimer = (flekstimer)
						end select 
						
						
						end if
						
						
						'*** Ferie ***'
                		call akttyper2009prop(14)
                		aty_Ferie_on = aty_on
                		aty_Ferie_tf = aty_tfval
                		if cint(aty_Ferie_on) = 1 then
		                '*** Ferie ****'   
                	    Ferietimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Ferietimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 14) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Ferietimer = oRec2("Ferietimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Ferietimer)) <> 0 then
		                Ferietimer = Ferietimer
		                else
		                Ferietimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manFerietimer = (Ferietimer)
						case 1
						tirFerietimer = (Ferietimer)
						case 2
						onsFerietimer = (Ferietimer)
						case 3
						torFerietimer = (Ferietimer)
						case 4
						freFerietimer = (Ferietimer)
						case 5
						lorFerietimer = (Ferietimer)
						case 6
						sonFerietimer = (Ferietimer)
						end select 
						
						
						end if
						
						
						
						'*** Syg ***'
                		call akttyper2009prop(20)
                		aty_Syg_on = aty_on
                		aty_Syg_tf = aty_tfval
                		if cint(aty_Syg_on) = 1 then
		                '*** Syg ****'   
                	    Sygtimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Sygtimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 20) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Sygtimer = oRec2("Sygtimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Sygtimer)) <> 0 then
		                Sygtimer = Sygtimer
		                else
		                Sygtimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manSygtimer = (Sygtimer)
						case 1
						tirSygtimer = (Sygtimer)
						case 2
						onsSygtimer = (Sygtimer)
						case 3
						torSygtimer = (Sygtimer)
						case 4
						freSygtimer = (Sygtimer)
						case 5
						lorSygtimer = (Sygtimer)
						case 6
						sonSygtimer = (Sygtimer)
						end select 
						
						
						end if
						
						
						'*** BarnSyg ***'
                		call akttyper2009prop(21)
                		aty_BarnSyg_on = aty_on
                		aty_BarnSyg_tf = aty_tfval
                		if cint(aty_BarnSyg_on) = 1 then
		                '*** BarnSyg ****'   
                	    BarnSygtimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS BarnSygtimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 21) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                BarnSygtimer = oRec2("BarnSygtimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(BarnSygtimer)) <> 0 then
		                BarnSygtimer = BarnSygtimer
		                else
		                BarnSygtimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manBarnSygtimer = (BarnSygtimer)
						case 1
						tirBarnSygtimer = (BarnSygtimer)
						case 2
						onsBarnSygtimer = (BarnSygtimer)
						case 3
						torBarnSygtimer = (BarnSygtimer)
						case 4
						freBarnSygtimer = (BarnSygtimer)
						case 5
						lorBarnSygtimer = (BarnSygtimer)
						case 6
						sonBarnSygtimer = (BarnSygtimer)
						end select 
						
						
						end if
						
						
						'*** Lage ***'
                		call akttyper2009prop(81)
                		aty_Lage_on = aty_on
                		aty_Lage_tf = aty_tfval
                		if cint(aty_Lage_on) = 1 then
		                '*** Lage ****'   
                	    Lagetimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Lagetimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 81) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Lagetimer = oRec2("Lagetimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Lagetimer)) <> 0 then
		                Lagetimer = Lagetimer
		                else
		                Lagetimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manLagetimer = (Lagetimer)
						case 1
						tirLagetimer = (Lagetimer)
						case 2
						onsLagetimer = (Lagetimer)
						case 3
						torLagetimer = (Lagetimer)
						case 4
						freLagetimer = (Lagetimer)
						case 5
						lorLagetimer = (Lagetimer)
						case 6
						sonLagetimer = (Lagetimer)
						end select 
						
						
						end if
						
						
						'*** Sund ***'
                		call akttyper2009prop(8)
                		aty_Sund_on = aty_on
                		aty_Sund_tf = aty_tfval
                		if cint(aty_Sund_on) = 1 then
		                '*** Sund ****'   
                	    Sundtimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Sundtimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 8) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Sundtimer = oRec2("Sundtimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Sundtimer)) <> 0 then
		                Sundtimer = Sundtimer
		                else
		                Sundtimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manSundtimer = (Sundtimer)
						case 1
						tirSundtimer = (Sundtimer)
						case 2
						onsSundtimer = (Sundtimer)
						case 3
						torSundtimer = (Sundtimer)
						case 4
						freSundtimer = (Sundtimer)
						case 5
						lorSundtimer = (Sundtimer)
						case 6
						sonSundtimer = (Sundtimer)
						end select 
						
						
						end if
                	    
                	    
                	    '*** Frokost ***'
                		call akttyper2009prop(10)
                		aty_Frokost_on = aty_on
                		aty_Frokost_tf = aty_tfval
                		if cint(aty_Frokost_on) = 1 then
		                '*** Frokost ****'   
                	    Frokosttimer = 0
		                strSQL2 = "SELECT sum(t.timer) AS Frokosttimer FROM timer t WHERE "_
		                &" (tmnr = "& medid &" AND tdato = '"& useSQLd &"' AND tfaktim = 10) GROUP BY tmnr " 
		                'Response.Write strSQL2
		                'Response.flush
                		
		                oRec2.open strSQL2, oConn, 3 
		                if not oRec2.EOF then
			                Frokosttimer = oRec2("Frokosttimer") * 60
		                end if
		                oRec2.close 
                		
                		
                		
		                if len(trim(Frokosttimer)) <> 0 then
		                Frokosttimer = Frokosttimer
		                else
		                Frokosttimer = 0
		                end if
		                
		                select case intcounter
						case 0
						manFrokosttimer = (Frokosttimer)
						case 1
						tirFrokosttimer = (Frokosttimer)
						case 2
						onsFrokosttimer = (Frokosttimer)
						case 3
						torFrokosttimer = (Frokosttimer)
						case 4
						freFrokosttimer = (Frokosttimer)
						case 5
						lorFrokosttimer = (Frokosttimer)
						case 6
						sonFrokosttimer = (Frokosttimer)
						end select 
						
						
						end if
	              
	                end if 'visning fradrag / tillæg
					
					
					
	
	next
	
	
	
	
	
	
	select case visning 
	case 0%>


	
	
	<h4><img src="../ill/ikon_stempelur_24.png" alt="" border="0">&nbsp; <%=tsa_txt_123 %>&nbsp;- 
		<%=tsa_txt_005 %>: <%=datepart("ww", stDato, 2, 2)%></h4>
		
		<!-- Denne uge, nuværende login -->
		<%if datepart("ww", stDato, 2, 2) =  datepart("ww", now, 2, 2) then%>
		<%=tsa_txt_134 %>: 
		<%
		
		sLoginTid = "00:00:00"
		
		strSQL = "SELECT l.id AS lid, l.login "_
		&" FROM login_historik l WHERE "_
		&" l.mid = " & medid &" AND stempelurindstilling <> -1"_
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
		logindiffSidste = datediff("n", sLoginTid, now, 2, 2) 
		%>
		
		<br /><%=tsa_txt_135 %>: 
		<%call timerogminutberegning(logindiffSidste) %>
		<b><%=thoursTot &":"& left(tminTot, 2) %>&nbsp;t.</b>
		
		
		<%end if '*** Denne uge / nuværende login **' %>
		
	<table cellspacing=1 cellpadding=2 border=0 width=100% bgcolor="#c4c4c4">
	<tr bgcolor="#d6dff5">
		<td style="width:100px;">
            &nbsp;</td>
		
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_128 %> d. <%=formatdatetime(stDato, 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_129 %> d. <%=formatdatetime(dateadd("d", 1, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_130 %> d. <%=formatdatetime(dateadd("d", 2, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_131 %> d. <%=formatdatetime(dateadd("d", 3, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_132 %> d. <%=formatdatetime(dateadd("d", 4, stDato), 2) %></b></td>
	    <td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_133 %> d. <%=formatdatetime(dateadd("d", 5, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_127 %> d. <%=formatdatetime(dateadd("d", 6, stDato), 2) %></b></td>
		<td width=50 bgcolor="#ffdfdf" valign=bottom align=right><b><%=global_txt_167 %></b></td>
	</tr>
	<tr bgcolor="#ffffff">
		<td><%=tsa_txt_137 %>:</td>
		<td valign=top align=right><%call timerogminutberegning(manMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt + manMin/1%>
		<td valign=top align=right><%call timerogminutberegning(tirMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt + tirMin/1%>
		<td valign=top align=right><%call timerogminutberegning(onsMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt + onsMin/1%>
		<td valign=top align=right><%call timerogminutberegning(torMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  torMin/1%>
		<td valign=top align=right><%call timerogminutberegning(freMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  freMin/1%>
		<td valign=top align=right><%call timerogminutberegning(lorMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  lorMin/1%>
		<td valign=top align=right><%call timerogminutberegning(sonMin)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIalt = ugeIalt +  sonMin/1%>
		<td valign=top align=right>
		<%call timerogminutberegning(ugeIalt)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	<tr bgcolor="lightgrey">
		<td><%=tsa_txt_138 %>:</td>
		<td valign=top align=right><%call timerogminutberegning(-manMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + manMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-tirMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + tirMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-onsMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + onsMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-torMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + torMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-freMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + freMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-lorMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + lorMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-sonMinPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltPause = ugeIaltPause + sonMinPause%>
		<td valign=top align=right><%call timerogminutberegning(-ugeIaltPause)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<!-- Fradrag / Tillæg via Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_168 %>:*</td>
		<td valign=top align=right><%call timerogminutberegning(manFraTimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (manFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFraTimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (tirFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFraTimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (onsFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFraTimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (torFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFraTimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (freFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorFraTimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (lorFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonFraTimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFraTilTimer = ugeIaltFraTilTimer + (sonFraTimer)%>
		<td valign=top align=right><%call timerogminutberegning(ugeIaltFraTilTimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<!-- total -->
	<%
	totMan = manMin - (manMinPause - (manFraTimer))
	totTir = tirMin - (tirMinPause - (tirFraTimer))
	totOns = onsMin - (onsMinPause - (onsFraTimer))
	totTor = torMin - (torMinPause - (torFraTimer))
	totFre = freMin - (freMinPause - (freFraTimer))
	totLor = lorMin - (lorMinPause - (lorFraTimer))
	totSon = sonMin - (sonMinPause - (sonFraTimer))
	%>
	
	 <tr bgcolor="#ffdfdf">
		<td><b><%=global_txt_167%>:</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totMan)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTir)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totOns)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTor)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
	    <td valign=top align=right><b><%call timerogminutberegning(totFre)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totLor)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totSon)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(ugeIalt - (ugeIaltPause - (ugeIaltFraTilTimer)))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<%loginTimerTot = ugeIalt - (ugeIaltPause - ugeIaltFraTilTimer)  %>
	</tr>
	
 
	<!-- Normtimer -->
	<tr bgcolor="#ffffff">
		<td><%=tsa_txt_259 %>:</td>
		
		<%call normtimerper(medid, varTjDatoUS_man, 6) %>
		
		<td valign=top align=right><%call timerogminutberegning(ntimMan*60)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimTir*60)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimOns*60)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimTor*60)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimFre*60)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<td valign=top align=right><%call timerogminutberegning(ntimLor*60)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		
		<td valign=top align=right><%call timerogminutberegning(ntimSon*60)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		
		<%
		NormTimerWeekTot = 0
		NormTimerWeekTot = (ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon) * 60 %>
		
		<td valign=top align=right><%call timerogminutberegning(NormTimerWeekTot)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		</tr>
	    
	  <!-- Saldo -->  
	  <tr bgcolor="#EFf3FF">
		<td style="height:20px;"><b><%=global_txt_163 %>:</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totMan - (ntimMan*60))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTir - (ntimTir*60))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totOns - (ntimOns*60))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totTor - (ntimTor*60))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
	    <td valign=top align=right><b><%call timerogminutberegning(totFre - (ntimFre*60))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totLor - (ntimLor*60))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning(totSon - (ntimSon*60))
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<td valign=top align=right><b><%call timerogminutberegning((ugeIalt - (ugeIaltPause - (ugeIaltFraTilTimer))) - NormTimerWeekTot)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</b></td>
		<%loginTimerTot = ugeIalt - (ugeIaltPause - ugeIaltFraTilTimer)  %>
	</tr>
	</table>
	
	<%if datepart("ww", stDato, 2, 2) =  datepart("ww", now, 2, 2) then%>
	
	<%=tsa_txt_139 %>: <% 
	call timerogminutberegning(logindiffSidste+(loginTimerTot))
	%>
	<b><%=thoursTot &":"& left(tminTot, 2) %>&nbsp;t.</b>
	
    <%end if ' *** Denne uge / Nuværende login **'%>
	
	<br /><br />
	<table><tr><td class=lille>
	<%
	
	Response.Write "*) Tillægs typer: " & akttypenavnTil
	Response.Write "<br>Fradrags typer: " & akttypenavnFra
	
	%>
	</td></tr></table>
	    
	<br /><a href="#" id="udspec" class="vmenu">+ Udspecificering</a> (fraværs typer)<br />
	<div id="udspecdiv" style="position:relative; visibility:hidden; display:none;">
	<table cellspacing=1 cellpadding=2 border=0 width=100% bgcolor="#c4c4c4">    
	    <tr bgcolor="#FFFFFF">
	    <td colspan=9><br /><b>Udspecificering på fraværs typer</b> <br />
	    Ikke medregnet i saldo, med mindre de er en del af <%=global_txt_168 %> typerne*.</td>
	</tr>
	<tr bgcolor="#d6dff5">
		<td style="width:100px;">
            &nbsp;</td>
		
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_128 %> d. <%=formatdatetime(stDato, 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_129 %> d. <%=formatdatetime(dateadd("d", 1, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_130 %> d. <%=formatdatetime(dateadd("d", 2, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_131 %> d. <%=formatdatetime(dateadd("d", 3, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_132 %> d. <%=formatdatetime(dateadd("d", 4, stDato), 2) %></b></td>
	    <td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_133 %> d. <%=formatdatetime(dateadd("d", 5, stDato), 2) %></b></td>
		<td width=50 valign=bottom align=center class=lille><b><%=tsa_txt_127 %> d. <%=formatdatetime(dateadd("d", 6, stDato), 2) %></b></td>
		<td width=50 bgcolor="#ffdfdf" valign=bottom align=right><b><%=global_txt_167 %></b></td>
	</tr>
	
	    
	    <%if cint(aty_fleks_on) = 1 then %>
	<!-- Fleks Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_147 &" "& aty_fleks_tf%></td>
		<td valign=top align=right><%call timerogminutberegning(manFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (manFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (tirFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (onsFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (torFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (freFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (lorFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFlekstimer = ugeIaltFlekstimer + (sonFlekstimer)%>
		<td valign=top align=right><%call timerogminutberegning(ugeIaltFlekstimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	
	
	<%if cint(aty_Ferie_on) = 1 then %>
	<!-- Ferie Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_135 &" "& aty_Ferie_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (manFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (tirFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (onsFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (torFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (freFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (lorFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFerietimer = ugeIaltFerietimer + (sonFerietimer)%>
		<td valign=top align=right><%call timerogminutberegning(ugeIaltFerietimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	<%if cint(aty_Syg_on) = 1 then %>
	<!-- Syg Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_138 &" "& aty_Syg_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (manSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (tirSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (onsSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(torSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (torSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(freSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (freSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (lorSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSygtimer = ugeIaltSygtimer + (sonSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(ugeIaltSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	
	<%if cint(aty_BarnSyg_on) = 1 then %>
	<!-- BarnSyg Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_139 &" "& aty_BarnSyg_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (manBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (tirBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (onsBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(torBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (torBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(freBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (freBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (lorBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltBarnSygtimer = ugeIaltBarnSygtimer + (sonBarnSygtimer)%>
		<td valign=top align=right><%call timerogminutberegning(ugeIaltBarnSygtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
	
	<%if cint(aty_Lage_on) = 1 then %>
	<!-- Lage Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_160 &" "& aty_Lage_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (manLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (tirLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (onsLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(torLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (torLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(freLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (freLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (lorLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltLagetimer = ugeIaltLagetimer + (sonLagetimer)%>
		<td valign=top align=right><%call timerogminutberegning(ugeIaltLagetimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>
	
		<%if cint(aty_Sund_on) = 1 then %>
	<!-- Sund Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_148 &" "& aty_Sund_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (manSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (tirSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (onsSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(torSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (torSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(freSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (freSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (lorSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltSundtimer = ugeIaltSundtimer + (sonSundtimer)%>
		<td valign=top align=right><%call timerogminutberegning(ugeIaltSundtimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>    
	
	<%if cint(aty_Frokost_on) = 1 then %>
	<!-- Frokost Realtimer -->
	<tr bgcolor="#ffffff">
		<td><%=global_txt_133 &" "& aty_Frokost_tf %></td>
		<td valign=top align=right><%call timerogminutberegning(manFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (manFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(tirFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (tirFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(onsFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (onsFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(torFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (torFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(freFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (freFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(lorFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (lorFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(sonFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
		<%ugeIaltFrokosttimer = ugeIaltFrokosttimer + (sonFrokosttimer)%>
		<td valign=top align=right><%call timerogminutberegning(ugeIaltFrokosttimer)
		Response.write "<font class=megetlillesort>"& thoursTot &":"& left(tminTot, 2)%> t</td>
	</tr>
	
	<%end if %>    
	    
	    
	</table>
	</div>
	
	
    
	
	
	<%
	
	case 2
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
    call timerogminutberegning(totalTimerPer100) %>
	<!-- formatnumber(totalTimerPer100, 2) -->
	
	<%=thoursTot &":"& left(tminTot, 2) %>
	
	<%
	case 3
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
	case 4 ''** Avg. Week ***' 
	
	if weekDiff <> 0 then
	weekDiff = weekDiff
	else
	weekDiff = 1
	end if
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)/weekDiff
	
	
	call timerogminutberegning(((totaltimerPer-totalpausePer)/weekDiff)) %>
	<!-- formatnumber(totalTimerPer100, 2) -->
	
	<%=thoursTot &":"& left(tminTot, 2) %>
	<%
	case else
	
	
	totalTimerPer100 = 0
	totalTimerPer100 = (totaltimerPer-totalpausePer)
	
	
	call timerogminutberegning(totaltimerPer-totalpausePer) %>
	<!-- formatnumber(totalTimerPer100, 2) -->
	
	<%=thoursTot &":"& left(tminTot, 2) %> 
	
	<%
	end select '** Visinng ***'%>
	
	

<%
end function


    '*** Akttyper ***'
 	public aty_fakbar, aty_real, aty_pre, aty_enh, aty_on, aty_tfval, aty_medpafak
	function akttyper2009Prop(fakturerbartype)
	
	aty_pre = 0
	aty_fakbar = 0
	aty_real = 0
	aty_medpafak = 0
	
	'Response.Write fakturerbartype
	'Response.flush
	
	
	 strSQL4 = "SELECT aty_id, aty_on, aty_label, aty_desc, aty_on_realhours, "_
	 &" aty_on_invoice, aty_on_invoiceble, aty_pre, aty_enh, et_navn, aty_on_workhours FROM akt_typer "_
	 &" LEFT JOIN enheds_typer ON (et_id = aty_enh) "_
	 &" WHERE aty_id = "& fakturerbartype 
	 
	'Response.Write strSQL4
	'Response.flush
	 
	 oRec4.open strSQL4, oConn, 3
	 
	
	 if not oRec4.EOF then
	
	  if oRec4("aty_on_invoiceble") = 1 then
	  aty_fakbar = 1
	  else
	  aty_fakbar = 0
	  end if
	  
	  if oRec4("aty_on_invoice") = 1 then
	  aty_medpafak = 1
	  else
	  aty_medpafak = 0
	  end if
	  
	  if oRec4("aty_on_realhours") = 1 then
	  aty_real = 1
	  else
	  aty_real = 0
	  end if
	  
	  if oRec4("aty_pre") <> 0 then
	  aty_pre = oRec4("aty_pre")
	  else
	  aty_pre = 0
	  end if
	  
	  aty_on = oRec4("aty_on")
	  
	  aty_enh = oRec4("et_navn")
	  
	  
	  select case oRec4("aty_on_workhours")
	  case 0
	  aty_tfval = ""
	  case 1 '** fradrag
	  aty_tfval = "(-)"
	  case 2 '** tillæg
	  aty_tfval = "(+)"
	  end select
	  
	  'Response.Write "aty_enh" & aty_enh
	  
	 end if 
	 oRec4.close
	
	end function
	
	
	'********** Er uge afsluttet??? ******'
	public ugeNrAfsluttet, showAfsuge, cdAfs 
    function erugeAfslutte(sm_aar, sm_sidsteugedag, sm_mid)
            
            showAfsuge = 1
            ugeNrAfsluttet = "1-1-2044"
            
            strSQLafslut = "SELECT status, afsluttet, uge FROM ugestatus WHERE WEEK(uge, 1) = "& sm_sidsteugedag &" AND YEAR(uge) = "& sm_aar &" AND mid = "& sm_mid
		    'Response.write strSQLafslut
		    oRec3.open strSQLafslut, oConn, 3 
		    if not oRec3.EOF then
    			
			    showAfsuge = 0
			    cdAfs = oRec3("afsluttet")
			    ugeNrAfsluttet = oRec3("uge")
    		
		    end if
		    oRec3.close 
            
            
    end function
    
    
    function valutakoder(i, valuta)
	%>
	<select name="FM_valuta_<%=i %>" id="FM_valuta_<%=i %>" style="width:65px; font-size:9px;">
		    
		    <%
		    strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
    		
    		
		    oRec3.open strSQL3, oConn, 3 
		    while not oRec3.EOF 
    		
		    if cint(valuta) = oRec3("id") then
		    valGrpCHK = "SELECTED"
		    else
		    valGrpCHK = ""
		    end if
		    
		   
		    %>
		    <option value="<%=oRec3("id")%>" <%=valGrpCHK %>><%=oRec3("valutakode")%></option>
		    <%
		    oRec3.movenext
		    wend
		    oRec3.close %>
		    </select>
	<%
	end function
	
	
	public dblKurs
	function valutaKurs(intValuta)
	    '**** Finder aktuel kurs ***'
       dblKurs = 100
       strSQL = "SELECT kurs FROM valutaer WHERE id = " & intValuta
       oRec.open strSQL, oConn, 3
       if not oRec.EOF then
       dblKurs = replace(oRec("kurs"), ",", ".")
       end if 
       oRec.close
	end function
	
	
	public valBelobBeregnet
	function beregnValuta(belob,frakurs,tilkurs)
	
	        valBelobBeregnet = belob * (frakurs/tilkurs)
    
    if len(valBelobBeregnet) <> 0 then
    valBelobBeregnet = valBelobBeregnet
    else
    valBelobBeregnet = 0
    end if
    
    
    
    valBelobBeregnet = valBelobBeregnet/1
    'Response.Write valBelobBeregnet & "<br>"
    'Response.flush
	end function
	
	
	public akttypenavn
	function akttyper(akttype, visning)
	
	if len(trim(akttype)) <> 0 then
	akttype = akttype
	else
	akttype = 0
	end if
	
	    select case akttype 
	    case 1 
	    akttypenavn = global_txt_129 '"Fakturerbar"
	    case 5 '2
	    akttypenavn = global_txt_130 '"Km"
	    case 2 '0
	    akttypenavn = global_txt_131 '"Ikke fakturerbar"
	    case 6
	    akttypenavn = replace(global_txt_132, "|", "&") '"Salg & NewBizz"
	    case 7
	    akttypenavn = global_txt_147 '"Flex brugt"
	    case 8
	    akttypenavn = global_txt_148 '"Læge / Massage / Fysioterapi"
	    case 9
	    akttypenavn = global_txt_150 '"Pause"
	    case 10
	    akttypenavn = global_txt_133 '"Frokost / Pause"
	    case 11
	    akttypenavn = global_txt_134 '"Ferie planlagt"
	    case 12
	    akttypenavn = global_txt_136 '"Ferie fridage optjent"
        case 13
	    akttypenavn = global_txt_137 '"Ferie fridage brugt"
	    case 14
	    akttypenavn = global_txt_135 '"Ferie afholdt"
	    case 15
	    akttypenavn = global_txt_143 '"Ferie optjent"
	    case 16
	    akttypenavn = global_txt_156 '"Ferie udbetalt"
	    case 17
	    akttypenavn = global_txt_149 '"Ferie fridage udbetalt"
	    case 18
	    akttypenavn = global_txt_164 '"Ferie Fridage Planlagt"
	    case 19
	    akttypenavn = global_txt_165 '"Ferie u. Løn"
	    case 20
	    akttypenavn = global_txt_138 '"Syg"
	    case 21
	    akttypenavn = global_txt_139 '"Barn syg"
	    case 30
	    akttypenavn = global_txt_144 '"Afspad. optjent"
        case 31
	    akttypenavn = global_txt_145 '"Afspad. brugt"
	    case 32
	    akttypenavn = global_txt_146 '"Afspad. udbetalt"
	    case 33
	    akttypenavn = global_txt_159 '"Afspad. Ø udbetalt"
	    case 50
	    akttypenavn = global_txt_166 '"Dag"
	    case 51
	    akttypenavn = global_txt_151 '"Nat"
	    case 52
	    akttypenavn = global_txt_152 '"Weekend"
	    case 53
	    akttypenavn = global_txt_153 '"Afspad. optjent"
	    case 54
	    akttypenavn = global_txt_154 '"Weekend Nat"
	    case 55
	    akttypenavn = global_txt_155 '"Weekend Aften"
	    case 60
	    akttypenavn = global_txt_157 '"Ad-hoc"
	    case 61
	    akttypenavn = global_txt_158 '"Stk. Antal"
	    case 81
	    akttypenavn = global_txt_160 '"Læge"
	    case 90
	    akttypenavn = global_txt_161 '"E1"
	    case 91
	    akttypenavn = global_txt_162 '"E2"
	    case else
	    akttypenavn = "-"
	    end select
	
	
	
	end function
	
	
    function sltque(slturl,slttxt,slturlalt,slttxtalt,lft,tp)%>
	
	<%if len(slttxtalt) <> 0 then
	usejaimg = "ja"
	else
	usejaimg = "sletja"
	end if %>
	
	<div id="slet" style="position:absolute; left:<%=lft%>px; top:<%=tp%>px; background-color:#ffffe1; visibility:visible; border:2px #8cAAe6 dashed;">
	<table cellspacing="0" cellpadding="10" border="0" bgcolor="#ffffe1">
	<tr>
	    <td bgcolor="#ffffff" style="border-bottom:1px #999999 solid;"><h4>Slet?</h4> </td>
	    <td bgcolor="#ffffff" align=right style="border-bottom:1px #999999 solid;"><img src="../ill/garbage_information.gif" alt="Slet?" border="0"></td>
	   
	    <%if len(slttxtalt) <> 0 then %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxtalt %>
	    </td>
	   </tr>
	  <tr><td colspan=2>
		<a href=<%=slturlalt %>><img src="../ill/sletja.gif" alt="Ja - slet" border="0"></a>
		</td>
	</tr>
	    
	    <%end if %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxt %>
	    </td>
	   </tr>
	  <tr><td>
		<a href=<%=slturl %>><img src="../ill/<%=usejaimg %>.gif" alt="Ja - slet" border="0"></a>
		</td>
		<td align=right>
		<a href="Javascript:history.back()"><img src="../ill/stop.gif" alt="Nej - tilbage" border="0"></a></td>
	</tr>
	</table>
	</div>
	
	<%
	end function
	
	 function sltque_Small(slturl,slttxt,slturlalt,slttxtalt,lft,tp)%>
	
	<%if len(slttxtalt) <> 0 then
	usejaimg = "ja"
	else
	usejaimg = "sletja"
	end if %>
	
	<div id="Div4" style="position:absolute; width:275px; left:<%=lft%>px; top:<%=tp%>px; padding:2px; background-color:#ffffe1; visibility:visible; border:1px #8cAAe6 solid;">
	<table cellspacing="0" cellpadding="2" border="0" bgcolor="#ffffe1" width=100%>
	<tr>
	    <td bgcolor="#ffffff" style="border-bottom:1px #999999 solid;"><b>Slet (nulstil)?</b> </td>
	    <td bgcolor="#ffffff" align=right style="border-bottom:1px #999999 solid;"><img src="../ill/garbage_information.gif" alt="Slet?" border="0"></td>
	   
	    <%if len(slttxtalt) <> 0 then %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxtalt %>
	    </td>
	   </tr>
	  <tr><td colspan=2>
		<a href=<%=slturlalt %>><img src="../ill/sletja.gif" alt="Ja - nulstil" border="0"></a>
		</td>
	</tr>
	    
	    <%end if %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1" style="padding:5px;"><%=slttxt %>
	    </td>
	   </tr>
	  <tr><td style="padding-left:5px;">
		<a href="Javascript:history.back()" class=rmenu><< Nej, tilbage</a>
		</td>
	    <td align=right style="padding-right:5px;">
		<a href=<%=slturl %> class=red>Ja, nulstil denne >></a></td>
	</tr>
	</table>
	</div>
	
	<%
	end function
	
	function sltquePopup(slturl,slttxt,slturlalt,slttxtalt,lft,tp)%>
	
	<%if len(slttxtalt) <> 0 then
	usejaimg = "ja"
	else
	usejaimg = "sletja"
	end if %>
	
	<div id="Div1" style="position:absolute; left:<%=lft%>px; top:<%=tp%>px; background-color:#ffffe1; visibility:visible; border:2px #8cAAe6 dashed;">
	<table cellspacing="0" cellpadding="10" border="0" bgcolor="#ffffe1">
	<tr>
	    <td bgcolor="#ffffff" style="border-bottom:1px #999999 solid;"><h4>Slet?</h4> </td>
	    <td bgcolor="#ffffff" align=right style="border-bottom:1px #999999 solid;"><img src="../ill/garbage_information.gif" alt="Slet?" border="0"></td>
	   
	    <%if len(slttxtalt) <> 0 then %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxtalt %>
	    </td>
	   </tr>
	  <tr><td colspan=2>
		<a href=<%=slturlalt %>><img src="../ill/sletja.gif" alt="Ja - slet" border="0"></a>
		</td>
	</tr>
	    
	    <%end if %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxt %>
	    </td>
	   </tr>
	  <tr><td>
		<a href=<%=slturl %>><img src="../ill/<%=usejaimg %>.gif" alt="Ja - slet" border="0"></a>
		</td>
		<td align=right>
		<a href="Javascript:window.close()"><img src="../ill/stop.gif" alt="Nej - Luk vindue" border="0"></a></td>
	</tr>
	</table>
	</div>
	
	<%
	end function
	
	function eksportogprint(ptop,pleft,pwdt)
	%>
	<div id=eksport style="position:absolute; background-color:#ffffff; width:<%=pwdt%>px; left:<%=pleft%>px; top:<%=ptop%>px; border:1px silver solid; padding:3px 3px 3px 3px;">
    <table cellpadding=5 cellspacing=0 border=0 width=100%>
    <tr>
    <td height=30 bgcolor="#FFFFe1" style="border-bottom:1px #ffff99 solid;" colspan=2><b>Eksport & Print:</b></td>
    </tr>
	<%
	end function
	
	function eksportogprint09(ptop,pleft,pwdt)
	%>
	<div id=Div2 style="position:relative; background-color:#ffffff; width:<%=pwdt%>px; left:<%=pleft%>px; top:<%=ptop%>px; padding:3px 3px 3px 3px;">
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
    <tr>
    <td bgcolor="#FFFFe1" style="border-bottom:1px #ffff99 solid; height:20px; padding:10px 10px 0px 10px;" colspan=2><h3>Eksport og Print</h3></td>
    </tr>
    <tr>
    <td valign=top style="padding:10px 10px 10px 10px;">
	<%
	end function
	
	function filteros09(ptop,pleft,pwdt,txt,visning, tdheight)
	
	select case visning
	case 1
	bgt = "#8CAAE6"
	bgt_border = "#5582d2"
	bgtd = ""
	divbg = "#EFF3FF" 
	case 2 '** print
	bgt = "#FFFFe1"
	bgt_border = "#ffff99"
	divbg = "#EFF3FF" 
	case 3
	bgt = "#C1D9F0"
	bgt_border = "#8CAAE6"
	divbg = "#EFF3FF" 
	case 4
	bgt = "#Cccccc"
	bgt_border = "#999999" 
	divbg = "#FFFFFF"
	
	case 5
	end select
	
	%>
	<div id=Div3 style="position:relative; width:<%=pwdt%>px; left:<%=pleft%>px; top:<%=ptop%>px; border:1px #cccccc solid; overflow:auto; height:<%=tdheight%>px; background-color:<%=divbg%>">
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
    <tr>
    <td bgcolor="<%=bgt %>" style="border-bottom:1px <%=bgt_border %> solid; padding:10px 10px 0px 10px;" colspan=2><h3><%=txt %></h3></td>
    </tr>
    <tr>
    <td valign=top style="padding:10px 10px 10px 10px;">
	<%
	end function
	
	
	function filterheader(ptop,pleft,pwdt,pTxt)
	
	pTxt = replace(global_txt_119, "|", "&")
	
	
	%>
	<div id="filter" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=pwdt %>px; border:1px #8caae6 solid; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#EFF3FF" style="border-bottom:1px #D6DfF5 solid;">
        <img src="../ill/find.png" /></td>
        <td bgcolor="#EFF3FF" align=left style="border-bottom:1px #D6DfF5 solid;"><b><%=pTxt%>:</b></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFFF" style="padding:5px;">
	
    
	
	<%
	end function
   
   
    function filterheaderid(ptop,pleft,pwdt,pTxt,fiVzb,fiDsp,fid,abrel)
	
	pTxt = replace(global_txt_119, "|", "&")
	
	
	%>
	<div id="<%=fid %>" style="position:<%=abrel%>; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=pwdt %>px; border:1px #8caae6 solid; left:<%=pleft%>px; top:<%=ptop%>px; visibility:<%=fiVzb%>; display:<%=vzDsp%>;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#EFF3FF" style="border-bottom:1px #D6DfF5 solid;">
        <img src="../ill/find.png" /></td>
        <td bgcolor="#EFF3FF" align=left style="border-bottom:1px #D6DfF5 solid;"><b><%=pTxt%>:</b></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFFF" style="padding:5px;">
	
    
	
	<%
	end function   
     
     
    function tableDiv(tTop,tLeft,tWdth)
	
	if print = "j" then
	bd = 0
	else
	bd = 1
	end if
	
	%>
	<div id="maintable" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=tWdth%>px; border:<%=bd%>px #8caae6 solid; left:<%=tLeft%>px; top:<%=tTop%>px; visibility:visible;">
    <%
	end function
	
	 function tableDivAbs(tTop,tLeft,tWdth,tHgt,tId)
	
	if print = "j" then
	bd = 0
	else
	bd = 1
	end if
	
	%>
	<div id="<%=tId%>" style="position:absolute; background-color:#ffffff; height:<%=tHgt%>px; padding:3px 3px 3px 3px; width:<%=tWdth%>px; border:<%=bd%>px #8caae6 solid; left:<%=tLeft%>px; top:<%=tTop%>px; visibility:visible; z-index:0; overflow:auto;">
    <%
	end function
	
	
	
	 function tableDivWid(tTop,tLeft,tWdth,tId, tVzb, tDsp)
	 if print = "j" then
	bd = 0
	else
	bd = 1
	end if
	 
	%>
	<div id="<%=tId%>" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=tWdth%>px; border:<%=bd%>px #8caae6 solid; left:<%=tLeft%>px; top:<%=tTop%>px; visibility:<%=tVzb%>; display:<%=tDsp%>; z-index:1000">
    <%
	end function
    
     
     
    function sideinfo(itop,ileft,iwdt)
	%>
	<div id="sideinfo" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=iWdt %>px; border:1px red solid; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:visible; z-index:1000000;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#FFFFe1" style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/lifebelt.png" /></td>
        <td bgcolor="#FFFFe1" align=left style="border-bottom:1px #C4c4c4 solid;"><b>Hjælp & Sideinfo:</b></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFfff" style="padding:5px;">
	
    
	
	<%
	end function
	
	
	 function sideinfoId(itop,ileft,iwdt,ihgt,iId,idsp,ivzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
	%>
	
	<script>
	
	$(document).ready(function() {
    
    
    $("#showpagehelp").click(function() {
    
    //alert("her")
    
    $("#pagehelp_bread").show("fast", function() {
        // use callee so don't have to name the function
    //$(this).next().show("fast", arguments.callee);

    $("#pagehelp_bread").css("display", "");
    $("#pagehelp_bread").css("visibility", "visible");
        
    });
	
	

	$("#pagehelp").hide("fast", function() {
	    // use callee so don't have to name the function
	    //$(this).next().show("fast", arguments.callee);
	});


});
    
    

    $("#hidepagehelp").click(function() {


    $("#pagehelp").show("slow", function() {
            // use callee so don't have to name the function
            //$(this).next().show("fast", arguments.callee);
        });


        $("#pagehelp_bread").hide("slow", function() {
            // use callee so don't have to name the function
            //$(this).next().show("fast", arguments.callee);
        });

     });   
	    
	    
        
    });



    
     

</script>
	
	<div id="<%=iId %>" style="position:absolute; background-color:#ffffff; padding:1px 1px 0px 1px; width:<%=iWdt%>px; border:1px silver solid; border-bottom:0px; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:<%=ivzb%>; display:<%=idsp%>; z-index:9000000; overflow:hidden;">
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
    <tr bgcolor="#FF6666"><td align=center>
        <a href="#" id="showpagehelp" class=alt>Hjælp & Sideinfo +</a></td>
        
   </tr>
        
        <!-- onclick="dsppagehelp('<%=iId%>')" -->
	</table>
	</div>
	
	
	<div id="<%=ibId %>" style="position:absolute; background-color:#ffffff; padding:5px 5px 5px 5px; width:<%=ibWdt %>px; height:<%=ibhgt %>px; border:1px silver solid; left:<%=ibleft %>px; top:<%=ibtop %>px; visibility:hidden; display:none; z-index:9000000; overflow:auto;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr bgcolor="#FF6666"><td align=center width=40 style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/lifebelt.png" /></td>
        <td align=left style="border-bottom:1px #C4c4c4 solid;" class=alt><b>Hjælp & Sideinfo</b></td>
         <td style="padding-right:20px;" align="right"><a href="#" id="hidepagehelp" class=alt>[x]</a></td>
       
        </tr>
        <!-- onclick="dsppagehelp('<%=iId%>')" -->
	<tr>
	</table>
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFff" style="padding:5px;">
	
    
	
	<%
	end function
	
    function sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb)
	%>
	
	<script>
	    function sidemsgclose(idthis) {
	        //alert(idthis)
	        document.getElementById(idthis).style.visibility = "hidden"
	        document.getElementById(idthis).style.display = "none"
	    }
	
	
	</script>
	
	 <div id="<%=iId %>" style="position:absolute; background-color:#FFFFFF; padding:3px 3px 3px 3px; width:<%=iWdt %>px; border:1px #999999 solid; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:<%=ivzb%>; display:<%=idsp%>; z-index:1000000;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#FF6666" style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/ikon_message_24.png" /></td>
        <td bgcolor="#FF6666" align=left style="border-bottom:1px #C4c4c4 solid;"><b>Meddelelse</b></td>
	    <td bgcolor="#FF6666" align=right style="border-bottom:1px #C4c4c4 solid; padding-right:5px;"><a href="#" onclick="sidemsgclose('<%=iId%>')" class="vmenu">[x]</a></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFe1" style="padding:5px;">
	
    
	
	<%
	end function
	
	  function sidemsgId2(itop,ileft,iwdt,iId,idsp,ivzb)
	%>
	
	<script>
	    function sidemsgclose2(idthis) {
	        //alert(idthis)
	        document.getElementById(idthis).style.visibility = "hidden"
	        document.getElementById(idthis).style.display = "none"
	    }
	
	
	</script>
	
	<div id="<%=iId %>" style="position:absolute; background-color:#FFFFFF; padding:3px 3px 3px 3px; width:<%=iWdt %>px; border:2px #c4c4c4 solid; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:<%=ivzb%>; display:<%=idsp%>; z-index:1000000;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#FFFF99" style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/ikon_message_24.png" /></td>
        <td bgcolor="#FFFF99" align=left style="border-bottom:1px #C4c4c4 solid;"><b>Meddelelse</b></td>
	    <td bgcolor="#FFFF99" align=right style="border-bottom:1px #C4c4c4 solid; padding-right:5px;"><a href="#" onclick="sidemsgclose2('<%=iId%>')" class="vmenu">[x]</a></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFFF" valing="top" style="padding:5px;">
	
    
	
	<%
	end function
	
	function opretNy(url, text, otoppx, oleftpx, owdtpx)
	%>
	<div style="position:relative; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td>
	<a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	
	function opretNyAB(url, text, otoppx, oleftpx, owdtpx)
	%>
	<div style="position:absolute; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td>
	<a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	
	
	
	function opretNy_blank(url, text, otoppx, oleftpx, owdtpx)
	%>
	<div style="position:relative; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td>
	<a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_blank"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_blank"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	
	function opretNyJava(url, text, otoppx, oleftpx, owdtpx, java)
	%>
	<div style="position:relative; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td>
	<a href='<%=url %>' onClick="<%=java %>" class='vmenu' alt="<%=text %>"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" onClick="<%=java %>"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	

	
	
	
	
	
	function insertDelhist(deltype, delid, delnr, delnavn, mid, mnavn)
	
	
	strSQLdelhist = "INSERT INTO delete_hist (deltype, delid, delnr, delnavn, mid, mnavn) VALUES "_
	&" ('"& deltype &"', "& delid &", '"& delnr &"', '"& delnavn &"', "& mid &", '"& mnavn &"')"
	
	oConn.execute(strSQLdelhist)
	
	end function
	
	
	
	
	public aty_sql_real, aty_sql_fak, aty_sql_fak_on, aty_options, aty_sql_frawhours, aty_sql_sel
	public aty_sql_ikfakbar, aty_sql_fakbar, aty_sql_onfak, aty_sql_realhours, aty_sql_realHoursFakbar
	public aktiveTyper, aty_sql_realHoursIkFakbar, aty_sql_tilwhours
	 
	function akttyper2009(dothis)
	
	'Response.Write "akttype_sel::"& akttype_sel
	 
	 aty_sql_fakbar = "fakturerbar = 0"
	 aty_sql_ikfakbar = "fakturerbar = 0"
	 aty_sql_onfak =  "aktiviteter.fakturerbar = 0"
	 
	 aty_sql_realhours = "tfaktim = 0"
	 aty_sql_realHoursFakbar = "t.tfaktim = 0"
	 aty_sql_realHoursIkFakbar = "t.tfaktim = 0"
	 
	 aty_sql_frawhours = "t.tfaktim = 0"
	 aty_sql_tilwhours = "t.tfaktim = 0"
	 aty_sql_sel = "t.tfaktim = 0"
	 
	 aty_lastsort = 0
	 
	 '** vis kun admin typer på internejob
	 if dothis = 1 then
	  
	  
	  call licKid()
	  '** finder kid på valgte job ***'
	  thisKid = 0
	  strSQLjid = "SELECT j.jobknr FROM job j "_
	  &" WHERE j.id = "& jobid 
	  
	  oRec4.open strSQLjid, oConn, 3
	  if not oRec4.EOF then
	  
	  thisKid = oRec4("jobknr")
	  
	  end if
	  oRec4.close
	  
	 if (cint(licensindehaverKid) <> cint(thisKid)) then 
	 'Vis kun admintyper på interne
	 AstrSQLwhKri = " AND ((aty_id BETWEEN 1 AND 6) OR (aty_id BETWEEN 50 AND 61) OR (aty_id BETWEEN 90 AND 100) OR aty_id = 30)"
	 else
	 AstrSQLwhKri = ""
	 end if
	
	end if
	 
	 'aty_on_invoiceble
	 '*** Typer der skal med på timeregsiden og medregnes i dagligt timeforbrug ***'
	 
	 strSQL4 = "SELECT aty_id, aty_label, aty_desc, aty_on_realhours, aty_on_invoice, "_
	 &" aty_on_invoice, aty_on_invoiceble, aty_sort, aty_on_workhours FROM akt_typer "_
	 &" WHERE aty_on = 1 "& AstrSQLwhKri &" ORDER BY aty_sort"
	
	 
	 'Response.Write strSQL4
	 'Response.flush
	 
	  oRec4.open strSQL4, oConn, 3
	 
	 xi = 0
	 while not oRec4.EOF
	 
	 if dothis = 1 then
	  
	  if oRec4("aty_on_realhours") = 1 then
	  meditimeregn = "M"
	  else
	  meditimeregn = "-"
	  end if
	  
	  if oRec4("aty_on_invoice") = 1 then
	  medpafak = "E"
	  else
	  medpafak = "-"
	  end if
	  
	  if oRec4("aty_on_invoiceble") = 1 then
	  fakturerbar = "Z"
	  else
	  fakturerbar = "-"
	  end if
	  
	  if oRec4("aty_on_workhours") = 1 then
	  fradrag = "F"
	  else
	  fradrag = "-"
	  end if
	  
	    call akttyper(oRec4("aty_id"), 1)
	    
	    if cint(strFakturerbart) = cint(oRec4("aty_id")) then
	    aktCHK = "SELECTED"
	    else
	    aktCHK = ""
	    end if
	    
	    aty_options = aty_options & "<option value='"& oRec4("aty_id")&"' "&aktCHK&">"& akttypenavn &""
	    
	    if level = 1 then
	    aty_options = aty_options & " ( "& meditimeregn & " "& medpafak &" "&fakturerbar&" "&  fradrag &" )"
	    end if
	    
	    aty_options = aty_options & "</option>"
	    
	 end if
	 
	 
	 if dothis = 2 then
	      
	      '** Aktiviteter tabel ***'  
	      '** Fakturerbare / ik fakturebare
	      if oRec4("aty_on_invoiceble") = 1 then
	      aty_sql_fakbar = aty_sql_fakbar & " OR fakturerbar = "& oRec4("aty_id")
	      else
	      aty_sql_fakbar = aty_sql_fakbar 
	      end if
	      
	      if oRec4("aty_on_invoiceble") = 2 then
	      aty_sql_ikfakbar = aty_sql_ikfakbar & " OR fakturerbar = "& oRec4("aty_id")
	      else
	      aty_sql_ikfakbar = aty_sql_ikfakbar 
	      end if
	        
	      
	      
	      
	      '** Timer tabel ***'
	        
	        '*** Tælles med i timereg **'
	        if oRec4("aty_on_realhours") = 1 then
	        aty_sql_realhours = aty_sql_realhours &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realhours = aty_sql_realhours
	        end if
	        
	        
	        '*** Fakturerbare timer **'
	        if oRec4("aty_on_invoiceble") = 1 then
	        aty_sql_realHoursFakbar = aty_sql_realHoursFakbar &" OR t.tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realHoursFakbar = aty_sql_realHoursFakbar
	        end if
	        
	         '*** Ikke Fakturerbare timer **'
	        if oRec4("aty_on_invoiceble") = 2 then
	        aty_sql_realHoursIkFakbar = aty_sql_realHoursIkFakbar &" OR t.tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realHoursIkFakbar = aty_sql_realHoursIkFakbar
	        end if
	        
	        
	          
	        '**** Fradrag fra løntimer ***'
	        if oRec4("aty_on_workhours") = 1 then
	        aty_sql_frawhours = aty_sql_frawhours &" OR t.tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_frawhours = aty_sql_frawhours
	        end if
	        
	        '**** Tillæg fra løntimer ***'
	        if oRec4("aty_on_workhours") = 2 then
	        aty_sql_tilwhours = aty_sql_tilwhours &" OR t.tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_tilwhours = aty_sql_tilwhours
	        end if
	        
	        
	         if oRec4("aty_on_invoice") = 1 then
          aty_sql_onfak = aty_sql_onfak & " OR aktiviteter.fakturerbar = "& oRec4("aty_id")
          else
          aty_sql_onfak = aty_sql_onfak 
          end if
	        
	        aktiveTyper = aktiveTyper & "#"&oRec4("aty_id") & "#, "
	        
	  end if
	 
	 
	 if dothis = 3 then
	 
	 if xi = 0 then
	 
	 %>
	 <b>Aktive aktivitets typer:</b> (Konfiguration, se kontrolpanel)<br />
			<table cellspacing=2 cellpadding=1 border=0 width=100%>
			<tr bgcolor="#EFf3ff">
			    <td valign=bottom><b>Id</b></td>
			    <td valign=bottom><b>Navn</b></td>
			    <td valign=bottom><b>Fakturerbar</b><br />
			    Ikke fakturerbar ell. administrativ</td>
			    <td valign=bottom><b>Medregnes i dagligt timeregnskab *)</b><br />
			    (Dvs. indgår denne type i de dagligt realiserede timer der udgør balancen mellem normeret tid pr. uge og realiseret tid)</td>
			    <%
                if session("stempelur") <> 0 then %>
			    <td valign=bottom><b>Tillæg/Fradrag +/- </b><br />
			    I fht. løntimer</td>
			    <%end if %>
			</tr>
	 <%end if
	 %>
			<tr>
			<td valign=top><%=oRec4("aty_id")%></td>
			<td>
			    <%call akttyper(oRec4("aty_id"), 1) %>
	            <%=akttypenavn%>
			</td>
			<td>
			<%select case oRec4("aty_on_invoiceble")
			case 1 %>
			Fakturerbar
			<%case 2 %>
			Ikke Fakt. bar
			<%case else %>
			Administrativ
			<%end select %></td>
			<td>
			<%if oRec4("aty_on_realhours") = 1 then %>
			Ja
			<%end if %>
			</td>
			<%
    if session("stempelur") <> 0 then %>
			<td>
			<%select case oRec4("aty_on_workhours") 
			case 1 %>
			Fradrag
			<% 
			case 2
			%>
			Tillæg
			<%
			case else
			end select %>
			</td>
			<%end if %>
			</tr>
			
			<%
	 
	 end if
	 
	 
	 
	 if dothis = 4 then
	 
	  if oRec4("aty_on_invoice") = 1 then
	  aty_sql_onfak = aty_sql_onfak & " OR aktiviteter.fakturerbar = "& oRec4("aty_id")
	  else
	  aty_sql_onfak = aty_sql_onfak 
	  end if
	 
	 
	 end if
	 
	 '******* Vælg akttyper joblog + bal_real_norm (medarbafstemning)
	  if dothis = 5 then
	  
	  if xi = 0 then  %>
	  
	  <!-- start tæller pga af trim felter -->
	  <input name="FM_akttype" type="hidden" value="#-99#" />
	  
	  <table cellspacing=0 cellpadding=1 border=0 width=100%>
	  
	        <%if thisfile = "bal_real_norm_2007.asp" then %>
			<tr>
			    <td><input id="chkalle_0" type="checkbox" class="akt_afst" /></td>
			    <td><b>Afstemning</b>
			    
			    </td>
			     <!-- <td><b>Time Regnsk.</b></td> -->
			</tr>
			
			
			
			<!-- faaste felter og sumtotaler -->
			
			<tr>
			<td style="border-bottom:1px #cccccc solid;" valign=top>
			
			<%
			'call erStempelurOn()
			
			if instr(akttype_sel, "#-5#") <> 0 AND stempelurOn = 1 then
            akttypeCHK = "CHECKED"
            else
            akttypeCHK = ""
            end if
			
			
			if stempelurOn = 1 then%>
			
            <input id="FM_akttype_id_0_0" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-5#" class="akt_afst" /></td>
			<td style="border-bottom:1px #cccccc solid;">
			Løntimer
			</td>
			
			<%end if %>
			
			</tr>
			
			<tr bgcolor="#EFF3ff">
			<td style="border-bottom:1px #cccccc solid;" valign=top>
			
			<%
			if instr(akttype_sel, "#-1#") <> 0 then
            akttypeCHK = "CHECKED"
            else
            akttypeCHK = ""
            end if
			%>
			
            <input id="FM_akttype_id_0_1" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-1#" class="akt_afst" /></td>
			<td style="border-bottom:1px #cccccc solid;">
			Vis sum-totaler på alle aktivitetstyper 
			</td>
			</tr>
			
			<tr>
			<td style="border-bottom:1px #cccccc solid;">
			
			<%
			if instr(akttype_sel, "#-2#") <> 0 then
            akttypeCHK = "CHECKED"
            else
            akttypeCHK = ""
            end if
			%>
			
            <input id="FM_akttype_id_0_2" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-2#" class="akt_afst" /></td>
			<td style="border-bottom:1px #cccccc solid;">
			Ressource timer
			</td>
			</tr>
			
			<tr bgcolor="#EFF3ff">
			<td style="border-bottom:1px #cccccc solid;">
			
			<%
			if instr(akttype_sel, "#-3#") <> 0 then
            akttypeCHK = "CHECKED"
            else
            akttypeCHK = ""
            end if
			%>
			
            <input id="FM_akttype_id_0_3" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-3#" class="akt_afst" /></td>
			<td style="border-bottom:1px #cccccc solid;">
			Faktureret timer
			</td>
			</tr>
			
			<tr>
			<td style="border-bottom:1px #cccccc solid;">
			
			<%
			if instr(akttype_sel, "#-4#") <> 0 then
            akttypeCHK = "CHECKED"
            else
            akttypeCHK = ""
            end if
			%>
			
            <input id="FM_akttype_id_0_4" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#-4#" class="akt_afst" /></td>
			<td style="border-bottom:1px #cccccc solid;">
			Mat. frb. / Udlæg
			</td>
			</tr>
			<%end if %>
			
	 <%
	 v = 5
	 end if 
	 
	  
	   if cint(left(oRec4("aty_sort"), 1)) <> cint(aty_lastsort) then ' AND cint(left(oRec4("aty_sort"), 1)) <> 2 then
	   %>
	   </table>
        <input id="antal_v_<%=aty_lastsort%>" type="hidden" value="<%=v %>" />
	    <%v = 0%>
	  </td>
	  <td valign=top style="padding:3px; width:125px; border:1px #D6DFf5 solid;">
	   
	  <table cellspacing=0 cellpadding=1 border=0 width=100%>
	   <%
			    select case left(oRec4("aty_sort"), 1)
			    'case 0
			    'straktgrpnavn = "Afstemning"
			    case 1
			    straktgrpnavn = "Udspecificering<br> atk. typer"
			    aktcls = "akt_udspec"
			    case 2
			    straktgrpnavn = "Udspecificering<br>akt. typer<br> (kat. 2)"
			    aktcls = "akt_flex"
			    case 3
			    straktgrpnavn = "Ferie"
			    aktcls = "akt_ferie"
			    case 4
			    straktgrpnavn = "Overarbejde / Afspadsering"
			    aktcls = "akt_overarb"
			    case 5
			    straktgrpnavn = "Sygdom / Fravær"
			    aktcls = "akt_syg"
			    case else
			    straktgrpnavn = ""
			    end select  %>
			<tr>
			    <td><input id="chkalle_<%=left(oRec4("aty_sort"), 1) %>" type="checkbox" class=<%=aktcls %> /></td>
			    <td><b><%=straktgrpnavn %></b></td>
			    <!-- <td><b>Time Regnsk.</b></td> -->
			</tr>
	 <%end if %>
	 
	  <%select case right(xi, 1)
	  case 2,4,6,8,0
	  bgthis = "#EFF3ff"
	  case else
	  bgthis = "#ffffff"
	  end select 
	  
	  
	  if instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	  akttypeCHK = "CHECKED"
	  else
	  akttypeCHK = ""
	  end if
	  %>
	        
	        <%
	        '**** Sygdom og sundhed kun admin ***'
	        if (oRec4("aty_id") = 20 OR oRec4("aty_id") = 21 OR oRec4("aty_id") = 8 OR oRec4("aty_id") = 81) AND level <> 1 AND thisfile = "bal_real_norm_2007.asp" then 
	        hdflt = 1
	        else
	        hdflt = 0
	        end if 
	        
	        if hdflt <> 1 then%>
	        
			<tr bgcolor="<%=bgthis%>">
			<td style="border-bottom:1px #cccccc solid;">
                <input id="FM_akttype_id_<%=left(oRec4("aty_sort"), 1)%>_<%=v%>" name="FM_akttype" type="checkbox" <%=akttypeCHK %> value="#<%=oRec4("aty_id")%>#" class=<%=aktcls %> /></td>
			<td style="border-bottom:1px #cccccc solid;">
			
			<%call akttyper(oRec4("aty_id"), 1) %>
			<%=akttypenavn%>
			</td>
			</tr>
			
			<%end if
	 
	 aty_lastsort = left(oRec4("aty_sort"), 1)
	 v = v + 1
	 end if
	 
	 '*** slut vælg akt typer **'
	 
	 
	 
	 
	 if dothis = 6 then
	 
	 aktiveTyper = aktiveTyper & "#"&oRec4("aty_id") & "#, "
	 
	 end if
	 
	 
	 '*** Medarbejder afstemning / Joblog **'
	 if dothis = 7 then
	 
	        
	        
	        '*** Tælles med i timereg **'
	        if oRec4("aty_on_realhours") = 1 AND instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_realhours = aty_sql_realhours &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realhours = aty_sql_realhours
	        end if
	        
	        
	        
	         
	        '*** Fakrurerbare **'
	        if oRec4("aty_on_invoiceble") = 1 AND instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_realHoursFakbar = aty_sql_realHoursFakbar &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realHoursFakbar = aty_sql_realHoursFakbar 
	        end if
	        
	      
	        '*** Ikke Fakturerbare **'
	        if oRec4("aty_on_invoiceble") = 2 AND instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_realHoursIkFakbar = aty_sql_realHoursIkFakbar &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_realHoursIkFakbar = aty_sql_realHoursIkFakbar
	        end if
	        
	        
	        '**** Fradrag fra løntimer ***'
	        if oRec4("aty_on_workhours") = 1 AND instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_frawhours = aty_sql_frawhours &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_frawhours = aty_sql_frawhours
	        end if
	        
	        '**** Tillæg fra løntimer ***'
	        if oRec4("aty_on_workhours") = 2 AND instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_tilwhours = aty_sql_tilwhours &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_tilwhours = aty_sql_tilwhours
	        end if
	        
	         '*** Aktive typer **'
	        if instr(akttype_sel, "#"&oRec4("aty_id")&"#") <> 0 then
	        aty_sql_sel = aty_sql_sel &" OR tfaktim = "& oRec4("aty_id")
	        else
	        aty_sql_sel = aty_sql_sel
	        end if
	        
	 end if
	 
	
	 
	 'aty_sql_real = aty_sql_real & " AND "
	 
	 
	 
	 xi = xi + 1
	 oRec4.movenext
	 wend
	 oRec4.close
	 
	 if dothis = 5 then
	 %>
	 <input id="antal_v_<%=aty_lastsort%>" type="hidden" value="<%=v %>" />
	 <%
	 end if
	 
	 if dothis = 3 then
	 %>
	 <tr>
	 <td colspan=5><br /><b>*)</b> Hvis man angiver den normerede uge som 37 timer excl. frokost skal typen <b>"frokost"</b> sættes til <b>ikke</b> at tælle med i det daglige timeregnskab (de realiserede timer pr. dag). I dette tilfælde sættes den til "Administrativ".<br /><br />
	 Typen <b>"Afspadsering"</b> sættes normalt til Ja, så den tæller med i den realiserede tid pr. uge, således at man ikke går i minus på den løbende fleks saldo ved at afspadsere. normalt kan man kun afspadsere optjent overarbejde. Afspadsering bliver automatisk altid modregnet typen "Overarbejde" <br /><br />
	 I det tilfælde hvor man har optjent et plus på norm/real tid saldoen bruges typen <b>"Fleks"</b> til at "nulle" saldoen. Derfor skal typen <b>"Fleks"</b> normalt ikke medregnes i de realiserede timer pr. dag.
	 </td>
	 </tr>
	 <%end if
	
	 
	end function
	
	
	
	
	public htmlparseTxt
	function htmlreplace(HTMLstring)
	
	
	    txtBRok = replace(HTMLstring, "<br>", "[#br#]") 
        txtBRok = replace(txtBRok, "<br />", "[#br#]") 
        txtBRok = replace(txtBRok, "<br/>", "[#br#]") 
        txtBRok = replace(txtBRok, "<b>", "[#b#]")
        txtBRok = replace(txtBRok, "</b>", "[#/b#]")
        
        txtBRok = replace(txtBRok, "<p>", "[#br#]")
        txtBRok = replace(txtBRok, "</p>", "[#br#]")
        
        txtBRok = replace(txtBRok, "<div>", "[#br#]")
        txtBRok = replace(txtBRok, "</div>", "[#br#]")
       
        txtBRok = replace(txtBRok, "<strong>", "[#strong#]")
        txtBRok = replace(txtBRok, "</strong>", "[#/strong#]")
       
        txtBRok = replace(txtBRok, "<i>", "[#i#]")
        txtBRok = replace(txtBRok, "</i>", "[#/i#]")
        
        txtBRok = replace(txtBRok, "<em>", "[#em#]")
        txtBRok = replace(txtBRok, "</em>", "[#/em#]")
       
        txtBRok = replace(txtBRok, "<u>", "[#u#]")
        txtBRok = replace(txtBRok, "</u>", "[#/u#]")
       
        
        HTMLstring = txtBRok
        
        
                Set RegularExpressionObject = New RegExp

                With RegularExpressionObject
                .Pattern = "<[^>]+>"
                .IgnoreCase = True
                .Global = True
                End With

                stripHTMLtags = RegularExpressionObject.Replace(HTMLstring, "")
                htmlparseTxt = replace(stripHTMLtags, "[#", "<")
                htmlparseTxt = replace(htmlparseTxt, "#]", ">")
                Set RegularExpressionObject = nothing
    	
    end function
    
   
 
 
    
 public utf_formatTxt
 function utf_format(utf_str)
            
            utf_str = replace(utf_str, "ø", "&#248;")
            utf_str = replace(utf_str, "æ", "&#230;")
            utf_str = replace(utf_str, "å", "&#229;")
            utf_str = replace(utf_str, "Ø", "&#216;")
            utf_str = replace(utf_str, "Æ", "&#198;")
            utf_str = replace(utf_str, "Å", "&#197;")
            utf_str = replace(utf_str, "Ö", "&#214;")
            utf_str = replace(utf_str, "ö", "&#246;")
            utf_str = replace(utf_str, "Ü", "&#220;")
            utf_str = replace(utf_str, "ü", "&#252;")
            utf_str = replace(utf_str, "Ä", "&#196;")
            utf_str = replace(utf_str, "ä", "&#228;")
            
            
            utf_formatTxt = utf_str
 
 end function
 
 public jq_formatTxt
 function jq_format(jq_str)
            
            jq_str = replace(jq_str, "ø", "&oslash;")
            jq_str = replace(jq_str, "æ", "&aelig;")
            jq_str = replace(jq_str, "å", "&aring;")
            jq_str = replace(jq_str, "Ø", "&Oslash;")
            jq_str = replace(jq_str, "Æ", "&AElig;")
            jq_str = replace(jq_str, "Å", "&Aring;")
            jq_str = replace(jq_str, "Ö", "&Ouml;")
            jq_str = replace(jq_str, "ö", "&ouml;")
            jq_str = replace(jq_str, "Ü", "&Uuml;")
            jq_str = replace(jq_str, "ü", "&uuml;")
            jq_str = replace(jq_str, "Ä", "&Auml;")
            jq_str = replace(jq_str, "ä", "&auml;")
            jq_str = replace(jq_str, "é", "&eacute;")
            jq_str = replace(jq_str, "É", "&Eacute;")
            jq_str = replace(jq_str, "á", "&aacute;")
            jq_str = replace(jq_str, "Á", "&Aacute;")
            
            jq_formatTxt = jq_str
 
 end function    
    
 public htmlparseCSVtxt
 Function htmlparseCSV(HTMLstring)
    
    
        Set RegularExpressionObject = New RegExp

        With RegularExpressionObject
        .Pattern = "<[^>]+>"
        .IgnoreCase = True
        .Global = True
        End With

        stripHTMLtags = RegularExpressionObject.Replace(HTMLstring, "")
        htmlparseCSVtxt = stripHTMLtags

        Set RegularExpressionObject = nothing

End Function






function opdaterFeriePl(level)

                '**** Ændrer status på planlagt ferie til afholdt ferie ***'
		        '*** (indlæser ny registrering så historik beholdes) ***'
		        '**** Tjekker om der findes reg. i forvejen *****'
		        '11 Ferie Planlagt
		        '14 Ferie Afholdt
		        
		        '18 Ferie Fridage Planlagt
		        '13 Ferie fridage brugt
		        
		        '*** Opdaterer for alle hvis det er en admin bruger der logger på ***'
		        
		        if level = 1 then
		        
		        for f = 1 to 2
		        
		        if f = 1 then
		        planlagtVal = 11
		        afholdtVal = 14
		        else
		        planlagtVal = 18
		        afholdtVal = 13
		        end if
		        
		        aktid = 0
		        LoginDato = year(now)&"/"& month(now)&"/"&day(now)
		        
		        '** Finder navn og id på afholdt ferie akt. ***'
		        strSQLfeafn = "SELECT a.id, a.navn FROM job j"_
		        &" LEFT JOIN aktiviteter a ON (a.fakturerbar = "& afholdtVal &" AND a.aktstatus = 1 AND a.job = j.id) "_
		        &" WHERE j.jobstatus = 1 AND a.id <> 'NULL' GROUP BY a.id ORDER BY a.id DESC"
		        
		        'AND a.id <> NULL GROUP BY a.id 
		        'Response.Write strSQLfeafn & "<br>"
		        'Response.flush
		        'Response.end
		        
		        oRec3.open strSQLfeafn, oCOnn, 3
		        if not oRec3.EOF then
		        
		        if oRec3("id") <> "" then
		        aktid = oRec3("id")
		        aktnavn = oRec3("navn")
		        end if
		        
		        end if
		        oRec3.close
		        
		        'Response.Write "<br>aktid" & aktid & "<br>"
		        
		        if cdbl(aktid) <> 0 then
		        
		        
		        '** Opdater ferie/feriefri for den bruger der logger på ***'
		        strSQLfepl = "SELECT * FROM timer WHERE tfaktim = "& planlagtVal &" AND tdato BETWEEN '2009-05-01' AND '"& LoginDato &"'"' AND tmnr = "& session("mid")
		        
		        'Response.Write strSQLfepl & "<br>"
		        
		        oRec4.open strSQLfepl, oConn, 3
		        while not oRec4.EOF 
		                
		                indtastningfindes = 0
		                strSQLfeaf = "SELECT timer, tdato FROM timer WHERE tfaktim = "& afholdtVal &" AND tdato = '"& year(oRec4("tdato")) &"/"& month(oRec4("tdato")) & "/"& day(oRec4("tdato")) &"' AND tmnr = "& oRec4("tmnr") 'session("mid")
		                
		                'Response.Write strSQLfeaf & "<br><br>"
		                
		                oRec3.open strSQLfeaf, oCOnn, 3
		                if not oRec3.EOF then
		                '*** ignorer da der allerede finsdes indtastning **'
		                indtastningfindes = 1
		                end if
		                oRec3.close
		                
		                'Response.Write indtastningfindes 
		                
		                
		                
		                if cint(indtastningfindes) = 0 then
		                '*** Indlæser afholdt ferie ***'
		                strSQLfeins = "INSERT INTO timer "_
		                &"("_
		                &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		                &" timerkom, TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		                &" editor, kostpris, offentlig, seraft, godkendtstatus, "_
		                &" godkendtstatusaf, sttid, sltid, valuta, kurs "_
		                &") "_
		                &" VALUES "_
		                &" (" _
		                & replace(oRec4("timer"), ",", ".") &", "& afholdtVal &", "_
		                & "'"& year(oRec4("tdato")) &"/"& month(oRec4("tdato")) & "/"& day(oRec4("tdato")) &"', "_
		                & "'"& oRec4("tmnavn") &"', "_
		                & oRec4("tmnr") &", "_
		                & "'"& oRec4("tjobnavn") &"', "_
		                & oRec4("tjobnr") &", "_
		                & "'"& oRec4("tknavn") &"', "_
		                & oRec4("tknr") &", "_
		                & "'"& replace(oRec4("timerkom"), "'", "''") &"', "_
		                & aktid &", "_
		                & "'"& aktnavn &"', "_
		                & oRec4("Taar") &", "_
		                & oRec4("TimePris") &", "_
		                & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', "_
		                & oRec4("fastpris") &", "_
		                & "'"& time &"', "_
		                & "'"& oRec4("editor") &"', "_
		                & replace(oRec4("kostpris"), ",", ".") &", "_
		                & oRec4("offentlig") &", "_
		                & oRec4("seraft") &", "_
		                & oRec4("godkendtstatus") &", "_
		                & "'"& oRec4("godkendtstatusaf") &"', "_
		                & "'"& oRec4("sttid") &"', "_
		                & "'"& oRec4("sltid") &"', "_
		                & oRec4("valuta") &", "_
		                & replace(oRec4("kurs"), ",", ".") &")"
		                
		                
		                'strSQLfeins = "INSERT LOW_PRIORITY INTO timer t1 SELECT * FROM timer t2 WHERE t2.tid = " & oRec4("tid")
		                
		                
		                'Response.Write strSQLfeins & "<br>"
		                'Response.end
		                oConn.execute(strSQLfeins)
		                
		                '*** Sletter den planlagte **'
		                strSQLdel = "DELETE FROM timer WHERE tid = "& oRec4("tid")
		                oConn.execute(strSQLdel)
		                
		                end if
		        
		                
		        
		        
		        oRec4.movenext
		        wend
		        oRec4.close
                
                end if '*** aktid <> 0
                
                next
                
                end if
                
                'Response.end

end function


public SY_usejoborakt_tp, SY_fastpris
function skaljobSync(jobid)

            '** Skal job sync ****'
			SY_usejoborakt_tp = 0
			SY_fastpris = 0
			strSQLj = "SELECT jobnavn, usejoborakt_tp, fastpris FROM job WHERE id = " & jobid
			
			'Response.Write strSQLj
			'Response.end
			oRec4.open strSQLj, oConn, 3
			if not oRec4.EOF then
			
			SY_usejoborakt_tp = oRec4("usejoborakt_tp")
			SY_fastpris = oRec4("fastpris")
			
			end if
			oRec4.close
			
			

end function

function syncJob(jobid)

                call akttyper2009(2)
                 			
			                strSQLaktSum = "SELECT SUM(budgettimer) sumakttimer, fakturerbar, SUM(aktbudgetsum) AS sumaktbudget FROM aktiviteter "_
			                &" WHERE job =  "& jobid & " AND("& aty_sql_fakbar &") AND aktfavorit = 0 GROUP BY job"
			                oRec2.open strSQLaktSum, oConn, 3
			                if not oRec2.EOF then
                			
			                sumakttimer = replace(oRec2("sumakttimer"), ",", ".")
			                sumaktbudget = replace(oRec2("sumaktbudget"), ",", ".")
                			
			                end if
			                oRec2.close
				
				strSQLsync = "UPDATE Job SET budgettimer = "& sumakttimer &", "_
				&" ikkebudgettimer = 0, jobtpris = "& sumaktbudget &" WHERE id = "& jobid
				
				oConn.execute(strSQLsync)
				
				'Response.Write strSQLsync
				'Response.Write " -- her"
				'Response.end

end function


'option explicit 

' Simple functions to convert the first 256 characters 
' of the Windows character set from and to UTF-8.

' Written by Hans Kalle for Fisz
' http://www.fisz.nl

'IsValidUTF8
'  Tells if the string is valid UTF-8 encoded
'Returns:
'  true (valid UTF-8)
'  false (invalid UTF-8 or not UTF-8 encoded string)
function IsValidUTF8(s)
  dim i
  dim c
  dim n

  IsValidUTF8 = false
  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      select case n
      case 1
        exit function
      case 2
        if (c and &HE0) <> &HC0 then
          exit function
        end if
      case 3
        if (c and &HF0) <> &HE0 then
          exit function
        end if
      case 4
        if (c and &HF8) <> &HF0 then
          exit function
        end if
      case else
        exit function
      end select
      i = i + n
    else
      i = i + 1
    end if
  loop
  IsValidUTF8 = true 
end function

'DecodeUTF8
'  Decodes a UTF-8 string to the Windows character set
'  Non-convertable characters are replace by an upside
'  down question mark.
'Returns:
'  A Windows string
function DecodeUTF8(s)
  dim i
  dim c
  dim n

  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      if n = 2 and ((c and &HE0) = &HC0) then
        c = asc(mid(s,i+1,1)) + &H40 * (c and &H01)
      else
        c = 191 
      end if
      s = left(s,i-1) + chr(c) + mid(s,i+n)
    end if
    i = i + 1
  loop
  DecodeUTF8 = s 
end function

'EncodeUTF8
'  Encodes a Windows string in UTF-8
'Returns:
'  A UTF-8 encoded string
function EncodeUTF8(s)
  dim i
  dim c

  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c >= &H80 then
      s = left(s,i-1) + chr(&HC2 + ((c and &H40) / &H40)) + chr(c and &HBF) + mid(s,i+1)
      i = i + 1
    end if
    i = i + 1
  loop
  EncodeUTF8 = s 
end function


public strDay_30
function dato_30(dagDato, mdDato, aarDato)

                if dagDato > 28 then 
				select case mdDato
				case "2"
				    
				    if len(trim(aarDato)) = 2 then
				    aarDato = "20" & aarDato
				    else
				    aarDato = aarDato
				    end if
				    
				    select case aarDato
				    case "2000", "2004", "2008", "2012", "2016", "2020", "2024", "2028", "2032", "2036", "2040", "2044"
				    strDay_30 = 29
				    case else
				    strDay_30 = 28
				    end select
				    
				case "4", "6", "9", "11"
				    if dagDato > 30 then
				    strDay_30 = 30
				    else
				    strDay_30 = dagDato
				    end if
				case else
				strDay_30 = dagDato
				end select
				else
				strDay_30 = dagDato
				end if

end function
%>





  
 
 