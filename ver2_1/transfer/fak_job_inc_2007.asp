
 <!-- Aktiviteter -->
	    


 <div id="aktdiv" style="position:absolute; visibility:hidden; display:none; z-index:2000; left:5px; top:105px; width:700px; border:1px yellowgreen solid; padding:5px 5px 5px 5px; background-color:#ffffff;">
    <table width=100% border=0 cellspacing=0 cellpadding=5>
	<tr>
	<td valign=top><h4>Aktiviteter (fakturalinier) <span style="font-size:11px; font-weight:normal;"> - kun aktive, med mindre der findes timer på en lukket/passiv aktivitet</span></h4>
	
	
      

<%
    '** Tjekker om der ligger timer på lukkede/passive aktiviteter ***
    '** Hvis ja vises lukkede og passive ellers vises kun åbne ***'
    strSQL2 = "SELECT aktiviteter.navn, aktiviteter.id, aktstatus, sum(t.timer) AS timerthis FROM aktiviteter "_
    &" LEFT JOIN timer AS t ON (t.taktivitetid = aktiviteter.id) "_
    &" WHERE aktstatus <> 1 AND (aktiviteter.job = "& jobid &") AND ("& aty_sql_onfak &") AND (t.tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"') GROUP BY t.taktivitetid"

    'Response.Write strSQL2
    'Response.flush
    timerPaLukkeogPassiveAkt = 0
    oRec2.open strSQL2, oConn, 3
    timerPaLukkeogPassiveAkt = 0                						
    while not oRec2.EOF 
	timerPaLukkeogPassiveAkt = timerPaLukkeogPassiveAkt + oRec2("timerthis")   
    oRec2.movenext
    wend 
    oRec2.close
				
                if cdbl(timerPaLukkeogPassiveAkt) <> 0 then	
                %>
                <span style="background-color:lightpink; width:500px; padding:3px;">Der er fundet en eller flere lukkede/passive aktiviteter med timer 
                på i den valgte periode, derfor vises lukkede og passive aktiviteter på denne faktura oprettelse.</span><br /><br />
                <%
                end if	                

   
	
	'******************************************************************************************'
	'*** Henter Medarbejdere og Aktiviteter ***************************************************'
	'******************************************************************************************'
	        
	        
	        isaktwritten = " aktiviteter.id <> 0 "
	        x = 0
	        
	        
	        if func = "red" then
        		
		        isTPused = "#"
        		
		        '*******************************************************************************'
		        '******************* Henter oprettede aktiviteter ******************************'
		        '*******************************************************************************'
		        strSQL = "SELECT faktura_det.id, aktid, aktiviteter.navn, aktiviteter.fakturerbar, aktiviteter.faktor, "_
		        &" faktura_det.valuta, faktura_det.beskrivelse, fak_sortorder, enhedsang, faktura_det.fase AS fase, momsfri, "_
		        &" aktiviteter.bgr, aktiviteter.aktbudget, aktiviteter.aktbudgetsum, aktiviteter.budgettimer, aktiviteter.antalstk"_
			    &" FROM faktura_det "_
		        &" LEFT JOIN aktiviteter ON (aktiviteter.id = aktid) WHERE fakid = "& intFakid &" GROUP BY aktid ORDER BY fase " 
		        
		        'Response.Write strSQL
		        'Response.flush
		        
		        oRec.open strSQL, oConn, 3
		       
		        While not oRec.EOF 
        		
			       thisaktnavn(x) = oRec("navn") 'oRec("beskrivelse")
			       thisaktid(x) = oRec("aktid")
			       thisaktfunc(x) = "red"
			       thisaktfakbar(x) = oRec("fakturerbar")
			       thisaktfaktor(x) = oRec("faktor")
			       thisAktText(x) = oRec("beskrivelse") 
			       thisAktValuta(x) = oRec("valuta")
			       thisaktsort(x) = oRec("fak_sortorder") 
			       thisaktEnh(x) = oRec("enhedsang")
			       thisaktCHK(x) = 1     
			       thisAktTimer(x) = 0
			       thisAktFase(x) = oRec("fase")
			       thisMomsfri(x) = oRec("momsfri")
			       
			       thisAktBudgetsum(x) = oRec("aktbudgetsum")
			       thisAktForkalk(x) = oRec("budgettimer")
				   thisAktForkalkStk(x) = oRec("antalstk")
				   
				   thisAktBgr(x) = oRec("bgr") 
				   select case thisAktBgr(x)
					case 0
					aktbgNavn = "Ingen"
					case 1
					aktbgNavn = "Timer"
					case 2
					aktbgNavn = "Stk."
					end select
			        
			        
			        '*** Antal timer total på denne akt. ***'
			        strSQLfantal = "SELECT sum(faktura_det.antal) AS antal, sum(faktura_det.aktpris) AS aktprissum FROM faktura_det "_
			        &" WHERE fakid = "& intFakid &" AND aktid = "& oRec("aktid") & " GROUP BY aktid"
			        oRec2.open strSQLfantal, oConn, 3
			        while not oRec2.EOF
			            thisAktTimer(x) = oRec2("antal")
			            thisAktBeloeb(x) = oRec2("aktprissum")
			        oRec2.movenext
			        wend
			        oRec2.close
			       
			       
        			
        		
        		isaktwritten = isaktwritten & " AND aktiviteter.id <> "& thisaktid(x) &"" 
        		
        		select case right(x, 1)
        		case 2,4,6,8,0
        		trbg = "#FFFFFF"
        		case else
        		trbg = "#FFFFFF"
        		end select
				
				
				    '** Marker lbn / fastpris kolonne
	                if cint(usefastpris) <> 1 then
	                bgfastpr = ""
	                bdfastim = 0
	                bglbntim = ""
	                bdlbntim = 1
	                else
	                bgfastpr = ""
	                bdfastim = 1
	                bglbntim = ""
	                bdlbntim = 0
	                end if
				
				
				if lcase(lastFase) <> lcase(thisAktFase(x)) AND len(trim(thisAktFase(x))) <> 0 then
				strAktAnchor = strAktAnchor & "<tr bgcolor=""#D6Dff5""><td colspan=8 class=lille>"
				strAktAnchor = strAktAnchor & "<b>"& replace(thisAktFase(x), "_", " ") &"</b></td></tr>"
				end if
				
				'strAktAnchor = strAktAnchor & "<tr class=""traktlink"" id=""tr_akt_"&thisaktid(x)&""" bgcolor="& trbg  &"><td style=""border-bottom:1px #cccccc solid;"" class=lille>"
				'strAktAnchor = strAktAnchor & "<a id='"& thisaktid(x) &"' href='#"& thisaktid(x) &"' class='qlinks'>"& thisaktnavn(x) & ": </a></td><td style=""border-bottom:1px #cccccc solid;"" align=right class=lille align=right>"& formatnumber(thisAktTimer(x), 2) &""_
				'&"</td><td style=""border-bottom:1px #cccccc solid;"" id=timeae_akt_"&thisaktid(x)&" align=right class=lille>&nbsp;</td></tr>"
				
				
				    strAktAnchor = strAktAnchor & "<tr class=""traktlink"" id=""tr_akt_"&thisaktid(x)&""" bgcolor="& trbg  &"><td class=lille style=""border-bottom:1px #cccccc solid;"">"
					strAktAnchor = strAktAnchor & "<a id='"& thisaktid(x) &"' href='#"& thisaktid(x) &"' class='qlinks'>"& thisaktnavn(x) & "</a></td>"_
					&"<td class=lille style=""border-bottom:1px #cccccc solid;"">"& aktbgNavn &"</td>"_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid;"">"& formatnumber(thisAktForkalkStk(x), 2) &"</td>"_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid;"">"& formatnumber(thisAktForkalk(x), 2) &"</td>"_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid; border-left:"&bdfastim&" solid #999999; border-right:"&bdfastim&" solid #999999; background-color:"& bgfastpr&";"">"& formatnumber(thisAktBudgetsum(x), 2) &"</td>"_
					&"<td style=""border-bottom:1px #cccccc solid;"" align=right class=lille>"& formatnumber(thisAktTimer(x), 2) &""_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid; border-left:"&bdlbntim&" solid #999999; border-right:"&bdlbntim&" solid #999999; background-color:#DCF5BD;"">"& formatnumber(thisAktBeloeb(x), 2) &"</td>"_
					&"</td><td align=right id=timeae_akt_"&thisaktid(x)&" class=lille style=""border-bottom:1px #cccccc solid;"">&nbsp;</td></tr>"
				
						
        		lastFase = oRec("fase")	
		        x = x + 1
        		
		        oRec.movenext
		        wend 
		        oRec.close
        	
	        end if
	        
	        
	        
	        
	        
	        
			
		    '*******************************************************************************'
		    '*** Henter alle de aktiviteter der ikke allerede er oprettet (ved rediger) ****'
		    '*** Henter alle aktiviteter ved opr. ******************************************'
		    '*******************************************************************************'
			
			'*** Henter fakbare, Ikke fakbare & Salg **'
			
			'**** Sidste faktura tidspunkt ****
			if len(strfaktidspkt) > 0 then
			LastFakTid = datepart("h", strfaktidspkt)&":"& datepart("n", strfaktidspkt)&":"& datepart("s", strfaktidspkt)
			else
			LastFakTid =  "00:00:02"
			end if
			
			
			'*** Henter aktiviter og finder registrede timer **'
			'*** siden sidste faktura dato. (periodeafgrænsning) ***'
			
            'vis lukkede og passive
            if cdbl(timerPaLukkeogPassiveAkt) <> 0 then
            sqlAktStatusKri = " aktstatus <> -1" 
			else
            sqlAktStatusKri = " aktstatus = 1"
            end if
			
			strSQL = "SELECT job.fastpris, job.jobtpris, job.budgettimer AS jbudgettimer, job.usejoborakt_tp, "_
			&" aktstatus, aktiviteter.job, aktiviteter.id AS aid, "_
			&" aktiviteter.navn AS anavn, aktiviteter.fakturerbar, aktiviteter.faktor, aktiviteter.fase, aktiviteter.aktstatus,  "_
			&" aty_on_invoice_chk, aty_enh, "_
			&" aktiviteter.bgr, aktiviteter.aktbudget, aktiviteter.aktbudgetsum, aktiviteter.budgettimer, aktiviteter.antalstk, aktiviteter.sortorder"_
			&" FROM aktiviteter LEFT JOIN job ON (job.id = aktiviteter.job) "_
			&" LEFT JOIN akt_typer aty ON (aty_id = aktiviteter.fakturerbar) "_
			&" WHERE (aktiviteter.job = "& jobid &") AND ("& sqlAktStatusKri &") AND ("& aty_sql_onfak &") AND ("& isaktwritten &") "_
			&" ORDER BY fase, aktiviteter.sortorder, aktiviteter.id"	
			'AND aktiviteter.aktstatus = 1, '"& aty_sql_onfak &"
			
			'Response.Write "<br>"& strSQL 
			'Response.flush
			
			oRec.open strSQL, oConn, 3
			
			
			lastAktId = 0
			timerThisT = 0
			timerThisD = 0
			'strKom = "Netto Kontant 8 dage." Hentes på kunde
			totaltimer = 0
			lastaktid = 0
            ir = 0
			
			
			while not oRec.EOF
			
			thisJoborAkt_tp(x) = oRec("usejoborakt_tp")	    
				    
				   if cint(lastaktid) <> cint(oRec("aid")) then
						
						
						thisaktid(x) = oRec("aid")
					    usefastpris = oRec("fastpris")
						
						    '** timepris på fastpris job ***'
							if cint(oRec("fastpris")) = 1 then
							
							    if cint(thisJoborAkt_tp(x)) = 0 then 'job
							    
							        if (oRec("jobtpris") <> 0 AND oRec("jbudgettimer")) <> 0 then
								    akttimepris = oRec("jobtpris")/oRec("jbudgettimer")
								    else
								    akttimepris = 0
								    end if
								    
								    
								    'timerThisD = oRec("budgettimer")
								    
								else
								    
								    '** Brug aktivteter ***'
								    if len(trim(oRec("aktbudget"))) <> 0 then
								    akttimepris = oRec("aktbudget")
								    else
								    akttimepris = 0
								    end if
								    
								    '*** Akt. grundlag ***'
								    '*** Timer eller ikke angivet
								    'if cint(oRec("bgr")) = 0 OR cint(oRec("bgr")) = 1 then
								    'timerThisD = oRec("budgettimer")
								    'else
								    'timerThisD = oRec("antalstk")
								    'end if
								
								end if
							
							'else	
								
							end if
								        
								        
								       
							            thisAktBeloeb(x) = 0
							            
							            '*** Skal ikke hentes ved kreditnota **'
							            if cint(intType) <> 1 then
						               
						                    '*** Realisrede timer ***'
						                    '** antal registerede timer pr. akt. **'
						                    strSQL2 = "SELECT sum(t.timer) AS timerthis, sum(t.timer*t.timepris) AS beloeb FROM timer AS t "_
						                    &" WHERE t.taktivitetid =" & oRec("aid") &" AND (t.tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"') GROUP BY t.taktivitetid"

                                            'Response.Write strSQL2
                                            'Response.flush
						                    oRec2.open strSQL2, oConn, 3
                    						
						                    if not oRec2.EOF then
							                    timerThisD = oRec2("timerthis")
							                    thisAktBeloeb(x) = oRec2("beloeb")
						                    end if
						                    oRec2.close
						                
						                end if
                						
							                if len(timerThisD) <> 0 then
							                timerThisD = timerThisD 
							                else
							                timerThisD = 0
							                end if
							
							
							
							
					
					
					thisTimePris(x) = akttimepris	
						
					thisaktnavn(x) = oRec("anavn") 
					thisAktText(x) = thisaktnavn(x)
					
					
					'** Hvis fastpris og job er grundlag strater aktviteter altid med 0 **'
                    '** Det giver forvirring, derfor vises real. timer, akt. linier er ikke tilvalgt hvis ikke der er forklakuleret på dem **'
					'if cint(thisJoborAkt_tp(x)) = 0 AND cint(oRec("fastpris")) = 1 then 'job
					'thisAktTimer(x) = 0
					'else
					thisAktTimer(x) = timerThisD
					'end if
					
					thisAktBudgetsum(x) = oRec("aktbudgetsum")
					thisaktfunc(x) = "opr"
					thisaktfakbar(x) = oRec("fakturerbar")
					thisaktfaktor(x) = oRec("faktor") 
					thisAktValuta(x) = valuta 'fra job
					thisaktEnh(x) = oRec("aty_enh")
					thisaktCHK(x) = oRec("aty_on_invoice_chk")
					
					timerThisD = 0
					thisBeloebD = 0
					
					thisAktFase(x) = oRec("fase")
					
					if thisaktfakbar(x) <> 5 then '** Kørsel
					thisMomsfri(x) = 0
					else
					thisMomsfri(x) = 1
					end if
					
					thisAktForkalk(x) = oRec("budgettimer")
					thisAktForkalkStk(x) = oRec("antalstk")
					thisAktBgr(x) = oRec("bgr") 
					
					select case thisAktBgr(x)
					case 0
					aktbgNavn = "Ingen"
					case 1
					aktbgNavn = "Timer"
					case 2
					aktbgNavn = "Stk."
					end select
					
					select case right(x, 1)
        		    case 2,4,6,8,0
        		    trbg = "#FFFFFF"
        		    case else
        		    trbg = "#FFFFFF"
        		    end select
        		    
        		    '** Maker lbn / fastpris kolonne
	                if cint(usefastpris) <> 1 then
	                bgfastpr = ""
	                bdfastim = 0
	                bglbntim = ""
	                bdlbntim = 1
	                else
	                bgfastpr = ""
	                bdfastim = 1
	                bglbntim = ""
	                bdlbntim = 0
	                end if

                    if ir = 0 AND func = "red" then
                    strAktAnchor = strAktAnchor & "<tr bgcolor=""#FFFFFF""><td colspan=8 class=lille>"
				    strAktAnchor = strAktAnchor & "<br><b>Aktiviteter ikke med på faktura:</b></td></tr>"
                    end if
				
					if (lcase(lastFase) <> lcase(thisAktFase(x)) AND len(trim(thisAktFase(x))) <> 0) OR (len(trim(lastFase)) = 0 AND len(trim(thisAktFase(x))) <> 0)  then
				    strAktAnchor = strAktAnchor & "<tr bgcolor=""#D6Dff5""><td colspan=8 class=lille>"
				    strAktAnchor = strAktAnchor & "<b>"& replace(thisAktFase(x), "_", " ") &"</b></td></tr>"

                    thisaktsort(x) = 100+x 'oRec("sortorder")

                    else

                    if lastSort <> "" then
                    thisaktsort(x) = lastSort
                    else
                    thisaktsort(x) = 1
                    end if
					
				    end if
					

                    select case oRec("aktstatus") 
                    case 0
                    aktstTXT = " (Lukket)"
                    case 2
                    aktstTXT = " (Passiv)"
                    case else
                    aktstTXT = "" 
                    end select

					strAktAnchor = strAktAnchor & "<tr class=""traktlink"" id=""tr_akt_"&thisaktid(x)&""" bgcolor="& trbg  &"><td class=lille style=""border-bottom:1px #cccccc solid;"">"
					strAktAnchor = strAktAnchor & "<a id='"& thisaktid(x) &"' href='#"& thisaktid(x) &"' class='qlinks'>"& thisaktnavn(x) & "</a>"& aktstTXT &"</td>"_
					&"<td class=lille style=""border-bottom:1px #cccccc solid;"">"& aktbgNavn &"</td>"_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid;"">"& formatnumber(thisAktForkalkStk(x), 2) &"</td>"_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid;"">"& formatnumber(thisAktForkalk(x), 2) &"</td>"_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid; border-left:"&bdfastim&" solid #999999; border-right:"&bdfastim&" solid #999999; background-color:"& bgfastpr&";"">"& formatnumber(thisAktBudgetsum(x), 2) &"</td>"_
					&"<td style=""border-bottom:1px #cccccc solid;"" align=right class=lille>"& formatnumber(thisAktTimer(x), 2) &""_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid; border-left:"&bdlbntim&" solid #999999; border-right:"&bdlbntim&" solid #999999; background-color:"& bglbntim&";"">"& formatnumber(thisAktBeloeb(x), 2) &"</td>"_
					&"</td><td align=right id=timeae_akt_"&thisaktid(x)&" class=lille style=""border-bottom:1px #cccccc solid;"">&nbsp;</td></tr>"
				
				
                if len(trim(thisAktFase(x))) <> 0 then	
			    lastFase = thisAktFase(x) 'oRec("fase")	
                else
                lastFase = ""
                end if

                lastSort = thisaktsort(x)
                
                ir = ir + 1
				x = x + 1		
				end if
			
          

			lastaktid = oRec("aid")		
			oRec.movenext
			wend
			oRec.close
			
			
	'end if
	'***********************************************************************************'
	
	
	'Response.flush
	
	
	'***********************************************************************************'
	'********************* Udskriver aktiviterer ***************************************'
	'***********************************************************************************'
	uTxt = "<font class=lillesort>Ved <b>fastpris job</b> er samle sum-aktiviteten forvalgt. Hvis job er grundlag for beregning af timepris, starter hver <b>samle sum-akt. med 0 timer</b> og ellers med det forkalkulerede på hver aktivitet.<br><br>"_
	&"Ved <b>lbn. timer job</b> er realiserede timer og omsætning forvalgt på faktura linier. <br>Hvis grundlag på aktivitet er sat til stk, er samle sum-aktiviteten forvalgt og <b>enhedsprisen beregnet udfra timeforbrug.</b></font>"
	uWdt = 500
	call infoUnisport(uWdt, uTxt)
	
	if cint(hidesumaktlinier) = 1 then
	hidesumaktlinierCHK = "CHECKED"
	else
	hidesumaktlinierCHK = ""
	end if
	
    if cint(hidefasesum) = 1 then
	hidefasesumCHK = "CHECKED"
	else
	hidefasesumCHK = ""
	end if

	%>
	
	 <br /><input id="FM_hidesumaktlinier" name="FM_hidesumaktlinier" value="1" type="checkbox" <%=hidesumaktlinierCHK %> /> Vis ikke udspecificering af aktivitets- og materiale -linier på faktura. (vis kun total beløb)
	 <br /><input id="FM_hidefasesum" name="FM_hidefasesum" value="1" type="checkbox" <%=hidefasesumCHK %> /> Vis ikke faseoverskift og mellemregning for hver fase på faktura.
	
	<br />
	<table width=100% cellspacing=0 cellpadding=2 border=0 bgcolor="#D6DFf5">
	<tr bgcolor="#FFFFFF">
	    <% if cint(usefastpris) <> 1 then%>
	     <td colspan=6><a name="top" class=vmenu><b>Links til aktiviteter og faser på denne faktura:</b> (klik for at redigere)</a></td>
	    <td class=lille style="border:1px solid #999999; border-bottom:0px; background-color:#D6Dff5;">Jobtype:<br /> <b>Lbn. timer</b></td>
	    <td>&nbsp;</td>
	 
	    <%else %>
	       <td colspan=4><a name="top" class=vmenu><b>Links til aktiviteter og faser på denne faktura:</b> (klik for at redigere)</a></td>
	    <td class=lille style="border:1px solid #999999; border-bottom:0px; background-color:#D6Dff5;">Jobtype:<br /><b>Fastpris</b></td>
	    <td colspan=3>&nbsp;</td>
	   
	    <%end if%>
	</tr>
	
	
	<tr bgcolor="#8cAAe6"><td style="width:250px;" class=lille><b>Faser og aktiviteter</b></td>
	<td class=lille style="width:75px;"><b>Ber. grundlag</b></td>
	<td align=right class=lille style="width:75px;"><b>Forkalk. stk.</b></td>
	<td align=right class=lille style="width:75px;"><b>Forkalk. tim.</b></td>
	<td align=right class=lille style="width:75px; padding-right:5px; background-color:<%=bgfastpr%>;"><b>Budget</b><br />(Fastpris)</td>
	<td align=right class=lille style="width:75px;">
	<b>Real. tim.</b> i periode
	<%if func = "red" then %>
	<br />/<b>Antal angivet på fak.</b>
	<%end if %>
	</td>
	
	<td align=right class=lille style="width:75px; background-color:<%=bglbntim%>;"><b>Omsætning</b><br />(Lbn. timer)
	<%if func = "red" then %>
	<br />/<b>Bel. angivet på fak.</b>
	<%end if %>
	</td>
	<td align=right class=lille style="width:75px;"><b>Ændring i antal</b><br /> denne faktura</td></tr>
	<%=strAktAnchor%>
	</table>
	<br /><br /><br />
	
	<%if fastpris = 1 then  %>
	<b>Dette job er et <u>fastpris job</u></b><br />
	
	    <%if usejoborakt_tp = 0 then %>
	    Det angivne timeantal og bruttoomsætning på jobbet er sat til at være grundlag for beregning af timepris på faktura.<br />
	    Budget / Pris: <b><%=formatnumber(intJobTpris, 2) &" "& valutaKode %></b><br />
        Fakturerbare timer forkalk: <b><%=formatnumber(intBudgettimer, 2) %></b> timer<br />
            <%
            if intBudgettimer <> 0 AND intJobTpris <> 0 then
            jobFastprisTp = (intJobTpris/intBudgettimer)
            else
            jobFastprisTp = 0
            end if
            %>
        Beregnet timepris: <b><%=formatnumber(jobFastprisTp) &" "& valutaKode%> </b> / time.
	    <%else %>
	    De forkalkulrede timer, antal og enhedspriser på hver aktivitet er sat til at danne grundlag for beregning af timepris / enhedspris på faktura.
	    <%end if %>
	<%else %>
	<b>Dette job er sat til <u>løbende timer</u></b><br />
	Timeprisen på den enkelte medarbejder danner grundlag for beregning. 
	<%end if %>
	
	
	
	</td>
	</tr>
	</table>
	
	
	
	
	
	<%

    '** Valuta_enheder Global **
    call enhval_gbl	
	
	
	
	intTimer = 0
	totalbelob = 0
	subtotalbelob = 0
	subtotaltimer = 0
	totalbelob_umoms = 0
	
	nialt = 0
	for x = 0 to x - 1 
	
	'if x = 0 then
    'sumDsp = ""
    'sumVzb = "visible"
	'else
    sumVzb = "hidden"
    sumDsp = "none"
    'end if
	
    if len(trim(thisAktFase(x))) <> 0 then
    thisAktFase(x) = thisAktFase(x)
    else
    thisAktFase(x) = ""
    end if
    
    %>
	
	
	
	
	
	
	
	
	
	<span class="sumaktdiv" id="sumakt_<%=thisaktid(x) %>" style="position:relative; display:<%=sumDsp%>; visibility:<%=sumVzb%>; padding:0px; border:0px #cccccc solid;">
	
	<table width=100% border=0 cellspacing=0 cellpadding=0>
	
	<tr><td colspan=2 valign=top style="padding:20px 5px 2px 5px; border-top:5px solid #Eff3FF;"><h4>

	<a id="Aname_<%=thisaktid(x)%>"><%=thisaktnavn(x)%></a> 
	<%
	call akttyper(thisaktfakbar(x), 1)
    %>
	
	&nbsp;<font class=roed>(<%=akttypenavn %>)</font>
	&nbsp;&nbsp;<a href="#" class="tiltoppen">^ Til toppen ^</a> 
	<!--&nbsp;&nbsp;<a href="#" id="<%=thisaktid(x)%>" class=naeste> Næste >></a> -->
	</h4>
	</td>
	</tr>
	<tr><td valign="top" style="padding:2px 20px 5px 0px;">
	Fase: <input type="text" name="FM_aktfase_<%=x%>" id="FM_aktfase_<%=x%>" value="<%=replace(thisAktFase(x), "_", " ") %>" style="font-size:9px; width:200px;" />&nbsp;&nbsp;
	Sortering: (hele fasen) <input type="text" name="aktsort_<%=x%>" id="aktsort_<%=x%>" value="<%=thisaktsort(x)%>" class="sortFase" style="font-size:9px; width:30px;">
    <input value="<%=lcase(trim(thisAktFase(x))) %>" id="fs_<%=x%>" type="hidden" />
	&nbsp;&nbsp;&nbsp;Faktor: <input type="text" name="aktfaktor_<%=x%>" id="aktfaktor_<%=x%>" value="<%=thisaktfaktor(x) %>" style="font-size:9px; width:30px;">
	<input id="beregnfaktor_<%=x%>" type="button" value="Beregn" style="font-size:9px;" onClick="opd_timer_faktor(<%=x%>)" /> (Undt. samle aktivitet) 
	
	</td>
	<td align=right style="padding:0px 20px 5px 0px;">
	<!--&nbsp;&nbsp;&nbsp;<a href="#" onclick="showlommeregner()">Lommerenger</a>-->
	<input name="subm_on_top_<%=x%>" id="subm_on_top_<%=x%>" type="submit" value=" Se faktura >> " style="font-size:9px;"/>
	</td></tr>
	
	</table>
	
	
	
	<%if thisaktfunc(x) <> "red" AND func = "red" then %>
	<table>
	<tr><td colspan=2 style="background-color:#ffff99; padding:5px;">
	<font class=roed>Denne aktivitet er ikke oprettet på faktura. # timer (opr. timeantal på aktivitet), er hentet fra timeregistrering i den valgte periode.<br />
	Timer til fakturering på medarbejder er sat til 0.</font>
	</td></tr>
	</table>
	<%end if%>
	
	
	
	
	<table cellspacing="0" cellpadding="1" border="0" width=100% bgcolor="silver">
	<%
	
	%>
	<input type="hidden" name="sumaktbesk_<%=x%>" id="sumaktbesk_<%=x%>" value="<%=thisaktnavn(x)%>">
	<%
		'if instr(aktiswritten, "#"&thisaktid(x)&"#") = 0 then
		n = 1
		a = 0
		newa = 0
		sa = 0
		sa_medlin = 0
	
					'if thisaktfunc(x) <> "red" then
					
					
							call projektgrupper()	
							
							
					'end if
					
					if thisaktid(x) <> 0 then%>
					<!--<tr bgcolor="#d6dff5">
						<td colspan=6><!--Deaktiverede medarbejdere vises ikke ved faktura oprettelse. pr. 29/8-2005.<img src="../ill/blank.gif" width="1" height="20" alt="" border="0"><br></td>
					</tr>-->
					
					<%
					if func = "red" then
					
					    'if thisaktfunc(x) = "red" then
					    bgthis = "#8CAAE6"
					    'else
					    'bgthis = "#cccccc"
					    'end if
					
					else
					    
					    bgthis = "#8CAAE6"
					    
					end if %>
					
					<tr bgcolor="<%=bgthis %>">
						<td>&nbsp;</td>
						<td class=lille style="padding:2px 2px 2px 5px;">Vis</td>
						<td class=lille style="width:70px;">Antal - #</td>
                        <td class=lille style="width:100px; padding:2px 2px 2px 5px;">Ventetimer <br />
                        <font class=megetlillesort>Saldo : Brugt : Ultimo</font></td>
						<td class=lille style="padding:2px 2px 2px 5px;">Medarbejder(e)</td>
						<td class=lille style="padding:2px 2px 2px 5px;">Enh.pris</td>
						<td class=lille style="padding:2px 2px 2px 5px;">Valuta</td>
						<td class=lille style="padding:2px 2px 2px 5px;">Enhed</td>
						<td class=lille style="padding:2px 2px 2px 5px;" align=center>Rabat</td>
						<td class=lille style="width:80px; padding:2px 2px 2px 5px;" align=right>Pris ialt&nbsp;&nbsp;&nbsp;</td>
						<td style="width:80px; padding:2px 2px 2px 5px;"><font class=megetlillesort>Afrund</td>
					</tr>
					<!--<tr bgcolor="#66CC33"><td colspan=6 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="2" alt="" border="0"><br></td></tr>-->
			
			
					<%end if
					
					
					
					
					'***************************************************'
					'Faktimer, Venter og Beløb på medarbejdere     *****'
					'***************************************************'
					
					faktot = 0
					medarbstrpris = ""
					
					
					
					'**** Henter alle allerede oprettede aktiviteter ***'
					if thisaktfunc(x) = "red" then
						
						strSQL10c = "SELECT fms.medrabat, fd.id, antal, "_
						&" fd.beskrivelse, aktpris, fd.showonfak AS akt_showonfak, "_
						&" fd.enhedspris, "_
						&" fd.aktid, "_
						&" fak, venter_brugt, fms.tekst, fms.enhedspris, beloeb, "_
						&" mid, fms.showonfak AS showonfak, venter AS venter_ultimo, "_
						&" fms.enhedsang, fms.valuta AS mvaluta, fms.kurs AS fmskurs "_
						&" FROM faktura_det fd "_
						& "LEFT JOIN fak_med_spec fms ON (fms.fakid = fd.fakid AND fms.aktid = fd.aktid)"_
						&" WHERE fd.fakid = "& intFakid &" AND fd.aktid = "& thisaktid(x) &" AND fms.fakid = "& intFakid &" AND fms.aktid = "& thisaktid(x) &" GROUP BY mid"
						
						
						
						    medidIsWrt = "AND (mid <> 0 "
					        medarbOptionList_0 = ""
					        medarbHiddenNameList_0 = ""
					        txt_0 = ""					
        					
					        '** Alle medarbejdere med timer på ***'
        						
					        oRec3.open strSQL10c, oConn, 3
					        while not oRec3.EOF
					        medarbOptionList_0 = "<option value="& oRec3("mid") &">"& oRec3("tekst") &"</option>"   		
					        medarbHiddenNameList_0 = "<input id=""a_"&thisaktid(x)&"_"& oRec3("mid") &""" value='"& oRec3("tekst") &"' type=""hidden"" />" 
        					txt_0 = oRec3("tekst") '(Mnavn)
        					
					        call fakturaMedarbLinjer(0)
					        nialt = nialt + 1
					        medidIsWrt = medidIsWrt & " AND mid <> "& oRec3("mid")
					        sa_medlin = sa_medlin + 1 
					        'Response.flush
					        oRec3.movenext
					        wend
					        oRec3.close
        					
        					
					        medidIsWrt = medidIsWrt & ")"
						
						
						
					else
					    
					'**** Ved red: Henter alle de aktiviteter der endnu ikke er oprettet ***'
                    '**** Ved opr: Hennter alle aktiviteter ********************************'
                    
                    
                   
				    
				    strSQL10a = "SELECT DISTINCT(tmnr) AS medarbejderid, mnavn, m.mid, "_
				    &" fms.venter_brugt, fms.venter AS venter_ultimo FROM timer "_
				    &" LEFT JOIN medarbejdere m ON (m.mid = tmnr) "_
				    &" LEFT JOIN fak_med_spec fms ON (fms.aktid = "& thisaktid(x) &" AND fms.mid = tmnr AND fms.fakid = "& lastFakid & ") "_
				    &" WHERE taktivitetid = "& thisaktid(x) &" AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"') GROUP BY tmnr"
				    
                   
                    medidIsWrt = "AND (mid <> 0 "
					medarbOptionList_0 = ""
					medarbHiddenNameList_0 = ""		
					txt_0 = ""			
					
					'** Alle medarbejdere med timer på ***'
					
					
					'Response.Write "<u>" & strSQL10a & "</u>"
					'Response.flush
						
					oRec3.open strSQL10a, oConn, 3
					while not oRec3.EOF
					medarbOptionList_0 = "<option value="& oRec3("medarbejderid") &">"& oRec3("mnavn") &"</option>"   		
					medarbHiddenNameList_0 = "<input id=""a_"&thisaktid(x)&"_"& oRec3("medarbejderid") &""" value='"& oRec3("mnavn") &"' type=""hidden"" />" 
					
					txt_0 = oRec3("mnavn")
					
					call fakturaMedarbLinjer(0)
					nialt = nialt + 1
					medidIsWrt = medidIsWrt & " AND mid <> "& oRec3("medarbejderid")
					sa_medlin = sa_medlin + 1
					'Response.flush
					oRec3.movenext
					wend
					oRec3.close
					
					
					medidIsWrt = medidIsWrt & ")"
                  
					
					end if
					
					
					
					'*****************************'
                    '*** VenteTime Saldo Primo ***'
                    '*****************************' 
                    '*** Er der medarbejdere med ventimer på skal de være aktive ved opret ***'
                    '*** Kun AKTIVTETER der ikke ALLEREDE ER MED PÅ FAK ***'
                    
                    if thisaktfunc(x) <> "red" then
                    'sqlFaknrKri =  " fid = "& id
                    'else
                    sqlFaknrKri =  " f.fid = "& lastFakid
                    'end if
                    
                    
					strSQL2 = "SELECT f.fid, fms.venter_brugt, fms.venter AS venter_ultimo, fms.mid AS medarbejderid, m.mid, m.mnavn FROM fakturaer f "_
					&" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid AND fms.aktid = "& thisaktid(x) &""_
					&" "& replace(medidIsWrt, "mid", "fms.mid") &")"_
					&" LEFT JOIN medarbejdere m ON (m.mid = fms.mid) "_
					&" WHERE "& sqlFaknrKri &" AND f.jobid = "& jobid &" GROUP BY fms.mid ORDER BY f.fakdato DESC, aktid" 
                    
                    'Response.Write "<b>" & strSQL2 & "</b>"
                    'Response.flush
                    
                    oRec3.open strSQL2, oConn, 3 
                    while not oRec3.EOF 
                        
                        if oRec3("venter_ultimo") <> 0 then
                        
                        medarbOptionList_0 = "<option value="& oRec3("medarbejderid") &">"& oRec3("mnavn") &"</option>"   		
					    medarbHiddenNameList_0 = "<input id=""a_"&thisaktid(x)&"_"& oRec3("medarbejderid") &""" value='"& oRec3("mnavn") &"' type=""hidden"" />" 
    					
					    txt_0 = oRec3("mnavn")
    					
					    call fakturaMedarbLinjer(0)
					    nialt = nialt + 1
					    medidIsWrt = medidIsWrt & " AND mid <> "& oRec3("medarbejderid")
					    sa_medlin = sa_medlin + 1
					    
					    end if
                        
                    oRec3.movenext
                    wend 
                    oRec3.close
                    
                    end if 'akt red
					'**** Ventimer ved opret faktura slut ***'
					
					
					
					'** Mulighed for tomme linier, med medarb, 
					'** der ikke allerede er skrevet
					'** (har timer eller VENTE timer på) ***'
					
					deakSQL = " AND mansat <> '2' AND mansat <> '3' " '** passive må IKKE vælges til ekstra linier, f.eks Epi med 350 medarb. bliver meget tung **'
                    pgrpSQL = "(projektgruppeid = "& gp1 &" OR projektgruppeid = "& gp2 &" OR projektgruppeid = "& gp3 &" OR projektgruppeid = "& gp4 &" "_
				    &" OR projektgruppeid = "& gp5 &" OR projektgruppeid = "& gp6 &" OR projektgruppeid = "& gp7 &" OR projektgruppeid = "& gp8 &" "_
				    &" OR projektgruppeid = "& gp9 &" OR projektgruppeid = "& gp10 &") "
				    
				    
				    usedmedabId = " Tmnr = 0 AND "
					
					
					'*** Opbygger optionlist **'
					medarbOptionList_1 = ""
					medarbHiddenNameList_1 = ""
					medarbOptionList_2 = ""
					medarbHiddenNameList_2 = ""
					medarbOptionList_3 = ""
					medarbHiddenNameList_3 = ""
					txt_1 = ""
					txt_2 = ""
					txt_3 = ""
					
					cn = 1
					countLimit = 0	
				    strSQL10b = "SELECT DISTINCT(medarbejderid), mnavn, mid FROM progrupperelationer, medarbejdere "_
				    &" WHERE mid = medarbejderid "& deakSQL &" AND "_
				    &""& pgrpSQL & medidIsWrt &""_
				    &" GROUP BY medarbejderid ORDER BY mnavn" 
					'& medidIsWrt
					'Response.Write "medidIsWrt" & medidIsWrt
					'Response.Write strSQL10b
					'Response.flush
						
					oRec3.open strSQL10b, oConn, 3
					while not oRec3.EOF
					'if countLimit = 0 then
					
					'end if
					
					opLSEL_1 = ""
					opLSEL_2 = ""
					opLSEL_3 = ""
					
					
					select case cn  
					case 1
					opLSEL_1 = "SELECTED"
					txt_1 = oRec3("mnavn")
					case 2
					opLSEL_2 = "SELECTED"
					txt_2 = oRec3("mnavn")
					case 3
					opLSEL_3 = "SELECTED"
					txt_3 = oRec3("mnavn")
					end select
					
					medarbOptionList_1 = medarbOptionList_1 & "<option value='"& oRec3("medarbejderid") &"' "&opLSEL_1&">"& oRec3("mnavn") &"</option>"   		
					medarbHiddenNameList_1 = medarbHiddenNameList_1 & "<input id=""a_"&thisaktid(x)&"_"& oRec3("medarbejderid") &""" value='"& oRec3("mnavn") &"' type=""hidden"" />" 
					
					medarbOptionList_2 = medarbOptionList_2 & "<option value='"& oRec3("medarbejderid") &"' "&opLSEL_2&">"& oRec3("mnavn") &"</option>"   		
					medarbHiddenNameList_2 = medarbHiddenNameList_2 & "<input id=""a_"&thisaktid(x)&"_"& oRec3("medarbejderid") &""" value='"& oRec3("mnavn") &"' type=""hidden"" />" 
					
					
					medarbOptionList_3 = medarbOptionList_3 & "<option value='"& oRec3("medarbejderid") &"' "&opLSEL_3&">"& oRec3("mnavn") &"</option>"   		
					medarbHiddenNameList_3 = medarbHiddenNameList_3 & "<input id=""a_"&thisaktid(x)&"_"& oRec3("medarbejderid") &""" value='"& oRec3("mnavn") &"' type=""hidden"" />" 
					if countLimit < 3 then
					sa_medlin = sa_medlin + 1
					end if
					
					countLimit = countLimit + 1
					cn = cn + 1
					'Response.flush
					oRec3.movenext
					wend
					oRec3.close
					
					if countLimit > 3 then
					countLimit = 3
					else
					countLimit = countLimit
					end if
					
					

					
					'*** Skriver 3 linier med medarbejdere der ikke er timer på **'	
					if countLimit <> 0 then
					%>
					<tr bgcolor="#FFFFFF"><td colspan=11 style="padding:10px;">
					    
					    <div style="color:darkred; border:1px #cccccc solid; padding:5px 5px 5px 5px;"><b>Kun medarbejder-linier</b> med tilsvarende timepris, som sum-aktiviteter der bliver vist på fakturaen, bliver indlæst til afstemning.<br />
                        Hvis <b>Sum-akt. Samle-linien er slået til</b> bliver alle medarb. linier indlæst.</div>
					    <br />
					    <b>Der kan tilføjes optil 3 ekstra medarbejder linier her:</b><br>Vælg mellem tilknyttede (via deres projektgrupper) <b>aktive</b> medarbejdere uden timeforbrug i den valgte periode/på valgte faktura.
                        
					   
	                   
					
					
					</td></tr>
					
					<%
					end if%>
					<%
				    strSQL10 = "SELECT DISTINCT(medarbejderid), mnavn, mid FROM progrupperelationer, medarbejdere "_
				    &" WHERE mid = medarbejderid "& deakSQL &" AND "_
				    &""& pgrpSQL & medidIsWrt &""_
				    &" GROUP BY medarbejderid ORDER BY mnavn LIMIT 0,"& countLimit 
					
					'Response.Write strSQL10
					'Response.flush
					cn = 1	
					oRec3.open strSQL10, oConn, 3
					while not oRec3.EOF
					
					call fakturaMedarbLinjer(cn)
					nialt = nialt + 1
					
					cn = cn + 1		
					'Response.flush
					oRec3.movenext
					wend
					oRec3.close
					
					%>
					
					</table>
					<br /><br />
					
					<!-- sum akt. table -->
					<table cellspacing="0" cellpadding="0" border="0" width="100%">
					
                    <%
                    
                    
                    call subtotalakt(x, 0)
                    
                    faktot = 0
					mbelobtot = 0
			
			        Response.write strAktsubtotal
			        strAktsubtotal = ""
			
			    if cint(thisMomsfri(x)) = 1 then
    	        subtotalbelobUmoms = subtotalbelob
    	        else
    	        subtotalbelobUmoms = 0
    	        end if
    	        
    	        
    	        '***** Subtotaler på sum-aktiviteter ***'
    	        Response.Write "<input type=""hidden"" name='timer_subtotal_akt_"&x&"' id='timer_subtotal_akt_"&x&"' value='"&subtotaltimer&"'>"_
                &"<input type=""hidden"" name='belob_subtotal_akt_"&x&"' id='belob_subtotal_akt_"&x&"' value='"&subtotalbelob&"'>"_
                &"<input type=""hidden"" name='belob_subtotal_akt_umoms_"&x&"' id='belob_subtotal_akt_umoms_"&x&"' value='"&subtotalbelobUmoms&"'>"_ 
			    &"<input type=""hidden"" name='txt_subtotal_akt_"&x&"' id='txt_subtotal_akt_"&x&"' value='"& thisSamleaktTxt &"'>" 
    			
    			
    			totalbelob_umoms = totalbelob_umoms + subtotalbelobUmoms
			
			
			%>
			<tr>
			    <td bgcolor="#D6Dff5" style="height:4px;"><img src="../ill/blank.gif" border=0 width=1 height=1 /></td>
			</tr>
			
			
			
			
			
		
            <input type="hidden" name="antal_subtotal_akt_<%=x%>" id="antal_subtotal_akt_<%=x%>" value="<%=sa%>">		   
			
			<input type="hidden" name="aktId_n_<%=x%>" id="aktId_n_<%=x%>" value="<%=thisaktid(x)%>">
			<input type="hidden" name="antal_n_<%=x%>" id="antal_n_<%=x%>" value="<%=n%>">
			<input type="hidden" name="highest_aval_<%=x%>" id="highest_aval_<%=x%>" value="<%=sa%>">
			<input type="hidden" name="highest_aval_m_<%=x%>" id="highest_aval_m_<%=x%>" value="<%=sa_medlin%>">
			 
			<% 
		subtotaltimer = 0 
		subtotalbelob = 0 
		'aktiswritten = aktiswritten & ",#"&thisaktid(x)&"#"		
		'end if	
		%></table>
		
		</span><!-- sum akt expand -->
		
		<%
		
		
    'Response.flush
    
    
	next
	
	
	%>
	
	
	
	
	
	<input type="hidden" name="lastactive_x" id="lastactive_x" value="0">
	<input type="hidden" name="lastactive_a" id="lastactive_a" value="0">
	<%
	
	
	'*** antal felter ***
	antalxx = x
	antalnn = n
	
	%>
	
	
	<input type="hidden" name="nialt" id="nialt" value="<%=nialt%>">
	<input type="hidden" name="antal_x" id="antal_x" value="<%=antalxx%>">
	<input type="hidden" name="antal_n" id="antal_n" value="<%=antalnn%>">
	
	<%
	if antalxx = 0 AND func = "red" then 
	%>
	<table>
	<tr bgcolor="#d6dff5">
		<td colspan=6 style="padding:10px;"><br><br><font color=red><b>Denne faktura <u>kladde</u> er oprettet uden der er valgt nogen aktiviteter.</b> <br>
		Der er derfor ikke nogen aktiviteter eller medarbejder timer der kan redigeres.<br><br>
		Slet denne kladde fra faktura-oversigten og opret en ny faktura med dette faktura nummer.</font><br><br>&nbsp;</td>
	<tr>
	<%
	end if
	%>
	</table>
	
	
	
	
	
    
     <div id=aktdiv_2 style="position:relative; visibility:hidden; display:none; top:80px; width:700px; left:5px; border:0px #8cAAe6 solid;">
        <table width=700 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('jobbesk')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('matdiv')" class=vmenu>Næste >></a></td></tr></table>
    </div>
    
	
	
	
	
	</div><!-- faktura linjer div -->
	
	
	<div id="aktsubtotal" style="position:absolute; left:730px; top:104px; width:200px; z-index:2000; border:1px yellowgreen solid; background-color:#ffffff; padding:5px;">
    <b>Fakturalinier:</b> (aktiviteter)<br />
    <table cellspacing=5 cellpadding=0 border=0 width=100%><tr>
	<td><!--Antal ialt:-->
	<%
	'*********************************************************
	'Subtotal
	'Total timer og beløb for aktiviteter
	'*********************************************************
	if intTimer <> 0 then
	intTimer = intTimer 
	else
	intTimer = 0
	end if
	
	%><input type="hidden" name="FM_timer" id="FM_Timer" value="0"><!--=intTimer-->
	<div style="position:relative; width:45; height:20px; border-bottom:0px YellowGreen dashed; background-color:#ffffff; padding-right:3px;" align="right" id="divtimertot">
	<!--<b><=intTimer%></b>--></div>
	</td>
	<td align="right">Subtotal:
		<!-- strBeloeb -->
		<%
		if len(totalbelob) <> 0 then
		thistotbelTimer = SQLBlessDot(formatnumber(totalbelob, 2))
		else
		thistotbelTimer = formatnumber(0, 2)
		end if
		
		if len(totalbelob_umoms) <> 0 then
		thistotbelTimer_umoms = SQLBlessDot(formatnumber(totalbelob_umoms, 2))
		else
		thistotbelTimer_umoms = formatnumber(0, 2)
		end if
		
		
		
		
		thistotbel = thistotbelTimer%>
		<input type="hidden" name="FM_timer_beloeb" id="FM_timer_beloeb" value="<%=thistotbelTimer%>">
		<input type="hidden" name="FM_timer_beloeb_umoms" id="FM_timer_beloeb_umoms" value="<%=thistotbelTimer_umoms%>">
		
		<div style="position:relative; width:95px; height:20px; border-bottom:2px YellowGreen dashed; background-color:#ffffff; padding-right:3px;" align="right" id="divtimerbelobtot"><b><%=thistotbel &" "& valutakodeSEL %></b></div>
		
		
		</td>
		</tr>
	</table>
	</div>
	
	

	
	
	
	
    
	
	
	
	
	
	
	
	
	