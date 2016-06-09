
 <!-- Aktiviteter -->
	    


 <div id="aktdiv" style="position:absolute; visibility:hidden; display:none; left:5px; top:105px; width:700px; border:1px yellowgreen solid; padding:10px 10px 10px 10px; background-color:#ffffff;">
    <table width=100% border=0 cellspacing=0 cellpadding=0>
	<tr>
	<td><h5>Aktiviteter</h5>
	
	<table>
	<tr>
		<td colspan=2 style="padding:2px 5px 2px 0px;"> 
		
                                <%if jobid <> 0 then 
                               
		                        select case intRabat
		                        case 0
		                        rSel0 = "SELECTED"
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 10
		                        rSel0 = ""
		                        rSel10 = "SELECTED"
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 25
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = "SELECTED"
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
	                            case 40
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = "SELECTED"
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 50
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = "SELECTED"
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 60
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = "SELECTED"
		                        rSel75 = ""
		                        case 75
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = "SELECTED"
		                        end select
		                        %>
                                    <select id="FM_rabat_all" name="FM_rabat_all" style="width:50px; font-family:verdana; font-size:9px;">
                                        <option value="0"  <%=rSel0%>>0%</option>
                                        <option value="0.10" <%=rSel10%>>10%</option>
                                        <option value="0.25" <%=rSel25%>>25%</option>
                                        <option value="0.40" <%=rSel40%>>40%</option>
                                        <option value="0.50" <%=rSel50%>>50%</option>
                                        <option value="0.60" <%=rSel60%>>60%</option>
                                        <option value="0.75" <%=rSel75%>>75%</option>
                                        </select>
                                    &nbsp;<input id="setRabtAll" type="button" value="Opdater (2-5 sek.)" style="font-size:9px;" onClick="opd_rabatrall(), setBeloebTotAll()" />
								
								<%end if %>
								
	
	
   
    
    <%select case intEnhedsang
	case 1
	chke3 = ""
	chke2 = "CHECKED"
	chke1 = ""
	case 2
	chke3 = "CHECKED"
	chke2 = ""
	chke1 = ""
	case else
	chke3 = ""
	chke2 = ""
	chke1 = "CHECKED"
	end select%>
          
		<br /><br /><b>Enhed på aktiviteter angives som:</b>
		<input type="radio" name="FM_enheds_ang" value="0" <%=chke1%> onclick="opd_akt_endhed('time','0')"> Timer&nbsp;&nbsp; 
		<input type="radio" name="FM_enheds_ang" value="1" <%=chke2%> onclick="opd_akt_endhed('stk.','1')"> Stk.&nbsp;&nbsp;
		<input type="radio" name="FM_enheds_ang" value="2" <%=chke3%> onclick="opd_akt_endhed('enhed','2')"> Enheder<br />
		
		<!--
		<br /><b>Vis antal som Timer ell. Faktor (Enheder)</b>:<br />
		Timer <input id="FM_timer_faktor0" name="FM_timer_faktor" type="radio" value=0 /> 
        Faktor / Enheder<input id="FM_timer_faktor1" name="FM_timer_faktor" type="radio" value=1 />
        <input id="Submit1" type="submit" value="Opdater" style="font-size:9px;" onClick="opd_timer_faktor_all(), setBeloebTotAll()" />
		-->
	
	    
       
	
	</td>
	</tr>
	</table>
<%


	
	'***********************************************************************************************
	'*** Henter Medarbejdere og Aktiviteter ************************************************************************
	'***********************************************************************************************
	        
	        
	        isaktwritten = " aktiviteter.id <> 0 "
	        x = 0
	        
	        
	        if func = "red" then
        		
		        isTPused = "#"
        		
		        '*******************************************************************************
		        '******************* Henter oprettede aktiviteter ******************************
		        '*******************************************************************************
		        strSQL = "SELECT faktura_det.id, aktid, aktiviteter.navn, aktiviteter.fakturerbar, aktiviteter.faktor FROM faktura_det LEFT JOIN aktiviteter ON (aktiviteter.id = aktid) WHERE fakid = "& intFakid &" GROUP BY aktid" 
		        oRec.open strSQL, oConn, 3
		       
		        While not oRec.EOF 
        		
			        'Redim preserve thisaktnavn(x)
			        thisaktnavn(x) = oRec("navn") 'oRec("beskrivelse")
			        
			        'Redim preserve thisaktid(x)
			        thisaktid(x) = oRec("aktid")
			        
			        'Redim preserve thisaktfunc(x)
			        thisaktfunc(x) = "red"
			        
			        thisaktfakbar(x) = oRec("fakturerbar")
			        thisaktfaktor(x) = oRec("faktor")
        			
        		
        		isaktwritten = isaktwritten & " AND aktiviteter.id <> "& thisaktid(x) &"" 
        			
		        x = x + 1
        		
		        oRec.movenext
		        wend 
		        oRec.close
        	
	        end if
	        
	        
	        
			
		    '*******************************************************************************
		    '*** Henter alle de aktiviteter der ikke allerede er oprettet (ved rediger) ****
		    '*** Henter alle aktiviteter ved opr. ******************************************
		    '*******************************************************************************
			
			'**** Sidste faktura tidspunkt ****
			if len(strfaktidspkt) > 0 then
			LastFakTid = datepart("h", strfaktidspkt)&":"& datepart("n", strfaktidspkt)&":"& datepart("s", strfaktidspkt)
			else
			LastFakTid =  "00:00:02"
			end if
			
			
			'*** Henter aktiviter og finder registrede timer siden sidste faktura dato. (periodeafgrænsning) *** 
			strSQL = "SELECT job.fastpris, job.jobtpris, job.budgettimer, "_
			&" aktstatus, aktiviteter.job, aktiviteter.id AS aid, aktiviteter.navn AS anavn, aktiviteter.fakturerbar, aktiviteter.faktor "_
			&" FROM aktiviteter LEFT JOIN job ON (job.id = aktiviteter.job) "_
			&" WHERE (aktiviteter.job = "& jobid &" AND aktiviteter.fakturerbar <> 2) AND ("& isaktwritten &") "_
			&" ORDER BY aktiviteter.id"	'AND aktiviteter.aktstatus = 1, 
			oRec.open strSQL, oConn, 3
			
			
			lastAktId = 0
			timerThisT = 0
			timerThisD = 0
			'strKom = "Netto Kontant 8 dage." Hentes på kunde
			totaltimer = 0
			lastaktid = 0
			
			
			while not oRec.EOF
				    
				    
				   if cint(lastaktid) <> cint(oRec("aid")) then
						
						'Redim preserve thisaktid(x)
						'Redim preserve usefastpris(x)
						thisaktid(x) = oRec("aid")
						usefastpris(x) = oRec("fastpris")
						
						'** timepris på fastpris job ***
							if cint(oRec("fastpris")) = 1 then
								akttimepris = oRec("jobtpris")/oRec("budgettimer")
							else
								akttimepris = 0
							end if
							
						
						'** antal registerede timer pr. akt.
						strSQL2 = "SELECT sum(timer.timer) AS timerthis FROM timer WHERE taktivitetid =" & oRec("aid") &" AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"
						' AND tfaktim = 1
						' OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"')"		
						oRec2.open strSQL2, oConn, 3
						
						if not oRec2.EOF then
							timerThisD = oRec2("timerthis")
						end if
						oRec2.close
						
							if len(timerThisD) <> 0 then
							timerThisD = timerThisD 
							else
							timerThisD = 0
							end if
							
						
						
					'Redim preserve thisaktnavn(x)
					thisaktnavn(x) = oRec("anavn") 
					
					'Redim preserve thisAktTimer(x)
					thisAktTimer(x) = timerThisD
					
					'Redim preserve thisAktBeloeb(x)
					thisAktBeloeb(x) = timerThisD * akttimepris
					
					'Redim preserve thisaktfunc(x)
			        thisaktfunc(x) = "opr"
					
					thisaktfakbar(x) = oRec("fakturerbar")
					
					thisaktfaktor(x) = oRec("faktor") 
					
					timerThisD = 0
					thisBeloebD = 0
					
					
				x = x + 1		
				end if
			
			lastaktid = oRec("aid")		
			oRec.movenext
			wend
			oRec.close
			
			
	'end if
	'***********************************************************************************
	
	
	
	
	'***********************************************************************************
	'********************* Udskriver aktiviterer ***************************************
	'***********************************************************************************
	
	intTimer = 0
	totalbelob = 0
	subtotalbelob = 0
	subtotaltimer = 0
	
	nialt = 0
	for x = 0 to x - 1 
	
	%>
	<br /><br /><br /><b><%=thisaktnavn(x)%></b> <br />
	Faktor: <input type="text" name="aktfaktor_<%=x%>" id="aktfaktor_<%=x%>" value="<%=thisaktfaktor(x) %>" style="font-size:9px; width:30px;">
	<input id="beregnfaktor_<%=x%>" type="button" value="Beregn" style="font-size:9px;" onClick="opd_timer_faktor(<%=x%>)" /><br /> ("Andet?" sum-aktiviteten skal beregnes manuelt) 
	
	<%if cint(thisaktfakbar(x)) <> 1 then %>
	&nbsp;&nbsp;<font class=roed><b>(Ej fakturerbar)</b></font>
	<%end if %>
	
        <!--&nbsp;&nbsp;&nbsp;<a href="#" onclick="showlommeregner()">Lommerenger</a>-->
	
	<div id=submitTop style="position:relative; left:550px; top:-15px; z-index:10000;">
	<input name="subm_on_top_<%=x%>" id="subm_on_top_<%=x%>" type="submit" value="Se faktura" style="font-size:9px;" />
	</div>
	
	
	<%if thisaktfunc(x) <> "red" AND func = "red" then %>
	&nbsp;<font class=roed>Denne aktivitet er ikke oprettet på faktura - timer er hentet fra timeregistrering i den valgte periode.</font>
	<%end if%>
	<table cellspacing="0" cellpadding="2" border="0" width=650 bgcolor="silver">
	<%
	
	%>
	<input type="hidden" name="sumaktbesk_<%=x%>" id="sumaktbesk_<%=x%>" value="<%=thisaktnavn(x)%>">
	<%
		'if instr(aktiswritten, "#"&thisaktid(x)&"#") = 0 then
		n = 1
		a = 0
		newa = 0
		sa = 0
	
					if thisaktfunc(x) <> "red" then
					
					
							call projektgrupper()	
							
							
					end if
					
					if thisaktid(x) <> 0 then%>
					<!--<tr bgcolor="#d6dff5">
						<td colspan=6><!--Deaktiverede medarbejdere vises ikke ved faktura oprettelse. pr. 29/8-2005.<img src="../ill/blank.gif" width="1" height="20" alt="" border="0"><br></td>
					</tr>-->
					
					<%
					if func = "red" then
					
					    if thisaktfunc(x) = "red" then
					    bgthis = "YellowGreen"
					    else
					    bgthis = "silver"
					    end if
					
					else
					    
					    bgthis = "YellowGreen"
					    
					end if %>
					
					<tr bgcolor="<%=bgthis %>">
						<td>&nbsp;</td>
						<td class=lille width=90 style="padding:2px 2px 2px 5px;">Vis - Antal</td>
                        <td class=lille width=120 style="padding:2px 2px 2px 5px;">Ventetimer <br />
                        <font class=megetlillesort>Saldo : Brugt : Ultimo</font></td>
						<td class=lille style="padding:2px 2px 2px 5px;">Medarbejder(e)</td>
						<td class=lille style="padding:2px 2px 2px 5px;">Enh.pris</td>
						<td class=lille style="padding:2px 2px 2px 5px;">Enhed</td>
						<td class=lille style="padding:2px 2px 2px 5px;" align=center>Rabat</td>
						<td class=lille style="padding:2px 2px 2px 5px;" align=right>Pris ialt&nbsp;&nbsp;&nbsp;</td>
						<td style="padding:2px 2px 2px 5px;"><font class=megetlillesort>Afrund</td>
					</tr>
					<!--<tr bgcolor="#66CC33"><td colspan=6 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="2" alt="" border="0"><br></td></tr>-->
			
					<%end if
					
					
					
					
					'***************************************************
					'Faktimer, Venter og Beløb på medarbejdere 
					'***************************************************
					
					faktot = 0
					medarbstrpris = ""
					
					
					
					'**** Henter alle allerede oprettede aktiviteter ***
					if thisaktfunc(x) = "red" then
						
						strSQL10 = "SELECT fms.medrabat, fd.id, antal, "_
						&" fd.beskrivelse, aktpris, fd.showonfak AS akt_showonfak, "_
						&" fd.enhedspris, "_
						&" fd.aktid, "_
						&" fak, venter_brugt, tekst, fms.enhedspris, beloeb, "_
						&" mid, fms.showonfak AS showonfak, venter AS venter_ultimo, fms.enhedsang "_
						&" FROM faktura_det fd "_
						& "LEFT JOIN fak_med_spec fms ON (fms.fakid = fd.fakid AND fms.aktid = fd.aktid)"_
						&" WHERE fd.fakid = "& intFakid &" AND fd.aktid = "& thisaktid(x) &" GROUP BY mid"
						
					else
					    
					'**** Ved red: Henter alle de aktiviteter der endnu ikke er oprettet ***
                    '**** Ved opr: Hennter alle aktiviteter ********************************
                    
                    usedmedabId = " Tmnr = 0 AND "
						
				    strSQL10 = "SELECT DISTINCT(medarbejderid), mnavn, mid, timeprisalt, medarbejdertype FROM progrupperelationer, medarbejdere "_
				    &" LEFT JOIN timepriser ON (aktid = "& thisaktid(x) &" AND medarbid = medarbejderid) WHERE mid = medarbejderid AND mansat <> 2 AND "_
				    &" (projektgruppeid = "& gp1 &" OR projektgruppeid = "& gp2 &" OR projektgruppeid = "& gp3 &" OR projektgruppeid = "& gp4 &" "_
				    &" OR projektgruppeid = "& gp5 &" OR projektgruppeid = "& gp6 &" OR projektgruppeid = "& gp7 &" OR projektgruppeid = "& gp8 &" "_
				    &" OR projektgruppeid = "& gp9 &" OR projektgruppeid = "& gp10 &") "_
				    &" GROUP BY medarbejderid ORDER BY projektgruppeid" 
    				
					
					
					end if
					
						
					oRec3.open strSQL10, oConn, 3
					while not oRec3.EOF
					
					
					call fakturaMedarbLinjer()
					nialt = nialt + 1
							
					Response.flush
					oRec3.movenext
					wend
					oRec3.close
					
					%>
					
					</table>
					<br /><br />
					<table cellspacing="0" cellpadding="0" border="0" width="600">
                    <%
                    faktot = 0
					mbelobtot = 0
			
			
			
			call subtotalakt(x, 0)
			
			%>
			
			<!--<tr>
			<td colspan=6 bgcolor="#ffffff" style="border-left:1px #8caae6 solid; border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="10" alt="" border="0"><br>
			</td>
			</tr>
			<tr><td colspan=6 bgcolor="#ffffff"><img src="../ill/blank.gif" width="1" height="8" alt="" border="0"><br></td></tr>
			-->
			
			
			<%
			Response.write strAktsubtotal
			strAktsubtotal = ""
			%>
			<tr><input type="hidden" name="antal_subtotal_akt_<%=x%>" id="antal_subtotal_akt_<%=x%>" value="<%=sa%>">
			<td bgcolor=Pink style="padding:10px 10px 10px 10px; border:0px red dashed;">
                <input id="kopier_sum_<%=x%>" type="button" value="^ Kopier til sum-aktiviet. ^" style="font-size:9px;" onclick="overtilsum(<%=x %>)" />&nbsp;&nbsp;&nbsp;
			<b><%=thisaktnavn(x)%></b> antal ialt:
			<input type="text" name="timer_subtotal_akt_<%=x%>" id="timer_subtotal_akt_<%=x%>" value="<%=subtotaltimer%>" onkeyup="offsetSumTtimer(<%=x %>)" style='border:1px yellowgreen solid; font-size:9px; width:55; background-color:#ffffff;'>
             
              
             &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Pris ialt:
			<input type="text" name="belob_subtotal_akt_<%=x%>" id="belob_subtotal_akt_<%=x%>" value="<%=subtotalbelob%>" onkeyup="offsetSumTbeloeb(<%=x %>)" style='border:1px #5582d2 solid; font-size:9px; width:85; background-color:#ffffff;'> kr.
			</td>
			</tr>
			
			<%if func <> "red" then%>
			<tr>
				<td bgcolor=Pink style="padding:5px 5px 5px 5px;">
				<%
				'** Gemte timer på medarbejdere der ikke længere har adgang til denne akt ***
				
				len_usedmedabId = len(usedmedabId)
				temp_usedmedabId = left(usedmedabId, len_usedmedabId - 4)
				usedmedabIdKri = temp_usedmedabId
				
				strSQL2 = "SELECT sum(timer) AS timer FROM timer WHERE "& usedmedabIdKri &" AND taktivitetid ="& thisaktid(x) &" AND tfaktim = 1 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')" 'OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"'
				'Response.write strSQL2
				oRec2.open strSQL2, oConn, 3 
				
				if not oRec2.EOF then 
					medarbejderGTimer = oRec2("timer")
				end if
				oRec2.close
				
				if len(medarbejderGTimer) <> 0 then
					medarbejderGTimer = medarbejderGTimer
				else
					medarbejderGTimer = 0
				end if
				
				if medarbejderGTimer <> 0 then
				Response.write "Skjulte timer på aktivitet:<b> " & formatnumber(medarbejderGTimer, 2) & "</b> timer"
				else
				Response.write "&nbsp;"
				end if
				%>
				</td>
			</tr>
			<%end if%>
			
			<input type="hidden" name="aktId_n_<%=x%>" id="aktId_n_<%=x%>" value="<%=thisaktid(x)%>">
			<input type="hidden" name="antal_n_<%=x%>" id="antal_n_<%=x%>" value="<%=n%>">
			<input type="hidden" name="highest_aval_<%=x%>" id="highest_aval_<%=x%>" value="<%=sa%>">
			<% 
		subtotaltimer = 0 
		subtotalbelob = 0 
		'aktiswritten = aktiswritten & ",#"&thisaktid(x)&"#"		
		'end if	
		%></table><%
	next
	%>
	
	
	
	
	<input type="hidden" name="lastactive_x" id="lastactive_x" value="0">
	<input type="hidden" name="lastactive_a" id="lastactive_a" value="0">
	<%
	
	
	'*** antal felter ***
	antalxx = x
	antalnn = n
	
	%>
	
	
	<input type="hidden" name="nialt" value="<%=nialt%>">
	<input type="hidden" name="antal_x" value="<%=antalxx%>">
	<input type="hidden" name="antal_n" id="antal_n" value="<%=antalnn%>">
	
	<%
	if antalxx = 0 AND func = "red" then 
	%>
	<table>
	<tr bgcolor="#d6dff5">
		<td colspan=6 style="padding:10px;"><br><br><font color=red><b>Denne faktura <u>kladde</u> er oprettet uden der er valgt nogen aktiviteter.</b> <br>
		Derer derfor ikke nogen aktiviteter eller medarbejder timer der kan redigeres.<br><br>
		Slet denne kladde fra faktura-oversigten og opret en ny faktura med dette faktura nummer.</font><br><br>&nbsp;</td>
	<tr>
	<%
	end if
	%>
	</table>
	
	
	<!--</div>
	
	
	
	
	<div id=aktsubtotal style="position:relative; width:700; visibility:visible; border:1px yellowgreen solid; left:5px; top:600px; padding:10px 10px 10px 10px; background-color:#ffffff;">
    -->
    <br /><br />
    <b>Sub-total aktiviteter:</b><br />
	<table cellspacing=5 cellpadding=0 border=0 width=700 bgcolor="#ffffff"><tr>
	<td style="padding:10px 10px 10px 20px;">Antal ialt:
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
	
	%><input type="hidden" name="FM_timer" value="<%=intTimer%>">
	<div style="position:relative; width:65; height:20px; border-bottom:2px YellowGreen dashed; background-color:#ffffff; padding-right:3px;" align="right" id="divtimertot"><b><%=intTimer%></b></div>
	</td>
	<td style="padding:10px 20px 10px 10px;" align="right">Subtotal:
		<!-- strBeloeb -->
		<%
		if len(totalbelob) <> 0 then
		thistotbelTimer = SQLBlessDot(formatnumber(totalbelob, 2))
		else
		thistotbelTimer = formatnumber(0, 2)
		end if
		
		
		
		thistotbel = thistotbelTimer%>
		<input type="hidden" name="FM_timer_beloeb" value="<%=thistotbelTimer%>">
		<div style="position:relative; width:65; height:20px; border-bottom:2px #86B5E4 dashed; background-color:#ffffff; padding-right:3px;" align="right" id="divtimerbelobtot"><b><%=thistotbel%></b></div>
		</td>
		</tr>
	</table>
	
	
	
    
     <div id=aktdiv_2 style="position:relative; visibility:hidden; display:none; top:80px; width:700px; left:5px; border:0px #8cAAe6 solid;">
        <table width=700 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('jobbesk')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('matdiv')" class=vmenu>Næste >></a></td></tr></table>
    </div>
    
	
	</div>
	
	
	
	

	
	
	
	
    
	
	                      <!-- Materialer --->
	                      
	                    <div id=matdiv style="position:absolute; width:700px; visibility:hidden; display:none; border:1px orange solid; top:96px; left:5px; padding:10px 10px 10px 10px; background-color:#ffffff;">
                       
	                        <h5>Materialer</h5>
	                        
	                        
	                        <table width=700 cellspacing=0 cellpadding=0 border=0>
	                        </tr>
	                        <tr bgcolor="orange">
	                        <td height=20>
                                &nbsp;</td>
	                            <td class=lille>Vis</td>
	                            <td class=lille>Navn</td>
	                            <td class=lille>Varenr</td>
	                            <td class=lille>Antal</td>
	                            <td class=lille>Enheds pris</td>
	                            <td class=lille>Enhed</td>
	                            <td class=lille>Rabat</td>
	                            <td class=lille>Pris ialt</td>
	                        </tr>
	                        <!-- Materialer -->
	                        <%
	                        '** Allerede fakturarede materialer 
	                        m = 0
	                        isMatWrt = " matid <> 0 "
                        	
	                        if func = "red" then
                        	
	                        strSQL = "SELECT matid, matnavn, matantal, matenhedspris AS matsalgspris, matenhed, matvarenr FROM fak_mat_spec WHERE matfakid = "& id &" ORDER BY matnavn"
	                        'Response.Write strSQL
	                        oRec.open strSQL, oConn, 3
                            while not oRec.EOF
                            
                            call materialer(1)
                            
                            isMatWrt = isMatWrt & " AND matid <> " & oRec("matid")
                            m = m + 1
                            oRec.movenext
                            wend
                            oRec.close
                        	
                        	
	                        
                        	
	                        end if
                        	
                        	
	                        '** Ikke fakturerede Materialer / Eller v. ny faktura = alle 
	                        strSQL = "SELECT matid, matnavn, sum(matantal) AS matantal, matsalgspris, matenhed, matvarenr FROM materiale_forbrug WHERE jobid = "& jobid &" AND forbrugsdato BETWEEN '"& stdatoKri &"' AND '"& slutdato &"' AND ("& isMatWrt &") GROUP BY  matid, matsalgspris ORDER BY  matnavn"
	                        
	                        
	                        oRec.open strSQL, oConn, 3
                            im = 0
                            while not oRec.EOF
                                
                                if im = 0 AND func = "red" then
                                %>
	                             <tr><td colspan=9>
                                     &nbsp;</td></tr>
    	                       
	                            <tr><td colspan=9><font class=roed>Ikke fakturerede materialer, registreret i den valgte periode.</font></td></tr>
	                             <tr bgcolor="silver">
	                            <td>
                                &nbsp;</td>
	                            <td class=lille>Vis</td>
	                            <td class=lille>Navn</td>
	                            <td class=lille>Varenr</td>
	                            <td class=lille>Antal</td>
	                            <td class=lille>Enheds pris</td>
	                            <td class=lille>Enhed</td>
	                            <td class=lille>Rabat</td>
	                            <td class=lille>Pris ialt</td>
	                        </tr>
	                            <%
                                end if
                                
                            call materialer(2)
                            
                           
                            
                            im = im + 1
                            m = m + 1
                            oRec.movenext
                            wend
                            oRec.close
	                        
	                        if m = 0 then%>
	                        <tr><td colspan=9><br />(Der er ikke fundet nogen materialer i det valgte interval.)</td></tr>
	                        <%end if  %>
                        	
                             <input id="FM_antal_materialer_ialt" name="FM_antal_materialer_ialt" value="<%=m%>" type="hidden" />
	                        </table>
                      
                        	
	
	
	<br /><br /><b>Sub-total materialer:</b><br />
	<table cellspacing=5 cellpadding=0 border=0 width=700 bgcolor="#ffffff"><tr>
	<td align=right style="padding:10px 20px 10px 20px;">Materialer ialt:
	<%
	'*********************************************************
	'Subtotal
	'*********************************************************
	%>
	
	
		<!-- matBeloeb -->
		<%
		if len(matSubTotalAll) <> 0 then
		matSubTotalAll = SQLBlessDot(formatnumber(matSubTotalAll, 2))
		else
		 matSubTotalAll = formatnumber(0, 2)
		end if
		
		
		
		%>
		<input type="hidden" name="FM_materialer_beloeb" value="<%= matSubTotalAll%>">
		<div style="position:relative; width:65; height:20px; border-bottom:2px orange dashed; background-color:#ffffff; padding-right:3px;" align="right" id="divmatbelobtot"><b><%= matSubTotalAll%></b></div>
		</td>
		</tr>
	</table>
	
	
	 <div id=matdiv_2 style="position:relative; visibility:hidden; display:none; top:80px; width:700px; left:5px; border:0px #8cAAe6 solid;">
        <table width=700 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('aktdiv')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('betdiv')" class=vmenu>Næste >></a></td></tr></table>
    </div>
	
	
    </div>
	 
	
	
	
	
	
	
	
	
	