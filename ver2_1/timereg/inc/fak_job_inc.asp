
<br><br>
<h4>Aktiviteter:</h4>
<table cellspacing="0" cellpadding="0" border="0" width="700" bgcolor="#FFFFFF">
	<%
	'***********************************************************************************************
	'*** Henter aktiviteter ************************************************************************
	'***********************************************************************************************
	if func = "red" then
		
		isTPused = "#"
		
		'******************* Henter oprettede aktiviteter *******************************
		strSQL = "SELECT faktura_det.id, aktid, aktiviteter.navn FROM faktura_det LEFT JOIN aktiviteter ON (aktiviteter.id = aktid) WHERE fakid = "& intFakid &" GROUP BY aktid" 
		oRec.open strSQL, oConn, 3
		x = 0
		While not oRec.EOF 
		
			Redim preserve thisaktnavn(x)
			thisaktnavn(x) = oRec("navn") 'oRec("beskrivelse")
			Redim preserve thisaktid(x)
			thisaktid(x) = oRec("aktid")
			
			
		x = x + 1
		
		oRec.movenext
		wend 
		oRec.close
	
	else
			'**** Sidste faktura tidspunkt ****
			if len(strfaktidspkt) > 0 then
			LastFakTid = datepart("h", strfaktidspkt)&":"& datepart("n", strfaktidspkt)&":"& datepart("s", strfaktidspkt)
			else
			LastFakTid =  "00:00:02"
			end if
			
			
			'*** Henter aktiviter og finder registrede timer siden sidste faktura dato. (periodeafgrænsning) ***' 
			strSQL = "SELECT job.fastpris, job.jobtpris, job.budgettimer, aktstatus, aktiviteter.job, aktiviteter.id AS aid, aktiviteter.navn AS anavn FROM aktiviteter LEFT JOIN job ON (job.id = aktiviteter.job) WHERE aktiviteter.job = "& jobid &" AND aktiviteter.fakturerbar = 1 "_
			&" ORDER BY aktiviteter.fase, aktiviteter.sortorder, aktiviteter.id"	'AND aktiviteter.aktstatus = 1	
			oRec.open strSQL, oConn, 3
			
			x = 0
			lastAktId = 0
			timerThisT = 0
			timerThisD = 0
			'strKom = "Netto Kontant 8 dage." Hentes på kunde
			totaltimer = 0
			lastaktid = 0
			
			
			while not oRec.EOF
				if cint(lastaktid) <> cint(oRec("aid")) then
						
						Redim preserve thisaktid(x)
						Redim preserve usefastpris(x)
						thisaktid(x) = oRec("aid")
						usefastpris(x) = oRec("fastpris")
						
						'** timepris på fastpris job ***
							if cint(oRec("fastpris")) = 1 then
								akttimepris = oRec("jobtpris")/oRec("budgettimer")
							else
								akttimepris = 0
							end if
							
						
						'** antal registerede timer pr. akt.
						strSQL2 = "SELECT sum(timer.timer) AS timerthis FROM timer WHERE taktivitetid =" & oRec("aid") &" AND tfaktim = 1 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"   ' OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"')"		
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
							
						
						
					Redim preserve thisaktnavn(x)
					thisaktnavn(x) = oRec("anavn") 
					Redim preserve thisAktTimer(x)
					thisAktTimer(x) = timerThisD
					'Redim preserve thisTimePris(x)
					'thisTimePris(x) = akttimepris
					Redim preserve thisAktBeloeb(x)
					thisAktBeloeb(x) = timerThisD * akttimepris
					
					
					
					timerThisD = 0
					thisBeloebD = 0
					
					
				x = x + 1		
				end if
			
			lastaktid = oRec("aid")		
			oRec.movenext
			wend
			oRec.close
			
			
	end if
	'***********************************************************************************
	
	
	
	
	
	'********************* Udskriver aktiviterer *********************
	'aktiswritten = "#0#,"
	for x = 0 to x - 1 
	
	%>
	<input type="hidden" name="sumaktbesk_<%=x%>" id="sumaktbesk_<%=x%>" value="<%=thisaktnavn(x)%>">
	<%
		'if instr(aktiswritten, "#"&thisaktid(x)&"#") = 0 then
		n = 1
		a = 0
		newa = 0
		sa = 0
	
					if func <> "red" then
							'**** projektgrupper på job ******************************************
							'Tjekker at der den projgp der findes på akt også findes på jobbet. 
							'Gælder kun job oprettet før 5/11-2004, da der her er lavet en 
							'funktion der tjekker det samme når et job redigeres. 
							'*********************************************************************
							
								jobPgstring = ""
							
								strSQL = "SELECT id, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM job WHERE id = "& jobid
								oRec2.open strSQL, oConn, 3
									if not oRec2.EOF then
									
									strProjektgr1 = oRec2("projektgruppe1")
									strProjektgr2 = oRec2("projektgruppe2")
									strProjektgr3 = oRec2("projektgruppe3")
									strProjektgr4 = oRec2("projektgruppe4")
									strProjektgr5 = oRec2("projektgruppe5")
									strProjektgr6 = oRec2("projektgruppe6")
									strProjektgr7 = oRec2("projektgruppe7")
									strProjektgr8 = oRec2("projektgruppe8")
									strProjektgr9 = oRec2("projektgruppe9")
									strProjektgr10 = oRec2("projektgruppe10")
									
									jobPgstring = "#"&strProjektgr1&"#,#"&strProjektgr2&"#,#"&strProjektgr3&"#,#"&strProjektgr4&"#,#"&strProjektgr5&"#,#"&strProjektgr6&"#,#"&strProjektgr7&"#,#"&strProjektgr8&"#,#"&strProjektgr9&"#,#"&strProjektgr10&"#"
									
									end if
								oRec2.close
							
							
							
							'*** Medarbejder udspecificering ***
							strSQL3 = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter WHERE id = "& thisaktid(x)
							oRec3.open strSQL3, oConn, 3
							if not oRec3.EOF then
							
								gp1 = oRec3("projektgruppe1")
								gp2 = oRec3("projektgruppe2")
								gp3 = oRec3("projektgruppe3")
								gp4 = oRec3("projektgruppe4")
								gp5 = oRec3("projektgruppe5")
								gp6 = oRec3("projektgruppe6")
								gp7 = oRec3("projektgruppe7")
								gp8 = oRec3("projektgruppe8")
								gp9 = oRec3("projektgruppe9")
								gp10 = oRec3("projektgruppe10")
							
							end if
							oRec3.Close 
							
							'*** Findes gruppen også på jobbet?? ****
							if len(gp1) <> 0 AND instr(jobPgstring, "#"&gp1&"#") > 0 then
							gp1 = gp1
							else
							gp1 = 1
							end if
							
							if len(gp2) <> 0 AND instr(jobPgstring, "#"&gp2&"#") > 0 then
							gp2 = gp2
							else
							gp2 = 1
							end if
							
							if len(gp3) <> 0 AND instr(jobPgstring, "#"&gp3&"#") > 0 then
							gp3 = gp3
							else
							gp3 = 1
							end if
							
							if len(gp4) <> 0 AND instr(jobPgstring, "#"&gp4&"#") > 0 then
							gp4 = gp4
							else
							gp4 = 1
							end if
							
							if len(gp5) <> 0 AND instr(jobPgstring, "#"&gp5&"#") > 0 then
							gp5 = gp5
							else
							gp5 = 1
							end if
							
							if len(gp6) <> 0 AND instr(jobPgstring, "#"&gp6&"#") > 0 then
							gp6 = gp6
							else
							gp6 = 1
							end if
							
							if len(gp7) <> 0 AND instr(jobPgstring, "#"&gp7&"#") > 0 then
							gp7 = gp7
							else
							gp7 = 1
							end if
							
							if len(gp8) <> 0 AND instr(jobPgstring, "#"&gp8&"#") > 0 then
							gp8 = gp8
							else
							gp8 = 1
							end if
							
							if len(gp9) <> 0 AND instr(jobPgstring, "#"&gp9&"#") > 0 then
							gp9 = gp9
							else
							gp9 = 1
							end if
							
							if len(gp10) <> 0 AND instr(jobPgstring, "#"&gp10&"#") > 0 then
							gp10 = gp10
							else
							gp10 = 1
							end if	
							
							
					end if
					
					if thisaktid(x) <> 0 then%>
					<!--<tr bgcolor="#d6dff5">
						<td colspan=6><!--Deaktiverede medarbejdere vises ikke ved faktura oprettelse. pr. 29/8-2005.<img src="../ill/blank.gif" width="1" height="20" alt="" border="0"><br></td>
					</tr>-->
					<tr bgcolor="#66CC33">
						<td style="border-left:1px #8caae6 solid; border-top:1px #8caae6 solid;">&nbsp;</td>
						<td width=205 style="border-top:1px #8caae6 solid;"><br><font class=megetlillesort>Vis - Timer - (Ventetim.) - Nye Ventetim.</td>
						<td width=335 style="border-top:1px #8caae6 solid;"><br>&nbsp;&nbsp;&nbsp;<b><%=thisaktnavn(x)%></b><font class=megetlillesort>&nbsp;&nbsp;&nbsp;Medarbejder(e)&nbsp;&nbsp;</td>
						<td width=105 style="border-top:1px #8caae6 solid;" align=center><br><font class=megetlillesort>Timepris</td>
						<td style="border-top:1px #8caae6 solid;" align=right><br><font class=megetlillesort>Beløb&nbsp;&nbsp;&nbsp;</td>
						<td style="border-top:1px #8caae6 solid; border-right:1px #8caae6 solid;"><br><font class=megetlillesort>&nbsp;&nbsp;Afrund</td>
					</tr>
					<!--<tr bgcolor="#66CC33"><td colspan=6 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="2" alt="" border="0"><br></td></tr>-->
			
					<%end if
					
					'***************************************************
					'Faktimer, Venter og Beløb på medarbejdere 
					'***************************************************
					
					faktot = 0
					medarbstrpris = ""
					
					
					if func = "red" then
						
						strSQL10 = "SELECT faktura_det.id, antal, faktura_det.beskrivelse, aktpris, faktura_det.showonfak AS akt_showonfak, "_
						&" faktura_det.enhedspris, "_
						&" faktura_det.aktid, "_
						&" fak, venter, tekst, fak_med_spec.enhedspris, beloeb, mid, fak_med_spec.showonfak AS showonfak "_
						&" FROM faktura_det "_
						& "LEFT JOIN fak_med_spec ON (fak_med_spec.fakid = faktura_det.fakid AND fak_med_spec.aktid = faktura_det.aktid)"_
						&" WHERE faktura_det.fakid = "& intFakid &" AND faktura_det.aktid = "& thisaktid(x) &" GROUP BY mid"
						
						
					else
						
						usedmedabId = " Tmnr = 0 AND "
						
						strSQL10 = "SELECT DISTINCT(medarbejderid), mnavn, mid, timeprisalt, medarbejdertype FROM progrupperelationer, medarbejdere "_
						&" LEFT JOIN timepriser ON (aktid = "& thisaktid(x) &" AND medarbid = medarbejderid) WHERE mid = medarbejderid AND mansat <> 2 AND "_
						&" (projektgruppeid = "& gp1 &" OR projektgruppeid = "& gp2 &" OR projektgruppeid = "& gp3 &" OR projektgruppeid = "& gp4 &" OR projektgruppeid = "& gp5 &" OR projektgruppeid = "& gp6 &" OR projektgruppeid = "& gp7 &" OR projektgruppeid = "& gp8 &" OR projektgruppeid = "& gp9 &" OR projektgruppeid = "& gp10 &") "_
						&" GROUP BY medarbejderid ORDER BY projektgruppeid"', timer DESC"
						
						'** 30/8-2005 fjernet ***
						'** PGA timer stod sommetider med dobbelt antal, beregnes nedenfor. ***
						'&" LEFT JOIN timer ON (Tmnr = medarbejderid AND taktivitetid ="& thisaktid(x) &" AND tfaktim = 1 AND Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"_  
						
						'Response.write strSQL10
						'Response.flush
						
					end if
					
					oRec3.open strSQL10, oConn, 3
					
					while not oRec3.EOF
					
								
								if func = "red" then
									
									fak = oRec3("fak")
									venter = oRec3("venter")
									'Response.write "fak "& fak  & "venter  " & venter & "<br>"
									medarbejderTimepris = oRec3("enhedspris")
									medarbejderBeloeb = oRec3("beloeb")
									txt = oRec3("tekst")
									medarbid = oRec3("mid")		
									showonfak = oRec3("showonfak")
									akt_showonfak = 0 'oRec3("akt_showonfak")
									
									
									if len(medarbejderTimepris) <> 0 then
									medarbejderTimepris = SQLBlessDOT(formatnumber(medarbejderTimepris, 2))
									else
									medarbejderTimepris = SQLBlessDot(formatnumber(0, 2))
									end if
								
								else
										
										'***************************************************************
										'Timer medarbejder 
										'***************************************************************
										
										strSQL2 = "SELECT sum(timer) AS timer FROM timer WHERE Tmnr = "&oRec3("medarbejderid")&" AND taktivitetid ="& thisaktid(x) &" AND tfaktim = 1 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')" 'OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"'
										oRec2.open strSQL2, oConn, 3 
										
										if not oRec2.EOF then 
											medarbejderTimer = oRec2("timer")
										end if
										oRec2.close
										
										
										'medarbejderTimer = oRec3("timer")
										if len(medarbejderTimer) <> 0 then
											medarbejderTimer = medarbejderTimer
										else
											medarbejderTimer = 0
										end if
										
										usedmedabId = usedmedabId & " Tmnr <> "& oRec3("medarbejderid") &" AND "
										
										'***************************************************************
										'Timepris pr. medarbejder ***
										'***************************************************************
										%>
										<!--#include file="fak_inc_mtimepris_opret.asp" -->
										<%
										
										'***************************************************************
										'Fak, Venter og Beløb
										'***************************************************************
										call fakventerbelob(thisaktid(x), oRec3("mid"), oRec3("mnavn"), medarbejderTimer, medarbejderTimepris)
										
											
								end if
								
							
							
							'**************************************************************************
							'Medarbejder timepris samme som forrige?? 
							'Kalder sum-aktivitet
							'**************************************************************************
							
								if instr(medarbstrpris, "#"&MedarbejderTimepris&"#") = 0 then
									newa = newa + 1
									Redim preserve aval(newa)
									aval(newa) = MedarbejderTimepris
									call subtotalakt(x, newa)
									sa = sa + 1
								end if
							'Response.write "a: "& newa  &"aval(a): "&aval(newa)&"<br>" 
							'***
							
							
							'** Henter tom sum-akt ***
							call subtotalakt(x, -(n))
							
							
							
							
							'***** Medarbejder timer på akt. timer skt. pris og omsætning  ************	
							%>	
							<tr>
								<td style="border-left:1px #86B5E4 solid;">&nbsp;</td>
								<td><input type="hidden" name="FM_mid_<%=n%>_<%=x%>" value="<%=medarbid%>">
								<%
								if showonfak = 1 then
								chkshow = "CHECKED"
								else
									if (request.cookies("erp")("tvmedarb") = "1" OR request.cookies("erp")("tvmedarb") = "j") AND func <> "red" then
									chkshow = "CHECKED"
									else
									chkshow = ""
									end if
								end if
								
								'request.cookies("tvmedarb") = "j"
								
								'*** finder den næste a- værdi i rækken. Der må ikke være "huller"
								'*** da javascriptet så fejler
								
								if instr(medarbstrpris, "#"&MedarbejderTimepris&"#") = 0 then
								a = newa
								else
									
									for intcounter = 0 to newa - 1
										
										if MedarbejderTimepris = aval(intcounter) then
											a = intcounter
										end if
										
									next
								end if
								
								
								 
								%>
								<input type="hidden" name="FM_m_aval_opr_<%=n%>_<%=x%>" id="FM_m_aval_opr_<%=n%>_<%=x%>" value="<%=a%>">
								<input type="hidden" name="FM_m_aval_<%=n%>_<%=x%>" id="FM_m_aval_<%=n%>_<%=x%>" value="<%=a%>">
								
								<%
								'*** Kun medarbejdere med mere end 0 timer skal være checked, hvis dette ellers er tilvalgt. ****
								if fak = 0 AND func <> "red" then
								chkshow = ""
								else
								chkshow = chkshow
								end if
								%>
								<input type="checkbox" name="FM_show_medspec_<%=n%>_<%=x%>" <%=chkshow%> id="FM_show_medspec_<%=n%>_<%=x%>" value="show">
								<!--<input type="text" name="FM_m_fak_<=n%>_<=x%>" id="FM_m_fak_<=n%>_<=x%>" onKeyup="offsetfak('<=x%>','<=n%>'), tjektimer2('<=x%>','<=n%>'), updantaltimerprakt('<=x%>','<=n%>', 1)" value="<=fak%>" size=3 style="!border: 1px; background-color: #FFFFFF; border-color: limegreen; border-style: solid; font-size:9px">&nbsp;-->
								<input type="text" maxlength="<%=maxl%>" name="FM_m_fak_<%=n%>_<%=x%>" id="FM_m_fak_<%=n%>_<%=x%>" onKeyup="offsetfak('<%=x%>','<%=n%>'), tjektimer2('<%=x%>','<%=n%>'), showborder('<%=x%>','<%=n%>')" value="<%=fak%>" size=4 style="!border: 1px; background-color: #FFFFFF; border-color: limegreen; border-style: solid; font-size:9px">
								<input type="button" name="beregn_<%=n%>_<%=x%>_a" id="beregn_<%=n%>_<%=x%>_a" value="Calc" onClick="enhedsprismedarb('<%=x%>','<%=n%>'), hideborder('<%=x%>','<%=n%>')" style="font-size:8px;">
								&nbsp;&nbsp;&nbsp;
								
								<input type="hidden" name="FM_m_fak_opr_<%=n%>_<%=x%>" id="FM_m_fak_opr_<%=n%>_<%=x%>" value="<%=fak%>">
								
								<%if func <> "red" then%>
								<%if formatnumber(showventer, 2) > 0 then %>
								<font class=roed><b>(<%=showventer%>)</b></font>&nbsp;&nbsp;
								<%else%>
								<img src="../ill/blank.gif" width="20" height="1" border="0">
								<%end if%>
								<input type="text" name="FM_m_vent_<%=n%>_<%=x%>" id="FM_m_vent_<%=n%>_<%=x%>" onKeyup="offsetventer('<%=x%>','<%=n%>'), tjektimer4('<%=x%>','<%=n%>')" value="<%=venter%>" size=1 style="!border: 1px; background-color: #FFFFFF; border-color: silver; border-style: solid; font-size:9px">
								<input type="hidden" name="FM_m_vent_opr_<%=n%>_<%=x%>" id="FM_m_vent_opr_<%=n%>_<%=x%>" value="<%=fak%>">
								<%else%>
							    <input type="text" name="FM_m_vent_<%=n%>_<%=x%>" id="FM_m_vent_<%=n%>_<%=x%>" onKeyup="offsetventer('<%=x%>','<%=n%>'), tjektimer4('<%=x%>','<%=n%>')" value="<%=venter%>" size=1 style="!border: 1px; background-color: #FFFFFF; border-color: silver; border-style: solid; font-size:9px">
								<input type="hidden" name="FM_m_vent_opr_<%=n%>_<%=x%>" id="FM_m_vent_opr_<%=n%>_<%=x%>" value="<%=venter%>">
								<%end if%>
								
								
								</td>
								<td><input type="text" name="FM_m_tekst_<%=n%>_<%=x%>" value="<%=txt%>" size=40 style="!border: 1px; background-color: #FFFFFF; border-color: silver; border-style: solid; font-size:9px">
								<%if func <> "red" then%>
								<%=tp%>
								<%end if%>
								</td>
								<td align=right style="padding-right:10;"><input type="hidden" name="FM_mtimepris_opr_<%=n%>_<%=x%>" id="FM_mtimepris_opr_<%=n%>_<%=x%>" value="<%=medarbejderTimepris%>">
								<input type="text" name="FM_mtimepris_<%=n%>_<%=x%>" id="FM_mtimepris_<%=n%>_<%=x%>" onKeyup="offsetmtp('<%=x%>','<%=n%>'), tjektimer3('<%=x%>','<%=n%>'), showborder('<%=x%>','<%=n%>')" value="<%=medarbejderTimepris%>" size=5 style="!border: 1px; background-color: #FFFFFF; border-color: #999999; border-style: solid; font-size:9px">
								<input type="button" name="beregn_<%=n%>_<%=x%>_b" id="beregn_<%=n%>_<%=x%>_b" value="Calc" onClick="enhedsprismedarb('<%=x%>','<%=n%>'), hideborder('<%=x%>','<%=n%>')" style="font-size:8px;">
								<!-- onFocus="higlightakt('<=x%>','<=n%>','1')" -->
								</td>
								<td style="padding-left:2;">
								<%
								if len(medarbejderBeloeb) <> 0 then
								mbelob = medarbejderBeloeb
								else
								mbelob = 0
								end if
								%>
								<div style="position:relative; left:5; top:0; width:65; border-bottom:1px #8caae6 solid; background-color:#ffffff; padding-right:3px;" align="right" name="medarbbelobdiv_<%=n%>_<%=x%>" id="medarbbelobdiv_<%=n%>_<%=x%>"><b><%=SQLBlessDot(formatnumber(mbelob, 2))%></b></div>
								<input type="hidden" name="FM_mbelob_<%=n%>_<%=x%>" id="FM_mbelob_<%=n%>_<%=x%>" value="<%=SQLBlessDot(formatnumber(mbelob, 2))%>">
								</td>
								<td style="border-right:1px #86B5E4 solid;" width="50">&nbsp;
								<input type="button" value="+" onClick="decimaler('plus', '<%=n%>', '<%=x%>')" style="font-size:9px; background-color:#8caae6; width:14px; height:17;">
							<input type="button" value="-" onClick="decimaler('minus', '<%=n%>', '<%=x%>')" style="font-size:9px; background-color:#8caae6; width:10px; height:17;">
							</td>
							</tr>
							<%
					
					medarbstrpris = medarbstrpris & ", #"&MedarbejderTimepris&"#"
					faktot = faktot + fak
					mbelobtot = mbelobtot + mbelob
					
					fak = 0 
					venter = 0
					medarbejderBeloeb = 0
					n = n + 1
					
					
					Response.flush
					oRec3.movenext
					wend
					oRec3.close
					

					faktot = 0
					mbelobtot = 0
			
			
			call subtotalakt(x, 0)
			%>
			
			<tr><td colspan=6 bgcolor="#ffffff" style="border-left:1px #8caae6 solid; border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="10" alt="" border="0"><br></td></tr>
			<tr><td colspan=6 bgcolor="#ffffff"><img src="../ill/blank.gif" width="1" height="8" alt="" border="0"><br></td></tr>
			
			<%
			Response.write strAktsubtotal
			strAktsubtotal = ""
			%>
			<tr><input type="hidden" name="antal_subtotal_akt_<%=x%>" id="antal_subtotal_akt_<%=x%>" value="<%=sa%>">
			<td colspan=3 bgcolor="#8caae6" style="padding-left:5;"><%=thisaktnavn(x)%> ialt:
			<input type="text" name="timer_subtotal_akt_<%=x%>" id="timer_subtotal_akt_<%=x%>" value="<%=subtotaltimer%>" style='border:0px; font-size:9; width:45; background-color:#d6dff5;'> timer</td>
			<td colspan=3 bgcolor="#8caae6" align=right style="padding-right:5;">
			<input type="text" name="belob_subtotal_akt_<%=x%>" id="belob_subtotal_akt_<%=x%>" value="<%=subtotalbelob%>" style='border:0px; font-size:9; width:55; background-color:#d6dff5;'> kr.</td>
			</tr>
			
			<%if func <> "red" then%>
			<tr>
				<td colspan=6 bgcolor="#d6dff5" style="padding-left:5;">
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
	next
	%>
	<input type="hidden" name="lastactive_x" id="lastactive_x" value="0">
	<input type="hidden" name="lastactive_a" id="lastactive_a" value="0">
	<%
	
	
	'*** antal felter ***
	antalxx = x
	antalnn = n
	%>
	<input type="hidden" name="antal_x" value="<%=antalxx%>">
	<input type="hidden" name="antal_n" id="antal_n" value="<%=antalnn%>">
	
	<%
	if antalxx = 0 AND func = "red" then 
	%>
	<tr bgcolor="#d6dff5">
		<td colspan=6 style="padding:10px;"><br><br><font color=red><b>Denne faktura <u>kladde</u> er oprettet uden der er valgt nogen aktiviteter.</b> <br>
		Derer derfor ikke nogen aktiviteter eller medarbejder timer der kan redigeres.<br><br>
		Slet denne kladde fra faktura-oversigten og opret en ny faktura med dette faktura nummer.</font><br><br>&nbsp;</td>
	<tr>
	<%
	end if
	%>
	
	<tr bgcolor="#d6dff5">
		<td colspan=6 style="border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="20" alt="" border="0"><br></td></tr>
	<tr>
	<td valign="top" style="border-left:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="30" alt="" border="0"></td>
	<td colspan="2" style="padding-left:40; padding-top:10;">
	Timer ialt:
	<%
	'*********************************************************
	'Total timer og beløb
	'*********************************************************
	if intTimer <> 0 then
	intTimer = intTimer 
	else
	intTimer = 0
	end if
	
	%><input type="hidden" name="FM_Timer" value="<%=intTimer%>">
	<div style="position:relative; width:65; border-bottom:2px limegreen solid; background-color:#ffffff; padding-right:3px;" align="right" name="divtimertot" id="divtimertot"><b><%=intTimer%></b></div>
	
	<td colspan="2" style="padding-right:10; padding-top:10;" align="right" style="padding-right:16;">Nettobeløb (subtotal):
		<!-- strBeloeb -->
		<%
		if len(totalbelob) <> 0 then
		thistotbel = SQLBlessDot(formatnumber(totalbelob, 2))
		else
		thistotbel = formatnumber(0, 2)
		end if
		%>
		<input type="hidden" name="FM_beloeb" value="<%=thistotbel%>">
		<div style="position:relative; width:65; border-bottom:2px #86B5E4 solid; background-color:#ffffff; padding-right:3px;" align="right" name="divbelobtot" id="divbelobtot"><b><%=thistotbel%></b></div>
		</td>
		<td valign="top" align=right style="border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="25" alt="" border="0"></td>
	</tr>
	<%
	'***********************************************************
	%>
	<tr>
		<td valign="top" style="border-left:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="190" alt="" border="0"></td>
		<td valign="top" colspan="4" style="padding-left:50;"><br><br><b>Betalings betingelser / Kommentar:</b> <br>
		<br><textarea cols="70" rows="12" name="FM_komm"><%=strKom%></textarea><br>
		<input type="checkbox" name="FM_gembetbet" id="FM_gembetbet" value="1">Gem som standard betalings betingelser for <b>denne</b> kontakt.<br>
		<%if level = 1 then%>
		<input type="checkbox" name="FM_gembetbetalle" id="FM_gembetbetalle" value="1">Gem som standard betalings betingelser for <b>alle</b> kontakter.<br>
		<%end if%>
		<input type="hidden" name="FM_kundeid" id="FM_kundeid" value="<%=intKid%>">
		<br>&nbsp;</td>
		<td valign="top" align=right style="border-right:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="190" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top" colspan="6" style="border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="700" height="1" alt="" border="0"></td>
	</tr>
	</table>
	<br><br>
	<br><br><br><br><br>
	
	<table cellpadding=0 border=0 cellspacing=0>
	<tr>
		<td valign="top">&nbsp;</td>
		<td><br><br>
		
		<!--<div id="oktxt" name=id="oktxt" style="position:relative; visibility:visible; display:; padding:10; background-color:#fffff1; border:1px #5582d2 solid;">
		<a href="#" onClick="showsubm_on()" class=vmenu><h4><font color="forestgreen">V</font>&nbsp;Jeg har gennemset min faktura.</h4>
		 - Jeg har sikret mig at der ikke er nogen "calc" knapper der mangler at blive beregnet<br> og vil gerne godkende den.</a><br><br>
		</div>-->
		
		<div id="subm_off" name=id="subm_off" style="position:relative; visibility:hidden; display:none; padding:10; background-color:#ffff99; left:150px; border:1px red dashed; width:400;">
		<h4><font color="red">!</font>&nbsp;Genberegn timeantal og beløb.</h4> 
		- Dette gøres ved at klikke på de calc-knapper med en rød markering omkring.<br>
		- Når "Calc" funktionen er udført vil "Godkend" knappen blive aktiv igen.</div>
		
		<img src="ill/blank.gif" width="260" height="1" alt="" border="0">
		<input name="subm_on" id="subm_on"  type="image" src="../ill/opretpil_fak.gif">
		
		
		<td valign="top" align=right>&nbsp;</td>
	</tr>
	<!-- Nedenstående Bruges af javascript --->
	<input type="hidden" name="FM_showalert" id="FM_showalert" value="0">
	
	
	
	</form>
	</table>
	<br><br>
