strSQL3 = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, id AS Akt, job FROM aktiviteter WHERE id = "& oRec("Aid") &""
	'strSQL3 = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, id AS Akt, job FROM aktiviteter WHERE id = "& oRec("Aid") &" AND projektgruppe1 = "& brugergruppeId(intcounter) &" OR id = "& oRec("Aid") &" AND projektgruppe2 = "& brugergruppeId(intcounter) &" OR id = "& oRec("Aid") &" AND projektgruppe3 = "& brugergruppeId(intcounter) &" OR id = "& oRec("Aid") &" AND projektgruppe4 = "& brugergruppeId(intcounter) &" OR id = "& oRec("Aid") &" AND projektgruppe5 = "& brugergruppeId(intcounter) &""
	oRec3.open strSQL3, oConn, 3
	
	while not oRec3.EOF 
	
			if searchstring = "0" then
		 	
				if oRec("Jobnr") = strLastjobnr then
				Response.write "<tr><td style='!border: 1px; background-color: #ffffff; border-color: #000000; border-style: solid; padding-left : 4px; padding-right : 4px;'>"& oRec3("Akt") & left(oRec("navn"), 20)&"</td>"
				else
					if de > "1" then
					Response.write "</table></div></td></tr>"
					end if
				Response.write "<tr><td width=80 style='!border: 1px; background-color: #cccccc; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><b>"_
				& oRec("Jobnr") &"</b></td><td width=420 height=20 style='!border: 1px; background-color: #cccccc; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><b>"
				%>
				<a href="javascript:expand('<%=de%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="Menub<%=de%>"></a>&nbsp;&nbsp;&nbsp;
				<%
				Response.write oRec("Jobnavn")&"</b>&nbsp;&nbsp;&nbsp;<font size=1>("&oRec("kkundenavn")&")</font></td></tr>"
				%><tr><td colspan="2"><DIV ID="Menu<%=de%>" Style="position: relative; display: none;">
				<table cellspacing="1" cellpadding="0" border="0" width="100%">
				<tr>
				<%
				Response.write "<td width=120 style='!border: 1px; background-color: #ffffff; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'>"& oRec3("Akt") &left(oRec("navn"), 20)&"</td>"
				end if
			   
			else
			
			 	if oRec("Jobnr") = strLastjobnr AND searchstring = cint(oRec("jobknr")) then
				Response.write "<tr><td style='!border: 1px; background-color: #ffffff; border-color: #000000; border-style: solid; padding-left : 4px; padding-right : 4px;'>"&left(oRec("navn"), 20)&"</td>"
				else
					if cint(Searchstring) = cint(oRec("jobknr")) then
						if deSearchNum > "1" then
						Response.write "</table></div></td></tr>"
						end if
						Response.write "<tr><td width=80 style='!border: 1px; background-color: #cccccc; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><b>"_
						& oRec("Jobnr") &"</b></td><td width=420 height=20 style='!border: 1px; background-color: #cccccc; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><b>"
						%>
						<a href="javascript:expand('<%=de%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="Menub<%=de%>"></a>&nbsp;&nbsp;&nbsp;
						<%
						Response.write oRec("Jobnavn")&"</b>&nbsp;&nbsp;&nbsp;<font size=1>("&oRec("kkundenavn")&")</font></td></tr>"
						%>
						<tr><td colspan=2><DIV ID="Menu<%=de%>" Style="position: relative; display: none;">
						<table cellspacing="1" cellpadding="0" border="0" width="100%">
						<tr>
						<%
						Response.write "<td width=120 style='!border: 1px; background-color: #ffffff; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'>"&left(oRec("navn"), 20)&"</td>"
						deSearchNum = deSearchNum + 1
					end if
				end if
			end if
			
		 	%>
			<input type="hidden" name="FM_jobnr_<%=de%>" value="<%=oRec("jobnr")%>">
			<input type="hidden" name="FM_Aid_<%=de%>" value="<%=oRec("Aid")%>">
			<%
			varTjDatoUS = cdate(convertDate(tjekdag(1)&"/"&strRegaar)) 
			
			oRec2.Open "SELECT Timer FROM timer WHERE TAktivitetId = "& oRec("Aid") &" AND Tjobnr = '"&oRec("Jobnr")&"' AND Tmnavn = '"&Session("user")&"' AND Tdato = #"& varTjDatoUS &"#", oConn, 3  
			if not oRec2.EOF then	
			sonTimerVal = oRec2("Timer")
			else
			sonTimerVal = ""
			end if
			oRec2.close
			
			varTjDatoUS = cdate(convertDate(tjekdag(2)&"/"&strRegaar)) 
			
			oRec2.Open "SELECT Timer FROM timer WHERE TAktivitetId = "& oRec("Aid") &" AND Tjobnr = '"&oRec("Jobnr")&"' AND Tmnavn = '"&Session("user")&"' AND Tdato = #"& varTjDatoUS &"#", oConn, 3  
			if not oRec2.EOF then	
			manTimerVal = oRec2("Timer")
			else
			manTimerVal = ""
			end if
			oRec2.close
			
			varTjDatoUS = cdate(convertDate(tjekdag(3)&"/"&strRegaar)) 
			
			oRec2.Open "SELECT Timer FROM timer WHERE TAktivitetId = "& oRec("Aid") &" AND Tjobnr = '"&oRec("Jobnr")&"' AND Tmnavn = '"&Session("user")&"' AND Tdato = #"& varTjDatoUS &"#", oConn, 3  
			if not oRec2.EOF then	
			tirTimerVal = oRec2("Timer")
			else
			tirTimerVal = ""
			end if
			oRec2.close
			
			varTjDatoUS = cdate(convertDate(tjekdag(4)&"/"&strRegaar)) 
			
			oRec2.Open "SELECT Timer FROM timer WHERE TAktivitetId = "& oRec("Aid") &" AND Tjobnr = '"&oRec("Jobnr")&"' AND Tmnavn = '"&Session("user")&"' AND Tdato = #"& varTjDatoUS &"#", oConn, 3  
			if not oRec2.EOF then	
			onsTimerVal = oRec2("Timer")
			else
			onsTimerVal = ""
			end if
			oRec2.close
			
			varTjDatoUS = cdate(convertDate(tjekdag(5)&"/"&strRegaar)) 
			
			oRec2.Open "SELECT Timer FROM timer WHERE TAktivitetId = "& oRec("Aid") &" AND Tjobnr = '"&oRec("Jobnr")&"' AND Tmnavn = '"&Session("user")&"' AND Tdato = #"& varTjDatoUS &"#", oConn, 3  
			if not oRec2.EOF then	
			torTimerVal = oRec2("Timer")
			else
			torTimerVal = ""
			end if
			oRec2.close
			
			varTjDatoUS = cdate(convertDate(tjekdag(6)&"/"&strRegaar)) 
			
			oRec2.Open "SELECT Timer FROM timer WHERE TAktivitetId = "& oRec("Aid") &" AND Tjobnr = '"&oRec("Jobnr")&"' AND Tmnavn = '"&Session("user")&"' AND Tdato = #"& varTjDatoUS &"#", oConn, 3  
			if not oRec2.EOF then	
			freTimerVal = oRec2("Timer")
			else
			freTimerVal = ""
			end if
			oRec2.close
			
			varTjDatoUS = cdate(convertDate(tjekdag(7)&"/"&strRegaar)) 
			
			
			oRec2.Open "SELECT Timer FROM timer WHERE TAktivitetId = "& oRec("Aid") &" AND Tjobnr = '"&oRec("Jobnr")&"' AND Tmnavn = '"&Session("user")&"' AND Tdato = #"& varTjDatoUS &"#", oConn, 3  
			if not oRec2.EOF then	
			lorTimerVal = oRec2("Timer")
			else
			lorTimerVal = ""
			end if
			oRec2.close
			
			if searchstring = "0" then%>
			<td align="center"><input type="hidden" name="FM_son_opr_<%=de%>" value="<%=sonTimerVal%>"><input type="Text" name="Timer_son_<%=de%>" maxlength="5" value="<%=sonTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'son'), tjektimer('son',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #cd853f; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_man_opr_<%=de%>" value="<%=manTimerVal%>"><input type="Text" name="Timer_man_<%=de%>" maxlength="5" value="<%=manTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'man'), tjektimer('man',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_tir_opr_<%=de%>" value="<%=tirTimerVal%>"><input type="Text" name="Timer_tir_<%=de%>" maxlength="5" value="<%=tirTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'tir'), tjektimer('tir',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_ons_opr_<%=de%>" value="<%=onsTimerVal%>"><input type="Text" name="Timer_ons_<%=de%>" maxlength="5" value="<%=onsTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'ons'), tjektimer('ons',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #cccccc; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_tor_opr_<%=de%>" value="<%=torTimerVal%>"><input type="Text" name="Timer_tor_<%=de%>" maxlength="5" value="<%=torTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'tor'), tjektimer('tor',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_fre_opr_<%=de%>" value="<%=freTimerVal%>"><input type="Text" name="Timer_fre_<%=de%>" maxlength="5" value="<%=freTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'fre'), tjektimer('fre',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_lor_opr_<%=de%>" value="<%=lorTimerVal%>"><input type="Text" name="Timer_lor_<%=de%>" maxlength="5" value="<%=lorTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'lor'), tjektimer('lor',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #cd853f; border-style: solid;"></td>
			</tr>
			<%
			else
			if searchstring = cint(oRec("jobknr")) then%>
			<td align="center"><input type="hidden" name="FM_son_opr_<%=de%>" value="<%=sonTimerVal%>"><input type="Text" name="Timer_son_<%=de%>" maxlength="5" value="<%=sonTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'son'), tjektimer('son',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #cd853f; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_man_opr_<%=de%>" value="<%=manTimerVal%>"><input type="Text" name="Timer_man_<%=de%>" maxlength="5" value="<%=manTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'man'), tjektimer('man',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_tir_opr_<%=de%>" value="<%=tirTimerVal%>"><input type="Text" name="Timer_tir_<%=de%>" maxlength="5" value="<%=tirTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'tir'), tjektimer('tir',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_ons_opr_<%=de%>" value="<%=onsTimerVal%>"><input type="Text" name="Timer_ons_<%=de%>" maxlength="5" value="<%=onsTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'ons'), tjektimer('ons',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_tor_opr_<%=de%>" value="<%=torTimerVal%>"><input type="Text" name="Timer_tor_<%=de%>" maxlength="5" value="<%=torTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'tor'), tjektimer('tor',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_fre_opr_<%=de%>" value="<%=freTimerVal%>"><input type="Text" name="Timer_fre_<%=de%>" maxlength="5" value="<%=freTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'fre'), tjektimer('fre',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;"></td>
			<td align="center"><input type="hidden" name="FM_lor_opr_<%=de%>" value="<%=lorTimerVal%>"><input type="Text" name="Timer_lor_<%=de%>" maxlength="5" value="<%=lorTimerVal%>" size="6" onkeyup="setTimerTot(<%=de%>, 'lor'), tjektimer('lor',<%=de%>)"; style="!border: 1px; background-color: #FFFFFF; border-color: #cd853f; border-style: solid;"></td>
			</tr>
			<%
			end if
			end if
			
			if len(sonTimerVal) <> "0" then
			sonTimerTot = sonTimerTot + sonTimerVal
			end if
			
			if len(manTimerVal) <> "0" then
			manTimerTot = manTimerTot + manTimerVal
			end if
			
			if len(tirTimerVal) <> "0" then
			tirTimerTot = tirTimerTot + tirTimerVal
			end if
			
			if len(onsTimerVal) <> "0" then
			onsTimerTot = onsTimerTot + onsTimerVal
			end if
			
			if len(torTimerVal) <> "0" then
			torTimerTot = torTimerTot + torTimerVal
			end if
			
			if len(freTimerVal) <> "0" then
			freTimerTot = freTimerTot + freTimerVal
			end if
			
			if len(lorTimerVal) <> "0" then
			lorTimerTot = lorTimerTot + lorTimerVal
			end if
			
			de = de + 1
			strLastjobnr = oRec("Jobnr")
			
		
		'** Slut Loop på Aktiviteter**
		'end if
		oRec3.movenext
		wend
		oRec3.close
		'end if
		
	'oRec.MoveNext
	'Wend 
	'end if
	
	oRec.Close
	
	'next
	
	
	
	
	
	'******* Slut gennemløb af Eksterne Job og Aktiviteter *******************
	
	if searchstring = "0" then
	varShowAnyJob = de
	else
	varShowAnyJob = deSearchNum
	end if
	
	if varShowAnyJob = 1 then
	%>
	<tr><td colspan="2"><font color=DarkRed>Du er ikke blevet tildelt adgang til nogen job.</font><br>
	<DIV Style="position: relative; display: none;">
				<table cellspacing="1" cellpadding="0" border="0" width="100%">
				<tr><td></td></tr>
	<%
	end if
	Response.write "</table></div></td></tr></table>"
	
	
	antal_de = de - 1
	
	Response.write "<table width='530' border='0' cellspacing='1' cellpadding='0'>"
	
	%><input type="hidden" name="FM_de" value="<%=antal_de%>">
	<%
	'** Udskriver job uden tilknyttede aktiviteter **
	areJobPrinted = 0
	showHeader = 0
	for intcounter = 0 to f - 1  
	
	strSQL = "SELECT job.id AS id, jobnr, jobnavn, jobknr, kkundenavn FROM job, kunder WHERE jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = jobknr AND job.projektgruppe1 = "&brugergruppeId(intcounter)&""_
	&" UNION ALL"_
	&" SELECT job.id AS id, jobnr, jobnavn, jobknr, kkundenavn FROM job, kunder WHERE jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = jobknr AND job.projektgruppe2 = "&brugergruppeId(intcounter)&""_
	&" UNION ALL"_
	&" SELECT job.id AS id, jobnr, jobnavn, jobknr, kkundenavn FROM job, kunder WHERE jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = jobknr AND job.projektgruppe3 = "&brugergruppeId(intcounter)&""_
	&" UNION ALL"_
	&" SELECT job.id AS id, jobnr, jobnavn, jobknr, kkundenavn FROM job, kunder WHERE jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = jobknr AND job.projektgruppe4 = "&brugergruppeId(intcounter)&""_
	&" UNION ALL"_
	&" SELECT job.id AS id, jobnr, jobnavn, jobknr, kkundenavn FROM job, kunder WHERE jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = jobknr AND job.projektgruppe5 = "&brugergruppeId(intcounter)&""_
	&" ORDER BY jobnavn"
	
	oRec.open strSQL, oConn, 3
	While Not oRec.EOF
	strSQL = "SELECT navn, job FROM aktiviteter WHERE job = " & oRec("id") & ""
	oRec2.open strSQL, oConn, 3
	
	areJobPrinted = instr(jobPrinted, ", "&oRec("Jobnr")&"#")
	
	if oRec2.EOF AND areJobPrinted = "0" then
	jobPrinted = jobPrinted &", "& oRec("Jobnr")&"#"
	if showHeader = 0 then
	Response.write "<tr><td height=20 colspan=9>&nbsp;</td>"
	Response.write "</tr>"
	Response.write "<tr><td height=20 colspan=9 style='!border: 1px; background-color: #ffffff; border-color: lightBlue; border-style: solid; padding-left : 4px;padding-right : 4px;'>Job uden aktiviteter:</td>"
	Response.write "</tr>"
	end if 
	Response.write "<tr><td style='!border: 1px; background-color: #cccccc; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><b>"_
	& oRec("Jobnr") &"</b></td><td width=420 height=20 colspan=8 style='!border: 1px; background-color: #cccccc; border-color: #000000; border-style: solid; padding-left : 4px;padding-right : 4px;'><a href='jobs.asp?menu=job'>"&oRec("jobnavn")&"</a>&nbsp;&nbsp;&nbsp;<font size=1>("&oRec("kkundenavn")&")</font></td>"
	Response.write "</tr>"
	showHeader = 1
	end if
	oRec2.close
	oRec.movenext
	wend
	oRec.close
	
	next
	%>
	</table>
	
