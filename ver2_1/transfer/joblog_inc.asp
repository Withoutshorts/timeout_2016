


<%

public lastmedarbnavn, medarbtimer, medarbEnheder, medarbFeriePlan
	public medarbFrokot, medarbAfspadOp, medarbKm
	
	function fordelpamedarbejdere()
	
	if len(medarbtimer) <> 0 then
	medarbtimer = medarbtimer
	else
	medarbtimer = 0
	end if
	
	  
	if len(medarbEnheder) <> 0 then
	medarbEnheder = medarbEnheder
	else
	medarbEnheder = 0
	end if
	
	if len(medarbFeriePlan) <> 0 then
	medarbFeriePlan = medarbFeriePlan
	else
	medarbFeriePlan = 0
	end if
	
	if len(medarbFrokot) <> 0 then
	medarbFrokot = medarbFrokot
	else
	medarbFrokot = 0
	end if
	
	if len(medarbAfspadOp) <> 0 then
	medarbAfspadOp = medarbAfspadOp
	else
	medarbAfspadOp = 0
	end if
	
	if len(medarbKm) <> 0 then
	medarbKm = medarbKm
	else
	medarbKm = 0
	end if
	               
	                
					
					
					
	%>
					
					<tr>
						<td bgcolor="#d6DFf5" colspan="11"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					<tr>
						<td bgcolor="#ffffe1" align=right colspan="11" height=30 style="padding-right:10;">
						
						<b><%=lastmedarbnavn%>:</b>&nbsp;<b><%=formatnumber(medarbtimer, 2)%></b> timer ~ <b>
						<%=formatnumber(medarbEnheder,2)%></b> enheder
						| <b><%=formatnumber(medarbKm,2) %></b> Km. 
						| Afspad. opsp.: <b><%=formatnumber(medarbAfspadOp,2) %> </b> t. 
	                    | Ferie planlagt: <b><%=formatnumber(medarbFeriePlan, 2) %> </b> t.
	                    | Frokost: <b><%=formatnumber(medarbFrokot,2) %></b> t.
						
						
						
						</td>
					</tr>
	
	<%end function
	
	
	
	public lastaktnavn, akttimer, aktEnheder, lastakttype
	
	function fordelpaakt()
	
	if len(akttimer) <> 0 then
	akttimer = akttimer
	else
	akttimer = 0
	end if
	
	  
	if len(aktEnheder) <> 0 then
	aktEnheder = aktEnheder
	else
	aktEnheder = 0
	end if
	
	
	               
	%>
					
					<tr>
						<td bgcolor="#d6DFf5" colspan="11"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
					</tr>
					<tr>
					    <td bgcolor="#EFf3FF" colspan=4>
                            &nbsp;</td>
						<td bgcolor="#EFf3FF" height=30 style="padding-right:3px;">
						<b><%=lastaktnavn%></b> <font class=lillesort> (<%=lastakttype %>)</font></td>
						<td bgcolor="#EFf3FF" align=right height=30 style="padding-right:15px;">
						<b><%=formatnumber(akttimer, 2)%></b>
						</td>
						<td bgcolor="#EFf3FF" align=right height=30 style="padding-right:15px;">
						<%=formatnumber(aktEnheder,2)%></b>
						</td>
						
						 <td bgcolor="#EFf3FF" colspan=4>
                            &nbsp;</td>
					</tr>
	
	<%end function
	
	
	
	public gtmTimer, gtmEnh, gtmFplan, gtmKm, gtmFro, gtmAfsp 
	function jobtotaler()
	
	   
	%>
	
	<tr>
	<td align=right bgcolor="pink" colspan="11" style="height:30px; padding-right:10; padding-top:3; border-top:1px #cccccc solid;">
	<%
	if len(timerTotpajob) <> 0 then
	timerTotpajob = timerTotpajob
	else
	timerTotpajob = 0
	end if
	
	gtmTimer = gtmTimer + timerTotpajob
	
	
	if len(kmTotpajob) <> 0 then
	kmTotpajob = kmTotpajob
	else
	kmTotpajob = 0
	end if
	
	gtmKm = gtmKm + kmTotpajob
	
	
	
	if len(enhederTotpajob ) <> 0 then
	enhederTotpajob  = enhederTotpajob 
	else
	enhederTotpajob  = 0
	end if
	
	
	gtmEnh = gtmEnh + enhederTotpajob 
	
	
	if len(afspadseringOpsTot) <> 0 then
	afspadseringOpsTot = afspadseringOpsTot
	else
	afspadseringOpsTot = 0
	end if
	
	afspadseringOpsTot = afspadseringOpsTot + gtmAfsp
	
	
	
	if len(timerFrokostTot) <> 0 then
	timerFrokostTot = timerFrokostTot
	else
	timerFrokostTot = 0
	end if
	
	
	timerFrokostTot = timerFrokostTot + gtmFro
	
	
	if len(ferieplanTot) <> 0 then
	ferieplanTot = ferieplanTot
	else
	ferieplanTot = 0
	end if
	
	
	gtmFplan = gtmFplan + ferieplanTot
	
	
	
	%>
	
	
	 <b>Total i periode:</b>&nbsp;<b><%=formatnumber(timerTotpajob, 2)%></b> timer ~ <b><%=formatnumber(enhederTotpajob, 2)%></b> enheder | <b><%=formatnumber(kmTotpajob,2) %></b> Km.
	 | Afspad. opsp.: <b><%=formatnumber(afspadseringOpsTot,2) %> </b> t. 
	 | Ferie planlagt: <b><%=formatnumber(ferieplanTot, 2) %> </b> t.
	 | Frokost: <b><%=formatnumber(timerFrokostTot,2) %></b> t.
	 </td>
	</tr>
    
    </table>

    </div>
    <!-- slut table DIV -->	
    
    <br /><br /><br /><br />
	
	
	
	<%
	end function
	
	
	function tableheader()
	%>
	
	<tr height="20" bgcolor="#8CAAe6">
				       <td style="width:5px;">
                           &nbsp;</td>
				    <td class=alt style="width:20px;"><b>Uge</b></td>
					<td class='alt' align=center style="width:100px; padding-right:10px;"><b>Dato</b></td>
					<%if komprimer = "1" then %>
					<td class='alt' style="width:180px;">Kontakt<br /><b>Jobnavn (Jobnr)</b><br />Jobansv.</td>
					<td class='alt' style="width:280px;"><b>Aktivitet</b> (type)<br />
					kommentar</td>
					<%else %>
					<td class='alt' style="width:10px;">
                        &nbsp;</td>
                    <td class='alt' style="width:450px;"><b>Aktivitet</b> (type)<br />
					kommentar</td>
					<%end if %>
					
					<td class='alt' align=right style="padding-right:15px; width:60px;"><b>Timer / Km / Antal</b><br>Klokkeslet</td>
					<td class='alt' align=right style="padding-right:15px; width:40px;"><b>Enh.</b></td>
					<td class='alt' style="padding-left:10px; width:100px;"><b>Medarb.</b></td>
					<td align="right" class='alt' style="padding-right:5px; width:80px;"><b>Timepris*</b><br>Taste dato</td>
					<td align="right" class='alt' style="width:40px;"><b>Status</b></td>
					<td class='alt' style="width:40px;"><b>Er fak.?</b></td>
				</tr>
	
	<%
	end function
	
	
	sub grandTotal
	
	
	if cint(visopdaterknap) = 1 And print <> "j" then%>
	
	<br />
	<table border="0" width=<%=globalWdt %> cellpadding="0" cellspacing="0">
    <tr>
     <td align=right style="padding:5px 0px 3px 3px; width:804px;">
         <input id="chkalle1" type="checkbox" onclick="checkAll(1)" /> Vælg alle,  <input id="chkalle2" type="checkbox" onclick="checkAll(2)" /> Fravælg alle
       </td>
    <td align=right style="padding:5px 50px 3px 3px; width:200px;">
        <input id="Submit1" type="submit" value="Godkend!" />
    
    </td>
	</tr>
        <input id="antal_v" name="antal_v" value="<%=v %>" type="hidden" />
	</form>
	</table>
	<%end if %>
	<br />
	<br /><h4>Joblog total timer og beløb</h4>  
    
    <%
                	
                tTop = 0
                tLeft = 0
                tWdth = globalWdt


                call tableDiv(tTop,tLeft,tWdth)



                %>          
    <table border="0" width=100% cellpadding="0" cellspacing="0" bgcolor="#ffffff">
    <tr>
	<td align=right colspan="11" valign="top" height=20 bgcolor="#ffffe1" style="padding:10px 10px 10px 10px; border:1px silver solid;">
	<% 
	
	if len(gtmTimer) <> 0 then
	gtmTimer = gtmTimer
	else
	gtmTimer = 0
	end if
	
	if len(gtmFplan) <> 0 then
	gtmFplan = gtmFplan
	else
	gtmFplan = 0
	end if
	
	if len(gtmEnh) <> 0 then
	gtmEnh = gtmEnh
	else
	gtmEnh = 0
	end if
	
	if len(gtmKm) <> 0 then
	gtmKm = gtmKm
	else
	gtmKm = 0
	end if
	
	if len(gtmFro) <> 0 then
	gtmFro = gtmFro
	else
	gtmFro = 0
	end if
	
	if len(gtmAfsp) <> 0 then
	gtmAfsp = gtmAfsp
	else
	gtmAfsp = 0
	end if
	
	%>
	
	Joblog total: 
	<b><%=formatnumber(gtmTimer, 2)%></b> timer ~ <b><%=formatnumber(gtmEnh, 2)%></b> enheder | <b><%=formatnumber(gtmKm,2) %></b> Km.
	 | Afspad. opsp.: <b><%=formatnumber(gtmAfsp,2) %> </b> t. 
	 | Ferie planlagt: <b><%=formatnumber(gtmFplan, 2) %> </b> t.
	 | Frokost: <b><%=formatnumber(gtmFro,2) %></b> t.
	 <br>&nbsp;</td>
	</tr>
	</table>

    </div>
    <!-- slut table DIV -->	
	
	<%
	end sub
	
	
	
	
	public lastfakdato, jobans1, jobans1txt, jobstatus
	public jobstartdato, jobslutdato, jobForkalkulerettimer, rekvnr
	public jobans2, jobans2txt
	', fid, faknr
	
	
	function jobansoglastfak()
	                    
	                    '*** Jobansvarlige ***'
						jobans1 = 0
						strSQL2 = "SELECT mnavn, mnr, mid, job.id AS jid, jobans1, jobstatus, jobstartdato, jobslutdato, budgettimer, ikkebudgettimer, rekvnr FROM job LEFT JOIN medarbejdere ON (mid = jobans1) WHERE job.jobnr = "& oRec("Tjobnr") &""
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobid = oRec2("jid") ''* Bruges til at finde sidste faktura
						jobans1txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						jobstatus = oRec2("jobstatus")
						jobstartdato = oRec2("jobstartdato")
						jobslutdato = oRec2("jobslutdato")
						jobForkalkulerettimer = (oRec2("budgettimer") + oRec2("ikkebudgettimer"))
						rekvnr = oRec2("rekvnr")
						jobans1 = oRec2("mid")
						end if
						oRec2.close
						
						
						'*** jobANS 2 **'
						jobans2 = 0
						strSQL2 = "SELECT mnavn, mnr, mid, jobans2 FROM job, medarbejdere WHERE job.jobnr = "& oRec("Tjobnr") &" AND mid = jobans2"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans2txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						jobans2 = oRec2("mid")
						end if
						oRec2.close
						
						
						
						'*** Lastfakdato **'
						'** tastedato **'
						if komprimeret = "1" then
						fakdatoSQLkri = year(oRec("Tdato")) & "/" & month(oRec("Tdato")) & "/" & day(oRec("Tdato")) 
						else
						fakdatoSQLkri = strAar&"/"&strMrd&"/"&strDag
						end if
					
						lastfakdato = "1/1/2001"
						fid = 0
						faknr = 0
					
						strSQLFAK = "SELECT f.fakdato, f.fid, f.faknr FROM fakturaer f WHERE f.jobid = "& jobid &" AND f.fakdato >= '"& fakdatoSQLkri &"' AND faktype = 0 ORDER BY f.fakdato DESC"
						
						'Response.Write strSQLFAK
						'Response.flush
						
						oRec2.open strSQLFAK, oConn, 3
						if not oRec2.EOF then
							if len(trim(oRec2("fakdato"))) <> 0 then
							lastfakdato = oRec2("fakdato")
							'fid = oRec2("fid")
							'faknr = oRec2("faknr")
							end if
						end if
						oRec2.close
	
	
	end function
	%>	

