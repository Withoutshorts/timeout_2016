


<%


    sub underskrift


    %>
                                <div style="position:relative; padding:40px; width:800px;">

                                <table width="100%" border="0" cellpadding="0" cellspacing="0"><tr><td>
                                Dato & underskrift medarbejder</td>
                                    <td style="width:200px; border-bottom:1px #000000 dashed;">&nbsp;</td>
                                    <td style="width:200px;">&nbsp;</td>
                                    <td style="">Dato & underskift (overordnet)</td>
                                    <td style="width:200px; border-bottom:1px #000000 dashed;">&nbsp;</td>

                                       </tr></table>

                            </div>
    <%

    end sub


        sub underskriftExcel
        
                    ekspTxt = ekspTxt & "xx99123sy#z" & "xx99123sy#z" & "xx99123sy#z" & "xx99123sy#z" & "xx99123sy#z"  
                         ekspTxt = ekspTxt & ";;Underskrift medarbejder:;;;;;;;;Dato & Underskift (overordnet);;;;;;;;;;;;;;;;;;;;;;;;"
                       
        end sub


        public timerTotprDagGT, timerTotDenneMedarb
        redim timerTotprDagGT(31)
        function medarbMdtot()
        timerTotDenneMedarb = 0

        %>
                                        <tr style="background-color:#FFC0CB;">
                                            <td colspan="2" style="border-top:1px #cccccc solid;">Total pr. dag:</td>
                                                <%
                                               for d = 1 to vlgtMdLastDay 
                                    
                                                   %>
                              
                                                    <td style="width:30px; border-left:1px #999999 solid; border-top:1px #cccccc solid;" align="right"><%=formatnumber(timerTotprDag(d), 2) %></td>

                                                <%
                                                timerTotprDagGT(d) = timerTotprDagGT(d) + timerTotprDag(d)
                                                timerTotDenneMedarb = timerTotDenneMedarb + timerTotprDag(d)
                                                 timerTotprDag(d) = 0   
                                                 next %>

                                            
                                        </tr>
                                        <tr>
                                            <td colspan="2" style="border-top:1px #cccccc solid;"><b>Total m�ned:</b><br />&nbsp;</td>
                                            <td colspan="<%=vlgtMdLastDay %>" align="right" style="border-top:1px #cccccc solid;"><b><%=formatnumber(timerTotDenneMedarb, 2) %></b><br />&nbsp;</td>
                                               
                                        </tr>


                                        <%

    
        end function



        public timerTotprDagJob
        redim timerTotprDagJob(31)
        function jobMdtot()
        timerTotDetteJob = 0

        %>
                                     
                                                <%
                                               for d = 1 to vlgtMdLastDay 
                                    
                                                 timerTotDetteJob = timerTotDetteJob + timerTotprDagJob(d)
                                                 timerTotprDagJob(d) = 0   
                                                 next %>

                                            
                                      
                                        <tr style="background-color:#F4F4F4;">
                                            <td colspan="2" style="border-top:1px #cccccc solid;">Total p� job:</td>
                                            <td colspan="<%=vlgtMdLastDay %>" style="border-top:1px #cccccc solid;" align="right"><%=formatnumber(timerTotDetteJob, 2) %></td>
                                               
                                        </tr>


                                        <%

    
        end function




         function medarbMdtotGT()
        timerTotGT = 0

        %>
                                        <tr style="background-color:#FFFFFF;">
                                            <td colspan="2"><b>Grandtotal:</b></td>
                                                <%
                                               for d = 1 to vlgtMdLastDay 
                                    
                                                   %>
                              
                                                    <td style="width:30px; border-left:1px #999999 solid;" align="right"><%=formatnumber(timerTotprDagGT(d), 2) %></td>

                                                <%
                                                timerTotGT = timerTotGT + timerTotprDagGT(d)
                                                timerTotprDagGT(d) = 0
                                                 
                                                 next %>


                                        </tr>
                                         <tr style="background-color:#cccccc;">
                                            <td colspan="2"><b>Grandtotal m�ned:</b><br />&nbsp;</td>
                                            <td colspan="<%=vlgtMdLastDay %>" align="right"><%=formatnumber(timerTotGT, 2) %><br />&nbsp;</td>
                                               
                                        </tr>

                                        <%


        end function






public lastmedarbnavn, medarbtimer, medarbEnheder, medarbFeriePlan
	public medarbFrokot, medarbAfspadOp, medarbKm
	
	function fordelpamedarbejdere()
	
	
	                
					
					
					
	%>
					
					
					<tr>
						 <td align=right bgcolor="#FFFFFF" colspan="9">&nbsp;</td>
	                     <td bgcolor="#FFFFFF" colspan="7" style="padding:5px 10px 5px 0px; border-top:1px #999999 solid;">
						Medarbejder<br />
						<b><%=lastmedarbnavn%>:</b>&nbsp;
						<%
	
	mtotArr = split(replace(akttype_sel, "#", ""), ",")
	for j = 0 to ubound(mtotArr)
	
	call akttyper(mtotArr(j), 1)
	
	if len(trim(vlgt_typerTotalerMed(mtotArr(j)))) <> 0 then
	antal = vlgt_typerTotalerMed(mtotArr(j))
	enh = vlgt_typerTotalerMedEnh(mtotArr(j))
	else
	antal = 0
	enh = 0
	end if
	
	if antal <> 0 then
	Response.write "<b>"& formatnumber(antal, 2)& "</b>" 'akttypenavn & 
	
	    if cint(hideenheder) = 0 then
	    Response.Write " ~ ("& formatnumber(enh, 2) & ")"
	    end if
	 
	 
	 Response.Write "<br>"   
	
	end if
	
	vlgt_typerTotalerMed(mtotArr(j)) = 0
	vlgt_typerTotalerMedEnh(mtotArr(j)) = 0
	
	next %>
				
						
						
						<br />&nbsp;</td>
					</tr>
	
	<%end function
	
	
	
	public lastaktnavn, akttimer, aktEnheder, lastakttype, faseCounta, faseCountenh
	
	function fordelpaakt(sft)
	
	
	
	
	               
	%>
					
					<tr>
						 <td align=right bgcolor="#FFFFFF" colspan="9">&nbsp;</td>
	                     <td bgcolor="#FFFFFF" colspan="7" style="padding:5px 10px 5px 0px; border-top:1px #999999 solid;">
					
                        Aktivitet<br /> <b><%=lastaktnavn%>:</b>
						<%
	
	atotArr = split(replace(akttype_sel, "#", ""), ",")
	for j = 0 to ubound(atotArr)
	
	call akttyper(atotArr(j), 1)
	
	if len(trim(vlgt_typerTotalerAkt(atotArr(j)))) <> 0 then
	antal = vlgt_typerTotalerAkt(atotArr(j))
	enh = vlgt_typerTotalerAktEnh(atotArr(j))
	else
	antal = 0
	enh = 0
	end if
	
	if antal <> 0 then
	Response.write formatnumber(antal, 2)
	    
	    if cint(hideenheder) = 0 then
	    Response.Write " ~ ("& formatnumber(enh, 2) & ")"
	    end if
	 
	 response.write "<br><span style=""color:#999999; font-size:9px;"">("& akttypenavn &")</span>"

	 'Response.Write "<br>"
	 
	 faseCounta = faseCounta + antal
	faseCountenh = faseCountenh + enh 
	
	end if
	
	vlgt_typerTotalerAkt(atotArr(j)) = 0
	vlgt_typerTotalerAktEnh(atotArr(j)) = 0
	
	next %>

           <br />&nbsp;</td>                 
					</tr>
	
	
	<%if (lcase(trim(lastFase)) <> lcase(trim(thisFase)) AND len(trim(lastFase)) <> 0 AND isNull(lastFase) <> true) OR (sft = 1 AND isNull(lastFase) <> true) AND hidefase <> 1 then  %>
	<%'if t <> 100 then %>
	<tr>
        
	     <td align=right bgcolor="#FFFFFF" colspan="9">&nbsp;</td>
	     <td bgcolor="#FFFFFF" colspan="7" style="padding:5px 10px 5px 0px; border-top:1px #999999 dashed;">Fase<br /> <b><%=replace(lastFase, "_", " ") %>:</b>&nbsp; <b><%=formatnumber(faseCounta, 2) %></b> ~ (<%=formatnumber(faseCountenh, 2) %>)<br /><br />&nbsp; </td>
	</tr>
	<%
    faseCounta = 0
	faseCountenh = 0

    'else
    
        'if (lcase(trim(lastFase)) <> lcase(trim(thisFase)) AND len(trim(lastFase)) <> 0) then
	    'faseCounta = 0
	    'faseCountenh = 0
	    'end if

	end if 
    
    
        if (len(trim(lastFase)) = 0) OR isNull(lastFase) = true then
	    faseCounta = 0
	    faseCountenh = 0
	    end if
    
    %>
	
	
	<%end function
	
	
	
	public gtmTimer, gtmEnh, gtmFplan, gtmKm, gtmFro, gtmAfsp 
	function jobtotaler()
	
	   
	%>
	
	<tr>
    <td align=right bgcolor="#FFFFFF" colspan="9">&nbsp;</td>
	<td bgcolor="#FFFFFF" colspan="7" style="padding:5px 10px 5px 0px; border-top:1px #999999 solid;">
	
	<%if cint(joblog_uge) = 2 then%>
	<b>Uge <%=lastWeekNum %></b>&nbsp;
	<%else %>
	Job<br /> <b><%=lastjobnavn%></b>&nbsp;
	 <%end if

     
        
	
	jtotArr = split(replace(akttype_sel, "#", ""), ",")
	for j = 0 to ubound(jtotArr)
	
	call akttyper(jtotArr(j), 1)

         'if jtotArr(j) = 1 then
         'akttypenavn = "Fakturerbar"
         'end if

         'if jtotArr(j) = 2 then
         'akttypenavn = "Ikke fakturerbar"
         'end if

         'if jtotArr(j) > 2 then
         'akttypenavn = "Anden type"
         'end if
	
	if len(trim(vlgt_typerTotaler(jtotArr(j)))) <> 0 then
	antal = vlgt_typerTotaler(jtotArr(j))
	enh = vlgt_typerTotalerEnh(jtotArr(j))
	else
	antal = 0
	enh = 0
	end if
	
	if antal <> 0 then
        
        'if cint(joblog_uge) = 2 then
	    Response.write "<br><span style=""color:#999999;"">"& akttypenavn & ":</span>  "& formatnumber(antal, 2) 
        'else
        'Response.write ""& akttypenavn &"  <b>"& formatnumber(antal, 2)& "</b>"
        'end if
	        
	    if cint(hideenheder) = 0 then
	    Response.Write " ~ ("& formatnumber(enh, 2) & ")"
	    end if
	 
	 
	 'Response.Write "<br>"
	        
	end if
	
	vlgt_typerTotaler(jtotArr(j)) = 0
	vlgt_typerTotalerEnh(jtotArr(j)) = 0
	
	next %>
	
	
	 <br />&nbsp;
	 </td>
	</tr>
	</table>

    </div>
    <!-- slut table DIV -->	
    
    <br /><br /><br /><br />
	
	
	
	<%
	end function
	
	
	public kmTotdag,timerTotdag,enhederTotdag,afspadTotdag,frokostTotdag,fplanTotdag
	function dagstotaler(visning)
	
					
					
	if visning <> 1 AND visning <> 0 then				
	%>
				
					<tr>
					    <td bgcolor="#ffffff" align=right colspan="9">
						&nbsp;</td>
						<td bgcolor="#ffffff" colspan="7" style="padding:5px 10px 5px 0px; border-top:1px #999999 solid;">
						
						<b><%=left(weekdayname(weekday(lastdate )), 3) %>. <%=day(lastdate ) &" "&left(monthname(month(lastdate )), 3) &". "& right(year(lastdate ), 2)%></b><br />
						<%=lastmedarbnavn%>:&nbsp; 
						<%
	
	mtotArr = split(replace(akttype_sel, "#", ""), ",")
	for j = 0 to ubound(mtotArr)
	
	call akttyper(mtotArr(j), 1)
	
	if len(trim(vlgt_typerTotalerMed(mtotArr(j)))) <> 0 then
	antal = vlgt_typerTotalerMed(mtotArr(j))
	enh = vlgt_typerTotalerMedEnh(mtotArr(j))
	else
	antal = 0
	enh = 0
	end if
	
	if antal <> 0 then
	Response.write "<b>"& formatnumber(antal, 2)& "</b>" 'akttypenavn & ":
	    
	    if cint(hideenheder) = 0 then
	    Response.Write " ~ ("& formatnumber(enh, 2) & ")"
	    end if
	 
	 
	 Response.Write "<br>"
	
	end if
	
	vlgt_typerTotalerMed(mtotArr(j)) = 0
	vlgt_typerTotalerMedEnh(mtotArr(j)) = 0
	
	next 
	
	end if
	%>
						
						
						
					<br />&nbsp;	</td>
					</tr>
					
	<%if visning <> 3 then
	
	
	
	    if (datepart("ww", oRec("tdato"), 2, 2) = datepart("ww", lastdate, 2, 2)) AND lastmedarb = oRec("Tmnr") then %>				
	    <tr>
						    <td bgcolor="#Eff3ff" colspan="16" style="height:20px; padding:5px 5px 5px 5px;">
						    <b><%=left(weekdayname(weekday(oRec("tdato"))), 3) %>. <%=day(oRec("tdato")) &" "&left(monthname(month(oRec("tdato"))), 3) &". "& right(oRec("tdato"), 2)%></b>
						    </td>
					    </tr>
        
        
        <%
        end if
    
    end if
    
    
    if visning = 0 then %>				
	<tr>
						<td bgcolor="#Eff3ff" colspan="16" style="height:20px; padding:5px 5px 5px 5px;">
						<b><%=left(weekdayname(weekday(oRec("tdato"))), 3) %>. <%=day(oRec("tdato")) &" "&left(monthname(month(oRec("tdato"))), 3) &". "& right(oRec("tdato"), 2)%></b>
						</td>
					</tr>
    
    
    <%
    end if
    
	end function 
	
	
	function tableheader()
	%>
	
	<tr height="20" bgcolor="#8CAAe6">
				       <td style="width:0px;">
                           &nbsp;</td>
				    <td valign=bottom class=alt><b>Uge</b></td>
					<td valign=bottom style="padding-left:10px;" class='alt'><b>Dato</b></td>
					<%if komprimer = "1" then %>
					<td valign=bottom class='alt'><b>Kunde</b><br />Jobnavn (Jobnr)<br />Jobansv.</td>

                     <%if hidefase <> 1 then %>
					<td valign=bottom class='alt'><b>Fase</b></td>
                    <%else %>
                    <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
                    <%end if %>

					<td valign=bottom class='alt'><b>Aktivitet</b> (type)<br />
					kommentar</td>
					<%else %>
					<td valign=bottom class='alt'>
                        &nbsp;</td>
                        <%if hidefase <> 1 then %>
                        <td valign=bottom class='alt'><b>Fase</b></td>
                        <%else %>
                        <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
                        <%end if %>
                    <td valign=bottom class='alt'><b>Aktivitet</b> (type)<br />
					kommentar</td>
					<%end if %>

                    <%if cint(showfor) = 1 then%>
                    <td valign=bottom class='alt' style="padding-right:5px;"><b>Forretnings-omr.</b></td>
					<%else %>
					<!--<td>&nbsp;</td>-->
					<%end if %>
					
					<td valign=bottom class='alt' style="padding-left:5px;"><b>Medarb.</b></td>
					
					
					<td valign=bottom class='alt' align=right style="padding-right:5px;">
					<%if print = "j" then
					%>
					<b>Timer <br /> Km  <br /> Antal</b>
					<%
					else %>
					<b>Timer / Km / Antal</b>
					<%end if %>
					<br>Klokkeslet</td>
					
					<%if cint(hideenheder) = 0 then %>
					<td valign=bottom class='alt' align=right style="padding-right:5px;">~ Enheder</td>
					<%else %>
					<td>&nbsp;</td>
					<%end if %>
					
					
					<%if (level <= 2 OR level = 6) AND cint(hidetimepriser) = 0 then  %>
					<td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b>Time / Stk. pris*</b></td>
					<td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b>Pris ialt</b></td>
					<%else %>
					<td>&nbsp;</td>
				    <td>&nbsp;</td>
					<%end if %>
					
					<%if level = 1 AND visKost = 1 then  %>
					<td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b>Kostpris</b></td>
					<td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b>Kostpris ialt</b></td>
					<%else %>
					<td>&nbsp;</td>
				    <td>&nbsp;</td>
					<%end if %>
					
					
					
					<%if cint(hidegkfakstat) <> 1 then  %>
                    <td valign=bottom align="right" class='alt' style="padding-right:5px;">
					Tastedato<br />Indtastet af</td>
					<td valign=bottom align="center" class='alt'><b>Status</b><br />Godkendt<br />Afvist</td>
					<td valign=bottom class='alt' align="center" style="padding-right:1px;"><b>Faktura</b></td>
				    <%else %>
                    <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				    <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
                    <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				    <%end if %>   
				 </tr>
	
	<%
	end function
	
	
	public grgrTotal
	sub grandTotal
	
	
	if cint(visopdaterknap) = 1 And print <> "j" AND cint(hidegkfakstat) <> 1 then%>
	
	<br />
	<table border="0" width=<%=globalWdt %> cellpadding="0" cellspacing="0">
    <tr>
     <td align=right style="padding:5px 20px 1px 1px;">
        <b>Opdater Status:</b><br /><input id="chkalle1" type="radio" onclick="checkAllGK(1)" /> Godkend alle<br />
        <input id="chkalle2" type="radio" onclick="checkAllGK(2)" /> Afvis alle  <br /> 
        <input id="chkalle0" type="radio" onclick="checkAllGK(0)" /> S�t alle = ingen.
       <br />
       <input id="Submit1" type="submit" value="Godkend!" /></td>
   
	</tr>
        <input id="antal_v" name="antal_v" value="<%=v %>" type="hidden" />
	</form>
	</table>
	<%end if %>
	<br />
	<br /><h4>Joblog Grandtotal</h4>  
    
    <%
                	
                tTop = 0
                tLeft = 0
                tWdth = globalWdt


                call tableDiv(tTop,tLeft,tWdth)



                %>          
    <table border="0" width=100% cellpadding="0" cellspacing="0" bgcolor="#ffffff">
    <tr>
    <td align=right colspan="16" valign="top" height=20 bgcolor="#ffffFF" style="padding:10px 10px 10px 10px;">
    <b>Grandtotal p� ovenst�ende liste, fordelt p� aktivitets-typer:<br /><br /></b>
    <b>timer</b> ~ (enheder)<br />
	<% 
	
	
	
	
	gtotArr = split(replace(akttype_sel, "#", ""), ",")
	for j = 0 to ubound(gtotArr)
	
	call akttyper(gtotArr(j), 1)
	
	if len(trim(vlgt_typerTotalerGrand(gtotArr(j)))) <> 0 then
	antal = vlgt_typerTotalerGrand(gtotArr(j))
	enh = vlgt_typerTotalerGrandEnh(gtotArr(j))
	else
	antal = 0
	enh = 0
	end if
	
	if antal <> 0 then
	Response.write  "<span style='color:#999999;'>"& akttypenavn &": </span> "& formatnumber(antal, 2) 'akttypenavn & ":
	    
	    if cint(hideenheder) = 0 then
	    Response.Write " ~ ("& formatnumber(enh, 2) & ")"
	    end if
	 
	 
	 Response.Write "<br>"   
	 end if
	
	vlgt_typerTotalerGrand(gtotArr(j)) = 0
	vlgt_typerTotalerGrandEnh(gtotArr(j)) = 0
	
	grgrTotal = grgrTotal + antal 
	
	next %>
	
	<br /><br />
	Grandtotal: <b><%=formatnumber(grgrTotal, 2) %> </b><br />
	============================
	</td>
	</tr>
	</table>

    </div>
    <!-- slut table DIV -->	
	
	<%
	end sub
	
	
	
	
	public lastfakdato, jobans1, jobans1txt, jobstatus, usejoborakt_tp
	public jobstartdato, jobslutdato, jobForkalkulerettimer, rekvnr
	public jobans2txt, fasttimepris, fastpris, jobRisiko
	', fid, faknr
	
	
	function jobansoglastfak()
	                    
	                    '*** Jobansvarlige ***'
						jobans1 = 0
						fasttimepris = 0
						fastpris = 0
						usejoborakt_tp = 0
						strSQL2 = "SELECT mnavn, mnr, mid, job.id AS jid, jobans1, "_
						&" jobstatus, jobstartdato, jobslutdato, budgettimer, ikkebudgettimer, "_
						&" rekvnr, fastpris, jobtpris, usejoborakt_tp, job.risiko "_
						&" FROM job "_
						&" LEFT JOIN medarbejdere ON (mid = jobans1) "_
						&" WHERE job.jobnr = '"& oRec("Tjobnr") & "'" 
						
						'Response.Write strSQL2
						'Response.flush
						
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						
						jobid = oRec2("jid") ''* Bruges til at finde sidste faktura
						jobans1txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						jobstatus = oRec2("jobstatus")
						jobstartdato = oRec2("jobstartdato")
						
						'Response.Write oRec2("jobslutdato")
						'Response.flush
						
						jobslutdato = oRec2("jobslutdato")
						jobForkalkulerettimer = (oRec2("budgettimer") + oRec2("ikkebudgettimer"))
						rekvnr = oRec2("rekvnr")
						jobans1 = oRec2("mid")
						
						if jobForkalkulerettimer <> 0 then
						jobForkalkulerettimer = jobForkalkulerettimer
						else
						jobForkalkulerettimer = 1
						end if
						
						usejoborakt_tp = oRec2("usejoborakt_tp")
						fasttimepris = (oRec2("jobtpris")/jobForkalkulerettimer)
						fastpris = oRec2("fastpris")    
						
                        jobRisiko = oRec2("risiko")

						end if
						oRec2.close
						
						
						
						
						'*** jobANS 2 **'
						jobans2 = 0
						strSQL2 = "SELECT mnavn, mnr, mid, jobans2 FROM job, medarbejdere WHERE job.jobnr = '"& oRec("Tjobnr") &"' AND mid = jobans2"
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











 public vlgtMd, vlgtMdLastDay, strExportOskriftDage, ekspTxt 

    sub maanedsoversigtArr



       if media <> "export" then%>

               <br /><br />
                  <table border="0" width=98% cellpadding="2" cellspacing="0" bgcolor="#ffffff">
                      <%

                        end if



                        ekspTxt = ""
                        lastMid = 0
                        lastKid = 0
                        lastJnr = 0

                        'response.write "her"
                        'response.end

                           for m = 0 TO x - 1

                   
                           
    

                                  if lastMid <> medarbtimerArr(m,1) then
        
        
                                   
                          
                                                        if m <> 0 then
                                                    
                                                             if media <> "export" then 

                                                             call medarbMdtot()

                                                             end if

                                                 
                                                            select case lto 
                                                            case "mmmi", "intranet - local"              


                                                                if print = "j" then
                                                                %><tr><td colspan="31">
                                                                    <%
                                                                    call underskrift
                                                
                                                                %></td></tr><%
                                                                end if

                                                 

                                                 
                                                                if media = "export" then

                                                                   call underskriftExcel  

                                                                  end if

                                                             end select

                                                
                                                        end if 'm = 0 



                                                     if media <> "export" then 


                                          
                                                    if m = 0 then
                                                         %>
                                                        <tr><td colspan="10" style="padding-top:50px;"><p></p>
                                                        <%else%>
                                                     <tr><td colspan="10"><p style="page-break-before: always; padding-top:50px;"></p>
                                                    <%     
                                                    end if%>

                                       
                                                  <b style="font-size:16px; line-height:18px;">
                                                        <%if len(trim(medarbtimerArr(m,12))) <> 0 then %>
                                                        <%=medarbtimerArr(m,0) & " ["& medarbtimerArr(m,12) &"]" %>
                                                        <%else %>
                                                        <%=medarbtimerArr(m,0)%>
                                                        <%end if  %>
                                                     </b>

                                                          <%if cint(showcpr) = 1 then%>
                               
                                
                                                         <%if len(trim(medarbtimerArr(m,13))) <> 0 then%>
                                                            <span style="font-size:11px;">CPR: <%=medarbtimerArr(m,13)%></span> 
                                                          <%end if %>
                                                         <%end if %>
                                                         </td>
                                     
                                       
                                                    <td colspan="30" align="right" style="padding-right:40px; padding-top:70px;"> <%=monthname(strMrd) &" - "& strAar %></td>
                                                </tr>
                                                <tr bgcolor="#cccccc">
                                                    <td style="width:150px;">Projekt / Aktivitet</td>
                                                    <td style="width:150px;">Kommentar</td>
                                
                                                    <% end if'excel

                                         
                                        vlgtMd = strDag &"/"& strMrd &"/"& strAar   
                                        vlgtMdNext = dateadd("m", 1, vlgtMd)
                                        vlgtMdLastDay = dateadd("d", -1, vlgtMdNext)
                                        vlgtMdLastDay = day(vlgtMdLastDay)     
                                       
                                        for d = 1 to vlgtMdLastDay 

                                    
                                            if media <> "export" then 
                                           %>
                              
                                            <td style="width:30px; border-left:1px #999999 solid;" align="center"><%=d %></td>
                                            <%else 
                                        
                                                strExportOskriftDage = strExportOskriftDage & d & ";"
                                        
                                            end if %>

                                        <%next 
                                    
                                    
                                             if media <> "export" then%>
                                            </tr>

                                    <%      end if
                             end if 'last mid

                             if media <> "export" then

                                 if lastKid <> medarbtimerArr(m,10) OR lastJnr <> medarbtimerArr(m,5) then

                                        if lastJnr <> "0" then

                                         if media <> "export" then 
                                            call jobMdtot()
                                         end if

                                        end if

                                 %>
                                <tr>
                                    <td colspan="35" style="border-bottom:1px #cccccc solid;border-top:1px #cccccc solid;"><br /><%=medarbtimerArr(m,2) & " ("& medarbtimerArr(m,3)  &")" %><br />
                                        <b><%=medarbtimerArr(m,4) &" ("& medarbtimerArr(m,5) &")"%></b>
                                    </td>
                                     
                                </tr>

                                <%end if

                            end if


                                if media <> "export" then

                                select case right(m, 1)
                                case 0,2,4,6,8
                                bgfCol = "#EFf3ff"
                                case else
                                bgfCol = "#ffffff"
                                end select

                               

                                %>
                                <tr style="background-color:<%=bgfCol%>;">
                                    <td style="vertical-align:top;"><%=left(medarbtimerArr(m,6), 25)%></td>
                                    <td style="color:#999999; font-size:10px; line-height:11px; width:175px; vertical-align:top;">
                                            <%if len(trim(medarbtimerArr(m,11))) <> 0 then %>
                                            <i><%=left(medarbtimerArr(m,11), 100) %></i>
                                            <%end if %>
                                    </td>

                                                <%
                                end if 'exp


                                          if media = "export" then
                                          ekspTxt = ekspTxt & Chr(34) & medarbtimerArr(m,0) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,12) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,13) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,2) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,3) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,4) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,5) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,6) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,11) & Chr(34) &";"
                                          end if

                                                   for d = 1 to vlgtMdLastDay 
                                                    
                                                    vlgtMdCntDays = dateAdd("d", d-1, vlgtMd)
                                                    
                                                    if cDate(medarbtimerArr(m,8)) = cDate(vlgtMdCntDays) then
                                                    timerTxt = formatnumber(medarbtimerArr(m,9), 2)
                                                    timerTxtExp = timerTxt
                                                    timerTotprDag(d) = timerTotprDag(d) + medarbtimerArr(m,9)
                                                    timerTotprDagJob(d) = timerTotprDagJob(d) + medarbtimerArr(m,9)
                                                    else
                                                    timerTxt = "&nbsp;"
                                                    timerTxtExp = ""
                                                    timerTotprDag(d) = timerTotprDag(d) + 0
                                                    timerTotprDagJob(d) = timerTotprDagJob(d) + 0
                                                    end if 

                                               
                                                     if media <> "export" then
                                                    %>
                              
                                                        <td style="width:30px; border-left:1px #cccccc solid;" align="right"><%=timerTxt %></td>

                                                    <%
                                                    else
                                                    ekspTxt = ekspTxt & ""& Chr(34) & timerTxtExp & Chr(34) &";"
                                                    end if
                                                     
                                                    next 
                                                        
                                        if media <> "export" then%>
                                        </tr>
                                        <%
                                        else  
                                            ekspTxt = ekspTxt & "xx99123sy#z"  
                                        end if

                            lastMid = medarbtimerArr(m,1)
                            lastKid = medarbtimerArr(m,10)
                            lastJnr = medarbtimerArr(m,5)
                            next

                          
                                            
                                            
                        if media <> "export" then
                                        
                                            call jobMdtot()

                                            call medarbMdtot()

                                            if cint(antalM) > 1 then
                                            call medarbMdtotGT()
                                            end if
        

                        %>
                        </table>
                   
                                    <br /><br /><br />&nbsp;
                        <%end if 'media=exp


    




    end sub
	%>	
