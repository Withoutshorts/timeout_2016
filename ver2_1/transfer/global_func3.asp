
<%





public strEksport, strEksportTxt, strEksportEmailTxt
	function medarbafstem(intMid, startDato, slutDato, visning, akttype_sel)
	
	'Response.Write "akttype_sel" & akttype_sel
	
	select case visning 
	case 1, 2, 3, 5, 7 
	call akttyper2009(2) '7
	visTotalTimerAlltyp = 1
	case else
	visTotalTimerAlltyp = 0
	end select
	
	
	'**Viser kun tom igår ***'
	'slutdato = dateadd("d", -1, slutdato)
	
	
	startDato = year(startdato)&"/"&month(startdato)&"/"&day(startdato)
	slutDato = year(slutdato)&"/"&month(slutdato)&"/"&day(slutdato)
	
	stdato = day(startdato)&"/"&month(startdato)&"/"&year(startdato)
	sldato = day(slutdato)&"/"&month(slutdato)&"/"&year(slutdato)
	dageDiff = dateDiff("d", stdato, sldato, 2, 2)
	weekDiff = dateDiff("ww", stdato, sldato, 2, 2)
	
	if len(weekDiff) <> 0 AND weekDiff <> 0 then
	weekDiff = cint(weekDiff)
	else
	weekDiff = 1
	end if
	
	
	dim normTimerDag
	redim normTimerDag(100)
	
	if visTotalTimerAlltyp = 1 then
	  
	if cint(intMid) <> 0 then
	medarbSQL = " mid = " & intMid & ""
	else
	medarbSQL = " mid <> 0 "
	end if
	
	x = 0
	dim realTimer, medarbNavn, medarbNr, medarbInit, realIfTimer
	redim realTimer(100), realIfTimer(100), medarbNavn(100), medarbNr(100), medarbInit(100)
	
	dim normTimerUge, normTimer
	redim normTimerUge(100), normTimer(100)
	
	 dim fakTimer, mfForbrug , resTimer
	 dim medarbEmail, fradragTimer, medarbId, realfTimer
	 
	 redim resTimer(100), fakTimer(100), mfForbrug(100)
	 redim medarbEmail(100), fradragTimer(100), medarbId(100), realfTimer(100)

	 redim strEksport(100)
	
	'*** Alle typer Arr ***'
	'dim akttype_sel_arr
	'redim akttype_sel_arr(100)
    'akttype_sel_arr = split(replace(aktiveTyper, "#", ""), ",") 'akttype_sel
	
	strEksportPer = "Periode afgrænsning: "& formatdatetime(stdato, 1) & " - "&  formatdatetime(sldato, 1) & "xx99123sy#z"
	strEksportTxt = strEksportTxt & strEksportPer
	'strEksportEmailTxt = strEksportEmailTxt & strEksportPer
	    
	'** Timer **'
	strSQL = "SELECT t.tid, sum(t.timer) AS realtimer, m.mnavn, m.mnr, m.init, m.email, "_
	&" t.tdato, m.mid, m.medarbejdertype, tfaktim FROM medarbejdere m "_
	&" LEFT JOIN timer t ON (t.tmnr = m.mid AND ("& aty_sql_realhours &") "_
	&" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"') "_
	&" WHERE "& medarbSQL &" AND mansat <> '2' AND mansat <> '3' "_
	&" GROUP BY m.mid ORDER BY m.mnavn"
	
	'Response.Write strSQL &"<br>"
	'Response.flush
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	ntimPer  = 0
	'Response.Write oRec("tid") & " - " & oRec("tdato") & " - " & oRec("realtimer") & " - " & oRec("mnavn") & "<br>"
	    
	    medarbEmail(x) = oRec("email")
	    medarbNavn(x) = oRec("mnavn")
	    medarbNr(x) = oRec("mnr")
	    medarbInit(x) = oRec("init")
	    medarbId(x) = oRec("mid")
	    
	     realTimer(x) = oRec("realtimer")
	     if len(trim(realTimer(x))) <> 0 then
	     realTimer(x) = realTimer(x)
	     else
	     realTimer(x) = 0
	     end if
	    
	    
	    '*** Normeret periode (skal altid med til dags beregning under ferie, fefri, og afspad)  ***'
	    ntimPer = 0
	    call normtimerPer(oRec("mid"), stDato, dageDiff)
	    
	     normTimer(x) = ntimPer 
	    
	     if len(trim(normTimer(x))) <> 0 then
	     normTimer(x) = normTimer(x)
	     else
	     normTimer(x) = 0
	     end if
	    
	    
	    '*** Normeret uge til (til bl.a ferie beregning + gns. fuld dag) ***'
	    '** Standard normtimer / dag uanset helligdage. 7,4 (37 tim.) skal altid give 1 hel dag.
	    call nortimerStandardDag(oRec("medarbejdertype"))
	    normTimerDag(x) = formatnumber(normtimerStDag) 'formatnumber(ntimPer/(5*weekDiff),1)
	    
	    if normTimerDag(x) <> 0 then
	    normTimerDag(x) = normTimerDag(x)
	    showNormTimerdag = normTimerDag(x)
	    else
	    normTimerDag(x) = 0
	    showNormTimerdag = 0
	    end if
	    
	    
	    normTimerUge(x) = ntimPer/weekDiff
	    
	   
	   
	    if normTimerUge(x) <> 0 then
	    normTimerUge(x) = normTimerUge(x)
	    else
	    normTimerUge(x) = 0
	    end if
	    
	    
	    select case visning
	    case 1, 5, 7
	    call fordelpaaaktType(oRec("mid"), startDato, slutDato, visning, aktiveTyper, x)
	    end select
	    
	    
	    
	  
	   
	   '*********************************'
	   '*** Kalder sum på samle-typer ***'
	   '*********************************'
	    if visning = 1 OR visning = 2 OR visning = 3 OR visning = 5 OR visning = 7 then
	    
	    
	     '*** Fradrag i løntimer ***'
	    
	    strSQLfradrag = "SELECT t.tid, sum(t.timer) AS fratimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& oRec("mid") &" AND ("& aty_sql_frawhours &")"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	   'Response.Write strSQLfradrag & "<br><br>"
	   'Response.flush
	   
	    oRec2.open strSQLfradrag, oConn, 3
	    while not oRec2.EOF
	    fradragTimer(x) = oRec2("fratimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    
	    strSQLtillag = "SELECT t.tid, sum(t.timer) AS tiltimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& oRec("mid") &" AND ("& aty_sql_tilwhours &")"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	   'Response.Write strSQLtillag
	   'Response.flush
	    
	    tilTimer = 0
	    oRec2.open strSQLtillag, oConn, 3
	    while not oRec2.EOF
	    tilTimer = oRec2("tiltimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    '*** - minus (omvendt fortegn) tillæg da angivet timer i tilæg skal gøren fradrags tallet mindre
	     if len(trim(fradragTimer(x))) <> 0 then
         fradragTimer(x) = tilTimer - (fradragTimer(x))  
         else
         fradragTimer(x) = tilTimer
         end if
	    
	    
	    
	    end if
	    
	    
	    
	    
	    
	    if visning = 1 OR visning = 2 OR visning = 3 OR visning = 7 then
	 
	  
	      
	    '***  Fakturerbare TOTAL ***'
	    strSQLf = "SELECT t.tid, sum(t.timer) AS faktimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& oRec("mid") &" AND ("& aty_sql_realHoursFakbar &")"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	   'Response.Write strSQLf
	   'Response.flush
	   
	    oRec2.open strSQLf, oConn, 3
	    while not oRec2.EOF
	    realfTimer(x) = oRec2("faktimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	   
	    if len(trim(realfTimer(x))) <> 0 then
         realfTimer(x) = realfTimer(x)
         else
         realfTimer(x) = 0
         end if
	    
	    
	    end if
	    
	    if visning = 1 then
	   
	    '*** Interne Timer / ikke fakturerbare TOTAL ***'
	    strSQLif = "SELECT t.tid, sum(t.timer) AS ikkefaktimer,"_
	    &" t.tdato FROM timer t WHERE t.tmnr = "& oRec("mid") &" AND ("& aty_sql_realHoursIkFakbar &")"_
	    &" AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY t.tmnr "
	   
	   'Response.Write strSQLif
	   'Response.flush
	   
	    oRec2.open strSQLif, oConn, 3
	    while not oRec2.EOF
	    realIfTimer(x) = oRec2("ikkefaktimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	   
	    if len(trim(realIfTimer(x))) <> 0 then
         realIfTimer(x) = realIfTimer(x)
         else
         realIfTimer(x) = 0
         end if
	    
	    
	
	    
	    '*** Ressource Timer ***'
	    strSQLres = "SELECT sum(r.timer) AS restimer FROM ressourcer_md r WHERE r.medid = "& oRec("mid") &" AND "_
	    &" (aar >= YEAR('"& startdato &"') AND md >= MONTH('"& startdato &"') "_
	    &" AND aar <= YEAR('"& slutdato &"') AND md <= MONTH ('"& slutdato &"'))"
	    
	    'Response.Write strSQLres & "<br>"
	    
	    oRec2.open strSQLres, oConn, 3
	    while not oRec2.EOF
	    resTimer(x) = oRec2("restimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	 
	    
	    
	   '*** Faktureret ***'
	    strSQLfak = "SELECT sum(fms.fak) AS faktimer, f.fid, f.fakdato FROM fakturaer f "_
	    &" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid AND fms.mid = "& oRec("mid")&")"_
	    &" WHERE f.fakdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY fms.mid"
	    
	    'Response.Write strSQLfak & "<br>"
	    'Response.Flush
	    
	    oRec2.open strSQLfak, oConn, 3
	    while not oRec2.EOF
	    fakTimer(x) = oRec2("faktimer")
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	  
	    '*** Materiale forbrug ***'
	    strSQLmf = "SELECT sum(mf.matantal * (mf.matsalgspris * (kurs/100))) AS mfforbrug FROM materiale_forbrug mf "_
	    &" WHERE mf.forbrugsdato BETWEEN '"& startdato &"' AND '"& slutdato &"' AND usrid = "& oRec("mid") &" GROUP BY mf.usrid"
	    
	    'Response.Write strSQLmf & "<br>"
	    'Response.Flush
	    
	    oRec2.open strSQLmf, oConn, 3
	    while not oRec2.EOF
	    mfForbrug(x) = oRec2("mfforbrug")/1
	    oRec2.movenext
	    wend
	    oRec2.close
	    
	    
	     if len(trim(resTimer(x))) <> 0 then
	     resTimer(x) = resTimer(x)
	     else
	     resTimer(x) = 0
	     end if
    	 
	     if len(trim(fakTimer(x))) <> 0 then
	     fakTimer(x) = fakTimer(x)
	     else
	     fakTimer(x) = 0
	     end if
	     
	     if len(trim(mfForbrug(x))) <> 0 then
	     mfForbrug(x) = mfForbrug(x)
	     else
	     mfForbrug(x) = 0
	     end if
	    
	   end if
	   
	    
	x = x + 1
	oRec.movenext
    wend
	oRec.close
	
	
	
	
	
	else 'visTotalTimerAlltyp = 1
	
	
	    '*** Normeret uge til (til bl.a ferie beregning + gns. fuld dag) ***'
	    '** Standard normtimer / dag uanset helligdage. 7,4 (37 tim.) skal altid give 1 hel dag.
	    strSQLmt = " SELECT medarbejdertype FROM medarbejdere WHERE mid = " & intMid
	    oRec.open strSQLmt, oConn, 3
	    mType = 0
	    if not oRec.EOF then
	    
	    mType = oRec("medarbejdertype")
	    
	    end if
	    oRec.close
	    
	    call nortimerStandardDag(mType)
	    normTimerDag(x) = formatnumber(normtimerStDag) 'formatnumber(ntimPer/(5*weekDiff),1)
	
	 
	 call fordelpaaaktType(intMid, startDato, slutDato, visning, akttype_sel, x)
	    
	
	end if
	
	
	
	
	
	
	'**** Præsentation ***'
	
	
	
	select case visning 
	case 1
		
	if instr(akttype_sel, "#-1#") <> 0 then
	cops = 8
	else
	cops = 0
	end if
	
	
	%>
	 <table cellspacing=1 cellpadding=2 border=0 bgcolor="#d6dff5"><!-- #5C75AA -->
	 <tr bgcolor="#5582d2">
	 <td class=alt><b><%=tsa_txt_147%></b></td>
	 
	  <%if instr(akttype_sel, "#-5#") <> 0 AND stempelurOn = 1 then %>
	 <td class=alt colspan=5><b>Løn timer</b><br />Stempelur</td>
	 <%end if
	 
	 
	 if instr(akttype_sel, "#-1#") <> 0 then %>
	 <td class=alt colspan=<%=cops %>><b><%=tsa_txt_148%></b> <br />
	 <%=tsa_txt_149%></td>
	 <%end if %>
	 
	 
	 
	 
	 <%if instr(akttype_sel, "#1#") <> 0 then %>
	 <td class=alt>&nbsp;</td>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#2#") <> 0 then %>
	<td class=alt>&nbsp;</td>
	<%end if %>
	
	 <%if instr(akttype_sel, "#6#") <> 0 then %>
	<td class=alt>&nbsp;</td>
	<%end if %>
	 
	   <%if instr(akttype_sel, "#50#") <> 0 then %>
	<td class=alt>&nbsp;</td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#51#") <> 0 then %>
	<td class=alt>&nbsp;</td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#52#") <> 0 then %>
	<td class=alt>&nbsp;</td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#53#") <> 0 then %>
	<td class=alt>&nbsp;</td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#54#") <> 0 then %>
	<td class=alt>&nbsp;</td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#55#") <> 0 then %>
	<td class=alt>&nbsp;</td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#60#") <> 0 then %>
	 <td class=alt>&nbsp;</td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#90#") <> 0 then %>
	<td class=alt>&nbsp;</td>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#91#") <> 0 then %>
	 <td class=alt>&nbsp;</td>
	 <%end if %>
	 
	 
	 
	 
	 
	 
	 
	   <%if instr(akttype_sel, "#7#") <> 0 then %>
	<td class=alt><b>Flex</b></td>
	 <%end if %>
	
	 
	 <%if instr(akttype_sel, "#5#") <> 0 then %>
	 <td class=alt><b><%=tsa_txt_265%></b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#10#") <> 0 then %>
	 <td class=alt><b><%=tsa_txt_150%></b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#9#") <> 0 then %>
	 <td class=alt><b>Pause</b></td>
	 <%end if %>
	 
	 
	 <%
	 copsff = 1
	 if instr(akttype_sel, "#12#") <> 0 then 
	 copsff = copsff + 1
	 end if
	 
	 if instr(akttype_sel, "#18#") <> 0 then 
	 copsff = copsff + 1
	 end if
	 
	 if instr(akttype_sel, "#13#") <> 0 then 
	 copsff = copsff + 1
	 end if
	 
	 if instr(akttype_sel, "#17#") <> 0 then 
	 copsff = copsff + 1
	 end if
	 
	 if instr(akttype_sel, "#12#") <> 0 OR instr(akttype_sel, "#18#") <> 0 OR instr(akttype_sel, "#13#") <> 0 OR instr(akttype_sel, "#17#") <> 0 then %>
	<td class=alt colspan=<%=copsff %>><b><%=tsa_txt_174%></b><br />
	(Sæt periode = ferieår for at se total)</td>
	<%end if %>
	 
	 
	 
	  <%
	 copsFe = 1
	 if instr(akttype_sel, "#11#") <> 0 then 
	 copsFe = copsFe + 1
	 end if
	 
	 if instr(akttype_sel, "#14#") <> 0 then 
	 copsFe = copsFe + 1
	 end if
	 
	 if instr(akttype_sel, "#19#") <> 0 then 
	 copsFe = copsFe + 1
	 end if
	 
	 if instr(akttype_sel, "#15#") <> 0 then 
	 copsFe = copsFe + 1
	 end if
	 
	 if instr(akttype_sel, "#16#") <> 0 then 
	 copsFe = copsFe + 1
	 end if
	 
	 if instr(akttype_sel, "#11#") <> 0 OR instr(akttype_sel, "#14#") <> 0 OR instr(akttype_sel, "#19#") <> 0 OR _
	 instr(akttype_sel, "#15#") <> 0 OR instr(akttype_sel, "#16#") <> 0 then %>
	<td class=alt colspan=<%=copsFe %>><b><%=tsa_txt_152%></b><br />
	(Sæt periode = ferieår for at se total)</td>
	<%end if %>
	  
	 
	 
	 <%
	 copsAf = 1
	 if instr(akttype_sel, "#30#") <> 0 then 
	 copsAf = copsAf + 1
	 end if
	 
	 if instr(akttype_sel, "#31#") <> 0 then 
	 copsAf = copsAf + 1
	 end if
	 
	 if instr(akttype_sel, "#32#") <> 0 then 
	 copsAf = copsAf + 1
	 end if
	 
	 if instr(akttype_sel, "#33#") <> 0 then 
	 copsAf = copsAf + 1
	 end if
	 
	 if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 OR _
	 instr(akttype_sel, "#32#") <> 0 OR instr(akttype_sel, "#33#") <> 0 then %>
	<td class=alt colspan=<%=copsAf %>><b><%=tsa_txt_151 %></b></td>
	<%end if %>
	 
	 
	 <%if instr(akttype_sel, "#20#") <> 0 AND instr(akttype_sel, "#21#") <> 0 then 
	 copsS = 2
	 else
	 copsS = 1
	 end if %>
	 
	  <%if instr(akttype_sel, "#20#") <> 0 OR instr(akttype_sel, "#21#") <> 0 then %>
	 <td class=alt colspan=<%=copsS%>><b><%=tsa_txt_153%></b></td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#8#") <> 0 then %>
	 <td class=alt><b>Sundh.</b></td>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#81#") <> 0 then %>
	 <td class=alt><b>Læge</b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-2#") <> 0 then %>
	 <td class=alt><b><%=tsa_txt_154%></b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-3#") <> 0 then %>
	 <td class=alt><b><%=tsa_txt_155%></b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#61#") <> 0 then %>
	 <td class=alt><b>Stk.</b></td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-4#") <> 0 then %>
	  <td class=alt><b><%=tsa_txt_156%></b></td>
	  <%end if %>
	 </tr>
	 
	 
	 
	 
	 <!-- Under overskrrifter --> <!--- Akt type navn -->
	 
	 <tr bgcolor="#8CAAe6">
	 
	 <td style="border-right:1px #cccccc solid;" class=alt>&nbsp;</td>
	 
	 <%strEksport(x) = strEksport(x) & tsa_txt_147 & ";" %>
	 
	  <%if instr(akttype_sel, "#-5#") <> 0 AND stempelurOn = 1 then %>
	 <td class=alt_lille align=right valign=bottom><%=tsa_txt_148 %>:<%=tsa_txt_141 %></td>
	 <td class=alt_lille align=right valign=bottom>Till./Frad. +/-</td>
	 <td class=alt_lille align=right valign=bottom bgcolor="#cccccc">Sum</td>
	 <td class=lille align=right valign=bottom bgcolor="#EFF3FF">Bal Lønt. / Real</td>
	 <td class=alt_lille align=right valign=bottom bgcolor="#9acd32">Bal Lønt. / Norm</td>
	 
	 <%strEksport(x) = strEksport(x) & tsa_txt_148 &":"& tsa_txt_141 &";Till./Frad. +/-;Sum;Bal Lønt./Real;Bal Lønt./Norm;" %>
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-1#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom bgcolor="#cccccc" style=" "><%=tsa_txt_157%></td>
	 <td class=alt_lille align=right valign=bottom bgcolor="#cccccc"><%=tsa_txt_172 &" "& tsa_txt_179%></td>
	 <td class=alt_lille align=right valign=bottom bgcolor="#cccccc"><%=tsa_txt_158%></td>
	 <td class=alt_lille align=right valign=bottom bgcolor="#cccccc"><%=tsa_txt_259%></td>
	 <td class=alt_lille align=right valign=bottom bgcolor="pink"><%=tsa_txt_159%></td>
	 <td class=alt_lille align=right valign=bottom bgcolor="#cccccc"><%=tsa_txt_160%></td>
	 <td class=alt_lille align=right valign=bottom bgcolor="#cccccc"><%=tsa_txt_161%></td>
	 <td class=alt_lille align=right valign=bottom bgcolor="#cccccc" style=""><%=tsa_txt_163%></td>
	 
	 <%strEksport(x) = strEksport(x) & tsa_txt_157 & ";"& tsa_txt_172 &" "& tsa_txt_179 &";" & tsa_txt_158 & ";"& tsa_txt_259 & ";" & tsa_txt_159 & ";"& tsa_txt_160 & ";" & tsa_txt_161 & ";" & tsa_txt_163 & ";"  %>
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#1#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(1, 1) %>
	 <%=akttypenavn%>
	  <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#2#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(2, 1) %>
	 <%=akttypenavn%>
	 </td>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	<%end if %>
	
	<%if instr(akttype_sel, "#50#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(50, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#51#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(51, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#52#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(52, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#53#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(53, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#54#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(54, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#55#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(55, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#60#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(60, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#90#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(90, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#91#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(91, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#6#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(6, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#7#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(7, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#5#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom style="">
	 <%call akttyper(5, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#10#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(10, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#9#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(9, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 
	 <!--Ferie Fridage -->
	 <%if instr(akttype_sel, "#12#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(12, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	<!-- <td class=alt_lille align=right valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksport(x) = strEksport(x) & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#18#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(18, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %> <br> >> dd.</td>
	<!-- <td class=alt_lille align=right valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksport(x) = strEksport(x) & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#13#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(13, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	<!-- <td class=alt_lille align=right valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksport(x) = strEksport(x) & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#17#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(17, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille align=right valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksport(x) = strEksport(x) & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 
	 
	 
	 <%if instr(akttype_sel, "#12#") <> 0 OR instr(akttype_sel, "#18#") <> 0 OR instr(akttype_sel, "#13#") <> 0 OR instr(akttype_sel, "#17#") <> 0 then %>
	 <td bgcolor="#FFC0CB" class=alt_lille align=right valign=bottom style=""><%=tsa_txt_282 &" "& tsa_txt_280 &" "& tsa_txt_275%></td>
	 <!--<td bgcolor="#FFC0CB" class=alt_lille align=right valign=bottom style=""><%=tsa_txt_282 &" "& tsa_txt_280 &" "& tsa_txt_276%></td>-->
	 <%strEksport(x) = strEksport(x) & tsa_txt_282 &" "& tsa_txt_280 &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	
	 
	 <!-- Ferie -->
	 <%if instr(akttype_sel, "#15#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(15, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille align=right valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksport(x) = strEksport(x) & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#11#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(11, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %> <br> >> dd.</td>
	 <!--<td class=alt_lille align=right valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksport(x) = strEksport(x) & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#14#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(14, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille align=right valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksport(x) = strEksport(x) & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#19#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(19, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille align=right valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksport(x) = strEksport(x) & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#16#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(16, 1) %>
	 <%=akttypenavn &" "& tsa_txt_275 %></td>
	 <!--<td class=alt_lille align=right valign=bottom><%=akttypenavn &" "& tsa_txt_276 %></td>-->
	 <%strEksport(x) = strEksport(x) & akttypenavn &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 <% if instr(akttype_sel, "#11#") <> 0 OR instr(akttype_sel, "#14#") <> 0 OR instr(akttype_sel, "#19#") <> 0 OR _
	 instr(akttype_sel, "#15#") <> 0 OR instr(akttype_sel, "#16#") <> 0 then %>
	  <td bgcolor="#FFC0CB" class=alt_lille align=right valign=bottom style=""><%=tsa_txt_281 &" "& tsa_txt_280 &" "& tsa_txt_275%></td>
	 <!--<td bgcolor="#FFC0CB" class=alt_lille align=right valign=bottom style=""><%=tsa_txt_281 &" "& tsa_txt_280 &" "& tsa_txt_276%></td>-->
	 <%strEksport(x) = strEksport(x) &" "& tsa_txt_281 &" "& tsa_txt_280 &" "& tsa_txt_275 &";"%>
	 <%end if %>
	 
	 
	 <!-- Afspad --->
	 <%if instr(akttype_sel, "#30#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(30, 1) %>
	 <%=akttypenavn%><br />(Tim.) Enh.</td>
	 <%strEksport(x) = strEksport(x) & akttypenavn &" Tim. ; " & akttypenavn &" Enh. ;"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#31#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(31, 1) %>
	 <%=akttypenavn%>.</td>
	 <%strEksport(x) = strEksport(x) & akttypenavn &".;"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#32#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(32, 1) %>
	 <%=akttypenavn%>.</td>
	 <%strEksport(x) = strEksport(x) & akttypenavn &".;"%>
	 <%end if %>
	 
	<%if instr(akttype_sel, "#33#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(33, 1) %>
	 <%=akttypenavn%></td>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 OR instr(akttype_sel, "#32#") <> 0 OR instr(akttype_sel, "#33#") <> 0 then %>
	 <td bgcolor="#FFC0CB" class=alt_lille align=right valign=bottom style=""><%=tsa_txt_283 &" "& tsa_txt_280%></td>
	 <%strEksport(x) = strEksport(x) & tsa_txt_283 &" "& tsa_txt_280 &";"%>
	 <%end if %>
	 
	 
	
	 <!-- Sygdom -->
	  <%if instr(akttype_sel, "#20#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(20, 1) %>
	 <%=akttypenavn%>
	  <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#21#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(21, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#8#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(8, 1) %>
	 <%=akttypenavn%>
	  <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#81#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(81, 1) %>
	 <%=akttypenavn%>
	  <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	
	 <%if instr(akttype_sel, "#-2#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom><%=tsa_txt_154%></td>
	  <%strEksport(x) = strEksport(x) & tsa_txt_154 &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-3#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom><%=tsa_txt_155%>
	 <%strEksport(x) = strEksport(x) & tsa_txt_155 &";"%>
	 </td>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#61#") <> 0 then %>
	 <td class=alt_lille align=right valign=bottom>
	 <%call akttyper(61, 1) %>
	 <%=akttypenavn%>
	 <%strEksport(x) = strEksport(x) & akttypenavn &";"%>
	 </td>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-4#") <> 0 then %>
	  <td style="width:90px;" class=alt_lille align=right valign=bottom><%=tsa_txt_171%></td>
	  <%strEksport(x) = strEksport(x) & tsa_txt_171 &";"%>
	  <%end if %>
	 </tr>
	 
	 <%
	 
	 strEksportTxt = strEksportTxt & strEksport(x)
	 strEksportEmailTxtHeader = strEksport(x)
	 
	 for x = 0 to x - 1 
	 
	 select case right(x, 1)
	 case 0,2,4,6,8
	 bgcl = "#ffffff"
	 'bgcolorTr = "#cccccc"
	 case else
	 bgcl = "#EFF3ff" '"#D6DFf5"
	 'bgcolorTr = "#ffffff"
	 end select%>
	 
	 <%strEksport(x) = strEksport(x) &"xx99123sy#z"%>
	 
	 <tr bgcolor="<%=bgcl %>">
	 <td class=lille style=" white-space: nowrap;"><b><%=medarbNavn(x) %></b> (<%=medarbNr(x) %>) 
	 <%if len(trim(medarbInit(x))) <> 0 then %>
	  - <%=medarbInit(x) %>
	  <%end if %>
	  </td>
	 
	 <%strEksport(x) = strEksport(x) & medarbNavn(x) &" ("& medarbNr(x) &");"%>
	 
	  <%if instr(akttype_sel, "#-5#") <> 0 AND  stempelurOn = 1 then %>
	 <td align=right style=" white-space: nowrap;" class=lille>
	 <%
	 
	 ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	 call fLonTimerPer(ltimerStDato, dageDiff, 2, medarbId(x)) 
	 strEksport(x) = strEksport(x) & thoursTot &":"& left(tminTot, 2) &";"
	 %> 
	 
	</td>
	
	  <%
	 fraDtimer = formatnumber(fradragTimer(x),2)
	 %>
	  <td align=right style=" white-space: nowrap;" class=lille>
	  <%call timerogminutberegning(fraDtimer*60) %>
	  <%=thoursTot &":"& left(tminTot, 2) %> </td>
	  
	  <%strEksport(x) = strEksport(x) & thoursTot &":"& left(tminTot, 2) &";"%>
	
	
	
	<%
	lontimerTotMfradrag = ((totalTimerPer100/60) + (fradragTimer(x)))
	normtime_lontime = -((normTimer(x) - (lontimerTotMfradrag)) * 60)
	 %>
	  <td align=right style=" white-space: nowrap;" class=lille>
	  <%call timerogminutberegning(lontimerTotMfradrag*60) %>
	  <%=thoursTot &":"& left(tminTot, 2) %> </td>
	  
	  <%strEksport(x) = strEksport(x) & thoursTot &":"& left(tminTot, 2) &";"%>
	  
	  
	  <%
	 bal_lon_real = formatnumber(lontimerTotMfradrag,2) - formatnumber(realTimer(x),2)
	 %>
	  <td align=right style=" white-space: nowrap;" class=lille>
	  <%call timerogminutberegning(bal_lon_real*60) %>
	  <%=thoursTot &":"& left(tminTot, 2) %> </td>
	  
	  <%strEksport(x) = strEksport(x) & thoursTot &":"& left(tminTot, 2) &";"%>
	  
	  
	 <td align=right style=" white-space: nowrap; " class=lille>
	  <%call timerogminutberegning(normtime_lontime) %>
	  <%=thoursTot &":"& left(tminTot, 2) %></td>
	  
	  <%strEksport(x) = strEksport(x) & thoursTot &":"& left(tminTot, 2) &";"%>
	  
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#-1#") <> 0 then %>
	 <td align=right style=" white-space: nowrap;" class=lille><b><%=formatnumber(realTimer(x),2)%></b></td>
	 <td align=right style=" white-space: nowrap;" class=lille><%=formatnumber(realTimer(x)/weekDiff,2)%></td>
	 <td align=right style=" white-space: nowrap;" class=lille><%=formatnumber(normTimer(x),2)%></td>
	 
	 <%strEksport(x) = strEksport(x) & formatnumber(realTimer(x),2) &";" & formatnumber(realTimer(x)/weekDiff,2) & ";"& formatnumber(normTimer(x),2) &";"%>
	
	 
	 <td align=right class=lille style=" white-space: nowrap;">
	 <%if normTimerDag(x) > 0 then %>
	 <%=formatnumber(normTimerDag(x), 1) %>
	 <%strEksport(x) = strEksport(x) & formatnumber(normTimerDag(x), 1) & ";" %>
	 <%else%>
	 <%=formatnumber(0, 2) %>
	 <%strEksport(x) = strEksport(x) & formatnumber(0, 2)  & ";" %>
	 <%end if %>
	 
	  
	 </td>
	 
	 <td align=right class=lille style="  white-space: nowrap;"><b><%=formatnumber((realTimer(x) - normTimer(x)),2)%></b></td>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(realfTimer(x),2)%></td>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(realIfTimer(x),2)%></td>
	 
	  <%
	  if realfTimer(x) <> 0 then
	  iebal = (realIfTimer(x)/realfTimer(x)) * 100 
	  else
	  iebal = 0
	  end if%>
	  
	  <%strEksport(x) = strEksport(x) & formatnumber((realTimer(x) - normTimer(x)),2) &";"& formatnumber(realfTimer(x),2) &";"& formatnumber(realIfTimer(x),2) & ";"& formatnumber(iebal,2) & ";" %>
	  
	  <td align=right class=lille style="  white-space: nowrap;"><%=formatnumber(iebal,2)%>%</td>
	   <%end if %>
	 
	 <%if instr(akttype_sel, "#1#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fTimer(x),2) %></td>
	 <% strEksport(x) = strEksport(x) & formatnumber(fTimer(x),2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#2#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ifTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(ifTimer(x),2) & ";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#50#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(dagTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(natTimer(x),2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#51#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(natTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(natTimer(x),2) & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#52#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(weekendTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(weekendTimer(x),2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#53#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(weekendnatTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(weekendnatTimer(x),2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#54#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(aftenTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(aftenTimer(x),2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#55#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(aftenweekendTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(aftenweekendTimer(x),2) & ";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#60#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(adhocTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(adhocTimer(x),2) & ";"%>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#90#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(e1Timer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(e1Timer(x),2) & ";"%>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#91#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(e2Timer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(e2Timer(x),2) & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#6#") <> 0 then %>
	  <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(sTimer(x),2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(sTimer(x),2) & ";"%>
	 <%end if %>
	  
	  <%if instr(akttype_sel, "#7#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(flexTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(flexTimer(x),2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#5#") <> 0 then %>
	   <td align=right class=lille style=" white-space: nowrap; "><%=formatnumber(km(x),2) %></td>
	  <%strEksport(x) = strEksport(x) & formatnumber(km(x),2) & ";"%>
	  <%end if %>
	  
	  <%if instr(akttype_sel, "#10#") <> 0 then %>
	  <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fpTimer(x),2) %></td>
	  <%strEksport(x) = strEksport(x) & formatnumber(fpTimer(x),2) & ";"%>
	  <%end if %>
	  
	  <%if instr(akttype_sel, "#9#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(pausTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(pausTimer(x),2) & ";"%>
	 <%end if %>
	 
	 
	 
	  <!--Ferie Fridage -->
	 <%if instr(akttype_sel, "#12#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	
	 <%if normTimerDag(x) <> 0 then 
	 fefriVal = fefriTimer(x)/normTimerDag(x)
	 else
	 fefriVal = 0
	 end if
	 %>
	 <%=formatnumber(fefriVal,2) %>
	 
	 </td>
	 
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fefriTimer(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(fefriVal,2) & ";"%>
	 <%end if %>
	 
	 
	  <%if instr(akttype_sel, "#18#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	
	 <%if normTimerDag(x) <> 0 then 
	 fefriplVal = fefriplTimer(x)/normTimerDag(x)
	 else
	 fefriplVal = 0
	 end if
	 %>
	 <%=formatnumber(fefriplVal,2) %>
	 
	 </td>
	 
	 
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fefriTimer(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(fefriplVal,2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#13#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	  <%if normTimerDag(x) <> 0 then 
	 fefribrVal = fefriTimerBr(x)/normTimerDag(x)
	 else
	 fefribrVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(fefribrVal,2) %></td> 
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fefriTimerBr(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(fefribrVal,2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#17#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 fefriudbVal = fefriTimerUdb(x)/normTimerDag(x)
	 else
	 fefriudbVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(fefriudbVal,2) %></td> 
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fefriTimerUdb(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(fefriudbVal,2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#12#") <> 0 OR instr(akttype_sel, "#18#") <> 0 OR instr(akttype_sel, "#13#") <> 0 OR instr(akttype_sel, "#17#") <> 0 then 
	 fefriBal = 0
	 fefriBal  = (fefriTimer(x) - (fefriTimerBr(x) + fefriTimerUdb(x)))
	 
	 if normTimerDag(x) <> 0 then 
	 fefriBalVal = fefriBal/normTimerDag(x)
	 else
	 fefriBalVal = 0
	 end if
	 
	 %>
	 <td align=right class=lille style="  white-space: nowrap;">
	 <%=formatnumber(fefriBalVal,2) %></td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fefriBal,2)%></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(fefriBalVal,2) & ";"%>
	 <%end if %>
	  
	
	 <!-- Ferie -->
	 
	  <%if instr(akttype_sel, "#15#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieVal = ferieOptjtimer(x)/normTimerDag(x)
	 else
	 ferieVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(ferieVal,2) %></td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ferieOptjtimer(x),2) %></td>-->
	  <%strEksport(x) = strEksport(x) & formatnumber(ferieVal,2) & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#11#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 feriePlVal = feriePLTimer(x)/normTimerDag(x)
	 else
	 feriePlVal = 0
	 end if
	 %>
	 
	 
	 <%=formatnumber(feriePlVal,2) %></td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(feriePLTimer(x),2) %></td>-->
	  <%strEksport(x) = strEksport(x) & formatnumber(feriePlVal,2) & ";"%>
	 
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#14#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieAftVal = ferieAFTimer(x)/normTimerDag(x)
	 else
	 ferieAftVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(ferieAftVal,2) %> </td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ferieAFTimer(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(ferieAftVal,2) & ";"%>
	 <%end if %>
	 
	 
	 <%if instr(akttype_sel, "#19#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieAftulonVal = ferieAFulonTimer(x)/normTimerDag(x)
	 else
	 ferieAftulonVal = 0
	 end if
	 %>

	 <%=formatnumber(ferieAftulonVal,2) %> </td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ferieAFTimer(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(ferieAftulonVal,2) & ";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#16#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">
	 
	 <%if normTimerDag(x) <> 0 then 
	 ferieUdbVal = ferieUdbTimer(x)/normTimerDag(x)
	 else
	 ferieUdbVal = 0
	 end if
	 %>
	 
	 <%=formatnumber(ferieUdbVal,2) %></td>
	 <!--<td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ferieUdbTimer(x),2) %></td>-->
	 <%strEksport(x) = strEksport(x) & formatnumber(ferieUdbVal,2) & ";"%>
	 <%end if %>
	 
	 <% if instr(akttype_sel, "#11#") <> 0 OR instr(akttype_sel, "#14#") <> 0 OR instr(akttype_sel, "#19#") <> 0 OR _
	 instr(akttype_sel, "#15#") <> 0 OR instr(akttype_sel, "#16#") <> 0 then 
	 
	 ferieBal = 0
	 ferieBal  = (ferieOptjtimer(x) - (ferieAFTimer(x) + ferieAFulonTimer(x) + ferieUdbTimer(x)))
	 
	 if normTimerDag(x) <> 0 then 
	 ferieBalVal = ferieBal/normTimerDag(x)
	 else
	 ferieBalVal = 0
	 end if
	 
	 %>
	 <td align=right class=lille style=" white-space: nowrap; ">
	 <%=formatnumber(ferieBalVal,2) %></td>
	<!-- <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(ferieBal,2)%></td>-->
	<%strEksport(x) = strEksport(x) & formatnumber(ferieBalVal,2) & ";"%>
	<%end if %>
	 
	 
     <!-- Afspad --->
	 <%if instr(akttype_sel, "#30#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;">(<%= formatnumber(afspTimerTim(x), 2) %>) <%=formatnumber(afspTimer(x), 2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(afspTimerTim(x), 2) &";"& formatnumber(afspTimer(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#31#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(afspTimerBr(x), 2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(afspTimerBr(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#32#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(afspTimerUdb(x), 2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(afspTimerUdb(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#33#") <> 0 then 
	 afspadUdbBal = 0
	 afspadUdbBal = (afspTimerUdb(x) - afspTimerOUdb(x)) %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(afspadUdbBal, 2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(afspadUdbBal, 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 OR _
	 instr(akttype_sel, "#32#") <> 0 OR instr(akttype_sel, "#33#") <> 0 then %>
	 
	 <%AfspadBal = 0 
	 
	 AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	 %>
	 <td align=right class=lille style=" white-space: nowrap; "><%=formatnumber(AfspadBal, 2)%></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(AfspadBal, 2) &";"%>
	 <%end if %>
	 
	 <!-- Syg -->
	  
	  <%if instr(akttype_sel, "#20#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(sygTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(sygTimer(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#21#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(BarnsygTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(BarnsygTimer(x), 2) &";"%>
	 <%end if %>
	 
	  <%if instr(akttype_sel, "#8#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(sundTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(sundTimer(x), 2) &";"%>
	 <%end if %>
	 
	   <%if instr(akttype_sel, "#81#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(lageTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(lageTimer(x), 2) &";"%>
	 <%end if %>
	 
	 <!-- Ressource -->
	 <%if instr(akttype_sel, "#-2#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(resTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(resTimer(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-3#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(fakTimer(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(fakTimer(x), 2) &";"%>
	 <%end if %> 
	  
	  <%if instr(akttype_sel, "#61#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(stkAntal(x),2) %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(stkAntal(x), 2) &";"%>
	 <%end if %>
	 
	 <%if instr(akttype_sel, "#-4#") <> 0 then %>
	 <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(mfForbrug(x)) & " DKK" %></td>
	 <%strEksport(x) = strEksport(x) & formatnumber(mfForbrug(x), 2) &";"%>
	 <%end if %>
	 </tr>
	 
	 
	 
	 
	 <%
	 
	 strEksportTxt = strEksportTxt & strEksport(x)
	 strEksportEmailTxt = strEksportEmailTxtHeader & strEksport(x)
	 
	 if sendemail = "j" then
	 
	 
	 
	 
	        '***** Oprettter Mail object ***
	        '**** Skal kun sende en linie til hver medarbejder ***'
	        '**** De må ikke se hinandens ****'
			if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\bal_real_norm_2007.asp" then
					
			
			
                txtEmailHtml = replace(strEksportEmailTxt, "xx99123sy#z", "</tr><tr><td class=lille align=right valign=bottom>")
			    txtEmailHtml = replace(txtEmailHtml, ";", "</td><td class=lille align=right valign=bottom>")
			   
				
				Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			    ' Sætter Charsettet til ISO-8859-1 
			    Mailer.CharSet = 2
			    ' Afsenderens navn 
			    Mailer.FromName = "TimeOut Email Service"
			    ' Afsenderens e-mail 
			    Mailer.FromAddress = "timeout_no_reply@outzource.dk"
			    Mailer.RemoteHost = "webmail.abusiness.dk"
    			
			    ' Mailens emne
			    Mailer.Subject = "TimeOut Afstemningsrapport"
			    ' Modtagerens navn og e-mail
			    Mailer.AddRecipient medarbNavn(x), medarbEmail(x)
			    Mailer.ContentType = "text/html"
			    
			    'Mailer.AddAttachment "d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\data\"& file 
			    
					' Selve teksten
					Mailer.BodyText = "<html><head><title></title>"_
					&"<LINK rel=""stylesheet"" type=""text/css"" href=""https://outzource.dk/timeout_xp/wwwroot/ver2_1/inc/style/timeout_style_print.css""></head>"_
					&"<body><br>Afstemningsrapport<br>"&replace(strEksportPer, "xx99123sy#z", "<br>")&"<table width=100% cellpadding=1 cellspacing=0 border=1><tr><td class=lille valign=bottom>"_ 
					& txtEmailHtml & "</table><br><br>Med venlig hilsen<br>" & session("user") & "</body></html>"
					
					
				
					If Mailer.SendMail Then
    				
					Else
    					Response.Write "Fejl...<br>" & Mailer.Response
  					End if
					
			
			end if
		
		
		strEksportEmailTxt = ""
		
	 
	 end if 'send email
	 
	 
	 
	 next%>
	 </table>
	 <%
	 
	 case 2,3 '**** minivisning på timereg siden ***
	 
	 %>
	
	 
	 
	 <%if visning = 2 then%>
	     	    
	   <br /><br />
	   
	     <h4> Md -> Dato</h4> <br /><b><%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %> </b> (<%=tsa_txt_318 %>)
	    
	 <%else%>
	    
	    
	    <%=tsa_txt_173%> pr. dag: <%=formatnumber(showNormTimerdag) %> 
	    <br /><%=tsa_txt_273%>
	    <br />
	    <br /><br />
	    
	    <h4>Uge <%=datepart("ww", stdato, 2, 2) %> - <%=datepart("yyyy", stdato, 2, 2) %></h4>
	    <b><%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %></b> 
	    <%if datepart("ww", stdato, 2, 2) = datepart("ww", now, 2, 2) AND datepart("yyyy", stdato, 2, 2) = datepart("yyyy", now, 2, 2) then%>
	    (<%=tsa_txt_318 %>)
	    <%end if %>
	 <%end if%>
	 
	 
	 
	 <%
	 x = 0
	 
	 %>
	 
	
	
	 <table cellspacing=0 cellpadding=2 border=0 width=100%>
	 <tr bgcolor="#EFF3FF">
	 <td class=lille style="border-bottom:1px silver dashed;">&nbsp;</td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_172%></b></td>
	 <%if lto <> "cst" then %>
	 <td class=lille bgcolor="#cccccc" style="border-bottom:1px silver dashed;">(heraf<br />fak.bare)</td>
	 <%end if %>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_173%></b></td>
	 
	 <%if session("stempelur") <> 0 then %> 
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Løn timer</b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Till. / <br />Frad. +/-</b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Sum</b></td>
	 <%end if %>
	
	  <%if lto <> "cst" then %>
	 <td bgcolor="pink" align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b> (Real / Norm)</td>
	 <%end if %>
	 
	 <%if session("stempelur") <> 0 then %> 
	 <td bgcolor="#9acd32" align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b> (Lønt. / Norm)</td>
	 
	     <%if lto <> "cst" then %>
	     <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_166%></b> (Real / Lønt.)</td>
	     <%end if %>
	 
	 <%end if %>
	 </tr>
	 <tr>
	 <td class=lille style="border-bottom:1px silver dashed;">
      <%=tsa_txt_148%></td>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(realTimer(x),2) %></td>
	 
	  <%if lto <> "cst" then %>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">(<%=formatnumber(realfTimer(x),2)%>)</td>
	 <%end if %>
	 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(normTimer(x),2)%></td>
	 <%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	  <% 
	  ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	 call fLonTimerPer(ltimerStDato, dageDiff, 2, intMid)
	 %>
	 </td>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille>
	  <%call timerogminutberegning(fradragTimer(x)*60) %>
	<%=thoursTot &":"& left(tminTot, 2) %>
	 </td>
	 <%
	 ltimerKorFrad = 0
	 ltimerKorFrad = ((totalTimerPer100/60) + (fradragTimer(x)))
	 %>
	  <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille>
	   <%call timerogminutberegning(ltimerKorFrad*60) %>
	<%=thoursTot &":"& left(tminTot, 2) %></td>
	 <%
	 end if %>
	 
	 <%if lto <> "cst" then %>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=formatnumber((realTimer(x) - normTimer(x)),2)%></b></td>
	 <%end if %>
	 
	 
	 <%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 <%normtime_lontime = -((normTimer(x) - (ltimerKorFrad)) * 60) %>
	 <%call timerogminutberegning(normtime_lontime) %>
	 <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	 </td>
	 <%if lto <> "cst" then %>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber((realTimer(x) - (ltimerKorFrad)),2)%></td>
	 <%end if %>
	 
	 <%end if %>
	 </tr>
	 
	 <%if visning = 2 then %>
	 <!-- MD -Dato  -->
	  <tr>
	         <td class=lille style="border-bottom:1px silver dashed;">
              <%=tsa_txt_272%></td>
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(realTimer(x)/weekDiff,2)%></td>
	         <%if lto <> "cst" then %>
	         <td align=right class=lille style="border-bottom:1px silver dashed;">(<%=formatnumber(realfTimer(x)/weekDiff,2)%>)</td>
	         <%end if %>
	 
	         
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(normTimerUge(x), 2)%></td>
	         <%if session("stempelur") <> 0 then %> 
	         <td align=right style="border-bottom:1px silver dashed;" class=lille>
	          <% 
	          ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	         call fLonTimerPer(ltimerStDato, dageDiff, 4, intMid)
	         %>
	         </td>
	         <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille>
	        <%call timerogminutberegning((fradragTimer(x)/weekDiff)*60) %>
	        <%=thoursTot &":"& left(tminTot, 2) %>
	         </td>
	         <%
	         ltimerKorFrad = 0
	         ltimerKorFrad = ((totalTimerPer100/60) + (fradragTimer(x))/weekDiff)
	         %>
	        <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille>
	         
	         
	         <%call timerogminutberegning(ltimerKorFrad*60) %>
	        <%=thoursTot &":"& left(tminTot, 2) %></td>
	 
	         <%end if %>
	         
	         
	         <%if lto <> "cst" then %>
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=formatnumber(((realTimer(x)/weekDiff) - normTimerUge(x)),2)%></b></td>
	        <%end if %>
	        
	        
	        
	        <%if session("stempelur") <> 0 then %> 
	        <td align=right style="border-bottom:1px silver dashed;" class=lille>
	       
	        <%
	        normtime_lontime = -((normTimerUge(x) - (ltimerKorFrad)) * 60)
	        call timerogminutberegning(normtime_lontime) %>
		    <b><%=thoursTot &":"& left(tminTot, 2) %></b>
	        </td>
	         <%if lto <> "cst" then %>
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(((realTimer(x)/weekDiff) - (ltimerKorFrad)),2)%></td>
        	 <%end if %>
        	<%end if %>
	         </tr>
	         
	       <%end if %>  
	        
	         </table>
	
	    <%       
	 case 6      
	          
	          
	         
	          
	         '***** MD --> Dato Aktyper fordeling / udspec 
	         '** Tjekker om Afspadsering og / eller overarbejde er aktive ***'
	         'Response.Write "###"& afspTimer(x)
	         call akttyper2009(2)
	         'aktiveTyper = akttype_sel
        	  %> 
	         <!-- Afspad / Overarb --->
	         <%if instr(aktiveTyper, "#30#") <> 0 OR instr(aktiveTyper, "#31#") <> 0 then %>
	           <br /><br />
	           
	           
	       
	           
	          <table cellspacing=0 cellpadding=2 border=0 width=100%>
	           <tr>
	         <td colspan=6><b>Overarbejde</b><br />
	         <%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %> (<%=tsa_txt_318 %>)</td>
	         </tr>
	         <tr bgcolor="#EFF3FF">
	        <td align=right class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b>Overarb.</b><br />(enh.)</td>
	          <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b>Afspads.</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Udbetalt</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Ønsk. Udbe.</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_283 &" "& tsa_txt_280 %></b></td>
	         </tr>
	          <tr>
        	 
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(afspTimer(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerBr(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerUdb(x), 2)%></td>
	     
	          <% 
	         afspadUdbBal = 0
	         afspadUdbBal = (afspTimerOUdb(x) - afspTimerUdb(x)) 
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspadUdbBal, 2)%></td>
	          <% 
	         AfspadBal = 0 
	         AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(AfspadBal, 2)%></td>
        	 
	         </tr>
	         </table>
	         <%end if %>
	         
	         
	         
        	 
        	 
            <!-- Syg -->
             <%if level = 1 OR (session("mid") = usemrn) then %>
              <br /><br />
              <b>Sygdom & Sundhed</b><br />
               <%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %> (<%=tsa_txt_318 %>)
	          <table cellspacing=0 cellpadding=2 border=0 width=100%>
	           <tr bgcolor="#EFF3FF">
	           <%if instr(akttype_sel, "#20#") <> 0 then %>
	        <td align=right class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b>Syg</b></td>
	          <%end if %>
	          <%if instr(akttype_sel, "#21#") <> 0 then %>
	          <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b>Barn syg</b></td>
	           <%end if %>
	           <%if instr(akttype_sel, "#8#") <> 0 then %>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Sundhed</b></td>
	           <%end if %>
	           <%if instr(akttype_sel, "#81#") <> 0 then %>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Læge</b></td>
	           <%end if %>
	          </tr>
	          <tr>
	          <%if instr(akttype_sel, "#20#") <> 0 then %>
	          <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(sygTimer(x),2) %></td>
	          <%end if %>
	         <%if instr(akttype_sel, "#21#") <> 0 then %>
	         <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(BarnsygTimer(x),2) %></td>
	         <%end if %>
	        <%if instr(akttype_sel, "#8#") <> 0 then %>
	        <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(sundTimer(x),2) %></td>
	         <%end if %>
	        <%if instr(akttype_sel, "#81#") <> 0 then %>
	        <td align=right class=lille style=" white-space: nowrap;"><%=formatnumber(lageTimer(x),2) %></td>
	        <%end if %>
	        </tr>
	        </table>
	        <%end if%>
        	
        	
	         <br /><br />
	            <table cellspacing=0 cellpadding=2 border=0 width="100%">
        	 
	         <td style="border-bottom:1px silver dashed;">
                 <b><%=tsa_txt_265%></b> <br /><%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %> (<%=tsa_txt_318 %>)</td>
        	 <%if len(trim(km(x))) <> 0 then
        	 km(x) = km(x)
        	 else
        	 km(x) = 0
        	 end if %>
	          <td align=right valign=bottom style="border-bottom:1px silver dashed;"><%=km(x) %> km.</td>
	         </tr>
	         </table>
	
	 
	 
	 
	 
	 <%
	 
	 case 5 '** Status total
	
	 x = 0
	 
	 %>
	 
	 <b><%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %></b> (<%=tsa_txt_319 %>)
	 <table cellspacing=0 cellpadding=2 border=0 width=100%>
	 <tr bgcolor="#EFF3FF">
	 <%
	 select case lto
	 case "x"
	 %>
	 <td style="border-bottom:1px silver dashed;">&nbsp;</td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_172%></b></td>
	  <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_173%></b></td>
	  
	  <%if session("stempelur") <> 0 then %> 
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Løn timer</b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Till. / <br />Frad. +/-</b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Sum</b></td>
	 <%end if %>
	  
	 <%end select%> 
	  
	  <%if lto <> "cst" then %>
	  <td bgcolor="pink" class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b><br /> (Real / Norm)</td>
	  <%end if %>
	  
	 
	  <%if session("stempelur") <> 0 then %> 
	 <td bgcolor="#9acd32" class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_284%></b><br /> (Lønt. / Norm)</td>
	 
	  <%if lto <> "cst" then %>
	  <td class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_166%></b><br /> (Real / Lønt.)</td>
	  <%end if %>  
	    
	 <%end if %>
	 </tr>
	 
	 
	 <tr>
	 <%
	 
	
	 ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	 call fLonTimerPer(ltimerStDato, dageDiff, 3, intMid)
	 
	 ltimerKorFrad = 0
	 ltimerKorFrad = ((totalTimerPer100/60) + (fradragTimer(x)))
	
	 select case lto
	 case "x"
	 %>
	 
	 <td class=lille style="border-bottom:1px silver dashed;">
      &nbsp;</td>
	
	<td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(realTimer(x),2)%></td>
	<td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(normTimer(x),2)%></td>
	 
	 <%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 <%call timerogminutberegning(totalTimerPer100 ) %>
    <%=thoursTot &":"& left(tminTot, 2) %>
	</td>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><%=formatnumber(fradragTimer(x), 2) %></td>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><%=formatnumber(ltimerKorFrad, 2) %></td>
	 <%
	 end if %>
	 
	 
	 <%end select%> 
	 
	 <%if lto <> "cst" then %>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=formatnumber((realTimer(x) - normTimer(x)),2)%></b></td>
	  <%end if %>
	  
	  
	  <%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 <%normtime_lontime = -((normTimer(x) - (ltimerKorFrad)) * 60) %>
	 <%call timerogminutberegning(normtime_lontime) %>
		<b><%=thoursTot &":"& left(tminTot, 2) %></b>
	 </td>
	    <%if lto <> "cst" then %>
	   <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber((realTimer(x) - (ltimerKorFrad)),2)%></td>
	    <%end if %>
	 <%end if %>
	 </tr>
	</table>
	
	       
	         <!-- Afspad / Overarb --->
	       <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>
	           <br /><br />
	          <table cellspacing=0 cellpadding=2 border=0 width=100%>
	           <tr>
	         <td colspan=6><b>Overarbejde</b><br />
	         <%=formatdatetime(startDato, 1) & " - " & formatdatetime(slutDato, 1) %> (<%=tsa_txt_319 %>)</td>
	         </tr>
	         <tr bgcolor="#EFF3FF">
	        <td align=right class=lille style="border-bottom:1px silver dashed; white-space:nowrap;"><b><%=tsa_txt_164%></b><br />(enh.)</td>
	          <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><b>Afspads.</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Udbetalt</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Ønsk. Udbe.</b></td>
	           <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_283 &" "& tsa_txt_280 %></b></td>
	         </tr>
	          <tr>
        	 
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(afspTimer(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerBr(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerUdb(x), 2)%></td>
	     
	          <% 
	         afspadUdbBal = 0
	         afspadUdbBal = (afspTimerOUdb(x) - afspTimerUdb(x)) 
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspadUdbBal, 2)%></td>
	          <% 
	         AfspadBal = 0 
	         AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(AfspadBal, 2)%></td>
        	 
	         </tr>
	         </table>
	         <br /><br />
	         <%end if %>
	         
	         
	  <%case 7 '**** afstem tot, år --> dato ***' 
	  
	  
	   x = 0
	 
	 %>
	 
	 <!-- st -->
	 <tr>
	 <td style="border-bottom:1px silver dashed;"><%=monthname(datepart("m", startDato,2,2)) %></td>
	 
	 <%
	  ltimerStDato = year(stDato) &"/"& month(stDato) &"/"& day(stDato)
	 call fLonTimerPer(ltimerStDato, dageDiff, 3, intMid)
	 
	 ltimerKorFrad = 0
	 ltimerKorFrad = ((totalTimerPer100/60) + (fradragTimer(x)))
	 
	 %>
	 
	
	
	<td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(realTimer(x),2)%></td>
    
    	   <%if lto <> "cst" then %>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">(<%=formatnumber(realfTimer(x),2)%>)</td>
	 <%arealfTimerTot = arealfTimerTot + (realfTimer(x)) %>
	 
	 <%end if %> 
    
    <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(normTimer(x),2)%></td>
	
	<%arealTimerTot = arealTimerTot + (realTimer(x)) %>
	<%anormTimerTot = anormTimerTot + (normTimer(x)) %>
	 
	 <%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 <%call timerogminutberegning(totalTimerPer100) %>
    <%=thoursTot &":"& left(tminTot, 2) %>
	</td>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><%=formatnumber(fradragTimer(x), 2) %></td>
	 <td align=right style="border-bottom:1px silver dashed; white-space:nowrap;" class=lille><%=formatnumber(ltimerKorFrad, 2) %></td>
	 <%
	 atotalTimerPer100 = atotalTimerPer100 + (totalTimerPer100)
	 afradragTimerTot = afradragTimerTot + (fradragTimer(x))
	 altimerKorFradTot = altimerKorFradTot + (ltimerKorFrad)
	 
	 end if %>
	 
	 
	 <%if lto <> "cst" then %>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=formatnumber((realTimer(x) - normTimer(x)),2)%></b></td>
	  <%
	 end if %>
	  

	  
	  <%if session("stempelur") <> 0 then %> 
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 <%normtime_lontime = -((normTimer(x) - (ltimerKorFrad)) * 60) %>
	 <%call timerogminutberegning(normtime_lontime) %>
		<b><%=thoursTot &":"& left(tminTot, 2) %></b>
	 </td>
	    <%if lto <> "cst" then %>
	   <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber((realTimer(x) - (ltimerKorFrad)),2)%></td>
	    <%end if %>
	 <%end if %>
	 
	 
	     <!-- Afspad / Overarb --->
	       <%if instr(akttype_sel, "#30#") <> 0 OR instr(akttype_sel, "#31#") <> 0 then %>
	           
	         <td align=right style="border-bottom:1px silver dashed;" class=lille><%=formatnumber(afspTimer(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerBr(x), 2)%></td>
	         <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspTimerUdb(x), 2)%></td>
	     
	          <% 
	          aafspTimerTot = aafspTimerTot + afspTimer(x) 
	          aafspTimerBrTot = aafspTimerBrTot + afspTimerBr(x)
	          aafspTimerUdbTot = aafspTimerUdbTot + afspTimerUdb(x)
	          
	         afspadUdbBal = 0
	         afspadUdbBal = (afspTimerOUdb(x) - afspTimerUdb(x)) 
	         
	         aafspadUdbBalTot = aafspadUdbBalTot + (afspadUdbBal)
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(afspadUdbBal, 2)%></td>
	          <% 
	         AfspadBal = 0 
	         AfspadBal = (afspTimer(x) - (afspTimerBr(x)+ afspTimerUdb(x)))
	         aAfspadBalTot = aAfspadBalTot + (AfspadBal)
	          %>
	          <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(AfspadBal, 2)%></td>
        	 
	         
	         
	         
	         <%end if %>
	         
	  
	  </tr>
	  
	  <!-- slut -->
	  
	  
	  
	  
	  
	  
	         
	  
	  <%case 4 '**** Ferie ***' %>       
	         
	   <%if lto <> "cst" then %>
	   <!-- Ferie fridage -->
	 <br /><br /><b><%=tsa_txt_282 %></b> (<%=year(startDato)%> - <%=year(slutDato)%>) 
	 <table cellspacing=0 cellpadding=2 border=0 width=100%>
	  <tr bgcolor="#FFFFe1">
	 
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_174 &" "& tsa_txt_164%></b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Planlagt <br> >> dd.</b></td>
	 
	  <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_165%></b></td>
	   <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Udbetalt</b></td>
	  <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_282 &" "& tsa_txt_280 %></b></td>
	 </tr>
	  <tr>
	
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 
	  <%if normTimerDag(x) <> 0 then
	 fefriVal = fefriTimer(x)/normTimerDag(x)
	 else
	 fefriVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(fefriVal,2) %> d. 
	 <!--<br /><%=formatnumber(fefriTimer(x),2) %> t.--></td>
	 
	  <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 
	  <%if normTimerDag(x) <> 0 then
	 fefriplVal = fefriplTimer(x)/normTimerDag(x)
	 else
	 fefriplVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(fefriplVal,2) %> d. 
	 <!--<br /><%=formatnumber(fefriTimer(x),2) %> t.--></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	  <%if normTimerDag(x) <> 0 then
	 fefriBrVal = fefriTimerBr(x)/normTimerDag(x)
	 else
	 fefriBrVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(fefriBrVal,2) %> d. 
	 <!--<br /><%=formatnumber(fefriTimerBr(x),2) %> t.--></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	 <%if normTimerDag(x) <> 0 then
	 fefriUdbVal = fefriTimerUdb(x)/normTimerDag(x)
	 else
	 fefriUdbVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(fefriUdbVal,2) %> d. 
	 <!--<br /><%=formatnumber(fefriTimerUdb(x),2) %> t.--></td>
	  
	  <% 
	 fefriBal = 0
	 fefriBal  = (fefriTimer(x) - (fefriTimerBr(x) + fefriTimerUdb(x)))
	 
	 if normTimerDag(x) <> 0 then
	 fefriBalVal = fefriBal/normTimerDag(x)
	 else
	 fefriBalVal = 0
	 end if
	 
	  %>
	  <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=formatnumber(fefriBalVal,2) %> d.<br />
	 <!-- <%=formatnumber(fefriBal,2)%> t.</b>--></td>
	 </tr>
	 </table>
	 <%end if %>
	 
	 
	 <!-- Ferie -->
	 
	    <br /><br /><b><%=tsa_txt_281 %></b> (<%=year(startDato)%> - <%=year(slutDato)%>) 
	
	  <table cellspacing=0 cellpadding=2 border=0 width=100%>
	 <tr bgcolor="#FFFFe1">
	 <td class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_152%> Opt.</b></td>
	 <td class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_317%><br> >> dd.</b></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Afholdt</b></td>
	   <td align=right class=lille style="border-bottom:1px silver dashed;"><b>Afholdt u. løn</b></td>
	   <td align=right style="border-bottom:1px silver dashed;" class=lille><b>Udbetalt</b></td>
	  <td align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_281 &" "& tsa_txt_280 %></b></td>
	 </tr>
	  <tr>
	
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 
	 <%if normTimerDag(x) <> 0 then
	 ferieOptjVal = ferieOptjtimer(x)/normTimerDag(x)
	 else
	 ferieOptjVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(ferieOptjVal,2) %> d.
	 <!--<br /><%=formatnumber(ferieOptjtimer(x),2) %> t.--></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	 <%if normTimerDag(x) <> 0 then
	 feriePlVal = feriePLTimer(x)/normTimerDag(x)
	 else
	 feriePlVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(feriePlVal,2) %> d. 
	 <!--<br /><%=formatnumber(feriePLTimer(x),2) %> t.--></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	 
	 
	 <%if normTimerDag(x) <> 0 then
	 ferieAfVal = ferieAFTimer(x)/normTimerDag(x)
	 else
	 ferieAfVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(ferieAfVal,2) %> d. 
	 <!--<br /><%=formatnumber(ferieAFTimer(x),2) %> t.--></td>
	 
	  <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	 
	 
	 <%if normTimerDag(x) <> 0 then
	 ferieAfulonVal = ferieAFulonTimer(x)/normTimerDag(x)
	 else
	 ferieAfulonVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(ferieAfulonVal,2) %> d. 
	 <!--<br /><%=formatnumber(ferieAFTimer(x),2) %> t.--></td>
	 
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 
	  <%if normTimerDag(x) <> 0 then
	 ferieUdbVal = ferieUdbTimer(x)/normTimerDag(x)
	 else
	 ferieUdbVal = 0
	 end if
	  %>
	 
	 <%=formatnumber(ferieUdbVal,2) %> d. 
	 <!--<br /><%=formatnumber(ferieUdbTimer(x),2) %> t.--></td>
	  
	  <% 
	 ferieBal = 0
	 ferieBal  = (ferieOptjtimer(x) - (ferieAFTimer(x) + ferieAFulonTimer(x) + ferieUdbTimer(x)))
	 
	 if normTimerDag(x) <> 0 then
	 ferieBalVal = ferieBal/normTimerDag(x)
	 else
	 ferieBalVal = 0
	 end if
	 
	  %>
	  <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(ferieBalVal,2) %> d. <br />
	  <!--<%=formatnumber(ferieBal,2)%> t.--></td>
	 </tr>
	 </table>
	 
	 <%
	 
	 
	 end select 
	 '** visning
	
	end function
	
	
	%>
	
	
	
	