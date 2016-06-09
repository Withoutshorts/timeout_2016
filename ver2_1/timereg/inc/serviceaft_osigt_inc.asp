<!--
Bruges af:
joblog_k.asp
serviceaft_osigt.asp
-->
<%
sub udspHeader
%>
<tr bgcolor="#8caae6">
			
			<td class=lille>&nbsp;Dato</td>
			<td class=lille>&nbsp;Job (nr.)</td>
			<td class=lille>&nbsp;Medarbejder (nr.)</td>
			<td class=lille>&nbsp;Enheder brugt</td>
			
		</tr>
<%
end sub





if thisfile = "serviceaft_osigt.asp" then
toppxThisDiv = 52
leftpxThisDiv = 750
else
toppxThisDiv = 80
leftpxThisDiv = 920
end if%>

<%if request("print") <> "j" then%>
<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv%>px; background-color:#d3d3d3; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
<div style="position:absolute; top:<%=toppxThisDiv%>px; left:<%=leftpxThisDiv+25%>px; padding:2px;">&nbsp;Lukket aftale.</div>
<div style="position:absolute; top:<%=toppxThisDiv+25%>px; left:<%=leftpxThisDiv%>px; background-color:#ffffff; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
<div style="position:absolute; top:<%=toppxThisDiv+25%>px; left:<%=leftpxThisDiv+25%>px; padding:2px;">&nbsp;Enh./ klip baseret.</div>
<div style="position:absolute; top:<%=toppxThisDiv+50%>px; left:<%=leftpxThisDiv%>px; background-color:#EFf3ff; border:1px #000000 solid; height:15px; width:20px; padding:2px;">&nbsp;</div>
<div style="position:absolute; top:<%=toppxThisDiv+50%>px; left:<%=leftpxThisDiv+25%>px; padding:2px;">&nbsp;Periode baseret.</div>		
<%end if%>



	
		
		
		 
		<%
		
		
		if len(request("filter_per")) <> 0 then
		filterKri = request("filter_per")
		response.cookies("so_filterKri")("periode") = filterKri
		else
			if len(request.cookies("so_filterKri")("periode")) <> 0 then
			filterKri = request.cookies("so_filterKri")("periode")
			else
 			filterKri = 0
			end if 
		end if
		
		select case filterKri
		case 0
		chkfilt0 = "CHECKED"
		chkfilt1 = ""
		chkfilt2 = ""
		useFilterKri = ""
		case 1
		chkfilt0 = ""
		chkfilt1 = ""
		chkfilt2 = "CHECKED"
		useFilterKri = " AND advitype = 1 "
		case 2
		chkfilt0 = ""
		chkfilt1 = "CHECKED"
		chkfilt2 = ""
		useFilterKri = " AND advitype = 2 "
		end select
		
		
			filterStatus = request("status")
			select case filterStatus
			case "vislukkede"
			chkfilt3 = ""
			chkfilt4 = ""
			chkfilt5 = "CHECKED"
			useFilterKri = useFilterKri & " AND status = 0 "
			case "visalle"
			chkfilt3 = "CHECKED"
			chkfilt4 = ""
			chkfilt5 = ""
			useFilterKri = useFilterKri & ""
			case else
			chkfilt3 = ""
			chkfilt4 = "CHECKED"
			chkfilt5 = ""
			useFilterKri = useFilterKri & " AND status = 1 "
			end select
	
		
		if id <> 0 then
		thiskid = id
		response.cookies("so_filterKri")("kundekri") = thiskid
		else
			if len(request.cookies("so_filterKri")("kundekri")) <> 0 AND len(request("id")) = 0 then
			thiskid = request.cookies("so_filterKri")("kundekri")
			else
			thiskid = 0
			end if
		end if
		
		
		response.cookies("so_filterKri").expires = date + 10
		%>
		
		
		<%if thisfile = "serviceaft_osigt.asp" then
		
		if func = "osigtall" then
		dtop = 0
		else
		dtop = 10
		end if
		
		
		call filterheader_2013(dtop,0,700,pTxt)
		%>
		
		
		<h4>Aftaleoversigt</h4>
		<table border=0 cellspacing=0 cellpadding=0 bgcolor="#ffffff" width="100%">
		
		<form id=filter name=filter method=post action="<%=thisfile%>?menu=kund&FM_soeg=<%=FM_soeg%>&func=<%=func%>">
				
				<tr><td style="padding-left:10px; padding-top:10px;" valign=top>
				<b>Kontakt:</b> (kunde)<br>
				<%
						strSQL = "SELECT Kkundenavn, Kkundenr, Kid, s.kundeid FROM serviceaft AS s LEFT JOIN kunder AS k ON (k.kid = s.kundeid) WHERE s.id <> 0 AND kid IS NOT NULL GROUP BY s.kundeid ORDER BY Kkundenavn"
                        'Response.write strSQL
                        'Response.flush
                        %>
                        <select name="id" style="width:405px; font-size:11px;" onChange="submit()";>
				        <option value="0">Alle</option>
                        <%
						oRec.open strSQL, oConn, 3
						while not oRec.EOF
						
						if cint(id) = cint(oRec("Kid")) then
						isSelected = "SELECTED"
						else
						isSelected = ""
						end if
						%>
						<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr") %>)</option>
						<%
						oRec.movenext
						wend
						oRec.close
						%>
						
				</select>
				</td>
			
				
			
				<td colspan=2 valign=top style="padding:10px;">
				
				<b>Type:</b><br>
			<input type="radio" name="filter_per" id="filter_per" value="0" <%=chkfilt0%>>Vis Alle&nbsp;&nbsp;<br>
			<input type="radio" name="filter_per" id="filter_per" value="1" <%=chkfilt2%>>Enheds/klip aftaler.&nbsp;&nbsp;<br>
			<input type="radio" name="filter_per" id="filter_per" value="2" <%=chkfilt1%>>Periode aftaler.
		
		</td>
			</tr>
			<tr>
			
			<td style="padding-left:10px; padding-top:10px;"><b>Periode:</b><br>
			<%if fmudato = 1 OR request("brugdatokri") = 1 then
			usedkri1 = "CHECKED"
			usedKri0 = ""
			usedKri2 = ""
			else
			    if fmudato = 0 then
			    usedkri1 = ""
			    usedKri0 = "CHECKED"
			    usedKri2 = ""
			    else
			    usedkri1 = ""
			    usedKri0 = ""
			    usedKri2 = "CHECKED"
			    end if
			end if%>
			<input type="radio" name="FM_usedatokri" value="0" <%=usedkri0%>> Ignorér datointerval.<br>
			<input type="radio" name="FM_usedatokri" value="1" <%=usedkri1%>> Vis kun aftaler med <u>startdato</u> i det valgte interval:<br>
			<input type="radio" name="FM_usedatokri" value="2" <%=usedkri2%>> Vis kun aftaler hvor slutdato er <u>overskreddet</u>:<br>
			<br />
			<!--#include file="weekselector_s.asp"--> <!-- b --><br>&nbsp;
			</td>
		
		<td style="padding:10px;" valign=top><b>Status:</b><br>
			<input type="radio" name="status" id="status" value="visalle" <%=chkfilt3%>>Vis Alle&nbsp;&nbsp;<br>
			<input type="radio" name="status" id="Radio1" value="visaktive" <%=chkfilt4%>>Aktive.&nbsp;&nbsp;<br>
			<input type="radio" name="status" id="Radio2" value="vislukkede" <%=chkfilt5%>>Lukkede.&nbsp;&nbsp;
		</td>
			
		<td valign=bottom align=right style="padding:10px;">
            <input id="Submit1" type="submit" value="Søg >>" /></td>
		</tr>
		</form>
		</table>
		</div>
		
		<!-- filterheader -->
		</table>
		</div>
		<%end if%>
		
		
		
		
		<%'if func <> "osigtall" then%>
			<%if thisfile = "serviceaft_osigt.asp" then
			
			'text = "Opret ny aftale"
			'otoppx = 20
			'oleftpx = 0
			'owdtpx = 150
			'url = "#"
			'java = "serviceaft('0', '"& id &"', '"&FM_soeg&"', '0', '"&func&"')"
			
			'call opretNyJava(url, text, otoppx, oleftpx, owdtpx, java)
              oprTopPX = 12 + (opTopPX)  %>
                <div style="position:absolute; left:750px; top:<%=oprTopPX%>px;">
                <%
                nWdt = 120
                nTxt = "Opret ny aftale"
                nLnk = "#"
                nTgt = "_blank"
                nJav = "serviceaft('0', '"& id &"', '"&FM_soeg&"', '0', '"&func&"')"
                call opretNy_2013_java(nWdt, nTxt, nLnk, nTgt, nJav) %>

                    </div>

			
			
			<%end if%>
		<%'end if%>
		
	 <br><br>
	
	
	
	
	<%
	
	tTop = 0
	tLeft = 0
	tWdth = 900
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>
	
	<table cellspacing="0" cellpadding="1" border="0" width=100%>
	<tr bgcolor="#8caae6">
		<td valign="top" height=20 style="border-left:1px #8caae6 solid;">&nbsp;</td>
		<td style="padding-left:5px;" width=70 class=alt><b>Status</b></td>
		<td style="padding-left:5px;" class=alt><b>Aftalenavn (nr)</b>&nbsp;Kontakt </td>
		<td style="padding-left:5px;" class=alt><b>Varenr</b></td>
		<td style="padding-left:5px;" class=alt align=right><b>Enh. tildelt</b></td>
		<td style="padding-left:5px;" class=alt align=center width=80><b>Periode</b></td>
		<td style="padding-left:5px;" class=alt><b>Realiseret antal<br>Saldo<br>% rest / fornyes ved</b></td>
		<td style="padding-left:5px;" class=alt>&nbsp;</td>
		<td style="padding-left:5px;" class=alt>&nbsp;</td>
		<td valign="top" style="border-right:1px #8caae6 solid;" class=alt>&nbsp;</td>
	</tr>
	<%
	'*** SQL datoer ***
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut

    call akttyper2009(2)
	
	'** Kundekri **
	if thisfile = "joblog_k" then
	kundeIdKri = kundeid 
	else 
	kundeIdKri = id
	end if
	
	if kundeIdKri <> 0 then
	kidKri = " kundeid = "& kundeIdKri & " "
	else
	kidKri = " kundeid <> 0 "
	end if
	
	
	'** Dato kri *** 'FM_visheleperiode
	if fmudato = "1" OR (thisfile = "joblog_k" AND request("FM_ignorerperiode") <> "1") then
	'Response.write "fmudato"& fmudato & "thisfile" & thisfile 
	strDatoKri = " AND stdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
	else
	    if fmudato = "2" then
	    strDatoKri = " AND sldato < '"& strSlutDato &"' AND perafg = 0"
	    else
	    strDatoKri = ""
	    end if
	end if
	
	
	strSQL = "SELECT s.id, s.enheder, s.stdato, s.sldato, s.status, s.navn, s.pris, s.perafg, "_
	&" s.advitype, s.advihvor, s.erfornyet, s.varenr, "_
	&" kkundenavn, kkundenr, kid, s.aftalenr, kundeans1, m.email, s.overfortsaldo "_
	&" FROM serviceaft s "_
	&" LEFT JOIN kunder ON (kid = kundeid)"_
	&" LEFT JOIN medarbejdere m ON (m.mid = kundeans1) "_
	&" WHERE "& kidKri &" "& useFilterKri &" "& strDatoKri &" ORDER BY kkundenavn, id DESC"
	oRec.open strSQL, oConn, 3 
	s = 0
	'enhTot = 0
	
	'Response.write strSQL
	
	antalBrugteEnhederTot = 0
	antalTildelteEnhederTot = 0
	
	
	while not oRec.EOF 
	
	
	select case oRec("status")
	case 1
		
		
		select case oRec("advitype")
		case 0, 1
		bgt = "#FFFFFF"
		sttus = "<font color=forestgreen>Aktiv</font><br> Type: Enh./klip"
		case 2
		bgt = "#EFf3FF"
		sttus = "<font color=forestgreen>Aktiv</font><br> Type: Periode"
		end select
	
	case 0
				select case oRec("advitype")
				case 0, 1
				sttus = "<font color=red>Lukket</font><br>Type: Enh./klip"
				bgt = "#dcdcdc"
				case 2
				sttus = "<font color=red>Lukket</font><br>Type: Periode"
				bgt = "#dcdcdc"
				end select
				
	
	end select
	
	%>
	
	<%if request("print") = "j" AND s <> 0 then%>
	<tr bgcolor="#ffffff">
		<td height=30 valign="top" style="border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=8 style="border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" style="border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<%end if%>
	
	
	
	<tr bgcolor="<%=bgt%>">
		<td valign="top">&nbsp;</td>
		<td style="padding:2px 5px 0px 5px;" valign=top>
		<a href="javascript:expand('<%=s%>');"><img src="ill/plus.gif" width="9" height="9" border="0" name="Menub<%=s%>" id="Menub<%=s%>"></a>
		&nbsp;<%=sttus%>
		</td>
		
		<td style="padding-left:5px;">
		<%if thisfile = "serviceaft_osigt.asp" then%>
		<a href="#" onClick="serviceaft('<%=oRec("id")%>', '<%=oRec("kid")%>', '<%=FM_soeg%>', '0')" class=vmenu><%=oRec("navn")%> (<%=oRec("aftalenr")%>)</a>
		<%else%>
		<a href="#" class=vmenu><%=oRec("navn")%> (<%=oRec("aftalenr")%>)</a>
		<%end if%>
		
		<br>
		
		<b><%if len(oRec("kkundenavn")) > 30 then
		Response.write left(oRec("kkundenavn"), 30) &".."
		else
		Response.write oRec("kkundenavn")
		end if%> (<%=oRec("kkundenr")%>)</b>
		
		<font class=megetlillesort>
		<%'*** job tilknytning ***
		
		strSQLjob = "SELECT j.jobnavn, j.jobnr FROM job j WHERE j.serviceaft = "& oRec("id") &" ORDER BY j.jobnavn"
		oRec2.open strSQLjob, oConn, 3
		j = 0
		while not oRec2.EOF
		if j <> 0 then%>
		,
		<%else %>
		<br />Job: 
		<%end if %>
		<%=oRec2("jobnavn") %> (<%=oRec2("jobnr") %>)
		 
		<%
		j = j + 1
		oRec2.movenext
		wend 
		oRec2.close
		
		 %></font>
		
		</td>
		<td style="padding-left:5px;" valign=top><%=oRec("varenr")%></td>
		<td style="padding-right:5px;" valign=top align=right><b><%=formatnumber(oRec("enheder"))%></b></td>
		<td class=lille style="padding:0px 10px 0px 5px;" valign=top align=right>
		<%if oRec("perafg") = 0 then%>
		<%=formatdatetime(oRec("stdato"), 2)%> <br> <%=formatdatetime(oRec("sldato"), 2)%>
		<%else%>
		<%=formatdatetime(oRec("stdato"), 2)%> <br /> (Ingen)
		<%end if%></td>
		<td style="padding-left:5px;" valign=top>	
				<%
				select case oRec("advitype")
				case 0
				'strForny = "Nej"
				skalfornyes = "-"
				bgcolforny = bgt
				case 1
				'strForny = "Enh"
						
						select case oRec("advihvor")
						case 1
						when = 50
						whenTxt = "50%"
						case 2
						when = 25
						whenTxt = "25%"
						case 3
						when = 5
						whenTxt = "5%"
						case 4
						when = -1
						whenTxt = "99%"
						case 5
						when = -100000000
						whenTxt = "(Aldrig)"
						end select 
						
						antalbrugteklip = 0
						strSQLsum = "SELECT sum(t.timer) AS sumtimer, t.taktivitetid, a.faktor "_
						&" FROM timer t LEFT JOIN aktiviteter a ON (a.id = t.taktivitetid) "_
						&" WHERE t.seraft = "& oRec("id") &" AND ("& aty_sql_realhours &") GROUP BY t.taktivitetid"
						
						'response.write strSQLsum
						'Response.flush
						oRec3.open strSQLsum, oConn, 3 
						while not oRec3.EOF 
							antalbrugteklip = antalbrugteklip + oRec3("sumtimer") * oRec3("faktor")
						oRec3.movenext
						wend 
						oRec3.close 
						
						antalBrugteEnhederTot = antalBrugteEnhederTot + antalbrugteklip
						'Response.write antalbrugteklip & "&nbsp;"
						
						
						if oRec("enheder") < 1 then
						enheder = 1
						else
						enheder = oRec("enheder")
						end if
						
						fstatus = 0
						fstatus = 100 - ((antalbrugteklip/enheder) * 100)
						
						skalfornyes = "<a href='#' onClick=""serviceaft('"&oRec("id")&"', '"&oRec("kid")&"', '"&FM_soeg&"', '1')"" class=vmenu>Forny nu</a> <font class=roed><b>!</b></font>"
							
						
						if cdbl(fstatus) <= 0 AND oRec("status") = 1 then 
							'if oRec("erfornyet") = 0 then
							bgcolforny = "pink"
							'else
							'skalfornyes = "<font class=lillesort>Er fornyet.</font>"
							'skalfornyes = ""
							'bgcolforny = bgt
							'end if
						else
						'skalfornyes = ""
						'skalfornyes = "<a href='#' onClick=""serviceaft('"&oRec("id")&"', '"&oRec("kid")&"', '"&FM_soeg&"', '1')"" class=vmenu>Forny nu</a> <font class=roed><b>!</b></font>"
						bgcolforny = bgt
						end if
						
						
						if oRec("status") = 1 then
						skalfornyes = "<a href='#' onClick=""serviceaft('"&oRec("id")&"', '"&oRec("kid")&"', '"&FM_soeg&"', '1')"" class=vmenu>Forny nu</a> <font class=roed><b>!</b></font>"
						else
						skalfornyes = ""
						end if
						
						
						Response.write "<b>"& formatnumber(antalbrugteklip, 2) & "</b> klip"
						if oRec("overfortsaldo") <> 0 then
						Response.write  "<font color=crimson> / overført: "& formatnumber(oRec("overfortsaldo"), 2) &"</font> "
						end if
						
						saldopaaaftale = 0
						saldopaaaftale = oRec("enheder") - (antalbrugteklip + (oRec("overfortsaldo")))
						
						Response.write "<br>Saldo: <b>" & formatnumber(saldopaaaftale, 2) &"</b>"
						
						Response.write  "<br><font class=megetlillesilver>"& formatnumber(fstatus, 0) &"% rest / forny v. " & whenTxt &"</font>&nbsp;&nbsp;"
						
				case 2
				'strForny = "Per"
				
						select case oRec("advihvor")
						case 1
						when = 90
						whenTxt = "3 md. før."
						case 2
						when = 30
						whenTxt = "30 dage før."
						case 3
						when = 10
						whenTxt = "10 dage før."
						case 4
						when = 1
						whenTxt = "1 dag før."
						case 5
						when = -100000000
						whenTxt = "(Aldrig)"
						end select 
						
						restperiode =  datediff("d", year(now)&"/"&month(now)&"/"&day(now), oRec("sldato"))
						
						if cint(restperiode) <= when AND oRec("status") = 1 then
							'if oRec("erfornyet") = 0 then
							bgcolforny = "pink"
							'else
							'skalfornyes = "<font class=lillesort>Er fornyet.</font>"
							'skalfornyes = ""
							'bgcolforny = bgt
							'end if
						else
						'skalfornyes = ""
						'skalfornyes = "<a href='#' onClick=""serviceaft('"&oRec("id")&"', '"&oRec("kid")&"', '"&FM_soeg&"', '1')"" class=vmenu>Forny nu</a> <font class=roed><b>!</b></font>"
						bgcolforny = bgt
						end if
						
						if oRec("status") = 1 then
						skalfornyes = "<a href='#' onClick=""serviceaft('"&oRec("id")&"', '"&oRec("kid")&"', '"&FM_soeg&"', '1')"" class=vmenu>Forny nu</a> <font class=roed><b>!</b></font>"
						else
						skalfornyes = ""
						end if
						
						Response.write "<b>"& restperiode &"</b> dage endnu.<br><font class=megetlillesilver>" & whenTxt &"</font>&nbsp;&nbsp;"
						
						
						
				end select 
				
				
				Response.write strWhen
				
				if thisfile <> "serviceaft_osigt.asp" then
				%>
				<br><a href="mailto:<%=oRec("email")%>&subject=Anmodning om fornyelse af aftale." class=vmenu>Anmod om ny aft?</a>
				<%
				end if
				%>
		</td>
		<%if thisfile = "serviceaft_osigt.asp" then%>
		<td style="padding-left:5px;" bgcolor="<%=bgcolforny%>">&nbsp;<%=skalfornyes%></td>
		<td style="padding-left:5px;"><a href="<%=thisfile%>?menu=kund&func=sletsaft&id=<%=oRec("kid")%>&saftid=<%=oRec("id")%>&FM_soeg=<%=FM_soeg%>&visalle=<%=func%>&FM_usedatokri=<%=fmudato%>&filter_per=<%=filterKri%>&status=<%=filterStatus%>"><img src="../ill/slet_16.gif" alt="Slet" border="0"></a></td>
		<%else%>
		<td colspan=2>&nbsp;</td>
		<%end if%>
		<td valign="top">&nbsp;</td>
	</tr>
	
	
	
	<!-------------------------------->	
	<tr bgcolor="<%=bgt%>">
	<td valign="top" colspan=9 bgcolor="<%=bgt%>" style="padding:20px;">
	
	
	<%
	if request("print") = "j" then
	udspDsp = ""
	else
	udspDsp = "none"
	end if%>
	
	<div ID="Menu<%=s%>" Style="position: relative; display: <%=udspDsp%>;">	
	<br><b>Forbrug:</b>
	<table cellspacing=0 cellpadding=2 border=0 width=100%>
	<tr>
		<td bgcolor="#ffffff" colspan=4 style="font-size:10px; color:#999999;">Overførte klip/enheder fra tidligere aftaler: <b><%=oRec("overfortsaldo")%></b></td>
	</tr>
	
	<%
	antalbrugteklipThis = 0
	'**** Udspecificering ****
	'strSQLuds = "SELECT j.id, j.jobnr, j.jobnavn, j.jobnr, t.timer, t.tdato, t.taktivitetid, a.faktor, t.Tmnavn, t.Tmnr "_
	'&" FROM job j "_
	'&" LEFT JOIN timer t ON (t.tjobnr = j.jobnr) "_
	'&" LEFT JOIN aktiviteter a ON (a.id = t.taktivitetid) "_
	'&" WHERE t.seraft = "& oRec("id") &" AND tfaktim = 1 "
	
	if thisfile = "joblog_k" then
	jobkundeSQLkri = " AND j.kundeok = 1"
	else
	jobkundeSQLkri = ""
	end if
	
	strSQLuds = "SELECT j.id, j.jobnr, j.jobnavn, j.jobnr, t.timer, t.tdato, t.taktivitetid, a.faktor, t.Tmnavn, t.Tmnr, j.jobstartdato "_
	&" FROM timer t "_
	&" LEFT JOIN job j ON (j.jobnr = t.tjobnr "& jobkundeSQLkri &")"_
	&" LEFT JOIN aktiviteter a ON (a.id = t.taktivitetid) "_
	&" WHERE t.seraft = "& oRec("id") &" AND ("& aty_sql_realhours &") ORDER BY t.tdato DESC"
	
	
	'response.write strSQLsum
	'Response.flush
	oRec3.open strSQLuds, oConn, 3 
	u = 0
	while not oRec3.EOF 
		if u = 0 then
		
		call udspHeader
		
		end if
		
		
		antalbrugteklipThis = oRec3("timer") * oRec3("faktor")
		
		select case right(u, 1)
		case 0, 2, 4, 6, 8
		bgt = "#FFFFFF"
		case else
		bgt = "#dcdcdc"
		end select
		
		%>
		<tr bgcolor="<%=bgt %>">
			
			<td class=lille>&nbsp;<%=oRec3("tdato")%></td>
			<td class=lille>&nbsp;<%=oRec3("jobnavn")%> (<%=oRec3("jobnr")%>)</td>
			<td class=lille>&nbsp;<%=oRec3("tmnavn")%> (<%=oRec3("tmnr")%>)</td>
			<td class=lille>&nbsp;<%=antalbrugteklipThis%></td>
			
		</tr>
		<%
		u = u + 1
	oRec3.movenext
	wend 
	oRec3.close 
	
	
	if u = 0 then
	
	call udspHeader
	%>
	<tr>
		<td bgcolor="#ffffff" colspan=4 class=lille>&nbsp;Ingen registreringer.</td>
	</tr>
	<%end if%>
	
	
	</table><br>&nbsp;
	</div>
	</td>
	<td valign="top"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<!-------------------------------->	
		
	
	<%
	if request("print") = "j" then
	hgt = 10
	else
	hgt = 1
	end if
	%>
	
	<tr bgcolor="<%=bgt%>">
		<td height=<%=hgt%> valign="top" style="border-bottom:1px #CCCCCC solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td colspan=8 style="border-bottom:1px #CCCCCC solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" style="border-bottom:1px #CCCCCC solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<%
	'ekspTxt = ekspTxt & left(oRec("kkundenavn"), 25) & chr(009) & chr(009) & chr(009) & oRec("kkundenr") & chr(009) & oRec("varenr") & chr(009) & left(oRec("navn"), 25) & chr(009)& chr(009)& chr(009) & oRec("enheder") & chr(009) & formatnumber(oRec("pris"), 2) & chr(009) & formatdatetime(oRec("stdato"), 2) & " til "& formatdatetime(oRec("sldato"), 2) & chr(013)
	
	'*** Kun enheds/klip afhængige aftaler bregnes med her ***
	select case oRec("advitype")
	case 1
	antalTildelteEnhederTot = antalTildelteEnhederTot + oRec("enheder")
	case else
	antalTildelteEnhederTot = antalTildelteEnhederTot
	end select
	
	s = s + 1
	oRec.movenext
	wend
	oRec.close 
	
	if s = 0 then%>
	<tr bgcolor="#ffffff">
		<td valign="top" style="border-left:1 #8caae6 solid; border-bottom:1px #8caae6 solid;">&nbsp;</td>
		<td colspan="8" style="border-bottom:1px #8caae6 solid; padding:15px;" height=30><br><b>Der blev ikke fundet nogen aftaler</b><br />
		der matcher de valgte filterkriterier.
		</td>
		<td valign="top" style=" border-bottom:1px #8caae6 solid; border-right:1 #8caae6 solid;">&nbsp;</td>
	</tr>
	<%end if%>
	
	
	
	</table>
	
	<!-- tablediv -->
	</div>
	
	
	
	<br><br><br>
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#FFFFFF" width=400>
	<tr>
		<td valign="top" style="border-left:1px #8caae6 solid; border-top:1px #8caae6 solid;">&nbsp;</td>
		<td colspan="8" style="border-top:1px #8caae6 solid;""><br>
		<b>Balance på <u>Enheds / Klip</u> afhængige aftaler:</b> <br>(Kun realiserede enheder er medregnet her. Ikke overførte saldi)
		</td>
		<td valign="top" style="border-right:1px #8caae6 solid; border-top:1px #8caae6 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td valign="top" style="border-left:1 #8caae6 solid;">&nbsp;</td>
		<td><br><b>Enheder tildelt ialt:</b></td>
		<td align=right><br><b><%=formatnumber(antalTildelteEnhederTot, 2)%></b></td>
		<td colspan="6">&nbsp;</td>
		<td valign="top" style="border-right:1 #8caae6 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td valign="top" style="border-left:1 #8caae6 solid;">&nbsp;</td>
		<td style="border-bottom:1px #8caae6 solid;"><b>Realiseret:</b></td>
		<td align=right style="border-bottom:1px #8caae6 solid;">
		<b><%=formatnumber(antalBrugteEnhederTot, 2)%></b></td>
		<td colspan="6">&nbsp;&nbsp;(Faktor korriget)</td>
		<td valign="top" style="border-right:1 #8caae6 solid;">&nbsp;</td>
	</tr>
	<tr>
		<td valign="top" style="border-left:1 #8caae6 solid;">&nbsp;</td>
		<td style="border-bottom:1px #8caae6 solid;"><b>Balance/Saldo:</b></td>
		<td align=right style="border-bottom:1px #8caae6 solid;"><u><b><%=formatnumber(antalTildelteEnhederTot - antalBrugteEnhederTot, 2)%></b></u></td>
		<td colspan="6">&nbsp;</td>
		<td valign="top" style="border-right:1 #8caae6 solid;">&nbsp;</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td valign="top" style="border-left:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
		<td colspan="8" style="border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		<td valign="top" align="right" style="border-right:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
	</tr>
	</table>
	
	<br><br>

    <%if thisfile = "serviceaft_osigt.asp" then%>

	<div style="position:relative; visibility:visible; padding:20px 20px 20px 5px; width:550px;"><h4>Hjælp & Sideinfo:</h4>
	
		
		
		Aftaler er rammeaftaler med <b>Kontakter</b> vedr. f.eks support, kampagner o.lign.<br>
		Når der oprettes et nyt job kan det vælges om det skal oprettes som en del af en eksisterende <b>Aftale</b>.<br>
		Kun timer/enheder indtastet på <b>fakturerbare aktiviteter</b> medregnes når antallet af brugte enheder skal findes.<br>&nbsp;
		
		
</div>
	<%end if%>
