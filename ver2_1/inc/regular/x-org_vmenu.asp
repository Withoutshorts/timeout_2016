<div id="vmenu_usr" style="position:absolute; left:10; top:60; visibility:visible;">
<table cellpadding="0" cellspacing="0" border="0">
<tr>
<%'************* TSA/CRM menu *********************
if menu = "crm" then
showmenu = "crm"
else
showmenu = request("mainmenu")
end if

select case showmenu
case "crm"
%>
<td><img src="../ill/v-menu_menu.gif" alt="" border="0"></td>
</tr>
</table>	
<table cellspacing="0" cellpadding="2" border="0" bgcolor="#EFF3FF" width="159">
<tr>
<td align="right"><%=varSelectedCRM_kal%></td>
<td><a href="crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal&status=0&id=0&emner=0" class='vmenu'>CRM Kalender</a></td>
</tr>
<tr>
<td align="right"><%=varSelectedCRM_osigt%></td>
<td><a href="kunder.asp?menu=crm&shokselector=1&ketype=e&selpkt=osigt" class="vmenu">CRM Firma kontakter</a></td>
</tr>
<!--
<tr>
<td align="right"><=varSelectedCRM_oprfirma%></td>
<td><a href='kunder.asp?menu=crm&ketype=e&func=opret&id=0&selpkt=oprfirma' class='vmenu'>Opret nyt firma&nbsp;<img src='../ill/pil_groen_lille.gif' width='17' height='17' alt='' border='0'></a></td>
</tr>-->
<td align="right"><%=varSelectedCRM_hist%></td>
<td><a href='crmhistorik.asp?menu=crm&ketype=e&func=hist&id=0&selpkt=hist' class='vmenu'>Aktions historik</a></td>
</tr>
<%if level <= 2 OR level = 6 then%>
<tr>
<td align="right"><%=varSelectedCRM_status%></td>
<td><a href='crmstatus.asp?menu=crm&ketype=e&selpkt=status' class='vmenu'>Status</a></td>
</tr>
<tr>
<td align="right"><%=varSelectedCRM_emn%></td>
<td><a href='crmemne.asp?menu=crm&ketype=e&selpkt=emn' class='vmenu'>Emner</a></td>
</tr>
<td align="right"><%=varSelectedCRM_kf%></td>
<td><a href='crmkontaktform.asp?menu=crm&ketype=e&selpkt=kf' class='vmenu'>Kontaktform</a></td>
</tr>
<tr>
	<td align="right">&nbsp;</td>
	<td><a href='medarb.asp?menu=medarb' class='vmenualt'>Medarbejdere</a></td>
</tr>
<%end if%>
<%case else%>
<td><img src="../ill/v-menu_menu.gif" alt="" border="0"></td>
</tr>
</table>	

<table cellspacing="0" cellpadding="2" border="0" bgcolor="#EFF3FF" width="159">
<tr>
	<td align="right"><%=varSelectedTimereg%></td>
	<td><a href="timereg.asp?menu=timereg" class='vmenu'>Timeregistrering</a></td>
</tr>
<%if level <= 2 OR level = 6 then%>
<tr>
	<td align="right"><%=varSelectedJobs%></td>
	<td><a href="jobs.asp?menu=job&shokselector=1" class='vmenu'>Jobs</a></td>
</tr>
<%end if%>
<%if level <= 2 OR level = 6 then%>
<tr>
	<td align="right"><%=varSelectedKund%></td>
	<td><a href="kunder.asp?menu=kund" class='vmenu'>Kunder</a></td>
</tr>
<%end if%>
<tr>
	<td align="right"><%=varSelectedMed%></td>
	<td><a href="medarb.asp?menu=medarb" class='vmenu'>Medarbejdere</a></td>
</tr>
<%if level <= 2 OR level = 6 then%>
<tr>
	<td align="right"><%=varSelectedStat%></td>
	<td><a href="stat.asp?menu=stat&shokselector=1&FM_kunde=0&FM_progrupper=10" class='vmenu'>Statistik og Fakturering</a></td>
</tr>
<%end if%>
<%end select%>
</table>
</div>



<%
'************ TSA submenu **************************

select case menu 
case "medarb"
varTopGif = "medarb"
	if level <= 2 OR level = 6 then 
	Links = 2
	else
	Links = 0
	end if
Redim link(Links)
link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='medarb.asp?menu=medarb' class='vmenu'>Medarbejdere</a></td>"
	if level <= 2 OR level = 6 then 
	link(1) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='bgrupper.asp?menu=medarb' class='vmenu'>Brugergrupper</a></td>"
	link(2) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='medarbtyper.asp?menu=medarb' class='vmenu'>Medarbejdertyper</a></td>"
	end if
case "kund"
varTopGif = "kund" 
Links = 1
Redim link(Links)
link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='kunder.asp?menu=kund' class='vmenu'>Kunder</a></td>"
link(1) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='kunder.asp?menu=kund&func=opret' class='vmenu'>Opret Kunde</a></td>"
case "job"
varTopGif = "jobs" 
Links = 5
Redim link(Links)
link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='jobs.asp?menu=job&shokselector=1' class='vmenu'>Jobs</a></td>"
link(1) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='jobs.asp?menu=job&func=opret&id=0&int=1' class='vmenu'>Opret nyt <img src='../ill/eksternt_job_trans.gif' width='15' height='15' alt='' border='0'> job&nbsp;</a></td>"
link(2) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='jobs.asp?menu=job&func=opret&id=0&int=0' class='vmenu'>Opret nyt <img src='../ill/internt_job_trans.gif' width='15' height='15' alt='' border='0'> job&nbsp;</a></td>"
link(3) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='projektgrupper.asp?menu=job' class='vmenu'>Projektgrupper</a></td>"
'link(4) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='aktiv.asp?menu=job&func=favorit&id=0' class='vmenu'>Stam-aktiviteter</a></td>"
link(4) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='akt_gruppe.asp?menu=job&func=favorit' class='vmenu'>Stam-aktiviteter og Grp.</a></td>"
link(5) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='pg_allokering.asp?menu=job' class='vmenu'>Jobplanner</a></td>"
case "stat"
varTopGif = "stat" 
Links = 0
Redim link(Links)
link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='stat.asp?menu=stat' class='vmenu'>Ny Kørsel</a></td>"
case "crm"
varTopGif = "blank" 
Links = 0
Redim link(Links)
case else
varTopGif = "timereg" 
Links = 3
Redim link(Links)
link(0) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='timereg.asp?menu=tiemreg' class='vmenu'>Timeregistrering</a></td>"
link(1) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='joblog.asp?menu=timereg&FM_medarb="&session("mid")&"&FM_job=0&selmedarb="&session("mid")&"' class='vmenu'>Medarbejder Log</a></td>"
link(2) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='joblog_korsel.asp?menu=timereg&FM_medarb="&session("mid")&"&FM_job=0&selmedarb="&session("mid")&"' class='vmenu'>Medarbejder Kørsel</a></td>"
link(3) = "<td align='right'><img src='../ill/pile_selected.gif' alt='' border='0'></td><td><a href='timereg.asp?func=showguide' class='vmenu'>Guiden dine aktive job&nbsp;<font color='darkred'>!</font></a></td>"
%>
<%end select%>


<div id="vmenudiv" style="position:absolute; left:10; top:200; width:130; height:300; visibility:visible;">
<%if menu = "crm" then
Response.write "<br><br>"
end if%>
<table cellspacing="0" cellpadding="0" border="0" width="159">
<tr>
	<td colspan="2"><img src="../ill/<%=varTopGif%>_top_blaa.gif" alt="" border="0"></td>
</tr>
</table>
</table>	
<table cellspacing="0" cellpadding="2" border="0" bgcolor="#EFF3FF" width="159">
<%for intCounter = 0 to Links%>
<tr bgcolor="EFF3FF">
<%=link(intCounter)%>
</tr>
<%next%>
</table>
<br><br>
<%
'** slut submenu ****


'**** Søg på kunde / firma ****
if menu = "timereg" OR menu = "" OR menu = "crm" then

			if menu = "crm" then
			ketypeKri = " ketype <> 'k' "
			
			select case selpkt
			case "hist"
			actionLine = "crmhistorik.asp?menu=crm&func=hist&selpkt=hist"
			case "osigt"
			actionLine = "kunder.asp?menu=crm&selpkt=osigt&ketype=e" 
			case else
			actionLine = "crmkalender.asp?selpkt=kal&menu=crm"
			end select
			
			
				if len(request("emner")) <> 0 then 
				actionLine = actionLine &"&emner="&request("emner")&""
				end if
				
				if len(request("status")) <> 0 then 
				actionLine = actionLine &"&status="&request("status")
				end if
				
				if len(request("medarb")) <> 0 then 
				actionLine = actionLine &"&medarb="&request("medarb")
				end if
			
			actionline = actionline &"&FM_start_dag="&strDag&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&""
			name = "id"
			imgsearch = "../ill/v-menu_soeg_f.gif"
			else
			ketypeKri = " ketype <> 'e' "
			actionLine = "timereg.asp?menu=timereg&FM_use_me="&request("FM_use_me")
			name = "searchstring"
			imgsearch = "../ill/v-menu_soeg.gif"
			end if
%>
<table cellspacing="0" cellpadding="0" border="0" width="120">
<tr>
	<td><img src="<%=imgsearch%>" alt="" border="0"></td>
</tr>
<tr bgcolor="#EFF3FF">
	<td height="30" valign="middle" align="center">	
	<table>
	<%
	if thisfile = "timereg" then
	writethis = "(Dine aktive job)"
	else
	writethis = ""
	end if
	%>
	<tr><form action="<%=actionLine%>" method="post">
	<td><select name="<%=name%>" size="1" style="font-size : 11px;">
	<option value="0">Alle <%=writethis%></option>
	<%
			strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" ORDER BY Kkundenavn"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			if cint(request(name)) = cint(oRec("Kid")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			%>
			<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=left(oRec("Kkundenavn"), 15)%></option>
			<%
			oRec.movenext
			wend
			oRec.close
			%>
	</select></td>
	<td><input type="image" src="../ill/pil_groen_lille.gif" border="0"></td>
	</tr></form>
	</table></td>
	</tr>
	

<%end if


				'**** Filter efter emne, status og medarbejder (Kun CRM)****
				if selpkt = "hist" then
					
					%>
					
					<tr bgcolor="#064CB9">
						<td height="20">&nbsp;&nbsp;<font class="stor-hvid">Vis emne:</td>
					</tr>
					<tr bgcolor="#EFF3FF">
						<td height="30" valign="middle" align="center">	
						<table>
						<tr><form action="crmhistorik.asp?menu=crm&func=hist&selpkt=hist&id=<%=id%>&status=<%=status%>&medarb=<%=medarb%>" method="post">
						<td><img src="../ill/blank.gif" alt="" border="0" width="10" height="1"><select name="emner" size="1" style="font-size : 11px;">
						<option value="0">Alle</option>
						<%
								strSQL = "SELECT navn, id FROM crmemne ORDER BY navn"
								oRec.open strSQL, oConn, 3
								while not oRec.EOF
								
								if cint(request("emner")) = cint(oRec("id")) then
								isSelected = "SELECTED"
								else
								isSelected = ""
								end if
								%>
								<option value="<%=oRec("id")%>" <%=isSelected%>><%=left(oRec("navn"), 15)%></option>
								<%
								oRec.movenext
								wend
								oRec.close
								%>
						</select></td>
						<td><input type="image" src="../ill/pil_groen_lille.gif" border="0"></td>
						</tr></form>
						</table></td>
						</tr>
					
					<tr bgcolor="#064CB9">
						<td height="20">&nbsp;&nbsp;<font class="stor-hvid">Vis status:</td>
					</tr>
					<tr bgcolor="#EFF3FF">
						<td height="30" valign="middle" align="center">	
						<table>
						<tr><form action="crmhistorik.asp?menu=crm&func=hist&selpkt=hist&id=<%=id%>&emner=<%=emner%>&medarb=<%=medarb%>" method="post">
						<td><img src="../ill/blank.gif" alt="" border="0" width="10" height="1"><select name="status" size="1" style="font-size : 11px;">
						<option value="0">Alle</option>
						<%
								strSQL = "SELECT navn, id FROM crmstatus ORDER BY navn"
								oRec.open strSQL, oConn, 3
								while not oRec.EOF
								
								if cint(request("status")) = cint(oRec("id")) then
								isSelected = "SELECTED"
								else
								isSelected = ""
								end if
								%>
								<option value="<%=oRec("id")%>" <%=isSelected%>><%=left(oRec("navn"), 15)%></option>
								<%
								oRec.movenext
								wend
								oRec.close
								%>
						</select></td>
						<td><input type="image" src="../ill/pil_groen_lille.gif" border="0"></td>
						</tr></form>
						</table></td>
						</tr>
					
					<%end if%>
					
					
			<%
			'**** Medarbejder filter ********
			if level <= 2 AND menu = "crm" then
			
			select case selpkt
			case "osigt"
			actionline = "kunder.asp?&ketype=e&menu=crm&func=hist&selpkt=osigt&id="&id&"&emner="&emner&"&status="&status 
			name = "medarb"
			case "hist"
			actionline = "crmhistorik.asp?menu=crm&func=hist&selpkt=hist&id="&id&"&emner="&emner&"&status="&status 
			'***tilføjer dato **
			actionline = actionline &"&FM_start_dag="&strDag&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&""
			name = "medarb"
			case else
			actionline = "crmkalender.asp?menu=crm&func=hist&selpkt=kal&id="&id&"&emner="&emner&"&status="&status 
			'***tilføjer dato **
			actionline = actionline &"&FM_start_dag="&strDag&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&""
			name = "medarb"
			end select
			
			
			%>
			
			<tr bgcolor="#064CB9">
						<td height="20">&nbsp;&nbsp;<font class="stor-hvid">Vis for medarbejder:</td>
					</tr>
			<tr bgcolor="#EFF3FF">
				<td height="30" valign="middle" align="center">	
				<table>
				<tr><form action="<%=actionline%>" method="post">
				<td><img src="../ill/blank.gif" alt="" border="0" width="10" height="1">
				<select name="<%=name%>" size="1" style="font-size : 11px;">
				<option value="0">Alle</option>
				<%
						strSQL = "SELECT mnavn, mid FROM medarbejdere ORDER BY mnavn"
						oRec.open strSQL, oConn, 3
						while not oRec.EOF
						
							if cint(medarb) = cint(oRec("mid")) then
							isSelected = "SELECTED"
							else
							isSelected = ""
							end if
						
						%>
						<option value="<%=oRec("mid")%>" <%=isSelected%>><%=left(oRec("mnavn"), 15)%></option>
						<%
						oRec.movenext
						wend
						oRec.close
						%>
				</select></td>
				<td><input type="image" src="../ill/pil_groen_lille.gif" border="0"></td>
				</tr></form>
				</table></td>
				</tr>
			
			<%end if%>

</table>
<%
'**** Stat ****
if menu = "job" AND request("shokselector") = 1 OR menu = "stat" AND request("shokselector") = 1 then

	if menu = "job" then
	actionLine = "jobs.asp?menu=job&sort=kunkunde&shokselector=1"
	end if
	
	if menu = "stat" then
	actionLine = "stat.asp?menu=stat&sort=kunkunde&shokselector=1"
	end if%>

<table cellspacing="0" cellpadding="0" border="0" width="120">
<tr><form method="post" action="<%=actionLine%>">
	<td colspan="2"><img src="../ill/v-menu_soeg.gif" alt="" border="0"></td>
</tr>
<tr bgcolor="#EFF3FF">
	<td height="30" valign="middle">&nbsp;
	<select name="FM_kunde" size="1" style="font-size : 11px;">
	<option value="0">Alle</option>
	<%
			strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			if cint(request("FM_Kunde")) = cint(oRec("Kid")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			%>
			<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=left(oRec("Kkundenavn"), 16)%></option>
			<%
			oRec.movenext
			wend
			oRec.close
			%>
	</select>&nbsp;
</td>
<td><input type="image" src="../ill/pil_groen_lille.gif" border="0"></td>
</tr>

<%if menu = "stat" AND request("shokselector") = 1 then%>
<tr bgcolor="#064CB9">
	<td height="20" colspan="2">&nbsp;&nbsp;<font class="stor-hvid">Vis Projektgrupper:</td>
</tr>
<tr bgcolor="#EFF3FF">
	<td height="30" valign="middle">&nbsp;
	<select name="FM_progrupper" size="1" style="font-size : 11px;">
	<%
			strSQL = "SELECT projektgrupper.id AS id, navn FROM projektgrupper ORDER BY navn"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			if cint(request("FM_progrupper")) = cint(oRec("id")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			%>
			<option value="<%=oRec("id")%>" <%=isSelected%>><%=left(oRec("navn"), 16)%></option>
			<%
			oRec.movenext
			wend
			oRec.close
			%>
	</select></td>
	<td><input type="image" src="../ill/pil_groen_lille.gif" border="0"></td>
</tr>

<%end if%>
</form>
</table>
<%end if%>
</div>

