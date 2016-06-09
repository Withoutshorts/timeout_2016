<%
function stamakt(useCookie,forvalgt,fsnavn,a)
'** fsnavn skal altid starte blank **'
'** Hvis nu folk ikke ønsker at indele i faser **'
fsnavn = ""
%>
<tr>
		<td>&nbsp;</td>
		<td colspan="2" style="padding-left:35; padding-top:10px; padding-bottom:4px;" valign="top">
		<% 
			strSQL = "SELECT ag.id, ag.navn, COUNT(a.id) AS antalakt FROM akt_gruppe ag "_
			&" LEFT JOIN aktiviteter a ON (a.aktfavorit = ag.id ) "_
			&" WHERE ag.id <> 2 GROUP BY ag.id ORDER BY ag.navn"
			
			
			'Response.Write strSQL
			'Response.end
			
		'Response.Write "lastStamaktgrp" & forvalgt	
	
		%>
		<select class="selstaktgrp" id="<%=a%>" name="FM_favorit" style="font-size:9px; width:250;">
		<option value="0">Ingen</option>
					
			<%
			
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
				
			
			lastStamaktgrp = forvalgt
			
			
			
			if cint(oRec("id")) = cint(lastStamaktgrp) then
			thisStaktgpSEL = "SELECTED"
			else
			thisStaktgpSEL = "" 
			end if
			
			
			if cint(oRec("antalakt")) < 10 then
			antalAkt = " "& oRec("antalakt")
			else
			antalAkt = oRec("antalakt")
			end if
			%>
			<option value="<%=oRec("id")%>" <%=thisStaktgpSEL%>><%=oRec("navn") & " ("& antalAkt &" stk.)" %></option>
			<%
			st = 1
			oRec.movenext
			wend
			oRec.close%>
			</select>
			&nbsp;&nbsp;&nbsp;Fase: <input id="stgrpfs_<%=a%>" name="FM_favorit_fase" value="<%=fsnavn%>" type="text" style="font-size:9px; width:200;" /> <span id="fs_ryd_<%=a %>" style="color:#999999; font-size:9px; cursor:hand;"><u>Hent navn</u></span>
				</td>
				<td>&nbsp;</td>
			</tr>
<%
end function


sub top

    oimg = "ikon_job_48.png"
	oleft = 20
	otop = 55
	owdt = 100
	oskrift = "Job"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)

end sub


sub kundeopl

'******************* Kunde, og kunderelateret oplysninger  *********************************'
	if showaspopup <> "y" then
	kwidth = 480
	else
	kwidth = 310
	end if
	
	'if func = "red" then
	'useKbg = "#eff3ff"
	'else
	'useKbg = "#ffffff"
	'end if
	
	if (func = "opret" AND seraft = 1) OR rdir = "sdsk" then
	strKnr = Request("kid") '** ??
	else
	'serchk = "checked"
	end if
	
	if len(trim(strKnr)) <> 0 then
	strKnr = strKnr
	else
	strKnr = 0
	end if
	
	
	if func = "opret" AND step = 1 then%>
	<tr>
		<td colspan=4 height="20" style="padding:10px 10px 10px 10px;">
		<b>Vælg Kontakt:</b>&nbsp;&nbsp;(<a href="kunder.asp?func=opret&ketype=k&hidemenu=1&showdiv=onthefly&rdir=1" target="_blank" class=vmenu>Opret ny her..</a>)</td>
	</tr>
	<%else %>
	  <tr>
	  <td colspan=4 style="padding:0px 20px 2px 10px;"><b>Kontaktoplysninger</b></td>
          </tr>
	<%end if%>
	
	
	<tr>
		
		<td style="padding:10px 10px 10px 10px;" colspan=4>
		
		
		<img src="../ill/ikon_kunder_24.png" alt="" border="0">&nbsp;<b>Kontakt:</b> 
		<%if func <> "red" then %>
		(Vælg flere hvis det samme job skal oprettes på flere kunder)
		<%end if %>
		<br>
		
		    <%if func = "opret" AND step = 1 then %>
            <input name="FM_kunde" id="FM_kunde" value=0 type="hidden" />
			<select name="FM_kunde" id="FM_kunde" size=10 multiple style="font-size:9px; width:<%=kwidth%>px;">
			<%else %>
			
			<%if func = "opret" AND step = 2 then 
			dsab = "DISABLED"
			%>
			<!--<input name="FM_kunde" id="Hidden1" value="<%=strKundeId %>" type="hidden" />-->
			<%
			else
			dsab = ""
			end if%>
			<select name="FM_kunde" id="FM_kunde" <%=dsab %> size=1 style="font-size:9px; width:<%=kwidth%>px;">
			
			<%end if 
			
			
			strSQL = "SELECT Kkundenavn, Kkundenr, Kid, kundeans1, kundeans2 FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
			oRec.open strSQL, oConn, 3
			kans1 = ""
			kans2 = ""
			while not oRec.EOF
				
				if func = "opret" AND step = 2 then
				    if cint(strKundeId) = oRec("Kid") then
				    kSel = "SELECTED"
				    else
				    kSel = ""
				    end if
				else
				    if cint(strKnr) = oRec("Kid") then
				    kSel = "SELECTED"
				    else
				    kSel = ""
				    end if
				end if
				
				strSQL2 = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = "&oRec("kundeans1")
				oRec2.open strSQL2, oConn, 3 
				if not oRec2.EOF then
				kans1 = oRec2("mnavn")
				end if
				oRec2.close
				
				strSQL2 = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = "&oRec("kundeans2")
				oRec2.open strSQL2, oConn, 3 
				if not oRec2.EOF then
				kans2 = " - &nbsp;&nbsp;" & oRec2("mnavn") 
				end if
				oRec2.close
				
				if len(kans1) <> 0 OR len(kans2) <> 0 then
				anstxt = "kontaktansv 1: "
				else
				anstxt = ""
				end if
				
			%>
			<option value="<%=oRec("Kid")%>" <%=kSel%>><%=oRec("Kkundenavn")%>&nbsp;&nbsp;(<%=oRec("Kkundenr")%>)&nbsp;&nbsp;-&nbsp;&nbsp;<%=anstxt%><%=kans1%>&nbsp;&nbsp;<%=kans2%></option>
			<%
			kans1 = ""
			kans2 = ""
			oRec.movenext
			wend
			oRec.close
			%>
		</select>
		<%if func = "red" AND strInternt <> "0" then%>
		<br><font class=megetlillesort>NB: Hvis du skifter kontakt skal du være opmærksom på at evt. valgte aftaler og kontaktpersoner
		også skal ændres.</font>
		<%end if%>
		
		
		</td>
		
	</tr>
	
	
	
	
	
	
	
	
	<%if func = "opret" AND step = 2 OR func = "red" then %>
	<tr>
		
		<td colspan="4" style="padding:10px 10px 10px 10px; border-bottom:1px silver solid;"><input type="checkbox" name="FM_kundese" id="FM_kundese" value="1" <%=kundechk%>>&nbsp;<b>Gør job tilgængeligt for kontakt.</b><br>
		Hvis tilgængelig for kontakt tilvælges, udsendes der en mail til kontakt-stamdata emailadressen, med link til kontakt loginside.
		
		
		<%if func = "opret" then
		hvchk1 = "checked"
		hvchk2 = ""
		else
			if cint(intkundeok) = 2 then
			hvchk1 = ""
			hvchk2 = "checked"
			else
			hvchk1 = "checked"
			hvchk2 = ""
			end if
		end if%>
		<br><br>
		<b>Hvis job åbnes for kontakt, hvornår skal registrerede timer så være tilgængelige?</b><br>
		<input type="radio" name="FM_kundese_hv" value="0" <%=hvchk1%>>Offentliggør timer, så snart de er indtastet.<br>
		<input type="radio" name="FM_kundese_hv" value="1" <%=hvchk2%>>Offentliggør først timer når jobbet er lukket. (afsluttet/godkendt)
		<br>&nbsp;</td>
		
	</tr>
	<%end if%>
	
	<%
	

end sub






sub kundeopl_aft_kpers 
    if showaspopup <> "y" then
	kwidth = 480
	else
	kwidth = 310
	end if
	%>
	<tr>
		<td colspan="4" style="padding:10px 10px 10px 10px;"><b>Kontaktperson hos kontakt:</b> (kunde)<br>
		
		<%
		
		strSQLkpers = "SELECT k.navn, k.id AS kid FROM kontaktpers k WHERE kundeid = "& strKundeId &" ORDER BY k.navn"
		'Response.Write strSQLkpers
		
		%>
		
		<select name="FM_opr_kpers" id="FM_opr_kpers" style="font-size:9px; width:<%=kwidth%>px;">
		<option value="0">Ingen</option>
	
		<%
		
		oRec.open strSQLkpers, oConn, 3 
		while not oRec.EOF 
		
		
		if cint(intkundekpers) = cint(oRec("kid")) then
		ts3 = "SELECTED"
		else
		ts3 = ""
		end if
		%>
		<option value="<%=oRec("kid")%>" <%=ts3%>><%=left(oRec("navn"), 30)%></option>
		<%
		oRec.movenext
		wend
		oRec.close 
		%>
		</select><br>&nbsp;</td>
	
	</tr>
	<tr>
		<td colspan="4" style="padding:10px 10px 10px 10px;">
		<b>Tilknyt job til aftale?</b><!--<input type="checkbox" name="FM_serviceaft" value="1" <=serchk%>>--><br>
		
		<%
		strSQL2 = "SELECT id, enheder, stdato, sldato, status, navn, pris, perafg, "_
		&" advitype, advihvor, erfornyet, varenr, aftalenr FROM serviceaft "_
		&" WHERE kundeid = "& strKundeId &" OR id = "& intServiceaft &" ORDER BY id DESC" 'AND status = 1
		
		'Response.write strSQL2
		%>
		<select name="FM_serviceaft" id="FM_serviceaft" style="font-size:9px; width:<%=kwidth%>px;">
		<option value="0">Nej (fjern)</option>
		
		<%
		
		oRec2.open strSQL2, oConn, 3 
		while not oRec2.EOF 
		
		if oRec2("advitype") <> 0 then
		udldato = "&nbsp;&nbsp;&nbsp; startdato: " & formatdatetime(oRec2("stdato"), 2) 
		else
		udldato = ""
		end if%>
		
		<%
		if cint(intServiceaft) = cint(oRec2("id")) then
		serChk = "SELECTED"
		else
		serChk = ""
		end if
		
		if oRec2("status") <> 0 then
		stThis = "Aktiv"
		else
		stThis = "Lukket"
		end if%>
		<option value="<%=oRec2("id")%>" <%=serChk%>> <%=oRec2("navn")%>&nbsp;(<%=oRec2("aftalenr") %>)  <%=udldato%> (<%=stThis%>) </option>
		<%
		oRec2.movenext
		wend
		oRec2.close
		%>
		</select>
		
		<%if func = "red" then %><br>
		<input type="checkbox" name="FM_overforGamleTimereg" value="1"> Overfør / fjern eksisterende 
		<b>timeregistreringer</b> og <b>materiale-forbrug</b> (på dette job) til den valgte aftale.<br />
		<br />Hvis der <b>ændres</b> aftale tilknytning undervejs i jobbets levetid, vil evt. oprettede
		fakturaer på jobbet ikke længere kunne ses under den <b>gamle aftale.</b> (faktura historik, aftale afstemning)<br />
		Fakturaer oprettet direkte på aftalen vil ikke blive berørt. 
		<%end if%>
		  <br>&nbsp;</td>
		
	</tr>
	
<%
end sub


               
               
               function tilknytstamakt(intAktfavgp, strAktFase)
				'Response.write "her " & intAktfavgp
				'Response.flush
				                
				                '** fase må ikke indeholde mellemrum **'
				                strAktFase = trim(strAktFase)
				                call illChar(strAktFase)
				                strAktFase = vTxt
				                
								if intAktfavgp <> 0 then
								fordeltimer = request("FM_timefordeling")
								antalstamakt = 0
								
								'** antal stamaktiviteter i den valgte gruppe ****
								strSQL0 = "select id FROM aktiviteter WHERE aktFavorit = "& intAktfavgp &""
								oRec2.open strSQL0, oConn, 3
								
								while not oRec2.EOF
								antalstamakt = antalstamakt + 1
								oRec2.movenext
								wend
								oRec2.close
								
								
								strSQL = "select id, navn, fakturerbar, beskrivelse, budgettimer, fomr, faktor, "_
								&" aktbudget, aktstatus, tidslaas, tidslaas_st, tidslaas_sl, antalstk, "_
								&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
						        &" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, fase, sortorder, bgr, aktbudgetsum, easyreg "_
								&" FROM aktiviteter WHERE aktFavorit = "& intAktfavgp &""
								
								    oRec2.open strSQL, oConn, 3
									while not oRec2.EOF
									
									aktid = oRec2("id")
									aktNavn = oRec2("navn")
									aktFakbar = oRec2("fakturerbar")
									aktFomr = oRec2("fomr")
									aktFaktor = replace(oRec2("faktor"), ",", ".")
									aktBudget = replace(oRec2("aktbudget"), ",", ".")
									aktstatus = oRec2("aktstatus")
									tidslaas = oRec2("tidslaas")
									tidslaas_st = oRec2("tidslaas_st")
									tidslaas_sl = oRec2("tidslaas_sl")
									
									tidslaas_man = oRec2("tidslaas_man")
									tidslaas_tir = oRec2("tidslaas_tir")
									tidslaas_ons = oRec2("tidslaas_ons")
									tidslaas_tor = oRec2("tidslaas_tor")
									tidslaas_fre = oRec2("tidslaas_fre")
									tidslaas_lor = oRec2("tidslaas_lor")
									tidslaas_son = oRec2("tidslaas_son")
									
									'** behold eller overskriv fase **'
									if len(trim(strAktFase)) <> 0 then
									strFase = strAktFase
									else
									    if len(trim(oRec2("fase"))) <> 0 then
									    strFase = replace(oRec2("fase"), "'", "")
									    else
									    strFase = NULL
									    end if
									end if
									
									if len(trim(oRec2("beskrivelse"))) <> 0 then
									
									beskrivelse = replace(oRec2("beskrivelse"), "'", "''")
									beskrivelse = replace(beskrivelse, "<span", "")
									beskrivelse = replace(beskrivelse, "</span>", "")
									beskrivelse = replace(beskrivelse, "<div", "")
									beskrivelse = replace(beskrivelse, "</div>", "")
									beskrivelse = replace(beskrivelse, "<table", "")
									beskrivelse = replace(beskrivelse, "</table>", "")
									beskrivelse = replace(beskrivelse, "<tr", "")
									beskrivelse = replace(beskrivelse, "</tr>", "")
									beskrivelse = replace(beskrivelse, "<td", "")
									beskrivelse = replace(beskrivelse, "</td>", "")
									'beskrivelse = replace(beskrivelse, "<p>", "")
									'beskrivelse = replace(beskrivelse, "</p>", "")
									
									else
									
									beskrivelse = ""
									
									end if
									
									antalstk = replace(oRec2("antalstk"), ",", ".")
									
									
									
									
										if len(tidslaas) <> 0 then
										tidslaas = tidslaas
										else
										tidslaas = 0
										end if
										
										
										if len(trim(tidslaas_st)) <> 0 AND len(trim(tidslaas_sl)) <> 0 then
										tidslaas_st = left(formatdatetime(tidslaas_st, 3), 5)
										tidslaas_sl = left(formatdatetime(tidslaas_sl, 3), 5)
										else
										tidslaas_st = "07:00:00"
										tidslaas_sl = "23:30:00"
										end if
			
									
									
									if len(aktFaktor) <> 0 then
									aktFaktor = aktFaktor
									else
									aktFaktor = 0
									end if
									
									select case fordeltimer
									case 1
									aktBudgettimer = SQLBless(oRec2("budgettimer"))
									case 2
										if antalstamakt > 0 then
										timerpaajob = cint(SQLBless3(strBudgettimer)) + cint(SQLBless3(ikkeBudgettimer))
										aktBudgettimer = cint((timerpaajob/antalstamakt))
										else
										aktBudgettimer = 0
										end if
									case 3
									aktBudgettimer = 0
									end select
									
									sortorder = oRec2("sortorder")
									bgr = oRec2("bgr")
									aktbudgetsum = replace(oRec2("aktbudgetsum"), ",", ".")
									
									easyreg = oRec2("easyreg")
								
									strSQLins = "INSERT INTO aktiviteter "_
									&" (navn, dato, editor, job, fakturerbar, "_
									&" projektgruppe1, projektgruppe2, projektgruppe3, "_
									&" projektgruppe4, projektgruppe5, projektgruppe6, "_
									&" projektgruppe7, projektgruppe8, projektgruppe9, "_
									&" projektgruppe10, aktstartdato, aktslutdato, "_
									&" budgettimer, fomr, faktor, aktBudget, aktstatus, tidslaas, "_
									&" tidslaas_st, tidslaas_sl, beskrivelse, antalstk, "_
									&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
						            &" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, fase, sortorder, bgr, aktbudgetsum, easyreg "_
						            &" ) VALUES "_
									&"('"& aktNavn &"', "_
									&"'"& strDato &"', "_ 
									&"'"& strEditor &"', "_
									&""& varjobId  &", "_ 
									&""& aktFakbar &", "_
									&""& strProjektgr1 &", "_ 
									&""& strProjektgr2 &", "_ 
									&""& strProjektgr3 &", "_ 
									&""& strProjektgr4 &", "_ 
									&""& strProjektgr5 &", "_
									&""& strProjektgr6 &", "_ 
									&""& strProjektgr7 &", "_ 
									&""& strProjektgr8 &", "_ 
									&""& strProjektgr9 &", "_ 
									&""& strProjektgr10 &", "_     
									&"'"& startDato &"', "_ 
									&"'"& slutDato &"', "_
									&""& aktBudgettimer & ", "& aktFomr &", "_
									&""& aktFaktor &", "& aktBudget &", "& aktstatus &", "_
									&""&tidslaas&", '"&tidslaas_st&"', '"&tidslaas_sl&"', '"& beskrivelse &"', "& antalstk &","_
									&""& tidslaas_man &", "& tidslaas_tir &", "& tidslaas_ons &", "_
						            &""& tidslaas_tor &", "& tidslaas_fre &", "& tidslaas_lor &", "& tidslaas_son &", "_
						            &" '"& strFase &"', "& sortorder &", "& bgr &", "& aktbudgetsum &", "& easyreg &""_
						            &")"
									
									
									'Response.write strSQLins
									'Response.flush
									oConn.execute(strSQLins)
									
									'*** Henter det netop oprettede akt-id ***
									strSQLid = "SELECT id FROM aktiviteter ORDER BY id DESC"
									oRec3.open strSQLid, oConn, 3
									if not oRec3.EOF then
									useNewAktid = oRec3("id")
									end if
									oRec3.close
									
									if len(useNewAktid) <> 0 then
									useNewAktid = useNewAktid
									else
									useNewAktid = 0
									end if
									
									'**** Overfører de timepriser hver enkelt st.akt er født med, for hver enkelt medarbejder ****
									if request("FM_timepriser") = 1 then
										
										strSQLfindtp = "SELECT jobid, aktid, medarbid, timeprisalt, 6timepris FROM timepriser WHERE aktid = "& aktid
										oRec3.open strSQLfindtp, oConn, 3 
										while not oRec3.EOF 
											
											strSQLtp = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris) VALUES "_
											&" ("& varjobId &", "& useNewAktid &", "& oRec3("medarbid") &", "_
											&" "& replace(oRec3("timeprisalt"), ",", ".") &", "& replace(oRec3("6timepris"), ",", ".") &")"
											oConn.execute(strSQLtp)
											
											'Response.write strSQLtp
											
										oRec3.movenext
										wend
										oRec3.close  
										
												
									end if
									
									oRec2.movenext
									wend
								oRec2.close
								end if
				
				end function
				
				
				
				
			
				
				
				
				
				
                function pjgcase(grp, thisprg)
				
				if cint(grp) = cint(strProjektgr1) OR cint(grp) = cint(strProjektgr2) OR cint(grp) = cint(strProjektgr3) OR cint(grp) = cint(strProjektgr4) OR cint(grp) = cint(strProjektgr5) OR cint(grp) = cint(strProjektgr6) OR cint(grp) = cint(strProjektgr7) OR cint(grp) = cint(strProjektgr8) OR cint(grp) = cint(strProjektgr9) OR cint(grp) = cint(strProjektgr10) Then
				else
					
					strSQL3 = "UPDATE aktiviteter SET projektgruppe"&thisprg&" = 1 WHERE id = " & oRec5("id")
					'Response.Write strSQL3
					'Response.flush
					
					oConn.execute(strSQL3)
				
				end if
				end function

%>
