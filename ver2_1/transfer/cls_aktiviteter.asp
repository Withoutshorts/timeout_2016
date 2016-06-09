<%




sub lastFaseSumSub

    '**lastFaseSum **'
	    strAktListe = strAktListe & "<tr bgcolor=""#FFFFFF""><td colspan=3 style='padding-left:5px; padding-bottom:2px;'><b>"& replace(lastFase, "_", " ") &"</b> ialt: (kun viste akt.)</td>"_
	    &"<td style='padding-bottom:2px;'><span style='width:40px; font-size:9px; font-family:arial; border:1px #FFc0CB solid; padding:2px;' id='sltimer_"&lcase(lastFase)&"'><b>"& formatnumber(lastFaseTimer, 2) &"</b></span></td>"_
	    &"<td colspan=3><img src=""ill/blank.gif"" width=1 height=1 border=""0""></td>"_
	    &"<td style='padding-bottom:2px;'>= <span style='width:60px; font-size:9px; font-family:arial; border:1px #FFc0CB solid; padding:2px;' id='slsum_"&lcase(lastFase)&"'><b>"& formatnumber(lastFaseSum, 2) &"</b></span></td>"_
	    &"<td><img src=""ill/blank.gif"" width=1 height=1 border=""0""></td>"_
	    &"<td><img src=""ill/blank.gif"" width=1 height=1 border=""0""></td>"_
	    &"<td><img src=""ill/blank.gif"" width=1 height=1 border=""0""></td>"_
	    &"</tr><input id='fa_"&lcase(lastFase)&"' value='"&fa&"' type=""hidden"" />"_ 
	    &"<input id='fatot_"&lcase(lastFase)&"' value='"& faTot &"' type=""hidden"" />"_
	    &"<input id='fatot_val_"&faTot&"' value='"&formatnumber(lastFaseSum, 2) &"' type=""hidden"" />"_
	    &"<input id='fatottimer_val_"&faTot&"' value='"&formatnumber(lastFaseTimer, 2)&"' type=""hidden"" />"
	
	     fa = 0
        lastFaseTimer = 0
        lastFaseSum = 0    
	    faTot = faTot + 1
	    
end sub


sub nyFaseSub

    strAktListe = strAktListe & "<tr bgcolor=""#D6DFf5""><td style='padding-left:5px; padding-top:3px;'>fase:</td>"_
    &"<td><input id='"&lcase(trim(thisFase))&"' name='' class=""faseoskrift_navn"" value='"&replace(thisFase, "_", " ")&"' type=""text"" style='width:85px; font-size:9px; font-family:arial;' /></td>"_
	&"<td><select name=""faseoskrift"" class=""faseoskrift"" id='"&lcase(trim(thisFase))&"' style='font-size:9px; font-family:arial;'>"_
	&"<option value=""1"">Vælg..</option>"_
	&"<option value=""1"">Aktiv</option>"_
	&"<option value=""0"">Lukket</option>"_
	&"<option value=""2"">Passiv</option>"_
	&"</select></td>"_
	&"<td colspan=6>&nbsp;</td><td><input id='sl_"&lcase(trim(thisFase))&"' class=""faseoskrift_slet"" type=""checkbox"" value=""1"" /></td><td>&nbsp;</td></tr>" 

end sub



public lastFase, lastFaseSum, thisFase, lastFaseTimer, fa, faTot 
public strAktListe, aktbudgetsamlet, akttfaktimtildelt
function hentaktiviterListe(jobid, func, vispasluk, sort)

    samletverdi = 0
    a = 0
	
	'select case sort
	'case "slutdato"
	'orderBy = "aktslutdato, a.navn"
	'case else
	orderBy = "a.fase, a.sortorder, a.navn"
	'end select
	
	lastFase = ""
	lastFaseSum = 0
	
	'if vispasluk = 1 then
	vispaslukKri = " AND aktstatus <> - 1"
	'else
	'vispaslukKri = " AND aktstatus = 1"
	'end if
	
	strSQL = "SELECT a.id, a.navn, job, beskrivelse, fakturerbar, aktstartdato, "_
	&" aktslutdato, budgettimer, aktstatus, aktbudget, fomr.navn AS fomr, antalstk, tidslaas, "_
	&" a.fase, a.sortorder, a.bgr, easyreg, a.aktbudgetsum "_
	&" FROM aktiviteter a LEFT JOIN fomr ON (fomr.id = a.fomr) "_
	&" WHERE job = "& jobid &" "& vispaslukKri &" ORDER BY "& orderBy 
	
	'Response.Write strSQl
	'Response.flush
	c = 0
	fa = 0
	faTot = 0
	oRec6.open strSQL, oConn, 3
	while not oRec6.EOF 
	
	
	thisFase = oRec6("fase")
	
	if lastFase <> lcase(trim(oRec6("fase"))) AND c <> 0 AND len(trim(oRec6("fase"))) <> 0 then
	
	    '**lastFaseSum **'
	   call lastFaseSumSub
	     
       call nyFaseSub
    end if
	
	if c = 0 then
	
	
            strAktListe = strAktListe &"<table cellspacing=0 cellpadding=0 border=0 id='incidentlist'>"_
            &"<tr><td><b>Navn</b></td>"_
            &"<td><b>Fase</b></td>"_
            &"<td><b>Status</b></td>"_
            &"<td><b>Timer</b></td>"_
            &"<td><b>Stk.</b></td>"_
            &"<td><b>Grundlag</b></td>"_
            &"<td><b>Pr. Stk. / Time</b></td>"_
            &"<td style=""padding-left:10px;""><b>Pris i alt DKK</b></td>"_
            &"<td>&nbsp;</td>"_
            &"<td><b>Slet</b></td>"_
            &"<td>&nbsp;</td>"_
            &"<tr>"
    
                if len(trim(oRec6("fase"))) <> 0 then
    
                call nyFaseSub
    
	            end if
	
	end if


    if vispasluk = 1 OR (vispasluk <> 1 AND oRec6("aktstatus") = 1) then
	
	select case right(c, 1)
    case 0,2,4,6,8
    bgcolor = "#EFf3FF"
    case else
    bgcolor = "#FFFFFF"
    end select
    
    bgrSEL0 = "SELECTED"
    bgrSEL1 = ""
    bgrSEL2 = ""
    
    '** budget grundlag **'
    select case oRec6("bgr")
    case 0 '** Intent angivet
    bgrSEL0 = "SELECTED"
    'aktsumtot = oRec6("aktbudget")
    case 1 '** timer
    bgrSEL1 = "SELECTED"
    'aktsumtot = oRec6("aktbudget") * budgettimer 
    case 2 '** Stk
    bgrSEL2 = "SELECTED"
    'aktsumtot = oRec6("aktbudget") * oRec6("antalstk")
    end select
    
    
    
    strAktListe = strAktListe & "<tr bgcolor="&bgcolor&">"_
    &"<td style='width:180px;'>"
    
    strAktListe = strAktListe & "<input type=""hidden"" name=""SortOrder"" value='"&oRec6("sortorder")&"' />"
	strAktListe = strAktListe & "<input type=""hidden"" name=""rowId"" value='"&oRec6("id")&"' />"
    
    strAktListe = strAktListe &"<input id=""Hidden1"" type=""text"" name=""FM_aktnavn"" value='"& oRec6("navn") &"' style='width:180px; font-size:9px; font-family:arial;' />"
    '&"<span style=""font-size:3px; font-family:arial; color:#ffffff;""><b>"&oRec6("navn")&"</b></span>"
    strAktListe = strAktListe &"<input id=""Hidden1"" type=""hidden"" name=""FM_aktnavn"" value='#' />"
	
    
    'if len(trim(oRec6("fase"))) <> 0 then
    'strAktListe = strAktListe & " ("& oRec6("fase") &")</td>"
    'end if
    
    strAktListe = strAktListe &"<td style='width:95px;'>"_
    &"<input class='aktFase_"& trim(lcase(thisFase)) &"' id='FM_aktfas_"& oRec6("id") &"' type=""text"" name=""FM_aktfase"" value='"& lcase(trim(replace(oRec6("fase"), "_", " "))) &"' style='width:95px; font-size:9px; font-family:arial;'  />"_
    &"</td>"

    strAktListe = strAktListe &"<input class='aktFase_"& trim(lcase(thisFase)) &"' id=""FM_aktfas_h_"& oRec6("id") &""" type=""hidden"" name="""" value='"& lcase(trim(oRec6("fase"))) &"' />"
    
	strAktListe = strAktListe &"<input id=""Hidden1"" type=""hidden"" name=""FM_aktfase"" value='#' />"
	strAktListe = strAktListe &"<input id=""Hidden1"" type=""hidden"" name=""FM_aktid"" value='"& oRec6("id") &"' />"
	
	strAktListe = strAktListe &"<td>"
	 
	select case oRec6("aktstatus")
	case 1
	stCHK0 = ""
	stCHK1 = "SELECTED"
	stCHK2 = ""
	selbgcol = "#DCF5BD"
	case 2
	stCHK0 = ""
	stCHK1 = ""
	stCHK2 = "SELECTED"
	selbgcol = "#cccccc"
	case 0
	stCHK0 = "SELECTED"
	stCHK1 = ""
	stCHK2 = ""
	selbgcol = "Crimson"
	end select
	
	strAktListe = strAktListe &"<select name=""FM_aktstatus"" id='af_"&lcase(trim(oRec6("fase")))&"_"&fa&"' style='background-color:"&selbgcol&"; font-family:arial; font-size:9px;' >"_
	&"<option value=""1"" "&stCHK1&">Aktiv</option>"_
	&"<option value=""0"" "&stCHK0&">Lukket</option>"_
	&"<option value=""2"" "&stCHK2&">Passiv</option>"_
	&"</select>"_
	&"</td>"
	
	'strAktListe = strAktListe &"<input id=""Hidden1"" type=""hidden"" name=""FM_aktstatus"" value='"& oRec6("aktstatus") &"' />"
	
	
	strAktListe = strAktListe &"<td><input class='jq_akttimerstk' id='FM_akttim_"& oRec6("id") &"' name=""FM_akttimer"" type=""text"" value='"& formatnumber(oRec6("budgettimer"), 2) &"' style='width:40px; font-size:9px; font-family:arial;' /></td>"_
    &"<td><input class='jq_akttimerstk' id='FM_aktstk_"& oRec6("id") &"' name=""FM_aktantalstk"" type=""text"" value='"& formatnumber(oRec6("antalstk"), 2) &"' style='width:40px; font-size:9px; font-family:arial;' /></td>"_
    &"<input name='af_timer_"&oRec6("id")&"' id='af_timer_"&lcase(trim(oRec6("fase")))&"_"&fa&"' type=""hidden"" value='"& formatnumber(oRec6("budgettimer"), 2) &"' />"_
    &"<input name='af_sum_"& oRec6("id") &"' id='af_sum_"&lcase(trim(oRec6("fase")))&"_"&fa&"' type=""hidden"" value='"& formatnumber(oRec6("aktbudgetsum"), 2) &"' />"_
    &"<input name='FM_sum_aid_fa_"&oRec6("id")&"' id='FM_sum_aid_fa_"&oRec6("id")&"' type=""hidden"" value='"&fa&"' />"_
    &"<td><select class='bgr' id='FM_aktbgr_"& oRec6("id") &"' name=""FM_aktbgr"" style='width:45px; font-size:9px; font-family:arial;'>"_
    &"<option value=0 "& bgrSEL0 &">Ingen</option>"_
    &"<option value=1 "& bgrSEL1 &">Timer</option>"_
    &"<option value=2 "& bgrSEL2 &">Stk.</option>"_
    &"</select></td>"_
    &"<td><input class='jq_akttimerstk' id='FM_aktpri_"& oRec6("id") &"' type=""text"" name=""FM_aktpris"" value="& formatnumber(oRec6("aktbudget"), 2) &" style='width:60px; font-size:9px; font-family:arial;' /></td>"_
    &"<td class=lille style='width:80px;'>= <input class='jq_akttotal' id='FM_akttotpris_"&oRec6("id")&"' name=""FM_akttotpris"" type=""text"" value="&formatnumber(oRec6("aktbudgetsum"), 2)&" style='width:60px; font-size:9px; font-family:arial;' />"_
    &"</td>"_
    &"<td align=right><a href='aktiv.asp?menu=job&func=red&id="&oRec6("id")&"&jobid="&jobid&"&jobnavn="&request("jobnavn")&"&rdir=job2&nomenu=1' target=""_blank"" class=vmenu>"_
    &"<img src=""../ill/ac0038-16.gif"" alt='"& oRec6("navn") &"' border=""0"" /></a>"_
    &"</td>"_
    &"<td align=right><input name='FM_slet_aid_"&c&"' id='af_sl_"&lcase(trim(oRec6("fase")))&"_"&fa&"' type=""checkbox"" value=""1"" />"_
    &"<td align=right><img src=""../ill/pile_drag.gif"" alt='Træk og sorter aktivitet' border=""0"" /></td>"_
    &"<input name=""FM_slet_aid"" type=""hidden"" value='"&c&"' />"_
    &"</tr>"
    
    end if 'aktstatus (før sum beregning) 

    lastFaseTimer = lastFaseTimer + oRec6("budgettimer")
    lastFaseSum = lastFaseSum + oRec6("aktbudgetsum") 
    lastFase = lcase(trim(oRec6("fase")))
    
    aktbudgetsamlet = aktbudgetsamlet + oRec6("aktbudgetsum") 
    akttfaktimtildelt = akttfaktimtildelt + oRec6("budgettimer")  
    
    fa = fa + 1
    c = c + 1
	
	
	oRec6.movenext
	wend
	oRec6.close
	
	'thisFase = lastFase
	
	
	call lastFaseSumSub
	    
	'akttfaktimtildelt  
    '*** Totaler ***'
    strAktListe = strAktListe & "<input id=""jq_akttottimer"" value='"& formatnumber(akttfaktimtildelt, 2) &"' type=""hidden"" />"
    strAktListe = strAktListe & "<input id=""jq_akttotsum"" value='"& formatnumber(aktbudgetsamlet, 2) &"' type=""hidden"" />"
    
    strAktListe = strAktListe & "<input id=""fatot_ialt"" value='"&faTot-1&"' type=""hidden"" />" 
    
    strAktListe = strAktListe & "</table>"
    '</td></tr>
    
   
    
    
end function   




'sub lastFaseTr
'    strAktListe = strAktListe & "<tr bgcolor=""#8CAAE6""><td colspan=9 style='padding-left:5px;'><b>"& thisFase &"</b></td></tr>" 
'end sub

'sub LastFaseTrSum

'end sub



function opdateraktliste(jobid, aktids, aktnavn, akttimer, aktantalstk, aktfaser, aktbgr, aktpris, aktstatus, akttotpris, aktslet, aktslet_aids)

        
    'Response.Write aktnavn & "<br>"    
    'Response.Write aktstatus
	'Response.end
	
	aktnavn = split(aktnavn, ", #,")
	akttimer = split(akttimer, ", ")
	aktantalstk = split(aktantalstk, ", ")
	aktpris = split(aktpris, ", ")
	aktids = split(aktids, ",")
	aktstatus = split(aktstatus, ", ")
	
	aktfaser =  split(aktfaser, ", #,")
	aktbgr = split(aktbgr, ",")
	
	'Response.Write request("FM_akt_totpris")
	'Response.end
	
	'akttotpris = split(akttotpris, ", #")
	akttotpris = split(akttotpris, ", ")
	
	
	
	aktslet = split(aktslet, ", ")
	
	'Response.Write aktSlet
	'Response.end
	
	
	
	for t = 0 to UBOUND(aktids)
	err = 0
	
	    akttimer(t) = replace(akttimer(t), ".", "")
	    akttimer(t) = replace(akttimer(t), ",", ".")
	    
	    if len(trim(akttimer(t))) <> 0 then
	    akttimer(t) = akttimer(t)
	    else
	    akttimer(t) = 0
	    end if
	    
	    call erDetInt(akttimer(t))
	    
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0
	    
	    aktantalstk(t) = replace(aktantalstk(t), ".", "")
	    aktantalstk(t) = replace(aktantalstk(t), ",", ".")
	    
	    if len(trim(aktantalstk(t))) <> 0 then
	    aktantalstk(t) = aktantalstk(t)
	    else
	    aktantalstk(t) = 0
	    end if
	    
	    call erDetInt(aktantalstk(t))
	        
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0
	    
	    aktpris(t) = replace(aktpris(t), ".", "")
	    aktpris(t) = replace(aktpris(t), ",", ".")
	    
	    if len(trim(aktpris(t))) <> 0 then
	    aktpris(t) = aktpris(t)
	    else
	    aktpris(t) = 0
	    end if
	    
	    call erDetInt(aktpris(t))
	      
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0
	      
	      'akttotpris(t) = replace(akttotpris(t), "#", "")
	      akttotpris(t) = replace(akttotpris(t), ".", "")
	      akttotpris(t) = replace(akttotpris(t), ",", ".")
	      
	     
	      'Response.Write "err:" & err & "<br>"
	      'Response.Write "int:" & isInt &"<br>"
	      
	      'Response.flush
	    aktFaser(t) = trim(aktFaser(t))
		call illChar(aktFaser(t))
		aktFaser(t) = vTxt
		
	    
	    
	    aktNavn(t) = trim(aktNavn(t))
	    aktNavn(t) = replace(aktNavn(t), "'", "")
	    
	    if cint(err) = 0 then
	    
	    aktSletval = 0
	    aktSletval = request("FM_slet_aid_"& aktSlet(t) &"")
		
		if aktSletval <> "1" then
		
		strSQL = "UPDATE aktiviteter SET navn = '"& aktNavn(t) &"', aktstatus = " & aktstatus(t) & ", "_
		&" budgettimer = "& akttimer(t) &", aktbudget = "& aktpris(t) &", antalstk = "& aktantalstk(t) &", "_
		&" fase = '"& aktFaser(t) &"', bgr = "& aktBgr(t) &", aktbudgetsum = "& akttotpris(t) &""_
		&" WHERE id = "& aktids(t)
		
		'Response.write strSQL & "<br>"
		'Response.flush
		oConn.execute(strSQL)
		else
		    
		    call delakt(aktids(t))
		
		end if
		
		'** Sync job ***'
		'if len(trim(request("opdjobv"))) <> 0 then 
		'opdjobv = 1		
		'else
		'opdjobv = 0
		'end if		
		
		
		'if cint(opdjobv) = 1 then
        '    call syncJob(jobid)
        'end if
		
		end if
	
	next




end function



function delakt(id)
	
	strSQL = "SELECT id, navn FROM aktiviteter WHERE id = "& id &"" 
	oRec5.open strSQL, oConn, 3
	while not oRec5.EOF 
		
		'*** Indsætter i delete historik ****'
	    call insertDelhist("akt", oRec5("id"), 0, oRec5("navn"), session("mid"), session("user"))
		
	oRec5.movenext
	wend
	oRec5.close
		
	
	
	oConn.execute("DELETE FROM aktiviteter WHERE id = "& id &"")
	oConn.execute("DELETE FROM timer WHERE TAktivitetId = "& id &"")
	
end function

function tilknytstamakt(a, intAktfavgp, strAktFase)
		'Response.write "her " & intAktfavgp
		'Response.flush
		                ause = a ' + 1
		                
		                '** fase må ikke indeholde mellemrum **'
		                strAktFase = trim(strAktFase)
		                call illChar(strAktFase)
		                strAktFase = vTxt
		                
						if intAktfavgp <> 0 then
						'fordeltimer = request("FM_timefordeling")
						
						'antalstamakt = 0
						
						'** antal stamaktiviteter i den valgte gruppe ****
						'strSQL0 = "select COUNT(id) AS antalstamaktiGrp FROM aktiviteter WHERE aktFavorit = "& intAktfavgp &" AND job = 0 GROUP BY job"
						'oRec2.open strSQL0, oConn, 3
						
						'if not oRec2.EOF then
						'antalstamakt = oRec2("antalstamaktiGrp")
						'end if
						'oRec2.close
						
						
						strSQL = "select id, navn, fakturerbar, beskrivelse, budgettimer, fomr, faktor, "_
						&" aktbudget, aktstatus, tidslaas, tidslaas_st, tidslaas_sl, antalstk, "_
						&" tidslaas_man, tidslaas_tir, tidslaas_ons, "_
				        &" tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son, fase, sortorder, bgr, aktbudgetsum, easyreg "_
						&" FROM aktiviteter WHERE aktFavorit = "& intAktfavgp &" AND job = 0"
						
						    oRec2.open strSQL, oConn, 3
							while not oRec2.EOF
							
							aktid = oRec2("id")
							aktNavn = trim(request("FM_stakt_navn_"& ause &"_"& oRec2("id"))) 'oRec2("navn")
							aktNavn = replace(aktNavn, "'", "")
							aktFakbar = oRec2("fakturerbar")
							aktFomr = oRec2("fomr")
							aktFaktor = replace(oRec2("faktor"), ",", ".")
							aktstatus = request("FM_stakt_status_"& ause &"_"& oRec2("id")) 'oRec2("aktstatus")
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
							'if len(trim(strAktFase)) <> 0 then
							'strFase = strAktFase
							'else
							    if len(trim(request("FM_stakt_fase_"& ause &"_"& oRec2("id")))) <> 0 then
							    strFase = request("FM_stakt_fase_"& ause &"_"& oRec2("id")) 'replace(oRec2("fase"), "'", "")
							    else
							    strFase = NULL
							    end if
							'end if
							
							if len(trim(oRec2("beskrivelse"))) <> 0 then
							
							'beskrivelse = replace(oRec2("beskrivelse"), "'", "''")
							'call htmlparseCSV(beskrivelse)
							'beskrivelse = htmlparseCSV
							
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
							
							''beskrivelse = replace(beskrivelse, "<p>", "")
							''beskrivelse = replace(beskrivelse, "</p>", "")
							
							else
							
							beskrivelse = ""
							
							end if
							
							'antalstk = replace(oRec2("antalstk"), ",", ".")
							
							
							
							
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
							
						
							
							sortorder = oRec2("sortorder")
							
							'********************************
							'** Henter fra forkalkulation ***
							'********************************
							'aktBudgettimer = replace(request("FM_stakt_timer_"& ause &"_"& oRec2("id") &""), ".", "") 'replace(oRec2("aktbudgetsum"), ",", ".")
							aktBudgettimer = replace(oRec2("budgettimer"), ",", ".")
							
							call erDetInt(aktBudgettimer)
							if isInt > 0 OR len(trim(aktBudgettimer)) = 0 then
							aktBudgettimer = 0
							else
							aktBudgettimer = aktBudgettimer
							end if
							
							
							antalstk = oRec2("antalstk") 'replace(request("FM_stakt_stk_"& ause &"_"& oRec2("id") &""), ".", "") 'replace(oRec2("aktbudgetsum"), ",", ".")
							antalstk = replace(antalstk, ",", ".")
							
							call erDetInt(antalstk)
							if isInt > 0 OR len(trim(antalstk)) = 0 then
							antalstk = 0
							else
							antalstk = antalstk
							end if
							
							bgr = oRec2("bgr") 'request("FM_stakt_bgr_"& ause &"_"& oRec2("id") &"") 'oRec2("bgr")
							
							
							'aktBudget = aktBudget 'replace(request("FM_stakt_aktsum_"& ause &"_"& oRec2("id") &""), ".", "") 'replace(oRec2("aktBudget"), ",", ".")
							'aktBudget = replace(aktBudget, ",", ".")
							aktBudget = replace(oRec2("aktbudget"), ",", ".")
							
							'Response.Write "aktBudget " & aktBudget
							
							call erDetInt(aktBudget)
							if isInt > 0 OR len(trim(aktBudget)) = 0 then
							aktBudget = 0
							else
							aktBudget = aktBudget
							end if
							
							aktbudgetsum = oRec2("aktbudgetsum") 'replace(request("FM_stakt_totaktsum_"& ause &"_"& oRec2("id") &""), ".", "") 'replace(oRec2("aktbudgetsum"), ",", ".")
							aktbudgetsum = replace(aktbudgetsum, ",", ".")
							
							call erDetInt(aktbudgetsum)
							if isInt > 0 OR len(trim(aktbudgetsum)) = 0 then
							aktbudgetsum = 0
							else
							aktbudgetsum = aktbudgetsum
							end if
							
							easyreg = oRec2("easyreg")
						    
						    '*** Tilføj fravalgt ***'
						    'Response.write "chk: "& request("FM_stakt_tilfoj_"& ause &"_"& oRec2("id") &"") & "<br>"
						    'Response.flush
						    if request("FM_stakt_tilfoj_"& ause &"_"& oRec2("id") &"") = "1" then
						    
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
							
							'**** Overfører de timepriser hver enkelt stamaktivitet er født med, for hver enkelt medarbejder ****
							'**** Eller nedarv fra job 
							'**** 1: Behold medarbejdertimepriser på aktiviteter   
							'**** 0: Nedarv fra job
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
							
							
							end if '** tilføj fravalgt
							
							oRec2.movenext
							wend
						oRec2.close
						end if
		
		end function

    
%>