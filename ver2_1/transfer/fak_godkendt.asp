<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/functions_inc.asp"-->  
<%

a = 0
dim aval
redim aval(a)

'** Bruges til PDF visning ***
if len(request("nosession")) <> 0 then
nosession = request("nosession")
else
nosession = 0
end if

thisfile = "fak_godkendt.asp"
 

if len(session("user")) = 0 AND nosession = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	if len(request("print")) <> 0 then
	print = request("print")
	    
	    %>
	    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	    <%
	    
	else
	print = "n"
	
	    
	    %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	  
	    <%
	    
	end if
	
	%>
	<!--#include file="inc/fak_layout_inc.asp"-->
	<%
	
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	id = request("id")
	showmoms = 0
	
	if len(request("jobid")) <> 0 then
	jobid = request("jobid")
	else
	jobid = 0
	end if
	
	if len(request("aftid")) <> 0 then
	aftid = request("aftid")
	else
	aftid = 0
	end if
	
	'*** aftale side variable ****
	
	'if len(request("aftaleid")) <> 0 then
	'aftid = request("aftaleid")
	'else
	'aftid = 0
	'end if
	
	'if len(request("FM_aftnr")) <> 0 then
	'sogaftnr = request("FM_aftnr")
	'else
	'sogaftnr = 0
	'end if
	
	'if len(request("thiskid")) <> 0 then
	'thiskid = request("thiskid")
	'else
	'thiskid = 0
	'end if
	
	
	visjoblog = 0
	vismatlog = 0
	
	
	
	
	'*** Faktura info ****
	if jobid <> 0 then
	strSQL_sel = " jobid, jobnavn, jobnr,"
	strSQL_lftJ = " job ON (id = jobid) "
	else
	strSQL_sel = " s.navn AS aftnavn, s.aftalenr, "
	strSQL_lftJ = " serviceaft s ON (s.id = aftaleid) "
	end if
	
	strSQL = "SELECT f.editor, tidspunkt, fid, faknr, fakdato, timer, beloeb, kommentar, "_
	&" tidspunkt, fakadr, att, faktype, konto, modkonto, "& strSQL_sel &" "_
	&" vismodland, vismodatt, vismodtlf, vismodcvr, visafstlf, visafsemail,"_
	&" visafsswift, visafsiban, visafscvr, moms, enhedsang, f.varenr, jobbesk, "_
	&" visjoblog, visrabatkol, vismatlog, rabat, "_
	&" visjoblog_timepris, visjoblog_enheder "_
	&" FROM fakturaer f LEFT JOIN "& strSQL_lftJ &" WHERE fid = "& id
	
	'Response.write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	strEditor = oRec("editor")
	fakTidspunkt = oRec("tidspunkt")
	intFaktype = oRec("faktype")
	intKundeid = oRec("fakadr")
	intAtt = oRec("att")
	intKonto = oRec("konto")
	intModKonto = oRec("modkonto")
	intFaktureretTimer = oRec("timer")
	intFaktureretBelob = oRec("beloeb")
	visjoblog = oRec("visjoblog")
	
	visjoblog_enheder = oRec("visjoblog_enheder")
	visjoblog_timepris = oRec("visjoblog_timepris")
	
	vismatlog = oRec("vismatlog")
	
	visrabatkol = oRec("visrabatkol")
	
	if intFaktype = 1 then
	intFaktureretBelob = -(intFaktureretBelob)
	end if
	
	if jobid <> 0 then
	strKomm = oRec("kommentar")
	strJobnavn = oRec("jobnavn")
	intJobnr = oRec("jobnr")
	strJobBesk = oRec("jobbesk")
	else
	
	
	 strAftNavn = oRec("aftnavn")
	 strAftVarenr = oRec("varenr")
	 intAftnr = oRec("aftalenr")
	    
	    if cdate(oRec("fakdato")) > cdate("1/5/2007") then
	    strKomm = oRec("kommentar")
	    strJobBesk = oRec("jobbesk")
	    else
	    strJobBesk = oRec("kommentar")
	    strKomm = "Netto Kontant 8 dage."
	    end if
	
	end if
	
	fakdato = oRec("fakdato") 
	varFaknr = oRec("faknr")
	
	
	intVismodland = oRec("vismodland")
	intVismodatt = oRec("vismodatt")
	intVismodtlf = oRec("vismodtlf")
	intVismodcvr = oRec("vismodcvr")
	
	
	intVisafstlf = oRec("visafstlf")
	intVisafsemail = oRec("visafsemail")
	intVisafsswift = oRec("visafsswift")
	intVisafsiban = oRec("visafsiban")
	intVisafscvr = oRec("visafscvr")
	
	intMoms = oRec("moms")
	if intFaktype = 1 then
	intMoms = -(intMoms)
	end if
	
	select case oRec("enhedsang")
	case 1	
	enhed = "Stk. pris"
	case 2
	enhed = "Enhedspris"
	case else
	enhed = "Timepris"
	end select
	
	
	
	intAftRabat = oRec("rabat")
	
	end if
	oRec.close 
	
	
	
	
	
	
	 
	
	' *** Modtager ****
	strSQL = "SELECT Kid, kkundenavn, kkundenr, adresse, postnr, city, land, telefon, cvr FROM kunder WHERE Kid =" & intKundeid		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			intKnr = oRec("kkundenr")
			strKnavn = oRec("kkundenavn")
			strKadr = oRec("adresse")
			strKpostnr = oRec("postnr")
			strBy = oRec("city")
			strLand = oRec("land")
			intKid = oRec("Kid")
			intCVR = oRec("cvr")
			intTlf = oRec("telefon")
		end if
		
		oRec.close
		
    kid = intKid
		
	'*** Att ******
	select case intAtt
	case 991
	showAtt = "Den økonomi ansvarlige"
	case 992
	showAtt = "Regnskabs afd."
	case 993
	showAtt = "Administrationen"
	case else
	
	strSQL3 = "SELECT id, navn FROM kontaktpers WHERE id="& intAtt
	oRec3.open strSQL3, oConn, 3
	if not oRec3.EOF then
	showAtt = oRec3("navn")
	end if
	oRec3.close
	
	if len(showAtt) <> 0 then
	showAtt = showAtt
	else
	showAtt = ""
	end if
	
	end select	
	
	
	'*** Afsender ***
	
		strSQL = "SELECT adresse, postnr, city, land, telefon, email, regnr, kkundenavn, kontonr, cvr, bank, swift, iban, kid FROM kunder WHERE useasfak = 1"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			yourbank = oRec("bank")
			yourRegnr = oRec("regnr")
			yourKontonr = oRec("kontonr")
			yourCVR = oRec("cvr")
			yourNavn = oRec("kkundenavn")
			yourAdr = oRec("adresse")
			yourPostnr = oRec("postnr")
			yourCity = oRec("city")
			yourLand = oRec("land")
			yourEmail = oRec("email")
			yourTlf = oRec("telefon")
			yourSwift = oRec("swift")
			yourIban = oRec("iban")
			afskid = oRec("kid")
		end if
		oRec.close
		
		
		
		'**** Konto / Modkonto ***
		
		'** modtagers konto **
		if len(intKonto) <> 0 then
		intKonto = intKonto
		else
		intKonto = 0
		end if 
		
		
		strKontoNavn = "-"
		intMomskode = "-"
		intKvotient = "-"
		intKontonr = "0"
		
		strSQL2 = "SELECT momskode, kvotient, kontoplan.navn, kontonr FROM kontoplan LEFT JOIN momskoder ON (momskoder.id = momskode) WHERE kontoplan.kontonr = "& intKonto
		
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
		strKontoNavn = oRec2("navn")
		intMomskode = oRec2("momskode")
		
		select case oRec2("kvotient")
		case 5 
		intKvotient = "0,25 %"
		case 21
		intKvotient = "0,5 %"
		case 1
		intKvotient = "0"
		end select
		
		intKontonr = oRec2("kontonr")
		end if
		oRec2.close
		
		
		
		'** afsenders konto **
		if len(intModKonto) <> 0 then
		intModKonto = intModKonto
		else
		intModKonto = 0
		end if 
		
		strModKontoNavn = "-"
		intModKontonr = "0"
		
		strSQL2 = "SELECT kontoplan.navn, kontonr FROM kontoplan WHERE kontoplan.kontonr = "& intModKonto
		oRec2.open strSQL2, oConn, 3 
		if not oRec2.EOF then
		strModKontoNavn = oRec2("navn")
		intModKontonr = oRec2("kontonr")
		end if
		oRec2.close
		
		
		'*********************************************************************************
		'**** Fakturaen ***
		'*********************************************************************************
		%>
		<div id="sindhold" style="position:absolute; left:0px; top:0px; visibility:visible;">

		
		<%if print = "n" then
		bd = 1
		else
		bd = 0
		end if %>
		<div id="fakturaside" style="position:relative; left:10px; top:4px; width:640px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; padding:10px 10px 10px 10px;">
		
		<% 
		'*** fak header layout *******
		call fakheader() 
	
		'*** Maintable layout ********
		call maintable
		
		'*** Modtager boks layout ****
		call modtager_layout()
		
		'*** Maintable 2 layout ********
		call maintable_2
		
		'*** Afsender boks layout ****										
		call afsender_layout()
		
		'*** Maintable 3 layout ********
		call maintable_3
		
												
		'*** Vedr. (upspecificering top) ***
		call vedr
		
		
		
		
		
		
		'**** Udspecificering ******
		strFakdet = ""
		strMedarbFakdet = ""
		strFakmat = ""
		
		if jobid <> 0 then
					'*** Fakturering af job /Aktiviteter ***
					h = 100
					
					strSQL = "SELECT antal, beskrivelse, aktpris, fakid, "_
					&" enhedspris, aktid, showonfak, rabat, enhedsang FROM faktura_det WHERE fakid = "& id &" AND showonfak = 1 ORDER BY beskrivelse"
					
					oRec.open strSQL, oConn, 3 
					while not oRec.EOF 
						
						aktEhpris = oRec("enhedspris")
						aktPris = oRec("aktpris")
						if intFaktype = 1 then
						aktPris = -(aktPris)
						aktEhpris = -(aktEhpris)
						end if
						
						rabat = oRec("rabat")
						
						select case oRec("enhedsang")
						case 0
						enhedsang = "pr. time"
						case 1
						enhedsang = "pr. stk."
						case 2
						enhedsang = "pr. enhed"
						end select
						
					
					strFakdet = strFakdet &"<tr>"_
				    &"<td valign=top align=right style='padding-right:2;'>"& formatnumber(oRec("antal"), 2) &"</td>"_
					&"<td colspan=2 valign=top style='padding-left:5;'>" & oRec("beskrivelse") &"</td>"_
					&"<td valign=top align=right style='padding-right:10;'>" & formatcurrency(aktEhpris, 2) &" "& enhedsang &"</td>"
					
					
					if cint(visrabatkol) = 1 then
					strFakdet = strFakdet &"<td valign=top align=right style='padding-right:10;'>"
					        
					        if cdbl(rabat) <> 0 then
					        rbtthis = (rabat * 100) 
					        strFakdet = strFakdet & rbtthis &" %</td>"
					        else
					        strFakdet = strFakdet &"&nbsp;</td>"
					        end if
					        
					end if
					
					
					
					strFakdet = strFakdet &"<td valign='top' align=right>"& formatcurrency(aktPris) &"</td>"_
					&"</tr>"
					
					aktsubtotal = aktsubtotal + aktPris
					
					h = h + 20		
			        rbtthis = 0
					
					
					
						
								
								ehpris = replace(oRec("enhedspris"), ",", ".")
								
								'**** Medarbejder timer der vises på print ***
								strSQL2 = "SELECT fakid, aktid, mid, fak, venter, tekst, enhedspris, beloeb, "_
								&" showonfak, medrabat, enhedsang FROM fak_med_spec WHERE fakid = "& id &" AND aktid = "& oRec("aktid") &" AND enhedspris = "& ehpris &" AND showonfak = 1"
								'Response.write strSQL2
								'Response.flush
								oRec2.open strSQL2, oConn, 3 
								while not oRec2.EOF 
								
								medEhpris = oRec2("enhedspris")
								medBelob = oRec2("beloeb")
								if intFaktype = 1 then
								medBelob = -(medBelob)
								medEhpris = -(medEhpris)
								end if
								
								medrabat = oRec2("medrabat")
								
								select case oRec("enhedsang")
						        case 0
						        enhedsang = "pr. time"
						        case 1
						        enhedsang = "pr. stk."
						        case 2
						        enhedsang = "pr. enhed"
						        end select
        								
								strFakdet = strFakdet &"<tr>"_
								&"<td valign=top align=right style='padding-right:2;'><font class=megetlillesort>"& formatnumber(oRec2("fak"), 2) &"</td>"_
								&"<td colspan=2 valign=top style='padding-left:5;'><font class=megetlillesort>" & oRec2("tekst") &"</td>"_
								&"<td valign=top align=right style='padding-right:10;'><font class=megetlillesort>" & formatcurrency(medEhpris, 2) &" "& enhedsang &"</td>"
								
								if cint(visrabatkol) = 1 then
								strFakdet = strFakdet &"<td valign=top align=right style='padding-right:10;'>"
								
								    if cdbl(medrabat) <> 0 AND cint(visrabatkol) = 1 then
								    strFakdet = strFakdet &"<font class=megetlillesort>" & (medrabat * 100) &" %</td>"
								    else
								    strFakdet = strFakdet &"&nbsp;</td>"
								    end if
								
								end if
								
								strFakdet = strFakdet &"<td valign='top' align=right><font class=megetlillesort>"& formatcurrency(medBelob) &"</td>"_
								&"</tr>"
								
								h = h + 20
								
								oRec2.movenext
								wend
								oRec2.close 
								
								'**** til ikke printbar medarbjeder timeoversigt (højre side) *** 
								if lastaktid <> oRec("aktid") then 
								strMedarbFakdet = strMedarbFakdet &"<tr><td colspan=4><b>"& oRec("beskrivelse") &"</b></td></tr>"
								end if
								
								mf = 0
								
								strSQL2 = "SELECT fakid, aktid, mid, fak, venter, tekst, enhedspris, beloeb, showonfak, medrabat FROM fak_med_spec WHERE fakid = "& id &" AND aktid = "& oRec("aktid") &" AND enhedspris = "& ehpris &" AND fak <> 0"
								oRec2.open strSQL2, oConn, 3 
								while not oRec2.EOF 
								
								medEhpris = oRec2("enhedspris")
								medBelob = oRec2("beloeb")
								
								if intFaktype = 1 then 'kreditnota
								medBelob = -(medBelob)
								medEhpris = -(medEhpris)
								end if
								
								medrabatThis = oRec2("medrabat")
								
								strMedarbFakdet = strMedarbFakdet &"<tr>"_
								&"<td valign=top align=right style='padding-right:5;'><font class=megetlillesort>"& formatnumber(oRec2("fak"), 2) &"</td>"_
								&"<td valign=top style='padding-left:5;'><font class=megetlillesort>" & left(oRec2("tekst"), 10) &"</td>"_
								&"<td valign=top align=right style='padding-right:10;'><font class=megetlillesort>" & formatcurrency(medEhpris, 2) &"</td>"
								
								if cint(visrabatkol) = 1 then
								strMedarbFakdet = strMedarbFakdet &"<td valign=top align=right style='padding-right:10;'>"
								
								    if cdbl(medrabatThis) <> 0 then
								    strMedarbFakdet = strMedarbFakdet &"<font class=megetlillesort>" & (medrabatThis * 100) &" %</td>"
								    else
								    strMedarbFakdet = strMedarbFakdet &"&nbsp;</td>"
								    end if
								
								end if
								
								strMedarbFakdet = strMedarbFakdet &"<td valign=top align=right><font class=megetlillesort>"& formatcurrency(medBelob,2) &"</td>"_
								&"</tr>"
								
								medarbejdertimer = medarbejdertimer + oRec2("fak")
								medarbejdersum = medarbejdersum + oRec2("beloeb")
								
								mf = 1
								oRec2.movenext
								wend
								oRec2.close 
								
								
								if mf = 0 then
								strMedarbFakdet = strMedarbFakdet &"<tr><td colspan=4 style='padding-left:20;'>-</td></tr>"
								end if
									
					lastaktid = oRec("aktid")			
					oRec.movenext
					wend
					oRec.close 
					
					
					if len(medarbejdertimer) <> 0 then
					medarbejdertimer = medarbejdertimer
					else
					medarbejdertimer = 0
					end if
					
					if len(medarbejdersum) <> 0 then
					medarbejdersum = medarbejdersum
					else
					medarbejdersum = 0
					end if
					
					if medarbejdertimer <> intFaktureretTimer then
					bgthis = "#ffffe1"
					else
					bgthis = "#ffffff"
					end if
					
					'** Er der overensstemmelse mellem antal fakturatimer og antal medarbjeder timer og beløb?
					if intFaktype = 1 then
					medarbejdersum = -(medarbejdersum)
					end if
					
					strMedarbFakdet = strMedarbFakdet &"<tr><td bgcolor="& bgthis &" colspan=4 class=lille><br>"_
					&" Medarbejder timer: "& formatnumber(medarbejdertimer, 2) &"<br>"_
					&" Medarbejder beløb: "& formatcurrency(medarbejdersum, 2) &"</td></tr>"
					
					
					            '*** Materialer *****
								strSQLmat = "SELECT matid, matnavn, matvarenr, matantal, matenhed, "_
							    &" matenhedspris, matrabat, matbeloeb, matshowonfak FROM fak_mat_spec WHERE matfakid = "&id
								
								'Response.write strSQLmat
								'Response.Flush
								
								oRec2.open strSQLmat, oConn, 3
                                while not oRec2.EOF
                                        
                                        
                                        if intFaktype <> 1 then
                                        matBelob = oRec2("matbeloeb")
                                        matEhpris = oRec2("matenhedspris")
                                        else
								        matBelob = -(oRec2("matbeloeb"))
								        matEhpris = -(oRec2("matenhedspris"))
								        end if
                                        
                                    	strFakmat = strFakmat &"<tr>"_
								        &"<td valign=top align=right style='padding-right:2;'>"& formatnumber(oRec2("matantal"), 2) &"</td>"_
								        &"<td colspan=2 valign=top style='padding-left:5;'>" & oRec2("matnavn") &"</td>"_
								        &"<td valign=top align=right style='padding-right:10;'>" & formatcurrency(matEhpris) &" pr. "& oRec2("matenhed") &"</td>"
								        
								        if cint(visrabatkol) = 1 then
								        strFakmat = strFakmat &"<td valign=top align=right style='padding-right:10;'>"
								       
								            if cdbl(oRec2("matrabat")) <> 0 then
								            strFakmat = strFakmat &"" & (oRec2("matrabat") * 100) &" %</td>"
								            else
								            strFakmat = strFakmat  &"&nbsp;</td>"
								            end if
								        
								        end if
								        
								        strFakmat = strFakmat &"<td valign='top' align=right>"& formatcurrency(matBelob) &"</td>"_
								        &"</tr>"
                                    
                                    matsubtotal = matsubtotal + (matBelob)
                                    
                                
                                oRec2.movenext
                                wend
                                oRec2.close
					
					
		
		else '**** Fakturering af Aftaler ***
		
				
				'** Betalingsbetingelder **
				
				strFakdet = strFakdet &"<tr>"_
				&"<td valign=top align=right style='padding-right:2;'>"& formatnumber(intFaktureretTimer, 2) &"</td>"_
				&"<td colspan=2 valign=top style='padding-left:5;'>" & strJobBesk &"</td>"
			
				if intFaktureretTimer <> 0 AND intFaktureretBelob <> 0 then
				intstkpris = (intFaktureretBelob/intFaktureretTimer)
				else
				intstkpris = 0
				end if
				
				strFakdet = strFakdet & "<td valign=top align=right style='padding-right:10px;'>" & formatcurrency(intstkpris, 2) &"</td>"_
				&"<td valign='top' align=right style='padding-right:10px;'>"& (100*intAftRabat) &" %</td>"_
				&"<td valign='top' align=right colspan=2>"& formatcurrency(intFaktureretBelob,2) &"</td>"_
				&"</tr>" 
		
				
				'strMedarbFakdet = "Der kan ikke angives timer på en medarbejder når det er en aftale der faktureres."
				
		end if
		
		
		'**** Aktiviteter / timer ***
		if len(trim(strFakdet)) <> 0 then
 		call udspecificering(strFakdet,1)
		
		
		if jobid <> 0 then
		
		    '**** Sub-tot. aktiviteter ***
		    call subtotakt()
		    end if
    		
		    '*** Materialer ***
		    if len(trim(strFakmat)) <> 0 then
		    call udspecificering(strFakmat,2)
    		
		    '**** Sub-tot. materialer ***
		    call subtotmat()
		    end if
		
		end if
		
		'*** Rykker gebyrer ***
		call rykkergebyrer()
		
		
		'**** Totaler og moms ***
		call totogmoms()
		
		
		'*** Komm. og bet. betingelser ***
		call komogbetbetingelser	
		
		
		Response.Write "</div>" 'fakturaside
		
		
		
		stDatoKri = request("FM_start_aar_ival") &"/"& request("FM_start_mrd_ival") &"/"& request("FM_start_dag_ival")  
		slutdato = request("FM_slut_aar_ival") &"/"& request("FM_slut_mrd_ival") &"/"& request("FM_slut_dag_ival")  
		
		
		'*** Joblog ****
		if visjoblog = 1 then%>
		
		<!-- joblog -->
		
	    <div id="joblog" style="position:relative; width:640px; left:10px; top:100px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; padding:10px 10px 10px 10px;">
		<table width=100% border=0 cellspacing=0 cellpadding=0 style="page-break-before:always;">
		<tr>
			<td><h4>Joblog bilag:</h4></td>
		</tr>
		<tr>
		
		<%
		call joblog(jobid, stdatoKri, slutdato, aftid)
		%>
		
		</table>
		</div>
		<!-- Joblog slut -->
		<%end if 	
		
		
		
		
		'*** Matlog ****
		if vismatlog = 1 then%>
		
		<!-- matlog -->
		<br /><br /><br /><br />
		 <div id="matlog" style="position:relative; width:640px; left:10px; top:100px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; padding:10px 10px 10px 10px;">
		
		<table border=0 cellspacing=0 width=620 cellpadding=0 style="page-break-before:always;">
		<tr>
			<td><h4>Materiale-log bilag:</h4></td>
		</tr>
		</table>
		
		<table width=100% cellspacing=0 cellpadding=0 border=0>
		<tr>
	                            <td><b>Dato</b></td>
	                            <td><b>Navn</b></td>
	                            <td><b>Varenr</b></td>
	                            <td align=right><b>Antal</b></td>
	                            <td align=right><b>Enheds pris</b></td>
	                            <td align=right><b>Enhed</b></td>
	                            <td align=right><b>Pris ialt</b></td>
	                        </tr>
		
		<%
		call matlog(jobid, stdatoKri, slutdato, aftid)
		%>
		
		</table>
		</div>
		
		
		<!-- matlog slut -->
		<%end if %>		
		
						
		<br /><br /><br /><br /><br /><br /><br /><br />
            &nbsp;
		
		</div>
		
											
		<%
		'**** Højreside (ikke printbar) ***
		'call ikkeprintbar
		
		
		if print = "n" then %>
		<div id=funktioner style="position:absolute; z-index:200; top:30px; width:300px; left:520px; border:2px #8caae6 DASHED; background-color:pink; padding:5px 5px 5px 5px;">
		<b>Udskrift funktioner:</b> <br />
		<a href="fak_godkendt.asp?print=j&jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=id%>&kid=<%=kid%>&FM_job=<%=Request("FM_job")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag_ival=<%=request("FM_start_dag_ival")%>&FM_start_mrd_ival=<%=request("FM_start_mrd_ival")%>&FM_start_aar_ival=<%=request("FM_start_aar_ival")%>&FM_slut_dag_ival=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd_ival=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar_ival=<%=request("FM_slut_aar_ival")%>&FM_fakint_ival=<%=request("FM_fakint_ival")%>" target="_blank" class=vmenu>Print</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<a href="erp_make_pdf.asp?lto=<%=lto%>&faknr=<%=varFaknr%>&jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=id%>&kid=<%=kid%>&FM_job=<%=Request("FM_job")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag_ival=<%=request("FM_start_dag_ival")%>&FM_start_mrd_ival=<%=request("FM_start_mrd_ival")%>&FM_start_aar_ival=<%=request("FM_start_aar_ival")%>&FM_slut_dag_ival=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd_ival=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar_ival=<%=request("FM_slut_aar_ival")%>&FM_fakint_ival=<%=request("FM_fakint_ival")%>" target="_blank" class=vmenu>PDF</a><br /><br />
		<b>Gem PDF ovenfor, og email faktura til:</b><br />
		<%
		strSQL = "SELECT navn, email FROM kontaktpers WHERE kundeid = " & kid
		'Response.Write strSQL 
		'Response.Flush
		
		oRec.open strSQL, oConn, 3
        while not oRec.EOF
        
        Response.Write "<i>"& oRec("navn") & "</i>, <a href='mailto:"&oRec("email")&"&subject=Faktura: "& varFaknr &"' class=vmenu>" & oRec("email") & "</a><br>"
        
        oRec.movenext
        wend
        oRec.close
		%>
		</div>
		<!--&nbsp;&nbsp;|&nbsp;&nbsp;Excel.-->
	    <%end if
	    
	    
	    
	    if print = "j" then
	    
	    Response.Write("<script language=""JavaScript"">window.print();</script>")
	    
	    end if
	    
		
		Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>her:"& cdate(fakdato)
		
                		        
		        
		
		
											
											
											
											
											
											
											
	
	
	
	
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->


