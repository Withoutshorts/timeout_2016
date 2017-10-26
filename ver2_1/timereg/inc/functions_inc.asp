<%





sub ltoPosteringer

                            select case lto 
							case "execon", "immenso"
							
							'** Netto Umoms
							call opretPosteringSingle(oprid, "2", "dbopr", "2010100", 0, 0, intBeloeb, 0, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorrefId)
		                    
		                    '** Netto Mmoms
							call opretPosteringSingle(oprid, "2", "dbopr", "2010200", 0, 0, intTotalDeb, 0, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorrefId)
		                    
		                     '** Kørsel Km ex. moms
		                     transbeloebPos = 0
		                     strSQLselmf = "SELECT SUM(aktpris*(kurs/100)) AS transbeloeb FROM faktura_det WHERE fakid = " & thisfakid & " AND enhedsang = 3 GROUP BY fakid"
							 oRec4.open strSQLselmf, oConn, 3
							 if not oRec4.EOF then
							 
							 transbeloebPos = oRec4("transbeloeb")
							    
							 end if
							 oRec4.close
							 
							 transbeloebPos = replace(transbeloebPos, ".", "")
		                     transbeloebPos = replace(transbeloebPos, ",", ".")
		                    
							call opretPosteringSingle(oprid, "2", "dbopr", "2010300", 0, 0, transbeloebPos, 0, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorrefId)
		                    
		                    '** Moms
		                    intMoms = Replace(intMoms, ".", "")
		                    intMoms = Replace(intMoms, ",", ".")
							call opretPosteringSingle(oprid, "2", "dbopr", "2010400", 0, 0, intMoms, 0, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorrefId)
		                    
		                    '** Udlæg ex. moms
		                     matbeloebPos = 0
		                     strSQLselmf = "SELECT SUM(matbeloeb*(kurs/100)) AS matbeloeb FROM fak_mat_spec WHERE matfakid = " & thisfakid & " GROUP BY matfakid"
							 oRec4.open strSQLselmf, oConn, 3
							 if not oRec4.EOF then
							 
							 matbeloebPos = oRec4("matbeloeb")
							    
							 end if
							 oRec4.close
							 
							 matbeloebPos = replace(matbeloebPos, ".", "")
		                     matbeloebPos = replace(matbeloebPos, ",", ".")
		                    
							 call opretPosteringSingle(oprid, "2", "dbopr", "2010500", 0, 0, matbeloebPos, 0, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorrefId)
		                    
		                    '** Aktiviteter ex. moms
		                     aktbeloebPos = 0
		                     strSQLselmf = "SELECT SUM(aktpris*(kurs/100)) AS aktbeloeb FROM faktura_det WHERE fakid = " & thisfakid & " AND enhedsang <> 3 GROUP BY fakid"
							 oRec4.open strSQLselmf, oConn, 3
							 if not oRec4.EOF then
							 
							 aktbeloebPos = oRec4("aktbeloeb")
							    
							 end if
							 oRec4.close
							 
							 aktbeloebPos = replace(aktbeloebPos, ".", "")
		                     aktbeloebPos = replace(aktbeloebPos, ",", ".")
		                    
							call opretPosteringSingle(oprid, "2", "dbopr", "2010600", 0, 0, aktbeloebPos, 0, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorrefId)
		
							
							'Response.end
							end select

end sub







function beregnTimerDage(strMrd)

Select case strMrd

case 1
varStrMrdA = 11
varStrMrdB = 12
case 2
varStrMrdA = 21
varStrMrdB = 22
case 3
varStrMrdA = 31
varStrMrdB = 32
case 4
varStrMrdA = 41
varStrMrdB = 42
case 5
varStrMrdA = 51
varStrMrdB = 52
case 6
varStrMrdA = 61
varStrMrdB = 62
case 7
varStrMrdA = 71
varStrMrdB = 72
case 8
varStrMrdA = 81
varStrMrdB = 82
case 9
varStrMrdA = 91
varStrMrdB = 92
case 10
varStrMrdA = 101
varStrMrdB = 102
case 11
varStrMrdA = 111
varStrMrdB = 112
case 12
varStrMrdA = 121
varStrMrdB = 122
end select

Redim preserve antalarbDageA(varStrMrdA)
Redim preserve totTimerA(varStrMrdA)
Redim preserve intFakbareTimerA(varStrMrdA)
Redim preserve omsA(varStrMrdA)
Redim preserve timerTjTotA(varStrMrdA)

Redim preserve avrTimerA(varStrMrdA) 
Redim preserve avrFaktimerA(varStrMrdA)


Redim preserve antalarbDageB(varStrMrdB) 
Redim preserve totTimerB(varStrMrdB)
Redim preserve intFakbareTimerB(varStrMrdB)
Redim preserve omsB(varStrMrdB)
Redim preserve timerTjTotB(varStrMrdB)

Redim preserve avrTimerB(varStrMrdB) 
Redim preserve avrFaktimerB(varStrMrdB)

if strDag < 15 then
			
			if objRec("Tdato") <> lastDato then
			lastDato = objRec("Tdato")
			antalarbDageA(varStrMrdA) = antalarbDageA(varStrMrdA) + 1
			else
			antalarbDageA(varStrMrdA) = antalarbDageA(varStrMrdA)
			end if
			
			totTimerA(varStrMrdA) = totTimerA(varStrMrdA) + objRec("timer")
			
				if objRec("Tfaktim") = "Yes" then
					intFakbareTimerA(varStrMrdA) = intFakbareTimerA(varStrMrdA) + objRec("timer")
					
					strSQL = "SELECT jobnr, jobTpris FROM job WHERE jobnr = '" & objRec("Tjobnr") & "'"
					oRec2.open strSQL, connection, 3
					
					omsA(varStrMrdA) = (omsA(varStrMrdA) + (oRec2("jobTpris") * objRec("timer")))
					timerTjTotA(varStrMrdA) = timerTjTotA(varStrMrdA) + objRec("timer")
					oRec2.close
				else
					intFakbareTimerA(varStrMrdA) = intFakbareTimerA(varStrMrdA)
				end if
			
		else
		
			if objRec("Tdato") <> lastDato then
			lastDato = objRec("Tdato")
			antalarbDageB(varStrMrdB) = antalarbDageB(varStrMrdB) + 1
			else
			antalarbDageB(varStrMrdB) = antalarbDageB(varStrMrdB)
			end if
			
			'Response.write "antalarbDageB(varStrMrdB):" & antalarbDageB(varStrMrdB) & "<br><hr>"
			
			totTimerB(varStrMrdB) = totTimerB(varStrMrdB) + objRec("timer")
			
				if objRec("Tfaktim") = "Yes" then
				
					intFakbareTimerB(varStrMrdB) = intFakbareTimerB(varStrMrdB) + objRec("timer")
					
					strSQL = "SELECT jobnr, jobTpris FROM job WHERE jobnr = '" & objRec("Tjobnr") & "'" 
					oRec2.open strSQL, connection, 3
					
					omsB(varStrMrdB) = (omsB(varStrMrdB) + (oRec2("jobTpris") * objRec("timer")))
					timerTjTotB(varStrMrdB) = timerTjTotB(varStrMrdB) + objRec("timer")
					oRec2.close
				else
					intFakbareTimerB(varStrMrdB) = intFakbareTimerB(varStrMrdB)
				end if
		end if
		
		
end function


public dagparset
function datofindes(dag, md, aar)


'Response.Write aar & " md:" & md & "<br>"

			Select case md
			case "4", "6", "9", "11"
				if dag = 31 then
				dagparset = 30
				else
				dagparset = dag
				end if
			case "2"
				
				
				if dag > 28 then
				    select case right(aar, 2)
				    case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
				    dagparset = 29
				    case else
				    dagparset = 28
				    end select
				else
				dagparset = dag
				end if
				
				
			case else
			dagparset = dag
			end select
	
	'Response.Write "dagparset" & dagparset & "<br>"
			
end function


public intMoms, intNetto
public intTotal, momsProcent 

function beregnmoms(fak_el_pos, intBelob, varKonto)
	    
	    
	    intTotal = replace(intBelob, ".", ",")
	    intMoms = 0
		intNetto = 0
		
		'** Beregner moms og netto **
		'*** Konto ***
		strSQL = "SELECT momskode, kvotient FROM kontoplan LEFT JOIN momskoder ON (momskoder.id = momskode) WHERE kontonr = "& varKonto
		
		'Response.Write strSQL
		'Response.flush
		'Response.End
		
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
		    
		    kvo = oRec("kvotient")
			
			if fak_el_pos = "fak" then
			
			
			
			    select case oRec("kvotient")
			    case 1
			    momskvo = 0
			    case 5
			    momskvo = 0.25
			    case 21
			    momskvo = 0.05
			    end select
    			
			    intMoms = formatnumber((intTotal * momskvo),2) 
			
			else
			
			    if cint(oRec("kvotient")) <> 1 then
			    intMoms = (intTotal / oRec("kvotient"))  'request("FM_moms")
			    else
			    intMoms = 0
			    end if
			
			end if
			
	    
	    
			
		if fak_el_pos = "pos" then
		intNetto = SQLBless3(intTotal) - SQLBless3(intMoms) 
		else
		intNetto = SQLBless3(intTotal) ' - SQLBless3(intMoms)
		intTotal = (intTotal/1) + (intMoms/1)
	    end if
		
		
		'Response.Write "her"
	    'Response.Write "fak_el_pos: " & fak_el_pos & " Kvo: "& oRec("kvotient") & ""_
	    '&"<br>momskvo:" & momskvo & ""
		
		'Response.Write "<hr>::##::<br>IntTotal: " & intTotal & "<br>"
		'Response.Write "Netto: "& intNetto &"<br>"
		'Response.Write "Moms beregnet " & intMoms &"<br>"
		
		'Response.flush
		
		end if
		oRec.close
		
	    
		
end function


public oprid, momsafsluttetDato, useleftdiv, tjkPosDato


function opretpos(fak_el_pos,func,intkontonr,modkontonr,intTotal_konto,intTotal_modkonto,intMoms,strEditor,varBilag, strTekst, posteringsdato, intStatus, vorref)
		
		strDato = year(now) &"/"& month(now) &"/"& day(now)
		
		'*** Momsperiode lukket ***
		momsafsluttetDato = "1/1/2001"
		strSQLmomsafs = "SELECT afslutdato FROM momsafsluttet WHERE id <> 0 ORDER BY afslutdato DESC"
		
		'Response.Write strSQLmomsafs
		'Response.flush
		
		oRec.open strSQLmomsafs, oConn, 3
		if not oRec.EOF then
		momsafsluttetDato = oRec("afslutdato")
		end if
		oRec.close
		
		'Response.Write cdate(posteringsdato) &" > "&  cdate(momsafsluttetDato)
		'Response.end
		'momsafsluttetDato = "1-1-1800"
		
		if fak_el_pos = 2 then
		tjkPosDato = cdate(showfakDato) 'day(posteringsdato)&"/"&month(posteringsdato)&"/"&year(posteringsdato)
		else
		tjkPosDato = cdate(posteringsdato)
		end if
		
		if cdate(tjkPosDato) < cdate(momsafsluttetDato) then
		
		    
		    %>
			<!--#include file="../../inc/regular/header_lysblaa_inc.asp"-->
			<%
			useleftdiv = "tt"
			errortype = 99
			call showError(errortype)
			
			Response.end
		    
		
		else
	
		if func = "dbopr" then
		
		strSQL = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att)"_
		&" VALUES ('"& strEditor &"', '"& strDato &"', '"& modkontonr &"', '"& intkontonr &"', '"& varBilag &"', "_
		&" "& intTotal_konto &", "& intTotal_konto &", 0, '"& strTekst &"', "_
		&" '"& posteringsdato &"', "_
		&" "& intStatus &", "& vorref &")"
		
		'Response.Write strSQL
		'Response.end
		
		
		oConn.execute(strSQL)
		
		
		'*** Finder det netop opr. id ***
		strSQL = "SELECT id FROM posteringer ORDER BY id DESC"
		oRec.Open strSQL, oConn, 3 
		if not oRec.EOF then
		oprid = oRec("id")
		end if
		oRec.close
		
		oConn.execute("UPDATE posteringer SET oprid = "& oprid &" WHERE id = "& oprid &"")
	    
	    if fak_el_pos = 2 then
	    oConn.execute("UPDATE fakturaer SET oprid = "& oprid &" WHERE Fid = "& thisfakid &"")
	    end if
		
		
		'**** Modkonto ***
		strSQL2 = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, "_
		&" beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
		&" VALUES ('"& strEditor &"', '"& strDato &"', '"& intkontonr  &"', '"& modkontonr &"', '"& varBilag &"', "_
		&" "& intTotal_modkonto &", "& intTotal_modkonto &", "& intMoms &", '"& strTekst &"', "_
		&" '"& posteringsdato &"', "_
		&" "&intStatus&", "& vorref &", "& oprid &")"
		
		'Response.Write "<br>"& strSQL2
	    'Response.end
	
		oConn.execute(strSQL2)
		
		
		
		else
		
		strSQL = "UPDATE posteringer SET editor = '" &strEditor &"', dato = '" & strDato &"',"_
		&" modkontonr = "& modkontonr &",  kontonr = "& intkontonr &", bilagsnr = '"& varBilag &"', beloeb = "& intTotal_konto &", "_
		&" nettobeloeb = "&intTotal_konto&", moms = 0, tekst = '"&strTekst&"', "_
		&" posteringsdato = '"& posteringsdato &"', status = "&intStatus&","_
		&" att = "& vorref &" WHERE id = "&id
		
		oConn.execute(strSQL)
		
		'Response.Write strSQL
		'Response.flush
		
		
		'*** Finder det netop opr. id ***
		strSQL = "SELECT oprid FROM posteringer WHERE id="& id
		oRec.Open strSQL, oConn, 3 
		if not oRec.EOF then
		oprid = oRec("oprid")
		end if
		oRec.close
		
		'**** modkonto ****
		strSQL2 =  "UPDATE posteringer SET editor = '" &strEditor &"', dato = '" & strDato &"',"_
		&" modkontonr = "& intkontonr  &",  kontonr = "& modkontonr &", beloeb = "& intTotal_modkonto &", "_
		&" nettobeloeb = "&intTotal_modkonto&", moms = "&intMoms&", tekst = '"&strTekst&"', posteringsdato = '"& posteringsdato &"', "_
		&" att = "& vorref &" WHERE (oprid = "& oprid &" AND id <> "&id&")"
		
		oConn.execute(strSQL2)
		
		'Response.Write "<br>"& strSQL2
		'Response.flush
		'Response.End 
		
		end if
				
				
		end if 'momsperiode afsluttet
		
end function



function opretPosteringSingle(oprid, fak_el_pos, func, intkontonr, modkontonr, intTotal, intNetto, intMoms, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorref)
		
		'Response.Write "oprid:" & oprid
		'Response.flush
		
		strDato = year(now) &"/"& month(now) &"/"& day(now)
		
		'*** Momsperiode Lukket ***'
		momsafsluttetDato = "1/1/2001"
		strSQLmomsafs = "SELECT afslutdato FROM momsafsluttet WHERE id <> 0 ORDER BY afslutdato DESC"
		
	    oRec.open strSQLmomsafs, oConn, 3
		if not oRec.EOF then
		momsafsluttetDato = oRec("afslutdato")
		end if
		oRec.close
		
		if fak_el_pos = 2 then
		tjkPosDato = cdate(showfakDato) 
		else
		tjkPosDato = cdate(posteringsdato)
		end if
		
		
		
		if cdate(tjkPosDato) < cdate(momsafsluttetDato) then
		
		    
		    %>
			<!--#include file="../../inc/regular/header_lysblaa_inc.asp"-->
			<%
			useleftdiv = "tt"
			errortype = 99
			call showError(errortype)
			
			Response.end
		    
		
		else
	
		
		
		strSQL = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, "_
		&" beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
		&" VALUES ('"& strEditor &"', '"& strDato &"', "& modkontonr &", "& intkontonr &", '"& varBilag &"', "_
		&" "& intTotal &", "& intNetto &", "& intMoms &", '"& strTekst &"', "_
		&" '"& posteringsdato &"', "_
		&" "& intStatus &", "& vorref &", "& oprid &")"
		
		'Response.Write strSQL & "<hr>"
		'Response.end 
		
		oConn.execute(strSQL)
		
		        if oprid = "0" then
        		
		        '*** Finder det netop opr. id ***'
		        strSQL = "SELECT id FROM posteringer ORDER BY id DESC"
		        oRec.Open strSQL, oConn, 3 
		        if not oRec.EOF then
		            oprid = oRec("id")
		        end if
		        oRec.close
		        
		        
        		
		        oConn.execute("UPDATE posteringer SET oprid = "& oprid &" WHERE id = "& oprid &"")
        	    
	               
		        end if
		
		
		
		
				
				
		end if 'momsperiode afsluttet


end function


 '*** Bruges af ERP fak og fak godkendt **'
public ekspTxt_kk
 function joblog(jobid,stdatoKri,slutdato,aftid, viskunfakturalinjer)

   session.LCID = 1030
   laktaktid = 0
  
  

       if jobid <> 0 then
       strSQLjobaftKri = "j.id = "& jobid &" AND tjobnr = j.jobnr"
       else
       strSQLjobaftKri = "seraft = "& aftid &" AND tjobnr = j.jobnr "
       end if

    if viskunfakturalinjer <> 0 then 
    ekspTxt_kk = ""
    '*** Henter fakturalinjer der er med på faktura på medarbejdere / ellers hebnter joblog
    'fakturaid = Vis KUN linjer med på den faktura BF diffrentirede linjer på faktura :: Kasserapport


                'strSQL = "SELECT tmnr, tmnavn, tdato, timer, timerkom, "_
				'&" taktivitetnavn, taktivitetid, tfaktim, jobnavn, jobnr, a.faktor, timepris, t.valuta, v.valutakode, avarenr, t.kurs AS kurs, tjobnavn, tjobnr "_
				'&" FROM faktura_det fd "_
                '&" LEFT JOIN timer t ON (taktivitetid = fd.aktid AND tfaktim <> 5 AND tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"_
                '&" LEFT JOIN job j ON (j.jobnr = t.tjobnr)"_
                '&" LEFT JOIN aktiviteter a ON (a.id = fd.aktid)"_
				'&" LEFT JOIN valutaer v ON (v.id = t.valuta)"_
				'&" WHERE a.id = fd.aktid AND fd.fakid = "& id &" ORDER BY jobnr, tdato DESC"

                strSQL = "SELECT fakid, aktid, mid, fak AS timer, tekst AS tmnavn, jobnavn, jobnr, fms.valuta, v.valutakode, "_
                &" fms.kurs AS kurs, avarenr, fakturerbar, fakdato, a.navn AS aktnavn, fms.beloeb FROM fak_med_spec fms "_
                &" LEFT JOIN aktiviteter a ON (a.id = fms.aktid) "_
                &" LEFT JOIN job j ON (j.id = a.job) "_
                &" LEFT JOIN fakturaer f ON (f.fid = fms.fakid) "_
                &" LEFT JOIN valutaer v ON (v.id = fms.valuta) "_ 
                &" WHERE fakid = "& id


    else

                ekspTxt_kk = ekspTxt_kk
                strSQL = "SELECT tmnr, tmnavn, tdato, timer, timerkom, "_
				&" taktivitetnavn, taktivitetid, tfaktim, jobnavn, jobnr, a.faktor, timepris, t.valuta, v.valutakode, avarenr, t.kurs AS kurs, tjobnavn, tjobnr "_
				&" FROM timer t "_
                &" LEFT JOIN job j ON (j.jobnr = t.tjobnr)"_
                &" LEFT JOIN aktiviteter a ON (a.id = t.taktivitetid)"_
				&" LEFT JOIN valutaer v ON (v.id = t.valuta)"_
				&" WHERE "&  strSQLjobaftKri &" AND tfaktim <> 5 AND "_
				&" (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"') AND a.id = taktivitetid ORDER BY jobnr, tdato DESC"

    end if
   
  
   
			    'if session("mid") = 1 then
				'Response.write strSQL & "viskunfakturalinjer:" & viskunfakturalinjer
                'Response.flush
				'end if

				oRec.open strSQL, oConn, 3 
				x = 0
				while not oRec.EOF 
					

                if cint(viskunfakturalinjer) = 0 then '*** Joblog eller faktura FAK_MED_SPEC LOG


               


			            if oRec("jobnr") <> lastjobnr then
						
						                    if oRec("tfaktim") = 1 then
						                    imgthis = "&nbsp;"
						                    else
						                    imgthis = "<font class=roed>"& txt_045 &"</font>"
						                    end if
            					
            					
					                        if x <> 0 then
					                        pdtop = "30px"
					                        else
					                        pdtop = "5px"
					                        end if
            					 
            					   
				                            if visjoblog_timepris = 1 then
				                            visjoblog_timepris_CHK = "CHECKED"
				                            else
				                            visjoblog_timepris_CHK = ""
				                            end if
                					
				                            if visjoblog_enheder = 1 then
				                            visjoblog_enheder_CHK = "CHECKED"
				                            else
				                            visjoblog_enheder_CHK = ""
				                            end if
				                    
				                            if visjoblog_mnavn = 1 then
				                            visjoblog_mnavn_CHK = "CHECKED"
				                            else
				                            visjoblog_mnavn_CHK = ""
				                            end if
            					    
            					   
                                    
            					
            		
            						
					                    Response.write "<tr><td colspan=6 style=padding-top:"&pdtop&";><h4>"& oRec("jobnavn") &" ("& oRec("jobnr") &")</h4></td></tr>"
					                    Response.Write "<tr><td width=100><b>"& erp_txt_042 &"</b></td>"_
					                    &"<td><b>"& erp_txt_046 &"</b> (type)</td>"_
					                    &"<td width=100 align=right><b>"& erp_txt_013 &"</b></td>"
            					
					                    if thisfile = "erp_oprfak_fs" then
					                    chkboxesShow1 = ""
					                    chkboxesShow2 = ""
					                    chkboxesShow3 = ""
					                    else
					                    chkboxesShow1 = "<input id='FM_vis_joblog_timepris' name='FM_vis_joblog_timepris' value='1' type='checkbox' "& visjoblog_timepris_CHK &" />"
					                    chkboxesShow2 = "<input id='FM_vis_joblog_enheder' name='FM_vis_joblog_enheder' value='1' type='checkbox' "& visjoblog_enheder_CHK &" />"
					                    chkboxesShow3 = "<input id='FM_vis_joblog_mnavn' name='FM_vis_joblog_mnavn' value='1' type='checkbox' "& visjoblog_mnavn_CHK &" />"
					            
					                    end if
            					
					                    if (thisfile = "erp_oprfak_fs" AND visjoblog_timepris = 1) OR thisfile = "erp_fak.asp" then
					                    Response.Write "<td width=100 align=right>"& chkboxesShow1 &"<b>"& txt_015 &"</b></td>"
					                    end if
            					
					                    if (thisfile = "erp_oprfak_fs" AND visjoblog_enheder = 1) OR thisfile = "erp_fak.asp" then
					                    Response.Write "<td width=100 align=right>"& chkboxesShow2 &"<b>"& txt_048 &"</b></td>"
					                    end if
            					
            					        if (thisfile = "erp_oprfak_fs" AND visjoblog_mnavn = 1) OR thisfile = "erp_fak.asp" then
					                    Response.Write "<td align=right>"& chkboxesShow3 &"<b>"& txt_047 &"</b></td>"
					                    end if
            					
            					
					                    Response.Write "</tr>"
					
					        end if



					
					                '*** Timepris ***'
					                '**** job på print ell. ved fakopr ***'
            					    if thisfile <> "erp_fak.asp" then
            					        if kurs <> 0 then
                                        joblogtpris = oRec("timepris") * (100/kurs)
                                        else
                                        joblogtpris = oRec("timepris")
                                        end if
                                    else
                                        joblogtpris = oRec("timepris")
                                        valutaISO = ""
                                    end if
					
					
					Response.write "<tr><td colspan=6 style='border-bottom:1px #cccccc solid; padding:5px 0px 0px 0px;'><img src='ill/blank.gif' height=1 width=1 border=0><br></td></tr>"
					
					
					Response.write "<tr><td style='padding:5px 0px 2px 0px;'>"& formatdatetime(oRec("tdato"), 1) &"</td>"_
					&"<td style='padding:5px 0px 2px 0px;'>"& oRec("taktivitetnavn") 
					
					call akttyper(oRec("tfaktim"), 4)
					
					
					Response.Write "<font class=megetlillesort> ("& akttypenavn &")</td> "
					Response.Write "<td align=right style='padding:5px 0px 2px 0px;'><b>"& formatnumber(oRec("timer"),2) &"</b></td>"
					
					if (thisfile = "erp_oprfak_fs" AND visjoblog_timepris = 1) OR thisfile = "erp_fak.asp" then
				    Response.Write "<td align=right style='padding:5px 0px 2px 0px;'>"& formatnumber(joblogtpris,2) &" "& oRec("valutakode") &"</td>"
					end if
					
					enheder = 0
					enheder = (oRec("faktor") * oRec("timer"))
					
					if len(trim(enheder)) <> 0 then
					enheder = enheder
					else
					enheder = 0
					end if
					
					if (thisfile = "erp_oprfak_fs" AND visjoblog_enheder = 1) OR thisfile = "erp_fak.asp" then
					Response.Write "<td align=right style='padding:5px 0px 2px 20px;'>"& formatnumber(enheder,2) &"</td>"
					end if
					
					if (thisfile = "erp_oprfak_fs" AND visjoblog_mnavn = 1) OR thisfile = "erp_fak.asp" then
					Response.write "<td align=right style='padding:0px 0px 2px 20px;'><i>"& oRec("tmnavn") &"</i></td>"
					end if
					
					Response.Write "</tr>"
					
					if len(oRec("timerkom")) <> 0 then
					Response.write "<tr><td colspan=6 style='padding:2px 5px 5px 5px;' class=lille><i>"& oRec("timerkom") & "</i></td></tr>"
					end if
					

                    else 'FAK_MED_SPEC

					
                        

                           if lto = "bf" OR lto = "intranet - local" then

                                    if x > 0 then
                                    ekspTxt_kk = ekspTxt_kk & "xx99123sy#z"
                                    else
                                    ekspTxt_kk = ekspTxt_kk 
                                    end if

                            
                                    if isNull(oRec("avarenr")) <> true AND len(trim(oRec("avarenr"))) >= 7 AND instr(lcase(oRec("avarenr")), "m") >= 5 AND (lto = "bf" OR lto = "intranet - local") then

                                    kontoTxt = trim(oRec("avarenr"))
                                    kontonrLen = len(kontoTxt)
                                    kontonrM = instr(lcase(kontoTxt), "m") 
                               
                                        'if session("mid") = 1 then
                                        'kontonrLeft = "_L_" & kontoTxt 'mid(kontoTxt, 2, kontonrM-2)
                                        'kontonrRight = "_R_" & kontoTxt 'mid(kontoTxt, kontonrM+1, kontonrLen)
                                        'else
                                        kontonrLeft = mid(kontoTxt, 2, kontonrM-2)
                                        kontonrRight = mid(kontoTxt, kontonrM+1, kontonrLen)
                                        'end if


                                    else

                                    kontonrLeft = oRec("avarenr")
                                    kontonrRight = 0

                                    end if

                                    call meStamdata(oRec("mid"))
                                    bilagsnr = "" 'year(oRec("tdato"))&month(oRec("tdato"))&day(oRec("tdato")) 

                                    ekspTxt_kk = ekspTxt_kk & chr(34) & formatdatetime(oRec("fakdato"), 2) & chr(34) &";"& chr(34) & oRec("jobnavn") &" ["& meInit &"]" & chr(34) &";" & bilagsnr &";"& chr(34) & kontonrLeft & chr(34) &";" & chr(34) & oRec("aktnavn") & chr(34) &";"
                            
                                    belob = formatnumber(oRec("beloeb"), 2)'formatnumber(oRec("timer") * oRec("enhedspris"), 2)
                                    frakurs = oRec("kurs")
                                    valBelobBeregnet = belob
                               

	                                ekspTxt_kk = ekspTxt_kk & replace(formatnumber(valBelobBeregnet, 2), ".", "") &";;;"& chr(34) & kontonrRight & chr(34) &""


				            end if 'lto BF kk
				
                        end if '** joblog / fak_med_Spec


				x = x + 1
				lastjobnr = oRec("jobnr")  
				oRec.movenext
				wend
				oRec.close 
				
				if x = 0 then
				Response.write "<tr><td colspan=3><br><br>&nbsp;</td></tr>"
				end if
				%>
				


									
    
    <%
    end function
    
    
   function matlog(jobid,stdatoKri,slutdato, aftid)
    
   if jobid <> 0 then
   strSQLjobaftKri = "jobid = "& jobid &""
   else
   strSQLjobaftKri = "serviceaft = "& aftid &""
   end if
    
   
            matbelobLogTotal = 0
            '**** Materiale log ***
            strSQL = "SELECT matid, matnavn, matantal AS matantal, matsalgspris, "_
            &" matenhed, matvarenr, forbrugsdato, mf.valuta, v.valutakode "_
            &" FROM materiale_forbrug mf "_
            &" LEFT JOIN valutaer v ON (v.id = mf.valuta)"_
			&" WHERE "& strSQLjobaftKri &" AND forbrugsdato BETWEEN '"& stdatoKri &"' AND '"& slutdato &"' ORDER BY  forbrugsdato DESC"
            
            'Response.Write strSQL
            'Response.Flush
           
            oRec.open strSQL, oConn, 3
            m = 0
            while not oRec.EOF
            
             '**** job på print ell. ved fakopr ***'
		    if thisfile <> "erp_fak.asp" then
		        if kurs <> 0 then
                matSpris = oRec("matsalgspris") * (100/kurs)
                else
                matSpris = oRec("matsalgspris")
                end if
            else
                matSpris = oRec("matsalgspris")
                valutaISO = ""
            end if
            
            %>
             <tr>
            
            <td style="padding:2px 0px 2px 0px;"><%=formatdatetime(oRec("forbrugsdato"), 1) %></td>
            <td style="padding:2px 0px 2px 0px;"><%=oRec("matnavn") %></td>
            <td style="padding:2px 0px 2px 0px;"><%=oRec("matvarenr") %></td>
            <td style="padding:2px 0px 2px 0px;" align=right><%=oRec("matantal") %></td>
            <td style="padding:2px 0px 2px 0px;" align=right><%=formatnumber(matSpris, 2) &" "& oRec("valutakode") %></td>
            <td style="padding:2px 0px 2px 0px;" align=right><%=oRec("matenhed") %></td>
            
              <td style="padding:2px 0px 2px 0px;" align=right>   
            <%
            matbelobLogTotal = (oRec("matantal") * matSpris)
            %>
            <%=valutaISO&" "&formatnumber(matbelobLogTotal, 2) &" "& oRec("valutakode") %>
            </td>
           
        </tr>
           
            <%
           
            m = m + 1
            oRec.movenext
            wend
            oRec.close
            
            if m = 0 then
            Response.write "<tr><td colspan=7><br><br>&nbsp;</td></tr>"
            end if
            
            
    end function
    
    
    function joblogdiv()
    %>
    
     <div id="joblogdiv" style="position:absolute; visibility:hidden; display:none; left:5px; height:650px; top:105px; width:700px; padding:10px 10px 10px 10px; background-color:#ffffff;">

	    <!-- joblog -->
		<%if visjoblog = "1" then
		visjoblogCHK = "CHECKED"
		else
		visjoblogCHK = ""
		end if 
		%>
		
		
		<table width=100% border=0 cellspacing=0 cellpadding=0>
		<tr>
		    <td><input id="FM_visjoblog" name="FM_visjoblog" <%=visjoblogCHK %> type="checkbox" /> <b><%=erp_txt_500 %></b></td>
		</tr>
		<tr>
			<td><%=erp_txt_434 %> <br />
			<%=erp_txt_435 %>: <%=formatdatetime(showStDato, 1)%> - <%=formatdatetime(showSlutDato, 1)%>
			<br />
              
                
                &nbsp;</td>
		</tr>
		<tr>
		<td><div style="position:relative; visibility:visible; left:0px; top:0px; height:550px; border:0px YellowGreen dashed; padding:10px 10px 10px 10px; overflow:auto; background-color:#ffffe1; z-index:200;">
		
		<table border=0 cellspacing=0 cellpadding=0><%
		viskunfakturalinjer = 0
		call joblog(jobid, stdatoKri, slutdato, aftid, viskunfakturalinjer)
		
		%>
		</table>
		
		</div></td></tr>
		</table>
	    </div>
	    
	<%if jobid <> 0 then
	nst = "aktdiv"
	forr = "jobbesk"
	else
	forr = "aktdiv"
	nst = "betdiv"
	end if %>
	    
	    <div id=joblogdiv_2 style="position:absolute; visibility:hidden; display:none; top:776px; width:700px; left:5px; border:0px #8cAAe6 solid;">
        <table width=700 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('<%=forr%>')" class=vmenu><< <%=erp_txt_436 %></a></td><td align=right><a href="#" onclick="showdiv('<%=nst%>')" class=vmenu><%=erp_txt_392 %> >></a></td></tr></table>
        </div>
	    
		<!-- Joblog slut -->
    
    <%
    end function
    
    
    function matlogdiv()
    %>
     <%if vismatlog = "1" then
       matLogChk = "CHECKED"
       else
       matLogChk = ""
       end if 
        %>
     <div id=matlogdiv style="position:absolute; width:700px; visibility:hidden; display:none; height:600px; border:0px orange solid; top:105px; left:5px; padding:10px 10px 10px 10px; background-color:#ffffff;">
    	 <table width=100% cellspacing=0 cellpadding=0 border=0>
        <tr>
        <td valign=top>
            <input id="FM_vismatlog" name="FM_vismatlog" <%=matLogChk %> type="checkbox" value="1" /> <b><%=erp_txt_437 %></b>
            <br />
            <%=erp_txt_438 %><br />
            <%=erp_txt_435 %>: <%=formatdatetime(showStDato, 1)%> - <%=formatdatetime(showSlutDato, 1)%>
			
        </td>
        </tr>
        </table>
        <br />
        <table width=100% cellspacing=0 cellpadding=0 border=0>
        <tr>
        <td>
           <div style="position:relative; visibility:visible; height:500px; border:0px orange dashed; overflow:auto; background-color:#ffffe1; z-index:200; padding:10px 10px 10px 10px;">
            <table width=100% cellspacing=0 cellpadding=0 border=0>
            <tr>
            <td><b><%=erp_txt_439 %></b></td>
            <td><b><%=erp_txt_440 %></b></td>
            <td><b><%=erp_txt_441 %></b></td>
            <td align=right><b><%=erp_txt_442 %></b></td>
            <td align=right><b><%=erp_txt_443 %></b></td>
            <td align=right><b><%=erp_txt_444 %></b></td>
            <td align=right><b><%=erp_txt_445 %></b></td>
        </tr><%
        
        call matLog(jobid, stdatoKri, slutdato, aftid)
        %>
        
        </table>
            </div>
        </td>
        </table>
       </div>
       
    <%if jobid <> 0 then
	nst = "matdiv"
	else
	nst = "betdiv"
	end if %>
       
        <div id=matlogdiv_2 style="position:absolute; visibility:hidden; display:none; top:776px; width:600px; left:5px; border:0px #8cAAe6 solid;">
        <table width=700 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('aktdiv')" class=vmenu><< <%=erp_txt_391 %></a></td><td align=right><a href="#" onclick="showdiv('<%=nst%>')" class=vmenu><%=erp_txt_392 %> >></a></td></tr></table>
        </div>
    <%
    end function

%>