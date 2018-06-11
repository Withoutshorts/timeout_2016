    
    <div id="aktdiv" style="position:absolute; visibility:hidden; display:none; left:5px; top:105px; width:718px;">
   
    



    <%
    if func = "red" then
        if cDate(strTdato) <= cDate("15-11-2010") AND cDate(strDato) <= cDate("15-11-2010") then
        showOld = 1
        else
        showOld = 0
        end if
    else
    showOld = 0
    end if

    if showOld = 1 then 


		
		if len(intEnheder) <> 0 then
		intEnheder = formatnumber(intEnheder, 2)
		else
		intEnheder = formatnumber(0, 2)
		end if
	
		intEnheder = replace(intEnheder, ".", "")
		


        '** Enhedspris ***
		'** Hvis perafg. = 0 then enhpris = pris
		'** Hvis klip afh. then enhpris = pris/enheder uanset om aftalen er pariodeafh.
		
		if intEnheder <> 0 then
		intEnheder = intEnheder
		else
		intEnheder = 1
		end if
		
		if func = "red" then
		    stkpris = (intPris * 1+intrabat) / intEnheder 
		 else
    		
		    if intEnheder > 0 AND cint(intType) = 0 then '*** Ikke kreditnota
			    stkpris = (intPris/intEnheder)
		    else
    		    stkpris = 0
    		end if
    		
	    end if
		
		
		if len(stkpris) <> 0 then
		stkpris = formatnumber(stkpris, 2)
		else
		stkpris = formatnumber(0, 2)
		end if
		stkpris = replace(stkpris, ".", "")
		
		                     
		
	
		if func <> "red" AND cint(intType) = 1 then '*** Opret kreditnota 
		
		intPris = formatnumber(0, 2)
		else
		if len(intPris) <> 0 then
		    
		    if intRabat <> 0 AND func <> "red" then
		    intPris = intPris - (intPris * intRabat/100)
		    else
		    intPris = intPris
		    end if
		    
		    intPris = formatnumber(intPris, 2)
		else
		intPris = formatnumber(0, 2)
		end if
		
		end if
		
		intPris = replace(intPris, ".", "")
		%>
		

<% 


 fl_antal(a) = strTimer 'oRec("antal")

call htmlparseCSV(strJobBesk)
fl_besk(a) = htmlparseCSVTxt

fl_vis(a) = "CHECKED"
fl_enhpris(a) = stkpris
fl_valuta(a) = valuta 'oRec("valuta")
thisaktfunc(a) = "red"
fl_enhed(a) = intEnhedsang 'oRec("enhedsang")
fl_rabat(a) = intRabat
fl_momsfri(a) = 0' oRec("momsfri")
fl_belob(a) = intPris 'strBeloeb 'intPris
fl_jobaktid(a) = 0

a = 1

else

a = 0

end if 'showOld




if func = "red" then
        		
		        
        		
		        '*******************************************************************************'
		        '******************* Henter oprettede aktiviteter ******************************'
		        '*******************************************************************************'
		        strSQL = "SELECT fd.id, fd.antal, "_
		        &" fd.valuta, fd.beskrivelse, fak_sortorder, enhedsang, fd.fase AS fase, momsfri, fd.enhedspris, fd.valuta, fd.rabat, fd.aktpris, fd.aktid, fd.showonfak "_
		        &" FROM faktura_det AS fd "_
		        &" WHERE fakid = "& intFakid &" ORDER BY fak_sortorder " 
		        
                'if session("mid") = 1 then
		        'Response.Write strSQL
		        'Response.flush
                'end if
		        ermedpaaFakSQL = " AND ( j.id <> 0"
                a = a
		        oRec.open strSQL, oConn, 3
		       
		        While not oRec.EOF 

                fl_antal(a) = oRec("antal")
                fl_besk(a) = oRec("beskrivelse")
            
                if cint(oRec("showonfak")) = 1 then
                fl_vis(a) = "CHECKED"
                else
                fl_vis(a) = ""
                end if
    
                fl_enhpris(a) = oRec("enhedspris")
                fl_valuta(a) = oRec("valuta")
                thisaktfunc(a) = "red"
                fl_enhed(a) = oRec("enhedsang")
                fl_rabat(a) = oRec("rabat")
                fl_momsfri(a) = oRec("momsfri")
                fl_belob(a) = oRec("aktpris")
                fl_jobaktid(a) = oRec("aktid")

                ermedpaaFakSQL = ermedpaaFakSQL &" AND j.id <> " & oRec("aktid")

                a = a + 1
                oRec.movenext
                wend
                oRec.close

                ermedpaaFakSQL = ermedpaaFakSQL & ")"

                '** Resten
                if a <> 0 then
                sta_tomme = a
                else
                sta_tomme = 0
                end if
    

                'if session("mid") = 1 then
                'Response.Write "a: "&a& " sta_tomme " & sta_tomme & "showOld" & showOld & " tdato:  "&cDate(strTdato) & " dato:  "&cDate(strDato)
                'Response.flush
                'end if

                'ALLE linjer <> 0 timer/enheder bliver gemt sammen med fakturaen.
                strSQLaftalejob = "SELECT jobnavn, jobnr, id AS jid, valuta FROM job j WHERE serviceaft = "& aftid & " AND jobstatus <> 0 AND serviceaft <> 0 "& ermedpaaFakSQL


                        'if session("mid") = 1 then
                        'Response.write "strSQLaftalejob " & strSQLaftalejob
                        'end if
                        'Response.end

                        jai = a
                       
                        oRec.open strSQLaftalejob, oConn, 3
                        while not oRec.EOF 

                      
                        fl_jobaktid(jai) = oRec("jid")
                        fl_besk(jai) = oRec("jobnavn") & " ("& oRec("jobnr") &")"
        
                
                                '** Henter hovedlinje til gl. fakturaer på aftaler før 15/11-2010 **'
            
                                fl_antal(jai) = 0
                                fl_vis(jai) = ""
                                fl_enhpris(jai) = 0
                                fl_valuta(jai) = oRec("valuta")
                                thisaktfunc(jai) = "opr"
                                fl_enhed(jai) = 2
                                fl_rabat(jai) = intRabat/100 '0
                                fl_momsfri(jai) = 0
                                fl_belob(jai) = 0

                               

                

                            jai = jai + 1
                            oRec.movenext
                            wend
                            oRec.close 


                sta_tomme = sta_tomme + jai


                for e = sta_tomme to 50


                fl_antal(e) = 0
                fl_besk(e) = ""
                fl_vis(e) = ""
                fl_enhpris(e) = 0
                fl_valuta(e) = 1
                thisaktfunc(e) = "opr"
                fl_enhed(e) = 1
                fl_rabat(e) = 0
                fl_momsfri(e) = 0
                fl_belob(e) = 0
                fl_jobaktid(e) = 0



                next



else            


                if ja <> 0 then
                ja = ja
                else
                ja = 0
                end if
                
                for a = ja to 50
                fl_antal(a) = 0
                fl_besk(a) = ""
                fl_vis(a) = ""
                fl_enhpris(a) = 0
                fl_valuta(a) = 1
                thisaktfunc(a) = "opr"
                fl_enhed(a) = 2
                fl_rabat(a) = intRabat/100
                fl_momsfri(a) = 0
                fl_belob(a) = 0
                fl_jobaktid(a) = 0
                next

                

end if

%>


<input id="antal_x" name="antal_x" value="21" type="hidden" />
<%

'a = 0

'** Valuta_enheder Global **
    call enhval_gbl	
 

 %>
                  <table width=100% border=0 cellspacing=0 cellpadding=2>

                <tr><td bgcolor="#FFFFFF" colspan=9 style="padding:10px 200px 10px 5px;">
                    <h4>Fakturalinier</h4>
              

                 Varenr. <input type="text" name="FM_varenr" value="<%=strVarenr%>" style="font-size:9px; width:150px;"><br /><br />
                <b>Job tilknyttet denne aftale:</b><br />           
                <%=strJobPaaAft %>

                <%if func = "red" then %>
                &nbsp;(antal: <%=a %>/<%=jai %>)    
                <%else %>
                &nbsp;(antal: <%=ja %>)
                <%end if %>
                <br /><br />

                <%
                uWdt = 300
                uTxt = "Job tilknyttet denne aftale bliver lukket for timereg. frem til fakturadato. <br> Alle linjer med enheder/timer på bliver gemt sammen med fakturaen. Kun dem der er checked under 'vis' bliver vist på fakturaren"
                call infoUnisport(uWdt, uTxt) %>
                
             

                <%if func <> "red" AND intType <> 1 then
                
                  'aft_stk_pris()	            
                '** Enheder i periode **
                call akttyper2009(4)
                sumEnhVedFak = 0
                sumTimerEfterFak = 0
                strSQLenh = "SELECT sum(timer * faktor) AS sumEnhEfterFak, sum(timer) AS sumTimerEfterFak FROM timer "_
                &" LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE "_
                &" ("& replace(aty_sql_onfak, "aktiviteter", "a") &") AND seraft = "& aftid &""_ 
                &" AND tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"' GROUP BY seraft, taktivitetid"
	            
                'Response.Write strSQLenh &"<br>"
	            
                oRec3.open strSQLenh, oConn, 3
                while not oRec3.EOF
                
                sumEnhVedFak = sumEnhVedFak + oRec3("sumEnhEfterFak")
                sumTimerEfterFak = sumTimerEfterFak + oRec3("sumTimerEfterFak")
                
                oRec3.movenext
                wend
                oRec3.close
                 
                 Response.Write "<br><br><b>Forbrug i Periode:</b><br>"
                 

                call akttyper2009(2)
                sumEnhVedFak_fakbar = 0
                sumTimerEfterFak_fakbar = 0
                strSQLenh = "SELECT sum(timer * faktor) AS sumEnhEfterFak, sum(timer) AS sumTimerEfterFak FROM timer "_
                &" LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE "_
                &" ("& aty_sql_fakbar &") AND seraft = "& aftid &""_ 
                &" AND tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"' GROUP BY seraft, taktivitetid"
	            
                'Response.Write strSQLenh &"<br>"
	            
                oRec3.open strSQLenh, oConn, 3
                while not oRec3.EOF
                
                sumEnhVedFak_fakbar = sumEnhVedFak_fakbar + oRec3("sumEnhEfterFak")
                sumTimerEfterFak_fakbar = sumTimerEfterFak_fakbar + oRec3("sumTimerEfterFak")
                
                oRec3.movenext
                wend
                oRec3.close
                 
                 Response.Write "<b>Fakturerbare</b> (total)<br>"
                 Response.Write "Timer: <b>"& formatnumber(sumEnhVedFak_fakbar, 2) &"</b> ("& formatnumber(sumTimerEfterFak, 2) & ")<br>"
                 Response.Write "Enheder: <b>"& formatnumber(sumTimerEfterFak_fakbar, 2) &"</b> ("& formatnumber(sumEnhVedFak, 2) & ")<br>&nbsp;"

                 
                
                end if %>

                </td>
                </tr>
            </table>
                   <table width=100% border=0 cellspacing=0 cellpadding=2>
                

                        <tr bgcolor="#8CAAE6">
						
						<td class=lille style="padding:2px 2px 2px 5px;">Vis</td>
						<td class=lille style="width:70px;">Antal - #</td>
                        <td class=lille style="width:100px; padding:2px 2px 2px 5px;">Fakturalinie tekst</td>
				        <td class=lille style="padding:2px 2px 2px 5px;">Enh.pris</td>
						<td class=lille style="padding:2px 2px 2px 5px;">Valuta</td>
						<td class=lille style="padding:2px 2px 2px 5px;">Enhed</td>
						<td class=lille style="padding:2px 2px 2px 5px;" align=center>Rabat</td>
                        <td class=lille style="padding:2px 2px 2px 5px;" align=center>Momsfri</td>
						<td class=lille style="width:80px; padding:2px 2px 2px 5px;" align=right>Pris ialt&nbsp;&nbsp;&nbsp;</td>
						
					</tr>


<%

totalbelob = 0
strAktsubtotal = ""
a = 1
for x = 0 to 50
            

            select case right(x, 1)
            case 0,2,4,6,8
            bgthis = "#EFf3ff"
            case else
            bgthis = "#FFFFFF"
            end select

            strAktsubtotal = strAktsubtotal & "<tr bgcolor='"&bgthis&"'><td style=""border-bottom:1px #CCCCCC solid; padding:2px 0px 2px 2px;"" valign=top>"

           


           

            '*** Faste generelle ***'
            strAktsubtotal = strAktsubtotal & "<input id='antal_n_"&x&"' name='antal_n_"&x&"' value=""2"" type=""hidden"" />"
            strAktsubtotal = strAktsubtotal & "<input id='aktsort_"&x&"' name='aktsort_"&x&"' value='"&x&"' type=""hidden"" />"
            strAktsubtotal = strAktsubtotal & "<input id='aktId_n_"&x&"' name='aktId_n_"&x&"' value='"&fl_jobaktid(x)&"' type=""hidden"" />"
            strAktsubtotal = strAktsubtotal & "<input id='highest_aval_"&x&"' name='highest_aval_"&x&"' value=""1"" type=""hidden"" />"
            strAktsubtotal = strAktsubtotal & "<input id='antal_subtotal_akt_"&x&"' name='antal_subtotal_akt_"&x&"' value=""-1"" type=""hidden"" />"
            strAktsubtotal = strAktsubtotal & "<input id='FM_hidden_timerthis_"&x&"_"&a&"' name='FM_hidden_timerthis_"&x&"_"&a&"' value='"& fl_antal(x) &"' type=""hidden"" />"
            strAktsubtotal = strAktsubtotal & "<input id='FM_hidden_timepristhis_"&x&"_"&a&"' name='FM_hidden_timepristhis_"&x&"_"&a&"' value='"& fl_enhpris(x) &"' type=""hidden"" />"
            
            strAktsubtotal = strAktsubtotal & "<input id='timer_subtotal_akt_"&x&"' name='timer_subtotal_akt_"&x&"' value='"&fl_antal(x)&"' type=""hidden"" />"
            strAktsubtotal = strAktsubtotal & "<input id='belob_subtotal_akt_"&x&"' name='belob_subtotal_akt_"&x&"' value='"&fl_belob(x)&"' type=""hidden"" />"


            if fl_momsfri(x) = 1 then
            akt_belobUmons = fl_belob(x)
            else
            akt_belobUmons = 0
            end if

            strAktsubtotal = strAktsubtotal & "<input id='belob_subtotal_akt_umoms_"&x&"' name='belob_subtotal_akt_umoms_"&x&"' value='"&akt_belobUmons&"' type=""hidden"" />"
            '******
            
            
            if fl_vis(x) = "CHECKED" then
            totalbelob = totalbelob + fl_belob(x)
            else
            totalbelob = totalbelob 
            end if
            
            
             '** Vis **'
            strAktsubtotal = strAktsubtotal &"<input type=checkbox name='FM_show_akt_"&x&"_"&a&"' id='FM_show_akt_"&x&"_"&a&"' value='1' "&fl_vis(x)&" onClick='setBeloebTot2("&x&")'></td>"
			
            
            '*** Timer **'
            strAktsubtotal = strAktsubtotal &"<td style='width:95px; padding:2px 0px 2px 2px; border-bottom:1px #CCCCCC solid;' valign=top><input type=text name='FM_timerthis_"&x&"_"&a&"' id='FM_timerthis_"&x&"_"&a&"' onKeyup='tjektimer("&x&","&a&"), showborder2("&x&","&a&")' value='"&fl_antal(x)&"' size=3 style='!border: 1px; background-color: #FFFFFF; font-size:10px; border-color: yellowgreen; border-style: solid;'>"_
			& "&nbsp;<input type='button' name='beregn2_"&x&"_"&a&"_a' id='beregn2_"&x&"_"&a&"_a' value='Calc' style='font-size:8px' onClick='tjektimer("&x&","&a&"), setBeloebThis2("&x&","&a&"), hideborder2("&x&","&a&")';></td>"
		    'tjektimer("&x&","&a&"), andreEnhprs("&x&","&a&"), setBeloebThis2("&x&","&a&"), hideborder2("&x&","&a&")

            
            '** Besk ***'
            strAktsubtotal = strAktsubtotal &"<td valign=top style='width:230px; padding:2px 2px 2px 0px; border-bottom:1px #CCCCCC solid;'><textarea name='FM_aktbesk_"&x&"_"&a&"' id='FM_aktbesk_"&x&"_"&a&"' style='width:220px; height:40px; font-size:11px;'>"&fl_besk(x)&"</textarea></td>"
		
	
            
            '*** Enh pris ****'
            strAktsubtotal = strAktsubtotal & "<td valign=top align='right' style='width:60px; padding:1px 2px 2px 2px; border-bottom:1px #CCCCCC solid;'>"
			strAktsubtotal = strAktsubtotal &"<input type='text' name='FM_enhedspris_"&x&"_"&a&"' id='FM_enhedspris_"&x&"_"&a&"' onKeyup='tjektimepris("&x&"), showborder2("&x&","&a&")' value='"&fl_enhpris(x)&"'"_ 
			&" style='width:50px; font-size:9px;'></td>"	



            '*** Valuta ***'
            strAktsubtotal = strAktsubtotal & "<td align=center valign=top style='width:55; padding:2px 2px 2px 2px; border-bottom:1px #CCCCCC solid;'>"
		    strAktsubtotal = strAktsubtotal & "<select name='FM_valuta_"&x&"_"&a&"' id='FM_valuta_"&x&"_"&a&"' style='width:50px; font-size:9px;' onchange='tjektimer("&x&","&a&"), setBeloebThis2("&x&","&a&")'>"
		
		    strSQL5 = "SELECT id, valutakode, grundvaluta, kurs FROM valutaer ORDER BY valutakode"
                        		
            oRec5.open strSQL5, oConn, 3 
            while not oRec5.EOF 
    		
    		if thisaktfunc(x) = "red" then
    		
    		
                if oRec5("id") = fl_valuta(x) then
                valGrpCHK = "SELECTED"
                else
                valGrpCHK = ""
                end if
            
            else
            
                if oRec5("id") = valuta then
                valGrpCHK = "SELECTED"
                else
                valGrpCHK = ""
                end if
            
            end if
            
           
    		strAktsubtotal = strAktsubtotal & " <option value='"&oRec5("id")&"' "&valGrpCHK&">"&oRec5("valutakode")&"</option>"
           
            oRec5.movenext
            wend
            oRec5.close
            
            strAktsubtotal = strAktsubtotal & "</select></td>"

            strAktsubtotal = strAktsubtotal & "<input name='FM_valuta_opr_"&x&"_"&a&"' id='FM_valuta_opr_"&x&"_"&a&"' type='hidden' value='"&valutaId&"'/>"
		    


            '** Enheder ***'
            'if func <> "red" then
            'intEnhedsang = intEnhedsang 
            'else
            intEnhedsang = fl_enhed(x)
            'end if
            
                     
                    '****** Enhedsangiverlse ****'
                    select case intEnhedsang
                    case -1
                    ehSel_none = "SELECTED"
                    ehSel_nul = ""
		            ehSel_et = ""
		            ehSel_to = ""
		            ehSel_tre = ""
		            ehLabel = "Ingen"
		            case 0
		            ehSel_none = ""
		            ehSel_nul = "SELECTED"
		            ehSel_et = ""
		            ehSel_to = ""
		            ehSel_tre = ""
		            ehLabel = "Pr. time"
		            case 1
		            ehSel_none = ""
		            ehSel_nul = ""
		            ehSel_et = "SELECTED"
		            ehSel_to = ""
		            ehSel_tre = ""
		            ehLabel = "Pr. stk."
		            case 2
		            ehSel_none = ""
		            ehSel_nul = ""
		            ehSel_et = ""
		            ehSel_to = "SELECTED"
		            ehSel_tre = ""
		            ehLabel = "Pr. enhed"
		            case 3
		            ehSel_none = ""
		            ehSel_nul = ""
		            ehSel_et = ""
		            ehSel_to = ""
		            ehSel_tre = "SELECTED"
		            ehLabel = "Pr. km."
		            case else
		            ehSel_none = ""
		            ehSel_nul = "SELECTED"
		            ehSel_et = ""
		            ehSel_to = ""
		            ehSel_tre = ""
		            ehLabel = "Pr. time"
		            end select




                    strAktsubtotal = strAktsubtotal & "<td valign=top align=right style='width:65px; padding:2px 2px 2px 2px; border-bottom:1px #CCCCCC solid;'>"
		            strAktsubtotal = strAktsubtotal &"<select name='FM_akt_enh_"&x&"_"&a&"' id='FM_akt_enh_"&x&"_"&a&"' style=""width:60px; font-size:9px;"">"_
                    &"<option value='0' "& ehSel_nul &">Pr. time</option>"_
                    &"<option value='1' "& ehSel_et &">Pr. stk.</option>"_
                    &"<option value='2' "& ehSel_to &">Pr. enhed</option>"_
                    &"<option value='3' "& ehSel_tre &">Pr. km.</option>"_
                    &"<option value='-1' "& ehSel_none &">Ingen</option>"_
                    &"</select></td>"


            '** Rabat **'

            thisakt_rabat = fl_rabat(x) 

            akt_rSel0 = ""
            akt_rSel10 = ""
            akt_rSel15 = ""
            akt_rSel25 = ""
            akt_rSel30 = ""
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = ""

            select case (thisakt_rabat*100)
            case 0
		    akt_rSel0 = "SELECTED"
            case 10
            akt_rSel10 = "SELECTED"
            case 15
            akt_rSel15 = "SELECTED"
            case 25
            akt_rSel25 = "SELECTED"
            case 30
            akt_rSel30 = "SELECTED"
            case 40
            akt_rSel40 = "SELECTED"
            case 50
            akt_rSel50 = "SELECTED"
            case 60
            akt_rSel60 = "SELECTED"
            case 75
            akt_rSel75 = "SELECTED"
		    end select
        
        strAktsubtotal = strAktsubtotal &"<td valign=top align='right' style='width:50px; padding:2px 0px 0px 0px; border-bottom:1px #CCCCCC solid;'><select id='FM_rabat_"&x&"_"&a&"' name='FM_rabat_"&x&"_"&a&"'  style='background-color:#ffffff; width:40px; font-size:9px;' onChange='setBeloebThis2("&x&","&a&")'>"
        strAktsubtotal = strAktsubtotal &"<option value='0' "&akt_rSel0&">0%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.10' "&akt_rSel10&">10%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.15' "&akt_rSel15&">15%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.25' "&akt_rSel25&">25%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.30' "&akt_rSel30&">30%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.40' "&akt_rSel40&">40%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.50' "&akt_rSel50&">50%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.60' "&akt_rSel60&">60%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.75' "&akt_rSel75&">75%</option></select></td>"
            					   
        
        '**** Moms ****'
        if fl_momsfri(x) <> 0 then
        momsfri_CHK = "CHECKED"
                if fl_vis(x) = "CHECKED" then
                totalbelob_umoms = totalbelob_umoms + fl_belob(x)
                else
                totalbelob_umoms = totalbelob_umoms 
                end if
        else
        totalbelob_umoms = totalbelob_umoms 
        momsfri_CHK = ""
        end if

        strAktsubtotal = strAktsubtotal & "<td valign=top align='center' style=""width:50px; padding:2px 2px 2px 2px; border-bottom:1px #CCCCCC solid;""><input id='FM_momsfri_"&x&"_"&a&"' name='FM_momsfri_"&x&"_"&a&"' type=checkbox value=""1"" "&momsfri_CHK&" onClick=""setBeloebTot2("&x&");""/></td>"
		

        '**** Beløb ***' 
		strAktsubtotal = strAktsubtotal & "<td valign=top style='width:110px; padding:2px 5px 0px 2px; border-bottom:1px #CCCCCC solid;'>" 
		
			strAktsubtotal = strAktsubtotal &"<div align='right' name='belobdiv_"&x&"_"&a&"' id='belobdiv_"&x&"_"&a&"'><b>"& fl_belob(x) &" "& valutaKodeSel &"</b></div>"
			strAktsubtotal = strAktsubtotal &"<input type='hidden' name='FM_beloebthis_"&x&"_"&a&"' id='FM_beloebthis_"&x&"_"&a&"' value='"& fl_belob(x) &"'></td>"
		
		
        'strAktsubtotal = strAktsubtotal &"<td align=right id=timeae_akt_"&x&" class=lille style=""border-bottom:1px #cccccc solid;"">&nbsp;</td>"

		'*** Afslut tabel *****'
		strAktsubtotal = strAktsubtotal &"</tr>"

next


'*** Afslut tabel *****'
strAktsubtotal = strAktsubtotal &"</table>"



Response.Write strAktsubtotal
%>



</div>

<div id=aktdiv_2 style="position:absolute; visibility:hidden; display:none; top:1336px; width:720px; left:5px; border:0px #8cAAe6 solid;">
<table width=100% cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('modtagdiv')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('betdiv')" class=vmenu>Næste >></a></td></tr></table>
 </div>



 <div id="aktsubtotal" style="position:absolute; left:730px; top:104px; width:200px; z-index:2000; border:1px yellowgreen solid; background-color:#ffffff; padding:5px;">
    <b>Fakturalinier:</b><br />
	<table cellspacing=5 cellpadding=0 border=0 width=100%><tr>
	<td><!--Antal ialt:-->
	<%
	'*********************************************************
	'Subtotal
	'Total timer og beløb for aktiviteter
	'*********************************************************
	if intTimer <> 0 then
	intTimer = intTimer 
	else
	intTimer = 0
	end if
	
	%><input type="hidden" name="FM_timer" id="FM_Timer" value="0"><!--=intTimer-->
	<div style="position:relative; width:45; height:20px; border-bottom:0px YellowGreen dashed; background-color:#ffffff; padding-right:3px;" align="right" id="divtimertot">
	<!--<b><=intTimer%></b>--></div>
	</td>
	<td align="right">Subtotal:
		<!-- strBeloeb -->
		<%
		if len(totalbelob) <> 0 then
		thistotbelTimer = SQLBlessDot(formatnumber(totalbelob, 2))
		else
		thistotbelTimer = formatnumber(0, 2)
		end if
		
		if len(totalbelob_umoms) <> 0 then
		thistotbelTimer_umoms = SQLBlessDot(formatnumber(totalbelob_umoms, 2))
		else
		thistotbelTimer_umoms = formatnumber(0, 2)
		end if
		
		
		
		
		thistotbel = thistotbelTimer%>
		<input type="hidden" name="FM_timer_beloeb" id="FM_timer_beloeb" value="<%=thistotbelTimer%>">
		<input type="hidden" name="FM_timer_beloeb_umoms" id="FM_timer_beloeb_umoms" value="<%=thistotbelTimer_umoms%>">
		
		<div style="position:relative; width:95px; height:20px; border-bottom:2px YellowGreen dashed; background-color:#ffffff; padding-right:3px;" align="right" id="divtimerbelobtot"><b><%=thistotbel &" "& valutakodeSEL %></b></div>
		
		
		</td>
		</tr>
	</table>
	</div>