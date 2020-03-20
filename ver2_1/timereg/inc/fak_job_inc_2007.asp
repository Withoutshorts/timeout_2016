
 <!-- Aktiviteter -->
	    




 <div id="aktdiv" style="position:absolute; visibility:hidden; display:none; z-index:2000; left:5px; top:105px; width:700px; border:1px yellowgreen solid; padding:5px 5px 5px 5px; background-color:#ffffff;">
    <table width=100% border=0 cellspacing=0 cellpadding=5>
	<tr>
	<td valign=top><h4><%=erp_txt_314 %> (<%=erp_txt_315 %>) <br />
        <span style="font-size:11px; font-weight:normal; color:#999999;"><%=erp_txt_316 %></span>

         <% response.Flush

    if cint(pcSpecial) = 1 then '1pc Grundfos, forvalgt antal timer altid sat til 1 
    %>
        <br /><span style="font-size:11px; font-weight:normal; color:red;"><%=erp_txt_317 %></span>

    <%
    end if
        %></h4>
	<%

    tblWdtHeader = "<tr bgcolor=""#FFFFFF"">"
    tblWdtHeader = tblWdtHeader & "<td><img src=""../ill/blank.gif"" width=""40"" height=""1"" border=""0""></td>"
    tblWdtHeader = tblWdtHeader & "<td><img src=""../ill/blank.gif"" width=""180"" height=""1"" border=""0""></td>"
    tblWdtHeader = tblWdtHeader & "<td><img src=""../ill/blank.gif"" width=""30"" height=""1"" border=""0""></td>"
    tblWdtHeader = tblWdtHeader & "<td><img src=""../ill/blank.gif"" width=""30"" height=""1"" border=""0""></td>"
    tblWdtHeader = tblWdtHeader & "<td><img src=""../ill/blank.gif"" width=""40"" height=""1"" border=""0""></td>"
    tblWdtHeader = tblWdtHeader & "<td><img src=""../ill/blank.gif"" width=""40"" height=""1"" border=""0""></td>"
    tblWdtHeader = tblWdtHeader & "<td><img src=""../ill/blank.gif"" width=""60"" height=""1"" border=""0""></td>"
    tblWdtHeader = tblWdtHeader & "<td><img src=""../ill/blank.gif"" width=""60"" height=""1"" border=""0""></td>"
    tblWdtHeader = tblWdtHeader & "<td><img src=""../ill/blank.gif"" width=""60"" height=""1"" border=""0""></td>"
    tblWdtHeader = tblWdtHeader & "<td><img src=""../ill/blank.gif"" width=""40"" height=""1"" border=""0""></td>"
    tblWdtHeader = tblWdtHeader & "</tr>"

    '** Tjekker om der ligger timer på lukkede/passive aktiviteter ***
    '** Hvis ja vises lukkede og passive ellers vises kun åbne ***'
    strSQL2 = "SELECT aktiviteter.navn, aktiviteter.id, aktstatus, sum(t.timer) AS timerthis FROM aktiviteter "_
    &" LEFT JOIN timer AS t ON (t.taktivitetid = aktiviteter.id) "_
    &" WHERE aktstatus <> 1 AND (aktiviteter.job = "& jobid &") AND ("& aty_sql_onfak &") AND (t.tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"') GROUP BY t.taktivitetid"

    'Response.Write strSQL2
    'Response.flush
    timerPaLukkeogPassiveAkt = 0
    oRec2.open strSQL2, oConn, 3
    timerPaLukkeogPassiveAkt = 0                						
    while not oRec2.EOF 
	timerPaLukkeogPassiveAkt = timerPaLukkeogPassiveAkt + oRec2("timerthis")   
    oRec2.movenext
    wend 
    oRec2.close
				
                if cdbl(timerPaLukkeogPassiveAkt) <> 0 then	
                %>
                <span style="background-color:lightpink; width:500px; padding:3px;"><%=erp_txt_318 %> 
                <%=erp_txt_319 %></span><br /><br />
                <%
                end if	                

   
	
	'******************************************************************************************'
	'*** Henter Medarbejdere og Aktiviteter ***************************************************'
	'******************************************************************************************'
	        
	        
	        isaktwritten = " aktiviteter.id <> 0 "
	        x = 0
	        incidentlistName = "incidentlist"
	        
	        if func = "red" then
        		
		        isTPused = "#"
        		
		        '*******************************************************************************'
		        '******************* Henter oprettede aktiviteter ******************************'
		        '*******************************************************************************'
		        strSQL = "SELECT faktura_det.id, aktid, aktiviteter.navn, aktiviteter.fakturerbar, aktiviteter.faktor, "_
		        &" faktura_det.valuta, faktura_det.beskrivelse, fak_sortorder, enhedsang, faktura_det.fase AS fase, momsfri, "_
		        &" aktiviteter.bgr, aktiviteter.aktbudget, aktiviteter.aktbudgetsum, aktiviteter.budgettimer, aktiviteter.antalstk, faktura_det.enhedspris"_
			    &" FROM faktura_det "_
		        &" LEFT JOIN aktiviteter ON (aktiviteter.id = aktid) WHERE fakid = "& intFakid &" GROUP BY aktid ORDER BY fase, fak_sortorder" 
		        
		        'Response.Write strSQL
		        'Response.flush
		        
		        oRec.open strSQL, oConn, 3
		       
		        While not oRec.EOF 
        		
			       thisaktnavn(x) = oRec("navn") 'oRec("beskrivelse")
			       thisaktid(x) = oRec("aktid")
			       thisaktfunc(x) = "red"
			       thisaktfakbar(x) = oRec("fakturerbar")
			       thisaktfaktor(x) = oRec("faktor")
			       thisAktText(x) = oRec("beskrivelse") 
			       thisAktValuta(x) = oRec("valuta")
			       thisaktsort(x) = oRec("fak_sortorder") 
			       thisaktEnh(x) = oRec("enhedsang")
			       thisaktCHK(x) = 1     
			       thisAktTimer(x) = 0
                   
			       thisAktFase(x) = oRec("fase")
			       thisMomsfri(x) = oRec("momsfri")
			       
			       thisAktBudgetsum(x) = oRec("aktbudgetsum")
			       thisAktForkalk(x) = oRec("budgettimer")
				   thisAktForkalkStk(x) = oRec("antalstk")

                   thisAktEnhpris(x) = oRec("enhedspris")
				   
				   thisAktBgr(x) = oRec("bgr") 
				   select case thisAktBgr(x)
					case 0
					aktbgNavn = "&nbsp;"
					case 1
					aktbgNavn = "Tim."
					case 2
					aktbgNavn = "Stk."
					end select
			        
			        
			        '*** Antal timer total på denne akt. ***'
			        strSQLfantal = "SELECT sum(faktura_det.antal) AS antal, sum(faktura_det.aktpris) AS aktprissum FROM faktura_det "_
			        &" WHERE fakid = "& intFakid &" AND aktid = "& oRec("aktid") & " GROUP BY aktid"
			        oRec2.open strSQLfantal, oConn, 3
			        while not oRec2.EOF
			            thisAktTimer(x) = oRec2("antal")
			            thisAktBeloeb(x) = oRec2("aktprissum")
			        oRec2.movenext
			        wend
			        oRec2.close
			       
			       
                   thisAktTimerSum(x) = thisAktTimer(x)
        			
        		
        		isaktwritten = isaktwritten & " AND aktiviteter.id <> "& thisaktid(x) &"" 
        		
        		select case right(x, 1)
        		case 2,4,6,8,0
        		trbg = "#FFFFFF"
        		case else
        		trbg = "#FFFFFF"
        		end select
				
				
				    '** Marker lbn / fastpris kolonne   0: lbn timer / 1: fastpris / 2: Commi / 3: Salesorder
	                if cint(usefastpris) = 0 then
	                bgfastpr = ""
	                bdfastim = 0
	                bglbntim = ""
	                bdlbntim = 1
	                else
	                bgfastpr = ""
	                bdfastim = 1
	                bglbntim = ""
	                bdlbntim = 0
	                end if
				
				
				if lcase(lastFase) <> lcase(thisAktFase(x)) AND len(trim(thisAktFase(x))) <> 0 then
				strAktAnchor = strAktAnchor & "</tr></table><table width=100% cellspacing=0 cellpadding=2 border=0 bgcolor=""#D6DFf5"" id=""incidentlist_"&x&""">"
                strAktAnchor = strAktAnchor & tblWdtHeader
                    
                strAktAnchor = strAktAnchor & "<tr bgcolor=""#D6Dff5"" class=""tr_fase""><td colspan=10 class=lille>&nbsp;"
				strAktAnchor = strAktAnchor & "<b>"& replace(thisAktFase(x), "_", " ") &"</b></td></tr>"

                incidentlistName = "incidentlist_"& x
				end if
				
				'strAktAnchor = strAktAnchor & "<tr class=""traktlink"" id=""tr_akt_"&thisaktid(x)&""" bgcolor="& trbg  &"><td style=""border-bottom:1px #cccccc solid;"" class=lille>"
				'strAktAnchor = strAktAnchor & "<a id='"& thisaktid(x) &"' href='#"& thisaktid(x) &"' class='qlinks'>"& thisaktnavn(x) & ": </a></td><td style=""border-bottom:1px #cccccc solid;"" align=right class=lille align=right>"& formatnumber(thisAktTimer(x), 2) &""_
				'&"</td><td style=""border-bottom:1px #cccccc solid;"" id=timeae_akt_"&thisaktid(x)&" align=right class=lille>&nbsp;</td></tr>"
				

                     
		
				
				    strAktAnchor = strAktAnchor & "<tr class=""traktlink"" id=""tr_akt_"&thisaktid(x)&""" bgcolor="& trbg  &"><td class=lille style=""border-bottom:1px #cccccc solid;"">"

                    
                    strAktAnchor = strAktAnchor & "<input type=""hidden"" id=""tblname_"& thisaktid(x) &""" value="& incidentlistName &" />"

                    strAktAnchor = strAktAnchor & "<input type=""hidden"" name=""SortOrder"" class=""SortOrder"" value="& thisaktsort(x) &" />"
		            strAktAnchor = strAktAnchor & "<input type=""hidden"" name=""rowId"" value='F"& thisaktid(x) &"' />"
                    strAktAnchor = strAktAnchor & "<input type=""hidden"" id='altsort_F"& thisaktid(x) &"' value=""1"" />"

                    strAktAnchor = strAktAnchor & "<input type=""checkbox"" class=""ch_akt_x"" id=""ch_akt_"&x&""" CHECKED>&nbsp;<img src=""../ill/pile_drag.gif"" alt="""& erp_txt_503 &""" border=""0"" id='drag_"&x&"' class=""drag"" />&nbsp;<span id='sort_F"& thisaktid(x) &"' style='color:#999999; font-size:7px;'>" & thisaktsort(x) & "</span>&nbsp;" 
					strAktAnchor = strAktAnchor & "</td><td style=""border-bottom:1px #cccccc solid; font-size:9px; white-space:inherit; width:200px;""><a id='"& thisaktid(x) &"' href='#"& thisaktid(x) &"' class='qlinks'>"& thisaktnavn(x) & "</a></td>"_
					&"<td style=""border-bottom:1px #cccccc solid; font-size:9px;"" align=center>"& aktbgNavn &"</td>"_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid;"">"& formatnumber(thisAktForkalk(x), 2) &"</td>"_
                    &"<td align=right class=lille style=""border-bottom:1px #cccccc solid;"">"& formatnumber(thisAktForkalkStk(x), 2) &" stk.</td>"_
                    &"<td align=right class=lille style=""border-bottom:1px #cccccc solid;"">"& formatnumber(thisAktEnhpris(x), 2) &"</td>"_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid; border-left:"&bdfastim&" solid #999999; border-right:"&bdfastim&" solid #999999; background-color:"& bgfastpr&";"">"& formatnumber(thisAktBudgetsum(x), 2) &"</td>"_
					&"<td style=""border-bottom:1px #cccccc solid; white-space:nowrap;"" align=right class=lille>"& formatnumber(thisAktTimer(x), 2) 

                    if thisAktTimer(x) <> 0 AND thisAktBeloeb(x) <> 0 then 
                    gnsTp = thisAktBeloeb(x)/thisAktTimer(x) 
                    else
                    gnsTp = 0
                    end if
                    
                    strAktAnchor = strAktAnchor & "<span style=""font-size:8px; color:#999999;""> ~ " & formatnumber(gnsTp, 2) & "</td>"

					strAktAnchor = strAktAnchor &"<td align=right class=lille style=""border-bottom:1px #cccccc solid; border-left:"&bdlbntim&" solid #999999; border-right:"&bdlbntim&" solid #999999; background-color:#DCF5BD;"">"& formatnumber(thisAktBeloeb(x), 2) &"</td>"_
					&"</td><td align=right id=timeae_akt_"&thisaktid(x)&" class=lille style=""border-bottom:1px #cccccc solid;"">&nbsp;</td></tr>"
				
						
        		lastFase = oRec("fase")	
		        x = x + 1
        		
		        oRec.movenext
		        wend 
		        oRec.close
        	
	        end if
	        
	        
	        
	        
	        
	        
			
		    '*******************************************************************************'
		    '*** Henter alle de aktiviteter der ikke allerede er oprettet (ved rediger) ****'
		    '*** Henter alle aktiviteter ved opr. ******************************************'
		    '*******************************************************************************'
			'ign_akttype_inst ignorer tpye
			'*** Henter fakbare, Ikke fakbare & Salg **'
			
			'**** Sidste faktura tidspunkt ****
			if len(strfaktidspkt) > 0 then
			LastFakTid = datepart("h", strfaktidspkt)&":"& datepart("n", strfaktidspkt)&":"& datepart("s", strfaktidspkt)
			else
			LastFakTid =  "00:00:02"
			end if
			
            
            incidentlistName = "incidentlist"
			
			'*** Henter aktiviter og finder registrede timer **'
			'*** siden sidste faktura dato. (periodeafgrænsning) ***'
			
            'vis lukkede og passive
            if cdbl(timerPaLukkeogPassiveAkt) <> 0 then
            sqlAktStatusKri = " aktstatus <> -1" 
			else
            sqlAktStatusKri = " aktstatus = 1"
            end if

            if cint(ign_akttype_inst) = 1 then
            aty_sql_onfak = ""
            else
            aty_sql_onfak = "AND ("& aty_sql_onfak &")"
            end if 
			
			strSQL = "SELECT job.fastpris, job.jobtpris, job.budgettimer AS jbudgettimer, job.usejoborakt_tp, "_
			&" aktstatus, aktiviteter.job, aktiviteter.id AS aid, "_
			&" aktiviteter.navn AS anavn, aktiviteter.fakturerbar, aktiviteter.faktor, aktiviteter.fase, aktiviteter.aktstatus,  "_
			&" aty_on_invoice_chk, aty_enh, "_
			&" aktiviteter.bgr, aktiviteter.aktbudget, aktiviteter.aktbudgetsum, aktiviteter.budgettimer, aktiviteter.antalstk, aktiviteter.sortorder"_
			&" FROM aktiviteter LEFT JOIN job ON (job.id = aktiviteter.job) "_
			&" LEFT JOIN akt_typer aty ON (aty_id = aktiviteter.fakturerbar) "_
			&" WHERE (aktiviteter.job = "& jobid &") AND ("& sqlAktStatusKri &") "& aty_sql_onfak &" AND ("& isaktwritten &") "_
			&" ORDER BY fase, aktiviteter.sortorder, aktiviteter.navn"	
			'AND aktiviteter.aktstatus = 1, '"& aty_sql_onfak &"
			
			'Response.Write "ign_akttype_inst: "& ign_akttype_inst &"<br>"& strSQL 
			Response.flush
			
			oRec.open strSQL, oConn, 3
			
			
			lastAktId = 0
			timerThisT = 0
			timerThisD = 0
			'strKom = "Netto Kontant 8 dage." Hentes på kunde
			totaltimer = 0
			lastaktid = 0
            ir = 0
			
			
			while not oRec.EOF
			
			thisJoborAkt_tp(x) = 0 'oRec("usejoborakt_tp")	   '**altid job (men viser forkalk på aktiviteter hvis <> 0 and job = fastpris) 6-12-2010
				    
				   if cdbl(lastaktid) <> cdbl(oRec("aid")) then
						
						
						thisaktid(x) = oRec("aid")
					    usefastpris = oRec("fastpris")
						
						   
								        
								        
								       
							            thisAktBeloeb(x) = 0
							            
							            '*** Skal ikke hentes ved kreditnota **'
							            if cint(intType) <> 1 then
						               
						                    '*** Realisrede timer ***'
						                    '** antal registerede timer pr. akt. **'
                                            select case lto 
                                            case "intranet - local", "bf" 
                                            '*** Aktivitetsspecifik dato. Kun Hvis den har været på faktura
                                                                                         

                                                        'if cint(oRec("aid")) = 6015 then
                                                        stdatoKriAktSpecifik(x) = stdatoKri '"2016/01/01"
                                                        strSQLAktDatospecifikSQL = "SELECT fakid, fakdato FROM faktura_det LEFT JOIN fakturaer ON (fid = fakid) WHERE aktid = "& oRec("aid") &" ORDER BY fakdato DESC" 
                                                        oRec6.open strSQLAktDatospecifikSQL, oConn, 3

                                                        'response.write strSQLAktDatospecifikSQL
                                                        'response.flush

                                                        if not oRec6.EOF then
                                                            stdatoKriAktSpecifikThis = dateAdd("d", 1, oRec6("fakdato"))
                                                            stdatoKriAktSpecifik(x) = year(stdatoKriAktSpecifikThis) & "/" & month(stdatoKriAktSpecifikThis) & "/"& day(stdatoKriAktSpecifikThis)
                                                        end if        
                                                        oRec6.close
                                                        
                                                        'response.write "stdatoKriAktSpecifik: " & stdatoKriAktSpecifik & "<br>"
                                                        
                                                        'else
                                                        'stdatoKriAktSpecifik = stdatoKri
                                                        'end if 

                                            case else

                                                        stdatoKriAktSpecifik(x) = stdatoKri

                                            end select

                            
                                            select case lto
                                            case "bf", "intranet - local" '** ONLY med_lincensindehaver 
                                                call alleMedIJurEnhed(session("med_lincensindehaver"), 0)
                                            case else
                                                alleMedIJurEnhedSQL = ""
                                            end select

                                           

						                    strSQL2 = "SELECT sum(t.timer) AS timerthis, sum(t.timer*t.timepris) AS beloeb FROM timer AS t "_
						                    &" WHERE t.taktivitetid =" & oRec("aid") &" AND (t.tdato BETWEEN '" & stdatoKriAktSpecifik(x) &"' AND '"& slutdato &"') "& alleMedIJurEnhedSQL &" GROUP BY t.taktivitetid"

                                            'Response.Write strSQL2
                                            'Response.flush
						                    oRec2.open strSQL2, oConn, 3
                    						
						                    if not oRec2.EOF then
							                    timerThisD = oRec2("timerthis")
							                    thisAktBeloeb(x) = oRec2("beloeb")
						                    end if
						                    oRec2.close
						                
						                end if
                						
							                if len(timerThisD) <> 0 then
							                timerThisD = timerThisD 
							                else
							                timerThisD = 0
							                end if
							
							
							
							
					
					thisAktEnhpris(x) = oRec("aktbudget")
					thisTimePris(x) = thisAktEnhpris(x)	
					akttimepris = thisTimePris(x)
                    	
					thisaktnavn(x) = oRec("anavn") 
					thisAktText(x) = thisaktnavn(x)
					
					
					
					thisAktTimer(x) = timerThisD
					thisAktTimerSum(x) = thisAktTimerSum(x) +  timerThisD
                    
					
					thisAktBudgetsum(x) = oRec("aktbudgetsum")
					thisaktfunc(x) = "opr"
					thisaktfakbar(x) = oRec("fakturerbar")
					thisaktfaktor(x) = oRec("faktor") 
					thisAktValuta(x) = valuta 'fra job
					thisaktEnh(x) = oRec("aty_enh")
					thisaktCHK(x) = oRec("aty_on_invoice_chk")
					
					timerThisD = 0
					thisBeloebD = 0
					
					thisAktFase(x) = oRec("fase")
					
					if thisaktfakbar(x) <> 5 then '** Kørsel
					thisMomsfri(x) = 0
					else
					thisMomsfri(x) = 1
					end if
					
					thisAktForkalk(x) = oRec("budgettimer")
					thisAktForkalkStk(x) = oRec("antalstk")
					thisAktBgr(x) = oRec("bgr") 
					
					select case thisAktBgr(x)
					case 0
					aktbgNavn = "&nbsp;"
					case 1
					aktbgNavn = "Tim."
					case 2
					aktbgNavn = "Stk."
					end select
					
					select case right(x, 1)
        		    case 2,4,6,8,0
        		    trbg = "#FFFFFF"
        		    case else
        		    trbg = "#FFFFFF"
        		    end select
        		    
        		    '** Maker lbn / fastpris kolonne
	                if cint(usefastpris) = 0 then
	                bgfastpr = ""
	                bdfastim = 0
	                bglbntim = ""
	                bdlbntim = 1
	                else
	                bgfastpr = ""
	                bdfastim = 1
	                bglbntim = ""
	                bdlbntim = 0
	                end if



                    if ir = 0 AND func = "red" then
                    strAktAnchor = strAktAnchor & "<tr bgcolor=""#FFFFFF""><td colspan=10 class=lille>"
				    strAktAnchor = strAktAnchor & "<br><b>"& erp_txt_322 &":</b></td></tr>"
                    end if
				
					if (lcase(lastFase) <> lcase(thisAktFase(x)) AND len(trim(thisAktFase(x))) <> 0) OR (len(trim(lastFase)) = 0 AND len(trim(thisAktFase(x))) <> 0)  then

                    strAktAnchor = strAktAnchor & "</tr></table><table width=100% cellspacing=0 cellpadding=2 border=0 bgcolor=""#D6DFf5"" id=""incidentlist_"&x&""">"
                    strAktAnchor = strAktAnchor & tblWdtHeader

                    strAktAnchor = strAktAnchor & "<tr bgcolor=""#D6Dff5"" class=""tr_fase""><td colspan=10 class=lille>&nbsp;"
				    strAktAnchor = strAktAnchor & "<b>"& replace(thisAktFase(x), "_", " ") &"</b></td></tr>"

                    thisaktsort(x) = 100+x 'oRec("sortorder")

                    incidentlistName = "incidentlist_"& x 

                    else

                    if lastSort <> "" then
                    thisaktsort(x) = lastSort
                    else
                    thisaktsort(x) = 1
                    end if
					
				    end if
					

                    select case oRec("aktstatus") 
                    case 0
                    aktstTXT = " ("& erp_txt_323 &")"
                    case 2
                    aktstTXT = " ("& erp_txt_324 &")"
                    case else
                    aktstTXT = "" 
                    end select



                   


					strAktAnchor = strAktAnchor & "<tr class=""traktlink"" id=""tr_akt_"&thisaktid(x)&""" bgcolor="& trbg  &"><td class=lille style=""border-bottom:1px #cccccc solid;"">"
                    strAktAnchor = strAktAnchor & "<input type=""hidden"" id=""tblname_"& x &""" value="& incidentlistName &" />"

                    strAktAnchor = strAktAnchor & "<input type=""hidden"" name=""SortOrder"" class=""SortOrder"" value="& thisaktsort(x) &" />"
		            strAktAnchor = strAktAnchor & "<input type=""hidden"" name=""rowId"" value='A"& thisaktid(x) &"' />"
                    strAktAnchor = strAktAnchor & "<input type=""hidden"" id='altsort_A"& thisaktid(x) &"' value=""0"" />"
		
                    
                    strAktAnchor = strAktAnchor & "<input type=""checkbox"" class=""ch_akt_x"" id=""ch_akt_"&x&""">&nbsp;<img src=""../ill/pile_drag.gif"" alt="""& erp_txt_503 &""" border=""0"" id='drag_"&x&"' class=""drag"" />&nbsp;<span id='sort_A"&thisaktid(x)&"' style='color:#999999; font-size:7px;'>" & thisaktsort(x) & "</span>&nbsp;"
					strAktAnchor = strAktAnchor & "</td><td style=""border-bottom:1px #cccccc solid; font-size:9px; white-space:inherit; width:200px;""><a id='"& thisaktid(x) &"' href='#"& thisaktid(x) &"' class='qlinks'>"& thisaktnavn(x) & "</a>"& aktstTXT &"</td>"_
					&"<td style=""border-bottom:1px #cccccc solid; font-size:9px;"" align=center>"& aktbgNavn &"</td>"_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid;"">"& formatnumber(thisAktForkalk(x), 2) &"</td>"_
                    &"<td align=right class=lille style=""border-bottom:1px #cccccc solid;"">"& formatnumber(thisAktForkalkStk(x), 2) &" stk.</td>"_
                    &"<td align=right class=lille style=""border-bottom:1px #cccccc solid;"">"& formatnumber(thisAktEnhpris(x), 2) &"</td>"_
					&"<td align=right class=lille style=""border-bottom:1px #cccccc solid; border-left:"&bdfastim&" solid #999999; border-right:"&bdfastim&" solid #999999; background-color:"& bgfastpr&";"">"& formatnumber(thisAktBudgetsum(x), 2) &"</td>"_
					&"<td style=""border-bottom:1px #cccccc solid; white-space:nowrap;"" align=right class=lille>"& formatnumber(thisAktTimer(x), 2) 

                    if thisAktTimer(x) <> 0 AND thisAktBeloeb(x) <> 0 then 
                    gnsTp = thisAktBeloeb(x)/thisAktTimer(x) 
                    else
                    gnsTp = 0
                    end if
                    strAktAnchor = strAktAnchor & "<span style=""font-size:8px; color:#999999;""> ~ " & formatnumber(gnsTp, 2) & "</td>"
                    

					strAktAnchor = strAktAnchor &"<td align=right class=lille style=""border-bottom:1px #cccccc solid; border-left:"&bdlbntim&" solid #999999; border-right:"&bdlbntim&" solid #999999; background-color:"& bglbntim&";"">"& formatnumber(thisAktBeloeb(x), 2) &"</td>"_
					&"</td><td align=right id=timeae_akt_"&thisaktid(x)&" class=lille style=""border-bottom:1px #cccccc solid;"">&nbsp;</td></tr>"
				
				
                if len(trim(thisAktFase(x))) <> 0 then	
			    lastFase = thisAktFase(x) 'oRec("fase")	
                else
                lastFase = ""
                end if

                lastSort = thisaktsort(x)
                
                ir = ir + 1
				x = x + 1		
				end if
			
          
    
			lastaktid = oRec("aid")		
	
            Response.flush        
            oRec.movenext
			wend
			oRec.close
			
			
	'end if
	'***********************************************************************************'
	
	
	'Response.flush
	
	
	'***********************************************************************************'
	'********************* Udskriver aktiviterer ***************************************'
	'***********************************************************************************'
	'uTxt = "<font class=lillesort><b>Fastpris job:</b> Samle-sumaktiviteten er forvalgt og forkalkulation på aktiviteter benyttes. Alle akt. linier er slået fra ved oprettelse af faktura på fastpris job.<br><br>"_
	'&"<b>Løbende timer job:</b> Realiserede timer og omsætning pr. medarbejder er forvalgt på faktura linier. <br>Hvis grundlag på aktivitet er sat til stk., og der ikke er angivet en stk. pris, er samle sum-aktiviteten forvalgt og <b>enhedsprisen beregnet udfra timeforbrug.</b></font>"
	'uWdt = 500
	'call infoUnisport(uWdt, uTxt)
	
	if cint(hidesumaktlinier) = 1 then
	hidesumaktlinierCHK = "CHECKED"
	else
	hidesumaktlinierCHK = ""
	end if
	
    if cint(hidefasesum) = 1 then
	hidefasesumCHK = "CHECKED"
	else
	hidefasesumCHK = ""
	end if

	%>
	<span> <br /><input id="FM_hidesumaktlinier" name="FM_hidesumaktlinier" value="1" type="checkbox" <%=hidesumaktlinierCHK %> /> <%=erp_txt_325 %> <b><%=erp_txt_326 %></b> <%=erp_txt_327 %>
	 <br /><input id="FM_hidefasesum" name="FM_hidefasesum" value="1" type="checkbox" <%=hidefasesumCHK %> /> <%=erp_txt_325 %> <b><%=erp_txt_328 %></b> <%=erp_txt_329 %>
     <br /><input id="akttilfra" name="FM_akttilfra" value="1" type="checkbox"/> <%=erp_txt_330 %> <b><%=erp_txt_331 %></b> (<%=erp_txt_332 %>)


     
	<br />
        
           <%
								 if cint(hideantenh) <> 0 then
								 hideantenhCHK = "CHECKED"
								 else
								 hideantenhCHK = ""
								 end if
								 %>
								 
                                <input id="hideantenh" name="FM_hideantenh" type="checkbox" <%=hideantenhCHK %>>&nbsp;<%=erp_txt_104 %>
            
       
        </span>
	    
        
      
        <table border="0" cellspacing="0" cellpadding="0"><tr>
            <td style="width:350px; padding:20px 10px 10px 10px;"><%=erp_txt_333 %> <input type="text" id="FM_globalfaktor" value="<%=formatnumber(globalfaktor,2) %>" style="width:30px; font-size:10px;" /> (<%=erp_txt_334 %>: <span id="gblfaktorbelob_udgor">0,00</span> DKK) &nbsp;&nbsp;</td>
            <td align="right" style="padding:20px 10px 10px 10px;"><input type="button" id="bt_beregn_globalfaktor" name="bt_beregn_globalfaktor" value="Tilføj >>" style="font-size:10px;"/>
        <input type="button" value="<< Fortryd" id ="bt_beregn_globalfaktor_fortryd" style="font-size:10px; visibility:hidden; display:none;" />
        </td></tr></table>
        <br />

    <%if cint(ign_akttype_inst) = 1 then %>
    <br /><br />&nbsp;
    <div style="padding:2px; background:lightpink;"><b><%=erp_txt_335 %> </b><br /> <%=erp_txt_336 %></div>
    <%end if %>

	<br />


	<table width=100% cellspacing=0 cellpadding=2 border=0 bgcolor="#D6DFf5" id="incidentlist">
    <%=tblWdtHeader %>

	<tr bgcolor="#FFFFFF">
	    <% if cint(usefastpris) = 0 then%>
	     <td colspan=8><a name="top" class=vmenu id="top"><b><%=erp_txt_337 %>:</b></a>
             <br><span style="color:#999999;"><%=erp_txt_338 %></span>
	     </td>
	    <td class=lille style="border:1px solid #999999; border-bottom:0px; background-color:#D6Dff5;"><%=erp_txt_339 %>:<br /> <b><%=erp_txt_340 %></b></td>
	    <td>&nbsp;</td>
	 
	    <%else %>
	       <td colspan=6><a name="top" class=vmenu id="top"><b><%=erp_txt_341 %>:</b></a>
               <br><span style="color:#999999;"><%=erp_txt_342 %> <br /> <%=erp_txt_343 %></span>
	       </td>
	    <td class=lille style="border:1px solid #999999; border-bottom:0px; background-color:#D6Dff5;"><%=erp_txt_344 %>:<br /><b><%=erp_txt_345 %></b></td>
	    <td colspan=3>&nbsp;</td>
	   
	    <%end if%>
	</tr>
	
	
	<tr bgcolor="#8cAAe6"><td class=lille><b><%=erp_txt_346 %></b><br /><%=erp_txt_346 %></td>
        <td class=lille><b><%=erp_txt_348 %></b></td>
	<td class=lille><b><%=erp_txt_504 %></b></td> 
	<td align=right class=lille><b><%=erp_txt_351 %></b></td>
    <td align=right class=lille><b><%=erp_txt_352 %></b></td>
    <td align=right class=lille><b><%=erp_txt_353 %></b></td>
	<td align=right class=lille style="padding-right:5px; background-color:<%=bgfastpr%>;"><b><%=erp_txt_354 %></b><br />(<%=erp_txt_355 %>)</td>
	<td align=right class=lille>
	<b><%=erp_txt_356 %></b> <%=erp_txt_357 %>
	<%if func = "red" then %>
	<br />/<b><%=erp_txt_358 %></b>
	<%end if %>
    <br />~ <%=erp_txt_359 %>
	</td>
	
	<td align=right class=lille style="background-color:<%=bglbntim%>;"><b><%=erp_txt_360 %></b><br />(<%=erp_txt_340 %>)
	<%if func = "red" then %>
	<br />/<b><%=erp_txt_361 %></b>
	<%end if %>
	</td>
	<td align=right class=lille><b><%=erp_txt_362 %></b><br /> <%=erp_txt_363 %></td></tr>
	<%=strAktAnchor%>
	</table>
	<br /><br /><br />
	

    &nbsp;
	
	
	
	</td>
	</tr>
	</table>
    <p id="qlinkstop"></p>
     
	
	
	
	
	
	<%

    '** Valuta_enheder Global **
    call enhval_gbl	
	
	
	
	intTimer = 0
	totalbelob = 0
	subtotalbelob = 0
	subtotaltimer = 0
	totalbelob_umoms = 0
	
	nialt = 0
	for x = 0 to x - 1 
	
	'if x = 0 then
    'sumDsp = ""
    'sumVzb = "visible"
	'else
    sumVzb = "hidden"
    sumDsp = "none"
    'end if
	
    if len(trim(thisAktFase(x))) <> 0 then
    thisAktFase(x) = thisAktFase(x)
    else
    thisAktFase(x) = ""
    end if
    
    %>
	
	
	
	
	
	
	
	
	
	<span class="sumaktdiv" id="sumakt_<%=thisaktid(x) %>" style="position:relative; display:<%=sumDsp%>; visibility:<%=sumVzb%>; padding:0px; border:0px #cccccc solid;">
	
	<table width=100% border=0 cellspacing=0 cellpadding=0>
	
	<tr><td colspan=2 valign=top style="padding:20px 5px 2px 5px; border-top:5px solid #Eff3FF;"><h4>

	<a id="aname_<%=thisaktid(x)%>" style="font-size:18px;"><%=thisaktnavn(x)%></a> 
	<%
	call akttyper(thisaktfakbar(x), 1)
    %>
	
	&nbsp;<font class=roed>(<%=akttypenavn %>)</font>
	&nbsp;&nbsp;<a href="#" class="tiltoppen">^ <%=erp_txt_364 %> ^</a> 
	<!--&nbsp;&nbsp;<a href="#" id="<%=thisaktid(x)%>" class=naeste> Næste >></a> -->
	</h4>
	</td>
	</tr>
	<tr><td valign="top" style="padding:2px 20px 5px 0px;">
	<%=erp_txt_365 %>: <input type="text" name="FM_aktfase_<%=x%>" id="FM_aktfase_<%=x%>" value="<%=replace(thisAktFase(x), "_", " ") %>" style="font-size:9px; width:200px;" />&nbsp;&nbsp;
	<%=erp_txt_366 %>: <input type="text" name="aktsort_<%=x%>" id="aktsort_<%=thisaktid(x)%>" value="<%=thisaktsort(x)%>" class="sortFase" style="font-size:9px; width:30px;"> 
        <!--<input type="checkbox" class="xsortFaseCHK" id="xaktsort_chk_<%=x%>" checked /><span style="font-size:9px; color:#999999;">Sortér alle aktiviteter i fase samlet</span>-->
    <input value="<%=lcase(trim(thisAktFase(x))) %>" id="fs_<%=x%>" type="hidden" />
	<br /><img src="../blank gif" width="1" height="5" /><br /><%=erp_txt_505 %>: <input type="text" name="aktfaktor_<%=x%>" id="aktfaktor_<%=x%>" value="<%=thisaktfaktor(x) %>" style="font-size:9px; width:30px;"> <!-- mangler -->
	<input id="beregnfaktor_<%=x%>" type="button" value="Beregn" style="font-size:9px;" onClick="opd_timer_faktor(<%=x%>)" /> (<%=erp_txt_367 %>.) 
	
	</td>
	<td align=right style="padding:0px 20px 5px 0px;">
	<!--&nbsp;&nbsp;&nbsp;<a href="#" onclick="showlommeregner()">Lommerenger</a>-->
	<input name="subm_on_top_<%=x%>" id="subm_on_top_<%=x%>" type="submit" value=" Se faktura >> " style="font-size:9px;"/>
	</td></tr>
	
	</table>
	
	
	
	<%if thisaktfunc(x) <> "red" AND func = "red" then %>
        <!--
	<table>
	<tr><td colspan=2 style="background-color:#ffff99; padding:5px;">
	<font class=roed>Denne aktivitet er ikke oprettet på faktura. # timer (opr. timeantal på aktivitet), er hentet fra timeregistrering i den valgte periode.<br />
	Timer til fakturering på medarbejder er sat til 0.</font>
	</td></tr>
	</table>-->
	<%end if%>
	
	
	
	
	<table cellspacing="0" cellpadding="1" border="0" width=100% bgcolor="silver">
	<%
	
	%>
	<input type="hidden" name="sumaktbesk_<%=x%>" id="sumaktbesk_<%=x%>" value="<%=thisaktnavn(x)%>">
	<%
		'if instr(aktiswritten, "#"&thisaktid(x)&"#") = 0 then
		n = 1
		a = 0
		newa = 0
		sa = 0
		sa_medlin = 0
	
					'if thisaktfunc(x) <> "red" then
					
					
							call projektgrupper()	
							
							
					'end if
					
					if thisaktid(x) <> 0 then%>
					<!--<tr bgcolor="#d6dff5">
						<td colspan=6><!--Deaktiverede medarbejdere vises ikke ved faktura oprettelse. pr. 29/8-2005.<img src="../ill/blank.gif" width="1" height="20" alt="" border="0"><br></td>
					</tr>-->
					
					<%
					if func = "red" then
					
					    'if thisaktfunc(x) = "red" then
					    bgthis = "#8CAAE6"
					    'else
					    'bgthis = "#cccccc"
					    'end if
					
					else
					    
					    bgthis = "#8CAAE6"
					    
					end if %>
					
					<tr bgcolor="<%=bgthis %>">
						<td>&nbsp;</td>
						<td class=lille style="padding:2px 2px 2px 5px;">Vis</td>
						<td class=lille style="width:70px;">Antal - #</td>
                        <td class=lille style="width:100px; padding:2px 2px 2px 5px;">Ventetimer <br />
                        <font class=megetlillesort><%=erp_txt_371 %> | <%=erp_txt_372 %> | <%=erp_txt_373 %></font></td>
						<td class=lille style="padding:2px 2px 2px 5px;"><%=erp_txt_374 %></td>
						<td class=lille style="padding:2px 2px 2px 5px;"><%=erp_txt_375 %></td>
						<td class=lille style="padding:2px 2px 2px 5px;"><%=erp_txt_376 %></td>
						<td class=lille style="padding:2px 2px 2px 5px;"><%=erp_txt_377 %></td>
						<td class=lille style="padding:2px 2px 2px 5px;" align=center><%=erp_txt_378 %></td>
						<td class=lille style="width:80px; padding:2px 2px 2px 5px;" align=right><%=erp_txt_379 %>&nbsp;&nbsp;&nbsp;</td>
						<td style="width:80px; padding:2px 2px 2px 5px;"><font class=megetlillesort><%=erp_txt_380 %></td>
					</tr>
					<!--<tr bgcolor="#66CC33"><td colspan=6 style="border-bottom:1px #003399 solid; border-left:1px #003399 solid; border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="2" alt="" border="0"><br></td></tr>-->
			
			
					<%end if
					
					
					
					
					'***************************************************'
					'Faktimer, Venter og Beløb på medarbejdere     *****'
					'***************************************************'

                    select case lto
                    case "bf", "intranet - local" '** ONLY med_lincensindehaver 
                        med_lincensindehaverSQL = " AND med_lincensindehaver = "& session("med_lincensindehaver")
                    case else
                        med_lincensindehaverSQL = ""
                    end select
                   

					
					faktot = 0
					medarbstrpris = ""
					
					
					
					'**** Henter alle allerede oprettede aktiviteter ***'
					if thisaktfunc(x) = "red" then
						
						strSQL10c = "SELECT fms.medrabat, fd.id, antal, "_
						&" fd.beskrivelse, aktpris, fd.showonfak AS akt_showonfak, "_
						&" fd.enhedspris, "_
						&" fd.aktid, "_
						&" fak, venter_brugt, fms.tekst, fms.enhedspris, beloeb, "_
						&" mid, fms.showonfak AS showonfak, venter AS venter_ultimo, "_
						&" fms.enhedsang, fms.valuta AS mvaluta, fms.kurs AS fmskurs "_
						&" FROM faktura_det fd "_
						& "LEFT JOIN fak_med_spec fms ON (fms.fakid = fd.fakid AND fms.aktid = fd.aktid)"_
						&" WHERE fd.fakid = "& intFakid &" AND fd.aktid = "& thisaktid(x) &" AND fms.fakid = "& intFakid &" AND fms.aktid = "& thisaktid(x) &" GROUP BY mid"
						
						
						
						    medidIsWrt = "AND (mid <> 0 "
					        medarbOptionList_0 = ""
					        medarbHiddenNameList_0 = ""
					        txt_0 = ""					
        					
					        '** Alle medarbejdere med timer på ***'
        						
					        oRec3.open strSQL10c, oConn, 3
					        while not oRec3.EOF
					        medarbOptionList_0 = "<option value="& oRec3("mid") &">"& oRec3("tekst") &"</option>"   		
					        medarbHiddenNameList_0 = "<input id=""a_"&thisaktid(x)&"_"& oRec3("mid") &""" value='"& oRec3("tekst") &"' type=""hidden"" />" 
        					txt_0 = oRec3("tekst") '(Mnavn)
        					
					        call fakturaMedarbLinjer(0)
					        nialt = nialt + 1
					        medidIsWrt = medidIsWrt & " AND mid <> "& oRec3("mid")
					        sa_medlin = sa_medlin + 1 
					        'Response.flush
					        oRec3.movenext
					        wend
					        oRec3.close
        					
        					
					        medidIsWrt = medidIsWrt & ")"
						
						
						
					else
					    
					'**** Ved red: Henter alle de aktiviteter der endnu ikke er oprettet ***'
                    '**** Ved opr: Hennter alle aktiviteter ********************************'
                    
                  
				    
				    strSQL10a = "SELECT DISTINCT(tmnr) AS medarbejderid, mnavn, m.mid, "_
				    &" fms.venter_brugt, fms.venter AS venter_ultimo FROM timer "_
				    &" LEFT JOIN medarbejdere m ON (m.mid = tmnr) "_
				    &" LEFT JOIN fak_med_spec fms ON (fms.aktid = "& thisaktid(x) &" AND fms.mid = tmnr AND fms.fakid = "& lastFakid & ") "_
				    &" WHERE taktivitetid = "& thisaktid(x) &" AND (Tdato BETWEEN '" & stdatoKriAktSpecifik(x) &"' AND '"& slutdato &"') "& med_lincensindehaverSQL &" GROUP BY tmnr"
				    
                        'stdatoKri
                   
                    medidIsWrt = "AND (mid <> 0 "
					medarbOptionList_0 = ""
					medarbHiddenNameList_0 = ""		
					txt_0 = ""			
					
					'** Alle medarbejdere med timer på ***'
					
					
					'Response.Write "<u>" & strSQL10a & "</u>"
					'Response.flush
						
					oRec3.open strSQL10a, oConn, 3
					while not oRec3.EOF
					medarbOptionList_0 = "<option value="& oRec3("medarbejderid") &">"& oRec3("mnavn") &"</option>"   		
					medarbHiddenNameList_0 = "<input id=""a_"&thisaktid(x)&"_"& oRec3("medarbejderid") &""" value='"& oRec3("mnavn") &"' type=""hidden"" />" 
					
					txt_0 = oRec3("mnavn")
					
					call fakturaMedarbLinjer(0)
					nialt = nialt + 1
					medidIsWrt = medidIsWrt & " AND mid <> "& oRec3("medarbejderid")
					sa_medlin = sa_medlin + 1
					'Response.flush
					oRec3.movenext
					wend
					oRec3.close
					
					
					medidIsWrt = medidIsWrt & ")"
                  
					
					end if
					
					
					
					'*****************************'
                    '*** VenteTime Saldo Primo ***'
                    '*****************************' 
                    '*** Er der medarbejdere med ventimer på skal de være aktive ved opret ***'
                    '*** Kun AKTIVTETER der ikke ALLEREDE ER MED PÅ FAK ***'
                    
                    if thisaktfunc(x) <> "red" then
                    'sqlFaknrKri =  " fid = "& id
                    'else
                    sqlFaknrKri =  " f.fid = "& lastFakid
                    'end if
                    
                    
					strSQL2 = "SELECT f.fid, fms.venter_brugt, fms.venter AS venter_ultimo, fms.mid AS medarbejderid, m.mid, m.mnavn FROM fakturaer f "_
					&" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid AND fms.aktid = "& thisaktid(x) &""_
					&" "& replace(medidIsWrt, "mid", "fms.mid") &")"_
					&" LEFT JOIN medarbejdere m ON (m.mid = fms.mid) "_
					&" WHERE "& sqlFaknrKri &" AND f.jobid = "& jobid &" GROUP BY fms.mid ORDER BY f.fakdato DESC, aktid" 
                    
                    'Response.Write "<b>" & strSQL2 & "</b>"
                    'Response.flush
                    
                    oRec3.open strSQL2, oConn, 3 
                    while not oRec3.EOF 
                        
                        if oRec3("venter_ultimo") <> 0 then
                        
                        medarbOptionList_0 = "<option value="& oRec3("medarbejderid") &">"& oRec3("mnavn") &"</option>"   		
					    medarbHiddenNameList_0 = "<input id=""a_"&thisaktid(x)&"_"& oRec3("medarbejderid") &""" value='"& oRec3("mnavn") &"' type=""hidden"" />" 
    					
					    txt_0 = oRec3("mnavn")
    					
					    call fakturaMedarbLinjer(0)
					    nialt = nialt + 1
					    medidIsWrt = medidIsWrt & " AND mid <> "& oRec3("medarbejderid")
					    sa_medlin = sa_medlin + 1
					    
					    end if
                        
                    oRec3.movenext
                    wend 
                    oRec3.close
                    
                    end if 'akt red
					'**** Ventimer ved opret faktura slut ***'
					
					
					
					'** Mulighed for tomme linier, med medarb, 
					'** der ikke allerede er skrevet
					'** (har timer eller VENTE timer på) ***'
					
					deakSQL = " AND mansat <> '2' AND mansat <> '3' AND mansat <> '4' " '** passive må IKKE vælges til ekstra linier, f.eks Epi med 350 medarb. bliver meget tung **'
                    
                    pgrpSQL = "(projektgruppeid = "& gp1 &" OR projektgruppeid = "& gp2 &" OR projektgruppeid = "& gp3 &" OR projektgruppeid = "& gp4 &" "_
				    &" OR projektgruppeid = "& gp5 &" OR projektgruppeid = "& gp6 &" OR projektgruppeid = "& gp7 &" OR projektgruppeid = "& gp8 &" "_
				    &" OR projektgruppeid = "& gp9 &" OR projektgruppeid = "& gp10 &") "
				    
                    'pgrpSQL = "(projektgruppeid <> 1000) "
				    
				    usedmedabId = " Tmnr = 0 AND "
					
					
					'*** Opbygger optionlist **'
					medarbOptionList_1 = ""
					medarbHiddenNameList_1 = ""
					medarbOptionList_2 = ""
					medarbHiddenNameList_2 = ""
					medarbOptionList_3 = ""
					medarbHiddenNameList_3 = ""
					txt_1 = ""
					txt_2 = ""
					txt_3 = ""
					
					cn = 1
					countLimit = 0	
				    strSQL10b = "SELECT DISTINCT(medarbejderid), mnavn, mid FROM progrupperelationer, medarbejdere "_
				    &" WHERE mid = medarbejderid "& deakSQL &" AND "_
				    &""& pgrpSQL & medidIsWrt & "" & med_lincensindehaverSQL &""_
				    &" GROUP BY medarbejderid ORDER BY mnavn" 
					'& medidIsWrt
					'Response.Write "medidIsWrt" & medidIsWrt
					'Response.Write strSQL10b
					'Response.flush
						
					oRec3.open strSQL10b, oConn, 3
					while not oRec3.EOF
					'if countLimit = 0 then
					
					'end if
					
					opLSEL_1 = ""
					opLSEL_2 = ""
					opLSEL_3 = ""
					
					
					select case cn  
					case 1
					opLSEL_1 = "SELECTED"
					txt_1 = oRec3("mnavn")
					case 2
					opLSEL_2 = "SELECTED"
					txt_2 = oRec3("mnavn")
					case 3
					opLSEL_3 = "SELECTED"
					txt_3 = oRec3("mnavn")
					end select
					
					medarbOptionList_1 = medarbOptionList_1 & "<option value='"& oRec3("medarbejderid") &"' "&opLSEL_1&">"& oRec3("mnavn") &"</option>"   		
					medarbHiddenNameList_1 = medarbHiddenNameList_1 & "<input id=""a_"&thisaktid(x)&"_"& oRec3("medarbejderid") &""" value='"& oRec3("mnavn") &"' type=""hidden"" />" 
					
					medarbOptionList_2 = medarbOptionList_2 & "<option value='"& oRec3("medarbejderid") &"' "&opLSEL_2&">"& oRec3("mnavn") &"</option>"   		
					medarbHiddenNameList_2 = medarbHiddenNameList_2 & "<input id=""a_"&thisaktid(x)&"_"& oRec3("medarbejderid") &""" value='"& oRec3("mnavn") &"' type=""hidden"" />" 
					
					
					medarbOptionList_3 = medarbOptionList_3 & "<option value='"& oRec3("medarbejderid") &"' "&opLSEL_3&">"& oRec3("mnavn") &"</option>"   		
					medarbHiddenNameList_3 = medarbHiddenNameList_3 & "<input id=""a_"&thisaktid(x)&"_"& oRec3("medarbejderid") &""" value='"& oRec3("mnavn") &"' type=""hidden"" />" 
					if countLimit < 3 then
					sa_medlin = sa_medlin + 1
					end if
					
					countLimit = countLimit + 1
					cn = cn + 1
					'Response.flush
					oRec3.movenext
					wend
					oRec3.close
					
					if countLimit > 3 then
					countLimit = 3
					else
					countLimit = countLimit
					end if
					
					'countLimit = 1

					
					'*** Skriver 3 linier med medarbejdere der ikke er timer på **'	
					if countLimit <> 0 then
					%>
					<tr bgcolor="#FFFFFF"><td colspan=11 style="padding:10px;">
					    
					    <div style="color:darkred; border:1px #cccccc solid; padding:5px 5px 5px 5px;"><b><%=erp_txt_381 %></b> <%=erp_txt_382 %><br />
                        <%=erp_txt_383 %> <b><%=erp_txt_384 %></b> <%=erp_txt_385 %></div>
					    
                        <!--
                        <br />
					    <b>Der kan tilføjes optil 3 ekstra medarbejder linier her:</b><br>Vælg mellem tilknyttede (via deres projektgrupper) <b>aktive</b> medarbejdere uden timeforbrug i den valgte periode/på valgte faktura.<br />
                        <br /><a href="aktiv.asp?menu=job&func=opdprogp&jobid=<%=jobid%>" class=vmenu target="_blank">klik her</a> for at tilføje projektgrupper og medarbejdere fra jobbet til denne aktivitet. (siden re-loader)
	                    
                        -->
					   
	                   
					
					
					</td></tr>
					
					<%
					end if%>
					<%
				    strSQL10 = "SELECT DISTINCT(medarbejderid), mnavn, mid FROM progrupperelationer, medarbejdere "_
				    &" WHERE mid = medarbejderid "& deakSQL &" AND "_
				    &""& pgrpSQL & medidIsWrt & med_lincensindehaverSQL &""_
				    &" GROUP BY medarbejderid ORDER BY mnavn LIMIT 0,"& countLimit 
					
					'Response.Write strSQL10
					Response.flush
					cn = 1	
					oRec3.open strSQL10, oConn, 3
					while not oRec3.EOF
					
					call fakturaMedarbLinjer(cn)
					nialt = nialt + 1
					
					cn = cn + 1		
					'Response.flush
					oRec3.movenext
					wend
					oRec3.close
					
					%>
					
					</table>
					<br /><br />
					
					<!-- sum akt. table -->
					<table cellspacing="0" cellpadding="0" border="0" width="100%">
					
                    <%
                    
                    
                    call subtotalakt(x, 0)
                    
                    faktot = 0
					mbelobtot = 0
			
			        Response.write strAktsubtotal
			        strAktsubtotal = ""
			
			    if cint(thisMomsfri(x)) = 1 then
    	        subtotalbelobUmoms = subtotalbelob
    	        else
    	        subtotalbelobUmoms = 0
    	        end if
    	        
    	        
    	        '***** Subtotaler på sum-aktiviteter ***'
    	        Response.Write "<input type=""hidden"" name='timer_subtotal_akt_"&x&"' id='timer_subtotal_akt_"&x&"' value='"&subtotaltimer&"'>"_
                &"<input type=""hidden"" name='belob_subtotal_akt_"&x&"' id='belob_subtotal_akt_"&x&"' value='"&subtotalbelob&"'>"_
                &"<input type=""hidden"" name='belob_subtotal_akt_umoms_"&x&"' id='belob_subtotal_akt_umoms_"&x&"' value='"&subtotalbelobUmoms&"'>"_ 
			    &"<input type=""hidden"" name='txt_subtotal_akt_"&x&"' id='txt_subtotal_akt_"&x&"' value='"& thisSamleaktTxt &"'>" 
    			
    			
    			totalbelob_umoms = totalbelob_umoms + subtotalbelobUmoms
			
			
			%>
			<tr>
			    <td bgcolor="#D6Dff5" style="height:4px;"><img src="../ill/blank.gif" border=0 width=1 height=1 /></td>
			</tr>
			
			
			
			
			
		
            <input type="hidden" name="antal_subtotal_akt_<%=x%>" id="antal_subtotal_akt_<%=x%>" value="<%=sa%>">		   
			
			<input type="hidden" name="aktId_n_<%=x%>" id="aktId_n_<%=x%>" value="<%=thisaktid(x)%>">
			<input type="hidden" name="antal_n_<%=x%>" id="antal_n_<%=x%>" value="<%=n%>">
			<input type="hidden" name="highest_aval_<%=x%>" id="highest_aval_<%=x%>" value="<%=sa%>">
			<input type="hidden" name="highest_aval_m_<%=x%>" id="highest_aval_m_<%=x%>" value="<%=sa_medlin%>">
			 
			<% 
		subtotaltimer = 0 
		subtotalbelob = 0 
		'aktiswritten = aktiswritten & ",#"&thisaktid(x)&"#"		
		'end if	
		%></table>
		
		</span><!-- sum akt expand -->
		
		<%
		
		
    'Response.flush
    
    
	next
	
	
	%>
	
	
	
	
	
	<input type="hidden" name="lastactive_x" id="lastactive_x" value="0">
	<input type="hidden" name="lastactive_a" id="lastactive_a" value="0">
	<%
	
	
	'*** antal felter ***
	antalxx = x
	antalnn = n
	
	%>
	
	
	<input type="hidden" name="nialt" id="nialt" value="<%=nialt%>">
	<input type="hidden" name="antal_x" id="antal_x" value="<%=antalxx%>">
	<input type="hidden" name="antal_n" id="antal_n" value="<%=antalnn%>">
	
	<%
	if antalxx = 0 AND func = "red" then 
	%>
	<table>
	<tr>
		<td colspan=6 style="padding:20px;"><br><br><b><%=erp_txt_386 %> <u><%=erp_txt_387 %></u> <%=erp_txt_388 %></b> <br>
		<%=erp_txt_389 %><br><br>
		<%=erp_txt_390 %></font><br><br>&nbsp;</td>
	<tr>
	<%
	end if
	%>
	</table>
	
	
	
	
	
    
     <div id=aktdiv_2 style="position:relative; visibility:hidden; display:none; top:80px; width:700px; left:5px; border:0px #8cAAe6 solid;">
        <table width=700 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('jobbesk')" class=vmenu><< <%=erp_txt_391 %></a></td><td align=right><a href="#" onclick="showdiv('matdiv')" class=vmenu><%=erp_txt_392 %> >></a></td></tr></table>
    </div>
    
	
	
	
	
	</div><!-- faktura linjer div -->
	
	
	<div id="aktsubtotal" style="position:absolute; left:730px; top:104px; width:200px; z-index:2000; border:1px yellowgreen solid; background-color:#ffffff; padding:5px;">
    <b><%=erp_txt_393 %>:</b> (<%=erp_txt_394 %>)<br />
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
	<td align="right"><%=erp_txt_395 %>:
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
	
	

	
	
	
	
    
	
	
	
	
	
	
	
	
	