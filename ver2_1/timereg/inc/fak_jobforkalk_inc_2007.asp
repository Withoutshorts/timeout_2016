
 <!-- Aktiviteter -->
	    


 <div id="aktdiv" style="position:absolute; visibility:hidden; display:none; z-index:2000; left:5px; top:105px; width:700px; border:1px yellowgreen solid; padding:10px 10px 10px 10px; background-color:#ffffff;">
    <table width=100% border=0 cellspacing=0 cellpadding=0>
	<tr>
	<td><h4>Aktiviteter</h4>
	
	
	<table width=50% border=0 cellspacing=0 cellpadding=0 bgcolor="#ffffe1">
	<tr>
	    <td colspan=2 style="padding:10px 5px 2px 20px; border:1px #CCCCCC solid; border-bottom:0px;">
	        <b>Fakturerings valuta p� faktura</b><br />
	        <%call selectAllValuta(2, jftp) %>
	   </td>
	</tr>
	<tr>
		<td colspan=2 style="padding:8px 5px 10px 20px; border:1px #999999 solid; border-top:0px;"> 
		<b>Rabat</b><br />
		
                                <%if jobid <> 0 then 
                               
		                        select case intRabat
		                        case 0
		                        rSel0 = "SELECTED"
		                        rSel10 = ""
		                        rSel15 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 10
		                        rSel0 = ""
		                        rSel10 = "SELECTED"
		                        rSel15 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 15
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel15 = "SELECTED"
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 25
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel15 = ""
		                        rSel25 = "SELECTED"
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
	                            case 40
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel15 = ""
		                        rSel25 = ""
		                        rSel40 = "SELECTED"
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 50
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel15 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = "SELECTED"
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 60
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel15 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = "SELECTED"
		                        rSel75 = ""
		                        case 75
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel15 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = "SELECTED"
		                        end select
		                        %>
                                    <select id="FM_rabat_all" name="FM_rabat_all" style="width:50px; font-family:verdana; font-size:9px;" onChange="opdvalAlleForkalk(2,<%=jftp%>,'rbt');">
                                    <!-- opd_rabatrall(), setBeloebTotAll()-->
                                    
                                        <option value="0"  <%=rSel0%>>0%</option>
                                        <option value="0.10" <%=rSel10%>>10%</option>
                                        <option value="0.15" <%=rSel15%>>15%</option>
                                        <option value="0.25" <%=rSel25%>>25%</option>
                                        <option value="0.40" <%=rSel40%>>40%</option>
                                        <option value="0.50" <%=rSel50%>>50%</option>
                                        <option value="0.60" <%=rSel60%>>60%</option>
                                        <option value="0.75" <%=rSel75%>>75%</option>
                                        </select> 
                                   
                                   &nbsp;<input id="Button1" type="button" value="Opdater (5-8 sek.)" style="font-size:9px;" onClick="opdvalAlleForkalk(2,<%=jftp%>,'rbt');" />
	        
                                 
								<%end if %>
								
	
	
   
    
    <%select case intEnhedsang
    case -1
    chke_1 = "CHECKED"
	chke3 = ""
	chke2 = ""
	chke1 = ""
	case 1
	chke_1 = ""
	chke3 = ""
	chke2 = "CHECKED"
	chke1 = ""
	case 2
	chke3 = "CHECKED"
	chke2 = ""
	chke1 = ""
	chke_1 = ""
	case 3
	chke4 = "CHECKED"
	chke3 = ""
	chke2 = ""
	chke1 = ""
	chke_1 = ""
	case else
	chke_1 = ""
	chke3 = ""
	chke2 = ""
	chke1 = "CHECKED"
	end select%>
          
		<br /><br /><b>Enheder angives som: </b><br />
		
		<input type="radio" name="FM_enheds_ang" id="FM_enheds_ang0" value="0" <%=chke1%> onclick="opd_akt_endhed_forkalk('Pr. time','0')"> Timer&nbsp;&nbsp; 
		<input type="radio" name="FM_enheds_ang" id="FM_enheds_ang1" value="1" <%=chke2%> onclick="opd_akt_endhed_forkalk('Pr. stk.','1')"> Stk.&nbsp;&nbsp;
		<input type="radio" name="FM_enheds_ang" id="FM_enheds_ang2" value="2" <%=chke3%> onclick="opd_akt_endhed_forkalk('Pr. enhed','2')"> Enheder&nbsp;&nbsp;
		<input type="radio" name="FM_enheds_ang" id="FM_enheds_ang3" value="3" <%=chke4%> onclick="opd_akt_endhed_forkalk('Pr. km.','3')"> Km.&nbsp;&nbsp;
		<input type="radio" name="FM_enheds_ang" id="FM_enheds_ang_1" value="-1" <%=chke_1%> onclick="opd_akt_endhed_forkalk('Ingen','-1')"> Ingen (Skjul)<br />
		
	
       
	
	</td>
	</tr>
	</table>
<%


	
	'***********************************************************************************************'
	'*** Henter Medarbejdere og Aktiviteter ********************************************************'
	'***********************************************************************************************'
	    
	Redim editor(415)    '115    
	Redim content(415)    '115
	        
	        isaktwritten = " a.id <> 0 "
	        x = 0
	        
	        
	        if func = "red" then
        		
		        isTPused = "#"
        		
		        '*******************************************************************************'
		        '******************* Henter oprettede aktiviteter ******************************'
		        '*******************************************************************************'
		        strSQL = "SELECT fd.id, fd.antal, fd.aktid, fd.beskrivelse, fd.aktpris, "_
		        &" fd.enhedsang, fd.enhedspris, fd.rabat, "_
		        &" a.fakturerbar, "_
		        &" a.faktor, a.navn AS aktnavn, fak_sortorder, enhedsang, a.fase FROM faktura_det fd "_
		        &" LEFT JOIN aktiviteter a ON (a.id = fd.aktid) "_
		        &" WHERE fd.fakid = "& intFakid &" GROUP BY fd.aktid ORDER BY a.fase" 
		        
		        'Response.Write strSQL
		        'Response.flush
		        
		        oRec.open strSQL, oConn, 3
		        While not oRec.EOF 
        		
			      
			        thisaktnavn(x) = oRec("aktnavn") 
			        thisaktbesk(x) = oRec("beskrivelse")  
			        thisaktid(x) = oRec("aktid")
			        
			        thisaktfunc(x) = "red"
			        
			        thisaktfakbar(x) = oRec("fakturerbar")
			        thisaktfaktor(x) = oRec("faktor")
			        
			       
				    thisaktForkalk(x) = oRec("antal") 
					
					
					thisAktBeloeb(x) = oRec("aktpris")
        			
        			thisAktEnhAng(x) = oRec("enhedsang")
					thisAktEnhPris(x) = oRec("enhedspris")
					
					thisaktsort(x) = oRec("fak_sortorder")
					
					thisaktrabat(x) = oRec("rabat")
					
					thisaktCHK(x) = 1
        		
        			
        		
        			    '** antal registerede timer pr. akt. i periode **'
						strSQL2 = "SELECT sum(timer.timer) AS timerthis FROM timer WHERE taktivitetid =" & oRec("aktid") &" AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"
						oRec2.open strSQL2, oConn, 3
						
						if not oRec2.EOF then
							timerThisD = oRec2("timerthis")
						end if
						oRec2.close
						
							
							
							if len(timerThisD) <> 0 then
							timerThisD = timerThisD 
							else
							timerThisD = 0
							end if
        		        
        		        
        		        thisAktTimer(x) = timerThisD
        		        
        		        
        		        isaktwritten = isaktwritten & " AND a.id <> "& thisaktid(x) &"" 
        		        'strAktAnchor = strAktAnchor & "<a href='#"& thisaktid(x) &"' class='rmenu'>"& thisaktnavn(x) &"</a> | "
				        
				        
				        thisAktFase(x) = oRec("fase")
					
					    if lcase(lastFase) <> lcase(thisAktFase(x)) AND len(trim(thisAktFase(x))) <> 0 then
				        if x <> 0 then
				        strAktAnchor = strAktAnchor & "<br>"
				        end if
    				    
				        strAktAnchor = strAktAnchor & "<br>fase, <b>"& thisAktFase(x) &"</b><br>"
				        end if
				        
				        
				        strAktAnchor = strAktAnchor & "<a href='#"& thisaktid(x) &"' class='rmenu'>"& thisaktnavn(x) & " : "& formatnumber(thisAktTimer(x), 2) &"</a> | "
				
        		
        		
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
			
			'*** Henter fakbare, Ikke fakbare & Salg **'
			
			'**** Sidste faktura tidspunkt ****
			if len(strfaktidspkt) > 0 then
			LastFakTid = datepart("h", strfaktidspkt)&":"& datepart("n", strfaktidspkt)&":"& datepart("s", strfaktidspkt)
			else
			LastFakTid =  "00:00:02"
			end if
			
			
			'*** Henter aktiviter og finder registrede timer siden sidste faktura dato. (periodeafgr�nsning) *** 
			strSQL = "SELECT job.fastpris, a.aktbudget, "_
			&" aktstatus, a.job, a.id AS aid, a.navn AS anavn, "_
			&" a.fakturerbar, a.faktor, a.budgettimer AS akttimer, a.beskrivelse AS abesk, a.antalstk, "_
			&" aty_on_invoice_chk, aty_enh, a.fase"_
			&" FROM aktiviteter a LEFT JOIN job ON (job.id = a.job) "_
			&" LEFT JOIN akt_typer aty ON (aty_id = a.fakturerbar) "_
			&" WHERE (a.job = "& jobid &") AND ("& replace(aty_sql_onfak, "aktiviteter", "a") &") AND ("& isaktwritten &") "_
			&" ORDER BY a.fase, a.sortorder, a.id"	
			
			'AND aktiviteter.aktstatus = 1, 
			
			'Response.write strSQL
			'Response.flush
			
			oRec.open strSQL, oConn, 3
			
			
			lastAktId = 0
			timerThisT = 0
			timerThisD = 0
			'strKom = "Netto Kontant 8 dage." Hentes p� kunde
			totaltimer = 0
			lastaktid = 0
			
			
			while not oRec.EOF
				    
				    
				   if cint(lastaktid) <> cint(oRec("aid")) then
						
					
						thisaktid(x) = oRec("aid")
					    usefastpris(x) = oRec("fastpris")
						
					
							
				    thisaktnavn(x) = oRec("anavn") 
					
					
					if len(trim(oRec("fase"))) <> 0 then
					thisaktbesk(x) = replace(oRec("fase"), "_", " ") & ", "
					else
					thisaktbesk(x) = ""
					end if
					
					thisaktbesk(x) = thisaktbesk(x) & " " & oRec("anavn")  & "<br>"& oRec("abesk") 
					
	                
					thisaktForkalk(x) = oRec("antalstk") 'oRec("akttimer") 
					
					thisAktBeloeb(x) = oRec("aktbudget")
					
					thisaktfunc(x) = "opr"
					
					thisaktfakbar(x) = oRec("fakturerbar")
					
					thisaktfaktor(x) = oRec("faktor") 
					
					thisAktEnhPris(x) = 0
					
					thisaktsort(x) = 1
					
					thisAktEnhAng(x) = oRec("aty_enh")
					thisaktCHK(x) = oRec("aty_on_invoice_chk")
					
					
					if len(trim(thisakt_rabat)) <> 0 then
					thisaktrabat(x) = thisakt_rabat/100
					else
					thisaktrabat(x) = 0
					end if
					
					    
					    if fastpris = 1 then 
					    
						'** antal registerede timer pr. akt. **'
						strSQL2 = "SELECT sum(timer.timer) AS timerthis FROM timer WHERE taktivitetid =" & oRec("aid") &" AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"
						oRec2.open strSQL2, oConn, 3
						
						if not oRec2.EOF then
							timerThisD = oRec2("timerthis")
						end if
						oRec2.close
						
						else
						
						'** antal registerede timer pr. akt. **'
						strSQL2 = "SELECT sum(timer.timer * timer.timepris) AS tpthis, "_
						&" sum(timer.timer) AS timerthis FROM timer WHERE taktivitetid = " & oRec("aid") &" AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"
						
						'Response.Write strSQL2 & "<br>"
						'Response.flush
						
						
						oRec2.open strSQL2, oConn, 3
						
						if not oRec2.EOF then
							timerThisD = oRec2("timerthis")
							thisAktEnhPris(x) = oRec2("tpthis")
						end if
						oRec2.close
						
						
						
						end if
							
		
							
							if len(timerThisD) <> 0 then
							timerThisD = timerThisD 
							else
							timerThisD = 0
							end if
					
					thisAktTimer(x) = timerThisD
					
					thisAktFase(x) = oRec("fase")
					
					if lcase(lastFase) <> lcase(thisAktFase(x)) AND len(trim(thisAktFase(x))) <> 0 then
				    if x <> 0 then
				    strAktAnchor = strAktAnchor & "<br>"
				    end if
				    
				    strAktAnchor = strAktAnchor & "<br>fase, <b>"& thisAktFase(x) &"</b><br>"
				    end if
					
					'strAktAnchor = strAktAnchor & "<a href='#"& thisaktid(x) &"' class='rmenu'>"& thisaktnavn(x) &"</a> | "
					strAktAnchor = strAktAnchor & "<a href='#"& thisaktid(x) &"' class='rmenu'>"& thisaktnavn(x) & " : "& formatnumber(thisAktTimer(x), 2) &"</a> | "
				
				lastFase = oRec("fase")		
				x = x + 1		
				end if
			
			lastaktid = oRec("aid")		
			oRec.movenext
			wend
			oRec.close
			
			
			            
                        
			
			
	'end if
	'***********************************************************************************'
	
	
	
	
	'***********************************************************************************'
	'********************* Udskriver aktiviterer ***************************************'
	'***********************************************************************************'
	
	%>
	<br /><a name="top" class=vmenu><b>Quicklinks til aktiviteter p� denne faktura:</b></a><br />
	<%=strAktAnchor%>
	<br /><br />
	<%
	
	
	intTimer = 0
	totalbelob = 0
	subtotalbelob = 0
	subtotaltimer = 0
	
	nialt = 0
	
	
	    n = 1
		a = 0
		newa = 0
	
	
	
	for x = 0 to x - 1 
	%>
	
	<br /><br />
	
	<%if len(trim(thisAktFase(x))) <> 0 then %>
	fase, <b><%=thisAktFase(x) %></b>
	<%end if %>
	
	<h4><a name="<%=thisaktid(x) %>"><%=thisaktnavn(x)%></a>
	
	<%
	
	call akttyper(thisaktfakbar(x), 1)
	%>
	
	&nbsp;&nbsp;<font class=roed><b>(<%=akttypenavn %>)</b></font>
	&nbsp;&nbsp;<a href="#top" class="vmenu">^ Top ^</a>
	</h4>
	
	Sortering p� faktura: <input type="text" name="aktsort_<%=x%>" id="aktsort_<%=x%>" value="<%=thisaktsort(x)%>" style="font-size:9px; width:30px;"><br /> 
	
	
	<div id=submitTop style="position:relative; left:600px; top:-35px; z-index:10000;">
	<input name="subm_on_top_<%=x%>" id="subm_on_top_<%=x%>" type="submit" value="Se faktura" style="font-size:9px;"/>
	</div>
	
	
	<%if thisaktfunc(x) <> "red" AND func = "red" then %>
		<div id=aktikkemedpafak  style="position:relative; left:0px; top:0px; z-index:10000; width:700px; background-color:#ffff99; border:1px red dashed; padding:5px;">
	    <font class=roed>Denne aktivitet er ikke oprettet p� faktura. # timer (opr. timeantal p� aktivitet), er hentet fra timeregistrering i den valgte periode.</font>
	    </div>
	<%end if%>
	
	
	
	
	<input type="hidden" name="sumaktbesk_<%=x%>" id="sumaktbesk_<%=x%>" value="<%=thisaktnavn(x)%>">
	<%
	
	
		'if instr(aktiswritten, "#"&thisaktid(x)&"#") = 0 then
		
	
					
					
					if thisaktid(x) <> 0 then%>
					
					
					<%
					if func = "red" then
					
					    if thisaktfunc(x) = "red" then
					    bgthis = "YellowGreen"
					    else
					    bgthis = "silver"
					    end if
					
					else
					    
					    bgthis = "YellowGreen"
					    
					end if 
					
					
		            '** Skal aktiviteten vises p� faktura (chk vis) ??? ***'
		            if func = "red" then 
		            
		                if thisaktfunc(x) <> "red" then
            		    chk = ""
            		    else
            		    chk = "CHECKED"
            		    end if
            		
		            else
            		
		                if (cint(thisaktCHK(x)) = 1) AND thisaktForkalk(x) <> 0 then
		                chk = "CHECKED"
			            else
		                chk = ""
		                end if
            		
		            end if
					
					%>
					
					
					<%end if
					
					'*** Beregning af enh pris og bel�b ***'
					
					
                        
                        
                        if thisaktfunc(x) <> "red" then
                        
                            if cint(fastpris) = 1 then 
                            '*** Fastpris **' 
                        
                                if thisaktForkalk(x) <> 0 then
                                enhPris = thisAktBeloeb(x) / thisaktForkalk(x)
                                else
                                enhPris = 0
                                end if
                            
                            
                            else
                            '*** LBN timer **'
                                
                                if thisaktForkalk(x) <> 0 then
                                enhPris = (thisAktEnhPris(x) / thisaktForkalk(x))
                                else
                                enhPris = 0 'thisAktEnhPris(x) / thisAktTimer(x)
                                end if
                                
                                if len(enhPris) <> 0 then
                                enhPris = enhPris
                                else
                                enhPris = 0
                                end if
                                
                                thisAktBeloeb(x) = (enhPris * thisaktForkalk(x))
                                
                            end if
                            
                        
                        else
                        
                         enhPris = thisAktEnhPris(x)
                        
                        end if
                        
                        
                        %>
					<table cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr bgcolor="<%=bgthis %>" style="height:20px;">
					
						<td style="width:120px;">
                            &nbsp;</td>
                            <td>
                                &nbsp;</td>
						
                        <td style="padding-left:30px;">Navn & Besk.</td>
						<td align=right>Pris ialt&nbsp;&nbsp;&nbsp;</td>
						
					</tr>
                    <tr>
                        <td>Vis</td>
                        <td><input type=checkbox name="FM_show_akt_<%=x%>_0" id="FM_show_akt_<%=x%>_0" <%=chk %> value='1' onclick="tjektimer(<%=x%>,0), beregnEnhedspris(<%=x %>)"></td>
                    
                       <td rowspan=6>
                        <div id=aktnavnogbesk style="position: relative; height:200; width:400; border:0px; padding:5px 5px 5px 20px;">
	                    <%
	                    
			
			            'dim content_&""&x
            					
			            content = trim(thisAktbesk(x)) 
            			
			            'Dim editor_&""&x
			            Set editor(x) = New CuteEditor
            					
			            editor(x).ID = "FM_aktbesk_"&x&"_0"
			            editor(x).Text = content
			            editor(x).FilesPath = "CuteEditor_Files"
			            editor(x).AutoConfigure = "Minimal"
            			
			            editor(x).Width = 380
			            editor(x).Height = 180
			            editor(x).Draw()
		                %>
	                    
	                   </div>
                        
                        <!--
                        <input id="FM_aktbesk_<=x%>_0" type="hidden" value="<=trim(thisAktbesk(x)) %>" />
                        -->
                       </td>
                        
                   
            <td valign=top align=right rowspan=6 style="padding:5px;">
                     
                   <input name="FM_beloebthis_<%=x%>_0" id="FM_beloebthis_<%=x%>_0" type="text" style="width:80px;" value="<%=SQLBlessDOT(formatnumber(thisAktBeloeb(x), 2)) %>" onKeyup="tjektimer(<%=x%>,0), beregnEnhedspris(<%=x %>)" />  
                    <input id="FM_beloebthis_opr_<%=x%>_0" name="FM_beloebthis_opr_<%=x%>_0" value="<%=SQLBlessDOT(formatnumber(thisAktBeloeb(x), 2))%>" type="hidden" /> 
                            
                            </td>
            </tr>
            <tr>
                <td>Antal stk. <br />
                # Timeforb.</td>
                <td> <input id="FM_hidden_timerthis_<%=x%>_0" name="FM_hidden_timerthis_<%=x%>_0" value="<%=SQLBlessDOT(formatnumber(thisaktForkalk(x), 2)) %>" type="hidden"/> 
                            <input id="FM_timerthis_<%=x%>_0" name="FM_timerthis_<%=x%>_0" value="<%=SQLBlessDOT(formatnumber(thisaktForkalk(x), 2)) %>" type="text" style="width:50px;" onKeyup="tjektimer(<%=x%>,0), beregnBelob(<%=x %>)" /> 
                            &nbsp;<font class=lillegray><%=SQLBlessDOT(formatnumber(thisAktTimer(x), 2)) %></font>
                   </td>
                
            </tr>
            <tr>
                <td>Enhed</td>
                <td><%
                        
                        
                    select case thisAktEnhAng(x)
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
                        
                        %>
                        
                        <select name="FM_akt_enh_<%=x%>_0" id=FM_akt_enh_<%=x%>_0">
                            <option value="0" <%=ehSel_nul%>>Pr. time</option>
                            <option value="1" <%=ehSel_et%>>Pr. stk.</option>
                            <option value="2" <%=ehSel_to%>>Pr. enhed</option>
                            <option value="3" <%=ehSel_tre%>>Pr. km.</option>
                            <option value="-1" <%=ehSel_none%>>Ingen</option>
                        </select></td>
                
            </tr>
            <tr>
                <td>Enh.pris</td>
                <td>
                         <input id="FM_enhedspris_<%=x%>_0" name="FM_enhedspris_<%=x%>_0" value="<%=SQLBlessDOT(formatnumber(enhPris, 2))%>" type="text" style="width:70px;" onKeyup="tjektimer(<%=x%>,0), beregnBelob(<%=x %>)" /> 
                         <input id="FM_enhedspris_opr_<%=x%>_0" name="FM_enhedspris_opr_<%=x%>_0" value="<%=SQLBlessDOT(formatnumber(enhPris, 2))%>" type="hidden" /> 
                       </td>
                
            </tr>
            <tr>
                <td>Valuta</td>
                <td> <div style="position:relative; left:0px; top:1px; border:0px #000000 solid; padding:1px;" id="valutadiv_<%=x%>_0"><b><%=valutaKode%></b></div>
                        <input type="hidden" name="FM_valuta_<%=x%>_0" id="FM_valuta_<%=x%>_0" value="<%=valuta%>">
                        </td>
                
            </tr>
            <tr>
                <td>Rabat:
                
                 </td>
                <td> <%
                             select case (thisaktrabat(x)*100)
                            case 0
                            akt_rSel0 = "SELECTED"
                            akt_rSel10 = ""
                            akt_rSel15 = ""
                            akt_rSel25 = ""
                            akt_rSel40 = ""
                            akt_rSel50 = ""
                            akt_rSel60 = ""
                            akt_rSel75 = ""
                            case 10
                            akt_rSel0 = ""
                            akt_rSel10 = "SELECTED"
                            akt_rSel15 = ""
                            akt_rSel25 = ""
                            akt_rSel40 = ""
                            akt_rSel50 = ""
                            akt_rSel60 = ""
                            akt_rSel75 = ""
                             case 15
                            akt_rSel0 = ""
                            akt_rSel10 = ""
                            akt_rSel15 = "SELECTED"
                            akt_rSel25 = ""
                            akt_rSel40 = ""
                            akt_rSel50 = ""
                            akt_rSel60 = ""
                            akt_rSel75 = ""
                            case 25
                            akt_rSel0 = ""
                            akt_rSel10 = ""
                            akt_rSel15 = ""
                            akt_rSel25 = "SELECTED"
                            akt_rSel40 = ""
                            akt_rSel50 = ""
                            akt_rSel60 = ""
                            akt_rSel75 = ""
                            case 40
                            akt_rSel0 = ""
                            akt_rSel10 = ""
                            akt_rSel15 = ""
                            akt_rSel25 = ""
                            akt_rSel40 = "SELECTED"
                            akt_rSel50 = ""
                            akt_rSel60 = ""
                            akt_rSel75 = ""
                            case 50
                            akt_rSel0 = ""
                            akt_rSel10 = ""
                            akt_rSel15 = ""
                            akt_rSel25 = ""
                            akt_rSel40 = ""
                            akt_rSel50 = "SELECTED"
                            akt_rSel60 = ""
                            akt_rSel75 = ""
                            case 60
                            akt_rSel0 = ""
                            akt_rSel10 = ""
                            akt_rSel15 = ""
                            akt_rSel25 = ""
                            akt_rSel40 = ""
                            akt_rSel50 = ""
                            akt_rSel60 = "SELECTED"
                            akt_rSel75 = ""
                            case 75
                            akt_rSel0 = ""
                            akt_rSel10 = ""
                            akt_rSel15 = ""
                            akt_rSel25 = ""
                            akt_rSel40 = ""
                            akt_rSel50 = ""
                            akt_rSel60 = ""
                            akt_rSel75 = "SELECTED"
                            end select
                            
                             %>
            
           
                <select id="FM_rabat_<%=x%>_0" name="FM_rabat_<%=x%>_0" style="width:50px;" onchange="beregnBelob(<%=x %>)">
                <option value="0" <%=akt_rSel0%>>0%</option>
                <option value="0.10" <%=akt_rSel10%>>10%</option>
                <option value="0.15" <%=akt_rSel15%>>15%</option>
                <option value="0.25" <%=akt_rSel25%>>25%</option>
                <option value="0.40" <%=akt_rSel40%>>40%</option>
               <option value="0.50" <%=akt_rSel50%>>50%</option>
                <option value="0.60" <%=akt_rSel60%>>60%</option>
                <option value="0.75" <%=akt_rSel75%>>75%</option></select></td>
                
            </tr>
            
            
          
			</table>
			
			<br /><br />
			
			
			<input type="hidden" name="timer_subtotal_akt_<%=x%>" id="timer_subtotal_akt_<%=x%>" value="<%=SQLBlessDOT(formatnumber(thisaktForkalk(x), 2)) %>">
            <input type="hidden" name="belob_subtotal_akt_<%=x%>" id="belob_subtotal_akt_<%=x%>" value="<%=SQLBlessDOT(formatnumber(thisAktBeloeb(x), 2)) %>"> 
		    <input type="hidden" name="antal_subtotal_akt_<%=x%>" id="antal_subtotal_akt_<%=x%>" value="0">
			
		    
			<input type="hidden" name="aktId_n_<%=x%>" id="aktId_n_<%=x%>" value="<%=thisaktid(x)%>">
			<input type="hidden" name="antal_n_<%=x%>" id="antal_n_<%=x%>" value="1">
			
			<% 
		 if chk <> "" then
		intTimer = intTimer + thisaktForkalk(x)
		totalbelob = totalbelob + thisAktBeloeb(x)
		end if
		
    'Response.flush
    n = n + 1
    
	next
	%>
	
	</td>
	</tr>
	</table>
	
	
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
	
	
	
	

    <div id=aktdiv_2 style="position:relative; visibility:hidden; display:none; top:80px; width:700px; left:5px; border:0px #8cAAe6 solid;">
        <table width=700 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('jobbesk')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('matdiv')" class=vmenu>N�ste >></a></td></tr></table>
    </div>
    
	
	</div>
	
	
	<div id="aktsubtotal" style="position:absolute; left:730px; top:104px; width:200px; z-index:2000; border:1px yellowgreen solid; background-color:#ffffff; padding:5px;">
    <b>Aktiviteter:</b><br />
	<table cellspacing=5 cellpadding=0 border=0 width=100%><tr>
	<td><!--Antal ialt:-->
	<%
	'*********************************************************
	'Subtotal
	'Total timer og bel�b for aktiviteter
	'*********************************************************
	if intTimer <> 0 then
	intTimer = intTimer 
	else
	intTimer = 0
	end if
	
	%><input type="hidden" name="FM_timer" id="FM_Timer" value="0">
	<div style="position:relative; width:45; height:20px; background-color:#ffffff; padding-right:3px;" align="right" id="divtimertot">
	<!--<b><=intTimer%></b>-->
	&nbsp;</div>
	
	</td>
	<td align="right">Subtotal:
		<!-- strBeloeb -->
		<%
		if len(totalbelob) <> 0 then
		thistotbelTimer = SQLBlessDot(formatnumber(totalbelob, 2))
		else
		thistotbelTimer = formatnumber(0, 2)
		end if
		
		
		
		
		thistotbel = thistotbelTimer%>
		
		<input type="hidden" name="FM_timer_beloeb" id="FM_timer_beloeb" value="<%=thistotbelTimer%>">
		<div style="position:relative; width:95px; height:20px; border-bottom:2px YellowGreen dashed; background-color:#ffffff; padding-right:3px;" align="right" id="divtimerbelobtot"><b><%=thistotbel &" "& valutakodeSEL %></b></div>
		
		
		</td>
		</tr>
	</table>
	</div>
	
	
	
	<%call sideinfo(300,730,200) %>
	
	Ved <b>fastpris</b> beregnes enhedspris udfra pris og antal angivet p� aktiviteten.<br /><br />
	Ved <b>lbn. timer</b> beregnes enhedspris udfra (timepris p� medarbejdere * real. timer) / antal der skal produceres angivet p� akt.
	
	</td>
	</tr>
	</table>
	</div>
	
	
		

	                   
	
	
    
	
	                    