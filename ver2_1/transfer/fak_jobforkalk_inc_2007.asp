
 <!-- Aktiviteter -->
	    


 <div id="aktdiv" style="position:absolute; visibility:hidden; display:none; z-index:2000; left:5px; top:105px; width:700px; border:1px yellowgreen solid; padding:10px 10px 10px 10px; background-color:#ffffff;">
    <table width=100% border=0 cellspacing=0 cellpadding=0>
	<tr>
	<td><h4>Aktiviteter</h4>
	
	
	<table width=50% border=0 cellspacing=0 cellpadding=0 bgcolor="#ffffe1">
	<tr>
	    <td colspan=2 style="padding:10px 5px 2px 20px; border:1px #999999 solid; border-bottom:0px;">
	        <b>Fakturerings valuta på faktura</b><br />
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
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 10
		                        rSel0 = ""
		                        rSel10 = "SELECTED"
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 25
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = "SELECTED"
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
	                            case 40
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = "SELECTED"
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 50
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = "SELECTED"
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 60
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = "SELECTED"
		                        rSel75 = ""
		                        case 75
		                        rSel0 = ""
		                        rSel10 = ""
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
		<input type="radio" name="FM_enheds_ang" id="FM_enheds_ang_1" value="-1" <%=chke_1%> onclick="opd_akt_endhed_forkalk('Ingen','-1')"> Ingen (Skjul)<br />
		
	
       
	
	</td>
	</tr>
	</table>
<%


	
	'***********************************************************************************************'
	'*** Henter Medarbejdere og Aktiviteter ********************************************************'
	'***********************************************************************************************'
	        
	        
	        isaktwritten = " a.id <> 0 "
	        x = 0
	        
	        
	        if func = "red" then
        		
		        isTPused = "#"
        		
		        '*******************************************************************************'
		        '******************* Henter oprettede aktiviteter ******************************'
		        '*******************************************************************************'
		        strSQL = "SELECT fd.id, fd.antal, fd.aktid, fd.beskrivelse, fd.aktpris, "_
		        &" fd.enhedsang, fd.enhedspris, "_
		        &" a.fakturerbar, "_
		        &" a.faktor, a.navn AS aktnavn FROM faktura_det fd "_
		        &" LEFT JOIN aktiviteter a ON (a.id = fd.aktid) "_
		        &" WHERE fd.fakid = "& intFakid &" GROUP BY fd.aktid" 
		        
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
        		
        		isaktwritten = isaktwritten & " AND a.id <> "& thisaktid(x) &"" 
        		strAktAnchor = strAktAnchor & "<a href='#"& thisaktid(x) &"' class='rmenu'>"& thisaktnavn(x) &"</a> | "
					
        		
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
			
			
			'*** Henter aktiviter og finder registrede timer siden sidste faktura dato. (periodeafgrænsning) *** 
			strSQL = "SELECT job.fastpris, a.aktbudget, "_
			&" aktstatus, a.job, a.id AS aid, a.navn AS anavn, "_
			&" a.fakturerbar, a.faktor, a.budgettimer AS akttimer, a.beskrivelse AS abesk "_
			&" FROM aktiviteter a LEFT JOIN job ON (job.id = a.job) "_
			&" WHERE (a.job = "& jobid &") AND (a.fakturerbar = 0 OR a.fakturerbar = 1 OR a.fakturerbar = 6) AND ("& isaktwritten &") "_
			&" ORDER BY a.id"	
			
			'AND aktiviteter.aktstatus = 1, 
			
			'Response.write strSQL
			'Response.flush
			
			oRec.open strSQL, oConn, 3
			
			
			lastAktId = 0
			timerThisT = 0
			timerThisD = 0
			'strKom = "Netto Kontant 8 dage." Hentes på kunde
			totaltimer = 0
			lastaktid = 0
			
			
			while not oRec.EOF
				    
				    
				   if cint(lastaktid) <> cint(oRec("aid")) then
						
					
						thisaktid(x) = oRec("aid")
					    usefastpris(x) = oRec("fastpris")
						
					
							
				    thisaktnavn(x) = oRec("anavn") 
					thisaktbesk(x) = oRec("anavn")  & "<br>"& oRec("abesk") 
					
	                
					thisaktForkalk(x) = oRec("akttimer") 
					
					thisAktBeloeb(x) = oRec("aktbudget")
					
					thisaktfunc(x) = "opr"
					
					thisaktfakbar(x) = oRec("fakturerbar")
					
					thisaktfaktor(x) = oRec("faktor") 
					thisAktEnhAng(x) = 1
					thisAktEnhPris(x) = 0
					
					
					    
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
						&" sum(timer.timer) AS timerthis FROM timer WHERE taktivitetid =" & oRec("aid") &" AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"
						
						'Response.Write strSQL2
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
					
					strAktAnchor = strAktAnchor & "<a href='#"& thisaktid(x) &"' class='rmenu'>"& thisaktnavn(x) &"</a> | "
					
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
	<br /><a name="top" class=vmenu><b>Quicklinks til aktiviteter på denne faktura:</b></a><br />
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
	
	
	<br /><h4><a name="<%=thisaktid(x) %>"><%=thisaktnavn(x)%></a>
	
	<%Select case cint(thisaktfakbar(x)) 
	case 6 %>
	&nbsp;&nbsp;<font class=roed><b>(Salg & Newbizz)</b></font>
	<%case 0 %>
	&nbsp;&nbsp;<font class=roed><b>(Ikke fakturerbar)</b></font>
	<%case 1%>
	&nbsp;&nbsp;<font class=roed><b>(Fakturerbar)</b></font>
	<%case else %>
	&nbsp;&nbsp;<font class=roed><b>..</b></font>
	<%end select %>
	&nbsp;&nbsp;<a href="#top" class="vmenu">^ Top ^</a>
	</h4>
	
	
	<div id=submitTop style="position:relative; left:550px; top:-35px; z-index:10000;">
	<input name="subm_on_top_<%=x%>" id="subm_on_top_<%=x%>" type="submit" value="Se faktura" style="font-size:9px;"/>
	</div>
	
	
	<%if thisaktfunc(x) <> "red" AND func = "red" then %>
		<div id=aktikkemedpafak  style="position:relative; left:0px; top:0px; z-index:10000; width:700px; background-color:#ffff99; border:1px red dashed; padding:5px;">
	    <font class=roed>Denne aktivitet er ikke oprettet på faktura. # timer (opr. timeantal på aktivitet), er hentet fra timeregistrering i den valgte periode.</font>
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
					
					
		            '** Skal aktiviteten vises på faktura (chk vis) ??? ***'
		            if func = "red" then 
		            
		                if thisaktfunc(x) <> "red" then
            		    chk = ""
            		    else
            		    chk = "CHECKED"
            		    end if
            		
		            else
            		
		                if (cint(thisaktfakbar(x)) = 1) then
		                chk = "CHECKED"
			            else
		                chk = ""
		                end if
            		
		            end if
					
					%>
					
					
					<%end if
					%>
					<table cellspacing="0" cellpadding="2" border="0" width="100%">
					<tr bgcolor="<%=bgthis %>">
					
						<td class=lille>Vis </td>
						<td class=lille>Antal - #</td>
						<td class=lille>Enhed</td>
                        <td class=lille>Navn & Besk.</td>
						<td class=lille>Enh.pris</td>
						<td class=lille>Valuta</td>
						<td class=lille align=center>Rabat</td>
						<td class=lille align=right>Pris ialt&nbsp;&nbsp;&nbsp;</td>
						
					</tr>
                    <tr>
                        <td valign=top>
                        <input type=checkbox name="FM_show_akt_<%=x%>_0" id="FM_show_akt_<%=x%>_0" <%=chk %> value='1'>
                        </td>
                        <td valign=top>
                            <input id="FM_hidden_timerthis_<%=x%>_0" name="FM_hidden_timerthis_<%=x%>_0" value="<%=SQLBlessDOT(formatnumber(thisaktForkalk(x), 2)) %>" type="hidden"/> 
                            <input id="FM_timerthis_<%=x%>_0" name="FM_timerthis_<%=x%>_0" value="<%=SQLBlessDOT(formatnumber(thisaktForkalk(x), 2)) %>" type="text" style="width:50px;" onKeyup="tjektimer(<%=x%>,0), beregnBelob(<%=x %>)" /> 
                            &nbsp;<font class=lillegray><%=SQLBlessDOT(formatnumber(thisAktTimer(x), 2)) %></font>
                        </td>
                        
                        <%
                        
                        
                        select case thisAktEnhAng(x)
		                            case -1
                                    ehSel_none = "SELECTED"
                                    ehSel_nul = ""
		                            ehSel_et = ""
		                            ehSel_to = ""
		                            ehLabel = "Ingen"
		                            case 0
		                            ehSel_none = ""
		                            ehSel_nul = "SELECTED"
		                            ehSel_et = ""
		                            ehSel_to = ""
		                            ehLabel = "Pr. time"
		                            case 1
		                            ehSel_none = ""
		                            ehSel_nul = ""
		                            ehSel_et = "SELECTED"
		                            ehSel_to = ""
		                            ehLabel = "Pr. stk."
		                            case 2
		                            ehSel_none = ""
		                            ehSel_nul = ""
		                            ehSel_et = ""
		                            ehSel_to = "SELECTED"
		                            ehLabel = "Pr. enhed"
		                            case else
		                            ehSel_none = ""
		                            ehSel_nul = "SELECTED"
		                            ehSel_et = ""
		                            ehSel_to = ""
		                            ehLabel = "Pr. time"
		                        end select
                        
                        %>
                        
                        <td valign=top><select name="FM_akt_enh_<%=x%>_0" id=FM_akt_enh_<%=x%>_0">
                            <option value="0" <%=ehSel_nul%>>Pr. time</option>
                            <option value="1" <%=ehSel_et%>>Pr. stk.</option>
                            <option value="2" <%=ehSel_to%>>Pr. enhed</option>
                            <option value="-1" <%=ehSel_none%>>Ingen</option>
                        </select>
                        </td>
                        <td>
                            <textarea id="FM_aktbesk_<%=x%>_0" name="FM_aktbesk_<%=x%>_0" cols="30" rows="6"><%=trim(thisAktbesk(x)) %></textarea></td>
                            
                        <td valign=top>
                        
                        
                        <%
                        
                        
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
                                enhPris = thisAktEnhPris(x) / thisaktForkalk(x)
                                else
                                enhPris = thisAktEnhPris(x) / thisAktTimer(x)
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
                         
                        
                       
                            <input id="FM_enhedspris_<%=x%>_0" name="FM_enhedspris_<%=x%>_0" value="<%=SQLBlessDOT(formatnumber(enhPris, 2))%>" type="text" style="width:70px;" onKeyup="tjektimer(<%=x%>,0), beregnBelob(<%=x %>)" /> 
                            <input id="FM_enhedspris_opr_<%=x%>_0" name="FM_enhedspris_opr_<%=x%>_0" value="<%=SQLBlessDOT(formatnumber(enhPris, 2))%>" type="hidden" /> 
                            
                           
                        </td>
                        
                      <td align=center valign=top>
		        
		       
            
            
            <div style="position:relative; left:0px; top:1px; border:0px #000000 solid; padding:1px;" id="valutadiv_<%=x%>_0"><b><%=valutaKode%></b></div>
           <input type="hidden" name="FM_valuta_<%=x%>_0" id="FM_valuta_<%=x%>_0" value="<%=valuta%>">
            
            </td>
            
            
            
            <%
             select case (thisakt_rabat*100)
            case 0
            akt_rSel0 = "SELECTED"
            akt_rSel10 = ""
            akt_rSel25 = ""
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = ""
            case 10
            akt_rSel0 = ""
            akt_rSel10 = "SELECTED"
            akt_rSel25 = ""
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = ""
            case 25
            akt_rSel0 = ""
            akt_rSel10 = ""
            akt_rSel25 = "SELECTED"
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = ""
            case 40
            akt_rSel0 = ""
            akt_rSel10 = ""
            akt_rSel25 = ""
            akt_rSel40 = "SELECTED"
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = ""
            case 50
            akt_rSel0 = ""
            akt_rSel10 = ""
            akt_rSel25 = ""
            akt_rSel40 = ""
            akt_rSel50 = "SELECTED"
            akt_rSel60 = ""
            akt_rSel75 = ""
            case 60
            akt_rSel0 = ""
            akt_rSel10 = ""
            akt_rSel25 = ""
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = "SELECTED"
            akt_rSel75 = ""
            case 75
            akt_rSel0 = ""
            akt_rSel10 = ""
            akt_rSel25 = ""
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = "SELECTED"
            end select
            
             %>
            
            <td valign=top align="right">
                <select id="FM_rabat_<%=x%>_0" name="FM_rabat_<%=x%>_0" style="width:50px;" onchange="beregnBelob(<%=x %>)">
                <option value="0" <%=akt_rSel0%>>0%</option>
                <option value="0.10" <%=akt_rSel10%>>10%</option>
                <option value="0.25" <%=akt_rSel25%>>25%</option>
                <option value="0.40" <%=akt_rSel40%>>40%</option>
               <option value="0.50" <%=akt_rSel50%>>50%</option>
                <option value="0.60" <%=akt_rSel60%>>60%</option>
                <option value="0.75" <%=akt_rSel75%>>75%</option></select>
            </td>   
            <td valign=top>
                     
                   <input name="FM_beloebthis_<%=x%>_0" id="FM_beloebthis_<%=x%>_0" type="text" style="width:80px;" value="<%=SQLBlessDOT(formatnumber(thisAktBeloeb(x), 2)) %>" onKeyup="tjektimer(<%=x%>,0), beregnEnhedspris(<%=x %>)" />  
                    <input id="FM_beloebthis_opr_<%=x%>_0" name="FM_beloebthis_opr_<%=x%>_0" value="<%=SQLBlessDOT(formatnumber(thisAktBeloeb(x), 2))%>" type="hidden" /> 
                            
                            </td>
            </tr>
			</table>
			<br /><br />
			
			
			<input type="hidden" name="timer_subtotal_akt_<%=x%>" id="timer_subtotal_akt_<%=x%>" value="<%=SQLBlessDOT(formatnumber(thisaktForkalk(x), 2)) %>">
            <input type="hidden" name="belob_subtotal_akt_<%=x%>" id="belob_subtotal_akt_<%=x%>" value="<%=SQLBlessDOT(formatnumber(thisAktBeloeb(x), 2)) %>"> 
		    <input type="hidden" name="antal_subtotal_akt_<%=x%>" id="antal_subtotal_akt_<%=x%>" value="0">
			
		    
			<input type="hidden" name="aktId_n_<%=x%>" id="aktId_n_<%=x%>" value="<%=thisaktid(x)%>">
			<input type="hidden" name="antal_n_<%=x%>" id="antal_n_<%=x%>" value="1">
			
			<% 
		 
		intTimer = intTimer + thisaktForkalk(x)
		totalbelob = totalbelob + thisAktBeloeb(x)
		
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
        <table width=700 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('jobbesk')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('matdiv')" class=vmenu>Næste >></a></td></tr></table>
    </div>
    
	
	</div>
	
	
	<div id="aktsubtotal" style="position:absolute; left:730px; top:104px; width:200px; z-index:2000; border:1px yellowgreen solid; background-color:#ffffff; padding:5px;">
    <b>Aktiviteter:</b><br />
	<table cellspacing=5 cellpadding=0 border=0 width=100%><tr>
	<td>Antal ialt:
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
	
	%><input type="hidden" name="FM_timer" id="FM_Timer" value="<%=intTimer%>">
	<div style="position:relative; width:45; height:20px; border-bottom:2px YellowGreen dashed; background-color:#ffffff; padding-right:3px;" align="right" id="divtimertot"><b><%=intTimer%></b></div>
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
	
	

	
	
	
	
    
	
	                    