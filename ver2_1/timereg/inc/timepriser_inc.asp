
<%
tTop = 0
	                tLeft = 465
	                tWdth = 813
	                tHgt = "" 
	                tId = "tpris"
	                tVzb = tpris_vzb
                    tDsp = tpris_dsp
                    tZindex = 2000

	                call tableDivAbs(tTop,tLeft,tWdth,tHgt,tId, tVzb, tDsp, tZindex)

					%>
                   <table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="100%">
					
					<tr bgcolor="#5582D2">
						<td colspan=4 class=alt valign=top style="padding:5px 5px 0px 5px;"><h3 class="hv">Timepriser (job) <br /><span style="font-size:10px;">Ved Lbn. timer bliver hver time omsat med følgende timepris for den enkelte medarbejder</span></h3></td>
					</tr>
					<tr>
						
						<td style="padding-left:5px;" colspan=4><br>
						<b>Tildel timepriser for de medarbejdere der er tilknyttet dette job.</b> (via projektgrupper)<br>
                            <input id="FM_sync_tp" name="FM_sync_tp" value="1" type="checkbox" /> Sync. aktiviteter til at følge de timepriser der er angivet her. (Ikke KM aktiviteter)

                        <%if func = "opret" then%>
						Jobbet skal være oprettet før der kan tildeles alternative timepriser!
						<%else%>

                        <%strSQLakt = "select id, navn, fase FROM aktiviteter WHERE job =" & id & " ORDER BY fase, navn" 
                        
                        %>
                        
                        <br /><br />
                        <b>Vælg job eller aktivitet:</b> <select name="FM_tp_jobaktid" style="width:350px;">
                        <option value="0">Job</option>
                        <option value="0">eller vælg aktivitet..:</option>
                        <%
                        oRec5.open strSQLakt, oConn, 3
                        while not oRec5.EOF
                       %>
                       <option value="<%=oRec5("id") %>"><%=oRec5("navn") %></option>
                       <%
                       oRec5.movenext
                       Wend
                       oRec5.close%>
						</select>

                        <br /><br />
                        Timepriser 1-5 er hentet fra medarbejdertypen.

						<br><br>
						<script language="javascript">
						    $(document).ready(function () {
						        $("#timepristable").table_checkall();
						    });</script>
								<table cellspacing=1 cellpadding=2 border=0 width=800 bgcolor=#8CAAe6 id="timepristable">
								<tr bgcolor=Gainsboro><td bgcolor=#d6dff5 class=lille>Medarbejder</td>
								<td class=lille>Generel timepris<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td class=lille>Timepris 1<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td class=lille>Timepris 2<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td class=lille>Timepris 3<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td class=lille>Timepris 4<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td class=lille>Timepris 5<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td bgcolor=GreenYellow class=lille>Valgt timepris<br /><input type="button" class="checkAll" value="kolonne" /></td></tr>
								<%
								usedmids = "0#"
								strSQL = "SELECT id, navn FROM projektgrupper WHERE id = "& gp1 &" OR id = "& gp2 &" OR id = "& gp3 &" OR id = "& gp4 &" OR id = "& gp5 &" OR id = "& gp6 &" OR id = "& gp7 &" OR id = "& gp8 &" OR id = "& gp9 &" OR id = "& gp10 &" ORDER BY navn"
								oRec.open strSQL, oConn, 0, 1
								
								'Response.write strSQL
								while not oRec.EOF
									
									strSQL3 = "SELECT medarbejderid, projektgruppeid, mid, mnavn, timepris, timepris_a1, "_
									&" timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
									&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta, "_
									&" v0.valutakode AS valkode_0, "_
									&" v1.valutakode AS valkode_1, "_
									&" v2.valutakode AS valkode_2, "_
									&" v3.valutakode AS valkode_3, "_
									&" v4.valutakode AS valkode_4, "_
									&" v5.valutakode AS valkode_5 "_
									&" FROM progrupperelationer "_
									&" LEFT JOIN medarbejdere ON (mid = progrupperelationer.medarbejderid) "_
									&" LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) "_
									&" LEFT JOIN valutaer v0 ON (v0.id = tp0_valuta) "_
									&" LEFT JOIN valutaer v1 ON (v1.id = tp1_valuta) "_
									&" LEFT JOIN valutaer v2 ON (v2.id = tp2_valuta) "_
									&" LEFT JOIN valutaer v3 ON (v3.id = tp3_valuta) "_
									&" LEFT JOIN valutaer v4 ON (v4.id = tp4_valuta) "_
									&" LEFT JOIN valutaer v5 ON (v5.id = tp5_valuta) "_
									&" WHERE projektgruppeid = "& oRec("id") &" AND mnavn <> '' AND mansat <> 2 ORDER BY mnavn"
									
									'Response.Write strSQL3
									'Response.flush
									
									oRec5.open strSQL3, oConn, 0, 1
									this6timepris = "-1"
									while not oRec5.EOF
									thissel = 0
									
										if instr(usedmids, "#"&oRec5("mid")&"#") = 0 then
										t = 0
										strSQL = "SELECT jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = "& id &" AND aktid = 0 AND medarbid =  "& oRec5("mid")
										oRec4.open strSQL, oConn, 3 
										
										if not oRec4.EOF then
										
										thissel = oRec4("timeprisalt")
										this6timepris = oRec4("6timepris")
										this6valuta = oRec4("6valuta")
										
										else
										
										thissel = 0
										this6timepris = "-1"
										this6valuta = 1
										
										end if 
										oRec4.close 
										
										
                                        call meStamdata(oRec5("mid"))
										%>
										
										<tr bgcolor=#ffffff><td class=lille><%=left(meNavn, 15)%> (<%=meNr %>)
                                        <%if len(trim(meInit)) <> 0 then %>
                                        - <b><%=meInit %></b>
                                        <%end if %>
										</td>
										<%
										v = 6
										for t = 0 to v
											
											
											
											if cint(thissel) = cint(t) then
											chk = "CHECKED" 
											else
											chk = ""
											end if
											
											
											
											select case t
											case 0
											val = oRec5("timepris")
											valutaId = oRec5("tp"&t&"_valuta")
											
											case 6 
											
											if this6timepris = "-1" OR thissel = 0 then 
											val = oRec5("timepris")
                                            valutaId = oRec5("tp0_valuta")
                                            else
											val = this6timepris
											valutaId = this6valuta
											end if
											
											case else
											val = oRec5("timepris_a"&t&"")
											valutaId = oRec5("tp"&t&"_valuta")
											end select
											
											%> 
											
											
											<%if t = 6 then%>
                                            <td class=lille style="background-color:#DCF5BD;">
											<input type="text" name="FM_6timepris_<%=oRec5("mid")%>" id="FM_6timepris_<%=oRec5("mid")%>" value="<%=formatnumber(val,2)%>" style="border:1px #86B5E4 solid; font-size:9px; width:60px;">
											<%call valutakoder(600&oRec5("mid"), valutaId) %>
											<%else%>
                                            <td class=lille>
											<input type="radio" name="FM_timepris_<%=oRec5("mid")%>" id="FM_timepris_<%=oRec5("mid")%>_<%=t%>" value="<%=t%>" <%=chk%> onclick="overfortp('<%=oRec5("mid")%>', '<%=t %>');">
											<%=formatnumber(val, 2) & " " & oRec5("valkode_"&t&"") %>
											<input type="hidden" id="FM_hd_timepris_<%=oRec5("mid")%>_<%=t%>" name="FM_hd_timepris_<%=oRec5("mid")%>_<%=t%>" value="<%=formatnumber(val,2)%>">
											<input type="hidden" id="FM_hd_valuta_<%=oRec5("mid")%>_<%=t%>" name="FM_hd_valuta_<%=oRec5("mid")%>_<%=t%>" value="<%=valutaId%>">
											
											<%end if
											
											%></td><%
										next 
										%>
										</tr>
										<input type="hidden" name="FM_use_medarb_tpris" id="FM_use_medarb_tpris" value="<%=oRec5("mid")%>">
										<%
										usedmids = usedmids &oRec5("mid")&"#"
										this6timepris = "-1"
										end if
									oRec5.movenext
									wend
									oRec5.close
									
								oRec.movenext
								wend
								oRec.close
								%>

                                	<tr>
					<td colspan=4 align="right" bgcolor="#FFFFFF" style="border-top:1px #CCCCCC solid; padding:20px 40px 10px 10px;">
					     <input id="Submit4" type="submit" value=" Opdater >> " /></td>
					</tr>
								</table>
								<br>
								<%
								
								uTxt = "<b>Ovenstående timepriser</b> bliver kun benyttet hvis jobbet <b>ikke er et fastpris job.</b> "_ 
								&"Ved fastpris benyttes den almindelige fastpris beregning der er angivet på job og aktiviteter. <br><br><b>En medarbejder kan kun have en timepris pr. aktiviet. </b><br />"_
								&"Hvis timepriser ændres, opdateres alle eksisterende time-registreringer på aktiviteter der nedarver timepris fra job.<br /> Gælder også selvom der foreligger en faktura, eller perioden er afsluttet.<br><br>"_
                                &"Hvis der ændres <b>projektgrupper</b>, skal du opdatere jobbet, før den aktuelle liste af medarbejdere vises her på timepris-siden."
								uWdt = 400
								
								call infoUnisport(uWdt, uTxt) 
								%>
								
								&nbsp;
						<%end if%>
						</td>
					</tr>
					<tr>
					<td colspan=4 align="center" bgcolor="#ffffff" height="20" style=" ">
					&nbsp;</td>
				</tr>
				
			
				</table>
				<br><br><br>&nbsp;
               </div>
