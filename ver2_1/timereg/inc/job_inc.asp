<%'GIT 20160811 - SK
public aktFaseSum
dim strStamgrp_pb
redim strStamgrp_pb(100,5), aktFaseSum(100)
function stamakt(useCookie,forvalgt,fsnavn,a)
'** fsnavn skal altid starte blank **'
'** Hvis nu folk ikke ønsker at indele i faser **'
fsnavn = ""

                    '**** Henter alle stam-aktiviteter i grupper ***'
                    '** Henter aktiviteter i gruppen til projektberegner ***'
			        strSQLakt = "SELECT a.navn, a.id, a.aktfavorit, ag.navn AS agnavn, aktbudget, "_
			        &" aktstatus, bgr, antalstk, budgettimer, fase, aty_desc, fakturerbar, fravalgt FROM "_
			        &" aktiviteter AS a "_
			        &" LEFT JOIN akt_gruppe AS ag ON (ag.id = a.aktfavorit) "_
                    &" LEFT JOIN akt_typer aty ON (aty_id = a.fakturerbar) "_
			        &" WHERE a.aktfavorit <> 0 AND a.job = 0 GROUP BY a.aktfavorit, a.id ORDER BY a.aktfavorit, a.sortorder, a.navn"
			        ak = 0
			        
			        'Response.Write strSQLakt
			        'Response.flush
			        aktFaseSumtot = 0
			        lastaktFavorit = 0
			        aktsumtot = 0
                    aktgrpialt = 0
			        oRec3.open strSQLakt, oConn, 3
			        while not oRec3.EOF
			        
			        if len(trim(oRec3("budgettimer"))) = "0" OR oRec3("budgettimer") = "0" then 
	                budgettimer = 0
	                else
	                budgettimer = oRec3("budgettimer")
	                end if
			        
			        if lastaktFavorit <> oRec3("aktfavorit") AND ak <> 0 then
			        
			        strStamgrp_pb(lastaktFavorit,a) = strStamgrp_pb(lastaktFavorit,a) &"</tr><tr><td colspan=10 style='padding:5px 5px 10px 5px; border:0px;' align=right><input class='overfortiljob_a' type=""button"" value=""Indlæs gruppe(r) på job >>"" style=""font-family:arial; font-size:9px;"" /></td></tr>"
			        
			        
			        aktFaseSumtot = 0
			        strStamgrp_pb(lastaktFavorit,a) = strStamgrp_pb(lastaktFavorit,a) & "</table></td>"
			        end if
			        
			        if cint(lastaktFavorit) <> cint(oRec3("aktfavorit")) then
			        'if ak = 0 then
                    aktgrpialt = aktgrpialt + 1 
					
					if cint(forvalgt) = cint(oRec3("aktfavorit")) then
					aDsp = ""
					aVzb = "visible"
					else
					aDsp = "none"
					aVzb = "hidden"
					end if
			        
			        
			        strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) & "<tr class=""akt_pb_"& oRec3("aktfavorit") &"_"&a&""" style='visibility:"&aVzb&"; display:"&aDsp&";'>"_						
				    &"<td colspan=6 bgcolor=""#FFFFFF"" style=""border:0px; padding:5px,"">"_
                    &"Angiv fase: "_
                    &"<input class=""fs_a"" id=""fs_"&oRec3("aktfavorit")&""" name=""FM_favorit_fase"" value="""&fsnavn&""" type=""text"" style=""font-size:9px; width:150px;"" />"_
				    &"<table cellspacing=0 cellpadding=0 border=0 width='100%'><tr bgcolor='#FFFFFF'>"_
				    &"<td><b>Tilføj</b></td>"_
				    &"<td><b>Navn</b></td>"_
				    &"<td><b>Fase</b></td>"_
				    &"<td><b>Status</b></td>"_
                    &"<td><b>Type</b></td>"_
				    &"<td align=right style='width:40px;'><b>Timer</b></td>"_
				    &"<td align=right style='width:40px; padding-right:5px;'><b>Stk.</b></td>"_
				    &"<td style='width:45px;'><b>Grundlag</b></td>"_
				    &"<td align=right style='width:45px;'><b>Pr. Stk./Time</b></td>"_
				    &"<td align=right style=""width:65px; padding-left:10px;""><b>Pris i alt "& basisValISO &"</b></td>"_
				    &"<tr>"
				    
				    'lastaktFavorit = oRec3("aktfavorit")
			        end if
			        
			        
			        select case right(ak, 1)
			        case 0,2,4,6,8
			        bgcolor = "#EFf3FF"
			        case else
			        bgcolor = "#FFFFFF"
			        end select
			        
			        bgrSEL0 = "SELECTED"
	                bgrSEL1 = ""
	                bgrSEL2 = ""
			        
			        '** budget grundlag **'
			        select case oRec3("bgr")
	                case 0 '** Intent angivet
	                bgrSEL0 = "SELECTED"
	                aktsumtot = oRec3("aktbudget")
	                case 1 '** timer
	                bgrSEL1 = "SELECTED"
	                aktsumtot = oRec3("aktbudget") * budgettimer 
	                case 2 '** Stk
	                bgrSEL2 = "SELECTED"
	                aktsumtot = oRec3("aktbudget") * oRec3("antalstk")
	                end select
			        
                    if oRec3("fravalgt") <> 1 then
                    chkFravalgt = "CHECKED"
                    else
                    chkFravalgt = ""
                    end if
                    

		            strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) &"<tr bgcolor="&bgcolor&">"_
			        &"<td><input name='FM_stakt_tilfoj_"&oRec3("aktfavorit")&"_"& oRec3("id") &"' value=""1"" type=""checkbox"" "& chkFravalgt &" /></td>"
			        
                    if oRec3("fakturerbar") <> 5 then '** Kørsel
                    strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) &"<td style='width:185px;'><input id=""Hidden1"" type=""text"" name='FM_stakt_navn_"&oRec3("aktfavorit")&"_"& oRec3("id") &"' value='"& oRec3("navn") &"' style=""width:180px; font-size:9px;"" /></td>"
                    else
			        strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) &"<td style='width:185px;'><input id=""Hidden1"" type=""hidden"" name='FM_stakt_navn_"&oRec3("aktfavorit")&"_"& oRec3("id") &"' value='"& oRec3("navn") &"'/><input id=""Hidden1"" type=""text"" name='hdFM_stakt_navn_"&a&"_"& oRec3("id") &"' value='"& oRec3("navn") &"' disabled style=""width:180px; font-size:9px;"" /></td>"
                    end if

                    strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) &"<td style='width:125px;'><input class='a_fs_"&oRec3("aktfavorit")&"' id='a_fs_"&oRec3("id")&"' type=""text"" name='FM_stakt_fase_"&oRec3("aktfavorit")&"_"& oRec3("id") &"' value='"& lcase(trim(oRec3("fase"))) &"' style=""width:120px; font-size:9px;""/></td>"
			        strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) &"<td>"
	 
	                select case oRec3("aktstatus")
	                case 1
	                stTxt = "Aktiv"
	                case 2
	                stTxt = "Passiv"
                    case 0
	                stTxt = "Lukket"
	                end select
                	
	                'strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) &"<select name='FM_stakt_status_"&a&"_"& oRec3("id") &"' style='background-color:"&selbgcol&"; font-family:arial; font-size:9px;' >"_
	                '&"<option value=""1"" "&stCHK1&">Aktiv</option>"_
	                '&"<option value=""0"" "&stCHK0&">Lukket</option>"_
	                '&"<option value=""2"" "&stCHK2&">Passiv</option>"_
	                '&"</select>"_

	                strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) &"<input id=""Hidden1"" type=""hidden"" name='FM_stakt_status_"&oRec3("aktfavorit")&"_"& oRec3("id") &"' value='"& oRec3("aktstatus") &"' />"& stTxt &"</td>"
			        strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) &"<td>"& left(oRec3("aty_desc"), 7) &".</td>"
			        
			        strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) &"<td align=right>"& formatnumber(budgettimer, 2) &"</td>"_
			        &"<td align=right style=""padding-right:5px;"">"& formatnumber(oRec3("antalstk"), 2) &"</td>"
			        
			         select case oRec3("bgr")
	                case 1
	                bgtTxt = "Timer"
	                case 2
	                bgtTxt = "Stk."
	                case 0
	                bgtTxt = "Ingen" 
	                end select
                	
	                strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) &"<td>"& bgtTxt &"</td>"_
			        &"<td align=right>"& formatnumber(oRec3("aktbudget"), 2) &"</td>"_
			        &"<td align=right class=lille>= <b>"& formatnumber(aktsumtot, 2)&"</b></td></tr>"_
			        &"</td></tr>"
			        
			         
			        lastaktFavorit = oRec3("aktfavorit")
			        aktFaseSumtot = aktFaseSumtot + aktsumtot  
			        
			        
			        ak = ak + 1
			        oRec3.movenext
			        wend
			        oRec3.close
			         
			         strStamgrp_pb(lastaktFavorit,a) = strStamgrp_pb(lastaktFavorit,a) &"<tr><td colspan=10 style='padding:5px 5px 10px 5px;' align=right><input class=""overfortiljob_a"" type=""button"" value=""Indlæs gruppe på job >>"" style=""font-family:arial; font-size:9px;"" />"
			         strStamgrp_pb(lastaktFavorit,a) = strStamgrp_pb(lastaktFavorit,a) & "<input id=""Text1"" type=""hidden"" value="& aktFaseSumtot &" /></td></tr></table></td>"
			       
			        aktFaseSum(lastaktFavorit) = aktFaseSumtot
			        
                   


if a = 1 then
%>


<!--- Stam aktivitetsgrupperne i select boks efter ulgt at til hørende aktiviteter ***' -->

<tr bgcolor="#FFFFFF">
		
		<td style="border:0px; padding-left:40px; font-size:11px;">
        <input id="aktgrpialt" value="<%=aktgrpialt%>" type="hidden" />
		<% 
			strSQL = "SELECT ag.id, ag.navn, COUNT(a.id) AS antalakt FROM akt_gruppe ag "_
			&" LEFT JOIN aktiviteter a ON (a.aktfavorit = ag.id AND job = 0) "_
			&" WHERE ag.id <> 2 AND skabelontype = 0 GROUP BY ag.id ORDER BY ag.navn"
		
	     %>
		 <b>Stamaktivitetsgruppe(r):</b><br />Kombiner gerne flere, hold [ctrl] nede<br /> <select class="selstaktgrp" id="selstaktgrp_<%=a%>" name="FM_favorit" size=12 multiple style="font-size:11px; width:360px;">
		<option value="0">(Ingen)</option>
					
			<%
			stamgrpVlgt = 0
			
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
				
			
			lastStamaktgrp = forvalgt
			
			
			
			if cint(oRec("id")) = cint(lastStamaktgrp) then
			thisStaktgpSEL = "SELECTED"
			stamgrpVlgt = oRec("id")
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
			<br /><br />&nbsp;

             

			</td><td colspan=3 style="border:0px; padding-top:10px;" valign=top>&nbsp;
            <!--Overskriv fase(r): <span id="fs_ryd_<%=a %>" style="color:#999999; font-size:9px; cursor:hand;"><u>(brug stamakt.grp. navn)</u></span> <input class="fs_a" id="stgrpfs_<%=a%>" name="FM_favorit_fase" value="<%=fsnavn%>" type="text" style="font-size:9px; width:150px;" /> 
				</td>
				    <td style="border:0px;">&nbsp;</td>
                    <td style="border:0px;">&nbsp;</td>
            -->
                    <input id="lastopen_<%=a%>" value="<%=forvalgt %>" type="hidden" />
			</tr>
			
			
<%   end if    

    aktFaseSumtot = 0

        for ag = 0 TO UBOUND(strStamgrp_pb)
                if len(trim(strStamgrp_pb(ag,a))) <> 0 then
			    Response.Write strStamgrp_pb(ag,a)
			    end if
	    next
  

end function


'dim u_navn, u_ipris, u_faktor, u_belob, u_fase
'redim u_navn(20), u_ipris(20), u_faktor(20), u_belob(20), u_fase(20)
		              
function xulev_pb(func, forvalgt, lastaktfavorit, a)
                            '*** Ulev ***'
		                    if func = "red" then
		                        
		                        uopen = 0
		                        
		                        strSQLUlev = "SELECT ju_navn, ju_ipris, ju_faktor, "_
		                        &" ju_belob, ju_stk, ju_stkpris FROM job_ulev_ju WHERE ju_jobid = "& id  & " ORDER BY ju_navn "
		                        oRec2.open strSQLUlev, oConn, 3
		                        
		                        u = 1
		                        
		                        while not oRec2.EOF 
		                     
		                            u_navn(u) = oRec2("ju_navn")
		                            
		                            if len(trim(u_navn(u))) then
		                            uopen = uopen + 1
		                            else
		                            uopen = uopen
		                            end if
		                            
		                            u_ipris(u) = oRec2("ju_ipris")
		                            u_faktor(u) = oRec2("ju_faktor")
		                            u_belob(u) = oRec2("ju_belob")
		                            u_istk(u) = oRec2("ju_stk")
                                    u_istkpris(u) = oRec2("ju_stkpris")
		                            
		                         u = u + 1
		                            
		                         oRec2.movenext
		                         wend
		                         oRec2.close
		                         
		                         if uopen = 0 then
		                         uopen = 1
		                         end if
		                    
		                  
		                    else
		                        
		                        for u = 1 to 5 
		                        
		                        if lto = "intranet - local" then
		                        select case u 
		                        case 1
		                        u_navn(u) = "Køb v. underlev."
		                        u_ipris(u) = 0
		                        u_faktor(u) = 1
		                        u_belob(u) = 0
		                        case 2
		                        u_navn(u) = "Kørsel og udlæg"
		                        u_ipris(u) = 0
		                        u_faktor(u) = 1
		                        u_belob(u) = 0
		                        case 3 
		                        u_navn(u) = "Webpanel"
		                        u_ipris(u) = 0
		                        u_faktor(u) = 1
		                        u_belob(u) = 0
		                        case else
		                        end select
		                        
		                        else
		                        
		                        u_navn(u) = ""
		                        u_ipris(u) = 0
		                        u_faktor(u) = 1
		                        u_belob(u) = 0
                                u_istk(u) = 1
                                u_istkpris(u) = 0
		                        
		                        end if
		                        
		                        next
		                        
		                        uopen = 3
		                        
		                    end if
		               
                         
                         
                         %>
                        
                       
						
						<tr bgcolor="#FFFFFF" class="akt_pb_<%=lastaktfavorit%>_<%=a%>">
						    <td style="padding-top:10px;" colspan=6><b>Salgsomkostninger</b>  (underleverandører / materialeforbrug) &nbsp;&nbsp; Antal linier: xx
                                <select id="antalulev" onchange="showulevlinier()" style="font-size:9px; font-family:arial;">
                                    
                                    <%for u = 1 to 5 
                                        
                                       if u = uopen then
                                       usel = "SELECTED"
                                       else
                                       usel = ""
                                       end if
                                    
                                    %>
                                    
                                    <option value="<%=u %>" <%=usel %>><%=u %></option>
                                    
                                    <%next %>
                                   
                                </select></td>
						</tr>  
						
                        
		                <tr bgcolor="#FFCC66" class="akt_pb_<%=lastaktfavorit%>_<%=a%>">
		                    <td>Fase / Aktivitet</td>
		                    <td>Udgift navn / txt.</td>
		                    <td>Indkøbspris</td>
		                    <td>Faktor</td>
		                    <td>Beløb</td>
		                    <td>&nbsp;</td>
		                </tr>
		                
		                <%
		                
		                for u = 1 to 5 
		                
		                select case right(u, 1)
		                case 0,2,4,6,8
		                bgulev = "#FFFFFF"
		                case else
		                bgulev = "#EFF3FF"
		                end select
		                
		                 if u > uopen then
		                 uvzb = "hidden"
		                 udsp = "none"
		                 else
		                 uvzb = "visible"
		                 udsp = ""
		                 end if
		                 
		                 
		                
		                %>
		                
		                <!-- <td><input type="text" id="ulevfase_<%=u%>" name="ulevfase_<%=u%>" value="<%=u_fase(u) %>" style="width:150px; font-size:9px; font-family:arial;"></td> -->
		                 <tr bgcolor="<%=bgulev %>" id="ulevlinie_<%=u%>" style="visibility:<%=uvzb%>; display:<%=udsp%>;">
		                    
		                    <td ><input type="text" id="ulevnavn_<%=u%>" name="ulevnavn_<%=u%>" value="<%=u_navn(u) %>" style="width:200px; font-size:9px; font-family:arial;"></td>
		                   <td><input type="text" id="ulevstk_<%=u%>" name="ulevstk_<%=u%>" value="<%=replace(formatnumber(u_istk(u), 2), ".", "") %>" style="width:60px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevstk_<%=u%>'), beregnulevbelob('<%=u%>')"></td>
							<td><input type="text" id="ulevstkpris_<%=u%>" name="ulevstkpris_<%=u%>" value="<%=replace(formatnumber(u_istkpris(u), 2), ".", "") %>" style="width:60px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevstkpris_<%=u%>'), beregnulevbelob('<%=u%>')"></td>
							 <td><input type="text" id="ulevpris_<%=u%>" name="ulevpris_<%=u%>" value="<%=replace(formatnumber(u_ipris(u), 2), ".", "") %>" style="width:60px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevpris_<%=u%>'), beregnulevbelob('<%=u%>')"></td>
							
                            <td><input type="text" id="ulevfaktor_<%=u%>" name="ulevfaktor_<%=u%>" value="<%=replace(formatnumber(u_faktor(u), 2), ".", "") %>" style="width:40px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevfaktor_<%=u%>'), beregnulevbelob('<%=u%>')"></td>
		                    <td bgcolor="#FFC0CB">=// <input type="text" id="ulevbelob_<%=u%>" name="ulevbelob_<%=u%>" value="<%=replace(formatnumber(u_belob(u), 2), ".", "") %>" style="width:60px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevfaktor_<%=u%>'), beregnulevipris('<%=u%>')"> <%=basisValISO %></td>
		                    <td>&nbsp;</td>
		                </tr>
		                
		                <%next
                        
                        
                        %>
                        <tr class="akt_pb_<%=lastaktfavorit%>_<%=a%>" bgcolor="#FFFFFF">
						    <td class=alt style="height:10px;" colspan=6>&nbsp;</td>
						</tr>  
                        <%

end function


sub top

    oimg = "ikon_job_48.png"
	oleft = 20
	otop = 55
	owdt = 100
	oskrift = "Job"
	
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)

end sub


sub kundeopl_options_2016


               
                    if cint(strKnr) = oRec("Kid") then
				    kSel = "SELECTED"
				    else
				    kSel = ""
				    end if
				
				



			
				
			%>
			<option value="<%=oRec("Kid")%>" <%=kSel%>><%=oRec("Kkundenavn")%>&nbsp;&nbsp;(<%=oRec("Kkundenr")%>)</option>
			<%
			

end sub


sub kundeopl_options


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
				anstxt = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;...............&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;kontaktansv 1: "
				else
				anstxt = ""
				end if
				
			%>
			<option value="<%=oRec("Kid")%>" <%=kSel%>><%=oRec("Kkundenavn")%>&nbsp;&nbsp;(<%=oRec("Kkundenr")%>) <%=anstxt%><%=kans1%>&nbsp;&nbsp;<%=kans2%></option>
			<%
			kans1 = ""
			kans2 = ""


end sub


sub kundeopl

'******************* Kunde, og kunderelateret oplysninger  *********************************'
	if showaspopup <> "y" then
	kwidth = 450
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
		<b>Vælg Kontakt:</b>&nbsp;&nbsp;(<a href="../to_2015/kunder.asp?func=opret&ketype=k&hidemenu=1&showdiv=onthefly&rdir=1" target="_blank" class=vmenu>Opret ny her..</a>)</td>
	</tr>
    <tr>
		<td colspan=4 height="20" style="padding:10px 10px 10px 10px;">
		<b>Søg:</b> (% wildcard / vis alle)&nbsp;&nbsp;<input id="kunde_sog_step1" name="FM_kunde_sog_step1" type="text" style="width:400px;" />&nbsp;&nbsp;
            <input id="kunde_sog_step1_but" type="button" value=" Søg >> " /></td>
	</tr>
	<%end if%>
	
	
	<tr>
		
		<td style="padding:10px 10px 2px 10px;" colspan=4>
		
		
		<b>Kunde:</b><br />
		<%if func = "opret" AND step = 1 then%>
		Vælg flere hvis det samme job skal oprettes på flere kunder.<br>
		<%end if %>
		

          
		
		    <%if func = "opret" AND step = 1 then %>
            <input name="FM_kunde" id="FM_kunde_nul" value=0 type="hidden" />
			<select name="FM_kunde" id="jq_kunde_sel" size="20" multiple="multiple" style="width:650px;">
			<%else %>
			
			            <%if func = "opret" AND step = 2 then 
			            dsab = "DISABLED"
			            %>
			            <!--<input name="FM_kunde" id="Hidden1" value="<%=strKundeId %>" type="hidden" />-->
			            <%
			            else
			            dsab = ""
			            end if%>
			<select name="FM_kunde" id="FM_kunde" <%=dsab %> size=1 style="width:<%=kwidth%>px;">
          
			
			<%end if 
			
            
                
            if func = "opret" AND step = 1 then

               
                %><option value="0" disabled>Kunder (flest job seneste 12 md.):</option><%


                    tdatodd = now
                    tdatodd = dateAdd("d", -365, tdatodd)
                    tdatodd = year(tdatodd)&"/"&month(tdatodd)&"/"& day(tdatodd)

            strSQL = "SELECT Kkundenavn, Kkundenr, Kid, kundeans1, kundeans2, count(j.id) AS antaljob FROM kunder "_
            &" LEFT JOIN job AS j ON (j.jobknr = kid AND jobstartdato >= '"& tdatodd &"') "_
            &" WHERE ketype <> 'e' GROUP BY kid ORDER BY antaljob DESC LIMIT 5"
			oRec.open strSQL, oConn, 3
                
            while not oRec.EOF

                    
				
                call kundeopl_options
				
			oRec.movenext
			wend
			oRec.close

            %><option></option>
            <option value="0" disabled>Kunder (alfabetisk):</option><%
            
            end if

			strSQL = "SELECT Kkundenavn, Kkundenr, Kid, kundeans1, kundeans2 FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
			oRec.open strSQL, oConn, 3
			kans1 = ""
			kans2 = ""
			while not oRec.EOF
				
                call kundeopl_options
				
			oRec.movenext
			wend
			oRec.close
		

				
		
			%>
		</select>



		
		</td>
		
	</tr>
	
	
	
	
	
	
	
	
	
	
	<%
	'Response.flush

end sub






sub kundeopl_aft_kpers 
    if showaspopup <> "y" then
	kwidth = 450
	else
	kwidth = 310
	end if
	%>
	<tr>
		<td colspan="4" style="padding:5px 10px 5px 10px;"><b>Kontaktperson / filial:</b> (hos kunde) <a href="#" onclick="popUp('kontaktpers.asp?id=0&kid=<%=strKundeId%>&func=opr','550','450','150','120');" class="vmenu">+ Opret ny (reload) </a><br>
		
		<%
		
		strSQLkpers = "SELECT k.navn, k.id AS kid FROM kontaktpers k WHERE kundeid = "& strKundeId &" ORDER BY k.navn"
		'Response.Write strSQLkpers
		
		%>
		
		<select name="FM_opr_kpers" id="FM_opr_kpers" style="width:<%=kwidth%>px;">
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
		</select>
        
          <br />
        <input type="checkbox" value="1" name="FM_brugaltfakadr" <%=altfakadrCHK %> />Brug kontaktperson / filial som modtager på faktura.
        </td>
	
	</tr>
	
	
	
	
	
<%
end sub


               
               
		
	
		
		
		
		
		
        function pjgcase(grp, thisprg)
		
		if cint(grp) = cint(strProjektgr1) OR cint(grp) = cint(strProjektgr2) OR cint(grp) = cint(strProjektgr3) OR cint(grp) = cint(strProjektgr4) OR cint(grp) = cint(strProjektgr5) OR cint(grp) = cint(strProjektgr6) OR cint(grp) = cint(strProjektgr7) OR cint(grp) = cint(strProjektgr8) OR cint(grp) = cint(strProjektgr9) OR cint(grp) = cint(strProjektgr10) Then
		else
			
			strSQL3 = "UPDATE aktiviteter SET projektgruppe"&thisprg&" = 1 WHERE id = " & oRec5("id")
			'Response.Write strSQL3
			'Response.flush
			
			oConn.execute(strSQL3)
		
		end if


end function





sub addUlev

                                vlgtUlevgrp = request("FM_jugrp_id")


                                'Response.Write "vlgtUlevgrp: " & vlgtUlevgrp & "<br>"
                                'Response.end
                                if vlgtUlevgrp <> 0 then
                                            

                                            strSQLUlev = "SELECT ju_id, ju_navn, ju_ipris, ju_faktor, ju_id, ju_fase, ju_stk, ju_stkpris, "_
		                                    &" ju_belob FROM job_ulev_ju WHERE ju_favorit = "& vlgtUlevgrp  & " AND ju_jobid = 0 ORDER BY ju_fase, ju_navn "
		                                    
                                            oRec5.open strSQLUlev, oConn, 3
	
	                                        while not oRec5.EOF 
	                                                
                                                        
                                                        if request("FM_ulev_linie_"& oRec5("ju_id") &"") = "1" then

                                                        ulevnavn = replace(oRec5("ju_navn"), "'", "''")
                                                        ju_fase = replace(oRec5("ju_fase"), "'", "''")

                                                        ulevpris = replace(oRec5("ju_ipris"), ",", ".") 
                                                        ulevstk = replace(oRec5("ju_stk"), ",", ".")
                                                        ulevstkpris = replace(oRec5("ju_stkpris"), ",", ".")

                                                        ulevfaktor = replace(oRec5("ju_faktor"), ",", ".") 
                                                        ulevbelob = replace(oRec5("ju_belob"), ",", ".") 

                                                        strSQLInsUlev = "INSERT INTO job_ulev_ju (ju_navn, ju_ipris, ju_faktor, ju_belob, ju_fase, ju_jobid, ju_stk, ju_stkpris) VALUES "_
							                            &" ('"& ulevnavn &"', "& ulevpris &", "_
							                            &""& ulevfaktor &", "& ulevbelob &", '"& ju_fase &"', "& varJobId &", "& ulevstk &", "& ulevstkpris &")"  
							                            
                                                        'Response.Write strSQLInsUlev
                                                        'Response.flush
                                                        oConn.execute(strSQLInsUlev)

                                                        end if


                                            oRec5.movenext
	                                        wend 
                                            oRec5.close


                                            'Response.end
                                end if

end sub




public strUlevTxt, strUlevSEL
'*** FaviritID, antalgrp **'
redim strUlevTxt(100, 3), strUlevSEL(100,3)
function ulev_tilfoj(ufvlgt, u)

                                


                                '*** Ulev linier i grp ***'
		              
		                        strSQLUlev = "SELECT ju_id, ju_navn, ju_ipris, ju_faktor, ju_id, ju_fase, ju_favorit, ju_stk, ju_stkpris, "_
		                        &" ju_belob, ju_fravalgt FROM job_ulev_ju WHERE ju_favorit <> 0 AND ju_jobid = 0 ORDER BY ju_favorit, ju_fase, ju_navn "
		                        oRec2.open strSQLUlev, oConn, 3
		                        
		                       
		                        while not oRec2.EOF 
		                     
		                            u_navn = oRec2("ju_navn")
		                            u_id = oRec2("ju_id")
		                            u_fravalgt = oRec2("ju_fravalgt")

                                    if cint(u_fravalgt) = 0 then
                                    u_fravalgtCHK = "CHECKED"
                                    else
                                    u_fravalgtCHK = ""
                                    end if

		                            u_ipris = oRec2("ju_ipris")
		                            u_faktor = oRec2("ju_faktor")
		                            u_belob = oRec2("ju_belob")
		                            u_id = oRec2("ju_id")
                                    u_fase = "" 'oRec2("ju_fase")

                                    u_istk = oRec2("ju_stk")
                                    u_istkpris = oRec2("ju_stkpris")

                                    u_favorit = oRec2("ju_favorit")

                                    strUlevTxt(u_favorit, u) =  strUlevTxt(u_favorit, u) & "<tr>"_
                                    &"<td style=""width:30px;""><input id=""Checkbox1"" type=""checkbox"" "& u_fravalgtCHK &" value='1' name='FM_ulev_linie_"& u_id &"' /></td>"_
                                    &"<td style=""width:150px;"">" & u_navn & "&nbsp;</td>"_
                                    &"<td style=""width:120px;"">" & formatnumber(u_istk) &" stk. á "& formatnumber(u_istkpris) &"</td>"_
                                    &"<td style=""width:90px;"" align=right>"& formatnumber(u_ipris, 2) &"</td>"_
                                    &"<td style=""width:60px;"" align=right>"& formatnumber(u_faktor, 2) &"</td>"_
                                    &"<td style=""width:90px;"" align=right>"& formatnumber(u_belob, 2) &"</td></tr>"
		                            
		                            strUlevSEL(u_favorit, u) = u_favorit
		                            
		                         oRec2.movenext
		                         wend
		                         oRec2.close
		                         
		                       
		                 
		               
                        
                         





end function



sub projektberegner

          if func = "red" OR step = "2" then

                    
                  


	                fTop = 0
                    fLeft = 487
                    fWdth = 700
                    
                    if func <> "red" then
                    fHgt = "2420" 
                    else
                    fHgt = ""
                    end if
                    
                    fId = "pb_div"
                    fZindex = 1000

                  

	                call tableDivAbs(fTop,fLeft,fWdth,fHgt,fId,fVzb,fDsp,fZindex) %>
                           
	                    
	                    



                       
                   <%
                   select case lto
                   case "xepi", "xintranet - local"
                    
                    if (formatnumber(jo_gnsbelob,2) = formatnumber(0,2)) AND func = "red" then
                    netno_vzb = "visible" 
                    netno_dsp = ""
                    else
                    netno_vzb = "hidden" 
                    netno_dsp = "none"
                    end if

                   case else
                   netno_vzb = "hidden" 
                   netno_dsp = "none"
                   end select %>
                   
                  
                   <div id="nettoomsnote" style="position:absolute; left:170px; top:105px; border:10px #CCCCCC solid; width:250px; visibility:<%=netno_vzb%>; display:<%=netno_vdsp%>; background-color:snow; padding:10px;">
                   <table cellspacing="0" cellpadding="0" border="0" width=100%><tr>
                   <td>Husk at indtaste bruttoomsætning, før der angives timer på medarbejdertyper. </td>
                   <td style="padding-left:5px;"><img src="../ill/pile.gif" border="0" /></td>
                   </tr></table>
                
                    </div> 

                   
	            <table cellspacing="0" cellpadding="0" border="0" bgcolor="#FFFFFF" class="pb_table">
				<tr bgcolor="#5582d2">
					<td colspan=6 class=alt style="padding:5px 5px 0px 5px; border:0px;"><h3 class="hv">Job forkalkulation og budget<br />
                    <span style="font-size:9px; color:#ffffff; line-height:12px;">Indtast direkte i nettoomsætning, eller benyt detaljeret budget på "aktiviter & faser" ell. "medarbejdertyper" nedenfor. (se. konfiguration) </span></h3><!-- Aktiviteter og Salgsomkost.  -->
                    </td>
				</tr>
                  <tr>
                <td align=right colspan=6 style="border:0px; padding:15px 15px 5px 5px;"> <input class="overfortiljob_u" type="button" value="Gem ændringer på job >>" style="font-family:arial; font-size:9px;"/></td>
                </tr>


                  <!--- Budget på medarbejertyper **** -->

                <% 
                showMtberegn = 0

                call bdgmtypon_fn()


                        'select case lto
                        'case "epi", "epi_no", "epi_sta", "intranet - local", "xwwf", "epi_ab", "epi_cati"
                        if cint(bdgmtypon_val) = 1 then
                        budgetmtyp_vzb = "visible"
                        budgetmtyp_dsp = ""
                        sync_mtCHK = "CHECKED"
                        showMtberegn = 1
                        else
                        budgetmtyp_vzb = "hidden"
                        budgetmtyp_dsp = "none"

                        sync_mtCHK = ""
                        end if

				%>

                
                 <input id="sync_budget_mt" type="hidden" value="<%=showMtberegn %>" />


                 <%
				if func = "red" then
                    
                     if cint(showMtberegn) = 1 then

                    %>     
                
                      
			  
				   <tr bgcolor="#DCF5BD">
					<td style="padding:20px 0px 3px 5px;" colspan="6"><h4><a href="#" id="a_budgetmtyp">[+]</a> Budget på medarbejdertyper

                    <%if level = 1 then %>
                       <a href="medarbtyper.asp" class=rmenu target="_blank">Se kost- og time -priser på medarbejdertyper her >></a>
                       <%end if %>

                    <!--
                    <br />
                    <span style="font-size:9px; color:#000000; line-height:12px;">Budget på medarbejdertyper indgår ikke i beregning af DB på job (nederst).</span>
                    -->
                    </h4>
                
                </td>
                </tr>

                 <tr class="tr_budgetmtyp" bgcolor="#FFFFFF">
                 

                 <td style="border:0px; padding-top:20px;" colspan="6">
                    
                     <table width=100% cellspacing=0 cellpadding=0 border=0>
                     <tr>
                         <td style="padding:12px 2px 2px 2px; font-size:11px; width:460px;"><b>Bruttoomsætning:</b><br /><span style="font-size:9px; color:#999999;">Nettooms. + Salgsomkostninger</span></td>
                         <td align=right style="padding:12px 20px 2px 2px; font-size:11px;">= <input type="text" id="FM_budget" name="FM_budget" value="<%=replace(formatnumber(jo_bruttooms, 2), ".", "")%>" style="width:75px; font-size:11px; font-family:arial; border:1px red solid;" onkeyup="tjektimer('FM_budget')">&nbsp;&nbsp;&nbsp;&nbsp; <%=basisValISO %></td>
			
                     </tr>
                     </table>

                     </td>
                </tr>

               

                <tr class="tr_budgetmtyp" id="tr_budgetmtyp" style="visibility:<%=budgetmtyp_vzb%>; display:<%=budgetmtyp_dsp%>;">
                <td style="border:0px; padding:20px 20px 20px 0px;" colspan=6>

                    

                <table cellpadding=0 cellspacing=0 border=0 width=100%>
                <tr>
                    <td><b>Medarbejdertype</b> (gruppe)</td>
                    <td><b>Timer</b></td>
                    <td><b>Timepris</b></td>
                    <td>Beløb<br /> før faktor</td>
                    <td><b>Faktor</b></td>
                    <td><b>Beløb</b></td>
                </tr>

                <%
                
                call meStamdata(session("mid")) 

                lastGrpId = 0
                
                strSQLmtyp = "SELECT typeid, timer, mt_tb.timepris, faktor, belob, mt.type, belob_ff, mt_tb.kostpris, "_
                &" mtg.navn AS gruppenavn, mtg.id AS mgruppe, mtg.opencalc "_
                &" FROM medarbejdertyper_timebudget AS mt_tb "_
                &" LEFT JOIN medarbejdertyper mt ON (mt.id = typeid) "_
                &" LEFT JOIN medarbtyper_grp AS mtg ON (mtg.id = mt.mgruppe)"_
                &" WHERE mt_tb.jobid = "& id &" AND (mt.sostergp <> -1) ORDER BY gruppenavn, mtsortorder"
                

                'Response.write strSQLmtyp
                'Response.flush
                tpno = 0
                oRec3.open strSQLmtyp, oConn, 3
                While not oRec3.EOF 

                if isNull(oRec3("timer")) <> true then
                timerThis = oRec3("timer")
                else
                timerThis = 0
                end if

                if isNull(oRec3("timepris")) <> true then
                timepris = oRec3("timepris")
                else
                timepris = 0
                end if

                if isNull(oRec3("faktor")) <> true then
                faktor = oRec3("faktor")
                else
                faktor = 0
                end if

                if isNull(oRec3("belob")) <> true then
                belob = oRec3("belob")
                else
                belob = 0
                end if


                if isNull(oRec3("belob_ff")) <> true then
                belob_for_fakor = oRec3("belob_ff")
                else
                belob_for_fakor = 0
                end if

                if isNull(oRec3("kostpris")) <> true then
                kostpris = oRec3("kostpris")
                else
                kostpris = 0
                end if

                kostBelob = (timerThis*kostpris)
                
                if isNull(oRec3("gruppenavn")) <> true then
                gruppenavnTxt = " ("& oRec3("gruppenavn") & ")"
                else
                gruppenavnTxt = ""
                end if 

                grpId = oRec3("mgruppe")
                opencalc = oRec3("opencalc")


                if cint(lastGrpId) <> cint(grpId) then
                 %>
               <input type="hidden" id="FM_mtype_startno_<%=grpId%>" name="FM_mtype_startno" value="<%=tpno%>"/>
               <%end if


               if tpno > 0 then
               call hovedgrpsum(lastGruppenavnTxt, gruppenavnTxt, timerTotThisGrp, timeprisTotThisGrp, faktorTotThisGrp, belobTotThisGrp, basisValISO, 0, lastGrpId, lastOpencalc, tpno, kostBelobTotThisGrp)
             
               end if
                
                %>
                <tr><td style="width:300px; font-size:9px;"><%=oRec3("type") %> <%=gruppenavnTxt %> <span style="font-size:9px; color:#999999;">(tp: <%=formatnumber(timepris) %> / kost: <%=formatnumber(kostpris) %>)</span>

                    <input type="hidden" id="FM_mtype_kp_<%=grpId%>_<%=oRec3("typeid") %>" name="FM_mtype_kostpris_<%=oRec3("typeid") %>" value="<%=formatnumber(kostpris) %>" />
                <input type="hidden" id="FM_mtype_ids_<%=tpno%>" name="FM_mtype_ids" value="<%=oRec3("typeid")%>"/>
                <input type="hidden" id="FM_mtype_id_<%=oRec3("typeid") %>" name="FM_mtype_id" value="<%=oRec3("typeid") %>" />
                <input type="hidden" id="FM_mtypegrp_id_<%=grpId%>_<%=oRec3("typeid") %>" name="FM_mtypegrp_id" value="<%=grpId%>" /></td>
                    <td><input type="text" class="mt_timepris" id="FM_mtype_ti_<%=grpId%>_<%=oRec3("typeid") %>" name="FM_mtype_timer_<%=oRec3("typeid") %>" style="width:50px; font-size:9px;" value="<%=formatnumber(timerThis, 2) %>" /></td>
                     <td style="width:150px; font-size:9px; white-space:nowrap;"><input type="text" id="FM_mtype_tp_<%=grpId%>_<%=oRec3("typeid") %>" class="mt_timepris" name="FM_mtype_timepris_<%=oRec3("typeid") %>" value="<%=formatnumber(timepris, 2) %>" style="width:75px; font-size:9px;" /> <%=basisValISO %></td>
                      <td style="width:50px; font-size:9px; white-space:nowrap;"> (<input type="text" class="mt_timer" id="FM_mtype_ff_<%=grpId%>_<%=oRec3("typeid") %>" value="<%=formatnumber(belob_for_fakor, 2) %>" name="FM_mtype_belob_ff_<%=oRec3("typeid") %>" style="width:40px; color:#999999; font-size:9px; border:0px;" maxlength="0" />)</td>
                      <td style="width:50px; font-size:9px; white-space:nowrap;"> x <input type="text" class="mt_faktor" id="FM_mtype_fa_<%=grpId%>_<%=oRec3("typeid") %>" value="<%=formatnumber(faktor, 2) %>" name="FM_mtype_faktor_<%=oRec3("typeid") %>" style="width:40px; font-size:9px;" maxlength="0" /></td>
                      <td style="width:120px; font-size:9px; white-space:nowrap;">= <input type="text" class="mt_timer" id="FM_mtype_be_<%=grpId%>_<%=oRec3("typeid") %>" value="<%=formatnumber(belob, 2) %>" name="FM_mtype_belob_<%=oRec3("typeid") %>" style="width:75px; font-size:9px;" maxlength="0" /> <%=basisValISO %>

                          <input type="hidden" id="FM_mtype_kostbe_<%=grpId%>_<%=oRec3("typeid") %>" value="<%=formatnumber(kostBelob, 2) %>" name="" />

                      </td>
                </tr>

                <%
                lastOpencalc = opencalc
                 lastGrpId = grpId  
                lastGruppenavnTxt = gruppenavnTxt
                tpno = tpno + 1
                typeidWrt = typeidWrt & " AND mt.id <> " & oRec3("typeid") 
                beltot_ff = beltot_ff + belob_for_fakor
                beltot = beltot + belob
                belobTotThisGrp = belobTotThisGrp + belob
                timerTotThisGrp = timerTotThisGrp + timerThis

                timeprisTotThisGrp = timeprisTotThisGrp + (timerThis*timepris) 'vægtet
                faktorTotThisGrp = faktorTotThisGrp + (faktor*timerThis) 'vægtet
                kostBelobTotThisGrp = kostBelobTotThisGrp + kostBelob


                oRec3.movenext
                wend
                oRec3.close 
                
                if tpno <> 0 then
                 call hovedgrpsum(lastGruppenavnTxt, gruppenavnTxt, timerTotThisGrp, timeprisTotThisGrp, faktorTotThisGrp, belobTotThisGrp, basisValISO, 1, lastGrpId, lastOpencalc, tpno, kostBelobTotThisGrp)
                 belobTotThisGrp = 0
                 timerTotThisGrp = 0
                 timeprisTotThisGrp = 0
                 faktorTotThisGrp = 0
                 kostBelobTotThisGrp = 0
                end if
              
                '*** Henter typer uden budget **' 
                strSQLmtyp = "SELECT mt.type, mt.id AS id, mt.timepris, mt.kostpris, mtg.navn AS gruppenavn, "_
                &" mtg.id AS mgruppe, opencalc FROM medarbejdertyper AS mt "_
                &" LEFT JOIN medarbtyper_grp AS mtg ON (mtg.id = mt.mgruppe AND mt.sostergp = 0)"_
                &" WHERE mt.id <> 0 "& typeidWrt &" AND (mt.sostergp = 0) ORDER BY gruppenavn, mt.mtsortorder"
                
                lastGrpId = 0
                tpii = 0
                oRec3.open strSQLmtyp, oConn, 3
                While not oRec3.EOF 

                if isNull(oRec3("timepris")) <> true then
                timepris = oRec3("timepris")
                else
                timepris = 0
                end if
                
                if isNull(oRec3("kostpris")) <> true then
                kostpris = oRec3("kostpris")
                else
                kostpris = 0
                end if

                if isNull(oRec3("gruppenavn")) <> true then
                gruppenavnTxt = " ("& oRec3("gruppenavn") & ")"
                else
                gruppenavnTxt = ""
                end if 

               if isNull(oRec3("mgruppe")) <> true then
               grpId = oRec3("mgruppe")
               else
               grpId = 0
               end if

               opencalc = oRec3("opencalc")



               if cint(grpId) <> cint(lastGrpId) then
                %>
                <input type="hidden" id="FM_mtype_startno_<%=grpId%>" name="FM_mtype_startno" value="<%=tpii%>"/>
                <% end if

                if tpii > 0 then
                 call hovedgrpsum(lastGruppenavnTxt, gruppenavnTxt,  0, 0, 0, 1, basisValISO, 0, lastGrpId, lastOpencalc, tpii, 0)
                
               end if
                


                if tpii = 0 then
                timerSt = 0
                    else
                   timerSt = 0%>
                

                <%end if %>
                <tr><td style="width:300px; font-size:9px;"><%=oRec3("type") %> <%=gruppenavnTxt %> <span style="font-size:9px; color:#999999;">(tp: <%=formatnumber(timepris) %> / kost: <%=formatnumber(kostpris) %>)</span>
                <input type="hidden" id="FM_mtype_kp_<%=grpId%>_<%=oRec3("id") %>" name="FM_mtype_kostpris_<%=oRec3("id") %>" value="<%=formatnumber(kostpris) %>" />
                <input type="hidden" id="FM_mtype_ids_<%=tpno%>" name="FM_mtype_ids" value="<%=oRec3("id")%>"/>
                <input type="hidden" id="FM_mtype_id_<%=oRec3("id")%>" name="FM_mtype_id" value="<%=oRec3("id")%>" />
                <input type="hidden" id="FM_mtypegrp_id_<%=grpId%>_<%=oRec3("id")%>" name="FM_mtypegrp_id" value="<%=grpId%>" /></td>
                    <td><input type="text" id="FM_mtype_ti_<%=grpId%>_<%=oRec3("id") %>" name="FM_mtype_timer_<%=oRec3("id") %>" value="<%=timerSt%>" style="width:50px; font-size:9px;" class="mt_timepris" /></td>
                     <td style="width:150px; font-size:9px; white-space:nowrap;"><input type="text" id="FM_mtype_tp_<%=grpId%>_<%=oRec3("id") %>" class="mt_timepris" name="FM_mtype_timepris_<%=oRec3("id") %>" value="<%=formatnumber(timepris, 2) %>" style="width:75px; font-size:9px;" /> <%=basisValISO %></td>
                       <td style="width:50px; font-size:9px; white-space:nowrap;"> (<input type="text" value="0" id="FM_mtype_ff_<%=grpId%>_<%=oRec3("id") %>" class="mt_timer" name="FM_mtype_belob_ff_<%=oRec3("id") %>" style="width:40px; font-size:9px; color:#999999; border:0px;" maxlength="0" />)</td>
                     
                      <td style="width:50px; font-size:9px; white-space:nowrap;"> x <input type="text" value="1" id="FM_mtype_fa_<%=grpId%>_<%=oRec3("id") %>" class="mt_faktor" name="FM_mtype_faktor_<%=oRec3("id") %>" style="width:40px; font-size:9px;" maxlength="0" /></td>
                      <td style="width:120px; font-size:9px; white-space:nowrap;">= <input type="text" value="0" id="FM_mtype_be_<%=grpId%>_<%=oRec3("id") %>" class="mt_timer" name="FM_mtype_belob_<%=oRec3("id") %>" style="width:75px; font-size:9px;" maxlength="0" /> <%=basisValISO %>

                            <input type="hidden" id="FM_mtype_kostbe_<%=grpId%>_<%=oRec3("id") %>" value="0" name="" />

                      </td>
                </tr>

                <%

                lastOpencalc = opencalc 
                lastGrpId = grpId 
                lastGruppenavnTxt = gruppenavnTxt
                tpno = tpno + 1
                beltot_ff = beltot_ff + 0 
                beltot = beltot + 0
                tpii = tpii + 1
                oRec3.movenext
                wend
                oRec3.close 
                
                if tpii <> 0 then
                  call hovedgrpsum(lastGruppenavnTxt, gruppenavnTxt, 0, 0, 0, 1, basisValISO, 1, lastGrpId, lastOpencalc, tpii, 0)
                end if
                %>
                <!--
                <tr>
                    <td style="border:0px;">&nbsp;</td>
                    <td style="height:22px; border:0px;"><input type="button" value="opdater >> " id="mt_calc" style="font-size:8px;" /></td>
                    <td style="border:0px;">&nbsp;</td>
                    <td style="font-size:9px; border:0px;"> 
                    (<input type="text" value="<%=formatnumber(beltot_ff,2) %>" id="FM_mtype_belob_ff_tot" name="FM_mtype_belob_ff_tot" style="width:40px; border:0px; color:#999999; font-size:9px;" />)</td>
                    <td style="border:0px;">&nbsp;</td>
                    <td style="font-size:9px; border:0px;">= <input type="text" value="<%=formatnumber(beltot,2) %>" id="FM_mtype_belob_tot" name="FM_mtype_belob_tot" style="width:75px; border:0px; color:#999999; font-size:9px;" /> <%=basisValISO %></td>
                </tr>
                -->

                 <input type="hidden" value="<%=formatnumber(beltot_ff,2) %>" id="FM_mtype_belob_ff_tot" name="FM_mtype_belob_ff_tot"/>
                 <input type="hidden" value="<%=formatnumber(beltot,2) %>" id="FM_mtype_belob_tot" name="FM_mtype_belob_tot" />

                </table>
               
              

                <input type="hidden" value="<%=tpno-1 %>" id="FM_mtype_natal" name="FM_mtype_natal" />
                <input type="hidden" value="<%=meType %>" id="FM_mtype_user" name="FM_mtype_user" />

            
                </td></tr>
                

                <!--
                  <tr class="tr_budgetmtyp" style="visibility:<%=budgetmtyp_vzb%>; display:<%=budgetmtyp_dsp%>;"><td colspan=6 align="right" style="border:0px; padding:10px 10px 30px 5px;">
                    <input class="overfortiljob_u" type="button" value="Gem ændringer på job >>" style="font-family:arial; font-size:9px;"/></td>
                    </tr>
                    --> 
                    <%else %>
                     <tr>
                     <td style="border:0px;" colspan="6">
                    
                     <table width=100% cellspacing=0 cellpadding=0 border=0>
                     <tr>
                         <td style="padding:12px 2px 2px 2px; font-size:11px; width:460px;"><b>Bruttoomsætning:</b><br /><span style="color:#999999; font-size:9px;">Nettooms. + Salgsomk.</span></td>
                         <td align=right style="padding:12px 20px 2px 2px; font-size:11px;">= <input type="text" id="FM_budget" name="FM_budget" value="<%=replace(formatnumber(jo_bruttooms, 2), ".", "")%>" style="width:75px; border:0;" onkeyup="tjektimer('FM_budget'), dbManuel()">&nbsp;&nbsp;&nbsp;&nbsp; <%=basisValISO %></td>
			
                     </tr>
                     </table>

                     </td>
                   	</tr>
                    <%end if %>

                  

                <%end if %>    





                 <%
                 'showMtberegn = 1
                    if cint(showMtberegn) = 1 then
                  
                   netto_bd = "border:0px;"
                   netto_bd2 = "border:0px;"
                   netto_fz = 11
                   netto_fc = "#000000" 
                   else
                   netto_bd = "border:1px red solid;"
                   netto_bd2 = ""
                   netto_fz = 11
                   netto_fc = "#000000"
                   end if%>

				<tr>
                    <td colspan=6 style="border:0px;" valign=top>

                   
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
					    <td style="width:270px; padding:10px 5px 2px 2px; font-size:11px;"><b>Nettoomsætning, timer:</b> <br /><span style="font-size:9px; color:#999999;">Oms. før salgsomk.</span></td>
					    <td valign="bottom" style="font-size:9px;"><span style="color:Red;">*</span> <b>Timer</b></td>
					    <td valign="bottom" style="width:195px; font-size:9px;"><b>Timepris</b></td>
					    <td valign="bottom" style="width:45px; font-size:9px;"><b>Faktor</b></td>
					    <td valign="bottom" style="width:100px; font-size:9px;"><span style="color:Red;">*</span><b> Beløb <%=basisValISO %></b></td>
				        <td valign="bottom" style="width:40px; font-size:8px; color:#999999;">Salgs<br />timepris</b></td>
				    </tr>
				
                        

				<tr bgcolor="#FFFFFF">
					
                    
                    <%'*** Gns. timepris og kostpris på oprettede medarbejdertyper ****' 
						    
					strSQLgt = "SELECT (mt.kostpris * count(m.mid)) AS virkgnskostpris, (mt.timepris * count(m.mid)) AS virkgnssalgspris, count(m.mid) AS antalm"_
					&" FROM medarbejdertyper mt "_
                    &" LEFT JOIN medarbejdere m ON "_
                    &" (m.medarbejdertype = mt.id AND m.mansat <> 2)"_
                    &" WHERE timepris <> 0 AND kostpris <> 0 GROUP BY mt.id"
						    
					'Response.Write strSQLgt
					'Response.flush
                    

                    virkgnssalgspris = 0	    
					virkgnskostpris = 0
					antalM = 0
						    
						    
					oRec2.open strSQLgt, oConn, 3
					while not oRec2.EOF
						    
					virkgnskostpris = virkgnskostpris + oRec2("virkgnskostpris")
                    virkgnssalgspris = virkgnssalgspris + oRec2("virkgnssalgspris")
					antalM = antalM + cint(oRec2("antalm"))
						    
					oRec2.movenext
					wend
					oRec2.close



                    '******* Gns timepris ***'
                    call gnstp_fn(antalM, virkgnskostpris)


                    
                    %>
                    <td id="xpb_jobnavn">Gns. timepris / kostpris: <span style="color:#999999; font-size:9px;"><%=gnsSalgsogKostprisTxt &" "& basisValISO %></span></td>
						  
					<td style="padding:3px;"><input type="text" id="FM_budgettimer" name="FM_budgettimer" value="<%=replace(formatnumber(strBudgettimer, 2), ".", "")%>" style="width:60px;"" onkeyup="tjektimer('FM_budgettimer'), beregnintbelob()" class="nettooms"></td>
					<td class=lille><input type="text" id="FM_gnsinttpris" name="FM_gnsinttpris" value="<%=replace(formatnumber(jo_gnstpris, 2), ".", "")%>" style="width:67px;" onkeyup="tjektimer('FM_gnsinttpris'), beregnintbelob()" class="nettooms"> <%=basisValISO %></td>
					<td>x <input type="text" id="FM_intfaktor" name="FM_intfaktor" value="<%=replace(formatnumber(jo_gnsfaktor, 2), ".", "")%>" style="width:30px;" onkeyup="tjektimer('FM_intfaktor'), beregnintbelob()" class="nettooms"></td>
					<td>= <input type="text" id="FM_interntbelob" name="FM_interntbelob" value="<%=replace(formatnumber(jo_gnsbelob, 2), ".", "")%>" style="width:75px;" onkeyup="tjektimer('FM_interntbelob'), beregninttp()" class="nettooms"></td>
					<input id="FM_interntomkost" name="FM_interntomkost" value="<%=replace(formatnumber(jo_udgifter_intern, 2), ".", "")%>" type="hidden" />
					
                    <%if strBudgettimer <> 0 then
                    tgt_tp = (jo_gnsbelob / strBudgettimer)
                    else
                    tgt_tp = 0
                    end if%>
                             
					<td ><input id="pb_tg_timepris" value="<%=tgt_tp%>" type="text" style="width:30px; font-size:9px; font-family:arial; border:0px;" maxlength="0"/></td>
							
					<!--&nbsp; antal fakturerbare timer.&nbsp;<br>-->
					<input type="hidden" name="FM_ikkebudgettimer_s" value="<%=SQLBless3(ikkeBudgettimer)%>">
														
				</tr>
                  </table>
                <!-- End Netto Table -->

                        <br /><br />


                            <span id="sp_fordeltb" style="color:#5582d2; font-size:11px;"><b>Fordel timebudget på finansår [+]</b></span>

                            <div id="dv_fordeltb" style="display:none; visibility:hidden;">
                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                           <tr>
                            <td style="height:30px;"><b>Budget FY</b><br />
                            
                            </td>
                            <td>Timer</td>
                            <td>Budgetår (FY)</td>
                            <td></td>
                            <td></td>
                            <td></td>
                           </tr>

                            

                                <% 

                                if len(trim(id)) <> 0 then
                                id = id
                                else
                                id = 0
                                end if

                                '** RAMME FY ***'
                                rammeFY0 = 0
                                fctimepris = 0
                                fctimeprish2 = 0 
                                strSQLrammeFY0 = "SELECT timer, fctimepris, fctimeprish2, aar FROM "_
                                &" ressourcer_ramme WHERE jobid = " & id & " AND aktid = 0 AND medid = 0 ORDER BY aar LIMIT 5" 'AND aar = "& year(Y0)    

                                'response.write strSQLrammeFY0
                                'response.flush

                                 fyA = 0
                                 oRec3.open strSQLrammeFY0, oConn, 3
                                 while not oRec3.EOF

                                   

                                    rammeFY0 = oRec3("timer")
                                    fctimepris = oRec3("fctimepris")
                                    fctimeprish2 = oRec3("fctimeprish2")
                                    fyAar = oRec3("aar")

                                    %>
                                       <tr>
                                           <td >År <%=fyA %></td>
                                           <td style="padding:2px;"><input type="text" name="FM_fy_hours_<%=fyA %>" value="<%=rammeFY0 %>" style="width:60px; text-align:right; font-size:9px; font-family:arial;"/></td>
                                           <td style="padding:2px;"><input type="text" name="FM_fy_aar_<%=fyA %>" value="<%=fyAar %>" style="width:60px; text-align:right; font-size:9px; font-family:arial;" /></td>
                                           <td >&nbsp;</td>
                                           <td >&nbsp;</td>
                                           <td >&nbsp;</td>
                                        </tr>
                                <%
                                   
                                     fyA = fyA + 1 
                                 oRec3.movenext
                                 wend
                                 oRec3.close



                                for fyA = fyA TO 4
                                    %>
                                       <tr>
                                         <td>År <%=fyA %></td>
                                            <td style="padding:2px;"><input type="text" name="FM_fy_hours_<%=fyA %>" value="" style="width:60px; text-align:right; font-size:9px; font-family:arial;" /></td>
                                            <td style="padding:2px;"><input type="text" name="FM_fy_aar_<%=fyA %>" value="<%=year(now) + fyA %>" style="width:60px; text-align:right; font-size:9px; font-family:arial;" /> </td>
                                     
                                            <td >&nbsp;</td>
                                            <td >&nbsp;</td>
                                            <td >&nbsp;</td>
                                       </tr>

                                    <%
                                next
                                    %>

                        

                </table>
                  <br /><br />&nbsp;
                  </div>
                       
                <!-- End Netto Table -->

                <input type="hidden" name="antalOpenCalc" id="antalOpenCalc" value="<%=antalOpenCalc %>" />
                <input type="hidden" name="antalMtypGrp" id="antalMtypGrp" value="<%=antalMtypGrp%>" />
                </td></tr>


                <%if cint(showMtberegn) = 0 then %>
                <tr>
					<td colspan=6 style="border:0px;"> <input id="FM_ign_tp" value="1" type="checkbox" /> Åbn for manuel indtastning og beregning af Brutto- og Netto -omsætning.</td>
				</tr>
                <%else 
                    if func = "dbred" OR step = "2" then%>
               <!-- <tr>
					<td colspan=6 style="border:0px; padding:3px 30px 5px 2px;" align="right"><input type="button" value="Opdater bruttooms. >> " id="bt_opdater_brutto" style="font-size:9px;" /><br /><br />&nbsp;</td>
				</tr> -->
                <%
                    end if

                end if


                 '*** Jobtype fastpris / lbn. timer ***
                 select case lto
                   case "epi", "epi_no", "epi_sta", "xintranet - local", "epi_ab", "epi_cati"
                   jt_vzb = "hidden" 'deaktiveret
                   jt_dsp = "none"
                   case else
                   jt_vzb = "visible" 
                   jt_dsp = ""
                   end select %>

                <tr><td colspan=6 valign="top" style="padding:18px 2px 2px 2px; border:0px; visibility:<%=jt_vzb%>; display:<%=jt_dsp%>; font-size:11px;"><b>Jobtype</b><br />
                    
                    <%'select case lto
                    'case "execon"
                    idfp = "xx_angfp0" 'deaktiveret
                    'case else
                    'idfp = "angfp0"
                    'end select %>

                    <input type="radio" id="<%=idfp %>" name="FM_fastpris" value="1" <%=varFastpris1%>> <b>Fastpris</b> (bruttoomsætning benyttes ved fakturering) 
                            

                    

                    <input id="Radio2" name="FM_usejoborakt_tp" type="hidden" value="0" <%=usejoborakt_tp0_CHK %> />

                    
                    <br /><input type="radio" id="angfp1" name="FM_fastpris" value="0" <%=varFastpris2%>> <b>Lbn. timer</b> (timeforbrug på hver enkelt aktivitet * <b>medarb. timepris</b> benyttes ved fakturering)
                    <br />&nbsp;</td>
                 </tr>


              


                 <%if func = "red" then %>

                <tr bgcolor="#FFFFFF">
					<td style="padding:10px 0px 7px 480px; border:0px; border-bottom:3px #D6Dff5 solid;" colspan="6">
                   
                   
                    <br />
		           <span style="padding:5px; background-color:#FFFFFF; border:1px #8caae6 solid; border-bottom:0px;"><a href="#stgrp_tilfoj" id="stgrp_tilfoj" name="stgrp_tilfoj">[+] Tilføj Stam-aktivitetsgrp. til job</a></span>
                    </td>
			    </tr>

                <%end if %>

                 <!--
                    <%if level <= 2 OR level = 6 then %>
		            &nbsp;&nbsp;&nbsp;<a href="akt_gruppe.asp?menu=job&func=favorit" class="kal_g" target="_blank">Konfigurer Stam-aktivitetsgrupper (job-skabeloner) >></a>
                    <%end if %>-->
              

                <!-- <h3>Tilføj Stam-aktiviteter & Faser:</h3><br /> -->
						
				<%if func = "red" then
				stgrpWzb = "hidden"
				stgrpDsp = "none"
				else
				stgrpWzb = "visible"
				stgrpDsp = ""
				end if %>
				<tr bgcolor="#FFFFFF" id="jq_stamaktgrp" style="visibility:<%=stgrpWzb%>; display:<%=stgrpDsp%>;">
					<td colspan="6" style="padding:5px 5px 5px 2px; border:0px #CCCCCC solid;">
                    
                                <div id="tilfojstamdiv" style="padding:5px 5px 5px 2px; border:10px #CCCCCC solid; background-color:#FFFFFF; position:absolute; top:200px; left:20px; height:650px; overflow:auto; width:650px; z-index:20000000;"><!-- AktTD Div -->
                               
                                <div id="jq_stamaktgrp_settings" style="position:relative; padding:0px; left:40px; top:20px; width:550px; border:0px; font-size:11px; background-color:#FFFFFF;">
                                <h4>Tilføj aktiviteter: <span style="font-size:11px; font-weight:normal;">(<a href="#a_indforstamgrp" class=vmenu id="a_indforstamgrp" name="a_indforstamgrp">+ Projektgrupper & medarb. timepriser</a>)</span>
                                <%if func = "red" then %>
                                &nbsp;&nbsp;<a href="#" <a href="#" id="stgrp_luk" class=red>[x]</a>
                                <%end if %>
                                </h4>
                                     <%

                                    uWdt = 400
                                    uTxt = "Vælg den ønskede stam-aktivitetsgruppe og <b>Indlæs gruppe på job</b>. Når gruppen er tilføjet, kan du begynde at estimere og angive timepriser på de forskellige aktiviteter."
                                    'call infoUnisport(uWdt, uTxt)
                                
                                    select case lto
                                    case "epi", "epi_no", "epi_sta", "epi_ab", "epi_cati"
                                    opfodCHK0 = ""
                                    opfodCHK1 = "CHECKED"
                                    case else
                                    opfodCHK0 = "CHECKED"
                                    opfodCHK1 = ""
                                    end select


                                
                                
                                %><br /><br /> 
                                <div id="div_indforstamgrp" style="position:relative; visibility:hidden; display:none; width:550px; padding:10px; border:1px #cccccc dashed;">
                                <b>Projektgrupper:</b><br />

                                    <input id="Radio1" name="pgrp_arvefode" value="0" type="radio" <%=opfodCHK0 %> /><b> 1) Nedarv</b> projektgrupper fra job til de stamaktiviter der tilføjes til jobbet.<br />
                                    <input id="Radio4" name="pgrp_arvefode" value="1" type="radio" <%=opfodCHK1 %> /><b> 2) Fød job</b> med de projektgrupper hver stam-aktivitet er født.
                                    <br /><br />
                                    <b>Medarbejder-timepriser:</b><br />
					                <%
					                '**** Nedarb timepriser fra job eller behold de medarbejder timepriser der er angive på aktiviteterne ***'
					                select case lto
					                case "dencker", "acc", "essens", "fe", "jttek", "wwf", "intranet - local", "nonstop"
					                tpCHK1 = ""
					                tpCHK2 = "CHECKED"
							
					                case "execon", "immenso", "epi", "epi_no", "epi_sta", "epi_ab", "epi_cati"
					                tpCHK1 = "CHECKED"
					                tpCHK2 = ""
							
					                case "jt"
					                tpCHK1 = "CHECKED"
					                tpCHK2 = ""
							
					                case else
					                tpCHK1 = "CHECKED"
					                tpCHK2 = ""
							
					                end select
					                %>
					                <input type="radio" name="FM_timepriser" value="0" <%=tpCHK1 %>> <b>1)</b> Nedarv fra job. (brug den timepris hver <b>medarbejdertype</b> er oprettet med) 
					                <br /><input type="radio" name="FM_timepriser" value="1" <%=tpCHK2 %>> <b>2)</b> Behold de medarbejder-timepriser <b>stam-aktiviteterne</b> er født med. (se stam-aktivitetsgrupper)
					               </div>




                                    </div>

                                    <br /><br />&nbsp;

					<table width=100% cellspacing="0" cellpadding="0" border="0">
						
					<%
	            

                        '***** Antal stamaktgrupper ialt ****'
                        
                        'antalStamGrp = 0
                        'strSQLst = "SELECT count(id) AS antalStamGrp FROM akt_gruppe WHERE id <> 0"
	                    'oRec2.open strSQLst, oConn, 3
	                    'if not oRec2.EOF then
	                    'antalStamGrp = oRec2("antalStamGrp")
	                    'end if
	                    'oRec2.close

	
	            if func <> "red" then
	                    '** Henter forvalgt stam-akt. grupper ***'
	                    for a = 1 to 1 'antalStamGrp - 1
                    	        
	                    forvlgtStamaktgrp = 0
	                    fsnavn = ""
	                    strSQLfv = "SELECT id, forvalgt, navn FROM akt_gruppe WHERE forvalgt = 1 AND skabelontype = 0"
	                    oRec2.open strSQLfv, oConn, 3
	                    if not oRec2.EOF then
	                    forvlgtStamaktgrp = oRec2("id")
	                    fsnavn = oRec2("navn")
	                    end if
	                    oRec2.close
                    	        
	                    call stamakt(2, forvlgtStamaktgrp, fsnavn, a)
                    	        
                    	       
	                    next
                    	
	            else
                    	   
                    	    
	                '*** Mulighed for at tilføje yderligere ved rediger **'
	                fsnavn = ""
	                for a = 1 to 1 'antalStamGrp - 1 
	                    call stamakt(2, 0, fsnavn, a)
	                next
                    	    
	            end if
	

	            %>
	                    
	            </table>

	            </td>
                </tr>
	            <!-- total aktiviteter og faser --->
						
						
				<% 
				if func = "red" then

                    'select case lto
                    'case "epi", "epi_no", "epi_sta", "intranet - local", "epi_ab", "epi_cati"
                    if cint(showMtberegn) = 1 then
                    aktlist_wzb = "hidden"
                    aktlist_dsp = "none"
                    else
                    aktlist_wzb = "visible"
                    aktlist_dsp = ""
                    end if
				
               
                %>
						
						
				
						 
				<tr bgcolor="#D6Dff5">
					<td style="padding:10px 0px 0px 5px;" colspan="6"><h4><a href="#" id="a_aktlisten">[+]</a> Aktiviteter & Faser

                    <%if func = "red" then%>
                     &nbsp;<a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&jobid=<%=id%>&jobnavn=<%=strNavn%>&rdir=<%=rdir%>&nomenu=1', '', 'width=1004,height=800,resizable=yes,scrollbars=yes')" class=rmenu>Rediger aktivitetliste udviddet >> </a>
                    <%end if%>
                    <br />
                    <span style="font-size:9px; color:#000000; line-height:12px;">Budget detaljeret - beregn dit projekt og sync. med nettoomsætning.<br />
                    Ved fastpris job bliver nedenstående overført til fakturering.</span></h4>
                    </td>
                   

				</tr>
                                          <%if func = "red" then %>
                                            <tr class="dv_aktlisten" style="visibility:<%=aktlist_wzb%>; display:<%=aktlist_dsp%>;"><td colspan="6" style="border:0px;">
                                                <input id="jq_vispasogluk" type="checkbox" /> Vis passive og lukkede aktiviteter
                                                </td></tr>
                                          <%end if %>


				                <!-- henter oprettede aktiviter jq -->
                                <%if func = "red" then %>
                                <tr class="dv_aktlisten" style="visibility:<%=aktlist_wzb%>; display:<%=aktlist_dsp%>;"><td colspan="6" style="border:0px;">
              
                                     <%
                     
                                    oleftpx = 20
	                                otoppx = 15
	                                owdtpx = 140
	                                java = "Javascript:window.open('aktiv.asp?menu=job&func=opret&jobid="&id&"&id=0&fb=1&rdir=job3&nomenu=1', '', 'width=1004,height=800,resizable=yes,scrollbars=yes')"    
	                                call opretNyJava("#", "Opret ny aktivitet", otoppx, oleftpx, owdtpx, java) 
                     
                                     %>
                       
                                    <br />&nbsp;
                                </td></tr>

                                <%end if %>

               

				<tr>
                <td colspan="6" style="border:0px;">
                    <div id="jq_aktlisten" class="dv_aktlisten" style="visibility:<%=aktlist_wzb%>; display:<%=aktlist_dsp%>;">Henter aktiviteter...</div>
                </td>
                </tr>
				
						

                <tr class="dv_aktlisten" style="visibility:<%=aktlist_wzb%>; display:<%=aktlist_dsp%>;"><!-- #FFC0CB -->
                    
                    <td style="border:0px;" colspan="6">
                    
                         <table width=100% cellspacing=0 cellpadding=0 border=0>
                         <tr>

					    <td style="padding:5px 50px 5px 5px; border:0px; font-size:11px; text-align:right;">Aktiviteter & Faser ialt: 
                            <span style="padding:2px 2px 2px 2px; font-size:11px; background-color:#FFFFFF; border-bottom:2px #999999 solid;" id="fasertimertot"><b>0,00</span> t.
                        <input id="FM_fasertimertot" value="0" type="hidden" />
					
					     = <span style="padding:2px 2px 2px 2px; font-size:11px; background-color:#FFFFFF; border-bottom:2px #999999 solid;;" id="fasersumtot"><b>0,00</b></span> DKK</td>
                        <input id="FM_fasersumtot" value="0" type="hidden" />
					
				        </tr>
                        </table>

                    </td>
                    </tr>

                    <%if cint(showMtberegn) = 0 then %>

                     <%
                     '** Starter altid med at være utjekket med mindre:
                     select case lto
                     case "synergi1", "xintranet - local"
                     syncCHK = "CHECKED"
                     case else
                     syncCHK = ""
                     end select %>


                     <%
                     '** Starter altid med at være utjekket med mindre:
                     if cint(laasmedtpbudget) = 1 then
                     mtpCHK = "CHECKED"
                     else
                     mtpCHK = ""
                     end if%>

                    <tr class="dv_aktlisten" style="visibility:<%=aktlist_wzb%>; display:<%=aktlist_dsp%>;">
					<td style="padding:20px 0px 10px 2px; border:0px; font-size:11px;" colspan="6">
                      <!--<input id="sync" type="button" value="Sync. forkalkulation >> " />--><input id="sync" type="checkbox" <%=syncCHK %> />Overfør budget på aktiviteter til nettoomsætning på job.
                     
                        <br /><br />
                             <input id="FM_opdmedarbtimepriser" name="FM_opdmedarbtimepriser" value="1" type="checkbox" <%=mtpCHK %> />Lås medarbejdertimepris til budget time-/stk. -pris på aktiviteterne. <span style="color:#999999;">Gælder for alle medarbejdere.</span><br />
                         <b>Opdaterer også ALLE eksisterende timeregistreringer!</b> (Klik "Gem ændringer på job" for at gemme)
                    </td>
                    </tr>
                    
                 
                   

                    <%else %>
                        <tr class="dv_aktlisten" style="visibility:<%=aktlist_wzb%>; display:<%=aktlist_dsp%>;">
					<td style="padding:20px 0px 10px 2px; border:0px;"" colspan="6">
                    &nbsp;</td>
                    </tr>
                    <%end if %>

                    <!--

                    <td style="padding:2px 0px 3px 75px; border:0px;""><span style="padding:2px 2px 2px 0px; background-color:#FFFFFF; width:40px; visibility:hidden; border-bottom:1px #999999 solid;" id="diff_timer"><b><%=formatnumber(diff_timer, 2) %></b></span></td>
                    <input id="FM_diff_timer" name="FM_diff_timer" value="<%=formatnumber(diff_timer, 2) %>" type="hidden" />
					<td colspan="2" style="border:0px;">
                            
                    <if diff_timer <> 0 then
                    df_t_Vzb = "hidden"
                    df_t_Dsp = "none"
                    else
                    df_t_Vzb = "visible"
                    df_t_Dsp = ""
                    end if %>
                            
                    <span id="diff_timer_ok" style="visibility:<%=df_t_Vzb%>; display:<%=df_t_Vzb%>;"><img src="../ill/ok.png" border="0" /></span>&nbsp;</td>
					<td style="padding:2px 0px 3px 40px; border:0px; white-space:nowrap;">= <span style="padding:2px 2px 2px 0px; background-color:#FFFFFF; width:60px; visibility:hidden; border-bottom:1px #999999 solid;" id="diff_sum"><b><%=formatnumber(diff_sum, 2) %></b></span></td>
                    <input id="FM_diff_sum" name="FM_diff_sum" value="<%=formatnumber(diff_sum, 2) %>" type="hidden" />
					<td style="border:0px;">
                             
                    <if diff_sum <> 0 then
                    df_s_Vzb = "hidden"
                    df_s_Dsp = "none"
                    else
                    df_s_Vzb = "visible"
                    df_s_Dsp = ""
                    end if %>
                            
                    <span id="diff_sum_ok" style="visibility:<%=df_s_Vzb%>; display:<%=df_s_Vzb%>;"><img src="../ill/ok.png" border="0" /></span>&nbsp;
                       
                       
                       
                       
                    </div>  
                    &nbsp;</td>
						   
				</tr>
                    -->
                    <!-- AktTD Div -->       


               <tr class="dv_aktlisten" style="visibility:<%=aktlist_wzb%>; display:<%=aktlist_dsp%>;"><td colspan=6 align="right" style="border:0px; padding:0px 10px 20px 5px;">
                    <input class="overfortiljob_u" type="button" value="Gem ændringer på job >>" style="font-family:arial; font-size:9px;"/></td></tr> 
                   
                    <input id="showtpalert" value="1" type="hidden" />
                           
      

                <%else %>
                <div style="visibility:hidden; display:none;">
                <input id="FM_fasertimertot" value="0" type="hidden" />
                <input id="FM_fasersumtot" value="0" type="hidden" />
                <span id="fasertimertot"></span>
                 <span id="fasersumtot"></span>

                 
                 <span id="diff_sum"></span>
                   <span id="diff_timer"></span>
                <input id="FM_diff_timer" value="0" type="hidden" />
                
       
                 <span id="diff_timer_ok"></span>
                <input id="FM_diff_sum" value="0" type="hidden" />
                 <span id="diff_sum_ok"></span>

                </div>
				<%end if%>		


                <%if func = "red" then %>

                	<tr bgcolor="#FFFFFF">
					<td style="height:10px; border:0px;" colspan=6><br />&nbsp;</td>
				</tr>    
				

                <%end if



                if func = "red" OR step = "2" then 

                 if cint(eksisterendeUlevLinier) < 30 then 
                    'atu_Vzb = "visible"
                    'atu_dsp = ""
                    sptu_Vzb = "hidden"
                    sptu_dsp = "none"
                    else
                    'atu_Vzb = "hidden"
                    'atu_dsp = "none"
                    sptu_Vzb = "visible"
                    sptu_dsp = ""
                    end if %>


                

                    <%if func <> "red" then
                        select case lto
                        case "intranet - local", "dencker"
                        ulg_Vsb = "hidden"
                        ulg_Dsp = "none"
                        case else
                        ulg_Vsb = "visible"
                        ulg_Dsp = ""
                        end select
                    else
                    ulg_Vsb = "hidden"
                    ulg_Dsp = "none"
                    end if %>

                      

                <%

                dim selUGrp 
                redim selUGrp(3)
                ufvlgt = 0
                ug = 1
                'for ug = 1 to 1 

                selUGrp(ug) = 0
                call ulev_tilfoj(ufvlgt, ug)
                %>
                       
                            
                    <%strSQLugrp = "SELECT jugrp_id, jugrp_navn, jugrp_forvalgt FROM job_ulev_jugrp WHERE jugrp_id <> 0 ORDER BY jugrp_forvalgt, jugrp_navn"
                            
                    'Response.Write strSQLugpr%>
                
                
                <%
                
                    if func = "red" then%>    
                <tr bgcolor="#FFFFFF"><td colspan=6 style="padding:10px 0px 7px 430px; border-bottom:3px #FFCC66 solid;"><br />
                 <span style="padding:5px; background-color:#FFFFFF; border:1px #FFCC66 solid; border-bottom:0px;">
                        
                        <a href="#" id="tilfoj_ulevgrp">[+] Tilføj Salgsomkostnings grupper til job</a>
                      
                        
                        </span>
                        
                </td></tr>
                <%end if %>
                


                <tr class="jq_ulevgrp" style="visibility:<%=ulg_Vsb%>; display:<%=ulg_Dsp%>;">

                <td colspan=6 style="border:0px; padding:0px 2px 2px 0px; font-size:11px;">

                               

                    <div id="span_tilfoj_ulevgrp" style="visibility:<%=sptu_Vzb%>; width:300px; border:1px red solid; padding:3px; background-color:lightpink; display:<%=sptu_dsp%>;">Der kan ikke tilføjes flere Underlev.- / Salgsomkost. -grupper da der er mere end 30 linier tilføjet allerede.<br />
                    Der kan manuelt tilføjes flere linier ovenfor. (optil 50)</div>

                     
                     <div id="tilfojulevdiv" style="padding:5px 5px 5px 0px; border:10px #CCCCCC solid; background-color:#FFFFFF; position:absolute; top:900px; left:20px; width:650px; z-index:20000000; height:400px; overflow:auto;"><!-- UlevTD Div -->
                     
                     <table cellpadding=0 cellspacing=0 border=0 width=100%>
                     <tr><td style="border:0px; padding:0px 0px 0px 20px;">
                     
                               


                    <br /><br />
                    <h4>Salgsomkostninger: 
                    
                    <span>&nbsp;&nbsp;<a href="#" <a href="#" id="luk_ulevgrp" class=red>[x]</a></span>
                   


                    </h4>
                    <b>Udlæg / Mat. gruppe: </b><br /><select id="jq_seljugrp_id" name="FM_jugrp_id" style="width:200px; font-family:arial; font-size:9px;">
                    <option value="0">Ingen - Vælg gruppe..</option>
                    <%
	                oRec5.open strSQLugrp, oConn, 3
	                while not oRec5.EOF 

                        if cint(ug) = oRec5("jugrp_forvalgt") AND func <> "red" then
                        ugrpSEL = "SELECTED"
                        selUGrp(ug) = oRec5("jugrp_id")
                        else
                        ugrpSEL = ""
                        end if
                    %>
                         
                        <option value="<%=oRec5("jugrp_id")%>" <%=ugrpSEL %>><%=oRec5("jugrp_navn")%></option>
                           

                    <%

                    lastUgrpId = oRec5("jugrp_id")
                    oRec5.movenext
                    wend
                    oRec5.close
                    %> </select>
                        <br /><br />&nbsp;

                    
               


                    </td></tr>

                <%
                        
                        
                ulgrp = 0
                uheadisWrt = 0
                for ulgrp = 0 to UBOUND(strUlevTxt) %>
                        
                <%if len(trim(strUlevTxt(ulgrp, ug))) <> 0 then 
                            
                    if cint(strUlevSEL(ulgrp, ug)) = cint(selUGrp(ug)) then
                    ulevLinieVzb = "visible"
                    ulevLinieDsp = ""
                    else
                    ulevLinieVzb = "hidden"
                    ulevLinieDsp = "none"
                    end if

                    %>
                    <tr bgcolor="#FFFFFF" style="visibility:<%=ulevLinieVzb%>; display:<%=ulevLinieDsp%>;" class="ulev_pb_<%=ulgrp%>">
                    <td colspan="6" style="border-top:0px;">


                    <table border="0" cellpadding="2" cellspacing="0" width="100%">

                    <%'if uheadisWrt = 0 then %>
                    <tr bgcolor="#FFCC66">
                    <td>&nbsp;</td>
		          
		            <td style="padding:2px 0px 2px 2px;"><b>Udgift navn / txt.</b></td>
		            <td style="padding:2px 0px 2px 2px;"><b>Stk. / stk. pris </b></td>
                    <td style="padding:2px 20px 2px 2px;" align=right><b>Indkøbspris</b></td>
		            <td style="padding:2px 20px 2px 2px;" align=right><b>Faktor</b></td>
		            <td style="padding:2px 20px 2px 2px;" align=right><b>Salgspris <%=basisValISO %></b></td>
		                 
		            </tr>
                    <%
                            
                    'uheadisWrt = 1
                    'end if %>

                    <%=strUlevTxt(ulgrp, ug) %>
                    </table>


                   

                </td>
                </tr>
                    <tr style="visibility:<%=ulevLinieVzb%>; display:<%=ulevLinieDsp%>;" class="ulev_pb_<%=ulgrp%>"><td colspan=6 align="right" style="border:0px; padding:10px 10px 5px 5px;">
                    <input class="overfortiljob_u" type="button" value="Indlæs gruppe på job >>" style="font-family:arial; font-size:9px;"/></td></tr> 
		               
                      
                <%end if %>

                <%next
                        
                        
                'next
	            
                if len(trim(uopen)) <> 0 then
                uopen = uopen
                else
                uopen = 0
                end if          

                
                
                
                
                
                
                'if func = "red" then

                         
                '*** Ulev ***'
		        dim u_navn, u_ipris, u_faktor, u_belob, u_fase, u_id, u_istk, u_istkpris
		        redim u_navn(250), u_ipris(250), u_faktor(250), u_belob(250), u_fase(250), u_id(250), u_istk(250), u_istkpris(250)
		                
		            if func = "red" then
		                        
		                uopen = 0
		                        
		                strSQLUlev = "SELECT ju_id, ju_navn, ju_ipris, ju_faktor, "_
		                &" ju_belob, ju_fase, ju_stk, ju_stkpris FROM job_ulev_ju WHERE ju_jobid = "& id &" ORDER BY ju_navn "
		                       
                        'Response.Write strSQLUlev & "<br>"
                               
                        oRec2.open strSQLUlev, oConn, 3
		                        
		                u = 1
		                        
		                while not oRec2.EOF 
		                     
		                    u_navn(u) = oRec2("ju_navn")
		                            
		                    if len(trim(u_navn(u))) then
		                    uopen = uopen + 1
		                    else
		                    uopen = uopen
		                    end if
		                            
                            u_id(u) = oRec2("ju_id")
		                    u_ipris(u) = oRec2("ju_ipris")
                            u_faktor(u) = oRec2("ju_faktor")
		                    u_belob(u) = oRec2("ju_belob")
		                    u_fase(u) = oRec2("ju_fase")
		                    u_istk(u) = oRec2("ju_stk")       
		                    u_istkpris(u) = oRec2("ju_stkpris")       
                                    
		                    u = u + 1
		                            
		                    oRec2.movenext
		                    wend
		                    oRec2.close
		                         
		                    if uopen = 0 then
		                    uopen = 1
		                    end if
		                    
		                  
		            else
		                        
		                for u = 1 to 50 
		                     
		                u_id(u) = 0
		                u_navn(u) = ""
		                u_ipris(u) = 0
		                u_faktor(u) = 1
		                u_belob(u) = 0
                        u_fase(u) = ""
		                u_istk(u) = 1       
		                u_istkpris(u) = 0        
		                next
		                        
		                uopen = 1
		                        
		            end if
		               
                         
                         
                    %>
                        
                     
                     
                      </td></tr></table>

                     
                        </div> <!-- ulevTDDiv -->
                        
                        
              <%'end if %>

                         	</td>
				</tr>

			   
                <% if func = "red" then%>
				<tr bgcolor="#FFCC66">
					<td style="padding:10px 0px 0px 5px;" colspan=6><h4><a href="#" id="a_salgsomk">[+]</a> Salgsomkostninger<br />
                    <span style="font-size:9px; color:#000000; line-height:12px;">Underleverandører / Materialeforbrug</span></h4> </td>
                 </tr>

                 <tr id="tr_salgsomk" class="tr_salgsomk" bgcolor="#FFFFFF">
					<td style="padding:3px 0px 3px 10px;" colspan=6>   
                     Antal linier: 
                        <select id="antalulev" onchange="showulevlinier()" style="font-size:9px; font-family:arial;">
                                    
                            <%for u = 1 to 50 
                                        
                                if u = uopen then
                                usel = "SELECTED"
                                else
                                usel = ""
                                end if
                                    
                            %>
                                    
                            <option value="<%=u %>" <%=usel %>><%=u %></option>
                                    
                            <%next %>
                                   
                        </select> (maks. 50 linier)</td>
				</tr>  
						
                   <tr class="tr_salgsomk" bgcolor="#FFFFFF">
                 <td colspan=6>
                    <table cellpadding="0" cellspacing="0" border=0 width=100%>
                        
		            <tr>
		            <td style="padding:10px 0px 3px 2px;"><b>* Udgift/Salgsomkost.</b></td>
                    <td style="padding:10px 0px 3px 2px;"><b>Stk.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Stk. pris.</b></td>
		            <td align="right" style="padding:10px 20px 3px 2px;"><b>Indkøbspris</b></td>
		            <td style="padding:10px 10px 3px 2px;"><b>Faktor</b></td>
		            <td style="padding:10px 0px 3px 10px;"><b>Salgspris <%=basisValISO %></b></td>
		            <td>&nbsp;</td>
		        </tr>

                
                    
                    
                    
                    
                    
		                
		        <%
                jo_salgspris_ulev = 0
		        eksisterendeUlevLinier = 1
		        for u = 1 to 50 
		                
		        select case right(u, 1)
		        case 0,2,4,6,8
		        bgulev = "#FFFFFF"
		        case else
		        bgulev = "#FFFFe1" '"#EFF3FF"
		        end select
		                
		            if u > uopen then
		            uvzb = "hidden"
		            udsp = "none"
		            else
                    eksisterendeUlevLinier = eksisterendeUlevLinier + 1 
		            uvzb = "visible"
		            udsp = ""
		            end if
		                 
		            if len(trim(u_id(u))) <> 0 then
                    u_id(u) = u_id(u)
                    else
                    u_id(u) = 0
                    end if
		                
		        %>
		                
		                
		            <tr bgcolor="<%=bgulev %>" id="ulevlinie_<%=u%>" style="visibility:<%=uvzb%>; display:<%=udsp%>;">


                        <input id="ulevid_<%=u %>" name="ulevid_<%=u %>" value="<%=u_id(u) %>" type="hidden" />
                        <input type="hidden" id="Hidden1" name="ulevfase_<%=u%>" value="<%=u_fase(u) %>">
		           
		            <td ><input type="text" id="ulevnavn_<%=u%>" name="ulevnavn_<%=u%>" value="<%=u_navn(u) %>" style="width:220px; font-size:9px; font-family:arial;"></td>
		            <td><input type="text" class="ulev" id="ulevstk_<%=u%>" name="ulevstk_<%=u%>" value="<%=replace(formatnumber(u_istk(u), 2), ".", "") %>" style="width:45px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevstk_<%=u%>'), beregnulevstkpris('<%=u%>')"> stk. á &nbsp;
					<input type="text" class="ulev" id="ulevstkpris_<%=u%>" name="ulevstkpris_<%=u%>" value="<%=replace(formatnumber(u_istkpris(u), 2), ".", "") %>" style="width:60px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevstkpris_<%=u%>'), beregnulevstkpris('<%=u%>')"></td>
					
                    <td>= <input type="text" class="ulev" id="ulevpris_<%=u%>" name="ulevpris_<%=u%>" value="<%=replace(formatnumber(u_ipris(u), 2), ".", "") %>" style="width:60px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevpris_<%=u%>'), beregnulevbelob('<%=u%>')"></td>
					<td>x <input type="text" class="ulev" id="ulevfaktor_<%=u%>" name="ulevfaktor_<%=u%>" value="<%=replace(formatnumber(u_faktor(u), 2), ".", "") %>" style="width:40px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevfaktor_<%=u%>'), beregnulevbelob('<%=u%>')"></td>
		            <td align=right style="padding-right:33px;">= <input type="text" id="ulevbelob_<%=u%>" name="ulevbelob_<%=u%>" value="<%=replace(formatnumber(u_belob(u), 2), ".", "") %>" style="width:60px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevfaktor_<%=u%>'), beregnulevipris('<%=u%>')"></td>
		            <td>&nbsp; <a href="#" id="ulev_ryd_<%=u%>" class="ulev_ryd">Ryd</a> </td>
		        </tr>
		                
		        <%
                    jo_salgspris_ulev = jo_salgspris_ulev + u_belob(u)
                    next %>
		            
                    </table>
                    
                    </td>
                </tr>

    
                    <tr class="tr_salgsomk">
                    <td align="right" style="padding:10px 5px 3px 5px; border:0px;" colspan="6"><a href="#" class="rmenu" id="ulev_tilfoj_line">Tilføje line >></a>
                    </td>
                    </tr>
                       
                       
                <%
                
                end if 'red
                
                end if 'if func = "red" OR step = "2" then 
                
               %>

                       <input id="ulevopen" value="<%=uopen %>" type="hidden" />
                        <input id="lastopen_ulev" value="0" type="hidden" />
                        

                       
                 
              
						
						
						
				
               
            

		        <tr bgcolor="#FFFFFF">
					<td class=alt style="padding-top:10px; border:0px;" colspan=6>

                <table width=100% border="0" cellpadding=0 cellspacing=0>
		        <tr>
						    
					<td colspan=4 style="padding:10px 0px 3px 5px; width:460px; font-size:11px; font-family:arial;"><b>Bruttoomsætning:</b></td>
					<td style="padding:2px 2px 2px 20px; width:100px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px #999999 solid; font-size:9px; font-family:arial;" id="SP_budget"><b><%=replace(formatnumber(jo_bruttooms, 2), ".", "")%></b></span></td>
							
					<!--&nbsp; antal fakturerbare timer.&nbsp;<br>-->
					<input type="hidden" name="FM_ikkebudgettimer" value="<%=SQLBless3(ikkeBudgettimer)%>">
                    <!--&nbsp;antal ikke fakturerbare timer.<br>&nbsp;-->
                    <!--<input id="FM_budget" name="FM_budget" value="<%=replace(formatnumber(jo_bruttooms, 2), ".", "")%>" type="hidden" />-->
						<td style="width:25px;">&nbsp;<%=basisValISO %></td>
							
		        </tr>
		        
                 <tr>
						    
					<td colspan=4 style="padding:10px 0px 3px 5px; width:460px; font-size:11px; font-family:arial;">Nettoomsætning: (timer)</td>
					<td style="padding:2px 2px 2px 20px; width:100px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px #999999 solid; font-size:9px; font-family:arial;" id="SP_netto"><b><%=replace(formatnumber(jo_gnsbelob, 2), ".", "")%></b></span></td>
							
				
						<td style="width:25px;">&nbsp;<%=basisValISO %></td>
							
		        </tr>

                       <tr>
						    
					<td colspan=4 style="padding:10px 0px 3px 5px; width:460px; font-size:11px; font-family:arial;">Varesalg: (salgspris på salgsomkostninger)</td>
					<td style="padding:2px 2px 2px 20px; width:100px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px #999999 solid; font-size:9px; font-family:arial;" id="SP_salgspris_ulev"><b><%=replace(formatnumber(jo_salgspris_ulev, 2), ".", "")%></b></span></td>
                           <input id="FM_salgspris_ulev" name="FM_salgspris_ulev" value="<%=replace(formatnumber(jo_salgspris_ulev, 2), ".", "")%>" type="hidden" />
							
				
						<td style="width:25px;">&nbsp;<%=basisValISO %></td>
							
		        </tr>

                    <tr bgcolor="#ffffff">
					<td colspan=4 style="padding:30px 0px 3px 5px; font-size:11px; font-family:arial;"><b>Udgifter ialt:</b> (intern kost. + salgsomk.)</td>
					<td style="padding:30px 2px 2px 20px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px; border-bottom:0px #999999 solid; font-size:9px; font-family:arial;" id="SP_udgifter"><b><%=replace(formatnumber(udgifter, 2), ".", "")%></b></span></td>
                        <input id="FM_udgifter" name="FM_udgifter" value="<%=replace(formatnumber(udgifter, 2), ".", "")%>" type="hidden" />
                    <td style="padding:30px 2px 2px 0px;">&nbsp;<%=basisValISO %></td>
				</tr>
                        
                <tr bgcolor="#ffffff">
					<td colspan=4 style="padding:10px 0px 3px 5px; font-size:11px; font-family:arial;">Intern kost.: (kostpris pr. time * timer)
                        
                         <%if cint(showMtberegn) = 1 then 
                             

                             if cint(bdgmtypon_val) = 1 then 

                             %><%
            
                                for t = 1 to UBOUND(mtypgrpids)

                                'Response.write "mtypgrpnavn(t): "& mtypgrpnavn(t) & "<br>"

                                if len(trim(mtypgrpnavn(t))) <> 0 then
           
         
            
                                %>
        
		                    <span style="color:#000000; font-size:9px;" id="span1"><%=mtypgrpnavn(t) %>:</span> <span style="color:#000000; font-size:9px;" id="span_mtype_totudspec_<%=mtypgrpid(t)%>"><%=formatnumber(thisMtypGrpBel, 2)%></span>
		
                           <%
                                end if
                                next%>
                                 
                            <%
                            end if%>

                      <%end if %>


                        
                       </td>
					<td style="padding:2px 2px 2px 20px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px; border-bottom:0px #CCCCCC solid; font-size:9px; font-family:arial;" id="SP_udgifter_intern"><%=replace(formatnumber(jo_udgifter_intern, 2), ".", "")%></span></td>
					<td>&nbsp;<%=basisValISO %></td>
                   <input id="FM_udgifter_intern" name="FM_udgifter_intern" value="<%=replace(formatnumber(jo_udgifter_intern, 2), ".", "")%>" type="hidden" />
				</tr>
                
                   

                <tr bgcolor="#ffffff">
					<td colspan=4 style="padding:10px 0px 3px 5px; font-size:11px; font-family:arial;">Salgsomkostninger:</td>
					<td style="padding:2px 2px 2px 20px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px; border-bottom:0px #CCCCCC solid; font-size:9px; font-family:arial;" id="SP_udgifter_ulev"><%=replace(formatnumber(jo_udgifter_ulev, 2), ".", "")%></span></td>
					<input id="FM_udgifter_ulev" name="FM_udgifter_ulev" value="<%=replace(formatnumber(jo_udgifter_ulev, 2), ".", "")%>" type="hidden" />
                    <td>&nbsp;<%=basisValISO %></td>
				</tr>
		                
                
                      <tr bgcolor="#ffffff">
					<td colspan=6 style="padding:30px 0px 3px 5px; font-size:11px; font-family:arial;">&nbsp;</td></tr>
		                
		            <tr bgcolor="#Eff3ff">
						    
					<td colspan=4 style="padding:10px 0px 3px 5px; font-size:11px; font-family:arial;"><b>Dækningsbidrag / Bruttofortjeneste:</b> (bruttooms. - udgifter)</td>
					<td style="padding:2px 2px 2px 20px;">= <span style="padding:2px 2px 2px 2px; background-color:#Eff3ff; width:60px; border:0px; border-bottom:0px #999999 solid; font-size:9px; font-family:arial;" id="SP_bruttofortj"><b><%=replace(formatnumber(jo_bruttofortj, 2), ".", "") %></b></span></td>
						<input id="FM_bruttofortj" name="FM_bruttofortj" value="<%=replace(formatnumber(jo_bruttofortj, 2), ".", "") %>" type="hidden" />
                        <td>&nbsp;<%=basisValISO %></td>
				</tr>
				<tr bgcolor="#ffffff">
						    
					<td colspan=4 style="padding:10px 0px 3px 5px; font-size:11px; font-family:arial;"><b>Dækningsbidrag (DB) %</b></td>
					<td style="padding:2px 2px 2px 20px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px; border-bottom:0px #999999 solid; font-size:9px; font-family:arial;" id="SP_db"><b><%=formatnumber(jo_dbproc, 0) %></b></span></td>
					<input id="FM_db" name="FM_db" value="<%=formatnumber(jo_dbproc, 0)%>" type="hidden" />
                    <td>&nbsp;%</td>
				</tr>
                </table> 

                &nbsp;</td>
				</tr>  
                
                <tr>
                <td align=right colspan=6 style="border:0px; padding:15px 15px 5px 5px;"> <input class="overfortiljob_u" type="button" value="Gem ændringer på job >>" style="font-family:arial; font-size:9px;"/></td>
                </tr>
                
               

                


                  


                  <%if func <> "red" then %>
                       <tr>
                <td align=right colspan=6 style="border:0px; padding:25px 15px 5px 5px; height:30px;">&nbsp;</td>
                </tr>
                <%end if %>
		                
				</table>


                <%if func = "red" then %>
                <br /><br /><br />
					&nbsp;
              


				
							
						
						    
						    
				


                    <table width="100%" cellspacing=0 cellpadding=0 border=0>
                    <tr>
                    <td colspan=4 align="right" style="padding:20px 15px 10px 10px; border-top:1px #cccccc solid;">
                    <input id="Submit3" type="submit" value=" <%=submtxt %>  >> " <%=sendmailtokunde%> /></td>
                    </tr>
                    </table>

              <%end if %>





              <%if step = "2" then %>
              <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
              <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
              <table width=100% cellspacing=0 cellpadding=0>
                <tr>
                <td align=right colspan=6 style="border:0px; padding:15px 15px 5px 5px;"> <input class="overfortiljob_u" type="button" value=" Videre >> " /></td>
                </tr>
              </table>
              <%end if %>
</div>
<!-- projekt beregner slut -->
<%
end if


end sub



sub minioverblik



	if thisfile <> "jobprintoverblik" then	
    oTop = 0
	oLeft = 487
	oWdth = 700
    else
    oTop = 20
	oLeft = 20
	oWdth = 700

    end if

	oHgt = "" 
	oId = "oblik_div"
	oZindex = 1000

	call tableDivAbs(oTop,oLeft,oWdth,oHgt,oId, oVzb, oDsp, oZindex)
	    
        %>
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <%
        
        if thisfile <> "jobprintoverblik" then
        %>
		
       
        <tr bgcolor="#5582d2">
						    <td colspan=6 class=alt style="padding:5px 5px 0px 5px; border:0px; height:32px;" valign=top><h3 class="hv">Joboverblik & Funktioner</h3></td>
						</tr>

                           <tr>
                <td align=right colspan=6 style="border:0px; padding:15px 15px 5px 5px;"> <input class="overfortiljob_o" type="button" value="Gem ændringer på job >>" style="font-family:arial; font-size:9px;"/></td>
                </tr>
		
		<tr>
		<td style="padding:10px 4px 5px 4px;">
	    <%
        
        
        
     
		


        'minusyear = dateadd("m", -12, now)
		
		'dtlink_stdag = datepart("d", minusyear, 2, 2)
		'dtlink_stmd = datepart("m", minusyear, 2, 2)
		'dtlink_staar = datepart("yyyy", minusyear, 2, 2)
		
        select case lto
        case "synergi1", "intranet - local"


        call licensStartDato()
        dtlink_stdag = startDatoDag
		dtlink_stmd = startDatoMd
		dtlink_staar = startDatoAar

        case else
        
        dtlink_stdag = datepart("d", strTdato, 2, 2)
		dtlink_stmd = datepart("m", strTdato, 2, 2)
		dtlink_staar = datepart("yyyy", strTdato, 2, 2)

		
        end select
		
        
        dtlink_sldag = datepart("d", now, 2, 2)
		dtlink_slmd = datepart("m", now, 2, 2)
		dtlink_slaar = datepart("yyyy", now, 2, 2)
		
		dtlink = "FM_usedatokri=1&FM_start_dag="&dtlink_stdag&""_
	    &"&FM_start_mrd="&dtlink_stmd&"&FM_start_aar="&dtlink_staar&"&FM_slut_dag="&dtlink_sldag&""_
	    &"&FM_slut_mrd="&dtlink_slmd&"&FM_slut_aar="&dtlink_slaar
		
		%>
		
        
        <a href="jobprintoverblik.asp?menu=job&id=<%=id%>&media=printjoboverblik" class=vmenu target="_blank">Print joboverblik >></a><br />
      
         <%call bdgmtypon_fn()
         if cint(bdgmtypon_val) = 1 then %>
         <a href="mtypgrp_real_konsolider.asp?dothis=1&first=2&jobid=<%=id %>" class=vmenu target="_blank">Konsolider job >></a><br />
		 <%end if %>   
            
       
		
        <%
         '*** projektleder rretigheder   
         select case lto
         case "oko"

            if level = 1 then
            sltjjobok = 1
            sltjfilok = 0
            sltjfakok = 0 
            else
            sltjjobok = 0
            sltjfilok = 0
            sltjfakok = 0 
            end if

          case else
            
           sltjjobok = 1
           sltjfilok = 1
           sltjfakok = 1  

          end select %>


        <%if cint(sltjjobok) = 1 then %>
        <a href="job_kopier.asp?func=kopier&id=<%=id%>&fm_kunde=<%=fm_kunde_sog%>&filt=<%=request("filt")%>" class=vmenu>Kopier job >></a>
        <br /><a href="jobs.asp?menu=job&func=slet&id=<%=id %>" class=slet>Slet / nulstil job?</a><br />
        <%end if %>

        <%else%>
        
            <% if thisfile <> "jobprintoverblik" then %>
       <tr bgcolor="#5582d2">
						    <td colspan=6 class=alt style="padding:5px 5px 0px 5px; border:0px;"><h3 class="hv">TimeOut - Joboverblik</h3></td>
						</tr>
            <%end if %>


		<tr>
		<td style="padding:10px 4px 5px 4px;">


            <% if thisfile <> "jobprintoverblik" then %>
            <%=""& strKnavn & " ("& strKnr &")<br><b>"& strNavn & "</b> ("& strjobnr &")<br /><br />"  %>
            <%else %>
            <%="<h4><span style=font-size:11px;>"& strKnavn & " ("& strKnr &")</span><br>"& strNavn & " ("& strjobnr &")</h4>"%>
            <%end if %>

        Status: <%=strStatusNavn %>
         <br>Periode: <span style="color:green;"><%=formatdatetime(strTdato, 1)%></span> - <span style="color:red;"><%=formatdatetime(strUdato, 1)%></span>
	 <br />
        Jobansvarlige: <span style="color:#999999;">
    <%
		
						'*** Jobansvarlige ***
                        if jobans1 <> 0 then
						call meStamData(jobans1)
                        %>
						<%=meNavn%> 
                             <%if jobans_proc_1  <> 0 then %>
                            (<%=jobans_proc_1 %>%)
                            <%end if %>

                        <%end if
                        

                        '*** Jobejer 2 ***
                        if jobans2 <> 0 then
						call meStamData(jobans2)
                        %>
						, <%=meNavn%> 
                        
                        <%if jobans_proc_2  <> 0 then %>
                        (<%=jobans_proc_2 %>%)
                        <%end if %>

                        <%end if

                        '*** Job medansvarlige ***
                        if jobans3 <> 0 then
						call meStamData(jobans3)
                        %>
						, <%=meNavn%> 
                         <%if jobans_proc_3  <> 0 then %>
                        (<%=jobans_proc_3 %>%)

                        <%end if
                        
                        end if


                        '*** Job medansvarlige ***
                        if jobans4 <> 0 then
						call meStamData(jobans4)
                        %>
						, <%=meNavn%> 
                             <%if jobans_proc_4 <> 0 then %>
                            (<%=jobans_proc_4 %>%)
                            <%end if %>

                        <%end if


                        '*** Job medansvarlige ***
                        if jobans5 <> 0 then
						call meStamData(jobans5)
                        %>
						, <%=meNavn%> 
                             <%if jobans_proc_5 <> 0 then %>
                            (<%=jobans_proc_5 %>%)
                            <%end if %>

                        <%end if

                   
                      
						
						'**********************
						%></span>


                        <%'*** kontakpersoner hos kunde ***'
                         
                         kpersNavn = ""
                         if cint(kundekpers) <> 0 then

                         strSQLkpers = "SELECT navn, email FROM kontaktpers WHERE id = " & kundekpers
                         oRec6.open strSQLkpers, oConn, 3
                         if not oRec6.EOF then

                         kpersNavn = oRec6("navn") & ", email: "& oRec6("email")

                         end if
                         oRec6.close

                         %>
                         <br />Kontaktpers. hos kunde: <span style="color:#5582d2;"><%=kpersNavn %></span>
                         <%

                         end if
                         
                         %>
							
						
	                   
                       
	                    <% 
	                    if len(trim(rekvnr)) <> 0 then
	                    Response.Write "<br>Rekvnr.: "& rekvnr
	                    end if

	
	                    if cint(intServiceaft) <> 0 then
	                    Response.Write "<br>Aftale: "& aftnavn & " ("& aftnr &")" 
	                    end if
	
	                    %>
                     




        <%end if%>


        </td>
        </tr>


        <%if intprio > -1 then '(thisfile <> "jobprintoverblik" AND lto = "synergi1")  AND 'OR lto <> "synergi1")%>
        <tr>
		    <td><br /><b>Timeforbrug</b> (overblik) 
                <%if cint(strStatus) <> 1 then 'lukket/passivt mv.%>
                <span style="color:darkred;"> - Seneste 15 md.</span>
                <%end if%>

		    </td>
		</tr>
        
        <tr>
        <td bgcolor="#FFFFFF" style="padding:4px 4px 5px 4px; border:1px #cccccc solid;" align=center>

         <table cellpadding=0 cellspacing=0 border=0 width=100%>
         <tr>
         <%call akttyper2009(2) %>

         <td valign=top class="lille">

             <!-- fordeling på medarbtyper --->

        <%
            call fn_medarbtyper()
            
            if cint(strStatus) <> 1 then 'lukket/passivt mv.
            'Viser hele job peridoen, ignorerdato
            mTypStDato = strTdato
            mTypSlDato = strUdato
            mtypVisningPer = 0
            else
            'Viser seneste 15 md
            mtypVisningPer = 1
            '& "/"& month(strTdato) &"/"& day(strTdato)
            mTypSlDato = now 'year(now) & "/"& month(now) &"/"& day(now)
            mTypStDato = dateAdd("m", -15, mTypSlDato) 'strTdato
            end if
            
            call timer_fordeling_medarb_typer(strjobnr, mtypVisningPer, 100, mTypStDato, mTypSlDato, id)
            
            forstetimereg = "1-1-2002"
            strSQlforstetimereg = "SELECT tdato FROM timer WHERE tjobnr = '"& strjobnr &"' ORDER BY tdato limit 1"
            oRec6.open strSQlforstetimereg, oConn, 3
            if not oRec6.EOF then
            forstetimereg = oRec6("tdato")  
            end if
            oRec6.close
        
            if cdate(forstetimereg) <> "1-1-2002" then
            %>
            <br />Første timreg.: <%=formatdatetime(forstetimereg, 2) %>
            <%
            end if
            %>
         </td>
         
         
         <td class="lille" valign=top>
     
		Timeforbrug fordelt på måneder:
        <table cellpadding=0 cellspacing=0 border=0>
        <tr bgcolor="#FFFFFf">
         <td style="width:40px; border-bottom:1px #cccccc solid; border-right:1px #cccccc solid; padding:2px; font-size:8px;" valign=top>Kvo.: <br /><%=hgtKvo%></td>
            <%
            

            '*** StDato ***'
            if strStatus <> 1 AND cDate(strUdato) <> cdate("1-1-2044") then 'passiv / lukket og uendelig

                    if len(trim(lkdato)) <> 0 AND cDate(lkdato) <> "01-01-2002" then
                    stDatoKriGrf = lkdato 'strUdato
                    else
                    stDatoKriGrf = strUdato
                    end if    

                    stDatoKriGrfDiff = datediff("m", strTdato, stDatoKriGrf, 2, 2) 'strUdato
            
                    if stDatoKriGrfDiff > 0 AND stDatoKriGrfDiff <= 15 then 
                    mts = stDatoKriGrfDiff
                    else
                    mts = 15
                    end if

                    else

            stDatoKriGrf = now
            
            mts = 15
            end if

            if (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati") AND strjobnr = "3" then
            mts = 3
            end if


            'response.write "mts" & mts & " stDatoKriGrfDiff: " & stDatoKriGrfDiff & " strTdato "& strTdato &" lkdato: "& lkdato & " cDate(lkdato) "& cDate(lkdato)
            'response.flush

            dim timerThis, kostThis, omsThis, resThis
            redim timerThis(mts), kostThis(mts), omsThis(mts), resThis(mts)
            
            
            for dt = 0 TO mts
             
             if dt <> 0 then
             dtnowHigh = dateadd("m", 1, dtnowHigh)
             else
             dtnowHigh = dateadd("m", -(mts), stDatoKriGrf)
             end if

            

            %>
            <td colspan=4 style="font-size:8px; border-bottom:1px #cccccc solid; width:24px;" align=center><%=left(monthname(month(dtnowHigh)), 3) &"<br> "& right(year(dtnowHigh), 2) %></td>
            <%next %>
         </tr>

         <tr bgcolor="#FFFFFF">
          <td style="border-bottom:1px #cccccc solid; border-right:1px #cccccc solid; padding:2px; font-size:8px;" valign=top><u><%=hgtKvoTxt %></u></td>
            <%for dt = 0 TO mts
             if dt <> 0 then
             dtnowHigh = dateadd("m", 1, dtnowHigh)
             else
             dtnowHigh = dateadd("m", -mts, stDatoKriGrf)
             end if

             select case month(dtnowHigh)
             case 1,3,5,7,8,10,12
                eday = 31
                case 2
                    select case right(year(dtnowHigh), 2)
                    case "00", "04", "08", "12", "16", "20", "24", "28", "30", "34", "38", "42", "46"
                    eday = 29
                    case else
                    eday = 28
                    end select
             case else
                eday = 30
             end select

             dtnowHighSQL = year(dtnowHigh) &"/"& month(dtnowHigh) &"/"& eday
             dtnowLowSQL = year(dtnowHigh) &"/"& month(dtnowHigh) &"/1"
               
               omsHgt = 0
               kostHgt = 0
               hgt = 0
               resHgt = 0
               timerThis(dt) = 0
               kostThis(dt) = 0
               resThis(dt) = 0

               if budgettimertot <> 0 then
               budgettimertot_use = budgettimertot
               else
               budgettimertot_use = 1
               end if

             strSQLjl = "SELECT timer, (kostpris * timer * (kurs / 100)) AS kostpris, timepris, tfaktim FROM timer WHERE tjobnr = '"& strjobnr &"' AND tdato BETWEEN '"& dtnowLowSQL &"' AND '"& dtnowHighSQL &"'"_
             &" AND ("& aty_sql_realhours &") ORDER BY tdato"

             oRec2.open strSQLjl, oConn, 3
             while not oRec2.EOF

 

             hgt = hgt + formatnumber(oRec2("timer")/ 2, 0) 
             timerThis(dt) = timerThis(dt) + oRec2("timer")

             if cint(jo_fastpris) = 1 then
             tpThis =  jo_bruttooms / budgettimertot_use
             else
             tpThis = oRec2("timepris")
             end if 'lbn timer
             
             call akttyper2009Prop(oRec2("tfaktim"))

             if aty_fakbar = 1 then
             omsHgt = omsHgt + formatnumber((oRec2("timer") * tpThis) / 1000, 0)
             omsThis(dt) = omsThis(dt) + (oRec2("timer") * tpThis)
             else
             omsHgt = omsHgt
             omsThis(dt) = omsThis(dt)
             end if

             kostHgt = kostHgt + formatnumber(oRec2("kostpris")/ 1000, 0)
             kostThis(dt) = kostThis(dt) + oRec2("kostpris")


             


             oRec2.movenext
             wend
             oRec2.close

             

            '*** Ressource timer ****'
             strSQLres = "SELECT timer FROM ressourcer_md "_
             &" WHERE jobid = "& id &" AND (aar = YEAR('"& dtnowLowSQL &"') AND md = MONTH('"& dtnowLowSQL &"'))"

             'Response.Write strSQLres
             'Response.flush
             oRec2.open strSQLres, oConn, 3
             while not oRec2.EOF 

             resHgt = resHgt + formatnumber(oRec2("timer")/ 2, 0)
             resThis(dt) = resThis(dt) + oRec2("timer")
             'Response.Write resThis(dt)  & "<br>"

             oRec2.movenext
             wend 
             oRec2.close

             '** tilpasser graf **'
             hgt = formatnumber(hgt/hgtKvo, 0)  
             resHgt = formatnumber(resHgt/hgtKvo, 0)  
             kostHgt = formatnumber(kostHgt/hgtKvo, 0)  
             omsHgt = formatnumber(omsHgt/hgtKvo, 0)  

             if cdbl(resHgt) >= 100 then
             resHgt = 100
             resbgThis = "#999999"
             else
             resHgt = resHgt
             resbgThis = "#999999"
             end if
             
             if cdbl(kostHgt) >= 100 then
             kostHgt = 100
             kostbgThis = "#8caae6"
             else
             kostHgt = kostHgt
             kostbgThis = "#8caae6"
             end if


             if cdbl(omsHgt) >= 100 then
             omsHgt = 100
             omsbgThis = "#9aCD32"
             else
             omsHgt = omsHgt
             omsbgThis = "#9aCD32"
             end if

             if cdbl(hgt) >= 100 then
             hgt = 100           
             bgThis = "#D6dff5"
             else
             hgt = hgt 
             bgThis = "#D6dff5"
             end if

           

            %>
            <td class=lille align=right height=100 valign=bottom style="border-bottom:1px #cccccc solid;">
            
            <%if hgt > 0 then 

                bdThis = 1
                timerThis(dt) = formatnumber(timerThis(dt), 2)
                kostThis(dt) = formatnumber(kostThis(dt) / 1000, 2)
                omsThis(dt) = formatnumber(omsThis(dt) / 1000, 2)
                pdd = 0

             

            else
            timerThis(dt) = 0
            kostThis(dt) = 0
            omsThis(dt) = 0
            bdThis = 0
            pdd = 0
            kostHgt = 0
            end if 
            
            if resHgt <> 0 then
            resThis(dt) = formatnumber(resThis(dt),0)
            else
            resHgt = 0
            end if
            

            if hgt = 0 then
            bgThis = "#ffffff"
            bdThis = 0
            end if

            if kostHgt = 0 then
            kostbgThis = "#ffffff"
            kostbdThis = 0
            end if

            if resHgt = 0 then
            resbgThis = "#ffffff"
            resbdThis = 0
            end if

            if omsHgt = 0 then
            omsbgThis = "#ffffff"
            omsbdThis = 0
            end if

            fz = 7

            'if hgt < 5 then
            'fz = 7
            'pdd = 0
            'end if


            %>
            <div style="height:<%=hgt%>px; border-top:<%=bdThis%>px #CCCCCC solid; width:6px; background-color:<%=bgThis%>;"><img src="../ill/blank.gif" border="0" width="1" height="1" /></div>
            </td>
            <td class=lille align=right height=100 valign=bottom style="border-bottom:1px #cccccc solid;">
            <div style="height:<%=resHgt%>px; border-top:<%=resbdThis%>px #CCCCCC solid; width:6px; background-color:<%=resbgThis%>;"><img src="../ill/blank.gif" border="0" width="1" height="1" /></div>
            </td>
            <td class=lille align=right height=100 valign=bottom style="border-bottom:1px #cccccc solid;">
            <div style="height:<%=omsHgt%>px; border-top:<%=omsbdThis%>px #CCCCCC solid; width:6px; background-color:<%=omsbgThis%>;"><img src="../ill/blank.gif" border="0" width="1" height="1" /></div>
            </td>
             <td class=lille align=right height=100 valign=bottom style="border-right:1px #CCCCCC solid; border-bottom:1px #cccccc solid;">
            <div style="height:<%=kostHgt%>px; border-top:<%=kostbdThis%>px #CCCCCC solid; width:6px; background-color:<%=kostbgThis%>;"><img src="../ill/blank.gif" border="0" width="1" height="1" /></div>
            </td>
            <%next %>
         </tr>
         <tr bgcolor="#FFFFFF">
         <td style="font-size:8px; border-bottom:1px #D6dff5 solid;" align=center>Timer</td>
            <%for dt = 0 TO mts %>
            <td colspan=4 style="border-bottom:1px #D6dff5 solid; font-family:Arial; padding:0px; font-size:7px;" align=center><%=timerThis(dt) %></td>
            
            

            <%next %>
	    </tr>

          <tr bgcolor="#FFFFFF">
           <td style="font-size:8px; border-bottom:1px #999999 solid;" align=center>Forecast</td>
            <%for dt = 0 TO mts %>
            
            <td colspan=4 style="border-bottom:1px #999999 solid; font-family:Arial; padding:0px; font-size:7px;" align=center><%=resThis(dt) %></td>
            

            <%next %>
	    </tr>

          <tr bgcolor="#FFFFFF">
           <td style="font-size:8px; border-bottom:1px #9aCD32 solid;" align=center>Oms.</td>
            <%for dt = 0 TO mts %>
            
            <td colspan=4 style="border-bottom:1px #9aCD32 solid; font-family:Arial; padding:0px; font-size:7px;" align=center><%=omsThis(dt) %> k</td>
            

            <%next %>
	    </tr>
          <tr bgcolor="#FFFFFF">
           <td style="font-size:8px; border-bottom:1px #8caae6 solid;" align=center>Kost.</td>
            <%for dt = 0 TO mts %>
            
            <td colspan=4 style="border-bottom:1px #8caae6 solid; font-family:Arial; padding:0px; font-size:7px;" align=center><%=kostThis(dt) %> k</td>
            

            <%next %>
	    </tr>
        </table>


        </td></tr></table>

      


        </td>
        </tr>
        
        <%
        end if 'synergi 'Risiko




       '***** Detaljeret timeoverblik *******************************************
       if thisfile <> "jobprintoverblik" then
        %>
        <tr>
            <td> <input type="checkbox" id="overfortiljob_tp" name="visrealtimerdetal" value="1" <%=visrealtimerdetalCHK %> /> Vis detaljeret timeforbrug</td>
        </tr>
        <%end if

        '**** Timer Real. detaljer '****************
        if len(trim(visrealtimerdetal)) <> 0 then 'sættes i printjoboverblik
            visrealtimerdetal = visrealtimerdetal
        else
            visrealtimerdetal = 0
        end if


        if cint(visrealtimerdetal) = 1 then
        %>

            <tr>
                <td>
                    <br /><br />
                    <b>Timeforbrug ialt:</b> (detaljeret)


                    <%if thisfile <> "jobprintoverblik" then
                        %>
                     <div style="width:680px; height:200px; overflow:scroll; border:1px #cccccc solid; padding:5px 5px 5px 5px;"><%
                       else
                        %>
                        <div style="width:680px; border:1px #cccccc solid; padding:5px 5px 5px 5px;"><%
                      end if %>

                    <table width="100%" cellpadding="1" cellspacing="0">
                        <tr>
                            
                            <th style="width:200px; text-align:left;">Medarbejder</th>
                            <th style="text-align:right;">Forecast (timebudget)</th>
                            <th style="text-align:right;">Real. Timer</th>
                        </tr>

                            <%strSQlrealtimer = "SELECT mid, init, SUM(t.timer) AS sumtimer, mnavn FROM medarbejdere AS m "_
                            &" LEFT JOIN timer AS t ON (t.tmnr = m.mid AND tjobnr = '"& strJobnr &"' AND ("& aty_sql_realhours &")) "_
                            &" WHERE mid <> 0 GROUP BY mid ORDER BY mnavn"
                             
                            'response.write strSQlrealtimer
                            'response.flush 
                            
                            realtimerDetalTot = 0
                            rt = 0

                         oRec2.open strSQlrealtimer, oConn, 3
                         while not oRec2.EOF




                                '**** Restimer ***'
                                restimerDetal = 0
                                strSQLres = "SELECT sum(timer) AS restimer FROM ressourcer_md AS r WHERE (r.medid = "& oRec2("mid") &" AND r.jobid = "& id &") GROUP BY medid"
                                oRec3.open strSQlres, oConn, 3
                                if not oRec3.EOF then

                                restimerDetal = oRec3("restimer")

                                end if
                                oRec3.close


                                restimerDetalTot = restimerDetalTot + restimerDetal 


                                if IsNull(oRec2("sumtimer")) <> true then
                                realTimer = oRec2("sumtimer")
                                else
                                realTimer = 0
                                end if



                        if cdbl(restimerDetal) <> 0 OR cdbl(realTimer) <> 0 then




                                select case right(rt, 1)
                                case 0,2,4,6,8
                                bgColRt = "#FFFFFF"
                                case else
                                bgColRt = "#D6DFf5"
                                end select


                                'if isNull(oRec2("kommentar")) <> true then
                                'kommThis = oRec2("kommentar")
                                'else
                                'kommThis = ""
                                'end if

                         %>
                        <tr style="background-color:<%=bgColRt%>;">

                            
                            <td style="font-size:10px;"><%=left(oRec2("mnavn"), 20) %>

                                <%if len(trim(oRec2("init"))) <> 0 then %>
                                &nbsp;[ <%=ucase(oRec2("init")) %> ]
                                <%end if %>

                            </td>
                            <td style="text-align:right; font-size:10px;"><%=formatnumber(restimerDetal, 2) %></td>
                            <td style="text-align:right; font-size:10px;"><%=formatnumber(realTimer, 2) %></td>
                        </tr>

                         <%
                             
                          realtimerDetalTot = realtimerDetalTot + realTimer 
                            rt = rt + 1


                        end if


                         oRec2.movenext   
                         wend 
                         oRec2.close
                          %>


                        <tr bgcolor="#FFDFDF">
                            <td>Ialt:</td>
                            <td style="text-align:right"><b><%=formatnumber(restimerDetalTot, 2) %></b></td>
                            <td style="text-align:right"><b><%=formatnumber(realtimerDetalTot, 2) %></b></td>
                        </tr>
                    </table>
                    </div>

                </td>

            </tr>

        <%
        end if



        if thisfile <> "jobprintoverblik" then
        
        
        
        '***forvalgtjobstatus i loblog *'' 
        viskunabnejob0 = "1"
        viskunabnejob1 = "" 
        viskunabnejob2 = ""
        select case strStatus
        case 0
        viskunabnejob0 = "1"
        case 1,2
        viskunabnejob2 = "1"
        case 3
        viskunabnejob1 = "1"
        end select
        
        call timesimon_fn()

        if cint(timesimon) = 1 then
        fcLinkj = "../to_2015/timbudgetsim.asp?FM_fy="&year(now)&"&FM_visrealprdato=1-1-"&year(now)&"&FM_sog="& strJobnr '& "&jobid="& jobid &"&func=forecast"
        else
        fcLinkj = "ressource_belaeg_jbpla.asp?FM_sog="& strJobnr
        end if

        %>

        <tr>
		    <td  style="padding:4px 4px 5px 4px;">
            <a href="joblog.asp?nomenu=1&FM_job=<%=id %>&FM_kunde=<%=strKnr %>&FM_jobsog=<%=strjobnr%>&viskunabnejob0=<%=viskunabnejob0%>&viskunabnejob1=<%=viskunabnejob1%>&viskunabnejob2=<%=viskunabnejob2%>&<%=dtlink %>" class=vmenu target="_blank">Joblog >></a><br />
            <%if level = 1 then %>
            <a href="medarbtyper.asp" class=vmenu target="_blank">Ret kostpriser her >></a><br />
            <%end if %>

            

            <a href="<%=fcLinkj%>" class=vmenu target="_blank">Ressource Forecast >></a>
        </td>
		</tr>
       
        

		
		<%end if 'thisfile%>
		</td>
		</tr>
		

        <%
        if cint(sltjfilok) = 1 then 
        
        if thisfile <> "jobprintoverblik" then
        %>
		<tr>
		    <td><br /><br /><b>Filarkiv</b> 
            <%if thisfile <> "jobprintoverblik" then%>
	        &nbsp;<a href="job_print.asp?id=<%=id%>" class=vmenu>Print / PDF Center >></a> <!--| <a href="javascript:popUp('upload.asp?type=job&id=0&kundeid=<%=strKnr%>&jobid=<%=id%>&nomenu=1','600','500','250','120');" target="_self" class=vmenu><img src="../ill/addmore55.gif" width="10" height="13" alt="" border="0">&nbsp;Upload fil >></a>-->

			<%end if %>
            </td>
		</tr>
		<tr>
		<td style="padding:5px 4px 5px 4px; border:1px #cccccc solid;">
        
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
	
		<tr>
		    <td class=lille width=300><b>Folder \ Fil</b></td>
		    <td class=lille align=right><b>Dato</b></td>
		</tr>
		<%strSQLdok = "SELECT f.filnavn, f.id AS fid, f.dato, fo.navn AS foldernavn FROM filer AS f  "_
        &" LEFT JOIN foldere fo ON (fo.id = f.folderid) WHERE fo.jobid =" & id & " OR f.jobid = "& id &" GROUP BY f.filnavn ORDER BY foldernavn, filnavn" 

        'AND f.filnavn <> ''
        
        'Response.Write strSQLdok
        'Response.flush

        oRec2.open strSQLdok, oConn, 3
        while not oRec2.EOF
        %>
            <tr>
		    <td class=lille><b><%=oRec2("foldernavn") %></b> \ 
		    
            <%if thisfile <> "jobprintoverblik" then%>
            <a href="../inc/upload/<%=lto%>/<%=oRec2("filnavn")%>" class='rmenu' target="_blank"><%=oRec2("filnavn") %></a>
            <%else %>
            <%=oRec2("filnavn") %>
            <%end if %></td>
		    <td class=lille align=right><%=oRec2("dato") %></td>
		</tr>
        <%
        oRec2.movenext
        wend
        oRec2.close
        %>
		
		</table>
		
		<br><br />
		
		</td>
		</tr>
        <%end if 
            
        end if%>


        <%
         if cint(sltjfakok) = 1 then    
            
         if ((thisfile <> "jobprintoverblik" AND lto = "synergi1") OR lto <> "synergi1") AND intprio > -1 then 'interne%>
		<tr>
		    <td><br /><br /><b>Faktura historik</b> </td>
		</tr>
		<tr>
		<td bgcolor="#FFFFFF" style="padding:5px 4px 5px 4px; border:1px #cccccc solid;">
        
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
	    <tr>
	        <td colspan=4>
	        <%if thisfile <> "jobprintoverblik" then 
                
                stDagF = day(strTdato)
                stMrdF = month(strTdato)
                stAarF = year(strTdato)

                slDagF = day(now)
                slMrdF = month(now)
                slAarF = year(now)
                
            if cint(useasfak) <= 2 then%>
	        <a href="erp_opr_faktura_fs.asp?func=opr&visfaktura=1&visjobogaftaler=1&visminihistorik=1&FM_kunde=<%=strKnr%>&FM_job=<%=id%>&FM_aftale=0&FM_start_dag=<%=stDagF%>&FM_start_mrd=<%=stMrdF%>&FM_start_aar=<%=stAarF%>&FM_slut_dag=<%=slDagF%>&FM_slut_mrd=<%=slMrdF%>&FM_slut_aar=<%=slAarF%>" class=vmenu target="_blank"><span style="color:green; font-size:16px;"><b>+</b></span> Opret faktura</a>
		    <%end if %>
           <%end if %>
		</td>
		</tr>
		<tr>
		    <td class=lille><b>Faktura nr.</b></td>
		    <td class=lille><b>Dato</b></td>
		    <td class=lille align=right><b>Beløb</b></td>
		    <td class=lille align=right><b>Beløb eksl. <br />materialer og km.</b></td>
		</tr>
		
		<%
		strSQLFak = "SELECT f.jobid, f.aftaleid, f.fid, f.fakdato, f.faknr, f.betalt, "_
		&" f.faktype, f.beloeb, f.shadowcopy, f.valuta, f.kurs, v.valutakode, SUM(fd.aktpris) AS aktbel, f.medregnikkeioms, f.brugfakdatolabel, f.labeldato, f.fakadr "_
		&" FROM fakturaer f LEFT JOIN valutaer v ON (v.id = f.valuta) "_
		&" LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid AND fd.enhedsang <> 3)"_
		&" WHERE f.jobid = "& id & " AND f.aftaleid = 0 AND f.shadowcopy = 0"_ 
		&" GROUP BY f.fid ORDER BY f.faknr DESC"
		
		'Response.Write strSQLfak
		'Response.Flush
		
		totFakbel = 0
		totFakbelKunTimer = 0
	    oRec2.open strSQLFak, oConn, 3
	    while not oRec2.EOF 
	    %>
	    <tr>
	        <td class=lille>
            <%if thisfile <> "jobprintoverblik" then %>
                    <a href="erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id=<%=oRec2("fid")%>&FM_jobonoff=<%=FM_jobonoffval%>&FM_kunde=<%=oRec2("fakadr")%>&FM_job=<%=oRec2("jobid")%>&FM_aftale=<%=oRec2("aftaleid")%>&fromfakhist=1" class="lgron" target="_blank"><%=oRec2("faknr")%></a>
          
            <%else %>
            <%=oRec2("faknr") %>
            <%end if %>
            
            <%if cint(oRec2("medregnikkeioms")) = 1 then %>
            (intern)
            <%end if %>

               <%if cint(oRec2("medregnikkeioms")) = 2 then %>
            (handelsfak.)
            <%end if %>
            </td>
	        <td class=lille>
            
                 <%if cint(oRec2("brugfakdatolabel")) = 1 then %>
        L: <b><%=replace(formatdatetime(oRec2("labeldato"),2),"-",".")  %></b>&nbsp;
        <span style="font-size:9px; color:#999999;">(<%=replace(formatdatetime(oRec2("fakdato"),2),"-",".") %>)</span>
        <%else %>
        F: <b><%=replace(formatdatetime(oRec2("fakdato"),2),"-",".") %></b>
        <%end if %>

            

            </td>
	        <%
	        call beregnValuta(minus&(oRec2("beloeb")),oRec2("kurs"),100)
            if oRec2("faktype") <> 1 then
            belobGrundVal = valBelobBeregnet
            else
            belobGrundVal = -valBelobBeregnet
            end if %>
	        
	        
	        <td class=lille align=right><%=formatnumber(belobGrundVal) &" "& basisValISO %></td>
	        
	        <%
	        '** Kun aktiviteter timer, enh. stk. IKKE mateiler og KM
                
                  if cDate(oRec2("fakdato")) < cDate("01-06-2010") AND (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati") then
                        belobKunTimerStk = belobGrundVal
                        else
       

	            call beregnValuta(minus&(oRec2("aktbel")),oRec2("kurs"),100)
                if oRec2("faktype") <> 1 then
                belobKunTimerStk = valBelobBeregnet
                else
                belobKunTimerStk = -valBelobBeregnet
                end if

                end if
            
            %>
	        
	        <td class=lille align=right><%=formatnumber(belobKunTimerStk) &" "& basisValISO %></td>
	    
	    </tr>
	    <%

        if cint(oRec2("medregnikkeioms")) = 0 then ' else: 1: intern/ 2: handels. fak.
	    totFakbel = totFakbel + (belobGrundVal)
	    totFakbelKunTimer = totFakbelKunTimer + (belobKunTimerStk)
        else
        totFakbel = totFakbel 
	    totFakbelKunTimer = totFakbelKunTimer
        end if
	    oRec2.movenext
	    wend
	    oRec2.close
		%>
		<tr bgcolor="#DBDB70">
            <td class=lille><b>Ialt:</b></td>
		    <td class=lille align=right colspan=2><b><%=formatnumber(totFakbel)%></b> <%=basisValISO %></td>
		    <td class=lille align=right><b><%=formatnumber(totFakbelKunTimer)%></b> <%=basisValISO %></td>
		</tr>
		</table>
		
		 
		
		</td>
		</tr>
        <%end if 'synergi 
            
        end if 'cint(sltjjobok) = 1 then %>


		<tr>
		    <td><br /><br /><b>Aktiviteter og faser:</b> (budget og timeforbrug)
            
            <%
            if cint(sltjjobok) = 1 then 
                if thisfile <> "jobprintoverblik" then %>
                &nbsp;<a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&jobid=<%=id%>&jobnavn=<%=strNavn%>&rdir=<%=rdir%>&nomenu=1', '', 'width=1004,height=800,resizable=yes,scrollbars=yes')" class=vmenu>Rediger aktivitetliste udviddet >></a>
                <%end if
            end if%>

            </td>
		</tr>
		<tr>
		<td bgcolor="#FFFFFF" style="padding:5px 4px 5px 4px; border:1px #cccccc solid;"">
		
		
	     <table cellpadding=0 cellspacing=0 border=0 width=100%>

             <%  if cint(sltjjobok) = 1 then
                 if thisfile <> "jobprintoverblik" then  %>
	     <tr>
	     <td>
	     <%
        
         
                            oleftpx = 500
	                        otoppx = 10
	                        owdtpx = 140
	                        java = "Javascript:window.open('aktiv.asp?menu=job&func=opret&jobid="&id&"&id=0&jobnavn="&strNavn&"&fb=1&rdir=job3&nomenu=1', '', 'width=1004,height=800,resizable=yes,scrollbars=yes')"    
	                        call opretNyJava("#", "Opret ny aktivitet", otoppx, oleftpx, owdtpx, java) 
         
                %>
         
         
	    <br />&nbsp;
	     
	     </td>
	     </tr>
             <% end if
            end if %>

	     <tr>
	        <td bgcolor="#FFFFFF">
            <%

            '**************************************************************
            '************** Aktiviteter budget og Real ********************
            '**************************************************************

            if thisfile <> "jobprintoverblik" then
            %>
            <div id="akt_faser" style="position:relative; overflow:auto; height:400px;">
            <%
            else
             %>
            <div id="akt_faser" style="position:relative;">
            <%
            end if
            %>
            

	        <table cellpadding=0 cellspacing=0 border=0 width=100%>
	        <tr bgcolor="#5582d2">
	        <td class=lille><b>Navn</b></td>
            <td class=lille><b>Periode</b></td>
	        <td class=lille><b>Status</b></td>
            <td class=lille><b>Type</b></td>
	        <td class=lille align=right><b>Forkalk.<br />Timer / Stk.</b></td>
            <td class=lille align=right><b>Time / Stk. pris</b></td>
            <td class=lille align=right><b>Budget</b></td>
            <td class=lille align=right><b>Real. tim.</b></td>
            <td class=lille align=right><b>Real. oms.</b> <br />(timer *<br> medarb. timepris)</td>
	        </tr>
            
            
	        
	        <%strSQLakt = "SELECT a.id, a.navn, job, beskrivelse, fakturerbar, aktstartdato, "_
	        &" aktslutdato, budgettimer, aktstatus, aktbudget, fomr.navn AS fomr, aty_id, "_
	        &" antalstk, tidslaas, a.fase, a.sortorder, a.bgr, aktbudgetsum, "_
            &" COALESCE(sum(t.timer),0) AS realiseret, COALESCE(SUM(timer * timepris *(kurs/100)),0) AS realbelob, COALESCE(SUM(timer * kostpris *(kurs/100)), 0) AS realtimerkost, aktstartdato, aktslutdato, aty_desc "_
	        &" FROM aktiviteter a LEFT JOIN fomr ON (fomr.id = a.fomr) "_
	        &" LEFT JOIN timer AS t ON (t.taktivitetid = a.id)"_
            &" LEFT JOIN akt_typer aty ON (aty_id = a.fakturerbar) "_
	        &" WHERE job = "& id &" AND a.aktfavorit = 0 GROUP BY a.id ORDER BY a.fase, a.sortorder, a.navn" 
        	
        	totSum = 0
	        totTimerforkalk = 0
	        totReal = 0
            realtimerkost = 0
	        'Response.Write strSQLakt
	        'Response.flush
	        lastFaseSum = 0
            lastFase = ""
            a = 0
            lastFaseRealTimer = 0
            lastFaseForkalkTimer = 0
	        oRec2.open strSQLakt, oConn, 3
	        while not oRec2.EOF  
	        
	            select case right(a, 1)
	            case 0,2,4,6,8
	            bga = "#FFFFFF"
                case else
	            bga = "#Eff3ff"
	            end select
	        
	            if lcase(lastFase) <> lcase(oRec2("fase")) AND len(trim(oRec2("fase"))) <> 0 then %>

                <%if a <> 0 then %>
               <tr bgcolor="#FFFFFF">
                
                <td>&nbsp;</td>
	            <td>&nbsp;</td>
	            <td>&nbsp;</td>
                <td>&nbsp;</td>
                
	            
                <td align=right class=lille><b><%=formatnumber(lastFaseForkalkTimer,2) %></b> t.</td>
                <td>&nbsp;</td>
	            <td align=right class=lille><b><%=formatnumber(lastFaseSum, 2) & "</b> "& basisValISO_f8 %></td>

                <td align=right class=lille><b><%=formatnumber(lastFaseRealTimer,2) %> t.</td>
	            <td align=right class=lille><b><%=formatnumber(lastFaseRealbel, 2) & "</b> "& basisValISO_f8 %></td>

                </tr>

              <tr>
                <td colspan=9 style="padding:2px; font-size:9px;">&nbsp;</td>

              </tr>
                <%
                lastFaseForkalkTimer = 0
                lastFaseRealTimer = 0
                lastFaseRealbel = 0
                end if %>
	            
	          
	            <%lastFaseSum = 0 %>

             

	            <tr bgcolor="#8cAAe6"><td colspan=9 style="padding:2px; font-size:9px;"><b><%=replace(oRec2("fase"), "_", " ")%></b></td></tr>
	            <%end if %>
	            
	            
	            <tr bgcolor="<%=bga %>">
	            
                <td class=lille>
                <%if thisfile <> "jobprintoverblik" AND cint(sltjjobok) = 1 then  %>
                <a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&func=red&id=<%=oRec2("id")%>&jobid=<%=id %>&jobnavn=<%=strNavn%>&rdir=job2&nomenu=1', '', 'width=850,height=600,resizable=yes,scrollbars=yes')" class="rmenu"><%=left(oRec2("navn"), 20) %></a>
                <%else %>
                <%=left(oRec2("navn"), 20) %>
                <%end if %>
                </td>
	            
                
                <td class=lillegray><%=formatdatetime(oRec2("aktstartdato"), 2)  &" - "& formatdatetime(oRec2("aktslutdato"), 2) %></td>
                <td class=lille>
	            <%select case oRec2("aktstatus")
	            case 1
	            aktstat = "Aktiv"
	            case 2
	            aktstat = "Passiv"
	            case else
	            aktstat = "Lukket"
	            end select
	             %>
	            <%=aktstat %></td>
                
                       <%call akttyper(oRec2("aty_id"), 1) %>

                 <td class=lille><%=left(akttypenavn, 10) %>.</td>
                 <%select case oRec2("bgr")
                 case 0
                 bgr = "ingen"
                 %>
                 <td class=lille align=right>-</td>
                 <%
                 case 1
                 bgr = "timer"
                 %>
                 <td class=lille align=right><%=formatnumber(oRec2("budgettimer"), 2) %> t.</td>
                 <%
                 case 2
                 bgr = "stk."
                 %>
                 <td class=lille align=right><%=formatnumber(oRec2("antalstk"), 2) %> stk.</td>
                 <%
                 end select %>

	            
	            <td class=lille align=right><%=formatnumber(oRec2("aktbudget"), 2) %></td>
	            <td class=lille align=right><%=formatnumber(oRec2("aktbudgetsum"), 2) &" "& basisValISO_f8 %></td>

                <td class=lille align=right><%=formatnumber(oRec2("realiseret"), 2) %> t.</td>
	            <td class=lille align=right><%=formatnumber(oRec2("realbelob"), 2) &" "& basisValISO_f8 %></td>
	            
	            </tr>
	        
	        <%
	        a = a + 1
	        lastFaseSum = lastFaseSum + oRec2("aktbudgetsum") 

            if len(trim(oRec2("fase"))) <> 0 then
	        lastFase = oRec2("fase")
            else
            lastFase = ""
            end if 
	        
           
           
            

	        call akttyper2009Prop(oRec2("fakturerbar"))
	        if cint(aty_real) = 1 OR (lto = "oko" AND oRec2("aty_id") = 90) then

                select case lto
                case "oko"

                    lastFaseRealTimer = lastFaseRealTimer + oRec2("realiseret")
                    lastFaseRealbel = lastFaseRealbel + oRec2("realbelob")

                    if oRec2("aty_id") = 90 then 'KUN NAV E 1 linjker

                        totSum = totSum + oRec2("aktbudgetsum")
	                    totTimerforkalk = totTimerforkalk + oRec2("budgettimer")

                        totReal = totReal
                        totRealbel = totRealbel
                        realtimerkost = realtimerkost
                      

                    else
                    
                        totSum = totSum
	                    totTimerforkalk = totTimerforkalk 
                    
                        totReal = totReal + oRec2("realiseret")
                        totRealbel = totRealbel + oRec2("realbelob") 
                        realtimerkost = realtimerkost + oRec2("realtimerkost")
                    

                    end if

                case else

                    totSum = totSum + oRec2("aktbudgetsum")
	                totTimerforkalk = totTimerforkalk + oRec2("budgettimer")
                    totReal = totReal + oRec2("realiseret")
                    totRealbel = totRealbel + oRec2("realbelob") 
                    lastFaseRealTimer = lastFaseRealTimer + oRec2("realiseret")
                    lastFaseRealbel = lastFaseRealbel + oRec2("realbelob")
                    realtimerkost = realtimerkost + oRec2("realtimerkost")
                
                end select

            

	        else

            totSum = totSum
	        totTimerforkalk = totTimerforkalk 
            totReal = totReal
            totRealbel = totRealbel
            lastFaseRealTimer = lastFaseRealTimer
            lastFaseRealbel = lastFaseRealbel 
            realtimerkost = realtimerkost
	        end if

            lastFaseForkalkTimer = lastFaseForkalkTimer + oRec2("budgettimer")
            

	        oRec2.movenext
	        wend
	        oRec2.close%>
	        
	        <tr bgcolor="#FFFFFF">
                
                 <td>&nbsp;</td>
	            <td>&nbsp;</td>
	            <td>&nbsp;</td>
                <td>&nbsp;</td>
                
	            <td align=right class=lille><b><%=formatnumber(lastFaseForkalkTimer,2) %></b> t.</td>
                <td>&nbsp;</td>
	            <td align=right class=lille><b><%=formatnumber(lastFaseSum, 2) %></b> <%=basisValISO_f8 %></td>

                <td align=right class=lille><b><%=formatnumber(lastFaseRealTimer,2) %></b> t.</td>
	            <td align=right class=lille><b><%=formatnumber(lastFaseRealbel, 2) %></b> <%=basisValISO_f8%></td>

            </tr>
            


                <%

                  select case lto
                   case "oko"
                    aktGtBgcol = "#FFDFDF"
                    'aktGtwdh = 300
                    'aktGtwdh2 = 80
                    case else 
                    aktGtBgcol = "#FFDFDF"
                    'aktGtwdh = 340
                    'aktGtwdh2 = 80
                       
                 end select 

              
                strAktiviteterGT = "<tr><td bgcolor=""#FFFFFF"" colspan=9 style=""padding:2px; font-size:9px;"">&nbsp;</td></tr>"
                strAktiviteterGT = strAktiviteterGT &"<tr bgcolor="& aktGtBgcol &"><td class=lille width="& aktGtwdh &"><b>Ialt:</b></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
                strAktiviteterGT = strAktiviteterGT &"<td align=right class=lille><b>"& formatnumber(totTimerforkalk,2) &" t.</b></td>"
                strAktiviteterGT = strAktiviteterGT &"<td width="& aktGtwdh2 &">&nbsp;</td>"
                strAktiviteterGT = strAktiviteterGT &"<td align=right class=lille><b>"& formatnumber(totSum, 2) &"</b> "& basisValISO_f8 &"</td>"
                strAktiviteterGT = strAktiviteterGT &"<td align=right class=lille><b>"& formatnumber(totReal,2) &" t.</td>"
                strAktiviteterGT = strAktiviteterGT &"<td align=right class=lille><b>"& formatnumber(totRealbel, 2) &"</b> "& basisValISO_f8 &"</td>"
                strAktiviteterGT = strAktiviteterGT &"</tr>"



                select case lto
                   case "xoko"
                    
                    case else    
                %>
               
                <%=strAktiviteterGT %>
              
             
                <%end select %>

                </table>
                </div><!-- akt_faser -->
           

             </td>
	     </tr>
         </table>
            
 
        <%'**** BUDGET akkumuleret ****'  
              
        if ((thisfile <> "jobprintoverblik" AND lto = "synergi1") OR lto <> "synergi1") then%>
        <tr>
		    <td style="padding:5px 4px 5px 0px;"><br /><br /><b>Salgsomkostninger budget/forbrugt akkumuleret</b></td>
		</tr>
        <tr>
        <td style="padding:4px 4px 5px 4px; border:1px #cccccc solid;">

            <table cellpadding=0 cellspacing=0 border=0 width=100%>
            <tr bgcolor="#FFCC66">
                <td class=lille><b>Navn</b></td>

                 <% select case lto
                case "oko", "intranet - local"
                case else %>
                <td class=lille align=right><b>Antal</b></td>
                <td class=lille align=right><b>Stk. pris</b></td>
                <td class=lille align=right><b>Købspris</b></td>
                <td class=lille align=right><b>Faktor</b></td>
                <%end select %>

                <td class=lille align=right><b>Budget Salgspris</b></td>

                <% select case lto
                case "oko", "intranet - local" %>
                <td class=lille align=right><b>Realiseret</b></td>
                <%end select %>
            </tr>
            <% 

                select case lto
                case "xoko", "xintranet - local" %>
              
                <%=strAktiviteterGT %>

                <%end select %>

                <%
                
                salgspris_tot = 0
                strSQLUlev = "SELECT ju_id, ju_navn, ju_ipris, ju_faktor, "_
		        &" ju_fase, ju_stk, ju_stkpris, "
                
                select case lto
                case "oko", "intranet - local"

                'strSQLUlev = strSQLUlev &", ju_konto, kp.navn AS kontonavn, kp.kontonr, COALESCE(SUM(matantal * matsalgspris), 0) AS realbelob FROM job_ulev_ju "
                'strSQLUlev = strSQLUlev &" LEFT JOIN kontoplan AS kp ON (kp.id = ju_konto)"
                'strSQLUlev = strSQLUlev &" LEFT JOIN materiale_forbrug AS mf ON (mf.mf_konto = ju_konto AND mf.jobid = "& id &")"
                'strSQLUlev = strSQLUlev &" WHERE ju_jobid = "& id & " GROUP BY ju_konto ORDER BY kontonr"

                strSQLUlev = strSQLUlev &" SUM(ju_belob) AS ju_belob, ju_konto, kp.navn AS kontonavn, kp.kontonr, kp.id AS kontoid "
                strSQLUlev = strSQLUlev &" FROM kontoplan AS kp "
                strSQLUlev = strSQLUlev &" LEFT JOIN job_ulev_ju ON (ju_konto = kp.id AND ju_jobid = "& id & ") "
                strSQLUlev = strSQLUlev &" WHERE kp.id <> 0 AND kontonr >= 200 GROUP BY kontonr ORDER BY kontonr" 
                'AND ju_navn IS NOT NULL

                case else

                strSQLUlev = strSQLUlev &" ju_belob FROM job_ulev_ju WHERE ju_jobid = "& id & " ORDER BY ju_navn"

                end select

                'if session("mid") = 1 then
                '    response.write strSQLUlev
                '    response.flush
                'end if

                u = 0
                salgreal_tot = 0
		        
                oRec2.open strSQLUlev, oConn, 3 
                while not oRec2.EOF 

                select case right(u, 1)
                case 0,2,4,6,8
                bgthis = "#FFFFFF"
                case else
                bgthis = "#FFFFe1"
                end select


                    select case lto
                    case "oko", "xintranet - local"
                    juNavn = oRec2("kontonr") & " " & oRec2("kontonavn") 
                        
                        if isNull(juNavn) <> true then
                        'juNavn = juNavn &" ("& oRec2("ju_navn") &")"
                               '*** SKAL VÆRE DEN FØRSTE POSTERING PÅ HVER KONTO
                               
                               strSQLunavn = "SELECT ju_navn AS forstepostering FROM job_ulev_ju WHERE ju_konto = "& oRec2("kontoid") &" AND ju_jobid = "& id & " ORDER BY ju_id LIMIT 1"
                               oRec3.open strSQLunavn, oConn, 3 
                               if not oRec3.EOF then

                               juNavn = juNavn &" ("& oRec3("forstepostering") &")"

                               end if
                               oRec3.close


                        end if

                    case else
                    juNavn = oRec2("ju_navn")
                    end select



                     select case lto
                        case "oko", "intranet - local" 
                            
                               realBelob = 0
                               strSQLUmf = "SELECT COALESCE(SUM(matantal * matsalgspris), 0) AS realbelob FROM materiale_forbrug AS mf WHERE mf.mf_konto = "& oRec2("kontoid") &" AND mf.jobid = "& id & ""
                               oRec3.open strSQLUmf, oConn, 3 
                               if not oRec3.EOF then


                                realBelob = oRec3("realbelob")

                               end if
                               oRec3.close

                               if len(trim(realBelob)) <> 0 then
                               realBelob = realBelob
                               else
                               realBelob = 0
                               end if

                     end select

                'if session("mid") = 1 AND oRec2("kontonr") = "205" then
                '  response.write "Konto:" & oRec2("kontonr") &" <br>"& strSQLUmf & "<br>realBelob: " & realBelob
                '  response.flush
                'end if

                 if oRec2("ju_stk") <> 0 OR cdbl(realBelob) <> 0 then          
                 %>

                
                <tr bgcolor="<%=bgthis %>">
                    <td class=lille><%=juNavn%></td>


                    <% select case lto
                    case "oko", "intranet - local"
                    case else %>
                    <td align=right class=lille ><%=oRec2("ju_stk") %></td>
                    <td align=right class=lille><%=formatnumber(oRec2("ju_stkpris"), 2) &" "& basisValISO_f8 %></td>
                    <td align=right class=lille><%=formatnumber(oRec2("ju_ipris"), 2) &" "& basisValISO_f8%> </td>
                    <td align=right class=lille><%=oRec2("ju_faktor") %></td>
                    <%end select %>

                    <%if isNull(oRec2("ju_stk")) <> true then %>
                    <td align=right class=lille><%=formatnumber(oRec2("ju_belob"), 2) &" "& basisValISO_f8%> </td>
                    <%else %>
                    <td align=right class=lille>0,00 <%=basisValISO_f8 %></td>
                    <%end if%>

                        <%select case lto
                        case "oko", "intranet - local"  %> 
                        <td align=right class=lille><%=formatnumber(realBelob, 2) &" "& basisValISO_f8%> </td>
                        <%end select %>
                    
                    
                </tr>
                <%

                if isNull(oRec2("ju_stk")) <> true then
                salgsomkostninger_tot = salgsomkostninger_tot + oRec2("ju_ipris") 
                salgspris_tot = salgspris_tot + oRec2("ju_belob")
                else
                salgsomkostninger_tot = salgsomkostninger_tot 
                salgspris_tot = salgspris_tot
                end if

                        select case lto
                        case "oko", "xintranet - local"
                        salgreal_tot = salgreal_tot + realBelob 'oRec2("realbelob")
                        end select


                end if' ju:stk > 0          


                u = u + 1
                oRec2.movenext
                wend
                oRec2.close 


                

                
                if u <> 0 then%>


                 <%select case lto
                case "oko"

                    %>

                 <tr bgcolor="#cccccc">
                    <td class=lille><b>Salgsomkostninger:</b></td>

                <%case else %>

                        <tr bgcolor="#FFDFDF">
                    <td class=lille><b>Salgsomkostninger Ialt:</b></td>

                <%end select %>


                      <% select case lto
                    case "oko", "intranet - local"
                    case else %>
                    <td align=right class=lille>&nbsp;</td>
                    <td align=right class=lille>&nbsp;</td>
                    <td align=right class=lille><b><%=formatnumber(salgsomkostninger_tot, 2) &" "& basisValISO_f8 %></b></td>
                    <td align=right class=lille>&nbsp;</td>
                    <%end select %>

                    <td align=right class=lille><b><%=formatnumber(salgspris_tot, 2) &" "& basisValISO_f8 %></b></td>

                      <% select case lto
                        case "oko", "intranet - local" %>
                        <td align=right class=lille><%=formatnumber(salgreal_tot, 2) &" "& basisValISO_f8%> </td>
                        <%end select %>
                     <td></td>
                    
                </tr>


                <%select case lto
                case "oko"

                    %>
                  <tr bgcolor="#999999">
                    <td class=lille><b>Lønomkostninger Ialt:</b> (100-190)</td>
                          <td align=right class=lille><b><%=formatnumber(totSum, 2) &" "& basisValISO_f8 %></b></td>
                      <td align=right class=lille><%=formatnumber(totRealbel, 2) &" "& basisValISO_f8%> </td>

                  

                 </tr>

                 <tr bgcolor="#FFDFDF">
                    <td class=lille><b>Omkostninger Ialt:</b></td>
                          <td align=right class=lille><b><%=formatnumber(totSum+salgspris_tot, 2) &" "& basisValISO_f8 %></b></td>
                      <td align=right class=lille><%=formatnumber(totRealbel+salgreal_tot, 2) &" "& basisValISO_f8%> </td>

                  

                 </tr>

                 <%
                 end select %>



                <%end if 


             
                 %>

            </table>


        </td>
        </tr>
         <%end if %>


            <tr>
		    <td bgcolor="#FFFFFF" style="padding:5px 4px 5px 0px;"><br /><br /><b>Salgsomkostninger udspecificeret</b> (materialeforbrug & udlæg) &nbsp; 
            <%if thisfile <> "jobprintoverblik" AND cint(sltjjobok) = 1 then %>
            <a href="materiale_stat.asp?FM_job=<%=id %>&FM_kunde=<%=strKnr %>&<%=dtlink %>&nomenu=1" class=vmenu target="_blank">Se materialer / udlæg's stat. >></a>
            <%end if %>
           
               </td>
		</tr>
        <tr>
        <td bgcolor="#FFFFFF" style="padding:4px 4px 5px 4px; border:1px #cccccc solid;">
        <%if thisfile <> "jobprintoverblik" AND cint(sltjjobok) = 1 then %>
                 <a href="materialer_indtast.asp?id=<%=id%>&fromsdsk=0&aftid=<%=intServiceaft%>" target="_blank" class="vmenu"><span style="color:green; font-size:16px;"><b>+</b></span> Opret salgsomkostning</a><br /><br />&nbsp;
        <%end if
        %>
		     
                
             <%
            if thisfile <> "jobprintoverblik" then
            %>
            <div style="height:400px; overflow:auto;">
            <%
            else
             %>
            <div>
            <%
            end if
            %>

                 
		        <table cellpadding=1 cellspacing=0 border=0 width=100%>
                <tr bgcolor="#FFCC66">
                    <td class=lille><b>Navn</b></td>
                    <td class=lille align=right><b>Antal stk.</b></td>
                    <td class=lille><b>Enhed</b></td>
                    <td class=lille><b>Dato</b></td>
                    <td class=lille align=right><b>Stk. pris</b></td>
                    <td class=lille align=right><b>Købspris</b></td>
                    <td class=lille align=right><b>Salgspris</b></td>
                    
                 </tr>
                
                 
                  
	            <tr>
	        
                    <%

                    select case lto
                    case "oko", "xintranet - local"
                        strMatforbrugSQLOrderBy = " k.kontonr, forbrugsdato DESC"
                        strKontonrKri = " AND (mf_konto > 14)" '** Ikke løn
                    case else
                        strMatforbrugSQLOrderBy = " matgrp, forbrugsdato DESC"
                        strKontonrKri = ""
                    end select


                    m = 0
                    antalmatreg = 0
                    salgsomkostSalg = 0
                    salgsomkostKost = 0
                    strSQLudl = "SELECT matnavn, forbrugsdato, matenhed, matantal AS antal, (matkobspris * (kurs/100)) AS stkkobspris, matkobspris * matantal * (kurs/100) AS matkobspris, "_
                     &" matsalgspris * matantal * (kurs/100) AS matsalgspris, mf_konto, mg.navn AS matgruppenavn, k.kontonr, k.navn AS kontonavn, matgrp FROM materiale_forbrug AS mf "_
                     &" LEFT JOIN materiale_grp AS mg ON (mg.id = matgrp) "_
                     &" LEFT JOIN kontoplan AS k ON (k.id = mf_konto) WHERE jobid = "& id & " "& strKontonrKri &" GROUP BY mf.id ORDER BY  "& strMatforbrugSQLOrderBy &"" 
                        
                    'Response.Write strSQLudl
                    'Response.flush
                    'salgsomkostReal = 0
                    matforbrugGrpSubTot = 0
                    lastGrpNavn = ""
                    lastGrp = 0
                    
                    oRec2.open strSQLudl, oConn, 3
	                while not oRec2.EOF 

                    select case right(m,1)
                    case 0,2,4,6,8
                    matBg = "#FFFFFF"
                    case else
                    matBg = "#FFFFe1"
                    end select

                    'call matforbrugGrpSubTot_fn


                    select case lto
                    case "oko", "xintranet - local"
                        thisGrp = oRec2("mf_konto")
                      
                   case else
                        thisGrp = oRec2("matgrp")
                       
                   end select


                    if lastGrp <> 0 AND lastGrp <> "" AND lastGrp <> thisGrp AND m <> 0 then
                    %>

                     <tr style="background-color:#cccccc;">
                        <td colspan="6" class="lille"><b><%=lastGrpNavn%> ialt:</b></td>
                        <td align="right" class="lille"><b><%=formatnumber(matforbrugGrpSubTot, 2) &" "& basisValISO_f8 %></b></td>
                     </tr>    

                    <%matforbrugGrpSubTot = 0
                      lastGrpNavn = ""
                     end if 
                        
                        
                        
                   %>


                    <tr bgcolor="<%=matBg %>">
                    <td class=lille><%=oRec2("matnavn") %></td>
                    <td align=right class=lille><%=oRec2("antal") %></td>
                    <td class=lille><%=oRec2("matenhed")%></td>
                    <td class=lille><%=oRec2("forbrugsdato")%></td>
                    <td class=lille align=right><%=formatnumber(oRec2("stkkobspris"), 2)%></td>
                    <td align=right class=lille><%=formatnumber(oRec2("matkobspris"), 2) &" "& basisValISO_f8 %></td>
                    <td align=right class=lille><%=formatnumber(oRec2("matsalgspris"), 2) &" "& basisValISO_f8 %></td>
                   
                    </tr>
                    <%
                         


                    
                    select case lto
                    case "oko", "xintranet - local"
                        lastGrp = oRec2("mf_konto")
                        lastGrpNavn = oRec2("kontonr") &" "& oRec2("kontonavn") 
                    case else
                        lastGrp = oRec2("matgrp")
                        lastGrpNavn = oRec2("matgruppenavn") 'oRec2("kontonavn") & " " & oRec2("kontonr")
                    end select

                    m = m + 1
                    antalmatreg = antalmatreg + 1
                    salgsomkostSalg = salgsomkostSalg + oRec2("matsalgspris")
                    matforbrugGrpSubTot = matforbrugGrpSubTot + oRec2("matsalgspris")
                    salgsomkostKost = salgsomkostKost + oRec2("matkobspris")

                   

                    oRec2.movenext
                    wend 
                    oRec2.close 
                    
                    %>
                       <tr style="background-color:#cccccc;">
                        <td colspan="6" class="lille"><b><%=lastGrpNavn%>:</b></td>
                         <td align="right" class="lille"><b><%=formatnumber(matforbrugGrpSubTot, 2) &" "& basisValISO_f8 %></b></td>
                
                    </tr>    
                    <%


                    'salgsomkostReal = salgsomkostKost

                 
                    %>
                    </table>
                </div>
                     <table cellpadding=1 cellspacing=0 border=0 width=100%>
                    <tr bgcolor="#FFDFDF"><td class=lille colspan=4><b>Ialt: <%=antalmatreg%></b><img src="ill/blank.gif" width="200" height="1" border="0" /></td>
                    <td class=lille align=right><b><%=formatnumber(salgsomkostKost, 2)%></b> <%=basisValISO_f8 %></td>
                    <td class=lille align=right><b><%=formatnumber(salgsomkostSalg, 2)%> </b><%=basisValISO_f8%></td></tr>
                   </table>
            
            </td>
	     </tr>





        <%if (thisfile <> "jobprintoverblik" AND lto = "synergi1") OR (lto <> "synergi1" AND lto <> "oko" AND lto <> "intranet - local") then%>
        <tr><td>



          <%if totReal <> 0 then 
	        gnstpris = totFakbelKunTimer/totReal
	        else 
	        gnstpris = 0
	        end if%>
            <table cellpadding=0 cellspacing=0 border=0 width=100%>
	        <tr>
	            <td colspan=7>
                <br /><br /><h4>Nøgletal</h4>
                <table cellpadding=2 cellspacing=0 border=0 width=90%>
               
                <tr><td><b>Restestimat</b></td><td align=right>
                <%select case stade_tim_proc 
                case 0
                stade_tim_proc_txt = " timer til rest."
                case 1
                stade_tim_proc_txt = " % afsluttet"
                end select %>
                <%=restestimat &" "& stade_tim_proc_txt %></td></tr>
                <tr><td> <b>Timer budgetteret:</b> (job)</td><td align=right> <%=formatnumber(strBudgettimer, 2) %> t.</td><td>&nbsp;</td></tr>
                <tr><td> <b>Timer realiseret:</b> (job)</td><td align=right> <%=formatnumber(totReal, 2) %> t.</td><td>&nbsp;</td></tr>
                 <tr><td colspan=3>&nbsp;</td></tr>
                
                     <tr><td style="border-bottom:1px #999999 solid;"><b>Budget</b></td><td align=right style="border-bottom:1px #999999 solid;">Kost.</td><td align=right style="border-bottom:1px #999999 solid;">Salg</td></tr>
                <tr><td> <b>Bruttoomsætning:</b> (job)</td><td>&nbsp;</td><td align=right> <%=formatnumber(jo_bruttooms, 2) &" "&basisValISO %></td></tr>
                <tr><td> <b>Nettoomsætning:</b> (job)</td><td align=right> <%=formatnumber(jo_udgifter_intern, 2) &" "&basisValISO%></td><td align="right"><%=formatnumber(jo_gnsbelob, 2) &" "&basisValISO%></td></tr>
                <tr><td><span style="color:#999999;"><b>Omsætning:</b> (aktiviteter)</span></td><td>&nbsp;</td><td align=right><span style="color:#999999;"><%=formatnumber(totsum, 2) &" "&basisValISO%></span></td></tr>
                  <tr><td><b>Salgsomkostninger:</b> (job)</td><td align="right"><%=formatnumber(salgsomkostninger_tot) &" "&basisValISO %></td><td align=right><%=formatnumber(salgspris_tot, 2) %></b> <%=basisValISO %></td></tr>
              
                 
               <tr><td><b>Dækningsbidrag / Bruttofortjeneste:</b> (budgetteret)</td><td align=right><%=formatnumber(jo_bruttofortj)& " "& basisValISO%></td><td>&nbsp;</td></tr>
               <tr><td><b>DB:</b></td><td align=right><%=formatnumber(jo_dbproc,0)%> %</td><td>&nbsp;</td></tr>
              
                    
                     <%'dbreal
                    realOmsialt = (totRealbel+salgsomkostSalg)
                    realKostialt = (realtimerkost+salgsomkostKost)
                    
                    dbreal = (totRealbel - realtimerkost) + (salgsomkostSalg - salgsomkostKost)
                    
                    call fn_dbproc(realOmsialt,dbreal)
                    dbrealpro = dbProc

                    %>


                    <tr><td colspan=3>&nbsp;</td></tr>
                   
                    <tr><td style="border-bottom:1px #999999 solid;"><b>Realiseret</b></td><td align=right style="border-bottom:1px #999999 solid;">Kost.</td><td align=right style="border-bottom:1px #999999 solid;">Salg</td></tr>
               
                       <tr><td><b>Bruttoomsætning:</b> (timer + varesalg)</td><td align=right><%=formatnumber(realKostialt, 2) %></b> <%=basisValISO %></td><td align=right><%=formatnumber(realOmsialt, 2) %></b> <%=basisValISO %></td></tr>
              
                      <tr><td><b>Nettoomsætning:</b> (timer)</td><td align=right><%=formatnumber(realtimerkost, 2) %></b> <%=basisValISO %></td><td align=right><%=formatnumber(totRealbel, 2) %></b> <%=basisValISO %></td></tr>
                
                <tr><td><b>Salgsomkostninger:</b></td><td align=right><%=formatnumber(salgsomkostKost, 2) %></b> <%=basisValISO %></td><td align=right><%=formatnumber(salgsomkostSalg, 2) %></b> <%=basisValISO %></td></tr>
                
                   
                 
                     <tr><td><b>Dækningsbidrag / Bruttofortjeneste:</b> (realiseret)</td><td align=right><%=formatnumber(dbreal)& " "& basisValISO%></td><td>&nbsp;</td></tr>
               <tr><td><b>DB:</b></td><td align=right><%=formatnumber(dbrealpro,0)%> %</td><td>&nbsp;</td></tr>
                    
                    <tr><td colspan=3>&nbsp;</td></tr>
                

                     <%'dbfaktisk
                     
        
                         dbfaktisk = (totFakbel-(realKostialt))
                         call fn_dbproc(totFakbel,dbfaktisk)
                         dbfaktiskpro = dbProc


               
                   
                    %>
                      <tr><td colspan="3" style="border-bottom:1px #999999 solid;"><b>Faktureret</b></td></tr>
                <tr><td><b>Faktureret omsætning</b></td><td align=right><%=formatnumber(totFakbel) &" "& basisValISO%> </td><td>&nbsp;</td></tr>
             
                <tr><td><b>Kost.</b> (timer kost. + varekøb):</td><td align=right><%=formatnumber(realKostialt, 2) &" "&basisValISO %></td><td align=right>&nbsp;</td></tr>
               
              
               <tr><td><b>Dækningsbidrag / Bruttofortjeneste:</b> (faktisk, faktureret oms. - real.kost)</td><td align=right><%=formatnumber(dbfaktisk)& " "& basisValISO%></td><td>&nbsp;</td></tr>
               <tr><td><b>DB:</b></td><td align=right><%=formatnumber(dbfaktiskpro,0)%> %</td><td>&nbsp;</td></tr>

                   <tr><td style="color:#999999;"><b>Faktisk timepris:</b><br />
                  (fakturet beløb ekskl. materialer og km. / realiseret timer)</td><td>&nbsp;</td><td align=right style="color:#999999;"><%=formatnumber(gnstpris) & " "& basisValISO%></td></tr>


              <!--<tr><td><b>DB realiseret:</b></td><td align=right><%=formatnumber(gnstpris) & " "& basisValISO%></td></tr>
              <tr><td><b>DB forventet:</b></td><td align=right><%=formatnumber(gnstpris) &" "&basisValISO%></td></tr>
              <tr><td><b>DB faktisk:</b></td><td align=right><%=formatnumber(gnstpris) &" "& basisValISO%></td></tr>-->
                
                </table>
                </td>
	        </tr>
            </table>

            
	        

         <br /><br /><br />&nbsp;
	     
	     </td>
		</tr>
            <%else 'luft i bunden %>
            <tr><td> <br /><br /><br />&nbsp;</td></tr>
        <%end if %>



	   	</table>
		</div>
        




<%
end sub


     sub matforbrugGrpSubTot_fn 
                    
    
    
            %>
            <tr>
                <td colspan="5" class="lille"><b><%=lastGrpNavn%> ialt:</b></td>
                 <td align="right" class="lille"><b><%=formatnumber(matforbrugGrpSubTot, 2) %></b></td>
                <td align="right" class="lille"><b><%=formatnumber(gruppekontoBudgetTot, 2) %></b></td>
            </tr>    
            <%
         

end sub


sub timepriser

                    
                    
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
						<td colspan=4 class=alt valign=top style="padding:5px 5px 0px 5px;"><h3 class="hv">Medarbejder timepriser (salgspris)<br /><span style="font-size:10px;">Hver realiseret time bliver omsat med følgende salgs-timepris for den enkelte medarbejder</span></h3></td>
					</tr>
					<tr>
						
						<td style="padding-left:5px;" colspan=4><br>
						Tildel timepriser for de medarbejdere der er tilknyttet dette job. (via projektgrupper)<br>
                        Timepriser 1-5 er hentet fra medarbejdertypen.<br />
                        <%if level = 1 then %>
                        <a href="medarbtyper.asp" class=vmenu target="_blank">Ret kostpriser her >></a><br />
                        <%end if %>  

                        <%if func = "opret" then%>
						Jobbet skal være oprettet før der kan tildeles alternative timepriser!
						<%else%>

                        <%
                        
                        akttyper2009(2)
                        
                        aty_sql_realhours = replace(aty_sql_realhours, "tfaktim", "fakturerbar")

                        strSQLakt = "select a.id, a.navn, a.fase, COUNT(t.id) AS antaltp, brug_fasttp FROM aktiviteter AS a "_
                        &" LEFT JOIN timepriser AS t ON (t.aktid = a.id) WHERE "_ 
                        &" a.job =" & id &" AND ("& aty_sql_realhours &") GROUP BY a.id ORDER BY a.fase, a.navn" 
                        

                        'Response.write strSQLakt
                        'Response.flush
                        %>

                        
                        <br />
                        <b>Vælg job eller aktivitet:</b><br />Antal timepriser angivet / nedarver fra job / fast timepris <br /> <!--(opdater gerne flere på en gang [CTRL] + [F5])<br />-->

                        <input type="hidden" name="FM_hd_tp_jobaktid" id="FM_hd_tp_jobaktid" value="<%=tp_jobaktid%>">
                        <select id="FM_tp_jobaktid" name="FM_tp_jobaktid" style="width:450px;" size="5">
                        <%if cdbl(tp_jobaktid) = 0 then
                        tp_jobaktidNULSel = "SELECTED"
                        else
                        tp_jobaktidNULSel = ""
                        end if %>
                        
                        <option value="0" <%=tp_jobaktidNULSel %>>Jobbet (<%=strNavn %>)</option>
                        <option value="0" disabled>eller vælg aktivitet....</option>
                        <option value="0" disabled></option>
                        <%

                       
                       lastFase = ""
                        oRec5.open strSQLakt, oConn, 3
                        while not oRec5.EOF
                       
                        if isNull(oRec5("fase")) = true then
                        
                        else
                            
                            if lastFase <> oRec5("fase") then%>
                            <option value="<%=oRec5("id") %>" disabled></option>
                            <option value="<%=oRec5("id") %>" disabled><%=oRec5("fase") %></option>
                        <%  end if

                        lastFase = oRec5("fase")
                        end if
                        
                        
                        if cdbl(tp_jobaktid) = oRec5("id") then
                        tp_jobaktidSEL = "SELECTED"
                        else
                        tp_jobaktidSEL = ""
                        end if

                        if cint(oRec5("antaltp")) = 0 then
                        nedarverTxt = " (nedarver)"
                        else
                        nedarverTxt = " (" & oRec5("antaltp") &")"
                        end if

                        if cint(oRec5("brug_fasttp")) = 1 then
                        brug_fasttpDIS = "disabled"
                        brug_fasttpTxt = " - fast (ens) timepris angivet på aktivitet"
                        else
                        brug_fasttpDIS = ""
                        brug_fasttpTxt = ""
                        end if
                        %>
                       <option value="<%=oRec5("id") %>" <%=tp_jobaktidSEL %> <%=brug_fasttpDIS %>><%=oRec5("navn") %> <%=nedarverTxt %> <%=brug_fasttpTxt %></option>
                       <%
                       'Response.flush
                       oRec5.movenext
                       Wend
                       oRec5.close%>
						</select>


                       


                        <br />
                        <%select case lto
                          case "xoko"
                            showAktTpsync = 0
                          case else
                            showAktTpsync = 1
                          end select 
                            
                            
                            
                         if cint(showAktTpsync) = 1 then%>
                         <input id="FM_sync_tp" name="FM_sync_tp" value="1" type="checkbox" />Sæt alle aktiviteter til at <b>nedarve</b> de timepriser der er <b>angivet på jobbet</b>. (Ikke KM aktiviteter)
                         <br /><span style="color:#999999;">Overskriver ikke hvis "fast (ens) timepris for alle medarbejdere" er slået til på aktiviteten.</span>
                         
                            <%if level = 1 then %>
                            <br /><input name="FM_sync_tp_rens" value="1" type="checkbox" /> Slet timepriser på medarbejdere der ikke længere har adgang til dette job, eller er blevet de-aktiveret.
                            <%end if %>
                         
                         <%end if 
                             
                           strSQLmt = "SELECT mt.id, mt.type, COUNT(m.mid) AS antalM, m.mansat FROM medarbejdertyper AS mt "_
                           &" LEFT JOIN medarbejdere AS m ON (m.medarbejdertype = mt.id AND m.mansat = 1) "_
                           &" WHERE mt.id <> 0 AND m.mansat = 1 GROUP BY mt.id ORDER BY mt.type" 
                       
                       'Response.Write "<br>"& strSQLmt & "<br>"
                       'Response.flush
                       %>

                       <br /><br />
                       <b>Medarbejdertype filter:</b><br />Sorter mellem de medarb. der er tilknyttet jobbet via deres projektgrupper.<br />
                            
                       <select style="width:350px;" id="FM_mtype" name="FM_mtype">
                       
                       <%if cint(mtype) = 0 then
                       mtypeNulSel = "SELECTED"
                       else
                       mtypeNulSel = ""
                       end if %>
                       
                       <option value="0" <%=mtypeNulSel %>>Alle</option>
                       
                       <% oRec5.open strSQLmt, oConn, 3
                        while not oRec5.EOF 
                        
                        if cint(mtype) = cint(oRec5("id")) then 
                        mtypeSEL = "SELECTED"
                        else
                        mtypeSEL = ""
                        end if%>
                        <option value="<%=oRec5("id") %>" <%=mtypeSEL%>><%=oRec5("type") %> (<%=oRec5("antalM") %>)</option>
                        <%
                        oRec5.movenext
                        wend
                        oRec5.close%>

                       </select>

						<br><br>
						    
                            <script language="javascript">
						    $(document).ready(function () {
						        $("#timepristable").table_checkall();
						    });</script>
								
								
                                
                                <%
                                
                                
                                if cdbl(tp_jobaktid) <> 0 then %>
                               
                                <%call timepriser_akt %>
                                <%else %>
                                <%call timepriser_job %>
                                <%end if %>

                             

                                	
								

                                </td>
					</tr>
					<tr>


                                <tr>
					<td colspan=4 align="right" bgcolor="#FFFFFF" style="border-top:1px #CCCCCC solid; padding:20px 40px 10px 10px;">
					     <input id="timepriser_opd" type="submit" value=" Opdater timepriser >> " /></td>
					</tr>
								 <tr>
					<td colspan=4 bgcolor="#FFFFFF" style="border-top:0px #CCCCCC solid; padding:20px 40px 10px 10px;">
								<%
								
								'uTxt = "<b>Ovenstående timepriser</b> bliver kun benyttet hvis jobbet <b>ikke er et fastpris job.</b> "_ 
								'&"Ved fastpris benyttes den almindelige fastpris beregning der er angivet på job og aktiviteter.<br><br>
                                uTxt = "<b>Opdater eksisterende timepriser</b><br />"_
								&"Hvis timepriser ændres, opdateres alle eksisterende time-registreringer (ikke de for-valgte timepriser på jobbet) på aktiviteter der nedarver timepris fra job.<br>"_
                                &"<br />Hvis der er valgt en specifik aktivitet opdateres kun denne aktivitet.<br><br>"_ 
                                &"Opdater eksisterende timepriser fra: <input type='text' name='FM_opdatertpfra' value='"& formatdatetime(now,2) &"'> til dd.<br>"_
                                &"Gælder også selvom der foreligger en faktura, eller perioden er afsluttet.<br><br>"_
                                &"Hvis der ændres <b>projektgrupper</b>, skal du opdatere jobbet, før den aktuelle liste af medarbejdere vises her på timepris-siden.<br><br>"

                                if cdbl(tp_jobaktid) <> 0 then
                                uTxt = uTxt &"<b>Multiopdater</b><br><input type='checkbox' name='FM_opdateralleakt' value='1'> Opdater aktiviteter med samme navn som valgte på ALLE job. (for valgte medarbejdere)"
                                end if
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
				
			    <tr>
					<td colspan=4 align="right" bgcolor="#FFFFFF" style="border-top:1px #CCCCCC solid; padding:20px 40px 10px 10px;">
					     <input id="Submit4" type="submit" value=" Opdater & Afslut >> " /></td>
					</tr>
				</table>
				<br><br><br>&nbsp;
               </div>


<%
end sub


public strFomr_navn, strFomr_id, visJobFomr
function forretningsomrJobId(jobid)


                             '*** Forretningsområder **' 
	                            strFomr_navn = ""
                                strFomr_id = ""
                                visJobFomr = 0


                                '**** Job ****'
                                strSQLfrel = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
                                &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_jobid = "& jobid & " AND for_aktid = 0 GROUP BY for_fomr"

                                'Response.Write strSQLfrel
                                'Response.flush
                                f = 0
                                oRec3.open strSQLfrel, oConn, 3
                                while not oRec3.EOF

                                if f = 0 then
                                strFomr_navn = "<b>Job:</b> "
                                end if

                                strFomr_navn = strFomr_navn & oRec3("navn") & ", " 
                                strFomr_id = strFomr_id &",#"& oRec3("for_fomr") & "#"

                                if instr(strFomr_rel, "#"&oRec3("for_fomr")&"#") <> 0 then
                                visJobFomr = 1
                                else
                                visJobFomr = visJobFomr
                                end if

                                f = f + 1
                                oRec3.movenext
                                wend
                                oRec3.close




                                '************ Aktiviteter *****'
                                 strSQLfrel = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
                                &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_jobid = "& jobid & " AND for_aktid <> 0 GROUP BY for_fomr"

                                'Response.Write strSQLfrel
                                'Response.flush
                                f = 0
                                oRec3.open strSQLfrel, oConn, 3
                                while not oRec3.EOF

                                if f = 0 then
                                strFomr_navn = strFomr_navn & "<b>Akt.: </b>"
                                end if

                                strFomr_navn = strFomr_navn & oRec3("navn") & ", " 
                                strFomr_id = strFomr_id &",#"& oRec3("for_fomr") & "#"

                                if instr(strFomr_rel, "#"&oRec3("for_fomr")&"#") <> 0 then
                                visJobFomr = 1
                                else
                                visJobFomr = visJobFomr
                                end if

                                f = f + 1
                                oRec3.movenext
                                wend
                                oRec3.close
                            



                    



                                if f <> 0 then
                                len_strFomr_navn = len(strFomr_navn)
                                left_strFomr_navn = left(strFomr_navn, len_strFomr_navn - 2)
                                strFomr_navn = left_strFomr_navn
                                end if		    


end function



%>
