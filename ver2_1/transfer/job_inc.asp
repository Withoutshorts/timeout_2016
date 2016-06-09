<%
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
			        &" aktstatus, bgr, antalstk, budgettimer, fase, aty_desc, fakturerbar FROM "_
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
				    &"<td align=right style=""width:65px; padding-left:10px;""><b>Pris i alt DKK</b></td>"_
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
			        
                    

		            strStamgrp_pb(oRec3("aktfavorit"),a) = strStamgrp_pb(oRec3("aktfavorit"),a) &"<tr bgcolor="&bgcolor&">"_
			        &"<td><input name='FM_stakt_tilfoj_"&oRec3("aktfavorit")&"_"& oRec3("id") &"' value=""1"" type=""checkbox"" CHECKED /></td>"
			        
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
			&" WHERE ag.id <> 2 GROUP BY ag.id ORDER BY ag.navn"
		
	     %>
		 <b>Stamaktivitetsgruppe(r):</b><br />Kombiner gerne flere, hold [ctrl] nede<br /> <select class="selstaktgrp" id="selstaktgrp_<%=a%>" name="FM_favorit" size=6 multiple style="font-size:9px; width:260px; font-family:arial;">
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
						    <td style="padding-top:10px;" colspan=6><b>Udgifter / Underleverandører</b> &nbsp;&nbsp; Antal linier: xx
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
		                    <td bgcolor="#FFC0CB">=// <input type="text" id="ulevbelob_<%=u%>" name="ulevbelob_<%=u%>" value="<%=replace(formatnumber(u_belob(u), 2), ".", "") %>" style="width:60px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevfaktor_<%=u%>'), beregnulevipris('<%=u%>')"> DKK</td>
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
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)

end sub


sub kundeopl

'******************* Kunde, og kunderelateret oplysninger  *********************************'
	if showaspopup <> "y" then
	kwidth = 380
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
    <tr>
		<td colspan=4 height="20" style="padding:10px 10px 10px 10px;">
		<b>Søg:</b> (% wildcard / vis alle)&nbsp;&nbsp;<input id="kunde_sog_step1" name="FM_kunde_sog_step1" type="text" style="width:<%=kwidth%>;" />&nbsp;&nbsp;
            <input id="kunde_sog_step1_but" type="button" value=" Søg >> " /></td>
	</tr>
	<%end if%>
	
	
	<tr>
		
		<td style="padding:10px 10px 2px 10px;" colspan=4>
		
		
		<img src="../ill/ikon_kunder_24.png" alt="" border="0">&nbsp;<b>Kontakt:</b> (kunde) <br />
		<%if func = "opret" AND step = 1 then%>
		Vælg flere hvis det samme job skal oprettes på flere kunder.<br>
		<%end if %>
		
		
		    <%if func = "opret" AND step = 1 then %>
            <input name="FM_kunde" id="FM_kunde_nul" value=0 type="hidden" />
			<select name="FM_kunde" id="jq_kunde_sel" size="6" multiple="multiple" style="width:650px;">
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
			
			
            'AND kkundenavn LIKE 'a%'
            'limit 4
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
                    't = 0
				    'if t = 1000 then
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
				anstxt = "....................kontaktansv 1: "
				else
				anstxt = ""
				end if
				
			%>
			<option value="<%=oRec("Kid")%>" <%=kSel%>><%=oRec("Kkundenavn")%>&nbsp;&nbsp;(<%=oRec("Kkundenr")%>) <%=anstxt%><%=kans1%>&nbsp;&nbsp;<%=kans2%></option>
			<%
			kans1 = ""
			kans2 = ""
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
	kwidth = 380
	else
	kwidth = 310
	end if
	%>
	<tr>
		<td colspan="4" style="padding:5px 10px 5px 10px;"><b>Kontaktperson:</b> (hos kunde) <a href="#" onclick="popUp('kontaktpers.asp?id=0&kid=<%=strKundeId%>&func=opr','550','450','150','120');" class="vmenu">+ Opret ny (reload) </a><br>
		
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
		</select></td>
	
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
		                        &" ju_belob FROM job_ulev_ju WHERE ju_favorit <> 0 AND ju_jobid = 0 ORDER BY ju_favorit, ju_fase, ju_navn "
		                        oRec2.open strSQLUlev, oConn, 3
		                        
		                       
		                        while not oRec2.EOF 
		                     
		                            u_navn = oRec2("ju_navn")
		                            u_id = oRec2("ju_id")
		                     
		                            u_ipris = oRec2("ju_ipris")
		                            u_faktor = oRec2("ju_faktor")
		                            u_belob = oRec2("ju_belob")
		                            u_id = oRec2("ju_id")
                                    u_fase = "" 'oRec2("ju_fase")

                                    u_istk = oRec2("ju_stk")
                                    u_istkpris = oRec2("ju_stkpris")

                                    u_favorit = oRec2("ju_favorit")

                                    strUlevTxt(u_favorit, u) =  strUlevTxt(u_favorit, u) & "<tr>"_
                                    &"<td style=""width:30px;""><input id=""Checkbox1"" type=""checkbox"" CHECKED value='1' name='FM_ulev_linie_"& u_id &"' /></td>"_
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
                           
	                    
	                     <%if func = "xxxxx" then %>
	                    <div id="huskbudget" style="position:absolute; left:100px; top:350px; width:300px; height:100px; padding:10px; border:1px #999999 solid; background-color:#FFFFe1;">
                        <table cellspacing=4 cellpadding=1 border=0 width=100%><tr><td valign=top><img src="../ill/about_32.png" border="0" /></td>
                        
                        <td valign=top><b>Tilføj Stam-aktivitetsgruppe</b><br /><br />
                        Husk at vælge og indlæse den korrekte Stam-aktivitetsgruppe.<br /><br />
                        Tilføj en eller flere grupper ved at klikke på <b>"Indlæs gruppe på job"</b> knappen nedenfor. (ingen re-load)
                            <br />&nbsp;</td>
                        <td valign=top align=right><a id="alukhuskbudget" href="#" class=red>[x]</a></td></tr>
                        </table>
                        </div>
                        <%end if %>



                       
                  

                  

	                    
				<table cellspacing="0" cellpadding="0" border="0" bgcolor="#FFFFFF" class="pb_table">
				<tr bgcolor="#5582d2">
					<td colspan=6 class=alt style="padding:5px 5px 0px 5px; border:0px;"><h3 class="hv">Job forkalkulation og budget<br />
                    <span style="font-size:9px; color:#ffffff; line-height:12px;">Budget hovedtal - bruges til at monitorere jobbet mens det er aktivt. (WIP, work in progress)<br />
                    Opret simpel på hovedlinien, eller benyt detaljeret budget på aktiviter & faser nedenfor.</span></h3><!-- Aktiviteter og Salgsomkost.  -->
                            

                              
                    </td>
				</tr>
				<tr>
					<th align=left valign=bottom><b>Job budget, hovedlinie</b> - 
					<b>Jobnavn</b><br />(angives under stamdata til venstre)</th>
					<th align=right valign=bottom style="padding-right:10px; white-space:nowrap;"><span style="color:Red;">*</span>&nbsp;<b>Timer forkalkuleret</b><br /> <input id="FM_ign_tp" value="1" type="checkbox" /> Auto beregn ikke</th>
					<th align=left valign=bottom><b>Gns. intern kostpris pr. time</b><br />
					<span style="color:#cccccc; font-size:9px; font-family:arial;">
					(fra medab. typer: 
						    
					<%
						    
					strSQLgt = "SELECT  (mt.kostpris * count(m.mid)) AS virkgnskostpris, count(m.mid) AS antalm"_
					&" FROM medarbejdertyper mt "_
                    &" LEFT JOIN medarbejdere m ON "_
                    &" (m.medarbejdertype = mt.id AND m.mansat <> 2)"_
                    &" WHERE kostpris <> 0 GROUP BY mt.id"
						    
					'Response.Write strSQLgt
					'Response.flush
						    
					virkgnskostpris = 0
					antalM = 0
						    
						    
					oRec2.open strSQLgt, oConn, 3
					while not oRec2.EOF
						    
					virkgnskostpris = virkgnskostpris + oRec2("virkgnskostpris")
					antalM = antalM + cint(oRec2("antalm"))
						    
					oRec2.movenext
					wend
					oRec2.close 
						    
						    
					if antalM <> 0 AND virkgnskostpris <> 0 then 
					gnsinttpris = virkgnskostpris/antalM%>
					<%else 
					gnsinttpris = 0%>
					<%end if %>
					~ <%=formatnumber(gnsinttpris, 2) %>)
						    
					</span>
						  
					<%if func <> "red" then
					jo_gnstpris = gnsinttpris 
					end if%>
						  
					</th>
					<th  style="width:45px;" align=left valign=bottom><b>Faktor</b></th>
					<th align="left" valign="bottom" style="width:100px;"><span style="color:Red;">*</span>&nbsp;<b>Nettooms. DKK</b>
					<br /> (oms. før salgsomk.<br />
					timer * intern kostpris * faktor)</th>
						 
					<th align=left valign=bottom style="width:40px;">Target<br />
					timepris</th>
						  
				</tr>
						
                        

				<tr bgcolor="#FFFFFF">
					<td style="width:250px; padding-top:8px; height:30px;" class="td_projekt"><span id="pb_jobnavn" style="font-size:11px; width:270px; padding:3px; font-family:arial; border:1px #cccccc dashed;"></span>&nbsp;</td>
						  
					<td align=right style="padding-right:10px;"><input type="text" id="FM_budgettimer" name="FM_budgettimer" value="<%=replace(formatnumber(strBudgettimer, 2), ".", "")%>" style="width:60px; font-size:11px; font-family:arial; border:1px red solid;" onkeyup="tjektimer('FM_budgettimer'), beregnintbelob()"></td>
					<td><input type="text" id="FM_gnsinttpris" name="FM_gnsinttpris" value="<%=replace(formatnumber(jo_gnstpris, 2), ".", "")%>" style="width:60px; font-size:11px; font-family:arial;" onkeyup="tjektimer('FM_gnsinttpris'), beregnintbelob()"></td>
					<td style="white-space:nowrap;">x <input type="text" id="FM_intfaktor" name="FM_intfaktor" value="<%=replace(formatnumber(jo_gnsfaktor, 2), ".", "")%>" style="width:30px; font-size:11px; font-family:arial;" onkeyup="tjektimer('FM_intfaktor'), beregnintbelob()"></td>
					<td style="padding:2px 2px 2px 10x;">= <input type="text" id="FM_interntbelob" name="FM_interntbelob" value="<%=replace(formatnumber(jo_gnsbelob, 2), ".", "")%>" style="width:75px; font-size:11px; font-family:arial; border:1px red solid;" onkeyup="tjektimer('FM_interntbelob'), beregninttp()"></td>
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


                <tr><td colspan=6 valign="top" style="padding:8px 2px 2px 2px; border:0px; font-size:11px; font-family:arial;">Jobtype:<br />
                    
                    <%'select case lto
                    'case "execon"
                    idfp = "xx_angfp0" 'deaktiveret
                    'case else
                    'idfp = "angfp0"
                    'end select %>

                    <input type="radio" id="<%=idfp %>" name="FM_fastpris" value="1" <%=varFastpris1%>> <b>Fastpris</b> 
                            

                     <div id="div_angfp" style="padding:0px 2px 0px 15px; width:450px; background-color:#FFFFFF; border:0px #cccccc solid; font-size:11px; font-family:arial;">
                     <table width=100% cellspacing=2 cellpadding=2><tr><td valign=top style="font-size:11px; border:0px;"">
                 
                    <!--
                    Skal det være forkalkulation på job (angivet ovenfor) eller på hver enkelt aktivitet (angivet nedenfor) der skal være grundlag for beregning
                    af <b>faktisk timepris</b> på aktiviteterne (og stå fortrykt ved fakturering)?
                    -->
                       
                        
                    <%          
                    '** Benyt simpel eller detaljeret er DE-aktiveret. Bruger altid simpel. Men er der forkalkuleret på aktiviteter benyttes disse ved faktura oprettelse.
                    '** Beregning af omsætning mm. følger enten fastpris beregning, eller lbn. timer (medarbejder timepriser)
                    '** Intern: intern kostpris pr. medarb. uanset job 

			   
				usejoborakt_tp1_CHK = ""
				usejoborakt_tp0_CHK = "CHECKED" 
				
					    
					    
				%>
               <!-- Den valgte fastpris indstilling bliver brugt ved faktura oprettelse og til at beregne realiseret timepris i statistikken.<br />-->
                <br /><b>Benyt simpel eller detaljeret budgettering:</b><br />

                    <input id="Radio2" name="FM_usejoborakt_tp" type="radio" value="0" <%=usejoborakt_tp0_CHK %> /> <b>1) </b>Benyt forkalkulation på joblinien ovenfor (simpel)<br /> 
                    <input id="Radio3" name="FM_usejoborakt_tp" type="radio" value="1" <%=usejoborakt_tp1_CHK %> /> <b>2)</b> Beregn hver aktivitet (detaljeret)<br />
                    
                    </td></tr></table>       
                    </div>




                    <br /><input type="radio" id="angfp1" name="FM_fastpris" value="0" <%=varFastpris2%>> <b>Lbn. timer</b> (medarb. timepriser benyttes)
                    
                    <%if func = "red" then %>
                    <br /><br /><input id="jq_vispasogluk" type="checkbox" /> <b>Vis passive og lukkede aktiviteter</b>
                    <%end if %>

                    </td>
                       
                </tr>

                

                <tr bgcolor="#FFFFFF">
					<td style="padding:10px 0px 7px 445px; border:0px; border-bottom:3px #D6Dff5 solid;" colspan="6">
                   
                    <br />
		           <span style="padding:5px; background-color:#FFFFFF; border:1px #8caae6 solid; border-bottom:0px;"><a href="#stgrp_tilfoj" id="stgrp_tilfoj" name="stgrp_tilfoj">[+] Tilføj Stam-aktivitetsgrp. til job</a></span>
                    </td>
			    </tr>
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
                    
                                <div id="tilfojstamdiv" style="padding:5px 5px 5px 2px; border:10px #CCCCCC solid; background-color:#FFFFFF; position:absolute; top:200px; left:20px; width:650px; z-index:20000000;"><!-- AktTD Div -->
                               
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
                                    case "epi"
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
					                case "dencker", "acc", "essens", "fe", "jttek", "wwf", "intranet - local"
					                tpCHK1 = ""
					                tpCHK2 = "CHECKED"
							
					                case "execon", "immenso"
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
					                <input type="radio" name="FM_timepriser" value="0" <%=tpCHK1 %>> <b>1)</b> Brug den timepris hver <b>medarbejder</b> er oprettet med. (nedarv timepriser fra job, se medarbejder-typer). 
					                <br /><input type="radio" name="FM_timepriser" value="1" <%=tpCHK2 %>> <b>2)</b> Behold de medarbejder-timepriser <b>stam-aktiviteterne</b> er født med. (se stam-aktivitetsgrupper)
					               </div>




                                    </div>

                                    <br /><br />&nbsp;

					<table width=100% cellspacing=0 cellpadding=0 border=0>
						
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
	                    strSQLfv = "SELECT id, forvalgt, navn FROM akt_gruppe WHERE forvalgt = 1"
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

                  select case lto
                    case "epi", "intranet - local"
                    aktlist_wzb = "hidden"
                    aktlist_dsp = "none"
                    case else
                    aktlist_wzb = "visible"
                    aktlist_dsp = ""
                    end select
				
               
                %>
						
						
				
						 
				<tr bgcolor="#D6Dff5">
					<td style="padding:10px 0px 0px 5px;" colspan="6"><h4><a href="#" id="a_aktlisten">[+]</a> Aktiviteter & Faser

                    <%if func = "red" then%>
                     &nbsp;<a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&jobid=<%=id%>&jobnavn=<%=strNavn%>&rdir=<%=rdir%>&nomenu=1', '', 'width=1004,height=800,resizable=yes,scrollbars=yes')" class=rmenu>Rediger aktivitetliste udviddet >> </a>
                    <%end if%>
                    <br />
                    <span style="font-size:9px; color:#000000; line-height:12px;">Budget detaljeret - beregn dit projekt og sync. tal til job budget hovedlinien foroven.<br />
                    Ved fastpris job bliver nedenstående overført til fakturering.</span></h4></td>
                   

				</tr>
				<!-- henter oprettede aktiviter jq -->
                <%if func = "red" then %>
                <tr class="dv_aktlisten" style="visibility:<%=aktlist_wzb%>; display:<%=aktlist_dsp%>;"><td colspan="6" style="border:0px;">
                     <%
                     
                    oleftpx = 20
	                otoppx = 15
	                owdtpx = 140
	                java = "Javascript:window.open('aktiv.asp?menu=job&func=opret&jobid="&id&"&id=0&fb=1&rdir=job3&nomenu=1', '', 'width=1004,height=800,resizable=yes,scrollbars=yes')"    
	                call opretNyJava("#", "Opret ny aktivitet", otoppx, oleftpx, owdtpx, java) 
                     
                     %><br />&nbsp;
                </td></tr>

                <%end if %>

				<tr><td colspan="6" style="border:0px;"><div id="jq_aktlisten" class="dv_aktlisten" style="visibility:<%=aktlist_wzb%>; display:<%=aktlist_dsp%>;">Henter aktiviteter...</div></td></tr>
				
						

                <tr class="dv_aktlisten" style="visibility:<%=aktlist_wzb%>; display:<%=aktlist_dsp%>;" bgcolor="#FFCOCB"><!-- #FFC0CB -->
					<td style="padding:5px 0px 5px 5px;"><b>Aktiviteter & Faser ialt:</b></td>
					<td style="padding:2px 0px 3px 75px;"><span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:40px; border-bottom:2px #FFC0CB solid;" id="fasertimertot"><b>0,00</span></td>
                    <input id="FM_fasertimertot" value="0" type="hidden" />
					<td colspan=2>&nbsp;</td>
					<td style="padding:2px 0px 3px 40px; width:140px; white-space:nowrap;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border-bottom:2px #FFC0CB solid;;" id="fasersumtot"><b>0,00</b></span></td>
                    <input id="FM_fasersumtot" value="0" type="hidden" />
					<td>&nbsp;</td>
				</tr>

                     <%
                     '** Starter altid med at være utjekket
                     'select case lto
                     'case "epi"
                     syncCHK = ""
                     'case else
                     'syncCHK = "CHECKED"
                     'end select %>

                    <tr class="dv_aktlisten" style="visibility:<%=aktlist_wzb%>; display:<%=aktlist_dsp%>;">
					<td style="padding:20px 0px 10px 5px;">
                      <!--<input id="sync" type="button" value="Sync. forkalkulation >> " />--><input id="sync" type="checkbox" <%=syncCHK %> /><b>Sync.</b> (overfør budget på aktiviteter til budget på job)
                       
                       
                      </td>
                    <td style="padding:2px 0px 3px 75px;"><span style="padding:2px 2px 2px 0px; background-color:#FFFFFF; width:40px; border-bottom:1px #999999 solid;" id="diff_timer"><b><%=formatnumber(diff_timer, 2) %></b></span></td>
                    <input id="FM_diff_timer" name="FM_diff_timer" value="<%=formatnumber(diff_timer, 2) %>" type="hidden" />
					<td colspan="2">
                            
                    <%if diff_timer <> 0 then
                    df_t_Vzb = "hidden"
                    df_t_Dsp = "none"
                    else
                    df_t_Vzb = "visible"
                    df_t_Dsp = ""
                    end if %>
                            
                    <span id="diff_timer_ok" style="visibility:<%=df_t_Vzb%>; display:<%=df_t_Vzb%>;"><img src="../ill/ok.png" border="0" /></span>&nbsp;</td>
					<td style="padding:2px 0px 3px 40px; white-space:no-wrap;">= <span style="padding:2px 2px 2px 0px; background-color:#FFFFFF; width:60px; border-bottom:1px #999999 solid;" id="diff_sum"><b><%=formatnumber(diff_sum, 2) %></b></span></td>
                    <input id="FM_diff_sum" name="FM_diff_sum" value="<%=formatnumber(diff_sum, 2) %>" type="hidden" />
					<td>
                             
                    <%if diff_sum <> 0 then
                    df_s_Vzb = "hidden"
                    df_s_Dsp = "none"
                    else
                    df_s_Vzb = "visible"
                    df_s_Dsp = ""
                    end if %>
                            
                    <span id="diff_sum_ok" style="visibility:<%=df_s_Vzb%>; display:<%=df_s_Vzb%>;"><img src="../ill/ok.png" border="0" /></span>&nbsp;
                       
                       
                       
                       
                    </div> <!-- AktTD Div -->        
                    &nbsp;</td>
						   
				</tr>
                
                <!--- Budget på medarbejertyper **** -->

                <% 
				if func = "red" then


                if lto = "epi" OR lto = "intranet - local" then

                select case lto
                case "epi", "intranet - local"
                budgetmtyp_vzb = "visible"
                budgetmtyp_dsp = ""
                case else
                budgetmtyp_vzb = "hidden"
                budgetmtyp_dsp = "none"
                end select

				%>
						
						
				<tr bgcolor="#FFFFFF">
					<td style="padding:10px 0px 0px 5px; border:0px;" colspan="6"><br /><br />&nbsp;</td>
                </tr>
						 
				<tr bgcolor="#DCF5BD">
					<td style="padding:10px 0px 0px 5px;" colspan="6"><h4><a href="#" id="a_budgetmtyp">[+]</a> Budget på medarbejdertyper

                    <%if level = 1 then %>
                       <a href="medarbtyper.asp" class=rmenu target="_blank">Se kost- og time -priser på medarbejdertyper her >></a>
                       <%end if %>

                    <br />
                    <span style="font-size:9px; color:#000000; line-height:12px;">Budget på medarbejdertyper indgår ikke i beregning af DB på job (nederst).</span>
                    
                    </h4>
                
                </td>
                </tr>
                <tr class="tr_budgetmtyp" id="tr_budgetmtyp" style="visibility:<%=budgetmtyp_vzb%>; display:<%=budgetmtyp_dsp%>;">
                <td style="border:0px; padding:20px 20px 20px 0px;" colspan=6>

                <table cellpadding=0 cellspacing=0 border=0 width=100%>
                <tr>
                    <td><b>Medarbejdertype</b></td>
                    <td><b>Timer</b></td>
                    <td><b>Timepris</b></td>
                    <td><b>Faktor</b></td>
                    <td><b>Beløb</b></td>
                </tr>

                <%strSQLmtyp = "SELECT typeid, timer, mt_tb.timepris, faktor, belob, mt.type FROM medarbejdertyper_timebudget AS mt_tb "_
                &" LEFT JOIN medarbejdertyper mt ON (mt.id = typeid) WHERE mt_tb.jobid = "& id &" ORDER BY type"
                

                'Response.write strSQLmtyp
                'Response.flush

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
                
                %>
                <tr><td style="width:300px; font-size:9px;"><%=oRec3("type") %>
                <input type="hidden" id="Hidden2" name="FM_mtype_id" value="<%=oRec3("typeid") %>" /></td>
                    <td><input type="text" id="FM_mtype_ti_<%=oRec3("typeid") %>" name="FM_mtype_timer_<%=oRec3("typeid") %>" style="width:50px; font-size:9px;" value="<%=formatnumber(timerThis, 2) %>" class="mt_belob" /></td>
                     <td style="width:150px; font-size:9px; white-space:nowrap;"><input type="text" id="FM_mtype_tp_<%=oRec3("typeid") %>" class="mt_belob" name="FM_mtype_timepris_<%=oRec3("typeid") %>" value="<%=formatnumber(timepris, 2) %>" style="width:75px; font-size:9px;" /> DKK</td>
                      <td style="width:50px; font-size:9px; white-space:nowrap;"> x <input type="text" class="mt_belob" value="<%=formatnumber(faktor, 2) %>" id="FM_mtype_fa_<%=oRec3("typeid") %>" name="FM_mtype_faktor_<%=oRec3("typeid") %>" style="width:40px; font-size:9px;" /></td>
                      <td style="width:120px; font-size:9px; white-space:nowrap;">= <input type="text" class="mt_timer" id="FM_mtype_be_<%=oRec3("typeid") %>" value="<%=formatnumber(belob, 2) %>" name="FM_mtype_belob_<%=oRec3("typeid") %>" style="width:75px; font-size:9px;" /> DKK</td>
                </tr>

                <%
                typeidWrt = typeidWrt & " AND id <> " & oRec3("typeid") 
                oRec3.movenext
                wend
                oRec3.close %>
                
                
                <%strSQLmtyp = "SELECT type, id, timepris FROM medarbejdertyper WHERE id <> 0 "& typeidWrt &" ORDER BY type"
                

                tpii = 0
                oRec3.open strSQLmtyp, oConn, 3
                While not oRec3.EOF 

                if isNull(oRec3("timepris")) <> true then
                timepris = oRec3("timepris")
                else
                timepris = 0
                end if
                
                if tpii = 0 then
                %>
                

                <%end if %>
                <tr><td style="width:300px; font-size:9px;"><%=oRec3("type") %>
                <input type="hidden" id="FM_mtype_id_<%=oRec3("id")%>" name="FM_mtype_id" value="<%=oRec3("id")%>" class="mt_belob" /></td>
                    <td><input type="text" id="FM_mtype_ti_<%=oRec3("id") %>" name="FM_mtype_timer_<%=oRec3("id") %>" value="0" style="width:50px; font-size:9px;" class="mt_belob" /></td>
                     <td style="width:150px; font-size:9px; white-space:nowrap;"><input type="text" id="FM_mtype_tp_<%=oRec3("id") %>" class="mt_belob" name="FM_mtype_timepris_<%=oRec3("id") %>" value="<%=formatnumber(timepris, 2) %>" style="width:75px; font-size:9px;" /> DKK</td>
                      <td style="width:50px; font-size:9px; white-space:nowrap;"> x <input type="text" value="1" id="FM_mtype_fa_<%=oRec3("id") %>" class="mt_belob" name="FM_mtype_faktor_<%=oRec3("id") %>" style="width:40px; font-size:9px;" /></td>
                      <td style="width:120px; font-size:9px; white-space:nowrap;">= <input type="text" value="0" id="FM_mtype_be_<%=oRec3("id") %>" class="mt_timer" name="FM_mtype_belob_<%=oRec3("id") %>" style="width:75px; font-size:9px;" /> DKK</td>
                </tr>

                <%
                tpii = tpii + 1
                oRec3.movenext
                wend
                oRec3.close %>
                </table>

                </td></tr>
                  <tr class="tr_budgetmtyp" style="visibility:<%=budgetmtyp_vzb%>; display:<%=budgetmtyp_dsp%>;"><td colspan=6 align="right" style="border:0px; padding:10px 10px 5px 5px;">
                    <input class="overfortiljob_u" type="button" value="Gem ændringer på job >>" style="font-family:arial; font-size:9px;"/></td></tr> 


                    <%end if %>

                <%end if %>               
      

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
                    ulg_Vsb = "visible"
                    ulg_Dsp = ""
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
                
                
                <%if func = "red" then %>    
                <tr bgcolor="#FFFFFF"><td colspan=6 style="padding:10px 0px 7px 335px; border-bottom:3px #FFCC66 solid;"><br />
                 <span style="padding:5px; background-color:#FFFFFF; border:1px #FFCC66 solid; border-bottom:0px;">
                        
                        <a href="#" id="tilfoj_ulevgrp">[+] Tilføj Underlev. / Salgsomkostnings grupper til job</a>
                      
                        
                        </span>
                        
                </td></tr>
                


                <tr class="jq_ulevgrp" style="visibility:<%=ulg_Vsb%>; display:<%=ulg_Dsp%>;">

                <td colspan=6 style="border:0px; padding:0px 2px 2px 0px; font-size:11px;">

                               

                    <div id="span_tilfoj_ulevgrp" style="visibility:<%=sptu_Vzb%>; width:300px; border:1px red solid; padding:3px; background-color:lightpink; display:<%=sptu_dsp%>;">Der kan ikke tilføjes flere Underlev.- / Salgsomkost. -grupper da der er mere end 30 linier tilføjet allerede.<br />
                    Der kan manuelt tilføjes flere linier ovenfor. (optil 50)</div>

                     
                     <div id="tilfojulevdiv" style="padding:5px 5px 5px 0px; border:10px #CCCCCC solid; background-color:#FFFFFF; position:absolute; top:500px; left:20px; width:550px; z-index:20000000;"><!-- UlevTD Div -->
                     
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

                        'if cint(ug) = oRec5("jugrp_forvalgt") AND func <> "red" then
                        'ugrpSEL = "SELECTED"
                        'selUGrp(ug) = oRec5("jugrp_id")
                        'else
                        ugrpSEL = ""
                        'end if
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
                            
                    'if cint(strUlevSEL(ulgrp, ug)) = cint(selUGrp(ug)) then
                    'ulevLinieVzb = "visible"
                    'ulevLinieDsp = ""
                    'else
                    ulevLinieVzb = "hidden"
                    ulevLinieDsp = "none"
                    'end if

                    %>
                    <tr bgcolor="#FFFFFF" style="visibility:<%=ulevLinieVzb%>; display:<%=ulevLinieDsp%>;" class="ulev_pb_<%=ulgrp%>">
                    <td colspan="6" style="border-top:0px;">


                    <table border="0" cellpadding="0" cellspacing="0" width="100%">

                    <%if uheadisWrt = 0 then %>
                    <tr bgcolor="#FFCC66">
                    <td>&nbsp;</td>
		          
		            <td style="padding:2px 0px 2px 2px;"><b>Udgift navn / txt.</b></td>
		            <td style="padding:2px 0px 2px 2px;"><b>Stk. / stk. pris </b></td>
                    <td style="padding:2px 20px 2px 2px;" align=right><b>Indkøbspris</b></td>
		            <td style="padding:2px 20px 2px 2px;" align=right><b>Faktor</b></td>
		            <td style="padding:2px 20px 2px 2px;" align=right><b>Salgspris DKK</b></td>
		                 
		            </tr>
                    <%
                            
                    uheadisWrt = 1
                    end if %>

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

                
                
                
                
                
                
                if func = "red" then

                         
                '*** Ulev ***'
		        dim u_navn, u_ipris, u_faktor, u_belob, u_fase, u_id, u_istk, u_istkpris
		        redim u_navn(50), u_ipris(50), u_faktor(50), u_belob(50), u_fase(50), u_id(50), u_istk(50), u_istkpris(50)
		                
		            if func = "red" then
		                        
		                uopen = 0
		                        
		                strSQLUlev = "SELECT ju_id, ju_navn, ju_ipris, ju_faktor, "_
		                &" ju_belob, ju_fase, ju_stk, ju_stkpris FROM job_ulev_ju WHERE ju_jobid = "& id  & " ORDER BY ju_navn "
		                       
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
                        
                        
                        <%end if %>

                         	</td>
				</tr>

			
				<tr bgcolor="#FFCC66">
					<td style="padding:10px 0px 0px 5px;" colspan=6><h4><a href="#" id="a_salgsomk">[+]</a> Salgsomkostninger / Underleverandører<br />
                    <span style="font-size:9px; color:#000000; line-height:12px;">Budget detaljeret</span></h4> </td>
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
                                   
                        </select> (maks 50 linier)</td>
				</tr>  
						
                   <tr class="tr_salgsomk" bgcolor="#FFFFFF">
                 <td colspan=6>
                    <table cellpadding=0 cellspacing=0 border=0 width=100%>
                        
		            <tr>
		            <td style="padding:10px 0px 3px 2px;"><b>* Udgift/Salgsomkost.</b></td>
                    <td style="padding:10px 0px 3px 2px;"><b>Stk. / Stk. pris.</b></td>
		            <td align="right" style="padding:10px 20px 3px 2px;"><b>Indkøbspris</b></td>
		            <td align="right" style="padding:10px 10px 3px 2px;"><b>Faktor</b></td>
		            <td style="padding:10px 0px 3px 10px;"><b>Salgspris DKK</b></td>
		            <td>&nbsp;</td>
		        </tr>

                
                    
                    
                    
                    
                    
		                
		        <%
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
		            <td><input type="text" id="ulevstk_<%=u%>" name="ulevstk_<%=u%>" value="<%=replace(formatnumber(u_istk(u), 2), ".", "") %>" style="width:45px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevstk_<%=u%>'), beregnulevstkpris('<%=u%>')"> stk. á &nbsp;
					<input type="text" id="ulevstkpris_<%=u%>" name="ulevstkpris_<%=u%>" value="<%=replace(formatnumber(u_istkpris(u), 2), ".", "") %>" style="width:60px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevstkpris_<%=u%>'), beregnulevstkpris('<%=u%>')"></td>
					
                    <td>= <input type="text" id="ulevpris_<%=u%>" name="ulevpris_<%=u%>" value="<%=replace(formatnumber(u_ipris(u), 2), ".", "") %>" style="width:60px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevpris_<%=u%>'), beregnulevbelob('<%=u%>')"></td>
					<td>x <input type="text" id="ulevfaktor_<%=u%>" name="ulevfaktor_<%=u%>" value="<%=replace(formatnumber(u_faktor(u), 2), ".", "") %>" style="width:40px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevfaktor_<%=u%>'), beregnulevbelob('<%=u%>')"></td>
		            <td align=right style="padding-right:33px;">= <input type="text" id="ulevbelob_<%=u%>" name="ulevbelob_<%=u%>" value="<%=replace(formatnumber(u_belob(u), 2), ".", "") %>" style="width:60px; font-size:9px; font-family:arial;" onkeyup="tjektimer('ulevfaktor_<%=u%>'), beregnulevipris('<%=u%>')"></td>
		            <td>&nbsp; <a href="#" id="ulev_ryd_<%=u%>" class="ulev_ryd">Ryd</a> </td>
		        </tr>
		                
		        <%next %>
		            
                    </table>
                    
                    </td>
                </tr>

    
                    <tr class="tr_salgsomk">
                    <td align="right" style="padding:10px 5px 3px 5px; border:0px;" colspan="6"><a href="#" class="rmenu" id="ulev_tilfoj_line">Tilføje line >></a>
                    </td>
                    </tr>
                       
                       
                <%
                
                end if
                
                
               %>

                       <input id="ulevopen" value="<%=uopen %>" type="hidden" />
                        <input id="lastopen_ulev" value="0" type="hidden" />
                        

                       
                 
              
						
						
						
				
               
            

		        <tr bgcolor="#FFFFFF">
					<td class=alt style="padding-top:10px; height:20px;" colspan=6>&nbsp;</td>
				</tr>  
		        <tr bgcolor="#ffC0CB">
						    
					<td colspan=4 style="padding:10px 0px 3px 5px; font-size:11px; font-family:arial;"><b>Brutto omsætning:</b></td>
					<td style="padding:2px 2px 2px 20px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:1px #999999 solid; font-size:9px; font-family:arial;" id="SP_budget"><b><%=replace(formatnumber(jo_bruttooms, 2), ".", "")%></b></span></td>
							
					<!--&nbsp; antal fakturerbare timer.&nbsp;<br>-->
					<input type="hidden" name="FM_ikkebudgettimer" value="<%=SQLBless3(ikkeBudgettimer)%>"><!--&nbsp;antal ikke fakturerbare timer.<br>&nbsp;-->
                    <input id="FM_budget" name="FM_budget" value="<%=replace(formatnumber(jo_bruttooms, 2), ".", "")%>" type="hidden" />
						<td>&nbsp;</td>
							
		        </tr>
		                
                    <tr bgcolor="#ffffff">
					<td colspan=4 style="padding:10px 0px 3px 5px; font-size:11px; font-family:arial;">Udgifter intern:</td>
					<td style="padding:2px 2px 2px 20px; width:140px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px; border-bottom:1px #CCCCCC solid; font-size:9px; font-family:arial;" id="SP_udgifter_intern"><%=replace(formatnumber(jo_udgifter_intern, 2), ".", "")%></span></td>
					<td>&nbsp;</td>
                    <input id="FM_udgifter_intern" name="FM_udgifter_intern" value="<%=replace(formatnumber(jo_udgifter_intern, 2), ".", "")%>" type="hidden" />
				</tr>

                    <tr bgcolor="#ffffff">
					<td colspan=4 style="padding:10px 0px 3px 5px; font-size:11px; font-family:arial;">Salgsomkostninger (underlev./indkøb):</td>
					<td style="padding:2px 2px 2px 20px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px; border-bottom:1px #CCCCCC solid; font-size:9px; font-family:arial;" id="SP_udgifter_ulev"><%=replace(formatnumber(jo_udgifter_ulev, 2), ".", "")%></span></td>
					<input id="FM_udgifter_ulev" name="FM_udgifter_ulev" value="<%=replace(formatnumber(jo_udgifter_ulev, 2), ".", "")%>" type="hidden" />
                    <td>&nbsp;</td>
				</tr>
		                
                <tr bgcolor="#ffffff">
					<td colspan=4 style="padding:10px 0px 3px 5px; font-size:11px; font-family:arial;"><b>Udgifter ialt: (intern + salgsomk.)</b></td>
					<td style="padding:2px 2px 2px 20px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px; border-bottom:1px #999999 solid; font-size:9px; font-family:arial;" id="SP_udgifter"><b><%=replace(formatnumber(udgifter, 2), ".", "")%></b></span></td>
                        <input id="FM_udgifter" name="FM_udgifter" value="<%=replace(formatnumber(udgifter, 2), ".", "")%>" type="hidden" />
                    <td>&nbsp;</td>
				</tr>
		                
		            <tr bgcolor="#Eff3ff">
						    
					<td colspan=4 style="padding:10px 0px 3px 5px; font-size:11px; font-family:arial;"><b>Dækningsbidrag / Bruttofortjeneste</b></td>
					<td style="padding:2px 2px 2px 20px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px; border-bottom:1px #999999 solid; font-size:9px; font-family:arial;" id="SP_bruttofortj"><b><%=replace(formatnumber(jo_bruttofortj, 2), ".", "") %></b></span></td>
						<input id="FM_bruttofortj" name="FM_bruttofortj" value="<%=replace(formatnumber(jo_bruttofortj, 2), ".", "") %>" type="hidden" />
                        <td>&nbsp;</td>
				</tr>
				<tr bgcolor="#ffffff">
						    
					<td colspan=4 style="padding:10px 0px 3px 5px; font-size:11px; font-family:arial;"><b>(DB) %</b></td>
					<td style="padding:2px 2px 2px 20px;">= <span style="padding:2px 2px 2px 2px; background-color:#FFFFFF; width:60px; border:0px; border-bottom:1px #999999 solid; font-size:9px; font-family:arial;" id="SP_db"><b><%=formatnumber(jo_dbproc, 0) %></b></span></td>
					<input id="FM_db" name="FM_db" value="<%=formatnumber(jo_dbproc, 0)%>" type="hidden" />
                    <td>&nbsp;</td>
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
</div>
<!-- projekt beregner slut -->
<%
end if


end sub



sub minioverblik



		
    oTop = 0
	oLeft = 487
	oWdth = 700
	oHgt = "" 
	oId = "oblik_div"
	oZindex = 1000

	call tableDivAbs(oTop,oLeft,oWdth,oHgt,oId, oVzb, oDsp, oZindex)
	%>
		
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr bgcolor="#5582d2">
						    <td colspan=6 class=alt style="padding:5px 5px 0px 5px; border:0px;"><h3 class="hv">Joboverblik & Funktioner</h3></td>
						</tr>
		<tr>
		<td style="padding:10px 4px 5px 4px;">
	   <%if func = "red" then 
		minusyear = dateadd("m", -12, now)
		
		dtlink_stdag = datepart("d", minusyear, 2, 2)
		dtlink_stmd = datepart("m", minusyear, 2, 2)
		dtlink_staar = datepart("yyyy", minusyear, 2, 2)
		
		dtlink_sldag = datepart("d", now, 2, 2)
		dtlink_slmd = datepart("m", now, 2, 2)
		dtlink_slaar = datepart("yyyy", now, 2, 2)
		
       
		
		dtlink = "FM_usedatokri=1&FM_start_dag="&dtlink_stdag&""_
	    &"&FM_start_mrd="&dtlink_stmd&"&FM_start_aar="&dtlink_staar&"&FM_slut_dag="&dtlink_sldag&""_
	    &"&FM_slut_mrd="&dtlink_slmd&"&FM_slut_aar="&dtlink_slaar
		
		%>
		

        <a href="job_kopier.asp?func=kopier&id=<%=id%>&fm_kunde=<%=fm_kunde_sog%>&filt=<%=request("filt")%>" class=vmenu>Kopier job >></a>
		<br /><a href="jobs.asp?menu=job&func=slet&id=<%=id %>" class=slet>Slet / nulstil job?</a><br />



        </td>
        </tr>
        <tr>
		    <td><br /><b>Joblog</b> (seneste 24 md. hvis job er aktivt, ellers jobperiode dog maks 24 md., timer er afrundet)</td>
		</tr>
        
        <tr>
        <td bgcolor="#Eff3ff" style="padding:4px 4px 5px 4px; border:1px #cccccc solid;" align=center>
		
        <table cellpadding=0 cellspacing=0 border=0>
        <tr bgcolor="#FFFFFf">
         <td style="width:40px; border-bottom:1px #cccccc solid; border-right:1px #cccccc solid;" align=center>&nbsp;</td>
            <%
            

            '*** StDato ***'
            if strStatus <> 1 AND cDate(strUdato) <> cdate("1-1-2044") then 'passiv / lukket og uendelig
            stDatoKriGrf = strUdato 
            stDatoKriGrfDiff = datediff("m", strTdato, strUdato , 2, 2) 
            
            if stDatoKriGrfDiff <= 24 then 
            mts = stDatoKriGrfDiff
            else
            mts = 24
            end if

            else
            stDatoKriGrf = now
            
            mts = 24
            end if

            dim timerThis, kostThis, omsThis, resThis
            redim timerThis(mts), kostThis(mts), omsThis(mts), resThis(mts)
            call akttyper2009(2)
            
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
          <td style="border-bottom:1px #cccccc solid; border-right:1px #cccccc solid;" align=center>&nbsp;</td>
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

             strSQLjl = "SELECT timer, (kostpris * timer * (kurs / 100)) AS kostpris, timepris, tfaktim FROM timer WHERE tjobnr = "& strjobnr &" AND tdato BETWEEN '"& dtnowLowSQL &"' AND '"& dtnowHighSQL &"'"_
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


             if cdbl(resHgt) > 100 then
             resHgt = 100
             resbgThis = "#999999"
             else
             resHgt = resHgt
             resbgThis = "#999999"
             end if
             
             if cdbl(kostHgt) > 100 then
             kostHgt = 100
             kostbgThis = "#8caae6"
             else
             kostHgt = kostHgt
             kostbgThis = "#8caae6"
             end if


             if cdbl(omsHgt) > 100 then
             omsHgt = 100
             omsbgThis = "#9aCD32"
             else
             omsHgt = omsHgt
             omsbgThis = "#9aCD32"
             end if

             if cdbl(hgt) > 100 then
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
            timerThis(dt) = formatnumber(timerThis(dt), 0)
            kostThis(dt) = formatnumber(kostThis(dt) / 1000, 0)
            omsThis(dt) = formatnumber(omsThis(dt) / 1000, 0)
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
           <td style="font-size:8px; border-bottom:1px #999999 solid;" align=center>Ressource</td>
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

       


        </td>
        </tr>
        
        <tr>
		    <td  style="padding:4px 4px 5px 4px;">
            <a href="joblog.asp?FM_job=<%=id %>&FM_kunde=<%=strKnr %>&<%=dtlink %>" class=rmenu target="_blank">Joblog >></a><br />
            <%if level = 1 then %>
        <a href="medarbtyper.asp" class=rmenu target="_blank">Ret kostpriser her >></a><br />
         <%end if %>
        <a href="ressource_belaeg_jbpla.asp?FM_sog=<%=strJobnr %>" class=rmenu target="_blank">Ressource timer / Forecast >></a>
        </td>
		</tr>
       
        <tr>
		    <td><br /><br /><b>Materialer / udlæg</b></td>
		</tr>
        <tr>
        <td bgcolor="#F6DC9C" style="padding:4px 4px 5px 4px; border:1px #cccccc solid;">
		<a href="materiale_stat.asp?FM_job=<%=id %>&FM_kunde=<%=strKnr %>&<%=dtlink %>" class=vmenu target="_blank">Materialer / udlæg >></a><br />
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr>
            <td class=lille align=right><b>Antal registreringer</b></td>
            <td class=lille align=right><b>Antal stk. ialt</b></td>
            <td class=lille align=right><b>Købspris</b></td>
            <td class=lille align=right><b>Salgspris</b></td>
            <td class=lille align=right>&nbsp;</b></td>
         </tr>
	    <tr>
	        
            <%strSQLudl = "SELECT COUNT(id) AS antalreg, SUM(matantal) AS antal, sum(matkobspris * (kurs/100)) AS matkobspris, SUM(matsalgspris * (kurs/100)) AS matsalgspris FROM materiale_forbrug WHERE jobid = "& id & " GROUP BY jobid" 
            
            'Response.Write strSQLudl
            'Response.flush

            oRec2.open strSQLudl, oConn, 3
	        if not oRec2.EOF then
            %>
            <tr bgcolor="#FFCC66">
            <td align=right class=lille><%=oRec2("antalreg") %></td>
            <td align=right class=lille><%=oRec2("antal") %></td>
            <td align=right class=lille><%=formatnumber(oRec2("matkobspris"), 2) %> DKK</td>
            <td align=right class=lille><%=formatnumber(oRec2("matsalgspris"), 2) %> DKK</td>
            <td></td>
            </tr>
            <%
            end if
            oRec2.close %>
        </table>

		
		<%end if%>
		</td>
		</tr>
		
		<tr>
		    <td><br /><br /><b>Filer</b></td>
		</tr>
		<tr>
		<td bgcolor="#Eff3ff" style="padding:5px 4px 5px 4px; border:1px #cccccc solid;">
        
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
	    <tr>
	        <td colspan=3>
	        <%if func = "red" then %>
	        <a href="job_print.asp?id=<%=id%>" class=vmenu>Print / PDF Center >></a> (print job som tilbud, budget etc.)
	        <%end if %>
		</td>
		</tr>
		<tr>
		    <td class=lille><b>Folder</b></td>
		    <td class=lille><b>Navn</b></td>
		    <td class=lille align=right><b>Dato</b></td>
		</tr>
		<%strSQLdok = "SELECT f.filnavn, f.id AS fid, f.dato, fo.navn AS foldernavn FROM filer f "_
        &" LEFT JOIN foldere fo ON (fo.id = f.folderid) WHERE f.jobid =" & id & " AND f.filnavn <> '' ORDER BY foldernavn, filnavn" 
        
        oRec2.open strSQLdok, oConn, 3
        while not oRec2.EOF
        %>
            <tr>
		    <td class=lille><%=oRec2("foldernavn") %></td>
		    <td class=lille><a href="../inc/upload/<%=lto%>/<%=oRec2("filnavn")%>" class='rmenu' target="_blank"><%=oRec2("filnavn") %></a></td>
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
		
		<tr>
		    <td><br /><br /><b>Faktura historik</b> </td>
		</tr>
		<tr>
		<td bgcolor="#EAEAAE" style="padding:5px 4px 5px 4px; border:1px #cccccc solid;">
        
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
	    <tr>
	        <td colspan=4>
	        <%if func = "red" then %>
	        <a href="erp_opr_faktura_fs.asp?FM_kunde=2&FM_job=<%=id%>&FM_aftale=<%=intServiceaft%>&reset=1" class=vmenu><span style="color:green; font-size:16px;"><b>+</b></span> Opret faktura</a>
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
		&" f.faktype, f.beloeb, f.shadowcopy, f.valuta, f.kurs, v.valutakode, SUM(fd.aktpris) AS aktbel, f.medregnikkeioms, f.brugfakdatolabel, f.labeldato "_
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
	        <td class=lille><b><%=oRec2("faknr") %></b>
            <%if cint(oRec2("medregnikkeioms")) = 1 then %>
            (intern)
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
	        
	        
	        <td class=lille align=right><%=formatnumber(belobGrundVal) & " DKK" %></td>
	        
	        <%
	        '** Kun aktiviteter timer, enh. stk. IKKE mateiler og KM
                
                  if cDate(oRec2("fakdato")) < cDate("01-06-2010") AND lto = "epi" then
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
	        
	        <td class=lille align=right><%=formatnumber(belobKunTimerStk) & " DKK" %></td>
	    
	    </tr>
	    <%

        if cint(oRec2("medregnikkeioms")) <> 1 then 'intern
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
		    <td class=lille align=right colspan=3>ialt: <b><%=formatnumber(totFakbel)%></b> DKK</td>
		    <td class=lille align=right><b><%=formatnumber(totFakbelKunTimer)%></b> DKK</td>
		</tr>
		</table>
		
		 
		
		</td>
		</tr>
		<tr>
		    <td><br /><br /><b>Aktiviteter og faser</b>
            
            <%if func = "red" then%>
            &nbsp;<a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&jobid=<%=id%>&jobnavn=<%=strNavn%>&rdir=<%=rdir%>&nomenu=1', '', 'width=1004,height=800,resizable=yes,scrollbars=yes')" class=vmenu>(+ rediger aktiviteter)</a>
            <%end if%>

            </td>
		</tr>
		<tr>
		<td style="padding:5px 4px 5px 4px; border:1px #cccccc solid;"">
		
		
	     <table cellpadding=0 cellspacing=0 border=0 width=100%>
	     <tr>
	     <td>
	     <%if func = "red" then %>
	     
         <a href="#" id="forkalk2_A" class="vmenu">Job & Aktiviteter (Budget) >></a><br />
	     
         <%
         
                    oleftpx = 500
	                otoppx = -10
	                owdtpx = 140
	                java = "Javascript:window.open('aktiv.asp?menu=job&func=opret&jobid="&id&"&id=0&jobnavn="&strNavn&"&fb=1&rdir=job3&nomenu=1', '', 'width=1004,height=800,resizable=yes,scrollbars=yes')"    
	                call opretNyJava("#", "Opret ny aktivitet", otoppx, oleftpx, owdtpx, java) 
         
         %>
         
         <!--<a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&func=opret&jobid=<%=id%>&id=0&jobnavn=<%=strNavn%>&fb=1&rdir=job2&nomenu=1', '', 'width=850,height=600,resizable=yes,scrollbars=yes')" class=vmenu><span style="color:green; font-size:16px;"><b>+</b></span> Opret ny aktivitet</a>-->

	    <br />&nbsp;
	     <!-- jobs.asp?menu=job&func=red&id=<%=id%>&rdir=<%=rdir%>&showdiv=fkalk -->
         <%end if%>
	     </td>
	     </tr>
	     <tr>
	        <td>
	        <table cellpadding=0 cellspacing=0 border=0 width=100%>
	        <tr>
	        <td class=lille><b>Navn</b></td>
            <td class=lille><b>Periode</b></td>
	        <td class=lille><b>Status</b></td>
            <td class=lille><b>Type</b></td>
	        <td class=lille align=right><b>Stk.</b></td>
	        <td class=lille align=right><b>Forkalk. timer (real.)</b></td>
	        <td class=lille align=right><b>Pris (oms.)</b></td>
	        </tr>
	        
	        <%strSQLakt = "SELECT a.id, a.navn, job, beskrivelse, fakturerbar, aktstartdato, "_
	        &" aktslutdato, budgettimer, aktstatus, aktbudget, fomr.navn AS fomr, "_
	        &" antalstk, tidslaas, a.fase, a.sortorder, a.bgr, aktbudgetsum, sum(t.timer) AS realiseret, aktstartdato, aktslutdato, aty_desc "_
	        &" FROM aktiviteter a LEFT JOIN fomr ON (fomr.id = a.fomr) "_
	        &" LEFT JOIN timer AS t ON (t.taktivitetid = a.id)"_
            &" LEFT JOIN akt_typer aty ON (aty_id = a.fakturerbar) "_
	        &" WHERE job = "& id &" AND a.aktfavorit = 0 GROUP BY a.id ORDER BY a.fase, a.sortorder, a.navn" 
        	
        	totSum = 0
	        totTimerforkalk = 0
	        totReal = 0
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
               <tr>
                
                 <td>&nbsp;</td>
	            <td>&nbsp;</td>
	            <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
	            <td align=right class=lille><b><%=formatnumber(lastFaseForkalkTimer,2) %></b> (<%=formatnumber(lastFaseRealTimer,2) %>)</td>
	            <td align=right class=lille><b><%=formatnumber(lastFaseSum, 2) %></b> DKK</td>

            </tr>
              <tr><td colspan=7 style="padding:2px; font-size:9px;">&nbsp;</td></tr>
                <%
                lastFaseForkalkTimer = 0
                lastFaseRealTimer = 0
                end if %>
	            
	          
	            <%lastFaseSum = 0 %>
	            <tr bgcolor="#8cAAe6"><td colspan=7 style="padding:2px; font-size:9px;"><b><%=replace(oRec2("fase"), "_", " ")%></b></td></tr>
	            <%end if %>
	            
	            
	            <tr bgcolor="<%=bga %>">
	            <td class=lille><a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&func=red&id=<%=oRec2("id")%>&jobid=<%=id %>&jobnavn=<%=strNavn%>&rdir=job2&nomenu=1', '', 'width=850,height=600,resizable=yes,scrollbars=yes')" class="rmenu"><%=left(oRec2("navn"), 20) %></a></td>
	            
                <td class=lillegray><%=formatdatetime(oRec2("aktstartdato"), 1)  &" - "& formatdatetime(oRec2("aktslutdato"), 1) %></td>
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
                
                 <td class=lille><%=left(oRec2("aty_desc"), 10) %>.</td>
	            <td class=lille align=right><%=formatnumber(oRec2("antalstk"), 2) %> stk.</td>
	            <td class=lille align=right><%=formatnumber(oRec2("budgettimer"), 2) %> (<%=formatnumber(oRec2("realiseret"), 2) %>)</td>
	            <td class=lille align=right><%=formatnumber(oRec2("aktbudgetsum"), 2) %> DKK </td>
	            
	            </tr>
	        
	        <%
	        a = a + 1
	        lastFaseSum = lastFaseSum + oRec2("aktbudgetsum") 
	        lastFase = oRec2("fase")
	        totSum = totSum + oRec2("aktbudgetsum")
	        totTimerforkalk = totTimerforkalk + oRec2("budgettimer")
	        call akttyper2009Prop(oRec2("fakturerbar"))
	        if cint(aty_real) = 1 then
	        totReal = totReal + oRec2("realiseret")
	        else
	        totReal = totReal
	        end if

            lastFaseForkalkTimer = lastFaseForkalkTimer + oRec2("budgettimer")
            lastFaseRealTimer = lastFaseRealTimer + oRec2("realiseret")

	        oRec2.movenext
	        wend
	        oRec2.close%>
	        
	        <tr>
                
                 <td>&nbsp;</td>
	            <td>&nbsp;</td>
	            <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
	            <td align=right class=lille><b><%=formatnumber(lastFaseForkalkTimer,2) %></b> (<%=formatnumber(lastFaseRealTimer,2) %>)</td>
	            <td align=right class=lille><b><%=formatnumber(lastFaseSum, 2) %></b> DKK</td>

            </tr>
             <tr><td colspan=7 style="padding:2px; font-size:9px;">&nbsp;</td></tr>

	        <tr bgcolor="#FFDFDF">
	            <td>&nbsp;</td>
	            <td>&nbsp;</td>
	            <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
	            <td align=right class=lille><b><%=formatnumber(totTimerforkalk,2) %></b> (<%=formatnumber(totReal,2) %>)</td>
	            <td align=right class=lille><b><%=formatnumber(totSum, 2) %></b> DKK</td>
	        </tr>    
	        <%if totReal <> 0 then 
	        gnstpris = totFakbelKunTimer/totReal
	        else 
	        gnstpris = 0
	        end if%>
	        <tr>
	            <td colspan=7><br /><br /><h3>Faktisk timepris = <%=formatnumber(gnstpris) & " DKK"%></h3>
	            (fakturet beløb ekskl. matererialer og km. / realiseret timer)</td>
	        </tr>
	        </table>
	        </td>
	     </tr>
	     
	     </table>
	     
	     </td>
		</tr>
	   	</table>
		</div>
		


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
						<td colspan=4 class=alt valign=top style="padding:5px 5px 0px 5px;"><h3 class="hv">Medarbejder timepriser<br /><span style="font-size:10px;">Ved Lbn. timer bliver hver time omsat med følgende timepris for den enkelte medarbejder</span></h3></td>
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

                        strSQLakt = "select a.id, a.navn, a.fase, COUNT(t.id) AS antaltp FROM aktiviteter AS a "_
                        &" LEFT JOIN timepriser AS t ON (t.aktid = a.id) WHERE "_ 
                        &" a.job =" & id &" AND ("& aty_sql_realhours &") GROUP BY a.id ORDER BY a.fase, a.navn" 
                        
                        %>

                        
                        <br />
                        <b>Vælg job eller aktivitet:</b><br />Antal timepriser angivet / nedarver fra job<br /> <!--(opdater gerne flere på en gang [CTRL] + [F5])<br />-->

                        <input type="hidden" name="FM_hd_tp_jobaktid" id="FM_hd_tp_jobaktid" value="<%=tp_jobaktid%>">
                        <select id="FM_tp_jobaktid" name="FM_tp_jobaktid" style="width:450px;" size="5">
                        <%if cint(tp_jobaktid) = 0 then
                        tp_jobaktidNULSel = "SELECTED"
                        else
                        tp_jobaktidNULSel = ""
                        end if %>
                        
                        <option value="0" <%=tp_jobaktidNULSel %>>Jobbet (<%=strNavn %>)</option>
                        <option value="0" disabled>eller vælg aktivitet....</option>
                        <option value="0" disabled></option>
                        <%

                       

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
                        nedarverTxt = ""
                        else
                        nedarverTxt = " (" & oRec5("antaltp") &")"
                        end if

                        %>
                       <option value="<%=oRec5("id") %>" <%=tp_jobaktidSEL %>><%=oRec5("navn") %> <%=nedarverTxt %></option>
                       <%
                       
                       oRec5.movenext
                       Wend
                       oRec5.close%>
						</select>
                        <br />
                          <input id="FM_sync_tp" name="FM_sync_tp" value="1" type="checkbox" /> <b>Sync.</b> alle aktiviteter til at følge de timepriser der er <b>angivet på jobbet</b>. (Ikke KM aktiviteter)
                       

                       <%strSQLmt = "SELECT mt.id, mt.type, COUNT(m.mid) AS antalM FROM medarbejdertyper AS mt "_
                       &" LEFT JOIN medarbejdere AS m ON (m.medarbejdertype = mt.id AND m.mansat = 1) WHERE mt.id <> 0 GROUP BY mt.id ORDER BY mt.type" 
                       
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

                                strSQLfrel = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
                                &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_jobid = "& jobid & " GROUP BY for_fomr"

                                'Response.Write strSQLfrel
                                'Response.flush
                                f = 0
                                oRec3.open strSQLfrel, oConn, 3
                                while not oRec3.EOF

                                if f = 0 then
                                strFomr_navn = " ("
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
                                strFomr_navn = left_strFomr_navn & ")"
                                end if		    


end function
%>
