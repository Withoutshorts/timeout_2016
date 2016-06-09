<%
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
					
					strSQL = "SELECT jobnr, jobTpris FROM job WHERE jobnr = " & objRec("Tjobnr")
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
					
					strSQL = "SELECT jobnr, jobTpris FROM job WHERE jobnr = " & objRec("Tjobnr") 
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


			Select case md
			case "4", "6", "9", "11"

				if dag = 31 then
				dagparset = 30
				else
				dagparset = dag
				end if

			case "2"

                '** SKUDÅR ***
                select case right(aar, 2)
                case "00", "04", "08", "12", "16", "20", "24", "28", "32", "36", "40", "44"

                if dag > 29 then
				dagparset = 29
				else
				dagparset = dag
				end if
    
                case else 

				if dag > 28 then
				dagparset = 28
				else
				dagparset = dag
				end if
	
                end select

    		case else
			dagparset = dag
			end select
			
end function


public intMoms, intNetto, intModMoms, intModNetto, showMoms, intTotal_konto, intTotal_modkonto
function beregnmoms(useDebKre, intBeloeb, varKonto, varModkonto)
	
	intTotal = intBeloeb
				
				
				if useDebKre = "d" then
				intTotal_konto = intTotal
				intTotal_modkonto = 0 - replace(intTotal, ".", ",")
				intTotal_modkonto = replace(intTotal_modkonto, ",", ".")
				else
				intTotal_konto = 0 - replace(intTotal, ".", ",")
				intTotal_konto = replace(intTotal_konto, ",", ".")
				intTotal_modkonto = intTotal
				end if
				
				
				intMoms = 0
				intNetto = 0
				intModMoms = 0
				intModNetto = 0
				
				showMoms = 0
				
				
				'** Beregner moms og netto **
				'*** Konto ***
				strSQL = "SELECT momskode, kvotient FROM kontoplan LEFT JOIN momskoder ON (momskoder.id = momskode) WHERE kontonr = "& varKonto
				oRec.open strSQL, oConn, 3 
				if not oRec.EOF then
					if oRec("kvotient") <> 1 then
						
							if oRec("kvotient") = 5 then
							intMoms = (replace(intTotal_konto, ".", ",") * 0.25)  'Alm
							else
							intMoms = (replace(intTotal_konto, ".", ",") * 0.05)  'Rep
							end if
							
							
							if intMoms < 0 then
							showmoms = -(intMoms)
							else
							showmoms = intMoms
							end if
						
						intMoms = replace(intMoms, ",", ".")
						
					else
					intMoms = 0
					end if
				intNetto = intTotal_konto 
				end if
				oRec.close
				
				
				
				'*** Komma til punktum
				showMoms = replace(showMoms, ",", ".")
				
				'** Beregner moms og netto **
				'*** Modkonto ***
				strSQL = "SELECT momskode, kvotient FROM kontoplan LEFT JOIN momskoder ON (momskoder.id = momskode) WHERE kontonr = "& varModkonto
				oRec.open strSQL, oConn, 3 
				if not oRec.EOF then
					if oRec("kvotient") <> 1 then
							if oRec("kvotient") = 5 then
							intModMoms = (replace(intTotal_modkonto, ".", ",") * 0.25)  'Alm
							else
							intModMoms = (replace(intTotal_modkonto, ".", ",") * 0.05) 'Rep
							end if
							
					intModMoms = replace(intModMoms, ",", ".")
					
					else
					intModMoms = 0
					end if
				intModNetto = intTotal_modkonto 
				end if
				oRec.close
				
end function


public oprid
function opretpos()
		
	
			'***************** Opretter posteringer på de valgte konti **********
				strSQLK = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att)"_
				&" VALUES ('"& strEditor &"', '"& strDato &"', '"& varModkonto &"', '"& varKonto &"', '"& intFaknum &"', "_
				&" "& intTotal_konto &", "&intNetto&", "&intMoms&", '"& posteringsTxt &"', "_
				&" '"& fakDato &"', "_
				&" 1, "& session("Mid") &")"
				
				'Response.write strSQLK
				oConn.execute(strSQLK)
				
				'*** Finder det netop opr. id ***
				strSQL = "SELECT id FROM posteringer ORDER BY id DESC"
				oRec.Open strSQL, oConn, 3 
				if not oRec.EOF then
				oprid = oRec("id")
				end if
				oRec.close
				
				oConn.execute("UPDATE posteringer SET oprid = "& oprid &" WHERE id = "& oprid &"")
				oConn.execute("UPDATE fakturaer SET oprid = "& oprid &" WHERE Fid = "& thisfakid &"")
				
				'**** Modkonto ***
				strSQLMK = "INSERT INTO posteringer (editor, dato, modkontonr, kontonr, bilagsnr, beloeb, nettobeloeb, moms, tekst, posteringsdato, status, att, oprid)"_
				&" VALUES ('"& strEditor &"', '"& strDato &"', '"& varKonto &"', '"& varModkonto &"', '"& intFaknum &"', "_
				&" "& intTotal_modkonto &", "&intModNetto&", "&intModMoms&", '"& posteringsTxt &"', "_
				&" '"& fakDato &"', "_
				&" 1, "& session("Mid") &", "& oprid &")"
				
				oConn.execute(strSQLMK)
	
	
end function



 function joblog(jobid,stdatoKri,slutdato,aftid)
   laktaktid = 0
   
   if jobid <> 0 then
   strSQLjobaftKri = "j.id = "& jobid &" AND tjobnr = j.jobnr"
   else
   strSQLjobaftKri = "seraft = "& aftid &" AND tjobnr = j.jobnr "
   end if
   %>
    
    
				<!--<table width=100% border=0 cellspacing=0 cellpadding=0>-->
				<%
				strSQL = "SELECT tmnr, tmnavn, tdato, timer, timerkom, "_
				&" taktivitetnavn, taktivitetid, tfaktim, jobnavn, jobnr, a.faktor, timepris "_
				&" FROM timer, job j, aktiviteter a "_
				&" WHERE "&  strSQLjobaftKri &" AND tfaktim <> 5 AND "_
				&" (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"') AND a.id = taktivitetid ORDER BY jobnr, tdato DESC"
				'response.write strSQL
				'Response.Flush
				oRec.open strSQL, oConn, 3 
				x = 0
				while not oRec.EOF 
					
					if oRec("jobnr") <> lastjobnr then
						
						if oRec("tfaktim") = 1 then
						imgthis = "&nbsp;"
						else
						imgthis = "<font class=roed>(Ej fakturerbar)</font>"
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
					
					
		
						
					Response.write "<tr><td colspan=3 style=padding-top:"&pdtop&";><h5>"& oRec("jobnavn") &" ("& oRec("jobnr") &")</h5></td></tr>"
					Response.Write "<tr><td width=100><b>Dato</b></td>"_
					&"<td><b>Aktivitet</b></td>"_
					&"<td width=100 align=right><b>Antal</b></td>"
					
					if thisfile = "fak_godkendt.asp" then
					chkboxesShow1 = ""
					chkboxesShow2 = ""
					else
					chkboxesShow1 = "<input id='FM_vis_joblog_timepris' name='FM_vis_joblog_timepris' value='1' type='checkbox' "& visjoblog_timepris_CHK &" />"
					chkboxesShow2 = "<input id='FM_vis_joblog_enheder' name='FM_vis_joblog_enheder' value='1' type='checkbox' "& visjoblog_enheder_CHK &" />"
					end if
					
					if (thisfile = "fak_godkendt.asp" AND visjoblog_timepris = 1) OR thisfile = "erp_fak.asp" then
					Response.Write "<td width=100 align=right>"& chkboxesShow1 &"<b>TimePris</b></td>"
					end if
					
					if (thisfile = "fak_godkendt.asp" AND visjoblog_enheder = 1) OR thisfile = "erp_fak.asp" then
					Response.Write "<td width=100 align=right>"& chkboxesShow2 &"<b>Enheder</b></td>"
					end if
					
					
					
					Response.Write "<td align=right><b>Medarbejder</b></td></tr>"
					
					end if
					
					Response.write "<tr><td style='padding:2px 0px 2px 0px;'>"& formatdatetime(oRec("tdato"), 1) &"</td>"_
					&"<td style='padding:2px 0px 2px 0px;'>"& oRec("taktivitetnavn") &"</td> "
					Response.Write "<td align=right style='padding:2px 0px 2px 0px;'><b>"& formatnumber(oRec("timer"),2) &"</b></td>"
					
					if (thisfile = "fak_godkendt.asp" AND visjoblog_timepris = 1) OR thisfile = "erp_fak.asp" then
				    Response.Write "<td align=right style='padding:2px 0px 2px 0px;'>"& formatcurrency(oRec("timepris"),2) &"</td>"
					end if
					
					enheder = 0
					enheder = (oRec("faktor") * oRec("timer"))
					
					if len(trim(enheder)) <> 0 then
					enheder = enheder
					else
					enheder = 0
					end if
					
					if (thisfile = "fak_godkendt.asp" AND visjoblog_enheder = 1) OR thisfile = "erp_fak.asp" then
					Response.Write "<td align=right style='padding:2px 0px 2px 20px;'>"& formatnumber(enheder,2) &"</td>"
					end if
					
					Response.write "<td align=right style='padding:2px 0px 2px 20px;'><i>"& oRec("tmnavn") &"</i></td></tr>"
					
					if len(oRec("timerkom")) <> 0 then
					Response.write "<tr><td colspan=3 style='padding:0px 5px 5px 0px;' class=lille><i>"& oRec("timerkom") & "</i></td></tr>"
					end if
				
				x = x + 1
				lastjobnr = oRec("jobnr")  
				oRec.movenext
				wend
				oRec.close 
				
				if x = 0 then
				Response.write "<tr><td colspan=3><br><br><b><font class=roed>!</font> Ingen registreringer!</b></td></tr>"
				end if
				%>
				<!--</table>-->
									
    
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
            strSQL = "SELECT matid, matnavn, matantal AS matantal, matsalgspris, matenhed, matvarenr, forbrugsdato FROM materiale_forbrug WHERE "& strSQLjobaftKri &" AND forbrugsdato BETWEEN '"& stdatoKri &"' AND '"& slutdato &"' ORDER BY  forbrugsdato DESC"
            
            'Response.Write strSQL
           
            oRec.open strSQL, oConn, 3
            m = 0
            while not oRec.EOF
            %>
             <tr>
            
            <td style="padding:2px 0px 2px 0px;"><%=formatdatetime(oRec("forbrugsdato"), 1) %></td>
            <td style="padding:2px 0px 2px 0px;"><%=oRec("matnavn") %></td>
            <td style="padding:2px 0px 2px 0px;"><%=oRec("matvarenr") %></td>
            <td style="padding:2px 0px 2px 0px;" align=right><%=oRec("matantal") %></td>
            <td style="padding:2px 0px 2px 0px;" align=right><%=oRec("matsalgspris") %></td>
            <td style="padding:2px 0px 2px 0px;" align=right><%=oRec("matenhed") %></td>
            
              <td style="padding:2px 0px 2px 0px;" align=right>   
            <%
            matbelobLogTotal = oRec("matantal") * oRec("matsalgspris")
            %>
            <%=formatnumber(matbelobLogTotal, 2) %>
            </td>
           
        </tr>
           
            <%
           
            m = m + 1
            oRec.movenext
            wend
            oRec.close
            
            if m = 0 then
            Response.write "<tr><td colspan=7><br><br><b><font class=roed>!</font> Ingen registreringer!</b></td></tr>"
            end if
            
            
    end function
    
    
    function joblogdiv()
    %>
    
     <div id="joblogdiv" style="position:absolute; visibility:hidden; display:none; left:5px; top:105px; width:700px; border:1px yellowgreen solid; padding:10px 10px 10px 10px; background-color:#ffffff;">

	    <!-- joblog -->
		<%if visjoblog = 1 then
		visjoblogCHK = "CHECKED"
		else
		visjoblogCHK = ""
		end if 
		%>
		
		
		<table width=100% border=0 cellspacing=0 cellpadding=0>
		<tr>
		    <td><input id="FM_visjoblog" name="FM_visjoblog" <%=visjoblogCHK %> type="checkbox" /> <b>Vedhæft joblog på faktura.</b></td>
		</tr>
		<tr>
			<td><b>Joblog i valgt periode.</b> (Både fakturerbare og ikke-faktorerbare timer vises i jobloggen.) <br />
			<b>Periode afgrænsning:</b> <%=formatdatetime(showStDato, 1)%> - <%=formatdatetime(showSlutDato, 1)%>
			<br />
              
                
                &nbsp;</td>
		</tr>
		<tr>
		<td><div style="position:relative; visibility:visible; left:0px; top:0px; height:250px; border:2px YellowGreen dashed; padding:10px 10px 10px 10px; overflow:auto; background-color:#ffffe1; z-index:200;">
		
		<table border=0 cellspacing=0 cellpadding=0><%
		
		call joblog(jobid, stdatoKri, slutdato, aftid)
		
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
	    
	    <div id=joblogdiv_2 style="position:absolute; visibility:hidden; display:none; top:456px; width:600px; left:5px; border:0px #8cAAe6 solid;">
        <table width=600 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('<%=forr%>')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('<%=nst%>')" class=vmenu>Næste >></a></td></tr></table>
        </div>
	    
		<!-- Joblog slut -->
    
    <%
    end function
    
    
    function matlogdiv()
    %>
     <%if vismatlog = 1 then
       matLogChk = "CHECKED"
       else
       matLogChk = ""
       end if 
        %>
     <div id=matlogdiv style="position:absolute; width:700px; visibility:hidden; display:none; border:1px orange solid; top:105px; left:5px; padding:10px 10px 10px 10px; background-color:#ffffff;">
    	 <table width=100% cellspacing=0 cellpadding=0 border=0>
        <tr>
        <td valign=top>
            <input id="FM_vismatlog" name="FM_vismatlog" <%=matLogChk %> type="checkbox" value="1" /> <b>Vedhæft materiale log på faktura.</b>
            <br />
            <b>Materiale-log i valgt periode.</b><br />
            <b>Periode afgrænsning:</b> <%=formatdatetime(showStDato, 1)%> - <%=formatdatetime(showSlutDato, 1)%>
			
        </td>
        </tr>
        </table>
        <br />
        <table width=100% cellspacing=0 cellpadding=0 border=0>
        <tr>
        <td>
           <div style="position:relative; visibility:visible; height:250px; border:2px orange dashed; overflow:auto; background-color:#ffffe1; z-index:200; padding:10px 10px 10px 10px;">
            <table width=100% cellspacing=0 cellpadding=0 border=0>
            <tr>
            <td><b>Dato</b></td>
            <td><b>Navn</b></td>
            <td><b>Varenr</b></td>
            <td align=right><b>Antal</b></td>
            <td align=right><b>Enheds pris</b></td>
            <td align=right><b>Enhed</b></td>
            <td align=right><b>Pris ialt</b></td>
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
       
        <div id=matlogdiv_2 style="position:absolute; visibility:hidden; display:none; top:456px; width:600px; left:5px; border:0px #8cAAe6 solid;">
        <table width=600 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('aktdiv')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('<%=nst%>')" class=vmenu>Næste >></a></td></tr></table>
        </div>
    <%
    end function

%>