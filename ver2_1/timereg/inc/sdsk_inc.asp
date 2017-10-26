<%


public rspDagogTidUse
public tilfojTimerIalt
function response_dagogTid(rspDagogTid)


thisWeekday = weekday(rspDagogTid, 1)
call startogsluttider(thisWeekday)

'Response.Write "<br>useDayStTimeThis " & formatdatetime(useDayStTimeThis, 3) & "<br>"
'Response.Write "<br>"& cint(tilfojTimerIalt) &"<br>"

if formatdatetime(useDayStTimeThis, 3) <> "00:00:00" AND formatdatetime(useDayEndTimeThis, 3) <> "00:00:00" then


'Response.Write "Kalder func.<br>"
'Response.Write formatdatetime(rspDagogTid, 3) &" >= "& formatdatetime(useDayStTimeThis, 3) &" AND "& formatdatetime(rspDagogTid, 3) &" <= "& formatdatetime(useDayEndTimeThis, 3) &"  OR " & cint(tilfojTimerIalt) &" = 1<br>"
'OR cint(tilfojTimerIalt) = 1

'*** Slut tid indefor åbingstid
if (formatdatetime(rspDagogTid, 3) >= formatdatetime(useDayStTimeThis, 3) AND formatdatetime(rspDagogTid, 3) <= formatdatetime(useDayEndTimeThis, 3)) OR cint(tilfojTimerIalt) = 1 then
    if cint(tilfojTimerIalt) = 1  then
    rspDagogTidUse = formatdatetime(rspDagogTid, 2) &" "& left(formatdatetime(useDayStTimeThis, 3), 5)
    else
    rspDagogTidUse = formatdatetime(rspDagogTid, 2) &" "& left(formatdatetime(useDayStTimeThis, 3), 5) 'rspDagogTid
    end if
end if

'*** Slut tid er før åbingstid
if (formatdatetime(rspDagogTid, 3) < formatdatetime(useDayStTimeThis, 3)) AND cint(tilfojTimerIalt) = 0 then
    rspDagogTidUse = formatdatetime(rspDagogTid, 2) &" "& left(formatdatetime(useDayStTimeThis, 3), 5)
end if


'**** Slut tid er efter lukke tid ***
if (formatdatetime(rspDagogTid, 3) > formatdatetime(useDayEndTimeThis, 3)) AND cint(tilfojTimerIalt) = 0 then
    
    tilfojTimerIalt = 1
    rspDagogTid = dateadd("h", 24, formatdatetime(rspDagogTid, 2) &" "& formatdatetime(rspDagogTid, 3))
    
    call startogsluttider(rspDagogTid)
    call response_dagogTid(rspDagogTid)

end if

else
    
    tilfojTimerIalt = 1
    rspDagogTid = dateadd("h", 24, formatdatetime(rspDagogTid, 2) &" 01:15:00")
    call startogsluttider(rspDagogTid)
    call response_dagogTid(rspDagogTid)


end if
                   

'Lørdag og Søndag ' isworkday

end function



public useDayHuorsThis 
public useDayStTimeThis
public useDayEndTimeThis
function startogsluttider(thisWeekday)


select case weekday(thisWeekday, 1)

case 2
useDayHuorsThis = manTimer
useDayStTimeThis = normtid_st_man
useDayEndTimeThis = normtid_sl_man

case 3
useDayHuorsThis = tirTimer
useDayStTimeThis = normtid_st_tir
useDayEndTimeThis = normtid_sl_tir

case 4
useDayHuorsThis = onsTimer
useDayStTimeThis = normtid_st_ons
useDayEndTimeThis = normtid_sl_ons

case 5
useDayHuorsThis = torTimer
useDayStTimeThis = normtid_st_tor
useDayEndTimeThis = normtid_sl_tor

case 6
useDayHuorsThis = freTimer
useDayStTimeThis = normtid_st_fre
useDayEndTimeThis = normtid_sl_fre

case 7
useDayHuorsThis = lorTimer
useDayStTimeThis = normtid_st_lor
useDayEndTimeThis = normtid_sl_lor

case 1
useDayHuorsThis = sonTimer
useDayStTimeThis = normtid_st_son
useDayEndTimeThis = normtid_sl_son
end select


end function

public bcc

public modNavn, modEmail, ansvNavn, ansvEmail
function eml_inciAns()

            '** Henter Modtager (Incident ansvarlig) ***
			strSQL = "SELECT mnavn, mnr, init, email FROM medarbejdere WHERE mid = " & intAns
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
				
			
			
			if instr(oRec("email"), "@") <> 0 then
			ansvNavn = oRec("mnavn") & "("& oRec("mnr") &") - " & oRec("init")
			ansvEmail = oRec("email")
			bcc = bcc & "Incident ansvarlig: "&oRec("mnavn")&", "& oRec("email") & "<br>"
		    end if
				
			oRec.movenext
			wend
				
			oRec.close


end function

function eml_inciCreator()

            '** Henter Modtager (Incident ansvarlig) ***'
            if len(creator) <> 0 then
            creator = creator
            else
            creator = 0
            end if
            
			strSQL = "SELECT mnavn, mnr, init, email FROM medarbejdere WHERE mid = " & creator
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
				
			
			
			if instr(oRec("email"), "@") <> 0 then
			modNavn = oRec("mnavn") & "("& oRec("mnr") &")  " & oRec("init")
			modEmail = oRec("email")
			'myMail.BCC = ""&modNavn&""&""& modEmail
            myMail.CC = myMail.CC & ";"& modEmail
			bcc = bcc & "Incident ejer: "&oRec("mnavn")&", "& oRec("email") & "<br>"
		    end if 
				
			oRec.movenext
			wend
				
			oRec.close


end function



public knavn, knr, adresse, postnr, by, telefon, modKansNavn, modKansEmail, kundeans1
function eml_kogjobAns()

                '*** Sender til kundeansvarlig ****
                '** Henter Kundeans ***

                kundeans1 = 0
		        strSQL = "SELECT m.mnavn, m.mnr, m.init, m.email, k.kundeans1, kkundenavn, kkundenr, adresse, postnr, city, telefon FROM kunder k "_
		        &" LEFT JOIN medarbejdere m ON (m.mid = k.kundeans1) WHERE k.kid = "& intKontakt 

		        'Response.write strSQL
		        'Response.flush
        		
		        oRec.open strSQL, oConn, 3
		        if not oRec.EOF then
        		
		        knavn = oRec("kkundenavn")
		        knr = oRec("kkundenr")	
		        adresse = oRec("adresse")
		        postnr = oRec("postnr")
		        by = oRec("city")
		        telefon = oRec("telefon")
		        modKansNavn = oRec("mnavn")
		        modKansEmail = oRec("email")
		        kundeans1 = oRec("kundeans1")
        		
        		if len(request("FM_sendtilkogjans")) <> 0 then
		        
        		    if instr(oRec("email"), "@") <> 0 then
                    'myMail.Bcc = ""&modKansNavn&""&""& modKansEmail
                    myMail.CC = myMail.CC & ";"& modKansEmail
                    bcc = bcc & "Pri. Kundeansv: "&oRec("mnavn")&", "& oRec("email") & "<br>"
                    end if
        		
        		end if
        		
		        end if
		        oRec.close
        		
        		
        if len(request("FM_sendtilkogjans")) <> 0 then
		        
		        '*** Sender til jobansvarlig(e) ****
		        strSQL = "SELECT m.mnavn, m.email FROM job j "_
		        &" LEFT JOIN medarbejdere m ON (m.mid = j.jobans1 OR m.mid = j.jobans2) WHERE j.id = "& jobid

		        'Response.write strSQL
		        'Response.flush
        		
		        oRec.open strSQL, oConn, 3
		        while not oRec.EOF 
        		
        		 if instr(oRec("email"), "@") <> 0 then
		         'myMail.Bcc = ""&oRec("mnavn")&""&""& oRec("email")
                 myMail.CC = myMail.CC & ";"& oRec("email")
        		 bcc = bcc & "Jobansvarlig: "&oRec("mnavn")&", "& oRec("email") & "<br>"
                 end if
        		
		        oRec.movenext
		        wend 
		        oRec.close

        end if

end function


function eml_Kpers(kpersid)


    '*** Tilføjer kontaktpersoner hos kontakt til mail advisering ***
	if len(request("FM_sendtilkontakter")) <> 0 then
	
		strSQL = "SELECT navn, email FROM kontaktpers WHERE kundeid = "& intKontakt & " AND id = "& kpersid
		oRec.open strSQL, oConn, 3
		while not oRec.EOF
			
		if instr(oRec("email"), "@") <> 0 then
		'myMail.Bcc = ""&oRec("navn")&""&""& oRec("email")
        myMail.CC = myMail.CC & ";"& oRec("email")  
		bcc = bcc & "Kontaktperson: "&oRec("navn")&", "& oRec("email") & "<br>"
	    end if
		
		oRec.movenext
		wend
		oRec.close
	
	
	end if
					

end function




function eml_proGrp()

            '*** Adviser projektgruppe ***	
			'if len(request("FM_sendtilprogrp")) <> 0 then
			    
				strSQL = "SELECT id, progrp FROM sdsk_status WHERE id = "& intStatus
				'Response.Write strSQL & "<br>"
				
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
					
					
					
					strSQL2 = "SELECT p.id, p.navn, m.mnavn, m.email FROM projektgrupper p"_
					&" LEFT JOIN progrupperelationer pr ON (pr.projektgruppeId = p.id) "_
					&" LEFT JOIN medarbejdere m ON (m.mid = pr.medarbejderId AND mansat <> '2') WHERE p.id = " & oRec("progrp")
					
					'Response.Write strSQL2
					
					oRec2.open strSQL2, oConn, 3
					while not oRec2.EOF
					
					
					
				        
						if instr(oRec2("email"), "@") <> 0 then 'AND instr(bcc, oRec2("email")) = 0 then
						'myMail.Bcc = ""&oRec2("mnavn")&""&""& oRec2("email") 
                        myMail.CC = myMail.CC & ";"& oRec2("email")
						bcc = bcc & "Projektgruppe medlem: "&oRec2("mnavn")&", "& oRec2("email") & "<br>" 
						end if
						
					oRec2.movenext
					wend
					oRec2.close
				
				
				end if 
				oRec.close
			
			'end if


end function


function tilfojtiljob()%>
		
		<h4>Tilføj Incident til job.</h4>
		<%if func = "red" then%>
		<font color="red"><b>Tilknytter kun Incident til job, gemmer ikke øvrige ændringer 
		på Incident.</b><br /></font>
		<%end if %>
		Bruges til intern timeregistrering, stopur og faktor beregning.<br>
		<table cellspacing=1 cellpadding=2 border=0>
		<tr>
		<%if func = "red" then%>
		</form>
		<%end if %>
		<form action="sdsk.asp?id=<%=lastedit%>&func=dbopr3&kid=<%=kid%>&bcc=<%=bcc%>" method="POST">
		
		<td colspan=2>
		<br><a href="javascript:popUp('jobs.asp?showaspopup=y&func=opret&id=0&int=1&rdir=sdsk&dato=<%=dato%>&datostkri=<%=jobstKri%>&kid=<%=kid%>&FM_kunde=<%=kontaktId%>','600','650','250','30');" target="_self" class=vmenu>Opret nyt job (Ekspress)..</a><br />
		<a href="job_kopier.asp?func=kopier&id=&rdir=sdsk&showaspopup=y" target="_blank" class=vmenu>Kopier eksisterende job..</a>
	
		<!--&datoslkri=<=jobslKri%>&stepialt=<=antalstepialt%> --></td>			
	</tr>
	<tr>
		<td width=100><br><img src="../ill/ac0063-16.gif" width="16" height="16" alt="" border="0">&nbsp;<b>Vælg job:</b>&nbsp;&nbsp;</td>
		<td><br>
		<%
		strSQL = "SELECT jobnavn, jobnr, id, jobknr, kkundenavn, kkundenr, kid FROM kunder "_
		&"LEFT JOIN job ON (jobstatus = 1 AND fakturerbart = 1 AND jobknr = kid) WHERE kid = "& kontaktId &" ORDER BY kkundenavn, jobnavn"
		'Response.write strSQL
		%>
		<select name="FM_jobid" id="FM_jobid" style="width:278px; font-size:10px;">
		<option value="0">Ingen (Tilføj ikke Incident til job)</option>
			<%	oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
			%>
			<option value="<%=oRec("id")%>"><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>) </option>
			<%
			oRec.movenext
			wend
			oRec.close %>
			
		</select>&nbsp;
		</td>
	</tr>
	<tr>
		
		<td><b>Forretningsområde:</b></td>
		<td><select name="FM_fomr" style="width:278px; font-size:10px;">
		<option value="0">Ingen</option>
		<%
		strSQL = "SELECT id, navn FROM fomr ORDER BY navn"
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		%>
		<option value="<%=oRec("id")%>"><%=oRec("navn")%></option>
		<%
		oRec.movenext
		wend
		oRec.close 
		%>
		</select></td>
	</tr>
	
	<%if request.Cookies("sdsk")("opretakt") = "1" OR lto = "kits" then
	aktCHK = "CHECKED"
	else
	aktCHK = ""
	end if %>
	
	<tr><td><b>Opret tilhørende aktivitet på job?</b></td><td><input type="checkbox" name="FM_opretakt" id="FM_opretakt" value="1" <%=aktCHK %>> Ja</td></tr>
	<tr>
		<td colspan=2 align=center><br><input type="submit" value="Tilføj til job >>"></td>
		<%if func <> "red" then %>
	    </form>
	    <%end if %>
	</tr>
	</table>
	<%end function%>

