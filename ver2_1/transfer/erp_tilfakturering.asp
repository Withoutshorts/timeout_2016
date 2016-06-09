<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato2.asp"-->
<!--#include file="../inc/regular/erp_func.asp"-->

<%

if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	menu = "erp"
	func = request("func")
    thisfile = "erp_tilfakturering.asp"
    print = request("print")
	
	
	if print <> "j" then
	
	dTop = "132"
	dLeft = "20" 
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call erpmainmenu(1)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call erptopmenu()
	%>
	</div>
	
	<%else 
	
	dTop = "20"
	dLeft = "20" 
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	
	<%end if %>
	
	
	<div id="sindhold" style="position:absolute; left:<%=dLeft %>px; top:<%=dTop %>px; visibility:visible;">
	<%
	'**********************************************************'
	'**************** Job til fakturering *********************'
	'**********************************************************'
	
	%>
	<h3><img src="../ill/ac0010-24.gif" /> Til fakturering</h3>
	<%
	
	call filterheader(0,0,900,pTxt)
	
	%>
	
	<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<form action="erp_tilfakturering.asp?FM_usedatokri=1" method="POST">
	<tr>
	<td valign=top width=300>
	
    <% 
    if len(request("FM_kunde")) <> 0 then
			
			if request("FM_kunde") = 0 then
			valgtKunde = 0
			sqlKundeKri = "t.tknr <> 0"
			sqlKundeKri2 = "k.kid <> 0"
			else
			valgtKunde = request("FM_kunde")
			sqlKundeKri = "t.tknr = "& valgtKunde &""
			sqlKundeKri2 = "k.kid = "& valgtKunde &""
			end if
			
	else
		
			if len(request.cookies("erp")("kid")) <> 0 AND request.cookies("erp")("kid") <> 0 then
			valgtKunde = request.cookies("erp")("kid")
			sqlKundeKri = "t.tknr = "& valgtKunde &""
			sqlKundeKri2 = "k.kid = "& valgtKunde &""
			else
			valgtKunde = 0
			sqlKundeKri = "t.tknr <> 0"
			sqlKundeKri2 = "k.kid <> 0"
			end if
			
	end if
	
	
	'*** Hvis er er søgt på jobnr ***'
	if len(trim(request("FM_sog"))) <> 0 then
	        
	        %>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(request("FM_sog")) 
			if isInt > 0 OR instr(request("FM_sog"), ".") <> 0 then
		    isInt = 0
		    sogval = 0
		    else
            sogval = request("FM_sog")
            end if
            
	jobnrKri = " AND j.jobnr = "& sogval
	
	strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM job j, kunder k WHERE j.jobnr = "& sogval &" AND k.kid = j.jobknr"
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	valgtKunde = oRec("kid")
	
	oRec.movenext
	wend
	oRec.close
	
	sqlKundeKri = "t.tknr = "& valgtKunde &""
	sqlKundeKri2 = "k.kid = "& valgtKunde &""
	
	else
	sogval = ""
	jobnrKri = ""
	end if
	
	
	
	
    
    %>
    <b>Kontakter:</b> <br>
        <%if print <> "j" then %>
        <select name="FM_kunde" size="1" style="font-size : 9px; width:255px;">
		<option value="0">Alle</option>
		<%end if
		ketypeKri = " ketype <> 'e'"
		strKundeKri = " AND kid <> 0 "
		vlgtKunde = "Alle"
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" "& strKundeKri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(valgtKunde) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				vlgtKunde = oRec("Kkundenavn") & " ("& oRec("Kkundenr") &")"
				else
				isSelected = ""
				end if
				
				
				if print <> "j" then
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr") %>)</option>
				<%
				end if
				oRec.movenext
				wend
				oRec.close
				
				if print <> "j" then%>
		        </select><br>
		        <% 
		        else
		        %>
				- <%=vlgtKunde %>
				<% 
				end if
				%>
	    
	    <!--
	    <b>Kontaktansvarlige (medarbejder):</b><br> <select name="FM_kans" size="1" style="font-size : 9px; width:285px;">
		<option value="0">Alle</option>
	    </select>-->
	    
	    
	        <%
	        if level <= 2 then
    		
    		
		        if len(trim(request("FM_ignorer_pg"))) <> 0 then
		        ign_progrp = request("FM_ignorer_pg")
		                if ign_progrp = 1 then
		                 chkThis1 = "CHECKED"
		                chkThis0 = ""
		                 else
		                 chkThis1 = ""
		                 chkThis0 = "CHECKED"
		                 end if       
		        else
		            if request.cookies("erp")("projgrp") <> "" then
		            ign_progrp = request.cookies("erp")("projgrp")
    		            
		                 if ign_progrp = 1 then
		                 chkThis1 = "CHECKED"
		                chkThis0 = ""
		                 else
		                 chkThis1 = ""
		                 chkThis0 = "CHECKED"
		                 end if       
    		                
		            else
		            ign_progrp = 0
		             chkThis1 = ""
		             chkThis0 = "CHECKED"
		            end if
		         
		        end if
		 
		     else
		        
		        ign_progrp = 0
		         chkThis1 = ""
		         chkThis0 = "CHECKED"
		    
		     
		    end if
		    
		    Response.Cookies("erp")("projgrp") = ign_progrp
    		
		    if print <> "j" then%>
		    <br /><b>Ignorer tilknytning til job via Projektgrupper:</b><br /><input type="radio" name="FM_ignorer_pg" value="1" <%=chkThis1%>> Ja - vis alle job.<br />
		    <input type="radio" name="FM_ignorer_pg" value="0" <%=chkThis0%>> Nej - vis kun job du er tilknyttet via dine projektgrupper.
		    
		    <br /><br />
		    
		    <b>Søg på job nr.:</b> <input id="FM_sog" name="FM_sog" value="<%=sogval %>" type="text" style="width:100px;" /><br />
		    (Ignorerer <b>IKKE</b> de andre søgekriterier)
		    
		    <%else %>
		    <%if sogval <> "" then %>
		        <br /><br />Der er søgt på jobnr: <b><%=sogval %></b><br />
		    <%end if %>
            <%end if%>
        
        
       
        <%if print = "j" AND ign_progrp = 1 then%>
        <br />- Vis alle job, uanset om du er tilknyttet via dine projektgrupper.
        <%end if%>
        
        
        <%if len(trim(request("FM_visjob"))) <> 0 then 
        visJobChk = "CHECKED"
        visikkejob = 1
        else
        visJobChk = ""
        visikkejob = 0
        end if%>
        
        <%if request("FM_nuljob") = 0 then 
        nulJobChk0 = "CHECKED"
        nulJobChk1 = ""
        visikkenuljob = 1
        else
        nulJobChk0 = ""
        nulJobChk1 = "CHECKED"
        visikkenuljob = 0
        end if%>
        
         <%if len(trim(request("FM_jobstatus"))) <> 0 then 
        jobStatChk = "CHECKED"
        visikkejobstatus = 1
        else
        jobStatChk = ""
        visikkejobstatus = 0
        end if%>
        
        <%
        '** Vis aft. cookie / value ********
        if len(trim(request("FM_usedatokri"))) <> 0 then 
            if request("FM_visaft") = 1 then
            visAftChk = "CHECKED"
            visikkeaft = 1
            else
            visAftChk = ""
            visikkeaft = 0
            end if
        else
            if request.Cookies("erp")("visaft") = "1" then
            visAftChk = "CHECKED"
            visikkeaft = 1
            else
            visAftChk = ""
            visikkeaft = 0
            end if
        end if
        
        
        %>
        
        
        <%
        if len(trim(request("FM_medarb_jobans"))) <> 0 then
        medarb_jobans = request("FM_medarb_jobans")
        else
                if request.Cookies("erp")("jk_ans") <> 0 then
                medarb_jobans = request.Cookies("erp")("jk_ans")
                else
                medarb_jobans = 0
                end if
        end if
        %>
         
        
        <% 
        
        if len(trim(request("job_kans"))) = 0 then
                
                if request.Cookies("erp")("jk_ans_filter") <> 0 then
                    
                    if request.Cookies("erp")("jk_ans_filter") = "1" then
                    job_kans1CHK = "CHECKED"
		            job_kans2CHK = ""
		            job_kans = "1"
		            else
		            job_kans1CHK = ""
		            job_kans2CHK = "CHECKED"
		            job_kans = "2"
		            end if
		        
                else
                    job_kans1CHK = "CHECKED"
		            job_kans2CHK = ""
		            job_kans = "1"
                end if
                
        else
        
            if request("job_kans") = "1" then
		        job_kans1CHK = "CHECKED"
		        job_kans2CHK = ""
		        job_kans = "1"
		    else
		        job_kans1CHK = ""
		        job_kans2CHK = "CHECKED"
		        job_kans = "2"
		    end if
		
		end if
		
		%>
        
        </td><td>
        <b>Job / kundeansvarlig:</b><br />
        Vis kun job og aftaler hvor
        
        <%if print <> "j" then%>
        <select name="FM_medarb_jobans" id="FM_medarb_jobans" style="font-size:9px; width:203px;">
            <option value="0">Alle - ignorer jobansv.</option>
            <%
            mNavn = "Alle"
            
             strSQL = "SELECT mnavn, mnr, mid, init FROM medarbejdere WHERE mansat <> '2' ORDER BY mnavn"
             oRec.open strSQL, oConn, 3
             while not oRec.EOF
                
                if cint(medarb_jobans) = oRec("mid") then
                selThis = "SELECTED"
                        
                        '** medarb navn til print **
                        mNavn = oRec("mnavn") & " ("& oRec("mnr") &") "
                       
                        if len(trim(oRec("init"))) <> 0 then 
                        mNavn = mNavn & " - "& oRec("init") 
                        end if 
                        
                else
                selThis = ""
                end if
                %>
             <option value="<%=oRec("mid")%>" <%=selThis%>><%=oRec("mnavn") %> (<%=oRec("mnr") %>)
             <%if len(trim(oRec("init"))) <> 0 then %>
              - <%=oRec("init") %>
              <%end if %>
              </option>
             <%
             oRec.movenext
             wend
             oRec.close
             %>
             
              
        </select>
        <br />er 
        <input name="job_kans" id="job_kans" value="1" type="radio" <%=job_kans1CHK%> /> jobansvarlig  <input name="job_kans" id="job_kans" value="2" type="radio" <%=job_kans2CHK%> /> kundeansvarlig 
        <%else %>
            
            <b><%=request("mNavn")%></b> er
            <%if job_kans = 1 then %>
            jobansvarlig.
            <%else %>
            kundeansvarlig.
            <%end if %>
        
        <%end if %>
        
        <%
        response.Cookies("erp")("visaft") = visikkeaft
        response.Cookies("erp")("jk_ans") = medarb_jobans
        response.Cookies("erp")("jk_ans_filter") = job_kans
        response.Cookies("erp")("kid") = valgtKunde
	    response.Cookies("erp").expires = date + 60
        %>
        
        <br /> <br><b>Job / Aftale filter:</b>
        
        <%if print <> "j" then%>
        
	    <br /><input id="FM_visjob" name="FM_visjob" type="checkbox" value=1 <%=visJobChk %> /> Skjul job
	    <input id="FM_nuljob0" name="FM_nuljob" type="radio" value=0 <%=nulJobChk0 %> />  Skjul "nul" job <input id="FM_nuljob1" name="FM_nuljob" type="radio" value=1 <%=nulJobChk1 %> /> Vis "nul" job <br />
        <input id="FM_jobstatus" name="FM_jobstatus" type="checkbox" value=1 <%=jobStatChk %> /> Vis lukkede og passive job<br> 
        <input id="FM_visaft" name="FM_visaft" type="checkbox" value=1 <%=visAftChk %>/> Skjul aftaler (aftaler på viste job skjules ikke)
        
        <%else %>
        
        <%if visikkejob = 1 then%> 
        <br />- Skjul job 
        <%end if %>
        
        <%if visikkeaft = 1 then%> 
        <br />- Skjul aftaler
        <%end if %>
        
        <%if visikkenuljob = 1 then %>
        <br />- Skjul "nul" job
        <%end if %>
        
        
        <%if visikkejobstatus = 1 then %>
        <br />- Vis lukkede og passive job
        <%end if %>
        
        
        <%end if %>
        </td>
        
        <%  
	    if ign_progrp <> "1" then
		call hentbgrppamedarb(session("mid"))
	    else
		strPgrpSQLkri = ""
	    end if
	    %>
        
        <% if print <> "j" then%>
        <td style="padding:2px 2px 2px 20px;" valign=top>
	    <b>Periode:</b><br /> 
	    <!--#include file="inc/weekselector_s.asp"-->&nbsp;<br />
	    <br /> <input id="Submit1" type="submit" value="  Vis timer til fak.  " />
        <br /><br />
	    Hvis der findes en faktura på de pågældende job i det valgte dato-interval medreges kun timer siden sidste faktura.
	    </td>
	    <%else
	    dontshowDD = 1
	    %>
	     <!--#include file="inc/weekselector_s.asp"-->
	    &nbsp;<%=chkNameValue %>
	    <%end if %>
	 
	 
	 </tr></form></table>
	 
	 </td>
	 </tr></table>
	 </div>
	<%
	
	 
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag  'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut 'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	
	'Response.Write sqlDatostart 
	'Response.Flush
	'Response.end
	
	%>
	<br><b>Periode:</b>&nbsp;
	<%=formatdatetime(strDag&"/"& strMrd &"/"& strAar, 1) & " - " & formatdatetime(strDag_slut &"/"& strMrd_slut &"/"& strAar_slut, 1)%>
	
	<%if print <> "j" then%>
	&nbsp;&nbsp;|&nbsp;&nbsp;<a href="erp_tilfakturering.asp?mNavn=<%=mNavn%>&FM_medarb_jobans=<%=medarb_jobans%>&job_kans=<%=job_kans%>&print=j&FM_jobstatus=<%=request("FM_jobstatus")%>&FM_nuljob=<%=request("FM_nuljob")%>&FM_ignorer_pg=<%=ign_progrp%>&FM_visjob=<%=request("FM_visjob")%>&FM_visaft=<%=request("FM_visaft")%>&FM_sog=<%=sogval%>" class=vmenu target=blank>Print venlig version</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="erp_opr_faktura_fs.asp" class=vmenu>Opret ny faktura skrivelse..</a>
	<%end if%>
	
	<br />
        &nbsp;
	<table cellspacing=0 cellpadding=0 border=0 width=95%>
	<% 
	
	dim kid
	dim knavn
	dim knr
	dim jobnavn
	dim jobid
	dim jobnr
	dim jobans1
	dim jobans2
	dim aftnavn
	dim aftnr
	dim varLastfak
	dim varFakfindes
	dim varTimer, varTimerFak, varTimerIG
	dim varAftOrJob
	dim aftId
	dim varMat
	dim varEnh
	dim varShadowCopy
	dim aftType
	dim varAftMat, jobstartdato, jobslutdato, jobtype,  aftStdato, aftSldato, jobstatus, aftStatus
	dim kans1, kans2, ventetimer
	dim lastFakuseStDato
	
	x = 2000
	redim kid(x)
	redim knavn(x)
	redim knr(x)
	redim jobid(x)
	redim jobnavn(x)
	redim jobnr(x)
	redim jobans1(x)
	redim jobans2(x)
	redim aftId(x)
	redim aftnavn(x)
	redim aftnr(x)
	redim varLastfak(x)
	redim varFakfindes(x)
	redim varTimer(x), varTimerFak(x), varTimerIG(x)
	redim varAftOrJob(x)
	redim varMat(x)
	redim varEnh(x)
	redim varShadowCopy(x) 
	redim aftType(x)
	redim varAftMat(x)
	redim jobstartdato(x), jobslutdato(x), jobtype(x), aftStdato(x), aftSldato(x), jobstatus(x), aftStatus(x)
	redim kans1(x), kans2(x), ventetimer(x)
	redim lastFakuseStDato(x)
	
	
	if visikkejobstatus = 1 then 'Via alle job uanset status
	jobStatusKri = " j.jobstatus <> 99 "
	else
	jobStatusKri = " j.jobstatus = 1 "
	end if
	
	if medarb_jobans <> 0 then
	    if job_kans = 1 then
	    kansKri = ""
	    jobansKri = " AND (j.jobans1 = "& medarb_jobans &" OR j.jobans2 = "& medarb_jobans &") "
	    else
	    kansKri = " AND (kundeans1 = "& medarb_jobans &" OR kundeans2 = "& medarb_jobans &")"
	    jobansKri = ""
	    end if
	else
	kansKri = ""
	jobansKri = ""
	end if
	
	
	lastKid = 0
	
	strSQL = "SELECT j.id, jobnavn, jobnr, jobknr, kkundenavn, jobans2, "_
	&" kkundenr, jobslutdato, jobstartdato, fastpris, "_
	&" j.beskrivelse, jobans1, "_
	&" kid, j.serviceaft, j.jobstatus, s.id AS aid, s.navn AS aftnavn, s.aftalenr, s.advitype, s.stdato, s.sldato, s.status, "_
	&" m.mnavn AS mnavn, m.mnr AS mnr, m2.mnavn AS m2navn, m2.mnr AS m2nr, m.init AS minit, m2.init AS m2init, "_
	&" m3.mnavn AS m3navn, m3.mnr AS m3nr, m4.mnavn AS m4navn, m4.mnr AS m4nr, m3.init AS m3init, m4.init AS m4init "_
	&" FROM kunder k "_
    &" LEFT JOIN job j ON (j.jobknr = k.kid AND "& jobStatusKri &" "& strPgrpSQLkri &" "& jobansKri &" "& jobnrKri &") "_
    &" LEFT JOIN serviceaft s ON (s.id = j.serviceaft) "_
    &" LEFT JOIN medarbejdere m ON (m.mid = j.jobans1)"_
    &" LEFT JOIN medarbejdere m2 ON (m2.mid = j.jobans2)"_
    &" LEFT JOIN medarbejdere m3 ON (m3.mid = k.kundeans1)"_
    &" LEFT JOIN medarbejdere m4 ON (m4.mid = k.kundeans2)"_
    &" WHERE  "& sqlKundeKri2 &" "& kansKri &""_ 
	&" ORDER BY kkundenavn, kkundenr, jobnavn, jobslutdato"  
	
	'Response.Write strSQL & "<br><br>"
	'Response.flush
	
	x = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	
	
	        '*** Henter aftaler ****
	        if oRec("kid") <> lastKid then 
	       
	        strSQLaft = "SELECT s.id AS aid, s.navn AS aftnavn, s.aftalenr, s.advitype, s.stdato, s.sldato, s.status "_
	        &" FROM serviceaft s WHERE s.kundeid = " & oRec("kid") & " AND s.status = 1 ORDER BY s.navn " 
	        oRec2.open strSQLaft, oConn, 3
	        while not oRec2.EOF
	        
	        kid(x) = oRec("kid") 
	        knavn(x) = oRec("kkundenavn") 
	        knr(x) = oRec("kkundenr")
	        jobid(x) = 0
	        
	        
	        if len(oRec("m3navn")) <> 0 then
	        kans1(x) = oRec("m3navn")&" ("& oRec("m3nr") &") - " & oRec("m3init")
	        else
	        kans1(x) = ""
	        end if
            
            if len(oRec("m4navn")) <> 0 then
            kans2(x) = oRec("m4navn")&" ("& oRec("m4nr") &") - " & oRec("m4init") 
            else
            kans2(x) = ""
            end if
	        
	        jobnavn(x) = ""
	        jobnr(x) = "" 
	        
	        jobans1(x) = "" 
	        jobans2(x) = "" 
        	
	        aftId(x) = oRec2("aid")
	        aftnavn(x) = oRec2("aftnavn")
	        aftnr(x) = oRec2("aftalenr")
	        
	        if oRec2("advitype") <> 2 then
	        aftType(x) = "Enh./Klip"
	        else
	        aftType(x) = "Periode"
	        end if
	        
	        aftStdato(x) = formatdatetime(oRec2("stdato"), 2) 
	        aftSldato(x) = formatdatetime(oRec2("sldato"), 2)            
	        
	        select case oRec2("status")
	        case 0
	        aftStatus(x) = "<font color=red><i>Lukket</i></font>"    
	        case 1
	        aftStatus(x) = "<font color=yellowgreen><i>Aktiv</i></font>"    
	        end select
	                
	                    
	                    '** Fakturaer på aftale ***
	                   
		                lastFak = "01/01/01"
		                aftfakfindes = 0
		                lastFaknr = 0
                		
		                varLastfak(x) = ""
	                    varFakfindes(x) = fakfindes
                		
		                
		                        strSQLFak = "SELECT f.fakdato, f.faknr FROM fakturaer f "_
		                        &" WHERE (f.aftaleid = "& oRec2("aid") &" AND "_
		                        &" f.fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"') "_
		                        &" ORDER BY f.fakdato DESC"
                        		
		                        'Response.Write strSQLFak &"<br>" 
                        		
		                        oRec3.open strSQLFak, oConn, 3
                                if not oRec3.EOF then
                                    
                                    aftfakfindes = 1
                                    lastaftFak = oRec3("fakdato")
                                    lastaftFaknr = oRec3("faknr")
                                    
                                    varLastfak(x) = lastaftFak &" ("& lastaftFaknr &")"
	                                varFakfindes(x) = aftfakfindes
                                    
                                end if
                                oRec3.close
                                
                                
                                
                                
                                    if aftfakfindes = 1 then
                
                                    lastaftFakuseSQL = dateadd("d", 1, lastaftFak)
                                    lastaftFakuseSQL = year(lastaftFakuseSQL) &"/"& month(lastaftFakuseSQL) &"/"& day(lastaftFakuseSQL)
                                    
                                    else
                                    
                                    lastaftFakuseSQL = sqlDatostart
                                    
                                    end if
                                    
                                   
                                    '** Enheder 
                                    
                                    sumEnhVedFak = 0
                                    strSQLenh = "SELECT sum(timer * faktor) AS sumEnhEfterFak FROM timer "_
                                    &" LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE "_
                                    &" (tfaktim = 1 OR tfaktim = 2 OR tfaktim = 6) AND seraft = "& oRec2("aid") &""_ 
	                                &" AND tdato BETWEEN '"& lastaftFakuseSQL &"' AND '"& sqlDatoslut &"' GROUP BY seraft, taktivitetid"
                    	            
	                                'Response.Write strSQLenh &"<br>"
                    	            
	                                oRec3.open strSQLenh, oConn, 3
                                    while not oRec3.EOF
                                    
                                    sumEnhVedFak = sumEnhVedFak + oRec3("sumEnhEfterFak")
                                    
                                    oRec3.movenext
                                    wend
                                    oRec3.close
                                    
                                    
                                    '*** Materialer
                
                                    matAftAntalEfterFak = 0
                                    strSQLaftMAT = "SELECT sum(matantal) AS matAftAntalEfterFak FROM materiale_forbrug WHERE "_
                                    &" serviceaft = "& oRec2("aid") &""_ 
	                                &" AND forbrugsdato BETWEEN '"& lastaftFakuseSQL &"' AND '"& sqlDatoslut &"' GROUP BY serviceaft"
                    	            
	                               'Response.Write strSQLaftMAT &"<br>"
                    	            
	                                oRec3.open strSQLaftMAT, oConn, 3
                                    while not oRec3.EOF
                                    
                                    matAftAntalEfterFak = oRec3("matAftAntalEfterFak")
                                    
                                    oRec3.movenext
                                    wend
                                    oRec3.close
	                    
        	
	        
	       
	        varAftOrJob(x) = 1  'aftale
	        varAftMat(x) = matAftAntalEfterFak 
	        varEnh(x) = sumEnhVedFak
	        
	        x = x + 1
	        oRec2.movenext
	        wend
	        oRec2.close
        	
        	
        	end if
	
	    
	    
	    '*** Job inf ***
	
        kid(x) = oRec("kid") 
        knavn(x) = oRec("kkundenavn") 
        knr(x) = oRec("kkundenr")
        
        jobid(x) = oRec("id")
        jobnavn(x) = oRec("jobnavn")
        jobnr(x) = oRec("jobnr")
        
        select case oRec("jobstatus")
        case 0
        jobstatus(x) = "<font color=red>Lukket</font>"
        case 1
        jobstatus(x) = "<font color=yellowgreen>Aktiv</font>"
        case 2
        jobstatus(x) = "<font color=silver>Passiv</font>"
        end select
        
            if len(oRec("m3navn")) <> 0 then
	        kans1(x) = oRec("m3navn")&" ("& oRec("m3nr") &") - " & oRec("m3init")
	        else
	        kans1(x) = ""
	        end if
            
            if len(oRec("m4navn")) <> 0 then
            kans2(x) = oRec("m4navn")&" ("& oRec("m4nr") &") - " & oRec("m4init") 
            else
            kans2(x) = ""
            end if
        
        if len(trim(oRec("jobstartdato"))) <> 0 AND len(trim(oRec("jobslutdato"))) <> 0 then
        jobstartdato(x) = formatdatetime(oRec("jobstartdato"), 2)
        jobslutdato(x) = formatdatetime(oRec("jobslutdato"), 2)
        else
        jobstartdato(x) = oRec("jobstartdato")
        jobslutdato(x) = oRec("jobslutdato")
        end if
        
        jobtype(x) = oRec("fastpris")
        
        if len(trim(oRec("mnavn"))) <> 0 then
        jobans1(x) = left(oRec("mnavn"), 7) &" ("& oRec("mnr") &")" &" - "& oRec("minit")
        else
        jobans1(x) = ""
        end if
        
        if len(trim(oRec("m2navn"))) <> 0 then
        jobans2(x) = left(oRec("m2navn"), 7) &" ("& oRec("m2nr") &")" &" - "& oRec("m2init") 
        else
        jobans2(x) = ""
        end if
    	
       
        varAftOrJob(x) = 0 'job 
        varEnh(x) = ""
	    varAftMat(x) = ""
	
	
        	
		       
        		
		if len(trim(oRec("id"))) <> 0 then
		
		
		         '***** Findes der fakturaer på job?? ***
		        lastFak = "01/01/01"
		        fakfindes = 0
		        lastFaknr = 0
        		
		        varLastfak(x) = ""
	            varFakfindes(x) = fakfindes
		
		        strSQLFak = "SELECT f.fakdato, f.faknr, f.shadowcopy FROM fakturaer f "_
		        &" WHERE (f.jobid = "& oRec("id") &" AND f.fakdato >= '"& sqlDatostart &"')"_
		        &" ORDER BY f.fakdato DESC"
        		
        		' &" f.fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"') "_
		        'Response.Write strSQLFak &"<br>" 
        		
		        oRec3.open strSQLFak, oConn, 3
                if not oRec3.EOF then
                    
                    fakfindes = 1
                    lastFak = oRec3("fakdato")
                    lastFaknr = oRec3("faknr")
                    
                    varLastfak(x) = lastFak &" ("& lastFaknr &")"
                    varShadowCopy(x) = oRec3("shadowcopy")
	                varFakfindes(x) = fakfindes
                    
                end if
                oRec3.close
                
                
                
                '*** Ventetimer ****
                strSQLven = "SELECT f.fid, SUM(fms.venter) AS ventetimer FROM fakturaer f "_
                &" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid)"_
		        &" WHERE (f.jobid = "& oRec("id") &" AND "_
		        &" f.shadowcopy <> 1) "_
		        &" GROUP BY f.jobid, f.fid ORDER BY f.fakdato DESC"
        		
		        'Response.Write strSQLven &"<br>" 
		        'Response.flush
        		
		        oRec3.open strSQLven, oConn, 3
                if not oRec3.EOF then
                    
                ventetimer(x) = oRec3("ventetimer")  
                    
                end if
                oRec3.close
                
                if ventetimer(x) <> 0 then
                ventetimer(x) = ventetimer(x)
                else
                ventetimer(x) = 0
                end if
                
                
                
                if cint(fakfindes) = 1 then
                
                lastFakuseSQL = dateadd("d", 1, lastFak)
                lastFakuseSQL = year(lastFakuseSQL) &"/"& month(lastFakuseSQL) &"/"& day(lastFakuseSQL)
                
                lastFakuseStDato(x) = dateadd("d", 1, lastFak)
                
                else
                
                lastFakuseSQL = sqlDatostart
                lastFakuseStDato(x) = strDag&"/"& strMrd &"/"& strAar
                
                end if
                
               
                
                '** Timer **'
                sumTimerVedFak = 0
                strSQLtimer = "SELECT sum(timer) AS sumtimerEfterFak FROM timer WHERE "_
                &" (tfaktim = 1 OR tfaktim = 2 OR tfaktim = 6) AND (tjobnr = "& oRec("jobnr") &""_ 
	            &" AND tdato BETWEEN '"& lastFakuseSQL &"' AND '"& sqlDatoslut &"') GROUP BY tjobnr"
	            
	           'Response.Write strSQLtimer &"<br>"
	            
	            oRec2.open strSQLtimer, oConn, 3
                while not oRec2.EOF
                
                sumTimerVedFak = oRec2("sumtimerEfterFak")
                
                oRec2.movenext
                wend
                oRec2.close
                
                '** Fak Timer **'
                sumtimerFak = 0
                strSQLtimer = "SELECT sum(timer) AS sumtimerFak FROM timer WHERE "_
                &" (tfaktim = 1) AND (tjobnr = "& oRec("jobnr") &""_ 
	            &" AND tdato BETWEEN '"& lastFakuseSQL &"' AND '"& sqlDatoslut &"') GROUP BY tjobnr"
	            
	           'Response.Write strSQLtimer &"<br>"
	            
	            oRec2.open strSQLtimer, oConn, 3
                while not oRec2.EOF
                
                sumtimerFak = oRec2("sumtimerFak")
                
                oRec2.movenext
                wend
                oRec2.close
                
                '** Ikke godkendte Timer **'
                sumtimerIG = 0
                strSQLtimer = "SELECT sum(timer) AS sumtimerIG FROM timer WHERE "_
                &" (tfaktim = 1 OR tfaktim = 2 OR tfaktim = 6) AND (tjobnr = "& oRec("jobnr") &""_ 
	            &" AND tdato BETWEEN '"& lastFakuseSQL &"' AND '"& sqlDatoslut &"') AND godkendtstatus = 0 GROUP BY tjobnr"
	            
	           'Response.Write strSQLtimer &"<br>"
	            
	            oRec2.open strSQLtimer, oConn, 3
                while not oRec2.EOF
                
                sumtimerIG = oRec2("sumtimerIG")
                
                oRec2.movenext
                wend
                oRec2.close
                
                
                
                '*** Materialer
                matAntalEfterFak = 0
                strSQLMAT = "SELECT sum(matantal) AS matAntalEfterFak FROM materiale_forbrug WHERE "_
                &" jobid = "& oRec("id") &""_ 
	            &" AND forbrugsdato BETWEEN '"& lastFakuseSQL &"' AND '"& sqlDatoslut &"' GROUP BY jobid"
	            
	           'Response.Write strSQLstrSQLMAT &"<br>"
	            
	            oRec2.open strSQLMAT, oConn, 3
                while not oRec2.EOF
                
                matAntalEfterFak = oRec2("matAntalEfterFak")
                
                oRec2.movenext
                wend
                oRec2.close
                
               
		
		
		
		varTimer(x) = sumTimerVedFak
		varTimerFak(x) = sumtimerFak
		varTimerIG(x) = sumtimerIG
		varMat(x) = matAntalEfterFak
	
		
		else
		varTimer(x) = -1
		varMat(x) = -1
		varTimerFak(x) = -1
		varTimerIG(x) = -1
		end if
		
		
	if oRec("aid") <> 0 then 
	        aftId(x) = oRec("aid")
	        aftnavn(x) = oRec("aftnavn")
	        aftnr(x) = oRec("aftalenr")
	        aftStdato(x) = oRec("stdato")
	        aftSldato(x) = oRec("sldato")
	        
	        
	        if oRec("advitype") <> 2 then
	        aftType(x) = "Enh./Klip"
	        else
	        aftType(x) = "Periode"
	        end if
	        
	        
	        select case oRec("status")
	        case 0
	        aftStatus(x) = "<i>Lukket</i>"    
	        case 1
	        aftStatus(x) = "<i>Aktiv</i>"    
	        end select
	        
	        
	else
	        aftId(x) = 0
	        aftnavn(x) = ""
	        aftnr(x) = ""
	        aftType(x) = ""
	end if
	

	x = x + 1
	lastKid = oRec("kid") 
	oRec.movenext
	wend
	oRec.close
	
	
	lastKid = 0
	c = 0
	for x = 0 to x - 1
	
	
	'*** Skjul job / aftaler
	if (visikkejob = 0 AND varAftOrJob(x) = 0) OR (visikkejob = 0 AND visikkenuljob = 0) OR (visikkeaft = 0 AND varAftOrJob(x) = 1) then 
	
	'*** Nuljob
	if ((varAftOrJob(x) = 0 AND visikkenuljob = 1 AND (varTimer(x) > 0 OR varMat(x) > 0 OR ventetimer(x) > 0)) OR (visikkeaft = 0 AND varAftOrJob(x) = 1) OR (varAftOrJob(x) = 0 AND visikkenuljob = 0)) then
	
	'** Er der fundet aktive job
	if varTimer(x) <> -1 then
	
	if varAftOrJob(x) = 1 then 'aftale = 1, job = 0
	varBG = "#ffffff"
	jbclass = "lillegray"
	atcls = ""
	fakclass = ""
	else
	varBG = "#EFF3FF"
	jbclass = ""
	atcls = "lillegray"
	    if varShadowCopy(x) = 1 then 'faktura oprettet på aft.
	    fakclass = "lillegray"
	    else
	    fakclass = ""
	    end if
	end if
	
	if lastKid <> kid(x) OR x = 0 then%>
	<tr><td colspan=12 height=15>
        &nbsp;</td></tr>
	<tr><td bgcolor="#ffffff" colspan=2 style="border-bottom:0px silver dashed; border-left:1px #8cAAE6 solid; border-right:1px #8cAAE6 solid; border-top:1px #8cAAE6 solid; padding:5px 2px 2px 6px;">
	
    <img src="../ill/ac0009-16.gif" />&nbsp;<b><%=knavn(x)%>
	&nbsp;(<%=knr(x)%>)</b>
	<% if trim(len(kans1(x))) <> 0 then%>
	<br /><%=kans1(x)%>
	<%end if %>
	<% if trim(len(kans2(x))) <> 0 then%>
	, <%=kans2(x)%>
	<%end if %>
	
	</td>
	<td colspan=10>
        &nbsp;</td>
	</tr>
	
	<tr bgcolor="#8cAAE6">
	    <td height=25 style="width:210px; border-top:0px #ffffff solid; border-left:0px #ffffff solid; padding:0px 0px 0px 4px;" class=alt>
            &nbsp;<b>Job</b> - type - periode - status</td>
	     <td class=alt style="border-top:0px #ffffff solid; padding:0px 4px 0px 0px;"><b>Job anvarlig(e)</b></td>
	    <td class=alt align=right style="padding:0px 10px 0px 0px; width:160px; border-top:0px #ffffff solid;">
	    <b>Timer</b> | Fak.bare | Ikke godk.
	    </td>
	    <td class=alt align=right style="padding:0px 10px 0px 0px; border-top:0px #ffffff solid;">(vente t.)</td>
	    <td class=alt align=right style="padding:0px 10px 0px 0px; border-top:0px #ffffff solid;"><b>Mat.</b></td>
	    <td class=alt style="padding:0px 10px 0px 0px; border-top:0px #ffffff solid;">
            &nbsp;Seneste Faktura</td>
           <td class=alt style="padding:0px 10px 0px 0px; border-top:0px #ffffff solid;">
            Opret ny</td>
	    <td bgcolor=#C4C4C4 style="padding:0px 10px 0px 6px; border-top:0px #ffffff solid;"><b>Aftale</b> - type - periode - status</td>
	    <td bgcolor=#C4C4C4 style="padding:0px 10px 0px 0px; border-top:0px #ffffff solid;" align=right><b>Enh.</b></td>
	    <td bgcolor=#C4C4C4 style="padding:0px 10px 0px 0px; border-top:0px #ffffff solid;" align=right><b>Mat.</b></td>
	    <td bgcolor=#C4C4C4 style="padding:0px 10px 0px 0px; border-top:0px #ffffff solid;">
            &nbsp;Seneste Fak.</td>
	    <td bgcolor=#C4C4C4 style="border-right:1px #ffffff solid; border-top:0px #ffffff solid;">
            Opret ny</td>
	</tr>
	
	<%end if %>
	<tr bgcolor="<%=varBG%>">
    <td bgcolor="#EFF3FF" valign=top class="<%=jbclass%>"  style="border-bottom:1px silver dashed; padding:3px 2px 2px 6px;"> 
	
	
	           
	           <%
	           if varAftOrJob(x) = 1 then 'aftale
	            
	           strSQLj = "SELECT j.jobnavn, j.jobnr, j.jobstartdato, j.jobslutdato, j.jobstatus, j.fastpris FROM job j WHERE j.serviceaft = "& aftId(x)
	           oRec3.open strSQLj, oConn, 3
	           js = 0
	           while not oRec3.EOF
	           if js <> 0 then%>
	           <br />
	           <%end if%>
	           <b><%=oRec3("jobnavn")%>&nbsp;(<%=oRec3("jobnr") %>)</b>
	           <%if cint(oRec3("fastpris")) = 1 then %>
	           <i>- fastpris</i>
	           <%else %>
	           <i>- lbn. timer</i>
	           <%end if %>
	           <br /><%=replace(formatdatetime(oRec3("jobstartdato"),2),"-",".") %> til <%=replace(formatdatetime(oRec3("jobslutdato"),2),"-",".")%>
	           
	           <% 
	            select case oRec3("jobstatus")
                case 0
                jobst = "Lukket"
                case 1
                jobst = "Aktiv"
                case 2
                jobst = "Passiv"
                end select
	            
	            Response.Write " - <i>"& jobst &"</i>"
	           
	           js = js + 1
	           oRec3.movenext
	           wend
	           oRec3.close
	           
	           else
	           %>
	           
	           <b><%=jobnavn(x)%>&nbsp;(<%=jobnr(x)%>)</b> 
	           <%if cint(jobtype(x)) = 1 then %>
	           <i>- fastpris</i>
	           <%else %>
	           <i>- lbn. timer</i>
	           <%end if %>
	           <br /><font class=megetlillesort><%=replace(formatdatetime(jobstartdato(x),2),"-",".") %> til <%=replace(formatdatetime(jobslutdato(x),2),"-",".")%> - <i><%=jobstatus(x) %></i></font>
	
	           
	           <%
	           end if
	           %>
	           
	 &nbsp;</td>
	<td bgcolor="#EFF3FF" class=lille valign=top style="border-bottom:1px silver dashed; padding:3px 2px 2px 2px;">
	<%=jobans1(x)%><br />
	<%=jobans2(x)%>
	</td>
	<td bgcolor="#EFF3FF" align=right valign=top style="border-bottom:1px silver dashed; padding:3px 5px 2px 2px;">
	
	<%if varTimer(x) <> 0 then %>
	<a href="joblog.asp?rdir=nwin&FM_medarb=0&FM_radio_projgrp_medarb=1&FM_kunde=<%=kid(x) %>&FM_job=<%=jobid(x)%>&FM_start_dag=<%=day(lastFakuseStDato(x))%>&FM_start_mrd=<%=month(lastFakuseStDato(x))%>&FM_start_aar=<%=year(lastFakuseStDato(x))%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&FM_usedatokri=1" target="_blank" class=vmenu>
	<%=formatnumber(varTimer(x), 2)%></a>
    
    | <%=formatnumber(varTimerFak(x),2) & " | "& formatnumber(varTimerIG(x),2) %>
    
   
    <%
    totaltimertilfak = totaltimertilfak + varTimer(x)
    
    
    end if %>
    
   &nbsp;</td>
   <td bgcolor="#EFF3FF" align=right valign=top style="border-bottom:1px silver dashed; padding:3px 5px 2px 2px;">
	<%if ventetimer(x) <> 0 then %>
	(<%=formatnumber(ventetimer(x), 2) %>)
	<%end if %>
	&nbsp;
	</td>
	<td bgcolor="#EFF3FF" align=right valign=top style="border-bottom:1px silver dashed; padding:3px 10px 2px 2px;">
	<%=varMat(x)%>
	&nbsp;
	</td>
	
	<td bgcolor="#EFF3FF" valign=top class=lillegray style="border-bottom:1px silver dashed; padding:5px 2px 2px 2px;">
	<%if varAftOrJob(x) = 0 then 'job %>
    &nbsp;<%=varLastfak(x)%>
    <%end if %>
        &nbsp;</td>
	
	<td bgcolor="#EFF3FF" align=right valign=top style="border-bottom:1px silver dashed; border-right:1px #C4C4C4 solid; padding:1px 5px 0px 2px;">
        <%if print <> "j" then
        
        if varAftOrJob(x) = 0 then 'job%>
	    
	    &nbsp;<a href="../timereg/erp_opr_faktura_fs.asp?FM_kunde=<%=kid(x)%>&FM_job=<%=jobid(x)%>&FM_aftale=0&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&reset=1" class=todo_mellem><img
            src="../ill/ac0010-16.gif" border="0" alt="Opret faktura" /> </a>
		
		<%end if%>
		<%end if%>
		&nbsp;</td>
	
	
	
	<td bgcolor="#FFFFFF" class="<%=atcls%>" valign=top style="border-bottom:1px silver dashed; padding:3px 2px 2px 6px;">
	<%if aftId(x) <> 0 then %>
	<b><%=aftnavn(x)%> (<%=aftnr(x) %>)</b> - <i><%=aftType(x)%></i><br />
	        
	        <%if varAftOrJob(x) = 1 then %>
	        <font class=megetlillesort><%= replace(formatdatetime(aftStdato(x),2),"-",".") %> til <%= replace(formatdatetime(aftSldato(x),2),"-",".") %> - <%=aftStatus(x) %></font>
	        <%else %>
	        <%= replace(formatdatetime(aftStdato(x),2),"-",".") %> til <%= replace(formatdatetime(aftSldato(x),2),"-",".") %> - <%=aftStatus(x) %>
	        <%end if %>
	        
	        
	        
	<%else %>
	&nbsp;
	<%end if %>
	</td>
	<td bgcolor="#FFFFFF" align=right valign=top style="border-bottom:1px silver dashed; padding:3px 10px 2px 2px;">
	<%if len(varEnh(x)) <> 0 then %>
	<b><%=formatnumber(varEnh(x), 2)%></b>
	<%end if %>
	&nbsp;
	</td>
	<td bgcolor="#FFFFFF" align=right valign=top style="border-bottom:1px silver dashed; padding:3px 10px 2px 2px;">
	<%=varAftMat(x)%>
	&nbsp;
	</td>
	
	
	<td bgcolor="#FFFFFF" valign=top class=lillegray style="border-bottom:1px silver dashed; padding:5px 2px 2px 2px;">
	<%if varAftOrJob(x) = 1 then 'aftale %>
    &nbsp;<%=varLastfak(x)%>
    <%end if %>
        &nbsp;</td>
	
	<td bgcolor="#FFFFFF" valign=top align=right style="border-bottom:1px silver dashed; padding:1px 5px 0px 2px;">
        <%if print <> "j" then
        
        if varAftOrJob(x) = 1 then 'aftale
        %>
	    &nbsp;<a href="../timereg/erp_opr_faktura_fs.asp?FM_kunde=<%=kid(x)%>&FM_job=0&FM_aftale=<%=aftId(x)%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd %>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&reset=1" class=todo_mellem><img
            src="../ill/ac0010-16.gif" border="0" alt="Opret faktura" /> </a>
		<%end if%>
		<%end if%>
		&nbsp;</td>
	
	</tr>
	
	<%
	lastKid = kid(x)
	c = c + 1
	
	end if
	
	end if
	
	end if
	
	
	next
	%>
	</table>
	<br />
	Ialt fundet: <b><%=c-1%></b> job og aftaler.
	<br />
	<%if len(totaltimertilfak) <> 0 then %>
	Timer til fakturering: <b><%=formatnumber(totaltimertilfak, 2) %> t.</b>
	<%end if %>
	
	<br /><br /><br />
        &nbsp;
	</div>
	
	
	
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->