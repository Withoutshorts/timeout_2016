<%

        '**** S�gekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
       

            '*** Denne bruges af 2015/materialerudlaeg.asp

          
             case "FN_vis_pas_medarb_matudlaeg2015"
      
                        usemrn = request("matreg_medid")
                        vis_passive = request("vispassive_medarb")
                        jqHTML = "1"

                
                        call selmedarbOptions_2018

                        '*** ��� **'
                        call jq_format(strMedarbOptionsHTML)
                        strMedarbOptionsHTML = jq_formatTxt

                response.write strMedarbOptionsHTML
                response.end


                        case "FN_sogjobogkunde_matudlaeg2015"

                        call selectJobogKunde_jq()


                        case "FN_sogakt_matudlaeg2015"
              
                        call selectAkt_jq


                        case "FN_indlasmat_matudlaeg2015"
            
                            jobid = request("jobid")
                            aktid = request("aktid")
                            medid = request("mid")
                            aftid = 0
                            matId = 0
                            intAntal = 1 'Ved udlaeg er antal altid 1
                            strDato = year(now)&"/"&month(now)&"/"&day(now)
                            strEditor = session("user")

                            if len(trim(request("forbrugsdato"))) <> 0 AND isdate(request("forbrugsdato")) = true then
                                regdato = year(request("forbrugsdato"))&"/"&month(request("forbrugsdato"))&"/"&day(request("forbrugsdato"))
                            else
                                regdato = year(now)&"/"&month(now)&"/"&day(now)
                            end if

                            bilagsnr = "" 
                            navn = request("navn")
                            FM_pris = request("FM_pris")
                            FM_pris = replace(FM_pris, ".", ",")
                            salgspris = 0
                            varenr = 0
                            valuta = request("valuta")
                            gruppe = request("gruppe")
                            personlig = request("personlig")
                            intkode = request("intkode")
                            betegnelse = ""
                            opretlager = 0
                            matregid = request("matregid")
                            matava = 0
                            otf = 1
                            matreg_opdaterpris = 0
                            mattype = 1 'Udl�g

                            mat_func = request("mat_func")
                            
                            basic_valuta = request("basic_valuta")
                            basic_kurs = request("basic_kurs")
                            basic_belob = request("basic_belob")
    	    	
                            call indlaes_mat(matregid, otf, medid, jobid, aktid, aftid, matId, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, FM_pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, mat_func, matreg_opdaterpris, matava, mattype, basic_valuta, basic_kurs, basic_belob)

            
       


        '** END 




        '*** DENNE Bruges af resource forecast ******************
        case "FN_getAkt"

         
    

        call aktTyper2009(2) '**** Akt typer der er med i ddagligt timeregnskab ***'
        aty_sql_realhoursAKt = replace(aty_sql_realhours, "tfaktim", "fakturerbar")
        
        if len(trim(request("medarbid"))) <> 0 then
        medarbId = request("medarbid")
        else
        medarbId = 0
        end if

        call positiv_aktivering_akt_fn()
	    
        'aty_sql_realhoursAKt = " fakturerbar = 1"
        jobid = request.form("cust")


           
        
                    
                    'strSQLa = "SELECT navn AS aktnavn, id AS aid, fase FROM aktiviteter WHERE job = "& jobid &" AND aktstatus = 1 AND ("& aty_sql_realhoursAkt &") GROUP BY id ORDER BY fase, sortorder, navn" 
                    
                    call hentaktiviteter(positiv_aktivering_akt_val, jobid, medarbId, aty_sql_realhoursAkt) 
                    
                     'Response.write strSQLa
                     'Response.flush 
                   
                    'strTXT = "<option value='0'>Udspecificer p� aktivitet..?</option>"
                    
                    strTXT = "<option value='0'>V�lg aktivitet..?</option>"
                    strTXT = strTXT & "<option value='0'>(uspecificeret)</option>"
                    
                    
                    

                    oRec4.open strSQLa, oConn, 3
                    While not oRec4.EOF 
                    
                     if ISNULL(oRec4("fase")) <> true AND len(trim(oRec4("fase"))) <> 0 then
                     fsNavn = " | fase: "& oRec4("fase")
                     else
                     fsNavn = ""
                     end if

                    strTXT = strTXT &"<option value='"&oRec4("aid")&"'>"& oRec4("aktnavn") &" "& fsNavn &"</option>"
                    
                    oRec4.movenext
                    wend
                    oRec4.close 

        'strTXT = "hej"
        'strTXT = strTXT & "</select>"

         '*** ��� **'
        call jq_format(strTXT)
        strTXT = jq_formatTxt
        Response.Write strTXT



        case "FN_medipgrp"
        

        medarbid = request("jq_medarb")
        progrp = request("jq_progrp")
        mtype = request("jq_mtyper")
        mtypegrp = request("jq_mtypergrp")    

        'Response.write progrp
        'Response.end

        'thisMiduse = request("jq_medarb")
    	'intMids = split(thisMiduse, ", ")

        'Response.write "<option>"& progrp &" / "& medarbid &" / "& mtype &" / "& mtypegrp &"</option>"
        'Response.end
        
        mansat = request("jq_mansat")
        mansatpas = request("jq_mansatpas")
        mansat12 = request("jq_mansat12")

        

        'Response.write "<option>"& mansat &"</option>"
        'Response.end

         strSQLmansat = "(m.mansat = 1"

        if mansat = "1" then 
      

                if len(trim(mansat12)) <> 0 AND mansat12 <> 0 then

           
                dd = now
                opsagtdatoKri = dateAdd("m", -mansat12, dd)
                opsagtdatoKri = year(opsagtdatoKri) &"/"& month(opsagtdatoKri) &"/1" '& day(opsagtdatoKri) 

                strSQLmansat = strSQLmansat & " OR (m.mansat = 2" 'de-aktiverede
                strSQLmansat = strSQLmansat & " AND opsagtdato > '"& opsagtdatoKri &"')"
                else
                opsagtdatoKri = ""
                strSQLmansat = strSQLmansat & " OR m.mansat = 2" 'de-aktiverede

                end if            


        end if  
    
        if mansatpas = "1" then 
        strSQLmansat = strSQLmansat & " OR m.mansat = 3" 'passive
        end if   

        strSQLmansat = strSQLmansat & ")"

        mft = 0 
	    mSel = ""

        
        if cint(mtypegrp) = 1 then
            
            vlgtmtypgrp = 0
            call mtyperIGrp_fn(vlgtmtypgrp,1)    
            
            strOptionsJq = "<option value='0' DISABLED>Medarbejdertypegruppe(r)</option>"
            strOptionsJq = strOptionsJq & "<option value='0'>Alle</option>"
            'strOptionsJq = "<option value='0' DISABLED></option>"
            
            for t = 1 to UBOUND(kunMtypgrp) 'UBOUND(mtypgrpids)

                if len(trim(kunMtypgrpNavn(t))) <> 0 then
                strOptionsJq =  strOptionsJq & "<option value='"& kunMtypgrp(t) &"'>"& kunMtypgrpNavn(t) &"</option>" 
                end if
            
            next 


        else 
	        call medarbiprojgrp(progrp, medarbid, mtype, -1)'-1
        end if    

	    
        '** ��� **'
        jq_format(strOptionsJq)
        strOptionsJq = jq_formatTxt
        Response.Write strOptionsJq
        Response.end

   


        case "GetGrandFlexSaldo" 'bruges p� week_real_norm_2010.asp (ugegodkendelse)
            
            medid = request("medid")

            

            call licensStartDato()
            call lonKorsel_lukketPer(now, -2, medid)
            call meStamdata(medid)
    
            if cdate(meAnsatDato) > cdate(licensstdato) then
            listartDato_GT = meAnsatDato
            else
            listartDato_GT = licensstdato
            end if
    
            lic_mansat_dt = listartDato_GT
    
            if cdate(lonKorsel_lukketPerDt) > cdate(listartDato_GT) then
            lonKorsel_lukketPerDt = dateAdd("d", 1, lonKorsel_lukketPerDt)
            listartDato_GT = year(lonKorsel_lukketPerDt) &"/"& month(lonKorsel_lukketPerDt) &"/"& day(lonKorsel_lukketPerDt)
            else
            listartDato_GT = year(listartDato_GT) &"/"& month(listartDato_GT) &"/"& day(listartDato_GT)
            end if
    
            slutDato_GT = day(now) &"-"& month(now) &"-"& year(now)
            slutDato_GT = Dateadd("d", -1, slutDato_GT)
            slutDato_GT = year(slutDato_GT) &"-"& month(slutDato_GT) &"-"& day(slutDato_GT)

            call akttyper2009(6)
	        akttype_sel = "#-99#, " & aktiveTyper
	        akttype_sel_len = len(akttype_sel)
	        left_akttype_sel = left(akttype_sel, akttype_sel_len - 2)
	        akttype_sel = left_akttype_sel
            
            call medarbafstem(medid, listartDato_GT, slutDato_GT, 5, akttype_sel, 0)
            response.Write formatnumber(realNormBal, 2)





        end select
        Response.end
        end if  

'******************************************************************************************
'*** AJAX / JQUERY SLUT *******************************************************************
'******************************************************************************************



function stattopmenu()
%><br>

<a href='joblog_timetotaler.asp' class='rmenu'>Grandtotal</a>&nbsp;&nbsp;|&nbsp;&nbsp;

<a href='joblog.asp?menu=stat' class='rmenu'>Joblog & Ugesedler</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<!--<a href='korsel_lukket.asp' class='rmenu'>K�rsel</a>&nbsp;&nbsp;|&nbsp;&nbsp;-->
<a href='oms.asp?menu=stat' class='rmenu'>Oms�tning</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<!--<a href='stat.asp?menu=stat&show=stat_pies' class='rmenu'>%-vis timefordeling</a>&nbsp;&nbsp;|&nbsp;&nbsp;-->

<!--<a href='stat.asp?menu=stat&show=word' class='rmenu'>Eksport</a>&nbsp;&nbsp;|&nbsp;&nbsp;-->

<a href='fomr.asp?menu=stat&func=stat' class='rmenu'>% Fordeling p� forret. omr.</a>&nbsp;&nbsp;|&nbsp;&nbsp;

<%if session("stempelur") <> 0 then %>
<a href='stempelur.asp?menu=stat&func=stat' class='rmenu'>Komme/G� tider (logind historik)</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<%end if %>

<a href='pipeline.asp?menu=stat&func=stat' class='rmenu'>Pipeline</a>&nbsp;&nbsp;|&nbsp;&nbsp;

<%if level = 1 then%>
<a href='saleandvalue.asp?menu=stat&func=stat' class='rmenu'>Medarbejder Salg & V�rdi</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<%end if%>

<%if level = 1 then '** Indtil teamleder er impl. p� k�rsels siden **'%>
<a href='smileystatus.asp?menu=stat' class='rmenu'>Smileys</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<%end if %>

<%if level = 1 then %>
<a href='bal_real_norm_2007.asp?menu=stat&dontdisplayresult=1' class='rmenu'>Medarb. afstemning & L�n (HR list.)</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<%end if %>

<%if level = 1 then %>
<a href='abonner.asp?menu=stat' class='rmenu'>Abonn�r</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<%end if %>

<a href='week_real_norm_2010.asp?menu=stat' class='rmenu'>Afstemn./Godkend Ugesedler</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<%if level <= 2 OR level = 6 then '** Indtil teamleder er impl. p� k�rsels siden **'%>
<a href='stat_korsel.asp?menu=stat' class='rmenu'>K�rsel</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<%end if %>
<a href='materiale_stat.asp?menu=stat' class='rmenu'>Mat. forbrug / Udl�g</a>

<%if level = 1 AND (lto = "intranet - local" OR lto = "epi" OR lto = "epi_cati") then %>
&nbsp;&nbsp;|&nbsp;&nbsp;<a href='timerimperr.asp?menu=stat' class='rmenu'>Timer Import Fejl</a>
&nbsp;&nbsp;|&nbsp;&nbsp;<a href='diff_timer_kons.asp?menu=stat' class='rmenu'>Konsoliderings-oversigt</a>
<%end if %>

<%


call stadeOn()

if jobasnvigv = 1 AND level = 1 then %>
&nbsp;&nbsp;|&nbsp;&nbsp;<a href='stat_opdater_igv.asp?menu=stat' class='rmenu'>Stade-indmeldinger</a>
<%end if %>


<br>&nbsp;



<%
end function





	
	
	'**** Projektgrupper medarbejdere, kunder job og aftler faste s�ge filterkriterier i stat.
	'**** Grnadtot, Joblog, Oms. Mat.forbrug, Forretningsomr�der
    public jobstKri, viskunabnejob0, viskunabnejob1, viskunabnejob2, segment, segmentSQLkri, jurMedidsSQL, jurJobidsSQL
	sub medkunderjob


    '********** SKAL felter v�re med oncahnge submit ******************'
    select case thisfile
        case "oms"
            useoncahnge = 0
        case else
            useoncahnge = 1
    end select




    '***** Multiple jur enheder ****'
            
            call multible_licensindehavereOn() 
            if cint(multible_licensindehavere) = 1 then

                   
            
                    if level = 1 then
                    
                      visalleJur = 0
                      jurJobidsSQL = ""
                      jurMedidsSQL = ""

                    else

                      visalleJur = 0
                      call meStamdata(session("mid"))
                      visalleJur = meMed_lincensindehaver
                    
                     
                      jurJobidsSQL = " AND instr(lincensindehaver_faknr_prioritet_job, '#"& visalleJur &"#') "
                      jurMedidsSQL = " AND med_lincensindehaver = "& visalleJur

                    end if


                


            else

            visalleJur = 0
            jurJobidsSQL = ""
            jurMedidsSQL = ""

            end if

    '*************** SLUT jur enheder *********************************************
              
        
   
	
	if print <> "j" AND media <> "export" AND media <> "print" then
	%>
	<tr><td colspan="5"><span id="sp_med" style="color:#5582d2;">[+] <%=joblog_txt_142 &" "%>&<%=" "& joblog_txt_143 %></span></td></tr>
	<tr id="tr_prog_med" style="display:none; visibility:hidden; z-index:1;">
	    
	    <%call progrpmedarb %>
	    
	  
	</tr>
   
    <% 
     end if 
	    

        if thisfile <> "smileystatus.asp" then

        	if print <> "j" AND media <> "export" AND media <> "print" then
        %>
         <tr><td colspan="5"><span id="sp_kun" style="color:#5582d2;">[+] <%=joblog_txt_144 %>

             <%if thisfile = "joblog_timetotaler" then%>
              & <%=joblog_txt_145 %>
             <%end if %>

                             </span></td></tr>
         

        <tr id="tr_kun" style="display:none; visibility:hidden;">   <%

        end if



	    call segment_kunder 
				
				
				if print <> "j" AND media <> "export" then
				%>
		
		
		          
            	
		            <td valign=top style="padding:30px 10px 20px 10px; width:350px; background-color:#F7F7F7;">
		            <input type="radio" name="FM_kundejobans_ell_alle" value="1" <%=kundejobansCHK1%>><b>B) <%=joblog_txt_154 %></b><br /><span style="font-size:11px; color:#000000;"><%=joblog_txt_155 %>:</span><br />
		            <input type="checkbox" name="FM_kundeans" id="cFM_kundeans" value="1" <%=kundeansChk%>>&nbsp;<%=joblog_txt_156 %> <br>
		            <input type="checkbox" name="FM_jobans" id="cFM_jobans" value="1" <%=jobansChk%>>&nbsp;<%=joblog_txt_157 %>
                    <input type="checkbox" name="FM_jobans2" id="cFM_jobans2" value="1" <%=jobansChk2%>>&nbsp;<%=joblog_txt_158 %>
                    <input type="checkbox" name="FM_jobans3" id="cFM_jobans3" value="1" <%=jobansChk3%>>&nbsp;<%=joblog_txt_159 &" 1-3"%>

                        <%if cint(showSalgsAnv) = 1 AND thisfile = "pipeline" then %>
                        <br /><input type="checkbox" name="FM_salgsans" id="cFM_slagsansv" value="1" <%=salgsansChk%>>&nbsp;<%=joblog_txt_160 &" "%>.
                        <%end if %>
                    <br /><br />
                        <span style="font-size:11px; color:#999999;">(overrules legal entity)</span>
                    
            		
	              </td>
            	 
            		
	            </tr>
	    
	
	        <%
	        end if 'media/print
	    
	    
	    '*** strKnrSQLkri Bruges ogs� i job functionen ***
		if cint(kundeid) = 0 then
		strKnrSQLkri = "("&strKnrSQLkri&")" '" jobknr = 0" '
		else
		strKnrSQLkri = "(jobknr = "& kundeid &")"
		end if

        if cint(jobans) = 1 OR cint(jobans2) = 1 OR cint(jobans3) = 1 then
        strKnrSQLkri = strKnrSQLkri & " OR (( " 'AND
        end if
		
        jansOneFundet = 0
        jansTWOFundet = 0

		'*** Er jobansvarlig tilvalgt ? ** 
		if cint(jobans) = 1 then
		strKnrSQLkri = ""&strKnrSQLkri &" "& jobAnsSQLkri
        jansOneFundet = 1
		end if
		
		'*** Er jobansvarlig 2 (jobejer) tilvalgt ? ** 
		if cint(jobans2) = 1 then
            
            if jansOneFundet = 1 then
            jobAns2SQLkri = replace(jobAns2SQLkri, "xx", "OR")
            else
            jobAns2SQLkri = replace(jobAns2SQLkri, "xx", "")
            end if

		strKnrSQLkri = ""&strKnrSQLkri &" "& jobAns2SQLkri
		jansTWOFundet = 1
        end if

        '*** Er jobansvarlig 3 (jobmedansvarlig 1-3) tilvalgt ? ** 
		if cint(jobans3) = 1 then
            
            if jansOneFundet = 1 OR jansTWOFundet = 1 then
            jobAns3SQLkri = replace(jobAns3SQLkri, "xx", "OR")
            else
            jobAns3SQLkri = replace(jobAns3SQLkri, "xx", "")
            end if

		strKnrSQLkri = ""&strKnrSQLkri &" "& jobAns3SQLkri
		end if

        if cint(jobans) = 1 OR cint(jobans2) = 1 OR cint(jobans3) = 1 then
        strKnrSQLkri = strKnrSQLkri & " )"
        end if


        '*** Salgsansv
        if cint(salgsans) = 1 then
        strKnrSQLkri = strKnrSQLkri & " OR (" & salgsansSQLkri &")" 
        end if

        
		'*** kundeans ****'
		if cint(kundeans) = 1 then
            if cint(jobans) = 1 OR cint(jobans2) = 1 OR cint(jobans3) = 1 then
            kundeAnsSQLKri = " OR (" & kundeAnsSQLKri &"))"
            else
		    kundeAnsSQLKri = " OR (" & kundeAnsSQLKri &")" 'AND
            end if
		else
            if cint(jobans) = 1 OR cint(jobans2) = 1 OR cint(jobans3) = 1 then
            strKnrSQLkri = strKnrSQLkri & ")"
            end if
		kundeAnsSQLKri = ""
		end if 

        if kundejobansCHK1 = "CHECKED" AND ((kundeAnsSQLKri = "") AND (strKnrSQLkri = "( jobknr <> 0 )")) then
        strKnrSQLkri = "( jobknr = -1 AND jobans1 = -1)"
        kundeAnsSQLKri = " AND kundeans1 = -1" 
        end if

        'Response.Write "<br><br>strKnrSQLkri" & strKnrSQLkri & "<br>" 
        'Response.Write "jobAns2SQLkri" & jobAns2SQLkri & "<br>"
        'Response.Write "jobAns3SQLkri" & jobAns3SQLkri & "<br>"
		

        viskunabnejob0 = 0
	    if len(trim(request("viskunabnejob0"))) <> 0 then
	    viskunabnejob0 = request("viskunabnejob0")
	    jost0CHK = "CHECKED"
        else
            if len(trim(request("FM_job"))) <> 0 then
            jost0CHK = ""
            else
            viskunabnejob0 = 1
            jost0CHK = "CHECKED"
            end if
        end if 

        viskunabnejob1 = 0
        if len(trim(request("viskunabnejob1"))) <> 0 then
	    viskunabnejob1 = request("viskunabnejob1")
	    jost1CHK = "CHECKED"
        else
        jost1CHK = ""
        end if 
		
        viskunabnejob2 = 0
        if len(trim(request("viskunabnejob2"))) <> 0 then
	    viskunabnejob2 = request("viskunabnejob2")
	    jost2CHK = "CHECKED"
        else
            if thisfile = "oms" then
                if len(request("FM_kundejobans_ell_alle")) <> 0 then'er der submitted
                 jost2CHK = ""
              
                else
                    viskunabnejob2 = 1
	                jost2CHK = "CHECKED"
                end if
            else    
            jost2CHK = ""
            end if
        end if 


        jobstKri = " AND ("

        '** Vis aktive

        if len(trim(viskunabnejob0)) <> 0 then
        viskunabnejob0 = viskunabnejob0
        else
        viskunabnejob0 = 1
        end if

        'Response.Write viskunabnejob0
        'Response.end

		if cint(viskunabnejob0) <> 0 then
	    jobstKri = jobstKri & " j.jobstatus = 1 "
        else
        jobstKri = jobstKri & " j.jobstatus = -1 "
	    end if

        '**Tilbud

         if len(trim(viskunabnejob1)) <> 0 then
        viskunabnejob1 = viskunabnejob1
        else
        viskunabnejob1 = 0
        end if

        if cint(viskunabnejob1) <> 0 then
	    jobstKri = jobstKri & " OR j.jobstatus = 3 "
	    end if

        '** Lukkede og passive

         if len(trim(viskunabnejob2)) <> 0 then
        viskunabnejob2 = viskunabnejob2
        else
        viskunabnejob2 = 0
        end if

        if cint(viskunabnejob2) <> 0 then
	    jobstKri = jobstKri & " OR j.jobstatus = 2 OR j.jobstatus = 0 OR j.jobstatus = 5 "
	    end if
	    
        '*** Hvis der slet ikke er valgt nogen status'er v�lges aktive automatisk ***'

        if viskunabnejob0 = 0 AND viskunabnejob1 = 0 ANd viskunabnejob2 = 0 then
        jobstKri = " AND (j.jobstatus = 1 "
        jost0CHK = "CHECKED"
        end if

        jobstKri = jobstKri & ")"


        'response.write "jobstKri: "& jobstKri

	    if cint(aftaleid) <> 0 then
	        if cint(aftaleid) <> -1 then
	        strJobAftSQL = " AND serviceaft = " & aftaleid
	        else
	        '** Viser alle uden aftale **'
	        strJobAftSQL = " AND serviceaft = 0" 
	        end if
		else
		strJobAftSQL = ""
		end if		
	    
	    
	    if cint(kundeid) = 0 then
	    strAftKidSQLkri = strAftKidSQLkri '" kundeid <> 0" '
		else
		strAftKidSQLkri = " kundeid = "& kundeid
		end if
	    
	
	
	if print <> "j" AND media <> "export" AND media <> "print" then
	%>
	 <tr><td colspan="5"><span id="sp_job" style="color:#5582d2;">[+] <%=joblog_txt_161 &" "%>&<%=" "& joblog_txt_162 %></span></td></tr>
         <tr id="tr_job" style="display:none; visibility:hidden;">
               
            


        
	    <td valign=top style="padding-top:20px;">
        <b><%=joblog_txt_163 %>:</b><br />
		
      
		<%
		
		
		
		strSQL = "SELECT jobnavn, jobnr, jobstatus, id, k.kkundenavn, k.kkundenr FROM job j "_
		&" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
		&" WHERE "& strKnrSQLkri &" "& kundeAnsSQLKri &" "& jobstKri &" "& strJobAftSQL &" "& jurJobidsSQL &" ORDER BY k.kkundenavn, j.jobnavn"
		
        'if session("mid") = 1 then
        'Response.write strSQL
		'Response.flush  
        'end if
		%>

		<%if cint(useoncahnge) = 1 then %>
		<select name="FM_job" id="FM_job" size=16 style="width:406px; font-size:11px;" onChange="clearJobsog(), submit();">
        <%else %>
        <select name="FM_job" id="FM_job" size=16 style="width:406px; font-size:11px;" onChange="clearJobsog()">"
        <%end if %>

        <%
        
        lastKnr = 0
        useForOmrKri = 0
        if jobid = 0 then
        selo = "SELECTED"
        else
        selo = ""
        end if %>
		<option value="0" <%=selo %>><%=joblog_txt_188 &" "%>(=<%=" "& joblog_txt_164 %>)</option>
		<%
				
				oRec.open strSQL, oConn, 3
				while not oRec.EOF

               

                if strFomr_reljobids <> "0" AND (thisfile = "joblog_timetotaler" OR thisfile = "pipeline") then
                useForOmrKri = 1
                else
                useForOmrKri = 0
                end if
            
				
                if cint(useForOmrKri) = 0 OR (cint(useForOmrKri) = 1 AND instr(strFomr_reljobids, "#"& oRec("id") &"#") <> 0) then 'IKKE EN DEL AF FORRETNINGSOMR�DE)
                

                if len(trim(oRec("kkundenavn"))) <> 0 then
                knavn = replace(oRec("kkundenavn"), "'", "")
                else
                knavn = ".."
                end if



                if lastKnr <> oRec("kkundenr") then
                %>
                <option value="<%=oRec("id")%>" DISABLED></option>
                <option value="<%=oRec("id")%>" DISABLED style="background-color:#Eff3ff;"><%=knavn &" ("& oRec("kkundenr") &")" %></option>
                <%
                end if

				if cdbl(jobid) = cdbl(oRec("id")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				
				jobnavnid = replace(oRec("jobnavn"), "'", "") & " ("& oRec("jobnr") &")"
				
			
                jstTxt = ""
                select case oRec("jobstatus")
                case 0
                jstTxt = " - " & joblog_txt_165
                case 2
                jstTxt = " - " & joblog_txt_166
                case 3
                jstTxt = " - " & joblog_txt_167
                case 4
                jstTxt = " - " & job_txt_098
                case 5
                jstTxt = " - Eval."
                end select
        
				%>
				<option value="<%=oRec("id")%>" <%=isSelected%>><%=jobnavnid &" "& jstTxt%></option>
				<%

                lastKnr = oRec("kkundenr")

                end if 'del af forretninsomr 

               

				oRec.movenext
				wend
				oRec.close
				
				
				
				if cdbl(jobid) = -1 then
				jChk = "SELECTED"
				else
				jChk = ""
				end if
				%>
				<option value="-1" <%=jChk %>><%=joblog_txt_168 %></option>
		</select>
		
	
		</td>
         <td valign=top style="padding-top:20px;">
         Jobfilter (apply to job):

          <%
              if thisfile = "joblog_timetotaler" OR thisfile = "pipeline" then
                    
                    %>
                    
	               <br />
                    <b><%=joblog_txt_145 %>:</b> <br />                              
                    <%
                                           
                                    strSQLf = "SELECT f.navn AS fnavn, f.id, f.konto, kp.kontonr AS kkontonr, kp.navn AS kontonavn FROM fomr AS f "_
                                    &" LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) WHERE f.id <> 0 AND f.jobok = 1 ORDER BY f.navn"
                                    %>
                                    <select name="FM_fomr" id="FM_fomr" multiple="multiple" size="5" style="width:350px;">
                                    <option value="0"><%=joblog_txt_152 %></option>
                                    
                                    <%
                                    
                                    strchkbox = ""
                                    oRec.open strSQLf, oConn, 3
                                    while not oRec.EOF
                                    
                                    if instr(strFomr_rel, "#"&oRec("id")&"#") <> 0 then
                                    fSel = "SELECTED"
                                    else
                                    fSel = ""
                                    end if

                                   
                                    
                                    %>
                                    <option value="<%=oRec("id")%>" <%=fSel %>><%=oRec("fnavn")%></option>
                                    <%
                                        
                                    oRec.movenext
                                    wend
                                    oRec.close
                                    %>
                                </select>
                              

           <% end if 'thisfile = "joblog_timetotaler" OR thisfile = "pipeline") then
               %>

	
	
        
        <%if thisfile <> "pipeline" then %>
         <br /><br />

        <b><%=joblog_txt_162 %>:</b> (<%=joblog_txt_169 %>)&nbsp;<br /><img src="../ill/blank.gif" width="50" height="5"  border="0"/><br />

		<%
			
		strSQL = "SELECT s.navn, s.aftalenr, s.id, k.kkundenavn, k.kkundenr FROM serviceaft s "_
		&" LEFT JOIN kunder k ON (k.kid = s.kundeid) "_
		&" WHERE "& strAftKidSQLkri &" ORDER BY k.kkundenavn, s.navn"
		
		'Response.write strSQL
		%>
		
        <%if cint(useoncahnge) = 1 then %>
		<select name="FM_aftaler" id="FM_aftaler" size=6 style="width:350px; font-size:11px;" onChange="clearJobsog(); submit();">
        <%else %>
        <select name="FM_aftaler" id="FM_aftaler" size=6 style="width:350px; font-size:11px;" onChange="clearJobsog();" >"
        <%end if %>

		<option value="0"><%=joblog_txt_188 %> - <%=joblog_txt_170 %>...</option>
		<%
		
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(aftaleid) = cint(oRec("id")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("id")%> - <%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>) ][ <%=oRec("navn")%>&nbsp;(<%=oRec("aftalenr")%>)</option>
				<%
				oRec.movenext
				wend
				oRec.close
				
				if cint(aftaleid) = -1 then
				aChk = "SELECTED"
				else
				aChk = ""
				end if%>
				<option value="-1" <%=aChk%>><%=joblog_txt_168 %> (<%=joblog_txt_171 %>)</option>
		</select>
        <br /><br /><input type="submit" style="font-size:9px;" value="Opdater jobliste >>">
        
	    <% 
        else
        %>
        &nbsp;
        <%
        end if

       %>
       </td></tr>
       <%

    end if


    
   if media <> "print" AND print <> "j" AND media <> "export" then

%>
            <tr><td colspan="5" style="padding-top:20px;">
            	
		<h4><%=joblog_txt_172 %>: <br /><span style="font-size:11px; font-weight:lighter;">(% wildcard, <b>231, 269</b><%=" "& joblog_txt_173 %>, <b>201--225</b><%=" "& joblog_txt_174 %>, <b><></b><%=" "& joblog_txt_175 %></span></h4>
        <input name="viskunabnejob0" id="viskunabnejob" type="checkbox" value="1" <%=jost0CHK %> /><%=joblog2_txt_148 %> &nbsp;
        <input name="viskunabnejob1" id="Radio3" type="checkbox" value="1" <%=jost1CHK %> /><%=jobstatus_txt_003 %> &nbsp;
        <input name="viskunabnejob2" id="Checkbox1" type="checkbox" value="1" <%=jost2CHK %> /><%=joblog2_txt_150 %> &nbsp;<br />

                <input type="text" name="FM_jobsog" id="FM_jobsog" value="<%=jobSogVal%>" style="width:350px; border:2px #6CAE1C solid; font-size:14px;">&nbsp;
                <input id="Submit1" type="submit" value=" <%=joblog_txt_189 %> >> " style="font-size:11px;" />
            
         

        </td> </tr>

     <%end if

    end if 'thisfile smileystatus.asp

	end sub
    '*******'















    '********************************************************'
    '********************************************************'
    sub segment_kunder

    select case thisfile
        case "oms"
            useoncahnge = 0
        case else
            useoncahnge = 1
    end select
    
    if len(trim(request("FM_segment"))) <> 0 AND request("FM_segment") <> 0 then 
    segment = request("FM_segment") 
    segmentSQLkri = " AND ktype = "& segment
       
    else
    segment = 0
    segmentSQLkri = " AND ktype <> -9999 "
    end if






    '** NACE
    if len(trim(request("FM_nacecode"))) <> 0 AND request("FM_nacecode") <> "0" then 
    naceArr = split(request("FM_nacecode"), ", ")
    nacecodeSQLkri = " AND (nace = '-9999' "

         for p = 0 TO UBOUND(naceArr)
         
            nacecodeSQLkri = nacecodeSQLkri & " OR nace = '"& naceArr(p) & "'" 

         next

    nacecodeSQLkri = nacecodeSQLkri & ")"

   
       
    else
    nace = 0
    nacecodeSQLkri = " AND nace <> -9999 "
    end if


    '** POSTnr
    if len(trim(request("FM_postalcode"))) <> 0 AND request("FM_postalcode") <> "0" then 
    postalcodeArr = split(request("FM_postalcode"), ", ")
    postalcodeSQLkri = " AND (postnr = '-9999' "

         for p = 0 TO UBOUND(postalcodeArr)
         
            postalcodeSQLkri = postalcodeSQLkri & " OR postnr = '"& postalcodeArr(p) & "'" 

         next

    postalcodeSQLkri = postalcodeSQLkri & ")"

   
       
    else
    postalcode = 0
    postalcodeSQLkri = " AND postnr <> -9999 "
    end if


    '** COUNTRY CODE
    if len(trim(request("FM_countrycode"))) <> 0 AND request("FM_countrycode") <> "0" then 
    countrycodeArr = split(request("FM_countrycode"), ", ")
    countrycodeSQLkri = " AND (land = '-9999' "

         for p = 0 TO UBOUND(countrycodeArr)
         
            countrycodeSQLkri = countrycodeSQLkri & " OR land = '"& countrycodeArr(p) & "'" 

         next

    countrycodeSQLkri = countrycodeSQLkri & ")"

   
       
    else
    countrycode = 0
    countrycodeSQLkri = " AND postnr <> -9999 "
    end if

    if print <> "j" AND media <> "export" AND media <> "print" AND media <> "chart" then
    %>
    <td valign=top style="padding-top:30px;">

        <input type="radio" name="FM_kundejobans_ell_alle" value="0" <%=kundejobansCHK0%> onclick="clearK_Jobans();"><b>A)<%=" "& joblog_txt_146 %></b>

        <br /><br /> <%=joblog_txt_238 %>:<br />


       <table>
	  
       <%
       strSQLsegm = "SELECT navn, id FROM kundetyper WHERE id <> 0 ORDER BY navn" 


       %>
	 

     
       <tr>
       <td>
       <%=joblog_txt_147 %>: 
       </td><td>


                <%if cint(useoncahnge) = 1 then %>
                    <select name="FM_segment" style="width:176px; font-size:11px;" onchange="submit();"><!-- clearJobsog(), -->
                <%else %>
                    <select name="FM_segment" style="width:176px; font-size:11px;">
                <%end if %>

      
       
       <option value=0><%=joblog_txt_153 %></option>
       <%
       oRec.open strSQLsegm, oConn, 3
       while not oRec.EOF 


       if cint(segment) = cint(oRec("id")) then
       segSEL = "SELECTED"
       else
       segSEL = ""
       end if
       %>
       <option value="<%=oRec("id") %>" <%=segSEL %>><%=oRec("navn") %></option>
       <%
       oRec.movenext
       wend
       oRec.close
       %>
       
      
       </select>
        </td></tr>

            <tr><td valign="top">NACE code:

                <%if lto = "lm" OR lto = "intranet - local" then %>
                <br /><span style="9px; color:#999999;">(Provsti)</span>
                <%end if %>

                </td><td>
             <%strSQLnacecode = "SELECT nace FROM kunder WHERE kid <> 0 AND nace IS NOT NULL AND nace <> '0' AND nace <> '' AND (useasfak = 0 OR useasfak = 1 OR useasfak = 6) AND ketype <> 'e' GROUP BY nace ORDER BY nace" %>
            <select multiple name="FM_nacecode" style="width:176px;" size="3">
                <option value="0">All</option>
                <%
                   nc = 0
                   oRec.open strSQLnacecode, oConn, 3
                   while not oRec.EOF 


                   if instr(nacecodeSQLkri, oRec("nace")) <> 0 then
                   nacecodeSEL = "SELECTED"
                   else
                   nacecodeSEL = ""
                   end if
                   %>
                   <option value="<%=oRec("nace") %>" <%=nacecodeSEL %>><%=oRec("nace") %></option>
                   <%
                   nc = nc + 1
                   oRec.movenext
                   wend
                   oRec.close

                   if nc = 0 then
                   %>
                   <option value="0">No NACE codes found </option>
                   <%end if %>
            </select>
            </td></tr>

       
          <tr>
            <td valign="top">Postal code:
                <%if lto = "lm" OR lto = "intranet - local" then %>
                <br /><span style="9px; color:#999999;">(Distrikt)</span>
                <%end if %>
            </td>
            <td>
            <%strSQLpostalcode = "SELECT postnr FROM kunder WHERE kid <> 0 AND postnr IS NOT NULL AND postnr <> '0' AND postnr <> '' AND (useasfak = 0 OR useasfak = 1 OR useasfak = 6) AND ketype <> 'e' GROUP BY postnr ORDER BY postnr" %>
            <select multiple name="FM_postalcode" style="width:176px;" size="3">
                <option value="0">All</option>
                <%
                   pc = 0
                   oRec.open strSQLpostalcode, oConn, 3
                   while not oRec.EOF 


                   if instr(postalcodeSQLkri, oRec("postnr")) <> 0 then
                   postalcodeSEL = "SELECTED"
                   else
                   postalcodeSEL = ""
                   end if
                   %>
                   <option value="<%=oRec("postnr") %>" <%=postalcodeSEL %>><%=oRec("postnr") %></option>
                   <%
                   pc = pc + 1
                   oRec.movenext
                   wend
                   oRec.close

                   if pc = 0 then
                   %>
                   <option value="0">No Postalcodes found </option>
                   <%end if %>
            </select>
            </td></tr>

           

            <tr><td valign="top">Country code:</td><td>
             <%strSQLcountrycode = "SELECT land FROM kunder WHERE kid <> 0 AND land IS NOT NULL AND land <> '0' AND land <> '' AND (useasfak = 0 OR useasfak = 1 OR useasfak = 6) AND ketype <> 'e' GROUP BY land ORDER BY land" %>
            <select multiple name="FM_countrycode" style="width:176px;" size="3">
                <option value="0">All</option>
                <%
                   nc = 0
                   oRec.open strSQLcountrycode, oConn, 3
                   while not oRec.EOF 


                   if instr(countrycodeSQLkri, oRec("land")) <> 0 then
                   countrycodeSEL = "SELECTED"
                   else
                   countrycodeSEL = ""
                   end if
                   %>
                   <option value="<%=oRec("land")%>" <%=countrycodeSEL %>><%=oRec("land") %></option>
                   <%
                   nc = nc + 1
                   oRec.movenext
                   wend
                   oRec.close

                   if nc = 0 then
                   %>
                   <option value="0">No countrycodes found </option>
                   <%end if %>
            </select>
            </td></tr>
            

            </table><br /><input type="submit" style="font-size:9px;" value="Opdater kundeliste >>">
           
           <br /><img src="../ill/blank.gif" width="1" height="5"  border="0"/><br />


           <%if thisfile = "xxpipeline" then %>
           <span style="float:right; padding-right:50px;"><input type="checkbox" DISABLED name="fomr_onjob" value="1" <%=fomr_onjobCHK %> /> Segment must match on job <br />(ignore customer)</span><br /><br />
           <%end if %>
		
		<%
		
		
	end if' print	
	
	
		strKnrSQLkri = " OR jobknr = -1 "
		strAftKidSQLkri = " OR kundeid = -1 "
		
	
		strSQL = "SELECT Kkundenavn, Kkundenr, Kid, useasfak FROM kunder WHERE (" & kundeAnsSQLKri & ") AND ketype <> 'e'"_
        &""& segmentSQLkri &" AND (useasfak = 0 OR useasfak = 1 OR useasfak = 6) "& postalcodeSQLkri &" "& nacecodeSQLkri &" "& countrycodeSQLkri &" ORDER BY useasfak, Kkundenavn"
		
        'if session("mid") = 1 then
        'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>KunderSQL:"& strSQL & "<br>"
		'Response.flush		
        'end if
	
	if print <> "j" AND media <> "export" AND media <> "print" AND media <> "chart" then
		
        %>
       <br /><b>V�lg kunde
           
           <%if lto = "nt" OR lto = "intranet - local" then %>
           / leverand�r
           <%end if %>
           :</b><br />
      

            <%if cint(useoncahnge) = 10000 then '�ndret 20190506 %>
		        <select name="FM_kunde" id="FM_kunde" style="width:406px; font-size:11px;" size="10" onChange="submit();"><!-- clearJobsog(), -->
            <%else %>
                <select name="FM_kunde" id="FM_kunde" style="width:406px; font-size:11px;" size="10">
            <%end if %>
       

        <%if kundeid = 0 then
        selo = "SELECTED"
        else
        selo = ""
        end if %>

		<option value="0" <%=selo %>><%=joblog_txt_150 %>...</option>
		<%
	end if
				
				oRec.open strSQL, oConn, 3
				k = 0
				while not oRec.EOF
				
				
				'*** Bruges til SQL kald l�ngere nede
				if cint(kundeans) = 1 OR cint(segment) <> 0 then
				strKnrSQLkri = strKnrSQLkri & " OR jobknr = "& oRec("kid")
				strAftKidSQLkri = strAftKidSQLkri & " OR kundeid = " & oRec("kid")
				else
				strAftKidSQLkri =  " OR kundeid <> 0 "
				strKnrSQLkri = " OR jobknr <> 0 "
				end if
				
				
				if print <> "j" AND media <> "export" ANd media <> "chart" AND media <> "print" then
			
            
                if lastuseAsFak <> oRec("useasfak") OR k = 0 then

                select case oRec("useasfak")
                case 0
                useAsFakTXT = tsa_txt_530 'kunde
                case 1
                useAsFakTXT = tsa_txt_532 'licensindehaver
                case 6
                useAsFakTXT = tsa_txt_531 'leverand�r
                end select

                   
                %>
			    <option value="0" DISABLED></option>
			    <%
                    

                %>
				<option value="0" DISABLED><%=useAsFakTXT %></option>
				<%
                end if 
            
            
            	
				if cint(kundeid) = cint(oRec("kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if

				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
				<%
				
				end if
				

                lastuseAsFak = oRec("useasfak")
				k = k + 1
				oRec.movenext
				wend
				oRec.close
				
				
				if print <> "j" AND media <> "export" AND media <> "print" then
				%>
				</select><br /><span style="font-size:9px; color:#999999;">(<%=k %>)</span><br /><br />&nbsp;
              
		        <%end if


            

             

				
                '*** TRIMMER KUNDE SEL SQL 
				len_strKnrSQLkri = len(strKnrSQLkri)
				right_strKnrSQLkri = right(strKnrSQLkri, len_strKnrSQLkri - 3)
				strKnrSQLkri = right_strKnrSQLkri
				
				len_strAftKidSQLkri = len(strAftKidSQLkri)
				right_strAftKidSQLkri = right(strAftKidSQLkri, len_strAftKidSQLkri - 3)
				strAftKidSQLkri = right_strAftKidSQLkri




                if print <> "j" AND media <> "export" then
				%>
			
                </td>
		        <%end if

    end sub


        








	
	
	public jobnrSQLkri, jobidFakSQLkri, jidSQLkri, jobnr, j, jobKri
	function valgtejob()
	
				
	'**** Job *****
	jobSQLfundet = 0
	jobKri = " mf.jobid = 0 "
	
	'Hvis der er valgt et job ==> V�lg et enkelt job
	if cdbl(jobid) <> 0 then
            
            '**Ignorer valgt status ved s�g p� jobnr **''
            'if len(trim(jobSogVal)) <> 0 then
            'jobstKri = ""
            'end if

	strSQL = "SELECT jobnr, id FROM job j WHERE id = " & jobid &" "& jobstKri '&" AND jobstatus <> 3 " '*tilbud
	jobSQLfundet = 1
	end if
	
	'*** job sog ***

    jobSogVal = replace(jobSogVal, " ", "")
	if len(trim(jobSogVal)) <> 0 AND jobSQLfundet = 0 then
	    
	            if instr(jobSogVal, ",") <> 0 then '** Komma **'
	            
                    jobSQLkri = " jobnr = '0' "
	                jobSogValuse = split(jobSogVal, ",")
	                for i = 0 TO UBOUND(jobSogValuse)
                        call erDetInt(jobSogValuse(i)) 
			            if isInt > 0 then
	                    jobSQLkri = jobSQLkri & " OR jobnr = '"& jobSogValuse(i) &"'"  
                        else
                        jobSQLkri = jobSQLkri & " OR (jobnr = "& jobSogValuse(i) &")"
                        end if 
	                next
	    
	                jobSQLkri = "("&jobSQLkri&") "
	    
	            else


                    if instr(jobSogVal, "--") <> 0 then '** Interval
	                jobSogValuse = split(jobSogVal, "--")
	                jobSQLkri = "(jobnr >= '"& trim(jobSogValuse(0)) &"' AND jobnr <= '" & trim(jobSogValuse(1)) & "')"   
	                else 
                        
                    
                        if instr(jobSogVal, ">") <> 0 OR instr(jobSogVal, "<") <> 0 then '** St�rre mindre <>
	                            
                                if instr(jobSogVal, ">") <> 0 then
                                jobSogValuse = replace(jobSogVal, ">", "")
                                jobSQLkri = "(jobnr > '"& trim(jobSogValuse) &"')"
                                else
                                jobSogValuse = replace(jobSogVal, "<", "")
                                jobSQLkri = "(jobnr < '"& trim(jobSogValuse) &"')"
                                end if          

                        else    '*** Alm s�gning 

                            jobSQLkri = "(jobnavn LIKE '" & jobSogVal &"%' OR jobnr = '"& jobSogVal &"')"
                        end if

                    end if

	            end if
	
	strSQL = "SELECT jobnr, id FROM job j WHERE "& jobSQLkri '&" AND "& strKnrSQLkri '&" AND jobstatus <> 3 "
	jobSQLfundet = 1
	
    'if session("mid") = 1 then
    'Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br># "&  strSQL
    'Response.flush
    'end if
    
    end if
	

    if cint(kundeans) = 1 AND (cint(jobans) = 1 OR cint(jobans2) = 1 OR cint(jobans3) = 1) then
       strKnrSQLkri = strKnrSQLkri & ")"
    end if

	if (cint(aftaleid) > 0 OR cint(aftaleid) = -1) AND jobSQLfundet = 0  then 'AND jobid <> 0
	    '*** Ingen eller en special aftale    
	    if cint(aftaleid) = -1 then
	    aftIdSQL = "serviceaft <> -99"
	    else
	    aftIdSQL = "serviceaft = " & aftaleid
	    end if
	
	strSQL = "SELECT id, jobnr, jobnavn FROM job j WHERE "& strKnrSQLkri &" AND "& aftIdSQL &" "& jobstKri 'AND jobstatus <> 3 " '*tilbud
	jobSQLfundet = 1
	end if	
		
	if jobSQLfundet = 0 then
	strSQL = "SELECT id, jobnr, jobnavn FROM job j WHERE "& strKnrSQLkri &" "& jobstKri 'AND jobstatus <> 3 " '*tilbud
	jobSQLfundet = 1
	end if		
	
	strSQL = strSQL & " ORDER BY jobnavn "			
				
                
				'Response.write "<br><br><br><br>GHer:"
                'if session("mid") = 1 then
				'Response.write strSQL &"<br><br>"
				'Response.flush
                'end if
				
				oRec.open strSQL, oConn, 3
				j = 0
				while not oRec.EOF
				
				if j = 0 then
					jobidFakSQLkri = ""
					jobnrSQLkri = ""
					jidSQLkri = ""
					jobKri =  " mf.jobid = 0 "
				end if


				
                    if strFomr_reljobids <> "0" AND (thisfile = "joblog_timetotaler" OR thisfile = "pipeline") then 
                    useForOmrKri = 1
                    else
                    useForOmrKri = 0
                    end if
            
				
                    'response.Write "<br><br><br><br><br>strKnrSQLkri: "& strKnrSQLkri &" useForOmrKri: "& useForOmrKri & " HER:  " & strFomr_reljobids & "<br>" & thisfile

                    if cint(useForOmrKri) = 0 OR (cint(useForOmrKri) = 1 AND instr(strFomr_reljobids, "#"& oRec("id") &"#") <> 0) then 'IKKE EN DEL AF FORRETNINGSOMR�DE)


					jobnrSQLkri = jobnrSQLkri & " OR tjobnr = '"& oRec("jobnr") & "'"
					jobidFakSQLkri = jobidFakSQLkri & " OR jobid = " & oRec("id")
					jidSQLkri = jidSQLkri & " OR id = " & oRec("id")
					jobKri = jobKri & " OR mf.jobid = "& oRec("id") 
					
					
				
				    j = j + 1
                    end if ' FORRETNINGSOMR�DE KRI


				oRec.movenext
				wend
				
				oRec.close
				

                'if session("mid") = 1 then
				'Response.write strSQL & "<br>#" & jidSQLkri &"<br><br>"

                    if len(trim(jidSQLkri)) = 0 then
                    jobidFakSQLkri = " OR jobid = -1 "
	                jobnrSQLkri = " OR tjobnr = '-1' "
	                jidSQLkri = " OR id = -1 "
	                seridFakSQLkri = " OR aftaleid = -1 "
                    end if

				'Response.flush
                'end if
				
	end function
	
	
	
	public seridFakSQLkri, a
	function valgteaftaler()
			
				
				a = 0
				
				aftSQLfundet = 0
				
				'*** Aftale valgt ****
				if cint(aftaleid) <> 0 then
				strSQL = "SELECT id, kundeid, navn, aftalenr FROM serviceaft WHERE id = "& aftaleid 
				aftSQLfundet = 1
				end if
				
				
				if aftSQLfundet = 0 then
				strSQL = "SELECT id, kundeid, navn, aftalenr FROM serviceaft WHERE "& strAftKidSQLkri 
				aftSQLfundet = 1
				end if
				
				strSQL = strSQL & " ORDER BY navn "	
				
				
				'Response.Write strSQL
				'Response.flush
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if a = 0 then
					seridFakSQLkri = ""
				end if
				
				seridFakSQLkri = seridFakSQLkri & " OR aftaleid = " & oRec("id")
				
				
				a = a + 1
				oRec.movenext
				wend
				
				oRec.close
			
			
	end function
	
	
	public strKnrSQLkri, visKundejobans, kundejobansCHK1, kundejobansCHK0, kontakt_keyaccountVAL
	public kundeans, kundeansChk, jobans, jobans2, jobansChk, jobansVal2, jobansVal, kansVal, jobansChk2, jobans1Val
    public jobans3, jobansChk3, jobansVal3, salgsans, salgsansChk, salgsansVal 
	function kundeogjobans()
	


	if len(request("FM_kundejobans_ell_alle")) <> 0 then
	visKundejobans = request("FM_kundejobans_ell_alle")
    response.cookies("tsa")("keyacc") = visKundejobans
	else
        if request.cookies("tsa")("keyacc") = "1" then
        visKundejobans = 1
        else
	    visKundejobans = 0
        end if
	end if
	
	
                    

	if cint(visKundejobans) = 1 then
	kundejobansCHK1 = "CHECKED"
	kundejobansCHK0 = ""
	
	kontakt_keyaccountVAL = "Key account" 
	kansVal = ""
	jobansVal = ""
	

                    
                if len(request("FM_kundejobans_ell_alle")) <> 0 then
				'if len(request("FM_kundeans")) <> 0 then
				kundeans = request("FM_kundeans")
					if cdbl(kundeans) = 1 then 
					kundeansChk = "CHECKED"
					kansVal = " - Kundeansvalig"
					else
					kundeans = 0
					kundeansChk = ""
					end if
				    
                    response.cookies("tsa")("kans") = kundeans
                else
                    if request.cookies("tsa")("kans") = "1" then
                    kundeans = 1
                    kundeansChk = "CHECKED"
					kansVal = " - Kundeansvalig"
                    else
				    kundeansChk = ""
				    kundeans = 0
                    end if
				end if
				
                 

				'*** Jobans ***
				if len(request("FM_kundejobans_ell_alle")) <> 0 then
                    
                    if len(request("FM_jobans")) <> 0 then
				    jobans = request("FM_jobans")
					else
                    jobans = 0
                    end if
                    
                    select case cdbl(jobans)
					case 1 
					jobansChk = "CHECKED"
					jobansVal = " - Jobansvarlig"
					case else
					jobansVal = ""    
				    jobansChk = ""
					end select

                    response.cookies("tsa")("jans") = jobans

				else
                        if request.cookies("tsa")("jans") = "1" then
                        jobans = 1
                        jobansChk = "CHECKED"
					    jobansVal = " - Jobansvarlig"
                        else
				        jobansVal = ""
				        jobansChk = ""
				        jobans = 0
                        end if
				
                end if
				
               jobans1Val = jobans

				'*** Jobejer jobans 2 ***
				'if len(request("FM_jobans2")) <> 0 then
				if len(request("FM_kundejobans_ell_alle")) <> 0 then
                jobans2 = request("FM_jobans2")
					select case cdbl(jobans2)
					case 1 
					jobansVal2 = " - Jobejer"
					jobansChk2 = "CHECKED"
					case else
					jobansChk2 = ""
					end select

                     response.cookies("tsa")("jans2") = jobans2
				else
                     if request.cookies("tsa")("jans2") = "1" then
                     jobansVal2 = " - Jobejer"
					jobansChk2 = "CHECKED"
                    jobans2 = 1
                    else
                    jobansChk2 = ""
				    jobans2 = 0
                    end if
				end if

                '*** Jobmedansvarlige 3 ***
				'if len(request("FM_jobans3")) <> 0 then
				if len(request("FM_kundejobans_ell_alle")) <> 0 then
                jobans3 = request("FM_jobans3")
					select case cdbl(jobans3)
					case 1 
					jobansVal3 = " - Jobmedansvarlig 1-3"
					jobansChk3 = "CHECKED"
					case else
					jobansChk3 = ""
					end select

                      response.cookies("tsa")("jans3") = jobans3
				else
                      if request.cookies("tsa")("jans3") = "1" then
				      jobansVal3 = " - Jobmedansvarlig 1-3"
					  jobansChk3 = "CHECKED"
                      jobans3 = 1
                      else
                      jobansChk3 = ""
				      jobans3 = 0
                      end if
				end if


                '*** Salgsansv ***
				if len(request("FM_kundejobans_ell_alle")) <> 0 then
                'if len(request("FM_jobans")) <> 0 then
				salgsans = request("FM_salgsans")
					select case cdbl(salgsans)
					case 1 
					salgsansChk = "CHECKED"
					salgsansVal = " - Salgsansvarlig 1-5"
					case else
					salgsansVal = ""    
				    salgsansChk = ""
					end select

                    response.cookies("tsa")("salgsans") = salgsans

				else
                        if request.cookies("tsa")("salgsans") = "1" then
                        salgsans = 1
                        jobansChk = "CHECKED"
					    jobansVal = " - Jobansvarlig"
                        else
				        jobansVal = ""
				        jobansChk = ""
				        salgsans = 0
                        end if
				
                end if
				
               salgsansVal = salgsans

                
	
	
	else
	
	kundejobansCHK1 = ""
	kundejobansCHK0 = "CHECKED"
	kontakt_keyaccountVAL = "Kontakt"
	
	kansVal = ""
	kundeansChk = ""
	kundeans = 0
	jobansChk = ""
	jobansChk2 = ""
    jobansChk3 = ""
	jobans = 0
	jobans2 = 0
    jobans3 = 0
	jobansVal = ""
	jobansVal2 = ""
    jobansVal3 = ""
                    
    salgsansChk = ""
	salgsans = 0
	salgsansVal = ""
	end if		
	
	
	end function
	
	
	public strAftKidSQLkri
	public projgrp_medarb1, projgrp_medarb2, jobAnsSQLkri, jobAns2SQLkri, jobAns3SQLkri, progrp, strMedarbPrgSQL, ugeAflsMidKri
	public fakmedspecSQLkri, kundeAnsSQLKri, progrp_medarb, medarbSQlKri, medarbigrp, selmedarb
	
    function xmedarbogprogrp(side)
	
	projgrp_medarb1 = "CHECKED"
	jobAnsSQLkri = "" 
	fakmedspecSQLkri = ""
	kundeAnsSQLKri = ""
	medarbigrp = ""
	
                   
	
	
	if len(request("FM_radio_projgrp_medarb")) <> 0 then
	
	    'Response.Write "-her-"
		
		progrp_medarb = request("FM_radio_projgrp_medarb") '1
		
		select case side
		'case "oms"
		'medarbSQlKri = " tmnr <> 0 " '** her
		case "fomr" '"timtot", 
		medarbSQlKri = "m.mid <> 0"
		case "matstat"
		medarbSQlKri = " usrid <> 0 "
		end select
		
		if cint(request("FM_radio_projgrp_medarb")) <> 2 then '** != <> Projektgrupper (Medarbjeder valgt)
		
		
		
		projgrp_medarb1 = "CHECKED"
		projgrp_medarb2 = ""
		
			        if len(request("FM_medarb")) <> 0 AND request("FM_medarb") <> "0" then
			        selmedarb = request("FM_medarb")
        			
					        select case side
					        'case "oms"
					        'medarbSQlKri = "Tmnr = "& selmedarb '** her
					        case "fomr" '"timtot", 
					        medarbSQlKri = " m.mid = "&selmedarb
					        case "matstat"
					        medarbSQlKri = " usrid = "& selmedarb 
					        end select
        					
        					
			        jobAnsSQLkri = " OR (jobans1 = "& selmedarb &") " 
			        jobAns2SQLkri = " OR (jobans2 = "& selmedarb &") " 
			        jobAns3SQLkri = " OR (jobans3 = "& selmedarb &" OR jobans4 = "& selmedarb &" OR jobans5 = "& selmedarb &") " 
                    fakmedspecSQLkri = " AND (fms.mid = "& selmedarb &")" 
			        kundeAnsSQLKri = " kundeans1 = "& selmedarb &" OR kundeans2 = "& selmedarb &""
			        ugeAflsMidKri = " (u.mid = "& selmedarb &")"
			        else
        			
			        selmedarb = 0
        					
					        select case side
					        'case "oms"
					        'medarbSQlKri = "Tmnr <> "& selmedarb '** her
					        case "fomr" '"timtot", 
					        medarbSQlKri = " m.mid <> "& selmedarb
					        case "matstat"
					        medarbSQlKri = " usrid <> "& selmedarb &""
					        end select
        					
			        jobAnsSQLkri = ""
			        jobAns2SQLkri = ""
                    jobAns3SQLkri = ""
			        fakmedspecSQLkri = ""
			        kundeAnsSQLKri = " (Kundeans1 <> -1 AND kundeans2 <> -1) "
			        ugeAflsMidKri = " (u.mid <> "& selmedarb &")"
			        end if
			
			'** Er der tilvalgt kundeansvarlige ?? ***
			if cint(kundeans) = 1 then
			kundeAnsSQLKri = kundeAnsSQLKri
			else
			kundeAnsSQLKri = " (Kundeans1 <> -1 AND kundeans2 <> -1) "
			end if
			
			'Response.write kundeAnsSQLKri & "<br>"
			'Response.flush
			
		
		
		else 
		'*** Projektgruppe valgt
		
		
		
		if len(request("FM_progrupper")) <> 0 AND request("FM_progrupper") <> 0 then 
		progrp = request("FM_progrupper")
		else
		progrp = 10 '** alle gruppen
		end if
		
		projgrp_medarb1 = ""
		projgrp_medarb2 = "CHECKED"
		
		intMedArbVal = Split(progrp, ", ")
		For b = 0 to Ubound(intMedArbVal)
			
			strMedarbPrgSQL = strMedarbPrgSQL & " OR ProjektgruppeId = "& intMedArbVal(b) 
			
		Next
		
			
			
					select case side
					'case "oms"
					'medarbSQlKri = " tmnr = 0 " '** her
					case "fomr" '"timtot", 
					medarbSQlKri = " m.mid = 0 "
					case "matstat"
					medarbSQlKri = " usrid = 0 "
					end select
					
					ugeAflsMidKri =  " u.mid = 0 " 
			
			
			selmedarb = 0 'request("FM_medarb")
			'**** Henter medarbejdere i valgte projektgruppe **'
			strSQL = "SELECT Mid, Mnavn, Mnr, ProjektgruppeId, MedarbejderId, mansat, init FROM progrupperelationer "_
			&"LEFT JOIN medarbejdere ON (mid = MedarbejderId) WHERE (ProjektgruppeId = 0 "& strMedarbPrgSQL &") ORDER BY Mnavn"
			
			'Response.write strSQL
			ja = 0
			oRec.open strSQL, oConn, 3
			while not oRec.EOF 
			
					
					select case side
					'case "oms"
					'medarbSQlKri = medarbSQlKri & " OR Tmnr = "& oRec("mid")
					case "fomr" '"timtot", 
					medarbSQlKri = medarbSQlKri & " OR m.mid = "& oRec("mid")
					case "matstat"
					medarbSQlKri = medarbSQlKri &" OR usrid = "& oRec("mid")
					end select
					
			        medarbigrp = medarbigrp & oRec("mnavn") & " ("& oRec("mnr") &") <br>" 
			
			jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& oRec("mid")  
			jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& oRec("mid") 
			jobAns3SQLkri = jobAns3SQLkri & " OR jobans3 = "& oRec("mid") & " OR jobans4 = "& oRec("mid")  & " OR jobans4 = "& oRec("mid") 
            fakmedspecSQLkri = fakmedspecSQLkri & " OR fms.mid = "& oRec("mid")
			kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& oRec("mid") &" OR kundeans2 = "& oRec("mid") &""
			ugeAflsMidKri = ugeAflsMidKri & " OR u.mid = " & oRec("mid")
			
			ja = ja + 1
			oRec.movenext
			wend
			oRec.close
			
			ugeAflsMidKri = "("& ugeAflsMidKri &")"
			
			len_jobAnsSQLkri = len(jobAnsSQLkri)
			if ja >= 1 then
			right_jobAnsSQLkri = right(jobAnsSQLkri, len_jobAnsSQLkri - 3)
			jobAnsSQLkri =  "OR (" & right_jobAnsSQLkri&")"
			end if	
			
			len_jobAns2SQLkri = len(jobAns2SQLkri)
			if ja >= 1 then
			right_jobAns2SQLkri = right(jobAns2SQLkri, len_jobAns2SQLkri - 3)
			jobAns2SQLkri =  "OR (" & right_jobAns2SQLkri&")"
			end if	
			
            len_jobAns3SQLkri = len(jobAns3SQLkri)
			if ja >= 1 then
			right_jobAns3SQLkri = right(jobAns3SQLkri, len_jobAns3SQLkri - 3)
			jobAns3SQLkri =  "OR (" & right_jobAns3SQLkri&")"
			end if	

			len_fakmedspecSQLkri = len(fakmedspecSQLkri)
			if ja >= 1 then
			right_fakmedspecSQLkri = right(fakmedspecSQLkri, len_fakmedspecSQLkri - 3)
			fakmedspecSQLkri =  "AND (" & right_fakmedspecSQLkri&")"
			end if	
			
			len_kundeAnsSQLKri = len(kundeAnsSQLKri)
			if ja >= 1 then
			right_kundeAnsSQLKri = right(kundeAnsSQLKri, len_kundeAnsSQLKri - 3)
			kundeAnsSQLKri =  " (" & right_kundeAnsSQLKri &") "
			end if	
			
			'** Er der tilvalgt kundeansvarlige ?? ***
			if cint(kundeans) = 1 then
			kundeAnsSQLKri = kundeAnsSQLKri
			else
			kundeAnsSQLKri = " (Kundeans1 <> -1 AND kundeans2 <> -1) "
			end if
			
			
			'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"& kundeAnsSQLKri
			'Response.flush
			
		end if
	
	else
	
	'Response.Write "selmedarb kuke" & selmedarb
	
				
					select case side
					'case "oms"
					'medarbSQlKri = " tmnr = " & selmedarb '<> 0 " '** Her
					'selmedarb = selmedarb '0
					case "fomr" '"timtot", 
					medarbSQlKri = " m.mid = " & selmedarb
					selmedarb = selmedarb
					case "matstat"
					medarbSQlKri = " usrid = " & selmedarb 
					selmedarb = selmedarb 
					end select
					
					ugeAflsMidKri = " u.mid = " & selmedarb 
				
	
	jobAnsSQLkri = ""
	jobAns2SQLkri = ""
    jobAns3SQLkri = ""
	fakmedspecSQLkri = ""
	
	'*** Hvis alle medarbejdere er valgt vises altid alle kunder uanset om "Vis kun for Kundeansvarlige" er sl�et til.
	kundeAnsSQLKri = " (Kundeans1 <> -1 AND kundeans2 <> -1) "
	progrp_medarb = 0
	
	end if
	
	
	end function
	
	
	
	public jobSogVal, jobSogValPrint, jobid, aktNavnSogVal
	function jobsog()
	
	'if lto = "essens" then
	'Response.Write "request(FM_jobsog): "& request("FM_jobsog") & " jobid = "& request("FM_job")
	'end if

    if len(trim(request("vis_fakbare_res"))) <> 0 then

        if len(trim(request("FM_aktnavnsog"))) <> 0 AND cint(vis_aktnavn) = 1 then
        aktNavnSogVal = trim(request("FM_aktnavnsog"))
        else
        aktNavnSogVal = ""
        end if

    else

        if request.cookies("cc_vis_fakbare_res")("aktn_val") <> "" AND cint(vis_aktnavn) = 1 then
        aktNavnSogVal = request.cookies("cc_vis_fakbare_res")("aktn_val")
        else
        aktNavnSogVal = ""
        end if
        

    end if


    response.cookies("cc_vis_fakbare_res")("aktn_val") = aktNavnSogVal


	
	if len(trim(request("FM_jobsog"))) <> 0 then
		
	jobSogVal = trim(request("FM_jobsog"))
	
	
			
			'*** fra Print **
			if instr(jobSogVal, "99ogprocent99") <> 0 then
			jobSogVal = replace(jobSogVal, "99ogprocent99", "%")
			else
			jobSogVal = jobSogVal 
			end if
			
			if instr(jobSogVal, "%") <> 0 then
			jobSogValPrint = replace(jobSogVal, "%", "99ogprocent99")
			else
			jobSogValPrint = jobSogVal 
			end if
		
		
		    'Response.Write "jobSogVal:" & jobSogVal
	        'Response.flush	
			
			
			'*** Brug det s�gte jobnr til at vise enkelt job og aktivitets detaljer **'
			call erDetInt(jobSogVal) 
			if isInt > 0 OR instr(jobSogVal, ".") <> 0 OR instr(jobSogVal, ",") <> 0 then
			jobid = 0
			else
			    
			    jobid = 0
			    
			    'jnrThis = trim(request("FM_jobsog"))
			    'strSQLj = "SELECT id FROM job WHERE jobnr = " & jnrThis
			    'oRec5.open strSQLj, oConn, 3
			    'if not oRec5.EOF then
			    'jobid = oRec5("id")
			    'end if
			    'oRec5.close
			     
			
			
			end if
	        
	        isInt = 0
	else
	jobSogVal = ""
			
			'** valgt
			if len(request("FM_job")) <> 0 then
			jobid = request("FM_job")
			else
			jobid = 0
			end if

            'if session("mid") = 1 then
            '       Response.write "Jobid"& jobid &"="& request("FM_job") 
            'end if
			
	end if
	
	end function
	
	
	public showthisMedarb
	function ekstraSogkri(useSogKri, useSogKriAfs, useSogKriGk, moreorless,timeKri,saldoKri,visning)
	
	if visning = 14 then
	
	
	   
	       if cint(useSogKri) = 1 then
	       'moreorless
	       'timeKri
	       'saldoKri
	           'Response.Write "<br>moreorless: "& moreorless &" normtime_lontime: "& normtime_lontime/60 &" balRealNormtimer: "& balRealNormtimer &" balRealLontimer: "& balRealLontimer &" timekri: "&  cdbl(timeKri) & " saldoKri "& saldoKri & " ugegodkendt: "& ugegodkendt &"<br>"
	           select case saldoKri
	           case 0 
	                  
	                  if cint(moreorless) = 0 then
	                    'Response.Write "her"
	                    if cdbl(balRealNormtimer) > cdbl(timeKri) then
	                    showthisMedarb = 1
	                    else
	                    showthisMedarb = 0
	                    end if
	                    
	                  else
	                    
	                    if cdbl(balRealNormtimer) < cdbl(timeKri) then
	                    showthisMedarb = 1
	                    else
	                    showthisMedarb = 0
	                    end if
	                  
	                  end if
	                  
	                  
	           case 1
	                     
	                    if cint(moreorless) = 0 then
	                    'Response.Write "<br>normtime_lontime: "& cdbl(normtime_lontime/60) & "timeKri" & timeKri
	                    if cdbl(normtime_lontime/60) > cdbl(timeKri) then
	                    showthisMedarb = 1
	                    else
	                    showthisMedarb = 0
	                    end if
	                    
	                  else
	                    
	                    if cdbl(normtime_lontime) < cdbl(timeKri) then
	                    showthisMedarb = 1
	                    else
	                    showthisMedarb = 0
	                    end if
	                  
	                  end if
	           case 2
	                     
	                     if cint(moreorless) = 0 then
	                    
	                    if cdbl(balRealLontimer) > cdbl(timeKri) then
	                    showthisMedarb = 1
	                    else
	                    showthisMedarb = 0
	                    end if
	                    
	                  else
	                    
	                    if cdbl(balRealLontimer) < cdbl(timeKri) then
	                    showthisMedarb = 1
	                    else
	                    showthisMedarb = 0
	                    end if
	                  
	                  end if
    	       end select 
	       end if 


           '** Vis kun afsluttet ***'
           if cint(useSogKriAfs) = 1 then
                
                
                if cint(showAfsuge) = 1 then
                showthisMedarb = 1
	            else
	            showthisMedarb = 0
                end if

           end if

           '** Vis kun godkendte ***'
           if cint(useSogKriGk) = 1 then

                if cint(ugegodkendt) = 0 then
                showthisMedarb = 1
	            else
	            showthisMedarb = 0
                end if
           end if


	       
	   end if 

                      'Response.Write "<br>moreorless: "& moreorless &" normtime_lontime: "& normtime_lontime/60 &" balRealNormtimer: "& balRealNormtimer &" balRealLontimer: "& balRealLontimer &" timekri: "&  cdbl(timeKri) & " saldoKri "& saldoKri & " ugegodkendt: "& ugegodkendt &" showthisMedarb: "& showthisMedarb &" <br>"
	         
	
	end function
	











  
	%>