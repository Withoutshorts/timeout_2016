<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%thisfile = "joblist" %>

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="inc/job_inc.asp"-->

<%
if Request.Form("AjaxUpdateField") = "true" then

Select Case Request.Form("control")


    case "simpleliste_cookie"

        
        response.Cookies("2015")("vis_simplejobliste") = request("simpleliste")


    case "statusfil_cookie"

        response.Cookies("2015")("statusfilt_open") = request("open")


case "FN_getKundeKperslisten"
              
              dim kundeKpersArr
              redim kundeKpersArr(500)
              
              if len(trim(request.Form("cust2"))) <> 0 AND request.Form("cust2") <> 0 then
              kidThis = request.Form("cust2")
              else
              kidThis = -1
              end if

              if len(trim(request("jq_kundekpers"))) <> 0 then
              jq_kundekpers = request("jq_kundekpers")
              else
              jq_kundekpers = 0
              end if


                strSQL = "SELECT navn, id, titel, email FROM kontaktpers WHERE kundeid = "& kidThis
			    
                'Response.write "<option>"& strSQL& "</option>"
                'Response.end
                

                kundeKpersArr(0) = "<option value='0'>"&job_txt_002&"</option>"
                z = 1

                oRec.open strSQL, oConn, 3
			   
			    while not oRec.EOF
				
			    
                if cint(jq_kundekpers) = oRec("id") then
                kpsSEL = "SELECTED"
                else
                kpsSEL = ""
                end if

                if len(trim(oRec("email"))) <> 0 then
                eTxt = ", " & oRec("email")
                else
                eTxt = ""
                end if


                if len(trim(oRec("titel"))) <> 0 then
                tTxt = oRec("titel")
                else
                tTxt = ""
                end if
                

            	
			kundeKpersArr(z) = "<option value='"&oRec("id")&"' "&kpsSEL&">"&oRec("navn")&" "& tTxt &" "& eTxt &"</option>"
            

            'Response.flush
		    
			z = z + 1
			
			oRec.movenext
			wend
			oRec.close

            if z = 1 then
            kundeKpersArr(z) = "<option value='0'>"&job_txt_003&"</option>"
            end if          


               for z = 0 to UBOUND(kundeKpersArr)
              '*** ��� **'
              call jq_format(kundeKpersArr(z))
              kundeKpersArr(z) = jq_formatTxt
	          
	          Response.Write kundeKpersArr(z)
              next
		
end select

Response.end
end if    
    
%>



<script src="js/job_list_jav_6.js" type="text/javascript"></script>

<%
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if
        
    func = request("func")
    select case func
    case "dbred"
        'rediger
    case "dbopr"
        'opret




    case "jobstatusUpdate"


        jobid = split(request("FM_selectedJobids_status"), ",")
        newJobstatus = request("FM_jobstatus")

        for t = 0 to UBOUND(jobid)

            if jobid(t) <> "" then
                'response.Write "<br> JS " & newJobstatus 
            
                call lukkedato(jobid(t), newJobstatus)

                strSQL = "UPDATE job SET jobstatus = "& newJobstatus & " WHERE id = "& jobid(t)
                'response.Write "<br> " & strSQL
                oConn.execute(strSQL)
            end if

        next

        'response.Write "DONE !"
        response.Redirect "job_list.asp?jobnr_sog="&request("jobnr_sog")&"&FM_kunde="&request("FM_kunde")


     case "easyregUpdate"


        jobid = split(request("FM_selectedJobids_easyreg"), ",")
        
        if len(trim(request("FM_tilfojeasyreg"))) <> 0 then
        easyReg = request("FM_tilfojeasyreg")
        else
        easyReg = 0
        end if

        for t = 0 to UBOUND(jobid)

                if cint(easyReg) <> 0 AND jobid(t) <> "" then
                            
                            
                if cint(easyReg) = 1 then
                easyRegVal = 1
                else
                easyRegVal = 0
                end if

                            
                                    '**** S�tter EASYreg aktiv p� aktiviteter for alle medarbejdere (hvis der findes EASY regaktiviter) ****'
                           
                                    strSQLea = "UPDATE aktiviteter SET easyreg = "& easyRegVal &" WHERE job = "& jobid(t)
                                    'Response.write strSQLea
                                    'Response.flush

				                    oConn.execute(strSQLea)
				                
                                    if cint(easyReg) = 1 then
                                    strSQLtreguse = "UPDATE timereg_usejob SET easyreg = " & jobid(t) & " WHERE jobid = " & jobid(t)
                                    else '= 2
                                    strSQLtreguse = "UPDATE timereg_usejob SET easyreg = 0 WHERE jobid = " & jobid(t)
                                    end if

	   	                            oConn.execute(strSQLtreguse)
				              


                end if

        next

        'response.Write "DONE !"
        response.Redirect "job_list.asp?jobnr_sog="&request("jobnr_sog")&"&FM_kunde="&request("FM_kunde")

    case "sletok"


         jobid = split(request("FM_selectedJobids_delete"), ",")
         kt = request("kt") 'slet kun timer
        
       

        for t = 0 to UBOUND(jobid)

             if jobid(t) <> "" then

            '*** Her slettes et job, dets aktiviteter og de indtastede timer p� jobbet ***
	        strSQL = "SELECT id, navn FROM aktiviteter WHERE job = "& jobid(t) &"" 
	        
	
	        oRec.open strSQL, oConn, 3
	        while not oRec.EOF 
		
		        '*** Inds�tter i delete historik ****'
	            call insertDelhist("akt", oRec("id"), 0, oRec("navn"), session("mid"), session("user"))
		
		        '** Slet eller nulstil? **'
		        if kt = "0" OR kt = "2" then
		        oConn.execute("DELETE FROM aktiviteter WHERE id = "& oRec("id") &"")
		        end if
		        oConn.execute("DELETE FROM timer WHERE TAktivitetId = "& oRec("id") &"")
		
	        oRec.movenext
	        wend
	        oRec.close
	
	        '*** Sletter ressource timer p� job ****
	        oConn.execute("DELETE FROM ressourcer WHERE jobid = "& jobid(t) &"")
	
	        '*** Sletter ressource_md timer p� job ****
	        oConn.execute("DELETE FROM ressourcer_md WHERE jobid = "& jobid(t) &"")
	
	        '*** Sletter fra Guiden aktive job timer p� job ****
	        oConn.execute("DELETE FROM timereg_usejob WHERE jobid = "& jobid(t) &"")
	
	        '*** Sletter materialeforbrug ****
	        oConn.execute("DELETE FROM materiale_forbrug WHERE jobid = "& jobid(t) &"")

	
	      
	
	
	        strSQL = "SELECT id, jobnavn, jobnr FROM job WHERE id = "& jobid(t) &"" 
	        oRec.open strSQL, oConn, 3
	        if not oRec.EOF then
		
		        '*** Inds�tter i delete historik ****'
	            call insertDelhist("job", jobid(t), oRec("jobnr"), oRec("jobnavn"), session("mid"), session("user"))
		
	
	        end if
	        oRec.close
	
	
	
	
	        'Response.flush
	
	        if kt = "0" then
	        oConn.execute("DELETE FROM job WHERE id = "& jobid(t) &"")
	        end if


            end if


        next



         response.Redirect "job_list.asp?jobnr_sog="&request("jobnr_sog")&"&FM_kunde="&request("FM_kunde")




    case "jobslutdatoUpdate"

        jobid = split(request("FM_selectedJobids_enddate"), ",")
        newEnddate = request("FM_newEnddate")
        newEnddate = year(newEnddate) &"-"& month(newEnddate) &"-"& day(newEnddate)

        for t = 0 to UBOUND(jobid)

            if jobid(t) <> "" AND isDate(newEnddate) = true then

                strSQL = "UPDATE job SET jobslutdato = '"& newEnddate &"' WHERE id = "& jobid(t)
                'response.Write "<br>" & strSQL
                oConn.execute(strSQL)

            end if

        next

        response.Redirect "job_list.asp?jobnr_sog="&request("jobnr_sog")&"&FM_kunde="&request("FM_kunde")


    case "foromUpdate"

        jobid = split(request("FM_selectedJobids_forOm"), ",")

        if len(trim(request("FM_formr_opdater_akt"))) <> 0 then
        fomrJaAkt = request("FM_formr_opdater_akt")
        else
        fomrJaAkt = 0
        end if

        for_faktor = 1

        fomrArr = split(request("FM_fomr"), ",")


        for t = 0 to UBOUND(jobid)

            if jobid(t) <> "" then

                '*** nulstiller job (IKKE akt) ****'
                strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& jobid(t)  & " AND for_aktid = 0"
                oConn.execute(strSQLfor)

                for afor = 0 to UBOUND(fomrArr)

                    fomrArr(afor) = replace(fomrArr(afor), "#", "")

                        if fomrArr(afor) <> 0 then

                            strSQLfomri = "INSERT INTO fomr_rel "_
                            &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                            &" VALUES ("& fomrArr(afor) &", "& jobid(t) &", 0, "& for_faktor &")"

                            oConn.execute(strSQLfomri)

                            'fomrArrLink = fomrArrLink & ",#"& fomrArr(afor) & "#" 
            

                            '**** IKKE MERE 11.3.2015 *****'
                            '*** Altid Sync aktiviteter ved multiopdater ****'
                            if cint(fomrJaAkt) = 1 then
                            strSQLa = "SELECT id FROM aktiviteter WHERE job = "& jobid(t)
                            oRec3.open strSQLa, oConn, 3
                            while not oRec3.EOF 

                            'response.Write "<br>" & strSQLa &" "& oRec3("id")

                                    '*** nulstiller form p� akt () ****'
                                strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& jobid(t)  & " AND for_aktid = "& oRec3("id")
                                'response.Write "<br>" & strSQLfor
                                oConn.execute(strSQLfor)

                                strSQLfomrai = "INSERT INTO fomr_rel "_
                                &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                &" VALUES ("& fomrArr(afor) &", "& jobid(t) &","& oRec3("id") &", "& for_faktor &")"
                                'response.Write "<br> " & strSQLfomrai
                                oConn.execute(strSQLfomrai)

                            oRec3.movenext
                            wend
                            oRec3.close
                            end if

                        end if

                next

            end if


        next


        response.Redirect "job_list.asp?jobnr_sog="&request("jobnr_sog")&"&FM_kunde="&request("FM_kunde")

    case else


    
    call simplejobfelter()
    


    'fomrArr = split(request("FM_fomr"), ",")
    if len(trim(request("FM_fomr"))) <> 0 then
    strFomr_rel = request("FM_fomr")
    strFomr_rel = replace(strFomr_rel, "X234", "#")
    else
    strFomr_rel = "#0#"
    end if


    

    '*** Sidste redigeret job id 
    if request.cookies("2015")("lastjobid") <> "" then
    lastjobid = request.cookies("2015")("lastjobid")
    else
    lastjobid = 0
    end if

    


    if len(trim(request("FM_kunde"))) <> 0 then
    FM_kunde = request("FM_kunde")
    else

        if request.cookies("2015")("fm_kunde") <> "" then
        FM_kunde = request.cookies("2015")("fm_kunde")
        else
        FM_kunde = 0
        end if

    
    end if

    response.cookies("2015")("fm_kunde") = FM_kunde

    select case FM_kunde
    case 0
    FM_kundeSqlFilter = " jobknr <> 0"
    case else
    FM_kundeSqlFilter = " jobknr = "& FM_kunde 'AND skal p�
    end select

    if len(trim(request("status_filter"))) <> 0 then
	statusFilt = request("status_filter")
    else
        if request.cookies("2015")("status_filter") <> "" then
        statusFilt = request.cookies("2015")("status_filter")
        else
        statusFilt = "1"
        end if
    end if

    response.cookies("2015")("status_filter") = statusFilt

    chk1 = ""
	chk2 = ""
	chk3 = ""
	chk4 = ""
	chk5 = ""

    varStatusFilt = " AND (jobstatus = -2"
    for f = 0 to 5
        'Response.write instr(filt, f) & "<br>.."

        if instr(statusFilt, f) <> 0 then
        varStatusFilt = varStatusFilt & " OR jobstatus = "& f
            
            select case f
            case 0
            chk0 = "CHECKED"
            case 1
            chk1 = "CHECKED"
            case 2
            chk2 = "CHECKED"
            case 3
            chk3 = "CHECKED"
            case 4
            chk4 = "CHECKED"
            case 5
            chk5 = "CHECKED"
            end select
        
        
        end if

    next

    varStatusFilt = varStatusFilt & ")"


    if len(request("FM_medarb_jobans")) <> 0 then
	medarb_jobans = request("FM_medarb_jobans")
	response.cookies("2015")("jobans") = medarb_jobans
	else
	    if request.cookies("2015")("jobans") <> "" then
	    medarb_jobans = request.cookies("2015")("jobans")
	    else

            select case lto
            case "epi2017"
            medarb_jobans = session("mid")
            case else
	        medarb_jobans = 0
	        end select

       end if
	end if

    if cdbl(medarb_jobans) = 0 then
	jobansKri = ""
	else

        select case lto
        case "hestia", "epi2017"
        jobansKri = " AND (jobans1 = " & medarb_jobans &" OR jobans2 = " & medarb_jobans &")"
        case else
	    jobansKri = " AND jobans1 = " & medarb_jobans 
        end select

	end if


    if len(trim(request("FM_prjgrp"))) <> 0 then
    prjgrp = request("FM_prjgrp")
    response.Cookies("2015")("prjgrp") = prjgrp
    else
        if request.Cookies("2015")("prjgrp") <> "" then
        prjgrp = request.Cookies("2015")("prjgrp")
        else
        prjgrp = "10" '** Alle
        end if
    end if

    select case prjgrp
    case "10", "-1"
    prjgrpSQL = ""
    case else
    prjgrpSQL = " AND projektgruppe1 = " & prjgrp & " OR projektgruppe2 = " & prjgrp & " OR projektgruppe3 = " & prjgrp & " OR projektgruppe4 = " & prjgrp & " OR projektgruppe5 = " & prjgrp & " OR projektgruppe6 = " & prjgrp & " OR projektgruppe7 = " & prjgrp & " OR projektgruppe8 = " & prjgrp & " OR projektgruppe9 = " & prjgrp & " OR projektgruppe10 = " & prjgrp
    end select



    if len(trim(request("FM_sogakt"))) <> 0 then
		sogakt = 1
		sogaktCHK = "CHECKED"
	else
	    sogakt = 0
	    sogaktCHK = ""		
    end if

    if len(trim(request("FM_hd_kpers"))) <> 0 then
    hd_kpers = request("FM_hd_kpers")
    else
    hd_kpers = -1
    end if
       
    '****** S�GNING **************

    'Fra menu
    if len(trim(request("hidelist"))) <> 0 then
    hidelist = request("hidelist")

     if hidelist = 2 then 'fra menu - nustiller helt 1: fra rediger slet, hsuker s�gnng, men viser ikke liste f�r klik s�g
     response.cookies("2015")("lastjobid") = 0
     response.cookies("2015")("jobsog") = ""
     lastjobid = 0
     end if

    else
    hidelist = 0
    end if


    if len(trim(request("FM_kunde"))) <> 0 OR hidelist = 0 then


       
       

                if len(trim(request("FM_kunde"))) <> 0 then 'har slettet sin s�gning p� joblisten

                jobnr_sog = request("jobnr_sog")

                else 'kommer retur fra rediger job

                    if request.cookies("2015")("jobsog") <> "" then
                    jobnr_sog = request.cookies("2015")("jobsog")

                        'if jobnr_sog = "%" then
                        'jobnr_sog = ""
                        'end if

                    else
                    jobnr_sog = ""
                    end if

                 end if

          

        response.cookies("2015")("jobsog") = jobnr_sog
        visliste = 1

    else
            if request.cookies("2015")("jobsog") <> "" then
           
                    jobnr_sog = request.cookies("2015")("jobsog")
            
                        'if jobnr_sog = "%" then
                        'jobnr_sog = ""
                        'end if

            else
            jobnr_sog = ""
            end if

            visliste = 0
    end if




    if ((len(jobnr_sog) <> 0 AND jobnr_sog <> "-1") OR (cint(visliste) = 1)) AND hidelist = 0 then
        jobnr_sog = jobnr_sog

        if cint(sogakt) = 1 then
			sogeKri = sogeKri & " AND (a.navn LIKE '%"& jobnr_sog &"%' "

		else
            if instr(jobnr_sog, ">") > 0 OR instr(jobnr_sog, "<") > 0 OR instr(jobnr_sog, "--") > 0 OR instr(jobnr_sog, ";") > 0 then
           
                    if instr(jobnr_sog, ">") > 0 then
                    sogeKri = sogeKri &" AND (j.jobnr > "& replace(trim(jobnr_sog), ">", "") &" "
                    end if

                    if instr(jobnr_sog, "<") > 0 then
                    sogeKri = sogeKri &" AND (j.jobnr < '"& replace(trim(jobnr_sog), "<", "") &"' "
                    end if

                    if instr(jobnr_sog, "--") > 0 then
                    jobnr_sogArr = split(jobnr_sog, "--")
               
                    for t = 0 to 1
                
                    if t = 0 then
                    jobSogKriA = jobnr_sogArr(0)
                    else
                    jobSogKriB = jobnr_sogArr(1)
                    end if
                
                    next

                    sogeKri = sogeKri &" AND (j.jobnr BETWEEN '"& trim(jobSogKriA) &"' AND '"& trim(jobSogKriB) &"'"
                    end if


                    if instr(jobnr_sog, ";") > 0 then
            
                    sogeKri = " AND (j.jobnr = '-1' "

                    jobnr_sogArr = split(jobnr_sog, ";")
               
                    for t = 0 to UBOUND(jobnr_sogArr)
                     sogeKri = sogeKri &" OR j.jobnr = '"& trim(jobnr_sogArr(t)) &"'"
                    next

                              

                    end if

                else


                    select case lto
                    case "dencker", "intranet - local", "mpt"
                    jobnrSqlkr = "'%"& jobnr_sog &"%'"
                    case else
                    jobnrSqlkr = "'"& jobnr_sog &"'"
                    end select

                    select case lto
                    case "mpt"
                    jobnavnSqlkr = "'%"& jobnr_sog &"%'"
                    case else
                    jobnavnSqlkr = "'"& jobnr_sog &"%'"
                    end select


                sogeKri = " AND (j.jobnr LIKE "& jobnrSqlkr &" OR j.jobnavn LIKE "& jobnavnSqlkr &" OR j.id LIKE '"& jobnr_sog &"' OR Kkundenavn LIKE '"& jobnr_sog &"%' OR Kkundenr LIKE '"& jobnr_sog &"' OR rekvnr LIKE '"& jobnr_sog &"' OR project_tier = '"& jobnr_sog &"'"
		        	


                end if
            
            end if
			
			sogeKri = sogeKri & ") "
			
			
			show_jobnr_sog = jobnr_sog
			
		else

            'show_jobnr_sog = ""
            show_jobnr_sog = jobnr_sog
			jobnr_sog = "-1"
			sogeKri = " AND (j.id = -1) "
			
		end if





    '*** EASY REG **'
    if len(trim(request("FM_sogakt_easyreg"))) <> 0 then
			sogakt_easyreg = 1
			sogakt_easyregCHK = "CHECKED"
	else
	        sogakt_easyreg = 0
	        sogakt_easyregCHK = ""		
    end if	   

    '**** Vis kun med Easyreg ***'
    if cint(sogakt_easyreg) = 1 then
    sogeKri = sogeKri & " AND a.easyreg = 1 "
    end if



    '** Service aftale 
     aftfilt = request("aftfilt")

    select case aftfilt
	case "serviceaft"
	varFilt = varFilt & " AND serviceaft <> 0" 'Service aftaler.
	chk6 = "CHECKED"
	chk7 = ""
	chk8 = ""
	case "ikkeserviceaft"
	varFilt = varFilt & " AND serviceaft = 0" 'Ikke Service aftaler.
	chk6 = ""
	chk7 = "CHECKED"
	chk8 = ""
	case else
	varFilt = varFilt & "" 'Vis alle.
	chk6 = ""
	chk7 = ""
	chk8 = "CHECKED"
	end select


    '**** Brug datokriterie filer ****
    usedatoKri = "n"
    if len(request("FM_kunde")) <> 0 then 'Er der s�gt

	    if len(request("usedatokri")) <> 0 then

		    usedatoKri = "j"
		    datoKriJa = "CHECKED"
	    else
		    usedatoKri = "n"
		    datoKriJa = ""
        end if

        response.Cookies("job")("cusedatokri") = usedatoKri

    else

            usedatoKri = request.Cookies("job")("cusedatokri")
		
            if usedatoKri = "j" then
		    datoKriJa = "CHECKED"
		    else
		    usedatoKri = "n"
		    datoKriJa = ""
		    end if

    end if

    usedatokri_st_or_end = request("usedatokri_st_or_end")
    if len(trim((usedatokri_st_or_end))) <> 0 then
        usedatokri_st_or_end = request("usedatokri_st_or_end")
    else
        usedatokri_st_or_end = 0
    end if

    select case usedatokri_st_or_end
    case "1"
        usedatokri_st_or_end_1SEL = "SELECTED"
        usedatokri_st_or_end_0SEL = ""
        usedatokri_st_or_end_SQLkri = "jobstartdato"
    case else
        usedatokri_st_or_end_1SEL = ""
        usedatokri_st_or_end_0SEL = "SELECTED"
        usedatokri_st_or_end_SQLkri = "jobslutdato"
    end select

     if len(trim(request("aar"))) <> 0 then
            aar = request("aar")

                response.cookies("2015")("gktimer_aar") = aar

            else

                if request.cookies("2015")("gktimer_aar") <> "" then
                aar = request.cookies("2015")("gktimer_aar")
                else
                aar = "1-"& month(now) &"-"& year(now)
                end if

            end if

            if len(trim(request("aarslut"))) <> 0 then
            aarslut = request("aarslut")

            response.cookies("2015")("gktimer_aarslut") = aarslut

            else
   
                if request.cookies("2015")("gktimer_aarslut") <> "" then
                aarslut = request.cookies("2015")("gktimer_aarslut")
                else
                aarslut = dateAdd("m", 1, aar)
                end if

     end if


    'if usedatoKri = "j" then
	'varOrder = "jobslutdato DESC"
	'else
	'varOrder = "kkundenavn, jobnavn"
	'end if


    if usedatoKri = "j" then
    strStartDato = year(aar) & "/"& month(aar) & "/" & day(aar)
    strSlutDato = year(aarslut) & "/"& month(aarslut) & "/" & day(aarslut)
	datoKri = " AND "& usedatokri_st_or_end_SQLkri &" BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
	else
	datoKri = ""
	end if	

    call basisValutaFN()

    call browsertype()


    if len(trim(request("FM_simplelist"))) <> 0 then
	vis_simplejobliste = request("FM_simplelist")
    else
        if request.cookies("2015")("vis_simplejobliste") <> "" AND request.cookies("2015")("vis_simplejobliste") <> "undefined" then
        vis_simplejobliste = request.cookies("2015")("vis_simplejobliste")
        else
        vis_simplejobliste = 0
        end if
    end if

    if browstype_client = "ip" then
        vis_simplejobliste = 0
        simple_joblist_jobnavn = 1 
        simple_joblist_jobnr = 1
        simple_joblist_kunde = 0 
        simple_joblist_ans = 0 
        simple_joblist_salgsans = 0 
        simple_joblist_fomr = 0
        simple_joblist_prg = 0
        simple_joblist_stdato = 0 
        simple_joblist_sldato = 0
        simple_joblist_status = 0 
        simple_joblist_tidsforbrug = 0 
        simple_joblist_budgettid = 0
    end if


    if cint(vis_simplejobliste) = 1 then
        simpledisplay = "none"
        advanceddisplay = "normal"
        btntxt = "< Simple >"
    else
        simpledisplay = "normal"
        advanceddisplay = "none"
        btntxt = "< Advanced >"
    end if


    strstatus = ""
    if chk1 = "CHECKED" then
        strstatus = strstatus & " ("& job_txt_221 & ")"
    end if
    if chk2 = "CHECKED" then
        strstatus = strstatus & " ("& job_txt_222 & ")"
    end if
    if chk3 = "CHECKED" then
        strstatus = strstatus & " ("& job_txt_063 & ")"
    end if
    if chk4 = "CHECKED" then
        strstatus = strstatus & " ("& job_txt_224 & ")"
    end if
    if chk0 = "CHECKED" then
        strstatus = strstatus & " ("& job_txt_225 & ")"
    end if
    if chk5 = "CHECKED" then
        strstatus = strstatus & " ("& job_txt_629 &")" 
    end if

    ' Henter cookie om drop down menu skal v�re sl�et ud
    if request.Cookies("2015")("statusfilt_open") <> "" then
        statusfiltopen = request.Cookies("2015")("statusfilt_open")
    else
        statusfiltopen = 0
    end if

    response.Write "<input type='hidden' value='"&statusfiltopen&"' id='statusfilt_open' />"

    if cint(statusfiltopen) = 1 then
        statusdiv = "block"
    else
        statusdiv = "none"
    end if
    
%>





<%
    if (browstype_client <> "ip") then
        call menu_2014 
    else
        %><!--#include file="../inc/regular/top_menu_mobile.asp"--><%
        call mobile_header
    end if
%>

<script src="js/job_jav.js" type="text/javascript"></script>

<style>

    .advanced_filter:hover,
    .advanced_filter:focus {
    text-decoration: none;
    cursor: pointer;
    }

        /* The Modal (background) */
    .modal {
        display: none; /* Hidden by default */
        position: fixed; /* Stay in place */
        z-index: 1; /* Sit on top */
        padding-top: 100px; /* Location of the box */
        left: 0;
        top: 0;
        width: 100%; /* Full width */
        height: 100%; /* Full height */
        overflow: auto; /* Enable scroll if needed */
        background-color: rgb(0,0,0); /* Fallback color */
        background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
    }

    /* Modal Content */
    .modal-content {
        background-color: #fefefe;
        margin: auto;
        padding: 20px;
        border: 1px solid #888;
        width: 200px;
        height: 350px;
    }

    .picmodal:hover,
    .picmodal:focus {
    text-decoration: none;
    cursor: pointer;
}


 


</style>




<%
'select case lto
'case "mpt", "outz"
'wrpwidth = ""
'case "xintranet - local", "hestia"
'wrpwidth = "style='width:1250px;'"
'case "dencker"
'wrpwidth = "style='width:1500px;'"
'case else
'wrpwidth = "style='width:1600px;'"
'end select

wrpwidth = "style='width:85%;'"

if browstype_client = "ip" then
    wrpwidth = "style='width:100%;'"
end if
    
%>

<style>
    .multiselect {
  width: 200px;
}

.selectBox {
  position: relative;
}

.overSelect {
  position: absolute;
  left: 0;
  right: 0;
  top: 0;
  bottom: 0;
}

#checkboxes {
  display: none;
  border: 1px #dadada solid;
}

#checkboxes label {
  display: block;
}

#checkboxes label:hover {
   background-color: #efefef;
}
</style> 

 <script type="text/javascript" src="//cdnjs.cloudflare.com/ajax/libs/moment.js/2.8.4/moment.min.js"></script>
 <script type="text/javascript" src="//cdn.datatables.net/plug-ins/1.10.19/sorting/datetime-moment.js"></script>

<meta name="viewport" content='width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0'>

<%if (browstype_client <> "ip") then %>
<div class="wrapper">
    <div class="content">
<%end if %>
        <div class="container" <%=wrpwidth %>>
            <div class="portlet">
                <h3 class="portlet-title"><u><%=job_txt_456 %></u></h3>
                <div class="portlet-body">

                    <%if lto <> "care" then %>
                        <%if browstype_client <> "ip" then %>
                            <form action="../timereg/jobs.asp?func=opret&id=0" method="post">
                                <div class="row">
                                    <div class="col-lg-10">&nbsp;</div>
                                    <div class="col-lg-2">
                                        <button class="btn btn-sm btn-success pull-right"><b><%=job_txt_307 %> +</b></button><br />&nbsp;
                                    </div>
                                </div>
                            </form>
                        <%end if %>
                    <%else %>
                        <form action="jobs.asp?func=opret&id=0" method="post">
                            <div class="row">
                                <div class="col-lg-10">&nbsp;</div>
                                <div class="col-lg-2">
                                    <button class="btn btn-sm btn-success pull-right"><b><%=job_txt_307 %> +</b></button><br />&nbsp;
                                </div>
                            </div>
                        </form>
                    <%end if %>

                    <%
                                            
                    select case lto
                    case "mpt", "outz", "intranet - local"%>
                       <%if browstype_client <> "ip" then %>
                       <form action="jobs.asp?func=opret&jobid=0" method="post">
                            <div class="row">
                                <div class="col-lg-10">&nbsp;</div>
                                <div class="col-lg-2">
                                    <button class="btn btn-sm btn-warning pull-right"><b>NY <%=job_txt_307 %> +</b></button><br />&nbsp;
                                </div>
                            </div>
                        </form>
                        <%end if %>
                    <%end select %>

                    <%if browstype_client <> "ip" then %>
                    <div class="well">
                        <form action="job_list.asp?" method="post">       

                            <div class="row">
                                <div class="col-lg-4"><%=job_txt_208 %>:</div>
                                
                                <div class="col-lg-4"><%=job_txt_021 %>:</div>
                            </div>
                        
                            <div class="row">
                                <div class="col-lg-4"><input name="jobnr_sog" type="text" class="form-control input-small" value="<%=show_jobnr_sog %>" />
                                  
                                    <span style="font-size:9px; color:#999999;">(% = wildcard) <%=job_txt_209 %> ";" <%=job_txt_210 %>, ">", "<" <%=job_txt_211 %> "--" (<%=job_txt_212 %>) <%=job_txt_213 %></span><br />
                                    <input type="checkbox" name="FM_sogakt" <%=sogaktCHK %> /> <%=job_txt_349 &" "& job_txt_232 %><br />&nbsp;</div>

                                <div class="col-lg-4">
                                    <select class="form-control input-small" id="FM_kunde" name="FM_kunde" onchange="submit();">
                                        <option value="0"><%=job_txt_323 %></option>

                                        <%
                                            strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
				                            oRec.open strSQL, oConn, 3
				                            while not oRec.EOF
                                            
                                            if cdbl(oRec("Kid")) = cdbl(FM_kunde) then
                                            FM_kundeSEL = "SELECTED"
                                            else
                                            FM_kundeSEL = ""
                                            end if 
                                        %>
                                        <option value="<%=oRec("Kid") %>" <%=FM_kundeSEL %>><%=oRec("Kkundenavn")%>&nbsp(<%=oRec("Kkundenr") %>)</option>
                                        <%
                                            oRec.movenext
                                            wend
                                            oRec.close 
                                        %>
                                    </select>
                                </div>                           
                            </div> 
                            
                            <div class="row">
                                <div class="col-lg-4"><%=job_txt_220 %>:</div>               
                                <div class="col-lg-2"><%=job_txt_023 %>
                                    <%select case lto
                                       case "epi2017"
                                        %>
                                        / PO.
                                        <%
                                       end select%>:</div>
                            </div>
                            
                            <div class="row">
                                    <div class="col-lg-4">
                                  <!--  <input type="CHECKBOX" name="status_filter" value="1" <%=chk1%>/> <%=job_txt_221 %><br />
                                    <input type="CHECKBOX" name="status_filter" value="2" <%=chk2%>/> <%=job_txt_222 %><br />
                                    <input type="CHECKBOX" name="status_filter" value="3" <%=chk3%>/> <%=job_txt_063 %><br />
                                    <input type="CHECKBOX" name="status_filter" value="4" <%=chk4%>/> <%=job_txt_224 %><br />
                                    <input type="CHECKBOX" name="status_filter" value="0" <%=chk0%>/> <%=job_txt_225 %><br />
                                    <input type="CHECKBOX" name="status_filter" value="5" <%=chk5%>/> Evaluering<br /> -->

                                 <div class="multiselect">
                                    <div id="multiselect" class="selectBox" onclick="showCheckboxes()" style="width:440px;">
                                        <select class="form-control input-small">
                                            <option id="statustxt"><%=strstatus %></option>
                                        </select>
                                        <div class="overSelect"></div>
                                    </div>

                                    <div id="checkboxes" style="background-color:white; display:<%=statusdiv%>; width:440px;">
                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="1" name="status_filter" value="1" <%=chk1%>/> <span id="status_navn_1"><%=job_txt_221 %></span></label>

                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="2" name="status_filter" value="2" <%=chk2%>/> <span id="status_navn_2"><%=job_txt_222 %></span></label>

                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="3" name="status_filter" value="3" <%=chk3%>/> <span id="status_navn_3"><%=job_txt_063 %></span></label>

                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="4" name="status_filter" value="4" <%=chk4%>/> <span id="status_navn_4"><%=job_txt_224 %></span></label>

                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="0" name="status_filter" value="0" <%=chk0%>/> <span id="status_navn_0"><%=job_txt_225 %></span></label>

                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="5" name="status_filter" value="5" <%=chk5%>/> <span id="status_navn_5"><%=job_txt_629 %></span></label>

                                    </div>
                                    </div>


                                </div>
                                <div class="col-lg-4">

                                    <%select case lto 
                                    case "epi2017"
                                        selBGCol = "" '"lightpink"
                                    case else
                                        selBGCol = ""
                                    end select
                                        %>

                                    <select name="FM_medarb_jobans" id="FM_medarb_jobans" class="form-control input-small" onchange="submit();" style="background-color:<%=selBGCol%>">
                                    <option value="0"><%=job_txt_002 & " " %>(<%=job_txt_205 %>)</option>
                                    <%
                                    mNavn = "Alle"
            
                                    select case lto
                                    case "hestia"
                                        passivSQL = " OR mansat = 3"
                                    case else
                                        passivSQL = ""
                                    end select

                                     strSQL = "SELECT mnavn, mnr, mid, init, mansat FROM medarbejdere WHERE mansat = 1 "& passivSQL &" ORDER BY mnavn"
                                     oRec.open strSQL, oConn, 3
                                     while not oRec.EOF
                
                                        if cdbl(medarb_jobans) = oRec("mid") then
                                        selThis = "SELECTED"
                                        else
                                        selThis = ""
                                        end if

                                        if cint(oRec("mansat")) = 3 then
                                        passvHTML = " - Passiv"
                                        else
                                        passvHTML = ""
                                        end if
                
                
                                        %>
                                     <option value="<%=oRec("mid")%>" <%=selThis%>><%=oRec("mnavn") %> 
                                     <%if len(trim(oRec("init"))) <> 0 then %>
                                      [<%=oRec("init") %>]
                                      <%end if %>

                                      <%=passvHTML %>

                                      </option>
                                     <%
                                     oRec.movenext
                                     wend
                                     oRec.close
                                     %>
             
              
                                    </select>
                                </div>

                            
                            </div>

                            
                            <div class="row">
                                <div class="col-lg-10 pad-t20"><a href="#"><span id="ud_s_span"><%=joblog_txt_007 %> +</span></a></div>
                                <!--<div class="col-lg-2" valign="bottom"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=godkend_txt_005 %> >></b></button></div>-->
                            </div>
                            
                            <div id="udvidet_sog" style="display:none; visibility:hidden;">
                           
                                <div class="row">
                                    <div class="col-lg-4"><%=job_txt_136 %>:</div>
                                    <div class="col-lg-4"><%=job_txt_144 %>: </div>
                                </div>
                                <div class="row">
                                    <div class="col-lg-4">
                                        <select name="FM_prjgrp" class="form-control input-small" onchange="submit();">
                                            <option value="-1"><%=job_txt_002 &" " %>(<%=job_txt_205 %>)</option>
                                            <%strSQL = "SELECT id, navn FROM projektgrupper "_
                                            &" WHERE id <> 1 ORDER BY navn"
    
                                            'Response.Write strSQL & "<br>"
                                            'Response.flush
    
                                            oRec3.Open strSQL, oConn, 0, 3
                                            while Not oRec3.EOF 
    
                                            if cdbl(prjgrp) = cdbl(oRec3("id")) then
                                            pgSEL = "SELECTED"
                                            else
                                            pgSEL = ""
                                            end if

                                            %>
                                            <option value="<%=oRec3("id") %>" <%=pgSEL %>><%=oRec3("navn") %></option>
                                            <%
            
    
   
                                            oRec3.movenext
                                            wend
                                            oRec3.close %>
                                        </select>

                                        <br />
                                         <input type="hidden" value="<%=hd_kpers %>" name="FM_hd_kpers" id="FM_hd_kpers" />
		                  
                                                  <%=job_txt_203 %>:<br />
		                                <select name="FM_kundekpers" id="FM_kundekpers" class="form-control input-small">
		                                <option value="0"><%=job_txt_002 & " " %>(<%=job_txt_204 %>)</option>
		                                </select>

                                </div>
                                
                                    
                                 <!-- Forretningsomr�der -->    
                                <div class="col-lg-4">
                                
                                <%strSQLf = "SELECT id, navn FROM fomr WHERE id <> 0 ORDER BY navn"%>
                                
                                <select name="FM_fomr" id="Select2" multiple="multiple" size="6" class="form-control input-small">
                                <%if instr(strFomr_rel, "#0#") <> 0 then 
                                f0sel = "SELECTED"
                                else
                                f0sel = ""
                                end if%>

                                <option value="#0#" <%=f0sel %>><%=job_txt_002 &" " %>(<%=job_txt_205 %>)</option>
                                    
                                    <%oRec.open strSQLf, oConn, 3
                                    while not oRec.EOF
                                    
                                    if instr(strFomr_rel, "#"&oRec("id")&"#") <> 0 then
                                    fSel = "SELECTED"
                                    else
                                    fSel = ""
                                    end if
                                    
                                    %>
                                    <option value="#<%=oRec("id")%>#" <%=fSel %>><%=oRec("navn") %></option>
                                    <%
                                    oRec.movenext
                                    wend
                                    oRec.close
                                    %>
                                </select>


                                    </div>
                                </div>


                                 

                           

                           <div class="row">
                                <div class="col-lg-4">

                                <!-- Service aftale -->
		                        <%=job_txt_226 %>?<br />
		                        <input type="radio" name="aftfilt" value="visalle" <%=chk8%>>&nbsp;<%=job_txt_323 %>&nbsp;&nbsp;
		                        <input type="radio" name="aftfilt" value="serviceaft" <%=chk6%>>&nbsp;<%=job_txt_018 %>&nbsp;&nbsp;
		                        <input type="radio" name="aftfilt" value="ikkeserviceaft" <%=chk7%>>&nbsp;<%=job_txt_019 %>&nbsp;&nbsp;
		
                                    
                                <br /><br />

                                <%
                                call showEasyreg_fn()    
                                if cint(showEasyreg_val) = 1 then %>
                                <input name="FM_sogakt_easyreg" type="checkbox" value="1" <%=sogakt_easyregCHK%>  /> Vis kun job hvor der er Easyreg aktiviteter
                                <%end if %>


                               </div>

                               </div>
                            <div class="row">
                                <!-- DATOER -->
                               

                                <div class="col-lg-4">
                                 <br />
                                 <input type="checkbox" name="usedatokri" value="j" <%=datoKriJa%>> Show project with 
                                    <select name="usedatokri_st_or_end">
                                        <option value="1" <%=usedatokri_st_or_end_1SEL %>>Start Date</option>
                                        <option value="0" <%=usedatokri_st_or_end_0SEL %>>End Date</option>
                                    </select> in selected period:
                                    <!--job_txt_227-->
                                </div>
                               <div class="col-lg-2">
                                <%=godkend_txt_003 %>:
                                <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="aar" value="<%=aar %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                                <br />
                           </div>
                                 <div class="col-lg-2">
                               
                                <%=godkend_txt_004 %>:<br />
                                <div class='input-group date' id='datepicker_sldato'>
                                <input type="text" class="form-control input-small" name="aarslut" value="<%=aarslut %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                           </div>
                             </div> <!-- ROW -->    


                             </div><!-- SLUT s�g projekter udviddet -->

                             <div class="row">
                                <div class="col-lg-10 pad-t20">&nbsp;</div>
                                <div class="col-lg-2" valign="bottom"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=godkend_txt_005 %> >></b></button></div>
                            </div>

                        </form>
                    </div>
                    <%else %>
                    <form action="job_list.asp?" method="post">  
                        <input  type="hidden" name="FM_kunde" value="0" />
                        <div class="row">
                            <div class="col-lg-12">
                                <div class="multiselect" style="width:100%;">
                                    <div id="multiselect" class="selectBox" onclick="showCheckboxes()" >
                                        <select class="form-control">
                                            <option id="statustxt"><%=strstatus %></option>
                                        </select>
                                        <div class="overSelect"></div>
                                    </div>

                                    <div id="checkboxes" style="background-color:white; display:<%=statusdiv%>;">
                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="1" name="status_filter" value="1" <%=chk1%>/> <span id="status_navn_1"><%=job_txt_221 %></span></label>

                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="2" name="status_filter" value="2" <%=chk2%>/> <span id="status_navn_2"><%=job_txt_222 %></span></label>

                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="3" name="status_filter" value="3" <%=chk3%>/> <span id="status_navn_3"><%=job_txt_063 %></span></label>

                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="4" name="status_filter" value="4" <%=chk4%>/> <span id="status_navn_4"><%=job_txt_224 %></span></label>

                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="0" name="status_filter" value="0" <%=chk0%>/> <span id="status_navn_0"><%=job_txt_225 %></span></label>

                                        <label style="padding-left:10px;"><input type="CHECKBOX" class="statusfilt" id="5" name="status_filter" value="5" <%=chk5%>/> <span id="status_navn_5"><%=job_txt_629 %></span></label>

                                    </div>
                                </div>
                            </div>
                        </div>

                        <div class="row">
                            <div class="col-lg-12">
                                <input name="jobnr_sog" type="text" class="form-control" value="<%=show_jobnr_sog %>" placeholder="S�g" style="width:60%; display:inline-block;" />
                                <button type="submit" class="btn btn-secondary btn" style="margin-bottom:5px; width:19%;"><b><%=godkend_txt_005 %></b></button>
                                
                                <a href="jobs.asp?func=opret&id=0" class="btn btn-warning btn pull-right" style="margin-bottom:5px; text-align:center; width:19%;"><b>Opret</b></a>
                            </div>
                        </div>
                    </form>
                    <%end if %>

                </div>


                <!-- Change Job Status pop up window -->
                <div id="actionMenu_jobstatus" class="modal">
                    <div class="modal-content" style="text-align:left; width:300px; height:175px; text-align:center">

                        <form action="job_list.asp?func=jobstatusUpdate&jobnr_sog=<%=jobnr_sog %>&FM_kunde=<%=FM_kunde %>&FM_medarb_jobans=<%=medarb_jobans %>" method="post">
                            <input type="hidden" id="lastjobid" value="<%=lastjobid %>" />


                            <h4>�ndre Status p� de valtge job til</h4>
                            <br />

                            <input type="hidden" id="FM_selectedJobids_status" name="FM_selectedJobids_status" value="" />

                            <select name="FM_jobstatus" class="form-control input-small">
                                <option disabled><%=job_txt_613 %>..</option>
                                <option value="1"><%=job_txt_221 %></option>
                                <option value="2"><%=job_txt_222 %></option>
                                <option value="3"><%=job_txt_063 %></option>
                                <option value="4"><%=job_txt_224 %></option>

                                <%if (lto = "mpt" AND level = 1) OR lto <> "mpt" then %>
                                <option value="0"><%=job_txt_225 %></option>
                                <%end if %>

                                <option value="5"><%=job_txt_629 %></option>
                            </select>

                       
                            <br />

                            <button type="submit" class="btn btn-sm btn-success"><b><%=job_txt_449 %></b></button>

                        </form>
                    </div>

                    
                </div>


                <!-- Change Easyreg pop up window -->
                <div id="actionMenu_easyreg" class="modal">
                    <div class="modal-content" style="text-align:left; width:300px; height:175px; text-align:center">

                        <form action="job_list.asp?func=easyregUpdate&jobnr_sog=<%=jobnr_sog %>&FM_kunde=<%=FM_kunde %>&FM_medarb_jobans=<%=medarb_jobans %>" method="post">


                              <h4><%=left(tsa_txt_358, 8) %>:</h4>
                             <input type="hidden" id="FM_selectedJobids_easyreg" name="FM_selectedJobids_easyreg" value="" />
                            Tilf�j/Fjern alle aktiviteter p� valgte job til easyreg<br />
                            <select name="FM_tilfojeasyreg" class="form-control input-small">
                                <option value="0" selected>G�r ingenting</option>
                                <option value="1">Tilf�j</option>
                                <option value="2">Fjern</option>
                                </select>

                            <br />

                            <button type="submit" class="btn btn-sm btn-success"><b><%=job_txt_449 %></b></button>

                        </form>
                    </div>

                    
                </div>


                <!-- DELTETE pop up window -->
                <div id="actionMenu_delete" class="modal">
                    <div class="modal-content" style="text-align:left; width:300px; height:175px; text-align:center">

                        <%
                            slturl = "job_list.asp?func=sletok&id="&id&"&kt=0&fm_kunde_sog="&request("fm_kunde_sog")&"&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")&"&rdir="&rdir
                            %>

                             <form action="<%=slturl%>" method="post">
                            <input type="hidden" id="FM_selectedJobids_delete" name="FM_selectedJobids_delete" value="" />

                            <%slttxt = "<span style=""color:red;""><b>"&job_txt_004&"</b></span><br />"_
	                        & job_txt_005 & " <br>"&job_txt_006&" "&job_txt_007&"."
	                        
	
	                        %>
                            <%=slttxt %>
                            
                            <br />
                            <br />
                           
                                    <button type="submit" class="btn btn-sm btn-danger"><b><%=job_txt_612 %></b></button>

                          </form>
                    </div>

                    
                </div>
                <!-- Change Job end state pop up window -->
                <div id="actionMenu_enddate" class="modal">
                    <div class="modal-content" style="text-align:left; width:300px; height:175px; text-align:center">
                        
                        <form action="job_list.asp?func=jobslutdatoUpdate&jobnr_sog=<%=jobnr_sog %>&FM_kunde=<%=FM_kunde %>&FM_medarb_jobans=<%=medarb_jobans %>" method="post">
                          
                            <h4>�ndre slutdatoen p� de valgte job til</h4>
                            <br />

                            <input type="hidden" id="FM_selectedJobids_enddate" name="FM_selectedJobids_enddate" value="" />

                            <div class='input-group date'>
                                <input type="text" style="width:100%;" class="form-control input-small" name="FM_newEnddate" id="jq_dato" value="" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small"><span class="fa fa-calendar"></span></span>
                            </div>

                            <br />

                            
                            <button type="submit" class="btn btn-sm btn-success"><b><%=job_txt_449 %></b></button>

                        </form>
                    </div>   
                </div>

                <!-- Change Job / Activities' business areas -->
                <div id="actionMenu_forOm" class="modal">
                    <div class="modal-content" style="text-align:left; width:500px; height:275px;">

                        <form action="job_list.asp?func=foromUpdate&jobnr_sog=<%=jobnr_sog %>&FM_kunde=<%=FM_kunde %>&FM_medarb_jobans=<%=medarb_jobans %>" method="post">
                            <input type="hidden" id="sogtekst" value="<%=kunder_txt_118 %>" />  

                            <h4><%=job_txt_283 %></h4>
                            <br />

                            <input type="hidden" id="FM_selectedJobids_forOm" name="FM_selectedJobids_forOm" value="" />

                            <div style="text-align:center;">
                                <%         
                                strSQLf = "SELECT f.id, f.navn, kp.navn AS kontonavn, kp.kontonr FROM fomr AS f LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) WHERE f.id <> 0 AND f.jobok = 1 ORDER BY f.navn"
                                
                                'response.write strSQLf
                                'response.flush         
                                %>
                                <select name="FM_fomr" id="Select3" multiple="multiple" size="5" class="form-control input-small">
                                <option value="#0#"><%=job_txt_129 &" " %>(<%=job_txt_282 %>)</option>
                               
                                    
                                    <%oRec3.open strSQLf, oConn, 3
                                    while not oRec3.EOF
                                    
                                    if oRec3("kontonr") <> 0 then
                                        kontnrTxt = " ("& oRec3("kontonavn") &" "& oRec3("kontonr") &")"
                                    else 
                                        kontnrTxt = ""
                                    end if
                                    
                                    %>
                                    <option value="#<%=oRec3("id") %>#"><%=oRec3("navn") & kontnrTxt %></option>
                                    <%
                                    oRec3.movenext
                                    wend
                                    oRec3.close
                                    %>
                                </select>
                            </div>

                            <span style="text-align:left;">
                               <!-- <input id="Checkbox3" type="checkbox" name="FM_formr_opdater" value="1" /> �ndre p� de valgte jobs <br /> -->
                                <input id="Checkbox3" type="checkbox" name="FM_formr_opdater_akt" value="1" /> <%=job_txt_284 %>
                            </span>

                            <br /><br /><br />

                            <div style="text-align:center">
                                <button type="submit" class="btn btn-sm btn-success"><b><%=job_txt_449 %></b></button>
                            </div>

                        </form>
                    </div>   
                </div>

                <%if browstype_client <> "ip" then %>
                <div class="row">
                    <div class="col-lg-2">
                        <button type="submit" id="skiftvisning" class="btn btn-default btn-sm"><b><%=btntxt %></b></button>
                    </div>
                </div>
                <%end if %>

                <input type="hidden" name="FM_simplelist" id="FM_simplelist" value="<%=vis_simplejobliste %>" />

               
                <div class="row">
                    <div class="col-lg-12">
                <table class="table dataTable table-striped table-bordered table-hover ui-datatable" id="xjobliste">
                   
                        
                     <thead>
                        <tr>
                            <th class="list_advanced" style="display:<%=advanceddisplay%>"><%=job_txt_022 %></th>
                            <th class="list_advanced" style="display:<%=advanceddisplay%>"><%=job_txt_232 %> <span style="font-size:10px; font-weight:lighter">(<%=job_txt_221 %>)</span></th>
                            <th class="list_advanced" style="display:<%=advanceddisplay%>"><%=job_txt_233 %></th>
                         
                            <th class="list_advanced" style="display:<%=advanceddisplay%>"><%=job_txt_236 %></th>
                            <th class="list_advanced" style="display:<%=advanceddisplay%>"><%=job_txt_238 %><br /><span style="font-size:10px;"><%=job_txt_239 %></span></th>                            
                   
                            <th class="list_advanced" style="display:<%=advanceddisplay%>"><%=job_txt_248 %></th>
                            <th class="list_advanced" style="display:<%=advanceddisplay%>"><%=job_txt_132 %></th>
                            <th class="list_advanced" style="display:<%=advanceddisplay%>"><%=job_txt_247 %></th>
                            <th class="list_advanced" style="display:<%=advanceddisplay%>"><%=job_txt_241 %></th>
                            <th class="list_advanced" style="display:<%=advanceddisplay%>"> 
                                <select class="form-control input-small" id="actionMenu" style="width:40px; padding:0px; overflow:hidden;">
                                    <option value="0">Action menu..</option>
                                    <%if level = 1 then %>
                                    <option value="1"><%=job_txt_144 %></option>
                                    <%end if %>
                                    <option value="2"><%=job_txt_114 %></option>

                                    <option value="3">Status</option>

                                    <% call showEasyreg_fn()     
                                    if cint(showEasyreg_val) = 1 then %>
                                    <option value="4">Easyreg</option>
                                    <%end if %>

                                    <%if level = 1 then %>
                                    <option value="5"><%=job_txt_612 %></option>
                                    <%end if %>
                                </select>

                                <input type="checkbox" id="selectall" />

                            </th>


                            <!-- Simple jobliste -->
                            
                                <%if cint(simple_joblist_jobnavn) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_055 %></th>
                                <%end if %>

                                <%if cint(simple_joblist_jobnr) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_057 %></th>
                                <%end if %>

                                <%if cint(simple_joblist_kunde) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_021 %></th>
                                <%end if %>

                                <%if cint(simple_joblist_ans) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_230 %></th>
                                <%end if %>

                                <%if cint(simple_joblist_salgsans) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_626 %></th>
                                <%end if %>

                                <%if cint(simple_joblist_fomr) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_144 %></th>
                                <%end if %>

                                <%if cint(simple_joblist_prg) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_132 %></th>
                                <%end if %>

                                <%if cint(simple_joblist_stdato) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_100 %></th>
                                <%end if %>

                                <%if cint(simple_joblist_sldato) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_114 %></th>
                                <%end if %>

                                <%if cint(simple_joblist_status) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_241 %></th>
                                <%end if %>

                                <%if cint(simple_joblist_tidsforbrug) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_627 %></th>
                                <%end if %>

                                <%if cint(simple_joblist_budgettid) = 1 then %>
                                <th class="list_simple" style="display:<%=simpledisplay%>;"><%=job_txt_628 %></th>
                                <%end if %>

                          

                        </tr>


                    </thead>
                    <tbody>
                        

                        <%
                            'strJob = "SELECT jobnavn, jobnr, jobknr, fastpris, id, budgettimer, ikkebudgettimer, jo_valuta, jo_bruttooms, jobstatus, jobstartdato, jobslutdato, "_
                            '&" projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, "_
                            '&" jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, virksomheds_proc, kundekpers, rekvnr, tilbudsnr, sandsynlighed, preconditions_met "_ 
                            '&" FROM job AS j LEFT JOIN kunder as k ON (k.kid = j.jobknr) WHERE" & FM_kundeSqlFilter & varStatusFilt & jobansKri & prjgrpSQL & sogeKri

                            '***** MAIN SQL ****'
                            strJob = "SELECT jobnavn, jobnr, jobknr, fastpris, j.id, j.budgettimer, ikkebudgettimer, jo_valuta, jo_bruttooms, jobstatus, jobstartdato, jobslutdato, "_
                            &" j.projektgruppe1, j.projektgruppe2, j.projektgruppe3, j.projektgruppe4, j.projektgruppe5, j.projektgruppe6, j.projektgruppe7, j.projektgruppe8, j.projektgruppe9, j.projektgruppe10, "_
                            &" jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, salgsans1, virksomheds_proc, kundekpers, rekvnr, tilbudsnr, sandsynlighed, preconditions_met, fakturerbart, k.kid, serviceaft, k.useasfak, s.navn AS aftnavn, j.jo_valuta_kurs "_ 
                            &" FROM job AS j LEFT JOIN kunder as k ON (k.kid = j.jobknr) LEFT JOIN serviceaft as s ON (s.id = j.serviceaft)"
                            
                            if cint(sogakt) = 1 OR cint(sogakt_easyreg) = 1 then
                            strJob = strJob + " LEFT JOIN aktiviteter as a ON (a.job = j.id)"
                            end if

                             '** kontaktpersoner hos kunde **'
                            if cdbl(hd_kpers) <> -1 AND cdbl(hd_kpers) <> 0 then
                            kpersSQLkri = " AND kundekpers = " & hd_kpers
                            else
                            kpersSQLkri = ""
                            end if

                            strJob = strJob + " WHERE" & FM_kundeSqlFilter & varStatusFilt & varFilt & jobansKri & datoKri & prjgrpSQL & kpersSQLkri & sogeKri
                            strJob = strJob + " GROUP BY j.id Order BY k.kkundenavn, jobnavn, jobnr"

                            'if session("mid") = 1 then
                            'response.Write strJob
                            'Response.flush
                            'end if

                            cnt = 0
                            jids = 0
                            jo_bruttooms_tot = 0
                            oRec.open strJob, oConn, 3
                            while not oRec.EOF
                            


                            if level <= 2 OR level = 6 then
	                        editok = 1
	                        else
	                            if cint(session("mid")) = oRec("jobans1") OR cint(session("mid")) = oRec("jobans2") OR (cint(oRec("jobans1")) = 0 AND cint(oRec("jobans2")) = 0) then
	                            editok = 1
	                            end if
	                        end if


                            select case lto '** �ndret 20190409 -- Alle skal kunne komme ind. De bliver tjekket inde p� jobbet og kan ikke �ndre status, ti�f�je timer mm. hvis status IKKE er aktivt.
                            case "xmpt"

                                   'if oRec("jobstatus") <> 1 AND oRec("jobstatus") <> 4 AND oRec("jobstatus") <> 3 then

                                   '     if level <> 1 then
                                   '         editok = 0
                                   '     end if

                                   'end if

                            case else
                            end select


                            call forretningsomrJobId(oRec("id"))
                            if cint(visJobFomr) = 1 OR instr(strFomr_rel, "#0#") <> 0 then



                            if len(trim(oRec("budgettimer"))) = "0" OR oRec("budgettimer") = "0" then 
	                        budgettimer = 0
	                        else
	                        budgettimer = oRec("budgettimer")
	                        end if
	
	                        if oRec("ikkebudgettimer") > 0 then 
	                        ikkebudgettimer = oRec("ikkebudgettimer")
	                        else
	                        ikkebudgettimer = 0
	                        end if 
                            

                            call akttyper2009(2)

                            strSQLproaf = "SELECT sum(timer) AS proafslut FROM timer WHERE Tjobnr = '" & oRec("jobnr") &"' AND ("& aty_sql_realhours &") "
	                        oRec2.open strSQLproaf, oConn, 3
	                        proaf = 0
                            if not oRec2.EOF then
	                            if len(oRec2("proafslut")) <> 0 then
	                            proaf = oRec2("proafslut")
	                            else
	                            proaf = 0
	                            end if
                            end if
	                        oRec2.close

                            realfakbare = 0
                            '*** Real. fakturerbare timer **************************
	                        strSQL = "SELECT sum(timer) AS realfakbare FROM timer WHERE Tjobnr = '" & oRec("jobnr") &"' AND tfaktim = 1"
	                        oRec3.open strSQL, oConn, 3
	                        if not oRec3.EOF then

	                        if len(oRec3("realfakbare")) <> 0 then
	                        realfakbare = oRec3("realfakbare")
	                        else
	                        realfakbare = 0
	                        end if

	                        end if
	                        oRec3.close

                            restt = (budgettimertot - proaf)

                            budgettimertot = (ikkebudgettimer + budgettimer)
                            
                            if budgettimer <> 0 then
	                        projektcomplt = ((proaf/budgettimertot)*100)
                            projektcompltFakbare = ((realfakbare/budgettimertot)*100)
	                        else
	                        projektcompltFakbare = 100
                            projektcomplt = 100
	                        end if
	
	                        'if projektcomplt > 100 then
                            if restt < 0 then 
	                        showprojektcomplt = projektcomplt
	                        projektcomplt = 100
	                        bgdiv = "Crimson"
                            bgdivFontcol = "#FFFFFF"
	                        else
	                        showprojektcomplt = projektcomplt
	                        projektcomplt = projektcomplt
	                        bgdiv = "yellowgreen"
                            bgdivFontcol = "#000000"
	                        end if

                            if projektcompltFakbare > 100 then
	                        showprojektcompltFakbare = projektcompltFakbare
	                        projektcompltFakbare = 100
                            else
	                        showprojektcompltFakbare = projektcompltFakbare
	                        projektcompltFakbare = projektcompltFakbare
                            end if
                            
                            preconditions_met = oRec("preconditions_met")    
                            





                            '*** Antal aktiviteter p� job *** KUN AKTIVE I DETTE VIEW 
                            x = 0
	                        strAktnavn = ""
                            lastFase = ""
                            strSQL2 = "SELECT id, navn, fase, aktstatus, easyreg FROM aktiviteter WHERE job = "& oRec("id") & " AND aktstatus = 1 ORDER BY fase, sortorder, navn"
	                        oRec5.open strSQL2, oConn, 3
	                        while not oRec5.EOF 
	                        x = x + 1

                            if lastFase <> oRec5("fase") AND isNull(oRec5("fase")) <> true AND len(trim(oRec5("fase"))) <> 0 then

                            strAktnavn = strAktnavn & "<br><b>"& replace(oRec5("fase"), "_", " ") & "</b><br>"
                            lastFase = oRec5("fase")
                            end if

                            strAktnavn = strAktnavn & left(oRec5("navn"), 20) 
        
                                if oRec5("aktstatus") = 0 then
                                strAktnavn = strAktnavn & " - lukket"
                                end if

                                if oRec5("aktstatus") = 2 then
                                strAktnavn = strAktnavn & " - passiv"
                                end if
        
                                if oRec5("easyreg") = 1 then
                                strAktnavn = strAktnavn & " <span style=""color:red;"">(E!)</span>"
                                end if

                            strAktnavn = strAktnavn & "<br>"


	                        oRec5.movenext
	                        wend
	                        oRec5.close
	                        Antal = x


                        if cdbl(lastjobid*1) = cdbl(oRec("id")*1) then
                            trbgcollastjobid = "#FFFF99"
                        else
                            trbgcollastjobid = ""
                        end if


                        %>
                        <tr id="tr_jobid_<%=oRec("id") %>">
                            <td class="list_advanced" style="display:<%=advanceddisplay%>; width:250px; background-color:<%=trbgcollastjobid%>;">
                                

                                
                               

                                <%
                                    strKunde = "SELECT kkundenavn, kkundenr, kid FROM kunder WHERE kid ="& oRec("jobknr")
                                    oRec2.open strKunde, oConn, 3
                                    if not oRec2.EOF then
                                    strKundenavn = oRec2("kkundenavn")
                                    strKundenr = oRec2("kkundenr")
                                    end if
                                    oRec2.close                                      
                                %>
                                <%response.Write strKundenavn &" ("& strKundenr &") <br>" %>
                                <%if editok = 1 then %>         
                                <%if lto <> "care" then %>
                                <a href="../timereg/jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=1&jobnr_sog=&filt=1&fm_kunde_sog=0"><%=oRec("jobnavn") &" ("& oRec("jobnr") & ")" %></a>
                                <%else %>
                                <a href="jobs.asp?menu=job&func=red&jobid=<%=oRec("id")%>&int=1&jobnr_sog=&filt=1&fm_kunde_sog=0"><%=oRec("jobnavn") &" ("& oRec("jobnr") & ")" %></a>
                                <%end if %>

                                    <%
                                      select case lto
                                      case "mpt", "intranet - local", "outz"
                                      
                                      'if session("mid") = 1 then%>
                                        <br />NY: <a href="jobs.asp?func=red&jobid=<%=oRec("id")%>" style="color:orange;"><%=oRec("jobnavn") &" ("& oRec("jobnr") & ")" %></a>
                                      <%
                                      'end if

                                      end select
                                      %>

                                <%else %>
                                <%=oRec("jobnavn") &" ("& oRec("jobnr") & ")" %>
                                <%end if %>

                                


                                <%
                                if oRec("fastpris") = 1 then
                                strFasptpris = job_txt_324

                                    %>
                                     <span style="color:green; font-size:10px; line-height:11px;">(<%=strFasptpris %>)</span>
                                    <%

                                else
                                strFasptpris = job_txt_325 'vises ikke
                                end if
                                %>
                            
                               
                                <!-- Her skal mere info ind se linje 8345 i gamle jobs -->
                                 
                                <span style="color:#999999; font-size:10px; line-height:11px;"><br /><i>
                                <%
		
						        '*** Jobansvarlige ***
                                if oRec("jobans1") <> 0 then
						        call meStamData(oRec("jobans1"))
                                %>
						        <%=meNavn%> 
                                        <%if oRec("jobans_proc_1")  <> 0 then %>
                                    (<%=oRec("jobans_proc_1") %>%)
                                    <%end if %>

                                <%end if
                        

                                '*** Jobejer 2 ***
                                if oRec("jobans2") <> 0 then
						        call meStamData(oRec("jobans2"))
                                %>
						        , <%=meNavn%> 
                        
                                <%if oRec("jobans_proc_2")  <> 0 then %>
                                (<%=oRec("jobans_proc_2") %>%)
                                <%end if %>

                                <%end if

                                '*** Job medansvarlige ***
                                if oRec("jobans3") <> 0 then
						        call meStamData(oRec("jobans3"))
                                %>
						        , <%=meNavn%> 
                                    <%if oRec("jobans_proc_3")  <> 0 then %>
                                (<%=oRec("jobans_proc_3") %>%)

                                <%end if
                        
                                end if


                                '*** Job medansvarlige ***
                                if oRec("jobans4") <> 0 then
						        call meStamData(oRec("jobans4"))
                                %>
						        , <%=meNavn%> 
                                        <%if oRec("jobans_proc_4")  <> 0 then %>
                                    (<%=oRec("jobans_proc_4") %>%)
                                    <%end if %>

                                <%end if


                                '*** Job medansvarlige ***
                                if oRec("jobans5") <> 0 then
						        call meStamData(oRec("jobans5"))
                                %>
						        , <%=meNavn%> 
                                        <%if oRec("jobans_proc_5")  <> 0 then %>
                                    (<%=oRec("jobans_proc_5") %>%)
                                    <%end if %>

                                <%end if

                                '*** Virksomhedsproc fjernet 20181010 level = 100 ****'
                                if level = 100 then
                                        if (lto <> "epi" AND lto <> "epi2017") OR (lto = "epi" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59 OR thisMid = 1720)) OR (lto = "epi_cati" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59)) OR (lto = "epi_no" AND thisMid = 2) OR (lto = "epi_ab" AND thisMid = 2) OR (lto = "epi_sta" AND thisMid = 2) OR (lto = "epi_uk" AND thisMid = 2) then 
                                        if oRec("virksomheds_proc") <> 0 then%>
                                        <br /><b><%=lto %>: </b> (<%=oRec("virksomheds_proc") %>%)
                                        <%end if
                                    end if
                                end if
                      
						
						        '**********************
						        %></i></span>

                                <%'*** kontaktpersoner hos kunde '******* 
                                
                                 kpersNavn = ""
                                 if cint(oRec("kundekpers")) <> 0 then

                                 strSQLkpers = "SELECT navn FROM kontaktpers WHERE id = " & oRec("kundekpers")
                                 oRec6.open strSQLkpers, oConn, 3
                                 if not oRec6.EOF then

                                 kpersNavn = oRec6("navn")

                                 end if
                                 oRec6.close 
                                %>
                                <span style="color:#8caae6; font-size:10px; line-height:11px;"><br /><%=job_txt_253 %>: <%=kpersNavn %></span>
                                <%end if 
                                    

                                '**** DATO overs�telser ****
                                if month(oRec("jobstartdato")) = 1 then jobstartdatoMonthTxt = job_txt_588 end if
                                if month(oRec("jobstartdato")) = 2 then jobstartdatoMonthTxt = job_txt_589 end if
                                if month(oRec("jobstartdato")) = 3 then jobstartdatoMonthTxt = job_txt_590 end if
                                if month(oRec("jobstartdato")) = 4 then jobstartdatoMonthTxt = job_txt_591 end if
                                if month(oRec("jobstartdato")) = 5 then jobstartdatoMonthTxt = job_txt_592 end if
                                if month(oRec("jobstartdato")) = 6 then jobstartdatoMonthTxt = job_txt_593 end if
                                if month(oRec("jobstartdato")) = 7 then jobstartdatoMonthTxt = job_txt_594 end if
                                if month(oRec("jobstartdato")) = 8 then jobstartdatoMonthTxt = job_txt_595 end if
                                if month(oRec("jobstartdato")) = 9 then jobstartdatoMonthTxt = job_txt_596 end if
                                if month(oRec("jobstartdato")) = 10 then jobstartdatoMonthTxt = job_txt_597 end if
                                if month(oRec("jobstartdato")) = 11 then jobstartdatoMonthTxt = job_txt_598 end if
                                if month(oRec("jobstartdato")) = 12 then jobstartdatoMonthTxt = job_txt_599 end if

                                if month(oRec("jobslutdato")) = 1 then jobslutdatoMonthTxt = job_txt_588 end if
                                if month(oRec("jobslutdato")) = 2 then jobslutdatoMonthTxt = job_txt_589 end if
                                if month(oRec("jobslutdato")) = 3 then jobslutdatoMonthTxt = job_txt_590 end if
                                if month(oRec("jobslutdato")) = 4 then jobslutdatoMonthTxt = job_txt_591 end if
                                if month(oRec("jobslutdato")) = 5 then jobslutdatoMonthTxt = job_txt_592 end if
                                if month(oRec("jobslutdato")) = 6 then jobslutdatoMonthTxt = job_txt_593 end if
                                if month(oRec("jobslutdato")) = 7 then jobslutdatoMonthTxt = job_txt_594 end if
                                if month(oRec("jobslutdato")) = 8 then jobslutdatoMonthTxt = job_txt_595 end if
                                if month(oRec("jobslutdato")) = 9 then jobslutdatoMonthTxt = job_txt_596 end if
                                if month(oRec("jobslutdato")) = 10 then jobslutdatoMonthTxt = job_txt_597 end if
                                if month(oRec("jobslutdato")) = 11 then jobslutdatoMonthTxt = job_txt_598 end if
                                if month(oRec("jobslutdato")) = 12 then jobslutdatoMonthTxt = job_txt_599 end if

                                %>
                                <span style="font-size:10px; line-height:11px;"><br><%=day(oRec("jobstartdato"))&". "& jobstartdatoMonthTxt &" "& year(oRec("jobstartdato")) %> - <%=day(oRec("jobslutdato"))&". "& jobslutdatoMonthTxt &" "& year(oRec("jobslutdato")) %></span>
                                

                                <% 
	                            if len(trim(oRec("rekvnr"))) <> 0 then
                                %>
                                <span style="font-size:10px; line-height:11px; color:#999999;">
                                <%
	                            Response.Write "<br>"&job_txt_254&": "& oRec("rekvnr") 
                                %>
                                 </span>
                                <%
	                            end if
                                
                                if len(trim(oRec("aftnavn"))) <> 0 then
                                %><span style="font-size:10px; background-color:#FFFFe1; color:#000000;"><%
	                            Response.Write "<br>"&job_txt_255&": "& oRec("aftnavn") 
                                %>
                                 </span>
                                <%
	                            end if 
	                           
	                            if oRec("jobstatus") = 3 then
                                %><span style="font-size:10px; background-color:#ffdfdf; color:#000000;"><%
	                            Response.Write "<br>"&job_txt_256&": "& oRec("tilbudsnr") &" ("& oRec("sandsynlighed") &" %)"
                                %>
                                 </span>
                                <%
	                            end if %>

                                <%
                                select case preconditions_met 
                                case 0
                                preconditions_met_Txt = ""
                                    case 1
                                preconditions_met_Txt = "<br><span style='color:#6CAE1C; font-size:10px; background:#DCF5BD;'>"&job_txt_257&"</span>"
                                case 2
                                preconditions_met_Txt = "<br><span style='color:red; font-size:10px; background:pink;'>"&job_txt_258&"!</span>"
                                case else
                                preconditions_met_Txt = ""
                                end select
                                %>
                                <%=preconditions_met_Txt %>                              
                            </td>

                            <td class="list_advanced" style="display:<%=advanceddisplay%>; background-color:<%=trbgcollastjobid%>;">
                            
                                <%
                                '***** Aktiviteter *****'
                                %>
                                
                                <div style="font-size:10px; width:120px; overflow:auto; color:#999999; padding-right:5px; max-height:80px;">           
                                <%=strAktnavn %></div>

                                <span style="font-size:10px;">
                                    <%if editok = 1 then%>
                                    ialt <a href="#" onclick="Javascript:window.open('../timereg/aktiv.asp?menu=job&jobid=<%=oRec("id")%>&jobnavn=<%=oRec("jobnavn")%>&rdir=job3&nomenu=1', '', 'width=1354,height=800,resizable=yes,scrollbars=yes')" class=vmenu><%=Antal%></a>
		                            <%else%>
		                            <b>(<%=Antal%>)</b>
		                            <%end if%>
                                </span>
                            </td>



                            <!-- Forretningsomr�der --> 
                            <td class="list_advanced" style="display:<%=advanceddisplay%>; max-width:150px; font-size:10px; background-color:<%=trbgcollastjobid%>;">                             
                                <%call forretningsomrJobId(oRec("id")) %>

                                <%=strFomr_navn %> &nbsp
                            </td>


                         

                            <!--- Budget, Stated, Realiseret -->

                            <%call valutakode_fn(oRec("jo_valuta")) %>

		                    <td class="list_advanced" style="display:<%=advanceddisplay%>; text-align:right; background-color:<%=trbgcollastjobid%>;">
		                    <%=formatnumber(oRec("jo_bruttooms"), 2) &" "& valutaKode_CCC %>
                                
                            <%
                                call beregnValuta(oRec("jo_bruttooms"),oRec("jo_valuta_kurs"),basisValKurs)

                                jo_bruttooms_tot = jo_bruttooms_tot + valBelobBeregnet*1
                                
                                 %>

		                    </td>

                            <td class="list_advanced" align=right valign=top style="padding:4px 15px 4px 4px; font-size:10px; display:<%=advanceddisplay%>; background-color:<%=trbgcollastjobid%>;">
                                
                                 <%=job_txt_261 %>: <span style="width:40px; height:20px; background-color:<%=bgdiv%>; color:<%=bgdivFontcol%>;  padding:2px 0px 2px 2px;">
                                      <%=formatnumber(budgettimertot, 2)%>
                                </span>
                                <br />
                                <%=job_txt_239 %>:  
                                
                                <%if showprojektcomplt > 0 then%>
		                           <%=formatnumber(proaf)%> (<%=formatpercent(showprojektcomplt/100, 0)%>)
		                        <%end if%>
                                <br>
                            <span style="color:#999999;"><%=job_txt_262 %>: <%=formatnumber(realfakbare, 2) %></span> <br />
                            -----------------------<br />
                            <%=job_txt_263 %> = <%=formatnumber(restt)%><br />
                            <span style="color:#999999;"><%=formatnumber(resttfakbare) %></span><br />
                            =====================
		                    </td>                           

                          

                            <td class="list_advanced" style="padding-right:10px; font-size:10px; vertical-align:top; display:<%=advanceddisplay%>; background-color:<%=trbgcollastjobid%>;">
		                   	<div style="position:relative; top:0px; left:0px; width:120px; max-height:80px; overflow:auto;">
                     <table cellspacing=0 border=0 cellpadding=0 width=100%>
		
		<%
		
		        '** Findes der fakturaer p� job kan det ikke slettes **'
			    
			    deleteok = 0
			    faktotbel = 0
                fakturaBel_tot = 0
                fakAftaleid = 0
			    strSQLffak = "SELECT f.fid, f.faknr, f.aftaleid, f.faktype, f.jobid, f.fakdato, f.beloeb, "_
			    &" f.faktype, f.kurs, brugfakdatolabel, labeldato, fakadr, shadowcopy, f.valuta, afsender FROM fakturaer f "_
			    &" WHERE (jobid = " & oRec("id") & " AND shadowcopy = 0) OR (aftaleid = "& oRec("serviceaft") &" AND aftaleid <> 0 AND shadowcopy = 0)"_
			    &" GROUP BY f.fid ORDER BY f.fakdato DESC"

                'OR aftid = "& oRec("serviceaft") &"
                '&" LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid)"_
                'AND fd.enhedsang <> 3
                'SUM(fd.aktpris) AS aktbel
                '&" AND aftaleid = 0 AND shadowcopy = 0"_


			    'Response.Write strSQLffak 
			    'Response.flush
			   
                f = 0
			    oRec3.open strSQLffak, oConn, 3
			    while not oRec3.EOF 
			        

                    if oRec3("brugfakdatolabel") = 1 then '** Labeldato
                    fakDato = job_txt_623&": "& oRec3("labeldato") 'VIS IKKE Faktura sys dato p� jioblsiten: Gnidret og manlger overblik & "<br><span style=""color:#999999; size:10px,"">"& oRec3("fakdato") &"</span>"
                    else
                    fakDato = job_txt_622&": "& oRec3("fakdato")
                    end if
			       

                              strFakAftNavn = ""

                             if oRec3("aftaleid") <> 0 then 'AND oRec2("aftaleid") <> 0 then 'oRec3("shadowcopy") = 1
               
                   

                                        strSQLFakorg = "SELECT f.fid, f.beloeb, f.valuta, f.kurs, f.faktype, f.aftaleid, fd.aktpris FROM fakturaer f "_
                                        &" LEFT JOIN faktura_det fd ON (fd.fakid = f.fid AND fd.aktid = "& oRec("id") &") WHERE fid = "& oRec3("fid") &" AND shadowcopy <> 1 "
                                        
                                        oRec8.open strSQLFakorg, oConn, 3
                                        if not oRec8.EOF then

                        
                                            fakValuta = oRec8("valuta")
                                            fakKurs = oRec8("kurs")
                                            fakbeloeb = oRec8("aktpris") 'oRec8("beloeb")
                                            fakType = oRec8("faktype")
                                            fakAftaleid = oRec8("aftaleid")
                                            fakAktbel = oRec8("aktpris")

                       
                                        end if
                                        oRec8.close

                                         
                                        if fakAftaleid <> "" then
                                        fakAftaleid = fakAftaleid
                                        else
                                        fakAftaleid = 0
                                        end if
                                        
            
                                        strSQLaftale = "SELECT navn, aftalenr FROM serviceaft WHERE id = " & fakAftaleid
                                        oRec8.open strSQLaftale, oConn, 3
                                        if not oRec8.EOF then
                                        
                                        strFakAftNavn = "<span style=""font-size:10px; color:#999999;""> - " & left(oRec8("navn"), 5) & " ("& oRec8("aftalenr") &")</span>" 
                                        
                                        end if
                                        oRec8.close



                            else
                        
                 
                        
                                fakValuta = oRec3("valuta")
                                fakKurs = oRec3("kurs")
                                fakBeloeb = oRec3("beloeb")
                                fakType = oRec3("faktype")
                                fakAktbel = 0 'oRec3("aktbel")

                            end if 'aftale fak

			          
                      %>
                      <tr>
                      <%
			        
                            'fidLink = 0
                            'if editok = 1 then 'cdate(oRec3("fakdato")) >= cdate("01-01-2006") AND
                            fidLink = oRec3("fid")    
                            'end if

                          
                          if cint(editok) = 1 then 'fidLink <> 0 then
                          %>
			              <td class="list_advanced" valign=top style="font-size:10px; display:<%=advanceddisplay%>; background-color:<%=trbgcollastjobid%>;"><a href="../timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id=<%=fidLink%>&FM_jobonoff=<%=FM_jobonoffval%>&FM_kunde=<%=oRec3("fakadr")%>&FM_job=<%=oRec3("jobid")%>&FM_aftale=<%=oRec3("aftaleid")%>&fromfakhist=1" class="lgron" target="_blank"><b><%=oRec3("faknr")%></b></a> 
                          <%=strFakAftNavn %>
                          </td>
                          <%else 
                          %><td class="list_advanced" valign=top style="font-size:10px; display:<%=advanceddisplay%>; background-color:<%=trbgcollastjobid%>;"><b><%=oRec3("faknr") &" "& strFakAftNavn %></b></td>
                          <%end if %>
                          
                         <td class="list_advanced" align=right valign=top style="white-space:nowrap; font-size:10px; display:<%=advanceddisplay%>; background-color:<%=trbgcollastjobid%>;"><%=fakDato %></td>
                       
                      </tr>
                      <%
                    


                  
                     
	                call beregnValuta(minus&(fakBeloeb),fakKurs,100)
                    if fakType <> 1 then
                    belobGrundVal = valBelobBeregnet
                    else
                    belobGrundVal = -valBelobBeregnet
                    end if 
                    

                        if cDate(oRec3("fakdato")) < cDate("01-06-2010") AND (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk") then
                        belobKunTimerStk = belobKunTimerStk + oRec3("beloeb")
                        else
       
                    
	                        '** Kun aktiviteter timer, enh. stk. IKKE materialer og KM
	                        call beregnValuta(minus&(fakAktbel),fakKurs,100)
                            if fakType <> 1 then
                            belobKunTimerStk = valBelobBeregnet
                            else
                            belobKunTimerStk = -valBelobBeregnet
                            end if 


                     end if
            
                fakturaBel_tot = fakturaBel_tot + (belobGrundVal)
                faktotbel = faktotbel + belobKunTimerStk
			    f = f + 1
			    oRec3.movenext
			    wend
			    oRec3.close

                fakturaBel_tot_gt = fakturaBel_tot_gt + fakturaBel_tot
			    %>
                </table>
               

			    </div>
                 <%=job_txt_270 %>:<br /> 
                 <b><%=formatnumber(fakturaBel_tot, 2) & " "& basisValISO %></b>

                 <%if totReal <> 0 then 
	        gnstpris = faktotbel/totReal
	        else 
	        gnstpris = 0
	        end if%>
	        
            <%if gnstpris <> 0 then %>
            <br /><%=job_txt_271 %>:<br />
	        <b><%=formatnumber(gnstpris) & " "& basisValISO %></b>
            
            <!-- <span class="qmarkhelp" id="qm0001" style="font-size:11px; color:#999999; font-weight:bolder;">?</span><span id="qmarkhelptxt_qm0001" style="visibility:hidden; color:#999999; display:none; padding:3px; z-index:4000;">(faktureret bel�b - (materialeforbrug + km)) / timer realiseret</span>-->
	        
	        <%
            end if

	        if f <> 0 AND proaf <> 0 then
	        gnsPrisTot = gnsPrisTot + (faktotbel)
	        else
	        gnsPrisTot = gnsPrisTot
	        'totRealialt = totRealialt
	        end if

            totRealialt = totRealialt + (proaf)
            %>
            </td>


            <%'*** END invoices 
                
                
            '***** Projektgrupper ******'%>



                            <td class="list_advanced" style="white-space:nowrap; padding:4px 4px 0px 4px; font-size:10px; line-height:11px; vertical-align:top; display:<%=advanceddisplay%>; background-color:<%=trbgcollastjobid%>;">
                            <% for p = 1 to 10
        
                            pgid = oRec("projektgruppe"&p)

                            if pgid <> 1 then
                                call prgNavn(pgid, 200)
                                Response.Write left(prgNavnTxt, 20) & "<br>"
                            end if
        
                            next %>
		   
		                    </td>


                            <td class="list_advanced" style="font-size:75%; display:<%=advanceddisplay%>; background-color:<%=trbgcollastjobid%>;">

                                  <%if editok = 1 then %>
                                        <a href="../timereg/job_print.asp?id=<%=oRec("id") %>" target="_blank">Print / PDF</a><br />
		                                <a href="../timereg/job_kopier.asp?func=kopier&id=<%=oRec("id")%>&fm_kunde=<%=oRec("kid") %>" target="_blank"><%=job_txt_264 %></a><br />
                                        <a href="../timereg/materialer_indtast.asp?id=<%=oRec("id")%>&fromsdsk=0&aftid=<%=oRec("serviceaft")%>" target="_blank"><%=job_txt_265 %></a><br />

                                        <%
         
                                        viskunabnejob0 = "1"
                                        viskunabnejob1 = "" 
                                        viskunabnejob2 = ""
                                        select case oRec("jobstatus")
                                        case 0
                                        viskunabnejob0 = "1"
                                        case 1,2
                                        viskunabnejob2 = "1"
                                        case 3
                                        viskunabnejob1 = "1"
                                        end select

                                         select case lto
                                        case "synergi1", "xintranet - local"


                                        call licensStartDato()
                                        dtlink_stdag = startDatoDag
		                                dtlink_stmd = startDatoMd
		                                dtlink_staar = startDatoAar

                                        case else
        
                                        dtlink_stdag = datepart("d", oRec("jobstartdato"), 2, 2)
		                                dtlink_stmd = datepart("m", oRec("jobstartdato"), 2, 2)
		                                dtlink_staar = datepart("yyyy", oRec("jobstartdato"), 2, 2)

		
                                        end select

        

		                                dtlink_sldag = datepart("d", now, 2, 2)
		                                dtlink_slmd = datepart("m", now, 2, 2)
		                                dtlink_slaar = datepart("yyyy", now, 2, 2)
		
       
		
		                                dtlink = "FM_usedatokri=1&FM_start_dag="&dtlink_stdag&""_
	                                    &"&FM_start_mrd="&dtlink_stmd&"&FM_start_aar="&dtlink_staar&"&FM_slut_dag="&dtlink_sldag&""_
	                                    &"&FM_slut_mrd="&dtlink_slmd&"&FM_slut_aar="&dtlink_slaar
         
                                        %>
                                        
                                        <a href="../timereg/joblog.asp?nomenu=1&FM_job=<%=oRec("id") %>&FM_kunde=<%=oRec("kid")%>&FM_jobsog=<%=oRec("jobnr")%>&viskunabnejob0=<%=viskunabnejob0%>&viskunabnejob1=<%=viskunabnejob1%>&viskunabnejob2=<%=viskunabnejob2%>&<%=dtlink %>" target="_blank"><%=job_txt_266&" " %></a><br />
                                        <a href="../timereg/jobprintoverblik.asp?menu=job&id=<%=oRec("id")%>&media=printjoboverblik" class=rmenu target="_blank"><%=job_txt_267 %> </a><br />
                                        <a href="../timereg/timereg_akt_2006.asp?FM_kontakt=<%=oRec("kid")%>&FM_ignorer_projektgrupper=1&jobid=<%=oRec("id")%>&FM_sog_job_navn_nr=<%=oRec("jobnr")%>&usemrn=<%=session("mid")%>&showakt=1&fromsdsk=1" target="_blank" ><%=job_txt_268 &" " %> </a><br />
                                        <%if len(trim(oRec("useasfak"))) <> 0 then %>
                                            <%if cint(oRec("useasfak")) <= 2 then 'links skal fixes - feltet her skal hentets %>
                                            <a href="../timereg/erp_opr_faktura_fs.asp?func=opr&visfaktura=1&visjobogaftaler=1&visminihistorik=1&FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&<%=dtlink %>" target="_blank"><%=job_txt_269 %> </a><br />
		                                                                        
                                            <%end if %>
                                        <%end if %>

                                        <%else %>
                                        &nbsp
                                        <%end if %>

                                
                            </td>
                            
                            <td class="list_advanced" style="text-align:center; display:<%=advanceddisplay%>; background-color:<%=trbgcollastjobid%>;" ><!-- Status -->

                               
                                <%
                                    select case cdbl(oRec("jobstatus"))
                                    case 1
                                    statusCHBAR1 = "SELECTED"
                                    statusCHBAR2 = ""
                                    statusCHBAR3 = ""
                                    statusCHBAR4 = ""
                                    statusCHBAR0 = ""
                                    statusTXT = job_txt_221
                                    case 2
                                    statusCHBAR1 = ""
                                    statusCHBAR2 = "SELECTED"
                                    statusCHBAR3 = ""
                                    statusCHBAR4 = ""
                                    statusCHBAR0 = ""
                                    statusTXT = job_txt_222
                                    case 3
                                    statusCHBAR1 = ""
                                    statusCHBAR2 = ""
                                    statusCHBAR3 = "SELECTED"
                                    statusCHBAR4 = ""
                                    statusCHBAR0 = ""
                                    statusTXT = job_txt_063
                                    case 4
                                    statusCHBAR1 = ""
                                    statusCHBAR2 = ""
                                    statusCHBAR3 = ""
                                    statusCHBAR4 = "SELECTED"
                                    statusCHBAR0 = ""
                                    statusTXT = job_txt_224
                                    case 5
                                    statusCHBAR1 = ""
                                    statusCHBAR2 = ""
                                    statusCHBAR3 = ""
                                    statusCHBAR4 = "SELECTED"
                                    statusCHBAR0 = ""
                                    statusTXT = job_txt_629
                                    case 0
                                    statusCHBAR1 = ""
                                    statusCHBAR2 = ""
                                    statusCHBAR3 = ""
                                    statusCHBAR4 = ""
                                    statusCHBAR0 = "SELECTED"
                                    statusTXT = job_txt_225

                                    end select
                                %>

                              <!-- <select class="form-control input-small">
                                    <option value="1" <%=statusCHBAR1 %>>Aktiv</option>
                                    <option value="2" <%=statusCHBAR2 %>>Passiv/Fak</option>
                                    <option value="3" <%=statusCHBAR3 %>>Tilbud</option>
                                    <option value="4" <%=statusCHBAR4 %>>Gen.syn</option>
                                    <option value="0" <%=statusCHBAR0 %>>Lukket</option>
                                </select> -->
                                <%=statusTXT %>

                            </td>

                            <td class="list_advanced" style="text-align:center; display:<%=advanceddisplay%>; background-color:<%=trbgcollastjobid%>;">
                                <%if ((lto = "mpt" AND oRec("jobstatus") <> 0 AND oRec("jobstatus") <> 2) OR level = 1) OR lto <> "mpt" then %>
                                <input class="job_bulkCHB" name="job_bulkCHB" id="<%=oRec("id") %>" type="checkbox" />
                                <%end if %>
                            </td>



                            <!-- Simple Jobliste -->
                            
                                     <%if cint(simple_joblist_jobnavn) = 1 then %>
                                     <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>;">
                                         <%if editok = 1 then %>  
                                            <%if lto = "care" then %>
                                            <a href="jobs.asp?func=red&jobid=<%=oRec("id")%>"><%=oRec("jobnavn") &" ("& oRec("jobnr") & ")" %></a>
                                            <%else %>
                                                <%if browstype_client <> "ip" then %>
                                                    <a href="../timereg/jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=1&jobnr_sog=&filt=1&fm_kunde_sog=0"><%=oRec("jobnavn")%></a>
                                                <%else %>
                                                    <a style="color:orange;" href="jobs.asp?func=red&jobid=<%=oRec("id")%>"><%=oRec("jobnavn")%></a>
                                                <%end if %>
                                            <%end if %>
                                         <%else %>
                                         <%=oRec("jobnavn") %>
                                         <%end if %>
                                     </td>
                                     <%end if %>

                                     <%if cint(simple_joblist_jobnr) = 1 then %>
                                     <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>;">
                                         <%if editok = 1 then %> 
                                            <%if lto = "care" then %>
                                            <a href="jobs.asp?func=red&jobid=<%=oRec("id")%>"><%=oRec("jobnr")%></a>
                                            <%else %>
                                                <%if browstype_client <> "ip" then %>
                                                <a href="../timereg/jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=1&jobnr_sog=&filt=1&fm_kunde_sog=0"><%=oRec("jobnr")%></a>
                                                <%else %>
                                                <%=oRec("jobnr")%>
                                                <%end if %>
                                            <%end if %>
                                                  <%
                                                  select case lto
                                                  case "mpt", "intranet - local", "outz"
                                      
                                                  if browstype_client <> "ip" then%>
                                                    <br />NY: <a href="jobs.asp?func=red&jobid=<%=oRec("id")%>" style="color:orange;"><%=oRec("jobnavn") &" ("& oRec("jobnr") & ")" %></a>
                                                  <%
                                                  end if

                                                  end select
                                                  %>

                                         <%else %>
                                            <%=oRec("jobnr") %>
                                         <%end if %>
                                     </td>
                                     <%end if %>

                                     <%if cint(simple_joblist_kunde) = 1 then %>
                                     <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>;"><%=strkundenavn %></td>
                                     <%end if %>

                                     <%if cint(simple_joblist_ans) = 1 then %>
                                     <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>;">
                                         <%
                                         if oRec("jobans1") <> 0 then
                                             call meStamData(oRec("jobans1"))
                                             response.Write meNavn 
                                         end if
                                         %>
                                     </td>
                                     <%end if %>

                                     <%if cint(simple_joblist_salgsans) = 1 then %>
                                     <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>;">
                                         <%
                                         if oRec("salgsans1") <> 0 then
                                             call meStamData(oRec("salgsans1"))
                                             response.Write meNavn 
                                         end if
                                         %>
                                     </td>
                                     <%end if %>

                                     <%if cint(simple_joblist_fomr) = 1 then %>
                                    <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>;">
                                        <%call forretningsomrJobId(oRec("id")) %>
                                        <%
                                        if len(strFomr_navn) > 45 then
                                            strFomr_navn = left(strFomr_navn, 45) & "..."
                                        else
                                            strFomr_navn = strFomr_navn
                                        end if
                                        %>
                                        <%=strFomr_navn %>
                                    </td>
                                     <%end if %>

                                     <%if cint(simple_joblist_prg) = 1 then %>
                                     <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>;">
                                         <% 
                                            strprg = ""
                                            for p = 1 to 10        
                                            pgid = oRec("projektgruppe"&p)

                                            if pgid <> 1 then
                                                call prgNavn(pgid, 200)
                                                if p = 1 then
                                                    strprg = prgNavnTxt
                                                else
                                                    strprg = strprg & ", " &prgNavnTxt
                                                end if
                                            end if    
                                        next 
                                     
                                        if len(strprg) > 20 then
                                             strprg = left(strprg, 20) & "..."
                                        else
                                             strprg = strprg
                                        end if
                                        response.Write strprg
                                        %>
                                     </td>
                                     <%end if %>

                                     <%if cint(simple_joblist_stdato) = 1 then %>
                                     <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>;">

                                         <% 
                                             srtyear = year(oRec("jobstartdato"))
                                             strmonth = month(oRec("jobstartdato"))

                                             if strmonth < 10 then
                                                strmonth = "0" & strmonth
                                             end if

                                             strday = day(oRec("jobstartdato"))

                                             if strday < 10 then
                                                strday = "0" & strday
                                             end if

                                             strdato = srtyear&"-"&strmonth&"-"&strday
                                        %>

                                         <span style="display:none;"><%=strdato %></span>
                                         <%=oRec("jobstartdato") %>

                                     </td>
                                     <%end if %>

                                     <%if cint(simple_joblist_sldato) = 1 then %>
                                     <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>;">

                                         <% 
                                             srtyear = year(oRec("jobslutdato"))
                                             strmonth = month(oRec("jobslutdato"))

                                             if strmonth < 10 then
                                                strmonth = "0" & strmonth
                                             end if

                                             strday = day(oRec("jobslutdato"))

                                             if strday < 10 then
                                                strday = "0" & strday
                                             end if

                                             strdato = srtyear&"-"&strmonth&"-"&strday
                                        %>

                                         <span style="display:none;"><%=strdato %></span>
                                         <%=oRec("jobslutdato") %>
                                     </td>
                                     <%end if %>

                                     <%if cint(simple_joblist_status) = 1 then %>
                                     <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>;"><%=statusTXT %></td>
                                     <%end if %>

                                     <%if cint(simple_joblist_tidsforbrug) = 1 then %>
                                     <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>; text-align:right;">
                                         
		                                   <%=formatnumber(proaf, 2)%> 
		                           
                                     </td>
                                     <%end if %>

                                     <%if cint(simple_joblist_budgettid) = 1 then %>
                                     <td class="list_simple" style="display:<%=simpledisplay%>; background-color:<%=trbgcollastjobid%>; text-align:right;"><%=formatnumber(budgettimertot, 2)%></td>
                                     <%end if %>

                        




                        </tr>



                        
                       



                        <%

                            '** Til Export fil ***
	                        jids = jids & "," & oRec("id")
                            jobids_all = jobids_all & "," & oRec("id")
                            cnt = cnt + 1




                           end if   'if cint(visJobFomr) = 1 OR instr(strFomr_rel, "#0#") <> 0 then

   


                            oRec.movenext
                            wend
                            oRec.close 
                        %>

                      
                           <input type="hidden" id="jq_jids" value="<%=jobids_all %>" />
  
                        </tbody>
                        
                        <%

	
	                    if cnt > 0 then%>
                        <tfoot class="list_advanced" style="display:<%=advanceddisplay%>;">
    	                <tr style="background-color:#FFFFFF;">
                         <td colspan=3>&nbsp;</td>
                            <td align="right"><br /><b><%=formatnumber(jo_bruttooms_tot, 2) & " "& basisValISO %></b></td>
                            <td>&nbsp;</td>
		                <td>
                            <%if fakturaBel_tot_gt <> 0 then %>
                            <%=job_txt_326 %>: <br /><b><%=formatnumber(fakturaBel_tot_gt, 2) & " "& basisValISO %></b>
                            <%end if %>

                            <%if totRealialt <> 0 AND gnsPrisTot <> 0 then %>
                            <br><br><%=job_txt_327 %>:<br><b><%=formatnumber(gnsPrisTot/totRealialt, 2) &" "& basisValISO  %></b> 
                            <%end if %>

		                </td>
                          <td colspan=4>&nbsp;</td>
                    </tr>
                      </tfoot>
                    <%end if %>


                    
                </table>
                </div>
                </div>
         


        <br /><br />
            <%if (browstype_client <> "ip") then %>
            <section>
                <div class="row">
                     <div class="col-lg-12">
                        <b><%=kunder_txt_112 & " & " & job_txt_247%>:</b>
                        </div>
                    </div>





                <form action="../timereg/job_eksport.asp" method="post" target="_blank">
                <div class="row">
                     <div class="col-lg-9 pad-l30">
                    

                    <input id="Hidden3" name="jids" value="<%=jids%>" type="hidden" />
                    <input type="hidden" value="1" name="eksDataStd" /><br />
                    <input type="checkbox" value="1" name="xeksDataStd" checked disabled /> <%=job_txt_297 %><br />
                    <input type="checkbox" value="1" name="eksDataNrl" /> <%=job_txt_298 %><br />
                    <input type="checkbox" value="1" name="eksDataJsv" /> 
                        <%
                        call salgsans_fn()    
                        if cint(showSalgsAnv) = 1 then  %>
                        <%=job_txt_300 %>
                        <%else %>
                        <%=job_txt_330 %>
                        <%end if %><br />
                    <input type="checkbox" value="1" name="eksDataAkt" /> <%=job_txt_232 %><br />
                    <input type="checkbox" value="1" name="eksDataMile" /> <%=job_txt_299 %>
                         <%if jobasnvigv = 1 then %>
                              (<%=job_txt_301 %>)
                             <%end if %>

                         <br />
                         <input id="Submit5" type="submit" value="A) <%=job_txt_329 %> >> " class="btn btn-sm" />
                     </div>

                       </form>
                 
             
                            <div class="col-lg-3 pull-right">
                                <%



                                
	
	            strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus <> 3"
	            oRec.open strSQL, oConn, 3 
	            if not oRec.EOF then
	            antalEksterneJob = oRec("antal")
	            end if
	            oRec.close 
	
	            if len(antalEksterneJob) <> 0 then
	            antalEksterneJob = antalEksterneJob
	            else
	            antalEksterneJob = 0
	            end if
	
	
	            strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 1"
	            oRec.open strSQL, oConn, 3 
	            if not oRec.EOF then
	            antalEksterneAktiveJob = oRec("antal")
	            end if
	            oRec.close 
	
	            if len(antalEksterneJob) <> 0 then
	            antalEksterneAktiveJob = antalEksterneAktiveJob
	            else
	            antalEksterneAktiveJob = 0
	            end if
	
	            strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 0"
	            oRec.open strSQL, oConn, 3 
	            if not oRec.EOF then
	            antalEksterneLukkedeJob = oRec("antal")
	            end if
	            oRec.close 
	
	            if len(antalEksterneJob) <> 0 then
	            antalEksterneLukkedeJob = antalEksterneLukkedeJob
	            else
	            antalEksterneLukkedeJob = 0
	            end if
	
	
	            '*** Tilbud ***'
	            strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 3"
	            oRec.open strSQL, oConn, 3 
	            if not oRec.EOF then
	            antalEksternePassiveJob = oRec("antal")
	            end if
	            oRec.close 
	
	            if len(antalEksterneJob) <> 0 then
	            antalEksternePassiveJob = antalEksternePassiveJob
	            else
	            antalEksternePassiveJob = 0
	            end if
	
	            antalTilbud = antalEksternePassiveJob
	

                '*** Gennemsyn ***'
	            gennemsyn = 0
                strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 4"
	            oRec.open strSQL, oConn, 3 
	            if not oRec.EOF then
	            gennemsyn = oRec("antal")
	            end if
	            oRec.close 
	
	            if len(gennemsyn) <> 0 then
	            gennemsyn = gennemsyn
	            else
	            gennemsyn = 0
	            end if


                '*** Til fakturering ***'
	            tilfakturering = 0
                strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 2"
	            oRec.open strSQL, oConn, 3 
	            if not oRec.EOF then
	            tilfakturering = oRec("antal")
	            end if
	            oRec.close 
	
	            if len(tilfakturering) <> 0 then
	            tilfakturering = tilfakturering
	            else
	            tilfakturering = 0
	            end if
	

                '*** Eval ***'
                antalEvalJob = 0
	            strSQL = "SELECT count(id) AS antal FROM job WHERE jobstatus = 5"
	            oRec.open strSQL, oConn, 3 
	            if not oRec.EOF then
	            antalEvalJob = oRec("antal")
	            end if
	            oRec.close 
	
	            if len(antalEvalJob) <> 0 then
	            antalEvalJob = antalEvalJob
	            else
	            antalEvalJob = 0
	            end if
	
	            antalTilbud = antalEksternePassiveJob

	
	                            uTxt = "<b>"&job_txt_311&":</b><br><b>"& antalEksterneAktiveJob & "</b> "&jobstatus_txt_001 _
	                            &"<br><b>" & antalEksterneLukkedeJob &"</b> "&jobstatus_txt_005 _
	                            &"<br><b>" & antalTilbud & "</b> "& jobstatus_txt_003 _
                                &"<br><b>" & gennemsyn & "</b> "& jobstatus_txt_004 _
                                &"<br><b>" & tilfakturering & "</b> "& jobstatus_txt_002 _
                                &"<br><b>" & antalEvalJob & "</b> "& job_txt_629   
	
                                if cint(vis_timepriser) <> 1 then
                                uTxt = uTxt  &"<br><b>" & cnt & "</b> "& job_txt_310
	                            else
                                uTxt = uTxt  &"<br><b>" & cnt & "</b> "& job_txt_313
                                end if

           
                                Response.write uTxt

                                    %>
                            </div>
                

                </div>
 
             


                <br />
                <form action="../timereg/job_eksport.asp?optiprint=3" method="post" target="_blank">
                  
                <div class="row">
                     <div class="col-lg-12 pad-l30">
                   
                      <input id="Submit3" type="submit" value="B) <%=job_txt_302 %> >> " class="btn btn-sm" />
                     </div>
                </div>
                <input id="Hidden5" name="jids" value="<%=jids%>" type="hidden" />
                <%sorttp = 1 %>
                <input id="Hidden6" name="sorttp" value="<%=sorttp%>" type="hidden" />
              
                </form>


                <form action="../timereg/job_eksport.asp?optiprint=1" method="post" target="_blank">
               <input id="Hidden4" name="jids" value="<%=jids%>" type="hidden" />
                   
                

                     <div class="row">
                     <div class="col-lg-12 pad-l30">
                       <input id="Submit6" type="submit" value="C) <%=job_txt_303 %> >>" class="btn btn-sm" />
                     
                     </div>
                </div>
                </form>




            <%if instr(lto, "epi") <> 0 OR lto = "intranet - local" then %>
            <form action="../timereg/job_eksport.asp?optiprint=6" method="post" target="_blank">
            <input id="Hidden5" name="jids" value="<%=jids%>" type="hidden" />
             

                 <div class="row">
                     <div class="col-lg-12 pad-l30">
                     <input id="Submit7" type="submit" value="D) <%=job_txt_304 %> >> " class="btn btn-sm" />
                </div>
                </div>

            </form>

            <%end if %>

            <%if lto = "dencker" OR lto = "dencker_test" OR lto = "intranet - local" OR lto = "cflow" then %>
            <form action="../timereg/job_eksport.asp?optiprint=7" method="post" target="_blank">
           <input id="Hidden5" name="jids" value="<%=jids%>" type="hidden" />
                
                  <div class="row">
                     <div class="col-lg-12 pad-l30">
                       <input id="Submit7" type="submit" value="E) <%=job_txt_306 %> >> " class="btn btn-sm" />
                    
                </div>
                </div>

                  
            </form>
            <%end if	
	
	
      


                     if lto = "outz" OR lto = "intranet - local" then %>
                      <br /><br />
                    <form action="../timereg/createJobFromFile.aspx?lto=<%=lto%>&editor=<%=session("user")%>" method="post" target="_blank">
                        <div class="row">
                            <div class="col-lg-3 pad-l30">
                                <button class="btn btn-sm btn-success"><b><%=job_txt_308 %> +</b></button><br />&nbsp;
                            </div>
                        </div>
                    </form>

                <br /><br />
                 <%end if

         %>
         </section>
        <%end if 'browsertype %>


           </div>
        </div>
<%if (browstype_client <> "ip") then %>
    </div>
</div>
<%end if %>

<%end select %>
<!--#include file="../inc/regular/footer_inc.asp"-->