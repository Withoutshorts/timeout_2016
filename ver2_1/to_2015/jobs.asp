<%response.buffer = true%>
<%thisfile = "jobs" %>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../timereg/inc/job_inc.asp"-->
<!--#include file="../inc/regular/cls_multipanel.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"--> 

<meta name="viewport" content='width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0'>

<%
'** Jquery section 
if Request.Form("AjaxUpdateField") = "true" then



Select Case Request.Form("control")
    case "FileuploadKlargoring"
            
             jobid = request("jobid")

             strSQLupdnt = "INSERT INTO filer (jobid, type, oprses, adg_alle) "_
             & " VALUES ("& jobid &", 5, 'x', 1)" 

            oConn.execute (strSQLupdnt)

            'response.Write "VAlgt jobid er " & jobid

    case "FN_kpers"
              
              dim kundeKpersArr
              redim kundeKpersArr(75)
              
              if len(trim(request("jq_kid"))) <> 0 AND request("jq_kid") <> 0 then
              kidThis = request("jq_kid")
              else
              kidThis = 0
              end if

              if len(trim(request("jq_kundekpers"))) <> 0 AND request("jq_kundekpers") <> 0 then
              kundekpersThis = request("jq_kundekpers")
              else
              kundekpersThis = 0
              end if
       
            
            

                strSQL = "SELECT navn, id, titel, email FROM kontaktpers WHERE kundeid = "& kidThis
			    
                'Response.write "<option>"& strSQL &"</option>"
                'Response.end
                

                kundeKpersArr(0) = "<option value='0'>Ingen</option>"
                z = 1

                oRec.open strSQL, oConn, 3
			   
			    while not oRec.EOF
				
			    
                if cint(kundekpersThis) = oRec("id") then
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
                kundeKpersArr(z) = "<option value='0'>Ingen kontaktpersoner fundet</option>"
            end if          

               for z = 0 to UBOUND(kundeKpersArr)
              '*** ��� **'
              call jq_format(kundeKpersArr(z))
              kundeKpersArr(z) = jq_formatTxt
	          
	          Response.Write kundeKpersArr(z)
              next



    case "FN_sogjobogkunde"


        call selectJobogKunde_jq()


        
    case "FN_sogakt"

               
        call selectAkt_jq
		

end select
Response.end
end if
%>




<!--------------- Github 1 ----------------->


<% 
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if


    func = request("func")
	if func = "dbopr" then
	jobid = 0
	else
	jobid = request("jobid")
	end if


    response.cookies("2015")("lastjobid") = jobid

    function SQLBlessP(s)
	dim tmp2
	tmp2 = s
	tmp2 = replace(tmp2, ",", ".")
	SQLBlessP = tmp2
	end function

    function SQLBless2(s)
    dim tmp
    tmp = s
    tmp = replace(tmp, "'", "''")
    SQLBless2 = tmp
    end function

    call browsertype()

    Select case func 

        case "deleteFile"

            fileid = request("fileid")
            jobid = request("jobid")

            if len(trim(fileId)) <> 0 then
                strPath = "d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\upload\"&lto&"\" & Request("filnavn")

                on Error resume Next 

	            Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	            Set fsoFile = FSO.GetFile(strPath)
	            fsoFile.Delete

            
	            oConn.execute("DELETE FROM filer WHERE id = "& fileid &"")
               ' response.Write "id er: " & id
            end if 

	        Response.redirect "jobs.asp?func=red&jobid="&jobid
            

        case "oprMat"

                matregid = 0
                otf = 1
                medid = session("mid") 'Man kan pt. kun registre for sig selv
                jobid = request("FM_jobid")
                aktid = 0 'pt kan man ikke v;lge aktivitetm skal man kunne det?
                aftid = 0
                matId = 0
                strEditor = session("user")
                strDato = year(now)&"/"&month(now)&"/"&day(now)
                intAntal = request("FM_matantal")
                intAntal = replace(intAntal,".",",") 
                regdato = year(request("FM_matdato"))&"/"&month(request("FM_matdato"))&"/"&day(request("FM_matdato"))
                valuta = 1 'Default er 1 skal det kunne aendres
            
                select case lto 
                case "synergi1", "intranet - local", "epi", "epi_as", "epi_se", "epi_os", "epi_uk", "alfanordic"
                intKode = 1 'Intern
                case else 
	            intKode = 2 'Ekstern = viderefakturering
                end select

                personlig = 0 'Pt firfa betalt som default
                bilagsnr = ""

                pris = request("FM_matpris")
                pris = replace(pris,".",",")
                salgspris = pris 'Skal denne have sin egen
        
                navn = request("FM_matnavn")
                gruppe = 0 'pt kan man ikke v;lge gruppe, skal man kunne det?
                varenr = 0 
                opretlager = 0
                betegnelse = ""
                mat_func = "dbopr"
                matreg_opdaterpris = 0
                matava = 0

                mattype = 0
                basic_valuta = request("basic_valuta")
                basic_kurs = request("basic_kurs")
                basic_belob = request("basic_belob")

                call indlaes_mat(matregid, otf, medid, jobid, aktid, aftid, matId, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, mat_func, matreg_opdaterpris, matava, mattype, basic_valuta, basic_kurs, basic_belob)


                response.Redirect "jobs.asp?func=red&jobid="&jobid

        case "sendchatmessage"


                chatmessage = request("FM_chatmessage")
                editdate = year(now) &"-"& month(now) &"-"& day(now)
                edittime = hour(now) &":"& minute(now) &":"& second(now)

                strSQL = "INSERT INTO chat SET jobid = "& jobid &", message = '"& chatmessage &"', editor = "& session("mid") & ", editdate = '"& editdate &"', edittime = '"& edittime &"'"
                response.Write strSQL
                oConn.execute(strSQL)

                response.Redirect "jobs.asp?func=red&jobid="&jobid

    case "dbopr", "dbred"

                                '********************** Opretter job ***************************


                                kid = request("FM_kunde")
                                varJobId = 0

                                if len(trim(request("FM_kpers"))) then
                                kunderef = request("FM_kpers")
                                else
                                kunderef = request("FM_kpers") = 0
                                end if   
        
                                jobnavn = request("FM_jobnavn")
                                jobnavn = replace(jobnavn, "'", "")

                                if jobnavn = "" then
                                    errortype = 207
		                            call showError(errortype)
		                            Response.end
                                end if

                                jobnr = request("FM_jobnr")
                                beskrivelse = request("FM_beskrivelse")
                                internnote = request("FM_internnote")
                                bruttooms = 0
                                editor = session("user")
                                dddato = year(now) & "-" & month(now) & "-" & day(now)
                                strdato = session("dato")

                                if len(request("FM_serviceaft")) <> 0 then
				                intServiceaft = request("FM_serviceaft") '1
				                else
				                intServiceaft = 0
				                end if
                
                                strOLDjobnr = request("FM_OLDjobnr")

                                jobstartdato = request("FM_jobstartdato")
                                jobstartdato = year(jobstartdato) & "/" & month(jobstartdato) & "/" & day(jobstartdato)
                               

                                if request("FM_datouendelig") = "j" then
					                jobslutdato =  "2044-1-1" 
					            else
                                    jobslutdato = request("FM_jobslutdato")
                                    jobslutdato = year(jobslutdato) & "/" & month(jobslutdato) & "/" & day(jobslutdato)
                                end if

                                 startdatoNum = day(jobstartdato) & "/" & month(jobstartdato) & "/" & year(jobstartdato)
                                 slutDatoNum = day(jobslutdato) & "/" & month(jobslutdato) & "/" & year(jobslutdato)

                                'Response.write "startdatoNum: "& startdatoNum
                                'Response.write "<br>slutDatoNum: "& slutDatoNum


                                '** Tjekker om SLUTDATO udfyldt korrekt ** EPI pga auto opret og startdato starter i minus 1
		                        if (cdate(startdatoNum) > cdate(slutDatoNum) AND instr(lto, "epi") = 0) OR _
                                (cdate(startdatoNum) > cdate(slutDatoNum) AND func = "dbred" AND instr(lto, "epi") <> 0) then
		
		                        'call visErrorFormat
		                        errortype = 184
		                        call showError(errortype)
		                        Response.end
                                end if

                                'jobans1 = request("FM_jobans1")
                                'jobans2 = 0

                                'rekvisitionsnr = ""


                                'Status
                                if request("FM_usetilbudsnr") = "j" then
						        strStatus = 3 'tilbud
						        else
							        if request("FM_status") <> 3 then
							        strStatus = request("FM_status")
							        else
							        strStatus = 1 'aktivt, hvis der ved en fejl at valgt tilbud
							        end if
						        end if

                                '*********  Sandsynlighed ***********************
                                if len(trim(request("FM_sandsynlighed"))) <> 0 then
			                    sandsynlighed = formatnumber(request("FM_sandsynlighed"), 0)
			                    else
			                    sandsynlighed = 0
			                    end if
                                %><!--#include file="../timereg/inc/isint_func.asp"--><%
			                    isInt = 0    
		                        call erDetInt(request("FM_sandsynlighed"))
                                if isInt > 0 OR (lto = "epi2017" AND strStatus = 3 AND (len(trim(sandsynlighed ) ) = 0 OR sandsynlighed > 100 OR sandsynlighed < 1 )) then
                                'call visErrorFormat
    		
		                        errortype = 31
		                        call showError(errortype)
		                        response.End
			                    end if


                                '0: lbn. timer / 1: fastpris 
                                if len(trim(request("FM_fastpris"))) <> 0 then 
				                strFastpris = request("FM_fastpris") 
				                else
				                strFastpris = 0
				                end if


                                '***** Jobansvarlige ***'
                                if len(trim(request("FM_jobans_1"))) <> 0 then
				                intJobans1 = request("FM_jobans_1")
                                else
                                intJobans1 = 0
                                end if

                                if len(trim(request("FM_jobans_2"))) <> 0 then
				                intJobans2 = request("FM_jobans_2")
                                else
                                intJobans2 = 0
                                end if

                                if len(trim(request("FM_jobans_3"))) <> 0 then
				                intJobans3 = request("FM_jobans_3")
                                else
                                intJobans3 = 0
                                end if

                                 if len(trim(request("FM_jobans_4"))) <> 0 then
				                intJobans4 = request("FM_jobans_4")
                                else
                                intJobans4 = 0
                                end if

                                 if len(trim(request("FM_jobans_5"))) <> 0 then
				                intJobans5 = request("FM_jobans_5")
                                else
                                intJobans5 = 0
                                end if


                                if len(trim(request("FM_jobans_proc_1"))) <> 0 then
				                jobans_proc_1 = request("FM_jobans_proc_1")
                                else
                                jobans_proc_1 = 0
                                end if
               
				                if len(trim(request("FM_jobans_proc_2"))) <> 0 then
				                jobans_proc_2 = request("FM_jobans_proc_2")
                                else
                                jobans_proc_2 = 0
                                end if
                
				                if len(trim(request("FM_jobans_proc_3"))) <> 0 then
				                jobans_proc_3 = request("FM_jobans_proc_3")
                                else
                                jobans_proc_3 = 0
                                end if
               
				                if len(trim(request("FM_jobans_proc_4"))) <> 0 then
				                jobans_proc_4 = request("FM_jobans_proc_4")
                                else
                                jobans_proc_4 = 0
                                end if
               
				                if len(trim(request("FM_jobans_proc_5"))) <> 0 then
				                jobans_proc_5 = request("FM_jobans_proc_5")
                                else
                                jobans_proc_5 = 0
                                end if


                                for i = 1 to 5
                                  select case i 
                                  case 1
                                  procVal = jobans_proc_1
                                  case 2
                                  procVal = jobans_proc_2
                                  case 3
                                  procVal = jobans_proc_3
                                  case 4
                                  procVal = jobans_proc_4
                                  case 5
                                  procVal = jobans_proc_5
                                  end select
                    
                   
                    
                                call erDetInt(procVal)
				    
                                if isInt > 0 then
                                errProcVal = 1
                                isInt = 0

                                else

                                if instr(lto, "epi2017") <> 0 OR lto = "intranet - local" then 
                                jobProcent100 = jobProcent100/1 + procVal 'Skal divideres med 10 da der fjernes komma
                                end if  

                                end if
                
                            next

                            jobans_proc_1 = replace(jobans_proc_1, ".", "")
                            jobans_proc_1 = replace(jobans_proc_1, ",", ".")
                            jobans_proc_2 = replace(jobans_proc_2, ".", "")
                            jobans_proc_2 = replace(jobans_proc_2, ",", ".")
                            jobans_proc_3 = replace(jobans_proc_3, ".", "")
                            jobans_proc_3 = replace(jobans_proc_3, ",", ".")
                            jobans_proc_4 = replace(jobans_proc_4, ".", "")
                            jobans_proc_4 = replace(jobans_proc_4, ",", ".")
                            jobans_proc_5 = replace(jobans_proc_5, ".", "")
                            jobans_proc_5 = replace(jobans_proc_5, ",", ".")


                            '******** Salgsansvarlige **********'
                            if len(trim(request("FM_salgsans_1"))) <> 0 then ' er slags ansvarlige sl�et til / vist
                            salgsans1 = request("FM_salgsans_1")

                            if len(trim(request("FM_salgsans_2"))) <> 0 then
				            salgsans2 = request("FM_salgsans_2")
                            else
                            salgsans2 = 0
                            end if

				            if len(trim(request("FM_salgsans_3"))) <> 0 then
				            salgsans3 = request("FM_salgsans_3")
                            else
                            salgsans3 = 0
                            end if

				            if len(trim(request("FM_salgsans_4"))) <> 0 then
				            salgsans4 = request("FM_salgsans_4")
                            else
                            salgsans4 = 0
                            end if

				            if len(trim(request("FM_salgsans_5"))) <> 0 then
				            salgsans5 = request("FM_salgsans_5")
                            else
                            salgsans5 = 0
                            end if


                            else
                            salgsans1 = 0
				            salgsans2 = 0
				            salgsans3 = 0
				            salgsans4 = 0
				            salgsans5 = 0
                            end if


                            if len(trim(request("FM_salgsans_proc_1"))) <> 0 then
				            salgsans_proc_1 = request("FM_salgsans_proc_1")
                            else
                            salgsans_proc_1 = 0
                            end if
               
				            if len(trim(request("FM_salgsans_proc_2"))) <> 0 then
				            salgsans_proc_2 = request("FM_salgsans_proc_2")
                            else
                            salgsans_proc_2 = 0
                            end if
               
				            if len(trim(request("FM_salgsans_proc_3"))) <> 0 then
				            salgsans_proc_3 = request("FM_salgsans_proc_3")
                            else
                            salgsans_proc_3 = 0
                            end if
              
                            if len(trim(request("FM_salgsans_proc_4"))) <> 0 then
				            salgsans_proc_4 = request("FM_salgsans_proc_4")
                            else
                            salgsans_proc_4 = 0
                            end if

				            if len(trim(request("FM_salgsans_proc_5"))) <> 0 then
				            salgsans_proc_5 = request("FM_salgsans_proc_5")
                            else
                            salgsans_proc_5 = 0
                            end if

               
                 
                            for i = 1 to 5
                                  select case i 
                                  case 1
                                  procVal = salgsans_proc_1
                                  case 2
                                  procVal = salgsans_proc_2
                                  case 3
                                  procVal = salgsans_proc_3
                                  case 4
                                  procVal = salgsans_proc_4
                                  case 5
                                  procVal = salgsans_proc_5
                                  end select
                    

                      

                                call erDetInt(procVal)
				    
                                if isInt > 0 then
                                errProcVal = 1
                                isInt = 0
                                else

                                if instr(lto, "epi2017") <> 0 OR lto = "xintranet - local"  then 
                                salgsProcent100 = salgsProcent100/1 + procVal 'Skal divideres med 10 da der fjernes komma
                                end if 

                                end if
                
                            next
                
                
                            salgsans_proc_1 = replace(salgsans_proc_1, ".", "")
                            salgsans_proc_2 = replace(salgsans_proc_2, ".", "")
                            salgsans_proc_3 = replace(salgsans_proc_3, ".", "")
                            salgsans_proc_4 = replace(salgsans_proc_4, ".", "")
                            salgsans_proc_5 = replace(salgsans_proc_5, ".", "")

                            salgsans_proc_1 = replace(salgsans_proc_1, ",", ".")
                            salgsans_proc_2 = replace(salgsans_proc_2, ",", ".")
                            salgsans_proc_3 = replace(salgsans_proc_3, ",", ".")
                            salgsans_proc_4 = replace(salgsans_proc_4, ",", ".")
                            salgsans_proc_5 = replace(salgsans_proc_5, ",", ".")





                            '***** Forrretningsomr�der **********'

                            fomrArr = split(request("FM_fomr"), ",")

                            for_faktor = 0
                            'for afor = 0 to UBOUND(fomrArr)
                            'for_faktor = for_faktor + 1 
                            'next

                            'if for_faktor <> 0 then
                            'for_faktor = for_faktor
                            'else
                            'for_faktor = 1
                            'end if

                            'for_faktor = formatnumber(100/for_faktor, 2)
                            'for_faktor = replace(for_faktor, ",", ".")

                            'if len(trim(request("FM_fomr_syncakt"))) <> 0 then
                            'syncForAkt = 1
                            'else
                            syncForAkt = 0
                            'end if
                                    
                         '****************************************


                            '** forretningsomr�de mandatory 
                            call jobopr_mandatory_fn()
                            if cint(fomr_mandatoryOn) = 1 AND len(trim(request("FM_fomr"))) = 0 then

                            'call visErrorFormat
			                errortype = 177
			                call showError(errortype)
			                Response.end


                            end if

                            rekvnr = replace(request("FM_rekvnr"), "'", "''")

                            if len(trim(rekvnr)) = 0 AND lto = "ets-track" then
		                    'call visErrorFormat
				            errortype = 149
                            call showError(errortype)
                            Response.end
		                    end if


                            if len(trim(request("FM_syncslutdato"))) <> 0 then
                            syncslutdato = 1
                            else
                            syncslutdato = 0
                            end if

                            if len(trim(request("FM_syncaktdatoer"))) <> 0 then
                            syncaktdatoer = 1
                            else
                            syncaktdatoer = 0
                            end if





                                if func = "dbopr" then



                                '**********************************'
				                '*** nyt jobnr / tilbudsnr ***'
				                '**********************************'
				                strSQL = "SELECT jobnr, tilbudsnr FROM licens WHERE id = 1"
				                oRec5.open strSQL, oConn, 3
				                if not oRec5.EOF then 
				    
				                    jobnr = oRec5("jobnr") + 1
				    
				                    if request("FM_usetilbudsnr") = "j" then
				                    tlbnr = oRec5("tilbudsnr") + 1
				                    else
				                    tlbnr = 0
				                    end if
				    
				                end if
				                oRec5.close


                                strSQLjob = ("INSERT INTO job (jobnavn, jobnr, jobknr, jobTpris, jobstatus, sandsynlighed, tilbudsnr, jobstartdato," _
                                & " jobslutdato, editor, dato, creator, createdate, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, " _
                                & " projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, " _
                                & " fakturerbart, budgettimer, fastpris, kundeok, beskrivelse, " _
                                & " ikkeBudgettimer, jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5," _
                                & " salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc, "_
                                & " serviceaft, kundekpers, syncslutdato, valuta, rekvnr, " _
                                & " risiko, job_internbesk, " _
                                & " jo_bruttooms, fomr_konto) VALUES " _
                                & "('" & SQLBless2(jobnavn) & "', " _
                                & "'" & jobnr & "', " _
                                & "" & kid & ", " _
                                & "0, " _
                                & "" & strStatus & ", " _
                                & "" & sandsynlighed & ", " _
                                & "'" & tlbnr & "', " _
                                & "'" & jobstartdato & "', " _
                                & "'" & jobslutdato & "', " _
                                & "'" & editor & "', " _
                                & "'" & strdato & "', " _
                                & "'" & editor & "', " _
                                & "'" & year(now) &"-"& month(now) &"-"& day(now) & "', " _
                                & "10, " _
                                & "1,1,1,1,1,1,1,1,1," _
                                & "1,0,"& strFastpris &",0," _
                                & "'" & SQLBless2(beskrivelse) & "', " _
                                & "0, " _
                                & "" & intJobans1 & ", " _
                                & "" & intJobans2 & ", " _
                                & "" & intJobans3 & ", " _
                                & "" & intJobans4 & ", " _
                                & "" & intJobans5 & ", " _
                                & "" & jobans_proc_1 & ", " _
                                & "" & jobans_proc_2 & ", " _
                                & "" & jobans_proc_3 & ", " _
                                & "" & jobans_proc_4 & ", " _
                                & "" & jobans_proc_5 & ", " _
                                &" "& salgsans1 &","& salgsans2 &","& salgsans3 &","& salgsans4 &","& salgsans5 &", "_
                                &" "& salgsans_proc_1 &","& salgsans_proc_2 &","& salgsans_proc_3 &","& salgsans_proc_4 &","& salgsans_proc_5 &", "_
                                &" "& intServiceaft &", "& kunderef & ", " _
                                & "" & syncslutdato & ", " _
                                & "1, '" & rekvnr & "', " _
                                & "100,'" & internnote & "'," _
                                & "" & bruttooms & ", 0)")
    							


                                response.write "strSQLjob: " & strSQLjob 
                                'response.flush

                                oConn.execute(strSQLjob)


                                '******************************************'
								'*** finder jobid p� det netop opr. job ***'
								'******************************************'
                                't = 10
								'if t = 0 then
								    lastID = 0
								    strSQL2 = "SELECT id FROM job WHERE id <> 0 ORDER BY id DESC " 'jobnr = " & strjnr
									oRec5.open strSQL2, oConn, 3
									if not oRec5.EOF then
									varJobIdThis = oRec5("id")
									varJobId = varJobIdThis
                                    lastID = varJobIdThis
									end if
									oRec5.close        


                                
                                '***************************************************************************************************
                                'PUSH - Aktiviteter - Timepriser
                                 '*********** timereg_usejob, s� der kan s�ges fra jobbanken KUN VED OPRET JOB *********************
                                Select Case lto
                                    Case "oko"
                                    Case else '"outz", "intranet - local", "demo"
                       
                                        strProjektgr1 = 10
                                        strProjektgr2 = 1
                                        strProjektgr3 = 1
                                        strProjektgr4 = 1
                                        strProjektgr5 = 1
                                        strProjektgr6 = 1
                                        strProjektgr7 = 1
                                        strProjektgr8 = 1
                                        strProjektgr9 = 1
                                        strProjektgr10 = 1
                                
                              
							
                                        strSQLpg = "SELECT MedarbejderId FROM progrupperelationer WHERE (" _
                                        & " ProjektgruppeId = " & strProjektgr1 & "" _
                                        & " OR ProjektgruppeId =" & strProjektgr2 & "" _
                                        & " OR ProjektgruppeId =" & strProjektgr3 & "" _
                                        & " OR ProjektgruppeId =" & strProjektgr4 & "" _
                                        & " OR ProjektgruppeId =" & strProjektgr5 & "" _
                                        & " OR ProjektgruppeId =" & strProjektgr6 & "" _
                                        & " OR ProjektgruppeId =" & strProjektgr7 & "" _
                                        & " OR ProjektgruppeId =" & strProjektgr8 & "" _
                                        & " OR ProjektgruppeId =" & strProjektgr9 & "" _
                                        & " OR ProjektgruppeId =" & strProjektgr10 & "" _
                                        & ") GROUP BY MedarbejderId"
								
                                        'Response.Write "strSQL "& strSQL & "<br><hr>"

                                        oRec2.open strSQLpg, oConn, 3
                                        While not oRec2.EOF
                
                                
                                            strSQL3 = "INSERT INTO timereg_usejob (medarb, jobid, forvalgt, forvalgt_sortorder, forvalgt_af, forvalgt_dt) VALUES " _
                                            & " (" & oRec2("MedarbejderId") & ", " & lastID & ", 1, 0, 0, '" & dddato & "')"

                                            oConn.execute(strSQL3)
                                                
                                       
									    oRec2.movenext
                                        Wend 
                                
                                        oRec2.Close()
                                
                                
                                End Select ' lto
                        
                        
                
                                '** V�lg Stamaktgruppe pbg. af projetktype
                                agforvalgtStamgrpKri = " ag.forvalgt = 1 "
                                
                                
                                
                                '*** Indl�ser STAMaktiviteter ***'
                                strSQLStamAkt = "SELECT a.navn AS aktnavn, aktkonto, fase, a.id AS stamaktid FROM akt_gruppe AS ag" _
                                & " LEFT JOIN aktiviteter AS a ON (a.aktfavorit = ag.id) WHERE " & agforvalgtStamgrpKri & " AND skabelontype = 0 ORDER BY a.sortorder DESC"
                              
                                oRec3.open strSQLStamAkt, oConn, 3 
                                While not oRec3.EOF 
                
                            
                            
                            
                    
                    
                                    strSQLaktins = "INSERT INTO aktiviteter (navn, job, fakturerbar, " _
                                    & "projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7," _
                                    & "projektgruppe8, projektgruppe9, projektgruppe10, aktstatus, budgettimer, aktbudget, aktbudgetsum, aktstartdato, aktslutdato, aktkonto, fase) VALUES " _
                                    & " ('" & oRec3("aktnavn") & "', " & lastID & ", 1," _
                                    & " 10,1,1,1,1,1,1,1,1,1,1,0,0,0,'" & jobstartdato & "', " _
                                    & "'" & jobslutdato & "', " & oRec3("aktkonto") & ", '" & oRec3("fase") & "')"
                                               
                                    oConn.execute(strSQLaktins)
                    
                    
                                    '*** Finder aktid ***
                                    strSQLlastAktID = "SELECT id FROM aktiviteter WHERE id <> 0 ORDER BY id DESC"
                                    oRec2.open strSQLlastAktID, oConn, 3
                                    lastAktID = 0
                                    If not oRec2.EOF Then
                
                                        lastAktID = oRec2("id")
                
                                    End If
                                    oRec2.close()
                    
                    
                    
                                    '** FOMR REL ***********
                                    strSQLaktfomr = "SELECT for_fomr, for_aktid FROM fomr_rel WHERE for_aktid =  " & oRec3("stamaktid")
                                             
                                    oRec4.open strSQLaktfomr, oConn, 3
                                    While not oRec4.EOF
                                
                                
                                
                                        strSQLaktinsfomr = "INSERT INTO fomr_rel (for_fomr, for_aktid, for_jobid, for_faktor) VALUES  (" & oRec4("for_fomr") & ", " & lastAktID & ", " & lastID & ", 100)"
                                               
                                        oConn.execute(strSQLaktinsfomr)
                                
                                
                                    oRec4.movenext
                                    Wend
                                    oRec4.close()
                            
                            
                    
                                    '**************************************************************'
                                    '*** Hvis timepris ikke findes p� job bruges Gen. timepris fra '
                                    '*** Fra medarbejdertype, og den oprettes p� job              *'
                                    '*** BLIVER ALTID HENTET FRA Medarb.type for �KO              *'
                                    '**************************************************************'
                                    tprisGen = 0
                                    valutaGen = 1
                                    kostpris = 0
                                    intTimepris = 0
                                    intValuta = valutaGen
                    
                                    SQLmedtpris = "SELECT medarbejdertype, timepris, timepris_a2, timepris_a3, tp0_valuta, kostpris, mnavn, mid FROM medarbejdere " _
                                    & " LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) " _
                                    & " WHERE mid <> 0 AND mansat = 1 AND medarbejdertyper.id = medarbejdertype GROUP BY mid"

                                    oRec2.open SQLmedtpris, oConn, 3
            
                                    While Not oRec2.EOF 
		 	
                                        If oRec2("kostpris") <> 0 Then
                                            kostpris = oRec2("kostpris")
                                        Else
                                            kostpris = 0
                                        End If
		
                            
                                                If oRec2("timepris") <> 0 Then
                                     
                                                    tprisGen = oRec2("timepris")
                                                Else
                                                    tprisGen = 0
                                                End If
                                    
                                                timeprisalt = 0
                           
                            
                                   
                        
                                        valutaGen = oRec2("tp0_valuta")
                        
                        
                                        '**** Indl�ser timepris p� aktiviteter ***'
                                        intTimepris = tprisGen
                                        intTimepris = Replace(intTimepris, ",", ".")
                            
                                        intValuta = valutaGen
							
                                        strSQLtpris = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) " _
                                        & " VALUES (" & lastID & ", " & lastAktID & ", " & oRec2("mid") & ", " & timeprisalt & ", " & intTimepris & ", " & intValuta & ")"
							
                                        oConn.execute(strSQLtpris)
                        
		                            oRec2.movenext
                                    wend

                                    oRec2.close()
                                    '** Slut timepris **
                
                
                   
							
                    
                    
                                oRec3.movenext
                                wend 
                                oRec3.close()
                

                                '*** End PUSH - Aktiviteter - Timepriser



















                                '*************************************************'
			                    '**** Opdaterer jobnr og tilbuds nr i Licens *****'
				                '*************************************************'
				                nytjobnr = jobnr
				                tilbudsnrKri = ""
				                
				                if request("FM_usetilbudsnr") = "j" then
				                nyttlbnr = tlbnr
				                tilbudsnrKri = ", tilbudsnr = "& nyttlbnr &""
				                end if
				                
				                strSQL = "UPDATE licens SET jobnr = "&  nytjobnr &" "& tilbudsnrKri &" WHERE id = 1"
				                oConn.execute(strSQL)





                                else 'hvis rediger



                                'Tjekker om jobnr findes
                                strSQL = "SELECT jobnr, id FROM job WHERE id <> "& jobid &" AND jobnr = '" & jobnr & "'"
                                response.Write strSQL
                                oRec5.open strSQL, oConn, 3
		                        if not oRec5.EOF then	
		            
		                        jobnrFindesNR = oRec5("jobnr")
		                        jobnrFindesID = oRec5("id")
		                        jobnrFindes = 1 
		            
		                        end if
					            oRec5.close
					
					            'Response.Write "<br> jobnrFindes" &  jobnrFindes & " "& jobnrFindesNR &" "& jobnrFindesID &"<br>"
					            'Response.flush
					
					            if cint(jobnrfindes) = 1 then
					                %>
					                <!--#include file="../inc/regular/header_inc.asp"-->
				                    <%	
					                errortype = 93
					                call showError(errortype)
					                Response.end
				                end if



                                call alfanumerisk(jobnr)
                                jobnr = alfanumeriskTxt
                                jobnr = left(jobnr,20)

				                tlbnr = request("FM_tnr")  


                                        
                                '*** Opdater LUKKE dato (f�r det skifter stastus) ***'
                                call lukkedato(jobid, strStatus)


                                strSQL = "UPDATE job SET "_
                                &" jobnavn = '"& SQLBless2(jobnavn) & "',"_
                                &" jobnr = '"& jobnr &"', "_
                                &" jobknr = "& kid &", "_
                                &" jobstatus = "& strStatus & ", "_
                                &" sandsynlighed = "& sandsynlighed & ", "_
                                &" tilbudsnr = "& tlbnr & ", "_
                                &" kundekpers = " & kunderef & ", "_
                                &" jobstartdato = '" & jobstartdato & "', "_
                                &" jobslutdato = '" & jobslutdato & "', "_
                                &" dato = '" & strdato & "', "_                                    
                                &" editor = '"& editor & "', "_
                                &" jobans1 = "& intJobans1 & ", "_
                                &" jobans2 = "& intJobans2 & ", "_
                                &" jobans3 = "& intJobans3 & ", "_
                                &" jobans4 = "& intJobans4 & ", "_
                                &" jobans5 = "& intJobans5 & ", "_
                                &" jobans_proc_1 = "& jobans_proc_1 & ", "_
                                &" jobans_proc_2 = "& jobans_proc_2 & ", "_
                                &" jobans_proc_3 = "& jobans_proc_3 & ", "_
                                &" jobans_proc_4 = "& jobans_proc_4 & ", "_
                                &" jobans_proc_5 = "& jobans_proc_5 & ", "_
                                &" salgsans1 = "& salgsans1 &", salgsans2 = "& salgsans2 &", salgsans3 = "& salgsans3 &", salgsans4 = "& salgsans4 &", salgsans5 = "& salgsans5 &", "_
                                &" salgsans1_proc = "& salgsans_proc_1 &", salgsans2_proc = "& salgsans_proc_2 &", salgsans3_proc = "& salgsans_proc_3 &", salgsans4_proc = "& salgsans_proc_4 &", salgsans5_proc = "& salgsans_proc_5 &", "_
                                &" beskrivelse = '" & SQLBless2(beskrivelse) & "', "_
                                &" rekvnr = '" & rekvnr & "', "_
                                &" syncslutdato = '" & syncslutdato & "', "_
                                &" job_internbesk = '" & internnote & "', serviceaft = "& intServiceaft &", fastpris = "& strFastpris &" "_
                                &" WHERE id = "& jobid

                                response.Write strSQL
                                    
                                oConn.execute(strSQL)

                                varJobId = jobid

                                end if 'opret / rediger

                                if func = "dbred" then



                                    'Henter kundenavn
                                    strkundenavn = ""
                                    strSQLkundenavn = "SELECT kkundenavn FROM kunder WHERE kid = "& kid
                                    oRec.open strSQLkundenavn, oConn, 3
                                    if not oRec.EOF then
                                        strkundenavn = oRec("kkundenavn")
                                    end if
                                    oRec.close


                                    '*** Overf�rer gamle timeregistreringer til ny aftale (hvis der skiftes aftale) **'
								        if request("FM_overforGamleTimereg") = "1" then
                								
								        '*** Opdaterer timereg tabellen ****
								        strSQLtimer = "UPDATE timer SET "_
								        & " Tknavn = '"& replace(strkundenavn, "'", "''") &"', Tknr = "& kid &", "_
								        & " Tjobnavn = '"& SQLBless2(jobnavn) &"', "_
								        & " Tjobnr = '"& jobnr &"', "_
								        & " fastpris = '"& strFastpris &"', "_
								        & " seraft = "& intServiceaft &" "_
								        & " WHERE Tjobnr = '"& strOLDjobnr & "'"
                								
                						'** Husk materiale forbrug ***
								        strSQLmat_forbrug = "UPDATE materiale_forbrug SET serviceaft = " & intServiceaft &""_
								        & " WHERE jobid = " & id
                								
								        oConn.execute(strSQLmat_forbrug)
                								
                								
								        else
                								
								        '*** Opdaterer timereg tabellen ****
								        strSQLtimer = "UPDATE timer SET "_
								        & " Tknavn = '"& replace(strkundenavn, "'", "''") &"', Tknr = "& kid &", "_
								        & " Tjobnavn = '"& SQLBless2(jobnavn) &"', "_
								        & " Tjobnr = '"& jobnr &"', "_
								        & " fastpris = '"& strFastpris &"' "_
								        & " WHERE Tjobnr = '"& strOLDjobnr & "'"
                								
								        end  if
                								
                								
								        'Response.write strSQLtimer
								        'Response.flush
                								
								        oConn.execute(strSQLtimer)







                                    'Rediger nuv�reende aktiviteter
                                    updated_activity = split(request("FM_updated_activity"), ",")
                                    for a = 0 to UBOUND(updated_activity)

                                        
                                        FM_up_akt_navn = "FM_akt_navn_" & updated_activity(a)
                                        FM_up_akt_navn = replace(FM_up_akt_navn, " ", "")
                                        FM_up_akt_navn = request(FM_up_akt_navn)

                                        FM_up_akt_status = "aktstatus_" & updated_activity(a)
                                        FM_up_akt_status = replace(FM_up_akt_status, " ", "")
                                        FM_up_akt_status = request(FM_up_akt_status)

                                        FM_up_akt_fakturerbart = "FM_fakturerbart_" & updated_activity(a)
                                        FM_up_akt_fakturerbart = replace(FM_up_akt_fakturerbart, " ", "")
                                        FM_up_akt_fakturerbart = request(FM_up_akt_fakturerbart)

                                        FM_up_akt_budget = "aktbudget_" & updated_activity(a)
                                        FM_up_akt_budget = replace(FM_up_akt_budget, " ", "")
                                        FM_up_akt_budget = request(FM_up_akt_budget)
                                        FM_up_akt_budget = replace(FM_up_akt_budget, ".", " ")

                                        FM_up_akt_budget = replace(FM_up_akt_budget, ".", "")
		                                FM_up_akt_budget = replace(FM_up_akt_budget, ",", ".")

                                        FM_up_akt_enhed = "aktEnhed_" & updated_activity(a)
                                        FM_up_akt_enhed = replace(FM_up_akt_enhed, " ", "")
                                        FM_up_akt_enhed = request(FM_up_akt_enhed)

                                        FM_up_stdate = "aktInputSTdate_" & updated_activity(a)
                                        FM_up_stdate = replace(FM_up_stdate, " ", "")
                                        FM_up_stdate = request(FM_up_stdate)
                                        FM_up_stdate = year(FM_up_stdate) &"-"& month(FM_up_stdate) &"-"& day(FM_up_stdate)

                                        FM_up_sldate = "aktInputSLdate_" & updated_activity(a)
                                        FM_up_sldate = replace(FM_up_sldate, " ", "")
                                        FM_up_sldate = request(FM_up_sldate)
                                        FM_up_sldate = year(FM_up_sldate) &"-"& month(FM_up_sldate) &"-"& day(FM_up_sldate)

                                        FM_up_akt_prg = "FM_akt_prg_" & updated_activity(a)
                                        FM_up_akt_prg = replace(FM_up_akt_prg, " ", "")
                                        FM_up_akt_prg = request(FM_up_akt_prg)

                                        FM_up_akt_prisenhed = "FM_akt_prisenhed_" & updated_activity(a)
                                        FM_up_akt_prisenhed = replace(FM_up_akt_prisenhed, " ", "")
                                        FM_up_akt_prisenhed = request(FM_up_akt_prisenhed)

			                            FM_up_akt_prisenhed = replace(FM_up_akt_prisenhed, ".", "")
		                                FM_up_akt_prisenhed = replace(FM_up_akt_prisenhed, ",", ".")



                                        response.Write "<br> Aktivit updateret " & updated_activity(a) & " Navn " & FM_up_akt_navn & " Ny status " & FM_up_akt_status & " faktvar " & FM_up_akt_fakturerbart & " Budget " & FM_up_akt_budget & " endhed " & FM_up_akt_enhed & " startdato " & FM_up_stdate & " slutdato " & FM_up_sldate & " prggrp1 " & FM_up_akt_prg & " prisendhed " & FM_up_akt_prisenhed


                                    'Rediger nuv�reende aktiviteter
                                        strSQLUpdateAkt = "UPDATE aktiviteter SET navn = '"& FM_up_akt_navn &"', aktstatus = "& FM_up_akt_status & ", fakturerbar = "& FM_up_akt_fakturerbart & ", budgettimer = "& FM_up_akt_budget &", bgr = "& FM_up_akt_enhed & ", aktstartdato = '"& FM_up_stdate &"', aktslutdato = '"& FM_up_sldate &"', projektgruppe1 = "& FM_up_akt_prg & ", aktbudget = "& FM_up_akt_prisenhed & " WHERE id = "& updated_activity(a)
                                        
                                        response.Write "<br> strSQLUpdateAkt " & strSQLUpdateAkt
                                        oConn.execute(strSQLUpdateAkt)

                                    next
                                    'response.End
                                    ' Opretter tilf�jede aktiviteter
                                    newActivitesId = split(request("FM_newActivities"), ",")

                                    for t = 0 to UBOUND(newActivitesId)
    
                                        'newActivityName = request("FM_newactivityName_" & newActivitesId(t))
                                        newActivityName = "FM_newactivityName_" & newActivitesId(t)
                                        newActivityName = replace(newActivityName, " ", "")
                                        newActivityName = request(newActivityName)

                                        newActivityStatus = "FM_newactivityStatus_" & newActivitesId(t)
                                        newActivityStatus = replace(newActivityStatus, " ", "")
                                        newActivityStatus = request(newActivityStatus)

                                        newActivityType = "FM_newactivityTypeSEL_" & newActivitesId(t)
                                        newActivityType = replace(newActivityType, " ", "")
                                        newActivityType = request(newActivityType)

                                        newActivitybudgetantal = "FM_newactivitybudgetantal_" & newActivitesId(t)
                                        newActivitybudgetantal = replace(newActivitybudgetantal, " ", "")
                                        newActivitybudgetantal = request(newActivitybudgetantal)
                                        newActivitybudgetantal = replace(newActivitybudgetantal, ".", "")
		                                newActivitybudgetantal = replace(newActivitybudgetantal, ",", ".")

                                        newActivityBGR = "FM_newactivityBGR_" & newActivitesId(t)
                                        newActivityBGR = replace(newActivityBGR, " ", "")
                                        newActivityBGR = request(newActivityBGR)

                                        newActivitySTdate = "FM_newactivitySTDate_" & newActivitesId(t)
                                        newActivitySTdate = replace(newActivitySTdate, " ", "")
                                        newActivitySTdate = request(newActivitySTdate)
                                        newActivitySTdateSQL = year(newActivitySTdate) &"-"& month(newActivitySTdate) &"-"& day(newActivitySTdate)

                                        newActivityENDdate = "FM_newactivityENDDate_" & newActivitesId(t)
                                        newActivityENDdate = replace(newActivityENDdate, " ", "")
                                        newActivityENDdate = request(newActivityENDdate)
                                        newActivityENDdateSQL = year(newActivityENDdate) &"-"& month(newActivityENDdate) &"-"& day(newActivityENDdate)

                                        newActivityPrg = "FM_newactivityPrgSEL_" & newActivitesId(t)
                                        newActivityPrg = replace(newActivityPrg, " ", "")
                                        newActivityPrg = request(newActivityPrg)

                                        newActivityPrisenhed = "FM_akt_prisenhed_" & newActivitesId(t)
                                        newActivityPrisenhed = replace(newActivityPrisenhed, " ", "")
                                        newActivityPrisenhed = request(newActivityPrisenhed)

                                        response.Write "<br> New activity " & newActivitesId(t) & " Name " & newActivityName & " Status " & newActivityStatus &" Budgetantal "& newActivitybudgetantal & " BGR " & newActivityBGR & " Start Date " & newActivitySTdate & " End date " & newActivityENDdate & " PRG " & newActivityPrg & " akttype " & newActivityType                                    

                                        strSQLActivities = "insert into aktiviteter (navn, aktstatus, budgettimer, bgr, aktstartdato, aktslutdato, projektgruppe1, job, fakturerbar)"_
                                        &" VALUES ('"& newActivityName &"', "& newActivityStatus &", "& SQLBlessP(newActivitybudgetantal) &", "& newActivityBGR &", '"& newActivitySTdateSQL &"', '"& newActivityENDdateSQL &"', "& newActivityPrg &", "& jobid &", "& newActivityType & ")"
                                        response.Write "<br> sql " & strSQLActivities
                                        oConn.execute(strSQLActivities)

                                        response.Write "<br> New activity " & newActivitesId(t) & " Name " & newActivityName & " Status " & newActivityStatus &" Budgetantal "& newActivitybudgetantal & " BGR " & newActivityBGR & " Start Date " & newActivitySTdate & " End date " & newActivityENDdate & " PRG " & newActivityPrg
                                    

                                    next
                                end if
                            



                                '********* Indl�ser forretningsomr�der
                                '********************************'
                    
                                '*** nulstiller job ****'
                                strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& varJobId & " AND for_aktid = 0"
                                oConn.execute(strSQLfor)


                                '*** IKKE MERE 11.3.2015 (aktiviteter t�lles altid med p� job) ***'
                                '*** nulstiller akt ved sync ****'
                                'if cint(syncForAkt) = 1 then
                                '    strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& varJobId & " AND for_aktid <> 0"
                                '    oConn.execute(strSQLfor)
                                'end if

                                'Response.Write "her"
                                'Response.Flush

                                for afor = 0 to UBOUND(fomrArr)

                                        'Response.Write "her2" & afor & "<br>"
                                        'Response.Flush

                                        if fomrArr(afor) <> 0 then

                                        strSQLfomri = "INSERT INTO fomr_rel "_
                                        &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                        &" VALUES ("& fomrArr(afor) &", "& varJobId &", 0, "& for_faktor &")"

                                        oConn.execute(strSQLfomri)

                                            '*** Sync aktiviteter ****'
                                            'if cint(syncForAkt) = 1 then
                                            'strSQLa = "SELECT id FROM aktiviteter WHERE job = "& varJobid
                                            'oRec3.open strSQLa, oConn, 3
                                            'while not oRec3.EOF 

                                            '    strSQLfomrai = "INSERT INTO fomr_rel "_
                                            '    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                            '    &" VALUES ("& fomrArr(afor) &", "& varJobId &","& oRec3("id") &", "& for_faktor &")"

                                            '    oConn.execute(strSQLfomrai)

                                            'oRec3.movenext
                                            'wend
                                            'oRec3.close

                                            'end if

                                        end if


                                next

                                '********************************'


                                '***** Sync Job slutDato ***********'
                                if strStatus = 0 then 'Kun ved luk job
                                syncslutdato = syncslutdato
                                else
                                syncslutdato = 0
                                end if

                                call syncJobSlutDato(varJobId, strjnr, syncslutdato)


                                '****** Sync. datoer p� aktiviteter *******'
                                if syncaktdatoer = 1 then
                                    
                                    '** Hvis sync jobslutdato er valgt ***
                                    if syncslutdato <> 1 then
                                    useSlutDato = jobslutdato
                                    else
                                    useSlutDato = useSyncDato 
                                    end if

                                strSQLaDatoer = "UPDATE aktiviteter SET aktstartdato = '"& jobstartdato &"', aktslutdato = '"& useSlutDato &"' WHERE job = "& varJobId
                                oConn.execute(strSQLaDatoer)

                                end if




                                '*************************************************************************'
                                '***** Adviser jobansvarlige ****************************
                                '*************************************************************************'
                                if len(trim(request.form("FM_adviser_jobans"))) <> 0 then
            	
		                    
                                    jobans1 = 0
                                    jobans2 = 0
                                    jobans3 = 0    
                                    jobans4 = 0
                                    jobans5 = 0

                                    jobans1email = ""
                                    jobans2email = ""
                                    jobans3email = ""
                                    jobans4email = ""
                                    jobans5email = ""
                                    jobstatus = 0
                                    lkDatoThis = "2002-01-01"                                    

				                    '**** Finder jobansvarlige *****
				                    strSQL = "SELECT job.id AS jid, jobnavn, jobnr, jobstatus, lukkedato, jobans1, jobans2, jobans3, jobans4, jobans5, job.beskrivelse, job_internbesk, "_
                                    &" m1.mnavn AS m1mnavn, m1.email AS m1email, m1.mansat AS m1mansat, "_
				                    &" m2.mnavn AS m2mnavn, m2.email AS m2email, m2.mansat AS m2mansat, "_
                                    &" m3.mnavn AS m3mnavn, m3.email AS m3email, m3.mansat AS m3mansat, "_
                                    &" m2.mnavn AS m4mnavn, m4.email AS m4email, m4.mansat AS m4mansat, "_    
                                    &" m2.mnavn AS m5mnavn, m5.email AS m5email, m5.mansat AS m5mansat, "_    
                                    &" kkundenavn, kkundenr, m2.init AS m2init, m1.init AS m1init, m3.init AS m3init, m4.init AS m4init, m5.init AS m5init FROM job "_
				                    &" LEFT JOIN medarbejdere AS m1 ON (m1.mid = jobans1)"_
				                    &" LEFT JOIN medarbejdere AS m2 ON (m2.mid = jobans2)"_
                                    &" LEFT JOIN medarbejdere AS m3 ON (m3.mid = jobans3)"_
                                    &" LEFT JOIN medarbejdere AS m4 ON (m4.mid = jobans4)"_
                                    &" LEFT JOIN medarbejdere AS m5 ON (m5.mid = jobans5)"_
                                    &" LEFT JOIN kunder AS k ON (k.kid = job.jobknr)"_
				                    &" WHERE job.id = "& varJobId

                                    'Response.Write strSQL
                                    'Response.end

				                    oRec5.open strSQL, oConn, 3
				                    x = 0
				                    if not oRec5.EOF then
            				
				                    jobid = oRec5("jid")

                                    jobstatus = oRec5("jobstatus")
                                    lkDatoThis = oRec5("lukkedato")
				            
                                    jobans1 = oRec5("m1mnavn")
                                    jobans1Init = oRec5("m1init")
                                    if isNull(oRec5("m1email")) <> true then 
                                    jobans1email = oRec5("m1email")
                                    else
                                    jobans1email = ""
                                    end if
                                    jobans1Mansat = oRec5("m1mansat")
                        
                                    jobans2 = oRec5("m2mnavn")
				                    jobans2Init = oRec5("m2init")
                                    if isNull(oRec5("m2email")) <> true then 
                                    jobans2email = oRec5("m2email")
                                    else
                                    jobans2email = ""
                                    end if
                                    jobans2Mansat = oRec5("m2mansat")
				            
                                    jobans3 = oRec5("m3mnavn")
				                    jobans3Init = oRec5("m3init")
                                    if isNull(oRec5("m3email")) <> true then 
                                    jobans3email = oRec5("m3email")
                                    else
                                    jobans3email = ""
                                    end if
                                    jobans3Mansat = oRec5("m3mansat")

                                    jobans4 = oRec5("m4mnavn")
				                    jobans4Init = oRec5("m4init")
                                    if isNull(oRec5("m4email")) <> true then 
                                    jobans4email = oRec5("m4email")
                                    else
                                    jobans4email = ""
                                    end if
                                    jobans4Mansat = oRec5("m4mansat")

                                    jobans5 = oRec5("m5mnavn")
				                    jobans5Init = oRec5("m5init")
                                    if isNull(oRec5("m5email")) <> true then 
                                    jobans5email = oRec5("m5email")
                                    else
                                    jobans5email = ""
                                    end if
                                    jobans5Mansat = oRec5("m5mansat")

                                    jobnavnThis = oRec5("jobnavn")
                                    intJobnr = oRec5("jobnr")
                                    strkkundenavn = oRec5("kkundenavn")

                                    strBesk = oRec5("beskrivelse")
                                    job_internbesk = oRec5("job_internbesk")

                                    'kkundenr = oRec("kkundenr")
            				
				            
				                    end if
				                    oRec5.close
            				
            				
				                    '*** Henter afsender **
				                    strSQL = "SELECT mnavn, email FROM medarbejdere"_
				                    &" WHERE mid = "& session("mid")
				                    oRec5.open strSQL, oConn, 3
            				
				                    if not oRec5.EOF then
            				
				                    afsNavn = oRec5("mnavn")
				                    afsEmail = oRec5("email")
            				
				                    end if
				                    oRec5.close
            				
            				
            					
            						
            					
				                    if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\jobs.asp" then
					  	                    'Response.Write "isNULL(jobans1) & isNULL(jobans2)" &  isNULL(jobans1) &" "& isNULL(jobans2)
                                            'Response.end


                                            for m = 1 to 5

                                            jobAnsThis = 0

                                            select case m
                                            case 1
                                            jobAnsThis = jobans1
                                            jobAnsThisEmail = jobans1email
                                            jobAnsAnsat = jobans1Mansat
                                            case 2
                                            jobAnsThis = jobans2
                                            jobAnsThisEmail = jobans2email
                                            jobAnsAnsat = jobans2Mansat
                                            case 3
                                            jobAnsThis = jobans3
                                            jobAnsThisEmail = jobans3email
                                            jobAnsAnsat = jobans3Mansat
                                            case 4
                                            jobAnsThis = jobans4
                                            jobAnsThisEmail = jobans4email
                                            jobAnsAnsat = jobans4Mansat
                                            case 5  
                                            jobAnsThis = jobans5
                                            jobAnsThisEmail = jobans5email
                                            jobAnsAnsat = jobans5Mansat
                                            end select


                                            if (jobAnsThis <> "0" AND isNULL(jobAnsThis) <> true) then
                                    
                                            Set myMail=CreateObject("CDO.Message")
                                            myMail.From="timeout_no_reply@outzource.dk" 'TimeOut Email Service 

                                        
						                    myMail.To= ""& jobAnsThis &"<"& jobAnsThisEmail &">"
                                   
                                    
						                    'Mailer.Subject = "Til de jobansvarlige p�: "& jobnavnThis &" ("& intJobnr &") | " & strkkundenavn  
                                            myMail.Subject = job_txt_020 &": "& jobnavnThis &" ("& intJobnr &") | " & strkkundenavn  


		                                    strBody = "<br>"
                                            strBody = strBody &"<b>"&job_txt_021&":</b> "& strkkundenavn & "<br>" 
						                    strBody = strBody &"<b>"&job_txt_022&":</b> "& jobnavnThis &" ("& intJobnr &")"
                                            if lto = "mpt" then
                                                strBody = strBody & "<br>"

                                                statusnavn = ""
                                                strlkDatoThis = ""
                                                select case cint(jobstatus)
									                case 1
									                strStatusNavn = job_txt_094
									                case 2
									                strStatusNavn = job_txt_095 'passiv
									                case 0
									                strStatusNavn = job_txt_096
                                                        if cdate(lkDatoThis) <> "01-01-2002" then
                                                        strlkDatoThis = " ("& formatdatetime(lkDatoThis, 2) & ")"
                                                        end if
									                case 3
									                strStatusNavn = job_txt_063
                                                    case 4
									                strStatusNavn = job_txt_098
                                                    case 5
									                strStatusNavn = "Evaluering"
									            end select

                                                strBody = strBody &"<b>"&job_txt_241&":</b> "& strStatusNavn & strlkDatoThis & "<br><br>"
                                            else
                                                strBody = strBody & "<br><br>"
                                            end if


                                            if jobans1 <> "0" AND isNULL(jobans1) <> true then
                                            strBody = strBody &"<b>"&job_txt_023&":</b> "& jobans1 &" "& jobans1init &"<br><br>"
                                            end if

                                            if jobans2 <> "0" AND isNULL(jobans2) <> true then
                                            strBody = strBody &"<b>"&job_txt_024&":</b> "& jobans2 &" ("& jobans2init &") <br><br>"
		                                    end if

                                            if len(trim(strBesk)) <> 0 then
                                            strBody = strBody &"<br><b>"&job_txt_025&":</b><br>"
                                            strBody = strBody & strBesk &"<br><br><br><br>"
                                            end if

                                            if len(trim(job_internbesk)) <> 0 then
                                            strBody = strBody &"<br><b>"&job_txt_026&":</b><br>"
                                            strBody = strBody & job_internbesk &"<br><br>"
                                            end if

                                    
                                            'strBody = strBody &"<br><br><br>https://outzource.dk/timeout_xp/wwwroot/ver2_10/login.asp?lto="&lto&"&tomobjid="&jobid&"&key="&strLicenskey

                                           'strBody = strBody &"<br><br><br>G� til TimeOut ved at <a href='https://outzource.dk/timeout_xp/wwwroot/ver2_10/login.asp?lto="&lto&"&tomobjid="&jobid&"'>klikke her..</a>"
                                           '&key="&strLicenskey&"
                                            if jobAnsAnsat = "1" then
                                                   if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                                   strBody = strBody &"<br><br><br>"&job_txt_027&":<br>https://outzource.dk/"&lto&"/default.asp?tomobjid="&jobid
                                                   else
                                                   strBody = strBody &"<br><br><br>"&job_txt_027&":<br>https://timeout.cloud/"&lto&"/default.asp?tomobjid="&jobid
                                                   end if
                                            end if
                                    
                                  


                                            strBody = strBody &"<br><br><br><br><br><br>"&job_txt_028&"<br><i>" 
		                                    strBody = strBody & session("user") & "</i><br><br>&nbsp;"



                                            select case lto 
                                            case "hestia"
            		                        strBody = strBody &"<br><br><br><br>_______________________________________________________________________________________________<br>"
                                            strBody = strBody &"HESTIA Ejendomme, Rosen�rnsgade 6, st., 8900 Randers C, Tlf. 70269010 - www.hestia.as<br><br>&nbsp;"
                                            end select
            		


                                            myMail.HTMLBody= "<html><head></head><body>" & strBody & "</body>"

                                            myMail.Configuration.Fields.Item _
                                            ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                            'Name or IP of remote SMTP server
                                    
                                                        if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                                           smtpServer = "webout.smtp.nu"
                                                        else
                                                           smtpServer = "formrelay.rackhosting.com" 
                                                        end if
                    
                                                        myMail.Configuration.Fields.Item _
                                                        ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                            'Server port
                                            myMail.Configuration.Fields.Item _
                                            ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                            myMail.Configuration.Fields.Update

                                            if isNull(jobAnsThisEmail) <> true then

                                            if len(trim(jobAnsThisEmail)) <> 0 then
                                            myMail.Send
                                            end if

                                            end if

                                            set myMail=nothing

            		
		                                    'Mailer.sendmail()
		                                    'Set Mailer = Nothing



                                            'Response.end

                                            end if
                              
                                            next
				            
                                    end if 'c drev





                            end if' adviser








                                '******* Opret folder **************'    
                                'Skal M�ske kun v�re hvis sl�et til i kontrol-panel 
                                
                                    'Tjekker om folderen findes i forvejen
                                    strSQLdok = "SELECT fo.navn FROM foldere fo  "_
                                    &" WHERE fo.jobid =" & varjobid & " AND fo.jobid <> 0" 

                                    folderfindes = 0
                                    oRec2.open strSQLdok, oConn, 3
                                    if not oRec2.EOF then
                                    folderfindes = 1
                                    end if
                                    oRec2.close
                                    
                                    
                                    'if request("FM_opret_folder") <> "" then 'Opretter hvis sl�et til i kontrol panel.
                                    if cint(folderfindes) = 0 then
                                    
                                        '*** Insert into db folder ***'
                                        sqldatoDt = year(now) &"/"& month(now) &"/"& day(now)
                                        strSQL = "INSERT INTO foldere (navn, kundeid, kundese, jobid, editor, dato) VALUES "_
		                                &" ('"& SQLBless2(jobnavn) &"_"& jobnr &"', "& kid &", 0, "& varJobId &", '"& session("user") &"', '"& sqldatoDt &"')"
		                                
		                                oConn.execute(strSQL)
                                    end if

                                    'end if

                                '********************************'






                                '** Rediger
                                'response.redirect("jobs.asp?func=red&id="&id)
                                    

                                '*** Liste
                                'response.redirect("../timereg/jobs.asp")

                                strSQLLastJoib = "SELECT id FROM job"
                                oRec.open strSQLLastJoib, oConn, 3
                                while not oRec.EOF
                                     
                                jobid = oRec("id")
                                    
                                oRec.movenext
                                wend
                                oRec.close


                                if func = "dbopr" then
                                    

                                    response.Redirect("jobs.asp?func=red&jobid="&jobid)

                                else


                                    if request("mp") = "1" then
                                        response.Redirect("jobs.asp?func=red&jobid="&jobid&"&mp=1")
                                    else
                                        if request("afslut_val") = "1" then 
                                            response.Redirect("jobs.asp?func=red&jobid="&varjobid)
                                        else
                                            response.Redirect("job_list.asp")
                                        end if
                                    end if
                                    

                                end if

                                 





    case "opret", "red"
        %>
        <script src="js/job_2018_jav5.js"></script>
        <style>

           /* table, tr, td, .tablecolor 
            {
                color:black;
                padding:0 15px 10px 0px;
            } */
    
    
    
            /* The Modal (background) */
            .modal {
                display: none; /* Hidden by default */
                position: fixed; /* Stay in place */
                z-index: 1000; /* Sit on top */
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
                width: 400px;
                height: 350px;
            }

            .picmodal:hover,
            .picmodal:focus {
            text-decoration: none;
            cursor: pointer;
            } 

        </style>

        <%
            if browstype_client <> "ip" then
                call menu_2014
            else
                %><!--#include file="../inc/regular/top_menu_mobile.asp"--><%
                call mobile_header
            end if
        %>

        <%


        if func = "red" then
       
            
            strSQL = "SELECT id, jobnavn, jobnr, jobknr, editor, dato, creator, createdate, kundekpers, jobstartdato, jobslutdato, jobans1, jobans2, jobstatus, tilbudsnr, beskrivelse, rekvnr, job_internbesk, syncslutdato, sandsynlighed, serviceaft, fastpris, lukkedato FROM job"_
            & " WHERE id = "& jobid

            'Response.Write strSQL
	        'Response.flush
	
	        oRec.open strSQL, oConn, 3
	
	        if not oRec.EOF then

	            strNavn = oRec("jobnavn")
	            strjobnr = oRec("jobnr")
	            'strKnavn = oRec("kkundenavn")
	            strKnr = oRec("jobknr")
                kundekpers = oRec("kundekpers")
                jobstdato = oRec("jobstartdato")
                jobsldato = oRec("jobslutdato")

                jobans1 = oRec("jobans1")
                jobans2 = oRec("jobans2")

                strStatus = oRec("jobstatus")

                strLastEditor = oRec("editor")
                strLastEditDate = oRec("dato")

                strCreator = oRec("creator")
                strCreatedate = oRec("createdate")

                if cdbl(oRec("tilbudsnr")) = 0 then	    
                    strtilbudsnr = oRec("tilbudsnr")

	                strSQLtb = "SELECT tilbudsnr FROM licens WHERE id = 1"
	                oRec3.open strSQLtb, oConn, 3
                    if not oRec3.EOF then

                    strNexttilbudsnr = oRec3("tilbudsnr") + 1
                    tbStyle = "color:#999999;"

                    end if
                    oRec3.close
                else
                    strtilbudsnr = oRec("tilbudsnr")
                    strNexttilbudsnr = strtilbudsnr
	                tbStyle = "color:#000000;"
                end if

                strbeskrivelse = oRec("beskrivelse")
                rekvnr = oRec("rekvnr")
                strinternbesk = oRec("job_internbesk")

                syncslutdato = oRec("syncslutdato")
                intSandsynlighed = oRec("sandsynlighed")
                intServiceaft = oRec("serviceaft")
                strFastpris = oRec("fastpris")

                lkdato = oRec("lukkedato")

        
            end if


            oRec.close

            dbfunc = "dbred"
            submitTxt = "Opdater"

        else 'func opr

            strNavn = ""
            strjobnr = 0
	        strKnavn = ""
	        strKnr = 0
            kundekpers = 0
            jobstdato = day(now) &"-"& month(now) &"-"& year(now)
            jobsldato = day(now) &"-"& month(now+1) &"-"& year(now)

            strNexttilbudsnr = 0
            strNexttilbudsnr = 0

            select case lto
	        case "intranet - local", "synergi1", "cisu", "wilke", "epi2017"
	        strFastpris = 1 'default fastpris
	        case else
	        strFastpris = 2 'default l�bende timer
	        end select

            fmkunde = 0

            dbfunc= "dbopr"
            submitTxt = "Opret"

            jobstdato = day(now) & "-" & month(now) & "-" & year(now)
            'jobstdato = dateadd("d", -7, jobstdato) 
            jobsldato = dateadd("m", 1, jobstdato)

            jobans1 = 0
            jobans2 = 0

            strbeskrivelse = ""
            rekvnr = ""
            strinternbesk = ""

            select case lto
            case "epi", "epi_no", "epi_sta", "intranet - local", "epi_ab", "epi_cati", "epi_uk", "epi2017"
	        virksomheds_proc = 50
	        syncslutdato = 0 '1
            intSandsynlighed = 10
            stade_tim_proc = 1
            case else
            virksomheds_proc = 0
	        syncslutdato = 0
            intSandsynlighed = 0
            stade_tim_proc = 0
            end select

            intServiceaft = 0

            lkdato = ""


        end if ' func red



        if level <= 2 OR level = 6 then
	    editok = 1
	    else
			if cint(session("mid")) = jobans1 OR cint(session("mid")) = jobans2 OR (cint(jobans1) = 0 AND cint(jobans2) = 0) then
			editok = 1
			end if
	    end if
%>







<style>
    .dropdownbtn:hover,
    .dropdownbtn:focus {
    text-decoration: none;
    cursor: pointer;
}
</style>



<%if browstype_client <> "ip" then %>
<div class="wrapper">
    <div class="content">
<%end if %>

        <%
        if lto = "mpt" OR lto = "xcare" OR lto = "intranet - local" OR request("mp") = 1 then
        accDisplay = "none"
        else
        accDisplay = "normal"
        end if
        %>
    
        <%
            if browstype_client = "ip" then
                constyle = "style='width:100%;'"
                
                %>
                <style>
                    .mobileOf {
                        display:none;
                    }
                </style>
                <%

            else
                constyle = ""
                sizecls = "input-small"
            end if
        %>

 <!------------------------------- Sideindhold------------------------------------->
 
        <div class="container" <%=constyle %>>

            
        <!--
        <div style="padding-left:75%; position:fixed; right:0; z-index:1100">
                <div style="width:250px;">
                <div class="panel-group accordion-panel" id="accordion-paneled1">

                    <div class="panel">

                    <div class="panel-heading">
                        <h4 class="panel-title">
                        <a class="accordion-toggle" data-toggle="collapse" href="#collapseOne">
                            Funktioner
                        </a>
                        </h4>
                    </div>

                    <div id="collapseOne" class="panel-collapse collapse">
                        <div class="panel-body">

                            <%
                                lnk = "../timereg/webblik_joblisten21.asp?nomenu=1&jobnr_sog="&fbjobnr&"&FM_kunde=32"&fbKnr
                            %>

                            <a href="#" onclick="Javascript:window.open('<%=lnk %>', '', 'width=1104,height=800,resizable=yes,scrollbars=yes')" class="alt">Gantt</a>
                        </div> 
                    </div> 

                    </div> </div>
                    </div>
            </div> -->


            <div class="portlet">
                <h3 class="portlet-title"><u>Job/Projekt</u></h3>


            <div class="portlet-body">


                <form id="opretproj" action="jobs.asp?func=<%=dbfunc %>&jobid=<%=jobid %>&mp=<%=request("mp") %>" method="post">
                <input type="hidden" name="" id="kundekpersopr" value="<%=kundekpers%>" />
                <input type="hidden" name="FM_OLDjobnr" value="<%=strJobnr%>">
                <input type="hidden" id="afslut_val" name="afslut_val" value="0" />
     	           
                <div class="row">
                    <div class="col-lg-12 pad-b5">
                        <table style="width:100%;">
                            <tr>
                                <td style="text-align:right;">
                                    <a id="opdaterafslut" class="btn btn-success btn-sm"><b><%=submitTxt %></b></a>
                                    <button type="submit" class="btn btn-success btn-sm"><b>Opdater og afslut</b></button>
                                </td>
                            </tr>
                        </table>
                    </div>
                </div>
     	           

                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse3">
                                Stamdata
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapse3" class="panel-collapse collapse in">
                            <div class="panel-body">
                            
                                
                                  <div class="row pad-b5">
                                  
                                      <div class="col-lg-1">Kunde: <span style="color:red">*</span></div>
                                      <div class="col-lg-5">

                                           <%
                                            if browstype_client = "ip" then
                                                widthsize = "100%"
                                            else
                                                widthsize = "95%"
                                            end if
                                        %>

                                          <select class="form-control <%=sizecls %>" name="FM_kunde" id="FM_kunde" style="display:inline-block; width:<%=widthsize%>;">
		                                   
		                                    <% if func = "red" then

				                                    strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' AND (useasfak = 0 OR useasfak = 1) ORDER BY Kkundenavn"
				                                    oRec.open strSQL, oConn, 3
				                                    while not oRec.EOF
				
				                                   
                                                    call kundeopl_options_2016

				                                    oRec.movenext
				                                    wend
				                                    oRec.close

                                            else

                                                  %><option value="0" disabled>Kunder (flest job seneste 12 md.):</option>


                                                <% 
                                                 tdatodd = now
                                                tdatodd = dateAdd("d", -365, tdatodd)
                                                tdatodd = year(tdatodd)&"/"&month(tdatodd)&"/"& day(tdatodd)

                                                strSQL = "SELECT Kkundenavn, Kkundenr, Kid, count(j.id) AS antaljob FROM kunder "_
                                                &" LEFT JOIN job AS j ON (j.jobknr = kid AND jobstartdato >= '"& tdatodd &"') "_
                                                &" WHERE ketype <> 'e' AND (useasfak = 0 OR useasfak = 1) GROUP BY kid ORDER BY antaljob DESC LIMIT 5"
			                                    
                                              
                                              'response.write strSQL
                                              'response.flush


                                                oRec.open strSQL, oConn, 3 
                                                while not oRec.EOF

                    
				
                                                    call kundeopl_options_2016
				
			                                    oRec.movenext
			                                    wend
			                                    oRec.close

                                                %><option>&nbsp;</option>
                                                <option value="0" disabled>Kunder (alfabetisk):</option><%
            
                                             

			                                    strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' AND (useasfak = 0 OR useasfak = 1) ORDER BY Kkundenavn"
			                                    oRec.open strSQL, oConn, 3
			                                    kans1 = ""
			                                    kans2 = ""
			                                    while not oRec.EOF
				
                                                    call kundeopl_options_2016
				
			                                    oRec.movenext
			                                    wend
			                                    oRec.close
		

				
		
			                                   



				                            end if   %>
		                                </select>
                                        <%if browstype_client <> "ip" then %>
                                          &nbsp <a style="font-size:125%;" href="../to_2015/kunder.asp?func=opret&ketype=k&hidemenu=1&showdiv=onthefly&rdir=1" target="_blank" class=vmenu><b>+</b></a>
                                        <%end if %>
                                      </div>
                              
                                 
                            
                                    <div class="col-lg-2">Kontaktpers. (deres ref.):
                                        
                                    </div>
                                    <div class="col-lg-3">
                                        
                                        <%
                                            if browstype_client = "ip" then
                                                widthsize = "100%"
                                            else
                                                widthsize = "90%"
                                            end if
                                        %>

		                                <select name="FM_kpers" id="FM_kpers" class="form-control <%=sizecls %>" style="display:inline-block; width:<%=widthsize%>;">
		                                    <option value="0">Ingen</option>
	
		                                   
		                                </select>
                                        <%if browstype_client <> "ip" then %>
                                        &nbsp <a style="cursor:pointer; font-size:125%;" onclick="window.open('kontaktpers.asp?id=0&kid=<%=strKnr%>&func=opr&OnTheFly=1','popUpWindow','height=700,width=1200,left=100,top=100,resizable=yes,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no, status=yes');"><b>+</b></a>
                                        <%end if %>
                                    </div>
                   
                            
                                  
                                </div>

                                <div class="row pad-b5">                                                                      
                                        <div class="col-lg-1">Navn: <span style="color:red">*</span></div>
                                        <div class="col-lg-5"><input type="text" name="FM_jobnavn" id="FM_jobnavn" value="<%=strNavn%>" class="form-control <%=sizecls %>" placeholder="Projekt Navn"></div>

                                        <%if func = "red" then %>
                                            <%if browstype_client <> "ip" then %>
                                            <div class="col-lg-2">Job nr:</div>
                                            <div class="col-lg-3"><input class="form-control <%=sizecls %>" type="text" name="FM_jobnr" value="<%=strjobnr %>" /></div>
                                            <%else %>
                                                <div class="col-lg-2"><span style="font-size:9px;">Jobnr: <%=strjobnr %></span></div>   
                                                <input type="hidden" name="FM_jobnr" value="<%=strjobnr %>" />
                                            <%end if %>
                                        <%else %>
                                        <input type="hidden" name="FM_jobnr" value="0" />
                                        <%end if %>
                                </div>

                                <div class="row pad-b5">
                                        
                                    <% 
                                    if dbfunc = "dbred" then
                                        lkDatoThis = ""
                                        select case strStatus
									    case 1
									    strStatusNavn = job_txt_094
									    case 2
									    strStatusNavn = job_txt_095 'passiv
									    case 0
									    strStatusNavn = job_txt_096
                                            if cdate(lkdato) <> "01-01-2002" then
                                            lkDatoThis = " ("& formatdatetime(lkdato, 2) & ")"
                                            end if
									    case 3
									    strStatusNavn = job_txt_063
                                        case 4
									    strStatusNavn = job_txt_098
                                        case 5
									    strStatusNavn = "Evaluering"
									    end select
                                    end if
                                    %>

                                    <div class="col-lg-1">Status:</div>
                                    <div class="col-lg-2">
                                        <%if (((lto = "mpt" AND strStatus <> 0 AND strStatus <> 2 AND dbfunc = "dbred") OR level = 1) OR dbfunc <> "dbred") OR lto <> "mpt" then %>
                                        <select class="form-control <%=sizecls %>" id="FM_status" name="FM_status">
									            <%
                                                
                                                if dbfunc = "dbred" then 
									                
									            %>
									            <option value="<%=strStatus%>" SELECTED><%=strStatusNavn%> <%=lkDatoThis %></option>
									            <%end if
                                    
                                                'call jobstatus_fn(0, 0, 1)
                                                %>
                                                <%'jobstatus_fn_options %>

									            <%'if (lto = "mpt" AND (cint(strStatus) = 1 OR cint(strStatus) = 4) OR level = 1) OR lto <> "mpt" OR func = "opret" then %>

									                <option value="1"><%=jobstatus_txt_007 %></option>
									                <option value="2"><%=jobstatus_txt_008 %></option>
                                            
                                                    <%if (lto = "mpt" AND level = 1) OR lto <> "mpt" then %>
									                <option value="0"><%=jobstatus_txt_009 %></option>
                                                    <%end if %>

                                                    <option value="4"><%=jobstatus_txt_004 %></option>
									
									                 <option value="3"><%=jobstatus_txt_003 %></option>
                                                     <option value="5"><%=jobstatus_txt_010 %></option>

                                                <%'end if %>    
                                            
                                        </select>
                                        <%else %>
                                        <input type="hidden" name="FM_status" value="<%=strStatus %>" />
                                        <%=strStatusNavn %>
                                        <%end if %>
                                    </div>

                                    <div class="col-lg-3">
                                        <%  if cint(strStatus) = 3 OR (func = "opret" AND cint(tilbud_mandatoryOn) = 1) then
					                        chkusetb = "CHECKED"
                                            tilbudvisibility = "inherit"
					                        else
					                        chkusetb = ""
                                            tilbudvisibility = "hidden"
					                        end if
					                    %>
					                    <input type="checkbox" id="FM_usetilbudsnr" name="FM_usetilbudsnr" value="j" <%=chkusetb%> style="visibility:<%=tilbudvisibility%>; display:none">

                                        <%if func = "red" then
                        
                                         if strNexttilbudsnr <> strtilbudsnr then 
                                         tlbplcholderTxt = strNexttilbudsnr

                                         %>
                                     <!--   &nbsp;<span class="tilbudsinfo"  style="color:#999999; visibility:<%=tilbudvisibility%>;">(<%=job_txt_065 %>: <%=tlbplcholderTxt %>)</span> -->
                                        <%

                                        else
                                        tlbplcholderTxt = ""
                                        end if %>

					                    <input type="text" id="FM_tnr" class="tilbudsinfo form-control input-small" name="FM_tnr" value="<%=strtilbudsnr%>" style="visibility:<%=tilbudvisibility%>; width:25%; display:inline-block;"> 
                  
					                    <%else%>
					                    <input type="hidden" name="FM_tnr" value="0">
					                    <%end if%>
				
					                    <input type="hidden" id="FM_nexttnr" value="<%=strNexttilbudsnr %>">

                                        <%select case lto
                                        case "epi2017"
                                            sandBdr = "border:1px red solid;"
                                        case else
                                            sandBdr = ""
                                        end select %>

                                       <!-- <br /><span style="font-size:10px; font-family:arial; color:#999999;">(<%=job_txt_067 &" " %>=<%=" "& job_txt_068 &" " %>*<%=" "& job_txt_069 %>)</span> -->

                                        &nbsp<input id="Text1" name="FM_sandsynlighed" class="tilbudsinfo form-control input-small" value="<%=intSandsynlighed %>" type="text" style="width:20%; <%=sandBdr%>; display:inline-block; visibility:<%=tilbudvisibility%>" />
                                        <span class="tilbudsinfo" style="display:inline-block; visibility:<%=tilbudvisibility%>">% Sandsynlighed </span>

                                     </div>

                                  <!--  <div class="col-lg-5 tilbudsinfo" style="visibility:<%=tilbudvisibility%>;">
                                        <input id="Text1" name="FM_sandsynlighed" class="form-control input-small" value="<%=intSandsynlighed %>" type="text" style="width:30px; <%=sandBdr%>; display:inline-block;" />
                                        <span style="display:inline-block;"><%="% "& job_txt_066 %></span>
                                    </div> -->


                                    <div class="col-lg-2 mobileOf">Projekttype:</div>

                                    <div class="col-lg-4 mobileOf">
                                        <select class="form-control input-small" style="display:inline-block; width:255px;">
                                            <option value="1">1</option>
                                            <option value="2">2</option>
                                            <option value="3">3</option>
                                            <option value="4">4</option>
                                        </select>


                                        <%
                                            if strFastpris = "1" then
		                                    varFastpris1 = "checked"      
                                            fastprisvalue = 1
		                                    'varFastpris2 = ""
		                                    else
		                                    varFastpris1 = ""
                                            fastprisvalue = 0
		                                    'varFastpris2 = "checked"
		                                    end if
                                        %>

                                        <%if editok = 1 then %>
                                        <span style="display:inline-block;">&nbsp<input type="checkbox" name="FM_fastpris" value="1" <%=varFastpris1 %> /> Fastpris</span>
                                        <%else %>
                                        <input type="hidden" name="FM_fastpris" value="<%=fastprisvalue %>" />
                                        <%end if %>

                                    </div>


                                </div>

                                
                                <div class="row pad-b5">
                                    <div class="col-lg-1">
                                        <%'if func = "red" then %>    
                                        <span id="modal_ans" class="picmodal"><u>Ansvarlige:</u></span>
                                        <%'else %>
                                       <!-- Jobansvarlig:<br />
                                        <span style="font-size:8px; white-space:nowrap;">Tilf�j flere ved rediger</span> -->
                                        <%'end if %>
                                    </div>
                                    <div class="col-lg-2">
                                        <select class="form-control <%=sizecls %>" name="FM_jobans_1" id="FM_jobans">
						                    <option value="0">Ingen</option>

                                            <%

                                               
                                                mTypeExceptSQL = ""         
                                               

                                                select case lto
                                                case "hestia", "intranet - local" 
							                    mPassivSQL = " OR mansat = 3"
                                                case else
                                                mPassivSQL = ""
                                                end select

                                                jobansField = "jobans1, jobans_proc_1"

                                                if func <> "red" then
                                                    strSQL = "SELECT mnavn, mnr, mid, init, mansat FROM medarbejdere WHERE ((mansat = 1 "& mPassivSQL &")" & mTypeExceptSQL & ")"
                                                else
                                                    strSQL = "SELECT mnavn, mnr, mid, mansat, "& jobansField &", init FROM medarbejdere "_
							                        &" LEFT JOIN job ON (job.id = "& jobid &") WHERE ((mansat = 1 "& mPassivSQL &") "& mTypeExceptSQL &") OR (mid = jobans1)"
                                                end if

                                                strSQL = strSQL & " ORDER BY mnavn"

                                                oRec.open strSQL, oConn, 3 
							                    while not oRec.EOF 


                                                
                                                if func <> "red" then

							                        if ja = 1 then
							                        usemed = session("mid")
                                                    jobans_proc = 100
							                        else
							                        usemed = 0
                                                    jobans_proc = 0
							                        end if
                            
							                    else

							                        'select case ja
						                            'case 1
						                            usemed = oRec("jobans1")
                                                    jobans_proc = oRec("jobans_proc_1")
						                            'case 2
						                            'usemed = oRec("jobans2")
                                                    'jobans_proc = oRec("jobans_proc_2")
						                            'case 3
						                            'usemed = oRec("jobans3")
                                                    'jobans_proc = oRec("jobans_proc_3")
						                            'case 4
						                            'usemed = oRec("jobans4")
                                                    'jobans_proc = oRec("jobans_proc_4")
						                            'case 5
						                            'usemed = oRec("jobans5")
                                                    'jobans_proc = oRec("jobans_proc_5")
						                            'end select

							                    end if 'func <> red


                                                if cint(usemed) = oRec("mid") then
								                medsel = "SELECTED"
								                else
								                medsel = ""
								                end if

                                                if len(trim(oRec("init"))) <> 0 then
                                                opTxt = oRec("init") &" - "& oRec("mnavn")
                                                else
                                                opTxt = oRec("mnavn") '&" (" & oRec("mnr") &")"
                                                end if

                                                if oRec("mansat") <> 1 then
                                                select case oRec("mansat")
                                                case 2
                                                opTxt = opTxt & " -" & job_txt_319
                                                case 3 
                                                opTxt = opTxt & " - " & job_txt_320
                                                end select
                                                end if

                                                %>
							                        <option value="<%=oRec("mid")%>" <%=medsel%>><%=opTxt %></option>
							                    <%


                                                oRec.movenext
                                                wend
                                                oRec.close
                                                
                                            %>
						                </select>                                     
                                    </div>
                                    <div class="col-lg-3 mobileOf"><input type="text" id="FM_jobans_proc_1" style="display:inline-block; width:25%" name="FM_jobans_proc_1" value="<%=formatnumber(jobans_proc, 1) %>" class="form-control input-small" />
                                        <span style="display:inline-block; width:20%;">&nbsp %</span>


                                        <%if func <> "red" then

                                            select case lto
                                            case "synergi1", "qwert", "hestia", "mpt"
                                            advJobansCHK = "CHECKED"
                                            case "dencker"

                                                if session("mid") = 34 OR session("mid") = 49 then 'Aministrationen 
                                                advJobansCHK = "CHECKED"
                                                else
                                                advJobansCHK = ""
                                                end if
                            
                                            case else
                                            advJobansCHK = ""
                                            end select
                        
                                        else
                            
                                            select case lto
                                            case "hestia"
                                            advJobansCHK = "CHECKED"
                                            case else
                                            advJobansCHK = ""
                                            end select

                                        end if %>

                                        <%
                                        '**** Tjekker adviser checkbox for dem der vil have det permenet ****'
                                        if lto = "mpt" then
                                            adCHB = "CHECKED"
                                        else
                                            adCHB = ""
                                        end if
                                        %>

                                        &nbsp <input <%=adCHB %> type="checkbox" name="FM_adviser_jobans" /> Adviser asvarlige
                                       <!-- <span id="modal_ans" style="color:cornflowerblue; display:inline-block; margin-top:5px; width:25%; text-align:center;" class="fa fa-plus picmodal"></span> -->
                                    </div>

                                    <!--  Valgfrie felter hentes -->
                                    <%call jobsidefelter %>

                                    <%if cint(job_felt_salgsans) = 1 then%>
                                    <div class="col-lg-2"><span id="modal_salgsans" class="picmodal"><u>Salgsansvarlige:</u></span></div>
                                    <div class="col-lg-2">
                                        <select name="FM_salgsans_1" class="form-control input-small">
                                            <option value="0"><%=job_txt_129 %></option>
                                            <%

                                                saansTxt = job_txt_322 & " 1:" 'job_txt_321
						                        salgsansField = "salgsans1, salgsans1_proc"

                                                if func <> "red" then
							                    strSQL = "SELECT mnavn, mnr, mid, init, mansat FROM medarbejdere WHERE (mansat = 1 " & mTypeExceptSQL &")"
							                    else
							                    strSQL = "SELECT mnavn, mnr, mid, mansat, "& salgsansField &", init FROM medarbejdere "_
							                    &" LEFT JOIN job ON (job.id = "& jobid &") WHERE (mansat = 1 "& mTypeExceptSQL &") OR (mid = salgsans1)"
							                    end if

                                                strSQL = strSQL & " ORDER BY mnavn"
                                                'response.Write "<br>" & strSQL
							                                
                                                oRec.open strSQL, oConn, 3 
							                    while not oRec.EOF 

                                                if func <> "red" then							  
							                    usemed = session("mid")
                                                salgsans_proc = 0
							                    else
						                        usemed = oRec("salgsans1")
                                                salgsans_proc = oRec("salgsans1_proc")						                            
							                    end if

                                                if cdbl(usemed) = oRec("mid") then
								                medsel = "SELECTED"
								                else
								                medsel = ""
								                end if

						                        if len(trim(oRec("init"))) <> 0 then
                                                opTxt = oRec("init") &" - "& oRec("mnavn")
                                                else
                                                opTxt = oRec("mnavn") '&" (" & oRec("mnr") &")"
                                                end if

                                                if oRec("mansat") <> 1 then
                                                select case oRec("mansat")
                                                case 2
                                                opTxt = opTxt & " -" & job_txt_319
                                                case 3 
                                                opTxt = opTxt & " - " & job_txt_320
                                                end select
                                                end if
                                            %>

                                            <option value="<%=oRec("mid")%>" <%=medsel%>><%=opTxt %> // <%=usemed %></option>

                                            <%
							                    oRec.movenext
							                    wend
							                    oRec.close 
							                %>
                                        </select>
                                    </div>

                                    <div class="col-lg-2"><input id="FM_salgsans_proc_1" style="display:inline-block; width:40%;" name="FM_salgsans_proc_1" value="<%=formatnumber(salgsans_proc, 1) %>" type="text" <%=sltuDatoCol%>;" <%=fltSaProcDis %> class="form-control input-small" />
                                        <span style="display:inline-block;">&nbsp %</span>
                                      <!--  <span id="modal_salgsans" style="color:cornflowerblue; display:inline-block; margin-top:5px; width:20%; text-align:center;" class="fa fa-plus picmodal"></span> -->
                                    </div>
                                    <%end if %>

                                </div>

                                <div id="myModal_ans" class="modal">
                                    <!-- Modal content -->
                                    <div class="modal-content">
                                     
                                         <%
                                            call salgsans_fn()

                                            select case lto 
                                            case "epi", "epi_ab", "epi_no", "epi_sta", "intranet - local", "epi_uk"
                                            mTypeExceptSQL = " AND (medarbejdertype <> 14 AND medarbejdertype <> 24)"
                                            case else
                                            mTypeExceptSQL = ""         
                                            end select
                                        %>

                                        <div class="row">
                                            <div class="col-lg-12">

                                                <table style="width:100%;">
                                                    
						                            <%
                                                    for ja = 2 to 5 
						                            select case ja
						                            case 1

						                                'jbansImg = "<img src='../ill/ac0019-24.gif' width='24' height='24' alt='Jobansvarlig' border='0'>"
                                                        if cint(showSalgsAnv) = 1 then 
						                                jbansTxt = job_txt_230 & ":"
                                                        else
                                                        jbansTxt = job_txt_023 & ":"
                                                        end if
						                                jobansField = "jobans1, jobans_proc_1"


                                                            select case lto
                                                            case "intranet - local", "epi2017"
                                                                fltDis = ""
                                                                fltProcDis = ""
                                                            case else 
                                                                fltDis = ""
                                                                fltProcDis = ""
                                                            end select

						                            case 2
						                            'jbansImg = "<img src='../ill/ac0020-24.gif' width='24' height='24' alt='Jobejer' border='0'>"
						                            jbansTxt = job_txt_024 & ":"
						                            jobansField = "jobans2, jobans_proc_2"


                                                            select case lto
                                                            case "intranet - local", "epi2017"
                                                                fltDis = ""
                                    
                                                                if func = "red" then 'AND cDate(jobstdato) < cDate("01-01-2018")
                                                                fltProcDis = ""
                                                                else
                                                                fltProcDis = "Disabled"
                                                                end if

                                                            case else 
                                                                fltDis = ""
                                                                fltProcDis = ""
                                                            end select

						                            case 3
						                            'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
                                                    select case lto
                                                    case "intranet - local", "epi2017"
						                            jbansTxt = job_txt_128 & ":"
                                                    case else 
                                                    jbansTxt = job_txt_128&" 1:"
                                                    end select
						                            jobansField = "jobans3, jobans_proc_3"

                                                            select case lto
                                                            case "intranet - local", "epi2017"
                                                                fltDis = ""
                                                                fltProcDis = ""
                                                            case else 
                                                                fltDis = ""
                                                                fltProcDis = ""
                                                            end select

						                            case 4
						                            'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						                             select case lto
                                                    case "intranet - local", "epi2017"
                                                          if func = "red" AND cDate(jobstdato) < cDate("02-01-2018") then
                                                          jbansTxt = "Co. resp. 2:"
                                                          else
                                                          jbansTxt = ""
                                                          end if
                                                    case else
						                            jbansTxt = job_txt_128&" 2:"
                                                    end select
						                            jobansField = "jobans4, jobans_proc_4"

                                                             select case lto
                                                            case "intranet - local", "epi2017"
                                    
                                                                if func = "red" AND cDate(jobstdato) < cDate("02-01-2018") then
                                                                    fltDis = ""
                                                                    fltProcDis = ""
                                                                else
                                                                    fltDis = "Disabled"
                                                                    fltProcDis = "Disabled"
                                                                end if

                                                            case else 
                                                                fltDis = ""
                                                                fltProcDis = ""
                                                            end select

						                            case 5
						                            'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
                                                    select case lto
                                                    case "intranet - local", "epi2017"
                                                         if func = "red" AND cDate(jobstdato) < cDate("02-01-2018") then
                                                          jbansTxt = "Co. resp. 3:"
                                                          else
                                                          jbansTxt = ""
                                                          end if
                                                    case else
						                            jbansTxt = job_txt_128&" 3:"
                                                    end select
						                            jobansField = "jobans5, jobans_proc_5"


                                                             select case lto
                                                            case "intranet - local", "epi2017"
                                                                if func = "red" AND cDate(jobstdato) < cDate("02-01-2018") then
                                                                    fltDis = ""
                                                                    fltProcDis = ""
                                                                else
                                                                    fltDis = "Disabled"
                                                                    fltProcDis = "Disabled"
                                                                end if
                                                            case else 
                                                                fltDis = ""
                                                                fltProcDis = ""
                                                            end select

						                            end select
						
						                            %>
						                            <tr>
						                            <!--<td><%=jbansImg  %></td>-->

						                            <td style="width:100px; padding-top:5px;"><b><%=jbansTxt %></b></td>
						                            
                                                    <td style="padding-top:5px;"><select name="FM_jobans_<%=ja %>" id="FM_jobans_<%=ja %>" class="form-control input-small" <%=fltDis %>>
						                           <option value="0"><%=job_txt_129 %></option>
							                            <%

                                                        select case lto
                                                        case "hestia", "intranet - local" 
							                            mPassivSQL = " OR mansat = 3"
                                                        case else
                                                        mPassivSQL = ""
                                                        end select

							                            if func <> "red" then
							                            strSQL = "SELECT mnavn, mnr, mid, init, mansat FROM medarbejdere WHERE ((mansat = 1 "& mPassivSQL &")" & mTypeExceptSQL & ")"
							                            else
							                            strSQL = "SELECT mnavn, mnr, mid, mansat, "& jobansField &", init FROM medarbejdere "_
							                            &" LEFT JOIN job ON (job.id = "& jobid &") WHERE ((mansat = 1 "& mPassivSQL &") "& mTypeExceptSQL &") OR (mid = jobans"& ja &")"
							                            end if

                                                        strSQL = strSQL & " ORDER BY mnavn"
                                                            response.Write strSQL
							                            oRec.open strSQL, oConn, 3 
							                            while not oRec.EOF 
							
							                            if func <> "red" then
							 
                                                          select case ja 
                                                          case 1 
							                              usemed = session("mid")
                                                          jobans_proc = 100
                                                          case 2
                                    
                                                                select case lto
                                                                case "intranet - local", "epi2017", "mpt"
                                                                usemed = jobans2 
                                                                jobans_proc = 0

                                                                if lto = "mpt" AND usemed = 0 then
                                                                    usemed = 4
                                                                end if

                                                                case else
                                                                usemed = 0
                                                                jobans_proc = 0
                                                                end select
							  
                                                          case else
							                              usemed = 0
                                                          jobans_proc = 0
							                              end select

                              


							                            else
							                                select case ja
						                                    case 1
						                                    usemed = oRec("jobans1")
                                                            jobans_proc = oRec("jobans_proc_1")
						                                    case 2
						                                    usemed = oRec("jobans2")
                                                            jobans_proc = oRec("jobans_proc_2")
						                                    case 3
						                                    usemed = oRec("jobans3")
                                                            jobans_proc = oRec("jobans_proc_3")
						                                    case 4
						                                    usemed = oRec("jobans4")
                                                            jobans_proc = oRec("jobans_proc_4")
						                                    case 5
						                                    usemed = oRec("jobans5")
                                                            jobans_proc = oRec("jobans_proc_5")
						                                    end select
							                            end if
							
								                            if cdbl(usemed) = oRec("mid") then
								                            medsel = "SELECTED"
								                            else
								                            medsel = ""
								                            end if

                                                            if len(trim(oRec("init"))) <> 0 then
                                                            opTxt = oRec("init") &" - "& oRec("mnavn")
                                                            else
                                                            opTxt = oRec("mnavn") '&" (" & oRec("mnr") &")"
                                                            end if

                                                            if oRec("mansat") <> 1 then
                                                            select case oRec("mansat")
                                                            case 2
                                                            opTxt = opTxt & " -" & job_txt_319
                                                            case 3 
                                                            opTxt = opTxt & " - " & job_txt_320
                                                            end select
                                                            end if

                                                        %>
							                            <option value="<%=oRec("mid")%>" <%=medsel%>><%=opTxt %></option>
							                            <%

							                            oRec.movenext
							                            wend
							                            oRec.close 
							                            %>
						                           </select></td>
						                                <td style="white-space:nowrap; padding-top:5px; padding-left:5px;"><input id="FM_jobans_proc_<%=ja %>" name="FM_jobans_proc_<%=ja %>" value="<%=formatnumber(jobans_proc, 1) %>" type="text" style="<%=sltuDatoCol%>;" <%=fltProcDis %> class="form-control input-small" /></td>
                                                        <td style="padding-top:5px; padding-left:5px;"> %</td>
						                            </tr>
                            
						                            <%next %>
						                            
						
						                            </table>

                                            </div>
                                        </div>

                                        <br /><br />
                                        <!-- Adviser knap -->
                                     <!--   <div class="row">
                                            <div class="col-lg-12" style="text-align:center"><span class="btn btn-default btn-sm"><b>Adviser valgte medarbejder</b></span></div>
                                        </div> -->

                                    </div>
                                </div>

                                <div id="myModal_salgsans" class="modal">
                                        <div class="modal-content">
                                            <div class="row">
                                                <div class="col-lg-12">
                                                    <table style="width:100%;">
												
						                                <%
                                                        for sa = 2 to 5 
					
						                                select case sa
						                                case 1
						                                'jbansImg = "<img src='../ill/ac0019-24.gif' width='24' height='24' alt='Jobansvarlig' border='0'>"
						                                saansTxt = job_txt_322 & " 1:" 'job_txt_321
						                                salgsansField = "salgsans1, salgsans1_proc"
						                                case 2
						                                'jbansImg = "<img src='../ill/ac0020-24.gif' width='24' height='24' alt='Jobejer' border='0'>"
						
						                                salgsansField = "salgsans2, salgsans2_proc"

                                                                select case lto
                                                                case "xintranet - local", "xepi2017"
                                                                    if func = "red" AND cDate(jobstdato) < cDate("02-01-2018") then
                                                                        fltSaDis = ""
                                                                        fltSaProcDis = ""
                                                                        saansTxt = job_txt_322 & " 2:"
                                                                    else
                                                                        fltSaDis = "Disabled"
                                                                        fltSaProcDis = "Disabled"
                                                                        saansTxt = ""
                                                                    end if
                                                                case else 
                                                                    fltSaDis = ""
                                                                    fltSaProcDis = ""
                                                                    saansTxt = job_txt_322 & " 2:"
                                                                end select

						                                case 3
						                                'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						 
						                                salgsansField = "salgsans3, salgsans3_proc"

                                                             select case lto
                                                                case "xintranet - local", "xepi2017"
                                                                    if func = "red" AND cDate(jobstdato) < cDate("02-01-2018") then
                                                                        fltSaDis = ""
                                                                        fltSaProcDis = ""
                                                                        saansTxt = job_txt_322 & " 3:"
                                                                    else
                                                                        fltSaDis = "Disabled"
                                                                        fltSaProcDis = "Disabled"
                                                                        saansTxt = ""
                                                                    end if
                                                                case else 
                                                                    fltSaDis = ""
                                                                    fltSaProcDis = ""
                                                                    saansTxt = job_txt_322 & " 3:"
                                                                end select

						                                case 4
						                                'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
					
						                                salgsansField = "salgsans4, salgsans4_proc"


                                                             select case lto
                                                                case "xintranet - local", "xepi2017"
                                                                    if func = "red" AND cDate(jobstdato) < cDate("02-01-2018") then
                                                                        fltSaDis = ""
                                                                        fltSaProcDis = ""
                                                                        saansTxt = job_txt_322 & " 4:"
                                                                    else
                                                                        fltSaDis = "Disabled"
                                                                        fltSaProcDis = "Disabled"
                                                                        saansTxt = ""
                                                                    end if
                                                                case else 
                                                                    fltSaDis = ""
                                                                    fltSaProcDis = ""
                                                                    saansTxt = job_txt_322 & " 4:"
                                                                end select

						                                case 5
						                                'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						
						                                salgsansField = "salgsans5, salgsans5_proc"


                                                             select case lto
                                                                case "xintranet - local", "xepi2017"
                                                                    if func = "red" AND cDate(jobstdato) < cDate("02-01-2018") then
                                                                        fltSaDis = ""
                                                                        fltSaProcDis = ""
                                                                        saansTxt = job_txt_322 & " 5:"
                                                                    else
                                                                        fltSaDis = "Disabled"
                                                                        fltSaProcDis = "Disabled"
                                                                        saansTxt = ""
                                                                    end if
                                                                case else 
                                                                    fltSaDis = ""
                                                                    fltSaProcDis = ""
                                                                    saansTxt = job_txt_322 & " 5:"
                                                                end select

						                                end select
						
						                                %>
						                                <tr>
						                                    <!--<td><%=jbansImg  %></td>-->
						                                    <td  style="width:100px; padding-top:5px;"><b><%=saansTxt%></b></td>

						                                    <td style="padding-top:5px;">
						                                        <select name="FM_salgsans_<%=sa %>" id="Select4" <%=fltSaDis %> class="form-control input-small">
						                                        <option value="0"><%=job_txt_129 %></option>
							                                        <%
							
							                                        if func <> "red" then
							                                        strSQL = "SELECT mnavn, mnr, mid, init, mansat FROM medarbejdere WHERE (mansat = 1 " & mTypeExceptSQL &")"
							                                        else
							                                        strSQL = "SELECT mnavn, mnr, mid, mansat, "& salgsansField &", init FROM medarbejdere "_
							                                        &" LEFT JOIN job ON (job.id = "& jobid &") WHERE (mansat = 1 "& mTypeExceptSQL &") OR (mid = salgsans"& sa &")"
							                                        end if

                                                                    strSQL = strSQL & " ORDER BY mnavn"
                                                                    'response.Write "<br>" & strSQL
							
                                
                                                                    oRec.open strSQL, oConn, 3 
							                                        while not oRec.EOF 
							
							                                        if func <> "red" then
							  
                                                                      if sa = 1 then
							                                          usemed = session("mid")
                                                                      salgsans_proc = 0
							                                          else
							                                          usemed = 0
                                                                      salgsans_proc = 0
							                                          end if

                              

							                                        else
							                                                    select case sa
						                                                        case 1
						                                                        usemed = oRec("salgsans1")
                                                                                salgsans_proc = oRec("salgsans1_proc")
						                                                        case 2
						                                                        usemed = oRec("salgsans2")
                                                                                salgsans_proc = oRec("salgsans2_proc")
						                                                        case 3
						                                                        usemed = oRec("salgsans3")
                                                                                salgsans_proc = oRec("salgsans3_proc")
						                                                        case 4
						                                                        usemed = oRec("salgsans4")
                                                                                salgsans_proc = oRec("salgsans4_proc")
						                                                        case 5
						                                                        usemed = oRec("salgsans5")
                                                                                salgsans_proc = oRec("salgsans5_proc")
						                                                        end select
							                                        end if
							
								                                        if cdbl(usemed) = oRec("mid") then
								                                        medsel = "SELECTED"
								                                        else
								                                        medsel = ""
								                                        end if


						                                                if len(trim(oRec("init"))) <> 0 then
                                                                        opTxt = oRec("init") &" - "& oRec("mnavn")
                                                                        else
                                                                        opTxt = oRec("mnavn") '&" (" & oRec("mnr") &")"
                                                                        end if

                                                                        if oRec("mansat") <> 1 then
                                                                        select case oRec("mansat")
                                                                        case 2
                                                                        opTxt = opTxt & " -" & job_txt_319
                                                                        case 3 
                                                                        opTxt = opTxt & " - " & job_txt_320
                                                                        end select
                                                                        end if

                                                                    %>
							                                       <option value="<%=oRec("mid")%>" <%=medsel%>><%=opTxt %> // <%=usemed %></option>
							                                        <%
							                                        oRec.movenext
							                                        wend
							                                        oRec.close 
							                                        %>
						                                        </select>
						                                    </td>
                                                            <td style="white-space:nowrap; padding-top:5px; padding-left:5px;"><input id="FM_salgsans_proc_<%=sa %>" name="FM_salgsans_proc_<%=sa %>" value="<%=formatnumber(salgsans_proc, 1) %>" type="text" style="width:40px; <%=sltuDatoCol%>;" <%=fltSaProcDis %> class="form-control input-small" /></td>
                                                            <td style="padding-top:5px; padding-left:5px;">%</td>
						                                </tr>
                            
						                                <%next %>												
						                            </table>
                                                </div>
                                            </div>
                                            </div>
                                            </div>

                               

                                

                                <div class="row pad-b5">   

                                        <div class="col-lg-1">Start dato:</div>
                                        <div class="col-lg-2">

                                                    <div class='input-group date' id='datepicker_stdato'>
                                                    <input type="text" class="form-control <%=sizecls %>" name="FM_jobstartdato" value="<%=jobstdato %>" placeholder="dd-mm-yyyy" />
                                                          <span class="input-group-addon input-small">
                                                                    <span class="fa fa-calendar">
                                                                    </span>
                                                                </span>
                                                           </div>

                                         </div>

                                       

                                        

                                        <div class="col-lg-1">Slut dato:</div>
                                        <div class="col-lg-2">

                                                    <div class='input-group date' id='datepicker_stdato'>
                                                    <input type="text" class="form-control <%=sizecls %>" name="FM_jobslutdato" value="<%=jobsldato %>" placeholder="dd-mm-yyyy" />
                                                          <span class="input-group-addon input-small">
                                                                    <span class="fa fa-calendar">
                                                                    </span>
                                                                </span>
                                                           </div>

                                         </div>
                                       <!-- <div class="col-lg-2"><input type="checkbox" name="FM_datouendelig" value="j" <%=chald%>> <%=job_txt_115 %></div> -->


                                    <!-- Forretningsomr�der -->
                                    <%if cint(job_felt_fomr) = 1 then %>
                                    <div class="col-lg-2"><span id="modal_fomr" class="picmodal"><u>Forretningsomr�der</u></span> &nbsp <!--<span id="modal_fomr" style="color:cornflowerblue;" class="fa fa-plus picmodal"></span>--></div>
                                    <div class="col-lg-3">
                                        <%
                                            if func = "red" then

                                                strSQL = "SELECT f.navn as fomrnavn FROM fomr_rel LEFT JOIN fomr as f ON for_fomr = f.id WHERE for_jobid = "& jobid &" AND for_aktid = 0"

                                                oRec.open strSQL, oConn, 3
                                                antalfomr = 0
                                                strFomrNavne = ""
                                                while not oRec.EOF

                                                    if antalfomr = 0 then
                                                    strFomrNavne = oRec("fomrnavn")
                                                    else
                                                    strFomrNavne = strFomrNavne & ", " & oRec("fomrnavn") 
                                                    end if
                                                    antalfomr = antalfomr + 1

                                                oRec.movenext
                                                wend
                                                oRec.close

                                                response.Write "<i>"&strFomrNavne&"</i>"

                                             end if

                                        %>
                                    </div>

                                    <div id="myModal_fomr" class="modal">
                                        <div class="modal-content">
                                            <%
                                                if func = "red" then
                                    
                                                strFomr_rel = ""
                                                strFomr_navn = ""
                    
                                                strSQLfrel = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
                                                &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_jobid = "& jobid & " AND for_aktid = 0 GROUP BY for_fomr"

                                                'Response.Write strSQLfrel
                                                'Response.flush
                                                f = 0
                                                oRec.open strSQLfrel, oConn, 3
                                                while not oRec.EOF

                                                if f = 0 then
                                                strFomr_navn = ""
                                                end if

                                                strFomr_rel = strFomr_rel & "#"& oRec("for_fomr") &"#"
                                                strFomr_navn = strFomr_navn & oRec("navn") & ", " 

                                                f = f + 1
                                                oRec.movenext
                                                wend
                                                oRec.close

                                                if f <> 0 then
                                                len_strFomr_navn = len(strFomr_navn)
                                                left_strFomr_navn = left(strFomr_navn, len_strFomr_navn - 2)
                                                strFomr_navn = left_strFomr_navn & ")"

                                                    if len(strFomr_navn) > 50 then
                                                    strFomr_navn = left(strFomr_navn, 50) & "..)"
                                                    end if
                                                end if

                                            else

                                            '**** forvalgt forretningsomr�der ****' 
                                            select case lto 
                                            case "hestia", "xintranet - local"
                    
                        
                                                        strSQLfrel = "SELECT id, navn FROM fomr WHERE id = 2"    
                           
                        
                                                        f = 0
                                                        oRec.open strSQLfrel, oConn, 3
                                                        while not oRec.EOF

                                                        'if f = 0 then
                                                        'strFomr_navn = " ("
                                                        'end if

                                                        strFomr_rel = strFomr_rel & "#"& oRec("id") &"#"
                                                        strFomr_navn = strFomr_navn & oRec("navn") & ", " 

                                                        f = f + 1
                                                        oRec.movenext
                                                        wend
                                                        oRec.close     

                                            case else
                                                 strFomr_navn = ""
                                                 strFomr_rel = ""
                                            end select


                                            end if

                                            %>                                       
                                           
                                            <%
                                            'strSQLf = "SELECT id, navn FROM fomr WHERE id <> 0 ORDER BY navn"
                                            strSQLf = "SELECT f.navn AS fnavn, f.id, f.konto, kp.kontonr AS kkontonr, kp.navn AS kontonavn, fomr_segment FROM fomr AS f "_
                                            &" LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) WHERE f.id <> 0 AND f.jobok = 1 ORDER BY f.navn"


                                            'if session("mid") = 1 then
                                            'response.write strSQLf
                                            'response.flush
                                            'end if

                                            %>

                                            <select name="FM_fomr" id="FM_fomr" multiple="multiple" size="12" class="form-control <%=sizecls %>">
                                                <option value="0"><%=job_txt_148 %></option>

                                                <%
                                                    fa = 0
                                                    strchkbox = ""
                                                    oRec.open strSQLf, oConn, 3
                                                    while not oRec.EOF

                                                        'if func = "opret" AND step = 2 then '*** Opret (forvalgt)
                                                        if func = "opret" then
                                                            'if instr(oRec("fomr_segment"), "#"& thisKtype_segment &"#") <> 0 then
                                                            'fSel = "SELECTED"
                                                            'else
                                                            'fSel = ""
                                                            'end if


                                                        else '** Rediger Forretningsomr�der
                                    
                                                            if instr(strFomr_rel, "#"&oRec("id")&"#") <> 0 then
                                                            fSel = "SELECTED"
                                                            else
                                                            fSel = ""
                                                            end if

                                                        end if



                                                    if oRec("konto") <> 0 then
                                                    kontonrVal = " ("& left(oRec("kontonavn"), 10) &" "& oRec("kkontonr") &")"

                                                    if cint(fomr_konto) = cint(oRec("id")) then
                                                    fokontoCHK = "CHECKED"
                                                    else
                                                    fokontoCHK = ""
                                                    end if

                                                    strchkbox = strchkbox & "<input type='radio' class='FM_fomr_konto' id='FM_fomr_konto_"& oRec("id") &"' name='FM_fomr_konto' value="& oRec("id") &" "& fokontoCHK &"> " & left(oRec("fnavn"), 20) &" "& kontonrVal &"<br>"
                                                    else
                                                    kontonrVal = ""
                                                    end if 
                                    
                                                    %>
                                                    <option value="<%=oRec("id")%>" <%=fSel %>><%=oRec("fnavn") &" "& kontonrVal %></option>
                                                    <%
                                                        fa = fa + 1

                                    
                                                    oRec.movenext
                                                    wend
                                                    oRec.close
                                                %>

                                            </select>
                                        </div>
                                    </div>

                                    <%end if %>

                                </div>
                                <%if cint(job_felt_rekno) = 1 then %>
                                <div class="row pad-b5">                                  
                                    <div class="col-lg-1">Rek. Nr.:</div>
                                    <div class="col-lg-2"><input type="text" name="FM_rekvnr" id="FM_rekvnr" value="<%=rekvnr%>" class="form-control <%=sizecls %>"></div>                                   
                                </div>
                                <%end if %>

                                <%
                                strbeskrivelseLngt = len(strbeskrivelse)
                                %>

                                <div class="row">
                                    <div class="col-lg-6">Beskrivelse: (<p style="display:inline-block; font-size:9px;" id="beskrivLNG"><%=strbeskrivelseLngt %>/1000</p>) <br />
                                        <textarea id="FM_beskrivelse" name="FM_beskrivelse" rows="5" class="form-control <%=sizecls %>"><%=strbeskrivelse %></textarea>
                                    </div>

                                    <%if cint(job_felt_intnote) = 1 then %>
                                    <div class="col-lg-6">Intern note: <br />
                                        <textarea id="TextArea2" name="FM_internnote" rows="5" class="form-control <%=sizecls %>"><%=strinternbesk %></textarea>
                                    </div>
                                    <%end if %>

                                </div>                                                        
                                
                                
                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->
                  </div>
                

                  <!-- Aktiviteter -->

                  <style>
                     .akttable
                      {
                          display:none;
                      }

                     /* table.dataTable tbody th,
                        table.dataTable tbody td {
                        white-space: nowrap;
                        } */

                  </style>
                 
                <%
                if func = "red" then
                    readonlystuff = "readonly"
                else
                    readonlystuff = ""
                end if
                %>

                 <div class="panel-group accordion-panel" id="accordion-paneled" style="display:<%=accDisplay%>">
                    <div class="panel panel-default">                        
                        <div class="panel-heading">
                        <h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapseAkt">Aktiviteter</a></h4></div> <!-- /.panel-heading -->
                        <div id="collapseAkt" class="panel-collapse collapse in">
                            <div class="panel-body">
                                
                                <%if lto <> "care" AND lto <> "outz" then %>
                                <div class="row">
                                    <div class="col-lg-6">
                                        <a id="aktView1_btn" class="btn btn-default btn-sm active"><b>Oversigt</b></a>
                                        <a id="aktView2_btn" class="btn btn-default btn-sm"><b>KPI �konomi</b></a>
                                        <a id="aktView3_btn" class="btn btn-default btn-sm"><b>KPI Tid</b></a>
                                    </div>
                                </div>
                                <%end if %>

                                <div id="updated_activities"></div>

                                <table id="activtyTable" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                    <thead>
                                        <tr>
                                            <th class="akttable akttable_navn" style="text-align:left;">Aktivitet</th>
                                            <th class="akttable akttable_status" id="aktheader_status">Status</th>
                                            <th class="akttable akttable_type" id="aktheader_type" >Type</th>
                                            <th class="akttable akttable_budget" id="aktheader_budget">Budget</th>
                                            <th class="akttable akttable_enhed" id="aktheader_enhed">Enhed</th>
                                            <th class="akttable akttable_stdato" id="aktheader_stdato">Start Dato</th>
                                            <th class="akttable akttable_sldato" id="aktheader_sldato">Slut Dato</th>
                                            <th class="akttable akttable_medarb" id="aktheader_medarb" style="width:130px;">Medarbejder / grp.</th>
                                            <th class="akttable akttable_alloker" id="aktheader_alloker">Allok�r</th>

                                            <th class="akttable akttable_prisenh" id="aktheader_prisenh">Pris/enh</th>
                                            <th class="akttable akttable_ialt" id="aktheader_ialt">ialt</th>
                                            <th class="akttable akttable_fctid" id="aktheader_fctid">FC tid</th>
                                            <th class="akttable akttable_fcdkk" id="aktheader_fcdkk">FC dkk</th>
                                            <th class="akttable akttable_realtid" id="aktheader_realtid">Real tid</th>
                                            <th class="akttable akttable_realdkk" id="aktheader_realdkk">Real dkk</th>

                                            <th class="akttable akttable_fc" id="aktheader_fc">FC</th>
                                            <th class="akttable akttable_real" id="aktheader_real">Real</th>
                                            <th class="akttable akttable_budgetfc" id="aktheader_budgetfc">Budget vs FC</th>
                                            <th class="akttable akttable_budgetreal" id="aktheader_budgetreal">Budget vs Real</th>
                                            <th class="akttable akttable_fcreal" id="aktheader_fcreal">FC vs Real</th>
                                            <th class="akttable akttable_faktid" id="aktheader_faktid">Fak tid</th>

                                            <th class="akttable akttable_actions"></th>



                                          <!--  <th >Status</th>
                                            <th >Budget</th>
                                            <th >Enhed</th>
                                            <th >Pris/Enh</th>
                                            <th >I alt</th>
                                            <th >FC (tid)</th>
                                            <th >FC (dkk)</th>
                                            <th >Real (tid)</th>
                                            <th >Real (dkk)</th>

                                            <th >Budget</th>
                                            <th >Enhed</th>
                                            <th >Pris/Enh</th>
                                            <th >FC</th>
                                            <th >Real</th>
                                            <th >Budget vs FC</th>
                                            <th >Budget vs Real</th>
                                            <th >FC vs Real</th>
                                            <th >Fak tid</th> -->

                                        </tr>
                                    </thead>

                                    <tbody>

                                        <%
                                        if func <> "opret" then
                                        '********* Henter aktiviteter *********'
                                        strSQLakt = "SELECT a.id as aktid, a.navn, aktstatus, aktstartdato, aktslutdato, projektgruppe1, aktbudget, antalstk, bgr, budgettimer, fakturerbar, p.navn as prgnavn FROM aktiviteter as a LEFT JOIN projektgrupper as p ON (p.id = projektgruppe1) WHERE job = "& jobid
                                        oRec.open strSQLakt, oConn, 3
                                        while not oRec.EOF

                                        strFakturerbart = oRec("fakturerbar")

                                        select case oRec("aktstatus")
                                            case 1
                                                status1SEL = "SELECTED"
	                                            status2SEL = ""
                                                status3SEL = ""
	                                        case 2
	                                            status1SEL = ""
	                                            status2SEL = "SELECTED"
                                                status3SEL = ""
	                                        case else
	                                            status1SEL = ""
	                                            status2SEL = ""
                                                status3SEL = "SELECTED"
                                        end select

                                        prisialt = cdbl(oRec("antalstk")) * cdbl(oRec("budgettimer"))

                                        budgettimer = oRec("budgettimer")


                                        startDateHTML = day(oRec("aktstartdato")) &"-"& month(oRec("aktstartdato")) &"-"& year(oRec("aktstartdato")) 
                                        endDateHTML = day(oRec("aktslutdato")) &"-"& month(oRec("aktslutdato")) &"-"& year(oRec("aktslutdato")) 

                                        if level < 2 Or level = 6 then
                                        disabledStatus = ""
                                        else
                                        disabledStatus = "disabled"
                                        end if
                                        %>

                                        <tr>
                                            <td class="akttable akttable_navn">
                                                <input <%=readonlystuff %> type="text" class="form-control input-small" id="FM_akt_navn_<%=oRec("aktid") %>" name="FM_akt_navn_<%=oRec("aktid") %>" value="<%=oRec("navn") %>" <%=disabledStatus%> />
                                            </td>

                                            <td class="akttable akttable_status">
                                                <select <%=readonlystuff %> class="form-control input-small" id="aktstatus_<%=oRec("aktid") %>" name="aktstatus_<%=oRec("aktid") %>" <%=disabledStatus%>>
                                                    <option value="1" <%=status1SEL %>><%=job_txt_094 %></option>
                                                    <option value="2" <%=status2SEL %>><%=job_txt_320 %></option>
                                                    <option value="3" <%=status3SEL %>><%=job_txt_246 %></option>
                                                </select>
                                            </td>

                                            <td class="akttable akttable_type" style="width:75px;">
                                                <%call akttyper2009(1) %>
				                                <select <%=readonlystuff %> id="FM_fakturerbart_<%=oRec("aktid") %>" class="form-control input-small" name="FM_fakturerbart_<%=oRec("aktid") %>" <%=disabledStatus%>>
                                                   <%
                                                   Response.Write aty_options
                                                   %>                
                                                </select>
                                            </td>

                                            <td class="akttable akttable_budget"><input <%=readonlystuff %> id="aktbudget_<%=oRec("aktid") %>" name="aktbudget_<%=oRec("aktid") %>" class="form-control input-small" value="<%=formatnumber(oRec("budgettimer"),2) %>" <%=disabledStatus%> /></td>

                                            <td class="akttable akttable_enhed">
                                                <%
                                                bgrSEL0 = "SELECTED"
                                                bgrSEL1 = ""
                                                bgrSEL2 = ""
                                                select case oRec("bgr")
                                                case 0
	                                                bgrSEL0 = "SELECTED"
	                                            case 1
	                                                bgrSEL1 = "SELECTED" 
	                                            case 2
	                                                bgrSEL2 = "SELECTED"
                                                end select
                                                %>

                                                <select <%=readonlystuff %> id="aktEnhed_<%=oRec("aktid") %>" name="aktEnhed_<%=oRec("aktid") %>" class="bgr form-control input-small" <%=disabledStatus%>>
                                                    <option value=0 <%=bgrSEL0 %>>Ingen</option>
                                                    <option value=1 <%=bgrSEL1 %>>Timer</option>
                                                    <option value=2 <%=bgrSEL2 %>>Stk.</option>
                                                </select>

                                            </td>

                                            <td class="akttable akttable_stdato" style="width:100px;">
                                                <div id="aktSTdate_<%=oRec("aktid") %>" class='input-group'>
                                                  <input <%=readonlystuff %> type="text" style="width:100%;" class="form-control input-small" id="aktInputSTdate_<%=oRec("aktid") %>" name="aktInputSTdate_<%=oRec("aktid") %>" value="<%=startDateHTML %>" placeholder="dd-mm-yyyy" <%=disabledStatus%> />
                                                    <span class="input-group-addon input-small">
                                                    <span class="fa fa-calendar"></span></span>
                                                </div>
                                            </td>

                                            <td class="akttable akttable_sldato" style="width:100px;">
                                                <div id="aktSLdate_<%=oRec("aktid") %>" class='input-group'>
                                                  <input <%=readonlystuff %> type="text" style="width:100%;" class="form-control input-small" id="aktInputSLdate_<%=oRec("aktid") %>" name="aktInputSLdate_<%=oRec("aktid") %>" value="<%=endDateHTML %>" placeholder="dd-mm-yyyy" <%=disabledStatus%> />
                                                    <span class="input-group-addon input-small">
                                                    <span class="fa fa-calendar"></span></span>
                                                </div>
                                            </td>

                                            <td class="akttable akttable_medarb">
                                                <select <%=readonlystuff %> id="FM_akt_prg_<%=oRec("aktid") %>" name="FM_akt_prg_<%=oRec("aktid") %>" class="form-control input-small" <%=disabledStatus%>>
                                                    <%
                                                        strSQL = "SELECT id, navn FROM projektgrupper"
                                                        oRec2.open strSQL, oConn, 3
                                                        while not oRec2.EOF

                                                            
                                                            if cint(oRec("projektgruppe1")) = cint(oRec2("id")) then
                                                            prgselected = "SELECTED"
                                                            else
                                                            prgselected = ""
                                                            end if
                                                            
                                                            response.Write "<option value='"&oRec2("id")&"' "&prgselected&">"& oRec2("navn") &"</option>"

                                                        oRec2.movenext
                                                        wend
                                                        oRec2.close
                                                    %>
                                                </select>
                                            </td>

                                            <td class="akttable akttable_alloker" style="text-align:center;">
                                                <%
                                                fsLink = "timbudgetsim.asp?FM_fy="&year(now)&"&FM_visrealprdato=1-1-"&year(now)&"&FM_sog="& strjobnr '& "&jobid="& jobid &"&func=forecast"
                                                %>
                                                <a href="<%=fsLink %>" target="_blank"><span class="fa fa-file-text"></span></a>
                                            </td>
                                            

                                            <td class="akttable akttable_prisenh"><input <%=readonlystuff %> type="text" id="FM_akt_prisenhed_<%=oRec("aktid") %>" name="FM_akt_prisenhed_<%=oRec("aktid") %>" class="form-control input-small" value="<%=formatnumber(oRec("aktbudget"),2) %>" <%=disabledStatus%> /></td>

                                            <td class="akttable akttable_ialt"><%=formatnumber(prisialt,2) %></td>

                                            <td class="akttable akttable_fctid">
                                                <%
                                                    timerFC = 0
                                                    strSQL2 = "SELECT sum(timer) as timerFC from ressourcer_md WHERE aktid = "& oRec("aktid") 
                                                    'response.Write strSQL2
                                                    oRec2.open strSQL2, oConn, 3
                                                    if not oRec2.EOF then
                                                        timerFC = oRec2("timerFC")
                                                    end if
                                                    oRec2.close

                                                    if len(trim(timerFC)) <> 0 then
                                                    'response.Write formatnumber(timerFC, 2)
                                                    timerFC = timerFC
                                                    else
                                                    'response.Write "-"
                                                    timerFC = 0
                                                    end if

                                                    'response.Write "<br> FC " & totTimerFC 
                                                %>
                                                <%=formatnumber(timerFC, 2) %>
                                            </td>

                                            <td class="akttable akttable_fcdkk">
                                                <%
                                                    forecastDKK = 0
                                                    forecastTotDKK = 0
                                                    strSQL2 = "SELECT sum(timer) as forecastetTimer, medid FROM ressourcer_md WHERE jobid = "& jobid &" AND aktid = "& oRec("aktid") & " GROUP BY medid"
                                                    'response.Write strSQL2
                                                    oRec2.open strSQL2, oConn, 3
                                                    while not oRec2.EOF
                                                    
                                                        
                                                        medarbTimepris = 0
                                                        brugAktTimepris = 1

                                                        strSQL3 = "SELECT 6timepris FROM timepriser WHERE jobid ="& jobid &" AND aktid ="& oRec("aktid")
                                                        oRec3.open strSQL3, oConn, 3
                                                        if not oRec3.EOF then
                                                            brugAktTimepris = 0
                                                            medarbTimepris = oRec3("6timepris")
                                                        end if
                                                        oRec3.close

                                                        if brugAktTimepris = 1 then
                                                        medarbTimepris = oRec("aktbudget")                                                        
                                                        else
                                                        medarbTimepris = medarbTimepris
                                                        end if

                                                        'response.Write "<br> Medid " & oRec2("medid") & " - timerfore " & oRec2("forecastetTimer") & " medarbTimepris " & medarbTimepris

                                                        forecastDKK = (oRec2("forecastetTimer") * medarbTimepris) 

                                                        forecastTotDKK = forecastTotDKK + forecastDKK

                                                        
                                                        
                                                    
                                                    oRec2.movenext
                                                    wend
                                                    oRec2.close
                                                    
                                                    response.Write formatnumber(forecastTotDKK, 2)
                                                    
                                                %>
                                            </td>

                                            <td class="akttable akttable_realtid">
                                                <%
                                                    timerReal = 0
                                                    strSQL2 = "SELECT sum(timer) as timerReal FROM timer WHERE tjobnr = '"& strjobnr &"' AND TAktivitetId = "& oRec("aktid")
                                                    'response.Write strSQL2
                                                    oRec2.open strSQL2, oConn, 3
                                                    if not oRec2.EOF then
                                                    timerReal = oRec2("timerReal")
                                                    end if
                                                    oRec2.close

                                                    if len(trim(timerReal)) <> 0 then
                                                    timerReal = timerReal
                                                    else
                                                    timerReal = 0
                                                    end if
                                                %>
                                                <%=formatnumber(timerReal, 2) & " t." %>
                                            </td>

                                            <td class="akttable akttable_realdkk">
                                                <%
                                                    realPrisTot = 0
                                                    medarbRealPris = 0
                                                    strlSQL2 = "SELECT sum(timer) as realTimer, tmnr FROM timer WHERE tjobnr = '"& strjobnr &"' AND TAktivitetId = "& oRec("aktid") &" GROUP BY tmnr"
                                                    'response.Write strlSQL2
                                                    oRec2.open strlSQL2, oConn, 3
                                                    while not oRec2.EOF

                                                        response.Write oRec2("realTimer")

                                                        brugMedarbTimepris = 0
                                                        medarbTimepris = 0
                                                        strSQL3 = "SELECT 6timepris FROM timepriser WHERE jobid = "& jobid &" AND aktid = "& oRec("aktid") &" AND medarbid = "& oRec2("tmnr")
                                                        'response.Write strSQL3
                                                        oRec3.open strSQL3, oConn, 3
                                                        while not oRec3.EOF
                                                        
                                                            brugMedarbTimepris = 1
                                                            medarbTimepris = oRec3("6timepris") 
                                                            
                                                        oRec3.movenext
                                                        wend
                                                        oRec3.close

                                                        if brugMedarbTimepris = 1 then
                                                            medarbTimepris = medarbTimepris
                                                        else
                                                            medarbTimepris = oRec("aktbudget")
                                                        end if

                                                        medarbRealPris = oRec2("realTimer") * medarbTimepris
                                                        realPrisTot = realPrisTot + medarbRealPris


                                                    oRec2.movenext
                                                    wend
                                                    oRec2.close

                                                    response.Write formatnumber(realPrisTot, 2)
                                                %>
                                            </td>

                                            <td class="akttable akttable_fc"><%=formatnumber(timerFC, 2) %></td>
                                            <td class="akttable akttable_real"><%=formatnumber(timerReal, 2) %></td>

                                            <td class="akttable akttable_budgetfc">
                                                <%
                                                    forhold_budget_fc = budgettimer - timerFC
                                                    forhold_budget_fc = budgettimer - timerFC
                                                    if forhold_budget_fc >= 0 then
                                                    spnColor = "green"
                                                    else
                                                    spnColor = "darkred"
                                                    end if
                                                %>
                                                <span style="color:<%=spnColor%>;"><%=forhold_budget_fc %></span>
                                            </td>

                                            <td class="akttable akttable_budgetreal">
                                                <%
                                                    forhold_budget_real = budgettimer - timerReal
                                                    if forhold_budget_real >= 0 then
                                                    spnColor = "green"
                                                    else
                                                    spnColor = "darkred"
                                                    end if
                                                %>
                                                <span style="color:<%=spnColor%>;"><%=forhold_budget_real %></span>
                                            </td>

                                            <td class="akttable akttable_fcreal">
                                                <%
                                                    forhold_fc_real = timerFC - timerReal
                                                    if forhold_fc_real >= 0 then
                                                    spnColor = "green"
                                                    else
                                                    spnColor = "darkred"
                                                    end if
                                                %>
                                                <span style="color:<%=spnColor%>;"><%=forhold_fc_real %></span>
                                            </td>
                                            <td class="akttable akttable_faktid">

                                            </td>


                                            <td class="akttable_actions" style="width:20px;">
                                                <a class="edit_akt_btn" id="<%=oRec("aktid") %>"><span class="fa fa-pencil pull-left"></span></a>
                                                <a href="#"><span style="color:darkred;" class="fa fa-times pull-right"></span></a>
                                            </td>

                                        </tr>

                                        <%
                                        oRec.movenext
                                        wend
                                        oRec.close

                                        else

                                        %>
                                        <tr>
                                            <td colspan="100" style="text-align:center;">Opret Job f�r aktiviteter</td>
                                        </tr>
                                        <%

                                        end if 'opret
                                        %>

                                    </tbody>

                                </table>


                                <%if func = "red" then %>
                                <div class="row newActivtyBtn_div akttable">
                                    <div class="col-lg-12" style="text-align:center;"><a id="newActivtyBtn" class="btn btn-default btn-sm"><b>+</b></a></div>
                                </div>

                                <div class="row" style="display:none;">
                                    <div class="col-lg-2">
                                        <select id="prgSEL" class="form-control input-small">
                                            <%
                                                strSQL = "SELECT id, navn FROM projektgrupper"
                                                oRec.open strSQL, oConn, 3
                                                while not oRec.EOF

                                                    response.Write "<option value='"&oRec("id")&"'>"& oRec("navn") &"</option>"

                                                oRec.movenext
                                                wend
                                                oRec.close
                                            %>
                                        </select>
                                    </div>

                                    <div class="col-lg-2">
                                        <%call akttyper2009(1) %>
				                        <select id="akttypeSEL" class="form-control input-small">
                                            <%
                                            Response.Write aty_options
                                            %>                
                                        </select>
                                    </div>
                                </div>
                                <%end if %>

                            </div>
                        </div>
                    </div>
                 </div>

                
                <!-- materialer -->
               <%if lto = "care" or lto = "outz" then
                    accDisplay = "none;"
               end if %>

              <div class="panel-group accordion-panel" id="accordion-paneled" style="display:<%=accDisplay%>">
                    <div class="panel panel-default">                        
                        <div class="panel-heading">
                        <h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapseMaterialer">Materialer</a></h4></div> <!-- /.panel-heading -->
                        <div id="collapseMaterialer" class="panel-collapse collapse in">
                            <div class="panel-body">
                                <div class="row">
                                    <div class="col-lg-12">
                                        <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                            <thead>
                                                <tr>
                                                    <th>Materiale</th>
                                                    <th>Antal</th>
                                                    <th>Salgspris</th>
                                                    <th>K�bspris</th>
                                                    <th>Faktor</th>
                                                    <th>Total pris</th>
                                                    <th>Dato</th>
                                                    <th>Budget</th>
                                                    <th>Real</th>
                                                    <th>Rest</th>
                                                </tr>
                                            </thead>
                                        </table>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>


                <!-- Avanceret -->
              <div class="panel-group accordion-panel" id="accordion-paneled" style="display:<%=accDisplay%>">
                    <div class="panel panel-default">                        
                        <div class="panel-heading">
                        <h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapseavanceret">Avanceret</a></h4></div> <!-- /.panel-heading -->
                        <div id="collapseavanceret" class="panel-collapse collapse in">
                            <div class="panel-body">


                                <div class="row pad-b5">
                                    <div class="col-lg-2"><%=job_txt_071 %> <a href="#" onClick="serviceaft('0', '<%=strKnr%>', '', '0')" class=vmenu><span style="color:cornflowerblue;" class="fa fa-plus"></span></a></div>
                                    <div class="col-lg-3">
                                        <%
		                                strSQL2 = "SELECT id, enheder, stdato, sldato, status, navn, pris, perafg, "_
		                                &" advitype, advihvor, erfornyet, varenr, aftalenr FROM serviceaft "_
		                                &" WHERE kundeid = "& strKnr &" OR id = "& intServiceaft &" ORDER BY id DESC" 'AND status = 1		
		                                'Response.write strSQL2
		                                %>
                                        <select name="FM_serviceaft" id="FM_serviceaft" class="form-control input-small">
		                                    <option value="0"><%=job_txt_073 %></option>
		
		                                    <%
		
		                                    oRec2.open strSQL2, oConn, 3 
		                                    while not oRec2.EOF 
		
		                                    if oRec2("advitype") <> 0 then
		                                    udldato = "&nbsp;&nbsp;&nbsp; "&job_txt_318&": " & formatdatetime(oRec2("stdato"), 2) 
		                                    else
		                                    udldato = ""
		                                    end if%>
		
		                                    <%
		                                    if cint(intServiceaft) = cint(oRec2("id")) then 
		                                    serChk = "SELECTED"
		                                    else
		                                    serChk = ""
		                                    end if
		
		                                    if oRec2("status") <> 0 then
		                                    stThis = "Aktiv"
		                                    else
		                                    stThis = "Lukket"
		                                    end if%>
		                                    <option value="<%=oRec2("id")%>" <%=serChk%>> <%=oRec2("navn")%>&nbsp;(<%=oRec2("aftalenr") %>)  <%=udldato%> (<%=stThis%>) </option>
		                                    <%
		                                    oRec2.movenext
		                                    wend
		                                    oRec2.close
		                                    %>
		                                </select>
                                    </div>
                                    <%if func = "red" then %>
                                    <div class="col-lg-6">
                                        <input type="checkbox" name="FM_overforGamleTimereg" value="1"> <%=job_txt_074 %> 
		                                <b><%=job_txt_075 %></b><%=" "&job_txt_076&" " %><b><%=job_txt_077 %></b><%=" "&job_txt_078 %> 
                                    </div>
                                    <%end if %>
                                </div>



                                <%if lto <> "epi2017" then %>
                                <div class="row">
                                   <%if syncslutdato = 1 then
                                    syncslutdatoCHK = "CHECKED"
                                    else
                                    syncslutdatoCHK = ""
                                    end if %>

                                    <div class="col-lg-12"><input type="checkbox" name="FM_syncslutdato" value="j" <%=syncslutdatoCHK %>> <%=job_txt_119 %> </div>

                                </div>
                                <%end if %>

                                <div class="row">
                                    <div class="col-lg-12"><input type="checkbox" name="FM_syncaktdatoer" value="j"> <%=job_txt_120 %></div>
                                </div>

                            </div>
                        </div>
                    </div>
                 </div>

                </form>

                <%if (lto = "mpt" OR lto = "intranet - local" OR lto = "xcare" OR request("mp") = 1) AND func = "red" then %>
                    <!-- Multi paneler -->
                    <script type="text/javascript" src="https://cdn.datatables.net/1.10.19/js/jquery.dataTables.min.js"></script>
                    <script src="js/multipanel_matindlaes2.js"></script>
                    <div class="row">                     
                        <div class="col-lg-6">
                            <div class="portlet portlet-boxed">                                
                                <div class="portlet-body">
                                    <h4>Chat</h4>

                                    <form action="jobs.asp?func=sendchatmessage" method="post">       
                                        <input type="hidden" name="jobid" value="<%=jobid %>" />
                                        <%call ChatModul(jobid) %>
                                    </form>

                                </div>
                            </div>
                        </div>

                        <%if browstype_client <> "ip" then %>
                        <div class="col-lg-6">
                            <div class="portlet portlet-boxed">                                
                                <div class="portlet-body"> 
                                    <h4>Timeregistrering &nbsp <a href="../timereg/joblog.asp?FM_job=<%=jobid %>" target="_blank"><span style="font-size:80%" class="fa fa-a fa-external-link"></span></a></h4>
                                    
                                    <%if (cint(strStatus) = 1 OR cint(strStatus) = 4) OR level = 1 then %>
                                     <%call TimeregPanel(jobid, 0) %>
                                    <%end if %>

                                    <br />
                                     <%call JoblogSimple(jobid) %>
                                </div>
                            </div>
                        </div>
                        <%end if %>
                    </div>

                    
                    <div class="row">
                        <div class="col-lg-6">
                            <div class="portlet portlet-boxed">
                                <div class="portlet-body">
                                    <h4>Filarkiv</h4>

                                    <%call fileoverview(jobid, strKnr) %>
                                </div>
                            </div>
                        </div>

                        <%if browstype_client <> "ip" then %>
                        <div class="col-lg-6" style="height:750px;">
                            <div class="portlet portlet-boxed">
                                <div class="portlet-body">
                                    <h4>Materialer</h4>

                                    <%if (cint(strStatus) = 1 OR cint(strStatus) = 4) OR level = 1 then %>
                                        <form id="oprmatform" action="jobs.asp?func=oprMat" method="post">
                                            <%call MatPanel(jobid) %>
                                        </form>
                                    <%end if %>

                                    <br />
                                    <%call matlistesimple(jobid) %>
                                </div>
                            </div>
                        </div>
                        <%end if %>
                    </div>

                    <%end if %>
                    
                    <div class="row">
                        <div class="col-lg-12">
                            <%=job_txt_053 %> <b><%=strLastEditDate %></b> <%=job_txt_054 %> <b><%=strLastEditor %></b>
                            <%if isdate(strCreatedate) = true AND strCreatedate <> "01-01-2002" then %>
                            <br />
                            <%=job_txt_630 %> <b><%=strCreatedate %></b> <%=job_txt_054 %> <b><%=strCreator %></b>
                            <%end if %>
                        </div>                   
                    </div> 
                    


                </div>
            </div><!-- /.portlet body -->
            </div><!-- /.container -->
          

        <%end select %>

        <%if browstype_client <> "ip" then %>
        </div><!-- /.wrapper -->
    </div><!-- /.content -->
    <%end if %>


<!--#include file="../inc/regular/footer_inc.asp"-->