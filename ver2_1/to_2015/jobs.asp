<%response.buffer = true%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../timereg/inc/job_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"--> 
<!--#include file="../timereg/inc/isint_func.asp"-->
<!--#include file = "../timereg/CuteEditor_Files/include_CuteEditor.asp" -->
 
<%
'** Jquery section 
if Request.Form("AjaxUpdateField") = "true" then

    'Oliver

Select Case Request.Form("control")

case "FN_kpers"
              
              dim kundeKpersArr
              redim kundeKpersArr(20)
              
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
                

                kundeKpersArr(0) = "<option value='0'>Vælg kontaktperson..</option>"
                z = 1

                oRec.open strSQL, oConn, 3
			   
			    while not oRec.EOF
				
			    
                if cint(v) = oRec("id") then
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

            kundeKpersArr(z) = "<option value='0'>Ingen kontaktpersoner fundet</option>"
              
               for z = 0 to UBOUND(kundeKpersArr)
              '*** ÆØÅ **'
              call jq_format(kundeKpersArr(z))
              kundeKpersArr(z) = jq_formatTxt
	          
	          Response.Write kundeKpersArr(z)
              next
		

end select
Response.end
end if
%>




<!--------------- Github 1 ----------------->


<% 

    function SQLBless3(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ".", ",")
		SQLBless3 = tmp
	end function
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	function SQLBless2(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless2 = tmp
	end function








    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if


    func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if



    if len(trim(request("FM_fomr"))) <> 0 then
    strFomr_rel = request("FM_fomr")
    strFomr_rel = replace(strFomr_rel, "X234", "#")
    else
    strFomr_rel = "#0#"
    end if



        




    '*** Forretningsområder **' 
	                strFomr_Gblnavn = ""
                    strFomr_relA = replace(strFomr_rel, "#", "")
                    strFomr_relA = split(strFomr_relA, ",")

                    fo = 0
                    for f = 0 to UBOUND(strFomr_relA)

                    strSQLfrel = "SELECT fomr.navn FROM fomr WHERE fomr.id = "& strFomr_relA(f) 

                    'Response.Write strSQLfrel
                    'Response.flush
                    
                    oRec3.open strSQLfrel, oConn, 3
                    if not oRec3.EOF then

                    if fo = 0 then
                    strFomr_Gblnavn = " ("
                    end if

                    strFomr_Gblnavn = strFomr_Gblnavn & oRec3("navn") & ", " 
                   
                    fo = fo + 1
                    end if
                    oRec3.close

                    next

                    if fo <> 0 then
                    len_strFomr_Gblnavn = len(strFomr_Gblnavn)
                    left_strFomr_Gblnavn = left(strFomr_Gblnavn, len_strFomr_Gblnavn - 2)
                    strFomr_Gblnavn = left_strFomr_Gblnavn & ")"

                    if len(strFomr_Gblnavn) > 50 then
                    strFomr_Gblnavn = left(strFomr_Gblnavn, 50) & "..)"
                    end if
                    end if 
                    fomrArr = split(request("FM_fomr"), ",")



                    formTemp = replace(request("FM_fomr"), "#", "X234")




    Select case func 
    case "dbopr" , "dbred"
                            
                                '********************** Opretter job ***************************

                                '********************** Henter variable ***************************
                                kid = request("FM_kunde")                             
                                kunderef = request("FM_kpers")
                                jobnr = request("FM_jobnr")
                                beskrivelse = ""
                                internnote = ""
                                bruttooms = 0
                                editor = session("user")
                                'dddato = year(now) & "-" & month(now) & "-" & day(now)
                                'jobstartdato = request("FM_jobstartdato")
                                'jobstartdato = year(jobstartdato) & "-" & month(jobstartdato) & "-" & day(jobstartdato)

                                jobslutdato = request("FM_jobslutdato")
                                'jobslutdato = year(jobslutdato) & "-" & month(jobslutdato) & "-" & day(jobslutdato)




        
                                fomrArr = split(request("FM_fomr"), ",")

                                jobans1 = request("FM_jobans_1")
                                jobans2 = request("FM_jobans_2")
                                jobans3 = request("FM_jobans_3")
                                jobans4 = request("FM_jobans_4")
                                jobans5 = request("FM_jobans_5")

                              
                                if len(trim(request("FM_jobans_proc_1"))) <> 0 then
				                jobans_proc_1 = request("FM_jobans_proc_1")
                                else
                                jobans_proc_1 = 0
                                end if

                                jobans_proc_1 = replace(jobans_proc_1, ".", "")
                                jobans_proc_1 = replace(jobans_proc_1, ",", ".")

				                if len(trim(request("FM_jobans_proc_2"))) <> 0 then
				                jobans_proc_2 = request("FM_jobans_proc_2")
                                else
                                jobans_proc_2 = 0
                                end if

                                jobans_proc_2 = replace(jobans_proc_2, ".", "")
                                jobans_proc_2 = replace(jobans_proc_2, ",", ".")

				                if len(trim(request("FM_jobans_proc_3"))) <> 0 then
				                jobans_proc_3 = request("FM_jobans_proc_3")
                                else
                                jobans_proc_3 = 0
                                end if

                                jobans_proc_3 = replace(jobans_proc_3, ".", "")
                                jobans_proc_3 = replace(jobans_proc_3, ",", ".")

				                if len(trim(request("FM_jobans_proc_4"))) <> 0 then
				                jobans_proc_4 = request("FM_jobans_proc_4")
                                else
                                jobans_proc_4 = 0
                                end if

                                jobans_proc_4 = replace(jobans_proc_4, ".", "")
                                jobans_proc_4 = replace(jobans_proc_4, ",", ".")

				                if len(trim(request("FM_jobans_proc_5"))) <> 0 then
				                jobans_proc_5 = request("FM_jobans_proc_5")
                                else
                                jobans_proc_5 = 0
                                end if

                                jobans_proc_5 = replace(jobans_proc_5, ".", "")
                                jobans_proc_5 = replace(jobans_proc_5, ",", ".")

                                errProcVal = 0

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
                                    end if
                
                                next



                                 '** forretningsområde mandatory 
                                call fomr_mandatory_fn()
                                if cint(fomr_mandatoryOn) = 1 AND len(trim(request("FM_fomr"))) = 0 then

                                call visErrorFormat
			                    errortype = 177
			                    call showError(errortype)
			                    Response.end


                                end if



                                
                                '**********************************
			                    '**** Jobdata ****
			                    '**********************************
			    
				
				                strNavn = replace(request("FM_navn"),chr(34), "")

                                '** Jobbeskrivelse **'
			                    strBesk = SQLBless2(request("FM_beskrivelse"))
			    
			                    '*** HTML Replace **'
                                call htmlreplace(strBesk)
			                    strBesk = htmlparseTxt
                



                                '***** Salg ansvarlige ******
                                if len(trim(request("FM_salgsans_1"))) <> 0 then ' er slags ansvarlige slået til / vist
                                salgsans1 = request("FM_salgsans_1")
				                salgsans2 = request("FM_salgsans_2")
				                salgsans3 = request("FM_salgsans_3")
				                salgsans4 = request("FM_salgsans_4")
				                salgsans5 = request("FM_salgsans_5")
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

                                salgsans_proc_1 = replace(salgsans_proc_1, ".", "")
                                salgsans_proc_1 = replace(salgsans_proc_1, ",", ".")

				                if len(trim(request("FM_salgsans_proc_2"))) <> 0 then
				                salgsans_proc_2 = request("FM_salgsans_proc_2")
                                else
                                salgsans_proc_2 = 0
                                end if

                                salgsans_proc_2 = replace(salgsans_proc_2, ".", "")
                                salgsans_proc_2 = replace(salgsans_proc_2, ",", ".")

				                if len(trim(request("FM_salgsans_proc_3"))) <> 0 then
				                salgsans_proc_3 = request("FM_salgsans_proc_3")
                                else
                                salgsans_proc_3 = 0
                                end if

                                salgsans_proc_3 = replace(salgsans_proc_3, ".", "")
                                salgsans_proc_3 = replace(salgsans_proc_3, ",", ".")

				                if len(trim(request("FM_salgsans_proc_4"))) <> 0 then
				                salgsans_proc_4 = request("FM_salgsans_proc_4")
                                else
                                salgsans_proc_4 = 0
                                end if

                                salgsans_proc_4 = replace(salgsans_proc_4, ".", "")
                                salgsans_proc_4 = replace(salgsans_proc_4, ",", ".")

				                if len(trim(request("FM_salgsans_proc_5"))) <> 0 then
				                salgsans_proc_5 = request("FM_salgsans_proc_5")
                                else
                                salgsans_proc_5 = 0
                                end if

                                salgsans_proc_5 = replace(salgsans_proc_5, ".", "")
                                salgsans_proc_5 = replace(salgsans_proc_5, ",", ".")

           

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
                                    end if
                
                                next
                                '***** Salg ansvarlige SLUT
                                


                                '*** Progrp på job ***'
                
                                strProjektgr1 = request("FM_projektgruppe_1")
				                strProjektgr2 = request("FM_projektgruppe_2")
				                strProjektgr3 = request("FM_projektgruppe_3")
				                strProjektgr4 = request("FM_projektgruppe_4")
				                strProjektgr5 = request("FM_projektgruppe_5")
				                strProjektgr6 = request("FM_projektgruppe_6")
				                strProjektgr7 = request("FM_projektgruppe_7")
				                strProjektgr8 = request("FM_projektgruppe_8")
				                strProjektgr9 = request("FM_projektgruppe_9")
				                strProjektgr10 = request("FM_projektgruppe_10")

                

                                'Response.write "strProjektgr1:" & strProjektgr1 &" strProjektgr2: "& strProjektgr2 &" strProjektgr3: "& strProjektgr3&" strProjektgr4: "& strProjektgr4&" strProjektgr5: "& strProjektgr5&" strProjektgr6: "& strProjektgr6&" strProjektgr7: "& strProjektgr7&" strProjektgr8: "& strProjektgr8&" strProjektgr9: "& strProjektgr9&" strProjektgr10: "& strProjektgr10
                                'Response.end

                                if request("FM_gemsomdefault") = "1" then
				                response.cookies("job")("defaultprojgrp") = strProjektgr1
				                response.cookies("job").expires = date + 65
				                end if
                                


                                if len(trim(request("FM_interntbelob"))) <> 0 AND lCase(request("FM_interntbelob")) <> "nan" then
                                jo_gnsbelob = replace(request("FM_interntbelob"), ".", "")
		                        jo_gnsbelob = replace(jo_gnsbelob, ",", ".")
                                else
                                jo_gnsbelob = 0
                                end if
                
                                call erdetint(jo_gnsbelob)
		                        if isInt <> 0 then
		                        call visErrorFormat
				                errortype = 125
                                call showError(errortype)
                                Response.end
		                        end if
		                        isInt = 0

                                if len(trim(request("FM_udgifter_ulev"))) <> 0 AND lCase(request("FM_udgifter_ulev")) <> "nan" then
                                jo_udgifter_ulev = replace(request("FM_udgifter_ulev"), ".", "")
		                        jo_udgifter_ulev = replace(jo_udgifter_ulev, ",", ".")
                                else
                                jo_udgifter_ulev = 0
                                end if

                                call erdetint(jo_udgifter_ulev)
		                        if isInt <> 0 then
		                        call visErrorFormat
				                errortype = 1252
                                call showError(errortype)
                                Response.end
		                        end if
		                        isInt = 0


                                if len(trim(request("FM_udgifter_intern"))) <> 0 AND lCase(request("FM_udgifter_intern")) <> "nan" then
                                jo_udgifter_intern = replace(request("FM_udgifter_intern"), ".", "")
		                        jo_udgifter_intern = replace(jo_udgifter_intern, ",", ".")
                                else
                                jo_udgifter_intern = 0
                                end if

                                call erdetint(jo_udgifter_intern)
		                        if isInt <> 0 then
		                        call visErrorFormat
				                errortype = 1251
                                call showError(errortype)
                                Response.end
		                        end if
		                        isInt = 0


                
               
		        
		                        if len(trim(request("FM_udgifter"))) <> 0 AND lCase(request("FM_udgifter")) <> "nan" then
		                        udgifter = replace(request("FM_udgifter"), ".", "")
                                udgifter = replace(udgifter, ",", ".")
                                else
                                udgifter = 0
                                end if
                
                                call erdetint(udgifter)
		                        if isInt <> 0 then
		                        call visErrorFormat
				                errortype = 129
                                call showError(errortype)
                                Response.end
		                        end if
		                        isInt = 0
				
				
                                if len(trim(request("FM_restestimat"))) <> 0 then
                                restestimat = request("FM_restestimat")

                                        call erdetint(restestimat)
		                                if isInt <> 0 then
		                                call visErrorFormat
				                        errortype = 156
                                        call showError(errortype)
                                        Response.end
		                                end if
		                                isInt = 0

                                        restestimat = abs(restestimat)

                                else
                                restestimat = 0
                                end if

                                call erDetInt(request("FM_budgettimer"&simpeludvEXT&""))
				
                                if isInt > 0 OR instr(request("FM_budget"&simpeludvEXT&""), "-") <> 0 OR trim(lcase(request("FM_budget"&simpeludvEXT&""))) = "nan" then
				                
                                response.End
                                end if


                                if len(trim(request("FM_budgettimer"&simpeludvEXT&""))) = 0 then
				                    strBudgettimer = 0
				                    else
				                    strBudgettimer = request("FM_budgettimer"&simpeludvEXT&"")
				                    strBudgettimer = replace(strBudgettimer, ".", "")
				                    strBudgettimer = replace(strBudgettimer, ",", ".")
				                end if




                                '''Tjekekr om brutto beløb er mindre end netto beløb
                                select case lto 
                                case "epi", "epi_no", "epi_sta", "epi_ab", "epi_uk"
                                strBudgetTjk = replace(strBudget, ".", ",")
                                jo_gnsbelobTjk = replace(jo_gnsbelob, ".", ",")
                                if cdbl(strBudgetTjk) < cdbl(jo_gnsbelobTjk) then
                
                                        call visErrorFormat
				                        errortype = 165
                                        call showError(errortype)
                                        Response.end
		               
                
                
                                end if 
                                end select            

            

                                 if len(trim(request("FM_forvalgt"))) <> 0 then
                                 forvalgt = 1
                                 else
                                 forvalgt = 0
                                 end if

                                
                                '*********************************************************************'
				                '************************** Internt / eksternt job *******************'
				                '*********************************************************************'
				                '*** Altid Eksternt job ***'
						
						                if request("FM_usetilbudsnr") = "j" then
						                strStatus = 3 'tilbud
						                else
							                if request("FM_status") <> 3 then
							                strStatus = request("FM_status")
							                else
							                strStatus = 1 'aktivt, hvis der ved en fejl at valgt tilbud
							                end if
						                end if

                            

                                if len(trim(request("FM_ikkebudgettimer"&simpeludvEXT&""))) = "0" then
				                ikkeBudgettimer = 0
				                else
				                ikkeBudgettimer = request("FM_ikkebudgettimer"&simpeludvEXT&"")
				                end if


                                if len(trim(request("FM_fastpris"))) = "0" then 
				                strFastpris = 0 '0: lbn. timer / 1: fastpris 
				                else
				                strFastpris = request("FM_fastpris")
				                end if


                                if len(request("FM_kundese")) <> 0 then
					            if request("FM_kundese_hv") = 1 then
					            intkundese = 2 '(når job er lukket)
					            else
					            intkundese = 1 '(når timer er indtastet)
					            end if
				                else
				                intkundese = 0
				                end if


                

                                if len(trim(request("FM_syncaktdatoer"))) <> 0 then
                                syncaktdatoer = 1
                                else
                                syncaktdatoer = 0
                                end if


                                if len(trim(request("FM_syncslutdato"))) <> 0 then
                                syncslutdato = 1
                                else
                                syncslutdato = 0
                                end if



                                strFakturerbart = request("FM_fakturerbart")



                                if func = "dbopr" then
				
				                '**********************************'
				                '*** nyt jobnr / tilbudsnr ***'
				                '**********************************'
				                strSQL = "SELECT jobnr, tilbudsnr FROM licens WHERE id = 1"
				                oRec5.open strSQL, oConn, 3
				                if not oRec5.EOF then 
				    
				                    strjnr = oRec5("jobnr") + 1
				    
				                    if request("FM_usetilbudsnr") = "j" then
				                    tlbnr = oRec5("tilbudsnr") + 1
				                    else
				                    tlbnr = 0
				                    end if
				    
				                end if
				                oRec5.close

                    
				
				                else
				  
				                  strjnr = request("FM_jnr")
                
                                 call alfanumerisk(strjnr)
                                 strjnr = alfanumeriskTxt
                                 strjnr = left(strjnr,20)

				                  tlbnr = request("FM_tnr")  
				    
				                end if



                                '*** tilbudsnr findes ***
					
					            if request("FM_usetilbudsnr") = "j" then
					            tilbudsnrFindes = 0
					
				                if func = "dbred" then
				                strSQL = "SELECT tilbudsnr, id FROM job WHERE id <> "& id &" AND tilbudsnr = '" &  tlbnr & "' AND tilbudsnr <> 0"
					            else
					            strSQL = "SELECT tilbudsnr, id FROM job WHERE tilbudsnr = '" & tlbnr & "'"
					            end if 
					
					            'Response.Write strSQL
					            'Response.Flush
					            'Response.end
					
					            oRec5.open strSQL, oConn, 3
		                        if not oRec5.EOF then	
		            
		                        tilbudsnrFindesNR = oRec5("tilbudsnr")
		                        tilbudsnrFindes = 1 
		            
		                        end if
					            oRec5.close
					
					            end if 
					
					
					            if cint(tilbudsnrFindes) = 1 then
					            %>
					            <!--#include file="../inc/regular/header_inc.asp"-->
				                <% 
					            errortype = 94
					            call showError(errortype)
					            Response.end
				                end if




                                if len(request("FM_serviceaft")) <> 0 then
				                intServiceaft = request("FM_serviceaft") '1
				                else
				                intServiceaft = 0
				                end if


                                 if len(trim(request("FM_restestimat"))) <> 0 then
                                restestimat = request("FM_restestimat")

                                        call erdetint(restestimat)
		                                if isInt <> 0 then
		                                call visErrorFormat
				                        errortype = 156
                                        call showError(errortype)
                                        Response.end
		                                end if
		                                isInt = 0

                                        restestimat = abs(restestimat)

                                else
                                restestimat = 0
                                end if


                                if len(request("FM_valuta")) <> 0 then
				                valuta = request("FM_valuta")
				                else
				                valuta = 1 'main valuta
				                end if

                               rekvnr = replace(request("FM_rekvnr"), "'", "''")
				
				                if len(trim(rekvnr)) = 0 AND lto = "ets-track" then
		                        call visErrorFormat
				                errortype = 149
                                call showError(errortype)
                                Response.end
		                        end if


                                '* Intern note **'
			                    strInternBesk = SQLBless2(request("FM_internbesk"))
			    
			                    '*** HTML Replace **'
			                    call htmlreplace(strInternBesk)
			                    strInternBesk = htmlparseTxt



                            isInt = 0
			                call erDetInt(request("prio")) 
			                if isInt > 0 then
			    
			                    call visErrorFormat
				
				                errortype = 132
				                call showError(errortype)
			
			                response.End
			                end if


                            if len(request("FM_opr_kpers")) <> 0 then
				            intKundekpers = request("FM_opr_kpers")
				            else
				            intKundekpers = 0
				            end if
				
                
                            if cint(intKundekpers) = 0 AND (lto = "dencker") AND rdir <> "redjobcontionue" then
		                    call visErrorFormat
				            errortype = 155
                            call showError(errortype)
                            Response.end
		                    end if


                            if len(trim(request("prio"))) <> 0 then
				            intprio = request("prio")
				            else
 				            intprio = -1
 				            end if


                            if len(trim(request("FM_fomr_konto"))) <> 0 then
                            fomr_konto = request("FM_fomr_konto")
                            else
                            fomr_konto = 0
                            end if



                            jfak_moms = request("FM_jfak_moms")
                            jfak_sprog = request("FM_jfak_sprog")


                            '************************************'
                             '***** Forrretningsområder **********'
                             '************************************'

                   


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
    




                            if len(trim(request("FM_brugaltfakadr"))) <> 0 then
                            altfakadr = 1
                            else
                            altfakadr = 0
                            end if





                            if len(trim(request("FM_preconditions_met"))) <> 0 then
                            preconditions_met = request("FM_preconditions_met")
                            else
                            preconditions_met = 0
                            end if



                               
                            '*** Jobnr findes ***
				            jobnrFindes = 0
				    
				            if func = "dbred" then
				            strSQL = "SELECT jobnr, id FROM job WHERE id <> "& id &" AND jobnr = '" & strjnr & "'"
					        else
					        strSQL = "SELECT jobnr, id FROM job WHERE jobnr = '" & strjnr &"'"
					        end if 
					
					        'Response.Write strSQL
					        'Response.Flush
					
					
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
                      


                        
                     
                            startDato = Request("FM_start_aar") &"/" & Request("FM_start_mrd") & "/" & strStartDay 
					        if request("FM_datouendelig") = "j" then
					        slutDato =  "2044/1/1" 
					        else
					        slutDato =  Request("FM_slut_aar") &"/" & Request("FM_slut_mrd") & "/" & strSlutDay 
					        end if

                            
                            if func = "dbopr" then


                                '**********************************'
				                '*** nyt jobnr / tilbudsnr ***'
				                '**********************************'
				                strSQL = "SELECT jobnr, tilbudsnr FROM licens WHERE id = 1"
				                oRec5.open strSQL, oConn, 3
				                if not oRec5.EOF then 
				    
				                    strjnr = oRec5("jobnr") + 1
				    
				                    if request("FM_usetilbudsnr") = "j" then
				                    tlbnr = oRec5("tilbudsnr") + 1
				                    else
				                    tlbnr = 0
				                    end if
				    
				                end if
				                oRec5.close

      

                                strSQLjob = ("INSERT INTO job (jobnavn, jobnr, jobknr, jobTpris, jobstatus, jobstartdato," _                    
                                & " jo_bruttooms, fomr_konto, risiko, preconditions_met, jfak_moms, jfak_sprog, altfakadr, syncslutdato, serviceaft) VALUES " _
                                & "('" & strNavn & "', " _
                                & "'" & strjnr & "', " _
                                & "" & kid & ", " _
                                &""& jo_gnsbelob &", "_
                                &""& strStatus &", "_
                                &"'"& startDato &"', "_
                                &"'"& slutDato &"', "_
                                & "'" & editor & "', " _
                                & "'" & dddato & "', " _
                                &""& strProjektgr1 &", "_ 
							    &""& strProjektgr2 &", "_ 
							    &""& strProjektgr3 &", "_ 
							    &""& strProjektgr4 &", "_ 
							    &""& strProjektgr5 &", "_
							    &""& strProjektgr6 &", "_ 
							    &""& strProjektgr7 &", "_ 
							    &""& strProjektgr8 &", "_ 
							    &""& strProjektgr9 &", "_ 
							    &""& strProjektgr10 &", "_
                                &""& strFakturerbart &", "_ 
							    &""& strBudgettimer  &", "_ 
							    &"'"& strFastpris & "',"_
							    &""& intkundese &", "_
                                & "'" & strBesk & "', " _
                                &""& SQLBless(ikkeBudgettimer) &", "_
                                &""& tlbnr &", "_                             
                                & "" & jobans1 & "," _
                                & "" & jobans2 & "," _
                                & "" & jobans3 & "," _
                                & "" & jobans4 & "," _
                                & "" & jobans5 & "," _                                                       
                                &" "& jobans_proc_1 & ", "& jobans_proc_2 & ", "& jobans_proc_3 & ", "& jobans_proc_4 & ", "& jobans_proc_5 & ","_
                                &" "& salgsans1 &","& salgsans2 &","& salgsans3 &","& salgsans4 &","& salgsans5 &", "_
                                &" "& salgsans_proc_1 &","& salgsans_proc_2 &","& salgsans_proc_3 &","& salgsans_proc_4 &","& salgsans_proc_5 & ","_
                                & "" & intServiceaft & ", " _
                                & "" & intKundekpers & ", " _
                                & "" & valuta & ", " _
                                & "'" & rekvnr & "', " _
                                & "" & intprio & ", " _
                                & "'" & strInternBesk & "', " _
                                & "" & strBudget & ", " _
                                & "" & fomr_konto & ", " _
                                & "" & intprio & ", " _
                                & "" & preconditions_met & ", " _
                                & "" & jfak_moms & ", " _
                                & "" & jfak_sprog & ", " _
                                & "" & altfakadr & ", " _
                                & "" & syncslutdato & ", " _
                                & "" & intServiceaft & "" _
                                &")")
                                
                                
                                response.write "strSQLjob: " & strSQLjob 
                                response.flush

                                oConn.execute(strSQLjob)
                                
                                

                                '******************************************'
								'*** finder jobid på det netop opr. job ***'
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
                                
                                
                                
                                
                                
                                '*** Opretter timepriser på job.       *'
								'***************************************'
                                
						        'Response.write "HER"
                                'response.end
                        
                                'timeE = now
                                'loadtime = datediff("s",timeA, timeE)
							    'Response.write "<br>Før timepriser: "& loadtime & "<br><br>"		


								strSQLpg = "SELECT id, navn FROM projektgrupper WHERE "_
								&" id = "& strProjektgr1 &" OR "_
								&" id = "& strProjektgr2 &" OR id = "& strProjektgr3 &" OR "_
								&" id = "& strProjektgr4 &" OR id = "& strProjektgr5 &" OR id = "& strProjektgr6 &" OR "_
								&" id = "& strProjektgr7 &" OR id = "& strProjektgr8 &" OR id = "& strProjektgr9 &" OR "_
								&" id = "& strProjektgr10 &" ORDER BY navn"
								oRec5.open strSQLpg, oConn, 3
								
								'Response.write strSQL
								while not oRec5.EOF
									
									strSQLmtp = "SELECT medarbejderid, projektgruppeid, mid, mnavn, timepris, timepris_a1, "_
									&" timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
									&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta "_
									&" FROM progrupperelationer "_
									&" LEFT JOIN medarbejdere ON (mid = progrupperelationer.medarbejderid) "_
									&" LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) "_
									&" WHERE projektgruppeid = "& oRec5("id") &" AND mnavn <> '' AND mansat <> 2 ORDER BY mnavn"
									oRec3.open strSQLmtp, oConn, 3
									this6timepris = 0
									while not oRec3.EOF
								        
								        
								        '*** Finder valuta på job og finder timepris ***'
								        '*** på medarbejder der matcher valuta ***'
								        '*** Hvis findes ellers vælges 1 = DKK, Grundvaluta ****'
								            
								           valutafundet = 0
								             
								           for i = 0 to 5
								           
								                   'Response.Write "her" & i & "<br>"
								           
								                   select case i
								                   case 0
								                   tpris = oRec3("timepris")
								                   tprisValuta = oRec3("tp0_valuta")
								                   tprisAlt = 0
								                   case 1
								                   tpris = oRec3("timepris_a1")
								                   tprisValuta = oRec3("tp1_valuta")
								                   tprisAlt = 1
								                   case 2
								                   tpris = oRec3("timepris_a2")
								                   tprisValuta = oRec3("tp2_valuta")
								                   tprisAlt = 2
								                   case 3
								                   tpris = oRec3("timepris_a3")
								                   tprisValuta = oRec3("tp3_valuta")
								                   tprisAlt = 3
								                   case 4
								                   tpris = oRec3("timepris_a4")
								                   tprisValuta = oRec3("tp4_valuta")
								                   tprisAlt = 4
								                   case 5
								                   tpris = oRec3("timepris_a5")
								                   tprisValuta = oRec3("tp5_valuta")
								                   tprisAlt = 5
								                   end select
								           
                                                    'if lto = "cisu" then
                                
                                                    '    response.write "<br>Medid: "& oRec3("medarbejderid") &":"& oRec3("mid") &" projektgruppeID: "& oRec5("id") &" i: "& i &" valuta: " & valuta &" tprisValuta = "& tprisValuta &" valutafundet: "& valutafundet
                                                    '    response.flush 

                                                    'end if

								           
								                    if cint(valuta) = cint(tprisValuta) AND cint(valutafundet) = 0 then
								                    tprisUse = tpris
								                    tprisValutaUse = tprisValuta
								                    tprisAltUse = tprisAlt
								                    valutafundet = 1
								            
								                    end if
								            
							                next
								           
								           if cint(valutafundet) = 0 then
								           tprisUse = oRec3("timepris")
								           tprisValutaUse = oRec3("tp0_valuta")
								           tprisAltUse = 0
								           end if
								            
                                           tprisUse = replace(tprisUse, ".", "")
                                           tprisUse = replace(tprisUse, ",", ".")
								        
								        '*** Opretter timepriser på job **'
								        if instr(usedmids, "#"&oRec3("mid")&"#") = 0 then
								        strSQLtp = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
								        &" VALUES ("& varJobId &", 0, "& oRec3("mid") &","& tprisAltUse &", "& tprisUse &", "& tprisValutaUse &")"
								        
                                        'if lto = "qwert" then
								        'Response.Write strSQLtp & "<br>"
								        'Response.flush
                                        'end if
								        
								        oconn.execute(strSQLtp)
								        end if
								   
								   usedmids = usedmids &oRec3("mid")&"#"
								   oRec3.movenext     
								   wend
								   oRec3.close
							    
							    
						     oRec5.movenext     
							 wend
							 oRec5.close
							   
                             'end if 't = 0
                            
                              'if lto = "qwert" then
							  'Response.end 
                              'end if 
								
                            'timeC = now
                            'loadtime = datediff("s",timeA, timeC)
							'Response.write "<br>Før stam: "& loadtime & "<br><br>"
                                   


                                
                                '***************************************************************************************************
                                'PUSH - Aktiviteter - Timepriser
                                 '*********** timereg_usejob, så der kan søges fra jobbanken KUN VED OPRET JOB *********************
                                Select Case lto
                                    Case "oko"
                                    Case "outz", "intranet - local", "demo"
                       
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
                        
                        
                
                                '** Vælg Stamaktgruppe pbg. af projetktype
                                agforvalgtStamgrpKri = " ag.forvalgt = 1 "
                                
                                
                                
                                '*** Indlæser STAMaktiviteter ***'
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
                
                                   end if
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
                                    '*** Hvis timepris ikke findes på job bruges Gen. timepris fra '
                                    '*** Fra medarbejdertype, og den oprettes på job              *'
                                    '*** BLIVER ALTID HENTET FRA Medarb.type for ØKO              *'
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
                        
                        
                                        '**** Indlæser timepris på aktiviteter ***'
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
				                nytjobnr = strjnr
				                tilbudsnrKri = ""
				                
				                if request("FM_usetilbudsnr") = "j" then
				                nyttlbnr = tlbnr
				                tilbudsnrKri = ", tilbudsnr = "& nyttlbnr &""
				                end if
				                
				                strSQL = "UPDATE licens SET jobnr = "&  nytjobnr &" "& tilbudsnrKri &" WHERE id = 1"
				                oConn.execute(strSQL)

                                else 'rediger


                                ''''''''''''''''''''''''''' UPDATE SKAL SKE HER '''''''''''''''''''
                            
                             
                                '*** Opdater LUKKE dato (før det skifter stastus) ***'
                                call lukkedato(id, strStatus)
                                

                                strSQL = "UPDATE job SET "_
							    &" jobnavn = '"& strNavn &"', "_
							    &" jobnr = '"& strjnr &"', "_
                                &" rekvnr = '"& rekvnr &"', "_ 
                                &" jobstatus = "& strStatus &", "_
                                &" jobans1 = "& Jobans1 &", "_
                                &" jobans2 = "& Jobans2 &", "_
                                &" jobans3 = "& Jobans3 &", "_
                                &" jobans4 = "& Jobans4 &", "_
                                &" jobans5 = "& Jobans5 &", "_ 
                                &" jobans_proc_1 = "& jobans_proc_1 & ", "_
                                &" jobans_proc_2 = "& jobans_proc_2 & ", "_
                                &" jobans_proc_3 = "& jobans_proc_3 & ", "_
                                &" jobans_proc_4 = "& jobans_proc_4 & ", "_
                                &" jobans_proc_5 = "& jobans_proc_5 & ", "_
                                &" beskrivelse = '"& strBesk &"', "_                              
                                &" salgsans1 = "& salgsans1 &", salgsans2 = "& salgsans2 &", salgsans3 = "& salgsans3 &", salgsans4 = "& salgsans4 &", salgsans5 = "& salgsans5 &", "_
                                &" salgsans1_proc = "& salgsans_proc_1 &", salgsans2_proc = "& salgsans_proc_2 &", salgsans3_proc = "& salgsans_proc_3 &", salgsans4_proc = "& salgsans_proc_4 &", salgsans5_proc = "& salgsans_proc_5 &", "_  
                                &" projektgruppe1 = "& strProjektgr1 &", "_ 
                                &" projektgruppe2 = "& strProjektgr2 &", "_
                                &" projektgruppe3 = "& strProjektgr3 &", "_
                                &" projektgruppe4 = "& strProjektgr4 &", "_
                                &" projektgruppe5 = "& strProjektgr5 &", "_
                                &" projektgruppe6 = "& strProjektgr6 &", "_
                                &" projektgruppe7 = "& strProjektgr7 &", "_
                                &" projektgruppe8 = "& strProjektgr8 &", "_
                                &" projektgruppe9 = "& strProjektgr9 &", "_
                                &" projektgruppe10 = "& strProjektgr10 &","_
                                &" fomr_konto = "& fomr_konto &", "_
                                &" risiko = "& intprio &", "_
                                &" preconditions_met = "& preconditions_met &", "_
                                &" valuta = "& valuta &", "_
                                &" jfak_moms = "& jfak_moms &", "_ 
                                &" jfak_sprog = "& jfak_sprog &", "_ 
                                &" kundeok = "& intkundese &", "_ 
                                &" altfakadr = "& altfakadr &", "_
                                &" syncslutdato = "& syncslutdato &""_                                                                                                       
                                &" WHERE id = "& id 

							   '&" serviceaft = "& intServiceaft &""_

							    'Response.Write strSQL
							    'Response.end
                                'Response.flush							
							    oConn.execute(strSQL)





                                '****** opdaterer tilbudsnr ved rediger ****'
                                                if request("FM_usetilbudsnr") = "j" then
				                                nyttlbnr = tlbnr
                                                     
                                                     senesteTilbnr = 0
                                                     strSQL = "SELECT tilbudsnr FROM licens WHERE id = 1"
                                                     oRec5.open strSQL, oConn, 3
                                                     if not oRec5.EOF then
                                                     senesteTilbnr = oRec5("tilbudsnr")
                                                     end if
                                                     oRec5.close

				                                end if
				                                
                                                if cdbl(nyttlbnr) <> 0 AND cdbl(nyttlbnr) >= cdbl(senesteTilbnr) then
				                                strSQL = "UPDATE licens SET tilbudsnr = "& nyttlbnr &" WHERE id = 1"
				                                oConn.execute(strSQL)
				                                end if






                                '***********************************'
                                                '** Opdaterer kundeoplysninger *****'
                                                '***********************************'

								
								                '*** Overfører gamle timeregistreringer til ny aftale (hvis der skiftes aftale) **'
								                'if request("FM_overforGamleTimereg") = "1" then
                								
								                '*** Opdaterer timereg tabellen ****
								                'strSQLtimer = "UPDATE timer SET "_
								                '& " Tknavn = '"& replace(oRec("kkundenavn"), "'", "''") &"', Tknr = "& oRec("kid")&", "_
								                '& " Tjobnavn = '"& strNavn &"', "_
								                '& " Tjobnr = '"& strjnr &"', "_
								                '& " fastpris = '"& strFastpris &"', "_
								                '& " seraft = "& intServiceaft &" "_
								               ' & " WHERE Tjobnr = '"& strOLDjobnr & "'"
                								
                								'** Husk materiale forbrug ***
								                'strSQLmat_forbrug = "UPDATE materiale_forbrug SET serviceaft = " & intServiceaft &""_
								                '& " WHERE jobid = " & id
                								
								                'oConn.execute(strSQLmat_forbrug)
                								
                								
								                'else
                								
								                '*** Opdaterer timereg tabellen ****
								                'strSQLtimer = "UPDATE timer SET "_
								                '& " Tknavn = '"& replace(oRec("kkundenavn"), "'", "''") &"', Tknr = "& oRec("kid") &", "_
								                '& " Tjobnavn = '"& strNavn &"', "_
								                '& " Tjobnr = '"& strjnr &"', "_
								                '& " fastpris = '"& strFastpris &"' "_
								                '& " WHERE Tjobnr = '"& strOLDjobnr & "'"
                								
								                'end  if
                								
                								
								                'Response.write strSQLtimer
								                'Response.flush
                								
								                'oConn.execute(strSQLtimer)
								                
								                
								                '*** Opdaterer faktura tabel så faktura kunde id passer hvis der er skiftet kunde  ved rediger job.
								                '*** Adr. i adresse felt på faktura behodes til revisor spor. **'
								                'strSQLFakadr = "UPDATE fakturaer SET fakadr = "& oRec("kid") &" WHERE jobid = " & id
								                'oConn.execute(strSQLFakadr)
								
								
								varJobId = id






                                '***** Sync Job slutDato ***********'
                                if strStatus = 0 then 'Kun ved luk job
                                syncslutdato = syncslutdato
                                else
                                syncslutdato = 0
                                end if

                                call syncJobSlutDato(varJobId, strjnr, syncslutdato)

                               
                                '****** Sync. datoer på aktiviteter *******'
                                if syncaktdatoer = 1 then
                                    
                                    '** Hvis sync jobslutdato er valgt ***
                                    if syncslutdato <> 1 then
                                    useSlutDato = slutDato
                                    else
                                    useSlutDato = useSyncDato 
                                    end if

                                strSQLaDatoer = "UPDATE aktiviteter SET aktstartdato = '"& startDato &"', aktslutdato = '"& useSlutDato &"' WHERE job = "& varJobId
                                oConn.execute(strSQLaDatoer)

                                end if



                                


                                '************************************************'
				                '**** Diffentierede timepriser ******************'
			                    '*** Opdater eksisterende time-registreringer ***'
			                    '************************************************'
                				
                                '** sync alle aktiviteter ***'
                		        if request("FM_sync_tp") = "1" then
                                syncAkt = 1
                                else
                                syncAkt = 0
                                end if

                                 


				                medarb_tpris = request("FM_use_medarb_tpris")
				                Dim intMedArbID 
				                Dim b
                				strMedabTimePriserSlet = " AND medarbid <> 0"

					                intMedArbID = Split(medarb_tpris, ", ")
					                For b = 0 to Ubound(intMedArbID)
                                        
                                         
                                         strMedabTimePriserSlet = strMedabTimePriserSlet & " AND medarbid <> "& intMedArbID(b)


                                        if cint(syncAkt) = 1 then
                                            
                                            '** ikke KM **
                                            fakbaraktid = ""
                                            strSQLatyp = "SELECT id, fakturerbar FROM aktiviteter WHERE job = " & id & " AND fakturerbar = 5"
                                            oRec5.open strSQLatyp, oConn, 3
                                            fb = 0
                                            while not oRec5.EOF 
                                                
                                                if fb <> 0 then
                                                fakbaraktid = fakbaraktid & " OR "
                                                end if

                                                if oRec5("fakturerbar") = 5 then
                                                fakbaraktid = fakbaraktid & "aktid <> "& oRec5("id") 
                                                else
                                                fakbaraktid = fakbaraktid 
                                                end if

                                                fb = fb + 1
                                            oRec5.movenext
                                            wend
                                            oRec5.close

                                            if len(trim(fakbaraktid)) <> 0 then
                                            fakbaraktid = " AND ("& fakbaraktid & ")"
                                            else
                                            fakbaraktid = ""
                                            end if



						                strSQLdeltp = "DELETE FROM timepriser WHERE jobid = "& id &" AND medarbid = "& intMedArbID(b) &" "& fakbaraktid 
                                        oConn.execute(strSQLdeltp)

                                        
                                        else '** ellers kun job // Eller valgte akt **'

                                        strSQLdeltp = "DELETE FROM timepriser WHERE jobid = "& id &" AND medarbid = "& intMedArbID(b) &" AND aktid = "& hd_tp_jobaktid 
                                        oConn.execute(strSQLdeltp)
                        
                                        end if

                                        
                                        strSQLdeltp = strSQLdeltp & "<br>" & strSQLdeltp


                                         
						               
                                        'if len(request("FM_6timepris_"&intMedArbID(b)&"")) <> 0 OR jobids(j) = 0 then
                						
						                if len(request("FM_6timepris_"&intMedArbID(b)&"")) <> 0 then
						                call erDetInt(request("FM_6timepris_"&intMedArbID(b)&""))
						                    if isInt > 0 then
						                    this6timepris = 0
						                    else
						                    this6timepris = SQLBless(replace(request("FM_6timepris_"&intMedArbID(b)&""),".",""))
						                    end if
						                
						                valuta6 = request("FM_valuta_600"& intMedArbID(b))
						                
						                else
						                this6timepris = 0
						                valuta6 = 1
						                end if
						               
                						
                                       ' if hd_tp_jobaktid = 0 then '** Job / Akt, Hviklen radio bt er valgt ***'
                                       '*** Ved sync nedarver alle og der skal derfor ikke indsættes en record heller ikke for den aktivitet man står på ***'
                                       '**** Kun hvis man står på selve jobbet ***'

                                        if len(request("FM_6timepris_"&intMedArbID(b)&"")) <> 0 AND (cint(syncAkt) <> 1 OR (cint(syncAkt) = 1 AND hd_tp_jobaktid = 0)) then

                                            radiobtVal = SQLBless(replace(request("FM_hd_timepris_"&intMedArbID(b)&"_"&request("FM_timepris_"&intMedArbID(b)&"")),".",""))

                						    if this6timepris = radiobtVal then
                						    tprisalt = request("FM_timepris_"&intMedArbID(b)&"")
                						    else
                						    tprisalt = 6
                                            end if
                						
                                            
                                            
						                    strSQLupdTimePriser = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
						                    &" VALUES ("& id &","& hd_tp_jobaktid &", "& intMedArbID(b) &", "& tprisalt &", "& this6timepris &", "& valuta6 &")"
                						
                                       
                                            'Response.Write strSQLupdTimePriser & "<br>"
                                            'Response.flush

                                             oConn.execute(strSQLupdTimePriser)

                                       
                                       

                                         end if
                                        






                                        
                                        '*********************************************************'
						                '** Opdaterer timereg (kun på aktiviteter der nedarver) **'
                                        '** Dvs dem der ikke findes i timepriser tabellen       **'

                                        '*** Else den specifikke aktivitet ***
						               
						                '*** Kræver at der minimun findes 1 akt. *****************
						              	'*** Ved sync er denne tom (lige slettet ovenfor) og alle **'
                                        '**' timeprsier på alle aktiviter forall medeb. opdateres ***'
                                        '*********************************************************'


                                        if cint(syncAkt) = 1 OR hd_tp_jobaktid = 0 then
										strSQLtimp = "SELECT aktid FROM timepriser WHERE jobid = "& id &" AND medarbid = "& intMedArbID(b) &" AND aktid <> 0"
										
										'Response.Write strSQLtimp & "<br>"
										'Response.flush
										
										notAkt = ""
										oRec5.open strSQLtimp, oConn, 3
										while not oRec5.EOF
										notAkt = notAkt & " AND taktivitetId <> "& oRec5("aktid") 
										oRec5.movenext
										wend
										oRec5.close 

                                        specifikakt = 0

                                        else

                                        specifikakt = 1
                                        notAkt = " AND taktivitetId = "& hd_tp_jobaktid 
                                        
                                        aktnavn = ""
                                        strSQLalleaktmnavn = "SELECT navn FROM aktiviteter WHERE id = "& hd_tp_jobaktid 
                                        oRec6.open strSQLalleaktmnavn, oConn, 3
                                        if not oRec6.EOF then
                                        aktnavn = oRec6("navn")
                                        end if
                                        oRec6.close

                                        end if






                                        '*** Opdaterer alle timeregistreringer på valgte akt. eller på alle aktiviteter ved sync ***'

										   '**** Finder aktuel kurs ***'
						                   strSQL = "SELECT kurs FROM valutaer WHERE id = " & valuta6
						                   oRec5.open strSQL, oConn, 3
        						            
						                   if not oRec5.EOF then
						                   nyKurs = replace(oRec5("kurs"), ",", ".")
						                   end if 
						                   oRec5.close

                                         '** Fra dato ved opdater timepriser ***'
                                        fraDato = request("FM_opdatertpfra") 
                                        if isDate(fraDato) = false then
                                        fraDato = year(now) &"/"& month(now) & "/"& day(now)
                                        else
                                        fraDato = year(fraDato) &"/"& month(fraDato) & "/"& day(fraDato) 
                                        end if

										
					                    
					                    
                                        if request("FM_opdateralleakt") = "1" AND specifikakt = 1 then 
                                        '** Specifik akt + alle med samme navn **' 
			                            strSQL = "UPDATE timer SET timepris = "& this6timepris  &", valuta = "& valuta6 &", kurs = "& nyKurs &""_
			                            &" WHERE (tmnr = "& intMedArbID(b) &" AND tdato >= '"& fraDato &"' AND (tfaktim <> 5)) AND taktivitetnavn = '" & aktnavn &"'"
                                        else
                                        '** Nedarv (IKKE kørsel) / specifik akt 
			                            strSQL = "UPDATE timer SET timepris = "& this6timepris  &", valuta = "& valuta6 &", kurs = "& nyKurs &""_
			                            &" WHERE (tmnr = "& intMedArbID(b) &" AND tjobnr = '"& strjnr &"' AND tdato >= '"& fraDato &"' AND (tfaktim <> 5)) " & notAkt
					                    end if

                                        'Response.Write strSQL & "<br>"
					                    'Response.end
					                    
					                    oConn.execute(strSQL)
                                                

                             next '** intMedArbID(b)





                            '*************************************************************************************
                            '*** Renser ud i timepriser på medarbejdere der ikke længere er tilknyttet jobbbet ***
                            '*************************************************************************************
                                       
                            '*** Pilot WWF 26-09-2011 *** 
                            'if lto = "wwf" then

                            if len(trim(request("FM_sync_tp_rens"))) <> 0 then
                            sync_tp_rens = 1
                            else
                            sync_tp_rens = 0
                            end if

                            '** KUN VED opdater timepriser bliver strMedabTimePriserSlet SAT ELLERS NULSTILLES timepriser på alle medarb for alle aktiviteter ***'
                            '** FARLIG ASSURATOR; OKO oplever alle deres timepriser på medarbejdere forsvinder ***'
                		    if cint(sync_tp_rens) = 1 AND strMedabTimePriserSlet <> " AND medarbid <> 0" then
                            strSQLRensTp = "DELETE FROM timepriser WHERE jobid = "& id &" "& strMedabTimePriserSlet 
                            oConn.execute(strSQLRensTp)
                            'end if

                            'if session("mid") = 1 then

                            'Response.Write strSQLRensTp
                            'Response.flush
                            'response.write "<hr>"

                            'response.write strSQLdeltp
                            
                            'Response.end
                            'end if

                            end if


                            '**** Timepriser END ***'






                            '**********************************************************'
								'*** Opdaterer allerede tilknyttet aktiviteter   **********'
								
								
								    if len(request("FM_aktnavn")) > 3 then
	                                len_aktnavn = len(request("FM_aktnavn"))
	                                left_aktnavn = left(request("FM_aktnavn"), len_aktnavn - 3)
	                                strAktnavn = left_aktnavn
	                                else
	                                strAktnavn = ""
	                                end if
                                	
	                                aktnavn = strAktnavn
	                                akttimer = request("FM_akttimer")

                                    if len(trim(request("FM_aktkonto"))) <> 0 then
                                    aktkonto = request("FM_aktkonto")
                                    else
                                    aktkonto = ""
                                    end if

	                                aktantalstk = request("FM_aktantalstk")
	                                aktpris = request("FM_aktpris")
	                                
	                                'Response.Write request("FM_aktpris")
	                                'Response.end
	                                aktids = request("FM_aktid")

                                    'Response.Write "aktids" & aktids
	                                'Response.end

	                                aktstatus = request("FM_aktstatus")
                                	
                                    'Response.Write request("FM_aktfase")
                                    'Response.end

	                                if len(request("FM_aktfase")) > 3 then
	                                len_Fasenavn = len(request("FM_aktfase"))
	                                left_Fasenavn = left(request("FM_aktfase"), len_Fasenavn - 3)
	                                strFasenavn = left_Fasenavn
	                                else
	                                strFasenavn = ", #"
	                                end if
                                	
	                                aktfaser = strFasenavn
	                                aktbgr = request("FM_aktbgr")
                                	
	                                'Response.Write request("FM_akt_totpris")
	                                'Response.end
                                	
	                                akttotpris = request("FM_akttotpris")
                                	
	                                'Response.Write request("FM_fase")
	                                'Response.end
                                	
	                                aktslet = request("FM_slet_aid")
                                	
	                                'Response.Write aktSlet
	                                'Response.end
								    
								    aktslet_aids = 0
								
								    call opdateraktliste(varJobId, aktids, aktnavn, akttimer, aktantalstk, aktfaser, aktbgr, aktpris, aktstatus, akttotpris, aktslet, aktslet_aids, aktkonto)

								
								'Response.end
								
								'**********************************************************'
								'**** Tilknytter flere Stam-aktivitets grupper til det  ***'
								'**** job der bliver redigeret                          ***'
								'**********************************************************'
								
                                'Response.write "intAktfavgp " & intAktfavgp & "<br>"
                                'Response.write "strAktFase " & strAktFase & "<br>"
                                'Response.end

								intAktfavgp_use = split(intAktfavgp, ",")
								strAktFase_use = split(strAktFase, ", ")
                                firstLoop = 0
								for a = 0 to UBOUND(intAktfavgp_use)
								'for a = 1 to 1
                                    'Response.Write "a" & a & " val: "& intAktfavgp_use(a) &"<br>"
                                    'Response.flush
                                    if len(trim(intAktfavgp_use(a))) <> 0 then
								    call tilknytstamakt(a, intAktfavgp_use(a), trim(strAktFase_use(1)), 0)
                                    end if
							
                                next
                                
                                'Response.write "her"
                                'Response.end





                                '*** Opdaterer projektgrupper på Eksisterende aktiviteter ved redigering af job ***'
                                                 '*** Så de følger jobbet. = Sync (nedarv på akt.) ***'
                    								
					                                if request("FM_opdaterprojektgrupper") = 1 then
                                                    
                                                    'Response.write "her"
                                                    'Response.end 
                                                    	
					                                strSQL = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM job WHERE id = "& varJobId 
                    	
					                                oRec5.open strSQL, oConn, 3
					                                if not oRec5.EOF then
                    									
					                                    oConn.execute("UPDATE aktiviteter SET"_
					                                    &" projektgruppe1 = "& oRec5("projektgruppe1") &" , projektgruppe2 = "& oRec5("projektgruppe2") &", "_
					                                    &" projektgruppe3 = "& oRec5("projektgruppe3") &", "_
					                                    &" projektgruppe4 = "& oRec5("projektgruppe4") &", "_
					                                    &" projektgruppe5 = "& oRec5("projektgruppe5") &", "_
					                                    &" projektgruppe6 = "& oRec5("projektgruppe6") &", "_
					                                    &" projektgruppe7 = "& oRec5("projektgruppe7") &", "_
					                                    &" projektgruppe8 = "& oRec5("projektgruppe8") &", "_
					                                    &" projektgruppe9 = "& oRec5("projektgruppe9") &", "_
					                                    &" projektgruppe10 = "& oRec5("projektgruppe10") &" WHERE job = "& varJobId &"")
                		
					                                end if
					                                oRec5.close
                    									
					                                end if
								                   
								
								
							                    '**** Indlæser Underlev grp / Salgsomkostninger **'
                                                call addUlev
							
							
							
							                   '***** Medarbejertyper timebudget ****'
                                               
                                               strSQLmtyp_tpDEL = "DELETE FROM medarbejdertyper_timebudget WHERE jobid = "& varJobId
                                               oConn.execute(strSQLmtyp_tpDEL)

                                               'Response.write request("FM_mtype_id") & "<hr>"
                                               
                                               mtype_id_arr = split(request("FM_mtype_id"), ",")
                                               


                                               for t = 0 to UBOUND(mtype_id_arr) 
                                                        
                                                        
                                                        tpy_id = trim(mtype_id_arr(t))

                                                        if len(trim(request("FM_mtype_timer_"& tpy_id &""))) <> 0 then
                                                        tpy_timer = request("FM_mtype_timer_"& tpy_id &"")
                                                        else
                                                        tpy_timer = 0
                                                        end if
                                                        
                                                        tpy_timer = replace(tpy_timer, ".", "")
                                                        tpy_timer = replace(tpy_timer, ",", ".") 

                                                        call erDetInt(tpy_timer)
                                                        if isint = 0 then
                                                        tpy_timer = tpy_timer
                                                        else
                                                        tpy_timer = 0
                                                        end if

                                                        '*** Nulstiller ikke ISINT da alle felter skal være iorden før record indlæses **'

                                                        if len(trim(request("FM_mtype_timepris_"& tpy_id &""))) <> 0 then
                                                        tpy_timepris = request("FM_mtype_timepris_"& tpy_id &"")
                                                        else
                                                        tpy_timepris = 0
                                                        end if
                                                        
                                                        tpy_timepris = replace(tpy_timepris, ".", "")
                                                        tpy_timepris = replace(tpy_timepris, ",", ".") 


                                                        call erDetInt(tpy_timepris)
                                                        if isint = 0 then
                                                        tpy_timepris = tpy_timepris
                                                        else
                                                        tpy_timepris = 0
                                                        end if


                                                        if len(trim(request("FM_mtype_kostpris_"& tpy_id &""))) <> 0 then
                                                        tpy_kostpris = request("FM_mtype_kostpris_"& tpy_id &"")
                                                        else
                                                        tpy_kostpris = 0
                                                        end if
                                                        
                                                        tpy_kostpris = replace(tpy_kostpris, ".", "")
                                                        tpy_kostpris = replace(tpy_kostpris, ",", ".") 


                                                        call erDetInt(tpy_timepris)
                                                        if isint = 0 then
                                                        tpy_timepris = tpy_timepris
                                                        else
                                                        tpy_timepris = 0
                                                        end if


                                                        if len(trim(request("FM_mtype_faktor_"& tpy_id &""))) <> 0 then
                                                        tpy_faktor = request("FM_mtype_faktor_"& tpy_id &"")
                                                        else
                                                        tpy_faktor = 0
                                                        end if
                                                        
                                                        tpy_faktor = replace(tpy_faktor, ".", "")
                                                        tpy_faktor = replace(tpy_faktor, ",", ".") 


                                                         call erDetInt(tpy_faktor)
                                                        if isint = 0 then
                                                        tpy_faktor = tpy_faktor
                                                        else
                                                        tpy_faktor = 0
                                                        end if


                                                        if len(trim(request("FM_mtype_belob_"& tpy_id &""))) <> 0 then
                                                        tpy_belob = request("FM_mtype_belob_"& tpy_id &"")
                                                        else
                                                        tpy_belob = 0
                                                        end if
                                                        
                                                        tpy_belob = replace(tpy_belob, ".", "")
                                                        tpy_belob = replace(tpy_belob, ",", ".") 

                                                        call erDetInt(tpy_belob)
                                                        if isint = 0 then
                                                        tpy_belob = tpy_belob
                                                        else
                                                        tpy_belob = 0
                                                        end if

                                                        if len(trim(request("FM_mtype_belob_ff_"& tpy_id &""))) <> 0 then
                                                        tpy_belob_ff = request("FM_mtype_belob_ff_"& tpy_id &"")
                                                        else
                                                        tpy_belob_ff = 0
                                                        end if
                                                        
                                                        tpy_belob_ff = replace(tpy_belob_ff, ".", "")
                                                        tpy_belob_ff = replace(tpy_belob_ff, ",", ".") 

                                                        call erDetInt(tpy_belob_ff)
                                                        if isint = 0 then
                                                        tpy_belob_ff = tpy_belob_ff
                                                        else
                                                        tpy_belob_ff = 0
                                                        end if

                                                        if isInt = 0 then

                                                        strSQLmtyp_tp = "INSERT INTO medarbejdertyper_timebudget "_
                                                        &" (jobid, typeid, timer, timepris, faktor, belob, belob_ff, kostpris) "_
                                                        &" VALUES "_
                                                        &"("& varJobId &", "& tpy_id &", "& tpy_timer &", "& tpy_timepris &", "& tpy_faktor &", "& tpy_belob &", "& tpy_belob_ff &", "& tpy_kostpris &")"


                                                        'Response.write strSQLmtyp_tp & "<br>"
                                                        'Response.flush

                                                        oConn.execute(strSQLmtyp_tp)

                                                        end if

                                               next


                                              
							
							
							end if 
                            'Response.end
							'**** Opret / Rediger Job ***'
				
                '******************************************************************************
                '****** WIP historik **********************************************************

                 'ddDato = year(now) &"/"& month(now) &"/"& day(now) 
             
                 'strSQLUpdjWiphist = "INSERT INTO wip_historik (dato, restestimat, stade_tim_proc, medid, jobid) VALUES ('"& ddDato &"', "& restestimat &", "& stade_tim_proc &", "& session("mid") &", "& id &")"
                 'oConn.Execute(strSQLUpdjWiphist)

                        		

                

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

				            '**** Finder jobansvarlige *****
				            strSQL = "SELECT job.id AS jid, jobnavn, jobnr, jobans1, jobans2, jobans3, jobans4, jobans5, job.beskrivelse, job_internbesk, "_
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
                                    
                                    'Sender notifikations mail
		                            'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            'Mailer.CharSet = 2
                                    'Mailer.FromName = "TimeOut Email Service | " & afsNavn &" "& afsEmail

                                    Set myMail=CreateObject("CDO.Message")
                                    myMail.From="timeout_no_reply@outzource.dk" 'TimeOut Email Service 

                                

		                            
                                
                                    '** Hvis problemer med mail pga spam filtre, kan der ændrs her så der bliver sendet fra en TO eller anden adresse der er godkendt af spamfilter
		                            'select case lto
                                    'case "hestia"
                                    'Mailer.FromAddress = "timeout_no_reply@outzource.dk"
                                    'case else
                                    'Mailer.FromAddress = afsEmail
		                            'end select            


                                    'Mailer.RemoteHost = "webout.smtp.nu" '"195.242.131.254" ' = smtp3.atznet.dk '"webout.smtp.nu" ' '"webmail.abusiness.dk" '"pasmtp.tele.dk"
                                    'Mailer.ContentType = "text/html"


                                     
						            'Mailer.AddRecipient "" & jobAnsThis & "", "" & jobAnsThisEmail & ""
                                    myMail.To= ""& jobAnsThis &"<"& jobAnsThisEmail &">"
                                   

						            'Mailer.Subject = "Til de jobansvarlige på: "& jobnavnThis &" ("& intJobnr &") | " & strkkundenavn  
                                    myMail.Subject= "Til de jobansvarlige på: "& jobnavnThis &" ("& intJobnr &") | " & strkkundenavn  


		                            strBody = "<br>"
                                    strBody = strBody &"<b>Kunde:</b> "& strkkundenavn & "<br>" 
						            strBody = strBody &"<b>Job:</b> "& jobnavnThis &" ("& intJobnr &") <br><br>"

                                    if jobans1 <> "0" AND isNULL(jobans1) <> true then
                                    strBody = strBody &"<b>Jobansvarlig:</b> "& jobans1 &" "& jobans1init &"<br><br>"
                                    end if

                                    if jobans2 <> "0" AND isNULL(jobans2) <> true then
                                    strBody = strBody &"<b>Jobejer:</b> "& jobans2 &" ("& jobans2init &") <br><br>"
		                            end if

                                    if len(trim(strBesk)) <> 0 then
                                    strBody = strBody &"<hr><b>Jobbeskrivelse:</b><br>"
                                    strBody = strBody & strBesk &"<br><br><br><br>"
                                    end if

                                    if len(trim(job_internbesk)) <> 0 then
                                    strBody = strBody &"<hr><b>Intern note:</b><br>"
                                    strBody = strBody & job_internbesk &"<br><br>"
                                    end if

                                    
                                    'strBody = strBody &"<br><br><br>https://outzource.dk/timeout_xp/wwwroot/ver2_10/login.asp?lto="&lto&"&tomobjid="&jobid&"&key="&strLicenskey

                                   'strBody = strBody &"<br><br><br>Gå til TimeOut ved at <a href='https://outzource.dk/timeout_xp/wwwroot/ver2_10/login.asp?lto="&lto&"&tomobjid="&jobid&"'>klikke her..</a>"
                                   '&key="&strLicenskey&"
                                    if jobAnsAnsat = "1" then
                                           if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                           strBody = strBody &"<br><br><br>Gå til TimeOut her:<br>https://outzource.dk/"&lto&"/default.asp?tomobjid="&jobid
                                           else
                                           strBody = strBody &"<br><br><br>Gå til TimeOut her:<br>https://timeout.cloud/"&lto&"/default.asp?tomobjid="&jobid
                                           end if
                                    end if
                                    
                                  


                                    strBody = strBody &"<br><br><br><br><br><br>Med venlig hilsen<br><i>" 
		                            strBody = strBody & session("user") & "</i><br><br>&nbsp;"



                                    select case lto 
                                    case "hestia"
            		                strBody = strBody &"<br><br><br><br>_______________________________________________________________________________________________<br>"
                                    strBody = strBody &"HESTIA Ejendomme, Rosenørnsgade 6, st., 8900 Randers C, Tlf. 70269010 - www.hestia.as<br><br>&nbsp;"
                                    end select
            		


            		
		                            'Mailer.BodyText = strBody
            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing


                                    'Mailer.BodyText = strBody
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
                            '*******************************************************'


    response.Redirect ("jobs.asp?menu=job&func=red&id="&varJobId&"&jobnr_sog="&strJobsog&"&filt="&filt&"&fm_kunde="&strKundeId&"&fm_kunde_sog="&vmenukundefilt&"&showdiv="&showdiv&"&FM_tp_jobaktid="&tp_jobaktid&"&FM_mtype="&mtype&"&visrealtimerdetal="&visrealtimerdetal)



    case "opret", "red"
        %>
        <script src="js/job_2015_jav.js"></script>
        <%call menu_2014 %>

        <%


        if func = "red" then
       



        vlgtmtypgrp = 0
        call mtyperIGrp_fn(vlgtmtypgrp,0) 
        call fn_medarbtyper()


        strSQL = "SELECT id, jobnavn, jobnr, rekvnr, jobstatus, jobans1, jobans2, jobstartdato, jobslutdato, job.beskrivelse, tilbudsnr, "_
            &" projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, "_
            &" projektgruppe9, projektgruppe10, fomr_konto, risiko, preconditions_met, valuta, jfak_moms, jfak_sprog, kundeok, altfakadr, syncslutdato, job.serviceaft "_ 
        &" FROM job, kunder WHERE id = " & id &" AND kunder.Kid = jobknr"

        Response.Write strSQL
	    Response.flush
	
	    oRec.open strSQL, oConn, 3
	
	    if not oRec.EOF then
             
	        strNavn = oRec("jobnavn")
	        strjobnr = oRec("jobnr")
            rekvnr = oRec("rekvnr")
            strStatus = oRec("jobstatus")
            jobans1 = oRec("jobans1")
	        jobans2 = oRec("jobans2")
            strStatus = oRec("jobstatus")
	        strTdato = oRec("jobstartdato")
            strBesk = oRec("beskrivelse")
            strProj_1 = oRec("projektgruppe1")
	        strProj_2 = oRec("projektgruppe2")
	        strProj_3 = oRec("projektgruppe3")
	        strProj_4 = oRec("projektgruppe4")
	        strProj_5 = oRec("projektgruppe5")
	        strProj_6 = oRec("projektgruppe6")
	        strProj_7 = oRec("projektgruppe7")
	        strProj_8 = oRec("projektgruppe8")
	        strProj_9 = oRec("projektgruppe9")
	        strProj_10 = oRec("projektgruppe10")
            fomr_konto = oRec("fomr_konto")
            intprio = oRec("risiko")

            preconditions_met = oRec("preconditions_met")
            valuta = oRec("valuta")


            jfak_moms = oRec("jfak_moms")
            jfak_sprog = oRec("jfak_sprog")

            intkundeok = oRec("kundeok")

            altfakadr = oRec("altfakadr")

            syncslutdato = oRec("syncslutdato")

            intServiceaft = oRec("serviceaft")



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
	       
        
            end if


        oRec.close

        dbfunc = "dbred"

        else 'opr


        select case lto 
        case "fe"
        intprio = 1
        case "dencker"
            if cint(strKundeId) = 1 then
            intprio = -3
            else
            intprio = 100
            end if        
        case else
        intprio = 100
        end select


    
        strjobnr = 0
        strNavn = "Jobnavn.."
	    rekvnr = ""
        jobans1 = 0
	    jobans2 = 0
        strNexttilbudsnr = 0

        dbfunc= "dbopr"

        jobstdato = day(now) & "-" & month(now) & "-" & year(now)
        jobstdato = dateadd("d", -7, jobstdato) 
        jobsldato = dateadd("m", 1, jobstdato)

        fomr_konto = 0


        preconditions_met = 0 '1 = ja, 0= ikke angivet, 2 = nej vent

        intkundeok = 0

        altfakadr = 0

        intServiceaft = 0


         select case lto
        case "epi", "epi_no", "epi_sta", "intranet - local", "epi_ab", "epi_cati", "epi_uk"
	    virksomheds_proc = 50
	    syncslutdato = 1
        case else
        virksomheds_proc = 0
	    syncslutdato = 0
        end select


        sandsynlighed = 0
        jfak_moms = 1
        jfak_sprog = 1
        sdskpriogrp = 0
        valuta = basisValId
   
        'if multibletrue = 0 then
    
        '** er der en SDSK priogruppe, valuta,moms og fak sprog på kunde ****
        'strSQL = "SELECT sdskpriogrp, kfak_moms, kfak_sprog, kfak_valuta FROM kunder WHERE kid = "& strKundeId
        'oRec.open strSQL, oConn, 3
        'if not oRec.EOF then
        'sdskpriogrp = oRec("sdskpriogrp")
        'jfak_moms = oRec("kfak_moms")
        'jfak_sprog = oRec("kfak_sprog")
        'valuta = oRec("kfak_valuta")
        'end if
        'oRec.close
    
        'end if

   
   

        end if 'func red


        if func = "red" then



            if altfakadr = 0 then
            altfakadrCHK = ""
            else
            altfakadrCHK = "CHECKED"
            end if

	
	        if jobfaktype = 0 then
	        jobfaktypeCHK0 = "CHECKED"
	        jobfaktypeCHK1 = ""
	        else
	        jobfaktypeCHK0 = ""
	        jobfaktypeCHK1 = "CHECKED"
	        end if 
	
	else
	    
           


	    if lto = "dencker" then
	        
	        jobfaktypeCHK0 = ""
	        jobfaktypeCHK1 = "CHECKED"
	    
	    else
	    
	        if request.Cookies("tsa")("faktype") <> "" then
    	        
	            if request.Cookies("tsa")("faktype") = "0" then
	            jobfaktypeCHK0 = "CHECKED"
	            jobfaktypeCHK1 = ""
	            else
	            jobfaktypeCHK0 = ""
	            jobfaktypeCHK1 = "CHECKED"
	            end if 
    	        
	        else
    	        
	            jobfaktypeCHK0 = "CHECKED"
	            jobfaktypeCHK1 = ""
    	    
	        end if
	    
	    end if
	
	end if

    if editok = 1 then 

        if cint(intkundeok) = 1 OR cint(intkundeok) = 2 then
	    kundechk = "checked"
	    else
	    kundechk = ""
	    end if


    end if

%>








<div class="wrapper">
    <div class="content">



    
 <!------------------------------- Sideindhold------------------------------------->
 
        <form id="opretproj" action="jobs.asp?func=<%=dbfunc %>" method="post">
        <input type="hidden" name="" id="kundekpersopr" value="<%=kundekpers%>" />
        <input id="showdiv" name="showdiv" value="<%=showdiv%>" type="hidden" />
	    <!--<input type="hidden" name="func" id="func" value="<=dbfunc%>">-->
	    <input type="hidden" name="fm_kunde_sog" value="<%=request("fm_kunde_sog")%>">
	    <input type="hidden" id="rdir" name="rdir" value="<%=request("rdir")%>">
	    <input type="hidden" id="jq_jobid" name="id" value="<%=id%>">
	    <input type="hidden" name="int" id="int" value="<%=strInternt%>">
	    <input type="hidden" name="jobnr_sog" value="<%=request("jobnr_sog")%>">
	    <input type="hidden" name="filt" value="<%=request("filt")%>">
	    <input type="hidden" name="showaspopup" value="<%=showaspopup%>">
        <input type="hidden" id="jq_func" name="jq_func" value="<%=func%>">
        <input type="hidden" id="lto" name="lto" value="<%=lto%>">

        <div class="container">

            <div class="portlet">
                <h3 class="portlet-title"><u>Projekt oprettelse</u></h3>
            </div>

            <div class="portlet-body">
                

                <div class="row">

                    <div class="col-lg-12">
                        <%if func = "opret" then %>
                        <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opret</b></button>
                        <%else %>
                        <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdater</b></button>
                        <%end if %>
                    </div>
                </div>

                <br />

               <div class="well well-white">
                             	           
               <!-- <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse3">
                                Stamdata
                            </a>
                          </h4>
                        </div> 
                        
                        <div id="collapse3" class="panel-collapse collapse in">
                            <div class="panel-body"> !-->
                            <input type="hidden" name="FM_fakturerbart" value="1">
                                
                                 <!-- <div class="row">
                                      
                                      <div class="col-lg-1">Kunde: <span style="color:red">*</span></div>
                                      <div class="col-lg-4">



                                          <select class="form-control input-small" name="FM_kunde" id="FM_kunde">
		                                   
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
                                      </div>
                              
                                                                                           

                                      
                                       <div class="col-lg-1">Start dato:</div>
                                        <div class="col-lg-2">
                                        <div class='input-group date' id='datepicker_stdato'>
                                        <input type="text" class="form-control input-small" name="FM_jobstartdato" value="<%=jobstdato %>" placeholder="dd-mm-yyyy" />
                                            <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                        </div>
                                    </div> 
                                  
                                    <div class="col-lg-1">Rek. nr:</div>
                                    <div class="col-lg-3"><input type="text" name="FM_rekvnr" id="FM_rekvnr" value="<%=rekvnr%>" class="form-control input-small" style="width:75%"></div>
                                       
                                </div>

                                      

                                <div class="row">
                                                                      
                                        <div class="col-lg-1">Navn: <span style="color:red">*</span></div>
                                        <div class="col-lg-4"><input type="text" name="FM_navn" id="FM_navn" value="<%=strNavn%>" class="form-control input-small"></div>
                                                                                                                                    
                                       
                                                                                                        
                                            <div class="col-lg-1">Slut dato:</div>
                                            <div class="col-lg-2">
                                                <div class='input-group date' id='datepicker_stdato'>
                                                <input type="text" class="form-control input-small" name="FM_jobslutdato" value="<%=jobsldato %>" placeholder="dd-mm-yyyy" />
                                                    <span class="input-group-addon input-small">
                                                        <span class="fa fa-calendar">
                                                        </span>
                                                    </span>
                                                </div>
                                            </div>
                                    
                                            <div class="col-lg-1">Status:</div>
                                            <div class="col-lg-3">
                                                <select name="FM_status" id="FM_status" class="form-control input-small" style="width:75%">
									            <%
                                                lkDatoThis = ""
                                                if dbfunc = "dbred" then 
									            select case strStatus
									            case 1
									            strStatusNavn = "Aktiv"
									            case 2
									            strStatusNavn = "Til Fakturering" 'passiv
									            case 0
									            strStatusNavn = "Lukket"
                                                    if cdate(lkdato) <> "01-01-2002" then
                                                    lkDatoThis = " ("& formatdatetime(lkdato, 2) & ")"
                                                    end if
									            case 3
									            strStatusNavn = "Tilbud"
                                                case 4
									            strStatusNavn = "Gennemsyn"
									            end select
									            %>
									            <option value="<%=strStatus%>" SELECTED><%=strStatusNavn%> <%=lkDatoThis %></option>
									            <%end if
                                    
                                  
                                                %>
									            <option value="1">Aktiv</option>
									            <option value="2">Til Fakturering</option> <!-- Passiv 
									            <option value="0">Lukket</option>
                                                <option value="4">Gennemsyn</option>
									
									            <option value="3">Tilbud</option>
									
									            </select> 
                                            </div>
                                    
                                                                  
                                    </div>

                                    

                                <div class="row">
                                     <div class="col-lg-1">Kontaktpers.:</div>
                                    <div class="col-lg-4">
                                      
		                                <select name="FM_kpers" id="FM_kpers" class="form-control input-small">
		                                    <option value="0">Ingen</option>
	
		                                   
		                                </select>
                                    </div>                               
                                    
                                </div>

                                <div class="row">
                                    <div class="col-lg-1">Projekt nr:</div>
                                    <div class="col-lg-4"><input type="text" name="FM_jnr" id="FM_jnr" value="<%=strJobnr%>" class="form-control input-small" ></div>
                                                                           
                                                             
                                </div>



                                <div class="row"> -->
                                        
                                      
            <!------------------------- Der skal en tekst ind her (Se det på den game jobs) ------------------------------------------------>
                               

                                
                                                             
                                
                             <!--   </div> -->

                            


                                <div class="row">

                                    <div class="col-lg-5">

                                        <table class="table">

                                            <tr style="border:hidden">
                                                <td style="color:black">Kunde: <span style="color:red;">*</span></td>
                                                <td>
                                                    <select class="form-control input-small" name="FM_kunde" id="FM_kunde">
		                                   
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
                                                </td>
                                            </tr>

                                            <tr style="border:hidden">
                                                <td style="color:black">Navn: <span style="color:red;">*</span></td>
                                                <td><input type="text" name="FM_navn" id="FM_navn" value="<%=strNavn%>" class="form-control input-small"></td>
                                            </tr>

                                            <tr style="border:hidden">
                                                <td style="color:black">Kontaktpers.:</td>
                                                <td><select name="FM_kpers" id="FM_kpers" class="form-control input-small">
		                                        <option value="0">Ingen</option>
	
		                                   
		                                        </select></td>
                                            </tr>

                                            <tr style="border:hidden">
                                                <td style="color:black">Projekt nr.:</t>
                                                <td><input type="text" name="FM_jnr" id="FM_jnr" value="<%=strJobnr%>" class="form-control input-small" ></td>
                                            </tr>

                                            <%for ja = 1 to 1 
					
						                    select case ja
						                    case 1
						                    jbansImg = "<img src='../ill/ac0019-24.gif' width='24' height='24' alt='Jobansvarlig' border='0'>"
                                            if cint(showSalgsAnv) = 1 then 
						                    jbansTxt = "Jobansv."
                                            else
                                            jbansTxt = "Jobansvarlig"
                                            end if
						                    jobansField = "jobans1, jobans_proc_1"
						                    end select
						                    %>
                                            <tr>
                                                <td style="color:black"><%=jbansTxt %>:</td>
                                                <td>
                                                    <select name="FM_jobans_<%=ja %>" id="FM_jobans_<%=ja %>" class="form-control input-small">
						                            <option value="0">Ingen</option>
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
							                            &" LEFT JOIN job ON (job.id = "& id &") WHERE ((mansat = 1 "& mPassivSQL &") "& mTypeExceptSQL &") OR (mid = jobans"& ja &")"
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
							                                select case ja
						                                    case 1
						                                    usemed = oRec("jobans1")
                                                            jobans_proc = oRec("jobans_proc_1")
						                                    end select
							                            end if
							
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
                                                            opTxt = opTxt & " - Deaktiveret"
                                                            case 3 
                                                            opTxt = opTxt & " - Passiv"
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
                                                </td>
                                            </tr>
                                            <%next %>

                                        </table>

                                    </div>

                                    <div class="col-lg-1">&nbsp</div>

                                    <div class="col-lg-6">

                                        <table class="table">

                                            <tr style="border:hidden">

                                                <td style="color:black">Start d.:</td>
                                                <td>
                                                    <div class='input-group date' id='datepicker_stdato'>
                                                    <input type="text" class="form-control input-small" name="FM_jobstartdato" value="<%=jobstdato %>" placeholder="dd-mm-yyyy" />
                                                        <span class="input-group-addon input-small">
                                                            <span class="fa fa-calendar">
                                                            </span>
                                                        </span>
                                                    </div>
                                                </td>

                                                <td style="color:black">Rek. nr.:</td>
                                                <td>
                                                    <input type="text" name="FM_rekvnr" id="FM_rekvnr" value="<%=rekvnr%>" class="form-control input-small">
                                                </td>
                                                                                           
                                            </tr>

                                            <tr style="border:hidden">

                                                <td style="color:black">Slut d.:</td>
                                                <td>
                                                    <div class='input-group date' id='datepicker_stdato'>
                                                    <input type="text" class="form-control input-small" name="FM_jobslutdato" value="<%=jobsldato %>" placeholder="dd-mm-yyyy" />
                                                    <span class="input-group-addon input-small">
                                                        <span class="fa fa-calendar">
                                                        </span>
                                                    </span>
                                                    </div>
                                                </td>


                                                <td style="color:black">Status:</td>
                                                <td>
                                                    <select name="FM_status" id="FM_status" class="form-control input-small">
									                <%
                                                    lkDatoThis = ""
                                                    if dbfunc = "dbred" then 
									                select case strStatus
									                case 1
									                strStatusNavn = "Aktiv"
									                case 2
									                strStatusNavn = "Til Fakturering" 'passiv
									                case 0
									                strStatusNavn = "Lukket"
                                                        if cdate(lkdato) <> "01-01-2002" then
                                                        lkDatoThis = " ("& formatdatetime(lkdato, 2) & ")"
                                                        end if
									                case 3
									                strStatusNavn = "Tilbud"
                                                    case 4
									                strStatusNavn = "Gennemsyn"
									                end select
									                %>
									                <option value="<%=strStatus%>" SELECTED><%=strStatusNavn%> <%=lkDatoThis %></option>
									                <%end if
                                    
                                  
                                                    %>
									                <option value="1">Aktiv</option>
									                <option value="2">Til Fakturering</option> <!-- Passiv -->
									                <option value="0">Lukket</option>
                                                    <option value="4">Gennemsyn</option>
									
									                <option value="3">Tilbud</option>
									
									                </select>
                                                </td>                                              
                                            </tr>

                                            <tr style="border:hidden">
                                                <td colspan="4"><textarea name="strBesk" id="FM_beskrivelse" class="form-control input-small" rows="6" placeholder="Jobbeskrivelse"></textarea></td>
                                            </tr>

                                        </table>

                                    </div>

                                </div>
                                                       

                               
                                <!--<div class="row">
                                    <div class="col-lg-11"><br />Intern besked: <br />
                                    <textarea id="TextArea1" name="FM_internbesk" cols="70" rows="3" class="form-control input-small"></textarea>   		               
                                    </div>                                
                                </div>-->

                              <!--
                                <br />
                                <div class="row">
                                    <%if func = "red" then %>
                                    <div class="col-lg-2">Tilbudsnr.:</div>
                                    <%else %>
                                    <div class="col-lg-2">Tilbud?</div>
                                    <%end if %>                                   
                                </div>

                                <div class="row">
                                    <%if cint(strStatus) = 3 then
					                chkusetb = "CHECKED"
					                else
					                chkusetb = ""
					                end if
					                %>
                                    <div class="col-lg-3"><input type="checkbox" id="FM_usetilbudsnr" name="FM_usetilbudsnr" value="j" <%=chkusetb%>> Dette job har status som tilbud.</div>
                                 </div>

                                    <div class="row">
                                        <%if func = "red" then      
                                                         
                                         if strNexttilbudsnr <> strtilbudsnr then 
                                         tlbplcholderTxt = strNexttilbudsnr
                                         else
                                        tlbplcholderTxt = ""
                                        end if %>
                                        <div class="col-lg-3">
                                        <span style="color:#999999">(Næste ledige nummer: <%=tlbplcholderTxt %>)</span>
                                        </div>
                                        <div class="col-lg-2"><input type="text" id="FM_tnr" name="FM_tnr" value="<%=strtilbudsnr%>" style="<%=tbStyle%>" size="20" class="form-control input-small"></div>
                                        <%else%>
					                    <input type="hidden" name="FM_tnr" value="0">
					                    <%end if%>

                                        <input type="hidden" id="FM_nexttnr" value="<%=strNexttilbudsnr %>">
                                    </div>

                                <div class="row">
                                    <div class="col-lg-1"><input id="Text1" name="FM_sandsynlighed" value="<%=intSandsynlighed %>" type="text" class="form-control input-small"/></div>
                                    <div class="col-lg-3">% sandsynlighed for at vinde tilbud. <span style="font-size:10px; font-family:arial; color:#999999;">(Pipelineværdi = Brutto oms. - Udgifter lev. * sandsynlighed)</span></div>
                                </div>
                                -->
                                  

                                <!--
                                <div class="row">
                                    <div class="col-lg-2">Projektgruppe:</div>
                                        <div class="col-lg-3">
                                            <select class="form-control input-small">
                                                <option value="1">Rengøring</option>
                                                <option value="2">Udvikler</option>
                                                <option value="3">Kørsel</option>
                                                <option value="4">Opvask</option>
                                                <option value="5">Madlavning</option>
                                            </select>
                                        </div>
                                    -->
                                    
                                
                                    
                                

                                <!--
                                <div class="row">
                                    <div class="col-lg-2"><a href="#">Tilføj grp.</a></div>
                                </div>
                                -->
                        

                      <!--      </div> <!-- /.panel-body -->
                     <!--   </div> <!-- /.panel-collapse -->
                   <!-- </div> <!-- /.panel -->
                  <!-- </div-->

                </div> <!-- Well  -->
                <br />

                <% if func = red then %>
                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse4">
                                Aktiviteter
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapse4" class="panel-collapse collapse">
                            <div class="panel-body">
                            


                            <div class="row">
                                <% java = "Javascript:window.open('../timereg/aktiv.asp?menu=job&func=opret&jobid="&id&"&id=0&jobnavn="&strNavn&"&fb=1&rdir=job3&nomenu=1', '', 'width=1004,height=800,resizable=yes,scrollbars=yes')" %>
                                <div class="col-lg-6">
                                     <div class="btn-group demo-element">
                                        <button type="button" class="btn btn-Tertiary dropdown-toggle" data-toggle="dropdown">Tilføj &nbsp<span class="caret"></span></button>

                                        <ul class="dropdown-menu" role="menu">
                                          <li><a onclick="<%=java %>" href="#">Aktivitet</a></li>
                                          <li><a href="javascript:;">Stam-aktivitet</a></li>                                                                                 
                                        </ul>
                                      </div> <!-- /.btn-gruop -->   
                                </div>                                                        
                            </div>


                             <script src="js/tableedittest.js" type="text/javascript"></script>
                              <table class="table dataTable table-bordered table-hover ui-datatable editabletable">
                                  <thead>
                                      <tr>
                                          <th style="width:20%">Aktivitet</th>
                                          <th style="width:10%">Status</th>
                                          <th style="width:9%">Type</th>
                                          <th style="width:5%">Timer/Stk.</th>
                                          <th style="width:5%">Pris ialt</th>
                                          <th style="width:5%">Start dato</th>
                                          <th style="width:5%">Slut dato</th>
                                          <%if avansproobr = 1 then%>
                                          <th style="width:5%">Start tidspunkt</th>
                                          <th style="width:5%">Slut tidspunkt</th>
                                          <%end if %>
                                          <th style="width:2%">Res</th>
                                          <th style="width:2%">Slet</th>
                                      </tr>
                                  </thead>
                                  
                                  <tbody>
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
	                                    oRec.open strSQLakt, oConn, 3
	                                    while not oRec.EOF  
	        
	                                        select case right(a, 1)
	                                        case 0,2,4,6,8
	                                        bga = "#FFFFFF"
                                            case else
	                                        bga = "#Eff3ff"
	                                        end select 
                                        if lcase(lastFase) <> lcase(oRec("fase")) AND len(trim(oRec("fase"))) <> 0 then

                                        if a <> 0 then %>

                                    <!--  <tr>
                                          <td>&nbsp;</td>
	                                      <td>&nbsp;</td>
	                                      <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                
	            
                                          <td><b><%=formatnumber(lastFaseForkalkTimer,2) %></b> t.</td>
                                          <td>&nbsp;</td>
	                                      <td><b><%=formatnumber(lastFaseSum, 2) & "</b> "& basisValISO_f8 %></td>

                                          <td><b><%=formatnumber(lastFaseRealTimer,2) %> t.</td>
	                                      <td><b><%=formatnumber(lastFaseRealbel, 2) & "</b> "& basisValISO_f8 %></td>
                                      </tr> -->

                                     <!-- <tr>
                                        <td colspan=9 style="padding:2px; font-size:9px;">&nbsp;</td>
                                      </tr> -->
                                      <%
                                        lastFaseForkalkTimer = 0
                                        lastFaseRealTimer = 0
                                        lastFaseRealbel = 0
                                        end if %>

                                      <%lastFaseSum = 0 %>

                                     <!-- <tr><td colspan=9 style="padding:2px; font-size:9px;"><b><%=replace(oRec("fase"), "_", " ")%></b></td></tr> -->
	                                 <%end if %>
                                      
                                      <tr>
                                          <td><a href="#" onclick="Javascript:window.open('../timereg/aktiv.asp?menu=job&func=red&id=<%=oRec("id")%>&jobid=<%=id %>&jobnavn=<%=strNavn%>&rdir=job2&nomenu=1', '', 'width=850,height=600,resizable=yes,scrollbars=yes')" class="rmenu"><%=left(oRec("navn"), 20) %></a></td>
                                          <td>
	                                        <%select case oRec("aktstatus")
	                                        case 1
	                                        aktstat = "Aktiv"
	                                        case 2
	                                        aktstat = "Passiv"
	                                        case else
	                                        aktstat = "Lukket"
	                                        end select
	                                            %>
	                                        <%=aktstat %>
                                          </td>
                                          <%call akttyper(oRec("aty_id"), 1) %>
                                          <td><%=left(akttypenavn, 10) %>.</td>

                                          <%select case oRec("bgr")
                                            case 0
                                            bgr = "ingen"
                                            %>
                                            <td class=lille align=right>-</td>
                                            <%
                                            case 1
                                            bgr = "timer"
                                            %>
                                            <td class=lille align=right><%=formatnumber(oRec("budgettimer"), 2) %> t.</td>
                                            <%
                                            case 2
                                            bgr = "stk."
                                            %>
                                            <td class=lille align=right><%=formatnumber(oRec("antalstk"), 2) %> stk.</td>
                                            <%
                                            end select %>


                                          <td><%=formatnumber(oRec("aktbudget"), 2) %></td> <!---  Viser time pris og ikke pris ialt  ------>
                                          <td><%'formatdatetime(oRec("aktstartdato"), 2)%>
                                          <td><%=formatdatetime(oRec("aktslutdato"), 2)%></td>
                                          <td>&nbsp</td>    
                                          <td><a href="../timereg/aktiv.asp?menu=job&func=slet&jobid=<%=jobid%>&jobnavn=<%=request("jobnavn")%>&id=<%=oRec("id")%>&rdir=<%=rdir%>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a></td>                             
                                      </tr>

                                     <%
	                                    a = a + 1
	                                    lastFaseSum = lastFaseSum + oRec("aktbudgetsum") 

                                        if len(trim(oRec("fase"))) <> 0 then
	                                    lastFase = oRec("fase")
                                        else
                                        lastFase = ""
                                        end if 
	        
           
           
            

	                                    call akttyper2009Prop(oRec("fakturerbar"))
	                                    if cint(aty_real) = 1 OR (lto = "oko" AND oRec("aty_id") = 90) then

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

                                            totSum = totSum + oRec("aktbudgetsum")
	                                        totTimerforkalk = totTimerforkalk + oRec("budgettimer")
                                            totReal = totReal + oRec("realiseret")
                                            totRealbel = totRealbel + oRec("realbelob") 
                                            lastFaseRealTimer = lastFaseRealTimer + oRec("realiseret")
                                            lastFaseRealbel = lastFaseRealbel + oRec("realbelob")
                                            realtimerkost = realtimerkost + oRec("realtimerkost")
                
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

                                        lastFaseForkalkTimer = lastFaseForkalkTimer + oRec("budgettimer")
            

	                                    oRec.movenext
	                                    wend
	                                    oRec.close%>


                                   <!--   <tr bgcolor="#FFFFFF">
                
                                            <td>&nbsp;</td>
	                                    <td>&nbsp;</td>
	                                    <td>&nbsp;</td>
                                        <td>&nbsp;</td>
                
	                                    <td align=right class=lille><b><%=formatnumber(lastFaseForkalkTimer,2) %></b> t.</td>
                                        <td>&nbsp;</td>
	                                    <td align=right class=lille><b><%=formatnumber(lastFaseSum, 2) %></b> <%=basisValISO_f8 %></td>

                                        <td align=right class=lille><b><%=formatnumber(lastFaseRealTimer,2) %></b> t.</td>
	                                    <td align=right class=lille><b><%=formatnumber(lastFaseRealbel, 2) %></b> <%=basisValISO_f8%></td>

                                    </tr> -->

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

              
                                    'strAktiviteterGT = "<tr><td colspan=9>&nbsp</td></tr>"
                                    strAktiviteterGT = strAktiviteterGT &"<tr style=background-color:#f9f9f9><td><b>Ialt:</b></td><td>&nbsp</td><td>&nbsp</td>"
                                    strAktiviteterGT = strAktiviteterGT &"<td align=right class=lille><b>"& formatnumber(totTimerforkalk,2) &" t.</b></td>"
                                    strAktiviteterGT = strAktiviteterGT &"<td align=right class=lille><b>"& formatnumber(totSum, 2) &"</b> "& basisValISO_f8 &"</td>"
                                    'strAktiviteterGT = strAktiviteterGT &"<td align=right class=lille><b>"& formatnumber(totReal,2) &" t.</td>"
                                    'strAktiviteterGT = strAktiviteterGT &"<td align=right class=lille><b>"& formatnumber(totRealbel, 2) &"</b> "& basisValISO_f8 &"</td>"
                                    strAktiviteterGT = strAktiviteterGT &"<td colspan=4>&nbsp</td>"
                                    strAktiviteterGT = strAktiviteterGT &"</tr>"



                                    select case lto
                                       case "xoko"
                    
                                        case else    
                                    %>
               
                                    <%=strAktiviteterGT %>
              
             
                                    <%end select %>

                                  </tbody>
                              </table>
                                <br />
                                       
                                                                                                                                                                                                                                                                                    
                                                
                                <!-- <a data-toggle="modal" href="#bigModal" class="btn btn-primary demo-element">Large Modal</a> -->

                                <div id="bigModal" class="modal modal-styled fade" tabindex="-1" role="dialog" style="margin-top:100px;">

                                    <div class="modal-dialog modal-lg">

                                      <div class="modal-content">

                                        <div class="modal-header">
                                          <button type="button" class="close" data-dismiss="modal" aria-hidden="true">×</button>
                                          <h3 class="modal-title" >Tilføj aktiviteter</h3>
                                        </div> <!-- /.modal-header -->

                                        <div class="modal-body">
                                                                                       
                                        </div> <!-- /.modal-body -->

                                        <div class="modal-footer">
                                          <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                                          <button type="button" class="btn btn-primary">Save changes</button>
                                        </div> <!-- /.modal-footer -->

                                      </div><!-- /.modal-content -->

                                    </div><!-- /.modal-dialog -->

                                  </div> <!-- /.modal -->


                            
                            <!-- <div class="row">
                                 <div class="col-lg-3"><a onclick="<%=java %>" class="btn btn-default btn-sm"><b>Tilføj aktivitet</b></a></div>
                             </div>
                            <div class="row">
                                 <div class="col-lg-3"><button type="submit" class="btn btn-default btn-sm"><b>Tilføj stam akt. gruppe</b></button></div>
                             </div> -->


                             
                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div>
                </div>
                <%end if %>

                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#progrp">
                                Projektgrupper <span style="font-size:9px; font-style:normal; font-weight:normal; line-height:11px;">- Hvem skal kunne registrere tid på dette job?</span>
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="progrp" class="panel-collapse collapse">
                            <div class="panel-body">

                             
                                <div class="row">
                                    <div class="col-lg-5">
                                    <table class="table table-stribed">
                                   
                                        <%
							            p = 0
							            for p = 1 to 5
							            varSelected = ""
							            select case p
							            case 1
							            strProj = strProj_1
							            case 2
							            strProj = strProj_2
							            case 3
							            strProj = strProj_3
							            case 4
							            strProj = strProj_4
							            case 5
							            strProj = strProj_5
							            case 6
							            strProj = strProj_6
							            case 7
							            strProj = strProj_7
							            case 8
							            strProj = strProj_8
							            case 9
							            strProj = strProj_9
							            case 10
							            strProj = strProj_10
							            end select
							           %>


                                        <tr style="border:hidden">
                                            <td style="color:black">Projektgruppe <%=p%>:</td>
                                            <td>
                                                <select name="FM_projektgruppe_<%=p%>" class="form-control input-small">
								                <%
									                strSQL = "SELECT id, navn FROM projektgrupper ORDER BY navn"
									                oRec.open strSQL, oConn, 3
									
									                While not oRec.EOF 
									                projgId = oRec("id")
									                projgNavn = oRec("navn")
									
                                                    'if projgNavn = "Alle" then
                                                    'projgNavn = "Alle-gruppen (alle medarbejdere)"
                                                    'end if

									                if cint(strProj) = cint(projgId) then
									                varSelected = "SELECTED"
									                gp = strProj
									                else
									                varSelected = ""
									                gp = gp
									                end if
									                %>
									                <option value="<%=projgId%>" <%=varSelected%>><%=projgNavn%></option>
									                <%
									                oRec.movenext
									                wend
									                oRec.close%>
						                        </select>
                                            </td>
                                        </tr>

                                       <%next %>

                                    </table>
                                    </div>

                                    <div class="col-lg-2">&nbsp</div>

                                    <div class="col-lg-5">
                                        <table class="table">
                                   
                                            <%
							                p = 5
							                for p = 6 to 10
							                varSelected = ""
							                select case p
							                case 1
							                strProj = strProj_1
							                case 2
							                strProj = strProj_2
							                case 3
							                strProj = strProj_3
							                case 4
							                strProj = strProj_4
							                case 5
							                strProj = strProj_5
							                case 6
							                strProj = strProj_6
							                case 7
							                strProj = strProj_7
							                case 8
							                strProj = strProj_8
							                case 9
							                strProj = strProj_9
							                case 10
							                strProj = strProj_10
							                end select
							               %>


                                                <tr style="border:hidden">
                                                    <td style="color:black">Projektgruppe <%=p%>:</td>
                                                    <td>
                                                        <select name="FM_projektgruppe_<%=p%>" class="form-control input-small">
								                        <%
									                    strSQL = "SELECT id, navn FROM projektgrupper ORDER BY navn"
									                    oRec.open strSQL, oConn, 3
									
									                    While not oRec.EOF 
									                    projgId = oRec("id")
									                    projgNavn = oRec("navn")
									
                                                        'if projgNavn = "Alle" then
                                                        'projgNavn = "Alle-gruppen (alle medarbejdere)"
                                                        'end if

									                    if cint(strProj) = cint(projgId) then
									                    varSelected = "SELECTED"
									                    gp = strProj
									                    else
									                    varSelected = ""
									                    gp = gp
									                    end if
									                    %>
									                    <option value="<%=projgId%>" <%=varSelected%>><%=projgNavn%></option>
									                    <%
									                    oRec.movenext
									                    wend
									                    oRec.close%>
						                            </select>
                                                </td>
                                            </tr>

                                           <%next %>

                                        </table>
                                    </div>

                                </div>


                                <%if func = "red" then %>


                                    <% if lto = "execon" OR lto = "immenso" OR lto = "synergi1" then
                                        syncAktProjGrpCHK = "CHECKED"
                                        else
                                        syncAktProjGrpCHK = ""
                                        end if %>


                                    <div class="row">
                                        <div class="col-lg-7"><input type="checkbox" name="FM_opdaterprojektgrupper" id="FM_opdaterprojektgrupper" value="1" <%=syncAktProjGrpCHK %>> Overfør <a data-toggle="modal" href="#modaloverfor"><span class="fa fa-info-circle"></span></a> </div>
                                    </div>
                                
                                <%end if %>



                                     <%
                           if lto <> "execon" then%>
                                <div class="row">
                                    <div class="col-lg-7">
								        <input type="checkbox" name="FM_gemsomdefault" id="FM_gemsomdefault" value="1"> Skift standard forvalgt projektgruppe <a data-toggle="modal" href="#styledModalSstGrp22"><span class="fa fa-info-circle"></span></a> <!--til den gruppe der her vælges som projektgruppe 1.
								        <span style="color:#999999;">Gemmes som cookie i 30 dage.</span> -->
                                    </div>
                                </div>
                               
                                <%end if %>


                            <%if func <> "red" then
                                 
                                    if lto = "jm" OR lto = "synergi1" OR lto = "micmatic" OR lto = "krj" then 'OR lto = "lyng" OR lto = "glad" then
                                    forvalgCHK = "CHECKED"
                                    else
                                    forvalgCHK = ""
                                    end if
                                 
                                 else
                                 
                                 forvalgCHK = ""

                                 end if %>

                            <div class="row">
                                <div class="col-lg-2"><input type="checkbox" name="FM_forvalgt" id="FM_forvalgt" <%=forvalgCHK %> value="1"> Tilføj "Push" <a data-toggle="modal" href="#modaltilfolpush"><span class="fa fa-info-circle"></span></a></div>
                            </div>
								
                               
                              

                                <%if func <> "red" then
                                 
                                    if lto = "jm" OR lto = "synergi1" OR lto = "micmatic" then 'OR lto = "lyng" OR lto = "glad" then
                                    forvalgCHK = "CHECKED"
                                    else
                                    forvalgCHK = ""
                                    end if
                                 
                                 else
                                 
                                 forvalgCHK = ""

                                 end if %>
                                                                                    

                             </div><!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->
                   </div>



                    <div class="panel-group accordion-panel" id="accordion-paneled">
                                <div class="panel panel-default">
                                    <div class="panel-heading">
                                        <h4 class="panel-title">
                                        <a class="accordion-toggle" data-toggle="collapse" data-target="#visansvarlig">
                                            Job & Salgs medansvarlige
                                        </a>
                                        </h4>
                                    </div> <!-- /.panel-heading -->
                                  <div id="visansvarlig" class="panel-collapse collapse">
                                     <div class="panel-body">
                                        <div class="row">
                                            <div class="col-lg-2"><b>Jobansvarlige:</b></div>
                                        </div>
                                        <%for ja = 2 to 5 
					
						                    select case ja
						                    case 1
						                    jbansImg = "<img src='../ill/ac0019-24.gif' width='24' height='24' alt='Jobansvarlig' border='0'>"
                                            if cint(showSalgsAnv) = 1 then 
						                    jbansTxt = "Jobansv."
                                            else
                                            jbansTxt = "Jobansvarlig"
                                            end if
						                    jobansField = "jobans1, jobans_proc_1"
						                    case 2
						                    'jbansImg = "<img src='../ill/ac0020-24.gif' width='24' height='24' alt='Jobejer' border='0'>"
						                    jbansTxt = "Jobejer"
						                    jobansField = "jobans2, jobans_proc_2"
						                    case 3
						                    'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						                    jbansTxt = "Medansv. 1"
						                    jobansField = "jobans3, jobans_proc_3"
						                    case 4
						                    'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						                    jbansTxt = "Medansv. 2"
						                    jobansField = "jobans4, jobans_proc_4"
						                    case 5
						                    'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						                    jbansTxt = "Medansv. 3"
						                    jobansField = "jobans5, jobans_proc_5"
						                    end select
						        
						                    %>

                                            <div class="row">
                                                <div class="col-lg-2"><%=jbansTxt %>:</div>
                                                <div class="col-lg-3">

                                                    <!-- data model - se gl. joboprettelse
                                                    tabel: medarbejdere
                                                    felter: id, navn
                                                    -->
                                        
                                       
                                                    <select name="FM_jobans_<%=ja %>" id="FM_jobans_<%=ja %>" class="form-control input-small">
						                                <option value="0">Ingen</option>
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
							                                &" LEFT JOIN job ON (job.id = "& id &") WHERE ((mansat = 1 "& mPassivSQL &") "& mTypeExceptSQL &") OR (mid = jobans"& ja &")"
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
                                                                opTxt = opTxt & " - Deaktiveret"
                                                                case 3 
                                                                opTxt = opTxt & " - Passiv"
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
                                              <div class="col-lg-1"><input id="FM_jobans_proc_<%=ja %>" name="FM_jobans_proc_<%=ja %>" value="<%=formatnumber(jobans_proc, 1) %>" type="text" class="form-control input-small" /></div>
                                              <!--  <div class="col-lg-1"><a href="#"><span class="glyphicon glyphicon-plus"></span></a></div> -->
                                            </div>
                                            <%Next %>

                                            <br />
                                            <div class="row">
                                                <div class="col-lg-2"><b>Salgsansvarlige:</b></div>
                                            </div>
                                            <%for sa = 1 to 5 
					
						                    select case sa
						                    case 1
						                    'jbansImg = "<img src='../ill/ac0019-24.gif' width='24' height='24' alt='Jobansvarlig' border='0'>"
						                    saansTxt = "Salgsansv. 1"
						                    salgsansField = "salgsans1, salgsans1_proc"
						                    case 2
						                    'jbansImg = "<img src='../ill/ac0020-24.gif' width='24' height='24' alt='Jobejer' border='0'>"
						                    saansTxt = "Salgsansv. 2"
						                    salgsansField = "salgsans2, salgsans2_proc"
						                    case 3
						                    'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						                    saansTxt = "Salgsansv. 3"
						                    salgsansField = "salgsans3, salgsans3_proc"
						                    case 4
						                    'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						                    saansTxt = "Salgsansv. 4"
						                    salgsansField = "salgsans4, salgsans4_proc"
						                    case 5
						                    'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						                    saansTxt = "Salgsansv. 5"
						                    salgsansField = "salgsans5, salgsans5_proc"
						                    end select
						
						                    %>
                                            <div class="row">
                                                <div class="col-lg-2"><%=saansTxt %></div>
                                                <div class="col-lg-3">
                                                    <select name="FM_salgsans_<%=sa %>" id="Select4" class="form-control input-small">
						                                <option value="0">Ingen</option>
							                                <%
							
							                                if func <> "red" then
							                                strSQL = "SELECT mnavn, mnr, mid, init, mansat FROM medarbejdere WHERE (mansat = 1 " & mTypeExceptSQL &")"
							                                else
							                                strSQL = "SELECT mnavn, mnr, mid, mansat, "& salgsansField &", init FROM medarbejdere "_
							                                &" LEFT JOIN job ON (job.id = "& id &") WHERE (mansat = 1 "& mTypeExceptSQL &") OR (mid = salgsans"& sa &")"
							                                end if

                                                            strSQL = strSQL & " ORDER BY mnavn"

							
                                
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
                                                                opTxt = opTxt & " - Deaktiveret"
                                                                case 3 
                                                                opTxt = opTxt & " - Passiv"
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
                                                <div class="col-lg-1"><input id="FM_salgsans_proc_<%=sa %>" name="FM_salgsans_proc_<%=sa %>" value="<%=formatnumber(salgsans_proc, 1) %>" type="text" class="form-control input-small" /></div>
                                            </div>
                                            <%next %>



                                         <br /><br /><br /><br />



                                         <%if level = 1 then %>
                                         <%if lto <> "epi" OR (lto = "epi" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59 OR thisMid = 1 OR thisMid = 1720)) OR (lto = "epi_no" AND thisMid = 2) OR (lto = "epi_ab" AND thisMid = 2) OR (lto = "epi_sta" AND thisMid = 2) OR (lto = "epi_cati" AND thisMid = 2) OR (lto = "epi_uk" AND thisMid = 2) then%>
                                         <div class="row">
                                             <div class="col-lg-5">Virksomhedsandel. af salg i %</div>
                                             <div class="col-lg-1"><input type="text" name="FM_virksomheds_proc" value="<%=formatnumber(virksomheds_proc, 0) %>" class="form-control input-small" /></div>
                                             <%else %>
                                            <input type="hidden" name="FM_virksomheds_proc" value="<%=formatnumber(virksomheds_proc, 0) %>"/>
                                             <%end if %>
                                         </div>

                                         <%else %>

                                        <input type="hidden" name="FM_virksomheds_proc" value="<%=formatnumber(virksomheds_proc, 0) %>"/>

                                         <%end if %>



                                         <%if func <> "red" then

                                            select case lto
                                            case "synergi1", "qwert", "hestia"
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
                                         <div class="row">
                                             <div class="col-lg-5">Adviser valgte medarbejdere om at de er tilføjet som jobans. / jobejer.</div>
                                             <div class="col-lg-1"><input type="checkbox" value="1" name="FM_adviser_jobans" <%=advJobansCHK %> /></div>
                                         </div>


                                      </div><!-- /.panel-body -->
                                    </div> <!-- /.panel-collapse -->
                                </div>
                                </div>



                
                    


                
                    <%
                    oprjobtype = 2    
                    if oprjobtype = 2 then %>

                    <%
                       
                    %>

               

                     



                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse5">
                                Avanceret
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapse5" class="panel-collapse collapse in">
                            <div class="panel-body">

                                                                

                            

                                <br /><br />

                               <div id="accordion-help" class="panel-group accordion-simple">
                                <div class="panel">
                                


                                <%if lto <> "execon" AND lto <> "xintranet - local" then %>
                                      <a data-toggle="collapse" data-parent="#accordion-help" href="#faq-general-1">Forretingsområder:</a>
                                      <!-- <br /> <span>Forretningsområder bruges bl.a. til at se tidsforbrug på tværs af kunder og job, og til at se hvilke slags opgaver man bruger sin tid på. 
                                       <br />
                                       Alle aktiviteter på dette job tæller altid med i de forretningsområder der er valgt på jobbet. Specifikke forretningsområder kan vælges på den enkelte aktivitet.</span>
                                       -->
                                  <div id="faq-general-1" class="panel-collapse collapse in">
                                    <div class="panel-body">
                                         <%
                                         ' uTxt = "Forretningsområder bruges bl.a. til at se tidsforbrug på tværs af kunder og job, og til at se hvilke slags opgaver man bruger sin tid på."
                                         ' uWdt = 300
								
								        'call infoUnisport(uWdt, uTxt) 

                                        call fomr_mandatory_fn()

                                        div_tild_forr_Pos = "relative"
                                        div_tild_forr_Lft = "0px"
                                        div_tild_forr_Top = "0px"
                                        div_tild_forr_z = "0"

                                        if cint(fomr_mandatoryOn) = 1 then
                                        div_tild_forr_VZB = "visible"
                                        div_tild_forr_DSP = ""

                                            if func = "opret" AND step = "2" then
                                            div_tild_forr_bdr = "10"
                                            div_tild_forr_pd = "20"
                                            else
                                            div_tild_forr_bdr = "0"
                                            div_tild_forr_pd = "20"
                                            end if

                                        
                                            if func = "opret" AND step = "2" then
                                            div_tild_forr_Pos = "absolute"
                                            div_tild_forr_Lft = "600px"
                                            div_tild_forr_Top = "500px"
                                            div_tild_forr_z = "400000"
                                            end if

                                        else
                                        div_tild_forr_VZB = "hidden"
                                        div_tild_forr_DSP = "none"

                                            div_tild_forr_bdr = "0"
                                            div_tild_forr_pd = "0"

                                        end if
                                        %>
                                       
                                        
                                            

                                           
                                        
                                        <div class="row">
                                            <div class="col-lg-5"><b>Konto:</b></div>
                                        </div>
                                            <div class="row">
                                                <div class="col-lg-12">    
                                                  
                                                    <%

                                                    '** Finder kundetype, til forvalgte forrretningsområder '***
                                                    thisKtype_segment = 0
                                                    if func = "opret" AND step = 2 AND strKundeId <> "" then 


                                                        strSQLktyp = "SELECT ktype FROM kunder WHERE kid = " & strKundeId
                                                        oRec5.open strSQLktyp, oConn, 3
                                                        if not oRec5.EOF then

                                                        thisKtype_segment = oRec5("ktype")

                                                        end if
                                                        oRec5.close


                                                    end if

                                
                                
                                                    'strSQLf = "SELECT id, navn FROM fomr WHERE id <> 0 ORDER BY navn"
                                                        strSQLf = "SELECT f.navn AS fnavn, f.id, f.konto, kp.kontonr AS kkontonr, kp.navn AS kontonavn, fomr_segment FROM fomr AS f "_
                                                        &" LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) WHERE f.id <> 0 AND f.jobok = 1 ORDER BY f.navn"


                                                        'if session("mid") = 1 then
                                                        'response.write strSQLf
                                                        'response.flush
                                                        'end if

                                                        %>
                                                    <select name="FM_fomr" id="FM_fomr" multiple="multiple" size="12" style="width:380px;">
                                                    <option value="0">Ingen valgt</option>
                                    
                                                        <%
                                                        fa = 0
                                                        strchkbox = ""
                                                        oRec.open strSQLf, oConn, 3
                                                        while not oRec.EOF

                                                            if func = "opret" AND step = 2 then '*** Opret (forvalgt)

                                                            if instr(oRec("fomr_segment"), "#"& thisKtype_segment &"#") <> 0 then
                                                            fSel = "SELECTED"
                                                            else
                                                            fSel = ""
                                                            end if


                                                            else '** Rediger Forretningsområder
                                    
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
                                        <br />
                                       <!-- <div class="row">
                                            <div class="col-lg-12"><b>Forvalgt konto på faktura / ERP system</b><br />
                                            Vælg herunder blandt de forretningsområder der har tilknyttet en omsætningskonto, og hvor fakturaer på dette job skal posteres på denne konto:<br />  
                                            <%=strchkbox %></div>
                                        </div>   -->
                                        
                                        <%if func <> "red" then

                                        select case lto
                                        case "hestia", "intranet - local"
                                        fomr_sync_CHK = "CHECKED"
                                        case else
                                        fomr_sync_CHK = ""
                                        end select

                                        else
                                        fomr_sync_CHK = ""
                                        end if

                                        %>                        

                                        </div> <!-- /.panel-body -->
                                    </div> <!-- /.panel-collapse -->
                                </div>
                                </div>

                                <%else 
                
                                fomrIDs = replace("0" & strFomr_rel, "#", ",")
                                fomrIDs = replace(fomrIDs, ",,", ",")
                                fomrIDs = left(fomrIDs, len(fomrIDs)-1)
                                %>
                                <input id="Hidden2" type="hidden" name="FM_fomr" value="<%=fomrIDs%>" />



                                <%end if%>
                                
                                
                                <br />

                                <div id="accordion-help" class="panel-group accordion-simple">
                                <div class="panel">
                                
                                      <a data-toggle="collapse" data-parent="#accordion-help" href="#faq-general-2">Avanceret indstillinger:</a>
                                       <!--<br /> <span>Tildel bla. prioitet, faktura-indstillinger, pre-konditioner, kundeadgang mm.</span>
                                        -->
                                  <div id="faq-general-2" class="panel-collapse collapse in">
                                    <div class="panel-body">
                                          
                                        

                                        <%strSQLr = "SELECT risiko FROM job WHERE risiko > -1 AND jobstatus = 1 ORDER BY risiko DESC"
									 
									     highestRiskval = 1
									     oRec4.open strSQLr, oConn, 3
									     if not oRec4.EOF then
									     highestRiskval = oRec4("risiko")
									     end if
									     oRec4.close
									 
									     %>
									 
									     <%strSQLr = "SELECT risiko FROM job WHERE risiko > -1 AND jobstatus  = 1 ORDER BY risiko"
									 
									     lowestRiskval = 0
									     oRec4.open strSQLr, oConn, 3
									     if not oRec4.EOF then
									     lowestRiskval = oRec4("risiko")
									     end if
									     oRec4.close
									 
									     %>

                                 
                                        <div class="row">
                                            <div class="col-lg-2">Prioitet: <a data-toggle="modal" href="#styledModalSstGrp23"><span class="fa fa-info-circle"></span></a></div>
                                            <div class="col-lg-3"><input id="prio" name="prio" type="text" value="<%=intprio %>" class="form-control input-small" /></div>

                                            <%if func = "opret" AND step = 2 OR func = "red" then %>
                                            <div class="col-lg-1"></div>

                                            <div class="col-lg-6">Hvis job åbnes for kontakt, hvornår skal registrerede timer så være tilgængelige?</div>
                                            <%end if %>

                                        </div>

                                        <%if func = "opret" then
		                                hvchk1 = "checked"
		                                hvchk2 = ""
		                                else
			                                if cint(intkundeok) = 2 then
			                                hvchk1 = ""
			                                hvchk2 = "checked"
			                                else
			                                hvchk1 = "checked"
			                                hvchk2 = ""
			                                end if
		                                end if%>

                                        <div class="row">
                                            <div class="col-lg-2">Pre-konditioner opfyldt:</div>

                                             <%
                                                preconditions_met_SEL0 = "SELECTED"
                                                preconditions_met_SEL1 = ""
                                                preconditions_met_SEL2 = ""

                                                select case cint(preconditions_met)
                                                case 0
                                                preconditions_met_SEL0 = "SELECTED"
                                                case 1
                                                preconditions_met_SEL1 = "SELECTED"
                                                case 2
                                                preconditions_met_SEL2 = "SELECTED"
                                                case else
                                                preconditions_met_SEL0 = "SELECTED"
                                                end select
                                             %>

                                            <div class="col-lg-3">
                                                <select name="FM_preconditions_met" id="Select1" size="1" class="form-control input-small">
                                                    <option value="0" <%=preconditions_met_SEL0 %>>Ikke angivet</option>
                                                    <option value="1" <%=preconditions_met_SEL1 %>>Ja</option>
                                                    <option value="2" <%=preconditions_met_SEL2 %>>Nej - afvent</option>
                                                </select>
                                            </div>

                                            <%if func = "opret" AND step = 2 OR func = "red" then %>
                                            <div class="col-lg-1"></div>

                                            <div class="col-lg-4">
                                                <input type="radio" name="FM_kundese_hv" value="0" <%=hvchk1%>> Offentliggør timer, så snart de er indtastet.
                                            </div>
                                            <%end if %>

                                        </div>
                                        
                                        

                                        <div class="row">
                                            <div class="col-lg-2">Valuta:</div>
                                            <div class="col-lg-3"><select name="FM_valuta" class="form-control input-small">
							                    <%strSQL = "SELECT id, valuta, valutakode, kurs FROM valutaer WHERE id <> 0 ORDER BY id " 
							                    oRec.open strSQL, oConn, 3
							                    while not oRec.EOF 
							                     if oRec("id") = cint(valuta) then
							                     vSEL = "SELECTED"
							                     else
							                     vSEL = ""
							                     end if%>
							                    <option value="<%=oRec("id") %>" <%=vSEL %>><%=oRec("valuta") %> | <%=oRec("valutakode") %> | kurs: <%=oRec("kurs")%></option>
							                    <%
							                    oRec.movenext
							                    wend
							                    oRec.close%>
							
							                    </select>
                                            </div>

                                            <%if func = "opret" AND step = 2 OR func = "red" then %>
                                            <div class="col-lg-1"></div>

                                            <div class="col-lg-5">
                                                <input type="radio" name="FM_kundese_hv" value="1" <%=hvchk2%>> Offentliggør først timer når jobbet er lukket. (afsluttet/godkendt)
                                            </div>
                                            <%end if %>

                                        </div>

                                        <div class="row">
                                            <div class="col-lg-2">Moms:</div>
                                            <div class="col-lg-3">
                                                <%strSQLmoms = "SELECT id, moms FROM fak_moms WHERE id <> 0 ORDER BY id " %>
                                                <select name="FM_jfak_moms" class="form-control input-small">

                                                    <%oRec6.open strSQLmoms, oConn, 3
                                                    while not oRec6.EOF 

                                                        if cint(jfak_moms) = cint(oRec6("id")) then
                                                        fakmomsSeL = "SELECTED"
                                                        else
                                                        fakmomsSeL = ""
                                                        end if

                                                    %><option value="<%=oRec6("id") %>" <%=fakmomsSeL %>><%=oRec6("moms") %>%</option><%
                  
                                                    oRec6.movenext
                                                    wend 
                                                    oRec6.close%>

                                                </select>
                                            </div>
                                            
                                            
                                        </div>

                                        <div class="row">
                                            <div class="col-lg-2">Sprog:</div>                                         
                                            <div class="col-lg-3">
                                                <%strSQLsprog = "SELECT id, navn FROM fak_sprog WHERE id <> 0 ORDER BY id " %>
                                                <select name="FM_jfak_sprog" class="form-control input-small">

                                                    <%oRec6.open strSQLsprog, oConn, 3
                                                    while not oRec6.EOF 

                                                        if cint(jfak_sprog) = cint(oRec6("id")) then
                                                        faksprogSeL = "SELECTED"
                                                        else
                                                        faksprogSeL = ""
                                                        end if

                                                    %><option value="<%=oRec6("id") %>" <%=faksprogSeL %>><%=oRec6("navn") %></option><%
                  
                                                    oRec6.movenext
                                                    wend 
                                                    oRec6.close%>

                                                </select>
                                            </div>                                  

                                        </div>

                                        <br />
                                        <br />


                                        <div class="row">
                                            <div class="col-lg-3">
                                                <%if strAar_slut = 44 then
												chald = "checked"
												else
												chald = ""
												end if%>
												<input type="checkbox" name="FM_datouendelig" value="j" <%=chald%>> Aldrig: (1. jan 2044)
                                            </div>
                                        </div>


                                        <%if syncslutdato = 1 then
                                                syncslutdatoCHK = "CHECKED"
                                                else
                                                syncslutdatoCHK = ""
                                         end if %>

                                        <div class="row">
                                            <div class="col-lg-5">
                                                <input type="checkbox" name="FM_syncslutdato" value="j" <%=syncslutdatoCHK %>> Sync. slutdato, så den følger sidste timereg. / sidste faktura / dd. <!-- (ved luk job) -->
                                            </div>
                                          
                                            <div class="col-lg-1"></div>

                                            <%if func = "opret" AND step = 2 OR func = "red" then %>
                                        

                                            <div class="col-lg-5"><input type="checkbox" name="FM_kundese" id="FM_kundese" value="1" <%=kundechk%>>&nbsp; Gør job tilgængeligt for kontakt.</div>
                                            
                                            <%end if %>
                                           

                                        </div>

                                        <div class="row">
                                            <div class="col-lg-5"><input type="checkbox" name="FM_syncaktdatoer" value="j"> Sync. start- og slut- datoer på aktiviteter, så de nedarver fra job.<br /></div>
                                            
                                            <div class="col-lg-1"></div>

                                            <div class="col-lg-5"><input type="checkbox" value="1" name="FM_brugaltfakadr" <%=altfakadrCHK %> /> Brug kontaktperson / filial som modtager på faktura.</div>
                                        </div>

                                        
                                                   
                                       <br />

                                            

                                        <br /><br />

                                        <div class="row">
                                            <div class="col-lg-2">Tilknyt job til aftale?</div>                                    
                                            <div class="col-lg-3">
                                               
		                                        
                                            </div>
                                                                                                                                                                                                    
                                        </div>
                                    

                                        <br /><br />

                                        <div class="row">
                                            <div class="col-lg-5">
                                            <textarea id="TextArea1" name="" class="form-control input-small" rows="4" placeholder="Intern note"></textarea> 
                                            </div>                                
                                        </div>


                                     
                                        </div> <!-- /.panel-body -->
                                    </div> <!-- /.panel-collapse -->
                                </div>
                                </div>




                                                                 


                   <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse7">
                                Økonomi
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapse7" class="panel-collapse collapse">
                            <div class="panel-body">
                            
                                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                    <thead>
                                        <tr style="font-size:75%">
                                            <th style="border-right:hidden">Bruttoomsætning<br /><span style="font-size:75%">Nettooms. + Salgsomk.</span></th>
                                            <th colspan="4" style="text-align:right; border-right:hidden">= 5100</th>
                                            <th>&nbsp</th>
                                        </tr>

                                        <tr style="font-size:75%">
                                            <th>Nettoomkostning, timer <br /><span style="font-size:75%">Oms. før salgsomk.	</span></th>
                                            <th style="width:10%">Timer</th>
                                            <th style="width:10%">Timepris</th>
                                            <th style="width:10%">Faktor</th>
                                            <th style="width:10%">Beløb</th>
                                            <th>Salgs<br />timepris</th>
                                        </tr>
                                    </thead>

                                    <tbody>
                                        <tr style="font-size:75%">
                                            <td>Gns. timepris / kostpris: 0,00 / 305,56</td>
                                            <td><input class="form-control input-small" type="number" name="FM_exch" value="23"/></td>
                                            <td><input class="form-control input-small" type="number" name="FM_exch" value="31"/></td>
                                            <td><input class="form-control input-small" type="number" name="FM_exch" value="0"/></td>
                                            <td><input class="form-control input-small" type="number" name="FM_exch" value=""/></td>
                                            <td><input class="form-control input-small" type="number" name="FM_exch" value="82,00"/></td>
                                        </tr>
                                    </tbody>
                                  
                                </table>

                                <br /><br />

                                <div id="accordion-help" class="panel-group accordion-simple">
                                <div class="panel">
                                
                                      <a data-toggle="collapse" data-parent="#accordion-help" href="#fordeltimbu">Fordel timebudget på finansår:</a>
                                       
                                  <div id="fordeltimbu" class="panel-collapse collapse">
                                    <div class="panel-body">

                                        <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                            <thead>
                                                <tr style="font-size:75%">
                                                    <th style="width:40%">Budget FY</th>
                                                    <td style="width:10%; text-align:right">Timer</td>
                                                    <td style="width:10%; text-align:right">Budgetår (FY)</td>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                <tr style="font-size:75%">
                                                    <td>År 0</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2016" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>År 1</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2017" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>År 2</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2018" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>År 3</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2019" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>År 4</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2020" style="text-align:right"/></td>
                                                </tr>
                                                
                                            </tbody>
                                        </table>

                                        
                                             
                                    </div> <!-- /.panel-body -->
                                    </div> <!-- /.panel-collapse -->
                                </div>                              
                                </div>

                                
                                <div class="row">                      
                                    <div class="col-lg-2"><b>Jobtype</b></div>
                                    <div class="col-lg-2"><input type="checkbox" name="#" value="#"> <b>Åbn</b> <a data-toggle="modal" href="#styledModalSstGrp26"><span class="fa fa-info-circle"></span></a> <!--for manuel indtastning og beregning af Brutto- og Netto -omsætning.--></div>
                                </div>
                                <div id="styledModalSstGrp26" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                            <div class="modal-dialog">
                                                <div class="modal-content" style="border:none !important;padding:0;">
                                                  <div class="modal-header">
                                                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                      <h5 class="modal-title"></h5>
                                                    </div>
                                                    <div class="modal-body">                                                
                                                        Åbn for manuel indtastning og beregning af Brutto- og Netto -omsætning.
                                                    </div>
                                                </div>
                                            </div>
                                        </div>                 
                                <div class="row">
                                    <div class="col-lg-2"><input type="checkbox" name="#" value="#"> Fastpris <a data-toggle="modal" href="#styledModalSstGrp27"><span class="fa fa-info-circle"></span></a> <!--(bruttoomsætning benyttes ved fakturering)--> </div>
                                </div>
                                 <div id="styledModalSstGrp27" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                    <div class="modal-dialog">
                                        <div class="modal-content" style="border:none !important;padding:0;">
                                            <div class="modal-header">
                                            <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                <h5 class="modal-title"></h5>
                                            </div>
                                            <div class="modal-body">                                                
                                               <b>Fastpris</b> (bruttoomsætning benyttes ved fakturering)
                                            </div>
                                        </div>
                                    </div>
                                </div>   
                                <div class="row">
                                    <div class="col-lg-2"><input type="checkbox" name="#" value="#"> Lbn. timer <a data-toggle="modal" href="#styledModalSstGrp28"><span class="fa fa-info-circle"></span></a> <!-- (timeforbrug på hver enkelt aktivitet * medarb. timepris benyttes ved fakturering) --> </div>
                                </div>
                                 <div id="styledModalSstGrp28" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                    <div class="modal-dialog">
                                        <div class="modal-content" style="border:none !important;padding:0;">
                                            <div class="modal-header">
                                            <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                <h5 class="modal-title"></h5>
                                            </div>
                                            <div class="modal-body">                                                
                                               <b>Lbn. timer</b> (timeforbrug på hver enkelt aktivitet * medarb. timepris benyttes ved fakturering)
                                            </div>
                                        </div>
                                    </div>
                                </div>   

                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->
                   </div>



                    <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#mater">
                                Materialer
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="mater" class="panel-collapse collapse">
                            <div class="panel-body">
                            
                              <div class="row">
                                  <div class="col-lg-2"><b>Salgsomkostninger</b></div>
                              </div>
                              <div class="row">
                                  <div class="col-lg-2">Antal linjer:</div>
                                  <div class="col-lg-1">
                                      <select class="form-control input-small">
                                            <option value="1">1</option>
                                            <option value="2">2</option>
                                            <option value="3">3</option>
                                            <option value="1">4</option>
                                            <option value="2">5</option>
                                            <option value="3">6</option>
                                            <option value="1">7</option>
                                            <option value="2">8</option>
                                            <option value="3">9</option>
                                            <option value="3">10</option>
                                      </select>
                                  </div>
                              </div>

                              <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                  <thead>
                                      <tr style="font-size:75%">
                                          <th style="width:40%">* udgift/salgsomkost.</th>
                                          <th style="width:11%">Stk.</th>
                                          <th style="width:11%">Stk. pris</th>
                                          <th style="width:11%">Indkøbspris</th>
                                          <th style="width:11%">Faktor</th>
                                          <th style="width:11%">Salgspris</th>
                                          <th>&nbsp</th>
                                      </tr>
                                  </thead>
                                  <tbody>
                                      <tr style="font-size:75%">
                                          <td>&nbsp</td>
                                          <td><input class="form-control input-small" type="number" name="FM_exch" value="" placeholder="0,00"/></td>
                                          <td><input class="form-control input-small" type="number" name="FM_exch" value="" placeholder="0,00"/></td>
                                          <td><input class="form-control input-small" type="number" name="FM_exch" value="" placeholder="0,00"/></td>
                                          <td><input class="form-control input-small" type="number" name="FM_exch" value="" placeholder="0,00"/></td>
                                          <td><input class="form-control input-small" type="number" name="FM_exch" value="" placeholder="0,00"/></td>
                                          <td><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>

                                      </tr>
                                  </tbody>
                              </table>

                                <div class="row">
                                    <div class="col-lg-12"><a href="#">Tilføj linje</a>
                                        <a data-toggle="modal" href="#styledModalSstGrp20"><span class="fa fa-info-circle"></span></a>
                                    </div>
                                

                                <div id="styledModalSstGrp20" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                        <div class="modal-dialog">
                                            <div class="modal-content" style="border:none !important;padding:0;">
                                              <div class="modal-header">
                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                  <h5 class="modal-title"><%=tsa_txt_medarb_069 %></h5>
                                                </div>
                                                <div class="modal-body">
                                                <%=tsa_txt_medarb_104 %>

                                                </div>
                                            </div>
                                        </div>
                                 </div></div>
                                
                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->
                   </div>
                 <%end if 'oprjobtype %>





               <br /><br />


        <!---------------------------------- ----------- Modals ------------------------------------------------------>

            <div id="styledModalSstGrp22" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                    <div class="modal-dialog">
                        <div class="modal-content" style="border:none !important;padding:0;">
                            <div class="modal-header">
                                          
                                <h5 class="modal-title">Skift standard forvalgt projektgruppe</h5>
                            </div>
                            <div class="modal-body">
                                Skift standard forvalgt projektgruppe til den gruppe der her vælges som projektgruppe 1. Gemmes som cookie i 30 dage. 
                            </div>
                        </div>
                    </div>
                </div>

            <div id="modaloverfor" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                <div class="modal-dialog">
                    <div class="modal-content" style="border:none !important;padding:0;">
                        <div class="modal-header">
                                          
                            <h5 class="modal-title">Overfør (synkroniser)</h5>
                        </div>
                        <div class="modal-body">
                            Overfør (synkroniser) valgte projektgrupper, til aktiviteterne på dette job.  
                        </div>
                    </div>
                </div>
            </div>

            <div id="modaltilfolpush" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                <div class="modal-dialog">
                    <div class="modal-content" style="border:none !important;padding:0;">
                        <div class="modal-header">
                                          
                            <h5 class="modal-title">Tilføj "Push"</h5>
                        </div>
                        <div class="modal-body">
                            Overfør (synkroniser) valgte projektgrupper, til aktiviteterne på dette job.  
                        </div>
                    </div>
                </div>
            </div>

            <div id="styledModalSstGrp23" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                <div class="modal-dialog">
                    <div class="modal-content" style="border:none !important;padding:0;">
                        <div class="modal-header">
                        <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                            <h5 class="modal-title"></h5>
                        </div>
                        <div class="modal-body">
                                Prioiteter på nuværende aktive job ligger mellem: <b><%=lowestRiskval%> - <%=highestRiskval%></b><br /><br />
						        <b>-1 = Internt job</b> vises ikke under fakturering og igangværende job.<br />
                                <b>-2 = HR job</b> vises i HR mode på timereg. siden<br />
                                <b>-3 = Internt job</b> men der skal kunne laves ressouceforecast på dette job. 
						        <br /><br />
                                -1 / -2 / -3 medfører enkel visning af aktivitetslinjer på timereg. siden, dvs. der bliver ikke vist tidsforbrug, start- og slut -datoer mv. på aktiviteterne. <br /><br />&nbsp;
									
                        </div>
                    </div>
                </div>
            </div>









            </div><!-- /.portlet body -->
            </div><!-- /.container -->
           

        <%end select %>


           </div></div> </form>

        </div><!-- /.wrapper -->
    </div><!-- /.content -->

    

<!--#include file="../inc/regular/footer_inc.asp"-->