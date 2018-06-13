<%response.buffer = true%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../timereg/inc/job_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"--> 

<%
'** Jquery section 
if Request.Form("AjaxUpdateField") = "true" then



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




    function SQLBlessP(s)
	dim tmp2
	tmp2 = s
	tmp2 = replace(tmp2, ",", ".")
	SQLBlessP = tmp2
	end function




    Select case func 
    case "dbopr", "dbred"

                                '********************** Opretter job ***************************


                                kid = request("FM_kunde")
                                kunderef = request("FM_kpers")
                                jobnavn = request("FM_jobnavn")
                                jobnr = request("FM_jobnr")
                                beskrivelse = ""
                                internnote = ""
                                bruttooms = 0
                                editor = session("user")
                                dddato = year(now) & "-" & month(now) & "-" & day(now)
                                jobstartdato = request("FM_jobstartdato")
                                jobstartdato = year(jobstartdato) & "-" & month(jobstartdato) & "-" & day(jobstartdato)

                                jobslutdato = request("FM_jobslutdato")
                                jobslutdato = year(jobslutdato) & "-" & month(jobslutdato) & "-" & day(jobslutdato)

                                jobans1 = request("FM_jobans1")
                                jobans2 = 0

                                rekvisitionsnr = ""


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
                                call visErrorFormat
    		
		                        errortype = 31
		                        call showError(errortype)
		                        response.End
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
                                & " jobslutdato, editor, dato, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, " _
                                & " projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, " _
                                & " fakturerbart, budgettimer, fastpris, kundeok, beskrivelse, " _
                                & " ikkeBudgettimer, tilbudsnr, jobans1," _
                                & " serviceaft, kundekpers, valuta, rekvnr, " _
                                & " risiko, job_internbesk, " _
                                & " jo_bruttooms, fomr_konto) VALUES " _
                                & "('" & jobnavn & "', " _
                                & "'" & strjnr & "', " _
                                & "" & kid & ", " _
                                & "0, " _
                                & "1, " _
                                & "'" & jobstartdato & "', " _
                                & "'" & jobslutdato & "', " _
                                & "'" & editor & "', " _
                                & "'" & dddato & "', " _
                                & "10, " _
                                & "1,1,1,1,1,1,1,1,1," _
                                & "1,0,0,0," _
                                & "'" & beskrivelse & "', " _
                                & "0,0, " _
                                & "" & jobans1 & "," _
                                & "0," & kunderef & ", " _
                                & "1, '" & rekvisitionsnr & "', " _
                                & "100,'" & internnote & "'," _
                                & "" & bruttooms & ", 0)")
    							


                                'response.write "strSQLjob: " & strSQLjob 
                                'response.flush

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


                                
                                '***************************************************************************************************
                                'PUSH - Aktiviteter - Timepriser
                                 '*********** timereg_usejob, så der kan søges fra jobbanken KUN VED OPRET JOB *********************
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





                                else 'hvis rediger



                                'Tjekker om jobnr findes
                                strSQL = "SELECT jobnr, id FROM job WHERE id <> "& jobid &" AND jobnr = '" & jobnr & "'"

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



                                strSQL = "UPDATE job SET "_
                                &" jobnavn = '"& jobnavn &"',"_
                                &" jobnr = '"& jobnr &"', "_
                                &" jobknr = "& kid &", "_
                                &" jobstatus = "& strStatus &""_
                                &" WHERE id = "& jobid

                                response.Write strSQL

                                oConn.execute(strSQL)

                                

                                end if

                                if func = "dbred" then
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
                                    
                                    response.Write "<br> New activity " & newActivitesId(t) & " Name " & newActivityName & " Status " & newActivityStatus &" Budgetantal "& newActivitybudgetantal & " BGR " & newActivityBGR & " Start Date " & newActivitySTdate & " End date " & newActivityENDdate & " PRG " & newActivityPrg & " akttype " & newActivityType                                    

                                    strSQLActivities = "insert into aktiviteter (navn, aktstatus, budgettimer, bgr, aktstartdato, aktslutdato, projektgruppe1, job, fakturerbar)"_
                                    &" VALUES ('"& newActivityName &"', "& newActivityStatus &", "& SQLBlessP(newActivitybudgetantal) &", "& newActivityBGR &", '"& newActivitySTdateSQL &"', '"& newActivityENDdateSQL &"', "& newActivityPrg &", "& jobid &", "& newActivityType &")"
                                    response.Write "<br> sql " & strSQLActivities
                                    oConn.execute(strSQLActivities)

                                    response.Write "<br> New activity " & newActivitesId(t) & " Name " & newActivityName & " Status " & newActivityStatus &" Budgetantal "& newActivitybudgetantal & " BGR " & newActivityBGR & " Start Date " & newActivitySTdate & " End date " & newActivityENDdate & " PRG " & newActivityPrg
                                    

                                next
                                end if
                            
                                '** Rediger
                                'response.redirect("jobs.asp?func=red&id="&id)
                                    

                                '*** Liste
                                'response.redirect("../timereg/jobs.asp")
                                if func = "dbopr" then
                                    strSQLLastJoib = "SELECT id FROM job"
                                    oRec.open strSQLLastJoib, oConn, 3
                                    while not oRec.EOF
                                     
                                    jobid = oRec("id")
                                    
                                    oRec.movenext
                                    wend
                                    oRec.close
                                end if

                                 response.Redirect("jobs.asp?func=red&jobid="&jobid)





    case "opret", "red"
        %>
        <script src="js/job_2018_jav3.js"></script>
        <%call menu_2014 %>

        <%


        if func = "red" then
       
            strSQL = "SELECT id, jobnavn, jobnr, jobknr, kundekpers, jobstartdato, jobslutdato, jobans1, jobans2, jobstatus, tilbudsnr FROM job"_
            & " WHERE id = "& jobid

            'Response.Write strSQL
	        'Response.flush
	
	        oRec.open strSQL, oConn, 3
	
	        if not oRec.EOF then

	            strNavn = oRec("jobnavn")
	            strjobnr = oRec("jobnr")
	            'strKnavn = oRec("kkundenavn")
	            strKnr = oRec("jobknr")
                'kundekpers = oRec("kundekpers")
                jobstdato = oRec("jobstartdato")
                jobsldato = oRec("jobslutdato")

                jobans1 = oRec("jobans1")
                jobans2 = oRec("jobans2")

                strStatus = oRec("jobstatus")

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

            fmkunde = 0

            dbfunc= "dbopr"
            submitTxt = "Opret"

            jobstdato = day(now) & "-" & month(now) & "-" & year(now)
            'jobstdato = dateadd("d", -7, jobstdato) 
            jobsldato = dateadd("m", 1, jobstdato)

            jobans1 = 0
            jobans2 = 0


        end if ' func red
%>










<div class="wrapper">
    <div class="content">



    
 <!------------------------------- Sideindhold------------------------------------->
 
        <form id="opretproj" action="jobs.asp?func=<%=dbfunc %>&jobid=<%=jobid %>" method="post">
        <input type="hidden" name="" id="kundekpersopr" value="<%=kundekpers%>" />
        <div class="container" style="width:1500px;">

            <div class="portlet">
                <h3 class="portlet-title"><u>Projekt oprettelse</u></h3>
            </div>

            <div class="portlet-body">

                            	           
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
                            
                                
                                  <div class="row">
                                  
                                      <div class="col-lg-1">Kunde: <span style="color:red">*</span></div>
                                      <div class="col-lg-5">



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
                              
                                 
                            
                                    <div class="col-lg-2">Kontaktpers. (deres ref.):</div>
                                    <div class="col-lg-3">
                                      
		                                <select name="FM_kpers" id="FM_kpers" class="form-control input-small">
		                                    <option value="0">Ingen</option>
	
		                                   
		                                </select>
                                    </div>
                   
                            
                                  
                                </div>

                                <br /> <br />

                       

                                <div class="row">
                                                                      
                                        <div class="col-lg-1">Navn: <span style="color:red">*</span></div>
                                        <div class="col-lg-5"><input type="text" name="FM_jobnavn" id="FM_jobnavn" value="<%=strNavn%>" class="form-control input-small" placeholder="Projekt Navn"></div>

                                        

                                        <div class="col-lg-2">Job nr:</div>
                                        <div class="col-lg-2"><input class="form-control input-small" type="text" name="FM_jobnr" value="<%=strjobnr %>" /></div>


                                </div>

                                <div class="row">
                                        
                                    <div class="col-lg-1">Status:</div>
                                    <div class="col-lg-2">
                                        <select class="form-control input-small" id="FM_status" name="FM_status">
                                            <option value="1">Aktiv</option>
                                            <option value="2">Til fakturering</option>
                                            <option value="3">Lukket</option>
                                            <option value="4">Gennemsyn</option>
                                            <option value="5">Tilbud</option>
                                        </select>
                                    </div>

                                    <div class="col-lg-2">
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

					                    <input type="text" id="FM_tnr" class="tilbudsinfo form-control input-small" name="FM_tnr" value="<%=strtilbudsnr%>" style="visibility:<%=tilbudvisibility%>;"> 
                  
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
                                     </div>

                                    <div class="col-lg-5 tilbudsinfo" style="visibility:<%=tilbudvisibility%>;">
                                        <input id="Text1" name="FM_sandsynlighed" class="form-control input-small" value="<%=intSandsynlighed %>" type="text" style="width:30px; <%=sandBdr%>; display:inline-block;" />
                                        <span style="display:inline-block;"><%="% "& job_txt_066 %></span>
                                    </div>
                                </div>

                                <div class="row">
                                    <div class="col-lg-1">Ansvarlig:</div>
                                    <div class="col-lg-2">
                                        <select class="form-control input-small" name="FM_jobans1" id="FM_jobans" >
						                    <option value="0">Ingen</option>

                                            <%

                                                select case lto 
                                                case "epi", "epi_ab", "epi_no", "epi_sta", "intranet - local", "epi_uk"
                                                mTypeExceptSQL = " AND (medarbejdertype <> 14 AND medarbejdertype <> 24)"
                                                case else
                                                mTypeExceptSQL = ""         
                                                end select


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

							                <!----- Jobs linje 5910 --->
						                </select>
                                    </div>
                                </div>

                                <div class="row">   

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
                                </div>

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
                                    
                                   <div class="row">
                                    <div class="col-lg-11"><br />Beskrivelse: <br />
                                    <textarea id="TextArea1" name="FM_beskrivelse" cols="70" rows="3" class="form-control input-small"></textarea>   
		               
                                    </div>
                                    
                                </div>

                                <!--
                                <div class="row">
                                    <div class="col-lg-2"><a href="#">Tilføj grp.</a></div>
                                </div>
                                -->

                                    <div class="row">
                                   <div class="col-lg-12 pad-t20">
                                                        <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=submitTxt %></b></button>
                                                 </div>
                                   </div><!-- /.row -->

                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->
                   </div>


                  <!-- Aktiviteter -->

                  <style>
                        .aktView2 {
                            display:none;
                        }
                        .aktView3 {
                            display:none;
                        }

                        .edit_akt_btn:hover,
                        .edit_akt_btn:focus {
                        text-decoration: none;
                        cursor: pointer;
                        }

                  </style>

                  <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">                        
                        <div class="panel-heading">
                        <h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapseAkt">Aktiviteter</a></h4></div> <!-- /.panel-heading -->
                        <div id="collapseAkt" class="panel-collapse collapse in">
                            <div class="panel-body">
                                
                                <div class="row">
                                    <div class="col-lg-6">
                                        <a id="aktView1_btn" class="btn btn-default btn-sm active"><b>Oversigt</b></a>
                                        <a id="aktView2_btn" class="btn btn-default btn-sm"><b>KPI Økonomi</b></a>
                                        <a id="aktView3_btn" class="btn btn-default btn-sm"><b>KPI Tid</b></a>
                                    </div>
                                </div>

                                <table id="activtyTable" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                    <thead>
                                        <tr>
                                            <th style="text-align:left; width:25%">Aktivitet</th>

                                            <th class="aktView1">Status</th>
                                            <th class="aktView1">Type</th>
                                            <th class="aktView1">Budget</th>
                                            <th class="aktView1">Enhed</th>
                                            <th class="aktView1">Start Dato</th>
                                            <th class="aktView1">Slut Dato</th>
                                            <th class="aktView1" style="width:130px;">Medarbejder / grp.</th>
                                            <th class="aktView1">Allokér</th>
                                            <th class="aktView1"></th>

                                            <th class="aktView2">Status</th>
                                            <th class="aktView2">Budget</th>
                                            <th class="aktView2">Enhed</th>
                                            <th class="aktView2">Pris/Enh</th>
                                            <th class="aktView2">I alt</th>
                                            <th class="aktView2">FC (tid)</th>
                                            <th class="aktView2">FC (dkk)</th>
                                            <th class="aktView2">Real (tid)</th>
                                            <th class="aktView2">Real (dkk)</th>

                                            <th class="aktView3">Budget</th>
                                            <th class="aktView3">Enhed</th>
                                            <th class="aktView3">Pris/Enh</th>
                                            <th class="aktView3">FC</th>
                                            <th class="aktView3">Real</th>
                                            <th class="aktView3">Budget vs FC</th>
                                            <th class="aktView3">Budget vs Real</th>
                                            <th class="aktView3">FC vs Real</th>
                                            <th class="aktView3">Fak tid</th>

                                        </tr>
                                    </thead>

                                    <tbody>

                                        <%
                                        if func <> "opret" then
                                        '********* Henter aktiviteter *********'
                                        strSQLakt = "SELECT a.id as aktid, a.navn, aktstatus, aktstartdato, aktslutdato, projektgruppe1, aktbudget, bgr, budgettimer, fakturerbar, p.navn as prgnavn FROM aktiviteter as a LEFT JOIN projektgrupper as p ON (p.id = projektgruppe1) WHERE job = "& jobid
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


                                        startDateHTML = day(oRec("aktstartdato")) &"-"& month(oRec("aktstartdato")) &"-"& year(oRec("aktstartdato")) 
                                        endDateHTML = day(oRec("aktslutdato")) &"-"& month(oRec("aktslutdato")) &"-"& year(oRec("aktslutdato")) 

                                        %>

                                        <tr>
                                            <td><%=oRec("navn") %></td>

                                            <td class="aktView1">
                                                <select class="form-control input-small" id="aktstatus_<%=oRec("aktid") %>" disabled>
                                                    <option value="1" <%=status1SEL %>><%=job_txt_094 %></option>
                                                    <option value="2" <%=status2SEL %>><%=job_txt_320 %></option>
                                                    <option value="3" <%=status3SEL %>><%=job_txt_246 %></option>
                                                </select>
                                            </td>

                                            <td style="width:75px;">
                                                <%call akttyper2009(1) %>
				                                <select id="FM_fakturerbart_<%=oRec("aktid") %>" class="form-control input-small" name="FM_fakturerbart" disabled>
                                                   <%
                                                   Response.Write aty_options
                                                   %>                
                                                </select>
                                            </td>

                                            <td class="aktView1"><input id="aktbudget_<%=oRec("aktid") %>" class="form-control input-small" value="<%=formatnumber(oRec("budgettimer"),2) %>" readonly /></td>

                                            <td class="aktView1">
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

                                                <select id="aktEnhed_<%=oRec("aktid") %>" class="bgr form-control input-small" disabled>
                                                    <option value=0 <%=bgrSEL0 %>>Ingen</option>
                                                    <option value=1 <%=bgrSEL1 %>>Timer</option>
                                                    <option value=2 <%=bgrSEL2 %>>Stk.</option>
                                                </select>

                                            </td>

                                            <td class="aktView1" style="width:100px;">
                                                <div id="aktSTdate_<%=oRec("aktid") %>" class='input-group'>
                                                  <input type="text" style="width:100%;" class="form-control input-small" id="aktInputSTdate_<%=oRec("aktid") %>" value="<%=startDateHTML %>" placeholder="dd-mm-yyyy" readonly />
                                                    <span class="input-group-addon input-small">
                                                    <span class="fa fa-calendar"></span></span>
                                                </div>
                                            </td>

                                            <td class="aktView1" style="width:100px;">
                                                <div id="aktSLdate_<%=oRec("aktid") %>" class='input-group'>
                                                  <input type="text" style="width:100%;" class="form-control input-small" id="aktInputSLdate_<%=oRec("aktid") %>" value="<%=endDateHTML %>" placeholder="dd-mm-yyyy" readonly />
                                                    <span class="input-group-addon input-small">
                                                    <span class="fa fa-calendar"></span></span>
                                                </div>
                                            </td>

                                            <td class="aktView1"><%=left(oRec("prgnavn"),20) %></td>

                                            <td class="aktView1" style="text-align:center;">
                                                <%
                                                fsLink = "timbudgetsim.asp?FM_fy="&year(now)&"&FM_visrealprdato=1-1-"&year(now)&"&FM_sog="& strjobnr '& "&jobid="& jobid &"&func=forecast"
                                                %>
                                                <a href="<%=fsLink %>" target="_blank"><span class="fa fa-file-text"></span></a>
                                            </td>
                                            

                                            <td class="aktView1">
                                                <a class="edit_akt_btn" id="<%=oRec("aktid") %>"><span class="fa fa-pencil pull-left"></span></a>
                                                <a href="#"><span style="color:darkred;" class="fa fa-times pull-right"></span></a>
                                            </td>


                                            <!-- View 2 -->
                                            <td class="aktView2"></td>
                                            <td class="aktView2"></td>
                                            <td class="aktView2"></td>
                                            <td class="aktView2"></td>
                                            <td class="aktView2"></td>
                                            <td class="aktView2"></td>
                                            <td class="aktView2"></td>
                                            <td class="aktView2"></td>
                                            <td class="aktView2"></td>

                                            <td class="aktView3"></td>
                                            <td class="aktView3"></td>
                                            <td class="aktView3"></td>
                                            <td class="aktView3"></td>
                                            <td class="aktView3"></td>
                                            <td class="aktView3"></td>
                                            <td class="aktView3"></td>
                                            <td class="aktView3"></td>
                                            <td class="aktView3"></td>

                                        </tr>

                                        <%
                                            oRec.movenext
                                            wend
                                            oRec.close
                                            end if
                                        %>

                                    </tbody>

                                </table>

                                <div class="row">
                                    <div class="col-lg-12" style="text-align:center;"><a id="newActivtyBtn" class="btn btn-default btn-sm"><b>+</b></a></div>
                                </div>

                                <div class="row" style="display:none;">
                                    <div class="col-lg-2">
                                        <select id="prgSEL" class="form-control input-small">
                                            <%
                                                strSQL = "SELECT id, navn FROM projektgrupper"
                                                oRec.open strSQL, oConn, 3
                                                while not oRec.EOF

                                                    if cint(oRec("id")) = 10 then
                                                    strSEL = "SELECTED"
                                                    else
                                                    strSEL = ""
                                                    end if


                                                    response.Write "<option value='"&oRec("id")&"' "& strSEL &">"& oRec("navn") &"</option>"

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


                            </div>
                        </div>
                    </div>
                  </div>



            </div><!-- /.portlet body -->
            </div><!-- /.container -->
           

        <%end select %>


            </form>

        </div><!-- /.wrapper -->
    </div><!-- /.content -->

    

<!--#include file="../inc/regular/footer_inc.asp"-->