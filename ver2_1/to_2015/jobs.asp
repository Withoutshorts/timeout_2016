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
                

                kundeKpersArr(0) = "<option value='0'>V�lg kontaktperson..</option>"
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
              '*** ��� **'
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
	id = 0
	else
	id = request("id")
	end if









    Select case func 
    case "dbopr"

                                '********************** Opretter job ***************************


                                kid = request("FM_kunde")
                                kunderef = request("FM_kpers")
                                jobnavn = request("FM_jobnavn")
                                'jobnr = request("FM_jobnr")
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
				                nytjobnr = strjnr
				                tilbudsnrKri = ""
				                
				                if request("FM_usetilbudsnr") = "j" then
				                nyttlbnr = tlbnr
				                tilbudsnrKri = ", tilbudsnr = "& nyttlbnr &""
				                end if
				                
				                strSQL = "UPDATE licens SET jobnr = "&  nytjobnr &" "& tilbudsnrKri &" WHERE id = 1"
				                oConn.execute(strSQL)

                            
                                '** Rediger
                                'response.redirect("jobs.asp?func=red&id="&id)
                                    

                                '*** Liste
                                response.redirect("../timereg/jobs.asp")





    case "opret", "red"
        %>
        <script src="js/job_2015_jav.js"></script>
        <%call menu_2014 %>

        <%


        if func = "red" then
       
        strSQL = "SELECT id, jobnavn, jobnr, kkundenavn, jobknr, kid, kundekpers, jobstartdato, jobslutdato FROM job "_
        &" LEFT JOIN kunder WHERE id = " & id &" AND kunder.kid = jobknr"

        'Response.Write strSQL
	    'Response.flush
	
	    oRec.open strSQL, oConn, 3
	
	    if not oRec.EOF then

	        strNavn = oRec("jobnavn")
	        strjobnr = oRec("jobnr")
	        strKnavn = oRec("kkundenavn")
	        strKnr = oRec("jobknr")
            kundekpers = oRec("kundekpers")
            jobstdato = oRec("jobstartdato")
            jobsldato = oRec("jobslutdato")


             jobans1 = 0
        jobans2 = 0
        jobans3 = 0
        jobans4 = 0
        jobans5 = 0
        
        end if


        oRec.close

        dbfunc = "dbred"

        else 'opr

        fmkunde = 0
        strKnr = 0
        kundekpers = 0

        dbfunc= "dbopr"

        jobstdato = day(now) & "-" & month(now) & "-" & year(now)
        'jobstdato = dateadd("d", -7, jobstdato) 
        jobsldato = dateadd("m", 1, jobstdato)

        jobans1 = 0
        jobans2 = 0
        jobans3 = 0
        jobans4 = 0
        jobans5 = 0

        end if
%>










<div class="wrapper">
    <div class="content">



    
 <!------------------------------- Sideindhold------------------------------------->
 
        <form id="opretproj" action="jobs.asp?func=<%=dbfunc %>" method="post">
        <input type="hidden" name="" id="kundekpersopr" value="<%=kundekpers%>" />
        <div class="container">

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

                                        

                                        <div class="col-lg-1">Projekt nr:</div>
                                        <div class="col-lg-2"><input class="form-control input-small" type="text" name="FM_jobnr" value="" placeholder="eks. 1234" disabled/></div>


                                </div>

                                <div class="row">
                                        
                                        <div class="col-lg-1">Status:</div>
                                        <div class="col-lg-2">
                                            <select class="form-control input-small" name="FM_status">
                                                <option value="1">Aktiv</option>
                                                <option value="2">Til fakturering</option>
                                                <option value="3">Lukket</option>
                                                <option value="4">Gennemsyn</option>
                                                <option value="5">Tilbud</option>
                                            </select>
                                        </div>
                               
                                    <div class="col-lg-1">Ansvarlig:</div>
                                    <div class="col-lg-2">
                                            <select class="form-control input-small" name="FM_jobans1" id="FM_jobans" >
						                        <option value="0">Ingen</option>
							                        <!----- Jobs linje 5361 --->
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
                                                <option value="1">Reng�ring</option>
                                                <option value="2">Udvikler</option>
                                                <option value="3">K�rsel</option>
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
                                    <div class="col-lg-2"><a href="#">Tilf�j grp.</a></div>
                                </div>
                                -->

                                    <div class="row">
                                   <div class="col-lg-12 pad-t20">
                                                        <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opret</b></button>
                                                 </div>
                                   </div><!-- /.row -->

                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div> <!-- /.panel -->
                   </div>


                
                    <%
                    oprjobtype = 2    
                    if oprjobtype = 2 then %>

                    <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse4">
                                Aktiviteter
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapse4" class="panel-collapse collapse in">
                            <div class="panel-body">
                            
                             <script src="js/tableedittest.js" type="text/javascript"></script>
                              <table class="table dataTable table-striped table-bordered table-hover ui-datatable editabletable">
                                  <thead>
                                      <tr>
                                          <th style="width:20%">Aktivitet</th>
                                          <th style="width:10%">Status</th>
                                          <th style="width:9%">Type</th>
                                          <th style="width:5%">Timer</th>
                                          <th style="width:5%">Pris</th>
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
                                      <tr>
                                          <td><a href="#">aa</a></td>
                                          <td><div contenteditable>hej </div></td>
                                          <td>fakt. bar</td>
                                          <td>6</td>
                                          <td>122 kr</td>
                                          <td>09-07</td>
                                          <td>12-07</td>
                                          <%if avansproobr = 1 then%>
                                          <td>08:30</td>
                                          <td>10:30</td>
                                          <%end if %>
                                          <td><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
                                          <td><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
                                      </tr>

                                      <tr>
                                          <td><a href="#">aa</a></td>
                                          <td><div contenteditable>hej </div></td>
                                          <td>fakt. bar</td>
                                          <td>6</td>
                                          <td>122 kr</td>
                                          <td>09-07</td>
                                          <td>12-07</td>
                                          <%if avansproobr = 1 then%>
                                          <td>08:30</td>
                                          <td>10:30</td>
                                          <%end if %>
                                          <td><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
                                          <td><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></td>
                                      </tr>
                                  </tbody>
                              </table>
                             <div class="row">
                                 <div class="col-lg-3"><a href="#">Tilf�j aktiitet</a></div>
                             </div>
                            <div class="row">
                                 <div class="col-lg-3"><a href="#">Tilf�j stam akt. gruppe</a></div>
                             </div>

                            </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                    </div>
                </div> <!-- /.panel-group -->



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
                                <div class="row">
                                    <div class="col-lg-2">Job medansvarlig:</div>
                                    <div class="col-lg-3">

                                        <!-- data model - se gl. joboprettelse
                                        tabel: medarbejdere
                                        felter: id, navn
                                        -->

                                        <select class="form-control input-small">
                                            <option value="1">Henrik</option>
                                            <option value="2">S�ren</option>
                                            <option value="3">jesper</option>
                                        </select>
                                    </div>
                                    <div class="col-lg-1"><input class="form-control input-small" type="number" name="FM_exch" value="100 %" style="text-align:right"/></div>
                                    <div class="col-lg-1"><a href="#"><span class="glyphicon glyphicon-plus"></span></a></div>
                                </div>
                                <br />
                                <div class="row">                                                           
                                    <div class="col-lg-2">Jobejer:</div>
                                    <div class="col-lg-3">
                                        <select class="form-control input-small">
                                            <option value="1">Henrik</option>
                                            <option value="2">S�ren</option>
                                            <option value="3">jesper</option>
                                        </select>
                                    </div>
                                </div>
                             
                                <br />

                                <%if func = "red" then %>
                                <div class="row">
                                    <div class="col-lg-7"><input type="checkbox" name="FM_opdaterprojektgrupper" id="FM_opdaterprojektgrupper" value="1" <%=syncAktProjGrpCHK %>> <b>Overf�r</b> <!--(synkroniser) valgte projektgrupper, 
                                til <b>aktiviteterne</b> p� dette job.--> </div>
                                </div>
                                <%end if %>

                                 <%
                           if lto <> "execon" then%>
                                <div class="row">
                                    <div class="col-lg-7">
								        <input type="checkbox" name="FM_gemsomdefault" id="FM_gemsomdefault" value="1"><b> Skift <!--standard--> forvalgt projektgruppe</b> <a data-toggle="modal" href="#styledModalSstGrp20"><span class="fa fa-info-circle"></span></a> <!--til den gruppe der her v�lges som projektgruppe 1.
								        <span style="color:#999999;">Gemmes som cookie i 30 dage.</span> -->
                                    </div>
                                </div>
                                <div id="styledModalSstGrp22" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
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
                                 </div>
                                <%end if %>
								
                               
                              

                                <%if func <> "red" then
                                 
                                    if lto = "jm" OR lto = "synergi1" OR lto = "micmatic" then 'OR lto = "lyng" OR lto = "glad" then
                                    forvalgCHK = "CHECKED"
                                    else
                                    forvalgCHK = ""
                                    end if
                                 
                                 else
                                 
                                 forvalgCHK = ""

                                 end if %>

                                <br /><br />

                               <div id="accordion-help" class="panel-group accordion-simple">
                                <div class="panel">
                                
                                      <a data-toggle="collapse" data-parent="#accordion-help" href="#faq-general-1">Forretingsomr�der:</a>
                                      <!-- <br /> <span>Forretningsomr�der bruges bl.a. til at se tidsforbrug p� tv�rs af kunder og job, og til at se hvilke slags opgaver man bruger sin tid p�. 
                                       <br />
                                       Alle aktiviteter p� dette job t�ller altid med i de forretningsomr�der der er valgt p� jobbet. Specifikke forretningsomr�der kan v�lges p� den enkelte aktivitet.</span>
                                       -->
                                  <div id="faq-general-1" class="panel-collapse collapse">
                                    <div class="panel-body">
                                         <%
                                         ' uTxt = "Forretningsomr�der bruges bl.a. til at se tidsforbrug p� tv�rs af kunder og job, og til at se hvilke slags opgaver man bruger sin tid p�."
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

                                            
                                            <%

                                            '** Finder kundetype, til forvalgte forrretningsomr�der '***
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

                                        
                                        <div class="row">
                                            <div class="col-lg-5"><b>Konto:</b></div>
                                        </div>
                                            <div class="row">
                                                <div class="col-lg-12">    
                                                   
                                            <select name="FM_fomr" id="FM_fomr" multiple="multiple" size="5" class="form-control input-small">
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
                                        <br />
                                       <!-- <div class="row">
                                            <div class="col-lg-12"><b>Forvalgt konto p� faktura / ERP system</b><br />
                                            V�lg herunder blandt de forretningsomr�der der har tilknyttet en oms�tningskonto, og hvor fakturaer p� dette job skal posteres p� denne konto:<br />  
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

                                
                                
                                
                                <br />

                                <div id="accordion-help" class="panel-group accordion-simple">
                                <div class="panel">
                                
                                      <a data-toggle="collapse" data-parent="#accordion-help" href="#faq-general-2">Avanceret indstillinger:</a>
                                       <!--<br /> <span>Tildel bla. prioitet, faktura-indstillinger, pre-konditioner, kundeadgang mm.</span>
                                        -->
                                  <div id="faq-general-2" class="panel-collapse collapse in">
                                    <div class="panel-body">
                                                     
                                        <div class="row">
                                            <div class="col-lg-2"><b>Prioitet:</b> <a data-toggle="modal" href="#styledModalSstGrp23"><span class="fa fa-info-circle"></span></a></div>
                                            <div class="col-lg-3">&nbsp</div>
                                            <div class="col-lg-4"><b>Pre-konditioner opfyldt:</b> <a data-toggle="modal" href="#styledModalSstGrp24"><span class="fa fa-info-circle"></span></a> <!--<br /><span style="font-size:90%; font-weight:lighter">Underleverand�r klar, materialer indk�bt mm.</span>--></div>
                                        </div>
                                        <div id="styledModalSstGrp23" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                            <div class="modal-dialog">
                                                <div class="modal-content" style="border:none !important;padding:0;">
                                                  <div class="modal-header">
                                                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                      <h5 class="modal-title"></h5>
                                                    </div>
                                                    <div class="modal-body">
                                                         Prioiteter p� nuv�rende aktive job ligger mellem: <b><%=lowestRiskval%> - <%=highestRiskval%></b><br /><br />
									                     <b>-1 = Internt job</b> vises ikke under fakturering og igangv�rende job.<br />
                                                         <b>-2 = HR job</b> vises i HR mode p� timereg. siden<br />
                                                         <b>-3 = Internt job</b> men der skal kunne laves ressouceforecast p� dette job. 
									                     <br /><br />
                                                         -1 / -2 / -3 medf�rer enkel visning af aktivitetslinjer p� timereg. siden, dvs. der bliver ikke vist tidsforbrug, start- og slut -datoer mv. p� aktiviteterne. <br /><br />&nbsp;
									
                                                    </div>
                                                </div>
                                            </div>
                                        </div>  
                                        <div id="styledModalSstGrp24" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                            <div class="modal-dialog">
                                                <div class="modal-content" style="border:none !important;padding:0;">
                                                  <div class="modal-header">
                                                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                      <h5 class="modal-title"></h5>
                                                    </div>
                                                    <div class="modal-body">                                                
                                                        Underleverand�r klar, materialer indk�bt mm.
                                                    </div>
                                                </div>
                                            </div>
                                        </div> 
                                        
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
                                            <div class="col-lg-1"> <input id="prio" name="prio" type="text" value="<%=intprio %>" class="form-control input-small" /></div>
                                           <!-- <div class="col-lg-3">Prioiteter p� nuv�rende aktive job ligger mellem: <b><%=lowestRiskval%> - <%=highestRiskval%></b></div> -->
                                            <div class="col-lg-4">&nbsp</div>
                                            <div class="col-lg-3">
                                                <select name="FM_preconditions_met" id="Select1" size="1" class="form-control input-small">
                                                <option value="0" <%=preconditions_met_SEL0 %>>Ikke angivet</option>
                                                <option value="1" <%=preconditions_met_SEL1 %>>Ja</option>
                                                <option value="2" <%=preconditions_met_SEL2 %>>Nej - afvent</option>
                                                </select>
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-12">
									     <!---<b>1 = Internt job</b>--> <!--vises ikke under fakturering og igangv�rende job.-->
                                         <!---<b><!---2 = HR job</b> --><!--vises i HR mode p� timereg. siden-->
                                         <!---<b><!-- 3 = Internt job</b>--> <!--men der skal kunne laves ressouceforecast p� dette job.--> 
									    
                                         <!---1 / -2 / -3 medf�rer enkel visning af aktivitetslinjer p� timereg. siden, dvs. der bliver ikke vist tidsforbrug, start- og slut -datoer mv. p� aktiviteterne. <br /><br />&nbsp;-->
									    </div>
                                        </div>  
                                        


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

                                       
                                        
                                          
                                        <br />
                                        

                                        <div class="row">
                                            <div class="col-lg-4"><b>Fakturaindstillinger:</b><!--<br /><span style="font-size:11px; font-weight:lighter;">(Nedarves fra kunde ved joboprettelse)</span>--></div>
                                            <div class="col-lg-1">&nbsp</div>
                                            <div class="col-lg-3"><b>Skal job v�re �ben for kunde?</b></div>

                                        </div>

                                        <div class="row">
                                            <div class="col-lg-1">Valuta:</div>
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
                                            <div class="col-lg-1">&nbsp</div>
                                            <div class="col-lg-3"><input type="checkbox" name="FM_kundese" id="FM_kundese" value="1" <%=kundechk%>>&nbsp;G�r job tilg�ngeligt for kontakt.</div>
                                        </div>

                                        <div class="row">
                                            <div class="col-lg-1">Moms:</div>
                                            <div class="col-lg-3"><%strSQLmoms = "SELECT id, moms FROM fak_moms WHERE id <> 0 ORDER BY id " %>
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
                                            <div class="col-lg-1">&nbsp</div>
                                            <div class="col-lg-7"><b>N�r job �bnes for kontakt:</b> <a data-toggle="modal" href="#styledModalSstGrp24"><span class="fa fa-info-circle"></span></a> <!--hvorn�r skal registrerede timer s� v�re tilg�ngelige? --></div>
                                        </div>
                                        <div id="styledModalSstGrp25" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                            <div class="modal-dialog">
                                                <div class="modal-content" style="border:none !important;padding:0;">
                                                  <div class="modal-header">
                                                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                      <h5 class="modal-title"></h5>
                                                    </div>
                                                    <div class="modal-body">                                                
                                                        Hvis job �bnes for kontakt, hvorn�r skal registrerede timer s� v�re tilg�ngelige?
                                                    </div>
                                                </div>
                                            </div>
                                        </div>

                                        <div class="row">
                                            <div class="col-lg-1">Sprog:</div>
                                            <%strSQLsprog = "SELECT id, navn FROM fak_sprog WHERE id <> 0 ORDER BY id " %>
                                            <div class="col-lg-3"><%strSQLmoms = "SELECT id, moms FROM fak_moms WHERE id <> 0 ORDER BY id " %>
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
                                            <div class="col-lg-1">&nbsp</div>
                                            <div class="col-lg-4"><input type="radio" name="FM_kundese_hv" value="0" <%=hvchk1%>> Offentligg�r timer, s� snart de er indtastet.</div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-5">&nbsp</div>
                                            <div class="col-lg-6"><input type="radio" name="FM_kundese_hv" value="1" <%=hvchk2%>> Offentligg�r f�rst timer n�r jobbet er lukket. (afsluttet/godkendt)</div>
                                        </div>

                                        <br /><br />
                                        <%if func = "opret" AND step = 2 OR func = "red" then %>


                                       <!-- <div class="row">
                                            <h5 class="col-lg-4">Skal job v�re �ben for kunde?</h5>
                                        </div>

                                        <div class="row">
                                            <div class="col-lg-6"><input type="checkbox" name="FM_kundese" id="FM_kundese" value="1" <%=kundechk%>>&nbsp;<b>G�r job tilg�ngeligt for kontakt.</b><br>
		                                    <!--Hvis tilg�ngelig for kontakt tilv�lges, udsendes der en mail til kontakt-stamdata emailadressen, med link til kontakt loginside.--></div>
                                        <!-- </div> -->

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
                                        <br />
                                        <!--<div class="row">
                                            <div class="col-lg-5">
                                                <b>Hvis job �bnes for kontakt, hvorn�r skal registrerede timer s� v�re tilg�ngelige?</b><br>
		                                        <input type="radio" name="FM_kundese_hv" value="0" <%=hvchk1%>>Offentligg�r timer, s� snart de er indtastet.<br>
		                                        <input type="radio" name="FM_kundese_hv" value="1" <%=hvchk2%>>Offentligg�r f�rst timer n�r jobbet er lukket. (afsluttet/godkendt)
                                            </div>
                                        </div> -->

                                        <%end if %>
                                        </div> <!-- /.panel-body -->
                                    </div> <!-- /.panel-collapse -->
                                </div>
                                </div>




                                                                 


                   <div class="panel-group accordion-panel" id="accordion-paneled">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapse7">
                                �konomi
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapse7" class="panel-collapse collapse in">
                            <div class="panel-body">
                            
                                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                    <thead>
                                        <tr style="font-size:75%">
                                            <th style="border-right:hidden">Bruttooms�tning<br /><span style="font-size:75%">Nettooms. + Salgsomk.</span></th>
                                            <th colspan="4" style="text-align:right; border-right:hidden">= 5100</th>
                                            <th>&nbsp</th>
                                        </tr>

                                        <tr style="font-size:75%">
                                            <th>Nettoomkostning, timer <br /><span style="font-size:75%">Oms. f�r salgsomk.	</span></th>
                                            <th style="width:10%">Timer</th>
                                            <th style="width:10%">Timepris</th>
                                            <th style="width:10%">Faktor</th>
                                            <th style="width:10%">Bel�b</th>
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
                                
                                      <a data-toggle="collapse" data-parent="#accordion-help" href="#fordeltimbu">Fordel timebudget p� finans�r:</a>
                                       
                                  <div id="fordeltimbu" class="panel-collapse collapse">
                                    <div class="panel-body">

                                        <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                            <thead>
                                                <tr style="font-size:75%">
                                                    <th style="width:40%">Budget FY</th>
                                                    <td style="width:10%; text-align:right">Timer</td>
                                                    <td style="width:10%; text-align:right">Budget�r (FY)</td>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                <tr style="font-size:75%">
                                                    <td>�r 0</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2016" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>�r 1</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2017" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>�r 2</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2018" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>�r 3</td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="23" style="text-align:right"/></td>
                                                    <td><input class="form-control input-small" type="number" name="FM_exch" value="2019" style="text-align:right"/></td>
                                                </tr>
                                                <tr style="font-size:75%">
                                                    <td>�r 4</td>
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
                                    <div class="col-lg-2"><input type="checkbox" name="#" value="#"> <b>�bn</b> <a data-toggle="modal" href="#styledModalSstGrp26"><span class="fa fa-info-circle"></span></a> <!--for manuel indtastning og beregning af Brutto- og Netto -oms�tning.--></div>
                                </div>
                                <div id="styledModalSstGrp26" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                            <div class="modal-dialog">
                                                <div class="modal-content" style="border:none !important;padding:0;">
                                                  <div class="modal-header">
                                                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                      <h5 class="modal-title"></h5>
                                                    </div>
                                                    <div class="modal-body">                                                
                                                        �bn for manuel indtastning og beregning af Brutto- og Netto -oms�tning.
                                                    </div>
                                                </div>
                                            </div>
                                        </div>                 
                                <div class="row">
                                    <div class="col-lg-2"><input type="checkbox" name="#" value="#"> Fastpris <a data-toggle="modal" href="#styledModalSstGrp27"><span class="fa fa-info-circle"></span></a> <!--(bruttooms�tning benyttes ved fakturering)--> </div>
                                </div>
                                 <div id="styledModalSstGrp27" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                    <div class="modal-dialog">
                                        <div class="modal-content" style="border:none !important;padding:0;">
                                            <div class="modal-header">
                                            <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                <h5 class="modal-title"></h5>
                                            </div>
                                            <div class="modal-body">                                                
                                               <b>Fastpris</b> (bruttooms�tning benyttes ved fakturering)
                                            </div>
                                        </div>
                                    </div>
                                </div>   
                                <div class="row">
                                    <div class="col-lg-2"><input type="checkbox" name="#" value="#"> Lbn. timer <a data-toggle="modal" href="#styledModalSstGrp28"><span class="fa fa-info-circle"></span></a> <!-- (timeforbrug p� hver enkelt aktivitet * medarb. timepris benyttes ved fakturering) --> </div>
                                </div>
                                 <div id="styledModalSstGrp28" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                    <div class="modal-dialog">
                                        <div class="modal-content" style="border:none !important;padding:0;">
                                            <div class="modal-header">
                                            <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                <h5 class="modal-title"></h5>
                                            </div>
                                            <div class="modal-body">                                                
                                               <b>Lbn. timer</b> (timeforbrug p� hver enkelt aktivitet * medarb. timepris benyttes ved fakturering)
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
                        <div id="mater" class="panel-collapse collapse in">
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
                                          <th style="width:11%">Indk�bspris</th>
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
                                    <div class="col-lg-12"><a href="#">Tilf�j linje</a>
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
            </div><!-- /.portlet body -->
            </div><!-- /.container -->
           

        <%end select %>


            </form>

        </div><!-- /.wrapper -->
    </div><!-- /.content -->

    

<!--#include file="../inc/regular/footer_inc.asp"-->