<!--#include file="../inc/connection/conn_db_inc.asp"-->

     
<%
 '**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")

    case "FN_sogjobogkunde" 

        
                '*** SØG kunde & Job            
                
                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                jobkundesog = request("jq_newfilterval")
                else
                filterVal = 0
                jobkundesog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")

                if filterVal <> 0 then
            
                 lastKid = 0
                
                  'strJobogKunderTxt = strJobogKunderTxt &"<span style=""color:red; font-size:9px; float:right;"" class=""luk_jobsog"">[X]</span>"    
                         


                strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM timereg_usejob AS tu "_ 
                &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                &" WHERE tu.medarb = "& medid &" AND (j.jobstatus = 1 OR j.jobstatus = 3) AND "_
                &" (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '%"& jobkundesog &"%' OR "_
                &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"')  AND kkundenavn <> ''"_
                &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT 50"       
    

                 'response.write "strSQL " &strSQL
                 'response.end

                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
               ' if lastKid <> oRec("kid") then
                'strJobogKunderTxt = strJobogKunderTxt &"<br><br><b>"& oRec("kkundenavn") &" "& oRec("kkundenr") &"</b><br>"
               ' end if 
                 
                'strJobogKunderTxt = strJobogKunderTxt & "<input type=""hidden"" id=""hiddn_job_"& oRec("jid") &""" value="""& oRec("jobnavn") & " ("& oRec("jobnr") &")"">"
                'strJobogKunderTxt = strJobogKunderTxt & "<input type=""checkbox"" class=""chbox_job"" id=""chbox_job_"& oRec("jid") &""" value="& oRec("jid") &"> "& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"<br>" 
                strJobogKunderTxt = strJobogKunderTxt & "<option value="& oRec("jid") &" "& jobSEL &">"& oRec("jobnavn") & " ("& oRec("jobnr") &")"
                
                lastKid = oRec("kid") 
                oRec.movenext
                wend
                oRec.close

              


                    '*** ÆØÅ **'
                    call jq_format(strJobogKunderTxt)
                    strJobogKunderTxt = jq_formatTxt

                    response.write strJobogKunderTxt

                end if



     case "FN_sogakt"

               
                '*** Søg Aktiviteter 
                

                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                aktsog = request("jq_newfilterval")
                else
                filterVal = 0
                aktsog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")
                aktid = request("jq_aktid")
    
                if len(trim(request("jq_jobid"))) <> 0 then        
                jobid = request("jq_jobid")
                else
                jobid = 0
                end if

                'positiv aktivering
                pa = 0
               ' if len(trim(request("jq_pa") )) <> 0 then
                'pa = request("jq_pa") 
                'else
                'pa = 0
               ' end if
        
                'pa = 0
            

                if filterVal <> 0 then
            
                 
    
                'strAktTxt = strAktTxt &"<span style=""color:#999999; font-size:9px; float:right;"" class=""luk_aktsog"">X</span>"    
                         

               if pa = "1" then
               strSQL= "SELECT a.id AS aid, navn AS aktnavn FROM timereg_usejob LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
               &" WHERE tu.medarb = "& usemrn &" AND tu.jobid = "& jobid &" AND aktid <> 0 AND a.navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"   


               else


                        '*** Finder medarbejders projektgrupper 
                        '** Medarbejder projektgrupper **'
                        medarbPGrp = "#0#" 
                        strMpg = "SELECT projektgruppeId, medarbejderId, teamleder FROM progrupperelationer WHERE medarbejderId = "& medid & " GROUP BY projektgruppeId"

                        oRec5.open strMpg, oConn, 3
                        while not oRec5.EOF
                        medarbPGrp = medarbPGrp & ",#"& oRec5("projektgruppeId") &"#"         
        
                        oRec5.movenext
                        wend
                        oRec5.close 


           




               strSQL= "SELECT a.id AS aid, navn AS aktnavn, projektgruppe1, projektgruppe2, projektgruppe3, "_
               &" projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter AS a "_
               &" WHERE a.job = " & jobid & " AND navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"      
    

               
            
                end if

                 'response.write "strSQL " &strSQL
                 'response.end


                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
                 showAkt = 0
                if instr(medarbPGrp, "#"& oRec("projektgruppe1") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe2") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe3") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe4") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe5") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe6") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe7") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe8") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe9") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe10") &"#") <> 0 then
                showAkt = 1
                end if 


                
                
                if cint(showAkt) = 1 then 
                 
                'strAktTxt = strAktTxt & "<input type=""hidden"" id=""hiddn_akt_"& oRec("aid") &""" value="""& oRec("aktnavn") &""">"
                'strAktTxt = strAktTxt & "<input type=""checkbox"" class=""chbox_akt"" id=""chbox_akt_"& oRec("aid") &""" value="& oRec("aid") &"> "& oRec("aktnavn") &"<br>" 
                strAktTxt = strAktTxt & "<option value="& oRec("aid") &" "& optionFcDis &">"& oRec("aktnavn") &" "& fcsaldo_txt &"</option>"             

                end if
                
                oRec.movenext
                wend
                oRec.close

              


                    '*** ÆØÅ **'
                    call jq_format(strAktTxt)
                    strAktTxt = jq_formatTxt


                    response.write strAktTxt

                end if    





        case "tilfoj_akt"

        jobid = request("FM_jobid")
        aktid = request("FM_aktid")
        medid = request("FM_medid_id")

        'oConn.execute("UPDATE timereg_usejob SET favorit = 1 WHERE aktid = "&aktid&"")
        Strtilfojakt = "INSERT INTO timereg_usejob SET jobid = "&jobid &", favorit = 1, aktid ="& aktid& ", medarb = "& medid
        oConn.execute(Strtilfojakt)


        akt_jobid = jobid
        akt_aktid = aktid


        'response.Write Strtilfojakt
        'response.end 

        'Response.redirect "favorit.asp?"

        'strSQLUpdjWiphist = "INSERT INTO wip_historik (dato, restestimat, stade_tim_proc, medid, jobid) VALUES ('"& ddDato &"', "& rest &", "& timerproc &", "& session("mid") &", "& jobid &")"
        'oConn.Execute(strSQLUpdjWiphist)


        end select
        response.end
        end if
    %>


<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../timereg/inc/convertDate.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<script src="js/favorit_jav.js"></script>
<link rel="stylesheet" href="../../bower_components/bootstrap-datepicker/css/datepicker3.css">

<style>

    .picmodal:hover,
    .picmodal:focus {
    text-decoration: none;
    cursor: pointer;
}
   
</style>

<%call menu_2014 %>
<div class="wrapper">
<div class="content">

    <%
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if


    func = request("func")


	'call menu_2014 


    stDato = request("stDato")

    redim tjekdag(7)
    tjekdag(1) = stDato
            
    for x = 2 to 7
    tjekdag(x) = dateAdd("d", x-1, stDato)
    next



    medid = request("FM_medid")


    varTjDatoUS_man = request("varTjDatoUS_man")
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

    datoMan = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
    datoSon = day(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& year(varTjDatoUS_son)

    next_varTjDatoUS_man = dateadd("d", 7, varTjDatoUS_man)
	next_varTjDatoUS_son = dateadd("d", 7, varTjDatoUS_son)
    next_varTjDatoUS_man = year(next_varTjDatoUS_man) &"/"& month(next_varTjDatoUS_man) &"/"& day(next_varTjDatoUS_man)
	next_varTjDatoUS_son = year(next_varTjDatoUS_son) &"/"& month(next_varTjDatoUS_son) &"/"& day(next_varTjDatoUS_son)


    prev_varTjDatoUS_man = dateadd("d", -7, varTjDatoUS_man)
	prev_varTjDatoUS_son = dateadd("d", -7, varTjDatoUS_son)
    prev_varTjDatoUS_man = year(prev_varTjDatoUS_man) &"/"& month(prev_varTjDatoUS_man) &"/"& day(prev_varTjDatoUS_man)
	prev_varTjDatoUS_son = year(prev_varTjDatoUS_son) &"/"& month(prev_varTjDatoUS_son) &"/"& day(prev_varTjDatoUS_son)




    select case func 

    case "fjernfavorit"


        id = request("id")

        %>
            <div class="container">
                <div class="portlet">
                    <h3 class="portlet-title"><u>Favorit liste</u></h3>
                    <div class="portlet-body">
                        <%response.Write "aktivi id" & id  %>
                    </div>

                </div>
            </div>
        <%
        

        oConn.execute("UPDATE timereg_usejob SET favorit = 0 WHERE aktid = "&id&"")

        response.Redirect "favorit.asp"


    case "tilfojfavorit"

        id = request("id")
                            
            %>
                <div class="container">
                <div class="portlet">
                    <h3 class="portlet-title"><u>Favorit liste</u></h3>
                    <div class="portlet-body">
                        <%response.Write "favor" & id  %>
                    </div>

                </div>
                </div>
            <%

        oConn.execute("UPDATE timereg_usejob SET favorit = 1 WHERE aktid = "&id&"")

        response.Redirect "favorit.asp"

    case else

%>




            <div class="container">
                <div class="portlet">
                    <h3 class="portlet-title"><u>Favorit liste</u></h3>
                    <div class="portlet-body">

                        <form action="favorit.asp?sogsubmitted=1" method="post">
                        <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
                        <div class="row">
                            <div class="col-lg-3">
                                 
                                <%
                                    strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe, init FROM medarbejdere WHERE mansat <> 2 GROUP BY mid ORDER BY Mnavn" 
                                %>
                                <select name="FM_medid" id="FM_medid" <%=progrpmedarbDisabled  %> class="form-control input-small"  onchange="submit();">
                                    <%

                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF
                                
                                    
                                    StrMnavn = oRec("Mnavn")
                                    StrMinit = oRec("init")
	
                                    if cdbl(medid) = cdbl(oRec("Mid")) then
				                    isSelected = "SELECTED"
				                    else
				                    isSelected = ""
				                    end if

				                    %>
                                     <option value="<%=oRec("Mid")%>" <%=isSelected%>><%=StrMnavn &" "& StrMinit%></option>
                                    <%
                                    oRec.movenext
                                    wend
                                    oRec.close  
                                    %>
                                </select>
                            </div>

                            <div class="col-lg-2">
                                <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="aarslut" value="<%=aarslut %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>
                            <div class="col-lg-5"></div>
                            <h4 class="col-lg-2" style="text-align:right"><a href="#" ><</a>&nbsp Uge 12 &nbsp<a href="#" >></a></h4>
                            
                        </div>
                        </form>
                        <%response.Write "id: " & medid%>


                        <table class="table dataTable table-striped table-bordered ui-datatable">

                            <thead>
                                <tr>
                                    <th>Kunde/Job</th>
                                    <th>Aktivitet</th>

                                    <%
                                            perInterval = 6 'dateDiff("d", varTjDatoUS_man, varTjDatoUS_son, 2,2) 
                                            perIntervalLoop = perInterval

                                            for l = 0 to perIntervalLoop 
        
                                            if l = 0 then
                                            varTjDatoUS_use = varTjDatoUS_man
                                            else
                                            varTjDatoUS_use = dateAdd("d", l, varTjDatoUS_man)
                                            end if 

                                            showdate = DatePart("yyyy",varTjDatoUS_use) & "-" & Right("0" & DatePart("m",varTjDatoUS_use), 2) & "-" & Right("0" & DatePart("d",varTjDatoUS_use), 2)

                                            'showweekdayname = weekdayname(weekday(varTjDatoUS_use, 1))
                                            daynamenum = weekday(varTjDatoUS_use,1)

                                            daynameword = WeekDayName(daynamenum,true)

                                            'response.Write weekdayname(weekday(varTjDatoUS_use, 1))

                                            %>
                                                <th style="width:75px"><%=UCase(Left(daynameword,1)) & Mid(daynameword,2) %> <br />
                                                    <%=showdate %>
                                                </th>
                                            <%

                                            next
                                    %>
                                    <th>Total</th>
                                    <th></th>
                                </tr>
                            </thead>

                            <tbody>
                                <%
                                     favoriter = 0
                                     StrSqlfav = "SELECT medarb, jobid, aktid, forvalgt_af FROM timereg_usejob WHERE medarb = "& medid & " AND favorit <> 0" 
                                     
                                     oRec.open StrSqlfav, oConn, 3
                                     while not oRec.EOF

                                     jobid = oRec("jobid")
                                     aktid = oRec("aktid")
                                     medarb = oRec("medarb")

                                     favoriter = favoriter + 1
                                     'response.Write "favs: " & favoriter
                                     'response.Write "<br>" & jobid
                                     for i = 1 to favoriter
                                        i = i + 1
                                        'response.Write i
                                        StrSQLjob = "SELECT id, jobnavn FROM job WHERE id ="& jobid

                                        oRec3.open StrSQLjob, oConn, 3
                                        if not oRec3.EOF then

                                        jobnavn = oRec3("jobnavn")

                                        StrSQLakt = "SELECT id, navn, beskrivelse, budgettimer FROM aktiviteter WHERE id ="& aktid

                                        oRec2.open StrSqlakt, oConn, 3
                                        if not oRec2.EOF then

                                        aktNavn = oRec2("navn")
                                        TaktId = oRec2("id")
                                        aktbudgettimer = oRec2("budgettimer")
                                        
                                        %>
                                        <tr>
                                            <td><%=jobnavn %></td>
                                            <td><%=aktNavn %>

                                                <span id="modal_<%=TaktId %>" style="color:cornflowerblue;" class="fa fa-book pull-right picmodal"></span>                                               
                                                <div id="myModal_<%=TaktId %>" style="display:none">
                                                
                                                    <%
                                                        StrSqltimerialt = "SELECT sum(Timer) as timer FROM timer WHERE TAktivitetId ="& TaktId
                                                        oRec6.open StrSqltimerialt, oConn, 3
                                                        if not oRec6.EOF then
                                                        timerforalle = oRec6("timer")
                                                        end if
                                                        oRec6.close

                                                        StrSqltotaltimer = "SELECT TAktivitetId, sum(timer) as timer FROM timer WHERE TAktivitetId ="& TaktId &" AND tmnr ="&medid
                                                        oRec5.open StrSqltotaltimer, oConn, 3
                                                        if not oRec5.EOF then
                                                
                                                        timertotal = oRec5("timer")
                                                        
                                                        if timerforalle > aktbudgettimer then
                                                           txtcolor = "red"
                                                        else
                                                            txtcolor = "green"
                                                        end if

                                                    %>
                                                        <span style="font-size:75%; color:#5582d2">Forkalk.: <%=aktbudgettimer %></span>
                                                        <span style="font-size:75%; color:<%=txtcolor%>;">Real:<%=timerforalle %></span>
                                                        <span style="font-size:75%; color:#5582d2;">Egne: <%=timertotal %></span>
                                                    <%
                                                        end if
                                                        oRec5.close 
                                                    %>
                                                
                                                </div>
                                            </td>

                                            <%
                                                
                                                for l = 0 to 6
                                                    

                                                    if l = 0 then
                                                        timerdato = varTjDatoUS_man
                                                    else
                                                        timerdato = dateAdd("d",1,timerdato)
                                                    end if

                                                    timerdato = year(timerdato) & "-" & month(timerdato) & "-" & day(timerdato) 

                                                    StrSQLtimer = "SELECT TAktivitetId, Timer FROM timer WHERE TAktivitetId ="& TaktId & " AND tdato = "& "'" & timerdato & "'"
                                                
                                                     oRec4.open StrSQLtimer, oConn, 3
                                                     if not oRec4.EOF then
                                                
                                                     timerdag = oRec4("Timer")
                                                    
                                                    %>
                                                        <td><input type="text" style="width:75px;" class="form-control input-small" name="FM_timerdag" value="<%=timerdag %>" /></td>
                                                    <%


                                                    else
                                                    timerdag = 0
                                                    %>
                                                        <td><input type="text" style="width:75px;" class="form-control input-small" name="FM_timerdag" value="<%=timerdag %>" /></td>
                                                    <%

                                                     end if 
                                                     oRec4.close

                                                next

                                                 'StrSQLtimer = "SELECT TAktivitetId, Timer FROM timer WHERE TAktivitetId ="& aktId
                                                
                                                 'oRec4.open StrSQLtimer, oConn, 3
                                                 'if not oRec4.EOF then
                                                
                                                    'timerdag = oRec4("Timer")

                                                 'end if 
                                                 'oRec4.close
                                                 
                                            %>

                                            <td>
                                                <%
                                                    ugestart_dato = year(datoMan) & "-" & month(datoMan) & "-" & day(datoMan)
                                                    ugeslut_dato = year(datoSon) & "-" & month(datoSon) & "-" & day(datoSon)

                                                    'response.Write ugestart_dato & "SØNDAG: " & ugeslut_dato
                                                    
                                                    StrSqlweektotal = "SELECT sum(timer) as timer FROM timer WHERE TAktivitetId ="& TaktId & " AND tdato BETWEEN '"& ugestart_dato &"' AND '"& ugeslut_dato &"'" 
                                                
                                                    oRec7.open StrSqlweektotal, oConn, 3
                                                    if not oRec7.EOF then
                                                    timerweektotal = oRec7("timer")
                                                    
                                                    if timerweektotal <> 0 then
                                                    else
                                                    timerweektotal = 0
                                                    end if

                                                    end if
                                                    oRec7.close    
                                                    
                                                %>
                                                <%=timerweektotal %>
                                            </td>

                                            <td>
                                                <a href="favorit.asp?id=<%=oRec2("id") %>&func=fjernfavorit"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a>
                                            </td>

                                        </tr>
                                        <%
                                        end if
                                        oRec2.close

                                        end if
                                        oRec3.close

                                    next

                                    oRec.movenext
                                    wend
                                    oRec.close 
                                    %>


                                  <!--  <tr>
                                        <td colspan="10"><input type="text" id="aktiviteter_sog" class="aktivitet_sog form-control input-small" />
                                            <select id="aktiviteterfelt" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none;">
                                            <option><%=week_txt_007 %>..</option>
                                            </select>
                                        </td>
                                        
                                    </tr> -->

                                    <tr style="visibility:hidden; display:none;" id="FN_akt_tilfojed">
                                        <td><input type="text" value="<%=akt_jobid %>" /></td>
                                        <td><input type="text" value="<%=akt_aktid %>" /></td>


                                        <%for i = 0 to 8 %>
                                        <td></td>
                                        <%next %>
                                    </tr>    

                                    <tr>  

                                        <input type="hidden" value="0" name="FM_pa" />
                                        <input type="hidden" id="FM_jobid" value=""/>

                                        <td><input type="text" class="FM_job form-control input-small" id="FM_job" value=""/>
                                          <!-- <div id="dv_job"></div> -->
                                            <select id="dv_job" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none;">
                                                <option><%=week_txt_007 %>..</option>
                                            </select>
                                        </td>                                           
                             

                                        <td>
                                            <input type="hidden" name="FM_aktivitetid" id="FM_aktid" value=""/>
                                            <input type="text" class="aktivitet_sog form-control input-small" id="FM_akt" value="" />
                                            <!--<div id="dv_akt"></div> -->
                                            <select id="dv_akt" class="form-control input-small chbox_akt" size="10" style="visibility:hidden; display:none;">
                                                <option><%=week_txt_007 %>..</option>
                                            </select>
                                        </td>
                                        <td style="text-align:center">
                                            <input type="hidden" id="FM_medid_id" value="<%=medid %>" />
                                            <button type="submit" class="tilfoj_akt btn btn-success btn-sm"><b>Tilføj</b></button>
                                            <div id="dv_akttil"></div>
                                        </td>
                                        <% for d = 0 to 7 %>
                                        <td></td>
                                        <%next %>
                                    </tr>

                            </tbody>

                        </table>

  
                           
                            <%
                                strSQL = "SELECT id, medarb, aktid, favorit FROM timereg_usejob WHERE medarb ="& medid & " AND aktid <> 0"
                            %>
                               <!-- <select name="FM_favorittilfoj" class="form-control input-small"  onchange="submit();"> -->
                            <%
                              

                                  oRec.open strSQL, oConn, 3

                                  While not oRec.EOF

                                  aktid = oRec("aktid")
                                %>
                                   <!-- <a href="favorit.asp?func=tilfojfavorit&id=<%=aktid %>"><%=aktid %></a> -->
                                <%
                                  oRec.movenext
                                  wend                     
                                  oRec.close
                              %>                               
                           <!-- </select> -->
                        

                    </div>
                </div>
            </div>

            



    <%end select %>



        </div>
    </div>


<script type="text/javascript">


$(".picmodal").click(function() {

    var modalid = this.id
    var idlngt = modalid.length
    var idtrim = modalid.slice(6, idlngt)

    //var modalidtxt = $("#myModal_" + idtrim);
    var modal = document.getElementById('myModal_' + idtrim);

    
    if (modal.style.display !== 'none')
    {
        modal.style.display = 'none';
    }
    else
    {
        modal.style.display = 'block';
    }
   

});


</script>


<!--#include file="../inc/regular/footer_inc.asp"-->