

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<!--#include file="inc/job_inc.asp"-->

<script src="js/job_list_jav.js" type="text/javascript"></script>

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
                response.Write "<br> JS " & newJobstatus 
            
                call lukkedato(jobid(t), newJobstatus)

                strSQL = "UPDATE job SET jobstatus = "& newJobstatus & " WHERE id = "& jobid(t)
                'response.Write "<br> " & strSQL
                oConn.execute(strSQL)
            end if

        next

        'response.Write "DONE !"
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

                                    '*** nulstiller form på akt () ****'
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


    if level <= 2 OR level = 6 then
	editok = 1
	else
	if cint(session("mid")) = jobans1 OR cint(session("mid")) = jobans2 OR (cint(jobans1) = 0 AND cint(jobans2) = 0) then
	editok = 1
	end if
	end if


    if len(trim(request("FM_kunde"))) <> 0 then
    FM_kunde = request("FM_kunde")
    else
    FM_kunde = 0
    end if

    select case FM_kunde
    case 0
    FM_kundeSqlFilter = " jobknr <> 0"
    case else
    FM_kundeSqlFilter = " jobknr = "& FM_kunde 'AND skal på
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
    for f = 0 to 4
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
	    medarb_jobans = 0
	    end if
	end if

    if cint(medarb_jobans) = 0 then
	jobansKri = ""
	else
	jobansKri = " AND jobans1 = " & medarb_jobans 
	end if


    if len(trim(request("FM_prjgrp"))) <> 0 then
    prjgrp = request("FM_prjgrp")
    response.Cookies("2015")("prjgrp") = prjgrp
    else
        if request.Cookies("2015")("prjgrp") <> "" then
        prjgrp = request.Cookies("2015")("prjgrp")
        else
        prjgrp = 10 '** Alle
        end if
    end if

    select case prjgrp
    case 10
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


       
    '****** SØGNING **************

    if len(trim(request("FM_kunde"))) <> 0 then
    jobnr_sog = request("jobnr_sog")
    response.cookies("2015")("jobsog") = jobnr_sog
    visliste = 1
    else
        if request.cookies("2015")("jobsog") <> "" then
           
                jobnr_sog = request.cookies("2015")("jobsog")
            
                    if jobnr_sog = "%" then
                    jobnr_sog = ""
                    end if

        else
        jobnr_sog = ""
        end if
    visliste = 0
    end if

    if len(jobnr_sog) <> 0 AND jobnr_sog <> "-1" OR (cint(visliste) = 1) then
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

                sogeKri = " AND (j.jobnr LIKE '"& jobnr_sog &"' OR j.jobnavn LIKE '"& jobnr_sog &"%' OR j.id LIKE '"& jobnr_sog &"' OR Kkundenavn LIKE '"& jobnr_sog &"%' OR Kkundenr LIKE '"& jobnr_sog &"' OR rekvnr LIKE '"& jobnr_sog &"'"
			

                end if
            
            end if
			
			sogeKri = sogeKri & ") "
			
			
			show_jobnr_sog = jobnr_sog
			
		else
			jobnr_sog = "-1"
			sogeKri = " AND (j.id = -1) "
			show_jobnr_sog = ""
		end if
%>





<%call menu_2014 %>

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


<div class="wrapper">
    <div class="content">

        <div class="container" style="width:1500px;">
            <div class="portlet">
                <h3 class="portlet-title"><u>Joboversigt</u></h3>
                <div class="portlet-body">

                    <form action="jobs.asp?func=opret&id=0" method="post">
                        <div class="row">
                            <div class="col-lg-10">&nbsp;</div>
                            <div class="col-lg-2">
                                <button class="btn btn-sm btn-success pull-right"><b>Opret +</b></button><br />&nbsp;
                            </div>
                        </div>
                    </form>

                    <div class="well">
                        <form action="job_list.asp?" method="post">                                                      
                            <div class="row">
                                <div class="col-lg-2">Projekt / Job søg:</div>
                                <div class="col-lg-2"></div>
                                <div class="col-lg-2">Kunde:</div>
                            </div>
                        
                            <div class="row">
                                <div class="col-lg-4"><input name="jobnr_sog" type="text" class="form-control input-small" value="<%=show_jobnr_sog %>" /><br /><input type="checkbox" name="FM_sogakt" <%=sogaktCHK %> /> Søg aktivitet</div>

                                <div class="col-lg-4">
                                    <select class="form-control input-small" name="FM_kunde">
                                        <option value="0">Alle</option>

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
                                <div class="col-lg-2"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=godkend_txt_005 %> >></b></button></div>
                            </div> 
                            
                            <div class="row">
                                <div class="col-lg-4">Jobansvarlig:</div>               
                                <div class="col-lg-2">Status filter:</div>
                            </div>
                            
                            <div class="row">
                                <div class="col-lg-4">
                                    <select name="FM_medarb_jobans" id="FM_medarb_jobans" class="form-control input-small">
                                    <option value="0"><%=job_txt_002 & " " %>(<%=job_txt_205 %>)</option>
                                    <%
                                    mNavn = "Alle"
            
                                     strSQL = "SELECT mnavn, mnr, mid, init FROM medarbejdere WHERE mansat <> '2' ORDER BY mnavn"
                                     oRec.open strSQL, oConn, 3
                                     while not oRec.EOF
                
                                        if cint(medarb_jobans) = oRec("mid") then
                                        selThis = "SELECTED"
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
                                </div>

                                <div class="col-lg-3">
                                    <input type="CHECKBOX" name="status_filter" value="1" <%=chk1%>/> <%=job_txt_221 %><br />
                                    <input type="CHECKBOX" name="status_filter" value="2" <%=chk2%>/> <%=job_txt_222 %><br />
                                    <input type="CHECKBOX" name="status_filter" value="3" <%=chk3%>/> <%=job_txt_223 %><br />
                                    <input type="CHECKBOX" name="status_filter" value="4" <%=chk4%>/> <%=job_txt_224 %><br />
                                    <input type="CHECKBOX" name="status_filter" value="0" <%=chk0%>/> <%=job_txt_225 %><br />
                                </div>
                            </div>

                           <!-- <div class="row"><div class="col-lg-2">Forretningsområder</div></div>
                            <div class="row">
                                    <div class="col-lg-4">
                                    <%
                                    strSQLf = "SELECT id, navn FROM fomr WHERE id <> 0 ORDER BY navn"%>
                                    <select name="FM_fomr" id="Select2" multiple="multiple" class="form-control input-small">
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
                            </div> -->
                            
                            <div class="row">
                                <div class="col-lg-2"><a class="advanced_filter">Udvidet søgning <span id="ud_s_span">+</span></a></div>
                            </div>
                            
                            <div id="udvidet_sog" style="display:none">
                                <div class="row">
                                    <div class="col-lg-4">Projektgruppe</div>
                                </div>
                                <div class="row">
                                    <div class="col-lg-4">
                                        <select name="FM_prjgrp" class="form-control input-small">
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
                                    </div>
                                </div>
                            </div>
                                                    
                        </form>
                    </div>
                </div>


                <!-- Change Job Status pop up window -->
                <div id="actionMenu_jobstatus" class="modal">
                    <div class="modal-content" style="text-align:left; width:300px; height:175px; text-align:center">

                        <form action="job_list.asp?func=jobstatusUpdate&jobnr_sog=<%=jobnr_sog %>&FM_kunde=<%=FM_kunde %>&FM_medarb_jobans=<%=medarb_jobans %>" method="post">

                            <h5>Ændre Status på de valtge job til</h5>
                            <br />

                            <input type="hidden" id="FM_selectedJobids_status" name="FM_selectedJobids_status" value="" />

                            <select name="FM_jobstatus" class="form-control input-small">
                                <option disabled>Vælg</option>
                                <option value="1">Aktiv</option>
                                <option value="2">Passiv/Fak</option>
                                <option value="3">Tilbud</option>
                                <option value="4">Gen.syn</option>
                                <option value="0">Lukket</option>
                            </select>

                            <br />

                            <button type="submit" class="btn btn-sm btn-success"><b>Gem</b></button>

                        </form>
                    </div>

                    
                </div>

                <!-- Change Job end state pop up window -->
                <div id="actionMenu_enddate" class="modal">
                    <div class="modal-content" style="text-align:left; width:300px; height:175px; text-align:center">
                        
                        <form action="job_list.asp?func=jobslutdatoUpdate&jobnr_sog=<%=jobnr_sog %>&FM_kunde=<%=FM_kunde %>&FM_medarb_jobans=<%=medarb_jobans %>" method="post">
                          
                            <h5>Ændre slutdatoen på de valtge job til</h5>
                            <br />

                            <input type="hidden" id="FM_selectedJobids_enddate" name="FM_selectedJobids_enddate" value="" />

                            <div class='input-group date'>
                                <input type="text" style="width:100%;" class="form-control input-small" name="FM_newEnddate" id="jq_dato" value="" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small"><span class="fa fa-calendar"></span></span>
                            </div>

                            <br />

                            
                            <button type="submit" class="btn btn-sm btn-success"><b>Gem</b></button>

                        </form>
                    </div>   
                </div>

                <!-- Change Job / Activities' business areas -->
                <div id="actionMenu_forOm" class="modal">
                    <div class="modal-content" style="text-align:left; width:500px; height:275px;">

                        <form action="job_list.asp?func=foromUpdate&jobnr_sog=<%=jobnr_sog %>&FM_kunde=<%=FM_kunde %>&FM_medarb_jobans=<%=medarb_jobans %>" method="post">

                            <h5>Ændre forretningsområder</h5>
                            <br />

                            <input type="hidden" id="FM_selectedJobids_forOm" name="FM_selectedJobids_forOm" value="" />

                            <div style="text-align:center;">
                                <%         
                                strSQLf = "SELECT f.id, f.navn, kp.navn AS kontonavn, kp.kontonr FROM fomr AS f LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) WHERE f.id <> 0 AND f.jobok = 1 ORDER BY f.navn"
                                
                                'response.write strSQLf
                                'response.flush         
                                %>
                                <select name="FM_fomr" id="Select3" multiple="multiple" size="5" style="width:450px;" class="form-control input-small">
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
                               <!-- <input id="Checkbox3" type="checkbox" name="FM_formr_opdater" value="1" /> Ændre på de valgte jobs <br /> -->
                                <input id="Checkbox3" type="checkbox" name="FM_formr_opdater_akt" value="1" /> Ændre også på aktiviteterne i de valgte jobs 
                            </span>

                            <br /><br /><br />

                            <div style="text-align:center">
                                <button type="submit" class="btn btn-sm btn-success"><b>Gem</b></button>
                            </div>

                        </form>
                    </div>   
                </div>

                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                    <thead>
                        <tr>
                            <th>Projekt</th>
                            <th>Aktiviter <span style="font-size:9px;">(Aktive)</span></th>
                            <th>Forretnings - <br />områder</th>
                            <th>Realiseret % <br /><span style="font-size:9px;">(<%=job_txt_235 %>)</span></th>
                            <th>Brutto oms.</th>
                            <th>Timer forkalk. <br /><span style="font-size:9px;">Realiseret</span></th>                            
                         <!-- <th>Funktioner</th> -->
                            <th>Faktura hist.</th>
                            <th>Projektgrupper</th>
                            <th>Links</th>
                            <th>Status</th>
                            <th>
                                <select class="form-control input-small" id="actionMenu">
                                    <option>Vælg</option>
                                    <%if level = 1 then %>
                                    <option value="1">Forretningsområder</option>
                                    <%end if %>
                                    <option value="2">Slut Dato</option>
                                    <option value="3">Jobstatus</option>
                                </select>
                            </th>
                        </tr>
                    </thead>
                    <tbody>
                        <%
                            'strJob = "SELECT jobnavn, jobnr, jobknr, fastpris, id, budgettimer, ikkebudgettimer, jo_valuta, jo_bruttooms, jobstatus, jobstartdato, jobslutdato, "_
                            '&" projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, "_
                            '&" jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, virksomheds_proc, kundekpers, rekvnr, tilbudsnr, sandsynlighed, preconditions_met "_ 
                            '&" FROM job AS j LEFT JOIN kunder as k ON (k.kid = j.jobknr) WHERE" & FM_kundeSqlFilter & varStatusFilt & jobansKri & prjgrpSQL & sogeKri


                            strJob = "SELECT jobnavn, jobnr, jobknr, fastpris, j.id, j.budgettimer, ikkebudgettimer, jo_valuta, jo_bruttooms, jobstatus, jobstartdato, jobslutdato, "_
                            &" j.projektgruppe1, j.projektgruppe2, j.projektgruppe3, j.projektgruppe4, j.projektgruppe5, j.projektgruppe6, j.projektgruppe7, j.projektgruppe8, j.projektgruppe9, j.projektgruppe10, "_
                            &" jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, virksomheds_proc, kundekpers, rekvnr, tilbudsnr, sandsynlighed, preconditions_met, fakturerbart, k.kid, serviceaft, k.useasfak, s.navn AS aftnavn "_ 
                            &" FROM job AS j LEFT JOIN kunder as k ON (k.kid = j.jobknr) LEFT JOIN serviceaft as s ON (s.id = j.serviceaft)"
                            
                            if cint(sogakt) = 1 then
                            strJob = strJob + " LEFT JOIN aktiviteter as a ON (a.job = j.id)"
                            end if

                            strJob = strJob + " WHERE" & FM_kundeSqlFilter & varStatusFilt & jobansKri & prjgrpSQL & sogeKri
                            strJob = strJob + " GROUP BY j.id Order BY k.kkundenavn"

                            'response.Write strJob
                            oRec.open strJob, oConn, 3
                            while not oRec.EOF
                            


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
	
	                        if projektcomplt > 100 then
	                        showprojektcomplt = projektcomplt
	                        projektcomplt = 100
	                        bgdiv = "Crimson"
	                        else
	                        showprojektcomplt = projektcomplt
	                        projektcomplt = projektcomplt
	                        bgdiv = "yellowgreen"
	                        end if

                            if projektcompltFakbare > 100 then
	                        showprojektcompltFakbare = projektcompltFakbare
	                        projektcompltFakbare = 100

	                        else
	                        showprojektcompltFakbare = projektcompltFakbare
	                        projektcompltFakbare = projektcompltFakbare
	
	                        end if
                            
                            preconditions_met = oRec("preconditions_met")                             
                        %>
                        <tr>
                            <td>
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
                                <a href="jobs.asp?func=red&jobid=<%=oRec("id") %>"><%=oRec("jobnavn") &" ("& oRec("jobnr") & ")" %></a>
                                <%else %>
                                <%=oRec("jobnavn") &" ("& oRec("jobnr") & ")" %>
                                <%end if %>


                                <%
                                if oRec("fastpris") = 1 then
                                strFasptpris = job_txt_324
                                else
                                strFasptpris = job_txt_325
                                end if
                                %>
                                <br />
                                <span style="color:green; font-size:10px;">(<%=strFasptpris %>)</span>
                                <!-- Her skal mere info ind se linje 8345 i gamle jobs -->

                                <span style="color:#999999; font-size:9px;"><i>
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

                                if level = 1 then
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
                                <span style="color:#8caae6; font-size:9px;"><br /><%=job_txt_253 %>: <%=kpersNavn %></span>
                                <%end if 
                                    

                                '**** DATO oversætelser ****
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
                                <br><span style="font-size:9px;"><%=day(oRec("jobstartdato"))&". "& jobstartdatoMonthTxt &" "& year(oRec("jobstartdato")) %> - <%=day(oRec("jobslutdato"))&". "& jobslutdatoMonthTxt &" "& year(oRec("jobslutdato")) %></span>
                                

                                <% 
	                            if len(trim(oRec("rekvnr"))) <> 0 then
                                %>
                                <span style="font-size:9px; color:#999999;">
                                <%
	                            Response.Write "<br>"&job_txt_254&": "& oRec("rekvnr") 
                                %>
                                 </span>
                                <%
	                            end if
                                
                                if len(trim(oRec("aftnavn"))) <> 0 then
                                %><span style="font-size:9px; background-color:#FFFFe1; color:#000000;"><%
	                            Response.Write "<br>"&job_txt_255&": "& oRec("aftnavn") 
                                %>
                                 </span>
                                <%
	                            end if 
	                           
	                            if oRec("jobstatus") = 3 then
                                %><span style="font-size:9px; background-color:#ffdfdf; color:#000000;"><%
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

                            <td>
                               <!-- <%=job_txt_259 %> --> 
                                <%
                                '***** Aktiviteter *****'
                
                                x = 0
                                strAktnavn = "" 
                                lastFase = ""
                                strSQL2 = "SELECT id, navn, fase, aktstatus FROM aktiviteter WHERE job = "& oRec("id") & " AND aktstatus = 1 ORDER BY fase, sortorder, navn"
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
        
                                strAktnavn = strAktnavn & "<br>"


	                            oRec5.movenext
	                            wend
	                            oRec5.close
	                            Antal = x %>
                                
                                <div style="font-size:10px; width:120px; overflow:auto; color:#999999; padding-right:5px; max-height:80px;">           
                                <%=strAktnavn %></div>

                                <span style="font-size:9px;">
                                    <%if editok = 1 then%>
                                    ialt <a href="#" onclick="Javascript:window.open('../timereg/aktiv.asp?menu=job&jobid=<%=oRec("id")%>&jobnavn=<%=oRec("jobnavn")%>&rdir=job3&nomenu=1', '', 'width=1354,height=800,resizable=yes,scrollbars=yes')" class=vmenu><%=Antal%></a>
		                            <%else%>
		                            <b>(<%=Antal%>)</b>
		                            <%end if%>
                                </span>
                            </td>

                            <!-- Forretningsområder --> 
                            <td style="max-width:150px;">                             
                                <%call forretningsomrJobId(oRec("id")) %>

                                <%=strFomr_navn %> &nbsp
                            </td>

                            <td>
                                <%if proaf <> 0 then %>				

		                        <div style="font-size:9px; color:#999999,"><%=job_txt_260 %>:
		                        <span style="width:<%=cint(left(projektcomplt, 3))%>px; background-color:<%=bgdiv%>; height:10px; padding:2px 0px 2px 2px; font-size:9px;">
		                        <%if showprojektcomplt > 0 then%>
		                        <%=formatpercent(showprojektcomplt/100, 0)%>
		                        <%end if%>
		                        </span>
                                </div>

                                <div style="font-size:9px; color:#999999;"><%=job_txt_617 &":" %>
                                <span style="width:<%=cint(left(projektcompltFakbare, 3))%>px; background-color:#CCCCCC; color:#000000; height:10px; padding:2px 0px 2px 2px; font-size:9px;">
		                        <%if showprojektcompltFakbare > 0 then%>
		                        <%=formatpercent(showprojektcompltFakbare/100, 0)%>
		                        <%end if%>
		                        </span>
                                </div>

                                <%end if %>
                            </td>

                            <%call valutakode_fn(oRec("jo_valuta")) %>
		                    <td align="right">
		                    <%=formatnumber(oRec("jo_bruttooms"), 2) &" "& valutaKode_CCC %> 
		                    </td>

                            <td align=right valign=top style="padding:4px 15px 4px 4px; font-size:9px;"><%=job_txt_261 %>: <%=formatnumber(budgettimertot)%><br>
		                    <%=job_txt_239 %>:  <%=formatnumber(proaf)%><br>
                            <span style="color:#999999;"><%=job_txt_262 %>: <%=formatnumber(realfakbare, 2) %></span> <br />
                            -----------------------<br />
                            <%=job_txt_263 %> = <%=formatnumber(restt)%><br />
                            <span style="color:#999999;"><%=formatnumber(resttfakbare) %></span>
                            =====================
		                    </td>                           

                           <!-- <td style="font-size:9px;">
                                <%if editok = 1 then %>
                                <a href="../timereg/job_print.asp?id=<%=oRec("id") %>" class=rmenu>Print / PDF</a><br />
		                        <a href="../timereg/job_kopier.asp?func=kopier&id=<%=oRec("id")%>&fm_kunde=<%=oRec("kid") %>" class=rmenu><%=job_txt_264 %></a><br />
                                <a href="../timereg/materialer_indtast.asp?id=<%=oRec("id")%>&fromsdsk=0&aftid=<%=oRec("serviceaft")%>" target="_blank" class=rmenu><%=job_txt_265 %></a>

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
                                <br />
                                <a href="../timereg/joblog.asp?nomenu=1&FM_job=<%=oRec("id") %>&FM_kunde=<%=oRec("kid")%>&FM_jobsog=<%=oRec("jobnr")%>&viskunabnejob0=<%=viskunabnejob0%>&viskunabnejob1=<%=viskunabnejob1%>&viskunabnejob2=<%=viskunabnejob2%>&<%=dtlink %>" class=rmenu target="_blank"><%=job_txt_266&" " %></a><br />
                                <a href="../timereg/jobprintoverblik.asp?menu=job&id=<%=oRec("id")%>&media=printjoboverblik" class=rmenu target="_blank"><%=job_txt_267 %> </a><br>
                                <a href="../timereg/timereg_akt_2006.asp?FM_kontakt=<%=oRec("kid")%>&FM_ignorer_projektgrupper=1&jobid=<%=oRec("id")%>&FM_sog_job_navn_nr=<%=oRec("jobnr")%>&usemrn=<%=session("mid")%>&showakt=1&fromsdsk=1" target="_blank" class=rmenu><%=job_txt_268 &" " %> </a><br />
          
                                <%if cint(oRec("useasfak")) <= 2 then 'links skal fixes - feltet her skal hentets %>
                                <a href="../timereg/erp_opr_faktura_fs.asp?FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&<%=dtlink %>&reset=1&FM_usedatokri=1" target="_blank" class=rmenu><%=job_txt_269 %> </a>
		                        <%end if %>

                                <%else %>
                                &nbsp
                                <%end if %>
                            </td> -->

                            <td class=lille valign=top style="padding-right:2px;">
		                    <div style="position:relative; top:0px, left:0px; width:160px; height:80px; overflow:auto; padding:5px;">
                            <table cellspacing=0 border=0 cellpadding=0 width=100%>
		
		                    <%
		
		                            '** Findes der fakturaer på job kan det ikke slettes **'
			    
			                        deleteok = 0
			                        faktotbel = 0
                                    fakturaBel_tot = 0
			                        strSQLffak = "SELECT f.fid, f.faknr, f.aftaleid, f.faktype, f.jobid, f.fakdato, f.beloeb, "_
			                        &" f.faktype, f.kurs, SUM(fd.aktpris) AS aktbel, brugfakdatolabel, labeldato, fakadr FROM fakturaer f "_
			                        &" LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid AND fd.enhedsang <> 3)"_
			                        &" WHERE jobid = " & oRec("id") & ""_
			                        &" AND aftaleid = 0 AND shadowcopy = 0"_
			                        &" GROUP BY f.fid ORDER BY f.fakdato DESC"
			                        'Response.Write "strSQLffak<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>" & strSQLffak 
			                        'Response.flush
			                        f = 0
			                        oRec3.open strSQLffak, oConn, 3
			                        while not oRec3.EOF 
			        

                                        if oRec3("brugfakdatolabel") = 1 then '** Labeldato
                                        fakDato = job_txt_623&": "& oRec3("labeldato") & "<br><span style=""color:#999999; size:9px,"">"& oRec3("fakdato") &"</span>"
                                        else
                                        fakDato = job_txt_622&": "& oRec3("fakdato")
                                        end if
			       
			          
                                          %>
                                          <tr><td class=lille valign=top style="font-size:9px;">
                                          <%
			                            if cdate(oRec3("fakdato")) >= cdate("01-01-2006") AND editok = 1 then%>
			                            <a href="../timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id=<%=oRec3("fid")%>&FM_jobonoff=<%=FM_jobonoffval%>&FM_kunde=<%=oRec3("fakadr")%>&FM_job=<%=oRec3("jobid")%>&FM_aftale=<%=oRec3("aftaleid")%>&fromfakhist=1" class="lgron" target="_blank"><b><%=oRec3("faknr") &"</b>"%></a></td>
                                        <td align=right class=lille valign=top style="font-size:9px;"> <%=fakDato %> <br />
                                        <%else%>
                                        <b><%=oRec3("faknr") &"</b></td><td align=right class=lille valign=top> "& fakDato %><br />
                                        <% end if
                    

                                          %>

                                          </td></tr>
                                          <%
                    
                     
	                                    call beregnValuta(minus&(oRec3("beloeb")),oRec3("kurs"),100)
                                        if oRec3("faktype") <> 1 then
                                        belobGrundVal = valBelobBeregnet
                                        else
                                        belobGrundVal = -valBelobBeregnet
                                        end if 
                    

                                            if cDate(oRec3("fakdato")) < cDate("01-06-2010") AND (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk") then
                                            belobKunTimerStk = belobKunTimerStk + oRec3("beloeb")
                                            else
       
                    
	                                            '** Kun aktiviteter timer, enh. stk. IKKE materialer og KM
	                                            call beregnValuta(minus&(oRec3("aktbel")),oRec3("kurs"),100)
                                                if oRec3("faktype") <> 1 then
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
            
                                <!-- <span class="qmarkhelp" id="qm0001" style="font-size:11px; color:#999999; font-weight:bolder;">?</span><span id="qmarkhelptxt_qm0001" style="visibility:hidden; color:#999999; display:none; padding:3px; z-index:4000;">(faktureret beløb - (materialeforbrug + km)) / timer realiseret</span>-->
	        
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

                            <td align=right class=lille style="white-space:nowrap; padding:4px 4px 0px 4px;" valign=top>
                            <% for p = 1 to 10
        
                            pgid = oRec("projektgruppe"&p)

                            if pgid <> 1 then
                                call prgNavn(pgid, 200)
                                Response.Write left(prgNavnTxt, 20) & "<br>"
                            end if
        
                            next %>
		   
		                    </td>

                            <td style="text-align:center">
                                <span style="font-size:125%" id="modal_<%=oRec("id") %>" class="fa fa-link picmodal"></span>

                                <div id="myModal_<%=oRec("id") %>" class="modal">

                                    <!-- Modal content -->
                                    <div class="modal-content" style="text-align:left">
                                        
                                        <%if editok = 1 then %>
                                        <h6><a href="../timereg/job_print.asp?id=<%=oRec("id") %>" class=rmenu>Print / PDF</a></h6>
		                                <h6><a href="../timereg/job_kopier.asp?func=kopier&id=<%=oRec("id")%>&fm_kunde=<%=oRec("kid") %>" class=rmenu><%=job_txt_264 %></a></h6>
                                        <h6><a href="../timereg/materialer_indtast.asp?id=<%=oRec("id")%>&fromsdsk=0&aftid=<%=oRec("serviceaft")%>" target="_blank" class=rmenu><%=job_txt_265 %></a></h6>

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
                                        
                                        <h6><a href="../timereg/joblog.asp?nomenu=1&FM_job=<%=oRec("id") %>&FM_kunde=<%=oRec("kid")%>&FM_jobsog=<%=oRec("jobnr")%>&viskunabnejob0=<%=viskunabnejob0%>&viskunabnejob1=<%=viskunabnejob1%>&viskunabnejob2=<%=viskunabnejob2%>&<%=dtlink %>" class=rmenu target="_blank"><%=job_txt_266&" " %></a></h6>
                                        <h6><a href="../timereg/jobprintoverblik.asp?menu=job&id=<%=oRec("id")%>&media=printjoboverblik" class=rmenu target="_blank"><%=job_txt_267 %> </a></h6>
                                        <h6><a href="../timereg/timereg_akt_2006.asp?FM_kontakt=<%=oRec("kid")%>&FM_ignorer_projektgrupper=1&jobid=<%=oRec("id")%>&FM_sog_job_navn_nr=<%=oRec("jobnr")%>&usemrn=<%=session("mid")%>&showakt=1&fromsdsk=1" target="_blank" class=rmenu><%=job_txt_268 &" " %> </a></h6>
          
                                        <%if cint(oRec("useasfak")) <= 2 then 'links skal fixes - feltet her skal hentets %>
                                        <h6><a href="../timereg/erp_opr_faktura_fs.asp?FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&<%=dtlink %>&reset=1&FM_usedatokri=1" target="_blank" class=rmenu><%=job_txt_269 %> </a></h6>
		                                <%end if %>

                                        <%else %>
                                        &nbsp
                                        <%end if %>

                                    </div>

                                </div>
                            </td>
                            
                            <td style="text-align:center" ><!-- Status -->

                                <%
                                    select case cdbl(oRec("jobstatus"))
                                    case 1
                                    statusCHBAR1 = "SELECTED"
                                    statusCHBAR2 = ""
                                    statusCHBAR3 = ""
                                    statusCHBAR4 = ""
                                    statusCHBAR0 = ""
                                    statusTXT = "Aktiv"
                                    case 2
                                    statusCHBAR1 = ""
                                    statusCHBAR2 = "SELECTED"
                                    statusCHBAR3 = ""
                                    statusCHBAR4 = ""
                                    statusCHBAR0 = ""
                                    statusTXT = "Pasiv"
                                    case 3
                                    statusCHBAR1 = ""
                                    statusCHBAR2 = ""
                                    statusCHBAR3 = "SELECTED"
                                    statusCHBAR4 = ""
                                    statusCHBAR0 = ""
                                    statusTXT = "Tilbud"
                                    case 4
                                    statusCHBAR1 = ""
                                    statusCHBAR2 = ""
                                    statusCHBAR3 = ""
                                    statusCHBAR4 = "SELECTED"
                                    statusCHBAR0 = ""
                                    statusTXT = "Gen. syn"
                                    case 0
                                    statusCHBAR1 = ""
                                    statusCHBAR2 = ""
                                    statusCHBAR3 = ""
                                    statusCHBAR4 = ""
                                    statusCHBAR0 = "SELECTED"
                                    statusTXT = "Lukket"

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

                            <td style="text-align:center">
                                <input class="job_bulkCHB" name="job_bulkCHB" id="<%=oRec("id") %>" type="checkbox" />
                            </td>

                        </tr>
                        <%
                            oRec.movenext
                            wend
                            oRec.close 
                        %>
                    </tbody>
                </table>

            </div>
        </div>

<script type="text/javascript">


$(".picmodal").click(function() {

    var modalid = this.id
    var idlngt = modalid.length
    var idtrim = modalid.slice(6, idlngt)

    //var modalidtxt = $("#myModal_" + idtrim);
    var modal = document.getElementById('myModal_' + idtrim);

    modal.style.display = "block";

    window.onclick = function (event) {
        if (event.target == modal) {
            modal.style.display = "none";
        }
    }

});


</script>


    </div>
</div>
<%end select %>
<!--#include file="../inc/regular/footer_inc.asp"-->