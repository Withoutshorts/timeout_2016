

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--include file="inc/budget_firapport_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->


<script type="text/javascript" src="js/libs/excanvas.compiled.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.orderBars.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.pie.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.resize.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.stack.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.tooltip.min.js"></script>
<script type="text/javascript" src="js/demos/flot/stacked-vertical.js"></script>
<script type="text/javascript" src="js/demos/flot/donut_saelger.js"></script>
<script type="text/javascript" src="js/demos/flot/vertical_saelger.js"></script>
<script type="text/javascript" src="js/demos/flot/scatter.js"></script>
<script type="text/javascript" src="js/demos/flot/line.js"></script>



<%call menu_2014 %>


<% if len(trim(request("FM_periode"))) <> 0 then
   usePeriode = request("FM_periode")
   else
   usePeriode = 1 'måned
   end if 
    

    usePeriodeSel0 = ""
    usePeriodeSel1 = ""
    usePeriodeSel2 = ""
    usePeriodeSel3 = ""

    select case usePeriode
    case 0
    usePeriodeSel0 = "CHECKED"
    case 1
    usePeriodeSel1 = "CHECKED"
    case 2
    usePeriodeSel2 = "CHECKED"
    case 3
    usePeriodeSel3 = "CHECKED"
    case else
    usePeriodeSel0 = "CHECKED"
    end select

    
    if len(trim(request("FM_aar"))) <> 0 then
    useYear = request("FM_aar")
    else
    useYear = year(now)
    end if
    
    if len(trim(request("FM_mth"))) <> 0 then
    useMonth = request("FM_mth")
    else
    useMonth = month(now)
    end if

    mthSel1 = ""
    mthSel2 = ""
    mthSel3 = ""
    mthSel4 = ""
    mthSel5 = ""
    mthSel6 = ""
    mthSel7 = ""
    mthSel8 = ""
    mthSel9 = ""
    mthSel10 = ""
    mthSel11 = ""
    mthSel12 = ""

    select case useMonth
    case 1
    mthSel1 = "SELECTED"
    case 2
    mthSel2 = "SELECTED"
    case 3
    mthSel3 = "SELECTED"
    case 4
    mthSel4 = "SELECTED"
    case 5
    mthSel5 = "SELECTED"
    case 6
    mthSel6 = "SELECTED"
    case 7
    mthSel7 = "SELECTED"
    case 8
    mthSel8 = "SELECTED"
    case 9
    mthSel9 = "SELECTED"
    case 10
    mthSel10 = "SELECTED"
    case 11
    mthSel11 = "SELECTED"
    case 12
    mthSel12 = "SELECTED"
    end select

    if len(trim(request("FM_kv"))) <> 0 then
    useQuater = request("FM_kv")
    else
    useQuater = 1
    end if

    useQuaterSel1 = ""
    useQuaterSel2 = ""
    useQuaterSel3 = ""
    useQuaterSel4 = ""

    select case useQuater
    case 1
    useQuaterSel1 = "SELECTED"
    case 2
    useQuaterSel2 = "SELECTED"
    case 3
    useQuaterSel3 = "SELECTED"
    case 4
    useQuaterSel4 = "SELECTED"
    case else
    useQuaterSel1 = "SELECTED"
    end select

    if len(trim(request("FM_medarb"))) <> 0 then
    usemrn = request("FM_medarb")
    else
    usemrn = session("mid") 
    end if


    jobidsStr = " AND (j.id <> 0 "
    if len(trim(request("FM_job"))) <> 0 then
    jobid = request("FM_job")
        
            if instr(jobid, ",") <> 0 then
            jobidsStr = " AND (j.id = 0 "
            jobids = split(request("FM_job"), ", ") 

                for d = 0 to UBOUND(jobids)

                    jobidsStr = jobidsStr & " OR j.id = " & jobids(d)
                
                next

            '** Nulstil jobid
            jobid = 0
            end if
    
    else
    jobid = 0
    end if

    jobidsStr = jobidsStr & ")"



    %>

<div class="wrapper">
    <div class="content">
    
    <div class="container">
 
        
         <div class="portlet">
        <h3 class="portlet-title"><u>Sælger Dashboard</u></h3>

        <div class="well">
              <form action="saelgerdashboard.asp?sogsubmitted=1" method="POST">
                <div class="row">
                <div class="col-lg-2">
                <h4 class="panel-title-well">Søgefilter</h4></div>
                <div class="col-lg-10">&nbsp</div>
                 
                </div>
                  
                <div class="row">
                <div class="col-lg-2">Navn: (kun sælger grp.)</div>
                <div class="col-lg-2"><input type="radio" name="FM_periode" value="1" onchange="submit()" <%=usePeriodeSel1 %> /> Måned:</div>
                <div class="col-lg-2"><input type="radio" name="FM_periode" value="3" onchange="submit()" <%=usePeriodeSel3 %> /> År:</div>
             
                <div class="col-lg-4 txt">Kunder & Projekter:</div>
                

            </div>
                  <div class="row">
                            <div class="col-lg-2">
                                 <select name="FM_medarb" class="form-control input-small" onchange="submit()">
                                <%
                                    select case lto
                                    case "intranet - local"
                                    pgrp = 10
                                    case "wilke"
                                    pgrp = 2
                                    case else
                                    pgrp = 10
                                    end select
                                    
                                    call medarbiprojgrp(pgrp, session("mid"), 0, -1)
                                    
                                    strSQLm = "SELECT mid, mnavn, init FROM medarbejdere WHERE mansat = 1 AND (mid = 0 "& medarbgrpIdSQLkri &") ORDER BY mnavn"

                                   

                                    oRec3.open strSQLm, oConn, 3
                                    while not oRec3.EOF 

                                    if cint(usemrn) = cint(oRec3("mid")) then
                                    mSel = "SELECTED"
                                    else
                                    mSel = ""
                                    end if
                                    %>
                                    <option value="<%=oRec3("mid") %>" <%=mSel %>><%=oRec3("mnavn")& " ["& oRec3("init") &"]" %></option>
                                    <%
                                    oRec3.movenext
                                    wend 
                                    oRec3.close
                                 %>
                  
                                </select>
                

                      </div>
                      <div class="col-lg-2">
                          <select name="FM_mth" class="form-control input-small" onchange="submit()">
                          <option value="1" <%=mthSel1 %>>Jan</option>
                          <option value="2" <%=mthSel2 %>>Feb</option>
                          <option value="3" <%=mthSel3 %>>Mar</option>
                          <option value="4" <%=mthSel4 %>>Apr</option>
                          <option value="5" <%=mthSel5 %>>Maj</option>
                          <option value="6" <%=mthSel6 %>>Jun</option>
                          <option value="7" <%=mthSel7 %>>Jul</option>
                          <option value="8" <%=mthSel8 %>>Aug</option>
                          <option value="9" <%=mthSel9 %>>Sep</option>
                          <option value="10" <%=mthSel10 %>>Okt</option>
                          <option value="11" <%=mthSel11 %>>Nov</option>
                          <option value="12" <%=mthSel12 %>>Dec</option>
                   
                        </select>
                      </div>
                      <div class="col-lg-2">
                        <select name="FM_aar" class="form-control input-small" onchange="submit()">
                            <%
                           yearSt = datePart("yyyy", dateAdd("yyyy", -3, now), 2,2) 
                           yearEnd = yearSt + 4 
     
                           for y = yearSt TO yearEnd  
                        
                        
                               if cint(y) = cint(useYear) then
                                ySel = "SELECTED"
                               else
                                ySel = ""
                                end if %>
                          <option value="<%=y %>" <%=ySel%>><%=y %></option>
                          <%next %>
                 
                        </select>
                    </div>
                     
                 
                   <div class="col-lg-4">
                      
                       <%           
                           
                           strSQLKproj = "SELECT kkundenavn, kkundenr, jobnavn, jobnr, j.id AS jid, kid FROM job AS j LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                           &" WHERE risiko > -1 AND (j.salgsans1 = "& usemrn &" OR j.salgsans2 = "& usemrn &" OR j.salgsans3 = "& usemrn &" OR j.salgsans4 = "& usemrn &" OR j.salgsans5 = "& usemrn &") "_
                           &" ORDER BY kkundenavn, jobnavn" 
                           '&" AND jobstartdato > '2015-01-01'"_
                           'response.write strSQLKproj
                           'response.flush

                           %>
                              <select name="FM_job" class="form-control input-small" multiple size="7">
                                  <%if cint(jobid) =  0 then
                                    jSel = "SELECTED"
                                    else
                                    jSel = ""
                                    end if %>

                            <option value="0" <%=jSel%>>Alle</option>
                       <%
                                    lastKid = 0
                                   k = 0
                                    oRec3.open strSQLKproj, oConn, 3
                                    while not oRec3.EOF 

                                    if lastKid <> oRec3("kid") then
                                  
                                    if k > 0 then %>
                                    <option value="0" DISABLED></option>
                                    <%end if %>

                                     <option value="0" DISABLED><%=oRec3("kkundenavn")%></option>
                                    <%
                                    end if


                                    if cint(jobid) = cint(oRec3("jid")) OR instr(jobidsStr, "OR j.id = "& oRec3("jid")) <> 0 then
                                    jSel = "SELECTED"
                                    else
                                    jSel = ""
                                    end if
                                    %>
                                    <option value="<%=oRec3("jid") %>" <%=jSel %>><%=oRec3("jobnavn") & " ("& oRec3("jobnr") &")"%></option>
                                    <%

                                    lastKid = oRec3("kid")
                                    k = k + 1
                                    oRec3.movenext
                                    wend 
                                    oRec3.close
                           
                           %>
                            </select>
                       </div>

                      <div class="col-lg-1">
                           <button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Søg >></b></button>
                      </div>

                 
                   </div>
            </form>
            </div>



        <%
         
         call akttyper2009(2)

         call meStamdata(usemrn)
        
            
        select case usePeriode
        case 0 'uge
        strDatoStartKri = useYear & "-9-1"
         strDatoEndKri = useYear & "-11-1"
        case 1 'month
        strDatoStartKri = useYear & "-"& useMonth &"-1"
        strDatoEndKri = dateAdd("m", 1, strDatoStartKri)
        strDatoEndKri = dateAdd("d", -1, strDatoEndKri)
        strDatoEndKri = year(strDatoEndKri) &"-"& month(strDatoEndKri) &"-"& day(strDatoEndKri)
        case 2 'quater

            select case useQuater
            case 1 
            strDatoStartKri = useYear & "-1-1"
            strDatoEndKri = useYear & "-3-31"
            case 2
            strDatoStartKri = useYear & "-4-1"
            strDatoEndKri = useYear & "-6-30"
            case 3
            strDatoStartKri = useYear & "-7-1"
            strDatoEndKri = useYear & "-9-30"
            case 4
            strDatoStartKri = useYear & "-10-1"
            strDatoEndKri = useYear & "-12-31"
            end select

        case 3 'year
        strDatoStartKri = useYear & "-1-1"
         strDatoEndKri = useYear & "-12-31"
        case else 'mth
        strDatoStartKri = useYear & "-9-1"
        strDatoEndKri = useYear & "-11-1"
        end select
           

        '*** JobSQLkri
        if jobid <> 0 then 'Arr eller valgt job
        JobSQLkri = " AND j.id = " & jobid 
        else
        JobSQLkri = jobidsStr
        end if


            '*** DB1 + DB2 - Valgte Medarb
            strSQLjobDB = "SELECT jobnavn, jobnr, SUM(budgettimer) AS budgettimer, SUM(jo_bruttofortj) AS jo_bruttofortj, SUM(jo_dbproc) AS jo_dbproc, SUM(jo_udgifter_ulev) AS jo_udgifter_ulev, SUM(jo_udgifter_intern) AS jo_udgifter_intern FROM job AS j WHERE risiko > -1 "& JobSQLkri & ""
            
            'response.write strSQLjobDB
            'response.flush
            
            jo_bruttofortj = 0
            jo_dbproc = 0
            jo_db2 = 0
            jo_db2_proc = 0
            budgettimer = 0
            oRec.open strSQLjobDB, oConn, 3
            if not oRec.EOF then

            jo_bruttofortj = oRec("jo_bruttofortj")
            jo_dbproc = oRec("jo_dbproc")

            jo_udgifter_intern = oRec("jo_udgifter_intern")
            jo_udgifter_ulev = oRec("jo_udgifter_ulev")
            jo_db2 = (jo_bruttofortj - jo_udgifter_intern)
            jo_db2_proc = (jo_db2/jo_bruttofortj)*100

            budgettimer = oRec("budgettimer")

            end if
            oRec.close


             '*** DB1 + DB2 - Gruppen
            strSQLjobDB = "SELECT jobnavn, jobnr, SUM(jo_bruttofortj) AS jo_bruttofortj, SUM(jo_dbproc) AS jo_dbproc, SUM(jo_udgifter_ulev) AS jo_udgifter_ulev, SUM(jo_udgifter_intern) AS jo_udgifter_intern FROM job AS j WHERE risiko > -1 AND jobstartdato >= '"& strDatoStartKri &"'"
            
            'response.write strSQLjobDB
            'response.flush
            
            grp_jo_bruttofortj = 0
            grp_jo_dbproc = 0
            grp_jo_db2 = 0
            grp_jo_db2_proc = 0
            oRec.open strSQLjobDB, oConn, 3
            if not oRec.EOF then

            grp_jo_bruttofortj = oRec("jo_bruttofortj")
            grp_jo_dbproc = oRec("jo_dbproc")

            grp_jo_udgifter_intern = oRec("jo_udgifter_intern")
            grp_jo_udgifter_ulev = oRec("jo_udgifter_ulev")
            grp_jo_db2 = (grp_jo_bruttofortj - grp_jo_udgifter_intern)

            if grp_jo_bruttofortj <> 0 then
            grp_jo_db2_proc = (grp_jo_db2/grp_jo_bruttofortj)*100
            else
            grp_jo_db2_proc = 0
            end if

            end if
            oRec.close



            '*** Henter jobnr på valgte job ****'
             '*** DB1 + DB2 - Valgte Medarb
            strSQLjobnrTimer = " AND (tjobnr = 0"
            strSQLjobnr = "SELECT jobnr FROM job AS j WHERE risiko > -1 "& JobSQLkri & ""
            oRec.open strSQLjobnr, oConn, 3
            while not oRec.EOF 

            strSQLjobnrTimer = strSQLjobnrTimer & " OR tjobnr = '" & oRec("jobnr") & "'"

            oRec.movenext
            wend 
            oRec.close
            strSQLjobnrTimer = strSQLjobnrTimer & ")"








        '** Real. timer ALLE MEDARB - Kun fakturerbare - Valgte job for slagsansvarlige --> Kun Lukkede job??
         strSQlmtimer = "SELECT SUM(timer) AS sumtimer FROM timer AS t WHERE tmnr <> 0 "& strSQLjobnrTimer &" AND ("& aty_sql_realHoursFakbar &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'"
       

        faktureringsgrad = budgettimer

        if budgettimer <> 0 then
        budgettimerProc = budgettimer
        else 
        budgettimerProc = 1
        end if
        
        'response.write strSQlmtimer
        'response.flush

        medarbSumtimer = 0

         oRec.open strSQlmtimer, oConn, 3
         if not oRec.EOF then

            faktureringsgrad = budgettimer - oRec("sumtimer")
            medarbSumtimer = oRec("sumtimer")
           
         end if
         oRec.close 

            if medarbSumtimer <> 0 then
            
                if faktureringsgrad > 0 then
                faktureringsgrad_proc = 100 + (medarbSumtimer / budgettimerProc) * 100
                faktureringsgradCol = "success"
                faktureringsgradFortegn = "+"
                else
                faktureringsgrad_proc = (budgettimerProc / medarbSumtimer) * 100
                faktureringsgradCol = "danger"
                faktureringsgradFortegn = "-"
                end if

            else

                faktureringsgrad_proc = 0
                faktureringsgradCol = ""
                faktureringsgradFortegn = ""

            end if






         '** Utilazation Real. timer valgt MEDARB - Kun fakturerbare ---> Kun Lukkede job??

         '*** Kun FOMR ***
            '** Pre = 2
            '** Pro = 4
            '** Post = 6

        
        '** Normtimer
        natalDage = dateDiff("d", strDatoStartKri, strDatoEndKri, 2,2)
        call normtimerPer(usemrn, strDatoStartKri, natalDage, 0)
        ntimPer = formatnumber(ntimPer,0)


        '**** PRE ****
         strSQlfomr = "SELECT for_aktid FROM fomr_rel WHERE for_fomr = 2 GROUP BY for_aktid"

         strSQLaktidsfomr = " AND (taktivitetid = 0 "
         oRec.open strSQlfomr, oConn, 3
         while not oRec.EOF

            strSQLaktidsfomr = strSQLaktidsfomr & " OR taktivitetid = "& oRec("for_aktid")
           
         oRec.movenext
         wend
         oRec.close 

        strSQLaktidsfomr = strSQLaktidsfomr & ")"


       strSQlmtimerUtil = "SELECT SUM(timer) AS sumtimer FROM timer AS t WHERE tmnr = "& usemrn &" "& strSQLjobnrTimer &" AND ("& aty_sql_realHoursFakbar &") "& strSQLaktidsfomr &" AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' GROUP BY tmnr"
      
        'response.write strSQlmtimerUtil
        'response.flush

        medarbSumtimerUtilpre = 0

         oRec.open strSQlmtimerUtil, oConn, 3
         if not oRec.EOF then

            
            medarbSumtimerUtilpre = oRec("sumtimer")

         oRec.movenext
         end if
         oRec.close 


         '**** PROD ****
         strSQlfomr = "SELECT for_aktid FROM fomr_rel WHERE for_fomr = 4 GROUP BY for_aktid"

         strSQLaktidsfomr = " AND (taktivitetid = 0 "
         oRec.open strSQlfomr, oConn, 3
         while not oRec.EOF

            strSQLaktidsfomr = strSQLaktidsfomr & " OR taktivitetid = "& oRec("for_aktid")
           
         oRec.movenext
         wend
         oRec.close 

        strSQLaktidsfomr = strSQLaktidsfomr & ")"


       strSQlmtimerUtil = "SELECT SUM(timer) AS sumtimer FROM timer AS t WHERE tmnr = "& usemrn &" "& strSQLjobnrTimer &" AND ("& aty_sql_realHoursFakbar &") "& strSQLaktidsfomr &" AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' GROUP BY tmnr"
      
        'response.write strSQlmtimerUtil
        'response.flush

        medarbSumtimerUtilprod = 0

         oRec.open strSQlmtimerUtil, oConn, 3
         if not oRec.EOF then

            medarbSumtimerUtilprod = oRec("sumtimer")

         oRec.movenext 
         end if
         oRec.close 


         '**** Post ****
         strSQlfomr = "SELECT for_aktid FROM fomr_rel WHERE for_fomr = 6 GROUP BY for_aktid"

         strSQLaktidsfomr = " AND (taktivitetid = 0 "
         oRec.open strSQlfomr, oConn, 3
         while not oRec.EOF

            strSQLaktidsfomr = strSQLaktidsfomr & " OR taktivitetid = "& oRec("for_aktid")
         
         oRec.movenext
         wend
         oRec.close 

        strSQLaktidsfomr = strSQLaktidsfomr & ")"


       strSQlmtimerUtil = "SELECT SUM(timer) AS sumtimer FROM timer AS t WHERE tmnr = "& usemrn &" "& strSQLjobnrTimer &" AND ("& aty_sql_realHoursFakbar &") "& strSQLaktidsfomr &" AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' GROUP BY tmnr"
      
        'response.write strSQlmtimerUtil
        'response.flush

        medarbSumtimerUtilpost = 0

         oRec.open strSQlmtimerUtil, oConn, 3
         if not oRec.EOF then

            
            medarbSumtimerUtilpost = oRec("sumtimer")
         
         oRec.movenext  
         end if
         oRec.close 



        '***** Bregn Util Proc Norm / Pro, Post

        if (ntimPer) <> 0 then
        'medarbSumtimerUtilprodUtilProc = (medarbSumtimerUtilprod / (medarbSumtimerUtilpost + medarbSumtimerUtilpre)) * 100
        medarbSumtimerUtilprodUtilProc = (medarbSumtimerUtilprod / ntimPer) * 100
        else
        medarbSumtimerUtilprodUtilProc = 0
        end if


        if (ntimPer) <> 0 then
            
                
                medarbSumtimerUtilprodUtilProc = (medarbSumtimerUtilprod / ntimPer) * 100
                
                if cdbl(medarbSumtimerUtilprodUtilProc) > 33 then 
                utilGradCol = "success"
                utilGradFortegn = "+"
                utilLevel = (medarbSumtimerUtilprodUtilProc - 33)
                else
                utilGradCol = "danger"
                utilGradFortegn = "-"
                utilLevel = (33 - medarbSumtimerUtilprodUtilProc)
                end if

                utilLevel = formatnumber(utilLevel, 0)

            else

                medarbSumtimerUtilprodUtilProc = 0
                faktureringsgradCol = ""
                faktureringsgradFortegn = ""
                utilLevel = 0

            end if
         
            
            donut100 = formatnumber((medarbSumtimerUtilpre/1 + medarbSumtimerUtilprod/1 + medarbSumtimerUtilpost/1), 0)
            medarbSumtimerUtilpreProc = formatnumber((medarbSumtimerUtilpre/donut100) * 100, 0)
            medarbSumtimerUtilprodProc = formatnumber((medarbSumtimerUtilprod/donut100) * 100, 0)
            medarbSumtimerUtilpostProc = formatnumber((medarbSumtimerUtilpost/donut100) * 100, 0)
        %>

             <form>
                
                 <input type="hidden" id="util_pre" value="<%=formatnumber(medarbSumtimerUtilpre, 0) %>" />
                 <input type="hidden" id="util_pre_proc" value="<%=formatnumber(medarbSumtimerUtilpreProc, 0) %>" />

                 <input type="hidden" id="util_prod" value="<%=formatnumber(medarbSumtimerUtilprod, 0) %>" />
                 <input type="hidden" id="util_prod_proc" value="<%=formatnumber(medarbSumtimerUtilprodProc, 0) %>" />

                 <input type="hidden" id="util_post" value="<%=formatnumber(medarbSumtimerUtilpost, 0) %>" />
                 <input type="hidden" id="util_post_proc" value="<%=formatnumber(medarbSumtimerUtilpostProc, 0) %>" />
             </form>


        <div class="portlet-body">

            <div class="row">
                <div class="col-lg-11">
                    <h4 class="portlet-title"><u>Nøgletal</u>
                    </h4>
                </div>
            </div>

            
            <div class="row">
            
            <div class="col-sm-6 col-md-3">
              <div class="row-stat">
                <p class="row-stat-label">DB1</p>
                <h3 class="row-stat-value"><%=formatnumber(jo_bruttofortj,2) %></h3>
                <span class="label label-<%=hoursDBindicatorColor %> row-stat-badge"><%=formatnumber(jo_dbproc, 0)%> %</span> <p><br /></p>

                <p class="row-stat-label">DB2</p>
                <h3 class="row-stat-value"><%=formatnumber(jo_db2,2) %></h3>
                <span class="label label-<%=hoursDBindicatorColor %> row-stat-badge"><%=formatnumber(jo_db2_proc, 0)%> %</span>
              </div> <!-- /.row-stat -->
            </div> <!-- /.col -->

            <div class="col-lg-1">&nbsp</div>
            <div class="col-sm-6 col-md-3">
              <div class="row-stat">
                <p class="row-stat-label">Faktureringsgrad</p>
                <h3 class="row-stat-value"><%=formatnumber(faktureringsgrad_proc,0) %> %</h3> 
                <span class="label label-<%=faktureringsgradCol %> row-stat-badge"><%=faktureringsgradFortegn &" "& formatnumber(faktureringsgrad, 0) %> t.</span>
                 <br /> <span style="font-size:9px; color:#999999;">Budget: <%=formatnumber(budgettimer, 2) %> Real.: <%=formatnumber(medarbSumtimer, 2) %> t.</span>
              </div> <!-- /.row-stat -->
            </div> <!-- /.col -->

            <div class="col-lg-1">&nbsp</div>
            <div class="col-sm-6 col-md-3">
              <div class="row-stat">
                <p class="row-stat-label">Utilization <%=meTxt %></p>
                <h3 class="row-stat-value"><%=formatnumber(medarbSumtimerUtilprodUtilProc, 0) %>%</h3>
                <span class="label label-<%=utilGradCol %> row-stat-badge"><%=utilGradFortegn &" "& formatnumber(utilLevel, 0) %> %</span>
                  <br /> <span style="font-size:9px; color:#999999;">Norm. / Production</span>
                   <!--<br /> <span style="font-size:9px; color:#999999;">Pre.: <%=formatnumber(medarbSumtimerUtilpre, 2) %> Prod.: <%=formatnumber(medarbSumtimerUtilprod, 2) %> Post: <%=formatnumber(medarbSumtimerUtilpost, 2) %> t.</span>-->
                 </div> <!-- /.row-stat -->
            </div> <!-- /.col -->

            </div> <!--Row-->

            <div class="row pad-t20">
                <div class="col-lg-3">
                    <h4 class="portlet-title"><u>Dækningsbidrag</u></h4>
                </div>
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-3">
                    <h4 class="portlet-title"><u>Faktureringsgrad</u></h4>
                </div>
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-3">
                    <h4 class="portlet-title"><u>Utilization</u></h4>
                </div>
            </div>

            </div>
            <div class="row">
                <div class="col-lg-3">  
                <div class="progress-stat">
                    <div class="progress-stat-label">
                        DB1<br /> 
                        <%=meTxt %>:
                    </div>
                    <div class="progress-stat-value">
                        <%=formatnumber(jo_bruttofortj,0) %> DKK<br />
                        <%=formatnumber(jo_dbproc, 0)%>%
                    </div>
                    <div class="progress progress-striped progress-sm active">
                        <div class="progress-bar progress-bar-secondary" role="progressbar" aria-valuenow="<%=formatnumber(jo_dbproc, 0)%>" aria-valuemin="0" aria-valuemax="100" style="width:<%=formatnumber(jo_dbproc, 0)%>%">
                        <span class="sr-only">0</span>
                        </div>
                    </div> 
                </div>
                    
                    <div class="progress-stat">
                    <div class="progress-stat-label">
                        DB1<br />Gruppen:
                    </div>
                    <div class="progress-stat-value">
                        <%=formatnumber(grp_jo_bruttofortj,0) %> DKK<br />
                        <%=formatnumber(grp_jo_dbproc,0) %>%
                    </div>
                    <div class="progress progress-striped progress-sm active">
                        <div class="progress-bar progress-bar-default" role="progressbar" aria-valuenow="<%=formatnumber(grp_jo_dbproc, 0)%>" aria-valuemin="0" aria-valuemax="100" style="width: <%=formatnumber(grp_jo_dbproc, 0)%>%">
                        <span class="sr-only">0</span>
                        </div>
                    </div> 
                </div> 

                    <div class="progress-stat">
                    <div class="progress-stat-label">
                        DB2<br /><%=meTxt %>:
                    </div>
                    <div class="progress-stat-value">
                       <%=formatnumber(jo_db2, 0) %> DKK<br />
                        <%=formatnumber(jo_db2_proc, 0) %>%
                    </div>
                    <div class="progress progress-striped progress-sm active">
                        <div class="progress-bar progress-bar-secondary" role="progressbar" aria-valuenow="<%=formatnumber(jo_db2_proc, 0)%>" aria-valuemin="0" aria-valuemax="100" style="width:<%=formatnumber(jo_db2_proc, 0)%>%">
                        <span class="sr-only">0</span>
                        </div>
                    </div> 
                </div>
                    
                    <div class="progress-stat">
                    <div class="progress-stat-label">
                        DB2<br />Gruppe:
                    </div>
                    <div class="progress-stat-value">
                         <%=formatnumber(grp_jo_db2, 0) %> DKK<br />
                        <%=formatnumber(grp_jo_db2_proc, 0)%>%
                    </div>
                    <div class="progress progress-striped progress-sm active">
                        <div class="progress-bar progress-bar-default" role="progressbar" aria-valuenow="<%=formatnumber(grp_jo_db2_proc, 0)%>" aria-valuemin="0" aria-valuemax="100" style="width: <%=formatnumber(grp_jo_db2_proc, 0)%>%">
                        <span class="sr-only">0</span>
                        </div>
                    </div> 
                         <span style="color:#999999; font-size:9px;">Gruppen udregnes ved alle job med st. dato > valgt dato.</span>
                </div>
               
            </div>   


                    <div class="col-lg-1">&nbsp</div>            
                    <div class="col-lg-3"><div id="vertical-chart" class="chart-holder-250"></div></div>

                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-3">            
                    <div id="donut-chart" class="chart-holder" style="width: 95%;"></div>
                    </div>


            </div>
        </div> 
    </div>


            <div class="container">
                <div class="portlet">
               
                    <div class="portlet-body">
                        
                        <div class="row">
                        <div class="col-lg-5">
                             <h4 class="portlet-title"><u>Aktiviteter</u></h4>
                        </div>
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-5">
                            <h4 class="portlet-title"><u>Interne og Eksterne timer</u></h4>
                        </div>
                        </div>

                        <div class="row">
                        <div class="col-lg-5">
                       <%=meTxt %> 
                        <table class="table table-bordered">
                            <thead>
                                <tr>
                                    <th style="width: 50%">Aktivitet</th>
                                    <th style="width: 20%">Forecast</th>

                                    <th style="width: 20%">Realiseret</th>
                                    <th style="width: 10%"><!--Indikator--></th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                
                                '*** PRE
                                    fctimer = 20

                                    if cdbl(fctimer) > cdbl(medarbSumtimerUtilpre) then
                                    indicatorCol = "yellowgreen"
                                    else
                                    indicatorCol = "crimson"
                                    end if
                                    %>

                                <tr>
                                    <td>Pre.:</td>
                                    <td class="text-right"><%=formatnumber(fctimer, 2) %></td>
                                    <td class="text-right"><%=formatnumber(medarbSumtimerUtilpre, 2) %></td>
                                    <td><div style="background-color:<%=indicatorCol%>; width:10px; height:10px;">&nbsp;</div></td>
                                </tr>


                               

                                <%
                                
                                '*** PROD
                                    fctimer = 20

                                    if cdbl(fctimer) > cdbl(medarbSumtimerUtilprod) then
                                    indicatorCol = "yellowgreen"
                                    else
                                    indicatorCol = "crimson"
                                    end if
                                    %>

                                <tr>
                                    <td>Prod.:</td>
                                    <td class="text-right"><%=formatnumber(fctimer, 2) %></td>
                                    <td class="text-right"><%=formatnumber(medarbSumtimerUtilprod, 2) %></td>
                                    <td><div style="background-color:<%=indicatorCol%>; width:10px; height:10px;">&nbsp;</div></td>
                                </tr>
                               
                                <%
                                
                                '*** POST
                                    fctimer = 20

                                    if cdbl(fctimer) > cdbl(medarbSumtimerUtilpost) then
                                    indicatorCol = "yellowgreen"
                                    else
                                    indicatorCol = "crimson"
                                    end if
                                    %>

                                <tr>
                                    <td>Post:</td>
                                    <td class="text-right"><%=formatnumber(fctimer, 2) %></td>
                                    <td class="text-right"><%=formatnumber(medarbSumtimerUtilpost, 2) %></td>
                                    <td><div style="background-color:<%=indicatorCol%>; width:10px; height:10px;">&nbsp;</div></td>
                                </tr>
                               
                            </tbody>
                        </table>
                    </div>
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-5">
                                <div id="line-chart" class="chart-holder"></div>
                            </div>
                    </div>
                    </div>

                </div>
            </div>

  </div>
</div>







<!--#include file="../inc/regular/footer_inc.asp"-->