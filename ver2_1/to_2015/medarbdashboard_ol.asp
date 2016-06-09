

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
<script type="text/javascript" src="js/demos/flot/donut.js"></script>
<script type="text/javascript" src="js/demos/flot/vertical.js"></script>
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
    %>

<div class="wrapper">
    <div class="content">
    
    <div class="container">
 
         <div class="portlet">
        <h3 class="portlet-title"><u>Medarbejder Dashboard</u></h3>

        <div class="well">
              <form action="medarbdashboard.asp?sogsubmitted=1" method="POST">
            <h4 class="panel-title-well">Periode</h4>
            <div class="row">
                <div class="col-lg-3">Medarbejder:</div>
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2"><input type="radio" name="FM_periode" value="1" onchange="submit()" <%=usePeriodeSel1 %> /> Måned:</div>
                <div class="col-lg-2"><input type="radio" name="FM_periode" value="2" onchange="submit()" <%=usePeriodeSel2 %> /> Kvartal:</div>
                <div class="col-lg-2"><input type="radio" name="FM_periode" value="3" onchange="submit()" <%=usePeriodeSel3 %> /> År:</div>
            </div>


            <div class="row">

                 <div class="col-lg-3">
                     <%=meTxt %> <!--På den der er logget på-->
                 </div>
                <div class="col-lg-1">&nbsp</div>
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
                <select name="FM_kv" class="form-control input-small" onchange="submit()">
                  <option value="1" <%=useQuaterSel1 %>>Kvt. 1</option>
                  <option value="2" <%=useQuaterSel2 %>>Kvt. 2</option>
                  <option value="3" <%=useQuaterSel3 %>>Kvt. 3</option>
                  <option value="4" <%=useQuaterSel4 %>>Kvt. 4</option>
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
           


        '** Normtimer
        natalDage = dateDiff("d", strDatoStartKri, strDatoEndKri, 2,2)
        call normtimerPer(usemrn, strDatoStartKri, natalDage, 0)
        ntimPer = ntimPer


        '** Real. timer SEL MEDARB
         strSQlmtimer = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris) AS sumtimeOms, SUM(timer*kostpris) AS sumtimeKost FROM timer WHERE tmnr = "& usemrn & " AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'"
            
         hoursDB = 0
         hoursDBindicator = 0
         hoursAvgRate = 0   
         hoursAvgindicator = 0
         utilz = 0

        'response.write strSQlmtimer
        'response.flush

         oRec.open strSQlmtimer, oConn, 3
         if not oRec.EOF then


            hoursDB = oRec("sumtimeOms") - oRec("sumtimeKost")
            hoursDBindicator = (oRec("sumtimeKost") / oRec("sumtimeOms")) * 100 'dbproc

            hoursAvgRate = oRec("sumtimeOms") / oRec("sumtimer")

            if hoursAvgRate > 600 then
            hoursAvgindicato = 100 - ((600 / hoursAvgRate * 100))
            else
            hoursAvgindicato = (hoursAvgRate / 600 * 100)
            end if

            utilz = (oRec("sumtimer") - ntimPer)

            if ntimPer > oRec("sumtimer") then
            utilzDB = (oRec("sumtimer") / ntimPer) * 100
            utilzCol = "danger"
            utilzDBFortegn = "-"
            else
            utilzDB = (ntimPer / oRec("sumtimer")) * 100
            utilzCol = "success"
            utilzDBFortegn = "+"
            end if

         end if
         oRec.close 



        '** Real. timer MEDARB PROJEKTGRUPPE ***'
        call hentbgrppamedarb(usemrn)
        
        '** FJERN ALLE GRUPPEN
        projektgruppeIds = replace(projektgruppeIds, ", 10", "") 

        call medarbiprojgrp(projektgruppeIds, usemrn, 0, -1)


        medarbgrpIdSQLkri = replace(medarbgrpIdSQLkri, "mid", "tmnr")
        
        strSQlmtimerProjgrp = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris) AS sumtimeOms, SUM(timer*kostpris) AS sumtimeKost, "_
        &" SUM(timer*timepris-timer*kostpris) AS pghoursDB FROM timer WHERE "_
        &" (tmnr = 0 "& medarbgrpIdSQLkri &") AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' GROUP BY tmnr"
            
         pghoursDB = 0
         pghoursDBindicator = 0
         pghoursAvgRate = 0   
         pghoursAvgindicator = 0
         pgutilz = 0
         m = 1
        'response.write strSQlmtimerProjgrp
        'response.flush

         oRec.open strSQlmtimerProjgrp, oConn, 3
         while not oRec.EOF


            pghoursDB = pghoursDB + oRec("pghoursDB")
            pghoursDBindicator = pghoursDBindicator + (oRec("sumtimeKost") / oRec("sumtimeOms")) * 100 'dbproc

            'pghoursAvgRate = oRec("sumtimeOms") / oRec("sumtimer")

            'if pghoursAvgRate > 600 then
            'pghoursAvgindicato = 100 - ((600 / hoursAvgRate * 100))
            'else
            'pghoursAvgindicato = (hoursAvgRate / 600 * 100)
            'end if

            'pgutilz = (oRec("sumtimer") - ntimPer)

            'if pgntimPer > oRec("sumtimer") then
            'pgutilzDB = (oRec("sumtimer") / ntimPer) * 100
            'pgutilzCol = "danger"
            'upgtilzDBFortegn = "-"
            'else
            'pgutilzDB = (pgntimPer / oRec("sumtimer")) * 100
            'pgutilzCol = "success"
            'pgutilzDBFortegn = "+"
            'end if

            m = m + 1

         oRec.movenext
         wend 
         oRec.close 


            pghoursDB = pghoursDB / m-1
            pghoursDBTxt = replace(formatnumber(pghoursDB, 0), ".", "")
            
            pghoursDBindicator = pghoursDBindicator / m-1

            
            if hoursDBindicator < 20 then
            hoursDBindicatorColor = "danger"
            else
            hoursDBindicatorColor = "success"
            end if

            if hoursAvgindicato < 0 then
            hoursAvgindicatoColor = "danger"
            hoursAvgindicatoFortegn = "-"
            else
            hoursAvgindicatoColor = "success"
            hoursAvgindicatoFortegn = "+"
            end if
            %>


        <div class="portlet-body">

            <div class="row">
                <div class="col-lg-11">
                    <h4 class="portlet-title"><u>Nøgletal</u>
                    </h4>
                </div>
            </div>

            <%if hoursDB <> 0 then
                hoursDB = hoursDB
              else
                hoursDB = 0
               end if %>
           
            
            <div class="row">
            
            <div class="col-sm-6 col-md-3">
              <div class="row-stat">
                <p class="row-stat-label">Dækningsbidrag1</p>
                <h3 class="row-stat-value"><%=formatnumber(hoursDB,2) %></h3>
                <span class="label label-<%=hoursDBindicatorColor %> row-stat-badge"><%=formatnumber(hoursDBindicator, 0)%> %</span> <p><br /></p>

                <p class="row-stat-label">Dækningsbidrag2</p>
                <h3 class="row-stat-value"><%=formatnumber(hoursDB,2) %></h3>
                <span class="label label-<%=hoursDBindicatorColor %> row-stat-badge"><%=formatnumber(hoursDBindicator, 0)%> %</span>
              </div> <!-- /.row-stat -->
            </div> <!-- /.col -->

            <div class="col-lg-1">&nbsp</div>
            <div class="col-sm-6 col-md-3">
              <div class="row-stat">
                <p class="row-stat-label">Faktureringsgrad</p>
                <h3 class="row-stat-value"><%=formatnumber(hoursAvgRate,2) %></h3>
                <span class="label label-<%=hoursAvgindicatoColor %> row-stat-badge"><%=hoursAvgindicatoFortegn &" "& formatnumber(hoursAvgindicato, 0) %> %</span>
              </div> <!-- /.row-stat -->
            </div> <!-- /.col -->

            <div class="col-lg-1">&nbsp</div>
            <div class="col-sm-6 col-md-3">
              <div class="row-stat">
                <p class="row-stat-label">Utilization</p>
                <h3 class="row-stat-value"><%=formatnumber(utilz,2) %> t.</h3>
                <span class="label label-<%=utilzCol %> row-stat-badge"><%=utilzDBFortegn &" "& formatnumber(utilzDB, 0) %> %</span>
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
                        DB1 
                    </div>
                    <div class="progress-stat-value">
                        <%=formatnumber(hoursDB,2) %> kr.
                    </div>
                    <div class="progress progress-striped progress-sm active">
                        <div class="progress-bar progress-bar-secondary" role="progressbar" aria-valuenow="<%=formatnumber(hoursDBindicator, 0)%>" aria-valuemin="0" aria-valuemax="100" style="width:100%">
                        <span class="sr-only"><%=formatnumber(utilzDB,0)%>%</span>
                        </div>
                    </div> 
                </div>
                    
                    <div class="progress-stat">
                    <div class="progress-stat-label">
                        <%=meTxt %>
                    </div>
                    <div class="progress-stat-value">
                        100%
                    </div>
                    <div class="progress progress-striped progress-sm active">
                        <div class="progress-bar progress-bar-default" role="progressbar" aria-valuenow="<%=formatnumber(pghoursDBindicator, 0)%>" aria-valuemin="0" aria-valuemax="100" style="width: 100%">
                        <span class="sr-only"><%=formatnumber(pghoursDBindicator,0) %>%</span>
                        </div>
                    </div> 
                </div> 

                    <div class="progress-stat">
                    <div class="progress-stat-label">
                        DB2 
                    </div>
                    <div class="progress-stat-value">
                       5.000 kr.
                    </div>
                    <div class="progress progress-striped progress-sm active">
                        <div class="progress-bar progress-bar-secondary" role="progressbar" aria-valuenow="<%=formatnumber(hoursDBindicator, 0)%>" aria-valuemin="0" aria-valuemax="100" style="width:100%">
                        <span class="sr-only"><%=formatnumber(utilzDB,0)%>%</span>
                        </div>
                    </div> 
                </div>
                    
                    <div class="progress-stat">
                    <div class="progress-stat-label">
                        <%=meTxt %>
                    </div>
                    <div class="progress-stat-value">
                        50%
                    </div>
                    <div class="progress progress-striped progress-sm active">
                        <div class="progress-bar progress-bar-default" role="progressbar" aria-valuenow="<%=formatnumber(pghoursDBindicator, 0)%>" aria-valuemin="0" aria-valuemax="100" style="width: 50%">
                        <span class="sr-only"><%=formatnumber(pghoursDBindicator,0) %>%</span>
                        </div>
                    </div> 
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
                                 '** Real. timer SEL MEDARB
                                 strSQlmtimer_aktiviater = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris) AS sumtimeOms, SUM(timer*kostpris) AS sumtimeKost, taktivitetnavn FROM timer "_
                                 &" WHERE tmnr = "& usemrn & " AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' GROUP BY tmnr, taktivitetnavn ORDER BY sumtimer DESC"
            
                                

                                 oRec.open strSQlmtimer_aktiviater, oConn, 3
                                 while not oRec.EOF 


                                    fctimer = 2

                                    if cdbl(fctimer) > cdbl(oRec("sumtimer")) then
                                    indicatorCol = "yellowgreen"
                                    else
                                    indicatorCol = "crimson"
                                    end if
                                    %>

                                <tr>
                                    <td><%=oRec("taktivitetnavn") %></td>
                                    <td class="text-right"><%=formatnumber(fctimer, 2) %></td>
                                    <td class="text-right"><%=formatnumber(oRec("sumtimer"), 2) %></td>
                                    <td><div style="background-color:<%=indicatorCol%>; width:10px; height:10px;">&nbsp;</div></td>
                                </tr>

                                <%
                                 oRec.movenext
                                 wend
                                 oRec.close %>


                               
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