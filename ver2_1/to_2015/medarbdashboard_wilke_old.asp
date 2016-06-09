
 <div class="well">
              <form action="medarbdashboard.asp?sogsubmitted=1" method="POST">
            <h4 class="panel-title-well">Søgefilter</h4>
            <div class="row">
                <div class="col-lg-3">Medarbejder</div>
               <!-- <div class="col-lg-2"><input type="radio" name="FM_periode" value="0" onchange="submit()" <%=usePeriodeSel0 %> /> Uge:</div>-->
                <div class="col-lg-2"><input type="radio" name="FM_periode" value="1" onchange="submit()" <%=usePeriodeSel1 %> /> Måned:</div>
                <div class="col-lg-2"><input type="radio" name="FM_periode" value="2" onchange="submit()" <%=usePeriodeSel2 %> /> Kvartal:</div>
                <div class="col-lg-2"><input type="radio" name="FM_periode" value="3" onchange="submit()" <%=usePeriodeSel3 %> /> År:</div>
            </div>


            <div class="row">

                 <div class="col-lg-3">
                <select name="FM_medarb" class="form-control input-small" onchange="submit()">
                    <%strSQLm = "SELECT mid, mnavn, init FROM medarbejdere WHERE mansat = 1 AND (medarbejdertype = 3 OR medarbejdertype = 10 OR medarbejdertype = 28) OR mid  < 60 ORDER BY mnavn"
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

                <!--
                 <div class="col-lg-2">
                <select class="form-control input-small" onchange="submit()">
                  <option value="41">41</option>
                  <option value="42">42</option>
                  <option value="43">43</option>
                  <option value="44">44</option>
                </select>
                 </div>
                -->

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



<!-- SLET -->


              <%
               f = 1   
               if hoursDB(f) <> 0 then
                hoursDB(f) = hoursDB(f)
              else
                hoursDB(f) = 0
               end if %>
           
            
            <div class="row">
            
             <div class="col-sm-6 col-md-3">
              <div class="row-stat">
                <p class="row-stat-label">DB1</p>
                <h3 class="row-stat-value"><%=formatnumber(hoursOms(f), 2)%></h3>
                <!--<span class="label label-danger row-stat-badge"><%=formatnumber(jo_dbproc, 0)%> %</span> -->
                     <br /> <span style="font-size:11px; color:#999999;">[ timer * timepris ]</span>
                  <br /><br />
                <p class="row-stat-label">DB2</p>
                <h3 class="row-stat-value"><%=formatnumber(hoursDB(f), 2)%></h3>
                <!--<span class="label label-succes row-stat-badge"><%=formatnumber(jo_db2_proc, 0)%> %</span> -->
                  <br /> <span style="font-size:11px; color:#999999;">[ timer * (timepris-kostpris) ]</span>

                  <br /><br />
                   <p class="row-stat-label">Gennemsnitlig timepris</p>
                <h3 class="row-stat-value"><%=formatnumber(hoursAvgRate(f),2) %></h3>
                <span class="label label-<%=hoursAvgindicatoColor %> row-stat-badge"><%=hoursAvgindicatoFortegn &" "& formatnumber(hoursAvgindicato(f), 0) %> %</span>


                 <!-- <br /> <span style="font-size:9px; color:#999999;">Budget: <%=formatnumber(jo_db2,2) %></span>-->
              </div> <!-- /.row-stat -->
            </div> <!-- /.col -->

              

             <div class="col-lg-1">&nbsp</div>
            <div class="col-sm-6 col-md-3">
              <div class="row-stat">
                <p class="row-stat-label">NPS</p>
                <h3 class="row-stat-value">8,2</h3>
                
              </div> <!-- /.row-stat -->
            </div> <!-- /.col -->

            <div class="col-lg-1">&nbsp</div>
            <div class="col-sm-6 col-md-3">
              <div class="row-stat">
                <p class="row-stat-label">Tidsfordeling Target</p>
                  <%
                   for f = 0 TO fo - 1 
                  %>
                  
                      <%=fomrnavn(f)%>: <b><%=fomrmal(f) %> %  </b> <span class="label label-<%=utilzCol(f) %> row-stat-badge"><%=utilzDBFortegn(f) &" "& formatnumber(utilzDB(f)*100, 0) %> %</span><br />

                  <%next%>

                  <p class="row-stat-label">Tidsfordeling Timer</p>

                  

                  <%
                  response.Flush    
                      
                 
                  f = 1%>

              </div> <!-- /.row-stat -->
            </div> <!-- /.col -->

            </div> <!--Row-->

            <div class="row pad-t20">
                <div class="col-lg-3">
                    <h4 class="portlet-title"><!--<u>DB</u>--></h4>
                </div>
                
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-3">
                    <h4 class="portlet-title"><u>Tidsfordeling - <span style="font-size:11px; color:#999999;"><%=meTxt %> / Peers</span></u></h4>
                </div>
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-3">
                    <h4 class="portlet-title"><u>Tidsfordeling %</u></h4>
                </div>
            </div>

            </div>

            
            <div class="row">


                

                  <div class="col-lg-1">&nbsp</div>            
                    <div class="col-lg-3"><!--<div id="vertical-chart" class="chart-holder-250"></div>--></div>
                
                <div class="col-lg-3">  
                <div class="progress-stat">
                    <div class="progress-stat-label">
                        <%=meTxt %> 
                    </div>
                    

                     <%
                        
                   for f = 0 TO fo - 1 
                   maksHoursBarThis = formatnumber(maksHoursBar*(fomrmal(f)/100), 0) 
                   if maksHoursBarThis <> 0 then
                   realThisProc = formatnumber((hours(f) / maksHoursBarThis) * 100, 0)
                    else
                    realThisProc = 0
                    end if
                      %>
                       <div class="progress-stat-value" style="font-size:10px;">
                         <%=fomrnavn(f)%>: <br /><%=formatnumber(hours(f),2) %> : <%=realThisProc %>%
                    </div>
                    <div class="progress progress-striped progress-sm active">
                        <div class="progress-bar progress-bar-secondary" role="progressbar" aria-valuenow="50" aria-valuemin="0" aria-valuemax="100" style="width:<%=realThisProc%>%">
                        <span class="sr-only">50%</span>
                        </div>
                    </div> 

                  <%next%>


                     </div>


                     <div class="progress-stat">
                    <div class="progress-stat-label">
                        Peers (proj.grp)
                    </div>

                       <%
                   for f = 0 TO fo - 1 
                   maksHoursBarThis = formatnumber(maksHoursBar*(fomrmal(f)/100), 0) 
                   realThisProc = formatnumber((hours(f)*2 / maksHoursBarThis) * 100, 0)
                      %>
                       <div class="progress-stat-value" style="font-size:10px;">
                         <%=fomrnavn(f)%>: <br /><%=formatnumber(hours(f),2) %> : <%=realThisProc %>%
                    </div>
                    <div class="progress progress-striped progress-sm active">
                        <div class="progress-bar progress-bar-primary" role="progressbar" aria-valuenow="50" aria-valuemin="0" aria-valuemax="100" style="width:<%=realThisProc%>%">
                        <span class="sr-only">50%</span>
                        </div>
                    </div> 

                  <%next
                      
                      f = 1%>

                           </div>
                   



              

               
            </div>   

                

                  

                    <div class="col-lg-1">&nbsp</div>
                    <div class="col-lg-3">            
                    <div id="xdonut-chart" class="chart-holder" style="width: 95%;"></div>
                    </div>


            </div>
        </div> 
    </div>


            <div class="container">
                <div class="portlet">
               
                    <div class="portlet-body">
                        
                        <div class="row">
                        <div class="col-lg-5">
                             <h4 class="portlet-title"><u>Projekter</u></h4>
                        </div>
                            <!--
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-5">
                            <h4 class="portlet-title"><u>Interne og Eksterne timer</u></h4>
                            </div>
                            -->
                        </div>

                        <div class="row">
                        <div class="col-lg-5">
                       <%=meTxt %> 
                        <table class="table table-bordered">
                            <thead>
                                <tr>
                                    <th style="width: 50%">Projekt</th>
                                    <th style="width: 20%">Forecast</th>

                                    <th style="width: 20%">Realiseret</th>
                                    <th style="width: 10%"><!--Indikator--></th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                 '** Real. timer SEL MEDARB
                                 strSQlmtimer_aktiviater = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris) AS sumtimeOms, SUM(timer*kostpris) AS sumtimeKost, tjobnavn, taktivitetnavn FROM timer "_
                                 &" WHERE tmnr = "& usemrn & " AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' GROUP BY tmnr, tjobnr ORDER BY sumtimer DESC"
            
                                

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
                                    <td><%=oRec("tjobnavn") %></td>
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

                            <!--
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-5">
                                <div id="line-chart" class="chart-holder"></div>
                            </div>
                            -->
                    </div>
                    </div>

                </div>
            </div>

        <!-- SÆET HERTIL -->