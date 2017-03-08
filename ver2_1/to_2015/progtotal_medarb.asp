 

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<%
    
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if


    if len(trim(request("FM_medarb"))) <> 0 then
    usemrn = request("FM_medarb")
    else
    usemrn = session("mid") 
    end if

    if len(trim(request("aar"))) <> 0 then
    aar = request("aar")
    else
    aar = "1-1-"& year(now)
    end if

    if len(trim(request("aarslut"))) <> 0 then
    aarslut = request("aarslut")
    else
    aarslut = "1-1-"& year(now)
    end if


    if len(trim(request("FM_job"))) <> 0 then
    jost0CHK = ""
    else
    viskunabnejob0 = 1
    jost0CHK = "CHECKED"
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



<%call menu_2014 %>

  <div class="wrapper">
      <div class="content">
          <script src="js/traveldietexp_jav.js" type="text/javascript"></script>
          <div class="container">
              <div class="portlet">
                  <h3 class="portlet-title"><u>Projekt Rapport</u></h3>
                  <div class="portlet-body">

                      <div class="well">
                        <form action="progtotal_medarb.asp?sogsubmitted=1" method="POST">

                        <div class="row">
                            <div class="col-lg-2">
                                <h4 class="panel-title-well"><%=dsb_txt_002 %></h4>
                            </div>
                            <div class="col-lg-10">&nbsp</div>                 
                        </div>

                
                        <div class="row">
                            <div class="col-lg-2"><%=dsb_txt_003 %>:</div>

                            <div class="col-lg-4">
                                
                                <%
                                    strSQLjob = "SELECT id, jobnr, jobnavn FROM job order by jobnavn"                                                 
                                %>
                                <select name="FM_job" id="FM_job" <%=progrpmedarbDisabled  %> class="form-control input-small"  onchange="submit();">
                                  <%

                                    oRec3.open strSQLjob, oConn, 3
                                    while not oRec3.EOF
                                
                                    Strjobnavn = oRec3("jobnavn")

                                    if cdbl(jobid) = cdbl(oRec3("id")) then
				                    isSelected = "SELECTED"
				                    else
				                    isSelected = ""
				                    end if
				
				                    jobnavnid = replace(oRec3("jobnavn"), "'", "") & " ("& oRec3("jobnr") &")"
        
				                %>
				                <option value="<%=oRec3("id")%>" <%=isSelected%>><%=jobnavnid &" "& jstTxt%></option>
				                <%

                                oRec3.movenext
                                wend
                                oRec3.close 

                                %>

                                </select>
               
                            </div>

                            <%
                                
                                'response.Write aar & " - " & aarslut 

                                antalmaaned = (DateDiff("m",aar,aarslut))
                                antalaar = (DateDiff("yyyy",aar,aarslut))
                                'response.Write(DateDiff("m",aar,aarslut))
                                'response.Write antalaar
                                                                                                                   
                            %>

                            <div class="col-lg-1">Fra:</div>
                            <div class="col-lg-2">
                                <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="aar" value="<%=aar %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>

                            <div class="col-lg-1">Til:</div>
                            <div class="col-lg-2">
                                <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="aarslut" value="<%=aarslut %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>  
                        </div>
                       
                        <div class="row">
                            <br />
                            <div class="col-lg-10">&nbsp</div>
                            <div class="col-lg-2"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Søg >></b></button></div>
                        </div>                       

                        </form>
                      </div>

                      <%
                           strSQLjobid = "SELECT jobnr FROM job WHERE id ="& jobid
                           oRec4.open strSQLjobid, oConn, 3

                           if not oRec4.EOF then
                          
                            Strjobid = oRec4("jobnr")

                           end if
                          
                          oRec4.close 
                          'response.Write "test" & Strjobid
                      %>

                     <!-- <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                          
                          <thead>
                              <tr>
                                  <th>Projekt: <%=jobid %> <br />
                                      Periode:
                                  </th>
                                  <th style="text-align:center">Jan</th>
                                  <th style="text-align:center">Feb</th>
                                  <th style="text-align:center">Mar</th>
                                  <th style="text-align:center">Apr</th>
                                  <th style="text-align:center">Maj</th>
                                  <th style="text-align:center">Jun</th>
                                  <th style="text-align:center">Jul</th>
                                  <th style="text-align:center">Aug</th>
                                  <th style="text-align:center">Sep</th>
                                  <th style="text-align:center">Okt</th>
                                  <th style="text-align:center">Nov</th>
                                  <th style="text-align:center">Dec</th>
                                  <th style="text-align:center">Total pr. medarbejder</th>
          
                                      
                              </tr>
                          </thead>



                          <tbody>
                            

                              <%
                                  
                                  strSQLtimer = "SELECT Tmnavn, Tmnr, sum(Timer) as Timer FROM timer WHERE tjobnr = "& Strjobid & " GROUP BY Tmnr "
                                  
                                  oRec5.open strSQLtimer, oConn, 3
      
                                  while not oRec5.EOF

                                  Strmedarb = oRec5("Tmnavn")
                                  Strtimer = oRec5("Timer")
                                   
                              %>

                              <tr>
                                  <td><%=Strmedarb %></td>                                  
                                  <td>&nbsp</td>
                                  <td>&nbsp</td>
                                  <td>&nbsp</td>
                                  <td>&nbsp</td>
                                  <td>&nbsp</td>
                                  <td>&nbsp</td>
                                  <td>&nbsp</td>
                                  <td>&nbsp</td>
                                  <td>&nbsp</td>
                                  <td>&nbsp</td>
                                  <td>&nbsp</td>
                                  <td>&nbsp</td>
                                  <td style="text-align:center"><%=Strtimer %></td>
                              </tr>

                              <%
                                oRec5.movenext
                                wend
                                oRec5.close 
                              %>

                          </tbody>

                          <tfoot>

                              <tr>
                                  <th>Total pr. månded</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0</th>
                                  <th style="text-align:center">0000</th>
                              </tr>

                          </tfoot>

                      </table> -->
                      <%

                        
                      %>
                      <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                          
                          <thead>
                              <tr>
                                  <th>Projekt: <%=jobid %> <br />
                                      Periode:
                                  </th>

                                  <%

                                      


                                      selectedstart_month = Month(aar)
                                      selectedstart_year = Year(aar)

                                      
                                      'response.write selectedstart_month

                                      startdate = "01/" & selectedstart_month & "/" & selectedstart_year

                                      

                                      months = selectedstart_month
                                      Stryear = selectedstart_year 
                                      
                                      'response.write startdate 

                                      for i = 0 to antalmaaned

                                            months = months + 1
                                            
                                            if months > 13 then
                                                months = 2
                                                Stryear = Stryear + 1
                                            end if
                                            
                                      
                                            'getdate = months & "/" & Stryear

                                               
                                           
                                           'response.write(monthname(i))
                                            'response.write MonthName(months)
                                                                          
                                        %> <th style="text-align:center;"><%=MonthName(months - 1) & " " & Stryear %></th> <%

                                      next
                                  %>

                                  <th>Total pr. medarbejder</th>
                              </tr>

                          </thead>

                          <tbody>
                            
                              <%
                                  startmonth = Month(aar)
                                  startyear = Year(aar)
                                  slutmonth = Month(aarslut)
                                  slutyear = Year(aarslut)
                                  
                                  startdato = startyear & "-" & startmonth & "-1"
                                  slutdato = slutyear & "-" & slutmonth & "-1"

                                  response.Write "start: " & startdato & "slut: " & slutdato 

                                    
                                  lastmid = 0
                                  x = 200
                                  dim timer_md, dato_medid, medarbid, medidnavn
                                  Redim timer_md(x), dato_medid(x), medarbid(x), medidnavn(x)
                                  
                                  strSQL = "SELECT tmnr, tmnavn, tdato, sum(timer) as Timer FROM timer WHERE tjobnr = "& Strjobid & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' AND tmnr is not null GROUP by tmnr, year(tdato), month(tdato) Order by tmnr, tdato "
                                  'response.Write strSQL
                                  'repsonse.flush
                                  oRec.open strSQL, oConn, 3
                                  d = 0
                                  While not oRec.EOF

                                  medarbid(oRec("tmnr")) = oRec("tmnr")
                                  timer_md(d) = oRec("timer")
                                  dato_medid(d) = year(oRec("tdato")) & month(oRec("tdato"))
                                  medidnavn(oRec("tmnr")) = oRec("tmnavn")
                                  d = d + 1
                                  oRec.movenext
                                  Wend 
                                  oRec.close
                                  d = 0
                                  for m = 0 TO UBound(medarbid)
                                  if medarbid(m) <> lastmid and len(trim(medidnavn(m))) <> 0 then
                                  %>
                                        <tr>
                                            <td><%=medidnavn(m) & " " & medarbid(m) %></td>

                                            <%  months = selectedstart_month 
                                                for d = 0 TO antalmaaned
                                                        months = months + 1
                                                        if months = 0 then
                                                        tjekdatoym = startdato
                                                        else
                                                        tjekdatoym = dateadd("m",1,tjekdatoym)
                                                        end if

                                                        tjekdatoym = startyear & startmonth
                                                        if tjekdatoym = dato_medid(months) then
                                                         %>
                                                        <td><%=timer_md(months) %> <%response.write months %></td>
                                                    <% 
                                                        else
                                                        
                                                    %>
                                                         
                                                        <td><%response.write d %></td>
                                                    <%
                                                        end if
                                                   

                                                next  
                                            %>

                                            <td>total</td>
                                        </tr>
                                  <%lastmid = medarbid(m) 
                                        end if

                                        
                                  next
       
                              %>

                          </tbody>

                          <tfoot>

                              <tr>
                                  <th>Total pr. månded</th>
                                  
                                  <%
                                      for i = 0 to antalmaaned  

                                        %> <th style="text-align:center;">0</th> <%

                                      next 
                                  %>

                                  <th style="text-align:center">0000</th>
                              </tr>

                          </tfoot>

                      </table>

                  </div>
              </div>
          </div>


      </div>
  </div>  

<!--#include file="../inc/regular/footer_inc.asp"-->