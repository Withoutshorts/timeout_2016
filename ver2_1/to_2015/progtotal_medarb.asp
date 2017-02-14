

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
                                Strstartdate = "2015/01/01"
                                Strslutdate = "2017/01/01"
                                
                                fromDate="2015-01-01"
                                toDate="2016-01-01"

                                antalmaaned = (DateDiff("m",fromDate,toDate)) 
                                response.Write antalmaaned
                             
                                Strheader = "<th>&nbsphej</th>"
                                                                                               
                            %>

                            <div class="col-lg-1">
                                <select name="FM_start_mrd" class="form-control input-small" onchange="submit();">
                                    <option value="<%=strDag%>"><%=strDag%></option>
		                            <option value="1">jan</option>
	   	                            <option value="2">feb</option>
	   	                            <option value="3">mar</option>
	   	                            <option value="4">apr</option>
	   	                            <option value="5">maj</option>
	   	                            <option value="6">jun</option>
	   	                            <option value="7">jul</option>
	   	                            <option value="8">aug</option>
	   	                            <option value="9">sep</option>
	   	                            <option value="10">okt</option>
	   	                            <option value="11">nov</option>
	   	                            <option value="12">dec</option>
                                </select>
                            </div>
                            <div class="col-lg-1">
                                <select name="FM_start_aar" class="form-control input-small" onchange="submit();">
		                            <option value="<%=strAar%>"><%=strAar%></option>
		                            <%for x = -10 to 10 
		                            useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		                            <option value="<%=useY%>"><%=right(useY, 2)%></option>
		                            <%next %>
		                        </select>
                            </div>

                           
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

                      <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                          
                          <thead>
                              <tr>
                                  <th>Projekt: <%=jobid %> <br />
                                      Periode:
                                  </th>

                                  <%
                                      antalmaaned = 12

                                      for i = 0 to antalmaaned  

                                        %> <th style="text-align:center;">header</th> <%

                                      next
                                  %>

                                  <th>Total pr. medarbejder</th>

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
                                                                  
                                  <%
                                      for i = 0 to antalmaaned  

                                        %> <td style="text-align:center;">&nbsp</td> <%

                                      next 
                                  %>

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