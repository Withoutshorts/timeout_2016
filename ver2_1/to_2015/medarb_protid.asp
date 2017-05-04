 

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
                  <h3 class="portlet-title"><u>Medarbejder - projekttid</u></h3>
                  <div class="portlet-body">

                      <div class="well">
                        <form action="medarb_protid.asp?sogsubmitted=1" method="POST">

                        <div class="row">
                            <div class="col-lg-2">
                                <h4 class="panel-title-well"><%=dsb_txt_002 %></h4>
                            </div>
                            <div class="col-lg-10">&nbsp</div>                 
                        </div>

                
                        <div class="row">
                          

                            <div class="col-lg-6">Projekt:<br />
                                
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

                            
                                
                                
                           
                            <div class="col-lg-2">Fra: (hele måneder)<br />
                                <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="aar" value="<%=aar %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>

                            
                            <div class="col-lg-2">Til:<br />
                                <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="aarslut" value="<%=aarslut %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>
                            <div class="col-lg-1"><br /><button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Søg >></b></button></div>  
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
                      <table style="background-color:white;" class="table dataTable table-striped table-bordered table-hover ui-datatable">

                          <thead>
                              <tr>
                                    <th>Medarbejder</th>

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
                                           
                                      
                                            getdate = Monthname(months -1 ,True) & "<br>" & Stryear
                                            'getdate = (Ucase(getdate))   
                                      
                                                                          
                                        %> <th style="text-align:center;"><%=UCase(Left(getdate,1)) & Mid(getdate,2) %></th> 
                                  <%

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
                                                                 
                                  startdato = startyear & "/" & (startmonth) & "/1"
                                  'slutdato = slutyear & "/" & (slutmonth + 1) & "/1"

                                  slutdato = dateAdd("m", 1, "1/"& slutmonth &"/"& slutyear) 
                                  slutdato = dateAdd("d", -1, slutdato)
                                  slutdato = year(slutdato) & "/" & month(slutdato) & "/" & day(slutdato)


                                  'if slutmonth = 12 then
                                  'slutdato = (slutyear + 1) & "/1/1"
                                  'end if
                                 
                                  'response.Write "start: " & startdato & "slut: " & slutdato 

                                    
                                  lastmid = 0
                                  m = 1000
                                  d = 10000
                                  lasttmnr = 0
                                  dim timer_md, dato_medid, medarbid, medidnavn, manedstot
                                  Redim timer_md(m,d), dato_medid(m,d), medarbid(m), medidnavn(m), manedstot(antalmaaned)
                                  
                                  strSQL = "SELECT tmnr, tmnavn, tdato, sum(timer) as Timer FROM timer WHERE tfaktim <> 5 AND tjobnr = '"& Strjobid & "' AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' AND tmnr is not null GROUP by tmnr, year(tdato), month(tdato) Order by tmnr, tdato "
                                  'response.Write strSQL
                                  'response.flush
                                  oRec.open strSQL, oConn, 3
                                  d = 0
                                  m = -1
                                  While not oRec.EOF

                                  if lasttmnr <> oRec("tmnr") then
                                     
                                     m = m + 1                                     
                                    
                                  end if
                                  
                                  'd = year(oRec("tdato")) & month(oRec("tdato"))
                                  medarbid(m) = oRec("tmnr")
                                  timer_md(m,d) = oRec("timer")
                                  dato_medid(m,d) = year(oRec("tdato")) & month(oRec("tdato"))
                                  medidnavn(m) = oRec("tmnavn")
                                  'response.Write "<br>m" & m & "d" & d & "dato" & dato_medid(m,d)
                                  'lasttmnr = oRec("tmnr")
                                  

                                  lasttmnr = oRec("tmnr")
                                  'm = m + 1
                                  d = d + 1
                                  oRec.movenext
                                  Wend 
                                  oRec.close

                                  m_end = m
                                  m = 0
                                  timerperiode = 0

                                  'response.Write "m" & m & "d" & d  
                                  for m = 0 TO m_end
                                  'if medarbid(m) <> lastmid and len(trim(medidnavn(m))) <> 0 then

                                  call meStamdata(medarbid(m)) 

                                  %>
                                        <tr>
                                            <td style="white-space:nowrap;"><%=left(medidnavn(m), 10) & " ["& meInit &"]" %></td>

                                            <%
                                                'expTxt = expTxt & medidnavn(m) & ";" 

                                               for i = 0 to antalmaaned  

                                                       
                                                               
                                                        if i = 0 then
                                                        tjekdato = startdato
                                                        else
                                                        tjekdato = dateadd("m",1,tjekdato)                                                       
                                                        end if
                                                        foundone = 0

                                                        
                                                        for d = 0 to Ubound(timer_md)

                                                            tjekdatoym = year(tjekdato) & month(tjekdato)

                                                            if tjekdatoym = dato_medid(m,d) then
                                                             
                                                            tdval = formatnumber(timer_md(m,d), 2) '& "<br>m" & m & "<br>d" & d & "<br>" & dato_medid(m,d) & ""
                                                            foundone = 1
                                                            manedstot(i) = manedstot(i) + timer_md(m,d) 
                                                            'expTxt = expTxt & timer_md(m,d) & ";"  
                                                            end if
                                                            if cint(foundone) = 0 then 
                                                            tdval = "&nbsp" '& m & "<br>" & d & "<b>" & tjekdatoym
                                                            'expTxt = expTxt & ";"  
                                                            end if
                                                                                                                       
                                                        next 
                                                    
                                                        %>
                                                        <td style="text-align:right;"><%=tdval %> </td>
                                                        <%
                                                        
                                                next
                                                
                                            %>

                                            <td style="text-align:right">
                                                <%

                                                    strSQLtimer = "SELECT sum(timer) as Timer FROM timer WHERE tfaktim <> 5 AND tjobnr = '"& Strjobid & "' AND tmnr ="& medarbid(m) & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'"
                                                    'response.Write strSQLtimer
                                                    'response.Flush
                                                    oRec.open strSQLtimer, oConn, 3
                                                    if not oRec.EOF then

                                                    timerperiode = oRec("Timer")

                                                    end if
                                                    oRec.close
                                                    
                                                %>
                                                <%=formatnumber(timerperiode, 2) %>
                                            </td>
                                        </tr>
                                  <%
                                        'end if

                                     lastmid = medarbid(m)    
                                  next
       
                              %>
                              
                          </tbody>

                          <tfoot>

                              <tr>
                                  <th>Total</th>
                                  
                                  <%

                                      for i = 0 to antalmaaned  

                                        %> <th style="text-align:right;"><%=formatnumber(manedstot(i), 2) %></th> <%

                                      next 
                                  %>

                                  <th style="text-align:right">
                                      <%
                                          strSQLtimertotal = "SELECT sum(timer) as Timer FROM timer WHERE tfaktim <> 5 AND tjobnr = '"& Strjobid & "' AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY tjobnr"

                                          'response.Write strSQLtimertotal

                                          oRec.open strSQLtimertotal, oConn, 3
                                          if not oRec.EOF then

                                            timertotal = oRec("Timer")
                                            if timertotal <> 0 then
                                            else
                                            timertotal = 0
                                            end if

                                          end if
                                          oRec.close
                                      %>

                                      <%=formatnumber(timertotal, 2) %>
                                  </th>
                              </tr>

                          </tfoot>
                          
                      </table>

                    </div>
                  </div>
              </div>
          </div>


      </div>
  </div>  



 <!--
    
     '************** STI OG MAPPE ******************
	
<!--
    strTxtExport = strTxtExport & "Kontakt;Kontakt id;Jobnavn;Jobnr.;Status;Startdato;Slutdato;Prioitet;"

	Set objFSO = server.createobject("Scripting.FileSystemObject")
	
	if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\job_eksport.asp" then
							
		Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", True, False)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", 8)
	
	else
		
		Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", True, false)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", 8, -1)
		
	end if
	
	file = "jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&""
    
    	objFSO.WriteLine(strTxtExport)
    objFSO.WriteLine(expTxt)
	'Response.write strTxtExport
	
	objFSO.close
    
    Response.redirect "../inc/log/data/"& file &""	
     %> -->

<!--#include file="../inc/regular/footer_inc.asp"-->