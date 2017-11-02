<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->


<%
if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if

    func = request("func")


    call akttyper2009(2)
    'if len(trim(request("FM_medarb"))) <> 0 then
    'usemrn = request("FM_medarb")
    'else
    usemrn = session("mid") 
    'end if

    brugugeCHK = ""
    if len(trim(request("bruguge"))) <> 0 OR (request.cookies("2015")("bruguge") = "1" AND len(trim(request("sogsubmitted"))) = 0) then

        brugugeCHK = "CHECKED"

        if len(trim(request("bruguge_week"))) <> 0 then 
        bruguge_week = request("bruguge_week")
        else
        bruguge_week = request.cookies("2015")("bruguge_week")
        end if

        response.cookies("2015")("bruguge_week") = bruguge_week
	                  
                        
	            stDato = "1/1/"&year(now)
	            datoFound = 0
	    
	            for u = 1 to 53 AND datoFound = 0
	   
	            if u = 1 then
	            tjkDato = stDato
	            else
	            tjkDato = dateadd("d",7,tjkDato)
	            end if
	    
                'response.Write "tjkDato: " & tjkDato & "<br>"

	            tjkDatoW = datepart("ww", tjkDato, 2,2)
	    
	            if cint(bruguge_week) = cint(tjkDatoW) then
	    
	            wDay = datepart("w", tjkDato, 2,2)
	       
	        
	                select case wDay
	                case 1
	                tjkDato = dateAdd("d", 0, tjkDato)
	                case 2
	                tjkDato = dateAdd("d", -1, tjkDato)
	                case 3
	                tjkDato = dateAdd("d", -2, tjkDato)
	                case 4
	                tjkDato = dateAdd("d", -3, tjkDato)
	                case 5
	                tjkDato = dateAdd("d", -4, tjkDato)
	                case 6
	                tjkDato = dateAdd("d", -5, tjkDato)
	                case 7
	                tjkDato = dateAdd("d", -6, tjkDato)
	                end select
	    
	            stDaguge = day(tjkDato)
	            stMduge = month(tjkDato)
	            stAaruge = year(tjkDato)
	    
	            tjkDato_slut = dateadd("d", 6, tjkDato)
	    
	            slDaguge = day(tjkDato_slut)
	            slMduge = month(tjkDato_slut)
	            slAaruge = year(tjkDato_slut)
	    
	       
	            datoFound = 1
	    
	            end if
	    
	            next

         aar = stDaguge &"-"& stMduge &"-"& stAaruge
        aarslut = slDaguge &"-"& slMduge &"-"& slAaruge

        response.cookies("2015")("bruguge") = "1"

    else

        response.cookies("2015")("bruguge") = ""

        bruguge_week = datepart("ww", now, 2,2)

            if len(trim(request("aar"))) <> 0 then
            aar = request("aar")

                response.cookies("2015")("gktimer_aar") = aar

            else

                if request.cookies("2015")("gktimer_aar") <> "" then
                aar = request.cookies("2015")("gktimer_aar")
                else
                aar = "1-"& month(now) &"-"& year(now)
                end if

            end if

            if len(trim(request("aarslut"))) <> 0 then
            aarslut = request("aarslut")
            else
   
                if request.cookies("2015")("gktimer_aarslut") <> "" then
                aarslut = request.cookies("2015")("gktimer_aarslut")
                else
                aarslut = dateAdd("m", 1, aar)
                end if

            end if


    end if


    'if len(trim(request("FM_job"))) <> 0 then
    'jost0CHK = ""
    'else
    'viskunabnejob0 = 1
    'jost0CHK = "CHECKED"
    'end if


    if len(trim(request("filter"))) <> 0 then
        filterThis = request("filter")
        response.cookies("2015")("gktimer_filter") = filterThis
    else
        if request.cookies("2015")("gktimer_filter") <> "" then
        filterThis = request.cookies("2015")("gktimer_filter")
        else
        filterThis = "0"
        end if
    end if

    filterSEL0 = ""
    filterSEL1 = ""
    filterSEL2 = ""
    select case filterThis
    case "0"
        filterSEL0 = "SELECTED"
        case "1"
        filterSEL1 = "SELECTED"
        case "2"
        filterSEL2 = "SELECTED"
    end select


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

    select case func
    case "XXdb"
        
        
        
        
        
        'response.Redirect "godkend_job_timer_2017.asp?aar="&aar&"&aarslut="&aarslut&"&FM_job="&jobid

        Response.write "<br><br><a href='godkend_job_timer_2017.asp?aar="&aar&"&aarslut="&aarslut&"&FM_job="&jobid &"'>The Work is done, continue >></a>"
        'Response.end

    case else
    
     
%>

<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<%call menu_2014 %>

<style>
    .medarbliste:hover,
    .medarbliste:focus {
    text-decoration: none;
    cursor: pointer;
    }
</style>

<div class="wrapper">
    <div class="content">

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Projekt timer</u></h3>

                <div class="portlet-body">
                    <div class="well">
                        <form action="akt_timer.asp?sogsubmitted=1" method="POST">

                             
                        

                            <div class="row">
                                <div class="col-lg-2">
                                    <h4 class="panel-title-well"><%=dsb_txt_002 %></h4>
                                </div>
                                <div class="col-lg-10">&nbsp</div>                 
                            </div>

                
                            <div class="row">
                          

                                <div class="col-lg-6"><br /><%=godkend_txt_002 %>:
                                
                                    <%

                                        select case lto
                                        case "tia"
                                        strSQLorderBy = "jobnr, jobnavn"

                                        salgsansvarligSQLKri = " OR salgsans1 = " & session("mid") 
                                        salgsansvarligSQLKri = salgsansvarligSQLKri & " OR salgsans2 = " & session("mid")
                                        salgsansvarligSQLKri = salgsansvarligSQLKri & " OR salgsans3 = " & session("mid")
                                        salgsansvarligSQLKri = salgsansvarligSQLKri & " OR salgsans4 = " & session("mid")
                                        salgsansvarligSQLKri = salgsansvarligSQLKri & " OR salgsans5 = " & session("mid")

                                        case else
                                        strSQLorderBy = "jobnavn"
                                        salgsansvarligSQLKri = ""
                                        end select


                                        strSQLjob = "SELECT id, jobnr, jobnavn FROM job WHERE jobstatus = 1 AND "_
                                        &" (jobans1 = "& session("mid") &" OR jobans2 = "& session("mid") &" OR jobans3 = "& session("mid") &""_
                                        &" OR jobans4 = "& session("mid") &" OR jobans5 = "& session("mid") &" "& salgsansvarligSQLKri &") ORDER BY "& strSQLorderBy          
                                    
                                        'Response.write strSQLjob
                                        'response.flush
                                                                            
                                    %><br />
                                    <select name="FM_job" id="FM_job" <%=progrpmedarbDisabled  %> class="form-control input-small" onchange="submit();">

                                        <%select case lto
                                        case "tia" %>
                                        <option value="0" DISABLED">Select your PM projects</option>
                                      <%end select

                                        oRec3.open strSQLjob, oConn, 3
                                        while not oRec3.EOF
                                
                                        Strjobnavn = oRec3("jobnavn")

                                        if cdbl(jobid) = cdbl(oRec3("id")) then
				                        isSelected = "SELECTED"
				                        else
				                        isSelected = ""
				                        end if
				    
                                        select case lto
                                        case "tia"
                                        jobnavnid = oRec3("jobnr") &" - "& oRec3("jobnavn")
                                        case else 
				                        jobnavnid = oRec3("jobnavn") & " ("& oRec3("jobnr") &")"
                                        end select
        
				                    %>
				                    <option value="<%=oRec3("id")%>" <%=isSelected%>><%=jobnavnid%></option>
				                    <%

                                    oRec3.movenext
                                    wend
                                    oRec3.close 

                                    %>

                                    </select>

                                    <br />
                                  <!--  Filter: <select name="filter" class="form-control input-small" onchange="submit();">
                                            <option value="0" <%=filterSEL0%>>Show All</option>
                                            <option value="1" <%=filterSEL1%>>Show only Approved</option>
                                            <option value="2" <%=filterSEL2%>>Show only Declined</option>
                                            </select> -->
               
                                </div>

                                <%
                                     antalmaaned = 0 'DateDiff("m",aar,aarslut, 2,2)
                                    antalaar = 0 'DateDiff("yyyy",aar,aarslut, 2,2)
                                                                                                   
                                %>

                            
                                
                                 
                           
                                <div class="col-lg-2">
                                    <b>Period or week:</b><br />
                                    <%=godkend_txt_003 %>:
                                    <div class='input-group date' id='datepicker_stdato'>
                                    <input type="text" class="form-control input-small" name="aar" value="<%=aar %>" placeholder="dd-mm-yyyy" />
                                    <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar">
                                            </span>
                                        </span>
                                    </div>
                               </div>
                                     <div class="col-lg-2">
                                   <br />
                                    <%=godkend_txt_004 %>:<br />
                                    <div class='input-group date' id='datepicker_sldato'>
                                    <input type="text" class="form-control input-small" name="aarslut" value="<%=aarslut %>" placeholder="dd-mm-yyyy" />
                                    <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar">
                                            </span>
                                        </span>
                                    </div>
                               </div>
                 
                            <div class="col-lg-1">
                            <br />
                            <input id="bruguge" name="bruguge" type="checkbox" value="1" <%=brugugeCHK %> />  Week:
			                <select name="bruguge_week" class="form-control input-small" onchange="submit();">
			                <% for w = 1 to 52
			
			                if w = 1 then
			                useWeek = 1
			                else
			                useWeek = dateadd("ww", 1, useWeek)
			                end if
			
			                if cint(datepart("ww", useWeek, 2,2)) = cint(bruguge_week) then
			                wSel = "SELECTED"
			                else
			                wSel = ""
			                end if
			
			                %>
			                <option value="<%=datepart("ww", useWeek, 2,2) %>" <%=wSel%>><%=datepart("ww", useWeek, 2,2) %></option>
			                <%
			                next %>
			
			                </select>
                        </div>
                   
                            <div class="col-lg-1">&nbsp;<br /><br /><button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=godkend_txt_005 %> >></b></button></div>  
                        </div>
                                             
                        </form>
                      </div>


                    <%
                        'response.Write "jobid: " & jobid & aar & aarslut  
                        
                        startdato = Year(aar) &"-"& Month(aar) &"-"& Day(aar)
                        slutdato = Year(aarslut) &"-"& Month(aarslut) &"-"& Day(aarslut)
                        
                        strSQLjobnr = "SELECT jobnr FROM job WHERE id ="& jobid
                        oRec.open strSQLjobnr, oConn, 3 
                        if not oRec.EOF then                         
                        Strjobnr = oRec("jobnr")
                        end if                          
                        oRec.close     
                        
                    %>

                    <script src="js/akt_timer_medarbliste.js" type="text/javascript"></script>
                    <table style="background-color:white;" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        
                        <thead>
                            <tr>
                                <th>Aktivitet</th>
                                <th>Timer</th>
                            </tr>                           
                        </thead>

                        <tbody>
                            <%
                            strSQLhentakt = "SELECT navn, id FROM aktiviteter WHERE job ="&jobid
                            oRec.open strSQLhentakt, oConn, 3
                            while not oRec.EOF 
                            aktnavn = oRec("navn")
                            aktid = oRec("id")

                            %>
                                
                           <tr>
                               <td><%=aktnavn %> <span class="fa fa-chevron-down pull-right medarbliste" id="<%=aktid %>"></span></td>
                               <td>
                                   <%
                                       strSQLhenttimer = "SELECT sum(timer) as timer FROM timer WHERE tjobnr = "& Strjobnr &" AND TAktivitetId ="& aktid &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'"
                                       'response.Write strSQLhenttimer & Strjobnr
                                       oRec2.open strSQLhenttimer, oConn, 3
                                       if not oRec2.EOF then
                                       Strtimer = oRec2("timer")
                                       if Strtimer <> 0 then
                                       else
                                       Strtimer = 0
                                       end if

                                       response.Write Strtimer
                                       end if
                                       oRec2.close 
                                   %>
                               </td>
                           </tr>

                           <tbody style="display:none" id="tr_medarbliste_<%=aktid %>">
                                 <%
                                     strSQLMedarb = "SELECT Tmnavn, Tmnr FROM timer WHERE tjobnr = "& Strjobnr &" AND TAktivitetId ="& aktid &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'"                                   
                                     oRec2.open strSQLMedarb, oConn, 3
                                     While not oRec2.EOF
                                    %>

                                    <tr>                                       
                                        <td><%=oRec2("Tmnavn") %></td>
                                        <td>
                                            <%
                                                strSQLmedarbtimer = "SELECT sum(timer) as timer FROM timer WHERE tjobnr = "& Strjobnr &" AND TAktivitetId ="& aktid &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' AND Tmnr ="&oRec2("Tmnr")
                                                oRec3.open strSQLmedarbtimer, oConn, 3
                                                if not oRec3.EOF then
                                                response.Write oRec3("timer")
                                                end if                                               
                                            %>
                                        </td>
                                    </tr>

                                    <%         
                                     oRec2.movenext
                                     wend
                                     oRec2.close 
                                    %>                                                            
                           </tbody>

                            <%
                            oRec.movenext
                            wend
                            oRec.close
                            %>     
                            
                        </tbody>

                    </table>


                </div>
            </div>
        </div>

    </div>
</div>

<%end select %>