

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->



<%
 '**** Søgekriterier AJAX **'
        'section for ajax calls

    if Request.Form("AjaxUpdateField") = "true" then
    Select Case Request.Form("control")


    case "godkenduge"

    tid = request("tid")
    'medid = request("medid")
    startdato = request("startdato")
    slutdato = request("slutdato")
    aktid = request("aktid")
    godkendjobid = request("godkendjobid")

    strSQLmedarb = "SELECT mnavn FROM medarbejdere WHERE mid ="& session("mid")
    oRec.open strSQLmedarb, oConn, 3
    if not oRec.EOF then
    usemedarbejder = oRec("mnavn")
    end if
    oRec.close

    redigerdato = Year(now) &"-"& Month(now) &"-"& Day(now) 

    strSQLgodkend = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& usemedarbejder &"', godkendtdato = '"& redigerdato &"'"_
    &" WHERE tid = "& tid
    '&" WHERE tmnr ="& medid &" AND TAktivitetId = "& aktid &" AND tjobnr = '"& godkendjobid &"' AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'"
    'Response.Write (strSQLgodkend)
    'Response.end 
    oConn.execute(strSQLgodkend)


    case "afvisuge"

    tid = request("tid")
    'medid = request("medid")
    startdato = request("startdato")
    slutdato = request("slutdato")
    aktid = request("aktid")
    godkendjobid = request("godkendjobid")

    strSQLmedarb = "SELECT mnavn FROM medarbejdere WHERE mid ="& session("mid")
    oRec.open strSQLmedarb, oConn, 3
    if not oRec.EOF then
    usemedarbejder = oRec("mnavn")
    end if
    oRec.close

    redigerdato = Year(now) &"-"& Month(now) &"-"& Day(now) 

    strSQLgodkend = "UPDATE timer SET godkendtstatus = 2, godkendtstatusaf = '"& usemedarbejder &"', godkendtdato = '"& redigerdato &"'"_
    &" WHERE tid = "& tid
    ' &" WHERE tmnr ="& medid &" AND TAktivitetId = "& aktid &" AND tjobnr = '"& godkendjobid &"' AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'"
    'Response.Write (strSQLgodkend)
    'Response.end 
    oConn.execute(strSQLgodkend)


    case "emailnoti"


  

    afsenderMid = session("mid")
    modtagerMid = request("medid")
    kommentar = request("kommentar")


                        'Response.Write "HER 7 modtagerMid: " & modtagerMid
                        'Response.end

                    '*** Henter afsender **
				        strSQL = "SELECT mnavn, email FROM medarbejdere"_
				        &" WHERE mid = "& afsenderMid
				        oRec.open strSQL, oConn, 3
            				
				        if not oRec.EOF then
            				
				        afsNavn = oRec("mnavn")
				        afsEmail = oRec("email")
            				
				        end if
				        oRec.close

                        '*** Henter modtager **
				        strSQL = "SELECT mnavn, email FROM medarbejdere"_
				        &" WHERE mid = "& modtagerMid
				        oRec.open strSQL, oConn, 3
            				
				        if not oRec.EOF then
            				
				        modtNavn = oRec("mnavn")
				        modtEmail = oRec("email")
            				
				        end if
				        oRec.close

                    Set myMail=CreateObject("CDO.Message")
                    myMail.Subject = "TimeOut - Declined project hours"
                    myMail.From = "timeout_no_reply@outzource.dk"
				                     

                    'myMail.To=strEmail
                    if len(trim(modtEmail)) <> 0 then
                    myMail.To= ""& modtNavn &"<"& modtEmail &">"
                    end if

                    strBody = "Hi " & modtNavn & vbCrLf & vbCrLf
                    strBody = strBody & "Declined project hours: " & vbCrLf & vbCrLf

                    tidsSQL = request("tids")
                    
                      '*** Henter modtager **
				        strSQL = "SELECT tjobnavn, tjobnr, taktivitetnavn, timer, tdato, godkendtstatusaf, godkendtdato FROM timer "_
				        &" WHERE tid = 0 " & tidsSQL
                        
                        'Response.Write strSQL
                        'Response.end

				        oRec.open strSQL, oConn, 3
            				
				        while not oRec.EOF
            				
				        strBody = strBody & oRec("tdato") &" - "& oRec("tjobnavn") &" ("& oRec("tjobnr") &") - "& oRec("taktivitetnavn") &" - "& oRec("timer") & vbCrLf 
            				
                        oRec.movenext
                        wend
				        oRec.close
    
                    'strBody = strBody & thisWeekTxt &" "& thisWeek &" - er afvist" & vbCrLf & vbCrLf
                    'strBody = strBody &"Begrundelse: " & vbCrLf 
					strBody = strBody & vbCrLf & vbCrLf
                    strBody = strBody & "Comment:" & vbCrLf & kommentar
                    strBody = strBody & vbCrLf & vbCrLf
		            strBody = strBody &"Best regards" & vbCrLf
		            strBody = strBody & afsNavn & ", "& afsEmail & vbCrLf & vbCrLf                       

                    myMail.TextBody= strBody

                        
                    myMail.Configuration.Fields.Item _
                    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                    'Name or IP of remote SMTP server
                                   
                    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                        smtpServer = "webout.smtp.nu"
                    else
                        smtpServer = "formrelay.rackhosting.com" 
                    end if
                    
                    myMail.Configuration.Fields.Item _
                    ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                    'Server port
                    myMail.Configuration.Fields.Item _
                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                    myMail.Configuration.Fields.Update
                    
                    if len(trim(modtEmail)) <> 0 then
                    myMail.Send
                    end if
                    set myMail=nothing



    end select     
    
    response.end             
    end if




    
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

    'select case func
        'case "db"
        'response.Redirect "godkend_job_timer.asp?aar="&aar&"&aarslut="&aarslut&"&FM_job="&jobid

        'Response.write "<br><br><a href='godkend_job_timer.asp?aar="&aar&"&aarslut="&aarslut&"&FM_job="&jobid &"'>Timer Godkendt - Videre >></a>"
        'Response.end
    'end select
     
%>
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

 <script src="js/godkend_job4.js" type="text/javascript"></script>
<style>
    .aktivilist:hover,
    .aktivilist:focus {
    text-decoration: none;
    cursor: pointer;
    }
</style>


<%call menu_2014 %>


  
  <div class="wrapper">
      <div class="content">
         
          <div class="container">
              <div class="portlet">
                  <h3 class="portlet-title"><u><%=godkend_txt_001 %></u> </h3>
                  <div class="portlet-body">

                   <!--<div id="div_jobid">DDD</div>-->

                      <div class="well">
                        <form action="godkend_job_timer.asp?sogsubmitted=1" method="POST">

                             
                        

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
                                Filter: <select name="filter" class="form-control input-small" onchange="submit();">
                                        <option value="0" <%=filterSEL0%>>Show All</option>
                                        <option value="1" <%=filterSEL1%>>Show only Approved</option>
                                        <option value="2" <%=filterSEL2%>>Show only Declined</option>
                                        </select>
               
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
                           strSQLjobid = "SELECT jobnr FROM job WHERE id ="& jobid
                           oRec4.open strSQLjobid, oConn, 3 

                           if not oRec4.EOF then
                          
                            Strjobid = oRec4("jobnr")

                           end if
                          
                          oRec4.close 
                          'response.Write "test" & Strjobid
                   
                        
                      %>

                      <!--onClick="reloadpage();"-->

                      <form action="godkend_job_timer.asp?func=db&aar=<%=aar %>&aarslut=<%=aarslut %>&FM_job=<%=jobid %>" method="post" id="godkendform" name="godkendform" >
                      <input type="hidden" value="<%=Strjobid %>" id="godkendjobid" />
                      <table style="background-color:white;" class="table dataTable table-striped table-bordered table-hover ui-datatable">

                          <thead>
                              <tr>
                                    <th><%=godkend_txt_006 %></th>

                                  <%

                                      


                                      selectedstart_month = Month(aar)
                                      selectedstart_year = Year(aar)

                                      
                                      'response.write selectedstart_month

                                    startdate = "01/" & selectedstart_month & "/" & selectedstart_year

                                    startdato = Year(aar) &"-"& Month(aar) &"-"& Day(aar)
                                    slutdato = Year(aarslut) &"-"& Month(aarslut) &"-"& Day(aarslut)
                                    response.Write "<input type=""hidden"" id=""startdato"" value="&startdato&" />"
                                    response.Write "<input type=""hidden"" id=""slutdato"" value="&slutdato&" />"

                                      months = selectedstart_month
                                      Stryear = selectedstart_year 
                                      
                                      'response.write startdate 

                                      for i = 0 to antalmaaned

                                            months = months + 1
                                            
                                            if months > 13 then
                                                months = 2
                                                Stryear = Stryear + 1
                                            end if
                                           
                                            select case months -1
                                            case 1
                                            getdate = godkend_txt_011 &"<br>"& Stryear
                                            case 2 
                                            getdate  = godkend_txt_012 &"<br>"& Stryear
                                            case 3 
                                            getdate  = godkend_txt_013 &"<br>"& Stryear
                                            case 4 
                                            getdate  = godkend_txt_014 &"<br>"& Stryear
                                            case 5 
                                            getdate  = godkend_txt_015 &"<br>"& Stryear
                                            case 6 
                                            getdate  = godkend_txt_016 &"<br>"& Stryear
                                            case 7 
                                            getdate  = godkend_txt_017 &"<br>"& Stryear
                                            case 8 
                                            getdate  = godkend_txt_018 &"<br>"& Stryear
                                            case 9 
                                            getdate  = godkend_txt_019 &"<br>"& Stryear
                                            case 10 
                                            getdate  = godkend_txt_020 &"<br>"& Stryear
                                            case 11 
                                            getdate  = godkend_txt_021 &"<br>"& Stryear
                                            case 12 
                                            getdate  = godkend_txt_022 &"<br>"& Stryear
                                            end select    

                                            'getdate = Monthname(months -1 ,True) & "<br>" & Stryear
                                            'getdate = (Ucase(getdate))   
                                             'getdate = getdate & "<br>" & Stryear 
                                                                          
                                        %> <th style="text-align:center;">
                                            Hours
                                            
                                            <!--<%=UCase(Left(getdate,1)) & Mid(getdate,2) %>-->

                                         
     
        

                                           </th> 
                                  <%

                                      next
                                  %>

                                  <th style="width:150px; text-align:center;">Comment</th>
                                  <th style="text-align:center;"><!--<%=godkend_txt_008 %><br />Declined<br />--> Dates<br />
                                       <span style="font-size:9px; font-weight:lighter; color:#999999;">
                                       <%=formatdatetime(aar, 1) & " <br> " & formatdatetime(aarslut, 1) %>

                                         <%if brugugeCHK = "CHECKED" then %>
                                           <br /> (Week: <%=bruguge_week %>)
                                            <%end if %>
                                           </span>

                                  </th>
                                    <th><%=godkend_txt_007 %></th>
                                  <th style="text-align:center;"><%=godkend_txt_009 %> <input type="checkbox" id="godkendalle" class="godkendalle" name="godkendalle" /></th>
                                  <th style="text-align:center;">Decline <input type="checkbox" id="afvisalle" name="godkendalle" class="afvisalle" /></th>
                                  <th style="text-align:center;" >App./Decl. Date</th>
                              </tr>

                          </thead>

                          <tbody>
                            
                              <%
                                 
                               

                                  

                                    
                                  lastmid = 0
                                  m = 1000
                                  d = 10000
                                  lasttmnr = 0
                                  dim timer_md, dato_medid, medarbid, medidnavn, manedstot
                                  Redim timer_md(m,d), dato_medid(m,d), medarbid(m), medidnavn(m), manedstot(antalmaaned)
                                  
                                  strSQL = "SELECT tmnr, tmnavn, tdato, sum(timer) as Timer FROM timer WHERE ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' AND tmnr is not null GROUP by tmnr, year(tdato), month(tdato) Order by tmnr, tdato "
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
                                        <tr id="aktvisfont_<%=medarbid(m) %>">
                                            <td style="white-space:nowrap; width:350px;"><%=left(medidnavn(m), 20) & " ["& meInit &"]" %> <span class="fa fa-chevron-down pull-right aktivilist" id="<%=medarbid(m) %>"></span></td>

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
                                                        <td style="text-align:right;">&nbsp; <!--<%=tdval %>--> </td>
                                                        <%
                                                        
                                                next
                                                
                                            %>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td style="text-align:center;">
                                                <%

                                                    strSQLtimer = "SELECT sum(timer) as Timer FROM timer WHERE ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND tmnr ="& medarbid(m) & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'"
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
                                           
                                            <td style="text-align:center"><input type="checkbox" class="godkendmedarb godkendmedarb_<%=medarbid(m)%>" id="medarbgodkend_<%=medarbid(m) %>" /></td> 
                                            <td style="text-align:center"><input type="checkbox" class="afvismedarb afvismedarb_<%=medarbid(m)%>" id="medarbafvis_<%=medarbid(m) %>" /></td> 
                                            <td>&nbsp;</td>
                                        </tr>
                                        <tbody style="display:none; font-size:95%" id="aktivilist_<%=medarbid(m) %>">
                                        <%
                                        select case filterThis
                                          case "1"
                                          filterSQL = " AND godkendtstatus = 1"
                                            case "2"
                                          filterSQL = " AND godkendtstatus = 2"
                                            case else
                                            filterSQL = ""
                                         end select 
                                        

                                        strSQLakt = "SELECT tid, taktivitetid, taktivitetnavn, timer, tdato, godkendtdato, godkendtstatus, godkendtstatusaf, timerkom FROM timer WHERE ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND tmnr ="& medarbid(m) & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& filterSQL &" GROUP BY taktivitetId, tdato ORDER BY taktivitetnavn"
                                        'response.Write strSQLtimer
                                        'response.Flush
                                        oRec.open strSQLakt, oConn, 3
                                        while not oRec.EOF

                                        tid = oRec("tid")
                                        aktid = oRec("taktivitetId")
                                        aktnavn = oRec("taktivitetnavn")
                                        aktdato = formatdatetime(oRec("tdato"), 2)
                                        akttimer = formatnumber(oRec("timer"), 2)
                                        aktgodkendtdato = formatdatetime(oRec("godkendtdato"), 2)
                                        aktgodkendtstatus = oRec("godkendtstatus")
                                        aktgodkendtstatusaf = oRec("godkendtstatusaf")
                                        aktTimerKom = oRec("timerkom")


                                        if cDate(aktgodkendtdato) >= "2002-01-01" then
                                        aktgodkendtdato = aktgodkendtdato
                                        else
                                        aktgodkendtdato = ""
                                        end if

                                        'response.Write "<br>id: " & aktid  
                                        %>
                                      
                                        <tr>                                      
                                            <td>&nbsp&nbsp-<span style="padding-left:25px;"><%=aktnavn %></span></td>
                                            <%
                                            aktstartmonth = Month(startdato)
                                            aktstartyear = Year(startdato)
                                            aktslutmonth = Month(slutdato)
                                            aktslutyear = Year(slutdato)  
                                            'response.Write aktstartmonth & aktstartyear & ":::"  
                                            for i = 0 to antalmaaned
                                            'response.Write "her:" & i
                                            if i = 0 then
                                            aktstartdato = startdato
                                            'aktslutdato = dateadd("m",1,aktstartdato)
                                            else 
                                            aktstartdato = dateAdd("m", i, "1-"& aktstartmonth &"-"& aktstartyear)
                                            aktstartdato= dateAdd("d", -1, aktstartdato)
                                            aktstartdato = year(aktstartdato) & "-" & month(aktstartdato) & "-" & day(aktstartdato)
                                            'aktslutdato = dateadd("m",1,aktstartdato)                                                      
                                            end if
                                            aktstartdatoUS = Year(aktstartdato) &"-"& Month(aktstartdato) &"-"& Day(aktstartdato)

                                            if i = antalmaaned then
                                            aktslutdato = slutdato
                                            else
                                            aktslutdato = dateAdd("m", i + 1, "1-"& aktstartmonth &"-"& aktstartyear)
                                            aktslutdato= dateAdd("d", -1, aktslutdato)
                                            aktslutdato = year(aktslutdato) & "-" & month(aktslutdato) & "-" & day(aktslutdato)
                                            end if
                                            aktslutdatoUS = Year(aktslutdato) &"-"& Month(aktslutdato) &"-"& Day(aktslutdato)
                                          
                                        
                                               
                                            'TimerPaaAkt = 0                                     
                                            'strSQLakttimer = "SELECT taktivitetnavn, sum(timer) as Timer FROM timer WHERE TAktivitetId ="& aktid &" AND tfaktim <> 5 AND tjobnr = '"& Strjobid & "' AND tmnr ="& medarbid(m) & " AND tdato BETWEEN '"& aktstartdatoUS &"' AND '"& aktslutdatoUS &"'" 
                                            'response.Write strSQLakttimer
                                           ' oRec2.open strSQLakttimer, oConn, 3
                                            'if not oRec2.EOF then
                                            '    TimerPaaAkt = oRec2("Timer")                                                                                                                                       
                                            'end if
                                            'oRec2.close 
                                            %>
                                            <td style="text-align:center;"><%=formatnumber(akttimer, 2) %>

                                               

                                            </td>
                                            <td style="text-align:left;">

                                                 <%if len(trim(aktTimerKom)) <> 0 then %>
                                                <span style="font-size:9px; font-weight:lighter;"><%=left(aktTimerKom, 250) %></span>
                                                <%end if %>
                                                &nbsp;

                                            </td>

                                            <%next %>


                                             <td style="text-align:center;"><%=aktdato%>
                                                <%
                                                    'sidstgodkendt = ""
                                                    'strSQLdato = "SELECT godkendtdato FROM timer WHERE ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND tmnr ="& medarbid(m) &" AND TAktivitetId ="& aktid & " AND godkendtstatus <> 0 AND godkendtdato > '2002-01-01' "
                                                    'oRec2.open strSQLdato, oConn, 3
                                                    'While not oRec2.EOF
                                                    '    sidstgodkendt = formatdatetime(oRec2("godkendtdato"), 2)'Day(oRec2("godkendtdato"))&"-"&Month(oRec2("godkendtdato"))&"-"&Year(oRec2("godkendtdato"))
                                                    'oRec2.movenext
                                                    'wend
                                                    'oRec2.close       
                                                                                               
                                                    'response.Write sidstgodkendt
                                                    'response.Write "awd" & sidstgodkendt
                                                    
                                                %>
                                            </td>

                                            <%

                                            if cdbl(lastAktid) <> cdbl(aktid) then
                                                    akTimerTot = 0
                                                    strSQLakttotal = "SELECT taktivitetnavn, sum(timer) as Timer FROM timer WHERE TAktivitetId ="& aktid &" AND ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND tmnr ="& medarbid(m) & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'"
                                                    oRec2.open strSQLakttotal, oConn, 3
                                                    if not oRec2.EOF then
                                                    akTimerTot = oRec2("Timer")
                                                    end if
                                                    oRec2.close 
                                                %>
                                            <td style="text-align:center">
                                                <%=formatnumber(akTimerTot,2) %>
                                            </td>

                                            <%else %>

                                               <td style="text-align:right">
                                              &nbsp;
                                            </td>

                                            <%end if %>

                                           
                                            <%
                                            select case cint(aktgodkendtstatus)
                                            case 0
                                                aktgodkendtstatus1CHK = ""
                                                aktgodkendtstatus2CHK = ""
                                            case 1
                                                aktgodkendtstatus1CHK = "CHECKED"
                                                aktgodkendtstatus2CHK = ""
                                            case 2
                                                 aktgodkendtstatus1CHK = ""
                                                aktgodkendtstatus2CHK = "CHECKED"
                                            end select    
                                            %>

                                            <td style="text-align:center">                                               
                                                <input type="radio" id="godkendstatus_<%=medarbid(m) %>_<%=aktid %>_<%=tid %>" class="godkendbox godkendbox_<%=medarbid(m)%>" name="godkendbox_<%=tid %>" value="<%=medarbid(m) %>_<%=aktid %>_<%=tid %>" <%=aktgodkendtstatus1CHK %> /></td> <!-- <%=medarbid(m)%>_<%=aktid %> --> 
                                            <td style="text-align:center">                                               
                                                <input type="radio" id="afvisstatus_<%=medarbid(m) %>_<%=aktid %>_<%=tid %>" class="afvisbox afvisbox_<%=medarbid(m)%>" name="godkendbox_<%=tid %>" value="<%=medarbid(m) %>_<%=aktid %>_<%=tid %>" <%=aktgodkendtstatus2CHK %> /></td> <!-- <%=medarbid(m)%>_<%=aktid %> -->
                                       
                                            <td style="text-align:center"><span style="font-size:9px;"><%=aktgodkendtdato %><br />
                                                <%=left(aktgodkendtstatusaf, 10) %></span>
                                            </td>    
                                        </tr>
                                        <%        
                                            
                                        lastAktid = aktid                     
                                        oRec.movenext
                                        wend
                                        oRec.close 
                                        %>
                                       

                                   <tr>
                                  <td>&nbsp;</td>
                                   <%

                                      for i = 0 to antalmaaned  

                                        %> <td style="text-align:right;">&nbsp;</td> <%

                                      next 
                                  %>
                                  <td>&nbsp;</td>
                                   <td>&nbsp;</td>
                                  <td>&nbsp;</td>
                                  <td style="text-align:right" colspan="3">Comment for declined<br /><textarea id="decline_comment_<%=medarbid(m) %>" name="decline_comment" class="form-control input-small"></textarea>
                                        <input type="hidden" value="" id="decl_tids_<%=medarbid(m) %>" />
                                        <input type="hidden" value="<%=medarbid(m) %>" id="decl_tids_mid_<%=medarbid(m) %>" class="decl_tids" />
                                      </td>
                              </tr>
                                 </tbody>
                                  <%
                                        'end if
                                  
                                     lastmid = medarbid(m)    
                                  next
       
                              %>
                              
                         

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
                                          strSQLtimertotal = "SELECT sum(timer) as Timer FROM timer WHERE ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY tjobnr"

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
                                  <th>&nbsp;</th>
                                  <th>&nbsp;</th>
                                  <th>&nbsp;</th>
                                   <th>&nbsp;</th>
                                  <th style="text-align:center; width:20px;"><a class="btn btn-success btn-sm" id="godkendknap">Submit</a><!-- =godkend_txt_009  --></th>
                                
                                   </tr>

                          </tfoot>
                          
                      </table>
                      </form>
                    </div>
                  </div>
              </div>
          </div>

      
      </div>
  <!--</div>-->




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