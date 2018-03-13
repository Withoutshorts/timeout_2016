

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->



<%
 '**** Søgekriterier AJAX **'
        'section for ajax calls

    if Request.Form("AjaxUpdateField") = "true" then
    Select Case Request.Form("control")


    case "godkenduge"


    

    multi = request("multi")
    if multi = "1" then
        tid = request("tid")
        tids = split(tid, ", ")
        
        tidsSQLKri = " (tid = 0"
        for t = 1 TO UBOUND(tids)
           
             '*** er den allerede godkendt? = DO NOTHING
            strSQLgodkendfindes = "SELECT tid FROM timer"
            strSQLgodkendfindes = strSQLgodkendfindes &" WHERE godkendtstatus = 1 AND tid = "& tids(t)  
    
            findes = 0
            oRec.open strSQLgodkendfindes, oConn, 3
            if not oRec.EOF then
            findes = 1
            end if
            oRec.close

            if cint(findes) = 0 then
            tidsSQLKri = tidsSQLKri & " OR tid = "& tids(t)  
            end if

        next
        tidsSQLKri = tidsSQLKri & ")"

    else
    tid = request("tid")
    end if

    'medid = request("medid")
    'startdato = request("startdato")
    'slutdato = request("slutdato")
    'aktid = request("aktid")
    'godkendjobid = request("godkendjobid")

    strSQLmedarb = "SELECT mnavn FROM medarbejdere WHERE mid ="& session("mid")
    oRec.open strSQLmedarb, oConn, 3
    if not oRec.EOF then
    usemedarbejder = oRec("mnavn")
    end if
    oRec.close

    redigerdato = Year(now) &"-"& Month(now) &"-"& Day(now) 


    strSQLgodkend = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& usemedarbejder &"', godkendtdato = '"& redigerdato &"'"

    if multi = "1" then
    strSQLgodkend = strSQLgodkend &" WHERE ("& tidsSQLKri &") AND overfort = 0"
    else
    strSQLgodkend = strSQLgodkend &" WHERE tid = "& tid & " AND overfort = 0"
    end if

    'Response.Write "HER 2 " & strSQLgodkend
    'Response.end 
    oConn.execute(strSQLgodkend)


    case "afvisuge"

    

     multi = request("multi")
    if multi = "1" then
        tid = request("tid")
        tids = split(tid, ", ")
        
        tidsSQLKri = " (tid = 0"
        for t = 1 TO UBOUND(tids)
             
            '*** er den allerede AFVIST? = DO NOTHING
            strSQLgodkendfindes = "SELECT tid FROM timer"
            strSQLgodkendfindes = strSQLgodkendfindes &" WHERE godkendtstatus = 2 AND tid = "& tids(t)  
    
            findes = 0
            oRec.open strSQLgodkendfindes, oConn, 3
            if not oRec.EOF then
            findes = 1
            end if
            oRec.close

            if cint(findes) = 0 then
            tidsSQLKri = tidsSQLKri & " OR tid = "& tids(t)  
            end if
 
        next
        tidsSQLKri = tidsSQLKri & ")"

    else
    tid = request("tid")
    end if

    
    'medid = request("medid")
    'startdato = request("startdato")
    'slutdato = request("slutdato")
    'aktid = request("aktid")
    'godkendjobid = request("godkendjobid")


    
    '*** smilaktiv = 1 AND autogk = 2 '**
    '*** Hvis timer bliver godkendt tentativt/godklendt direkte når en uge aflsutttes af medarbejder, skal denne uge godkendelse slettes igen.
    call smileyAfslutSettings()
    call autogktimer_fn()
    call ersmileyaktiv()

   
    strSQLmedarb = "SELECT mnavn FROM medarbejdere WHERE mid ="& session("mid")
    oRec.open strSQLmedarb, oConn, 3
    if not oRec.EOF then
    usemedarbejder = oRec("mnavn")
    end if
    oRec.close

    
    if cint(smilaktiv) = 1 AND (cint(autogk) = 1 OR cint(autogk) = 2) then



    redigerdato = Year(now) &"-"& Month(now) &"-"& Day(now) 
    strSQLgodkend = "UPDATE timer SET godkendtstatus = 2, godkendtstatusaf = '"& usemedarbejder &"', godkendtdato = '"& redigerdato &"'"
    
    if multi = "1" then
    strSQLgodkend = strSQLgodkend &" WHERE ("& tidsSQLKri &") AND overfort = 0"
    else
    strSQLgodkend = strSQLgodkend &" WHERE tid = "& tid & " AND overfort = 0"
    end if

    'Response.write strSQLgodkend
    'Response.end
    oConn.execute(strSQLgodkend)


    'Response.Write "cint(smilaktiv) = 1 AND cint(autogk) = 2: " & cint(smilaktiv) &" AND "& cint(autogk) 
           
             if multi = "1" then
            strSQLhentertimeDatoerKri = " WHERE ("& tidsSQLKri &") AND overfort = 0"
            else
            strSQLhentertimeDatoerKri = " WHERE tid = "& tid & " AND overfort = 0"
            end if

            strSQLhentertimeDatoer = "SELECT tdato, tmnr FROM timer " & strSQLhentertimeDatoerKri
           
            'strAlle = "<br>strSQLhentertimeDatoer: " & strSQLhentertimeDatoer & " SmiWeekOrMonth: " & SmiWeekOrMonth
            'Response.end

            oRec.open strSQLhentertimeDatoer, oConn, 3
            while not oRec.EOF

                                     

                                    select case cint(SmiWeekOrMonth) 
                                    case 0
					                
                                    denneuge = datepart("ww", oRec("tdato"), 2,2)
					                detteaar = datepart("yyyy", oRec("tdato"), 2,2)
                                    perSqlKri = "YEAR(uge) = '"& detteaar &"' AND WEEK(uge, 3) = '"& denneuge &"'"
                                    
                                    if cint(autogktimer) = 2 then 'Nulstil tentative timer
                                
                                                    thisUeId = 0
                                                    strSQLthisPeriode = "SELECT uge, mid, id FROM ugestatus WHERE mid = "& oRec("tmnr") &" AND ("& perSqlKri &")"
                                                    oRec6.open strSQLthisPeriode, oConn, 3
                                                    if not oRec6.EOF then

                                                    thisUeId = oRec6("id")

                                                    end if
                                                    oRec6.close

                                            call nulstilTentative(autogktimer, thisUeId)

                                    end if


                                    perSqlKriRem = "DELETE FROM ugestatus WHERE mid = "& oRec("tmnr") &" AND ("& perSqlKri &")"
                                    oConn.execute(perSqlKriRem)


                                    case 1

                                    dennemd = datepart("m", oRec("tdato"), 2,2)
					                detteaar = datepart("yyyy", oRec("tdato"), 2,2)

                                    perSqlKri = "YEAR(uge) = '"& detteaar &"' AND MONTH(uge) = '"& dennemd &"'"

                                        if cint(autogktimer) = 2 then 'Nulstil tentative timer
                                
                                                    thisUeId = 0
                                                    strSQLthisPeriode = "SELECT uge, mid, id FROM ugestatus WHERE mid = "& oRec("tmnr") &" AND ("& perSqlKri &")"
                                                    oRec6.open strSQLthisPeriode, oConn, 3
                                                    if not oRec6.EOF then

                                                    thisUeId = oRec6("id")

                                                    end if
                                                    oRec6.close

                                            call nulstilTentative(autogktimer, thisUeId)

                                         end if


                                    perSqlKriRem = "DELETE FROM ugestatus WHERE mid = "& oRec("tmnr") &" AND ("& perSqlKri &")"
                                    oConn.execute(perSqlKriRem)

                                    case 2 '** Dagligt: DO NOTHING 


                                    end select


                                    
            oRec.movenext
            wend
            oRec.close


            'Response.write "SQL:" & strAlle
            'Response.end
   


    end if


 
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

    call smileyAfslutSettings()

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

            response.cookies("2015")("gktimer_aarslut") = aarslut

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
    filterSEL3 = ""
    filterSEL4 = ""

    select case filterThis
    case "0"
        filterSEL0 = "SELECTED"
        case "1"
        filterSEL1 = "SELECTED"
        case "2"
        filterSEL2 = "SELECTED"
        case "3"
        filterSEL3 = "SELECTED"
        case "4"
        filterSEL4 = "SELECTED"
    end select

    select case filterThis
        case "1"
        filterSQL = " AND godkendtstatus = 1"
        case "2"
        filterSQL = " AND godkendtstatus = 2"
        case "3"
        filterSQL = " AND (godkendtstatus = 0 OR godkendtstatus = 3)"
        case "4"
        filterSQL = " AND godkendtstatus = 3" 'Tentative
        case else
        filterSQL = ""
    end select


    if len(trim(request("aproveviewfiler"))) <> 0 then
        aproveview = request("aproveviewfiler")
        response.cookies("2015")("gktimer_aproveview") = aproveview
    else
        if request.cookies("2015")("gktimer_aproveview") <> 0 then
        aproveview = request.cookies("2015")("gktimer_aproveview")
        else
        aproveview = 0
        end if
    end if
    
    aproveviewfilerSEL0 = ""
    aproveviewfilerSEL1 = ""
    select case aproveview 
        case 0
        aproveviewfilerSEL0 = "SELECTED"
        aproveviewfilerTXT = godkend_txt_023
        case 1 
        aproveviewfilerSEL1 = "SELECTED"
        aproveviewfilerTXT = godkend_txt_006
    end select


     


    jobidsStr = " AND (j.id <> 0 "
    jobid = 0
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
    
        response.cookies("2015")("gktimer_seljob") = jobid

    else
        if request.cookies("2015")("gktimer_seljob") <> "" then
        jobid = request.cookies("2015")("gktimer_seljob")
        else
        jobid = 0
        end if
    end if

    jobidsStr = jobidsStr & ")"

    select case func
    case "db"
        
        
        
        
        
        'response.Redirect "godkend_job_timer_2017.asp?aar="&aar&"&aarslut="&aarslut&"&FM_job="&jobid
        %>
        <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
        <%
        Response.write "<br><br><div style='position:relative; top:40px; left:40px;'><a href='godkend_job_timer_2017.asp?aar="&aar&"&aarslut="&aarslut&"&FM_job="&jobid &"'>The Work is done, continue >></a></div>"
        'Response.end

    case else
    
     
%>
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

 <script src="js/godkend_job7.js" type="text/javascript"></script>

<style>
    .aktivilist:hover,
    .aktivilist:focus {
    text-decoration: none;
    cursor: pointer;
    }

    .medarbliste:hover,
    .medarbliste:focus {
    text-decoration: none;
    cursor: pointer;
    }
</style>

<%if session("mid") = 1 AND lto = "tia" then
    session("mid") = 9 '1645
  end if %>

<%call menu_2014 %>


  
  <div class="wrapper">
      <div class="content">
         
          <div class="container">
              <div class="portlet">
                  <h3 class="portlet-title"><u><%=godkend_txt_001 %></u> </h3>
                  <div class="portlet-body">

                   <!--<div id="div_jobid">DDD</div>-->

                      <div class="well">
                        <form action="godkend_job_timer_2017.asp?sogsubmitted=1" method="POST">

                             
                        

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
                                
                                    strJobnavn = oRec3("jobnavn")

                                    if cdbl(jobid) = cdbl(oRec3("id")) then
				                    isSelected = "SELECTED"
				                    else
				                    isSelected = ""
				                    end if

                                    neutraltimer = 0
                                    strSQLneutraltimer = "SELECT sum(timer) as timer FROM timer WHERE tjobnr = '"& oRec3("jobnr") & "' AND (godkendtstatus = 0 OR godkendtstatus = 3) GROUP BY tjobnr"
                                    oRec4.open strSQLneutraltimer, oConn, 3
                                    if not oRec4.EOF then
                                      neutraltimer = oRec4("timer")
                                    end if
                                    oRec4.close
				    
                                    select case lto
                                    case "tia"

                                        if cdbl(neutraltimer) <> 0 then
                                        jobnavnid = oRec3("jobnr") &" - "& oRec3("jobnavn") &"&nbsp - &nbsp"& formatnumber(neutraltimer, 2) &" t."
                                        else
                                        jobnavnid = oRec3("jobnr") &" - "& oRec3("jobnavn")
                                        end if

                                    case else 

				                        if cdbl(neutraltimer) <> 0 then
                                        jobnavnid = oRec3("jobnavn") & " ("& oRec3("jobnr") &")" &"&nbsp - &nbsp"& formatnumber(neutraltimer, 2) &" t."
                                        else
                                        jobnavnid = oRec3("jobnavn") & " ("& oRec3("jobnr") &")"
                                        end if
                                    
                                      end select
        
				                %>
				                <option value="<%=oRec3("id")%>" <%=isSelected%>><%=jobnavnid%></option>
				                <%

                                oRec3.movenext
                                wend
                                oRec3.close 

                                %>

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
                                <br />
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

                <input class="dateweeknum" id="bruguge" name="bruguge" type="checkbox" value="1" <%=brugugeCHK %> />  Week:
                <input type="text" id="bruguge_selector_of" class="form-control input-small" value="<%=Datepart("ww",now) %>" readonly>
			<select name="bruguge_week" id="bruguge_selector" class="form-control input-small" onchange="submit();" style="display:none" >
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

                            <div class="row">
                                <div class="col-lg-6">
                                    Filter: <select name="filter" class="form-control input-small" onchange="submit();">
                                    <option value="0" <%=filterSEL0%>>Show All</option>
                                    <option value="1" <%=filterSEL1%>>Show only Approved</option>
                                    <option value="2" <%=filterSEL2%>>Show only Declined</option>
                                    <option value="3" <%=filterSEL3%>>Activities to be approved</option>
                                    <option value="4" <%=filterSEL4%>>Show only Tentative</option>
                                    </select>
                                </div>

                                <div class="col-lg-2">View:
                                    <select name="aproveviewfiler" class="form-control input-small" onchange="submit();">
                                    <option value="1" <%=aproveviewfilerSEL1 %>>Based by employee</option>
                                    <option value="0" <%=aproveviewfilerSEL0 %>>Based by activity</option>
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

                      <!--onClick="reloadpage();"-->
                      <!-- godkend_job_timer_2017.asp?func=db&aar=<%=aar %>&aarslut=<%=aarslut %>&FM_job=<%=jobid %> -->
                      <form action="godkend_job_timer_2017.asp?func=db&aar=<%=aar %>&aarslut=<%=aarslut %>&FM_job=<%=jobid%>" method="post" id="godkendform" name="godkendform" >
                      <input type="hidden" value="<%=Strjobid %>" id="godkendjobid" />
                      <table style="background-color:white;" class="table dataTable table-striped table-bordered table-hover ui-datatable">

                          <thead>
                              <tr>
                                    <th><%=aproveviewfilerTXT %></th>

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
                                    <th style="text-align:center;"><%=godkend_txt_007 %> &nbsp;<span style="font-size:9px;">(Appr.)</span></th>
                                  <th style="text-align:center;"><%=godkend_txt_009 %> <input type="checkbox" id="godkendalle" class="godkendalle" name="godkendalle" /></th>
                                  <th style="text-align:center;">Decline <input type="checkbox" id="afvisalle" name="godkendalle" class="afvisalle" /></th>
                                  <th style="text-align:center;" >App./Decl. Date</th>
                              </tr>

                          </thead>

                          <%if aproveview = 1 then %>

                          <tbody>
                            
                              <%
                                 
                               
   
                                  lastmid = 0
                                  m = 1000
                                  d = 10000
                                  lasttmnr = 0
                                  dim timer_md, dato_medid, medarbid, medidnavn, manedstot
                                  Redim timer_md(m,d), dato_medid(m,d), medarbid(m), medidnavn(m), manedstot(antalmaaned)
                                  
                                  strSQL = "SELECT tmnr, tmnavn, tdato, sum(timer) as Timer FROM timer WHERE ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' "& filterSQL &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' AND tmnr is not null GROUP by tmnr, year(tdato), month(tdato) Order by tmnr, tdato "
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
                                  dec_ids = "" 
                                  'response.Write "m" & m & "d" & d  
                                  for m = 0 TO m_end

                                  dec_ids = "" 
                                  
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
                                                <%   '** Timer pr. medarbejder i periode
                                                    timerperiode = 0
                                                    strSQLtimer = "SELECT sum(timer) as Timer FROM timer WHERE ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND tmnr ="& medarbid(m) & filterSQL & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'"
                                                    'response.Write strSQLtimer
                                                    'response.Flush
                                                    oRec.open strSQLtimer, oConn, 3
                                                    if not oRec.EOF then

                                                    timerperiode = oRec("Timer")

                                                    end if
                                                    oRec.close
                                                    
                                                %>
                                                <%=formatnumber(timerperiode, 2) %>


                                                 <%   '** Godkendte: Timer pr. medarbejder i periode

                                                    gktimerperiode = 0
                                                    strSQLtimer = "SELECT sum(timer) as Timer FROM timer WHERE ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND tmnr ="& medarbid(m) & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' AND godkendtstatus = 1"
                                                    'response.Write strSQLtimer
                                                    'response.Flush
                                                    oRec.open strSQLtimer, oConn, 3
                                                    if not oRec.EOF then

                                                    gktimerperiode = oRec("Timer")

                                                    end if
                                                    oRec.close
                                                    
                                                
                                                     'if session("mid") = 1 then

                                                     'Response.write "HER: "& gktimerperiode

                                                     'end if
                                    
                                                if gktimerperiode <> 0 then
                                                gktimerperiode = gktimerperiode 
                                                else 
                                                gktimerperiode = 0
                                                end if%>

                                                &nbsp;<span style="font-size:9px;">
                                                    <%
                                                        'select case filterThis
                                                        'case "1"
                                                        response.Write "("& formatnumber(gktimerperiode,2) &")"
                                                        'case "2"
                                                        
                                                        'case "3"
                                                        
                                                        'case else
                                                        'response.Write "("& formatnumber(gktimerperiode,2) &")"
                                                        'end select 
                                                    %>
                                                      </span>

                                                <%'end if %>

                                            </td>
                                           
                                            <td style="text-align:center">
                                            <input type="checkbox" class="godkendmedarb godkendmedarb_<%=medarbid(m)%>" id="medarbgodkend_<%=medarbid(m) %>" /></td> 
                                            <td style="text-align:center"><input type="checkbox" class="afvismedarb afvismedarb_<%=medarbid(m)%>" id="medarbafvis_<%=medarbid(m) %>" /></td> 
                                            <td>&nbsp;</td>
                                        </tr>
                                        <tbody style="display:none; font-size:95%" id="aktivilist_<%=medarbid(m) %>">
                                        <%
                                         
                                        

                                        strSQLakt = "SELECT tid, taktivitetid, taktivitetnavn, timer, tdato, godkendtdato, godkendtstatus, godkendtstatusaf, timerkom, overfort FROM timer WHERE ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND tmnr ="& medarbid(m) & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& filterSQL &" ORDER BY taktivitetnavn"
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
                                        overfort = oRec("overfort")


                                        if cDate(aktgodkendtdato) >= cDate("01-01-2002") then
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
                                                    strSQLakttotal = "SELECT taktivitetnavn, sum(timer) as Timer FROM timer WHERE TAktivitetId ="& aktid &" AND ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND tmnr ="& medarbid(m) & filterSQL & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'"
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
                                            case 0, 3 'Ikke taget stilling / Tentative
                                                aktgodkendtstatus1CHK = ""
                                                aktgodkendtstatus2CHK = ""
                                            case 1
                                                aktgodkendtstatus1CHK = "CHECKED"
                                                aktgodkendtstatus2CHK = ""
                                            case 2
                                                 aktgodkendtstatus1CHK = ""
                                                aktgodkendtstatus2CHK = "CHECKED"
                                                dec_ids = dec_ids & " OR tid = " & tid
                                            end select    


                                           'Kun timer fra medarbejer godkedte uger må godkendes

                                            strMrd_sm = datepart("m", oRec("tdato"), 2, 2)
                                            strAar_sm = datepart("yyyy", oRec("tdato"), 2, 2)
                                            strWeek = datepart("ww", oRec("tdato"), 2, 2)
                                            strAar = datepart("yyyy", oRec("tdato"), 2, 2)

                                            if cint(SmiWeekOrMonth) = 0 then
                                            usePeriod = strWeek
                                            useYear = strAar
                                            else
                                            usePeriod = strMrd_sm
                                            useYear = strAar_sm
                                            end if

                
                                            call erugeAfslutte(useYear, usePeriod, medarbid(m), SmiWeekOrMonth, 0)

                                            call lonKorsel_lukketPer(oRec("Tdato"), -2)

                                            tjkDato = oRec("Tdato")

                                            call tjkClosedPeriodCriteria(tjkDato, ugeNrAfsluttet, usePeriod, SmiWeekOrMonth, splithr, smilaktiv, autogk, autolukvdato, lonKorsel_lukketIO)    
                                        
                                            'if (cint(ugeerAfsl_og_autogk_smil) = 1 OR cint(overfort) = 1 OR cint(aktgodkendtstatus) = 0) AND cint(level) <> 1 then
                                            'if session("mid") = 1 OR session("mid") = 1505 then
                                            '    response.write "ugeerAfsl_og_autogk_smil: "& cint(ugeerAfsl_og_autogk_smil) & "O: "& overfort &" aktgodkendtstatus: " & cint(aktgodkendtstatus) & " ugeerAfsl_og_autogk_smil2: "& ugeerAfsl_og_autogk_smil & " AND ugegodkendt: " & ugegodkendt & " AND L:" & cint(level) & "<br>"
                                            'end if
                                             infoTxt = ""
                                             if ((cint(ugeAfsluttetAfMedarb) = 0 OR cint(overfort) = 1) OR (cint(ugeerAfsl_og_autogk_smil) = 1 AND cint(ugegodkendt) = 1 AND cint(aktgodkendtstatus) = 1)) then

                                                if cint(level) = 1 then
                                                chkDis = ""
                                                else
                                                chkDis = "disabled"
                                                end if

                                                if cint(ugeAfsluttetAfMedarb) = 0 then
                                                infoTxt = "Week not completed by employee"
                                                end if
                                            
                                            else
                                            chkDis = "" 
                                            end if %>


                                            
                                            <%if len(trim(infoTxt)) <> "" AND chkDis = "disabled" AND cint(aktgodkendtstatus) = 0 then %>
                                             <td style="text-align:center" colspan="2">  <span style="font-size:9px;"><%=infoTxt %> (<%=aktgodkendtstatus %>)</span></td>
                                            <%else %>
                                            
                                            <td style="text-align:center">                                               
                                                <input <%=chkDis %> type="radio" id="godkendstatus_<%=medarbid(m) %>_<%=aktid %>_<%=tid %>" class="godkendbox godkendbox_<%=medarbid(m)%>" name="godkendbox_<%=tid %>" value="<%=medarbid(m) %>_<%=aktid %>_<%=tid %>" <%=aktgodkendtstatus1CHK %> />
                                                  <%if len(trim(infoTxt)) <> "" AND cint(level) = 1 then %>
                                                <br /><span style="font-size:9px; color:#CCCCCC;"><%=left(infoTxt, 15) %></span>
                                                <%end if %>
                                            </td> 
                                            <td style="text-align:center">          
                                                <input <%=chkDis %> type="radio" id="afvisstatus_<%=medarbid(m) %>_<%=aktid %>_<%=tid %>" class="afvisbox afvisbox_<%=medarbid(m)%>" name="godkendbox_<%=tid %>" value="<%=medarbid(m) %>_<%=aktid %>_<%=tid %>" <%=aktgodkendtstatus2CHK %> />
                                                   <%if len(trim(infoTxt)) <> "" AND cint(level) = 1 then %>
                                                <br /><span style="font-size:9px; color:#CCCCCC;"><%=left(infoTxt, 15) %></span>
                                                <%end if %>

                                            </td>
                                               
                                            <%end if %>
                                            <!--
                                             <td style="text-align:center">                                               
                                                <input type="radio" id="godkendstatus_<%=medarbid(m) %>_<%=aktid %>_<%=tid %>" class="godkendbox godkendbox_<%=medarbid(m)%>" name="godkendbox_<%=tid %>" value="<%=tid %>" <%=aktgodkendtstatus1CHK %> /></td> 
                                            <td style="text-align:center">                                               
                                                <input type="radio" id="afvisstatus_<%=medarbid(m) %>_<%=aktid %>_<%=tid %>" class="afvisbox afvisbox_<%=medarbid(m)%>" name="godkendbox_<%=tid %>" value="<%=tid %>" <%=aktgodkendtstatus2CHK %> /></td>
                                       
                                                -->

                                            <td style="text-align:center">
                                                <%if cint(aktgodkendtstatus) = 1 OR cint(aktgodkendtstatus) = 2 then %>
                                                <span style="font-size:9px;"><%=aktgodkendtdato %><br />
                                                <%=left(aktgodkendtstatusaf, 10) %></span>
                                                <%end if %>
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
                                  <td style="text-align:right" colspan="3">Comment for declined<br /><textarea id="decline_comment_<%=medarbid(m) &"_"&lastAktid %>" name="decline_comment" class="form-control input-small"></textarea>
                                      <br />
                                        
                                         <%if len(trim(dec_ids)) <> 0 then %>
                                        <a id="sendemail_<%=medarbid(m)&"_"&lastAktid %>" class="sendemail">Send email with declined info >></a>
                                       <br><span style="color:#999999; visibility:hidden; display:;" class="noemail">You need to submit before sending email</span>
                                        <%else %>
                                        <span style="color:#999999;">You need to submit before sending email</span>
                                       <%end if %>
                                        <!--<div id="sendemail_send_<%=medarbid(m) %>">Mail send!</div>-->
                                        <input type="hidden" value="<%=dec_ids %>" id="decl_tids_<%=medarbid(m)&"_"&lastAktid %>" />
                                        <input type="hidden" value="<%=medarbid(m) %>" id="decl_tids_mid_<%=medarbid(m)&"_"&lastAktid %>" class="decl_tids" />
                                      </td>
                              </tr>
                             </tbody>
                                  <%
                                      dec_ids = ""

                                        'end if
                                  
                                     lastmid = medarbid(m)    
                                  next
       
                              %>
                              
                         

                         <!-- <tfoot>

                               

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
                                 
                                  <th style="text-align:center; width:20px;">
                                      <input type="submit" class="btn btn-success btn-sm" value="Submit" />
                                      <!--<a class="btn btn-success btn-sm" id="godkendknap">Submit</a><!-- =godkend_txt_009  -></th>
                                
                                  </tr>

                          </tfoot> -->

                          <%else 'akt - > medarb 
                          '*** GROUP BY activity%>

                          <tbody>
                             <%
                            strSQLhentakt = "SELECT navn, id FROM aktiviteter AS a LEFT JOIN timer as t ON (a.id = t.TAktivitetId) WHERE job ="& jobid & " AND t.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& filterSQL & " ORDER BY a.id"

                            oRec.open strSQLhentakt, oConn, 3
                            while not oRec.EOF 

                            aktnavn = oRec("navn")
                            aktid = oRec("id")

                            if lastaktid <> aktid then 

                            %>
                                
                           <tr id="aktvisfont_<%=aktid %>">
                               <td style="white-space:nowrap; width:350px;"><%=aktnavn %> <span class="fa fa-chevron-down pull-right medarbliste" id="<%=aktid %>"></span></td>
                               <td style="text-align:center">
                                   <%
                                       strTimer = 0
                                       strSQLhenttimer = "SELECT sum(timer) as timer FROM timer WHERE tjobnr = '"& Strjobid &"' AND TAktivitetId ="& aktid & filterSQL &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY tjobnr, TAktivitetId"
                                       oRec2.open strSQLhenttimer, oConn, 3
                                       if not oRec2.EOF then
                                    
                                       strTimer = oRec2("timer")
                                    

                                       'response.Write formatnumber(strTimer,2)
                                       end if
                                       oRec2.close 

                                       if strTimer <> 0 then
                                       strTimer = formatnumber(strTimer, 2)
                                       else
                                       strTimer = 0
                                       end if
                                   %>
                               </td>
                               <td><!-- comment --></td>
                               <td><!-- dates --></td>

                               <td style="text-align:center">
                                   <%
                                        gktimerperiode = 0
                                        strSQLtimer = "SELECT sum(timer) as Timer FROM timer WHERE ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND TAktivitetId ="& aktid & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' AND godkendtstatus = 1 GROUP BY tjobnr, TAktivitetId"
                                        'response.Write strSQLtimer
                                        'response.Flush
                                        oRec2.open strSQLtimer, oConn, 3
                                        if not oRec2.EOF then

                                        gktimerperiode = oRec2("Timer")
                                   

                                        end if
                                        oRec2.close

                                       if gktimerperiode <> 0 then
                                       gktimerperiode = gktimerperiode
                                       else
                                       gktimerperiode = 0
                                       end if

                                       'select case filterThis
                                       'case "1"
                                       response.Write strTimer &"&nbsp <span style=""font-size:9px;"">("& formatnumber(gktimerperiode,2) &")</span>"
                                       'case "2"
                                       'response.Write strTimer
                                       'case "3"
                                       'response.Write strTimer
                                       'case else
                                       'response.Write strTimer &"&nbsp <span style=""font-size:9px;"">("& formatnumber(gktimerperiode,2) &")</span>"
                                       'end select
                                       
                                   %>
                               </td>  

                               <td><!-- godkedn --></td>
                               <td><!-- avis --></td>
                               <td><!-- Apr/decl date --></td>
                           </tr>

                           <tbody style="display:none" id="tr_medarbliste_<%=aktid %>">
                                 <%
                                     m = 0
                                     strSQLMedarb = "SELECT Tmnavn, Tmnr, Tid, godkendtstatusaf, godkendtdato, Tdato, timer, godkendtstatus, timerkom, overfort FROM timer WHERE tjobnr = '"& Strjobid &"' AND TAktivitetId ="& aktid &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "&filterSQL&" ORDER BY Tmnavn"                                   
                                     oRec2.open strSQLMedarb, oConn, 3
                                     While not oRec2.EOF

                                     aktgodkendtstatus = oRec2("godkendtstatus")

                                    select case cint(oRec2("godkendtstatus"))
                                    case 0, 3 '3: Tentative
                                        aktgodkendtstatus1CHK = ""
                                        aktgodkendtstatus2CHK = ""
                                    case 1
                                        aktgodkendtstatus1CHK = "CHECKED"
                                        aktgodkendtstatus2CHK = ""
                                    case 2
                                        aktgodkendtstatus1CHK = ""
                                        aktgodkendtstatus2CHK = "CHECKED"
                                        dec_ids = dec_ids & " OR tid = " & oRec2("tid")
                                    end select    

                                    overfort = oRec2("overfort")

                                    %>

                                    <%if cdbl(lastTmnr) <> cdbl(oRec2("Tmnr")) AND m > 0 then 
                                    %>
                                    <tr>
                                        <td>&nbsp</td>
                                        <td>&nbsp</td>
                                        <td>&nbsp</td>
                                        <td>&nbsp</td>
                                        <td>&nbsp</td>
                                     
                                      <td style="text-align:right" colspan="3">Comment for declined<br /><textarea id="decline_comment_<%=lastTmnr&"_"&aktid %>" name="decline_comment" class="form-control input-small"></textarea>
                                      <br />
                                        <%if len(trim(dec_ids)) <> 0 then %>
                                        <a id="sendemail_<%=lastTmnr&"_"&aktid %>" class="sendemail">Send email with declined info >></a>
                                          <br><span style="color:#999999; visibility:hidden; display:;" class="noemail">You need to submit before sending email</span>
                                        <%else%>
                                        <span style="color:#999999;">You need to submit before sending email</span>
                                          <%end if %>
                                        <!--<div id="sendemail_send_<%=lastTmnr %>">Mail send!</div>-->
                                        <input type="hidden" value="<%=dec_ids %>" id="decl_tids_<%=lastTmnr&"_"&aktid %>" />
                                        <input type="hidden" value="<%=lastTmnr %>" id="decl_tids_mid_<%=lastTmnr&"_"&aktid %>" class="decl_tids" />
                                      </td>
                                    </tr>
                                    <%
                                    dec_ids = ""     
                                    end if %>

                                    <tr>                                      
                                        <td>&nbsp - &nbsp<%=oRec2("Tmnavn") %></td>
                                        <td style="text-align:center">
                                            <%
                                                response.Write formatnumber(oRec2("timer"),2)                                               
                                            %>
                                        </td>
                                        <td>
                                            <%if len(trim(oRec2("timerkom"))) <> 0 then %>
                                                <span style="font-size:9px; font-weight:lighter;"><%=left(oRec2("timerkom"), 250) %></span>
                                            <%end if %>
                                        </td>
                                        <td style="text-align:center"><%=oRec2("Tdato") %></td>

                                        <%
                                            if cdbl(lastTmnr) <> cdbl(oRec2("Tmnr")) then
                                                    akTimerTot = 0
                                                    strSQLakttotal = "SELECT taktivitetnavn, sum(timer) as Timer FROM timer WHERE TAktivitetId ="& aktid &" AND ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "' AND tmnr ="& oRec2("Tmnr") & filterSQL &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY TAktivitetId, tjobnr, tmnr"
                                                    oRec3.open strSQLakttotal, oConn, 3
                                                    if not oRec3.EOF then
                                                    akTimerTot = oRec3("Timer")
                                                    end if
                                                    oRec3.close 
                                        %>

                                            <td style="text-align:center">
                                                <%=formatnumber(akTimerTot,2)  %>
                                            </td> 

                                            <%else %>

                                               <td style="text-align:right">
                                              &nbsp;
                                            </td>

                                            <%end if %>


                                        <%'Kun timer fra medarbejer gdokedte uger må godkendes

                                        strMrd_sm = datepart("m", oRec2("tdato"), 2, 2)
                                        strAar_sm = datepart("yyyy", oRec2("tdato"), 2, 2)
                                        strWeek = datepart("ww", oRec2("tdato"), 2, 2)
                                        strAar = datepart("yyyy", oRec2("tdato"), 2, 2)

                                        if cint(SmiWeekOrMonth) = 0 then
                                        usePeriod = strWeek
                                        useYear = strAar
                                        else
                                        usePeriod = strMrd_sm
                                        useYear = strAar_sm
                                        end if

                
                                        call erugeAfslutte(useYear, usePeriod, oRec2("Tmnr"), SmiWeekOrMonth, 0)

                                        call lonKorsel_lukketPer(oRec2("Tdato"), -2)

                                        tjkDato = oRec2("Tdato")

                                        call tjkClosedPeriodCriteria(tjkDato, ugeNrAfsluttet, usePeriod, SmiWeekOrMonth, splithr, smilaktiv, autogk, autolukvdato, lonKorsel_lukketIO)    
                                        
                                        '** TIA
                                        '** Hvis cint(ugeerAfsl_og_autogk_smil) = 1 OR cint(overfort) = 1
                                        '** eller cint(aktgodkendtstatus) = 0
                                        '** Hvis uge er afsluttet KAN cint(aktgodkendtstatus) ikke være = 0. De er da Tentative. 
                                        infoTxt = ""
                                        'if (cint(ugeerAfsl_og_autogk_smil) = 1 OR cint(overfort) = 1 OR cint(aktgodkendtstatus) = 0) AND cint(level) <> 1 then
                                        if ((cint(ugeAfsluttetAfMedarb) = 0 OR cint(overfort) = 1) OR (cint(ugeerAfsl_og_autogk_smil) = 1 AND cint(ugegodkendt) = 1 AND cint(aktgodkendtstatus) = 1)) then
                                                
                                                 if cint(level) = 1 then
                                                chkDis = ""
                                                else
                                                chkDis = "disabled"
                                                end if

                                                  if cint(ugeAfsluttetAfMedarb) = 0 then
                                                infoTxt = "Week not completed by employee"
                                                end if
                                            
                                            else
                                            chkDis = "" 
                                            end if %>


                                            
                                            <%if len(trim(infoTxt)) <> "" AND chkDis = "disabled" AND cint(aktgodkendtstatus) = 0 then %>
                                             <td style="text-align:center" colspan="2">  <span style="font-size:9px;"><%=infoTxt %> (<%=aktgodkendtstatus %>)</span></td>
                                            <%else %>


                                        <td style="text-align:center">     
                                            
                                                                               
                                                <input type="radio" <%=chkDis %> id="godkendstatus_<%=oRec2("Tmnr") %>_<%=aktid %>_<%=oRec2("Tid") %>" class="godkendbox godkendbox_<%=oRec2("Tmnr")%>" name="godkendbox_<%=oRec2("Tid") %>" value="<%=oRec2("Tmnr") %>_<%=aktid %>_<%=oRec2("Tmnr") %>" <%=aktgodkendtstatus1CHK %> />
                                                <%if len(trim(infoTxt)) <> "" AND cint(level) = 1 then %>
                                                <br /><span style="font-size:9px; color:#CCCCCC;"><%=left(infoTxt, 15) %></span>
                                                <%end if %>
                                        </td>
                                        <td style="text-align:center">
                                            <input type="radio" <%=chkDis %> id="afvisstatus_<%=oRec2("Tmnr") %>_<%=aktid %>_<%=oRec2("Tid") %>" class="afvisbox afvisbox_<%=oRec2("Tmnr")%>" name="godkendbox_<%=oRec2("Tid") %>" value="<%=oRec2("Tmnr") %>_<%=aktid %>_<%=oRec2("Tid") %>" <%=aktgodkendtstatus2CHK %> />
                                               <%if len(trim(infoTxt)) <> "" AND cint(level) = 1 then %>
                                                <br /><span style="font-size:9px; color:#CCCCCC;"><%=left(infoTxt, 15) %></span>
                                                <%end if %>
                                        </td>
                                        <%end if %>

                                        <td style="text-align:center">
                                            <% if cint(oRec2("godkendtstatus")) = 1 OR cint(oRec2("godkendtstatus")) = 2 then
                                                aktgodkendtdato = oRec2("godkendtdato")
                                                if cDate(aktgodkendtdato) >= cDate("01-01-2002") then
                                                aktgodkendtdato = aktgodkendtdato
                                                else 
                                                aktgodkendtdato = ""
                                                end if 
                                            %>
                                            <span style="font-size:9px;"><%=aktgodkendtdato %><br />
                                            <%=left(oRec2("godkendtstatusaf"), 10) %></span>
                                            <%end if %>
                                        </td>

                                    </tr>

                                       

                                    <%
                                    
                                     lastAktid = aktid
                                     lastTmnr = oRec2("Tmnr")  
                                     m = m + 1      
                                     oRec2.movenext
                                     wend
                                     oRec2.close 
                                    %>
                                    <tr>
                                        <td>&nbsp</td>
                                        <td>&nbsp</td>
                                        <td>&nbsp</td>
                                        <td>&nbsp</td>
                                        <td>&nbsp</td>
                                     
                                      <td style="text-align:right" colspan="3">Comment for declined<br /><textarea id="decline_comment_<%=lastTmnr&"_"&lastAktid %>" name="decline_comment" class="form-control input-small"></textarea>
                                      <br />
                                        <%if len(trim(dec_ids)) <> 0 then %>
                                        <a id="sendemail_<%=lastTmnr&"_"&lastAktid %>" class="sendemail">Send email with declined info >></a>
                                          <br><span style="color:#999999; visibility:hidden; display:;" class="noemail">You need to submit before sending email</span>
                                        <%else %>
                                        <span style="color:#999999;">You need to submit before sending email</span>
                                          <%end if %>
                                          <!--<div id="sendemail_send_<%=lastTmnr %>">Mail send!</div>-->
                                       <input type="hidden" value="<%=dec_ids %>" id="decl_tids_<%=lastTmnr&"_"&lastAktid %>" />
                                        <input type="hidden" value="<%=lastTmnr %>" id="decl_tids_mid_<%=lastTmnr&"_"&lastAktid %>" class="decl_tids" />
                                      </td>
                                    </tr>                                                            
                           </tbody>


                           <%

                            dec_ids = "" 
                            end if
                            lastaktid = aktid
                            lastTmnr = 0                            
                            oRec.movenext
                            wend
                            oRec.close  
                            %>

                          <!--</tbody>-->

                          <%end if %>

                          <tfoot>
                              <tr>
                                  <th>Total</th>

                                  <th style="text-align:center">
                                      <%
                                          totaltimer = 0
                                          strTotalTimer = "SELECT sum(timer) as timer FROM timer WHERE ("& aty_sql_realhours &") AND tjobnr = '"& Strjobid & "'" & filterSQL &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY tjobnr"
                                          oRec.open strTotalTimer, oConn, 3
                                          if not oRec.EOF then
                                          totaltimer = oRec("timer")
                                          end if
                                          oRec.close 

                                          if totaltimer <> 0 then
                                          totaltimer = totaltimer
                                          else
                                          totaltimer = 0
                                          end if

                                          response.Write formatnumber(totaltimer,2)
                                       %>
                                  </th>

                                  <th>&nbsp</th>
                                  <th>&nbsp</th>
                                  <th>&nbsp</th>
                                  <th>&nbsp</th>
                                  <th>&nbsp</th>
                                  <th style="text-align:center; width:20px;">
                                      <!--<input type="submit" class="btn btn-success btn-sm"  value="Submit" /> //-->
                                      <a class="btn btn-success btn-sm" id="godkendknap">Submit</a><!-- =godkend_txt_009  --></th>
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


<!--#include file="../inc/regular/footer_inc.asp"-->
<%end select %>




