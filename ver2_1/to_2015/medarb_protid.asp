 

<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->

<%
    if Request.Form("AjaxUpdateField2") = "true" then
        Select Case Request.Form("control")
            case "SogJobFilter"

                sogval = request("FM_sogval")

                select case lto
                    case "cflow"
                        sqlorderby = "kid, kkundenavn, jobnr"
                    case else
                        sqlorderby = "kid, kkundenavn, jobnavn"
                end select

                strOpt = ""

                if sogval <> "" then

                    level = session("rettigheder")
                    
                    call GetMedarbJobAns(session("mid"))

                    strPrgids = "ingen"
                    select case strmedarbjobans
                        case "alle"
                            strJobKri = " AND id <> 0"
                        case "ingen"
                            response.Write "<option value='-1'>"& medarb_protid_txt_037 &"</option>"
                            response.End
                        case else
                            strJobKri = " AND id IN ("& strmedarbjobans &")"
                    end select

                    
                    strSqL = "SELECT id as jobid, jobnr, jobnavn, kkundenavn, kid FROM job LEFT JOIN kunder ON (jobknr = kid) WHERE jobstatus = 1 AND (jobnavn LIKE '%"& sogval &"%' OR jobnr LIKE '%"& sogval &"%')"& strJobKri &" ORDER BY "& sqlorderby
                    lastkid = -1
                    'response.Write strSqL
                    'response.End
                    oRec.open strSqL, oConn, 3
                    while not oRec.EOF

                        if oRec("kid") <> lastkid then
                            'response.Write "<option disabled></option>"
                            'response.Write "<option disabled>"& oRec("kkundenavn") &"</option>"          
                            strOpt = strOpt & "<option disabled></option>"
                            strOpt = strOpt & "<option disabled>"& oRec("kkundenavn") &"</option>"    
                        end if

                        jobSEL = ""
                        if checkSelectedJobs = 1 then
                            for j = 0 TO uBOUND(jobnr)
                                if oRec("jobnr") = jobnr(j) then
                                    jobSEL = "SELECTED"
                                end if
                            next
                        else
                            jobSEL = ""
                        end if

                        select case lto
                        case "cflow"
                            'response.Write "<option value='"&oRec("jobnr")&"' "& jobSEL &" data-jobnavn='"& oRec("jobnavn") &"' >"& oRec("jobnr") & " " & oRec("jobnavn") &"</option>"
                            strOpt = strOpt & "<option value='"&oRec("jobnr")&"' "& jobSEL &" data-jobnavn='"& oRec("jobnavn") &"' >"& oRec("jobnr") & " " & oRec("jobnavn") &"</option>"
                        case else
                            'response.Write "<option value='"&oRec("jobnr")&"' "& jobSEL &" data-jobnavn='"& oRec("jobnavn") &"' >"& oRec("jobnavn") & " ("&oRec("jobnr")&")" &"</option>"
                            strOpt = strOpt & "<option value='"&oRec("jobnr")&"' "& jobSEL &" data-jobnavn='"& oRec("jobnavn") &"' >"& oRec("jobnavn") & " ("&oRec("jobnr")&")" &"</option>"
                        end select
                                        
                        lastkid = oRec("kid")
                    oRec.movenext
                    wend
                    oRec.close
                end if 'sogval

                response.Write strOpt

        end select
    response.End
    end if
%>

<!--#include file="../inc/errors/error_inc.asp"-->
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

    if len(trim(request("version"))) <> 0 then
        version = request("version") '0 = Projekt, akt, medarb, eller Projekt, medarb, akt.... 1 = Medarb, projekt, akt 
    else
        version = 0
    end if

    media = request("media")

    if media <> "print" AND media <> "print_sim" then
        if len(trim(request("aar"))) <> 0 then
        aar = request("aar")
        else
        aar = "1-1-"& year(now)
        end if

        if len(trim(request("aarslut"))) <> 0 then
        aarslut = request("aarslut")
        else
        aarslut = day(now) &"-"& month(now) &"-"& year(now) '"1-1-"& year(now)
        end if
    else
        aar = request("aar")
        aarslut = request("aarslut")
    end if

    antalmaaned = (DateDiff("m",aar,aarslut))
    antalaar = (DateDiff("yyyy",aar,aarslut))

    'response.Write "<br><br><br>><br>><br>><br>><br> ---------------------------------------------------------------------------- JOBVALGT " & request("FM_job")

    if lto <> "cflow" then
        if antalmaaned > 11 then
            errortype = 226
	        call showError(errortype)
            response.end
        end if
    end if

    if len(trim(request("FM_job"))) <> 0 then
    jost0CHK = ""
    else
    viskunabnejob0 = 1
    jost0CHK = "CHECKED"
    end if

    'if media <> "print" then

    call GetMedarbJobAns(session("mid")) 

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

    'response.write "<br><br><br>><br>><br>><br>><br> ------------------------------------------------------------------------------------------ jobvalgt " & jobidsStr 

    strSQLJobWhere = ""
    checkSelectedJobs = 0
    if len(trim(request("FM_job"))) <> 0 then
        if request("FM_job") <> "-1" AND request("FM_job") <> -1 then
            checkSelectedJobs = 1
            jobnr = split(request("FM_job"), ", ") 
            for j = 0 to UBOUND(jobnr)
                if j = 0 then
                    strSQLJobWhere = " AND (jobnr = '"& jobnr(j) &"'"
                else
                    strSQLJobWhere = strSQLJobWhere  & " OR jobnr =  '" & jobnr(j) & "'"
                end if           
            next  
            strSQLJobWhere = strSQLJobWhere & ")"
        else
            if strmedarbjobans = "alle" then
                strSQLJobWhere = ""
            else
                if strmedarbjobans <> "ingen" then
                    strSQLJobWhere = " AND id IN ("& strmedarbjobans &")"
                end if
            end if
        end if
    else
        if strmedarbjobans = "alle" then
                strSQLJobWhere = ""
            else
                if strmedarbjobans <> "ingen" then
                    strSQLJobWhere = " AND id IN ("& strmedarbjobans &")"
                end if
            end if
    end if

    if strmedarbjobans = "ingen" then
        strSQLJobWhere = " AND (jobnr = 0)"
    end if

    jobidsStr = jobidsStr & ")"


    'else
    'jobid = request("jobid")
    'end if


    if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
    progrp = 0
	end if

    strSQLMedWhere = ""
    checkValgteMedarb = 0
    if len(trim(request("FM_medarb"))) <> 0 then
        if request("FM_medarb") <> "-1" AND request("FM_medarb") <> -1 then
            checkValgteMedarb = 1
            medid = split(request("FM_medarb"), ", ")
            intMids = split(request("FM_medarb"), ", ")
            for m = 0 TO UBOUND(medid)
                if m = 0 then
                    strSQLMedWhere = " AND tmnr IN ("& medid(m)
                else    
                    strSQLMedWhere = strSQLMedWhere & ", " & medid(m)
                end if
            next
            strSQLMedWhere = strSQLMedWhere & ")"
        else
            strSQLMedWhere = " AND tmnr IN ("& session("mid") & ")"
            intMids = split(request("FM_medarb"), ", ")
        end if
    else
        strSQLMedWhere = " AND tmnr IN ("& session("mid") & ")"
        intMids = split(session("mid"), ", ")
    end if


        call akttyper2009(2)

    strSQLtfaktimWhere = " AND ("& aty_sql_realhours &")" '" AND tfaktim <> 5"


%>



<%
    if media <> "print" AND media <> "print_sim" then
    call menu_2014
    oskrift = ""
    else
    
    'strSQLJobnavn = "SELECT jobnavn, jobnr from job WHERE id = "& jobid
    'oRec.open strSQLJobnavn, oConn, 3
    'if not oRec.EOF then
    'jobnavnPrint = oRec("jobnavn")
    'jobnavnNr = oRec("jobnr")
    'end if
    'oRec.close

        'if jobid <> 0 then
            'oskrift = "("& jobnavnPrint &" "& jobnavnNr &")"
        'else
            'oskrift = ""
        'end if

    end if

    FM_sortering = request("FM_sortering")
    CHBsort1 = ""
    CHBsort0 = ""
    if FM_sortering = 1 OR FM_sortering = "1" then
        CHBsort1 = "SELECTED"
    else
        CHBsort0 = "SELECTED"
    end if
    
    listsort = request("FM_listsort")

    listsortSEL0 = ""
    listsortSEL1 = ""
    select case cint(listsort)
        case 0
            listSorteringSQL = "ORDER BY jobnavn"
            listsortSEL0 = "SELECTED"
        case 1
            listSorteringSQL = "ORDER BY jobnr"
            listsortSEL1 = "SELECTED"
    end select

    startmonth = Month(aar)
    startyear = Year(aar)
    slutmonth = Month(aarslut)
    slutyear = Year(aarslut)

    startdato = startyear & "-" & (startmonth) & "-1"
    'slutdato = slutyear & "/" & (slutmonth + 1) & "/1"

    slutdato = dateAdd("m", 1, "1-"& slutmonth &"-"& slutyear) 
    slutdato = dateAdd("d", -1, slutdato)
    slutdato = year(slutdato) & "-" & month(slutdato) & "-" & day(slutdato)

    if media <> "print" then
        displayList = "none"
    else
        displayList = "normal"
    end if

%>


<%
    public strOskrifter
    function TableHeader %>
       

<thead>
    <tr>
    <%if version = 0 then %>
    <th style="color:black;"><%=medarb_protid_txt_023 %></th>
    <%strOskrifter = medarb_protid_txt_023 & ";" %>
    <%else %>
    <th style="color:black;"><%=medarb_protid_txt_024 %></th>
    <%strOskrifter = medarb_protid_txt_024 & ";" %>
    <%end if %>
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
                                           
                select case (months -1)
                case 1
                getdate = medarb_protid_txt_011
                case 2
                getdate = medarb_protid_txt_012
                case 3
                getdate = medarb_protid_txt_013
                case 4
                getdate = medarb_protid_txt_014
                case 5
                getdate = medarb_protid_txt_015
                case 6
                getdate = medarb_protid_txt_016
                case 7
                getdate = medarb_protid_txt_017
                case 8
                getdate = medarb_protid_txt_018
                case 9
                getdate = medarb_protid_txt_019
                case 10
                getdate = medarb_protid_txt_020
                case 11
                getdate = medarb_protid_txt_021
                case 12
                getdate = medarb_protid_txt_022

                end select
                                  
                expdate = getdate & " - " & Stryear & ";"
                getdate = getdate & "<br>" & Stryear 
                
                'getdate = Monthname(months -1 ,True) & "<br>" & Stryear
                'getdate = (Ucase(getdate))   
                                      
                                                                          
            %> <!--<th style="text-align:center;"><%=UCase(Left(getdate,1)) & Mid(getdate,2) %></th> -->
                <th style="text-align:center; color:black;"><%=getdate %></th>
        <%
                strOskrifter = strOskrifter & expdate

            next
        %>
        <%if version = 0 then %>
        <th style="color:black;"><%=medarb_protid_txt_025 %></th>
        <%else %>
        <th style="color:black;"><%=medarb_protid_txt_026 %></th>
        <%end if %>
    </tr>
</thead>

<%end function %>

<%
    if version = 1 then
        oskrift1 = medarb_protid_txt_024 &" - " & medarb_protid_txt_031
    else
        oskrift1 = medarb_protid_txt_010 &" - "& medarb_protid_txt_032
    end if
%>

  <%if media <> "print" AND media <> "print_sim" then %>
 <!-- <div class="wrapper">
      <div class="content"> -->
  <%end if %>

  <style>
      .zoom {
          zoom: 80%;
      }
  </style>    


          <div class="luft"><br /><br /><br /><br /><br /><br /></div>

          <script src="js/traveldietexp_jav.js" type="text/javascript"></script>
          <script src="js/medarb_protid_jav7.js" type="text/javascript"></script>
        <script src="../timereg/inc/header_lysblaa_stat_20170808.js" type="text/javascript"></script>

          <input type="hidden" id="FM_sortering" value="<%=FM_sortering %>" />
          <input type="hidden" id="version" value="<%=version %>" />

          <input type="hidden" id="alletxt" value="<%=medarb_protid_txt_027 %>" />

          <%
              if media = "export" then
                conDis = "none"
              else
                conDis = "normal"
              end if
          %>

          <div class="container" style="display:<%=conDis%>;">
              <div class="portlet">
                  <h3 class="portlet-title"><u><%=oskrift1 &" "& oskrift %> <%=level %></u></h3>
                  <div class="portlet-body">

                      <%if lto = "foa" then
                        if version = 0 AND (level > 2 AND level <> 6) then
                            response.Write "Du har ikke adgang til denne side <a href='medarb_protid.asp?version=1'>Klik for at g� tilbage</a>"
                            response.End
                        end if
                      end if %>

                      <%if media <> "print" AND media <> "print_sim" then %>
                      <div class="well printhide">
                        <form action="medarb_protid.asp?sogsubmitted=1&version=<%=version %>" method="POST">
                        
                       <!-- <div class="row">
                            <div class="col-lg-2">
                                <h4 class="panel-title-well"><%=dsb_txt_002 %></h4>
                            </div>
                            <div class="col-lg-10">&nbsp</div>                 
                        </div> -->
                
                        <div class="row">
                          
                            <%
                            if version = 0 then
                                colSize = "5"
                            else
                                colSize = "6"
                            end if
                            %>

                            <%if version = 0 then %>
                            <div class="col-lg-<%=colSize %>">
                                <%if version = 0 then %>
                                <%=medarb_protid_txt_010 %>: 
                                <%else %>
                                <%=medarb_protid_txt_024 %>:
                                <%end if %>
                                <br />

                               <input type="hidden" value="<%=strmedarbjobans %>" id="strmedarbjobans" />

                               <input type="text" class="form-control input-small" id="jogsog" autocomplete="off" placeholder="<%=dsb_txt_002 %>" />
                               <select multiple size="10" name="FM_jobxxx" id="FM_job" class="form-control input-small">
                                    <%if strmedarbjobans <> "ingen" then %>
                                    <option value="-1" data-jobnavn="<%=medarb_protid_txt_027 %>" ><%=medarb_protid_txt_027 %></option>
                                    <%else %>
                                    <option disabled><%=medarb_protid_txt_037 %></option>
                                    <%end if %>
                                    <%
                                        
                                        'select case lto
                                        'case "cflow"
                                        'sqlorderby = "kid, kkundenavn, jobnr"
                                        'case else
                                        'sqlorderby = "kid, kkundenavn, jobnavn"
                                        'end select


                                        'strSqL = "SELECT id as jobid, jobnr, jobnavn, kkundenavn, kid FROM job LEFT JOIN kunder ON (jobknr = kid) WHERE jobstatus = 1 ORDER BY "& sqlorderby
                                        'lastkid = -1
                                        'oRec.open strSqL, oConn, 3
                                        'while not oRec.EOF

                                            'if oRec("kid") <> lastkid then
                                                'response.Write "<option disabled></option>"
                                                'response.Write "<option disabled>"& oRec("kkundenavn") &"</option>"                                                
                                            'end if

                                            'jobSEL = ""
                                            'if checkSelectedJobs = 1 then
                                                'for j = 0 TO uBOUND(jobnr)
                                                    'if oRec("jobnr") = jobnr(j) then
                                                        'jobSEL = "SELECTED"
                                                    'end if
                                                'next
                                            'else
                                                'jobSEL = ""
                                            'end if

                                            'select case lto
                                            'case "cflow"
                                            'response.Write "<option value='"&oRec("jobnr")&"' "& jobSEL &" data-jobnavn='"& oRec("jobnavn") &"' >"& oRec("jobnr") & " " & oRec("jobnavn") &"</option>"
                                            'case else
                                            'response.Write "<option value='"&oRec("jobnr")&"' "& jobSEL &" data-jobnavn='"& oRec("jobnavn") &"' >"& oRec("jobnavn") & " ("&oRec("jobnr")&")" &"</option>"
                                            'end select
                                        
                                            'lastkid = oRec("kid")
                                        'oRec.movenext
                                        'wend
                                        'oRec.close
                                   


                                    %>
                               </select>
                                </div>
                                <%end if %>


                                <%if version = 1 then %>

                                <%call progrpmedarb_2018 %>

                               <!-- <select multiple size="10" class="form-control input-small" name="FM_medarb">
                                    <option value="-1"><%=medarb_protid_txt_027 %></option>
                                    <%
                                    strSQL = "SELECT mnavn, init, mid FROM medarbejdere WHERE mansat = 1 ORDER BY mnavn"
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF

                                        mSEL = ""
                                        if checkValgteMedarb = 1 then
                                            for m = 0 TO UBOUND(medid)
                                                if cdbl(oRec("mid")) = cdbl(medid(m)) then
                                                    mSEL = "SELECTED"
                                                end if
                                            next
                                        end if

                                        response.Write "<option value='"& oRec("mid") &"' "& mSEL &" >"& oRec("mnavn") &" ("& oRec("init") &")</option>"
                                    oRec.movenext
                                    wend
                                    oRec.close
                                    %>
                                </select> -->

                                <%end if %>
                            

                            <%
                                
                                'response.Write aar & " - " & aarslut 

                                'antalmaaned = (DateDiff("m",aar,aarslut))
                                'antalaar = (DateDiff("yyyy",aar,aarslut))
                                'response.Write(DateDiff("m",aar,aarslut))
                                'response.Write antalaar
                                                                                                                   
                            %>

                            
                                
                                
                           
                            <div class="col-lg-2"><%=medarb_protid_txt_001 %><br />

                                <%
                                if len(day(aar)) = 1 then
                                    aarStrDay = "0" & day(aar)
                                else
                                    aarStrDay = day(aar)
                                end if

                                if len(month(aar)) = 1 then
                                    aarStrMonth = "0" & month(aar)
                                else
                                    aarStrMonth = month(aar)
                                end if  

                                aarStr = aarStrDay &"-"& aarStrMonth &"-"& year(aar)
                                %>

                                <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="aar" value="<%=aarStr %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>

                                <%if version = 0 then %>
                                <br />
                                <%=medarb_protid_txt_030 %>: <br />
                                <select name="FM_listsort" class="form-control input-small">
                                    <option value="0" <%=listsortSEL0 %>><%=medarb_protid_txt_034 %></option>
                                    <option value="1" <%=listsortSEL1 %>><%=medarb_protid_txt_035 %></option>
                                </select>

                                <br />
                                <%=medarb_protid_txt_036 %>:
                                    <%if strmedarbjobans <> "ingen" then %>
                                        <div id="jobbox" style="width:400%;">
                                            <%
                                            if len(trim(request("FM_job"))) <> 0 AND request("FM_job") <> "-1" then
                                                for j = 0 TO UBOUND(jobnr)

                                                    jobnavn = ""
                                                    strSQLJobNavn = "SELECT jobnavn, jobnr FROM job WHERE jobnr = '"& jobnr(j) & "'"
                                                    oRec.open strSQLJobNavn, oConn, 3
                                                    if not oRec.EOF then
                                                        jobnavn = oRec("jobnavn") & " ("& oRec("jobnr") &")"
                                                    end if
                                                    oRec.close
                                                    

                                                    if jobnr(j) <> "0" then
                                                        divBox = "<div class='jobboxElement' id='jobbox_"& jobnr(j) &"' style='height:25px; display:inline-block; margin:2px; margin-right:7px;'>"& jobnavn &" &nbsp <span style='cursor:pointer;' class='fa fa-times fjernjob' id='" & jobnr(j) & "' ></span> &nbsp"
                                                        divBox = divBox & "<input type='hidden' value='" & jobnr(j) & "' name='FM_job' class='FM_job' id='" & jobnr(j) & "' />"
                                                        divBox = divBox & "</div>"
                                                        response.Write divBox
                                                    end if
                                                next
                                            end if

                                            if request("FM_job") = "-1" OR len(trim(request("FM_job"))) = 0 then
                                                divBox = "<div class='jobboxElement' id='jobbox_-1' style='height:25px; display:inline-block; margin:2px;'>"& medarb_protid_txt_027 &" &nbsp <span style='cursor:pointer;' class='fa fa-times fjernjob' id='-1' ></span> &nbsp"
                                                divBox = divBox & "<input type='hidden' value='-1' name='FM_job' class='FM_job' id='-1' />"
                                                divBox = divBox & "</div>"
                                                response.Write divBox
                                            end if
                                            %>
                                        </div>
                                    <%end if %>

                                <%end if %>
                            </div>

                            
                            <div class="col-lg-2"><%=medarb_protid_txt_002 %>:<br />

                                <%
                                if len(day(aarslut)) = 1 then
                                    aarStrDay = "0" & day(aarslut)
                                else
                                    aarStrDay = day(aarslut)
                                end if

                                if len(month(aarslut)) = 1 then
                                    aarStrMonth = "0" & month(aarslut)
                                else
                                    aarStrMonth = month(aarslut)
                                end if  

                                aarslStr = aarStrDay &"-"& aarStrMonth &"-"& year(aarslut)
                                %>

                                <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="aarslut" value="<%=aarslStr %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>

                            <%if version = 0 then %>
                            <div class="col-lg-2"><%=medarb_protid_txt_033 %>:<br />
                                <select name="FM_sortering" class="form-control input-small">
                                    <option value="0" <%=CHBsort0 %> ><%=medarb_protid_txt_028 %></option>
                                    <option value="1" <%=CHBsort1 %> ><%=medarb_protid_txt_024 %></option>
                                </select>
                            </div>

                            <%end if %>

                            <div class="col-lg-1"><br /><button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=medarb_protid_txt_003 %> >></b></button></div>  
                        </div>
                                             
                        </form>
                      </div>
                      <%end if 'media print %>

                      <!-- 2.0 -->
                      <%if cint(FM_sortering) = 0 AND version = 0 then %>
                      <table class="table dataTable table-bordered table-hover">

                          <%call TableHeader %>
                          <%if request("sogsubmitted") = 1 then %>
                          <tbody>
                              <%
                                jobsTotal = 0
                                dim periodTotal

                                if lto <> "cflow" then
                                    redim periodTotal(12) 'antal m�nneder
                                else
                                  redim periodTotal(72)
                                end if

                                'jobstatus = 1 AND
                                strSQL = "SELECT jobnavn, id as jobid, jobnr, kkundenavn FROM job LEFT JOIN timer ON (tjobnr = jobnr) LEFT JOIN kunder ON (jobknr = kid) WHERE tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'" & strSQLJobWhere & strSQLtfaktimWhere & " GROUP BY jobnr " & listSorteringSQL
                                'response.Write strSQL
                                oRec.open strSQL, oConn, 3
                                while not oRec.EOF

                                dim timerMonth
                                redim timerMonth(12, 2) 'antal m�neder, �r-nummer
                                jobtot = 0
                                lastyear = -1
                                strSQLTimer = "SELECT sum(timer) as timer, month(tdato) as month, year(tdato) as year FROM timer WHERE tjobnr = '"& oRec("jobnr") &"' AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP by year(tdato), month(tdato) Order by tdato"
                                'response.Write " HERHER " & strSQLTimer
                                oRec2.open strSQLTimer, oConn, 3
                                monthNumber = 0
                                while not oRec2.EOF                                 
                                    if oRec2("year") <> year(startdato) then
                                    yearNumber = 1
                                    else
                                    yearNumber = 0
                                    end if

                                    timerMonth(oRec2("month"), yearNumber) = oRec2("timer")
                                    jobtot = jobtot + oRec2("timer")

                                    monthNumber = monthNumber + 1
                                oRec2.movenext
                                wend
                                oRec2.close

                                if listsort = 0 then
                                    tdOskrift = oRec("jobnavn") &" ("& oRec("jobnr") &")"
                                else
                                    tdOskrift = oRec("jobnr") &" ("& oRec("jobnavn") &")"
                                end if
                              %>

                              <tr style="background-color:#f9f9f9;">
                                  <td><b><span class="expandJob" style="cursor:pointer;" id="<%=oRec("jobid") %>"><%=tdOskrift %> <span style="font-size:12px;"> - <%=oRec("kkundenavn") %></span> <%if media <> "print" AND media <> "print_sim" then %> <span id="icon_<%=oRec("jobid") %>" class="fa fa-angle-down"></span><%end if %> </span></b></td>
                                  
                                  <%
                                      ekspTxt = ekspTxt & oRec("kkundenavn") &"-"& oRec("kkundenavn") &";"
                                  %>

                                  <%
                                  for i = 0 to antalmaaned 
                                      if i = 0 then
                                        thisMonth = month(startdato)
                                        else
                                            thisMonth = thisMonth + 1
                                            if thisMonth > 12 then
                                                thisMonth = 1
                                                thisYear = 1
                                            end if
                                        end if

                                  %>
                                  <%periodTotal(i) = periodTotal(i) + timerMonth(thisMonth, thisYear) %>
                                  <td style="text-align:right;"><b><%=formatnumber(timerMonth(thisMonth, thisYear), 2) %></b></td>

                                  <%ekspTxt = ekspTxt & formatnumber(timerMonth(thisMonth, thisYear), 2) & ";" %>

                                  <%next %>

                                  <td style="text-align:right;"><b><%=formatnumber(jobtot, 2) %></b></td>

                                  <%ekspTxt = ekspTxt & formatnumber(jobtot, 2) & ";" & chr(013) %>

                              </tr>

                              <%
                                  strSQLAkt = "SELECT timer, taktivitetid, a.navn as aktnavn FROM timer LEFT JOIN aktiviteter a ON (taktivitetid = a.id) WHERE tjobnr = '"& oRec("jobnr") & "' AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP BY taktivitetid ORDER BY a.navn" 
                                  'response.Write strSQLAkt
                                  oRec2.open strSQLAkt, oConn, 3
                                  while not oRec2.EOF
                                  aktTot = 0
                              %>
                                    <tr style="display:<%=displayList%>;" class="akt_<%=oRec("jobid") %>">
                                        <td> &nbsp&nbsp - <b><span class="expandAkt" style="cursor:pointer;" id="<%=oRec("jobid") %>_<%=oRec2("taktivitetid") %>">&nbsp <%=oRec2("aktnavn") %> <%if media <> "print" then %> <span id="icon_<%=oRec("jobid") %>_<%=oRec2("taktivitetid") %>" class="fa fa-angle-down"></span><%end if %> </span></b></td>

                                        <%ekspTxt = ekspTxt & " - " & oRec2("aktnavn") & ";" %>

                                        <%
                                        Dim timerMonthAkt
                                        redim timerMonthAkt(12, 2) 'antal m�neder, �r-nummer
                                        
                                        strSQLAktTimer = "SELECT sum(timer) as timer, year(tdato) as year, month(tdato) as month FROM timer WHERE taktivitetid = "& oRec2("taktivitetid") &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP by year(tdato), month(tdato) Order by tdato"
                                        'response.Write " akt timer " & strSQLAktTimer
                                        monthNumber = 0
                                        oRec3.open strSQLAktTimer, oConn, 3
                                        while not oRec3.EOF
                                            if oRec3("year") <> year(startdato) then
                                            yearNumber = 1
                                            else
                                            yearNumber = 0
                                            end if

                                            timerMonthAkt(oRec3("month"),yearNumber) = oRec3("timer")
                                            aktTot = aktTot + oRec3("timer")

                                            monthNumber = monthNumber + 1
                                        oRec3.movenext
                                        wend
                                        oRec3.close                                                
                                        %>

                                        <%
                                        for i = 0 to antalmaaned
                                            if i = 0 then
                                            thisMonth = month(startdato)
                                            else
                                                thisMonth = thisMonth + 1
                                                if thisMonth > 12 then
                                                    thisMonth = 1
                                                    thisYear = 1
                                                end if
                                            end if
                                        %>
                                        <td style="text-align:right;"><b><%=formatnumber(timerMonthAkt(thisMonth, thisYear), 2) %></b></td>

                                        <%ekspTxt = ekspTxt & formatnumber(timerMonthAkt(thisMonth, thisYear), 2) & ";" %>

                                        <%next %>

                                        <td style="text-align:right;"><b><%=formatnumber(aktTot, 2) %></b></td>

                                        <%ekspTxt = ekspTxt & formatnumber(aktTot, 2) & ";" & chr(013) %>

                                    </tr>

                                    <%
                                    strSQLMedarb = "SELECT tmnr, mnavn FROM timer LEFT JOIN medarbejdere ON (tmnr = mid) WHERE taktivitetid = "& oRec2("taktivitetid") &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP BY tmnr ORDER BY mnavn"
                                    'response.Write " Medarb " & strSQLMedarb
                                    oRec3.open strSQLMedarb, oConn, 3
                                    while not oRec3.EOF
                                    medarbTot = 0
                                    %>
                                    <tr style="display:<%=displayList%>;" class="medarb_<%=oRec("jobid") %>_<%=oRec2("taktivitetid") %> medarb_<%=oRec("jobid") %>">
                                        <td> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp - &nbsp <span><%=oRec3("mnavn") %></span></td>

                                        <%ekspTxt = ekspTxt & "           - " & oRec3("mnavn") & ";" %>

                                        <%
                                        dim timerMonthMedarb
                                        redim timerMonthMedarb(12, 2) 'antal m�neder, �r-nummer

                                        strSQLMedarbTimer = "SELECT sum(timer) as timer, year(tdato) as year, month(tdato) as month FROM timer WHERE tmnr = "& oRec3("tmnr") &" AND taktivitetid = "& oRec2("taktivitetid") &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP by year(tdato), month(tdato) Order by tdato"
                                        'response.Write " MOnth timer medarb " & strSQLMedarbTimer
                                        oRec4.open strSQLMedarbTimer, oConn, 3
                                        monthNumber = 0
                                        while not oRec4.EOF
                                            if oRec4("year") <> year(startdato) then
                                            yearNumber = 1
                                            else
                                            yearNumber = 0
                                            end if

                                            timerMonthMedarb(oRec4("month"),yearNumber) = oRec4("timer")
                                            medarbTot = medarbTot + oRec4("timer")

                                            monthNumber = monthNumber + 1 
                                        oRec4.movenext
                                        wend
                                        oRec4.close
                                        %>

                                        <%
                                        for i = 0 to antalmaaned 
                                            if i = 0 then
                                            thisMonth = month(startdato)
                                            else
                                                thisMonth = thisMonth + 1
                                                if thisMonth > 12 then
                                                    thisMonth = 1
                                                    thisYear = 1
                                                end if
                                            end if
                                        %>
                                        <td style="text-align:right;"><%=formatnumber(timerMonthMedarb(thisMonth, thisYear), 2) %></td>

                                        <%ekspTxt = ekspTxt & formatnumber(timerMonthMedarb(thisMonth, thisYear), 2) & ";" %>

                                        <%next %>

                                        <td style="text-align:right;"><%=formatnumber(medarbTot, 2) %></td>

                                        <%ekspTxt = ekspTxt & formatnumber(medarbTot, 2) & ";" & chr(013) %>

                                    </tr>
                                    <%
                                    oRec3.movenext
                                    wend
                                    oRec3.close
                                    %>
                              <%
                                  oRec2.movenext
                                  wend
                                  oRec2.close
                              %>

                              <%
                                ekspTxt = ekspTxt & chr(13) & chr(13)

                                oRec.movenext                                 
                                wend
                                oRec.close
                              %>

                              <!-- TOTAL -->
                              <tr class="total">
                                  <th style="color:black;"><%=medarb_protid_txt_029 %></th>
                                   <%for i = 0 to antalmaaned %>
                                        <%jobsTotal = jobsTotal + periodTotal(i) %>
                                        <th style="text-align:right; color:black;"><%=formatnumber(periodTotal(i), 2) %></th>
                                   <%next %>
                                  <th style="text-align:right; color:black;"><%=formatnumber(jobsTotal, 2) %></th>
                              </tr>

                          </tbody>
                          <%else 'sogogsubmitted %>
                          <tbody>
                              <tr>
                                  <td colspan="15" style="text-align:center;">Ingen s�gekriterier valgt</td>
                              </tr>
                          </tbody>
                          <%end if %>

                         <!-- <tfoot>
                              <tr>
                                  <th><%=medarb_protid_txt_029 %></th>
                                   <%for i = 0 to antalmaaned %>
                                        <%jobsTotal = jobsTotal + periodTotal(i) %>
                                        <th style="text-align:right;"><%=formatnumber(periodTotal(i), 2) %></th>
                                   <%next %>
                                  <th style="text-align:right;"><%=formatnumber(jobsTotal, 2) %></th>
                              </tr>
                          </tfoot> -->

                      </table>
                      <%end if %>

                    <!-- Medarbedjer f�r aktviteter -->
                    <%if cint(FM_sortering) = 1 AND version = 0 then %>
                    <table class="table dataTable table-bordered table-hover">
                        <%call TableHeader %>
                        <%if request("sogsubmitted") = 1 then %>
                        <tbody>
                            <%
                                jobsTotal = 0
                                dim periodTotal2
                                redim periodTotal2(12)  'antal m�neder
                                'jobstatus = 1 AND
                                strSQL = "SELECT jobnavn, id as jobid, jobnr, kkundenavn FROM job LEFT JOIN timer ON (tjobnr = jobnr) LEFT JOIN kunder ON (jobknr = kid) WHERE tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'" & strSQLJobWhere & strSQLtfaktimWhere & " GROUP BY jobnr "& listSorteringSQL 
                                'response.Write strSQL
                                oRec.open strSQL, oConn, 3
                                while not oRec.EOF

                                dim timerMonth2
                                redim timerMonth2(12, 2) 'antal m�neder, �r-nummer
                                jobtot = 0
                                lastyear = -1
                                strSQLTimer = "SELECT sum(timer) as timer, month(tdato) as month, year(tdato) as year FROM timer WHERE tjobnr = '"& oRec("jobnr") &"' AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP by year(tdato), month(tdato) Order by tdato"
                                'response.Write " HERHER " & strSQLTimer
                                oRec2.open strSQLTimer, oConn, 3
                                monthNumber = 0
                                while not oRec2.EOF
                                    if oRec2("year") <> year(startdato) then
                                    yearNumber = 1
                                    else
                                    yearNumber = 0
                                    end if
                                    

                                    timerMonth2(oRec2("month"), yearNumber) = oRec2("timer")
                                    jobtot = jobtot + oRec2("timer")

                                    monthNumber = monthNumber + 1
                                oRec2.movenext
                                wend
                                oRec2.close

                                if listsort = 0 then
                                    tdOskrift = oRec("jobnavn") &" ("& oRec("jobnr") &")"
                                else
                                    tdOskrift = oRec("jobnr") &" ("& oRec("jobnavn") &")"
                                end if
                              %>

                              <tr style="background-color:#f9f9f9;">
                                  <td><b><span class="expandJob" style="cursor:pointer;" id="<%=oRec("jobid") %>"><%=tdOskrift %> <span style="font-size:12px;"> - <%=oRec("kkundenavn") %></span> <%if media <> "print" then %> <span id="icon_<%=oRec("jobid") %>" class="fa fa-angle-down"></span><%end if %></span></b></td>
                                  
                                  <%ekspTxt = ekspTxt & tdOskrift & " - " & oRec("kkundenavn") & ";" %>

                                  <%
                                      for i = 0 to antalmaaned 
                                         if i = 0 then
                                            thisMonth = month(startdato)
                                        else
                                            thisMonth = thisMonth + 1
                                            if thisMonth > 12 then
                                                thisMonth = 1
                                                thisYear = 1
                                            end if
                                        end if
                                  %>
                                  <%periodTotal2(i) = periodTotal2(i) + timerMonth2(thisMonth, thisYear) %>
                                  <td style="text-align:right;"><b><%=formatnumber(timerMonth2(thisMonth, thisYear), 2) %></b></td>

                                  <%ekspTxt = ekspTxt & formatnumber(timerMonth2(thisMonth, thisYear), 2) & ";" %>

                                  <%next %>

                                  <td style="text-align:right;"><b><%=formatnumber(jobtot, 2) %></b></td>

                                  <%ekspTxt = ekspTxt & ";" & chr(13) %>

                              </tr>

                              <%
                              strSQLMedarb = "SELECT mnavn, mid FROM medarbejdere LEFT JOIN timer ON (tmnr = mid) WHERE tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' AND tjobnr = '"& oRec("jobnr") &"' "& strSQLtfaktimWhere &" GROUP BY mnavn"
                              oRec2.open strSQLMedarb, oConn, 3
                              while not oRec2.EOF
                              dim timerMonthMedarb2
                              redim timerMonthMedarb2(12, 2) 'Antal m�nededer, �r-nummer
                              %>
                              <tr style="display:<%=displayList%>;" class="medarb_<%=oRec("jobid") %>">
                                  <td> &nbsp&nbsp - <b><span class="expandMedarb" style="cursor:pointer;" id="<%=oRec("jobid") %>_<%=oRec2("mid") %>"><%=oRec2("mnavn") %> <%if media <> "print" then %> <span id="icon_<%=oRec("jobid") %>_<%=oRec2("mid") %>" class="fa fa-angle-down"></span><%end if %></span></b></td>

                                  <%ekspTxt = ekspTxt & " - " & oRec2("mnavn") & ";" %>

                                  <%
                                  strSQLMedarbTimer = "SELECT sum(timer) as timer, year(tdato) as year, month(tdato) as month FROM timer WHERE tmnr = "& oRec2("mid") & " AND tjobnr = '"& oRec("jobnr") & "' AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP by year(tdato), month(tdato) Order by tdato"
                                  'response.Write strSQLMedarbTimer
                                  oRec3.open strSQLMedarbTimer, oconn, 3
                                  monthNumber = 0
                                  medarbTot = 0
                                  while not oRec3.EOF
                                        if oRec3("year") <> year(startdato) then
                                        yearNumber = 1
                                        else
                                        yearNumber = 0
                                        end if

                                      timerMonthMedarb2(oRec3("month"), yearNumber) = oRec3("timer")
                                      medarbTot = medarbTot + oRec3("timer")

                                      monthNumber = monthNumber + 1
                                  oRec3.movenext
                                  wend
                                  oRec3.close
                                  %>

                                  <%
                                    for i = 0 to antalmaaned 
                                        if i = 0 then
                                            thisMonth = month(startdato)
                                        else
                                            thisMonth = thisMonth + 1
                                            if thisMonth > 12 then
                                                thisMonth = 1
                                                thisYear = 1
                                            end if
                                        end if
                                  %>
                                  <td style="text-align:right;"><b><%=formatnumber(timerMonthMedarb2(thisMonth, thisYear), 2) %></b></td>

                                  <%ekspTxt = ekspTxt & formatnumber(timerMonthMedarb2(thisMonth, thisYear), 2) & ";" %>

                                  <%next %>

                                  <td style="text-align:right;"><b><%=formatnumber(medarbTot, 2) %></b></td>

                                  <%ekspTxt = ekspTxt & formatnumber(medarbTot, 2) & ";" & chr(13) %>

                              </tr>

                                <%
                                strSQLAkt = "SELECT navn, id FROM aktiviteter LEFT JOIN timer ON (taktivitetid = id) WHERE tjobnr = '"& oRec("jobnr") &"' AND tmnr = "& oRec2("mid") & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP BY id ORDER BY navn"
                                oRec3.open strSQLAkt, oConn, 3
                                Dim timerMonthAkt2
                                redim timerMonthAkt2(12, 2) 'antal m�neder �r-nummer                              
                                while not oRec3.EOF
                                %>
                                <tr style="display:<%=displayList%>;" class="akt_<%=oRec("jobid") %>_<%=oRec2("mid") %> akt_<%=oRec("jobid") %>">
                                    <td> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp - &nbsp <span><%=oRec3("navn") %></span></td>

                                    <%ekspTxt = ekspTxt & "           - " & oRec3("navn") & ";" %>

                                    <%
                                    strSQLaktTimer = "SELECT sum(timer) as timer, year(tdato) as year, month(tdato) as month FROM timer WHERE taktivitetid = "& oRec3("id") & " AND tmnr = "& oRec2("mid") &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP by year(tdato), month(tdato) Order by tdato"
                                    oRec4.open strSQLaktTimer, oConn, 3
                                    monthNumber = 0 
                                    yearNumber = 0
                                    aktTot = 0 
                                    while not oRec4.EOF
                                        if oRec4("year") <> year(startdato) then
                                        yearNumber = 1
                                        else
                                        yearNumber = 0
                                        end if

                                        timerMonthAkt2(oRec4("month"), yearNumber) = oRec4("timer")
                                        aktTot = aktTot + oRec4("timer")

                                        monthNumber = monthNumber + 1
                                    oRec4.movenext
                                    wend
                                    oRec4.close
                                    %>

                                    <%
                                    for i = 0 to antalmaaned 
                                        if i = 0 then
                                            thisMonth = month(startdato)
                                        else
                                            thisMonth = thisMonth + 1
                                            if thisMonth > 12 then
                                                thisMonth = 1
                                                thisYear = 1
                                            end if
                                        end if
                                    %>
                                    <td style="text-align:right;"><%=formatnumber(timerMonthAkt2(thisMonth, thisYear), 2) %></td>

                                    <%ekspTxt = ekspTxt & formatnumber(timerMonthAkt2(thisMonth, thisYear), 2) & ";"  %>

                                    <%next %>

                                    <td style="text-align:right;"><%=formatnumber(aktTot, 2) %></td>

                                    <%ekspTxt = ekspTxt & formatnumber(aktTot, 2) & ";" & chr(13) %>

                                </tr>
                                <%
                                oRec3.movenext
                                wend
                                oRec3.close
                                %>

                               <%
                                oRec2.movenext
                                wend
                                oRec2.close
                               %>
                                
                                <%ekspTxt = ekspTxt & chr(13) & chr(13) %>

                              <%
                                oRec.movenext                                 
                                wend
                                oRec.close
                              %>


                            <!-- TOTAL -->
                            <tr style="background-color:#f5f5f5;">
                                <th style="color:black;"><%=medarb_protid_txt_029 %></th>
                                <%for i = 0 to antalmaaned %>
                                    <%jobsTotal = jobsTotal + periodTotal2(i) %>
                                    <th style="text-align:right; color:black;"><%=formatnumber(periodTotal2(i), 2) %></th>
                                <%next %>     
                                <th style="text-align:right; color:black;"><%=formatnumber(jobsTotal, 2) %></th>
                            </tr>

                        </tbody>
                        <%else 'sogogsubmitted %>
                          <tbody>
                              <tr>
                                  <td colspan="15" style="text-align:center;">Ingen s�gekriterier valgt</td>
                              </tr>
                          </tbody>
                        <%end if %>
                        
                       <!-- <tfoot>
                            <tr>
                                <th><%=medarb_protid_txt_029 %></th>
                                <%for i = 0 to antalmaaned %>
                                    <%jobsTotal = jobsTotal + periodTotal2(i) %>
                                    <th style="text-align:right;"><%=formatnumber(periodTotal2(i), 2) %></th>
                                <%next %>     
                                <th style="text-align:right;"><%=formatnumber(jobsTotal, 2) %></th>
                            </tr>
                        </tfoot> -->

                    </table>
                    <%end if %>

                    <!-- Medarbejder aller f�rst --> 
                    <%if version = 1 then %>
                    <table class="table dataTable table-bordered table-hover">
                        <%call TableHeader %>
           
                        <%if request("sogsubmitted") = 1 then %>
                        <tbody>
                            <%
                            medarbsTotal = 0
                            dim totPeriod
                            redim totPeriod(12) 'antal m�neder
                            strSQLmedarb = "SELECT mnavn, init, mid FROM medarbejdere LEFT JOIN timer ON (mid = tmnr) WHERE tdato BETWEEN '"& startdato &"' AND '"& slutdato &"'" & strSQLMedWhere & strSQLtfaktimWhere & " GROUP BY mid ORDER BY mnavn"
                            
                            'Response.write strSQLmedarb
                            'Response.flush
                                
                            oRec.open strSQLmedarb, oConn, 3
                            while not oRec.EOF
                            %>
                            <tr>
                                <td><b><span class="expandMedarb" style="cursor:pointer;" id="<%=oRec("mid") %>"><%=oRec("mnavn") & " (" &oRec("init")& ")" %> <%if media <> "print" AND media <> "print_sim" then %> <span id="icon_<%=oRec("mid") %>" class="fa fa-angle-down"></span><%end if %></span></b></td>

                                <%
                                    ekspTxt = ekspTxt & oRec("mnavn") &" ("&oRec("init")&");"

                                    dim monthTimermedarb
                                    redim monthTimermedarb(12, 2) 'Antal m�nneder �r-nummer
                                    strSQLMedTimer = "SELECT sum(timer) as timer, year(tdato) as year, month(tdato) as month FROM timer WHERE tmnr = "& oRec("mid") & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP by year(tdato), month(tdato) Order by tdato"
                                    oRec2.open strSQLMedTimer, oConn, 3
                                    monthNumber = 0
                                    yearNumber = 0
                                    medarbTotal = 0
                                    while not oRec2.EOF
                                        if oRec2("year") <> year(startdato) then
                                        yearNumber = 1
                                        else
                                        yearNumber = 0
                                        end if

                                        monthTimermedarb(oRec2("month"), yearNumber) = oRec2("timer")
                                        medarbTotal = medarbTotal + oRec2("timer")

                                        monthNumber = monthNumber + 1                                  
                                    oRec2.movenext
                                    wend
                                    oRec2.close
                                %>

                                <%
                                    for i = 0 to antalmaaned
                                        if i = 0 then
                                            thisMonth = month(startdato)
                                        else
                                            thisMonth = thisMonth + 1
                                            if thisMonth > 12 then
                                                thisMonth = 1
                                                thisYear = 1
                                            end if
                                        end if
                                %>
                                    <%totPeriod(i) = totPeriod(i) + monthTimermedarb(thisMonth, thisYear) %>
                                    <td style="text-align:right;"><b><%=formatnumber(monthTimermedarb(thisMonth, thisYear), 2) %></b></td>
                                    
                                    <%ekspTxt = ekspTxt & formatnumber(monthTimermedarb(thisMonth, thisYear), 2) & ";" %>

                                <%next %>

                                <td style="text-align:right;"><b><%=formatnumber(medarbTotal, 2) %></b></td>

                                <%ekspTxt = ekspTxt & formatnumber(medarbTotal, 2) & ";" & chr(013) %>

                            </tr>

                            <%
                            strSQLJob = "SELECT jobnavn, jobnr, id as jobid, kkundenavn FROM job LEFT JOIN timer ON (tjobnr = jobnr) LEFT JOIN kunder ON (jobknr = kid) WHERE tmnr = "& oRec("mid") & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP BY jobnr ORDER BY jobnavn"
                            oRec2.open strSQLJob, oConn, 3
                            while not oRec2.EOF
                            %>
                            <tr class="job_<%=oRec("mid") %>" style="display:<%=displayList%>;">
                                <td> &nbsp&nbsp - <b><span class="expandJob" style="cursor:pointer;" id="<%=oRec2("jobid") %>_<%=oRec("mid") %>"><%=oRec2("jobnavn") & " ("&oRec2("jobnr")&")" %> <span style="font-size:12px;"> - <%=oRec2("kkundenavn") %></span> <%if media <> "print" then %> <span class="fa fa-angle-down" id="icon_<%=oRec2("jobid") %>_<%=oRec("mid") %>"></span><%end if %></span></b></td>

                                <%ekspTxt = ekspTxt & " - " & oRec2("jobnavn") & "("&oRec2("jobnr")&") - " & oRec2("kkundenavn") & ";" %>

                                <%
                                dim monthtimerJob
                                redim monthtimerJob(12, 2) 'antal m�nender, �r-nummer

                                strSQLJobTimer = "SELECT sum(timer) as timer, year(tdato) as year, month(tdato) as month FROM timer WHERE tjobnr = '"& oRec2("jobnr") & "' AND tmnr = "& oRec("mid") & " AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP by year(tdato), month(tdato) ORDER BY tdato"
                                oRec3.open strSQLJobTimer, oConn, 3
                                monthNumber = 0
                                yearNumber = 0
                                jobTot = 0
                                while not oRec3.EOF
                                    if oRec3("year") <> year(startdato) then
                                    yearNumber = 1
                                    else
                                    yearNumber = 0
                                    end if

                                    monthtimerJob(oRec3("month"), yearNumber) = oRec3("timer")
                                    jobTot = jobTot + oRec3("timer")

                                    monthNumber = monthNumber + 1
                                oRec3.movenext
                                wend
                                oRec3.close
                                %>

                                <%
                                thisYear = 0
                                for i = 0 to antalmaaned 
                                    if i = 0 then
                                        thisMonth = month(startdato)
                                    else
                                        thisMonth = thisMonth + 1
                                        if thisMonth > 12 then
                                            thisMonth = 1
                                            thisYear = 1
                                        end if
                                    end if
                                %>
                                    <td style="text-align:right;"><b><%=formatnumber(monthtimerJob(thisMonth, thisYear), 2) %></b></td>

                                    <%ekspTxt = ekspTxt & formatnumber(monthtimerJob(thisMonth, thisYear), 2) & ";" %>

                                <%next %>
                                
                                <td style="text-align:right;"><b><%=formatnumber(jobTot, 2) %></b></td>

                                <%ekspTxt = ekspTxt & formatnumber(jobTot, 2) & ";" & chr(013) %>

                            </tr>   
                            
                            <%
                            strSQLAkt = "SELECT id, navn FROM aktiviteter LEFT JOIN timer ON (taktivitetid = id) WHERE tmnr = "& oRec("mid") & " AND tjobnr = '"& oRec2("jobnr") &"' AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP BY taktivitetid ORDER BY navn"
                            oRec3.open strSQLAkt, oConn, 3
                            while not oRec3.EOF
                            %>
                            <tr class="akt_<%=oRec2("jobid") %>_<%=oRec("mid") %> akt_<%=oRec("mid") %>" style="display:<%=displayList%>;">
                                <td> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp - &nbsp <span><%=oRec3("navn") %></span></td>

                                <%ekspTxt = ekspTxt & "           - " & oRec3("navn") & ";" %>

                                <%
                                dim monthTimerAkt
                                redim monthTimerAkt(12, 2) 'antal m�neder, �r nummer
                                strSQLAktTimer = "SELECT sum(timer) as timer, year(tdato) as year, month(tdato) as month FROM timer WHERE tmnr = "& oRec("mid") & " AND taktivitetid = "& oRec3("id") &" AND tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' "& strSQLtfaktimWhere &" GROUP by year(tdato), month(tdato) ORDER BY tdato"
                                'response.Write "strSQLAktTimer " & strSQLAktTimer
                                oRec4.open strSQLAktTimer, oConn, 3
                                aktTot = 0
                                monthNumber = 0
                                yearNumber = 0
                                while not oRec4.EOF

                                    if oRec4("year") <> year(startdato) then
                                    yearNumber = 1
                                    else
                                    yearNumber = 0
                                    end if

                                    monthTimerAkt(oRec4("month"), yearNumber) = oRec4("timer")
                                    aktTot = aktTot + oRec4("timer")

                                    monthNumber = monthNumber + 1
                                oRec4.movenext
                                wend
                                oRec4.close
                                %>

                                <%
                                thisYear = 0
                                for i = 0 to antalmaaned
                                    if i = 0 then
                                     thisMonth = month(startdato)
                                    else
                                        thisMonth = thisMonth + 1
                                        if thisMonth > 12 then
                                            thisMonth = 1
                                            thisYear = 1
                                        end if
                                    end if
                                %>

                                   <td style="text-align:right;"><%=formatnumber(monthTimerAkt(thisMonth,thisYear), 2) %></td>

                                   <%ekspTxt = ekspTxt & formatnumber(monthTimerAkt(thisMonth,thisYear), 2) & ";" %>

                                <%next %>

                                <td style="text-align:right;"><%=formatnumber(aktTot, 2) %></td>

                                <%ekspTxt = ekspTxt & formatnumber(aktTot, 2) & ";" & chr(013)  %>

                            </tr>
                            <%
                            oRec3.movenext
                            wend
                            oRec3.close
                            %>

                            <%
                            oRec2.movenext
                            wend
                            oRec2.close
                            %>

                            <%

                            ekspTxt = ekspTxt & chr(013)

                            oRec.movenext
                            wend
                            oRec.close
                            %>


                            <!-- TOTAL -->
                            <tr style="background-color:#f5f5f5;">
                                <th style="color:black;"><%=medarb_protid_txt_029 %></th>
                                <%for i = 0 to antalmaaned %>
                                    <%medarbsTotal = medarbsTotal + totPeriod(i)%>
                                    <th style="text-align:right; color:black;"><%=formatnumber(totPeriod(i), 2) %></th>
                                <%next %>
                                <th style="text-align:right; color:black;"><%=formatnumber(medarbsTotal, 2) %></th>
                            </tr>

                        </tbody>
                        <%else 'sogogsubmitted %>
                          <tbody>
                              <tr>
                                  <td colspan="15" style="text-align:center;">Ingen s�gekriterier valgt</td>
                              </tr>
                          </tbody>
                        <%end if %>
                       <!-- <tfoot>
                            <tr>
                                <th><%=medarb_protid_txt_029 %></th>
                                <%for i = 0 to antalmaaned %>
                                    <%medarbsTotal = medarbsTotal + totPeriod(i)%>
                                    <th style="text-align:right;"><%=formatnumber(totPeriod(i), 2) %></th>
                                <%next %>
                                <th style="text-align:right;"><%=formatnumber(medarbsTotal, 2) %></th>
                            </tr>
                        </tfoot> -->

                    </table>
                    <%end if %>


                      <%if media <> "print" AND media <> "print_sim" then %>
                      <br /><br /><br />
                      <section class="printhide">
                        <div class="row">
                             <div class="col-lg-12">
                                <b><%=medarb_protid_txt_006 %></b>
                                </div>
                            </div>

                          <!--
                            <form action="medarb_protid.asp?media=print&aar=<%=aar %>&aarslut=<%=aarslut %>&version=<%=version %>&FM_job=<%=request("FM_job") %>&FM_medarb=<%=request("FM_medarb") %>" method="post" target="_blank">               
                                <div class="row">
                                    <div class="col-lg-12 pad-r30">
                                        <input style="width:125px;" id="Submit5" type="submit" value="<%=medarb_protid_txt_007 %>" class="btn btn-sm" /><br />
                                    </div>
                                </div>
                           </form>

                          <form action="medarb_protid.asp?media=print_sim&aar=<%=aar %>&aarslut=<%=aarslut %>&version=<%=version %>&FM_job=<%=request("FM_job") %>&FM_medarb=<%=request("FM_medarb") %>" method="post" target="_blank">          
                                <div class="row">
                                    <div class="col-lg-12 pad-r30">
                                        <input style="width:125px;" type="submit" value="Printvenlig simpel" class="btn btn-sm" /><br />
                                    </div>
                                </div>
                           </form> -->

                        <div class="row">
                            <div class="col-lg-12 pad-r30">
                                <input style="width:125px;" type="submit" value="<%=medarb_protid_txt_007 %>" id="nyprint" class="btn btn-default btn-sm" /><br />
                            </div>
                        </div>

                        <%if version = 1 then %>
                        <form action="medarb_protid.asp?media=export&version=1&FM_medarb=<%=request("FM_medarb") %>&FM_progrp=<%=request("FM_progrp") %>&sogsubmitted=1&aar=<%=aar %>&aarslut=<%=request("aarslut") %>" method="post" target="_blank">                           
                            <div class="row">
                                <div class="col-lg-12 pad-r30">
                                    <input style="width:125px;" type="submit" value="Eksport" class="btn btn-default btn-sm" /><br />
                                </div>
                            </div>
                        </form>
                        <%else %>
                        <form action="medarb_protid.asp?media=export&sogsubmitted=1&FM_job=<%=request("FM_job") %>&FM_sortering=<%=FM_sortering %>" method="post" target="_blank"> 
                            <div class="row">
                                <div class="col-lg-12 pad-r30">
                                    <input style="width:125px;" type="submit" value="Eksport" class="btn btn-default btn-sm" /><br />
                                </div>
                            </div>
                        </form>
                        <%end if %>
                    </section>
                    <%
                        
                    else
                    Response.Write("<script language=""JavaScript"">window.print();</script>")                                         
                    end if 
                    %>


                      <%if media = "export" then

                        if version = 1 then
                            strOskrifter = strOskrifter & "Total per medarbejder;"
                        else
                          strOskrifter = strOskrifter & "Total per job;"
                        end if

                       ' ekspTxt = "HER;HER" 'ekspTxt 'request("FM_ekspTxt")
	                    ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	
	
	                    filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                    filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				        Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				        if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\medarb_protidexp_.asp" then
					        Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarb_protidexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					        Set objNewFile = nothing
					        Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarb_protidexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				        else
					        Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarb_protidexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					        Set objNewFile = nothing
					        Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarb_protidexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				        end if
				
				
				
				        file = "medarb_protidexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				        '**** Eksport fil, kolonne overskrifter ***
			
				        'strOskrifter = "Medarbejder; Nr.; Init; Status; Medarbejertype; Brugergruppe; Email; Ansatdato"
				
				
				
				
				        objF.WriteLine(strOskrifter & chr(013))
				        objF.WriteLine(ekspTxt)
				        objF.close
				
				        
                      end if %>


            </div>
            </div>
        </div>

        <%if media = "export" then %>
		    <div style="position:absolute; top:100px; left:200px; width:300px;">
	        <table border=0 cellspacing=1 cellpadding=0 width="100%">
	        <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	        <img src="../ill/outzource_logo_200.gif" />
	        </td>
	        </tr>
	        <tr>
	        <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
                
                             
	        <a href="../inc/log/data/<%=file%>" target="_blank" ><%=medarb_txt_115 & " "  %>>></a>
	        </td></tr>
	        </table>
            </div>
        <%end if %>

       <%if media <> "print" AND media <> "print_sim" then %>
         <!-- </div>
      </div> -->
      <%end if %>

    <div class="luft"><br /><br /><br /><br /><br /><br /></div>


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