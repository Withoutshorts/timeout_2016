

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<% 
func = request("func")
media = request("media")

select case func
case "export"

    ekspTxt = request("exporttxt")

     '******************* Eksport **************************' 
                if media = "export" then

                    call TimeOutVersion()
    

                    ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
                    
	               
	                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)


                    fileext = "csv"
                    
                    
	
				                Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\medarbdashboard.asp" then
					                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\mdashexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& fileext, True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\mdashexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& fileext, 8)
				                else
					                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\mdashexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& fileext, True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\mdashexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& fileext, 8)
				                end if
				
				
                                file = "mdashexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."& fileext
				
				              

				                objF.WriteLine(ekspTxt)
				                objF.close
				
				                %>
				                <div style="padding:20px 40px 40px 40px;">
	                            <table border=0 cellspacing=1 cellpadding=0 width="300">
	                            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	                            <img src="../ill/outzource_logo_200.gif" />
	                            </td>
	                            </tr>
	                            <tr>
	                            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	                            <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank" onClick="Javascript:window.close()">Your .csv file is ready >></a>
	                            </td></tr>
	                            </table>
	                            </div>
	          
	            
	                            <%
                
                                'Response.redirect "../inc/log/data/"& file &""	
                                response.end


                end if
case else
%>




<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--include file="inc/budget_firapport_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->


<script type="text/javascript" src="js/libs/excanvas.compiled.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.js"></script>

<!--
<script type="text/javascript" src="js/plugins/flot/jquery.flot.orderBars.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.resize.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.stack.js"></script>
<script type="text/javascript" src="js/plugins/flot/jquery.flot.tooltip.min.js"></script>
<script type="text/javascript" src="js/demos/flot/stacked-vertical.js"></script>
<script type="text/javascript" src="js/demos/flot/vertical.js"></script>
<script type="text/javascript" src="js/demos/flot/scatter.js"></script>
<script type="text/javascript" src="js/demos/flot/line.js"></script>
 -->
<!--
<script type="text/javascript" src="js/demos/flot/donut_saelger2.js"></script>
-->

<script src="js/medarbdashboard_jav.js"></script>
<%select case lto
  case "epi2017"
    %>
    <script type="text/javascript" src="js/demos/flot/donut_epi2017.js"></script>
    <%
  case else %>
  <script type="text/javascript" src="js/demos/flot/donut.js"></script>
<%end select %>

<script type="text/javascript" src="js/plugins/flot/jquery.flot.pie.js"></script>


  
<%call menu_2014 %>


<% 
    lastMjobansBgtDBClosedGT = 0
    lastMsalgsansBgtDBClosedGT = 0
    antalMLastMidClosedGt = 0
    lastMUtil = 0 

    sub MatrixPrMedarb

    jobansBgtGT = jobansBgtGT/1 + lastMjobansBgt/1
    salgsansBgtGT = salgsansBgtGT/1 + lastMsalgsansBgt/1
    
    fakbeloebGT = fakbeloebGT/1 + lastMfakbeloeb/1
    matforbrugsumTot = matforbrugsumTot/1 + lastMmedarbSalgsOkost/1

    'lastMdbthis = (lastMfakbeloeb/1 - lastMmedarbSalgsOkost/1)
    'lastMdbthisproc = 100 - ((lastMmedarbSalgsOkost*1/lastMfakbeloeb*1) * 100)

    
    if lastMjobansBgtDBClosed <> 0 then 'lastMmedarbJobOkostProc <> 0 AND
    lastMjobansBgtDBClosed = ((lastMjobansBgtDBClosed - lastMmedarbJobOkostProc) / lastMjobansBgtDBClosed) * 100 'antalMLastMidClosed 
    else
    lastMjobansBgtDBClosed = 0
    end if

    if lastMsalgsansBgtDBClosed <> 0 then 'lastMmedarbSalgsOkostProc <> 0 AND
    lastMsalgsansBgtDBClosed = ((lastMsalgsansBgtDBClosed - lastMmedarbSalgsOkostProc) / lastMsalgsansBgtDBClosed) * 100 'antalMLastMidClosed 
    else
    lastMsalgsansBgtDBClosed = 0
    end if

    

    'lastMjobansBgtDBClosed = lastMjobansBgtDBClosed

    'lastMsalgsansBgtDBClosed = (lastMsalgsansBgtDBClosed/antalMLastMidClosed) 
    'else
    'lastMjobansBgtDBClosed = 0
    'lastMsalgsansBgtDBClosed = 0
    'end if

    lastMjobansBgtDBClosedGT = lastMjobansBgtDBClosedGT + lastMjobansBgtDBClosed
    lastMsalgsansBgtDBClosedGT = lastMsalgsansBgtDBClosedGT + lastMsalgsansBgtDBClosed

    antalMLastMidClosedGt = antalMLastMidClosedGt + 1

    if ntimPerytdprM(lastMid) <> 0 AND faktimerGTselmedarbAar(lastMid) <> 0 then

    lastMUtil = (faktimerGTselmedarbAar(lastMid)*1 / ntimPerytdprM(lastMid)*1)
    
    else
    lastMUtil = 0
    end if

    '/ faktimerGTselmedarbAar(lastMid)*1) / ntimPerytdprM(lastMid)*1

    'dbGT = dbGT*1 + lastMdbthis*1

     %> 
                                           
    <td><%=LastMnavn %> [<%=lastMinit %>]</td>
        
     <!--
    <td style="text-align:right; white-space:nowrap;"><%=formatnumber(lastMfakbeloeb, 2) &" "& basisValISO_f8%></td>
    <td style="text-align:right; white-space:nowrap;"><%=formatnumber(lastMmedarbSalgsOkost, 2) &" "& basisValISO_f8%></td>
      -->
                                          
    <td style="text-align:right; white-space:nowrap;"> <%=formatnumber(lastMjobansBgt, 2) &" "& basisValISO_f8 %>
    
    </td>
    <td style="text-align:right; white-space:nowrap;"> <%=formatnumber(lastMjobansBgtDBClosed, 0) &" %"%></td>    
    <td style="text-align:right; white-space:nowrap;"> <%=formatnumber(lastMsalgsansBgt, 2) &" "& basisValISO_f8 %></td>
    <td style="text-align:right; white-space:nowrap;"> <%=formatnumber(lastMsalgsansBgtDBClosed, 0) &" %"%></td>

    <!--<td style="text-align:right; white-space:nowrap;"><%=formatnumber(lastMdbthis, 2) &" "& basisValISO_f8 %></td>-->

    <td style="text-align:right; white-space:nowrap;"><%=formatnumber(lastMUtil, 2) %> <br />

        
        <!--// <%="ntimPerytdprM(lastMid) - faktimerGTselmedarbAar(lastMid)) / ntimPerytdprM(lastMid) "& ntimPerytdprM(lastMid) &" - " & faktimerGTselmedarbAar(lastMid) &" / "& ntimPerytdprM(lastMid) %>-->

    </td>

    <%

                                       
    strEksport = strEksport & LastMnavn &";"& lastMinit
    'strEksport = strEksport &";"& formatnumber(lastMjo_bruttooms, 2) &";"& formatnumber(lastMfakbeloeb, 2) 
    'strEksport = strEksport &";"& formatnumber(lastMmedarbSalgsOkost/1, 2) 
    
    strEksport = strEksport &";"& formatnumber((lastMjobansBgt/1), 2) &";"& formatnumber((lastMjobansBgtDBClosed/1), 0)
    strEksport = strEksport &";"& formatnumber((lastMsalgsansBgt/1), 2) &";"& formatnumber((lastMsalgsansBgtDBClosed/1), 0) 
                                                
    strEksport = strEksport &";"& formatnumber(lastMUtil, 2) & "xx99123sy#z"


   

          

                                          
    lastMfakbeloeb = 0
    lastMmedarbSalgsOkost = 0
    lastMjobansBgt = 0
    lastMsalgsansBgt = 0
    lastMdbthis = 0 

    lastMjobansBgtDBClosed = 0
    lastMsalgsansBgtDBClosed = 0
    antalMLastMid = 0                                  
    'antalMLastMidClosed = 0

    lastMjobsansProc = 0
    lastMsalgssansProc = 0

    lastMmedarbSalgsOkostProc = 0
    lastMmedarbJobOkostProc = 0

    end sub



    
   call basisValutaFN() 
    
   if len(trim(request("FM_periode"))) <> 0 then
   usePeriode = request("FM_periode")
   else
   usePeriode = 1 'm�ned
   end if 


   if len(trim(request("FM_sog"))) <> 0 then
    jobSogVal = request("FM_sog")
    jobSog = 1
   else
    jobSogVal = ""
    jobSog = 0
    end if
    

    if cint(jobSog) = 1 then
    progrpmedarbDisabled = "DISABLED"
    else
    progrpmedarbDisabled = ""
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

    
    if len(trim(request("business_area_filter"))) <> 0 then
    business_area_filter = request("business_area_filter")
    else
    business_area_filter = ""
    end if

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

    antalMids = 0
    multipleMids = 0
    useMids = 0
    strSQLmids = " AND (mid = 0"
    if len(trim(request("FM_medarb"))) <> 0 AND request("FM_medarb") <> "0" AND request("sogsubmitted") = "1" then

        if instr(request("FM_medarb"), ",") <> 0 then 'Array / Multible 
        
            multipleMids = 1
            useMids = split(request("FM_medarb"), ", ")
            usemrn = session("mid") 


            for m = 0 TO UBOUND(useMids)
            strSQLmids = strSQLmids & " OR mid = "& useMids(m) &""
        
                if m = 0 then
                usemrn = useMids(m) 'f�rst valgte
                end if

            antalMids = antalMids + 1
            next

            strSQLmids = strSQLmids & ")"

        else
        usemrn = request("FM_medarb")
        end if
    else
        if request("sogsubmitted") = "1" then
        usemrn = -1
        else
        usemrn = session("mid") 
        end if
    end if

   


    if len(trim(request("FM_progrp"))) <> 0 then
    pgrp = request("FM_progrp")
    else
            if level = 1 then
            pgrp = 10
            else
            pgrp = -1
            end if
    end if


    level = session("rettigheder")

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

    select case usePeriode
    case 1
    maksHoursBar = 185
    case 2
    maksHoursBar = 185 * 3
    case 3
    maksHoursBar = 185 * 12
    end select
    %>

<div class="wrapper">
    <div class="content">
    
    <div class="container">
 
      

        <div class="portlet">
        <h3 class="portlet-title"><u><%=dsb_txt_001 %></u></h3>
        

        <div class="well">
              <form action="medarbdashboard.asp?sogsubmitted=1" name="dashboard_filter" id="dashboard_filter" method="POST">
                <div class="row">
                    <div class="col-lg-2">
                    <h4 class="panel-title-well"><%=dsb_txt_002 %></h4></div>

                    

                    <div class="col-lg-10">&nbsp</div>
                 
                    </div>
                  
                    <div class="row">
                    <div class="col-lg-4"><%=dsb_txt_003 %>:</div>
                    <!--<div class="col-lg-2"><%=dsb_txt_004 %>:</div>-->
                    <div class="col-lg-2"><input type="radio" name="FM_periode" value="1" onchange="submit()" <%=usePeriodeSel1 %> /> <%=dsb_txt_005 %>:</div>
                    <div class="col-lg-2"><input type="radio" name="FM_periode" value="2" onchange="submit()" <%=usePeriodeSel2 %> /> <%=dsb_txt_006 %>:</div>
                    <div class="col-lg-2"><input type="radio" name="FM_periode" value="3" onchange="submit()" <%=usePeriodeSel3 %> /> <%=dsb_txt_007 %>:</div>
             
                   
                

            </div>
            <div class="row">

                

                <div class="col-lg-4">

                                 <%
                                   
                                    if level = 1 then

                                    select case lto
                                    case "epi2017"
                                    strSQLpgrp = "SELECT id, navn FROM projektgrupper WHERE id <> 0 AND instr(navn, 'hr') > 0 OR (id = 2 OR id = 3 OR id = 11 OR id = 12 OR id = 13) ORDER BY orgvir, navn"
                                    case else
                                    strSQLpgrp = "SELECT id, navn FROM projektgrupper WHERE id <> 0 ORDER BY navn"
                                    end select

                                    else

                                    strProgrpSQLkri = " id = -1 "

                                        strSQLprogrpErTeamleder = "SELECT projektgruppeid FROM progrupperelationer WHERE medarbejderId = "& session("mid") &" AND teamleder = 1 GROUP BY projektgruppeid"
                                        
                                        'response.write strSQLprogrpErTeamleder
                                        'response.flush

                                        oRec3.open strSQLprogrpErTeamleder, oConn, 3
                                        while not oRec3.EOF
                                    
                                        strProgrpSQLkri = strProgrpSQLkri &" OR id = "&  oRec3("projektgruppeid")    
                                  
                                        oRec3.movenext
                                        wend
                                        oRec3.close

                                    
                                        strProgrpSQLkri = strProgrpSQLkri 

                                      strSQLpgrp = "SELECT id, navn FROM projektgrupper WHERE "& strProgrpSQLkri &" ORDER BY navn"

                                    end if


                                    'response.write strSQLpgrp
                                   
                                %>
                                <select id="jq_progrp" name="FM_progrp" <%=progrpmedarbDisabled  %> class="form-control input-small"> <!-- onchange="submit();" -->
                                  <%

                                    oRec3.open strSQLpgrp, oConn, 3
                                    while not oRec3.EOF 

                                    if cint(pgrp) = cint(oRec3("id")) then
                                    pSel = "SELECTED"
                                    else
                                    pSel = ""
                                    end if
                                    %>
                                    <option value="<%=oRec3("id") %>" <%=pSel %>><%=oRec3("navn")%></option>
                                    <%
                                    oRec3.movenext
                                    wend 
                                    oRec3.close


                                if cint(pgrp) = -1 then
                                 ingenSEL = "SELECTED"
                                else
                                 ingenSEL = ""
                                end if

                                 %>
                                <option value="-1" <%=ingenSEL %>>None</option>
                                </select>

                                <br />        

                                <%select case lto
                                 case "xepi2017"
                                    med_multiple  = ""
                                    med_size = 1
                                    med_oc = "submit();"
                                 case else
                                    med_multiple  = "multiple"
                                    med_size = 5
                                    med_oc = ""
                                 end select %>

                                <%
                                    meFound = 0
                                    if level <= 2 OR level = 6 then
                                    
                                    call medarbiprojgrp(pgrp, session("mid"), 0, -1)
                                    
                                    strSQLm = "SELECT mid, mnavn, init FROM medarbejdere WHERE mansat = 1 AND (mid = 0 "& medarbgrpIdSQLkri &") ORDER BY mnavn"

                                    else

                                    strSQLm = "SELECT mid, mnavn, init FROM medarbejdere WHERE mansat = 1 AND (mid = "& session("mid") &") ORDER BY mnavn"

                                    end if

                                    'Response.write strSQLm

                                    %>
                                     <select id="jq_medid" name="FM_medarb" <%=progrpmedarbDisabled%> size="<%=med_size %>" <%=med_multiple%> class="form-control input-small" onchange="<%=med_oc%>"> <!-- onchange="submit();"--->
                                
                                    <%

                                    if med_multiple = "multiple" then
                                    %><option value="0">All</option> <%
                                    end if


                                    

                                    oRec3.open strSQLm, oConn, 3
                                    while not oRec3.EOF 

                                    if (cdbl(usemrn) = cdbl(oRec3("mid")) AND cint(multipleMids) = 0) OR (cint(multipleMids) = 1 AND instr(strSQLmids, "OR mid = "& oRec3("mid") &"") <> 0)  then
                                    mSel = "SELECTED"
                                    meFound = 1
                                    else
                                    mSel = ""
                                    
                                    end if
                                    %>
                                    <option value="<%=oRec3("mid") %>" <%=mSel %>><%=oRec3("mnavn")& " ["& oRec3("init") &"]" %></option>
                                    <%
                                    oRec3.movenext
                                    wend 
                                    oRec3.close
                                 
                                        
                                     if cint(meFound) = 0 AND med_multiple <> "multiple" then%>
                                           <option value="0" SELECTED>Choose..</option>
                                     <%end if %>
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
                <select name="FM_kv" class="form-control input-small" onchange="submit()">
                  <option value="1" <%=useQuaterSel1 %>>Kvt. 1</option>
                  <option value="2" <%=useQuaterSel2 %>>Kvt. 2</option>
                  <option value="3" <%=useQuaterSel3 %>>Kvt. 3</option>
                  <option value="4" <%=useQuaterSel4 %>>Kvt. 4</option>
                </select>
                 </div>

                      <div class="col-lg-2">

                          <% 
                              yearSt = datePart("yyyy", dateAdd("yyyy", -3, now), 2,2) 
                           yearEnd = yearSt + 4 
                              
                               %>
                        <select name="FM_aar" class="form-control input-small" onchange="submit()">
                            <%
                          
     
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
                  <div class="row">
                  <%if level <= 2 OR level = 6 then %>
                   
                            <div class="col-lg-4">
                                <div style="padding:4px 0px 4px 0px;"><%=dsb_txt_010 %>:</div> 
                        
                                <input type="text" name="FM_sog" id="FM_sog" value="<%=jobSogVal %>" class="form-control input-small" />
                          </div>


                        <%select case lto
                          case "xxepi2017" '** Projektgrupepr bruges istedet for


                           business_area_filterADM = ""
                           business_area_filterAVI = ""
                           business_area_filterTRA = ""
                           business_area_filterPUB = ""
                           business_area_filterPRI = ""

                            select case business_area_filter
                            case "ADM"
                            business_area_filterADM = "SELECTED"
                            case "AVI"
                            business_area_filterAVI = "SELECTED"
                            case "TRA"
                            business_area_filterTRA = "SELECTED"
                            case "PUB"
                            business_area_filterPUB = "SELECTED"
                            case "PRI"
                            business_area_filterPRI = "SELECTED"
                            end select

                            %>
                            <div class="col-lg-4">Business area:
                            <select name="business_area_filter" class="form-control input-small" onchange="submit();">
                                <option value="">All</option>
                                <option value="ADM" <%=business_area_filterADM %>>Admin</option>
                                <option value="AVI" <%=business_area_filterAVI %>>Avi</option>
                                <option value="TRA" <%=business_area_filterTRA %>>Trans</option>
                                <option value="PUB" <%=business_area_filterPUB %>>Public</option>
                                <option value="PRI" <%=business_area_filterPRI %>>Private</option>
                            
                            </select>
                             </div>
                           <%end select%>

                         <div class="col-lg-7"><br />
                           <button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=dsb_txt_008 %> >></b></button>

                      </div>


                <%else %>
                   <div class="col-lg-11" style="vertical-align:bottom;">
                           <button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=dsb_txt_008 %> >></b></button>

                      </div>
                  <%end if %>

                   
                  </div>

                  
                 
            </form>
            </div>


            

       






        <%

        if usemrn <> "-1" then
         
         call akttyper2009(2)

         meTxt = "-"
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



        '*****************************************************************
        '*** Vis tot p� job
        '*****************************************************************
        if cint(jobSog) = 1 then


        %>
            <!-- Alle projekter med Real. timer p� for den valgte medarb. --->

            <br /><br />
              <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseSix">Job <span style="font-size:12px; font-weight:lighter;">Alle medarbejdere i valgt periode</span></a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseSix" class="panel-collapse">
                        <div class="panel-body">

                


            
                        <div class="row">
                        <div class="col-lg-6">

           
                  <table class="table table-stribed">
                            <thead>
                                <tr>
                                    <th>Medarbejder</th>
                                    <th style="text-align:right;">Forecast</th>
                                    <th style="text-align:right;">Realiseret</th>

                                </tr>
                            </thead>

                      <tbody>

            
                    
                    
                             <%
                             strSQLjob = "SELECT id AS jobid, jobslutdato, jobnr, jobnavn FROM job WHERE jobnr LIKE '"& jobSogVal &"%' OR jobnavn LIKE '"& jobSogVal &"%'"
                             levDato = "01-01-2002"
                             jobid = 0
                             jobnr = 0
                             jobnavn = ""
                             oRec2.open strSQLjob, oConn, 3
                             if not oRec2.EOF then
                        
                             levDato = oRec2("jobslutdato")
                             jobid = oRec2("jobid")
                             jobnr = oRec2("jobnr")
                             jobnavn = oRec2("jobnavn")

                             end if 
                             oRec2.close

                             %>
                            <tr>
                                <td colspan="3">
                                    <b><%=jobnavn & " ("& jobnr &")" %></b><br />
                                    Lev. dato: <%=levDato %></td>

                            </tr>

                             <%'*** Henter ALLE medarbejdere ****'
                             strSQLmedarb = "SELECT mid, mnavn FROM medarbejdere WHERE mid > 0 ORDER BY mnavn"
                             re = 0
                             sumTimerGT = 0
                             sumfcGT = 0
                             oRec3.open strSQLmedarb, oConn, 3
                             while not oRec3.EOF 
                        
                            
                                            strSQLrt = "SELECT tmnavn, tjobnavn, tjobnr, SUM(t.timer) AS sumtimer, tmnr FROM timer AS t "_
                                            &" WHERE (tjobnr = '"& jobnr &"') AND (tmnr = "& oRec3("mid") &") AND ("& aty_sql_realhours &") "_
                                            &" AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' GROUP BY t.tjobnr, tmnavn" 

                        
                                            realTimer = 0
                                            oRec2.open strSQLrt, oConn, 3
                                            if not oRec2.EOF then
                                            
                                            realTimer = oRec2("sumtimer")

                                            end if
                                            oRec2.close
                                 
                                 
                                             '*** INDENFOR budget�r? 
                                             fctimer = 0
                                             strSQLfc = "SELECT sum(timer) AS fctimer FROM ressourcer_md WHERE jobid = "& jobid &" AND medid = " & oRec3("mid") & " GROUP BY medid, jobid"
                                             oRec2.open strSQLfc, oConn, 3
                                             if not oRec2.EOF then
                        
                                             fctimer = oRec2("fctimer")

                                             end if 
                                             oRec2.close
                                        
                                 
                                 
                                    if realTimer <> 0 OR fctimer <> 0 then
                                     

                                    select case right(re, 1)
                                    case 0,2,4,6,8
                                    bgcr = "#FFFFFF"
                                    case else
                                    bgcr = "#EfF3ff"
                                    end select

                            
                         

                                    %>
                                    <tr style="background-color:<%=bgcr%>;">
                                       
                                        <td><%=oRec3("mnavn") %></td>
                                        <td align="right"><%=formatnumber(fctimer, 2) %> t.</td>
                                        <td align="right"><%=formatnumber(realTimer, 2) %> t.</td>
                                    </tr>
                                    <%



                                    sumTimerGT = sumTimerGT + realTimer
                                    sumfcGT = sumfcGT + fcTimer
                                    re = re + 1
                                    
                                    end if


                    oRec3.movenext
                    wend 
                    oRec3.close
                        
                     
                    if re = 0 then   
                    %>
                    <tr bgcolor="#FFFFFF"><td colspan="3">- ingen </td></tr>
                    <%else %>
                       <tr bgcolor="#FFFFFF">
                           
                           <td>&nbsp;</td>
                           <td align="right"><%=formatnumber(sumfcGT, 2) %> t.</td>
                           <td align="right"><%=formatnumber(sumTimerGT, 2) %> t.</td></tr>    

                    <%end if %>
                      </tbody>
                    </table>
                              </div><!-- coloum -->
                              </div><!-- /ROW -->


                               </div><!-- END panel-body -->
              </div><!-- END panel-collapse -->
             </div><!-- END panel deafult -->
        <%

        else




           
        '** Mtype m�l ***

        if cint(multipleMids) = 0 then
        meType = meType
        else
        meType = meType
        end if

        mTypeNavn = ""
        strSQLmtypnavn = "SELECT type FROM medarbejdertyper WHERE id =  " & meType
        oRec4.open strSQLmtypnavn, oConn, 3
        if not oRec4.EOF then

            mTypeNavn = oRec4("type")

        end if
        oRec4.close


        fo = 0
        foBA = 0
        dim fomrnavn, fomrmal, fomrid, fomrBAnavn, fomrBAmal, fomrBAid
        redim fomrnavn(0), fomrmal(0), fomrid(0), fomrBAnavn(0), fomrBAmal(0), fomrBAid(0)

        select case lto
        case "epi2017" '5 vertikaler

            fo = 5 'fo + 1
            redim preserve fomrnavn(fo), fomrmal(fo), fomrid(fo)

            dim strSQLpr_business_unit
            redim strSQLpr_business_unit(4)

            if len(trim(business_area_filter)) <> 0 then
                strSQLbusiness_unit = " AND (business_unit LIKE '"& business_area_filter &"%')"
            else
           
                strSQLbusiness_unit = " AND (business_unit LIKE 'AVI%' OR "_
                &" business_unit LIKE 'PRI%' OR "_
                &" business_unit LIKE 'PUB%' OR "_
                &" business_unit LIKE 'T%' OR "_
                &" business_unit LIKE 'ADM%' "_
                &")"
            
            end if

               strSQLsel_mmff = "SELECT fomr.navn AS fomrnavn, business_area_label, business_unit, fomr.id AS fomrid FROM fomr "_
               &" WHERE fomr.id <> 0 "& strSQLbusiness_unit &" LIMIT 100"

     

        case else

            strSQLbusiness_unit = ""

                strSQLsel_mmff = "SELECT mmff_id, mmff_mal, fomr.navn AS fomrnavn, mmff_fomr, business_area_label, business_unit, fomr.id AS fomrid FROM fomr "_
                &" LEFT JOIN mtype_mal_fordel_fomr ON (mmff_fomr = fomr.id AND mmff_mtype = "& meType & ") "_
                &" WHERE fomr.id <> 0 "& strSQLbusiness_unit &" LIMIT 30"

        end select

    
        'if session("mid") = 1 AND lto = "epi2017" then
        'response.write strSQLsel_mmff
        'response.flush
        'end if


        
              
         oRec4.open strSQLsel_mmff, oConn, 3
        while not oRec4.EOF
            
            if oRec4("business_unit") = "Salg" then 'Sekund�r Donut
            'OR (lto = "epi2017" AND instr(oRec4("business_unit"), "AVI") <> 0) 


            redim preserve fomrBAnavn(foBA), fomrBAmal(foBA), fomrBAid(foBA)
            fomrBAid(foBA) = oRec4("fomrid")
            fomrBAnavn(foBA) = oRec4("business_area_label")
            fomrBAmal(foBA) = oRec4("mmff_mal")

            'response.write "fomrBAnavn(foBA): " & fomrBAnavn(foBA) & "<br>"

            foBA = foBA + 1 
    
            else

                
               

                select case lto
                case "epi2017" '5 vertikaler

             

                    if instr(oRec4("business_unit"), "AVI") <> 0 then
                        strSQLpr_business_unit(0) = strSQLpr_business_unit(0) & " OR for_fomr = "& oRec4("fomrid")
            
                          fomrid(0) = oRec4("fomrid")
                          fomrnavn(0) = left(oRec4("business_unit"), 3)
                          fomrmal(0) = 0
                   
                    end if
                
                    if instr(oRec4("business_unit"), "PRI") <> 0 then
                        strSQLpr_business_unit(1) = strSQLpr_business_unit(1) & " OR for_fomr = "& oRec4("fomrid")   
           
                          fomrid(1) = oRec4("fomrid")
                          fomrnavn(1) = left(oRec4("business_unit"), 3)
                          fomrmal(1) = 0             
                
                    end if

                    if instr(oRec4("business_unit"), "PUB") <> 0 then
                        strSQLpr_business_unit(2) = strSQLpr_business_unit(2) & " OR for_fomr = "& oRec4("fomrid")   
            
                          fomrid(2) = oRec4("fomrid")
                          fomrnavn(2) = left(oRec4("business_unit"), 3)
                          fomrmal(2) = 0
                
                    end if

                    if left(oRec4("business_unit"), 1) = "T" then
                        strSQLpr_business_unit(3) = strSQLpr_business_unit(3) & " OR for_fomr = "& oRec4("fomrid") 

                          fomrid(3) = oRec4("fomrid")
                          fomrnavn(3) = left(oRec4("business_unit"), 3)
                          fomrmal(3) = 0
                               
                    end if

                    
                    if instr(oRec4("business_unit"), "ADM") <> 0 then
                        strSQLpr_business_unit(4) = strSQLpr_business_unit(4) & " OR for_fomr = "& oRec4("fomrid") 

                          fomrid(4) = oRec4("fomrid")
                          fomrnavn(4) = left(oRec4("business_unit"), 3)
                          fomrmal(4) = 0
                               
                    end if

                  
                    'fo = 5 'fo + 1

                case else

                redim preserve fomrnavn(fo), fomrmal(fo), fomrid(fo)

                fomrid(fo) = oRec4("fomrid")
                fomrnavn(fo) = oRec4("fomrnavn")

                    if isNull(oRec4("mmff_mal")) <> true then
                    fomrmal(fo) = oRec4("mmff_mal")
                    else
                    fomrmal(fo) = 0
                    end if    

                fo = fo + 1
                end select


               

         
            end if

        
        oRec4.movenext
        wend 
        oRec4.close



        
        '*************************************************************************************
        '*** Forretningsomr�de aktids / Finder alle akti job p� de valgte fomr
        '*************************************************************************************
        fomr_jobidsOnlySQL = " AND (jobnr = '0'"
        dim fomraktids
        redim fomraktids(fo-1)


             'if f = 0 OR lastF <> f then
         
            'end if




                    for f = 0 TO fo - 1 

                      select case lto
                        case "epi2017", "alfanordic", "cflow"
                        fomraktids(f) = "AND (tjobnr = 0"
                        useJobFO = 1
                        useAktFO = 0
                        case else 
                        fomraktids(f) = "AND (taktivitetid = 0"
                        useAktFO = 1
                        useJobFO = 0
                        end select
        
       

                        if cint(useAktFO) = 1 then '** Viser fomr knyttet til akt
                
                         strSQLsel_mmff = "SELECT for_aktid FROM fomr_rel WHERE for_fomr = "& fomrid(f)
            
                        else    '** Viser fomr knyttet til Job         
                    

                            select case lto
                            case "epi2017"
                             strSQLsel_mmff = "SELECT for_aktid, for_jobid, j.jobnr FROM fomr_rel "
                             strSQLsel_mmff = strSQLsel_mmff & " LEFT JOIN job j ON (j.id = for_jobid)"
                             strSQLsel_mmff = strSQLsel_mmff & " WHERE for_fomr = 0 "& strSQLpr_business_unit(f) 
                            case else
                             strSQLsel_mmff = "SELECT for_aktid, for_jobid, j.jobnr FROM fomr_rel "
                             strSQLsel_mmff = strSQLsel_mmff & " LEFT JOIN job j ON (j.id = for_jobid)"
                             strSQLsel_mmff = strSQLsel_mmff & " WHERE for_fomr = "& fomrid(f)
                            end select

                       end if

                    'response.write strSQLsel_mmff
                    'response.flush

                                    oRec4.open strSQLsel_mmff, oConn, 3
                                    while not oRec4.EOF
            

                                        if cint(useAktFO) = 1 then

                                            if oRec4("for_aktid") <> 0 then
                                            fomraktids(f) = fomraktids(f) & " OR taktivitetid = "& oRec4("for_aktid")
                                            end if

                                        else

                                            if oRec4("for_jobid") <> 0 AND len(trim(oRec4("jobnr"))) <> 0 then
                                             fomraktids(f) = fomraktids(f) & " OR tjobnr = '"& oRec4("jobnr") & "'"
                
                                        
                                              fomr_jobidsOnlySQL = fomr_jobidsOnlySQL & " OR jobnr = '"& oRec4("jobnr") & "'"
                                            end if

                                           

                                        end if
            
                                    oRec4.movenext
                                    wend 
                                    oRec4.close

                    fomraktids(f) = fomraktids(f) &")"

                    lastF = f
                    next



                    fomr_jobidsOnlySQL = fomr_jobidsOnlySQL  & ")"
        
            
            
            
                        '*** Forretningsomr�de aktids BA // PIE 2, kun fomr indenfor bestemt business_unit
                        dim fomraktidsBA
                        redim fomraktidsBA(foBA-1)

         

                            foBAc = 0
                            for foBAc = 0 TO foBA - 1 

           
                            'if f = 0 OR lastF <> f then
                            select case lto
                            case "epi2017"
                            fomraktidsBA(foBAc) = "AND (tjobnr = 0"
                            useJobFO = 1
                            useAktFO = 0
                            case else 
                            fomraktidsBA(foBAc) = "AND (taktivitetid = 0"
                            useAktFO = 1
                            useJobFO = 0
                            end select
                            'end if

            
                            if cint(useAktFO) = 1 then
                             strSQLsel_mmff = "SELECT for_aktid FROM fomr_rel "
                             strSQLsel_mmff = strSQLsel_mmff & " WHERE for_fomr = "& fomrBAid(foBAc)
                            else
                             strSQLsel_mmff = "SELECT for_aktid, for_jobid, j.jobnr FROM fomr_rel "
                             strSQLsel_mmff = strSQLsel_mmff & " LEFT JOIN job j ON (j.id = for_jobid)"
                             strSQLsel_mmff = strSQLsel_mmff & " WHERE for_fomr = "& fomrBAid(foBAc)
                            end if

                            'response.write strSQLsel_mmff
                            'response.flush

                            oRec4.open strSQLsel_mmff, oConn, 3
                            while not oRec4.EOF
            
                            if cint(useAktFO) = 1 then

                                if oRec4("for_aktid") <> 0 then
                                fomraktidsBA(foBAc) = fomraktidsBA(foBAc) & " OR taktivitetid = "& oRec4("for_aktid")
                                end if

                            else

                                if oRec4("for_jobid") <> 0 then
                                fomraktidsBA(foBAc) = fomraktidsBA(foBAc) & " OR tjobnr = '"& oRec4("jobnr") & "'"
                                end if

                            end if
            
                            oRec4.movenext
                            wend 
                            oRec4.close

                            fomraktidsBA(foBAc) = fomraktidsBA(foBAc) &")"

                            lastF = foBAc
                            next

        

        '*** SLUT finder jobids 7 aktid tilknyttet fomr

         if cint(multipleMids) = 0 then
            mHighEnd = 0
            antalMids = 1
         else
            mHighEnd = antalMids - 1
         end if



         dim hoursDB, hoursDBindicator, hoursAvgRate, hoursAvgindicator, utilz, hoursAvgindicato, hours, hoursOms, hoursBA
         dim utilzDB, utilzCol, utilzDBFortegn, utilzDBBA
         dim hoursDBindex, matforbrugsumPrJob, ntimPerytdprM, faktimerGTselmedarbAar
         redim utilzDB(fo-1), utilzCol(fo-1), utilzDBFortegn(fo-1), hours(fo-1), hoursOms(fo-1) 
         redim hoursDB(fo-1), hoursDBindicator(fo-1), hoursAvgRate(fo-1), hoursAvgindicato(fo-1), utilz(fo-1)
         redim hoursDBindex(fo-1), hoursBA(foBA-1), utilzDBBA(foBA-1), matforbrugsumPrJob(100000), ntimPerytdprM(100000), faktimerGTselmedarbAar(100000)

                            
              for f = 0 TO fo - 1 

                    hours(f) = 0
                    hoursDB(f) = 0
                    hoursOms(f) = 0

            
                    hoursDBindicator(f) = 0
                    hoursAvgRate(f) = 0

                    hoursAvgindicato(f) = 0

             next


         medarbSelDBGT = 0
         medarbSelOmsGT = 0
         medarbSelDBindexGT = 0
         medarbSelHoursFomrGT = 0


           ntimPerArr = 0
           indexFcArr = 0
           medarbSelHoursGTArr = 0
           faktimerGTselmedarbArr = 0
           omsGTselmedarbArr = 0
           faktimerDBGTselmedarbArr = 0



        '***********************************************************************************************
        '****  Medarbejder LOOP
        '***********************************************************************************************
        for m = 0 TO mHighEnd

        if cint(multipleMids) = 0 then
        usemrn = usemrn
        else
        usemrn = useMids(m)
        end if


        '**' Normtimer ATD
        'if year(now) = year(strDatoStartKri) then
        natalDageytd = dateDiff("d", strDatoStartKri, day(now) & "/" & month(now) & "/" & year(now), 2,2)
        call normtimerPer(usemrn, strDatoStartKri, natalDageytd, 0)
        ntimPerytd = ntimPer
        ntimPerytdprM(usemrn) = ntimPer
        'end if

        '** Normtimer
        natalDage = dateDiff("d", strDatoStartKri, strDatoEndKri, 2,2)
        call normtimerPer(usemrn, strDatoStartKri, natalDage, 0)
        ntimPer = ntimPer
        ntimPerArr = ntimPerArr + ntimPer

        'Response.write m & "usemrn: "& usemrn &" ntimPer: " & ntimPer & "multipleMids: "& multipleMids &"<br>"

        '** Target timepris
        tgtimpris = 0 '600




            '*****************************************************
            '**** RESSOURCE FORECAST timer TOTAL til INDEX ******'
            '*****************************************************
           
            if m = 0 then
            jobisMedLevDato = " AND (jobid = 0" 
            strSQLjobmLevDato = "SELECT j.id FROM job AS j WHERE jobslutdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'"
            
            'response.Write strSQLjobmLevDato
            'response.flush
            oRec5.open strSQLjobmLevDato, oConn, 3
            while not oRec5.EOF 
            
            jobisMedLevDato = jobisMedLevDato & " OR jobid =  " & oRec5("id") 

            oRec5.movenext
            wend 
            oRec5.close 

            jobisMedLevDato = jobisMedLevDato & ")"

            useEndMonth = month(strDatoEndKri)
            strSQLResPer = " AND aar = "& useYear  &"" 'I FY = AND (md >= "& useMonth &" AND md <= "& useEndMonth &")"
           
            end if

            
            indexFc = 0
            strSQLforecast = "SELECT COALESCE(SUM(timer),0) AS fctimer FROM ressourcer_md WHERE medid = "& usemrn &" "& jobisMedLevDato &" GROUP BY medid"
            'strSQLResPer
            'response.write strSQLforecast
            'response.flush
            
            oRec5.open strSQLforecast, oConn, 3

            if not oRec5.EOF then
            indexFc = oRec5("fctimer")
            end if
            oRec5.close

            if indexFc <> 0 then
            indexFc = indexFc
            indexFcArr = indexFcArr + indexFc
            else
            indexFc = 1
            indexFcArr = indexFcArr
            end if

           



        '*******************************
        '*** Sygdom og Ferie periode 
        '*******************************

        select case lto
        case "epi2017", "wilke"
        use_medarbSelNormFradragGT = 0
        case else
        use_medarbSelNormFradragGT = 1
        end select




         strSQlmtimer = "SELECT SUM(timer) AS sumtimer FROM timer AS t WHERE tmnr = "& usemrn & " AND (tfaktim = 11 OR tfaktim = 13 OR tfaktim = 14 OR tfaktim = 20 OR tfaktim = 21 OR tfaktim = 22 OR tfaktim = 25) AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'" 
         medarbSelNormFradragGT = 0
        'response.write strSQlmtimer & "<hr>"
        'response.flush
         oRec.open strSQlmtimer, oConn, 3
         if not oRec.EOF then
            medarbSelNormFradragGT = oRec("sumtimer")  
         end if
         oRec.close 


        if medarbSelNormFradragGT <> 0 then
        medarbSelNormFradragGT = medarbSelNormFradragGT
        medarbSelNormFradragGTArr = medarbSelNormFradragGTArr + medarbSelNormFradragGT
        else
        medarbSelNormFradragGT = 0
       medarbSelNormFradragGTArr = medarbSelNormFradragGTArr
        end if


        '*** Korrigerer Norm for frav�r
        if cint(use_medarbSelNormFradragGT) = 1 then
            ntimerPer = (ntimerPer - medarbSelNormFradragGT)
            ntimerPerytd = (ntimerPerytd - medarbSelNormFradragGT)
        end if 
    

        '*******************************
        '*** Alle timer periode 
        '*******************************
         strSQlmtimer = "SELECT SUM(timer) AS sumtimer FROM timer AS t WHERE tmnr = "& usemrn & " AND ("& aty_sql_realHours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'" 
         medarbSelHoursGT = 0
        'response.write strSQlmtimer & "<hr>"
        'response.flush
         oRec.open strSQlmtimer, oConn, 3
         if not oRec.EOF then
            medarbSelHoursGT = oRec("sumtimer")  
         end if
         oRec.close 


        if medarbSelHoursGT <> 0 then
        medarbSelHoursGT = medarbSelHoursGT
        medarbSelHoursGTArr = medarbSelHoursGTArr + medarbSelHoursGT
        else
        medarbSelHoursGT = 0
        medarbSelHoursGTArr = medarbSelHoursGTArr
        end if


           

        '*******************************
        '*** Fakurerbare timer periode 
        '*******************************
        omsGTselmedarb = 0
        faktimerDBGTselmedarb = 0
        faktimerGTselmedarb = 0
        
         strSQlmtimer = "SELECT tmnr, SUM(timer) AS sumtimer, SUM(timer*timepris) AS omsGTselmedarb, SUM(timer*(timepris-kostpris)) AS faktimerDBGTselmedarb FROM timer AS t WHERE tmnr = "& usemrn & " AND ("& aty_sql_realHoursFakbar &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'" 
            
        
        'response.write strSQlmtimer & "<hr>"
        'response.flush

         faktimerGTselmedarbAar(usemrn) = 0 
         oRec.open strSQlmtimer, oConn, 3
         if not oRec.EOF then

        faktimerGTselmedarb = oRec("sumtimer")  
        omsGTselmedarb = oRec("omsGTselmedarb")
        faktimerDBGTselmedarb = oRec("faktimerDBGTselmedarb")

        faktimerGTselmedarbAar(usemrn) = oRec("sumtimer")  

         end if
         oRec.close 


        if faktimerGTselmedarb <> 0 then
        faktimerGTselmedarb = faktimerGTselmedarb
        faktimerGTselmedarbArr = faktimerGTselmedarbArr + faktimerGTselmedarb
        else
        faktimerGTselmedarb = 0
        faktimerGTselmedarbArr = faktimerGTselmedarbArr
        end if


        if omsGTselmedarb <> 0 then
        omsGTselmedarb = omsGTselmedarb
        omsGTselmedarbArr = omsGTselmedarbArr + omsGTselmedarb
        else
        omsGTselmedarb = 0
        omsGTselmedarbArr = omsGTselmedarbArr
        end if

        if faktimerDBGTselmedarb <> 0 then
        faktimerDBGTselmedarb = faktimerDBGTselmedarb
        faktimerDBGTselmedarbArr = faktimerDBGTselmedarbArr + faktimerDBGTselmedarb
        else
        faktimerDBGTselmedarb = 0
        faktimerDBGTselmedarbArr = faktimerDBGTselmedarbArr
        end if



        '*********************************************************
        '** Real. timer SEL MEDARB Fordelt p� Fomr
        '*********************************************************
     
             
      


                       for f = 0 TO fo - 1 

                        '************************************************
                        '************** DE 7 - (30) MAIN FORM ***********
                        '************************************************
                        
                        ' select case lto
                        ' case "epi2017"
                         'strSQlmtimer = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tmnr = "& usemrn & " AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' "& fomraktids(f)
                        
                         'case else 
                        '& replace(fomr_jobidsOnlySQL, "jobnr", "tjobnr")
                         strSQlmtimer = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris*kurs/100) AS sumtimeOms, SUM(timer*kostpris*kpvaluta_kurs/100) AS sumtimeKost FROM timer "_
                         &" WHERE tmnr = "& usemrn & " AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' "& fomraktids(f) & " GROUP BY tmnr"
                         'end select

                         if m = 0 then
                         hoursDB(f) = 0
                         hoursDBindicator(f) = 0
                         hoursAvgRate(f) = 0   
                         hoursAvgindicato(f) = 0
                         hoursOms(f) = 0 
                         hoursDBindex(f) = 0
                         hours(f) = 0
                         end if

                        'response.write strSQlmtimer & "<hr>"
                        'response.flush

                         oRec.open strSQlmtimer, oConn, 3
                         if not oRec.EOF then

                            if isNULL(oRec("sumtimer")) <> true AND oRec("sumtimer") <> 0 then ' AND isNull(oRec("sumtimeOms")) <> true AND oRec("sumtimeOms") <> 0 then
                                
                                hours(f) = oRec("sumtimer")

                                if isNull(oRec("sumtimeOms")) <> true AND oRec("sumtimeOms") <> 0 then
                                hoursDB(f) = oRec("sumtimeOms") - oRec("sumtimeKost")
                                hoursOms(f) = oRec("sumtimeOms")

          
                        
                                hoursDBindicator(f) = (oRec("sumtimeKost") / oRec("sumtimeOms")) * 100 'dbproc
                                hoursAvgRate(f) = oRec("sumtimeOms") / oRec("sumtimer")

                              
                                if hoursAvgRate(f) > tgtimpris then
                                hoursAvgindicato(f) = 100 - ((tgtimpris / hoursAvgRate(f) * 100))
                                else
                                hoursAvgindicato(f) = (hoursAvgRate(f) / tgtimpris * 100)
                                end if

                                else


                                hoursDB(f) = 0
                                hoursOms(f) = 0

                                hoursDBindicator(f) = 0
                                hoursAvgRate(f) = 0

                                hoursAvgindicato(f) = 0
                             

                                end if

          
                            end if

            
           

                         end if
                         oRec.close 


           

            
                            medarbSelHoursFomrGT = medarbSelHoursFomrGT + hours(f)
                            'medarbSelHoursGT = medarbSelHoursGT + hours(f)
            
                            '*** 2016 06 13 Lavet om til faktimerGTselmedarb nedenfor
                            'medarbSelDBGT = medarbSelDBGT + hoursDB(f)
                            'medarbSelOmsGT = medarbSelOmsGT + hoursOms(f)

                            'response.write "medarbSelHoursGT: " & formatnumber(medarbSelHoursGT, 0) &"<br>"
            

                        next

            
                        'medarbSelDBGT = faktimerDBGTselmedarb   'DB fakturerbare timer * (timepris - kost) 
                        'medarbSelOmsGT = omsGTselmedarb         'Oms fakturerbare timer * timepris
                        'medarbSelHoursGT = medarbSelHoursGT     'ALLE Timer


                    next 'mHighEnd


            '****************************************************************************
            '**** Medarbejder LOOP SLUT
            '****************************************************************************







            '*** Knytter Gt v�rdier for alle valgte medarbejdere ved multible til den oprindelige variable
            ntimPer = ntimPerArr/antalMids
            
            
            'Response.write "ntimPerArr/antalMids" & ntimPer &" = "& ntimPerArr &"/"& antalMids

            indexFc = indexFcArr/antalMids
            medarbSelHoursGT = medarbSelHoursGTArr/antalMids
            faktimerGTselmedarb = faktimerGTselmedarbArr/antalMids
            omsGTselmedarb = omsGTselmedarbArr/antalMids
            faktimerDBGTselmedarb = faktimerDBGTselmedarbArr/antalMids


                        medarbSelDBGT = faktimerDBGTselmedarb   'DB fakturerbare timer * (timepris - kost) 
                        medarbSelOmsGT = omsGTselmedarb         'Oms fakturerbare timer * timepris
                        medarbSelHoursGT = medarbSelHoursGT     'ALLE Timer



                            '** IKKE i forhold til norm, men i forhold til reg FAKTURERBARE timer
                            for f = 0 TO fo - 1 

                            utilz(f) = 0
                            utilzDB(f) = 0
                            utilzCol(f) = ""
                            utilzDBFortegn(f) = ""

                                    utilz(f) = (hours(f) - medarbSelHoursGT)

                                    if hours(f) <> 0 then

                                    utilz(f) = hours(f)

                                        if medarbSelHoursGT > hours(f) then
                                        utilzDB(f) = (hours(f) / medarbSelHoursGT) * 100
                                        utilzCol(f) = "danger"
                                        utilzDBFortegn(f) = "-"
                                        else
                                        utilzDB(f) = (medarbSelHoursGT / hours(f)) * 100
                                        utilzCol(f) = "success"
                                        utilzDBFortegn(f) = "+"
                                        end if

                                    else

                                    utilz(f) = 0
                                    utilzDB(f) = 0
                                    utilzCol(f) = "success"
                                    utilzDBFortegn(f) = "+"

                                    end if

                            next


        


            if medarbSelOmsGT <> 0 then
            medarbSelOmsGT = formatnumber(medarbSelOmsGT, 0)
            else
            medarbSelOmsGT = 0
            end if

            if medarbSelDBGT <> 0 then
            medarbSelDBGT = formatnumber(medarbSelDBGT, 0)
            else
            medarbSelDBGT = 0
            end if

            if medarbSelHoursGT <> 0 then
            medarbSelHoursGT = formatnumber(medarbSelHoursGT, 2)
            else
            medarbSelHoursGT = 0
            end if

            medarbSelIndex = 0

            '** Fakturerings index HVIS Grundlag er forecast
            'if cdbl(indexFc) > cdbl(medarbSelHoursGT) then
            'medarbSelIndex = (1 - (medarbSelHoursGT/indexFc)) + 1
            'else
            'medarbSelIndex = (indexFc/medarbSelHoursGT)
            'end if

            '** Fakturerings index HVIS Grundlag er normtid

            if usePeriode = 3 AND year(now) = year(strDatoStartKri) then 'ATD

                 if cdbl(ntimperytd) > cdbl(faktimerGTselmedarb) then
                    if ntimperytd <> 0 AND faktimerGTselmedarb <> 0 then
                    medarbSelIndex = (faktimerGTselmedarb/ntimperytd)
                    else
                    medarbSelIndex = 0
                    end if
                else
                    if ntimperytd <> 0 AND faktimerGTselmedarb <> 0 then
                    medarbSelIndex = (ntimperytd/faktimerGTselmedarb)
                    else
                    medarbSelIndex = 0
                    end if
                end if

            else

                if cdbl(ntimper) > cdbl(faktimerGTselmedarb) then
                    if ntimper <> 0 AND faktimerGTselmedarb <> 0 then
                    medarbSelIndex = (faktimerGTselmedarb/ntimper)
                    else
                    medarbSelIndex = 0
                    end if
                else
                    if ntimper <> 0 AND faktimerGTselmedarb <> 0 then
                    medarbSelIndex = (ntimper/faktimerGTselmedarb)
                    else
                    medarbSelIndex = 0
                    end if
                end if
         
            end if


            'if (medarbSelIndex) < 1 then
            'medarbSelIndex = medarbSelIndex
            'else
            'medarbSelIndex = 0
            'end if

            medarbSelDBindexGT = (medarbSelIndex) * medarbSelDBGT


        select case lto
        case "epi2017"

        jobansInvAllClosed = 0
        salgsansInvAllClosed = 0
       
       fakbeloebThisClosed = 0
       fakbeloebThisClosedOmkost = 0

        jobansInvAllClosedMinusKost = 0
        salgsansInvAllClosedMinusKost = 0

        jobansInvAll = 0
        salgsansInvAll = 0
       fakbeloebThis = 0

            salgsansInvAllClosedPercentPrM = 0
                jobansInvAllClosedPercentPrM = 0 

        '************************************************* 
        '*** Job og salgsavrslig oms *********************
        '*************************************************

            if len(trim(business_area_filter)) <> 0 then
            fomr_jobidsOnlySQL = fomr_jobidsOnlySQL
            else
            fomr_jobidsOnlySQL = "" 
            end if

            if cint(multipleMids) = 0 then

            'KID as mid for ikke at fejle. Bare Pseydo 20171120
            strSQLjobansv = "SELECT j.id AS mid, j.id AS jid, jobnavn, jobnr, "_
            &" jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, "_
            &" salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc, jobstatus "_
            &" FROM job j "_
            &" LEFT JOIN fakturaer f ON (f.jobid = j.id)"_
            &" WHERE "_
            &" ((jobans1 = "& usemrn &" OR "_ 
            &" jobans2 = "& usemrn &" OR "_
            &" jobans3 = "& usemrn &" OR "_
            &" jobans4 = "& usemrn &" OR "_
            &" jobans5 = "& usemrn &")"_
            &" OR (salgsans1 = "& usemrn &" OR "_ 
            &" salgsans2 = "& usemrn &" OR "_
            &" salgsans3 = "& usemrn &" OR "_
            &" salgsans4 = "& usemrn &" OR "_
            &" salgsans5 = "& usemrn &")) "_
            &" AND jobstatus <> 3 AND (jobstartdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'"_
            &" OR (((brugfakdatolabel = 0 AND f.fakdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"') "_
            &" OR (brugfakdatolabel = 1 AND f.labeldato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"')) AND shadowcopy = 0) "& fomr_jobidsOnlySQL &""_
            &") GROUP BY j.id ORDER BY jobstartdato, f.labeldato DESC" 

            else 'Multibple Mids

                    jobansKrijobids = "mid <> 0 "& strSQLmids &" AND (jobstartdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'"_ 
                    &" OR (((brugfakdatolabel = 0 AND f.fakdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"') "_
                    &" OR (brugfakdatolabel = 1 AND f.labeldato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"')) AND shadowcopy = 0)) "& fomr_jobidsOnlySQL &""

                    strSQLjobansv = "SELECT mid, init, mnavn, "_
                    &" kkundenavn, kkundenr, jobnavn, jobnr, j.id AS jid, jobstartdato, jobslutdato, jo_bruttooms AS jo_bruttooms, jo_valuta, jo_valuta_kurs, "_
                    &" jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, "_ 
                    &" salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc, f.labeldato, jobstatus "
                                         
                    strSQLjobansv = strSQLjobansv &" FROM medarbejdere "_
                    &" LEFT JOIN job AS j ON (jobans1 = mid OR jobans2 = mid OR jobans3 = mid OR jobans4 = mid OR jobans5 = mid OR salgsans1 = mid OR salgsans2 = mid OR salgsans3 = mid OR salgsans4 = mid OR salgsans5 = mid) "_
                    &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr)"_
                    &" LEFT JOIN fakturaer f ON (f.jobid = j.id)"
                                         
                    strSQLjobansv = strSQLjobansv &" WHERE "& jobansKrijobids &" GROUP BY mid, jid ORDER BY mnavn"

            end if

            
            
            'jobansInv = 0
            'salgsansInv = 0
            'if session("mid") = 1 then
            'Response.write strSQLjobansv
            'Response.flush
            'end if
            fakbeloebThis = 0
            fakbeloeb = 0
            jobansInv = 0
            salgsansInv = 0   
            isJobWrt = ""
            a = 0
            antalClosed = 0
            jobansInvMinusKost = 0
            salgsansInvMinusKost = 0

            oRec4.open strSQLjobansv, oConn, 3
            while not oRec4.EOF 


            'Response.write "<br>"& oRec4("jobnavn") & " ("& oRec4("jobnr") &"): "


                            '** Faktureret bel�b ****************
                            strSQLjobansvInv = "SELECT IF(faktype = 0, COALESCE(sum(f.beloeb * (f.kurs/100)),0), COALESCE(sum(f.beloeb * -1 * (f.kurs/100)),0)) AS fakbeloeb FROM fakturaer f WHERE jobid = " & oRec4("jid") &""_
                            &" AND ((brugfakdatolabel = 0 AND f.fakdato BETWEEN '"& year(strDatoStartKri) &"-01-01' AND '"& year(strDatoEndKri) &"-12-31') "_
                            &" OR (brugfakdatolabel = 1 AND f.labeldato BETWEEN '"& year(strDatoStartKri) &"-01-01' AND '"& year(strDatoEndKri) &"-12-31')) AND shadowcopy = 0 GROUP BY jobid, faktype"
                            
                            'if session("mid") = 1 then
                            'response.write strSQLjobansvInv
                            'response.flush
                            'end if
                            'if cint(multipleMids) = 0 then
                            fakbeloeb = 0
                            'end if            

                            oRec2.open strSQLjobansvInv, oConn, 3
                            while not oRec2.EOF
                                    
                                fakbeloeb = fakbeloeb + oRec2("fakbeloeb")
                            
                            oRec2.movenext
                            wend 
                            oRec2.close 
                                     


                          '** Faktureret bel�b TOTal no periode ****************
                            strSQLjobansvInv = "SELECT IF(faktype = 0, COALESCE(sum(f.beloeb * (f.kurs/100)),0), COALESCE(sum(f.beloeb * -1 * (f.kurs/100)),0)) AS fakbeloeb FROM fakturaer f WHERE jobid = " & oRec4("jid") &" AND shadowcopy = 0 GROUP BY jobid, faktype"
                            
                            fakbeloebTot = 0
                           
                            oRec2.open strSQLjobansvInv, oConn, 3
                            while not oRec2.EOF
                                    
                                fakbeloebTot = fakbeloebTot + oRec2("fakbeloeb")
                            
                            oRec2.movenext
                            wend 
                            oRec2.close 
                               





                           '***** External cost *********
                           omkostningerIaltprJob = 0
                           matforbrugsum = 0
                           select case lto
                           case "epi2017"

                                   'if (instr(isJobWrt, ",#"& oRec4("jid") &"#") = 0 AND cint(multipleMids) = 1) OR cint(multipleMids) = 0 then                 

                                    strSQLjobansvInv = "SELECT COALESCE(sum(m.matkobspris * matantal * (m.kurs/100)),0) AS matforbrugsum FROM materiale_forbrug m"_
                                    &" WHERE m.jobid = " & oRec4("jid") &""
                                    '&" AND m.forbrugsdato BETWEEN '"& year(strDatoStartKri) &"-01-01' AND '"& year(strDatoEndKri) &"-12-31' GROUP BY m.jobid "
                           
                                    'if session("mid") = 1 then
                                    'response.write strSQLjobansvInv
                                    'response.flush
                                    'end if

                                    matforbrugsumPrJob(oRec4("jid")) = 0
                                    oRec2.open strSQLjobansvInv, oConn, 3
                                    if not oRec2.EOF then
                                    
                                        matforbrugsum = oRec2("matforbrugsum")
                                        matforbrugsumPrJob(oRec4("jid")) = oRec2("matforbrugsum")
                                        'Response.Write " matforbrugsum: "& oRec2("matforbrugsum") & "<br>"  
                            
                                     end if
                                    oRec2.close 


                                   'end if

                          



                           '*** Salgsomkostninger Internt ****'
                           salgsomKostningerIntern = 0

                                   'if (instr(isJobWrt, ",#"& oRec4("jid") &"#") = 0 AND cint(multipleMids) = 1) OR cint(multipleMids) = 0 then

                                   strSQlmtimerKost = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris*kurs/100) AS sumtimeOms, SUM(timer*kostpris*kpvaluta_kurs/100) AS sumtimeKost FROM timer "_
                                    &" WHERE tjobnr = '"& oRec4("jobnr") & "' AND ("& aty_sql_realhours &") GROUP BY tjobnr"
                                    'AND tdato BETWEEN '"& year(strDatoStartKri) &"-01-01' AND '"& year(strDatoEndKri) &"-12-31'  
                          
                                   'if oRec4("jobnr") = "21544" then
                                   ' response.write strSQlmtimerKost & "<br>"
                                   ' end if

                                    oRec2.open strSQlmtimerKost, oConn, 3
                                    if not oRec2.EOF then

                                    salgsomKostningerIntern = oRec2("sumtimeKost")

                                    stimThis = oRec2("sumtimer")

                                    end if
                                    oRec2.close 

                            
                                    if len(trim(salgsomKostningerIntern)) <> 0 then
                                    salgsomKostningerIntern = salgsomKostningerIntern
                                    else
                                    salgsomKostningerIntern = 0
                                    end if

                                    'end if' iswrt
            
                            
                                    'response.write "salgsomKostningerIntern: "& oRec4("jobnr") &": timer: "& stimThis &" kost: " & salgsomKostningerIntern & "<br>"

                                    'if cint(multipleMids) = 0 then
                                    matforbrugsumPrJob(oRec4("jid")) = matforbrugsumPrJob(oRec4("jid")) + salgsomKostningerIntern
                                    omkostningerIaltprJob = matforbrugsum + salgsomKostningerIntern
                                    'else

                                        'if instr(isJobWrt, ",#"& oRec4("jid") &"#") = 0 then
                                        'matforbrugsumPrJob(oRec4("jid")) = matforbrugsumPrJob(oRec4("jid")) + salgsomKostningerIntern
                                        'omkostningerIaltprJob = matforbrugsum + salgsomKostningerIntern
                                        'else
                                        'matforbrugsumPrJob(oRec4("jid")) = matforbrugsumPrJob(oRec4("jid"))
                                        'omkostningerIaltprJob = omkostningerIaltprJob
                                        'end if

                                                        

                                    'end if


                           end select



                                jobansInv = 0
                                salgsansInv = 0    
                                jobansInvMinusKost = 0
                                salgsansInvMinusKost = 0

                                jobansInvMinusKostClosed = 0
                                jobansInvClosed = 0

                                salgsansInvMinusKostClosed = 0
                                salgsansInvClosed = 0   

                                jobansvFundet = 0
                                jobansvFundetFak = 0
                                j = 1
                                for j = 1 to 5

                                    if isNull(oRec4("jobans"&j)) <> true then

                                        if (cdbl(oRec4("jobans"&j)) = cdbl(usemrn) AND cint(multipleMids) = 0) OR (cdbl(oRec4("jobans"&j)) = cdbl(oRec4("mid")) AND cint(multipleMids) = 1) then
                                        jobansInv = jobansInv + ((fakbeloeb) * (oRec4("jobans_proc_"& j) / 100))
                                        jobansInvMinusKost = jobansInvMinusKost + ((omkostningerIaltprJob) * (oRec4("jobans_proc_"& j) / 100))


                                        if (fakbeloebTot) <> 0 AND oRec4("jobstatus") = 0 AND cint(jobansvFundetFak) = 0 then
                                        jobansInvClosed = jobansInvClosed/1 + ((fakbeloebTot) * 1) '(oRec4("jobans_proc_"& j) / 100))
                                        jobansvFundetFak = 1
                                        else
                                        jobansInvClosed = jobansInvClosed
                                        end if
                                        
                                       
                                        if (omkostningerIaltprJob) <> 0 AND oRec4("jobstatus") = 0 AND cint(jobansvFundet) = 0 then
                                        jobansInvMinusKostClosed = jobansInvMinusKostClosed/1 + (omkostningerIaltprJob * 1)  '(oRec4("jobans_proc_"& j) / 100))
                                        'Response.write oRec4("jobnavn") & "("& oRec4("jobnr") &") jobansInvMinusKostClosed: " & jobansInvMinusKostClosed & "<br>"
                                        jobansvFundet = 1
                                        else
                                        jobansInvMinusKostClosed = jobansInvMinusKostClosed
                                        end if

                                                    

                                        end if

                                    end if      
            
                                    salgsansvFundet = 0
                                    salgsansvFundetFak = 0
                                    if isNull(oRec4("salgsans"&j)) <> true then

                                        if (cdbl(oRec4("salgsans"&j)) = cdbl(usemrn) ANd cint(multipleMids) = 0) OR (cdbl(oRec4("salgsans"&j)) = cdbl(oRec4("mid")) AND cint(multipleMids) = 1) then
                                        salgsansInv = salgsansInv + ((fakbeloeb) * (oRec4("salgsans"&j&"_proc") / 100))
                                        salgsansInvMinusKost = salgsansInvMinusKost + ((omkostningerIaltprJob) * (oRec4("salgsans"&j&"_proc") / 100))


                                        if (fakbeloebTot) <> 0 AND oRec4("jobstatus") = 0 AND cint(salgsansvFundetFak) = 0  then
                                        salgsansInvClosed = salgsansInvClosed + ((fakbeloebTot) * 1) '(oRec4("salgsans"&j&"_proc")  / 100))
                                        salgsansvFundetFak = 1
                                        else
                                        salgsansInvClosed = salgsansInvClosed
                                        end if
                                        
                                       
                                        if (omkostningerIaltprJob) <> 0 AND oRec4("jobstatus") = 0 AND cint(salgsansvFundet) = 0 then
                                        salgsansInvMinusKostClosed = salgsansInvMinusKostClosed/1 + (omkostningerIaltprJob * 1)  '(oRec4("salgsans"&j&"_proc") / 100))
                                        salgsansvFundet = 1
                                        else
                                        salgsansInvMinusKostClosed = salgsansInvMinusKostClosed
                                        end if

                                        end if

                                    end if  

                                next         


                            'if session("mid") = 1 AND cint(multipleMids) = 1 then
                            'response.write oRec4("mnavn") &"; "& oRec4("jobnavn") &"; "& oRec4("jobnr") &"; "& (jobansInv/1) & "; "& fakbeloeb &";" & omkostningerIaltprJob & ";iswrt: "& instr(isJobWrt, ",#"& oRec4("jid") &"#") &"<br>"
                            'end if


                            jobansInvAll = jobansInvAll + (jobansInv/1) '- (matforbrugsumPrJob(oRec4("jid")))/1
                            salgsansInvAll = salgsansInvAll + (salgsansInv/1) '- (matforbrugsumPrJob(oRec4("jid")))/1)
                
                            jobansInvMinusKostAll = jobansInvMinusKostAll + (jobansInvMinusKost/1)
                            salgsansInvMinusKostAll = salgsansInvMinusKostAll + (salgsansInvMinusKost/1)

                            fakbeloebThis = fakbeloebThis + (fakbeloeb/1) '- (matforbrugsumPrJob(oRec4("jid")))/1)
                            fakbeloebThisTot = fakbeloebThisTot + (fakbeloebTot/1)
                            matforbrugsumAll = matforbrugsumAll + (matforbrugsumPrJob(oRec4("jid")))/1 
                           
                            'Response.write " <div class=row><div>Jobnr: " & oRec4("jobnr") &" ("& instr(isJobWrt, ",#"& oRec4("jid") &"#") &") : <b>"& formatnumber(jobansInvAll, 2) & "</b> ### jobansInv: "& formatnumber(jobansInv, 2) & " - matforbrugsumPrJob(oRec4(jid): " & formatnumber(matforbrugsumPrJob(oRec4("jid")), 2) & " a.: "& a &"</div></div>"
                            'salgsomKostningerInternAll = salgsomKostningerInternAll + salgsomKostningerIntern/1
                            '((instr(isJobWrt, ",#"& oRec4("jid") &"#") = 0 AND cint(multipleMids) = 1) OR cint(multipleMids) = 0) AND

                            if oRec4("jobstatus") = 0 then 'LUKKET job til KPI DB beregning

                            jobansInvAllClosed = jobansInvAllClosed + (jobansInvClosed/1)
                            salgsansInvAllClosed = salgsansInvAllClosed + (salgsansInvClosed/1)

                            jobansInvAllClosedMinusKost = jobansInvAllClosedMinusKost + (jobansInvMinusKostClosed/1)
                            salgsansInvAllClosedMinusKost = salgsansInvAllClosedMinusKost + (salgsansInvMinusKostClosed/1)
                            
                            'Response.write oRec("jobnavn") & " "& oRec("jobnr") & ":  salgsansInvAllClosedMinusKost &" / " & (salgsansInvMinusKostClosed/1) & " salgsansInvAllClosed: " & salgsansInvAllClosed &" = "& salgsansInvMinusKostClosed/salgsansInvAllClosed &"<br>"

                            'fakbeloebThisClosed = fakbeloebThisClosed + (jobansInv*1+salgsansInv*1)/1 '- (matforbrugsumPrJob(oRec4("jid"))/1)
                            'fakbeloebThisClosedOmkost = fakbeloebThisClosedOmkost + (matforbrugsumPrJob(oRec4("jid"))/1)

                            antalClosed = antalClosed + 1 

                            end if

                            if cint(multipleMids) = 1 then


                                if cdbl(lastMidkpi) <> cdbl(oRec4("mid"))  then
                               

                                        antallastMidkpi = antallastMidkpi + 1

                                        if a > 0 then

                                            if jobansInvAllClosed <> 0 then
                                            jobansInvAllClosedPercentPrM = jobansInvAllClosedPercentPrM*1 + ((jobansInvAllClosed - jobansInvAllClosedMinusKost) / jobansInvAllClosed)
                                            else
                                            jobansInvAllClosedPercentPrM = jobansInvAllClosedPercentPrM
                                            end if


                                            

                                            if salgsansInvAllClosed <> 0 then
                                            salgsansInvAllClosedPercentPrM = salgsansInvAllClosedPercentPrM*1 + ((salgsansInvAllClosed - salgsansInvAllClosedMinusKost) / salgsansInvAllClosed)
                                            else
                                            salgsansInvAllClosedPercentPrM = salgsansInvAllClosedPercentPrM
                                            end if

                                            'response.write "antallastMidkpi: "& antallastMidkpi & " a: "& a & " jobansInvAllClosedPercentPrM: " & jobansInvAllClosedPercentPrM & " salgsansInvAllClosedPercentPrM : "& salgsansInvAllClosedPercentPrM &"<br>"

                                        jobansInvAllClosed = 0
                                        salgsansInvAllClosed = 0

                    
                                        jobansInvAllClosedMinusKost = 0
                                        salgsansInvAllClosedMinusKost = 0
                           
                                        end if
                                

                                end if

                                lastMidkpi = oRec4("mid")

                            else
                            antallastMidkpi = 1
                            end if

                            a = a + 1
                           isJobWrt = isJobWrt &",#"& oRec4("jid") &"#"   


            oRec4.movenext
            wend
            oRec4.close


            if cint(multipleMids) = 1 then


                if a > 0 then
                               

                if jobansInvAllClosed <> 0 then
                jobansInvAllClosedPercentPrM = jobansInvAllClosedPercentPrM*1 + ((jobansInvAllClosed - jobansInvAllClosedMinusKost) / jobansInvAllClosed)
                else
                jobansInvAllClosedPercentPrM = jobansInvAllClosedPercentPrM
                end if

                   

                if salgsansInvAllClosed <> 0 then
                salgsansInvAllClosedPercentPrM = salgsansInvAllClosedPercentPrM *1 + ((salgsansInvAllClosed - salgsansInvAllClosedMinusKost) / salgsansInvAllClosed)
                else
                salgsansInvAllClosedPercentPrM = salgsansInvAllClosedPercentPrM
                end if

                    'response.write "S antallastMidkpi: "& antallastMidkpi & " a: "& a & " jobansInvAllClosedPercentPrM: " & jobansInvAllClosedPercentPrM & " salgsansInvAllClosedPercentPrM: "& salgsansInvAllClosedPercentPrM &"<br>"

                           
                else

                salgsansInvAllClosedPercentPrM = 0
                jobansInvAllClosedPercentPrM = 0 
                
                end if

            end if




            jobansInv = formatnumber(jobansInvAll, 2)
            salgsansInv = formatnumber(salgsansInvAll, 2)

            salgsansInvMinusKost = formatnumber(salgsansInvMinusKostAll, 2)
            jobansInvMinusKost = formatnumber(jobansInvMinusKostAll, 2)

            'closed
            jobansInvClosedMinusKost = formatnumber(jobansInvAllClosedMinusKost,2)
            salgsansInvClosedMinusKost = formatnumber(salgsansInvAllClosedMinusKost,2)
         

            jobansInvClosed = formatnumber(jobansInvAllClosed, 2)
            salgsansInvClosed = formatnumber(salgsansInvAllClosed, 2)

        end select

             

        '**************************************************
        '*** PEERS
        '**************************************************


        '** Real. timer MEDARB PROJEKTGRUPPE ***'
        call hentbgrppamedarb(usemrn)
        
        '** FJERN ALLE GRUPPEN
        projektgruppeIds = replace(projektgruppeIds, ", 10", "") 

        call medarbiprojgrp(projektgruppeIds, usemrn, 0, -1)
        medarbgrpIdSQLkri = replace(medarbgrpIdSQLkri, "mid", "tmnr")


      
        visPeers = 0
        if cint(visPeers) = 1 then

        strSQlmtimerProjgrp = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris*kurs/100) AS sumtimeOms, SUM(timer*kostpris*kpvaluta_kurs/100) AS sumtimeKost, "_
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

         

            m = m + 1

         oRec.movenext
         wend 
         oRec.close 


            

            pghoursDB = pghoursDB / m-1
            pghoursDBTxt = replace(formatnumber(pghoursDB, 0), ".", "")
            
            pghoursDBindicator = pghoursDBindicator / m-1


        end if 'vis Peers

            
            f = 1
            if hoursDBindicator(f) < 20 then
            hoursDBindicatorColor = "danger"
            else
            hoursDBindicatorColor = "success"
            end if

            if hoursAvgindicato(f) < 0 then
            hoursAvgindicatoColor = "danger"
            hoursAvgindicatoFortegn = "-"
            else
            hoursAvgindicatoColor = "success"
            hoursAvgindicatoFortegn = "+"
            end if

            if cdbl(medarbSelIndex) < 1 then
            medarbSelIndexColor = "danger"
            medarbSelIndexFortegn = "-"
            else
            medarbSelIndexColor = "success"
            medarbSelIndexFortegn = "+"
            end if
            

            


            '***********************************************************************************************************************************'
            '*** Personlig DB 1 og DB 2, Fakturaindex og Timefordelign
            '***********************************************************************************************************************************'

                '*** PIE 1 ***
               %>
                 <form>
                <input type="hidden" id="timefordeling_proc_norm" value="<%=formatnumber(ntimPer, 0) %>" />
                <!--<input type="hidden" id="timefordeling_proc_fravar" value="<%=20 %>" />-->
                <%for f = 0 TO fo - 1  %>

              
                <input type="hidden" id="timefordeling_proc_<%=f %>" value="<%=replace(formatnumber(utilzDB(f), 0), ".", "") %>" />
                <input type="hidden" id="" value="<%=replace(formatnumber(utilz(f), 0), ".", "") %>" />
                <input type="hidden" id="timefordeling_navn_<%=f %>" value="<%=fomrnavn(f)%>" />
                <%next %>
                </form>

              

     


        <div class="portlet-body">
            <br />
            <div class="row">
                <div class="col-lg-11">
                    <h4 class="portlet-title"><u><%=dsb_txt_009%> 
                        <%if cint(multipleMids) = 0 then%>
                        <span style="font-size:12px; font-weight:lighter;"><%=meTxt %></span>
                        <%else %>
                        <span style="font-size:12px; font-weight:lighter;">Multible employees selected (peers)</span>
                        <%end if %>
                         </u>
                    </h4>
                </div>
            </div>


                 <div class="row">
            
                 <div class="col-lg-4">


                    
                      <div class="row-stat">
                       
                        <!--<span class="label label-succes row-stat-badge"><%=formatnumber(jo_db2_proc, 0)%> %</span> -->
                          <%select case lto
                          case "epi2017"
                          %>
                        <p class="row-stat-label">Invoiced job & sales resp. <br />Job startdate in selected period OR invoice date in FY <!--in selected period--></p>
                        <h3 class="row-stat-value"><span style="font-size:14px;">Job:</span> <%=formatnumber(jobansInv, 2)%> <span style="font-size:14px;">DKK</span></h3>
                        <br /><span style="font-size:11px; color:#999999;"> [Invoiced in period * job %]</span>

                          <br /><br />
                          <%

                             if cint(multipleMids) = 0 then

                            if jobansInvClosed <> 0 AND jobansInvClosedMinusKost <> 0 then
                            dbJobprocClosedJob = ((jobansInvClosed - jobansInvClosedMinusKost)/jobansInvClosed)*100
                            else
                            dbJobprocClosedJob = 0
                            end if
                              
                            else

                                if antallastMidkpi > 0 then
                                dbJobprocClosedJob = (jobansInvAllClosedPercentPrM*100)/antallastMidkpi
                                else
                                dbJobprocClosedJob = 0
                                end if
                
                            end if
                              
                              %>
                           <h3 class="row-stat-value"><span style="font-size:14px;">Job resp. GM Closed job:</span> <%=formatnumber(dbJobprocClosedJob, 0)%> <span style="font-size:14px;">%</span></h3> 
                          <!--(<%=jobansInvClosed & "  -" & jobansInvClosedMinusKost %>)-->
                           <br /><span style="font-size:11px; color:#999999;">[total profit / invoiced tot.]</span> <!-- Invoiced tot. / sales & external cost -->
                          <br /><br />
                       
                        <h3 class="row-stat-value"><span style="font-size:14px;">Sales:</span> <%=formatnumber(salgsansInv, 2)%> <span style="font-size:14px;">DKK</span></h3>
                        <br /><span style="font-size:11px; color:#999999;">[Invoiced in period * sales %]</span>
                          <br /><br />
                         

                          <%
                             if cint(multipleMids) = 0 then

                            if salgsansInvClosed <> 0 AND salgsansInvClosedMinusKost <> 0 then
                            dbSalgsprocClosedJob = ((salgsansInvClosed - salgsansInvClosedMinusKost)/salgsansInvClosed)*100
                            else
                            dbSalgsprocClosedJob = 0
                            end if
                              
                            else

                                if antallastMidkpi > 0 then
                                dbSalgsprocClosedJob = (salgsansInvAllClosedPercentPrM*100)/antallastMidkpi
                                else
                                dbSalgsprocClosedJob = 0
                                end if
                              
                            end if%>

                           <h3 class="row-stat-value"><span style="font-size:14px;">Sales resp. GM Closed job:</span> <%=formatnumber(dbSalgsprocClosedJob, 0)%> <span style="font-size:14px;">%</span></h3>
                             <br /><span style="font-size:11px; color:#999999;">[total profit / invoiced tot.]</span>

                          <!--<%="salgsansInvClosed: "& salgsansInvClosed &" - salgsansInvClosedMinusKost : "&salgsansInvClosedMinusKost %> HUSK kreditnotaer-->

                        

                     
                         

                          <%
                          case else
                          %> <p class="row-stat-label">GM1</p>
                         <h3 class="row-stat-value"><%=formatnumber(medarbSelDBGT, 2)%> <span style="font-size:14px;">DKK</span></h3>
                     
                          <br /><br />
                        <p class="row-stat-label">GM2</p>
                        <h3 class="row-stat-value"><%=formatnumber(medarbSelDBindexGT, 2)%> <span style="font-size:14px;">DKK</span></h3>
                          <br /> <span style="font-size:11px; color:#999999;"> [ Timer * (timepris - Kost.) * <%=dsb_txt_012 %> ]</span>
                          <%
                          end select%>
                          

                     </div> <!-- /.row-stat -->
         

                     
                 </div>
                     <div class="col-lg-3">
                         <%select case lto
                         case "epi2017"
                             dsb_txt_012 = "Utilazation"
                         case else
                             dsb_txt_012 = dsb_txt_012
                         end select%>
                         <%=dsb_txt_012 %>: 
                         <span class="label label-<%=medarbSelIndexColor%> row-stat-badge"><%=formatnumber(medarbSelIndex, 2)%></span>
                         <br /> <span style="font-size:11px; color:#999999;">[ <%=dsb_txt_014 %> / <%=dsb_txt_013 %> ]</span>
                         <br /><br />

                         <%if cint(multipleMids) = 0 then
                             avgTxt = ""
                         else
                             avgTxt = "Avg. "
                         end if%>

                         <span style="font-size:9px; color:#999999;"><%=avgTxt%></span><%=dsb_txt_014 %>: <%=formatnumber(ntimper, 2) %> t. <br />

                         <%if usePeriode = 3 AND year(now) = year(strDatoStartKri) AND lto <> "wilke" then 'ATD 
                             %>
                             <span style="font-size:9px; color:#999999;"><%=avgTxt%></span>Norm. YTD: <%=formatnumber(ntimperytd, 2) %> t.<br />
                             <%
                         end if%>

                        
                        

                         <%select case lto
                         case "epi2017" %>
                         <span style="font-size:9px; color:#999999;"><%=avgTxt%></span>Absence Hours: <%=formatnumber(medarbSelNormFradragGT, 2) %> t.<br /> 
                          <span style="font-size:9px; color:#999999;"><%=avgTxt%></span>Stated Hours: <%=formatnumber(medarbSelHoursGT, 2) %> t.<br />
                         <span style="font-size:9px; color:#999999;"><%=avgTxt%></span>Hereby invoicable: <b><%=formatnumber(faktimerGTselmedarb, 2) %> t.</b><br />
                         <%case else %>
                         <span style="font-size:9px; color:#999999;"><%=avgTxt%></span>Frav�r timer: <%=formatnumber(medarbSelNormFradragGT, 2) %> t.<br /> 
                          <span style="font-size:9px; color:#999999;"><%=avgTxt%></span>Timer: <%=formatnumber(medarbSelHoursGT, 2) %> t.<br />
                         <span style="font-size:9px; color:#999999;"><%=avgTxt%></span>Heraf fakturerbare: <b><%=formatnumber(faktimerGTselmedarb, 2) %> t.</b><br />
                         <span style="font-size:9px; color:#999999;"><%=avgTxt%></span>Forecasttimer: <%=formatnumber(indexFc, 2) %> t.<br />
                         <span style="font-size:11px; color:#999999;">Forecast p� projekter med lev. dato i valgt periode.</span>
                         <%end select %>
                         
                         <!-- <br /><br />
                        <div class="row-stat">
                        <p class="row-stat-label">NPS</p>
                        <h3 class="row-stat-value">8,2</h3>
                
                      </div> <!-- /.row-stat -->
                         

                 </div>
                                           
                         <div class="col-lg-5">
                        <%=dsb_txt_011 %> (<%=formatnumber(medarbSelHoursFomrGT, 2) %> t.)
                        
                         <% if cint(multipleMids) = 0 then%>    
                        <br />
                        <span style="color:#999999;"><%=mTypeNavn %></span>
                        <%end if %>
                       
                             <!--
                              <%select case lto 
                        case "epi2017"
                            %>
                             <span style="color:#999999;"><%=mTypeNavn %></span>
                             <%
                        case else%>
                        <span style="font-size:11px; color:#999999;"><%=dsb_txt_011 %> based on Employeetype: <b><%=mTypeNavn %></b></span>
                        <%end select %>
                             -->

                        <div id="donut-chart" class="chart-holder" style="width:95%"></div>
                        </div>

                  </div><!-- END ROW -->









             <%if cint(multipleMids) = 0 then%>
            <!-- Alle projekter med Real. timer p� for den valgte medarb. --->


            <%select case lto
              case "epi2017"
               tfTxt = "Hour distribution from donut" 
              case else
               tfTxt = "Timefordeling fra donut"
              end select%>

            <br /><br />
              <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseFive"><%=tfTxt %> - <span style="font-size:12px; font-weight:lighter;"><%=meTxt %></span></a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseFive" class="panel-collapse collapse">
                        <div class="panel-body">

                


            
                        <div class="row">
                        <div class="col-lg-12">

           
                  <table class="table table-stribed" style="width:80%;">
                            <thead>
                                <tr><th style="width:30%;">Jobnavn</th>
                                    <th>Forretningsomr�de</th>
                                    <th>Business Unit</th>
                                    <th style="text-align:right;">Realiseret</th>
                                </tr>
                            </thead>

                      <tbody>

            
                    
                    
                    <%
                        strSQLrt = "SELECT tjobnavn, tjobnr, SUM(t.timer) AS sumtimer, j2.id AS jobid FROM timer AS t "_
                        &" LEFT JOIN "_
                        &" (Select j.jobnr, j.id, j.jobnavn FROM job AS j) "_
                        &" AS j2 ON j2.jobnr = t.tjobnr "_
                        &" WHERE (tmnr = "& usemrn & ") AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' GROUP BY t.tjobnr ORDER BY sumtimer DESC LIMIT 100" 

                         '&" LEFT JOIN aktiviteter AS a ON (a.job = j.id) "_

                        'response.write strSQLrt
                        'response.flush
                        
                        sumTimerGT = 0
                        re = 0
                        oRec.open strSQLrt, oConn, 3
                        While not oRec.EOF 

                        select case right(re, 1)
                        case 0,2,4,6,8
                        bgcr = "#FFFFFF"
                        case else
                        bgcr = "#EfF3ff"
                        end select



                                strSQlfomr = "SELECT for_fomr, for_jobid, f2.fomr AS fomr, f2.business_unit FROM fomr_rel AS frel "_
                                &" LEFT JOIN "_
                                &" (Select f.navn AS fomr, f.business_unit, f.id FROM fomr AS f) "_
                                &" AS f2 ON f2.id = frel.for_fomr "_
                                &" WHERE for_jobid = "& oRec("jobid") & " AND for_fomr IS NOT NULL "

                                'response.write strSQlfomr
                                'response.flush
                                fomrNavn = ""
                                business_unit = ""
                                oRec2.open strSQlfomr, oConn, 3
                                if not oRec2.EOF then
                                
                                fomrNavn = oRec2("fomr")
                                business_unit = oRec2("business_unit")

                                end if
                                oRec2.close

                        %>
                        <tr style="background-color:<%=bgcr%>;">
                            <td style="white-space:nowrap;"><%=left(oRec("tjobnavn"), 40) & " ("& oRec("tjobnr") &")" %></td>
                            <td><%=fomrNavn %></td>
                            <td><%=business_unit %></td>
                            <td align="right"><%=formatnumber(oRec("sumtimer"), 2) %> t.</td>
                        </tr>
                        <%

                        sumTimerGT = sumTimerGT + oRec("sumtimer")

                        re = re + 1
                        oRec.movenext
                        wend
                        oRec.close      
                        
                     
                    if re = 0 then   
                    %>
                    <tr bgcolor="#FFFFFF"><td colspan="3">- ingen </td></tr>
                    <%else %>
                       <tr bgcolor="#FFFFFF"><td colspan="4" align="right"><%=formatnumber(sumTimerGT, 2) %> t.</td></tr>    

                    <%end if %>
                      </tbody>
                    </table>
                              </div><!-- coloum -->
                              </div><!-- /ROW -->


                               </div><!-- END panel-body -->
              </div><!-- END panel-collapse -->
             </div><!-- END panel deafult -->





           



            <!-- Alle projekter med forecast / budet p� --->

            <%select case lto
             case "epi2017"
                
             case else %>
             
              <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseFour">Dine job med forecast p� - <span style="font-size:12px; font-weight:lighter;"><%=meTxt %></span></a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseFour" class="panel-collapse collapse">
                        <div class="panel-body">

                


            
                        <div class="row">
                        <div class="col-lg-6">

           
                  <table class="table table-stribed">
                            <thead>
                                <tr><th>Jobnavn</th>
                                    <th style="text-align:right;">Forecast</th>
                                    <th style="text-align:right;">Realiseret</th>
                                    <th style="text-align:right;">Lev. dato</th>
                                </tr>
                            </thead>

                      <tbody>

            
                    
                    
                    <%
                    select case lto
                    case "wilke"
                        sqlMineJobOrderBy = "jobslutdato DESC"
                     case else
                        sqlMineJobOrderBy = "jobnavn"
                    end select    
                        
                    strSQLr = "SELECT r.jobid, SUM(r.timer) AS restimer, j.jobnavn, j.jobnr, j.jobslutdato FROM ressourcer_md AS r "_
                    &" LEFT JOIN job AS j ON (j.id = r.jobid) "_
                    &" WHERE r.medid = "& usemrn & " AND r.aktid <> 0 AND jobstatus = 1 GROUP BY r.jobid ORDER BY "& sqlMineJobOrderBy &" LIMIT 150" 

                        'response.write strSQLr
                        'response.flush
                        re = 0
                        oRec.open strSQLr, oConn, 3
                        While not oRec.EOF 

                        select case right(re, 1)
                        case 0,2,4,6,8
                        bgcr = "#FFFFFF"
                        case else
                        bgcr = "#EfF3ff"
                        end select
                        %>
                        <tr style="background-color:<%=bgcr%>;">
                            <td class="lille"><%=left(oRec("jobnavn"), 20) & " ("& oRec("jobnr") &")" %></td>
                            <td class="lille" align="right"><%=oRec("restimer") %> t.</td>

                            <%
                               sumRealtimer = 0
                               strSQLtimer = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tjobnr = '"& oRec("jobnr") &"' AND tmnr = "& usemrn & " GROUP BY tmnr"
                               oRec2.open strSQLtimer, oConn, 3
                               if not oRec2.EOF then

                                sumRealtimer = oRec2("sumtimer") 

                               end if
                               oRec2.close

                                if len(trim(sumRealtimer)) <> 0 then
                                sumRealtimer = formatnumber(sumRealtimer, 2)
                                else
                                sumRealtimer = formatnumber(0, 2)
                                end if

                                 %>

                            <td class="lille" align="right"><%=sumRealtimer%> t.</td>
                            <td class="lille" align="right"><%=oRec("jobslutdato")%></td>
                        </tr>
                        <%

                        re = re + 1
                        oRec.movenext
                        wend
                        oRec.close      
                        
                     
                    if re = 0 then   
                    %>
                    <tr bgcolor="#FFFFFF"><td colspan="3">- ingen </td></tr>
                    <%end if %>
                      </tbody>
                    </table>
                              </div><!-- coloum -->
                              </div><!-- /ROW -->


                               </div><!-- END panel-body -->
              </div><!-- END panel-collapse -->
             </div><!-- END panel deafult -->
            <%end select %>






            <!-- KUN S�LGERE -->

            <%
               select case lto
               case "epi2017"

               case else


              erSaelgerOk = 0  
              strSQLprogrp = "SELECT medarbejderid FROM progrupperelationer WHERE projektgruppeid = 2 OR projektgruppeid = 12"
                
              oRec5.open strSQLprogrp, oConn, 3
              if not oRec5.EOF then
              erSaelgerOk = 1  
              end if
                
              oRec5.close%>

            <%if erSaelgerOk = 1 then %>
            
              <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne">Kunder top 10 - <span style="font-size:12px; font-weight:lighter;"><%=meTxt %></span></a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseOne" class="panel-collapse collapse">
                        <div class="panel-body">

                


            
                        <div class="row">
                        <div class="col-lg-8">
                     
                        <table class="table table-stribed">
                            <thead>
                                <tr>
                                    <th style="width: 50%">Kunder (kundenr)</th>
                                    <th style="width: 20%">Forecast</th>

                                    <th style="width: 20%">Realiseret</th>
                                    <th style="width: 10%"><!--Indikator--></th>

                                     <%for foBAc = 0 TO foBA - 1%>
                                    <th><%=fomrBAnavn(foBAc)%></th>

                                    <%next%>

                                </tr>
                            </thead>
                            <tbody>
                                <%
                                 '** Real. timer SEL MEDARB
                                 strSQlmtimer_k = "SELECT SUM(t.timer) AS sumtimer, SUM(t.timer*t.timepris*kurs/100) AS sumtimeOms, SUM(t.timer*t.kostpris*kpvaluta_kurs/100) AS sumtimeKost, j.id AS jid, "_
                                 &" t.tknavn, t.tknr, t.tjobnavn, t.tjobnr, k.kkundenr "_
                                 &" FROM timer AS t "_
                                 &" LEFT JOIN job AS j ON (jobnr = tjobnr)"_ 
                                 &" LEFT JOIN kunder AS k ON (k.kid = t.tknr) "_
                                 &" WHERE t.tmnr = "& usemrn & " AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' GROUP BY tknr ORDER BY sumtimer DESC LIMIT 10"
            
                                'if session("mid") = 1 then
                                'response.write strSQlmtimer_k
                                'response.flush
                                'end if
                                
                                 oRec.open strSQlmtimer_k, oConn, 3
                                 while not oRec.EOF 

                                    restimerThisKunde = 0 
                                    strSQLresTimer = "SELECT SUM(r.timer) AS fctimer, r.medid, r.aktid, r.aar, r.md FROM ressourcer_md AS r WHERE r.jobid = "& oRec("jid") &" AND medid = "& usemrn 

                                    oRec2.open strSQLresTimer, oConn, 3
                                    if not oRec2.EOF then 

                                    if IsNull(oRec2("fctimer")) <> true then
                                    restimerThisKunde = oRec2("fctimer")
                                    else
                                    restimerThisKunde = 0
                                    end if

                                    end if
                                    oRec2.close

                                    if cint(restimerThisKunde) <> 0 then
                                    fctimer = restimerThisKunde
                                    else
                                    fctimer = 0
                                    end if

                                    if cdbl(fctimer) > cdbl(oRec("sumtimer")) then
                                    indicatorCol = "yellowgreen"
                                    else
                                    indicatorCol = "crimson"
                                    end if
                                    %>

                                <tr>
                                    <td><%=oRec("tknavn") %> (<%=oRec("kkundenr") %>)</td>
                                    <td class="text-right;"><%=formatnumber(fctimer, 2) %></td>
                                    <td class="text-right;"><%=formatnumber(oRec("sumtimer"), 2) %></td>
                                    <td><div style="background-color:<%=indicatorCol%>; width:10px; height:10px;">&nbsp;</div></td>

                                    <%for foBAc = 0 TO foBA - 1
                                        
                                        '*****************************************************************************
                                        '************** SALGS FORM Business Area = 'SALG' fordelt p� kunde ***********
                                         strSQlmtimer_fomr = "SELECT SUM(timer) AS sumtimer FROM timer "_
                                         &" WHERE tmnr = "& usemrn & " AND ("& aty_sql_realhours &") AND tdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"' "& fomraktidsBA(foBAc) & " AND tknr = "& oRec("tknr") &" GROUP BY tknr"
            
                                         
                                        'response.write strSQlmtimer_fomr & "<hr>"
                                        'response.flush
                                        thisTimerFomrBA = 0
                                         oRec5.open strSQlmtimer_fomr, oConn, 3
                                         if not oRec5.EOF then

                                            if isNULL(oRec5("sumtimer")) <> true then
                                            hoursBA(foBAc) = oRec5("sumtimer")
                                            else
                                            hoursBA(foBAc) = 0
                                            end if
                                            
                                           if isNULL(oRec5("sumtimer")) <> true then
                                            thisTimerFomrBA = formatnumber(oRec5("sumtimer"), 0)
                                           else
                                            thisTimerFomrBA = 0
                                           end if

                                         end if
                                         oRec5.close 

                                            %>
                                            <td style="text-align:right;"><%=thisTimerFomrBA %></td>

                                            <%

                                         medarbSelHoursBAGT = (medarbSelHoursBAGT*1 + hoursBA(foBAc)*1)
           

                                     next

                                   
                                        'medarbSelHoursBAGT = replace(formatnumber(medarbSelHoursBAGT,0), ".", "")
                                                

                                        
                                        %>
                                   
                                </tr>

                                <%
                                 oRec.movenext
                                 wend
                                 oRec.close 
                                    
                                    
                                 for foBAc = 0 TO foBA - 1

                                            'response.write "medarbSelHoursGT: " & medarbSelHoursGT & " hoursBA(foBA): "& hoursBA(foBA) &"#<br>"

                                            if medarbSelHoursGT > hoursBA(foBAc) then
                                                 if medarbSelHoursGT <> 0 then
                                                 utilzDBBA(foBAc) = (hoursBA(foBAc) / medarbSelHoursGT)*100
                                                 else
                                                 utilzDBBA(foBAc) = 0
                                                 end if
                                            else
                                                if hoursBA(foBAc) <> 0 then
                                                utilzDBBA(foBAc) = (hoursBA(foBAc) / medarbSelHoursGT) * 100
                                                else
                                                utilzDBBA(foBAc) = 0
                                                end if
                                            end if
                                        
                                             

                                  next   
                                    
                                    
                                 %>


                               
                            </tbody>
                        </table>
                    </div>

                            <%
                 '*** PIE 2 ***
               %>
                 <form>
               
                <!--<input type="hidden" id="timefordeling_proc_fravar" value="<%=20 %>" />-->
                <%for foBAc = 0 TO foBA - 1  %>

              
                <input type="hidden" id="timefordeling_proc_saelger_<%=foBAc %>" value="<%=replace(formatnumber(utilzDBBA(foBAc), 0), ".", "") %>" />
                <input type="hidden" id="timefordeling_navn_saelger_<%=foBAc %>" value="<%=fomrBAnavn(foBAc)%>" />
                <%next %>
                </form>


                           
                    <div class="col-lg-4">
                    <!--Tidsfordeling Fomr. - Salg %-->
                    <div id="xdonut-chart_saelger" class="chart-holder" style="width: 95%;"></div>
                    </div>


                    </div>
                    </div><!-- END panel-body -->
              </div><!-- END panel-collapse -->
             </div><!-- END panel deafult -->

            <%end if %>

            <%end select 'lto %>


            <%end if 'Multible select medarb %>







           <%if cint(multipleMids) = 0 OR lto = "epi2017" then%>

            <%if cint(multipleMids) <> 0 then %>
            <br /><br />
            <%end if %>



            <%select case lto
             case "intranet - local", "epi2017"
                 if cint(multipleMids) = 0 then
                 antalLoopsJobansv = 2
                 else
                 antalLoopsJobansv = 1
                 end if

             case else 
                antalLoopsJobansv = 1
             end select%>

               <!-- JoB HVIS jobsansvarlig -->

                <%


                for anloops = 0 TO antalLoopsJobansv - 1

                 jobansFundet = 0
                 
                    
                 select case lto
                 case "epi2017"

                 if anloops = 0 then
                 jobansOskift = "Job & Sales Responsible % <span style=""font-size:12px; font-weight:lighter;"">(Job startdate in selected period OR invoice date in FY - Sales cost in period)</span>"
                 else
                 jobansOskift = "Job & Sales Responsible CLOSED job % <span style=""font-size:12px; font-weight:lighter;"">(Job startdate in selected period OR invoice date in FY - Total sales cost no period)</span>"
                 end if
                
                    if cint(multipleMids) = 0 then

                   
                    jobansKrijobids = " ((jobans1 = "& usemrn &" OR "_ 
                    &" jobans2 = "& usemrn &" OR "_
                    &" jobans3 = "& usemrn &" OR "_
                    &" jobans4 = "& usemrn &" OR "_
                    &" jobans5 = "& usemrn &") OR "_
                    &" (salgsans1 = "& usemrn &" OR "_ 
                    &" salgsans2 = "& usemrn &" OR "_
                    &" salgsans3 = "& usemrn &" OR "_
                    &" salgsans4 = "& usemrn &" OR "_
                    &" salgsans5 = "& usemrn &"))"

                    if anloops = 0 then
                     jobansKrijobids = jobansKrijobids & " AND jobstatus <> 3 " 
                    else
                     jobansKrijobids = jobansKrijobids & " AND jobstatus = 0 " 
                    end if

                    jobansKrijobids = jobansKrijobids & " AND (jobstartdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'"_ 
                    &" OR (((brugfakdatolabel = 0 AND f.fakdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"') "_
                    &" OR (brugfakdatolabel = 1 AND f.labeldato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"')) AND shadowcopy = 0)) "& fomr_jobidsOnlySQL &""

                    else

                    jobansKrijobids = ""

                   
                    jobansKrijobids = "mid <> 0 "& strSQLmids &" AND jobstatus <> 3 AND (jobstartdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"'"_ 
                    &" OR (((brugfakdatolabel = 0 AND f.fakdato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"') "_
                    &" OR (brugfakdatolabel = 1 AND f.labeldato BETWEEN '"& strDatoStartKri &"' AND '"& strDatoEndKri &"')) AND shadowcopy = 0)) "& fomr_jobidsOnlySQL &""

                    end if
                 
                    jobansFundet = 1
                    jobsalgsansBgtTotGT = 0
                    jobsalgsansInvTotGT = 0
                    fakbeloebGT = 0
                    toBeInvoicedGT = 0
                    toBeInvoiced = 0

                 case else   

                 jobansKrijobids = " AND (tjobnr = '0' "  
                 jobansOskift = "Projektansvarlig"
                 strSQKjobans = "SELECT jobnr FROM job WHERE (jobans1 = "& usemrn &" OR jobans2 = "& usemrn &") AND jobslutdato >= '"& strDatoStartKri &"'"
                    
                       'response.write strSQKjobans
                    'response.flush 
                    
                     oRec.open strSQKjobans, oConn, 3
                     while not oRec.EOF 
                 
                       
                        jobansKrijobids = jobansKrijobids & " OR tjobnr = '" & oRec("jobnr") & "'"

                        jobansFundet = 1
                     oRec.movenext
                     wend
                     oRec.close 

                    jobansKrijobids = jobansKrijobids & ")"   
                     
                 end select

                 
                 
                    
                    


                 if cint(jobansFundet) = 1 then

                 jobansBgtGT = 0
                 salgsansBgtGT = 0
                 matforbrugsumTot = 0
                 dbGT = 0

                 %>

             
                   <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                               <% if cint(multipleMids) = 0 then %>
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTwo<%=anloops %>"><%=jobansOskift %></a> 
                              <%else %>
                                 <%=jobansOskift %>
                              <%end if %>
                          </h4>
                        </div> <!-- /.panel-heading -->
                       <% if cint(multipleMids) = 0 then %>
                        <div id="collapseTwo<%=anloops %>" class="panel-collapse collapse">
                        <%else %>
                        <div id="collapseTwo<%=anloops %>" class="panel-collapse">
                        <%end if %>
                        <div class="panel-body">

            

                 <div class="row">
            
                 <div class="col-lg-12">
                   
                     <table class="table table-stribed">
                            <thead>
                                <tr>

                                    <%select case lto
                                     case "epi2017" 
                                        
                                        
                                     if cint(multipleMids) = 0 then %>
                                    <th style="width: 20%">Customer</th>
                                    <th style="width: 20%">Job</th>
                                    <th style="width: 12%">Proj. period</th>
                                   
                                    <% if anloops = 0 then 
                                    Budget_GTxt = "Budget"
                                    Budget_GTxtExp = "Budget"

                                    %>
                                    
                                        <th style="width: 13%"><%=Budget_GTxt %></th>
                                        <th style="width: 10%">Invoiced Total <%=useYear %></th>
                                        <th style="width: 10%">Sales & External Cost <%=useYear %></th>
                                    
                                    <%

                                    else
                                    Budget_GTxt = "GM"  
                                    Budget_GTxtExp = "GM;GM %"    
                                        %>
                                        
                                        
                                        <th style="width: 10%">Invoiced Total</th>
                                        <th style="width: 10%">Sales & External Cost</th>
                                        <th style="width: 13%; text-align:right;"><%=Budget_GTxt %></th>

                                        <%

                                    end if%>

                                  

                                     
                                    <!--<th style="width: 5%">Job resp.</th>-->
                                    <th style="width: 10%; text-align:right;">Job %</th>
                                    <!--<th style="width: 5%">Sales resp.</th>-->
                                    <th style="width: 8%; text-align:right;">Sales %</th>
                                    <!--<th style="width: 13%">Budget %</th>-->
                                    
                                    <!--<th style="width: 10%">To Be Invoiced</th>-->
                                    <th style="width: 2%"><!--Indikator--></th>
                                    <%
                                    
                                        strEkspHeader = "Init;Customer;Customer No;Job;Job No.;Status;Startdate;Delevery date; "& Budget_GTxtExp &"; Invoiced Total;Latest Invoicedate; Sales & External Cost.; Job %; Job % value;Sales %;Sales % value;xx99123sy#z"
                                        'To Be invoiced
                                    else 'TEAMLEDER matrix

                                        %>
                                    <th style="width: 20%;">Employee</th>
                                    <!--<th style="width: 15%; text-align:right;">Invoiced Total</th>-->
                                    <!--<th style="width: 15%; text-align:right;">Sales & External Cost</th>-->
                                    <th style="width: 15%; text-align:right;">Job resp. % <br /><span style="font-size:9px; color:#999999; font-weight:lighter;">[Inv. in per. job * job %]</span></th>
                                    <th style="width: 15%; text-align:right;">Job % GM<br /> <span style="font-size:9px; color:#999999; font-weight:lighter;">[Inv. in per. closed job / cost]</span></th>
                                    <th style="width: 15%; text-align:right;">Sales resp. % <br /><span style="font-size:9px; color:#999999; font-weight:lighter;">[Inv. in per. job * sales %]</span></th>
                                    <th style="width: 15%; text-align:right;">Sales % GM <br /> <span style="font-size:9px; color:#999999; font-weight:lighter;">[Inv. in per. closed job / cost]</span></th>
                                   
                                    
                                <th style="width: 5%; text-align:right;">Util.</th>
                                 

                                        <%

                                      strEkspHeader = "Employee; Job %; Job % DB closed; Sales DB %;Sales % DB closed;Util.;xx99123sy#z"


                                    end if

                                         
                                    case else %>

                                     <th style="width: 25%">Kunder (kundenr)</th>
                                    <th style="width: 25%">Job</th>
                                    <th style="width: 10%">Lev. dato</th>
                                    <th style="width: 10%">Forecast</th>
                                    <th style="width: 10%">Realiseret</th>
                                    <th style="width: 10%">Oms�tning</th>
                                    <th style="width: 10%"><!--Indikator--></th>
                                    <%end select %>
                                </tr>
                            </thead>
                            <tbody>
                                <%

                                    select case lto
                                     case "epi2017"
                                       
                                         '** Job SEL MEDARB JOB+salgsANSVARLIG

                                          if cint(multipleMids) = 0 then

                                         strSQlmtimer_k = "SELECT kid AS mid, "_
                                         &" kkundenavn, kkundenr, jobnavn, jobnr, j.id AS jid, jobstartdato, jobslutdato, jo_bruttooms AS jo_bruttooms, jo_valuta, jo_valuta_kurs, "_
                                         &" jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, "_ 
                                         &" salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc, f.labeldato, jobstatus "
                                         
                                         strSQlmtimer_k = strSQlmtimer_k &" FROM job AS j "_
                                         &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                                         &" LEFT JOIN fakturaer f ON (f.jobid = j.id)"
                                         
                                         strSQlmtimer_k = strSQlmtimer_k &" WHERE "& jobansKrijobids &" GROUP BY j.id ORDER BY jobstartdato, f.labeldato DESC"
            
                                          else 'MATRIX

                                        
                                             strSQlmtimer_k = "SELECT mid, init, mnavn, "_
                                             &" kkundenavn, kkundenr, jobnavn, jobnr, j.id AS jid, jobstartdato, jobslutdato, jo_bruttooms AS jo_bruttooms, jo_valuta, jo_valuta_kurs, "_
                                             &" jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, "_ 
                                             &" salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc, f.labeldato, jobstatus "
                                         
                                             strSQlmtimer_k = strSQlmtimer_k &" FROM medarbejdere "_
                                             &" LEFT JOIN job AS j ON (jobans1 = mid OR jobans2 = mid OR jobans3 = mid OR jobans4 = mid OR jobans5 = mid OR salgsans1 = mid OR salgsans2 = mid OR salgsans3 = mid OR salgsans4 = mid OR salgsans5 = mid) "_
                                             &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr)"_
                                             &" LEFT JOIN fakturaer f ON (f.jobid = j.id)"
                                         
                                             strSQlmtimer_k = strSQlmtimer_k &" WHERE "& jobansKrijobids &" GROUP BY mid, j.id ORDER BY mnavn"


                                          end if

                                    
                                     case else

                                         '** Real. timer SEL MEDARB JOBANSVARLIG
                                         strSQlmtimer_k = "SELECT SUM(t.timer) AS sumtimer, SUM(t.timer*t.timepris*kurs/100) AS sumtimeOms, SUM(t.timer*t.kostpris*t.kpvaluta_kurs/100) AS sumtimeKost, "_
                                         &" t.tknavn, t.tknr, t.tjobnavn, t.tjobnr, j.id AS jid, jobslutdato, k.kkundenr, jo_bruttooms, "_
                                         &" jobans1, jobans2, jobans3, jobans4, jobans5, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, "_ 
                                         &" salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc "_
                                         &" FROM timer AS t "_
                                         &" LEFT JOIN job AS j ON (j.jobnr = tjobnr) "_
                                         &" LEFT JOIN kunder AS k ON (k.kid = t.tknr) "_
                                         &" WHERE tjobnr <> 0 "& jobansKrijobids &" AND ("& aty_sql_realhours &") GROUP BY tjobnr ORDER BY jobslutdato DESC, t.tknr, tjobnavn DESC LIMIT 30"
            
                              
                                    
                                     end select 
                                 
                                  'response.write strSQlmtimer_k
                                  'response.flush

                                 
                                lastMid = 0
                                antalM = 0

                                lastMfakbeloeb = 0
                                lastMmedarbSalgsOkost = 0
                                lastMjobansBgt = 0
                                lastMsalgsansBgt = 0
                                lastMdbthis = 0 

                                medarbSalgsOkost = 0
                                antalMLastMid = 0
                                antalMLastMidClosed = 0
                                antalMLastMidClosedGt = 0

                                lastMjobsansProc = 0
                                lastMsalgssansProc = 0

                                'lastMmedarbSalgsOkostProc = 0
                                medarbJobOkostProc = 0
                                medarbSalgsOkostProc = 0 

                                 oRec.open strSQlmtimer_k, oConn, 3
                                 while not oRec.EOF 

                                    
                                    select case lto
                                    case "epi2017"


                                            
                                           
                                            


                                            strSQLjobansvInv = "SELECT IF(faktype = 0, COALESCE(sum(f.beloeb * (f.kurs/100)),0), COALESCE(sum(f.beloeb * -1 * (f.kurs/100)),0)) AS fakbeloeb FROM fakturaer f WHERE jobid = " & oRec("jid") &""

                                            if anloops = 0 then
                                            strSQLjobansvInv = strSQLjobansvInv &" AND ((brugfakdatolabel = 0 AND f.fakdato BETWEEN '"& year(strDatoStartKri) &"-01-01' AND '"& year(strDatoEndKri) &"-12-31') "
                                            strSQLjobansvInv = strSQLjobansvInv &" OR (brugfakdatolabel = 1 AND f.labeldato BETWEEN '"& year(strDatoStartKri) &"-01-01' AND '"& year(strDatoEndKri) &"-12-31')) AND shadowcopy = 0"
                                            else
                                            strSQLjobansvInv = strSQLjobansvInv 
                                            end if

                                            
                                            if cint(multipleMids) = 0 then
                                            strSQLjobansvInv = strSQLjobansvInv  &" GROUP BY jobid, faktype"
                                            else
                                            strSQLjobansvInv = strSQLjobansvInv  &" GROUP BY jobid, faktype"
                                            end if
                                            
                                            'if cint(multipleMids) = 0 OR (cdbl(lastMid) <> cdbl(oRec("mid")) AND antalM > 0) then
                                            fakbeloeb = 0
                                            'end if

                                            'if session("mid") = 1 then
                                            'response.write "<br>"& strSQLjobansvInv
                                            'response.flush
                                            'end if

                                            oRec2.open strSQLjobansvInv, oConn, 3
                                            while not oRec2.EOF 
                                    
                                            fakbeloeb = fakbeloeb + (oRec2("fakbeloeb"))
                                        
                                            oRec2.movenext
                                            wend
                                            oRec2.close 



                                            '**** BruttoOms i valuta
                                            call valutakode_fn(oRec("jo_valuta")) 

                                            jo_bruttooms = oRec("jo_bruttooms")

                                            if cint(oRec("jo_valuta")) <> cint(basisValId) then
                                            call beregnValuta(oRec("jo_bruttooms"),oRec("jo_valuta_kurs"),100)
                                            jo_bruttooms = valBelobBeregnet
                                            end if

                                     
                                            '*** Salgsomkotninger ***'
                                           'if cint(multipleMids) = 0 OR (cdbl(lastMid) <> cdbl(oRec("mid")) AND antalM > 0) then
                                            medarbSalgsOkost = 0
                                            
                                            jobansBgt = 0
                                            jobansInv = 0
                                            salgsansBgt = 0
                                            salgsansInv = 0
                                            salgsansProc = 0
                                            jobsansProc = 0
                                            'end if
                                            

                                           
                                    
                                            '***** External cost *********
                                          
                                                   select case lto
                                                   case "epi2017"
                                                            
                                                            omkostningerIaltprJob = 0
                                                            matforbrugsum = 0
                                                            strSQLjobansvInv = "SELECT COALESCE(sum(m.matkobspris * matantal * (m.kurs/100)),0) AS matforbrugsum FROM materiale_forbrug m WHERE "

                                                            if anloops = 0 then
                                                            strSQLjobansvInv = strSQLjobansvInv &" m.forbrugsdato BETWEEN '"& year(strDatoStartKri) &"-01-01' AND '"& year(strDatoEndKri) &"-12-31' AND "
                                                            end if

                                                            strSQLjobansvInv = strSQLjobansvInv &" m.jobid = " & oRec("jid") & " GROUP BY m.jobid"
                           
                                                    
                                                            oRec2.open strSQLjobansvInv, oConn, 3
                                                            if not oRec2.EOF then
                                    
                                                                matforbrugsum = oRec2("matforbrugsum")
                                                             
                                                             end if
                                                            oRec2.close 


                                        

                                                           '*** Salgsomkostninger Internt ****'
                                                           salgsomKostningerIntern = 0

                                                           strSQlmtimerKost = "SELECT SUM(timer) AS sumtimer, SUM(timer*timepris*kurs/100) AS sumtimeOms, SUM(timer*kostpris*kpvaluta_kurs/100) AS sumtimeKost FROM timer WHERE "
                                                
                                                           if anloops = 0 then
                                                           strSQlmtimerKost = strSQlmtimerKost & " tdato BETWEEN '"& year(strDatoStartKri) &"-01-01' AND '"& year(strDatoEndKri) &"-12-31' AND " 
                                                           end if

                                                           strSQlmtimerKost = strSQlmtimerKost &" tjobnr = '"& oRec("jobnr") & "' AND ("& aty_sql_realhours &") GROUP BY tjobnr" 

                                                           
                         
                                                            oRec2.open strSQlmtimerKost, oConn, 3
                                                            if not oRec2.EOF then

                                                            salgsomKostningerIntern = oRec2("sumtimeKost")

                                                            'stimThis = oRec2("sumtimer")

                                                            end if
                                                            oRec2.close 

                            
                                                         
                                                            omkostningerIaltprJob = matforbrugsum + salgsomKostningerIntern
                                                            medarbSalgsOkost = omkostningerIaltprJob


                                                   end select



                                            'end if 'anloops


                                            
                                            
                                            jobansInvMinusKost = 0
                                            salgsansInvMinusKost = 0
                                           
                                            medarbJobOkostProc = 0
                                            medarbSalgsOkostProc = 0
                                    
                                            jobansvFundet = 0
                                            jobansvFundetFak = 0
                                            
                                            j = 1
                                            for j = 1 to 5
                    
                                                if isNull(oRec("jobans"&j)) <> true then

                                                    if (cdbl(oRec("jobans"&j)) = cdbl(usemrn) AND cint(multipleMids) = 0) OR (cdbl(oRec("jobans"&j)) = cdbl(oRec("mid")) AND cint(multipleMids) = 1) then
                                                    jobansBgt = jobansBgt + (jo_bruttooms * (oRec("jobans_proc_"&j) / 100)) 'jo_bruttooms
                                                    jobansInv = jobansInv + ((fakbeloeb * oRec("jobans_proc_"&j)) / 100)

                                                    if (fakbeloeb) <> 0 AND oRec("jobstatus") = 0 AND cint(jobansvFundetFak) = 0 then
                                                    jobansInvMinusKost = jobansInvMinusKost/1 + (fakbeloeb * 1) '(oRec("jobans_proc_"& j) / 100))
                                                    jobansvFundetFak = 1
                                                    else
                                                    jobansInvMinusKost = jobansInvMinusKost
                                                    end if

                                                    if (medarbSalgsOkost) <> 0 AND oRec("jobstatus") = 0 AND cint(jobansvFundet) = 0 then
                                                    medarbJobOkostProc = medarbJobOkostProc/1 + (medarbSalgsOkost * 1) '(oRec("jobans_proc_"& j) / 100)) 
                                                    jobansvFundetFak = 1
                                                    else
                                                    medarbJobOkostProc = medarbJobOkostProc
                                                    end if
                                    
                                                    jobsansProc = jobsansProc + oRec("jobans_proc_"&j)
                                                   
                                                    end if

                                                end if
                                           

                                                salgsansvFundet = 0
                                                salgsansvFundetFak = 0
                                                if isNull(oRec("salgsans"&j)) <> true then

                                                    if (cdbl(oRec("salgsans"&j)) = cdbl(usemrn) AND cint(multipleMids) = 0) OR (cdbl(oRec("salgsans"&j)) = cdbl(oRec("mid")) AND cint(multipleMids) = 1) then
                                                    salgsansBgt = salgsansBgt + (jo_bruttooms * (oRec("salgsans"&j&"_proc") / 100)) 'jo_bruttooms
                                                    salgsansInv = salgsansInv + ((fakbeloeb * oRec("salgsans"&j&"_proc")) / 100)
                                            
                                                   

                                                    if (fakbeloeb) <> 0 AND oRec("jobstatus") = 0 AND cint(salgsansvFundetFak) = 0 then
                                                    salgsansInvMinusKost = salgsansInvMinusKost/1 + (fakbeloeb * 1) '(oRec("salgsans"&j&"_proc") / 100))
                                                    salgsansvFundetFak = 1
                                                    else
                                                    salgsansInvMinusKost = salgsansInvMinusKost
                                                    end if

                                                    if (medarbSalgsOkost) <> 0 AND oRec("jobstatus") = 0 AND cint(salgsansvFundet) = 0 then
                                                    medarbSalgsOkostProc = medarbSalgsOkostProc/1 + (medarbSalgsOkost * 1) '(oRec("salgsans"&j&"_proc") / 100)) 
                                                    salgsansvFundet = 1
                                                    else
                                                    medarbSalgsOkostProc = medarbSalgsOkostProc
                                                    end if

                                                    salgsansProc = salgsansProc + oRec("salgsans"&j&"_proc")
                                                    
                                                    end if

                                                end if
                                            next   
                                    
                                        
                                             if cint(multipleMids) = 0 then
                                             salgsansBgtGT = salgsansBgtGT + salgsansInv
                                             end if

                                                 
                                            if cint(multipleMids) = 0 then
                                            jobansBgtGT = jobansBgtGT + jobansInv
                                            end if
        


                                            jobsalgsansBgtTot = (salgsansBgt/1 + jobansBgt/1)
                                            jobsalgsansInvTot = (salgsansInv/1 + jobansInv/1)

                                            if cdbl(fakbeloeb) >= cdbl(medarbSalgsOkost) then 'oRec("jo_bruttooms")
                                            indicatorCol = "yellowgreen"
                                            else
                                            indicatorCol = "crimson"
                                            end if

                                    case else

                                        fctimer = 0
                                        strSQL_r = "SELECT COALESCE(SUM(timer), 0) AS fctimer FROM ressourcer_md AS r WHERE jobid = " & oRec("jid") 
                                   
                                     
                                        oRec2.open strSQL_r, oConn, 3
                                        if not oRec2.EOF then
                                    
                                        fctimer = oRec2("fctimer")

                                        end if
                                        oRec2.close 

                                        if fctimer <> 0 then
                                        fctimer = formatnumber(fctimer, 0)
                                        else
                                        fctimer = 0
                                        end if



                                        if cdbl(fctimer) > cdbl(oRec("sumtimer")) then
                                        indicatorCol = "yellowgreen"
                                        else
                                        indicatorCol = "crimson"
                                        end if

                                    end select

                                    %>

                                <tr>
                                   
                                   
                                     <%select case lto
                                     case "epi2017"

                                             
                                           

                                            if cint(multipleMids) = 0 then 

                                            if oRec("jobstatus") = 0 AND fakbeloeb = 0 AND anloops = 0 then 'Vis ikke lukkedejob uden fakturering - Vis alle p� CLOSED joblsiten
                                            dontshowemptyJob = 1
                                            else
                                            dontshowemptyJob = 0
                                            end if
                                           

                                            'jobsalgsansBgtTotGT = jobsalgsansBgtTotGT + jobansInv/1 'jobsalgsansBgtTot/1
                                            'jobsalgsansInvTotGT = jobsalgsansInvTotGT + slagsansInv/1 'jobsalgsansInvTot/1
                                            fakbeloebGT = fakbeloebGT + fakbeloeb/1
                                            toBeInvoiced = (jo_bruttooms - fakbeloeb)
                                            toBeInvoicedGT = (toBeInvoicedGT + (toBeInvoiced)) 


                                            if cint(dontshowemptyJob) <> 1 then

                                        %>
                                            <td><%=oRec("kkundenavn") %> (<%=oRec("kkundenr") %>)</td>
                                            <td><%=oRec("jobnavn") %> (<%=oRec("jobnr") %>)
                                                <%if cint(oRec("jobstatus")) = 0 then %>
                                                <span style="font-size:9px; color:red;">Closed</span>
                                                <%end if %>
                                            </td>
                                            <td style="font-size:9px; white-space:nowrap;"><%=oRec("jobstartdato") &" - "& oRec("jobslutdato") %></td>

                                         <% if anloops = 0 then %>
                                          <td style="text-align:right; white-space:nowrap;">
                                             
                                              <%=formatnumber(jo_bruttooms, 2) & " "& basisValISO_f8%>
                                         </td>
                                        <%end if %>


                                             
                                        
                                              <td style="text-align:right;"><%=formatnumber(fakbeloeb, 2) &" "& basisValISO_f8%>

                                              <%if isDate(oRec("labeldato")) = true then %>
                                              <br />
                                              <span style="font-size:9px;"><%=formatdatetime(oRec("labeldato"), 2) %></span>
                                              <%end if %>
                                             </td>

                                          <td style="text-align:right;">
                                              <%'if anloops = 0 then %>
                                              <%'formatnumber(matforbrugsumPrJob(oRec("jid")), 2) &" "& basisValISO_f8%>
                                              <%'else %>
                                              <%=formatnumber(medarbSalgsOkost, 2) &" "& basisValISO_f8%>
                                              <%'end if %>
                                          </td>    


                                            <%if anloops > 0 then %>
                                                <td style="text-align:right;">
                                                  <%if cdbl(fakbeloeb) <> 0 then
                                                  
                                                      totDBclosedProc = ((fakbeloeb*1-medarbSalgsOkost*1)/fakbeloeb*1)*100%>

                                                      <%=formatnumber(totDBclosedProc, 0) & "%"%>
                                                       <br />
                                                      <%=formatnumber(fakbeloeb-medarbSalgsOkost, 2) & " "& basisValISO_f8%>
                                             
                                                  <%end if%>
                                                </td>
                                              <%end if %>

                                          <td style="text-align:right; white-space:nowrap;"><%=formatnumber(jobsansProc, 0)%>% <br /> <%=formatnumber((jobansInv/1), 2) &" "& basisValISO_f8 %></td>
                                          <td style="text-align:right; white-space:nowrap;"><%=formatnumber(salgsansProc, 0)%>% <br /> <%=formatnumber((salgsansInv/1), 2) &" "& basisValISO_f8 %></td>

                                          <!--<td class="text-right; white-space:nowrap;"><%=formatnumber(jobsalgsansBgtTot, 2) & " "& basisValISO_f8%></td>-->

                                           
                                      
                                          <!--<td class="text-right;"><%=formatnumber(toBeInvoiced, 2) &" "& basisValISO_f8 %></td>-->
                                    
                                        <%

                                           

                                            strEksport = strEksport & meInit &";"& oRec("kkundenavn") &";"& oRec("kkundenr") &";"& oRec("jobnavn") &";"& oRec("jobnr") &";" & oRec("jobstatus") &";"& oRec("jobstartdato") &";"& oRec("jobslutdato") &";"

                                            if anloops = 0 then 
                                            strEksport = strEksport & formatnumber(jo_bruttooms, 2) &";"
                                            else
                                            strEksport = strEksport & formatnumber(fakbeloeb-medarbSalgsOkost, 2) &";"& formatnumber(totDBclosedProc, 0) & ";"
                                            end if
                                            
                                            strEksport = strEksport & formatnumber(fakbeloeb, 2) &";"

                                            if isDate(oRec("labeldato")) = true then 
                                            strEksport = strEksport &";"& formatdatetime(oRec("labeldato"), 2)  
                                            else
                                            strEksport = strEksport &";"
                                            end if
                                            
                                            if anloops = 0 then
                                            strEksport = strEksport &";"& formatnumber(matforbrugsumPrJob(oRec("jid")), 2) 
                                            else
                                            strEksport = strEksport &";"& formatnumber(medarbSalgsOkost, 2)
                                            end if

                                            strEksport = strEksport &";"& formatnumber(jobsansProc, 0) &";"& formatnumber((jobansInv/1), 2) &";"& formatnumber(salgsansProc, 0) &";"& formatnumber((salgsansInv/1), 2) & "xx99123sy#z"
                                            
                                            
                                            'strEksport = strEksport &";"& formatnumber(toBeInvoiced, 2) & "xx99123sy#z"

                                            jobansInvTot = jobansInvTot/1 + jobansInv/1
                                            salgsansInvTot = salgsansInvTot/1 + salgsansInv/1
                                            matforbrugsumTot = matforbrugsumTot + matforbrugsumPrJob(oRec("jid"))

                                            end if 'dontshowemptyJob

                                        else 'multiple

                                            dbthis = 0

                                            if cdbl(lastMid) <> cdbl(oRec("mid")) AND cint(antalM) > 0 then

                                                call MatrixPrMedarb

                                            end if

                                            

                                        end if


                                         

                                         
                                      case else   
                                      %>
                                         <td><%=oRec("tknavn") %> (<%=oRec("kkundenr") %>)</td>
                                        <td><%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)</td>
                                        <td><%=oRec("jobslutdato") %></td>
                                        <td style="text-align:right;"><%=formatnumber(fctimer, 2) %></td>
                                        <td style="text-align:right;"><%=formatnumber(oRec("sumtimer"), 2) %></td>
                                        <td style="text-align:right;"><%=formatnumber(oRec("sumtimeOms"), 2) %></td>
                                    <%end select %>

                                   
                                    <%if cint(multipleMids) = 0  AND cint(dontshowemptyJob) <> 1 then %>
                                    
                                    <td><div style="background-color:<%=indicatorCol%>; width:10px; height:10px;">&nbsp;</div></td>
                                    <%end if %>

                                  
                                   
                                </tr>

                                <%
                              
                                if cint(multipleMids) = 1 AND lto = "epi2017" then
                                lastMid = oRec("mid")
                                lastMnavn = oRec("mnavn")
                                lastMinit = oRec("init")

                                lastMfakbeloeb = lastMfakbeloeb + fakbeloeb
                                lastMmedarbSalgsOkost = lastMmedarbSalgsOkost + medarbSalgsOkost

                                lastMjobansBgt = lastMjobansBgt + jobansInv 
                                lastMsalgsansBgt = lastMsalgsansBgt + salgsansInv

                                lastMmedarbJobOkostProc = lastMmedarbJobOkostProc + medarbJobOkostProc

                                lastMmedarbSalgsOkostProc = lastMmedarbSalgsOkostProc + medarbSalgsOkostProc

                                if oRec("jobstatus") = 0 then

                                    lastMjobsansProc = jobsansProc

                                    if jobansInvMinusKost <> 0 then
                                    lastMjobansBgtDBClosed = lastMjobansBgtDBClosed/1 + (jobansInvMinusKost/1) '((jobansInv-medarbSalgsOkost)/jobansInv) 
                                    else
                                    lastMjobansBgtDBClosed = lastMjobansBgtDBClosed
                                    end if

                                    if salgsansInvMinusKost <> 0 then
                                    lastMsalgsansBgtDBClosed = lastMsalgsansBgtDBClosed/1 + (salgsansInvMinusKost/1) '((salgsansInv-medarbSalgsOkost)/salgsansInv)
                                    else
                                    lastMsalgsansBgtDBClosed = lastMsalgsansBgtDBClosed
                                    end if


                                antalMLastMidClosed = antalMLastMidClosed + 1 
                                
                                end if

                                lastMdbthis = dbthis 

                                antalMLastMid = antalMLastMid + 1 
                                antalM = antalM + 1 
                                end if
                                 
                                    
                                 oRec.movenext
                                 wend
                                 oRec.close
                                    
                                  if cint(multipleMids) = 1 And cint(antalM) > 0 AND lto = "epi2017" then  
                                    call MatrixPrMedarb
                                  end if

                                    %>

                                 </tbody>

                                <%select case lto
                                 case "epi2017"

                                    %>
                                        <tfoot>
                                            <tr>

                                                <%if cint(multipleMids) = 0 then %>

                                                <%if anloops = 0 then %>
                                                <td colspan="4">&nbsp;</td>
                                                <%else %>
                                                <td colspan="3">&nbsp;</td>
                                                <%end if %>
                                                <!--<td style="white-space:nowrap;"><%=formatnumber(jobsalgsansBgtTotGT, 2) &" "& basisValISO_f8 %></td>-->

                                                  
                                                 <!--<td>&nbsp;</td>-->

                                                <td style="white-space:nowrap; text-align:right;"><%=formatnumber(fakbeloebGT, 2) &" "& basisValISO_f8 %></td>
                                                <td style="white-space:nowrap; text-align:right;"><%=formatnumber(matforbrugsumTot, 2) &" "& basisValISO_f8 %></td> 


                                                 <%if anloops > 0 then %>
                                                <td>&nbsp;</td>
                                                <%end if%>


                                                 <td style="white-space:nowrap; text-align:right;">
                                                     <%=formatnumber(jobansInvTot, 2) &" "& basisValISO_f8 %>
                                                     <!--<br />(<%=formatnumber(jobansInvTot, 2) &" "& basisValISO_f8 %>)-->
                                                 </td>
                                                 <!--<td>&nbsp;</td>-->
                                                 <td style="white-space:nowrap; text-align:right;">
                                                     <%=formatnumber(salgsansInvTot, 2) &" "& basisValISO_f8 %>
                                                     <!--<br />(<%=formatnumber(salgsansInvTot, 2) &" "& basisValISO_f8 %>)-->

                                                 </td>
                                                <td>&nbsp;</td>
                                                
                                                <!--<td style="white-space:nowrap;"><%=formatnumber(toBeInvoicedGT, 2) &" "& basisValISO_f8 %></td>-->

                                                <%else %>

                                              
                                                <td>&nbsp;</td>
                                               
                                                <!-- <td style="white-space:nowrap; text-align:right;"><%=formatnumber(fakbeloebGT, 2) &" "& basisValISO_f8 %></td>
                                                 <td style="white-space:nowrap; text-align:right;"><%=formatnumber(matforbrugsumTot, 2) &" "& basisValISO_f8 %></td> 
                                                -->
                                                
                                              

                                                
                                                <%
                                                if antalMLastMidClosedGt <> 0 then
                                                lastMjobansBgtDBClosedGT = (lastMjobansBgtDBClosedGT/antalMLastMidClosedGt) 
                                                lastMsalgsansBgtDBClosedGT = (lastMsalgsansBgtDBClosedGT/antalMLastMidClosedGt) 
                                                else
                                                lastMjobansBgtDBClosedGT = 0
                                                lastMsalgsansBgtDBClosedGT = 0
                                                end if
                                                %>

                                                  <td style="white-space:nowrap; text-align:right;">
                                                    <%=formatnumber(jobansBgtGT, 2) &" "& basisValISO_f8 %>
                                                 </td>
                                                  <td style="white-space:nowrap; text-align:right;">
                                                    <%=formatnumber(lastMjobansBgtDBClosedGT,0) &" %"%>
                                                 </td>

                                                 <td style="white-space:nowrap; text-align:right;">
                                                   <%=formatnumber(salgsansBgtGT, 2) &" "& basisValISO_f8 %>
                                                 </td>
                                                 <td style="white-space:nowrap; text-align:right;">
                                                    <%=formatnumber(lastMsalgsansBgtDBClosedGT,0) &" %"%>
                                                 </td>

                                              
                                                <td style="text-align:right; white-space:nowrap;">&nbsp;</td>
                                               

                                                <%end if %>
                                                </tr>
                                        </tfoot>
                                    <%
                                 case else

                                 end select
                                     %>

                               
                           
                        </table>
                 </div>
               </div><!-- END ROW -->


           


                                  

            <%select case lto
             case "epi2017"%>
                            <br /><br />
                
            <section>
                <div class="row">
                     <div class="col-lg-12">
                        <b>Export</b>
                        </div>
                    </div>
                 <form action="medarbdashboard.asp?media=export&func=export" method="post" target="_blank">
                    
                     <%
                      if anloops = 0 then    
                     strEksport = strEkspHeader & strEksport
                         %>
                     <input id="Hidden3" name="exporttxt" value="<%=strEksport%>" type="hidden" />
                         <%
                       strEksport = ""
                       strEkspHeader = ""
                      else
                     strEksport = strEkspHeader & strEksport  
                                %>
                     <input id="Hidden3" name="exporttxt" value="<%=strEksport%>" type="hidden" />
                         <%

                      end if%>
                    
                 
                     <div class="row">
                     <div class="col-lg-12 pad-r30">
                         <input id="Submit6" type="submit" value="Export to .csv file >>" class="btn btn-sm" />
                        <!--Eksporter viste kunder som .csv fil-->
                         
                         </div>


                </div>
                    </form>
             
               
            </section>
            <%end select %>

                
                    </div><!-- END panel-body -->
              </div><!-- END panel-collapse -->
             </div><!-- END panel deafult -->

               <%end if 'jobansfundet %>

              <%next 'antalLoopsJobansv %>



           





             <!-- KUNDER HVIS kundeansvarlig -->

                <%

                select case lto
                case "epi2017"

                case else

                 kundeansFundet = 0
                 kansKrijobids = " AND (tjobnr = '0' "    
                 strSQKjobans = "SELECT jobnr, kid FROM kunder "_
                 &" LEFT JOIN job ON (jobknr = kid AND jobslutdato >= '"& strDatoStartKri &"') WHERE (kundeans1 = "& usemrn &" OR kundeans2 = "& usemrn &") " 
                 
                    'response.write strSQKjobans
                    'response.flush
                    
                 oRec.open strSQKjobans, oConn, 3
                 while not oRec.EOF 
                 
                       
                    kansKrijobids = kansKrijobids & " OR tjobnr = '" & oRec("jobnr") & "'"

                    kundeansFundet = 1
                 oRec.movenext
                 wend
                 oRec.close 
                 
                    kansKrijobids = kansKrijobids & ")"   
                    


                 if cint(kundeansFundet) = 1 then
                 %>

             
                   <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTree">Kundeansvarlig - <span style="font-size:12px; font-weight:lighter;">Keyaccount</span></a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTree" class="panel-collapse collapse">
                        <div class="panel-body">

            

                 <div class="row">
            
                 <div class="col-lg-12">
                   
                     <table class="table table-stribed">
                            <thead>
                                <tr>
                                    <th style="width: 25%">Kunder (kundenr)</th>
                                    <th style="width: 25%">Job</th>
                                    <th style="width: 10%">Lev. dato</th>
                                    <th style="width: 10%">Forecast</th>
                                    <th style="width: 10%">Realiseret</th>
                                    <th style="width: 10%">Oms�tning</th>
                                    <th style="width: 10%"><!--Indikator--></th>


                                </tr>
                            </thead>
                            <tbody>
                                <%
                                 '** Real. timer SEL MEDARB KUNDEANSVARLIG
                                 strSQlmtimer_k = "SELECT SUM(t.timer) AS sumtimer, SUM(t.timer*t.timepris*kurs/100) AS sumtimeOms, SUM(t.timer*t.kostpris*kpvaluta_kurs/100) AS sumtimeKost, "_
                                 &" t.tknavn, t.tknr, t.tjobnavn, t.tjobnr, j.id AS jid, jobslutdato, k.kkundenr "_
                                 &" FROM timer AS t "_
                                 &" LEFT JOIN job AS j ON (j.jobnr = tjobnr) "_
                                 &" LEFT JOIN kunder AS k ON (k.kid = t.tknr) "_
                                 &" WHERE tjobnr <> 0 "& kansKrijobids &" AND ("& aty_sql_realhours &") GROUP BY tjobnr ORDER BY t.tknr, tjobnavn DESC LIMIT 30"
            
                                'response.write strSQlmtimer_k
                                'response.flush

                                 oRec.open strSQlmtimer_k, oConn, 3
                                 while not oRec.EOF 

                                    'if isNull(oRec("fctimer")) <> true then
                                    'fctimer = oRec("fctimer")
                                    'else
                                    fctimer = 0
                                    'end if
                                    strSQL_r = "SELECT COALESCE(SUM(timer), 0) AS fctimer FROM ressourcer_md AS r WHERE jobid = " & oRec("jid") 
                                    oRec2.open strSQL_r, oConn, 3
                                    if not oRec2.EOF then
                                    
                                    fctimer = oRec2("fctimer")

                                    end if
                                    oRec2.close 

                                    if fctimer <> 0 then
                                    fctimer = formatnumber(fctimer, 0)
                                    else
                                    fctimer = 0
                                    end if



                                    if cdbl(fctimer) > cdbl(oRec("sumtimer")) then
                                    indicatorCol = "yellowgreen"
                                    else
                                    indicatorCol = "crimson"
                                    end if
                                    %>

                                <tr>
                                    <td><%=oRec("tknavn") %> (<%=oRec("kkundenr") %>)</td>
                                    <td><%=oRec("tjobnavn") %> (<%=oRec("tjobnr") %>)</td>
                                    <td><%=oRec("jobslutdato") %></td>
                                    <td class="text-right;"><%=formatnumber(fctimer, 2) %></td>
                                    <td class="text-right;"><%=formatnumber(oRec("sumtimer"), 2) %></td>
                                    <td class="text-right;"><%=formatnumber(oRec("sumtimeOms"), 2) %></td>
                                    <td><div style="background-color:<%=indicatorCol%>; width:10px; height:10px;">&nbsp;</div></td>

                                  
                                   
                                </tr>

                                <%
                                 oRec.movenext
                                 wend
                                 oRec.close %>


                               
                            </tbody>
                        </table>
                 </div>
                  
                     
                  </div><!-- END ROW -->
                
                    </div><!-- END panel-body -->
              </div><!-- END panel-collapse -->
             </div><!-- END panel deafult -->

               

               <%end if 'jobansfundet %>

             <%end select %>

            <%end if 'jobSog %>
        
            <%end if 'usemrn = -1 %> 


            <!------------------------------- END PAGE ---------------------------------------------------------------------->

            <%end if 'Multible select medarb %>
           

  </div>
</div>






<%end select 'func %>
<!--#include file="../inc/regular/footer_inc.asp"-->