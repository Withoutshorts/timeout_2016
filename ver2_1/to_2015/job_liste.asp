

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<!--#include file="../timereg/inc/job_inc.asp"-->

<% 
    
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	end if
    response.End
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030    
    
    func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if


    call menu_2014
    
%>

<div class="wrapper">
    <div class="content">


<%
    Public left_intX
	Public function kommaFunc(x)
	if len(x) <> 0 then
	instr_komma = instr(x, ",")
		
		if instr_komma > 0 then
		left_intX = left(x, instr_komma + 2)
		else
		left_intX = x
		end if
	else
	left_intX = 0
	end if
	
	Response.write left_intX 
	end function
	
    if len(trim(request("filt"))) <> 0 then
	filt = request("filt")
    else
        if request.cookies("tsa")("statusfilt") <> "" then
        filt = request.cookies("tsa")("statusfilt")
        else
        filt = "1"
        end if
    end if
	
    response.cookies("tsa")("statusfilt") = filt
    
    aftfilt = request("aftfilt")
	sort = Request("sort")
	usedletter = request("l") 
	

   
	
	oimg = "ikon_job_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Joboversigt"
	
    if media = "print" OR request("print") = "j" then
	    call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    end if
	

    'Response.write "filt" & filt
    'Response.end
    
    chk1 = ""
	chk2 = ""
	chk3 = ""
	chk4 = ""
	chk5 = ""

    varFilt = " AND (jobstatus = -2"
    for f = 0 to 4
        'Response.write instr(filt, f) & "<br>.."

        if instr(filt, f) <> 0 then
        varFilt = varFilt & " OR jobstatus = "& f
            
            select case f
            case 0
            chk0 = "CHECKED"
            case 1
            chk1 = "CHECKED"
            case 2
            chk2 = "CHECKED"
            case 3
            chk3 = "CHECKED"
            case 4
            chk4 = "CHECKED"
            end select
        
        
        end if



    next

    varFilt = varFilt & ")"

    'Response.write "varFilt" & varFilt


	

	
	select case aftfilt
	case "serviceaft"
	varFilt = varFilt & " AND serviceaft <> 0" 'Service aftaler.
	chk6 = "CHECKED"
	chk7 = ""
	chk8 = ""
	case "ikkeserviceaft"
	varFilt = varFilt & " AND serviceaft = 0" 'Ikke Service aftaler.
	chk6 = ""
	chk7 = "CHECKED"
	chk8 = ""
	case else
	varFilt = varFilt & "" 'Vis alle.
	chk6 = ""
	chk7 = ""
	chk8 = "CHECKED"
	end select
	
	
	'* er der valgt et bogstav ****
	if len(usedletter) <> 0 then
	varFilt = varFilt & " AND jobnavn LIKE '"&request("l")&"%'" 
	'showFilter = showFilter & " der starter med:<b> " & request("l") &"</b>"
	else
	varFilt = varFilt 
	'showFilter = showFilter
	end if
	
	
		if len(request("FM_kunde")) <> 0 then
		fmkunde = request("FM_kunde")
		else
		fmkunde = 0
		end if
	    
        if len(request("frompost")) <> 0 then
            if len(trim(request("FM_vis_timepriser"))) <> 0 then
            vis_timepriser = 1
            else
            vis_timepriser = 0
            end if

           
        else
            if request.Cookies("job")("vis_timepriser") <> "" then
            vis_timepriser = request.Cookies("job")("vis_timepriser")
            else
            vis_timepriser = 0
            end if
        end if

        Response.Cookies("job")("vis_timepriser") = vis_timepriser

        if vis_timepriser = 1 then
        vis_timepriserCHK = "CHECKED"
        else
        vis_timepriserCHK = ""
        end if

	
	'''** søg i akt ***'
	if len(trim(request("FM_sogakt"))) <> 0 then
			sogakt = 1
			sogaktCHK = "CHECKED"
	else
	        sogakt = 0
	        sogaktCHK = ""		
    end if	
	
	    
        
        
        '*************************************************
        '*** Er der søgt på en specifikt jobnr/jobnavn ***
		'*************************************************
        
        if len(request("FM_kunde")) <> 0 then 'er der søgt/submitted??
        jobnr_sog = request("jobnr_sog")
        response.cookies("tsa")("jobsog") = jobnr_sog
        visliste = 1
        else

            if request.cookies("tsa")("jobsog") <> "" then
           
            jobnr_sog = request.cookies("tsa")("jobsog")
            
                if jobnr_sog = "%" then
                jobnr_sog = ""
                end if

            else
            jobnr_sog = ""
            end if
        visliste = 0
        end if




		if len(jobnr_sog) <> 0 AND jobnr_sog <> "-1" OR (cint(visliste) = 1) then
			jobnr_sog = jobnr_sog
			
			if cint(sogakt) = 1 then
			sogeKri = sogeKri & " (a.navn LIKE '%"& jobnr_sog &"%' "
			else
			    

                if instr(jobnr_sog, ">") > 0 OR instr(jobnr_sog, "<") > 0 OR instr(jobnr_sog, "--") > 0 then
           
                if instr(jobnr_sog, ">") > 0 then
                sogeKri = sogeKri &" (j.jobnr > "& replace(trim(jobnr_sog), ">", "") &" "
                end if

                if instr(jobnr_sog, "<") > 0 then
                sogeKri = sogeKri &" (j.jobnr < '"& replace(trim(jobnr_sog), "<", "") &"' "
                end if

                if instr(jobnr_sog, "--") > 0 then
                jobnr_sogArr = split(jobnr_sog, "--")
               
                for t = 0 to 1
                
                if t = 0 then
                jobSogKriA = jobnr_sogArr(0)
                else
                jobSogKriB = jobnr_sogArr(1)
                end if
                
                next

                sogeKri = sogeKri &" (j.jobnr BETWEEN '"& trim(jobSogKriA) &"' AND '"& trim(jobSogKriB) &"'"
                end if

                else

                sogeKri = " (j.jobnr LIKE '"& jobnr_sog &"' OR j.jobnavn LIKE '"& jobnr_sog &"%' OR j.id LIKE '"& jobnr_sog &"' OR Kkundenavn LIKE '"& jobnr_sog &"%' OR Kkundenr LIKE '"& jobnr_sog &"' OR rekvnr LIKE '"& jobnr_sog &"'"
			

                end if
            
            end if
			
			sogeKri = sogeKri & ") AND "
			
			
			show_jobnr_sog = jobnr_sog
			
		else
			jobnr_sog = "-1"
			sogeKri = " (j.id = -1) AND "
			show_jobnr_sog = ""
		end if


		
    '**** Projektgrupper ****'
    if len(trim(request("FM_prjgrp"))) <> 0 then
    prjgrp = request("FM_prjgrp")
    response.Cookies("job")("prjgrp") = prjgrp
    else
        if request.Cookies("job")("prjgrp") <> "" then
        prjgrp = request.Cookies("job")("prjgrp")
        else
        prjgrp = 10 '** Alle
        end if
    end if
		
	
	call datocookie
	
	'**** Brug datokriterie filer ****
	if len(request("usedatokri")) <> 0 then
		if request("usedatokri") = "j" then
		usedatoKri = "j"
		datoKriJa = "CHECKED"
		datoKriNej = ""
		else
		usedatoKri = "n"
		datoKriJa = ""
		datoKriNej = "CHECKED"
		end if
	else
		usedatoKri = request.Cookies("job")("cusedatokri")
		if usedatoKri = "j" then
		datoKriJa = "CHECKED"
		datoKriNej = ""
		else
		usedatoKri = "n"
		datoKriJa = ""
		datoKriNej = "CHECKED"
		end if
	end if
	


    if len(trim(request("FM_hd_kpers"))) <> 0 then
    hd_kpers = request("FM_hd_kpers")
    else
    hd_kpers = -1
    end if

	'*** Indsætter cookie ***
	Response.Cookies("job")("cusedatokri") = usedatoKri
	Response.Cookies("job").Expires = date + 65
	
	if usedatoKri = "j" then
	datoKri = " AND jobslutdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
	else
	datoKri = ""
	end if	
	'***** 
	
	%>

<!--

 <script src="../timereg/inc/job_listen_jav.js"></script>
    
 <div id="loadbar" style="position:absolute; display:; visibility:visible; top:300px; left:280px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid: 5-45 sek. <br />(ved mere end 1000 job)
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	
	</td></tr></table>

	</div> -->

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Joboversigt</u></h3>
                <div class="portlet-body">
                    <form action="job_liste.asp?menu=job&shokselector=1&frompost=1" method="post" name="jobnr" id="jobnr">
                    <div class="well well-white">
                        <div class="row">
                            <div class="col-lg-1">Kunde:</div>
                            <div class="col-lg-3">
                                <select name="FM_kunde" id="FM_kunde" class="form-control input-small" onchange="submit();">
		                            <option value="0">Alle <%=writethis%></option>
		                            <%
				                        strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
				                        oRec.open strSQL, oConn, 3
				                        while not oRec.EOF
				
				                        if cint(fmkunde) = cint(oRec("Kid")) then
				                        isSelected = "SELECTED"
				                        else
				                        isSelected = ""
				                        end if
				                        %>
				                        <option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
				                        <%
				                        oRec.movenext
				                        wend
				                        oRec.close
				                        %>
		                        </select>
                            </div>
                            <div class="col-lg-1">Kontaktperson</div>
                            <div class="col-lg-3">
                                <input type="hidden" value="<%=hd_kpers %>" name="FM_hd_kpers" id="FM_hd_kpers" />
                                <select name="FM_kundekpers" id="FM_kundekpers" class="form-control input-small">
		                            <option value="0">Alle (vælg kontakt for at se kontaktpersoner)</option>
		                        </select>
                            </div>
                        </div>
                        <br />
                        <div class="row">
                            <div class="col-lg-1">Søg:</div>
                            <div class="col-lg-7">
                                <input type="text" name="jobnr_sog" id="jobnr_sog" value="<%=show_jobnr_sog%>" class="form-control input-small">
                            </div>
                            <div class="col-lg-4">
                                <input id="FM_sogakt" name="FM_sogakt" type="checkbox" value="1" <%=sogaktCHK%> /> Vis kun job hvor søgekriterie indgår i en akt. på jobbet.
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1">Prg.:</div>
                            <div class="col-lg-3">
                                <select name="FM_prjgrp" id="FM_prjgrp" class="form-control input-small">
                                 <%call progrupperIdNavn(prjgrp) %>
                                </select>
                            </div>
                        </div>

                        <br />
                       
                        
                        <div id="accordion-paneled">                                                                                                       
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#forret_acc">Forretningsområder</a>
                            <div id="forret_acc" class="panel-collapse collapse">
                                <div class="panel-body" style="padding-left:95px;">
                                    <div class="row">
                                        <div class="col-lg-8">                                     
                                        <%
                                        strSQLf = "SELECT id, navn FROM fomr WHERE id <> 0 ORDER BY navn"
                                        %>
                                        <select name="FM_fomr" id="Select2" multiple="multiple" size="6" class="form-control input-small">
                                        <%if instr(strFomr_rel, "#0#") <> 0 then 
                                        f0sel = "SELECTED"
                                        else
                                        f0sel = ""
                                        end if%>

                                        <option value="#0#" <%=f0sel %>>Alle (ignorer)</option>
                                    
                                            <%oRec.open strSQLf, oConn, 3
                                            while not oRec.EOF
                                    
                                            if instr(strFomr_rel, "#"&oRec("id")&"#") <> 0 then
                                            fSel = "SELECTED"
                                            else
                                            fSel = ""
                                            end if
                                    
                                            %>
                                            <option value="#<%=oRec("id")%>#" <%=fSel %>><%=oRec("navn") %></option>
                                            <%
                                            oRec.movenext
                                            wend
                                            oRec.close
                                            %>
                                        </select>
                                        </div>
                                    </div>
                                                                                                 
                                </div> <!-- /.panel-body -->                               
                            </div>                            
                        </div>
                           

                        <div id="accordion-paneled">                                                                                                       
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#avans_acc">Avanceret</a>
                            <div id="avans_acc" class="panel-collapse collapse">
                                <div class="panel-body" style="padding-left:95px;">
                                    
                                    <div class="row">
                                        <div class="col-lg-8">
                                            <div style="border:1px #CCCCCC solid; padding:5px;">
                                            <input type="checkbox" name="FM_vis_timepriser" <%=vis_timepriserCHK %> value="1" />
                                            <b>Vis som liste med medarbejder-timepriser</b> (maks. 10000 linier)
                                            <br />
                                            Sorter efter: <input type="radio" value="0" name="FM_sorter_tp" <%=sorttpCHK0 %> /> Aktiviteter, ell.
                                            &nbsp;<input type="radio" value="1" name="FM_sorter_tp" <%=sorttpCHK1 %> />  Medarbejdere
                                            </div>
                                        </div>
                                    </div>
                                    <br />
                                    <div class="row">
                                        <div class="col-lg-3">Er jobbet tilknyttet en aftale?</div>
                                        <div class="col-lg-4"><input type="radio" name="aftfilt" value="visalle" <%=chk8%>>&nbsp;Vis alle&nbsp;&nbsp;
		                                <input type="radio" name="aftfilt" value="serviceaft" <%=chk6%>>&nbsp;Ja&nbsp;&nbsp;
		                                <input type="radio" name="aftfilt" value="ikkeserviceaft" <%=chk7%>>&nbsp;Nej&nbsp;&nbsp;</div>                                        
                                    </div>
                                    <br />
                                    <div class="row">
                                        <div class="col-lg-8">Vis kun job med slutdato i følgende periode:</div>
                                    </div>
                                    <div class="row">
                                        <div class="col-lg-8">
                                           <input type="radio" name="usedatokri" value="n" <%=datoKriNej%>> Nej
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="col-lg-12">
                                            <input type="radio" name="usedatokri" value="j" <%=datoKriJa%>> Ja 
                                            <table><tr><!--#include file="../timereg/inc/weekselector_b.asp"--></tr></table>
                                        </div>
                                    </div>
                                    <br />
                                    <div class="row">
                                        <div class="col-lg-8">Jobansvarlig:</div>
                                    </div>
                                    <div class="row">
                                        <%
	                                    if len(request("FM_medarb_jobans")) <> 0 then
	                                    medarb_jobans = request("FM_medarb_jobans")
	                                    response.cookies("tsa")("jobans") = medarb_jobans
	                                    else
	                                        if request.cookies("tsa")("jobans") <> "" then
	                                        medarb_jobans = request.cookies("tsa")("jobans")
	                                        else
	                                        medarb_jobans = 0
	                                        end if
	                                    end if
	
	                                    response.cookies("tsa").expires = date + 45
	                                    %>
                                        <div class="col-lg-5">
                                            <select name="FM_medarb_jobans" id="FM_medarb_jobans" class="form-control input-small">
                                                <option value="0">Alle (ignorer)</option>
                                                <%
                                                mNavn = "Alle"
            
                                                 strSQL = "SELECT mnavn, mnr, mid, init FROM medarbejdere WHERE mansat <> '2' ORDER BY mnavn"
                                                 oRec.open strSQL, oConn, 3
                                                 while not oRec.EOF
                
                                                    if cint(medarb_jobans) = oRec("mid") then
                                                    selThis = "SELECTED"
                                                    else
                                                    selThis = ""
                                                    end if
                
                
                                                    %>
                                                 <option value="<%=oRec("mid")%>" <%=selThis%>><%=oRec("mnavn") %> (<%=oRec("mnr") %>)
                                                 <%if len(trim(oRec("init"))) <> 0 then %>
                                                  - <%=oRec("init") %>
                                                  <%end if %>
                                                  </option>
                                                 <%
                                                 oRec.movenext
                                                 wend
                                                 oRec.close
                                                 %>                        
                                            </select>
                                        </div>
                                    </div>
                                    <br />
                                    <div class="row">
                                    <div class="col-lg-1">Status:</div>
                                    </div>
                                    <div class="row">
                                        <div class="col-lg-3">
                                        &nbsp&nbsp&nbsp&nbsp<input type="CHECKBOX" name="filt" value="1" <%=chk1%>/> Aktive <br />
                                        &nbsp&nbsp&nbsp&nbsp<input type="CHECKBOX" name="filt" value="2" <%=chk2%>/> Passive / Til fakturering <br />
                                        &nbsp&nbsp&nbsp&nbsp<input type="CHECKBOX" name="filt" value="3" <%=chk3%>/> Tilbud <br />
                                        &nbsp&nbsp&nbsp&nbsp<input type="CHECKBOX" name="filt" value="4" <%=chk4%>/> Gennemsyn <br />
                                        &nbsp&nbsp&nbsp&nbsp<input type="CHECKBOX" name="filt" value="0" <%=chk0%>/> Lukkede
                                        </div>
                                    </div>                                                                                                                                   
                                </div> <!-- /.panel-body -->                               
                            </div>                            
                        </div>
                        
                        <div class="row">
                        <div class="col-lg-10">&nbsp</div>
                        <div class="col-lg-2 pad-b10">
                            <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                        </div>
                        </div>
                                                                       
                    </div> <!-- Søgefilter -->
                    </form>

                    <%
                        '** Vælger sql sætning efter sortering og filter ***
	
	                    if fmkunde <> 0 then
	                    varJobknrKri = "j.jobknr = "& Request("FM_Kunde") &" AND "
	                    else
		                    'if request("fromvemenu") = "j" then
		                    'varJobknrKri = "j.jobknr = 0 AND "
		                    'else
		                    varJobknrKri = ""
		                    'end if
	                    end if
	
	                    if usedatoKri = "j" then
	                    varOrder = "jobslutdato DESC"
	                    else
	                    varOrder = "kkundenavn, jobnavn"
	                    end if
	
                        '** jobans **'
	                    if cint(medarb_jobans) = 0 then
	                    jobansKri = ""
	                    else
	                    jobansKri = " (jobans1 = " & medarb_jobans & " OR jobans2 = "& medarb_jobans &") AND "
	                    end if
	

                        '** kontaktpersoner hos kunde **'
                        if cint(hd_kpers) <> -1 AND cint(hd_kpers) <> 0  then
                        kpersSQLkri = " AND kundekpers = " & hd_kpers
                        else
                        kpersSQLkri = ""
                        end if


                        if prjgrp <> 0 AND prjgrp <> 10 then
                        prjgrpSQLkri = " AND (projektgruppe1 = "&prjgrp&" "_
                        &" OR projektgruppe2 = "&prjgrp&" OR projektgruppe3 = "&prjgrp&" OR projektgruppe4 = "&prjgrp&" "_
                        &" OR projektgruppe5 = "&prjgrp&" OR projektgruppe6 = "&prjgrp&" OR projektgruppe7 = "&prjgrp&" OR projektgruppe8 = "&prjgrp&" OR projektgruppe9 = "&prjgrp&" OR projektgruppe10 = "&prjgrp&")"
                        else
                        prjgrpSQLkri = ""
                        end if
                        
                        
                        
                        if cint(vis_timepriser) <> 1 then  
                    %>

                    
                    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <%sub tabletop %>
                        <tr>
                            <th>Job&nbsp;/&nbsp;Jobansv.<br />Kunde&nbsp;/&nbsp;Per.</th>
                            <th>Aktiviteter</th>
                            <th>Forretningsomr.</th>
                            <th>Realiseret %<br />(heraf fakturerbare)</th>
                            <th>Brutto oms<br />(Budget)</th>
                            <th>Timer forkalk.<br>Realiseret<br />Balance
                            <th>Status</th>
                            <th>Funktioner</th>
                            <th>Faktura hist.<br /> Faktisk timepris</th>
                            <th>Faktureret beløb <br />(ekskl. mat. og km.) <br />/ real. timer</th>
                            <th>Projekgrupper</th>
                        </tr>
                        <%end sub %>

                        <%
                        '***********************************************************
	                    '*** Main SQL **********************************************
	                    '***********************************************************
	                    strSQL = "SELECT j.id, jobnavn, jobnr, kkundenavn, kid, jobknr, jobTpris, jobstatus, jobstartdato, "_
	                    &" jobslutdato, j.budgettimer, fakturerbart, Kkundenr, ikkebudgettimer, jobans1, jobans2, jobans3, jobans4, jobans5, fastpris, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, virksomheds_proc, "_
	                    &" s.navn AS aftnavn, rekvnr, tilbudsnr, sandsynlighed, jo_bruttooms, kundekpers, serviceaft, lukkedato, preconditions_met, jo_valuta"
	
                        strSQL = strSQL &", j.projektgruppe1, j.projektgruppe2, j.projektgruppe3, j.projektgruppe4, j.projektgruppe5, j.projektgruppe6, j.projektgruppe7, j.projektgruppe8, j.projektgruppe9, j.projektgruppe10 "
	   
	
	                    if cint(sogakt) = 1 then
	                        strSQL = strSQL &" FROM aktiviteter AS a "
	                        strSQL = strSQL &" LEFT JOIN job AS j ON (j.id = a.job)"
	                        strSQL = strSQL &" LEFT JOIN kunder ON (kunder.kid = jobknr)"
	                    else
                            strSQL = strSQL &" FROM job AS j "
                            strSQL = strSQL &" LEFT JOIN kunder ON (kunder.kid = jobknr)"
                        end if

	
	                    strSQL = strSQL &" LEFT JOIN serviceaft s ON (s.id = serviceaft)"_
	                    &" WHERE "& varJobknrKri &" "& sogeKri &" "& jobansKri &" kunder.Kid = jobknr " & varFilt &" "& datoKri &" "& prjgrpSQLkri &" "& kpersSQLkri &""_
	                    &" GROUP BY j.id ORDER BY "&varOrder&" LIMIT 5000"
	
	
	
	                    x = 0
	                    cnt = 0
	                    jids = 0
	                    totReal = 0
	                    gnsPrisTot = 0
	                    totRealialt = 0
                        thisMid = session("mid")
	
	                    'Response.write strSQL
	                    'Response.Flush
	
	                    oRec.open strSQL, oConn, 3
	                    while not oRec.EOF

                        call forretningsomrJobId(oRec("id"))


                         if cint(visJobFomr) = 1 OR instr(strFomr_rel, "#0#") <> 0 then


	
	                    '** Til Export fil ***
	                    jids = jids & "," & oRec("id")
	
	                    if cnt = 0 then
	                    call tabletop
	                    end if
	

                        if len(trim(oRec("budgettimer"))) = "0" OR oRec("budgettimer") = "0" then 
	                    budgettimer = 0
	                    else
	                    budgettimer = oRec("budgettimer")
	                    end if
	
	                    if oRec("ikkebudgettimer") > 0 then 
	                    ikkebudgettimer = oRec("ikkebudgettimer")
	                    else
	                    ikkebudgettimer = 0
	                    end if
	
	                    '** tildelt total ***
	                    budgettimertot = (ikkebudgettimer + budgettimer)


	                    call akttyper2009(2)

	                    '*** Real timer og Proafsluttet **************************
	                    strSQL = "SELECT sum(timer) AS proafslut FROM timer WHERE Tjobnr = '" & oRec("jobnr") &"' AND ("& aty_sql_realhours &") "
	                    oRec3.open strSQL, oConn, 3
	
	                    if len(oRec3("proafslut")) <> 0 then
	                    proaf = oRec3("proafslut")
	                    else
	                    proaf = 0
	                    end if

	                    oRec3.close

                        totReal = proaf
                        restt = (budgettimertot - proaf)
    
                        realfakbare = 0
                        '*** Real. fakturerbare timer **************************
	                    strSQL = "SELECT sum(timer) AS realfakbare FROM timer WHERE Tjobnr = '" & oRec("jobnr") &"' AND tfaktim = 1"
	                    oRec3.open strSQL, oConn, 3
	                    if not oRec3.EOF then

	                    if len(oRec3("realfakbare")) <> 0 then
	                    realfakbare = oRec3("realfakbare")
	                    else
	                    realfakbare = 0
	                    end if

	                    end if
	                    oRec3.close
	

                        resttfakbare = (budgettimertot - realfakbare)

                        'timerTotFakbarePajob = realfakbare 


                        'if len(timerTotFakbarePajob) <> 0 then
	                    'timerTotFakbarePajob = timerTotFakbarePajob
	                    'else
	                    'timerTotFakbarePajob = 0
	                    'end if



		
		                    '*** bgcolor ***
		                    if cdbl(id) = oRec("id") then
		                    bc = "#FFFF99"
		                    else
			                    select case right(cnt, 1) 
			                    case 0,2,4,6,8
			                    bc = "#FFFFFF"
			                    case else
			                    bc = "#EFF3FF" '"#5582D2" '"#d2691e"
			                    end select
		                    end if
	
	
	
	                    sletkanp = "<img src='../ill/slet_16.gif' alt='Slet' border='0'>" 
	
	                    'if oRec("fakturerbart") = "1" then
		                    if oRec("jobstatus") = 3 then
		                    inteksimg = "<img src='../ill/filtertilbud.gif' width='10' height='21' alt='' border='0'>"
		                    else
		                    inteksimg = "<img src='../ill/blank.gif' width='15' height='15' alt='' border='0'>"
		                    'eksternt_job_trans
		                    end if
	                    'else
	                    'inteksimg = "<img src='../ill/internt_job_trans.gif' width='15' height='15' alt='' border='0'>"
	                    'end if
	
	


                        '*** Oms. timer ****'
                        jo_bruttooms = oRec("jo_bruttooms")
	                    jo_fastpris = oRec("fastpris")


                        '** Prekonditioner ****'
                        preconditions_met = oRec("preconditions_met")
    
	
	                    if budgettimer <> 0 then
	                    projektcomplt = ((proaf/budgettimertot)*100)
                        projektcompltFakbare = ((realfakbare/budgettimertot)*100)
	                    else
	                    projektcompltFakbare = 100
                        projektcomplt = 100
	                    end if
	
	                    if projektcomplt > 100 then
	                    showprojektcomplt = projektcomplt
	                    projektcomplt = 100
	                    bgdiv = "Crimson"
	                    else
	                    showprojektcomplt = projektcomplt
	                    projektcomplt = projektcomplt
	                    bgdiv = "yellowgreen"
	                    end if

                        if projektcompltFakbare > 100 then
	                    showprojektcompltFakbare = projektcompltFakbare
	                    projektcompltFakbare = 100

	                    else
	                    showprojektcompltFakbare = projektcompltFakbare
	                    projektcompltFakbare = projektcompltFakbare
	
	                    end if

    
	
	                    '*** Antal aktiviteter på job *** KUN AKTIVE I DETTE VIEW 
	                    strAktnavn = ""
                        lastFase = ""
                        strSQL2 = "SELECT id, navn, fase, aktstatus FROM aktiviteter WHERE job = "& oRec("id") & " AND aktstatus = 1 ORDER BY fase, sortorder, navn"
	                    oRec5.open strSQL2, oConn, 3
	                    while not oRec5.EOF 
	                    x = x + 1

                        if lastFase <> oRec5("fase") AND isNull(oRec5("fase")) <> true AND len(trim(oRec5("fase"))) <> 0 then

                        strAktnavn = strAktnavn & "<br><b>"& replace(oRec5("fase"), "_", " ") & "</b><br>"
                        lastFase = oRec5("fase")
                        end if

                        strAktnavn = strAktnavn & left(oRec5("navn"), 20) 
        
                            if oRec5("aktstatus") = 0 then
                            strAktnavn = strAktnavn & " - lukket"
                            end if

                            if oRec5("aktstatus") = 2 then
                            strAktnavn = strAktnavn & " - passiv"
                            end if
        
                        strAktnavn = strAktnavn & "<br>"


	                    oRec5.movenext
	                    wend
	                    oRec5.close
	                    Antal = x
	
	                    ''** Antal brugte fakturerbare timer **
	                    'strSQL3 = "SELECT sum(timer) AS timerTotFaktimerpajob FROM timer WHERE Tjobnr= " & oRec("jobnr") &" AND ("& aty_sql_realhours &")"
	                    'oRec3.open strSQL3, oConn, 3
	                    'if not oRec3.EOF then
	                    'timerTotFakbarePajob = oRec3("timerTotFaktimerpajob")
	                    'end if

	                    'oRec3.close
	
	                    'if len(timerTotFakbarePajob) <> 0 then
	                    'timerTotFakbarePajob = timerTotFakbarePajob
	                    'else
	                    'timerTotFakbarePajob = 0
	                    'end if
	
	                    '*** fakturerbare timer tildelt på aktiviteter **** 
	                    'strSQL3 = "SELECT sum(budgettimer) AS akttimer FROM aktiviteter WHERE job = " & oRec("id") &" AND fakturerbar = 1 ORDER BY budgettimer"
	                    'oRec3.open strSQL3, oConn, 3
	                    'if not oRec3.EOF then
	                    '	akttfaktimtildelt = oRec3("akttimer")
	                    'end if
	                    'oRec3.close
	
	                    'if len(akttfaktimtildelt) <> 0 then
	                    'akttfaktimtildelt = akttfaktimtildelt
	                    'else
	                    'akttfaktimtildelt = 0
	                    'end if
	
	
	                    '***Ikke fakturerbare timer tildelt på aktiviteter **** 
	                    'strSQL3 = "SELECT sum(budgettimer) AS aktnotftimer FROM aktiviteter WHERE job = " & oRec("id") &" AND fakturerbar = 0 ORDER BY budgettimer"
	                    'oRec3.open strSQL3, oConn, 3
	                    'if not oRec3.EOF then
	                    '	akttnotfaktimtildelt = oRec3("aktnotftimer")
	                    'end if
	                    'oRec3.close
	
	                    'if len(akttnotfaktimtildelt) <> 0 then
	                    'akttnotfaktimtildelt = akttnotfaktimtildelt
	                    'else
	                    'akttnotfaktimtildelt = 0
	                    'end if



	                    %>

                        <tr>
		<td bgcolor="#d6dff5" colspan="11"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bc%>">
	
	<%
	'*** Tjekker rettigehder eller om man er jobanssvarlig ***
	editok = 0
	if level <= 2 OR level = 6 then
	editok = 1
	else
			if cint(session("mid")) = oRec("jobans1") OR cint(session("mid")) = oRec("jobans2") OR (cint(oRec("jobans1")) = 0 AND cint(oRec("jobans2")) = 0) then
			editok = 1
			else
			editok = 0   
			end if
	end if
	
	
	              
	
	
	'********* Jobnavn og nr ****
	%>
	<td valign=top style="padding:4px 10px 4px 2px; width:250px;">
      <b><%=oRec("kkundenavn")%></b>&nbsp;(<%=oRec("Kkundenr")%>)<br />
	<%
	'Response.write  cint(session("mid")) &"="& oRec("jobans1") &" OR " & cint(session("mid")) &" = " & oRec("jobans2") &" OR "& cint(oRec("jobans1")) &" = 0 OR" & cint(oRec("jobans2")) &" = 0 then"
			
	
	if editok = 1 then
		'if oRec("fakturerbart") = "1" then%>
		<a href="jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=<%=oRec("fakturerbart")%>&jobnr_sog=<%=jobnr_sog%>&filt=<%=filt%>&fm_kunde_sog=<%=fmkunde%>"><%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)</a>&nbsp;
		<%'else%>
		<!--<a href="jobs.asp?menu=job&func=red&id=<=oRec("id")%>&int=<=oRec("fakturerbart")%>&jobnr_sog=<=jobnr_sog%>&filt=<=filt%>&fm_kunde_sog=<=fmkunde%>" class=vmenu><=oRec("jobnavn")%>&nbsp;&nbsp;(<=oRec("jobnr")%>)</a>&nbsp;-->
		<%'end if
	else
		%>
		<b><%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)</b>&nbsp;
		<%
	end if
	

    if oRec("fastpris") = 1 then
    strFasptpris = "Fastpris"
    else
    strFasptpris = "Lbn. timer"
    end if


    %>
    <span style="color:green; font-size:10px;">(<%=strFasptpris %>)</span>

    <%
	
	
	
	'******************************
	%>
    <span style="color:#999999; font-size:9px;"><i>
    <%
		
						'*** Jobansvarlige ***
                        if oRec("jobans1") <> 0 then
						call meStamData(oRec("jobans1"))
                        %>
						<%=meNavn%> 
                             <%if oRec("jobans_proc_1")  <> 0 then %>
                            (<%=oRec("jobans_proc_1") %>%)
                            <%end if %>

                        <%end if
                        

                        '*** Jobejer 2 ***
                        if oRec("jobans2") <> 0 then
						call meStamData(oRec("jobans2"))
                        %>
						, <%=meNavn%> 
                        
                        <%if oRec("jobans_proc_2")  <> 0 then %>
                        (<%=oRec("jobans_proc_2") %>%)
                        <%end if %>

                        <%end if

                        '*** Job medansvarlige ***
                        if oRec("jobans3") <> 0 then
						call meStamData(oRec("jobans3"))
                        %>
						, <%=meNavn%> 
                         <%if oRec("jobans_proc_3")  <> 0 then %>
                        (<%=oRec("jobans_proc_3") %>%)

                        <%end if
                        
                        end if


                        '*** Job medansvarlige ***
                        if oRec("jobans4") <> 0 then
						call meStamData(oRec("jobans4"))
                        %>
						, <%=meNavn%> 
                             <%if oRec("jobans_proc_4")  <> 0 then %>
                            (<%=oRec("jobans_proc_4") %>%)
                            <%end if %>

                        <%end if


                        '*** Job medansvarlige ***
                        if oRec("jobans5") <> 0 then
						call meStamData(oRec("jobans5"))
                        %>
						, <%=meNavn%> 
                             <%if oRec("jobans_proc_5")  <> 0 then %>
                            (<%=oRec("jobans_proc_5") %>%)
                            <%end if %>

                        <%end if

                        if level = 1 then
                             if (lto <> "epi" AND lto <> "epi2017") OR (lto = "epi" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59 OR thisMid = 1720)) OR (lto = "epi_cati" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59)) OR (lto = "epi_no" AND thisMid = 2) OR (lto = "epi_ab" AND thisMid = 2) OR (lto = "epi_sta" AND thisMid = 2) OR (lto = "epi_uk" AND thisMid = 2) then 
                                if oRec("virksomheds_proc") <> 0 then%>
                                <br /><b><%=lto %>: </b> (<%=oRec("virksomheds_proc") %>%)
                                <%end if
                            end if
                        end if
                      
						
						'**********************
						%></i></span>


                        <%'*** kontakpersoner hos kunde ***'
                         
                         kpersNavn = ""
                         if cint(oRec("kundekpers")) <> 0 then

                         strSQLkpers = "SELECT navn FROM kontaktpers WHERE id = " & oRec("kundekpers")
                         oRec6.open strSQLkpers, oConn, 3
                         if not oRec6.EOF then

                         kpersNavn = oRec6("navn")

                         end if
                         oRec6.close

                         %>
                         <span style="color:#8caae6; font-size:9px;"><br />Kontaktpers.: <%=kpersNavn %></span>
                         <%

                         end if
                         
                         %>
							
						
	<br><span style="font-size:9px;"><%=formatdatetime(oRec("jobstartdato"), 1)%> - <%=formatdatetime(oRec("jobslutdato"), 1)%></span>
	
    
	<% 
	if len(trim(oRec("rekvnr"))) <> 0 then
    %>
    <span style="font-size:9px; color:#999999;">
    <%
	Response.Write "<br>Rekvnr.: "& oRec("rekvnr") 
    %>
     </span>
    <%
	end if

  
	if len(trim(oRec("aftnavn"))) <> 0 then
    %><span style="font-size:9px; background-color:#FFFFe1; color:#000000;"><%
	Response.Write "<br>Aftale: "& oRec("aftnavn") 
    %>
     </span>
    <%
	end if
	
	if oRec("jobstatus") = 3 then
    %><span style="font-size:9px; background-color:#ffdfdf; color:#000000;"><%
	Response.Write "<br>Tilbud: "& oRec("tilbudsnr") &" ("& oRec("sandsynlighed") &" %)"
    %>
     </span>
    <%
	end if


    select case preconditions_met    
      case 0
    preconditions_met_Txt = ""
      case 1
    preconditions_met_Txt = "<br><span style='color:#6CAE1C; font-size:10px; background:#DCF5BD;'>Pre-konditioner opfyldt</span>"
    case 2
    preconditions_met_Txt = "<br><span style='color:red; font-size:10px; background:pink;'>Pre-konditioner ikke opfyldt!</span>"
    case else
    preconditions_met_Txt = ""
    end select


    %>
    <%=preconditions_met_Txt %>
	
	&nbsp;</td>
	
        <td valign=top style="padding:4px 10px 0px 5px;">
            Aktiviteter (aktive)
		 <%
		'********* Aktiviteter ****
		if editok = 1 then%>
        <a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&jobid=<%=oRec("id")%>&jobnavn=<%=oRec("jobnavn")%>&rdir=job3&nomenu=1', '', 'width=1354,height=800,resizable=yes,scrollbars=yes')" class=vmenu>(<%=Antal%>)</a>
		<%else%>
		<b>(<%=Antal%>)</b>
		<%end if%>

        <%
        'if visAkt <> 1 then
        %>
        <div style="font-size:10px; height:80px; width:120px; overflow:auto; color:#999999; padding-right:5px;">
           
            
            <%=strAktnavn %></div>

        <%'end if
		'**************************
		%>
		</td>
		

        

        <td valign=top style="padding:4px 5px 5px 5px;" class=lille><%=strFomr_navn %>&nbsp;</td>

		<td align="left" valign=top style="padding:10px 5px 0px 5px;">
		<%if proaf <> 0 then %>
		
		
		<div style="font-size:9px; color:#999999,">Real.:
		<span style="width:<%=cint(left(projektcomplt, 3))%>px; background-color:<%=bgdiv%>; height:10px; padding:2px 0px 2px 2px; font-size:9px;">
		<%if showprojektcomplt > 0 then%>
		<%=formatpercent(showprojektcomplt/100, 0)%>
		<%end if%>
		</span>
        </div>

        <div style="font-size:9px; color:#999999;">Fakt.: 
        <span style="width:<%=cint(left(projektcompltFakbare, 3))%>px; background-color:#CCCCCC; color:#000000; height:10px; padding:2px 0px 2px 2px; font-size:9px;">
		<%if showprojektcompltFakbare > 0 then%>
		<%=formatpercent(showprojektcompltFakbare/100, 0)%>
		<%end if%>
		</span>
            </div>

      

		
		
		
		<%end if %>
		</td>

        <%call valutakode_fn(oRec("jo_valuta")) %>
		<td align=right valign=top style="padding:4px 5px 0px 5px;" class=lille>
		<%=formatnumber(oRec("jo_bruttooms"), 2) &" "& valutaKode_CCC %> 
		</td>
		
		<td align=right valign=top style="padding:4px 15px 4px 4px;" class=lille>Forkalk: <%=formatnumber(budgettimertot)%><br>
		Realiseret:  <%=formatnumber(proaf)%><br>
        <span style="color:#999999;">Heraf fakturerbar.: <%=formatnumber(realfakbare, 2) %></span> <br />
        -----------------------<br />
        Bal. = <%=formatnumber(restt)%><br />
        <span style="color:#999999;"><%=formatnumber(resttfakbare) %></span>
        =====================
		</td>
		
		<td valign=top style="padding:4px 5px 0px 5px;"><%
		
		'if oRec("jobstatus") <> "3" then 'Tilbud
		
        stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
        stCHK3 = ""
        stCHK4 = ""
        stBgcol0 = "#999999"
		stBgcol1 = "#999999"
		stBgcol2 = "#999999"
        stBgcol3 = "#999999"
        stBgcol4 = "#999999"

        lkDato = ""

		select case oRec("jobstatus")
        case "0"
        lkDato = "("& formatdatetime(oRec("lukkedato"), 2) &")"
        stCHK0 = "CHECKED"
        stBgcol0 = "darkred"
		case "1"
		stCHK1 = "CHECKED"
        stBgcol1 = "green"
        case "2"
        stCHK2 = "CHECKED"
        stBgcol2 = "#999999"
		case "0"
        stCHK0 = "CHECKED"
        stBgcol0 = "Crimson"
        case "3"
        stCHK3 = "CHECKED"
        stBgcol3 = "#000000"
        case "4"
        stCHK4 = "CHECKED"
        stBgcol4 = "#5582d2"
		end select
		
		if editok = 1 then
		%>
        <!--
		<select name="FM_listestatus" id="FM_listestatus" style="background-color:<%=selbgcol%>; font-size:9px;">
		<option value="1" <%=stCHK1%>>Aktiv</option>
		<option value="0" <%=stCHK0%>>Lukket</option>
		<option value="2" <%=stCHK2%>>Passiv</option>
		</select>
        -->
        
        <input type="radio" class="FM_listestatus_1" name="FM_listestatus_<%=oRec("id")%>" value="1" id="FM_listestatus_1_<%=oRec("id")%>" <%=stCHK1%>/><span style="color:<%=stBgcol1%>; font-size:9px;">Aktiv</span><br />
        <input type="radio" class="FM_listestatus_2" name="FM_listestatus_<%=oRec("id")%>" value="2" id="FM_listestatus_2_<%=oRec("id")%>" <%=stCHK2%>/><span style="color:<%=stBgcol2%>; font-size:9px;">Passiv/Fak.</span><br />
        <input type="radio" class="FM_listestatus_3" name="FM_listestatus_<%=oRec("id")%>" value="3" id="FM_listestatus_3_<%=oRec("id")%>" <%=stCHK3%>/><span style="color:<%=stBgcol3%>; font-size:9px;">Tilbud<br />
        <input type="radio" class="FM_listestatus_4" name="FM_listestatus_<%=oRec("id")%>" value="4" id="FM_listestatus_4_<%=oRec("id")%>" <%=stCHK4%>/><span style="color:<%=stBgcol4%>; font-size:9px;">Gen.syn</span><br />
        <input type="radio" class="FM_listestatus_0" name="FM_listestatus_<%=oRec("id")%>" value="0" id="FM_listestatus_0_<%=oRec("id")%>" <%=stCHK0%>/><span style="color:<%=stBgcol0%>; font-size:9px;">Lukket <br /><%=lkDato %></span>

		<%else%>
		<%=stNavn%>
		<input type="hidden" name="FM_listestatus" id="Hidden1" value="<%=oRec("jobstatus")%>">
		
		<%end if %>
		
		<%'else%>
        <!--
		<b>Tilbud</b><br />
		<span style="font-size:9px; font-family:arial;"><%=oRec("tilbudsnr") %> <br /><%=oRec("sandsynlighed") %> %</span>
		<input type="hidden" name="FM_listestatus" id="FM_listestatus" value="<%=oRec("jobstatus")%>">
        -->
		<%'end if%>
		
		
		<input type="hidden" name="FM_listeid" id="FM_listeid" value="<%=oRec("id")%>">
		</td>
	
		<td valign=top style="padding:4px 5px 0px 5px; white-space:nowrap;">
		<%if editok = 1 then %>
		<a href="job_print.asp?id=<%=oRec("id")%>" class=rmenu>Print / PDF >></a><br />
		<a href="job_kopier.asp?func=kopier&id=<%=oRec("id")%>&fm_kunde=<%=oRec("kid")%>&filt=<%=request("filt")%>" class=rmenu>Kopier job >></a><br />
         <a href="materialer_indtast.asp?id=<%=oRec("id")%>&fromsdsk=0&aftid=<%=oRec("serviceaft")%>" target="_blank" class=rmenu>Indtast mat./udlæg >></a>

         <%
         
        viskunabnejob0 = "1"
        viskunabnejob1 = "" 
        viskunabnejob2 = ""
        select case oRec("jobstatus")
        case 0
        viskunabnejob0 = "1"
        case 1,2
        viskunabnejob2 = "1"
        case 3
        viskunabnejob1 = "1"
        end select

         select case lto
        case "synergi1", "xintranet - local"


        call licensStartDato()
        dtlink_stdag = startDatoDag
		dtlink_stmd = startDatoMd
		dtlink_staar = startDatoAar

        case else
        
        dtlink_stdag = datepart("d", oRec("jobstartdato"), 2, 2)
		dtlink_stmd = datepart("m", oRec("jobstartdato"), 2, 2)
		dtlink_staar = datepart("yyyy", oRec("jobstartdato"), 2, 2)

		
        end select

        

		dtlink_sldag = datepart("d", now, 2, 2)
		dtlink_slmd = datepart("m", now, 2, 2)
		dtlink_slaar = datepart("yyyy", now, 2, 2)
		
       
		
		dtlink = "FM_usedatokri=1&FM_start_dag="&dtlink_stdag&""_
	    &"&FM_start_mrd="&dtlink_stmd&"&FM_start_aar="&dtlink_staar&"&FM_slut_dag="&dtlink_sldag&""_
	    &"&FM_slut_mrd="&dtlink_slmd&"&FM_slut_aar="&dtlink_slaar
         
          %>

          <br />
          <a href="joblog.asp?nomenu=1&FM_job=<%=oRec("id") %>&FM_kunde=<%=oRec("kid")%>&FM_jobsog=<%=oRec("jobnr")%>&viskunabnejob0=<%=viskunabnejob0%>&viskunabnejob1=<%=viskunabnejob1%>&viskunabnejob2=<%=viskunabnejob2%>&<%=dtlink %>" class=rmenu target="_blank">Joblog >></a><br />
          <a href="jobprintoverblik.asp?menu=job&id=<%=oRec("id")%>&media=printjoboverblik" class=rmenu target="_blank">Print joboverblik >></a><br>
          <a href="timereg_akt_2006.asp?FM_kontakt=<%=oRec("kid")%>&FM_ignorer_projektgrupper=1&jobid=<%=oRec("id")%>&FM_sog_job_navn_nr=<%=oRec("jobnr")%>&usemrn=<%=session("mid")%>&showakt=1&fromsdsk=1" target="_blank" class=rmenu>Timeregistrering >> </a><br />
          
          <%if cint(useasfak) <= 2 then %>
          <a href="../timereg/erp_opr_faktura_fs.asp?FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&<%=dtlink %>&reset=1&FM_usedatokri=1" target="_blank" class=rmenu>Opret faktura >> </a>
		  <%end if %>

		<%else %>
		&nbsp;
		<%end if %></td>
		<td class=lille valign=top style="padding-right:2px;">
		<div style="position:relative; top:0px, left:0px; width:160px; height:100px; overflow:auto; padding:5px;">
        <table cellspacing=0 border=0 cellpadding=0 width=100%>
		
		<%
		
		        '** Findes der fakturaer på job kan det ikke slettes **'
			    
			    deleteok = 0
			    faktotbel = 0
                fakturaBel_tot = 0
			    strSQLffak = "SELECT f.fid, f.faknr, f.aftaleid, f.faktype, f.jobid, f.fakdato, f.beloeb, "_
			    &" f.faktype, f.kurs, SUM(fd.aktpris) AS aktbel, brugfakdatolabel, labeldato, fakadr FROM fakturaer f "_
			    &" LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid AND fd.enhedsang <> 3)"_
			    &" WHERE jobid = " & oRec("id") & ""_
			    &" AND aftaleid = 0 AND shadowcopy = 0"_
			    &" GROUP BY f.fid ORDER BY f.fakdato DESC"
			    'Response.Write "strSQLffak<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>" & strSQLffak 
			    'Response.flush
			    f = 0
			    oRec3.open strSQLffak, oConn, 3
			    while not oRec3.EOF 
			        

                    if oRec3("brugfakdatolabel") = 1 then '** Labeldato
                    fakDato = "L: "& oRec3("labeldato") & "<br><span style=""color:#999999; size:9px,"">"& oRec3("fakdato") &"</span>"
                    else
                    fakDato = "F: "& oRec3("fakdato")
                    end if
			       
			          
                      %>
                      <tr><td class=lille valign=top>
                      <%
			        if cdate(oRec3("fakdato")) >= cdate("01-01-2006") AND editok = 1 then%>
			        <a href="erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id=<%=oRec3("fid")%>&FM_jobonoff=<%=FM_jobonoffval%>&FM_kunde=<%=oRec3("fakadr")%>&FM_job=<%=oRec3("jobid")%>&FM_aftale=<%=oRec3("aftaleid")%>&fromfakhist=1" class="lgron" target="_blank"><b><%=oRec3("faknr") &"</b>"%></a></td>
                    <td align=right class=lille valign=top> <%=fakDato %> <br />
                    <%else%>
                    <b><%=oRec3("faknr") &"</b></td><td align=right class=lille valign=top> "& fakDato %><br />
                    <% end if
                    

                      %>

                      </td></tr>
                      <%
                    
                     
	                call beregnValuta(minus&(oRec3("beloeb")),oRec3("kurs"),100)
                    if oRec3("faktype") <> 1 then
                    belobGrundVal = valBelobBeregnet
                    else
                    belobGrundVal = -valBelobBeregnet
                    end if 
                    

                        if cDate(oRec3("fakdato")) < cDate("01-06-2010") AND (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk") then
                        belobKunTimerStk = belobKunTimerStk + oRec3("beloeb")
                        else
       
                    
	                        '** Kun aktiviteter timer, enh. stk. IKKE materialer og KM
	                        call beregnValuta(minus&(oRec3("aktbel")),oRec3("kurs"),100)
                            if oRec3("faktype") <> 1 then
                            belobKunTimerStk = valBelobBeregnet
                            else
                            belobKunTimerStk = -valBelobBeregnet
                            end if 


                     end if
            
                fakturaBel_tot = fakturaBel_tot + (belobGrundVal)
                faktotbel = faktotbel + belobKunTimerStk
			    f = f + 1
			    oRec3.movenext
			    wend
			    oRec3.close

                fakturaBel_tot_gt = fakturaBel_tot_gt + fakturaBel_tot
			    %>
                </table>
               

			    </div>
                 Faktureret (basisvaluta):<br /> 
                 <b><%=formatnumber(fakturaBel_tot, 2) & " "& basisValISO %></b>

                 <%if totReal <> 0 then 
	        gnstpris = faktotbel/totReal
	        else 
	        gnstpris = 0
	        end if%>
	        
            <%if gnstpris <> 0 then %>
            <br />Faktisk timepris:<br />
	        <b><%=formatnumber(gnstpris) & " "& basisValISO %></b>
            
            <!-- <span class="qmarkhelp" id="qm0001" style="font-size:11px; color:#999999; font-weight:bolder;">?</span><span id="qmarkhelptxt_qm0001" style="visibility:hidden; color:#999999; display:none; padding:3px; z-index:4000;">(faktureret beløb - (materialeforbrug + km)) / timer realiseret</span>-->
	        
	        <%
            end if

	        if f <> 0 AND proaf <> 0 then
	        gnsPrisTot = gnsPrisTot + (faktotbel)
	        else
	        gnsPrisTot = gnsPrisTot
	        'totRealialt = totRealialt
	        end if

            totRealialt = totRealialt + (proaf)
                %>


         </td>
		<td align=right class=lille style="white-space:nowrap; padding:4px 4px 0px 4px;" valign=top>
        <% for p = 1 to 10
        
        pgid = oRec("projektgruppe"&p)

        if pgid <> 1 then
            call prgNavn(pgid, 200)
            Response.Write left(prgNavnTxt, 20) & "<br>"
        end if
        
        next %>
		   
		</td>

		<td valign=top style="padding:4px 2px 4px 10px;">
		<%
		    
		    if f = 0 AND editok = 1 then
		    deleteok = 1
		    else
		    deleteok = 0
		    end if
		    
		    
		if deleteok = 1 then%>
		<a href="jobs.asp?menu=job&func=slet&id=<%=oRec("id")%>&jobnr_sog=<%=jobnr_sog%>&filt=<%=filt%>&fm_kunde_sog=<%=fmkunde%>"><%=sletkanp%></a>&nbsp;
		<%else %>
		&nbsp;
		<%end if%>
		</td>
	</tr>
	<%
	cnt = cnt + 1
	x = 0
	lastKid = oRec("kid")

    select case right(cnt,1)
    case 0
    Response.flush
    end select

    end if' forretningsområde


    
	oRec.movenext
	wend
	
	if cnt > 0 then%>
    	<tr style="background-color:#FFFFFF;">
         <td colspan=8>&nbsp;</td>
		<td>
            <%if fakturaBel_tot_gt <> 0 then %>
            <br />Faktureret total: <br /><b><%=formatnumber(fakturaBel_tot_gt, 2) & " "& basisValISO %></b>
            <%end if %>

            <%if totRealialt <> 0 AND gnsPrisTot <> 0 then %>
            <br><br>Gns. faktisk timepris:<br><b><%=formatnumber(gnsPrisTot/totRealialt, 2) &" "& basisValISO  %></b> 
            <%end if %>

		</td>
             <td colspan=11>&nbsp;</td>
        </tr>
        <tr style="background-color:#FFFFFF;">
            <td colspan=20 style="padding-right:30px; padding-top:5px;" align=right><br /><input type="submit" name="statusliste" value="Opdater liste >>"><br />&nbsp;</td>
	    </tr>
    <%end if %>
	</table>

                    
                    
                    <%end if %>

                </div>
            </div>
        </div>


    </div>
</div>
   




 

<!--#include file="../inc/regular/footer_inc.asp"-->