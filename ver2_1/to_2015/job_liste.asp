

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<%call menu_2014 %>

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



 <script src="../timereg/inc/job_listen_jav.js"></script>
    
 <div id="loadbar" style="position:absolute; display:; visibility:visible; top:300px; left:280px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid: 5-45 sek. <br />(ved mere end 1000 job)
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	
	</td></tr></table>

	</div>

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Joboversigt</u></h3>
                <div class="portlet-body">
                    
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
                       
                        <div class="row">
                            <div id="accordion-help" class="col-lg-4 panel-group accordion-simple">
                                <div class="panel">
                                  <div class="panel-heading">
                                    <h4 class="panel-title">
                                      <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion-help" href="#faq-general-1">Forretningsområder</a>
                                    </h4>
                                  </div><!-- .panel-heading -->

                                  <div id="faq-general-1" class="panel-collapse collapse">
                                    <div class="panel-body">
                                
                                        forret

                                    </div><!-- .panel-body -->
                                  </div> <!-- ./panel-collapse -->
                                </div><!-- .panel -->
                            </div>
                        </div>

                        <div class="row">
                            <div class="col-lg-2"><span style="color:cornflowerblue">Avanceret +</span></div>
                            <div id="avanceret_div">
                                
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
                    </div>

                </div>
            </div>
        </div>


    </div>
</div>
   




 

<!--#include file="../inc/regular/footer_inc.asp"-->