

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../timereg/inc/isint_func.asp"-->
<!--include file="../inc/regular/header_lysblaa_inc.asp"-->
<!-- Header_lysblaa_inc (den game) bruges til sub progrpmedarb, finder medarbedjere efter den valgte projektgruppe -->



<% 
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if


    if len(trim(request("FM_aar"))) <> 0 then
	aar = request("FM_aar")
	else
    aar = 2015
	end if

    thisfile = "1234.asp"


    '*** Medarbejdere / projektgrupper
	'selmedarb = session("mid")
	'call medarbogprogrp("timtot")
	medarbSQlKri = ""
	kundeAnsSQLKri = ""
	jobAns2SQLkri = ""
	jobAnsSQLkri = ""
	ugeAflsMidKri = ""
	fakmedspecSQLkri = ""
	

    if len(trim(request("nomenu"))) <> 0 then
    nomenu = request("nomenu")
    else
    nomenu = 0
    end if
	
    if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
    progrp = 0
	end if
	
	'*** Rettigheder på den der er logget ind **'
	medarbid = session("mid")
    media = request("media") 'bruges ikke da printversion kaldes


     
	if len(request("FM_medarb")) <> 0 OR func = "export" then
	
	    if left(request("FM_medarb"), 1) = "0" then 'ikke længere mulig efer jq vælg alle automatisk
	        
            if media <> "print" then
	        thisMiduse = request("FM_medarb_hidden")
    	    else
    	    thisMiduse = request("FM_medarb")
    	    end if
    	
    	intMids = split(thisMiduse, ", ")
	    else
	    thisMiduse = request("FM_medarb")
	    intMids = split(thisMiduse, ", ")
	    end if
	
	else
	    
	    thisMiduse = session("mid") 
	    intMids = split(thisMiduse, ", ")
	   
	end if
	
    'fomrLNK = ""
    strFomr_rel = "0"
	strFomr_reljobids = "0"
    strFomr_relaktids = "0"
    
    
    
    if len(trim(request("FM_fomr"))) <> 0 then
    fomr = request("FM_fomr")
    'fomrLNK = fomr


    if right(fomr, 1) = "," then 
    fomr_len = len(fomr)
    fomr_left = left(fomr, fomr_len - 1)
    fomr = fomr_left
    end if

    'fomr = ""
    'response.write "<br><br><br><br><br><br>fomr: " & fomr

    fomrArr = split(fomr, ", ")

        for f = 0 TO UBOUND(fomrArr)
         
             strFomr_rel = strFomr_rel & ", #"& fomrArr(f) & "#"  

            '*** Forrretningsområder Kriteerie strFomr_rel            
            strSQLfomrrel = "SELECT for_fomr, for_jobid, for_aktid FROM fomr_rel WHERE for_fomr = " & fomrArr(f) & " GROUP BY for_jobid, for_aktid"
            oRec3.open strSQLfomrrel, oConn, 3
            while not oRec3.EOF 

            strFomr_reljobids = strFomr_reljobids & ", #"& oRec3("for_jobid") & "#" 
            strFomr_relaktids = strFomr_relaktids & ", #"& oRec3("for_aktid") & "#"
        
            oRec3.movenext
            wend 
            oRec3.close

        next

    else
    fomrArr = 0
    fomr = 0
    end if


    for m = 0 to UBOUND(intMids)
			    
		if m = 0 then
		medarbSQlKri = "(m.mid = " & intMids(m)
		kundeAnsSQLKri = "kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
		jobAnsSQLkri = "jobans1 = "& intMids(m)  
		jobAns2SQLkri = "jobans2 = "& intMids(m)
        jobAns3SQLkri = "jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m) 
		else
		medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
		kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
		jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
		jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
        jobAns3SQLkri = jobAns3SQLkri & " OR jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
		end if
			    
        antalm = antalm + 1  
	next
			    
	medarbSQlKri = medarbSQlKri & ")"


   

    call menu_2014
%>

<div class="wrapper">
    <div class="content">
        <div class="container" style="width:75%;">
            <div class="portlet">
                <h3 class="portlet-title"><u>Øko Rapport</u></h3>
                <div class="portlet-body">
                
                   <!-- <div class="row">
                        <div class="col-lg-2">
                            <select id="FM_progrp" name="FM_progrp" class="form-control input-small" multiple="multiple" size=9>
                                <option value="0" SELECTED><%=godkendweek_txt_083 %></option>                                              
                            </select>
                        </div>
                    </div> -->

                    <form action="1234.asp?" method="post">

                    <%call progrpmedarb %>


                    <br /><br /><br />                   
                   <!-- <div class="row">
                        <div class="col-lg-1"><b>Periode:</b></div>
                        <div class="col-lg-2">
                            <select name="FM_aar" class="form-control input-small">
                                <option value="2015-1-31">2015</option>
                                <option value="2016-1-31">2016</option>
                                <option value="2017-1-31">2017</option>
                            </select>
                        </div>
                    </div> -->
                        <table>
                            <tr>
                                <td style="color:black"><b>Periode:<%=aar %></b></td>
                                <td style="padding-left:63px;">
                                   <!-- <select name="FM_aar" class="form-control input-small" style="width:100px" onchange="submit();">
                                    <option value="2015-1-31">2015</option>
                                    <option value="2016-1-31">2016</option>
                                    <option value="2017-1-31">2017</option>
                                    </select> -->

                                    <select name="FM_aar" class="form-control input-small" style="width:100px" onchange="submit();">
                                        <%
                                        aarstal = 2013                                      
                                        for i = 0 to 5
                                        aarstal = aarstal + 1 
                                        if aarstal = aar then
                                        optionselect = "selected"
                                        else
                                        optionselect = ""
                                        end if
                                        strOption = "<option value="&aarstal&" "&optionselect&">"&aarstal&"</option>"

                                        response.Write strOption
                                        next                                         
                                        %>
                                    </select>
                                </td>
                            </tr>
                        </table>

                    </form>

                    <%
                        
                   valgtprogrp = progrp
                   'valgtprogrp = Replace(valgtprogrp,"0","")
                   valgtprogrp = Replace(valgtprogrp,"0,","")
                   valgtprogrp = Replace(valgtprogrp,","," OR p.id =") 
                            
                   'response.Write "medarbejdere2: " & valgtprogrp  
                     
                   for m = 0 to UBOUND(intMids) 
                   
                        'response.Write "medarb"
                        
                        strSQLmedarb = "SELECT mnavn FROM medarbejdere WHERE mid ="& intMids(m)
                        oRec.open strSQLmedarb, oConn, 3
                        if not oRec.EOF then
                        'response.Write oRec("mnavn")
                        end if
                        oRec.close

                   next
                           
                    %>

                    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th style="vertical-align:bottom; text-align:left;">Afd</th>
                                <th style="vertical-align:bottom; text-align:left">Navn</th>
                                <th style="vertical-align:bottom; text-align:left">T/uge</th>
                                <th style="vertical-align:bottom; text-align:left">Ansat <br />mdr</th>
                                <th style="vertical-align:bottom; text-align:left">Årsnorm</th>
                                <th style="vertical-align:bottom; text-align:left">Ferie/FF</th>
                                <th style="vertical-align:bottom; text-align:left">Interne timer<br /> forening</th>
                                <th style="vertical-align:bottom; text-align:left">interne timer<br /> afdeling</th>
                                <th style="vertical-align:bottom; text-align:left">Faktisk <br /> syge, fravær, barsel mm.</th>
                                <th style="vertical-align:bottom; text-align:left">Kapacitet til projekt <br /> t/år</th>
                                <th>&nbsp&nbsp</th>
                                <th style="vertical-align:bottom; text-align:left">Forecast</th>
                                <th style="vertical-align:bottom; text-align:left">Interne timer  (3) +</th>
                                <th style="vertical-align:bottom; text-align:left">Projekt tid</th>
                                <th style="vertical-align:bottom; text-align:left">Overskud/underskud %</th>
                                <th style="vertical-align:bottom; text-align:left">Overskud/underskud t.</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%

                            'for m = 0 to UBOUND(intMids)
                            'response.Write 
                            'next
                            'strSQLprogrp = "SELECT navn, mnavn FROM projektgrupper LEFT JOIN medarbejdere AS m ON (m.mid = 68) WHERE id ="& valgtprogrp
                            'SELECT "navn, p.id, pr.MedarbejderId FROM projektgrupper as p LEFT JOIN progrupperelationer AS pr ON (pr.ProjektgruppeId = p.id) WHERE p.id = 3 OR p.id = 0"
                            strSQLprogrp = "SELECT navn, p.id, pr.MedarbejderId, m.mnavn FROM projektgrupper AS p LEFT JOIN progrupperelationer AS pr ON (pr.ProjektgruppeId = p.id) LEFT JOIN medarbejdere as m ON (m.mid = pr.MedarbejderId) WHERE p.id ="& valgtprogrp

                            
                            'response.Write strSQLprogrp 
                            oRec.open strSQLprogrp, oConn, 3
                            while not oRec.EOF
                            'response.Write "Her " & oRec("MedarbejderId")
                            'response.Write oRec("mnavn")
                            'response.Write oRec("m.mid")
                            'response.Write oRec("navn")

                            'for m = 0 to UBOUND(intMids) 
                   
                            'response.Write "medarb"
                        
                            'strSQLmedarb = "SELECT mnavn FROM medarbejdere WHERE mid ="& intMids(m)
                            'oRec2.open strSQLmedarb, oConn, 3
                            'if not oRec2.EOF then
                            'medarbnavn = oRec2("mnavn")
                            'end if
                            'oRec2.close


                            'response.Write medarbnavn
                            'next
                                                                                    
                            %>
                            <tr>
                                <td><div style="width:200px;"><%=oRec("navn") %></div></td>
                                <td><div style="width:200px;"><%=oRec("mnavn") %></div></td>

                                <td><div style="width:50px;">
                                    <%
                                    varTjDatoUS_man = year(now) &"/"& month(now) &"/"& day(now)

                                    call normtimerPer(oRec("MedarbejderId"), varTjDatoUS_man, 6, 0)
                                    norm_ugetotal = ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon                                    

                                    response.Write formatnumber(norm_ugetotal,1)  
                                    %>
                                    </div></td>

                                <td><div style="width:75px;">
                                        <%
                                            dagsdato = Year(now) &"-"& Month(now) &"-"& Day(now)
                                            strSQLansat = "SELECT ansatdato FROM medarbejdere WHERE mid = "& oRec("MedarbejderId")
                                            'response.Write strSQLansat
                                            oRec2.open strSQLansat, oConn, 3
                                            if not oRec2.EOF then
                                            maanederansat = DateDiff("m",oRec2("ansatdato"),dagsdato)
                                            if maanederansat > 12 then
                                            maanederansat = 12
                                            end if
                                            response.Write formatnumber(maanederansat,1)
                                            end if
                                            oRec2.close 
                                        %>
                                    </div></td>

                                <td><div style="width:75px;"> 
                                    <%
                                        'norm_aarstotal = norm_ugetotal * 52
                                        norm_aarstotal = (maanederansat/12) * (norm_ugetotal * 52)   
                                        response.Write formatnumber(norm_aarstotal,1)
                                    %>
                                    </div></td>

                                <td><div style="width:75px;">
                                    <%
                                        arrsferie = norm_ugetotal * 6
                                        response.Write formatnumber(arrsferie,1)   
                                    %>
                                    </div></td>

                                <td><div style="width:75px;">
                                    <%
                                        'strSQLinterne = "SELECT sum(timer) as timer FROM timer WHERE tmnr ="& oRec("MedarbejderId") & " AND tjobnr = 179 AND tdato BETWEEN '2016-1-1' AND '2016-12-31'"
                                        strSQLinterneprio2 = "SELECT sum(timer) as timer FROM timer AS t LEFT JOIN job AS j ON (j.jobnr = t.tjobnr) WHERE tmnr ="& oRec("MedarbejderId") &" AND j.risiko = -2 AND tdato BETWEEN '"&aar&"-1-1' AND '"&aar&"-12-31'"
                                        'response.Write strSQLinterneprio2
                                        oRec3.open strSQLinterneprio2, oConn, 3
                                        if not oRec3.EOF then
                                        internepro3 = oRec3("timer")
                                        interneprio2 = oRec3("timer")
                                        if interneprio2 <> 0 then
                                        else
                                        interneprio2 = 0
                                        end if                                         
                                        end if
                                        oRec3.close                                         
                                    %>
                                    <%=formatnumber(interneprio2,1) %></div></td>

                               <td><div style="width:75px;"><!---- interne afd --->
                                   <%
                                        strSQLinterneprio3 = "SELECT sum(timer) as timer FROM timer AS t LEFT JOIN job AS j ON (j.jobnr = t.tjobnr) WHERE tmnr ="& oRec("MedarbejderId") &" AND j.risiko = -3 AND tdato BETWEEN '"&aar&"-1-1' AND '"&aar&"-12-31'"
                                        'response.Write strSQLinterneprio3 
                                        oRec3.open strSQLinterneprio3, oConn, 3
                                        if not oRec3.EOF then
                                        interneprio3 = oRec3("timer")
                                        if interneprio3 <> 0 then
                                        else
                                        interneprio3 = 0
                                        end if
                                        end if
                                        oRec3.close                                         
                                    %>
                                   <%=formatnumber(interneprio3,1) %></div></td>

                                <td><div style="width:75px;">
                                    <% 
                                        
                                        strSQLbarsel = "SELECT sum(timer) as timer FROM timer WHERE tmnr ="& oRec("MedarbejderId") & " AND tfaktim = 22 AND tdato BETWEEN '"&aar&"-1-1' AND '"&aar&"-12-31'"
                                        'response.Write strSQLbarsel
                                        oRec3.open strSQLbarsel, oConn, 3
                                        if not oRec3.EOF then
                                        fravaer = oRec3("timer")
                                        if fravaer <> 0 then
                                        else
                                        fravaer = 0
                                        end if
                                        end if
                                        oRec3.close 
                                    %>
                                    <%=formatnumber(fravaer,1) %></div></td>

                                <td><div style="width:75px;">
                                    <%
                                        'strkapacitet = norm_aarstotal - arrsferie - internepro3 - interneafd
                                        aarskapacitet = norm_aarstotal - arrsferie - interneprio2 - interneprio3 - fravaer     
                                        response.Write formatnumber(aarskapacitet,1)      
                                    %>
                                    </div></td>

                                <td><div style="width:75px;"></div></td>
                                <td><div style="width:75px;">
                                    <%
                                        strSQLress = "SELECT sum(timer) as timer FROM ressourcer_ramme WHERE medid ="& oRec("MedarbejderId") &" AND aar ="& aar
                                        oRec2.open strSQLress, oConn, 3
                                        if not oRec2.EOF then
                                        aarforecast = oRec2("timer")
                                        if aarforecast <> 0 then
                                        else
                                        aarforecast = 0
                                        end if
                                        end if
                                        oRec2.close
                                    %>
                                    <%=formatnumber(aarforecast,1) %></div></td>

                                <td><div style="width:75px;">
                                    <%
                                        interntid = (arrsferie + interneprio2 + fravaer + interneprio3)
                                        interntid = interntid - (arrsferie + fravaer)
                                        response.Write formatnumber(protid,1)    
                                    %>
                                    </div></td>

                                <td><div style="width:75px;"><!--Projekt Tid -->
                                    <%
                                        protid = aarforecast-interntid 
                                        response.Write formatnumber(protid,1)  
                                    %>
                                    </div></td>

                                <td>
                                    <%
                                        opgjortprocent = 0
                                        opgjortprocent = aarskapacitet - protid
                                        opgjortprocent = (opgjortprocent * 100) / aarskapacitet 
                                        opgjortprocent = 100 - opgjortprocent     
                                    %>                                    
                                    <div style="width:100%; display:inline-block">
                                    <div style="width:<%=opgjortprocent%>%; background-color:dodgerblue; display:inline-block; text-align:right"><%=formatnumber(opgjortprocent,1) %>%</div>
                                    </div></td>

                                <td>
                                    <%
                                        resultskud = aarskapacitet - protid
                                        if resultskud > 0 then 
                                        rescolor = "green"
                                        else
                                        rescolor = "red"
                                        end if
                                    %>
                                    <div style="width:75px; color:<%=rescolor%>;"><%=resultskud %></div></td>
                            </tr>
                            <%
                           
                            oRec.movenext
                            wend
                            oRec.close
                                
                            'next 
                            %>
                        </tbody>
                    </table>

                </div>
            </div>
        </div>
    </div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->