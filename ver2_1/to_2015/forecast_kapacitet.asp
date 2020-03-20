

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../timereg/inc/isint_func.asp"-->
<script src="js/fc_kap_list_jav.js" type="text/javascript"></script>
<script src="../timereg/inc/header_lysblaa_stat_20170808.js" type="text/javascript"></script>




<% 
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if
	
    if len(trim(request("media"))) <> 0 then
    media = request("media")
    else
    media = ""
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
    aar = year(now)
	end if

    call browsertype()

    thisfile = "forecast_kapacitet.asp"


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
    
    
    
 

    for m = 0 to UBOUND(intMids)
			    
		if m = 0 then
		medarbSQlKri = "(m.mid = " & intMids(m)
		'kundeAnsSQLKri = "kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
		'jobAnsSQLkri = "jobans1 = "& intMids(m)  
		'jobAns2SQLkri = "jobans2 = "& intMids(m)
        'jobAns3SQLkri = "jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m) 
		else
		medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
		'kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
		'jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
		'jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
        'jobAns3SQLkri = jobAns3SQLkri & " OR jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
		end if
			    
        antalm = antalm + 1  
	next
			    
	medarbSQlKri = medarbSQlKri & ")"


   '*** Antal helligDage Pr år
   helligdageIalt = 0
   'stDagAarHelligDage = "1-1-"&aar
   'for d = 0 to 365

   '     if d = 0 then
   '     tjekdennedag = stDagAarHelligDage
   '     else
   '     tjekdennedag = dateAdd("d", 1, tjekdennedag)
   '     end if

   '     if cint(year(tjekdennedag)) = cint(aar) then
   '     call helligdage(tjekdennedag, 0, lto, session("mid"))
   '     end if

   'next

    
   


    if media <> "export" then
   

    call menu_2014
%>

<div class="wrapper">
    <div class="content">
        <div class="container" style="width:75%;">
            <div class="portlet">
                <h3 class="portlet-title"><u>Forecast & Kapacitet</u></h3>
                <div class="portlet-body">
             
                   <!-- <div class="row">
                        <div class="col-lg-2">
                            <select id="FM_progrp" name="FM_progrp" class="form-control input-small" multiple="multiple" size=9>
                                <option value="0" SELECTED><%=godkendweek_txt_083 %></option>                                              
                            </select>
                        </div>
                    </div> -->
                  
                    <form action="forecast_kapacitet.asp?" method="post">
                    <table><tr>

                    <%call progrpmedarb %>
                    </tr></table>

                                  
                
                        <table>
                            <tr>
                                <td style="color:black"><b>Periode:</b></td>
                                <td style="padding-left:63px;">
                                  

                                    <select name="FM_aar" class="form-control input-small" style="width:100px" onchange="submit();">
                                        <%
                                        aarstal = 2013                                      
                                        for i = 0 to 10
                                        aarstal = aarstal + 1

                                        if cint(aarstal) = cint(aar) then
                                        optionselect = "selected"
                                        else
                                        optionselect = ""
                                        end if
                                        %>
                                        <option value="<%=aarstal%>" <%=optionselect%>><%=aarstal%></option>
                                        <%
                                       
                                        next                                         
                                        %>
                                    </select>
                                </td>
                            </tr>
                        </table>

                    </form>

                    <%
                        
                   valgtprogrp = progrp
                   valgtprogrp = Replace(valgtprogrp,"0, ","")
                   valgtprogrp = Replace(valgtprogrp,","," OR p.id = ") 
                            
                   'response.Write "medarbejdere2: "& progrp & "##" & valgtprogrp  
                    

                  
                           
                    %>

                   <!-- <a href="timbudgetsim.asp">Simulering >></a> -->

                    <br /><br />
                    <table id="fckap" style="background-color:#FFFFFF;" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th style="vertical-align:bottom; text-align:left;">Afd. (org. projektgrp.)</th>
                                <th style="vertical-align:bottom; text-align:left">Navn <br />Med.type hvis skiftet</th>
                                <th style="vertical-align:bottom; text-align:left">A. Norm.<br />/uge</th>
                                <th style="vertical-align:bottom; text-align:left">B. Ansat <br />mdr</th>
                                <th style="vertical-align:bottom; text-align:left">C. Årsnorm</th>
                                <th style="vertical-align:bottom; text-align:left">D. Ferie/FF</th>
                                <th style="vertical-align:bottom; text-align:left">E. Forecast interne timer<br /> forening (3)</th>
                                <th style="vertical-align:bottom; text-align:left">F. Forecast interne timer<br /> afdeling (7XXX) <span style="font-size:9px;">(prioritet -3)</span></th>
                                <th style="vertical-align:bottom; text-align:left">G. Faktisk <br /> syge, fravær, barsel mm.</th>
                                <th style="vertical-align:bottom; text-align:left">H. Kapacitet til projekt <br /> t/år</th>
                         
                                <th style="vertical-align:bottom; text-align:left">I. Forecast</th>
                                <th style="vertical-align:bottom; text-align:left">J. Interne timer (E+F)</th>
                                <th style="vertical-align:bottom; text-align:left">K. Projekt tid (I-J)</th>
                                <th style="vertical-align:bottom; text-align:left">L. Overskud <br />Proj. tid %</th>
                                <th style="vertical-align:bottom; text-align:left">M. <br />Timer<br />oversk.<br />/undersk.<br />(C-(D+G+I))</th>
                              
                            </tr>
                        </thead>
                        <tbody>
                            <%


            end if 'media


          


                            '*** Projektfordleing KUN EXP **'
                            if media = "export" then

                                        j = 0
                                        strTxtExportHeaderJob = ""
                                        dim antaljobids
                                        redim antaljobids(1000)
                                        strSQLress = "SELECT sum(timer) as timer, j.jobnavn, j.jobnr, r.jobid FROM ressourcer_md r "_
                                        &"LEFT JOIN job j ON (j.id = r.jobid) WHERE r.aar ="& aar & " GROUP BY r.jobid ORDER BY j.jobnr"
                                        
                                        oRec2.open strSQLress, oConn, 3
                                        while not oRec2.EOF 
                                        aarforecastprprojekt = oRec2("timer")
                                        jobnr = oRec2("jobnr")
                                        jobnavn = oRec2("jobnavn")


                                        antaljobids(j) = oRec2("jobid")
                                        strTxtExportHeaderJob = strTxtExportHeaderJob & chr(34) & jobnr &" "& jobnavn & chr(34) &";"

                                        j = j + 1
                                        oRec2.movenext
                                        wend
                                        oRec2.close


                            end if

                            totalAntalMedarbs = 0
                            totalMedarbInterneprio2 = 0
                            totalMedarbNorm = 0
                            totalMedarbFerie = 0
                            totalMedarbfravaar = 0
                            totalMedarbKapa = 0
                            totalMedarbForecast = 0
                            totalMedarbInterntid = 0
                            totalMedarbProtid = 0
                            resultskud = 0
                                
                            i = 100
                            dim  norm_ugetotalArr, intervalWeeksArr, intervalWMnavn
                            redim  norm_ugetotalArr(i), intervalWeeksArr(i), intervalWMnavn(i) 


                            strSQLSelmed = "SELECT m.mnavn, m.init, mid, ansatdato, opsagtdato FROM medarbejdere m WHERE "& medarbSQlKri &" ORDER BY mnavn"

                            'response.write "strSQLSelmed: "& strSQLSelmed
                            'response.flush
                            
                            oRec.open strSQLSelmed, oConn, 3
                            while not oRec.EOF
                            totalAntalMedarbs = totalAntalMedarbs + 1

                         


                            '**************************************************
                            '** Måneder ansat, ferie, årsnorm
                            '**************************************************

                          

                            startperiode = aar &"/01/01"
                            slutperiode = aar &"/12/31"
                            call kapcitetsnorm(oRec("mid"), oRec("ansatdato"), oRec("opsagtdato"), startperiode, slutperiode, 1) 
                            
                                
                            
                            
                                


                            '**************************************************
                            '** projektgruppe, medarbejder navn mm
                            '**************************************************
                            strProgrupppeNavn = "-"
                            strSQLprogrp = "SELECT p.navn, p.id, pr.MedarbejderId FROM progrupperelationer pr "_
                            &" LEFT JOIN projektgrupper AS p ON (p.id = pr.ProjektgruppeId AND p.orgvir = 1) "_
                            &" WHERE MedarbejderId = "& oRec("mid") &" AND p.orgvir = 1 LIMIT 1"

                            'response.write "strSQLprogrp: "& strSQLprogrp
                            'response.flush

                            oRec2.open strSQLprogrp, oConn, 3
                            while not oRec2.EOF

                                strProgrupppeNavn = oRec2("navn")

                            oRec2.movenext
                            wend
                            oRec2.close








                            if media <> "export" then                                 
                            %>
                            <tr>
                                <td style="width:100px; overflow-wrap:break-word;"><%=left(strProgrupppeNavn, 15)%></td>
                                <td style="width:200px; overflow-wrap:break-word;"><%=left(oRec("mnavn"), 20) & " ["& oRec("init") &"]" %>

                                    <%if mt > 9999999 then '0 %>
                                    <span style="font-size:9px;"><%=mtText %></span>

                                    <%end if %>

                                </td>

                                <td style="width:80px; overflow-wrap:break-word;">
                                    <%=formatnumber(norm_ugetotal,2)%>
                                    
                                </td>

                            <%else 
                                
                                strTxtExport = strTxtExport & Chr(34) & strProgrupppeNavn & Chr(34) &";"& Chr(34) & oRec("mnavn") & Chr(34) &";"& Chr(34) & oRec("init") & Chr(34) &";"& formatnumber(norm_ugetotal,2) &";" %>


                            <%end if
                                
                                
                                
                             if media <> "export" then    %>

                                <td style="width:35px;">
                                        <%
                               
                             end if

                                           

                                    if media <> "export" then
                                    response.Write maanederansat '& ": "& asDato& "#" & endDtKri & " // "& year(oRec("opsagtdato")) &"<>"& aar
                                    else
                                    strTxtExport = strTxtExport & maanederansat &";"
                                    end if


                                    if media <> "export" then
                                    %>
                                    </td>
                                    <td><%=formatnumber(norm_aarstotal,0)%>
                                        <br />
                                        <%'"maanederansat/12:" & maanederansat/12 &" * norm_ugetotal * 52:"& norm_ugetotal * 52 &" - antalhelligdagetimer:" & antalhelligdagetimer %>
                                    <%else
                                    strTxtExport = strTxtExport & formatnumber(norm_aarstotal,0) &";"
                                    end if
                                   
                                    if media <> "export" then%>
                                    </td><td>
                                    <%=formatnumber(arrsferie,2)%>
                                    <%
                                    else
                                    strTxtExport = strTxtExport & formatnumber(arrsferie,2) &";"
                                    end if
                                        
                                        
                                        
                                    if media <> "export" then%>
                                    </td><td>
                                    <%end if


                                         interneprio2 = 0
                                        'strSQLinterne = "SELECT sum(timer) as timer FROM timer WHERE tmnr ="& oRec("MedarbejderId") & " AND tjobnr = 179 AND tdato BETWEEN '2016-1-1' AND '2016-12-31'"
                                        'strSQLinterneprio2 = "SELECT sum(timer) as timer FROM timer AS t LEFT JOIN job AS j ON (j.jobnr = t.tjobnr) WHERE tmnr ="& oRec("mid") &" AND j.risiko = -2 AND taktivitetid = 11 AND tdato BETWEEN '"&aar&"-1-1' AND '"&aar&"-12-31'"
                                        'response.Write strSQLinterneprio2

                                        strSQLinterneprio2 = "SELECT sum(timer) as timer, j.jobnavn, j.jobnr, r.jobid FROM ressourcer_md r "_
                                        &"LEFT JOIN job j ON (j.id = r.jobid) WHERE r.aar ="& aar & " AND j.risiko = -2 AND r.medid = "& oRec("mid") &" AND r.aktid = 11 GROUP BY r.jobid ORDER BY j.jobnr"

                                        oRec3.open strSQLinterneprio2, oConn, 3
                                        if not oRec3.EOF then
                                        internepro3 = oRec3("timer")
                                        interneprio2 = oRec3("timer")
                                                                                    
                                        end if
                                        oRec3.close   
                                        
                                            if interneprio2 <> 0 then
                                            interneprio2 = interneprio2
                                            else
                                            interneprio2 = 0
                                            end if  
                                        
                                    if  totalAntalMedarbs = 1 then
                                    totalMedarbInterneprio2 = interneprio2 
                                    else 
                                    totalMedarbInterneprio2 = totalMedarbInterneprio2 + interneprio2 
                                    end if

                                    if media <> "export" then                                    
                                    %>
                                   <%=formatnumber(interneprio2,2) %>
                                    <%else
                                    strTxtExport = strTxtExport & formatnumber(interneprio2,2) &";" 
                                    end if
                                        
                                        
                                    if media <> "export" then       %>
                                    </td><td><!---- interne afd --->
                                   <%end if

                                        'strSQLress = "SELECT sum(timer) as timer, j.jobnavn, j.jobnr, r.jobid FROM ressourcer_md r "_
                                        '&"LEFT JOIN job j ON (j.id = r.jobid) WHERE r.aar ="& aar & " GROUP BY r.jobid ORDER BY j.jobnr"

                                        interneprio3 = 0
                                        'strSQLinterneprio3 = "SELECT sum(timer) as timer FROM timer AS t LEFT JOIN job AS j ON (j.jobnr = t.tjobnr) WHERE tmnr ="& oRec("mid") &" AND j.risiko = -3 AND tdato BETWEEN '"&aar&"-1-1' AND '"&aar&"-12-31'"
                                        'response.Write strSQLinterneprio3 
                                        
                                        strSQLinterneprio3 = "SELECT sum(timer) as timer, j.jobnavn, j.jobnr, r.jobid FROM ressourcer_md r "_
                                        &"LEFT JOIN job j ON (j.id = r.jobid) WHERE r.aar ="& aar & " AND j.risiko = -3 AND r.medid = "& oRec("mid") &" GROUP BY r.jobid ORDER BY j.jobnr"

                                        
                                        oRec3.open strSQLinterneprio3, oConn, 3
                                        if not oRec3.EOF then
                                        interneprio3 = oRec3("timer")
                                        
                                        end if
                                        oRec3.close  
                                       
                                            if interneprio3 <> 0 then
                                            interneprio3 = interneprio3
                                            else
                                            interneprio3 = 0
                                            end if
                                       
                                       
                                    if media <> "export" then                 
                                    %>
                                   <%=formatnumber(interneprio3,2) %>
                                   <%else 
                                       strTxtExport = strTxtExport & formatnumber(interneprio3,2) &";" 
                                    end if 
                                
                                    if media <> "export" then                 
                                    %>
                                    </td><td>
                                    <%end if
                                        
                                        '*+ ORLOV
                                        fravaer = 0
                                        strSQLbarsel = "SELECT sum(timer) as timer FROM timer WHERE tmnr ="& oRec("mid") & " AND (tfaktim = 22 OR tfaktim = 20 OR tfaktim = 21 OR tfaktim = 23 OR tfaktim = 24 Or tfaktim = 25) AND tdato BETWEEN '"&aar&"-1-1' AND '"&aar&"-12-31' GROUP BY tmnr"
                                        'response.Write strSQLbarsel
                                        oRec3.open strSQLbarsel, oConn, 3
                                        if not oRec3.EOF then
                                        fravaer = oRec3("timer")
                                        end if
                                        oRec3.close 

                                        if fravaer <> 0 then
                                        fravaer = fravaer
                                        else
                                        fravaer = 0
                                        end if

                                        if  totalAntalMedarbs = 1 then
                                        totalMedarbfravaar = fravaer 
                                        else 
                                        totalMedarbfravaar = totalMedarbfravaar + fravaer
                                        end if

                                    if media <> "export" then
                                    %>
                                    <%=formatnumber(fravaer,2) %>
                                    <%else
                                     strTxtExport = strTxtExport & formatnumber(fravaer,2) &";"     
                                    end if
                                        
                                    if media <> "export" then%>
                                    </td><td>
                                    <%end if

                                        'strkapacitet = norm_aarstotal - arrsferie - internepro3 - interneafd
                                        '*******************************************************
                                        '************** ÅRS KAPACITET **************************
                                        aarskapacitet = norm_aarstotal - arrsferie - interneprio2 - interneprio3 - fravaer
                                        
                                        if  totalAntalMedarbs = 1 then
                                        totalMedarbKapa = aarskapacitet
                                        else 
                                        totalMedarbKapa = totalMedarbKapa + aarskapacitet
                                        end if

                                        if media <> "export" then     
                                        response.Write formatnumber(aarskapacitet,2)  
                                        else
                                        strTxtExport = strTxtExport & formatnumber(aarskapacitet,2) &";" 
                                        end if    
                                    
                                        
                                    if media <> "export" then%>
                                    </td>

                           
                                    <%end if

                                        aarforecast = 0 
                                        resultskud = 0
                                        strSQLress = "SELECT sum(timer) as timer FROM ressourcer_md WHERE medid ="& oRec("mid") &" AND aar = "& aar &" GROUP BY medid"
                                        oRec2.open strSQLress, oConn, 3
                                        if not oRec2.EOF then
                                        aarforecast = oRec2("timer")
                                        end if
                                        oRec2.close

                                        if aarforecast <> 0 then
                                        aarforecast = aarforecast
                                        else
                                        aarforecast = 0
                                        end if

                                        interntid = (interneprio2 + interneprio3) 'arrsferie + + fravaer
                                        'interntid = interntid - (arrsferie + fravaer)

                                        'strSQLinterneprio3 = "SELECT sum(timer) as timer FROM timer AS t LEFT JOIN job AS j ON (j.jobnr = t.tjobnr) WHERE tmnr ="& oRec("mid") &" AND j.risiko = -3 AND tdato BETWEEN '"&aar&"-1-1' AND '"&aar&"-12-31'"
                                        'response.Write strSQLinterneprio3 

                                        protid = ((aarforecast/1)-(interntid/1)) 
                                        'protidGraf = aarforecast-interntid 

                                        opgjortprocent = 0
                                        opgjortprocent = aarskapacitet - protid

                                        if aarskapacitet <> 0 then
                                        opgjortprocent = (opgjortprocent * 100) / aarskapacitet
                                        else
                                        opgjortprocent = 0 
                                        end if
                                        
                                        opgjortprocent = 100 - opgjortprocent  

                                        'projekskud =  - protid
                                        resultskud = aarskapacitet - protid

                                        if  totalAntalMedarbs = 1 then
                                        totalMedarbForecast = aarforecast
                                        totalMedarbInterntid = interntid
                                        totalMedarbProtid = protid
                                        else 
                                        totalMedarbForecast = totalMedarbForecast + aarforecast
                                        totalMedarbInterntid = totalMedarbInterntid + interntid
                                        totalMedarbProtid = totalMedarbProtid + protid
                                        end if
                                        
                                    if media <> "export" then%>
                                    
                                    <td>
                                    <%=formatnumber(aarforecast,2) %>
                                    
                                    </td><td><!--Intern Tid -->
                                    <%=formatnumber(interntid,2)%>    
                                    
                                    </td><td><!--Projekt Tid I -->
                                    <%=formatnumber(protid,2)%>
                                    </td>

                                
                                    
                                    <td><!-- opgjot % L -->                                    
                                        <%=formatnumber(opgjortprocent,0) %><span style="font-size:9px;">%</span>
                                        <!--
                                    <div style="width:100%; display:inline-block">
                                    <div style="width:<%=formatnumber(opgjortprocent,0)%>%; background-color:dodgerblue; display:inline-block; text-align:right; color:#ffffff;"><%=formatnumber(opgjortprocent,0) %><span style="font-size:9px;">%</span></div>
                                    </div>-->

                                    </td>

                                    <td>

                                         <%


                                       
                                        if resultskud >= 0 then 
                                        rescolor = "green"
                                        else
                                        rescolor = "red"
                                        end if

                                        if  totalAntalMedarbs = 1 then
                                        totalMedarbResult = resultskud
                                        else 
                                        totalMedarbResult = totalMedarbResult + resultskud
                                        end if
                                    %>
                                    <div style="color:<%=rescolor%>;"><%=formatnumber(resultskud, 2) %></div></td>
                                    </tr>

                               
                                    <%

                                    else

                                        


                                    strTxtExport = strTxtExport & formatnumber(aarforecast,2) &";"
                                    strTxtExport = strTxtExport & formatnumber(interntid,2) &";"
                                    strTxtExport = strTxtExport & formatnumber(protid,2) &";"
                                    strTxtExport = strTxtExport & formatnumber(opgjortprocent,0) &";" & formatnumber(resultskud, 2) &";"
                                        
                                        '*** Fordeling pr. projekt *'
                                        aarforecastprprojekt = 0 
                                       
                                        for jj = 0 TO j - 1

                                        timerFundet = 0
                                        strSQLress = "SELECT sum(timer) as timer FROM ressourcer_md r "_
                                        &" WHERE r.medid ="& oRec("mid") &" AND r.jobid = "& antaljobids(jj) &" AND r.aar ="& aar & " GROUP BY r.jobid, r.medid"
                                        
                                        'response.write strSQLress
                                        'response.flush

                                        oRec2.open strSQLress, oConn, 3
                                        if not oRec2.EOF then
                                        aarforecastprprojekt = oRec2("timer")
                                        strTxtExport = strTxtExport & formatnumber(aarforecastprprojekt,0) &";"

                                        timerFundet = 1
                                        
                                        end if
                                        oRec2.close

                                        if cint(timerFundet) = 0 then
                                        strTxtExport = strTxtExport &";"
                                        end if


                                        next
                                        
                                        strTxtExport = strTxtExport &";"& vbcrlf

                                    end if 'media



                           


                                
		                  

                            oRec.movenext
                            wend
                            oRec.close
                                
                            'next 
                            %>
                        </tbody>

                        <tfoot>
                            <tr>
                                <th>Total</th>
                                <th>&nbsp</th>
                                <th>&nbsp</th>
                                <th>&nbsp</th>
                                <th><%=formatnumber(totalMedarbNorm,0) %></th>
                                <th><%=formatnumber(totalMedarbFerie,2) %></th>
                                <th><%=formatnumber(totalMedarbInterneprio2,2) %></th>
                                <th>&nbsp</th>
                                <th><%=formatnumber(totalMedarbfravaar,2) %></th>
                                <th><%=formatnumber(totalMedarbKapa,2) %></th>
                                <th><%=formatnumber(totalMedarbForecast,2) %></th>
                                <th><%=formatnumber(totalMedarbInterntid,2) %></th>
                                <th><%=formatnumber(totalMedarbProtid,2) %></th>
                                <th>&nbsp</th>
                                <th>
                                    <%
                                        if totalMedarbResult >= 0 then 
                                        totalRescolor = "green"
                                        else
                                        totalRescolor = "red"
                                        end if
                                        
                                    %>
                                    <div style="color:<%=totalRescolor%>;"><%=formatnumber(totalMedarbResult, 2) %></div>
                                </th>
                            </tr>
                        </tfoot>

                    </table>

                    <div class="row">
                          <div class="col-lg-12">
                            <br /><br /><br />
                                </div>
                    </div>

                    <div class="row">
                          <div class="col-lg-6">
                            <%'if lto = "sduuas" then %>
                            <script type="text/javascript" src="js/libs/excanvas.compiled.js"></script>
                            <script type="text/javascript" src="js/plugins/flot/jquery.flot.js"></script>
                            <script type="text/javascript" src="js/demos/flot/stacked-vertical_dashboard_sduuas.js"></script>
                            <div id="stacked-vertical-chart" class="chart-holder-200"></div>
                           <%'end if %>
                        </div>
                    </div>

               

                    <% '*** CSV eksport **''
                        
                if media = "export" then

                        call TimeOutVersion()
                        
                        filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                    filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	                    Set objFSO = server.createobject("Scripting.FileSystemObject")
	
	                    if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\forecast_kapacitet.asp" then
							
		                    Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\fckapacitetexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		                    Set objNewFile = nothing
		                    Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\fckapacitetexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
	
	                    else
		
		                    Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\fckapacitetexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, false)
		                    Set objNewFile = nothing
		                    Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\fckapacitetexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8, -1)
		
	                    end if
	
	                    file = "fckapacitetexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
                        
                        
                        
                        strEksportHeader = "Afd. (org. projektgrp.); "
                        strEksportHeader = strEksportHeader & "Navn; Init;"
                        strEksportHeader = strEksportHeader & "A. Norm. /uge;"
                        strEksportHeader = strEksportHeader & "B. Ansat mdr;"
                        strEksportHeader = strEksportHeader & "C. Årsnorm;"
                        strEksportHeader = strEksportHeader & "D. Ferie/FF;"
                        strEksportHeader = strEksportHeader & "E. Forecast interne timer forening (3);"
                        strEksportHeader = strEksportHeader & "F. Forecast interne timer afdeling (7XXX);"
                        strEksportHeader = strEksportHeader & "G. Faktisk syge, fravær, barsel mm.;"
                        strEksportHeader = strEksportHeader & "H. Kapacitet til projekt t/år;"
                        strEksportHeader = strEksportHeader & "I. Forecast;"
                        strEksportHeader = strEksportHeader & "J. Interne timer (E+F);"
                        strEksportHeader = strEksportHeader & "K. Projekt tid (I-J);"
                        strEksportHeader = strEksportHeader & "L. Overskud Projekttid %;"
                        strEksportHeader = strEksportHeader & "M. Timer Overskud/Underskud (C-(D+G+I));"
                        strEksportHeader = strEksportHeader & strTxtExportHeaderJob & vbcrlf
                                        
                       
                        strTxtExport = strEksportHeader & strTxtExport 
                        
                        objF.WriteLine(strTxtExport)

	                    'Response.write strTxtExport
	                    'Response.flush
	                    objF.close


                        'Response.end
                        Response.write "<div style=""background-color:#ffffff; padding:10px;""><a href='../inc/log/data/"& file &"' target='_blank' onClick=""Javascript:window.close()"";>Din .csv fil er klar >></a>"
	                    'Response.redirect "../inc/log/data/"& file &""      
                       
                        
                       end if


                     if media <> "export" then
                               
                           %>
            

             <section>
                 <br /><br />
                <div class="row">
                     <div class="col-lg-12">
                        <b><%=kunder_txt_112 %></b>
                        </div>
                    </div>
                  <form action="forecast_kapacitet.asp?media=export" method="post" target="_blank">
                      <input id="Hidden2" name="FM_aar" value="<%=aar%>" type="hidden" />
                      <input id="Hidden2" name="FM_medarb" value="<%=thisMiduse%>" type="hidden" />
                      <input id="Hidden2" name="FM_medarb_hidden" value="<%=thisMiduse%>" type="hidden" />

                     <div class="row">
                     <div class="col-lg-12 pad-r30">
                         
                    <input id="Submit5" type="submit" value=".CSV export" class="btn btn-sm" /><br />
                    <!--Eksporter viste kunder og kontaktpersoner som .csv fil-->
                         
                         </div>


                </div>
               
            </section>



                     </div>
                </div>


       </div><!-- cointainer -->
</div><!--content -->
</div><!-- wrapper -->

    <%end if 'media %>


<!--#include file="../inc/regular/footer_inc.asp"-->