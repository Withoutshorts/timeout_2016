

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->


<div class="wrapper">
<div class="content">
<%
    thisfile = "forecast_kapacitet_graf"
    
    if len(session("user")) = 0 then	
        errortype = 5
        call showError(errortype)
        response.end
    end if
    

    if len(trim(request("FM_visning"))) <> 0 then
    grafvisning = request("FM_visning")
    else
    grafvisning = 1
    end if
    

    grafvisningSEL_aar = ""
    grafvisningSEL_halv = ""
    grafvisningSEL_kva = ""

    select case grafvisning
    case 1
        grafvisningSEL_aar = "SELECTED"
        datoFraSQL = Year(now) &"-01-01"
        datoTilSQL = Year(now + 1) & "31-12"    
    case 2
        grafvisningSEL_halv = "SELECTED"
    case 3
        grafvisningSEL_kva = "SELECTED"
    case 4
        grafvisningSEL_monthly = "SELECTED"
    end select
    
    function EndDayOfMonth(pDate)
 
            Select Case Month(pDate)
             Case 1, 3, 5, 7, 8, 10, 12
            EndDayOfMonth = "31"
            Case 4, 6, 9, 11
             EndDayOfMonth = "30"
            Case 2
             If (Year(pDate) Mod 4) = 0 Then
              EndDayOfMonth = "29"
             Else
              EndDayOfMonth = "28"
            End If
        End Select

    end function
    
    call menu_2014 

    fomr = request("FM_fomr")
    fomrArr = split(fomr, ", ")
    strfomrSQL = ""
    afslutFomrStr = 0
    for f = 0 TO UBOUND(fomrArr)   
    
        strFomr_rel = strFomr_rel & ", #"& fomrArr(f) & "#" 

        if fomrArr(f) <> 0 then
            afslutFomrStr = 1
            if f = 0 then
            'strfomrSQL = strfomrSQL + " AND (f.for_fomr = "& fomrArr(f)
            strfomrSQL = strfomrSQL + " WHERE (id = "& fomrArr(f)
            else
            'strfomrSQL = strfomrSQL + " OR f.for_fomr = "& fomrArr(f)
            strfomrSQL = strfomrSQL + " OR id = "& fomrArr(f)
            end if
        end if
    next

    if afslutFomrStr <> 0 then
        strfomrSQL = strfomrSQL + ")"
    end if

    'response.Write "<br> HerHerHerherHerHerHerherHerHerHerherHerHerHerherHerHerHerherHerHerHerherHerHerHerherHerHerHerher fomr " & strfomrSQL

    if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
    progrp = 0
	end if

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

    antalm = 0
    for m = 0 to UBOUND(intMids)			    
		if m = 0 then
		medarbSQlKri = " (mid = " & intMids(m)
		'kundeAnsSQLKri = "kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
		'jobAnsSQLkri = "jobans1 = "& intMids(m)  
		'jobAns2SQLkri = "jobans2 = "& intMids(m)
        'jobAns3SQLkri = "jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m) 
		else
		medarbSQlKri = medarbSQlKri & " OR mid = " & intMids(m)
		'kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
		'jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
		'jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
        'jobAns3SQLkri = jobAns3SQLkri & " OR jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
		end if
			    
        antalm = antalm + 1  
	next

	if antalm <> 0 then     
	    medarbSQlKri = medarbSQlKri & ")"
    end if



    totalAntalMedarbs = 1


    'response.Write "<br> medmedmedmedmedmedmedmedmedmedmedmedmedmedmedmedmedmedmdemdemdemd her - " & medarbSQlKri
%>

   <!-- <script type="text/javascript" src="js/libs/excanvas.compiled.js"></script>
    <script type="text/javascript" src="js/plugins/flot/jquery.flot.js"></script>
    <script type="text/javascript" src="js/demos/flot/stacked-vertical_fk.js"></script> -->


   <!--<script type="text/javascript" src="https://canvasjs.com/assets/script/jquery.canvasjs.min.js"></script>-->
   <!--<script src="canvasjs/canvasjs.min.js"></script> -->

    <script type="text/javascript" src="js/demos/flot/stacked-vertical_FK_2.js"></script>
    <script src="../timereg/inc/header_lysblaa_stat_20170808.js" type="text/javascript"></script>

    <style>
       #ColorCom_btn {
            text-decoration: none;
            cursor: pointer;
        }
    </style>

    <div class="container" style="width:1500px;">
        <div class="portlet">
            <h3 class="portlet-title"><u><%=forecastkap_txt_001 %> & <%=forecastkap_txt_002 %></u></h3>
            <div class="portlet-body">

                <form action="forecast_kapacitet_graf.asp" method="post">
                    <div class="well">
                        <h4 class="panel-title-well"><%=medarb_txt_098 %></h4>

                        <div class="row">
                            <div class="col-lg-2"><%=forecastkap_txt_003 %>:</div>
                            <div class="col-lg-2"><%=forecastkap_txt_004 %>:</div>
                            <div class="col-lg-3"><%=forecastkap_txt_005 %>:</div>
                            <div class="col-lg-2"><%=forecastkap_txt_006 %>:</div>
                        </div>

                        <div class="row">
                            <div class="col-lg-2">
                                <select class="form-control input-small" name="FM_visning">
                                    <option value="1" <%=grafvisningSEL_aar %> ><%=forecastkap_txt_007 %></option>
                                    <option value="2" <%=grafvisningSEL_halv %> ><%=forecastkap_txt_008 %></option>
                                    <option value="3" <%=grafvisningSEL_kva %> ><%=forecastkap_txt_009 %></option>
                                  <!-- <option value="4" <%=grafvisningSEL_monthly %> >monthly</option> -->
                                </select>
                            </div>
                            <div class="col-lg-2">
                                <select class="form-control input-small" name="FM_fomr" multiple size="9">
                                    <option value="0"><%=forecastkap_txt_010 %></option>
                                <%
                                    strSQL = "SELECT id, navn FROM fomr ORDER BY navn"
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF

                                        if instr(strFomr_rel, "#"&oRec("id")&"#") <> 0 then
                                        fSel = "SELECTED"
                                        else
                                        fSel = ""
                                        end if

                                    %>
                                        <option value="<%=oRec("id") %>" <%=fSel %> ><%=oRec("navn") %></option>
                                    <%
                                    oRec.movenext
                                    wend
                                    oRec.close
                                %>
                                </select>
                            </div>

                            <%call progrpmedarb_2018 %>

                        </div>

                        <div class="row">
                            <div class="col-lg-2" id="ColorCom_btn"><span class="badge badge-default demo-element">&nbsp</span> <b><u><%=forecastkap_txt_011 %></u></b></div>
                            <div class="col-lg-12"><button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button></div>

                        </div>

                       <div class="row" id="farvebetydninger" style="display:none;">
                            <div class="col-lg-4">

                                <%

                                    dim fomrid, fomrnavn
                                    redim fomrid(50), fomrnavn(50) 

                                    color1 = "yellow"
                                    color2 = "#ead700"
                                    color3 = "orange"
                                    color4 = "deepskyblue"
                                    color5 = "cornflowerblue"
                                    color6 = "dodgerblue"
                                    color7 = "cadetblue"
                                    color8 = "greenyellow"
                                    color9 = "lawngreen"
                                    color10 = "grey"
                                    color11 = "green"
                                    color12 = "red"
                                    color13 = "blue"
                                    color14 = "purple"
                                    color15 = "#f48042"

                                    strSQL = "SELECT id, navn FROM fomr order by id"
                                    oRec.open strSQL, oConn, 3

                                    x = 1
                                    while not oRec.EOF

                                        fomrid(x) = oRec("id")
                                        fomrnavn(x) = oRec("navn") 
                                        'response.Write "<br> id" & oRec("id") 
                                    x = x + 1
                                    oRec.movenext
                                    wend
                                    oRec.close

                                    'response.Write " X " & x 

                                    i = 1
                                    for i = 1 TO UBOUND(fomrid)
                                        
                                        select case fomrid(i)
                                        case 1
                                            fomrColor = color1
                                        case 2
                                            fomrColor = color2
                                        case 3
                                            fomrColor = color3
                                        case 4
                                            fomrColor = color4
                                        case 5
                                            fomrColor = color5
                                        case 6
                                            fomrColor = color6
                                        case 7
                                            fomrColor = color7
                                        case 8
                                            fomrColor = color8
                                        case 9
                                            fomrColor = color9
                                        case 10
                                            fomrColor = color10
                                        case 11
                                            fomrColor = color11
                                        case 12
                                            fomrColor = color12
                                        case 13
                                            fomrColor = color13
                                        case 14
                                            fomrColor = color14
                                        case 15
                                            fomrColor = color15
                                        end select

                                        'response.Write "<br> hep id " & fomrid(i) & " i " & i
                                        if i < x then
                                            response.Write "<input type='hidden' id='color_fomr_"&fomrid(i)&"' value='"&fomrColor&"' />"
                                            if i <> 1 then
                                                response.Write "<br> <span style='background-color:"&fomrColor&";' class='badge demo-element'>&nbsp</span>"& fomrnavn(i)
                                            else
                                                response.Write "<span style='background-color:"&fomrColor&";' class='badge demo-element'>&nbsp</span>"& fomrnavn(i)
                                            end if
                                        end if

                                    next

                                %>




                              <!--  <span style="background-color:grey;" class="badge demo-element">&nbsp</span> Other Tasks <br />
                                <span style="background-color:yellow;" class="badge demo-element">&nbsp</span> Phd supervision <br />                                
                                <span style="background-color:#ead700;" class="badge demo-element">&nbsp</span> Administration <br />
                                <span style="background-color:orange;" class="badge demo-element">&nbsp</span> Funding - applications & external presentations <br />
                                <span style="background-color:deepskyblue;" class="badge demo-element">&nbsp</span> Project - pipeline <br />
                                <span style="background-color:cornflowerblue;" class="badge demo-element">&nbsp</span> Projects - self-financed hours <br />
                                <span style="background-color:dodgerblue;" class="badge demo-element">&nbsp</span> Projects - project-financed hours <br />
                                <span style="background-color:cadetblue;" class="badge demo-element">&nbsp</span> Projects - Ph D <br />
                                <span style="background-color:greenyellow;" class="badge demo-element">&nbsp</span> Education - project supervision <br />
                                <span style="background-color:lawngreen;" class="badge demo-element">&nbsp</span> Education - courses -->



                                <!-- Hidden felter til graf -->
                            <!--    <input type="hidden" id="color_fomr_otherTasks" value="grey" />
                                <input type="hidden" id="color_fomr_1" value="yellow" />
                                <input type="hidden" id="color_fomr_2" value="#ead700" />
                                <input type="hidden" id="color_fomr_3" value="orange" />
                                <input type="hidden" id="color_fomr_4" value="deepskyblue" />
                                <input type="hidden" id="color_fomr_5" value="cornflowerblue" />
                                <input type="hidden" id="color_fomr_6" value="dodgerblue" />
                                <input type="hidden" id="color_fomr_7" value="cadetblue" />
                                <input type="hidden" id="color_fomr_8" value="greenyellow" />
                                <input type="hidden" id="color_fomr_9" value="lawngreen" /> -->

                            </div>
                        </div>

                    </div>
                </form>
                <br />
                <div class="row">
                    <div class="col-lg-12" style="text-align:center">
                        <h3>Total</h3> <br />
                        <div id="totalChart" style="width:100%; height:300px"></div>
                    </div>
                </div>

                <%
                  
                    public normIper
                    function normTidIPeriode(medid, startDatoPeriode, slutDatoPeriode)

                        'response.Write "func start per " & startDatoPeriode

                        startdatePerSQL = year(startDatoPeriode) &"-"& month(startDatoPeriode) &"-"& day(startDatoPeriode)
                        slutdatePerSQL = year(slutDatoPeriode) &"-"& month(slutDatoPeriode) &"-"& day(slutDatoPeriode)
                        
                        typerIperiode = 0
                        strSQL = "SELECT mtypedato FROM medarbejdertyper_historik WHERE mid = "& medid &" AND mtypedato BETWEEN '"& startdatePerSQL &"' AND '"& slutdatePerSQL &"'"
                        'response.Write strSQL
                        oRec6.open strSQL, oConn, 3
                        while not oRec6.EOF 
                             typerIperiode = typerIperiode + 1
                        oRec6.movenext
                        wend
                        oRec6.close

                        if typerIperiode = 0 then  
                    
                            strSQL = "SELECT ansatdato FROM medarbejdere WHERE mid = "& medid
                            oRec6.open strSQL, oConn, 3
                            if not oRec6.EOF then
                                ansatdato = oRec6("ansatdato")
                            end if
                            oRec6.close

                            
                            findestypehistorik = 0
                            strSQL = "SELECT mtypedato, mtype FROM medarbejdertyper_historik WHERE mid = "& medid &" AND mtypedato < '"& startdatePerSQL &"' ORDER BY mtypedato DESC, id DESC"
                            'response.Write strSQL
                            oRec5.open strSQL, oConn, 3
                            if not oRec5.EOF then
                                findestypehistorik = 1
                                
                                strSQL2 = "SELECT feriesats, normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, normtimer_son FROM medarbejdertyper WHERE id = "& oRec5("mtype")
                                oRec6.open strSQL2, oConn, 3
                                if not oRec6.EOF then
                                    ugeNorm = oRec6("normtimer_man") + oRec6("normtimer_tir") + oRec6("normtimer_ons") + oRec6("normtimer_tor") + oRec6("normtimer_fre") + oRec6("normtimer_lor") + oRec6("normtimer_son")
                                    ugediff = DateDiff("w", startdatePerSQL, slutdatePerSQL)
                                    'response.Write "wip wip iwp uge nrom " & oRec5("mtype")
                                    monthDiff = DateDiff("m", startdatePerSQL, slutdatePerSQL)
                                    totFerie = oRec6("feriesats") * monthDiff

                                    'finder helldigdage i periode
                                    holidayTimer = 0
                                    strSQL3 = "SELECT nh_date FROM national_holidays WHERE nh_date BETWEEN '"& startdatePerSQL &"' AND '"& slutdatePerSQL &"'"
                                    'response.Write strSQL3
                                    oRec3.open strSQL3, oConn, 3
                                    while not oRec3.EOF
                                        'response.Write " Heldidag " & oRec3("nh_date")
                                        holyday = weekday(oRec3("nh_date"))
                                        holydayname = weekdayname(holyday)
                                        'response.Write " DAg " & holydayname

                                        select case holydayname
                                        case "mandag"
                                            holidayTimer = holidayTimer + oRec6("normtimer_man")
                                        case "tirsdag"
                                            holidayTimer = holidayTimer + oRec6("normtimer_tir")
                                        case "onsdag"
                                            holidayTimer = holidayTimer + oRec6("normtimer_ons")
                                        case "torsdag"
                                            holidayTimer = holidayTimer + oRec6("normtimer_tor")
                                        case "fredag"
                                            holidayTimer = holidayTimer + oRec6("normtimer_fre")
                                        case "lørdag"       
                                            holidayTimer = holidayTimer + oRec6("normtimer_lor")
                                        case "søndag"
                                            holidayTimer = holidayTimer + oRec6("normtimer_son")
                                        end select

                                        'response.write "<br> holiday timer " & holidayTimer 
                                    oRec3.movenext
                                    wend
                                    oRec3.close

                                    normIper = (ugeNorm * ugediff) - totFerie
                                    normIper = normIper - holidayTimer
                                end if
                                oRec6.close                                
                            end if
                            oRec5.close

                            if findestypehistorik = 0 then 
                                if CDate(ansatdato) > CDate(startdatePerSQL) AND CDate(ansatdato) > CDate(slutdatePerSQL) then
                                    normIper = 0
                                end if

                                if CDate(ansatdato) > CDate(startdatePerSQL) AND CDate(ansatdato) < CDate(slutdatePerSQL) then
                                    antalUgermednorm = dateDiff("w", ansatdato, slutdatePerSQL)
                                    manSQLdate = ansatdato
                                    manSQLdate = DateAdd("d", 1 - WeekDay(manSQLdate, 2), manSQLdate)
                                    call normtimerPer(medid, manSQLdate, 6, 0)
                                    normIper = normIper
                                end if
                    
                                if CDate(ansatdato) < CDate(startdatePerSQL) then
                                    
                                    strSQL5 = "SELECT mtype FROM medarbejdertyper_historik WHERE mid = "& medid &" AND mtypedato = "& year(ansatdato) &"-"& month(ansatdato) &"-"& day(ansatdato)
                                    oRec5.open strSQL5, oConn, 3
                                    if not oRec5.EOF then
                                        strSQL2 = "SELECT feriesats, normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, normtimer_son FROM medarbejdertyper WHERE id = "& oRec5("mtype")
                                        oRec6.open strSQL2, oConn, 3
                                        if not oRec6.EOF then
                                            ugeNorm = oRec6("normtimer_man") + oRec6("normtimer_tir") + oRec6("normtimer_ons") + oRec6("normtimer_tor") + oRec6("normtimer_fre") + oRec6("normtimer_lor") + oRec6("normtimer_son")
                                            ugediff = DateDiff("w", startdatePerSQL, slutdatePerSQL)
                                            monthDiff = DateDiff("m", startdatePerSQL, slutdatePerSQL)
                                            totFerie = oRec6("feriesats") * monthDiff

                                            'finder helldigdage i periode
                                            holidayTimer = 0
                                            strSQL3 = "SELECT nh_date FROM national_holidays WHERE nh_date BETWEEN '"& startdatePerSQL &"' AND '"& slutdatePerSQL &"'"
                                            'response.Write strSQL3
                                            oRec3.open strSQL3, oConn, 3
                                            while not oRec3.EOF
                                                'response.Write " Heldidag " & oRec3("nh_date")
                                                holyday = weekday(oRec3("nh_date"))
                                                holydayname = weekdayname(holyday)
                                                'response.Write " DAg " & holydayname

                                                select case holydayname
                                                case "mandag"
                                                    holidayTimer = holidayTimer + oRec6("normtimer_man")
                                                case "tirsdag"
                                                    holidayTimer = holidayTimer + oRec6("normtimer_tir")
                                                case "onsdag"
                                                    holidayTimer = holidayTimer + oRec6("normtimer_ons")
                                                case "torsdag"
                                                    holidayTimer = holidayTimer + oRec6("normtimer_tor")
                                                case "fredag"
                                                    holidayTimer = holidayTimer + oRec6("normtimer_fre")
                                                case "lørdag"       
                                                    holidayTimer = holidayTimer + oRec6("normtimer_lor")
                                                case "søndag"
                                                    holidayTimer = holidayTimer + oRec6("normtimer_son")
                                                end select

                                                'response.write "<br> holiday timer " & holidayTimer 
                                            oRec3.movenext
                                            wend
                                            oRec3.close


                                            normIper = (ugeNorm * ugediff) - totFerie 
                                            normIper = normIper - holidayTimer 

                                        end if
                                        oRec6.close
                                    end if
                                    oRec5.close
                                        
                                end if
                            end if

                        end if


                        if (typerIperiode <> 0 AND findestypehistorik = 0) OR (CDate(ansatdato) > cDate(startdatePerSQL) AND findestypehistorik = 0) then                            
                            dageiperiode = DateDiff("d", startdatePerSQL, slutdatePerSQL)
                             
                            call normtimerPer(medid, startdatePerSQL, dageiperiode, 0)
                            normIper = ntimPer                            
                        end if

                        'response.Write "<br> " & "normen i perioden er udregnet til " & normIper & "<br>" 
                        
                        

                    end function
                %>

                <div class="row">
                <%

                    i = 10
                    dim  norm_ugetotalArr, intervalWeeksArr, intervalWMnavn
                    redim  norm_ugetotalArr(i), intervalWeeksArr(i), intervalWMnavn(i)

                    dim totalReg, totalFc
                    redim totalReg(50,4), totalFc(50,4) 'Fomr, periode

                    totalReg_per1 = 0
                    totalReg_per2 = 0
                    totalReg_per3 = 0
                    totalReg_per4 = 0

                    totalFc_per1 = 0
                    totalFc_per2 = 0
                    totalFc_per3 = 0
                    totalFc_per4 = 0

                    totalnorm_per1 = 0
                    totalnorm_per2 = 0
                    totalnorm_per3 = 0
                    totalnorm_per4 = 0

                    strSQL = "SELECT mid, mnavn, mnr, ansatdato, opsagtdato FROM medarbejdere WHERE"& medarbSQlKri &" ORDER BY mnavn"
                    'response.Write strSQL 
                    oRec.open strSQL, oConn, 3
                    while not oRec.EOF

                        normDatoFra = year(now) &"-01-01"
                        manSQLdate = year(now) &"-"& month(now) &"-"& day(now)
                        manSQLdate = DateAdd("d", 1 - WeekDay(manSQLdate, 2), manSQLdate)

                        'dageIPeriode = DateDiff("d", (year(now)&"-01-01"), ((year(now) + 3)&"-12-31"))
                        'call normtimerPer(oRec("mid"), manSQLdate, 6, 0)                      
                        'norm_ugetotal = ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon
                        'feriefradrag = 150
                        'ntimPer = ntimPer - feriefradrag
                        'ntimPer = ntimPer * 52

                        'response.Write "<br> " & oRec("mid")

                        'response.Write "<br><br> ¤ års norm herherherherherh - " & (ntimPer / 4) & "Uge norm " & norm_ugetotal & " dato mandag " & manSQLdate & " MEDarbejder " & oRec("mid")
                        
                        'response.Write "<br> Medarb " & oRec("medid") & " Timer " & oRec("timer")
                %>

                        <%

                            sqlYearST = year(datoFraSQL)
                            sqldatoST = year(now) &"-"& month(now) &"-01"
                            
                            sqldatoHalvAar = year(now) &"-01-01"
                            sqldatoHalvAarSlut = DateAdd("m", 6, sqldatoHalvAar)
                            sqldatoHalvAarSlut = DateAdd("d", -1, sqldatoHalvAarSlut)

                            sqldatoQAST = year(now) &"-01-01"
                            sqldatoQASL = DateAdd("q", 1, sqldatoQAST)
                            sqldatoQASL = DateADd("d", -1, sqldatoQASL)
                            
                            i = 0 
                            while i < 4
                            'response.Write "<br>" & sqldatoST
                            timer_for_fomr1 = 0
                            timer_for_fomr2 = 0
                            timer_for_fomr3 = 0
                            timer_for_fomr4 = 0
                            timer_for_fomr5 = 0
                            timer_for_fomr6 = 0
                            timer_for_fomr7 = 0
                            timer_for_fomr8 = 0
                            timer_for_fomr9 = 0
                            timer_for_fomr10 = 0
                            timer_for_fomr11 = 0
                            timer_for_fomr12 = 0
                            timer_for_fomr13 = 0
                            timer_for_fomr14 = 0
                            timer_for_fomr15 = 0


                            reg_timer_fomr1 = 0
                            reg_timer_fomr2 = 0
                            reg_timer_fomr3 = 0
                            reg_timer_fomr4 = 0
                            reg_timer_fomr5 = 0
                            reg_timer_fomr6 = 0
                            reg_timer_fomr7 = 0
                            reg_timer_fomr8 = 0
                            reg_timer_fomr9 = 0
                            reg_timer_fomr10 = 0
                            reg_timer_fomr11 = 0
                            reg_timer_fomr12 = 0
                            reg_timer_fomr13 = 0
                            reg_timer_fomr14 = 0
                            reg_timer_fomr15 = 0

                            normforperiode = 0

                            select case cint(grafvisning)
                            case 1
                                '******* Årligt *********
                                'response.Write "<br> År valgt " & sqlYearST
                                periodefiltler_forecasttimer_SQL = " AND md BETWEEN 1 AND 12 AND aar =" & sqlYearST
                                periodefiltler_regtimer_SQL = " AND tdato BETWEEN '"&sqlYearST&"-01-01' AND '"&sqlYearST&"-12-31'" 
                                underskrift = right(sqlYearST,2) 


                                functionSt = sqlYearST &"-01-01"
                                functionSl = sqlYearST &"-12-31"
                                
                                call kapcitetsnorm(oRec("mid"), oRec("ansatdato"), oRec("opsagtdato"), functionSt, functionSl, 1)

                                'response.Write "<br> Norm i per " & totalMedarbNorm & "<br>"
                                
                                  
                                'call normTidIPeriode(oRec("mid"), functionSt, functionSl)
                                'response.Write "<br><br> Norm i periode = " & normIper
                                'ntimPer = normIper
                                'manSQLdate = sqlYearST &"-01-01"
                                'manSQLdate = DateAdd("d", 1 - WeekDay(manSQLdate, 2), manSQLdate)

                                'strSQL = "SELECT mtypedato FROM medarbejdertyper_historik WHERE mid = "& oRec("mid") &" AND mtypedato BETWEEN '"&sqlYearST&"-01-01' AND '"&sqlYearST&"-12-31'"
                                'response.Write strSQL
                                'oRec2.open strSQL, oConn, 3
                                'if not oRec2.EOF then 
                                 '   response.Write "<br> Ny Type Dato " & oRec2("mtypedato") 
                                'end if
                                'oRec2.close
                                
                               ' call normtimerPer(oRec("mid"), manSQLdate, 6, 1)

                               ' ugerIperiodeFraDato = "2018-01-01"
                               ' ugerIperiodeTilDato = "2018-12-31"

                               ' ugerIPeriode = DateDiff("w", ugerIperiodeFraDato, ugerIperiodeTilDato)
                               ' feriefradrag = 150
                               ' ntimPer = nTimerPerIgnHellig
                                'ntimPer = ntimPer - feriefradrag
                               ' ntimPer = ntimPer * ugerIPeriode

                               ' response.Write "<br> normtimer herher -- " & ntimPer & " i År " & sqlYearST & " uger I periode " & ugerIPeriode & " Uge Nprm " & nTimerPerIgnHellig
                                
                                'dageIPeriode = DateDiff("d", (year(sqlYearST)&"-01-01"), (year(sqlYearST)&"-12-31")) 
                                'response.Write "<br> dage i per " & dageIPeriode
                                'normDatoFra = year(now) &"-01-01"
                                'call normtimerPer(oRec("mid"), normDatoFra, dageIPeriode, 0)
                                'response.Write "<br> ntimper " & ntimPer
                                'response.Write "<br> rrrrrr" & sqlYearST&"-01-01 - " & sqlYearST&"-12-31" 
                            case 2
                                '******* Halvårligt *********
                                'response.Write "<br> her - " & sqldatoHalvAar & "-" & sqldatoHalvAarSlut
                                periodefiltler_forecasttimer_SQL = " AND md BETWEEN "& month(sqldatoHalvAar) & " AND "& Month(sqldatoHalvAarSlut) & " AND aar =" & year(sqldatoHalvAarSlut)
                                periodefiltler_regtimer_SQL = " AND tdato BETWEEN '"& year(sqldatoHalvAar) &"-"& month(sqldatoHalvAar) &"-01' AND '"& year(sqldatoHalvAarSlut) &"-"& month(sqldatoHalvAarSlut) &"-"& day(sqldatoHalvAarSlut) &"'"
                            
                                select case month(sqldatoHalvAar)
                                case 1,2,3,4,5,6
                                underskrift = "H1 - " & year(sqldatoHalvAar)
                                case else
                                underskrift = "H2 - " & year(sqldatoHalvAar)
                                end select

                                manSQLdate = sqldatoHalvAar
                                manSQLdate = DateAdd("d", 1 - WeekDay(manSQLdate, 2), manSQLdate)

                                'response.Write "<br> herherher1234555555555" & sqldatoHalvAar &" - "& manSQLdate  
                                
                                dageIPeriode = DateDiff("d", (year(sqldatoHalvAar) &"-"& month(sqldatoHalvAar) &"-"& day(sqldatoHalvAar)), (year(sqldatoHalvAarSlut) &"-"& month(sqldatoHalvAarSlut) &"-"& day(sqldatoHalvAarSlut))) 
                                    
                                functionSt = year(sqldatoHalvAar) &"-"& month(sqldatoHalvAar) &"-"& day(sqldatoHalvAar)
                                functionSl = year(sqldatoHalvAarSlut) &"-"& month(sqldatoHalvAarSlut) &"-"& day(sqldatoHalvAarSlut)

                                call kapcitetsnorm(oRec("mid"), oRec("ansatdato"), oRec("opsagtdato"), functionSt, functionSl, 2)
                                'response.Write "<br> Norm i per " & totalMedarbNorm & "<br>"

                                'call normTidIPeriode(oRec("mid"), functionSt, functionSl)
                                'response.Write "<br> dage i per " & dageIPeriode
                                'normDatoFra = year(sqldatoHalvAar) &"-"& month(sqldatoHalvAar) &"-"& day(sqldatoHalvAar) 
                                'call normtimerPer(oRec("mid"), normDatoFra, dageIPeriode, 0)
                                'response.Write "<br> ntimper " & ntimPer & "år " & year(sqldatoHalvAar)
                                'ntimPer = normIper
                            case 3
                                '******* Kvartal *********
                            'response.Write "<br> qqqqq " & sqldatoQAST &" - "& sqldatoQASL
                                periodefiltler_forecasttimer_SQL = " AND md BETWEEN "& month(sqldatoQAST) & " AND "& Month(sqldatoQASL) & " AND aar = "& year(sqldatoQASL)
                                periodefiltler_regtimer_SQL = " AND tdato BETWEEN '"& year(sqldatoQAST) &"-"& month(sqldatoQAST) &"-01' AND '"& year(sqldatoQASL) &"-"& month(sqldatoQASL) &"-"& day(sqldatoQASL) &"'"
                                
                                select case month(sqldatoQAST)
                                case 1,2,3
                                underskrift = "Q1 " & right(year(sqldatoQAST),2)
                                case 4,5,6
                                underskrift = "Q2 " & right(year(sqldatoQAST),2)
                                case 7,8,9
                                underskrift = "Q3 " & right(year(sqldatoQAST),2)
                                case 10,11,12
                                underskrift = "Q4 " & right(year(sqldatoQAST),2)
                                end select

                                dageIPeriode = DateDiff("d", (year(sqldatoQAST) &"-"& month(sqldatoQAST) &"-"& day(sqldatoQAST)), (year(sqldatoQASL) &"-"& month(sqldatoQASL) &"-"& day(sqldatoQASL))) 
                                
                                functionSt = year(sqldatoQAST) &"-"& month(sqldatoQAST) &"-"& day(sqldatoQAST)
                                functionSl = year(sqldatoQASL) &"-"& month(sqldatoQASL) &"-"& day(sqldatoQASL)

                                call kapcitetsnorm(oRec("mid"), oRec("ansatdato"), oRec("opsagtdato"), functionSt, functionSl, 3)
                                'response.Write "<br> Norm i per " & totalMedarbNorm & "<br>"
                                'call normTidIPeriode(oRec("mid"), functionSt, functionSl)
                                'response.Write "<br> dage i per " & dageIPeriode
                                'normDatoFra = year(sqldatoQAST) &"-"& month(sqldatoQAST) &"-"& day(sqldatoQAST) 
                                'call normtimerPer(oRec("mid"), normDatoFra, dageIPeriode, 0)
                                'response.Write "<br> ntimper " & ntimPer
                                'ntimPer = normIper
                            case 4
                                '******* Monthly *********
                                periodefiltler_forecasttimer_SQL = " AND md = "& month(sqldatoST)
                                periodefiltler_regtimer_SQL = " AND tdato BETWEEN '"& year(sqldatoST) &"-"& month(sqldatoST) &"-01' AND '"& year(sqldatoST) &"-"& month(sqldatoST) &"-"& EndDayOfMonth(sqldatoST) &"'"
                                underskrift = left(monthname(month(sqldatoST)),2)

                                dageIPeriode = DateDiff("d", (year(sqldatoST) &"-"& month(sqldatoST) &"-"& day(sqldatoST)), (year(sqldatoST) &"-"& month(sqldatoST) &"-"& EndDayOfMonth(sqldatoST))) 
                                
                                functionSt = year(sqldatoST) &"-"& month(sqldatoST) &"-"& day(sqldatoST)
                                functionSl = year(sqldatoST) &"-"& month(sqldatoST) &"-"& EndDayOfMonth(sqldatoST)

                                call kapcitetsnorm(oRec("mid"), oRec("ansatdato"), oRec("opsagtdato"), functionSt, functionSl, 3)
                                'response.Write "<br> Norm i per " & totalMedarbNorm & " uger i per " & antalUger & " --" & " medarb " & oRec("mid") & "<br>"

                                'call normTidIPeriode(oRec("mid"), functionSt, functionSl)
                                'response.Write "<br> dage i per " & dageIPeriode
                                'normDatoFra = year(sqldatoST) &"-"& month(sqldatoST) &"-"& EndDayOfMonth(sqldatoST) 
                                'call normtimerPer(oRec("mid"), normDatoFra, dageIPeriode, 0)
                                'response.Write "<br> ntimper " & ntimPer
                                'ntimPer = normIper
                            end select

                            ntimPer = totalMedarbNorm

                            select case cint(i)
                                case 0
                                    totalnorm_per1 = totalnorm_per1 + ntimPer
                                case 1
                                    totalnorm_per2 = totalnorm_per2 + ntimPer
                                case 2
                                    totalnorm_per3 = totalnorm_per3 + ntimPer
                                case 3
                                    totalnorm_per4 = totalnorm_per4 + ntimPer
                                case 4
                                    totalnorm_per5 = totalnorm_per5 + ntimPer
                            end select

                            'response.Write "<br> norm i per " & ntimPer
                            'response.Write "<br> herher" & sqlYearST
                            'strSQL2 = "SELECT timer, f.for_fomr as for_fomr FROM ressourcer_md LEFT JOIN fomr_rel as f ON (jobid = f.for_jobid) WHERE medid = "& oRec("mid") & " AND for_aktid = 0" & strfomrSQL & periodefiltler_forecasttimer_SQL
                            'response.Write "<br>" & strSQL2
                            'oRec2.open strSQL2, oConn, 3
                            'while not oRec2.EOF

                                'select case oRec2("for_fomr")
                                'case 1
                                'timer_for_fomr1 = timer_for_fomr1 + oRec2("timer")
                                'case 2
                                'timer_for_fomr2 = timer_for_fomr2 + oRec2("timer")
                                'case 3
                                'timer_for_fomr3 = timer_for_fomr3 + oRec2("timer")
                                'case 4
                                'timer_for_fomr4 = timer_for_fomr4 + oRec2("timer")
                                'case 5
                                'timer_for_fomr5 = timer_for_fomr5 + oRec2("timer")
                                'case 6
                                'timer_for_fomr6 = timer_for_fomr6 + oRec2("timer")
                                'case 7
                                'timer_for_fomr7 = timer_for_fomr7 + oRec2("timer")
                                'case 8
                                'timer_for_fomr8 = timer_for_fomr8 + oRec2("timer")
                                'case 9
                                'timer_for_fomr9 = timer_for_fomr9 + oRec2("timer")
                                'case 10
                                'timer_for_fomr10 = timer_for_fomr10 + oRec2("timer")
                                'case 11
                                'timer_for_fomr11 = timer_for_fomr11 + oRec2("timer")
                                'case 12
                                'timer_for_fomr12 = timer_for_fomr12 + oRec2("timer")
                                'case 13
                                'timer_for_fomr13 = timer_for_fomr13 + oRec2("timer")
                                'case 14
                                'timer_for_fomr14 = timer_for_fomr14 + oRec2("timer")
                                'case 15
                                'timer_for_fomr15 = timer_for_fomr15 + oRec2("timer")
                                'end select
                                'response.Flush
                            'oRec2.movenext
                            'wend
                            'oRec2.close 
                            
                            
                            Dim fomr, aktSQLWHERE, fomrRegTimer, fomrFCTimer, aktFCSQLWHERE
                            Redim fomr(50), aktSQLWHERE(50), fomrRegTimer(50), fomrFCTimer(50), aktFCSQLWHERE(50)
                            strSQL = "SELECT id, navn FROM fomr "& strfomrSQL &" ORDER BY id"
                            oRec3.open strSQL, oConn, 3
                            fo = 0
                            while not oRec3.EOF
                                fomr(fo) = oRec3("id")
                                response.Write "<input type='hidden' value='"&oRec3("navn")&"' id='fomrnavn_"&fo&"' />"
                            fo = fo + 1
                            oRec3.movenext
                            wend
                            oRec3.close

                            for f = 0 TO fo
                                if fomr(f) <> "" then
                                    strSQLAkt = "SELECT for_aktid FROM fomr_rel WHERE for_fomr = "& fomr(f) & " AND for_aktid <> 0"
                                    'response.Write "<br>" & strSQLAkt
                                    oRec3.open strSQLAkt, oConn, 3
                                    a = 0
                                    while not oRec3.EOF
                                        if a = 0 then
                                        aktSQLWHERE(f) = " AND (taktivitetid = "& oRec3("for_aktid")
                                        aktFCSQLWHERE(f) = " AND (aktid = " & oRec3("for_aktid")
                                        else
                                        aktSQLWHERE(f) = aktSQLWHERE(f) & " OR taktivitetid = "& oRec3("for_aktid")
                                        aktFCSQLWHERE(f) = aktFCSQLWHERE(f) & " OR aktid = "& oRec3("for_aktid")
                                        end if
                                    
                                    a = a + 1
                                    oRec3.movenext
                                    wend
                                    oRec3.close

                                    if a <> 0 then
                                        aktSQLWHERE(f) = aktSQLWHERE(f) & ")"
                                        aktFCSQLWHERE(f) = aktFCSQLWHERE(f) & ")"
                                    end if
                                end if

                            next

                            for f2 = 0 TO fo
                                'response.Write "<br><br> FOR " & fomr(f2) & " WHERE " & aktSQLWHERE(f2)
                                fomrRegTimer(f2) = 0
                                if aktSQLWHERE(f2) <> "" then
                                
                                    strSQLTimer = "SELECT sum(timer) as sumtimer FROM timer WHERE tmnr = "& oRec("mid") & periodefiltler_regtimer_SQL & aktSQLWHERE(f2)
                                    oRec3.open strSQLTimer, oConn, 3
                                    if not oRec.EOF then
                                        fomrRegTimer(f2) = oRec3("sumtimer")
                                        if isnull(oRec3("sumtimer")) = false then

                                            totalReg(f2,i) = totalReg(f2,i) + oRec3("sumtimer")

                                            select case cint(i)
                                                case 0
                                                    totalReg_per1 = totalReg_per1 + oRec3("sumtimer")
                                                case 1
                                                    totalReg_per2 = totalReg_per2 + oRec3("sumtimer")
                                                case 2
                                                    totalReg_per3 = totalReg_per3 + oRec3("sumtimer")
                                                case 3
                                                    totalReg_per4 = totalReg_per4 + oRec3("sumtimer")   
                                            end select
                                        end if
                                    oRec3.close
                                    end if
                                end if
                                

                                fomrFCTimer(f2) = 0
                                if aktFCSQLWHERE(f2) <> "" then
                                    strSQLFCTimer = "SELECT sum(timer) as sumtimer FROM ressourcer_md WHERE medid = "& oRec("mid") & periodefiltler_forecasttimer_SQL & aktFCSQLWHERE(f2)
                                    'response.Write strSQLFCTimer & "<br>"
                                    oRec3.open strSQLFCTimer, oConn, 3
                                    if not oRec3.EOF then
                                        fomrFCTimer(f2) = oRec3("sumtimer")
                                        if isNull(oRec3("sumtimer")) = false then
                                            totalFc(f2,i) = totalFc(f2,i) + oRec3("sumtimer")

                                            select case cint(i)
                                                case 0
                                                    totalFc_per1 = totalFc_per1 + oRec3("sumtimer")
                                                case 1
                                                    totalFc_per2 = totalFc_per2 + oRec3("sumtimer")
                                                case 2
                                                    totalFc_per3 = totalFc_per3 + oRec3("sumtimer")
                                                case 3
                                                    totalFc_per4 = totalFc_per4 + oRec3("sumtimer")
                                            end select
                                        end if
                                    end if
                                    oRec3.close
                                end if
                                'response.Write "<br> " & fomrFCTimer(f2)
                            next

                            timer_for_fomr1 = fomrFCTimer(0)
                            timer_for_fomr2 = fomrFCTimer(1)
                            timer_for_fomr3 = fomrFCTimer(2)
                            timer_for_fomr4 = fomrFCTimer(3)
                            timer_for_fomr5 = fomrFCTimer(4)
                            timer_for_fomr6 = fomrFCTimer(5)
                            timer_for_fomr7 = fomrFCTimer(6)
                            timer_for_fomr8 = fomrFCTimer(7)
                            timer_for_fomr9 = fomrFCTimer(8)
                            timer_for_fomr10 = fomrFCTimer(9)
                            timer_for_fomr11 = fomrFCTimer(10)
                            timer_for_fomr12 = fomrFCTimer(11)
                            timer_for_fomr13 = fomrFCTimer(12)
                            timer_for_fomr14 = fomrFCTimer(13)
                            timer_for_fomr15 = fomrFCTimer(14)

                            reg_timer_fomr1 = fomrRegTimer(0)
                            reg_timer_fomr2 = fomrRegTimer(1)
                            reg_timer_fomr3 = fomrRegTimer(2)
                            reg_timer_fomr4 = fomrRegTimer(3)
                            reg_timer_fomr5 = fomrRegTimer(4)
                            reg_timer_fomr6 = fomrRegTimer(5)
                            reg_timer_fomr7 = fomrRegTimer(6)
                            reg_timer_fomr8 = fomrRegTimer(7)
                            reg_timer_fomr9 = fomrRegTimer(8)
                            reg_timer_fomr10 = fomrRegTimer(9)
                            reg_timer_fomr11 = fomrRegTimer(10)
                            reg_timer_fomr12 = fomrRegTimer(11)
                            reg_timer_fomr13 = fomrRegTimer(12)
                            reg_timer_fomr14 = fomrRegTimer(13)
                            reg_timer_fomr15 = fomrRegTimer(14)

                            'response.Write "<bR> REGTIMER på " & fomr(0) & " TIMER " &  reg_timer_fomr1
                            'response.Write "<bR> REGTIMER på " & fomr(1) & " TIMER " &  reg_timer_fomr2
                            'response.Write "<bR>REGTIMER på " & fomr(2) & " TIMER " &  reg_timer_fomr3
                            'response.Write "<bR>REGTIMER på " & fomr(3) & " TIMER " &  reg_timer_fomr4
                            'response.Write "<bR>REGTIMER på " & fomr(4) & " TIMER " &  reg_timer_fomr5
                            'response.Write "<bR>REGTIMER på " & fomr(5) & " TIMER " &  reg_timer_fomr6
                            'response.Write "<bR>REGTIMER på " & fomr(6) & " TIMER " &  reg_timer_fomr7
                            'response.Write "<bR>REGTIMER på " & fomr(7) & " TIMER " &  reg_timer_fomr8
                            'response.Write "<bR>REGTIMER på " & fomr(8) & " TIMER " &  reg_timer_fomr9
                            'response.Write "<bR>REGTIMER på " & fomr(9) & " TIMER " &  reg_timer_fomr10
                            'response.Write "<bR>REGTIMER på " & fomr(10) & " TIMER" &  reg_timer_fomr11
                            'response.Write "<bR>REGTIMER på " & fomr(11) & " TIMER " &  reg_timer_fomr12
                            'response.Write "<bR>REGTIMER på " & fomr(12) & " TIMER " &  reg_timer_fomr13
                            'response.Write "<bR>REGTIMER på " & fomr(13) & " TIMER " &  reg_timer_fomr14
                            'response.Write "<bR>REGTIMER på " & fomr(14) & " TIMER " &  reg_timer_fomr15

                            'RESPONSE.Write "<br><br>"

                            'response.Write "<bR> Forecast på " & fomr(0) & " TIMER " &  timer_for_fomr1
                            'response.Write "<bR> Forecast på " & fomr(1) & " TIMER " &  timer_for_fomr2
                            'response.Write "<bR>Forecast på " & fomr(2) & " TIMER " &  timer_for_fomr3
                            'response.Write "<bR>Forecast på " & fomr(3) & " TIMER " &  timer_for_fomr4
                            'response.Write "<bR>Forecast på " & fomr(4) & " TIMER " &  timer_for_fomr5
                            'response.Write "<bR>Forecast på " & fomr(5) & " TIMER " &  timer_for_fomr6
                            'response.Write "<bR>Forecast på " & fomr(6) & " TIMER " &  timer_for_fomr7
                            'response.Write "<bR>Forecast på " & fomr(7) & " TIMER " &  timer_for_fomr8
                            'response.Write "<bR>Forecast på " & fomr(8) & " TIMER " &  timer_for_fomr9
                            'response.Write "<bR>Forecast på " & fomr(9) & " TIMER " &  timer_for_fomr10
                            'response.Write "<bR>Forecast på " & fomr(10) & " TIMER" &  timer_for_fomr11
                            'response.Write "<bR>Forecast på " & fomr(11) & " TIMER " &  timer_for_fomr12
                            'response.Write "<bR>Forecast på " & fomr(12) & " TIMER " &  timer_for_fomr13
                            'response.Write "<bR>Forecast på " & fomr(13) & " TIMER " &  timer_for_fomr14
                            'response.Write "<bR>Forecast på " & fomr(14) & " TIMER " &  timer_for_fomr15

                            'strSQL3 = "SELECT timer, f.for_fomr as for_fomnr FROM timer LEFT JOIN job as j ON (tjobnr = jobnr) LEFT JOIN fomr_rel as f ON (j.id = f.for_jobid) WHERE tmnr = "& oRec("mid") & " AND for_aktid = 0" & strfomrSQL & periodefiltler_regtimer_SQL &" ORDER BY f.for_fomr"
                            'response.Write "<br> her: " & strSQL3 &" medid - "& oRec("mid")
                            'oRec3.open strSQL3, oConn, 3
                            'while not oRec3.EOF

                                'select case oRec3("for_fomnr")
                                'case 1
                                'reg_timer_fomr1 = reg_timer_fomr1 + oRec3("timer")
                                'case 2
                                'reg_timer_fomr2 = reg_timer_fomr2 + oRec3("timer")
                                'case 3
                                'reg_timer_fomr3 = reg_timer_fomr3 + oRec3("timer")
                                'case 4
                                'reg_timer_fomr4 = reg_timer_fomr4 + oRec3("timer")
                                'case 5
                                'reg_timer_fomr5 = reg_timer_fomr5 + oRec3("timer")
                                'case 6
                                'reg_timer_fomr6 = reg_timer_fomr6 + oRec3("timer")
                                'case 7
                                'reg_timer_fomr7 = reg_timer_fomr7 + oRec3("timer")
                                'case 8
                                'reg_timer_fomr8 = reg_timer_fomr8 + oRec3("timer")
                                'case 9
                                'reg_timer_fomr9 = reg_timer_fomr9 + oRec3("timer")
                                'case 10
                                'reg_timer_fomr10 = reg_timer_fomr10 + oRec3("timer")
                                'case 11
                                'reg_timer_fomr11 = reg_timer_fomr11 + oRec3("timer")
                                'case 12
                                'reg_timer_fomr12 = reg_timer_fomr12 + oRec3("timer")
                                'case 13
                                'reg_timer_fomr13 = reg_timer_fomr13 + oRec3("timer")
                                'case 14
                                'reg_timer_fomr14 = reg_timer_fomr14 + oRec3("timer")
                                'case 15
                                'reg_timer_fomr15 = reg_timer_fomr15 + oRec3("timer")
                                'end select

                                'response.Flush
                            'oRec3.movenext
                            'wend
                            'oRec3.close

                            'response.Write "<br> ntimper " & ntimper
                            'response.Write "<br> timer forecast 1 " & timer_for_fomr1 & " i " & i
                            'response.Write "<br> timer forecast 2 " & timer_for_fomr2 & " i " & i
                            'response.Write "<br> timer forecast 3 " & timer_for_fomr3 & " i " & i
                            'response.Write "<br> timer forecast 4 " & timer_for_fomr4 & " i " & i
                            'response.Write "<br> timer forecast 5 " & timer_for_fomr5 & " i " & i
                            'response.Write "<br> timer forecast 6 " & timer_for_fomr6 & " i " & i
                            'response.Write "<br> timer forecast 7 " & timer_for_fomr7 & " i " & i
                            'response.Write "<br> timer forecast 8 " & timer_for_fomr8 & " i " & i
                            'response.Write "<br> timer forecast 9 " & timer_for_fomr9 & " i " & i

                            'response.Write "<br> timer reg 1 " & reg_timer_fomr1 & " i " & i & " medid " & oRec("mid")
                            'response.Write "<br> timer reg 2 " & reg_timer_fomr2 & " i " & i & " medid " & oRec("mid")
                            'response.Write "<br> timer reg 3 " & reg_timer_fomr3 & " i " & i & " medid " & oRec("mid")
                            'response.Write "<br> timer reg 4 " & reg_timer_fomr4 & " i " & i & " medid " & oRec("mid")
                            'response.Write "<br> timer reg 5 " & reg_timer_fomr5 & " i " & i & " medid " & oRec("mid")
                            'response.Write "<br> timer reg 6 " & reg_timer_fomr6 & " i " & i & " medid " & oRec("mid")
                            'response.Write "<br> timer reg 7 " & reg_timer_fomr7 & " i " & i & " medid " & oRec("mid")
                            'response.Write "<br> timer reg 8 " & reg_timer_fomr8 & " i " & i & " medid " & oRec("mid")
                            'response.Write "<br> timer reg 9 " & reg_timer_fomr9 & " i " & i & " medid " & oRec("mid")
                            'response.Write "<br> timer reg 10 " & reg_timer_fomr10 & " i " & i & " medid " & oRec("mid")

                            i = i + 1
                            sqlYearST = sqlYearST + 1
                            sqldatoST = DateAdd("m",1,sqldatoST)
                            sqldatoHalvAar = DateAdd("m", 6, sqldatoHalvAar)
                            sqldatoHalvAarSlut = DateAdd("m", 6, sqldatoHalvAar)
                            sqldatoHalvAarSlut = DateAdd("d", -1, sqldatoHalvAarSlut)
                            sqldatoQAST = DateAdd("q",1,sqldatoQAST)
                            sqldatoQASL = DateAdd("q",1,sqldatoQAST)
                            sqldatoQASL = DateADd("d", -1, sqldatoQASL)
                            
                            if ntimPer <> 0 then
                                timer_fomr1_procent = (timer_for_fomr1 / ntimPer) * 100 'skal divideres med normen for den givende periode
                                timer_fomr2_procent = (timer_for_fomr2 / ntimPer) * 100
                                timer_fomr3_procent = (timer_for_fomr3 / ntimPer) * 100
                                timer_fomr4_procent = (timer_for_fomr4 / ntimPer) * 100
                                timer_fomr5_procent = (timer_for_fomr5 / ntimPer) * 100
                                timer_fomr6_procent = (timer_for_fomr6 / ntimPer) * 100
                                timer_fomr7_procent = (timer_for_fomr7 / ntimPer) * 100
                                timer_fomr8_procent = (timer_for_fomr8 / ntimPer) * 100
                                timer_fomr9_procent = (timer_for_fomr9 / ntimPer) * 100
                                timer_fomr10_procent = (timer_for_fomr10 / ntimPer) * 100
                                timer_fomr11_procent = (timer_for_fomr11 / ntimPer) * 100
                                timer_fomr12_procent = (timer_for_fomr12 / ntimPer) * 100
                                timer_fomr13_procent = (timer_for_fomr13 / ntimPer) * 100
                                timer_fomr14_procent = (timer_for_fomr14 / ntimPer) * 100
                                timer_fomr15_procent = (timer_for_fomr15 / ntimPer) * 100


                                reg_timer_fomr1_procent = (reg_timer_fomr1 / ntimPer) * 100 'skal divideres med normen for den givende periode
                                reg_timer_fomr2_procent = (reg_timer_fomr2 / ntimPer) * 100
                                reg_timer_fomr3_procent = (reg_timer_fomr3 / ntimPer) * 100
                                reg_timer_fomr4_procent = (reg_timer_fomr4 / ntimPer) * 100
                                reg_timer_fomr5_procent = (reg_timer_fomr5 / ntimPer) * 100
                                reg_timer_fomr6_procent = (reg_timer_fomr6 / ntimPer) * 100
                                reg_timer_fomr7_procent = (reg_timer_fomr7 / ntimPer) * 100
                                reg_timer_fomr8_procent = (reg_timer_fomr8 / ntimPer) * 100
                                reg_timer_fomr9_procent = (reg_timer_fomr9 / ntimPer) * 100
                                reg_timer_fomr10_procent = (reg_timer_fomr10 / ntimPer) * 100
                                reg_timer_fomr11_procent = (reg_timer_fomr11 / ntimPer) * 100
                                reg_timer_fomr12_procent = (reg_timer_fomr12 / ntimPer) * 100
                                reg_timer_fomr13_procent = (reg_timer_fomr13 / ntimPer) * 100
                                reg_timer_fomr14_procent = (reg_timer_fomr14 / ntimPer) * 100
                                reg_timer_fomr15_procent = (reg_timer_fomr15 / ntimPer) * 100
                            else
                                if timer_for_fomr1 <> 0 then timer_for_fomr1 = 100 else timer_for_fomr1 = 0 end if
                                if timer_for_fomr2 <> 0 then timer_for_fomr2 = 100 else timer_for_fomr2 = 0 end if
                                if timer_for_fomr3 <> 0 then timer_for_fomr3 = 100 else timer_for_fomr3 = 0 end if
                                if timer_for_fomr4 <> 0 then timer_for_fomr4 = 100 else timer_for_fomr4 = 0 end if
                                if timer_for_fomr5 <> 0 then timer_for_fomr5 = 100 else timer_for_fomr5 = 0 end if
                                if timer_for_fomr6 <> 0 then timer_for_fomr6 = 100 else timer_for_fomr6 = 0 end if
                                if timer_for_fomr7 <> 0 then timer_for_fomr7 = 100 else timer_for_fomr7 = 0 end if
                                if timer_for_fomr8 <> 0 then timer_for_fomr8 = 100 else timer_for_fomr8 = 0 end if
                                if timer_for_fomr9 <> 0 then timer_for_fomr9 = 100 else timer_for_fomr9 = 0 end if
                                if timer_for_fomr10 <> 0 then timer_for_fomr10 = 100 else timer_for_fomr10 = 0 end if
                                if timer_for_fomr11 <> 0 then timer_for_fomr11 = 100 else timer_for_fomr11 = 0 end if
                                if timer_for_fomr12 <> 0 then timer_for_fomr12 = 100 else timer_for_fomr12 = 0 end if
                                if timer_for_fomr13 <> 0 then timer_for_fomr13 = 100 else timer_for_fomr13 = 0 end if
                                if timer_for_fomr14 <> 0 then timer_for_fomr14 = 100 else timer_for_fomr14 = 0 end if
                                if timer_for_fomr15 <> 0 then timer_for_fomr15 = 100 else timer_for_fomr15 = 0 end if
                         

                                if reg_timer_fomr1 <> 0 then reg_timer_fomr1_procent = 100 else reg_timer_fomr1_procent = 0 end if
                                if reg_timer_fomr2 <> 0 then reg_timer_fomr2_procent = 100 else reg_timer_fomr2_procent = 0 end if
                                if reg_timer_fomr3 <> 0 then reg_timer_fomr3_procent = 100 else reg_timer_fomr3_procent = 0 end if
                                if reg_timer_fomr4 <> 0 then reg_timer_fomr4_procent = 100 else reg_timer_fomr4_procent = 0 end if
                                if reg_timer_fomr5 <> 0 then reg_timer_fomr5_procent = 100 else reg_timer_fomr5_procent = 0 end if
                                if reg_timer_fomr6 <> 0 then reg_timer_fomr6_procent = 100 else reg_timer_fomr6_procent = 0 end if
                                if reg_timer_fomr7 <> 0 then reg_timer_fomr7_procent = 100 else reg_timer_fomr7_procent = 0 end if
                                if reg_timer_fomr8 <> 0 then reg_timer_fomr8_procent = 100 else reg_timer_fomr8_procent = 0 end if
                                if reg_timer_fomr9 <> 0 then reg_timer_fomr9_procent = 100 else reg_timer_fomr9_procent = 0 end if
                                if reg_timer_fomr10 <> 0 then reg_timer_fomr10_procent = 100 else reg_timer_fomr10_procent = 0 end if
                                if reg_timer_fomr11 <> 0 then reg_timer_fomr11_procent = 100 else reg_timer_fomr11_procent = 0 end if
                                if reg_timer_fomr12 <> 0 then reg_timer_fomr12_procent = 100 else reg_timer_fomr12_procent = 0 end if
                                if reg_timer_fomr13 <> 0 then reg_timer_fomr13_procent = 100 else reg_timer_fomr13_procent = 0 end if
                                if reg_timer_fomr14 <> 0 then reg_timer_fomr14_procent = 100 else reg_timer_fomr14_procent = 0 end if
                                if reg_timer_fomr15 <> 0 then reg_timer_fomr15_procent = 100 else reg_timer_fomr15_procent = 0 end if
                            end if
                            'response.Write "<br> periode_"&i&"_fomr_1_forecast_"&oRec("mid")
                            
                            %>   
                                <!-- Dato underskrifter -->            
                                <input type="hidden" id="underskrift_periode_<%=i %>" value="<%=underskrift %>" />

                                <!-- Forecastet felter -->
                                <input type="hidden" id="periode_<%=i %>_fomr_1_forecast_<%=oRec("mid") %>" value="<%=timer_fomr1_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_2_forecast_<%=oRec("mid") %>" value="<%=timer_fomr2_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_3_forecast_<%=oRec("mid") %>" value="<%=timer_fomr3_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_4_forecast_<%=oRec("mid") %>" value="<%=timer_fomr4_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_5_forecast_<%=oRec("mid") %>" value="<%=timer_fomr5_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_6_forecast_<%=oRec("mid") %>" value="<%=timer_fomr6_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_7_forecast_<%=oRec("mid") %>" value="<%=timer_fomr7_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_8_forecast_<%=oRec("mid") %>" value="<%=timer_fomr8_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_9_forecast_<%=oRec("mid") %>" value="<%=timer_fomr9_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_10_forecast_<%=oRec("mid") %>" value="<%=timer_fomr10_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_11_forecast_<%=oRec("mid") %>" value="<%=timer_fomr11_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_12_forecast_<%=oRec("mid") %>" value="<%=timer_fomr12_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_13_forecast_<%=oRec("mid") %>" value="<%=timer_fomr13_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_14_forecast_<%=oRec("mid") %>" value="<%=timer_fomr14_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_15_forecast_<%=oRec("mid") %>" value="<%=timer_fomr15_procent %>" />

                                <!-- Registreret felter -->
                               <input type="hidden" id="periode_<%=i %>_fomr_1_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr1_procent %>" />
                               <input type="hidden" id="periode_<%=i %>_fomr_2_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr2_procent %>" />
                               <input type="hidden" id="periode_<%=i %>_fomr_3_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr3_procent %>" />
                               <input type="hidden" id="periode_<%=i %>_fomr_4_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr4_procent %>" />
                               <input type="hidden" id="periode_<%=i %>_fomr_5_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr5_procent %>" />
                               <input type="hidden" id="periode_<%=i %>_fomr_6_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr6_procent %>" />
                               <input type="hidden" id="periode_<%=i %>_fomr_7_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr7_procent %>" />
                               <input type="hidden" id="periode_<%=i %>_fomr_8_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr8_procent %>" />
                               <input type="hidden" id="periode_<%=i %>_fomr_9_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr9_procent %>" />
                               <input type="hidden" id="periode_<%=i %>_fomr_10_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr10_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_11_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr11_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_12_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr12_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_13_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr13_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_14_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr14_procent %>" />
                                <input type="hidden" id="periode_<%=i %>_fomr_15_registreret_<%=oRec("mid") %>" value="<%=reg_timer_fomr15_procent %>" />
                            <%

                            wend                            
                        %>

                        
                        <div class="col-lg-4" style="padding-top:50px;">
                           
                           <input type="hidden" id="medarbnavn_<%=oRec("mid") %>" value="<%=oRec("mnavn") %>" />

                           <input type="hidden" id="fomr_1_procent_<%=oRec("mid") %>" value="<%=timer_fomr1_procent %>" />
                           <input type="hidden" id="fomr_2_procent_<%=oRec("mid") %>" value="<%=timer_fomr2_procent %>" />
                           <input type="hidden" id="fomr_3_procent_<%=oRec("mid") %>" value="<%=timer_fomr3_procent %>" />
                           <input type="hidden" id="fomr_4_procent_<%=oRec("mid") %>" value="<%=timer_fomr4_procent %>" />
                           <input type="hidden" id="fomr_5_procent_<%=oRec("mid") %>" value="<%=timer_fomr5_procent %>" />
                           <input type="hidden" id="fomr_6_procent_<%=oRec("mid") %>" value="<%=timer_fomr6_procent %>" />
                           <input type="hidden" id="fomr_7_procent_<%=oRec("mid") %>" value="<%=timer_fomr7_procent %>" />
                           <input type="hidden" id="fomr_8_procent_<%=oRec("mid") %>" value="<%=timer_fomr8_procent %>" />
                           <input type="hidden" id="fomr_9_procent_<%=oRec("mid") %>" value="<%=timer_fomr9_procent %>" />
                            <input type="hidden" id="fomr_10_procent_<%=oRec("mid") %>" value="<%=timer_fomr10_procent %>" />
                            <input type="hidden" id="fomr_11_procent_<%=oRec("mid") %>" value="<%=timer_fomr11_procent %>" />
                            <input type="hidden" id="fomr_12_procent_<%=oRec("mid") %>" value="<%=timer_fomr12_procent %>" />
                            <input type="hidden" id="fomr_13_procent_<%=oRec("mid") %>" value="<%=timer_fomr13_procent %>" />
                            <input type="hidden" id="fomr_14_procent_<%=oRec("mid") %>" value="<%=timer_fomr14_procent %>" />
                            <input type="hidden" id="fomr_15_procent_<%=oRec("mid") %>" value="<%=timer_fomr15_procent %>" />

                           <input type="hidden" id="reg1_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr1_procent %>" />
                           <input type="hidden" id="reg2_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr2_procent %>" />
                           <input type="hidden" id="reg3_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr3_procent %>" />
                           <input type="hidden" id="reg4_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr4_procent %>" />
                           <input type="hidden" id="reg5_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr5_procent %>" />
                           <input type="hidden" id="reg6_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr6_procent %>" />
                           <input type="hidden" id="reg7_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr7_procent %>" />
                           <input type="hidden" id="reg8_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr8_procent %>" />
                           <input type="hidden" id="reg9_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr9_procent %>" />
                           <input type="hidden" id="reg10_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr10_procent %>" />
                           <input type="hidden" id="reg11_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr11_procent %>" />
                            <input type="hidden" id="reg12_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr12_procent %>" />
                            <input type="hidden" id="reg13_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr13_procent %>" />
                            <input type="hidden" id="reg14_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr14_procent %>" />
                            <input type="hidden" id="reg15_timer_procent_<%=oRec("mid") %>" value="<%=reg_timer_fomr15_procent %>" />

                            <h4><%=oRec("mnavn") %></h4><br />
                           <div class="chartContainer" id="<%=oRec("mid") %>" style="width:100%; height:300px"></div>
                        </div>

                <%
                    oRec.movenext
                    wend
                    oRec.close
                %>
                </div>

                <%
                    'response.Write "<br> totalReg_per1 " & totalReg_per1   
                    'response.Write "<br> totalReg_per2 " & totalReg_per2 
                    'response.Write "<br> totalReg_per3 " & totalReg_per3 
                    'response.Write "<br> totalReg_per4 " & totalReg_per4
                    
                    'response.Write "<br> totalFc_per1 " & totalFc_per1   
                    'response.Write "<br> totalFc_per2 " & totalFc_per2 
                    'response.Write "<br> totalFc_per3 " & totalFc_per3 
                    'response.Write "<br> totalFc_per4 " & totalFc_per4

                    for p = 0 TO 3
                        for f3 = 0 TO ubound(fomr)

                                'response.Write "<br> fomr " & f3 &" HH "& totalReg(f3, p) & " " & totalnorm_per

                                select case cint(p)
                                    case 0
                                        if totalReg(f3, p) <> 0 AND totalReg(f3, p) <> "" AND totalnorm_per1 <> 0 AND totalnorm_per1 <> "" then
                                        totalRegProcent = totalReg(f3, p) / totalnorm_per1 * 100
                                        else
                                        totalRegProcent = 0
                                        end if

                                        if totalFc(f3, p) <> 0 AND totalFc(f3, p) <> "" AND totalnorm_per1 <> 0 AND totalnorm_per1 <> "" then
                                        totalFcProcent = totalFc(f3, p) / totalnorm_per1 * 100
                                        else
                                        totalFcProcent = 0
                                        end if
                                    case 1    
                                        if totalReg(f3, p) <> 0 AND totalReg(f3, p) <> "" AND totalnorm_per2 <> 0 AND totalnorm_per2 <> "" then 
                                        totalRegProcent = totalReg(f3, p) / totalnorm_per2 * 100
                                        else
                                        totalRegProcent = 0
                                        end if

                                        if totalFc(f3, p) <> 0 AND totalFc(f3, p) <> "" AND totalnorm_per2 <> 0 AND totalnorm_per2 <> "" then
                                        totalFcProcent = totalFc(f3, p) / totalnorm_per2 * 100
                                        else
                                        totalFcProcent = 0
                                        end if
                                    case 2              
                                        if totalReg(f3, p) <> 0 AND totalReg(f3, p) <> "" AND totalnorm_per3 <> 0 AND totalnorm_per3 <> "" then
                                        totalRegProcent = totalReg(f3, p) / totalnorm_per3 * 100
                                        else
                                        totalRegProcent = 0
                                        end if

                                        if totalFc(f3, p) <> 0 AND totalFc(f3, p) <> "" AND totalnorm_per3 <> 0 AND totalnorm_per3 <> "" then
                                        totalFcProcent = totalFc(f3, p) / totalnorm_per3 * 100
                                        else
                                        totalFcProcent = 0
                                        end if
                                    case 3                          

                                        if totalReg(f3, p) <> 0 AND totalReg(f3, p) <> "" AND totalnorm_per4 <> 0 AND totalnorm_per4 <> "" then
                                        totalRegProcent = totalReg(f3, p) / totalnorm_per4 * 100 
                                        else
                                        totalRegProcent = 0
                                        end if

                                        if totalFc(f3, p) <> 0 AND totalFc(f3, p) <> "" AND totalnorm_per4 <> 0 AND totalnorm_per4 <> "" then
                                        totalFcProcent = totalFc(f3, p) / totalnorm_per4 * 100
                                        else
                                        totalFcProcent = 0
                                        end if
                                end select
                            'response.Write "<br> TIMER " & totalFc(fomr(f3), p)
                            'response.Write "totreg_"&f3&"_"&p & " tofc_"&f3&"_"&p
                            response.Write "<input type='hidden' value='"& totalRegProcent &"' id='totreg_"&f3&"_"&p&"' />"
                            response.Write "<input type='hidden' value='"& totalFcProcent &"' id='tofc_"&f3&"_"&p&"' />"
                           ' response.Write "<input type='text' value='"& totalFc(fomr(f3), p) &"' class='totalFc' data-fomr='"&fomr(f3)&"' data-periode='"&p&"' />"
                        next
                    next

                    'response.Write "<br><br> Norm per 1 " & totalnorm_per1
                    'response.Write "<br> Norm per 2 " & totalnorm_per2
                    'response.Write "<br> Norm per 3 " & totalnorm_per3
                    'response.Write "<br> Norm per 4 " & totalnorm_per4



                    
                %>                
            </div>
        </div>
    </div>

    


    <script type="text/javascript">

        farvebetydningerVist = 0
        $("#ColorCom_btn").click(function () {

            if (farvebetydningerVist == 0)
            {
                $("#farvebetydninger").show();
                farvebetydningerVist = 1
            } else
            {
                $("#farvebetydninger").hide();
                farvebetydningerVist = 0
            }
            
            

        });


    </script>





</div>
</div>
    
<script src="canvasjs/canvasjs.min.js"></script>

<!--#include file="../inc/regular/footer_inc.asp"-->