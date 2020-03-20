<%
thisfile = "timebudget_scroll"
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<div class="wrapper">
    <div class="content">

<%
    if len(session("user")) = 0 then
	    errortype = 5
	    call showError(errortype)
        response.End
    end if
    
    media = request("media")

    Session.LCID = 1030

    if media <> "print" then
        call menu_2014()
    end if

    if len(trim(request("FM_start_mrd"))) <> 0 then
	strMrd = request("FM_start_mrd")
	else
        select case lto
        case "adra"
        strMrd = month(now)
        case else 
	    strMrd = month(now)
        end select
	end if

    if len(trim(request("yuse"))) <> 0 then
	ysel = request("yuse")
	else
	ysel = year(now)
	end if


    '**** PERIOD **********
      'per_interval = request("per_interval") 
      if media <> "print" then
      per_interval = 12
      else
      per_interval = request("per_interval") 
      end if

      select case per_interval
      case "1"
      per_i_enCHK = "SELECTED"
      per_interval = 1
      daysInInterval = 31
      faktor = 4
      weekspan = 7

      case "3", ""

      per_i_treCHK = "SELECTED"
      per_interval = 3
      daysInInterval = 170
      faktor = 4
      weekspan = 7

      case "6"
            per_interval = 6
            per_i_6CHK = "SELECTED"
      case else

        select case lto
        case "adra"

          per_i_treCHK = "SELECTED"
          per_interval = 3
          daysInInterval = 170
          faktor = 4
          weekspan = 7

        case else

          per_i_tolvCHK = "SELECTED" 
          per_interval = 12 
          daysInInterval = 353
          faktor = 1
          weekspan = 1
        
        end select

    end select

    if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
	progrp = 0
	end if

    if len(request("FM_medarb")) <> 0 OR func = "export" then
	
	    if left(request("FM_medarb"), 1) = "0" then
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

    ynow = year(now)
    mnow = month(now)

    if mnow < 5 then
        startDatoFEOSQL = (ynow - 1) &"/5/1"
        slutDatoFEOSQL = ynow &"/4/30"
    else 
        startDatoFEOSQL = ynow &"/5/1"
        slutDatoFEOSQL = (ynow+1) &"/4/30"
    end if

    if lto = "lm" then
    startferiefriaar = startDatoFEOSQL
    slutferiefriaar = slutDatoFEOSQL
    else
    startferiefriaar = year(now) &"-01-01" 
    slutferiefriaar = year(now) &"-12-31"
    end if


%>

        <script src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
        <script src="https://cdn.datatables.net/fixedcolumns/3.2.4/js/dataTables.fixedColumns.min.js"></script> 


        <link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
        <link rel="stylesheet" href="https://cdn.datatables.net/fixedcolumns/3.2.4/css/fixedColumns.dataTables.min.css" />

        <script src="../timereg/inc/header_lysblaa_stat_20170808.js" type="text/javascript"></script>
       <script src="js/feriekalender2.js"></script>

        <%
            if media = "print" then 
                conwidth = "100%"
                conZoom = "zoom:45%;"
            else
                conwidth = "85%"
                conZoom = ""
            end if
        %>

        <div class="container" style="width:<%=conwidth%>; <%=conZoom%>">
            <div class="portlet">

                <h3 class="portlet-title"><u><%=feriekalender_txt_001 &" & "& feriekalender_txt_002 %></u></h3>

                <div class="portlet-body">

                    <%if media <> "print" then %>
                    <form action="feriekalender.asp" method="post">
                        <div class="well">
                            <div class="row">
                                <div class="col-lg-6"><b><%=feriekalender_txt_003 %>:</b></div>
                                <div class="col-lg-2"><b><%=feriekalender_txt_004 %>:</b></div>
                                <div class="col-lg-1"><b><%=feriekalender_txt_005 %>:</b></div>
                                
                            </div>


                           <!-- <div class="row">
                                <div class="col-lg-12">                            
                                    <div style="height:25px; width:25px; background-color:red;" class="btn btn-default demo-element ui-popover" data-toggle="tooltip" data-placement="top" data-trigger="hover" data-content="Tekst" title="Title"></div>
                                </div>
                            </div> -->

                            <div class="row">
                                <%call progrpmedarb_2018 %>

                                <div class="col-lg-1">
                                  <select name="FM_start_mrd" class="form-control input-small">
                                        <%
                                        for m = 1 TO 12

                                            if cint(m) = cint(strMrd) then
                                                mSEL = "SELECTED"
                                            else
                                                mSEL = ""
                                            end if

                                            select case cint(m)
                                                case 1
                                                    strmonth = feriekalender_txt_038
                                                case 2
                                                    strmonth = feriekalender_txt_039
                                                case 3
                                                    strmonth = feriekalender_txt_040
                                                case 4
                                                    strmonth = feriekalender_txt_041
                                                case 5
                                                    strmonth = feriekalender_txt_042
                                                case 6
                                                    strmonth = feriekalender_txt_043
                                                case 7
                                                    strmonth = feriekalender_txt_044
                                                case 8
                                                    strmonth = feriekalender_txt_045
                                                case 9
                                                    strmonth = feriekalender_txt_046
                                                case 10
                                                    strmonth = feriekalender_txt_047
                                                case 11
                                                    strmonth = feriekalender_txt_048
                                                case 12
                                                    strmonth = feriekalender_txt_049
                                            end select

                                            response.Write "<option value='"&m&"' "&mSEL&" >"&strmonth&"</option>"
                                        next
                                        %>

                                    </select>
                                </div>

                                <div class="col-lg-1">
                                    <select id="Select2" name="yuse" class="form-control input-small">
                                        <%
	                                    'ysel = now
	                                    for y = -5 to 10 
                                            yShow = dateAdd("yyyy", y, now) 
        
                                            if cint(year(yShow)) = cint(ysel) then
                                            ySele = "SELECTED"
                                            else
                                            ySele = ""
                                            end if 
                                            %>
	                                        <option value="<%=datePart("yyyy", yShow, 2,2)%>" <%=ySele %>><%=datePart("yyyy", yShow, 2,2)%></option> 
	                                    <%next %>
            
                                    </select>
                                </div>

                                <div class="col-lg-2">
                                   <!-- <select id="per_interval" name="per_interval" class="form-control input-small">
                                        <option value="1" <%=per_i_enCHK %>>1 m�neder frem</option>
                                        <option value="3" <%=per_i_treCHK %>>3 m�neder frem</option>
                                        <option value="6" <%=per_i_6CHK %>>6 m�neder frem</option>
                                        <option value="12" <%=per_i_tolvCHK %>>12 m�neder frem</option>
                                    </select> -->

                                    <%=feriekalender_txt_006 %>

                                </div>

                                <div class="col-lg-1">
                                    <button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=feriekalender_txt_007 %></b></button>
                                </div>

                            </div>
                        </div>
                    </form>
                    <%end if %>


                    <%
                    stDato = ysel &"-"& strMrd & "-1"
                    stDato = day(stDato) &"-"& month(stDato) &"-"& year(stDato)
                    slDato = dateAdd("m", per_interval, stDato)
                    'response.Write stDato & " til " & slDato

                    antalDage = dateDiff("d", stDato, slDato, 2,2)
                    'response.Write "<br>" & antalDage
                    %>

                  <!--  <style>
                        tr:nth-child(even) {background-color: white;}
                        tr:nth-child(odd) {background-color: #f7f7f7;}

                        td
                        {
                            font-size:12px;
                            border:none !important;
                        }

                        th {
                            font-size:12px !important;
                        }

                        .tablegraf > thead > tr > th,
                        .tablegraf > tbody > tr > th,
                        .tablegraf > tfoot > tr > th,
                        .tablegraf > thead > tr > td,
                        .tablegraf > tbody > tr > td,
                        .tablegraf > tfoot > tr > td {
                          padding: 1px !important;
                          line-height: 1.42857143;
                          vertical-align: top;
                          border-top: 1px solid #ddd;
                        }

                    </style> -->

                    <style>
                        th, td {
                            padding:1px;
                            font-size: 12px !important;
                        }
                    </style>
                    
                    <%
                    if media <> "print" then
                        tableid = "main_datatable_forecast"
                        caltdWidth = "10px"
                        boxSize = "11px;"
                    else
                        tableid = ""
                        caltdWidth = "25px"

                        if per_interval = 1 then
                            boxSize = "25px;"
                            caltdWidth = "50px"
                            %>
                            <style>
                                th, td {
                                    font-size: 25px !important;
                                }
                            </style>
                            <%

                        else
                            boxSize = "11px;"
                        end if

                    end if
                    %>

                    <%
                    call browsertype()

                    userAgent = request.servervariables("HTTP_USER_AGENT")

                    if instr(userAgent , "Edge") <> 0 then
                        browstype_client = "eg"
                    end if

                    if browstype_client <> "ch" then
                        tableid = "x"
                        caltdWidth = "25px;"
                    end if
                    %>
                     
                    <%if tableid = "x" AND media <> "print" then %>
                    <div style="width:100%; overflow-x:scroll; max-height:800px; overflow-y:scroll;">
                    <%end if %>

                    <table id="<%=tableid %>" class="table table-bordered">
                        <thead>
                            <tr>
                                <th style="min-width:125px;"><%=feriekalender_txt_008 %></th>
                                
                                <%
                                    'duse = "2019-12-31"
                                    'response.Write "Week " & DatePart("ww", duse, 2,1) %>

                                <%
                                loopDate = stDato
                                lastWeek = 0
                                loopWeek = DatePart("ww", loopDate, 2,2)
                                for d = 0 TO (antalDage -1)
                                    'response.Write d
                                    'response.Write "<br> UGE" & loopWeek & " DD " & loopDate
                                    if d <> 0 then
                                        loopDate = dateAdd("d", 1, loopDate)
                                        loopWeek = DatePart("ww", loopDate, 2,2)

                                        If loopWeek = 53 And DatePart("ww", DateAdd("d", 7, loopDate), vbMonday, vbFirstFourDays) = 2 Then
                                            loopWeek = 1
                                        End If

                                    end if
                                    
                                    if loopWeek <> lastWeek then
                                        weekStartdato = loopDate
                                        weekStartDay = weekday(weekStartdato, 0)
                                        'response.Write "<br> UGE" & loopWeek & " DD " & loopDate & " WS " & weekStartDay
                                        loopWeekEndDay = loopWeek + 1
                                        
                                        'daysInWeek = (weekStartDay - 1)
                                        select case weekStartDay
                                            case 1
                                                daysInWeek = 7
                                            case 2
                                                daysInWeek = 6
                                            case 3
                                                daysInWeek = 5
                                            case 4
                                                daysInWeek = 4
                                            case 5
                                                daysInWeek = 3
                                            case 6
                                                daysInWeek = 2
                                            case 7
                                                daysInWeek = 1
                                        end select
                                       ' daysInWeek = (7 - daysInWeek)
                                    end if

                                    'response.Write "<br> UGE" & loopWeek & " DD " & loopDate & " WS "

                                    if d = 0 then
                                        response.Write "<th style='text-align:center;' colspan='"&cint(daysInWeek)&"'>"&loopWeek&"</th>"
                                    else
                                        if loopWeek <> lastWeek then
                                            response.Write "<th style='text-align:center;' colspan='7'>"&loopWeek&"</th>"
                                        end if
                                    end if

                                lastWeek = loopWeek
                                next
                                %>
                            </tr>


                            <tr>
                                <th><%=feriekalender_txt_009 %></th>

                                <%
                                loopDate = stDato
                                loopMonth = month(loopDate)
                                lastMonth = 0
                                for d = 0 TO (antalDage -1)

                                    if d <> 0 then
                                        loopDate = dateAdd("d", 1, loopDate)
                                        loopMonth = month(loopDate)
                                    end if

                                    if loopMonth <> lastMonth then

                                        firstDayInMonth = loopDate
                                        lastDayInMonth = dateAdd("m", 1, loopDate)
                                        lastDayInMonth = dateAdd("d", -1, lastDayInMonth)
                                        daysInMonth = dateDiff("d", firstDayInMonth, lastDayInMonth,2,2)
                                        daysInMonth = daysInMonth + 1

                                        select case cint(loopMonth)
                                            case 1
                                                strmonth = feriekalender_txt_038
                                            case 2
                                                strmonth = feriekalender_txt_039
                                            case 3
                                                strmonth = feriekalender_txt_040
                                            case 4
                                                strmonth = feriekalender_txt_041
                                            case 5
                                                strmonth = feriekalender_txt_042
                                            case 6
                                                strmonth = feriekalender_txt_043
                                            case 7
                                                strmonth = feriekalender_txt_044
                                            case 8
                                                strmonth = feriekalender_txt_045
                                            case 9
                                                strmonth = feriekalender_txt_046
                                            case 10
                                                strmonth = feriekalender_txt_047
                                            case 11
                                                strmonth = feriekalender_txt_048
                                            case 12
                                                strmonth = feriekalender_txt_049
                                        end select

                                        response.Write "<th style='text-align:center;' colspan='"&daysInMonth&"'>"&strmonth&"</th>"
                                    end if

                                    lastMonth = loopMonth
                                next
                                %>

                            </tr>



                            <%'if per_interval = 1 OR per_interval = 3 then %>
                            <tr>
                                <th><%=feriekalender_txt_010 %></th>
                                <%
                                loopDate = stDato
                                for d = 0 TO (antalDage -1)

                                    if d <> 0 then
                                        loopDate = dateAdd("d", 1, loopDate)
                                    end if

                                    dayNum = DatePart("w",loopDate ,2 ,2)
                                    if dayNum = 6 OR dayNum = 7 then
                                        thbgColor = "#d4d4d4"
                                    else
                                        thbgColor = ""
                                    end if

                                    response.Write "<th style='text-align:center; max-width:20px !important; padding:1px !important; background-color:"& thbgColor &";'>"&day(loopDate)&"</th>"

                                next
                                %>
                            </tr>
                            <%'end if %>
                        </thead>


                        <%
                        normTimerGns5 = 1
                        dim medarbnavn, medarbNr, medarbInit, ferieafholdtTot, ferieplanlagt, ferieudbetalt, ferieulon, feriefribrugt, feriefriudbetalt, rejsedageTot, ferieaarOptjent, ferieaarAfholdt, ferieaarUdbetalt, ferieaarAfholdtUdlon, ferieaarOverfort, ferieaarOptUdenLon, sygPertot, barnsygPertot, afspadPertot, feriefriaarsaldo, feriefriaarbrugt, feriefriaaropt
                        redim medarbnavn(700), medarbNr(700), medarbInit(700), ferieafholdtTot(700), ferieplanlagt(700), ferieudbetalt(700), ferieulon(700), feriefribrugt(700), feriefriudbetalt(700), rejsedageTot(700), ferieaarsaldo(700), ferieaarOptjent(700), ferieaarAfholdt(700), ferieaarUdbetalt(700), ferieaarAfholdtUdlon(700), ferieaarOverfort(700), ferieaarOptUdenLon(700), sygPertot(700), barnsygPertot(700), afspadPertot(700), feriefriaarsaldo(700), feriefriaarbrugt(700), feriefriaaropt(700)
                             
                      
                        startDatoPerSQL = year(stDato) &"-"& month(stDato) &"-"& day(stDato)
                        slutDatoPerSQL = year(slDato) &"-"& month(slDato) &"-"& day(slDato)

                        if vis = 1 then '** Ferie rapport 

                        sqlTypKri = "tfaktim = 14 OR tfaktim = 16 OR tfaktim = 13 OR tfaktim = 17 OR tfaktim = 19"

                        else
	
	                    if level <= 2 OR level = 6 then
	                    sqlTypKri = "tfaktim = 11 OR tfaktim = 14 OR tfaktim = 16 OR tfaktim = 20 OR tfaktim = 21 OR tfaktim = 31 OR tfaktim = 13 OR tfaktim = 17 OR tfaktim = 18 OR tfaktim = 19 OR tfaktim = 125"
	                    else
	                    sqlTypKri = "tfaktim = 11 OR tfaktim = 14 OR tfaktim = 16 OR tfaktim = 31 OR tfaktim = 13 OR tfaktim = 17 OR tfaktim = 18 OR tfaktim = 19 OR tfaktim = 125"
	                    end if
	
                        if lto = "cst" then
                        sqlTypKri = sqlTypKri & " OR tfaktim = 7"
                        end if
    
                        end if


                        sqlGrpBy = "tdato, tfaktim, tmnr"

                        for m = 0 TO UBOUND(intMids)
                            
                            dim antalTimer, exportTxt
                            redim antalTimer(12,2,365,150), exportTxt(12,2,365,150)

                            'Ferie Saldo

                            'response.Write "Start " & startDatoFEOSQL
                            'response.Write "Slut " & slutDatoFEOSQL

                            ferieOptjent = 0
                            strSQL =  "SELECT sum(timer) AS timer, tdato, month(tdato) AS month, tfaktim, tastedato FROM timer WHERE tmnr = " & intMids(m)  & ""_
	                        &" AND (tfaktim = 15 OR tfaktim = 111 OR tfaktim = 112) AND tdato BETWEEN '"& startDatoFEOSQL &"' AND '"& slutDatoFEOSQL &"' GROUP BY tfaktim ORDER BY tdato"
                            
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF

                                sqlNow = year(now) &"-"& month(now) &"-"& day(now)
                                call normtimerPer(intMids(m) , oRec("tdato"), 6, 0)

                                if ntimPer <> 0 then
                                'ntimPerUse = ntimPer/antalDageMtimer
                                normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
                                else
                                normTimerGns5 = 1
                                end if 
                                
                                if oRec("tfaktim") = 15 then
                                    ferieaarOptjent(m) = oRec("timer") / normTimerGns5
                                end if
                                
                                if oRec("tfaktim") = 111 then
                                    ferieaarOverfort(m) = oRec("timer") / normTimerGns5
                                end if

                                if oRec("tfaktim") = 112 then
                                    ferieaarOptUdenLon(m) = oRec("timer") / normTimerGns5
                                end if

                            
                            oRec.movenext
                            wend
                            oRec.close

                            ferieafholdt = 0
                            strSQL = "SELECT timer, tdato, month(tdato) AS month, tfaktim, tastedato FROM timer WHERE tmnr = " & intMids(m)  & ""_
	                        &" AND (tfaktim = 14 OR tfaktim = 19 OR tfaktim = 16) AND tdato BETWEEN '"& startDatoFEOSQL &"' AND '"& slutDatoFEOSQL &"' ORDER BY tdato"
                            'response.Write "sql afholdt " & strSQL
                            oRec.open strSQL, oConn, 3
                            
                            while not oRec.EOF

                                call normtimerPer(intMids(m) , oRec("tdato"), 0, 0)
	                            if ntimPer <> 0 then
                                ntimPerUse = ntimPer
                                else
                                ntimPerUse = 1
                                end if 

                                if oRec("tfaktim") = 14 then
                                    ferieaarAfholdt(m) = ferieaarAfholdt(m) + (oRec("timer") / ntimPerUse)
                                end if

                                if oRec("tfaktim") = 16 then
                                    ferieaarUdbetalt(m) = ferieaarUdbetalt(m) + (oRec("timer") / ntimPerUse)
                                end if

                                if oRec("tfaktim") = 19 then
                                    ferieaarAfholdtUdlon(m) = ferieaarAfholdtUdlon(m) + (oRec("timer") / ntimPerUse)
                                end if
                                
                            oRec.movenext
                            wend
                            oRec.close
                            
                            feriebal = 0


                            call ferieBal_fn(ferieaarOptjent(m), ferieaarOverfort(m), ferieaarOptUdenLon(m), ferieaarAfholdt(m), ferieaarAfholdtUdlon(m), ferieaarUdbetalt(m))
                             
                            ferieaarsaldo(m) = feriebal 'ferieaarOptjent(m) - (ferieaarAfholdt(m) + ferieaarUdbetalt(m))



                            'Feriefri hele �r
                            strSQLfeau = "SELECT timer, tdato, month(tdato) AS month, tfaktim, tastedato FROM timer WHERE tmnr = " & intMids(m)  & ""_
	                        &" AND (tfaktim = 13 OR tfaktim = 17 OR tfaktim = 12) AND tdato BETWEEN '"& startferiefriaar &"' AND '"& slutferiefriaar &"' ORDER BY tdato"
                            'response.Write "FERIEFRI " & strSQLfeau
                            oRec.open strSQLfeau, oConn, 3
                            while not oRec.EOF
                                if oRec("tfaktim") = 13 OR oRec("tfaktim") = 17 then

                                    call normtimerPer(intMids(m) , oRec("tdato"), 0, 0)
	                                if ntimPer <> 0 then
                                    ntimPerUse = ntimPer
                                    else
                                    ntimPerUse = 1
                                    end if 

                                    feriefriaarbrugt(m) = feriefriaarbrugt(m) + (oRec("timer") / ntimPerUse)
                                else

                                    sqlNow = year(now) &"-"& month(now) &"-"& day(now)
                                    call normtimerPer(intMids(m) , oRec("tdato"), 6, 0)

                                    if ntimPer <> 0 then
                                    'ntimPerUse = ntimPer/antalDageMtimer
                                    normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
                                    else
                                    normTimerGns5 = 1
                                    end if 

                                    feriefriaaropt(m) = feriefriaaropt(m) + (oRec("timer") / normTimerGns5)
                                end if
                            oRec.movenext
                            wend
                            oRec.close

                            'response.Write "<br> FEFRI OPTJ " & feriefriaaropt(m)
                            'response.Write "<br> FEFRI BRUGHT " & feriefriaarbrugt(m)

                            feriefriaarsaldo(m) = feriefriaaropt(m) - feriefriaarbrugt(m)

                            'response.Write "<br> ferieOptjent " & ferieOptjent & " ferieafholdt " & ferieafholdt

                            'Henter medarb info
                            if intMids(m) <> 0 then
                                strSQL = "SELECT mid, mnavn, mnr, init FROM medarbejdere WHERE mid = "& intMids(m) 
                                oRec.open strSQL, oConn, 3
                                if not oRec.EOF then
                                    medarbnavn(m) = oRec("mnavn") 
                                    medarbNr(m) = oRec("mnr")
                                    medarbInit(m) = oRec("init")
                                end if
                                oRec.close


                                '* Main SQL '*
                                strSQLt = "SELECT sum(timer) AS timer, tdato, month(tdato) AS month, tfaktim, tmnr, timerkom, tastedato FROM timer WHERE tmnr = " & intMids(m)  & ""_
	                                      &" AND ("& sqlTypKri &") AND tdato BETWEEN '"& startDatoPerSQL &"' AND '"& slutDatoPerSQL &"' GROUP BY "& sqlGrpBy &" ORDER BY tdato"
                                'response.Write strSQLt
                                useYear = 1
                                antalReg = 0
                                divTitle = ""
                                divText = ""
                                bgColor = ""
                                giff = ""
                                hoursval = 0
                                dagval = 0
                                title = ""

                                oRec.open strSQLt, oConn, 3
                                while not oRec.EOF


                                    call normtimerPer(intMids(m), oRec("tdato"), 0, 0)
                                    nomrTimerPrDag = ntimper

                                    'Response.Write oRec("tdato") & "ntim: "& ntimper &"  erH:"& erHellig & "<br>"

                                    if len(trim(nomrTimerPrDag)) <> 0 AND nomrTimerPrDag <> 0 then
                                    nomrTimerPrDag = nomrTimerPrDag
                                    nomrTimerPrDagTxt = nomrTimerPrDag
                                    else
                                    nomrTimerPrDag = 1
                                    nomrTimerPrDagTxt = 0
                                    end if


                                    thisyear = year(oRec("tdato"))
                                    if antalReg <> 0 then
                                        if thisyear <> lastyear then
                                            useYear = 2
                                        end if
                                    end if

                                    select case cint(oRec("tfaktim"))
                                        case 11, 18
                                            bgColor = "silver" 'Ferie planlagt, Feriefridag pl
                                            giff = "dot_graae.gif" 

                                            call akttyper(oRec("tfaktim"), 1)
                                            title = akttypenavn
                                            hoursval = oRec("timer")
                                            dagval = hoursval / nomrTimerPrDag
                                            'dageVal = hoursval / 7

                                            ferieplanlagt(m) = ferieplanlagt(m) + dagval

                                        case 14 'Ferie afholdt 
                                            bgColor = "green"
                                            giff = "dot_gron.gif" 

                                            call akttyper(oRec("tfaktim"), 1)
                                            title = akttypenavn
                                            hoursval = oRec("timer")
                                            dagval = hoursval / nomrTimerPrDag

                                            if vis = 1 then
                                                ferieafholdtTot(m) = ferieafholdtTot(m) + (oRec("timer") / nomrTimerPrDag)
                                            else
                                                ferieafholdtTot(m) = ferieafholdtTot(m) + (oRec("timer") / nomrTimerPrDag)
                                            end if

                                        case 16 'Ferie Udbetalt
                                            bgColor = "pink"
                                            giff = "dot_darkpink.gif"

                                            call akttyper(oRec("tfaktim"), 1)
	                                        title = akttypenavn '"Ferie udbetalt"
                                
                                            hoursval = oRec("timer")
                                            dagval = hoursval / nomrTimerPrDag

                                            if vis = 1 then
                                             'dageVal_16(m, mtThis) = dageVal_16(m, mtThis) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
                                             'intMidsFeUb(m) = intMidsFeUb(m) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
                                              ferieudbetalt(m) = ferieudbetalt(m) + (oRec("timer") / nomrTimerPrDag)
                                            else
	                                         'dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
                                             'intMidsFeUb(m) = intMidsFeUb(m) + dageVal/1
                                              ferieudbetalt(m) = ferieudbetalt(m) + (oRec("timer") / nomrTimerPrDag)
                                            end if

                                        case 19 'Ferie u l�n
                                            bgColor = "yellowgreen"
                                            giff = "dot_yellowgron.gif"

                                            title = feriekalender_txt_011
                                            hoursval = oRec("timer")
                                            dagval = hoursval / nomrTimerPrDag

                                            if vis = 1 then
                                            'dageVal_19(m, mtThis) = dageVal_19(m, mtThis) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
                                            'intMidsFeAfUL(m) = intMidsFeAfUL(m) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
                                            ferieulon(m) = ferieulon(m) + (oRec("timer") / nomrTimerPrDag)
                                            else
	                                        'dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
                                            'intMidsFeAfUL(m) = intMidsFeAfUL(m) + dageVal/1
                                            ferieulon(m) = ferieulon(m) + (oRec("timer") / nomrTimerPrDag)
                                            end if

                                        case 20 'Syg
                                            bgColor = "red"
                                            giff = "dot_rod.gif"

                                            title = feriekalender_txt_012
                                            hoursval = oRec("timer")
                                            dagval = hoursval / nomrTimerPrDag

                                            sygPertot(m) = sygPertot(m) + dagval

                                        case 21 'Barnsyg
                                            bgColor = "orange"
                                            giff = "dot_orange.gif"
    
                                            title = feriekalender_txt_013
                                            hoursval = oRec("timer")
                                            dagval = hoursval / nomrTimerPrDag
                                            
                                            barnsygPertot(m) = barnsygPertot(m) + dagval
                                
                                        case 31 'Afsp. 
                                            bgColor = "#8cAAe6"
                                            giff = "dot_blaa.gif"

                                            title = feriekalender_txt_014
                                            hoursval = oRec("timer")
                                            dagval = hoursval / nomrTimerPrDag
                                
                                            afspadPertot(m) = afspadPertot(m) + dagval

                                        case 13 'Feriefri brugt
                                            bgColor = "yellow"
                                            giff = "dot_gul.gif"

                                            hoursval = oRec("timer")       
                                            dagval = hoursval / nomrTimerPrDag
                                            call akttyper(oRec("tfaktim"), 1)
                                            title = akttypenavn
                                            if vis = 1 then
                                            'dageVal_13(m, mtThis) = dageVal_13(m, mtThis) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
                                            'intMidsFeFri(m) = intMidsFeFri(m) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
                                            feriefribrugt(m) = feriefribrugt(m) + (oRec("timer") / nomrTimerPrDag)
                                            else
	                                        'dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
                                            'intMidsFeFri(m) = intMidsFeFri(m) + dageVal/1
                                            feriefribrugt(m) = feriefribrugt(m) + (oRec("timer") / nomrTimerPrDag)  
                                            end if

                                        case 17 'Feriefri Udbetalt
                                            bgColor = "#D592E1"
                                            giff = "dot_lightpink.gif"

                                            hoursval = oRec("timer")  
                                            dagval = hoursval / nomrTimerPrDag
                                            call akttyper(oRec("tfaktim"), 1)
                                            title = akttypenavn
                                            if vis = 1 then
                                            'dageVal_13(m, mtThis) = dageVal_13(m, mtThis) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
                                            'intMidsFeFri(m) = intMidsFeFri(m) + (formatnumber(oRec("timer")/(nomrTimerPrDag), 2)/1)
                                            feriefriudbetalt(m) = feriefriudbetalt(m) + (oRec("timer") / nomrTimerPrDag)
                                            else
	                                        'dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
                                            'intMidsFeFri(m) = intMidsFeFri(m) + dageVal/1
                                            feriefriudbetalt(m) = feriefriudbetalt(m) + (oRec("timer") / nomrTimerPrDag)
                                            end if

                                        case 125 'Rejsedage
                                            bgColor = "#E7A1EF"
                                            giff = "dot_E7A1EF.gif"

                                            call akttyper(oRec("tfaktim"), 1)
	                                        title = akttypenavn '"Rejsedage"
                                            hoursval = oRec("timer")

                                            call normtimerPer(intMids(m) , oRec("tdato"), 6, 0)

                                            if ntimPer <> 0 then
                                            'ntimPerUse = ntimPer/antalDageMtimer
                                            normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
                                            else
                                            normTimerGns5 = 1
                                            end if 

                                            dagval = hoursval / normTimerGns5
                                            'dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
                                            'intMidsRejsedage(m) = intMidsRejsedage(m) + dageVal/1
                                            rejsedageTot(m) = rejsedageTot(m) + dagval
                                    end select

                                    if lto = "cst" then
                                        if oRec("tfaktim") = 7 then
                                            bgColor = "#E7A1EF"
                                            giff = "dot_E7A1EF.gif"

                                            title = feriekalender_txt_015
                                            hoursval = oRec("timer")
                                            dagval = hoursval / nomrTimerPrDag
                                            'dageVal = formatnumber(oRec("timer")/(nomrTimerPrDag), 2)
                                        end if
                                        

                                    end if
                                    'antalTimer(month(oRec("tdato")), useYear, day(oRec("tdato")), oRec("tfaktim")) = oRec("timer") &"_"& oRec("timerkom") &"_"&oRec("tastedato")

                                    divTitle = title
                                    divText = oRec("tdato") &" "& feriekalender_txt_034 &": "& hoursval

                                    'antalTimer(month(oRec("tdato")), useYear, day(oRec("tdato")), oRec("tfaktim")) = "<img src='../ill/"&giff&"' style='width:10px; border-radius:5px;' /> <div class='"& boxclass &"' style='height:12px; width:12px; display:inline-block; cursor:pointer; margin:auto; background-color:"&bgColor&"; border-radius:5px;' class='ui-popover' data-toggle='tooltip' data-placement='top' data-trigger='hover' data-content='"&divText&"' title='"&divTitle&"'></div><br>"
                                    antalTimer(month(oRec("tdato")), useYear, day(oRec("tdato")), oRec("tfaktim")) = "<div><img src='../ill/"&giff&"' style='width:"& boxSize &"; border-radius:5px; cursor:pointer;' class='ui-popover' data-trigger='hover' data-content='"&divText&"' title='"&divTitle&"' /></div>"
                                    exportTxt(month(oRec("tdato")), useYear, day(oRec("tdato")), oRec("tfaktim")) = title &";"& hoursval &";"& formatnumber(dagval,2) &";"& oRec("tdato")

                                    lastyear = thisyear
                                    antalReg = antalReg + 1
                                oRec.movenext
                                wend
                                oRec.close

                                response.Write "<tr> <td style='border:solid 1px #ddd !important; padding-left:10px !important;'>"&medarbnavn(m)&" <br><br></td>"

                                lastYear = 0
                                thisYear = 0
                                loopDate = stDato
                                useYear = 1
                                
                                for d = 0 TO (antalDage -1)

                                    thisYear = year(loopDate)                                    

                                    if d <> 0 then
                                        loopDate = dateAdd("d", 1, loopDate)
                                        if thisYear <> lastYear then
                                            useYear = 2
                                        end if
                                    end if

                                    lastDayOfMonthDateToUse = thisYear&"-"&month(loopDate)&"-01"
                                    lastDayOfMonth = DateAdd("m", 1, lastDayOfMonthDateToUse)
                                    lastDayOfMonth = DateADd("d", -1, lastDayOfMonth) 

                                    if loopDate = lastDayOfMonth then
                                        tdBorder = "border-right:solid 1px #ddd !important;'"
                                    else
                                        tdBorder = ""
                                    end if

                                    dayNum = DatePart("w",loopDate ,2 ,2)
                                    if dayNum = 6 OR dayNum = 7 then
                                        tdbgColor = "#d4d4d4"
                                    else
                                        tdbgColor = ""
                                    end if

                                    response.Write "<td style='text-align:center; border-bottom:solid 1px #ddd !important; min-width:"& caltdWidth &"; background-color:"& tdbgColor &";'>"
                                    insideTd = ""
                                    for tfak = 0 TO 140
                                        timerDag = antalTimer(month(loopDate), useYear, day(loopDate), tfak)
                                        
                                        if timerDag <> "" then
                                            insideTd = insideTd & timerDag

                                            ekspTxt = ekspTxt & medarbnavn(m) &";"& exportTxt(month(loopDate), useYear, day(loopDate), tfak) & chr(013)

                                        end if

                                    next
                                    

                                    response.Write insideTd  & "</td>"

                                    lastYear = thisYear
                                next


                                response.Write "</tr>"
                            end if

                        next
                        %>

                        <%if media = "print" then %>
                            <tfoot>
                                <th colspan="500">
                                    <img src="../ill/dot_graae.gif" style="width:<%=boxSize%>; display:inline-block; border-radius:5px;" /> <%=feriekalender_txt_016 %>,
                                    <img src="../ill/dot_gron.gif" style="width:<%=boxSize%>; display:inline-block; border-radius:5px;" /> <%=feriekalender_txt_017 %>,
                                    <img src="../ill/dot_darkpink.gif" style="width:<%=boxSize%>; display:inline-block; border-radius:5px;" /> <%=feriekalender_txt_018 %>,
                                    <img src="../ill/dot_yellowgron.gif" style="width:<%=boxSize%>; display:inline-block; border-radius:5px;" /> <%=feriekalender_txt_019 %>,
                                    <img src="../ill/dot_rod.gif" style="width:<%=boxSize%>; display:inline-block; border-radius:5px;" /> <%=feriekalender_txt_012 %>,
                                    <img src="../ill/dot_orange.gif" style="width:<%=boxSize%>; display:inline-block; border-radius:5px;" /> <%=feriekalender_txt_013 %>,
                                    <img src="../ill/dot_blaa.gif" style="width:<%=boxSize%>; display:inline-block; border-radius:5px;" /> <%=feriekalender_txt_014 %>,
                                    <img src="../ill/dot_gul.gif" style="width:<%=boxSize%>; display:inline-block; border-radius:5px;" /> <%=feriekalender_txt_020 %>,
                                    <img src="../ill/dot_lightpink.gif" style="width:<%=boxSize%>; display:inline-block; border-radius:5px;" /> <%=feriekalender_txt_021 %>,

                                    <img src="../ill/dot_E7A1EF.gif" style="width:<%=boxSize%>; display:inline-block; border-radius:5px;" /> <%=feriekalender_txt_022 %>

                                    <%if lto = "cst" then %>
                                    ,<img src="../ill/dot_E7A1EF.gif" style="width:<%=boxSize%>; display:inline-block; border-radius:5px;" /> <%=feriekalender_txt_023 %>
                                    <%end if %>

                                </th>
                            </tfoot>
                        <%end if %>

                    </table>
                    <%if tableid = "x" AND media <> "print" then %>
                    </div>
                    <%end if %>
                    
                    


                    <%if media <> "print" then %>
                    <br /><br />
                    <%
                    if request("extype") = "sum" then 
                        ekspTxt = ""
                    end if
                    %>

                    <div class="row">
                        <div class="col-lg-12">
                            <table class="table table-bordered table-striped">
                                <thead>

                                    <tr>
                                        <th><%=feriekalender_txt_024 %></th>
                                        <th colspan="10"><%=feriekalender_txt_025 %></th>
                                        <th colspan="1"><%=day(startDatoFEOSQL) &"-"& month(startDatoFEOSQL) &"-"& year(startDatoFEOSQL) %> - <%=day(slutDatoFEOSQL) &"-"& month(slutDatoFEOSQL) &"-"& year(slutDatoFEOSQL) %></th>
                                        <th colspan="1"><%=day(startferiefriaar) &"-"& month(startferiefriaar) &"-"& year(startferiefriaar) %> - <%=day(slutferiefriaar) &"-"& month(slutferiefriaar) &"-"& year(slutferiefriaar) %></th>
                                    </tr>

                                    <tr>
                                        <th></th>
                                        <th><%=feriekalender_txt_017 %></th>
                                        <th><%=feriekalender_txt_016 %></th>
                                        <th><%=feriekalender_txt_018 %></th>
                                        <th><%=feriekalender_txt_019 %></th>

                                        <th><%=feriekalender_txt_020 %></th>
                                        <th><%=feriekalender_txt_021 %></th>

                                        <th><%=feriekalender_txt_012 %></th>
                                        <th><%=feriekalender_txt_014 %></th>
                                        <th><%=feriekalender_txt_027 %></th>
                                        <th><%=feriekalender_txt_022 %></th>

                                        <th><%=feriekalender_txt_026 %></th>
                                        <th><%=feriekalender_txt_028 %></th>
                                    </tr>
                                </thead>

                                <tbody>
                                    <%
                                    for m = 0 TO UBOUND(intmids)
                                    %>
                                    <tr>
                                        <th><%=medarbnavn(m) %></th>
                                        <th><%=formatnumber(ferieafholdtTot(m), 1) %></th>
                                        <th><%=formatnumber(ferieplanlagt(m), 1) %></th>
                                        <th><%=formatnumber(ferieudbetalt(m), 1) %></th>
                                        <th><%=formatnumber(ferieulon(m), 1) %></th>

                                        <th><%=formatnumber(feriefribrugt(m), 1) %></th>
                                        <th><%=formatnumber(feriefriudbetalt(m), 1) %></th>

                                        <th><%=formatnumber(sygPertot(m), 1) %></th>
                                        <th><%=formatnumber(barnsygPertot(m), 1) %></th>
                                        <th><%=formatnumber(afspadPertot(m), 1) %></th>
                                        <th><%=formatnumber(rejsedageTot(m), 1) %></th>

                                        <th><%=formatnumber(ferieaarsaldo(m), 1) %></th>
                                        <th><%=formatnumber(feriefriaarsaldo(m), 1) %></th>
                                    </tr>
                                    <%
                                        if request("extype") = "sum" then 
                                            if m <> 0 then
                                                ekspTxt = ekspTxt & chr(013)
                                            end if

                                            ekspTxt = ekspTxt & medarbnavn(m) &";"& formatnumber(ferieafholdtTot(m), 1) &";"
                                            ekspTxt = ekspTxt & formatnumber(ferieplanlagt(m), 1) &";"
                                            ekspTxt = ekspTxt & formatnumber(ferieudbetalt(m), 1) &";"
                                            ekspTxt = ekspTxt & formatnumber(ferieulon(m), 1) &";"
                                            ekspTxt = ekspTxt & formatnumber(feriefribrugt(m), 1) &";"
                                            ekspTxt = ekspTxt & formatnumber(feriefriudbetalt(m), 1) &";"
                                            ekspTxt = ekspTxt & formatnumber(sygPertot(m), 1) &";"
                                            ekspTxt = ekspTxt & formatnumber(barnsygPertot(m), 1) &";"
                                            ekspTxt = ekspTxt & formatnumber(afspadPertot(m), 1) &";"
                                            ekspTxt = ekspTxt & formatnumber(rejsedageTot(m), 1) &";"
                                            ekspTxt = ekspTxt & formatnumber(ferieaarsaldo(m), 1) &";"
                                            ekspTxt = ekspTxt & formatnumber(feriefriaarsaldo(m), 1) &";"
                                        end if
                                    next
                                    %>
                                </tbody>

                            </table>
                        </div>
                    </div>

                   <br />
                   <%end if %>

                    <%if media <> "print" then %>
                   <div class="row">
                       <div class="col-lg-3"><b><%=feriekalender_txt_029 %></b></div>
                   </div>

                    <form action="feriekalender.asp?media=print&FM_medarb=<%=thisMiduse %>&yuse=<%=ysel %>&FM_start_mrd=<%=strMrd %>&per_interval=3" method="post" target="_blank">
                        <div class="row">
                            <div class="col-lg-3">
                                <button class="btn btn-sm btn-success" style="width:200px;"><b><%=feriekalender_txt_030 %></b></button>
                            </div>
                        </div>
                    </form>

                    <form action="feriekalender.asp?media=print&FM_medarb=<%=thisMiduse %>&yuse=<%=ysel %>&FM_start_mrd=<%=strMrd %>&per_interval=1" method="post" target="_blank">
                        <div class="row">
                            <div class="col-lg-3">
                                <button class="btn btn-sm btn-success" style="width:200px;"><b><%=feriekalender_txt_031 %></b></button>
                            </div>
                        </div>
                    </form>

                    <form action="feriekalender.asp?media=export&FM_medarb=<%=thisMiduse %>&yuse=<%=ysel %>&FM_start_mrd=<%=strMrd %>&per_interval=1" method="post">
                        <div class="row">
                            <div class="col-lg-3">
                                <button class="btn btn-sm btn-secondary" style="width:200px;"><b><%=feriekalender_txt_032 %></b></button>
                            </div>
                        </div>
                    </form>

                    <form action="feriekalender.asp?media=export&FM_medarb=<%=thisMiduse %>&yuse=<%=ysel %>&FM_start_mrd=<%=strMrd %>&per_interval=1&extype=sum" method="post">
                        <div class="row">
                            <div class="col-lg-3">
                                <button class="btn btn-sm btn-secondary" style="width:200px;"><b><%=feriekalender_txt_033 %></b></button>
                            </div>
                        </div>
                    </form>

                    <%end if %>

                    <%
                        if media = "print" then
                            Response.Write("<script language=""JavaScript"">window.print();</script>")     
                        end if
                    %>



                  <!-- <table id="xmain_datatable_forecast" class="table table-condensed table-striped table-bordered">
                        <thead>

                          <!-,-  <tr>
                                <th>1</th>
                                <th colspan="31">Hej</th>
                            </tr> -,->

                            <tr>
                                 <th>a</th>

                                <th style="max-width:1px;">a</th>

                                <th>a</th>

                                 <th>a</th>
                                <th>a</th>

                                 <th>HEADer2</th>
                                 <th>HEADer3</th>
                                 <th>HEADer4</th>
                                <th>HEADer1</th>
                                 <th>HEADer2</th>
                                 <th>HEADer3</th>
                                 <th>HEADer4</th>
                                <th>HEADer1</th>
                                 <th>HEADer2</th>
                                 <th>HEADer3</th>
                                 <th>HEADer4</th>

                                <th>HEADer1</th>
                                 <th>HEADer2</th>
                                 <th>HEADer3</th>
                                 <th>HEADer4</th>
                                <th>HEADer1</th>
                                 <th>HEADer2</th>
                                 <th>HEADer3</th>
                                 <th>HEADer4</th>
                                <th>HEADer1</th>
                                 <th>HEADer2</th>
                                 <th>HEADer3</th>
                                 <th>HEADer4</th>
                                <th>HEADer1</th>
                                 <th>HEADer2</th>
                                 <th>HEADer3</th>
                                 <th>HEADer4</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                 <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                 <td>a</td>

                                <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                                 <td>a</td>
                            </tr>
                        </tbody>
                    </table> -->

                </div>

            </div>
        </div>



    </div>
</div>


<%

if media = "export" then
    
    if request("extype") = "sum" then
        'strOskrifter = "Medarbejder;Ferie Afholdt;Ferie planlagt;Ferie udbetalt;Ferie afh. u. lon;Feriefri afholdt; Feriefri udbetalt;Syg;Barnsyg;Afspad;Rejsedage"
        strOskrifter = feriekalender_txt_024 &";"&feriekalender_txt_017&";"&feriekalender_txt_016&";"&feriekalender_txt_018&";"&feriekalender_txt_019&";"&feriekalender_txt_020&";"&feriekalender_txt_021&";"&feriekalender_txt_012&";"&feriekalender_txt_013&";"&feriekalender_txt_027&";"&feriekalender_txt_022&";"&feriekalender_txt_026&";"&feriekalender_txt_028
    else
        'strOskrifter = "Medarbejder;Type;Timer;Dage;Dato"
        strOskrifter = feriekalender_txt_024&";"&feriekalender_txt_037&";"&feriekalender_txt_034&";"&feriekalender_txt_035&";"&feriekalender_txt_036&";"
    end if

    ekspTxt = ekspTxt

    call TimeOutVersion()
    

	ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	datointerval = ysel &", "& strMrd & " +"& per_interval & " m�neder"
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	Set objFSO = server.createobject("Scripting.FileSystemObject")
				
	if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\feriekalender.asp" then
		Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\feriekalenderexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\feriekalenderexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
	else
		Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\feriekalenderexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\feriekalenderexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
	end if
				
				
				
	file = "feriekalenderexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
	'**** Eksport fil, kolonne overskrifter ***
				
				
				
	'objF.writeLine("Periode afgr�nsning: "& datointerval & vbcrlf)

    if vis <> 1 then
	objF.WriteLine(strOskrifter & chr(013))
    end if

	objF.WriteLine(ekspTxt)
	objF.close		

	Response.redirect "../inc/log/data/"& file &""	
	Response.end


end if

%>






    

<!--#include file="../inc/regular/footer_inc.asp"-->