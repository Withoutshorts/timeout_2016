

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%
    if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")

            case "findaktiviteter"

                jobnr = request("jobnr")                
                stropt = ""
                strSQL = "SELECT aktiviteter.id as aktid, navn, jobnr FROM aktiviteter LEFT JOIN job ON (aktiviteter.job = job.id) WHERE jobnr = "& jobnr                
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
                    stropt = stropt & "<option value="&oRec("aktid")&">"&oRec("navn")&"</option>"
                oRec.movenext
                wend
                oRec.close

                call jq_format(stropt)
                stropt = jq_formatTxt

                response.Write stropt
            
        end select
    response.end
    end if   
%>















<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../timereg/inc/convertDate.asp"-->
<!--#include file="../timereg/inc/timereg_akt_2006_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<%if request("func") = "print" OR request("func") = "print2" then %>
<style>
    .container {
        zoom:74%;
    }
</style>
<%end if %>

<%if request("func") <> "print" AND request("func") <> "print2" then %>
<div class="wrapper">
<div class="content">
<%end if %>
<% 
    if len(session("user")) = 0 then

    errortype = 5
    call showError(errortype)
        response.End
    end if
 
    func = request("func")

    if func <> "print" AND func <> "print2" then
    call menu_2014()
    end if

    varTjDatoUS_man = request("varTjDatoUS_man")
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

    datoMan = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
    datoSon = day(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& year(varTjDatoUS_son)

    next_varTjDatoUS_man = dateadd("d", 7, varTjDatoUS_man)
	next_varTjDatoUS_son = dateadd("d", 7, varTjDatoUS_son)
    next_varTjDatoUS_man = year(next_varTjDatoUS_man) &"/"& month(next_varTjDatoUS_man) &"/"& day(next_varTjDatoUS_man)
	next_varTjDatoUS_son = year(next_varTjDatoUS_son) &"/"& month(next_varTjDatoUS_son) &"/"& day(next_varTjDatoUS_son)


    prev_varTjDatoUS_man = dateadd("d", -7, varTjDatoUS_man)
	prev_varTjDatoUS_son = dateadd("d", -7, varTjDatoUS_son)
    prev_varTjDatoUS_man = year(prev_varTjDatoUS_man) &"/"& month(prev_varTjDatoUS_man) &"/"& day(prev_varTjDatoUS_man)
	prev_varTjDatoUS_son = year(prev_varTjDatoUS_son) &"/"& month(prev_varTjDatoUS_son) &"/"& day(prev_varTjDatoUS_son)

    strYear = year(varTjDatoUS_man)

    select case cint(month(varTjDatoUS_man))
        case 1
        strMonth = ugetast_txt_020
        case 2
            strMonth = ugetast_txt_021
        case 3
            strMonth = ugetast_txt_022
        case 4
            strMonth = ugetast_txt_023
        case 5
            strMonth = ugetast_txt_024
        case 6
            strMonth = ugetast_txt_025
        case 7
            strMonth = ugetast_txt_026
        case 8
            strMonth = ugetast_txt_027
        case 9
            strMonth = ugetast_txt_028
        case 10
            strMonth = ugetast_txt_029
        case 11
            strMonth = ugetast_txt_030
        case 12
            strMonth = ugetast_txt_031
    end select

    call thisWeekNo53_fn(varTjDatoUS_man)
    strWeek = thisWeekNo53 'datepart("ww", varTjDatoUS_man, 2,2)

    'medid = session("mid")

    call fTeamleder(session("mid"), 0, 0)

    if erTeamLeder = 1 OR level = 1 then
        if len(trim(request("FM_medid"))) <> 0 then
        medid = request("FM_medid")
        else
        medid = session("mid") 
        end if
    else
        medid = session("mid") 
    end if

    usemrn = medid
    

    if func = "db" then
        feltnr = request("FM_feltnr")
        feltnr = replace(feltnr, " ", "")
        feltnr = split(feltnr, ",")
        
        medid = usemrn
        call meStamdata(medid)

        strMnavn = meNavn

        for f = 0 TO UBOUND(feltnr)
            
            tdato = request("FM_dato_" & feltnr(f))
            jobnr = request("FM_jobnr_" & feltnr(f))
            aktid = request("FM_aktid_" & feltnr(f))

            sttid = request("FM_sttid_" & feltnr(f))
            sltid = request("FM_sltid_" & feltnr(f))

            timeid = request("FM_timeid_" & feltnr(f))

            
            'response.Write "<br> ---------------- feltnr "& feltnr(f) &" tdato " & tdato & " jobnr " & jobnr & " aktid " & aktid & " sttid " & sttid & " sltid " & sltid

            if jobnr <> "-1" AND aktid <> "-1" then 'Valdiering
                strSQL = "SELECT jobnavn, jobknr, k.kkundenavn as kundenavn, j.fastpris as fastpris, j.serviceaft as serviceaft FROM job j LEFT JOIN kunder k ON (j.jobknr = k.kid) WHERE jobnr ='"& jobnr &"'"
                'response.Write strSQL
                oRec.open strSQL, oConn, 3
                if not oRec.EOF then
                    strJobknavn = oRec("kundenavn")
                    strJobnavn = oRec("jobnavn")
                    strJobknr = oRec("jobknr")
                    strFastpris = oRec("fastpris")
                    intServiceAft = oRec("serviceaft")
                end if
                oRec.close

                dblkostpris = 0

                strSQL = "SELECT navn, fakturerbar, kostpristarif FROM aktiviteter WHERE id = "& aktid
                'response.Write strSQL
                oRec.open strSQL, oConn, 3
                if not oRec.EOF then
                    strAktNavn = oRec("navn")
                    tfaktimvalue = oRec("fakturerbar")

                    kostpristarif = oRec("kostpristarif")
                    select case kostpristarif
                    case "0"
                    dblkostpris = dblkostpris
                    case "A"
                    dblkostpris = mkostpristarif_A
                    case "B"
                    dblkostpris = mkostpristarif_B
                    case "C"
                    dblkostpris = mkostpristarif_C
                    case "D"
                    dblkostpris = mkostpristarif_D
                    case else
                    dblkostpris = dblkostpris 
                    end select

                oRec.close
                end if

                intJobnr = jobnr
                useDato = tdato

                if len(sttid) <> 0 AND len(sltid) <> 0 then

                if InStr(sttid, ":") = 0 then
                    sttid = sttid & ":00"
                end if

                if InStr(sltid, ":") = 0 then
                    sltid = sltid & ":00"
                end if

                totalminutes = datediff("n", sttid, sltid)
                useTimer = totalminutes/60
                useTimer = replace(useTimer, ",", ".")
                else
                useTimer = 0
                end if

                'useTimer = 2

                strKomm = ""
                if len(tprisGen) <> 0 then
                intTimepris = replace(tprisGen, ",", ".")
                else
                intTimepris = 0
                end if
                offentlig = 0
                strYear = year(now)
                visTimerelTid = 2 'S� start og slut tiderne bliver lagt i tabellen
                stopur = 0
                intValuta = valutaGen
                bopal = 0
                destination = ""
                useDage = ""
                tildeliheledage = 0
                origin = 30
                extsysid = 0
                mtrx = 0

                useDato = dateadd("d",0,useDato) 'Rigtig format

            
                call opdaterTimer(aktid, strAktNavn, tfaktimvalue, strFastpris, intJobnr, strJobnavn, strJobknr,_
		        strJobknavn, medid, strMnavn,_
		        useDato, useTimer, strKomm, intTimepris,_
		        dblkostpris, offentlig, intServiceAft, strYear,_
		        sttid, sltid, visTimerelTid, stopur, intValuta, bopal, destination, useDage, tildeliheledage, origin, extsysid, mtrx, intKpValuta, timeid)
            
            

            'response.Write "<br><br><br>----------------------------------------------- aktid " & aktid & " strAktNavn " & strAktNavn & " tfaktimvalue " & tfaktimvalue & " strFastpris " & strFastpris & " intJobnr " & intJobnr & " strJobnavn " & strJobnavn & " strJobknr "& strJobknr

		    'response.Write " strJobknavn " & strJobknavn & " medid " & medid & "strMnavn" & strMnavn

		    'response.Write " useDato " & useDato &" useTimer "& useTimer &" strKomm "& strKomm & " intTimepris " & intTimepris

		    'response.Write  " dblkostpris " & dblkostpris & " offentlig " & offentlig & " intServiceAft " & intServiceAft & " strYear " & strYear

		    'response.Write  " sttid " & sttid & " sltid " & sltid & " visTimerelTid " & visTimerelTid & " stopur " & stopur & " intValuta " & intValuta & " bopal " & bopal & " destination " & destination & " useDage " & useDage & " tildeliheledage " & tildeliheledage & " origin " & origin & " extsysid " & extsysid & " mtrx "& mtrx & " intKpValuta " & intKpValuta & " totalminutes " & totalminutes
            end if 'validering
        next

        'response.End
        response.Redirect "ugetast.asp?varTjDatoUS_man="& request("varTjDatoUS_man") &"&FM_medid="& medid
    end if



    call meStamdata(usemrn)

%>

    <script src="js/ugetast_jav2.js" type="text/javascript"></script>

    <div class="container">
        <div class="portlet">
        <!-- <h3 class="portlet-title"><u>Uge indtastning:</u></h3> -->
            <div class="portlet-body">

                <%'if func <> "print" then %>

                <%if func <> "print" AND func <> "print2" then %>
                    <%if erTeamleder = 1 OR level = 1 then %>
                    <form action="ugetast.asp?varTjDatoUS_man=<%=request("varTjDatoUS_man")%>" method="post">
                        <input type="hidden" name="varTjDatoUS_man" value="<%=request("varTjDatoUS_man") %>" />

                        <div class="row">
                            <div class="col-lg-4"></div>
                            <div class="col-lg-4">
                                <%
                                select case cint(level) 
                                case 1 
                                sqgMw = " mansat <> 2 AND mansat <> 4 "
                                case 2,6

                                    call medarb_teamlederfor

                                    sqgMw = " mansat <> 2 AND mansat <> 4 "& medarbgrpIdSQLkri
                                case else
                                    sqgMw = " mid = "& session("mid") &" AND mansat <> 2 AND mansat <> 4"
                                end select

                                strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe, init FROM medarbejdere WHERE "& sqgMw &" GROUP BY mid ORDER BY Mnavn" 
                                %>
                                <select name="FM_medid" id="FM_medid" <%=progrpmedarbDisabled  %> class="form-control input-small"  onchange="submit();" style="width:240px">
                                    <%

                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF
                                
                                    
                                    StrMnavn = oRec("Mnavn")
                                    StrMinit = oRec("init")
	
                                    if cdbl(medid) = cdbl(oRec("Mid")) then
				                    isSelected = "SELECTED"
                                    else
				                    isSelected = ""
				                    end if

				                    %>
                                        <option value="<%=oRec("Mid")%>" <%=isSelected%>><%=StrMnavn &" ["& StrMinit & "]"%></option>
                                    <%
                                    oRec.movenext
                                    wend
                                    oRec.close  
                                    %>
                                </select>
                            </div>
                        </div>
                    </form>
                    <%end if %>
                <%end if %>

                <div class="row">
                    <div class="col-lg-12" style="text-align:center"><img src="img/flash_ugeseddel.PNG" style="width:100%;" /></div>
                </div>

                <br /><br />
                <%'end if %>

                <%if func = "print" or func = "print2" then 
                    tableStyle = "border:solid 2px; border-right:solid 2px;"

                    %>
                    <style>
                        table, th, td {
                            border-color:black !important;
                        }
                    </style>
                    <%

                else 
                    tableClass = "table-bordered"
                    tableStyle = ""
                %>

                <%end if %>

                <!-- <table style="width:100%">
                    <tr>
                        <td style="background-color:#64727b;" colspan="2"><span style="color:white; font-size:250%; padding-left:25px;"><b>Ugeseddel</b></span></td>
                    </tr>

                    <tr>
                        <td style="background-color:#07619c;" rowspan="4"><img src="img/flash_ikon.png" style="width:50%;" /></td>
                    </tr>
                    <tr>
                        <td style="background-color:#07619c;"><span style="color:white; font-size:175%; padding-left:75px;"><b>Ouh</b></span></td>
                    </tr>
                     <tr>
                        <td style="background-color:#07619c;"><span style="color:white; font-size:175%; padding-left:75px;"><b>Odense</b></span></td>
                    </tr>
                    <tr>
                        <td style="background-color:#07619c;"><span style="color:white; font-size:175%; padding-left:75px;"><b>Svenborg</b></span></td>
                    </tr>
                </table> -->

                <table class="table <%=tableClass %>" style="<%=tableStyle%>">
                    <tr>
                        <td style="width:50%; height:50px; border-top:hidden; border-left:hidden;"><span style="font-size:125%"><%=ugetast_txt_043 %></span></td>
                        <td style="width:25%; <%=tableStyle%>"><b><%=ugetast_txt_001 %>:</b>&nbsp <%=meCPR %></td>
                        <td style="width:25%; <%=tableStyle%>"><b><%=ugetast_txt_002 %>:</b></td>
                    </tr>
                    <tr>

                        <%
                        strnavn = ""
                        fornavn = ""
                        efternavn = ""
                        navne = split(Menavn, " ")
                        antalEfternavne = UBOUND(navne)
                        for n = 0 TO UBOUND(navne)
                            if n < antalEfternavne then
                            fornavn = fornavn & " " & navne(n)
                            else
                            efternavn = navne(n)
                            end if
                        next
                        %>

                        <td style="<%=tableStyle%> height:50px;"><b><%=ugetast_txt_003 %>:</b>&nbsp <%=efternavn %></td>
                        <td style="<%=tableStyle%>" colspan="2"><b><%=ugetast_txt_004 %>:</b>&nbsp <%=fornavn %></td>
                    </tr>
                    <tr>
                        <td style="<%=tableStyle%> height:50px;"><b><%=ugetast_txt_019 %>:</b></td>
                        <td style="<%=tableStyle%>" colspan="2"><b><%=ugetast_txt_005 %></b>: <br />
                            &nbsp Odense <input type="checkbox" name="FM_matrikel" />
                            &nbsp Svendborg <input type="checkbox" name="FM_matrikel" />
                        </td>
                    </tr>
                    <tr>
                        <td style="height:50px;" colspan="3"><b><%=ugetast_txt_006 %>:</b></td>
                    </tr>
                </table>
                
                <%if func <> "print" then %>
                <div class="row">
                    <h3 class="col-lg-12" style="text-align:center;">
                       <%if func <> "print" and func <> "print2" then %><a href="ugetast.asp?varTjDatoUS_man=<%=prev_varTjDatoUS_man %>&FM_medid=<%=medid %>" class="btn btn-default btn-sm" ><b><</b></a><%end if %>
                        &nbsp&nbsp&nbsp<%=strMonth &" "& strYear &" - "& ugetast_txt_007 &" "& strWeek %>&nbsp&nbsp&nbsp
                        <%if func <> "print" and func <> "print2" then %><a href="ugetast.asp?varTjDatoUS_man=<%=next_varTjDatoUS_man %>&FM_medid=<%=medid %>" class="btn btn-default btn-sm" ><b>></b></a><%end if %>
                    </h3>
                </div>
                <%end if %>

                <!-- Reg delen -->
                <form action="ugetast.asp?func=db" method="post">
                    <input type="hidden" name="varTjDatoUS_man" value="<%=varTjDatoUS_man %>" />
                    <input type="hidden" name="FM_medid" value="<%=medid %>" />
    
                    <%
                    if func = "print" then
                        tableBorder = "border-bottom:solid 2px;"
                        headerBorder = "border:solid 2px;"
                        textalign = "text-align:center;"
                    else
                        tableBorder = ""
                        headerBorder = ""
                        textalign = ""

                        if func = "print2" then
                            tableBorder = "border:solid 2px;"
                            textalign = "text-align:center;"
                            headerBorder = "border:solid 2px;"
                        end if

                    end if
                    %>

                    <table class="table <%=tableClass %>" style="<%=tableBorder%>;">

                        <%if func = "print" then %>
                        <tr style="height:50px;">       
                            <td style="vertical-align:middle; text-align:center; <%=headerBorder%>"></td>
                            <td style="vertical-align:middle; text-align:center; <%=headerBorder%>"><b><%=ugetast_txt_039 %>:</b> <%=strYear %></td>
                            <td style="vertical-align:middle; text-align:center; <%=headerBorder%>"><b><%=ugetast_txt_040 %>:</b> <%=strMonth %></td>
                            <td style="vertical-align:middle; text-align:center; <%=headerBorder%>"><b><%=ugetast_txt_007 %>:</b> <%=strWeek %></td>
                        </tr>
                        <%end if %>

                        <tr style="background-color:#f9f9f9;">
                            <td style="text-align:center; <%=headerBorder%>"><b><%=ugetast_txt_008 %></b></td>
                            <td style="<%=headerBorder%> <%=textalign%>"><b><%=ugetast_txt_009 %></b></td>
                            <td style="<%=headerBorder%> <%=textalign%>"><b><%=ugetast_txt_010 %></b></td>
                            <%if func <> "print" then %>

                            <%
                            if func = "print2" then
                                textAlign = "text-align:center;"
                            else
                                textAlign = ""
                            end if
                            %>

                            <td style="width:25%; <%=headerBorder%> <%=textAlign%>"><b><%=ugetast_txt_011 %></b></td>
                            <td style="width:25%; <%=headerBorder%> <%=textAlign%>"><b><%=ugetast_txt_012 %></b></td>
                            <%else %>
                            <td style="width:35%; border-right:solid 2px; text-align:center;"><b><%=ugetast_txt_013 %></b></td>
                            <%end if %>
                        </tr>

                        <%
                        stroptJob = "<option value=""-1"">"&ugetast_txt_014&"</option>"
                        strSQL = "SELECT id, jobnavn, jobnr FROM job WHERE jobstatus = 1"
                        oRec.open strSQL, oConn, 3
                        while not oRec.EOF
                            stroptJob = stroptJob & "<option value="&oRec("jobnr")&">"& oRec("jobnavn") &"</option>"
                        oRec.movenext
                        wend
                        oRec.close


                        response.Write "<input type='hidden' value='"&stroptJob&"' id='defaultjobopr' />"

                        perIntervalLoop = 6
                        feltnr = 0
                        for d = 0 to perIntervalLoop  

                            if d = 0 then
                            varTjDatoUS_use = varTjDatoUS_man
                            varTjDatoUS_use = day(varTjDatoUS_use) &"-"& month(varTjDatoUS_use) &"-"& year(varTjDatoUS_use)
                            else
                            varTjDatoUS_use = dateAdd("d", d, varTjDatoUS_man)
                            end if

                            varTjDatoUS_useSQL = year(varTjDatoUS_use) &"-"& month(varTjDatoUS_use) &"-"& day(varTjDatoUS_use)

                            select case cint(weekday(varTjDatoUS_use, 0))
                                case 1
                                    strDay = ugetast_txt_032
                                case 2
                                    strDay = ugetast_txt_033
                                case 3
                                    strDay = ugetast_txt_034
                                case 4
                                    strDay = ugetast_txt_035
                                case 5
                                    strDay = ugetast_txt_036
                                case 6
                                    strDay = ugetast_txt_037
                                case 7
                                    strDay = ugetast_txt_038
                            end select

                            'strDay = weekdayname(weekday(varTjDatoUS_use), false, 1)
                            strthisdate = day(varTjDatoUS_use) &"-"& month(varTjDatoUS_use) &"-"& year(varTjDatoUS_use)
                        %>

                            <%
                            rowCount = 1
                            strSQL = "SELECT count(tid) as totregs FROM timer WHERE tdato = '"& varTjDatoUS_useSQL &"' AND tmnr = "& usemrn & " AND origin = 30"
                            oRec.open strSQL, oConn, 3
                            if not oRec.EOF then
                            rowCount = oRec("totregs")
                            end if
                            oRec.close
                            
                            rowSpanAmount = rowCount
                            if func = "print" then
                                if rowCount < 3 then
                                    rowSpanAmount = 3
                                end if
                            end if

                            strSQL = "SELECT tid, timer, sttid, sltid, taktivitetnavn, taktivitetid, tjobnr, jobnavn, job.id as jobid FROM timer LEFT JOIN job ON (tjobnr = job.jobnr) WHERE tdato = '"& varTjDatoUS_useSQL &"' AND tmnr = "& usemrn & " AND origin = 30 ORDER BY sttid, sltid"
                            'response.Write strSQL
                            oRec.open strSQL, oConn, 3                            
                            r = 0
                            while not oRec.EOF
                                sttid = right(oRec("sttid"), 8)
                                sltid = right(oRec("sltid"), 8)
                                rowspan = 1

                                if func = "print" then
                                    borderStyleSolid = "border:solid 2px;"

                                    if r > 0 AND r < rowCount then
                                            borderStyleDashed =  "border-top:dashed 2px; border-right:solid 2px;"
                                    else
                                            borderStyleDashed = "border-top:solid 2px; border-right:solid 2px;"
                                    end if
                                else
                                    borderStyleSolid = ""
                                    borderStyleDashed = ""

                                    if func = "print2" then
                                        borderStyleSolid = "border-top:solid 2px; border-right:solid 2px;"
                                        borderStyleDashed = "border-top:solid 2px; border-right:solid 2px;"
                                    end if
                                end if

                            %> 
                                <tr class="tr_<%=d %>">

                                    <input type="hidden" name="FM_feltnr" value="<%=feltnr %>" />
                                    <input type="hidden" name="FM_dato_<%=feltnr %>" value="<%=varTjDatoUS_useSQL %>" />
                                    <input type="hidden" name="FM_timeid_<%=feltnr %>" value="<%=oRec("tid") %>" />

                                    <%if r = 0 then %>
                                    <td id="title_<%=d %>" rowspan="<%=rowSpanAmount %>" style="text-align:center; vertical-align:middle; <%=borderStyleSolid%>">
                                        <b>
                                        <%=strDay %> <br />
                                        <%=strthisdate %></b>
                                    </td>
                                    <%end if %>

                                    <td style="vertical-align:middle; text-align:center; <%=borderStyleDashed%>">
                                        <%if func <> "print" and func <> "print2" then %>
                                            <input name="FM_sttid_<%=feltnr %>" type="time" value="<%=sttid %>" class="form-control input-small" />
                                        <%else %>
                                            <span><%=left(sttid, 5)%></span>
                                        <%end if %>
                                    </td>
                                    <td style="vertical-align:middle; text-align:center; <%=borderStyleDashed%>"">
                                        <%if func <> "print" and func <> "print2" then %>
                                        <input name="FM_sltid_<%=feltnr %>" type="time" value="<%=sltid %>" class="form-control input-small" />
                                        <%else %>
                                            <span><%=left(sltid, 5)%></span>
                                        <%end if %>
                                    </td>
                                    
                                    <%if func <> "print" then %>

                                    <%
                                    if func = "print2" then
                                        textAlign = "text-align:center;"
                                    else
                                        textAlign = ""
                                    end if
                                    %>

                                    <td style="vertical-align:middle; <%=borderStyleSolid%> <%=textAlign%>"><%=oRec("jobnavn") %> <input type="hidden" name="FM_jobnr_<%=feltnr %>" value="<%=oRec("tjobnr") %>" /></td>

                                    <td style="vertical-align:middle; <%=borderStyleSolid%> <%=textAlign%>">
                                        <%if func <> "print2" then %>
                                       <select name="FM_aktid_<%=feltnr %>" class="form-control input-small">

                                            <%
                                            stroptaktreg = ""
                                            strSQL = "SELECT id, navn FROM aktiviteter WHERE job = "& oRec("jobid")
                                            oRec2.open strSQL, oConn, 3
                                            while not oRec2.EOF

                                                if cdbl(oRec2("id")) = cdbl(oRec("taktivitetid")) then
                                                    aktSEL = "SELECTED"
                                                else
                                                    aktSEL = ""
                                                end if

                                                stroptaktreg = stroptaktreg & "<option value='"&oRec2("id")&"' "&aktSEL&">"&oRec2("navn")&"</option>"
                                            oRec2.movenext
                                            wend
                                            oRec2.close
                                            %>

                                            <%=stroptaktreg %>
                                        </select>          
                                        <%else %>

                                        <%
                                        aktnavn = ""
                                        strSQL = "SELECT id, navn FROM aktiviteter WHERE id = "& oRec("taktivitetid")
                                        oRec2.open strSQL, oConn, 3
                                        if not oRec2.EOF then
                                            aktnavn = oRec2("navn")
                                        end if
                                        oRec2.close

                                        response.Write aktnavn
                                        %>

                                        <%end if %>
                                    </td>
                                    <%else %>
                                    <td style="<%=borderStyleDashed%>""></td>
                                    <%end if %>

                                    <%if r = 0 AND func <> "print" AND func <> "print2" then %>
                                    <td style="width:15px;"><span style="color:#afafaf;" id="<%=d %>" data-thisdate="<%=varTjDatoUS_useSQL %>" class="btn btn-sm btn-default fa fa-plus addRow"></span></td>
                                    <%end if %>
                                </tr>
                            <%
                            r = r + 1
                            feltnr = feltnr + 1
                            oRec.movenext
                            wend
                            oRec.close %>

                            <%
                            if func = "print" then
                                rowSpanAmount = 3
                            else
                                rowSpanAmount = 1
                            end if
                            %>

                            <%if r = 0 then %>

                                <%
                                if func = "print" or func = "print2" then
                                    borderStyle = "border-top:solid 2px; border-right:solid 2px; border-left:solid 2px;"
                                else
                                    borderStyle = ""
                                end if
                                %>
                                

                                <tr class="tr_<%=d %>">
                                    <input type="hidden" name="FM_feltnr" value="<%=feltnr %>" />
                                    <input type="hidden" name="FM_dato_<%=feltnr %>" value="<%=varTjDatoUS_useSQL %>" />
                                    <input type="hidden" name="FM_timeid_<%=feltnr %>" value="-1" />

                                    <td id="title_<%=d %>" rowspan="<%=rowSpanAmount %>" style="text-align:center; vertical-align:middle; <%=borderStyle%>">
                                        <b>
                                        <%=strDay %> <br />
                                        <%=strthisdate %></b>
                                    </td>

                                    <td style="vertical-align:middle; <%=borderStyle%>">
                                        <%if func <>"print" AND func <> "print2" then %>
                                            <input name="FM_sttid_<%=feltnr %>" type="time" value="" class="form-control input-small" />
                                         <%else %>
                                        <span>&nbsp</span>
                                        <%end if %>
                                    </td>

                                    <td style="vertical-align:middle; <%=borderStyle%>">
                                        <%if func <> "print" AND func <> "print2" then %>
                                            <input name="FM_sltid_<%=feltnr %>" type="time" value="" class="form-control input-small" />
                                        <%else %>
                                        <span>&nbsp</span>
                                        <%end if %>
                                    </td>

                                    <%if func <> "print" then %>
                                    <td style="vertical-align:middle; <%=borderStyle%>">
                                        <%if func <> "print2" then %>
                                        <select name="FM_jobnr_<%=feltnr %>" class="form-control input-small jobSEL" id="<%=feltnr %>">
                                            <%=stroptJob %>
                                        </select>
                                        <%end if %>
                                    </td>

                                    <td style="vertical-align:middle; <%=borderStyle%>">
                                        <%if func <> "print2" then %>
                                        <select name="FM_aktid_<%=feltnr %>" class="form-control input-small" id="aktSEL_<%=feltnr %>">
                                            <option value="-1">-</option>
                                        </select>
                                        <%end if %>
                                    </td>
                                    <%else %>
                                    <td style="<%=borderStyle%>"></td>
                                    <%end if %>

                                    <%if func <> "print" AND func <> "print2" then %>
                                    <td style="width:15px;"><span style="color:#afafaf;" id="<%=d %>" data-thisdate="<%=varTjDatoUS_useSQL %>" class="btn btn-sm btn-default fa fa-plus addRow"></span></td>
                                    <%end if %>
                                </tr>
                            <%feltnr = feltnr + 1 %>
                            <%end if %>
                            

                            <%
                            if func = "print" then
                                if rowCount = 0 OR rowCount = 1 then
                                    response.Write "<tr><td style='border-top:dashed 2px; border-right:solid 2px;'><span>&nbsp</span></td> <td style='border-top:dashed 2px; border-right:solid 2px;'><span>&nbsp</span></td> <td style='border-top:dashed 2px; border-right:solid 2px;'><span>&nbsp</span></td> </tr>"
                                    response.Write "<tr><td style='border-top:dashed 2px; border-right:solid 2px;'><span>&nbsp</span></td> <td style='border-top:dashed 2px; border-right:solid 2px;'><span>&nbsp</span></td> <td style='border-top:dashed 2px; border-right:solid 2px;'><span>&nbsp</span></td> </tr>"
                                end if

                                if rowCount = 2 then
                                    response.Write "<tr><td style='border-top:dashed 2px; border-right:solid 2px;'><span>&nbsp</span></td> <td style='border-top:dashed 2px; border-right:solid 2px;'><span>&nbsp</span></td> <td style='border-top:dashed 2px; border-right:solid 2px;'><span>&nbsp</span></td> </tr>"
                                end if
                            end if
                            %>

                        <%
                        next
                        %>

                    </table>
                    
                    <%if func <> "print" AND func <> "print2" then %>
                    <div class="row">
                        <div class="col-lg-4"><a href="ugetast.asp?varTjDatoUS_man=<%=varTjDatoUS_man %>&func=print&FM_medid=<%=medid %>" target="_blank" style="width:100px;" class="btn btn-default btn-sm"><b><%=ugetast_txt_015 %></b></a>
                        <a href="ugetast.asp?varTjDatoUS_man=<%=varTjDatoUS_man %>&func=print2&FM_medid=<%=medid %>" target="_blank" style="width:100px;" class="btn btn-default btn-sm"><b><%=ugetast_txt_041 %> </b></a></div>
                        <div class="col-lg-8"><button type="submit" style="width:100px;" class="btn btn-success btn-sm pull-right"><b><%=ugetast_txt_016 %></b></button></div>
                    </div>
                    <%end if %>

                </form>
                <input type="hidden" id="lastfeltnr" value="<%=feltnr %>" />

                <%
                if func = "print" or func = "print2" then
                    tableborderstyle = "border:solid 2px;"
                    tdborderstyle = "border-right:solid 2px;"
                else
                    tableborderstyle = ""
                    tdborderstyle = ""
                end if
                %>

                <%if func = "print" OR func = "print2" then %>
                <table class="table <%=tableClass %>" style="<%=tableborderstyle%>">
                    <tr>
                        <td style="width:50%; <%=tdborderstyle%>"><b><%=ugetast_txt_017 %> &nbsp /  &nbsp <%=ugetast_txt_042 %> </b> <br /><br /><br /><br /></td>
                        <td style="width:50%;"><b><%=ugetast_txt_018 %> &nbsp /  &nbsp <%=ugetast_txt_042 %>  </b> <br /><br /><br /><br /></td>
                    </tr>
                </table>
                <%end if %>

            </div>
        </div>
    </div>

    <%if func = "print" OR func = "print2" then %>

        <%Response.Write("<script language='JavaScript'>window.print();</script>") %>

    <%end if %>
<%if request("func") <> "print" AND request("func") <> "print2" then %>
</div>
</div>
<%end if %>
    

<!--#include file="../inc/regular/footer_inc.asp"-->