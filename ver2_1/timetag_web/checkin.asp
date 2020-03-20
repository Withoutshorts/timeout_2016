
<%
    thisfile = "timetag_mobile" 
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

      
    <%

    if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/errors/error_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	response.End
    end if

    call meStamdata(session("mid"))

    usemrn = session("mid")

    func = request("func")


    'Henter de aktivitetstyper der slået til at blive vidst som knap på mobilen
    strAktTyperKri = ""
    strSQL = "SELECT aty_id, aty_desc FROM akt_typer WHERE aty_on = 1 AND aty_mobile_btn = 1"
    oRec.open strSQL, oConn, 3
    at = 0
    while not oRec.EOF

        if at = 0 then
        strAktTyper = " tfaktim = " & oRec("aty_id")
        else
        strAktTyper = strAktTyper & " OR tfaktim = " & oRec("aty_id")
        end if

    at = at + 1
    oRec.movenext
    wend
    oRec.close

    if at > 0 then
    strAktTyperKri = " AND (" & strAktTyper & ")"
    else
    strAktTyperKri = " AND tfaktim = 9999"
    end if
                     

    'response.Write "strAktTyperKri " & strAktTyperKri



    'Tjekker om registreringer skal afsluttes
    strSQL = "SELECT tid, sttid, tdato FROM timer WHERE tmnr =" & session("mid") & strAktTyperKri &" AND timer = 0 AND sltid is null"
    'response.Write strSQL
    oRec.open strSQL, oConn, 3
    while not oRec.EOF


        startDato = year(oRec("tdato")) &"-"& month(oRec("tdato")) &"-"& day(oRec("tdato"))
        startTidspunkt = hour(oRec("sttid")) &":"& minute(oRec("sttid")) &":"& second(oRec("sttid"))

        slutDato = year(now)&"-"&month(now)&"-"&day(now)
        slutTidspunkt = hour(now) &":"& minute(now) &":"& second(now)

        start = startDato & " " & startTidspunkt
        slut = slutDato & " " & slutTidspunkt


        ' Henter timer siden start, bruger minutter og dividere med 60 for at få decimaler på.
        timerSidenStart = (DateDiff("n", start, slut, 2,2)) / 60
        timerSidenStart = Round(timerSidenStart, 2)
             
        'Tjekker om timer er mere end en dag.
        if timerSidenStart >= 7.4 then
        timerSidenStart = 7.4
        end if

        timerSidenStart = Replace(CStr(timerSidenStart), ",", ".")

        slutTidspunkt = hour(now) &":"& minute(now) &":"& second(now)

        oConn.execute("UPDATE timer SET timer = "& timerSidenStart &", sltid = '"& slutTidspunkt &"' WHERE tid = "& oRec("tid") &" AND tmnr = "& session("mid"))

    oRec.movenext
    wend
    oRec.close

    select case func 
    case "knap"

        akttype = request("FM_akttype")

        timeNow = hour(now)&":"&minute(now)&":"&second(now)

        oConn.execute("INSERT INTO timer SET tmnr = "& session("mid") &", tmnavn = '"& menavn &"', tdato = '"&year(now)&"-"&month(now)&"-"&day(now)&"', tastedato = '"&year(now)&"-"&month(now)&"-"&day(now)&"', timerkom = '', tfaktim = "& akttype &", sttid = '"& timeNow &"'" )            

        'response.Write "<br> skalregafsluttes" & skalregafsluttes

        response.Redirect "../sesaba.asp"


    case "logud"

        response.Redirect "../sesaba.asp"

    end select

    %>


    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">

    
    <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
    <!--#include file="../inc/regular/top_menu_mobile.asp"-->



<head>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TimeOut mobile</title>





    <style type="text/css">

        input[type="text"] 
        {
          height:125%;
          font-size:125%;
        }
        input[type="button"] 
        {
          height:125%;
          font-size:125%;
        }

        .span_job {
        list-style:none;
        font-size:125%;
        }

         .span_akt {
        list-style:none;
        font-size:125%;
        }

        .span_mat {
        list-style:none;
        font-size:125%;
        }

        

    </style>

</head>
    

<%call mobile_header %>


<div class="container">
    <div class="portlet">


        <div class="row">
            <div class="col-lg-12"><b>Afbrydelser</b></div>
        </div>

        <div class="row">
            <div class="col-lg-12">
                <%
                ddDato = year(now) &"/"& month(now) &"/"& day(now)
                stdatoSQL = dateadd("d", -7, ddDato)
                stdatoSQL = year(stdatoSQL) &"/"& month(stdatoSQL) &"/"& day(stdatoSQL)
                sldatoSQL = ddDato
                'call logindhistorik_week_60_100(session("mid"), 3, stdatoSQL, sldatoSQL) 
                %>


                <%

                    Dim timeId, akttype, aktivitetsTypeNavn, startTid, slutTid, timerReg
                    Redim timeId(10), akttype(10), aktivitetsTypeNavn(10), startTid(10), slutTid(10), timerReg(10)
                    strSQL = "SELECT tid, sttid, sltid, timer, tfaktim, at.aty_desc as akttypenavn FROM timer LEFT JOIN akt_typer as at ON (aty_id = tfaktim) WHERE tmnr = "& usemrn & strAktTyperKri &" AND tdato = '"& ddDato &"'"
                    oRec.open strSQL, oConn, 3
                    x = 0
                    while not oRec.EOF

                        timeId(x) = oRec("tid")
                        akttype(x) = "(" & oRec("tfaktim") & ")"

                        'if oRec("tfaktim") = 90 then
                        'aktivitetsTypeNavn(x) = "Kursus"
                        'else
                        'aktivitetsTypeNavn(x) = oRec("akttypenavn")
                        'end if

                        akttypenavn = ""
                        call akttyper(oRec("tfaktim"), 1)
                        aktivitetsTypeNavn(x) = akttypenavn
                        
                        startTid(x) = hour(oRec("sttid")) &":"& minute(oRec("sttid")) &":"& second(oRec("sttid")) 
                        slutTid(x) = hour(oRec("sltid")) &":"& minute(oRec("sltid")) &":"& second(oRec("sltid"))
                        timerReg(x) = oRec("timer")

                        'response.Write "<br> Akttype " & oRec("tfaktim")
                        if slutTid(x) <> "::" then
                            slutTid(x) = slutTid(x)
                        else
                            slutTid(x) = "-"
                        end if
                        

                    x = x + 1
                    oRec.movenext
                    wend
                    oRec.close

                    if x > 0 then
                        response.Write "<table style='width:100%;'><tr><td>Type</td><td <td style='text-align:center; padding-left:10px;'>Start</td><td <td style='text-align:center; padding-left:10px;'>Slut</td><td style='padding-left:10px;'>Ialt</td></tr>"

                        i = 0
                        for i = 0 TO UBOUND(timeId) 

                            if slutTid(i) <> "-" then
                            heleTimer = formatnumber(DateDiff("n", startTid(i), slutTid(i), 2, 2)/60, 2)
                            heleMinutterDB = formatnumber(DateDiff("n", startTid(i), slutTid(i), 2, 2), 0)

                            if cdbl(heleTimer) < 1 then
                                heleTimer = 0
                            end if

                            heleTimer_komma = instr(heleTimer, ",")
                            if heleTimer_komma <> 0 then
                            heleTimer = left(heleTimer, heleTimer_komma - 1)
                            end if

                            heleTimer_komma = instr(timerThis, ".")
                            if  heleTimer_komma <> 0 then
                            heleTimer = left(heleTimer, heleTimer_komma - 1)
                            end if

                            heleMinutter = 0
                            heleMinutter = formatnumber(heleMinutterDB/60, 2)
                            heleMinutter = right(heleMinutter, 2)
                            heleMinutter = formatnumber((heleMinutter*60)/100, 0)

                            if heleMinutter < 10 then
                                heleMinutter = "0"& heleMinutter
                            end if

                            if heleTimer < 10 then
                                heleTimer = "0"& heleTimer
                            end if
                    
                            else
                            heleTimer = "00"
                            heleMinutter = "00"
                            end if


                            strStartTid = Right("0" & hour(startTid(i)), 2) &":"& Right("0" & minute(startTid(i)), 2)
                            if slutTid(i) <> "-" then
                            strSlutTid = Right("0" & hour(slutTid(i)), 2) &":"& Right("0" & minute(slutTid(i)), 2)
                            else
                            strSlutTid = " - "
                            end if
        
                            if i < x then
                                response.Write "<tr>"
                        
                                response.Write "<td>"& aktivitetsTypeNavn(i) &" "&akttype(i)&"</td> <td style='text-align:center; padding-left:10px;'>"&strStartTid&"</td> <td style='text-align:center; padding-left:10px;'>"&strSlutTid&"</td> <td style='padding-left:10px;'>"& heleTimer &":"& heleMinutter &"</td>"

                                response.Write "</tr>"
                            end if
                        next

                        response.Write "</table>"
                    end if

                    if x = 0 then
                        response.Write "<table><tr><td>Ingen afbrydelser fundet</td></tr></table>"
                    end if

                %>

            </div>
        </div>


        <br />

         <%if lto = "mielexxxx" then %>

            <form action="checkin.asp?func=logud" method="post">
                <div class="row pad-b5">
                    <div class="col-lg-12" style="text-align:center;"><button type="submit" class="btn btn-danger" style="width:100%"><b style="font-size:150%;">Afslut dagen</b></button></div>
                </div>
            </form>

            <form action="checkin.asp?func=knap&FM_akttype=21" method="post">
                <div class="row pad-b5">
                    <div class="col-lg-12" style="text-align:center;"><button type="submit" class="btn btn-default" style="width:100%"><b style="font-size:150%;">Barns 1. Sygedag</b></button></div>
                </div>
            </form>

            <form action="checkin.asp?func=knap&FM_akttype=10" method="post">
                <div class="row pad-b5">
                    <div class="col-lg-12" style="text-align:center;"><button type="submit" class="btn btn-default" style="width:100%"><b style="font-size:150%;">Frokost</b></button></div>
                </div>
            </form>

            <form action="checkin.asp?func=knap&FM_akttype=90" method="post">
                <div class="row pad-b5">
                    <div class="col-lg-12" style="text-align:center;"><button type="submit" class="btn btn-default" style="width:100%"><b style="font-size:150%;">Kursus</b></button></div>
                </div>
            </form>

            <form action="checkin.asp?func=knap&FM_akttype=81" method="post">
                <div class="row pad-b5">
                    <div class="col-lg-12" style="text-align:center;"><button type="submit" class="btn btn-default" style="width:100%"><b style="font-size:150%;">Lægebesøg</b></button></div>
                </div>
            </form>

            <!-- Skal denne vare ferie uden lon -->
            <form action="checkin.asp?func=knap&FM_akttype=31" method="post">
                <div class="row pad-b5">
                    <div class="col-lg-12" style="text-align:center;"><button type="submit" class="btn btn-default" style="width:100%"><b style="font-size:150%;">Selvbetalt tid</b></button></div>
                </div>
            </form>

            <form action="checkin.asp?func=knap&FM_akttype=20" method="post">
                <div class="row pad-b5">
                    <div class="col-lg-12" style="text-align:center;"><button type="submit" class="btn btn-default" style="width:100%"><b style="font-size:150%;">Sygdom</b></button></div>
                </div>
            </form>

        <%end if %>

        <%'if lto = "outz" then %>


            <form action="checkin.asp?func=logud" method="post">
                <div class="row pad-b5">
                    <div class="col-lg-12" style="text-align:center;"><button type="submit" class="btn btn-danger" style="width:100%"><b style="font-size:150%;">Afslut dagen</b></button></div>
                </div>
            </form>

        <%
            strSQL = "SELECT aty_id, aty_desc FROM akt_typer WHERE aty_on = 1 AND aty_mobile_btn = 1 ORDER BY aty_mobile_btn_order, aty_desc"
            oRec.open strSQL, oConn, 3
            while not oRec.EOF

            call akttyper(oRec("aty_id"), 1)

        %>
                <form action="checkin.asp?func=knap&FM_akttype=<%=oRec("aty_id") %>" method="post">
                    <div class="row pad-b5">
                        <div class="col-lg-12" style="text-align:center;"><button type="submit" class="btn btn-default" style="width:100%"><b style="font-size:150%;"><%=akttypenavn %></b></button></div>
                    </div>
                </form>
        <%
            oRec.movenext
            wend
            oRec.close
        %>

        <%'end if %>

        <br />

        <div class="row">
            <div class="col-lg-12"><b>Komme - gå tider</b></div>
        </div>

        <div class="row">
            <div class="col-lg-5">
                <%
                    stdatoSQL = year(now) &"-"& month(now) &"-"& day(now)
                    call logindhistorik_week_60_100(usemrn, 2, stdatoSQL, stdatoSQL)
                %>
            </div>
        </div>



    </div>
</div>


           


<!--#include file="../inc/regular/footer_inc.asp"-->
