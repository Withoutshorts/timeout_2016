

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<%
if len(session("user")) = 0 then
	
	errortype = 5
	call showError(errortype)

    response.End

end if

call menu_2014 %>

<div class="wrapper">
    <div class="content">


     <%

        if len(session("user")) = 0 then
	    errortype = 5
	    call showError(errortype)
        response.End
	    end if
        

        if len(trim(request("FM_medarb"))) <> 0 then
            FM_medarb = request("FM_medarb")
        else
            FM_medarb = 1
        end if

        if len(trim(request("grp_tpID"))) <> 0 then
        grp_tpID = request("grp_tpID")
        else
        grp_tpID = 0
        end if

        select case FM_medarb
            case 1

                chbMedarb1 = "checked"
                SELstatus1 = ""
                chbMedarb2 = ""
                SELstatus2 = "DISABLED"            
                if grp_tpID <> 0 then
                grp_typIDSQL = " WHERE medarbejdertype = "& grp_tpID & " AND (mansat = 1)"
                else
                grp_typIDSQL = " WHERE (mansat = 1)"
                end if

            case 2

                chbMedarb1 = ""
                SELstatus1 = "DISABLED"
                chbMedarb2 = "checked"
                SELstatus2 = ""
                if grp_tpID <> 0 then
                grp_typIDSQL = " WHERE pr.ProjektgruppeId = "& grp_tpID & " AND (mansat = 1)"
                else
                grp_typIDSQL = " WHERE (mansat = 1)"
                end if

        end select


        if len(trim(request("ferieaar"))) then
            ferieaar = request("ferieaar")
        else
            nowYear = year(now)

            nowDate = year(now)&"-"&month(now) &"-"& day(now)
            ferieNextYearStart = year(now) &"-5-1"
            
            if (Datevalue(nowDate) < DateValue(ferieNextYearStart)) then

            feriestartDato = nowYear
            feriesslutDato = nowYear + 1
            else
            feriestartDato = (year(nowDate) + 1)
            feriesslutDato = nowYear + 2
            end if

            ferieaar = feriestartDato

         end if


         ferieStarArr = ferieaar
         ferieSlutArr = ferieaar + 1 
        

         optjentferieStartDato = (ferieaar -1) &"-01-01"
         optjentferieSluttDato = (ferieaar -1) &"-12-31"

         nuFerieSaldoStartDato = (ferieaar -1) &"-05-01"
         nuFerieSaldoSlutDato = ferieaar &"-04-30"

         saldoFraDato = ferieStarArr & "-05-01"
         saldoTilDato = ferieSlutArr & "-04-30"

         optjentferieStartDatoDDMMYYY = day(optjentferieStartDato) &"/"& month(optjentferieStartDato) &"/"& year(optjentferieStartDato)


         public ferieBalVal
         function getFerieSaldo(mid, datofra, datotil)

            ferieOptjtimer = 0
            ferieOptjOverforttimer = 0
            ferieUdbTimer = 0
            ferieAFTimer = 0
            ferieOptjUlontimer = 0
            ferieKorrigering = 0
            ferieAFulonTimer = 0

            'response.Write "medidmedidmedid " & mid 
            strSQLFerieOptjent = "SELECT tfaktim, sum(timer) as sumtimer FROM timer WHERE Tmnr = "& mid &" AND tdato between '" & datofra & "' AND '" & datotil & "'"_
            &" AND (tfaktim = 15 OR tfaktim = 111 OR tfaktim = 16 OR tfaktim = 14 OR tfaktim = 112 OR tfaktim = 19 OR tfaktim = 126) "_
            &" GROUP BY tfaktim"
            'response.Write  "<br>" & strSQLFerieOptjent
            'response.flush
            oRec2.open strSQLFerieOptjent, oConn, 3
            while not oRec2.EOF
                                            
            select case oRec2("tfaktim")
            case 15
            ferieOptjtimer = ferieOptjtimer + oRec2("sumtimer")
            case 111
            ferieOptjOverforttimer = ferieOptjOverforttimer + oRec2("sumtimer")
            case 16
            ferieUdbTimer = ferieUdbTimer + oRec2("sumtimer")
            case 14
            ferieAFTimer = ferieAFTimer + oRec2("sumtimer")
            case 112
            ferieOptjUlontimer = ferieOptjUlontimer + oRec2("sumtimer")
            case 19
            ferieAFulonTimer = ferieAFulonTimer + oRec2("sumtimer")
            case 126
            ferieKorrigering = ferieKorrigering + oRec2("sumtimer")

            end select
                                          
            oRec2.movenext
            wend
            oRec2.close


            'response.Write "<br><br> herFO " & ferieOptjtimer
            'response.Write "<br><br> her" & ferieOptjOverforttimer(x)
            'response.Write "<br><br> her" & ferieUdbTimer(x)
            'response.Write "<br><br> her" & ferieAFTimer(x)
            'response.Write "<br><br> her" & ferieOptjUlontimer(x)
            'response.Write "<br><br> her" & ferieAFulonTimer(x)


            ferieBal = 0
            call ferieBal_fn(ferieOptjtimer, ferieOptjOverforttimer, ferieOptjUlontimer, ferieAFTimer, ferieAFulonTimer, ferieUdbTimer)
	 


            ' Sker efter kaldelse af funktionen
	        ' if normTimerDag(x) <> 0 then 
	        'ferieBalVal = ferieBal/normTimerDag(x)
	        ' else
	        ' ferieBalVal = 0
	        ' end if
	 
            'ferieBalVal = ferieBal

            'Response.write "HER: " & ferieBal


            'L�gger feriekorrigering til saldeon
            ferieBalVal = ferieBal + ferieKorrigering

            if ferieBalVal <> 0 then 
	        ferieBalValTxt = formatnumber(ferieBalVal,2)
            ferieBalValTxtExp = ferieBalValTxt
	        else
	        ferieBalValTxt = ""
            ferieBalValTxtExp = 0
	        end if

         end function




         function getFerieFri (mid, datofra, datotil)


            fefriTimer = 0
            fefriplTimer = 0
            fefriTimerBr = 0
            fefriTimerUdb = 0

            strSQLFeFri = "SELECT tfaktim, sum(timer) as sumtimer FROM timer WHERE Tmnr = "& mid &" AND tdato between '" & datofra & "' AND '" & datotil & "' AND (tfaktim = 12 OR tfaktim = 18 OR tfaktim = 13 OR tfaktim = 17) GROUP BY tfaktim"
            'response.Write strSQLFeFri 
            oRec2.open strSQLFeFri, oConn, 3
            while not oRec2.EOF

            select case oRec2("tfaktim")
            case 12
            fefriTimer = fefriTimer + oRec2("sumtimer")
            case 18
            fefriplTimer = fefriplTimer + oRec2("sumtimer")
            case 13
            fefriTimerBr = fefriTimerBr + oRec2("sumtimer")
            case 17
            fefriTimerUdb = fefriTimerUdb + oRec2("sumtimer")

            end select

            oRec2.movenext
            wend
            oRec2.close

            
         
	  

	        fefriBal = 0
            fefriBal  = (fefriTimer - (fefriTimerBr + fefriTimerUdb))
	       
        

         end function


         
         func = request("func")

         select case func
         
         case "opdater"

            
             medarbid = request("medarbid")
             medarbids = split(medarbid, ", ")

             ferieoptentIAlt = ferieoptjent + ferieKorri

               ferieoptjent = 0
                ferieKorri = 0
                ferieoverfort = 0
                saerligFerie = 0
                godkendMedarb = 0
                ferieoverfortulon = 0

            for t = 0 TO UBOUND(medarbids)
                

                ferieoptjent = request("ferieoptjent_" & medarbids(t))
                ferieKorri = request("ferieKorri_" & medarbids(t))
                ferieoverfort = request("ferieoverfort_" & medarbids(t))
                saerligFerie = request("saerligFerie_" & medarbids(t))
                godkendMedarb = request("godkend_" & medarbids(t))
                ferieoverfortulon = request("ferieoverfortulon_" & medarbids(t))

                if godkendMedarb = 1 then
                    godkendMedarb = godkendMedarb
                else
                    godkendMedarb = 0
                end if

                antaldagemedTimer = request("antalDageMtimer_" & medarbids(t))
                medarbnorm = request("medarbnorm_" & medarbids(t))

                gennemsnitdagstimer = cdbl(medarbnorm/antaldagemedTimer)

                '' Laver felter om til dage i stedet for timer
                ferieoptjent = ferieoptjent * gennemsnitdagstimer
                ferieKorri = ferieKorri * gennemsnitdagstimer
                ferieoverfort = ferieoverfort * gennemsnitdagstimer
                saerligFerie = saerligFerie * gennemsnitdagstimer
                ferieoverfortulon = ferieoverfortulon * gennemsnitdagstimer

                
               if cint(godkendMedarb) = 1 then

            
                    'NULSTILLER 
                    'Overskriv gamle med nye tal, dvs. slet de gamle f�rst
                    oConn.execute("DELETE FROM timer WHERE tmnr = "& medarbids(t) &" AND (tfaktim = 15 OR tfaktim = 111 OR tfaktim = 112 OR tfaktim = 126 OR tfaktim = 12) AND tdato BETWEEN '"& saldoFraDato &"' AND '"& saldoTilDato &"'")


                    ' ferie optjent
                    ferieoptjent = (ferieoptjent*1 + (cdbl(ferieKorri)*1))

                    if cdbl(ferieoptjent) > 0 then
                        call indlasFerieSaldo(lto, medarbids(t), cdbl(ferieoptjent), 15, ferieStarArr)
                    end if
               
                    'ferie overf�rt
                    if cdbl(ferieoverfort) >= 0 then
                        call indlasFerieSaldo(lto, medarbids(t), cdbl(ferieoverfort), 111, ferieStarArr)
                    end if

                    'ferie overf�rt ulon
                    if cdbl(ferieoverfortulon) >= 0 then
                        call indlasFerieSaldo(lto, medarbids(t), cdbl(ferieoverfortulon), 112, ferieStarArr)
                    end if
         
                    'ferie optjent uden l�n - ferie korrigering (kab b�de v�re minus og plus)
                    'KUN TIL AT HOLDE �JE MED korrektion, ved rediger p� ferietildeling
                    if cdbl(ferieKorri) > 0 then
                        call indlasFerieSaldo(lto, medarbids(t), cdbl(ferieKorri), 126, ferieStarArr) 
                    end if
  
                  

                    'S�rlig ferie optjent
                    if cdbl(saerligFerie) >= 0 then              
                        call indlasFerieSaldo(lto, medarbids(t), cdbl(saerligFerie), 12, ferieStarArr)
                    end if

                end if

                
            next

           
            response.Redirect ("ferietildel.asp?ferieaar="&ferieaar&"&FM_medarb="&FM_medarb&"&grp_tpID="&grp_tpID)


         case else

     %>


        <script src="js/ferietildel_jav.js" type="text/javascript"></script>   

        <div class="container" style="width:1500px;">
            <div class="portlet">
                <h3 class="portlet-title"><u>Ferie tildeling </u></h3>

                <div class="portlet-body">

                    <form action="ferietildel.asp?" method="post" >
                    <div class="well">
                        <div class="row">
                            <div class="col-lg-12">
                                <h4 class="panel-title-well">S�gefilter</h4>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1">Ferie �r:</div>
                            <div class="col-lg-1">&nbsp;</div>
                            <div class="col-lg-3">Medarbejdertype <input type="radio" name="FM_medarb" value="1" <%=chbMedarb1 %> /></div>
                            <div class="col-lg-3">Projektgrupper <input type="radio" name="FM_medarb" value="2" <%=chbMedarb2 %> /></div>                            
                        </div>
                        <div class="row">

                            <% 
                                nowYear = year(now)

                                'forvalgYear1 = nowYear + 1
                                'forvalgYear2 = nowYear + 2

                                nowDate = year(now)&"-"&month(now) &"-"& day(now)
                                ferieNextYearStart = year(now) &"-5-1"

                                'response.Write ferieStartDato  &"<br>"
                                'response.Write nowDate  &"<br>"
                                if (Datevalue(nowDate) < DateValue(ferieNextYearStart)) then
                                'response.Write "dette �r skal starte som ferie �r"
                                feriestartDato = nowYear
                                feriesslutDato = nowYear + 1
                                else
                                feriestartDato = (year(nowDate) + 1)
                                feriesslutDato = nowYear + 2
                                end if

                                'response.Write nowYear &"<br>"& feriestartDato &"<br>"& feriesslutDato
                            %>
                            <div class="col-lg-2">
                                <select class="form-control input-small" name="ferieaar">

                                    <%
                                    i = 0
                                    while i < 10 
                                        
                                    if cint(feriestartDato) = cint(ferieaar) then
                                        ferieSEL = "selected"
                                    else
                                        ferieSEL = ""
                                    end if                                      
                                    %>
                                    
                                    <option value="<%=feriestartDato %>" <%=ferieSEL %>><%=feriestartDato%> / <%=feriesslutDato %></option>

                                    <%
                                    feriestartDato = feriestartDato + 1
                                    feriesslutDato = feriesslutDato + 1

                                    i = i + 1

                                    wend
                                    %>

                                </select>
                            </div>


                            
                            <div class="col-lg-3">
                                <select class="form-control input-small FM_medarbtype_SEL" name="grp_tpID" <%=SELstatus1 %> onchange="submit()">
                                    <option value="0">Alle</option>

                                    <%
                                        'if FM_medarb = 1 then

                                        strSQL = "SELECT id, type FROM medarbejdertyper ORDER BY type"
                                        oRec.open strSQL, oConn, 3
                                        while not oRec.EOF
                                        
                                        if cint(grp_tpID) = cint(oRec("id")) then
                                        SELOP = "SELECTED"
                                        else
                                        SELOP = ""
                                        end if

                                        %><option value="<%=oRec("id") %>" <%=SELOP %>><%=oRec("type") %></option><%

                                        oRec.movenext
                                        wend
                                        oRec.close
                                        

                                        'end if
                                    %>

                                </select>
                            </div>
                            <div class="col-lg-3">
                                <select class="form-control input-small FM_progrb_SEL" name="grp_tpID" <%=SELstatus2 %> onchange="submit()">

                                    <%
                                        'if FM_medarb = 2 then

                                            strSQL = "SELECT id, navn FROM projektgrupper WHERE id <> 1 ORDER BY navn"
                                            oRec.open strSQL, oConn, 3
                                            while not oRec.EOF
                                        
                                            if cint(grp_tpID) = cint(oRec("id")) then
                                            SELOP = "SELECTED"
                                            else
                                            SELOP = ""
                                            end if

                                            %><option value="<%=oRec("id") %>" <%=SELOP %>><%=oRec("navn") %></option><%

                                            oRec.movenext
                                            wend
                                            oRec.close

                                        'end if
                                    %>

                                </select>
                            </div>

                            <div class="col-lg-1"><button type="submit" class="btn btn-secondary btn-sm pull-left" style="width:80px;"><b>Vis</b></button></div>
                            
                        </div>
                       
                    </div>
                    </form>

                    <form action="ferietildel.asp?func=opdater&ferieaar=<%=ferieaar %>&FM_medarb=<%=FM_medarb %>&grp_tpID=<%=grp_tpID %>" method="post">

                        <div class="row">
                            <div class="col-lg-9">&nbsp</div>
                            <div class="col-lg-3 pad-b10"><button type="submit" class="btn btn-success btn-sm pull-right" style="width:80px;"><b><%=medarb_txt_020 %></b></button></div>
                        </div>

                    <style>
                        table tr td {
                            text-align:center;
                        }
                    </style>

                    <%select case lto
                    case "esn"
                        FeriefridageTxt = "S�r. feriedage"
                    case else
                        FeriefridageTxt = "Feriefridage"
                    end select%>

                    <table id="xscrollable" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th>Navn</th>
                                <th>Ansat d.</th>
                                <th>Norm</th>
                                <th>Ferie saldo <br /><span style="font-size:9px;">Nuv�rende</span></th>
                                <!--<th>Medarbejdertype</th>-->
                                <th>Optjent ferie <br /><span style="font-size:9px;">1.1.<%=((ferieaar)-1) %> - 31.12.<%=((ferieaar)-1) %></span></th>
                                <th>Korrigering <br /><span style="font-size:9px;">+/- dage</span> <div style="padding-top:5px;"><input type="text" id="feriekorri_multi_val" style="width:60px;" class="form-control input-small"/></div></th>
                                <th>Optjent u. l�n <br /> <span style="font-size:9px;">dage</span><div style="padding-top:5px;"><input type="text" id="ferieulon_multi_val" style="width:60px;" class="form-control input-small"/></div></th>
                                <th>Overf�rt Ekstra<br /> <span style="font-size:9px;">Evt. rest-ferie</span> </th>
                                <th>Ferie tildelt ialt <br /><span style="font-size:9px;"><%=ferieaar %></span></th>
                                <!--<th><%=FeriefridageTxt %> opt. <br /><span style="font-size:9px;">1.1.<%=((ferieaar)-1) %> - 31.12.<%=((ferieaar)-1) %></span></th>-->
                                <th><%=FeriefridageTxt %> tildelt<br /><span style="font-size:9px;"><%=ferieaar %></span><br /> <div style="padding-top:5px;"><input type="text" id="feriefriOpt_multi_val" style="width:60px;" class="form-control input-small"/></div></th>
                                <!--<th><%=FeriefridageTxt %> saldo <br /><span style="font-size:9px;"><%=ferieaar %></span></th>-->
                                <th style="text-align:center">Godkend <br /> <input type="checkbox" id="godkendalle" /></th>
                            </tr>
                        </thead>

                        <tbody>

                            <%
                                'ferieafholdt = 0
                                select case FM_medarb
                                case 1
                                    strSQL = "SELECT Mid, Mnavn, ansatdato, init, Medarbejdertype FROM medarbejdere "
                                case 2
                                    strSQL = "SELECT Mid, Mnavn, ansatdato, init, Medarbejdertype FROM medarbejdere LEFT JOIN progrupperelationer pr ON (pr.MedarbejderId = Mid) "
                                end select

                                strSQL = strSQL & grp_typIDSQL
                                strSQL = strSQL & " ORDER BY mnavn" 
                                
                                'response.Write strSQL
                                'response.Flush

                                oRec.open strSQL, oConn, 3 
                                while not oRec.EOF

                                %>
                                    <tr>
                                        <td style="text-align:left"><%=oRec("Mnavn") %>
                                            <% if (len(trim(oRec("init"))) <> 0) then%>
                                                &nbsp;[<%=oRec("init") %>]
                                            <%end if %>

                                            <input type="hidden" name="medarbid" value="<%=oRec("Mid") %>" />
                                        </td>

                                        <td style="white-space:nowrap;"><%=oRec("ansatdato") %></td>

                                        <td>
                                            <%

                                                '*** Finder medarbejdertyper_historik ***'

                                                strSQLmth = "SELECT mid, mtype, mtypedato FROM medarbejdertyper_historik WHERE "_
                                                &" mid = "& oRec("Mid") &" AND mtypedato > '"& optjentferieStartDato &"' ORDER BY mtypedato, id"

                                                t = 0

                                                oRec2.open strSQLmth, oConn, 3

                                                while not oRec2.EOF

                                                mtyperIntvTyp = oRec2("mtype")

                                                mtyperIntvDato = oRec2("mtypedato")

                                                 t = t + 1

                                                oRec2.movenext
                                                wend
                                                oRec2.close


                                                if t = 0 then
                                                    medarbtypeDato = year(oRec("ansatdato")) &"-"& month(oRec("ansatdato")) &"-"& day(oRec("ansatdato"))
                                                else
                                                    medarbtypeDato = year(mtyperIntvDato) &"-"& month(mtyperIntvDato) &"-"& day(mtyperIntvDato)
                                                end if

                                                call normtimerPer(oRec("mid"), medarbtypeDato, 6, 0)

                                                response.Write formatnumber(ntimper,2)

                                                
                                                'if year(medarbtypeDato) = nowYear then
                                                    %><!--&nbsp;<span style="color:red">!</span>--><%
                                                'end if


                                                normTimerDag = 0
                                                normTimerDag = ntimper / antalDageMtimer
                                                'response.Write "<br>" &  medarbtypeDato & "mid" & oRec("mid") & " mtype " & oRec("Medarbejdertype")

                                            %>
                                            <input type="hidden" name="medarbnorm_<%=oRec("mid") %>" value="<%=ntimper %>" />
                                            <input type="hidden" name="antalDageMtimer_<%=oRec("mid") %>" value="<%=antalDageMtimer %>" />
                                        </td>

                                        <!-- Ferie Saldo Nuv�rende -->
                                        <td>
                                            
                                        <%
                                         
                                              
                                            '15, 111, 16, 14, 112, 19: disse skal hentes fra databasen, det er tfaktim, ogs� skal de bruges i feriebal_fn
                                            ' 15 = ferieoptjent
                                            ' 111 = ferieoverf�rt
                                            ' 16 = ferieudbetalt
                                            ' 14 = ferieafhold
                                            ' 112 = optjent uden l�n
                                            ' 19 = ferie afholdt uden l�n

                                            call getFerieSaldo(oRec("Mid"),nuFerieSaldoStartDato,nuFerieSaldoSlutDato)

                                            if normTimerDag <> 0 then
                                            ferieBalVal = ferieBalVal / normTimerDag
                                            else
                                            ferieBalVal = 0
                                            end if

                                        %>

                                            <%=formatnumber(ferieBalVal,2) %>
                                        </td>
                                     
                                            <%

                                                'totalsDaysInYear = DateDiff("d", optjentferieStartDato, optjentferieSluttDato)
                                                'totalsDaysInYear = totalsDaysInYear + 1

                                                'daysPerMonth = totalsDaysInYear / 12 
                                                'daysPerMonth = cdbl(daysPerMonth)
                                                'response.Write "Days in year " & totalsDaysInYear & "Timer pr. month" & daysPerMonth & "<br>"


                                                'daysPerMonth = 12
                                                monthDiff = 1
                                               
                                                'response.Write "<br> Records " & totalRecords
                                                totalRecords = 0
                                                ferieoptjent = 0 
                                                ferieoptjentTot = 0

                                                first_optjentferieSlutDatoMtypetjk = optjentferieStartDatoDDMMYYY
                                              
                                              
                                                
                                                strSQLmth = "SELECT mid, mtype, mtypedato, mt.feriesats, mt.type AS mtypenavn FROM medarbejdertyper_historik mth LEFT JOIN medarbejdertyper as mt ON (mt.id = mtype) WHERE "_                                           
                                                &" mid = "& oRec("Mid") &" GROUP BY mtypedato ORDER BY mtypedato DESC, mth.id DESC"
                                              
                                                '&" AND mtypedato > '"& optjentferieStartDato &"'
                                                '&" mid = "& oRec("Mid") &" ORDER BY mtypedato, id"
                                                'response.Write "<br><br>"& strSQLmth & "<br>"
                                                'AND '"& optjentferieSluttDato &"'

                                                oRec2.open strSQLmth, oConn, 3
                                                t = 0
                                                stopLoopDatoererForGamle = 0
                                                while not oRec2.EOF
                                                
                                                
                                                if cint(stopLoopDatoererForGamle) = 0 then


                                                'AND mtypedato BETWEEN '"& optjentferieStartDato &"' AND '"& optjentferieSluttDato &"'

                                                ferieoptjent = 0

                                               

                                                 if cint(t) = 0 then
                                                    
                                                    if year(cDate(oRec("ansatdato"))) < optjentferieStartDato then
                                                    optjentferieSlutDatoMtypetjk = optjentferieSluttDato
                                                    else
                                                    optjentferieSlutDatoMtypetjk = oRec("ansatdato")
                                                    end if

                                                    first_optjentferieSlutDatoMtypetjk = optjentferieSlutDatoMtypetjk

                                                 else
                                                    optjentferieSlutDatoMtypetjk = dateAdd("d", 1, lastmtypedato) 
                                                 end if
                                                

                                                visMtype = 1
                                                if cDate(oRec2("mtypedato")) < cDate(optjentferieStartDatoDDMMYYY) then
                                                mtypedato = optjentferieStartDatoDDMMYYY 'year(optjentferieStartDatoDDMMYYY)&"-"&month(oRec2("mtypedato"))&"-"&day(oRec2("mtypedato"))

                                                stopLoopDatoererForGamle = 1
                                                else
                                                mtypedato = day(oRec2("mtypedato"))&"-"&month(oRec2("mtypedato"))&"-"&year(oRec2("mtypedato"))
                                                
                                                end if

                                                


                                                nuvarendeferiesats = oRec2("feriesats")

                                                dDiff = DateDiff("d", mtypedato, optjentferieSlutDatoMtypetjk) 
                                                monthDiff = DateDiff("m", mtypedato, optjentferieSlutDatoMtypetjk) + 1
                                               

                                                if monthDiff > 0 then
                                                monthDiff = monthDiff
                                                else
                                                monthDiff = 1
                                                end if

                                                'if (dDiff/monthDiff) > 3 then
                                                'monthDiff = monthDiff + 1
                                                'else
                                                'monthDiff = monthDiff
                                                'end if


                                                ferieoptjent = (nuvarendeferiesats * monthDiff) 
                                                
                                                'if totalRecords = 0 then
                                                '    if DateValue(mtypedato) > DateValue(year(now)&"-"&month(now)&"-"&day(now)) then
                                                '        ferieoptjent = ferieoptjent
                                                '    else
                                                '        if DateValue(year(now)&"-"&month(now)&"-"&day(now)) > DateValue(optjentferieSluttDato) then
                                                '            ferieoptjent = ferieoptjent + oRec2("feriesats") * (DateDiff("d", mtypedato, optjentferieSluttDato)+1)/daysPerMonth
                                                '        else
                                                '            ferieoptjent = ferieoptjent + oRec2("feriesats") * (DateDiff("d", mtypedato, year(now)&"-"&month(now)&"-"&day(now))+1)/daysPerMonth
                                                '        end if
                                                '    end if
                                                'end if

                                                'if t < totalRecords and t > 0 then
                                                '    if DateValue(year(now)&"-"&month(now)&"-"&day(now)) > DateValue(mtypedato) then
                                                '        ferieoptjent = ferieoptjent + lastferiesats * (DateDiff("d", lastmtypedato, mtypedato)+1)/daysPerMonth
                                                '        'response.Write lastmtypedato &" - "& mtypedato & "sats" & lastferiesats
                                                '    else
                                                '        ferieoptjent = ferieoptjent + lastferiesats * (DateDiff("d", lastmtypedato, year(now)&"-"&month(now)&"-"&day(now))+1)/daysPerMonth
                                                '    end if
                                                'end if

                                                'if t = (totalRecords - 1) then
                                                '    if DateValue(year(now)&"-"&month(now)&"-"&day(now)) > DateValue(optjentferieSluttDato) then
                                                '        ferieoptjent = ferieoptjent + oRec2("feriesats") * (DateDiff("d", mtypedato, optjentferieSluttDato)+1)/daysPerMonth
                                                '        'response.Write " <br><br>"& mtypedato &" - "& optjentferieSluttDato & "sats " & oRec2("feriesats")
                                                '    else
                                                '        ferieoptjent = ferieoptjent + oRec2("feriesats") * (DateDiff("d", mtypedato, year(now)&"-"&month(now)&"-"&day(now))+1)/daysPerMonth
                                                '    end if
                                                'end if

                                                ferieoptjentTot = ferieoptjentTot*1 + cdbl(ferieoptjent)

                                                'response.Write "t = " & t &" optjentferieSlutDatoMtypetjk: "& optjentferieSlutDatoMtypetjk &" mtypedato = "& mtypedato & " feriesats = " & oRec2("feriesats") & " ferieoptjent = " & ferieoptjent & " ferieoptjentTot: " & ferieoptjentTot &"<br><br>" 
                                                
                                                
                                                'Response.write oRec2("mtypenavn") & " "& oRec2("mtypedato") &" sats: "& nuvarendeferiesats &"<br>" ' = "& formatnumber(ferieoptjent, 0) & "<br>"
                                                'Response.write oRec2("mtypenavn") & " sats: "& nuvarendeferiesats &"<br>"

                                                'if t > 0 then
                                                'response.Write "<br>" & DateDiff("m",lastmtypedato,mtypedato)
                                                'end if

                                                'response.Write "<br>" & oRec2("mtypedato")

                                                lastmtypedato = year(oRec2("mtypedato"))&"-"&month(oRec2("mtypedato"))&"-"&day(oRec2("mtypedato"))
                                                lastferiesats = oRec2("feriesats")

                                                t = t + 1

                                                end if 'stopLoopDatoererForGamle = 1

                                                oRec2.movenext
                                                wend
                                                oRec2.close


                                            

                                                totalRecords = t
                                                ferieoptjent = ferieoptjentTot

                                              
                                                if cint(totalRecords) = 0 then 'Ingen medarbejderhistorik i ferie�r

                                                forDato = lastmtypedato
                                                nuvarendeferiesats = 0
                                                strSQLForFerieArr = "SELECT Medarbejdertype, feriesats, type AS mtypenavn FROM medarbejdere LEFT JOIN medarbejdertyper as mt ON (mt.id = Medarbejdertype) WHERE mid ="& oRec("mid") 
                                                'response.Write strSQLForFerieArr 
                                                'response.flush
                                                oRec2.open strSQLForFerieArr, oConn,3 
                                                while not oRec2.EOF
                                                'forDato = oRec2("mtypedato")
                                                nuvarendeferiesats = oRec2("feriesats")
                                                mtypenavn = oRec2("mtypenavn")
                                              
                                                oRec2.movenext
                                                wend
                                                oRec2.close

                                               
                                                'Response.write mtypenavn &" sats: "& nuvarendeferiesats
                                               
                                               

                                                'mtypedato = year(oRec2("mtypedato"))&"-"&month(oRec2("mtypedato"))&"-"&day(oRec2("mtypedato"))
                                                'nuvarendeferiesats = oRec2("feriesats")

                                                    if year(cDate(oRec("ansatdato"))) < optjentferieStartDato then
                                                    optjentferieSlutDatoMtypetjk = optjentferieStartDato
                                                    else
                                                    optjentferieSlutDatoMtypetjk = oRec("ansatdato")
                                                    end if

                                                monthDiff = DateDiff("m", optjentferieSlutDatoMtypetjk, optjentferieSluttDato) + 1
                                                'dDif = DateDiff("d", optjentferieStartDato, optjentferieSluttDato) + 1
                                                'Response.write "nuvarendeferiesats: "& nuvarendeferiesats & " dDif: "& dDif  & " monthDiff: " & monthDiff

                                                    if cDate(oRec("ansatdato")) < cDate(optjentferieSluttDato) then
                                                    ferieoptjent = (nuvarendeferiesats * monthDiff)
                                                    else
                                                    ferieoptjent = 0
                                                    end if
                                                
                                                else

                                                    if cDate(oRec("ansatdato")) < cDate(optjentferieSluttDato) then
                                                    ferieoptjent = ferieoptjent
                                                    else
                                                    ferieoptjent = 0
                                                    end if

                                                end if

                                                'if cint(totalRecords) = 0 then
                                                   
                                                'else

                                                'end if
                                                'response.Write "before date " & forDato & "Before sats " & nuvarendeferiesats &"<br>" 


                                               'if t = 0 AND forDato <> "" then
                                               '     'response.Write "<br>t = 0 "& nuvarendeferiesats &" dato "& forDato& "<br>"
                                               '     if DateValue(year(now)&"-"&month(now)&"-"&day(now)) > DateValue(optjentferieSluttDato) then
                                               '         ferieoptjent = ferieoptjent + nuvarendeferiesats * (DateDiff("d", optjentferieStartDato, optjentferieSluttDato)+1)/daysPerMonth
                                               '     else
                                               '         ferieoptjent = ferieoptjent + nuvarendeferiesats * (DateDiff("d", optjentferieStartDato, year(now)&"-"&month(now)&"-"&day(now))+1)/daysPerMonth
                                               '     end if
                                               'end if

                                               'if t = 0 AND forDato = "" then
                                               '     ferieoptjent = 0
                                               'end if
                                                   
                                                if ferieoptjent > 25 then
                                                 ferieoptjent = 25
                                                end if

                                                if ferieoptjent > 0 then
                                                ferieoptjent = formatnumber(ferieoptjent, 0)
                                                else
                                                ferieoptjent = 0
                                                end if



                                            %>
                                      
                                        <td>
                                            <input type="hidden" name="ferieoptjent_<%=oRec("Mid") %>" class="form-control input-small" style="width:60px;" value="<%=ferieoptjent %>" />
                                            <%=formatnumber(ferieoptjent,0) & ",00" %>
                                        </td>

                                        <td>
                                            <%
                                                
                                                feriekorrigering = 0
                                                strSQLFerieKori = "SELECT sum(timer) as sumtimer FROM timer WHERE Tmnr = "& oRec("Mid") &" AND tfaktim = 126 AND tdato BETWEEN '"& saldoFraDato &"' AND '"& saldoTilDato &"' GROUP BY tfaktim"
                                                'response.Write strSQLFerieKori
                                                oRec2.open strSQLFerieKori, oConn, 3
                                                if not oRec2.EOF then
                                                feriekorrigering = oRec2("sumtimer")
                                                end if
                                                oRec2.close

                                                if normTimerDag <> 0 then
                                                feriekorrigering = feriekorrigering / normTimerDag
                                                else
                                                feriekorrigering = 0
                                                end if

                                                if feriekorrigering <> 0 then
                                                feriekorrigering = feriekorrigering
                                                else
                                                feriekorrigering = 0
                                                end if

                                            %>
                                            <input type="text" name="ferieKorri_<%=oRec("Mid") %>" id="korrigering_opt_<%=oRec("Mid") %>" style="width:60px;" class="form-control input-small korrigering_input" value="<%=formatnumber(feriekorrigering, 2) %>" />

                                        
                                             <td>
                                            <%
                                                
                                                ferieulon = 0
                                                strSQLFerieulon = "SELECT sum(timer) as sumtimer FROM timer WHERE Tmnr = "& oRec("Mid") &" AND tfaktim = 112 AND tdato BETWEEN '"& saldoFraDato &"' AND '"& saldoTilDato &"' GROUP BY tfaktim"
                                                'response.Write strSQLFerieulon
                                                oRec2.open strSQLFerieulon, oConn, 3
                                                if not oRec2.EOF then
                                                ferieulon = oRec2("sumtimer")
                                                end if
                                                oRec2.close

                                                if normTimerDag <> 0 then
                                                ferieulon = ferieulon / normTimerDag
                                                else
                                                ferieulon = 0
                                                end if

                                                if ferieulon <> 0 then
                                                ferieulon = ferieulon
                                                else
                                                ferieulon = 0
                                                end if

                                            %>
                                            <input type="text" name="ferieoverfortulon_<%=oRec("Mid") %>" id="ferieoverfortulon_<%=oRec("Mid") %>" class="form-control input-small ulon_input" style="width:60px;" value="<%=formatnumber(ferieulon, 2) %>" />
                                                 </td>

                                        <td>

                                            <%
                                               
                                                strSQLFerieOverfort = "SELECT sum(timer) as sumtimer FROM timer WHERE Tmnr = "& oRec("Mid") &" AND tfaktim = 111 AND tdato BETWEEN '"& saldoFraDato &"' AND '"& saldoTilDato &"' GROUP BY tfaktim"
                                                'response.Write strSQLFerieOverfort
                                                oRec2.open strSQLFerieOverfort, oConn, 3
                                                if not oRec2.EOF then
                                                ferieoverfort = oRec2("sumtimer")
                                                end if
                                                oRec2.close

                                                if normTimerDag <> 0 then
                                                ferieoverfort = ferieoverfort / normTimerDag
                                                else
                                                ferieoverfort = 0
                                                end if

                                               if ferieoverfort <> 0 then
                                               ferieoverfort = ferieoverfort
                                               else
                                               ferieoverfort = 0
                                               end if


                                            %>
                                            <input name="ferieoverfort_<%=oRec("Mid") %>" class="form-control input-small" style="width:60px;" value="<%=formatnumber(ferieoverfort, 2) %>" />

                                        </td>

                                        <td>
                                           
                                            <%  '**FERIE TILDELT KOMMENDE FERIE�R

                                       
                                               
                                                ferietildeltnytferieaar = 0
                                                strSQLFerieOverfort = "SELECT sum(timer) as sumtimer FROM timer WHERE Tmnr = "& oRec("Mid") &" AND (tfaktim = 15 OR tfaktim = 111 OR tfaktim = 112) AND tdato BETWEEN '"& saldoFraDato &"' AND '"& saldoTilDato &"' GROUP BY tfaktim"
                                                'response.Write strSQLFerieOverfort
                                                oRec2.open strSQLFerieOverfort, oConn, 3
                                                while not oRec2.EOF 
                                                ferietildeltnytferieaar = ferietildeltnytferieaar + oRec2("sumtimer")
                                                oRec2.movenext
                                                wend
                                                oRec2.close

                                                if normTimerDag <> 0 then
                                                ferietildeltnytferieaar = ferietildeltnytferieaar / normTimerDag
                                                else
                                                ferietildeltnytferieaar = 0
                                                end if

                                               if ferietildeltnytferieaar <> 0 then
                                               ferietildeltnytferieaar = ferietildeltnytferieaar
                                               else
                                               ferietildeltnytferieaar = 0
                                               end if


                                            %>
                                            <b><%=formatnumber(ferietildeltnytferieaar, 2) %></b>
                                        </td>

                                        <!--
                                        <td>
                                            <%
                                                call getFerieFri(oRec("mid"), optjentferieStartDato, optjentferieSluttDato)
                                                if normTimerDag <> 0 then
                                                fefriBal = fefriBal / normTimerDag
                                                else
                                                fefriBal = 0
                                                end if

                                            %>


                                           
                                            <%=formatnumber(fefriBal,2) %>
                                        </td>
                                        -->

                                        <td>
                                            <% 

                                           
                                            strSQLSaerligFerie = "SELECT sum(timer) as sumtimer FROM timer WHERE Tmnr = "& oRec("Mid") &" AND tfaktim = 12 AND tdato BETWEEN '"& saldoFraDato &"' AND '"& saldoTilDato &"' GROUP BY tfaktim"
                                               ' response.Write strSQLSaerligFerie
                                                oRec2.open strSQLSaerligFerie, oConn, 3
                                                if not oRec2.EOF then
                                                saerligFerie = oRec2("sumtimer")
                                                end if
                                                oRec2.close
                                               
                                                if normTimerDag <> 0 then
                                                saerligFerie = saerligFerie / normTimerDag 
                                                else
                                                saerligFerie = 0
                                                end if

                                                if saerligFerie <> 0 then
                                                saerligFerie = saerligFerie
                                                else
                                                saerligFerie = 0
                                                end if
                                            %>
                                            <input type="text" name="saerligFerie_<%=oRec("mid") %>" class="feriefriOpt_input form-control input-small" style="width:60px;" value="<%=formatnumber(saerligFerie, 2) %>" />
                                        </td>

                                        <!--
                                        <td>
                                            <%
                                               
                                                call getFerieFri(oRec("mid"), saldoFraDato, saldoTilDato)

                                                if normTimerDag <> 0 then
                                                fefriBal = fefriBal / normTimerDag
                                                else
                                                fefriBal = 0
                                                end if
                                            %>

                                            <%=formatnumber(fefriBal,2) %>
                                        </td>
                                        -->

                                        <td><input type="checkbox" name="godkend_<%=oRec("mid") %>" class="godkend_medarb" value="1" /></td>
                                    </tr>
                                <%
                                    oRec.movenext
                                    wend
                                    oRec.close
                                %>

                        </tbody>

                    </table>
                  

               
                  <div class="row">
                            <div class="col-lg-9">&nbsp</div>
                            <div class="col-lg-3 pad-b10"><button type="submit" class="btn btn-success btn-sm pull-right" style="width:80px;"><b><%=medarb_txt_020 %></b></button></div>
                        </div>

                  </form>
                     </div>
            </div>
        </div>



        <%end select %>













    </div>
</div>

<!-- hvis tabbellen skal scrolle -->
<!-- <script src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script> -->

<!--#include file="../inc/regular/footer_inc.asp"-->

