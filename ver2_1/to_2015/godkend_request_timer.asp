


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->





       
        
        <%

            

            dim tfaktimTot
            redim tfaktimTot(200)
            sub subTotaler


                                        select case oRec5("tfaktim")
                                        case 1
                                                tfaktimTot(0) = tfaktimTot(0) + oRec5("overtidtimer_onsket")
                                        
                                        case 10
                                                tfaktimTot(1) = tfaktimTot(1) + oRec5("overtidtimer_onsket")
                                        case 30
                                                tfaktimTot(2) = tfaktimTot(2) + oRec5("overtidtimer_onsket")
                                       
                                        case 50
                                                tfaktimTot(3) = tfaktimTot(3) + oRec5("overtidtimer_onsket")
                                        case 51
                                                tfaktimTot(4) = tfaktimTot(4) + oRec5("overtidtimer_onsket")
                                        case 52
                                                tfaktimTot(5) = tfaktimTot(5) + oRec5("overtidtimer_onsket")

                                         case 53
                                                tfaktimTot(6) = tfaktimTot(6) + oRec5("overtidtimer_onsket")
                                        case 54
                                                tfaktimTot(7) = tfaktimTot(7) + oRec5("overtidtimer_onsket")

                                         case 55
                                                tfaktimTot(8) = tfaktimTot(8) + oRec5("overtidtimer_onsket")

                                        case 60
                                                tfaktimTot(9) = tfaktimTot(9) + oRec5("overtidtimer_onsket")
                                        case 61
                                                tfaktimTot(10) = tfaktimTot(10) + oRec5("overtidtimer_onsket")

                                        case 90
                                                tfaktimTot(11) = tfaktimTot(11) + oRec5("overtidtimer_onsket")

                                        case 7
                                                tfaktimTot(12) = tfaktimTot(12) + oRec5("overtidtimer_onsket")


                                        case 28
                                                tfaktimTot(28) = tfaktimTot(28) + oRec5("overtidtimer_onsket")

                                        case 29
                                                tfaktimTot(29) = tfaktimTot(29) + oRec5("overtidtimer_onsket")
                                        case 120
                                                tfaktimTot(120) = tfaktimTot(120) + oRec5("overtidtimer_onsket")
                                        case 121
                                                tfaktimTot(121) = tfaktimTot(121) + oRec5("overtidtimer_onsket")
                                        case 123
                                                tfaktimTot(123) = tfaktimTot(123) + oRec5("overtidtimer_onsket")

                                        case 124
                                                tfaktimTot(124) = tfaktimTot(124) + oRec5("overtidtimer_onsket")

                                        case 125
                                                tfaktimTot(125) = tfaktimTot(125) + oRec5("overtidtimer_onsket")
                                            
                                        case else
                                         tfaktimTot(13) = tfaktimTot(13) + oRec5("overtidtimer_onsket") 
                                        end select

            end sub



            public overtidplusfaktorTot, overtidplusfaktorPrDag, overtidplusfaktorPrDagLunch 
            function tr_gk_request_timer(dothis, medid)

                                    '** dothis = 0 'Standard
                                    '** dothis = 1 'Tiembank Cflow - ekstra tom linje


                                     select case lto
                                        case "intranet - local"
                                            strSQlFakturerbar = "fakturerbar <> 0 "
                                            strSQLJjobnr = "jobnr = 863"

                                        case "wap"
                                            strSQlFakturerbar = "fakturerbar = 51"
                                           strSQLJjobnr = "j.risiko < 0 AND jobnr = 3"

                                        case "cflow"

                                            strSQlFakturerbar = " (fakturerbar = 7 OR fakturerbar = 10 OR fakturerbar = 30 OR fakturerbar = 32 OR fakturerbar = 50 OR fakturerbar = 51 OR fakturerbar = 52 OR fakturerbar = 54 OR fakturerbar = 50 OR fakturerbar = 53 OR fakturerbar = 55 OR fakturerbar = 60 OR fakturerbar = 61 OR fakturerbar = 90 OR fakturerbar = 91 OR fakturerbar = 92 OR fakturerbar = 93 OR fakturerbar = 17 OR fakturerbar = 20 OR fakturerbar = 21 OR fakturerbar = 120 OR fakturerbar = 121)"
                                            strSQLJjobnr = "j.risiko < 0 AND jobnr = 3"


                                       
                                        case else

                                           strSQlFakturerbar = "fakturerbar <> 1 AND fakturerbar <> 2"
                                           strSQLJjobnr = "j.risiko < 0" 'AND jobnr = 3

                                        end select

                                    
         
                                    if cint(dothis) = 0 then

                                    ntimPer = 0
                                    realtimerIalt = 0
                                    strSQL2 = "SELECT tmnavn, SUM(timer) AS realtimer FROM timer WHERE tmnr = "& medid &" AND ("& aty_sql_realhours &" OR tfaktim = 114) AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoEnd &"' GROUP BY tmnr"

                                    'if session("mid") = 1 then
                                    'Response.write strSQL2
                                    'Response.flush
                                    'end if

                                    oRec2.open strSQL2, oConn, 3
                                    if not oRec2.EOF then
                                    '114 = Korrektion Ultimo
                                    realtimerIalt = oRec2("realtimer")
                                    end if

                                    oRec2.close
                              
                                    call normtimerPer(medid, useStDatoKri, daysansat, 0)
                                    ntimPerFlex = ntimPer
                                    'balRealNormTxt = formatnumber(( realTimer(x) + korrektionReal(x) ) - normTimer(x),2)
                                    flexsaldo = (realtimerIalt - ntimPerFlex)

                                 %>
                                 <tr>
                                    <input type="hidden" name="tids" value="<%=oRec5("tid") %>" />
                                    <input type="hidden" name="tfaktim_<%=oRec5("tid")%>" value="<%=oRec5("tfaktim") %>" />
                                    <td valign="top"><%=oRec5("tmnavn") %></td>
                                    <td valign="top"><%=left(weekdayname(weekday(oRec5("tdato"), 1)), 3) &". "& formatdatetime(oRec5("tdato"), 2) %></td>

                                    <td valign="top">

                                        <%
                                            
                                        call normtimerPer(medid, oRec5("tdato"), 0, 0)  %>

                                        <%=ntimPer %>

                                    </td>


                                   


                                    <td>
                                         <select name="FM_aktids" class="form-control input-small" style="width:300px;">
                                        <%
                                        'call akttyper2009(2)    
                                        '"& replace(aty_sql_realhours, "tfaktim", "fakturerbar") &" AND     


                                        strSQLA = "SELECT a.id, a.navn FROM aktiviteter a WHERE a.id = " & oRec5("taktivitetid") 
                                        oRec.open strSQLA, oConn, 3
                                        if not oRec.EOF then

                                            %>
                                             <option value="<%=oRec("id")%>" SELECTED><%=oRec("navn")%></option>
                                             <%

                                        end if
                                        oRec.close

                                       

                                        strSQLakttype = "SELECT a.id, a.navn FROM job j "_
                                        &" LEFT JOIN aktiviteter a ON (a.job = j.id) WHERE ("& strSQlFakturerbar &") AND  ("& strSQLJjobnr &") AND a.id <> "& oRec5("taktivitetid") &" GROUP BY a.fakturerbar ORDER BY a.navn"
                                        oRec.open strSQLakttype, oConn, 3
                                        while not oRec.EOF
                                        
                                            
                                            'if cdbl(oRec5("taktivitetid")) = cdbl(oRec("id")) then
                                            'selAkt = "SELECTED"
                                            'else
                                            selAkt = ""
                                            'end if%>


                                       
                                            <option value="<%=oRec("id")%>" <%=selAkt %>><%=oRec("navn")%></option>
                                       
                                        <%
                                         oRec.movenext
                                         wend
                                         oRec.close%>
                                         </select>
                                        
                                        </td>

                                    <td valign="top"><%=oRec5("overtidtimer_onsket") %>
                                        <input type="hidden" id="timer_gk_opr_<%=oRec5("tid")%>" value="<%=oRec5("overtidtimer_onsket") %>" />
                                    </td>


                                     <%
                                         
                                     'select case datepart("w", oRec5("tdato"), 2,2) 
                                     'case "6","7"
                                     'faktor = "2"
                                     'case else
                                     'faktor = "1,5"
                                     'end select

                                     '*** FAKTOR 
                                     '***
                                     select case lto
                                     case "cflow"
                                         '** Altid 1:1, tillæg sker udelukkende på lønseddel
                                         faktor = "1"

                                     case else


                                         select case oRec5("tfaktim") 
                                         case "51","53"
                                         faktor = "2"
                                         case "54","55"
                                         faktor = "1,5"
                                         case else
                                         faktor = "1"
                                         end select

                                     end select



                                         faktor1sel = ""
                                         faktor25sel = ""
                                         faktor50sel = ""
                                         faktor75sel = ""
                                         faktor100sel = ""
                                         faktor150sel = ""
                                         faktor200sel = ""
                                     
                                        select case faktor
                                        case "1,5"
                                        faktor50sel = "SELECTED"
                                        case "2"
                                        faktor100sel = "SELECTED"
                                        case else 
                                        faktor1sel = "SELECTED"
                                        end select

                                     %>

                                    <td valign="top"><select  id="faktor_gk_<%=oRec5("tid")%>" name="faktor_gk_<%=oRec5("tid")%>" class="form-control input-small faktor" style="width:60px;">

                                                       <%select case lto
                                                        case "cflow"%>

                                                        <option value="1" <%=faktor1sel %>>1:1</option>
                                                        <option value="50" <%=faktor50sel %>>50%</option>
                                                        <option value="100" <%=faktor100sel %>>100%</option>
                                                        
                                                           
                                                        <%case else%>

                                                        <option value="1" <%=faktor1sel %>>1:1</option>
                                                        <option value="25" <%=faktor25sel %>>25%</option>
                                                        <option value="50" <%=faktor50sel %>>50%</option>
                                                        <option value="75" <%=faktor75sel %>>75%</option>
                                                        <option value="100" <%=faktor100sel %>>100%</option>
                                                        <option value="150" <%=faktor150sel %>>150%</option>
                                                        <option value="200" <%=faktor200sel %>>200%</option>

                                                           
                                                        <%end select%>
                                                
                                                       
                                                     </select>
                                    </td>
                                    <td valign="top">
                                        
                                        <%overtidplusfaktor = formatnumber((oRec5("overtidtimer_onsket") * faktor), 2)  %>


                                        <input type="text" id="timer_gk_<%=oRec5("tid")%>" name="timer_gk_<%=oRec5("tid")%>" value="<%=overtidplusfaktor %>" class="form-control input-small" style="width:60px; text-align:right;"  /></td>
                                        
                                        <%select case lto
                                        case "cflow"

                                            if oRec5("tfaktim") <> 10 then
                                            overtidplusfaktorTot = (overtidplusfaktorTot * 1) + (overtidplusfaktor * 1)
                                            overtidplusfaktorPrDag = (overtidplusfaktorPrDag*1) + (overtidplusfaktor * 1)
                                            else
                                            overtidplusfaktorTot = (overtidplusfaktorTot * 1) - (overtidplusfaktor * 1)
                                            overtidplusfaktorPrDag = (overtidplusfaktorPrDag*1) ' - (overtidplusfaktor * 1)
                                            overtidplusfaktorPrDagLunch = (overtidplusfaktorPrDagLunch*1) + (overtidplusfaktor * 1)
                                            end if

                                            
                                              

                                        case else%>
                                        <%overtidplusfaktorTot = (overtidplusfaktorTot * 1) + (overtidplusfaktor * 1) 
                                        overtidplusfaktorPrDag = (overtidplusfaktorPrDag*1) + (overtidplusfaktor * 1)
                                        overtidplusfaktorPrDagLunch = 0
                                        %>
                                        <%end select %>


                                        <!--(<%=overtidplusfaktorTot %>)-->
                                     
                                     <td>
                                         <%
                                         godkendtstatusSEL0 = ""
                                         godkendtstatusSEL1 = ""
                                         godkendtstatusSEL2 = ""
                                         godkendtstatusSEL3 = ""

                                         select case oRec5("godkendtstatus")
                                         case 1
                                         godkendtstatusSEL1 = "SELECTED"
                                         case 2 
                                         godkendtstatusSEL2 = "SELECTED"
                                         case 3 
                                         godkendtstatusSEL3 = "SELECTED"
                                         case else
                                         godkendtstatusCHK0 = ""    
                                         end select%>
                                         
                                       <select name="opdater_gkstatus_<%=oRec5("tid")%>" class="form-control input-small gkstatus">
                                           <option value="0" <%=godkendtstatusSEL0 %>>None</option>
                                           <option value="1" <%=godkendtstatusSEL1 %>>Godkendt</option>
                                           <option value="2" <%=godkendtstatusSEL2 %>>Afvist</option>
                                           <option value="3" <%=godkendtstatusSEL3 %>>Tentative</option>
                                       </select></td>
                                      
                                     <td><input type="checkbox" name="opdater_gk_<%=oRec5("tid")%>" value="1" class="gk" /></td>
                                     
                                     
                                     
                                     <%
                                         
                                         
                                         
                                     else 
                                     '*** Til TimebanK CFLOW%>



                                        <input type="hidden" name="tids7" value="0" />
                                        <input type="hidden" name="tids7_tmnr" value="<%=medid %>" />
                                         <td colspan="1">Tilføj:</td>
                                         <td colspan="2"><input class="form-control input-small" type="text" name="FM_dato7_<%=medid %>" value="dd-mm-yyyy" /></td>
                                         <td>
                                         
                                        <%
                                        'strSQLakttype = "SELECT a.id, a.navn FROM job j LEFT JOIN aktiviteter a ON (a.fakturerbar <> 0 AND a.job = j.id) WHERE (fakturerbar <> 0) AND j.risiko < 0 AND jobnr = 3 " '(fakturerbar = 7)
                                        
                                        strSQLakttype = "SELECT a.id, a.navn FROM job j "_
                                        &" LEFT JOIN aktiviteter a ON (a.job = j.id) WHERE ("& strSQlFakturerbar &") AND ("& strSQLJjobnr &") GROUP BY a.fakturerbar ORDER BY a.navn"
                                            
                                        'response.write strSQLakttype

                                            %>
                                        <select name="FM_aktids7_<%=medid %>" class="form-control input-small"><%

                                        oRec.open strSQLakttype, oConn, 3
                                        while not oRec.EOF
                                        
                                            
                                           %>
                                            <option value="<%=oRec("id")%>" <%=selAkt %>><%=oRec("navn")%></option>
                                       
                                        <%
                                         oRec.movenext
                                         wend
                                         oRec.close%>
                                         </select>
                                        
                                        </td>
                                      <td>&nbsp;</td>
                                       <td>&nbsp;</td>
                                     
                                      <td valign="top">

                                          <input type="text" id="timer_gk_7_<%=medid %>" name="timer_gk_7_<%=medid %>" value="0" class="form-control input-small" style="width:60px; text-align:right;"  /></td>
                                     
                                     <td>&nbsp;</td>
                                     <td><input type="checkbox" name="opdater_gk_7_<%=medid %>" value="1" class="gk" /></td>

                                     <%end if ' dothhis %>
                                    
                                    </tr>
                                <%


            end function





           

            func = request("func")

            if len(trim(request("usemrn"))) <> 0 then
            usemrn = request("usemrn")
            else
            usemrn = 0
            end if
            
            call meStamdata(usemrn)


         select case func
         case "gk"
           
                    '**** Opdaterer/flytter Godkendte timer ***'
                    dtNow = year(now) &"-"& month(now) &"-"& day(now)


                    'Response.write "TIDS: "& request("tids")

                    tids = split(request("tids"), ", ")
                    aktids = split(request("FM_aktids"), ", ")

                    
                    for t = 0 to UBOUND(tids)
            
                    if request("timer_gk_"& tids(t)) > 0 AND len(trim(request("timer_gk_"& tids(t)) )) <> 0 AND request("opdater_gk_"& tids(t)) = "1" then

                    timerThis = request("timer_gk_"& tids(t))
                    timerThis = replace(timerThis, ",", ".")

                    gkstatus = request("opdater_gkstatus_"& tids(t))

                   
                            
                               '*** Finder AKT data ***
                                strSQLakt = "SELECT a.navn, a.fakturerbar, k.kkundenavn, kid, jobnavn, jobnr FROM aktiviteter a "_
                                & "LEFT JOIN job j ON (j.id = a.job) "_
                                & "LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE a.id = " & aktids(t)
                                oRec9.open strSQLakt, oConn, 3
                                if not oRec9.EOF then

                                aktnavn = oRec9("navn")
                                fakturerbar = oRec9("fakturerbar")
                                kid = oRec9("kid")
                                kkundenavn = oRec9("kkundenavn")
                                jobnavn = oRec9("jobnavn") 
                                jobnr = oRec9("jobnr")

                                end if
                                oRec9.close


                    tFakTimSQL = " tfaktim = "& fakturerbar &","
                    aktNavnSQL = " taktivitetid = "& aktids(t) &", taktivitetnavn = '"& aktnavn &"',"
                    tjobkundeSQL =  " tjobnavn = '"& jobnavn &"', tjobnr = '"& jobnr &"', tknavn = '"& kkundenavn &"', tknr = "& kid &","
		                        
		                 

                    strSQlupd = "UPDATE timer SET "& tFakTimSQL &" "& aktNavnSQL &" "& tjobkundeSQL &" timer = "& timerThis &", godkendtstatus = "& gkstatus &", godkendtstatusaf = '"& session("user") &"', godkendtdato = '"& dtNow &"' WHERE tid = "& tids(t)
                    'Response.write strSQlupd
                    'Response.flush

                    oConn.execute(strSQlupd)

                    end if

                
                    '*** Slet ***'    
                    if request("timer_gk_"& tids(t)) = 0 AND len(trim(request("timer_gk_"& tids(t)) )) <> 0 AND request("opdater_gk_"& tids(t)) = "1" then

                    strSQlupd = "DELETE FROM timer WHERE tid = "& tids(t)
                    'SET godkendtstatus = 2, godkendtstatusaf = '"& session("user") &"', godkendtdato = '"& dtNow &"' 

                    oConn.execute(strSQlupd)

                    end if


                    next
            

                    '*** Timebank CFLOW


                    'Response.write "    HER: "& request("tids7")
                    'Response.end

                    tids7 = split(request("tids7"), ", ")
                    'aktids7 = split(request("FM_aktids7"), ", ")
                    tids7_tmnr = split(request("tids7_tmnr"), ", ")
                    
                                
                                   
                    insertUpdate = 1 '1: FROM godkendrequest timer
                    t = 0
                    for t = 0 to UBOUND(tids7)

                    'Response.write "    HER: "& t & " timer: "& request("timer_gk_7_"& tids7_tmnr(t)) &" "& request("timer_gk_7_"& tids7_tmnr(t)) &  "#"& request("opdater_gk_7_"& tids7_tmnr(t))  &" dato: "& request("FM_dato7") &"<br>"
                    'Response.flush
                    'Response.end
            
                    if request("timer_gk_7_"& tids7_tmnr(t)) > 0 AND len(trim(request("timer_gk_7_"& tids7_tmnr(t)) )) <> 0 AND request("opdater_gk_7_"& tids7_tmnr(t)) = "1" AND tids7_tmnr(t) <> 0 then


                        if isDate(request("FM_dato7_"& tids7_tmnr(t))) = true then

                                 

                        'Response.write "<br>    HER 999: "& aktids7(t)
                        timerThis = request("timer_gk_7_"& tids7_tmnr(t))
                        tids7_dato = request("FM_dato7_"& tids7_tmnr(t))   
                        tids_aktid = request("FM_aktids7_"& tids7_tmnr(t))
                        extsysid = 0
                        timerkom = "" 
                        koregnr = ""
                        destination = ""
                        usebopal = 0

                        call indlasTimerTfaktimAktid(lto, tids7_tmnr(t), timerThis, 7, tids_aktid, insertUpdate, tids7_dato, extsysid, timerkom, koregnr, destination, usebopal)

                        else


                                %>

                                            <div class="wrapper">
                                             <div class="content">   

                                                    <div class="container">
                                                        <div class="portlet">
                                                                 <h3 class="portlet-title"><u>Fejl</u> </h3>
                                                                    <div class="portlet-body">

                                                             <div class="well well-white">

                                                            <div class="row">

                                                                    
                                                                     <div class="col-lg-3">
                                                                         Fejl<br />
                                                                         Den valgte dato er ikke gyldig.<br /><br />
                                                                         <a href="Javascript:history.back();"><< Tilbage</a>
                                                                     </div>

                                                            </div>
                                                                    </div><!-- well -->

                                                                        </div>
                                        </div>
                                                        </div>
                                                 </div>
                                                </div>


                                <% Response.end


                        end if
                    
                        thisMrn = tids7_tmnr(t)

                    end if



                    next

                    'Response.end
                    response.redirect "godkend_request_timer.asp?visning=0&varTjDatoUS_man="& request("varTjDatoUS_man") &"&FM_medarb="& thisMrn

              

        case else


        call akttyper2009(2)
        call licensStartDato()
        licensstdato = licensstdato
        meAnsatDato = meAnsatDato

        'if cDate(meAnsatDato) > cDate(licensstdato) then
        'useStDatoKri = meAnsatDato 
        'else
        'useStDatoKri = licensstdato
        'end if

        call lonKorsel_lukketPerPrDato(now, usemrn)
        'if cDate(lonKorsel_lukketPerDt) > cDate(useStDatoKri) then
        'useStDatoKri = dateAdd("d", 1, day(lonKorsel_lukketPerDt) &"/"& month(lonKorsel_lukketPerDt) &"/"& year(lonKorsel_lukketPerDt))
        'else
        'useStDatoKri = useStDatoKri
        'end if



        '**** DATOER *****'



            brugugeCHK = ""
    if len(trim(request("bruguge"))) <> 0 OR (request.cookies("2015")("bruguge") = "1" AND len(trim(request("sogsubmitted"))) = 0) then

        brugugeCHK = "CHECKED"

        if len(trim(request("bruguge_week"))) <> 0 then 
        bruguge_week = request("bruguge_week")
        else
        bruguge_week = request.cookies("2015")("bruguge_week")
        end if

        response.cookies("2015")("bruguge_week") = bruguge_week
	                  
                        
	            stDato = "1/1/"&year(now)
	            datoFound = 0
	    
	            for u = 1 to 53 AND datoFound = 0
	   
	            if u = 1 then
	            tjkDato = stDato
	            else
	            tjkDato = dateadd("d",7,tjkDato)
	            end if
	    
                'response.Write "tjkDato: " & tjkDato & "<br>"
                call thisWeekNo53_fn(tjkDato)
	            tjkDatoW = thisWeekNo53 'datepart("ww", tjkDato, 2,2)
	    
	            if cint(bruguge_week) = cint(tjkDatoW) then
	    
	            wDay = datepart("w", tjkDato, 2,2)
	       
	        
	                select case wDay
	                case 1
	                tjkDato = dateAdd("d", 0, tjkDato)
	                case 2
	                tjkDato = dateAdd("d", -1, tjkDato)
	                case 3
	                tjkDato = dateAdd("d", -2, tjkDato)
	                case 4
	                tjkDato = dateAdd("d", -3, tjkDato)
	                case 5
	                tjkDato = dateAdd("d", -4, tjkDato)
	                case 6
	                tjkDato = dateAdd("d", -5, tjkDato)
	                case 7
	                tjkDato = dateAdd("d", -6, tjkDato)
	                end select
	    
	            stDaguge = day(tjkDato)
	            stMduge = month(tjkDato)
	            stAaruge = year(tjkDato)
	    
	            tjkDato_slut = dateadd("d", 6, tjkDato)
	    
	            slDaguge = day(tjkDato_slut)
	            slMduge = month(tjkDato_slut)
	            slAaruge = year(tjkDato_slut)
	    
	       
	            datoFound = 1
	    
	            end if
	    
	            next

         aar = stDaguge &"-"& stMduge &"-"& stAaruge
        aarslut = slDaguge &"-"& slMduge &"-"& slAaruge

        response.cookies("2015")("bruguge") = "1"

    else

        response.cookies("2015")("bruguge") = ""
        call thisWeekNo53_fn(now)
        bruguge_week = thisWeekNo53 'datepart("ww", now, 2,2)

            if len(trim(request("aar"))) <> 0 then
            aar = request("aar")

                response.cookies("2015")("gktimer_aar") = aar

            else

                if request.cookies("2015")("gktimer_aar") <> "" then
                aar = request.cookies("2015")("gktimer_aar")
                else
                aar = "1-"& month(now) &"-"& year(now)
                end if

            end if

            if len(trim(request("aarslut"))) <> 0 then
            aarslut = request("aarslut")

            response.cookies("2015")("gktimer_aarslut") = aarslut

            else
   
                if request.cookies("2015")("gktimer_aarslut") <> "" then
                aarslut = request.cookies("2015")("gktimer_aarslut")
                else
                aarslut = dateAdd("m", 1, aar)
                end if

            end if


    end if





        '***** End Datoer





      



        varTjDatoUS_manSQL = year(aar) & "/" & month(aar) & "/" & day(aar)
        varTjDatoUS_sonSQL = year(aarslut) & "/" & month(aarslut) & "/" & day(aarslut)


        








       
         visning = 0
         if len(trim(request("FM_visning"))) <> 0 then 

            visning = request("FM_visning")
            select case visning
            case 0
            visningSEL0 = "SELECTED"
            case 1
            visningSEL1 = "SELECTED"
            end select

                                        
        else 
            visningSEL0 = "SELECTED"
            visning = 0
        end if


        'if cint(visning) = 0 then

        if len(trim(request("FM_medarb"))) <> 0 then
        usemrn = request("FM_medarb")
        else
             if request.Cookies("tsa")("selmedreqtimer") <> "" then
             usemrn = request.Cookies("tsa")("selmedreqtimer")
             else
             usemrn = 0
             end if
        end if

        'else

        '    usemrn = 0

        'end if


        response.Cookies("tsa")("selmedreqtimer") = usemrn
                    
        if cdbl(usemrn) <> 0 then
            usrmnrSQL = " tmnr = "& usemrn &""
        else
            usrmnrSQL = " tmnr <> 0"
        end if


        

        ignperDis = ""
        if len(trim(request("FM_medarb"))) <> 0 then

                if len(trim(request("FM_ignper"))) <> 0 then
                ignper = 1
                ignperCHK = "CHECKED"
                else
                ignper = 0
                ignperCHK = ""
                end if

        else

                if request.Cookies("tsa")("selmedreqignper") = "1" then
                    ignper = 1
                    ignperCHK = "CHECKED"
                else
                    ignper = 0
                    ignperCHK = ""
                end if

        end if

          if usemrn = 0 OR visning = 1 then
                ignperDis = "DISABLED"
                ignper = 0
                ignperCHK = ""
          end if

        response.Cookies("tsa")("selmedreqignper") = ignper


        if cint(ignper) <> 1 then
        datoSQLkri = " AND tdato BETWEEN '"& varTjDatoUS_manSQL & "' AND '"& varTjDatoUS_sonSQL &"'"
        else
        datoSQLkri = ""
        end if   
                 
                        
        if len(trim(request("FM_medarb"))) <> 0 then

                if len(trim(request("FM_viskunikkegk"))) <> 0 then
                viskunikkegk = 1
                viskunikkegkCHK = "CHECKED"
                else
                viskunikkegk = 0
                viskunikkegkCHK = ""
                end if

        else
                  
                        
                 if request.Cookies("tsa")("selmedreqviskunikkegk") = "1" then
                    viskunikkegk = 1
                    viskunikkegkCHK = "CHECKED"
                 else
                    viskunikkegk = 0
                    viskunikkegkCHK = ""

                end if

     
        end if

        response.Cookies("tsa")("selmedreqviskunikkegk") = viskunikkegk




        if len(trim(request("FM_medarb"))) <> 0 then

                if len(trim(request("FM_viskunikkegk1"))) <> 0 then
                viskunikkegk1 = 1
                viskunikkegkCHK1 = "CHECKED"
                else
                viskunikkegk1 = 0
                viskunikkegkCHK1 = ""
                end if

        else
                  
                        
                 if request.Cookies("tsa")("selmedreqviskunikkegk1") = "1" then
                    viskunikkegk1 = 1
                    viskunikkegkCHK1 = "CHECKED"
                 else
                    viskunikkegk1 = 0
                    viskunikkegkCHK1 = ""

                end if

     
        end if

        response.Cookies("tsa")("selmedreqviskunikkegk1") = viskunikkegk1



        if cint(viskunikkegk) = 1 then
        gkSQLkri = " AND (godkendtstatus = 0 OR godkendtstatus = 1 OR godkendtstatus = 3"
        else
        gkSQLkri = " AND (godkendtstatus = 0 OR godkendtstatus = 3"
        end if

        if cint(viskunikkegk1) = 1 then
        gkSQLkri1 = " OR godkendtstatus = 2"
        else
        gkSQLkri1 = ""
        end if

        gkSQLkri = gkSQLkri & gkSQLkri1 & ")"
        
      

        if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> 0 then
        nomenu = 1
        else
        nomenu = 0
        end if



        if cint(nomenu) = 0 then
            call menu_2014()
        end if

        %>
        <script src="js/godkend_request_timer_jav.js" type="text/javascript"></script>

        <div class="wrapper">
 <div class="content">   

        <div class="container">
            <div class="portlet">
                   
                  <h3 class="portlet-title"><u>Godkend Ferie, Overtid & Løntimer</u> </h3>
                  

                <div class="portlet-body">
                     <!--<div class="portlet-title">
                          <table style="width:100%;">
                              <tr>
                                  <td><h3 style="color:black"><%=medarb_navn %></h3></td>
                                  <td style="text-align:right"><img style="width:15%; margin-bottom:2px; margin-left:55px" src="img/cflow_logo.png" /></td>
                              </tr>
                          </table>
                        
                        </div>-->

                  
                       <div class="well well-white">
                      <form action="godkend_request_timer.asp?sogsubmitted=1" method="post">
                        
                                <div class="row">
                                    <div class="col-lg-6">Medarbejder: 

                                        <%

                                            call selectMedarbPrProgrp(1, 0, "FM_medarb", 1, 0)

                                            %>
                                      </div>


                                   <%if session("stempelur") <> 0 then %>

                                    <div class="col-lg-3">Visning: <select name="FM_visning" class="form-control input-small" onchange="submit();">
                                        <option value="0" <%=visningSEL0 %>>Vis åben for godkendelse (rediger)</option>
                                        <option value="1" <%=visningSEL1 %>>Vis liste og afvigelser (oversigt)</option>
                                                                   </select></div>

                                    <%end if %>

                                  

                                     </div>





                             <%
                                 antalmaaned = 0 'DateDiff("m",aar,aarslut, 2,2)
                                antalaar = 0 'DateDiff("yyyy",aar,aarslut, 2,2)
                                                                                                   
                            %>

                            

                          <div class="row">
                                
                                 
                           
                            <div class="col-lg-2">
                                
                                <%=godkend_txt_003 %>:
                                <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="aar" value="<%=aar %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                                <br />
                           </div>
                                 <div class="col-lg-2">
                               
                                <%=godkend_txt_004 %>:<br />
                                <div class='input-group date' id='datepicker_sldato'>
                                <input type="text" class="form-control input-small" name="aarslut" value="<%=aarslut %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                           </div>
                 
                  <div class="col-lg-1">
                    
                      <%call thisWeekNo53_fn(now) %>

                <input class="dateweeknum" id="bruguge" name="bruguge" type="checkbox" value="1" <%=brugugeCHK %> />  Week:
                <input type="text" id="bruguge_selector_of" class="form-control input-small" value="<%=thisWeekNo53 %>" readonly>
			<select name="bruguge_week" id="bruguge_selector" class="form-control input-small" onchange="submit();" style="display:none" >
			<% for w = 1 to 52
			
			'if w = 1 then
			'useWeek = 1
			'else
			'useWeek = dateadd("ww", 1, useWeek)
			'end if
			
			if cint(w) = cint(bruguge_week) then
			wSel = "SELECTED"
			else
			wSel = ""
			end if
			
			%>
			<option value="<%=w%>" <%=wSel%>><%=w %></option>
			<%
			next %>
			
			</select>
               
          
                </div>
        </div><!-- row -->


                          <div class="row">
                                    <%if cint(visning) = 1 then

                                        
                                        viskunDIS = "DISABLED"
                                        viskunDIS1 = "DISABLED"

                                    end if %>



                              

                                    <div class="col-lg-9">
                              
                                        <!--<input type="radio" value="0" name="FM_ignper" id="FM_ignper0" <%=ignper0CHK %> <%=ignper0Dis %> /> Vis uger<br />
                                        <span class="inline"><input type="radio" value="1" name="FM_ignper" id="FM_ignper" <%=ignperCHK %> <%=ignperDis %> /> Ignorer periode <input type="text" name="FM_fradato" value="<%=fraDato %>" class="form-control input-small" style="width:80px;" /> - <input type="text" name="FM_tildato" value="<%=tilDato %>" class="form-control input-small" style="width:80px;" /></span><br />
                                        <input type="radio" value="2" name="FM_ignper" id="FM_ignper2" <%=ignper2CHK %> <%=ignper2Dis %> /> Vis løn periode &nbsp;<br />-->

                                    <input type="checkbox" value="1" name="FM_viskunikkegk" <%=viskunikkegkCHK %> <%=viskunDIS %> /> Vis godkendte timer<br />
                                    <input type="checkbox" value="1" name="FM_viskunikkegk1" <%=viskunikkegkCHK1 %> <%=viskunDIS1 %> /> Vis afviste timer


                                    </div>
                                    
                                      <div class="col-lg-3" style="text-align:right; padding-right:40px;"><button type="submit" class="btn btn-info"><%=medarb_txt_100 %> >> </button></div>
                                  
                                </div>
                        
                    
                           
                         
                          </form>
                       </div>
              
            

                        <%  
                            
                            select case lto
                            case "wap"
                                 sqlTfaktim = " AND tfaktim = 51 "
                                 timeKolTxt = "AFTALE-home ønsket"
                            case "intranet - local"
                                 sqlTfaktim = " AND tfaktim <> 0 "
                                 timeKolTxt = "AFTALE-home ønsket"
                            case "cflow"
                                 sqlTfaktim = " AND (tfaktim = 7 OR tfaktim = 10 OR tfaktim = 30 OR tfaktim = 32 OR tfaktim = 50 OR tfaktim = 51 OR tfaktim = 52 OR tfaktim = 54 OR tfaktim = 50 OR tfaktim = 53 OR tfaktim = 55 OR tfaktim = 60 OR tfaktim = 61 OR tfaktim = 90 OR tfaktim = 91 OR tfaktim = 92 OR tfaktim = 17 OR tfaktim = 20 OR tfaktim = 21 OR tfaktim = 28 OR tfaktim = 29 OR tfaktim = 123 OR tfaktim = 124 OR tfaktim = 125 OR tfaktim = 120 OR tfaktim = 121)"
                                 timeKolTxt = "Timer ønsket"
                            case else
                                 sqlTfaktim = " AND (tfaktim <> 1 AND tfaktim <> 2 ANd tfaktim <> 5)"
                                 timeKolTxt = "Timer ønsket"
                            end select 
                            
                            cspan = 9%>
                       
                      
                            
                       



                                 <%  if (session("stempelur") <> 0 AND usemrn <> 0 AND ignper <> 1 AND visning = 0) OR (cint(visning) = 1) then  %>
                                <div class="row">

                                 <%if cint(visning) = 0 then %>
                                <div class="col-lg-5">
                                <%else %>
                                <div class="col-lg-12">
                                <%end if %>

                                    <br />
                                    <%if cint(visning) = 0 then %>
                                    <%call meStamdata(usemrn) %>
                                    
                                    <h4><%=meNavn & " ["& meInit &"]"%></h4> 
                                    <%end if %>

                                          
                                            <%call logindhistorik_week_60_100(usemrn, visning, varTjDatoUS_manSQL, varTjDatoUS_sonSQL) %> 
                                
                                        </div>

                                    

                                </div>
                                 <%end if %>



                            <%if visning = 0 then %>

                            <div class="row">
                             <div class="col-lg-12">

                                   <form action="godkend_request_timer.asp?func=gk&varTjDatoUS_man=<%=varTjDatoUS_man %>" method="post">
                            <!--Flexsaldo: <b><%=flexsaldo %></b><br />-->
                            <table class="table table-striped">
                                <thead>
                                    <tr>
                                    <th>Navn</th>
                                    <th>Dato</th>
                                    <th>Norm.</th>
                                    <%if session("stempelur") <> 0 then
                                       cspan = cspan + 1%>
                                    <!--<th>Komme/gå tid</th>-->
                                    <%end if %>
                                    <th>Aktivitet</th>
                                    <th><%=timeKolTxt %></th>
                                    <th>+ Faktor</th>
                                    <th>Godkend antal<br /> <span style="background-color:yellow; padding:4px;">(0: slet)</span></th>
                                        <th>Status
                                           <select id="gkstatustall" class="form-control input-small">
                                           <option value="0">None</option>
                                           <option value="1">Godkendt</option>
                                           <option value="2">Afvist</option>
                                           <option value="3">Tentative</option>
                                       </select>

                                        </th>
                                    <th>Opdater <input type="checkbox" id="gkall" /> 
                                        
                                    </th>
                                    </tr>
                                </thead>
                                
                             <%
                         
                            overtidplusfaktorTot = 0
                            overtidplusfaktorPrDag = 0
                            overtidplusfaktorPrDagLunch = 0
                            lasttdato = "01-01-2001"
                            'overtidtot = 0
                            m = 0
                            '(godkendtstatus = 0 OR godkendtstatus = 3)
                            strSQL6 = "SELECT tmnr, tmnavn, timer As overtidtimer_onsket, tdato, tid, tfaktim, taktivitetnavn, taktivitetid, godkendtstatus FROM timer WHERE "& usrmnrSQL &" "& sqlTfaktim &" "& datoSQLkri &" "& gkSQLkri &" ORDER BY tmnavn, tdato, tfaktim" '"& usemrn &" AND tdato BETWEEN '"& sqlDatoStartATD &"' AND '"& sqlDatoEnd &"'
                            'Response.write strSQL6
                            'Response.Flush
                            oRec5.open strSQL6, oConn, 3
                            while not oRec5.EOF


                                 if cdbl(lasttmnr) <> cdbl(oRec5("tmnr")) AND m > 0 then   
            
                                        call tr_gk_request_timer(1, oRec5("tmnr")) 

                                 %>
                                 <tr><td colspan="<%=cspan %>"><br />&nbsp;</td></tr>
                                 <%
                                 end if


                                 if cdate(lasttdato) <> cdate(oRec5("tdato")) AND m > 0 then   

                                  overtidplusfaktorPrDagMinusLunch = (overtidplusfaktorPrDag - overtidplusfaktorPrDagLunch)
                                 %>
                                  <tr><td colspan="<%=cspan-4 %>">&nbsp;</td><td align="right"><%=formatnumber(overtidplusfaktorPrDag, 2) %> - <%=formatnumber(overtidplusfaktorPrDagLunch, 2) %> = <%=formatnumber(overtidplusfaktorPrDagMinusLunch, 2) %> </td><td colspan="2">&nbsp;</td></tr>
                                 <%
                                     overtidplusfaktorPrDag = 0
                                     overtidplusfaktorPrDagLunch = 0
                                     overtidplusfaktorPrDagMinusLunch = 0
                                 end if

                                     


                                call tr_gk_request_timer(0, oRec5("tmnr"))
                            
                                select case lto 
                                        case "cflow"
                                        call subTotaler
                                        case else
                                        end select
                                    
                               
                                lasttmnr = oRec5("tmnr")
                                lasttdato = oRec5("tdato")
                                

                                m = m + 1

                            oRec5.movenext
                            wend 
                            oRec5.close
                                     

                                 if m > 0 then   
                                 overtidplusfaktorPrDagMinusLunch = (overtidplusfaktorPrDag - overtidplusfaktorPrDagLunch)
                                 %>
                                  <tr><td colspan="<%=cspan-4 %>">&nbsp;</td><td align="right"><%=formatnumber(overtidplusfaktorPrDag, 2) %> - <%=formatnumber(overtidplusfaktorPrDagLunch, 2) %> = <%=formatnumber(overtidplusfaktorPrDagMinusLunch, 2) %> </td><td colspan="2">&nbsp;</td></tr>
                                 <%
                                     overtidplusfaktorPrDag = 0
                                     overtidplusfaktorPrDagLunch = 0
                                     overtidplusfaktorPrDagMinusLunch = 0
                                 end if

                            if m > 0 then
                            medid = lasttmnr
                            else

                                if usemrn <> 0 then 
                                medid = usemrn
                                else
                                medid = 0
                                end if

                            end if

                              call tr_gk_request_timer(1, medid)%>


                                <%

                                 select case lto
                                 case "cflow"

                                         totTfaktimlabelTxt = ""   
                                         for t = 1 to 12 


                                            if tfaktimTot(t) <> 0 then
                                    
                                            select case t
                                            case 1 '10
                                            totTfaktimlabelTxt = "Lunch"
                                            case 2 '30
                                            totTfaktimlabelTxt = "Ordinær"
                                            case 3 '50
                                            totTfaktimlabelTxt = "Utetillæg innland"
                                            case 4 '51
                                            totTfaktimlabelTxt = "Overtid 100%"
                                            case 5 '52
                                            totTfaktimlabelTxt = "Reisetid"
                                            case 6 '53
                                            totTfaktimlabelTxt = "Reisetid 100%"
                                            case 7 '54
                                            totTfaktimlabelTxt = "Overtid 50%"
                                            case 8 '55
                                            totTfaktimlabelTxt = "Reisetid 50%"
                                            case 9 '60
                                            totTfaktimlabelTxt = "Utetillæg udland"
                                            case 10 '61
                                            totTfaktimlabelTxt = "Teamleder tillæg"
                                            case 11 '90
                                            totTfaktimlabelTxt = "Etter Midnatt"
                                            case 12 '7
                                            totTfaktimlabelTxt = "Timebank"
                                            end select
                                   
                                            %>

                                          <tr>
                                        <td colspan="<%=cspan-5 %>">&nbsp;</td>
                                              <td><%=totTfaktimlabelTxt %>:</td>
                                        <td>  <input type="text" value="<%=formatnumber(tfaktimTot(t), 2) %>" class="form-control input-small" style="width:60px; text-align:right;" disabled  /></td>
                                        <td>&nbsp;</td><td>&nbsp;</td>
                                        </tr>


                                        <%
                                         end if 'tfaktimTot(t) <> 0   
                                            
                                         next 
                                    
                                end select

                                 select case lto
                                 case "cflow"
                                    totTimerTxt = "Ialt til H&L" 
                                 case else
                                    totTimerTxt = "Ialt:"
                                 end select           
                                            
                                  %>
                                                            <tr>
                                <td colspan="<%=cspan-5 %>">&nbsp;</td>
                                <td>&nbsp;<%=totTimerTxt %>:</td>
                                <td>  <input type="text" value="<%=formatnumber(overtidplusfaktorTot, 2) %>" class="form-control input-small" style="width:60px; text-align:right;" disabled  /></td>
                                <td>&nbsp;</td><td>&nbsp;</td>
                            </tr>
                            <tr><td colspan="<%=cspan %>" align="right"> <input type="submit" class="btn btn-small btn-success" value="Opdater >>"></td></tr>

                            </table><br />

                                
                                       </form>
                          </div>
                                
                            </div>
                       <%end if 'visning %>
                           

                                        </div>

                        
            </div>
        </div>


              </div> <!-- content -->
        </div><!-- wrapper -->

<%end select %>

<!--#include file="../inc/regular/footer_inc.asp"-->