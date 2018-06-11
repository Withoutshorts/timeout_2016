


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->





       
        
        <%

            public timerThisDayTot, thisDayC
            function dagsTotal(timerThisDayTot)

                        

                        'timerThisTot = 0
                        timerThisTot = formatnumber(timerThisDayTot/60, 2)
                        timerThisTot_komma = instr(timerThisTot, ",")
                        timerThisTot = left(timerThisTot,  timerThisTot_komma - 1)

                        minutterThisTot = 0
                        minutterThisTot = formatnumber(timerThisDayTot/60, 2)
                        minutterThisTot = right(minutterThisTot, 2)
                        minutterThisTot = formatnumber((minutterThisTot*60)/100, 0)
                                          

                        'minutterThis = formatnumber(minutterThis * 60/100, 0)
                        minutter100Tot = formatnumber(minutterThisTot/60*100, 0) 

                        if minutterThisTot < 10 then
                        minutterThisTot = "0"& minutterThisTot
                        end if

                        if minutter100Tot < 10 then
                        minutter100Tot = "0"& minutter100Tot
                        end if


               

                Response.write "<tr>"_
                &"<td>&nbsp;</td>"_
                &"<td>&nbsp;</td>"_
                &"<td>&nbsp;</td>"_
                &"<td style=""width:100px;"">&nbsp;</td><td align=right><b>"& timerThisTot &":"& minutterThisTot &"</b></td>"_
                &"<td align=right><b>"& timerThisTot &","& minutter100Tot &"</b></td>"_
                &"</tr>" 

                timerThisDayTot = 0
                thisDayC = 0

                

            end function

            dim tfaktimTot
            redim tfaktimTot(13)
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

                                            strSQlFakturerbar = " (fakturerbar = 7 OR fakturerbar = 10 OR fakturerbar = 30 OR fakturerbar = 51 OR fakturerbar = 52 OR fakturerbar = 54 OR fakturerbar = 50 OR fakturerbar = 53 OR fakturerbar = 55 OR fakturerbar = 60 OR fakturerbar = 61 OR fakturerbar = 90)"
                                            strSQLJjobnr = "j.risiko < 0 AND jobnr = 3"

                                        case else
                                           strSQlFakturerbar = "fakturerbar <> 1 AND fakturerbar <> 2"
                                           strSQLJjobnr = "j.risiko < 0 AND jobnr = 3"
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
                                        &" LEFT JOIN aktiviteter a ON (a.job = j.id) WHERE ("& strSQlFakturerbar &") AND  ("& strSQLJjobnr &") AND a.id <> "& oRec5("taktivitetid") &""
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
                                        &" LEFT JOIN aktiviteter a ON (a.job = j.id) WHERE ("& strSQlFakturerbar &") AND ("& strSQLJjobnr &")"
                                            
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

                        call indlasTimerTfaktimAktid(lto, tids7_tmnr(t), timerThis, 7, tids_aktid, insertUpdate, tids7_dato)

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

        call lonKorsel_lukketPerPrDato(now)
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

	            tjkDatoW = datepart("ww", tjkDato, 2,2)
	    
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

        bruguge_week = datepart("ww", now, 2,2)

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





        'if len(trim(request("varTjDatoUS_man"))) <> 0 then

         '   varTjDatoUS_man = request("varTjDatoUS_man")
        
        'else

         '   varTjDatoUS_man = year(now)&"/"&month(now)&"/"&day(now)

        'end if

         
         '   mandag = datePart("w", varTjDatoUS_man, 2,2) 
           

         '       if cint(mandag) > 1 then
         '       varTjDatoUS_man = dateAdd("d", -(mandag - (mandag-1)), varTjDatoUS_man)
         '       end if

          

         '   if cint(mandag) > 1 then
         '   varTjDatoUS_man = dateAdd("d", -(mandag - (mandag-1)), varTjDatoUS_man)
         '   end if

      

        'varTjDatoUS_man_minus = dateAdd("d", -7, varTjDatoUS_man)
        'varTjDatoUS_man_plus = dateAdd("d", 7, varTjDatoUS_man)
        'varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

        'useStDatoKri = varTjDatoUS_man

        'sqlDatoStart = year(useStDatoKri)&"/"&month(useStDatoKri)&"/"&day(useStDatoKri)
        'sqlDatoStartATD =  year(now)&"/1/1"

     
        'ddNowMinusOne = dateAdd("d", -1, now)

        'sqlDatoEnd = year(ddNowMinusOne)&"/"&month(ddNowMinusOne)&"/"&day(ddNowMinusOne)

        'daysAnsat = dateDiff("d", sqlDatoStart, ddNowMinusOne)

        'varTjDatoUS_manSQL = year(varTjDatoUS_man) & "/" & month(varTjDatoUS_man) & "/" & day(varTjDatoUS_man)
        'varTjDatoUS_sonSQL = year(varTjDatoUS_son) & "/" & month(varTjDatoUS_son) & "/" & day(varTjDatoUS_son)




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


        if cint(viskunikkegk) = 1 then
        gkSQLkri = " AND (godkendtstatus = 0 OR godkendtstatus = 1 OR godkendtstatus = 2 OR godkendtstatus = 3)"
        else
        gkSQLkri = " AND (godkendtstatus = 0 OR godkendtstatus = 3)"
        end if

        
      

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
                                    <div class="col-lg-6">Medarbejder: <select name="FM_medarb" class="form-control input-small" onchange="submit();">
                                        <option value="0">Vis alle</option>

                                        <%strSQLm = "SELECT mnavn, mid, init FROM medarbejdere WHERE mansat = 1 ORDER BY mnavn"
                                            
                                            oRec.open strSQLm, oConn, 3
                                            While not oRec.EOF 
                                            
                                            if len(trim(oRec("init"))) <> 0 then
                                            init = "["& oRec("init") &"]"
                                            else
                                            init = ""
                                            end if

                                            if cdbl(usemrn) = cdbl(oRec("mid")) then
                                            mSEL = "SELECTED"
                                            else
                                            mSEL = ""
                                            end if
                                            %>
                                          <option value="<%=oRec("mid") %>" <%=mSEL %>><%=oRec("mnavn") & " "& init %></option>
                                           <%
                                            oRec.movenext
                                            Wend
                                            oRec.close
                                            
                                         %>

                                    


                                                 </select>
                                      </div>


                                   <%if session("stempelur") <> 0 then %>

                                    <div class="col-lg-3">Visning: <select name="FM_visning" class="form-control input-small" onchange="submit();">
                                        <option value="0" <%=visningSEL0 %>>Vis åben for godkendelse (rediger)</option>
                                        <option value="1" <%=visningSEL1 %>>Vis liste og afvigelser (oversigt)</option>
                                                                   </select></div>

                                    <%end if %>

                                    <div class="col-lg-3"><br /><button type="submit" class="btn btn-sm"><b><%=medarb_txt_100 %> >></b></button></div>

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
                    

                <input class="dateweeknum" id="bruguge" name="bruguge" type="checkbox" value="1" <%=brugugeCHK %> />  Week:
                <input type="text" id="bruguge_selector_of" class="form-control input-small" value="<%=Datepart("ww",now) %>" readonly>
			<select name="bruguge_week" id="bruguge_selector" class="form-control input-small" onchange="submit();" style="display:none" >
			<% for w = 1 to 52
			
			if w = 1 then
			useWeek = 1
			else
			useWeek = dateadd("ww", 1, useWeek)
			end if
			
			if cint(datepart("ww", useWeek, 2,2)) = cint(bruguge_week) then
			wSel = "SELECTED"
			else
			wSel = ""
			end if
			
			%>
			<option value="<%=datepart("ww", useWeek, 2,2) %>" <%=wSel%>><%=datepart("ww", useWeek, 2,2) %></option>
			<%
			next %>
			
			</select>
               
          
                </div>
        </div><!-- row -->


                          <div class="row">
                                    <%if cint(visning) = 1 then

                                        
                                        viskunDIS = "DISABLED"

                                    end if %>



                              

                                    <div class="col-lg-6">
                              
                                        <!--<input type="radio" value="0" name="FM_ignper" id="FM_ignper0" <%=ignper0CHK %> <%=ignper0Dis %> /> Vis uger<br />
                                        <span class="inline"><input type="radio" value="1" name="FM_ignper" id="FM_ignper" <%=ignperCHK %> <%=ignperDis %> /> Ignorer periode <input type="text" name="FM_fradato" value="<%=fraDato %>" class="form-control input-small" style="width:80px;" /> - <input type="text" name="FM_tildato" value="<%=tilDato %>" class="form-control input-small" style="width:80px;" /></span><br />
                                        <input type="radio" value="2" name="FM_ignper" id="FM_ignper2" <%=ignper2CHK %> <%=ignper2Dis %> /> Vis løn periode &nbsp;<br />-->

                                    <input type="checkbox" value="1" name="FM_viskunikkegk" <%=viskunikkegkCHK %> <%=viskunDIS %> /> Vis også godkendte og afviste timer</div>
                                    
                                    
                                  
                                </div>
                        
                    
                           
                           <!--
                           <div class="row">
                                <div class="col-lg-4"><br /><h4> <a href="godkend_request_timer.asp?varTjDatoUS_man=<%=varTjDatoUS_man_minus %>&FM_visning=<%=visning %>"><</a>  Week <%=datePart("ww", varTjDatoUS_man, 2,2) %> <a href="godkend_request_timer.asp?varTjDatoUS_man=<%=varTjDatoUS_man_plus %>&FM_visning=<%=visning %>">></a></h4> </div>
                                    
                           </div>-->

                  
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
                                 sqlTfaktim = " AND (tfaktim = 7 OR tfaktim = 10 OR tfaktim = 30 OR tfaktim = 51 OR tfaktim = 52 OR tfaktim = 54 OR tfaktim = 50 OR tfaktim = 53 OR tfaktim = 55 OR tfaktim = 60 OR tfaktim = 61 OR tfaktim = 90)"
                                 timeKolTxt = "Timer ønsket"
                            case else
                                 sqlTfaktim = " AND tfaktim = 51 "
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

                                          <%if cint(visning) = 1 then %>
                                          <table width="100%" border="0" class="table table-striped">
                                              <thead>
                                           <%else %>
                                            <table width="100%" border="0">
                                            <%end if %>
                                              <tr>

                                                   <%if cint(visning) = 1 then %>
                                                   <td>Medarbejder</td><td>Init</td>
                                                   <%end if %>
                                                  
                                                  <td colspan="2">Dato</td><td colspan="2">Klokkeslet</td><td align=right>hh:mm</td><td align=right>100 tals</td>

                                                  <%if cint(visning) = 1 then %>
                                                   <td align=right>Timer til løn</td><td align=right>Pause</td><td align=right>Projekttimer</td><td>&nbsp;</td>
                                                   <%end if %>

                                                  <%if visning = 1 then %>
                                                  </thead>
                                                  <%end if %>
                                              </tr>

                                         <%
                                          if cint(visning) = 0 then
                                            logintjkDatoSql = " AND DATE(login) BETWEEN '"& varTjDatoUS_manSQL & "' AND '"& varTjDatoUS_sonSQL &"'"
                                            'logintjkDatoSql = year(oRec5("tdato")) & "/" & month(oRec5("tdato")) &"/"& day(oRec5("tdato")) 

                                            '******* Tjekker logindtid ********'
                                            timerMin100Tot = 0
                                            strSQL = "SELECT id, logud, login, minutter, stempelurindstilling FROM login_historik WHERE mid = "& usemrn &""_
                                            &" AND stempelurindstilling <> -10 "& logintjkDatoSql &" AND logud IS NOT NULL ORDER BY login LIMIT 50"
                                            
                                            else 'viser alle


                                             if usemrn <> 0 then 
                                             midSQL = "l.mid = "& usemrn
                                             else
                                             midSQL = "l.mid <> 0 "
                                             end if

                                            logintjkDatoSql = " AND DATE(l.login) BETWEEN '"& varTjDatoUS_manSQL & "' AND '"& varTjDatoUS_sonSQL &"'"
                                            timerMin100Tot = 0
                                            strSQL = "SELECT id, logud, l.login, minutter, mnavn, init, m.mid, stempelurindstilling FROM login_historik l "_
                                            &" LEFT JOIN medarbejdere m ON (m.mid = l.mid) WHERE "& midSQL &" "_
                                            &" AND stempelurindstilling > 0 "& logintjkDatoSql &" AND logud IS NOT NULL ORDER BY m.mid, l.login LIMIT 300"
                                            

                                            end if

                                            'if session("mid") = 1 then
                                            'Response.write strSQL
                                            'Response.flush
                                            'end if
                                            
                                            l = 0
                                            thisDayC = 0
                                            oRec.open strSQL, oConn, 3
                                            while not oRec.EOF 
                                            
                                            login = oRec("login")
                                            logud = oRec("logud")
                                            
                                            thisDay = datepart("w", login, 2,2)
                                            

                                            timerThis = 0
                                            timerThis = formatnumber(oRec("minutter")/60, 2)
                                            timerThis_komma = instr(timerThis, ",")
                                            timerThis = left(timerThis,  timerThis_komma - 1)

                                            minutterThis = 0
                                            minutterThis = formatnumber(oRec("minutter")/60, 2)
                                            minutterThis = right(minutterThis, 2)
                                            minutterThis = formatnumber((minutterThis*60)/100, 0)
                                          

                                            'minutterThis = formatnumber(minutterThis * 60/100, 0)
                                            minutter100 = formatnumber(minutterThis/60*100, 0) 

                                            if minutterThis < 10 then
                                            minutterThis = "0"& minutterThis
                                            end if

                                            if minutter100 < 10 then
                                            minutter100 = "0"& minutter100
                                            end if

                                            

                                            if cint(visning) = 0 then

                                            'Response.write "<tr><td>"& left(weekdayname(weekday(login)), 3) &".</td><td>"& formatdatetime(login, 2) &"</td><td>"& left(formatdatetime(login, 3), 5) & " - " & left(formatdatetime(logud, 3), 5) & "</td><td style=""width:100px;"">&nbsp;</td><td align=right>"& timerThis &":"& minutterThis &"</td><td align=right>"& timerThis &","& minutter100 &"</td></tr>" 


                                            
                                                if cint(lastDay) <> cint(thisday) AND l > 0 AND thisDayC > 1 then
                                                    call dagsTotal(timerThisDayTot)
                                                end if

                                                if cint(lastDay) <> cint(thisday) AND l > 0 then
                                                timerThisDayTot = 0
                                                thisDayC = 0
                                                end if
                                             

                                             Response.write "<tr>"_
                                             &"<td>"& left(weekdayname(weekday(login)), 3) &".</td>"_
                                             &"<td>"& formatdatetime(login, 2) &"</td>"

                                             if oRec("stempelurindstilling") <> -1 then
                                             Response.write "<td>"& left(formatdatetime(login, 3), 5) & " - " & left(formatdatetime(logud, 3), 5) & "</td>"
                                             else
                                             Response.write "<td>Pause</td>"
                                             end if
                                             
                                             Response.write "<td style=""width:100px;"">&nbsp;</td><td align=right>"& timerThis &":"& minutterThis &"</td>"_
                                             &"<td align=right>"& timerThis &","& minutter100 &"</td>"_
                                             &"</tr>" 

                                            else

                                             Response.write "<tr>"_
                                             &"<td>"& oRec("mnavn")&"</td><td>"& oRec("init")&"</td>"_
                                             &"<td>"& left(weekdayname(weekday(login)), 3) &".</td>"_
                                             &"<td>"& formatdatetime(login, 2) &"</td>"_
                                             &"<td>"& left(formatdatetime(login, 3), 5) & " - " & left(formatdatetime(logud, 3), 5) & "</td>"_
                                             &"<td style=""width:100px;"">&nbsp;</td><td align=right>"& timerThis &":"& minutterThis &"</td>"_
                                             &"<td align=right>"& timerThis &","& minutter100 &"</td>"


                                             datoSQL = year(login) &"/"& month(login) &"/"& day(login)

                                             lonTimerTreg = 0
                                             strSQLtimerLon = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tdato = '"& datoSQL &"' AND tmnr = "& oRec("mid") &" AND tfaktim > 5 AND tfaktim <> 10 GROUP BY tmnr"

                                             'Response.write strSQLtimerLon
                                             'Response.flush
                                             oRec4.open strSQLtimerLon, oConn, 3
                                             if not oRec4.EOF then

                                             lonTimerTreg = oRec4("sumtimer")

                                             end if

                                             oRec4.close

                                             Response.write "<td align=right><a href=""godkend_request_timer.asp?FM_medarb="& oRec("mid") &"&visning=0&varTjDatoUS_man="& varTjDatoUS_man &""" target=""_blank"">"& formatnumber(lonTimerTreg, 2) &"</a></td>"


                                             pauseTimerTreg = 0
                                             strSQLtimerPause = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tdato = '"& datoSQL &"' AND tmnr = "& oRec("mid") &" AND tfaktim = 10 GROUP BY tmnr"

                                             'Response.write strSQLtimerLon
                                             'Response.flush
                                             oRec4.open strSQLtimerPause, oConn, 3
                                             if not oRec4.EOF then

                                             pauseTimerTreg = oRec4("sumtimer")

                                             end if

                                             oRec4.close

                                             Response.write "<td align=right>"& formatnumber(pauseTimerTreg, 2) &"</td>"

                                             projektTimerTreg = 0
                                             strSQLtimerprojekt = "SELECT SUM(timer) AS sumtimer FROM timer WHERE tdato = '"& datoSQL &"' AND tmnr = "& oRec("mid") &" AND (tfaktim = 1 OR tfaktim = 2) GROUP BY tmnr"
                                             oRec4.open strSQLtimerprojekt, oConn, 3
                                             if not oRec4.EOF then

                                             projektTimerTreg = oRec4("sumtimer")

                                             end if
                                             oRec4.close

                                             Response.write "<td align=right>"& formatnumber(projektTimerTreg, 2) &"</td>"

                                             'ALERT HVIS løntimer afviger? 
                                             alert = 0
                                             lonTimerstempel100 = timerThis &","& minutter100
                                             if ((formatnumber(lonTimerstempel100, 2) - formatnumber(lonTimerTreg+pauseTimerTreg, 2)) > "0,25") then
                                             alert = 1
                                             end if


                                             'SKAL DER VISES ALERT HVIS projekttimer afviger? 
                                             if (cdbl(lonTimerstempel100) - formatnumber(projektTimerTreg, 2)) > 1 then
                                             'alert = 1
                                             end if

                                             if cint(alert) = 1 then
                                             Response.write "<td align=right>&nbsp;&nbsp;&nbsp;<span style=color:red>!</span></td>"
                                             else
                                             Response.write "<td align=right>&nbsp;&nbsp;&nbsp;</td>"
                                             end if


                                             Response.write "</tr>" 


                                            end if

                                             
                                             if oRec("stempelurindstilling") <> -1 then
                                             timerMin100this = (timerThis &","& minutter100) * 1
                                             else
                                             timerMin100this = (timerThis &","& minutter100) * -1
                                             end if

                                             timerMin100Tot = timerMin100Tot*1 + (timerMin100this*1)

                                             lastDay = datepart("w", login, 2,2)
                                             
                                             if oRec("stempelurindstilling") <> -1 then
                                             timerThisDayTot = (timerThisDayTot*1)+(oRec("minutter")*1)
                                             else
                                             timerThisDayTot = (timerThisDayTot*1)+(oRec("minutter")*-1)
                                             end if

                                             thisDayC = thisDayC + 1

                                            l = l + 1
                                            oRec.movenext
                                            wend
                                            oRec.close
                                        
                                        if l = 0 then
                                        %>
                                        <tr><td>Ingen loginds i valgte periode</td></tr>
                                            <%else 
                                                
                                                      if visning <> 1 then
                                                
                                                        if thisDayC > 1 then

                                                            call dagsTotal(timerThisDayTot)
                                                   
                                                        end if
                                                
                                                      %>
                                                      <tr>
                                                          <td><b>Total:</b></td>
                                                          <td>&nbsp;</td>
                                                          <td>&nbsp;</td>
                                                          <td>&nbsp;</td>
                                                          <td>&nbsp;</td>
                                                          <td align=right><b><%=formatnumber(timerMin100Tot, 2) %></b></td>
                                                  
                                                      </tr>
                                                       <%    
                                                       end if
                                            
                                         end if %>
                                        </table></div>

                                    

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