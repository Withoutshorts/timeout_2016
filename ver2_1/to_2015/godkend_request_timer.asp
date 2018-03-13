


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->





       
        
        <%
           
            function tr_gk_request_timer(dothis, medid)

                                    '** dothis = 0 'Standard
                                    '** dothis = 1 'Tiembank Cflow - ekstra tom linje
         
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


                                     <%if session("stempelur") <> 0 then 
                                         
                                         logintjkDatoSql = year(oRec5("tdato")) & "/" & month(oRec5("tdato")) &"/"& day(oRec5("tdato")) 

                                         %>
                                        <td valign="top"><%
                                           '******* Tjekker logindtid ********'
                                            strSQL = "SELECT id, logud, login FROM login_historik WHERE mid = "& medid &" AND stempelurindstilling > 0 AND DATE(login) = '"& logintjkDatoSql &"' AND logud IS NOT NULL ORDER BY login LIMIT 10"
                                            
                                            'Response.write strSQL
                                            'Response.flush
                                            
                                            oRec.open strSQL, oConn, 3
                                            while not oRec.EOF 
                                            
                                            login = oRec("login")
                                            logud = oRec("logud")
                                            
                                            Response.write left(formatdatetime(login, 3), 5) & " - " & left(formatdatetime(logud, 3), 5) & "<br>" 

                                            oRec.movenext
                                            wend
                                            oRec.close
                                        
                                        %></td>

                                     <%end if %>


                                    <td>
                                         <select name="FM_aktids" class="form-control input-small">
                                        <%
                                        'call akttyper2009(2)    
                                        '"& replace(aty_sql_realhours, "tfaktim", "fakturerbar") &" AND     
                                        strSQLakttype = "SELECT a.id, a.navn FROM job j LEFT JOIN aktiviteter a ON (a.job = j.id) WHERE (fakturerbar <> 1 AND fakturerbar <> 2)  AND j.risiko < 0 AND jobnr = 3 "
                                        oRec.open strSQLakttype, oConn, 3
                                        while not oRec.EOF
                                        
                                            
                                            if cdbl(oRec5("taktivitetid")) = cdbl(oRec("id")) then
                                            selAkt = "SELECTED"
                                            else
                                            selAkt = ""
                                            end if%>


                                       
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
                                         
                                       <select name="opdater_gkstatus_<%=oRec5("tid")%>" class="form-control input-small">
                                           <option value="0" <%=godkendtstatusSEL0 %>>None</option>
                                           <option value="1" <%=godkendtstatusSEL1 %>>Godkendt</option>
                                           <option value="2" <%=godkendtstatusSEL2 %>>Afvist</option>
                                           <option value="3" <%=godkendtstatusSEL3 %>>Tentative</option>
                                       </select></td>
                                      
                                     <td><input type="checkbox" name="opdater_gk_<%=oRec5("tid")%>" value="1" class="gk" /></td>
                                     
                                     
                                     
                                     <%else 
                                     '*** Til TimebanK CFLOW%>



                                        <input type="hidden" name="tids7" value="0" />
                                        <input type="hidden" name="tids7_tmnr" value="<%=medid %>" />
                                         <td colspan="4">&nbsp;</td>
                                         <td>
                                         <select name="FM_aktids7" class="form-control input-small">
                                        <%
                                       strSQLakttype = "SELECT a.id, a.navn FROM job j LEFT JOIN aktiviteter a ON (a.fakturerbar = 7 AND a.job = j.id) WHERE (fakturerbar = 7) AND j.risiko < 0 AND jobnr = 3 "
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
		                        
		                 

                    strSQlupd = "UPDATE timer SET "& tFakTimSQL &" "& aktNavnSQL &" "& tjobkundeSQL &" timer = "& timerThis &", godkendtstatus = 1, godkendtstatusaf = '"& session("user") &"', godkendtdato = '"& dtNow &"' WHERE tid = "& tids(t)
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
                                   
                                   

                    t = 0
                    for t = 0 to UBOUND(tids7)

                    'Response.write "    HER: "& t & " "& request("timer_gk_7_"& tids7_tmnr(t)) &  "#"& request("opdater_gk_7_"& tids7_tmnr(t))  &"<br>"
                    'Response.end
            
                    if request("timer_gk_7_"& tids7_tmnr(t)) > 0 AND len(trim(request("timer_gk_7_"& tids7_tmnr(t)) )) <> 0 AND request("opdater_gk_7_"& tids7_tmnr(t)) = "1" then


                        'Response.write "<br>    HER 999: "& aktids7(t)
                        timerThis = request("timer_gk_7_"& tids7_tmnr(t))
                        call indlasTimerTfaktimAktid(lto, tids7_tmnr(t), timerThis, 7, 0)
                    
                    end if

                    next

                                    'Response.end
                    %>


   <div class="container" style="width:100%; height:100%">
            <div class="portlet">
                   
                  <h3 class="portlet-title"><u>Godkend ønsket Ferie, Tillæg, Flex og Overtid</u> </h3>
                  

                <div class="portlet-body">
                     

                       
                        <div class="row" style="padding-left:100px;">
                            <br /><br />
                               <div class="col-lg-12">
                    
                            <h3>Timer opdateret!</h3>
                   
                                   
                            <a href="Javascript:window.close()">Luk faneblad</a><br /><br />
                            <a href="godkend_request_timer.asp">Godkend flere timer</a>

                            </div>
                            </div>
                            </div>
                        </div>
               </div>

                    <%
                    Response.end
                    
            
            
                    '******* Tjekker medarbejderen ********'
                

                    '******* Tjekker logindtid ********'
                    'strSQL = "SELECT id, logud FROM login_historik WHERE mid = "& medid &" AND stempelurindstilling > 0 ORDER BY id DESC"
                    'oRec.open strSQL, oConn, 3
                    'if not oRec.EOF then
                    'logud = oRec("logud")
                    'id = oRec("id")
                    'end if
                    'oRec.close
                   

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


        if len(trim(request("varTjDatoUS_man"))) <> 0 then

            varTjDatoUS_man = request("varTjDatoUS_man")

            mandag = datePart("w", varTjDatoUS_man, 2,2) 

            if cint(mandag) > 1 then
            varTjDatoUS_man = dateAdd("d", -(mandag - (mandag-1)), ddNu)
            end if

            else

            ddNu = now
            mandag = datePart("w", ddNu, 2,2) 

            if cint(mandag) > 1 then
            varTjDatoUS_man = dateAdd("d", -(mandag - (mandag-1)), ddNu)
            end if

        end if

        varTjDatoUS_man_minus = dateAdd("d", -7, varTjDatoUS_man)
        varTjDatoUS_man_plus = dateAdd("d", 7, varTjDatoUS_man)

        useStDatoKri = varTjDatoUS_man

        sqlDatoStart = year(useStDatoKri)&"/"&month(useStDatoKri)&"/"&day(useStDatoKri)
        sqlDatoStartATD =  year(now)&"/1/1"

     
        ddNowMinusOne = dateAdd("d", -1, now)

        sqlDatoEnd = year(ddNowMinusOne)&"/"&month(ddNowMinusOne)&"/"&day(ddNowMinusOne)

        daysAnsat = dateDiff("d", sqlDatoStart, ddNowMinusOne)

        if len(trim(request("FM_medarb"))) <> 0 then
        usemrn = request("FM_medarb")
        else
             if request.Cookies("tsa")("selmedreqtimer") <> "" then
             usemrn = request.Cookies("tsa")("selmedreqtimer")
             else
             usemrn = 0
             end if
        end if

        response.Cookies("tsa")("selmedreqtimer") = usemrn
                    
        if cdbl(usemrn) <> 0 then
            usrmnrSQL = " tmnr = "& usemrn &""
        else
            usrmnrSQL = " tmnr <> 0"
        end if


      

        %>
        <script src="js/godkend_request_timer_jav.js" type="text/javascript"></script>
        <div class="container" style="width:100%; height:100%">
            <div class="portlet">
                   
                  <h3 class="portlet-title"><u>Godkend ønsket Ferie, Tillæg, Flex og Overtid</u> </h3>
                  

                <div class="portlet-body" style="padding-left:100px;">
                     <!--<div class="portlet-title">
                          <table style="width:100%;">
                              <tr>
                                  <td><h3 style="color:black"><%=medarb_navn %></h3></td>
                                  <td style="text-align:right"><img style="width:15%; margin-bottom:2px; margin-left:55px" src="img/cflow_logo.png" /></td>
                              </tr>
                          </table>
                        
                        </div>-->
                       <section>
                       <div class="well well-white" style="width:800px;">
                      <form action="godkend_request_timer.asp?varTjDatoUS_man=<%=varTjDatoUS_man%>"" method="post">
                        
                                <div class="row">
                                      <div class="col-lg-1">&nbsp</div>
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

                                    <div class="col-lg-1"><br /><button type="submit" class="btn btn-sm"><b><%=medarb_txt_100 %> >></b></button></div>
                                  
                                </div>
                        
                    
                           </div>
                           <div class="row">
                                <div class="col-lg-1"><br /> <a href="godkend_request_timer.asp?varTjDatoUS_man=<%=varTjDatoUS_man_minus %>"><<</a>  Uge: <b><%=datePart("ww", varTjDatoUS_man, 2,2) %></b>   <a href="godkend_request_timer.asp?varTjDatoUS_man=<%=varTjDatoUS_man_plus %>">>></a> </div>
                                    
                           </div>

                  
                          </form>
                       </div>
                </section>
                <br /><br /><br />&nbsp;


                        <%  
                            
                            select case lto
                            case "wap"
                                 sqlTfaktim = " AND tfaktim = 51 "
                                 timeKolTxt = "AFTALE-home ønsket"
                            case "cflow"
                                 sqlTfaktim = " AND (tfaktim = 30 OR tfaktim = 51 OR tfaktim = 52 OR tfaktim = 54 OR tfaktim = 50 OR tfaktim = 53 OR tfaktim = 55 OR tfaktim = 60 OR tfaktim = 61 OR tfaktim = 90)"
                                 timeKolTxt = "Timer ønsket"
                            case else
                                 sqlTfaktim = " AND tfaktim = 51 "
                                 timeKolTxt = "Timer ønsket"
                            end select 
                            
                            cspan = 9%>
                       
                        <form action="godkend_request_timer.asp?func=gk" method="post">
                            
                        <div class="row" style="padding-left:100px;">
                             <div class="col-lg-12">
                            <!--Flexsaldo: <b><%=flexsaldo %></b><br />-->
                            <table style="width:60%" class="table table-striped">
                                <thead>
                                    <tr>
                                    <th>Navn</th>
                                    <th>Dato</th>
                                    <th>Norm.</th>
                                    <%if session("stempelur") <> 0 then
                                       cspan = cspan + 1%>
                                    <th>Komme/gå tid</th>
                                    <%end if %>
                                    <th>Aktivitet</th>
                                    <th><%=timeKolTxt %></th>
                                    <th>+ Faktor</th>
                                    <th>Godkend antal<br /> <span style="background-color:yellow; padding:4px;">(0: slet)</span></th>
                                        <th>Status</th>
                                    <th>Opdater <br /><input type="checkbox" id="gkall" />
                                        
                                    </th>
                                    </tr>
                                </thead>
                                
                             <%
                         

                            'overtidtot = 0
                            m = 0
                            strSQL6 = "SELECT tmnr, tmnavn, timer As overtidtimer_onsket, tdato, tid, tfaktim, taktivitetnavn, taktivitetid, godkendtstatus FROM timer WHERE "& usrmnrSQL &" "& sqlTfaktim &" AND (godkendtstatus = 0 OR godkendtstatus = 3) ORDER BY tmnavn, tdato" '"& usemrn &" AND tdato BETWEEN '"& sqlDatoStartATD &"' AND '"& sqlDatoEnd &"'
                            oRec5.open strSQL6, oConn, 3
                            while not oRec5.EOF


                                    if cdbl(lasttmnr) <> cdbl(oRec5("tmnr")) AND m > 0 then   
            
                                        call tr_gk_request_timer(1, oRec5("tmnr"))


                                 %>
                                 <tr><td colspan="<%=cspan %>"><br />&nbsp;</td></tr>
                                 <%
                                 end if

                                call tr_gk_request_timer(0, oRec5("tmnr"))
                            
        

                               
                                lasttmnr = oRec5("tmnr")

                                m = m + 1

                            oRec5.movenext
                            wend 
                            oRec5.close 

                            medid = lasttmnr
                                 
                              call tr_gk_request_timer(1, medid)%>

                            <tr><td colspan="<%=cspan %>" align="right"> <input type="submit" class="btn btn-small btn-success" value="Godkend/Afvis >>"></td></tr>

                            </table><br />

                                

                          </div>
                        </div>
                       
                           </form>

                                        </div>
            </div>
        </div>
<%end select %>

<!--#include file="../inc/regular/footer_inc.asp"-->