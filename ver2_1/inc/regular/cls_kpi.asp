<!------ Flex saldo ----->

    <script type ="text/javascript">
        $(document).ready(function () {

            
           /* $("#bt_overtid").click(function () {

                alert("HER" + $("#sp_overtid").html())
               timerThis = $("#sp_overtid").html()
               $("#fm_overtid").val(""+timerThis+"")

            });
            */

        });
    

    </script>

    <%
           
        
    function keyfigures(touchMonitor, login)

        
        if session("mid") = 1 AND lto = "hestia" then
        usemrn = 84
        else
        usemrn = session("mid")
        end if

        call akttyper2009(2)
        call meStamdata(usemrn)
        call licensStartDato()
        licensstdato = licensstdato
        meAnsatDato = meAnsatDato

        yNow = year(now)

        if cDate(meAnsatDato) > cDate(licensstdato) then
        useStDatoKri = meAnsatDato 
        else
        useStDatoKri = licensstdato
        end if

        call lonKorsel_lukketPerPrDato(now)
        if lonKorsel_lukketPerDt > useStDatoKri then
        useStDatoKri = dateAdd("d", 1, day(lonKorsel_lukketPerDt) &"/"& month(lonKorsel_lukketPerDt) &"/"& year(lonKorsel_lukketPerDt))
        else
        useStDatoKri = useStDatoKri
        end if


        sqlDatoStart = year(useStDatoKri)&"/"&month(useStDatoKri)&"/"&day(useStDatoKri)
        sqlDatoStartATD =  year(now)&"/1/1"

        '** Er dd i år 
        if month(now) <= 4 then 
        sqlDatoStartFerie =  year(now)-1&"/5/1"
        sqlDatoEndFerie =  year(now)&"/4/30"
        else
        sqlDatoStartFerie =  year(now)&"/5/1"
        sqlDatoEndFerie =  year(now)+1&"/4/30"
        end if

        ddNowMinusOne = dateAdd("d", -1, now)

        sqlDatoEnd = year(ddNowMinusOne)&"/"&month(ddNowMinusOne)&"/"&day(ddNowMinusOne)

        daysAnsat = dateDiff("d", sqlDatoStart, ddNowMinusOne)

        ntimPerIdag = 0
        call normtimerPer(usemrn, useStDatoKri, 0, 0)
        ntimPerIdag = ntimPer '/3600
                    
        ntimPer = 0
        realtimerIalt = 0
        strSQL2 = "SELECT tmnavn, SUM(timer) AS realtimer FROM timer WHERE tmnr = "& usemrn &" AND ("& aty_sql_realhours &" OR tfaktim = 114) AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoEnd &"' GROUP BY tmnr"

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

        ntimerPeriode = dateDiff("d", useStDatoKri, daysansat, 2,2)
                              
        call normtimerPer(usemrn, useStDatoKri, ntimerPeriode-1, 0)
        ntimPerFlex = ntimPer

      
        'balRealNormTxt = formatnumber(( realTimer(x) + korrektionReal(x) ) - normTimer(x),2)

        flexsaldo = (realtimerIalt - ntimPerFlex)

        'call nortimerStandardDag(meType)
            
        %>

        <!------ Ferie saldo ----->
        <%
                    
        ferieafholdt = 0
        call ferieAfholdtPer(sqlDatoStartFerie, sqlDatoEndFerie, usemrn, 0) 

                    
        'strSQL3 = "SELECT SUM(timer) AS ferieafholdt FROM timer WHERE tmnr = "& usemrn &" AND tdato BETWEEN '"& sqlDatoStartFerie &"' AND '"& sqlDatoEndFerie &"' AND tfaktim = 14 GROUP BY tmnr"
        'oRec3.open strSQL3, oConn, 3
        'if not oRec3.EOF then
        'ferieafholdt = oRec3("ferieafholdt")
        'end if
        'oRec3.close
                   

        if ferieAFPerTimer(0) <> 0 then
            ferieafholdt = ferieAFPerTimer(0)
        else
            ferieafholdt = 0
        end if

        'response.Write "<br> af:" & ferieAFPerTimer(0)
                    
        ferieoptjent = 0
        strSQL4 = "SELECT timer AS ferieoptjent, tdato FROM timer WHERE tmnr = "& usemrn &" AND tdato BETWEEN '"& sqlDatoStartFerie &"' AND '"& sqlDatoEndFerie &"' AND (tfaktim = 15 OR tfaktim = 111 OR tfaktim = 112) ORDER BY tdato"
                    
        'if session("mid") = 1 then
        'Response.write strSQL4
        'Response.flush
        'end if
                    
        oRec8.open strSQL4, oConn, 3
        while not oRec8.EOF 
                    
                call normtimerPer(usemrn, oRec8("tdato"), 6, 0)
	            if ntimPer <> 0 then
                normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
                else
                normTimerGns5 = 1
                end if 

               
        ferieoptjent = ferieoptjent + (oRec8("ferieoptjent")*1/normTimerGns5)
                    
        oRec8.movenext
        wend
        oRec8.close
                   
                
        ferieSaldo = (ferieoptjent - ferieafholdt)

            'response.Write "<br> tal: " & ferieoptjent
        %>

        <!------ sygetal ----->
        <%

        sygtot = 0
        strSQL5 = "SELECT timer AS sygtimer, tdato FROM timer WHERE tmnr = "& usemrn &" AND tfaktim = 20 AND tdato BETWEEN '"& sqlDatoStartATD &"' AND '"& sqlDatoEnd &"' ORDER BY tdato"
                    
        oRec8.open strSQL5, oConn, 3
        while not oRec8.EOF 

                call normtimerPer(usemrn, oRec8("tdato"), 0, 0)
	            if ntimPer <> 0 then
                ntimPerUse = ntimPer
                else
                ntimPerUse = 1
                end if 

        sygtot = sygtot + (oRec8("sygtimer")*1 / ntimPerUse)
        oRec8.movenext
        wend 
        oRec8.close
                    

        barnsygtot = 0
        strSQL6 = "SELECT timer As barnsygtimer, tdato FROM timer WHERE tmnr = "&usemrn&" AND tfaktim = 21 AND tdato BETWEEN '"& sqlDatoStartATD &"' AND '"& sqlDatoEnd &"' ORDER BY tdato"
        oRec8.open strSQL6, oConn, 3
        while not oRec8.EOF

                call normtimerPer(usemrn, oRec8("tdato"), 0, 0)
	            if ntimPer <> 0 then
                ntimPerUse = ntimPer
                else
                ntimPerUse = 1
                end if 

        barnsygtot = barnsygtot + (oRec8("barnsygtimer")*1/ntimPerUse)
        oRec8.movenext
        wend 
        oRec8.close
                   

        if cint(touchMonitor) = 1 then

        overtidtot = 0
        strSQL6 = "SELECT SUM(timer) As overtidtimer, tdato FROM timer WHERE tmnr = "&usemrn&" AND tfaktim = 30 GROUP BY tmnr" 'AND tdato BETWEEN '"& sqlDatoStartATD &"' AND '"& sqlDatoEnd &"'
        oRec8.open strSQL6, oConn, 3
        if not oRec8.EOF then

        overtidtot = overtidtot + oRec8("overtidtimer")*1
        
        end if 
        oRec8.close

        realtimerIdag = 0
        sqlDatoIdag = year(now) &"/"& month(now) &"/"& day(now)
        strSQL2 = "SELECT tmnavn, SUM(timer) AS realtimer FROM timer WHERE tmnr = "& usemrn &" AND ("& aty_sql_realhours &") AND tdato = '"& sqlDatoIdag &"' GROUP BY tmnr"


        'response.write strSQL2
        'response.flush
       
        oRec2.open strSQL2, oConn, 3
        if not oRec2.EOF then
        realtimerIdag = oRec2("realtimer")
        end if

        oRec2.close

        end if

                   

        if ferieSaldo < 0 then
        flexferieCol = "red"
        else
        flexferieCol = "greenyellow"
        end if

        if flexsaldo < 0 then
        flexsaldoCol = "red"
        else
        flexsaldoCol = "greenyellow"
        end if

        if sygtot < 0 then
        sygtotCol = "red"
        else
        sygtotCol = "greenyellow"
        end if

        if barnsygtot < 0 then
        barnsygtotCol = "red"
        else
        barnsygtotCol = "greenyellow"
        end if

        if cint(touchMonitor) <> 0 then
        %>
             <!--<div class="col-lg-3"> -->
                <!-- monitor.asp?func=logud&medid=<%=medid %>&medarb_navn=<%=medarb_navn %> -->
                <style type="text/css">
         
                .blink {
  animation: blink-animation 1s steps(5, start) infinite;
  -webkit-animation: blink-animation 1s steps(5, start) infinite;
}
@keyframes blink-animation {
  to {
    visibility: hidden;
  }
}
@-webkit-keyframes blink-animation {
  to {
    visibility: hidden;
  }
}
                    </style>

               <form action="../sesaba.asp" method="post">
               <table style="font-size:80%; width:450px; display:inline-table;" border="0">  
                      <tr>
                        <td style="padding-top:5px; padding-left:10px; font-size:15px; text-align:left; width:250px;">Flex saldo (opgjort pr. igår):</td>
                        <td style="padding-top:5px; padding-left:20px; font-size:15px; text-align:right;"><%=formatnumber(flexsaldo, 2) %></td>
                    </tr> 
                   <tr>
                        <td style="padding-top:5px; padding-left:10px; font-size:15px; text-align:left;">Overtid saldo (godkendte):</td>
                        <td style="padding-top:5px; padding-left:20px; font-size:15px; text-align:right;"><%=formatnumber(overtidtot, 2) %></td>
                    </tr> 
                   <tr>
                        <td style="padding-top:5px; padding-left:10px; font-size:15px; text-align:left;">Komponent/aktivitet idag:</td>
                        <td style="padding-top:5px; padding-left:20px; font-size:15px; text-align:right;"><%=formatnumber(realtimerIdag, 2)%></td>
                    </tr>
                   <tr><td colspan="2"><br /><br />&nbsp;</td></tr>

                   <tr>
                        <td style="padding-top:20px; padding-left:10px; font-size:15px; text-align:left;">Komme/gå tid:</td>
                        <td style="padding-top:20px; padding-left:20px; font-size:15px; text-align:right;">
                            <%
                                
                                
                            timerLoggetPaa = dateDiff("n", login, now, 2,2)
                            timerLoggetPaa100 = formatnumber((timerLoggetPaa*100/60) / 100, 2)
                            overTidberegnet = formatnumber((ntimPerIdag) - timerLoggetPaa100, 2)
                            
                            if cdbl(overTidberegnet) < 0 then
                                overTidberegnet = overTidberegnet
                            else
                                overTidberegnet = 0
                            end if

                            'timerogminutberegning(totalTimerPer100) 

                            'if cdbl(timerLoggetPaa100) > 1 then
                                
                                timerLoggetPaa = (timerLoggetPaa100)
                               
                                komma = instr(timerLoggetPaa, ",") 'efter formatnumber er der altid komma
                                timerLoggetPaaTimer = left(timerLoggetPaa, komma-1)
                                timerLoggetPaaMin = right(timerLoggetPaa, 2)
                                timerLoggetPaaMin = formatnumber((timerLoggetPaaMin / 100) * 60, 0)

                                timerLoggetPaaMin = left(timerLoggetPaaMin,2)

                                if timerLoggetPaaMin < 10 then
                                timerLoggetPaaMin = "0"& timerLoggetPaaMin
                                end if
                            
                                timerLoggetPaa = timerLoggetPaaTimer& ":"& timerLoggetPaaMin

                            %>

                            <%=left(formatdatetime(login, 3),5) %> - <%=left(formatdatetime(now, 3), 5) %> = <span class="blink"><b><%=timerLoggetPaa%></b></span></td>
                    </tr>
                    
                   
                    <tr>
                        <td style="padding-top:5px; padding-left:10px; font-size:15px;">Overtid idag (beordret):<br /><span style="font-size:11px;">30 min = 0,5</span>
                        <td style="padding-top:5px; padding-left:60px; font-size:15px; text-align:right;" align="right"><input type="number" id="fm_overtid" name="FM_timer_overtid" value="<%=overTidberegnet %>" class="form-control input-lg" style="text-align:right;" /></td>
                    </tr>
                    <tr>
                        
                        <td style="font-size:15px; text-align:right;" colspan="2"><input type="checkbox" name="fm_overtid_ok" value="1" /> Godkend overtid</td>
                    </tr>
                   <tr>
                        <td style="padding-top:40px; padding-left:10px; font-size:15px; vertical-align:top;">Arbejde ute:</td>
                        <td style="padding-top:40px; padding-left:65px; font-size:15px;"><input type="checkbox" name="FM_arbute_no" value="1" /> NO<br />
                            <input type="checkbox" name="FM_arbute_world" value="1" /> Utland<br />
                            <input type="checkbox" name="FM_arbute_teamleder" value="1" /> Teamleder</td>
                    </tr>

                   <!--
                    <tr>
                        <td style="padding-top:20px; padding-left:10px; font-size:15px; text-align:left;">Ferie saldo</td>
                        <td style="padding-top:20px; padding-left:20px; font-size:15px; text-align:right;"><%=formatnumber(ferieSaldo,2) %></td>
                    </tr>
                 
                    <tr>
                        <td style="padding-top:23px; padding-left:10px; font-size:15px; text-align:left;">Syg ÅTD</td>
                        <td style="padding-top:23px; padding-left:20px; font-size:15px; text-align:right;"><%=formatnumber(sygtot, 2) %></td>
                    </tr>
                    <tr>
                        <td style="padding-top:27px; padding-left:10px; font-size:15px; text-align:left;">Barn syg ÅTD</td>
                        <td style="padding-top:27px; padding-left:20px; font-size:15px; text-align:right;"><%=formatnumber(barnsygtot, 2) %></td>
                    </tr>
                   -->
                   
                   
                   <tr>
                       <td colspan="2" style="padding-top:40px; padding-left:10px; text-align:right;">
                           <%if cdbl(realtimerIdag) > -1 then %>
                           <input type="submit" class="btn btn-lg btn-danger" value="Stemple ut >>">
                           <%else %>
                           <a href="#" class="btn btn-lg btn-warning"><b>Stemple ut >> </b></a><br />
                           <span style="color:#999999; font-size:11px;">
                           Du behøver mindst 4 timer på Komponent/aktivitet for at kunne logge ud. 
                               </span>
                           <%end if %>

                           

                       </td>
                   </tr>
                </table>
                   </form>
            <!--</div> -->
        <%
        end if


    end function
%>