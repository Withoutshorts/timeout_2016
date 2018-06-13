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

        nt = 0
        if nt = 1 then

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
            
        end if 'nt
       


      

       

                   


        if cint(touchMonitor) <> 0 then
        %>
             
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

                <%    
                mobil_week_reg_job_dd = 1
                mobil_week_reg_akt_dd = 1
                %>

               <form action="../sesaba.asp" method="post" id="monitorform" name="monitorform">
               <input type="hidden" id="mobil_week_reg_akt_dd" name="" value="<%=mobil_week_reg_akt_dd %>"/>
               <input type="hidden" id="mobil_week_reg_job_dd" name="" value="<%=mobil_week_reg_job_dd %>"/>

                <% 

                

                'if request.Cookies("monitor_job")(session("mid")) <> "" then
                'jobidC = request.Cookies("monitor_job")(session("mid"))
                'aktidC = request.Cookies("monitor_akt")(session("mid"))
                'else
                jobidC = "-1"
                aktidC = "-1"
                'end if

                
                '** Er medarbejder med i Aumatisjon Eller Enginnering
                call medariprogrpFn(usemrn)

                'Response.write "MED I PROGRP: "& medariprogrpTxt

                if instr(medariprogrpTxt, "#14#") = 0 AND instr(medariprogrpTxt, "#16#") = 0 then

                '** IKKE PÅ COOKIE DA DET SKAL GÆLDE ALLE TERMINALER
                strSQLsel = "SELECT lha_jobid FROM login_historik_aktivejob_rel WHERE lha_mid = "& session("mid") & " AND lha_jobid <> 0"

                'response.write strSQLsel
                'response.flush

                oRec.open strSQLsel, oConn, 3
                if not oRec.EOF then

                    jobidC = oRec("lha_jobid")

                end if
                oRec.close
                

                 '** IKKE PÅ COOKIE DA DET SKAL GÆLDE ALLE TERMINALER
                strSQLsel = "SELECT lha_aktid FROM login_historik_aktivejob_rel WHERE lha_mid = "& session("mid") & " AND lha_aktid <> 0"
                oRec.open strSQLsel, oConn, 3
                if not oRec.EOF then

                    aktidC = oRec("lha_aktid")

                end if
                oRec.close

                end if

                'jobidC = "-1"
                'aktidC = "-1"

                'jobidC = "909"
                'aktidC = "-1"    
                    
                %>

               <input type="hidden" id="jq_jobidc" name="" value="<%=jobidC %>"/>
               <input type="hidden" id="jq_aktidc" name="" value="<%=aktidC %>"/>

               <input type="hidden" id="FM_medid" name="FM_medid" value="<%=usemrn %>"/>
               <input type="hidden" id="lto" value="<%=lto%>">
               <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
                 
                    
            

                   <%   
                       
                       if browstype_client <> "ip" then
                        
                        txtClass = "lg"
                        inputHeight = "150%"
                        inputFont = "150%"
                        paddingTop = "30px"
                        txtboxWdt = "" 
                        tblwdt = "450px"
                         paddingTop2 = "15px"
                        paddingTop3 = "40px"

                       else

                        txtClass = "lg"
                        inputHeight = "100%"
                        inputFont = "150%"
                        paddingTop = "10px"
                        'txtboxWdt = "width:300px;"
                        txtboxWdt = "" 
                        tblwdt = "300px"
                        paddingTop2 = "5px"
                        paddingTop3 = "10px"

                       end if

                        rdir_timereg = "touchMonitor" 
                    %>

                   <table style="font-size:80%; width:<%=tblwdt%>px; display:inline-table;" border="0">

                       <%if instr(medariprogrpTxt, "#14#") = 0 AND instr(medariprogrpTxt, "#16#") = 0 then  %>

                     <tr><td colspan="2" style="padding-left:10px; padding-top:0px; font-size:15px; text-align:left;">Projekt</td></tr>
                   <tr><td colspan="2">

                       <!--<span style="background-color:yellow; font-size:14px;">Hei <%=session("user") %> <br />- pga en nulstilling af projekter, er du nødt til at vælge projekt på nyt.</span>-->

                        <!--<input type="text" id="test" />-->
                            <%
                                
                                     

                          
                              if cint(mobil_week_reg_job_dd) = 1 then %>
                            
                            <input type="hidden" id="FM_job_0" value="-1"/>
                             <select id="dv_job_0" name="FM_jobid" style="font-size:<%=inputFont %>; height:<%=inputHeight%>; <%=txtboxWdt%>" class="form-control input-<%=txtClass %> chbox_job">
                                 <option value="-1"><%=left(tsa_txt_145, 4) %>..</option>
                                 <!--<option value="0">..</option>-->
                             </select>

                            <%else %>
                            <input type="text" style="font-size:<%=inputFont%>; height:<%=inputHeight%>" id="FM_job_0" name="FM_job" placeholder="<%=tsa_txt_066 %>/<%=tsa_txt_236 %>" class="FM_job form-control input-<%=txtClass %>"/>
                           <!-- <div id="dv_job_0" class="dv-closed dv_job" style="border:1px #cccccc solid; padding:10px; visibility:hidden; display:none;"></div>--> <!-- dv_job -->

                             <select id="dv_job_0" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none;">
                                 <option><%=tsa_txt_534%>..</option>
                             </select>

                            <input type="hidden" id="FM_jobid_0" name="FM_jobid" value="0"/>
                            <%end if %>

                       </td></tr>
                        <tr><td colspan="2" style="padding-left:10px; padding-top:10px; font-size:15px; text-align:left;">Komponent/Aktivitet</td></tr>

                            <tr><td colspan="2">

                                <%
                                    
                              mobil_week_reg_akt_dd = 1      
                              if cint(mobil_week_reg_akt_dd) = 1 then %>
                                 <input type="hidden" id="FM_akt_0" value="-1"/>
                                 <!--<textarea id="dv_akt_test"></textarea>-->
                                 <select id="dv_akt_0" name="FM_aktivitetid" class="form-control input-<%=txtClass %> chbox_akt" style="font-size:<%=inputFont%>; height:<%=inputHeight%>; <%=txtboxWdt%>" DISABLED>
                                      <option>..</option>
                                  </select>

                             <%else %>
                              <input style="font-size:<%=inputFont%>; height:<%=inputHeight%>" type="text" id="FM_akt_0" name="activity" placeholder="<%=tsa_txt_068%>" class="FM_akt form-control input-<%=txtClass %>"/>
                                <!--<div id="dv_akt_0" class="dv-closed dv_akt" style="border:1px #cccccc solid; padding:10px; visibility:hidden; display:none;"></div>--> <!-- dv_akt -->

                                  <select id="dv_akt_0" class="form-control input-small chbox_akt" size="10" style="visibility:hidden; display:none;"> <!-- chbox_akt -->
                                      <option><%=tsa_txt_534%>..</option>
                                  </select>
                              <input type="hidden" id="FM_aktid_0" name="FM_aktivitetid" value="0"/>
                             <%end if %>
                           
                        </td></tr>

                    <tr>
                       <td colspan="2" style="padding-top:5px; padding-left:10px; text-align:right;">
                         
                           <input type="button" id="bt_indlaspaajob" class="btn btn-sm btn-info" value="Indlæs timer og vælg etterpå nyt projekt >>">
                           <input type="hidden" value="0" name="indlaspaajob" id="indlaspaajob" />
                       </td>
                   </tr>

                   <%if browstype_client <> "xip" then %>
                   <tr><td colspan="2"><br />&nbsp;</td></tr>
                   <%end if %>


                    <%else %>
                       <input type="hidden" id="FM_jobid_0" name="FM_jobid" value="0"/>
                       <input type="hidden" id="FM_aktid_0" name="FM_aktivitetid" value="0"/>
                    <%end if 'instr(medariprogrpTxt, "#14#") = 0 AND instr(medariprogrpTxt, "#16#") = 0%>



                   <tr>
                        <td style="padding-top:<%=paddingTop2%>; padding-left:10px; font-size:15px; text-align:left;">Komme/gå tid:</td>
                        <td style="padding-top:<%=paddingTop2%>; padding-left:20px; font-size:15px; text-align:right;">
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
                                
                                'response.write "timerLoggetPaa100:"&  timerLoggetPaa100
                                'response.Flush

                                timerLoggetPaa = (timerLoggetPaa100)
                               
                                komma = instr(timerLoggetPaa, ",") 'efter formatnumber er der altid komma
                                if komma > 0 then
                                timerLoggetPaaTimer = left(timerLoggetPaa, komma-1)
                                else
                                timerLoggetPaaTimer = "00"
                                end if
                                
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
                    
                   <!--
                    <tr>
                        <td style="padding-top:5px; padding-left:10px; font-size:15px;">Overtid idag (beordret):<br /><span style="font-size:11px;">30 min = 0,5</span>
                        <td style="padding-top:5px; padding-left:60px; font-size:15px; text-align:right;" align="right"><input type="text" id="fm_overtid" name="FM_timer_overtid" value="<%=overTidberegnet %>" class="form-control input-lg" style="text-align:right;" /></td>
                    </tr>
                    <tr>
                        
                        <td style="font-size:15px; text-align:right;" colspan="2"><input type="checkbox" name="fm_overtid_ok" value="1" /> Godkend overtid</td>
                    </tr>
                   -->

                    <tr>
                        <td style="padding-top:<%=paddingTop2%>; padding-left:10px; font-size:15px;">Evt. overtid ønskes udbetalt:
                        <td style="padding-top:<%=paddingTop2%>; padding-left:60px; font-size:15px; text-align:right;" align="right"><input type="text" id="fm_overtidonskesudbetalt" name="fm_overtidonskesudbetalt" value="0" class="form-control input-lg" style="text-align:right;" /></td>
                    </tr>

                    <tr>
                        <td style="padding-top:<%=paddingTop2%>; padding-left:10px; font-size:15px;">Reisetid:<br /><span style="font-size:11px;">30 min = 0,5</span>
                        <td style="padding-top:<%=paddingTop2%>; padding-left:60px; font-size:15px; text-align:right;" align="right"><input type="text" id="fm_rejsetid" name="FM_rejsetid" value="0" class="form-control input-lg" style="text-align:right;" /></td>
                    </tr>
                   


                   <tr>
                        <td style="padding-top:<%=paddingTop3%>; padding-left:10px; font-size:15px; vertical-align:top;">Arbejde ute:</td>
                        <td style="padding-top:<%=paddingTop3%>; padding-left:65px; font-size:15px;"><input type="checkbox" name="FM_arbute_no" value="1" /> NO<br />
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
                       <td colspan="2" style="padding-top:20px; padding-left:10px; text-align:right;">
                           <%if cdbl(realtimerIdag) > -1 then %>
                           <input type="submit" class="btn btn-lg btn-danger" id="sb_monitor_stempleud" value="Stemple ut >>">
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