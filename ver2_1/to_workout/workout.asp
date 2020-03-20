


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->




<%if Request.Form("AjaxUpdateField") = "true" then
Select Case Request.Form("control")

case "FN_hentakt"

        if len(trim(request("theme"))) <> 0 then
		theme = request("theme")
        else
        theme = 0
        end if
    
         strAktListe = "<option value=0>Choose Activity</option>"
         a = 0

                strSQL = "SELECT for_aktid, a.id AS aktid, a.navn As aktnavn, a.fase FROM fomr_rel "_
                &" LEFT JOIN aktiviteter a ON (a.id = for_aktid) WHERE for_fomr = "& theme & " AND job = 4 AND a.navn IS NOT NULL"
                oRec.open strSQL, oConn, 3
                while not oRec.EOF

                    if len(trim(oRec("fase"))) <> 0 then
                    strFase = " ("& oRec("fase") & ")"
                    else
                    strFase = ""
                    end if
                    
                    if isNull(oRec("aktnavn")) <> true then
                    strAktListe = strAktListe & "<option value="& oRec("aktid") &">" & oRec("aktnavn") & strFase &"</option>" 
                    end if

                a = a + 1
                oRec.movenext
                wend
                oRec.close

                if cint(a) = 0 then
                      strAktListe = strAktListe & "<option value=0>None</option>"
                  
                end if


                 '*** ÆØÅ **'
                call jq_format(strAktListe)
                strAktListe = jq_formatTxt
                response.write strAktListe 
                response.end

end select
Response.end
end if


public timerIalt
sub totalprdaglinje

    timerIalt = timerIalt * 60/100 * 100
    %>
    <tr>

            <%if browstype_client <> "ip" then 'Coach %>
            <td>&nbsp;</td>
            <%end if %>


            <%if level = 1 AND browstype_client <> "ip" then 'conducted %>
            <td>&nbsp;</td>
            <%end if %>

                <%if browstype_client <> "ip" then 'theme %>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
            <%end if %>

            <td colspan="2">&nbsp;</td>

            <td><b><%=formatnumber(timerIalt, 0) %></b> min.</td>

            <%if browstype_client <> "ip" then 'slet + submit %>
            <td colspan="2"></td>
        <%else %>
        <td>&nbsp;</td>
        <%end if %>
        </tr>

<%

    timerIalt = 0
end sub
    
    

sub divGrafTheme
    



    timerThemeWHt = 0
    
    
    if timerThemePl <> 0 OR timerThemeCon then
    %>
    <div class="progress">
    <%
    end if


    if timerThemePl <> 0 then

    timerThemeWHt = formatnumber(timerThemePl * 5, 0)

    %>
    
        <div style="width:<%=timerThemeWHt%>%;" class="progress-bar progress-bar-info"><%=formatnumber(timerThemePl, 0)%> t.</div>
   
    <%
    end if


    if timerThemeCon <> 0 then

    timerThemeWHt = formatnumber(timerThemeCon * 5, 0)
    %>
    <div id="dvcon_<%=t%><%=f%>" class="progress-bar <%=themeCol%> dvcon" style="width:<%=timerThemeWHt%>%;"><%=formatnumber(timerThemeCon, 0)%> t.</div>

    <%
    end if


    if timerThemePl <> 0 OR timerThemeCon then
    %>
    </div>
    <%end if



    if cint(periode) = 0 then

        dsp = "none"
        wzb = "hidden"
        thisid = "dvakt_"&t&f
        thisclass = "dvakt"
        bdr = "1px #999999 solid"
        

    else

        if timerThemeCon <> 0 then
        dsp = ""
        wzb = "visible"
        else
        dsp = "none"
        wzb = "hidden"
        end if
        thisid = ""
        thisclass = ""
        bdr = "0px"
    end if

    %>
        <div id="<%=thisid%>" class="<%=thisclass%>" style="position:relative; top:2px; font-size:12px; padding:5px 5px 5px 20px; border:<%=bdr%>; display:<%=dsp%>; visibility:<%=wzb%>;">
            <b>Conducted exercises:</b><br />
            <%=taktivitetnavnStrAf%>
            <br /><br />&nbsp;

        </div>
    <%  


end sub
    
    
    
%>



<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<script src="js/workout.js"></script>




        <%
            call browsertype()
           
            if len(session("user")) = 0 then
            errortype = 5
            call showError(errortype)
            response.End
            end if

            

            func = request("func")


            if len(trim(request("FM_aar"))) <> 0 then
                aar = request("FM_aar")  
            else
                aar = year(now) 
            
                if month(now) <= 6 then
                aar = aar - 1
                end if

            end if


            if len(trim(request("FM_datoFra"))) <> 0 then
                fradato = request("FM_datoFra")  
            else
                fradato = day(now) &"-"& month(now) &"-"& year(now) 
            end if

            
            if len(trim(request("FM_progrpid"))) <> 0 then
            progrpid = request("FM_progrpid")
            else
            progrpid = "-1"
            end if


            if len(trim(request("FM_periode"))) <> 0 then
             periode = request("FM_periode")
            else
             periode = 0
            end if



            select case func
            
            case "dbopr"

            planlagt_afholdt = request("FM_planlagt_afholdt")

            progrp = request("FM_progrp")
            medid = request("FM_medid")
            fomr = request("FM_fomr")
            aktivitet = request("FM_akt")

               


            if cint(aktivitet) = 0 Or cint(fomr) = 0 then

            %>
            <div style="position:absolute; top:100px; left:200px; padding:20px; width:400px; background-color:#FFFFFF; border:10px solid red;">
                <b>Error!</b>
                <br /><br />
                You are missing either Theme or Activity<br /><br />
                <a href="Javascript:history.back()"><< Back</a>
            </div>
            <%
            Response.end
            end if

                aktnavn = "-"
                 strSQLa = "SELECT navn FROM aktiviteter WHERE id = " &  aktivitet 
                oRec.open strSQLa, oConn, 3
                if not oRec.EOF then

                aktnavn = replace(oRec("navn"), "'", "")

                end if
                oRec.close

               

            timerthis = request("FM_timer")
            timerthis = (timerthis * 1 * planlagt_afholdt * 1)
            timerthis = replace(timerthis, ",", ".")

            dato = request("FM_dato")
            datoSQL = year(dato) & "-" & month(dato) & "-" & day(dato)
            dddato = year(now) & "/" & month(now) & "/" & day(now)

            timerkom = replace(request("FM_timerkom"), "'", "")

            team = progrpid

                        strSQLm = "SELECT mnavn FROM medarbejdere WHERE mid = "& medid
                        oRec.open strSQLm, oCOnn, 3
                        if not oRec.EOF then 

                               mnavn = replace(oRec("mnavn"), "'", "")

                        end if
                        oRec.close

            strSQLins = "INSERT INTO timer (tmnr, tmnavn, tdato, tastedato, tjobnavn, tjobnr, tknr, tknavn, timer, taktivitetid, taktivitetnavn, offentlig, timerkom, team) VALUES "_
            &" ("& medid &", '"& mnavn &"', '"& datoSQL &"', '"& dddato &"', 'Træning', 3, 1, 'A27', "& timerthis &", "& aktivitet &",  '"& aktnavn &"', 1, '"& timerkom &"', "& team &")"
            
            'response.write strSQLins
            'response.flush
            
            oConn.execute(strSQLins)


                        if cint(planlagt_afholdt) <> -1 then
                        '**Indsætter på hver spiller 
                        
                        strSQLp = "SELECT medarbejderid, mnavn FROM progrupperelationer "_
                        &" LEFT JOIN medarbejdere ON (mid = medarbejderid) "_
                        &" WHERE projektgruppeid = "& progrp & " AND mansat <> 0 AND mid = 0"
                        oRec.open strSQLp, oCOnn, 3
                        while not oRec.EOF 

                                mnavn = replace(oRec("mnavn"), "'", "")
                            
                                strSQLins = "INSERT INTO timer (tmnr, tmnavn, tdato, tastedato, tjobnavn, tjobnr, tknr, tknavn, timer, taktivitetid, taktivitetnavn, team) VALUES "_
                                &" ("& oRec("medarbejderid") &", '"& mnavn &"', '"& datoSQL &"', '"& dddato &"', 'Træning', 3, 1, 'A27', "& timerthis &", "& aktivitet &",  '"& aktnavn &"', "& team &")"

                                oConn.execute(strSQLins)

                        oRec.movenext
                        wend
                        oRec.close

                        end if           


            response.Redirect "workout.asp?FM_progrpid="&progrpid&"&FM_aar="&aar&"&FM_periode="&periode

            
            case "slet"

                tid = request("tid")
               

                strSQL = "DELETE FROM timer WHERE tid = "& tid
                oConn.execute(strSQL)

                response.Redirect "workout.asp?FM_progrpid="&progrpid&"&FM_aar="&aar&"&FM_periode="&periode
            
            
           
        %>

        <%case else 
            
            if browstype_client <> "ip" then 
                call menu_2014 
            end if%>
        


        <script src="js/datepicker.js" type="text/javascript"></script>
        

    <div class="wrapper">
        <div class="content">

            <div class="container" style="width:90%;">
            <div class="portlet">
                <h3 class="portlet-title"><u>Workout <%=ucase(lto) %></u></h3>
                <div class="portlet-body">

                    <form action="workout.asp" method="post" name="hold">
                    <div class="well well-white">
                         <div class="row">
                     <div class="col-lg-3">
                         
                           Choose Team:

                            <%
                            if level = 1 then
                                progrpKri = " id <> 10"
                            else
                                io = 1
                                medid = session("mid")
                                prgid = 0
                                progrpKri = " (id = 0 "
                                call fTeamleder(medid, prgid, io)

                                strPrgidsArr = split(strPrgids, ",")

                                for p = 0 to UBOUND(strPrgidsArr)

                                   progrpKri = progrpKri & " OR id = "& strPrgidsArr(p)

                                next

                                progrpKri = progrpKri & ")"

                            end if
                                

                                
                            strSQLh = "SELECT id, navn FROM projektgrupper WHERE " & progrpKri
                            
                            'Response.write strSQLh
                            'Response.flush

                           %>
                          <select class="form-control input-small" name="FM_progrpid" onchange="submit()">
                            <%
                                
                           oRec.open strSQLh, oConn, 3
                            while not oRec.EOF 
                                
                                if cint(progrpid) = cint(oRec("id")) then
                                progrpidSEL = "SELECTED"
                                else
                                progrpidSEL = ""
                                end if

                                %><option value="<%=oRec("id") %>" <%=progrpidSEL %>><%=oRec("navn")%></option>
                              <%
                            oRec.movenext
                            wend 
                            oRec.close 
                                %>

                              </select>
                            
                          
                                </div>
                             <div class="col-lg-2">Season:<br />
                                 <select name="FM_aar" class="form-control input-small" onChange="submit()">

                                     <%for a = -5 to 5
                                         
                                         
                                         yThis = year(now)+a
                                         

                                         if cint(aar) = cint(yThis) then
                                         ySel = "SELECTED"
                                         else
                                         ySel = ""
                                         end if
                                         %>

                                     <option value="<%=yThis %>" <%=ySel%>> <%=yThis %> - <%=yThis+1 %></option>
                                     <%next %>


                                 </select>
                                 </div>
                                  
                             

                          <div class="col-lg-1">Period:<br />
                                 <select name="FM_periode" class="form-control input-small" onChange="submit()">

                                     <%
                                         periodeSel0 = ""
                                         periodeSel1 = ""
                                         periodeSel2 = ""
                                         periodeSel3 = ""
                                         periodeSel4 = ""
                                         periodeSel5 = ""
                                         periodeSel6 = ""
                                         periodeSel7 = ""
                                         periodeSel8 = ""
                                         periodeSel9 = ""
                                         periodeSel10 = ""
                                         periodeSel11 = ""
                                         periodeSel12 = ""

                                         select case cint(periode) 
                                         case 0
                                         periodeSel0 = "SELECTED"
                                         case 1
                                         periodeSel1 = "SELECTED" 
                                         case 2
                                         periodeSel2 = "SELECTED"
                                         case 3
                                         periodeSel3 = "SELECTED"
                                         case 4
                                         periodeSel4 = "SELECTED"
                                         case 5
                                         periodeSel5 = "SELECTED"
                                         case 6
                                         periodeSel6 = "SELECTED"
                                         case 7
                                         periodeSel7 = "SELECTED"
                                         case 8
                                         periodeSel8 = "SELECTED"
                                         case 9
                                         periodeSel9 = "SELECTED"
                                         case 10
                                         periodeSel10 = "SELECTED"
                                         case 11
                                         periodeSel11 = "SELECTED"
                                         case 12
                                         periodeSel12 = "SELECTED"
                                         end select
                                         %>

                                     <option value="0" <%=periodeSel0 %>>Year</option>
                                     <option value="1" <%=periodeSel1 %>>Jan</option>
                                     <option value="2" <%=periodeSel2 %>>Feb</option>
                                     <option value="3" <%=periodeSel3 %>>Mar</option>
                                     <option value="4" <%=periodeSel4 %>>Apr</option>
                                     <option value="5" <%=periodeSel5 %>>May</option>
                                     <option value="6" <%=periodeSel6 %>>Jun</option>
                                     <option value="7" <%=periodeSel7 %>>Jul</option>
                                     <option value="8" <%=periodeSel8 %>>Aug</option>
                                     <option value="9" <%=periodeSel9 %>>Sep</option>
                                     <option value="10" <%=periodeSel10 %>>Okt</option>
                                     <option value="11" <%=periodeSel11 %>>Nov</option>
                                     <option value="12" <%=periodeSel12 %>>Dec</option>
                                  
                                 </select>
                                 </div>
                                  <div class="col-lg-6">
                                      <img src="img/A27-2019_logo_136.png" style="float:right" />
                                      </div>
                                  </div>

                        </div>

                        </form>
                    <br /><br />

                    <%if progrpid <> "-1" then %>


                    <div class="panel-group accordion-panel" id="accordion-paneled">
                        <div class="panel panel-default">

                        <%   strSQLh = "SELECT id, navn FROM projektgrupper WHERE id = " & progrpid


                                'Response.write "strSQLh. " & strSQLh

                             oRec.open strSQLh, oConn, 3
                             if not oRec.EOF then
                             holdnavn = oRec("navn")
                             end if
                             oRec.close
                            %>

                        
                                <div class="panel-heading">
                                    <h4 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" href="#0"><%=holdnavn %> Trup & Curriculum</a></h4>
                                </div>

                                 <div id="0" class="panel-collapse collapse">
                                       <div class="panel-body">

                                           <%select case progrpid
                                            case 11
                                            curriname = "CurriculumU13.png"
                                            case 2
                                            curriname = "CurriculumU17.png"
                                            case 3
                                            curriname = "CurriculumU15.png"
                                            case 4
                                            curriname = "CurriculumU14.png"
                                            end select

                                    %><div class="row">

                                        <div class="col-lg-12"><img src="img/<%=curriname %>" border="0" /></div>

                                      </div>
                                       <div class="row">

                                        <div class="col-lg-12">
                                            <h4>Trup</h4>
                                           
                                           <%

                                          strSQLh = "SELECT mnavn, mid FROM progrupperelationer p "_
                                             &" LEFT JOIN medarbejdere ON (mid = p.medarbejderid) WHERE " & progrpid & " AND medarbejdertype = 10 GROUP BY mid ORDER BY mnavn"
                                             oRec.open strSQLh, oConn, 3
                                             while not oRec.EOF 
                                           
                                              %>
                                            
                                             <input type="checkbox" value="<%=oRec("mid") %>" checked /> <%=oRec("mnavn") %><br />
                                           
                                              <%

                                             oRec.movenext
                                             wend
                                             oRec.close %>
                                            </div>
                                           </div><!-- row -->
                                        </div>
                                  </div>

                            </div>
                        </div>
                          
                    <%  
                        
                        
                        if browstype_client <> "ip" then 

                            if cint(periode) = 0 then 'year
                            endpoint = 12
                            else
                            endpoint = 5
                            end if


                        else
                        endpoint = 1
                        end if
                        
                     %>

                    
                    <div class="well well-white">
                       <table class="table table-striped table-bordered">



                             <%  if browstype_client <> "ip" then 
                                 
                                  if cint(periode) = 0 then 'year%>
                                    <thead>
                          
                                        <tr>
                                
                                            <th style="width:8%;">July</th>
                                            <th style="width:8%;">August</th>
                                            <th style="width:8%;">September</th>
                                            <th style="width:8%;">October</th>
                                            <th style="width:8%;">November</th>
                                            <th style="width:8%;">December</th>
                                            <th style="width:8%;">January</th>
                                            <th style="width:8%;">February</th>
                                            <th style="width:8%;">March</th>
                                            <th style="width:8%;">April</th>
                                            <th style="width:8%;">May</th>
                                            <th style="width:8%;">June</th>
                               
                                
                                
                                        </tr>
                                    </thead>


                                    <%else %>

                                     <thead>
                          
                                        <tr>
                                            <th colspan="5"><%=monthname(periode) %></th>
                                        </tr>
                                        <tr>
                                        <%for t = 1 to endpoint 
                                           

                                           
                                            tMonth = periode
                                    
                                            if tMonth < 7 then
                                            tYear = aar+1
                                            else
                                            tYear = aar
                                            end if

                                            if t = 1 then
                                            tdagSt = t
                                            else
                                            tdagSt = (t-1) * 7
                                            end if

                                           thisDato = tdagSt & "/"& tMonth &"/"&tYear

                                           
                                            %>

                                  

                                            <th style="width:20%;">Week <%=datePart("ww", thisDato, 2,2) %> - <%=formatdatetime(thisDato, 2)%></th>
                                           
                                            <%next %>
                                        </tr>

                                     </thead>

                                    <%end if %>

                           <%end if %>

                        <tbody>
                             <tr>
                                 <%for t = 1 to endpoint %>
                                <td>
                                    
                                    <div class="progress">
                                        <div class="progress-bar progress-bar-info" style="width:50%;">Planned</div>
                                        <div class="progress-bar progress-bar-success" style="width:50%;">Conducted</div>
                                    </div>
                                    
                                </td>
                               
                                 <%next %>
                               
                            </tr>
                            <tr>
                                <%for t = 1 to endpoint 
                                    
                                   
                                    bgcol = "#FFFFFF"

                                    if periode = 0 then
                                    
                                        if t < 7 then
                                        tMonth = t + 6
                                        tYear = aar
                                        else
                                        tMonth = t - 6
                                        tYear = aar+1
                                        end if

                                        tdagSt = 1
                                        tdagEnd = 31

                                    else

                                    tMonth = periode
                                    
                                    if tMonth < 7 then
                                    tYear = aar+1
                                    else
                                    tYear = aar
                                    end if

                                    if t = 1 then
                                    tdagSt = t
                                    else
                                    tdagSt = (t-1) * 7
                                    end if

                                    tdagEnd = tdagSt + 7


                                        if tdagEnd > 28 then ' sidste uge i md
                                            select case tMonth
                                            case 1,3,5,7,8,10,12
                                            tdagEnd = 31
                                            case 2
                                                select case aar 
                                                case 2012, 2016, 2020, 2024, 2028, 2032, 2036, 2040, 2044
                                                tdagEnd = 29
                                                case else
                                                tdagEnd = 28
                                                end select
                                            case else
                                            tdagEnd = 30
                                            end select

                                        end if

                                    end if

                                    %>
                                <td style="background-color:<%=bgcol%>;">
                                    <%
                                        lastFomrId = 0
                                        taktivitetnavnStr = ""
                                        f = 0
                                        strSQL = "SELECT for_aktid, f.navn, for_fomr, business_area_label FROM fomr f "_
                                        &" LEFT JOIN fomr_rel ON (for_fomr = f.id ) WHERE id <> 0 AND for_aktid IS NOT NULL ORDER BY for_fomr"
                                        oRec.open strSQL, oConn, 3
                                        while not oRec.EOF

                                            'if lastT <> t then
                                            ' lastFomrId = 0
                                            'end if
                                            if lastFomrId <> oRec("for_fomr") then
                                            
                                            if f <> 0 then
                                            call divGrafTheme
                                            taktivitetnavnStrPl = ""
                                            end if

                                            timerThemePl = 0
                                            end if
                                            strSQLt = "SELECT SUM(timer*-1) AS timer, taktivitetnavn FROM timer WHERE tdato BETWEEN '"& tYear &"-"&tMonth&"-"& tdagSt &"' AND '"& tYear &"-"&tMonth&"-"& tdagEnd &"' AND taktivitetid = "& oRec("for_aktid") &" AND timer < 0 AND offentlig = 1 AND team = "& progrpid &" GROUP BY taktivitetid" 
                                            
                                            'Response.write strSQLt & "<br>"
                                            oRec2.open strSQLt, oConn, 3
                                            while not oRec2.EOF

                                                %>
                                                <!-- <%=oRec2("taktivitetnavn") %>: <%=oRec2("timer") %> t.<br /> -->
                                                <%
                                                    taktivitetnavnStrPl = taktivitetnavnStrPl &", " & oRec2("taktivitetnavn")
                                                    timerThemePl = (timerThemePl*1) + replace(oRec2("timer"), ".", ",")

                                            oRec2.movenext
                                            wend
                                            oRec2.close

                                              themeCol = oRec("business_area_label")

                                            
                                            if lastFomrId <> oRec("for_fomr") then
                                            timerThemeCon = 0
                                            taktivitetnavnStrAf = ""
                                            end if
                                            strSQLt = "SELECT SUM(timer) AS timer, taktivitetnavn FROM timer WHERE tdato BETWEEN '"& tYear &"-"&tMonth&"-"& tdagSt &"' AND '"& tYear &"-"&tMonth&"-"& tdagEnd &"' AND taktivitetid = "& oRec("for_aktid") &" AND timer > 0 AND offentlig = 1 AND team = "& progrpid &" GROUP BY taktivitetid" 
                                            
                                            'Response.write strSQLt & "<br>"
                                            oRec2.open strSQLt, oConn, 3
                                            while not oRec2.EOF

                                                   
                                                    taktivitetnavnStrAf = taktivitetnavnStrAf & oRec2("taktivitetnavn") & " ("& oRec2("timer") &" t. / "& formatnumber((oRec2("timer")/1.5),0) &" pas.)<br>"
                                                    timerThemeCon = (timerThemeCon * 1) + replace(oRec2("timer"), ".", ",") 'oRec2("timer") '

                                            oRec2.movenext
                                            wend
                                            oRec2.close


                                                    'response.write lastFomrId &"<>"& oRec("for_fomr") &" ANd " & timerThemePl & "<> 0 OR "& timerThemeCon & "<> 0" 

                                                    if lastFomrId <> oRec("for_fomr") ANd (timerThemePl <> 0 OR timerThemeCon <> 0) then 
                                                    %><b><%=oRec("navn") %></b><br /> <%
                                                    lastFomrId = oRec("for_fomr")
                                                    end if

                                                    


                                        f = f + 1
                                        
                                        
                                        oRec.movenext
                                        wend
                                        oRec.close

                                        call divGrafTheme


                                        '**** Intensitet '***********

                                                      
                                            
                                        
                                            timerIntensitetPl = 0
                                            
                                            strSQLt = "SELECT SUM(timer) AS timer, taktivitetnavn, a.fase FROM timer "_
                                            &" LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE tdato BETWEEN '"& tYear &"-"&tMonth&"-"& tdagSt &"' AND '"& tYear &"-"&tMonth&"-"& tdagEnd &"' "_
                                            &" AND timer > 0 AND offentlig = 1 AND fase <> 'ingen' AND team = "& progrpid &" GROUP BY fase ORDER BY fase" 
                                            i = 0
                                            'Response.write strSQLt & "<br>"
                                            oRec2.open strSQLt, oConn, 3
                                            while not oRec2.EOF


                                                    timerIntensitetPl = replace(oRec2("timer"), ".", ",")
                                                    timerThemeWHt = formatnumber(timerIntensitetPl * 1, 0)

                                                    select case lcase(left(oRec2("fase"), 3))
                                                    case "ing"
                                                    bgCol = ""
                                                    case "lav"
                                                    bgCol = "info"
                                                    case "mid"
                                                    bgCol = "warning"
                                                    case "høj"
                                                    bgCol = "danger"
                                                    end select
                                                    
                                                    if i = 0 then
                                                %>

                                                   <br />
                                        
                                              <b>Intensity:</b>
                                              <div class="progress">
                                                <%  end if
                                                    %>
                                                    
                                                       <!--
                                                    <div style="background-color:<%=bgCol%>; height:<%=timerThemeWHt%>px; width:40px; font-size:8px; padding:1px; border:1px #999999 solid;"><%=formatnumber(timerIntensitetPl, 0)%></div>
                                                        -->
                                                

                                                        <div class="progress-bar progress-bar-<%=bgCol %>" style="width:<%=timerThemeWHt%>%;"><%=left(oRec2("fase"), 3) %>: <%=timerIntensitetPl %></div>
                                      
                                                        <%
                                               i = i + 1    
                                            oRec2.movenext
                                            wend
                                            oRec2.close

                             

                                        if i <> 0 then%>  
                                    </div>
                                    <%end if %>

                                </td>
                               


                                <%
                                    lastT = t
                                    next %>
                            </tr>
                            </tbody>
                           </table>
                       

                    <br /><br />
                    
                    <form action="workout.asp?func=dbopr" method="post" name="indlas">
                          <input type="hidden" name="FM_progrp" value="<%=progrpid%>" />
                          <input type="hidden" name="FM_progrpid" value="<%=progrpid%>" />
                          <input type="hidden" name="FM_aar" value="<%=aar%>" />
                        

                    <h3>Add activity</h3>
                    <a href="mailto:haase.a27@gmail.com?subject=A27 Mangler aktivitet;">A27 Mangler aktivitet >></a> 
                    <table class="table table-striped table-bordered">
                        <thead>
                            <tr>
                                <%if browstype_client <> "ip" then %>
                                <th>Coach</th>
                                <%end if %>

                                <%if level = 1 then %>
                                <th>Planned/Conducted</th>
                                <%end if %>

                                <th>Date</th>
                                <th>Theme</th>
                                <th>Activity</th>
                                <th>Comment</th>
                                <th>Duration</th>
                                <th>&nbsp;</th>
                                <th>&nbsp;</th>
                            </tr>
                        </thead>

                        <tbody>

                            
                                <tr>
                                    <%if browstype_client <> "ip" then %>
                                    <td>
                                        <select class="form-control input-small" name="FM_medid">
                                            <%strSQLm = "SELECT mid, mnavn, init FROM medarbejdere WHERE mansat = 1 AND medarbejdertype = 9 ORDER BY mnavn"
                                                
                                                oRec.open strSQLm, oConn, 3
                                                While not oRec.EOF 

                                                if cint(session("mid")) = cint(oRec("mid")) then
                                                mSEL = "SELECTED"
                                                else
                                                mSel = ""
                                                end if

                                                %>    
                                                   <option value="<%=oRec("mid")%>" <%=mSel %>><%=oRec("mnavn") %></option>
                                                <%
                                                oRec.movenext
                                                Wend 
                                                oRec.close
                                                %>
                                            
                                            
                                     
                                       </select>
                                    </td>
                                    <%else %>
                                    <input type="hidden" name="FM_medid" value="<%=session("mid") %>">
                                    <%end if %>


                                    <%if level = 1 then %>
                                    <td><!-- progruppe -->
                                        <select class="form-control input-small" name="FM_planlagt_afholdt">
                                            <option value="1" SELECTED>Conducted</option>
                                           <option value="-1">Planned</option>
                                          
                                            

                                        </select>
                                        <!--<input type="text" value=".." id="test" />-->
                                    </td>
                                    <%
                                    else
                                        %>
                                    <input type="hidden" name="FM_planlagt_afholdt" value="1" />
                                    <%
                                    end if %>


                                    <td>
                                        <div class='input-group date'>
                                            <input type="text" style="width:100%;" class="form-control input-small" name="FM_dato" value="<%=fradato %>" placeholder="dd-mm-yyyy" />
                                            <span class="input-group-addon input-small"><span class="fa fa-calendar"></span></span>
                                        </div>
                                    </td>
                                   
                                      <td><!-- fomr -->
                                        <select class="form-control input-small" name="FM_fomr" id="theme">
                                            <option value="0">Choose Theme..</option>
                                           <%strSQLp = "SELECT id, navn FROM fomr WHERE id <> 0 ORDER BY navn"
                                                
                                                oRec.open strSQLp, oConn, 3
                                                While not oRec.EOF 
                                                %>    
                                                   <option value="<%=oRec("id")%>"><%=oRec("navn") %></option>
                                                <%
                                                oRec.movenext
                                                Wend 
                                                oRec.close
                                                %>
                                        </select>
                                   </td>
                                     <td><!-- Aktiviteter -->

                                        <select class="form-control input-small" name="FM_akt" id="FM_activity" style="width:200px;">
                                            <option value="0">Choose Theme first..</option>
                                        </select>
                                   </td>
                                  
                                  <td><input type="text" style="width:100%;" class="form-control input-small" name="FM_timerkom" value="" placeholder="Kommentar" /></td>
                                     <td>
                                        <select class="form-control input-small" name="FM_timer">
                                            <option value="0,17">10 min.</option>
                                            <option value="0,25">15 min.</option>
                                            <option value="0,33">20 min.</option>
                                            <option value="0,5">30 min.</option>
                                            <option value="0,75">45 min.</option>
                                            <option value="1">1 time</option>
                                            <option value="1,5">1,5 time</option>
                                            <%if level = 1 then %>
                                            
                                            <option value="3">10 min. pr. træning (3 timer)</option>
                                            <option value="4">15 min. pr. træning (4 timer)</option>
                                            <option value="5">20 min. pr. træning (5 timer)</option>
                                            <option value="8">30 min. pr. træning (8 timer)</option>
                                            <option value="12">45 min. pr. træning (12 timer)</option>
                                            <option value="19">1 Uge (9 timer)</option>
                                            <option value="18">2 Uger (18 timer)</option>
                                            <option value="36">1 md. (36 timer)</option>
                                            <%end if %>

                                        </select>
                                   </td>
                                    <td style="text-align:center"></td>

                                    <!-- Opret knap -->
                                    <td style="text-align:center"><button type="submit" class="btn btn-default btn-sm"><b>Indlæs</b></button>
                                    </td>

                                </tr>
                             
                            <%if browstype_client = "ip" then %>
                            </table>

                            <h3>Latest activities</h3>
                            <table class="table table-striped table-bordered">
                            <%else %>
                                <tr><td colspan="20">
                                    <br /><br />
                                       <b>Latest activities</b>

                                    </td></tr>
                            <%end if %>

                            <%
                                timerIalt = 0
                                if browstype_client <> "ip" then
                                    if periode = 0 then
                                    fradato = tYear & "/"& month(now) &"/1"
                                    else
                                    fradato = tYear & "/" & periode & "/1"
                                    end if

                                else
                                fradato = dateadd("d", -7, fradato)
                                end if

                                datoSQL = year(fradato) & "-" & month(fradato) & "-" & day(fradato)

                                x = 0
                                lastDato = "01-01-2002"

                                if level = 1 then
                                strSQLplancon = ""
                                else
                                strSQLplancon = " AND timer > 0"
                                end if

                                if browstype_client <> "ip" then

                                    if periode = 0 then
                                    tildato = dateadd("m", 3, fradato)
                                    else
                                    tildato = dateadd("m", 1, fradato)
                                    end if
                                else
                                tildato = dateadd("d", 14, fradato)
                                end if

                                datoSQL_slut = year(tildato) & "-" & month(tildato) & "-" & day(tildato)

                                strSQLt = "SELECT tid, tmnavn, taktivitetnavn, tdato, timer, timerkom FROM timer WHERE team = "& progrpid &" "& strSQLplancon &" AND tdato BETWEEN '"& datoSQL &"' AND '"& datoSQL_slut &"' ORDER BY tdato DESC LIMIT 300 "
                                'response.Write strSQLt
                                oRec.open strSQLt, oConn, 3
                                while not oRec.EOF
                                
                                if lastDato <> oRec("tdato") AND x <> 0 then
                                %>
                                        <%call totalprdaglinje %>
                                <%
                                end if
                                
                                x = x + 1
                            %>
                                <tr>
                                    <%if browstype_client <> "ip" then %>
                                    <td><%=oRec("tmnavn") %></td>
                                    <%end if %>
                                    
                                    <%if browstype_client <> "ip" then %>
                                    <td>
                                        <%if oRec("timer") > 0 then %>
                                        Conducted
                                        <%else %>
                                        Planned
                                        <%end if %>
                                    </td>
                                    <%end if %>

                                    <td><%=left(weekdayname(weekday(oRec("tdato"))), 3) %>. 

                                        <%if browstype_client = "ip" then %>
                                        <br />
                                        <%end if %>
                                        <%=formatdatetime(oRec("tdato"), 1) %></td>

                                     <%if browstype_client <> "ip" then %>
                                    <td></td>
                                    <%end if %>

                                    <td>
                                         <%if browstype_client = "ip" then %>
                                        <%if oRec("timer") < 0 then %>
                                        Planned:<br />
                                        <%end if %>
                                        <%end if %>

                                        <%=oRec("taktivitetnavn") %></td>
                                    <%if browstype_client <> "ip" then %>
                                    <td><%=oRec("timerkom") %></td>
                                    <%end if %>
                                    <td><%=formatnumber(oRec("timer") * 60/100 * 100, 0) %> min.</td>
                                    <td><a href="workout.asp?func=slet&tid=<%=oRec("tid") %>&FM_progrpid=<%=progrpid%>&FM_aar=<%=aar %>" style="color:red;">[X]</a></td>

                                     <%if browstype_client <> "ip" then 'slet + submit %>
                                    <td>&nbsp;</td>
                                    <%end if%>
                                </tr>
                            <%


                                lastDato = oRec("tdato")
                                timerIalt = timerIalt*1 + oRec("timer")*1

                                oRec.movenext
                                wend
                                oRec.close
                            %>


                            <%if x = 0 then %>

                                <tr>
                                    <td colspan="20" style="text-align:center;">No records</td>
                                </tr>

                            <%
                            else
                               %>

                                   <%call totalprdaglinje %>

                            <%end if %>

                        </tbody>

                    </table>
                         </form>

                    </div>

                    <%end if' progrpid <> -1 %>

            </div>
        </div>
            </div></div>
</div>

        <%end select 'func %>













<!--#include file="../inc/regular/footer_inc.asp"-->