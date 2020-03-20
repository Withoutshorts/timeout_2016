<%
    
    
    
function TimeregPanel(jobid, io)
    

    if len(trim(request("FM_datoer"))) <> 0 then 'redirect fra timereg_akt_2006 HUSK Dato
        tregDato = day(request("FM_datoer")) &"/"& month(request("FM_datoer")) &"/"& year(request("FM_datoer"))

        ddDato_ugedag = day(tregDato) &"/"& month(tregDato) &"/"& year(tregDato)
        ddDato_ugedag_w = datepart("w", ddDato_ugedag, 2, 2)
            
        varTjDatoUS_man = dateAdd("d", -(ddDato_ugedag_w-1), ddDato_ugedag)
    else      
        tregDato = day(now) &"/"& month(now) &"/"& year(now)
    end if

    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = session("mid")
    end if

    call erStempelurOn() 

    call positiv_aktivering_akt_fn()

    call mobil_week_reg_dd_fn()

    call akt_maksbudget_treg_fn()

    call aktBudgettjkOn_fn()

    'response.Write "mobil_week_reg_job_dd " & mobil_week_reg_job_dd
    'response.Write "stempelurOn " & stempelurOn
    'response.Write "lto " & lto

    if cint(stempelurOn) = 1 then 
        showStartstop = 1

        select case lto
        case "cflow", "intranet - local"
        visTimerelTid = 2
        case else
        visTimerelTid = 0 '0, 1 eller 2 (optinal) altid = 0 da start og slut tid skal være optional
        end select
        'visTimerelTidoptinal = 1 'der kan indtastet klokkeslet eller bare timer. Det er optinal, så der kan nøjes med at blvie indtastet timer
            
    else
        
        visTimerelTid = 0
        showStartstop = 0

    end if


    lcid_sprog_val = Session.LCID


    rdir_timereg = "jobreg" 

    if lto = "dencker" then 'Dette er bare s[ dencker kan teste multi panel
        mobil_week_reg_akt_dd = 1
        showStartstop = 0 'Kun til test
        rdir_timereg = "jobregmp" 'Kun til test de skal tilabge til multi siden
    end if
    
    %>

      <!--  <script src="../to_2015/js/timeregsimple.js"></script> -->

         <SCRIPT src="../to_2015/js/ugeseddel_2011_jav7.js"></script>

        <form id="container" action="../timereg/timereg_akt_2006.asp?func=db&rdir=<%=rdir_timereg %>" method="post">

            <input type="hidden" id="stempelurOn" name="" value="<%=stempelurOn%>"/> <!-- StempelurOn værdig mangler -->
            <input type="hidden" id="varTjDatoUS_man" name="varTjDatoUS_man" value="2018-10-01">
            <input type="hidden" name="usemrn" id="treg_usemrn" value="<%=usemrn%>"> <!-- SKal v;re efter teamleder osv. pr er kun ham der er logget paa -->
            <input type="hidden" id="lto" value="<%=lto%>">
            <input type="hidden" id="Hidden4" name="FM_dager" value="0"/>
            <input type="hidden" id="Hidden2" name="FM_feltnr" value="0"/>
            <input type="hidden" id="FM_pa" name="FM_pa" value="<%=pa_aktlist %>"/>
            <input type="hidden" id="FM_medid" name="FM_medid" value="<%=usemrn %>"/>
            <input type="hidden" id="FM_medid_k" name="FM_medid_k" value="<%=usemrn%>"/>
            <input type="hidden" id="" name="FM_vistimereltid" value="<%=visTimerelTid %>"/>
            <input type="hidden" id="lcid_sprog" name="" value="<%=lcid_sprog_val %>"/>

            <input type="hidden" id="mobil_week_reg_akt_dd" name="" value="<%=mobil_week_reg_akt_dd %>"/>
            <input type="hidden" id="mobil_week_reg_job_dd" name="" value="<%=mobil_week_reg_job_dd %>"/>

            <input type="hidden" id="jq_jobidc" name="" value="-1"/> 
            <input type="hidden" id="jq_aktidc" name="" value="-1"/>


            <!-- Forecast max felter -->
            <input type="hidden" id="aktnotificer_fc" name="" value="0"/>
            <input type="hidden" id="akt_maksbudget_treg" value="<%=akt_maksbudget_treg%>">
            <input type="hidden" id="akt_maksforecast_treg" value="<%=akt_maksforecast_treg%>">
            <input type="hidden" id="aktBudgettjkOn_afgr" value="<%=aktBudgettjkOn_afgr%>">

            <input type="hidden" id="regskabsaarStMd" value="<%=datePart("m", aktBudgettjkOnRegAarSt, 2,2)%>">
            <input type="hidden" id="regskabsaarUseAar" value="<%=datepart("yyyy", varTjDatoUS_man, 2,2)%>">
            <input type="hidden" id="regskabsaarUseMd" value="<%=datepart("m", varTjDatoUS_man, 2,2)%>">
        
            
                <style>
                    #timereg td 
                    {
                        padding-left:5px;
                    }
                </style>

                <table id="timereg" style="width:100%;">
                    <tr>
                    <td>

                        <input type="hidden" id="Hidden5" name="year" value="<%=year(now) %>"/>
                               
                        <%
                        if lto <> "tbg" OR level = 1 then 
                            inputDatoTXT = ""
                            inputDatoCLS = "date"
                        else 
                            inputDatoTXT = "readonly"
                            inputDatoCLS = ""
                            tregDato = day(now) &"-"& month(now) &"-"& year(now)
                        end if
                        %>

                        <div class='input-group <%=inputDatoCLS %>'>
                                <input type="text" autocomplete="off" class="form-control input-small" name="FM_datoer" id="jq_dato" value="<%=tregDato %>" placeholder="dd-mm-yyyy" <%=inputDatoTXT %> />
                                <span class="input-group-addon input-small">
                                <span class="fa fa-calendar">
                                </span>
                            </span>
                        </div>

                    </td>

                    <%if io <> 0 then %>
                    <td>
                        <%if cint(mobil_week_reg_job_dd) = 1 then %>

                          
                             <input type="hidden" id="FM_job_0" value="-1"/>
                             <select id="dv_job_0" name="FM_jobid" class="form-control input-small chbox_job">
                                 <option value="-1"><%=left(tsa_txt_145, 4) %>..</option>
                             </select>                            

                        <%else %>
                            <input type="text" autocomplete="off" id="FM_job_0" name="FM_job" placeholder="<%=tsa_txt_066 %>/<%=tsa_txt_236 %>" class="FM_job form-control input-small"/>
                           <!-- <div id="dv_job_0" class="dv-closed dv_job" style="border:1px #cccccc solid; padding:10px; visibility:hidden; display:none;"></div>--> <!-- dv_job -->

                             <select id="dv_job_0" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none;">
                                 <option><%=tsa_txt_534%>..</option>
                             </select>

                            <input type="hidden" id="FM_jobid_0" name="FM_jobid" value="0"/>
                        <%end if %>
                    </td>
                    <%else %>
                    <input type="hidden" id="dv_job_0" value="<%=jobid %>" />
                    <input type="hidden" id="FM_job_0" value="<%=jobid %>" />
                    <input type="hidden" name="FM_jobid" value="<%=jobid %>" />
                    <%end if %>

                    <td>
                       
                        <%if cint(mobil_week_reg_akt_dd) = 1 then %>
                            <input type="hidden" id="FM_akt_0" value="-1"/>
                            <!--<textarea id="dv_akt_test"></textarea>-->
                            <select id="dv_akt_0" name="FM_aktivitetid" class="form-control input-small chbox_akt" DISABLED>
                                <option>..</option>
                            </select>

                        <%else %>
                            <input type="text" autocomplete="off" id="FM_akt_0" name="activity" placeholder="<%=tsa_txt_068%>" class="FM_akt form-control input-small"/>
                                <!--<div id="dv_akt_0" class="dv-closed dv_akt" style="border:1px #cccccc solid; padding:10px; visibility:hidden; display:none;"></div>--> <!-- dv_akt -->

                            <select id="dv_akt_0" class="form-control input-small chbox_akt" size="10" style="visibility:hidden; display:none;"> <!-- chbox_akt -->
                                <option><%=tsa_txt_534%>..</option>
                            </select>
                            <input type="hidden" id="FM_aktid_0" name="FM_aktivitetid" value="0"/>
                        <%end if %>
                    </td>

                    <td>
                        <%if cint(showStartstop) = 0 OR lto = "outz" then 'or lto = outz kun test til mpt %>
                             <input type="hidden" id="FM_sttid" name="FM_sttid" value="00:00"/>
                             <input type="hidden" id="FM_sltid" name="FM_sltid" value="00:00"/>
                             <input type="text" id="FM_timer" name="FM_timer" placeholder="<%=tsa_txt_137%>" class="form-control input-small" /><!-- brug type number for numerisk tastatur -->
                        <%else %>
                        <%timeNow = left(formatdatetime(now, 3), 5)    %>
                        <table width="100%">
                                 <tr><td colspan="4" style="font-size:9px;">[Optional start and end time]</td></tr>
                                 <tr>
                                     <td style="width:70px;"><input type="text" id="FM_sttid" name="FM_sttid" placeholder="00:00" value="<%=timeNow %>" class="form-control input-small" style="width:60px;" /></td>
                                     <td style="width:70px;"><input type="text" id="FM_sltid" name="FM_sltid" placeholder="00:00" value="<%=timeNow %>" class="form-control input-small" style="width:60px;" /></td>
                                     <td style="width:30px; text-align:center;">=</td>
                                     <td><input type="text" id="FM_timer" name="FM_timer" placeholder="<%=tsa_txt_137%>" class="form-control input-small" style="width:160px; text-align:right;" /><!-- brug type number for numerisk tastatur --></td>
                                 </tr>
                        </table>

                        <%end if %>
                    </td>

                    <td>
                        <input type="text" id="FM_kom" name="FM_kom_0" placeholder="<%=tsa_txt_051%>" class="form-control input-small" />
                    </td>

                    <td>
                        <button id="btn_indlas" class="btn btn-sm btn-success" <%=btn_dis %>><b>>></b></button>
                    </td>
                    </tr>


                </table>

             <!-- <div class="row">
                    <div class="col-lg-5">
                        <input type="text" id="FM_kom" name="FM_kom_0" placeholder="<%=tsa_txt_051%>" class="form-control input-small" />
                    </div>
                    <div class="col-lg-2"><button id="btn_indlas" class="btn btn-sm btn-success" <%=btn_dis %>><b><%=tsa_txt_085%> >></b></button></div>
                </div> -->

        </form>
    <%
    
end function   %>




<%
function JoblogSimple(jobid)


    'Finder jobnr
    strSQL = "SELECT jobnr FROM job WHERE id = "& jobid
    'response.Write strSQL 
    oRec.open strSQL, oConn, 3
    if not oRec.EOF then
    jobnr = oRec("jobnr")
    end if
    oRec.close



%>



    <table id="joblogsimple" class="table dataTable table-striped table-bordered table-hover ui-datatable" style="width:100%;">
        <thead>
            <tr>
                <th>Dato</th>
                <th>Aktivitet</th>
                <th>Medarbejder</th>
                <th>Kommentar</th>
                <th>Timer</th>
            </tr>
        </thead>
        <tbody>

            <%

            if len(trim(jobnr)) <> 0 then
            jobnr = jobnr
            else
            jobnr = "-1"
            end if
              
            strSQL = "SELECT tdato, taktivitetid, a.navn as aktnavn, tmnavn, timerkom, timer, tid, tmnr FROM timer LEFT JOIN aktiviteter a ON (taktivitetid = a.id) WHERE tjobnr = '"& jobnr & "'"
            'response.Write "<br>" & strSQL
            oRec.open strSQL, oConn, 3 
            totalTimer = 0
            while not oRec.EOF 
                totalTimer = totalTimer + oRec("timer")
            %>

            <tr>
                <td><%=oRec("tdato") %></td>
                <td><%=oRec("aktnavn") %></td>
                <td><%=oRec("tmnavn") %></td>
                <td><%=oRec("timerkom") %></td>

                <td style="text-align:right;">
                    <%if cdbl(oRec("tmnr")) = cdbl(session("mid")) OR level = 1 then %>                        
                        <a href="#" onclick="Javascript:window.open('rediger_tastede_dage_2006.asp?id=<%=oRec("tid") %>', '', 'width=450,height=675,resizable=yes,scrollbars=yes')" class=vmenu style="color:<%=gkbgcol%>;"><%=formatnumber(oRec("timer"), 2) %></a>
                    <%else %>
                        <%=formatnumber(oRec("timer"), 2) %>
                    <%end if %>
                </td>
            </tr>

            <%
            oRec.movenext
            wend
            oRec.close
            %>


        </tbody>

        <tfoot>
            <tr>
                <th style="text-align:right;"><%=formatnumber(totalTimer, 2) %></th>
            </tr>
        </tfoot>

    </table>

    <script type="text/javascript">
        //alert("hep")
        var table = $('#joblogsimple').DataTable({
            scrollY: "250px",
            scrollX: true,
            scrollCollapse: true,
            paging: false,
            ordering: false

        });
        
    </script>

<%
end function
%>




<%

function MatPanel(jobid)

%>

        <input type="hidden" name="FM_jobid" value="<%=jobid %>" />
    
        <% 
        matdato = day(now) &"-"& month(now) &"-"& year(now)
        %>

        <style>
            table td 
            {
                padding-left:5px;
            }
        </style>

        <%
            'Henter grundvaluta
            basic_valuta = "DKK"
            valutaid = 1
            strSQL = "SELECT id, valutakode FROM valutaer WHERE grundvaluta = 1"
            oRec.open strSQL, oconn, 3
            if not oRec.EOF then
                valutaid = oRec("id")
                basic_valuta = oRec("valutakode")
            end if
            oRec.close
        %>

        <input type="hidden" id="FM_valutakode" value="<%=basic_valuta %>" />
        <input type="hidden" id="FM_valuta" name="FM_valuta" value="<%=valutaid %>" /> <!-- grundvaluta da man ikke ænadre valuta'en på denne side -->
        <input type="hidden" id="basic_valuta" name="basic_valuta" value="<%=basic_valuta %>" /> 

        <input type="hidden" id="basic_kurs" name="basic_kurs" />
        <input type="hidden" id="basic_belob" name="basic_belob" />

        <table>
            <tr>
                <td>
                    <div class='input-group date'>
                        <input type="text" autocomplete="off" class="form-control input-small" name="FM_matdato" id ="FM_matdato" value="<%=matdato %>" placeholder="dd-mm-yyyy" <%=inputDatoTXT %> />
                        <span class="input-group-addon input-small">
                            <span class="fa fa-calendar"></span>
                        </span>
                    </div>
                </td>

                <td><input type="text" autocomplete="off" name="FM_matnavn" id="FM_matnavn" class="form-control input-small" placeholder="Navn" /></td>

                <td><input type="text" autocomplete="off" name="FM_matantal" id="FM_matantal" class="form-control input-small" placeholder="Antal" /></td>

                <td><input type="text" autocomplete="off" name="FM_matpris" id="FM_matpris" class="form-control input-small" placeholder="Pris" /></td>

                <td>
                    <span id="oprmat" class="btn btn-sm btn-success" <%=btn_dis %>><b>>></b></span>
                </td>
            </tr>
        </table>




<%end function %>


<%function matlistesimple(jobid) %>

    <%
        'Henter grundvaluta
        basic_valuta = "DKK"
        valutaid = 1
        strSQL = "SELECT id, valutakode FROM valutaer WHERE grundvaluta = 1"
        oRec.open strSQL, oconn, 3
        if not oRec.EOF then
            valutaid = oRec("id")
            basic_valuta = oRec("valutakode")
        end if
        oRec.close
    %>


    <%
        'usemrn = session("mid")
        'useSog = 0
        'sogBilagOrJob = 0
        'matreg_personlig = 0
        'matreg_visallemed = 1
        'sogliste = ""

        'call senstematReg(usemrn, useSog, sogBilagOrJob, matreg_personlig, matreg_visallemed, sogliste)
    %>
    
    <table id="matliste" class="table dataTable table-striped table-bordered table-hover ui-datatable" style="width:100%;">
        <thead>
            <tr>
               <!-- <th>Kunde/Job</th> -->
                <th>Medarbejder</th>
                <th>Navn</th>
                <th>Antal</th>
                <th style="white-space:nowrap;">Stk. pris</th>
                <th>Pris</th> <!-- Købspris -->
                <th></th>
            </tr>
        </thead>


        <tbody>

            <%
                sqlWh = ""
                strSQLper = ""
                prisIAlt = 0

                strSQLmat = "SELECT m.mnavn AS medarbejdernavn, mnr, init, mf.id AS mfid, mf.matvarenr AS varenr, mg.navn AS gnavn, mf.matenhed AS enhed, "_
                &" mf.matnavn AS navn, mf.matantal AS antal, mf.dato, mf.editor, "_
                & "mf.matkobspris, mf.matsalgspris, mf.jobid, mf.matgrp, mf.basic_belob as basic_belob, "_
                &" mf.usrid, mf.forbrugsdato, j.id, j.jobnr, j.jobnavn, "_
                &" mg.av, f.fakdato, k.kkundenavn, "_
                &" k.kkundenr, mf.valuta, mf.intkode, mf.bilagsnr, v.valutakode as valutakode, mf.personlig, j.serviceaft, "_
                &" s.navn AS aftnavn, s.aftalenr, mf.kurs, ma.lokation, j.risiko, a.navn AS aktnavn, a.id AS aktid "_
                &" FROM materiale_forbrug mf"_
                &" LEFT JOIN materialer ma ON (ma.id = matid) "_
                &" LEFT JOIN materiale_grp mg ON (mg.id = mf.matgrp) "_
                &" LEFT JOIN medarbejdere m ON (mid = usrid) "_
                &" LEFT JOIN job j ON (j.id = mf.jobid) "_
                &" LEFT JOIN aktiviteter a ON (a.id = mf.aktid) "_
                &" LEFT JOIN serviceaft s ON (s.id = j.serviceaft) "_
                &" LEFT JOIN fakturaer f ON (f.jobid = mf.jobid AND f.faktype = 0) "_
                &" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
                &" LEFT JOIN valutaer v ON (v.id = mf.valuta) "_
                &""& sqlWh &" "& strSQLper &" WHERE mf.jobid = "& jobid  &" GROUP BY mf.id ORDER BY mf.id DESC, mf.forbrugsdato DESC, f.fakdato DESC LIMIT 100"
                'response.Write "strSQLmat " & strSQLmat
                oRec.open strSQLmat, oConn, 3
                while not oRec.EOF
                

                if len(oRec("fakdato")) <> 0 then
                fakdato = oRec("fakdato")
                else
                fakdato = "01/01/2002"
                end if

                prisIAlt = prisIAlt + cdbl(oRec("basic_belob") * oRec("antal"))

                valutaspan = ""
                if oRec("valutakode") <> basic_valuta then
                    valutaspan = "<span style='font-size:9px;'>"& oRec("valutakode") &"</span>"
                end if

            %>

                <tr>
                  <!--  <td><%=oRec("kkundenavn") %> <br /> <%=oRec("jobnavn") %></td> -->
                    <td style="white-space:nowrap;"><span style="font-size:9px;"><%=oRec("forbrugsdato") %></span><br /><%=oRec("medarbejdernavn") %></td>
                    <td><%=oRec("navn") %></td>
                    <td style="text-align:right;"><%=oRec("antal") %></td>
                    <td style="text-align:right; white-space:nowrap;"><%=formatnumber(oRec("matkobspris"), 2) & " " & valutaspan %> </td>
                    <td style="text-align:right; white-space:nowrap;"><%=formatnumber(cdbl(oRec("matkobspris") * oRec("antal")), 2) & " " & valutaspan %> </td>


                    <%if level = 1 then %>
                        <td style="text-align:center;"><a target="_blank" href="../timereg/materialer_indtast.asp?sogliste=<%=oRec("navn") %>&vasallemed=on"><span class="fa fa-pencil"></span></a></td>
                    <%
                    else                        
                        select case lto
                            case "dencker", "jttek"
                            %><td style="text-align:center;"><a target="_blank" href="../timereg/materialer_indtast.asp?sogliste=<%=oRec("navn") %>&vasallemed=on"><span class="fa fa-pencil"></span></a></td><%
                            case else
                                if cdbl(oRec("usrid")) = cdbl(session("mid")) then
                                    %><td style="text-align:center;"><a target="_blank" href="../timereg/materialer_indtast.asp?sogliste=<%=oRec("navn") %>&vasallemed=on"><span class="fa fa-pencil"></span></a></td><%
                                else
                                    %><td style="text-align:center;"></td><%
                                end if
                        end select
                    end if 'level 1
                    %>
                </tr>


            <%
                oRec.movenext
                wend
                oRec.close
            %>

        </tbody>

        <tfoot>
            <tr>
                <th style="text-align:right;" colspan="5"><%=formatnumber(prisIAlt, 2) %> <span style="font-size:7px;"><%=basic_valuta %></span></th>
                <th></th>
            </tr>
        </tfoot>

    </table>

    <script type="text/javascript">
        var table = $('#matliste').DataTable({
            scrollY: "303px",
            scrollX: true,
            scrollCollapse: true,
            paging: false,
            ordering: false

        });
    </script>


<%end function %>






<%function ChatModul (jobid) %>

        <div id="chatscroll" style="height:320px; overflow-y:scroll;">
            <table id="chat" style="width:100%" class="table dataTable table-striped table-bordered table-hover ui-datatable">            
                <tbody>
                <%
                if len(trim(jobid)) <> 0 then
                jobid = jobid
                else
                jobid = 0
                end if

                findesbesked = 0
                strSQL = "SELECT message, c.editor, editdate, edittime, m.mnavn FROM chat c LEFT JOIN medarbejdere m ON (c.editor = m.mid) WHERE jobid = "& jobid
                'response.Write strSQL
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
                findesbesked = 1

                %>
            
                <tr>
                    <td><%=oRec("mnavn") %> - <%=oRec("edittime") %> <br />
                        <b><%=oRec("message") %></b>
                    </td>
                </tr>

                <%
                    oRec.movenext
                    wend
                    oRec.close
                %>


                <%if findesbesked <> 1 then %>
                    <tr>
                        <td>Ingen beskeder</td>
                    </tr>
                <%end if %>
                </tbody>
            </table>
        </div>

        <table style="width:100%; margin-top:7px;">
            <tr>
                <td><textarea name="FM_chatmessage" class="form-control input-small"></textarea></td>
                <td><button type="submit" class="btn btn-success btn-sm pull-right"><b>Send</b></button></td>
            </tr>
        </table>


        <script type="text/javascript">

            $("#chatscroll").scrollTop($("#chatscroll")[0].scrollHeight);

        </script>

<%end function %>



<%function fileoverview(jobid, kundeid) %>

    

    <%
        call TimeOutVersion()

        if len(trim(request("FM_sog"))) <> 0 OR request("fms") = "1" then
        
            if len(trim(request("FM_sog"))) <> 0 then

                sogeKri = request("FM_sog")
                sogTxt = sogeKri

            else

                'sogeKri = "fdsfD3fve#"
                sogeKri = ""
                sogTxt = ""

            end if

            response.cookies("tsa")("filersog") = sogTxt

        else

            if jobid <> 0 then
                strSQLjob = "SELECT id, jobnavn, jobnr FROM job WHERE id = "& jobid
                oRec.open strSQLjob, oConn, 3
                if not oRec.EOF then

                'sogeKri = oRec("jobnavn")
                sogeKri = ""
                sogTxt = sogeKri

                end if
                oRec.close

            else

                if request.cookies("tsa")("filersog") <> "" then
                    sogeKri = request.cookies("tsa")("filersog")
                    sogTxt = sogeKri
                else
                    'sogeKri = "fdsfD3fve#"
                    sogeKri = ""
                    sogTxt = ""
                end if

            end if
    end if

    if len(request("fms")) <> 0 then
        if request("vistommefoldere") <> 0 then
        vistommefoldere = 1
        vistommefoldereCHK = "CHECKED"
        else
        vistommefoldere = 0
        vistommefoldereCHK = ""
        end if

        response.cookies("TSA")("tommefoldere") = vistommefoldere

    else
        
        if request.cookies("TSA")("tommefoldere") = "1" then
        vistommefoldere = 1
        vistommefoldereCHK = "CHECKED"
        else
        vistommefoldere = 0
        vistommefoldereCHK = ""
        end if

    end if


    %>

    <SCRIPT src="../timereg/inc/filer_jav.js"></script>

    <%if (level <= 2 OR level = 6) then %>

    <FORM ENCTYPE="multipart/form-data" ACTION="upload_bin.asp" METHOD="POST" id="fileform">
       
        <div class="row">
            <div class="col-lg-12">
                <div class="fileinput fileinput-new" data-provides="fileinput">
                <div class="fileinput-preview thumbnail" data-trigger="fileinput" style="width:100%; height: 30px; border:none;"></div>
                <div>
                    <span class="btn btn-default btn-file"><span class="fileinput-new" style="width:75px;">Vælg fil</span><span class="fileinput-exists" style="width:75px;">Change</span><INPUT id="fileToUpload" NAME="fileupload1" TYPE="file"></span>
                    <a href="#" class="btn btn-default fileinput-exists" data-dismiss="fileinput" style="width:75px;">Remove</a>

                    <a style="width:75px;" class="btn btn-sm btn-success" id="uploadbtn">Upload</a>
                </div>
                </div>
            </div>

        </div>

    </form>
    <input type="hidden" value="<%=jobid %>" id="uploadjobid" />

    <script type="text/javascript">
        $("#uploadbtn").bind('click', function () {
          // alert("klik")
            
            var file = $("#fileToUpload").val();
            
            if (file != null && file.length > 0) {
                uploadjobid = $("#uploadjobid").val();
               // alert(uploadjobid)
                
                $.post("?jobid=" + uploadjobid, { control: "FileuploadKlargoring", AjaxUpdateField: "true" }, function (data) {
                    
                    // Efter der er oprettet en r;kke til den opkommende fil, bliver selve filen lagt up
                    $("#fileform").submit();
               });

            }
        });
    </script>

    <%end if %>

    <table id="filer" style="width:100%;" class="table dataTable table-striped table-bordered table-hover ui-datatable">
        
        <thead>
            <tr>
                <th>Folder / Filer navn</th>
                <th>Dato</th>
                <th>Medarbejder</th>
                <th>Slet</th>
            </tr>
        </thead>

        <tbody>
            <%
            strSQL = "SELECT id, filnavn, dato, editor FROM filer WHERE jobid = "& jobid & " GROUP BY filnavn"
            oRec.open strSQL, oConn, 3
            while not oRec.EOF
            %>
            <tr>
                <td><a href="https://timeout.cloud/timeout_xp/wwwroot/<%=toVer %>/inc/upload/<%=lto%>/<%=oRec("filnavn")%>" class='vmenulille' target="_blank"><%=oRec("filnavn")%></a></td>
                <td><%=oRec("dato") %></td>
                <td><%=oRec("editor") %></td>
                <td style="text-align:center;"><a href="jobs.asp?func=deleteFile&fileid=<%=oRec("id") %>&jobid=<%=jobid %>"><span style="color:darkred;" class="fa fa-times"></span></a></td>
            </tr>
            <%
            oRec.movenext
            wend
            oRec.close 
            %>
        </tbody>


        <!-- Meget advanceret metode som ikke er nøvendig pt. derfor display:none -->
        <tbody style="display:none;">
            <%
                if sogeKri = "fdsfD3fve#" then
                'response.end
                end if


                jobidSQL2 = ""

                kundeseSQL = ""
	    
                kundeIdSQL = " k.kkundenavn LIKE '%"& sogeKri &"%' OR k.kkundenr = '"& sogeKri &"'"
                'jobIdSQL = " OR j2.jobnavn LIKE '%"& sogeKri &"%' OR j2.jobnr = '"& sogeKri &"'"
                jobIdSQL = ""
                filerIdSQL = " OR fi.filnavn LIKE '%"& sogeKri &"%'"
                folderIdSQL = " OR fo.navn LIKE '%"& sogeKri &"%'"
	
	            strSQL = "SELECT fo.kundeid AS kundeid, fo.navn AS foldernavn, "_
	            &" fi.adg_kunde, fi.adg_admin, fi.adg_alle, fo.id AS foid, fo.kundese, "_
	            &" fo.jobid AS jobid, filnavn, fi.id AS fiid, COUNT(fi.id) AS antalfiler,"_
	            &" fi.dato, fi.editor as streditor,"_
	            &" j1.jobnr AS j1jobnr, j1.jobnavn AS j1jobnavn, j1.jobans1 AS j1jobans1, j1.jobans2 AS j1jobans2, "_
	            &" j2.jobnr As j2jobnr, j2.jobnavn As j2jobnavn, j2.jobans1 AS j2jobans1, j2.jobans2 AS j2jobans2, "_
	            &" fo.dato AS fodato, kkundenavn, kkundenr"_
	            &" FROM foldere fo "_
	            &" LEFT JOIN filer AS fi ON (fi.folderid = fo.id "& jobidSQL2 &") "_
	            &" LEFT JOIN job AS j1 ON (j1.id = fo.jobid) "_
	            &" LEFT JOIN job AS j2 ON (j2.id = fi.jobid) "_
                &" LEFT JOIN kunder AS k ON (k.kid = fo.kundeid) "_
	            &" WHERE ("& kundeIdSQL &" "& jobIdSQL &" "& filerIdSQL &" "& folderIdSQL &") "& kundeseSQL &" AND fo.jobid = "& jobid &" GROUP BY foid, fiid ORDER BY fo.navn"

                'response.Write strSQL
                'Response.flush

                x = 0
                y = 0
	            lastfolderid = 0
	            oRec.open strSQL, oConn, 3 
	            while not oRec.EOF 


                '******************** Folder *****************************

                if cint(vistommefoldere) = 1 OR oRec("antalfiler") = "1" then

                    if lastfolderid <> oRec("foid") then

                        y = y + 1

                        '** Er det kunde eller medarbejder der er logget ind ?
		                if kundelogin <> 1 then 'ikke kunde logget ind => Medarbjeder logget ind
                        
                            '** Tjekker rettigheder eller om man er jobanssvarlig ***

		                    editok = 0
                            '** På foldere der er tilknyttet et job ***
			                if len(oRec("j1jobnavn")) <> 0 then
			
			
				                if level = 1 then '** Administrator
				                editok = 1
				                else
						                '*** jobans 
						                if cint(session("mid")) = oRec("j1jobans1") OR cint(session("mid")) = oRec("j1jobans2") OR _
						                (cint(oRec("j1jobans1")) = 0 AND cint(oRec("j1jobans2")) = 0) then
						                editok = 1
						                end if
				                end if
			
			                else
			                '** På foldere der IKKE er tilknyttet et job ***
				
				                if level <= 3 OR level = 6 then '** Admin eller niveau 1 må redigere FOLDER
				                editok = 1
				                else
				                editok = 0
				                end if
			
			                end if
            %>
                            <tr id="trid_<%=oRec("foid")%>" class="trfo">
                                <td>
                                    <%if oRec("antalfiler") = "1" then 
                                    spcls = "showfolder"
                                    spcol = "#000000"%>
					                <span id="foid_<%=oRec("foid")%>" class="showfolder"><img src="../ill/folder.png"  alt="" border="0"/></span>
                                    <%else 
                                    spcls = ""
                                    spcol = "#999999"%>
                                    <img src="../ill/folder_blue.png"  alt="" border="0"/>
					                <%end if %>

                                    <span id="fxid_<%=oRec("foid")%>" class="<%=spcls %>" style="color:<%=spcol %>;"><b><u>+ <%=left(oRec("foldernavn"), 20)%></u></b></span>
					                <%if oRec("foid") <> 14 AND oRec("foid") <> 500 AND oRec("foid") <> 1000 AND editok = 1 then%>
                                    &nbsp <a href="../timereg/filer.asp?func=fored&kundeid=<%=oRec("kundeid")%>&id=<%=oRec("foid")%>&nomenu=<%=nomenu%>" target="_black"><span class="fa fa-pencil"></span></a>
				                    <%end if%>
                                </td>

                                <td style="text-align:center;"><%=oRec("fodato")%>&nbsp;</td>

                                <td>
                                   <!-- <%if len(trim(oRec("kkundenavn"))) <> 0 then%>
					                <b><%=left(oRec("kkundenavn"), 30)%> (<%=oRec("kkundenr")%>)</b>
					                <%end if%>

					                <%if len(trim(oRec("j1jobnavn"))) <> 0 then%>
					                <br /><%=left(oRec("j1jobnavn"), 30)%> (<%=oRec("j1jobnr")%>)
					                <%end if%> -->
					            </td>
                            </tr>

                        <%else 'kundelogin %>

                            <tr>
                                <td> <img src="../ill/folder_ikon.gif" width="17" height="15" alt="" border="0"><b><%=oRec("foldernavn")%></b></td>
                                <td style="text-align:center;"><%=oRec("fodato")%></td>
                                <td>                                   
                                    <!-- 
			                        <%if len(trim(oRec("j1jobnavn"))) <> 0 then%>
			                        <%=oRec("j1jobnavn")%> (<%=oRec("j1jobnr")%>)
			                        <%end if%> -->
			                    </td>
                            </tr>
                
                        <%end if 'kundelogin

                    end if 'lastfolder 



                    '******************** Filer *****************************
	                if isNull(oRec("filnavn")) <> true then
                        '** Er det kunde der er logget ind ?
		                if kundelogin <> 1 then

                            filnavnTxt = replace(oRec("filnavn"), "_txt", ".txt")
                %>
                            <tr class="foid_<%=oRec("foid")%>">
                                <td><img src="../ill/addmore55.gif" width="10" height="13" alt="" border="0">&nbsp;<a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/upload/<%=lto%>/<%=filnavnTxt%>" target="_blank"><%=left(filnavnTxt, 40)%></a> [<%=right(filnavnTxt, 3) %>]
					            <%if (level = 1 OR oRec("adg_alle") = 1) then%>
					            <a href="../timereg/filer.asp?kundeid=<%=oRec("kundeid")%>&jobid=<%=jobid%>&func=redfil&id=<%=oRec("fiid")%>&nomenu=1" target="_blank"><span class="fa fa-pencil"></span></a>
					            <%end if%>
					            </td>

                                <td style="text-align:center;"><%=oRec("dato")%></td>

                                <td>
                                    <%=left(oRec("streditor"), 12) %>
                                    <!--
						            <%if len(trim(oRec("j2jobnavn"))) <> 0 then%>
					                <%=left(oRec("j2jobnavn"), 30)%> (<%=oRec("j2jobnr")%>)
					                <%end if%> -->
						        </td>
                            </tr>
                
                        <%else 'kundelogin

                            if len(oRec("filnavn")) <> 0 then %>
			                <tr>	
					
					                <td style="padding-left:20px; border-bottom:1px #cccccc solid;"><img src="../ill/addmore55.gif" width="10" height="13" alt="" border="0">&nbsp;<a href="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/upload/<%=lto%>/<%=oRec("filnavn")%>" class='vmenulille' target="_blank"><%=oRec("filnavn")%></a>&nbsp;</td>
					                <td align=center class=lille ><%=oRec("dato")%></td>
					                <td colspan=8 >&nbsp;</td>
			                </tr>
			                <%end if

                        end if

                    end if 'filnavn

                lastfilnavn = oRec("filnavn")
	            lastfolderid = oRec("foid") 
	
	            x = x + 1

                end if 'tommefiler oRec("antalfiler") = "1"

                response.flush
            
                oRec.movenext
                wend
                oRec.close
            %>

        </tbody>


    </table>

    <script type="text/javascript">
        var table = $('#filer').DataTable({
            scrollY: "303px",
            scrollX: true,
            scrollCollapse: true,
            paging: false,
            ordering: false

        });
    </script>


<%end function %>





