

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->


<% '**** Søgekriterier AJAX **'
        'section for ajax calls

    if Request.Form("AjaxUpdateField") = "true" then
    Select Case Request.Form("control")

   
    'case "eval_submit"

    case "xeval_submit_lbn_hours"

   
    


    end select
    response.End
    end if

    jobid_til_eval = request("jobid_til_eval")

%>



<%
    if len(session("user")) = 0 then

    errortype = 5
    call showError(errortype)
    response.End
    end if 

    func = request("func")


    if len(trim(request("noupdate"))) <> 0 then
    noupdate = request("noupdate")
    else
    noupdate = 0
    end if
%>






        <%
            select case func 

            case "dbred"

                if len(trim(request("eval_vur"))) = 0 OR _
                len(trim(request("timerBrugt"))) = 0 OR _
                len(trim(request("timePris"))) = 0 OR _
                len(trim(request("omregnetOms"))) = 0 OR len(trim(request("eval_comment"))) = 0 then
                
                errortype = 189
		        call showError(errortype)
		        Response.end

                
                end if

                eval_vur = request("eval_vur")
                eval_jobid = request("eval_jobid")
                eval_comment = request("eval_comment")
                
                timerBrugt = request("timerBrugt")
                timerBrugt = replace(timerBrugt,".","")
                timerBrugt = replace(timerBrugt,",",".")
                
                timePris = request("timePris")
                timePris = replace(timePris,",",".")
                
                omregnetOms = request("omregnetOms")
                omregnetOms = replace(omregnetOms,".","")
                omregnetOms = replace(omregnetOms,",",".")
                
                eval_lbn_forskel = request("eval_lbn_forskel")
                eval_lbn_forskel = replace(eval_lbn_forskel, ".","")
                eval_lbn_forskel = replace(eval_lbn_forskel, ",",".")

                CHECK_fastpris = request("CHECK_fastpris")
                if cint(CHECK_fastpris) = 1 then
                eval_original_price = request("fastprisjob_oms")
                else
                eval_original_price = request("for_oms_result")
                end if

                eval_original_price = replace(eval_original_price,".","")
                eval_original_price = replace(eval_original_price,",",".")
    
                lastid = 0
                strSQLevaler = "SELECT eval_id FROM eval ORDER BY eval_id"
                oRec.open strSQLevaler, oConn, 3
                while not oRec.EOF
                lastid = oRec("eval_id")
                oRec.movenext
                wend
                oRec.close
                lastid = lastid + 1

                strRedorOpr = "SELECT eval_id FROM eval WHERE eval_jobid = "&eval_jobid
                oRec.open strRedorOpr, oConn, 3

                if not oRec.EOF then
                strEvalSub = "UPDATE eval SET eval_evalvalue = "& eval_vur &", eval_jobvaluesuggested = "& omregnetOms & ", eval_comment = '"&eval_comment&"', eval_suggested_hours = "& timerBrugt &", eval_suggested_hourly_rate = "& timePris &", eval_diff = "& eval_lbn_forskel &", eval_original_price = "& eval_original_price &" WHERE eval_jobid = "& eval_jobid 
                else
                strEvalSub = "INSERT INTO eval SET eval_id = "& lastid &", eval_jobid = "& eval_jobid &", eval_evalvalue = "& eval_vur &", eval_jobvaluesuggested = "& omregnetOms & ", eval_comment = '"&eval_comment&"', eval_diff = "& eval_lbn_forskel &", eval_suggested_hours = "& timerBrugt &", eval_suggested_hourly_rate = "& timePris &", eval_original_price = "& eval_original_price
                end if
                oRec.close 


               'response.write strEvalSub
               'response.flush
                oConn.execute(strEvalSub)


                '*** Skift jobstatus ************************ 
                strSQLjob = "UPDATE job SET jobstatus = 4 WHERE id = "& eval_jobid
                oConn.execute(strSQLjob)

                '********************************************************************

                'Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
                'Response.Write("<script language=""JavaScript"">window.close();</script>")

                'Response.end

                %>
                <div style="padding:40px;">
                Tak for din evaluering. Jobbet har skiftet status til "Gennemsyn",<br />
                <a href="#" onclick="Javascript:window.close();">[Luk dette vindue]</a>
                    </div>
                <%

               

            case "opr", "red"

            call akttyper2009(2)
          

            strsql = "SELECT id, jobnavn, jo_bruttooms, fastpris, jobnr, job_internbesk, jobknr, kkundenavn, kkundenr FROM job "_
            &" LEFT JOIN kunder k ON (kid = jobknr) WHERE id = "& jobid_til_eval

            oRec.open strsql, oConn, 3
            if not oRec.EOF then
            strJobnavn = oRec("jobnavn")
            strBruttooms = oRec("jo_bruttooms")
            strJobid = oRec("id")
            strFastpris = oRec("fastpris")
            strJobnr = oRec("jobnr")
            job_internbesk = oRec("job_internbesk")
            kkundenavn = oRec("kkundenavn") 
            kkundenr = oRec("kkundenr")
            end if
            oRec.close

            'if strFastpris = 0 then
            'lbn timer
            fast_Pris_akt_gens = 0
            TotalGNStimepris = 0
            GNStimePris_ikke_fastpris_aktiviteter = 0
            fastPris_Aktiviteter_Total_TP = 0
            totalTimer = 0
            omsResultat = 0
            antalFastTP = 0
            strSQLfastpris = "SELECT id, brug_fasttp, fasttp, navn FROM aktiviteter WHERE job = "& strJobid &" AND brug_fasttp = 1"
            oRec2.open strSQLfastpris, oConn, 3
            while not oRec2.EOF
            antalFastTP = antalFastTP + 1
            'response.Write "<br>" & oRec2("navn") & " fastpris " & oRec2("fasttp") & "rg: " & oRec2("brug_fasttp")

            if cint(antalFastTP) = 0 then
            fastPris_Aktiviteter_Total_TP = oRec2("fasttp")
            else
            fastPris_Aktiviteter_Total_TP = fastPris_Aktiviteter_Total_TP + oRec2("fasttp")
            end if

            oRec2.movenext
            wend
            oRec2.close
            if antalFastTP <> 0 then
            fast_Pris_akt_gens = fastPris_Aktiviteter_Total_TP / antalFastTP
            else
            fast_Pris_akt_gens = 0
            end if
            'response.Write "<br>" & fast_Pris_akt_gens


            strSQLAVG_ikke_fastpris_aktiviteter = "SELECT AVG(TimePris) as GNStimepris FROM timer "_
            &" LEFT JOIN aktiviteter as a ON (a.job = "& strJobid &") WHERE tjobnr = '"& strJobnr &"' AND a.brug_fasttp = 0 AND ("& aty_sql_realhours &")"
            oRec2.open strSQLAVG_ikke_fastpris_aktiviteter, oConn, 3

            if not oRec2.EOF then
            if oRec2("GNStimepris") <> 0 then
            GNStimePris_ikke_fastpris_aktiviteter = oRec2("GNStimepris")
            else
            GNStimePris_ikke_fastpris_aktiviteter = 0
            end if
            'response.Write "<br> T" & GNStimePris_ikke_fastpris_aktiviteter
            end if
            oRec2.close
            
            if antalFastTP <> 0 AND GNStimePris_ikke_fastpris_aktiviteter <> 0 then
            TotalGNStimepris = (GNStimePris_ikke_fastpris_aktiviteter + fast_Pris_akt_gens) / 2
            else
            TotalGNStimepris = (GNStimePris_ikke_fastpris_aktiviteter + fast_Pris_akt_gens) / 1
            end if

            'response.Write "<br>TTP" & TotalGNStimepris

            StrSQLTimerTOT = "SELECT sum(timer) as totTimer FROM timer WHERE tjobnr = '"& strJobnr &"' AND ("& aty_sql_realhours &")"
            oRec2.open StrSQLTimerTOT, oConn, 3
            if not oRec2.EOF then
            totalTimer = oRec2("totTimer")
            end if
            oRec2.close

            if totalTimer <> 0 then
            totalTimer = totalTimer
            else
            totalTimer = 0
            end if
            
            omsResultat = totalTimer * TotalGNStimepris
            'response.Write "<br>" & "Tottimer" & totalTimer
            'lbn timer slut
            'end if

            if func = "red" then
            
                strSQLEval = "SELECT eval_jobid, eval_evalvalue, eval_jobvaluesuggested, eval_comment, eval_diff, eval_suggested_hours, eval_suggested_hourly_rate FROM eval WHERE eval_jobid = "& jobid_til_eval 
                oRec2.open strSQLEval, oConn, 3
                if not oRec2.EOF then

                    select case cdbl(oRec2("eval_evalvalue"))
                    case 4
                    valueBarCH4 = "CHECKED"
                    valueBarCH3 = ""
                    valueBarCH2 = ""
                    valueBarCH1 = ""
                    case 3
                    valueBarCH4 = ""
                    valueBarCH3 = "CHECKED"
                    valueBarCH2 = ""
                    valueBarCH1 = ""
                    case 2
                    valueBarCH4 = ""
                    valueBarCH3 = ""
                    valueBarCH2 = "CHECKED"
                    valueBarCH1 = ""
                    case 1
                    valueBarCH4 = ""
                    valueBarCH3 = ""
                    valueBarCH2 = ""
                    valueBarCH1 = "CHECKED"
                    end select

                    strEval_comment = oRec2("eval_comment")
                    strEval_suggested_hours = oRec2("eval_suggested_hours")
                    strEval_suggested_houry_rate = oRec2("eval_suggested_hourly_rate")

                end if
                oRec2.close

            else 'Ikke rediger

                valueBarSEL4 = ""
                valueBarSEL3 = ""
                valueBarSEL2 = ""
                valueBarSEL1 = ""
                strEval_comment = ""
                strEval_suggested_hours = totalTimer
                strEval_suggested_houry_rate = TotalGNStimepris


            end if

             
        %>

<script src="js/eval_jav3.js" type="text/javascript"></script>


<div class="wrapper"><br /><br />
<!--  
    <div class="content">
      -->

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title">Projekt evaluering</h3>
                <div class="portlet-body">
                   <b>Kunde & Job:</b><br />
                   <%response.Write kkundenavn & " ("& kkundenr &")<br>"& strJobnavn %>
                   <input type="hidden" name="CHECK_fastpris" id="CHECK_fastpris" value="<%=strFastpris %>" />
                   <div class="row">
                       <div class="col-lg-12" style="text-align:center"><h5>Du bedes evaluere om projektet har været God +, God, Mellem, eller Dårligt</h5>
                         
                           <span style="font-size:12px">(Vurderingen skal være på Økonomi - Tid - Kvalitet)</span>
                       </div>
                   </div>


                    
                    <form id="xeval_vur_form" method="post" action="eval.asp?func=dbred">
                    <div class="row">
                        <div class="col-lg-12" style="text-align:center;">
                        <b style="color:cornflowerblue">God +</b> <input type="radio" name="eval_vur" value="4" <%=valueBarCH4 %> />
                        &nbsp&nbsp&nbsp
                        <b style="color:yellowgreen">God</b> <input type="radio" name="eval_vur" value="3" <%=valueBarCH3 %> />
                        &nbsp&nbsp&nbsp
                        <b style="color:orange">Middel</b> <input type="radio" name="eval_vur" value="2" <%=valueBarCH2 %> />
                        &nbsp&nbsp&nbsp
                        <b style="color:red">Dårlig</b> <input type="radio" name="eval_vur" value="1" <%=valueBarCH1 %> />

                        <!--&nbsp&nbsp&nbsp
                        <b>Ikke evalueret</b> <input type="radio" name="eval_vur" value="0" <%=valueBarCH0 %> />-->
                        </div>
                    </div>
                    <br /><br />
                    <div class="row">
                        <div class="col-lg-3"></div>
                        <div class="col-lg-4">Intern note:<br />
                            <%=job_internbesk %>

                        </div>
                    </div>
                    <br /><br />
                    <div class="row">
                        <div class="col-lg-3"></div>
                        <div class="col-lg-4">Kommentar:</div>
                    </div>
                    <div class="row">
                        <div class="col-lg-3"></div>
                        <div class="col-lg-6">
                            <table>
                                <tr>
                                    <td><textarea class="form-control input-small" cols="100" rows="5" name="eval_comment" id="eval_comment"><%=strEval_comment %></textarea></td>
                                </tr>
                            </table>
                        </div>
                    </div>
                    <br /><br />

                    <%if strFastpris = 1 then%>
                    <div class="row">
                        <div class="col-lg-3"></div>
                        <div class="col-lg-6">
                            <table>
                                <tr>
                                    <td style="width:13px"></td>
                                    <td style="color:black; padding-right:5px">Fastpris:</td>
                                    <td style="width:350px"></td>                                    
                                    <td style="width:125px; padding-right:10px;"><input type="text" name="fastprisjob_oms" id="fastprisjob_oms" class="form-control input-small" value="<%=strBruttooms %>" readonly/></td>
                                </tr>
                            </table>
                        </div>
                    </div>
                    
                    <%end if %>
                   <div class="row">
                        <div class="col-lg-3"></div>
                        <div class="col-lg-6">
                            <table>
                                <tr>
                                    <td style="width:13px"></td>
                                    <td style="color:black; padding-right:10px">Timer brugt</td>
                                    <td style="width:100px; padding-right:10px"><input type="text" name="LBN_totaltimer" id="LBN_totaltimer" class="form-control input-small" readonly value="<%=replace(formatnumber(totalTimer, 2), ".", "") %>" /></td>
                                    <td style="padding-right:10px">*</td>
                                    <td style="color:black; padding-right:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px"><input type="text" name="LBN_gns_timepris" id="LBN_gns_timepris" class="form-control input-small" readonly value="<%=replace(formatnumber(TotalGNStimepris, 2), ".", "") %>"/></td>
                                    <td style="padding-right:10px">=</td>
                                    <td style="width:125px; padding-right:10px"><input type="text" name="for_oms_result" id="for_oms_result" class="form-control input-small" readonly value="<%=replace(formatnumber(omsResultat, 2), ".", "") %>"/></td>
                                </tr>
                            </table>
                        </div>
                    </div>                   

                    <div class="row">
                        <div class="col-lg-3"></div>
                        <div class="col-lg-6">
                            <table>
                                <tr>
                                    <td style="width:13px"></td>
                                    <td style="color:black; padding-right:26px">Foreslået</td>
                                    <td style="width:100px; padding-right:10px"><input type="text" name="timerBrugt" id="timerBrugt" class="form-control input-small opdaterOms" value="<%=strEval_suggested_hours %>" /></td>
                                    <td style="padding-right:10px">*</td>
                                    <td style="color:black; padding-right:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px"><input type="text" name="timePris" id="timePris" class="form-control input-small opdaterOms" value="<%=strEval_suggested_houry_rate %>" /></td>
                                    <td style="padding-right:10px">=</td>
                                    <td style="width:125px; padding-right:10px"><input type="text" name="omregnetOms" id="omregnetOms" class="form-control input-small" readonly/>
                                        <input type="hidden" name="eval_lbn_forskel" id="eval_lbn_forskel" class="form-control input-small" readonly/>
                                    </td>
                                </tr>
                            </table>
                        </div>
                    </div>

                    <%'end if %>

                    <%if cint(noupdate) <> 1 then %>
                    <br /><br /><br />
                    <div class="row">
                        <div class="col-lg-8"></div>
                        <input type="hidden" name="eval_jobid" id="eval_jobid" value="<%=strJobid %>" />
                        <div class="col-lg-2"><!--<button class="btn btn-success btn-sm pull-right eval_sub_btn"><b>Opdatér</b></button>-->
                            <input type="submit" value="Opdatér" class="btn btn-success btn-sm pull-right eval_sub_btn">
                        </div>
                    </div>
                    <%end if %>
                </div>
            </div>
        </div>
        </form>

        <%end select %>

    </div>
</div> 

<!--#include file="../inc/regular/footer_inc.asp"-->