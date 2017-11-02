

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

                'if len(trim(request("eval_vur"))) = 0 OR _
                'len(trim(request("timerBrugt"))) = 0 OR _
                'len(trim(request("timePris"))) = 0 OR _
                'len(trim(request("omregnetOms"))) = 0 OR len(trim(request("eval_comment"))) = 0 then
                if len(trim(request("eval_vur"))) = 0 OR _
                len(trim(request("forslaet_oms"))) = 0 OR _
                len(trim(request("eval_comment"))) = 0 then
                 

                errortype = 189
		        call showError(errortype)
		        Response.end

                
                end if

                eval_vur = request("eval_vur")
                eval_jobid = request("eval_jobid")
                eval_comment = request("eval_comment")
                
                'timerBrugt = request("timerBrugt")
                'timerBrugt = replace(timerBrugt,".","")
                'timerBrugt = replace(timerBrugt,",",".")
                
                'timePris = request("timePris")
                'timePris = replace(timePris,",",".")
                
                'omregnetOms = request("omregnetOms")
                'omregnetOms = replace(omregnetOms,".","")
                'omregnetOms = replace(omregnetOms,",",".")

                fak_svendetimer = request("fak_svendetimer")
                fak_svendetimer = replace(fak_svendetimer, ".","")
                fak_svendetimer = replace(fak_svendetimer, ",",".")
                fak_svendetimePris = request("fak_svendetimePris") 
                fak_svendetimePris = replace(fak_svendetimePris, ".","")
                fak_svendetimePris = replace(fak_svendetimePris, ",",".")
                           
                ubemandet_maskine_timer = request("ubemandet_maskine_timer")
                ubemandet_maskine_timer = replace(ubemandet_maskine_timer, ".","")
                ubemandet_maskine_timer = replace(ubemandet_maskine_timer, ",",".")
                ubemandet_maskine_timePris = request("ubemandet_maskine_timePris")
                ubemandet_maskine_timePris = replace(ubemandet_maskine_timePris, ".","")
                ubemandet_maskine_timePris = replace(ubemandet_maskine_timePris, ",",".")

                laer_timer = request("laer_timer")
                laer_timer = replace(laer_timer, ".","")
                laer_timer = replace(laer_timer, ",",".")
                laer_timepris = request("laer_timepris")
                laer_timepris = replace(laer_timepris, ".","")
                laer_timepris = replace(laer_timepris, ",",".")

                easy_reg_timer = request("easy_reg_timer")
                easy_reg_timer = replace(easy_reg_timer, ".","")
                easy_reg_timer = replace(easy_reg_timer, ",",".")
                easy_reg_timepris = request("easy_reg_timepris")
                easy_reg_timepris = replace(easy_reg_timepris, ".","")
                easy_reg_timepris = replace(easy_reg_timepris, ",",".")

                ikke_fakbar_tid_timer = request("ikke_fakbar_tid_timer")
                ikke_fakbar_tid_timer = replace(ikke_fakbar_tid_timer, ".","")
                ikke_fakbar_tid_timer = replace(ikke_fakbar_tid_timer, ",",".")
                ikke_fakbar_tid_timepris = request("ikke_fakbar_tid_timepris")
                ikke_fakbar_tid_timepris = replace(ikke_fakbar_tid_timepris, ".","")
                ikke_fakbar_tid_timepris = replace(ikke_fakbar_tid_timepris, ",",".")

                forslaet_timer_total = request("forslaet_timer_total")
                forslaet_timer_total = replace(forslaet_timer_total, ".","")
                forslaet_timer_total = replace(forslaet_timer_total, ",",".")
                'forslaet_timepris_total = request("forslaet_timepris_total")
                'forslaet_timepris_total = replace(forslaet_timepris_total, ".","")
                'forslaet_timepris_total = replace(forslaet_timepris_total, ",",".")

                forslaet_oms = request("forslaet_oms")
                forslaet_oms = replace(forslaet_oms, ".","")
                forslaet_oms = replace(forslaet_oms, ",",".")

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
                'response.Write "Her: " & eval_original_price & "her: " & forslaet_oms &"res "& eval_lbn_forskel & "<br>"

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
                strEvalSub = "UPDATE eval SET eval_evalvalue = "& eval_vur &", eval_jobvaluesuggested = "& forslaet_oms & ", eval_comment = '"&eval_comment&"', eval_diff = "& eval_lbn_forskel &", eval_original_price = "& eval_original_price _
                &", eval_fakbartimer = "& fak_svendetimer &", eval_fakbartimepris = "& fak_svendetimePris &", ubemandet_maskine_timer = "& ubemandet_maskine_timer &", ubemandet_maskine_timePris = "& ubemandet_maskine_timePris _
                &", laer_timer = "& laer_timer &", laer_timepris = "& laer_timepris &", easy_reg_timer = "& easy_reg_timer &", easy_reg_timepris = "& easy_reg_timepris & ", ikke_fakbar_tid_timer = "& ikke_fakbar_tid_timer &", ikke_fakbar_tid_timepris = "& ikke_fakbar_tid_timepris _
                &", eval_suggested_hours = "& forslaet_timer_total _
                &" WHERE eval_jobid = "& eval_jobid 
                else
                strEvalSub = "INSERT INTO eval SET eval_id = "& lastid &", eval_jobid = "& eval_jobid &", eval_evalvalue = "& eval_vur &", eval_jobvaluesuggested = "& forslaet_oms & ", eval_comment = '"&eval_comment&"', eval_diff = "& eval_lbn_forskel &", eval_suggested_hours = "& forslaet_timer_total &", eval_original_price = "& eval_original_price _
                &", eval_fakbartimer = "& fak_svendetimer &", eval_fakbartimepris = "& fak_svendetimePris &", ubemandet_maskine_timer = "& ubemandet_maskine_timer &", ubemandet_maskine_timePris = "& ubemandet_maskine_timePris _
                &", laer_timer = "& laer_timer &", laer_timepris = "& laer_timepris &", easy_reg_timer = "& easy_reg_timer &", easy_reg_timepris = "& easy_reg_timepris & ", ikke_fakbar_tid_timer = "& ikke_fakbar_tid_timer &", ikke_fakbar_tid_timepris = "& ikke_fakbar_tid_timepris
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
            totalTimerFastprisAkt = 0
            omsResultat = 0
            antalFastTP = 0
            sumtimePris_ikke_fastpris_aktiviteter = 0



                '***** Fastpris AKT *****
          
                strSQLfastpris = "SELECT a.id, brug_fasttp, fasttp, navn, SUM(t.timer) AS sumtimer FROM aktiviteter a "_
                &" LEFT JOIN timer t ON (t.taktivitetid = a.id) WHERE job = "& strJobid &" AND brug_fasttp = 1 AND ("& aty_sql_realhours &") GROUP BY taktivitetid"

                'Response.write strSQLfastpris & "<br>"
                oRec2.open strSQLfastpris, oConn, 3
                while not oRec2.EOF
                
                'response.Write "<br>" & oRec2("navn") & " fastpris " & oRec2("fasttp") & "rg: " & oRec2("brug_fasttp")

                'if cint(antalFastTP) = 0 then
                'fastPris_Aktiviteter_Total_TP = oRec2("fasttp")
                'else
                fastPris_Aktiviteter_Total_TP = fastPris_Aktiviteter_Total_TP + (oRec2("fasttp")*oRec2("sumtimer")*1)
               'end if
                totalTimerFastprisAkt = totalTimerFastprisAkt + (oRec2("sumtimer")*1)

                antalFastTP = antalFastTP + 1
                oRec2.movenext
                wend
                oRec2.close


            'Response.write "fastPris_Aktiviteter_Total_TP: "& fastPris_Aktiviteter_Total_TP & "<br>totalTimerFastprisAkt: "& totalTimerFastprisAkt & "<br>"
            'if antalFastTP <> 0 then
            'fast_Pris_akt_gens = fastPris_Aktiviteter_Total_TP / antalFastTP
            'else
            'fast_Pris_akt_gens = 0
            'end if
            'response.Write "<br>" & fast_Pris_akt_gens


            '******* Medarb TP aktiviteter ****'
            strSQLAVG_ikke_fastpris_aktiviteter = "SELECT SUM(timer*TimePris) as sumtimepris, SUM(timer) AS timer FROM timer "_
            &" LEFT JOIN aktiviteter as a ON (a.job = "& strJobid &") WHERE tjobnr = '"& strJobnr &"' AND a.brug_fasttp = 0 AND ("& aty_sql_realhours &") GROUP BY taktivitetid"
            'response.Write strSQLAVG_ikke_fastpris_aktiviteter
            totalTimerIkkeFastprisAkt = 0
            sumtimePris_ikke_fastpris_aktiviteter = 0

            oRec2.open strSQLAVG_ikke_fastpris_aktiviteter, oConn, 3

            while not oRec2.EOF 

            sumtimePris_ikke_fastpris_aktiviteter = sumtimePris_ikke_fastpris_aktiviteter + oRec2("sumtimepris")
            totalTimerIkkeFastprisAkt = totalTimerIkkeFastprisAkt + oRec2("timer")

            oRec2.movenext
            wend
            oRec2.close

            'response.write "<br>sumtimePris_ikke_fastpris_aktivitete: "& sumtimePris_ikke_fastpris_aktiviteter & "<br>totalTimerIkkeFastprisAkt: "& totalTimerIkkeFastprisAkt

            totalTimepris = (fastPris_Aktiviteter_Total_TP*1+sumtimePris_ikke_fastpris_aktiviteter*1) 
            totalTimer = (totalTimerFastprisAkt*1+totalTimerIkkeFastprisAkt*1)

            'if antalFastTP <> 0 AND GNStimePris_ikke_fastpris_aktiviteter <> 0 then
            'TotalGNStimepris = (GNStimePris_ikke_fastpris_aktiviteter + fast_Pris_akt_gens) / 2
            'else
            'TotalGNStimepris = (GNStimePris_ikke_fastpris_aktiviteter + fast_Pris_akt_gens) / 1
            'end if

            'response.Write "<br>TTP" & TotalGNStimepris

            'StrSQLTimerTOT = "SELECT sum(timer) as totTimer FROM timer WHERE tjobnr = '"& strJobnr &"' AND ("& aty_sql_realhours &")"
            'oRec2.open StrSQLTimerTOT, oConn, 3
            'if not oRec2.EOF then
            'totalTimer = oRec2("totTimer")
            'end if
            'oRec2.close

            if totalTimer <> 0 then
            totalTimer = totalTimer
            else
            totalTimer = 0
            end if
            
            omsResultat = totalTimepris 'totalTimer * 

            if totalTimepris <> 0 then
            TotalGNStimepris = (omsResultat*1/totalTimer*1)
            else
            TotalGNStimepris = 0
            end if
            'response.Write "<br>" & "Tottimer" & totalTimer
            'lbn timer slut
            'end if



            '***************************************************************
            '**** udsepcificering **************
            '***************************************************************
            fastPris_svendetimer_Total_TP = 0
            totalTimerFastpris_svendetimer = 0
            fakbarSvendeTimer_fastprisSQL = "SELECT a.id, brug_fasttp, fasttp, navn, SUM(t.timer) AS sumtimer FROM aktiviteter a "_
            &" LEFT JOIN timer t ON (t.taktivitetid = a.id) WHERE job = "& strJobid &" AND brug_fasttp = 1 AND Tfaktim = 1 AND (Timer >= 0.25) GROUP BY taktivitetid"
            oRec2.open fakbarSvendeTimer_fastprisSQL, oConn, 3
            while not oRec2.EOF
            fastPris_svendetimer_Total_TP = fastPris_svendetimer_Total_TP + (oRec2("fasttp")*oRec2("sumtimer")*1)
            totalTimerFastpris_svendetimer = totalTimerFastpris_svendetimer + (oRec2("sumtimer")*1)
            oRec2.movenext
            wend
            oRec2.close

            svendetimepris_ikkeFP_sum = 0
            svendetimer_ikkeFP_sumtimer = 0
            fakbarSvendeTimer_ikkefastprisSQL = "SELECT sum(timer*timepris) as sumSvendeTimepris_ikkeFP, sum(timer) as sumSvendeTimer_ikkeFP FROM timer LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE tjobnr = '"& strJobnr &"' AND Tfaktim = 1 AND (Timer >= 0.25) AND brug_fasttp = 0 GROUP BY taktivitetid"
            'response.Write fakbarSvendeTimer_ikkefastprisSQL
            oRec2.open fakbarSvendeTimer_ikkefastprisSQL, oConn, 3
            while not oRec2.EOF
            svendetimepris_ikkeFP_sum = svendetimepris_ikkeFP_sum + oRec2("sumSvendeTimepris_ikkeFP")
            svendetimer_ikkeFP_sumtimer = svendetimer_ikkeFP_sumtimer + oRec2("sumSvendeTimer_ikkeFP")
            oRec2.movenext
            wend
            oRec2.close 

            svendetimepristotal = (fastPris_svendetimer_Total_TP*1 + svendetimepris_ikkeFP_sum*1)
            svendetimertotal = (totalTimerFastpris_svendetimer*1 + svendetimer_ikkeFP_sumtimer*1)
               
            svendetimer_oms = svendetimepristotal '(svendetimepristotal * svendetimertotal) 
            'response.Write "<br><br><br>HER1 "& svendetimepris_ikkeFP_sum & "<br>" & "her2 " & svendetimer_ikkeFP_sumtimer

            'if svendetimerTot <> 0 then svendetimerTot = svendetimerTot else svendetimerTot = 0 end if
            'if svendetimeprisGNS <> 0 then svendetimeprisGNS = svendetimeprisGNS else svendetimeprisGNS = 0 end if
            if svendetimer_oms <> 0 then svendetimer_oms = svendetimer_oms else svendetimer_oms = 0 end if

            if svendetimer_oms <> 0 then
            totalGNSsvendeTimepris = (svendetimer_oms*1/svendetimertotal*1)
            else
            totalGNSsvendeTimepris = 0
            end if



            fastPris_maskine_Total_TP = 0
            totalTimerFastpris_maskine = 0
            fakbarMaskineTimer_fastprisSQL = "SELECT a.id, brug_fasttp, fasttp, navn, SUM(t.timer) AS sumtimer FROM aktiviteter a "_
            &" LEFT JOIN timer t ON (t.taktivitetid = a.id) WHERE job = "& strJobid &" AND brug_fasttp = 1 AND Tfaktim = 91 GROUP BY taktivitetid"
            oRec2.open fakbarMaskineTimer_fastprisSQL, oConn, 3
            while not oRec2.EOF
            fastPris_maskine_Total_TP = fastPris_maskine_Total_TP + (oRec2("fasttp")*oRec2("sumtimer")*1)
            totalTimerFastpris_maskine = totalTimerFastpris_maskine + (oRec2("sumtimer")*1)
            oRec2.movenext
            wend
            oRec2.close
            
            
            maskine_ikkeFP_sum = 0
            maskine_ikkeFP_sumtimer = 0
            fakbarmaskineTimer_ikkefastprisSQL = "SELECT sum(timer*timepris) as sumMaskineTimepris_ikkeFP, sum(timer) as sumMaskineTimer_ikkeFP FROM timer LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE tjobnr = '"& strJobnr &"' AND Tfaktim = 91 AND brug_fasttp = 0 GROUP BY taktivitetid"
            'response.Write fakbarmaskineTimer_ikkefastprisSQL
            oRec2.open fakbarmaskineTimer_ikkefastprisSQL, oConn, 3
            while not oRec2.EOF
            maskine_ikkeFP_sum = maskine_ikkeFP_sum + oRec2("sumMaskineTimepris_ikkeFP")
            maskine_ikkeFP_sumtimer = maskine_ikkeFP_sumtimer + oRec2("sumMaskineTimer_ikkeFP")
            oRec2.movenext
            wend
            oRec2.close 
            'response.Write "<br><br><br>HER1 "& maskine_ikkeFP_sum & "<br>" & "her2 " & maskine_ikkeFP_sumtimer

            maskinetimepristotal = (fastPris_maskine_Total_TP*1 + maskine_ikkeFP_sum*1)
            maskinetimertotal = (totalTimerFastpris_maskine*1 + maskine_ikkeFP_sumtimer*1)
            ubemandetMaskine_oms = maskinetimepristotal
            if ubemandetMaskine_oms <> 0 then 
            ubemandetMaskine_oms = ubemandetMaskine_oms
            maskine_GNS_timepris = (ubemandetMaskine_oms*1/maskinetimertotal*1) 
            else 
            ubemandetMaskine_oms = 0
            maskine_GNS_timepris = 0 
            end if


            fastPris_laer_Total_TP = 0
            totalTimerFastpris_lear = 0
            fakbarLearTimer_fastprisSQL = "SELECT a.id, brug_fasttp, fasttp, navn, SUM(t.timer) AS sumtimer FROM aktiviteter a "_
            &" LEFT JOIN timer t ON (t.taktivitetid = a.id) WHERE job = "& strJobid &" AND brug_fasttp = 1 AND Tfaktim = 93 GROUP BY taktivitetid"
            oRec2.open fakbarLearTimer_fastprisSQL, oConn, 3
            while not oRec2.EOF
            fastPris_laer_Total_TP = fastPris_laer_Total_TP + (oRec2("fasttp")*oRec2("sumtimer")*1)
            totalTimerFastpris_lear = totalTimerFastpris_lear + (oRec2("sumtimer")*1)
            oRec2.movenext
            wend
            oRec2.close
            'response.Write "<br><br><br>HER1 "& fastPris_laer_Total_TP & "<br>" & "her2 " & totalTimerFastpris_lear

            lear_ikkeFP_sum = 0
            lear_ikkeFP_sumtimer = 0
            fakbarLearTimer_ikkefastprisSQL = "SELECT sum(timer*timepris) as sumLearTimepris_ikkeFP, sum(timer) as sumLearTimer_ikkeFP FROM timer LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE tjobnr = '"& strJobnr &"' AND Tfaktim = 93 AND brug_fasttp = 0 GROUP BY taktivitetid"
            'response.Write fakbarLearTimer_ikkefastprisSQL
            oRec2.open fakbarLearTimer_ikkefastprisSQL, oConn, 3
            while not oRec2.EOF
            lear_ikkeFP_sum = lear_ikkeFP_sum + oRec2("sumLearTimepris_ikkeFP")
            lear_ikkeFP_sumtimer = lear_ikkeFP_sumtimer + oRec2("sumLearTimer_ikkeFP")
            oRec2.movenext
            wend
            oRec2.close 
            'response.Write "<br><br><br>HER1 "& lear_ikkeFP_sum & "<br>" & "her2 " & lear_ikkeFP_sumtimer

            laertimepristotal = (fastPris_laer_Total_TP*1 + lear_ikkeFP_sum*1)
            laertimertotal = (totalTimerFastpris_lear*1 + lear_ikkeFP_sumtimer*1)
            learTimer_oms = laertimepristotal

            if learTimer_oms <> 0 then 
            learTimer_oms = learTimer_oms
            laarTimeprisGNS = (learTimer_oms*1/laertimertotal*1) 
            else 
            learTimer_oms = 0
            laarTimeprisGNS = 0 
            end if


            fastPris_easyReg_Total_TP = 0
            totalTimerFastpris_easyReg = 0
            fakbareasyRegTimer_fastprisSQL = "SELECT a.id, brug_fasttp, fasttp, navn, SUM(t.timer) AS sumtimer FROM aktiviteter a "_
            &" LEFT JOIN timer t ON (t.taktivitetid = a.id) WHERE job = "& strJobid &" AND brug_fasttp = 1 AND Tfaktim = 1 AND Timer < 0.25 GROUP BY taktivitetid"
            oRec2.open fakbareasyRegTimer_fastprisSQL, oConn, 3
            while not oRec2.EOF
            fastPris_easyReg_Total_TP = fastPris_easyReg_Total_TP + (oRec2("fasttp")*oRec2("sumtimer")*1)
            totalTimerFastpris_easyReg = totalTimerFastpris_easyReg + (oRec2("sumtimer")*1)
            oRec2.movenext
            wend
            oRec2.close
            'response.Write "<br><br><br>HER1 "& fastPris_easyReg_Total_TP & "<br>" & "her2 " & totalTimerFastpris_easyReg

            easyReg_ikkeFP_sum = 0
            easyReg_ikkeFP_sumtimer = 0
            fakbareasyRegTimer_ikkefastprisSQL = "SELECT sum(timer*timepris) as sumeasyRegTimepris_ikkeFP, sum(timer) as sumeasyRegTimer_ikkeFP FROM timer LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE tjobnr = '"& strJobnr &"' AND Tfaktim = 1 AND Timer < 0.25 AND brug_fasttp = 0 GROUP BY taktivitetid"
            'response.Write fakbarLearTimer_ikkefastprisSQL
            oRec2.open fakbareasyRegTimer_ikkefastprisSQL, oConn, 3
            while not oRec2.EOF
            easyReg_ikkeFP_sum = easyReg_ikkeFP_sum + oRec2("sumeasyRegTimepris_ikkeFP")
            easyReg_ikkeFP_sumtimer = easyReg_ikkeFP_sumtimer + oRec2("sumeasyRegTimer_ikkeFP")
            oRec2.movenext
            wend
            oRec2.close
            'response.Write "<br><br><br>HER1 "& easyReg_ikkeFP_sum & "<br>" & "her2 " & easyReg_ikkeFP_sumtimer 

            easyRegtimepristotal = (fastPris_easyReg_Total_TP*1 + easyReg_ikkeFP_sum*1)
            easyRegtimertotal = (totalTimerFastpris_easyReg*1 + easyReg_ikkeFP_sumtimer*1)
            easyRegTimer_oms = easyRegtimepristotal

            if easyRegTimer_oms <> 0 then 
            easyRegTimer_oms = easyRegTimer_oms
            easyRegTimeprisGNS = (easyRegTimer_oms*1/easyRegtimertotal*1)
            else 
            easyRegTimer_oms = 0
            easyRegTimeprisGNS = 0 
            end if


            fastPris_ikkeFakbar_Total_TP = 0
            totalTimerFastpris_ikkeFakbar = 0
            ikkefakbarTimer_fastprisSQL = "SELECT a.id, brug_fasttp, fasttp, navn, SUM(t.timer) AS sumtimer FROM aktiviteter a "_
            &" LEFT JOIN timer t ON (t.taktivitetid = a.id) WHERE job = "& strJobid &" AND brug_fasttp = 1 AND Tfaktim = 2 GROUP BY taktivitetid"
            oRec2.open ikkefakbarTimer_fastprisSQL, oConn, 3
            while not oRec2.EOF
            fastPris_ikkeFakbar_Total_TP = fastPris_ikkeFakbar_Total_TP + (oRec2("fasttp")*oRec2("sumtimer")*1)
            totalTimerFastpris_ikkeFakbar = totalTimerFastpris_ikkeFakbar + (oRec2("sumtimer")*1)
            oRec2.movenext
            wend
            oRec2.close
            'response.Write "<br><br><br>HER1 "& fastPris_ikkeFakbar_Total_TP & "<br>" & "her2 " & totalTimerFastpris_ikkeFakbar


            ikkefakbar_ikkeFP_sum = 0
            ikkefakbar_ikkeFP_sumtimer = 0
            ikkefakbarTimer_ikkefastprisSQL = "SELECT sum(timer*timepris) as sumikkefakbarTimepris_ikkeFP, sum(timer) as sumikkefakbarTimer_ikkeFP FROM timer LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE tjobnr = '"& strJobnr &"' AND Tfaktim = 1 AND Timer < 0.25 AND brug_fasttp = 0 GROUP BY taktivitetid"
            'response.Write fakbarLearTimer_ikkefastprisSQL
            oRec2.open ikkefakbarTimer_ikkefastprisSQL, oConn, 3
            while not oRec2.EOF
            ikkefakbar_ikkeFP_sum = ikkefakbar_ikkeFP_sum + oRec2("sumikkefakbarTimepris_ikkeFP")
            ikkefakbar_ikkeFP_sumtimer = ikkefakbar_ikkeFP_sumtimer + oRec2("sumikkefakbarTimer_ikkeFP")
            oRec2.movenext
            wend
            oRec2.close
            'response.Write "<br><br><br>HER1 "& ikkefakbar_ikkeFP_sum & "<br>" & "her2 " & ikkefakbar_ikkeFP_sumtimer

            ikkefakbartimepristotal = (fastPris_ikkeFakbar_Total_TP*1 + ikkefakbar_ikkeFP_sum*1)
            ikkefakbartimertotal = (totalTimerFastpris_ikkeFakbar*1 + ikkefakbar_ikkeFP_sumtimer*1)
            ikkeFabbarTimer_oms = ikkefakbartimepristotal
            if ikkeFabbarTimer_oms <> 0 then 
            ikkeFabbarTimer_oms = ikkeFabbarTimer_oms
            ikkefakbarGNStimepris = (ikkeFabbarTimer_oms*1/ikkefakbartimertotal*1)
            else 
            ikkeFabbarTimer_oms = 0
            ikkefakbarGNStimepris = 0 
            end if

            if func = "red" then
            
                strSQLEval = "SELECT eval_jobid, eval_evalvalue, eval_jobvaluesuggested, eval_comment, eval_diff, eval_suggested_hours, eval_suggested_hourly_rate, eval_fakbartimer, eval_fakbartimepris, " _ 
                & "ubemandet_maskine_timer, ubemandet_maskine_timePris, laer_timer, laer_timepris, easy_reg_timer, easy_reg_timepris, ikke_fakbar_tid_timer, ikke_fakbar_tid_timepris " _
                & "FROM eval WHERE eval_jobid = "& jobid_til_eval
                'response.Write strSQLEval
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
                    strEval_svendetimer_sug = oRec2("eval_fakbartimer")
                    strEval_svendetimePris_sug = oRec2("eval_fakbartimepris")
                    ubemandet_maskine_timer = oRec2("ubemandet_maskine_timer")
                    ubemandet_maskine_timePris = oRec2("ubemandet_maskine_timePris")
                    laer_timer = oRec2("laer_timer")
                    laer_timepris = oRec2("laer_timepris")
                    easy_reg_timer = oRec2("easy_reg_timer")
                    easy_reg_timepris = oRec2("easy_reg_timepris")
                    ikke_fakbar_tid_timer = oRec2("ikke_fakbar_tid_timer")
                    ikke_fakbar_tid_timepris = oRec2("ikke_fakbar_tid_timepris")                   
                end if
                oRec2.close

                if strEval_svendetimer_sug <> 0 then strEval_svendetimer_sug = strEval_svendetimer_sug else strEval_svendetimer_sug = 0 end if
                if strEval_svendetimePris_sug <> 0 then strEval_svendetimePris_sug = strEval_svendetimePris_sug else strEval_svendetimePris_sug = 0 end if
                if ubemandet_maskine_timer <> 0 then ubemandet_maskine_timer = ubemandet_maskine_timer else ubemandet_maskine_timer = 0 end if
                if ubemandet_maskine_timePris <> 0 then ubemandet_maskine_timePris = ubemandet_maskine_timePris else ubemandet_maskine_timePris = 0 end if
                if laer_timer <> 0 then laer_timer = laer_timer else laer_timer = 0 end if
                if laer_timepris <> 0 then laer_timepris = laer_timepris else laer_timepris = 0 end if
                if easy_reg_timer <> 0 then easy_reg_timer = easy_reg_timer else easy_reg_timer = 0 end if
                if easy_reg_timepris <> 0 then easy_reg_timepris = easy_reg_timepris else easy_reg_timepris = 0 end if
                if ikke_fakbar_tid_timer <> 0 then ikke_fakbar_tid_timer = ikke_fakbar_tid_timer else ikke_fakbar_tid_timer = 0 end if
                if ikke_fakbar_tid_timepris <> 0 then ikke_fakbar_tid_timepris = ikke_fakbar_tid_timepris else ikke_fakbar_tid_timepris = 0 end if

            else 'Ikke rediger

                valueBarSEL4 = ""
                valueBarSEL3 = ""
                valueBarSEL2 = ""
                valueBarSEL1 = ""
                strEval_comment = ""
                strEval_suggested_hours = totalTimer
                strEval_suggested_houry_rate = TotalGNStimepris
                strEval_svendetimer_sug = svendetimerTot
                strEval_svendetimePris_sug = svendetimeprisGNS
                ubemandet_maskine_timer = ubemandetMaskineTimerTot
                ubemandet_maskine_timePris = ubemandetMaskineTimeprisGNS
                laer_timer = learTimer
                laer_timepris = learTimepris
                easy_reg_timer = easyRegTimer
                easy_reg_timepris = easyRegTimepris
                ikke_fakbar_tid_timer = ikkeFabbarTimer
                ikke_fakbar_tid_timepris = ikkeFabbarTimepris

            end if

             
        %>

<script src="js/eval_jav5.js" type="text/javascript"></script>


<div class="wrapper"><br /><br />
<!--  
    <div class="content">
      -->

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title">Projekt Evaluering</h3>
                <div class="portlet-body">
                   <b>Kunde & Job:</b><br />
                   <%response.Write kkundenavn & " ("& kkundenr &")<br>"& strJobnavn & " ("& strJobnr &")<br> Bruttooms.: " & formatnumber(strBruttooms, 2) & "<br> <a href=""#"" onclick=""Javascript:window.open('../timereg/jobprintoverblik.asp?menu=job&id="&strJobid&"', '', 'width=1200,height=700,resizable=yes,scrollbars=yes')"">Timefordeling >></a>" %>
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
                        <div class="col-lg-2"></div>
                        <div class="col-lg-8">
                            <table>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:5px">Fastpris:</td>
                                    <td style="width:443px"></td>                                    
                                    <td style="width:125px; padding-right:10px;"><input type="text" name="fastprisjob_oms" id="fastprisjob_oms" class="form-control input-small" value="<%=strBruttooms %>" readonly/></td>
                                </tr>
                            </table>
                        </div>
                    </div>
                    
                    <%end if %>
                   <div class="row">
                        <div class="col-lg-2"></div>
                        <div class="col-lg-8">
                            <table>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timer brugt</td>
                                    <td style="width:100px; padding-right:10px; padding-bottom:10px"><input type="text" name="LBN_totaltimer" id="LBN_totaltimer" class="form-control input-small" readonly value="<%=formatnumber(totalTimer, 2) %>" /></td>
                                    <td style="padding-right:10px; padding-bottom:10px">*</td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" name="LBN_gns_timepris" id="LBN_gns_timepris" class="form-control input-small" readonly value="<%=formatnumber(TotalGNStimepris, 2)%>"/></td>
                                    <td style="padding-right:10px; padding-bottom:10px">=</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small" readonly value="<%=formatnumber(omsResultat, 2) %>"/>
                                        <input type="hidden" name="for_oms_result" id="for_oms_result" class="form-control input-small" readonly value="<%=replace(formatnumber(omsResultat, 2), ".", "") %>"/></td>
                                </tr>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:26px; padding-bottom:10px">Fakturerbar svendetimer</td>
                                    <td style="width:100px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small" value="<%=formatnumber(svendetimertotal, 2) %>" readonly />
                                    <td style="padding-right:10px; padding-bottom:10px">*</td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small opdaterOms" value="<%=formatnumber(totalGNSsvendeTimepris, 2) %>" readonly /></td>
                                    <td style="padding-right:10px; padding-bottom:10px">=</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small" readonly value="<%=formatnumber(svendetimer_oms, 2) %>"/></td>
                                </tr>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:26px; padding-bottom:10px">Ubemandet mask. tid</td>
                                    <td style="width:100px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small" value="<%=formatnumber(maskinetimertotal, 2) %>" readonly />
                                    <td style="padding-right:10px; padding-bottom:10px">*</td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small opdaterOms" value="<%=formatnumber(maskine_GNS_timepris, 2) %>" readonly /></td>
                                    <td style="padding-right:10px; padding-bottom:10px">=</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small" readonly value="<%=formatnumber(ubemandetMaskine_oms, 2) %>"/></td>
                                </tr>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:26px; padding-bottom:10px">Lærlinge timer</td>
                                    <td style="width:100px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small" value="<%=formatnumber(laertimertotal, 2) %>" readonly />
                                    <td style="padding-right:10px; padding-bottom:10px">*</td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small opdaterOms" value="<%=formatnumber(laarTimeprisGNS, 2) %>" readonly /></td>
                                    <td style="padding-right:10px; padding-bottom:10px">=</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small" readonly value="<%=formatnumber(learTimer_oms, 2) %>"/></td>
                                </tr>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:26px; padding-bottom:10px">Easy reg. timer</td>
                                    <td style="width:100px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small" value="<%=formatnumber(easyRegtimertotal, 2) %>" readonly />
                                    <td style="padding-right:10px; padding-bottom:10px">*</td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small opdaterOms" value="<%=formatnumber(easyRegTimeprisGNS, 2) %>" readonly /></td>
                                    <td style="padding-right:10px; padding-bottom:10px">=</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small" readonly value="<%=formatnumber(easyRegTimer_oms, 2) %>"/></td>
                                </tr>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:26px; padding-bottom:10px">Ikke fakturerbar tid</td>
                                    <td style="width:100px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small" value="<%=formatnumber(ikkefakbartimertotal, 2) %>" readonly />
                                    <td style="padding-right:10px; padding-bottom:10px">*</td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small opdaterOms" value="<%=formatnumber(ikkefakbarGNStimepris, 2) %>" readonly /></td>
                                    <td style="padding-right:10px; padding-bottom:10px">=</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" class="form-control input-small" readonly value="<%=formatnumber(ikkeFabbarTimer_oms, 2) %>"/></td>
                                </tr>
                            </table>
                        </div>
                    </div>                   
                    <br />
                    <div class="row">
                        <div class="col-lg-2"></div>
                        <div class="col-lg-8">
                            <!--<table>
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
                            </table>-->
                            <table>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:26px; padding-bottom:10px">Fakturerbar svendetimer</td>
                                    <td style="width:100px; padding-right:10px; padding-bottom:10px"><input type="text" name="fak_svendetimer" id="fak_svendetimer" class="form-control input-small opdaterOms" value="<%=strEval_svendetimer_sug %>" />
                                    <td style="padding-right:10px; padding-bottom:10px">*</td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" name="fak_svendetimePris" id="fak_svendetimerPris" class="form-control input-small opdaterOms" value="<%=strEval_svendetimePris_sug %>" /></td>
                                    <td style="padding-right:10px; padding-bottom:10px">=</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" id="omregnetOms_svendetimer" class="form-control input-small" readonly/></td>
                                </tr>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:26px; padding-bottom:10px">Ubemandet mask. tid</td>
                                    <td style="width:100px; padding-right:10px; padding-bottom:10px"><input type="text" name="ubemandet_maskine_timer" id="ubemandet_maskine_timer" class="form-control input-small opdaterOms" value="<%=ubemandet_maskine_timer %>" />
                                    <td style="padding-right:10px; padding-bottom:10px">*</td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" name="ubemandet_maskine_timePris" id="ubemandet_maskine_timePris" class="form-control input-small opdaterOms" value="<%=ubemandet_maskine_timePris %>" /></td>
                                    <td style="padding-right:10px; padding-bottom:10px">=</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" id="omregnetOms_ubemandet_maskinetimer" class="form-control input-small" readonly/></td>
                                </tr>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:26px; padding-bottom:10px">Lærlinge timer</td>
                                    <td style="width:100px; padding-right:10px; padding-bottom:10px"><input type="text" name="laer_timer" id="laer_timer" class="form-control input-small opdaterOms" value="<%=laer_timer %>" />
                                    <td style="padding-right:10px; padding-bottom:10px">*</td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" name="laer_timepris" id="laer_timepris" class="form-control input-small opdaterOms" value="<%=laer_timepris %>" /></td>
                                    <td style="padding-right:10px; padding-bottom:10px">=</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" id="omregnet_lear_timer" class="form-control input-small" readonly/></td>
                                </tr>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:26px; padding-bottom:10px">Easy reg. timer</td>
                                    <td style="width:100px; padding-right:10px; padding-bottom:10px"><input type="text" name="easy_reg_timer" id="easy_reg_timer" class="form-control input-small opdaterOms" value="<%=easy_reg_timer %>" />
                                    <td style="padding-right:10px; padding-bottom:10px; padding-bottom:10px">*</td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" name="easy_reg_timepris" id="easy_reg_timepris" class="form-control input-small opdaterOms" value="<%=easy_reg_timepris %>" /></td>
                                    <td style="padding-right:10px; padding-bottom:10px; padding-bottom:10px">=</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" id="omregnet_easy_reg_timer" class="form-control input-small" readonly/></td>
                                </tr>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:26px; padding-bottom:10px">Ikke fakturerbar tid</td>
                                    <td style="width:100px; padding-right:10px; padding-bottom:10px"><input type="text" name="ikke_fakbar_tid_timer" id="ikke_fakbar_tid_timer" class="form-control input-small opdaterOms" value="<%=ikke_fakbar_tid_timer %>" />
                                    <td style="padding-right:10px; padding-bottom:10px">*</td>
                                    <td style="color:black; padding-right:10px; padding-bottom:10px">Timepris</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" name="ikke_fakbar_tid_timepris" id="ikke_fakbar_tid_timepris" class="form-control input-small opdaterOms" value="<%=ikke_fakbar_tid_timepris %>" /></td>
                                    <td style="padding-right:10px; padding-bottom:10px">=</td>
                                    <td style="width:125px; padding-right:10px; padding-bottom:10px"><input type="text" id="omregnet_ikke_fakbar" class="form-control input-small" readonly/></td>
                                </tr>
                            </table>
                        </div>
                    </div>
                    <br />
                    <div class="row">
                        <div class="col-lg-2"></div>
                        <div class="col-lg-8">
                            <table>
                                <tr>
                                    <td style="width:65px"></td>
                                    <td style="color:black; padding-right:56px">Forslået omsætning</td>
                                    <td style="width:100px; padding-right:10px"><input type="text" name="forslaet_timer_total" id="forslaet_timer_total" class="form-control input-small opdaterOms" value="" readonly />
                                    <td style="width:80px"></td>
                                    <td style="width:125px; padding-right:10px"><input type="text" name="forslaet_timepris_total" id="forslaet_timepris_total" class="form-control input-small opdaterOms" value="" readonly /></td>
                                    <td style="width:19px"></td>
                                    <td style="width:125px; padding-right:10px"><input type="text" name="forslaet_oms" id="forslaet_oms" class="form-control input-small" readonly /></td>
                                    <td><input type="hidden" name="eval_lbn_forskel" id="eval_lbn_forskel" class="form-control input-small" readonly/></td>
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