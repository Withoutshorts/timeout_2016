

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
    
    if len(trim(request("FM_jobans"))) <> 0 then
    FM_jobans = request("FM_jobans")
    else
    FM_jobans = 0
    end if

    select case FM_jobans
    case 0
    jobansSQLfilter = ""
    case else
    jobansSQLfilter = " AND (j.jobans1 = "& FM_jobans &" OR j.jobans2 = "& FM_jobans & " OR j.jobans3 = "& FM_jobans & " OR j.jobans4 = "& FM_jobans & " OR j.jobans5 = "& FM_jobans &")" 
    end select 
    
      
    if len(trim(request("FM_fomr"))) <> 0 then
    FM_fomr = request("FM_fomr")
    else
    FM_fomr = 0
    end if 

    select case FM_fomr
    case 0
    fomrSQLfilter = ""
    case else
    fomrSQLfilter = " AND f.for_fomr = "& FM_fomr
    end select
    
    if len(request("FM_kunde")) <> 0 then
	FM_kunde = request("FM_kunde")
	else
	FM_kunde = 0
	end if
    
    select case FM_kunde
    case 0
    kundeSQLfilter = ""
    case else
    kundeSQLfilter = " AND j.jobknr = "& FM_kunde
    end select
    
    if len(request("fradato")) <> 0 then
    fradato = request("fradato")
    else
    fradato = day(now) &"-"& month(now) &"-"& year(now)
    end if
    
    if len(request("tildato")) <> 0 then
    tildato = request("tildato")
    else
    tildato = day(now) &"-"& month(now) &"-"& year(now)  
    end if

    sqlFraDato = year(fradato) &"-"& month(fradato) &"-"& day(fradato)
    sqlTilDato = year(tildato) &"-"& month(tildato) &"-"& day(tildato)

    if len(request("datoType")) <> 0 then
    datoType = request("datoType")
    else
    datoType = 0
    end if

    select case datoType
    case 0
    sqlDato = ""
    chbDato1 = "checked"
    chbDato2 = ""
    chbDato3 = ""
    case 1
    sqlDato = " AND j.jobstartdato BETWEEN '"& sqlFraDato &"' AND '"& sqlTilDato &"'"
    chbDato1 = ""
    chbDato2 = "checked"
    chbDato3 = ""
    case 2
    sqlDato = " AND j.lukkedato BETWEEN '"& sqlFraDato &"' AND '"& sqlTilDato &"'"
    chbDato1 = ""
    chbDato2 = ""
    chbDato3 = "checked"
    end select

    if len(request("statusType")) <> 0 then
    statusType = request("statusType")
    else
    statusType = 5
    end if

    select case statusType

    case 5
    
    statusSEL5 = "selected"
    statusSEL1 = ""
    statusSEL2 = ""
    statusSEL3 = ""
    statusSEL4 = ""
    statusSEL0 = ""
    case 1
    'statusSQLfilter = ""
    statusSEL5 = ""
    statusSEL1 = "selected"
    statusSEL2 = ""
    statusSEL3 = ""
    statusSEL4 = ""
    statusSEL0 = ""
    case 2
    'statusSQLfilter = ""
    statusSEL5 = ""
    statusSEL1 = ""
    statusSEL2 = "selected"
    statusSEL3 = ""
    statusSEL4 = ""
    statusSEL0 = ""
    case 3
    'statusSQLfilter = ""
    statusSEL5 = ""
    statusSEL1 = ""
    statusSEL2 = ""
    statusSEL3 = "selected"
    statusSEL4 = ""
    statusSEL0 = ""
    case 4
    'statusSQLfilter = ""
    statusSEL5 = ""
    statusSEL1 = ""
    statusSEL2 = ""
    statusSEL3 = ""
    statusSEL4 = "selected"
    statusSEL0 = ""
    case 0
    'statusSQLfilter = ""
    statusSEL5 = ""
    statusSEL1 = ""
    statusSEL2 = ""
    statusSEL3 = ""
    statusSEL4 = ""
    statusSEL0 = "selected"
    case else
   
    end select


    statusSQLfilter = " AND j.jobstatus = "& statusType

     call akttyper2009(2)

%>

<%call menu_2014 %> 

<script src="js/eval_jav6.js" type="text/javascript"></script>
<div class="wrapper">
    <div class="content">


        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Projekt Evaluering</u></h3>
                <div class="portlet-body">


                    <div class="well">
                        <form action="eval_liste.asp" method="post">

                        <div class="row">
                            <div class="col-lg-12">
                                <h4 class="panel-title-well"><%=medarb_txt_021 %></h4>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-lg-4">Kunde:</div>
                            <div class="col-lg-4">Periode:</div>                           
                        </div>
                        <div class="row">
                            <div class="col-lg-4">
                                <select name="FM_kunde" class="form-control input-small" onchange="submit();">
                                    <option value="0">Ignorer</option>
                                    <%
                                        strSQLKunder = "SELECT Kkundenavn, Kkundenr, Kid, useasfak FROM kunder WHERE ketype <> 'e' AND (useasfak = 0 OR useasfak = 1 OR useasfak = 6) ORDER BY useasfak, Kkundenavn"
                                        oRec.open strSQLKunder, oConn, 3
                                        while not oRec.EOF
                                        
                                        if cdbl(FM_kunde) = cdbl(oRec("Kid")) then
                                        FM_kundeSEL = "SELECTED"
                                        else
                                        FM_kundeSEL = ""
                                        end if
                                        %>
                                            <option value="<%=oRec("Kid") %>" <%=FM_kundeSEL %>><%=oRec("kKundenavn") %></option>
                                        <%
                                        oRec.movenext
                                        wend
                                        oRec.close                             
                                    %>
                                </select>
                            </div>

                            <div class="col-lg-2">
                                <div class='input-group date' id='datepicker_stdato'>
                                <input name="fradato" type="text" class="form-control input-small" placeholder="dd-mm-yyyy" value="<%=fradato %>" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>
                            <div class="col-lg-2">
                                <div class='input-group date' id='datepicker_stdato'>
                                <input name="tildato" type="text" class="form-control input-small" placeholder="dd-mm-yyyy" value="<%=tildato %>" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>
                            <div class="col-lg-4">
                                <input type="radio" name="datoType" value="0" <%=chbDato1 %> /> Ignorer
                                &nbsp<input type="radio" name="datoType" value="1" <%=chbDato2 %> /> Jobstart-dato                                
                                &nbsp<input type="radio" name="datoType" value="2" <%=chbDato3 %> /> Lukket-dato
                            </div>
                        </div>

                        <div class="row">
                            <div class="col-lg-4">Jobansvarlig:</div>
                            <div class="col-lg-4">Forretningsområde:</div>
                            <div class="col-lg-2">Status:</div>
                        </div>
                        <div class="row">
                            <div class="col-lg-4">
                                <select name="FM_jobans" class="form-control input-small" onchange="submit();">
                                    <option value="0">Ignorer</option>

                                    <%
                                        jobansSEL = ""
                                        strJobans = "SELECT mnavn, mid, init FROM medarbejdere WHERE mansat = 1 ORDER BY mnavn"
                                        oRec.open strJobans, oConn, 3
                                        while not oRec.EOF
                                        
                                        if cdbl(FM_jobans) = cdbl(oRec("mid")) then
                                        jobansSEL = "SELECTED"
                                        else
                                        jobansSEL = ""
                                        end if

                                        if len(trim(oRec("init"))) <> 0 then
                                        strInit = " ["& oRec("init") &"]"
                                        else
                                        strInit = ""
                                        end if

                                        %>
                                        <option value="<%=oRec("mid") %>" <%=jobansSEL %>><%=oRec("mnavn") & strInit %></option>
                                        <%
                                        oRec.movenext
                                        wend
                                        oRec.close
                                    %>

                                </select>
                            </div>

                            <div class="col-lg-4">
                                <select name="FM_fomr" class="form-control input-small" onchange="submit();">
                                    <option value="0">Ignorer</option>
                                    <% 
                                        strFomr = "SELECT id, navn FROM fomr ORDER BY navn"
                                        oRec.open strFomr, oConn, 3
                                        while not oRec.EOF

                                        if cdbl(FM_fomr) = cdbl(oRec("id")) then
                                        fomrSEL = "SELECTED"
                                        else
                                        fomrSEL = ""
                                        end if

                                    %>
                                        <option value="<%=oRec("id") %>" <%=fomrSEL %>><%=oRec("navn") %></option>
                                    <%
                                        oRec.movenext
                                        wend
                                        oRec.close
                                    %>
                                </select>
                            </div>
                            
                            <div class="col-lg-2">
                                <select name="statusType" class="form-control input-small" onchange="submit();">
                                    <option value="5" <%=statusSEL5 %>>Evaluering</option>
                                    <option value="1" <%=statusSEL1 %>><%=job_txt_221 %></option>
                                    <option value="2" <%=statusSEL2 %>><%=job_txt_222 %></option>
                                    <option value="3" <%=statusSEL3 %>><%=job_txt_223 %></option>
                                    <option value="4" <%=statusSEL4 %>><%=job_txt_224 %></option>
                                    <option value="0" <%=statusSEL0 %>><%=job_txt_225 %></option>
                                </select>
                            </div>
                                
                            <div class="col-lg-2">
                                <button type="submit" class="btn btn-secondary btn-sm pull-right"><b> Søg </b></button>
                            </div>
                                                    
                        </div>

                        </form>
                    </div>
                    
                    <script src="js/eval_liste_jav.js" type="text/javascript"></script>
                    <table id="eval_liste" class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th>Kunde</th>
                                <th>Projekt</th>
                                <th>Forretningsomr.</th>
                                <th>Eval</th>
                                <th>Værdi</th>
                                <%'if statusType = 5 then %>
                                <th>Pris</th>
                                <th>Forslået pris</th>
                                <th>Forskel</th>
                                <th>Timer <br /><span style="font-size:9px;">Real. - Forslået</span></th>
                                <%'end if %>
                                <th>Kommentar</th>
                            </tr>
                        </thead>
                        <%
                        'if statusType = 5 then 
                        %>
                        <tbody>

                            <%                               
                                evalTotal = 0
                                evalPrisTotal = 0
                                evalFPrisTotal = 0
                                evalForskelTotal = 0
                                eval_suggested_hoursTotal = 0
                                antal = 0
                                strEval = "SELECT j.id AS jobid, j.jobnavn, k.kkundenavn, jobnr, "_
                                &" e.eval_jobid, e.eval_evalvalue, eval_comment, eval_jobvaluesuggested, eval_diff, e.eval_original_price, j.jo_bruttooms, j.fastpris, eval_suggested_hours "_
                                &" FROM job AS j "_
                                &" LEFT JOIN eval as e ON (e.eval_jobid = j.id) "_
                                &" LEFT JOIN fomr_rel as f ON (f.for_jobid = j.id) "_
                                &" LEFT JOIN kunder as k ON (k.Kid = j.jobknr) WHERE risiko >= 0 " & jobansSQLfilter & fomrSQLfilter & kundeSQLfilter & sqlDato & statusSQLfilter & " GROUP BY j.id ORDER BY kkundenavn, jobnavn"  
                                
                                'response.write strEval
                                'response.flush
                                
                                oRec.open strEval, oConn, 3 
                                while not oRec.EOF

                                jobnr = oRec("jobnr")

                                if oRec("eval_evalvalue") <> 0 then
                                eval_evalvalue = oRec("eval_evalvalue")
                                else 
                                eval_evalvalue = 0
                                end if

                                eval_valueTXT = ""
                                select case cdbl(eval_evalvalue)
                                case 4
                                eval_valueTXT = "God +"
                                eval_valueTXTcol = "cornflowerblue"
                                case 3
                                eval_valueTXT = "God"
                                eval_valueTXTcol = "yellowgreen"
                                case 2
                                eval_valueTXT = "Middel"
                                eval_valueTXTcol = "orange"
                                case 1
                                eval_valueTXT = "Dårlig"
                                eval_valueTXTcol = "red"
                                case else
                                eval_valueTXT = "+"
                                eval_valueTXTcol = "#999999"
                                end select
                                  
                                if antal = 0 then
                                evalTotal = oRec("eval_evalvalue")
                                else
                                evalTotal = evalTotal + oRec("eval_evalvalue")
                                end if

                                'if antal = 0 then
                                 '   if oRec("fastpris") = 1 then
                                 '   evalPrisTotal = oRec("jo_bruttooms")
                                 '   else
                                 '   evalPrisTotal = oRec("eval_original_price")
                                 '   end if
                                'else
                                    if oRec("fastpris") = 1 then
                                    evalPrisTotal = evalPrisTotal + oRec("jo_bruttooms")
                                    else
                                        if oRec("eval_original_price") <> 0 then
                                        evalPrisTotal = evalPrisTotal + oRec("eval_original_price")
                                        end if
                                    end if
                                'end if

                                'if antal = 0 then
                                'evalFPrisTotal = oRec("eval_jobvaluesuggested")
                                'else
                                if oRec("eval_jobvaluesuggested") <> 0 then
                                evalFPrisTotal = evalFPrisTotal + oRec("eval_jobvaluesuggested")
                                end if
                                'end if
                                 
                               

                                'strJob = "SELECT j.jobnavn, k.kKundenavn FROM job as j LEFT JOIN kunder as k ON (k.Kid = j.jobknr) WHERE j.id = "& oRec("eval_jobid") 
                                'oRec2.open strJob, oConn, 3
                                'if not oRec2.EOF then
                                jobnavn = oRec("jobnavn")
                                kundenavn = oRec("kKundenavn")
                                'end if
                                'oRec2.close 
                            %>

                            <tr>
                                <td style="width:200px;"><%=kundenavn %></td>
                                <td style="width:200px;"><%=jobnavn & " ("& jobnr &")" %></td>
                                <td style="width:200px">
                                    <%
                                        
                                        strFomr = "SELECT r.for_fomr, f.navn FROM fomr_rel as r LEFT JOIN fomr as f ON (f.id = r.for_fomr) WHERE r.for_jobid = "& oRec("jobid") & " GROUP BY for_fomr" 
                                        oRec2.open strFomr, oConn, 3
                                        while not oRec2.EOF                                                                                
                                        response.Write oRec2("navn") & ", "                                                                             
                                        oRec2.movenext
                                        wend
                                        oRec2.close
                                    %>
                                </td>
                                <td style="width:50px;"><a href="#" onclick="Javascript:window.open('eval.asp?func=red&jobid_til_eval=<%=oRec("jobid") %>', '', 'width=1200,height=700,resizable=yes,scrollbars=yes')" style="color:<%=eval_valueTXTcol%>;"><%=eval_valueTXT %></a></td>
                               
                                 <td style="width:15px; text-align:center"><% if oRec("eval_evalvalue") <> 0 then %> <%=oRec("eval_evalvalue") %> <% end if %></td>

                                <td style="text-align:right">
                                     <%
                                    if oRec("fastpris") = 1 then 
                                    response.Write formatnumber(oRec("jo_bruttooms"),2)
                                    else

                                         if oRec("eval_original_price") <> 0 then
                                         eval_original_price = oRec("eval_original_price")
                                         else
                                         eval_original_price = 0
                                         end if
                                    response.Write formatnumber(eval_original_price ,2)
                                    end if
                                     %>
                                </td>

                                <%if len(trim(oRec("eval_jobvaluesuggested"))) <> 0 then
                                  eval_jobvaluesuggested = oRec("eval_jobvaluesuggested")
                                else
                                  eval_jobvaluesuggested = 0
                                end if    %>

                                <td style="text-align:right"><%=formatnumber(eval_jobvaluesuggested,2) %></td>


                                 <%
                                if oRec("eval_diff") > 0 then
                                    forskelColor = "green"
                                    eval_diff = oRec("eval_diff") 
                                else
                                    if oRec("eval_diff") < 0 then
                                    forskelColor = "red"
                                    eval_diff = oRec("eval_diff") 
                                    else
                                    forskelColor = "#999999"
                                    eval_diff = 0 
                                    end if
                                end if
                                %>
                                <td style="color:<%=forskelColor%>; text-align:right"><%=formatnumber(eval_diff,2) %></td>


                                <%StrSQLTimerTOT = "SELECT sum(timer) as totTimer FROM timer WHERE tjobnr = '"& oRec("jobnr") &"' AND ("& aty_sql_realhours &")"
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
                                   
                                if oRec("eval_suggested_hours") <> 0 then
                                eval_suggested_hours = oRec("eval_suggested_hours")
                                else
                                eval_suggested_hours = 0
                                end if
                               
                                eval_suggested_hoursBal = (totalTimer*1 - (eval_suggested_hours*1)) 
                                eval_suggested_hoursTotal = eval_suggested_hoursTotal + (eval_suggested_hoursBal)      
                                %>

                                   <td style="text-align:right; white-space:nowrap;"><%=formatnumber(eval_suggested_hoursBal, 2) %><br /><span style="font-size:9px; line-height:10px;"><%=formatnumber(totalTimer,2) %> - <%=formatnumber(eval_suggested_hours,2)%>
                                   </span></td>
                                
                               
                                
                                <td style="width:250px"><%=oRec("eval_comment") %></td>
                            </tr>
                            <%
                                antal = antal + 1
                                oRec.movenext
                                wend
                                oRec.close 
                            %>
                        </tbody>
                        <%'if statusType = 5 then 
                            
                            if evalPrisTotal <> 0 then
                            evalPrisTotal = evalPrisTotal
                            else
                            evalPrisTotal = 0
                            end if
                            
                            %>
                        <tfoot>
                            <tr>
                                <th colspan="4" style="border-right:hidden">Total<br /><span style="font-size:10px">Gennemsnit</span></th> 
                                <th style="text-align:center; border-right:hidden"><%if evaltotal <> 0 then %><%=evalTotal %> <%end if %> <br /> <span style="font-size:10px"><%if evalTotal <> 0 and antal <> 0 then %> (<%=formatnumber((evalTotal/antal),2) %>) <% end if %></span></th>
                                <th style="border-right:hidden; text-align:right"><%=formatnumber(evalPrisTotal,2) %></th>
                                <th style="border-right:hidden; text-align:right"><%=formatnumber(evalFPrisTotal,2) %></th>
                                <%
                                evalForskelTotal = evalPrisTotal - evalFPrisTotal

                                if evalForskelTotal <> 0 then
                                    evalForskelTotal = evalForskelTotal
                                else
                                    evalForskelTotal = 0
                                end if

                                if evalForskelTotal > 0 then
                                forskelTotalColor = "yellowgreen"
                                else
                                    if evalForskelTotal = 0 then
                                    forskelTotalColor = "#999999"
                                    else
                                    forskelTotalColor = "red"
                                    end if
                                end if

                                 if eval_suggested_hoursTotal <> 0 then
                                    eval_suggested_hoursTotal = eval_suggested_hoursTotal
                                else
                                    eval_suggested_hoursTotal = 0
                                end if

                                if eval_suggested_hoursTotal > 0 then
                                      forskelTotalHoursColor = "yellowgreen"
                                else

                                        if eval_suggested_hoursTotal = 0 then
                                        forskelTotalHoursColor = "#999999"
                                        else
                                        forskelTotalHoursColor = "red"
                                        end if

                                end if

                               


                                %>
                                <th style="border-right:hidden; color:<%=forskelTotalColor%>; text-align:right"><%=formatnumber(evalForskelTotal,2) %></th>
                                <th style="border-right:hidden; color:<%=forskelTotalHoursColor%>; text-align:right"><%=formatnumber(eval_suggested_hoursTotal,2) %></th>
                                <th></th>
                            </tr>
                        </tfoot>
                        <%'end if %>

                        <%'else %>

                        <!--
                        <tbody>
                            <%
                                strSQL = "SELECT j.id, j.jobnavn, k.kkundenavn FROM job AS j LEFT JOIN kunder AS k ON (k.kid = j.jobknr) LEFT JOIN eval AS e ON (e.eval_jobid <> j.id) LEFT JOIN fomr_rel AS f ON (f.for_jobid = j.id) WHERE j.jobstatus = "& statusType & jobansSQLfilter & fomrSQLfilter & kundeSQLfilter & sqlDato &" GROUP BY j.id"
                                oRec.open strSQL, oConn, 3
                                while not oRec.EOF

                            %>
                            <tr>
                                <td style="width:200px;"><%=oRec("kkundenavn") %></td>
                                <td style="width:200px;"><%=oRec("jobnavn") %></td>
                                <td style="width:200px">
                                    <%
                                        strFomr = "SELECT r.for_fomr, f.navn FROM fomr_rel as r LEFT JOIN fomr as f ON (f.id = r.for_fomr) WHERE r.for_jobid = "& oRec("id") & " GROUP BY for_fomr" 
                                        oRec2.open strFomr, oConn, 3
                                        while not oRec2.EOF                                                                                
                                        response.Write oRec2("navn") & ", "                                                                             
                                        oRec2.movenext
                                        wend
                                        oRec2.close 
                                    %>
                                </td>
                                <td style="text-align:center; width:50px"><a href="#" onclick="Javascript:window.open('eval.asp?func=opr&jobid_til_eval=<%=oRec("id") %>', '', 'width=1200,height=700,resizable=yes,scrollbars=yes')" class=vmenu style="font-size:15px">+</a></td>
                                <td style="width:15px; text-align:center">-</td>
                                <td style="width:250px; text-align:center">-</td>
                            </tr>
                            <%
                                oRec.movenext
                                wend
                                oRec.close 
                            %>
                            
                        </tbody>
                        -->

                        <%'end if %>
                    </table>

                </div>
            </div>
        </div>

    </div>
</div> 

<!--#include file="../inc/regular/footer_inc.asp"-->