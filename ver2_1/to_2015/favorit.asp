<!--#include file="../inc/connection/conn_db_inc.asp"-->

     
<%
 '**** Søgekriterier AJAX **'
        'section for ajax calls

        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")

    case "FN_sogjobogkunde" 

        
                '*** SØG kunde & Job            
                
                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                jobkundesog = request("jq_newfilterval")
                else
                filterVal = 0
                jobkundesog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")

                if filterVal <> 0 then
            
                 lastKid = 0
                
                  'strJobogKunderTxt = strJobogKunderTxt &"<span style=""color:red; font-size:9px; float:right;"" class=""luk_jobsog"">[X]</span>"    
                         


                strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM timereg_usejob AS tu "_ 
                &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                &" WHERE tu.medarb = "& medid &" AND (j.jobstatus = 1 OR j.jobstatus = 3) AND "_
                &" (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '%"& jobkundesog &"%' OR "_
                &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"')  AND kkundenavn <> ''"_
                &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT 50"       
    

                 'response.write "strSQL " &strSQL
                 'response.end

                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
               ' if lastKid <> oRec("kid") then
                'strJobogKunderTxt = strJobogKunderTxt &"<br><br><b>"& oRec("kkundenavn") &" "& oRec("kkundenr") &"</b><br>"
               ' end if 
                 
                'strJobogKunderTxt = strJobogKunderTxt & "<input type=""hidden"" id=""hiddn_job_"& oRec("jid") &""" value="""& oRec("jobnavn") & " ("& oRec("jobnr") &")"">"
                'strJobogKunderTxt = strJobogKunderTxt & "<input type=""checkbox"" class=""chbox_job"" id=""chbox_job_"& oRec("jid") &""" value="& oRec("jid") &"> "& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"<br>" 
                strJobogKunderTxt = strJobogKunderTxt & "<option value="& oRec("jid") &" "& jobSEL &">"& oRec("jobnavn") & " ("& oRec("jobnr") &")"
                
                lastKid = oRec("kid") 
                oRec.movenext
                wend
                oRec.close

              


                    '*** ÆØÅ **'
                    call jq_format(strJobogKunderTxt)
                    strJobogKunderTxt = jq_formatTxt

                    response.write strJobogKunderTxt

                end if



     case "FN_sogakt"

               
                '*** Søg Aktiviteter 
                

                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                aktsog = request("jq_newfilterval")
                else
                filterVal = 0
                aktsog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")
                aktid = request("jq_aktid")
    
                if len(trim(request("jq_jobid"))) <> 0 then        
                jobid = request("jq_jobid")
                else
                jobid = 0
                end if

                'positiv aktivering
                pa = 0
               ' if len(trim(request("jq_pa") )) <> 0 then
                'pa = request("jq_pa") 
                'else
                'pa = 0
               ' end if
        
                'pa = 0
            

                if filterVal <> 0 then
            
                 
    
                'strAktTxt = strAktTxt &"<span style=""color:#999999; font-size:9px; float:right;"" class=""luk_aktsog"">X</span>"    
                         

               if pa = "1" then
               strSQL= "SELECT a.id AS aid, navn AS aktnavn FROM timereg_usejob LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
               &" WHERE tu.medarb = "& usemrn &" AND tu.jobid = "& jobid &" AND aktid <> 0 AND a.navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"   


               else


                        '*** Finder medarbejders projektgrupper 
                        '** Medarbejder projektgrupper **'
                        medarbPGrp = "#0#" 
                        strMpg = "SELECT projektgruppeId, medarbejderId, teamleder FROM progrupperelationer WHERE medarbejderId = "& medid & " GROUP BY projektgruppeId"

                        oRec5.open strMpg, oConn, 3
                        while not oRec5.EOF
                        medarbPGrp = medarbPGrp & ",#"& oRec5("projektgruppeId") &"#"         
        
                        oRec5.movenext
                        wend
                        oRec5.close 


           




               strSQL= "SELECT a.id AS aid, navn AS aktnavn, projektgruppe1, projektgruppe2, projektgruppe3, "_
               &" projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter AS a "_
               &" WHERE a.job = " & jobid & " AND navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"      
    

               
            
                end if

                 'response.write "strSQL " &strSQL
                 'response.end


                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
                 showAkt = 0
                if instr(medarbPGrp, "#"& oRec("projektgruppe1") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe2") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe3") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe4") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe5") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe6") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe7") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe8") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe9") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe10") &"#") <> 0 then
                showAkt = 1
                end if 


                
                
                if cint(showAkt) = 1 then 
                 
                'strAktTxt = strAktTxt & "<input type=""hidden"" id=""hiddn_akt_"& oRec("aid") &""" value="""& oRec("aktnavn") &""">"
                'strAktTxt = strAktTxt & "<input type=""checkbox"" class=""chbox_akt"" id=""chbox_akt_"& oRec("aid") &""" value="& oRec("aid") &"> "& oRec("aktnavn") &"<br>" 
                strAktTxt = strAktTxt & "<option value="& oRec("aid") &" "& optionFcDis &">"& oRec("aktnavn") &" "& fcsaldo_txt &"</option>"             

                end if
                
                oRec.movenext
                wend
                oRec.close

              


                    '*** ÆØÅ **'
                    call jq_format(strAktTxt)
                    strAktTxt = jq_formatTxt


                    response.write strAktTxt

                end if    





        case "tilfoj_akt"

        jobid = request("FM_jobid")
        aktid = request("FM_aktid")
        medid = request("FM_medid_id")

        'oConn.execute("UPDATE timereg_usejob SET favorit = 1 WHERE aktid = "&aktid&"")
        Strtilfojakt = "INSERT INTO timereg_usejob SET jobid = "&jobid &", favorit = 1, aktid ="& aktid& ", medarb = "& medid
        oConn.execute(Strtilfojakt)


        akt_jobid = jobid
        akt_aktid = aktid


        'response.Write Strtilfojakt
        'response.end 

        'Response.redirect "favorit.asp?"

        'strSQLUpdjWiphist = "INSERT INTO wip_historik (dato, restestimat, stade_tim_proc, medid, jobid) VALUES ('"& ddDato &"', "& rest &", "& timerproc &", "& session("mid") &", "& jobid &")"
        'oConn.Execute(strSQLUpdjWiphist)




        case "tilfoj_mat"
    
        mat_jobid = request("mat_jobid")
        mat_aktid = request("mat_aktid")
        mat_antal = request("mat_antal")
        mat_navn = request("mat_navn")
        mat_kobpris = request("mat_kobpris")
        mat_enhed = request("mat_enhed")
        mat_dato = request("mat_dato")
        mat_forbrugsdato = request("mat_forbrugsdato")
        mat_editor = request("mat_editor")
        mat_usrid = request("mat_usrid")
        mat_gruppe = request("mat_gruppe")
        mat_salgspris = request("mat_salgspris")
        mat_bilagsnr = request("mat_bilagsnr")
        mat_valuta = request("mat_valuta")
        mat_varenr = request("mat_varenr") 
        
        strsqlmat = "INSERT INTO materiale_forbrug SET jobid ="& mat_jobid & ", aktid ="& mat_aktid & ", matantal ="& mat_antal & ", matnavn ='"& mat_navn &"', matkobspris ="& mat_kobpris & ", matenhed ='"& mat_enhed &"'" & ", dato ='"& mat_dato &"'" & ", forbrugsdato ='"& mat_forbrugsdato & "'" & ", editor ='"& mat_editor &"'" & ", usrid =" & mat_usrid & ", matgrp ="& mat_gruppe & ", matsalgspris ="& mat_salgspris & ", bilagsnr ='"& mat_bilagsnr & "', valuta ="& mat_valuta & ", matvarenr ='" & mat_varenr & "'"   
        oConn.execute(strsqlmat)
        

        end select                  
        response.end
        end if

        
    %>


<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../timereg/inc/convertDate.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<script src="js/favorit_jav.js"></script>
<!-- <link rel="stylesheet" href="../../bower_components/bootstrap-datepicker/css/datepicker3.css"> -->

<script type="text/javascript" src="js/plugins/flot/jquery.flot.js"></script>
<script type="text/javascript" src="js/demos/flot/stacked-vertical_favorit.js"></script>


<style>

    
    /* The Modal (background) */
    .modal {
        display: none; /* Hidden by default */
        position: fixed; /* Stay in place */
        z-index: 2; /* Sit on top */
        padding-top: 200px; /* Location of the box */
        left: 0;
        top: 0;
        width: 100%; /* Full width */
        height: 100%; /* Full height */
        overflow: auto; /* Enable scroll if needed */
        background-color: rgb(0,0,0); /* Fallback color */
        background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
    }

    /* Modal Content */
    .modal-content {
        background-color: #fefefe;
        margin: auto;
        padding: 20px;
        border: 1px solid #888;
        width: 500px;
        height: 450px;
    }

    .picmodal:hover,
    .picmodal:focus {
    text-decoration: none;
    cursor: pointer;
    }

    .kommodal:hover,
    .kommodal:focus {
    text-decoration: none;
    cursor: pointer;
    }
   
</style>


<div class="wrapper">
<div class="content">

    <%
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if

    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = session("mid")
    end if

    struSQLusemedarb = "SELECT mnavn FROM medarbejdere WHERE mid ="& usemrn
    oRec.open struSQLusemedarb, oConn, 3
    if not oRec.EOF then
    medarbejdernavn = oRec("mnavn")
    end if
    oRec.close

    call menu_2014

    func = request("func")


	if func = "db" then


        aktids = request("FM_aktivitetids")
        strAktNavn = request("FM_aktnavn")

        response.Write "test tekst: " & aktids & strAktNavn
        response.End


    end if


    stDato = request("stDato")

    redim tjekdag(7)
    tjekdag(1) = stDato
            
    for x = 2 to 7
    tjekdag(x) = dateAdd("d", x-1, stDato)
    next



    medid = request("FM_medid")
    response.Write "medid: " & medid

    varTjDatoUS_selectedday = request("varTjDatoUS_selectedday")
    mondayofsameweek = DateAdd("d", -((Weekday(varTjDatoUS_selectedday) + 7 - 2) Mod 7), varTjDatoUS_selectedday)
                  

    varTjDatoUS_man = mondayofsameweek
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

    datoMan = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
    datoSon = day(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& year(varTjDatoUS_son)

    next_varTjDatoUS_man = dateadd("d", 7, varTjDatoUS_man)
	next_varTjDatoUS_son = dateadd("d", 7, varTjDatoUS_son)
    next_varTjDatoUS_man = year(next_varTjDatoUS_man) &"/"& month(next_varTjDatoUS_man) &"/"& day(next_varTjDatoUS_man)
	next_varTjDatoUS_son = year(next_varTjDatoUS_son) &"/"& month(next_varTjDatoUS_son) &"/"& day(next_varTjDatoUS_son)


    prev_varTjDatoUS_man = dateadd("d", -7, varTjDatoUS_man)
	prev_varTjDatoUS_son = dateadd("d", -7, varTjDatoUS_son)
    prev_varTjDatoUS_man = year(prev_varTjDatoUS_man) &"/"& month(prev_varTjDatoUS_man) &"/"& day(prev_varTjDatoUS_man)
	prev_varTjDatoUS_son = year(prev_varTjDatoUS_son) &"/"& month(prev_varTjDatoUS_son) &"/"& day(prev_varTjDatoUS_son)

    weeknumber = year(varTjDatoUS_man) & "-" & month(varTjDatoUS_man) & "-" & day(varTjDatoUS_man)

    select case func 

    case "fjernfavorit"

        medid = request("FM_medid")
        varTjDatoUS_man = request("varTjDatoUS_man")
        id = request("id")

        %>
            <div class="container">
                <div class="portlet">
                    <h3 class="portlet-title"><u>Favorit liste</u></h3>
                    <div class="portlet-body">
                        <%response.Write "aktivi id" & id  %>
                    </div>

                </div>
            </div>
        <%
        

        oConn.execute("UPDATE timereg_usejob SET favorit = 0 WHERE aktid = "&id&"")

        response.Redirect "favorit.asp?FM_medid="&medid&"&varTjDatoUS_man="&varTjDatoUS_man


    case "tilfojfavorit"

        id = request("id")
                            
            %>
                <div class="container">
                <div class="portlet">
                    <h3 class="portlet-title"><u>Favorit liste</u></h3>
                    <div class="portlet-body">
                        <%response.Write "favor" & id  %>
                    </div>

                </div>
                </div>
            <%

        oConn.execute("UPDATE timereg_usejob SET favorit = 1 WHERE aktid = "&id&"")

        response.Redirect "favorit.asp"




    case else

%>



            <div class="container">
                <div class="portlet">
                    <h3 class="portlet-title"><u>Favorit liste</u></h3>
                    <div class="portlet-body">
                                                                                           
                        <form action="favorit.asp?" method="post">
                                              
                        <!-- <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>"> -->

                      <!--  <div class="row">
                            <div class="col-lg-2">
                                 
                                <%
                                    strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe, init FROM medarbejdere WHERE mansat <> 2 GROUP BY mid ORDER BY Mnavn" 
                                %>
                                <select name="FM_medid" id="FM_medid" <%=progrpmedarbDisabled  %> class="form-control input-small"  onchange="submit();" style="width:210px">
                                    <%

                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF
                                
                                    
                                    StrMnavn = oRec("Mnavn")
                                    StrMinit = oRec("init")
	
                                    if cdbl(medid) = cdbl(oRec("Mid")) then
				                    isSelected = "SELECTED"
				                    else
				                    isSelected = ""
				                    end if

				                    %>
                                     <option value="<%=oRec("Mid")%>" <%=isSelected%>><%=StrMnavn &" "& StrMinit%></option>
                                    <%
                                    oRec.movenext
                                    wend
                                    oRec.close  
                                    %>
                                </select>
                            </div>
                           
                           <div class="col-lg-1" style="z-index:1;">
                               <div class='input-group date' style="padding-left:30px; width:50px">
                                    <input type="text" class="form-control input-small" name="varTjDatoUS_selectedday" id="varTjDatoUS_selectedday" value="<%=varTjDatoUS_selectedday %>" />
                                    <span class="input-group-addon input-small">
                                    <span class="fa fa-calendar">
                                    </span>
                                    </span>
                                </div>
                            </div>

                            <div class="col-lg-2">
                                <button type="submit" class="btn btn-sm btn-default"><b>Gå</b></button>
                            </div>

                            <div class="col-lg-3"></div>
                            <h4 class="col-lg-2" style="text-align:right"><a href="favorit.asp?FM_medid=<%=medid %>&varTjDatoUS_man=<%=prev_varTjDatoUS_man %>" ><</a>&nbsp Uge <%=datepart("ww",weeknumber)  %> &nbsp<a href="favorit.asp?FM_medid=<%=medid %>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>" >></a></h4>
                            
                        </div> -->

                            <table>
                                <tr>
                                    <td>
                                        <%
                                        strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe, init FROM medarbejdere WHERE mansat <> 2 GROUP BY mid ORDER BY Mnavn" 
                                        %>
                                        <select name="FM_medid" id="FM_medid" <%=progrpmedarbDisabled  %> class="form-control input-small"  onchange="submit();" style="width:210px">
                                            <%

                                            oRec.open strSQL, oConn, 3
                                            while not oRec.EOF
                                
                                    
                                            StrMnavn = oRec("Mnavn")
                                            StrMinit = oRec("init")
	
                                            if cdbl(medid) = cdbl(oRec("Mid")) then
				                            isSelected = "SELECTED"
				                            else
				                            isSelected = ""
				                            end if

				                            %>
                                             <option value="<%=oRec("Mid")%>" <%=isSelected%>><%=StrMnavn &" "& StrMinit%></option>
                                            <%
                                            oRec.movenext
                                            wend
                                            oRec.close  
                                            %>
                                        </select>
                                    </td>
                                    <td>
                                        <div class="col-lg-2" style="z-index:1;">
                                            <div class='input-group date' style="width:135px">
                                                <input type="text" class="form-control input-small" name="varTjDatoUS_selectedday" id="varTjDatoUS_selectedday" value="<%=varTjDatoUS_selectedday %>" />
                                                <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                                </span>
                                            </div>
                                        </div>                                                                                                                           
                                    </td>
                                    <td>                                       
                                        <button type="submit" class="btn btn-sm btn-default"><b>Gå</b></button>                                       
                                    </td>
                                </tr>
                            </table>

                        </form>
                        <%'response.Write "id: " & medid%>


                        <form action="../timereg/timereg_akt_2006.asp?func=db&rdir=favorit&medid=<%=medid %>&varTjDatoUS_man=<%=varTjDatoUS_man%>" method="post">
                            
                            
                            <input type="hidden" name="varTjDatoUS_man" value="<%=varTjDatoUS_man %>" />
                            <input type="hidden" name="FM_medid" value="<%=medid %>" />
                            <input type="hidden" id="Hidden4" name="FM_dager" value="7"/>
                            <input type="hidden" name="FM_sttid" value="00:00"/>
                            <input type="hidden" name="FM_sltid" value="00:00"/>
                            <input type="hidden" id="" name="FM_vistimereltid" value="0"/>
                            <input type="hidden" id="Hidden5" name="year" value="<%=year(now) %>"/>
                            <input type="hidden" value="0" name="extsysId" />

                        <script src="js/fav_mat.js" type="text/javascript"></script>
                        <table class="table table-striped dataTable table-bordered ui-datatable">
                            
                            

                            <thead>
                                <tr>
                                    <th>Kunde/Job</th>
                                    <th>Aktivitet</th>

                                    <%
                                            perInterval = 6 'dateDiff("d", varTjDatoUS_man, varTjDatoUS_son, 2,2) 
                                            perIntervalLoop = perInterval

                                            for l = 0 to perIntervalLoop 
        
                                            if l = 0 then
                                            varTjDatoUS_use = varTjDatoUS_man
                                            else
                                            varTjDatoUS_use = dateAdd("d", l, varTjDatoUS_man)
                                            end if 

                                            showdate = DatePart("yyyy",varTjDatoUS_use) & "-" & Right("0" & DatePart("m",varTjDatoUS_use), 2) & "-" & Right("0" & DatePart("d",varTjDatoUS_use), 2)

                                            'showweekdayname = weekdayname(weekday(varTjDatoUS_use, 1))
                                            daynamenum = weekday(varTjDatoUS_use,1)

                                            daynameword = WeekDayName(daynamenum,true)

                                            'response.Write weekdayname(weekday(varTjDatoUS_use, 1))

                                            %>
                                                <th style="width:75px"><%=UCase(Left(daynameword,1)) & Mid(daynameword,2) %> <br />
                                                    <%=showdate %>
                                                </th>
                                            <%

                                            next
                                    %>
                                    <th style="text-align:center; width:55px;">Total</th>
                                    <th></th>
                                </tr>
                            </thead>

                            <tbody>
                                <%
                                     
                                     favoriter = 0
                                     lastaktid = 0
                                     i = 10

                                     'Dim jobid, aktid, medarb
                                     Redim jobid(i), aktid(i), medarb(i)

                                     StrSqlfav = "SELECT medarb, jobid, aktid, forvalgt_af FROM timereg_usejob WHERE medarb = "& medid & " AND favorit <> 0" 
                                     
                                     oRec.open StrSqlfav, oConn, 3

                                      i = 0

                                     while not oRec.EOF

                                     

                                     jobid(i) = oRec("jobid")
                                     aktid(i) = oRec("aktid")
                                     medarb(i) = oRec("medarb")
                                     'response.Write jobid(i)

                                    if lastaktid <> oRec("aktid") then

                                        i = i + 1
                                     
                                     end if

                                   

                                    'response.Write jobid(i) & "<br>"
                                    'response.Write aktid(i) & "<br>"
                                    
                                     lastaktid = oRec("aktid")
                                     
                                     favoriter = favoriter + 1

                                     'response.write aktid(i) & "<br>"

                                     oRec.movenext
                                     wend
                                     oRec.close

                                    i_end = i
                                    i = 0
                                    'response.write "<br> d" & i_end
                                    'response.Write "lastid: " & lastjobid 
                                    
                                    y = 0                                
                                     for i = 0 to i_end -1
                                                          
                                        'response.Write "aktid1: " & aktid(i)

                                        StrSQLjob = "SELECT id, jobnavn, jobstartdato, jobslutdato, jobans1, jobans2, jobans3, jobans4, jobans5, jobknr, beskrivelse, budgettimer FROM job WHERE id ="& jobid(i)
                                        oRec3.open StrSQLjob, oConn, 3
                                        if not oRec3.EOF then
                                        jobids = oRec3("id")
                                        jobnavn = oRec3("jobnavn")

                                        
                                        StrSQLakt = "SELECT id, navn, beskrivelse, budgettimer FROM aktiviteter WHERE id ="& aktid(i)
                                        oRec2.open StrSqlakt, oConn, 3
                                        if not oRec2.EOF then
                                        aktNavn = oRec2("navn")
                                        TaktId = oRec2("id")
                                        aktbudgettimer = oRec2("budgettimer")
                                        'response.Write jobnavn
                                        %>
                                        <tr>
                                            <td style="vertical-align:middle; width:200px"><input type="hidden" value="<%=jobids %>" name="FM_jobid" />
                                                <%=jobnavn %><a data-toggle="modal" href="#styledModalSstGrp20"><span id="jobinfo_<%=jobids %>" style="color:#8c8c8c" class="fa fa-file-text pull-right jobinfo"></span></a>                                                                                              
                                                 <div id="jobmodal_<%=jobids %>" class="modal">
                                                     <div class="modal-content" style="width:400px;">
                                                         <%
                                                            strSQLkunde = "SELECT kundeans1, kkundenavn FROM kunder WHERE kid ="& oRec3("jobknr")
                                                            oRec4.open strSQLkunde, oConn, 3
                                                            if not oRec4.EOF then
                                                            kundeans = oRec4("kundeans1")
                                                            kundenavn = oRec4("kkundenavn")
                                                            end if
                                                            oRec4.close

                                                            strSQLjobnas = "SELECT mnavn FROM medarbejdere WHERE mid="& oRec3("jobans1")  
                                                            oRec4.open strSQLjobnas, oConn, 3
                                                            if not oRec4.EOF then
                                                            jobansvarlig = oRec4("mnavn")
                                                            end if
                                                            oRec4.close

                                                            strSQLkundeans = "SELECT mnavn FROM medarbejdere WHERE mid="& kundeans  
                                                            oRec4.open strSQLkundeans, oConn, 3
                                                            if not oRec4.EOF then
                                                            kundeansvarlig = oRec4("mnavn")
                                                            end if
                                                            oRec4.close

                                                            strSQLrealiseret = "SELECT sum(timer) as timer FROM timer WHERE tjobnr ="& jobids
                                                            oRec4.open strSQLrealiseret, oConn, 3
                                                            if not oRec4.EOF then
                                                             realiserettimer = oRec4("timer")
                                                            end if 
                                                            oRec4.close

                                                         %>
                                                       <!--  <div class="row">
                                                            <div class="col-lg-4">
                                                                <table>
                                                                     <tr>
                                                                         <td><b>Start & slut dato</b>
                                                                             <br />
                                                                             <%=oRec3("jobstartdato") & " - " & oRec3("jobslutdato")  %>
                                                                            <br /><br />
                                                                         </td>
                                                                     </tr>
                                                                    <tr>
                                                                         <td><b>Jobansvarlig:</b> <%=jobansvarlig %><br />
                                                                             <b>Kundeansvarlig:</b> <%=kundeansvarlig  %>                                                                             
                                                                         </td>
                                                                       
                                                                     </tr>                                                                                                                                    
                                                                </table>
                                                            </div>
                                                             <div class="col-lg-4">
                                                                <table>
                                                                    <tr>
                                                                        <td><b>Forkalk.</b>: <%=oRec3("budgettimer") %> t.<br />
                                                                            <b>Realiseret:</b> <%=realiserettimer %> t.
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </div>
                                                            <div class="col-lg-4">
                                                                <table>
                                                                    <tr>
                                                                        <td><b>Jobbeskrivelse:</b>
                                                                            <br />
                                                                            <textarea rows="5" class="form-control input-small"><%=oRec3("beskrivelse") %></textarea>                                                                            
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </div>
                                                        </div> -->

                                                        <div class="row">
                                                            <div class="col-lg-12">
                                                                <b><%=kundenavn %></b> - <%=jobnavn %>
                                                            </div>
                                                        </div>
                                                         <br /><br />

                                                        <div class="row">
                                                             <div class="col-lg-4"><b>Start dato:</b></div>
                                                             <div class="col-lg-5"><%=oRec3("jobstartdato")%></div>
                                                        </div>
                                                        <div class="row">
                                                             <div class="col-lg-4"><b>Slut dato:</b></div>
                                                             <div class="col-lg-5"><%=oRec3("jobslutdato")  %></div>
                                                        </div>
                                                        <br />
                                                        <div class="row">
                                                            <div class="col-lg-4"><b>Jobansvarlig:</b></div>
                                                            <div class="col-lg-5"><%=jobansvarlig %></div>
                                                        </div>
                                                        <div class="row">
                                                            <div class="col-lg-4"><b>Kundeansvarlig:</b></div>
                                                            <div class="col-lg-5"><%=kundeansvarlig %></div>
                                                        </div>
                                                        <br />

                                                        <div class="row">
                                                            <div class="col-lg-4"><b>Forkalkuleret:</b></div>
                                                            <div class="col-lg-5"><%=oRec3("budgettimer") %> t.</div>
                                                        </div>
                                                        <div class="row">
                                                            <div class="col-lg-4"><b>Realiseret:</b></div>
                                                            <div class="col-lg-5"><%=realiserettimer %> t.</div>
                                                        </div>
                                                        <br />

                                                        <div class="row">
                                                            <div class="col-lg-4"><b>Jobbeskrivelse:</b></div>
                                                        </div>
                                                        <div class="row">
                                                            <div class="col-lg-12"><textarea rows="5" class="form-control input-small"><%=oRec3("beskrivelse") %></textarea></div>
                                                        </div>                                                    

                                                     </div>
                                                 </div>

                                            </td>

                                            <td style="vertical-align:middle; width:195px;">
                                                <input type="hidden" value="<%=TaktId %>" name="FM_aktivitetid" />
                                                <%=aktNavn %>

                                                <span id="modal_<%=TaktId %>" class="fa fa-chevron-down pull-right picmodal" style="color:#8c8c8c"></span>                                               
                                                <div id="myModal_<%=TaktId %>" style="display:none">
                                                
                                                    <%
                                                        StrSqltimerialt = "SELECT sum(Timer) as timer FROM timer WHERE TAktivitetId ="& TaktId
                                                        oRec6.open StrSqltimerialt, oConn, 3
                                                        if not oRec6.EOF then
                                                        timerforalle = oRec6("timer")
                                                        end if
                                                        oRec6.close

                                                        StrSqltotaltimer = "SELECT TAktivitetId, sum(timer) as timer FROM timer WHERE TAktivitetId ="& TaktId &" AND tmnr ="&medid
                                                        oRec5.open StrSqltotaltimer, oConn, 3
                                                        if not oRec5.EOF then
                                                
                                                        timertotal = oRec5("timer")
                                                        
                                                        if timerforalle > aktbudgettimer then
                                                           txtcolor = "red"
                                                        else
                                                            txtcolor = "green"
                                                        end if

                                                    %>
                                                        <span style="font-size:75%; color:#5582d2">Forkalk.: <%=aktbudgettimer %></span>
                                                        <span style="font-size:75%; color:<%=txtcolor%>;">Real:<%=timerforalle %></span>
                                                        <span style="font-size:75%; color:#5582d2;">Egne: <%=timertotal %></span>
                                                    <%
                                                        end if
                                                        oRec5.close  
                                                    %>
                                                
                                                </div>
                                            </td>
                                            
                                            <%
                                                
                                                for l = 0 to 6
                                                     
                                                    y = y + 1
                                                    'response.Write y

                                                    if l = 0 then
                                                        timerdato = varTjDatoUS_man
                                                    else
                                                        timerdato = dateAdd("d",1,timerdato)
                                                    end if


                                                    timerdato = year(timerdato) & "-" & month(timerdato) & "-" & day(timerdato) 

                                                    StrSQLtimer = "SELECT TAktivitetId, sum(timer) as Timer, extsysId, tdato, Timerkom, origin FROM timer WHERE TAktivitetId ="& TaktId & " AND tdato = "& "'" & timerdato & "' AND tmnr ="& medid
                                                
                                                     oRec4.open StrSQLtimer, oConn, 3
                                                     if not oRec4.EOF then
                                                
                                                     Stryear = oRec4("tdato")
                                                     timerdag = oRec4("Timer")
                                                     extsysid = oRec4("extsysId")
                                                     timerkcoment = oRec4("Timerkom")
                                                     origin = oRec4("origin")
                                                     
                                                     todaydate = DatePart("yyyy",Date) _ 
                                                                &"-"& Right("0" & DatePart("m",Date), 2) _ 
                                                                &"-"& Right("0" & DatePart("d",Date), 2)

                                                    %>
                                                        <td>                                           
                                                            <input type="hidden" name="FM_feltnr" value="<%=y %>" />
                                                           <!-- <input type="hidden" value="<%=oRec4("extsysId")%>" name="extsysId" /> -->
                                                            <input type="hidden" value="<%=timerdato %>" name="FM_datoer" />
                                                            <input type="hidden" value="dist" name="FM_destination_<%=y %>" />
                                                            <%if origin <> 0 then %>
                                                            <input type="hidden" name="FM_timer" value=""/>
                                                            <input type="text" class="form-control input-small" style="width:75px;" value="<%=timerdag %>" readonly />
                                                            <span style="font-size:50%"><a style="color:dimgrey;" href="ugeseddel_2011.asp?usemrn=<%=medid %>&varTjDatoUS_man=<%=varTjDatoUS_man %>">Se ugeseddel</a></span>
                                                            <%else %>
                                                            <div class="row">                                                     
                                                            <div class="col-lg-10" style="padding-right:5px!important"><input type="text" class="form-control input-small" name="FM_timer" value="<%=timerdag %>" /></div>
                                                            <div class="col-lg-1" style="padding-left :0px!important"><span id="modal_<%=y%>" class="kommodal">+</span></div>
                                                            </div>
                                                            
                                                                <div id="kommentarmodal_<%=y%>" class="modal">
                                                                        <div class="modal-content">                                                                                                                        
                                                                            <div class="row">
                                                                                <div class="col-lg-2"><b>Kommentar:</b></div>
                                                                            </div>
                                                                            <div class="row">
                                                                                <div class="col-lg-12"><textarea rows="2" name="FM_kom_<%=y %>" class="form-control input-small"><%=oRec4("Timerkom") %></textarea></div>
                                                                            </div>

                                                                            <br /><br />

                                                                            <div class="panel-group accordion-panel" id="accordion-paneled">
                                                                    
                                                                            <div class="panel panel-default">
                                                                                <div class="panel-heading">
                                                                                    <h6 class="panel-title"><a class="accordion-toggle" data-toggle="collapse" data-target="#collapse_<%=y %>">Udlæg</a></h6>
                                                                                </div>
                                                                                <div id="collapse_<%=y %>" class="panel-collapse collapse">  
                                                                                    
                                                                                   <!-- <input type="hidden" id="matreg_lto" value="<%=lto %>" />
                                                                                    <input type="hidden" id="matreg_func" value="dbopr" />
                                                                                    <input type="hidden" id="matregid" name="matregid" value="0" />
                                                                                    <input type="hidden" id="matreg_jobid" name="jobid" value="0" />
                                                                                    <input type="hidden" id="matreg_aftid" name="aftid" value="0" />
                                                                                    <input type="hidden" id="matreg_medid" value="<%=medid %>" />
                                                                                    <input type="hidden" id="matreg_aktid" name="aktid" value="0" />
                                                                                    <input type="hidden" id="matreg_regdato_0" name="regdato" value="01-01-2002" /> -->


                                                                                    <input type="hidden" id="mat_id" value="0" />
                                                                                    <input type="hidden" id="mat_jobid_<%=y %>" value="<%=jobids %>" />
                                                                                    <input type="hidden" id="mat_dato_<%=y %>" value="<%=todaydate%>" />
                                                                                    <input type="hidden" id="mat_editor" value="<%=medarbejdernavn %>" />                                                                                  
                                                                                    <input type="hidden" id="mat_userid" value="<%=medid %>" />
                                                                                    <input type="hidden" id="mat_forbrugsdato_<%=y %>" value="<%=timerdato %>" />
                                                                                    <input type="hidden" id="mat_serviceaft" value="0" />
                                                                                    <input type="hidden" id="mat_endhed_<%=y %>" value="Stk." />
                                                                                    <input type="hidden" id="mat_aktid_<%=y %>" value="<%=Taktid %>" />
                                                                                    <input type="hidden" id="mat_bilagsnr" value="" />
                                                                                    <input type="hidden" id="mat_varenr" value="0" /> 

                                                                                                                                                          
                                                                                    <div class="panel-body">
                                                                                        
                                                                                        <div class="row" id="error_felt_<%=y %>" style="visibility:hidden">
                                                                                            <span id="error_txt_<%=y %>" class="col-lg-12" style="color:red"></span>
                                                                                        </div>
                                                                                        
                                                                                        <div class="row">
                                                                                            <div class="col-lg-4">Antal:</div>
                                                                                            <div class="col-lg-4"><input type="text" value="1" id="mat_antal_<%=y %>" class="form-control input-small" /></div>
                                                                                        </div>
                                                                                        <div class="row">
                                                                                            <div class="col-lg-4">Mat. navn:</div>
                                                                                            <div class="col-lg-4"><input type="text" value="" id="mat_navn_<%=y %>" class="form-control input-small" /></div>
                                                                                        </div>
                                                                                        <div class="row">
                                                                                            <div class="col-lg-4">Indkøbspris:</div>
                                                                                            <div class="col-lg-3"><input type="text" value="" id="mat_kobpris_<%=y %>" class="form-control input-small" /></div>
                                                                                            <div class="col-lg-4">
                                                                                                <select name="FM_valuta" id="mat_valuta<%=y %>" class="form-control input-small">
		                                                                                            <!--<option value="0"><=tsa_txt_229 %></option>-->
		                                                                                            <%
		                                                                                            strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
    		
    		
		                                                                                            oRec5.open strSQL3, oConn, 3 
		                                                                                            while not oRec5.EOF 
    		
		                                                                                            if cint(valuta) = oRec5("id") then
		                                                                                            valGrpCHK = "SELECTED"
		                                                                                            else
		                                                                                            valGrpCHK = ""
		                                                                                            end if
		    
		   
		                                                                                            %>
		                                                                                            <option value="<%=oRec5("id")%>" <%=valGrpCHK %>><%=oRec5("valutakode")%></option>
		                                                                                            <%
		                                                                                            oRec5.movenext
		                                                                                            wend
		                                                                                            oRec5.close %>
		                                                                                        </select>
                                                                                            </div>                                                      
                                                                                        </div>
                                                                                        <div class="row">
                                                                                            <div class="col-lg-4">Gruppe:</div>
                                                                                            <div class="col-lg-4">                                                                                               
                                                                                                <select class="form-control input-small" name="gruppe" id="mat_gruppe_<%=y %>"><!-- onchange="beregnsalgsprisOTF(0)" -->
		                                                                                            <option value="0"><%=tsa_txt_200 %></option>
		                                                                                            <%
		                                                                                            strSQL = "SELECT id, navn, av FROM materiale_grp ORDER BY navn"
		                                                                                            oRec.open strSQL, oConn, 3 
		
		                                                                                                    while not oRec.EOF 
		
		                                                                                                    if cint(matgrp) = oRec("id") then
		                                                                                                    matgrpSel = "SELECTED"
		                                                                                                    else
		                                                                                                    matgrpSel = ""
		                                                                                                    end if
		
		
    		                                                                                                'matgrpVal = matgrpVal &  "<input id=""avagrpval_"&oRec("id")&""" name=""avagrpval_"&oRec("id")&""" type=""hidden"" value="& oRec("av") &" />"
    		
		
		                                                                                                %>
		                                                                                                <option value="<%=oRec("id")%>" <%=matgrpSel %>><%=oRec("navn")%>
		                                                                                                <%if level <= 2 OR level = 6 then %>
		                                                                                                &nbsp;(<%=oRec("av") %>%)
		                                                                                                <%end if %></option>

		                                                                                            <%
		                                                                                            oRec.movenext
		                                                                                            wend
		                                                                                            oRec.close %>
		                                                                                        </select>
                                                                                            </div>
                                                                                        </div>

                                                                                        <div class="row">
                                                                                            <div class="col-lg-4">Salgspris:</div>
                                                                                            <div class="col-lg-4"><input type="text" id="mat_salgspris_<%=y %>" class="form-control input-small" /></div>
                                                                                        </div>
                                                                                       
                                                                                        
                                                                                        
                                                                                        
                                                                                    <!------------- Henter allerede udlæg ---------------->                                                                  
                                                                                    <%
                                                                                    strSQLudlag = "SELECT matnavn, m.forbrugsdato, m.matsalgspris, v.valutakode FROM materiale_forbrug as m "_ 
                                                                                    & "LEFT JOIN valutaer as v ON (v.id = m.valuta) WHERE aktid ="& Taktid &" AND forbrugsdato ='"& timerdato &"'"
                                                                                    oRec8.open strSQLudlag, oConn, 3
                                                                                    'response.write strSQLudlag
                                                                                    'response.Flush
                                                                                    'if oRec8("forbrugsdato") <> 0 then
                                                                                    %>
                                                                                        <div class="row">
                                                                                            <div class="col-lg-6"><span style="font-size:75%">Indlæste materialer:</span></div>
                                                                                        </div> 
                                                                                    <%
                                                                                    'end if

                                                                                    while not oRec8.EOF
                                                                                    strforbrugsdato = oRec8("Forbrugsdato")
                                                                                    matnavn = oRec8("matnavn")
                                                                                    matsalgspris = oRec8("matsalgspris")
                                                                                    valutakode = oRec8("valutakode")                                                                                                                                                                                                                        
                                                                                        %>                                                                                                                                
                                                                                        <div class="row"><div class="col-lg-12" style="font-size:75%"><%response.Write matnavn & " " & matsalgspris & " " & valutakode %></div></div>
                                                                                        <% 
                                                                                        oRec8.movenext
                                                                                        wend                                                                                 
                                                                                        oRec8.close 
                                                                                        %>                           
                                                                                        <div class="row">
                                                                                            <div class="col-lg-12 pull-right">
                                                                                              <!--  <a id="<%=Taktid %>" class="btn btn-sm btn-default pull-right mat_save"><b>Gem</b></a> -->
                                                                                                <a class="btn btn-sm btn-default pull-right mat_save matreg_sb" id="<%=y %>"><b>Gem</b></a>
                                                                                            </div>
                                                                                        </div>
                                                                              
                                                                                    </div>


                                                                                </div>
                                                                                </div>
                                                                                </div>
                                                                            </div>
                                                                            </div>
                                                                            <%end if  %>
                                                            <input type="hidden" name="FM_timer" value="xx"/>
                                                        </td>
                                                    <%


                                                    else

                                                    timerdag = 0

                                                    'if timerdag = 0 then 
                                                        'extsysid = 0
                                                    'end if

                                                    %>
                                                        <td>
                                                            <input type="hidden" name="FM_feltnr" value="<%=y %>" />
                                                         <!--   <input type="hidden" value="0" name="extsysId" /> -->
                                                            <input type="hidden" value="<%=timerdato %>" name="FM_datoer" />
                                                            <input type="text" style="width:75px;" class="form-control input-small" name="FM_timer" value="<%=timerdag %>" />
                                                            <input type="hidden" name="FM_timer" value="xx"/>
                                                            <input type="hidden" id="FM_kom" name="FM_kom_<%=y %>" placeholder="<%=tsa_txt_051%>" class="form-control input-small"/>
                                                        </td>
                                                    <%

                                                     end if 
                                                     oRec4.close
                                                    

                                                    strforbrugsdato = 0


                                                next

                                                 'StrSQLtimer = "SELECT TAktivitetId, Timer FROM timer WHERE TAktivitetId ="& aktId
                                                
                                                 'oRec4.open StrSQLtimer, oConn, 3
                                                 'if not oRec4.EOF then
                                                
                                                    'timerdag = oRec4("Timer")

                                                 'end if 
                                                 'oRec4.close
                                                 
                                            %>

                                            <td style="text-align:center; vertical-align:middle;">
                                                <%
                                                    ugestart_dato = year(datoMan) & "-" & month(datoMan) & "-" & day(datoMan)
                                                    ugeslut_dato = year(datoSon) & "-" & month(datoSon) & "-" & day(datoSon)

                                                    'response.Write ugestart_dato & "SØNDAG: " & ugeslut_dato
                                                    
                                                    StrSqlweektotal = "SELECT sum(timer) as timer FROM timer WHERE TAktivitetId ="& TaktId & " AND tdato BETWEEN '"& ugestart_dato &"' AND '"& ugeslut_dato &"' AND tmnr ="&medid 
                                                
                                                    oRec7.open StrSqlweektotal, oConn, 3
                                                    if not oRec7.EOF then
                                                    timerweektotal = oRec7("timer")
                                                    
                                                    if timerweektotal <> 0 then
                                                    else
                                                    timerweektotal = 0
                                                    end if

                                                    end if
                                                    oRec7.close    
                                                    
                                                %>
                                                <%=timerweektotal %>
                                            </td>

                                            <td style="vertical-align:middle;">
                                                <a href="favorit.asp?id=<%=oRec2("id") %>&FM_medid=<%=medid %>&varTjDatoUS_man=<%=varTjDatoUS_man%>&func=fjernfavorit"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a>
                                            </td>

                                        </tr>
                                        <%
                                        end if
                                        oRec2.close

                                        end if
                                        oRec3.close

                                    next

                                     
                                    %>


                                  <!--  <tr>
                                        <td colspan="10"><input type="text" id="aktiviteter_sog" class="aktivitet_sog form-control input-small" />
                                            <select id="aktiviteterfelt" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none;">
                                            <option><%=week_txt_007 %>..</option>
                                            </select>
                                        </td>
                                        
                                    </tr> -->

                                    <%
                                        next_akt_id = 0 
                                        for a = 0 to 10
                                        next_akt_id = next_akt_id + 1

                                        'response.Write "next_akt_id :" & next_akt_id
                                    %>
                                    <tr class="next_akt_id" style="visibility:hidden; display:none;" id="FN_akt_tilfojed_<%=next_akt_id %>">
                                        <td><input type="text" value="" id="next_akt_jobid_<%=next_akt_id %>" class="form-control input-small" readonly /></td>
                                        <td><input type="text" value="" id="next_akt_aktid_<%=next_akt_id %>" class="form-control input-small" readonly /></td>


                                        <%for i = 0 to 6 %>
                                        <td><input type="text" style="width:75px;" class="form-control input-small" name="FM_timerdag" value="0" /></td>
                                        <%next %>
                                        <td>&nbsp</td>
                                        <td>&nbsp</td>
                                    </tr>    
                                    <%
                                        next 
                                    %>

                                    <tr style="border-bottom:inherit">  

                                        <input type="hidden" value="0" name="FM_pa" />
                                        <input type="hidden" id="FM_jobid" value=""/>

                                        <td><input type="text" class="FM_job form-control input-small" id="FM_job" value="" placeholder="Søg kunde/job"/>
                                          <!-- <div id="dv_job"></div> -->
                                            <select id="dv_job" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none;">
                                                <option><%=week_txt_007 %>..</option>
                                            </select>
                                        </td>                                           
                             

                                        <td>
                                            <input type="hidden" name="FM_aktivitetid" id="FM_aktid" value=""/>
                                            <input type="text" class="aktivitet_sog form-control input-small" id="FM_akt" value="" placeholder="Søg aktivitet" />
                                            <!--<div id="dv_akt"></div> -->
                                            <select id="dv_akt" class="form-control input-small chbox_akt" size="10" style="visibility:hidden; display:none;">
                                                <option><%=week_txt_007 %>..</option>
                                            </select>                                     
                                        </td>
                                        
                                        <td style="text-align:center;" colspan="9">
                                           <input type="hidden" id="FM_medid_id" class="form-control input-small" value="<%=medid %>" />
                                            <input type="hidden" value="1" id="next_akt_id" />                                                                                      
                                            <a class="tilfoj_akt btn btn-default btn-sm" id="1" style="width:100%"><b>Tilføj</b></a>
                                            <div id="dv_akttil"></div>
                                        </td>                                     
                                    </tr>
                                    

                            </tbody>

                        </table>
                            <div class="row">
                            <div class="col-lg-10">&nbsp</div>
                            <div class="col-lg-2 pad-b10">
                                <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                            </div>
                        </div>
                        </form>
                        <br /><br />
                           
                            <%
                                strSQL = "SELECT id, medarb, aktid, favorit FROM timereg_usejob WHERE medarb ="& medid & " AND aktid <> 0"
                            %>
                               <!-- <select name="FM_favorittilfoj" class="form-control input-small"  onchange="submit();"> -->
                            <%
                              

                                  oRec.open strSQL, oConn, 3

                                  While not oRec.EOF

                                  aktid = oRec("aktid")
                                %>
                                   <!-- <a href="favorit.asp?func=tilfojfavorit&id=<%=aktid %>"><%=aktid %></a> -->
                                <%
                                  oRec.movenext
                                  wend                     
                                  oRec.close
                              %>                               
                           <!-- </select> -->
                        


                        <%
                             select case lto
                                case "tec", "esn", "intranet - local", "outz"
                                aty_sql_realhours = " tfaktim <> 0"
                                case "xdencker", "xintranet - local"
                                aty_sql_realhours = " tfaktim = 1"
                                case else
                                aty_sql_realhours = aty_sql_realhours &""_
		                        & "tfaktim = 30 OR tfaktim = 31 OR tfaktim = 7 OR tfaktim = 11" 'spørg søren om det er rigtigt at  tfaktim skal være noget her
                                end select

                            timerdenneuge_dothis = 0
                            SmiWeekOrMonth = 0
                            'response.Write SmiWeekOrMonth
                            call timerDenneUge(medid, lto, varTjDatoUS_man, aty_sql_realhours, timerdenneuge_dothis, SmiWeekOrMonth)
                            
                            'response.Write manTimer
                            
                            
                            
                            
                            'henter norm timer minus dencker
                            Select case lto
                            case "dencker"
                            
                            case else
                            'response.Write "medid: " & varTjDatoUS_man
                            call normtimerPer(medid, varTjDatoUS_man, 6, 0)

                            end select


                            'response.Write "<br> norm:" & ntimMan
                             

                        %>




                        


                        <input type="hidden" id="timerdagman" value="<%=replace(manTimer, ",", ".") %>" />
                        <input type="hidden" id="timerdagtir" value="<%=replace(tirTimer, ",", ".") %>" />
                        <input type="hidden" id="timerdagons" value="<%=replace(onsTimer, ",", ".") %>" />
                        <input type="hidden" id="timerdagtor" value="<%=replace(torTimer, ",", ".") %>" />
                        <input type="hidden" id="timerdagfre" value="<%=replace(freTimer, ",", ".") %>" />
                        <input type="hidden" id="timerdaglor" value="<%=replace(lorTimer, ",", ".") %>" />
                        <input type="hidden" id="timerdagson" value="<%=replace(sonTimer, ",", ".") %>" />


                        <input type="hidden" id="normdagman" value="<%=replace(ntimMan, ",", ".") %>" />
                        <input type="hidden" id="normdagtir" value="<%=replace(ntimTir, ",", ".") %>" />
                        <input type="hidden" id="normdagons" value="<%=replace(ntimOns, ",", ".") %>" />
                        <input type="hidden" id="normdagtor" value="<%=replace(ntimTor, ",", ".") %>" />
                        <input type="hidden" id="normdagfre" value="<%=replace(ntimFre, ",", ".") %>" />
                        <input type="hidden" id="normdaglor" value="<%=replace(ntimLor, ",", ".") %>" />
                        <input type="hidden" id="normdagson" value="<%=replace(ntimSon, ",", ".") %>" />


                       <!--<div class="row">
                           <div class="col-lg-5"><div id="stacked-vertical-chart" class="chart-holder-200"></div></div>
                        </div>
                        -->

                        <%
                            'maxHeight = "200px"
                            'Width = "9%"
                            
                            dateMan = DateAdd("d",0,varTjDatoUS_man)
                            dateTir = DateAdd("d",1,varTjDatoUS_man)
                            dateOns = DateAdd("d",2,varTjDatoUS_man)
                            dateTor = DateAdd("d",3,varTjDatoUS_man)
                            dateFre = DateAdd("d",4,varTjDatoUS_man)
                            dateLor = DateAdd("d",5,varTjDatoUS_man)
                            dateSon = DateAdd("d",6,varTjDatoUS_man)

                            'timerHeight_man = manTimer * 30
                            'timerHeight_tir = tirTimer * 30
                            'timerHeight_ons = onsTimer * 30
                            'timerHeight_tor = torTimer * 30
                            'timerHeight_fre = freTimer * 30
                            'timerHeight_lor = lorTimer * 30
                            'timerHeight_son = sonTimer * 30


                            balMan = manTimer - ntimMan
                            balTir = tirTimer - ntimTir
                            balOns = onsTimer - ntimOns
                            balTor = torTimer - ntimTor
                            balFre = freTimer - ntimFre
                            balLor = lorTimer - ntimLor
                            balSon = sonTimer - ntimSon

                            weekhourstotal = manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer
                            normtotal = ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon

                            baltotal = weekhourstotal - normtotal

                            if baltotal < 0 then 
                                balcolor = "red;"
                            else
                                balcolor = "green;"
                            end if

                            if balMan < 0 then 
                                mancolor = "red;"
                            else
                                mancolor = "green;"
                            end if

                            if balTir < 0 then 
                                tircolor = "red;"
                            else
                                tircolor = "green;"
                            end if

                            if balOns < 0 then 
                                onscolor = "red;"
                            else
                                onscolor = "green;"
                            end if

                            if balTor < 0 then 
                                torcolor = "red;"
                            else
                                torcolor = "green;"
                            end if

                            if balFre < 0 then 
                                frecolor = "red;"
                            else
                                frecolor = "green;"
                            end if

                            'response.Write balTir 
                            'response.Write timerHeight_man
                             
                        %>


                        <div class="row">
                            <div class="col-lg-12">

                                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                    <thead>
                                        <tr>
                                            <th style="width:405px;"></th>
                                            <!--<th style="width:10%; text-align:center"><%response.Write Month(dateMan) & "-" & Day(dateMan)%></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateTir) & "-" & Day(dateTir) %></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateOns) & "-" & Day(dateOns) %></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateTor) & "-" & Day(dateTor) %></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateFre) & "-" & Day(dateFre) %></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateLor) & "-" & Day(dateLor) %></th>
                                            <th style="width:10%; text-align:center"><%response.Write Month(dateSon) & "-" & Day(dateSon) %></th>-->
                                            <th style="width:75px; text-align:center">Ma</th>
                                            <th style="width:75px; text-align:center">Ti</th>
                                            <th style="width:75px; text-align:center">On</th>
                                            <th style="width:75px; text-align:center">To</th>
                                            <th style="width:75px; text-align:center">Fr</th>
                                            <th style="width:75px; text-align:center">Lø</th>
                                            <th style="width:75px; text-align:center">Sø</th>
                                            <th style="width:75px; text-align:center">Total</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr>
                                            <td>Timer:</td>
                                            <td style="text-align:center"><%=replace(manTimer, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(tirTimer, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(onsTimer, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(torTimer, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(freTimer, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(lorTimer, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(sonTimer, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(weekhourstotal, ",", ".") %></td>
                                        </tr>
                                    
                                        <tr>
                                            <td>Norm:</td>
                                            <td style="text-align:center"><%=replace(ntimMan, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(ntimTir, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(ntimOns, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(ntimTor, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(ntimFre, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(ntimLor, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(ntimSon, ",", ".") %></td>
                                            <td style="text-align:center"><%=replace(normtotal, ",", ".") %></td>
                                        </tr>

                                        <tr>
                                            <td>Balance:</td>
                                            <td style="text-align:center; color:<%=mancolor%>"><%=balMan %></td>
                                            <td style="text-align:center; color:<%=tircolor%>"><%=balTir %></td>
                                            <td style="text-align:center; color:<%=onscolor%>"><%=balOns %></td>
                                            <td style="text-align:center; color:<%=torcolor%>"><%=balTor %></td>
                                            <td style="text-align:center; color:<%=frecolor%>"><%=balFre %></td>
                                            <td style="text-align:center;"><%=balLor %></td>
                                            <td style="text-align:center;"><%=balSon %></td>
                                            <td style="text-align:center; color:<%=balcolor%>"><%=replace(baltotal, ",", ".") %></td>
                                        </tr>
                                    </tbody>
                                </table>

                            </div>
                        </div>

                        


                       <!-- <div class="row">
                            <div class="col-lg-8">
                                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                                    <tr>
                                        <th colspan="8"><h3>Ugetotal</h3></th>
                                    </tr>

                                    <tr>
                                        <td></td>
                                        <td>Ma<br /><%=varTjDatoUS_man %></td>
                                        <td>Ti<br /><%=dateTir %></td>
                                        <td>On<br /><%=dateOns %></td>
                                        <td>To<br /><%=dateTor %></td>
                                        <td>Fr<br /><%=dateFre %></td>
                                        <td>Lø<br /><%=dateLor %></td>
                                        <td>Sø<br /><%=dateSon %></td>
                                    </tr>

                                    <tr>
                                        <td><div style="height:200px;"></div></td>
                                        <td style="vertical-align:bottom; width:<%=Width%>;"><div style="width:100%; height:<%=timerHeight_man%>px; background-color:forestgreen; max-height:222px;"></div></td>
                                        <td style="vertical-align:bottom; width:<%=Width%>;"><div style="width:100%; height:<%=timerHeight_tir%>px; background-color:forestgreen; max-height:200px;"></div></td>
                                        <td style="vertical-align:bottom; width:<%=Width%>;"><div style="width:100%; height:<%=timerHeight_ons%>px; background-color:forestgreen; max-height:200px;"></div></td>
                                        <td style="vertical-align:bottom; width:<%=Width%>;"><div style="width:100%; height:<%=timerHeight_tor%>px; background-color:forestgreen; max-height:200px;"></div></td>
                                        <td style="vertical-align:bottom; width:<%=Width%>;"><div style="width:100%; height:<%=timerHeight_fre%>px; background-color:forestgreen; max-height:200px;"></div></td>
                                        <td style="vertical-align:bottom; width:<%=Width%>;"><div style="width:100%; height:<%=timerHeight_lor%>px; background-color:forestgreen; max-height:200px;"></div></td>
                                        <td style="vertical-align:bottom; width:<%=Width%>;"><div style="width:100%; height:<%=timerHeight_son%>px; background-color:forestgreen; max-height:200px;"></div></td>
                                    </tr>

                                    <tr style="text-align:center">
                                        <td style="text-align:left; width:25%">Opgjorte Timer:</td>
                                        <td><%=replace(manTimer, ",", ".") %></td>
                                        <td><%=replace(TirTimer, ",", ".") %></td>
                                        <td><%=replace(OnsTimer, ",", ".") %></td>
                                        <td><%=replace(TorTimer, ",", ".") %></td>
                                        <td><%=replace(FreTimer, ",", ".") %></td>
                                        <td><%=replace(LorTimer, ",", ".") %></td>
                                        <td><%=replace(SonTimer, ",", ".") %></td>
                                    </tr>

                                    <tr style="text-align:center">
                                        <td style="text-align:left">Norm:</td>
                                        <td><%=replace(ntimTir, ",", ".") %></td>
                                        <td><%=replace(ntimOns, ",", ".") %></td>
                                        <td><%=replace(ntimTir, ",", ".") %></td>
                                        <td><%=replace(ntimTor, ",", ".") %></td>
                                        <td><%=replace(ntimFre, ",", ".") %></td>
                                        <td><%=replace(ntimLor, ",", ".") %></td>
                                        <td><%=replace(ntimSon, ",", ".") %></td>
                                    </tr>
                                    
                                </table>
                            </div>
                        </div> -->


                        <%                         
                            'd = "25-05-2017" 
                            'mondayofsameweek = DateAdd("d", -((Weekday(d) + 7 - 2) Mod 7), d)
                            'response.Write "mandag:" & mondayofsameweek 
                        %>


                    </div>
                </div>
            </div>

            



    <%end select %>



        </div>
    </div>


<script type="text/javascript">


$(".picmodal").click(function() {

    var modalid = this.id
    var idlngt = modalid.length
    var idtrim = modalid.slice(6, idlngt)

    //var modalidtxt = $("#myModal_" + idtrim);
    var modal = document.getElementById('myModal_' + idtrim);

    
    if (modal.style.display !== 'none')
    {
        modal.style.display = 'none';
    }
    else
    {
        modal.style.display = 'block';
    }
   

});


$(".kommodal").click(function () {

    //alert("klik")

    var modalid = this.id
    var idlngt = modalid.length
    var idtrim = modalid.slice(6, idlngt)

    //var modalidtxt = $("#myModal_" + idtrim);
    var modal = document.getElementById('kommentarmodal_' + idtrim);

    modal.style.display = "block";

    window.onclick = function (event) {
        if (event.target == modal) {
            modal.style.display = "none";
        }
    }

});



$(".jobinfo").click(function () {
    
    var modalid = this.id
    var idlngt = modalid.length
    var idtrim = modalid.slice(8, idlngt)
    
    //var modalidtxt = $("#myModal_" + idtrim);
    var jobmodal = document.getElementById('jobmodal_' + idtrim);

    jobmodal.style.display = "block";

    window.onclick = function (event) {
        if (event.target == jobmodal) {
            jobmodal.style.display = "none";
        }
    }


});


</script>


<!--#include file="../inc/regular/footer_inc.asp"-->