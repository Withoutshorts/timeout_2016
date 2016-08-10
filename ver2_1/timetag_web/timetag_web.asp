
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<%'GIT 20160810 - SK 3
response.buffer = true

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

                'if lto = "essens" then
                'medid = 3
                'end if



                if filterVal <> 0 then
            
                 lastKid = 0
                
                strJobogKunderTxt = strJobogKunderTxt &"<span style=""color:red; font-size:9px; float:right;"" id=""luk_jobsog"">[X]</span>"    
                         
                strJobogKunderTxt = strJobogKunderTxt &"<ul>"

                strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM timereg_usejob AS tu "_ 
                &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                &" WHERE tu.medarb = "& medid &" AND (j.jobstatus = 1 OR j.jobstatus = 3)" 
        

                'Personlig aktiv jobliste
                'select case lto
                'case "hestia" 'Tvunget
                'strSQL = strSQL & ""
                'case else
                'strSQL = strSQL & " AND tu.forvalgt = 1" 

                        call positiv_aktivering_akt_fn()
                        if cint(pa_aktlist) = 0 then 'PA = 0 kan søge i jobbanken / PA = 1 kan kun søge på aktivjobliste
                        strSQL = strSQL & "" 'standard
                        else
                        strSQL = strSQL & " AND tu.forvalgt = 1" 
                        end if

                'end select
    
                strSQL = strSQL & " AND "_
                &" (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '"& jobkundesog &"%' OR "_
                &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"')  AND kkundenavn <> ''"_
                &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT 50"       
    
                'if lto = "essens" then
                'response.write strSQL
                'response.end 
                'end if
                           

                c = 0
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
                if lastKid <> oRec("kid") then
                
                    if c <> 0 then
                    strJobogKunderTxt = strJobogKunderTxt &"<br>"
                    end if            

                strJobogKunderTxt = strJobogKunderTxt &"<b><u>"& oRec("kkundenavn") &" "& oRec("kkundenr") &"</b></u><br>"
                end if 
                 
                strJobogKunderTxt = strJobogKunderTxt & "<li class=""span_job"" id=""chbox_job_"& oRec("jid") &"""><input type=""hidden"" id=""hiddn_job_"& oRec("jid") &""" value="""& oRec("jobnavn") & " ("& oRec("jobnr") &")"">"
                strJobogKunderTxt = strJobogKunderTxt & "<input type=""hidden"" id=""hiddn_jid_"& oRec("jid") &""" value="& oRec("jid") &"> "& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"</li>" 
                
                lastKid = oRec("kid") 
                c = c + 1
                oRec.movenext
                wend
                oRec.close

               strJobogKunderTxt = strJobogKunderTxt &"</ul>"


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
                pa = request("jq_pa") 


                call akttyper2009(2)
                
                
                'response.write "aty_sql_realhours: "& aty_sql_realhours & " aty_sql_hide_on_treg: " & replace(aty_sql_hide_on_treg, )           
                'response.end


                if filterVal <> 0 then
            
                 
    
                strAktTxt = strAktTxt &"<span style=""color:red; font-size:9px; float:right;"" class=""luk_aktsog"">[X]</span>"    

                strAktTxt = strAktTxt &"<ul>"
                         
               'pa = 0 '***ÆNDRES til variable
               if pa = 1 then
               strSQL= "SELECT a.id AS aid, navn AS aktnavn FROM timereg_usejob LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
               &" WHERE tu.medarb = "& medid &" AND tu.jobid = "& jobid &" AND aktid <> 0 AND a.navn LIKE '%"& aktsog &"%' AND aktstatus = 1 AND ("& replace(aty_sql_realhours, "tfaktim", "a.fakturerbar") &") AND ("& aty_sql_hide_on_treg &") ORDER BY navn LIMIT 20"   


               else


                        '*** Finder medarbejders projektgrupper 
                        '** Medarbejder projektgrupper **'
                        medarbPGrp = "#0#" 
                        strMpg = "SELECT projektgruppeId, medarbejderId, teamleder FROM progrupperelationer WHERE medarbejderId = "& medid & " GROUP BY projektgruppeId"

                     
                        'response.write "strMpg " & strMpg
                        'response.end

                        oRec5.open strMpg, oConn, 3
                        while not oRec5.EOF
                        medarbPGrp = medarbPGrp & ",#"& oRec5("projektgruppeId") &"#"         
        
                        oRec5.movenext
                        wend
                        oRec5.close 


               


               strSQL= "SELECT a.id AS aid, navn AS aktnavn, projektgruppe1, projektgruppe2, projektgruppe3, "_
               &" projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter AS a "_
               &" WHERE a.job = " & jobid & " AND a.job <> 0 AND navn LIKE '%"& aktsog &"%' AND aktstatus = 1 AND ("& replace(aty_sql_realhours, "tfaktim", "a.fakturerbar") &") AND ("& aty_sql_hide_on_treg &") ORDER BY navn LIMIT 20"      
    

              
                end if


                'response.write "strSQL " & strSQL
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
               
                  

                'showAkt = 1

                
                
                if cint(showAkt) = 1 then 
                 
                strAktTxt = strAktTxt & "<li class=""span_akt"" id=""span_akt_"& oRec("aid") &"""><input type=""hidden"" id=""hiddn_akt_"& oRec("aid") &""" value="""& oRec("aktnavn") &""">" 
                strAktTxt = strAktTxt & ""& oRec("aktnavn") &"</li>" 
                
                end if
                
                oRec.movenext
                wend
                oRec.close

              
                strAktTxt = strAktTxt &"</ul>"        


                    '*** ÆØÅ **'
                    call jq_format(strAktTxt)
                    strAktTxt = jq_formatTxt


                    response.write strAktTxt

                end if    

          


           case "FN_sogmat"


                

                '*** SØG Mat            
                
                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                jobkundesog = request("jq_newfilterval")
                else
                filterVal = 0
                jobkundesog = "6xxxxxfsdf554"
                end if
        
                

                if filterVal <> 0 then
            
                
                
                strMat = strMat &"<span style=""color:#999999; font-size:9px; float:right;"" id=""luk_matsog"">X</span>"    
                         
                strMat = strMat &"<ul>"

                strSQL = "SELECT m.navn AS matnavn, m.id AS matid, m.matgrp, m.varenr, m.enhed, mg.navn AS grpnavn, mg.id AS mgid FROM materialer AS m "_
                &" LEFT JOIN materiale_grp AS mg ON (mg.id = m.id) WHERE m.id <> 0 AND (m.navn LIKE '%"& jobkundesog &"%') ORDER BY m.navn LIMIT 100"   
    

                'response.write strSQL
                'response.end            

                c = 0
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
                if lastgrp <> oRec("grpnavn") then
                
                    if c <> 0 then
                    strMat = strMat &"<br>"
                    end if            

                strMat = strMat &"<b><u>"& oRec("grpnavn") &"</b></u><br>"
                end if 
                 
                strMat = strMat & "<li class=""span_mat"" id=""chbox_mat_"& oRec("matid") &"""><input type=""hidden"" id=""hiddn_mat_"& oRec("matid") &""" value="""& oRec("matnavn") &""">"
                strMat = strMat & "<input type=""hidden"" id=""hiddn_matid_"& oRec("matid") &""" value="& oRec("matid") &"> "& oRec("matnavn") &" ("& oRec("enhed") &") </li>" 
                
                lastgrp = oRec("mgid") 
                c = c + 1
                oRec.movenext
                wend
                oRec.close

               strMat = strMat &"</ul>"


                    '*** ÆØÅ **'
                    call jq_format(strMat)
                    strMat = jq_formatTxt


                    response.write strMat

                end if    

        
      
          case "FN_tomobjid"


                if len(trim(request("jq_jobid"))) <> 0 then
                jq_jobid = request("jq_jobid")
                else
                jq_jobid = 0
                end if

                
                strJobnavn = ""
                strSQL = "SELECT jobnavn, jobnr FROM job WHERE id = "& jq_jobid
                

                oRec.open strSQL, oConn, 3
                if not oRec.EOF then
        
                strJobnavn = oRec("jobnavn") & " ("& oRec("jobnr") &")"           

                end if
                oRec.close
                
            
                  '*** ÆØÅ **'
                    call jq_format(strJobnavn)
                    strJobnavn = jq_formatTxt


                    response.write strJobnavn


        end select
        response.end
        end if

    %>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">




    <% 

    thisfile = "timetag_mobile"

    if len(trim(request("flushsessionuser"))) <> 0 then ' flushsessionuser = kommer fra logindsiden
     flushsessionuser = 1
    else
     flushsessionuser = 0
    end if

    '*** Hvis der ligger coookie, skal telefonen blive på / flushsessionuser = kommer fra logind ****'
    if request.Cookies("timeouttimeoutcloud")("mobileuser") <> "" AND cint(flushsessionuser) <> 1 then

    session("mid") = request.Cookies("timeoutcloud")("mobilemid")
    session("user") = request.Cookies("timeoutcloud")("mobileuser")
    session("rettigheder") = request.Cookies("timeoutcloud")("mobilerettigheder")

    end if




    if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/errors/error_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else

    call meStamdata(session("mid"))


    if len(trim(request("indlast"))) <> 0 AND session("timetag_web_indMsgShown") <> "1" then
    indlast = 1
    else
    indlast = 0
    end if


    tomobjid = session("tomobjid")

    select case lto
    case "hestia", "xoutz", "micmatic"
        showAfslutJob = 1
        showMatreg = 1
        showStop = 0
    case "intranet - local", "sdutek", "nonstop", "xoutz", "cc"
        showAfslutJob = 0
        showMatreg = 0
        showStop = 1
    case else
        showAfslutJob = 0
        showMatreg = 0
        showStop = 0
    end select
    %>


<head>
<!--<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />-->
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TimeOut mobile</title>
<link href='http://fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css'>

<!--<link href="../inc/menu/css/chronograph.css" rel="stylesheet" type="text/css" />-->
<link href="css/style.css" rel="stylesheet" type="text/css" />


    
  <link href="../inc/jquery/jquery-ui-1.7.1.custom.css" rel="stylesheet" type="text/css" />
	<script src="../inc/jquery/jquery-1.3.2.min.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery-ui-1.7.1.custom.min.js" type="text/javascript"></script>

	<script src="../inc/jquery/timeout.jquery.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery.coookie.js" type="text/javascript"></script>
	<script src="../inc/jquery/jquery.scrollTo-1.4.2-min.js" type="text/javascript"></script>
    <script src="../inc/jquery/jquery.timer.js" type="text/javascript"></script>
    <script src="../inc/jquery/jquery.corner.js" type="text/javascript"></script>


<script type="text/javascript">
  less = {
    env: "development",
    async: false,
    fileAsync: false,
    poll: 1000,
    functions: {},
    dumpLineNumbers: "comments",
    relativeUrls: false,
    rootpath: ""
  };
</script>



<script src="js/timetag_web_jav.js" type="text/javascript"></script>
<script src="js/less.js" type="text/javascript"></script>
</head>
<body>
    <div id="header">TimeOut mobile</div>

      <div id="dvindlaes_msg" style="position:absolute; top:0px; left:0px; height:100%; width:100%; background-color:#cccccc; visibility:hidden; display:none;">Indlæser timer...vent</div>

    <%call positiv_aktivering_akt_fn() %>

    <form id="container" action="../timereg/timereg_akt_2006.asp?func=db&rdir=timetag_web" method="post">

          <%if cint(indlast) = 1 then %>
            <div id="timer_indlast" class="approved">Timer indlæst</div>
         <%
        
        session("timetag_web_indMsgShown") = "1"
        end if%>

        <input type="hidden" id="Hidden5" name="year" value="<%=year(now) %>"/>
        <input type="hidden" id="Hidden1" name="FM_datoer" value="<%=day(now) &"/"& month(now) &"/"& year(now) %>"/>
        <!--<input type="hidden" id="Hidden3" name="FM_dager" value=""/>-->
        <input type="hidden" id="Hidden4" name="FM_dager" value="0"/><!-- , xx -->
        <input type="hidden" id="Hidden2" name="FM_feltnr" value="0"/>
        <input type="hidden" id="FM_pa" name="FM_pa" value="<%=positiv_aktivering_akt_val %>"/>
        <input type="hidden" id="FM_medid" name="FM_medid" value="<%=session("mid") %>"/>
        <input type="hidden" id="FM_medid_k" name="FM_medid_k" value="<%=session("mid") %>"/>
        <input type="text" id="FM_job" name="FM_job" value="Kunde/job"/>
        <input type="hidden" id="tomobjid" name="tomobjid" value="<%=tomobjid %>"/>
        <input type="hidden" id="showstop" name="showstop" value="<%=showStop %>"/>
        <input type="hidden" id="" name="FM_vistimereltid" value="<%=showStop %>"/>
        

        <input type="hidden" id="FM_jobid" name="FM_jobid" value="0"/>
        <div id="dv_job" class="dv-closed"></div> <!-- dv_job -->
        <input type="text" id="FM_akt" name="activity" value="Aktivitet"/>
        <input type="hidden" id="FM_aktid" name="FM_aktivitetid" value="0"/>
        <div id="dv_akt" class="dv-closed"></div> <!-- dv_akt -->


        <%if cint(showStop) = 1 then%>

        <div style="white-space:nowrap; width:100%; border:0;">
            <table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td>
           <input type="button" id="bt_stst" value="St. / Stop" style="background-color:#428bca; color:#FFFFFF; padding:2px 5px 2px 5px;" /> </td><td>
           <input type="text" id="FM_sttid" name="FM_sttid" value="00:00" style="color:#cccccc; width:65px;"/></td><td>
          <input type="text" id="FM_sltid" name="FM_sltid" value="00:00" style="color:#cccccc; width:65px;"/></td><td style="padding:2px 5px 12px 5px; width:85px;" >
             = <span id="FM_timerlbl" style="color:#999999; font-size:18px; width:65px;">0</span>
              <input type="hidden" id="FM_timer" name="FM_timer" value="0" style="color:#cccccc; width:65px;"/></td></tr></table>
        </div>      
            
        <%else %>
          <input type="hidden" id="FM_sttid" name="FM_sttid" value="00:00"/>
          <input type="hidden" id="FM_sltid" name="FM_sltid" value="00:00"/>
         <input type="text" id="FM_timer" name="FM_timer" value="Antal timer"/><!-- brug type number for numerisk tastatur -->
        <% end if%>
        
       
        <input type="text" id="FM_kom" name="FM_kom_0" value="Kommentar"/>
      
        

       <%if cint(showMatreg) = 1 then%>
        
        <a href="#" id="a_mat">+ Tilføj materialeforbrug</a> 
       
        <div id="dv_mat_container" class="dvm-closed">
           
            <table width="100%"><tr><td style="width:60%;">
         <input type="text" id="FM_matnavn" name="FM_matnavn" value="Tilføj Materiale" style="width:100%;"/></td> 
               <td style="width:20%;"><input type="text" id="FM_matantal" name="FM_matantal" value="Ant." style="width:100%;"/>
                
            </td><td style="width:20%;"><input type="button" value=">>" id="sbmmat" style="width:100%;" /></td></tr>
            </table>

                <input type="hidden" id="FM_matid" name="FM_matid" value="0"/>
                <div id="dv_mat" class="dv-closed"></div> <!-- dv_mat -->
             <input type="hidden" id="FM_matids" name="FM_matids" value=""/>
            <input type="hidden" id="FM_matantals" name="FM_matantals" value=""/>

             
        <div id="dv_mat_sbm" style="position:relative; border:0px; padding-left:10px; height:40px; overflow:auto;"></div>     

        </div>
        
        <%end if %>
        <!--
        <a href="#">+ Udlæg</a> 

         <input type="text" id="Text1" name="FM_matnavn" value="Tilføj materiale"/>
        <input type="hidden" id="Hidden3" name="FM_matid" value="0"/>
        -->
        
        <br /><br />
        <%if cint(showAfslutJob) = 1 then %>
        <span><input type="checkbox" value="2" name="FM_lukjobstatus" id="FM_lukjobstatus" /> Job er afsluttet</span> 
        <%end if %>
        <input type="submit" id="sbm_timer" class="active" value="Gem registrering >>"/>


      
        <br />
        <%
            call akttyper2009(2)

             
             ddDato = year(now) &"/"& month(now) &"/"& day(now) 
            timerIdagTxt = ""
            timerIdag = 0


            if cint(showStop) = 1 then

             timerIdagTxt = "<table cellpadding=0 cellspacing=0 border=0 width=100% >"

             timerIdagTxt = timerIdagTxt & "<tr><td colspan=4><b>Timer i dag:</b></td></tr>"
           
            
             strSQLtimer = "SELECT timer, tjobnavn, taktivitetnavn, sttid, sltid, timer FROM timer WHERE tmnr = "& session("mid") &" AND tdato = '"& ddDato &"' AND ("& aty_sql_realhours &")"
            
            oRec.open strSQLtimer, oConn, 3
            while not oRec.EOF 

            timerIdagTxt = timerIdagTxt & "<tr><td>"& left(formatdatetime(oRec("sttid"), 3), 5) & " - "& left(formatdatetime(oRec("sltid"), 3), 5) &"</td><td>"& left(oRec("tjobnavn"), 15) &"</td><td>"& left(oRec("taktivitetnavn"), 10) &"</td><td align=right>"&  formatnumber(oRec("timer"), 2) & "</td></tr>"


            timerIdag = timerIdag + oRec("timer")
            oRec.movenext
            wend
            oRec.close 

           


             if len(trim(timerIdag)) <> 0 then
            timerIdag = timerIdag
            else
            timerIdag = 0 
            end if

            timerIdag = formatnumber(timerIdag, 2)

            timerIdagTxt = timerIdagTxt & "<tr><td colspan=4 align=right><b><u>"& timerIdag &"</u></b></td></tr>"

            timerIdagTxt = timerIdagTxt & "</table>"

            %> 
            <div id="dv_timeridag" style="font-size:12px; width:100%; text-align:left; border:0px; color:#999999; white-space:nowrap;">
                <%=timerIdagTxt %>
            </div>
            <%

            else

            
            strSQLtimer = "SELECT SUM(timer) AS timer FROM timer WHERE tmnr = "& session("mid") &" AND tdato = '"& ddDato &"' AND ("& aty_sql_realhours &")"
            
            oRec.open strSQLtimer, oConn, 3
            if not oRec.EOF then

            timerIdag = oRec("timer")

            end if
            oRec.close 

            if len(trim(timerIdag)) <> 0 then
            timerIdag = timerIdag
            else
            timerIdag = 0 
            end if

            timerIdag = formatnumber(timerIdag, 2)

            %>
            <div id="dv_timeridag" style="font-size:32px; width:100%; text-align:right; border:0px; color:#999999;"><span style="font-size:12px; line-height:13px; vertical-align:baseline; color:#999999;"><%=meNavn %><br />Timer i dag:<br /></span> <a href="../to_2015/ugeseddel_2011.asp"><%=timerIdag%></a> </div>
            <%

            end if

           
            
            %>


       
        <!--
        <span style="font-size:10px; color:#999999;">[<%=lto %>:<%=meNavn %>]</span>
        -->

        <!--
        <input type="submit" class="inactive" value="Tilføj timer & materialer"/>
        -->
    </form>

    <!--#include file="../inc/regular/footer_inc.asp"-->
    <%end if 'session %>