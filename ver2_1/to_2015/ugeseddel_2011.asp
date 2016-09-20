<%response.buffer = true%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../timereg/inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/treg_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<!---------------------  Github 3 -------------------------------->


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

                call positiv_aktivering_akt_fn()
                if cint(pa_aktlist) = 0 then 'PA = 0 kan søge i jobbanken / PA = 1 kan kun søge på aktivjobliste
                strSQLPAkri = strSQLPAkri & ""
                else
                strSQLPAkri = strSQLPAkri & " AND tu.forvalgt = 1" 
                end if
            

                varTjDatoUS_man = request("varTjDatoUS_man")
                varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

                varTjDatoUS_man = year(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& day(varTjDatoUS_man)
                varTjDatoUS_son = year(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& day(varTjDatoUS_son)

                '*** Datospærring Vis først job når stdato er oprindet
                call lukaktvdato_fn()
                ignJobogAktper = lukaktvdato

                select case ignJobogAktper
                case 0,1
                strSQLDatokri = " AND ((j.jobstartdato <= '"& varTjDatoUS_son &"' AND j.jobstatus = 1) OR (j.jobstatus = 3))"
                case 3
                strSQLDatokri = " AND ((j.jobstartdato <= '"& varTjDatoUS_son &"' AND j.jobslutdato >= '"& varTjDatoUS_man &"' AND j.jobstatus = 1) OR (j.jobstatus = 3))"
                case else
                strSQLDatokri = ""
                end select

                'strSQLDatokri = " AND ((j.jobstartdato <= '"& varTjDatoUS_son &"' AND j.jobslutdato >= '"& varTjDatoUS_man &"' AND j.jobstatus = 1) OR (j.jobstatus = 3))"
                

                if filterVal <> 0 then
            
                 lastKid = 0
                
                'strJobogKunderTxt = strJobogKunderTxt &"<span style=""color:red; font-size:9px; float:right;"" class=""luk_jobsog"">[X]</span>"    
                         
               

                strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM timereg_usejob AS tu "_ 
                &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                &" WHERE tu.medarb = "& medid &" AND (j.jobstatus = 1 OR j.jobstatus = 3) "& strSQLPAkri &" "& strSQLDatokri &" AND "_
                &" (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '"& jobkundesog &"%' OR "_
                &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"')  AND kkundenavn <> ''"_
                &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT 50"       
    

                 'response.write "<option>strSQL " & strSQL & "</option>"
                 'response.end
                
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        



                if lastKid <> oRec("kid") then
    

                    if lastKid <> 0 then
                    ' strJobogKunderTxt = strJobogKunderTxt &"<br>"
                    strJobogKunderTxt = strJobogKunderTxt & "<option DISABLED></option>"
                    end if
    
                ' strJobogKunderTxt = strJobogKunderTxt & oRec("kkundenavn") &" "& oRec("kkundenr") &"<br>"

                 strJobogKunderTxt = strJobogKunderTxt & "<option DISABLED>"& oRec("kkundenavn") &" "& oRec("kkundenr") &"</option>"
    
                end if 
                 
               ' strJobogKunderTxt = strJobogKunderTxt & "<input type=""hidden"" id=""hiddn_job_"& oRec("jid") &""" value="""& oRec("jobnavn") & " ("& oRec("jobnr") &")"">"
               ' strJobogKunderTxt = strJobogKunderTxt & "<a class=""chbox_job"" id=""chbox_job_"& oRec("jid") &""" value="& oRec("jid") &">"& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"</a><br>" 
                
                strJobogKunderTxt = strJobogKunderTxt & "<option value="& oRec("jid") &">"& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"</option>"

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
                if len(trim(request("jq_pa") )) <> 0 then
                pa = request("jq_pa") 
                else
                pa = 0
                end if
        
                'pa = 0
            
                varTjDatoUS_man = request("varTjDatoUS_man")
                varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

                varTjDatoUS_man = year(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& day(varTjDatoUS_man)
                varTjDatoUS_son = year(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& day(varTjDatoUS_son)

                
                '*** Datospærring Vis først job når stdato er oprindet
                call lukaktvdato_fn()
                ignJobogAktper = lukaktvdato

                '** HÅRD styrring af aktiviteer slået til
	            '** Aktivitetsstatus vises først når der hentes aktiviteter, men gør SQL kald hurtigere. 
	            if (cint(ignJobogAktper) = 1 OR cint(ignJobogAktper) = 2 OR cint(ignJobogAktper) = 3) then
	            strSQLDatoKri = " AND ((a.aktstartdato <= '"& varTjDatoUS_son &"' AND a.aktslutdato >= '"& varTjDatoUS_man &"') OR (a.fakturerbar = 6))" 
	            end if


                if filterVal <> 0 then
            
                 
    
                 'strAktTxt = strAktTxt &"<span style=""color:#999999; font-size:9px; float:right;"" class=""luk_aktsog"">X</span>"    
                         

                   if pa = "1" then

                   strSQL= "SELECT a.id AS aid, navn AS aktnavn FROM timereg_usejob AS tu LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
                   &" WHERE tu.medarb = "& medid &" AND tu.jobid = "& jobid &" AND aktid <> 0 AND a.navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"   


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
                   &" WHERE a.job = " & jobid & " "& strSQLDatoKri &" AND navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"      
    

               
            
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
                'strAktTxt = strAktTxt & "<a class=""chbox_akt"" id=""chbox_akt_"& oRec("aid") &""" value="& oRec("aid") &">"& oRec("aktnavn") &"</a><br>" 

                
                strAktTxt = strAktTxt & "<option value="& oRec("aid") &">"& oRec("aktnavn") &"</option>" 
                
                end if
                
                oRec.movenext
                wend
                oRec.close

              


                    '*** ÆØÅ **'
                    call jq_format(strAktTxt)
                    strAktTxt = jq_formatTxt


                    response.write strAktTxt

                end if    




        end select
        response.end
        end if




tloadA = now
 

if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	thisfile = "ugeseddel_2011.asp"
	media = request("media")




    if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> 0 then
    nomenu = 1
    else
    nomenu = 0
    end if

    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = session("mid")
    end if
    


            
         

    level = session("rettigheder")
 
    rdir = request("rdir")
    select case rdir
    case "logindhist"
    rdirfile = "logindhist_2011.asp"
    case "godkenduge"
    rdirfile = "godkenduge.asp"
    case else
	rdirfile = "ugeseddel_2011.asp"
	end select
    

    if len(trim(request("medarbsel_form"))) <> 0 then

        'response.write "HER"
        'response.end

    if len(trim(request("FM_visAlleMedarb"))) <> 0 then
    visAlleMedarbCHK = "CHECKED"
    visAlleMedarb = 1
    else
    visAlleMedarbCHK = ""
    visAlleMedarb = 0
    end if

    else

                    if request.cookies("tsa")("visAlleMedarb") = "1" then
                    visAlleMedarbCHK = "CHECKED"
                    visAlleMedarb = 1
                    else
                    visAlleMedarbCHK = ""
                    visAlleMedarb = 0
                    'usemrn = session("mid")
                    end if
    end if
    response.cookies("tsa")("visAlleMedarb") = visAlleMedarb



    
    call ersmileyaktiv()
	
    call medarb_teamlederfor
    
        
    
        
    varTjDatoUS_man = request("varTjDatoUS_man")

    'if len(trim(request("varTjDatoUS_son"))) <> 0 then
	'varTjDatoUS_son = request("varTjDatoUS_son")
    'else
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)
    'end if

    'Response.Write varTjDatoUS_man
    'Response.end

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





    call positiv_aktivering_akt_fn()
    call smileyAfslutSettings()

	'Response.Write "media " & media
	
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	
    select case func 
	case "-"

    case "opdaterstatus"


        
        'response.write "request(ids)" & request("ids")
        
          ujid = split(request("ids"), ",")

                if len(trim(request("FM_godkendt"))) <> 0 then
                uGodkendt = request("FM_godkendt")
                else
                uGodkendt = 0
                end if
	
    for u = 0 to UBOUND(ujid)
	
	            editor = session("user")


                'if len(trim(request("FM_godkendt_"& trim(ujid(u))))) <> 0 then
                'uGodkendt = request("FM_godkendt_"& trim(ujid(u)))
                'else
                'uGodkendt = 0
                'end if
              

				strSQL = "UPDATE timer SET godkendtstatus = "& uGodkendt &", "_
				&"godkendtstatusaf = '"& editor &"' WHERE tid = " & ujid(u)
				
			    'Response.write strSQL &"<br>"
				
				oConn.execute(strSQL)
				
	next

       'response.end

    'Response.Redirect "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son&"&nomenu="&nomenu

    Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	
	case "godkendugeseddel"
	
	'*** Godkender ugeseddel ***'
	
	    varTjDatoUS_manSQL = year(varTjDatoUS_man)&"/"&month(varTjDatoUS_man)&"/"&day(varTjDatoUS_man)
        varTjDatoUS_sonSQL = year(varTjDatoUS_son)&"/"&month(varTjDatoUS_son)&"/"&day(varTjDatoUS_son)
	    
	    strSQLup = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& session("user") &"' WHERE tmnr = "& usemrn 
	    if cint(SmiWeekOrMonth) = 0 then
        strSQLup = strSQLup & " AND tdato BETWEEN '"& varTjDatoUS_manSQL &"' AND '" & varTjDatoUS_sonSQL & "'" 
        else
        varTjDatoUS_man_mth = datepart("m", varTjDatoUS_man,2,2)
        strSQLup = strSQLup & " AND MONTH(tdato) = '"& varTjDatoUS_man_mth & "'" 
        end if

        strSQLup = strSQLup & " AND godkendtstatus <> 1" 

	    oConn.execute(strSQLup)
	    
        'if session("mid") = 1 then
	    'Response.Write strSQLup
	    'Response.end
        'end if
        
        
        '*** Godkend uge status ****'
        call godekendugeseddel(thisfile, session("mid"), usemrn, varTjDatoUS_man)
	    
	
	
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	

    case "afvisugeseddel"
	
	'*** Afviser ugeseddel ***'
    if len(trim(request("FM_afvis_grund"))) <> 0 then
    txt = replace(request("FM_afvis_grund"), "'", "")
    else
    txt = ""
    end if

	call afviseugeseddel(thisfile, session("mid"), usemrn, varTjDatoUS_man, varTjDatoUS_son, txt)
	    
	
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	
    
    
    case "adviserugeafslutning"

    call afslutugereminder(thisfile, session("mid"), usemrn, varTjDatoUS_man, varTjDatoUS_son, txt)

  
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	

    case "slet_tip"

        %>
        	
        <%

        id = request("id")

        slttxt = "Du er ved at slette en timeregistrering uden match.<br> Er Du sikker på Du vil gøre dette?"
        slturl = rdirfile & "?func=slet_tip_ok&id="&id&"&usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son

       call sltque(slturl,slttxt,slturlalt,slttxtalt,110,90)

    

    case "slet_tip_ok"

        id = request("id")

        sqlInTimerImpTemp = "DELETE FROM timer_import_temp WHERE id = "& id
        oConn.Execute(sqlInTimerImpTemp)

    Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son

    case else



     if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	response.end
    end if

   

    call akttyper2009(2)
	
	
	
	if media <> "export" then
	
	
	if media <> "print" then
	
	
	%>

	
    <SCRIPT src="js/ugeseddel_2011_jav.js"></script>
    <SCRIPT src="../timereg/inc/smiley_jav.js"></script>

    <%call browsertype() 
    if browstype_client <> "ip" then
        
        call menu_2014()

    end if
    %>    


   <%      
         
     'call browsertype()
                
                'if cint(nomenu) <> 1 then
                'call menu_2014() 
                'else   
                'end if
    else 
	

    end if 'print%>
	
	<%end if 'eksport
        
        


   
        %>

 

 <script type="text/javascript" src="js/plugins/flot/jquery.flot.js"></script>
    <script type="text/javascript" src="js/demos/flot/stacked-vertical_ugetotal.js"></script>
   
    <!--<div id="wrapper">
        <div class="to-content-hybrid-fullzize" style="position:absolute; top:102px; left:90px; background-color:#FFFFFF;"> -->

   

	
	
	
	<!-------------------------------Sideindhold------------------------------------->

   
   <%if browstype_client = "ip" then %>

              <div id="header_ugeseddel_2011_ip" style="width: 100%;
              height: 40px;
              color: white;
              background-color: #004d7b;
              margin: 0 0 20px 0;
              line-height: 40px;
              text-align: left;
              font-weight: 600;
              padding-left: 40px;
              background: #004d7b url('img/logo_emblem.png') no-repeat left center;">TimeOut mobile</div>

   <%end if %>



   <%if browstype_client <> "ip" then %>
<div class="wrapper">
 <div class="content">   
     <%end if %>

 
    <div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u><%=tsa_txt_337%></u><!-- ugeseddel -->
        </h3>
        <div class="portlet-body">
            
      
        <%   
                                

            call timerDenneUge(usemrn, lto, varTjDatoUS_man, aty_sql_realhours)

                ugetotal = manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer
                manTimer = replace(manTimer, ",",".")
                tirTimer = replace(tirTimer, ",",".")
                onsTimer = replace(onsTimer, ",",".")
                torTimer = replace(torTimer, ",",".")
                lorTimer = replace(lorTimer, ",",".")
                sonTimer = replace(sonTimer, ",",".")

                call normtimerPer(usemrn, varTjDatoUS_man, 6, 0)

            norm_ugetotal = ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon

            if norm_ugetotal <> 0 then
            'weeknormpro = ugetotal / norm_ugetotal * 100
            weeknormpro = ugetotal / norm_ugetotal * 100
            weeknormpro = replace(weeknormpro, ",",".")
            else
            weeknormpro = 0
            end if

            %>

            <input type="hidden" id="dagman" value="<%=manTimer %>" />
            <input type="hidden" id="dagtir" value="<%=tirTimer %>" />
            <input type="hidden" id="dagons" value="<%=onsTimer %>" />
            <input type="hidden" id="dagtor" value="<%=torTimer %>" />
            <input type="hidden" id="dagfre" value="<%=freTimer %>" />
            <input type="hidden" id="daglor" value="<%=lorTimer %>" />
            <input type="hidden" id="dagson" value="<%=sonTimer %>" />

        
                

            
            <%
          

           
    if media <> "print" AND len(trim(strSQLmids)) > 0 then 'Hvis man er level 1 eller teamleder vil len(trim(strSQLmids)) ALTID VÆRE > 16 %>
    <%
        
        if browstype_client <> "ip" then
    %>
    <div class="well well-white"> 
        
        <%if cint(stempelurOn) = 1 then 
            wdth = 205
        else
            wdth = 100
        end if%>
          
    <div style="position:relative; background-color:#ffffff; border:1px #cccccc solid; border-bottom:0; padding:4px; width:<%=wdth%>px; top:-70px; left:880px; z-index:0;"><a href="../timereg/<%=lnkTimeregside%>" class="vmenu"><%=left(tsa_txt_031, 7) %>.</a>
    
    <%if cint(stempelurOn) = 1 then %>
    &nbsp;| &nbsp;<a href="../timereg/<%=lnkLogind%>" class="vmenu"><%=tsa_txt_340 %></a>
    <%end if
        %>
        </div>


      
       
        <form id="filterkri" method="post" action="ugeseddel_2011.asp">
            <input type="hidden" name="FM_sesMid" id="FM_sesMid" value="<%=session("mid") %>">
            <input type="hidden" name="medarbsel_form" id="medarbsel_form" value="1">
            <input type="hidden" name="nomenu" id="nomenu" value="<%=nomenu %>">
            <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
           
            <div class="row">
                <div class="col-lg-1"><input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> Admin </div>
                <div class="col-lg-5">
                    <% 
				    call medarb_vaelgandre
                    %>
                </div>
            </div>
                  
        </form>
       

	
       
	
	<%
   
       end if 'media
	
        'if datePart("d", now, 2,2) < 4 OR datePart("d", now, 2,2) > 29 then 

	   if media <> "print" then
        %> 
        <%
            
            if browstype_client <> "ip" then
        %>
        <form id="container" action="../timereg/timereg_akt_2006.asp?func=db&rdir=ugeseddel_2011" method="post">
        <div class="row">
            <div class="col-lg-6">
                <table style="font-size:100%; color:black">
                    
                    <tr>
                        <td style="padding-right:5px; vertical-align:text-top;"><b>Dato:</b></td>
                        <td style="padding-left:10px">
                            <input type="hidden" id="Hidden5" name="year" value="<%=year(now) %>"/>

                              <div class='input-group date'>
                                      <input type="text" style="width:300px;" class="form-control input-small" name="FM_datoer" value="<%=day(now) &"/"& month(now) &"/"& year(now) %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>

                            <!--
                            <input style="width:300px;" type="text" id="Hidden1" name="xFM_datoer" value="<%=day(now) &"/"& month(now) &"/"& year(now) %>" class="form-control input-small"/>
                            --> 
                        </td>
                        <td>&nbsp;</td>
                    </tr>
                    
                    <tr>
                        <td style="padding-right:5px; padding-top:10px; vertical-align:text-top;"><b>Kunde/job:</b></td>
                        <td style="padding-top:10px; padding-left:10px;">
                             <input type="hidden" name="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
        
                            <input type="hidden" name="usemrn" id="" value="<%=usemrn%>">

                            <!--<input type="hidden" id="Hidden3" name="FM_dager" value=""/>-->
                            <input type="hidden" id="Hidden4" name="FM_dager" value="0"/><!-- , xx -->
                            <input type="hidden" id="Hidden2" name="FM_feltnr" value="0"/>
                            <input type="hidden" id="FM_pa" name="FM_pa" value="<%=positiv_aktivering_akt_val %>"/>
                            <input type="hidden" id="FM_medid" name="FM_medid" value="<%=usemrn %>"/>
                            <input type="hidden" id="FM_medid_k" name="FM_medid_k" value="<%=usemrn%>"/>
                            <input type="hidden" id="" name="FM_vistimereltid" value="0"/>

                    
                            <input type="text" id="FM_job_0" name="FM_job" placeholder="Kunde/job" class="FM_job form-control input-small"/>
                           <!-- <div id="dv_job_0" class="dv-closed dv_job" style="border:1px #cccccc solid; padding:10px; visibility:hidden; display:none;"></div>--> <!-- dv_job -->

                             <select id="dv_job_0" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none;">
                                 <option>Søger..</option>
                             </select>

                            <input type="hidden" id="FM_jobid_0" name="FM_jobid" value="-1"/>
                        </td>
                     
                    </tr>

                    <tr>
                        <td style="padding-right:5px; padding-top:10px; vertical-align:text-top;"><b>Aktivitet:</b></td>
                         <td style="padding-top:10px; padding-left:10px; width:225px">
                            <input type="text" id="FM_akt_0" name="activity" placeholder="Aktivitet" class="FM_akt form-control input-small"/>
                                <!--<div id="dv_akt_0" class="dv-closed dv_akt" style="border:1px #cccccc solid; padding:10px; visibility:hidden; display:none;"></div>--> <!-- dv_akt -->

                                  <select id="dv_akt_0" class="form-control input-small chbox_akt" size="10" style="visibility:hidden; display:none;">
                                      <option>Søger..</option>
                             </select>

                            <input type="hidden" id="FM_aktid_0" name="FM_aktivitetid" value="-1"/>
                        </td>
                       
                    </tr>

                    <tr>
                        <td style="padding-right:5px; padding-top:10px; vertical-align:text-top;"><b>Timer:</b></td>
                         <td style="padding-top:10px; padding-left:10px">
                             <input type="hidden" id="FM_sttid" name="FM_sttid" value="00:00"/>
                             <input type="hidden" id="FM_sltid" name="FM_sltid" value="00:00"/>
                             <input type="text" id="FM_timer" name="FM_timer" placeholder="Antal timer" class="form-control input-small"/><!-- brug type number for numerisk tastatur -->
                        </td>
                        
                    </tr>

                    <tr>
                        <td style="padding-right:5px; padding-top:10px; vertical-align:text-top;"><b>Kommentar:</b></td>
                        <td style="padding-top:10px; padding-left:10px; text-align:right;">
                            <input type="text" id="FM_kom" name="FM_kom_0" placeholder="Kommentar" class="form-control input-small"/><br />
                            <button class="btn btn-sm btn-success"><b>Indlæs >></b></button>
                        </td>
                       
                    </tr>

      
                </table>
            </div>
           
            <div class="col-lg-5"><div id="stacked-vertical-chart" class="chart-holder-200"></div></div>

        </div>

       

    
        </form>
        <%end if %>
        

        </div>
            <%end if  %>


        <%
            
            if browstype_client = "ip" then
            %>
            <div class="row">
            <div class="col-lg-12"><div id="stacked-vertical-chart" class="chart-holder-200"></div></div>
            </div>
            <%end if %>

            
        <%
            end if

        'end if
    

    
   ugeseddelvisning = 1

   perInterval = 6 'dateDiff("d", varTjDatoUS_man, varTjDatoUS_son, 2,2) 
   perIntervalLoop = perInterval

   'response.write "perIntervalLoop " & perIntervalLoop
   for l = 0 to perIntervalLoop 
        
        if l = 0 then
        varTjDatoUS_use = varTjDatoUS_man
        showheader = 1
        showtotal = 0
        else
        varTjDatoUS_use = dateAdd("d", l, varTjDatoUS_man)
        showheader = 0
        showtotal = 0
         end if 

       if l = perIntervalLoop then
       showtotal = 1
       end if

         varTjDatoUS_son = varTjDatoUS_man

        call ugeseddel(usemrn, varTjDatoUS_use, varTjDatoUS_use, ugeseddelvisning, showtotal, showheader)  
   
   next
	
	
        select case lto
        case "tec", "intranet - local", "esn"
        aty_sql_realhours = " tfaktim <> 0"
        case else
        aty_sql_realhours = aty_sql_realhours &""_
		& " OR tfaktim = 30 OR tfaktim = 31 OR tfaktim = 7 OR tfaktim = 11"
        end select 
   
                                

            'response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>aty_sql_realhours "& aty_sql_realhours
            'response.flush

            call timerDenneUge(usemrn, lto, varTjDatoUS_man, aty_sql_realhours)

                ugetotal = manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer
                manTimer = replace(manTimer, ",",".")
                tirTimer = replace(tirTimer, ",",".")
                onsTimer = replace(onsTimer, ",",".")
                torTimer = replace(torTimer, ",",".")
                lorTimer = replace(lorTimer, ",",".")
                sonTimer = replace(sonTimer, ",",".")

                call normtimerPer(usemrn, varTjDatoUS_man, 6, 0)

            norm_ugetotal = ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon


            'weeknormpro = ugetotal / norm_ugetotal * 100
            if norm_ugetotal <> 0 then
                
           
                weeknormpro = ugetotal / norm_ugetotal * 100
                weeknormpro = formatnumber(weeknormpro, 0)
                weeknormpro = replace(weeknormpro, ",",".")
                

             if ugetotal <= norm_ugetotal then
             weeknormproWdt = weeknormpro
             else
             weeknormproWdt = norm_ugetotal / norm_ugetotal * 100 
             end if

            else
            weeknormpro = 0
            end if

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

          <div class="row">
                <div class="col-lg-12">
                    <h4 style="text-align:right">Norm timer: <%=norm_ugetotal %> &nbsp  Opgjort: <%=ugetotal %> &nbsp&nbsp</h4>
                    <table class="table table">
                        <tr>
                            <th colspan="11"><div style="background-color:lightgray; height:30px; width: 100%;"><div style="background-color:#207800; height:30px; width:<%=weeknormproWdt %>%; color:#ffffff; text-align:right; padding:5px;"><%=weeknormpro %> %</div></div></th>
                        </tr>
                        <tr>
                            <th style="border-top:none; width:10%;"><h6>00</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>20</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>40</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>60</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>80</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>100</h6></th>
                                        
                        </tr>
                       <!-- <tr>
                            <th style="text-align:left; border-top:none; font-size:85%" colspan="11">Norm timer: <%=norm_ugetotal %></th>
                        </tr>
                        <tr>
                            <th style="text-align:left; border-top:none; font-size:85%" colspan="11">Opgjort: <%=weeknormpro %>%</th>
                        </tr> -->
                    </table>
                </div>
            </div>







            <%if media <> "print" then %>

                <!--
                <form action="ugeseddel_2011.asp?media=export" method="post" target="_blank">
                    
                 <div class="row">
                     <div class="col-lg-12 pad-r30">
                         <input id="Submit6" type="submit" value="Eksport til .csv simpel" class="btn btn-sm" />
                      </div>


                </div>
                    </form>
        -->

          
            <br /><br />


         <%if browstype_client = "ip" then %>
           
            <br /><br /><br /><br />
            <form action="../timetag_web/timetag_web.asp">
                <div class="row">
                    <div class="col-lg-12">
                       <!-- <a href="../timetag_web/timetag_web.asp?flushsessionuser=1" class="btn btn-sm btn-default" style="text-align:center; width:100%"><b>Tilbage</b></a> -->
                        <button class="btn btn-sm btn-default" style="text-align:center; width:50%"><b><< Tidsregistrering</b></button>
                    </div>
                </div>
            </form>
           
        <%end if %>

             
        <%
            
            if browstype_client <> "ip" then
        %>
    <div class="well" style="width:35%;">
      <div class="portlet">
        <h3 class="portlet-title">
          <u>Funktioner:</u><!-- ugeseddel -->
        </h3>

          

           

              <form action="ugeseddel_2011.asp?media=print&usemrn=<%=usemrn %>&varTjDatoUS_man=<%=varTjDatoUS_man%>" method="post" target="_blank">
                    
                 <div class="row">
                     <div class="col-lg-12 pad-r30">
                         <input id="Submit6" type="submit" value="Printvenlig >>" class="btn btn-sm btn-default" />
                      </div>


                </div>
                    </form>


         

        


          
<table>

    <tr><td colspan="2">

           <!-- hvis level = 1 OR teamleder -->
        <%if (level <=2 OR level = 6) AND cint(smilaktiv) <> 0 then %>
          

        <%
        '** Tjekker for Uge 53. SKAL i virkeligheden være om søndag er i et andet år end mandag - da år så skifter.
        if datepart("ww", varTjDatoUS_son, 2,2) = 53 then
        tjkAar = year(varTjDatoUS_son) + 1
        else
        tjkAar = year(varTjDatoUS_son)
        end if    
            
        call erugeAfslutte(tjkAar, datepart("ww", varTjDatoUS_son, 2,2), usemrn) %>
        <%call godkendugeseddel(fmlink, usemrn, varTjDatoUS_man, rdir) %>
    
        <%end if 'level %>
    

             </td>
        </tr>



</table>
  </div> <!-- well -->
  </div> <!-- portlet title -->
            <%end if %>
      

<%else

    
Response.Write("<script language=""JavaScript"">window.print();</script>")
    
end if 

      %>


    </div>

    <%
       '** Wrapper / Content
        if browstype_client <> "ip" then %>
      </div></div>
     <%end if %>


   
    <br /><br />&nbsp;
    <%
	
	
	end select
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
