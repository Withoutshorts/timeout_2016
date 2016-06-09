  

    <%
    
    if request("usegl2006") = "0" then
    Response.Cookies("tsa")("usegl2006") = "0"
    else
        if request.Cookies("tsa")("usegl2006") = "1" AND lto <> "dencker" then
        Response.Redirect "../timereg2006/timereg_2006_fs.asp?usegl2006=1"
        else
        Response.Cookies("tsa")("usegl2006") = "0"
        end if
    end if
    
    'Response.Write "request.Cookies(tsa)(usegl2006) " & request.Cookies("tsa")("usegl2006")
    %>
    
    
    <!--#include file="../inc/connection/conn_db_inc.asp"-->
    <!--#include file="../inc/xml/tsa_xml_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/global_func.asp"-->
	<!--#include file="inc/convertDate.asp"-->
	<!--#include file="inc/timereg_akt_2006_inc.asp"-->
	<!--#include file="inc/smiley_inc.asp"-->
	<!--#include file="inc/isint_func.asp"-->
    <!--#include file="../inc/regular/treg_func.asp"-->
	
	
	
	<%

    
	
	
	   '**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_fjernjob"
                  
                 jobid = request("jobid")   
                 medid = request("usemrn")

                 '** Nulstillerforvalgt job for denne medarbejder ****'
                 strSQLUpdFvOff = "UPDATE timereg_usejob SET forvalgt = 0 WHERE medarb = "& medid &" AND jobid = "& jobid 
                 oConn.Execute(strSQLUpdFvOff)
                    '******************************************************'

        case "FM_sortOrder"
               Call AjaxUpdateTreg("timereg_usejob","forvalgt_sortorder","")
        
        
        case "FN_updatejobkom"
        
                 jobid = request("jobid")   
                 jobkom = request("jobkom") 
                 jobkom = replace(jobkom, "''", "") 
                 jobkom = replace(jobkom, "'", "")
                 jobkom = replace(jobkom, "<", "")
                 jobkom = replace(jobkom, ">", "")

                 oldKom = ""

                 strSQLUpdj = "SELECT kommentar FROM job WHERE id = "& jobid
                 oRec3.open strSQLUpdj, oConn, 3
                 if not oRec3.EOF then 
                 oldKom = oRec3("kommentar")
                 end if
                 oRec3.close

                 session_user = request("session_user")
                 dt = request("dt")

                 jobkom = "<span style=""color:#999999;"">" & dt &", "& session_user &":</span> " & jobkom & "<br>"& oldKom
                 
                 
                 '** Job kommentar ****'
                 strSQLUpdj = "UPDATE job SET kommentar = '"& jobkom &"' WHERE id = "& jobid 
                 oConn.Execute(strSQLUpdj)
                 
                 
                 
                
        case "FN_jobidenneuge"
         

         stDato = request("stDato")
         slDato = request("slDato")

         stDato = year(stDato) & "/"& month(stDato) & "/"& day(stDato)
         slDato = year(slDato) & "/"& month(slDato) & "/"& day(slDato)

         strJoiDUge =  "<table cellspacing=0 cellpadding=0 border=0 width='100%'>"



         strSQLfv = "select j.jobnavn, j.jobnr, j.jobslutdato FROM job AS j WHERE jobstatus = 1 AND jobslutdato BETWEEN '"& stDato &"' AND '"& slDato &"' ORDER BY jobslutdato DESC"
         
         'Response.write strSQLfv
         'Response.Flush

         lastSlutDato = 0
         oRec3.open strSQLfv, oConn ,3
         while not oRec3.EOF 
         
         
         if lastSlutDato <> oRec3("jobslutdato") OR lastSlutDato = 0 then
         strJoiDUge = strJoiDUge &"<tr bgcolor=""#DCF5BD""><td class=lille><b>"&oRec3("jobslutdato")&"</b></td></tr>"
         end if 

        
         strJoiDUge = strJoiDUge &"<tr><td class=lille style=""border-bottom:1px #CCCCCC solid;"">"& oRec3("jobnavn") &" ("& oRec3("jobnr") &")</td></tr>"

        

         lastSlutDato = oRec3("jobslutdato")
         oRec3.movenext
         wend
         oRec3.close  

        strJoiDUge = strJoiDUge &"<tr bgcolor=""#EfF3FF""><td><b>Aktiviteter:</b><br /><span style=""font-size:9px;"">Hvor slutdato afviger fra job</span></td></tr>"


          strSQLfv = "SELECT a.navn, aktslutdato, j.jobnavn, j.jobnr FROM aktiviteter AS a LEFT JOIN job AS j ON (j.id = a.job) "_
          &" WHERE aktstatus = 1 AND aktslutdato <> jobslutdato AND aktslutdato BETWEEN '"& stDato &"' AND '"& slDato &"' ORDER BY aktslutdato DESC"
         
         'Response.write strSQLfv
         'Response.Flush

         lastSlutDato = 0
         oRec3.open strSQLfv, oConn ,3
         while not oRec3.EOF 
         
         
         if lastSlutDato <> oRec3("aktslutdato") OR lastSlutDato = 0 then
         strJoiDUge = strJoiDUge &"<tr bgcolor=""#DCF5BD""><td class=lille><b>"& oRec3("aktslutdato") &"</b></td></tr>"
         end if 

         
         strJoiDUge = strJoiDUge & "<tr><td class=lille style=""border-bottom:1px #CCCCCC solid;"">"& oRec3("navn") &"<br />"_
         &"<span style=""font-size:9px; color:#999999;"">"& oRec3("jobnavn") &" ("& oRec3("jobnr") &")</span></td></tr>"


         lastSlutDato = oRec3("aktslutdato")
         oRec3.movenext
         wend
         oRec3.close  

         strJoiDUge = strJoiDUge &"</table>"       


          '*** ÆØÅ **'
          call jq_format(strJoiDUge)
          strJoiDUge = strJoiDUge
                

         Response.write strJoiDUge

        case "FN_opdtodo"
                
                if len(trim(request.form("cust"))) <> 0 then
                todoID = request.form("cust")
                else
                todoId = 0
                end if
                
                strSQL = "UPDATE todo_new SET afsluttet = 1 WHERE id = "& todoId & " OR parent = "& todoId
                oConn.execute(strSQL)
                
                'Response.redirect "timereg_akt_2006.asp" 
                
        case "FM_get_destinations"
        
            	if len(trim(request("kid"))) <> 0 then
				kid = request("kid") 
				else
				kid = 0
				end if
				
				if len(trim(request("xvalbegin"))) <> 0 then
				xvalbegin = (request("xvalbegin") + 1)/1
				else
				xvalbegin = 2
				end if
				
				    'visalle = request("visalle")
				    if len(trim(request("visalle"))) <> 0 then
				    visalle = request("visalle")
				    else
				    visalle = 0
				    end if
				    
				    if visalle <> 0 then
				        strSQLKundKri = " AND kid <> 0"
				    else
				        strSQLKundKri = " AND kid = " & kid
				    end if
				
				 more = request("more")
				
				'Response.Write "visalle: "& request("visalle")  & "<br>"
				
				
				if len(trim(request.form("cust"))) <> 0 then
				sogeKri = " (kp.navn LIKE '"& request.form("cust") &"%' OR k.kkundenavn LIKE '"& request.form("cust") &"%')"
				else
				sogeKri = " kkundenavn <> 'sdfsdfd556irp'"
				end if
				
				
				strSQL2 = "SELECT kp.id, kp.navn, kp.email, kp.dirtlf, kp.mobiltlf, "_
				&" kp.adresse As kpadr, kp.postnr As kpzip, kp.town As kptown, kp.land As kpland, kp.afdeling AS kpafd, "_
				&" k.kid, k.kkundenavn, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, k.land FROM kunder k "_ 
				&" LEFT JOIN kontaktpers kp ON (kp.kundeid = k.kid)"_
				&" WHERE "& sogeKri & strSQLKundKri & " ORDER BY k.kkundenavn, kp.navn LIMIT 10"
			   
			    'Response.Write strSQL2 & "<br><br>"
			    'Response.flush
			   
			    strFil_Kon_Txt = ""
			   
			    oRec2.open strSQL2, oConn, 3
                x = xvalbegin
                k = 0
                while not oRec2.EOF
                
                
                        if ((lastknavn <> oRec2("kkundenavn") AND instr(lcase(oRec2("kkundenavn")), lcase(request.form("cust"))) <> 0)) OR (cint(more) = 0 AND cint(k) = 0) then 'lastknavn <> oRec2("kkundenavn")) then
                            strFil_Kon_Txt = strFil_Kon_Txt &  "<br /><input id='ko"&x&"chk' type=checkbox />"_ 
                            &"<select id='ko"&x&"sel' style='font-size:9px;'>"_
                            &"<option value=0>-</option>"_
                            &"<option value=2>"&tsa_txt_363&"</option>"_
                            &"<option value=3>"&tsa_txt_364&"</option>"_
                            &"<option value=1>"&tsa_txt_362&"</option>"_
                            &"</select><input id='ko"&x&"kid' type=hidden value='k_"&oRec2("kid")&"' />"_ 
                            &"<b>&nbsp;"& left(oRec2("kkundenavn"), 30) &"</b> ("&oRec2("kkundenr")&") " & tsa_txt_370
                         
                           strFil_Kon_Txt = strFil_Kon_Txt &  "<br /><img src='../ill/blank.gif' width=1 height=5 /><br>"_
                           &"<div id='ko"&x&"' style=""padding:5px; border:1px #cccccc solid; top:5px; background-color:#DCF5BD;"">"
                        
                                if len(trim(oRec2("kkundenavn"))) <> 0 then
                                strFil_Kon_Txt = strFil_Kon_Txt &  oRec2("kkundenavn") &"<br />"_
                                & oRec2("adresse") &"<br />"_
                                & oRec2("postnr") &" "& oRec2("city")&"<br />"_
                                & oRec2("land") &"<br />"
                                end if
                        
                        strFil_Kon_Txt = strFil_Kon_Txt &  "</div>"
                        strFil_Kon_Txt = strFil_Kon_Txt &  "<br />"
                            
                        x = x + 1   
                        k = 1 
                        end if
                        
                
                                if len(trim(oRec2("id"))) <> 0 AND ((cint(more) = 0) OR (instr(lcase(oRec2("navn")), lcase(request.form("cust"))) <> 0) AND cint(more) = 1 ) then
                                
                                strFil_Kon_Txt = strFil_Kon_Txt &  "<br /><input id='ko"&x&"chk' type=checkbox />"_ 
                                &"<select id='ko"&x&"sel' style='font-size:9px;'>"_
                                &"<option value=0>-</option>"_
                                &"<option value=2>"&tsa_txt_363&"</option>"_
                                &"<option value=3>"&tsa_txt_364&"</option>"_
                                &"<option value=1>"&tsa_txt_362&"</option>"_
                                &"</select><input id='ko"&x&"kid' type=hidden value='kp_"&oRec2("id")&"' />"_ 
                                &"<b>&nbsp;"& oRec2("navn")&"</b> " & tsa_txt_369 & ""_
                                &" <span style=""color:#999999;"">(<b>"& left(oRec2("kkundenavn"), 30)&"</b> "&oRec2("kkundenr")&")</span>"
                                    
                                           
                                           
                                        strFil_Kon_Txt = strFil_Kon_Txt &  "<br /><img src=""../ill/blank.gif"" width=1 height=5 /><br>"_
                                        &"<div id='ko"&x&"' style=""padding:5px; border:1px #cccccc solid; top:5px; background-color:#EFF3FF;"">"
                                        if len(trim(oRec2("navn"))) <> 0 then
                                        strFil_Kon_Txt = strFil_Kon_Txt &  oRec2("navn") &"<br />"
                                        
                                        if len(trim(oRec2("kpafd"))) <> 0 then
                                        strFil_Kon_Txt = strFil_Kon_Txt &  oRec2("kpafd") &"<br />"
                                        end if
                                        
                                        strFil_Kon_Txt = strFil_Kon_Txt &  oRec2("kpadr") &"<br />"_
                                        & oRec2("kpzip") &" "& oRec2("kptown") &"<br />"_
                                        & oRec2("kpland") &"<br />"
                                        end if
                                        
                                        strFil_Kon_Txt = strFil_Kon_Txt &  "</div>"
                                
                                'if len(trim(oRec2("email"))) <> 0 then
                                'strFil_Kon_Txt = strFil_Kon_Txt &  tsa_txt_025 &": <a href='mailto:"&oRec2("email")&"' class=rmenu>"& oRec2("email") &"</a><br>"
                                'end if
                                
                                'if len(trim(oRec2("dirtlf"))) <> 0 then
                                'strFil_Kon_Txt = strFil_Kon_Txt &  tsa_txt_026 &": "& oRec2("dirtlf") &"<br>"
                                'end if
                                
                                'if len(trim(oRec2("mobiltlf"))) <> 0 then
                                'strFil_Kon_Txt = strFil_Kon_Txt &  tsa_txt_027 &": "& oRec2("mobiltlf") &"<br><br />"
                                'end if
                                
                                x = x + 1
                                end if 'oRec2("id") <> 0
                
                lastknavn = oRec2("kkundenavn") 
                oRec2.movenext
                wend
                oRec2.close
                
                
                
                 '*** ÆØÅ **'
                call jq_format(strFil_Kon_Txt)
                strFil_Kon_Txt = jq_formatTxt
                
                Response.Write strFil_Kon_Txt & "<input id=""kperfil_fundet"" type=""hidden"" value='"&x-1&"'>" 
                '&"</div><input type=""text"" value='"&x&"'>"
                
                 
                 
                 
                 
                    
        
        case "FN_getCustDesc"
        
        'Response.end
        
        'Response.Write "<select>"
        if len(trim(request("usemrn"))) <> 0 then
        usemrn = request("usemrn")
        else
        usemrn = session("mid")
        end if
        
        
        if len(trim(request("ignprg"))) <> 0 then
        ignProj = request("ignprg")
        else
        ignProj = 0
                
                '**** Henter projektgrupper ***
                '*** ER allerede hentet **'
                '** Side SKAL re-loades ved skift medarbejder ***' eller??
				'call hentbgrppamedarb(usemrn)
        
        end if
        
        
        if len(trim(request("visnyeste"))) <> 0 then
        visnyeste = request("visnyeste")
        else
        visnyeste = 0
        end if
        
        '*** Easyreg ***'
        if len(trim(request("viseasyreg"))) <> 0 then
        viseasyreg = request("viseasyreg")
            if cint(viseasyreg) = 1 then
            easyregFlt = " ,COUNT(a.id) AS antal_aeasy "
            easyWh = " AND tu.easyreg <> 0"
            else
            easyregFlt = ""
            easyWh = ""
            end if
        
        else
        viseasyreg = 0
        easyregFlt = ""
        easyWh = ""
        end if

        '** dato
        ugeStDato = trim(request("ugeStDato"))
        ugeStDato = year(ugeStDato) &"/"& month(ugeStDato) &"/"& day(ugeStDato)
        
        if cint(visnyeste) = 1 then
        sqlOrderBy = "j.jobstartdato DESC, j.jobnavn"
        lmt = " LIMIT 0, 25"
        dtWhcls = " AND j.jobstartdato <= '"& year(now) & "/"& month(now) & "/"& day(now) &"' "
        else
            
     
        
        if cint(viseasyreg) = 1 then
        sqlOrderBy = ""
        lmt = ""
        dtWhcls = ""
        else
        sqlOrderBy = ""
        lmt = " LIMIT 0, 100"
        'lmt = ""
        dtWhcls = ""
        end if


        sortby = request("sortby")
        if sortby = 1 then
        sqlOrderBy = sqlOrderBy & "j.jobnavn "
        end if
        
        if sortby = 2 then
        sqlOrderBy = sqlOrderBy & "j.jobnr DESC "
        end if
        
        if sortby = 3 then
        sqlOrderBy = sqlOrderBy & "k.kkundenavn, j.jobnr DESC "
        end if
        
        if sortby = 0 then
        sqlOrderBy = sqlOrderBy & "k.kkundenavn, j.jobnavn "
        end if
      
        
        end if

        
        
        
        
        'Response.Write "usemrn "& usemrn & " ignProj " & ignProj & "easyreg " & viseasyreg
        'Response.end
        
        if len(trim(request("jobsog"))) <> 0 then
        jobsog = request("jobsog")
        else
        jobsog = ""
        end if
       
		if len(trim(jobsog)) <> 0 then 		
		strJobSogKri = " AND (jobnr LIKE '"& jobsog &"%' OR jobnavn LIKE '%"& jobsog &"%' OR kkundenavn LIKE '%"& jobsog &"%' OR kkundenr = '"& jobsog &"') "
		else
		strJobSogKri = ""
		end if				
        
        
        if cint(ignProj) = 0 then
        strSQLwh = " (j.jobstatus = 1 OR j.jobstatus = 3) "& strPgrpSQLkri & "" 
	    else
	    '** Henter alle aktive job + tilbud **'
	    strSQLwh = " (j.jobstatus = 1 OR j.jobstatus = 3) " 
	    end if
        
        if Request.Form("cust") <> 0 AND strJobSogKri = "" then
        strSQLwh = strSQLwh & " AND j.jobknr = "& Request.Form("cust")
        else
        strSQLwh = strSQLwh & " AND j.jobknr <> 0 " '"& Request.Form("cust")
        end if

        

        strSQLwh = strSQLwh & " AND (j.jobstartdato <= '"& ugeStDato &"')"
        
        if cint(ignProj) = 0 then
        
            strSQL = "SELECT j.id, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr "& easyregFlt &" FROM timereg_usejob AS tu "_
            &"LEFT JOIN job AS j ON (j.id = tu.jobid) "
            
            if cint(viseasyreg) = 1 then
            strSQL = strSQL &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND a.easyreg = 1) "
            end if
            
            strSQL = strSQL &"LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
            &"WHERE tu.medarb = "& usemrn &" "& easyWh &" AND "_
            & strSQLwh &" "& strJobSogKri &" "& dtWhcls &" AND kkundenavn <> '' GROUP BY j.id ORDER BY "& sqlOrderBy & lmt
            
        else
            
            '** Timereg use jober ignoreret her. Både for job og easyreg akt.
                
            strSQL = "SELECT j.id, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr "& easyregFlt &" FROM job AS j "_
            &"LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "
            
            if cint(viseasyreg) = 1 then
            strSQL = strSQL &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND a.easyreg = 1) "
            end if
            
            strSQL = strSQL & "WHERE "& strSQLwh &" "& strJobSogKri &" "& dtWhcls &" AND kkundenavn <> '' GROUP BY j.id ORDER BY " & sqlOrderBy & lmt
            
        end if
        
        
        if len(trim(request("seljobid"))) <> 0 then
        seljobidArr = request("seljobid")
             
            jobids = split(seljobidArr, ",") 
            for j = 0 to UBOUND(jobids)
            seljobid = seljobid & ",#"& jobids(j) &"#"
            next


        else
        seljobid = "#0#"
        end if
        
        'Response.Write strSQL & " seljobid: " & seljobid
        'Response.write "</option></select>"
         'Response.end
        
        lastknr = 0
        jobids_easyreg = 0
        j = 0


                 

        oRec.open strSQL, oConn, 3
        while not oRec.EOF
        
            if cint(viseasyreg) = 1 then
            antal_easy = oRec("antal_aeasy")
            else
            antal_easy = 0
            end if
            
            if cint(viseasyreg) = 0 OR (cint(viseasyreg) = 1 AND cint(antal_easy) > 0) then
            
                if instr(seljobid, "#"& oRec("id") &"#") <> 0 then
	            jSel = "SELECTED"
                else
	            jSel = ""
	            end if
        	    
        	    
	            delit = "..................................................................................................................................."
                'delit = "...................."
                
                if sortBy <> 2 AND sortBy <> 3 then
                jobnavnid = replace(oRec("jobnavn"), "'", "") & " ("& oRec("jobnr") &")"
                else
                jobnavnid = replace(oRec("jobnr"), "'", "") & " - "& oRec("jobnavn") 
                end if

                if oRec("jobstatus") = 3 then
                jobnavnid = jobnavnid & " - Tilbud"
                end if
                
                '*** ÆØÅ **'
                jq_format(jobnavnid)
                jobnavnid = jq_formatTxt
                
                
                
                len_jobnavnid = len(jobnavnid)
                if len(len_jobnavnid) < 50 then
                lenst = 50
                else
                lenst = len_jobnavnid + 2
                end if 
                
                if len(lenst) <> 0 AND len(len_jobnavnid) <> 0 AND (lenst-len_jobnavnid) > 0 then 
                delit_antal = left(delit, (lenst-len_jobnavnid)) 
                else
                delit_antal = "...................."
                end if
                
                jq_format(oRec("kkundenavn"))
                knavn = jq_formatTxt
                
                
                
                if (sortby = 0 OR sortby = 3) AND lastknr <> oRec("kkundenr") AND cint(visnyeste) <> 1 then
                if j <> 0 then
                Response.Write "<option value='"& oRec("id") &"' disabled></option>" 
                end if
                Response.Write "<option value='"& oRec("id") &"' disabled>"& knavn &" ("& oRec("kkundenr") &")</option>" 
                end if
                
                Response.Write "<option value='"& oRec("id") &"' "&jSel&" >"& jobnavnid &" "& delit_antal &" "& knavn &" ("& oRec("kkundenr") &")</option>" 
                
                if right(j, 1) = 0 then
                Response.flush
                end if
                
                lastknr = oRec("kkundenr")
                
                'jobids_easyreg = jobids_easyreg &", " & oRec("id")
                
                j = j + 1
            
            end if
        
        oRec.movenext
        wend
        oRec.close
        
        if j = 0 then
          Response.Write "<option value='0'>Ingen aktive job matcher din s&oslash;gning!</option>"
          Response.Write "<option value='0'>Tjek dine projektgrupper, din personlig aktivliste og job startdato.</option>"  
        end if
        
       
        
        'Response.Write "</select>"
        '<input id='jobids_easyreg' type='text' value='"& jobids_easyreg &"' />"
        'Response.Write strSQL
        end select
        Response.end
        end if  
	
	


    'loadA =  now
    

	
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	if len(trim(request("fromsdsk"))) <> 0 then
	fromsdsk = request("fromsdsk")
	fromdeskKomm = request("sdskkomm")
	else
	fromsdsk = 0
	fromdeskKomm = ""
	end if
	
	print = request("print")
	
	thisfile = "timereg_akt_2006"
	func = request("func")
	
	if len(request("stopur")) <> 0 then
	stopur = request("stopur")
	else
	stopur = 0
	end if
	
	
	
	
	call ersmileyaktiv()
	
	'******************************************************
	'*************** Opdaterer timeregistreringer *********
	'******************************************************
	if func = "db" then
	
	if len(trim(request("jobid"))) <> 0 then
	jobid = request("jobid")
	else
	jobid = 0
	end if
	
	'if len(request("showjobinfo")) <> 0 then
	'showjobinfo = request("showjobinfo")
	'else
	'showjobinfo = 0
	'end if
	'Response.Cookies("showjobinfo") = showjobinfo
	
	if len(trim(request("tildelalle"))) <> 0 then
	multitildel = 1
	else
	multitildel = 0
	end if
	
	if len(trim(request("tildeliheledage"))) <> 0 then
	tildeliheledage = 1
	else
	tildeliheledage = 0
	end if
	
	
	'*** Valgte medarbejdere ****'
    dim tildelselmedarb
    'redim tildelselmedarb(200)
	
    if len(trim(request("tildelselmedarb"))) <> 0 AND cint(multitildel) = 1 then
    tildelselmedarb = split(request("tildelselmedarb"), ", ")
    else
    tildelselmedarb = split(1, ", ")
    end if
	
	
	            '*** => Timer er indtastet på timereg. siden ***'
	            '*** Eller via stopur                        ***'
	            if stopur <> "1" then 
	            
	                
	                
	                '*** Sætter cookies på faser ****'
	                faseshowVal = split(request("faseshow"), ", ")
	                faserids = split(request("faseshowid"), ", ")
	                for f = 0 to UBOUND(faserids)
	                    'thisFaseVal = request("faseshow_"& faserids(f) &"")
	                    'Response.Write "#"& faserids(f) & "#" &faseshowVal(f) & "<br>"
	                    Response.Cookies("tsa_fase")(""& trim(faserids(f)) &"") = trim(faseshowVal(f))
	                
	                next
	                
	                
	                'Response.end
	                'Response.write "Opdaterer DB<br><br>"
                	
	                strdag = request("FM_start_dag")
	                strmrd = request("FM_start_mrd")
	                straar = request("FM_start_aar")
                	
	                strYear = Request("year")
                	
                	'Response.Write "request(FM_jobid): " & request("FM_jobid") & "<br>"
                	
                	
                	'Response.Write " St: " & request("FM_sttid") & "<br>" 
                	
                	
	                jobids = split(request("FM_jobid"), ",")
	                
	                aktids = split(request("FM_aktivitetid"), ",")
	                
	                medarbejderid = request("FM_mnr")
                	
	                tTimertildelt_temp = replace(request("FM_timer"), ", #,", ";")
	                tTimertildelt_temp2 = replace(tTimertildelt_temp, ", #", "")
	                tTimertildelt = split(tTimertildelt_temp2, ";") 
                	
                	'Response.Write "tTimertildelt: " & request("FM_timer") & "<br>"
                	'Response.flush
                	
	                datoer = split(request("FM_datoer"), ",")
	                
	                'Response.Write "datoer:" & datoer(0)
                	'Response.end
                	
	                tSttid = split(request("FM_sttid"), ",") 
	                tSltid = split(request("FM_sltid"), ",") 
                	
                	
                	'Response.Write tSttid &" til: "& tSltid 
                	'Response.end
                	
                	'Response.Write "Feltnr: "& request("FM_feltnr") & "<hr>"
	                feltnr = split(request("FM_feltnr"), ",")
                	for y = 0 to UBOUND(feltnr)
	                high_fetnr = feltnr(y)
	                z = y
	                next
                	
                	
                	'Response.Write  "<br>Antal felter:" & z & "<br>"
                	
	                if len(request("FM_vistimereltid")) <> 0 then
	                visTimerelTid = request("FM_vistimereltid")
	                else
	                visTimerelTid = 0 '** Vis timer 
	                end if
	                
	                
	               
                	
	                 
	                
                                '**** ETS ****'
			                    '***** Kode til beregning af tidslås på timreg. med klokkeslet angivelse her ***'          
			                    '*** Skal array forlænges??? ***'
            				    
				    
				                 
				                 if lto = "ets-track" then 'lto = "intranet - local"
			                       
			                                        '*** Sletter indtastninger på den valgte uge på et valge job,
                            		                '*** så man ikke risikerer at have indtastninger stående i et
                            		                '*** tidsrum der ved redigering ikke længere er i det indstastede tidsinterval ***'
                            		                '*** Kun nat, dat etc og E1 typen. IKKE på grundregistreringen ***'
                                                    thisJobid = 0
                                                    strSQLjobid = "SELECT jobnr FROM job WHERE id = "& jobids(0)
                                                    'Response.Write strSQLjobid
                                                    'Response.flush
                                                    
                                                    oRec.open strSQLjobid, oConn, 3
                                                    if not oRec.EOF then
                                                    thisJobnr = oRec("jobnr")
                                                    end if
                                                    oRec.close
                                                    
                                                    vlgtDato = strDag &"/"& strMrd &"/"& strAar
                                                    select case weekday(vlgtDato,2)
                                                    case 1
                                                    stDeldato = dateadd("d",0, vlgtDato)
                                                    slDeldato = dateadd("d",6, vlgtDato)
                                                    case 2
                                                    stDeldato = dateadd("d",-1, vlgtDato)
                                                    slDeldato = dateadd("d",5, vlgtDato)
                                                    case 3
                                                    stDeldato = dateadd("d",-2, vlgtDato)
                                                    slDeldato = dateadd("d",4, vlgtDato)
                                                    case 4
                                                    stDeldato = dateadd("d",-3, vlgtDato)
                                                    slDeldato = dateadd("d",3, vlgtDato)
                                                    case 5
                                                    stDeldato = dateadd("d",-4, vlgtDato)
                                                    slDeldato = dateadd("d",2, vlgtDato)
                                                    case 6
                                                    stDeldato = dateadd("d",-5, vlgtDato)
                                                    slDeldato = dateadd("d",1, vlgtDato)
                                                    case 7
                                                    stDeldato = dateadd("d",-6, vlgtDato)
                                                    slDeldato = dateadd("d",0, vlgtDato)
                                                    end select
                                                    
                                                    stDeldato = year(stDeldato) &"/"& month(stDeldato) &"/"& day(stDeldato)
                                                    slDeldato = year(slDeldato) &"/"& month(slDeldato) &"/"& day(slDeldato)
                                                    
                                                    for m = 0 to UBOUND(tildelselmedarb)
	
	                                                'Response.Write tildelselmedarb(m) & "<br>"
                                                	
	                                                if cint(multitildel) = 1 then
	                                                medarbejderid = tildelselmedarb(m) 
	                                                else
	                                                medarbejderid = medarbejderid
	                                                end if
	                                                
                                                    strSQLdelweek = "DELETE FROM timer WHERE tjobnr = "& thisJobnr &" AND "_
                                                    &" tdato BETWEEN '"& stDeldato &"' AND '"& slDeldato &"' AND tmnr = " & medarbejderid & ""_
                                                    &" AND (tfaktim = 50 OR tfaktim = 51 OR tfaktim = 52 OR tfaktim = 53 OR tfaktim = 54 OR tfaktim = 55 OR tfaktim = 90 OR tfaktim = 91)"  
                            				
                            				        'Response.Write strSQLdelweek & "<br><br>"
                            				        oConn.execute(strSQLdelweek)
                            				        
                            				        next
                            				        'Response.end   
			                       
			                       
			                       
			                       
			                       
			                       
			                       
			                        
			                    
	                             'newArrSizeTim = tTimertildelt
	                             icc = 0
	                             cta = 0
			                     jarrforl = ""
			                     for j = 0 to UBOUND(jobids) 'antal akt linier
			             
			                        
			                        if j = 0 then
			                        ystart = 0
			                        else
			                        ystart = (j * 7) 
			                        end if
                        			
			                        yslut = (ystart + 6)
			                     
			                     
			                     'gnlob = 1
			                     'oprArrSize = ubound(tTimertildelt)
			                     lasty2 = 0
			                     for y = ystart to yslut    
			                         
			                         
			                                        '*** Nulstiller alle DAG / NAT / E1 typer så der ikke kan indtastes timer på disse typer
                            			            '*** Selvom de evt. er åbne og der tastes timer ind i felterne ***'
                            			            'if t = 9999 then
                            			            aktType = 0
                            			            strSQLtype = "SELECT fakturerbar "_
                                                    &" FROM aktiviteter WHERE "_
                                                    &" id = "& aktids(j) 
                                                    
                                                    'Response.Write strSQLtype & "<br>"
                            		                
                            		                oRec5.open strSQLtype, oConn, 3
                            		                if not oRec5.EOF then
                            		                
                            		                aktType = oRec5("fakturerbar")
                            		                
                            		                end if
                            		                oRec5.close
                            		                
                            		                if aktType <> 1 then
                            		                tSttid(y) = ""
                            		                tSltid(y) = ""
                            		                tTimertildelt(y) = ""
                            		                end if  
                            		                
                            		                'end if 't 
			        
			                         'response.Write "y: nyt arr gennemløb: " & y & "<br>"
			                         'Response.flush
			                         '***** ER klokkeslet Sttid og Sltid benyttet eller timeangivelse ****'
			                         '*** tjekker for slet tSttid(y) <> 0 ***'
                                     if cint(visTimerelTid) <> 0 AND len(trim(tSttid(y))) <> 0 then
                                     
                                     if len(trim(tSttid(y))) <> 1 then ' == Slet
                                     
                                     
                                     'Response.Write "<br><b>tSttid(y):</b>" & tSttid(y) & "<br>"
                            			
                            			           
                            		                
                            				    
                                                    '**** Tjekker for tidslås ***'
                                                    '**** Skal timer lægges på en efterfølgende akt? ***' 
                                                    weekdayThis = weekday(datoer(y), 2)
                                                    'Response.write "weekdayThis: " & weekdayThis & "<br>"
                            				        
                                                    select case weekdayThis
                                                    case 1
                                                    flt = "tidslaas_man"
                                                    case 2
                                                    flt = "tidslaas_tir"
                                                    case 3
                                                    flt = "tidslaas_ons"
                                                    case 4
                                                    flt = "tidslaas_tor"
                                                    case 5
                                                    flt = "tidslaas_fre"
                                                    case 6
                                                    flt = "tidslaas_lor"
                                                    case 7
                                                    flt = "tidslaas_son"
                                                    end select 
                            				        
                                                    tidslaas = 0
				                                    tidslaasDayOn = 0
                            				        
                            				        '** Finder grund aktivitet. Skal være = 1 fakturerbar.
                                                    strSQLtidslaas = "SELECT tidslaas, tidslaas_st, "_
                                                    &" tidslaas_sl, "& flt &" AS tidslaasDayOn FROM aktiviteter WHERE "_
                                                    &" id = "& aktids(j) & " AND fakturerbar = 1"
                                                    
                                                    'Response.Write strSQLtidslaas
                                                    'Response.flush
                                                    
                                                    oRec5.open strSQLtidslaas, oConn, 3
                                                    if not oRec5.EOF then
                            				        
                                                    tidslaas = oRec5("tidslaas")
                            				        
                                                    tidslaas_st = formatdatetime(oRec5("tidslaas_st"), 3)
                                                    tidslaas_sl = formatdatetime(oRec5("tidslaas_sl"), 3)
                            				        
                                                    tidslaasDayOn = oRec5("tidslaasDayOn")
                            				        
                                                    end if
                                                    oRec5.close
                                                    
                                                    'Response.Write "<br>A tidslaas " & tidslaas
                                                    'Response.Write "<br>A tidslaas_st " & tidslaas_st
                                                    'Response.Write "<br>A tidslaas_sl" & tidslaas_sl
                                                    'Response.Write "<br>A tidslaasDayOn "& tidslaasDayOn & "<br><br>"
                            				        
				                                    tidslaasErr = 0
				                                    
				                                       
                                                                 
                            				        
                                                    if tidslaas <> 0 then
                                                        
                                                        
                                                        
                                                            if tidslaasDayOn <> 0 then
                                                            
                                                            
                                                                     '*** Forlænger Array strings ***'
        						                                     '** Hvis der er fundet en aktivitet forlænges j arr.
        						                                     '** En indtasting via tidslås går forud for en manuel indtastning direkte i feltet ***'
                                                                    if instr(jarrforl, ",#"& aktids(j) &"#") = 0 then
                                                                    
                                                                    
                                                                    
                                                                    '*** Skal finde op til 8 aktiviteter ***'
                                                                    '** 1+2 Man-Fre Dag/Man-Tor Nat aktiviteter på medarbejder (type:Dag+Nat)
                                                                    '** 3+4 Dag/Nat aktiviteter på kunde (type:E1) 
                                                                    '** 5+6 Mandag morgen akt. aktivitet på kunde (type:Nat+E1)
                                                                    '** 7+8 Lørdag morgen akt. aktivitet på kunde (type:Nat+E1)
                                                                    '** 9+10 Fredag aften + Søndag aften (type:Nat+E1)
                                                                    '** 11+12 Weekend Dag medarb + kunde (type:Weekend+E1)
                                                                    '** 13+14 Weekend Nat medarb + kunde (type:Weekend Nat+E1)
                                                                    
                                                                   
                                                                    cta = 7
                                                                   
                                                                    
                                                                    newArrSizeAkt = ubound(aktids)+100 
                                                                    Redim preserve jobids(newArrSizeAkt) 
                                                                    Redim preserve aktids(newArrSizeAkt)
                                                                    
                                                                    
                                                                    
                                                                    jarrforl = jarrforl & ",#"& aktids(j) & "#"
                                                                    
                                                                    '** Gør plads til 14 akt. (se ovenfor) ***'
                                                                    'st_high_fetnr = high_fetnr
                                                                    for aa = 1 to 100
                                                                    
                                                                    jobids(newArrSizeAkt-100+aa) = jobids(j)
                                                                    aktids(newArrSizeAkt-100+aa) = 0 'aktids(j) '=0?
                                                                    
                                                                    Redim preserve tTimertildelt(ubound(tTimertildelt)+cta)
                                                                    Redim preserve datoer(ubound(tTimertildelt)+cta)
                                                                    Redim preserve tSttid(ubound(tTimertildelt)+cta)
                                                                    Redim preserve tSltid(ubound(tTimertildelt)+cta)
                                                                    Redim preserve feltnr(ubound(tTimertildelt)+cta)
                                                                    
                                                                            for y2 = (ubound(tTimertildelt) - 6) to ubound(tTimertildelt)
                                                                                
                                                                                if y2 = (ubound(tTimertildelt) - 6) then
                                                                                ic = 1
                                                                                else
                                                                                ic = ic + 1
                                                                                end if
                                                                            
                                                                            tTimertildelt(y2) = -9001
                                                                            datoer(y2) = datoer(y)
                                                                            tSttid(y2) = ""
                                                                            tSltid(y2) = ""
                                                                            feltnr(y2) = high_fetnr+ic+3 
                                                                            '** 3 for at få 7 til at blive næste 10'er
                                                                            '** dvs 17 bliver til 20 for at få første felt i næste række
                                                                            
                                                                            'Response.Write "high_fetnr+y2:" & high_fetnr &" - "& y2&"<br>"
                                                                            'Response.Write "feltnr(y2):" &  feltnr(y2) & " y2:"& y2 &"<br>"
                                                                            '*** ialtcounter **'
                                                                            'iic = iic + 1
                                                                            next
                                                                            
                                                                            high_fetnr = high_fetnr + 10
                                                                    
                                                                    next
                                                                    
                                                                end if 'jarrforl
                                                                
                                                                
                                                                
                                                                
                                                                '*** Kig efter ny akt. der er åben iflg. tidslås **'
                                                                '*** Kun hvis der ikke eer angivet NUL for slet **'
                                                                
                                                                    
                                                                    strGrundAktnavn = ""
                                                                    strSQLgrundAkt = "SELECT navn FROM aktiviteter WHERE id = "& aktids(j) 
                                                                    oRec5.open strSQLgrundAkt, oConn ,3
                                                                    if not oRec5.EOF then
                                                                    
                                                                    strGrundAktnavn = oRec5("navn")
                                                                    
                                                                    end if
                                                                    oRec5.close
                                                                    
                                                                    'select case weekdayThis
                                                                    'case 5 'fre
                                                                    'eksFlt = " OR (tidslaas_lor = 1 AND fakturerbar = 53)" 'weekend nat
                                                                    'case else
                                                                    eksFlt = ""
                                                                    'end select
                                                                    
                                                                    eksFlt = ""
                                                                    
                                                                    aktfundet = 0
                                                                    strSQLfind = "SELECT a.id AS id, a.navn, a.job, tidslaas, tidslaas_st, "_
                                                                    &" tidslaas_sl, "& flt &" AS tidslaasDayOn, a.fakturerbar, "_
                                                                    &" tidslaas_man, tidslaas_tir, tidslaas_ons, tidslaas_tor, tidslaas_fre, tidslaas_lor, tidslaas_son"_
                                                                    &" FROM job j "_
                                                                    &" LEFT JOIN aktiviteter a ON (("_
                                                                    &" a.id <> "& aktids(j) & " AND tidslaas = 1 AND ("& flt &" = 1 "& eksFlt &")) AND "_
                                                                    &" (a.fakturerbar = 91 OR a.fakturerbar = 90 OR a.fakturerbar = 50 OR a.fakturerbar = 51 OR a.fakturerbar = 52 OR a.fakturerbar = 53)"_
                                                                    &" AND a.aktstatus <> 0 AND a.job = j.id AND navn LIKE '"& left(strGrundAktnavn, 6) &"%')"_
                                                                    &" WHERE j.id = "& jobids(j) &" AND a.job <> ''"
                                                                    
                                                                    
                                                                   
                                                                    
                                                                    'Response.Write  "lasty2: "& lasty2 & "<br>"
                                                                    'Response.Write "<b>"& strSQLfind & "</b><br>"
                                                                    'Response.flush
                            				                        
                            				                        gnlob = 1
                            				                        'icc = 0
                                                                    oRec5.open strSQLfind, oConn, 3
                                                                    while not oRec5.EOF 
                                                                    
                                                                    
                                                                    
                                                                    'Response.Write "<br><u>Undersøger: "& aktids(j) &" Finder:"& oRec5("id") & "</u><br><br>"
                                    						        'aktidprfelt(y) = oRec5("id")
                                    						        
                                    						                                 
                                    						        
                                    						        
                                    						        B_type = oRec5("fakturerbar")
                                    						        
                                                                    tidslaas_B = oRec5("tidslaas")
                                    						        
                                                                    tidslaas_st_B = formatdatetime(oRec5("tidslaas_st"), 3)
                                                                    tidslaas_sl_B = formatdatetime(oRec5("tidslaas_sl"), 3)
                                    						        
                                                                    tidslaasDayOn_B = oRec5("tidslaasDayOn")
                                                                    
                                                                    tidslaasManOn_B = oRec5("tidslaas_man")
                                                                    tidslaasTirOn_B = oRec5("tidslaas_tir")
                                                                    tidslaasOnsOn_B = oRec5("tidslaas_ons")
                                                                    tidslaasTorOn_B = oRec5("tidslaas_tor")
                                                                    tidslaasFreOn_B = oRec5("tidslaas_fre")
                                                                    tidslaasLorOn_B = oRec5("tidslaas_lor")
                                                                    tidslaasSonOn_B = oRec5("tidslaas_son")
                                                                    
                                    						        
        						                                    strAktNavn_B = replace(oRec5("navn"), "'", "")
                                    						        
        						                                    
        						                                    '*** Redigere tidspunkter efter tidslås ***'
        						                                    
        						                                    '* Hvis tidslås A = 1
        						                                    '* Hvis tidslås A for dagen = 1
        						                                    
        						                                    
        						                                   
        						                                    '*** Trimmer grundreg **'
        						                                    if cint(gnlob) = 1 then
        						                                    
        						                                    '*** Tjekker Dato format her ***'
        						                                    
        						                                    'KODE
        						                                    
        						                                    '****'
        						                                    gRegSt = tSttid(y)
        						                                    gRegSl = tSltid(y)
        						                                    
        						                                    
        						                                    
        						                                            
        						                                           
        						                                    end if 'gnlob   
        						                                    
        						                                    
        						                                    
        						                                    if request("minindtast_on") = "1" then
        						                                    '**** Min 7,4 timer på medarb / 8 timer på kunde ****'
					                                                '**** Det skal være samlet for dagen             ****'
					                                                '**** hvilken akt. skal den placere det på?      ****'
					                                                '**** Denne beregning skal ligge i INC filen     ****'
    					                                            
    					                                            idag = formatdatetime(now, 2)
                                                                    minutberegnDiffmin = datediff("n", idag &" "& gRegSt, idag &" "& gRegSl, 2, 2) 
                                						                                           
    					                                            
					                                                select case B_type
					                                                case "50", "51", "52", "53", "54", "55"
					                                                    if minutberegnDiffmin < 444 AND minutberegnDiffmin > 0 then
					                                                    tillaeg = (444 - minutberegnDiffmin) 
					                                                    gRegSl = dateAdd("n", tillaeg, idag &" "& gRegSl)
					                                                    gRegSl = left(formatdatetime(gRegSl, 3), 5)
					                                                    
					                                                    end if
    					                                                
    					                                               
					                                                    if minutberegnDiffmin < 0 then 'AND minutberegnDiffmin > -444 then
					                                                    
					                                                        imorgen = formatdatetime(dateadd("d", 1, idag), 2)
					                                                        minutberegnDiffmin = datediff("n", idag &" "& gRegSt, imorgen &" "& gRegSl, 2, 2) 
                                    						            
					                                                        tillaeg = (444 - (minutberegnDiffmin)) 
					                                                        gRegSl = dateAdd("n", tillaeg, idag &" "& gRegSl)
					                                                        gRegSl = left(formatdatetime(gRegSl, 3), 5)
    					                                                    
					                                                    end if
    					                                            
					                                                case "90", "91"
    					                                                
    					                                                
					                                                    if minutberegnDiffmin < 480 AND minutberegnDiffmin > 0 then
					                                                        tillaeg = (480 - minutberegnDiffmin) 
					                                                        gRegSl = dateAdd("n", tillaeg, idag &" "& gRegSl)
					                                                        gRegSl = left(formatdatetime(gRegSl, 3), 5)
					                                                    end if
    					                                                
    					                                                
					                                                    if minutberegnDiffmin < 0 then 'AND minutberegnDiffmin > -480 then
    					                                                    
					                                                        imorgen = formatdatetime(dateadd("d", 1, idag), 2)
					                                                        minutberegnDiffmin = datediff("n", idag &" "& gRegSt, imorgen &" "& gRegSl, 2, 2) 
                                    						            
					                                                        tillaeg = (480 - (minutberegnDiffmin)) 
					                                                        gRegSl = dateAdd("n", tillaeg, idag &" "& gRegSl)
					                                                        gRegSl = left(formatdatetime(gRegSl, 3), 5)
    					                                                    
    					                                                    
    					                                                    'Response.Write "tillaeg"
    					                                                    'Response.end
					                                                    end if
    					                                            
					                                                'case else
					                                                'minutberegnDiffmin = minutberegnDiffmin
					                                                end select
        						                                    
        						                                    
        						                                    end if '*** force min indtast 7,4 / 8,0
        						                                    
        						                                   
        						                                   
        						                                     
        						                                    '*** Er tidslås slåettil på dagen ***'
        						                                     if cint(tidslaas_B) = 1 AND cint(tidslaasDayOn_B) = 1 then
        						                                    
        						                                     aktfundet = 1
        						                                   
        						                                          
                                                                    '*** Finder de korrekte tidspunkter ***'
                                                                    
                                                                    'if cint(gnlob) = 1 then
                                                                    'if y = ystart then 
                                                                    if lasty2 = 0 then
                                                                    gnlob_st_high_fetnr = UBOUND(feltnr) - (707) '(105) '7*14 = 98 (+ 7)
                                                                    else
                                                                    gnlob_st_high_fetnr = gnlob_st_high_fetnr + 7 '** 7 dage
                                                                    end if
                                                                    
                                                                    'icc = 0
                                                                                    
                                                                                    for y2 = gnlob_st_high_fetnr + 1 to gnlob_st_high_fetnr + 7 '*** Lægges 7 til for 7 dage ***'
                                                                                    
                                                                                           
                                                                                            
                                                                                            aktids(newArrSizeAkt-100+icc+1) = oRec5("id") '+gnlob
                                                                                            
                                                                                            if right(feltnr(y2), 1) = right(feltnr(y),1) then
                                                                                            
                                                                                             
                                                                                                'y2 = y2 + (cint(right(feltnr(y),1)) * 7)
                                                                                            
                                                                                            
                                                                                                     '*** Trimmer de funde aktiviteter **'
        						                                                                     'Response.Write "<b>Aktid:</b>"& oRec5("id") &"/"& aktids(newArrSizeAkt-7+gnlob) &"y2 val: " & y2 &"/"& newArrSizeAkt-7+gnlob &" gRegSt: "& gRegSt & " til "& gRegSl &" - tidslaas_st_B: " & tidslaas_st_B & " tidslaas_sl_B: " & tidslaas_sl_B & "cint(right(feltnr(y2), 1)) >= cint(weekdayThis)"& cint(right(feltnr(y2), 1)) &" >= "& cint(weekdayThis) &"<br>"
                                        						                                     
                						                                                            '*** Er sttid større end start tidslås og er sttid mindre end tidslås slut.
        						                                                                    if (cDate(gRegSt&":00") >= cDate(tidslaas_st_B) AND cDate(tidslaas_st_B) >= cDate(tidslaas_sl_B)) _ 
        						                                                                    OR (cDate(gRegSt&":00") >= cDate(tidslaas_st_B) AND cDate(gRegSt&":00") <= cDate(tidslaas_sl_B))then 
        						                                                                        'if cDate(gRegSt&":00") > cDate(tidslaas_sl_B) then
                                                                                                        'tSttid(y2) = ""
                                                                                                        'else
                                                                                                        tSttid(y2) = gRegSt
                                                                                                        'end if
                                                                                                        'Response.Write "Bruger st.tid: " & gRegSt 
                                                                                                    else 
                                                                                                    tSttid(y2) = left(tidslaas_st_B, 5)
        						                                                                        'Response.Write "Bruger tidslås start: " & left(tidslaas_st_B, 5)
        						                                                                    end if
                                						                                            
                                						                                            
        						                                                                    '* Hvis angivet SLUT tidspunkt er MINDRE end tidslås slut A
        						                                                                    '>>
        						                                                                    '** Er slut tid næste morgen, dvs efter 24:00 **'
        						                                                                    '_
        						                                                                    'OR (cDate(gRegSl&":00") <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) <= cDate(gRegSt&":00"))
        						                                                                    
        						                                                                    if (cDate(gRegSl&":00") <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) <= cDate(tidslaas_sl_B) AND cDate(gRegSl&":00") > cDate(tidslaas_st_B)) _ 
        						                                                                    OR (cDate(gRegSl&":00") > cDate(tidslaas_st_B) AND cDate(gRegSl&":00") > cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) >= cDate(tidslaas_sl_B)) _
        						                                                                    OR (cDate(gRegSl&":00") <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) >= cDate(tidslaas_sl_B) AND cDate(gRegSt&":00") >= cDate(gRegSl&":00")) _
        						                                                                    OR (cDate(gRegSl&":00") <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) <= cDate(tidslaas_sl_B) AND cDate(tidslaas_st_B) <= cDate(gRegSt&":00") AND cDate(tidslaas_sl_B) <= cDate(gRegSt&":00")) then
        						                                                                    '===   TID   ======    DAG      =====   NAT   ========
        						                                                                    '1
        						                                                                    '2
        						                                                                    '3 21:00 - 07:00 => 06:00 - 07:00 => 21:00 - 06:00
        						                                                                    '5 10:45 - 04:30 => 10:45 - 18:00 => 18:00 - 04:30    
        						                                                                    '=====================================================  
        						                                                                    tSltid(y2) = gRegSl
        						                                                                    else
        						                                                                    tSltid(y2) = left(tidslaas_sl_B, 5)
        						                                                                    end if
        						                                                                    '* Hvis angivet SLUT tidspunkt er STØRRE end tidslås slut A
        						                                                                    '>>
                                						                                    
        						                                                                    idag = formatdatetime(now, 2)
        						                                                                    minutberegning = datediff("n", idag &" "& tSttid(y2), idag &" "& tSltid(y2), 2, 2) 
                                						                                            
                                						                                            
                                						                                            
                                						                                            
        						                                                                    'Response.Write "minutberegning tidslås st: "& tidslaas_st_B &" tidslås sl: "& tidslaas_sl_B &" gregST: "& cDate(gRegSt) &" gRegSl: "& cDate(gRegSl) &" yST: "& cDate(tSttid(y2)) &" ySL "&tSltid(y2)&" typ: "& B_type &": " & minutberegning & "<br>"
        						                                                                    'Response.flush
                                						                                            
        						                                                                    '*** Tjekker om der skal indlæses på akt selvom minutbe. er negativt **'
        						                                                                    '*** AND cDate(tSltid(y2)) < cDate(gRegSl)
        						                                                                    if (minutberegning < 0 AND cDate(tSttid(y2)) > cDate(gRegSt) AND cDate(gRegSt) < cDate(gRegSl))_
        						                                                                    OR (minutberegning < 0 AND cDate(tidslaas_sl_B) < cDate(tidslaas_st_B) AND cDate(tidslaas_sl_B) <= cDate(gRegSl) AND cDate(tidslaas_st_B) <= cDate(tSltid(y2)))_
        						                                                                    OR (minutberegning < 0 AND cDate(tidslaas_st_B) < cDate(tidslaas_sl_B) AND cDate(gRegSt) >= cDate(tidslaas_sl_B) AND cDate(gRegSl) <= cDate(tidslaas_st_B))_
        						                                                                    
        						                                                                    then
        						                                                                    
        						                                                                    'Response.Write "minutbereg. negativt, indtastning IKKE godkendt<br><br>"
        						                                                                    
        						                                                                    
        						                                                                    tSttid(y2) = ""
        						                                                                    tSltid(y2) = ""
        						                                                                    end if
                                                                                                    
                                                                                                    
                                                                                                    '** Korrigere dato (lør morgen / Mandag morgen mv.)
                                                                                                    datoer(y2) = datoer(y)
                                                                                                    'if cint(weekdayThis) = 5 AND cint(tidslaasFreOn_B) = 0 AND cint(tidslaasLorOn_B) = 1 then 
                                                                                                    'datoer(y2) = dateadd("d", 1, datoer(y))
                                                                                                    'feltnr(y2) = feltnr(y2) + 1
                                                                                                    'end if
                                                                                                    'feltnr(y2) = feltnr(y)
                                                                                                    tTimertildelt(y2) = ""
                                                                                                    
                                                                                                    'Response.Write "<b>Resultat af fundet akt:</b><br>"
                                                                                                    'Response.Write "FELTNR: "& feltnr(y2) &" y2 val2 /y: "& y2 &"/ "& y &" # "& tSttid(y2) &" - "& tSltid(y2) & " Dato" & datoer(y2) & "dserial: <br><br>"
                                                                                                    'lastUsedy2serie = y2
                                                                                                    
                                                                                                    
                                                                                                    'tSttid(y) = tSttid(y2)
        						                                                                    'tSltid(y) = tSttid(y2)
        						                                                                    'tTimertildelt(y2) = ""
                                                                                                    
                                                                                            else
                                                                                            
                                                                                            'Response.Write "Nulstiller felt<br>"
                                                                                            
                                                                                            'tSttid(y2) = ""
        						                                                            'tSltid(y2) = ""
        						                                                            'tTimertildelt(y2) = -9001
                                                                                            
                                                                                            end if
                                                                                            
                                                                                          
                                                                                    lasty2 = y2
                                                                                    next
                                                                                    
                                                                   
                                                                    end if 'tidslås slået til
                                                                    
                                                                    
                                                                    icc = icc + 1  
        						                                    gnlob = gnlob + 1
                                                                    oRec5.movenext
                                                                    wend 
                                                                    oRec5.close
                                                                    
                                                 
                                                                    
                                                                    
                                                                   
                                                                    
                                                            '*** err hvis der ikke findes en akt med en dækkende tidslås ***'
                                                            if cint(aktfundet) = 0 then
                                                             tidslaasErr = 1
                                                            end if
                                                                        
                                                                        
                                                                        
                                                                
                                                        
                                                else
                                                
                                                tidslaasErr = 1
                                                
                                                end if 'tidslaasDayOn
                                                
                    				        
                                            end if 'tidslåas
                                                    
                                                    
                                                    
                                                    
                                                    'if tidslaasErr <> 0 then
                                                    if tidslaasErr = 1110 then
                                                    %>
                                                    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
                                                    <% 
                                    			    
                                                    useleftdiv = "t"
                                                    errortype = 136
                                                    call showError(errortype)
                                    		        
                                                    Response.end
                                                    end if
                                                    
                                                    
                            			        
                            			
                                        end if 'slet
                            		
                                    else
                            		
                                    '** Hvis alm. timereg er benyttet, 
                                    '** dvs der ikke er angivet klokkeslet for start og slut ***'
                                    sTtid = ""
                                    sLtid = ""
                            		
                                    end if
                                   
                                   
                                   
                                   
                                   next 'y
                                   next 'j 
                                    
				                    
				                    
				                    
				                            
				                
				                'Response.Write "aktids(newArrSizeAkt)<br>" 
				                
				                for j = 0 to UBOUND(aktids) 
				                'Response.Write aktids(j) & " == "
				                
				                     if j = 0 then
			                        ystart = 0
			                        else
			                        ystart = (j * 7) 
			                        end if
                        			
			                        yslut = (ystart + 6)
			                     
			                     'oprArrSize = ubound(tTimertildelt)
			                     for y = ystart to yslut    
			                        
			                        'Response.Write  feltnr(y) & ":("& y &"):" & tSttid(y) &" - "& tSltid(y) & " | "
			                        
			                        
			                     next
			                     
			                     
			                     'Response.Write "<hr>"
				                
				                next
				                
				                'Response.end        
				            
				                '**** Slut ETS ****'
	                             end if '*** ETS track kode slut
				               
	                
	
			
			
			'*** Validerer ***
			antalErr = 0
			for y = 0 to UBOUND(tTimertildelt)
			
				
				call erDetInt(SQLBless(trim(tTimertildelt(y))))
				if isInt > 0 then
					antalErr = 1
					errortype = 28
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0
				
			
			next
			
			
			
			idagErrTjek = day(now)&"/"&month(now)&"/"&year(now)
			
			for y = 0 to UBOUND(tSttid)	
				
				'Response.write SQLBless3(trim(tSttid(y))) & "<br>"
				'Response.flush
				
				tSttid(y) = replace(tSttid(y), ".", ":")
				tSttid(y) = replace(tSttid(y), ",", ":")
				
				if instr(tSttid(y), ":") = 0 AND len(trim(tSttid(y))) = 4 then
				  
				  left_tid = left(tSttid(y), 2)
				  right_tid = right(tSttid(y), 2)
				  nyTid = left_tid & ":"& right_tid
				  tSttid(y) = nyTid
				
				
				end if 
				  
				
				
				'Response.write "len(trim(tSttid(y))" & len(trim(tSttid(y))) & "<br>"
				if len(trim(tSttid(y))) <> 1 then ' == Slet
				
				
				
				call erDetInt(SQLBless3(trim(tSttid(y))))
				if isInt > 0 then
					antalErr = 1
					errortype = 63
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0
				
				'*** Punktum i angivelse ved registrering af klokkeslet
				'if instr(trim(tSttid(y)), ".") <> 0 OR instr(trim(tSttid(y)), ",") <> 0 then
				'	antalErr = 1
			    '	errortype = 66
				'	useleftdiv = "t"
				'	<include file="../inc/regular/header_lysblaa_inc.asp"--><
				'	call showError(errortype)
				'	response.end
				'end if	
				
				if len(trim(tSttid(y))) <> 0 then
				
				'Response.write idagErrTjek &" "& tSttid(y)&":00" &" - "& isdate(idagErrTjek &" "& tSttid(y)&":00") &"<br>"
					if isdate(idagErrTjek &" "& tSttid(y)&":00") = false OR instr(idagErrTjek &" "& tSttid(y),":") = 0 then
						errTid = tSttid(y)
						antalErr = 1
						errortype = 64
						useleftdiv = "t"
						%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						call showError(errortype)
						response.end
					end if
				end if
				
				
				end if ' tSltid(y) <> 0 
				
			next
			
			
			
			
			for y = 0 to UBOUND(tSltid)	
				
				tSltid(y) = replace(tSltid(y), ".", ":")
				tSltid(y) = replace(tSltid(y), ",", ":")
				
				if instr(tSltid(y), ":") = 0 AND len(trim(tSltid(y))) = 4 then
				  
				  left_tid = left(tSltid(y), 2)
				  right_tid = right(tSltid(y), 2)
				  nyTid = left_tid & ":"& right_tid
				  tSltid(y) = nyTid
				
				end if 
				
				
				call erDetInt(SQLBless3(trim(tSltid(y))))
				if isInt > 0 then
					antalErr = 1
					errortype = 63
					useleftdiv = "t"
					%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
					call showError(errortype)
					response.end
				end if	
				isInt = 0
				
				
				'*** Punktum  i angivelse ved registrering af klokkeslet
				'if instr(trim(tSltid(y)), ".") <> 0 OR instr(trim(tSltid(y)), ",") <> 0 then
				'	antalErr = 1
				'	errortype = 66
				'	useleftdiv = "t"
				'	include file="../inc/regular/header_lysblaa_inc.asp"--
				'	call showError(errortype)
				'	response.end
				'end if	
				
				
				if len(trim(tSltid(y))) <> 0 then
				'Response.Write "len(trim(tSltid(y))) " & len(trim(tSltid(y))) & "<br>"
					if isdate(idagErrTjek &" "& tSltid(y)&":00") = false OR instr(idagErrTjek &" "& tSltid(y),":") = 0  then
						errTid = tSltid(y)
						antalErr = 1
						errortype = 64
						useleftdiv = "t"
						%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						call showError(errortype)
						response.end
					end if
				end if
				
			next
			
			
			
	
	else
	
	
	'**** Hvis indlæsnng sker fra Stopur  ***'
	
	
	
    antalErr = 0
	jobidSQLkri =  ""
	antalsids = split(request("sids"), ", ")
	for t = 1 to UBOUND(antalsids)
	
	if t = 1 then
	jobidSQLkri = jobidSQLkri & " w.id = "& antalsids(t)
	else
	jobidSQLkri = jobidSQLkri & " OR w.id = "& antalsids(t)
	end if  
	
	next
	 
	medid = request("medid")
	medarbejderid = medid
	sjobids = ""
	saktids = ""
	timerthis = ""
	datoerthis = ""
	sids = ""
	
	'*** Henter jobids og aktids ***'
	strSQLstop = "SELECT w.id AS wid, w.jobid, w.aktid, w.sttid, w.sltid, w.kommentar, w.incident, "_
	&" s.dato, s.tidspunkt FROM stopur w "_
	&" LEFT JOIN sdsk s ON (s.id = w.incident) WHERE ("& jobidSQLkri &") AND timereg_overfort = 0"
	
	'Response.write strSQLstop &"<br>"
	'Response.flush 
	
	oRec2.open strSQLstop, oConn, 3
	while not oRec2.EOF
	
	sids = sids & oRec2("wid") &","
	sjobids = sjobids & oRec2("jobid") &","
	saktids = saktids & oRec2("aktid") &","
	
	    if isDate(oRec2("sttid")) AND isDate(oRec2("sltid")) then 
	    tidThis = datediff("n", oRec2("sttid"),oRec2("sltid"), 2,2)
	    else
	    tidThis = 0 
	    end if
	    
	    
	    call timerogminutberegning(tidThis)
	    timerthis = timerthis & thoursTot&"."& tminProcent &","
	    
	   
	datoerthis = datoerthis & oRec2("sttid") &","   
	datoogklokkeslet = left(weekdayname(weekday(oRec2("dato"))), 3) &". "&formatdatetime(oRec2("dato"), 2) &" "& left(formatdatetime(oRec2("tidspunkt"), 3), 5)
    kommentars = kommentars & "<br><br> == Incident Id "& oRec2("incident") &" == <br>Opr. "& datoogklokkeslet &"<br>Udført ("& oRec2("wid") &") "& oRec2("sttid") &" - "& oRec2("sltid") &" <br> "& replace(oRec2("kommentar"), "'", "&#39;") &"#,#"
	
	oRec2.movenext
	Wend
	oRec2.close
	
	if len(kommentars) > 3 then
	kommentars_len = len(kommentars)
	kommentars_left = left(kommentars, kommentars_len - 3) 
	kommentars = kommentars_left
	end if
	
	
	if len(sjobids) > 1 then
	sjobids_len = len(sjobids)
	sjobids_left = left(sjobids, sjobids_len - 1) 
	sjobids = sjobids_left
	end if
	
	if len(saktids) > 1 then
	saktids_len = len(saktids)
	saktids_left = left(saktids, saktids_len - 1) 
	saktids = saktids_left
	end if
	
	if len(timerthis) > 1 then
	timerthis_len = len(timerthis)
	timerthis_left = left(timerthis, timerthis_len - 1) 
	timerthis = timerthis_left
	end if
	
	if len(datoerthis) > 1 then
	datoerthis_len = len(datoerthis)
	datoerthis_left = left(datoerthis, datoerthis_len - 1) 
	datoerthis = datoerthis_left
	end if
	
	if len(sids) > 1 then
	sids_len = len(sids)
	sids_left = left(sids, sids_len - 1) 
	sids = sids_left
	end if
	
	kommentarer = split(kommentars, "#,#")
	entrysids = split(sids, ",")
	jobids = split(sjobids, ",")
	aktids = split(saktids, ",")
	useTimer = split(timerthis, ",")
	useDatoer = split(datoerthis, ",")
	
	
	end if '***Fra  Stopur****'
	
	
	
	
	
	
	
	
	
	'*******************************'
	'***** Indlæser timer array ****'
	'*** indlæser i db ***
	
	
	'Response.Write "<br><br>antalErr" & antalErr & "<br>"
	
	
	if antalErr = 0 then
	
	
	
	
	
	
	for m = 0 to UBOUND(tildelselmedarb)
	
	'Response.Write tildelselmedarb(m) & "<br>"
	
	if cint(multitildel) = 1 then
	medarbejderid = tildelselmedarb(m) 
	else
	medarbejderid = medarbejderid
	end if
	
	    
        
	
	    '**** Finder medarbejder timepris ************
		call mNavnogKostpris(medarbejderid)
	    
	    for j = 0 to UBOUND(jobids)
			    
			    '*****************************************************'
				'*** Henter Akt og Job oplysninger                  **'
				'*** Finder evt. alternativ medarbejder timepris    **'
				'*** på den valgte aktivitet                        **'
				'*****************************************************'
				
                 'foundone_0 = 0
                 'intTimepris_0 = 0
                 'intValuta_0 = 0
                
				
				strSQL = "SELECT job.id AS jobid, aktiviteter.navn, fakturerbar, jobnavn, jobnr, "_
				&" kkundenavn, jobknr, fastpris, job.serviceaft"_
				&" FROM aktiviteter, job, kunder "_
				&" WHERE aktiviteter.id = "& aktids(j) &" AND job.id = "& jobids(j) &" AND kunder.Kid = jobknr "
				
				
				
				'Response.write "<br><br>"& strSQL &"<br><br>"
				'Response.flush				
				
				oRec.Open strSQL, oConn, 3
				if Not oRec.EOF then
					
							 intjobid = oRec("jobid")
							 strJobnavn = oRec("Jobnavn")
					 		 strJobknavn = oRec("kkundenavn")
							 strJobknr = oRec("Jobknr") 	
							 strAktNavn = oRec("navn")
							 strFakturerbart = oRec("fakturerbar")
							 intServiceAft = oRec("serviceaft")
							 intJobnr = oRec("jobnr")
							 
							 '**** 2009 t.tfaktim er = a.fakturerbar (akt. type) også i timer tabellen. **'
							 tfaktimvalue = strFakturerbart
							 strFastpris = oRec("fastpris")
						
					        
					        '*** tjekker for alternativ timepris på aktivitet
							call alttimepris(aktids(j), jobids(j), medarbejderid, 1)
							
							'** Er der alternativ timepris på jobbet
							if foundone = "n" then
								'if cint(foundone_0) <> 1 then
                                call alttimepris(0, jobids(j), medarbejderid, 1)
                                'intTimepris_0 = intTimepris
                                'intValuta_0 = intValuta
                                'foundone_0 = 1
                                'else
                                'intTimepris = intTimepris_0 
                                'intValuta = intValuta_0
                                'end if
							end if
							
							'**************************************************************'
							'*** Hvis timepris ikke findes på job bruges Gen. timepris fra '
							'*** Fra medarbejdertype, og den oprettes på job **************'
							'**************************************************************'
							if foundone = "n" then
							intTimepris = replace(tprisGen, ",", ".") 
							intValuta = valutaGen
							
							strSQLtpris = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
							&" VALUES ("& intjobid &", 0, "& medarbejderid &", 0, "& intTimepris &", "& intValuta &")"
							
							oConn.execute(strSQLtpris)
							
							end if
							
				end if
				oRec.Close
	
	        
	        
	        
	        
	        if stopur <> "1" then 
		    '****************************************************'
			'***** Dag 1 - 7 Indlæses i db fra Timereg **********'
			'****************************************************'
			
			if j = 0 then
			ystart = 0
			else
			ystart = (j * 7) 
			end if
			
			yslut = (ystart + 6)
			
			isAktMedidWrt = ""
			multiDagArrUse = 0
			
			'for y = 0 to UBOUND(tTimertildelt)
			for y = ystart to yslut   
			    
			    '*** tjekker om felt id liggen i den rigtige aktlinie *'
			    '**  DVS mellem 11 - 17, eller 21-27 etc..          ***'
				if y >= ystart AND y <= yslut then
				
				feltvalThis = cint(replace(feltnr(y), "_", ""))
				feltvalThis = cstr(feltvalThis)
				
				
				strKomm = replace(request("FM_kom_"&feltvalThis), "'", "&#39;")
				
			    '*** Tjekker at komm. er udfyldt ved pre-def aktiviteter ***'
				call akttyper2009prop(tfaktimvalue)
				if len(trim(tTimertildelt(y))) <> 0 AND aty_pre <> 0 AND tTimertildelt(y) > -9000 then
				        
				        if cdbl(replace(tTimertildelt(y), ".", ",")) <> cdbl(aty_pre) AND len(trim(strKomm)) = 0  then
				        'Response.Write cdbl(replace(tTimertildelt(y), ".", ",")) &" <> "& cdbl(aty_pre) & "<br>"
				        antalErr = 1
						errortype = 131
						useleftdiv = "t"
						%><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						call showError(errortype)
						response.end
						end if
				
				end if
				
				if len(request("FM_off_"&feltvalThis)) <> 0 then
				offentlig = request("FM_off_"&feltvalThis)
				else
				offentlig = 0
				end if
				
				if len(request("FM_bopal_"&feltvalThis)) <> 0 then
				bopal = request("FM_bopal_"&feltvalThis)
				else
				bopal = 0
				end if
				
				
				if len(request("FM_destination_"&feltvalThis)) <> 0 then
				destination = request("FM_destination_"&feltvalThis)
				else
				destination = ""
				end if
				
				if len(trim(tTimertildelt(y))) <> 0 then
				tTimertildelt(y) = SQLBless(tTimertildelt(y))
				else
					if visTimerelTid <> 0 then
					tTimertildelt(y) = -9002 '-2 Opdater
					else
					tTimertildelt(y) = -9001 '-1 Slet
					end if
				end if
				
				
				    'Response.Write "fltnt:"& feltnr(y) &"aktid: "& aktids(y) &" timer:" & tTimertildelt(y) & " Dato: " & datoer(y) & " tSttid(y): "& tSttid(y) &" medarbejderid: "& medarbejderid&"<br>"
                    'Response.flush
				
			    
				    '**** * Multitildel for flere dage ****'
				    if instr(tTimertildelt(y), "*") <> 0 then
    				    
				        lngt = len(tTimertildelt(y))
				        st = instr(tTimertildelt(y), "*")
				        multiDagArr = mid(tTimertildelt(y), st+1, lngt)
				        'multiDagArr = (multiDagArr/5) * 7
				            's = 0
				            'if s = 100 then
				            if instr(multiDagArr, ",") <> 0 OR instr(multiDagArr, ".") <> 0 then
				            antalErr = 1
						    errortype = 135
						    useleftdiv = "t"
						    %><!--#include file="../inc/regular/header_lysblaa_inc.asp"--><%
						    call showError(errortype)
						    response.end
						    end if
    				    
    				    
				        findTimerForStar = left(tTimertildelt(y), (st-1))
				        multiDagArrUse = multiDagArr - 1
    				    
				        if left(trim(findTimerForStar), 1) = "." then
			            useTimer = "0"&trim(findTimerForStar)
			            else
			            useTimer = findTimerForStar
			            end if    
    				    
    				    
				    '*** tildeler timer på valgte dage, ved multiDage funktion tildeles ikke på hellgidage ***'
				    '*** Hvad med personlige fridage, der læses ike timer ind på dage ved indtasting i periode *''
				    '**  via * hvis medarbejderen ikke 
				    '*** ikke har normeret tid de pågældende dage **'
				    muDag = 0
				    for muDag = 0 to multiDagArrUse 
				    'Response.Write  "muDag " & muDag & "<br>"
    				'Response.write AktMedidWrt  & "<br>"
    				
				    useDato = dateAdd("d", muDag, datoer(y))
				    '**** Indlæser kun timer på dage hvor den valgte medarbejder har arbejdsdag, og ikke på helligdage **'
    				
				    ntimPer = 0
				    call normtimerPer(medarbejderid, useDato, 0)
    				
				        'Response.Write "timer:" & useTimer & " Dato: " & useDato & " ntimPer:"& ntimPer &" medarbejderid: "& medarbejderid&"<br>"
                        'Response.flush
    				
				        if cint(ntimPer) <> 0 then
    				    useTimerThis = useTimer
    				    else
    				    useTimerThis = 0
    				    end if
    				   
    				      
            				
				            '**** Indlæser timer i DB, alm timereg. i timer *****'
				            call opdaterimer(aktids(j), strAktNavn, tfaktimvalue, strFastpris, intJobnr, strJobnavn, strJobknr,_
				            strJobknavn, medarbejderid, strMnavn,_
				            useDato, useTimerThis, strKomm, intTimepris,_
				            dblkostpris, offentlig, intServiceAft, strYear,_
				            tSttid(y), tSltid(y), visTimerelTid, stopur, intValuta, bopal, destination)
            				
    				    
    				    
				        isAktMedidWrt = isAktMedidWrt & ",#"&aktids(j)&"_"&medarbejderid&"#"
				        'end if
    				    
				        findTimerForStar = ""
				        
    				 
				    next 
				    
    				
    				
    				
				    else
				    '*** else flere dage * ==> dvs. alm
    				        
    				        
				            if left(trim(tTimertildelt(y)), 1) = "." then
			                useTimer = "0"&trim(tTimertildelt(y))
			                else
			                useTimer = tTimertildelt(y)
			                end if    
    			            
    				        muDag = 0
				            useDato = datoer(y)
				            
				                 
				                 
				             '*** ETS HER ? ***  
				                 
				            
				            
				            'Response.write AktMedidWrt  & "<br>"
    				        
				            '*** Er akt/medarb. allerede opdateret via multiDatoer?****' 
				            '** Så skal de ikke opdateresd igen, da de så overskriver den allerede opdaterede værdi ***'
				            if instr(isAktMedidWrt, ",#"&aktids(j)&"_"&medarbejderid&"#") = 0 then
				            
				           
    	                    '**** Indlæser timer i DB, alm timereg. i timer *****'
				            
				            if aktids(j) <> 0 AND len(trim(aktids(j))) <> 0 then
				            call opdaterimer(aktids(j), strAktNavn, tfaktimvalue, strFastpris, intJobnr, strJobnavn, strJobknr,_
				            strJobknavn, medarbejderid, strMnavn,_
				            useDato, useTimer, strKomm, intTimepris,_
				            dblkostpris, offentlig, intServiceAft, strYear,_
				            tSttid(y), tSltid(y), visTimerelTid, stopur, intValuta, bopal, destination)
    				        
    				        
    				        isAktMedidDatoWrt = isAktMedidDatoWrt & ",#"&aktids(j)&"_"&medarbejderid&"_"&useDato&"#"
				            end if
				            
				            
				            end if
    				
    				
				    end if
			   
				        
				end if 'y <= ystart
			
			next 'y
			
			else ' = fra Stopur
			    
			        
			        tSttid = ""
	                tSltid = "" 
                	visTimerelTid = 0 '** Vis timer
                	offentlig = 0
                	strYear = year(useDatoer(j))
                    
                    
                    destination = ""
	                bopal = 0
               
               
                '****************************************************************'
                '*** Er uge alfsuttet af medarb, er smiley og autogk slået til **'
                '****************************************************************'
                regdato = useDatoer(j)
                firstWeekDay = weekday(regdato, 2)
                
                select case firstWeekDay
                case 1
                stTil = 0
                slTil = 6
                case 2
                stTil = 1
                slTil = 5
                case 3
                stTil = 2
                slTil = 4
                case 4
                stTil = 3
                slTil = 3
                case 5
                stTil = 4
                slTil = 2
                case 6
                stTil = 5
                slTil = 1
                case 7
                stTil = 6
                slTil = 0
                end select
                
                stwDayDato = dateadd("d", -stTil, regDato)
                slwDayDato = dateadd("d", slTil, regDato)
                
                regdatoStSQL = year(stwDayDato) &"/"& month(stwDayDato) &"/"& day(stwDayDato)
                regdatoSlSQL = year(slwDayDato) &"/"& month(slwDayDato) &"/"& day(slwDayDato)
                
                
                'Response.Write "firstWeekDay "& firstWeekDay
                
                call afsluger(medarbejderid, regdatoStSQL, regdatoSlSQL)
                
                
                erugeafsluttet = instr(afslUgerMedab(medarbejderid), "#"&datepart("ww", regdato,2,2)&"_"& datepart("yyyy", regdato) &"#")
                
                'Response.Write useDatoer(j) & "<br>"
                'Response.Write "erugeafsluttet --" & erugeafsluttet  &"<br>"
                'Response.Write "autogk --" & autogk  &"<br>"
                'Response.Write "smilaktiv  --" & smilaktiv   &"<br>" 
                'Response.flush
                'Response.end
                
                call lonKorsel_lukketPer(regdato)
              
                 if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", regdato) = year(now) AND DatePart("m", regdato) < month(now)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", regdato) < year(now) AND DatePart("m", regdato) = 12)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", regdato) < year(now) AND DatePart("m", regdato) <> 12) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", regdato) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
              
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = 0
                end if 
		        
		                
		                
		        '*** Tjekker sidste fakdato ***'
		        'if len(trim(intjobid)) <> 0 then
		        'intjobid = intjobid
		        'else
		        'intjobid = 0
		        'end if
		        
		        
		        
		        lastFakdato = "01/01/2002"
		        strSQL = "SELECT fakdato FROM fakturaer WHERE jobid = "& intjobid &" AND faktype = 0 ORDER BY fakdato DESC LIMIT 0,1"
		        
		       
		        oRec.open strSQL, oConn, 3
		        if not oRec.EOF then
		        lastFakdato = oRec("fakdato")
		        end if
		        oRec.close
		        
		        if ugeerAfsl_og_autogk_smil = 1 OR cdate(lastFakdato) >= cdate(regdato) then
		        
		        %>
			    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			    <% 
			    
			    if lastFakdato = "01/01/2002" then
			    lastFakdato = "(ingen)"
			    end if
			    
			    useleftdiv = "c2"
			    errortype = 117
			    call showError(errortype)
		        
		        Response.end
		        end if
               
                strKomm = kommentarer(j)
			    
			    
			    'Response.Write "tSttid "& tSttid &",  tSltid "& tSltid &", visTimerelTid "& visTimerelTid &", stopur "& stopur
			    'Response.flush
			    
			    'aktids(j) ", strAktNavn, tfaktimvalue, strFastpris, intJobnr, strJobnavn, strJobknr,_
				'strJobknavn, medarbejderid, strMnavn,_
				'useDatoer(j), useTimer(j), strKomm, intTimepris,_
				'dblkostpris, offentlig, intServiceAft, strYear,
			    
			    '**** Indlæser timer i DB *****'
				call opdaterimer(aktids(j), strAktNavn, tfaktimvalue, strFastpris, intJobnr, strJobnavn, strJobknr,_
				strJobknavn, medarbejderid, strMnavn,_
				useDatoer(j), useTimer(j), strKomm, intTimepris,_
				dblkostpris, offentlig, intServiceAft, strYear, tSttid, tSltid, visTimerelTid, stopur, intValuta, bopal, destination)
 
			'*** Opdaterer stopur timer overført *****'
			strSQLstopU = "UPDATE stopur SET timereg_overfort = 1 WHERE id = "& entrysids(j) 
	        oConn.execute(strSQLstopU)
			       
			end if '**** Fra Stopur ***'
			
		
		
		next 'j
		
		next 'm (multideldel)
	
	    
	    '** Tjekker indlæs værdier **' 
	    'Response.end
	
	
	
	
	
	            '*************************************'
	            '*** Lukker job og afsender email ****'
	            '*************************************'
	            if len(trim(request("FM_lukjob"))) <> 0 then
            	
		            strSQL = "UPDATE job SET jobstatus = 0 WHERE id = "& jobid  'jobnr = "& intJobnr
		            oConn.execute(strSQL)
            			
            				
				            '**** Finder jobansvarlige *****
				            strSQL = "SELECT job.id AS jid, jobnavn, jobans1, jobans2, m1.mnavn AS m1mnavn, m1.email AS m1email,"_
				            &" m2.mnavn AS m2mnavn, m2.email AS m2email FROM job "_
				            &" LEFT JOIN medarbejdere m1 ON (m1.mid = jobans1)"_
				            &" LEFT JOIN medarbejdere m2 ON (m2.mid = jobans2)"_
				            &" WHERE job.id = "& jobid
				            oRec.open strSQL, oConn, 3
				            x = 0
				            if not oRec.EOF then
            				
				            jobid = oRec("jid")
				            jobans1 = oRec("m1mnavn")
				            jobans2 = oRec("m2mnavn")
				            jobans1email = oRec("m1email")
				            jobans2email = oRec("m2email")
				            jobnavnThis = oRec("jobnavn")
            				
				            x = x + 1
				            end if
				            oRec.close
            				
            				
				            '*** Henter afsender **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& medarbejderid
				            oRec.open strSQL, oConn, 3
            				
				            if not oRec.EOF then
            				
				            afsNavn = oRec("mnavn")
				            afsEmail = oRec("email")
            				
				            end if
				            oRec.close
            				
            				
            					
            						
            					
				            if x <> 0 AND request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\timereg_akt_2006.asp" then
					  	            'Sender notifikations mail
		                            Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            Mailer.CharSet = 2
		                            Mailer.FromName = "TimeOut /" & afsNavn 
		                            Mailer.FromAddress = afsEmail
		                            Mailer.RemoteHost = "webmail.abusiness.dk" '"pasmtp.tele.dk"
						            Mailer.AddRecipient "" & jobans1 & "", "" & jobans1email & ""
		                            Mailer.AddRecipient "" & jobans2 & "", "" & jobans2email & ""
            						
            						    
						            Mailer.Subject = jobnavnThis &" ("& intJobnr &") - Afsluttet."
		                            strBody = "Hej jobansvarlige." & vbCrLf & vbCrLf
						            strBody = strBody &"Job: "& jobnavnThis &" ("& intJobnr &") er nu afsluttet/lukket.  "& vbCrLf & vbCrLf
		                            strBody = strBody & "Med venlig hilsen" & vbCrLf
		                            strBody = strBody & session("user") & vbCrLf & vbCrLf
            		                
            		
            		
		                            Mailer.BodyText = strBody
            		
		                            Mailer.sendmail()
		                            Set Mailer = Nothing
				            end if' x
            	
	                    if request("FM_kopierjob") = "1" then
	                    
	                    'Response.Write "kopier job"
	                    'Response.end
	                    Response.Redirect "job_kopier.asp?func=kopier&id="&jobid&"&fm_kunde=0&filt=aaben&showaspopup=y&rdir=timereg&usemrn="&medarbejderid&"&showakt=1"
	                    
	                    %>
	                    
	                 
	                    <%
	                   
	                       
                    	    
	                    end if '*** Kopier job **'
            	
	            end if '** Luk job **'
	
	end if '** AntalErr **'
	
	
	 '** hvis der skal tjekkes values fra timeindlæsningen brus .end her ***'
	 '** Response.end
	
	    '**** Opdaterer stopurs siden / Redirect til Timregside ****'
	    if stopur = "1" then 
	    Response.Redirect "stopur_2008.asp?reload=1"
	    else
	    Response.Redirect "timereg_akt_2006.asp?showakt=1"    
	    end if
	
	
	
	end if '** func = db **'
	
	
	
	
	
	'******************************************************'
	'************ Opdater Smiley **************************'
	'******************************************************'
	 if func = "opdatersmiley" then
	 call opdaterSmiley
     
     
     
    
        
        '*** Vender tilbage til timereg ****'
        %>
        <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
        <div style="position:absolute; top:100px; left:200px; width:400px; padding:2px; border:2px red dashed; background-color:#ffffff;">
        
        <table cellspacing="0" cellpadding="10" border="0" bgcolor="#ffffe1" width=100%>
	    <tr>
	        <td bgcolor="#ffffff" valign=bottom style="border-bottom:1px #999999 solid; padding-top:15px;">
	        <h4><%=tsa_txt_001%>!</h4></td>
	        <td valign="top" align=right bgcolor="#ffffff" style="border-bottom:1px #999999 solid;">
		    <img src="../ill/about_32.png" alt="Info" border="0">
		    </td>
	        </tr>
	        <tr>
	        <td colspan=2>
	        <%=tsa_txt_002%>:<br /> <b><%=weekdayname(datepart("w",cDateUge,1)) %> d. 
	        <%=formatdatetime(cDateUge, 2) &" "& tsa_txt_003 %>. 
	        <%=formatdatetime(cDateUge, 3) %></b><br />
	        <br />
	        <%=tsa_txt_004 %>:<br /> <b><%=weekdayname(datepart("w",cDateAfs,1)) %> d. 
	        <%=formatdatetime(cDateAfs, 2) &" "& tsa_txt_003 &". "& formatdatetime(cDateAfs, 3) %></b>
	        <br />
	        <b><%=tsa_txt_005 &" "& datepart("ww", cDateUgeTilAfslutning, 2, 2)%> <%=smileysttxt %></b>
	        
	        
	        </td>
	    </tr>
	<tr>
	<td>
	 <a href="timereg_akt_2006.asp?fromsdsk=<%=fromsdsk%>" class=vmenu><%=tsa_txt_006 %> >></a>
        
		</td>
	</tr>
</table>
</div>
        <%
        Response.end
	   
	    
	
	end if '*** Opdater smiley func **'
	
	
	
	
	
	
	
	   
	
	
	   
	%>
	
	
	
	
	<%
	
	%><!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<SCRIPT language=javascript src="inc/timereg_2006_func.js"></script>
	
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
    <div id="sindhold" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(1)%>
	</div>
	
	
	
	<div id="sekmenu" style="position:absolute; left:15; top:97; visibility:visible;">
	<%call tregsubmenu() %>
	</div>
	
	

	
	
	 
	<%
	
	
	
	
	'**** Aktiviteter på det valgte job ***********'
	'** Ved forste logon i sessionen må cookie ikke vælges
	'** Således at man altid får vist sig selv når men logger på selvom man har tastet timer inde
	'** for en anden sidst man var logget på ***'
	if session("forste") <> "n" then
	session("forste") = "j"
	end if
	'****** Findes nu i login.asp ****'
	
	
	'**** Jobid, showakt og usemrn sættes i sharepoint.asp link og hentes i timreg__2006_fs.asp **'
	if session("fromsharepoint") = "a6196f3646c2e60eddb95cd2134d457f" then
	
	showakt = 1
    jobid = session("shpjobid")
    usemrn = session("mid")      	
	
	
	else
	             
    

     '******************************************************************************
    '****************** Medarb *******************************************************
    '******************************************************************************


                '** Valgt mid ****
                if len(request("usemrn")) <> 0 then
                usemrn = request("usemrn")
                else
                    if len(request.Cookies("tsa")("usemrn")) <> 0 AND session("forste") = "n" then
                    usemrn = request.Cookies("tsa")("usemrn")
                    else
                    usemrn = session("mid")
                    end if
                end if
            	
            	response.cookies("timereg_2006")("usemrn") = usemrn

    
    '******************************************************************************
    '****************** Job *******************************************************
    '******************************************************************************


             



                 '*** Det valgte job ****'
                if len(request("jobid")) <> 0 then

                jobid = replace(request("seljobid"), "0,", "") 
                jobidsSeljids = split(request("jobid"), ",")
                for i = 0 TO UBOUND(jobidsSeljids)
                    
                    if cint(instr(jobid, "#"& jobidsSeljids(i) & "#,")) = 0 AND cint(instr(jobid, ","& jobidsSeljids(i) & ",")) = 0 then
                      jobid = jobid &",#"& trim(jobidsSeljids(i)) & "#" 
                    end if

                next

                'jobid = replace(jobid, ",##", "")
                jobid = replace(jobid, " ", "")
                jobid = replace(jobid, "#", "")
                jobid = trim(jobid)
                jobid = "0,"&jobid &",0"

               

                    '** Nulstillerforvalgt job for denne medarbejder ****'
                    strSQLUpdFvOff = "UPDATE timereg_usejob SET forvalgt = 0 WHERE medarb = "& usemrn &"" 
                    oConn.Execute(strSQLUpdFvOff)
                    '******************************************************'
                   

                else	
                    'if len(request.Cookies("tsa")("jobid")) <> 0 then
                    'jobid = request.Cookies("tsa")("jobid")
                    'else
            		    
                        select case lto
                        case "demo"
                        jobid = 42
                        case "external"
                        jobid = 127
                        case else
                        jobid = 0
                        
                        '*** Ellers åbnes det job der sidst er reg. timer på ***'
                        'strSQLlastjobid = "SELECT j.id AS jid FROM timer "_
                        '&" LEFT JOIN job j ON (j.jobnr = tjobnr) WHERE tmnr = "& session("mid") &" ORDER BY tid DESC LIMIT 1"
                        ''Response.Write "<br><br><br><br><br><br>"& strSQLlastjobid
                        ''Response.flush 
                        
                        'oRec.open strSQLlastjobid, oConn, 3
                        'if not oRec.EOF then
                        'jobid = oRec("jid")
                        'end if
                        'oRec.close
                        

                           
                        jobid = "0"
                        strSQLfv = "SELECT jobid FROM timereg_usejob WHERE medarb = "& usemrn & " AND forvalgt = 1"
                        oRec5.open strSQLfv, oConn, 3
                        while not oRec5.EOF
            
                        jobid = jobid &","& oRec5("jobid") 

                        oRec5.movenext
                        wend
                        oRec5.close

                        jobid = jobid & ",0"
                        end select
            		    
                    'end if
                end if

                
                '** Er der valgt et eller flere job ***'
                'dim jobidss
                'redim jobidss(2)
                if instr(jobid, ",") <> 0 then

                seljobidSQL = " j.id = 0" 

                jobids = split(jobid, ",") 
                for j = 0 to UBOUND(jobids)
                seljobidSQL = seljobidSQL & " OR j.id = "& replace(jobids(j), "#", "") &" "
                        

                        '** Sætter forvalgt job for denne medarbejder ****'
                        strSQLUpdFvOn = "UPDATE timereg_usejob SET forvalgt = 1 WHERE medarb = "& usemrn &" AND jobid = "& jobids(j)
                        oConn.Execute(strSQLUpdFvOn)
                        '*************************************************'

                next

              
                else

                seljobidSQL = seljobidSQL & " j.id = "& jobid &" "
                Redim jobids(0) 
                jobids(0) = jobid

                        '** Sætter forvalgt job for denne medarbejder ****'
                        strSQLUpdFvOn = "UPDATE timereg_usejob SET forvalgt = 1 WHERE medarb = "& usemrn &" AND jobid = "& jobid
                        oConn.Execute(strSQLUpdFvOn)
                        '*************************************************'


                end if


            	
            	
                    if len(trim(request("showakt"))) <> 0 then
                    showakt = request("showakt")
                   
                    else
                        if len(request.Cookies("showakt")) <> 0 then
                        showakt = request.Cookies("showakt")
                        else
                             'if lto = "demo" then
                             'showakt = 0
                             showakt = 1
                             'else
                             'showakt = 1
                             'Response.Write "her"
                             'end if
                        end if
                    end if
                    
               
    
    
	
	end if 'sharepointlink
	
	'showakt = 0
	
	if request.Cookies("showjobinfo") <> "" then
	showjobinfo = request.Cookies("showjobinfo")
	else  
	    select case lto
	    case "demo"
	    showjobinfo = 0
	    case else
	    showjobinfo = 1
	    end select
	end if
	
	if len(request("FM_vistimereltid")) <> 0 then
	visTimerElTid = request("FM_vistimereltid")
	else
	
		if len(request.Cookies("tsa")("timereltid")) <> 0 then
		visTimerElTid = request.Cookies("tsa")("timereltid")
		else
		visTimerElTid = 0
		end if
		
	end if
	
	'** tidslås ***
	if len(request("FM_chkfor_ignorertidslas")) <> 0 then
	    
	   
	    if len(request("FM_igtlaas")) <> 0 then
	    ignorertidslas = 1
	    else
	    ignorertidslas = 0
	    end if
	
	else
	    
	    if request.Cookies("tsa")("igtidslas") <> "" then
		ignorertidslas = request.Cookies("tsa")("igtidslas")
		else
		ignorertidslas = 0
		end if
	   
	end if
	
	
	'**** Job og Akt. periode ***'
	if len(request("FM_chkfor_ignJobogAktper")) <> 0 then
	    
	   
	    if len(request("FM_ignJobogAktper")) <> 0 then
	    ignJobogAktper = 1
	    else
	    ignJobogAktper = 0
	    end if
	
	else
	    
	    if request.Cookies("tsa")("ignJobogAktper") <> "" then
		ignJobogAktper = request.Cookies("tsa")("ignJobogAktper")
		else
		ignJobogAktper = 0
		end if
	   
	end if
	
	response.Cookies("tsa")("ignJobogAktper") = ignJobogAktper
	response.Cookies("tsa")("igtidslas") = ignorertidslas

    if len(trim(jobid)) <> 0 then
    jobid = jobid
    else
    jobid = 0
    end if
	
    'response.Cookies("tsa")("jobid") = jobid
	response.Cookies("showakt") = showakt
	response.Cookies("tsa")("usemrn") = usemrn
	response.Cookies("tsa")("timereltid") = visTimerElTid
	response.Cookies("tsa").expires = date + 32
	
	
	call erSDSKaktiv()
	
	%>
	
	
	
	
	
	
	<%
	level = session("rettigheder")
	timenow = formatdatetime(now, 3)
	%>
	
	
	<%
	
	'**** Divs ***'
	
	Select case showakt
	case 1
	dKalVzb = "visible"
	dKalDsp = ""
	dTimVzb = "visible"
	dTimDsp = ""
	
    dAfstemVzb = "hidden"
	dAfstemDsp = "none"
	dStempelVzb = "hidden"
	dStempelDsp = "none"
	sVzb = "visible"
	sDsp = ""
	urVzb = "visible"
	urDsp = ""
	phVzb = "visible"
	phDsp = ""
	
	fiVzb = "Visible"
	fiDsp = ""
	
	jinf_knap_vzb = "visible"
	jinf_knap_dsp = ""
	
	
	
	case 2
	dKalVzb = "hidden"
	dKalDsp = "none"
	dTimVzb = "hidden"
	dTimDsp = "none"
	dJobVzb = "hidden"
	dJobDsp = "none"
	dAfstemVzb = "visible"
	dAfstemDsp = ""
	dStempelVzb = "hidden"
	dStempelDsp = "none"
	sVzb = "hidden"
	sDsp = "none"
	urVzb = "hidden"
	urDsp = "none"
	phVzb = "hidden"
	phDsp = "none"
	
    fiVzb = "hidden"
	fiDsp = "none"
	jinftdBG = "#EFF3FF"
	jinf_knap_vzb = "hidden"
	jinf_knap_dsp = "none"
	
	case 3
	dKalVzb = "hidden"
	dKalDsp = "none"
	dTimVzb = "hidden"
	dTimDsp = "none"
	dJobVzb = "hidden"
	dJobDsp = "none"
	dAfstemVzb = "hidden"
	dAfstemDsp = "none"
	dStempelVzb = "visible"
	dStempelDsp = ""
	sVzb = "hidden"
	sDsp = "none"
	urVzb = "hidden"
	urDsp = "none"
	phVzb = "hidden"
	phDsp = "none"
	
	jinftdBG = "#EFF3FF"
    fiVzb = "hidden"
	fiDsp = "none"
	jinf_knap_vzb = "hidden"
	jinf_knap_dsp = "none"
	
	
	case else
	
	dKalVzb = "hidden"
	dKalDsp = "none"
	dTimVzb = "hidden"
	dTimDsp = "none"
		
	dAfstemVzb = "hidden"
	dAfstemDsp = "none"
	dStempelVzb = "hidden"
	dStempelDsp = "none"
	
	sVzb = "hidden"
	sDsp = "none"
	urVzb = "hidden"
	urDsp = "none"
	phVzb = "hidden"
	phDsp = "none"
	
	
    fiVzb = "visible"
	fiDsp = ""
	
	jinf_knap_vzb = "visible"
	jinf_knap_dsp = ""
	
	end select%>
	
	<% 
	if cint(fromsdsk) = 1 then
	addVal = 20
	else
	addVal = 100
	end if
	
	'**srctip **'
    
   
    
    %>
     <div id="kalender" style="position:absolute; left:780px; top:153px; width:210px; visibility:<%=dKalVzb%>; display:<%=dKalDsp%>;">
	<!--#include file="../inc/regular/calender_2006.asp"-->
	<!--#include file="inc/timereg_dage_2006_inc.asp"-->
	</div>
    
   <%
   varTjDatoUS_man = convertDateYMD(tjekdag(1))
	varTjDatoUS_tir = convertDateYMD(tjekdag(2))
	varTjDatoUS_ons = convertDateYMD(tjekdag(3)) 
	varTjDatoUS_tor = convertDateYMD(tjekdag(4)) 
	varTjDatoUS_fre = convertDateYMD(tjekdag(5))
	varTjDatoUS_lor = convertDateYMD(tjekdag(6))
	varTjDatoUS_son = convertDateYMD(tjekdag(7))
   
    '** redirect til afstem toto, hvis det er den er vist sidst. ***
    if showakt = "5" then
    Response.Redirect "afstem_tot.asp?usemrn="&usemrn&"&show=5&varTjDatoUS_man="&varTjDatoUS_man 
    Response.end
    end if

    
    call treg_3menu(thisfile)
    %>
    
    
    
    
    
   
	
    
    <div id="loadbar" style="position:absolute; display:; visibility:visible; top:260px; left:200px; width:300px; background-color:#ffffff; border:1px #cccccc solid; padding:2px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>

	</div>
	
	
	
	
	
	
	
    
    
	
	<%
	
	
	
	
	 
 '************************************************************************************************** 
  'Opbygger timereg SQL states
  '**************************************************************************************************
 	            
 	        
				
				
				if len(trim(request("FM_kontakt"))) <> 0 then
				show_sogjobnavn_nr = request("FM_sog_job_navn_nr")
				response.cookies("tsa")("jobsog") = show_sogjobnavn_nr
				else
				    if request.cookies("tsa")("jobsog") <> "" then
				    show_sogjobnavn_nr = request.cookies("tsa")("jobsog")
				    else
				    show_sogjobnavn_nr = ""
				    end if
				end if
				
				
				if len(request("FM_kontakt")) <> 0 then
				sogKontakter = request("FM_kontakt")
				
				
				response.cookies("tsa")("kontakt") = sogKontakter
	            response.cookies("tsa").expires = date + 65
				
				else
				    
				        if request.Cookies("tsa")("kontakt") <> "" _
				        AND request.Cookies("tsa")("kontakt") <> 0 _
				        then
				        sogKontakter = request.Cookies("tsa")("kontakt")
				        else
				        sogKontakter = 0
                        end if
				end if
				
				
				if len(request("FM_ignorer_projektgrupper")) <> 0 then
				ignorerProgrp = request("FM_ignorer_projektgrupper") '1
					if cint(ignorerProgrp) = 1 then
					selIgn = "CHECKED"
					else
					selIgn = ""
					end if
				else
				ignorerProgrp = 0
				selIgn = ""
				end if
				
				if len(request("FM_nyeste")) <> 0 then
				    intNyeste = request("FM_nyeste") '1
					if cint(intNyeste) = 1 then
					selNye = "CHECKED"
					else
					selNye = ""
					end if
					
				else
				    if request.cookies("nyeste") = "1" then
				    intNyeste = 1
				    selNye = "CHECKED"
				    else
				    intNyeste = 0
				    selNye = ""
				    end if
				end if
				
				if len(request("FM_easyreg")) <> 0 then
				    intEasyreg = request("FM_easyreg") '1
					if cint(intEasyreg) = 1 then
					selEasy = "CHECKED"
					else
					selEasy = ""
					end if
					
				else
				    if request.cookies("easyreg") = "1" then
				    intEasyreg = 1
				    selEasy = "CHECKED"
				    else
				    intEasyreg = 0
				    selEasy = ""
				    end if
				end if

                if len(request("FM_smartreg")) <> 0 then
				    intSmartreg = request("FM_smartreg") '1
					if cint(intSmartreg) = 1 then
					selSmart = "CHECKED"
					else
					selSmart = ""
					end if
					
				else
				    if request.cookies("smartreg") = "1" then
				    intSmartreg = 1
				    selSmart = "CHECKED"
				    else
				    intSmartreg = 0
				    selSmart = ""
				    end if
				end if



                 if len(request("FM_classicreg")) <> 0 then
				    intClassicreg = request("FM_classicreg") '1
					if cint(intClassicreg) = 1 then
					selClassic = "CHECKED"
					else
					selClassic = ""
					end if
					
				else
				    if request.cookies("classicreg") = "1" OR (intSmartreg = 0 AND intEasyreg = 0 AND intNyeste = 0) then
				    intClassicreg  = 1
				    selClassic = "CHECKED"
				    else
				    intClassicreg  = 0
				    selClassic = ""
				    end if
				end if


				


				if selNye = "CHECKED" then
				sotDivDSP = "none"
				sotDivVzb = "hidden"
				else
				sotDivDSP = ""
				sotDivVzb = "visible"
				end if
				
				
				if len(request("sort")) <> 0 then
				sortVal = request("sort")
				else
				    if request.cookies("treg_sort") <> "" then
				    sortVal = request.cookies("treg_sort")
				    else
				    select case lto
				    case "dencker"
				    sortVal = 3
				    case else 
				    sortVal = 0
				    end select
				    end if
				end if
				
				
				sort1CHK = ""
				sort2CHK = ""
				sort3CHK = ""
				sort0CHK = ""
				
				select case sortVal
				case 1
				sort1CHK = "CHECKED"
				case 2
				sort2CHK = "CHECKED"
				case 3
				sort3CHK = "CHECKED"
				case else
				sort0CHK = "CHECKED"
				end select
				
				
	call hentbgrppamedarb(usemrn)
				
	
	'**************************************************************************************************
	
	
	
	
	
	call filterheaderid(80,20,744,pTxt,fiVzb,fiDsp,"filter", "relative")%>
	<table cellspacing=0 width=100% cellpadding=1 border=0>
	<form action="timereg_akt_2006.asp" id="filterkri">
    <input type="hidden" name="dtson" id="datoSonSQL" value="<%=now%>">
	<%if level = 1 then
	SQLkriEkstern = ""
    end if
	
	''*** Finder medarbejdere i de progrp hvor man er teamleder ***'
	call projgrp(-1,level,session("mid"),1)
	    
	     medarbgrpIdSQLkri = "AND (mid = "& session("mid")
    
	    
	    for p = 0 to prgAntal
	     
	     if prjGoptionsId(p) <> 0 then
	        call medarbiprojgrp(prjGoptionsId(p), session("mid"))
	     end if
	    
	    next 
	    
	     medarbgrpIdSQLkri = medarbgrpIdSQLkri & ")"
	    
	strSQLmids = medarbgrpIdSQLkri '" AND mid = "& usemrn
	%>
	<tr bgcolor="#ffffff">
	<td valign=top rowspan=2 style="padding-top:5px; padding-bottom:5px;"><b><%=tsa_txt_333 &" "& tsa_txt_077 %>:</b> (<%=tsa_txt_357 %>)
	<br />
				<%
					strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe FROM medarbejdere WHERE mansat <> 2 "& strSQLmids &" GROUP BY mid ORDER BY Mnavn"
					
					'Response.Write strsQL
					'Response.flush
					
					%>
					<select name="usemrn" id="usemrn" style="font-size: 9px; width:250px;" onchange="submit();">
					<%
					
					oRec.open strSQL, oConn, 3
					while not oRec.EOF 
					
					if cint(oRec("Mid")) = cint(usemrn) then
					rchk = "SELECTED"
					else
					rchk = ""
					end if%>
					<option value="<%=oRec("Mid")%>" <%=rchk%>><%=oRec("mnavn")%></option>
					<%
					
					
					oRec.movenext
					wend
					oRec.close
				%></select>
				
				 <!--
				 &nbsp;&nbsp;&nbsp;&nbsp;<a href="../timereg2006/timereg_2006_fs.asp" class=red>tilbage til den gamle timereg. side..</a>
				 -->
				
				<%
	'**** Finder kunder til dd *****'
	'**** Tag højde for ignorer projektgrupper? **'
	if cint(ignorerProgrp) = 0 then
	
	    
		
		strKundeKri = " AND ("
		strSQL3 = "SELECT j.id, medarb, jobid, easyreg, kid, jobknr FROM timereg_usejob "_
		&" LEFT JOIN job j ON (j.id = jobid) "_
		&" LEFT JOIN kunder ON (kid = j.jobknr) "_
		&" WHERE medarb = "& usemrn &" AND kid <> '' AND jobstatus = 1  GROUP BY kid ORDER BY kid"
	    'response.write strSQL3
		'Response.flush
		oRec3.open strSQL3, oConn, 3
		
		while not oRec3.EOF 
		strKundeKri = strKundeKri & " kid = "& oRec3("kid") & " OR "
		oRec3.movenext
		wend 
		
		oRec3.close 
    else
    strKundeKri = " AND (kid <> 0 OR "
    end if
	%>
	<br /><br /><b><%=tsa_txt_075 %>:</b><br />
	
	<%
	
	    if len(strKundeKri) <> 0 then
		strKundeKri = strKundeKri &" kid = 0)"
		else
		strKundeKri = strKundeKri &" AND (kid = 0)"
		end if
		
		ketypeKri = " ketype <> 'e'"
		
		
		if cint(ignorerProgrp) = 0 then
	    strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" "& strKundeKri &" GROUP BY kid ORDER BY Kkundenavn"
	    else
	    strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder, job WHERE "& ketypeKri &" "& strKundeKri &" AND jobstatus = 1 AND jobknr = kid GROUP BY kid ORDER BY Kkundenavn"
	    end if
	    'Response.write strSQL
		'Response.flush
	
	%>
	<select name="FM_kontakt" id="FM_kontakt" style="font-size:9px; width:250px;" size=8 onchange="renssog();"> <!-- onChange="renssog()" --> 
		
		<%if sogKontakter = 0 then
		allSel = "SELECTED"
		else
		allSel = ""
		end if %>
		
		<option value="0" <%=allSel %>><%=tsa_txt_076 %></option>
		<%
		
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(sogKontakter) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn") & " ("& oRec("Kkundenr")&")"%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		
		</select>
       	
       	<br /><br />
       	<b><%=tsa_txt_078 %>:</b> (<%=tsa_txt_073 %>)<br />
		<input type="text" name="FM_sog_job_navn_nr" id="FM_sog_job_navn_nr" value="<%=show_sogjobnavn_nr%>" style="font-family:arial; font-size:10px; width:250px;"> 
		<br /><input type="checkbox" name="FM_ignorer_projektgrupper" id="FM_ignorer_projektgrupper" value="1" <%=selIgn%>> <%=tsa_txt_074%> 
		<a href="javascript:popUp('guiden_2006.asp?mid=<%=usemrn%>','650','620','150','120');" target="_self" class=rmenu>
		(+ <%=lcase(tsa_txt_082)%>)
		</a>
		<br /><br>
            <div style="position:relative; padding:5px; border:1px #cccccc dashed; width:250px; background-color:#eff3ff">
            <b>Visning:</b><br />
            <input id="FM_nyeste" name="FM_nyeste" type="radio" value="1" <%=selNye%> /> <%=tsa_txt_334 %>
            
            <%if (lto = "dencker" AND level = 1) OR lto = "epi" OR lto = "wowern" OR lto = "outz" then %>
            <br /><input id="FM_easyreg" name="FM_easyreg" type="radio" value="1" <%=selEasy%> /> <%=tsa_txt_358 %> 
            <%end if %>
		
			
            <!--<br /><input id="FM_smartreg" name="FM_smartreg" type="radio" value="1" <%=selSmart%> /> <%=tsa_txt_373 %> -->
            <br /><input id="FM_classicreg" name="FM_classicreg" type="radio" value="1" <%=selClassic%> /> <%=tsa_txt_374 %> 
				
            </div>

				</td>
				
				
	    <td valign=top style="height:200px; padding-top:5px;"><b><%=tsa_txt_336 %>:</b>
	    
        
	    <%if level <= 3 OR level = 6 then %>
	    &nbsp;<a href="jobs.asp?menu=job&func=opret&id=0&int=1&rdir=treg" class=rmenu><%=tsa_txt_335 %> >></a><br />
        <%end if %>
	    
	   
	    <div id="sorterdiv" style="display:<%=sotDivDsp%>; visibility:<%=sotDivVzb%>; padding:3px 3px 3px 0px; border:0px #cccccc solid;">
	        <b><%=tsa_txt_341 %>:</b> <input id="sort1" name="sort" type="radio" value=1 <%=sort1CHK%> /> <%=tsa_txt_342 %>&nbsp;
             <input id="sort2" name="sort" type="radio" value=2 <%=sort2CHK%> /> <%=tsa_txt_343 %>&nbsp;
              <input id="sort0" name="sort" type="radio" value=0 <%=sort0CHK%> /> <%=tsa_txt_344&", "&lcase(tsa_txt_342) %>&nbsp;
              <input id="sort3" name="sort" type="radio" value=3 <%=sort3CHK%> /> <%=tsa_txt_344&", "&lcase(tsa_txt_343) %>
              </div>
	   
	  <!--
	   <textarea id="fajl" cols="20" rows="2"></textarea>
       -->
       
	   
	   
	   
	   
	   
	    <select name="jobid" id="jobid" style="width:450px; font-size:9px;" size="20" multiple>
	    <!-- henter job fra jquery -->
	    </select>
        <br>
        <span style="font-size:9px; color:#999999;">Der vises maks. 100 job i job-oversigten.</span>
	    
            
         <input id="jobid_nul" name="jobid" type="hidden" value="0" /> 
         <input id="seljobid" name="seljobid" type="hidden" value="<%=jobid %>" />
         <input id="showakt" name="showakt" type="hidden" value="1" />
	    
	  
              
	    </td>
				
				
				
	</tr>
	
	
	
	    <tr><td valign=bottom align=right style="height:30px;">
	    <input type="submit" id="sogsubmit" value="<%=tsa_txt_314 &" "& lcase(tsa_txt_244) %> >>">
	    </td></tr>
	    </table>
	    </td>
	    </form>
	</tr>
	
	
	</table>
		<!-- filter header -->
	</td></tr></table>

 </div>
	
	
	
	
	<!--- tjekker adgang via projektgrupper --->
	
	<%
	'if ignorerProgrp = 1 then
	
	select case lto 
	case "acc"
	pgeditok = 1
	case else
	    if level <= 2 OR level = 6 then
	    pgeditok = 1
	    else
	    pgeditok = 0
	    end if
	end select
	
	
	strJobmedadgang = "#0#"
	
		strSQL = "SELECT j.id AS id FROM job j "_
		&" WHERE "& varUseJob &" (j.jobstatus = 1 OR j.jobstatus = 3) "& strPgrpSQLkri & "  ORDER BY j.id" 
		
        'AND (j.jobstartdato <= '"& varTjDatoUS_man &"')
        'Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"& strSQL
        'Response.flush

		oRec.open strSQL, oConn, 3
		while not oRec.EOF
		
		strJobmedadgang = strJobmedadgang &",#"& oRec("id")&"#"
		
		oRec.movenext
		wend
		
		oRec.close
	
	
	
    
    
    
    if (instr(strJobmedadgang, "#"& jobids(0) &"#") <> 0) then
		projektgrpOK = 1
		
			strprojektgrpOK = "&nbsp;"
		
		else
		projektgrpOK = 0
			if pgeditok = 1 then
			strprojektgrpOK = "<b>"& tsa_txt_351 &"</b><br>"& tsa_txt_352 &" <a href=""javascript:popUp('tilknytprojektgrupper.asp?id="&jobid&"&medid="&usemrn&"','600','500','150','120');"" target=""_self""; class=vmenu>"& tsa_txt_353 &"..</a>"  
			'"<b>Du har ikke adgang til dette job.</b><br>Du kan tilmelde dig jobbet ved at tilføje en projektgruppe du er medlem af. Du kan tilføje en projektgruppe ved at klikke <a href=""javascript:popUp('tilknytprojektgrupper.asp?id="&jobid&"&medid="&usemrn&"','600','500','150','120');"" target=""_self""; class=vmenu>her..</a>"
			else
			
			strprojektgrpOK = "<b>"& tsa_txt_351 &"</b>" '"<b>Du har ikke adgang til dette job.</b>"
			end if
		end if
	
	
	'Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>projektgrpOK "& projektgrpOK
	
	'else
	'projektgrpOK = 1
	'end if
	
    '**************** Aktive job, job til afv. +TODO LIST ********'
     %>

        <div id="div2" style="position:absolute; left:779px; top:418px; width:210px; height:280px; visibility:<%=jinf_knap_vzb%>; display:<%=jinf_knap_dsp%>; z-index:100; border:1px #cccccc solid; padding:2px; background-color:#ffffff; overflow:auto;">
	     <table cellpadding=2 cellspacing=0 border=0 width=100%>
         <tr bgcolor="#5C75AA"><td class=alt style="width:165px;"><b>
         Aktive job på medarb.:</b>
         </td>
         <td class=lille>Fjern</td>
         </tr>


         <%strSQLfv = "select jobid, medarb, j.jobnavn, j.jobnr, mnavn, mnr, tu.id AS tuid FROM timereg_usejob AS tu"_
         &" LEFT JOIN job AS j on (j.id = tu.jobid) "_
         &" LEFT JOIN medarbejdere AS m ON (m.mid = tu.medarb) WHERE forvalgt = 1 ORDER BY medarb, forvalgt_sortorder"
         
         'Response.write strSQLfv
         'Response.Flush

         lastMid = 0
         oRec3.open strSQLfv, oConn ,3
         while not oRec3.EOF 
         
         
         if lastMid <> oRec3("medarb") OR lastMid = 0 then%>
         <tr bgcolor="8caae6"><td colspan=2 class=lille><b><%=oRec3("mnavn") %></b></td></tr>
         <% end if 

         %>
         <tr><td class=lille style="border-bottom:1px #CCCCCC solid;"><%=oRec3("jobnavn") %> (<%=oRec3("jobnr") %>)</td>
         <td style="border-bottom:1px #CCCCCC solid;"><input id="tuid_<%=oRec3("tuid") %>" class="aktivejobid" type="checkbox" value="1" /></td>
         </tr>

         <%

         lastMid = oRec3("medarb")
         oRec3.movenext
         wend
         oRec3.close   %>


         </table>
        </div>


        <div id="div3" style="position:absolute; left:779px; top:718px; width:210px; height:300px; visibility:<%=jinf_knap_vzb%>; display:<%=jinf_knap_dsp%>; z-index:100; border:1px #cccccc solid; padding:2px; background-color:#ffffff; overflow:auto;">
	     <table cellpadding=2 cellspacing=0 border=0 width=100%>
         <tr bgcolor="#5C75AA"><td class=alt><b>
         Slutdato i denne uge:</b>
         </td></tr>
         <tr><td class=lille>
         Vis kun job hvor Du er: <br /><input type="checkbox" value="1" name="denneuge_jobans" /> Jobansvarlig / jobejer<br /><input type="checkbox" value="1" name="denneuge_jobans" /> Tilknyttet via projektgrp.   
         </td></tr>

          <tr bgcolor="#EfF3FF"><td><b>
         Job:</b>
         </td></tr>

         <tr><td>

         <div id="jobidenneuge">
        </div>

         </td></tr>

         </table>
        </div>
	    
        <div id="div_todolist" style="position:absolute; left:779px; top:1028px; width:210px; height:200px; visibility:<%=jinf_knap_vzb%>; display:<%=jinf_knap_dsp%>; z-index:100; border:1px #cccccc solid; padding:2px; background-color:#ffffff; overflow:auto;">
	    
       
        <table cellpadding=2 cellspacing=0 border=0 width=100%>
         <tr bgcolor="#5C75AA"><td id="todotd" colspan="2" style="padding:3px; height:20px;" class=alt><b>ToDo liste</b> (seneste 20)</td></tr>
         <tr><td colspan="2"> <a href="webblik_todo.asp?nomenu=1" class=rmenu target="_blank">Se / rediger ToDo listen >></a></td></tr>
         <tr><td class=lille style="padding-left:5px;">ToDo's</td><td class=lille align=right style="padding-right:5px;">Afslut</td></tr>
        <% 
        
        idag_365 = dateadd("d", -365, now)
        idag_365 = year(idag_365) &"/"& month(idag_365) & "/" & day(idag_365)

        strSQLtodo = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, "_
	    &" t.afsluttet, t.sortorder, trn.medarbid, trn.todoid, t.tododato, t.forvafsl, t.public, t.parent "_
	    &" FROM todo_rel_new trn "_
	    &" LEFT JOIN todo_new t ON (t.id = trn.todoid AND t.afsluttet <> 1)"_
	    &" WHERE ((medarbid = "& useMrn &" AND trn.todoid = t.id) OR (t.public = 1)) AND afsluttet = 0 AND t.dato > '"& idag_365 &"' GROUP BY t.id ORDER BY t.dato DESC, t.id DESC LIMIT 20"

        'AND parent = 0
        'Response.Write strSQLtodo

        oRec.open strSQLtodo, oConn, 3
        t = 0
        while not oRec.EOF

        select case right(t, 1)
        case 0,2,4,6,8
        tBg = "#FFFFFF"
        case else
        tBg = "#EFf3ff"
        end select


        if t = 0 AND cdate(oRec("dato")) = cdate(date()) then
        tbg = "#FFFFe1"
        end if

        if len(trim(oRec("navn"))) > 200 then
        todoTxt = left(oRec("navn"), 200) & ".."
        else
        todoTxt = oRec("navn")
        end if

        if len(todoTxt) > 15 then
        len_todoTxt = len(todoTxt)
        left_todoTxt = "<b>"& left(todoTxt, 15) & "</b>"
        right_todoTxt = right(todoTxt, len_todoTxt- 15)
        
        todoTxt = left_todoTxt & right_todoTxt 

        else

        todoTxt = "<b>"& todoTxt &"</b>"

        end if

        if oRec("public") = 0 then
        privatTxt = ""
        else
        privatTxt = "(off.lig)"
        end if

        if oRec("forvafsl") = 1 then
        todoDatoTxt = "Forv. udført: " & day(oRec("tododato")) & ". "& left(monthname(month(oRec("tododato"))), 3)&" "& year(oRec("tododato"))
        else
        todoDatoTxt = ""
        end if

       %><tr bgcolor="<%=tBg%>">
       <td class=lille valign=top style="border-top:1px #cccccc solid; padding:4px 5px 4px 5px;">
       <%if oRec("parent") <> 0 then 
       
       strSQLpar = "SELECT navn FROM todo_new WHERE id = "& oRec("parent")
       oRec2.open strSQLpar, oConn, 3
       if not oRec2.EOF then%>
       <%=left(oRec2("navn"), 20) %> >> 
       <%end if
       oRec2.close %>


       <%end if %>
       
       <%=todoTxt%> - <span style="color:#999999; font-size:9px;"><%=day(oRec("dato")) & ". "& left(monthname(month(oRec("dato"))), 3)&" "& year(oRec("dato"))%></span>
       <span style="color:darkred; font-size:9px;"><%=todoDatoTxt %></span></td>
       <td align=right valign=top style="border-top:1px #cccccc solid; padding:2px 5px 2px 5px;">
       <%if oRec("public") <> 1 then %>
       <input id="todoid_<%=oRec("id")%>" class="todoid" type="checkbox" value="1" />
       <%else %>
       <span style="color:#5582d2; font-size:9px;"><%=privatTxt %></span>
       <%end if %></td>
           
       </tr>
       
       <%
        t = t + 1
        oRec.movenext
        wend
        oRec.close
        %>
        <tr>
         <td valign=top colspan=2 style="border-top:1px #cccccc solid; padding:4px 5px 4px 5px;">&nbsp;</td>
        </tr>
        </table>

       
        </div>


	<%
	if cint(projektgrpOK) = 0 then
	
	
 
            
	
	'Response.end
	else%>
	 
	
	        

	
	     <%
		
				strSQL = "SELECT j.id, j.jobnavn, j.jobnr, k.kid, k.kkundenavn, k.kkundenr, "_
				&" j.jobans1, j.jobans2, jobstartdato, jobslutdato, j.beskrivelse, j.lukafmjob, "_
				&" m1.mnavn AS medjobans1, m1.mnr AS medjobans1mnr, m1.init AS medjobans1init, "_
				&" m2.mnavn AS medjobans2, m2.mnr AS medjobans2mnr, m2.init AS medjobans2init, "_
				&" m3.mnavn AS medkundeans1, m3.mnr AS mnr, m3.init AS init, j.jobstatus, "_
				&" j.budgettimer, j.ikkebudgettimer, j.jobtpris, j.fastpris, j.jobfaktype, s.navn AS saftnavn, "_
				&" s.status AS saftstatus, s.aftalenr AS saftnr, s.id AS aftid, jo_bruttooms, "_
				&" kundekpers, rekvnr, kk.navn AS kkpersnavn, k.sdskpriogrp, j.usejoborakt_tp, j.job_internbesk, j.restestimat, j.stade_tim_proc, k.useasfak "_
				&" FROM job j, kunder k, medarbejdere m "_
				&" LEFT JOIN medarbejdere m1 ON m1.mid = j.jobans1"_
				&" LEFT JOIN medarbejdere m2 ON m2.mid = j.jobans2"_
				&" LEFT JOIN medarbejdere m3 ON m3.mid = k.kundeans1"_
				&" LEFT JOIN serviceaft s ON (s.id = j.serviceaft)"_
				&" LEFT JOIN kontaktpers kk ON (kk.id = kundekpers)"_
				&" WHERE j.id = " & jobids(0) & " AND k.kid = j.jobknr"
				
				'Response.write strSQL
				'Response.flush
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
				
				'*** Jobbet er blevet lukket ved timregistrering ***
				if (cint(oRec("jobstatus")) <> 1 AND cint(oRec("jobstatus")) <> 3) OR cdate(oRec("jobstartdato")) > cDate(now) then 'tjekdag(7)
				
				Response.Redirect "timereg_akt_2006.asp?jobid=0&showakt=1"
                
				
				end if
				
				'** til timepriser **'
				j_intjobtype = oRec("fastpris")
				j_usejoborakt_tp = oRec("usejoborakt_tp") 
				j_timerforkalk = oRec("budgettimer") + oRec("ikkebudgettimer")
				j_budget = oRec("jobtpris")
				
				j_jobnavn_nr = oRec("jobnavn") & " ("& oRec("jobnr") &")"

                if oRec("jobstatus") = 3 then
                j_jobnavn_nr = j_jobnavn_nr & "<span style='font-size:12px; color:red;'> - Tilbud<span>"
	            end if

                restestimat = oRec("restestimat")
                stade_tim_proc = oRec("stade_tim_proc")
                aftid = oRec("aftid")

                if len(oRec("jo_bruttooms")) <> 0 then
		        jobbudget = oRec("jo_bruttooms")
		        else
		        jobbudget = 0
		        end if     

                useasfak = oRec("useasfak")
	    %>

		<div id="jinf_knap" style="position:absolute; left:779px; top:632px; width:210px; visibility:<%=jinf_knap_vzb%>; display:<%=jinf_knap_dsp%>; z-index:100000; border:1px #cccccc solid; padding:2px; background-color:#ffffff;">
	    <table cellpadding=2 cellspacing=0 border=0 width=100%>
	     <tr bgcolor="#5C75AA"><td style="padding:3px; height:20px;" class=alt><b>Links</b></td></tr>
	     <tr bgcolor="#FFFFFF">		        
				<td>
                <a href="materialer_indtast.asp?id=<%=jobid%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aftid%>" target="_blank" class=rmenu><%=tsa_txt_021 %> >></a>
				</td>
	    </tr> 
        <tr>
            <td><%
			    if cint(dsksOnOff) = 1 then
			    %>	
				<a href="javascript:popUp('stopur_2008.asp','1000','650','20','20')" class=rmenu><%=tsa_txt_080 %> >></a>
				<%end if%></td>
        </tr>
         <tr>
            <td>
			 <%if level <= 3 OR level = 6 then
				%>
				<a href="week_real_norm_2010.asp" class=rmenu><%=tsa_txt_356 %> >></a>
			    <%end if
				%></td>
        </tr>
		</table>
        </div>



      
	    
	    
	    
		
		<%
				
				
				
				lukafmjob = oRec("lukafmjob")
				kid = oRec("kid")
				
				
				
				end if
				oRec.close
				%>
			
                <%response.flush %>
                
                
                <%
            
                    '**** Valgt medarbejer *****'
                    strSQLmn = "SELECT mnavn, mnr, init, madr, mpostnr, mcity, mland FROM medarbejdere WHERE mid = "& usemrn 
                    oRec4.open strSQLmn, oConn, 3
                    if not oRec4.EOF then
                    
                    mMnavn = oRec4("mnavn")
                    mMnr = oRec4("mnr")
                    mInit = oRec4("init")
                    mAdr = oRec4("madr")
                    mPostnr = oRec4("mpostnr")
                    mCity = oRec4("mcity")
                    mLand = oRec4("mland")
                    
                    end if
                    oRec4.close
                 %>
                 
                 
                 
    
                <!-- Kontakt Kørsel info div -->
	            <div id="korseldiv" name="korseldiv" style="position:absolute; left:200px; top:500px; background-color:#ffffff; width:500px; border:2px #6CAE1C solid; padding:1px; visibility:hidden; display:none; z-index:2000000; height:500px; overflow:auto;">
				<div style="width:100%; border:0px #8caae6 solid; padding:20px;">
				<a href="javascript:popUp('kontaktpers.asp?func=opr&id=0&kid=<%=kid%>&rdir=treg','550','500','20','20')" target="_self" class=remnu>+ Tilføj ny kontakt</a>
				<table border=0 cellspacing=0 cellpadding=0 width=100%>
				<form action="#" method="post">
				<tr><td align=right><a href="#" id="jq_hideKpersdiv" onClick="hideKpersdiv()" class=red>[x]</a></td></tr>
				
				<tr>
				    <td><h3><%=tsa_txt_360 %>:</h3>
                    <%=tsa_txt_361 %><br />
                    <%=tsa_txt_365 %><br /><br />
				   
				     <input id="ko0chk" value="1" type="checkbox" /> <select id="ko0sel" style="font-size:9px;">
                            <option value=1><%=tsa_txt_362 %></option>
                            <option value=0>-</option>
                            <option value=2><%=tsa_txt_363 %></option>
                            <option value=3><%=tsa_txt_364 %></option>
                            </select>
                        <b><%=tsa_txt_367%></b>
                        <div id="ko0">
                        <%if len(trim(mAdr)) <> 0 then %>
                        <%=mMnavn %><br />
                        <%=mAdr%><br />
                        <%=mPostnr &" "& mCity%><br />
                        <%=mLand%> 
                        <%end if %>
                        </div>
				   </td>
				</tr>
				
				<tr>
				    <td>
                        <br /><input id="ko1chk" type="checkbox" />
                           <select id="ko1sel" style="font-size:9px;">
                            <option value=2><%=tsa_txt_363 %></option>
                            <option value=1><%=tsa_txt_362 %></option>
                            <option value=3><%=tsa_txt_364 %></option>
                            <option value=0>-</option>
                            </select>
                        
                         <b><%=tsa_txt_331 %></b>
                        <%
                        strSQLklicens = "SELECT k.kid, k.kkundenavn, k.adresse, k.postnr, k.city, k.land FROM kunder k WHERE useasfak = 1"
                        oRec.open strSQLklicens, oConn, 3
                        if not oRec.EOF then
                        
                        koAdr = oRec("adresse")
                        koPostnr = oRec("postnr")
                        koCity = oRec("city")
                        koLand = oRec("land")
                        koKid = oRec("kid")
                        kokkundenavn = oRec("kkundenavn")
                        
                        end if
                        oRec.close
                        %>
                        <br /><img src="../ill/blank.gif" width=1 height=5 /><br>
                        <div id="ko1" style="padding:5px; border:1px #cccccc solid; top:5px; background-color:#ffffe1;">
                        <%
                        if len(trim(koAdr)) <> 0 then %>
                        <%=kokkundenavn %><br />
                        <%=koAdr %> <br />
                        <%=koPostnr&" "&koCity %> <br />
                        <%=koLand%> <br />
                        <%end if %>
                        </div>
                        <br />
                        
                        <input id="ko1kid" type="hidden" value="k_<%=koKid%>" />
                        
                     
                        
				   </td>
				</tr>
				
				<tr><td align=right><input id="Button2" type="button" value=" Indlæs >> " onclick="koadr()" />
                    </tr>
				
				<tr><td valign=top>
                <br /><br />
				<h3><%=tsa_txt_368 %>:</h3>
                 <%=tsa_txt_361 %><br />
                    <%=tsa_txt_365 %><br /><br />

				<b><%=tsa_txt_022 %> (<%=lcase(tsa_txt_243) %>), <%=tsa_txt_024 %>:</b><br />
				
				<b><%=tsa_txt_078 %>:</b> 
                    <input id="FM_sog_kpers_dist" name="FM_sog_kpers_dist" value="" type="text" style="width:200px;" /> 
                    &nbsp;<input id="FM_sog_kpers_but" type="button" value="<%=tsa_txt_078 %> >> " />
                    <input id="FM_sog_kpers_dist_kid" value="<%=kid %>" type="hidden" />
                    <br />
                    <input id="FM_sog_kpers_dist_all" value="1" type="checkbox" /> <%=tsa_txt_359 %>
                    <!--<br /><input id="FM_sog_kpers_dist_more" value="1" type="checkbox" /> <=tsa_txt_372 %>-->
				
				<br /><br />
				<div id="filialerogkontakter">
				<!--jquery TXT -->
				</div>
				<br />&nbsp
				</td></tr>
				<tr><td align=right>
				<%call erkmDialog() 
				%>
				    <input id="koKmDialog" value="<%=kmDialogOnOff%>" type="hidden" />
                    <input id="koFlt" value="0" type="hidden" />
                    <input id="koFltx" value="1" type="hidden" />
                    
                    <input id="Button1" type="button" value=" Indlæs >> " onclick="koadr();" /><br />&nbsp;
                    </td></tr>
				</form>
				</table>
				</div>
				</div>
				<!--KM dialog div -->
				<%response.flush %>
				
				
				
				
				
				<!-- Kontakt info div -->
	            <div id="kpersdiv" name="kpersdiv" style="position:absolute; left:0px; top:0px; width:400px; border:2px #6CAE1C solid; padding:10px; visibility:hidden; display:none; background-color:#ffffff; z-index:2000000; height:300px; overflow:auto;">
				<table border=0 cellspacing=0 cellpadding=0 width=100%>
		        <tr><td align=right><a href="#" onClick="hideKpersdiv()" class=red>[x]</a></td></tr>
				<tr><td valign=top>
		        <b><%=tsa_txt_022 %>:</b><br />
				
				<%
				
				if kid <> 0 then
				kid = kid 
				else
				kid = 0
				end if
				
				'strSQL2 = "SELECT kp.navn, kp.email, kp.dirtlf, kp.mobiltlf, "_
				'&" kp.adresse As kpadr, kp.postnr As kpzip, kp.town As kptown, kp.land As kpland, kp.afdeling AS kpafd, "_
				'&" k.kkundenavn, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, k.land FROM kontaktpers kp, kunder k WHERE kp.kundeid = "& kid &" AND k.kid =" & kid
			    
			    strSQL2 = "SELECT kp.id, kp.navn, kp.email, kp.dirtlf, kp.mobiltlf, "_
				&" kp.adresse As kpadr, kp.postnr As kpzip, kp.town As kptown, kp.land As kpland, kp.afdeling AS kpafd, "_
				&" k.kid, k.kkundenavn, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, k.land FROM kunder k "_ 
				&" LEFT JOIN kontaktpers kp ON (kp.kundeid = k.kid)"_
				&" WHERE k.kid =" & kid
			    
			    oRec2.open strSQL2, oConn, 3
                x = 2
                while not oRec2.EOF
                
                        if x = 2 then%>
                         <br /><b><%=oRec2("kkundenavn") %></b> (<%=oRec2("kkundenr") %>)
                         <div id="Div4" style="padding:5px; border:1px #cccccc solid; top:5px; background-color:#DCF5BD;">
                        <%if len(trim(oRec2("adresse"))) <> 0 then %>
                        <%=oRec2("adresse") %> <br />
                        <%=oRec2("postnr") %>, <%=oRec2("city") %> <br />
                        <%=oRec2("land") %> <br />
                        <%end if %>
                       
                        
                            <%if len(trim(oRec2("telefon"))) <> 0 then %>
                            <%=tsa_txt_023 %>: <%=oRec2("telefon") %> <br />
                           
                            <%end if %>
                         
                          </div>
                            
                         <%x = x + 1 %>    
                         
                         <br /><br /><b><%=tsa_txt_024 %>:</b><br />
                         
                        <%end if %>
                        
                <%if len(trim(oRec2("id"))) <> 0 then %>
                <br /><b><%=oRec2("navn")%></b>
                        <div id="Div5" style="padding:5px; border:1px #cccccc solid; top:5px; background-color:#EFF3FF;">
                        <%if len(trim(oRec2("kpadr"))) <> 0 then %>
                        <%=oRec2("kpafd") %> <br />
                        <%=oRec2("kpadr") %> <br />
                        <%=oRec2("kpzip") %>, <%=oRec2("kptown") %> <br />
                        <%=oRec2("kpland") %> <br />
                        <%end if %>
                        
                
                <%if len(trim(oRec2("email"))) <> 0 then %>
                <%=tsa_txt_025 %>: <a href="mailto:<%=oRec2("email")%>" class=vmenu><%=oRec2("email")%></a><br>
                <%end if %>
                
                <%if len(trim(oRec2("dirtlf"))) <> 0 then %>
                <%=tsa_txt_026 %>: <%=oRec2("dirtlf")%><br>
                <%end if %>
                
                <%if len(trim(oRec2("mobiltlf"))) <> 0 then %>
                <%=tsa_txt_027 %>: <%=oRec2("mobiltlf")%><br><br />
                <%end if %>
                
                </div>
                
                <%end if %>
                
                <%
                x = x + 1
                oRec2.movenext
                wend
                oRec2.close
				%>
				
				<br />&nbsp
				</td></tr>
				
				</table>
				
				</div>
				
    
  
	
	
	<%
	
	end if '******* Denne ??
	'Response.flush
	
	
	call browsertype()
	if browstype_client = "mz" then
	topAdd_1 = 16
	else
	topAdd_1 = 0
	end if
	
	

   
	
     '*** SMILEY faneblade ****'
    
    if smilaktiv <> 0 AND showakt <> 0 AND cint(projektgrpOK) <> 0 then
    '*** Smiley *** 
    'sdTop = (333+topAdd)
     %>           
                
  <div id="s" style="position:relative; left:25px; top:122px; width:200px; visibility:<%=sVzb%>; display:<%=sDsp%>; z-index:2;">
   
        <table cellpadding=0 cellspacing=0 border=0>
        <tr>
            <td align=center id="s1_td" style="white-space:nowrap; width:100px; padding:1px; background-color:#EFF3FF; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" id="s1_k" class=rmenu>+ <%=tsa_txt_345 %></a></td>
             <td align=center id="s2_td" style="white-space:nowrap; width:120px; padding:1px; background-color:#EFF3FF; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" id="s2_k" class=rmenu>+ <%=tsa_txt_346 %></a></td>
			</td>
			
        </tr>
    </table>
    </div>
	    
		<%
	
	
	
	
	   
	    call medrabSmilord(usemrn)
	    
	    '**** Smiley vises om fredagen, første gang me logger på. ***'
	    if datepart("w", now, 2,2) = 5 AND datepart("h", now, 2,2) >= 13 AND session("smvist") <> "j" AND showakt = "1" then
    	smVzb = "visible"
    	smDsp = ""
    	session("smvist") = "j"
    	showweekmsg = "j"
    	else
    	smVzb = "hidden"
    	smDsp = "none"
    	showweekmsg = "n"
    	end if
    	
    	'showweekmsg = "j"
    	
	        if cint(smilaktiv) = 1 then%>
	        <div id="s1" style="position:absolute; left:105px; top:395px; width:250px; visibility:<%=smVzb%>; display:<%=smDsp%>; z-index:20000000; background-color:#FFFFFF; padding:2px; border:1px #999999 solid;">
	       <%
	        call showsmiley()
	        
	        call afslutuge()
	        
	        %><br />&nbsp;
	         </div>
	        <div id="s2" style="position:absolute; left:205px; top:395px; width:475px; visibility:hidden; display:none; z-index:20000000; background-color:#FFFFFF; padding:2px; border:1px #999999 solid;">
	        <%
	        useYear = year(tjekdag(4))
	        call smileystatus(usemrn, 1)
	        %>
	        <br />&nbsp;
	        </div>
        	
	        <%end if
    	
	    end if

	
	


	
	if showakt <> 0 AND cint(projektgrpOK) <> 0 then
	
         
    
    tTop = 123 + topAdd_1 '(225+topAdd)
	tLeft = 20
	tWdth = 745
	tId = "timereg"
	
	call tableDivWid(tTop,tLeft,tWdth,tId, dTimVzb, dTimDsp)


    


	
	%>
	

	
	<table cellspacing="0" cellpadding="2" border="0" width="100%" bgcolor="#8caae6"><!-- id="inputTable" onkeydown="doKeyDown()" -->
	<tr bgcolor="#ffffff">
	    <td colspan=7 valign=top>
	    <%
	    oskrift = tsa_txt_031 &" "& datepart("ww", tjekdag(1), 2 ,2) & " "& datepart("yyyy", tjekdag(1), 2 ,3) 
	    %>
	<b><%=oskrift %></b> <a href="#" name="anc_s0" id="pagesetadv_but" class="rmenu">(+ <%=tsa_txt_302%>)</a>
	<br />
	<%if cint(intEasyreg) = 1 then%>
	<span><h3>Easyreg. listen 
	    <%if level = 1 then %>
	    <a href="javascript:popUp('easyreg_2010.asp','800','720','20','20')" class=rmenu>(+ <%=lcase(tsa_txt_251)%>)</a>
	    <%end if %>
	</h3>  </span>
	<%else %>

	<!--
    <h3><=j_jobnavn_nr %>  <a href="#" id="showjobbesk" class=rmenu>(+ <=tsa_txt_029 %>: <=len(trim(jobbeskrivelse)) &" "& tsa_txt_371 &"." %>)</a></h3>
    -->
	
	<%end if %>
	</td>
	<td colspan=2 valign=top align=right>
	                        <!-- skift uge pile --->
	                        <table cellpadding=0 cellspacing=0 border=0 width=80>
	                        <tr>
	                        <td valign=top align=right><a href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=1&strdag=<%=day(useDatePrevWeek)%>&strmrd=<%=month(useDatePrevWeek)%>&straar=<%=year(useDatePrevWeek)%>&fromsdsk=<%=fromsdsk%>"><img src="../ill/nav_left_blue.png" border="0" /></a></td>
                           <td style="pading-left:20px;" valign=top align=right><a href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=1&strdag=<%=day(useDateNextWeek)%>&strmrd=<%=month(useDateNextWeek)%>&straar=<%=year(useDateNextWeek)%>&fromsdsk=<%=fromsdsk%>"><img src="../ill/nav_right_blue.png" border="0" /></a></td>
	                        </tr>
	                        </table>
	</td>
	</tr>
	
	
	
	
    
    <!--<input id="showjobinfo" name="showjobinfo" type="hidden" value="<=showjobinfo %>" />-->
	
    
		
		<tr bgcolor="#FFFFFF">
        <form action="timereg_akt_2006.asp?FM_use_me=<%=usemrn%>&func=db&fromsdsk=<%=fromsdsk%>" method="POST" name="timereg">
		<td colspan=7 valign=top>
		
		 <!-- luk job ved indtatning --->
        <%
	            lukafm = 0
	            strSQL = "SELECT lukafm FROM licens WHERE id = 1"
	            oRec.open strSQL, oConn, 3
	            while not oRec.EOF
	            lukafm = oRec("lukafm")
	            oRec.movenext
	            wend
	
	            oRec.close
	
	     %>
                    <!--<div style="position:relative; border:4px #cccccc solid; padding:5px; width:550px;">-->
                    <table cellpadding=1 cellspacing=0 border=0 width=100%><%
	
	                if cint(lukafm) = 1 then%>
	
		
		                 <%select case lukafmjob
		                 case 1
		                 %>
		                    <tr><td colspan=2>
		                    <input type="checkbox" name="FM_lukjob" id="FM_lukjob" value="1"> <span style="color:#000000;">
		                    <%if lto = "bowe" then%>
		                    <%=tsa_txt_040 %>
		                    <%else%>
		                    <%=tsa_txt_041 %>
		                    <%end if
		    
		                    %></span></td></tr><%
		                 case 2
		                 %>
		    
		
		                 <tr><td>
		   
		                    <input type="checkbox" name="FM_lukjob" id="FM_lukjob" value="1"> </td><td> <span style="color:#000000;">
		                    <%if lto = "bowe" then%>
		                    <%=tsa_txt_043 %>
		                    <%else%>
		                    <%=tsa_txt_044 %>
		                    <%end if
		    
		                    %> </span>
		                    <input type="hidden" name="FM_kopierjob" id="FM_kopierjob" value="1">
		    
		                  </td></tr>
		                <%
		                 end select
        
        
                        end if

                        if cint(useasfak) = 1 then '*** Kun interne (licens ejer)
                        %>
		
		                 <tr><td colspan=2>
		                <input id="tildeliheledage" name="tildeliheledage" type="checkbox" CHECKED value="1"/> <span style="color:#000000;"><%=tsa_txt_277 %> </span><br />&nbsp;
		                 </td></tr>
		
		                <%'*** Min indtast ETS **'
		
		                if lto = "ets-track" OR lto = "intranet - local" then %>
		                <tr><td colspan=2><input id="minindtast_on" name="minindtast_on" type="checkbox" value="1" CHECKED /> <span style="color:#000000;">Gennemtving min. indtastning 7,4 timer / 8,0 timer</span>
		                 </td></tr>
		                <%end if %>

                        <%end if 'licensejer%>
		
		                <%
		                '*** Multi tildel ***'
		                'if level = 1 then
		                %>
		                <tr><td colspan=2>
                        <input id="tildelalle" name="tildelalle" type="checkbox" value="1" onclick="showmultiDiv()" /> <span style="color:#000000;"><%=tsa_txt_268 %> (<%=tsa_txt_357 %>)</span> <br />
		
		                                <div id="multivmed" style="position:relative; left:0px; top:0px; padding:10px; visibility:hidden; display:none;">
		                                <%=tsa_txt_267 %>:<br /> <select name="tildelselmedarb" id="tildelselmedarb" size=10 multiple style="width:350px;">
				                                <%
					                                strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe FROM medarbejdere WHERE mansat <> 2 "& medarbgrpIdSQLkri &" ORDER BY Mnavn"
					                                oRec.open strSQL, oConn, 3
					
					                                while not oRec.EOF 
					
					                                if cint(oRec("Mid")) = cint(usemrn) then
					                                rchk = "SELECTED"
					                                else
					                                rchk = ""
					                                end if%>
					                                <option value="<%=oRec("Mid")%>" <%=rchk%>><%=oRec("mnavn")%></option>
					                                <%
					
					
					                                oRec.movenext
					                                wend
					                                oRec.close
				                                %></select>
				                                </div>
		
		
		
		                                <%
		                                'end if
		                                %>
		
		                                </td></tr>
                                        </table>

                                        <!--</div>-->
		                                <br />&nbsp;

        

		
		
		
		                </td>
		                <td colspan=2 align=right valign=bottom style="padding-bottom:5px;"><input type="submit" id="sm1" value="<%=tsa_txt_085 %> >> "></td>
	                </tr>
	
	                 <%
                     if cint(intEasyreg) = 1 then
     
                     %>
                     <tr bgcolor="#FFFFFF"><td colspan=9 style="height:5px;"><img src="ill/blank.gif" border=0 /></td></tr>
                     <tr bgcolor="#DCF5BD">
                        <td colspan=2 style="border-right:1px #CCCCCC solid; border-left:1px #CCCCCC solid; border-top:1px #CCCCCC solid;">Fordel timer på Easyreg. aktiviteter</td>
        
                        <%for e = 1 to 7 %>
                        <td style="border-right:1px #CCCCCC solid; border-top:1px #CCCCCC solid;">
                            <input id="ea_<%=e %>" type="text" style="width:46px; font-size:9px;" /><a href="#" id="ea_k_<%=e%>" class="ea_kom"><font color="#999999"> + </font></a></td>
                        <%next %>
     
     
                     </tr>
                     <tr bgcolor="#DCF5BD">
                     <td colspan=2 style="border-right:1px #CCCCCC solid; border-left:1px #CCCCCC solid; border-top:1px #CCCCCC solid;">(0 = nulstil)</td>
                     <%for e = 1 to 7 %>
                            <td style="border-right:1px #CCCCCC solid; border-top:1px #CCCCCC solid;">
                                <input id="ea_kn_<%=e %>" type="button" value="Calc. =" style="font-family:arial; font-size:9px;" /></td>
                        <%next %>
     

                     </tr>
                     <%
                     end if
                     %>
     

    

      </table>
     </div><!--tableDiv-->
     

     <div id="timereg_job" style="position:relative; width:780px; border:0px; top:123px; visibility:visible; left:20px; display:; z-index:1000">
    
     <%
    'tTop = 153
	'tLeft = 20
	'tWdth = 0 '760
	'tId = "timereg_job"
	
	'call tableDivWid(tTop,tLeft,tWdth,tId, dTimVzb, dTimDsp)
    
     %>
   
    
    
    <table id="incidentlist" cellspacing=0 cellpadding=0 border=0 width=100%>
   <tbody>

  
   <tr>
	<td>

	<input type="hidden" name="jobid" value="<%=jobid%>">
	<input type="Hidden" name="FM_vistimereltid" value="<%=visTimerElTid%>">
	<input type="Hidden" name="FM_start_dag" value="<%=strdag%>">
	<input type="Hidden" name="FM_start_mrd" value="<%=strmrd%>">
	<input type="Hidden" name="FM_start_aar" value="<%=straar%>">
	<!--<input type="Hidden" name="searchstring" value="<=searchstring%>">-->
	<input type="hidden" name="FM_mnr" value="<%=usemrn%>">

	<input type="hidden" name="year" value="<%=strRegAar%>">
	<input type="hidden" id="datoMan" name="datoMan" value="<%=tjekdag(1)%>">
	<input type="hidden" name="datoTir" value="<%=tjekdag(2)%>">
	<input type="hidden" name="datoOns" value="<%=tjekdag(3)%>">
	<input type="hidden" name="datoTor" value="<%=tjekdag(4)%>">
	<input type="hidden" name="datoFre" value="<%=tjekdag(5)%>">
	<input type="hidden" name="datoLor" value="<%=tjekdag(6)%>">
	<input type="hidden" id="datoSon" name="datoSon" value="<%=tjekdag(7)%>">
	<input type="hidden" id="lastopendiv" name="lastopendiv" value="jobinfo">

    <input type="hidden" id="FM_session_user" name="FM_session_user" value="<%=session("user")%>">
    <input type="hidden" id="FM_now" name="FM_now" value="<%=formatdatetime(now, 2)%>">

    </td></tr>

   

	
    <%

	
	totTimMan = 0
	totTimTir = 0
	totTimOns = 0
	totTimTor = 0
	totTimFre = 0
	totTimLor = 0
	totTimSon = 0
	
	'*** Brugergrupper er fundet i linie 2697 ***
	'call hentbgrppamedarb(usemrn)
	
	iRowLoop = 1
	m = 58 '16
    if cint(intEasyreg) = 1 then
    x = 3500 ' Bør kun være 750, men der kan forekomme mere end 100 Easyreg. aktiviteter indtil loft effketureres
    else
	x = 350 '6500 '2500 (maks 100 aktiviteter * 7 dages timereg.)
    end if
			
	dim aktdata
	redim aktdata(x, m) 
	
					
					
					
    '*** Aktiviteter og Timer SQL MAIN **** 
	strSQL = "SELECT a.navn, a.id AS aid, a.fakturerbar, t.timer, t.tdato, t.timerkom, t.offentlig, "_
	&" t.sttid, t.sltid, j.jobnr AS jnr, j.id AS jid, j.jobnavn AS jnavn, j.risiko, job_internbesk, "_
    &" j.serviceaft, j.beskrivelse AS jobbesk, j.jobans1, jobans2, jobstartdato, jobslutdato, j.fastpris, jo_bruttooms, j.rekvnr, j.kommentar, "_
	&" a.incidentid, a.aktstartdato, a.aktslutdato, a.beskrivelse, t.godkendtstatus, "_
	&" a.tidslaas, a.tidslaas_st, a.tidslaas_sl, a.tidslaas_man, a.tidslaas_tir, a.aktbudget, a.budgettimer AS aktbudgettimer, a.bgr, a.antalstk, "_
	&" a.tidslaas_ons, a.tidslaas_tor, a.tidslaas_fre, a.tidslaas_lor, a.tidslaas_son, t.bopal, t.destination, a.fase, j.jobstatus, "_
    &"tu.forvalgt_sortorder, tu.id AS tuid, tu.forvalgt_af, tu.forvalgt_dt "
	
	'if cint(intEasyreg) = 1 then
	strSQL = strSQL &", kkundenavn, kkundenr, kid, kundeans1 "
	'end if
	
	strSQL = strSQL &" FROM job j, aktiviteter a "
	
	if cint(intEasyreg) = 1 then
	strSQL = strSQL &", timereg_usejob AS tu "
	else
    strSQL = strSQL &" LEFT JOIN timereg_usejob AS tu ON (tu.jobid = j.id AND tu.medarb = "& usemrn &") "
    end if
	
    strSQL = strSQL &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "

	strSQL = strSQL &" LEFT JOIN timer t ON"_
	&" (t.taktivitetid = a.id "_ 
    &" AND t.tmnr = "& usemrn &" AND (t.tdato = '"& varTjDatoUS_son &"'"_
	&" OR t.tdato = '"& varTjDatoUS_man &"'"_
	&" OR t.tdato = '"& varTjDatoUS_tir &"'"_
	&" OR t.tdato = '"& varTjDatoUS_ons &"'"_
	&" OR t.tdato = '"& varTjDatoUS_tor &"'"_
	&" OR t.tdato = '"& varTjDatoUS_fre &"'"_
	&" OR t.tdato = '"& varTjDatoUS_lor &"'))"
	
   
	    '*** Vis som easyreg. 
	    if cint(intEasyreg) = 1 then
            
            if cint(sogKontakter) <> 0 then
            strSQL = strSQL &" WHERE (k.kid = " & sogKontakter
            else
            strSQL = strSQL &" WHERE (k.kid <> 0 "
            end if

            strSQL = strSQL &" AND j.id <> 0 AND a.easyreg = 1 AND j.jobstatus = 1 "_ 
            &" AND tu.medarb = "& usemrn &" AND tu.jobid = j.id AND tu.easyreg <> 0) AND "& strSQLkri3

        
        else
            strSQL = strSQL &" WHERE ("& seljobidSQL &") AND (j.jobstartdato <= '"& varTjDatoUS_son &"') AND "& strSQLkri3
        end if
	
	
    
	
	'Response.Write  "strSQLkri3" &  strSQLkri3  
	
	if cint(ignJobogAktper) = 1 then
	useDateSQL = year(useDate) &"/"& month(useDate) &"/"& day(useDate)
	strSQL = StrSQL & " AND (a.aktstartdato <= '"& varTjDatoUS_son &"' AND a.aktslutdato >= '"& useDateSQL &"')"
	end if
	
    '*** Vis ikke disse typer på treg. siden **'
    strSQL = StrSQL & " AND ("& aty_sql_hide_on_treg &")"
    
	if lto = "Q2con" then
	strSQL = StrSQL & " ORDER BY j.jobnavn, j.id, LTRIM(a.navn)"
	else
	    if cint(intEasyreg) = 1 then
	    strSQL = StrSQL & " ORDER BY j.jobnavn, j.id, a.fase, a.sortorder, LTRIM(a.navn) " 'kkundenavn, 
	    else
	    strSQL = StrSQL & " ORDER BY tu.forvalgt_sortorder, j.risiko, j.jobnavn, j.id, a.fase, a.sortorder, LTRIM(a.navn) LIMIT 550"
        'j.risiko, j.jobnavn, j.id, a.fase, a.sortorder, LTRIM(a.navn) LIMIT 550"
	    end if
	end if
	
	
    'Response.write strSQL
	'Response.flush
	
	tr_vzb = ""
    tr_dsp = ""
	lastFase = ""
    lastJobId = 0
    'tjkLastJidFak = 0
	
    antalJobLinier = 0

	oRec.open strSQL, oConn, 3
	while not oRec.EOF
            

                   
	
	aktdata(iRowLoop, 0) = oRec("tdato")
	aktdata(iRowLoop, 1) = oRec("timer")
	aktdata(iRowLoop, 2) = oRec("navn")
	aktdata(iRowLoop, 3) = oRec("aid")
	aktdata(iRowLoop, 4) = oRec("jid")
	aktdata(iRowLoop, 5) = oRec("jnavn")
	aktdata(iRowLoop, 6) = oRec("timerkom")
	aktdata(iRowLoop, 7) = oRec("offentlig")
	aktdata(iRowLoop, 11) = oRec("fakturerbar")
	aktdata(iRowLoop, 12) = oRec("incidentid")
	aktdata(iRowLoop, 13) = oRec("aktstartdato")
	aktdata(iRowLoop, 14) = oRec("aktslutdato")
	aktdata(iRowLoop, 15) = oRec("beskrivelse")
	aktdata(iRowLoop, 16) = oRec("godkendtstatus")
	
	aktdata(iRowLoop, 17) = oRec("tidslaas")
	aktdata(iRowLoop, 18) = oRec("tidslaas_st")
	aktdata(iRowLoop, 19) = oRec("tidslaas_sl")
	aktdata(iRowLoop, 20) = oRec("tidslaas_man")
	aktdata(iRowLoop, 21) = oRec("tidslaas_tir")
	aktdata(iRowLoop, 22) = oRec("tidslaas_ons")
	aktdata(iRowLoop, 23) = oRec("tidslaas_tor")
	aktdata(iRowLoop, 24) = oRec("tidslaas_fre")
	aktdata(iRowLoop, 25) = oRec("tidslaas_lor")
	aktdata(iRowLoop, 26) = oRec("tidslaas_son")
	aktdata(iRowLoop, 27) = oRec("bopal")
	aktdata(iRowLoop, 28) = oRec("destination")
	aktdata(iRowLoop, 29) = oRec("fase")
	aktdata(iRowLoop, 30) = oRec("aktbudget")
	aktdata(iRowLoop, 31) = oRec("jnr")

    aktdata(iRowLoop, 33) = oRec("aktbudgettimer")
	aktdata(iRowLoop, 34) = oRec("bgr")
    aktdata(iRowLoop, 35) = oRec("antalstk")

    aktdata(iRowLoop, 40) = oRec("job_internbesk")
    aktdata(iRowLoop, 41) = oRec("jobbesk") 


   
	aktdata(iRowLoop, 32) = oRec("kkundenavn")
    aktdata(iRowLoop, 56) = "("& oRec("kkundenr") &")"

	aktdata(iRowLoop, 43) = oRec("kundeans1")
    aktdata(iRowLoop, 44) = oRec("jobans1")
    aktdata(iRowLoop, 50) = oRec("jobans2")

    aktdata(iRowLoop, 42) = oRec("kid")


    aktdata(iRowLoop, 45) = oRec("jobstartdato")
    aktdata(iRowLoop, 46) = oRec("jobslutdato")

    aktdata(iRowLoop, 47) = oRec("fastpris")
    aktdata(iRowLoop, 48) = oRec("jo_bruttooms")

    aktdata(iRowLoop, 49) = oRec("rekvnr")

    aktdata(iRowLoop, 51) = oRec("serviceaft")
    
    aktdata(iRowLoop, 52) = oRec("forvalgt_sortorder")	       
    aktdata(iRowLoop, 53) = oRec("tuid")

    aktdata(iRowLoop, 54) = oRec("forvalgt_af")
    aktdata(iRowLoop, 55) = oRec("forvalgt_dt")
    

    aktdata(iRowLoop, 58) = oRec("kommentar")

    select case aktdata(iRowLoop, 34) 
    case 1 
    aktdata(iRowLoop, 36) = "Fork.: "& formatnumber(aktdata(iRowLoop, 33), 2)  & " t."
    case 2
    aktdata(iRowLoop, 36) = "Ant.: "& formatnumber(aktdata(iRowLoop, 35), 2)  & " stk."
    case else
    aktdata(iRowLoop, 36) = ""
    end select

    '** Timer real + kost (lsået fra) **'
    aktdata(iRowLoop, 37) = 0
    aktdata(iRowLoop, 39) = 0
    strSQLttot = "SELECT SUM(timer) AS timertot FROM timer WHERE taktivitetid = "& aktdata(iRowLoop, 3) & " GROUP BY taktivitetid"
    oRec2.open strSQLttot, oConn, 3
    if not oRec2.EOF then  
    aktdata(iRowLoop, 37) = oRec2("timertot")
        
        ''if lto = "intranet - local" OR lto = "epi" then
        ', SUM(timer*kostpris*(kurs/100)) AS internKost
        'aktdata(iRowLoop, 39) = oRec2("internkost") 
        ''end if

    end if
    oRec2.close

     '** Timer real kun valgte medarb **'
    aktdata(iRowLoop, 38) = 0
    strSQLtmnr = "SELECT SUM(timer) AS timermnr FROM timer WHERE taktivitetid = "& aktdata(iRowLoop, 3) & " AND tmnr = "& usemrn &" GROUP BY taktivitetid"
    oRec2.open strSQLtmnr, oConn, 3
    if not oRec2.EOF then  
    aktdata(iRowLoop, 38) = oRec2("timermnr")
    end if
    oRec2.close

	

	
	if len(trim(oRec("sttid"))) <> 0 then
		if left(formatdatetime(oRec("sttid"), 3), 5) <> "00:00" then
		aktdata(iRowLoop, 8) = left(formatdatetime(oRec("sttid"), 3), 5)
		else
		aktdata(iRowLoop, 8) = ""
		end if
	else
	aktdata(iRowLoop, 8) = ""
	end if
	
	if len(trim(oRec("sltid"))) <> 0 then
		if left(formatdatetime(oRec("sltid"), 3), 5) <> "00:00" then
		aktdata(iRowLoop, 9) = left(formatdatetime(oRec("sltid"), 3), 5)
		else
		aktdata(iRowLoop, 9) = ""
		end if
	else
	aktdata(iRowLoop, 9) = ""
	end if

	
	iRowLoop = iRowLoop + 1
	oRec.movenext
	wend
	oRec.close
	
	'** If iRowLoop > 100 show alert **'


    '**
	
	
	if cint(iRowLoop) <> 1 then
	
	dim flt
	Redim flt(7)
	
    loops = 0
	antalaktlinier = 0
	lastAktid = 0
	for iRowLoop = 1 TO UBOUND(aktdata)
		
		    loops = loops + 1
		    
		    
		    if lastAktid <> aktdata(iRowLoop, 3) AND len(trim(aktdata(iRowLoop, 3))) <> 0 then
		
			if iRowLoop <> 0 then
			
					
					for x = 1 to 7
					
					Response.write flt(x) 
					
					next
					antalaktlinier = antalaktlinier + 1 
					%>
					
			</tr>
			<%end if%>
			
			
			<%
            '*** Job ***'
            if lastJobid <> aktdata(iRowLoop, 4) AND cint(intEasyreg) <> 1 then

                     
                   
            if antalJobLinier <> 0 then
            %>

            <tr>
            <td bgcolor="#FFFFFF" colspan=9 align=right valign=bottom style="padding:5px 5px 5px 5px;"><input type="submit" id="Submit1" value="<%=tsa_txt_085 %> >> "></td>
            </tr>
            <!--
             <tr>
                  <td colspan=9 style="background-color:#D6dff5; background-image:url('../ill/stripe_10.png'); height:2px;">
                    <img src="../ill/blank.gif" width=1 height=1 /></td>  </tr>
                    -->
            </table>
            </div>

            </td></tr>

            <%else %>

             <tr>
               
               <td>
                   <input type="hidden" name="SortOrder" value="0" />
	            <input type="hidden" name="rowId" value="0" />
        

          

            <div id="Div1" style="position:relative; width:745px; border:0px; visibility:visible; display:; z-index:1000">
            <table width=100% cellspacing=0 cellpadding=0 border=0>
            <tr><td><img src="../ill/blank.gif" width=1 height=1 border=0 /></td></tr>
            </table>
            </div>
            </td></tr>

            
            <%end if %>

            
              

               <tr>
               
               <td>
                   <input type="hidden" name="SortOrder" value="<%=aktdata(iRowLoop, 52)%>" />
	            <input type="hidden" name="rowId" value="<%=aktdata(iRowLoop, 53)%>" />
        

          

            <div id="ad_timereg_<%=aktdata(iRowLoop, 4)%>" style="position:relative; width:745px; border:0px; visibility:visible; display:; z-index:1000">
            <table width=100% cellspacing=0 cellpadding=0 border=0 bgcolor="#ffffff">
             
             <!--
            <tr><td style="background-color:#D6Dff5; border-bottom:0px #cccccc solid; padding:10px 0px 0px 0px;" colspan=2 class=lille>
             < <%=antalJobLinier %> > 
             
             <%if cint(usemrn) <> cint(aktdata(iRowLoop, 54)) then %>
             &nbsp;<span style="font-size:9px; color:#999999;"><i>tildelt d. <%=aktdata(iRowLoop, 55)%> af 
             <%call meStamdata(aktdata(iRowLoop, 54)) %>
             
             <%=meTxt %>
             </i></span>
             <%end if %>
            <img src="../ill/blank.gif" width=1 height=10 border=0 />
            </td>
            
            
            <td colspan=2 style="background-color:#D6Dff5; border-bottom:0px #cccccc solid; padding:10px 5px 0px 0px;" valign="bottom" align=right class=lille>&nbsp;</td>
            </tr> 
            -->    
            
                 <tr><td style="background-color:#D6Dff5; border-bottom:0px #cccccc solid; padding:10px 0px 0px 0px;" colspan=5 class=lille><img src="../ill/blank.gif" width=1 height=5 border=0 /></td>
                 </tr>
                <tr>

                    
                            
                    <td style="width:240px; background-color:#5C75AA; border-bottom:1px #8caae6 solid; border-top:1px #FFFFFF solid; padding:3px 3px 0px 4px;" valign=top> <a class="a_treg" id="a_timereg_<%=aktdata(iRowLoop, 4)%>" href="#">+ <%=left(aktdata(iRowLoop, 5), 35) %> (<%=aktdata(iRowLoop, 31) %>)</a>
                    <br /><span style="color:#cccccc; font-size:9px;"><%=left(aktdata(iRowLoop, 32), 35) & " " & aktdata(iRowLoop, 56) %></span></td>

                    <td style="width:15px; background-color:#5C75AA; border:0px #cccccc solid; border-bottom:1px #8caae6 solid; border-top:1px #FFFFFF solid; border-right:0px #cccccc solid; padding:3px 3px 0px 4px;" valign=top><a class="a_det" id="a_det_<%=aktdata(iRowLoop, 4)%>" href="#"><img src="../ill/document.png" border="0" alt="Stamdata, job beskrivelse & Materiale registrering"/></a></td>
                    
                    <td valign="top" style="width:155px; background-color:#FFFFFF; border-bottom:1px #8caae6 solid; padding:2px 2px 0px 1px;">
                    <%'if cint(usemrn) <> cint(aktdata(iRowLoop, 54)) then %>
                    &nbsp;<span style="font-size:10px; color:#999999;"> <%=antalJobLinier %> // <i>Af  
                    <%call meStamdata(aktdata(iRowLoop, 54)) %>
                    <a href="mailto:<%=meEmail %>?subject=Vedr. job <%=left(aktdata(iRowLoop, 5), 35) %> (<%=aktdata(iRowLoop, 31) %>)" class="rmenu"><%=meInit %></a> d. <%=aktdata(iRowLoop, 55)%> 
                    </i></span>
                    <%'end if %>
                    &nbsp;
                    </td>
                    
                    

                    <td valign="bottom" class=lille style="background-color:#FFFFFF; width:300px; border-bottom:1px #8caae6 solid; border-top:0px #CCCCCC solid; height:4px; padding:2px 2px 0px 30px;">
                    <input type="text" style="width:258px; font-size:9px; font-family:Arial; color:#999999; font-style:italic;" maxlength="101" value="Job tweet..(åben for alle)" class="FM_job_komm" name="FM_job_komm_<%=aktdata(iRowLoop, 4)%>" id="FM_job_komm_<%=aktdata(iRowLoop, 4)%>"> <a href="#" class="aa_job_komm" id="aa_job_komm_<%=aktdata(iRowLoop, 4)%>">Gem</a>
                    <div id="dv_job_komm_<%=aktdata(iRowLoop, 4)%>" style="width:258px; font-size:9px; font-family:Arial; color:#000000; font-style:italic; overflow-y:auto; overflow-x:hidden; height:30px;"><%=aktdata(iRowLoop, 58) %></div>
                    </td>
                    <td valign="top" align=right class=lille style="background-color:#FFFFFF; width:40px; border-bottom:1px #8caae6 solid; border-top:0px #cccccc solid; border-right:2px #cccccc solid; padding:2px 2px 0px 10px;">
                    <span style="color:#999999; font-size:9px; font-style:oblique;">Fjern<br /><input class="fjern_job" type="checkbox" value="1" id="fjern_job_<%=aktdata(iRowLoop, 4)%>"/></span></td>
                </tr>
               
            </table>
            </div>
            <%response.flush %>



             <!-- job detaljer / stamdata -->
        

             <div id="div_det_<%=aktdata(iRowLoop, 4)%>"" style="position:relative; background-color:#ffffff; padding:0px 3px 3px 3px; width:745px; border:1px #8caae6 solid; visibility:hidden; display:none; z-index:2000">
   
            <table cellspacing="0" cellpadding="2" border="0" width="100%" bgcolor="#8caae6">
            

            <tr>
                  <td colspan=9 style="background-color:#D6dff5; background-image:url('../ill/stripe_10_graa.png'); height:2px;">
                    <img src="../ill/blank.gif" width=1 height=1 /></td>  </tr>

                    <tr>
                  <td colspan=9 style="height:20px;" bgcolor="#FFFFFF">
                    <img src="../ill/blank.gif" width=1 height=1 /></td>  </tr>



                                                 <tr bgcolor="#FFFFFF"><td valign=top class=lille colspan=2 style="padding:5px 5px 5px 10px; width:250px;">


                                                <table cellpadding=0 cellspacing=0 border=0>
                                                <tr>
                                                    <%
                                                    'call akttyper2009(2)
            
                                                    for dt = 0 TO 3 
                                                     if dt <> 0 then
                                                     dtnowHigh = dateadd("m", 1, dtnowHigh)
                                                     else
                                                     dtnowHigh = dateadd("m", -3, now)
                                                     end if

            

                                                    %>
                                                    <td align=center style="font-size:7px;"><%=left(monthname(month(dtnowHigh)), 3) %></td>
                                                    <%next %>
                                                 </tr>

                                                 <tr bgcolor="#FFFFFF">
                                                    <%for dt = 0 TO 3
                                                     if dt <> 0 then
                                                     dtnowHigh = dateadd("m", 1, dtnowHigh)
                                                     else
                                                     dtnowHigh = dateadd("m", -3, now)
                                                     end if
             
             
                                                     select case month(dtnowHigh)
                                                     case 1,3,5,7,8,10,12
                                                        eday = 31
                                                        case 2
                                                            select case right(year(dtnowHigh), 2)
                                                            case "00", "04", "08", "12", "16", "20", "24", "28", "30", "34", "38", "42", "46"
                                                            eday = 29
                                                            case else
                                                            eday = 28
                                                            end select
                                                     case else
                                                        eday = 30
                                                     end select

                                                     dtnowHighSQL = year(dtnowHigh) &"/"& month(dtnowHigh) &"/"& eday
                                                     dtnowLowSQL = year(dtnowHigh) &"/"& month(dtnowHigh) &"/1"

                                                       hgt = 0
                                                       timerThis = 0

                                                     strSQLjl = "SELECT sum(timer) AS timer FROM timer WHERE tjobnr = "& aktdata(iRowLoop, 31) &" AND tdato BETWEEN '"& dtnowLowSQL &"' AND '"& dtnowHighSQL &"'"_
                                                     &" AND ("& aty_sql_realhours &") GROUP BY tjobnr"

                                                     'Response.Write strSQLjl
                                                     'Response.flush

                                                     oRec2.open strSQLjl, oConn, 3
                                                     if not oRec2.EOF then

                                                     hgt = formatnumber(oRec2("timer"), 0)
                                                     timerThis = oRec2("timer")

                                                     end if
                                                     oRec2.close

                                                     if hgt > 100 then
                                                     hgt = 20
                                                     else
                                                     hgt = hgt/5 
                                                     end if
                                                    %>
                                                    <td class=lille align=right height=20 valign=bottom style="border-right:1px #CCCCCC solid;">
                                                    <%if hgt <> 0 then 
                                                    bdThis = 1
                                                    timerThis = formatnumber(timerThis,2)
                                                    else
                                                    timerThis = ""
                                                    bdThis = 0
            
                                                    end if 
            
                                                    bgThis = "#D6dff5"

                                                    if hgt = 0 then
                                                    bgThis = "#ffffff"
                                                    end if

           
          
                                                    %>
                                                    <div style="height:<%=hgt%>px; width:20px; border-top:<%=bdThis%>px #CCCCCC solid; padding:0px; background-color:<%=bgThis%>;"><img src="../ill/blank.gif" width="20" height="<%=hgt%>" border="0" alt="<%=timerThis %>" /></div>
            
                                                    </td>
                                                    <%next %>
                                                 </tr>
	    
                                                </table>
                                                <br />
                                                <b>Stamdata:</b><br />
                                                Start & slutdato: <%=formatdatetime(aktdata(iRowLoop, 45), 2) %> - <%=formatdatetime(aktdata(iRowLoop, 46), 2) %>	


                                                <br />

                                               <%call meStamdata(aktdata(iRowLoop, 43)) %>
                                               <%=tsa_txt_007 %>: <%=meTxt %>

                                               <br />
                                               <%call meStamdata(aktdata(iRowLoop, 44)) %>
                                               <%=tsa_txt_010 %>: <%=meTxt %>

                                                  <br />
                                               <%call meStamdata(aktdata(iRowLoop, 50)) %>
                                               <%=tsa_txt_011 %>: <%=meTxt %>

      
		
                                                <%if level <= 2 OR level = 6 then %>	

		                                                <%select case aktdata(iRowLoop, 47) '** Fastpris
		                                                case 1
		                                                jobtype = tsa_txt_013 '"Fastpris"
				                                                'if oRec("usejoborakt_tp") <> 0 then
				                                                'jobtype = jobtype & " ("& lcase(tsa_txt_125) &")"
				                                                'end if
		                                                case else
		                                                jobtype = tsa_txt_014 '"Lbn. timer"
		                                                end select
					
		
		                                        %>
		                                        <br /><%=tsa_txt_015 %>: <%=jobtype%>		
		                                         <%end if %>		
		 
         
                                                 <%
             
			

				

						
						                                        '*** Aftale ***
						                                        if aktdata(iRowLoop, 51) <> 0 then

                                                                %>
                                                                <br /><%=tsa_txt_018 %>:

						                                        <%
                                                                strSQLaft = "SELECT navn, aftalenr, status FROM serviceaft WHERE id = " & aktdata(iRowLoop, 51)
						                                        oRec5.open strSQLaft, oConn, 3
                                                                if not oRec5.EOF then
							
							                                        %>
                            
							                                        <%=oRec5("navn")%> (<%=oRec5("aftalenr")%>)
							
							                                        <%
							                                        select case oRec5("status")
							                                        case 1
							                                        %>
							                                         - <i><%=tsa_txt_019 %></i>
							                                        <%case else
							                                        %> 
							                                         - <i> <%=tsa_txt_020 %></i>
							                                        <%end select
							
						
						                                        end if
						                                        oRec5.close

                                                                end if
                        
                                                                %>
                  
				
				                                        <%if len(trim(rekvnr)) <> 0 then %>
				                                        <br /><%=tsa_txt_269 %>: <%=aktdata(iRowLoop, 49) %>
				                                        <%end if %>

                                              </td>
                                               <td style="border-right:1px #cccccc solid;">&nbsp;</td>
                                               <td colspan=3 valign=top class=lille style="padding:5px 5px 5px 10px; width:250px;">

                                                 <%if level <= 2 OR level = 6 then %>	
     
                                                <!--Budget -->
                                                            <b><%=tsa_txt_016 %></b>
                    
				                                            <br />Timer: <%=formatnumber((j_timerforkalk), 2)%> t.
                                                            <br />Bruttoomsætning: <%=formatnumber(aktdata(iRowLoop, 48), 2) %> DKK

                                                            <br />
                                                            <br /><b><%=tsa_txt_017 %>:</b><br />
					                                        <%
					                                        timerbrugtthis = 0
                                                            internKost = 0
                                                            oms = 0

                                                            if j_timerforkalk <> 0 then
                                                            j_timerforkalk = j_timerforkalk
                                                            else
                                                            j_timerforkalk = 1
                                                            end if 
							
							                                        '* finder antal brugte timer på job ***
							                                        strSQL2 = "SELECT timer AS sumtimer, (kostpris * timer * kurs / 100) AS kostpris, timepris FROM timer WHERE tjobnr = "& aktdata(iRowLoop, 31) &" AND ("& aty_sql_realhours &") ORDER BY timer"
							                                        oRec2.open strSQL2, oConn, 3
							                                        while not oRec2.EOF 

							                                        timerbrugtthis = timerbrugtthis + oRec2("sumtimer")
                            
                                                                    if aktdata(iRowLoop, 47) = 1 then '** Fastpris
                                                                    oms = oms + (oRec2("sumtimer") * (aktdata(iRowLoop, 48)/j_timerforkalk))
                                                                    else
                                                                    oms = oms + (oRec2("sumtimer") * oRec2("timepris"))
                                                                    end if

                                                                    internKost = internKost + oRec2("kostpris")
							                                        oRec2.movenext
                                                                    wend 
							                                        oRec2.close
							
							                                        if len(timerbrugtthis) > 0 then
							                                        timerbrugtthis = timerbrugtthis 
							                                        else
							                                        timerbrugtthis = 0
							                                        end if

                                                                    if len(internKost) > 0 then
							                                        internKost = internKost 
							                                        else
							                                        internKost = 0
							                                        end if
							
							                                        %>
					
					                                        Timer: <%=formatnumber(timerbrugtthis, 2)%> t. 
                                                            <br />
                    
                                                            Omsætning: <%=formatnumber(oms) %> DKK<br />
                                                            Intern kostpris: <%=formatnumber(internKost) %> DKK
					

                    
                                                            <%end if %>


                                                            <!-- stade --->

                                                            <br />
                                                            <span style="color:darkred;">
                                                            <%if cint(stade_tim_proc) = 0 then %>
                                                            <b>Restestimat: </b><%=restestimat &" timer"%>
                                                            <%else%>
                                                            <b>Afsluttet: </b><%=restestimat &" %"%>
                                                            <%end if %>
                                                            </span> (angivet) 
		
        
                                                </td>
                                                 <td style="border-right:1px #cccccc solid;">&nbsp;</td>
                                                <td colspan=2 valign=top style="padding:5px 5px 5px 10px; width:235px;">


                                                            <%'**** Rettigheder
				                                            if level <= 2 OR level = 6 then
				                                            editok = 1
				                                            else
					                                            if cint(session("mid")) = aktdata(iRowLoop, 44) OR cint(session("mid")) = aktdata(iRowLoop, 50) OR (cint(aktdata(iRowLoop, 44)) = 0 AND cint(aktdata(iRowLoop, 50)) = 0) then
					                                            editok = 1
					                                            end if
				                                            end if 
                                                            %>
                

                	                                        <%if editok = 1 AND fromsdsk = 0 then  %>

                   
                                                            <a href="jobs.asp?menu=job&func=red&id=<%=aktdata(iRowLoop, 4)%>&int=1&rdir=treg" class=rmenu target="_top">Rediger job >></a> 
				                                        <br />
				                                        <a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&func=opret&jobid=<%=aktdata(iRowLoop, 4)%>&id=0&jobnavn=<%=aktdata(iRowLoop, 5)%>&fb=1&rdir=treg2&nomenu=1', '', 'width=850,height=600,resizable=yes,scrollbars=yes')" class=rmenu>Opret ny aktivitet >></a>
			
				                                        <br />
				                                        <a href="job_print.asp?menu=job&id=<%=aktdata(iRowLoop, 4)%>&kid=<%=aktdata(iRowLoop, 42)%>" target="_blank" class='rmenu'>Print / PDF >></a>
				
				                                        <br />
				                                        <a href="job_kopier.asp?func=kopier&id=<%=aktdata(iRowLoop, 4)%>&fm_kunde=0&filt=aaben&showaspopup=y&rdir=timereg&usemrn=<%=session("mid")%>&showakt=1" class='rmenu'>Kopier job >></a>
				                                        <%end if %>





                
				
				                                        <%if (level < 7) then %>
				
                                                        <br /><br />		
				                                        <a href="materialer_indtast.asp?id=<%=aktdata(iRowLoop, 4)%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aktdata(iRowLoop, 51)%>" target="_blank" class=rmenu><img src="../ill/ikon_materiale.png" border="0" /><%=tsa_txt_021 %> >> </a>
				                                        <br />
				                                        <%
				                                         antalmatreg = 0
				                                         fordeling = 0
				                                         strSQLmr = "SELECT COUNT(id) AS fordeling, SUM(matantal) AS matantal FROM materiale_forbrug WHERE jobid = "& aktdata(iRowLoop, 4) & " AND usrid = "& usemrn & " GROUP BY jobid, usrid"
				                                         'Response.Write strSQLmr 
				                                         'Response.flush
				 
				 
				                                         oRec4.open strSQLmr, oConn, 3
				                                         if not oRec4.EOF then
				                                         fordeling = oRec4("fordeling")
				                                         antalmatreg = oRec4("matantal")
				                                         end if
				                                         oRec4.close
				                                         %>
				
				                                        <%=tsa_txt_326 %> <a href="materiale_stat.asp?id=0&hidemenu=1&FM_job=<%=aktdata(iRowLoop, 4)%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aktdata(iRowLoop, 51)%>&FM_visprjob_ell_sum=0" target="_blank" class=vmenu><%=antalmatreg%> </a> <%=tsa_txt_327 %> <b><%=fordeling%></b> <%=tsa_txt_328 %>.
				
				                                        <%end if %>
				
				
			                                             <br /><br />
			                                             Send besked til jobansvarlige om at dit arbejdere på dette job er afsluttet.

            
                                                     &nbsp;

                                                        </td>
                                                        </tr>

                 

                                                         <!--- job besk --->

                                                         <%if (len(trim(aktdata(iRowLoop, 41))) <> 0 OR len(trim(aktdata(iRowLoop, 40))) <> 0) then 
	 
	                                                     select case lto
	                                                     case "dencker", "intranet - local"
	                                                     jobbeskdivSHOW = "visible"
	                                                     jobbeskdivDSP = ""
	                                                     case else
	                                                     jobbeskdivSHOW = "hidden"
	                                                     jobbeskdivDSP = "none"
	                                                     end select
	 
	                                                     %>
	                                                                 <!---job besk -->
	                    
                                                                    <tr bgcolor="#FFFFFF"><td colspan=9>
                                                                    <br />
	                        
                          
                                                                        <a href="#" class="showjobbesk" id="showjobbesk_<%=aktdata(iRowLoop, 4) %>" class=rmenu>(+ Se <%=tsa_txt_029 %>: <%=len(trim(aktdata(iRowLoop, 40))) &" "& tsa_txt_371 &"." %>)</a>
                             
	           
	          
				                                                    <div id="jobbeskdiv_<%=aktdata(iRowLoop, 4)%>" name="jobbeskdiv" style="overflow:auto; height:150px; padding:10px 0px 10px 5px; visibility:<%=jobbeskdivSHOW%>; display:<%=jobbeskdivDSP%>; border:1px #cccccc solid; background-color:snow;">
				                                                    <table border=0 cellspacing=0 cellpadding=5 width=100%>
                                                                    <tr><td valign=top style="width:450px;"><b>Job beskrivelse:</b><br />
				                                                    <%=trim(aktdata(iRowLoop, 41))%>
				                                                    </td>
				
				                                                    <%if len(trim(aktdata(iRowLoop, 40))) <> 0 then %>
				                                                    <td valign=top style="width:250px;">
				                                                    <b>Intern note:</b><br />
				                                                    <%=trim(aktdata(iRowLoop, 40)) %></td>
				                                                    <%end if %>
				
				                                                    </tr></table>
				                                                    </div>
	
	                                                                    <br />&nbsp;
	
	
	                                                                    </td></tr>
                    
                           





	                                                    <%end if 
                                                        '***** SLUT Job stamdata + beskrivesle ****'


                                                        %>

                                                          <tr>
                  <td colspan=9 style="height:20px;" bgcolor="#FFFFFF">
                    <img src="../ill/blank.gif" width=1 height=1 /></td>  </tr>
                                                        </table>
                                                        </div>

                                                        <%


          
			if request.Cookies("job_"&aktdata(iRowLoop, 4)&"") <> "visible" then
                
               
			    job_vzb = "hidden"
			    job_dsp = "none"
               
			else
			job_vzb = "visible"
			job_dsp = ""
			end if
			
			
			
			%>



            

            <div id="div_timereg_<%=aktdata(iRowLoop, 4)%>" style="position:relative; background-color:#ffffff; height:250px; overflow:auto; padding:0px 3px 3px 3px; width:745px; border:1px #8caae6 solid; visibility:<%=job_vzb %>; display:<%=job_dsp%>; z-index:2000">
   
            <table cellspacing="0" cellpadding="2" border="0" width="100%" bgcolor="#8caae6">
            

            <tr>
                  <td colspan=9 style="background-color:#D6dff5; background-image:url('../ill/stripe_10.png'); height:2px;">
                    <img src="../ill/blank.gif" width=1 height=1 /></td>  </tr>

         
         
         
         
         
                                                
                
               
            
      <%

                    '*** LastFakdato ***'
                    
                    
					lastfakdato = "1/1/2001"
					
					strSQLFAK = "SELECT f.fakdato FROM fakturaer f WHERE f.jobid = "& aktdata(iRowLoop, 4) &" AND f.fakdato >= '"& varTjDatoUS_man &"' AND faktype = 0 ORDER BY f.fakdato DESC"
					
                    
                    oRec2.open strSQLFAK, oConn, 3
					if not oRec2.EOF then
						if len(trim(oRec2("fakdato"))) <> 0 then
						lastfakdato = oRec2("fakdato")
						end if
					end if
					oRec2.close
	                

	            call dageDatoer()
	          

            antalJobLinier = antalJobLinier + 1 
            end if

			'*** Faser ***'
			
			
			if len(trim(aktdata(iRowLoop, 29))) <> 0 then
			fsVal = lcase(aktdata(iRowLoop, 29))
			else
			faVal = "nonexx99w"
			end if
			
			
			
			
			if lastFase <> lcase(aktdata(iRowLoop, 29)) AND len(trim(aktdata(iRowLoop, 29))) <> 0 AND cint(intEasyreg) <> 1 then %>
            <tr>
				<td bgcolor="#FFFFFF" colspan=9 valign=top style="padding:0px 0px 0px 0px; height:10px; border-bottom:0px #cccccc solid;"><img src="../ill/blank.gif" width="1" height="1" border="0" alt="" /><br />
                </td>
            </tr>
			<tr>
				<td bgcolor="#Eff3ff" colspan=2 valign=top style="padding:5px 5px 4px 2px; border:1px #cccccc solid; border-bottom:0px; border-left:0px;">
			    <a id="<%=lcase(aktdata(iRowLoop, 29)) %>" name="<%=lcase(aktdata(iRowLoop, 29)) %>" href="#<%=lcase(aktdata(iRowLoop, 29)) %>" class=afase> + <%=tsa_txt_329 %>: <%=replace(aktdata(iRowLoop, 29), "_", " ") %></a>
                </td>

                <td bgcolor="#FFFFFF" colspan=7 valign=top style="padding:5px 5px 4px 2px; border-bottom:1px #cccccc solid;">
			  &nbsp;
                </td>
			    
			</tr>
			<%
			'if request.Cookies("tsa_fase")(""&jobid&"_"& lcase(aktdata(iRowLoop, 29)) &"") <> "visible" then
                
                'if iRowLoop < 10 then
                tr_vzb = "visible"
			    tr_dsp = ""
                'else
			    'tr_vzb = "hidden"
			    'tr_dsp = "none"
                'end if

			'else
			'tr_vzb = "visible"
			'tr_dsp = ""
			'end if
			
			
			%>
			<input name="faseshow" id="faseshow_<%=lcase(aktdata(iRowLoop, 29)) %>" type="hidden" value="<%=tr_vzb %>" />
			<input name="faseshowid" type="hidden" value="<%=aktdata(iRowLoop, 4)&"_"&lcase(aktdata(iRowLoop, 29)) %>" />
			<%
			
			
			lastFase = lcase(aktdata(iRowLoop, 29))
            
			
			else 
			        if len(trim(tr_vzb)) = 0 then
			        tr_vzb = "visible"
			        tr_dsp = ""
			        else
			        tr_vzb = tr_vzb
			        tr_dsp = tr_dsp
			        end if
			%>
			<%end if '**faser
			
			
			
			
			
			%>
			
			
			
			<tr class="td_<%=lcase(aktdata(iRowLoop, 29)) %>" style="visibility:<%=tr_vzb%>; display:<%=tr_dsp%>;">
				<td bgcolor="#ffffff" valign=top height=30 style="width:250px; padding-right:3px; border-bottom:1px #cccccc solid;">
				
                
                <!-- akt linje -->

                <!-- kunde -->
                 <%if cint(intEasyreg) = 1 then%>
		        
                <span style="font-size:9px;"><%=aktdata(iRowLoop, 32) %></span>
		        <br />
		        
		        <%end if%>

		        
		        <!-- jobnavn & fase -->
                   <%if cint(intEasyreg) = 1 then%>
		      
		        <span style="font-size:9px; color:#000000;"><b><%=left(aktdata(iRowLoop, 5), 50) %></b> (<%=aktdata(iRowLoop, 31) %>)</span><br />

                <%if len(trim(aktdata(iRowLoop, 29))) <> 0 then %>
                <span style="font-size:9px; color:yellowgreen;"><%=left(aktdata(iRowLoop, 29), 50) %></span><br />
                <%end if %>
                

		        <%end if%>


                <!-- akt. navn -->
                <!-- 2009 bruges ikke mere, 
				da der ikke oprettes ikke fakbare akt. fra SDSK. 
				Gemmes til gamle aktiviteter -->
				<b><%=replace(aktdata(iRowLoop, 2), "(ej fakturerbar)", "")%></b><br />

		     
				
                

                <span style="font-size:8px; line-height:10px; color:#999999;">
                <%=datepart("d", aktdata(iRowLoop, 13), 2,2) &". "& left(monthname(datepart("m", aktdata(iRowLoop, 13), 2,2)), 3) &". "& datepart("yyyy", aktdata(iRowLoop, 13), 2,2) %> 
				- 
                <%=datepart("d", aktdata(iRowLoop, 14), 2,2) &". "& left(monthname(datepart("m", aktdata(iRowLoop, 14), 2,2)), 3) &". "& datepart("yyyy", aktdata(iRowLoop, 14), 2,2) %>
				<%if aktdata(iRowLoop, 12) <> 0 then %>
				&nbsp;<%="<i>"& tsa_txt_036 &": "& aktdata(iRowLoop, 12) &"</i>" %>
				<%end if %>
                </span>

				
				
                <br />
				<span style="font-family:Arial; line-height:10px; font-size:8px; color:#5582d2;">
				<%
				call akttyper(aktdata(iRowLoop, 11), 1)
                %>

		        <%=left(akttypenavn, 12)%> 
				
                <%
                
                '** Timepris **'
                'level = 100
				'if level <= 2 OR level = 6 then
				 '** Admin eller jobansvarlig **'
                if level = 1 OR session("mid") = aktdata(iRowLoop, 44) OR session("mid") = aktdata(iRowLoop, 50) then
                'if level = 1000 then
				        call akttyper2009Prop(aktdata(iRowLoop, 11))
				        if cint(aty_fakbar) = 1 OR aktdata(iRowLoop, 11) = 5 then '** Viser tpriser på fakturerbare eller kørsel
				
				            if cint(j_intjobtype) = 1 then 'fastpris (altid angivet i DKK (Indtild videre!))
				                'if cint(j_usejoborakt_tp) = 0 then 'job ell. akt. grundlag
				        
				                    'if j_timerforkalk <> 0 then
				                    'j_timerforkalk = j_timerforkalk
				                    'else
				                    'j_timerforkalk = 1
				                    'end if
				        
				                aktTpThis = formatnumber(j_budget/j_timerforkalk, 2) '&" DKK"
				                'else
				                'aktTpThis = formatnumber(aktdata(iRowLoop, 30), 2) '&" DKK"
				                'end if
				    
				            else 'medarbtimepriser
				                    '*** tjekker for alternativ timepris på aktivitet
				                    call alttimepris(aktdata(iRowLoop, 3), aktdata(iRowLoop, 4), usemrn, 0)  
				      
				                    '** Er der alternativ timepris på jobbet
                                    if foundone = "n" then
                                    call alttimepris(0, aktdata(iRowLoop, 4), usemrn, 0) 
                                    end if
				       
				                    aktTpThis = formatnumber(intTimepris, 2) '& " " & valutaISO
				                    'aktTpThis = "..."

				            end if  %>
				
				        // <%=aktTpThis & " DKK" %> 
				
				        <%end if %>
				
				
				
				
				<%end if %>


                
                // <%=aktdata(iRowLoop, 36) %> </span>



                <%if aktdata(iRowLoop, 34) = 1 then 'bg = timer
                    if aktdata(iRowLoop, 37) > aktdata(iRowLoop, 33) then
                    fcolreal = "red"
                    else
                    fcolreal = "#2c962d"
                    end if
                else
                fcolreal = "#999999"
                end if
                     %>
                <!-- Real -->
                
                <span style="font-family:Arial; font-size:8px; color:<%=fcolreal%>;">//  <%=formatnumber(aktdata(iRowLoop, 37), 2) %></span>
				<span style="font-family:Arial; font-size:8px; color:#5582d2;"> (<%=formatnumber(aktdata(iRowLoop, 38), 2) %>)</span>
				
				<%
                'if level <= 2 OR level = 6 then
                '** Admin eller jobansvarlig **'
                if level = 1000 then '**  Viser ikke kospriser 2011.7.5
                'if level = 1 OR session("mid") = aktdata(iRowLoop, 44) OR session("mid") = aktdata(iRowLoop, 50) then
                if (cint(aty_fakbar) = 1 OR aktdata(iRowLoop, 11) = 2) AND formatnumber(aktdata(iRowLoop, 39), 0) <> 0 then 'if lto = "local-intranet" OR lto = "epi" then %>
               
                <span style="font-family:Arial; font-size:8px; color:#999999">// Kost.: <%=formatnumber(aktdata(iRowLoop, 39), 0) %>,-</span>
                <%end if %>
				<%end if %>
				
				<input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=aktdata(iRowLoop, 4)%>">
				<input type="hidden" name="FM_aktivitetid" id="FM_aktivitetid" value="<%=aktdata(iRowLoop, 3)%>"> 
				
				
				 <%if len(trim(aktdata(iRowLoop, 15))) <> 0 then %>
		       
		        
		        
		        
		         
		        <%
		        'len_abesk = len(aktdata(iRowLoop, 15))
		        'if len_abesk > 50 then
		        'len_abesk_erskrevet = 50
		        'else
		        'len_abesk_erskrevet = len(aktdata(iRowLoop, 15))
		        'end if
		        
		        %>
		        
		        <div id="aktbesk_<%=aktdata(iRowLoop, 3)%>" style="position:relative; visibility:visible; display:; left:0px; top:0px; width:200px; height:14px; overflow:hidden; z-index:200000;">
		        <table bgcolor="#ffffff" border=0 cellspacing=0 cellpadding=0 width=100%>
		        <tr>
		        <td style="padding:0px;" valign=top class=lille>
		        <a href="javascript:ShowHideAktBesk('aktbesk_<%=aktdata(iRowLoop, 3)%>');" class="silverlille"><%=aktdata(iRowLoop, 15)%></a>
		        </td></tr>
		        </table>
		        </div>
		        
		        <%end if %>
				
				</td><td bgcolor="#ffffff" valign=top style="padding:8px 3px 3px 3px; border-bottom:1px #cccccc solid; border-right:1px #cccccc solid;">
				<%if level = 1000 then 'cint(dsksOnOff) = 1  stopur slået fra %>
				<a href="javascript:popUp('stopur_2008.asp?func=ins&incid=0&logentry=0&jobid=<%=aktdata(iRowLoop, 4)%>&FM_mid=<%=usemrn%>&aktid=<%=aktdata(iRowLoop, 3)%>','1000','650','20','20')" class=vmenu><img src="../ill/stopur_st_stop.png" alt="<%=tsa_txt_037 %>" border=0 /></a>
				<%else %>
				&nbsp;
				<%end if %>
				</td>
		        
		      
		        
			<%
			
			ugedagUdskrevet = "#"
		
		end if%>
			
			<%
			
			for x = 1 to 7
				
				'*** Standard ****
                maxl = 5
                fmbgcol = "#ffffff"
                fmborcl = "1px #999999 solid"
				
				
						if tjekdag(x) = aktdata(iRowLoop, 0) then%>
						<%			
								
								
								
								if len(aktdata(iRowLoop, 6)) <> 0 then
								kommtrue = "+"
								else
								kommtrue = "<font color='#999999'>+</font>"
								end if 
						
								
								if x = 6 OR x = 7 OR helligdageCol(x) = 1 then
								bgColor = "gainsboro"
								else
								bgColor = "#FFFFFF"
								end if
								
								erIndlast = 1
								
								call fakfarver(lastfakdato, tjekdag(1), tjekdag(x), erIndlast, usemrn, aktdata(iRowLoop, 11))
								
								felt = "<td valign=top bgcolor='"&bgColor&"' style=""border-bottom:1px #cccccc solid; border-right:1px #cccccc solid; padding-top:6px;"" id='td_"&iRowLoop&""&x&"'>" 
								
								call akttyper2009Prop(aktdata(iRowLoop, 11)) 
								
								if visTimerElTid <> 1 OR aktdata(iRowLoop, 11) = 5 OR aktdata(iRowLoop, 11) = 61 OR aty_pre <> 0 OR cint(intEasyreg) = 1 then
								timtype = "text"
								tidtype = "hidden"
								br = ""
								else
								timtype = "hidden"
								tidtype = "text"
								br = "<br>"
								end if
								
								'if lto <> "execon" then
								felt = felt &"<input class='dcls_"&x&"' type='"& timtype &"' name='FM_timer' id='FM_timer_"&iRowLoop&"_"&x&"' value='"& aktdata(iRowLoop, 1) &"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; border:"&fmborcl&"; height:16px; width:45px; font-family:arial; font-size:10px; line-height:12px;"" onkeyup=""doKeyDown()"" onfocus=""markerfelt('"&iRowLoop&""&x&"','"&x&"'), showKMdailog('"&aktdata(iRowLoop, 11)&"', '"&iRowLoop&""&x&"', '"& aktdata(iRowLoop, 27) &"', '"& iRowLoop &"', '"& topAdd &"')"";>"
								felt = felt &"<input type='hidden' name='FM_timer_opr' id='FM_timer_opr_"&iRowLoop&"_"&x&"' value='"& aktdata(iRowLoop, 1) &"'>"
								'else
								'felt = felt &"<input type='"& timtype &"' name='FM_timer' id='FM_timer' value='"& aktdata(iRowLoop, 1) &"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; border:"&fmborcl&"; height:16px; width:45px; font-family:arial; font-size:10px; line-height:12px;"" onfocus=""markerfelt('"&iRowLoop&""&x&"','"&x&"')"";>"
								'end if
								
								
								felt = felt &"<input type='"& tidtype &"' name='FM_sttid' id='FM_sttid' value='"& aktdata(iRowLoop, 8) &"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; height:16px; width:35px; font-family:arial; font-size:10px; line-height:12px; border:"&fmborcl&";"">"_
								&""& br &"<input type='"& tidtype &"' name='FM_sltid' id='FM_sltid' value='"& aktdata(iRowLoop, 9) &"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; height:16px; width:35px; font-family:arial; font-size:10px; line-height:12px; border:"&fmborcl&";"">"
								
								if maxl <> 0 then
								felt = felt &"&nbsp;<a class=""a_showkom"" href=""#anc_s0"" name='anc_"&iRowLoop&"_"&x&"' onClick=""expandkomm('"&iRowLoop&"', '"&x&"')"">"& kommtrue &"</a>"
								end if
								
								felt = felt &"<input type='hidden' name='FM_timer' id='FM_timer' value='#'><input type='hidden' name='FM_datoer' id='FM_datoer' value='"& tjekdag(x) &"'>"_
								&"<input class='d_kom_cls_"&x&"' type='hidden' name='FM_kom_"&iRowLoop&""&x&"' id='FM_kom_"&iRowLoop&""&x&"' value='"&aktdata(iRowLoop, 6)&"'>"_
								&"<input type='hidden' name='FM_feltnr' id='FM_feltnr' value='"&iRowLoop&""&x&"'>"
								

								if len(aktdata(iRowLoop, 7) ) <> 0 then
								kommOnOff = aktdata(iRowLoop, 7) 
								else
								kommOnOff = 0
								end if
								
								
								felt = felt &"<input type='hidden' name='FM_off_"&iRowLoop&""&x&"' id='FM_off_"&iRowLoop&""&x&"' value='"&kommOnOff&"'>"
								felt = felt &"<input id='FM_bopal_"&iRowLoop&""&x&"' name='FM_bopal_"&iRowLoop&""&x&"' type=""hidden"" value="""& aktdata(iRowLoop, 27) &"""/>"
								felt = felt &"<input id='FM_destination_"&iRowLoop&""&x&"' name='FM_destination_"&iRowLoop&""&x&"' type=""hidden"" value="""& aktdata(iRowLoop, 28) &"""/>"
								
								felt = felt &"</td>"
								flt(x) = felt
								
							    
							    
							    
							    if cint(aty_real) = 1 then
							    
							    select case weekday(tjekdag(x), 2)
									case 1
									totTimMan = totTimMan + aktdata(iRowLoop, 1)
									case 2
									totTimTir = totTimTir + aktdata(iRowLoop, 1)
									case 3
									totTimOns = totTimOns + aktdata(iRowLoop, 1)
									case 4
									totTimTor = totTimTor + aktdata(iRowLoop, 1)
									case 5
									totTimFre = totTimFre + aktdata(iRowLoop, 1)
									case 6
									totTimLor = totTimLor + aktdata(iRowLoop, 1)
									case 7
									totTimSon = totTimSon + aktdata(iRowLoop, 1)
									end select
							  
							    
							    end if
							        
						       
						%>
						<%else
								if instr(ugedagUdskrevet, weekday(tjekdag(x), 2)) = 0 then 
								
								
								
								
								if x = 6 OR x = 7 OR helligdageCol(x) = 1 then
								bgColor = "gainsboro"
								else
								bgColor = "#FFFFFF"
								end if
								
								call akttyper2009Prop(aktdata(iRowLoop, 11)) 
								
								if aty_pre <> 0 AND (x < 6) ANd helligdageCol(x) <> 1 then
								preVal = aty_pre
								else
								preVal = ""
								end if
								
								erIndlast = 0
								
								
                                call fakfarver(lastfakdato, tjekdag(1), tjekdag(x), erIndlast, usemrn, aktdata(iRowLoop, 11))
								
								felt = "<td valign=top bgcolor='"&bgColor&"' style=""border-bottom:1px #cccccc solid; border-right:1px #cccccc solid; padding-top:6px;"" id='td_"&iRowLoop&""&x&"'>" 
								
								
								
								if visTimerElTid <> 1 OR aktdata(iRowLoop, 11) = 5 OR aktdata(iRowLoop, 11) = 61 OR aty_pre <> 0 OR cint(intEasyreg) = 1 then
								timtype = "text"
								tidtype = "hidden"
								br = ""
								else
								timtype = "hidden"
								tidtype = "text"
								br = "<br>"
								end if
								
								
								'if lto <> "execon" then
								felt = felt &"<input class='dcls_"&x&"' type='"& timtype &"' name='FM_timer' id='FM_timer_"&iRowLoop&"_"&x&"' value='"&preVal&"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; border:"&fmborcl&"; height:16px; width:45px; font-family:arial; font-size:10px; line-height:12px;"" onkeyup=""doKeyDown()""  onfocus=""markerfelt('"&iRowLoop&""&x&"','"&x&"'), showKMdailog('"&aktdata(iRowLoop, 11)&"', '"&iRowLoop&""&x&"', '0', '"& iRowLoop &"', '"& topAdd &"')"";>"
								felt = felt &"<input type='hidden' name='FM_timer_opr' id='FM_timer_opr_"&iRowLoop&"_"&x&"' value='"& aktdata(iRowLoop, 1) &"'>"
								
								'else
								'felt = felt &"<input type='"& timtype &"' name='FM_timer' id='FM_timer' value='"&preVal&"' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; border:"&fmborcl&"; height:16px; width:45px; font-family:arial; font-size:10px; line-height:12px;"" onfocus=""markerfelt('"&iRowLoop&""&x&"','"&x&"')"";>"
							    'end if
								
								felt = felt &"<input type='"& tidtype &"' name='FM_sttid' id='FM_sttid' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; height:16px; width:35px; font-family:arial; font-size:10px; line-height:12px; border:"&fmborcl&";"">"_
								&""& br &"<input type='"& tidtype &"' name='FM_sltid' id='FM_sltid' maxlength='"&maxl&"' style=""background-color:"&fmbgcol&"; height:16px; width:35px; font-family:arial; font-size:10px; line-height:12px; border:"&fmborcl&";"">"
								
								if maxl <> 0 then
								felt = felt &"&nbsp;<a class=""a_showkom"" href=""#anc_s0"" name='anc_"&iRowLoop&"_"&x&"' onClick=""expandkomm('"&iRowLoop&"', '"&x&"')""><font color='#999999'>+</font></a>"
								end if
								
								felt = felt &"<input type='hidden' name='FM_timer' id='FM_timer' value='#'><input type='hidden' name='FM_datoer' id='FM_datoer' value='"& tjekdag(x) &"'>"_
								&"<input type='hidden' class='d_kom_cls_"&x&"' name='FM_kom_"&iRowLoop&""&x&"' id='FM_kom_"&iRowLoop&""&x&"' value=''>"_
								&"<input type='hidden' name='FM_feltnr' id='FM_feltnr' value='"&iRowLoop&""&x&"'>"
								
								
								if len(aktdata(iRowLoop, 7) ) <> 0 then
								kommOnOff = aktdata(iRowLoop, 7) 
								else
								kommOnOff = 0
								end if
								
								
								
								felt = felt &"<input type='hidden' name='FM_off_"&iRowLoop&""&x&"' id='FM_off_"&iRowLoop&""&x&"' value='"&kommOnOff&"'>"
							    felt = felt &"<input id='FM_bopal_"&iRowLoop&""&x&"' name='FM_bopal_"&iRowLoop&""&x&"' type=""hidden"" value=""0"" />"
								felt = felt &"<input id='FM_destination_"&iRowLoop&""&x&"' name='FM_destination_"&iRowLoop&""&x&"' type=""hidden"" value=""""/>"
								felt = felt &"</td>"
								flt(x) = felt
								
								
								
								end if%>
						
						<%end if%>
				<%
				ugedagUdskrevet = ugedagUdskrevet & "," & weekday(tjekdag(x), 2) & "#"
				'Response.write ugedagUdskrevet
				%>
				
				<%'end if%>
			<%next%>
		 	
	
	<%
	lastAktid = aktdata(iRowLoop, 3)
    lastJobid = aktdata(iRowLoop, 4)
	
	next
	
	
	'if writeendline = 1 then
		for x = 1 to 7
			Response.write flt(x) 
		next
	'end if
	%>
	
	<%end if 'iRowLoop%>
	
	</tr>
            <%if iRowLoop <> 0 then
            %>
            </table>
            </div>
            <!-- table Div-->


            <%end if %>
        
            
            <!-- id=incidentlist -->
            
            </td></tr>
            </tbody>
            </table>

            <%if antalJobLinier <> 0 then %>

             <br /><br />

     <div id="dagstotaler" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:745px; border:1px #8caae6 solid; left:0px; visibility:visible; display:; z-index:1000">
   


     <table cellspacing="0" cellpadding="2" border="0" width="100%" bgcolor="#8caae6">

     <% call dageDatoer() %>

	<tr bgcolor="#FFDFDF">
		<td colspan=2 style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><b><%=tsa_txt_038 %>:</b> (timer på viste job)</td>
		<td style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_1" type="text" size=4 value="<%=formatnumber(totTimMan, 2)%>" style="border:0px #8caae6 solid; background-color:#FFDFDF; font-size:10px;" /></td>
	<td style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_2" type="text" size=4 value="<%=formatnumber(totTimTir, 2)%>" style="border:0px #8caae6 solid; background-color:#FFDFDF; font-size:10px;" /></td>
	<td style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_3" type="text" size=4 value="<%=formatnumber(totTimOns, 2)%>" style="border:0px #8caae6 solid; background-color:#FFDFDF; font-size:10px;" /></td>
	<td style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_4" type="text" size=4 value="<%=formatnumber(totTimTor, 2)%>" style="border:0px #8caae6 solid; background-color:#FFDFDF; font-size:10px;" /></td>
	<td style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_5" type="text" size=4 value="<%=formatnumber(totTimFre, 2)%>" style="border:0px #8caae6 solid; background-color:#FFDFDF; font-size:10px;" /></td>
	<td bgcolor="#cccccc" style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_6" type="text" size=4 value="<%=formatnumber(totTimLor, 2)%>" style="border:0px #8caae6 solid; background-color:#cccccc; font-size:10px;" /></td>
	<td bgcolor="#cccccc" style="border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><input id="timer_7" type="text" size=4 value="<%=formatnumber(totTimSon, 2)%>" style="border:0px #8caae6 solid; background-color:#cccccc; font-size:10px;" /></td>
	</tr>
	
	
	
	
	<tr bgcolor="#ffffff">
		<td colspan=9 align=right valign=top style="padding-top:10px;">
		<input id="sm2" type="submit" value="<%=tsa_txt_085 %> >> "></td>
	</tr>
	<tr bgcolor="#ffffff"><td colspan=9>
    <br />Antal job: <%=antalJobLinier %>
    <br />Antal aktivitetslinier: <%=antalaktlinier %> (<%=loops %>)</td></tr>
    <input id="antalaktlinier" value="<%=antalaktlinier %>" type="hidden" />
   <input id="loops" value="<%=loops %>" type="hidden" />
    
    </table>

    
    </div>
    <%end if %>



            </div>



           
    
     
	
	
	<%'***** kommentar layer *** 



    %>
    
    <div id="kom_easy" name="kom_easy" style="position:absolute; visibility:hidden; display:none; left:140px; top:410px; z-index:950000000; width:300px; background-color:#ffffff; padding:10px 10px 10px 10px; border:10px #CCCCCC solid;">
	<table width=100% cellpadding=0 cellspacing=2 border=0><tr><td></td>
	<textarea cols="32" rows="5" id="FM_kommentar_easy" name="FM_kommentar_easy"></textarea><br>
        <input id="FM_kom_dagtype" value="1" type="hidden" />
      </td></tr>
      <tr><td align=right>
	<a href="#" class="vmenu" id="close_kom_easy"><%=tsa_txt_056 %>
    &nbsp;<img src="../ill/ikon_luk_komm.png" alt="<%=tsa_txt_056 %>" border="0"> </a><br />
    <br />Kommentar bliver først gemt når timer indlæses. Indlæses på alle aktiviteter den valgte dag.
	</td></tr></table>
	</div>
    
	<div id="kom" name="kom" style="position:absolute; visibility:hidden; display:none; left:140px; top:1010px; z-index:950000000; width:500px; background-color:#ffffFF; padding:10px 10px 10px 10px; border:10px #CCCCCC solid;">
	<table width=100% cellspacing=0 cellpadding=0 border=0>
	<tr>
	<td>
	<b><%=tsa_txt_051 %>:</b>
	<img src="../ill/blank.gif" width="190" height="1" alt="" border="0">
	(<%=tsa_txt_052 %>)&nbsp;
	<input type="hidden" name="daytype" id="daytype" value="1">
	<input type="hidden" name="rowcounter" id="rowcounter" value="0">
	<input type="text" name="antch" id="antch" size="3" maxlength="4"><br>
	<textarea cols="60" rows="10" id="FM_kommentar" name="FM_kommentar" onKeyup="antalchar();"></textarea><br>
	<br /><%=tsa_txt_053 %>?
	<select name="FM_off" id="FM_off" style="font-size:9px;">
	<option value="0" SELECTED><%=tsa_txt_054 %></option>
	<option value="1"><%=tsa_txt_055 %></option>
	</select>
	<%if kmDialogOnOff = 1 then %>
	<br />
	Ved kørselsregistrering,
	er dette en bopælstur?
	<select name="FM_bopalstur" id="FM_bopalstur" style="font-size:9px;">
	<option value="0" SELECTED><%=tsa_txt_054 %></option>
	<option value="1"><%=tsa_txt_055 %></option>
	</select>
	<%else %>
        <input name="FM_bopalstur" id="FM_bopalstur" value="0" type="hidden" />
	<%end if %>
	</td>
	</tr>
	<tr><td align=right>
	<a href="#" onclick="closekomm();" class="vmenu" id="clskom"><%=tsa_txt_056 %>
    &nbsp;<img src="../ill/ikon_luk_komm.png" alt="<%=tsa_txt_056 %>" border="0"> </a><br />
    <br />Kommentar bliver først gemt når timer indlæses.
    &nbsp;</td></tr>
	</table>
	</div>


<input type="hidden" name="sdskKomm" id="sdskKomm" value="<%=fromdeskKomm%>">
<input type="hidden" name="lasttd" id="lasttd" value="0">
<input type="hidden" name="lastkmdiv" id="lastkmdiv" value="0">
<input type="hidden" name="lasttddaytype" id="lasttddaytype" value="0">
</form>
	
	
	
	
	   <!---- Side indstillinger --->
	   <div id="pagesetadv" style="position:absolute; display:none; visibility:hidden; top:500px; left:125px; width:380px; height:155px; z-index:10000000000; background-color:#ffffff; border:1px #999999 solid; padding:3px;">
        
        <table cellspacing=0 cellpadding=2 border="0" bgcolor="#FFFFFF" width=100%>
         <form action="timereg_akt_2006.asp?FM_use_me=<%=usemrn%>&showakt=1&jobid=<%=jobid%>&fromsdsk=<%=fromsdsk%>" method="POST" name="timereg">
	   <tr bgcolor="#8caae6">
	   <td><img src="../ill/gear.png"/></td>
	   <td class=alt align=left style="width:150px;"><b><%=tsa_txt_307%></b></td>
	   <td align=right><a href="#" id="pagesetadv_close" class=red>[x]</a></td></tr>
        <tr><td colspan=3>
        
		<b><%=tsa_txt_033 %>:</b><br /> 
		<%if visTimerElTid <> 1 then
		chk0 = "CHECKED"
		chk1 = ""
		else
		chk1 = "CHECKED"
		chk0 = ""
		end if%>
		<input type="radio" name="FM_vistimereltid" value="0" <%=chk0%>> <%=tsa_txt_034 %> <input type="radio" name="FM_vistimereltid" value="1" <%=chk1%>> <%=tsa_txt_035 %> 
		
        <br>
		
		<%
		if cint(ignorertidslas) = 1 then
		ignorertidslasCHK = "checked"
		else
		ignorertidslasCHK = ""
		end if
		%>
		
		<input type="checkbox" name="FM_igtlaas" id="FM_igtlaas" value="1" <%=ignorertidslasCHK%>> <%=tsa_txt_046 %>
            <input id="FM_chkfor_ignorertidslas" name="FM_chkfor_ignorertidslas" type="hidden" value="1"/>
		
		
		<br />
		
		<%
		if cint(ignJobogAktper) = 1 then
		ignJobogAktperCHK = "checked"
		else
		ignJobogAktperCHK = ""
		end if
		%>
		
		<input type="checkbox" name="FM_ignJobogAktper" id="FM_ignJobogAktper" value="1" <%=ignJobogAktperCHK%>> <%=tsa_txt_286 %>
         <input id="Hidden1" name="FM_chkfor_ignJobogAktper" type="hidden" value="1"/>
		
	   </td>
	   </tr>
	   <tr>
	   <td colspan=3 align=right><input type="submit" value="<%=tsa_txt_086 %> <> " style="height:18px; font-family:arial; font-size:10px;"></td>
	   </tr>
	   	</form>
	   </table>
	   </div>
	
	
	<!--
	</div>--><!-- table div -->
  
	
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	&nbsp;
	<!--Udfyld resterende timer (op til din normeret uge) på følgende aktivitet:<br><br>-->
	<% 




    if antalJobLinier > 10 OR antalaktlinier > 100 AND session("forste") = "j" then
    
    itop = 780
ileft = 240
iwdt = 300
idsp = ""
ivzb = "visible"
iId = "tregaktmsg1"
call sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb)
%>
			
            <b>Der er valgt mere end 10 job eller mere end 100 aktiviteter på denne liste</b><br><br />
	        Hvis der er vises mere end 10 job (<b><%=antalJobLinier %></b>) eller over 100 aktiviteter (<b><%=antalaktlinier%></b>), bliver visningen langsommere (> 10 sek.) og det er lettere at miste overblikket. 
            <br><br>
            Du bør derfor overveje at vælge færre job eller rydde op i antallet af aktiviteter.
            <br><br>
	        
	        
	</td></tr></table>
	</div>
	

<%

    end if


	at = 9999
	if iRowLoop <= 1 AND showakt = 1 AND at <> 9999 then
 
itop = 180
ileft = 540
iwdt = 300
idsp = ""
ivzb = "visible"
iId = "tregaktmsg1"
call sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb)
%>
			
            <b><%=tsa_txt_059 %></b><br><br />
	        <%=tsa_txt_060 %><br><br>
	        
	        
	</td></tr></table>
	</div>
	

<%
end if

t = 1000
	if t = 999 AND session("forste") = "j" then
 
itop = 120
ileft = 400
iwdt = 300
idsp = ""
ivzb = "visible"
iId = "tregaktmsg1"
call sidemsgId2(itop,ileft,iwdt,iId,idsp,ivzb)
%>

			
            <b>Velkommen til vores nye søgefilter på timereg. siden</b><br><br />
	        Vi har forsøgt at gøre timeregistreringen enklere, ved at skille
	        væsentlig information fra uvæsentlig og tilføje et nyt søge filter - forhåbentlig til glæde for alle.<br /><br />
	        Hvis I har spørgmål sidder vi klar ved support telefonen på 2684 2000, eller på <a href="mailto:support@outzource.dk" class=vmenu>support@outzource.dk</a>.<br />
	        <br />Mvh.<br />
	        OutZourCE
	        <br><br>
	        
	        
	</td></tr></table>
	</div>
	
<%end if
%>
	



<div id="velkommen" style="width:1px; height:1px; display:none; visibility:hidden;"></div>






<%response.flush %>


	
	
	
	
	
<!--pagehelp-->

<%
if showakt = 1 then

itop = 137
ileft = 635
iwdt = 120
ihgt = 0
ibtop = 40 
ibleft = 150
ibwdt = 600
ibhgt = 550
iId = "pagehelp"
ibId = "pagehelp_bread"
call sideinfoId(itop,ileft,iwdt,ihgt,iId,phDsp,phVzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
%>
        <b><%=tsa_txt_278 %></b><br /><%=tsa_txt_279 %><br /><br />
		<!--<b><=tsa_txt_047 %>:</b>&nbsp;<=tsa_txt_048 %><br><br /> -->
		<b><%=tsa_txt_049 %>:</b>&nbsp;<%=tsa_txt_050 %> <br /><br />
		<b><%=tsa_txt_296 %>:</b><br />
		<%=tsa_txt_297 %>
		<br /><br />
		<b><%=tsa_txt_298 %>:</b><br />
		<%=tsa_txt_299 %>
		
		<br /><br /><b><%=tsa_txt_300 %></b><br />
		<%=tsa_txt_301 %>
		<br /><br /><b><%=tsa_txt_295 %>:</b><br />
		
		<%for f = 1 to 6
		
		select case f
		case 1
		xxfmborcl = "1px #999999 solid"
		xxfmbgcol = "#ffffff"
		xxtxt = tsa_txt_285 '"Åben. Timer ikke indlæst. (Pre-udfyldte aktiviteter, f.eks. frokost)"
		case 2
		xxfmborcl = "1px yellowgreen solid"
		xxfmbgcol = "#ffffff"
		xxtxt = tsa_txt_290 '"Åben. Timer indlæst i database."
		case 3
		xxfmborcl = "1px yellowgreen dashed"
		xxfmbgcol = "#cccccc"
		xxtxt = tsa_txt_291 '"Lukket for indtastning. Registrering er godkendt. (kan ydermere også være faktureret eller lukket via smiley-ordning)"
		case 4
		xxfmborcl = "1px #999999 dashed"
		xxfmbgcol = "#cccccc"
		xxtxt = tsa_txt_292 '"Lukket for indtastning. Er enten faktureret eller uge er lukket via smiley-ordning."
		case 5
		xxfmborcl = "1px #999999 dashed"
		xxfmbgcol = "#ffffff"
		xxtxt = tsa_txt_293 '"Lukket for indtastning. Uge er lukket via smiley ell. fast periode, men admin. brugere kan ændre."
		case 6
		xxfmborcl = "1px #cccccc dashed"
		xxfmbgcol = "#EfF3FF"
		xxtxt = tsa_txt_294 '"Tidslåst."
		end select
		
		Response.write "<input type='text' name='xx' id='xx' value='' maxlength='0' style=""background-color:"&xxfmbgcol&"; border:"&xxfmborcl&"; height:16px; width:45px; font-family:arial; font-size:10px; line-height:12px;""> "& xxtxt &"<img src='ill/blank.gif' height=3 width=1><br><br>"
								
		
		next %>
		
		
		</td>
	</tr></table></div>
	
	<%end if %>
	


    <%
    response.flush
    
   k = 0
    if k = 999 then %>
   <!-- ugeseddel -->
   <div id="afstem" style="position:absolute; left:20px; top:158px; display:<%=dAfstemDsp%>; visibility:<%=dAfstemVzb%>; width:<%=gblDivWdt%>px; height:10020px; overflow:auto; background-color:#FFFFFF; padding:2px; border:1px #8caae6 solid; z-index:2000;">
   <%
   ugeseddelvisning = 1
   call ugeseddel(usemrn, varTjDatoUS_man, varTjDatoUS_son, ugeseddelvisning)  %>
   </div>
   
   <%end if %>
   
    
   
    
    <!-- login historik --->
    
    <%
    if k = 999 then
    
    if session("stempelur") <> 0 then %>
    <div id=stempelur style="position:absolute; left:10px; top:158px; display:<%=dStempelDsp%>; visibility:<%=dStempelVzb%>; width:<%=gblDivWdt%>px; height:1020px; overflow:auto; background-color:#FFFFFF; padding:3px; border:1px #8caae6 solid; z-index:2000;">
    <table cellpadding=0 cellspacing=0 border=0 width=100%><tr>
     <td valign=top width=70% style="padding:10px;">
     
	<%call fLonTimerPer(varTjDatoUS_man, 7, 0, useMrn) %>
	
	
	
	</td>
	<td align=right valign=top style="padding:10px 10px 0px 0px;">
	<table cellpadding=0 cellspacing=0 border=0 width=80>
	<tr>
	<td valign=top align=right><a href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=3&strdag=<%=day(useDatePrevWeek)%>&strmrd=<%=month(useDatePrevWeek)%>&straar=<%=year(useDatePrevWeek)%>&fromsdsk=<%=fromsdsk%>"><img src="../ill/nav_left_blue.png" border="0" /></a></td>
   <td style="pading-left:20px;" valign=top align=right><a href="timereg_akt_2006.asp?jobid=<%=jobid%>&usemrn=<%=usemrn%>&showakt=3&strdag=<%=day(useDateNextWeek)%>&strmrd=<%=month(useDateNextWeek)%>&straar=<%=year(useDateNextWeek)%>&fromsdsk=<%=fromsdsk%>"><img src="../ill/nav_right_blue.png" border="0" /></a></td>
	</tr>
	</table>
	</td></tr>
	<tr>
	<td valign=top colspan=2 style="padding:10px;">
	
	
	<%call stempelurlist(useMrn, 0, 1, varTjDatoUS_man, varTjDatoUS_son,0) %>
	
	</td></tr>
    </table>
    
    
    </td></tr>
    </table>
    
    </div>
    <%
	end if
	
	end if 'k
    %>
	
	




<%


'** ikke længere første logon-. == Cookies må benyttes. ***
'** bruges bl.a til i kalender til valg af dato eller cookie **'
session("forste") = "n"


	
else





            itop = 550
            ileft = 20
            iwdt = 300
            idsp = ""
            ivzb = "visible"
            iId = "tregaktmsg2"
            call sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb)
            %>
            			
                        
            	         <b><%=tsa_txt_347 %></b><br />
            	         <%=tsa_txt_348 %>
            	         <br /><br><b><%=tsa_txt_349 %></b><br />
            	         <%=tsa_txt_350 %>
	                    
	                    <br /><br />
            	        
	                    <%=strprojektgrpOK %>

                        <br /><br />
                        <b>Nulstil</b><br />
                        Hvis du søgningen er stoppet kan du nulstille søgningen ved at klikke her.<br />
                        <a href="timereg_akt_2006.asp?jobid=0&showakt=1" class=vmenu>Nulstil søgning >></a>
	                    <br /><br />
                       
	                   
	                    
	                   &nbsp;
	            </td></tr></table>
	            </div>
            	

            <%
       
        
end if 'projgrpOK, adgang via projgrp eller vi link til tilføj sig til job 



end if 'Session User





%>
<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />

 <%


 'if showakt = 1 then
 '   call helpandguides(lto, 1) 
 'else
 '   call helpandguides(lto, 0)
 'end if
 %>

<!--
<br /><br /><br /><br /><br />&nbsp;
  <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
   <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
     <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
    
-->
 

 
  
<!--#include file="../inc/regular/footer_inc.asp"-->
