<%response.buffer = true%>


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file="CuteEditor_Files/include_CuteEditor.asp" -->
<!--#include file="../inc/regular/topmenu_inc.asp"--> 
<!--#include file="inc/convertDate.asp"-->
<!--#include file="inc/timbudgetsim_inc.asp"-->



<%'GIT 20160811 - SK


 'response.write "request(FM_fomr): "  & request("FM_fomr") & "<br>"
    

'** Jquery section 
if Request.Form("AjaxUpdateField") = "true" then




Select Case Request.Form("control")

case "FN_getMatPris"

    matid = request("matid")

    strMatPris = 0

     if len(trim(matid)) <> 0 then

    strSQLmat = "SELECT m.salgspris FROM materialer m "_
    &" WHERE m.id = "& matid
    
    oRec.open strSQLmat, oConn, 3
	if not oRec.EOF then 

    strMatPris = oRec("salgspris") 

    end if
    oRec.close

    end if

    Response.write strMatPris
    response.end


case "FN_getMatlisten"

    sog_val = request("jq_sog_val")
    uval = request("uval")
    strMat = ""

    

    if len(trim(sog_val)) <> 0 then

    strSQLmat = "SELECT m.navn AS mnavn, m.matgrp, m.varenr AS mvnr, m.id, m.antal, minlager FROM materialer m "_
    &" WHERE m.navn LIKE '%"& sog_val &"%' OR m.varenr LIKE '%"& sog_val &"%' OR m.betegnelse LIKE '"& sog_val &"%' OR m.lokation LIKE '"& sog_val &"%' ORDER BY m.navn DESC LIMIT 100" 
	
	'Response.Write strSQLmat
	'Response.end

    'Response.write "HEJ: "& sog_val &" - "& strMat
    'Response.end
	
	oRec.open strSQLmat, oConn, 3
	while not oRec.EOF 

                    itilbud = 0
                    itilbudTxt = ""
                    strSQLitilbud = "SELECT j.id, SUM(ju_stk) AS itilbud FROM job j LEFT JOIN job_ulev_ju ju ON (ju_jobid = j.id AND ju_matid = "& oRec("id") &") WHERE jobstatus = 3 AND ju_matid = "& oRec("id") &" GROUP BY ju_matid"
                    oRec9.open strSQLitilbud, oConn, 3
	                if not oRec9.EOF then

                        if oRec9("itilbud") <> 0 then 'isnull(oRec9("itilbud")) <> true then 
                        itilbudTxt = " - I tilbud: "& oRec9("itilbud")
                        else
                        itilbudTxt = ""
                        end if

                    end if
                    oRec9.close

    strMat = strMat & "<u><span style='font-size:14px; line-height:16px;' class=valgtmateriale id=valgtmateriale_"& uval &"_"& oRec("id") &">" & oRec("mnavn") & " ("& oRec("mvnr") &")</span></u> P� lager: "& oRec("antal") &" "& itilbudTxt &" <br>" '&nbsp;&nbsp;[ Min. lager: "& oRec("minlager") &"]<br>" 

    oRec.movenext
    wend 
    oRec.close

      call jq_format(strMat)
      strMat = jq_formatTxt

      Response.write strMat
      'Response.end

    end if
    Response.end

case else

%>
    
   
 <script src="../inc/jquery/jquery-1.10.1.min.js" type="text/javascript"></script>
    <script src="../inc/jquery/jquery-ui-1.10.3.custom/js/jquery-ui-1.10.3.custom.min.js" type="text/javascript"></script>

    <style>
        #sortable { list-style-type: none; margin: 0; padding: 0; width: 60%; }
        #sortable li { margin: 0 3px 3px 3px; padding: 0.4em; padding-left: 1.5em; font-size: 1.4em; height: 18px; }
        #sortable li span { position: absolute; margin-left: -1.3em; }
    </style>

    <script type="text/javascript">

        /*
        Restore globally scoped jQuery variables to the first version loaded
        (the newer version)
        */
        jq1101 = jQuery.noConflict(true);

    </script>
    <script src="../inc/jquery/dragSort.jquery.js" type="text/javascript"></script>
<!--End of added by Lei 17-06-2013-->


<%


Select Case Request.Form("control")
case "FM_sortOrder"
              Call AjaxUpdate("aktiviteter","sortorder","")

case "FN_getAktlisten"

               sfunc = request("func")

               if sfunc = "dbred" then          
              %>
              <script src="inc/job_jav_aktliste.js"></script> 
              <%
              
              jq_func = "red"
              jq_sort = ""
              jq_id = jq_visalle
             
              if len(trim(request.Form("cust"))) <> 0 then
              jq_id = request.Form("cust")
              else
              jq_id = 0
              end if
              
              jq_vispasluk = request("jq_visalle")

              'slto = request("slto") 'bruges kun ved test
              
               
              
              '*** Henter eksisterende ***'
	          call hentaktiviterListe(jq_id, jq_func, jq_vispasluk, sort)
	          'Response.write "<table><tr><td>Hej der1</td></tr></table>"


	          '*** ��� **'
              call jq_format(strAktListe)
              strAktListeTxt = jq_formatTxt
	          
	          Response.Write strAktListeTxt

             end if




case "FN_getKundelisten_2013"
              
              'Session.LCID = 1033

              dim kundeArr_2013
              redim kundeArr_2013(50)
              
              if len(trim(request.Form("cust"))) <> 0 then
              'jq_sog_val = request.Form("cust")

                '*** ��� **'
              'call jq_format(jq_sog_val)
              'jq_sog_val = jq_formatTxt

              'call utf_format(jq_sog_val)
              'jq_sog_val = utf_formatTxt

              'call utf_format(jq_sog_val)
              'jq_sog_val = utf_formatTxt

               'call htmlparseCSV(HTMLstring)
               'jq_sog_val = htmlparseCSVTxt 
                  'jq_sog_val = Server.URLEncode(jq_sog_val)

                  'jq_sog_val = Server.HTMLDecode(jq_sog_val)

                   'jq_sog_val = replace(jq_sog_val, "%C3%A6%", "&aelig;")
                   'jq_sog_val = replace(jq_sog_val, "C3", "&aelig;")
                   'jq_sog_val = replace(jq_sog_val, "A6", "")
                   'jq_sog_val = replace(jq_sog_val, "%", "")
                   'jq_sog_val = replace(jq_sog_val, "oe", "�")

                   'jq_sog_val = Server.HTMLEncode(jq_sog_val)

                   'call utf_format(jq_sog_val)
                   'jq_sog_val = utf_formatTxt

                   'call htmlparseCSV(HTMLstring)
                   'jq_sog_val = htmlparseCSVTxt 

                   '*** ��� **'
                   'call jq_format(jq_sog_val)
                   'jq_sog_val = jq_formatTxt
                   jq_sog_val = request("sog_val")

                  
              else
              jq_sog_val = ""
              end if

             'if session("mid") = 1 then
             '  jq_sog_val = "U+00E6"
             'end if
                

              if jq_sog_val <> "" then
              jq_sog_valSQL = " AND (Kkundenavn LIKE '%"& jq_sog_val &"%' OR Kkundenr LIKE '"& jq_sog_val &"%')"
              else
              jq_sog_valSQL = ""
              end if

              'jq_sog_valSQL = " AND (Kkundenavn LIKE 'p%')"

                strSQL = "SELECT Kkundenavn, Kkundenr, Kid, kundeans1, kundeans2 FROM kunder WHERE ketype <> 'e' AND (useasfak = 0 OR useasfak = 1 OR useasfak = 5) "& jq_sog_valSQL &" GROUP BY kid ORDER BY Kkundenavn LIMIT 500"
			    
                'if session("mid") = 1 then
                'Response.write "<option>"& jq_sog_valSQL &" sQL:_ "& strSQL & "</option>"
                'response.write(Server.HTMLEncode("The image tag: <img>"))
                'Response.end
                'end if
                
                z = 0
                oRec.open strSQL, oConn, 3
			    kans1 = ""
			    kans2 = ""
			    while not oRec.EOF


                if cdbl(oRec("kundeans1")) <> 0 then

            	strSQL2 = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = "&oRec("kundeans1")
				oRec2.open strSQL2, oConn, 3 
				if not oRec2.EOF then
				kans1 = oRec2("mnavn")
				end if
				oRec2.close

                end if
				

                if cdbl(oRec("kundeans2")) <> 0 then
				strSQL2 = "SELECT mnavn, mnr FROM medarbejdere WHERE mid = "&oRec("kundeans2")
				oRec2.open strSQL2, oConn, 3 
				if not oRec2.EOF then
				kans2 = " - &nbsp;&nbsp;" & oRec2("mnavn") 
				end if
				oRec2.close

                end if
				
				if len(kans1) <> 0 OR len(kans2) <> 0 then
				anstxt = " --- "& job_txt_001 &": "
				else
				anstxt = ""
				end if
				
			'strKlistTxt = strKlistTxt & "<option value='"&oRec("Kid")&"'>"&oRec("Kkundenavn")&"&nbsp;&nbsp;("&oRec("Kkundenr")&") " & anstxt &""& kans1 &"&nbsp;&nbsp;"& kans2 &"</option>"
            kundeArr_2013(z) = "<option value='"&oRec("Kid")&"'>"&oRec("Kkundenavn")&"&nbsp;&nbsp;("&oRec("Kkundenr")&") " & anstxt &""& kans1 &"&nbsp;&nbsp;"& kans2 &"</option>"
            

            'Response.flush
		    
			z = z + 1
			kans1 = ""
			kans2 = ""
			oRec.movenext
			wend
			oRec.close

              
              for z = 0 to z - 1
              '*** ��� **'
              call jq_format(kundeArr_2013(z))
              kundeArr_2013(z) = jq_formatTxt
	          
	          Response.Write kundeArr_2013(z)
              next


case "FN_getKundeKperslisten"
              
              dim kundeKpersArr
              redim kundeKpersArr(500)
              
              if len(trim(request.Form("cust2"))) <> 0 AND request.Form("cust2") <> 0 then
              kidThis = request.Form("cust2")
              else
              kidThis = -1
              end if

              if len(trim(request("jq_kundekpers"))) <> 0 then
              jq_kundekpers = request("jq_kundekpers")
              else
              jq_kundekpers = 0
              end if


                strSQL = "SELECT navn, id, titel, email FROM kontaktpers WHERE kundeid = "& kidThis
			    
                'Response.write "<option>"& strSQL& "</option>"
                'Response.end
                

                kundeKpersArr(0) = "<option value='0'>"&job_txt_002&"</option>"
                z = 1

                oRec.open strSQL, oConn, 3
			   
			    while not oRec.EOF
				
			    
                if cint(jq_kundekpers) = oRec("id") then
                kpsSEL = "SELECTED"
                else
                kpsSEL = ""
                end if

                if len(trim(oRec("email"))) <> 0 then
                eTxt = ", " & oRec("email")
                else
                eTxt = ""
                end if


                if len(trim(oRec("titel"))) <> 0 then
                tTxt = oRec("titel")
                else
                tTxt = ""
                end if
                

            	
			kundeKpersArr(z) = "<option value='"&oRec("id")&"' "&kpsSEL&">"&oRec("navn")&" "& tTxt &" "& eTxt &"</option>"
            

            'Response.flush
		    
			z = z + 1
			
			oRec.movenext
			wend
			oRec.close

            kundeKpersArr(z) = "<option value='0'>"&job_txt_003&"</option>"
              
               for z = 0 to UBOUND(kundeKpersArr)
              '*** ��� **'
              call jq_format(kundeKpersArr(z))
              kundeKpersArr(z) = jq_formatTxt
	          
	          Response.Write kundeKpersArr(z)
              next
		


end select
end select
Response.end
end if



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** S�tter lokal dato/kr format. Skal inds�ttes efter kalender.
	Session.LCID = 1030
	
    '** NOT READY ***'
    if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> "0" then                       
	nomenu = 1
    else
    nomenu = 0
    end if

    nomenu = 0
    '*********************


	function SQLBless3(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ".", ",")
		SQLBless3 = tmp
	end function
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	function SQLBless2(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless2 = tmp
	end function
	

    select case lto
         case "intranet - local", "epi2017"
            maxCharJobNavn = 50
         case else
            maxCharJobNavn = 100
         end select

	
	'*** Altid = 1 for en sikkedrheds skyld, men den bruges ikke mere pr. 1/10-2005 ***
	'*** Kun eksterne job herefter *** 
	strInternt = 1 ' = request("int") 
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if

    tp_jobid = id

    response.cookies("2015")("lastjobid") = id
	
    '**** Hvilken div skal vises? ****
	if len(request("showdiv")) <> 0 then
	showdiv = request("showdiv")
	else
	showdiv = "stamdata"
	end if

    if len(trim(request("FM_tp_jobaktid"))) <> 0 then
    tp_jobaktid = request("FM_tp_jobaktid")
    else
    tp_jobaktid = 0
    end if


    if len(trim(request("FM_hd_tp_jobaktid"))) <> 0 then
    hd_tp_jobaktid = request("FM_hd_tp_jobaktid")
    else
    hd_tp_jobaktid = 0
    end if


    if len(trim(request("FM_mtype"))) <> 0 then 'AND request("FM_mtype") <> 0 
    mtype = request("FM_mtype")
    else
        select case lto
        case "epi", "epi_no", "epi_sta", "epi_ab", "epi_cati", "epi_uk", "epi2017"
        mtype = 10
        case else
        mtype = 0
        end select
    
    end if
    
    if mtype <> 0 then
    mtypeSQLKri = " medarbejdertype = " & mtype
    else
    mtypeSQLKri = " medarbejdertype <> 0"
    end if
    

    if len(trim(request("FM_sorter_tp"))) <> 0 then
    sorttp = request("FM_sorter_tp")
        if sorttp = "1" then
        sorttpCHK0 = ""
        sorttpCHK1 = "CHECKED"
        else
        sorttpCHK0 = "CHECKED"
        sorttpCHK1 = ""
        end if
    else
        
        if request.cookies("tsa")("sorttp") <> "" then
            

        if request.cookies("tsa")("sorttp") = "1" then
        sorttp = 1
        sorttpCHK0 = ""
        sorttpCHK1 = "CHECKED"
        else
        sorttp = 0
        sorttpCHK0 = "CHECKED"
        sorttpCHK1 = ""
        end if


        else
        sorttp = 0
        sorttpCHK0 = "CHECKED"
        sorttpCHK1 = ""
        end if
    end if

    response.cookies("tsa")("sorttp") = sorttp


    if len(trim(request("FM_fomr"))) <> 0 then
    strFomr_rel = request("FM_fomr")
    strFomr_rel = replace(strFomr_rel, "X234", "#")
    else
    strFomr_rel = "#0#"
    end if

   
    

                    '*** Forretningsomr�der **' 
	                strFomr_Gblnavn = ""
                    strFomr_relA = replace(strFomr_rel, "#", "")
                    strFomr_relA = split(strFomr_relA, ",")

                    fo = 0
                    for f = 0 to UBOUND(strFomr_relA)

                    strSQLfrel = "SELECT fomr.navn FROM fomr WHERE fomr.id = "& strFomr_relA(f) 

                    'Response.Write strSQLfrel
                    'Response.flush
                    
                    oRec3.open strSQLfrel, oConn, 3
                    if not oRec3.EOF then

                    if fo = 0 then
                    strFomr_Gblnavn = " ("
                    end if

                    strFomr_Gblnavn = strFomr_Gblnavn & oRec3("navn") & ", " 
                   
                    fo = fo + 1
                    end if
                    oRec3.close

                    next

                    if fo <> 0 then
                    len_strFomr_Gblnavn = len(strFomr_Gblnavn)
                    left_strFomr_Gblnavn = left(strFomr_Gblnavn, len_strFomr_Gblnavn - 2)
                    strFomr_Gblnavn = left_strFomr_Gblnavn & ")"

                    if len(strFomr_Gblnavn) > 50 then
                    strFomr_Gblnavn = left(strFomr_Gblnavn, 50) & "..)"
                    end if
                    end if		    

	
	thisfile = "jobs"
	rdir = request("rdir")
	
	
	select case func
	case "slet"
	'*** Her sp�rges om det er ok at der slettes en medarbejder ***
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<%
	
	slttxt = "<b>A) "&job_txt_004&"</b><br />"_
	& job_txt_005 & "<b>"&job_txt_006&"<br> "&job_txt_007&"</b>"&job_txt_008
	slturl = "jobs.asp?menu=job&func=sletok&id="&id&"&kt=0&fm_kunde_sog="&request("fm_kunde_sog")&"&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")&"&rdir="&rdir
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,110,90)
	
	slttxtb = "<b>B) "&job_txt_009&"</b><br>"&job_txt_010&"<br> "&job_txt_011
    slturlb = "jobs.asp?menu=job&func=sletok&id="&id&"&kt=1&fm_kunde_sog="&request("fm_kunde_sog")&"&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")&"&rdir="&rdir
	
	
	call sltque(slturlb,slttxtb,slturlalt,slttxtalt,540,90)
	
	slttxtc = "<b>C) "&job_txt_012&"</b><br>"&job_txt_013&"<br>"&job_txt_14
    slturlc = "jobs.asp?menu=job&func=sletok&id="&id&"&kt=2&fm_kunde_sog="&request("fm_kunde_sog")&"&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")&"&rdir="&rdir
	
	
	call sltque(slturlc,slttxtc,slturlalt,slttxtalt,970,90)
	
	
	
	case "sletok"
	'*** Her slettes et job, dets aktiviteter og de indtastede timer p� jobbet ***
	strSQL = "SELECT id, navn FROM aktiviteter WHERE job = "& id &"" 
	kt = request("kt") 'slet kun timer
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
		
		'*** Inds�tter i delete historik ****'
	    call insertDelhist("akt", oRec("id"), 0, oRec("navn"), session("mid"), session("user"))
		
		'** Slet eller nulstil? **'
		if kt = "0" OR kt = "2" then
		oConn.execute("DELETE FROM aktiviteter WHERE id = "& oRec("id") &"")
		end if
		oConn.execute("DELETE FROM timer WHERE TAktivitetId = "& oRec("id") &"")
		
	oRec.movenext
	wend
	oRec.close
	
	'*** Sletter ressource timer p� job ****
	oConn.execute("DELETE FROM ressourcer WHERE jobid = "& id &"")
	
	'*** Sletter ressource_md timer p� job ****
	oConn.execute("DELETE FROM ressourcer_md WHERE jobid = "& id &"")
	
	'*** Sletter fra Guiden aktive job timer p� job ****
	oConn.execute("DELETE FROM timereg_usejob WHERE jobid = "& id &"")
	
	'*** Sletter materialeforbrug ****
	oConn.execute("DELETE FROM materiale_forbrug WHERE jobid = "& id &"")

	
	'*** Sletter fakturaer p� job ****'
	'*** Kan ikke slette job der findes fakturaer p� pr. 16/11-2008 **'
	
	'strSQLfak = "SELECT fid FROM fakturaer WHERE jobid = "& id &""
	'oRec.open strSQLfak, oConn, 3
	
	'while not oRec.EOF 
	        
	'        oConn.execute("DELETE FROM faktura_det WHERE fakid = "& oRec("fid") &"")
	'        oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& oRec("fid") &"")
	'        oConn.execute("DELETE FROM fak_mat_spec WHERE matfakid = "& oRec("fid") &"")
	        
	'oRec.movenext
	'wend
	'oRec.close
	
	
	'oConn.execute("DELETE FROM fakturaer WHERE jobid = "& id &"")
	
	
	strSQL = "SELECT id, jobnavn, jobnr FROM job WHERE id = "& id &"" 
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
		
		'*** Inds�tter i delete historik ****'
	    call insertDelhist("job", id, oRec("jobnr"), oRec("jobnavn"), session("mid"), session("user"))
		
	
	end if
	oRec.close
	
	
	
	
	'Response.flush
	
	if kt = "0" then
	oConn.execute("DELETE FROM job WHERE id = "& id &"")
	end if
	
	if rdir <> "webblik" then
	'Response.redirect "jobs.asp?menu=job&shokselector=1&fm_kunde="&request("fm_kunde_sog")&"&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")
    Response.redirect "../to_2015/job_list.asp?hidelist=1"
	else
	Response.redirect "../to_2015/webblik_joblisten.asp"
	end if
	
	case "sletfil"
	
	'*** Her sp�rges om det er ok at der slettes en fil ***
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190px; top:150px; visibility:visible;">
	<br><br><br>
	<table cellspacing="2" cellpadding="2" border="0">
	<tr>
	    <td><img src="../ill/alert.gif" width="44" height="45" alt="" border="0"><%=job_txt_015 %> <b><%=job_txt_016 %></b> <%=job_txt_017 %></td>
	</tr>
	<tr>
	   <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	   <a href="jobs.asp?menu=jobs&func=sletfilok&id=<%=id%>&filnavn=<%=request("filnavn")%>&jobid=<%=request("jobid")%>"><%=job_txt_018 %></a>&nbsp;&nbsp;&nbsp;<a href="javascript:history.back()"><%=job_txt_019 %></a></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	case "sletfilok"
	'*** Her slettes en fil ***
	'ktv
	'strPath =  "E:\www\timeout_xp\wwwroot\ver2_1\upload\"&lto&"\" & Request("filnavn")
	'Qwert
	strPath =  "d:\dkdomains\outzource\wwwroot\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\" & Request("filnavn")
	'Response.write strPath
	
	on Error resume Next 

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set fsoFile = FSO.GetFile(strPath)
	fsoFile.Delete
	
	
	oConn.execute("DELETE FROM filer WHERE id = "& id &"")
	Response.redirect "jobs.asp?menu=job&shokselector=1&func=red&id="& request("jobid")&"&int=1"
	
	

    '************************************************************'
    '************** Opdater jobliste ****************************'
	'************************************************************'
    case "opdjobliste"
	
	'Response.write "Opdaterer liste"
	
    if len(trim(request("FM_formr_opdater"))) <> 0 then
    fomrJa = request("FM_formr_opdater")
    else
    fomrJa = 0
    end if

    if len(trim(request("FM_formr_opdater_akt"))) <> 0 then
    fomrJaAkt = request("FM_formr_opdater_akt")
    else
    fomrJaAkt = 0
    end if

    

    slutdato = request("FM_opdaterslutdato_dato")

    if len(trim(request("FM_opdaterslutdato"))) <> 0 then
    opdaterslutdato = 1
    else
    opdaterslutdato = 0
    end if

    slutdato = replace(slutdato, ".", "-")

    'Response.Write slutdato & IsDate('slutdato)
    if IsDate(slutdato) = True then

    'Response.Write "<br>" & slutdato
    slutdatoSQL = year(slutdato) &"/"& month(slutdato) & "/"& day(slutdato) 
    else

    opdaterslutdato = 0

    end if


   

    fomrArr = split(request("FM_fomr"), ",")

    if len(trim(request("FM_tilfojeasyreg"))) <> 0 then
    easyReg = request("FM_tilfojeasyreg")
    else
    easyReg = 0
    end if

                    'for_faktor = 0
                    'for afor = 0 to UBOUND(fomrArr)
                    'for_faktor = for_faktor + 1 
                    'next

                    'if for_faktor <> 0 then
                    'for_faktor = for_faktor
                    'else
                    for_faktor = 1
                    'end if

                    'for_faktor = formatnumber(100/for_faktor, 2)
                    'for_faktor = replace(for_faktor, ",", ".")

	jobids = split(request("FM_listeid"), ",")
	'jobstatus = split(request("FM_listestatus"), ",")
	for t = 0 to UBOUND(jobids)
		

        jobStatus = request("FM_listestatus_"& trim(jobids(t)) &"")
        'Response.Write "FM_listestatus_"& trim(jobids(t)) &"" & "<br>"

         '*** S�t lukkedato (skal v�re f�r det skifter status) ***'
        call lukkedato(jobids(t), jobStatus)

		strSQL = "UPDATE job SET jobstatus = " & jobStatus 
		
        
        if opdaterslutdato = 1 then
        strSQL = strSQL & ", jobslutdato = '"& slutdatoSQL & "'"
        end if


        strSQL = strSQL & " WHERE id = "& jobids(t)

        oConn.execute(strSQL)

        '**** SyncDatoer ****'
        jobnrThis = 0
        syncslutdato = 0
        strSQLj = "SELECT jobnr, syncslutdato FROM job WHERE id =  "& jobids(t)
        oRec5.open strSQLj, oConn, 3 
        if not oRec5.EOF then

        jobnrThis = oRec5("jobnr")
        syncslutdato = oRec5("syncslutdato")

        end if
        oRec5.close

        if jobStatus = 0 AND syncslutdato = 1 then
            call syncJobSlutDato(jobids(t), jobnrThis, syncslutdato)
        end if


       

        


		'Response.write strSQL & "<br>"

        if cint(easyReg) <> 0 then
                            
                            
        if cint(easyReg) = 1 then
        easyRegVal = 1
        else
        easyRegVal = 0
        end if

                            
                            '**** S�tter EASYreg aktiv p� aktiviteter for alle medarbejdere (hvis der findes EASY regaktiviter) ****'
                           
                            strSQLea = "UPDATE aktiviteter SET easyreg = "& easyRegVal &" WHERE job = "& jobids(t)
				            oConn.execute(strSQLea)
				                
                            if cint(easyReg) = 1 then
                            strSQLtreguse = "UPDATE timereg_usejob SET easyreg = " & jobids(t) & " WHERE jobid = " & jobids(t)
                            else '= 2
                            strSQLtreguse = "UPDATE timereg_usejob SET easyreg = 0 WHERE jobid = " & jobids(t)
                            end if

	   	                    oConn.execute(strSQLtreguse)
				              


        end if


        '*** Forretningsomr�der ****'
	    if cint(fomrJa) = 1 then
                    

                    '*** nulstiller job (IKKE akt) ****'
                    strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& jobids(t)  & " AND for_aktid = 0"
                    oConn.execute(strSQLfor)

                    'Response.Write "her"
                    'Response.Flush
                    'fomrArrLink = "#0#"
                    for afor = 0 to UBOUND(fomrArr)

                            'Response.Write "her2" & afor & "<br>"
                            'Response.Flush

                            fomrArr(afor) = replace(fomrArr(afor), "#", "")

                            if fomrArr(afor) <> 0 then

                            strSQLfomri = "INSERT INTO fomr_rel "_
                            &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                            &" VALUES ("& fomrArr(afor) &", "& jobids(t) &", 0, "& for_faktor &")"

                            oConn.execute(strSQLfomri)

                            'fomrArrLink = fomrArrLink & ",#"& fomrArr(afor) & "#" 
            

                                '**** IKKE MERE 11.3.2015 *****'
                                '*** Altid Sync aktiviteter ved multiopdater ****'
                                if cint(fomrJaAkt) = 1 then
                                strSQLa = "SELECT id FROM aktiviteter WHERE job = "& jobids(t)
                                oRec3.open strSQLa, oConn, 3
                                while not oRec3.EOF 


                                     '*** nulstiller form p� akt () ****'
                                    strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& jobids(t)  & " AND for_aktid = " & oRec3("id")
                                    oConn.execute(strSQLfor)

                                    strSQLfomrai = "INSERT INTO fomr_rel "_
                                    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                    &" VALUES ("& fomrArr(afor) &", "& jobids(t) &","& oRec3("id") &", "& for_faktor &")"

                                   oConn.execute(strSQLfomrai)

                                oRec3.movenext
                                wend
                                oRec3.close
                                end if

                                

                            end if


                    next

                    '********************************'



        end if

	next
	'response.write "jobs.asp?menu=job&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")&"&fm_kunde_sog="&request("fmkunde")&"&FM_kunde="&request("FM_kunde")&"&FM_fomr="&request("FM_fomr")
    'Response.end
	
    formTemp = replace(request("FM_fomr"), "#", "X234")

    Response.redirect "jobs.asp?menu=job&jobnr_sog="&request("jobnr_sog")&"&filt="&request("filt")&"&fm_kunde_sog="&request("fmkunde")&"&FM_kunde="&request("FM_kunde")&"&FM_fomr="&formTemp
	
	
	
	
	case "dbopr", "dbred"
	'*** Her inds�ttes et nyt job i db ****
	
	call jobopr_mandatory_fn()
    showaspopup = request("showaspopup")
	
	public useleftdiv
	sub visErrorFormat
	
		if showaspopup <> "y" then
		%>
		<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
		
		<%
		else
		%>
		<!--#include file="../inc/regular/header_hvd_inc.asp"-->
		<%
		useleftdiv = "j"
		end if
	
	end sub
	
	
				 '*tjekker om dato eksisterer og tildeler ny dato hvis den ikke g�r ** 
				call dato_30(Request("FM_start_dag"), Request("FM_start_mrd"), Request("FM_start_aar"))
				strStartDay = strDay_30
				
				call dato_30(Request("FM_slut_dag"), Request("FM_slut_mrd"), Request("FM_slut_aar"))
				strSlutDay = strDay_30
	
		
		slutaar = right(Request("FM_slut_aar"), 2)
		if slutaar = "44" then
		slutaar = "2044"
		else
		slutaar = slutaar
		end if
		
		startDatoNum = cdate(strStartDay &"/"& Request("FM_start_mrd") &"/"& Request("FM_start_aar")) 
		slutDatoNum = cdate(strSlutDay &"/"& Request("FM_slut_mrd")  &"/"& slutaar)
		
	

		'** Tjekker om alle felter er udfyldt korrekt **
		if len(request("FM_navn")) = 0 OR _
        ((len(request("FM_navn")) > 50 AND (instr(lto, "epi") <> 0 OR lto = "intranet - local")) OR (len(request("FM_navn")) > 100 AND instr(lto, "epi") = 0)) OR _
        len(trim(request("FM_jnr"))) = 0 OR len(request("FM_kunde")) = 0 OR _
		len(request("FM_jnr")) > 20 OR _
		cint(instr(request("FM_navn"), "'")) > 0 OR _
        ((lcase(request("FM_navn")) = "jobnavn.." AND func = "dbred" AND instr(lto, "epi") <> 0) OR _
        (lcase(request("FM_navn")) = "jobnavn.." AND func = "dbred" AND instr(lto, "epi") = 0)) then
		
		call visErrorFormat
		errortype = 14
		call showError(errortype)
		Response.end
		
		else
		
            	'** Tjekker om SLUTDATO udfyldt korrekt ** EPI pga auto opret og startdato starter i minus 1
		        if (startDatoNum > slutDatoNum AND instr(lto, "epi") = 0) OR _
                (startDatoNum > slutDatoNum AND func = "dbred" AND instr(lto, "epi") <> 0) then
		
		        call visErrorFormat
		        errortype = 184
		        call showError(errortype)
		        Response.end
                end if
		
		    %>
			<!--#include file="inc/isint_func.asp"-->
			<%
			
			isInt = 0
			call erDetInt(request("prio")) 
			if isInt > 0 then
			    
			    call visErrorFormat
				
				errortype = 132
				call showError(errortype)
			
			response.End
			end if
			
            if cint(budget_mandatoryOn) = 1 then 'Budget  / InternL�nOmk. mandatory

                if len(trim(request("FM_interntbelob"))) <> 0 then
                interntbelobTjk = request("FM_interntbelob")
                else
                interntbelobTjk = 0
                end if
			
                isInt = 0    
		        call erDetInt(interntbelobTjk)
                if (isInt > 0 OR cdbl(interntbelobTjk) < 100) AND func = "dbred" then 
                call visErrorFormat
    		
		        errortype = 182
		        call showError(errortype)
		        response.End
			    end if

             


            end if
			

            '*********************************************************************'
				'************************** Internt / eksternt job *******************'
				'*********************************************************************'
				'*** Altid Eksternt job ***'
						
						if request("FM_usetilbudsnr") = "j" then
						strStatus = 3 'tilbud
						else
							if request("FM_status") <> 3 then
							strStatus = request("FM_status")
							else
							strStatus = 1 'aktivt, hvis der ved en fejl at valgt tilbud
							end if
						end if



            '*********  Sandsynlighed ***********************
            if len(trim(request("FM_sandsynlighed"))) <> 0 then
			sandsynlighed = formatnumber(request("FM_sandsynlighed"), 0)
			else
			sandsynlighed = 0
			end if

			isInt = 0    
		    call erDetInt(request("FM_sandsynlighed"))
            if isInt > 0 OR (lto = "epi2017" AND strStatus = 3 AND (len(trim(sandsynlighed ) ) = 0 OR sandsynlighed > 100 OR sandsynlighed < 1 )) then
            call visErrorFormat
    		
		    errortype = 31
		    call showError(errortype)
		    response.End
			end if


            '** forretningsomr�de mandatory 
            call jobopr_mandatory_fn()
            if cint(fomr_mandatoryOn) = 1 AND len(trim(request("FM_fomr"))) = 0 then

            call visErrorFormat
			errortype = 177
			call showError(errortype)
			Response.end


            end if
			
			
			'Response.Write  request("FM_jnr")
			'Response.flush
			        
                    '*** Jobnr og tilbudsnummer m� gerne indeholde nr. **'
                    isInt = 0
			        'call erDetInt(request("FM_jnr")) 
			        'if cint(isInt) > 0 OR instr(request("FM_jnr"), ".") <> 0 OR (request("FM_jnr") = "0" AND func = "dbred") then
				
				     '   call visErrorFormat
				
				     '   errortype = 15
				     '   call showError(errortype)
			
			        'isInt = 0
			
			        'else
			                'call erDetInt(request("FM_tnr"))
			                'if isInt > 0 then
				
				            'call visErrorFormat
				
				            'errortype = 15
				            'call showError(errortype)
			
			                'isInt = 0
			
			                'else
			    
			    if request("simpeludv") = "1" then
			    simpeludvEXT = "_s"
			    response.cookies("tsa")("be_simpel_udv") = "1"
			    else
			    simpeludvEXT = ""
			    response.cookies("tsa")("be_simpel_udv") = "2"
			    end if
			    
			    
				call erDetInt(request("FM_budget"&simpeludvEXT&""))
				if isInt > 0 OR trim(lcase(request("FM_budget"&simpeludvEXT&""))) = "nan" then 'OR instr(request("FM_budget"&simpeludvEXT&""), "-") <> 0 20200114
				
				call visErrorFormat
				
				errortype = 16
				call showError(errortype)
				isInt = 0
						
						else
						
						
						call erDetInt(request("FM_budgettimer"&simpeludvEXT&""))
						if isInt > 0 then
						
						call visErrorFormat
						
						errortype = 20
						call showError(errortype)
						isInt = 0
								
								else
								
								call erDetInt(request("FM_ikkebudgettimer"&simpeludvEXT&""))
								if isInt > 0 then
								
								call visErrorFormat
								
								errortype = 20
								call showError(errortype)
								isInt = 0
								
								else
								
								
								if request("FM_fastpris") = "1" AND len(trim(request("FM_budgettimer"))) = 0 then 'OR request("FM_fastpris") = "1" AND request("FM_budgettimer") = "0" then
									
									call visErrorFormat
									
									errortype = 30
									call showError(errortype)
									
								
								
								
				                
				  else
				
				
				'**** Validering OK ***
				'** kundeid og navn *** 
				
				
				strKundeId = request("FM_kunde")
				varJobId = 0
				
				'Response.Write strKundeid
				'Response.end
				
				'*** Multible opret ***
				if instr(request("FM_kunde"), ",") <> 0 then
				    
				    kids = split(request("FM_kunde"), ",")
				    for x = 0 to UBOUND(kids)
				        
				        if kids(x) <> 0 then
				        kidSQLKri = kidSQLKri & " kid = "& kids(x) & " OR"
				        end if
				        
				        if x = 1 then
				        strKidUseOne = kids(x)
				        end if 
				    next
				    
				    len_kidSQLKri = len(kidSQLKri)
				    left_kidSQLKri = left(kidSQLKri, len_kidSQLKri - 3)
				    kidSQLKri = " ("& left_kidSQLKri &") "
				    
				else
			    kidSQLKri = " kid = "& strKundeId 
			    strKidUseOne = strKundeId
				end if
				
				
				
				

				
				
			    '**********************************
			    '**** Jobdata ****
			    '**********************************
			    
				
				strNavn = replace(request("FM_navn"),chr(34), "")
				
				'** Jobbeskrivelse **'
			    strBesk = SQLBless2(request("FM_beskrivelse"))
			    
			    '*** HTML Replace **'
                call htmlreplace(strBesk)
			    strBesk = htmlparseTxt
			    
			    '* Intern note **'
			    strInternBesk = SQLBless2(request("FM_internbesk"))
			    
			    '*** HTML Replace **'
			    call htmlreplace(strInternBesk)
			    strInternBesk = htmlparseTxt
			    
			    
			    
			    if len(request("FM_ski")) <> 0 then
			    ski = 1
			    else
			    ski = 0
			    end if
			    
			    if len(request("FM_abo")) <> 0 then
			    abo = 1
			    else
			    abo = 0
			    end if
			    
			    if len(request("FM_ubv")) <> 0 then
			    ubv = 1
			    else
			    ubv = 0
			    end if
			    
			    
			   
                			    
			    
			    
				strOLDjobnr = request("FM_OLDjobnr")
				

                '***** Jobansvarlige ***'
                if len(trim(request("FM_jobans_1"))) <> 0 then
				intJobans1 = request("FM_jobans_1")
                else
                intJobans1 = 0
                end if

                 if len(trim(request("FM_jobans_2"))) <> 0 then
				intJobans2 = request("FM_jobans_2")
                else
                intJobans2 = 0
                end if

                 if len(trim(request("FM_jobans_3"))) <> 0 then
				intJobans3 = request("FM_jobans_3")
                else
                intJobans3 = 0
                end if

                 if len(trim(request("FM_jobans_4"))) <> 0 then
				intJobans4 = request("FM_jobans_4")
                else
                intJobans4 = 0
                end if

                 if len(trim(request("FM_jobans_5"))) <> 0 then
				intJobans5 = request("FM_jobans_5")
                else
                intJobans5 = 0
                end if

				

                if len(trim(request("FM_jobans_proc_1"))) <> 0 then
				jobans_proc_1 = request("FM_jobans_proc_1")
                else
                jobans_proc_1 = 0
                end if

                

				if len(trim(request("FM_jobans_proc_2"))) <> 0 then
				jobans_proc_2 = request("FM_jobans_proc_2")
                else
                jobans_proc_2 = 0
                end if

                

				if len(trim(request("FM_jobans_proc_3"))) <> 0 then
				jobans_proc_3 = request("FM_jobans_proc_3")
                else
                jobans_proc_3 = 0
                end if

               
				if len(trim(request("FM_jobans_proc_4"))) <> 0 then
				jobans_proc_4 = request("FM_jobans_proc_4")
                else
                jobans_proc_4 = 0
                end if

               

				if len(trim(request("FM_jobans_proc_5"))) <> 0 then
				jobans_proc_5 = request("FM_jobans_proc_5")
                else
                jobans_proc_5 = 0
                end if

               

                errProcVal = 0

                for i = 1 to 5
                      select case i 
                      case 1
                      procVal = jobans_proc_1
                      case 2
                      procVal = jobans_proc_2
                      case 3
                      procVal = jobans_proc_3
                      case 4
                      procVal = jobans_proc_4
                      case 5
                      procVal = jobans_proc_5
                      end select
                    
                   
                    
                    call erDetInt(procVal)
				    
                    if isInt > 0 then
                    errProcVal = 1
                    isInt = 0

                    else

                    if instr(lto, "epi2017") <> 0 OR lto = "intranet - local" then 
                    jobProcent100 = jobProcent100/1 + procVal 'Skal divideres med 10 da der fjernes komma
                    end if  

                    end if
                
                next

                jobans_proc_1 = replace(jobans_proc_1, ".", "")
                jobans_proc_1 = replace(jobans_proc_1, ",", ".")
                jobans_proc_2 = replace(jobans_proc_2, ".", "")
                jobans_proc_2 = replace(jobans_proc_2, ",", ".")
                 jobans_proc_3 = replace(jobans_proc_3, ".", "")
                jobans_proc_3 = replace(jobans_proc_3, ",", ".")
                jobans_proc_4 = replace(jobans_proc_4, ".", "")
                jobans_proc_4 = replace(jobans_proc_4, ",", ".")
                jobans_proc_5 = replace(jobans_proc_5, ".", "")
                jobans_proc_5 = replace(jobans_proc_5, ",", ".")
                '***** Jobanavarlige *****


                '***** Salg ansvarlige ******
                if len(trim(request("FM_salgsans_1"))) <> 0 then ' er slags ansvarlige sl�et til / vist
                salgsans1 = request("FM_salgsans_1")

                if len(trim(request("FM_salgsans_2"))) <> 0 then
				salgsans2 = request("FM_salgsans_2")
                else
                salgsans2 = 0
                end if

				if len(trim(request("FM_salgsans_3"))) <> 0 then
				salgsans3 = request("FM_salgsans_3")
                else
                salgsans3 = 0
                end if

				if len(trim(request("FM_salgsans_4"))) <> 0 then
				salgsans4 = request("FM_salgsans_4")
                else
                salgsans4 = 0
                end if

				if len(trim(request("FM_salgsans_5"))) <> 0 then
				salgsans5 = request("FM_salgsans_5")
                else
                salgsans5 = 0
                end if


                else
                salgsans1 = 0
				salgsans2 = 0
				salgsans3 = 0
				salgsans4 = 0
				salgsans5 = 0
                end if

                if len(trim(request("FM_salgsans_proc_1"))) <> 0 then
				salgsans_proc_1 = request("FM_salgsans_proc_1")
                else
                salgsans_proc_1 = 0
                end if

               

				if len(trim(request("FM_salgsans_proc_2"))) <> 0 then
				salgsans_proc_2 = request("FM_salgsans_proc_2")
                else
                salgsans_proc_2 = 0
                end if

               


				if len(trim(request("FM_salgsans_proc_3"))) <> 0 then
				salgsans_proc_3 = request("FM_salgsans_proc_3")
                else
                salgsans_proc_3 = 0
                end if

                
               
                if len(trim(request("FM_salgsans_proc_4"))) <> 0 then
				salgsans_proc_4 = request("FM_salgsans_proc_4")
                else
                salgsans_proc_4 = 0
                end if

              
                

				if len(trim(request("FM_salgsans_proc_5"))) <> 0 then
				salgsans_proc_5 = request("FM_salgsans_proc_5")
                else
                salgsans_proc_5 = 0
                end if

               
                
                
           

                for i = 1 to 5
                      select case i 
                      case 1
                      procVal = salgsans_proc_1
                      case 2
                      procVal = salgsans_proc_2
                      case 3
                      procVal = salgsans_proc_3
                      case 4
                      procVal = salgsans_proc_4
                      case 5
                      procVal = salgsans_proc_5
                      end select
                    

                      

                    call erDetInt(procVal)
				    
                    if isInt > 0 then
                    errProcVal = 1
                    isInt = 0
                    else

                    if instr(lto, "epi2017") <> 0 OR lto = "xintranet - local"  then 
                    salgsProcent100 = salgsProcent100/1 + procVal 'Skal divideres med 10 da der fjernes komma
                    end if 

                    end if
                
                next
                
                
                salgsans_proc_1 = replace(salgsans_proc_1, ".", "")
                salgsans_proc_2 = replace(salgsans_proc_2, ".", "")
                salgsans_proc_3 = replace(salgsans_proc_3, ".", "")
                salgsans_proc_4 = replace(salgsans_proc_4, ".", "")
                salgsans_proc_5 = replace(salgsans_proc_5, ".", "")

                salgsans_proc_1 = replace(salgsans_proc_1, ",", ".")
                salgsans_proc_2 = replace(salgsans_proc_2, ",", ".")
                salgsans_proc_3 = replace(salgsans_proc_3, ",", ".")
                salgsans_proc_4 = replace(salgsans_proc_4, ",", ".")
                salgsans_proc_5 = replace(salgsans_proc_5, ",", ".")

                
                
                '***** Salg ansvarlige SLUT
			 
               
				
                if len(request("FM_virksomheds_proc")) <> 0 then
                virksomheds_proc = request("FM_virksomheds_proc")
                else
                virksomheds_proc = 0
                end if 

                call erDetInt(virksomheds_proc)
                if isInt > 0 then
                errProcVal = 1
                isInt = 0
                end if



                if errProcVal <> 0 then
                call visErrorFormat
				errortype = 151
                call showError(errortype)

                Response.end
                end if


                if instr(lto, "epi2017") <> 0 OR lto = "xintranet - local" then
                
                    if (cdbl(salgsProcent100) <> 100 OR cdbl(jobProcent100) <> 100) AND func = "dbred" then
                
                            call visErrorFormat
				            errortype = 183
                            call showError(errortype)

                            Response.end

                
                    end if
                
                end if 


                'if len(trim(lcase(request("FM_budget"&simpeludvEXT&"")))) = "nan" then
                ''call visErrorFormat
				'errortype = 159
                'call showError(errortype)

                'Response.end
                'end if


				if len(trim(request("FM_budget"&simpeludvEXT&""))) = "0" then
				strBudget = 0
				else
				strBudget = request("FM_budget"&simpeludvEXT&"")
				strBudget = replace(strBudget, ".", "")
				strBudget = replace(strBudget, ",", ".")
				end if
				
				if len(trim(request("FM_fastpris"))) = "0" then 
				strFastpris = 0 '0: lbn. timer / 1: fastpris 
				else
				strFastpris = request("FM_fastpris")
				end if
				
				if len(trim(request("FM_jobfaktype"))) <> 0 then
				jobfaktype = request("FM_jobfaktype")
				else
				jobfaktype = 0
				end if
				
				
				if len(trim(request("prio"))) <> 0 then
				intprio = request("prio")
				else
 				intprio = -1
 				end if
 				
				response.Cookies("tsa")("faktype") = jobfaktype
				response.cookies("tsa").expires = date + 45
				
				
				usejoborakt_tp = request("FM_usejoborakt_tp")
				
				
				if len(trim(request("FM_budgettimer"&simpeludvEXT&""))) = 0 then
				strBudgettimer = 0
				else
				strBudgettimer = request("FM_budgettimer"&simpeludvEXT&"")
				strBudgettimer = replace(strBudgettimer, ".", "")
				strBudgettimer = replace(strBudgettimer, ",", ".")
				end if
				
				if len(trim(request("FM_ikkebudgettimer"&simpeludvEXT&""))) = "0" then
				ikkeBudgettimer = 0
				else
				ikkeBudgettimer = request("FM_ikkebudgettimer"&simpeludvEXT&"")
				end if
				
				strKunde = request("FM_kunde")
				
				
				'Response.Write "request(FM_kundese): " & request("FM_kundese")
				'Response.end
				'** Tilg�ngelig for kunde ***
				if len(request("FM_kundese")) <> 0 then
					if request("FM_kundese_hv") = 1 then
					intkundese = 2 '(n�r job er lukket)
					else
					intkundese = 1 '(n�r timer er indtastet)
					end if
				else
				intkundese = 0
				end if
				
				if len(request("FM_opr_kpers")) <> 0 then
				intKundekpers = request("FM_opr_kpers")
				else
				intKundekpers = 0
				end if
				
                
                if cint(intKundekpers) = 0 AND (lto = "dencker") AND rdir <> "redjobcontionue" then
		        call visErrorFormat
				errortype = 155
                call showError(errortype)
                Response.end
		        end if

				if len(request("FM_serviceaft")) <> 0 then
				intServiceaft = request("FM_serviceaft") '1
				else
				intServiceaft = 0
				end if
				
				
				
				strEditor = session("user")
				strDato = session("dato")

                ddupdate = year(now) & "/" & month(now) & "/" & day(now)


				intAktfavgp = "0, "& request("FM_favorit")
				strAktFase = replace("0, "& request("FM_favorit_fase"), "'", "")
				strAktFase = replace(strAktFase, "''", "")
				
				intKorsel = 0 'request("FM_favorit_korsel")
				
				if len(request("FM_lukafmjob")) <> 0 then
				lukafmjob = request("FM_lukafmjob")
				else
				lukafmjob = 0
				end if
				
				if len(request("FM_valuta")) <> 0 then
				valuta = request("FM_valuta")
				else
				valuta = 1 'main valuta
				end if
				
				rekvnr = replace(request("FM_rekvnr"), "'", "''")
				
				if len(trim(rekvnr)) = 0 AND lto = "ets-track" then
		        call visErrorFormat
				errortype = 149
                call showError(errortype)
                Response.end
		        end if

                if len(trim(request("FM_forvalgt"))) <> 0 then
                forvalgt = 1
                else
                forvalgt = 0
                end if
				
		        if len(trim(request("FM_bruttofortj"))) <> 0 AND lCase(request("FM_bruttofortj")) <> "nan" then
		        jo_bruttofortj = replace(request("FM_bruttofortj"), ".", "")
		        jo_bruttofortj = replace(jo_bruttofortj, ",", ".")
		        else
		        jo_bruttofortj = 0
		        end if
		        
		        call erdetint(jo_bruttofortj)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 126
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0

                
		        
		        if len(trim(request("FM_db"))) <> 0 AND lCase(request("FM_db")) <> "nan" then
		        
		        jo_dbproc = request("FM_db")
		         
		            call erdetint(jo_dbproc)
		            if isInt <> 0 then
		            call visErrorFormat
				    errortype = 127
                    call showError(errortype)
                    Response.end
                    else
                    jo_dbproc = formatnumber(request("FM_db"),0)
		            end if
		            isInt = 0
		            
				
				else
				jo_dbproc = 0
				end if
				
				
				
				if len(trim(request("FM_gnsinttpris"))) <> 0 AND lCase(request("FM_gnsinttpris")) <> "nan" then
				jo_gnstpris = replace(request("FM_gnsinttpris"), ".", "")
		        jo_gnstpris = replace(jo_gnstpris, ",", ".")
				else
				jo_gnstpris = 0
				end if
				
				call erdetint(jo_gnstpris)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 124
                call showError(errortype)
                Response.end
		        end if
				isInt = 0
				
				if len(trim(request("FM_intfaktor"))) <> 0 then
				jo_gnsfaktor = replace(request("FM_intfaktor"), ",", ".")
				else
				jo_gnsfaktor = 0
				end if
				
				call erdetint(jo_gnsfaktor)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 128
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0
				
                '** Diff p� timer og Sum job vs aktiviteter		
                diff_timer = request("FM_diff_timer")
                diff_sum = request("FM_diff_sum")

                diff_timer = replace(diff_timer, ".", "")
		        diff_timer = replace(diff_timer, ",", ".")

                diff_sum = replace(diff_sum, ".", "")
		        diff_sum = replace(diff_sum, ",", ".")

                if len(trim(diff_timer)) <> 0 then
                diff_timer = diff_timer
                else
                diff_timer = 0
                end if

                 if len(trim(diff_sum)) <> 0 then
                diff_sum = diff_sum
                else
                diff_sum = 0
                end if
 
                        
                if len(trim(request("FM_interntbelob"))) <> 0 AND lCase(request("FM_interntbelob")) <> "nan" then
                jo_gnsbelob = replace(request("FM_interntbelob"), ".", "")
		        jo_gnsbelob = replace(jo_gnsbelob, ",", ".")
                else
                jo_gnsbelob = 0
                end if
                
                call erdetint(jo_gnsbelob)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 125
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0

                if len(trim(request("FM_udgifter_ulev"))) <> 0 AND lCase(request("FM_udgifter_ulev")) <> "nan" then
                jo_udgifter_ulev = replace(request("FM_udgifter_ulev"), ".", "")
		        jo_udgifter_ulev = replace(jo_udgifter_ulev, ",", ".")
                else
                jo_udgifter_ulev = 0
                end if

                call erdetint(jo_udgifter_ulev)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 1252
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0


                if len(trim(request("FM_udgifter_intern"))) <> 0 AND lCase(request("FM_udgifter_intern")) <> "nan" then
                jo_udgifter_intern = replace(request("FM_udgifter_intern"), ".", "")
		        jo_udgifter_intern = replace(jo_udgifter_intern, ",", ".")
                else
                jo_udgifter_intern = 0
                end if

                call erdetint(jo_udgifter_intern)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 1251
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0


                
               
		        
		        if len(trim(request("FM_udgifter"))) <> 0 AND lCase(request("FM_udgifter")) <> "nan" then
		        udgifter = replace(request("FM_udgifter"), ".", "")
                udgifter = replace(udgifter, ",", ".")
                else
                udgifter = 0
                end if
                
                call erdetint(udgifter)
		        if isInt <> 0 then
		        call visErrorFormat
				errortype = 129
                call showError(errortype)
                Response.end
		        end if
		        isInt = 0
				
				
                if len(trim(request("FM_restestimat"))) <> 0 then
                restestimat = request("FM_restestimat")

                        call erdetint(restestimat)
		                if isInt <> 0 then
		                call visErrorFormat
				        errortype = 156
                        call showError(errortype)
                        Response.end
		                end if
		                isInt = 0

                        restestimat = abs(restestimat)

                else
                restestimat = 0
                end if


                '''Tjekekr om brutto bel�b er mindre end netto bel�b
                select case lto 
                case "epi", "epi_no", "epi_sta", "epi_ab", "epi_uk", "epi2017", "intranet - local"
                strBudgetTjk = replace(strBudget, ".", ",")
                jo_gnsbelobTjk = replace(jo_gnsbelob, ".", ",")
                if cdbl(strBudgetTjk) < cdbl(jo_gnsbelobTjk) AND func = "dbred" then
                
                        call visErrorFormat
				        errortype = 165
                        call showError(errortype)
                        Response.end
		               
                
                
                end if 
                end select


               

                stade_tim_proc = request("FM_stade_tim_proc")
       
				
                if len(trim(request("FM_syncaktdatoer"))) <> 0 then
                syncaktdatoer = 1
                else
                syncaktdatoer = 0
                end if

                
                if len(trim(request("FM_syncslutdato"))) <> 0 then
                syncslutdato = 1
                else
                syncslutdato = 0
                end if

                if len(trim(request("FM_brugaltfakadr"))) <> 0 then
                altfakadr = 1
                else
                altfakadr = 0
                end if
				
                if len(trim(request("FM_preconditions_met"))) <> 0 then
                preconditions_met = request("FM_preconditions_met")
                else
                preconditions_met = 0
                end if

                if len(trim(request("FM_fomr_konto"))) <> 0 then
                fomr_konto = request("FM_fomr_konto")
                else
                fomr_konto = 0
                end if


                jfak_moms = request("FM_jfak_moms")
                jfak_sprog = request("FM_jfak_sprog")


                if len(trim(request("FM_useFYbudgetinGT"))) <> 0 then
                useFYbudgetinGT = 1
                else
                useFYbudgetinGT = 0
                end if



                if len(trim(request("FM_opdmedarbtimepriser"))) <> 0 then
                laasmedtpbudget = 1
                else
                        
                        if func = "dbopr" then
                        select case lto 
                        case "jttek", "intranet - local"
                        laasmedtpbudget = 1
                        case else
                        laasmedtpbudget = 0
                        end select    
                        
                        else
                        laasmedtpbudget = 0

                        end if
                end if


                if len(trim(request("FM_alert"))) <> 0 then
                alert = 1 
                else
                alert = 0
                end if

                if len(trim(request("FM_jo_valuta"))) <> 0 then
                jo_valuta = request("FM_jo_valuta")
                else
                jo_valuta = 1
                end if


                '*** Finder kurs ***
                call valutaKurs(jo_valuta)
                dblKurs = dblKurs 

              
                if len(trim(request("project_tier"))) <> 0 then
                project_tier = request("project_tier")
                project_tier = replace(project_tier, "'", "")
                else
                project_tier = ""
                end if


                '*** GDPR Epinion ************************************
               

                if len(trim(request("FM_data_outside_eu"))) <> 0 then
                data_outside_eu = request("FM_data_outside_eu")
                else
                data_outside_eu = 0
                end if

                 if len(trim(request("FM_categories_process"))) <> 0 then
                categories_process = request("FM_categories_process")
                categories_process = replace(categories_process, "'", "")
                else
                categories_process = ""
                end if
                
                if len(trim(request("FM_safeguard"))) <> 0 then
                safeguard = request("FM_safeguard")
                safeguard = replace(safeguard, "'", "")
                else
                safeguard = ""
                end if


                '******* Sti til dokumenter p� egen filserver *****'
                if len(trim(request("FM_filepath1"))) <> 0 then 
                filepath1 = request("FM_filepath1")
                filepath1 = replace(filepath1, "'", "")
                filepath1 = replace(filepath1, "\", "#")
                else
                filepath1 = ""
                end if

                if len(trim(request("FM_filepath2"))) <> 0 then 
                filepath2 = request("FM_filepath2")
                filepath2 = replace(filepath2, "'", "")
                filepath2 = replace(filepath2, "\", "#")
                else
                filepath2 = ""
                end if

                if len(trim(request("FM_filepath3"))) <> 0 then 
                filepath3 = request("FM_filepath3")
                filepath3 = replace(filepath3, "'", "")
                filepath3 = replace(filepath3, "\", "#")
                else
                filepath3 = ""
                end if


                if len(trim(request("FM_gdpr_projecttype"))) <> 0 then 
                gdpr_projecttype = request("FM_gdpr_projecttype")
                else
                gdpr_projecttype = 0
                end if

                if len(trim(request("FM_gdpr_personaldata"))) <> 0 then 
                gdpr_personaldata = request("FM_gdpr_personaldata")
                else
                gdpr_personaldata = 0
                end if

                if len(trim(request("FM_gdpr_safeguard_io"))) <> 0 then 
                gdpr_safeguard_io = request("FM_gdpr_safeguard_io")
                else
                gdpr_safeguard_io = 0
                end if

                
                if len(trim(request("FM_gdpr_personaldatavendor"))) <> 0 then 
                gdpr_personaldatavendor = request("FM_gdpr_personaldatavendor")
                else
                gdpr_personaldatavendor = 0
                end if
                  


                if len(trim(request("FM_gdpr_strippedforpersonal_data"))) <> 0 then 
                gdpr_strippedforpersonal_data = request("FM_gdpr_strippedforpersonal_data")
                else
                gdpr_strippedforpersonal_data = 0
                end if

                if len(trim(request("FM_gdpr_personal"))) <> 0 then 
                gdpr_personal = request("FM_gdpr_personal")
                else
                gdpr_personal = 0
                end if

                if len(trim(request("FM_gdpr_sensitive"))) <> 0 then 
                gdpr_sensitive = request("FM_gdpr_sensitive")
                else
                gdpr_sensitive = 0
                end if


                if session("mid") = 32821 OR session("mid") = 1 then
                 
                if (lto = "epi2017" OR lto = "intranet - local") AND func = "dbred" AND strStatus = 0 AND cint(gdpr_strippedforpersonal_data) = 0 AND cint(gdpr_personaldata) = 0 then

                call visErrorFormat
				errortype = 201
                call showError(errortype)
                Response.end

                end if

                end if

                
                if lto = "epi2017" AND func = "dbred" AND strStatus = 1 AND level <> 1 AND _
                ( len(trim(filepath1)) = 0 OR _
                ( cint(gdpr_personaldata) = 0 AND ( len(trim(filepath2)) = 0 OR len(trim(categories_process)) = 0 ) ) OR _
                ( cint(data_outside_eu) = 1 AND len(trim(safeguard)) = 0 ) ) then
                
                call visErrorFormat
				errortype = 201
                call showError(errortype)
                Response.end
		        
				
                end if

				
				'**********************************'
				'**** Henter valgte kunde(r) ****'
				'**********************************'
			    strSQLkun = "SELECT kid, kkundenavn FROM kunder WHERE ketype <> 'e' AND "& kidSQLKri &" ORDER BY Kkundenavn"
			    
			    'Response.Write strSQLkun 
			    'Response.flush
			    'Response.end
			    
			    oRec.open strSQLkun, oConn, 3
			    while not oRec.EOF
				
				
				if func = "dbopr" then
				
				'**********************************'
				'*** nyt jobnr / tilbudsnr ***'
				'**********************************'
				strSQL = "SELECT jobnr, tilbudsnr FROM licens WHERE id = 1"
				oRec5.open strSQL, oConn, 3
				if not oRec5.EOF then 
				    
				    strjnr = oRec5("jobnr") + 1
				    
				    if request("FM_usetilbudsnr") = "j" then
				    tlbnr = oRec5("tilbudsnr") + 1
				    else
				    tlbnr = 0
				    end if
				    
				end if
				oRec5.close

                    
				
				else
				  
				 
                 strjnr = request("FM_jnr")
                
                 call alfanumerisk(strjnr)
                 strjnr = alfanumeriskTxt
                 strjnr = left(strjnr,20)

				 tlbnr = request("FM_tnr")  
				    
				end if
				


				
                 '************************************'
                 '***** Forrretningsomr�der **********'
                 '************************************'

                   


                    fomrArr = split(request("FM_fomr"), ",")

                    for_faktor = 0
                    'for afor = 0 to UBOUND(fomrArr)
                    'for_faktor = for_faktor + 1 
                    'next

                    'if for_faktor <> 0 then
                    'for_faktor = for_faktor
                    'else
                    'for_faktor = 1
                    'end if

                    'for_faktor = formatnumber(100/for_faktor, 2)
                    'for_faktor = replace(for_faktor, ",", ".")

                    'if len(trim(request("FM_fomr_syncakt"))) <> 0 then
                    'syncForAkt = 1
                    'else
                    syncForAkt = 0
                    'end if
              

                 '****************************************
				
				
				
				


                
                '*******************************************'		
                '** Projektgrupper ****
                '*******************************************'		
				
                
                '*** Nedarv / F�d ***'
                '** 0 = Nedarv ******'
                '** 1 = F�d *********'

                pgrp_arvefode = request("pgrp_arvefode") ''Tilf�jes l�ngere ned hvor Stam-aktiviteter tilf�jes
                
                '*** Progrp p� job ***'
                
                strProjektgr1 = request("FM_projektgruppe_1")
				strProjektgr2 = request("FM_projektgruppe_2")
				strProjektgr3 = request("FM_projektgruppe_3")
				strProjektgr4 = request("FM_projektgruppe_4")
				strProjektgr5 = request("FM_projektgruppe_5")
				strProjektgr6 = request("FM_projektgruppe_6")
				strProjektgr7 = request("FM_projektgruppe_7")
				strProjektgr8 = request("FM_projektgruppe_8")
				strProjektgr9 = request("FM_projektgruppe_9")
				strProjektgr10 = request("FM_projektgruppe_10")

                

                'Response.write "strProjektgr1:" & strProjektgr1 &" strProjektgr2: "& strProjektgr2 &" strProjektgr3: "& strProjektgr3&" strProjektgr4: "& strProjektgr4&" strProjektgr5: "& strProjektgr5&" strProjektgr6: "& strProjektgr6&" strProjektgr7: "& strProjektgr7&" strProjektgr8: "& strProjektgr8&" strProjektgr9: "& strProjektgr9&" strProjektgr10: "& strProjektgr10
                'Response.end

                if request("FM_gemsomdefault") = "1" then
				response.cookies("job")("defaultprojgrp") = strProjektgr1
				response.cookies("job").expires = date + 65
				end if

              
				
				strFakturerbart = request("FM_fakturerbart")
				
					startDato = Request("FM_start_aar") &"/" & Request("FM_start_mrd") & "/" & strStartDay 
					if request("FM_datouendelig") = "j" then
					slutDato =  "2044/1/1" 
					else
					slutDato =  Request("FM_slut_aar") &"/" & Request("FM_slut_mrd") & "/" & strSlutDay 
					end if
				
				'************************************************************************************
				
				    
				    
				    '*** Jobnr findes ***
				    jobnrFindes = 0
				    
				    if func = "dbred" then
				    strSQL = "SELECT jobnr, id FROM job WHERE id <> "& id &" AND jobnr = '" & strjnr & "'"
					else
					strSQL = "SELECT jobnr, id FROM job WHERE jobnr = '" & strjnr &"'"
					end if 
					
					'Response.Write strSQL
					'Response.Flush
					
					
					oRec5.open strSQL, oConn, 3
		            if not oRec5.EOF then	
		            
		            jobnrFindesNR = oRec5("jobnr")
		            jobnrFindesID = oRec5("id")
		            jobnrFindes = 1 
		            
		            end if
					oRec5.close
					
					'Response.Write "<br> jobnrFindes" &  jobnrFindes & " "& jobnrFindesNR &" "& jobnrFindesID &"<br>"
					'Response.flush
					
					if cint(jobnrfindes) = 1 then
					%>
					<!--#include file="../inc/regular/header_inc.asp"-->
				    <%	
					errortype = 93
					call showError(errortype)
					Response.end
				    end if 
					
				    
					'*** tilbudsnr findes ***
					
					if request("FM_usetilbudsnr") = "j" then
					tilbudsnrFindes = 0
					
				    if func = "dbred" then
				    strSQL = "SELECT tilbudsnr, id FROM job WHERE id <> "& id &" AND tilbudsnr = '" &  tlbnr & "' AND tilbudsnr <> 0"
					else
					strSQL = "SELECT tilbudsnr, id FROM job WHERE tilbudsnr = '" & tlbnr & "'"
					end if 
					
					'Response.Write strSQL
					'Response.Flush
					'Response.end
					
					oRec5.open strSQL, oConn, 3
		            if not oRec5.EOF then	
		            
		            tilbudsnrFindesNR = oRec5("tilbudsnr")
		            tilbudsnrFindes = 1 
		            
		            end if
					oRec5.close
					
					end if 
					
					
					if cint(tilbudsnrFindes) = 1 then
					%>
					<!--#include file="../inc/regular/header_inc.asp"-->
				    <% 
					errortype = 94
					call showError(errortype)
					Response.end
				    end if 
					
					 if len(trim(request("FM_lincensindehaver_faknr_prioritet_job"))) <> 0 then
                     lincensindehaver_faknr_prioritet_job = request("FM_lincensindehaver_faknr_prioritet_job")
                     else
                     lincensindehaver_faknr_prioritet_job = 0
                     end if
					
					 if cint(jobnrFindes) <> 1 AND cint(tilbudsnrFindes) <> 1 then		
					 
					 'Response.Write "her oprettes"
					 		
					 	'*** Opretter Job ***
					 	if func = "dbopr" then  
					 	
					 	
							    strSQLjob = ("INSERT INTO job (jobnavn, jobnr, jobknr, jobTpris, jobstatus, jobstartdato,"_
							    &" jobslutdato, editor, dato, creator, createdate, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, "_
							    &" projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10, "_
							    &" fakturerbart, budgettimer, fastpris, kundeok, beskrivelse "_
							    &", ikkeBudgettimer, tilbudsnr, jobans1, jobans2, "_
							    &" serviceaft, kundekpers, lukafmjob, valuta, jobfaktype, rekvnr, "_
							    &" jo_gnstpris, jo_gnsfaktor, jo_gnsbelob, jo_bruttofortj, jo_dbproc, udgifter, "_
							    &" risiko, usejoborakt_tp, ski, job_internbesk, abo, ubv, sandsynlighed, jobans3, jobans4, jobans5, "_
                                &" diff_timer, diff_sum, jo_udgifter_intern, jo_udgifter_ulev, jo_bruttooms, "_
                                &" jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, restestimat, stade_tim_proc, "_
                                &" virksomheds_proc, syncslutdato, altfakadr, preconditions_met, laasmedtpbudget,"_
                                &" salgsans1, salgsans2, salgsans3, salgsans4, salgsans5, "_
                                &" salgsans1_proc, salgsans2_proc, salgsans3_proc, salgsans4_proc, salgsans5_proc, filepath1, fomr_konto, "_
                                &" jfak_sprog, jfak_moms, alert, lincensindehaver_faknr_prioritet_job, jo_valuta, jo_valuta_kurs, jo_usefybudgetingt, filepath2, filepath3, "_
                                &" categories_process, data_outside_eu, safeguard, gdpr_safeguard_io, gdpr_personaldata, gdpr_projecttype, project_tier, gdpr_personaldatavendor, gdpr_strippedforpersonal_data, gdpr_personal, gdpr_sensitive "_
                                &") VALUES "_
							    &"('"& strNavn &"', "_
							    &"'"& strjnr &"', "_ 
							    &""& oRec("kid") &", "_  
							    &""& jo_gnsbelob &", "_ 
							    &""& strStatus &", "_ 
							    &"'"& startDato &"', "_ 
							    &"'"& slutDato &"', "_
							    &"'"& strEditor &"', "_
							    &"'"& strDato &"', "_ 
                                &"'"& strEditor &"', "_
                                &"'"& year(now) &"-"& month(now) &"-"& day(now) &"', "_ 
							    &""& strProjektgr1 &", "_ 
							    &""& strProjektgr2 &", "_ 
							    &""& strProjektgr3 &", "_ 
							    &""& strProjektgr4 &", "_ 
							    &""& strProjektgr5 &", "_
							    &""& strProjektgr6 &", "_ 
							    &""& strProjektgr7 &", "_ 
							    &""& strProjektgr8 &", "_ 
							    &""& strProjektgr9 &", "_ 
							    &""& strProjektgr10 &", "_  
							    &""& strFakturerbart &", "_ 
							    &""& strBudgettimer  &", "_ 
							    &"'"& strFastpris & "',"_
							    &""& intkundese &", "_
							    &"'"& strBesk &"', "& SQLBless(ikkeBudgettimer) &", "_
							    &"'"& tlbnr &"', "& intJobans1 &", "& intJobans2 &", "_
							    &" "& intServiceaft &", "& intKundekpers &", "& lukafmjob &", "_
							    &" "& valuta &", "& jobfaktype &", '"& rekvnr &"', "_
							    &" "& jo_gnstpris &", "& jo_gnsfaktor &", "_
							    &" "& jo_gnsbelob &","& jo_bruttofortj &","& jo_dbproc &", "& udgifter &", "_
							    &" "& intprio &", "& usejoborakt_tp &", "& ski &", '"& strInternBesk &"',"_
							    &" "& abo &", "& ubv &", "& sandsynlighed &", "& intJobans3  &", "& intJobans4  &", "& intJobans5  &", "_
                                &" "& diff_timer &", "& diff_sum &", "& jo_udgifter_intern &", "& jo_udgifter_ulev &", "& strBudget &", "_
                                &" "& jobans_proc_1 & ", "& jobans_proc_2 & ", "& jobans_proc_3 & ", "& jobans_proc_4 & ", "& jobans_proc_5 & ", "& restestimat &", "& stade_tim_proc &","_
                                &" "& virksomheds_proc &", "& syncslutdato &", "& altfakadr &", "& preconditions_met &", "& laasmedtpbudget &", "_
                                &" "& salgsans1 &","& salgsans2 &","& salgsans3 &","& salgsans4 &","& salgsans5 &", "_
                                &" "& salgsans_proc_1 &","& salgsans_proc_2 &","& salgsans_proc_3 &","& salgsans_proc_4 &","& salgsans_proc_5 &", "_
                                &" '"& filepath1 &"', "& fomr_konto &", "& jfak_sprog &", "& jfak_moms &", "& alert &", '"& lincensindehaver_faknr_prioritet_job &"', "_
                                &""& jo_valuta &", "& dblKurs &", "& useFYbudgetinGT &", '"& filepath2 &"', '"& filepath3 &"', "_
                                &"'"& categories_process &"', "& data_outside_eu &", '"& safeguard &"', "& gdpr_safeguard_io &", "& gdpr_personaldata &", "& gdpr_projecttype &", '"& project_tier &"', "_
                                &"'"& gdpr_personaldatavendor &"',"& gdpr_strippedforpersonal_data &", '"& gdpr_personal &"', '"& gdpr_sensitive &"'"_
                                &")")
    							
							    'Response.write strFakturerbart & "<br><br>"
                                'if session("mid") = 1 then
							    'Response.write strSQLjob
							    'Response.flush 
                                'end if
    							
							    oConn.execute(strSQLjob)	
								
								
								'******************************************'
								'*** finder jobid p� det netop opr. job ***'
								'******************************************'
                                't = 10
								'if t = 0 then
								
								    strSQL2 = "SELECT id FROM job WHERE id <> 0 ORDER BY id DESC " 'jobnr = " & strjnr
									oRec5.open strSQL2, oConn, 3
									if not oRec5.EOF then
									varJobIdThis = oRec5("id")
									varJobId = varJobIdThis
									end if
									oRec5.close
									
									
								'***************************************'
								'*** Opretter timepriser p� job.       *'
								'***************************************'
                                
						        'Response.write "HER"
                                'response.end
                        
                                'timeE = now
                                'loadtime = datediff("s",timeA, timeE)
							    'Response.write "<br>F�r timepriser: "& loadtime & "<br><br>"		


								strSQLpg = "SELECT id, navn FROM projektgrupper WHERE "_
								&" id = "& strProjektgr1 &" OR "_
								&" id = "& strProjektgr2 &" OR id = "& strProjektgr3 &" OR "_
								&" id = "& strProjektgr4 &" OR id = "& strProjektgr5 &" OR id = "& strProjektgr6 &" OR "_
								&" id = "& strProjektgr7 &" OR id = "& strProjektgr8 &" OR id = "& strProjektgr9 &" OR "_
								&" id = "& strProjektgr10 &" ORDER BY navn"
								oRec5.open strSQLpg, oConn, 3
								
								'Response.write strSQL
								while not oRec5.EOF
									
									strSQLmtp = "SELECT medarbejderid, projektgruppeid, mid, mnavn, timepris, timepris_a1, "_
									&" timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
									&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta "_
									&" FROM progrupperelationer "_
									&" LEFT JOIN medarbejdere ON (mid = progrupperelationer.medarbejderid) "_
									&" LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) "_
									&" WHERE projektgruppeid = "& oRec5("id") &" AND mnavn <> '' AND mansat <> 2 AND mansat <> 4 ORDER BY mnavn"
									oRec3.open strSQLmtp, oConn, 3
									this6timepris = 0
									while not oRec3.EOF
								        
								        
								        '*** Finder valuta p� job og finder timepris ***'
								        '*** p� medarbejder der matcher valuta ***'
								        '*** Hvis findes ellers v�lges 1 = DKK, Grundvaluta ****'
								            
								           valutafundet = 0
								             
								           for i = 0 to 5
								           
								                   'Response.Write "her" & i & "<br>"
								           
								                   select case i
								                   case 0
								                   tpris = oRec3("timepris")
								                   tprisValuta = oRec3("tp0_valuta")
								                   tprisAlt = 0
								                   case 1
								                   tpris = oRec3("timepris_a1")
								                   tprisValuta = oRec3("tp1_valuta")
								                   tprisAlt = 1
								                   case 2
								                   tpris = oRec3("timepris_a2")
								                   tprisValuta = oRec3("tp2_valuta")
								                   tprisAlt = 2
								                   case 3
								                   tpris = oRec3("timepris_a3")
								                   tprisValuta = oRec3("tp3_valuta")
								                   tprisAlt = 3
								                   case 4
								                   tpris = oRec3("timepris_a4")
								                   tprisValuta = oRec3("tp4_valuta")
								                   tprisAlt = 4
								                   case 5
								                   tpris = oRec3("timepris_a5")
								                   tprisValuta = oRec3("tp5_valuta")
								                   tprisAlt = 5
								                   end select
								           
                                                    'if lto = "cisu" then
                                
                                                    '    response.write "<br>Medid: "& oRec3("medarbejderid") &":"& oRec3("mid") &" projektgruppeID: "& oRec5("id") &" i: "& i &" valuta: " & valuta &" tprisValuta = "& tprisValuta &" valutafundet: "& valutafundet
                                                    '    response.flush 

                                                    'end if

								           
								                    if cint(valuta) = cint(tprisValuta) AND cint(valutafundet) = 0 then
								                    tprisUse = tpris
								                    tprisValutaUse = tprisValuta
								                    tprisAltUse = tprisAlt
								                    valutafundet = 1
								            
								                    end if
								            
							                next
								           
								           if cint(valutafundet) = 0 then
								           tprisUse = oRec3("timepris")
								           tprisValutaUse = oRec3("tp0_valuta")
								           tprisAltUse = 0
								           end if
								            
                                           tprisUse = replace(tprisUse, ".", "")
                                           tprisUse = replace(tprisUse, ",", ".")
								        
								        '*** Opretter timepriser p� job **'
								        if instr(usedmids, "#"&oRec3("mid")&"#") = 0 then
								        strSQLtp = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
								        &" VALUES ("& varJobId &", 0, "& oRec3("mid") &","& tprisAltUse &", "& tprisUse &", "& tprisValutaUse &")"
								        
                                        'if lto = "qwert" then
								        'Response.Write strSQLtp & "<br>"
								        'Response.flush
                                        'end if
								        
								        oconn.execute(strSQLtp)
								        end if
								   
								   usedmids = usedmids &oRec3("mid")&"#"
								   oRec3.movenext     
								   wend
								   oRec3.close
							    
							    
						     oRec5.movenext     
							 wend
							 oRec5.close
							   
                             'end if 't = 0
                            
                              'if lto = "qwert" then
							  'Response.end 
                              'end if 
								
                            'timeC = now
                            'loadtime = datediff("s",timeA, timeC)
							'Response.write "<br>F�r stam: "& loadtime & "<br><br>"


					            '******************************************************************'
								'Tilknytter stam aktiviteter til det job der bliver oprettet her.**'
								'******************************************************************'
								
								intAktfavgp_use = split(intAktfavgp, ",")
								strAktFase_use = split(strAktFase, ", ")
                                firstLoop = 0
								for a = 0 to UBOUND(intAktfavgp_use)
								'for a = 1 to 1
                                
                                   ' if session("mid") = 1 then
                                   '     intAktfavgp_1 = intAktfavgp_use(a)
								'	    Response.Write "intAktfavgp_1 "& intAktfavgp_1 & " a: " & a & "<br>"
                                  '  end if

								
                                    if len(trim(intAktfavgp_use(a))) <> 0 then
								    call tilknytstamakt(a, intAktfavgp_use(a), trim(strAktFase_use(1)), 0, varjobId)

                                  
								
								        'if a = 0 then
									    'intAktfavgp_1 = intAktfavgp_use(a)
									    'Response.Write "intAktfavgp_1 "& intAktfavgp_1
									    'Response.end
									    'end if
								
                                    end if

								next
								
                            'if session("mid") = 1 then
                            '    Response.write "Aktiviterer tilknyttet"
                            '    Response.end
                            'end if


                                'timeD = now
                                'loadtime = datediff("s",timeA, timeD)
							    'Response.write "<br>Efter stam: "& loadtime & "<br><br>"


                                '**** Indl�ser Underlev grp / Salgsomkostninger **'
                                call addUlev
                               

							
								
				                '*************************************************'
			                    '**** Opdaterer jobnr og tilbuds nr i Licens *****'
				                '*************************************************'
				                nytjobnr = strjnr
				                tilbudsnrKri = ""
				                
				                if request("FM_usetilbudsnr") = "j" then
				                nyttlbnr = tlbnr
				                tilbudsnrKri = ", tilbudsnr = "& nyttlbnr &""
				                end if
				                
				                strSQL = "UPDATE licens SET jobnr = "& nytjobnr &" "& tilbudsnrKri &" WHERE id = 1"
				                oConn.execute(strSQL)
				               
				                
				            '****************************************************************************************    
							'** REDIGER
                            '****************************************************************************************
							else 
							


                             '*** Opdater LUKKE dato (f�r det skifter stastus) ***'
                             call lukkedato(id, strStatus)



							strSQL = "UPDATE job SET "_
							&" jobnavn = '"& strNavn &"', "_
							&" jobnr = '"& strjnr &"', "_ 
							&" jobknr = "& oRec("kid") &", "_  
							&" jobTpris = "& jo_gnsbelob &", "_ 
							&" jobstatus = "& strStatus &", "_ 
						    &" jobstartdato = '"& startDato &"', "_ 
							&" jobslutdato = '"& slutDato &"', "_
							&" editor = '"& strEditor &"', "_
							&" dato = '"& strDato &"', "_
							&" projektgruppe1 = "& strProjektgr1 &", "_ 
							&" projektgruppe2 = "& strProjektgr2 &", "_ 
							&" projektgruppe3 = "& strProjektgr3 &", "_ 
							&" projektgruppe4 = "& strProjektgr4 &", "_
							&" projektgruppe5 = "& strProjektgr5 &", "_
							&" projektgruppe6 = "& strProjektgr6 &", "_ 
							&" projektgruppe7 = "& strProjektgr7 &", "_ 
							&" projektgruppe8 = "& strProjektgr8 &", "_ 
							&" projektgruppe9 = "& strProjektgr9 &", "_
							&" projektgruppe10 = "& strProjektgr10 &", "_
							&" fakturerbart = "& strFakturerbart &", "_
							&" Budgettimer  = "& strBudgettimer &", "_
							&" fastpris = '"& strFastpris & "', kundeok = "& intkundese &", "_
							&" beskrivelse = '"& strBesk &"', ikkeBudgettimer = "& SQLBless(ikkeBudgettimer) &", "_
							&" jobans1 =  "& intJobans1 &", "_
							&" jobans2 = "& intJobans2 &", serviceaft = "& intServiceaft &", "_
							&" kundekpers = "& intKundekpers &", lukafmjob = "& lukafmjob &", "_
							&" valuta = "& valuta &", jobfaktype = "& jobfaktype &", rekvnr = '"& rekvnr &"', "_
							&" jo_gnstpris = "& jo_gnstpris &", jo_gnsfaktor = "& jo_gnsfaktor &", "_
							&" jo_gnsbelob = "& jo_gnsbelob &", jo_bruttofortj = "& jo_bruttofortj &", "_
							&" jo_dbproc = "& jo_dbproc &", udgifter = "& udgifter &", "_
							&" risiko = "& intprio &", usejoborakt_tp = "& usejoborakt_tp &", ski = "& ski &", "_
							&" job_internbesk = '"& strInternBesk &"', abo = "& abo &", ubv = "& ubv &", "_
							&" sandsynlighed = "& sandsynlighed &", tilbudsnr = '"& tlbnr &"', "_
                            &" jobans3 = "& intJobans3 &", jobans4 = "& intJobans4 &", jobans5 = "& intJobans5 &", "_
                            &" diff_timer = "& diff_timer &", diff_sum = "& diff_sum &", "_
                            &" jo_udgifter_intern = "& jo_udgifter_intern &", jo_udgifter_ulev = "& jo_udgifter_ulev &", jo_bruttooms = "& strBudget &", "_
                            &" jobans_proc_1 = "& jobans_proc_1 & ", jobans_proc_2 = "& jobans_proc_2 & ", jobans_proc_3 = "& jobans_proc_3 & ", jobans_proc_4 = "& jobans_proc_4 & ", jobans_proc_5 = "& jobans_proc_5 & ", "_
                            &" restestimat = "& restestimat &", stade_tim_proc = "& stade_tim_proc &", virksomheds_proc = "& virksomheds_proc &", "_
                            &" syncslutdato = "& syncslutdato &", altfakadr = "& altfakadr &", preconditions_met = "& preconditions_met &", laasmedtpbudget = "& laasmedtpbudget &", "_
                            &" salgsans1 = "& salgsans1 &", salgsans2 = "& salgsans2 &", salgsans3 = "& salgsans3 &", salgsans4 = "& salgsans4 &", salgsans5 = "& salgsans5 &", "_
                            &" salgsans1_proc = "& salgsans_proc_1 &", salgsans2_proc = "& salgsans_proc_2 &", salgsans3_proc = "& salgsans_proc_3 &", salgsans4_proc = "& salgsans_proc_4 &", "_
                            &" salgsans5_proc = "& salgsans_proc_5 &", filepath1 = '"& filepath1 &"', fomr_konto = "& fomr_konto &","_
                            &" jfak_sprog = "& jfak_sprog &", jfak_moms = "& jfak_moms &", alert = "& alert &", "_
                            &" lincensindehaver_faknr_prioritet_job = '"& lincensindehaver_faknr_prioritet_job &"', "_
                            &" jo_valuta = "& jo_valuta &", jo_valuta_kurs = "& dblKurs &", jo_usefybudgetingt = "& useFYbudgetinGT &", "_
                            &" filepath2 = '"& filepath2 &"', filepath3 = '"& filepath3 &"', categories_process = '"& categories_process &"', data_outside_eu = "& data_outside_eu &", safeguard = '"& safeguard &"',"_
                            &" gdpr_safeguard_io = "& gdpr_safeguard_io &", gdpr_personaldata = "& gdpr_personaldata &", gdpr_projecttype = "& gdpr_projecttype &", project_tier = '"& project_tier &"',"_
                            &" gdpr_personaldatavendor = '"& gdpr_personaldatavendor &"', gdpr_strippedforpersonal_data = "& gdpr_strippedforpersonal_data&", gdpr_personal = '"& gdpr_personal &"',gdpr_sensitive = '"& gdpr_sensitive &"'"_
							&" WHERE id = "& id 
							
							'Response.Write strSQL
							'Response.end
                            'Response.flush							
							oConn.execute(strSQL)
							                    



                                                '********************************************
                                                '****** opdaterer tilbudsnr ved rediger ****'
                                                '********************************************
                                                if request("FM_usetilbudsnr") = "j" then
				                                nyttlbnr = tlbnr
                                                     
                                                     senesteTilbnr = 0
                                                     strSQL = "SELECT tilbudsnr FROM licens WHERE id = 1"
                                                     oRec5.open strSQL, oConn, 3
                                                     if not oRec5.EOF then
                                                     senesteTilbnr = oRec5("tilbudsnr")
                                                     end if
                                                     oRec5.close

				                                end if
				                                
                                                if cdbl(nyttlbnr) <> 0 AND cdbl(nyttlbnr) >= cdbl(senesteTilbnr) then
				                                strSQL = "UPDATE licens SET tilbudsnr = "& nyttlbnr &" WHERE id = 1"
				                                oConn.execute(strSQL)
				                                end if

                                
                                                '***********************************'
                                                '** Opdaterer kundeoplysninger *****'
                                                '***********************************'
                                                
                                                
                            
								
								                '*** Overf�rer gamle timeregistreringer til ny aftale (hvis der skiftes aftale) **'
								                if request("FM_overforGamleTimereg") = "1" then
                								
								                '*** Opdaterer timereg tabellen ****
								                strSQLtimer = "UPDATE timer SET "_
								                & " Tknavn = '"& replace(oRec("kkundenavn"), "'", "''") &"', Tknr = "& oRec("kid")&", "_
								                & " Tjobnavn = '"& strNavn &"', "_
								                & " Tjobnr = '"& strjnr &"', "_
								                & " fastpris = '"& strFastpris &"', "_
								                & " seraft = "& intServiceaft &" "_
								                & " WHERE Tjobnr = '"& strOLDjobnr & "'"
                								
                								'** Husk materiale forbrug ***
								                strSQLmat_forbrug = "UPDATE materiale_forbrug SET serviceaft = " & intServiceaft &""_
								                & " WHERE jobid = " & id
                								
								                oConn.execute(strSQLmat_forbrug)
                								
                								
								                else
                								
								                '*** Opdaterer timereg tabellen ****
								                strSQLtimer = "UPDATE timer SET "_
								                & " Tknavn = '"& replace(oRec("kkundenavn"), "'", "''") &"', Tknr = "& oRec("kid") &", "_
								                & " Tjobnavn = '"& strNavn &"', "_
								                & " Tjobnr = '"& strjnr &"', "_
								                & " fastpris = '"& strFastpris &"' "_
								                & " WHERE Tjobnr = '"& strOLDjobnr & "'"
                								
								                end  if
                								
                								
								                'Response.write strSQLtimer
								                'Response.flush
                								
								                oConn.execute(strSQLtimer)
								                
								                
								                '*** Opdaterer faktura tabel s� faktura kunde id passer hvis der er skiftet kunde  ved rediger job.
								                '*** Adr. i adresse felt p� faktura behodes til revisor spor. **'
								                strSQLFakadr = "UPDATE fakturaer SET fakadr = "& oRec("kid") &" WHERE jobid = " & id
								                oConn.execute(strSQLFakadr)
								
								
						                        
                        
                                varJobId = id
								

                               


                                '***** Sync Job slutDato ***********'
                                if strStatus = 0 then 'Kun ved luk job
                                syncslutdato = syncslutdato
                                else
                                syncslutdato = 0
                                end if

                                call syncJobSlutDato(varJobId, strjnr, syncslutdato)

                               
                                '****** Sync. datoer p� aktiviteter *******'
                                if syncaktdatoer = 1 then
                                    
                                    '** Hvis sync jobslutdato er valgt ***
                                    if syncslutdato <> 1 then
                                    useSlutDato = slutDato
                                    else
                                    useSlutDato = useSyncDato 
                                    end if

                                strSQLaDatoer = "UPDATE aktiviteter SET aktstartdato = '"& startDato &"', aktslutdato = '"& useSlutDato &"' WHERE job = "& varJobId
                                oConn.execute(strSQLaDatoer)

                                end if

                               






								
				                '************************************************'
				                '**** Diffentierede timepriser ******************'
			                    '*** Opdater eksisterende time-registreringer ***'
			                    '************************************************'
                				
                                '** sync alle aktiviteter ***'
                		        if request("FM_sync_tp") = "1" then
                                syncAkt = 1
                                else
                                syncAkt = 0
                                end if

                                 


				                medarb_tpris = request("FM_use_medarb_tpris")
				                Dim intMedArbID 
				                Dim b
                				strMedabTimePriserSlet = " AND medarbid <> 0"

					                intMedArbID = Split(medarb_tpris, ", ")
					                For b = 0 to Ubound(intMedArbID)
                                        
                                         
                                         strMedabTimePriserSlet = strMedabTimePriserSlet & " AND medarbid <> "& intMedArbID(b)


                                        if cint(syncAkt) = 1 then
                                            
                                            '** ikke KM **
                                            fakbaraktid = ""
                                            strSQLatyp = "SELECT id, fakturerbar FROM aktiviteter WHERE job = " & id & " AND fakturerbar = 5"
                                            oRec5.open strSQLatyp, oConn, 3
                                            fb = 0
                                            while not oRec5.EOF 
                                                
                                                if fb <> 0 then
                                                fakbaraktid = fakbaraktid & " OR "
                                                end if

                                                if oRec5("fakturerbar") = 5 then
                                                fakbaraktid = fakbaraktid & "aktid <> "& oRec5("id") 
                                                else
                                                fakbaraktid = fakbaraktid 
                                                end if

                                                fb = fb + 1
                                            oRec5.movenext
                                            wend
                                            oRec5.close

                                            if len(trim(fakbaraktid)) <> 0 then
                                            fakbaraktid = " AND ("& fakbaraktid & ")"
                                            else
                                            fakbaraktid = ""
                                            end if



						                strSQLdeltp = "DELETE FROM timepriser WHERE jobid = "& id &" AND medarbid = "& intMedArbID(b) &" "& fakbaraktid 
                                        oConn.execute(strSQLdeltp)

                                        
                                        else '** ellers kun job // Eller valgte akt **'

                                        strSQLdeltp = "DELETE FROM timepriser WHERE jobid = "& id &" AND medarbid = "& intMedArbID(b) &" AND aktid = "& hd_tp_jobaktid 
                                        oConn.execute(strSQLdeltp)
                        
                                        end if

                                        
                                        strSQLdeltp = strSQLdeltp & "<br>" & strSQLdeltp


                                         
						               
                                        'if len(request("FM_6timepris_"&intMedArbID(b)&"")) <> 0 OR jobids(j) = 0 then
                						
						                if len(request("FM_6timepris_"&intMedArbID(b)&"")) <> 0 then
						                call erDetInt(request("FM_6timepris_"&intMedArbID(b)&""))
						                    if isInt > 0 then
						                    this6timepris = 0
						                    else
						                    this6timepris = SQLBless(replace(request("FM_6timepris_"&intMedArbID(b)&""),".",""))
						                    end if
						                
						                valuta6 = request("FM_valuta_600"& intMedArbID(b))
						                
						                else
						                this6timepris = 0
						                valuta6 = 1
						                end if
						               
                						
                                       ' if hd_tp_jobaktid = 0 then '** Job / Akt, Hviklen radio bt er valgt ***'
                                       '*** Ved sync nedarver alle og der skal derfor ikke inds�ttes en record heller ikke for den aktivitet man st�r p� ***'
                                       '**** Kun hvis man st�r p� selve jobbet ***'

                                        if len(request("FM_6timepris_"&intMedArbID(b)&"")) <> 0 AND (cint(syncAkt) <> 1 OR (cint(syncAkt) = 1 AND hd_tp_jobaktid = 0)) then

                                            radiobtVal = SQLBless(replace(request("FM_hd_timepris_"&intMedArbID(b)&"_"&request("FM_timepris_"&intMedArbID(b)&"")),".",""))

                						    if this6timepris = radiobtVal then
                						    tprisalt = request("FM_timepris_"&intMedArbID(b)&"")
                						    else
                						    tprisalt = 6
                                            end if
                						
                                            
                                            
						                    strSQLupdTimePriser = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
						                    &" VALUES ("& id &","& hd_tp_jobaktid &", "& intMedArbID(b) &", "& tprisalt &", "& this6timepris &", "& valuta6 &")"
                						
                                       
                                            'Response.Write strSQLupdTimePriser & "<br>"
                                            'Response.flush

                                             oConn.execute(strSQLupdTimePriser)

                                       
                                       

                                         end if
                                        






                                        
                                        '*********************************************************'
						                '** Opdaterer timereg (kun p� aktiviteter der nedarver) **'
                                        '** Dvs dem der ikke findes i timepriser tabellen       **'

                                        '*** Else den specifikke aktivitet ***
						               
						                '*** Kr�ver at der minimun findes 1 akt. *****************
						              	'*** Ved sync er denne tom (lige slettet ovenfor) og alle **'
                                        '**' timeprsier p� alle aktiviter forall medeb. opdateres ***'
                                        '*********************************************************'


                                        if cint(syncAkt) = 1 OR hd_tp_jobaktid = 0 then
										strSQLtimp = "SELECT aktid FROM timepriser WHERE jobid = "& id &" AND medarbid = "& intMedArbID(b) &" AND aktid <> 0"
										
										'Response.Write strSQLtimp & "<br>"
										'Response.flush
										
										notAkt = ""
										oRec5.open strSQLtimp, oConn, 3
										while not oRec5.EOF
										notAkt = notAkt & " AND taktivitetId <> "& oRec5("aktid") 
										oRec5.movenext
										wend
										oRec5.close 

                                        specifikakt = 0

                                        else

                                        specifikakt = 1
                                        notAkt = " AND taktivitetId = "& hd_tp_jobaktid 
                                        
                                        aktnavn = ""
                                        strSQLalleaktmnavn = "SELECT navn FROM aktiviteter WHERE id = "& hd_tp_jobaktid 
                                        oRec6.open strSQLalleaktmnavn, oConn, 3
                                        if not oRec6.EOF then
                                        aktnavn = oRec6("navn")
                                        end if
                                        oRec6.close

                                        end if
										
										

                                        '*** Opdaterer alle timeregistreringer p� valgte akt. eller p� alle aktiviteter ved sync ***'

										   '**** Finder aktuel kurs ***'
						                   strSQL = "SELECT kurs FROM valutaer WHERE id = " & valuta6
						                   oRec5.open strSQL, oConn, 3
        						            
						                   if not oRec5.EOF then
						                   nyKurs = replace(oRec5("kurs"), ",", ".")
						                   end if 
						                   oRec5.close

                                         '** Fra dato ved opdater timepriser ***'
                                        fraDato = request("FM_opdatertpfra") 
                                        if isDate(fraDato) = false then
                                        fraDato = year(now) &"/"& month(now) & "/"& day(now)
                                        else
                                        fraDato = year(fraDato) &"/"& month(fraDato) & "/"& day(fraDato) 
                                        end if

										
					                    
					                    
                                        if request("FM_opdateralleakt") = "1" AND specifikakt = 1 then 
                                        '** Specifik akt + alle med samme navn **' 
			                            strSQL = "UPDATE timer SET timepris = "& this6timepris  &", valuta = "& valuta6 &", kurs = "& nyKurs &""_
			                            &" WHERE (tmnr = "& intMedArbID(b) &" AND tdato >= '"& fraDato &"' AND (tfaktim <> 5)) AND taktivitetnavn = '" & aktnavn &"'"
                                        else
                                        '** Nedarv (IKKE k�rsel) / specifik akt 
			                            strSQL = "UPDATE timer SET timepris = "& this6timepris  &", valuta = "& valuta6 &", kurs = "& nyKurs &""_
			                            &" WHERE (tmnr = "& intMedArbID(b) &" AND tjobnr = '"& strjnr &"' AND tdato >= '"& fraDato &"' AND (tfaktim <> 5)) " & notAkt
					                    end if

                                        'Response.Write strSQL & "<br>"
					                    'Response.end
					                    
					                    oConn.execute(strSQL)
                                                

                             next '** intMedArbID(b)




                            


                            '*************************************************************************************
                            '*** Renser ud i timepriser p� medarbejdere der ikke l�ngere er tilknyttet jobbbet ***
                            '*************************************************************************************
                                       
                            '*** Pilot WWF 26-09-2011 *** 
                            'if lto = "wwf" then

                            if len(trim(request("FM_sync_tp_rens"))) <> 0 then
                            sync_tp_rens = 1
                            else
                            sync_tp_rens = 0
                            end if

                            '** KUN VED opdater timepriser bliver strMedabTimePriserSlet SAT ELLERS NULSTILLES timepriser p� alle medarb for alle aktiviteter ***'
                            '** FARLIG ASSURATOR; OKO oplever alle deres timepriser p� medarbejdere forsvinder ***'
                		    if cint(sync_tp_rens) = 1 AND strMedabTimePriserSlet <> " AND medarbid <> 0" then
                            strSQLRensTp = "DELETE FROM timepriser WHERE jobid = "& id &" "& strMedabTimePriserSlet 
                            oConn.execute(strSQLRensTp)
                            'end if

                            'if session("mid") = 1 then

                            'Response.Write strSQLRensTp
                            'Response.flush
                            'response.write "<hr>"

                            'response.write strSQLdeltp
                            
                            'Response.end
                            'end if

                            end if


                                '**** Timepriser END ***'

				                
				                '**********************************************************'
								'*** Opdaterer allerede tilknyttet aktiviteter   **********'
								
								
								    if len(request("FM_aktnavn")) > 3 then
	                                len_aktnavn = len(request("FM_aktnavn"))
	                                left_aktnavn = left(request("FM_aktnavn"), len_aktnavn - 3)
	                                strAktnavn = left_aktnavn
	                                else
	                                strAktnavn = ""
	                                end if
                                	
	                                aktnavn = strAktnavn
	                                akttimer = request("FM_akttimer")

                                    if len(trim(request("FM_aktkonto"))) <> 0 then
                                    aktkonto = request("FM_aktkonto")
                                    else
                                    aktkonto = ""
                                    end if

                                    

                                    if len(trim(request("FM_avarenr"))) <> 0 then
                                    avarenr = trim(request("FM_avarenr"))
                                    else
                                    avarenr = ""
                                    end if

                                    'if session("mid") = 1 then
                                    'Response.write "aktavarnr: "& request("FM_avarenr") & "<br>aktkonto: " & request("FM_aktkonto")
                                    'Response.end
                                    'end if


	                                aktantalstk = request("FM_aktantalstk")
	                                aktpris = request("FM_aktpris")
	                                
	                                'Response.Write request("FM_aktpris")
	                                'Response.end
	                                aktids = request("FM_aktid")

                                    'Response.Write "aktids" & aktids
	                                'Response.end

	                                aktstatus = request("FM_aktstatus")
                                	
                                    'Response.Write request("FM_aktfase")
                                    'Response.end

	                                if len(request("FM_aktfase")) > 3 then
	                                len_Fasenavn = len(request("FM_aktfase"))
	                                left_Fasenavn = left(request("FM_aktfase"), len_Fasenavn - 3)
	                                strFasenavn = left_Fasenavn
	                                else
	                                strFasenavn = ", #"
	                                end if
                                	
	                                aktfaser = strFasenavn
	                                aktbgr = request("FM_aktbgr")
                                	
	                                'Response.Write request("FM_akt_totpris")
	                                'Response.end
                                	
	                                akttotpris = request("FM_akttotpris")
                                	
	                                'Response.Write request("FM_fase")
	                                'Response.end
                                	
	                                aktslet = request("FM_slet_aid")
                                	
	                                'Response.Write aktSlet
	                                'Response.end
								    
								    aktslet_aids = 0
								
								    call opdateraktliste(varJobId, aktids, aktnavn, akttimer, aktantalstk, aktfaser, aktbgr, aktpris, aktstatus, akttotpris, aktslet, aktslet_aids, aktkonto, avarenr)

								'Response.write "HEJ"
								'Response.end
								
								'**********************************************************'
								'**** Tilknytter flere Stam-aktivitets grupper til det  ***'
								'**** job der bliver redigeret                          ***'
								'**********************************************************'
								
                                'Response.write "intAktfavgp " & intAktfavgp & "<br>"
                                'Response.write "strAktFase " & strAktFase & "<br>"
                                'Response.end

								intAktfavgp_use = split(intAktfavgp, ",")
								strAktFase_use = split(strAktFase, ", ")
                                firstLoop = 0

                               

								        for a = 0 to UBOUND(intAktfavgp_use)
								        'for a = 1 to 1
                                            'Response.Write "a" & a & " val: "& intAktfavgp_use(a) &"<br>"
                                            'Response.flush
                                            if len(trim(intAktfavgp_use(a))) <> 0 then
								            call tilknytstamakt(a, intAktfavgp_use(a), trim(strAktFase_use(1)), 0, varjobId)
                                            end if
							
                                        next

                                
                                'Response.write "her"
                                'Response.end

                             
								
								                

                                                 '*** Opdaterer projektgrupper p� Eksisterende aktiviteter ved redigering af job ***'
                                                 '*** S� de f�lger jobbet. = Sync (nedarv p� akt.) ***'
                    								
					                                if request("FM_opdaterprojektgrupper") = 1 then
                                                    
                                                    'Response.write "her"
                                                    'Response.end 
                                                    	
					                                strSQL = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM job WHERE id = "& varJobId 
                    	
					                                oRec5.open strSQL, oConn, 3
					                                if not oRec5.EOF then
                    									
					                                    oConn.execute("UPDATE aktiviteter SET"_
					                                    &" projektgruppe1 = "& oRec5("projektgruppe1") &" , projektgruppe2 = "& oRec5("projektgruppe2") &", "_
					                                    &" projektgruppe3 = "& oRec5("projektgruppe3") &", "_
					                                    &" projektgruppe4 = "& oRec5("projektgruppe4") &", "_
					                                    &" projektgruppe5 = "& oRec5("projektgruppe5") &", "_
					                                    &" projektgruppe6 = "& oRec5("projektgruppe6") &", "_
					                                    &" projektgruppe7 = "& oRec5("projektgruppe7") &", "_
					                                    &" projektgruppe8 = "& oRec5("projektgruppe8") &", "_
					                                    &" projektgruppe9 = "& oRec5("projektgruppe9") &", "_
					                                    &" projektgruppe10 = "& oRec5("projektgruppe10") &" WHERE job = "& varJobId &"")
                		
					                                end if
					                                oRec5.close
                    									
					                                end if
								                   
								
								
							                    '**** Indl�ser Underlev grp / Salgsomkostninger **'
                                                call addUlev
							
							
							
							                   '***** Medarbejertyper timebudget ****'
                                               
                                               strSQLmtyp_tpDEL = "DELETE FROM medarbejdertyper_timebudget WHERE jobid = "& varJobId
                                               oConn.execute(strSQLmtyp_tpDEL)

                                               'Response.write request("FM_mtype_id") & "<hr>"
                                               
                                               mtype_id_arr = split(request("FM_mtype_id"), ",")
                                               


                                               for t = 0 to UBOUND(mtype_id_arr) 
                                                        
                                                        
                                                        tpy_id = trim(mtype_id_arr(t))

                                                        if len(trim(request("FM_mtype_timer_"& tpy_id &""))) <> 0 then
                                                        tpy_timer = request("FM_mtype_timer_"& tpy_id &"")
                                                        else
                                                        tpy_timer = 0
                                                        end if
                                                        
                                                        tpy_timer = replace(tpy_timer, ".", "")
                                                        tpy_timer = replace(tpy_timer, ",", ".") 

                                                        call erDetInt(tpy_timer)
                                                        if isint = 0 then
                                                        tpy_timer = tpy_timer
                                                        else
                                                        tpy_timer = 0
                                                        end if

                                                        '*** Nulstiller ikke ISINT da alle felter skal v�re iorden f�r record indl�ses **'

                                                        if len(trim(request("FM_mtype_timepris_"& tpy_id &""))) <> 0 then
                                                        tpy_timepris = request("FM_mtype_timepris_"& tpy_id &"")
                                                        else
                                                        tpy_timepris = 0
                                                        end if
                                                        
                                                        tpy_timepris = replace(tpy_timepris, ".", "")
                                                        tpy_timepris = replace(tpy_timepris, ",", ".") 


                                                        call erDetInt(tpy_timepris)
                                                        if isint = 0 then
                                                        tpy_timepris = tpy_timepris
                                                        else
                                                        tpy_timepris = 0
                                                        end if


                                                        if len(trim(request("FM_mtype_kostpris_"& tpy_id &""))) <> 0 then
                                                        tpy_kostpris = request("FM_mtype_kostpris_"& tpy_id &"")
                                                        else
                                                        tpy_kostpris = 0
                                                        end if
                                                        
                                                        tpy_kostpris = replace(tpy_kostpris, ".", "")
                                                        tpy_kostpris = replace(tpy_kostpris, ",", ".") 


                                                        call erDetInt(tpy_timepris)
                                                        if isint = 0 then
                                                        tpy_timepris = tpy_timepris
                                                        else
                                                        tpy_timepris = 0
                                                        end if


                                                        if len(trim(request("FM_mtype_faktor_"& tpy_id &""))) <> 0 then
                                                        tpy_faktor = request("FM_mtype_faktor_"& tpy_id &"")
                                                        else
                                                        tpy_faktor = 0
                                                        end if
                                                        
                                                        tpy_faktor = replace(tpy_faktor, ".", "")
                                                        tpy_faktor = replace(tpy_faktor, ",", ".") 


                                                         call erDetInt(tpy_faktor)
                                                        if isint = 0 then
                                                        tpy_faktor = tpy_faktor
                                                        else
                                                        tpy_faktor = 0
                                                        end if


                                                        if len(trim(request("FM_mtype_belob_"& tpy_id &""))) <> 0 then
                                                        tpy_belob = request("FM_mtype_belob_"& tpy_id &"")
                                                        else
                                                        tpy_belob = 0
                                                        end if
                                                        
                                                        tpy_belob = replace(tpy_belob, ".", "")
                                                        tpy_belob = replace(tpy_belob, ",", ".") 

                                                        call erDetInt(tpy_belob)
                                                        if isint = 0 then
                                                        tpy_belob = tpy_belob
                                                        else
                                                        tpy_belob = 0
                                                        end if

                                                        if len(trim(request("FM_mtype_belob_ff_"& tpy_id &""))) <> 0 then
                                                        tpy_belob_ff = request("FM_mtype_belob_ff_"& tpy_id &"")
                                                        else
                                                        tpy_belob_ff = 0
                                                        end if
                                                        
                                                        tpy_belob_ff = replace(tpy_belob_ff, ".", "")
                                                        tpy_belob_ff = replace(tpy_belob_ff, ",", ".") 

                                                        call erDetInt(tpy_belob_ff)
                                                        if isint = 0 then
                                                        tpy_belob_ff = tpy_belob_ff
                                                        else
                                                        tpy_belob_ff = 0
                                                        end if

                                                        if isInt = 0 then

                                                        strSQLmtyp_tp = "INSERT INTO medarbejdertyper_timebudget "_
                                                        &" (jobid, typeid, timer, timepris, faktor, belob, belob_ff, kostpris) "_
                                                        &" VALUES "_
                                                        &"("& varJobId &", "& tpy_id &", "& tpy_timer &", "& tpy_timepris &", "& tpy_faktor &", "& tpy_belob &", "& tpy_belob_ff &", "& tpy_kostpris &")"


                                                        'Response.write strSQLmtyp_tp & "<br>"
                                                        'Response.flush

                                                        oConn.execute(strSQLmtyp_tp)

                                                        end if

                                               next


                                              
							
							
							end if 
                            'Response.end
							'**** Opret / Rediger Job ***'
				
                '******************************************************************************
                '****** WIP historik **********************************************************

                 ddDato = year(now) &"/"& month(now) &"/"& day(now) 
             
                 strSQLUpdjWiphist = "INSERT INTO wip_historik (dato, restestimat, stade_tim_proc, medid, jobid) VALUES ('"& ddDato &"', "& restestimat &", "& stade_tim_proc &", "& session("mid") &", "& id &")"
                 oConn.Execute(strSQLUpdjWiphist)

                        		



                '*************************************************************************'
                '***** Adviser jobansvarlige ****************************
                '*************************************************************************'
	            if len(trim(request.form("FM_adviser_jobans"))) <> 0 then
            	
		                    
                            jobans1 = 0
                            jobans2 = 0
                            jobans3 = 0    
                            jobans4 = 0
                            jobans5 = 0

                            jobans1email = ""
                            jobans2email = ""
                            jobans3email = ""
                            jobans4email = ""
                            jobans5email = ""

                            jobstatus = 0
                            lkDatoThis = "2002-01-01"


				            '**** Finder jobansvarlige *****
				            strSQL = "SELECT job.id AS jid, jobnavn, jobnr, jobstatus, lukkedato, jobans1, jobans2, jobans3, jobans4, jobans5, job.beskrivelse, job_internbesk, "_
                            &" m1.mnavn AS m1mnavn, m1.email AS m1email, m1.mansat AS m1mansat, "_
				            &" m2.mnavn AS m2mnavn, m2.email AS m2email, m2.mansat AS m2mansat, "_
                            &" m3.mnavn AS m3mnavn, m3.email AS m3email, m3.mansat AS m3mansat, "_
                            &" m2.mnavn AS m4mnavn, m4.email AS m4email, m4.mansat AS m4mansat, "_    
                            &" m2.mnavn AS m5mnavn, m5.email AS m5email, m5.mansat AS m5mansat, "_    
                            &" kkundenavn, kkundenr, m2.init AS m2init, m1.init AS m1init, m3.init AS m3init, m4.init AS m4init, m5.init AS m5init FROM job "_
				            &" LEFT JOIN medarbejdere AS m1 ON (m1.mid = jobans1)"_
				            &" LEFT JOIN medarbejdere AS m2 ON (m2.mid = jobans2)"_
                            &" LEFT JOIN medarbejdere AS m3 ON (m3.mid = jobans3)"_
                            &" LEFT JOIN medarbejdere AS m4 ON (m4.mid = jobans4)"_
                            &" LEFT JOIN medarbejdere AS m5 ON (m5.mid = jobans5)"_
                            &" LEFT JOIN kunder AS k ON (k.kid = job.jobknr)"_
				            &" WHERE job.id = "& varJobId

                            'Response.Write strSQL
                            'Response.end

				            oRec5.open strSQL, oConn, 3
				            x = 0
				            if not oRec5.EOF then
            				
				            jobid = oRec5("jid")
				                
                            jobstatus = oRec5("jobstatus")
                            lkDatoThis = oRec5("lukkedato")

                            jobans1 = oRec5("m1mnavn")
                            jobans1Init = oRec5("m1init")
                            if isNull(oRec5("m1email")) <> true then 
                            jobans1email = oRec5("m1email")
                            else
                            jobans1email = ""
                            end if
                            jobans1Mansat = oRec5("m1mansat")
                        
                            jobans2 = oRec5("m2mnavn")
				            jobans2Init = oRec5("m2init")
                            if isNull(oRec5("m2email")) <> true then 
                            jobans2email = oRec5("m2email")
                            else
                            jobans2email = ""
                            end if
                            jobans2Mansat = oRec5("m2mansat")
				            
                            jobans3 = oRec5("m3mnavn")
				            jobans3Init = oRec5("m3init")
                            if isNull(oRec5("m3email")) <> true then 
                            jobans3email = oRec5("m3email")
                            else
                            jobans3email = ""
                            end if
                            jobans3Mansat = oRec5("m3mansat")

                            jobans4 = oRec5("m4mnavn")
				            jobans4Init = oRec5("m4init")
                            if isNull(oRec5("m4email")) <> true then 
                            jobans4email = oRec5("m4email")
                            else
                            jobans4email = ""
                            end if
                            jobans4Mansat = oRec5("m4mansat")

                            jobans5 = oRec5("m5mnavn")
				            jobans5Init = oRec5("m5init")
                            if isNull(oRec5("m5email")) <> true then 
                            jobans5email = oRec5("m5email")
                            else
                            jobans5email = ""
                            end if
                            jobans5Mansat = oRec5("m5mansat")

                            jobnavnThis = oRec5("jobnavn")
                            intJobnr = oRec5("jobnr")
                            strkkundenavn = oRec5("kkundenavn")

                            strBesk = oRec5("beskrivelse")
                            job_internbesk = oRec5("job_internbesk")

                            'kkundenr = oRec("kkundenr")
            				
				            
				            end if
				            oRec5.close
            				
            				
				            '*** Henter afsender **
				            strSQL = "SELECT mnavn, email FROM medarbejdere"_
				            &" WHERE mid = "& session("mid")
				            oRec5.open strSQL, oConn, 3
            				
				            if not oRec5.EOF then
            				
				            afsNavn = oRec5("mnavn")
				            afsEmail = oRec5("email")
            				
				            end if
				            oRec5.close
            				
            				
            					
            						
            					
				            if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\jobs.asp" then
					  	            'Response.Write "isNULL(jobans1) & isNULL(jobans2)" &  isNULL(jobans1) &" "& isNULL(jobans2)
                                    'Response.end


                                    for m = 1 to 5

                                    jobAnsThis = 0

                                    select case m
                                    case 1
                                    jobAnsThis = jobans1
                                    jobAnsThisEmail = jobans1email
                                    jobAnsAnsat = jobans1Mansat
                                    case 2
                                    jobAnsThis = jobans2
                                    jobAnsThisEmail = jobans2email
                                    jobAnsAnsat = jobans2Mansat
                                    case 3
                                    jobAnsThis = jobans3
                                    jobAnsThisEmail = jobans3email
                                    jobAnsAnsat = jobans3Mansat
                                    case 4
                                    jobAnsThis = jobans4
                                    jobAnsThisEmail = jobans4email
                                    jobAnsAnsat = jobans4Mansat
                                    case 5  
                                    jobAnsThis = jobans5
                                    jobAnsThisEmail = jobans5email
                                    jobAnsAnsat = jobans5Mansat
                                    end select


                                    if (jobAnsThis <> "0" AND isNULL(jobAnsThis) <> true) then
                                    
                                    Set myMail=CreateObject("CDO.Message")
                                    myMail.From="timeout_no_reply@outzource.dk" 'TimeOut Email Service 

                                        
						            myMail.To= ""& jobAnsThis &"<"& jobAnsThisEmail &">"
                                   
                                    
						            'Mailer.Subject = "Til de jobansvarlige p�: "& jobnavnThis &" ("& intJobnr &") | " & strkkundenavn  
                                    myMail.Subject = job_txt_020 &": "& jobnavnThis &" ("& intJobnr &") | " & strkkundenavn  


		                            strBody = "<br>"
                                    strBody = strBody &"<b>"&job_txt_021&":</b> "& strkkundenavn & "<br>" 
						            strBody = strBody &"<b>"&job_txt_022&":</b> "& jobnavnThis &" ("& intJobnr &")"
                                    if lto = "mpt" then
                                        strBody = strBody & "<br>"
                                        
                                        statusnavn = ""
                                        strlkDatoThis = ""
                                        select case cint(jobstatus)
									        case 1
									        strStatusNavn = job_txt_094
									        case 2
									        strStatusNavn = job_txt_095 'passiv
									        case 0
									        strStatusNavn = job_txt_096
                                                if cdate(lkDatoThis) <> "01-01-2002" then
                                                strlkDatoThis = " ("& formatdatetime(lkDatoThis, 2) & ")"
                                                end if
									        case 3
									        strStatusNavn = job_txt_063
                                            case 4
									        strStatusNavn = job_txt_098
                                            case 5
									        strStatusNavn = "Evaluering"
									    end select

                                        strBody = strBody &"<b>"&job_txt_241&":</b> "& strStatusNavn & strlkDatoThis & "<br><br>"

                                    else
                                        strBody = strBody & "<br><br>"
                                    end if

                                    if jobans1 <> "0" AND isNULL(jobans1) <> true then
                                    strBody = strBody &"<b>"&job_txt_023&":</b> "& jobans1 &" "& jobans1init &"<br><br>"
                                    end if

                                    if jobans2 <> "0" AND isNULL(jobans2) <> true then
                                    strBody = strBody &"<b>"&job_txt_024&":</b> "& jobans2 &" ("& jobans2init &") <br><br>"
		                            end if

                                    if len(trim(strBesk)) <> 0 then
                                    strBody = strBody &"<br><b>"&job_txt_025&":</b><br>"
                                    strBody = strBody & strBesk &"<br><br><br><br>"
                                    end if

                                    if len(trim(job_internbesk)) <> 0 then
                                    strBody = strBody &"<br><b>"&job_txt_026&":</b><br>"
                                    strBody = strBody & job_internbesk &"<br><br>"
                                    end if

                                    
                                    'strBody = strBody &"<br><br><br>https://outzource.dk/timeout_xp/wwwroot/ver2_10/login.asp?lto="&lto&"&tomobjid="&jobid&"&key="&strLicenskey

                                   'strBody = strBody &"<br><br><br>G� til TimeOut ved at <a href='https://outzource.dk/timeout_xp/wwwroot/ver2_10/login.asp?lto="&lto&"&tomobjid="&jobid&"'>klikke her..</a>"
                                   '&key="&strLicenskey&"
                                    if jobAnsAnsat = "1" then
                                           if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                           strBody = strBody &"<br><br><br>"&job_txt_027&":<br>https://outzource.dk/"&lto&"/default.asp?tomobjid="&jobid
                                           else
                                           strBody = strBody &"<br><br><br>"&job_txt_027&":<br>https://timeout.cloud/"&lto&"/default.asp?tomobjid="&jobid
                                           end if
                                    end if
                                    
                                  


                                    strBody = strBody &"<br><br><br><br><br><br>"&job_txt_028&"<br><i>" 
		                            strBody = strBody & session("user") & "</i><br><br>&nbsp;"



                                    select case lto 
                                    case "hestia"
            		                strBody = strBody &"<br><br><br><br>_______________________________________________________________________________________________<br>"
                                    strBody = strBody &"HESTIA Ejendomme, Rosen�rnsgade 6, st., 8900 Randers C, Tlf. 70269010 - www.hestia.as<br><br>&nbsp;"
                                    end select
            		


                                    myMail.HTMLBody= "<html><head></head><body>" & strBody & "</body>"

                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                    'Name or IP of remote SMTP server
                                    
                                                if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                                   smtpServer = "webout.smtp.nu"
                                                else
                                                   smtpServer = "formrelay.rackhosting.com" 
                                                end if
                    
                                                myMail.Configuration.Fields.Item _
                                                ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                    'Server port
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                    myMail.Configuration.Fields.Update

                                    if isNull(jobAnsThisEmail) <> true then

                                    if len(trim(jobAnsThisEmail)) <> 0 then
                                    myMail.Send
                                    end if

                                    end if

                                    set myMail=nothing

            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing



                                    'Response.end

                                    end if
                              
                                    next
				            
                            end if 'c drev





                    end if' adviser


                    '******** Fast advisering Dencker
                    if lto = "dencker" AND func = "dbopr" then

                         if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\jobs.asp" then

					  	            Set myMail=CreateObject("CDO.Message")
                                    myMail.From="timeout_no_reply@outzource.dk" 'TimeOut Email Service 

                                     strSQL = "SELECT job.id AS jid, jobnavn, jobnr, job.beskrivelse, job_internbesk, k.kkundenavn "_
                                     &" FROM job "_
				                     &" LEFT JOIN kunder AS k ON (k.kid = job.jobknr)"_
				                     &" WHERE job.id = "& varJobId

                          
				                    oRec5.open strSQL, oConn, 3
				                    if not oRec5.EOF then

                                    jobnavnThis = oRec5("jobnavn")
                                    intJobnr = oRec5("jobnr")
                                    strkkundenavn = oRec5("kkundenavn") 

                                    end if
                                    oRec5.close

                                  
                                    myMail.To= "Dencker - Ordre<ordre@dencker.net>"
                                    myMail.Bcc= "Dencker - Ordre<sk@outzource.dk>"
                                    myMail.Subject= "Ny ordre: "& jobnavnThis &" ("& intJobnr &") | " & strkkundenavn  


		                            strBody = "<br>"
                                    strBody = strBody &"<b>Kunde:</b> "& strkkundenavn & "<br>" 
						            strBody = strBody &"<b>Job:</b> "& jobnavnThis &" ("& intJobnr &") <br><br>"

                                   
                                    if len(trim(strBesk)) <> 0 then
                                    strBody = strBody &"<hr><b>Jobbeskrivelse:</b><br>"
                                    strBody = strBody & strBesk &"<br><br><br><br>"
                                    end if

                                    if len(trim(job_internbesk)) <> 0 then
                                    strBody = strBody &"<hr><b>Intern note:</b><br>"
                                    strBody = strBody & job_internbesk &"<br><br>"
                                    end if


                                    strBody = strBody &"<br><br><br><br><br><br>Med venlig hilsen<br><i>" 
		                            strBody = strBody & session("user") & "</i><br><br>&nbsp;"


                                    'Mailer.BodyText = strBody
                                    myMail.HTMLBody= "<html><head></head><body>" & strBody & "</body>"

                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                    'Name or IP of remote SMTP server
                                    
                                                if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                                   smtpServer = "webout.smtp.nu"
                                                else
                                                   smtpServer = "formrelay.rackhosting.com" 
                                                end if
                    
                                                myMail.Configuration.Fields.Item _
                                                ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                    'Server port
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                    myMail.Configuration.Fields.Update

                                   
                                    myMail.Send
                                   
                                    set myMail=nothing

                        end if
                    end if 'Notificering Dencker
                    '*******************************************************


                       



                            '********************************************************
                            '*** Opret folder ****'
                            '********************************************************
                                if request("FM_opret_folder") <> "" then

                                                    
                                            '**** Opretter ikke fysisk en folder endnu ***'
                                        
                                             

                                           
                                            select case lto 
                                            case "xdencker", "xintranet - local"
                                            'Set objFSO = server.createobject("Scripting.FileSystemObject")
                                            'folderPath = "\\195.189.130.210\" 'timeout_xp\wwwroot\"
                                            'Set objnewFolder = objFSO.CreateFolder(folderPath & "DER1")
                                            'Set objnewFolder = nothing






                                            'ServerShare = "\\outzource.dk\dencker_job"
                                            'ServerShare = "\\remote.dencker.net\JOB_Timeout"
                                            ServerShare = "\\remote.dencker.net\to_job4"
                                            
                                            
                                            
                                            'UserName = "Administrator"
                                            UserName = "ad"
                                            'UserName = "TimeOut"
                                            'Password = "Sok!2637"
                                            Password = "ad1996"
                                            'Password = "betina00"
                                            'Password = "timeout"
                                            'Password = ""

                                            Set NetworkObject = CreateObject("WScript.Network")
                                            Set FSO = CreateObject("Scripting.FileSystemObject")

                                            NetworkObject.MapNetworkDrive "", ServerShare, False ', UserName, Password

                                            Set objnewFolder = FSO.CreateFolder(ServerShare &"/"& strNavn &"_"& strjnr)

                                            Set Directory = FSO.GetFolder(ServerShare)
                                            For Each FileName In Directory.Files
                                                response.write FileName.Name & "<br>"
                                                response.write fso.GetParentFolderName(FileName.Name) & "<br>"
                                            Next

                                            Set FileName = Nothing
                                            Set Directory = Nothing
                                            Set FSO = Nothing
                                            Set objnewFolder = nothing

                                            response.write job_txt_315 &"!"
                                            response.end

                                            end select
                         
            	                            oprFysisk = 1
                                            if oprFysisk = 1000 then

	                                                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\jobs.asp" then
            							
		                                                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
		                                                Set objNewFile = nothing
		                                                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
            	
	                                                else
            		
		                                                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, false)
		                                                Set objNewFile = nothing
		                                                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8, -1)
            		
	                                                end if
            	
	                                        file = "ressexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
            	
            	
	                                        objF.WriteLine(ekspTxt)


                                            objF.close

                                            end if

                                        '*** Insert into db folder ***'

                                        sqldatoDt = year(now) &"/"& month(now) &"/"& day(now)
                                        strSQL = "INSERT INTO foldere (navn, kundeid, kundese, jobid, editor, dato) VALUES "_
		                                &" ('"& strNavn &"_"& strjnr &"', "& oRec("kid") &", 0, "& varJobid &", '"& session("user") &"', '"& sqldatoDt &"')"
		                                
		
		
		
		                                oConn.execute(strSQL)
		


                                 end if

							
							
							'**** Tilknytter / Sletter / Opdaterer Uleverand�rer ***'
							
							'strSQLDelUlev = "DELETE FROM job_ulev_ju WHERE ju_jobid = "& varJobId
							'oConn.execute(strSQLDelUlev)

                             call budgetakt_fn()

							
							For u = 1 to 50
							    
                                if len(trim(request("ulevid_"&u&""))) <> 0 then
                                ulevid = request("ulevid_"&u&"")
                                else
                                ulevid = 0
                                end if

                                if len(trim(request("ulevmatid_"&u&""))) <> 0 then
                                ulevmatid = request("ulevmatid_"&u&"")
                                else
                                ulevmatid = 0
                                end if

                                ulevfase = "" 'replace(request("ulevfase_"&u&""), "'", "''")
							    ulevnavn = replace(request("ulevnavn_"&u&""), "'", "''")
							    
                                ulevstk = replace(request("ulevstk_"&u&""), ".", "")
							    ulevstk = replace(ulevstk, ",", ".")

                                ulevstkpris = replace(request("ulevstkpris_"&u&""), ".", "")
							    ulevstkpris = replace(ulevstkpris, ",", ".")

							    ulevpris = replace(request("ulevpris_"&u&""), ".", "")
							    ulevpris = replace(ulevpris, ",", ".")
							    
							    ulevfaktor = replace(request("ulevfaktor_"&u&""), ".", "")
							    ulevfaktor = replace(ulevfaktor, ",", ".")
							    
							    ulevbelob = replace(request("ulevbelob_"&u&""), ".", "")
							    ulevbelob = replace(ulevbelob, ",", ".")

                                ulevkonto = request("ulevkonto_"&u&"")
							    if len(trim(ulevkonto)) <> 0 then
                                ulevkonto = ulevkonto
                                else
                        
                                    if cint(budgetakt) = 2 then
                                    ulevkonto = "" '0
                                    else
                                    ulevkonto = 0
                                    end if

                                end if
    							
    							

                     
                     
    							
		                     if len(trim(ulevnavn)) <> 0 then


                                        if len(trim(ulevstk)) <> 0 then
    							        ulevstk = ulevstk
    							    
    							            call erDetInt(ulevstk)
    							            if isInt = 0 then
    							            ulevstk = ulevstk
    							            else
    							            ulevstk = 0
    							            end if
    							            isInt = 0
    							    
    							        else
    							        ulevstk = 0
    							        end if
                                        
                                        if len(trim(ulevstkpris)) <> 0 then
    							        ulevstkpris = ulevstkpris
    							    
    							            call erDetInt(ulevstkpris)
    							            if isInt = 0 then
    							            ulevstkpris = ulevstkpris
    							            else
    							            ulevstkpris = 0
    							            end if
    							            isInt = 0
    							    
    							        else
    							        ulevstkpris = 0
    							        end if
                                            
					            
					                    if len(trim(ulevpris)) <> 0 then
    							        ulevpris = ulevpris
    							    
    							            call erDetInt(ulevpris)
    							            if isInt = 0 then
    							            ulevpris = ulevpris
    							            else
    							            ulevpris = 0
    							            end if
    							            isInt = 0
    							    
    							        else
    							        ulevpris = 0
    							        end if
    							
    							
    							        if len(trim(ulevfaktor)) <> 0 then
    							        ulevfaktor = ulevfaktor
    							    
    							            call erDetInt(ulevfaktor)
    							            if isInt = 0 then
    							            ulevfaktor = ulevfaktor
    							            else
    							            ulevfaktor = 0
    							            end if
    							            isInt = 0
    							    
    							        else
    							        ulevfaktor = 0
    							        end if
    							
    							        if len(trim(ulevbelob)) <> 0 then
    							        ulevbelob = ulevbelob
    							        
    							            call erDetInt(ulevbelob)
    							            if isInt = 0 then
    							            ulevbelob = ulevbelob
    							            else
    							            ulevbelob = 0
    							            end if
    							            isInt = 0
    							        
    							        else
    							        ulevbelob = 0
    							        end if
    							

                                            if cdbl(ulevid) <> 0 then

                                            strSQLUpdUlev = "UPDATE job_ulev_ju SET ju_date = '"& ddupdate &"', ju_editor = '"& strEditor &"', "_
							                &" ju_fase = '"& ulevfase &"', ju_navn = '"& ulevnavn &"', ju_ipris = "& ulevpris &", "_
							                &" ju_faktor = "& ulevfaktor &", ju_belob = "& ulevbelob &",  ju_jobid = "& varJobId & ", ju_stk = "& ulevstk &", ju_stkpris = "& ulevstkpris &", ju_matid = "& ulevmatid &", "
                        
                                             if cint(budgetakt) = 2 then 
                                             strSQLUpdUlev = strSQLUpdUlev & " ju_konto_label = '"& ulevkonto &"' WHERE ju_id = " & ulevid
                                             else
                                             strSQLUpdUlev = strSQLUpdUlev & " ju_konto = "& ulevkonto &" WHERE ju_id = " & ulevid
						                     end if	                

                                            'Response.Write strSQLInsUlev & "<br>"
                                            oConn.execute(strSQLUpdUlev)


                                            else
    							
						                    strSQLInsUlev = "INSERT INTO job_ulev_ju SET ju_date = '"& ddupdate &"', ju_editor = '"& strEditor &"',"_
							                &" ju_fase = '"& ulevfase &"', ju_navn = '"& ulevnavn &"', ju_ipris = "& ulevpris &", "_
							                &" ju_faktor = "& ulevfaktor &", ju_belob = "& ulevbelob &",  ju_jobid = "& varJobId& ", ju_stk = "& ulevstk &", ju_stkpris = "& ulevstkpris & ", ju_matid = "& ulevmatid &", "
                        
                                             
                                             if cint(budgetakt) = 2 then 
                                             strSQLInsUlev = strSQLInsUlev & " ju_konto_label = '"& ulevkonto &"'" ' WHERE ju_id = " & ulevid
                                             else
                                             strSQLInsUlev = strSQLInsUlev & " ju_konto = "& ulevkonto &"" 'WHERE ju_id = " & ulevid
						                     end if	
                                    
                        

                                            'if session("mid") = 1 then
                                            'Response.Write "ulevid:" & ulevid & " -- "&  strSQLInsUlev & "<br>"
                                            'Response.end
                                            'end if
                        
                                            oConn.execute(strSQLInsUlev)

                                            end if
    							
							   
                               
                               else
                                        
                                         if cint(ulevid) <> 0 then

                                         ulevDel = "DELETE FROM job_ulev_ju WHERE ju_id = " & ulevid
                                         oConn.execute(ulevDel)

                                         end if
                               
                               
                               end if
    							
							  
							
							next
							
						    'Response.end
							




							
					    '*****************************************************************************'
					    '*** tilf�jer job i timereg_usejob (Vis guide), b�de ved opret og rediger ****'
					    '*****************************************************************************'
						
                        'repsonse.write "her"
                        'response.end

                                '**** Aktive aktiviteter p� dette job

                  

                                progrpFindespaaAkt10 = 0
                                aktiveAktids = "#0#"
                                strSQlAkt = "SELECT id FROM aktiviteter WHERE job = " & varJobId & " AND aktstatus = 1"
                                oRec4.open strSQlAkt, oConn, 3
                                while not oRec4.EOF
                                    
                                    for ai = 1 to 10

                                        select case ai
                                        case 1
                                        pid = strProjektgr1
                                        case 2
                                        pid = strProjektgr2
                                        case 3
                                        pid = strProjektgr3
                                        case 4
                                        pid = strProjektgr4
                                        case 5
                                        pid = strProjektgr5
                                        case 6
                                        pid = strProjektgr6
                                        case 7
                                        pid = strProjektgr7
                                        case 8
                                        pid = strProjektgr8
                                        case 9
                                        pid = strProjektgr9
                                        case 10
                                        pid = strProjektgr9
                                        end select                

                                        call projktgrpPaaAktids(pid, oRec4("id"))
                                        progrpFindespaaAkt10 = progrpFindespaaAkt10 & progrpFindespaaAktivitet

                                    next

                                    if instr(progrpFindespaaAkt10, "1") <> 0 then '001001100 Prokjektgruppen findes p� aktiviteten som progruppe 1-10
                                    aktiveAktids = aktiveAktids & ",#" & oRec4("id") & "#"  
                                    end if

                                oRec4.movenext
                                wend
                                oRec4.close

                            

                         
							medarUseJobWrt = ""
							dtNow = year(now) & "-"& month(now) & "-"& day(now)

                      
                      'Tester hastighed n�r EPI redigere elelr opretter job. 
                      'Optimer nedenst�ende f.eks v�lge alle medarbejdere p� job og tjek med instr.
                      'strSQLex = "SELECT id, aktid FROM timereg_usejob WHERE jobid = "& varJobId &" AND medarb = " & oRec5("MedarbejderId") & " AND aktid = 0"
                     
                     'if instr(lto, "epi") <> 0 then
                     'opdaterTimereg_usejob = 0
                     'else
                     opdaterTimereg_usejob = 1
                     'end if

                    
                        
                    if cint(opdaterTimereg_usejob) = 1 then
							

                                '** KUN AKTIVE og passive medarbjedere
                                select case lto
                                case "epi2017"
                                mansatSQLkri = "m.mansat = 1" 'kun aktive
                                case else
                                mansatSQLkri = "m.mansat = 1 OR m.mansat = 3" 'Aktive og passive
                                end select


								strSQL = "SELECT DISTINCT(MedarbejderId), mansat, mid FROM progrupperelationer "_
                                &" LEFT JOIN medarbejdere m ON (m.mid = MedarbejderId AND ("& mansatSQLkri &")) WHERE ("_
								&" ProjektgruppeId = "& strProjektgr1 &""_
								&" OR ProjektgruppeId =" & strProjektgr2 &""_
								&" OR ProjektgruppeId =" & strProjektgr3 &""_
								&" OR ProjektgruppeId =" & strProjektgr4 &""_
								&" OR ProjektgruppeId =" & strProjektgr5 &""_
								&" OR ProjektgruppeId =" & strProjektgr6 &""_
								&" OR ProjektgruppeId =" & strProjektgr7 &""_
								&" OR ProjektgruppeId =" & strProjektgr8 &""_
								&" OR ProjektgruppeId =" & strProjektgr9 &""_
								&" OR ProjektgruppeId =" & strProjektgr10 &""_
								&") AND mid IS NOT NULL GROUP BY MedarbejderId"
								
                                'if session("mid") = 1 then
                                'Response.Write "strSQL "& strSQL & "<br><hr>"
                                'Response.end
                                'end if
								
								oRec5.open strSQL, oConn, 3
								while not oRec5.EOF

                                    
                                
                                 medarbfundet = 0

                                    
                                            
                                            '** Henter b�de job og specifikke aktiviter
                                            strSQLex = "SELECT id, aktid FROM timereg_usejob WHERE jobid = "& varJobId &" AND medarb = " & oRec5("MedarbejderId") & " AND aktid = 0"

                                             'if session("mid") = 1 then
                                             'Response.Write "strSQL "& strSQLex & "<br><hr>"
                                             'Response.end
                                             'end if
                                    
                                            oRec4.open strSQLex, oConn, 3
                                            if not oRec4.EOF then
                                            medarbfundet = 1
                                            
                                                    if cint(forvalgt) = 1 then 's�t aktiv for alle tilmeldte medarbejdere i valgte projektgrupper

                                    
                                                        strSQLfvlgt = "UPDATE timereg_usejob SET forvalgt = 1, forvalgt_sortorder = 0, forvalgt_af = "& session("mid") &", forvalgt_dt = '"& dtNow &"', favorit = 1 WHERE id = " & oRec4("id") 
                                                        oConn.execute(strSQLfvlgt)
                                                        'Response.Write strSQLfvlgt & "<br>"


                                                        'Aktiviteten er blevet lukket/passiv. Sletter aldrig den p� jobbet aktid = 0
                                                        strSQLfvlgtDel = "DELETE FROM timereg_usejob WHERE jobid = " & varJobId & " AND medarb = " & oRec5("MedarbejderId") & " AND aktid <> 0"
                                                        oConn.execute(strSQLfvlgtDel)

                                                        aktids = replace(aktiveAktids, "#", "")
                                                        aktidsarr = split(aktids, ",")
                                
                                                        for a = 1 to UBOUND(aktidsarr)


                                                                strSQL3 = "INSERT INTO timereg_usejob (medarb, jobid, forvalgt, forvalgt_sortorder, forvalgt_af, forvalgt_dt, favorit, aktid) VALUES "_
                                                                &" ("& oRec5("MedarbejderId") &", "& varJobId &", 1, 0, "& session("mid") &", '"& dtNow &"', 1, "& aktidsarr(a) &")"
                                                                oConn.execute(strSQL3)

                                                        next


                                                    end if 'forvalgt

                                            End if
                                            oRec4.close



                                 

                                    if cint(medarbfundet) = 0 then 
                                        
                                        if cint(forvalgt) = 1 then 's�t aktiv
                                       

                                                strSQL3 = "INSERT INTO timereg_usejob (medarb, jobid, forvalgt, forvalgt_sortorder, forvalgt_af, forvalgt_dt, favorit, aktid) VALUES "_
                                                &" ("& oRec5("MedarbejderId") &", "& varJobId &", 1, 0, "& session("mid") &", '"& dtNow &"', 1, 0)"
                                                oConn.execute(strSQL3)

                                                aktids = replace(aktiveAktids, "#", "")
                                                aktidsarr = split(aktids, ",")
                                
                                                for a = 1 to UBOUND(aktidsarr)


                                                        strSQL3 = "INSERT INTO timereg_usejob (medarb, jobid, forvalgt, forvalgt_sortorder, forvalgt_af, forvalgt_dt, favorit, aktid) VALUES "_
                                                        &" ("& oRec5("MedarbejderId") &", "& varJobId &", 1, 0, "& session("mid") &", '"& dtNow &"', 1, "& aktidsarr(a) &")"
                                                        oConn.execute(strSQL3)

                                                next

                                        
                                        else
                                        
                                                '**** S�tter i jobbanken. G�lder ALLE undt WWF der har manuel positiv aktivering (Kontrolpanel)

                                                 'if instr(lto, "epi") <> 0 then
                                                 'opdaterTimereg_usejob3 = 0
                                                 'else
                                                 opdaterTimereg_usejob3 = 1
                                                 'end if


                                                if cint(opdaterTimereg_usejob3) = 1 then

                                                    call positiv_aktivering_akt_fn() 'wwf
                                                    if cint(positiv_aktivering_akt_val) <> 1 then
                                                    strSQL3 = "INSERT INTO timereg_usejob (medarb, jobid) VALUES ("& oRec5("MedarbejderId") &", "& varJobId &")"
                                                    oConn.execute(strSQL3)
                                                    end if

                                                end if


                                        end if

                                   
                                   
									'Response.Write strSQL3 & "<br>"
									'Response.Flush
                                    end if
									
                                    
                                    'Response.End 
									
								     medarUseJobWrt = medarUseJobWrt & " AND medarb <> "& oRec5("MedarbejderId")
									
									
								oRec5.movenext
								wend
								oRec5.close

                                'Response.End

							
                            


							if func = "dbred" then
							    '*** Sletter ikke l�ngere aktuelle medarbejdere fra timereg usejob **'
                                'if len(trim(medarUseJobWrt)) <> 0 then
                                if lto = "tia" then
							    strSQLdelUseJob = "DELETE FROM timereg_usejob WHERE jobid = "& varJobId & "" & medarUseJobWrt & " AND favorit <> 1"
                                else
                                strSQLdelUseJob = "DELETE FROM timereg_usejob WHERE jobid = "& varJobId & "" & medarUseJobWrt & ""
                                end if

                                'Response.Write "<br><br>"& strSQLdelUseJob

                                oConn.execute(strSQLdelUseJob)
							    'end if
                            end if



                            end if 'opdaterTimereg_usejob = 1 insrt(epi)




                            '**** S�tter EASYreg aktiv p� aktiviteter for alle medarbejdere ****'
                            call showEasyreg_fn()
                            if cint(showEasyreg_val) = 1 then
                            oEasyReg = 0
                            strSQLea = "SELECT easyreg, id FROM aktiviteter WHERE job = "& id & " AND easyreg = 1"
				            oRec5.open strSQLea, oConn, 3
				            while not oRec5.EOF 
				                
                              oEasyReg = 1
                            
                            oRec5.movenext
				            wend 
				            oRec5.close
				                

                                if cint(oEasyReg) = 1 then
				                strSQLtreguse = "UPDATE timereg_usejob SET easyreg = " & id & " WHERE jobid = " & id
	   	                        'Response.Write strSQLtreguse & "<br>"
	   	                        'Response.flush
	   	                        oConn.execute(strSQLtreguse)
				                end if
				               
				            'Response.end
				         
					        end if 'showEasyreg_val


                            



					        '************************************************'
							' Opdaterer prjgp (fjerner) p� aktiviteter      *'
							' s� der ikke kan v�re progp p� aktiviteter     *'
							' der ikke ogs� findes p� jobbet.               *'
							'************************************************'
							if func = "xxx" then	
                            '** Sl�et fra da det giver problemer iforhold til nedarv. B�DE ved opret og rediger job, f.eks ved F�D job med nye stamaktivitetyer ved rediger job	
							'** Funktionen kan dog bruges til oprydning ***'		
							
							
								    strSQL = "SELECT id, job, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter WHERE job = "& varJobId
								    'Response.write strSQL
								    oRec5.open strSQL, oConn, 3
								    
									while not oRec5.EOF 
									
									call pjgcase(oRec5("projektgruppe1"), 1)
									call pjgcase(oRec5("projektgruppe2"), 2)
									call pjgcase(oRec5("projektgruppe3"), 3)
									call pjgcase(oRec5("projektgruppe4"), 4)
									call pjgcase(oRec5("projektgruppe5"), 5)
									call pjgcase(oRec5("projektgruppe6"), 6)
									call pjgcase(oRec5("projektgruppe7"), 7)
									call pjgcase(oRec5("projektgruppe8"), 8)
									call pjgcase(oRec5("projektgruppe9"), 9)
									call pjgcase(oRec5("projektgruppe10"), 10)
									
									oRec5.movenext
									wend 
								    oRec5.close
							
					
					        end if
					
					

                    '******************************************************************************************'
                    '**** Opdaterer medarbejder timeprsier til at f�lge budgetprisen p� aktivitets linjerne ***'
                    '******************************************************************************************'
                    if cint(laasmedtpbudget) = 1 then

                   
                        call opd_aktfasttp(id,opdekstp,valuta)




                    end if



                    showFordelpFinansaar = request("showFordelpFinansaar")
                    if func = "dbred" AND cint(showFordelpFinansaar) = 1 then
                    '***************************************
                    '**** BUDGET FY ************************
                    '***************************************
                    f = 0
                    if len(trim(request("antal_usefybudgetingt"))) <> 0 then
                    FyMaksAar = request("antal_usefybudgetingt")
                    else
                    FyMaksAar = 0
                    end if

                    for f = 0 to FyMaksAar

                    if len(trim(request("FM_fy_aar_"& f))) <> 0 then
                    FY0 = request("FM_fy_aar_"& f)
                    else
                    FY0 = "2001"
                    end if
                    'FY1 = request("FM_fy_aar_1")
                    'FY2 = request("FM_fy_aar_2")
                    fctimeprisFY0 = 0
                    fctimeprish2FY0 = 0

                    if len(trim(request("FM_fy_hours_" & f))) <> 0 then
                    jobBudgetFY0 = request("FM_fy_hours_" & f)
                    else
                    jobBudgetFY0 = 0
                    end if

                    if len(trim(request("FM_fy_belob_" & f))) <> 0 then
                    jobBudgetBelobFY0 = request("FM_fy_belob_" & f)
                    else
                    jobBudgetBelobFY0 = 0
                    end if
                    'jobBudgetFY2 = request("FM_fy_hours_2")
                    jobids = id
                    aktid = 0



                   
                        call opdaterRessouceRamme_job(f, FY0, jobBudgetFY0, jobBudgetBelobFY0, jobids, aktid)
                    next


                    'response.end
                    end if
					

                    '********************************'
                    '***** Forretningsomr�der ******'
                    '********************************'
                    
                    '*** nulstiller job ****'
                    strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& varJobId & " AND for_aktid = 0"
                    oConn.execute(strSQLfor)


                    '*** IKKE MERE 11.3.2015 (aktiviteter t�lles altid med p� job) ***'
                    '*** nulstiller akt ved sync ****'
                    'if cint(syncForAkt) = 1 then
                    '    strSQLfor = "DELETE FROM fomr_rel WHERE for_jobid = "& varJobId & " AND for_aktid <> 0"
                    '    oConn.execute(strSQLfor)
                    'end if

                    'Response.Write "her"
                    'Response.Flush

                    for afor = 0 to UBOUND(fomrArr)

                            'Response.Write "her2" & afor & "<br>"
                            'Response.Flush

                            if fomrArr(afor) <> 0 then

                            strSQLfomri = "INSERT INTO fomr_rel "_
                            &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                            &" VALUES ("& fomrArr(afor) &", "& varJobId &", 0, "& for_faktor &")"

                            oConn.execute(strSQLfomri)

                                '*** Sync aktiviteter ****'
                                'if cint(syncForAkt) = 1 then
                                'strSQLa = "SELECT id FROM aktiviteter WHERE job = "& varJobid
                                'oRec3.open strSQLa, oConn, 3
                                'while not oRec3.EOF 

                                '    strSQLfomrai = "INSERT INTO fomr_rel "_
                                '    &" (for_fomr, for_jobid, for_aktid, for_faktor) "_
                                '    &" VALUES ("& fomrArr(afor) &", "& varJobId &","& oRec3("id") &", "& for_faktor &")"

                                '    oConn.execute(strSQLfomrai)

                                'oRec3.movenext
                                'wend
                                'oRec3.close

                                'end if

                            end if


                    next

                    '********************************'

					
                    end if 'jobnummer findes		
							
					oRec.movenext
					wend
					oRec.close
                    '*** End for next Multible opret p� kunder ***'
					
                    'Response.Write "func:" & func
                    'Response.end	
							
							 
							 
							 if len(request("fm_kunde_sog")) <> 0 then
									vmenukundefilt = request("fm_kunde_sog")
									else
									vmenukundefilt = 0
									end if
									
									if len(trim(request("jobnr_sog"))) <> 0 then
									strJobsog = request("jobnr_sog")
									else
									strJobsog = left(strNavn, 10)
									end if
							
							
							if request("FM_usetilbudsnr") = "j" then
							filt = "tilbud"
							else
							filt = request("filt") 
							end if
							

                             if len(trim(request("visrealtimerdetal"))) <> 0 AND request("visrealtimerdetal") <> 0 then
                             visrealtimerdetal = 1
                             else
                             visrealtimerdetal = 0
                             end if
							
							
                          
							
                            
							'timeB = now
                            'loadtime = datediff("s",timeA, timeB)
							'Response.write loadtime
                            'response.flush
							'Response.Write "<br>" & rdir
							'Response.end
						
                            
                        	


							select case request("rdir")
							case "redjobcontionue"
							rdirLink = "jobs.asp?menu=job&func=red&id="&varJobId&"&jobnr_sog="&strJobsog&"&filt="&filt&"&fm_kunde="&strKundeId&"&fm_kunde_sog="&vmenukundefilt&"&showdiv="&showdiv&"&FM_tp_jobaktid="&tp_jobaktid&"&FM_mtype="&mtype&"&visrealtimerdetal="&visrealtimerdetal
                            'Response.Write rdirLink
                            'Response.end
                            case "jbpla_w"
							strKlokkeslet = request("sttid")
							dato = request("dato")
							jobstKri = request("datostkri")
							jobslKri = request("datoslkri")
							rdirLink = "jbpla_w_opr.asp?func=step1&sttid="&strKlokkeslet&"&dato="&dato&"&datostkri="&jobstKri&"&datoslkri="&jobslKri&"&id=0&step=2&jobid="&varJobId&"&stepialt=3"
							case "sdsk"
							rdirLink = "<script language=""JavaScript"">window.opener.location.reload();</script>"
							rdirLink = rdirLink & "<script language=""JavaScript"">window.close();</script>"
							case "sdsk2"
							rdirLink = "sdsk.asp?func=red&id=0"
							case "webblik"
							rdirLink = "../to_2015/webblik_joblisten.asp"
							case "webblik_tilfakturering"
							rdirLink = "webblik_tilfakturering.asp"
							case "webblik_milepale"
							'rdirLink = "webblik_milepale.asp"
                            rdirLink = "<script language=""JavaScript"">window.opener.location.href(""webblik_milepale.asp"");</script>"
							rdirLink = rdirLink & "<script language=""JavaScript"">window.close();</script>"
							case "pg"
							rdirLink = "pg_allokering.asp?menu=job"
							case "kon"
							rdirLink = "kunder.asp?menu=kund"
							case "pipe"
							'rdirLink = "pipeline.asp?menu=job"
                            rdirLink = "<script language=""JavaScript"">window.opener.location.reload();</script>"
							rdirLink = rdirLink & "<script language=""JavaScript"">window.close();</script>"
							case "treg"	
						    rdirLink = "timereg_akt_2006.asp?showakt=1"
												
							case "seraft"
							rdirLink = "kunder.asp?menu=kund&func=red&id="&strKundeId&"&showdiv=seraft"
							
							case else
							
							'**til joblisten
						    'rdirLink = "jobs.asp?menu=job&shokselector=1&id="&varJobId&"&jobnr_sog="&strJobsog&"&filt="&filt&"&fm_kunde="&vmenukundefilt
							
                            '*** Skal v�re s�ledes at listyen ikke bliver vist med det samme, da f.eks Hestia opfatter det som en fejl at den st�r l�nge og venter p� listen (1400 job)
                            if func = "dbopr" then
                            rdirLink = "../to_2015/job_list.asp?hidelist=1"
                            else
                            rdirLink = "../to_2015/job_list.asp?hidelist=1" '?id="&varJobId&"&jobnr_sog="&strJobsog&"&filt="&filt&"&fm_kunde="&vmenukundefilt
                            end if

                          end select
							
						  
						    '*****************************************'
				            '******* Opret opgave Inceidnt i SDSK ****'
				            '*****************************************'
            				
				             if len(request("FM_opr_incident")) <> 0 then
            							    
				                if func = "dbred" then
				                kidforSDSK = strKundeId
				                else
				                kidforSDSK = strKidUseOne 'trim(kids(0))
				                end if
				            
				            Response.Write("<script language=""JavaScript"">window.open('sdsk.asp?func=opr1&id=0&FM_kontakt="&kidforSDSK&"&FM_emne="&strNavn&" ("& strjnr &")"&"','sdskwin');</script>")
                            %>
                            <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
                            
                            <body>
                            <%
                            itop = 100
                            ileft = 150
                            iwdt = 400
                            call sideinfo(itop,ileft,iwdt)
                            
                             %>
                            
                            
			                       <tr><td style="padding:10px;">
                                   <%=job_txt_029 %> <br /><a href="<%=rdirLink %>" class=vmenu><%=job_txt_030 %>..</a>
                            
	                        </td></tr></table>
	                        </div>
	                        <br><br><br><br><br>&nbsp;
                            </body>
                            
                            
						    <%
						    Response.end
						    else
						    
						     if rdir = "sdsk" OR rdir = "pipe" or rdir = "webblik_milepale" then
						     response.Write rdirLink
						     else
						     response.Redirect rdirLink
						     end if
						     
						    end if
				
                '** Validering		  
						
					'end if
				'end if
			  end if
			 end if
			end if
		end if
	end if
		
	
	
	
	
	case "opret", "red"
	'*** Tjekker om det er fra ressourcebooking/popup at jobbet bliver oprettet ***
	showaspopup = request("showaspopup")
	headergif = "../ill/job_oprred_header.gif"
	
	
	
	
	if showaspopup <> "y" then%>
	
	
	
	
	<%
	if len(trim(request("step"))) <> 0 then
	step = request("step")
	else
		if request("func") = "red" then
		step = 2
		else
		step = 1
		end if
	end if
	
	
	if request("FM_kunde") = "0" AND step = 2 then
	%>
					 <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					 
					 
					 
			          <%
					errortype = 92
					call showError(errortype)
					isInt = 0
			        Response.End
			        
			        
	end if
	
    
    %>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    
    <%if step = 2 then %>
    <script src="inc/job_jav_2020_3.js"></script>
    <%else %>
    <script src="inc/job_listen_jav2.js"></script>
    <%end if 
    
    
    if nomenu <> "1" then%>
    
    

    <!--
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(3)%>
	</div>

	<div id="sekmenu" style="position:absolute; left:15px; top:82px; visibility:visible;">
	<%
	call jobtopmenu()
	%>
    </div>
    <%end if%>
        -->
	
	<%call menu_2014() %>



	<%else%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
    
	<%
	'level = 1
	'editok = 1 %>
	
	
	<%end if 'showaspopup%>
	<!--#include file="inc/job_inc.asp"-->
	<%
	
	if showaspopup <> "y" then
    call top
	varbroedkrumme = job_txt_031
	
	else
	varbroedkrumme = job_txt_032
	%>
		<br>&nbsp;&nbsp;&nbsp;&nbsp;<font class="stor-blaa">job_txt_033</font>
	<%end if

	
	

	
	
	
	'*** Placering af side div ***
	if showaspopup <> "y" then
	sidedivLeft = 90

    if step = 1 then
    sidedivTop = 132
    else
	sidedivTop = 152
	end if  
      
	    if func <> "red" then
	    
	    if step = 1 then
	    mainTableWidth = 675
	    else
	    mainTableWidth = 475
	    end if
	    
	    
	    select case step
	    case 1
	    varbroedkrumme = "<h3 class=""hv"">"&job_txt_034&" <span style='font-size:10px;'> - "&job_txt_035&"</span></h3>"
	    case 2  
	    varbroedkrumme = "<h3 class=""hv"">"&job_txt_036&" <span style='font-size:10px;'> - "&job_txt_316&"</span></h3>"
	    end select
	    
	    else
	    
	    varbroedkrumme = "<h3 class=""hv"">Jobstamdata<span style='font-size:10px;'> - "&job_txt_037&"</span></h3>"
	    mainTableWidth = 475
	    
	    end if
	    
	else
	step = 2
	stdata_vzb = "visible"
	stdata_dsp = ""
	sidedivLeft = 90
	sidedivTop = 132
	mainTableWidth = 475
	'varbroedkrumme = varbroedkrumme & " (Ekspres oprettelse)"
	end if
	
     '************ Faneblade ******************
    if func = "red" then 'AND strInternt = "1" 
    'strJobnr & strKnr = request("FM_kunde")
    strSQLfb = "SELECT id, jobnavn, jobnr, jobknr FROM job, kunder WHERE id = " & id 
	
	'Response.Write strSQLfb
	'Response.flush
	
	oRec.open strSQLfb, oConn, 3
	
	if not oRec.EOF then
	fbNavn = oRec("jobnavn")
	fbjobnr = oRec("jobnr")
	'fbKnavn = oRec("kkundenavn")
	fbKnr = oRec("jobknr")
    
	end if
    oRec.close
    
    
    call faneblade("job", showdiv)
    end if
	

    if func = "opret" AND step = 2 then
	%>
	<!-- manual --->
    <div id="help" style="position:absolute; background-color:#ffffff; padding:1px 1px 0px 1px; width:170px; border:1px silver solid; border-bottom:0px; left:1020px; top:135px; visibility:visible; display:; z-index:9000000; overflow:hidden;">
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
    <tr bgcolor="#FF6666"><td align=center>
    <%
    select case lto
    case "execon" %>
    <a href="../help_and_faq/TimeOut_job_oprettelse_rev_100924_execon.pdf" class=alt target="_blank"><%=job_txt_038 %>..</a>
    <%case "wwf"
    %>
    <a href="../help_and_faq/TimeOut_job_oprettelse_rev_110523_wwf.pdf" class=alt target="_blank"><%=job_txt_038 %>..</a>
    <%
    case else %>
    <a href="../help_and_faq/TimeOut_job_oprettelse_rev_110523.pdf" class=alt target="_blank"><%=job_txt_038 %>..</a>
    <%end select%>

    
    
        </td>
     </tr>
    </table>
	</div>
    <%end if %>

    <%select case lto
    case "epi", "epi_cati", "epi_no", "epi_sta", "epi_se", "xintranet - local", "epi_uk"
    %>
      <!-- Popup MSG -->
	<div id="mtyp_msg1" style="position:absolute; top:200px; left:1250px; width:275px; border:1px #999999 solid; z-index:900000000; background-color:#FFFFe1; padding:10px 10px 10px 10px; visibility:hidden; display:none;">
    <b><%=job_txt_039 %></b>, <%=job_txt_040 %> <br /><br />
     <%=job_txt_317 %><br /><br />
        <b>1) <%=job_txt_041 %></b> <%=job_txt_042 %><br /><br />
        <b>2) <%=job_txt_043 %></b> <%=job_txt_044 %><br /><br />
        <span id="lukms1" style="color:#ea0957; font-size:9px;">[luk]</span>
	</div>
	
     <!-- Popup MSG -->
	<div id="mtyp_msg2" style="position:absolute; top:200px; left:1250px; width:275px; border:1px #999999 solid; z-index:900000000; background-color:#FFFFe1; padding:10px; visibility:hidden; display:none;">
    <b><%=job_txt_046 %></b><%=" " & job_txt_047 %><br /><br />
          <span id="lukms2" style="color:#ea0957; font-size:9px;">[<%=job_txt_045 %>]</span>
	</div>
	
    <%
    case else
    end select %>

  



    <!-------------------------------Sideindhold------------------------------------->

    <form action="jobs.asp?menu=job" method="post" name="jobdata" id="jobdata">
	<div id="sindhold" style="position:absolute; left:<%=sidedivLeft%>px; top:<%=sidedivTop%>px; visibility:<%=stdata_vzb%>; display:<%=stdata_dsp%>; z-index:50;">
	
	
	<%
    
   
	
	
	

    tTop = 0
	tLeft = 0
	tWdth = mainTableWidth
	
	
	call tableDiv_j(tTop,tLeft,tWdth)
	%>

    <table cellspacing="0" cellpadding="0" border="0" bgcolor="#FFFFFF" width="100%">
	
    <input id="showdiv" name="showdiv" value="<%=showdiv%>" type="hidden" />
	<!--<input type="hidden" name="func" id="func" value="<=dbfunc%>">-->
	<input type="hidden" name="fm_kunde_sog" value="<%=request("fm_kunde_sog")%>">
	<input type="hidden" id="rdir" name="rdir" value="<%=request("rdir")%>">
	<input type="hidden" id="jq_jobid" name="id" value="<%=id%>">
	<input type="hidden" name="int" id="int" value="<%=strInternt%>">
	<input type="hidden" name="jobnr_sog" value="<%=request("jobnr_sog")%>">
	<input type="hidden" name="filt" value="<%=request("filt")%>">
	<input type="hidden" name="showaspopup" value="<%=showaspopup%>">
    <input type="hidden" id="jq_func" name="jq_func" value="<%=func%>">
    <input type="hidden" id="lto" name="lto" value="<%=lto%>">
	
	<tr bgcolor="#5582D2">
		<td width="8" rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=2 valign="top"><img src="../ill/blank.gif" width="<%=(mainTableWidth-16)%>" height="1" alt="" border="0"></td>
		<td align=right rowspan="2" valign=top><img src="../ill/blank.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td colspan=2 class="alt" valign="top" style="padding-top:4px;"><%=varbroedkrumme %></td>
	</tr>
	<%
	select case step
	case "1" 
	editok = 1
	
	%>
	<input type="hidden" name="step" id="step" value="<%=step+1%>">
	<input type="hidden" name="func" id="func" value="<%=func%>">
	
	<input type="hidden" name="FM_fastpris1" id="FM_fastpris1" value="0">
	
	
	<%if showaspopup = "y" then%>
	<input type="hidden" name="sttid" value="<%=request("sttid")%>">
	<input type="hidden" name="dato" value="<%=request("dato")%>">
	<input type="hidden" name="datostkri" value="<%=request("datostkri")%>">
	<input type="hidden" name="datoslkri" value="<%=request("datoslkri")%>">
	<%end if%>
	<%
	
	call kundeopl


	
	
	case "2" '"3"
	

    



	
	if func = "opret" then
	
	%>
	<input type="hidden" name="step" value="<%=step+1%>">
	<!-- Fra Step 1 og 2 -->
	<input type="hidden" name="FM_kunde" value="<%=request("FM_kunde")%>">
	<!--
	<input type="hidden" name="FM_kundese" value="<%=request("FM_kundese")%>">
	<input type="hidden" name="FM_kundese_hv" value="<%=request("FM_kundese_hv")%>">
	-->
	
	
	<%
	
	end if
	
	if func = "opret" AND step = 2 OR func = "red" then
	sendmailtokunde = "onClick='mailtokunde()'"
	else
	sendmailtokunde = ""
	end if
	
	thisMid = session("mid")

	
	                '*** Er der p� licebns niveau mulighe for at lukke og afmelde job? ***'
	                lukafm = 0
					strSQLlicens = "SELECT lukafm FROM licens WHERE id = 1"
					oRec5.open strSQLlicens, oConn, 3
					if not oRec5.EOF then
					
					lukafm = oRec5("lukafm")
					
					end if
					oRec5.close 
	
	
   
     call jobopr_mandatory_fn()

	'*** Her indl�ses form til rediger/oprettelse af job ***
	if func = "opret" then
	
	strjobnr = 0
	strtilbudsnr = 0
    strNexttilbudsnr = 0
	strNavn = job_txt_055 & ".."
	
	dbfunc = "dbopr"
	strFakturerbart = 1
	rowspan = "28"

	    select case lto
	    case "intranet - local", "synergi1", "cisu", "wilke", "epi2017"
	    strFastpris = 1 'default fastpris
	    case else
	    strFastpris = 2 'default l�bende timer
	    end select
	
    intkundeok = 0
	varSubVal = "opretpil"
	

    '** Forvalgte projektgrupper *********************************************************
	'if len(request.cookies("job")("defaultprojgrp")) <> 0 AND lto <> "execon" then 
	'strProj_1 = request.cookies("job")("defaultprojgrp")
	'else

	    select case lto 
        case "execon"
        strProj_1 = 1
        case "cisu", "synergi1"
        strProj_1 = 1
        case else
        strProj_1 = 10
        end select

	'end if
	
      select case lto 
        
        case "cisu" '** F�der job fra de priojektgrupper der er tilknyttet aktiviteterne A-E
        strProj_2 = 1
	    strProj_3 = 1
	    strProj_4 = 1
	    strProj_5 = 1
	    strProj_6 = 1
	    strProj_7 = 1
	    strProj_8 = 1
	    strProj_9 = 1
	    strProj_10 = 1
        case else
        strProj_2 = 1
	    strProj_3 = 1
	    strProj_4 = 1
	    strProj_5 = 1
	    strProj_6 = 1
	    strProj_7 = 1
	    strProj_8 = 1
	    strProj_9 = 1
	    strProj_10 = 1
        end select


	
    '****************************************************************************************

	
	intkundekpers = 0
	
	editok = 1
	
	ikkeBudgettimer = 0
	strBudgettimer = 0
	
	rekvnr = ""
	
	ski = 0
	abo = 0
	ubv = 0
	
	'*** Mulitible opret *** (bruges til kontakt personer) **'
	

	                
				    strKundeId = 0
				    kids = split(request("FM_kunde"), ",")
				    for x = 0 to UBOUND(kids)
				        
				        if x > 1 then
				        strKundeId = 0
				        else
				        strKundeId = kids(x)
				        end if
				        
				        if x > 1 then
				        multibletrue = 1
				        else
				        multibletrue = 0
				        end if
				    
				    next
				    
				   
	
	
	
	laasmedtpbudget = 0
	intServiceaft = 0
	
	if cint(lukafm) <> 0 then
	lukafmjob = 1
	else 
	lukafmjob = 0
	end if
	
	

	jobfaktype = 0
	jo_gnstpris = 0

    select case lto 
    case "epi", "epi_no", "epi_sta", "outz", "epi_ab", "epi_cati", "epi_uk", "epi2017"
    jo_gnsfaktor = 2
	case else
    jo_gnsfaktor = 1
    end select

    jo_gnsbelob = 0
    jo_bruttofortj = 0
    jo_dbproc = 0
    
    udgifter = 0
    usejoborakt_tp = 0
    

    if len(trim(strKundeId)) <> 0 then
    strKundeId = strKundeId
    else
    strKundeId = 0
    end if


    jobans1 = 0
	
    select case lto
    case "epi2017", "intranet - local"

        '* Jobans 2 NEDARVES FRA KUNDE kudneans1 
        strSQLforvalgt_kans = "SELECT kundeans1 FROM kunder WHERE kid  = "& strKundeId
        oRec.open strSQLforvalgt_kans, oConn, 3
        if not oRec.EOF then
        
        jobans2 = oRec("kundeans1")
       
        end if
        oRec.close

    case "mpt"
    jobans2 = 4
    case else
    jobans2 = 0
    end select

    

    select case lto 
    case "fe"
    intprio = 1
    case "dencker"
        if cint(strKundeId) = 1 then
        intprio = -3
        else
        intprio = 100
        end if        
    case else
    intprio = 100
    end select
    
  
    
    sandsynlighed = 0
    jfak_moms = 1
    jfak_sprog = 1
    sdskpriogrp = 0
    valuta = basisValId
    jo_valuta = valuta
   
    if multibletrue = 0 then
    
    '** er der en SDSK priogruppe, valuta,moms og fak sprog p� kunde ****
    strSQL = "SELECT sdskpriogrp, kfak_moms, kfak_sprog, kfak_valuta, lincensindehaver_faknr_prioritet FROM kunder WHERE kid = "& strKundeId
    oRec.open strSQL, oConn, 3
    if not oRec.EOF then
    sdskpriogrp = oRec("sdskpriogrp")
    jfak_moms = oRec("kfak_moms")
    jfak_sprog = oRec("kfak_sprog")
    valuta = oRec("kfak_valuta")
    lincensindehaver_faknr_prioritet = oRec("lincensindehaver_faknr_prioritet")
    end if
    oRec.close

        '* NEDARVES FRA KUNDE Licensindehaver VALUTA p� faktura
        strSQLforvalgt_jo_valuta = "SELECT kfak_valuta FROM kunder WHERE lincensindehaver_faknr_prioritet = "& lincensindehaver_faknr_prioritet &" AND useasfak = 1"
        oRec.open strSQLforvalgt_jo_valuta, oConn, 3
        if not oRec.EOF then
        
        jo_valuta = oRec("kfak_valuta")
       
        end if
        oRec.close
    
    end if 



   



    restestimat = 0
    

    select case lto
    case "epi", "epi_no", "epi_sta", "intranet - local", "epi_ab", "epi_cati", "epi_uk", "epi2017"
	virksomheds_proc = 50
	syncslutdato = 0 '1
    intSandsynlighed = 10
    stade_tim_proc = 1
    case else
    virksomheds_proc = 0
	syncslutdato = 0
    intSandsynlighed = 0
    stade_tim_proc = 0
    end select
    
    altfakadrCHK = ""
    altfakadr = 0

    preconditions_met = 0 '1 = ja, 0= ikke angivet, 2 = nej vent

    filepath1 = ""
    filepath2 = ""
    filepath3 = ""
    fomr_konto = 0

    useasfak = 0

    jo_valuta_kurs = 100
   
    select case lto
    case "intranet - local", "cisu"
    jo_usefybudgetingt = 1
    case else
    jo_usefybudgetingt = 0
    end select

    lincensindehaver_faknr_prioritet_job = "0"
    project_tier = ""

	else '*** REDIGER JOB *****'
    
    vlgtmtypgrp = 0
    call mtyperIGrp_fn(vlgtmtypgrp,0) 
    call fn_medarbtyper()
    

	strSQL = "SELECT id, jobnavn, jobnr, kkundenavn, jobknr, "_
	&" jobTpris, jobstatus, jobstartdato, jobslutdato, projektgruppe1, projektgruppe2, "_
	&" projektgruppe3, projektgruppe4, projektgruppe5, job.dato, job.editor, creator, createdate, "_
	&" fakturerbart, budgettimer, fastpris, kundeok, job.beskrivelse, "_
	&" ikkeBudgettimer, tilbudsnr, projektgruppe6, projektgruppe7, "_
	&" projektgruppe8, projektgruppe9, projektgruppe10, job.serviceaft, "_
	&" kundekpers, jobans1, jobans2, lukafmjob, valuta, jobfaktype, rekvnr, "_
	&" jo_gnstpris, jo_gnsfaktor, jo_gnsbelob, jo_bruttofortj, jo_dbproc, "_
	&" udgifter, risiko, sdskpriogrp, usejoborakt_tp, ski, job_internbesk, abo, ubv, sandsynlighed, "_
    &" diff_timer, diff_sum, jo_udgifter_ulev, jo_udgifter_intern, jo_bruttooms, restestimat, stade_tim_proc, virksomheds_proc, "_
    &" syncslutdato, lukkedato, altfakadr, preconditions_met, laasmedtpbudget, filepath1, filepath2, filepath3, fomr_konto, jfak_moms, jfak_sprog, useasfak, alert,"_
    &" lincensindehaver_faknr_prioritet_job, jo_valuta, jo_valuta_kurs, jo_usefybudgetingt, categories_process, data_outside_eu, safeguard, gdpr_safeguard_io, gdpr_personaldata, gdpr_projecttype, project_tier, "_
    &" gdpr_personaldatavendor, gdpr_strippedforpersonal_data, gdpr_personal, gdpr_sensitive"_
    &" FROM job, kunder WHERE id = " & id &" AND kunder.Kid = jobknr"
	
        'if session("mid") = 1 then
	    'Response.Write strSQL
	    'Response.flush
	    'end if

	oRec.open strSQL, oConn, 3
	
	if not oRec.EOF then
	strNavn = oRec("jobnavn")
	strjobnr = oRec("jobnr")
	strKnavn = oRec("kkundenavn")
	strKnr = oRec("jobknr")
	strBudget = oRec("jo_bruttooms") 'oRec("jobTpris")
	strStatus = oRec("jobstatus")
	strTdato = oRec("jobstartdato")
	strUdato = oRec("jobslutdato")
	strProj_1 = oRec("projektgruppe1")
	strProj_2 = oRec("projektgruppe2")
	strProj_3 = oRec("projektgruppe3")
	strProj_4 = oRec("projektgruppe4")
	strProj_5 = oRec("projektgruppe5")
	strProj_6 = oRec("projektgruppe6")
	strProj_7 = oRec("projektgruppe7")
	strProj_8 = oRec("projektgruppe8")
	strProj_9 = oRec("projektgruppe9")
	strProj_10 = oRec("projektgruppe10")
	strDato = oRec("dato")
	strLastUptDato = oRec("dato")
	strEditor = oRec("editor")
    strCreatedDate = oRec("createdate")
    strCreator = oRec("creator")
	strFakturerbart = oRec("fakturerbart")
	
	
	strFastpris = oRec("fastpris")
	intkundeok = oRec("kundeok")
	strBesk = oRec("beskrivelse")


    useasfak = oRec("useasfak")
    
        if cdbl(oRec("tilbudsnr")) = 0 then
	    
        strtilbudsnr = oRec("tilbudsnr")

	    strSQLtb = "SELECT tilbudsnr FROM licens WHERE id = 1"
	    oRec3.open strSQLtb, oConn, 3
        if not oRec3.EOF then

        strNexttilbudsnr = oRec3("tilbudsnr") + 1
        tbStyle = "color:#999999;"

        end if
        oRec3.close

        else

        strtilbudsnr = oRec("tilbudsnr")
        strNexttilbudsnr = strtilbudsnr
	    tbStyle = "color:#000000;"
        end if
    
    intServiceaft = oRec("serviceaft")
	intkundekpers = oRec("kundekpers")
	
	jobans1 = oRec("jobans1")
	jobans2 = oRec("jobans2")

   
	
	'*** Ikke fakturerbare timer bruges ikke mere, men gemmes pga. gamle job ***'
	strBudgettimer = oRec("budgettimer") + oRec("ikkeBudgettimer")
	
	
		if oRec("ikkeBudgettimer") > 0 then
		ikkeBudgettimer = oRec("ikkeBudgettimer")
		else
		ikkeBudgettimer = 0
		end if

	
	
	lukafmjob = oRec("lukafmjob") 
	
    valuta = oRec("valuta")
    jobfaktype = oRec("jobfaktype")
	
	rekvnr = oRec("rekvnr")
	
	jo_gnstpris = oRec("jo_gnstpris")
	jo_gnsfaktor = oRec("jo_gnsfaktor")
	jo_gnsbelob = oRec("jo_gnsbelob")
	jo_bruttofortj = oRec("jo_bruttofortj")
	jo_dbproc = oRec("jo_dbproc")
	
    jo_udgifter_ulev = oRec("jo_udgifter_ulev")
    jo_udgifter_intern = oRec("jo_udgifter_intern")

    jo_bruttooms = oRec("jo_bruttooms")

	udgifter = oRec("udgifter")
	
	intprio = oRec("risiko")
	sdskpriogrp = oRec("sdskpriogrp")
	
	usejoborakt_tp = oRec("usejoborakt_tp")
	
	ski = oRec("ski")
	abo = oRec("abo")
	ubv = oRec("ubv")
	job_internbesk = oRec("job_internbesk")
	
	intSandsynlighed = oRec("sandsynlighed")

    diff_timer = oRec("diff_timer")
    diff_sum = oRec("diff_sum")
	
    restestimat = oRec("restestimat")
    stade_tim_proc = oRec("stade_tim_proc")

    virksomheds_proc = oRec("virksomheds_proc")

    syncslutdato = oRec("syncslutdato")
    lkdato = oRec("lukkedato")
    altfakadr = oRec("altfakadr")

    preconditions_met = oRec("preconditions_met")
    laasmedtpbudget = oRec("laasmedtpbudget")

  
    if len(trim(oRec("filepath1"))) <> 0 then
    filepath1 = oRec("filepath1")
    filepath1 = replace(filepath1, "#", "\")
    else
    filepath1 = "" 
    end if

     if len(trim(oRec("filepath2"))) <> 0 then
    filepath2 = oRec("filepath2")
    filepath2 = replace(filepath2, "#", "\")
    else
    filepath2 = "" 
    end if

    if len(trim(oRec("filepath3"))) <> 0 then
    filepath3 = oRec("filepath3")
    filepath3 = replace(filepath3, "#", "\")
    else
    filepath3 = "" 
    end if

    fomr_konto = oRec("fomr_konto")

    jfak_moms = oRec("jfak_moms")
    jfak_sprog = oRec("jfak_sprog")

	job_internbesk_alert = oRec("alert")
    lincensindehaver_faknr_prioritet_job = oRec("lincensindehaver_faknr_prioritet_job")
    jo_valuta = oRec("jo_valuta")
    jo_valuta_kurs = oRec("jo_valuta_kurs")

    jo_usefybudgetingt = oRec("jo_usefybudgetingt")

    categories_process = oRec("categories_process")
    data_outside_eu = oRec("data_outside_eu")
    safeguard = oRec("safeguard")

        gdpr_safeguard_io = oRec("gdpr_safeguard_io") 
        gdpr_personaldata = oRec("gdpr_personaldata") 
        gdpr_projecttype = oRec("gdpr_projecttype")

        project_tier = oRec("project_tier")

        gdpr_personaldatavendor = oRec("gdpr_personaldatavendor")
        gdpr_strippedforpersonal_data = oRec("gdpr_strippedforpersonal_data")


    gdpr_personal = oRec("gdpr_personal") 
    gdpr_sensitive = oRec("gdpr_sensitive")

    end if
	oRec.close
	
	
	headergif = "../ill/job_oprred_header.gif"
	dbfunc = "dbred"
	varSubVal = "opdaterpil" 
	rowspan = "28"
	
	strKundeId = strKnr
	
	editok = 0
	
	
	end if





    if cint(jo_usefybudgetingt) <> 0 then
    jo_usefybudgetingtCHK = "CHECKED"
    else
    jo_usefybudgetingtCHK = ""
    end if

    if cint(jo_valuta) = 0 then
    call basisValutaFN()
    jo_valuta = basisValId
    end if

    '** S�tter valuta til job budget
    call valutakode_fn(jo_valuta)
    'basisValISO = valutaKode_CCC
    public jo_bgt_basisValISO
    jo_bgt_basisValISO = valutaKode_CCC 
    jo_bgt_basisValISO_f8 = valutaKode_CCC_f8 
	
	if func = "red" then



            if altfakadr = 0 then
            altfakadrCHK = ""
            else
            altfakadrCHK = "CHECKED"
            end if

	
	        if jobfaktype = 0 then
	        jobfaktypeCHK0 = "CHECKED"
	        jobfaktypeCHK1 = ""
	        else
	        jobfaktypeCHK0 = ""
	        jobfaktypeCHK1 = "CHECKED"
	        end if 

            if cint(job_internbesk_alert) = 1 then
            alertCHK = "CHECKED"
            else
            alertCHK = ""
            end if
	
	else
	    
        alertCHK = ""


	    if lto = "dencker" then
	        
	        jobfaktypeCHK0 = ""
	        jobfaktypeCHK1 = "CHECKED"
	    
	    else
	    
	        if request.Cookies("tsa")("faktype") <> "" then
    	        
	            if request.Cookies("tsa")("faktype") = "0" then
	            jobfaktypeCHK0 = "CHECKED"
	            jobfaktypeCHK1 = ""
	            else
	            jobfaktypeCHK0 = ""
	            jobfaktypeCHK1 = "CHECKED"
	            end if 
    	        
	        else
    	        
	            jobfaktypeCHK0 = "CHECKED"
	            jobfaktypeCHK1 = ""
    	    
	        end if
	    
	    end if
	
	end if
	
	
	%>
	<!--#include file="inc/dato2.asp"-->
	<%
	'****************************************************************
	'*** Tjekker admin rettigeheder eller om man er jobanssvarlig ***
	'****************************************************************
	
	
	if level <= 2 OR level = 6 then
	editok = 1
	else
			if cint(session("mid")) = jobans1 OR cint(session("mid")) = jobans2 OR (cint(jobans1) = 0 AND cint(jobans2) = 0) then
			editok = 1
			end if
	end if
	
	
	if editok = 1 then
	
	
	
		if cint(intkundeok) = 1 OR cint(intkundeok) = 2 then
		kundechk = "checked"
		else
		kundechk = ""
		end if
	
	
		if strFakturerbart = 1 then
		varFakEks = "checked"
		varFakInt = ""
		else
		varFakEks = ""
		varFakInt = "checked"
		end if
		
		if strFastpris = "1" then
		varFastpris1 = "checked"
		varFastpris2 = ""
		else
		varFastpris1 = ""
		varFastpris2 = "checked"
		end if
	
	
	%>
    <input type="hidden" name="slto" id="slto" value="<%=lto %>" />
           
            
	<input type="hidden" name="func" id="func" value="<%=dbfunc%>">
	<input type="hidden" name="FM_OLDjobnr" value="<%=strJobnr%>">
	
	<%if showaspopup = "y" then%>
	<input type="hidden" name="sttid" value="<%=request("sttid")%>">
	<input type="hidden" name="dato" value="<%=request("dato")%>">
	<input type="hidden" name="datostkri" value="<%=request("datostkri")%>">
	<input type="hidden" name="datoslkri" value="<%=request("datoslkri")%>">
	<%end if%>
	
	<!--<input type="hidden" name="seraft" value="<=intServiceaft%>">-->
	    
    <%
	    if func = "red" then
	        if len(request("fm_kunde_sog")) <> 0 then
		    fm_kunde_sog = request("fm_kunde_sog")
		    else
		    fm_kunde_sog = 0
		    end if
		end if
	
		rowspan = rowspan
		if func = "opret" then
		tp_height = "1529"
		else
		tp_height = "1529"
		end if
	

              
	
	if request("showoprtxt") = "j" then 'AND request("int") <> 0 then%>
	<div id=visopretmed name=visopretmed style="position:absolute; left:140; top:20; width:300px; visibility:visible; z-index:100; background-color:#fffff1; border:1px darkred solid; padding:3px;">
	<font color="ForestGreen"><b><%=job_txt_048 %>!</b></font><br> <%=job_txt_049 &" " %><a href="jobs.asp?menu=job&shokselector=1" class=vmenu><%=job_txt_050 %></a><%=" " & job_txt_051 %><br>
	<br><br>
	<a href="#" onClick="hideoprmed()" class=red>[X] <%=job_txt_052 %></a> 
	</div>
	<%end if%>
	
    <%if func = "red" then%>
	<tr>
		<td colspan=4 bgcolor="#ffffff" height="30" style="  padding-top:5px; padding-left:20px;">
        <%if isDate(strCreatedDate) = true AND strCreatedDate <> "01-01-2002" then %>
            <%=job_txt_630 %> <b><%=strCreatedDate %></b> <%=job_txt_054 %> <b><%=strCreator %></b>
            <br />
        <%end if %>

        <%if isdate(strLastUptDato) = true then %>
		<%=job_txt_053 %> <b><%=formatdatetime(strLastUptDato, 2)%></b><%=" "&job_txt_054&" " %><b><%=strEditor%></b>
        <%else %>
           strLastUptDato: <%=strLastUptDato %> Err. NOT VALID date format.
        <%end if %>
		</td>
	</tr>
    <%end if%>

 
        <%
        '**** kundeoplysninger, eksternt job *****'
	    call kundeopl
	
	    '******** Kontaktpersoner ***********'
	    call kundeopl_aft_kpers
        %>
       
	
	
	
	<tr bgcolor="#FFFFFF">
		
		<td style="padding:10px 5px 10px 10px; white-space:nowrap;" colspan="4">
            <font color=red size=2>*</font> <b><%=job_txt_055 %>:</b> 
        
                <!--
                <br /><a href="file:///C:\SK\skole\Fra Historie Online.docx" target="_blank">C:\SK\skole\Fra Historie Online.docx</a>

                <div class="fileupload">
                <input type="file" class="file" name="f1" style="width:400px;" />
                </div>
                <input  type="text" id="myInputText" value="C:\SK\skole\Fra Historie Online.docx" />
                -->

        <%if func <> "red" then %>
        (<%=job_txt_056 %>)
        <%end if %>

        <br />
        <%if func <> "red" then 
        'focol = "#999999"
        twd = 420
        else
        'focol = "#000000"
        twd = 290
        end if%>
		<input type="text" name="FM_navn" id="FM_navn" value="<%=strNavn%>" style="width:<%=twd%>px; border:1px red solid; font-size:14px; padding:3px;">
        <%if func = "red" then %>
        <font color=red size=2>*</font> <%=job_txt_057 %>: <input type="text" name="FM_jnr" id="FM_jnr" value="<%=strJobnr%>" style="width:80px; color:#999999; font-style:italic;">
        <%else %>
        <input type="hidden" name="FM_jnr" id="Text2" value="0">
        <%end if %>
        
        

        <br /><span style="font-size:10px; font-family:arial; color:#999999;"><%=job_txt_058 &" " %><%=maxCharJobNavn %><%=" "& job_txt_059 %>. "<%=" "&job_txt_060 %></span>
 
       
      
                        <%
                        
                        select case lto
                        case "synergi1", "intranet - local", "outz", "dencker", "essens", "jttek"
                        
                        if func <> "red" then
                        opFolderCHK = "CHECKED"
                        else
                        opFolderCHK = ""
                        end if

                        case else

                        opFolderCHK = ""

                        end select

                        if len(trim(id)) <> 0 then
                        joidThis = id
                        else
                        joidThis = 0
                        end if

                        strSQLdok = "SELECT fo.navn FROM foldere fo  "_
                        &" WHERE fo.jobid =" & joidThis & " AND fo.jobid <> 0" 

                        folderfindes = 0
                        oRec2.open strSQLdok, oConn, 3
                        if not oRec2.EOF then
                        folderfindes = 1
                        end if
                        oRec2.close
                        
                        
                        if cint(folderfindes) = 0 then %>
                        <br />
                        <input type="checkbox" value="1" name="FM_opret_folder" <%=opFolderCHK %> /> <%=job_txt_061 %>
                        

                        <%end if %>


        </td>
	</tr>


        
        <%select case lto
        case "intranet - local", "dencker"%>
        <tr>
		<td style="">&nbsp;</td>
		<td valign=top style="padding:5px 5px 10px 0px;" colspan=2><b>Project tier:</b><br /><input type="text" name="project_tier" value="<%=project_tier %>" style="width:300px;" />
            </td>
	     </tr>
        <%case else %>
        <input type="hidden" name="project_tier" value="<%=project_tier %>" />
        <%end select %>

	<tr>
		<td style="">&nbsp;</td>
		<td valign=top style="padding:5px 5px 10px 0px;" colspan=2>
		<%if func = "red" then%>
		<b><%=job_txt_062 %>:</b><br />
		<%else %>
		<b><%=job_txt_063 %>?</b><br /> 
		<%end if %>
		
		        <%  if cint(strStatus) = 3 OR (func = "opret" AND cint(tilbud_mandatoryOn) = 1) then
					chkusetb = "CHECKED"
					else
					chkusetb = ""
					end if
					%>
					
					<input type="checkbox" id="FM_usetilbudsnr" name="FM_usetilbudsnr" value="j" <%=chkusetb%>><%=job_txt_064 %>
		
	
			
					<%if func = "red" then
                        
                     if strNexttilbudsnr <> strtilbudsnr then 
                     tlbplcholderTxt = strNexttilbudsnr

                     %>
                    &nbsp;<span style="color:#999999">(<%=job_txt_065 %>: <%=tlbplcholderTxt %>)</span>
                    <%

                     else
                    tlbplcholderTxt = ""
                    end if %>

                   

					<br /><input type="text" id="FM_tnr" name="FM_tnr" value="<%=strtilbudsnr%>" style="<%=tbStyle%>" size="20"> 
                  
					<%else%>
					<input type="hidden" name="FM_tnr" value="0">
					<%end if%>
				
					<input type="hidden" id="FM_nexttnr" value="<%=strNexttilbudsnr %>">

                    <%select case lto
                    case "epi2017", "intranet - local"
                        sandBdr = "border:1px red solid;"
                    case else
                        sandBdr = ""
                    end select %>

					&nbsp;&nbsp;<input id="Text1" name="FM_sandsynlighed" value="<%=intSandsynlighed %>" type="text" style="width:30px; <%=sandBdr%>"/> <%="% "& job_txt_066 %> &nbsp;
            <br /><span style="font-size:10px; font-family:arial; color:#999999;">(<%=job_txt_067 &" " %>=<%=" "& job_txt_068 &" " %>*<%=" "& job_txt_069 %>)</span>
            
	       
	        </td>
            <td>&nbsp;</td>
	        </tr>	

  
                        <%
                            
                        showfilepath_gdpr = 1
                           
                        if cint(showfilepath_gdpr) = 1 then
                           
                        select case lto 
                        case "essens", "intranet - local", "epi2017"

                            select case lto
                            case "epi2017", "intranet - local"
                            filepathTxt = "GDPR information" 
                            case else
                            filepathTxt = job_txt_070 
                            end select 

                           %> 
                        <tr><td colspan="4" style="padding:20px 20px 20px 20px;">
                        <div style="padding:10px 10px 10px 10px; background-color:#faf0f0;">    
                        <b><%=filepathTxt %>:</b><br />
                        Link to contract <font color=red size=2>*</font> <input type="text" name="FM_filepath1" placeholder="URL" style="width:410px;" value="<%=filepath1 %>" /> <br />


                         <%
                        gdpr_projecttype0SEL = ""
                        gdpr_projecttype1SEL = ""
                        gdpr_projecttype2SEL = ""
                        select case gdpr_projecttype 
                        case 0
                            gdpr_projecttype0SEL = "SELECTED"
                            gdpr_projecttype1SEL = ""
                            gdpr_projecttype2SEL = ""
                        case 1
                            gdpr_projecttype0SEL = ""
                            gdpr_projecttype1SEL = "SELECTED"
                            gdpr_projecttype2SEL = ""
                        case 2
                            gdpr_projecttype0SEL = ""
                            gdpr_projecttype1SEL = ""
                            gdpr_projecttype2SEL = "SELECTED"
                        end select%>
                        
                        <br />
                        Project type <font color=red size=2>(*)</font> <select name="FM_gdpr_projecttype">
                            <option value="0" <%=gdpr_projecttype0SEL %>>Ad-hoc (Consulting)</option>
                            <option value="1" <%=gdpr_projecttype1SEL %>>Data collection</option>
                            <option value="2" <%=gdpr_projecttype2SEL %>>Hybrid</option>
                        </select><br />
                            <br />
                            

                        <%
                        gdpr_personaldata0SEL = ""
                        gdpr_personaldata1SEL = ""
                        gdpr_personaldata2SEL = ""
                        select case gdpr_personaldata 
                        case 0
                            gdpr_personaldata0SEL = "SELECTED"
                            gdpr_personaldata1SEL = ""
                            gdpr_personaldata2SEL = ""
                        case 1
                            gdpr_personaldata0SEL = ""
                            gdpr_personaldata1SEL = "SELECTED"
                            gdpr_personaldata2SEL = ""
                        case 2
                            gdpr_personaldata0SEL = ""
                            gdpr_personaldata1SEL = ""
                            gdpr_personaldata2SEL = "SELECTED"
                        end select
                        %>
                        
                        <br />
                          <div style="padding-bottom:2px">
                        Personal Data <font color=red size=2>(*)</font> <select id="FM_gdpr_personaldata" name="FM_gdpr_personaldata">
                            <option value="0" <%=gdpr_personaldata0SEL %>>Dataprosessor (add text below)</option>
                            <option value="1" <%=gdpr_personaldata1SEL %>>Data controller</option>
                            <option value="2" <%=gdpr_personaldata2SEL %>>No personal data</option>
                        </select>
                              </div>
                       
                       <div id="dv_filepath2">
                           <%if session("mid") = 1 OR session("mid") = 32821 OR session("mid") = 33407 then 

                              end if 


                               gdpr_personaldatavendor1SEL = ""
                               gdpr_personaldatavendor2SEL = ""
                               gdpr_personaldatavendor3SEL = ""
                               gdpr_personaldatavendor4SEL = ""
                               gdpr_personaldatavendor5SEL = ""
                               gdpr_personaldatavendor6SEL = ""
                               gdpr_personaldatavendor7SEL = ""
                               gdpr_personaldatavendor8SEL = ""
                               gdpr_personaldatavendor9SEL = ""
                               gdpr_personaldatavendor10SEL = ""
                               gdpr_personaldatavendor11SEL = ""
                               gdpr_personaldatavendor12SEL = ""
                               
                               'if instr(gdpr_personaldatavendor, "01") <> 0 then
                               'gdpr_personaldatavendor1SEL = "SELECTED"
                               'end if
                               'if instr(gdpr_personaldatavendor, "02") <> 0 then
                               'gdpr_personaldatavendor2SEL = "SELECTED"
                               'end if
                               'if instr(gdpr_personaldatavendor, "3") <> 0 then
                               'gdpr_personaldatavendor3SEL = "SELECTED"
                               'end if
                               'if instr(gdpr_personaldatavendor, "4") <> 0 then
                               'gdpr_personaldatavendor4SEL = "SELECTED"
                               ' end if
                               'if instr(gdpr_personaldatavendor, "5") <> 0 then 
                               'gdpr_personaldatavendor5SEL = "SELECTED"
                               ' end if
                               'if instr(gdpr_personaldatavendor, "6") <> 0 then
                               'gdpr_personaldatavendor6SEL = "SELECTED"
                               ' end if
                               'if instr(gdpr_personaldatavendor, "7") <> 0 then
                               'gdpr_personaldatavendor7SEL = "SELECTED"
                               ' end if
                               'if instr(gdpr_personaldatavendor, "8") <> 0 then
                               'gdpr_personaldatavendor8SEL = "SELECTED"
                               ' end if
                               'if instr(gdpr_personaldatavendor, "9") <> 0 then
                               'gdpr_personaldatavendor9SEL = "SELECTED"
                               ' end if
                               'if instr(gdpr_personaldatavendor, "10") <> 0 then
                               'gdpr_personaldatavendor10SEL = "SELECTED"
                               ' end if
                               'if instr(gdpr_personaldatavendor, "11") <> 0 then
                               'gdpr_personaldatavendor11SEL = "SELECTED"
                               'end if
                               'if instr(gdpr_personaldatavendor, "12") <> 0 then
                               'gdpr_personaldatavendor12SEL = "SELECTED"
                               'end if
                               
                               
                               %>
                           <br />
                           Data Vendor:<font color=red size=2>(*)</font><br /><select name="FM_gdpr_personaldatavendor" size="13" style="width:412px;" multiple>
                            <option value="0" <%=gdpr_personaldatavendor0SEL %>>Vendor (add text below)</option>
                            <!--
                                <option value="01" <%=gdpr_personaldatavendor1SEL %>>Marius Pedersen (Shredding of sensitive material)</option>
                            <option value="02" <%=gdpr_personaldatavendor2SEL %>>Microsoft (Operates cloud service O365 and Azure, that is used for e.g. E-mail and server operations)</option>
                               <option value="3" <%=gdpr_personaldatavendor3SEL %>>Timengo DPG A/S (Backup of server data)</option>
                               -->

                               <%strSQLdatavendors = "SELECT id, navn FROM selectboxoptions WHERE selectboxlist = 0 ORDER BY id"
                                   oRec2.open strSQLdatavendors, oConn, 3 
		                            while not oRec2.EOF 

                                     if instr(gdpr_personaldatavendor, oRec2("id")) <> 0 then
                                       gdpr_personaldatavendorSEL = "SELECTED"
                                       end if
                                    %> 
                                        <option value="<%=oRec2("id") %>" <%=gdpr_personaldatavendorSEL%> ><%=oRec2("navn") %></option>
                                    <%
                                    oRec2.movenext
                                    wend
                                    oRec2.close%>
                               <!--
                               <option value="4" <%=gdpr_personaldatavendor4SEL %>>Compaya A/S (SMS send outs)</option>
                               <option value="5" <%=gdpr_personaldatavendor5SEL %>>Informi GIS A/S (RVU location treatment)</option>
                               <option value="6" <%=gdpr_personaldatavendor6SEL %>>Focus vision (Mobile etnography)</option>
                               <option value="7" <%=gdpr_personaldatavendor7SEL %>>Charlie Tango (KPMG - KMD processes personal data in connection with the parties' agreement of delivery of documents to e-Boks. The processing consists of sending documents received from the customer to e-Boks)</option>
                               <option value="8" <%=gdpr_personaldatavendor8SEL %>>Express Paperbased (Print, scanning, enveloping. Treating name, e-mail and address)</option>
                               <option value="9" <%=gdpr_personaldatavendor9SEL %>>Via scan (Print, scanning, enveloping. Treating name, e-mail and address)</option>
                               <option value="10" <%=gdpr_personaldatavendor10SEL %>>JT Digital Solutions v/ Jan Selchau Thomsen (JTDS receives a merge file from Epinion and then prints data, content and any attached folder. Then letters and content are enveloped and sent to PostNord.)</option>
                               <option value="11" <%=gdpr_personaldatavendor11SEL %>>AWS (Data processing, web hosting and similar services)</option>
                               <option value="12" <%=gdpr_personaldatavendor12SEL %>>Zoom (Mobile ethnography)</option>
                               -->
                        </select>
                        <br /><br />
                           Stripped for personal data
                             

                             <%if cint(gdpr_strippedforpersonal_data) = 0 then
                               gdpr_strippedforpersonal_dataCHK0 = "CHECKED"
                             end if%>
                           <br />
                           <input type="radio" name="FM_gdpr_strippedforpersonal_data" value="0" <%=gdpr_strippedforpersonal_dataCHK0 %> /> N/A

                           <%if cint(gdpr_strippedforpersonal_data) = 1 then
                               gdpr_strippedforpersonal_dataCHK1 = "CHECKED"
                             end if%>
                           <br />
                           <input type="radio" name="FM_gdpr_strippedforpersonal_data" value="1" <%=gdpr_strippedforpersonal_dataCHK1 %>/> Yes 

                            <%if cint(gdpr_strippedforpersonal_data) = 2 then
                               gdpr_strippedforpersonal_dataCHK1 = "CHECKED"
                             end if%>
                           <br />
                           <input type="radio" name="FM_gdpr_strippedforpersonal_data" value="2" <%=gdpr_strippedforpersonal_dataCHK2 %> /> No
                           

                        <br />
                           

                           
                           

                       <br />    
                       Link to Data processor agreement <font color=red size=2>(*)</font><br /><input type="text" name="FM_filepath2" placeholder="URL" style="width:410px;" value="<%=filepath2 %>" /> <br />
                      


                        <br />
                        Categories of processing <font color=red size=2>(*)</font><br />
                           <!--<input type="text" id="FM_categories_process" name="FM_categories_process" style="width:410px;" value="<%=categories_process %>" /> <br />-->
                           <select name="FM_categories_process" id="FM_categories_process"  style="width:410px;">

                               <!--<option value="">Select categories..</option>-->
                           <%   strSQLcatofproc = "SELECT id, navn FROM selectboxoptions WHERE selectboxlist = 1 ORDER BY id"
                                   oRec2.open strSQLcatofproc, oConn, 3 
		                            while not oRec2.EOF 

                                     if instr(categories_process, oRec2("navn")) <> 0 then
                                       gdpr_catofprocSEL = "SELECTED"
                                       end if
                                    %> 
                                        <option value="<%=oRec2("navn") %>" <%=gdpr_catofprocSEL%> ><%=oRec2("navn") %></option>
                                    <%
                                    oRec2.movenext
                                    wend
                                    oRec2.close%> 
                               
                               </select><br />



                           <div id="dv_ordinary_personal_data" style="display:; visibility:visible;">
                           <br /><br />
                           Ordinary personal data
                           <br />
                            <%   strSQLcatofproc = "SELECT id, navn FROM selectboxoptions WHERE selectboxlist = 2 ORDER BY id"
                                   oRec2.open strSQLcatofproc, oConn, 3 
		                            while not oRec2.EOF 
                                    gdpr_personalSEL = ""

                                     if instr(gdpr_personal, "#"& oRec2("id") &"#") <> 0 then
                                       gdpr_personalSEL = "CHECKED"
                                       end if
                                    %> 
                                        <input type="checkbox" name="FM_gdpr_personal" value="#<%=oRec2("id")%>#" <%=gdpr_personalSEL%> /><%=oRec2("navn") %><br />
                                    <%
                                    oRec2.movenext
                                    wend
                                    oRec2.close%> 
                               </div>

                           <div id="dv_sensitive_personal_data" style="display:; visibility:visible;">
                           <br /><br />
                           Sensitive personal data
                           <br />
                            <%   strSQLcatofproc = "SELECT id, navn FROM selectboxoptions WHERE selectboxlist = 3 ORDER BY id"
                                   oRec2.open strSQLcatofproc, oConn, 3 
		                            while not oRec2.EOF 
                                    gdpr_sensitiveSEL = ""

                                     if instr(gdpr_sensitive, "#"& oRec2("id") &"#") <> 0 then
                                       gdpr_sensitiveSEL = "CHECKED"
                                       end if
                                    %> 
                                        <input type="checkbox" name="FM_gdpr_sensitive" value="#<%=oRec2("id")%>#" <%=gdpr_sensitiveSEL%> /><%=oRec2("navn") %><br />
                                    <%
                                    oRec2.movenext
                                    wend
                                    oRec2.close%> 




                                        
                         <%'*************** data_outside_eu *************************
                        data_outside_eu0SEL = ""
                        data_outside_eu1SEL = ""
                        select case data_outside_eu 
                        case 0
                            data_outside_eu0SEL = "SELECTED"
                            data_outside_eu1SEL = ""
                        case 1
                            data_outside_eu0SEL = ""
                            data_outside_eu1SEL = "SELECTED"
                        end select%>
                        
                        <br />
                        Personal data are transferred out side EU <font color=red size=2>(*)</font> <br /><select id="FM_data_outside_eu" name="FM_data_outside_eu">
                            <option value="0" <%=data_outside_eu0SEL %>>No</option>
                            <option value="1" <%=data_outside_eu1SEL %>>Yes</option>
                        </select><br />

                         </div>

                        <div id="dv_gdpr_data_outside_eu">

                        <%
                        gdpr_safeguard_io0SEL = ""
                        gdpr_safeguard_io1SEL = ""
                        select case gdpr_safeguard_io
                        case 0
                            gdpr_safeguard_io0SEL = "SELECTED"
                            gdpr_safeguard_io1SEL = ""
                        case 1
                            gdpr_safeguard_io0SEL = ""
                            gdpr_safeguard_io1SEL = "SELECTED"
                        end select%>
                        
                        <br />
                        
                            <div style="padding-bottom:2px">
                        Safeguards for exceptional transfers of personal data <font color=red size=2>(*)</font> <select name="FM_gdpr_safeguard_io">
                            <option value="0" <%=gdpr_safeguard_io0SEL %>>N/A</option>
                            <option value="1" <%=gdpr_safeguard_io1SEL %>>Yes (add text below)</option>
                        </select>
                        </div>

                        <input type="text" name="FM_safeguard" style="width:410px;" value="<%=safeguard%>" /> <br />


                        <%if level = 1 then %>
                        <br />General description of technical and organisational security measures:<br /> <input type="text" name="FM_filepath3" placeholder="URL" style="width:410px;" value="<%=filepath3 %>" /> <br /><br />
                        <%else%>
                        <input type="hidden" name="FM_filepath3" value="<%=filepath3 %>" />
                        <%end if%>

                        </div> <!-- dv_gdpr_data_outside_eu -->

                        </div>
                        </td></tr>
                        <%
                        end select
        
    end if 'showfilepath & GDPR%>

	
    	
	<tr bgcolor="#FFFFFF">
		<td colspan="4" style="padding:10px 10px 2px 10px;">
		<b><%=job_txt_071 %>?</b> <a href="#" onClick="serviceaft('0', '<%=strKundeId%>', '', '0')" class=vmenu>+<%=" "& job_txt_072 %></a> <br>
		
		<%
		strSQL2 = "SELECT id, enheder, stdato, sldato, status, navn, pris, perafg, "_
		&" advitype, advihvor, erfornyet, varenr, aftalenr FROM serviceaft "_
		&" WHERE kundeid = "& strKundeId &" OR id = "& intServiceaft &" ORDER BY id DESC" 'AND status = 1
		
		'Response.write strSQL2
		%>
		<select name="FM_serviceaft" id="FM_serviceaft" style="width:450px;">
		<option value="0"><%=job_txt_073 %></option>
		
		<%
		
		oRec2.open strSQL2, oConn, 3 
		while not oRec2.EOF 
		
		if oRec2("advitype") <> 0 then
		udldato = "&nbsp;&nbsp;&nbsp; "&job_txt_318&": " & formatdatetime(oRec2("stdato"), 2) 
		else
		udldato = ""
		end if%>
		
		<%
		if cint(intServiceaft) = cint(oRec2("id")) then
		serChk = "SELECTED"
		else
		serChk = ""
		end if
		
		if oRec2("status") <> 0 then
		stThis = "Aktiv"
		else
		stThis = "Lukket"
		end if%>
		<option value="<%=oRec2("id")%>" <%=serChk%>> <%=oRec2("navn")%>&nbsp;(<%=oRec2("aftalenr") %>)  <%=udldato%> (<%=stThis%>) </option>
		<%
		oRec2.movenext
		wend
		oRec2.close
		%>
		</select>
		
		<%if func = "red" then %><br>
		<input type="checkbox" name="FM_overforGamleTimereg" value="1"> <%=job_txt_074 %> 
		<b><%=job_txt_075 %></b><%=" "&job_txt_076&" " %><b><%=job_txt_077 %></b><%=" "&job_txt_078 %> 
      
        
        <!--<br />
		<br />Hvis der <b>�ndres</b> aftale tilknytning undervejs i jobbets levetid, vil evt. oprettede
		fakturaer p� jobbet ikke l�ngere kunne ses under den <b>gamle aftale.</b> (faktura historik, aftale afstemning)<br />
		Fakturaer oprettet direkte p� aftalen vil ikke blive ber�rt. -->
		<%end if%>
		  <br>&nbsp;</td>
		
	</tr>
	
	
	<%
	'**** Beskrivelse *****
	if showaspopup <> "y" then
	bcols = 62
	else
	bcols = 52
	end if%>

    <%if (func = "red"  AND lto = "hestia" OR lto = "dencker" OR lto = "hidalgo") AND id <> 0 then %>

        <%
        filnavn = ""
        findesfil = 0
        strSQL = "SELECT filnavn FROM filer WHERE filertxt = '"&id&"_beskrivelsesdokument' ORDER BY id DESC"
        oRec.open strSQL, oConn, 3
        if not oRec.EOF then
            findesfil = 1
            filnavn = oRec("filnavn")
        end if
        oRec.close

        if cint(findesfil) = 1 then
            uploadtxt = "V�lg nyt"
            'response.Write "<span> strSQL "& strSQL & " " & filnavn &"</span>"
        else
            uploadtxt = "Tilf�j"
        end if
        %>

        <tr>
            <td>&nbsp</td>
            <td style="padding:5px 0px 0px 2px;" colspan=2>
                <b>Dokument:</b>                
               <a target="_blank" class="vmenu" href="../inc/upload/<%=lto %>/<%=filnavn %>"><%=filnavn %></a>
               &nbsp -  &nbsp <a style="cursor:pointer; color:#3366BB" class="vmenu" onclick="Javascript:window.open('../to_2015/upload.asp?menu=fob&func=opret&id=<%=id%>', '', 'width=650,height=600,resizable=yes,scrollbars=yes')"><b><%=uploadtxt %></b></a> 
            </td>
        </tr>
    <%end if %>

	<tr>
		<td style="">&nbsp;</td>
		<td style="padding:5px 0px 0px 2px;" colspan=2><b><%=job_txt_079 %>:</b>
		<%if showaspopup <> "y" then %>
		&nbsp;<a href="#" id="a_internnote" class="vmenu"> +<%=" "&job_txt_080 %></a><br>
		<%end if%>

		                <%
		                dim content

	                   select case lto 
                            case "synergi1", "xintranet - local"
                                if func = "red" then 
                            content = strBesk
                            else
                            content = "Deadline:<br><br>Tidsestimat:"
                            end if
                            case else
	                    content = strBesk
                            end select
            			
            			
			            
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_beskrivelse"
			            editorK.Text = content
			            editorK.FilesPath = "CuteEditor_Files"
			            
                        select case lto
                        case "dencker", "jttek", "intranet - local"
                        editorK.AutoConfigure = "Compact"
                        case else
			            editorK.AutoConfigure = "Minimal"
                        end select
            			
			            editorK.Width = 450
			            editorK.Height = 280
			            editorK.Draw()
		                %>
		                
		    <br />
            &nbsp;
            
                           
                            
		
		</td>
		<td>
		
		<%if showaspopup <> "y" then %>
		<!-- intern note -->
        <div id="div_internnote" style="position:absolute; top:420px; left:20px; visibility:hidden; padding:20px; border:10px #CCCCCC solid; background-color:#FFFFFF; width:650px; z-index:2000000;">
	    
                           <table cellpadding="0" cellspacing="0" border="0" width=100%>
							  <tr>
							  <td style="padding:15px 20px 2px 10px;"><b><%=job_txt_080 %>:</b><br />
							  <%=job_txt_081 %></td>
	                            <td align=right style="padding:0px 20px 2px 0px;"><a href="#" id="a2_internnote" class=red>[x]</a></td>
	                        </tr>
                            
                            	<tr>
		
		<td style="padding-left:5px; padding-top:6px; padding-bottom:4px;" colspan=2>
		
		                <%
		                dim content2
	                    content2 = job_internbesk
            			
			            
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_internbesk"
			            editorK.Text = content2
			            editorK.FilesPath = "CuteEditor_Files"
                        select case lto
                        case "dencker"
                        editorK.AutoConfigure = "Compact"
                        case else
			            editorK.AutoConfigure = "Minimal"
                        end select
            			
			            editorK.Width = 420
			            editorK.Height = 220
			            editorK.Draw()
		                %>
		                
		    <br />
            &nbsp;
		    <input type=checkbox value="1" name="FM_alert" <%=alertCHK %> /> <%=job_txt_082 %>
		    </td>
		   
	    </tr>
        </table>
        </div>
		<%end if%>
		
		&nbsp;</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td style="padding:0px 20px 5px 0px; width:140px;"><b><%=job_txt_083 %>:</b><br />
        (<%=job_txt_084 %>)</td>
		<td style="padding:0px 0px 5px 5px;"><input type="text" name="FM_rekvnr" id="FM_rekvnr" value="<%=rekvnr%>" style="width:260px;"></td>
		<td>&nbsp;</td>
	</tr>
	
    <%if (lto = "execon" OR lto = "immenso") then %>

	<%if cint(ski) = 1 then
	skiCHK = "CHECKED"
	else
	skiCHK = ""
	end if
	 %>
	<tr>
		<td style="" width=8>&nbsp;</td>
		<td colspan=2 style="padding:0px 20px 10px 0px;"><input id="Checkbox1" name="FM_ski" type="checkbox" value="1" <%=skiCHK %> /> <b><%=job_txt_085 %></b>&nbsp;
		(<%=job_txt_086 %>)</td>
	    <td style="" width=8>&nbsp;</td>
	</tr>
	
	
	
	<%if cint(abo) = 1 then
	aboCHK = "CHECKED"
	else
	aboCHK = ""
	end if
	 %>
	<tr>
		<td style="" width=8>&nbsp;</td>
		<td colspan=2 style="padding:0px 20px 10px 0px;"><input id="FM_abo" name="FM_abo" type="checkbox" value="1" <%=aboCHK %> /> <b><%=job_txt_087 %></b>&nbsp;(<%=job_txt_088 %>)</td>
	    <td style="" width=8>&nbsp;</td>
	</tr>
	
	
	<%if cint(ubv) = 1 then
	ubvCHK = "CHECKED"
	else
	ubvCHK = ""
	end if
	 %>
	<tr>
		<td style="" width=8>&nbsp;</td>
		<td colspan=2 style="padding:0px 20px 10px 0px;"><input id="Checkbox2" name="FM_ubv" type="checkbox" value="1" <%=ubvCHK %> /> <b><%=job_txt_089 %></b>&nbsp;
		(<%=job_txt_090 %>)</td>
	    <td style="" width=8>&nbsp;</td>
	</tr>
	
    <%end if%>
   

	  <%if lto = "dencker" then
                    '*** Dencker opret SDSK incident ****' 
                 
                 call erSDSKaktiv() 
	
	
	
	                '**** Opret Incident p� job ****'
	                if dsksOnOff <> 0 AND showaspopup <> "y" ANd multibletrue <> 1 AND sdskpriogrp <> 0 then
	
	                if lto = "xdencker" AND func <> "red" then
	                opIncCHK = "CHECKED"
	                else
	                opIncCHK = ""
	                end if%>
	                <tr bgcolor="#ffffff">
		                <td colspan=4 style="padding:10px 0px 5px 10px;">
                            <input id="FM_opr_incident" type="checkbox" name="FM_opr_incident" value="1" <%=opIncCHK %>/> <b><%=job_txt_091 %></b> (<%=job_txt_092 %>)<br />&nbsp;</td>
	                </tr>
	                <%end if 
                 
                 
        end if         %>				
					
					
					
					
                   
					
					
				
					
					
					<tr bgcolor="#FFFFFF">
						<td>&nbsp;</td>
						<td colspan=2 valign=top style="padding:10px 5px 0px 0px; white-space:nowrap;">
						<h3><%=job_txt_093 %></h3></td>
						<td>&nbsp;</td>	
				   </tr>
				   <tr bgcolor="#FFFFFF">
				        <td>&nbsp;</td>
                        <td valign=top style="padding-top:3px; padding-bottom:4px;">
                        <b><%=job_txt_241 %>:</b></td>
						<td valign=top style="padding-top:0px; padding-bottom:4px;">
						
						             <select name="FM_status" id="FM_status" style="width:160px;">
									<%
                                    lkDatoThis = ""
                                    if dbfunc = "dbred" then 
									    select case strStatus
									    case 1
									    strStatusNavn = job_txt_094
									    case 2
									    strStatusNavn = job_txt_095 'passiv
									    case 0
									    strStatusNavn = job_txt_096
                                            if cdate(lkdato) <> "01-01-2002" then
                                            lkDatoThis = " ("& formatdatetime(lkdato, 2) & ")"
                                            end if
									    case 3
									    strStatusNavn = job_txt_063
                                        case 4
									    strStatusNavn = job_txt_098
                                        case 5
									    strStatusNavn = "Evaluering"
									    end select
									%>
									<option value="<%=strStatus%>" SELECTED><%=strStatusNavn%> <%=lkDatoThis %></option>
									<%end if
                                    
                                    'call jobstatus_fn(0, 0, 1)
                                    %>
                                    <%'jobstatus_fn_options %>

                                           <%select case lto
                                            case "mpt", "local - intranet"

                                                'if level = 1 then
                                                'wprotec = ""
                                                'else
                
                                                    'if stCHK0 = "CHECKED" then 'Hvis job er lukket m� kun admin �ndre status
                                                    'wprotec = "readonly"
                                                    'else
                                                    'wprotec = ""
                                                    'end if

                                                'end if

                                                if dbfunc = "dbred" then
                                                    if level = 1 then
                                                        wprotec = ""
                                                    else
                                                        if strStatus <> 0 AND strStatus <> 2 then
                                                            wprotec = ""
                                                        else
                                                            wprotec = "readonly"
                                                        end if
                                                    end if
                                                end if

                                            case else

                                                if level <= 2 OR level = 6 then
                                                wprotec = ""
                                                else

                                                    if stCHK0 = "CHECKED" then
                                                    wprotec = "readonly"
                                                    else
                                                    wprotec = ""
                                                    end if

                                                end if

                                         end select  %> 


                                    <%if wprotec <> "readonly" then %>
									<option value="1"><%=jobstatus_txt_007 %></option>
									<option value="2"><%=jobstatus_txt_008 %></option> 
									<option value="4"><%=jobstatus_txt_004 %></option>
							        <option value="3"><%=jobstatus_txt_003 %></option>
                                     <option value="5"><%=jobstatus_txt_010 %></option>
                                    <%end if %>

                                    <%if (lto = "mpt" AND level = 1) OR lto <> "mpt" then %>
									<option value="0"><%=jobstatus_txt_009 %></option>
                                    <%end if %>

									</select> 

                                    <%if strStatus = 3 then
                                    %>
                                    <br /><span style="color:#999999;"><%=job_txt_099 %></span>
                                    <%
                                    end if %>
									
						<td>&nbsp;</td>	
				   </tr>		
			       
                   
                  
					
					
		
	
								<%if showaspopup <> "y" then%>
									<tr bgcolor="#ffffff">
										<td style="">&nbsp;</td>
										<td style="padding-top:4px;"><b><%=job_txt_100 %>:</b> <br /><span style="color:#999999; font-size:9px;"><%=job_txt_101 %></span></td>
										<td style="padding-top:4px;"><select name="FM_start_dag" id="FM_start_dag">
										<option value="<%=strDag%>"><%=strDag%></option> 
										<option value="1">1</option>
									   	<option value="2">2</option>
									   	<option value="3">3</option>
									   	<option value="4">4</option>
									   	<option value="5">5</option>
									   	<option value="6">6</option>
									   	<option value="7">7</option>
									   	<option value="8">8</option>
									   	<option value="9">9</option>
									   	<option value="10">10</option>
									   	<option value="11">11</option>
									   	<option value="12">12</option>
									   	<option value="13">13</option>
									   	<option value="14">14</option>
									   	<option value="15">15</option>
									   	<option value="16">16</option>
									   	<option value="17">17</option>
									   	<option value="18">18</option>
									   	<option value="19">19</option>
									   	<option value="20">20</option>
									   	<option value="21">21</option>
									   	<option value="22">22</option>
									   	<option value="23">23</option>
									   	<option value="24">24</option>
									   	<option value="25">25</option>
									   	<option value="26">26</option>
									   	<option value="27">27</option>
									   	<option value="28">28</option>
									   	<option value="29">29</option>
									   	<option value="30">30</option>
										<option value="31">31</option></select>&nbsp;
										
                                        <%
                                            if strMrdNavn = "jan" then strMrdNavnTranslate = job_txt_102
                                            if strMrdNavn = "feb" then strMrdNavnTranslate = job_txt_103
                                            if strMrdNavn = "mar" then strMrdNavnTranslate = job_txt_104
                                            if strMrdNavn = "apr" then strMrdNavnTranslate = job_txt_105
                                            if strMrdNavn = "maj" then strMrdNavnTranslate = job_txt_106
                                            if strMrdNavn = "jun" then strMrdNavnTranslate = job_txt_107
                                            if strMrdNavn = "jul" then strMrdNavnTranslate = job_txt_108
                                            if strMrdNavn = "aug" then strMrdNavnTranslate = job_txt_109
                                            if strMrdNavn = "sep" then strMrdNavnTranslate = job_txt_110
                                            if strMrdNavn = "okt" then strMrdNavnTranslate = job_txt_111
                                            if strMrdNavn = "nov" then strMrdNavnTranslate = job_txt_112
                                            if strMrdNavn = "dec" then strMrdNavnTranslate = job_txt_113
                                        %>

										<select name="FM_start_mrd" id="FM_start_mrd">
										<option value="<%=strMrd%>"><%=strMrdNavnTranslate%></option>
										<option value="1"><%=job_txt_102 %></option>
									   	<option value="2"><%=job_txt_103 %></option>
									   	<option value="3"><%=job_txt_104 %></option>
									   	<option value="4"><%=job_txt_105 %></option>
									   	<option value="5"><%=job_txt_106 %></option>
									   	<option value="6"><%=job_txt_107 %></option>
									   	<option value="7"><%=job_txt_108 %></option>
									   	<option value="8"><%=job_txt_109 %></option>
									   	<option value="9"><%=job_txt_110 %></option>
									   	<option value="10"><%=job_txt_111 %></option>
									   	<option value="11"><%=job_txt_112 %></option>
									   	<option value="12"><%=job_txt_113 %></option></select>
										
										
										<select name="FM_start_aar" id="FM_start_aar">
										<option value="<%=strAar%>">
										<%if id <> 0 then%>
										20<%=strAar%>
										<%else%>
										<%=strAar%>
										<%end if%></option>
										
										<%for x = -10 to 15
		                                useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		                                  <option value="<%=right(useY, 2)%>"><%=useY%></option>
		                                <%next %>
										
									  
										</select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=6')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a></td>
										
                                        <%select case lto
                                        case "epi2017", "xintranet - local"
                                        sltuDatoCol = "border:1px red solid;"
                                                
                                        case else
                                        sltuDatoCol = ""    
                                        end select %>
										
										<td style="">&nbsp;</td>
										</tr>
										<tr bgcolor="#ffffff">
											<td style="padding-top:0px; padding-left:5px;" rowspan=2 align=right>&nbsp;</td>
											<td valign=top style="padding-top:5px; padding-bottom:4px;"><b><%=job_txt_114 %>:</b>&nbsp;
											</td>
											<td style="padding-bottom:4px;"><select name="FM_slut_dag" id="FM_slut_dag" style="<%=sltuDatoCol%>">
										<option value="<%=strDag_slut%>"><%=strDag_slut%></option> 
									   	<option value="1">1</option>
									   	<option value="2">2</option>
									   	<option value="3">3</option>
									   	<option value="4">4</option>
									   	<option value="5">5</option>
									   	<option value="6">6</option>
									   	<option value="7">7</option>
									   	<option value="8">8</option>
									   	<option value="9">9</option>
									   	<option value="10">10</option>
									   	<option value="11">11</option>
									   	<option value="12">12</option>
									   	<option value="13">13</option>
									   	<option value="14">14</option>
									   	<option value="15">15</option>
									   	<option value="16">16</option>
									   	<option value="17">17</option>
									   	<option value="18">18</option>
									   	<option value="19">19</option>
									   	<option value="20">20</option>
									   	<option value="21">21</option>
									   	<option value="22">22</option>
									   	<option value="23">23</option>
									   	<option value="24">24</option>
									   	<option value="25">25</option>
									   	<option value="26">26</option>
									   	<option value="27">27</option>
									   	<option value="28">28</option>
									   	<option value="29">29</option>
									   	<option value="30">30</option>
										<option value="31">31</option></select>&nbsp;
										
                                        <%
                                            if strMrdNavn_slut = "jan" then strMrdNavnTranslate_slut = job_txt_102
                                            if strMrdNavn_slut = "feb" then strMrdNavnTranslate_slut = job_txt_103
                                            if strMrdNavn_slut = "mar" then strMrdNavnTranslate_slut = job_txt_104
                                            if strMrdNavn_slut = "apr" then strMrdNavnTranslate_slut = job_txt_105
                                            if strMrdNavn_slut = "maj" then strMrdNavnTranslate_slut = job_txt_106
                                            if strMrdNavn_slut = "jun" then strMrdNavnTranslate_slut = job_txt_107
                                            if strMrdNavn_slut = "jul" then strMrdNavnTranslate_slut = job_txt_108
                                            if strMrdNavn_slut = "aug" then strMrdNavnTranslate_slut = job_txt_109
                                            if strMrdNavn_slut = "sep" then strMrdNavnTranslate_slut = job_txt_110
                                            if strMrdNavn_slut = "okt" then strMrdNavnTranslate_slut = job_txt_111
                                            if strMrdNavn_slut = "nov" then strMrdNavnTranslate_slut = job_txt_112
                                            if strMrdNavn_slut = "dec" then strMrdNavnTranslate_slut = job_txt_113
                                        %>

										<select name="FM_slut_mrd" id="FM_slut_mrd" style="<%=sltuDatoCol%>">
										<option value="<%=strMrd_slut%>"><%=strMrdNavnTranslate_slut%></option>
										<option value="<%=strMrd%>"><%=strMrdNavn%></option>
										<option value="1"><%=job_txt_102 %></option>
									   	<option value="2"><%=job_txt_103 %></option>
									   	<option value="3"><%=job_txt_104 %></option>
									   	<option value="4"><%=job_txt_105 %></option>
									   	<option value="5"><%=job_txt_106 %></option>
									   	<option value="6"><%=job_txt_107 %></option>
									   	<option value="7"><%=job_txt_108 %></option>
									   	<option value="8"><%=job_txt_109 %></option>
									   	<option value="9"><%=job_txt_110 %></option>
									   	<option value="10"><%=job_txt_111 %></option>
									   	<option value="11"><%=job_txt_112 %></option>
									   	<option value="12"><%=job_txt_113 %></option></select>
										
										
										<select name="FM_slut_aar" id="FM_slut_aar" style="<%=sltuDatoCol%>">
										<option value="<%=strAar_slut%>">
										<%if id <> 0 then%>
										20<%=strAar_slut%>
										<%else%>
										<%=strAar_slut%>
										<%end if%></option>
										
									   
									   <%for x = -10 to 15 
		                                useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		                                <option value="<%=right(useY, 2)%>"><%=useY%></option>
		                                <%next %>
									   
									   </select>&nbsp;&nbsp;<a href="javascript:NewWin_popupcal('../inc/regular/popupcalender_inc.asp?use=5')"><img src="../ill/popupcal.gif" alt="" border="0" width="16" height="15"></a>&nbsp;&nbsp;
												
                                                <%if lto <> "epi2017" then %>
                                                
                                                <br />
												<%if strAar_slut = 44 then
												chald = "checked"
												else
												chald = ""
												end if%>
												<input type="checkbox" name="FM_datouendelig" value="j" <%=chald%>> <%=job_txt_115 %>
                                                <%end if %>
                                               </td>
										<td style="">&nbsp;</td>
									</tr>
                                     
                                     <%if lto <> "epi2017" then %>
									<tr bgcolor="#ffffff">
										
										<td colspan=2 style="padding:10px 5px 10px 0px;"><b><%=job_txt_116 %></b>
										 (<%=job_txt_117 %>)
										 <input type="text" name="FM_antaluger" id="FM_antaluger" size="4" style="font-size:9px;">
										 <input type="button" value="<%=job_txt_582 %>" onClick="beregnuger();" style="font-size:9px;">&nbsp;&nbsp;
										</td>
										<td style="">&nbsp;</td>
									</tr>
                                    <%end if %>

                                    <tr bgcolor="#ffffff"><td colspan=4 style="padding:10px 0px 0px 8px;">
                                               
                                         <%if lto <> "epi2017" then %>
                                        
                                                 <%if syncslutdato = 1 then
                                                syncslutdatoCHK = "CHECKED"
                                                else
                                                syncslutdatoCHK = ""
                                                end if %>
                                                <b><%=job_txt_118 %>:</b>
                                                 <br />	<input type="checkbox" name="FM_syncslutdato" value="j" <%=syncslutdatoCHK %>> <%=job_txt_119 %> 

                                        <%end if %>    

                                                  <br />	<input type="checkbox" name="FM_syncaktdatoer" value="j"> <%=job_txt_120 %><br />&nbsp;
                                       </td></tr>
									
	<tr bgcolor="#FFFFFF">
        <td colspan=4 style="padding:10px 0px 10px 8px;">
        <b><%=job_txt_121 %>: (WIP)</b>&nbsp;&nbsp;<input id="FM_restestimat" name="FM_restestimat" type="text" size="6" value="<%=restestimat %>" />&nbsp;
        <select id="FM_stade_tim_proc" name="FM_stade_tim_proc">
           
           <%
           sele0 = ""
           sele1 = ""
           
           select case stade_tim_proc
           case 1
           sele1 = "SELECTED"
           case else
           sele0 = "SELECTED"
           end select
            %>
           
            <option <%=sele0 %> value="0"><%=job_txt_122 %></option>
            <option <%=sele1 %> value="1">% <%=job_txt_123 %></option>
        </select><br />&nbsp;
       </td>
    </tr>    
									
									
								<%else
								'*** Start og slut dato s�ttes = med hinanden n�r der oprettes fra ressource booking 
								'*** �ndres senere til den dato der v�lges i jbpla_w_opr.asp ***
								strDag = day(request("dato"))
								strMrd = month(request("dato"))
								strAar = year(request("dato"))%>
								<input type="hidden" name="FM_start_dag" value="<%=strDag%>">
								<input type="hidden" name="FM_start_mrd" value="<%=strMrd%>">
								<input type="hidden" name="FM_start_aar" value="<%=strAar%>">
								<input type="hidden" name="FM_slut_dag" value="<%=strDag%>">
								<input type="hidden" name="FM_slut_mrd" value="<%=strMrd%>">
								<input type="hidden" name="FM_slut_aar" value="<%=strAar%>">
								<%end if%>
	
	
	
	<!-- fordel timepriser -->
	
	
	
	
	<!-- Til stam akt. grupper -->
	
	<!-- stam akt tilde SLUT --->
	
	
	                 <!-- Forkalk. tmer Nettooms�tning og brutto oms�tning ---->
	                <input type="hidden" name="FM_fakturerbart" value="1">
	
					<%if showaspopup <> "y" then%>
				   <input id="jobfaktype" name="FM_jobfaktype" value="0" type="hidden" />   
	
	
	
	
						
			
                                <%  
                            call salgsans_fn()
                             
                            if cint(showSalgsAnv) = 1 then 
                            jobansvHeaderTxt = job_txt_124
                            selWdt = "80px"
                            else
                            jobansvHeaderTxt = job_txt_125
                            selWdt = "240px"       
                            end if
                                    
                            select case lto 
                            case "epi", "epi_ab", "epi_no", "epi_sta", "intranet - local", "epi_uk"
                            mTypeExceptSQL = " AND (medarbejdertype <> 14 AND medarbejdertype <> 24)"
                            case else
                            mTypeExceptSQL = ""         
                            end select      
                            %>
			
	               
					<tr>
						<td style="">&nbsp;</td>
						<td colspan="2" style="padding-left:5px; padding-top:15px;" valign=top><h3><%=jobansvHeaderTxt %>
						
						<span style="font-size:9px; font-style:normal; font-weight:normal; line-height:12px;">- <%=job_txt_126&" " %> 
						<%=job_txt_127 %></span></h3>
						
                            <table style="width:100%;"><tr>  <td>
                           


						<table style="width:100%;">
						
						
						<%for ja = 1 to 5 
					
						select case ja
						case 1

						    'jbansImg = "<img src='../ill/ac0019-24.gif' width='24' height='24' alt='Jobansvarlig' border='0'>"
                            if cint(showSalgsAnv) = 1 then 
						    jbansTxt = job_txt_230 & ":"
                            else
                            jbansTxt = job_txt_023 & ":"
                            end if
						    jobansField = "jobans1, jobans_proc_1"


                                select case lto
                                case "intranet - local", "epi2017"
                                    fltDis = ""
                                    fltProcDis = ""
                                case else 
                                    fltDis = ""
                                    fltProcDis = ""
                                end select

						case 2
						'jbansImg = "<img src='../ill/ac0020-24.gif' width='24' height='24' alt='Jobejer' border='0'>"
						jbansTxt = job_txt_024 & ":"
						jobansField = "jobans2, jobans_proc_2"


                                select case lto
                                case "intranet - local", "epi2017"
                                    fltDis = ""
                                    
                                    if func = "red" then 'AND cDate(strTdato) < cDate("01-01-2018")
                                    fltProcDis = ""
                                    else
                                    fltProcDis = "Disabled"
                                    end if

                                case else 
                                    fltDis = ""
                                    fltProcDis = ""
                                end select

						case 3
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
                        select case lto
                        case "intranet - local", "epi2017"
						jbansTxt = job_txt_128 & ":"
                        case else 
                        jbansTxt = job_txt_128&" 1:"
                        end select
						jobansField = "jobans3, jobans_proc_3"

                                select case lto
                                case "intranet - local", "epi2017"
                                    fltDis = ""
                                    fltProcDis = ""
                                case else 
                                    fltDis = ""
                                    fltProcDis = ""
                                end select

						case 4
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						 select case lto
                        case "intranet - local", "epi2017"
                              if func = "red" AND cDate(strTdato) < cDate("02-01-2018") then
                              jbansTxt = "Co. resp. 2:"
                              else
                              jbansTxt = ""
                              end if
                        case else
						jbansTxt = job_txt_128&" 2:"
                        end select
						jobansField = "jobans4, jobans_proc_4"

                                 select case lto
                                case "intranet - local", "epi2017"
                                    
                                    if func = "red" AND cDate(strTdato) < cDate("02-01-2018") then
                                        fltDis = ""
                                        fltProcDis = ""
                                    else
                                        fltDis = "Disabled"
                                        fltProcDis = "Disabled"
                                    end if

                                case else 
                                    fltDis = ""
                                    fltProcDis = ""
                                end select

						case 5
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
                        select case lto
                        case "intranet - local", "epi2017"
                             if func = "red" AND cDate(strTdato) < cDate("02-01-2018") then
                              jbansTxt = "Co. resp. 3:"
                              else
                              jbansTxt = ""
                              end if
                        case else
						jbansTxt = job_txt_128&" 3:"
                        end select
						jobansField = "jobans5, jobans_proc_5"


                                 select case lto
                                case "intranet - local", "epi2017"
                                    if func = "red" AND cDate(strTdato) < cDate("02-01-2018") then
                                        fltDis = ""
                                        fltProcDis = ""
                                    else
                                        fltDis = "Disabled"
                                        fltProcDis = "Disabled"
                                    end if
                                case else 
                                    fltDis = ""
                                    fltProcDis = ""
                                end select

						end select
						
						%>
						<tr>
						<!--<td><%=jbansImg  %></td>-->
						<td>
						<b><%=jbansTxt %></b></td><td>
						&nbsp;&nbsp;<select name="FM_jobans_<%=ja %>" id="FM_jobans_<%=ja %>" style="width:<%=selWdt%>;" <%=fltDis %>>
						<option value="0"><%=job_txt_129 %></option>
							<%

                            select case lto
                            case "hestia", "intranet - local" 
							mPassivSQL = " OR mansat = 3"
                            case else
                            mPassivSQL = ""
                            end select

							if func <> "red" then
							strSQL = "SELECT mnavn, mnr, mid, init, mansat FROM medarbejdere WHERE ((mansat = 1 "& mPassivSQL &")" & mTypeExceptSQL & ")"
							else
							strSQL = "SELECT mnavn, mnr, mid, mansat, "& jobansField &", init FROM medarbejdere "_
							&" LEFT JOIN job ON (job.id = "& id &") WHERE ((mansat = 1 "& mPassivSQL &") "& mTypeExceptSQL &") OR (mid = jobans"& ja &")"
							end if

                            strSQL = strSQL & " ORDER BY mnavn"

							oRec.open strSQL, oConn, 3 
							while not oRec.EOF 
							
							if func <> "red" then
							 
                              select case ja 
                              case 1 
							  usemed = session("mid")
                              jobans_proc = 100
                              case 2
                                    
                                    select case lto
                                    case "intranet - local", "epi2017", "mpt"
                                    usemed = jobans2 
                                    jobans_proc = 0
                                    case else
                                    usemed = 0
                                    jobans_proc = 0
                                    end select
							  
                              case else
							  usemed = 0
                              jobans_proc = 0
							  end select

                              


							else
							    select case ja
						        case 1
						        usemed = oRec("jobans1")
                                jobans_proc = oRec("jobans_proc_1")
						        case 2
						        usemed = oRec("jobans2")
                                jobans_proc = oRec("jobans_proc_2")
						        case 3
						        usemed = oRec("jobans3")
                                jobans_proc = oRec("jobans_proc_3")
						        case 4
						        usemed = oRec("jobans4")
                                jobans_proc = oRec("jobans_proc_4")
						        case 5
						        usemed = oRec("jobans5")
                                jobans_proc = oRec("jobans_proc_5")
						        end select
							end if
							
								if cdbl(usemed) = oRec("mid") then
								medsel = "SELECTED"
								else
								medsel = ""
								end if

                                if len(trim(oRec("init"))) <> 0 then
                                opTxt = oRec("init") &" - "& oRec("mnavn")
                                else
                                opTxt = oRec("mnavn") '&" (" & oRec("mnr") &")"
                                end if

                                if oRec("mansat") <> 1 then
                                select case oRec("mansat")
                                case 2
                                opTxt = opTxt & " -" & job_txt_319
                                case 3 
                                opTxt = opTxt & " - " & job_txt_320
                                end select
                                end if

                            %>
							<option value="<%=oRec("mid")%>" <%=medsel%>><%=opTxt %></option>
							<%

							oRec.movenext
							wend
							oRec.close 
							%>
						</select>
						</td><td style="white-space:nowrap;"><input id="FM_jobans_proc_<%=ja %>" name="FM_jobans_proc_<%=ja %>" value="<%=formatnumber(jobans_proc, 1) %>" type="text" style="width:40px; <%=sltuDatoCol%>;" <%=fltProcDis %> /> %</td></tr>
                            
						<%next %>
						
						
						</table>
						

                             <%
                                
                           
                            if cint(showSalgsAnv) = 1 then %>
                            </td><td>
                           <table style="width:100%;">
						
						
						<%for sa = 1 to 5 
					
						select case sa
						case 1
						'jbansImg = "<img src='../ill/ac0019-24.gif' width='24' height='24' alt='Jobansvarlig' border='0'>"
						saansTxt = job_txt_322 & " 1:" 'job_txt_321
						salgsansField = "salgsans1, salgsans1_proc"
						case 2
						'jbansImg = "<img src='../ill/ac0020-24.gif' width='24' height='24' alt='Jobejer' border='0'>"
						
						salgsansField = "salgsans2, salgsans2_proc"

                                select case lto
                                case "xintranet - local", "xepi2017"
                                    if func = "red" AND cDate(strTdato) < cDate("02-01-2018") then
                                        fltSaDis = ""
                                        fltSaProcDis = ""
                                        saansTxt = job_txt_322 & " 2:"
                                    else
                                        fltSaDis = "Disabled"
                                        fltSaProcDis = "Disabled"
                                        saansTxt = ""
                                    end if
                                case else 
                                    fltSaDis = ""
                                    fltSaProcDis = ""
                                    saansTxt = job_txt_322 & " 2:"
                                end select

						case 3
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						 
						salgsansField = "salgsans3, salgsans3_proc"

                             select case lto
                                case "xintranet - local", "xepi2017"
                                    if func = "red" AND cDate(strTdato) < cDate("02-01-2018") then
                                        fltSaDis = ""
                                        fltSaProcDis = ""
                                        saansTxt = job_txt_322 & " 3:"
                                    else
                                        fltSaDis = "Disabled"
                                        fltSaProcDis = "Disabled"
                                        saansTxt = ""
                                    end if
                                case else 
                                    fltSaDis = ""
                                    fltSaProcDis = ""
                                    saansTxt = job_txt_322 & " 3:"
                                end select

						case 4
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
					
						salgsansField = "salgsans4, salgsans4_proc"


                             select case lto
                                case "xintranet - local", "xepi2017"
                                    if func = "red" AND cDate(strTdato) < cDate("02-01-2018") then
                                        fltSaDis = ""
                                        fltSaProcDis = ""
                                        saansTxt = job_txt_322 & " 4:"
                                    else
                                        fltSaDis = "Disabled"
                                        fltSaProcDis = "Disabled"
                                        saansTxt = ""
                                    end if
                                case else 
                                    fltSaDis = ""
                                    fltSaProcDis = ""
                                    saansTxt = job_txt_322 & " 4:"
                                end select

						case 5
						'jbansImg = "<img src='../ill/blank.gif' width='24' height='24' alt='Jobejer' border='0'>"
						
						salgsansField = "salgsans5, salgsans5_proc"


                             select case lto
                                case "xintranet - local", "xepi2017"
                                    if func = "red" AND cDate(strTdato) < cDate("02-01-2018") then
                                        fltSaDis = ""
                                        fltSaProcDis = ""
                                        saansTxt = job_txt_322 & " 5:"
                                    else
                                        fltSaDis = "Disabled"
                                        fltSaProcDis = "Disabled"
                                        saansTxt = ""
                                    end if
                                case else 
                                    fltSaDis = ""
                                    fltSaProcDis = ""
                                    saansTxt = job_txt_322 & " 5:"
                                end select

						end select
						
						%>
						<tr>
						<!--<td><%=jbansImg  %></td>-->
						<td>
						<b><%=saansTxt%></b></td><td>
						&nbsp;&nbsp;<select name="FM_salgsans_<%=sa %>" id="Select4" style="width:<%=selWdt%>;" <%=fltSaDis %>>
						<option value="0"><%=job_txt_129 %></option>
							<%
							
							if func <> "red" then
							strSQL = "SELECT mnavn, mnr, mid, init, mansat FROM medarbejdere WHERE (mansat = 1 " & mTypeExceptSQL &")"
							else
							strSQL = "SELECT mnavn, mnr, mid, mansat, "& salgsansField &", init FROM medarbejdere "_
							&" LEFT JOIN job ON (job.id = "& id &") WHERE (mansat = 1 "& mTypeExceptSQL &") OR (mid = salgsans"& sa &")"
							end if

                            strSQL = strSQL & " ORDER BY mnavn"

							
                                
                            oRec.open strSQL, oConn, 3 
							while not oRec.EOF 
							
							if func <> "red" then
							  
                              if sa = 1 then
							  usemed = session("mid")
                              salgsans_proc = 0
							  else
							  usemed = 0
                              salgsans_proc = 0
							  end if

                              

							else
							            select case sa
						                case 1
						                usemed = oRec("salgsans1")
                                        salgsans_proc = oRec("salgsans1_proc")
						                case 2
						                usemed = oRec("salgsans2")
                                        salgsans_proc = oRec("salgsans2_proc")
						                case 3
						                usemed = oRec("salgsans3")
                                        salgsans_proc = oRec("salgsans3_proc")
						                case 4
						                usemed = oRec("salgsans4")
                                        salgsans_proc = oRec("salgsans4_proc")
						                case 5
						                usemed = oRec("salgsans5")
                                        salgsans_proc = oRec("salgsans5_proc")
						                end select
							end if
							
								if cdbl(usemed) = oRec("mid") then
								medsel = "SELECTED"
								else
								medsel = ""
								end if


						        if len(trim(oRec("init"))) <> 0 then
                                opTxt = oRec("init") &" - "& oRec("mnavn")
                                else
                                opTxt = oRec("mnavn") '&" (" & oRec("mnr") &")"
                                end if

                                if oRec("mansat") <> 1 then
                                select case oRec("mansat")
                                case 2
                                opTxt = opTxt & " -" & job_txt_319
                                case 3 
                                opTxt = opTxt & " - " & job_txt_320
                                end select
                                end if

                            %>
							<option value="<%=oRec("mid")%>" <%=medsel%>><%=opTxt %> // <%=usemed %></option>
							<%
							oRec.movenext
							wend
							oRec.close 
							%>
						</select>
						</td><td style="white-space:nowrap;"><input id="FM_salgsans_proc_<%=sa %>" name="FM_salgsans_proc_<%=sa %>" value="<%=formatnumber(salgsans_proc, 1) %>" type="text" style="width:40px; <%=sltuDatoCol%>;" <%=fltSaProcDis %> /> %</td></tr>
                            
						<%next %>
						
						
						</table>
                            <%end if %>

                            </td></tr></table>







						<br>
						<%
                        
                        if level = 1 then 
                             'if (lto <> "epi" AND lto <> "epi2017") OR (lto = "epi" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59 OR thisMid = 1 OR thisMid = 1720)) OR (lto = "epi_no" AND thisMid = 2) OR (lto = "epi_ab" AND thisMid = 2) OR (lto = "epi_sta" AND thisMid = 2) OR (lto = "epi_cati" AND thisMid = 2) OR (lto = "epi_uk" AND thisMid = 2) then
                            showVAndel = 0
                            if showVAndel = 1 then
                            %>
                            
                            <input type="text" name="FM_virksomheds_proc" value="<%=formatnumber(virksomheds_proc, 0) %>" style="width:40px;" /> <%=job_txt_130 & " %" %>
                            <br />
                            <%else %>
                              <input type="hidden" name="FM_virksomheds_proc" value="<%=formatnumber(virksomheds_proc, 0) %>"/>
                            <%end if %>
                        <%else %>
                        <input type="hidden" name="FM_virksomheds_proc" value="<%=formatnumber(virksomheds_proc, 0) %>"/>
                        <%end if %>
					    
                        
                        <%if func <> "red" then

                            select case lto
                            case "synergi1", "qwert", "hestia", "xmpt"
                            advJobansCHK = "CHECKED"
                            case "dencker"

                                if session("mid") = 34 OR session("mid") = 49 then 'Aministrationen 
                                advJobansCHK = "CHECKED"
                                else
                                advJobansCHK = ""
                                end if
                            
                            case else
                            advJobansCHK = ""
                            end select
                        
                        else
                            
                            select case lto
                            case "hestia"
                            advJobansCHK = "CHECKED"
                            case else
                            advJobansCHK = ""
                            end select

                        end if %>

                        <input type="checkbox" value="1" name="FM_adviser_jobans" <%=advJobansCHK %> /><%=job_txt_131 %>

                       
						</td>
						<td style="">&nbsp;</td>
					</tr>
					
		       
		
		        <!-- Avanceret indstillinger -->
				<tr bgcolor="#FFFFFF">
								<td colspan=4 style="padding:20px 0px 0px 10px;">
                              

                            
						 <%'**** Projektgrupper *****%>
                           
                            <table cellpadding=0 cellspacing=0 border=0 width=100%>
							<tr>
								<td style="">&nbsp;</td>
								<td colspan=2 style="padding-left:5px;">
                                 <h4><%=job_txt_132 %> <span style="font-size:9px; font-style:normal; font-weight:normal; line-height:11px;">- <%=job_txt_133 %>?</span></h4>
                                
								<% 
                                uTxt = job_txt_134 _
                                &" "&job_txt_135
								uWdt = 300
								
								call infoUnisport(uWdt, uTxt) 

                              %>

								<br>&nbsp;</td>
								<td style="">&nbsp;</td>
							</tr>
							<%
							p = 0
							for p = 1 to 10
							varSelected = ""
							select case p
							case 1
							strProj = strProj_1
							case 2
							strProj = strProj_2
							case 3
							strProj = strProj_3
							case 4
							strProj = strProj_4
							case 5
							strProj = strProj_5
							case 6
							strProj = strProj_6
							case 7
							strProj = strProj_7
							case 8
							strProj = strProj_8
							case 9
							strProj = strProj_9
							case 10
							strProj = strProj_10
							end select

                                strSQLpg = "SELECT id, navn FROM projektgrupper WHERE navn IS NOT NULL AND (orgvir = 0 OR orgvir = 1) ORDER BY navn LIMIT 250"

                             

							%>
							<tr>
								<td style="">&nbsp;</td>
								<td valign="top" style="padding-left:5px; width:140px;"><%=job_txt_136 %> <%=p%>:</td>
								<td><select name="FM_projektgruppe_<%=p%>" style="width:200px;">
								<%
									
									oRec3.open strSQLpg, oConn, 3
									
									While not oRec3.EOF 
									projgId = oRec3("id")
									projgNavn = oRec3("navn")
									
                                    'if projgNavn = "Alle" then
                                    'projgNavn = "Alle-gruppen (alle medarbejdere)"
                                    'end if

									if cint(strProj) = cint(projgId) then
									varSelected = "SELECTED"
									gp = strProj
									else
									varSelected = ""
									gp = gp
									end if
									%>
									<option value="<%=projgId%>" <%=varSelected%>><%=projgNavn%></option>
									<%
									oRec3.movenext
									wend
									oRec3.close%>


						</select>
                        
                       
                        
                        
                        </td>
							<td style="">&nbsp;</td>
							</tr>
							<%
							select case p
							case 1
							gp1 = gp
							case 2
							gp2 = gp
							case 3
							gp3 = gp
							case 4
							gp4 = gp
							case 5
							gp5 = gp
							case 6
							gp6 = gp
							case 7
							gp7 = gp
							case 8
							gp8 = gp
							case 9
							gp9 = gp
							case 10
							gp10 = gp
							end select
							
							next
				
				


                %>
               <tr><td colspan=4 style="padding:20px 60px 40px 10px;">

                            
                     <%if func = "red" then
                                    
                                    if lto = "execon" OR lto = "immenso" OR lto = "synergi1" OR lto = "bf" then
                                    syncAktProjGrpCHK = "CHECKED"
                                    else
                                    syncAktProjGrpCHK = ""
                                    end if
                                
                                %>
                                
                                
                               
								<input type="checkbox" name="FM_opdaterprojektgrupper" id="FM_opdaterprojektgrupper" value="1" <%=syncAktProjGrpCHK %>> <b><%=job_txt_137 %></b><%=" "&job_txt_138 %> 
								<br />
                                <%end if%>
                
                  <%
                                if lto = "xxxx" then ' <> execon 20161013 fjernet%>
                                <br>
								<input type="checkbox" name="FM_gemsomdefault" id="FM_gemsomdefault" value="1"><b><%=job_txt_139 %></b><%=" "&job_txt_140 %>
								<span style="color:#999999;"><%=job_txt_141 %></span>
                                <%end if %>
								
                               
                              

                                <%if func <> "red" then
                                 
                                    if lto = "jm" OR lto = "synergi1" OR lto = "micmatic" OR lto = "krj" OR lto = "hestia" then 'OR lto = "lyng" OR lto = "glad" then
                                    forvalgCHK = "CHECKED"
                                    else
                                    forvalgCHK = ""
                                    end if
                                 
                                 else
                                 
                                      if lto = "hestia" then
                                      forvalgCHK = "CHECKED"
                                      else
                                      forvalgCHK = ""
                                      end if

                                 end if 
                                    
                                 call positiv_aktivering_akt_fn() 'wwf
                                 if cint(positiv_aktivering_akt_val) <> 1 then%>

                                 <br />
                              
								<input type="checkbox" name="FM_forvalgt" id="FM_forvalgt" <%=forvalgCHK %> value="1"><b><%=job_txt_142 %> / Favorites</b> 
                                <%=job_txt_143 %> 

                                <%end if %>
                              
                              
                           
                                </td></tr>
                       </table>
                    



                </td></tr>
                
                <!-- forretningsomr�der --->
                <%
                
                
                if func = "red" then
                                    
                    strFomr_rel = ""
                    strFomr_navn = ""
                    
                    strSQLfrel = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
                    &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_jobid = "& id & " AND for_aktid = 0 GROUP BY for_fomr"

                    'Response.Write strSQLfrel
                    'Response.flush
                    f = 0
                    oRec.open strSQLfrel, oConn, 3
                    while not oRec.EOF

                    if f = 0 then
                    strFomr_navn = ""
                    end if

                    strFomr_rel = strFomr_rel & "#"& oRec("for_fomr") &"#"
                    strFomr_navn = strFomr_navn & oRec("navn") & ", " 

                    f = f + 1
                    oRec.movenext
                    wend
                    oRec.close

                    if f <> 0 then
                    len_strFomr_navn = len(strFomr_navn)
                    left_strFomr_navn = left(strFomr_navn, len_strFomr_navn - 2)
                    strFomr_navn = left_strFomr_navn & ")"

                        if len(strFomr_navn) > 50 then
                        strFomr_navn = left(strFomr_navn, 50) & "..)"
                        end if
                    end if

                else

                '**** forvalgt forretningsomr�der ****' 
                select case lto 
                case "hestia", "xintranet - local"
                    
                        
                            strSQLfrel = "SELECT id, navn FROM fomr WHERE id = 2"    
                           
                        
                            f = 0
                            oRec.open strSQLfrel, oConn, 3
                            while not oRec.EOF

                            'if f = 0 then
                            'strFomr_navn = " ("
                            'end if

                            strFomr_rel = strFomr_rel & "#"& oRec("id") &"#"
                            strFomr_navn = strFomr_navn & oRec("navn") & ", " 

                            f = f + 1
                            oRec.movenext
                            wend
                            oRec.close     

                case else
                     strFomr_navn = ""
                     strFomr_rel = ""
                end select


                end if
                
                if lto <> "execon" AND lto <> "xintranet - local" then %>

                <tr bgcolor="#FFFFFF">
								<td colspan=4 style="padding:10px 0px 0px 10px;">
                                <a href="#" class=vmenu id="a_tild_forr">+ <%=job_txt_144 %> </a><br />
                                <%=job_txt_145 %> 
                                  <br /><br />
                                  <%=job_txt_146 %>  
                                  <br /><br />
                              
                                                      
                                <%
                                 ' uTxt = "Forretningsomr�der bruges bl.a. til at se tidsforbrug p� tv�rs af kunder og job, og til at se hvilke slags opgaver man bruger sin tid p�."
                                 ' uWdt = 300
								
								'call infoUnisport(uWdt, uTxt) 
                                
                                    
                                    call jobopr_mandatory_fn()

                                    div_tild_forr_Pos = "relative"
                                    div_tild_forr_Lft = "0px"
                                    div_tild_forr_Top = "0px"
                                    div_tild_forr_z = "0"
                                    fomr_div_wdt = 390

                                    if cint(fomr_mandatoryOn) = 1 then
                                    div_tild_forr_VZB = "visible"
                                    div_tild_forr_DSP = ""

                                        if func = "opret" AND step = "2" then
                                        div_tild_forr_bdr = "10"
                                        div_tild_forr_pd = "20"
                                        else
                                        div_tild_forr_bdr = "0"
                                        div_tild_forr_pd = "20"
                                        end if

                                        
                                        if func = "opret" AND step = "2" then

                                            select case lto
                                            case "xdencker", "xintranet - local"
                                            div_tild_forr_Pos = "absolute"
                                            div_tild_forr_Lft = "20px"
                                            div_tild_forr_Top = "80px"
                                            div_tild_forr_z = "400000"
                                            fomr_div_wdt = 390 
                                            case "epi2017", "dencker", "intranet - local"
                                            div_tild_forr_Pos = "absolute"
                                            div_tild_forr_Lft = "30px"
                                            div_tild_forr_Top = "-30px"
                                            div_tild_forr_z = "400000"
                                            fomr_div_wdt = 500
                                            case else
                                            div_tild_forr_Pos = "absolute"
                                            div_tild_forr_Lft = "250px"
                                            div_tild_forr_Top = "250px"
                                            div_tild_forr_z = "400000"
                                            fomr_div_wdt = 390
                                            end select

                                        end if

                                    else
                                    div_tild_forr_VZB = "hidden"
                                    div_tild_forr_DSP = "none"

                                        div_tild_forr_bdr = "0"
                                        div_tild_forr_pd = "0"
                                        
                                    end if
                                    %>

                                

						
                           
                            <div id="div_tild_forr" style="position:<%=div_tild_forr_Pos%>; left:<%=div_tild_forr_Lft%>; top:<%=div_tild_forr_Top%>; z-index:<%=div_tild_forr_z%>; width:<%=fomr_div_wdt%>px; visibility:<%=div_tild_forr_VZB%>; display:<%=div_tild_forr_DSP%>; border:<%=div_tild_forr_bdr%>px #CCCCCC solid;  padding:<%=div_tild_forr_pd%>px; background-color:#FFFFFF;">
                              <!--Forretningsomr�der er baseret p� de forretningsomr�der der er tildelt p� aktiviteteterne: <br /><b><%=strFomr_navn%></b>-->
                            
                            <table cellpadding=0 cellspacing=0 border=0 width=100%>
                                <% if cint(fomr_mandatoryOn) = 1 AND func = "opret" AND step = "2" then%>
                                <tr><td align="right"> <a href="#" class=vmenu id="a_tild_forr2" style="color:red;">[X]</a></td></tr>
                                <%end if %>
                            <tr><td valign=top>
                            <h4><%=job_txt_144 %>: <span style="font-size:11px; font-weight:lighter;">(<%=job_txt_147 %>)</span></h4>  
                                
                               
                                                            
                                <%

                                '** Finder kundetype, til forvalgte forrretningsomr�der '***
                                thisKtype_segment = 0
                                if func = "opret" AND step = 2 AND strKundeId <> "" then 


                                    strSQLktyp = "SELECT ktype FROM kunder WHERE kid = " & strKundeId
                                    oRec5.open strSQLktyp, oConn, 3
                                    if not oRec5.EOF then

                                    thisKtype_segment = oRec5("ktype")

                                    end if
                                    oRec5.close


                                end if

                                
                                
                                'strSQLf = "SELECT id, navn FROM fomr WHERE id <> 0 ORDER BY navn"
                                    strSQLf = "SELECT f.navn AS fnavn, f.id, f.konto, kp.kontonr AS kkontonr, kp.navn AS kontonavn, fomr_segment FROM fomr AS f "_
                                    &" LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) WHERE f.id <> 0 AND f.jobok = 1 ORDER BY f.navn"


                                    'if session("mid") = 1 then
                                    'response.write strSQLf
                                    'response.flush
                                    'end if

                                    %>
                                <select name="FM_fomr" id="FM_fomr" multiple="multiple" size="12" style="width:380px;">
                                <option value="0"><%=job_txt_148 %></option>
                                    
                                    <%
                                    fa = 0
                                    strchkbox = ""
                                    oRec.open strSQLf, oConn, 3
                                    while not oRec.EOF

                                        if func = "opret" AND step = 2 then '*** Opret (forvalgt)

                                        if instr(oRec("fomr_segment"), "#"& thisKtype_segment &"#") <> 0 then
                                        fSel = "SELECTED"
                                        else
                                        fSel = ""
                                        end if


                                        else '** Rediger Forretningsomr�der
                                    
                                        if instr(strFomr_rel, "#"&oRec("id")&"#") <> 0 then
                                        fSel = "SELECTED"
                                        else
                                        fSel = ""
                                        end if

                                        end if



                                    if oRec("konto") <> 0 then
                                    kontonrVal = " ("& left(oRec("kontonavn"), 10) &" "& oRec("kkontonr") &")"

                                    if cint(fomr_konto) = cint(oRec("id")) then
                                    fokontoCHK = "CHECKED"
                                    else
                                    fokontoCHK = ""
                                    end if

                                    strchkbox = strchkbox & "<input type='radio' class='FM_fomr_konto' id='FM_fomr_konto_"& oRec("id") &"' name='FM_fomr_konto' value="& oRec("id") &" "& fokontoCHK &"> " & left(oRec("fnavn"), 20) &" "& kontonrVal &"<br>"
                                    else
                                    kontonrVal = ""
                                    end if 
                                    
                                    %>
                                    <option value="<%=oRec("id")%>" <%=fSel %>><%=oRec("fnavn") &" "& kontonrVal %></option>
                                    <%
                                        fa = fa + 1

                                    
                                    oRec.movenext
                                    wend
                                    oRec.close
                                    %>
                                </select>
                                <br />
                                <br />
                                <b><%=job_txt_149 %></b><br />
                                <%=job_txt_150 %>:<br />  
                                <%=strchkbox %>

                               
                                
                                <%if func <> "red" then

                                    select case lto
                                    case "hestia", "intranet - local"
                                    fomr_sync_CHK = "CHECKED"
                                    case else
                                    fomr_sync_CHK = ""
                                    end select

                                else
                                fomr_sync_CHK = ""
                                end if

                                %>

                                <!--

                                <input id="FM_fomr_syncakt" type="checkbox" value="1" name="FM_fomr_syncakt" <%=fomr_sync_CHK%> /> Tildel / Synkroniser disse forretningsomr�der p� underliggende aktiviteter.<br />
                                Alle valgte forretningsomr�der bliver fordelt ligeligt p� hver aktivitet.
                                -->

                                <!--
                                <span style="font-size:11px; color:#999999;"><br />
                                Forretningsomr�der <b>skal</b> fordeles ned p� aktiviteter for at blive talt med i statitikken over timeforbrug p� hvert forretningsomr�de. Der kan manuelt tildeles forretningsomr�der nede p� hver enkelt aktivitet.
                                </span>
                                -->
                                </td>
                                </tr></table>
                            </div>
                                    
                </td>
                </tr>
                <%else 
                
                fomrIDs = replace("0" & strFomr_rel, "#", ",")
                fomrIDs = replace(fomrIDs, ",,", ",")
                fomrIDs = left(fomrIDs, len(fomrIDs)-1)
                %>
                <input id="Hidden2" type="hidden" name="FM_fomr" value="<%=fomrIDs%>" />



                <%end if%>

                 <!-- Avanceret indstillinger 2 -->
				<tr bgcolor="#FFFFFF">
								<td colspan=4 style="padding:0px 0px 0px 10px;">
                                <a href="#" class=vmenu id="a_tild_ava">+ <%=job_txt_151 %> </a><br />
                                <%=job_txt_152 %>

								<br /><br />

                    
                      <div id="div_tild_ava" style="visibility:hidden; display:none; background-color:#FFFFFF;">
                     <table cellpadding=0 cellspacing=0 border=0 width=100%>
                    


                     <tr bgcolor="#FFFFFF">
				        <td>&nbsp;</td>
                        <td valign=top style="padding-top:30px;" colspan=3>
                        <h3><%=job_txt_153 %>:</h3></td>
                        <td style="">&nbsp;</td>
                    </tr>
                    <tr>
                        <td style="">&nbsp;</td>
						<td valign=top colspan=3>		
									
									
								
									
											<%strSQLr = "SELECT risiko FROM job WHERE risiko > -1 AND jobstatus = 1 ORDER BY risiko DESC"
									 
									 highestRiskval = 1
									 oRec4.open strSQLr, oConn, 3
									 if not oRec4.EOF then
									 highestRiskval = oRec4("risiko")
									 end if
									 oRec4.close
									 
									 %>
									 
									 <%strSQLr = "SELECT risiko FROM job WHERE risiko > -1 AND jobstatus  = 1 ORDER BY risiko"
									 
									 lowestRiskval = 0
									 oRec4.open strSQLr, oConn, 3
									 if not oRec4.EOF then
									 lowestRiskval = oRec4("risiko")
									 end if
									 oRec4.close
									 
									 %>

                                     <input id="prio" name="prio" type="text" value="<%=intprio %>" style="width:40px;" />
                                     <%=job_txt_154 %>: <b><%=lowestRiskval%> - <%=highestRiskval%></b><br /><br />
									 <b>-1 =<%=" "&job_txt_155 %></b><%=" "&job_txt_156 %><br />
                                     <b>-2 =<%=" "&job_txt_157 %></b><%=" "&job_txt_158 %><br />
                                     <b>-3 =<%=" "&job_txt_159 %></b><%=" "&job_txt_160 %> 
									 <br /><br />
                                     -1 / -2 / -3 <%=job_txt_161 %> <br /><br />&nbsp;
									
								
							</td>
							<td style="">&nbsp;</td>
						</tr>


                        
                        <%
                        preconditions_met_SEL0 = "SELECTED"
                        preconditions_met_SEL1 = ""
                        preconditions_met_SEL2 = ""

                        select case cint(preconditions_met)
                        case 0
                        preconditions_met_SEL0 = "SELECTED"
                        case 1
                        preconditions_met_SEL1 = "SELECTED"
                        case 2
                        preconditions_met_SEL2 = "SELECTED"
                        case else
                        preconditions_met_SEL0 = "SELECTED"
                        end select
                        %>


                           <tr bgcolor="#FFFFFF">
				        <td>&nbsp;</td>
                        <td valign=top style="padding-top:30px;" colspan=3>
                        <h3><%=job_txt_162 %>:<br /><span style="font-size:11px; font-weight:normal;">
                        <%=job_txt_163 %></span></h3></td>
                        <td style="">&nbsp;</td>
                    </tr>
                    <tr>
                        <td style="">&nbsp;</td>
						<td valign=top colspan=3>	
                        
                        <select name="FM_preconditions_met" id="Select1" size="1" style="width:200px;">
                        <option value="0" <%=preconditions_met_SEL0 %>><%=job_txt_164 %></option>
                        <option value="1" <%=preconditions_met_SEL1 %>><%=job_txt_165 %></option>
                        <option value="2" <%=preconditions_met_SEL2 %>><%=job_txt_166 %></option>
                        </select>
									
					</td>
							<td style="">&nbsp;</td>
						</tr>

                        <%
					
					
					
					'** Altid �ben, indstilling v�lges p� timereg. siden **'
					'cint(lukafm) = 1
                    if cint(lukafm) = 1000 then%>
					<tr bgcolor="#FFFFFF">
						<td style="">&nbsp;</td>
						<td style="padding-left:5px; padding-top:30px; padding-bottom:4px;" colspan=3><h3><%=job_txt_167 %></h3>	
						<select name="FM_lukafmjob" style="font-size:9px; width:300px;">
									<%

                                    lukafmjob0CHK = "" 
									lukafmjob2CHK = ""
									lukafmjob3CHK = ""
                                    lukafmjob1CHK = ""

									select case lukafmjob
									case 1
									lukafmjob1CHK = "SELECTED"
									case 2
									lukafmjob2CHK = "SELECTED"
									case 3
									lukafmjob3CHK = "SELECTED"
									case 0
									lukafmjob0CHK = "SELECTED" 
									end select
									%>
									
									<option value="1" <%=lukafmjob1CHK %>>A - <%=job_txt_168 %></option>
									<option value="2" <%=lukafmjob2CHK %>>B - <%=job_txt_169 %></option>
                                    <option value="3" <%=lukafmjob1CHK %>>C - <%=job_txt_170 %></option>
									<option value="0" <%=lukafmjob0CHK %>><%=job_txt_171 %></option>
								
									</select>
									
							</td>
							<td style="">&nbsp;</td>
						</tr>
					    <%else %>
					    <input type="hidden" name="FM_lukafmjob" value="<%=lukafmjob %>">
					    <%end if %>
					    
					
					<%else%>
					<input type="hidden" name="FM_fastpris" value="0">
					<input type="hidden" name="FM_status" value="1">
					<%end if%>	
                
                  
                 


                   <tr>
						<td>&nbsp;</td>
						<td colspan=3 style="padding:30px 5px 10px 0px;"><h4><%=job_txt_172 %>:<br /><span style="font-size:11px; font-weight:lighter;">(<%=job_txt_173 %>)</span></h4>
					 
                            <%call multible_licensindehavereOn() 
                             if cint(multible_licensindehavere) = 1 then
                                multiKSQL = " useasfak = 1 "
                                selMultible = "multiple"
                                multipleSz = 3
                             else
                                multiKSQL = " useasfak = 1"
                                selMultible = ""
                                multipleSz = 1
                             end if
                            lincensindehaver_faknr_prioritet = "0"
                            %>
                            Faktureres af f�lgende licensindehaver (juridisk enhed):<br /><select name="FM_lincensindehaver_faknr_prioritet_job" <%=selMultible %> size="<%=multipleSz %>" style="width:380px;">
							<%strSQL = "SELECT kid, kkundenavn, kkundenr, lincensindehaver_faknr_prioritet FROM kunder WHERE "& multiKSQL &" ORDER BY kkundenavn" 
							oRec.open strSQL, oConn, 3
							while not oRec.EOF 


                             lincensindehaver_faknr_prioritet = ""& oRec("lincensindehaver_faknr_prioritet") &""

							 if instr(lincensindehaver_faknr_prioritet_job, lincensindehaver_faknr_prioritet) <> 0 then 
                             '= lincensindehaver_faknr_prioritet_job then 'V�R OPM�RKSOM HER EPI2017, DENCKER, hvis det bliver multiple
							 kSEL = "SELECTED"
							 else
							 kSEL = ""
							 end if
                            
                            %>
							<option value="#<%=oRec("lincensindehaver_faknr_prioritet")%>#" <%=kSEL %>><%=oRec("kkundenavn") &" ("& oRec("kkundenr")&")" %></option>
							<%
							oRec.movenext
							wend
							oRec.close%>
							
							</select>
                            

                            <br /><br />
					  
							<%=job_txt_296 %>:<br /><select name="FM_valuta" style="width:380px;">
							<%strSQL = "SELECT id, valuta, valutakode, kurs FROM valutaer WHERE id <> 0 ORDER BY id " 
							oRec.open strSQL, oConn, 3
							while not oRec.EOF 
							 if oRec("id") = cint(valuta) then
							 vSEL = "SELECTED"
							 else
							 vSEL = ""
							 end if%>
							<option value="<%=oRec("id") %>" <%=vSEL %>><%=oRec("valuta") %> | <%=oRec("valutakode") %> | <%=job_txt_174 %>: <%=oRec("kurs")%></option>
							<%
							oRec.movenext
							wend
							oRec.close%>
							
							</select>

                            <br /><br />
                            <%=job_txt_175 %>: <br />

                           

                        <%strSQLmoms = "SELECT id, moms FROM fak_moms WHERE id <> 0 ORDER BY id " %>
                        <select name="FM_jfak_moms">

                            <%oRec6.open strSQLmoms, oConn, 3
                            while not oRec6.EOF 

                                if cint(jfak_moms) = cint(oRec6("id")) then
                                fakmomsSeL = "SELECTED"
                                else
                                fakmomsSeL = ""
                                end if

                            %><option value="<%=oRec6("id") %>" <%=fakmomsSeL %>><%=oRec6("moms") %>%</option><%
                  
                            oRec6.movenext
                            wend 
                            oRec6.close%>

                        </select>
        
                <br /><br />
        
                <%=job_txt_176 %>: <br />
                      <%strSQLsprog = "SELECT id, navn FROM fak_sprog WHERE id <> 0 ORDER BY id " %>
                    <select name="FM_jfak_sprog">

                        <%oRec6.open strSQLsprog, oConn, 3
                        while not oRec6.EOF 

                            if cint(jfak_sprog) = cint(oRec6("id")) then
                            faksprogSeL = "SELECTED"
                            else
                            faksprogSeL = ""
                            end if

                        %><option value="<%=oRec6("id") %>" <%=faksprogSeL %>><%=oRec6("navn") %></option><%
                  
                        oRec6.movenext
                        wend 
                        oRec6.close%>

                    </select>

							
						</td><td style="">&nbsp;</td></tr>
                
                
                
                <%if func = "opret" AND step = 2 OR func = "red" then %>
  
	          
        
                    
	              
                    <tr>
		            <td style="padding:30px 10px 10px 0px;" colspan=4>
                    <h3><%=job_txt_177 %>?</h3>
                    <input type="checkbox" name="FM_kundese" id="FM_kundese" value="1" <%=kundechk%>>&nbsp;<b><%=job_txt_178 %></b><br>
		            <%=job_txt_179 %>
		
		
		            <%if func = "opret" then
		            hvchk1 = "checked"
		            hvchk2 = ""
		            else
			            if cint(intkundeok) = 2 then
			            hvchk1 = ""
			            hvchk2 = "checked"
			            else
			            hvchk1 = "checked"
			            hvchk2 = ""
			            end if
		            end if%>
		            <br><br>
		            <b><%=job_txt_180 %>?</b><br>
		            <input type="radio" name="FM_kundese_hv" value="0" <%=hvchk1%>><%=job_txt_181 %><br>
		            <input type="radio" name="FM_kundese_hv" value="1" <%=hvchk2%>><%=job_txt_182 %>
		            <br>&nbsp;
                    </td></tr>
                  <%end if%>
                
                
                
                 
                
                   </table>
                
                </div>
                <br /><br /><br /><br /><br />&nbsp;

                </td>
                </tr>

                <%




				
				

end if '** Rettigheder	
end select '*** Step %>


	
	
<%if editok = 1 then '*** rediger rettigheder ok%>	
	
	
	
	
	<%if func = "red" then
	submtxt = job_txt_183&" & "&job_txt_184
	else
        if step = "2" then
	    submtxt = job_txt_185&" & "&job_txt_184
        else
        submtxt = job_txt_185
        end if
	end if %>
	
	<tr bgcolor="#ffffff">
		<td colspan=4 align="right" style="padding:40px 40px 10px 10px;">
		<input id="Submit2" type="submit" value=" <%=submtxt %>  >> " <%=sendmailtokunde%> /></td>
	</tr>
	
	
	</table>
   
    </td>
	</tr>
	</table>
	</div><!-- tablediv -->
	
	

	<br><br><br>&nbsp;
	
	                     
	                    
	            <%
                '******************************************************************************
                '*** div projekt beregner ****'
                '******************************************************************************** 
                
                'call browsertype()

                'if browstype_client <> "ch" then

                
                call projektberegner

               

                'else
                
                'uWdt = 300
                'uLeft = 700
                'uTop = 100
                'uTxt = "Projektberegneren er pt. ikke underst�ttet af Chrome. Vi arbejder p� en l�sning, men indtil da henvises du til at benytte Internet Explorer, Firefox eller Safari."
                'call infoUnisportAB(uWdt, uTxt, uTop, uLeft)

                'end if    
                   
                        
						
	
	
	
	
	
	'*** Funktioner og Mini joboverblik ***'
	'if request("int") <> 0 AND showaspopup <> "y" then
	mini = 1
	if func = "red" AND showaspopup <> "y" AND mini = 1 then


     if len(trim(request("visrealtimerdetal"))) <> 0 AND request("visrealtimerdetal") <> 0 then
     visrealtimerdetal = 1
     visrealtimerdetalCHK = "CHECKED"
    else
     visrealtimerdetal = 0
     visrealtimerdetalCHK = ""
    end if


    call minioverblik

       
        
    
        '********************************************************'
        '**************** Milep�le ******************************'
        '********************************************************'

    oTop = 0
	oLeft = 487
	oWdth = 700
	oHgt = "" 
	oId = "milep_div"
	oZindex = 1000
    oVzb = "hidden"
    oDsp = "none"

	call tableDivAbs(oTop,oLeft,oWdth,oHgt,oId, oVzb, oDsp, oZindex)
    %>

    <table cellspacing="0" cellpadding="3" border="0" bgcolor="#EFF3FF" width="100%">
	<tr>
	<tr bgcolor="#5582D2">
		<td width="8" valign=top><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
		<td colspan=4 style="padding-top:3px;"><h3 class="hv"><%=job_txt_186 %></h3></td>
		<td align=right valign=top><img src="../ill/blank.gif" width="8" height="30" alt="" border="0"></td>
	</tr>




	
	    

        <%for m = 0 to 1 %>
        

        <%if m = 0 then %>
        <tr bgcolor="#FFFFFF">
		<td valign="top" height=10>&nbsp;</td>
		<td colspan=3><br /><h4><%=job_txt_187 %></h4></td>
        <td align=right>
        	<a href="javascript:popUp('milepale.asp?menu=job&func=opr&jid=<%=id%>&rdir=redjobcontionue&type=0','650','500','250','120');" target="_self"><span style="color:green; font-size:16px;"><b>+</b></span> <%=job_txt_188 %></a>
        </td>
		<td valign="top" align=right>&nbsp;</td>
	    </tr>
		<tr>
			<td valign="top" height=25 style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;">&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"">
			<b><%=formatdatetime(strTdato, 1)%></b></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;""><%=job_txt_189 %></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"" colspan=2>&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"" align=right>&nbsp;</td>
		</tr>
        <%else %>
         <tr bgcolor="#FFFFFF">
		<td valign="top" height=10>&nbsp;</td>
		<td colspan=3><br /><h4><%=job_txt_190 %></h4></td>
        <td align=right>
        	<a href="javascript:popUp('milepale.asp?menu=job&func=opr&jid=<%=id%>&rdir=redjobcontionue&type=1','650','500','250','120');" target="_self"><span style="color:green; font-size:16px;"><b>+</b></span> <%=job_txt_191 %></a>
        </td>
		<td valign="top" align=right>&nbsp;</td>
	    </tr>
        <%end if %>
		
		<%
        if m = 0 then
        mtp_typ = " AND mpt_fak <> 1"
        else
        mtp_typ = " AND mpt_fak = 1"
        end if

		strSQL = "SELECT milepale.id AS id, milepale.navn AS navn, milepal_dato, "_
		&" milepale_typer.navn AS type, ikon, beskrivelse, milepale.editor, belob FROM milepale "_
		&" LEFT JOIN milepale_typer ON (milepale_typer.id = milepale.type) "_
		&" WHERE jid = "& id &" "& mtp_typ &" ORDER BY milepal_dato"
		x = 0
		oRec.open strSQL, oConn, 3
		while not oRec.EOF 
		
		select case right(x, 1)
		case 0,2,4,6,8
		bgthis = "#FFFFFF"
		case else
		bgthis = "#EFF3FF"
		end select
		%>
		<tr bgcolor="<%=bgthis %>">
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;">&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;">
            <%if isNull(oRec("milepal_dato")) = false then %>
            <b><%=formatdatetime(oRec("milepal_dato"), 1)%></b>
            <%else %>
            - Ikke angivet
            <%end if %>
            <br />
			<span style="color:#999999;"><i><%=oRec("editor") %></i></span></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;">
			<span style="color:#999999;"><%=oRec("type")%></span>
            <%if m = 1 then %><br />
            <b><%=formatnumber(oRec("belob")) &" "& basisValISO%> </b>
            <%end if %></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;">
			<a href="javascript:popUp('milepale.asp?func=red&id=<%=oRec("id")%>&jid=<%=id%>&rdir=redjobcontionue','650','500','250','120');" target="_self" class=vmenu><img src="../ill/ac0038-16.gif" alt="Rediger" border="0"><%=oRec("navn")%></a>
			<%if len(oRec("beskrivelse")) <> 0 then%>
			<br><i><%=oRec("beskrivelse")%></i>
			<%end if%></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"><a href="javascript:popUp('milepale.asp?milepale.asp?menu=job&id=<%=oRec("id")%>&jid=<%=id%>&func=slet&rdir=redjobcontionue','650','500','250','120');" target="_self"><img src="../ill/delete2.gif" alt="Slet" border="0"></a></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;" align="right">&nbsp;</td>
		</tr>
		
		<%
		x = x + 1
		oRec.movenext
		wend
		oRec.close

        select case right(x, 1)
		case 0,2,4,6,8
		bgthis = "#FFFFFF"
		case else
		bgthis = "#EFF3FF"
		end select
		%>
		

         <%if m = 0 then %>
		<tr bgcolor="<%=bgthis %>">
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"><img src="../ill/blank.gif" width="1" height="25" alt="" border="0"></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;" height=25><b><%=formatdatetime(strUdato, 1)%></b></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;"><%=job_txt_192 %></td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;" colspan=2>&nbsp;</td>
			<td valign="top" style="border-bottom:1px #CCCCCC solid; padding:4px 4px 2px 4px;" align=right><img src="../ill/blank.gif" width="1" height="25" alt="" border="0"></td>
		</tr>
        <%end if %>
		

        <%next %>

	</table>
    </div> 
		
		
		<%end if
		'************************* Joboverblik / funktioner ***'
		%>
	
	
	
    <!--
	
	
	</div>side div

    -->
	
	<%
				
	
	
	
	
	
	
				'**********************************************************
				'**** Nedenst�ende (Timepriser) vises kun ved rediger *****
                '**** Og ved klik p� timepriser ***'

                if cdbl(antalmedarb) < 25 then
                tpKriShow = 1
                else
                    if showdiv = "tpriser" then
                    tpKriShow = 1
                    else
                    tpKriShow = 0
                    end if
                end if

				if step = "2" AND func = "red" AND  tpKriShow = 1 then 'AND strInternt <> 0 '2

                call timepriser%>
                
                <%    
				end if%>	

             
	
	
	<%else%>
	<tr>
		<td colspan=4 bgcolor="#ffffff" height="100" style="border:1px #003399 solid; padding:20px;">
		<img src="../ill/alert_lille.gif" width="22" height="19" alt="" border="0">&nbsp;<font class=error><b>Fejl!</b></font><br>
		<%=job_txt_193&" " %><b><%=job_txt_194 %></b><%=" "&job_txt_195 %> 
		<%=job_txt_196 %><br><br>
		<%=job_txt_197&" " %><b><%=job_txt_198 %></b> <%=job_txt_199 %>
		<br>
		<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
		&nbsp;</td>
	</tr>
	</table>
	</div><!-- tablediv -->
	
	
	<%end if '** rettigheder%>
	
	
	</form>
	
	<br><br><br>&nbsp;
	</div><!--- side div -->

	
	
	<%'** AUTO ORPET ***
        
    if (lto = "epi2017" OR lto = "mpt") AND func <> "red" AND step = 2 then 

  
   
   
    '** AUTO OPRET job_jav_2017.js ***'
    '** EPI f�rst autoopret efter der er valgt FOMR
    select case lto
    case "mpt"


        %>
    <div id="load" style="position:absolute; display:none; visibility:hidden; top:245px; left:200px; width:860px; height:400px; background-color:#ffffff; border:10px yellowgreen solid; padding:10px; z-index:100000000;">
    <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<!--<img src="../ill/outzource_logo_200.gif" /><br /><br />-->
	<h4>TimeOut prepares your new project</h4><br /><br />
   
        
    TimeOut ties it all together and making sure you will get alle the activities you need to get a correct timerecording.<br /><br />
    It should take no more than 5 seconds...

   

        <br /><br /><br />&nbsp;

     

	</td></tr></table>
	
	</div>

    <%
    Response.Write("<script>autoopret();</script>") 
    'response.end
    end select

    end if %>


	<%case else
        
        
        
    'response.write "|"& request("FM_fomr") &"|"
    'response.write "strFomr_rel: "& strFomr_rel
        %>
	
	
	
	
	
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <script src="inc/job_listen_jav2.js"></script>
    
 <div id="loadbar" style="position:absolute; display:; visibility:visible; top:300px; left:280px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	<%=job_txt_200 %> <br />(<%=job_txt_201 %>)
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	
	</td></tr></table>

	</div>


	

    <%call menu_2014() %>
	
	
	

	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:102px;">
	
	
	<%


	
	
	
	
	Public left_intX
	Public function kommaFunc(x)
	if len(x) <> 0 then
	instr_komma = instr(x, ",")
		
		if instr_komma > 0 then
		left_intX = left(x, instr_komma + 2)
		else
		left_intX = x
		end if
	else
	left_intX = 0
	end if
	
	Response.write left_intX 
	end function
	
    if len(trim(request("filt"))) <> 0 then
	filt = request("filt")
    else
        if request.cookies("tsa")("statusfilt") <> "" then
        filt = request.cookies("tsa")("statusfilt")
        else
        filt = "1"
        end if
    end if
	
    response.cookies("tsa")("statusfilt") = filt
    
    aftfilt = request("aftfilt")
	sort = Request("sort")
	usedletter = request("l") 
	

   
	
	oimg = "ikon_job_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = job_txt_625
	
    if media = "print" OR request("print") = "j" then
	    call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
    end if
	

    'Response.write "filt" & filt
    'Response.end
    chk0 = ""
    chk1 = ""
	chk2 = ""
	chk3 = ""
	chk4 = ""
	chk5 = ""

    varFilt = " AND (jobstatus = -2"
    for f = 0 to 5
        'Response.write instr(filt, f) & "<br>.."

        if instr(filt, f) <> 0 then
        varFilt = varFilt & " OR jobstatus = "& f
            
            select case f
            case 0
            chk0 = "CHECKED"
            case 1
            chk1 = "CHECKED"
            case 2
            chk2 = "CHECKED"
            case 3
            chk3 = "CHECKED"
            case 4
            chk4 = "CHECKED"
            case 5
            chk5 = "CHECKED"
            end select
        
        
        end if



    next

    varFilt = varFilt & ")"

    'Response.write "varFilt" & varFilt


	

	
	select case aftfilt
	case "serviceaft"
	varFilt = varFilt & " AND serviceaft <> 0" 'Service aftaler.
	chk6 = "CHECKED"
	chk7 = ""
	chk8 = ""
	case "ikkeserviceaft"
	varFilt = varFilt & " AND serviceaft = 0" 'Ikke Service aftaler.
	chk6 = ""
	chk7 = "CHECKED"
	chk8 = ""
	case else
	varFilt = varFilt & "" 'Vis alle.
	chk6 = ""
	chk7 = ""
	chk8 = "CHECKED"
	end select
	
	
	'* er der valgt et bogstav ****
	if len(usedletter) <> 0 then
	varFilt = varFilt & " AND jobnavn LIKE '"&request("l")&"%'" 
	'showFilter = showFilter & " der starter med:<b> " & request("l") &"</b>"
	else
	varFilt = varFilt 
	'showFilter = showFilter
	end if
	
	
		if len(request("FM_kunde")) <> 0 then
		fmkunde = request("FM_kunde")
		else
		fmkunde = 0
		end if
	    
        if len(request("frompost")) <> 0 then
            if len(trim(request("FM_vis_timepriser"))) <> 0 then
            vis_timepriser = 1
            else
            vis_timepriser = 0
            end if

           
        else
            if request.Cookies("job")("vis_timepriser") <> "" then
            vis_timepriser = request.Cookies("job")("vis_timepriser")
            else
            vis_timepriser = 0
            end if
        end if

        Response.Cookies("job")("vis_timepriser") = vis_timepriser

        if vis_timepriser = 1 then
        vis_timepriserCHK = "CHECKED"
        else
        vis_timepriserCHK = ""
        end if

	
	'''** s�g i akt ***'
	if len(trim(request("FM_sogakt"))) <> 0 then
			sogakt = 1
			sogaktCHK = "CHECKED"
	else
	        sogakt = 0
	        sogaktCHK = ""		
    end if	
	
	if len(trim(request("FM_sogakt_easyreg"))) <> 0 then
			sogakt_easyreg = 1
			sogakt_easyregCHK = "CHECKED"
	else
	        sogakt_easyreg = 0
	        sogakt_easyregCHK = ""		
    end if	    
   
        
        
        '*************************************************
        '*** Er der s�gt p� en specifikt jobnr/jobnavn ***
		'*************************************************
        
        if len(request("FM_kunde")) <> 0 then 'er der s�gt/submitted??
        jobnr_sog = request("jobnr_sog")
        response.cookies("tsa")("jobsog") = jobnr_sog
        visliste = 1
        else

            if request.cookies("tsa")("jobsog") <> "" then
           
            jobnr_sog = request.cookies("tsa")("jobsog")
            
                if jobnr_sog = "%" then
                jobnr_sog = ""
                end if

            else
            jobnr_sog = ""
            end if
        visliste = 0
        end if




		if len(jobnr_sog) <> 0 AND jobnr_sog <> "-1" OR (cint(visliste) = 1) then
			jobnr_sog = jobnr_sog
			
			if cint(sogakt) = 1 then
			sogeKri = sogeKri & " (a.navn LIKE '%"& jobnr_sog &"%' "
			else
			    

                if instr(jobnr_sog, ">") > 0 OR instr(jobnr_sog, "<") > 0 OR instr(jobnr_sog, "--") > 0 OR instr(jobnr_sog, ";") > 0 then
           
                    if instr(jobnr_sog, ">") > 0 then
                    sogeKri = sogeKri &" (j.jobnr > "& replace(trim(jobnr_sog), ">", "") &" "
                    end if

                    if instr(jobnr_sog, "<") > 0 then
                    sogeKri = sogeKri &" (j.jobnr < '"& replace(trim(jobnr_sog), "<", "") &"' "
                    end if

                    if instr(jobnr_sog, "--") > 0 then
                    jobnr_sogArr = split(jobnr_sog, "--")
               
                    for t = 0 to 1
                
                    if t = 0 then
                    jobSogKriA = jobnr_sogArr(0)
                    else
                    jobSogKriB = jobnr_sogArr(1)
                    end if
                
                    next

                    sogeKri = sogeKri &" (j.jobnr BETWEEN '"& trim(jobSogKriA) &"' AND '"& trim(jobSogKriB) &"'"
                    end if


                    if instr(jobnr_sog, ";") > 0 then
            
                    sogeKri = " (j.jobnr = '-1' "

                    jobnr_sogArr = split(jobnr_sog, ";")
               
                    for t = 0 to UBOUND(jobnr_sogArr)
                     sogeKri = sogeKri &" OR j.jobnr = '"& trim(jobnr_sogArr(t)) &"'"
                    next

                              

                    end if

                else

                select case lto
                case "dencker", "intranet - local"
                jobnrSqlkr = "'%"& jobnr_sog &"%'"
                case else
                jobnrSqlkr = "'"& jobnr_sog &"'"
                end select


                sogeKri = " (j.jobnr LIKE "& jobnrSqlkr &" OR j.jobnavn LIKE '"& jobnr_sog &"%' OR j.id LIKE '"& jobnr_sog &"' OR Kkundenavn LIKE '"& jobnr_sog &"%' OR Kkundenr LIKE '"& jobnr_sog &"' OR rekvnr LIKE '"& jobnr_sog &"'"
			

                end if
            
            end if
			
			sogeKri = sogeKri & ") AND "
			
			
			show_jobnr_sog = jobnr_sog
			
		else
			jobnr_sog = "-1"
			sogeKri = " (j.id = -1) AND "
			show_jobnr_sog = ""
		end if


    '**** Vis kun med Easyreg ***'
    if cint(sogakt_easyreg) = 1 then
    sogeKri = sogeKri & " a.easyreg = 1 AND "
    end if

		
    '**** Projektgrupper ****'
    if len(trim(request("FM_prjgrp"))) <> 0 then
    prjgrp = request("FM_prjgrp")
    response.Cookies("job")("prjgrp") = prjgrp
    else
        if request.Cookies("job")("prjgrp") <> "" then
        prjgrp = request.Cookies("job")("prjgrp")
        else
        prjgrp = 10 '** Alle
        end if
    end if
		
	
	call datocookie
	
	'**** Brug datokriterie filer ****
	if len(request("usedatokri")) <> 0 then
		if request("usedatokri") = "j" then
		usedatoKri = "j"
		datoKriJa = "CHECKED"
		datoKriNej = ""
		else
		usedatoKri = "n"
		datoKriJa = ""
		datoKriNej = "CHECKED"
		end if
	else
		usedatoKri = request.Cookies("job")("cusedatokri")
		if usedatoKri = "j" then
		datoKriJa = "CHECKED"
		datoKriNej = ""
		else
		usedatoKri = "n"
		datoKriJa = ""
		datoKriNej = "CHECKED"
		end if
	end if
	


    if len(trim(request("FM_hd_kpers"))) <> 0 then
    hd_kpers = request("FM_hd_kpers")
    else
    hd_kpers = -1
    end if

	'*** Inds�tter cookie ***
	Response.Cookies("job")("cusedatokri") = usedatoKri
	Response.Cookies("job").Expires = date + 65
	
	if usedatoKri = "j" then
	datoKri = " AND jobslutdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
	else
	datoKri = ""
	end if	
	'***** 
	
	%>
	
	<%call filterheader_2013(0,0,900,oskrift)%>
	<table cellspacing=0 width=100% cellpadding=0 border=0>
    <tr><td valign=top>

    <table border=0 cellspacing=0 cellpadding=0>
    <form action="jobs.asp?menu=job&shokselector=1&frompost=1" method="post" name="jobnr" id="jobnr">
	<tr>
		<td colspan=2><b><%=job_txt_202 %>:</b>&nbsp;&nbsp;
		<select name="FM_kunde" id="FM_kunde" style="width:380px;" onchange="submit();">
		<option value="0"><%=job_txt_002 %> <%=writethis%></option>
		<%
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(fmkunde) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select>
		</td>
	</tr>


    <tr>
    <input type="hidden" value="<%=hd_kpers %>" name="FM_hd_kpers" id="FM_hd_kpers" />
		<td colspan=2 style="padding:5px 2px 2px 0px;"><b><%=job_txt_203 %>:</b>&nbsp;&nbsp;
		<select name="FM_kundekpers" id="FM_kundekpers" style="width:342px;">
		<option value="0"><%=job_txt_002 & " " %>(<%=job_txt_204 %>)</option>
		</select>
        </td>
	</tr>




	<tr><td colspan=2>
	<%
	if len(request("FM_medarb_jobans")) <> 0 then
	medarb_jobans = request("FM_medarb_jobans")
	response.cookies("tsa")("jobans") = medarb_jobans
	else
	    if request.cookies("tsa")("jobans") <> "" then
	    medarb_jobans = request.cookies("tsa")("jobans")
	    else
	    medarb_jobans = 0
	    end if
	end if
	
	response.cookies("tsa").expires = date + 45
	%>
	
    <br /><br />
	<b><%=job_txt_023 %>:</b> (<%=job_txt_206 %>)&nbsp;&nbsp; 
	<select name="FM_medarb_jobans" id="FM_medarb_jobans" style="width:310px;">
            <option value="0"><%=job_txt_002 & " " %>(<%=job_txt_205 %>)</option>
            <%
            mNavn = "Alle"
            
             strSQL = "SELECT mnavn, mnr, mid, init FROM medarbejdere WHERE mansat <> '2' AND mansat <> '4' ORDER BY mnavn"
             oRec.open strSQL, oConn, 3
             while not oRec.EOF
                
                if cdbl(medarb_jobans) = oRec("mid") then
                selThis = "SELECTED"
                else
                selThis = ""
                end if
                
                
                %>
             <option value="<%=oRec("mid")%>" <%=selThis%>><%=oRec("mnavn") %> (<%=oRec("mnr") %>)
             <%if len(trim(oRec("init"))) <> 0 then %>
              - <%=oRec("init") %>
              <%end if %>
              </option>
             <%
             oRec.movenext
             wend
             oRec.close
             %>
             
              
        </select>
	</td></tr>
    <tr><td colspan=2><br />
    <b><%=job_txt_132 %></b>: (<%=job_txt_207 %>..)<br />
    <select  name="FM_prjgrp" id="FM_prjgrp" style="width:433px;">
    <%call progrupperIdNavn(prjgrp) %>
    
    </select>
    
    </td></tr>
   

	<tr bgcolor="#FFFFFF">
		<td colspan=2><br /><b><%=job_txt_208 %>:</b><br />
		<input type="text" name="jobnr_sog" id="jobnr_sog" value="<%=show_jobnr_sog%>" style="width:433px; border:2px #6CAE1C solid;">&nbsp;
		<br />
        (% = wildcard) <%=job_txt_209 %> ";" <%=job_txt_210 %>, ">", "<" <%=job_txt_211 %> "--" (<%=job_txt_212 %>) <%=job_txt_213 %><br />
        
        <input id="FM_sogakt" name="FM_sogakt" type="checkbox" value="1" <%=sogaktCHK%> /> <%=job_txt_214 %><br />
        <%
        call showEasyreg_fn()    
        if cint(showEasyreg_val) = 1 then %>
        <input name="FM_sogakt_easyreg" type="checkbox" value="1" <%=sogakt_easyregCHK%> /> Vis kun job hvor der er Easyreg aktiviteter
        <%end if %>
        <br /><br /><br />
        <div style="border:1px #CCCCCC solid; padding:5px;">
        <input type="checkbox" name="FM_vis_timepriser" <%=vis_timepriserCHK %> value="1" />
        <b><%=job_txt_215 %></b> (<%=job_txt_216 %>)
        <br />
        <%=job_txt_217 %>: <input type="radio" value="0" name="FM_sorter_tp" <%=sorttpCHK0 %> /> <%=job_txt_218 %>
        &nbsp;<input type="radio" value="1" name="FM_sorter_tp" <%=sorttpCHK1 %> />  <%=job_txt_219 %>
        </div>
        <br /><br />&nbsp;
		</td>
	</tr>
    </table>
    </td>
   
    <td valign=top>
    &nbsp;
     </td>
   
    <td valign=top>

    <table border=0 cellspacing=0 cellpadding=0>

     <tr><td colspan=2>
         <b><%=job_txt_144 %>:</b> <!--<%=strFomr_rel %> <%=left(strFomr_Gblnavn, 75) %>--><br />

                              
                                <%
                                
                                
                                strSQLf = "SELECT id, navn FROM fomr WHERE id <> 0 ORDER BY navn"%>
                                <select name="FM_fomr" id="Select2" multiple="multiple" size="6" style="width:410px;">
                                <%if instr(strFomr_rel, "#0#") <> 0 then 
                                f0sel = "SELECTED"
                                else
                                f0sel = ""
                                end if%>

                                <option value="#0#" <%=f0sel %>><%=job_txt_002 &" " %>(<%=job_txt_205 %>)</option>
                                    
                                    <%oRec.open strSQLf, oConn, 3
                                    while not oRec.EOF
                                    
                                    if instr(strFomr_rel, "#"&oRec("id")&"#") <> 0 then
                                    fSel = "SELECTED"
                                    else
                                    fSel = ""
                                    end if
                                    
                                    %>
                                    <option value="#<%=oRec("id")%>#" <%=fSel %>><%=oRec("navn") %></option>
                                    <%
                                    oRec.movenext
                                    wend
                                    oRec.close
                                    %>
                                </select>

    </td></tr>

	<tr>
		<td colspan=2 valign=top style="padding-top:15px;"><b><%=job_txt_220 %>:</b><br />
       
		
        <input type="CHECKBOX" name="filt" value="1" <%=chk1%>/> <%=jobstatus_txt_001 %><br />
        <input type="CHECKBOX" name="filt" value="2" <%=chk2%>/> <%=jobstatus_txt_002 %><br />
        <input type="CHECKBOX" name="filt" value="3" <%=chk3%>/> <%=jobstatus_txt_003 %><br />
        <input type="CHECKBOX" name="filt" value="4" <%=chk4%>/> <%=jobstatus_txt_004 %><br />
        <input type="CHECKBOX" name="filt" value="0" <%=chk0%>/> <%=jobstatus_txt_005 %><br />
        <input type="CHECKBOX" name="filt" value="5" <%=chk5%>/> <%=jobstatus_txt_006 %><br />

		</td>
	</tr>
	<tr>
		<td colspan=2><br /><br /><b><%=job_txt_226 %>?</b><br />
		<input type="radio" name="aftfilt" value="visalle" <%=chk8%>>&nbsp;<%=job_txt_323 %>&nbsp;&nbsp;
		<input type="radio" name="aftfilt" value="serviceaft" <%=chk6%>>&nbsp;<%=job_txt_018 %>&nbsp;&nbsp;
		<input type="radio" name="aftfilt" value="ikkeserviceaft" <%=chk7%>>&nbsp;<%=job_txt_019 %>&nbsp;&nbsp;
		</td>
	</tr>
   
	<tr>
		<td colspan=2><br /><br />

		<b><%=job_txt_227 %>:</b><br /> 
            <%=job_txt_019 %> <input type="radio" name="usedatokri" value="n" <%=datoKriNej%>><br />
            <%=job_txt_018 %>:
		<input type="radio" name="usedatokri" value="j" <%=datoKriJa%>> <span style="font-size:9px; color:#999999;">(<%=job_txt_228 %>)</span> <br />
		<table width=100% cellpadding=0 cellspacing=0>
			<tr><!--#include file="inc/weekselector_b.asp"--></tr>
		</table>
		</td>
    </tr>
    <tr>
		<td colspan=2 align=right valign=bottom><br /><br /><br /><br /><input id="Submit1" type="submit" value="<%=job_txt_229 %> >>"/>
            &nbsp;</td>
	</tr>
	</form>
	</table>

    </td>
    </tr>
    </table>
	
	<!-- filter header -->
	</td></tr></table>
	</div>
	<br /><br />

	
	<%
	
	tTop = 0
	tLeft = 0
	tWdth = 1224
	
	
	call tableDiv(tTop,tLeft,tWdth)
	

    '** V�lger sql s�tning efter sortering og filter ***
	
	if fmkunde <> 0 then
	varJobknrKri = "j.jobknr = "& Request("FM_Kunde") &" AND "
	else
		'if request("fromvemenu") = "j" then
		'varJobknrKri = "j.jobknr = 0 AND "
		'else
		varJobknrKri = ""
		'end if
	end if
	
	if usedatoKri = "j" then
	varOrder = "jobslutdato DESC"
	else
	varOrder = "kkundenavn, jobnavn"
	end if
	
    '** jobans **'
	if cdbl(medarb_jobans) = 0 then
	jobansKri = ""
	else
	jobansKri = " (jobans1 = " & medarb_jobans & " OR jobans2 = "& medarb_jobans &") AND "
	end if
	

    '** kontaktpersoner hos kunde **'
    if cdbl (hd_kpers) <> -1 AND cdbl(hd_kpers) <> 0 then
    kpersSQLkri = " AND kundekpers = " & hd_kpers
    else
    kpersSQLkri = ""
    end if


    if prjgrp <> 0 AND prjgrp <> 10 then
    prjgrpSQLkri = " AND (projektgruppe1 = "&prjgrp&" "_
    &" OR projektgruppe2 = "&prjgrp&" OR projektgruppe3 = "&prjgrp&" OR projektgruppe4 = "&prjgrp&" "_
    &" OR projektgruppe5 = "&prjgrp&" OR projektgruppe6 = "&prjgrp&" OR projektgruppe7 = "&prjgrp&" OR projektgruppe8 = "&prjgrp&" OR projektgruppe9 = "&prjgrp&" OR projektgruppe10 = "&prjgrp&")"
    else
    prjgrpSQLkri = ""
    end if 
	




    if cint(vis_timepriser) <> 1 then
	%>
	
	
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#ffffff" width=100%>
	<form action="jobs.asp?menu=job&func=opdjobliste&jobnr_sog=<%=jobnr_sog%>&filt=<%=filt%>&fm_kunde_sog=<%=fmkunde%>&FM_kunde=<%=fmkunde%>" method="post">
    <input type="hidden" value="<%=lto %>" id="jq_lto" />
    <input type="hidden" value="<%=level %>" id="jq_level" />

	<%sub tabletop%>
	
	<tr bgcolor="#5582D2">
		<td class='alt' valign=bottom style="width:200px;"><b><%=job_txt_022 %>&nbsp;/&nbsp;<%=job_txt_230 %><br /><%=job_txt_021 %>&nbsp;/&nbsp;<%=job_txt_231 %></b></td>
		<td class='alt' style="padding-right:10px;" valign=bottom><b><%=job_txt_232 %></b></td>
        <td class='alt' style="padding-right:10px;" valign=bottom><b><%=job_txt_233 %></b></td>
        <td class='alt' valign=bottom style="width:150px; padding-left:20px;"><b><%=job_txt_234 %> %</b><br />(<%=job_txt_235 %>)</td>
		<td class='alt' valign=bottom style="padding-right:5px; white-space:nowrap;" align=right><b><%=job_txt_236 %></b><br />(<%=job_txt_237 %>)</td>
		<td align=right style="padding-right:5px;" valign=bottom class=alt style="width:100px;"><b><%=job_txt_238 %></b><br>
		<%=job_txt_239 %><br>
		<b><%=job_txt_240 %></b></td>
		<td class='alt' valign=bottom style="padding-left:5px;"><b><%=job_txt_241 %></b>&nbsp;

   <table cellpadding=0 cellspacing=0 border=0 width="100">
       
   <tr><td style="font-size:9px; color:#CCCCCC; white-space:nowrap;"><input type="radio" name="st_sel" value="1" id="st_cls_1" /><%=jobstatus_txt_007 %></td>
   <td style="font-size:9px; color:#CCCCCC; white-space:nowrap;"><input type="radio" name="st_sel"  value="2" id="st_cls_2" /><%=jobstatus_txt_011 %></td></tr>  
   <tr><td style="font-size:9px; color:#CCCCCC; white-space:nowrap;"><input type="radio" name="st_sel"  value="3" id="st_cls_3"/><%=jobstatus_txt_003 %></td> 
   <td style="font-size:9px; color:#CCCCCC; white-space:nowrap;"><input type="radio" name="st_sel" value="4" id="st_cls_4" /><%=jobstatus_txt_012 %></td></tr> 
   <tr><td style="font-size:9px; color:#CCCCCC; white-space:nowrap;"><input type="radio" name="st_sel" value="0" id="st_cls_0" /><%=jobstatus_txt_009 %></td>
	<td style="font-size:9px; color:#CCCCCC; white-space:nowrap;"><input type="radio" name="st_sel" value="0" id="st_cls_5" /><%=jobstatus_txt_006 %></td></tr>
  </table>	
            
            </td>
		<td class='alt' valign=bottom align=center style="padding-left:5px; width:100px;"><b><%=job_txt_247 %></b></td>
		<td class='alt' valign=bottom style="width:160px;"><b><%=job_txt_248 %></b>
        
        <!--<br />
        <%=job_txt_249 %><br />
        <span style="font-size:9px; line-height:10px; color:#CCCCCC;"><%=job_txt_250 %> <br />(<%=job_txt_251 %>) <br />/ <%=job_txt_252 %></span></td>
        -->

		<td class='alt' valign=bottom colspan=2><b><%=job_txt_132 %></b>
		</td>
		
    </tr>
	<%
	end sub
	
	
	
	
	'***********************************************************
	'*** Main SQL **********************************************
	'***********************************************************
	strSQL = "SELECT j.id, jobnavn, jobnr, kkundenavn, kid, jobknr, jobTpris, jobstatus, jobstartdato, "_
	&" jobslutdato, j.budgettimer, fakturerbart, Kkundenr, ikkebudgettimer, jobans1, jobans2, jobans3, jobans4, jobans5, fastpris, jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, virksomheds_proc, "_
	&" s.navn AS aftnavn, rekvnr, tilbudsnr, sandsynlighed, jo_bruttooms, kundekpers, serviceaft, lukkedato, preconditions_met, jo_valuta, jo_valuta_kurs"
	
    strSQL = strSQL &", j.projektgruppe1, j.projektgruppe2, j.projektgruppe3, j.projektgruppe4, j.projektgruppe5, j.projektgruppe6, j.projektgruppe7, j.projektgruppe8, j.projektgruppe9, j.projektgruppe10 "
	   
	
	if cint(sogakt) = 1 OR cint(sogakt_easyreg) = 1 then
	    strSQL = strSQL &" FROM aktiviteter AS a "
	    strSQL = strSQL &" LEFT JOIN job AS j ON (j.id = a.job)"
	    strSQL = strSQL &" LEFT JOIN kunder ON (kunder.kid = jobknr)"
	else
        strSQL = strSQL &" FROM job AS j "
        strSQL = strSQL &" LEFT JOIN kunder ON (kunder.kid = jobknr)"
    end if

	
	strSQL = strSQL &" LEFT JOIN serviceaft s ON (s.id = serviceaft)"_
	&" WHERE "& varJobknrKri &" "& sogeKri &" "& jobansKri &" kunder.Kid = jobknr " & varFilt &" "& datoKri &" "& prjgrpSQLkri &" "& kpersSQLkri &""_
	&" GROUP BY j.id ORDER BY "&varOrder&" LIMIT 5000"
	
	
    jobids_all = 0
	x = 0
	cnt = 0
	jids = 0
	totReal = 0
	gnsPrisTot = 0
	totRealialt = 0
    thisMid = session("mid")
	jo_valuta_kurs = 100

    'if session("mid") = 1 then
	'Response.write strSQL
	'Response.Flush
    'end if
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF


    jo_valuta_kurs = oRec("jo_valuta_kurs")

    call forretningsomrJobId(oRec("id"))


     if cint(visJobFomr) = 1 OR instr(strFomr_rel, "#0#") <> 0 then


	
	'** Til Export fil ***
	jids = jids & "," & oRec("id")
	
	if cnt = 0 then
	call tabletop
	end if
	

    if len(trim(oRec("budgettimer"))) = "0" OR oRec("budgettimer") = "0" then 
	budgettimer = 0
	else
	budgettimer = oRec("budgettimer")
	end if
	
	if oRec("ikkebudgettimer") > 0 then 
	ikkebudgettimer = oRec("ikkebudgettimer")
	else
	ikkebudgettimer = 0
	end if
	
	'** tildelt total ***
	budgettimertot = (ikkebudgettimer + budgettimer)


	call akttyper2009(2)

	'*** Real timer og Proafsluttet **************************
    select case lto
    case "dencker"
        aty_sql_realhours = aty_sql_realhours & " OR tfaktim = 90"
    case else
        aty_sql_realhours = aty_sql_realhours
    end select

	strSQL = "SELECT sum(timer) AS proafslut FROM timer WHERE Tjobnr = '" & oRec("jobnr") &"' AND ("& aty_sql_realhours &") "
	oRec3.open strSQL, oConn, 3
	
	if len(oRec3("proafslut")) <> 0 then
	proaf = oRec3("proafslut")
	else
	proaf = 0
	end if

	oRec3.close

    totReal = proaf
    restt = (budgettimertot - proaf)
    
    realfakbare = 0
    '*** Real. fakturerbare timer **************************
	strSQL = "SELECT sum(timer) AS realfakbare FROM timer WHERE Tjobnr = '" & oRec("jobnr") &"' AND tfaktim = 1"
	oRec3.open strSQL, oConn, 3
	if not oRec3.EOF then

	if len(oRec3("realfakbare")) <> 0 then
	realfakbare = oRec3("realfakbare")
	else
	realfakbare = 0
	end if

	end if
	oRec3.close
	

    resttfakbare = (budgettimertot - realfakbare)

    'timerTotFakbarePajob = realfakbare 


    'if len(timerTotFakbarePajob) <> 0 then
	'timerTotFakbarePajob = timerTotFakbarePajob
	'else
	'timerTotFakbarePajob = 0
	'end if



		
		'*** bgcolor ***
		if cdbl(id) = oRec("id") then
		bc = "#FFFF99"
		else
			select case right(cnt, 1) 
			case 0,2,4,6,8
			bc = "#FFFFFF"
			case else
			bc = "#EFF3FF" '"#5582D2" '"#d2691e"
			end select
		end if
	
	
	
	sletkanp = "<img src='../ill/slet_16.gif' alt='Slet' border='0'>" 
	
	'if oRec("fakturerbart") = "1" then
		if oRec("jobstatus") = 3 then
		inteksimg = "<img src='../ill/filtertilbud.gif' width='10' height='21' alt='' border='0'>"
		else
		inteksimg = "<img src='../ill/blank.gif' width='15' height='15' alt='' border='0'>"
		'eksternt_job_trans
		end if
	'else
	'inteksimg = "<img src='../ill/internt_job_trans.gif' width='15' height='15' alt='' border='0'>"
	'end if
	
	


    '*** Oms. timer ****'
    jo_bruttooms = oRec("jo_bruttooms")
	jo_fastpris = oRec("fastpris")


    '** Prekonditioner ****'
    preconditions_met = oRec("preconditions_met")
    
	
	if budgettimer <> 0 then
	projektcomplt = ((proaf/budgettimertot)*100)
    projektcompltFakbare = ((realfakbare/budgettimertot)*100)
	else
	projektcompltFakbare = 100
    projektcomplt = 100
	end if
	
	if projektcomplt > 100 then
	showprojektcomplt = projektcomplt
	projektcomplt = 100
	bgdiv = "Crimson"
	else
	showprojektcomplt = projektcomplt
	projektcomplt = projektcomplt
	bgdiv = "yellowgreen"
	end if

    if projektcompltFakbare > 100 then
	showprojektcompltFakbare = projektcompltFakbare
	projektcompltFakbare = 100

	else
	showprojektcompltFakbare = projektcompltFakbare
	projektcompltFakbare = projektcompltFakbare
	
	end if

    
	
	'*** Antal aktiviteter p� job *** KUN AKTIVE I DETTE VIEW 
	strAktnavn = ""
    lastFase = ""
    strSQL2 = "SELECT id, navn, fase, aktstatus, easyreg FROM aktiviteter WHERE job = "& oRec("id") & " AND aktstatus = 1 ORDER BY fase, sortorder, navn"
	oRec5.open strSQL2, oConn, 3
	while not oRec5.EOF 
	x = x + 1

    if lastFase <> oRec5("fase") AND isNull(oRec5("fase")) <> true AND len(trim(oRec5("fase"))) <> 0 then

    strAktnavn = strAktnavn & "<br><b>"& replace(oRec5("fase"), "_", " ") & "</b><br>"
    lastFase = oRec5("fase")
    end if

    strAktnavn = strAktnavn & left(oRec5("navn"), 20) 
        
        if oRec5("aktstatus") = 0 then
        strAktnavn = strAktnavn & " - lukket"
        end if

        if oRec5("aktstatus") = 2 then
        strAktnavn = strAktnavn & " - passiv"
        end if
        
        if oRec5("easyreg") = 1 then
        strAktnavn = strAktnavn & " <span style=""color:red;"">(E!)</span>"
        end if

    strAktnavn = strAktnavn & "<br>"


	oRec5.movenext
	wend
	oRec5.close
	Antal = x
	
	
	%>
	<tr>
		<td bgcolor="#d6dff5" colspan="11"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bc%>">
	
	<%
	'*** Tjekker rettigehder eller om man er jobanssvarlig ***
	editok = 0
	if level <= 2 OR level = 6 then
	editok = 1
	else
			if cint(session("mid")) = oRec("jobans1") OR cint(session("mid")) = oRec("jobans2") OR (cint(oRec("jobans1")) = 0 AND cint(oRec("jobans2")) = 0) then
			editok = 1
			else
			editok = 0   
			end if
	end if
	
	
	              
	
	
	'********* Jobnavn og nr ****
	%>
	<td valign=top style="padding:4px 10px 4px 2px; width:250px;">
      <b><%=oRec("kkundenavn")%></b>&nbsp;(<%=oRec("Kkundenr")%>)<br />
	<%
	'Response.write  cint(session("mid")) &"="& oRec("jobans1") &" OR " & cint(session("mid")) &" = " & oRec("jobans2") &" OR "& cint(oRec("jobans1")) &" = 0 OR" & cint(oRec("jobans2")) &" = 0 then"
			
	
	if editok = 1 then
		'if oRec("fakturerbart") = "1" then%>
		<a href="jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=<%=oRec("fakturerbart")%>&jobnr_sog=<%=jobnr_sog%>&filt=<%=filt%>&fm_kunde_sog=<%=fmkunde%>"><%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)</a>&nbsp;
		<%'else%>
		<!--<a href="jobs.asp?menu=job&func=red&id=<=oRec("id")%>&int=<=oRec("fakturerbart")%>&jobnr_sog=<=jobnr_sog%>&filt=<=filt%>&fm_kunde_sog=<=fmkunde%>" class=vmenu><=oRec("jobnavn")%>&nbsp;&nbsp;(<=oRec("jobnr")%>)</a>&nbsp;-->
		<%'end if
	else
		%>
		<b><%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)</b>&nbsp;
		<%
	end if
	

    if oRec("fastpris") = 1 then
    strFasptpris = job_txt_324
    else
    strFasptpris = job_txt_325
    end if


    %>
    <span style="color:green; font-size:10px;">(<%=strFasptpris %>)</span>

    <%
	
	
	
	'******************************
	%>
    <span style="color:#999999; font-size:9px;"><i>
    <%
		
						'*** Jobansvarlige ***
                        if oRec("jobans1") <> 0 then
						call meStamData(oRec("jobans1"))
                        %>
						<%=meNavn%> 
                             <%if oRec("jobans_proc_1")  <> 0 then %>
                            (<%=oRec("jobans_proc_1") %>%)
                            <%end if %>

                        <%end if
                        

                        '*** Jobejer 2 ***
                        if oRec("jobans2") <> 0 then
						call meStamData(oRec("jobans2"))
                        %>
						, <%=meNavn%> 
                        
                        <%if oRec("jobans_proc_2")  <> 0 then %>
                        (<%=oRec("jobans_proc_2") %>%)
                        <%end if %>

                        <%end if

                        '*** Job medansvarlige ***
                        if oRec("jobans3") <> 0 then
						call meStamData(oRec("jobans3"))
                        %>
						, <%=meNavn%> 
                         <%if oRec("jobans_proc_3")  <> 0 then %>
                        (<%=oRec("jobans_proc_3") %>%)

                        <%end if
                        
                        end if


                        '*** Job medansvarlige ***
                        if oRec("jobans4") <> 0 then
						call meStamData(oRec("jobans4"))
                        %>
						, <%=meNavn%> 
                             <%if oRec("jobans_proc_4")  <> 0 then %>
                            (<%=oRec("jobans_proc_4") %>%)
                            <%end if %>

                        <%end if


                        '*** Job medansvarlige ***
                        if oRec("jobans5") <> 0 then
						call meStamData(oRec("jobans5"))
                        %>
						, <%=meNavn%> 
                             <%if oRec("jobans_proc_5")  <> 0 then %>
                            (<%=oRec("jobans_proc_5") %>%)
                            <%end if %>

                        <%end if

                        if level = 1 then
                             if (lto <> "epi" AND lto <> "epi2017") OR (lto = "epi" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59 OR thisMid = 1720)) OR (lto = "epi_cati" AND (thisMid = 6 OR thisMid = 11 OR thisMid = 59)) OR (lto = "epi_no" AND thisMid = 2) OR (lto = "epi_ab" AND thisMid = 2) OR (lto = "epi_sta" AND thisMid = 2) OR (lto = "epi_uk" AND thisMid = 2) then 
                                if oRec("virksomheds_proc") <> 0 then%>
                                <br /><b><%=lto %>: </b> (<%=oRec("virksomheds_proc") %>%)
                                <%end if
                            end if
                        end if
                      
						
						'**********************
						%></i></span>


                        <%'*** kontakpersoner hos kunde ***'
                         
                         kpersNavn = ""
                         if cint(oRec("kundekpers")) <> 0 then

                         strSQLkpers = "SELECT navn FROM kontaktpers WHERE id = " & oRec("kundekpers")
                         oRec6.open strSQLkpers, oConn, 3
                         if not oRec6.EOF then

                         kpersNavn = oRec6("navn")

                         end if
                         oRec6.close

                         %>
                         <span style="color:#8caae6; font-size:9px;"><br /><%=job_txt_253 %>: <%=kpersNavn %></span>
                         <%

                         end if
                         
                         %>
							
                        <%

                            if month(oRec("jobstartdato")) = 1 then jobstartdatoMonthTxt = job_txt_588 end if
                            if month(oRec("jobstartdato")) = 2 then jobstartdatoMonthTxt = job_txt_589 end if
                            if month(oRec("jobstartdato")) = 3 then jobstartdatoMonthTxt = job_txt_590 end if
                            if month(oRec("jobstartdato")) = 4 then jobstartdatoMonthTxt = job_txt_591 end if
                            if month(oRec("jobstartdato")) = 5 then jobstartdatoMonthTxt = job_txt_592 end if
                            if month(oRec("jobstartdato")) = 6 then jobstartdatoMonthTxt = job_txt_593 end if
                            if month(oRec("jobstartdato")) = 7 then jobstartdatoMonthTxt = job_txt_594 end if
                            if month(oRec("jobstartdato")) = 8 then jobstartdatoMonthTxt = job_txt_595 end if
                            if month(oRec("jobstartdato")) = 9 then jobstartdatoMonthTxt = job_txt_596 end if
                            if month(oRec("jobstartdato")) = 10 then jobstartdatoMonthTxt = job_txt_597 end if
                            if month(oRec("jobstartdato")) = 11 then jobstartdatoMonthTxt = job_txt_598 end if
                            if month(oRec("jobstartdato")) = 12 then jobstartdatoMonthTxt = job_txt_599 end if

                            if month(oRec("jobslutdato")) = 1 then jobslutdatoMonthTxt = job_txt_588 end if
                            if month(oRec("jobslutdato")) = 2 then jobslutdatoMonthTxt = job_txt_589 end if
                            if month(oRec("jobslutdato")) = 3 then jobslutdatoMonthTxt = job_txt_590 end if
                            if month(oRec("jobslutdato")) = 4 then jobslutdatoMonthTxt = job_txt_591 end if
                            if month(oRec("jobslutdato")) = 5 then jobslutdatoMonthTxt = job_txt_592 end if
                            if month(oRec("jobslutdato")) = 6 then jobslutdatoMonthTxt = job_txt_593 end if
                            if month(oRec("jobslutdato")) = 7 then jobslutdatoMonthTxt = job_txt_594 end if
                            if month(oRec("jobslutdato")) = 8 then jobslutdatoMonthTxt = job_txt_595 end if
                            if month(oRec("jobslutdato")) = 9 then jobslutdatoMonthTxt = job_txt_596 end if
                            if month(oRec("jobslutdato")) = 10 then jobslutdatoMonthTxt = job_txt_597 end if
                            if month(oRec("jobslutdato")) = 11 then jobslutdatoMonthTxt = job_txt_598 end if
                            if month(oRec("jobslutdato")) = 12 then jobslutdatoMonthTxt = job_txt_599 end if

                        %>
						
	<br><span style="font-size:9px;"><%=day(oRec("jobstartdato"))&". "& jobstartdatoMonthTxt &" "& year(oRec("jobstartdato")) %> - <%=day(oRec("jobslutdato"))&". "& jobslutdatoMonthTxt &" "& year(oRec("jobslutdato")) %></span>
	
    
	<% 
	if len(trim(oRec("rekvnr"))) <> 0 then
    %>
    <span style="font-size:9px; color:#999999;">
    <%
	Response.Write "<br>"&job_txt_254&": "& oRec("rekvnr") 
    %>
     </span>
    <%
	end if

  
	if len(trim(oRec("aftnavn"))) <> 0 then
    %><span style="font-size:9px; background-color:#FFFFe1; color:#000000;"><%
	Response.Write "<br>"&job_txt_255&": "& oRec("aftnavn") 
    %>
     </span>
    <%
	end if
	
	if oRec("jobstatus") = 3 then
    %><span style="font-size:9px; background-color:#ffdfdf; color:#000000;"><%
	Response.Write "<br>"& job_txt_063 &": "& oRec("tilbudsnr") &" ("& oRec("sandsynlighed") &" %)"
    %>
     </span>
    <%
	end if


    select case preconditions_met
      case 0
    preconditions_met_Txt = ""
      case 1
    preconditions_met_Txt = "<br><span style='color:#6CAE1C; font-size:10px; background:#DCF5BD;'>"&job_txt_257&"</span>"
    case 2
    preconditions_met_Txt = "<br><span style='color:red; font-size:10px; background:pink;'>"&job_txt_258&"!</span>"
    case else
    preconditions_met_Txt = ""
    end select


    %>
    <%=preconditions_met_Txt %>
	
	&nbsp;</td>
	
        <td valign=top style="padding:4px 10px 0px 5px;">
            <%=job_txt_259 %>
		 <%
		'********* Aktiviteter ****
		if editok = 1 then%>
        <a href="#" onclick="Javascript:window.open('aktiv.asp?menu=job&jobid=<%=oRec("id")%>&jobnavn=<%=oRec("jobnavn")%>&rdir=job3&nomenu=1', '', 'width=1354,height=800,resizable=yes,scrollbars=yes')" class=vmenu>(<%=Antal%>)</a>
		<%else%>
		<b>(<%=Antal%>)</b>
		<%end if%>

        <%
        'if visAkt <> 1 then
        %>
        <div style="font-size:10px; height:80px; width:120px; overflow:auto; color:#999999; padding-right:5px;">
           
            
            <%=strAktnavn %></div>

        <%'end if
		'**************************
		%>
		</td>
		

        

        <td valign=top style="padding:4px 5px 5px 5px;" class=lille><%=strFomr_navn %>&nbsp;</td>

		<td align="left" valign=top style="padding:10px 5px 0px 5px;">
		<%if proaf <> 0 then %>
		
		
		<div style="font-size:9px; color:#999999,"><%=job_txt_260 %>:
		<span style="width:<%=cint(left(projektcomplt, 3))%>px; background-color:<%=bgdiv%>; height:10px; padding:2px 0px 2px 2px; font-size:9px;">
		<%if showprojektcomplt > 0 then%>
		<%=formatpercent(showprojektcomplt/100, 0)%>
		<%end if%>
		</span>
        </div>

        <div style="font-size:9px; color:#999999;"><%=job_txt_617 &":" %>
        <span style="width:<%=cint(left(projektcompltFakbare, 3))%>px; background-color:#CCCCCC; color:#000000; height:10px; padding:2px 0px 2px 2px; font-size:9px;">
		<%if showprojektcompltFakbare > 0 then%>
		<%=formatpercent(showprojektcompltFakbare/100, 0)%>
		<%end if%>
		</span>
            </div>

      

		
		
		
		<%end if %>
		</td>

        <%call valutakode_fn(oRec("jo_valuta")) %>
		<td align=right valign=top style="padding:4px 5px 0px 5px;" class=lille>
		<%=formatnumber(oRec("jo_bruttooms"), 2) &" "& valutaKode_CCC %> 
		</td>
		
		<td align=right valign=top style="padding:4px 15px 4px 4px;" class=lille><%=job_txt_261 %>: <%=formatnumber(budgettimertot)%><br>
		<%=job_txt_239 %>:  <%=formatnumber(proaf)%><br>
        <span style="color:#999999;"><%=job_txt_262 %>: <%=formatnumber(realfakbare, 2) %></span> <br />
        -----------------------<br />
        <%=job_txt_263 %> = <%=formatnumber(restt)%><br />
        <span style="color:#999999;"><%=formatnumber(resttfakbare) %></span>
        =====================
		</td>
		
		<td valign=top style="padding:4px 5px 0px 5px;"><%
		
		'if oRec("jobstatus") <> "3" then 'Tilbud
		
        stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
        stCHK3 = ""
        stCHK4 = ""
        stCHK5 = ""
        stBgcol0 = "#999999"
		stBgcol1 = "#999999"
		stBgcol2 = "#999999"
        stBgcol3 = "#999999"
        stBgcol4 = "#999999"
        stBgcol5 = "#999999"

        lkDato = ""

		select case oRec("jobstatus")
        case "0"
        lkDato = "("& formatdatetime(oRec("lukkedato"), 2) &")"
        stCHK0 = "CHECKED"
        stBgcol0 = "darkred"
		case "1"
		stCHK1 = "CHECKED"
        stBgcol1 = "green"
        case "2"
        stCHK2 = "CHECKED"
        stBgcol2 = "#999999"
		case "0"
        stCHK0 = "CHECKED"
        stBgcol0 = "Crimson"
        case "3"
        stCHK3 = "CHECKED"
        stBgcol3 = "#000000"
        case "4"
        stCHK4 = "CHECKED"
        stBgcol4 = "#5582d2"
        case "5"
        stCHK5 = "CHECKED"
        stBgcol5 = "#5582d2"
		end select
		
		if editok = 1 then
		%>
        <!--
		<select name="FM_listestatus" id="FM_listestatus" style="background-color:<%=selbgcol%>; font-size:9px;">
		<option value="1" <%=stCHK1%>>Aktiv</option>
		<option value="0" <%=stCHK0%>>Lukket</option>
		<option value="2" <%=stCHK2%>>Passiv</option>
		</select>
        -->

        <%select case lto
            case "mpt", "local - intranet"

                if level = 1 then
                wprotec = ""
                else
                
                    if stCHK0 = "CHECKED" then 'Hvis job er lukket m� kun admin �ndre status
                    wprotec = "readonly"
                    else
                    wprotec = ""
                    end if

                end if

            case else

                if level <= 2 OR level = 6 then
                wprotec = ""
                else

                    if stCHK0 = "CHECKED" then
                    wprotec = "readonly"
                    else
                    wprotec = ""
                    end if

                end if

         end select  %>
        
       
        <%if wprotec <> "readonly" then%>
        <input type="radio" class="FM_listestatus_1" name="FM_listestatus_<%=oRec("id")%>" value="1" id="FM_listestatus_1_<%=oRec("id")%>" <%=stCHK1%>/><span style="color:<%=stBgcol1%>; font-size:9px;"><%=jobstatus_txt_007 %></span><br />
        <input type="radio" class="FM_listestatus_2" name="FM_listestatus_<%=oRec("id")%>" value="2" id="FM_listestatus_2_<%=oRec("id")%>" <%=stCHK2%>/><span style="color:<%=stBgcol2%>; font-size:9px;"><%=jobstatus_txt_011 %></span><br />
        <input type="radio" class="FM_listestatus_3" name="FM_listestatus_<%=oRec("id")%>" value="3" id="FM_listestatus_3_<%=oRec("id")%>" <%=stCHK3%>/><span style="color:<%=stBgcol3%>; font-size:9px;"><%=jobstatus_txt_003 %><br />
        <input type="radio" class="FM_listestatus_4" name="FM_listestatus_<%=oRec("id")%>" value="4" id="FM_listestatus_4_<%=oRec("id")%>" <%=stCHK4%>/><span style="color:<%=stBgcol4%>; font-size:9px;"><%=jobstatus_txt_012 %></span><br />
        <input type="radio" class="FM_listestatus_5" name="FM_listestatus_<%=oRec("id")%>" value="5" id="FM_listestatus_5_<%=oRec("id")%>" <%=stCHK5%>/><span style="color:<%=stBgcol5%>; font-size:9px;"><%=jobstatus_txt_006 %></span><br />
        <%end if %>

        <input type="radio" class="FM_listestatus_0" name="FM_listestatus_<%=oRec("id")%>" value="0" id="FM_listestatus_0_<%=oRec("id")%>" <%=stCHK0%>/><span style="color:<%=stBgcol0%>; font-size:9px;"><%=jobstatus_txt_009 %> <br /><%=lkDato %></span>
        


            <%jobids_all = jobids_all & ", "& oRec("id") %>
         


		<%else%>
		<%=stNavn%>
		<input type="hidden" name="FM_listestatus" id="Hidden1" value="<%=oRec("jobstatus")%>">
		
		<%end if %>
		
		<%'else%>
        <!--
		<b>Tilbud</b><br />
		<span style="font-size:9px; font-family:arial;"><%=oRec("tilbudsnr") %> <br /><%=oRec("sandsynlighed") %> %</span>
		<input type="hidden" name="FM_listestatus" id="FM_listestatus" value="<%=oRec("jobstatus")%>">
        -->
		<%'end if%>
		
		
		<input type="hidden" name="FM_listeid" id="FM_listeid" value="<%=oRec("id")%>">
		</td>
	
		<td valign=top style="padding:4px 5px 0px 5px; white-space:nowrap;">
		<%if editok = 1 then %>
		<a href="job_print.asp?id=<%=oRec("id")%>" class=rmenu>Print / PDF >></a><br />
		<a href="job_kopier.asp?func=kopier&id=<%=oRec("id")%>&fm_kunde=<%=oRec("kid")%>&filt=<%=request("filt")%>" class=rmenu><%=job_txt_264 %> >></a><br />
         <a href="materialer_indtast.asp?id=<%=oRec("id")%>&fromsdsk=0&aftid=<%=oRec("serviceaft")%>" target="_blank" class=rmenu><%=job_txt_265 %> >></a>

         <%
         
        viskunabnejob0 = "1"
        viskunabnejob1 = "" 
        viskunabnejob2 = ""
        select case oRec("jobstatus")
        case 0
        viskunabnejob0 = "1"
        case 1,2
        viskunabnejob2 = "1"
        case 3
        viskunabnejob1 = "1"
        end select

         select case lto
        case "synergi1", "xintranet - local"


        call licensStartDato()
        dtlink_stdag = startDatoDag
		dtlink_stmd = startDatoMd
		dtlink_staar = startDatoAar

        case else
        
        dtlink_stdag = datepart("d", oRec("jobstartdato"), 2, 2)
		dtlink_stmd = datepart("m", oRec("jobstartdato"), 2, 2)
		dtlink_staar = datepart("yyyy", oRec("jobstartdato"), 2, 2)

		
        end select

        

		dtlink_sldag = datepart("d", now, 2, 2)
		dtlink_slmd = datepart("m", now, 2, 2)
		dtlink_slaar = datepart("yyyy", now, 2, 2)
		
       
		
		dtlink = "FM_usedatokri=1&FM_start_dag="&dtlink_stdag&""_
	    &"&FM_start_mrd="&dtlink_stmd&"&FM_start_aar="&dtlink_staar&"&FM_slut_dag="&dtlink_sldag&""_
	    &"&FM_slut_mrd="&dtlink_slmd&"&FM_slut_aar="&dtlink_slaar
         
          %>

          <br />
          <a href="joblog.asp?nomenu=1&FM_job=<%=oRec("id") %>&FM_kunde=<%=oRec("kid")%>&FM_jobsog=<%=oRec("jobnr")%>&viskunabnejob0=<%=viskunabnejob0%>&viskunabnejob1=<%=viskunabnejob1%>&viskunabnejob2=<%=viskunabnejob2%>&<%=dtlink %>" class=rmenu target="_blank"><%=job_txt_266&" " %>>></a><br />
          <a href="jobprintoverblik.asp?menu=job&id=<%=oRec("id")%>&media=printjoboverblik" class=rmenu target="_blank"><%=job_txt_267 %> >></a><br>
          <a href="timereg_akt_2006.asp?FM_kontakt=<%=oRec("kid")%>&FM_ignorer_projektgrupper=1&jobid=<%=oRec("id")%>&FM_sog_job_navn_nr=<%=oRec("jobnr")%>&usemrn=<%=session("mid")%>&showakt=1&fromsdsk=1" target="_blank" class=rmenu><%=job_txt_268 &" " %>>> </a><br />
          
          <%if cint(useasfak) <= 2 then %>
          <a href="../timereg/erp_opr_faktura_fs.asp?func=opr&visfaktura=1&visjobogaftaler=1&visminihistorik=1&FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&<%=dtlink %>" target="_blank" class=rmenu><%=job_txt_269 %> >> </a>
		  <%end if %>

		<%else %>
		&nbsp;
		<%end if %></td>
		<td class=lille valign=top style="padding-right:2px;">
		<div style="position:relative; top:0px, left:0px; width:160px; height:100px; overflow:auto; padding:5px;">
        <table cellspacing=0 border=0 cellpadding=0 width=100%>
		
		<%
		
		        '** Findes der fakturaer p� job kan det ikke slettes **'
			    
			    deleteok = 0
			    faktotbel = 0
                fakturaBel_tot = 0
                fakAftaleid = 0
			    strSQLffak = "SELECT f.fid, f.faknr, f.aftaleid, f.faktype, f.jobid, f.fakdato, f.beloeb, "_
			    &" f.faktype, f.kurs, brugfakdatolabel, labeldato, fakadr, shadowcopy, f.valuta, afsender FROM fakturaer f "_
			    &" WHERE (jobid = " & oRec("id") & " AND shadowcopy = 0) OR (aftaleid = "& oRec("serviceaft") &" AND aftaleid <> 0 AND shadowcopy = 0)"_
			    &" GROUP BY f.fid ORDER BY f.fakdato DESC"

                'OR aftid = "& oRec("serviceaft") &"
                '&" LEFT JOIN faktura_det AS fd ON (fd.fakid = f.fid)"_
                'AND fd.enhedsang <> 3
                'SUM(fd.aktpris) AS aktbel
                '&" AND aftaleid = 0 AND shadowcopy = 0"_


			    'Response.Write strSQLffak 
			    'Response.flush
			    
                f = 0
			    oRec3.open strSQLffak, oConn, 3
			    while not oRec3.EOF 
			        

                    if oRec3("brugfakdatolabel") = 1 then '** Labeldato
                    fakDato = job_txt_623&": "& oRec3("labeldato") 'VIS IKKE Faktura sys dato p� jioblsiten: Gnidret og manlger overblik & "<br><span style=""color:#999999; size:9px,"">"& oRec3("fakdato") &"</span>"
                    else
                    fakDato = job_txt_622&": "& oRec3("fakdato")
                    end if
			       

                              strFakAftNavn = ""

                             if oRec3("aftaleid") <> 0 then 'AND oRec2("aftaleid") <> 0 then 'oRec3("shadowcopy") = 1
               
                   

                                        strSQLFakorg = "SELECT f.fid, f.beloeb, f.valuta, f.kurs, f.faktype, f.aftaleid, fd.aktpris FROM fakturaer f "_
                                        &" LEFT JOIN faktura_det fd ON (fd.fakid = f.fid AND fd.aktid = "& oRec("id") &") WHERE fid = "& oRec3("fid") &" AND shadowcopy <> 1 "
                                        
                                        oRec8.open strSQLFakorg, oConn, 3
                                        if not oRec8.EOF then

                        
                                            fakValuta = oRec8("valuta")
                                            fakKurs = oRec8("kurs")
                                            fakbeloeb = oRec8("aktpris") 'oRec8("beloeb")
                                            fakType = oRec8("faktype")
                                            fakAftaleid = oRec8("aftaleid")
                                            fakAktbel = oRec8("aktpris")

                       
                                        end if
                                        oRec8.close

                                         
                                        if fakAftaleid <> "" then
                                        fakAftaleid = fakAftaleid
                                        else
                                        fakAftaleid = 0
                                        end if
                                        
            
                                        strSQLaftale = "SELECT navn, aftalenr FROM serviceaft WHERE id = " & fakAftaleid
                                        oRec8.open strSQLaftale, oConn, 3
                                        if not oRec8.EOF then
                                        
                                        strFakAftNavn = "<span style=""font-size:9px; color:#999999;""> - " & left(oRec8("navn"), 5) & " ("& oRec8("aftalenr") &")</span>" 
                                        
                                        end if
                                        oRec8.close



                            else
                        
                 
                        
                                fakValuta = oRec3("valuta")
                                fakKurs = oRec3("kurs")
                                fakBeloeb = oRec3("beloeb")
                                fakType = oRec3("faktype")
                                fakAktbel = 0 'oRec3("aktbel")

                            end if 'aftale fak

			          
                      %>
                      <tr>
                      <%
			        
                            'fidLink = 0
                            'if editok = 1 then 'cdate(oRec3("fakdato")) >= cdate("01-01-2006") AND
                            fidLink = oRec3("fid")    
                            'end if

                          
                          if cint(editok) = 1 then 'fidLink <> 0 then
                          %>
			              <td class=lille valign=top><a href="erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&id=<%=fidLink%>&FM_jobonoff=<%=FM_jobonoffval%>&FM_kunde=<%=oRec3("fakadr")%>&FM_job=<%=oRec3("jobid")%>&FM_aftale=<%=oRec3("aftaleid")%>&fromfakhist=1" class="lgron" target="_blank"><b><%=oRec3("faknr")%></b></a> 
                          <%=strFakAftNavn %>
                          </td>
                          <%else 
                          %><td class=lille valign=top><b><%=oRec3("faknr") &" "& strFakAftNavn %></b></td>
                          <%end if %>
                          
                         <td align=right class=lille valign=top style="white-space:nowrap;"><%=fakDato %></td>
                       
                      </tr>
                      <%
                    


                  
                     
	                call beregnValuta(minus&(fakBeloeb),fakKurs,100)
                    if fakType <> 1 then
                    belobGrundVal = valBelobBeregnet
                    else
                    belobGrundVal = -valBelobBeregnet
                    end if 
                    

                        if cDate(oRec3("fakdato")) < cDate("01-06-2010") AND (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk") then
                        belobKunTimerStk = belobKunTimerStk + oRec3("beloeb")
                        else
       
                    
	                        '** Kun aktiviteter timer, enh. stk. IKKE materialer og KM
	                        call beregnValuta(minus&(fakAktbel),fakKurs,100)
                            if fakType <> 1 then
                            belobKunTimerStk = valBelobBeregnet
                            else
                            belobKunTimerStk = -valBelobBeregnet
                            end if 


                     end if
            
                fakturaBel_tot = fakturaBel_tot + (belobGrundVal)
                faktotbel = faktotbel + belobKunTimerStk
			    f = f + 1
			    oRec3.movenext
			    wend
			    oRec3.close

                fakturaBel_tot_gt = fakturaBel_tot_gt + fakturaBel_tot
			    %>
                </table>
               

			    </div>
                 <%=job_txt_270 %>:<br /> 
                 <b><%=formatnumber(fakturaBel_tot, 2) & " "& basisValISO %></b>

                 <%if totReal <> 0 then 
	        gnstpris = faktotbel/totReal
	        else 
	        gnstpris = 0
	        end if%>
	        
            <%if gnstpris <> 0 then %>
            <br /><%=job_txt_271 %>:<br />
	        <b><%=formatnumber(gnstpris) & " "& basisValISO %></b>
            
            <!-- <span class="qmarkhelp" id="qm0001" style="font-size:11px; color:#999999; font-weight:bolder;">?</span><span id="qmarkhelptxt_qm0001" style="visibility:hidden; color:#999999; display:none; padding:3px; z-index:4000;">(faktureret bel�b - (materialeforbrug + km)) / timer realiseret</span>-->
	        
	        <%
            end if

	        if f <> 0 AND proaf <> 0 then
	        gnsPrisTot = gnsPrisTot + (faktotbel)
	        else
	        gnsPrisTot = gnsPrisTot
	        'totRealialt = totRealialt
	        end if

            totRealialt = totRealialt + (proaf)
            %>


         </td>
		<td align=right class=lille style="white-space:nowrap; padding:4px 4px 0px 4px;" valign=top>
        <% for p = 1 to 10
        
        pgid = oRec("projektgruppe"&p)

        if pgid <> 1 then
            call prgNavn(pgid, 200)
            Response.Write left(prgNavnTxt, 20) & "<br>"
        end if
        
        next %>
		   
		</td>

		<td valign=top style="padding:4px 2px 4px 10px;">
		<%
		    
		    if f = 0 AND editok = 1 then
		    deleteok = 1
		    else
		    deleteok = 0
		    end if
		    
		    
		if deleteok = 1 then%>
		<a href="jobs.asp?menu=job&func=slet&id=<%=oRec("id")%>&jobnr_sog=<%=jobnr_sog%>&filt=<%=filt%>&fm_kunde_sog=<%=fmkunde%>"><%=sletkanp%></a>&nbsp;
		<%else %>
		&nbsp;
		<%end if%>
		</td>
	</tr>
	<%
	cnt = cnt + 1
	x = 0
	lastKid = oRec("kid")

    select case right(cnt,1)
    case 0
    Response.flush
    end select

    end if' forretningsomr�de


    
	oRec.movenext
	wend
    oRec.Close
    
        %>
           <input type="hidden" id="jq_jids" value="<%=jobids_all %>" />
        <%

	
	if cnt > 0 then%>
    	<tr style="background-color:#FFFFFF;">
         <td colspan=8>&nbsp;</td>
		<td>
            <%if fakturaBel_tot_gt <> 0 then %>
            <br /><%=job_txt_326 %>: <br /><b><%=formatnumber(fakturaBel_tot_gt, 2) & " "& basisValISO %></b>
            <%end if %>

            <%if totRealialt <> 0 AND gnsPrisTot <> 0 then %>
            <br><br><%=job_txt_327 %>:<br><b><%=formatnumber(gnsPrisTot/totRealialt, 2) &" "& basisValISO  %></b> 
            <%end if %>

		</td>
             <td colspan=11>&nbsp;</td>
        </tr>
        <tr style="background-color:#FFFFFF;">
            <td colspan=20 style="padding-right:30px; padding-top:5px;" align=right><br /><input type="submit" name="statusliste" value="<%=job_txt_272 %> >>"><br />&nbsp;</td>
	    </tr>
    <%end if %>
	</table>
	</div>
	<!-- end table div -->



    
	<%if cnt = 0 AND request("fromvemenu") <> "j" then%>
	
    <br /><br /><br /><br />

         <table cellspacing="0" cellpadding="2" border="0" bgcolor="#ffffff" width=100%>
	<tr>
		<td style="padding-left:20;"><br>
		<b><%=job_txt_273 %></b><br>
		<%=job_txt_274 %>:<br>
			<ul>
			<li><%=job_txt_275 &" " %><u><%=job_txt_276 %></u>
			<li><%=job_txt_328 & " " %><u><%=job_txt_277 %></u> 
			<li><%=job_txt_278 &" " %><u><%=job_txt_279 %></u></ul>
			<%=job_txt_280 %>
		</td>
	</tr> <!-- 'her -->
	
    </table>
    <br /><br />&nbsp;

	<%
    'response.end
    end if%>	
	
	
    <%
    if cnt > 0 then
    
    tTop = 40
	tLeft = 0
	tWdth = 600
	
	
	call tableDiv(tTop,tLeft,tWdth)
	 %>

    <br />
    <h3><%=job_txt_247 %></h3>
	<table cellspacing=1 cellpadding=2 border=0 width=100%>
	

    <%if level = 1 then %>
     <tr style="background-color:#FFFFFF;"><td colspan=2>
 
         <b><%=job_txt_281 %>:</b><br />

                              
                                <%
                                
                                strSQLf = "SELECT f.id, f.navn, kp.navn AS kontonavn, kp.kontonr FROM fomr AS f LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) WHERE f.id <> 0 AND f.jobok = 1 ORDER BY f.navn"
                                
                                'response.write strSQLf
                                'response.flush
         
                                %>
                                <select name="FM_fomr" id="Select3" multiple="multiple" size="5" style="width:450px;">
                                <option value="#0#"><%=job_txt_129 &" " %>(<%=job_txt_282 %>)</option>
                               
                                    
                                    <%oRec3.open strSQLf, oConn, 3
                                    while not oRec3.EOF
                                    
                                    if oRec3("kontonr") <> 0 then
                                        kontnrTxt = " ("& oRec3("kontonavn") &" "& oRec3("kontonr") &")"
                                    else 
                                        kontnrTxt = ""
                                    end if
                                    
                                    %>
                                    <option value="#<%=oRec3("id") %>#"><%=oRec3("navn") & kontnrTxt %></option>
                                    <%
                                    oRec3.movenext
                                    wend
                                    oRec3.close
                                    %>
                                </select>

        <br />
         <input id="Checkbox3" type="checkbox" name="FM_formr_opdater" value="1" /> <%=job_txt_283 %><br />
          <input id="Checkbox3" type="checkbox" name="FM_formr_opdater_akt" value="1" /> <%=job_txt_284 %> 


    </td></tr>
    <%end if %>


    <%dtNow = day(now) & "-"& month(now) & "-"& year(now) %>

    <tr style="background-color:#FFFFFF;">
		<td style="padding-right:30px; padding-top:10px;">
        <b><%=job_txt_285 %>:</b><br />
        <input type="checkbox" name="FM_opdaterslutdato" value="1"><%=job_txt_286 %> = <input type="text" name="FM_opdaterslutdato_dato" value="<%=dtNow%>" style="width:75px;" /> dd-mm-yyyy</td>
	</tr>
        <%
       call showEasyreg_fn()     
       if cint(showEasyreg_val) = 1 then %>
      <tr style="background-color:#FFFFFF;">
		<td style="padding-right:30px; padding-top:10px;">
        <b><%=left(tsa_txt_358, 8) %>:</b><br />
        Tilf�j/Fjern alle aktiviteter p� ovenst�ende job til easyreg<br />
        <select name="FM_tilfojeasyreg">
            <option value="0" selected>G�r ingenting</option>
            <option value="1">Tilf�j</option>
            <option value="2">Fjern</option>
            </select>
            </td>
	</tr>
        <%end if %>

	<tr style="background-color:#FFFFFF;">
		<td style="padding-right:30px; padding-top:5px;" align=right><br /><input type="submit" name="statusliste" value="<%=job_txt_272 %> >>"></td>
	</tr></form>
    </table>

    	</table>
	</div>
	<!-- end table div -->

	<%end if%>
	

    <%
    'oRec.Close
    
    
    else 'vis_timepriser%>
    <br />
    <b><%=job_txt_287 %></b> <br />
    <%=job_txt_288 %>

    <table cellspacing="0" cellpadding="2" border="0" bgcolor="#ffffff" width=100%>
	
            <tr bgcolor="#5582D2">
                <td class="alt"><%=job_txt_289 %></td>
                <td class="alt"><%=job_txt_290 %></td>
                <td class="alt"><%=job_txt_291 %></td>
                <td class="alt"><%=job_txt_292 %></td>
                <td class="alt"><%=job_txt_293 %></td>
                <td class="alt"><%=job_txt_294 %></td>
                <td class="alt"><%=job_txt_295 %></td>
                <td class="alt"><%=job_txt_296 %></td>
                
            </tr>
            <%
            
            jids = 0

            if sorttp = 0 then
            sortTpBy = "a.fase, a.navn, m.mnavn"
            else
            sortTpBy = "m.mnavn, a.fase, a.navn"
            end if
              

                        strSQL = "SELECT tp.jobid, tp.aktid, tp.medarbid, tp.timeprisalt, tp.6timepris, tp.6valuta, j.jobnavn, j.jobnr, a.navn AS aktnavn, a.fase, m.mnavn, m.init, m.mnr, "_
                        &" v.valutakode, j.jobans1, j.jobans2, j.id, "_
                        &" k.kkundenavn, k.kkundenr FROM job AS j"_
                        &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                        &" LEFT JOIN timepriser AS tp ON (tp.jobid = j.id) "_
                        &" LEFT JOIN aktiviteter AS a ON (a.id = tp.aktid) "_
                        &" LEFT JOIN medarbejdere AS m ON (m.mid = tp.medarbid) "_
                        &" LEFT JOIN valutaer AS v ON (v.id = tp.6valuta) "_
                        &" WHERE "& varJobknrKri &" "& sogeKri &" "& jobansKri &" k.kid = j.jobknr " & varFilt &" "& datoKri &" "& replace(prjgrpSQLkri, "projektgruppe", "j.projektgruppe") &""_ 
                        &" AND j.jobnavn IS NOT NULL AND m.mnavn IS NOT NULL GROUP BY tp.jobid, tp.aktid, tp.medarbid ORDER BY k.kkundenavn, j.jobnavn, j.jobnr, "& sortTpBy &" LIMIT 0,10000"
						
                        'Response.write strSQL & "<br>"
						'Response.flush


                        oRec2.open strSQL, oConn, 3 
						    
                            cnt = 0
                            while not oRec2.EOF 


                             call forretningsomrJobId(oRec2("id"))


                            if cint(visJobFomr) = 1 OR instr(strFomr_rel, "#0#") <> 0 then


                            select case right(cnt,1)
                            case 0,2,4,6,8
                            bgthis = "#FFFFFF"
                            case else
                            bgthis = "#EFF3FF"
                            end select




						    'thissel = oRec2("timeprisalt")
						    this6timepris = formatnumber(oRec2("6timepris"), 2)
						    this6valuta = oRec2("valutakode")

                            
                            if IsNull(oRec2("aktnavn")) = False OR oRec2("aktid") = 0 then

                            %>
                            <tr bgcolor="<%=bgthis%>">
                                <td style="padding:3px; border-bottom:1px #999999 solid; white-space:nowrap;"><b><%=left(oRec2("kkundenavn"), 65) & "</b> ("& oRec2("kkundenr") &")" %></td>

                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;">

                                <%
                                editok = 0
	                            if level <= 2 OR level = 6 then
	                            editok = 1
	                            else
			                            if cint(session("mid")) = oRec2("jobans1") OR cint(session("mid")) = oRec2("jobans2") OR (cint(oRec2("jobans1")) = 0 AND cint(oRec2("jobans2")) = 0) then
			                            editok = 1
			                            else
			                            editok = 0   
			                            end if
	                            end if
	
	
	              
	
	
	                           
	                            if editok = 1 then
		                           %>
                                   <a href="jobs.asp?menu=job&func=red&id=<%=oRec2("id")%>&int=1&jobnr_sog=<%=jobnr_sog%>&filt=<%=filt%>&fm_kunde_sog=<%=fmkunde%>&showdiv=tpriser" class=vmenu><%=oRec2("jobnavn")%>&nbsp;&nbsp;(<%=oRec2("jobnr")%>)</a>&nbsp;
		                           <%
	                            else
		                            %>
		                            <b><%=oRec2("jobnavn")%>&nbsp;&nbsp;(<%=oRec2("jobnr")%>)</b>&nbsp;
		                            <%
	                            end if
                                %>


                                </td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=oRec2("fase") %>&nbsp;</td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=oRec2("aktnavn") & " <span style=""font-size:8px; color:#cccccc;"">"& oRec2("aktid") %></span>&nbsp;</td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=oRec2("mnavn") & " ("& oRec2("mnr") &")" %></td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=oRec2("init")  %>&nbsp;</td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=this6timepris %></td>
                                <td style="border-bottom:1px #999999 solid; white-space:nowrap;"><%=this6valuta %></td>
                            </tr>
                          
                            <%
                           
                            select case right(cnt,1)
                            case 0
                            Response.flush
                            case else
                            
                            end select
                           
                            if lastJob <> oRec2("jobid") then
                            jids = jids & "," & oRec2("jobid")
                            lastJob = oRec2("jobid")
                            end if

                            cnt = cnt + 1


                            end if 'Is null (akt slettet)

                            end if 'Forretningsomr

                            oRec2.movenext
						    wend 
						oRec2.close 

    
            
            
            %>
            </table>
	</div>
	<!-- end table div -->



    <%end if %>
	
	
	
	
	

    <%  

    'if cint(cnt) <> 0 then


ptop = 0
pleft = 946
pwdt = 210

call eksportogprint(ptop,pleft, pwdt)
%>

<form action="job_eksport.asp" method="post" target="_blank">
<tr> <input id="Hidden3" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit5" type="submit" value="A) <%=job_txt_329 %> >> " style="font-size:9px; width:130px;" />
    <input type="hidden" value="1" name="eksDataStd" /><br />
    <input type="checkbox" value="1" name="xeksDataStd" checked disabled /> <%=job_txt_297 %><br />
    <input type="checkbox" value="1" name="eksDataNrl" /> <%=job_txt_298 %><br />
    <input type="checkbox" value="1" name="eksDataJsv" /> 
        <%
        call salgsans_fn()    
        if cint(showSalgsAnv) = 1 then  %>
        <%=job_txt_300 %>
        <%else %>
        <%=job_txt_330 %>
        <%end if %><br />
    <input type="checkbox" value="1" name="eksDataAkt" /> <%=job_txt_232 %><br />
    <input type="checkbox" value="1" name="eksDataMile" /> <%=job_txt_299 %>
         <%if jobasnvigv = 1 then %>
              (<%=job_txt_301 %>)
             <%end if %>
    </td>
</tr>
</form>


<form action="job_eksport.asp?optiprint=3" method="post" target="_blank">
<tr> <input id="Hidden5" name="jids" value="<%=jids%>" type="hidden" />
<input id="Hidden6" name="sorttp" value="<%=sorttp%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit3" type="submit" value="B) <%=job_txt_302 %> >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>




<!--
<form action="job_eksport.asp?optiprint=4" method="post" target="_blank">
<tr> <input id="Hidden6" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit4" type="submit" value="C) Eksport�r jobansvarlige >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>

<form action="job_eksport.asp?optiprint=5" method="post" target="_blank">
<tr> <input id="Hidden7" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit7" type="submit" value="D) Eksport�r job og akt. >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>

-->

<form action="job_eksport.asp?optiprint=1" method="post" target="_blank">
<tr>
<input id="Hidden4" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/printer3.png" />
    </td><td class=lille><input id="Submit6" type="submit" value="E) <%=job_txt_303 %> >>" style="font-size:9px; width:130px;" /></td>
</tr>
</form>


<%if instr(lto, "epi") <> 0 OR lto = "intranet - local" then %>
<form action="job_eksport.asp?optiprint=6" method="post" target="_blank">
<tr> <input id="Hidden5" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
    <td class=lille><input id="Submit7" type="submit" value="F) <%=job_txt_304 %> >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>

<%end if %>

<%if lto = "dencker" OR lto = "dencker_test" OR lto = "intranet - local" OR lto = "cflow" then %>
<form action="job_eksport.asp?optiprint=7" method="post" target="_blank">
<tr> <input id="Hidden5" name="jids" value="<%=jids%>" type="hidden" />
    <td valign=top align=center>
   <input type=image src="../ill/export1.png" />
    </td>
   <td>

        <input id="Submit7" type="submit" value="F) <%=job_txt_306 %> >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>

<!--  <td class=lille>
         d.d -
        <select name="antaldage" style="font-size:9px;">
            <%for a = 0 TO 10 %>
            <option value="<%=a %>"> <%=a %> <%=" "& job_txt_305 %></option>
            <%next %>
        </select>
        
        </td>
    </tr>
<tr><td>&nbsp;</td>
 -->

<%end if %>
	
	
	
    <!--
	<br /><br /><a href="job_eksport.asp?jids=<%=jids%>" class=vmenu target="_blank">&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>&nbsp;
	<br><a href="job_eksport.asp?jids=<%=jids%>&optiprint=1" class=vmenu target="_blank">Eksporter ovenst�ende liste af job. (optimeret til print)&nbsp;<img src="../ill/soeg-knap.gif" width="16" height="16" alt="" border="0"></a>&nbsp;
	<br><br>
    -->
	
	
	
	

	<%
	
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus <> 3"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksterneJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksterneJob = antalEksterneJob
	else
	antalEksterneJob = 0
	end if
	%>
	
	<!--<b><=antalEksterneJob%></b> Eksterne job.
	<br>-->
	
	
	<%
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 1"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksterneAktiveJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksterneAktiveJob = antalEksterneAktiveJob
	else
	antalEksterneAktiveJob = 0
	end if
	
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 0"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksterneLukkedeJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksterneLukkedeJob = antalEksterneLukkedeJob
	else
	antalEksterneLukkedeJob = 0
	end if
	
	
	'*** Tilbud ***'
	strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 3"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	antalEksternePassiveJob = oRec("antal")
	end if
	oRec.close 
	
	if len(antalEksterneJob) <> 0 then
	antalEksternePassiveJob = antalEksternePassiveJob
	else
	antalEksternePassiveJob = 0
	end if
	
	antalTilbud = antalEksternePassiveJob
	

    '*** Gennemsyn ***'
	gennemsyn = 0
    strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 4"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	gennemsyn = oRec("antal")
	end if
	oRec.close 
	
	if len(gennemsyn) <> 0 then
	gennemsyn = gennemsyn
	else
	gennemsyn = 0
	end if


    '*** Til fakturering ***'
	tilfakturering = 0
    strSQL = "SELECT count(id) AS antal FROM job WHERE fakturerbart = 1 AND jobstatus = 2"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	tilfakturering = oRec("antal")
	end if
	oRec.close 
	
	if len(tilfakturering) <> 0 then
	tilfakturering = tilfakturering
	else
	tilfakturering = 0
	end if
	
	
       

        %><tr><td colspan="2"><br /><br />
            <%
                nWdt = 120
                nTxt = job_txt_307
                nLnk = "jobs.asp?menu=job&func=opret&id=0&int=1"
                nTgt = ""
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>
          
                <br />
            
            <%if lto = "outz" OR lto = "intranet - local" then %>

            <%
                nWdt = 120
                nTxt = job_txt_308
                nLnk = "createJobFromFile.aspx?lto="&lto&"&editor="&session("user")
                nTgt = "_blank"
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>

                <br /><br />
                 <%end if
	
        
	
	if totRealialt < 1 then
	totRealialt = 1
	end if
	
	
	uTxt = "<b>"&job_txt_311&":</b><br><b>"& antalEksterneAktiveJob & "</b> "&jobstatus_txt_001 _
	&"<br><b>" & antalEksterneLukkedeJob &"</b> "&jobstatus_txt_005 _
	&"<br><b>" & antalTilbud & "</b> "& jobstatus_txt_003 _
    &"<br><b>" & gennemsyn & "</b> "& jobstatus_txt_004 _
    &"<br><b>" & tilfakturering & "</b> "& jobstatus_txt_002
	
    if cint(vis_timepriser) <> 1 then
    uTxt = uTxt  &"<br><b>" & cnt & "</b> "& job_txt_310
	else
    uTxt = uTxt  &"<br><b>" & cnt & "</b> "& job_txt_313
    end if

    'uTxt = uTxt  &"<br><br>Gns. faktisk timepris:<br><b> "& formatnumber(gnsPrisTot/totRealialt) &" "& basisValISO &" </b>"
	'&" Gns. faktisk timepris = (faktureret bel�b ekskl. mat. og km. / real. timer)<br>"_
	'&" G�lder for de viste job, hvor der forefindes fakturaer og realiserede timer.<br>Gns. timepris er v�gtet i forhold til timeforbrug p� de enkelte job."
	
                Response.write uTxt


	%>

	</td></tr>
</table></div>	
<%'end if %>




	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>

	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>


<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br>
	<br>
	<br><br>
	<br><br>
	<br><br>
	<br><br>
	<br>
&nbsp;





	<a href="Javascript:history.back()" class="vmenu"><< <%=job_txt_314 %> </a>
	<br>
	<br>
	</div>
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->

