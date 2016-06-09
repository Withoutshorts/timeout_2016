<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" -->



<%
'** Bruges til PDF visning ***
if len(request("nosession")) <> 0 then
nosession = request("nosession")
else
nosession = 0
end if


if Request.Form("AjaxUpdateField") = "true" then
Select Case Request.Form("control")
case "FN_getCustDesc"

strSQL = "SELECT ts_txt FROM tilbuds_skabeloner WHERE ts_id = " & Request.Form("cust")
oRec.open strSQL, oConn, 3
while not oRec.EOF
Response.Write oRec("ts_txt") 
oRec.movenext
wend
oRec.close

End select
Response.End
end if








if len(trim(session("user"))) = 0 AND cint(nosession) = 0  then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	id = request("id")
	
	if len(trim(request("kid"))) <> 0 then
	kid = request("kid")
	else
	kid = 0
	end if
	
	thisfile = "job_print.asp"
	
	if len(trim(request("hideudskriftfunc"))) <> 0 then
	    hideuf = request("hideudskriftfunc")
	    else
	    hideuf = 0
	    end if
	
	if func <> "print" then%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%else %>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%end if %>
	
	
	
	<script>

	    $(document).ready(function() {

	        $("#FM_vis_akt").click(function() {

	            var n = $("#FM_vis_akt:checked").length;

	            if (n == 1) {
	                $(".chkakt").attr("checked", "checked")
	                //$("#af_95").checked = false;
	                //document.getElementById("af_95").checked = false
	            } else {
	                $(".chkakt").removeAttr("checked");

	            }



	        });


	        $(".faseoskrift").click(function() {

	            var thisid = this.id
	            //alert(thisid)
	            var n = $("#" + thisid + ":checked").length;
	            //alert(n)

	            for (i = 0; i < 100; i++) {
	                if (n == 1) {
	                    $("#af_" + thisid + "_" + i).attr("checked", "checked")
	                } else {
	                    $("#af_" + thisid + "_" + i).removeAttr("checked");

	                }
	            }


	        });



	        //$("#BtnCustDescUpdate").click(function() {
	        //    $(this).hide();
	        //    $("#LoadCustDescUpdate").show();
	        //    $.post("?", { AjaxUpdateField: "true", id: $("#BtnCustDescUpdate").data("cust"), control: "FM_CustDesc", value: $("#custDesc").val() }, function() {
	        //        TONotifie("Opdateringen i kundebeskrivelsen blev gemt", true);
	        //        $("#BtnCustDescUpdate").show();
	        //        $("#LoadCustDescUpdate").hide();
	        //    });
	        //});

	        function GetCustDesc() {
	            $("#custDesc").val("henter text...");
	            var thisC = $("#tb_skabeloner")
	            $("#BtnCustDescUpdate").data("cust", thisC.val());
	            $.post("?", { control: "FN_getCustDesc", AjaxUpdateField: "true", cust: thisC.val() }, function(data) { $("#custDesc").val(data); });




	        }
	        GetCustDesc();
	        $("#tb_skabeloner").change(function() { GetCustDesc(), showSkabtxt(); });
	        //, showSkabtxt()
	        //showSkabtxt() 

	    });

	    function showSkabtxt() {

	        document.getElementById("sletskabelonid").value = document.getElementById("tb_skabeloner").value
	                   
			
	        //alert(document.getElementById("custDesc").value)
	        //document.getElementById("FM_vis_indledning").value = document.getElementById("custDesc").value
	        // get the cute editor instance
	        //var editorK = document.getElementById('FM_vis_indledning');

	        // setting the content of Cute Editor
	        // setting the content of Cute Editor
	        //editorK.setHTML("Hello World");
	         
	        
	        
	    }
        
	 
	</script>
	<%
	
	if func = "sletskabelon" then
	
	ts_id = request("sletskabelonid")
	strSQLdelts = "DELETE FROM tilbuds_skabeloner WHERE ts_id = " & ts_id
	oConn.execute(strSQLdelts)
	
	Response.Redirect "job_print.asp?id="&id
	
	
	end if
	
	
	if func = "print" AND hideuf = 0 then
	
	'*** Opdaterer felter ***'
	'jobbeskupd = replace(request("FM_jobbesk"), "'", "''")
	'strSQLu = "UPDATE job SET beskrivelse = '"& jobbeskupd &"' WHERE id = " & id
	'oConn.execute(strSQLu) 
	
	
	'*** Gemmer skabelon ****'
	if request("gemsomskabelon") = "1" then
	
	skabelonNavn = replace(request("skabelonnavn"), "'", "''")
	skabelonTxt = replace(request("FM_vis_indledning"), "'", "''")
	tsDato = year(now) & "/"& month(now) &"/"& day(now)
	tsEditor = session("user")
	ts_kundeid = request("gemsomskabelon_kid")
	
	strSQLinsSkab = "INSERT INTO tilbuds_skabeloner (ts_navn, ts_txt, ts_dato, ts_editor, ts_kundeid) VALUES "_
	&" ('"& skabelonNavn& "', '"& skabelonTxt &"', '"& tsDato &"', '"& tsEditor &"', "& ts_kundeid &")"
	
	oConn.execute(strSQLinsSkab)
	
	end if
	
	end if
	
	
	if func <> "print" then
	lft = 20
	tp = 0
	
	'*** Henter det job der skal printes ***
		strSQL = "SELECT id, jobnavn, jobnr,"_
		&" k.kkundenavn, k.kid, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, k.land, "_
		&" jobknr, "_
		&" jobTpris, jobstatus, jobstartdato, jobslutdato, "_
		&" job.dato, job.editor, tilbudsnr, "_
		&" fakturerbart, budgettimer, fastpris, kundeok, job.beskrivelse, "_
		&" ikkeBudgettimer, tilbudsnr, serviceaft, kundekpers, jobans1, jobans2 "_
		&" FROM job "_
		&" LEFT JOIN kunder k ON (k.Kid = jobknr) "_
		&" WHERE id = " & id 
		oRec.open strSQL, oConn, 3
				
		
		if not oRec.EOF then	
				
			kid = oRec("kid")	
			strNavn = oRec("jobnavn")
			strjobnr = oRec("jobnr")
			tilbudsnr = oRec("tilbudsnr")
			
			strKnavn = oRec("kkundenavn")
			strKundeId = oRec("jobknr")
			strKundenr = oRec("kkundenr")
			strAdresse = oRec("adresse")
			strPostnr = oRec("postnr")
			strBy = oRec("city")
			strTLF = oRec("telefon")
			strLand = oRec("land")
			
			strBudget = oRec("jobTpris")
			if strBudget > 0 then
			strBudget = strBudget
			else
			strBudget = 0
			end if
			
			if oRec("ikkeBudgettimer") > 0 then
			ikkeBudgettimer = oRec("ikkeBudgettimer")
			else
			ikkeBudgettimer = 0
			end if
			
			if oRec("budgettimer") > 0 then
			strBudgettimer = oRec("budgettimer")
			else
			strBudgettimer = 0
			end if
			
			strTdato = oRec("jobstartdato")
			strUdato = oRec("jobslutdato")
			
			strFastpris = oRec("fastpris")
			
			strBesk = oRec("beskrivelse")
			'strtilbudsnr = oRec("tilbudsnr")
			
			'if request("FM_aftaler") = 1 then
			'intServiceaft = oRec("serviceaft")
			'else
			'intServiceaft = 0
			'end if
			
			intkundekpers = oRec("kundekpers")
			
				strSQL2 = "SELECT mnavn AS jobans_mnavn1, email, init, mnr FROM "_
				&" medarbejdere WHERE mid = "& oRec("jobans1") 
				oRec2.open strSQL2, oConn, 3 
				while not oRec2.EOF 
				
				jobans1 = oRec2("jobans_mnavn1") &" ("& oRec2("mnr") &") "
				
				if len(oRec2("init")) <> 0 then
				jobans1 = jobans1 & " - " & oRec2("init")
				end if
				
				if len(oRec2("email")) <> 0 then
				jobans1 = jobans1 & ", " & oRec2("email")
				end if
				
			    oRec2.movenext
				wend
				oRec2.close
				
				strSQL3 = "SELECT mnavn AS jobans_mnavn2, email, init, mnr FROM "_
				&" medarbejdere WHERE mid =  "& oRec("jobans2")
				oRec2.open strSQL3, oConn, 3 
				while not oRec2.EOF 
				
				jobans2 = oRec2("jobans_mnavn2") &" ("& oRec2("mnr") &") "
				
				if len(oRec2("init")) <> 0 then
				jobans2 = jobans2 & " - " & oRec2("init")
				end if
				
				if len(oRec2("email")) <> 0 then
				jobans2 = jobans2 & ", " & oRec2("email")
				end if
				
				oRec2.movenext
				wend
				oRec2.close
				 
			
		end if
		oRec.close
    
    else
		lft = 20
	    tp = 100
		indledningVal = request("FM_vis_indledning")
	end if '** Print%>
	
	
	
	
	<div id="sindhold" style="position:absolute; left:<%=lft%>; top:<%=tp%>; visibility:visible; z-index:50; width:800px;">
	    
	    
	    <%if func = "print" AND hideuf = 0 then %>
	    
	     <table cellpadding=0 cellspacing=0 border=0 width=80>
	<tr>
	<td valign=top style="padding:20px 10px 0px 0px;"><a href="#" onclick="Javascript:history.back()"><img src="../ill/nav_left_blue.png" border="0" /></a></td>
    </tr>
	</table>
	    
	    <%end if %>
	    
	    <%if func <> "print" then %>
	    
	   
	    
	    
	    
	    <table cellspacing=0 cellpadding=0 border=0 width="100%">
	    <tr>
	        <td align=right valign=bottom><br /><br />
	         <form action="job_print.asp?func=sletskabelon&id=<%=id %>&kid=<%=kid %>" method="post">
        <input id="sletskabelonid" name="sletskabelonid" value="0" type="hidden" />
        <input id="Submit3" type="submit" value=" Slet Skabelon >> " />
	    </form>
	        </td>
	    </tr>
	    <form action="job_print.asp?id=<%=id%>&func=print&kid=<%=kid %>" method="post">
	    <tr><td valign=top>
	  
	
	<%
		strSQLlogo = "SELECT f.filnavn, k.kid, k.logo, k.kkundenavn, k.kkundenr, k.adresse, "_
		&" k.postnr, k.city, k.telefon, k.land, k.fax, k.cvr, k.bank, k.regnr, k.kontonr FROM kunder k "_
		&" LEFT JOIN filer f ON (f.id = k.logo) WHERE k.useasfak = 1"
		
		'Response.Write strSQLlogo
		'Response.flush
		
		oRec.open strSQLlogo, oConn, 3
		if not oRec.EOF then
		
		afsKnavn = oRec("kkundenavn")
		afsAdr = oRec("adresse")
		afsPostnr = oRec("postnr")
		afsLand = oRec("land")
		afsBy = oRec("city")
		afsTlf = oRec("telefon")
		afsFax = oRec("fax")
		
		afsBank = oRec("bank")
		afsRegnr = oRec("regnr")
		afsKontonr = oRec("kontonr")
		afsCVR = oRec("cvr")
				
		end if
		oRec.close
		
		strkundeHTML = "<br><table cellspacing=0 cellpadding=0 border=0 width=""100%""><tr><td valign=top>"
		
		'** Kunde / Modtager **'
		
		strkundeHTML = strkundeHTML & strKnavn &"<br>"& strAdresse & "<br>"_
		& strPostnr &" "& strBy & "<br>" & strLand & "<br><br>"  
		
		strSQLkpers = "SELECT navn FROM kontaktpers WHERE id = "& intkundekpers 
	    oRec.open strSQLkpers, oConn, 3
	    while not oRec.EOF
	    
	    strkundeHTML = strkundeHTML & "Att:"& oRec("navn") & "<br><br>"
	    
	    oRec.movenext
	    wend
	    oRec.close 
	    
	    
	    strkundeHTML = strkundeHTML & "</td><td valign=bottom align=right>"
	    strkundeHTML = strkundeHTML & afsBy &", d. "& formatdatetime(date(),1) &"</td></tr></table>"
	    
	   
        strjobHTML = "<br><br><table cellspacing=0 cellpadding=0 border=0 width=""100%""><tr><td valign=top>"
		
	    %>
	    
	    
	    
	    <%if tilbudsnr <> 0 then%>
	    <h4>Tilbud: <%=tilbudsnr %></h4>
		<%
		strJobHTML = strjobHTML & "<h4>Tilbud: "& tilbudsnr &"</h4>"
		
		end if %>
		
		
		<%
		'*** Jobnavn / vedr. text ***
		vedrtxtval = "<b>Vedr. job: "
		strJobHTML = strJobHTML & vedrtxtval &" "& strNavn & " ("&strjobnr&")</b><br>"
		
		if len(trim(jobans1)) <> 0 then
		strJobHTML = strJobHTML & "Jobansvarlig / vor. ref.: " & jobans1 & "<br>"
        end if  
        
        if len(trim(jobans2)) <> 0 AND prodtilbud = 2 then
		strJobHTML = strJobHTML & "Jobejer:" & jobans2 & "<br>"
        end if   
		
		strJobHTML = strJobHTML & "<br><b>Beskrivelse</b><br>"& strBesk
		
		strJobHTML = strJobHTML & "<br><h4>Pris: DKK "& formatnumber(strBudget, 2) &"</h4>"
		
		if prodtilbud = 2 then
		strJobHTML = strJobHTML & "Arbejdet forventes udført i perioden "& formatdatetime(strTdato, 1) &" - "& formatdatetime(strUdato, 1) & "<br>"
		end if
		
		      
                  
        strJobHTML = strJobHTML & "</td></tr></table>"
        
        
        
        strAktHTML = "<br><br><h4>Udspecificering:</h4><table cellspacing=0 cellpadding=0 border=0 width=""100%"">"
		
		
	
			strSQLselAkt = "SELECT id AS aktid, navn, beskrivelse, job, fakturerbar,  "_
			&" aktstartdato, aktslutdato, editor, dato, budgettimer, aktfavorit, aktstatus, fomr, faktor, aktbudget, fase "_
			&" FROM aktiviteter WHERE job =" & id &" ORDER BY fase, navn"
			
			oRec.open strSQLselAkt, oConn, 3
			
			af = 0
			lastFase = ""
			while not oRec.EOF 
				
				strNavn = oRec("navn")
				strBeskrivelse = oRec("beskrivelse")
				strTdato = oRec("aktstartdato")
				strUdato = oRec("aktslutdato")
				strFakturerbart = oRec("fakturerbar")
				useBudgettimer = oRec("budgettimer")
				intaktstatus = oRec("aktstatus")
			    aktbudget = oRec("aktbudget")
			    strFase = oRec("fase")
				
				
				if lcase(lastFase) <> lcase(strFase) then
				strAktHTML = strAktHTML & "<tr>"
				strAktHTML = strAktHTML & "<td colspan=5><br><br><b>"&strFase&"</b></td></tr>"
				    
				af = 0
				end if 
				
				
				
				strAktHTML = strAktHTML & "<tr><td style=""padding-right:20px; width:400px;"">"&strNavn&""
				
				if len(trim(strBeskrivelse)) <> 0 then
				strAktHTML = strAktHTML & "<br />"& strBeskrivelse 
				end if
					
				strAktHTML = strAktHTML & "</td>"	
					
				if prodtilbud = 2 then	
				strAktHTML = strAktHTML & "<td class=lille>"& formatdatetime(strTdato, 2) & " til<br> "& formatdatetime(strUdato, 2) &"</td>"       
				strAktHTML = strAktHTML & "<td align=right>"& formatnumber(useBudgettimer, 2) &"</td>"        
				end if	
				
				strAktHTML = strAktHTML & "<td align=right>DKK "& formatnumber(aktbudget, 2) &"</td></tr>"        
					
					     
			
			af = af + 1	
			lastFase = strFase	
			oRec.movenext
			wend
			oRec.close
							
			strAktHTML = strAktHTML & "</table>"
			
	    if prodtilbud = 2 then
		strMileHTML = "<br><br><h4>Milepæle:</h4><table cellspacing=0 cellpadding=0 border=0 width=""100%"">"
		
		        strSQL = "SELECT navn, milepal_dato, jid, beskrivelse FROM milepale WHERE jid = "& id & " ORDER BY milepal_dato"
		        oRec.open strSQL, oConn, 3 
		        x = 0
		        while not oRec.EOF 
		        strMileHTML = strMileHTML &"<tr><td valign=top>"& formatdatetime(oRec("milepal_dato"), 2) & ": "& oRec("navn") & "<br>"& oRec("beskrivelse") & "<br><br></td></tr>"
		        
		        x = x + 1
		        oRec.movenext
		        wend
		        oRec.close 
		      
		      strMileHTML = strMileHTML &"</table>"
		
		
		end if
		
		        
                    %>
                    <br /><br />
                    <h4>Tilbudstekst / Skabelon:</h4>
                    <table cellspacing=2 cellpadding=2 border=0 width=800>
                    <tr><td valign=top><b>Vælg skabelon / revider tilbud</b> 
                    
                    
                    <%
                    strSQL_skab = "SELECT ts_navn, ts_id, ts_dato, ts_txt FROM tilbuds_skabeloner WHERE ts_id <> 0 ORDER BY ts_navn"
                    
                    %>
                    <br />
                    <select id="tb_skabeloner" name="tb_skabeloner" size=10 style="width:300px;">
                    <option value="0">Vælg..</option>
                     
                    
                    <%
                    oRec4.open strSQL_skab, oConn, 3
                    while not oRec4.EOF 
                    %>
                    <option value="<%=oRec4("ts_id") %>"><%=oRec4("ts_navn") %> (<%=oRec4("ts_dato") %>)</option>
                    
                    
                       
                    <%
                    oRec4.movenext
                    wend
                    oRec4.close
                       %>
                        </select>
                        </td>
                        <td valign=top>
                        Kopier denne HTMLtekst, og paste den ind i editoren nedenfor i HTML view.<br />
                        <textarea cols="70" rows="8" style="width:478px;" id="custDesc"></textarea><br />
		                </td>
                          </tr>
                          </table>  
                          
                          
                          
                          
                          
                          
                          <br /><br /><br /><br />                                    
                        <h4>Job data og aktiviteter hentet fra det aktuelle job / tilbud:</h4>
                    <%
		                dim content
	                    content = strkundeHTML & strJobHTML & strAktHTML & strMileHTML
            			
			            
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_vis_indledning"
			            editorK.Text = content
			            editorK.FilesPath = "CuteEditor_Files"
			            editorK.AutoConfigure = "Simple"
            			
			            editorK.Width = 790
			            editorK.Height = 680
			            editorK.Draw()
		                %>
                    
                    <br />
                    
                    
                     
                    <input id="gemsomskabelon_kid" name="gemsomskabelon_kid" value="<%=kid %>" type="hidden" />
                    
                    <input id="gemsomskabelon" name="gemsomskabelon" type="checkbox" value="1" /> Gem tilbud som skabelon / revision, angiv navn:
                    <input name="skabelonnavn" id="skabelonnavn" style="width:200px;" type="text" /> &nbsp;&nbsp;<input id="Submit2" type="submit" value=" Gem / Se >> " />
                    
        <%else %>
        
            <%if len(indledningVal) <> 0 then %>
            <%=indledningVal %>
            <%end if %>
        
        <%end if %> 
                  
                  
           
		
		
		<%
		
	    if showfiler = 99999 then
		
		if len(trim(request("FM_vis_filer"))) <> 0 AND request("FM_vis_filer") <> 0 then
		visfiler = 1
		else
		visfiler = 0
		end if
		
		if cint(visfiler) = 1 OR func <> "print" then %>
		                <br><br>
		                <h4><input id="FM_vis_filer" name="FM_vis_filer" type="checkbox" value="1" />Filer:</h4>
		                <%
                		
                		
	                        jobidSQL = "AND fo.jobid = " & id
	                        kundeseSQL = ""
	                        kundeseFilSQL = ""
	                        gamleFilerKri = ""
	                        kundeid = strKundeId
                	
                	
                	
		                strSQL = "SELECT fo.kundeid AS kundeid, fo.navn AS foldernavn, "_
		                &" fi.adg_kunde, fi.adg_admin, fi.adg_alle, fo.id AS foid, fo.kundese, "_
		                &" fo.jobid AS jobid, filnavn, fi.id AS fiid, fi.dato, jobnr, jobnavn, jobans1, jobans2, fo.dato AS fodato"_
		                &" FROM foldere fo "_
		                &" LEFT JOIN filer fi ON (fi.folderid = fo.id "& kundeseFilSQL &") "_
		                &" LEFT JOIN job j ON (j.id = fo.jobid) "_
		                &" WHERE "& gamleFilerKri &" fo.kundeid = "& kundeid &" "& jobidSQL &" "& kundeseSQL &" ORDER BY foid, filnavn"
                		
		                'Response.write strSQL
		                'Response.flush
                		
		                oRec.open strSQL, oConn, 3 
		                x = 0
		                lastfoid = 0
		                while not oRec.EOF 
			                if lastfoid <> oRec("foid") then%>
			                <br><b><%=oRec("foldernavn")%></b>
			                <%end if%>
		                &nbsp;<a href="../inc/upload/<%=lto%>/<%=oRec("filnavn")%>" class='vmenulille' target="_blank"><%=oRec("filnavn")%></a><br>
		                <%
		                lastfoid = oRec("foid")
		                x = x + 1
		                oRec.movenext
		                wend
		                oRec.close 
		               
		               
                        
        end if
        end if 'showfiler
   
		'**** afslutning ****'
		
		
		if len(trim(request("FM_vis_afslutning"))) <> 0 then
		afslutningVal = trim(request("FM_vis_afslutning"))
		else
		    
		    if request.Cookies("job")("afslutning") <> "" then
		    afslutningVal = request.Cookies("job")("afslutning")
		    else
		    afslutningVal = ""
		    end if
		
		
		end if
		
		
		if len(request("FM_gem_afs")) <> 0 then
		gemafs = 1
		else
		gemafs = 0
		end if
		
		if gemafs = 1 then
		response.Cookies("job")("afslutning") = afslutningVal
		end if
		
		
		        
                    if func <> "print" then%>
                    <br /><br /><b>Afslutning:</b> (Med venlig hilsen...)&nbsp;
                    <br />
                    <textarea id="FM_vis_afslutning" name="FM_vis_afslutning" cols="93" rows="4"><%=afslutningVal %></textarea>
                    <br /><input id="FM_gem_afs" name="FM_gem_afs" type="checkbox" value="1" />gem som cookie, alle kunder.<br />
		            <%else %>
		            
		                <%if len(afslutningVal) <> 0 then %>
                        <table width=100% cellpadding=2 cellspacing=2 border=0><tr><td><br /><br /><%=afslutningVal %></td></tr></table>
                        <%end if %>
                    
                    <%end if %> 
                    
                    <br /><br /><br /><br /><br /><br />&nbsp;
	
	
		
	
	
	<%if func <> "print" then %>
	</td></tr>
	<tr>
	<td align=right style="padding:30px 0px 50px 0px;"> <input id="Submit1" type="submit" value="Gem / Se >>" /></td>
	</tr>
		</form>
	</table>
	<%end if %>
	
	
	
	
	
	
	

	</div><!-- side div -->
	
	<%
	    
	
	    if func = "print" AND cint(hideuf) < 1 then %>
		<div id=funktioner style="position:absolute; z-index:200; top:30px; width:300px; left:600px; border:2px #8caae6 DASHED; background-color:pink; padding:5px 5px 5px 5px;">
		<b>Udskrift funktioner:</b> <br />
		<form action="job_print.asp?func=print&id=<%=id%>&hideudskriftfunc=1&kid=<%=kid %>" method="post">
        <input id="FM_vis_indledning" name="FM_vis_indledning" value="" type="hidden" />
        <input id="Submit4" type="submit" value="Print" style="font-size:9px;" />
		</form>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<a href="job_make_pdf.asp?lto=<%=lto%>&id=<%=id%>&hideudskriftfunc=2&kid=<%=kid %>" class=vmenu target="_blank">PDF</a><br /><br />
		
	
		<b>Gem PDF ovenfor, og email faktura til:</b><br />
		<%
		strSQL = "SELECT navn, email FROM kontaktpers WHERE kundeid = " & kid
		'Response.Write strSQL 
		'Response.Flush
		
		oRec.open strSQL, oConn, 3
        while not oRec.EOF
        
        Response.Write "<i>"& oRec("navn") & "</i>, <a href='mailto:"&oRec("email")&"&subject=Vedr.: "& strjobnr&"_"&tilbudsnr &"' class=vmenu>" & oRec("email") & "</a><br>"
        
        oRec.movenext
        wend
        oRec.close
        
        %>
        <br />Kontaktpersoner oprettes under fanebladet "kontakter" i hovemenuen.
		</div>
		<!--&nbsp;&nbsp;|&nbsp;&nbsp;Excel.-->
	    <%end if
	    
	    
	    
	    if func = "print" AND cint(hideuf) = 1 then
	    
	    Response.Write("<script language=""JavaScript"">window.print();</script>")
	    
	    end if
	
	        
	       
	
	
	response.Cookies("job").expires = date + 31
	
	
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
