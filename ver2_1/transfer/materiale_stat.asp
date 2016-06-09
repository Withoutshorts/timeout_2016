<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/xml/tsa_xml_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	
	if len(request("hidemenu")) <> 0 then
	hidemenu = request("hidemenu")
	else
	hidemenu = 0
	end if
	
	thisfile = "materiale_stat.asp"
	func = request("func")
	id = request("id") '** Jobid
	print = request("print")
	
	if len(id) <>  0 then
	id = id
		else
			if len(request.cookies("matstat")("jobidmat")) <> 0 then
			id = request.cookies("matstat")("jobidmat")
			else
			id = 0
			end if
	end if
	Response.cookies("matstat")("jobidmat") = id
	
	session.lcid = 1030 'DK
	editok = 0
	
	if len(request("medid")) <> 0 then
	medid = request("medid")
	else
		if len(request.cookies("matstat")("medid")) <> 0 then
		medid = request.cookies("matstat")("medid")
		else
		medid = 0
		end if
	end if
	
	level = session("rettigheder")
	
	Response.cookies("matstat")("medid") = medid
	Response.cookies("matstat").expires = date + 1
	
	'*** Altid = 1 (Brug datointerval)
	fmudato = 1
	
	select case func
	case "afregnalle"
	
	if len(trim(request("afregnalle_id"))) <> 0 then
	afregnalle_id = request("afregnalle_id")
	else
	afregnalle_id = 0
	end if
	
	strSQL = "UPDATE materiale_forbrug SET afregnet = 1 WHERE bilagsnr = '"& afregnalle_id & "'"
	'Response.Write strSQL & "<br>"
	oConn.execute(strSQL)
	
	Response.redirect "materiale_stat.asp?hidemenu="&hidemenu
	
	
	case "opdafregnet"
	
	
	'Response.Write request("erafregnet") & "<br>"
	
	fids = split(trim(request("erafregnet")), ", ")
	for x = 0 to UBOUND(fids)
	
	thisval = trim(request("mfid_"&fids(x)&""))
	thisvalGK = trim(request("mfidgk_"&fids(x)&""))
	gkDate = year(now) & "/"& month(now) &"/"& day(now)
	
	'** kun hvis den ikke har været afregnet / udbewtalt før skal gkaf og gkdato ændres ***'
	godkendt = 0
	findes = 0
	strSQLfindes = "SELECT godkendt FROM materiale_forbrug WHERE id = "& fids(x)
	oRec.open strSQLfindes, oConn , 3
	if not oRec.EOF then
	
	godkendt = oRec("godkendt")
	findes = 1
	
	end if
	oRec.close
	
	if cint(godkendt) <> 1 then
	
	strSQL = "UPDATE materiale_forbrug SET afregnet = "& thisval &", godkendt = "& thisvalGK 
	    if thisvalGK = 1 then
	    strSQL = strSQL &", gkaf = '"& session("user") &"', gkdato = '"& gkDate &"'"
	    end if
	strSQL = strSQL &" WHERE id = "& fids(x)
	'Response.Write strSQL & "<br>"
	oConn.execute(strSQL)
	
	else
	'*** Opdater ikke timestamp for gokenels, samt godkendt af ***'
	strSQL = "UPDATE materiale_forbrug SET afregnet = "& thisval &", godkendt = "& thisvalGK &" WHERE id = "& fids(x)
	'Response.Write strSQL & "<br>"
	oConn.execute(strSQL)
	
	
	end if
	
	next
	
	'Response.end
	
	Response.redirect "materiale_stat.asp?hidemenu="&hidemenu
   
	
	
	
	case "dbred"
		
		
		antal = request("FM_antal")
		matfbrecordid = request("matrecid")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		
		    matid = 0
            matantal = 0
            matvarenr = 0
            
            '*** Henter materiale id og antal ***'
            strSQLsel = "SELECT matantal, matid, matvarenr FROM materiale_forbrug WHERE id = " & matfbrecordid
            oRec.open strSQLsel, oConn, 3
            if not oRec.EOF then
            
            matvarenr = oRec("matvarenr")
            matid = oRec("matid")
            matantal = oRec("matantal")
            
            end if
            oRec.close
		
		if antal = 0 then
		    
		    
		    '** Opdaterer antal på lager ***
	        if matvarenr <> 0 then
	        strSQL2 = "UPDATE materialer SET antal = (antal+("&matantal&")) WHERE id = "&matid
	        oCOnn.execute(strSQL2)
	        end if
		    
		
		strSQL = "DELETE FROM materiale_forbrug WHERE id = "& matfbrecordid 
		else
		        
		    '** Opdaterer antal på lager ***
	        if matvarenr <> 0 then
	        strSQL2 = "UPDATE materialer SET antal = (antal+("&(matantal-(antal))&")) WHERE id = "&matid
	        oCOnn.execute(strSQL2)
	        end if
		    
		
		strSQL = "UPDATE materiale_forbrug SET matantal = "& antal &", editor = '"& session("user")&"', dato = '"& strDato &"' WHERE id = "& matfbrecordid 
		end if
		'Response.write strSQL
		oConn.execute(strSQL)	
		
		
		progrp_medarb = request("FM_radio_projgrp_medarb")
		progrp = request("FM_progrupper")
		selmedarb = request("FM_medarb")
		kundeid = request("FM_kunde")
		kundeans = request("FM_kundeans")
		jobans = request("FM_jobans")
		visKundejobans = request("FM_kundejobans_ell_alle")
		jobid = request("FM_job")
		
		response.redirect "materiale_stat.asp?id="&id&""_
		&"&FM_radio_projgrp_medarb="&progrp_medarb&""_
		&"&FM_progrupper="&progrp&""_
		&"&FM_medarb="&selmedarb&""_
		&"&FM_kunde="&kundeid&""_
		&"&FM_kundeans="&kundeans&""_
		&"&FM_jobans="&jobans&""_
		&"&FM_kundejobans_ell_alle="&visKundejobans&""_
		&"&FM_job="&jobid&"&hidemenu="&hidemenu
		
		
		
	case else
	
	thisJobnr = 0
	%>
	
	
	
	
	
	
	<script>



	  
	
	
	function BreakItUp2()
	{
	  //Set the limit for field size.
	  var FormLimit = 102399
	
	  //Get the value of the large input object.
	  var TempVar = new String
	  TempVar = document.theForm2.BigTextArea.value
	
	  //If the length of the object is greater than the limit, break it
	  //into multiple objects.
	  if (TempVar.length > FormLimit)
	  {
	    document.theForm2.BigTextArea.value = TempVar.substr(0, FormLimit)
	    TempVar = TempVar.substr(FormLimit)
	
	    while (TempVar.length > 0)
	    {
	      var objTEXTAREA = document.createElement("TEXTAREA")
	      objTEXTAREA.name = "BigTextArea"
	      objTEXTAREA.value = TempVar.substr(0, FormLimit)
	      document.theForm2.appendChild(objTEXTAREA)
	
	      TempVar = TempVar.substr(FormLimit)
	    }
	  }
	}
	
	function visjob() {
	id = document.getElementById("id").value
	//alert("default.asp?typethis="+tp+"&clicked="+cl+"&FM_fab="+fab+"")
	window.location.href = "materiale_stat.asp?menu=stat&id="+id
	}	
	
	function clearJobsog(){
	document.getElementById("FM_jobsog").value = ""
}


function clearK_Jobans() {
    //alert("her")
    document.getElementById("cFM_jobans").checked = false
    document.getElementById("cFM_jobans2").checked = false
    document.getElementById("cFM_kundeans").checked = false
}


function setshowprogrpval() {
    document.getElementById("showprogrpdiv").value = 1
}


function closeprogrpdiv() {
    document.getElementById("progrpdiv").style.visibility = "hidden"
    document.getElementById("progrpdiv").style.display = "none"
}


function curserType(tis) {
    document.getElementById(tis).style.cursor = 'hand'
}

 


	           
        
        
        
        
	</script>
	
	
	<%if print <> "j" AND cint(hidemenu) <> 1 then%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(7)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call stattopmenu()
	end if
	
	
	tdclass = ""
	tblwdt = 900
	tdwtd = 120%>
	</div>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sideindhold" style="position:absolute; left:20px; top:132px; visibility:visible;">
	
	<%else%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sideindhold" style="position:absolute; left:20px; top:20px; visibility:visible;">
	<%
	tdclass = "lille"
	tblwdt = 600
	tdwtd = 80
	
	
	end if '*Print%>
	
	<script>

	    $(document).ready(function() {

	    if (document.getElementById("FM_visprjob_ell_sum2").checked == true) { 
	        $("#ysiu").css("display", "");
	        $("#ysiu").css("visibility", "visible");
            }
            
        
        $("#FM_visprjob_ell_sum2").click(function() {
            $("#ysiu").css("display", "");
	        $("#ysiu").css("visibility", "visible");
	        $("#ysiu").show("slow");
	    });


	    $("#FM_visprjob_ell_sum0").click(function() {
	    $("#ysiu").css("display", "none");
	    $("#ysiu").css("visibility", "hidden");
	    $("#ysiu").hide("slow");
	    });

	    $("#FM_visprjob_ell_sum1").click(function() {
	        $("#ysiu").css("display", "none");
	        $("#ysiu").css("visibility", "hidden");
	        $("#ysiu").hide("slow");
	    });
            
	    });     
	
	
	</script>
	
	
	<%
	
	 oimg = "ikon_matforbrug.png"
	oleft = 0
	otop = 0
	owdt = 400
	oskrift = "Materialeforbrug / Udlæg"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	call filterheader(0,0,800,pTxt)
	
	
	
	'**************************************************
	'***** Faste Filter kriterier *********************
	'**************************************************
		
	
	'*** Brug Kundeans / Jobans ***
	call kundeogjobans()
	
	'*** Medarbejdere / projektgrupper
	selmedarb = session("mid")
	call medarbogprogrp("matstat")
	
	
	'**** Kundekri ***
	if len(request("FM_kunde")) <> 0 then
	kundeid = request("FM_kunde")
	else
	kundeid = 0
	end if
	
	strKnrSQLkri = " OR jobknr = 0 "
	
	'***** Valgt job eller søgt på Job ****
	'** hvis Sog = Ja
	call jobsog()
	
	
	'************ slut faste filter var **************		
	
	
		
	
	
			
			
			'**** Fordeling pr job og medarb eller sum ***
			if len(request("FM_visprjob_ell_sum")) <> 0 then
			visprjob_ell_sum = request("FM_visprjob_ell_sum")
				
				select case cint(visprjob_ell_sum) 
				case 0 
				visprjob_ell_sum_chk0 = "CHECKED"
				visprjob_ell_sum_chk1 = ""
				visprjob_ell_sum_chk2 = ""
				case 2
				visprjob_ell_sum_chk0 = ""
				visprjob_ell_sum_chk1 = ""
				visprjob_ell_sum_chk2 = "CHECKED"
				case else
				visprjob_ell_sum_chk0 = ""
				visprjob_ell_sum_chk1 = "CHECKED"
				visprjob_ell_sum_chk2 = ""
				end select
				
				Response.Cookies("matstat")("vtype") = visprjob_ell_sum
				
			else
			    
			    if request.Cookies("matstat")("vtype") <> "" then
			    
			    visprjob_ell_sum = request.Cookies("matstat")("vtype")
			    
			        select case cint(visprjob_ell_sum) 
				    case 0 
				    visprjob_ell_sum_chk0 = "CHECKED"
				    visprjob_ell_sum_chk1 = ""
				    visprjob_ell_sum_chk2 = ""
				    case 2
				    visprjob_ell_sum_chk0 = ""
				    visprjob_ell_sum_chk1 = ""
				    visprjob_ell_sum_chk2 = "CHECKED"
				    case else
				    visprjob_ell_sum_chk0 = ""
				    visprjob_ell_sum_chk1 = "CHECKED"
				    visprjob_ell_sum_chk2 = ""
				    end select
			    
			    else
			    visprjob_ell_sum = 0
			    visprjob_ell_sum_chk0 = "CHECKED"
			    visprjob_ell_sum_chk1 = ""
			    visprjob_ell_sum_chk2 = ""
			    end if
			    
			end if
			
			        select case cint(visprjob_ell_sum) 
				    case 0 
				    visningsTypVal = "Materialeforbr. / Udlæg fordelt på job"
				    case 2
				    visningsTypVal = "Udgiftsbilag"
				    case else
				    visningsTypVal = "Totalforbrug"
				    end select
	
		
	
	%>
	
	<table cellspacing=0 cellpadding=5 border=0 width=100% bgcolor="#FFFFFF">
	
	<%if print <> "j" then%>
	<form action="materiale_stat.asp?hidemenu=<%=hidemenu %>" method="post">
	
	
	
	<tr>
	<td style="padding:5px;"><input type="radio" name="FM_radio_projgrp_medarb" value="1" <%=projgrp_medarb1%>>&nbsp;<b>Medarbejder(e):</b><br>
	<select name="FM_medarb" style="font-size : 11px; width:205px;">
	
	<%
	end if
	
	if level <=2 OR level = 6 then
	mwhSQLkri = " mansat <> 2 "
	
	if print <> "j" then%>
	<option value="0">Alle</option>
	<%else
	medVal = "Alle"
	end if
	
	else
    mwhSQLkri = " mid = " & selmedarb
	end if
	
	strSQL = "SELECT Mid, Mnavn, Mnr, mansat, init FROM medarbejdere WHERE "& mwhSQLkri &" ORDER BY Mnavn"
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
		
		if cint(selmedarb) = oRec("mid") then
		thisChecked = "SELECTED"
		medVal = oRec("mnavn") & " ("& oRec("mnr") &") - "& oRec("init")
		else
		thisChecked = ""
		end if
		
		if print <> "j" then%>
		<option value="<%=oRec("mid")%>" <%=thisChecked%>><%=oRec("mnavn")%>&nbsp;(<%=oRec("mnr")%>) - <%=oRec("init")%></option>
	    <%  
	    end if
	
	oRec.movenext
	wend
	oRec.close
	
	
	if print <> "j" then%>
	</select>
	</td>
	<%else %>
	<tr><td><b>Medarbejder(e):</b></td><td><%=medVal %></td>
	<%end if%>
	    
	    
	 
	
	<%
	if level <=2 OR level = 6 then
	
	if print <> "j" then %>
	<td>
	<input type="radio" name="FM_radio_projgrp_medarb" value="2" <%=projgrp_medarb2%>>&nbsp;<b>Projektgruppe(r):</b><br>
	<%
	end if
	
	strSQL = "SELECT p.id AS id, navn, count(medarbejderid) AS antal FROM projektgrupper p "_
	&" LEFT JOIN progrupperelationer ON (ProjektgruppeId = p.id) GROUP BY p.id ORDER BY navn"
	oRec.open strSQL, oConn, 3
	
	'Response.write strSQL
	'Response.flush
	
	if print <> "j" then%>
	<select name="FM_progrupper" style="font-size : 11px; width:205px;">
	<%end if
			
			
			while not oRec.EOF
			
			if cint(progrp) = cint(oRec("id")) then
			isSelected = "SELECTED"
			progrpVal = oRec("navn")
			else
			isSelected = ""
			end if
			
			if cint(oRec("antal")) > 0 AND print <> "j" then%>
			<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("navn")%></option>
			<%end if
			
			oRec.movenext
			wend
			oRec.close
			
    if print <> "j" then%>
	</select>
	</td>
	<td align=right>&nbsp;<input type="submit" value=" Vælg medarb. >> "></td>
	</tr>
	</table>
	
	<%if progrp_medarb = 2 then %>
	
	<div id="progrpdiv" style="position:absolute; left:810px; top:120px; width:200px; height:250px; overflow:auto; background-color:#ffffff; border:1px #c4c4c4 solid; padding:10px; font-size:9px; z-index:20000;" onclick="closeprogrpdiv();" onmouseover="curserType('progrpdiv');">
	
	    <b>Medarbejdere i valgt projektgruppe:</b><br />
	    <%=medarbigrp %><br />
	    <a href="#" onclick="closeprogrpdiv();" class=red>[Klik her for at lukke]</a>
	    </div>
	    
	    
	    <input id="showprogrpdiv" name="showprogrpdiv" value="0" type="hidden" />
	
	    
	<%end if %>
	
	<%else 'print
	
	    if progrp_medarb = 2 then %>
	    <tr><td><b>Projektgruppe:</b></td><td> <%=progrpVal%></td></tr>
	    <%end if %>
	    
    <%end if %>
    
    <%else %>
    <td>
        &nbsp;</td>
        </tr>
        </table>
    <%end if %> <!-- level -->
	    
	
	
	<%if print <> "j" then %>
	<table cellspacing=0 cellpadding=5 border=0 width=100% bgcolor="#FFFFFF">
	
	
	<tr bgcolor="#Eff3ff">
		<td><br />
		<input type="radio" name="FM_kundejobans_ell_alle" value="1" <%=kundejobansCHK1%>><b>A) For Key account</b>, vis kun kunder og job hvor
		valgte ovenstående medarbejder(e) er:<br />
		<input type="checkbox" name="FM_kundeans" id="cFM_kundeans" value="1" <%=kundeansChk%>>&nbsp;Kundeansvarlig. (1 ell. 2)<br>
		<input type="checkbox" name="FM_jobans" id="cFM_jobans" value="1" <%=jobansChk%>>&nbsp;Jobansvarlig (jobansv. 1)<br />
		<input type="checkbox" name="FM_jobans2" id="cFM_jobans2" value="1" <%=jobansChk2%>>&nbsp;Jobejer (jobansv. 2)
		
		</td>
		<td bgcolor="#Eff3ff">&nbsp;</td>
	</tr>
	
	
	<tr bgcolor="#Eff3ff">
		<td valign=top><input type="radio" name="FM_kundejobans_ell_alle" value="0" <%=kundejobansCHK0%> onclick="clearK_Jobans();"><b>B) Som Virksomhed</b><br>
		<br><b>Kontakt(er):</b><br />
		<%
		
		
   end if
   
   
		strKnrSQLkri = " OR jobknr = -1 "
		strAftKidSQLkri = " OR kundeid = -1 "
		
		strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE (" & kundeAnsSQLKri & ") AND ketype <> 'e' ORDER BY Kkundenavn"
		'Response.write strSQL & "<br>"
		'Response.flush		
	if print <> "j" then%>
		<select name="FM_kunde" id="Select1" style="font-size : 11px; width:466px;">
		<%
		
    end if
				
				oRec.open strSQL, oConn, 3
				k = 0
				while not oRec.EOF
				
				if k = 0 ANd print <> "j" then
				%>
				<option value="0">Alle</option>
				<%end if
				
      
				if cint(kundeans) = 1 then
				strKnrSQLkri = strKnrSQLkri & " OR jobknr = "& oRec("kid")
				strAftKidSQLkri = strAftKidSQLkri & " OR kundeid = " & oRec("kid")
				else
				strAftKidSQLkri =  " OR kundeid <> 0 "
				strKnrSQLkri = " OR jobknr <> 0 "
				end if
				
				if print <> "j" then
				
				if cint(kundeid) = cint(oRec("kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%>&nbsp;(<%=oRec("Kkundenr")%>)</option>
				<%
				
				end if
				
				k = k + 1
				oRec.movenext
				wend
				oRec.close
				
		if print <> "j" then
				
				if cint(kundeid) = -1 then
				ingChk = "SELECTED"
				else
				ingChk = ""
				end if%>
				<option value="-1" <%=ingChk%>>Ingen</option>
				
				
		</select>
		
		<% end if '** print
				
				len_strKnrSQLkri = len(strKnrSQLkri)
				right_strKnrSQLkri = right(strKnrSQLkri, len_strKnrSQLkri - 3)
				strKnrSQLkri = right_strKnrSQLkri
				
				len_strAftKidSQLkri = len(strAftKidSQLkri)
				right_strAftKidSQLkri = right(strAftKidSQLkri, len_strAftKidSQLkri - 3)
				strAftKidSQLkri = right_strAftKidSQLkri
				
				
				
		if print <> "j" then		%>
		
		
		</td><td align=right valign=bottom>
	        <input type="submit" value=" Vælg Kontakt(er) >> ">
		</td>
		
	</tr>
	</table>
	<%
		else
		%>
		<tr><td valign=top><b>Vis for valgte medarb. som key account<br /> ell. samlet for valgte kontakt:</b></td>
		<td valign=top><%=kontakt_keyaccountVAL %>
		<%if cint(visKundejobans) = 1 then%>
		<br /><%=jobansVal %>
		<br /><%=kansVal %>
		<%end if %>
		</td></tr>
		<%
		end if '** print 
	
	
	
	if print <> "j" then %>
	<table cellspacing=0 cellpadding=5 border=0 width=100%>
	<tr>
	<!--
	<form action="materiale_stat.asp?hidemenu=<%=hidemenu %>" method="post">
		<input type="hidden" name="FM_radio_projgrp_medarb" id="FM_radio_projgrp_medarb" value="<%=progrp_medarb%>">
		<input type="hidden" name="FM_progrupper" id="FM_progrupper" value="<%=progrp%>">
		<input type="hidden" name="FM_medarb" id="FM_medarb" value="<%=selmedarb%>">
		<input type="hidden" name="FM_kunde" id="FM_kunde" value="<%=kundeid%>">
		<input type="hidden" name="FM_kundeans" id="FM_kundeans" value="<%=kundeans%>">
		<input type="hidden" name="FM_jobans" id="FM_jobans" value="<%=jobans%>">
		<input type="hidden" name="FM_kundejobans_ell_alle" value="<%=visKundejobans%>">
		-->
		
		<td>
		<b>Vælg job:</b><br />
		<%
		end if
		
		
		
		
		if cint(kundeid) = 0 then
		strKnrSQLkri = "("&strKnrSQLkri&")" '" jobknr = 0" '
		else
		strKnrSQLkri = " jobknr = "& kundeid
		end if
		
		'*** Er jobansvarlig tilvalgt ? ** 
		if cint(jobans) = 1 then
		strKnrSQLkri = ""&strKnrSQLkri &" "& jobAnsSQLkri
		end if
		
		'*** Er jobansvarlig 2 (jobejer) tilvalgt ? ** 
		if cint(jobans2) = 1 then
		strKnrSQLkri = ""&strKnrSQLkri &" "& jobAns2SQLkri
		end if
		
		strSQL = "SELECT jobnavn, jobnr, id, k.kkundenavn, k.kkundenr FROM job j "_
		&" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
		&" WHERE "& strKnrSQLkri &" AND jobstatus <> 3 ORDER BY jobnavn"
		'Response.write strSQL
		'Response.flush
		
		
		if print <> "j" then%>
		<select name="FM_job" id="FM_job" style="font-size : 11px; width:466px;" onChange="clearJobsog(), submit();">
		<option value="0">Alle - eller vælg fra liste...</option>
		<%
		else
		jobVal = "Alle"
		kontaktVal = "Alle"
		end if
				
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(jobid) = cint(oRec("id")) then
				isSelected = "SELECTED"
				jobVal = oRec("jobnavn") & " ("& oRec("jobnr") &")"
				kontaktVal = oRec("kkundenavn") & " ("& oRec("kkundenr") &")"  
				else
				isSelected = ""
				end if
				
				if print <> "j" then%>
				<option value="<%=oRec("id")%>" <%=isSelected%>><%=oRec("jobnavn")%>&nbsp;(<%=oRec("jobnr")%>) ][ <%=oRec("kkundenavn")%>&nbsp;(<%=oRec("kkundenr")%>) </option>
				<%end if
				
				oRec.movenext
				wend
				oRec.close
				
				if cint(jobid) = -1 then
				mChk = "SELECTED"
				kontaktVal = "Ingen"
				jobVal = "Ingen"
				else
				mChk = ""
				end if
				
				if print <> "j" then%>
				<option value="-1" <%=mChk%>>Ingen</option>
		        </select>
		        
		        <br><img src="ill/blank.gif" width="1" height="5" /><br />
		        <b>ell. søg på jobnavn ell. nr:</b>&nbsp;&nbsp;<input type="text" name="FM_jobsog" id="FM_jobsog" value="<%=jobSogVal%>" style="width:200px;">&nbsp;<input id="Submit3" type="submit" value=" Søg >> " />&nbsp;(% wildcard, f.eks: "%analyse")
		        </td>
	</tr>
	<%else 
	        if len(trim(jobSogVal)) <> 0 then%>
    	    <tr><td><b>Jobsøgning:</b></td><td><%=jobSogVal %></td></tr>
	        <%else %>
	        <tr><td><b>Kontakt(er):</b></td><td><%=kontaktVal %></td></tr>
	        <tr><td><b>Job:</b></td><td><%=jobVal %></td></tr>
	        <%end if %>
	    
    <%end if %>
	
	
	<%
	'*** valgte job ***
	call valgtejob()

				
				if j >= 1 then
				len_jobKri = len(jobKri)
				right_jobKri = right(jobKri, len_jobKri - 3)
				jobKri = right_jobKri & ")"
				end if
				
				if cint(kundeid) = 0 AND cint(kundeans) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND len(trim(jobSogVal)) = 0 then
				jobKri = " mf.jobid <> 0 "
				else
				jobKri = jobKri
				end if
				
	
	%>
	
	<%if print = "j" then
	dontshowDD = "1"
	%>
	<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
	
	<tr><td><b>Periode:</b></td><td>
	<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>
	</td></tr>
	
	<%
	else
	%>
	<tr bgcolor="#Eff3ff">
			<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<td><b>Periode:</b><br>
			<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
			</td>
		</tr>
	<%end if%>
	
	
	<%
	'*** visningstype kriteriuer, bilagsnr, personlige mm.
	
	if len(trim(request("FM_bilagsnr"))) <> 0 then 
	        sogliste = trim(request("FM_bilagsnr"))
	        useSog = 1
	        else
	        sogliste = ""
	        useSog = 0
	        end if
	        %> 
    
    
    <%if len(request("FM_intkode")) <> 0 then
			intKode = request("FM_intkode")
			
			
			'Response.Write "intKode  "& intKode 
			    
			    select case intKode
			    case 1
			    ikSEL0 = ""
			    ikSEL1 = "SELECTED"
			    ikSEL2 = ""
			    ikSEL3 = ""
			    intKodeSQLkri = " AND intkode = 1 "
			    intKodeVal = "Intern" 
			    case 2
			    ikSEL0 = ""
			    ikSEL1 = ""
			    ikSEL2 = "SELECTED"
			    ikSEL3 = ""
			    intKodeSQLkri = " AND intkode = 2 "
			    intKodeVal = "Ekstern"
			    'case 3
			    'ikSEL0 = ""
			    'ikSEL1 = ""
			    'ikSEL2 = ""
			    'ikSEL3 = "SELECTED"
			    'intKodeSQLkri = " AND intkode = 3 "
			    'intKodeVal = "Personlig"
			    case else
			    ikSEL0 = "SELECTED"
			    ikSEL1 = ""
			    ikSEL2 = ""
			    ikSEL3 = ""
			    intKodeSQLkri = " AND intkode <> -1 "
			    intKodeVal = "Alle"
			    end select
			
			
			else
			intKode = 0
			ikSEL0 = "SELECTED"
			ikSEL1 = ""
			ikSEL2 = ""
			ikSEL3 = ""
			intKodeSQLkri = " AND intkode <> -1 "
			intKodeVal = "Alle"
			end if %>
            
            
            
              <%
           if len(trim(request("FM_intkode"))) <> 0 AND print <> "j" then
                if len(trim(request("showonlypers"))) <> 0 then
                showonlypers = 1
                showonlypersCHK = "CHECKED"
                else
                showonlypers = 0
                showonlypersCHK = ""    
                end if
           else
               
               if Request.Cookies("mat")("cshowonlypers") <> "" then
                    showonlypers = request.Cookies("mat")("cshowonlypers")
                    
                    if showonlypers = 1 then
                    showonlypersCHK = "CHECKED"
                    else
                    showonlypersCHK = ""
                    end if
                    
                else
                    showonlypers = 0
                    showonlypersCHK = ""
               end if
         
           end if 
           
           Response.Cookies("mat")("cshowonlypers") = showonlypers
           
           
           
            if len(trim(request("FM_intkode"))) <> 0 AND print <> "j" then
                if len(trim(request("showonlyikkeafreg"))) <> 0 then
                showonlyikkeafreg = 1
                showonlyikkeafregCHK = "CHECKED"
                else
                showonlyikkeafreg = 0
                showonlyikkeafregCHK = ""    
                end if
           else
               
               if Request.Cookies("mat")("cshowonlyikkeafreg") <> "" then
                    showonlyikkeafreg = request.Cookies("mat")("cshowonlyikkeafreg")
                    
                    if showonlyikkeafreg = 1 then
                    showonlyikkeafregCHK = "CHECKED"
                    else
                    showonlyikkeafregCHK = ""
                    end if
                    
                else
                    showonlyikkeafreg = 0
                    showonlyikkeafregCHK = ""
               end if
         
           end if 
           
           
           Response.Cookies("mat")("cshowonlyikkeafreg") = showonlyikkeafreg
           
           if len(trim(request("FM_intkode"))) <> 0 AND print <> "j" then
                if len(trim(request("showonlyikkegk"))) <> 0 then
                showonlyikkegk = 1
                showonlyikkegkCHK = "CHECKED"
                else
                showonlyikkegk = 0
                showonlyikkegkCHK = ""    
                end if
           else
               
               if Request.Cookies("mat")("cshowonlyikkegk") <> "" then
                    showonlyikkegk = request.Cookies("mat")("cshowonlyikkegk")
                    
                    if showonlyikkegk = 1 then
                    showonlyikkegkCHK = "CHECKED"
                    else
                    showonlyikkegkCHK = ""
                    end if
                    
                else
                    showonlyikkegk = 0
                    showonlyikkegkCHK = ""
               end if
         
           end if 
           
           
           Response.Cookies("mat")("cshowonlyikkegk") = showonlyikkegk
           
           
           
           '**** slut visning type kriterier ***'%>     
	
	
	<%if print <> "j" then %>
	<table cellspacing=0 cellpadding=5 border=0 width=100%>
		<tr>
			<td valign=top><br />
			<b>Visningsttype:</b><br />
			<input type="radio" name="FM_visprjob_ell_sum" id="FM_visprjob_ell_sum0" value="0" <%=visprjob_ell_sum_chk0%>> Vis Materialeforbr. / Udlæg fordelt på job<br />
			<input type="radio" name="FM_visprjob_ell_sum" id="FM_visprjob_ell_sum1" value="1" <%=visprjob_ell_sum_chk1%>> Vis totalforbrug<br />
			<input type="radio" name="FM_visprjob_ell_sum" id="FM_visprjob_ell_sum2" value="2" <%=visprjob_ell_sum_chk2%>> Vis som udgiftsbilag
			</td>
			
			<td valign=top style="padding-top:30px;">
			
			<div id="ysiu" style="postion:relative; display:none; visibility:hidden; border:1px #cccccc dashed; padding:10px; background-color:#ffffe1;">
			<b>Yderligere søgekriterier ved udgiftsbilag's visning.</b><br />
		    
		    <br />Søg på bilags nummer:<br /><input id="FM_bilagsnr" name="FM_bilagsnr" size="20" type="text" value="<%=sogliste %>" />&nbsp;% = wildcard<br />
			<%=tsa_txt_231 %>: <br />
			
			
			
			<select name="FM_intkode" id="FM_intkode" style="width:55px; font-size:9px;">
		    <option value="0" <%=ikSEL0 %>><%=tsa_txt_076 %></option>
		    <option value="1" <%=ikSEL1 %>><%=tsa_txt_232 %></option>
		    <option value="2" <%=ikSEL2 %>><%=tsa_txt_233 %></option>
		    <!--
		    <option value="3" <%=ikSEL3 %>><=tsa_txt_234 %></option>
		    -->
		    
    		</select>
    		
    		<br />
                
                
           
            <input id="Checkbox2" name="showonlypers" id="showonlypers" type="checkbox" <%=showonlypersCHK %> /> <%=tsa_txt_320 %>
             <br /><input id="Checkbox1" name="showonlyikkeafreg" id="showonlyikkeafreg" type="checkbox" <%=showonlyikkeafregCHK %> /> <%=tsa_txt_322 %>
             <br /><input id="Checkbox3" name="showonlyikkegk" id="showonlyikkegk" type="checkbox" <%=showonlyikkegkCHK %> /> <%=tsa_txt_324 %>
             
             
                </div>  
		    
		    </td>
		    <td style="width:150px;">&nbsp;</td>
		
		</tr>
	
    
    
    
		
		<tr><td align=right colspan=3>
		   <input type="submit" value=" Kør statistik >> ">
			&nbsp;
		</td></tr>
		</form>
		</table>
		
	
	
	<%else %>
	
	<!-- for at undgå javafejl onload  -->
	<div id="ysiu" style="postion:relative; display:none; visibility:hidden; width:1px; height:1px;">
        <input id="FM_visprjob_ell_sum2" type="checkbox" />
	</div>		
	<!-- -->
	
	
	  <tr><td><b>Visningstype:</b></td><td><%=visningsTypVal %></td></tr>
	  
	<%if visningsTypVal = "Udgiftsbilag" then %>
	<tr><td><b>Søgning på bilags nr:</b></td><td><%=sogliste %>&nbsp;</td></tr>
	<tr><td><b>Intern kode:</b></td><td><%=intKodeVal%></td></tr>
	<tr><td><b>Personlig:</b></td><td>
	<%if cint(showonlypers) <> 0 then %>
	<%=tsa_txt_320 %>
	<%else %>
	Alle
	<%end if %></td></tr>
	
	<tr><td><b>Afregnet:</b></td><td>
	<%if cint(showonlyikkeafreg) <> 0 then %>
	<%=tsa_txt_322 %>
	<%else %>
	Alle
	<%end if %></td></tr>
	
	
	<tr><td><b><%=tsa_txt_323 %>:</b></td><td>
	<%if cint(showonlyikkegk) <> 0 then %>
	<%=tsa_txt_324 %>
	<%else %>
	Alle
	<%end if %></td></tr>
	
	<%end if '**udgiftbilag %>
	
	<%end if %>
	
	</table>
	
	<!-- filter div -->
	</td></tr></table>
	</div><br /><br />
	
	
	

	
	<%
	'*** SQL datoer ***
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	
	if cint(fmudato) = 1 then
	strDatoKri = " AND mf.forbrugsdato BETWEEN '"& strStartDato &"' AND '"& strSlutDato &"'"
	else
	strDatoKri = ""
	end if
	
	orderby = request("orderby") 
	if len(request("orderby")) = 0 OR orderby = "dato" then
	orderbyKri = "mf.forbrugsdato DESC, mf.matnavn"
	else
	    if orderby = "matnr" then
	    orderbyKri = "mf.matvarenr"
	    else
	    orderbyKri = "mf.matnavn"
	    end if
	end if
	
	if cint(jobans) = 1 OR cint(kundeans) = 1 then
	medarbSQlKri = " usrid <> 0 "
	else
	medarbSQlKri = medarbSQlKri
	end if
	
	if cint(jobid) <> 0 then
	jobKri = "mf.jobid = "& jobid &""
	else
	jobKri = jobKri
	end if
	
	
	
	                if print <> "j" then
                    mainTableWth = "1104"
                	else
                	mainTableWth = "950"
                	end if
	
	
	select case visprjob_ell_sum 
	case 0 '** Matforbrug fordelt pr. job ***'
	
	
	                '**** Lukket for redigering *'
	                call ersmileyaktiv()
	                call licensStartDato()
                	
	                stDato = startDatoAar &"/"& startDatoMd &"/"& startDatoDag
	                slDato = year(now) &"/"& month(now) &"/"& day(now)
                	
                	
	                strSQL = "SELECT  mid FROM medarbejdere WHERE mansat <> '2' ORDER BY mid"
	                oRec.open strSQL, oConn, 3
	                while not oRec.EOF 
                	
	                call afsluger(oRec("mid"), stdato, sldato)
                	
	                oRec.movenext
	                wend
	                oRec.close
                	
                	
	               
		        tTop = 20
	            tLeft = 0
	            tWdth = mainTableWth
            	
            	
	            call tableDiv(tTop,tLeft,tWdth)
	            %>
	                
	                <table cellspacing="0" cellpadding="2" border="0" bgcolor="#ffffff" width=100%>
	                
	                <tr bgcolor="#5582D2">
		                <td>&nbsp;</td>
		                <td class=alt><b>Job/Kunde</b></td>
		                <td height=20 class=alt><a href="materiale_stat.asp?menu=stat&orderby=dato&id=<%=id%><%=link%>&hidemenu=<%=hidemenu%>" class=alt><u><b>Dato</b></u></a> (Medarb.)</td>
		                <td class=alt style="padding-left:5px;"><b>Gruppe / lager</b> (Ava. %)</td>
		                <td class=alt style="padding-left:5px;"><a href="materiale_stat.asp?menu=stat&orderby=navn&id=<%=id%><%=link%>&hidemenu=<%=hidemenu%>" class=alt><u><b>Navn</b></u></a>&nbsp;&nbsp;
		                <a href="materiale_stat.asp?menu=stat&orderby=matnr&id=<%=id%><%=link%>&hidemenu=<%=hidemenu%>" class=alt><u><b>(Mat. nr.)</b></u></a></td>
		                <td class=alt><b>Antal</b></td>
		                <td class=alt>&nbsp;</td>
		                <td class=alt align=right><b>På lager</b></td>
		                <td class=alt align=right><b>Min. lager</b></td>
		                <td class=alt align=right><b>Købspris pr. stk</b></td>
		                <td class=alt align=right><b>Salgspris pr. stk.</b></td>
		                <td class=alt><b>Valuta</b></td>
		                <td class=alt align=right><b>Kurs</b></td>
		                <td class=alt align=right><b>Købspris ialt</b></td>
		                <td class=alt align=right><b>Salgspris ialt</b></td>
		                <td class=alt><b>Valuta</b></td>
		                <td class=alt>&nbsp;</td>
		                <td>&nbsp;</td>
	                </tr>
	                <%
                	
	                if lto = "bowe" then
	                strExport = "Jobnavn;Job Nr;KA nr.;Böwe kode;Kundennavn;Kunde Nr;Forbrugs dato;Medarbejder;Medarb. Nr;Initialer;Gruppe;Navn;Varenr;Antal;Enhed;Købspris pr. stk.;Salgspris pr. stk; Valuta; Kurs; Købspris ialt; Salgspris ialt; Valuta;"
	                else
	                strExport = "Jobnavn;Job Nr;Kundennavn;Kunde Nr;Forbrugs dato;Medarbejder;Medarb. Nr;Initialer;Gruppe;Navn;Varenr;Antal;Enhed;Købspris pr. stk.;Salgspris pr. stk;Valuta; Kurs; Købspris ialt; Salgspris ialt; Valuta;"
	                end if
                	
	                strSQL = "SELECT m.mnavn AS medarbejdernavn, mf.id, mf.matvarenr AS varenr, mg.navn AS gnavn, mf.matenhed AS enhed, "_
	                &" mf.matnavn AS navn, mf.matantal AS antal, mf.dato, mf.editor, mf.matid, "_
	                & "mf.matkobspris, mf.matsalgspris, mf.jobid, mf.matgrp, mf.usrid, mf.forbrugsdato, mg.av, "_
	                &" j.jobans1, j.jobans2, j.jobnavn, j.jobnr, k.kkundenavn, k.kkundenr, m.init, f.fakdato, "_
	                &" m.mnr, ma.antal AS palager, ma.minlager, "_
	                &" mf.valuta, mf.intkode, v.valutakode, mf.kurs AS mfkurs "_
	                &" FROM materiale_forbrug mf"_
	                &" LEFT JOIN materiale_grp mg ON (mg.id = mf.matgrp) "_
	                &" LEFT JOIN medarbejdere m ON (mid = mf.usrid) "_
	                &" LEFT JOIN job j ON (j.id = mf.jobid) "_
	                &" LEFT JOIN kunder k ON (kid = jobknr)"_
	                &" LEFT JOIN materialer ma ON (ma.id = mf.matid)"_
	                &" LEFT JOIN fakturaer f ON (f.jobid = mf.jobid AND f.faktype = 0)"_
	                &" LEFT JOIN valutaer v ON (v.id = mf.valuta) "_
	                &" WHERE "&jobKri&" AND ("& medarbSQlKri &") "& strDatoKri &" GROUP BY mf.id ORDER BY "& orderbyKri	
                	
	                'response.write strSQL
	                'Response.flush
                	
	                x = 0
                	
	                strMatids = "0"
	                'strMatAntal = "0"
	                strAntalFM = ""
	                oRec.open strSQL, oConn, 3
	                while not oRec.EOF 
                	
	                strMatids = strMatids &","& oRec("matid")
	                'strMatAntal = strMatAntal &","& oRec("antal")
                					
                					
                					
					                editok = 0
					                '*** rediger rettigheder ***
					                if level = 1 then
					                editok = 1
					                else
							                if cint(session("mid")) = jobans1 OR cint(session("mid")) = jobans2 OR (cint(jobans1) = 0 AND cint(jobans2) = 0) then
							                editok = 1
							                end if
					                end if
                					
                					
                			    
			                    '*** Er uge alfsuttet af medarb, er smiley og autogk slået til
                                erugeafsluttet = instr(afslUgerMedab(oRec("usrid")), "#"&datepart("ww", oRec("forbrugsdato"),2,2)&"_"& datepart("yyyy", oRec("forbrugsdato")) &"#")
                                
                                'Response.Write "erugeafsluttet --" & erugeafsluttet  &"<br>"
                                'Response.flush
                                
                              
                                 if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) = year(now) AND DatePart("m", oRec("forbrugsdato")) < month(now)) OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) = 12)) OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) <> 12) OR _
                                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", oRec("forbrugsdato")) > 1))) then
                              
                                ugeerAfsl_og_autogk_smil = 1
                                else
                                ugeerAfsl_og_autogk_smil = 0
                                end if 
                				
				                if len(oRec("fakdato")) <> 0 then
	                            fakdato = oRec("fakdato")
	                            else
	                            fakdato = "01/01/2002"
	                            end if
                				
                				
				                if (ugeerAfsl_og_autogk_smil = 0 _
				                OR ugeerAfsl_og_autogk_smil = 1 AND level = 1) _
				                AND cdate(fakdato) < cdate(oRec("forbrugsdato")) then
				                editok = editok
				                else
				                editok = 0
				                end if
                					
                	
	                bgthis = "#ffffff"
	                %>
	                <form name="<%=oRec("id")%>" id="<%=oRec("id")%>" action="materiale_stat.asp?func=dbred&id=<%=id%>&matrecid=<%=oRec("id")%>&hidemenu=<%=hidemenu%>" method="post">
	                <input type="hidden" name="FM_radio_projgrp_medarb" id="FM_radio_projgrp_medarb" value="<%=progrp_medarb%>">
		                <input type="hidden" name="FM_progrupper" id="FM_progrupper" value="<%=progrp%>">
		                <input type="hidden" name="FM_medarb" id="FM_medarb" value="<%=selmedarb%>">
		                <input type="hidden" name="FM_kunde" id="FM_kunde" value="<%=kundeid%>">
		                <input type="hidden" name="FM_kundeans" id="FM_kundeans" value="<%=kundeans%>">
		                <input type="hidden" name="FM_jobans" id="FM_jobans" value="<%=jobans%>">
		                <input type="hidden" name="FM_kundejobans_ell_alle" value="<%=visKundejobans%>">
		                <input type="hidden" name="FM_job" id="FM_job" value="<%=jobid%>">
		                <input type="hidden" name="FM_visprjob_ell_sum" id="FM_visprjob_ell_sum" value="<%=visprjob_ell_sum%>">
		               
		                
		                
	                <%if x <> 0 then%>
	                <tr>
		                <td bgcolor="#ffffff" colspan="18"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
                    </tr>
	                <%end if%>
	                <tr bgcolor="<%=bgthis%>">
		                <td height=20>&nbsp;</td>
		                <td><b><%=oRec("jobnavn")%> (<%=oRec("jobnr")%>)</b><br>
		                <%=oRec("kkundenavn")%> (<%=oRec("kkundenr")%>)</td>
                		
		                <%
		                if lto = "bowe" then
                		
					                '*** KA nr ****
					                kanr_st = instr(oRec("jobnavn"), "(")
					                kanr_sl = instr(oRec("jobnavn"), ")")
                					
					                if kanr_st <> 0 AND kanr_sl <> 0 then
					                kanr_length = (kanr_sl - kanr_st) 
					                kanr_str = mid(oRec("jobnavn"), kanr_st+1, kanr_length-1)
					                kanr = kanr_str
					                else
					                kanr = ""
					                end if
                					
					                '** bowekode **
					                bowekode_st = instr(oRec("jobnavn"), ")")
                					
					                if bowekode_st <> 0 then
					                bowekode_length = 2
					                bowekode = mid(oRec("jobnavn"), bowekode_st+1, 3)
					                else
					                bowekode = ""
					                end if
                					
                					
		                strExport = strExport & "xx99123sy#z" & oRec("jobnavn") &";"&oRec("jobnr")&";"&kanr&";"&bowekode&";"&oRec("kkundenavn")&";"&oRec("kkundenr")&";"
		                else
		                strExport = strExport & "xx99123sy#z" & oRec("jobnavn") &";"&oRec("jobnr")&";"&oRec("kkundenavn")&";"&oRec("kkundenr")&";"
		                end if%>
                		
		                <td>
		                <%if len(oRec("forbrugsdato")) <> 0 then%>
		                <%=formatdatetime(oRec("forbrugsdato"), 2)%>
		                <%
		                strExport = strExport & formatdatetime(oRec("forbrugsdato"), 2) &";" 
		                else
		                strExport = strExport &";"  
		                end if
		                %>
		                <br><font class="megetlillesilver"><%=oRec("medarbejdernavn")%> (<%=oRec("mnr")%>) - <%=oRec("init")%></font></td>
                		
		                <%
		                strExport = strExport & oRec("medarbejdernavn")&";"&oRec("mnr")&";"&oRec("init")&";"
		                %>
                		
		                <td style="padding-left:5px;">
		                <%if len(oRec("gnavn")) <> 0 then %>
		                <%=oRec("gnavn")%> (<%=oRec("av") %> %)
		                <%end if %>
		                &nbsp;
		                </td>
		                <td style="padding-left:5px;"><b><%=oRec("navn")%>&nbsp;&nbsp;(<%=oRec("varenr")%>)</b></td>
                		
		                <%
		                strExport = strExport & oRec("gnavn")&" ("& oRec("av") &"%);"&oRec("navn")&";"&oRec("varenr")&";"
		                %>
                		
		                <%if editok = 1 AND print <> "j" then%>
		                <td><input type="text" name="FM_antal" id="FM_antal" size="2" style="font-size:9px;" value="<%=oRec("antal")%>">&nbsp;<%=oRec("enhed")%>&nbsp;</td>
		                <td><input type="image" src="../ill/soeg-knap.gif"></td>
		                <%else%>
		                <td><%=oRec("antal")%>&nbsp;<%=oRec("enhed")%></td>
		                <td>&nbsp;</td>
		                <%end if%>
                		
		                <td align=right><%=oRec("palager")%>&nbsp;</td>
		                <td align=right><%=oRec("minlager")%>&nbsp;</td>
                		
		                <%
		                strExport = strExport & oRec("antal")&";"&oRec("enhed")&";"
		                %>
                		
		                <td align=right style="padding-right:3px;"><b><%=formatnumber(oRec("matkobspris"),2)%></b></td>
		                <td align=right style="padding-right:3px;"><%=formatnumber(oRec("matsalgspris"),2)%></td>
		                <td align=center><%=oRec("valutakode") %></td>
		                
		            <td align=right style="padding-right:5px;"><%=oRec("mfkurs") %></td>
            		<td align=right style="padding-right:3px;">
		                <%kobsprisialt = formatnumber(oRec("matkobspris")/1 * oRec("antal")/1 * oRec("mfkurs")/100 , 2) %>
		                <b><%=kobsprisialt %></b></td>
		                <%
		                strExport = strExport & formatnumber(oRec("matkobspris"), 2) &";"
		                kprisTot = kprisTot + (kobsprisialt/1)
		                %>
		              
		              <td align=right style="padding-right:3px;">
		                <%salgsprisialt = formatnumber(oRec("matsalgspris")/1 * oRec("antal")/1 * oRec("mfkurs")/100 , 2) %>
		                <%=salgsprisialt %>
		                <%
		                strExport = strExport & formatnumber(oRec("matsalgspris"), 2) &";"& oRec("valutakode") & ";" & oRec("mfkurs") & ";"& kobsprisialt &";"&salgsprisialt&";DKK;"
		                sprisTot = sprisTot + (salgsprisialt/1)
		                %>
		                </td>
		                <td align=center>DKK</td>
            		
                		
		                <td>&nbsp;</td>
		                <td>&nbsp;</td>
	                </tr>
	                 <tr>
		                <td style="border-bottom:1px #999999 dashed;" colspan="18"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	                </tr>
	                
	                </form>
	                <% 
                	
	                strAntalFM = strAntalFM &"<input id='FM_antal_b_"&oRec("matid")&"' name='FM_antal_b_"&oRec("matid")&"' type=hidden value='"&oRec("antal")&"' />" 
                	
	                x = x + 1
	                oRec.movenext
	                wend
                	
	                if x = 0 then 
	                %>
	              
	                <tr bgcolor="#ffffff">
		                <td height=50 style="border-left:1px #003399 solid;">&nbsp;</td>
		                <td colspan=16><font color="red"><b>!</b></font>&nbsp;Der er ikke registreret materialer der matcher de valgte søgekriterier.</td>
		                <td style="border-right:1px #003399 solid;">&nbsp;</td>
	                </tr>
	                <%else%>
	                <tr bgcolor="#ffdfdf">
		                <td height="50">&nbsp;</td>
		                <td colspan=12 align=right><b>Total:</b></td>
		                <td align=right>
		                <%if len(kprisTot) <> 0 then%>
		                <b><u><%=kprisTot%></u></b>
		                <%end if%></td>
                		
                		
		                <td align=right>
		                <%if len(sprisTot) <> 0 then%>
		                <b><u><%=sprisTot%></u></b>
		                <%end if%></td>
                		
                		
		                <td><b>DKK</b></td>
		                <td>
                            &nbsp;</td>
		                <td>&nbsp;</td>
	                </tr>
	                <%end if%>	
	               
	                </table>
	                </div>
	
	<%
	case 1  '** Total forbrug ***'
	%>
		
		        <% 
		        tTop = 20
	            tLeft = 0
	            tWdth = 880
            	
            	
	            call tableDiv(tTop,tLeft,tWdth)
	            %>
	            
		        <table cellspacing=0 cellpadding=2 width=100% border=0>
		        <tr bgcolor="#5582d2">
			        <td style="padding:5px;" class=alt><b>Lager / Gruppe</b> (Ava. %)</td>
			        <td style="padding:5px;" class=alt><b>Navn</b></td>
			        <td style="padding:5px;" class=alt><b>Varenr</b></td>
			        <td style="padding:5px;" class=alt align=right><b>Antal brugt i per.</b></td>
			        <td style="padding:5px;" class=alt align=right><b>På lager</b></td>
			        <td style="padding:5px;" class=alt align=right><b>Min. lager</b></td>
		        </tr>
        		
        		
		        <%
        		
		        strExport = "Gruppe;Navn;Varenr;Antal;Enhed;På lager;Min. lager;"
        	
        		
		        strSQL = "SELECT mf.id, mf.matvarenr AS varenr, mg.navn AS gnavn, mf.matenhed AS enhed, "_
		        &" mf.matnavn AS navn, sum(mf.matantal) AS antal, mf.dato, mf.editor, "_
		        &" mf.matkobspris, mf.matsalgspris, mf.jobid, mf.matgrp, mf.usrid, mf.forbrugsdato, "_
		        &" m.antal AS palager, m.minlager, m.id, mf.matid, mg.av "_
		        &" FROM materiale_forbrug mf"_
		        &" LEFT JOIN materialer m ON (m.id = mf.matid) "_
		        &" LEFT JOIN materiale_grp mg ON (mg.id = mf.matgrp) "_
		        &" WHERE "&jobKri&" AND ("& medarbSQlKri &") "& strDatoKri &" GROUP BY mf.navn, mf.matgrp, mf.matgrp ORDER BY mg.gnavn, mf.matnavn"
        		
		        'Response.write strSQL
		        'Response.flush
		        x = 0
		        oRec.open strSQL, oConn, 3
		        while not oRec.EOF
        		
        		
		        'if oRec("minlager") > oRec("palager") then
		        'bg = "#ffdfdf"
		        'else
		        'bg = "#ffffff"
		        'end if		
		        
		        select case right(x,1)
		        case 0,2,4,6,8
		        bgcol = "#ffffff"
		        case else
		        bgcol = "#EFf3FF"
		        end select
		        			
        								
		        %>
		        <tr bgcolor="<%=bgcol %>">
			        <td height=20 style="border-top:1px #cccccc dashed; padding-left:5px;">	
	            <%if len(oRec("gnavn")) <> 0 then %>
		        <%=oRec("gnavn")%> (<%=oRec("av") %> %)
		        <%end if %>
		        &nbsp;</td>
			        <td style="border-top:1px #cccccc dashed; padding-left:5px;" class="<%=tdclass%>"><b><%=oRec("navn")%></b></td>
			        <td style="border-top:1px #cccccc dashed; padding-left:5px;" class="<%=tdclass%>"><%=oRec("varenr")%></td>
			        <td style="border-top:1px #cccccc dashed; padding-right:5px;" align=right class="<%=tdclass%>"><%=oRec("antal")%>&nbsp;
			        <%if len(oRec("enhed")) <> 0 then%>
			        <%enh = oRec("enhed")%>
			        <%else 
			        enh = "Stk."
			        end if %>
        			
			        <%=enh%></td>
        			
			        <%if oRec("varenr") <> "0" then %>
			        <td style="border-top:1px #cccccc dashed; padding-right:5px;" align=right class="<%=tdclass%>"><%=oRec("palager")%>&nbsp;<%=oRec("enhed")%>
        			
			        <%if oRec("minlager") > oRec("palager") then
			        Response.write "&nbsp;<font class=roed><b>!</b>"
			        end if
			        %>
			        </td>
			        <td style="border-top:1px #cccccc dashed; padding-right:5px;" align=right><%=oRec("minlager")%>&nbsp;<%=oRec("enhed")%></td>
		            <%else %>
		            <td style="border-top:1px #cccccc dashed; padding-right:5px;" align=right class="<%=tdclass%>">-</td>
			        <td style="border-top:1px #cccccc dashed; padding-right:5px;" align=right>-</td>
		            <%end if %>
		        </tr>
		        <%
		        strExport = strExport & "xx99123sy#z" & oRec("gnavn") &";"&oRec("navn")&";"&oRec("varenr")&";"&oRec("antal")&";"&enh&";"&oRec("palager")&";"&oRec("minlager")&";"
        		
        		
		        x = x + 1
		        oRec.movenext
		        wend
        		
		        oRec.close
        		
        		
		        if x = 0 then%>
        		
		        <tr bgcolor="#ffffff">
			        <td colspan=8 height=50 style="padding:20px;">
			        <font color="red"><b>!</b></font>&nbsp;Der er ikke registreret materialer der matcher de valgte søgekriterier.</td>
		        </tr>
		        <%end if%>
		        <tr bgcolor="#5582d2">
			        <td colspan=8>&nbsp;</td>
		        </tr>
		        </table>
		        </div>
        	
	<% 
	case 2 '*** Som udgiftsbilag ***'
	
	
	            tTop = 0
	            tLeft = 0
	            tWdth = mainTableWth
            	
            	
	            call tableDiv(tTop,tLeft,tWdth)
            	
	            %>
	            <table width=100% cellspacing=0 cellpadding=2 border=0>
            		<form method="post" action="materiale_stat.asp?func=opdafregnet&hidemenu=<%=hidemenu %>">
		            <tr bgcolor="#5582d2">
            		       
		                <td class=alt><b><%=tsa_txt_230 %></b></td>
		                
		                
		                <td class=alt><b><%=tsa_txt_231%></b><br />
		                <%=tsa_txt_241 %></td>
		                <td class=alt><b><%=tsa_txt_183 %></b></td>
            			
            			<td class=alt><b><%=tsa_txt_244 %></b><br />
            			<%=tsa_txt_243 %></td>
            			
            			<td class=alt><%=tsa_txt_077%></td>
            			
            			
            			
			            <td align=right class=alt><b><%=tsa_txt_202 %></b></td>
			            <td class=alt><b><%=tsa_txt_245%></b></td>
			            <td style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_209 %></b></td>
			            <td align=center style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_217 %></b></td>
            			
            			
            			
            			
			            <td align=right style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_219 %></b></td>
			            <td align=right style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_246%></b></td>
			            <td align=right style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_247%></b></td>
            			<td align=right style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_248%></b></td>
            			<td align=right style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_246%></b></td>
            			
            			
			            <td class=alt><%=tsa_txt_221 %></td>
			            <td class=alt><%=tsa_txt_323 %>?</td>
			            <td class=alt><%=tsa_txt_253 %></td>
            			
            			
		            </tr>
	            <%
	            
	            strExport = strExport & tsa_txt_230 &";"& tsa_txt_231 &";"& tsa_txt_241 &";"& tsa_txt_183 &";"_
	            &""&tsa_txt_243&";"&tsa_txt_249&";"&tsa_txt_244&";"&tsa_txt_250&";"& tsa_txt_077 &";"&tsa_txt_321&";"& tsa_txt_202 &";"&tsa_txt_245&";"& tsa_txt_209 &";"&tsa_txt_217&";"&tsa_txt_219&";"_
	            &""&tsa_txt_246&";"&tsa_txt_247&";"&tsa_txt_248&";"&tsa_txt_246&";"&tsa_txt_323&";"&tsa_txt_253&";"  
	            ' "xx99123sy#z" 
	            
	            
	            if useSog = 1 then
	            sqlWh = "mf.bilagsnr LIKE '"& sogliste &"'"
	            else
	            sqlWh = medarbSQlKri '"usrid = "& usemrn 
	            end if
	            
	            if showonlypers = 1 then
	            strSQLper = " AND personlig = 1"
	            else
	            strSQLper = ""
	            end if
            	
            	
            	if showonlyikkeafreg = 1 then
            	SQLafr = " AND mf.afregnet = 0"
            	else
            	SQLafr = ""
            	end if
            	
            	
            	if showonlyikkegk = 1 then
            	SQLgk = " AND godkendt = 0"
            	else
            	SQLgk = ""
            	end if
            	
            	totsum = 0
            	
	            strSQLmat = "SELECT m.mnavn AS medarbejdernavn, mf.id AS mfid, mf.matvarenr AS varenr, mg.navn AS gnavn, mf.matenhed AS enhed, "_
	            &" mf.matnavn AS navn, mf.matantal AS antal, mf.dato, mf.editor, "_
	            & "mf.matkobspris, mf.matsalgspris, mf.jobid, mf.matgrp, "_
	            &" mf.usrid, mf.forbrugsdato, j.id, j.jobnr, j.jobnavn, "_
	            &" mg.av, f.fakdato, k.kkundenavn, mf.afregnet, "_
	            &" k.kkundenr, mf.valuta, mf.intkode, mf.bilagsnr, v.valutakode, mf.kurs AS mfkurs, "_
	            &" mf.personlig, m.init, m.mnr, mf.godkendt, mf.gkaf "_
	            &" FROM materiale_forbrug mf"_
	            &" LEFT JOIN materiale_grp mg ON (mg.id = matgrp) "_
	            &" LEFT JOIN medarbejdere m ON (mid = usrid) "_
	            &" LEFT JOIN job j ON (j.id = mf.jobid) "_
	            &" LEFT JOIN fakturaer f ON (f.jobid = mf.jobid AND f.faktype = 0) "_
	            &" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
	            &" LEFT JOIN valutaer v ON (v.id = mf.valuta) "_
	            &" WHERE ("& sqlWh &") "& strDatoKri &" AND "&jobKri&" "& intKodeSQLkri &" "& strSQLper &" "& SQLafr &" "& SQLgk &" GROUP BY mf.id ORDER BY mf.forbrugsdato DESC, f.fakdato DESC, mf.id DESC" 	
            	
	            'mf.jobid = "& id &"
	            'response.write strSQLmat
	            'Response.flush
            	
	            x = 0
	            oRec.open strSQLmat, oConn, 3
	            while not oRec.EOF
            	 
            	 
	             if len(oRec("fakdato")) <> 0 then
	             fakdato = oRec("fakdato")
	             else
	             fakdato = "01/01/2002"
	             end if
            	 
	             if oRec("mfid") = cint(lastId) then
	             bgthis = "#FFFF99"
	             else
	                 select case right(x, 1) 
	                 case 0,2,4,6,8
	                 bgthis = "#EFf3FF"
	                 case else
	                 bgthis = "#FFFFff"
	                 end select
	             end if
	             
	             strExport = strExport & "xx99123sy#z"
	             
	             %>
	            <tr bgcolor="<%=bgthis %>">
            	    
	                <td class=lille style="border-bottom:1px #cccccc dashed; white-space:nowrap;">
                    <%if len(oRec("bilagsnr")) <> 0 then%>
		            <%=oRec("bilagsnr") %> 
		            <%end if %>
            	        &nbsp;
                        </td>
                        
                        <%
                          strExport = strExport & oRec("bilagsnr") &";"
	           
                        %>
                        
                        <td class=lille style="border-bottom:1px #cccccc dashed; padding-right:5px;">
                        <%select case oRec("intkode")
		            case 0
	                intKodeTxt = "-"
		            case 1
		            intKodeTxt = tsa_txt_239
		            case 2
		            intKodeTxt = tsa_txt_240
		            'case 3
		            'intKode = tsa_txt_241
		            end select %>
            		
		            
                        
                        <%if intKodeTxt <> "-" then %>
		                <b><%=intKodeTxt%></b> <br />
		                <%end if %>
                            
            	    
            	            <%
                          strExport = strExport & intKodeTxt &";"
	           
                        %>
                        
                        <%if oRec("personlig") <> 0 then %>
		                    <font class=megetlillesilver><%=tsa_txt_241 %></font>
		                    <%
		                    
		                    end if %>
		                    
		                  
		                    
		<%
                          strExport = strExport & oRec("personlig") &";"
	           
                        %>
		
		&nbsp;</td>
            	    
            	    
		            <td width=60 class=lille style="border-bottom:1px #cccccc dashed;">
		            <%if len(oRec("forbrugsdato")) <> 0 then%>
		            <b><%=formatdatetime(oRec("forbrugsdato"), 2)%></b>
		            <%end if%>
		           
	                </td>
	                 <%
                          strExport = strExport & formatdatetime(oRec("forbrugsdato"), 2) &";"
	           
                        %>
	                
	                
	                <td class=lille style="border-bottom:1px #cccccc dashed;"><b><%=oRec("jobnavn")%></b>&nbsp;(<%=oRec("jobnr")%>) <br />
	                <%=oRec("kkundenavn") %> (<%=oRec("kkundenr") %>)</td>
            		
	                
	                <%
                          strExport = strExport & oRec("kkundenavn") & ";"& oRec("kkundenr") &";" & oRec("jobnavn") & ";" & oRec("jobnr") & ";"
	           
                        %>
	                
	                <td width=100 class=lille style="border-bottom:1px #cccccc dashed;"><%=oRec("medarbejdernavn")%> (<%=oRec("mnr") %>)
	                <%if len(trim(oRec("init"))) <> 0 then %>
	                - <%=oRec("init") %>
	                <%end if%>
	                </td>
            	     
            	     <%
                          strExport = strExport & oRec("medarbejdernavn") &";"& oRec("init") &";"& oRec("antal") &";"
	           
                        %>
            	     
		            <td align=right style="padding-right:5px; border-bottom:1px #cccccc dashed;" class=lille><b><%=oRec("antal")%></b>&nbsp;</td>
		            
		            <td style="border-bottom:1px #cccccc dashed;"> <%if len(oRec("enhed")) <> 0 then
		            enh = oRec("enhed")
		            else
		            enh = tsa_txt_222 '"Stk."
		            end if %>
            		
		            <%=enh%></td>
            	    
            	   
            	    
	                <td class=lille style="border-bottom:1px #cccccc dashed;"><b><%=oRec("navn")%></b>&nbsp;</td>
		            <td align=right style="border-bottom:1px #cccccc dashed; padding-right:5px;" class=lille><%=oRec("varenr")%>
                        &nbsp;</td>
            	    
            	     <%
            	    strNavn = replace(oRec("navn"), "'", "")
            	    strNavn = replace(strNavn, chr(34), "")
                
            	    
                          strExport = strExport & enh &";"& strNavn &";" & oRec("varenr") & ";" & formatnumber(oRec("matkobspris"), 2) & ";"& oRec("valutakode") &";"& oRec("mfkurs") &";"
	           
                        %>
		            
		            <td align=right style="border-bottom:1px #cccccc dashed; padding-right:5px;" class=lille>
	                <b><%=formatnumber(oRec("matkobspris"), 2) %></b>
		            </td>
		            
		            <td align=right style="border-bottom:1px #cccccc dashed; padding-right:5px;" class=lille><%=oRec("valutakode") %></td>
            		
            		<td align=right style="border-bottom:1px #cccccc dashed; padding-right:5px;" class=lille><%=oRec("mfkurs") %></td>
            		<td align=right style="border-bottom:1px #cccccc dashed; padding-right:5px;" class=lille>
            		<%udlgSum = formatnumber(oRec("matkobspris")/1 * oRec("antal")/1 * oRec("mfkurs")/100 , 2) %>
            		<b><%=udlgSum %></b>
            		</td>
            		
            		<%
                          strExport = strExport & udlgSum &";DKK;"
	           
                        %>
            	     
            		
            		<%totsum = totsum + udlgSum %>
            		
            		<td align=right style="border-bottom:1px #cccccc dashed; padding-right:5px;" class=lille>DKK</td>
            		
	                <td style="border-bottom:1px #cccccc dashed;">
            	    
	                <%
	                   '*** Er uge alfsuttet af medarb, er smiley og autogk slået til
                            erugeafsluttet = instr(afslUgerMedab(usemrn), "#"&datepart("ww", oRec("forbrugsdato"),2,2)&"_"& datepart("yyyy", oRec("forbrugsdato")) &"#")
                            
                            'Response.Write "erugeafsluttet --" & erugeafsluttet  &"<br>"
                            'Response.flush
                            
                          
                             if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                            (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) = year(now) AND DatePart("m", oRec("forbrugsdato")) < month(now)) OR _
                            (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) = 12)) OR _
                            (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) <> 12) OR _
                            (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", oRec("forbrugsdato")) > 1))) then
                          
                            ugeerAfsl_og_autogk_smil = 1
                            else
                            ugeerAfsl_og_autogk_smil = 0
                            end if 
            				
            				
            				
				            if (ugeerAfsl_og_autogk_smil = 0 _
				            OR ugeerAfsl_og_autogk_smil = 1 AND level = 1) _
				            AND cdate(fakdato) < cdate(oRec("forbrugsdato")) AND print <> "j" then
            				
			            %>
	                &nbsp;<a href="materialer_indtast.asp?id=<%=id%>&func=slet&matregid=<%=oRec("mfid")%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aftid%>"><img src="../ill/slet_16.gif" alt="<%=tsa_txt_221 %>" border=0 /></a>
	                <%else%>
	                &nbsp;
	                <%end if %>
            	    
	                </td>
	                
	                <td style="border-bottom:1px #cccccc dashed; padding-top:5px;" align=center>
	                
	                 <%if cint(oRec("godkendt")) = 1 then
                        gkendt1CHK = "SELECTED"
                        gkendt0CHK = ""
                        gkendt = "Ja"
                        gkAf = oRec("gkaf")
                        else
                        gkendt1CHK = ""
                        gkendt0CHK = "SELECTED"
                        gkendt = "Nej" 
                        gkAf = "&nbsp;"
                        end if 
            
             strExport = strExport & gkendt &";"
            
            if print <> "j" then%>
           
         
            <select id="Select2" name="mfidgk_<%=oRec("mfid")%>" style="width:50px; font-size:9px;">
                <option value="0" <%=gkendt0CHK %>>Nej</option>
                <option value="1" <%=gkendt1CHK %>>Ja</option>
            </select><br />
            <font class="megetlillesilver"><%=left(gkAf, 10)%>.</font>
           
            <%else 
                if oRec("godkendt") = 1 then
                Response.Write "<b>Ja</b><br><font class=megetlillesilver>"& gkAf
                else
                Response.Write "Nej"
                end if
            end if
            %>
	                
	                </td>
	                
	                
	                <td style="border-bottom:1px #cccccc dashed; padding-top:5px;" align=center>
	                
	                 <%if oRec("afregnet") = 1 then
            efbCHK = "SELECTED"
            ikbCHK = ""
            afreg = "Ja;"
            else
            ikbCHK = "SELECTED"
            efbCHK = ""
            afreg = "Nej;" 
            end if 
            
             strExport = strExport & afreg &";"
            
            if print <> "j" then%>
           
           <!--
           <input id="mfid_<%=oRec("mfid")%>" name="mfid_<%=oRec("mfid")%>" value="1" type="radio" <%=efbCHK %> /> <b>Ja</b><br />
           <input id="mfid_<%=oRec("mfid")%>" name="mfid_<%=oRec("mfid")%>" value="0" type="radio" <%=ikbCHK %> /> Nej
            -->
            
            <select id="mfid_<%=oRec("mfid")%>" name="mfid_<%=oRec("mfid")%>" style="width:50px; font-size:9px;">
                <option value="0" <%=ikbCHK %>>Nej</option>
                <option value="1" <%=efbCHK %>>Ja</option>
            </select><br />&nbsp;
            
            <input id="erafregnet" name="erafregnet" type="hidden" value="<%=oRec("mfid")%>" />
           
            <%else 
                if oRec("afregnet") = 1 then
                Response.Write "<b>Ja</b>"
                else
                Response.Write "Nej"
                end if
            end if
            %>
	                
	                </td>
            		
	            </tr>
	           
	            <%
	            x = x + 1
	            oRec.movenext
	            wend
            	
	            oRec.close
	            %>
	            <tr>
	                <td colspan=13>
                        &nbsp;</td>
	                <td align=right style="padding-right:5px;" class=lille><b><%=formatnumber(totsum,2) %></b></td>
	                <td align=right style="padding-right:5px;" class=lille>DKK</td>
	                <td colspan=2>
                        &nbsp;<%if print <> "j" then %>
                        <input id="Submit1" type="submit" value="<%=tsa_txt_144 %>" />
                        <%end if %>
                        </td>
	            </tr>
	             </form>
	            </table>
	            
	           </div>
	           
	           <%if print <> "j" then %>
	           <br />
	            <br />
	            <form method="post" action="materiale_stat.asp?func=afregnalle&hidemenu=<%=hidemenu %>" >
	            Marker alle bilags poster på følgende bilag som afregnet: <br />
                <input id="afregnalle_id" name="afregnalle_id" type="text" size=20 value="<%=sogliste %>" />
                &nbsp;<input id="Submit2" type="submit" value="Afregn --> " />
                </form>
                <%end if %>
	
	
	
	<%end select '* pr job eller sum eller som udgiftsbilag%>
	
	
	
	
	
	
	
	
	
	
	
	<%'if x <> 0 then%>
	
	
	
	<%if print <> "j" then%>
	
	<br><br>
	
	<%if visprjob_ell_sum = 0 then %>
	<form action="materialer_ordrer.asp?func=opr" method="post" target="_blank"> 
			<input type="hidden" name="matids" id="matids" value="<%=strMatids%>">
			<%=strAntalFM%>
		    <input type="submit" value="Opret mat. ordre >>">
			
			</form>
	<%end if %>
	        
	        
	        
	         <%
	            
	           
	'*** Generalt link med alle variabler ***
	link = "&FM_visprjob_ell_sum="&visprjob_ell_sum&"&FM_job="&jobid&"&FM_jobans="&jobans&""_
	&"&FM_jobans2="&jobans2&"&FM_kundeans="&kundeans&"&FM_kunde="&kundeid&"&FM_medarb="&selmedarb&""_
	&"&FM_radio_projgrp_medarb="&progrp_medarb&"&FM_intkode="&intKode&"&FM_progrupper="&progrp&""_
	&"&FM_kundejobans_ell_alle="&visKundejobans&"&FM_jobsog="&jobSogValPrint

            
                ptop = 72
                pleft = 810
                pwdt = 165

                call eksportogprint(ptop,pleft, pwdt)
                %>

                <form action="materiale_stat_eksport.asp" method="post" name=theForm2 onsubmit="BreakItUp2()" target="_blank"> <!--  -->
			    <input type="hidden" name="datointerval" id="datointerval" value="<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>">
			    <input type="hidden" name="txt1" id="txt1" value="">
			    <input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=strExport%>">
			    <input type="hidden" name="txt20" id="txt20" value="">

                <tr>
                    <td align=center><input type="image" src="../ill/export1.png">
                    </td>
                    </td><td>.csv fil eksport</td>
                    </tr>
                    <tr>
                    <td align=center>
                   <a href="materiale_stat.asp?menu=stat&print=j&id=<%=id%><%=link%>&hidemenu=<%=hidemenu %>" target="_blank"  class='vmenu'>
                   &nbsp;<img src="../ill/printer3.png" border=0 alt="" /></a>
                </td><td>Print version</td>
                   </tr>
                   
	                </form>
                   </table>
                </div>
                
                <%if cint(visprjob_ell_sum) = 1 OR cint(visprjob_ell_sum)  = 0 then %>
                <div style="position:relative; background-color:#ffffe1; visibility:visible; border:1px red dashed; top:40px; padding:20px; width:400px;"><img src="../ill/ac0005-24.gif" width="24" height="24" alt="" border="0">&nbsp;<b>Side note(r):<br></b>
	            Hvis du er Jobansvarlig eller Administrator har du mulighed for at ændre i antal. For at slette en materiale-registrering tast <b>0 (nul) i antal</b> og opdater. Antal på lager vil samtidig blive opdateret.
	            </div>
	            <%end if %>
	            <br><br>&nbsp;
                
                <%else%>

                <% 
                Response.Write("<script language=""JavaScript"">window.print();</script>")
                %>
                <%end if%>
	
	
	<%'end if%><!-- x <> 0 -->
	
	<br><br>&nbsp;
	</div>
	<%end select%>
	
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
