<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--include file="inc/dato2.asp"-->


<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<% 
	errortype = 5
	call showError(errortype)
	else
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	%>
	<script>
	
	function njob(){
	//alert("her")
	document.getElementById("FM_job").value = "0"
	document.getElementById("FM_aftale").style.backgroundColor = "#ffff99"
	document.getElementById("FM_job").style.backgroundColor = "#ffffff"
	document.getElementById("jobelaft2").checked = true;
	}
	
	function naft(){
	//alert("her")
	document.getElementById("FM_aftale").value = "0"
	document.getElementById("FM_aftale").style.backgroundColor = "#ffffff"
	document.getElementById("FM_job").style.backgroundColor = "#ffff99"
	document.getElementById("jobelaft").checked = true;
	}
	
	function nextstep2() {
    //kid = document.getElementById("FM_kunde").value
    //type = document.getElementById("FM_type").value
    //jobid = document.getElementById("FM_job").value
   //aftaleid = document.getElementById("FM_aftale").value
    //jobelaft = document.getElementById("jobelaft").value
    
    //strDag = document.getElementById("FM_start_dag").value
    //strMrd = document.getElementById("FM_start_mrd").value
    //strAar = document.getElementById("FM_start_aar").value
    //strDag_slut = document.getElementById("FM_slut_dag").value
    //strMrd_slut = document.getElementById("FM_slut_mrd").value
    //strAar_slut = document.getElementById("FM_slut_aar").value
    
    //window.top.frames['erp2_1'].location.href = "erp_opr_faktura_kontojob.asp"
	//window.top.frames['erp2_2'].location.href = "erp_opr_faktura.asp?FM_usedatokri=1&FM_job="+jobid+"&FM_aftale="+aftid+"&jobelaft="+jobelaft+"&FM_kunde="+kid+"&FM_start_dag="+strDag+"&FM_start_mrd="+strMrd+"&FM_start_aar="+strAar+"&FM_slut_dag="+strDag_slut+"&FM_slut_mrd="+strMrd_slut+"&FM_slut_aar="+strAar_slut   
    window.top.frames['erp3'].location.href = "erp_opr_faktura_blank.asp"
    }	
    </script>
    <%
	
	func = request("func")
    thisfile = "erp_opr_faktura_kontojob"
    print = request("print")
    
    
    if len(trim(request("FM_kunde"))) <> 0 then
    kid = request("FM_kunde")
  
        if kid <> 0 then
        kidSQL = " jobknr = "& kid
        aftKidSQL = " kundeid = "& kid 
        else
        kidSQL = " jobknr <> "& kid 
        aftKidSQL = " kundeid <> "& kid 
        end if 
       
    else
        if len(trim(request.Cookies("erp")("kid"))) <> 0 then
        kidSQL = " jobknr = "& request.Cookies("erp")("kid")
        aftKidSQL = " kundeid = "& request.Cookies("erp")("kid") 
        else
        kid = 0
        kidSQL = " jobknr <> "& kid 
        aftKidSQL = " kundeid <> "& kid 
        end if
    end if
    
    Response.Cookies("erp")("kid") = kid
    
    if len(request("FM_job")) <> 0 then
        jobid = request("FM_job")
    else
    
        if len(trim(request.cookies("erp")("jid"))) <> 0 then
        jobid = request.cookies("erp")("jid")
        else
        jobid = 0
        end if
    
    end if
    
    if len(request("FM_aftale")) <> 0 then
    aftaleid = request("FM_aftale")
    else
   
    if len(trim(request.cookies("erp")("aid"))) <> 0 then
    aftaleid = request.cookies("erp")("aid")
    else
    aftaleid = 0
    end if
    
    end if
    
    
    
    
    Response.Cookies("erp").expires = date + 60
    
	%>
	
	
	
	
	<div id="sindhold" style="position:absolute; left:10px; top:0px; visibility:visible;">
	
	<%
	'**********************************************************
	'**************** Orpet / red Faktura Step 2 **************
	'**********************************************************
	 %>
	

	    
	   <!-- <table cellspacing=2 cellpadding=0 border=0>
	    <tr><td>
	    <a href="erp_opr_faktura_fs.asp" target="_top" class="rmenu">1) - Vælg kontakt (debitor) </a>
	    </td></tr>
	    <tr><td>
	    <a href="#" target="erp2_1" class="rmenu"><u>2) - Vælg job / aftale og datointerval</u></a>
	     </td></tr>
	    </table>-->
	 
	<img src="../ill/blank.gif" height="4" width="1" border="0"/><br />
	 <table cellspacing=0 cellpadding=0 border=0 width="275">
	 <form action="erp_opr_faktura.asp" method="POST" target="erp2_2">
      <input id="FM_kunde" name="FM_kunde" type="hidden" value="<%=kid %>" />
      <input id="FM_usedatokri" name="FM_usedatokri" type="hidden" value="1"/>
	
	 
	
	  <tr>
	      <td colspan=2 bgcolor="#ffffe1" style="padding:10px 10px 10px 10px; border:1px #8caae6 solid; border-bottom:0px;"><b>Vælg job eller aftale:</b><br />
	   
	    <%strSQL = "SELECT id, jobnavn, jobnr, jobstatus FROM job WHERE "& kidSQL &" AND jobstatus = 1 ORDER BY jobnavn"
	    'Response.Write strSQL
	     %>
	     <!--<input id="jobelaft" name="jobelaft" type="radio" value="1" <%=jobelaft1CHK %> onclick="naft()" />-->
	     <select id="FM_job" name="FM_job" style="width:225px; font-family:verdana; font-size:10px;" onchange="naft()">
	     <option value="0">Vælg job..</option>
	     <%
	        oRec.open strSQL, oConn, 3
            while not oRec.EOF
            if cint(jobid) = oRec("id") then
            jidSel = "SELECTED"
            else
            jidSel = ""
            end if %>
            <option value="<%=oRec("id") %>" <%=jidSel %>><%=oRec("jobnavn") %> (<%=oRec("jobnr") %>)</option>
            <%
            oRec.movenext
            wend
            oRec.close%>
           </select>
           
          <img src="../ill/blank.gif" height="4" width="1" border="0"/><br />
             
	    <%strSQL = "SELECT id, navn, aftalenr, status FROM serviceaft WHERE "& aftKidSQL &" AND status = 1"
	    'Response.Write strSQL
	     %>
	     <!--<input id="jobelaft2" name="jobelaft" type="radio" value="2" <%=jobelaft2CHK %> onclick="njob()" />-->
	     <select id="FM_aftale" name="FM_aftale" style="width:225px; font-family:verdana; font-size:10px;" onchange="njob()">
	     <option value="0">Vælg aftale..</option>
	     <%
	        oRec.open strSQL, oConn, 3
            while not oRec.EOF
             if cint(aftaleid) = oRec("id") then
            aidSel = "SELECTED"
            else
            aidSel = ""
            end if%>
            <option value="<%=oRec("id") %>" <%=aidSel %>><%=oRec("navn") %> (<%=oRec("aftalenr") %>)</option>
            <%
            oRec.movenext
            wend
            oRec.close%>
           </select></td>
	 </tr>
	
	    <tr>
        <td bgcolor="#ffffe1" style="padding:10px 10px 10px 10px; border:1px #8caae6 solid; border-right:0px; border-top:0px;">
	    <b>Dato interval:</b> (slutdato = fakturadato)<br />
	    <!--#include file="inc/weekselector_s.asp"-->
	  
	 
	  </td><td bgcolor="#ffffe1" valign=top style="padding:42px 4px 0px 0px; border:1px #8caae6 solid; border-left:0px; border-top:0px;">
           <input id="Button1" type="image" src="../ill/pilstorxp.gif" onClick="nextstep2()" />
	    </td></tr>
	     </form>
	</table>
	

	</div>
	
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->