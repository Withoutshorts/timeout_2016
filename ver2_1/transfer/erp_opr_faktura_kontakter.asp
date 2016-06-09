<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato2.asp"-->


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
	
	<script>
	function nextstep1() {
	kid = document.getElementById("FM_kunde").value
    //window.location.href = "default.asp?clicked="+dd_clicked
    window.top.frames['erp2_1'].location.href = "erp_opr_faktura_kontojob.asp?fm_kunde="+kid
	//window.top.frames['erp2_2'].location.href = "erp_opr_faktura.asp?fm_kunde="+kid+"&func=viskontakt"   
    }	
    
    function rensSog() {
    document.getElementById("FM_sog").value = ""
    }	
    </script>
	
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
	func = request("func")
    thisfile = "webblik_tilfakturering.asp"
    print = request("print")
    
    if len(trim(request("FM_kunde"))) <> 0 then
    kid = request("FM_kunde")
    else
        if len(trim(request.Cookies("erp")("kid"))) <> 0 then
        kid = request.Cookies("erp")("kid")
        else
        kid = 0
        end if
    end if
    
    
   
   
    
    
    if len(request("FM_sog")) <> 0 then
    sogKri = request("FM_sog")
    kSQLkri = "AND kkundenavn LIKE '"& sogKri &"%' OR kkundenr = '"& sogKri &"'"
    else
    sogKri = "Søg.."
    kSQLkri = "AND kid <> 0"
    end if
	%>
	<div id="sindhold" style="position:absolute; left:10px; top:0px; visibility:visible;">
	
	<%
	'**********************************************************
	'**************** Orpet / red Faktura Step 1 **************
	'**********************************************************
	 %>
	 
	 <!--
	    <table cellspacing=2 cellpadding=0 border=0>
	    <tr><td>
	    <a href="erp_opr_faktura_kontakter.asp" target="erp2_1" class="rmenu"><u>1) - Vælg kontakt (debitor)</u></a>
	    </td></tr>
	    <tr><td>
	    <a href="#" target="erp2_1" class="erp_gray">2) - Vælg job / aftale og datointerval</a>
	    </td></tr>
	    </table>
	    -->
	 
	  <img src="../ill/blank.gif" height="4" width="1" border="0"/><br />
	 <table cellspacing=0 cellpadding=0 border=0 width="275">
	 <form action="erp_opr_faktura_kontakter.asp" method="POST" target="erp2_0">
	 <tr>
	   <td bgcolor="#ffffff" style="padding:10px 10px 10px 10px; border:1px #8caae6 solid;">
            <input id="FM_sog" name="FM_sog" type="text" value="<%=sogKri %>"  onfocus="rensSog()" style="width:150px; font-size:9px;">&nbsp;(% wildcard)&nbsp;
            <input id="Submit0" type="submit" value="Søg" style="font-size:9px;" />
	
	 </td></tr> </form>
	 </table>
	    <img src="../ill/blank.gif" height="4" width="1" border="0"/><br />
	 <table cellspacing=0 cellpadding=0 border=0 width="275">
	 <form action="erp_opr_faktura_kontojob.asp" target="erp2_1" method="POST">
	  <tr>
	   <td bgcolor="#ffffe1" style="padding:10px 0px 10px 10px; border:1px #8caae6 solid; border-right:0px;">
	
	 <b>Kontakter:</b><br />
	    <select name="FM_kunde" size="1" style="width:225px; font-size:9px;" onchange="nextstep1()">
		<%
		
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' "& kSQLkri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr") %>)</option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
				<option value="0">Ingen</option>
		</select> 
		</td><td bgcolor="#ffffe1" valign=top style="padding:22px 7px 0px 0px; border:1px #8caae6 solid; border-left:0px;">
           <input id="Button1" type="image" src="../ill/pilstorxp.gif" onclick="nextstep1()" />
	    </td></tr></form>
	 </table>
        
	</div>
	
	<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->