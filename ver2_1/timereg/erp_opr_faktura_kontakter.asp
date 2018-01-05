




	
	
	

    
	
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
	 
     <%select case lto
            case "xintranet - local", "bf" 
              %>
              <form action="erp_opr_faktura_fs.asp?formsubmitted=1&visjobogaftaler=1" method="POST">
             <% 
            case else
               %>
            <form action="erp_opr_faktura_fs.asp?formsubmitted=1" method="POST">
           <%
             end select  %>
	
        <table style="width:280px;">
	 
	 <tr>
	   <td bgcolor="#ffffff" style="padding:10px 10px 10px 10px; border:0px #8caae6 solid;">
           <%select case lto
            case "xintranet - local", "bf" 
              %>
             Search Project:
             <% 
            case else
               %>
           Søg på kunde:
           <%
             end select  %>
            
            <br /><input id="FM_sog" name="FM_sog" type="text" value="<%=sogKri %>" style="width:205px; border:2px yellowgreen solid; padding:2px;">&nbsp;<input id="Submit0" type="submit" value=">>" style="font-size:9px;" />
           
           <br /><span style="color:#999999; font-size:9px;">(% wildcard)</span></td></tr> 
	 </table>
         </form>
	    

    <%select case lto
      case "bf", "xintranet - local"
        
       case else %>

        <form action="erp_opr_faktura_fs.asp?formsubmitted=1&visjobogaftaler=1" method="POST">
            <input type="hidden" name="FM_sog" value="<%=sogKri %>" />
	 <table style="width:280px;">
	 
	  <tr>
	   <td bgcolor="#FFFFFF" style="padding:10px 10px 10px 10px; border:0px #8caae6 solid;">
	
	 Kunde(r):<br />
      <select name="FM_kunde" id="FM_kunde" size="1" style="width:215px; font-size:11px;">
		<%
		
		        foundKid = 0
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' AND (useasfak <= 2) "& kSQLkri &" AND kstatus = 1 ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
                foundKid = 1
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr") %>)</option>
				<%
				oRec.movenext
				wend
				oRec.close

                if cint(foundKid) = 0 then
				%>
				<option value="0" SELECTED>Ingen</option>
                <%else %>
                <option value="0">Ingen</option>
                <%end if %>

		</select>&nbsp;<input type="submit" value=">>" style="font-size: 9px;"  />
           <!--<input id="Button1" type="image" src="../ill/pilstorxp.gif" onclick="nextstep1()" />-->

		<br />
		  
		 <input id="FM_jobonoff" name="FM_jobonoff" type="checkbox" value="j" <%=jobonoffCHK %> /> Vis lukkede job og aft. 
		
		</td></tr>
	 </table>
        </form>

    <%end select %>
    <!--
	</div>
	-->