
   
    


    <script type="text/javascript">


    function njob(){
	document.getElementById("kFM_job").value = "0"
	document.getElementById("kFM_aftale").style.backgroundColor = "#ffff99"
	document.getElementById("kFM_job").style.backgroundColor = "#ffffff"
	//document.getElementById("jobelaft2").checked = true;
	}
	
	function naft(){
	//alert("her")
	document.getElementById("kFM_aftale").value = "0"
	document.getElementById("kFM_aftale").style.backgroundColor = "#ffffff"
	document.getElementById("kFM_job").style.backgroundColor = "#ffff99"
	//document.getElementById("jobelaft").checked = true;
	}
	
	
    </script>
    <%
	

    
   
    
   
  
    
    
    jobonoff = vislukkedejob
    
    
    
    
	


	'**********************************************************
	'**************** Orpet / red Faktura Step 2 **************
	'**********************************************************
	 %>
	

	    
	 
	 
	

         <form action="erp_opr_faktura_fs.asp?formsubmitted=2&visjobogaftaler=1&visminihistorik=1" method="POST">
      <input id="FM_kunde" name="FM_kunde" type="hidden" value="<%=kid %>" />
      <input id="FM_usedatokri" name="FM_usedatokri" type="hidden" value="1"/>
	 <table style="width:280px;">
	
	
	 
	
	  <tr>
	      <td colspan=2 bgcolor="#FFFFFF" style="padding:10px 10px 10px 10px; border:0px #8caae6 solid; border-bottom:0px;">
              
            <%select case lto
            case "xintranet - local", "bf" 
             %>
            Choose project:
             <% 

                vaelgJObTxt = "Select.."
	            
	            jobonoffSQLkri = " AND (jobstatus = 1 OR jobstatus = 2) "
	             
                 kSQLkri = replace(kSQLkri, "AND", "")
                 kSQLkri = replace(kSQLkri, "kkundenavn", "jobnavn")
                 kSQLkri = replace(kSQLkri, "kkundenr", "jobnr")

                strSQL = "SELECT id, jobnavn, jobnr, jobstatus FROM job WHERE ("& kSQLkri &") "& jobonoffSQLkri &" ORDER BY jobnavn"

            case else
               %>
            Vælg job eller aftale:
             <%

                  vaelgJObTxt = erp_txt_489

                  if cint(jobonoff) = 1 then
	                jobonoffSQLkri = " AND jobstatus <> 99 "
	                else
	                jobonoffSQLkri = " AND (jobstatus = 1 OR jobstatus = 2) "
	                end if

                    strSQL = "SELECT id, jobnavn, jobnr, jobstatus FROM job WHERE "& kidSQL &" "& jobonoffSQLkri &" ORDER BY jobnavn"

             end select  %>
	
              <br />
	   
	    <%
	   
	    
	    
	  
	    'Response.Write strSQL
	     %>
	     <!--<input id="jobelaft" name="jobelaft" type="radio" value="1" <%=jobelaft1CHK %> onclick="naft()" />-->
	     <select id="kFM_job" name="FM_job" style="width:250px;" onchange="naft()">

              <%select case lto
            case "xintranet - local", "bf" 
             case else%>

	     <option value="0"><%=vaelgJObTxt %></option>
	     <% end select

             antalJob = 0

	        oRec.open strSQL, oConn, 3
            while not oRec.EOF


            if cdbl(jobid) = oRec("id") then
            jidSel = "SELECTED"
            else
            jidSel = ""
            end if %>
            <option value="<%=oRec("id") %>" <%=jidSel %>><%=oRec("jobnavn") %> (<%=oRec("jobnr") %>)
            <%select case oRec("jobstatus")
            case 0 
                jobstatusTxt = " - "& erp_txt_323 &""
            case 1 
                jobstatusTxt = ""
            case 2 
                jobstatusTxt = " - "& erp_txt_490 &""
            case 3 
                jobstatusTxt = " - "& erp_txt_427 &""
            case 4 
                jobstatusTxt = " - "& erp_txt_428 &""
            case else
                jobstatusTxt = " - "& erp_txt_324 &""
            end select %>
                
                <%=jobstatusTxt %>
                
            
            
            </option>
            <%
            antalJob = antalJob + 1 
            oRec.movenext
            wend
            oRec.close%>

              <%select case lto
            case "intranet - local", "bf" 

                  if cint(antalJob) = 0 then
                  %>
                   <option value="0">No match - search again</option>
                  <%end if
             case else%>

	    
	     <% end select
             %>
           </select>
           
          <br /><img src="../ill/blank.gif" height="4" width="1" border="0"/><br />
             
	    <%
	    
             select case lto
            case "xintranet - local", "bf" 

            case else

	                if cint(jobonoff) = 1 then
	                aftonoffSQLkri = " AND status <> 99 "
	                else
	                aftonoffSQLkri = " AND status = 1 "
	                end if
	    
	    
	                strSQL = "SELECT id, navn, aftalenr, status FROM serviceaft WHERE "& aftKidSQL &" " & aftonoffSQLkri & " ORDER BY navn"
	                'Response.Write strSQL
	                 %>
	                 <!--<input id="jobelaft2" name="jobelaft" type="radio" value="2" <%=jobelaft2CHK %> onclick="njob()" />-->
	                 <select id="kFM_aftale" name="FM_aftale" style="width:250px;" onchange="njob()">
	                 <option value="0">Vælg aftale..</option>
	                 <%
	                    oRec.open strSQL, oConn, 3
                        while not oRec.EOF
                         if cint(aftaleid) = oRec("id") then
                        aidSel = "SELECTED"
                        else
                        aidSel = ""
                        end if%>
                        <option value="<%=oRec("id") %>" <%=aidSel %>><%=oRec("navn") %> (<%=oRec("aftalenr") %>)
                
                            <%if oRec("status") <> 1 then %>
                            - <%=erp_txt_429 %>
                            <%end if %>
            
                        </option>
                        <%
                        oRec.movenext
                        wend
                        oRec.close%>
                       </select>

                <%end select %>

	      </td>
	             </tr>
         </table>
	

          <table style="width:280px;">
	    <tr>
        <td bgcolor="#FFFFFF" colspan="2" style="padding:10px 10px 10px 10px;">
	    <b><%=erp_txt_297 %>:</b> (<%=erp_txt_461 %>)<br />
	    <!--#include file="inc/weekselector_s.asp"-->
	  
	 
	    <br /><input name="FM_igdato" id="Checkbox2" type="checkbox" value="1" <%=chkigDato%> /> <%=erp_txt_462 %>
         <br /><span style="float:right; clear:left;"><input id="Button1" type="submit" value=">>" style="font-size: 9px;"/></span>
           <!--<input id="Button1" type="image" src="../ill/pilstorxp.gif" onClick="nextstep2()" />-->
	    </td></tr>
	     
	</table>
	</form>


	