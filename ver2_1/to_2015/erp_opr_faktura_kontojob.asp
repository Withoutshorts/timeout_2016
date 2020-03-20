
   
    


    <%
	

    
   
    
   
  
    
    
    jobonoff = vislukkedejob
    
    
    
    
	


	'**********************************************************
	'**************** Orpet / red Faktura Step 2 **************
	'**********************************************************
	 %>
	


   <form id="form_jobaft" action="erp_opr_faktura_fs.asp?formsubmitted=2&visjobogaftaler=1&visminihistorik=1" method="POST">
      <input id="FM_kunde" name="FM_kunde" type="hidden" value="<%=kid %>" />
      <input id="FM_usedatokri" name="FM_usedatokri" type="hidden" value="1"/>
   <div class="row">
	    
	
	 
	

      
	
              
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
            

                  vaelgJObTxt = erp_txt_489

                  if cint(jobonoff) = 1 then
	                jobonoffSQLkri = " AND jobstatus <> 99 "
	                else
	                jobonoffSQLkri = " AND (jobstatus = 1 OR jobstatus = 2) "
	                end if

                    strSQL = "SELECT id, jobnavn, jobnr, jobstatus FROM job WHERE "& kidSQL &" "& jobonoffSQLkri &" ORDER BY jobnavn"

             end select  %>
	
            
	    <%
	   
	    
	    
	  
	    'Response.Write strSQL
	     %>
           
             <div class="col-lg-4">

	     <!--<input id="jobelaft" name="jobelaft" type="radio" value="1" <%=jobelaft1CHK %> onclick="naft()" />-->
	     <select id="kFM_job" name="FM_job" style="width:300px;" onchange="naft()" class="form-control input-small">

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
           
        <br />
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
	                 <select id="kFM_aftale" name="FM_aftale" onchange="njob()" style="width:300px;" class="form-control input-small">
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

	    
	
      

	

         </div>
        
              <!--  <div class="col-lg-2"><br /><br />
               <input id="Button1" type="submit" value=" >> " class="btn btn-success">
                </div>-->
        </div>
</form>
	