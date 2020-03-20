




	
	
	

    
	
	<%
	'**********************************************************
	'**************** Orpet / red Faktura Step 1 **************
	'**********************************************************
	 %>
	 
	
    
            <!--<form action="erp_opr_faktura_fs.asp?formsubmitted=1" method="POST">   </form>-->
       
         <form action="erp_opr_faktura_fs.asp?formsubmitted=1&visjobogaftaler=1" method="POST" id="form_kunde">
             <div class="row">
           <div class="col-lg-4">
               <!--<input type="text" id="help" />-->

             <input id="FM_sogkunde" name="FM_sogkunde" type="text" value="<%=sogKri %>" placeholder="Search customer" class="form-control input-small" style="width:300px;">
             <!--<span style="color:#999999; font-size:9px;">(% wildcard)</span>-->
           
               <!--</div>
          <div class="col-lg-2"><br />
               <input id="Submit0" type="submit" value=">>" class="btn btn-sm btn-success" />
             -->
           
           
	
      
	    

   
       
            <!--<input type="hidden" name="FM_sog" value="<%=sogKri %>" />-->
	
    


				<%
                    '**Skriver valgte kunde ***'
                    
                if cdbl(kid) <> 0 then    
                strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE kid = "& kid 
				oRec.open strSQL, oConn, 3
				if not oRec.EOF then
				
                    kundeValgt = oRec("kkundenavn") & " ("& oRec("kkundenr") &")"

                    Response.Write "<br><span style=""color:#999999; font-size:75%;"">Valgt kunde:</span><br><b>"& kundeValgt & "</b><br>"
			
				end if
				oRec.close 
                    
                 end if%>




      <select name="FM_kunde" id="FM_kunde" size="10" class="form-control input-small" style="width:300px; visibility:hidden; display:none;">

          </select>

        <input id="FM_jobonoff" name="FM_jobonoff" type="checkbox" value="j" <%=jobonoffCHK %>/> Vis lukkede job og aft. <br />

		
		</div>
        </div><!-- row -->
        </form>

        
    