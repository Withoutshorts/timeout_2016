
<!-- frames -->
	<%
	kid = request("FM_kunde")
	
	if len(request("FM_job")) <> 0 then
	jobid = request("FM_job")
	else
	jobid = 0
	end if
	%>
	
	
	<frameset rows="126,*" framespacing="0" border="0" id="erp1">
    
    <frame src="erp_top.asp" name="top" id="top" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
	        
	         <frameset cols="290,*" framespacing="0" id="erp6">
		            
		            
		        <frameset rows="106,157,*" framespacing="0" id="erp2">
    		            
    		            <frame src="erp_opr_faktura_kontakter.asp?FM_kunde=<%=kid%>" name="erp2_0" id="erp2_0" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
	                    <%if jobid <> 0 then %>
	                    <frame src="erp_opr_faktura_kontojob.asp?FM_kunde=<%=kid%>&FM_job=<%=jobid%>" name="erp2_1" id="erp2_1" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
	                    <%else %>
	                    <frame src="erp_opr_faktura_blank.asp" name="erp2_1" id="erp2_1" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
	                    <%end if %>
	                    
	                    <frame src="erp_opr_faktura_blank.asp" name="erp2_2" id="erp2_2" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
	                    <!--<frame src="erp_opr_faktura_dato.asp" name="erp2_3" id="erp2_3" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">-->
    	               
    	                
		        </frameset>
		        
		        
		        <frame src="erp_opr_faktura_blank.asp" name="erp3" id="erp3" frameborder="0" scrolling="auto" marginwidth="0" marginheight="0">
	 
		        
		        
		    </frameset>
    
    
                   
                   
		    
		  
		    
    </frameset>
	
	
	
	






