



	
	
	
	
    
	
	                    <!-- Materialer --->
	                    
	                   
	                      
	                    <div id=matdiv style="position:absolute; width:700px; visibility:hidden; display:none; border:1px orange solid; top:104px; left:5px; padding:10px 0px 5px 10px; background-color:#ffffff;">
                        
                         <div id=matvaldiv style="position:absolute; left:725px; top:340px; width:210px; border:0px #8cAAe6 solid; padding:0px;">
	                      <table width=100% border=0 cellspacing=0 cellpadding=0 bgcolor="#ffffff">
	                    <tr>
	                        <td colspan=2 style="padding:5px 5px 5px 10px; border:1px #CCCCCC solid;">
	                            <b><%=erp_txt_396 %></b><br />
	                            <%call selectAllValuta(3,jftp) %>
                    	        
	                        </td>
	                    </tr>
	                    </table>
	                    </div>
                        
                        <h4><%=erp_txt_397 %></h4>
                        
                      
                         
	                        <!--
	                        <a href="materialer_indtast.asp?id=<%=jobid %>&fromsdsk=0&aftid=<%=aftid%>&useFM_jobsog=0" class=vemnu target="_blank">Indtast materiale forbrug..</a>
                            -->

                             <%
                nWdt = 220
                nTxt = erp_txt_398
                nLnk = "materialer_indtast.asp?id="&jobid&"&fromsdsk=0&aftid="&aftid&"&useFM_jobsog=0"
                nTgt = "_blank"
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt)
                                 
                                 
                nWdt = 220
                nTxt = "Indtast fra lager" 'mangler
                nLnk = "materialer_indtast.asp?id="&jobid&"&fromsdsk=0&aftid="&aftid&"&useFM_jobsog=0&vis=lag"
                nTgt = "_blank"
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>

	                        <br />
                            &nbsp;
	                        <%
	                        
                            select case lto 
                                case "nt", "intranet - local"

                                viskeks = 0
                                viskeksCHK0 = "CHECKED"
	                            viskeksCHK1 = ""

                                matignperCHK = "CHECKED"

                                case else



	                                if request.cookies("erp")("matvke") <> "" then
	                                viskeks = request.cookies("erp")("matvke")
	                                else
	                                viskeks = 1
	                                end if
	                        
	                                if viskeks = 1 then
	                                viskeksCHK1 = "CHECKED"
	                                viskeksCHK0 = ""
	                                else
	                                viskeksCHK0 = "CHECKED"
	                                viskeksCHK1 = ""
	                                end if
	                        
	                                if request.cookies("erp")("matignper") <> "" then
	                                matignperCHK = "CHECKED"
	                                else
	                                matignperCHK = ""
	                                end if

                            end select
	                        %>
	                        
	                        
	                        <table cellpadding=2 cellspacing=1 border=0 width=100%>
                           
	                        <tr>
	                            <td colspan=2><b><%=erp_txt_399 %></b><br />
                                    <input id="FM_mat_viskuneks1" name="FM_mat_viskuneks" value="1" type="radio" <%=viskeksCHK1 %> /> <%=erp_txt_400 %>
                                    <br /><input id="FM_mat_viskuneks0" name="FM_mat_viskuneks" value="0" type="radio" <%=viskeksCHK0 %>/> <%=erp_txt_401 %></td>
	                        </tr>
	                        <tr>
	                            <td><br /><b><%=erp_txt_402 %></b><br />
                                    <input id="FM_mat_ignper" name="FM_mat_ignper" value="1" type="checkbox" <%=matignperCHK %> /> <%=erp_txt_403 %></td>
                                     <td>
                                   &nbsp;</td>
	                        </tr>
                                 <tr>
                                    <td colspan="2" style="padding-right:20px;">
                                    <input id="FM_mat_hentnye" type="button" value=" <%=erp_txt_404 %> >> " /></td>
                                </tr>
	                       
	                        </table>
	                       
	                       
	                        
	                        
	                        <%
                                showmatasgrpSEL0 = ""
                                showmatasgrpSEL1 = ""
                                showmatasgrpSEL2 = ""


                             select case showmatasgrp 
                             case 1 
	                         showmatasgrpSEL1 = "SELECTED"
                             case 2
	                         showmatasgrpSEL2 = "SELECTED"
                             case else
	                         showmatasgrpSEL0 = "SELECTED"
	                         end select%>
                            <br />
                            <b><%=erp_txt_405 %>:</b><br /> 
                                <select id="FM_showmatasgrp" name="FM_showmatasgrp">
                                    <option value="0" <%=showmatasgrpSEL0 %>><%=erp_txt_406 %></option>
                                <option value="1" <%=showmatasgrpSEL1 %>><%=erp_txt_407 %></option> <!-- (materialer uden gruppe vises som diverse) -->
                                    <option value="2" <%=showmatasgrpSEL2 %>><%=erp_txt_408 %></option>
                                </select><br />&nbsp;
                            
                            <!--<input id="FM_showmatasgrp" name="FM_showmatasgrp" value="1" type="checkbox" <%=showmatasgrpCHK%> /> -->
	                        
	                        <div id="divmatreg">
	                        <!-- henter mat fra jquery --->
	                        </div>
                        	
	
	
	
	
	
	 <div id=matdiv_2 style="position:relative; visibility:hidden; display:none; top:80px; width:700px; left:5px; border:0px #8cAAe6 solid;">
        <table width=700 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('aktdiv')" class=vmenu><< <%=erp_txt_391 %></a></td><td align=right><a href="#" onclick="showdiv('betdiv')" class=vmenu><%=erp_txt_392 %> >></a></td></tr></table>
    </div>
	
	
    </div>
	 
	
	
	<div id="matsubtotal" style="position:absolute; left:730px; top:184px; width:200px; z-index:2000; border:1px orange solid; background-color:#ffffff; padding:5px;">
    <b><%=erp_txt_411 %>:</b><br />
	<table cellspacing=5 cellpadding=0 border=0 width=100%><tr>
	<td align=right><%=erp_txt_412 %>:
	<%
	'*********************************************************
	'Subtotal
	'*********************************************************
	%>
	
	
		<!-- matBeloeb -->
		<%
		'if len(matSubTotalAll) <> 0 then
		'matSubTotalAll = SQLBlessDot(formatnumber(matSubTotalAll, 2))
		'else
		'matSubTotalAll = formatnumber(0, 2)
		'end if
		'<%=matSubTotalAll
		
		
		%>
		
		
		<input type="hidden" name="FM_materialer_beloeb" id="FM_materialer_beloeb" value="0">
		<div style="position:relative; width:95; height:20px; border-bottom:2px orange dashed; background-color:#ffffff; padding-right:3px;" align="right" id="divmatbelobtot"><b><%= matSubTotalAll &" "& valutakodeSEL%></b></div>
		</td>
		</tr>
	</table>
	</div>
	
	
	
	
	
	
	
	