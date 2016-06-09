



	
	
	
	
    
	
	                    <!-- Materialer --->
	                    
	                   
	                      
	                    <div id=matdiv style="position:absolute; width:700px; visibility:hidden; display:none; border:1px orange solid; top:104px; left:5px; padding:10px 0px 5px 10px; background-color:#ffffff;">
                        
                         <div id=matvaldiv style="position:absolute; left:725px; top:340px; width:210px; border:0px #8cAAe6 solid; padding:0px;">
	                      <table width=100% border=0 cellspacing=0 cellpadding=0 bgcolor="#ffffff">
	                    <tr>
	                        <td colspan=2 style="padding:5px 5px 5px 10px; border:1px #CCCCCC solid;">
	                            <b>Fakturerings valuta</b><br />
	                            <%call selectAllValuta(3,jftp) %>
                    	        
	                        </td>
	                    </tr>
	                    </table>
	                    </div>
                        
                        <h4>Materialer / Udlæg</h4>
                        
                      
                         
	                        <!--
	                        <a href="materialer_indtast.asp?id=<%=jobid %>&fromsdsk=0&aftid=<%=aftid%>&useFM_jobsog=0" class=vemnu target="_blank">Indtast materiale forbrug..</a>
                            -->

                             <%
                nWdt = 220
                nTxt = "Materiale forbrug / Udlæg"
                nLnk = "materialer_indtast.asp?id="&jobid&"&fromsdsk=0&aftid="&aftid&"&useFM_jobsog=0"
                nTgt = "_blank"
                call opretNy_2013(nWdt, nTxt, nLnk, nTgt)
                                 
                                 
                nWdt = 220
                nTxt = "Indtast fra lager"
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
	                            <td colspan=2><b>Intern kode</b><br />
                                    <input id="FM_mat_viskuneks1" name="FM_mat_viskuneks" value="1" type="radio" <%=viskeksCHK1 %> /> Vis kun materialer til videre-fakturering (oprettet som ekstern)
                                    <br /><input id="FM_mat_viskuneks0" name="FM_mat_viskuneks" value="0" type="radio" <%=viskeksCHK0 %>/> Vis alle</td>
	                        </tr>
	                        <tr>
	                            <td><br /><b>Periode</b><br />
                                    <input id="FM_mat_ignper" name="FM_mat_ignper" value="1" type="checkbox" <%=matignperCHK %> /> Ignorer periode (vis alle ikke fakturerede materialer)</td>
                                     <td>
                                   &nbsp;</td>
	                        </tr>
                                 <tr>
                                    <td colspan="2" style="padding-right:20px;">
                                    <input id="FM_mat_hentnye" type="button" value=" Hent nye >> " /></td>
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
                            <b>Fordel materialer på fakturalayout:</b><br /> 
                                <select id="FM_showmatasgrp" name="FM_showmatasgrp">
                                    <option value="0" <%=showmatasgrpSEL0 %>>Materialeforbrug vises som liste i bunden af faturaen</option>
                                <option value="1" <%=showmatasgrpSEL1 %>>Materialer vises som sum fordelt på grupper.</option> <!-- (materialer uden gruppe vises som diverse) -->
                                    <option value="2" <%=showmatasgrpSEL2 %>>Materialer vises fordelt på aktiveter/faser</option>
                                </select><br />&nbsp;
                            
                            <!--<input id="FM_showmatasgrp" name="FM_showmatasgrp" value="1" type="checkbox" <%=showmatasgrpCHK%> /> -->
	                        
	                        <div id="divmatreg">
	                        <!-- henter mat fra jquery --->
	                        </div>
                        	
	
	
	
	
	
	 <div id=matdiv_2 style="position:relative; visibility:hidden; display:none; top:80px; width:700px; left:5px; border:0px #8cAAe6 solid;">
        <table width=700 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('aktdiv')" class=vmenu><< Forrige</a></td><td align=right><a href="#" onclick="showdiv('betdiv')" class=vmenu>Næste >></a></td></tr></table>
    </div>
	
	
    </div>
	 
	
	
	<div id="matsubtotal" style="position:absolute; left:730px; top:184px; width:200px; z-index:2000; border:1px orange solid; background-color:#ffffff; padding:5px;">
    <b>Materialer:</b><br />
	<table cellspacing=5 cellpadding=0 border=0 width=100%><tr>
	<td align=right>Subtotal:
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
	
	
	
	
	
	
	
	