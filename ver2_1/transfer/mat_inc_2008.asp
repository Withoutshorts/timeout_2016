



	
	
	
	
    
	
	                    <!-- Materialer --->
	                    
	                   
	                      
	                    <div id=matdiv style="position:absolute; width:700px; visibility:hidden; display:none; border:1px orange solid; top:104px; left:5px; padding:10px 0px 5px 10px; background-color:#ffffff;">
                        
                         <div id=matvaldiv style="position:absolute; left:725px; top:300px; width:200px; border:0px #8cAAe6 solid; padding:0px;">
	                      <table width=100% border=0 cellspacing=0 cellpadding=0 bgcolor="#ffffe1">
	                    <tr>
	                        <td colspan=2 style="padding:5px 5px 5px 10px; border:1px #CCCCCC solid;">
	                            <b>Fakturerings valuta</b><br />
	                            <%call selectAllValuta(3,jftp) %>
                    	        
	                        </td>
	                    </tr>
	                    </table>
	                    </div>
                        
                        <h4>Materialer / Udlæg</h4>
                        
                      
                       
	                        
	                        <a href="materialer_indtast.asp?id=<%=jobid %>&fromsdsk=0&aftid=<%=aftid%>" class=vemnu target="_blank">Indtast materiale forbrug..</a>
	                        <br />
                            &nbsp;
	                        <%
	                        
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
	                        %>
	                        
	                        <div style="border:1px #8cAAe6 solid; background-color:#Eff3FF; padding:5px; width:400px;">
	                        <table cellpadding=2 cellspacing=1 border=0 width=100%>
	                        <tr>
	                            <td colspan=2><b>Intern kode</b><br />
                                    <input id="FM_mat_viskuneks1" name="FM_mat_viskuneks" value="1" type="radio" <%=viskeksCHK1 %> /> Vis kun materialer til videre-fakturering (oprettet som ekstern)
                                    <br /><input id="FM_mat_viskuneks0" name="FM_mat_viskuneks" value="0" type="radio" <%=viskeksCHK0 %>/> Vis alle</td>
	                        </tr>
	                        <tr>
	                            <td><b>Periode</b><br />
                                    <input id="FM_mat_ignper" name="FM_mat_ignper" value="1" type="checkbox" <%=matignperCHK %> /> Ignorer periode (vis alle ikke fakturerede materialer)</td>
                                     <td>
                                    <input id="FM_mat_hentnye" type="button" value=" Hent nye >> " /></td>
	                        </tr>
	                       
	                        </table>
	                        </div>
	                        <br /><br />
	                        
	                        
	                        <%if showmatasgrp = 1 then
	                        showmatasgrpCHK = "CHECKED"
	                        else
	                        showmatasgrpCHK = ""
	                        end if%>
                            <input id="FM_showmatasgrp" name="FM_showmatasgrp" value="1" type="checkbox" <%=showmatasgrpCHK%> /> Vis sum af udlæg / materialeforbrug fordelt på grupper på faktura layout. (materialer uden gruppe vises som diverse)
	                        
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
	
	
	
	
	
	
	
	