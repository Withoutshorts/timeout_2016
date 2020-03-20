


<%


    sub underskrift


    %>
                                <div style="position:relative; padding:40px; width:800px;">

                                <table width="100%" border="0" cellpadding="0" cellspacing="0"><tr><td>
                                <%=joblog2_txt_072 &" & "& joblog2_txt_073 %></td>
                                    <td style="width:200px; border-bottom:1px #000000 dashed;">&nbsp;</td>
                                    <td style="width:200px;">&nbsp;</td>
                                    <td style=""><%=joblog2_txt_072 &" & "& joblog2_txt_075 &" "& joblog2_txt_074 %></td>
                                    <td style="width:200px; border-bottom:1px #000000 dashed;">&nbsp;</td>

                                       </tr></table>

                            </div>
    <%

    end sub


        sub underskriftExcel
        
                    ekspTxt = ekspTxt & "xx99123sy#z" & "xx99123sy#z" & "xx99123sy#z" & "xx99123sy#z" & "xx99123sy#z"  
                         ekspTxt = ekspTxt & ";;Underskrift medarbejder:;;;;;;;;Dato & Underskift (overordnet);;;;;;;;;;;;;;;;;;;;;;;;"
                       
        end sub


        public timerTotprDagGT, timerTotDenneMedarb
        redim timerTotprDagGT(31)
        function medarbMdtot()
        timerTotDenneMedarb = 0

        %>
                                        <tr style="background-color:#FFC0CB;">
                                            <td colspan="2" style="border-top:1px #cccccc solid;"><%=joblog2_txt_076 %>:</td>
                                                <%
                                               for d = 1 to vlgtMdLastDay 
                                    

                                                    if timerTotprDag(d) <> 0 then
                                                    timerTotprDagTxt = formatnumber(timerTotprDag(d), 2) 
                                                    else
                                                    timerTotprDagTxt = ""
                                                    end if
                                                   %>
                              
                                                    <td style="width:30px; border-left:1px #999999 solid; border-top:1px #cccccc solid;" align="right"><%=timerTotprDagTxt %></td>

                                                <%
                                                timerTotprDagGT(d) = timerTotprDagGT(d) + timerTotprDag(d)
                                                timerTotDenneMedarb = timerTotDenneMedarb + timerTotprDag(d)
                                                 timerTotprDag(d) = 0   
                                                 next 

                                                if cint(mthrap_grpbyakt) = 1 then
                                                       %>
                                                        <td style="width:30px; border-left:1px #999999 solid; border-top:1px #cccccc solid;"" align="right"><b><%=formatnumber(timerTotDenneMedarb, 2) %></b></td>
                                                        <%

                                               vlgtMdLastDaySpan = vlgtMdLastDay + 1
                                               else
                                               vlgtMdLastDaySpan = vlgtMdLastDay 
                                               end if
                                                    %>

                                            
                                        </tr>


                                        <%if cint(mthrap_grpbyakt) <> 1 then %>
                                        <tr>
                                            <td colspan="2" style="border-top:1px #cccccc solid;"><b><%=joblog2_txt_077 %>:</b><br />&nbsp;</td>
                                            <td colspan="<%=vlgtMdLastDaySpan %>" align="right" style="border-top:1px #cccccc solid;"><b><%=formatnumber(timerTotDenneMedarb, 2) %></b><br />&nbsp;</td>
                                               
                                        </tr>


                                        <%end if

    
        end function



        public timerTotprDagJob
        redim timerTotprDagJob(31)
        function jobMdtot()
        timerTotDetteJob = 0

        %>
                                     
                                                <%
                                               for d = 1 to vlgtMdLastDay 
                                    
                                                 timerTotDetteJob = timerTotDetteJob + timerTotprDagJob(d)
                                                 timerTotprDagJob(d) = 0   
                                                 next 
                                                    
                                               
                                                    
                                               if cint(mthrap_grpbyakt) = 1 then
                                               vlgtMdLastDaySpan = vlgtMdLastDay + 1
                                               else
                                               vlgtMdLastDaySpan = vlgtMdLastDay 
                                               end if
                                               %>



                                            
                                      
                                        <tr style="background-color:#F4F4F4;">
                                            <td colspan="2" style="border-top:1px #cccccc solid;"><%=joblog2_txt_078 %>:</td>
                                            <td colspan="<%=vlgtMdLastDaySpan %>" style="border-top:1px #cccccc solid;" align="right"><%=formatnumber(timerTotDetteJob, 2) %></td>
                                               
                                        </tr>


                                        <%

    
        end function




         function medarbMdtotGT()
        timerTotGT = 0

        %>
                                        <tr style="background-color:#FFFFFF;">
                                            <td colspan="2"><b><%=joblog2_txt_079 %>:</b></td>
                                                <%
                                               for d = 1 to vlgtMdLastDay 
                                    
                                                   %>
                              
                                                    <td style="width:30px; border-left:1px #999999 solid;" align="right"><%=formatnumber(timerTotprDagGT(d), 2) %></td>

                                                <%
                                                timerTotGT = timerTotGT + timerTotprDagGT(d)
                                                timerTotprDagGT(d) = 0
                                                 
                                                 next 
                                                    
                                               if cint(mthrap_grpbyakt) = 1 then
                                               vlgtMdLastDaySpan = vlgtMdLastDay + 1
                                               else
                                               vlgtMdLastDaySpan = vlgtMdLastDay 
                                               end if%>


                                        </tr>
                                         <tr style="background-color:#cccccc;">
                                            <td colspan="2"><b><%=joblog2_txt_080 %>:</b><br />&nbsp;</td>
                                            <td colspan="<%=vlgtMdLastDaySpan %>" align="right"><%=formatnumber(timerTotGT, 2) %><br />&nbsp;</td>
                                               
                                        </tr>

                                        <%


        end function






public lastmedarbnavn, medarbtimer, medarbEnheder, medarbFeriePlan
	public medarbFrokot, medarbAfspadOp, medarbKm
	
	function fordelpamedarbejdere()
	
	
	                
					
					
					
	%>
					
					
					<tr>
						 <td align=right bgcolor="#FFFFFF" colspan="9">&nbsp;</td>
	                     <td bgcolor="#FFFFFF" colspan="7" style="padding:5px 10px 5px 0px; border-top:1px #999999 solid;">
						<%=joblog2_txt_081 %><br />
						<b><%=lastmedarbnavn%>:</b>&nbsp;
						<%
	
	mtotArr = split(replace(akttype_sel, "#", ""), ",")
	for j = 0 to ubound(mtotArr)
	
	call akttyper(mtotArr(j), 1)
	
	if len(trim(vlgt_typerTotalerMed(mtotArr(j)))) <> 0 then
	antal = vlgt_typerTotalerMed(mtotArr(j))
	enh = vlgt_typerTotalerMedEnh(mtotArr(j))
	else
	antal = 0
	enh = 0
	end if
	
	if antal <> 0 then
	Response.write "<b>"& formatnumber(antal, 2)& "</b>" 'akttypenavn & 
	
	    if cint(hideenheder) = 0 then
	    Response.Write " ~ ("& formatnumber(enh, 2) & ")"
	    end if
	 
	 
	 Response.Write "<br>"   
	
	end if
	
	vlgt_typerTotalerMed(mtotArr(j)) = 0
	vlgt_typerTotalerMedEnh(mtotArr(j)) = 0
	
	next %>
				
						
						
						<br />&nbsp;</td>
					</tr>
	
	<%end function
	
	
	
	public lastaktnavn, akttimer, aktEnheder, lastakttype, faseCounta, faseCountenh
	
	function fordelpaakt(sft)
	
	
	
	
	               
	%>
					
					<tr>
						 <td align=right bgcolor="#FFFFFF" colspan="9">&nbsp;</td>
	                     <td bgcolor="#FFFFFF" colspan="7" style="padding:5px 10px 5px 0px; border-top:1px #999999 solid;">
					
                        <%=joblog2_txt_086 %><br /> <b><%=lastaktnavn%>:</b>
						<%
	
	atotArr = split(replace(akttype_sel, "#", ""), ",")
	for j = 0 to ubound(atotArr)
	
	call akttyper(atotArr(j), 1)
	
	if len(trim(vlgt_typerTotalerAkt(atotArr(j)))) <> 0 then
	antal = vlgt_typerTotalerAkt(atotArr(j))
	enh = vlgt_typerTotalerAktEnh(atotArr(j))
	else
	antal = 0
	enh = 0
	end if
	
	if antal <> 0 then
	Response.write formatnumber(antal, 2)
	    
	    if cint(hideenheder) = 0 then
	    Response.Write " ~ ("& formatnumber(enh, 2) & ")"
	    end if
	 
	 response.write "<br><span style=""color:#999999; font-size:9px;"">("& akttypenavn &")</span>"

	 'Response.Write "<br>"
	 
	 faseCounta = faseCounta + antal
	faseCountenh = faseCountenh + enh 
	
	end if
	
	vlgt_typerTotalerAkt(atotArr(j)) = 0
	vlgt_typerTotalerAktEnh(atotArr(j)) = 0
	
	next %>

           <br />&nbsp;</td>                 
					</tr>
	
	
	<%if (lcase(trim(lastFase)) <> lcase(trim(thisFase)) AND len(trim(lastFase)) <> 0 AND isNull(lastFase) <> true) OR (sft = 1 AND isNull(lastFase) <> true) AND hidefase <> 1 then  %>
	<%'if t <> 100 then %>
	<tr>
        
	     <td align=right bgcolor="#FFFFFF" colspan="9">&nbsp;</td>
	     <td bgcolor="#FFFFFF" colspan="7" style="padding:5px 10px 5px 0px; border-top:1px #999999 dashed;"><%=joblog2_txt_082 %><br /> <b><%=replace(lastFase, "_", " ") %>:</b>&nbsp; <b><%=formatnumber(faseCounta, 2) %></b> ~ (<%=formatnumber(faseCountenh, 2) %>)<br /><br />&nbsp; </td>
	</tr>
	<%
    faseCounta = 0
	faseCountenh = 0

    'else
    
        'if (lcase(trim(lastFase)) <> lcase(trim(thisFase)) AND len(trim(lastFase)) <> 0) then
	    'faseCounta = 0
	    'faseCountenh = 0
	    'end if

	end if 
    
    
        if (len(trim(lastFase)) = 0) OR isNull(lastFase) = true then
	    faseCounta = 0
	    faseCountenh = 0
	    end if
    
    %>
	
	
	<%end function
	
	
	
	public gtmTimer, gtmEnh, gtmFplan, gtmKm, gtmFro, gtmAfsp 
	function jobtotaler()
	
	   
	%>
	
	<tr>
    <td align=right bgcolor="#FFFFFF" colspan="9">&nbsp;</td>
	<td bgcolor="#FFFFFF" colspan="7" style="padding:5px 10px 5px 0px; border-top:1px #999999 solid;">
	
	<%if cint(joblog_uge) = 2 then%>
	<b>Uge <%=lastWeekNum %></b>&nbsp;
	<%else %>
	Job<br /> <b><%=lastjobnavn%></b>&nbsp;
	 <%end if

     
        
	
	jtotArr = split(replace(akttype_sel, "#", ""), ",")
	for j = 0 to ubound(jtotArr)
	
	call akttyper(jtotArr(j), 1)

         'if jtotArr(j) = 1 then
         'akttypenavn = "Fakturerbar"
         'end if

         'if jtotArr(j) = 2 then
         'akttypenavn = "Ikke fakturerbar"
         'end if

         'if jtotArr(j) > 2 then
         'akttypenavn = "Anden type"
         'end if
	
	if len(trim(vlgt_typerTotaler(jtotArr(j)))) <> 0 then
	antal = vlgt_typerTotaler(jtotArr(j))
	enh = vlgt_typerTotalerEnh(jtotArr(j))
	else
	antal = 0
	enh = 0
	end if
	
	if antal <> 0 then
        
        'if cint(joblog_uge) = 2 then
	    Response.write "<br>"& akttypenavn & ": "& formatnumber(antal, 2) 
        'else
        'Response.write ""& akttypenavn &"  <b>"& formatnumber(antal, 2)& "</b>"
        'end if
	        
	    if cint(hideenheder) = 0 then
	    Response.Write " ~ ("& formatnumber(enh, 2) & ")"
	    end if
	 
	 
	 'Response.Write "<br>"
	        
	end if
	
	vlgt_typerTotaler(jtotArr(j)) = 0
	vlgt_typerTotalerEnh(jtotArr(j)) = 0
	
	next %>
	
	
	 <br />&nbsp;
	 </td>
	</tr>
    </table>
    
    
    <%if cint(vismat) <> 1 then %>
    </div>
    <!-- slut table DIV -->	
    <br /><br /><br /><br />
	<%end if %>
	
	
	<%
	end function
	
	
	public kmTotdag,timerTotdag,enhederTotdag,afspadTotdag,frokostTotdag,fplanTotdag
	function dagstotaler(visning)
	
					
					
	if visning <> 1 AND visning <> 0 then				
	%>
				
					<tr>
					    <td bgcolor="#ffffff" align=right colspan="9">
						&nbsp;</td>
						<td bgcolor="#ffffff" colspan="7" style="padding:5px 10px 5px 0px; border-top:1px #999999 solid;">
						
						<b><%=left(weekdayname(weekday(lastdate )), 3) %>. <%=day(lastdate ) &" "&left(monthname(month(lastdate )), 3) &". "& right(year(lastdate ), 2)%></b><br />
						<%=lastmedarbnavn%>:&nbsp; 
						<%
	
	mtotArr = split(replace(akttype_sel, "#", ""), ",")
	for j = 0 to ubound(mtotArr)
	
	call akttyper(mtotArr(j), 1)
	
	if len(trim(vlgt_typerTotalerMed(mtotArr(j)))) <> 0 then
	antal = vlgt_typerTotalerMed(mtotArr(j))
	enh = vlgt_typerTotalerMedEnh(mtotArr(j))
	else
	antal = 0
	enh = 0
	end if
	
	if antal <> 0 then
	Response.write "<b>"& formatnumber(antal, 2)& "</b>" 'akttypenavn & ":
	    
	    if cint(hideenheder) = 0 then
	    Response.Write " ~ ("& formatnumber(enh, 2) & ")"
	    end if
	 
	 
	 Response.Write "<br>"
	
	end if
	
	vlgt_typerTotalerMed(mtotArr(j)) = 0
	vlgt_typerTotalerMedEnh(mtotArr(j)) = 0
	
	next 
	
	end if
	%>
						
						
						
					<br />&nbsp;	</td>
					</tr>
					
	<%if visning <> 3 then
	
	    call thisWeekNo53_fn(oRec("tdato")) 
        thisWeekNo53_tdato = thisWeekNo53

        call thisWeekNo53_fn(lastdate) 
        thisWeekNo53_lastdate = thisWeekNo53
	
	    if (cint(thisWeekNo53_tdato) = cint(thisWeekNo53_lastdate)) AND lastmedarb = oRec("Tmnr") then %>				
	    <tr>
						    <td bgcolor="#Eff3ff" colspan="16" style="height:20px; padding:5px 5px 5px 5px;">
						    <!--<b><%=left(weekdayname(weekday(oRec("tdato"))), 3) %>. <%=day(oRec("tdato")) &" "&left(monthname(month(oRec("tdato"))), 3) &". "& right(oRec("tdato"), 2)%></b>-->
                                <%=formatdatetime(oRec("tdato"), 2) %>
						    </td>
					    </tr>
        
        
        <%
        end if
    
    end if
    
    
    if visning = 0 then %>				
	<tr>
						<td bgcolor="#Eff3ff" colspan="16" style="height:20px; padding:5px 5px 5px 5px;">
						<!--<b><%=left(weekdayname(weekday(oRec("tdato"))), 3) %>. <%=day(oRec("tdato")) &" "&left(monthname(month(oRec("tdato"))), 3) &". "& right(oRec("tdato"), 2)%></b>-->
                         <%=formatdatetime(oRec("tdato"), 2) %>
						</td>
					</tr>
    
    
    <%
    end if
    
	end function 
	
	
	function tableheader()
	%>
	
	<tr height="20" bgcolor="#8CAAe6">
				       <td style="width:0px;">
                           &nbsp;</td>
				    <td valign=bottom class=alt><b><%=joblog2_txt_006 %></b></td>
					<td valign=bottom style="padding-left:10px;" class='alt'><b><%=joblog2_txt_072 %></b></td>
					<%if komprimer = "1" then %>
					<td valign=bottom class='alt'><b><%=joblog2_txt_083 %></b><br /><%=joblog2_txt_084 %><br /><%=joblog2_txt_085 %></td>

                     <%if hidefase <> 1 then %>
					<td valign=bottom class='alt'><b><%=joblog2_txt_082 %></b></td>
                    <%else %>
                    <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
                    <%end if %>

					<td valign=bottom class='alt'><b><%=joblog2_txt_086 %></b> (<%=joblog2_txt_087 %>)<br />
					<%=joblog2_txt_088 %></td>
					<%else %>
					<td valign=bottom class='alt'>
                        &nbsp;</td>
                        <%if hidefase <> 1 then %>
                        <td valign=bottom class='alt'><b><%=joblog2_txt_082 %></b></td>
                        <%else %>
                        <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
                        <%end if %>
                    <td valign=bottom class='alt'><b><%=joblog2_txt_086 %></b> (<%=joblog2_txt_087 %>)<br />
					<%=joblog2_txt_088 %></td>
					<%end if %>

                    <%if cint(showfor) = 1 then%>
                    <td valign=bottom class='alt' style="padding-right:5px;"><b><%=joblog2_txt_089 %></b></td>
					<%else %>
					<!--<td>&nbsp;</td>-->
					<%end if %>
					
					<td valign=bottom class='alt' style="padding-left:5px;"><b><%=joblog2_txt_115 %></b></td>
					
					
					<td valign=bottom class='alt' align=right style="padding-right:5px;">
					<%if print = "j" then
					%>
					<b><%=joblog2_txt_090 %> <br /> <%=joblog2_txt_091 %>  <br /> <%=joblog2_txt_092 %></b>
					<%
					else %>
					<b><%=joblog2_txt_090 %> / <%=joblog2_txt_091 %> / <%=joblog2_txt_092 %></b>
					<%end if %>
					<br><%=joblog2_txt_116 %></td>
					
					<%if cint(hideenheder) = 0 then %>
					<td valign=bottom class='alt' align=right style="padding-right:5px;">~ <%=joblog2_txt_093 %></td>
					<%else %>
					<td>&nbsp;</td>
					<%end if %>
					
					
					<%if (level <= 2 OR level = 6) AND cint(hidetimepriser) = 0 then  %>
					<td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b><%=joblog2_txt_094 %>*</b></td>
					<td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b><%=joblog2_txt_095 %></b></td>
					<%else %>
					<td>&nbsp;</td>
				    <td>&nbsp;</td>
					<%end if %>
					
					<%if level = 1 AND visKost = 1 then  %>
					<td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b><%=joblog2_txt_096 %></b></td>
					<td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b><%=joblog2_txt_097 %></b></td>
					<%else %>
					<td>&nbsp;</td>
				    <td>&nbsp;</td>
					<%end if %>
					
					
					
					<%if cint(hidegkfakstat) <> 1 then  %>
                    <td valign=bottom class='alt'>
					<%=joblog2_txt_098 %><br /><%=joblog2_txt_099 %></td>
					<td valign=bottom align="center" class='alt'><b><%=joblog2_txt_100 %></b><br /><%=joblog2_txt_053 %><br /><%=joblog2_txt_050 %></td>
					<td valign=bottom class='alt' align="center" style="padding-right:1px;"><b><%=joblog2_txt_101 %></b></td>
				    <%else %>
                    <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				    <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
                    <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				    <%end if %>   
				 </tr>
	
	<%
	end function
	
	

    function materialeforbrug(jobnr, medid)

        thisJobid = 0
        thisJobRisiko = 0
        thisKid = 0
        thisJobnr = 0
        strSQljob = "SELECT id, risiko, jobknr, jobnr FROM job WHERE jobnr = '"& jobnr &"'"
        oRec9.open strSQljob, oConn, 3
        if not oREc9.EOF then

        thisJobid = oRec9("id")
        thisJobRisiko = oRec9("risiko")
        thisKid = oRec9("jobknr")
        thisJobnr = oRec9("jobnr")

        end if
        oRec9.close

        JobRisiko = thisJobRisiko


        'thisJObid = 77

        %>
        <h3>Materialeforbrug</h3>
        <table cellpadding="2" cellspacing="0" style="background-color:#ffffff; width:100%;"><%


            call tablematheader()

    mf = 0
    strSQLmat = "SELECT id, matnavn, matantal, forbrugsdato, godkendt, matkobspris, (matsalgspris*kurs/100) AS matsalgspris, valuta, godkendt, mf.dato, mf.editor, matvarenr, usrid, mnavn, init, matantal*(matsalgspris*kurs/100) AS salgsbeloeb, matantal*(matkobspris*kurs/100) AS kostbeloeb, (matkobspris*kurs/100) AS matkostpris FROM materiale_forbrug mf "_
    &" LEFT JOIN medarbejdere m ON (m.mid = usrid) "_
    &" WHERE jobid = "& thisJobid &" AND ("& replace(medarbSQlKri, "t.tmnr", "usrid") &") AND forbrugsdato BETWEEN '"& startDatoKriSQL &"' AND '"& slutDatoKriSQL &"' ORDER BY forbrugsdato "


    'response.write strSQLmat
    'response.flush

    oRec9.open strSQLmat, oConn, 3
    while not oRec9.EOF

            select case right(mf,1)
            case 0,2,4,6,8
            tdbgM = "#FFFFFF"
            case else
            tdbgM = "#C4C4C4"
            end select

            matsalgsprisbeloeb = oRec9("salgsbeloeb")
            v = v + 1

            call thisWeekNo53_fn(oRec("forbrugsdato")) 
            
                            %>
                            <tr style="background-color:<%=tdbgM%>;">
                                <td></td>
                                <td><%=thisWeekNo53%></td>
                                <td><%=oRec9("forbrugsdato") %></td>
                                <td><%=oRec9("matnavn") %></td>
                                <td><%=oRec9("matvarenr") %></td>
                                <td><%=oRec9("mnavn") & " ["& oRec9("init") &"]"%></td>

                                <% 
                                    tjkDag = oRec9("forbrugsdato")
                                    erugeafsluttet = instr(afslUgerMedab(oRec9("usrid")), "#"&thisWeekNo53&"_"& datepart("yyyy", oRec9("forbrugsdato")) &"#")
                                    strMrd_sm = datepart("m", oRec9("forbrugsdato"), 2, 2)
                                    strAar_sm = datepart("yyyy", oRec9("forbrugsdato"), 2, 2)
                                    strWeek = thisWeekNo53 'datepart("ww", oRec9("forbrugsdato"), 2, 2)
                                    strAar = datepart("yyyy", oRec9("forbrugsdato"), 2, 2)

                                    if cint(SmiWeekOrMonth) = 0 then
                                    usePeriod = strWeek
                                    useYear = strAar
                                    else
                                    usePeriod = strMrd_sm
                                    useYear = strAar_sm
                                    end if


                                     call erugeAfslutte(useYear, usePeriod, oRec9("usrid"), SmiWeekOrMonth, 0, oRec9("forbrugsdato"))
		        
		                            'Response.Write "smilaktiv: "& smilaktiv & "<br>"
		                            'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
		                            'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
		                            'Response.Write "tjkDag: "& tjkDag & "<br>"
		                            'Response.Write "autolukvdato: "& autolukvdato & "<br>"
		                            'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"
		        
		                            call lonKorsel_lukketPer(oRec9("forbrugsdato"), jobRisiko, oRec9("usrid"))
		         
                          
				
                                 '*** tjekker om uge er afsluttet / lukket / lønkørsel
                                call tjkClosedPeriodCriteria(oRec9("forbrugsdato"), ugeNrAfsluttet, usePeriod, SmiWeekOrMonth, splithr, smilaktiv, autogk, autolukvdato, lonKorsel_lukketIO, ugegodkendt)


                                if ((oRec9("godkendt") <> 1 AND request("print") <> "j" AND ugeerAfsl_og_autogk_smil = 0) _
				                OR (oRec9("godkendt") <> 1 AND request("print") <> "j" AND ugeerAfsl_og_autogk_smil = 1 AND level = 1)) _
				                AND (cdate(lastfakdato) < cdate(oRec9("forbrugsdato"))) then  
                                    
                                'if oRec("varenr") = "0" then	
                                'matregid=<%=oRec9("id") &func=red
                                    
                                 %>

                                

                                <td align="right"><a href="materialer_indtast.asp?sogliste=<%=thisJobnr %>&vasallemed=1" target="_blank">
                                    
                                <%=oRec9("matantal") %></a></td>

                                <%else %>

                                <td align="right"><%=oRec9("matantal") %></td>
                                <%end if %>

                                <%if (level <= 2 OR level = 6) AND cint(hidetimepriser) = 0 then  %>
                                <td align="right"><%=formatnumber(oRec9("matsalgspris"), 2) %></td>
                                <td align="right"><b><%=formatnumber(oRec9("salgsbeloeb"), 2) & " " & basisValISO %></b></td>
                                <%else %>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <%end if %>

                                <%if level = 1 AND visKost = 1 then  %>
					            <td align="right"><%=formatnumber(oRec9("matkostpris"), 2)%></td>
                                <td align="right"><%=formatnumber(oRec9("kostbeloeb"), 2) & " " & basisValISO%></td>
					            <%else %>
					            <td>&nbsp;</td>
				                <td>&nbsp;</td>
					            <%end if %>



                                <%if cint(hidegkfakstat) <> 1 then  %>
                                <td><span style="color:#999999;"><%=oRec9("dato") %><br />
                                    <%=left(oRec9("editor"), 15) %></span>
                                </td>
                                <td class=lille>
                                     <%
                                         
                                     
                                 erGk = 0
                                 gkCHK0 = ""
                                 gkCHK1 = ""
                                 gkCHK2 = ""
                                 'gkCHK3 = ""
                                 'gk3bgcol = "#999999"
                                 gk2bgcol = "#999999"
                                 gk1bgcol = "#999999"
                                 gk0bgcol = "#999999"

                                     select case cint(oRec9("godkendt"))
                                     case 2
                                     gkCHK2 = "CHECKED"
                                     erGk = 2
                                     gk2bgcol = "red"
                                     case 1
                                     gkCHK1 = "CHECKED"
                                     erGk = 1
                                     gk1bgcol = "green"
                                     'case 3
                                     'gkCHK3 = "CHECKED"
                                     'erGk = 3
                                     'gk3bgcol = "orange"
                                     case else
                                     gkCHK0 = "CHECKED"
                                     erGk = 0
                                     gk0bgcol = "#000000"
                                     end select
                 
                             erGkaf = ""
                                         

                                         
                                     '*** Godkendelse ***'
					                 if print <> "j" then%>
    					
					               <input name="matids" id="matids" value="<%=oRec9("id")%>" type="hidden" />
					               <input type="radio" name="FM_godkendt_<%=oRec9("id")%>" id="FM_godkendt1_<%=v%>" class="FM_godkendt_1" value="1" <%=gkCHK1 %>><span style="color:<%=gk1bgcol%>;"><%=joblog2_txt_048 %></span><br />
                                   <input type="radio" name="FM_godkendt_<%=oRec9("id")%>" id="FM_godkendt2_<%=v%>" class="FM_godkendt_2" value="2" <%=gkCHK2 %>><span style="color:<%=gk2bgcol%>;"><%=joblog2_txt_050 %></span><br />
                                   <input type="radio" name="FM_godkendt_<%=oRec9("id")%>" id="FM_godkendt0_<%=v%>" class="FM_godkendt_0" value="0" <%=gkCHK0 %>><span style="color:<%=gk0bgcol%>;"><%=joblog2_txt_051 %></span>

                                    
					                <%
					                else
    					            %>
					                &nbsp;
    					            <%
					                end if
                                    %>
                                    
                                    </td>
                                <td style="text-align:right; padding-right:20px;">

                                    <%
                                        erFak = 0
                                        faknr = 0
                                    strSQLermatfak = "SELECT id, matfakid, faknr FROM fak_mat_spec LEFT JOIN fakturaer f ON (f.fid = matfakid) WHERE matfrb_id = "& oRec9("id") & " AND shadowcopy = 0"
                                    oRec8.open strSQLermatfak, oConn, 3
                                    if not oRec8.EOF then

                                        erFak = 1
                                        faknr = oRec8("faknr")

                                    end if
                                    oRec8.close
                                    
                                        
                                       if cint(erFak) = 1 then%>
                                       <a href="erp_fakhist.asp?FM_kunde=<%=thisKid%>&FM_job=<%=thisJobId %>" target="_blank">Yes</a>  <!--(<%=faknr %>)-->
                                       <%end if%>


                                </td>
                                <%else %>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <%end if %>

                            </tr>
                            <%

        mf = mf + 1

    oRec9.movenext
	wend
	oRec9.Close


    %></table>


  </div>
    <!-- slut table DIV -->	
    <br /><br /><br /><br />
   <%


    end function



    function tablematheader()
	%>
	
	<tr height="20" bgcolor="#8CAAe6">
				       <td style="width:0px;">
                           &nbsp;</td>
				    <td valign=bottom class=alt><b><%=joblog2_txt_006 %></b></td>
					<td valign=bottom class='alt'><b><%=joblog2_txt_072 %></b></td>
				
      
					<td valign=bottom class='alt'><b>Name / desc.</b></td>
                    <td valign=bottom class='alt'>Mat. No.</td>
		

                    <%if cint(showfor) = 1 then%>
                    <td valign=bottom class='alt' style="padding-right:5px;"><b><%=joblog2_txt_089 %></b></td>
					<%else %>
					<!--<td>&nbsp;</td>-->
					<%end if %>
					
					<td valign=bottom class='alt' style="padding-left:5px;"><b><%=joblog2_txt_115 %></b></td>
					<td valign=bottom class='alt' align=right style="padding-right:5px;"><b><%=joblog2_txt_092 %></b></td>
					
					
					
					
					<%if (level <= 2 OR level = 6) AND cint(hidetimepriser) = 0 then  %>
					<td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b><%=joblog2_txt_094 %>*</b></td>
					<td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b><%=joblog2_txt_095 %></b></td>
					<%else %>
					<td>&nbsp;</td>
				    <td>&nbsp;</td>
					<%end if %>
					
					<%if level = 1 AND visKost = 1 then  %>
					<td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b><%=joblog2_txt_096 %></b></td>
                    <td valign=bottom align="right" class='alt' style="padding-right:5px;">
					<b><%=joblog2_txt_097 %></b></td>
					<%else %>
					<td>&nbsp;</td>
				    <td>&nbsp;</td>
					<%end if %>
					
					
					
					<%if cint(hidegkfakstat) <> 1 then  %>
                    <td valign=bottom  class='alt'>
					<%=joblog2_txt_098 %><br /><%=joblog2_txt_099 %></td>
					<td valign=bottom align="center" class='alt'><b><%=joblog2_txt_100 %></b><br /><%=joblog2_txt_053 %><br /><%=joblog2_txt_050 %></td>
					<td valign=bottom class='alt' align="center" style="padding-right:1px;"><b><%=joblog2_txt_101 %></b></td>
				    <%else %>
                    <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				    <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
                    <td><img src="ill/blank.gif" width="1" height="1" border="0" /></td>
				    <%end if %>   
				 </tr>
	
	<%
	end function




	public grgrTotal
	sub grandTotal
	
	
	if cint(visopdaterknap) = 1 And print <> "j" AND cint(hidegkfakstat) <> 1 then%>
	
	<br />
	<table border="0" width=<%=globalWdt %> cellpadding="0" cellspacing="0">
    <tr>
     <td align=right style="padding:5px 20px 1px 1px;">
        <b><%=joblog2_txt_102 %>:</b><br /><input name="gkall" id="chkalle1" type="radio" onclick="checkAllGK(1)" /><%=" "& joblog2_txt_103 %><br />
        <input name="gkall" id="chkalle2" type="radio" onclick="checkAllGK(2)" /><%=" "& joblog2_txt_104 %>  <br /> 
        <input name="gkall" id="chkalle3" type="radio" onclick="checkAllGK(3)" /><%=" "& joblog2_txt_105 &" " %><%=joblog2_txt_052 %>.<br />
        <input name="gkall" id="chkalle0" type="radio" onclick="checkAllGK(0)" /><%=" "& joblog2_txt_105 %> = "" (<%=joblog2_txt_117 %>)
       <br />
       <input id="Submit1" type="submit" value="<%=joblog2_txt_107 %>!" /></td>
   
	</tr>
        <input id="antal_v" name="antal_v" value="<%=v %>" type="hidden" />
	</form>
	</table>
	<%end if %>
	<br />
	<br /><h4><%=joblog2_txt_108 %></h4>  
    
    <%
                	
                tTop = 0
                tLeft = 0
                tWdth = globalWdt


                call tableDiv(tTop,tLeft,tWdth)



                %>          
    <table border="0" width=100% cellpadding="0" cellspacing="0" bgcolor="#ffffff">
    <tr>
    <td align=right colspan="16" valign="top" height=20 bgcolor="#ffffFF" style="padding:10px 10px 10px 10px;">
    <b><%=joblog2_txt_109 %>:<br /><br /></b>
    <b><%=joblog2_txt_110 %></b> ~ (<%=joblog2_txt_111 %>)<br />
	<% 
	
	
	
	
	gtotArr = split(replace(akttype_sel, "#", ""), ",")
	for j = 0 to ubound(gtotArr)
	
	call akttyper(gtotArr(j), 1)
	
	if len(trim(vlgt_typerTotalerGrand(gtotArr(j)))) <> 0 then
	antal = vlgt_typerTotalerGrand(gtotArr(j))
	enh = vlgt_typerTotalerGrandEnh(gtotArr(j))
	else
	antal = 0
	enh = 0
	end if
	
	if antal <> 0 then
	Response.write  akttypenavn &": " & formatnumber(antal, 2) 'akttypenavn & ":
	    
	    if cint(hideenheder) = 0 then
	    Response.Write " ~ ("& formatnumber(enh, 2) & ")"
	    end if
	 
	 
	 Response.Write "<br>"   
	 end if
	
	vlgt_typerTotalerGrand(gtotArr(j)) = 0
	vlgt_typerTotalerGrandEnh(gtotArr(j)) = 0
	
	grgrTotal = grgrTotal + antal 
	
	next %>
	
	<br /><br />
	<%=joblog2_txt_079 %>: <b><%=formatnumber(grgrTotal, 2) %> </b><br />
	============================
	</td>
	</tr>
	</table>

    </div>
    <!-- slut table DIV -->	
	
	<%
	end sub
	
	
	
	
	public lastfakdato, jobans1, jobans1txt, jobstatus, usejoborakt_tp
	public jobstartdato, jobslutdato, jobForkalkulerettimer, rekvnr
	public jobans2txt, fasttimepris, fastpris, jobRisiko
	', fid, faknr
	
	
	function jobansoglastfak()
	                    
	                    '*** Jobansvarlige ***'
						jobans1 = 0
						fasttimepris = 0
						fastpris = 0
						usejoborakt_tp = 0
						strSQL2 = "SELECT mnavn, mnr, mid, job.id AS jid, jobans1, "_
						&" jobstatus, jobstartdato, jobslutdato, budgettimer, ikkebudgettimer, "_
						&" rekvnr, fastpris, jobtpris, usejoborakt_tp, job.risiko "_
						&" FROM job "_
						&" LEFT JOIN medarbejdere ON (mid = jobans1) "_
						&" WHERE job.jobnr = '"& oRec("Tjobnr") & "'" 
						
						'Response.Write strSQL2
						'Response.flush
						
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						
						jobid = oRec2("jid") ''* Bruges til at finde sidste faktura
						jobans1txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						jobstatus = oRec2("jobstatus")
						jobstartdato = oRec2("jobstartdato")
						
						'Response.Write oRec2("jobslutdato")
						'Response.flush
						
						jobslutdato = oRec2("jobslutdato")
						jobForkalkulerettimer = (oRec2("budgettimer") + oRec2("ikkebudgettimer"))
						rekvnr = oRec2("rekvnr")
						jobans1 = oRec2("mid")
						
						if jobForkalkulerettimer <> 0 then
						jobForkalkulerettimer = jobForkalkulerettimer
						else
						jobForkalkulerettimer = 1
						end if
						
						usejoborakt_tp = oRec2("usejoborakt_tp")
						fasttimepris = (oRec2("jobtpris")/jobForkalkulerettimer)
						fastpris = oRec2("fastpris")    
						
                        jobRisiko = oRec2("risiko")

						end if
						oRec2.close
						
						
						
						
						'*** jobANS 2 **'
						jobans2 = 0
						strSQL2 = "SELECT mnavn, mnr, mid, jobans2 FROM job, medarbejdere WHERE job.jobnr = '"& oRec("Tjobnr") &"' AND mid = jobans2"
						oRec2.open strSQL2, oConn, 3 
						if not oRec2.EOF then
						jobans2txt = oRec2("mnavn") & " ("& oRec2("mnr") &")"
						jobans2 = oRec2("mid")
						end if
						oRec2.close
						
						
						
						'*** Lastfakdato **'
						'** tastedato **'
						if komprimeret = "1" then
						fakdatoSQLkri = year(oRec("Tdato")) & "/" & month(oRec("Tdato")) & "/" & day(oRec("Tdato")) 
						else
						fakdatoSQLkri = strAar&"/"&strMrd&"/"&strDag
						end if
					
						lastfakdato = "1/1/2001"
						fid = 0
						faknr = 0
					
						strSQLFAK = "SELECT f.fakdato, f.fid, f.faknr FROM fakturaer f WHERE f.jobid = "& jobid &" AND f.fakdato >= '"& fakdatoSQLkri &"' AND faktype = 0 ORDER BY f.fakdato DESC"
						
						'Response.Write strSQLFAK
						'Response.flush
						
						oRec2.open strSQLFAK, oConn, 3
						if not oRec2.EOF then
							if len(trim(oRec2("fakdato"))) <> 0 then
							lastfakdato = oRec2("fakdato")
							'fid = oRec2("fid")
							'faknr = oRec2("faknr")
							end if
						end if
						oRec2.close
	
	
	end function











 public vlgtMd, vlgtMdLastDay, strExportOskriftDage, ekspTxt, timerTotprAkt

    sub maanedsoversigtArr

        

       if media <> "export" then%>

               
                  <table border="0" width=100% cellpadding="2" cellspacing="0" bgcolor="#ffffff">
                      <%

                        end if



                        ekspTxt = ""
                        lastMid = 0
                        lastKid = 0
                        lastJnr = 0
                        LASTaktid = 0
                        timerTotprAkt = 0

                        'response.write "her"
                        'response.end

                           for m = 0 TO x - 1

                   
                           
    

                                  if lastMid <> medarbtimerArr(m,1) then
        
        
                                   
                          
                                                        if m <> 0 then
                                                    
                                                             if media <> "export" then 


                                                                       

                                                             call medarbMdtot()

                                                             end if

                                                        
                                                            '*** Vis underskrift **'
                                                            select case lto 
                                                            case "mmmi", "intranet - local"              


                                                                if print = "j" then
                                                                %><tr><td colspan="31">
                                                                    <%
                                                                    call underskrift
                                                
                                                                %></td></tr><%
                                                                end if

                                                 

                                                 
                                                                if media = "export" then

                                                                   call underskriftExcel  

                                                                  end if

                                                             end select

                                                
                                                        end if 'm = 0 



                                                     if media <> "export" then 


                                          
                                                    if m = 0 then
                                                         %>
                                                        <tr><td colspan="10" style="padding-top:50px;"><p></p>
                                                        <%else%>
                                                     <tr><td colspan="10"><p style="page-break-before: always; padding-top:50px;"></p>
                                                    <%     
                                                    end if%>

                                       
                                                  <b style="font-size:16px; line-height:18px;">
                                                        <%if len(trim(medarbtimerArr(m,12))) <> 0 then %>
                                                        <%=medarbtimerArr(m,0) & " ["& medarbtimerArr(m,12) &"]" %>
                                                        <%else %>
                                                        <%=medarbtimerArr(m,0)%>
                                                        <%end if  %>
                                                     </b>

                                                          <%if cint(showcpr) = 1 then%>
                               
                                
                                                         <%if len(trim(medarbtimerArr(m,13))) <> 0 then%>
                                                            <span style="font-size:11px;">CPR: <%=medarbtimerArr(m,13)%></span> 
                                                          <%end if %>
                                                         <%end if %>
                                                         </td>
                                     
                                       
                                                    <td colspan="30" align="right" style="padding-right:40px; padding-top:70px;"> <%=monthname(strMrd) &" - "& strAar %></td>
                                                </tr>
                                                <tr bgcolor="#cccccc">
                                                    <td style="width:150px;"><%=joblog2_txt_112 &" "%>/<%=" "& joblog2_txt_113 %></td>
                                                    <td style="width:150px;"><%=joblog2_txt_114 %></td>
                                
                                                    <% end if'excel

                                         
                                        vlgtMd = strDag &"/"& strMrd &"/"& strAar   
                                        vlgtMdNext = dateadd("m", 1, vlgtMd)
                                        vlgtMdLastDay = dateadd("d", -1, vlgtMdNext)
                                        vlgtMdLastDay = day(vlgtMdLastDay)     
                                       
                                        for d = 1 to vlgtMdLastDay 

                                    
                                            if media <> "export" then 
                                           %>
                              
                                            <td style="width:30px; border-left:1px #999999 solid;" align="center"><%=d %></td>
                                            <%else 
                                        
                                                strExportOskriftDage = strExportOskriftDage & d & ";"

                                              
                                            end if %>

                                        <%next 
                                    

                                            if media = "export" then

                                             if cint(mthrap_grpbyakt) = 1 then

                                                strExportOskriftDage = strExportOskriftDage & "Total;"

                                                end if

                                            end if


                                    
                                             if media <> "export" then
                                            

                                                     if cint(mthrap_grpbyakt) = 1 then
                                                    %>
                                                    <td style="width:30px; border-left:1px #999999 solid;" align="center">Total</td>
                                                    <%

                                                     end if
                                            
                                             %>
                                            </tr>

                                    <%      end if
                             end if 'last mid

                              

                               if media <> "export" then


                                 if (cint(mthrap_grpbyakt) = 1 AND medarbtimerArr(m,7) <> LASTaktid) then


                                        if cint(mthrap_grpbyakt) = 1 AND m > 0 then        

                                                                      
                                                                       for d = lastEndKri to vlgtMdLastDay
                                                            
                                                                               if cint(d) <= cint(vlgtMdLastDay) then
                                                                               strTommeTds = strTommeTds & "<td style=""width:30px; border-left:1px #cccccc solid;"" align=""right"">&nbsp;</td>"
                                                                               'strTommeTdsExp = strTommeTdsExp & ";"
                                                                                end if

                                                                       next

                              
                              
                                         Response.write strTommeTds & "<td style=""width:30px; border-left:1px #cccccc solid;"" align=""right"">"& formatnumber(timerTotprAkt, 2) &"</td>"
                                         strTommeTds = ""
                                         timerTotprAkt = 0
                                         end if
                               

                                end if


                                 if lastKid <> medarbtimerArr(m,10) OR lastJnr <> medarbtimerArr(m,5) then

                                        if lastJnr <> "0" then


                                         'Response.write strTommeTds
                                         'strTommeTds = ""

                                         if media <> "export" then 
                                            call jobMdtot()
                                         end if

                                        

                                        end if

                                 %>
                                <tr>
                                    <td colspan="35" style="border-bottom:1px #cccccc solid;border-top:1px #cccccc solid;"><br /><%=medarbtimerArr(m,2) & " ("& medarbtimerArr(m,3)  &")" %><br />
                                        <b><%=medarbtimerArr(m,4) &" ("& medarbtimerArr(m,5) &")"%></b>
                                    </td>
                                     
                                </tr>

                                <%end if


                            else


                                    if (cint(mthrap_grpbyakt) = 1 AND medarbtimerArr(m,7) <> LASTaktid) then

                                            if cint(mthrap_grpbyakt) = 1 AND m > 0 then

                                                            for d = lastEndKri to vlgtMdLastDay
                                                            
                                                                    if cint(d) <= cint(vlgtMdLastDay) then
                                                                    ekspTxt = ekspTxt & ";" 
                                                                    end if

                                                            next

                                                ekspTxt = ekspTxt & formatnumber(timerTotprAkt, 2) &";"
                                                timerTotprAkt = 0
                                            end if

                                    end if

                            end if 'media


                                if media <> "export" then


                                if cint(mthrap_grpbyakt) = 0 then

                                select case right(m, 1)
                                case 0,2,4,6,8
                                bgfCol = "#EFf3ff"
                                case else
                                bgfCol = "#ffffff"
                                end select


                                else

                                    'select case right(m2, 1)
                                    'case 0,2,4,6,8
                                    'bgfCol = "#EFf3ff"
                                    'case else
                                    bgfCol = "#ffffff"
                                    'end select

                                end if
                              
                               

                                    if (cint(mthrap_grpbyakt) = 1 AND medarbtimerArr(m,7) <> LASTaktid) OR cint(mthrap_grpbyakt) = 0 then


                                    %>
                                    <tr style="background-color:<%=bgfCol%>;">
                                        <td style="vertical-align:top;"><%=left(medarbtimerArr(m,6), 25)%></td>
                                        <td style="color:#999999; font-size:10px; line-height:11px; width:175px; vertical-align:top;">
                                                <%if len(trim(medarbtimerArr(m,11))) <> 0 then %>
                                                <i><%=left(medarbtimerArr(m,11), 100) %></i>
                                                <%end if %>
                                        </td>

                                                    <%
                                                        m2 = m2 + 1

                                    end if



                                 else     
                                                        
                                          
                                          if (cint(mthrap_grpbyakt) = 1 AND medarbtimerArr(m,7) <> LASTaktid) OR cint(mthrap_grpbyakt) = 0 then

                                                if cint(mthrap_grpbyakt) = 1 AND m > 0 then
                                                        ekspTxt = ekspTxt & "xx99123sy#z"  
                                                end if
                                          'if media = "export" then
                                          ekspTxt = ekspTxt & Chr(34) & medarbtimerArr(m,0) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,12) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,13) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,2) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,3) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,4) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,5) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,6) & Chr(34) &";"& Chr(34) & medarbtimerArr(m,11) & Chr(34) &";"
                                          end if

                                end if 'exp



                                         


                                                   'Timer fordelt på dage
                                                   if cint(mthrap_grpbyakt) = 1 then
                                                        
                                                        if (medarbtimerArr(m,7) <> LASTaktid) then
                                                        dStkri = 1
                                                        else
                                                        dStkri = lastEndKri
                                                        end if

                                                   else

                                                        dStkri = 1

                                                   end if

                                                   lastEndKri = 32

                                                   for d = dStkri to vlgtMdLastDay 
                                                    
                                                   
                                                    if (d < lastEndKri AND cint(mthrap_grpbyakt) = 1) OR (cint(mthrap_grpbyakt) = 0) then

                                                    vlgtMdCntDays = dateAdd("d", d-1, vlgtMd)
                                                    
                                                    if cDate(medarbtimerArr(m,8)) = cDate(vlgtMdCntDays) then
                                                    timerTxt = formatnumber(medarbtimerArr(m,9), 2)
                                                    timerTxtExp = timerTxt
                                                    timerTotprDag(d) = timerTotprDag(d) + medarbtimerArr(m,9)
                                                    timerTotprDagJob(d) = timerTotprDagJob(d) + medarbtimerArr(m,9)
                                                    timerTotprAkt = timerTotprAkt + medarbtimerArr(m,9)
                                                    lastEndKri = d + 1
                                                    else
                                                    timerTxt = "&nbsp;"
                                                    timerTxtExp = ""
                                                    timerTotprDag(d) = timerTotprDag(d) + 0
                                                    timerTotprDagJob(d) = timerTotprDagJob(d) + 0
                                                    timerTotprAkt = timerTotprAkt + 0
                                                    end if 

                                                    
                                                    

                                                        if media <> "export" then
                                                        %>
                              
                                                            <td style="width:30px; border-left:1px #cccccc solid;" align="right"><%=timerTxt %></td>

                                                        <%
                                                        else
                                                        'timerTxtExp = replace(timerTxtExp, "''", "")
                                                        ekspTxt = ekspTxt & ""& Chr(34) & timerTxtExp & Chr(34) &";"
                                                        end if

                                                     
                                                    end if 'lastEndKri < 31

                                                    next 
                                                        
                                        if media <> "export" then
                                                        

                                              


                                            if cint(mthrap_grpbyakt) = 0 then
                                            %>
                                            </tr>
                                            <%
                                            end if

                                        else  

                                            if cint(mthrap_grpbyakt) = 0 then '(cint(mthrap_grpbyakt) = 1 AND medarbtimerArr(m,7) <> LASTaktid) OR
                                            ekspTxt = ekspTxt & "xx99123sy#z"  
                                            end if

                                        end if

                            lastMid = medarbtimerArr(m,1)
                            lastKid = medarbtimerArr(m,10)
                            lastJnr = medarbtimerArr(m,5)
                            LASTaktid = medarbtimerArr(m,7) 
                         
                            next

                          
                                            
                                            
                        if media <> "export" then


                             


                                            if cint(mthrap_grpbyakt) = 1 AND m > 0 then        

                                                                           
                                                                           for d = lastEndKri to vlgtMdLastDay
                                                            
                                                                                   if cint(d) <= cint(vlgtMdLastDay) then
                                                                                   strTommeTds = strTommeTds & "<td style=""width:30px; border-left:1px #cccccc solid;"" align=""right"">&nbsp;</td>"
                                                                                   'strTommeTdsExp = strTommeTdsExp & ";"
                                                                                    end if

                                                                           next

                                                                    
                                                                

                              
                              
                                             Response.write strTommeTds & "<td style=""width:30px; border-left:1px #cccccc solid;"" align=""right"">"& formatnumber(timerTotprAkt, 2) &"</td>"
                                             strTommeTds = ""
                                             timerTotprAkt = 0
                                             end if
                               

                               

                                        
                                            call jobMdtot()

                                            call medarbMdtot()

                                            if cint(antalM) > 1 then
                                            call medarbMdtotGT()
                                            end if
        

                        %>
                        </table>
                   
                                    <br /><br /><br />&nbsp;
                        <%
                        else    
                                    

                                    if (cint(mthrap_grpbyakt) = 1 AND medarbtimerArr(m,7) <> LASTaktid) then

                                        if cint(mthrap_grpbyakt) = 1 AND m > 0 then

                                                        for d = lastEndKri to vlgtMdLastDay
                                                            
                                                                if cint(d) <= cint(vlgtMdLastDay) then
                                                                ekspTxt = ekspTxt & ";" 
                                                                end if

                                                        next

                                            ekspTxt = ekspTxt & formatnumber(timerTotprAkt, 2) &";"
                                            timerTotprAkt = 0
                                        end if

                                    end if

                        end if 'media=exp


    




    end sub
	%>	

