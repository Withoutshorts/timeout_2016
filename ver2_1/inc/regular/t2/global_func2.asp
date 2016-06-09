




<%
       
       sub subTotaler_gt


            '***************************'
			'*** Sub totaler i bunden af hvert job ved udspec ******'
			'***************************'
			
			

               
				
			    strJobLinie_Subtotal = "<tr>"
				strJobLinie_Subtotal = strJobLinie_Subtotal & "<td style='padding:4px; border-top:1px #CCCCCC solid;' valign=bottom bgcolor=snow><b>Jobtotal:</b></td>"
						
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<td class=lille valign=bottom align=right "&tdstyleTimOms2&" bgcolor=snow>" 
						
						
						

                        strJobLinie_Subtotal = strJobLinie_Subtotal & formatnumber(subbudgettimer, 2)

                        if cint(vis_enheder) = 1 then
					    strJobLinie_Subtotal = strJobLinie_Subtotal & "<br>"
						end if

						
					    if cint(visfakbare_res) = 1 then
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<br>DKK "& formatnumber(subbudget, 2)&"<br>"
                        
                            if jobid = 0 then
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "<span style='color:green; font-size:10px;'>"& basisValISO &" "& formatnumber(TimeprisFaktiskSub, 2) &"</span><br></td>"
                            else
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "</td>"
                            end if
							
								
							
						else
						strJobLinie_Subtotal = strJobLinie_Subtotal & "</td>"
						end if
						

                        '**** Saldo ***'
                        if cint(visPrevSaldo) = 1 then
                        strJobLinie_Subtotal = strJobLinie_Subtotal & "<td class=lille valign=bottom align=right "&tdstyleTimOms2&" bgcolor=snow>"
                        
                            if cint(vis_restimer) = 1 then
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "<span style='color:#999999;'>"& formatnumber(saldoRestimerSub,0) &"</span><br>"
                            end if


                        strJobLinie_Subtotal = strJobLinie_Subtotal & formatnumber(saldoJobSub, 2) &"</td>"
                        end if 

						
						'*** Jobtotaler IALT ***'
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<td class=lille valign=bottom align=right "&tdstyleTimOms2&" bgcolor=snow>" 
						

                        '**** Res timer ***'
                           if cint(vis_restimer) = 1 then
                         strJobLinie_Subtotal = strJobLinie_Subtotal & "<span style='color:#999999;'>"& formatnumber(restimerSubJob,0) &"</span><br>"
                         end if

                         strJobLinie_Subtotal = strJobLinie_Subtotal & formatnumber(subtotaljboTimerIalt,2)

						'*** Enheder ***'
						if cint(vis_enheder) = 1 then
					    strJobLinie_Subtotal = strJobLinie_Subtotal & "<br>enh. " & formatnumber(subJobEnh,2)  
					    end if
                                

                               
						if cint(visfakbare_res) = 1 then
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<br>DKK "&formatnumber(subtotaljobOmsIalt, 2) 
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<br><font class=megetlillesilver>bal.: "&formatnumber(subdbialt, 2)&"</td>"
						else
						strJobLinie_Subtotal = strJobLinie_Subtotal & "</td>"
						end if
						
						
						
				
				
				
				'if cint(jobid) = 0 then
				
						for v = 0 to v - 1
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<td class=lille colspan="& md_split_cspan &" valign=top align=right "&tdstyleTimOms3&" bgcolor=snow>"
                        
                         if cint(vis_restimer) = 1 then
                            if subMedabRestimer(v) <> 0 then
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "<span style='color:#999999;'>"& formatnumber(subMedabRestimer(v),0) &"</span><br>"
                            else
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "&nbsp;<br>"
                            end if
                         subMedabRestimer(v) = 0
                         end if
                        
                        if subMedabTottimer(v) <> 0 then
                        strJobLinie_Subtotal = strJobLinie_Subtotal & formatnumber(subMedabTottimer(v), 2)
                        else
                        strJobLinie_Subtotal = strJobLinie_Subtotal & "&nbsp;<br>"
                        end if
                          
                        subMedabTottimer(v) = 0

                        'strJobLinie_Subtotal = strJobLinie_Subtotal & " % "& formatnumber(subMedabTottimer(v), 2)  
						
						if cint(vis_enheder) = 1 then
						    if subMedabTotenh(v) <> 0 then
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "<br>enh. " & formatnumber(subMedabTotenh(v), 2)
                            else
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "&nbsp;<br>"
                            end if
                        subMedabTotenh(v) = 0
						end if
						        
                               
                            if cint(visfakbare_res) = 1 then
							
                                if omsSubTot(v) <> 0 then
                                strJobLinie_Subtotal = strJobLinie_Subtotal & "<br>DKK "&formatnumber(omsSubTot(v), 2)&"</td>"
                                else
                                strJobLinie_Subtotal = strJobLinie_Subtotal & "<br>&nbsp;</td>"
                                end if
                            omsSubTot(v) = 0
								
							else
							strJobLinie_Subtotal = strJobLinie_Subtotal & "</td>"
							end if
						
							
						next
				 ' end if 'if visMedarb = 1 then
			strJobLinie_Subtotal = strJobLinie_Subtotal & "</tr>"
			

            subbudgettimer = 0
            subbudget = 0
            TimeprisFaktiskSub = 0
            saldoRestimerSub = 0
            saldoJobSub = 0
            restimerSubJob = 0
            subtotaljboTimerIalt = 0
            subJobEnh = 0
            subtotaljobOmsIalt = 0
            subdbialt = 0
            
            
            
          
            

			
			'Response.write strJobLinie_Subtotal
            strJobLinie = strJobLinie & strJobLinie_Subtotal

            

       end sub




        sub tommeCSVfelter


        if cint(jobid) <> 0 then
				                                expTxt = expTxt &";;;;;;;;;;"
				                                else
				                                expTxt = expTxt &";;;;;;;;;;"
				                                end if

                                                if cint(visPrevSaldo) = 1 then
                                                expTxt = expTxt &";"
                                                     if cint(vis_restimer) = 1 then
                                                     expTxt = expTxt &";"
                                                     end if
                                                end if


                                                if cint(vis_restimer) = 1 then
                                                expTxt = expTxt &";"
                                                end if

                                                if cint(vis_enheder) = 1 then
							                    expTxt = expTxt &";" 
							                    end if
							
							                    if cint(visfakbare_res) = 1 then
							                    expTxt = expTxt &";;;"
                                                    
                                                    if jobid = 0 then
                                                    expTxt = expTxt & ";"
                                                    end if

							                 
							                    end if    

       end sub


       public restimerThis, li
       function resTimer(jobid, aktid, medid, md, aar, md_split_cspan, saldo)


       'Response.Write "medid" & medid & "<br>"

       if md_split_cspan = 1 OR medid = 0 then

       
       mdSt = month(datoStart)
       aarSt = year(datoStart)
       mdSl = month(datoSlut)
       aarSL = year(datoSlut)

       if year(datoStart) = year(datoSlut)  then
	   orandval = " AND "
	   else
	   orandval = " OR "
	   end if

       if medid = 0 then 'jobtotal
       medidKri = replace(medarbSQlKri, "m.mid", "rmd.medid")
       grpByKri = "rmd.jobid"
       else
       medidKri = "rmd.medid = "& jobmedtimer(x,4) &""
       grpByKri = "rmd.jobid, rmd.medid"
       end if

       if saldo <> 1 then
       datoKri = "((rmd.md >= "& mdSt &" AND rmd.aar = "& aarSt & ") "& orandval &" (rmd.md <= "& mdSl &" AND rmd.aar = "& aarSl &"))"
       else '** Prev saldo
       datoKri = "(rmd.md < "& mdSt &" AND rmd.aar <= "& aarSt & ")"
       end if
       

       else
       md = md
       aar = aar
       datoKri = "rmd.md = "& md &" AND rmd.aar = "& aar 

       medidKri = "rmd.medid = "& jobmedtimer(x, 4) &""
       grpByKri = "rmd.jobid, rmd.medid"
       end if


       if aktid <> 0 then 'aktivtet valgt
       jobaktIDKri = "rmd.aktid = "& aktid
       else
       jobaktIDKri = "rmd.jobid = "& jobid
       end if

       restimerThis = 0
                                    strSQLrestimer = "SELECT sum(timer) AS restimer FROM ressourcer_md AS rmd WHERE "& jobaktIDKri &""_
                                    &" AND ("& medidKri &") AND ("& datoKri &") GROUP BY "& grpByKri &""

                                    'if medid <> 0 then
                                    'Response.Write li &" (lastMD "& LastMd &"): "& strSQLrestimer & "<br>"
                                    'Response.flush
                                    'end if
			                        
                                    oRec4.open strSQLrestimer, oConn, 3
                                    if not oRec4.EOF then
                                    restimerThis = oRec4("restimer")
                                    end if
                                    oRec4.close

      if resTimerThis <> 0 then
      restimerThis = formatnumber(restimerThis,0)
      else
      restimerThis = ""
      end if

       end function
       

        public guidjobids, guideasyids
        function projktgrpPaaJobids(pid)
        
                            '*** Sætter guiden aktive job ready ***'
					        strSQLjob = "SELECT id AS jid FROM job WHERE "_
					        &" projektgruppe1 = "& pid &" OR " _
					        &" projektgruppe2 = "& pid &" OR " _
					        &" projektgruppe3 = "& pid &" OR " _
					        &" projektgruppe4 = "& pid &" OR " _
					        &" projektgruppe5 = "& pid &" OR " _
					        &" projektgruppe6 = "& pid &" OR " _
					        &" projektgruppe7 = "& pid &" OR " _
					        &" projektgruppe8 = "& pid &" OR " _
					        &" projektgruppe9 = "& pid &" OR " _
					        &" projektgruppe10 = "& pid 
					        
					        oRec4.open strSQLjob, oConn, 3
					        while not oRec4.EOF
					        
					        guidjobids = guidjobids & oRec4("jid") &", #, "
					        
					        oRec4.movenext
					        wend
					        oRec4.close
					        
					        
					        '*** Sætter guiden aktive job Easyreg ready ***'
					        strSQLeasy = "SELECT j.id AS jid, COUNT(a.id) AS antal_aeasy FROM job AS j "_
					        &" LEFT JOIN aktiviteter AS a ON (a.job = j.id AND a.easyreg = 1) WHERE "_
					        &" j.projektgruppe1 = "& pid &" OR " _
					        &" j.projektgruppe2 = "& pid &" OR " _
					        &" j.projektgruppe3 = "& pid &" OR " _
					        &" j.projektgruppe4 = "& pid &" OR " _
					        &" j.projektgruppe5 = "& pid &" OR " _
					        &" j.projektgruppe6 = "& pid &" OR " _
					        &" j.projektgruppe7 = "& pid &" OR " _
					        &" j.projektgruppe8 = "& pid &" OR " _
					        &" j.projektgruppe9 = "& pid &" OR " _
					        &" j.projektgruppe10 = "& pid & " GROUP BY j.id" 
					        
					        'Response.Write strSQLeasy
					        'Response.flush
					        
					        oRec4.open strSQLeasy, oConn, 3
					        while not oRec4.EOF
					        
					        if oRec4("antal_aeasy") <> "0" then
					        guideasyids = guideasyids & oRec4("jid") &", #, "
					        else
					        guideasyids = guideasyids &"0, #, "
					        end if
					        
					        oRec4.movenext
					        wend
					        oRec4.close
        
        
        end function

        function setGuidenUsejob(medid, useJob, del, useEasy)
		
		'Response.Write useJob & "<hr>"
		'Response.Write useEasy & "<hr>"
		
		if len(trim(useJob)) <> 0 then
		len_useJob = len(useJob)
		left_useJob = left(useJob, (len_useJob-3))
		useJob = left_useJob
		end if
		
		'Response.Write useEasy & "<br>"
		'Response.flush
		
		if len(trim(useEasy)) <> 0 then
		len_useEasy = len(useEasy)
		left_useEasy = left(useEasy, (len_useEasy-3))
		useEasy = left_useEasy
		end if
		
		'Response.flush
		if cint(del) = 0 then
		oConn.execute("DELETE FROM timereg_usejob WHERE medarb = "& medid &"")
		end if
		
		
		j = 0
		
        'Response.write "useJob " & useJob & "<br>"
        intuseJob = Split(useJob, "#, ")
        'Respose.write "useEasy " & useEasy & "<br>"
        intUseEasy = split(useEasy, "#, ")
	   	For j = 0 to Ubound(intuseJob)
	   	    
	   	    'Response.Write len(trim(intuseJob(j))) & "Job: " & intuseJob(j) &" Easy: "& intUseEasy(j) &"<br>"
	   	    'Response.flush
	   	    
	   	    if len(trim(intuseJob(j))) > 1 then
	   	    intuseJob(j) = trim(left(intuseJob(j), len(intuseJob(j)) - 2))
	   	    else
	   	    intuseJob(j) = 0
	   	    end if
	   	    
	   	    if len(trim(intUseEasy(j))) > 1 then
	   	    intUseEasy(j) = trim(left(intUseEasy(j), len(intUseEasy(j)) - 2))
	   	    else
	   	    intUseEasy(j) = 0
	   	    end if
	   	
	   	'Response.Write " intusejob DB: " & intuseJob(j) & "<br>"
	   	
	   	if intuseJob(j) <> 0 then
		strSQL = "INSERT INTO timereg_usejob (medarb, jobid, easyreg) VALUES ("& medid &", "& intuseJob(j) &","& intUseEasy(j) &")"
		'Response.write strSQL & "<br>"
		'Response.flush
		oConn.execute(strSQL)
		end if
		next
	    
	    'Response.end
	    
		end function


function sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	%>
	<div id="sideoverskrift" style="position:relative; left:<%=oleft%>px; top:<%=otop%>px; width:<%=owdt%>px; visibility:visible; border:0px #000000 solid; z-index:1200;">
	<!-- sideoverskrift -->
	<table cellspacing="0" cellpadding="5" border="0" width=100%>
	<tr>
	    <td style="width:48px;"><img src="../ill/<%=oimg%>" alt="" border="0"></td>
	    <td align=left style="padding-top:30px; padding-left:20px;"><h3><%=oskrift %></h3></td>
	</tr>
	</table>
	</div>	
	<%
	end function


    
    
    public chkMedarblinier, chklog, chktlffax, chkemail, chkdeak, chklukjob, ign_akttype_inst
	function setFakPreInd()
	
	
	if len(request("FM_chkmed")) <> 0 then
	    if request("FM_chkmed") = 1 then
	    chkMedarblinier = 1
	    else
	    chkMedarblinier = 0
	    end if
	else
	    if len(trim(request.cookies("erp")("tvmedarb"))) <> 0 then
	    chkMedarblinier = request.cookies("erp")("tvmedarb")
	    else
	    chkMedarblinier = 0
	    end if
	end if
	
	Response.Cookies("erp")("tvmedarb") = chkMedarblinier
	
	
	if len(request("FM_chklog")) <> 0 then
	    if request("FM_chklog") = 1 then
	    chklog = 1
	    else
	    chklog = 0
	    end if
	else
	    chklog = request.cookies("erp")("tvlogs")
	end if
	
	if len(request("FM_chktlffax")) <> 0 then
	    if request("FM_chktlffax") = 1 then
	    chktlffax = 1
	    else
	    chktlffax = 0
	    end if
	else
	    chktlffax = request.cookies("erp")("tvtlffax")
	end if
	
	if len(request("FM_chkemail")) <> 0 then
	    if request("FM_chkemail") = 1 then
	    chkemail = 1
	    else
	    chkemail = 0
	    end if
	else
	    chkemail = request.cookies("erp")("tvemail")
	end if
	
	'if len(request("FM_chkdeak")) <> 0 then
	'    if cint(request("FM_chkdeak")) = 1 then
	'    chkdeak = 1
	'    else
	    chkdeak = 0
	'    end if
	'else
	'    if request.cookies("erp")("deak") <> "" then
	'    chkdeak = request.cookies("erp")("deak")
	'    else
	'    chkdeak = 0
	'    end if
	'end if
	
	if len(request("FM_chklukjob")) <> 0 then
	    if cint(request("FM_chklukjob")) = 1 then
	    chklukjob = 1
	    else
	    chklukjob = 0
	    end if
	else
	    if request.cookies("erp")("lukjob") <> "" then
	    chklukjob = request.cookies("erp")("lukjob")
	    else
	    chklukjob = 0
	    end if
	end if

    '** ignorer akttype inst. **'
    if len(request("FM_chkmed")) <> 0 then '** Er form submitted **'
        if len(trim(request("ign_akttype_inst"))) <> 0 then
        ign_akttype_inst = request("ign_akttype_inst")
        else
        ign_akttype_inst = 0
        end if
    else
        if request.cookies("erp")("ignakttype") <> "" then
	    ign_akttype_inst = request.cookies("erp")("ignakttype")
	    else
	    ign_akttype_inst = 0
	    end if
        
    end if
	
	Response.cookies("erp")("lukjob") = chklukjob 
	Response.cookies("erp")("deak") = chkdeak 
	Response.Cookies("erp")("tvemail") = chkemail
	Response.Cookies("erp")("tvtlffax") = chktlffax
	Response.Cookies("erp")("tvlogs") = chklog
	Response.Cookies("erp")("ignakttype") = ign_akttype_inst

	end function
    
    

    public  betCHK9, betCHK8, betCHK7,  betCHK6,  betCHK5,  betCHK4,  betCHK3,  betCHK2,  betCHK1, betCHKu1
	function betalingsbetDage(betbetint,disa)
	
	'if cint(disa) = 1 then
	'disTxt = "DISABLED"
	'else
	'disTxt = ""
	'end if
	    
        betCHK9 = ""
	    betCHK8 = ""
	    betCHK7 = ""
	    betCHK6 = ""
	    betCHK5 = ""
	    betCHK4 = ""
	    betCHK2 = ""
	    betCHK3 = ""
	    betCHK1 = ""
	    betCHKu1 = ""
	
	select case betbetint
	    case 1
	    betCHK1 = "SELECTED"
	    case 2
	    betCHK2 = "SELECTED"
	    case 3
	    betCHK3 = "SELECTED"
	    case 4
	    betCHK4 = "SELECTED"
	    case 5
	    betCHK5 = "SELECTED"
	    case 6
	    betCHK6 = "SELECTED"
	    case 7
	    betCHK7 = "SELECTED"
	    case "-1"
	    betCHKu1 = "SELECTED"
	    case 8
	    betCHK8 = "SELECTED"
	    case 9
	    betCHK9 = "SELECTED"
	    
	    end select %>
	
		
		
		<select id="FM_betbetint" name="FM_betbetint">
		
            <option value="0">Ikke angivet</option>
            <option value="-1" <%=betCHKu1 %>>Angiver selv</option>
                <option value="1" <%=betCHK1 %>>8 dage</option>
                <option value="2" <%=betCHK2 %>>14 dage</option>
                <option value="5" <%=betCHK5 %>>21 dage</option>
                <option value="3" <%=betCHK3 %>>30 dage</option>
                 <option value="6" <%=betCHK6 %>>45 dage</option>
                 <option value="9" <%=betCHK9 %>>60 dage</option>
                <option value="4" <%=betCHK4 %>>Lbn. månd + 15 dage</option>
                <option value="8" <%=betCHK8 %>>Lbn. månd + 30 dage</option>
                <option value="7" <%=betCHK7 %>>Lbn. månd + 45 dage</option>
       
            </select>
	
	<%
	end function

public antalFak
function antalFakturaerKid(kid)
'*** kunde må kun slettes hvis der ikke findes fakturaer **'
			antalfak = 0
			strSQLfak = "SELECT COUNT(fid) AS antalfak FROM fakturaer WHERE fakadr = " & kid & " GROUP BY fakadr"
			oRec4.open strSQLfak, oConn, 3
			if not oRec4.EOF then
			
			antalfak = oRec4("antalfak")
			
			end if
			oRec4.close 
end function





    '*** Akttyper ***'
 	public aty_fakbar, aty_real, aty_pre, aty_enh, aty_on, aty_tfval, aty_medpafak
	function akttyper2009Prop(fakturerbartype)
	
	aty_pre = 0
	aty_fakbar = 0
	aty_real = 0
	aty_medpafak = 0
	
	'Response.Write fakturerbartype
	'Response.flush
	
	
	 strSQL4 = "SELECT aty_id, aty_on, aty_label, aty_desc, aty_on_realhours, "_
	 &" aty_on_invoice, aty_on_invoiceble, aty_pre, aty_enh, et_navn, aty_on_workhours FROM akt_typer "_
	 &" LEFT JOIN enheds_typer ON (et_id = aty_enh) "_
	 &" WHERE aty_id = "& fakturerbartype 
	 
	'Response.Write strSQL4
	'Response.flush
	 
	 oRec4.open strSQL4, oConn, 3
	 
	
	 if not oRec4.EOF then
	
	  if oRec4("aty_on_invoiceble") = 1 then
	  aty_fakbar = 1
	  else
	  aty_fakbar = 0
	  end if
	  
	  if oRec4("aty_on_invoice") = 1 then
	  aty_medpafak = 1
	  else
	  aty_medpafak = 0
	  end if
	  
	  if oRec4("aty_on_realhours") = 1 then
	  aty_real = 1
	  else
	  aty_real = 0
	  end if
	  
	  if oRec4("aty_pre") <> 0 then
	  aty_pre = oRec4("aty_pre")
	  else
	  aty_pre = 0
	  end if
	  
	  aty_on = oRec4("aty_on")
	  
	  aty_enh = oRec4("et_navn")
	  
	  
	  select case oRec4("aty_on_workhours")
	  case 0
	  aty_tfval = ""
	  case 1 '** fradrag
	  aty_tfval = "(-)"
	  case 2 '** tillæg
	  aty_tfval = "(+)"
	  end select
	  
	  'Response.Write "aty_enh" & aty_enh
	  
	 end if 
	 oRec4.close
	
	end function
	
	
	'********** Er uge afsluttet??? ******'
	public ugeNrAfsluttet, showAfsuge, cdAfs 
    function erugeAfslutte(sm_aar, sm_sidsteugedag, sm_mid)
            
            showAfsuge = 1
            ugeNrAfsluttet = "1-1-2044"
            
            strSQLafslut = "SELECT status, afsluttet, uge FROM ugestatus WHERE WEEK(uge, 1) = "& sm_sidsteugedag &" AND YEAR(uge) = "& sm_aar &" AND mid = "& sm_mid
		    'Response.write strSQLafslut
		    oRec3.open strSQLafslut, oConn, 3 
		    if not oRec3.EOF then
    			
			    showAfsuge = 0
			    cdAfs = oRec3("afsluttet")
			    ugeNrAfsluttet = oRec3("uge")
    		
		    end if
		    oRec3.close 
            
            
    end function
    
    
    function valutakoder(i, valuta)
	%>
	<select name="FM_valuta_<%=i %>" id="FM_valuta_<%=i %>" style="width:55px; font-size:9px;">
		    
		    <%
		    strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
    		
    		
		    oRec3.open strSQL3, oConn, 3 
		    while not oRec3.EOF 
    		
		    if cint(valuta) = oRec3("id") then
		    valGrpCHK = "SELECTED"
		    else
		    valGrpCHK = ""
		    end if
		    
		   
		    %>
		    <option value="<%=oRec3("id")%>" <%=valGrpCHK %>><%=oRec3("valutakode")%></option>
		    <%
		    oRec3.movenext
		    wend
		    oRec3.close %>
		    </select>
	<%
	end function
	
	
	public dblKurs
	function valutaKurs(intValuta)
	    '**** Finder aktuel kurs ***'
       dblKurs = 100
       strSQL = "SELECT kurs FROM valutaer WHERE id = " & intValuta
       oRec.open strSQL, oConn, 3
       if not oRec.EOF then
       dblKurs = replace(oRec("kurs"), ",", ".")
       end if 
       oRec.close
	end function
	
	
	public valBelobBeregnet
	function beregnValuta(belob,frakurs,tilkurs)
	        
        'Response.write lastFaknr & " belob: "& belob & "<br>"
        'Response.flush

        if isNull(belob) <> true AND belob <> "-" AND isNull(frakurs) <> true then
        valBelobBeregnet = belob * (frakurs/tilkurs)
        else
        valBelobBeregnet = 0
        end if
    
    if len(valBelobBeregnet) <> 0 then
    valBelobBeregnet = valBelobBeregnet
    else
    valBelobBeregnet = 0
    end if
    
    
    
    valBelobBeregnet = valBelobBeregnet/1
    'Response.Write valBelobBeregnet & "<br>"
    'Response.flush
	end function
	
	
	
	
	
    function sltque(slturl,slttxt,slturlalt,slttxtalt,lft,tp)%>
	
	<%if len(slttxtalt) <> 0 then
	usejaimg = "ja"
	else
	usejaimg = "sletja"
	end if %>
	
	<div id="slet" style="position:absolute; left:<%=lft%>px; top:<%=tp%>px; background-color:#ffffe1; visibility:visible; border:2px #8cAAe6 dashed;">
	<table cellspacing="0" cellpadding="10" border="0" bgcolor="#ffffe1">
	<tr>
	    <td bgcolor="#ffffff" style="border-bottom:1px #999999 solid;"><h4>Slet?</h4> </td>
	    <td bgcolor="#ffffff" align=right style="border-bottom:1px #999999 solid;"><img src="../ill/garbage_information.gif" alt="Slet?" border="0"></td>
	   
	    <%if len(slttxtalt) <> 0 then %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxtalt %>
	    </td>
	   </tr>
	  <tr><td colspan=2>
		<a href=<%=slturlalt %>><img src="../ill/sletja.gif" alt="Ja - slet" border="0"></a>
		</td>
	</tr>
	    
	    <%end if %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxt %>
	    </td>
	   </tr>
	  <tr><td>
		<a href=<%=slturl %>><img src="../ill/<%=usejaimg %>.gif" alt="Ja - slet" border="0"></a>
		</td>
		<td align=right>
		<a href="Javascript:history.back()"><img src="../ill/stop.gif" alt="Nej - tilbage" border="0"></a></td>
	</tr>
	</table>
	</div>
	
	<%
	end function
	
	 function sltque_Small(slturl,slttxt,slturlalt,slttxtalt,lft,tp)%>
	
	<%if len(slttxtalt) <> 0 then
	usejaimg = "ja"
	else
	usejaimg = "sletja"
	end if %>
	
	<div id="Div4" style="position:absolute; width:275px; left:<%=lft%>px; top:<%=tp%>px; padding:2px; background-color:#ffffe1; visibility:visible; border:1px #8cAAe6 solid;">
	<table cellspacing="0" cellpadding="2" border="0" bgcolor="#ffffe1" width=100%>
	<tr>
	    <td bgcolor="#ffffff" style="border-bottom:1px #999999 solid;"><b>Slet?</b> </td>
	    <td bgcolor="#ffffff" align=right style="border-bottom:1px #999999 solid;"><img src="../ill/garbage_information.gif" alt="Slet?" border="0"></td>
	   
	    <%if len(slttxtalt) <> 0 then %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxtalt %>
	    </td>
	   </tr>
	  <tr><td colspan=2>
		<a href=<%=slturlalt %>><img src="../ill/sletja.gif" alt="Ja - nulstil" border="0"></a>
		</td>
	</tr>
	    
	    <%end if %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1" style="padding:5px;"><%=slttxt %>
	    </td>
	   </tr>
	  <tr><td style="padding-left:5px;">
		<a href="Javascript:history.back()" class=rmenu><< Nej, tilbage</a><br />&nbsp;
		</td>
	    <td align=right style="padding-right:25px;">
		<a href=<%=slturl %> class=red>Ja, slet denne >></a><br />&nbsp;</td>
	</tr>
	</table>
	</div>
	
	<%
	end function
	
	function sltquePopup(slturl,slttxt,slturlalt,slttxtalt,lft,tp)%>
	
	<%if len(slttxtalt) <> 0 then
	usejaimg = "ja"
	else
	usejaimg = "sletja"
	end if %>
	
	<div id="Div1" style="position:absolute; left:<%=lft%>px; top:<%=tp%>px; background-color:#ffffe1; visibility:visible; border:2px #8cAAe6 dashed;">
	<table cellspacing="0" cellpadding="10" border="0" bgcolor="#ffffe1">
	<tr>
	    <td bgcolor="#ffffff" style="border-bottom:1px #999999 solid;"><h4>Slet?</h4> </td>
	    <td bgcolor="#ffffff" align=right style="border-bottom:1px #999999 solid;"><img src="../ill/garbage_information.gif" alt="Slet?" border="0"></td>
	   
	    <%if len(slttxtalt) <> 0 then %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxtalt %>
	    </td>
	   </tr>
	  <tr><td colspan=2>
		<a href=<%=slturlalt %>><img src="../ill/sletja.gif" alt="Ja - slet" border="0"></a>
		</td>
	</tr>
	    
	    <%end if %>
	    <tr>
	    <td colspan=2 bgcolor="#ffffe1"><%=slttxt %>
	    </td>
	   </tr>
	  <tr><td>
		<a href=<%=slturl %>><img src="../ill/<%=usejaimg %>.gif" alt="Ja - slet" border="0"></a>
		</td>
		<td align=right>
		<a href="Javascript:window.close()"><img src="../ill/stop.gif" alt="Nej - Luk vindue" border="0"></a></td>
	</tr>
	</table>
	</div>
	
	<%
	end function
	
	function eksportogprint(ptop,pleft,pwdt)
	%>
	<div id=eksport style="position:absolute; background-color:#ffffff; width:<%=pwdt%>px; left:<%=pleft%>px; top:<%=ptop%>px; border:1px #dadada solid; padding:3px 3px 3px 3px;">
    <table cellpadding=5 cellspacing=0 border=0 width=100%>
    <tr>
    <td height=30 bgcolor="#eaeaea" style="border-bottom:1px #cccccc solid;" colspan=2><b>Eksport & Print:</b></td>
    </tr>
	<%
	end function
	
	function eksportogprint09(ptop,pleft,pwdt)
	%>
	<div id=Div2 style="position:relative; background-color:#ffffff; width:<%=pwdt%>px; left:<%=pleft%>px; top:<%=ptop%>px; padding:3px 3px 3px 3px;">
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
    <tr>
    <td bgcolor="#FFFFFF" style="border-bottom:1px #EFf3ff solid; height:20px; padding:10px 10px 0px 10px;" colspan=2><h3>Eksport og Print</h3></td>
    </tr>
    <tr>
    <td valign=top style="padding:10px 10px 10px 10px;">
	<%
	end function
	
	function filteros09(ptop,pleft,pwdt,txt,visning,tdheight)
	
	select case visning
	case 1
	bgt = "#8CAAE6"
	bgt_border = "#5582d2"
	bgtd = ""
	divbg = "#EFF3FF" 
	case 2 '** print
	bgt = "#FFFFe1"
	bgt_border = "#ffff99"
	divbg = "#EFF3FF" 
	case 3
	bgt = "#d6dff5" 'C1D9F0
	bgt_border = "#8CAAE6"
	divbg = "#EFF3FF" 
	case 4
	bgt = "#Cccccc"
	bgt_border = "#999999" 
	divbg = "#FFFFFF"
	
	case 5
	end select
	
	%>
	<div id=Div3 style="position:relative; width:<%=pwdt%>px; left:<%=pleft%>px; top:<%=ptop%>px; border:1px #cccccc solid; overflow:auto; height:<%=tdheight%>px; background-color:<%=divbg%>">
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
    <tr>
    <td bgcolor="<%=bgt %>" style="border-bottom:1px <%=bgt_border %> solid; padding:10px 10px 0px 10px;" colspan=2><h3><%=txt %></h3></td>
    </tr>
    <tr>
    <td valign=top style="padding:10px 10px 10px 10px;">
	<%
	end function
	
	
	function filterheader(ptop,pleft,pwdt,pTxt)
	
	pTxt = replace(global_txt_119, "|", "&")
	
	
	%>
	<div id="filter" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=pwdt %>px; border:1px #8caae6 solid; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#EFF3FF" style="border-bottom:1px #D6DfF5 solid;">
        <img src="../ill/find.png" /></td>
        <td bgcolor="#EFF3FF" align=left style="border-bottom:1px #D6DfF5 solid;"><b><%=pTxt%>:</b></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFFF" style="padding:5px;">
	
    
	
	<%
	end function
   
   
    function filterheaderid(ptop,pleft,pwdt,pTxt,fiVzb,fiDsp,fid,abrel)
	
	pTxt = replace(global_txt_119, "|", "&")
	
	
	%>
	<div id="<%=fid %>" style="position:<%=abrel%>; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=pwdt %>px; border:1px #8caae6 solid; left:<%=pleft%>px; top:<%=ptop%>px; visibility:<%=fiVzb%>; display:<%=vzDsp%>;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#EFF3FF" style="border-bottom:1px #D6DfF5 solid;">
        <img src="../ill/find.png" /></td>
        <td bgcolor="#EFF3FF" align=left style="border-bottom:1px #D6DfF5 solid;"><b><%=pTxt%>:</b></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFFF" style="padding:5px;">
	
    
	
	<%
	end function   
     
     
    function tableDiv(tTop,tLeft,tWdth)
	
	if print = "j" then
	bd = 0
	else
	bd = 1
	end if
	
	%>
	<div id="maintable" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=tWdth%>px; border:<%=bd%>px #8caae6 solid; left:<%=tLeft%>px; top:<%=tTop%>px; visibility:visible;">
    <%
	end function
	
	 function tableDivAbs(tTop,tLeft,tWdth,tHgt,tId, tVzb, tDsp, tZindex)
	
	if print = "j" then
	bd = 0
	else
	bd = 1
	end if
	
	%>
	<div id="<%=tId%>" style="position:absolute; background-color:#ffffff; height:<%=tHgt%>px; padding:3px 3px 3px 3px; width:<%=tWdth%>px; border:<%=bd%>px #8caae6 solid; left:<%=tLeft%>px; top:<%=tTop%>px; visibility:<%=tVzb%>; display:<%=tDsp%>; z-index:<%=tZindex%>; overflow:auto;">
    <%
	end function
	
	
	
	 function tableDivWid(tTop,tLeft,tWdth,tId, tVzb, tDsp)
	 if print = "j" then
	bd = 0
	else
	bd = 1
	end if
	 
	%>
	<div id="<%=tId%>" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=tWdth%>px; border:<%=bd%>px #8caae6 solid; left:<%=tLeft%>px; top:<%=tTop%>px; visibility:<%=tVzb%>; display:<%=tDsp%>; z-index:1000">
    <%
	end function
    
     
     
    function sideinfo(itop,ileft,iwdt)
	%>
	<div id="sideinfo" style="position:relative; background-color:#ffffff; padding:3px 3px 3px 3px; width:<%=iWdt %>px; border:1px red solid; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:visible; z-index:1000000;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#FFFFe1" style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/lifebelt.png" /></td>
        <td bgcolor="#FFFFe1" align=left style="border-bottom:1px #C4c4c4 solid;"><b>Hjælp & Sideinfo:</b></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFfff" style="padding:5px;">
	
    
	
	<%
	end function
	
	
	 function sideinfoId(itop,ileft,iwdt,ihgt,iId,idsp,ivzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
	%>
	
	<script>
	
	$(document).ready(function() {
    
    
    $("#showpagehelp").click(function() {
    
    //alert("her")
    
    $("#pagehelp_bread").show("fast", function() {
        // use callee so don't have to name the function
    //$(this).next().show("fast", arguments.callee);

    $("#pagehelp_bread").css("display", "");
    $("#pagehelp_bread").css("visibility", "visible");
        
    });
	
	

	$("#pagehelp").hide("fast", function() {
	    // use callee so don't have to name the function
	    //$(this).next().show("fast", arguments.callee);
	});


});
    
    

    $("#hidepagehelp").click(function() {


    $("#pagehelp").show("slow", function() {
            // use callee so don't have to name the function
            //$(this).next().show("fast", arguments.callee);
        });


        $("#pagehelp_bread").hide("slow", function() {
            // use callee so don't have to name the function
            //$(this).next().show("fast", arguments.callee);
        });

     });   
	    
	    
        
    });



    
     

</script>
	
	<div id="<%=iId %>" style="position:absolute; background-color:#ffffff; padding:1px 1px 0px 1px; width:<%=iWdt%>px; border:1px silver solid; border-bottom:0px; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:<%=ivzb%>; display:<%=idsp%>; z-index:9000000; overflow:hidden;">
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
    <tr bgcolor="#FF6666"><td align=center>
        <a href="#" id="showpagehelp" class=alt>Hjælp & Sideinfo +</a></td>
     </tr>
    </table>
	</div>
	
	
	<div id="<%=ibId %>" style="position:absolute; background-color:#ffffff; padding:5px 5px 5px 5px; width:<%=ibWdt %>px; height:<%=ibhgt %>px; border:1px silver solid; left:<%=ibleft %>px; top:<%=ibtop %>px; visibility:hidden; display:none; z-index:9000000; overflow:auto;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr bgcolor="#FF6666"><td align=center width=40 style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/lifebelt.png" /></td>
        <td align=left style="border-bottom:1px #C4c4c4 solid;" class=alt><b>Hjælp & Sideinfo</b></td>
         <td style="padding-right:20px;" align="right"><a href="#" id="hidepagehelp" class=alt>[x]</a></td>
       
        </tr>
        <!-- onclick="dsppagehelp('<%=iId%>')" -->
	<tr>
	</table>
    <table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFff" style="padding:5px;">
	
    
	
	<%
	end function
	
    function sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb)
	%>
	
	<script>
	    function sidemsgclose(idthis) {
	        //alert(idthis)
	        document.getElementById(idthis).style.visibility = "hidden"
	        document.getElementById(idthis).style.display = "none"
	    }
	
	
	</script>
	
	 <div id="<%=iId %>" style="position:absolute; background-color:#FFFFFF; padding:3px 3px 3px 3px; width:<%=iWdt %>px; border:1px #999999 solid; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:<%=ivzb%>; display:<%=idsp%>; z-index:1000000;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#FF6666" style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/ikon_message_24.png" /></td>
        <td bgcolor="#FF6666" align=left style="border-bottom:1px #C4c4c4 solid;"><b>Meddelelse</b></td>
	    <td bgcolor="#FF6666" align=right style="border-bottom:1px #C4c4c4 solid; padding-right:5px;"><a href="#" onclick="sidemsgclose('<%=iId%>')" class="vmenu">[x]</a></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFe1" style="padding:5px;">
	
    
	
	<%
	end function
	
	  function sidemsgId2(itop,ileft,iwdt,iId,idsp,ivzb)
	%>
	
	<script>
	    function sidemsgclose2(idthis) {
	        //alert(idthis)
	        document.getElementById(idthis).style.visibility = "hidden"
	        document.getElementById(idthis).style.display = "none"
	    }
	
	
	</script>
	
	<div id="<%=iId %>" style="position:absolute; background-color:#FFFFFF; padding:3px 3px 3px 3px; width:<%=iWdt %>px; border:2px #c4c4c4 solid; left:<%=iLeft%>px; top:<%=iTop%>px; visibility:<%=ivzb%>; display:<%=idsp%>; z-index:1000000;">
    <table cellpadding=2 cellspacing=0 border=0 width=100%>
    <tr><td align=center width=40 bgcolor="#FFFF99" style="border-bottom:1px #C4c4c4 solid;">
        <img src="../ill/ikon_message_24.png" /></td>
        <td bgcolor="#FFFF99" align=left style="border-bottom:1px #C4c4c4 solid;"><b>Meddelelse</b></td>
	    <td bgcolor="#FFFF99" align=right style="border-bottom:1px #C4c4c4 solid; padding-right:5px;"><a href="#" onclick="sidemsgclose2('<%=iId%>')" class="vmenu">[x]</a></td></tr>
	<tr>
	</table>
	<table cellpadding=0 cellspacing=0 border=0 width=100%>
	</tr><td bgcolor="#FFFFFF" valing="top" style="padding:5px;">
	
    
	
	<%
	end function
	
	function opretNy(url, text, otoppx, oleftpx, owdtpx)
	%>
	<div style="position:relative; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td style="padding:1px 0px 0px 10px;">
	<a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	
	function opretNyAB(url, text, otoppx, oleftpx, owdtpx)
	%>
	<div style="position:absolute; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td style="padding:1px 0px 0px 10px;">
	<a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_top"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	
	
	
	function opretNy_blank(url, text, otoppx, oleftpx, owdtpx)
	%>
	<div style="position:relative; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td style="padding:1px 0px 0px 10px;">
	<a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_blank"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" target="_blank"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	
	function opretNyJava(url, text, otoppx, oleftpx, owdtpx, java)
	%>
	<div style="position:relative; top:<%=otoppx%>px; left:<%=oleftpx%>px; width:<%=owdtpx%>px; border:1px #8cAAe6 solid; padding:3px 2px 1px 2px; background-color:#ffffff;">
	<table cellpadding=0 cellspacing=0 border=0 width=100%><tr><td style="border-bottom:0px; padding:1px 0px 0px 10px;">
	<a href='<%=url %>' onclick="<%=java%>" class='vmenu' alt="<%=text %>"><%=text %></a>
        </td><td style="padding:3px 0px 0px 0px; border-bottom:0px;">
        <a href='<%=url %>' class='vmenu' alt="<%=text %>" onclick="<%=java %>"><img src="../ill/add2.png" border="0" /></a>
        </td></tr></table>
    </div>
	<%
	end function
	

	
	
	
	
	
	function insertDelhist(deltype, delid, delnr, delnavn, mid, mnavn)
	
	
	strSQLdelhist = "INSERT INTO delete_hist (deltype, delid, delnr, delnavn, mid, mnavn) VALUES "_
	&" ('"& deltype &"', "& delid &", '"& delnr &"', '"& delnavn &"', "& mid &", '"& mnavn &"')"
	
	oConn.execute(strSQLdelhist)
	
	end function
	
	
	
	
	
	
	
	
	
	public htmlparseTxt
	function htmlreplace(HTMLstring)
	
	
	    txtBRok = replace(HTMLstring, "<br>", "[#br#]") 
        txtBRok = replace(txtBRok, "<br />", "[#br#]") 
        txtBRok = replace(txtBRok, "<br/>", "[#br#]") 
        txtBRok = replace(txtBRok, "<b>", "[#b#]")
        txtBRok = replace(txtBRok, "</b>", "[#/b#]")
        
        txtBRok = replace(txtBRok, "<p>", "[#br#]")
        txtBRok = replace(txtBRok, "</p>", "[#br#]")
        
        txtBRok = replace(txtBRok, "<div>", "[#br#]")
        txtBRok = replace(txtBRok, "</div>", "[#br#]")
       
        txtBRok = replace(txtBRok, "<strong>", "[#strong#]")
        txtBRok = replace(txtBRok, "</strong>", "[#/strong#]")
       
        txtBRok = replace(txtBRok, "<i>", "[#i#]")
        txtBRok = replace(txtBRok, "</i>", "[#/i#]")
        
        txtBRok = replace(txtBRok, "<em>", "[#em#]")
        txtBRok = replace(txtBRok, "</em>", "[#/em#]")
       
        txtBRok = replace(txtBRok, "<u>", "[#u#]")
        txtBRok = replace(txtBRok, "</u>", "[#/u#]")

        txtBRok = replace(txtBRok, "<img", "[#img")
        txtBRok = replace(txtBRok, "</img>", "[#/img#]")
       
        
        HTMLstring = txtBRok
        
        
                Set RegularExpressionObject = New RegExp

                With RegularExpressionObject
                .Pattern = "<[^>]+>"
                .IgnoreCase = True
                .Global = True
                End With

                stripHTMLtags = RegularExpressionObject.Replace(HTMLstring, "")
                htmlparseTxt = replace(stripHTMLtags, "[#", "<")
                htmlparseTxt = replace(htmlparseTxt, "#]", ">")
                Set RegularExpressionObject = nothing
    	
    end function
    
   
 
 
  public assci_formatTxt
 function assci_format(assci_str)
            
            assci_str = replace(assci_str, "&#248;", "ø")
            assci_str = replace(assci_str, "&#230;", "æ")
            assci_str = replace(assci_str, "&#229;", "å")
            assci_str = replace(assci_str, "&#216;", "Ø")
            assci_str = replace(assci_str, "&#198;", "Æ")
            assci_str = replace(assci_str, "&#197;", "Å")
            assci_str = replace(assci_str, "&#214;", "Ö")
            assci_str = replace(assci_str, "&#246;", "ö")
            assci_str = replace(assci_str, "&#220;", "Ü")
            assci_str = replace(assci_str, "&#252;", "ü")
            assci_str = replace(assci_str, "&#196;", "Ä")
            assci_str = replace(assci_str, "&#228;", "ä")
            
            
            assci_formatTxt = assci_str
 
 end function

    
 public utf_formatTxt
 function utf_format(utf_str)
            
            utf_str = replace(utf_str, "ø", "&#248;")
            utf_str = replace(utf_str, "æ", "&#230;")
            utf_str = replace(utf_str, "å", "&#229;")
            utf_str = replace(utf_str, "Ø", "&#216;")
            utf_str = replace(utf_str, "Æ", "&#198;")
            utf_str = replace(utf_str, "Å", "&#197;")
            utf_str = replace(utf_str, "Ö", "&#214;")
            utf_str = replace(utf_str, "ö", "&#246;")
            utf_str = replace(utf_str, "Ü", "&#220;")
            utf_str = replace(utf_str, "ü", "&#252;")
            utf_str = replace(utf_str, "Ä", "&#196;")
            utf_str = replace(utf_str, "ä", "&#228;")
            
            
            utf_formatTxt = utf_str
 
 end function
 
 public jq_formatTxt
 function jq_format(jq_str)
            
            jq_str = replace(jq_str, "ø", "&oslash;")
            jq_str = replace(jq_str, "æ", "&aelig;")
            jq_str = replace(jq_str, "å", "&aring;")
            jq_str = replace(jq_str, "Ø", "&Oslash;")
            jq_str = replace(jq_str, "Æ", "&AElig;")
            jq_str = replace(jq_str, "Å", "&Aring;")
            jq_str = replace(jq_str, "Ö", "&Ouml;")
            jq_str = replace(jq_str, "ö", "&ouml;")
            jq_str = replace(jq_str, "Ü", "&Uuml;")
            jq_str = replace(jq_str, "ü", "&uuml;")
            jq_str = replace(jq_str, "Ä", "&Auml;")
            jq_str = replace(jq_str, "ä", "&auml;")
            jq_str = replace(jq_str, "é", "&eacute;")
            jq_str = replace(jq_str, "É", "&Eacute;")
            jq_str = replace(jq_str, "á", "&aacute;")
            jq_str = replace(jq_str, "Á", "&Aacute;")
            
            jq_formatTxt = jq_str
 
 end function    
    
 public htmlparseCSVtxt
 Function htmlparseCSV(HTMLstring)
    
    
        Set RegularExpressionObject = New RegExp

        With RegularExpressionObject
        .Pattern = "<[^>]+>"
        .IgnoreCase = True
        .Global = True
        End With

        stripHTMLtags = RegularExpressionObject.Replace(HTMLstring, "")
        htmlparseCSVtxt = stripHTMLtags

       
       

        Set RegularExpressionObject = nothing

End Function






function opdaterFeriePl(level)

                '**** Ændrer status på planlagt ferie til afholdt ferie ***'
		        '*** (indlæser ny registrering så historik beholdes) ***'
		        '**** Tjekker om der findes reg. i forvejen *****'
		        '11 Ferie Planlagt
		        '14 Ferie Afholdt
		        
		        '18 Ferie Fridage Planlagt
		        '13 Ferie fridage brugt
		        
		        '*** Opdaterer for alle hvis det er en admin bruger der logger på ***'
		        
		        if level = 1 then
		        
		        for f = 1 to 2
		        
		        if f = 1 then
		        planlagtVal = 11
		        afholdtVal = 14
		        else
		        planlagtVal = 18
		        afholdtVal = 13
		        end if
		        
		        aktid = 0
		        LoginDato = year(now)&"/"& month(now)&"/"&day(now)
		        
		        '** Finder navn og id på afholdt ferie akt. ***'
		        strSQLfeafn = "SELECT a.id, a.navn FROM job j"_
		        &" LEFT JOIN aktiviteter a ON (a.fakturerbar = "& afholdtVal &" AND a.aktstatus = 1 AND a.job = j.id) "_
		        &" WHERE j.jobstatus = 1 AND a.id <> 'NULL' GROUP BY a.id ORDER BY a.id DESC"
		        
		        'AND a.id <> NULL GROUP BY a.id 
		        'Response.Write strSQLfeafn & "<br>"
		        'Response.flush
		        'Response.end
		        
		        oRec3.open strSQLfeafn, oCOnn, 3
		        if not oRec3.EOF then
		        
		        if oRec3("id") <> "" then
		        aktid = oRec3("id")
		        aktnavn = oRec3("navn")
		        end if
		        
		        end if
		        oRec3.close
		        
		        'Response.Write "<br>aktid" & aktid & "<br>"
		        
		        if cdbl(aktid) <> 0 then
		        
		        
		        '** Opdater ferie/feriefri for den bruger der logger på ***'
		        strSQLfepl = "SELECT * FROM timer WHERE tfaktim = "& planlagtVal &" AND tdato BETWEEN '2009-05-01' AND '"& LoginDato &"'"' AND tmnr = "& session("mid")
		        
		        'Response.Write strSQLfepl & "<br>"
		        
		        oRec4.open strSQLfepl, oConn, 3
		        while not oRec4.EOF 
		                
		                indtastningfindes = 0
		                strSQLfeaf = "SELECT timer, tdato FROM timer WHERE tfaktim = "& afholdtVal &" AND tdato = '"& year(oRec4("tdato")) &"/"& month(oRec4("tdato")) & "/"& day(oRec4("tdato")) &"' AND tmnr = "& oRec4("tmnr") 'session("mid")
		                
		                'Response.Write strSQLfeaf & "<br><br>"
		                
		                oRec3.open strSQLfeaf, oCOnn, 3
		                if not oRec3.EOF then
		                '*** ignorer da der allerede finsdes indtastning **'
		                indtastningfindes = 1
		                end if
		                oRec3.close
		                
		                'Response.Write indtastningfindes 
		                
		                
		                
		                if cint(indtastningfindes) = 0 then
		                '*** Indlæser afholdt ferie ***'
		                strSQLfeins = "INSERT INTO timer "_
		                &"("_
		                &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		                &" timerkom, TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		                &" editor, kostpris, offentlig, seraft, godkendtstatus, "_
		                &" godkendtstatusaf, sttid, sltid, valuta, kurs "_
		                &") "_
		                &" VALUES "_
		                &" (" _
		                & replace(oRec4("timer"), ",", ".") &", "& afholdtVal &", "_
		                & "'"& year(oRec4("tdato")) &"/"& month(oRec4("tdato")) & "/"& day(oRec4("tdato")) &"', "_
		                & "'"& oRec4("tmnavn") &"', "_
		                & oRec4("tmnr") &", "_
		                & "'"& oRec4("tjobnavn") &"', "_
		                & "'"& oRec4("tjobnr") &"', "_
		                & "'"& oRec4("tknavn") &"', "_
		                & oRec4("tknr") &", "_
		                & "'"& replace(oRec4("timerkom"), "'", "''") &"', "_
		                & aktid &", "_
		                & "'"& aktnavn &"', "_
		                & oRec4("Taar") &", "_
		                & oRec4("TimePris") &", "_
		                & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', "_
		                & oRec4("fastpris") &", "_
		                & "'"& time &"', "_
		                & "'"& oRec4("editor") &"', "_
		                & replace(oRec4("kostpris"), ",", ".") &", "_
		                & oRec4("offentlig") &", "_
		                & oRec4("seraft") &", "_
		                & oRec4("godkendtstatus") &", "_
		                & "'"& oRec4("godkendtstatusaf") &"', "_
		                & "'"& oRec4("sttid") &"', "_
		                & "'"& oRec4("sltid") &"', "_
		                & oRec4("valuta") &", "_
		                & replace(oRec4("kurs"), ",", ".") &")"
		                
		                
		                'strSQLfeins = "INSERT LOW_PRIORITY INTO timer t1 SELECT * FROM timer t2 WHERE t2.tid = " & oRec4("tid")
		                
		                
		                'Response.Write strSQLfeins & "<br>"
		                'Response.end
		                oConn.execute(strSQLfeins)
		                
		                '*** Sletter den planlagte **'
		                strSQLdel = "DELETE FROM timer WHERE tid = "& oRec4("tid")
		                oConn.execute(strSQLdel)
		                
		                end if
		        
		                
		        
		        
		        oRec4.movenext
		        wend
		        oRec4.close
                
                end if '*** aktid <> 0
                
                next
                
                end if
                
                'Response.end

end function


public SY_usejoborakt_tp, SY_fastpris
function skaljobSync(jobid)

            '** Skal job sync ****'
			SY_usejoborakt_tp = 0
			SY_fastpris = 0
			strSQLj = "SELECT jobnavn, usejoborakt_tp, fastpris FROM job WHERE id = " & jobid
			
			'Response.Write strSQLj
			'Response.end
			oRec4.open strSQLj, oConn, 3
			if not oRec4.EOF then
			
			SY_usejoborakt_tp = oRec4("usejoborakt_tp")
			SY_fastpris = oRec4("fastpris")
			
			end if
			oRec4.close
			
			

end function

function syncJob(jobid)

                call akttyper2009(2)
                 			
			                strSQLaktSum = "SELECT SUM(budgettimer) sumakttimer, fakturerbar, SUM(aktbudgetsum) AS sumaktbudget FROM aktiviteter "_
			                &" WHERE job =  "& jobid & " AND("& aty_sql_fakbar &") AND aktfavorit = 0 GROUP BY job"
			                oRec2.open strSQLaktSum, oConn, 3
			                if not oRec2.EOF then
                			
			                sumakttimer = replace(oRec2("sumakttimer"), ",", ".")
			                sumaktbudget = replace(oRec2("sumaktbudget"), ",", ".")
                			
			                end if
			                oRec2.close
				
				strSQLsync = "UPDATE Job SET budgettimer = "& sumakttimer &", "_
				&" ikkebudgettimer = 0, jobtpris = "& sumaktbudget &" WHERE id = "& jobid
				
				oConn.execute(strSQLsync)
				
				'Response.Write strSQLsync
				'Response.Write " -- her"
				'Response.end

end function


'option explicit 

' Simple functions to convert the first 256 characters 
' of the Windows character set from and to UTF-8.

' Written by Hans Kalle for Fisz
' http://www.fisz.nl

'IsValidUTF8
'  Tells if the string is valid UTF-8 encoded
'Returns:
'  true (valid UTF-8)
'  false (invalid UTF-8 or not UTF-8 encoded string)
function IsValidUTF8(s)
  dim i
  dim c
  dim n

  IsValidUTF8 = false
  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      select case n
      case 1
        exit function
      case 2
        if (c and &HE0) <> &HC0 then
          exit function
        end if
      case 3
        if (c and &HF0) <> &HE0 then
          exit function
        end if
      case 4
        if (c and &HF8) <> &HF0 then
          exit function
        end if
      case else
        exit function
      end select
      i = i + n
    else
      i = i + 1
    end if
  loop
  IsValidUTF8 = true 
end function

'DecodeUTF8
'  Decodes a UTF-8 string to the Windows character set
'  Non-convertable characters are replace by an upside
'  down question mark.
'Returns:
'  A Windows string
function DecodeUTF8(s)
  dim i
  dim c
  dim n

  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      if n = 2 and ((c and &HE0) = &HC0) then
        c = asc(mid(s,i+1,1)) + &H40 * (c and &H01)
      else
        c = 191 
      end if
      s = left(s,i-1) + chr(c) + mid(s,i+n)
    end if
    i = i + 1
  loop
  DecodeUTF8 = s 
end function

'EncodeUTF8
'  Encodes a Windows string in UTF-8
'Returns:
'  A UTF-8 encoded string
function EncodeUTF8(s)
  dim i
  dim c

  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c >= &H80 then
      s = left(s,i-1) + chr(&HC2 + ((c and &H40) / &H40)) + chr(c and &HBF) + mid(s,i+1)
      i = i + 1
    end if
    i = i + 1
  loop
  EncodeUTF8 = s 
end function


Function EncodeUTF8(s)
        Dim i
        Dim c
       
        i = 1
        Do While i <= len(s)
            c = asc(mid(s, i, 1))
            If c >= &H80 Then
                s = left(s, i - 1) + chr(&HC2 + ((c And &H40) / &H40)) + chr(c And &HBF) + mid(s, i + 1)
                i = i + 1
            End If
            i = i + 1
        Loop
        
        's = Replace(s, "Ã,", "&oslash;")
        EncodeUTF8 = s
End Function


public strDay_30
function dato_30(dagDato, mdDato, aarDato)

                if dagDato > 28 then 
				select case mdDato
				case "2"
				    
				    if len(trim(aarDato)) = 2 then
				    aarDato = "20" & aarDato
				    else
				    aarDato = aarDato
				    end if
				    
				    select case aarDato
				    case "2000", "2004", "2008", "2012", "2016", "2020", "2024", "2028", "2032", "2036", "2040", "2044"
				    strDay_30 = 29
				    case else
				    strDay_30 = 28
				    end select
				    
				case "4", "6", "9", "11"
				    if dagDato > 30 then
				    strDay_30 = 30
				    else
				    strDay_30 = dagDato
				    end if
				case else
				strDay_30 = dagDato
				end select
				else
				strDay_30 = dagDato
				end if

end function
%>





  
 
 