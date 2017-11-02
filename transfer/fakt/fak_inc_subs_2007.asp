<% 
public valutakodeSEL, valutaKursFak
function selectAllValuta(a, tp)


'** Faktura Forklak, Lbn timer ell. på aftale **'
select case tp
case 0
jfunc = "opdatervalutAllelinier("&a&",0,"&tp&")"
case 1
jfunc = "opdvalAlleForkalk("&a&",0,'val')"
case else
jfunc = "opdatervalutAllelinier("&a&",0,"&tp&")"
end select

 %>
	

        <select name="FM_valuta_all_<%=a%>" id="FM_valuta_all_<%=a%>" onChange="<%=jfunc %>" style="width:70px;">
		<%strSQL = "SELECT id, valuta, valutakode, kurs FROM valutaer WHERE id <> 0 ORDER BY id " 
		oRec.open strSQL, oConn, 3
		while not oRec.EOF 
		
         if oRec("id") = cint(valuta) then
		 vSEL = "SELECTED"
		 valutakode = oRec("valutakode")
		 valutakodeSEL = valutakode
		 aktuelKurs = oRec("kurs")
		 
		 if func <> "red" then
		 valutaKursFak = aktuelKurs
		 end if
		 
		 else
		 vSEL = ""
		 end if%>
		<option value="<%=oRec("id") %>" <%=vSEL %>><%=oRec("valutakode") %></option>
		<%
		oRec.movenext
		wend
		oRec.close%>
		
		</select>
		&nbsp;<input id="setAll" type="button" value=" <%=erp_txt_499 %> >> " style="font-size:9px;" onClick=<%=jfunc%> />
		
		<%if func = "red" then %><br />
		Kurs gemt på faktura: <b><%=kurs %></b> - Aktuel kurs: <%=aktuelKurs %>
		<%else
		
		'if jobid <> 0 then
		'hfra = "job"
		'else
		'hfra = "aftale"
		'end if 
		%>
		<!--<br /> <span style="font-size:11px; color:#999999;">(valuta nedarvet fra <%=hfra %>)</span>&nbsp;-->
		<%end if %>
	     
        

	     
        

<% 
end function



function timeprisertilFak(ftpset, mtype)


        strSQL4 = "SELECT timepris, timepris_a1, timepris_a2, "_
        &" timepris_a3, timepris_a4, timepris_a5, "_
        &" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta, "_
        &" v0.valutakode AS valkode_0, "_
        &" v1.valutakode AS valkode_1, "_
        &" v2.valutakode AS valkode_2, "_
        &" v3.valutakode AS valkode_3, "_
        &" v4.valutakode AS valkode_4, "_
        &" v5.valutakode AS valkode_5, "_
        &" v0.kurs AS kurs_0, "_
        &" v1.kurs AS kurs_1, "_
        &" v2.kurs AS kurs_2, "_
        &" v3.kurs AS kurs_3, "_
        &" v4.kurs AS kurs_4, "_
        &" v5.kurs AS kurs_5 "_
        &" FROM medarbejdertyper m "_
        &" LEFT JOIN valutaer v0 ON (v0.id = tp0_valuta) "_
        &" LEFT JOIN valutaer v1 ON (v1.id = tp1_valuta) "_
        &" LEFT JOIN valutaer v2 ON (v2.id = tp2_valuta) "_
        &" LEFT JOIN valutaer v3 ON (v3.id = tp3_valuta) "_
        &" LEFT JOIN valutaer v4 ON (v4.id = tp4_valuta) "_
        &" LEFT JOIN valutaer v5 ON (v5.id = tp5_valuta) "_ 
        &" WHERE m.id = "& mtype
        
        'Response.Write strSQL4
        'Response.flush
        
        oRec4.open strSQL4, oConn, 3 
        
        
        if not oRec4.EOF then
        
        
                select case oRec2("timeprisalt")
                case 0
                medarbejderTimepris = oRec4("timepris")
                valutaKode = oRec4("valkode_0")
                valutaID = oRec4("tp0_valuta")
                valutaKodeKurs = oRec4("kurs_0")
                case 1
                medarbejderTimepris = oRec4("timepris_a1")
                valutaKode = oRec4("valkode_1")
                valutaID = oRec4("tp1_valuta")
                valutaKodeKurs = oRec4("kurs_1")
                case 2
                medarbejderTimepris = oRec4("timepris_a2")
                valutaKode = oRec4("valkode_2")
                valutaID = oRec4("tp2_valuta")
                valutaKodeKurs = oRec4("kurs_2")
                case 3
                medarbejderTimepris = oRec4("timepris_a3")
                valutaKode = oRec4("valkode_3")
                valutaID = oRec4("tp3_valuta")
                valutaKodeKurs = oRec4("kurs_3")
                case 4
                medarbejderTimepris = oRec4("timepris_a4")
                valutaKode = oRec4("valkode_4")
                valutaID = oRec4("tp4_valuta")
                valutaKodeKurs = oRec4("kurs_4")
                case 5
                medarbejderTimepris = oRec4("timepris_a5")
                valutaKode = oRec4("valkode_5")
                valutaID = oRec4("tp5_valuta")
                valutaKodeKurs = oRec4("kurs_5")
                end select
        
        '*** Timepriser fundet **'
        if ftpset = 1 then
        ftp = 1
        end if
        
        end if
        oRec4.close 

end function


public timeforbrug 
public lastFakdato, lastFakid
public strfaktidspkt
function faktureredetimerogbelob()
        
        if jobid <> 0 then
        jobaftKri = " jobid = " & jobid 
        else
        jobaftKri = " aftaleid = "& aftid 
        end if
        
        lastFakid = 0
        
		strSQL = "SELECT fid, beloeb, timer, fakdato, tidspunkt, faktype FROM fakturaer WHERE "& jobaftKri &" AND fid <> "& id &" AND faktype = 0 AND medregnikkeioms = 0 ORDER BY fakdato, fid"
		
	    'Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"& strSQL
		'Response.Flush
		
	    oRec.open strSQL, oConn, 3
		while not oRec.EOF 
				
				'strSQL2 = "SELECT Fid, parentfak FROM fakturaer WHERE parentfak = " &oRec("Fid")&" AND faktype = 2"
				'oRec2.open strSQL2, oConn, 3 
				'hasfak = 0
				'if not oRec2.EOF then
				'hasfak = 1
				'end if
				'oRec2.close
				
				
				'Select case oRec("faktype") 
				'case 0
					'if hasfak = 0 then
				'	faktureretBeloeb = faktureretBeloeb + oRec("beloeb")
					'timeforbrug = timeforbrug + oRec("timer")
					'else
					'faktureretBeloeb = faktureretBeloeb
					'timeforbrug = timeforbrug
					'end if
				'case 1
				'faktureretBeloeb = faktureretBeloeb - oRec("beloeb")
				'timeforbrug = timeforbrug - oRec("timer")
				'case 2
				'faktureretBeloeb = faktureretBeloeb + oRec("beloeb")
				'timeforbrug = timeforbrug + oRec("timer")
				'end select
			
			lastFakid = oRec("fid")
			lastFakdato = convertDateYMD(oRec("fakdato")) 
			strfaktidspkt = oRec("tidspunkt")
		oRec.movenext
		Wend
		oRec.close



end function




'***************************************************************
'Fak, Venter og Beløb
'***************************************************************
public medarbejderBeloeb 
public txt 
public medarbid 
public showonfak
public showventer
public venter
public fak
public venter_brugt
public venter_ultimo





'*********************************************'
'**** Sum-Aktiviteter subtotal func **********'
'*********************************************'

	
	public totalbelob
	public intTimer
	public strAktsubtotal
	public subtotalbelob
	public subtotaltimer
	public aktSubTotalAlluMoms
	'public antalenhialt
	
	
	
	function subtotalakt(x, a)
	
	
	
	if a < 0 then
	
	useMedarbejderTimepris = -1
	thisAkttimer(0) = ""
	thisAktBeloeb(0) = ""
	nul = 1
	bgcol = "#ffffff"
	bgcol2 = "#FFFFFF"
	
	hidden = "hidden"
	display = "none"
	
	if func = "red" then
	akt_showonfak = 0
	end if
	
	thistxt = thisAktText(x)
	
	'hidden = "visible"
	'display = ""

   
	end if
	
	
	
	    '********** Andre? sum-aktiviteten *****
	    if a = 0 then
	
	
		        if thisaktfunc(x) = "red" then
				
				
				
				'*** Finder timepris på ekstra /andre? sum-aktiviteten **
				strSQL = "SELECT enhedspris, kurs, valuta, showonfak, rabat, beskrivelse FROM faktura_det WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &" AND valuta = "& thisAktValuta(x)
				'AND enhedspris = "& SQLBless(useMedarbejderTimepris) 
				
				'Response.Write strSQL
				'Response.flush
				
				oRec.open strSQL, oConn, 3 
				while not oRec.EOF 
					
					'Response.Write "her"
					
					if instr(isTPused, "#"& oRec("valuta") & "-"& SQLBlessDOT(formatnumber(oRec("enhedspris"), 2))&"#") = 0 then
					useEkstraMedarbejderTimepris = oRec("enhedspris")
					andreSumValKurs = oRec("kurs")
					valutaKodeKurs = andreSumValKurs
					andreSumVal = oRec("valuta")
					akt_showonfak = oRec("showonfak")
				    aktrabat = oRec("rabat")
				    thistxt = oRec("beskrivelse")
				    
				    end if
					
				oRec.movenext
				wend
				oRec.close 
				
				
				
				if len(useEkstraMedarbejderTimepris) <> 0 then
				useMedarbejderTimepris = useEkstraMedarbejderTimepris
				else
				useMedarbejderTimepris = -2
				end if
				
				useEkstraMedarbejderTimepris = ""
				
				else '*** Opr
		            
		            '**** fastpris / lbn timer ****'
		            if cint(usefastpris) = 0 then '** = lbn timer
		                    
		                    '*** For prod virk DENCKER løsningen ***'
		                    '*** Beregn enhedspris udfra forkalk antal og medarbejder timepris og timeforbrug
		                    '*** + Hvis job = lbn timer
		                    '*** + Hvis antal stk på akt <> 0
		                    '*** + Akt beregningsgrundlag = 2 (stk)
                            '*** + at der ikke er angivet en enhedspris pr. stk.
		                    if thisAktForkalkStk(x) <> 0 AND thisAktBgr(x) = 2 then
		                    
		                        if cdbl(thisAktEnhpris(x)) <> 0 then '** Er der angivet enhedspris
		                        useMedarbejderTimepris = thisAktEnhpris(x) 'thisTimePris(x)
    		                    else
                                '*** beregner timepris udfra timeforbrug og antal stk der skulle produceres ***'
                                useMedarbejderTimepris = ((mbelobtot)/thisAktForkalkStk(x))
                                end if
                            
                            thisAkttimer(x) = 0 'thisAktTimer(x)
		                    thisAktBeloeb(x) = 0 'thisAktBeloeb(x)
    		                akt_showonfak = 0
				            aktrabat = 0
		                    
		                    else '*** Alm løbende timer, baseret pr. medarbejdertimepris på hver akt.
		                    useMedarbejderTimepris = -2
		                    thisAkttimer(x) = 0
		                    thisAktBeloeb(x) = 0
		                    akt_showonfak = 0
				            aktrabat = 0
				            end if
		            
		            else
		            '** fastpris ***
		                
		                
		                '***************************************************************************************
		                '** Beregning mellem jobfastpris eller om det er aktiviteter der skal være grundlag
		               
    		            useMedarbejderTimepris = thisTimePris(x) 'akttimepris 'jobFastprisTp 'thisTimePris(x)
    		            
                        thisAkttimer(x) = thisAktTimer(x)
                        if len(trim(thisAkttimer(x))) <> 0 then
                        thisAkttimer(x) = thisAkttimer(x)
                        else 
                        thisAkttimer(x) = 0		            
                        end if

                        if len(trim(thisAktBeloeb(x))) <> 0 then
                        thisAktBeloeb(x) = thisAktBeloeb(x)
                        else 
                        thisAktBeloeb(x) = 0            
                        end if


                       

                        konsulentMode = 0
                        select case lto
                        case "synergi1"
                        konsulentMode = 1
                        case else
                        konsulentMode = 0
                        end select

		                
                        if thisAkttimer(x) > 0 AND cint(konsulentMode) = 1 then 
                        akt_showonfak = 1
                        else
                        akt_showonfak = 0
                        end if
				        
                        aktrabat = 0
				    
				    end if
				    
		        end if
	
	
	
	nul = 2
	bgcol = "#ffffff"
	bgcol2 = "#FFFFFF"
	hidden = "visible"
	display = ""
	end if '* a = 0
	
	
	
	
	
	
	if a > 0 then
	'Response.Write "a" & a & " MedarbejderTimepris: " & MedarbejderTimepris
	useMedarbejderTimepris = MedarbejderTimepris	
	
	nul = 1
	bgcol = "#FFFFFF"
	bgcol2 = "#FFFFFF"
	
				
				if thisaktfunc(x) = "red" then
						
						'** Finder showonfak på aktiviteten **
						strSQL2 = "SELECT showonfak, rabat, enhedsang FROM faktura_det WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &" AND enhedspris = "& SQLBless(useMedarbejderTimepris) 'SQLBless(thisTimePris(x)) 
						'Response.Write strSQL2
						'Response.Flush
						oRec2.open strSQL2, oConn, 3
						if not oRec2.EOF then
							
							akt_showonfak = oRec2("showonfak")
							aktrabat = oRec2("rabat")
							
				        end if
						oRec2.close
						
						'*** Skal aktiviteten vises? 
						if akt_showonfak = 1 then
						display = ""
						hidden = "visible"
						else
						display = "none"
						hidden = "hidden"
						end if
					
				else
                'akt_showonfak sættes kun på NUL aktiviteten

				display = ""
				hidden = "visible"
				end if
	
	
	
	end if
	
		
		if a = 0 then
		sadiv = 0
		else
		sadiv = 1
		end if
		
		
		'strAktsubtotal = strAktsubtotal & "<tr name='sumaktdiv_"&x&"_"&a&"' id='sumaktdiv_"&x&"_"&a&"' style='visibility:"&hidden&"; display:"&display&"; background-color:"&bgcol2&"; border-top:"&sadiv&"px lightgrey solid; padding:5px 5px 5px 0px;'><td valign=top>"
	    strAktsubtotal = strAktsubtotal & "<tr><td valign=top>"_
	    &"<div name='sumaktdiv_"&x&"_"&a&"' id='sumaktdiv_"&x&"_"&a&"' style=""visibility:"&hidden&"; display:"&display&"; background-color:"&bgcol2&"; border-top:"&sadiv&"px lightgrey solid; padding:5px 5px 5px 0px;"">"
		strAktsubtotal = strAktsubtotal &"<table cellspacing='0' cellpadding='0' border='0' width='680'>"
		
		'*** Kopier til sum FUNKTION ***
	    if a = 0 then
    	        
    	        'if len(trim(thisAktFase(x))) <> 0 then
    	        'thisSamleaktTxt = thisAktFase(x) &", "& thisAktText(x)
    	        'else
    	        thisSamleaktTxt = thisAktText(x)
    	        'end if
    	        
    	       
    	        strAktsubtotal = strAktsubtotal & "<tr><td colspan=9 style=""padding:10px 10px 0px 0px; border-top:1px #cccccc solid; border-bottom:0px #cccccc solid;"">Overfør ovenstående sum-aktiviteter til nedensåtende sum. aktivitet samle linie.<br>"_
                &"<input id='kopier_sum_"&x&"' type=""button"" value="" Overfør til samle linie >> "" style=""font-size:9px;"" onclick=""overtilsum("&x&")"" /><br><br><b>Sum-aktivitet samle linie:</b> (ell. ekstra fak. linie)</td></tr>"
               
    			
    			
    			
	    end if
		
		if a = 1 OR a = 0 then
		strAktsubtotal = strAktsubtotal &"<tr bgcolor=""#D6DFf5"">"_
		&"<td class=lille style=""padding:2px 2px 2px 5px;"">Vis</td>"_
		&"<td class=lille>Antal</td>"_
        &"<td class=lille style=""padding:2px 2px 2px 5px;"">Faktura linje txt.</td>"_
	    &"<td class=lille style=""padding:2px 2px 2px 5px;"">Enh.pris</td>"_
		&"<td class=lille style=""padding:2px 2px 2px 5px;"">Valuta</td>"_
		&"<td class=lille style=""padding:2px 2px 2px 5px;"">Enhed</td>"_
		&"<td class=lille style=""padding:2px 2px 2px 5px;"" align=center>Rabat</td>"_
		&"<td class=lille style=""padding:2px 2px 2px 5px;"" align=center>Momsfri</td>"_
		&"<td class=lille style=""padding:2px 2px 2px 5px;"" align=right>Pris ialt</td></tr>"
	    end if
	    
	    if nul = 1 then
	    tTpx = 5 
	    else
	    tTpx = 2
	    end if
		
		strAktsubtotal = strAktsubtotal &"<tr><td valign=top style=""width:25px; padding:"&tTpx&"px 2px 2px 2px;"">"
		strAktsubtotal = strAktsubtotal &"<input type=hidden name='FM_aktid_"&x&"_"&a&"' value='"&thisaktid(x)&"'>"_
		&"<input type=hidden name='FM_hidden_timepristhis_"&x&"_"&a&"' id='FM_hidden_timepristhis_"&x&"_"&a&"' value='"&useMedarbejderTimepris&"'>"
		
		
		'************************************************'
		'*** Beregner timer pr. sum aktivitet ***********'
		'************************************************'
		
        '*** ved rediger akt ***'
		if thisaktfunc(x) = "red" then
				    
				    prisThis = 0
				    
				    if a <> 0 then
                        if len(trim(valutaId)) <> 0 then
				        valID = valutaId
                        else
                        valID = 1
                        end if
				    else
                        if len(trim(thisAktValuta(x))) <> 0 then
	    			    valID = thisAktValuta(x)
                        else
                        valID = 1
                        end if
				    end if
				    
                    valutaID = valID

					strSQL2 = "SELECT sum(antal) AS antaltimer, sum(aktpris) AS aktpris, beskrivelse, momsfri FROM faktura_det "_
					&" WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &""_
					&" AND enhedspris = "& SQLBless(useMedarbejderTimepris) &""_
					&" AND valuta = "& valID & " GROUP BY aktid, enhedspris, valuta"
					'SQLBless(thisTimePris(x)) 
					
					'Response.Write "x"& x & " - "& strSQL2 & "<br>"
                    'Response.Write "valkutaId: "& valutaId & ", "& thisAktValuta(x) & ", a:"& a & "<br>"
					'Response.Flush
					
					oRec2.open strSQL2, oConn, 3
					
					if not oRec2.EOF then
					
					    if len(oRec2("aktpris")) <> 0 then
						prisThis = oRec2("aktpris")
						end if
					            
					    hiddentimer = oRec2("antaltimer")
						valueTimer = oRec2("antaltimer")
						
						thistxt = oRec2("beskrivelse")
						
						
						
					end if
					oRec2.close
					
					if len(hiddentimer) <> 0 then
					hiddentimer = hiddentimer
					valueTimer = valueTimer
					else
					hiddentimer = 0
					valueTimer = 0
					end if
					
		else '*** henter nye / ikke oprettede **'
		    
		                
		                
                if len(valutaID) <> 0 then
				valutaID = valutaID 
				else
				valutaID = 1
				end if 
				
				
		        
			            '***** Finder timforbrug i periode *****'
				        if a > 0 AND cint(intType) <> 1 then '** Kun på Faktura / Ikke på Kreditnotaer
    						
						    strSQL2 = "SELECT sum(timer.timer) AS antaltimer FROM timer WHERE taktivitetid =" & thisaktid(x) &""_
						    &" AND timepris = "&SQLBless(useMedarbejderTimepris)&" AND "_
						    &" (Tdato BETWEEN '" & stdatoKriAktSpecifik(x) &"' AND '"& slutdato &"') AND valuta = "& valutaID 
    						
                            'stdatoKri
						    'Response.Write strSQL2
						    oRec2.open strSQL2, oConn, 3
						    if not oRec2.EOF then
							    hiddentimer = oRec2("antaltimer")
							    valueTimer = oRec2("antaltimer")
						    end if
						    oRec2.close
    						

                                 pcValueTimer = 0
                                 valueTimerTxt = ""        
                    
                                 if cint(pcSpecial) = 1 AND valueTimer > 0 then
                

                                    '1pc setting grundfos, basere på forretningsområde
                                    valueTimerTxt = formatnumber(valueTimer, 2)
                                    pcValueTimer = valueTimer
                                    valueTimer = 1
                                    hiddentimer = 1
                    

                                end if


					   
						
						        'thistxt = oRec2("beskrivelse")

                                'if len(trim(valueTimerTxt)) <> 0 then
                                'thistxt = thistxt & "("&valueTimerTxt&")"
                                'end if

                         
    						
						    if hiddentimer > 0 then
						    hiddentimer = hiddentimer
						    valueTimer = valueTimer 
						    else
						    hiddentimer = 0
						    valueTimer = 0 
						    end if
						    
						   
					   end if 
    					
    					
    					if a < 0 then	
						    
						    hiddentimer = 0
						    valueTimer = 0
    						
					    end if
			    
			    
			            if a = 0 then
			                
			                '*** For prod virk DENCKER løsningen ***'
		                    '*** Beregn enhedspris udfra forkalk antal og medarbejder timepris og timeforbrug
		                    '*** + Hvis job = lbn timer
		                    '*** + Hvis antal stk på akt <> 0
		                    '*** + Akt beregningsgrundlag = 2 (stk)
		                    
			            
			                '*** Fastpris / lbn timer ***'
			                if cint(usefastpris) = 0 then '*** Lbn timer job
                            '** OR (thisAktForkalkStk(x) <> 0 AND thisAktBgr(x) = 2 AND cint(usefastpris) <> 1) then 
    			            
                            '** Lbn. timer ***'
    			            '** sum af forbrug på akt **'
    			            '** Ell. forklak antal STK (beregner tp udfra antal stk forkalk og timeforbrug på opgaven = DENCKER) '**
    			                
                                if thisAktForkalkStk(x) <> 0 AND thisAktBgr(x) = 2 then
    			                valueTimer = thisAktForkalkStk(x)
					            hiddenTimer = valueTimer
    			                else
    			                valueTimer = thisAktTimer(x)
					            hiddenTimer = valueTimer
					            end if
					        
					               
    			            
			                else
			                '*** Fastpris: 1 / 2: Commi / 3: Salesorder ****'

			                     
        			                
			                            '*** Forkalk på aktivitet **' 0: ingen, 1: timer, 2 stk
					                    select case thisAktBgr(x)
					                    case 0
					                    valueTimer = thisAktTimer(x) '0 ændret fra NUL til antal brugte timer 4.5.2015 == Bedre at fakturere for meget
					                    hiddenTimer = valueTimer
					                    case 1
					                    valueTimer = thisAktForkalk(x)
					                    hiddenTimer = valueTimer
					                    case 2
					                    valueTimer = thisAktForkalkStk(x)
					                    hiddenTimer = valueTimer
                                        
					                    end select 
        					        
                                        '100

                                        'valueTimer = thisAktBeloeb(x)
					                    'hiddenTimer = valueTimer

                                                    
					               
            			            
			                end if
			            
			            end if
			            
			            
					
			        thistxt = thisAktText(x)
			        
                                if len(trim(valueTimerTxt)) <> 0 then
                                thistxt = thistxt & " ("&valueTimerTxt&")"
                                end if
			   
			
     end if '** Opret / rediger akt
     
     
      
	if a <> 0 then	
		
	if thisaktfunc(x) = "red" then
	timepris = SQLBlessDot(formatnumber(useMedarbejderTimepris, 2))
	else
	timepris = useMedarbejderTimepris
	end if	
	
	else
	
	timepris = useMedarbejderTimepris '(akttimepris(x))
	
	end if
		
		
		
		            '*****************************************************************'
		            '**** Skal aktivitet vises på faktura ?? chk vis ***'
		            '*****************************************************************'
		
		

                    '*** Rediger fak og ikke rediger Akt, 
		            '**  dvs den blev fravalgt ved fakopr skal akt aldrig være sået til **'
		             if func = "red" then
		    
                        if thisaktfunc(x) <> "red" then '** Henter allerede oprettede akt. altid = CHECkED
		                chk = ""
                        else
                            if cint(akt_showonfak) = 1 then
                            chk = "CHECKED"
                            else
                            chk = ""
                            end if
                        end if
		
		            else
		    
		                '**** Fastptris => Samle aktivitet slået til ****'
		                if cint(usefastpris) = 1 then '*** Fastpris: 1 / aldrig noget slået til
		            
                                'tstF = 1 'akt_showonfak
                                if cint(akt_showonfak) = 1 then 'Hvis der ligger timer på aktiviteten og konsulentmode = 1 AND nulakt <> 0
                                chk = "CHECKED"
                                else
                                chk = ""
                                end if
		    
		                else '*** lbn timer: 0 / salesorder: 3 / Commission: 2

		                        if a >= 0 then
		                        '*****************************************************************'
		                        '** Denne kan være misvisende da en alternativ akt. (a = 0 akt) **'
		                        '** godt kan være slået til men med 0 timer **'
			                    '*****************************************************************'
			                

                                        if a = 0 then
			                            '*** Produktions løsning, beregnet samlet stk pris udfra timeforbrug (DENCKER) ****'
			                                 

                                             if (thisAktForkalkStk(x) <> 0 AND thisAktBgr(x) = 2) then
        			                         '** i dette tikfælde skal 0 akt. være slået til 
        			          
			                                    if hiddentimer <> 0 AND cint(thisaktCHK(x)) = 1 AND thisAktTimerSum(x) > 0 then
			                                    chk = "CHECKED"
                                                else
				                                chk = ""
				                                end if
			                   
			                                else
			                                chk = ""
			                                end if
        			                        
                                            'chk = ""

                                             '*** NT altid slået til    
                                            select case lto
                                            case "nt"
                                            chk = "CHECKED"
                                            case else
                                            chk = chk
                                            end select

			                            else
        			            
        			                        '** alm lbn. timer, kun a > 0 skal være slået til
        			                        if hiddentimer <> 0 AND cint(thisaktCHK(x)) = 1 AND (thisAktBgr(x) <> 2) then
                                            '** hvis akt er sat til stk. er nul akt. slået til og timforbrugs aktiviteter skal ikke være slået til selvom der er timer på
                                            '** ovenståedne er LAVET om DENCker pr. 04022013 // IGANG
			                                chk = "CHECKED"
			                                else
				                            chk = ""
				                            end if
			                    
			                                
                                           
			                    
        		            
		                                end if '** a = 0
        		            
		            
		                        else '** hvis a < 0 altid slået fra
		                        chk = ""
		                        end if '** a>= 0
		            
		                end if '** fastpris / lbn timer

		            end if '** red



		
		'*************'
		'**** Vis ***'
		'*************'
		strAktsubtotal = strAktsubtotal & "<input type=hidden name='FM_hidden_timerthis_"&x&"_"&a&"' id='FM_hidden_timerthis_"&x&"_"&a&"' value='"&hiddentimer&"'>"
		
		'nulstilfaktimer("&x&","&a&")
		
		'strAktsubtotal = strAktsubtotal & thisaktfunc(x) &" akt_showonfak" & akt_showonfak & " hiddentimer: "& hiddentimer & " thisaktCHK(x): " & thisaktCHK(x)
		
		
		'**** Timer (antal) *****'
		if nul = 1 then '*** alm. Sum akt ***'
		    strAktsubtotal = strAktsubtotal &"<input type=checkbox class=""visakt_"&x&""" name='FM_show_akt_"&x&"_"&a&"' id='FM_show_akt_"&x&"_"&a&"' value='1' "&chk&" onClick='andreEnhprs("&x&","&a&"), setBeloebTot2("&x&")'></td>"
			strAktsubtotal = strAktsubtotal &"<td style='width:65px; padding:10px 2px 2px 2px;' valign=top><div style='width:60px;' name='timeprisdiv_"&x&"_"&a&"' id='timeprisdiv_"&x&"_"&a&"'><b>"& formatnumber(valueTimer, 2) &"</b></div>"
			strAktsubtotal = strAktsubtotal &"<input type='hidden' name='FM_timerthis_"&x&"_"&a&"' id='FM_timerthis_"&x&"_"&a&"' value='"&valueTimer&"'></td>"
		    
		else '** Nul akt / sum samle linje **'
        
            strAktsubtotal = strAktsubtotal &"<input class=""visakt_"&x&""" type=checkbox name='FM_show_akt_"&x&"_"&a&"' id='FM_show_akt_"&x&"_"&a&"' value='1' "&chk&" onClick='setBeloebTot2("&x&"), andreEnhprs("&x&","&a&")'></td>"
			strAktsubtotal = strAktsubtotal &"<td style='width:65px; padding:2px 0px 2px 2px;' valign=top><input type=text name='FM_timerthis_"&x&"_"&a&"' id='FM_timerthis_"&x&"_"&a&"' onKeyup='tjektimer("&x&","&a&"), showborder2("&x&","&a&")' value='"&valueTimer&"' size=3 style='!border: 1px; background-color: #FFFFFF; font-size:10px; border-color: yellowgreen; border-style: solid;'>"_
            &"<input type=hidden name='FM_timerthisSum_"&x&"_"&a&"' id='FM_timerthisSum_"&x&"_"&a&"'  value='"&thisAktTimerSum(x)&"'>"_
			& "<br><input type='button' name='beregn2_"&x&"_"&a&"_a' id='beregn2_"&x&"_"&a&"_a' value='Calc' style='font-size:8px' onClick='hideborder2("&x&","&a&"), tjektimer("&x&","&a&"), andreEnhprs("&x&","&a&"), setBeloebThis2("&x&","&a&")';></td>"
		    'tjektimer("&x&","&a&"), andreEnhprs("&x&","&a&"), setBeloebThis2("&x&","&a&"), hideborder2("&x&","&a&")
		
        end if
		
		
		
		
		
		if nul = 1 then
		pd = "0px"
		wdt = 220
		hgt = 40
		else
		pd = "0px"
		hgt = 40
		wdt = 220
		end if
		
        strAktsubtotal = strAktsubtotal &"<td valign=top style='width:"&wdt+10&"; padding:2px 2px 2px "&pd&";'><textarea name='FM_aktbesk_"&x&"_"&a&"' id='FM_aktbesk_"&x&"_"&a&"' style='width:"&wdt&"px; height:"&hgt&"px; font-size:11px;'>"&thistxt&"</textarea></td>"
		'<input id='button_bold_"&x&"_"&a&"' type='button' value='< fed >' style=""font-size:8px;"" /><input id='button_kurs_"&x&"_"&a&"' type='button' value='< kursiv >' style=""font-size:8px;"" /><input id='button_br_"&x&"_"&a&"' type='button' value='< br >' style=""font-size:8px;""  onClick=""addbr('"&x&"', '"&a&"')"" /><br>
		
		
		'**** Enhedspris ***'
		if nul = 1 then
		    strAktsubtotal = strAktsubtotal & "<td valign=top align='right' style='width:60px; padding:10px 2px 2px 2px;'>"
			strAktsubtotal = strAktsubtotal & "<div align='right' name='enhprisdiv_"&x&"_"&a&"' id='enhprisdiv_"&x&"_"&a&"'><b>"& timepris &"</b></div>"
			strAktsubtotal = strAktsubtotal & "<input type='hidden' name='FM_enhedspris_"&x&"_"&a&"' id='FM_enhedspris_"&x&"_"&a&"' value='"& timepris &"'></td>"
		else '** Nul akt. **'
			strAktsubtotal = strAktsubtotal & "<td valign=top align='right' style='width:60px; padding:1px 2px 2px 2px;'>"
			strAktsubtotal = strAktsubtotal &"<input type='text' name='FM_enhedspris_"&x&"_"&a&"' id='FM_enhedspris_"&x&"_"&a&"' onKeyup='tjektimepris("&x&"), showborder2("&x&","&a&")' value='"& timepris &"'"_ 
			&" style='width:50px; font-size:9px;'></td>"
		end if
		
		
		'*** Valuta ***'
		if nul = 1 then
		strAktsubtotal = strAktsubtotal & "<td align=center valign=top style='width:55px; padding:10px 2px 2px 2px;'>"
		strAktsubtotal = strAktsubtotal & "<div name='valutadiv_"&x&"_"&a&"' id='valutadiv_"&x&"_"&a&"'><b>"& valutaKode &"</b></div>"
	    strAktsubtotal = strAktsubtotal & "<input type='hidden' name='FM_valuta_"&x&"_"&a&"' id='FM_valuta_"&x&"_"&a&"' value='"& valutaId &"'></td>"
		else
		strAktsubtotal = strAktsubtotal & "<td align=center valign=top style='width:55; padding:2px 2px 2px 2px;'>"
		strAktsubtotal = strAktsubtotal & "<select name='FM_valuta_"&x&"_"&a&"' id='FM_valuta_"&x&"_"&a&"' style='width:50px; font-size:9px;' onchange='tjektimer("&x&","&a&"), andreEnhprs("&x&","&a&"), setBeloebThis2("&x&","&a&")'>"
		
		    strSQL5 = "SELECT id, valutakode, grundvaluta, kurs FROM valutaer ORDER BY valutakode"
                        		
            oRec5.open strSQL5, oConn, 3 
            while not oRec5.EOF 
    		
    		if thisaktfunc(x) = "red" then
    		
    		
                if oRec5("id") = andreSumVal then
                valGrpCHK = "SELECTED"
                else
                valGrpCHK = ""
                end if
            
            else
            
                if oRec5("id") = valuta then
                valGrpCHK = "SELECTED"
                else
                valGrpCHK = ""
                end if
            
            end if
            
           
    		strAktsubtotal = strAktsubtotal & " <option value='"&oRec5("id")&"' "&valGrpCHK&">"&oRec5("valutakode")&"</option>"
           
            oRec5.movenext
            wend
            oRec5.close
            
            strAktsubtotal = strAktsubtotal & "</select></td>"
            					   
		
		end if
		
		strAktsubtotal = strAktsubtotal & "<input name='FM_valuta_opr_"&x&"_"&a&"' id='FM_valuta_opr_"&x&"_"&a&"' type='hidden' value='"&valutaId&"'/>"
			
		
		
		           
		            '*** Enheds angivelse ****'
		             'if thisaktfunc(x) = "red" then 
                     'thisakt_enhedsang = aktenhedsang
                     'else
                     if nul <> 1 AND thisAktBgr(x) = 2 AND func <> "red" then
                     thisakt_enhedsang = 1
                     else
                     thisakt_enhedsang = thisaktEnh(x)
                     end if
                     'end if
                     
                     
                    '****** Enhedsangiverlse ****'
                    select case thisakt_enhedsang
                    case -1
                    ehSel_none = "SELECTED"
                    ehSel_nul = ""
		            ehSel_et = ""
		            ehSel_to = ""
		            ehSel_tre = ""
		            ehLabel = "Ingen"
		            case 0
		            ehSel_none = ""
		            ehSel_nul = "SELECTED"
		            ehSel_et = ""
		            ehSel_to = ""
		            ehSel_tre = ""
		            ehLabel = "Pr. time"
		            case 1
		            ehSel_none = ""
		            ehSel_nul = ""
		            ehSel_et = "SELECTED"
		            ehSel_to = ""
		            ehSel_tre = ""
		            ehLabel = "Pr. stk."
		            case 2
		            ehSel_none = ""
		            ehSel_nul = ""
		            ehSel_et = ""
		            ehSel_to = "SELECTED"
		            ehSel_tre = ""
		            ehLabel = "Pr. enhed"
		            case 3
		            ehSel_none = ""
		            ehSel_nul = ""
		            ehSel_et = ""
		            ehSel_to = ""
		            ehSel_tre = "SELECTED"
		            ehLabel = "Pr. km."
		            case else
		            ehSel_none = ""
		            ehSel_nul = "SELECTED"
		            ehSel_et = ""
		            ehSel_to = ""
		            ehSel_tre = ""
		            ehLabel = "Pr. time"
		            end select
            		
            		
            		
            		if nul = 1 then
            		
            		strAktsubtotal = strAktsubtotal & "<td valign=top align=right style='width:65px; padding:10px 2px 2px 2px;'>"
		            strAktsubtotal = strAktsubtotal &"<div name='ehdiv_"&x&"_"&a&"' id='ehdiv_"&x&"_"&a&"'>"& ehLabel &"</div>"_
            		&"<input name='FM_akt_enh_"&x&"_"&a&"' id='FM_akt_enh_"&x&"_"&a&"' type='hidden' value='"& thisakt_enhedsang &"' />"_
            		&"</td>"
            		
            		else
            		
            		
            		
            		strAktsubtotal = strAktsubtotal & "<td valign=top align=right style='width:65px; padding:2px 2px 2px 2px;'>"
		            strAktsubtotal = strAktsubtotal &"<select name='FM_akt_enh_"&x&"_"&a&"' id='FM_akt_enh_"&x&"_"&a&"' style=""width:60px; font-size:9px;"">"_
                    &"<option value='0' "& ehSel_nul &">Pr. time</option>"_
                    &"<option value='1' "& ehSel_et &">Pr. stk.</option>"_
                    &"<option value='2' "& ehSel_to &">Pr. enhed</option>"_
                    &"<option value='3' "& ehSel_tre &">Pr. km.</option>"_
                    &"<option value='-1' "& ehSel_none &">Ingen</option>"_
                    &"</select></td>"
                    
                    end if
                    
                    
            		
        
        
            		
		
		'*** Rabat ****'
	    
	     if thisaktfunc(x) = "red" then 
         thisakt_rabat = aktrabat
         else
             if intRabat <> 0 then
             thisakt_rabat = (intRabat/100)
             else
             thisakt_rabat = 0
             end if
         end if
	    
	    if nul = 1 then
	    
	    strAktsubtotal = strAktsubtotal & "<td valign=top align='right' style='width:50px; padding:10px 2px 2px 2px;'><div name='rabatdiv_"&x&"_"&a&"' id='rabatdiv_"&x&"_"&a&"'>"& (thisakt_rabat*100) &" %</div></td>"
	    strAktsubtotal = strAktsubtotal & "<input type='hidden' name='FM_rabat_"&x&"_"&a&"' id='FM_rabat_"&x&"_"&a&"' value='"&thisakt_rabat&"'>"
		
	    else
        
            akt_rSel0 = ""
            akt_rSel5 = ""
            akt_rSel7 = ""
            akt_rSel10 = ""
            akt_rSel15 = ""
            akt_rSel20 = ""
            akt_rSel25 = ""
            akt_rSel30 = ""
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = ""

            select case (thisakt_rabat*100)
            case 0
		    akt_rSel0 = "SELECTED"
            case 5
            akt_rSel5 = "SELECTED"
            case 7
            akt_rSel7 = "SELECTED"
            case 10
            akt_rSel10 = "SELECTED"
            case 15
            akt_rSel15 = "SELECTED"
            case 20
            akt_rSel20 = "SELECTED"
            case 25
            akt_rSel25 = "SELECTED"
            case 30
            akt_rSel30 = "SELECTED"
            case 40
            akt_rSel40 = "SELECTED"
            case 50
            akt_rSel50 = "SELECTED"
            case 60
            akt_rSel60 = "SELECTED"
            case 75
            akt_rSel75 = "SELECTED"
		    end select
         
           
        
        strAktsubtotal = strAktsubtotal &"<td valign=top align='right' style='width:50px; padding:2px 0px 0px 0px;'><select id='FM_rabat_"&x&"_"&a&"' name='FM_rabat_"&x&"_"&a&"'  style='background-color:#ffffff; width:40px; font-size:9px;' onChange='andreEnhprs("&x&","&a&"), setBeloebThis2("&x&","&a&")'>"
        strAktsubtotal = strAktsubtotal &"<option value='0' "&akt_rSel0&">0%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.05' "&akt_rSel5&">5%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.07' "&akt_rSel7&">7%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.10' "&akt_rSel10&">10%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.15' "&akt_rSel15&">15%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.15' "&akt_rSel20&">20%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.25' "&akt_rSel25&">25%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.30' "&akt_rSel30&">30%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.40' "&akt_rSel40&">40%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.50' "&akt_rSel50&">50%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.60' "&akt_rSel60&">60%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.75' "&akt_rSel75&">75%</option></select></td>"
        
        'strAktsubtotal = strAktsubtotal &"<div id='angivtxt_"&x&"_"&a&"' style='position:relative; left:-10px; top:2px; visibility:hidden; display:none;'><input id=asum_rbt_"&x&"_"&a&" type=button value='Match timepris' style='font-size:9px; width:100px;' onClick='genberegntimeprissumakt("&x&")'/></div></td>"
        
		end if                           
		
		
		'*** Moms fritagelse ***'
		if cint(thisMomsfri(x)) = 1 then
		momsfri_CHK = "CHECKED"
		else
		momsfri_CHK = ""
		end if
		
		if nul = 1 then
		mTpx = 5
		else
		mTpx = 1
		end if
		
		strAktsubtotal = strAktsubtotal & "<td valign=top align='center' style=""width:50px; padding:"& mTpx &"px 2px 2px 2px;""><input id='FM_momsfri_"&x&"_"&a&"' name='FM_momsfri_"&x&"_"&a&"' type=checkbox value=""1"" "&momsfri_CHK&" onClick=""setAllmomsFri("&x&","&a&"), setBeloebTot2("&x&");""/></td>"
		
		
		'**** Beløb *****'
		'*** Rabatberegning ****
	    if thisaktfunc(x) = "red" then
	        
	        prisThis = prisThis
		    intRabatUse = prisThis
		     
		else
		    

            'select case lto
            'case "xintranet - local", "nt"
            'prisThis = 700
            'case else

            if cint(pcSpecial) = 1 AND valueTimer > 0 AND func <> "red" then
            prisThis = pcValueTimer * timepris
            else
		    prisThis = valueTimer * timepris
            end if
            'end select
            
		    
		    if intRabat <> 0 then
		    intRabatUse = (prisThis) - (prisThis * (intRabat/100))
		    else
		    intRabatUse = prisThis
		    end if
		    
		    
		    '*** Valuta omregning ***'
	        call beregnValuta(intRabatUse,valutaKodeKurs,valutaKursFak)
            
            if len(trim(valBelobBeregnet)) <> 0 then
            intRabatUse = valBelobBeregnet
            else
            intRabatUse = 0
            end if
    		    
		end if
		
		
		
		if len(trim(intRabatUse)) <> 0 then
		intRabatUse = intRabatUse
		else
		intRabatUse = 0
		end if
		
		
		
		
		thisbel = SQLBlessDot(formatnumber(intRabatUse, 2))
        'thisbel = 500
		useforTotal = intRabatUse
	    'useforTotal = thisbel	

		'** toal beløb u moms ***'
		
		if cint(thisMomsfri(x)) = 1 then
	    aktSubTotalAlluMoms = aktSubTotalAlluMoms + intRabatUse
	    else
	    aktSubTotalAlluMoms = aktSubTotalAlluMoms
        end if
		
		'**** Beløb ***' 
		strAktsubtotal = strAktsubtotal & "<td valign=top style='width:110px; padding:10px 0px 0px 2px;'>" 
		
            'useforTotal = 1000
            'thisbel = 1000
            'valueTimer = 1    

			strAktsubtotal = strAktsubtotal &"<div align='right' name='belobdiv_"&x&"_"&a&"' id='belobdiv_"&x&"_"&a&"'><b>"& thisbel &" "& valutaKodeSel &"</b></div>"
			strAktsubtotal = strAktsubtotal &"<input type='hidden' name='FM_beloebthis_"&x&"_"&a&"' id='FM_beloebthis_"&x&"_"&a&"' value='"& thisbel &"'></td>"
		
		
		'*** Afslut tabel *****'
		strAktsubtotal = strAktsubtotal &"</tr></table></div></td></tr>"
	
		
	if a >= 0 AND chk = "CHECKED" then
	intTimer = intTimer + valueTimer 'thisAkttimer(x)
	totalbelob = totalbelob + useforTotal 'cdbl(thisAktBeloeb(x))
	subtotalbelob = subtotalbelob + useforTotal 
	subtotaltimer = subtotaltimer + valueTimer
	'timeprisThisjob = thisTimePris(x)
	end if
	
	if func = "red" then
	isTPused = isTPused & ",#"& valutaId &"-"& useMedarbejderTimepris &"#"
	end if
	
	hiddentimer = 0
	valuetimer = 0
	end function
	
	
	
	
	
	
	
	
	'*********************************************************************'
	'******* Medarbejder linier ***********'
	'*********************************************************************'
	
	public ftp, valutaKode, valutaID, valKursThisMed, valutaKodeKurs
	public medarbstrpris, faktot, mbelobtot, n, x, useMedarbejderTimepris
	public nialt, MedarbejderTimepris
	
	function fakturaMedarbLinjer(medlinieid)
	                                venter = 0
	                                showventer = 0
	                                venter_brugt = 0
									venter_ultimo = 0
	                                    
	                                '*** Henter de rigtige dd og medarb info **'
	                                select case medlinieid
	                                case 0
	                                medarbOptionList = medarbOptionList_0
					                medarbHiddenNameList = medarbHiddenNameList_0
					                txt = txt_0 
					                prevSelIndex = 0
	                                case 1
	                                medarbOptionList = medarbOptionList_1
					                medarbHiddenNameList = medarbHiddenNameList_1
					                txt = txt_1
					                prevSelIndex = 0
	                                case 2
	                                medarbOptionList = medarbOptionList_2
					                medarbHiddenNameList = medarbHiddenNameList_2
					                txt = txt_2
					                prevSelIndex = 1
	                                case 3
	                                medarbOptionList = medarbOptionList_3
					                medarbHiddenNameList = medarbHiddenNameList_3
					                txt = txt_3
					                prevSelIndex = 2
	                                end select
	                                
	                                
	                                    
	                               
	                                if medlinieid = 0 then
	                                venter = oRec3("venter_ultimo")
	                                showventer = oRec3("venter_ultimo")
	                                    '** ved fak oprettelse er venter brugt og ultimo altid = 0
	                                    if id <> 0 then '** opret / rediger fak
	                                    venter_brugt = oRec3("venter_brugt")
									    venter_ultimo = oRec3("venter_ultimo")
									    else
									    venter_brugt = 0
									    venter_ultimo = 0
									end if
	                                
	                                end if
	                                
	                                'Response.flush
	                                
	                           
	                            '*** Henter akt. der allerede er med på fak. ved rediger **'
	                            if thisaktfunc(x) = "red" AND medlinieid = 0 then
									
									
								  
								    fak = oRec3("fak")
									fakopr = fak
									
									medarbejderTimepris = oRec3("enhedspris")
									medarbejderBeloeb = oRec3("beloeb")
									'txt = oRec3("tekst") '** medarbnavn
									'medarbid = oRec3("mid")		
									showonfak = oRec3("showonfak")
									akt_showonfak = 0 
									
									
									
									if len(medarbejderTimepris) <> 0 then
									medarbejderTimepris = SQLBlessDOT(formatnumber(medarbejderTimepris, 2))
									else
									medarbejderTimepris = SQLBlessDot(formatnumber(0, 2))
									end if
									
									medrabat = (oRec3("medrabat") * 100)
									'medenhedsang = oRec3("enhedsang")
									
									if len(trim(oRec3("mvaluta"))) <> 0 then
									valutaId = oRec3("mvaluta")
									else
									valutaId = 1
									end if
									    
									    strSQLvkode = "SELECT v.valutakode FROM valutaer v WHERE v.id = "& valutaId
					                    
					                    'Response.Write strSQLvkode
					                    'Response.flush
					                    oRec4.open strSQLvkode, oConn, 3
					                    if not oRec4.EOF then
					                        valutaKode = oRec4("valutaKode")
					                    end if
					                    oRec4.close
					                    
					                valutaKodeKurs = oRec3("fmskurs")
									
									
								
								else
								
								    
										
										'***************************************************************'
										'** Timer medarbejder                                          *' 
										'***************************************************************'
										
										'** Kun medarb. der er timer på ***'
										'** På aktiviteter der ikke er med på fak i forvejen ved rediger faktura **'
										'** Hent ikke timer ved Kreditnota **'
										if medlinieid = 0 AND cint(intType) <> 1 then
										
										strSQL2 = "SELECT sum(timer) AS timer FROM timer WHERE Tmnr = "&oRec3("medarbejderid")&" "_
										&" AND taktivitetid ="& thisaktid(x) &" AND "_
										&" (Tdato BETWEEN '" & stdatoKriAktSpecifik(x) &"' AND '"& slutdato &"')"
										
										
										'Response.Write strSQL2 & "<br>"
										'Response.flush
										oRec2.open strSQL2, oConn, 3 
										
										if not oRec2.EOF then 
											medarbejderTimer = oRec2("timer")
										end if
										oRec2.close
										
										end if
										
										
										'medarbejderTimer = oRec3("timer")
										if len(medarbejderTimer) <> 0 then
											medarbejderTimer = medarbejderTimer
										else
											medarbejderTimer = 0
										end if
										
										usedmedabId = usedmedabId & " Tmnr <> "& oRec3("medarbejderid") &" AND "
										
										'***************************************************************'
										'Timepris pr. medarbejder                                    ***'
										'***************************************************************'
										%>
										<!--#include file="fak_inc_mtimepris_opret_2007.asp" -->
										<%
										
										'***************************************************************'
										'Fak, Venter og Beløb                                          *'
										'***************************************************************'
										
										
										
										'** For at undgå at der vises timer på akt. der ikke er med på
										'** faktura ved redigering, og som fejlagtigt kommer med efter godkend.
										
										fakopr = medarbejderTimer
                                        if len(fakopr) <> 0 then
                                        fakopr = fakopr
                                        else
                                        fakopr = 0
                                        end if
                                        
                                        if func = "red" then
                                        fak = 0
                                        else
                                        fak = fakopr
                                        end if

                                        'venter_brugt = 0
                                        'venter = 0
                                        'venter_ultimo = 0
                                        medarbejderBeloeb = (fak * medarbejderTimepris) 
		                                
		                                
                                        
                                        'txt = oRec3("mnavn") 
                                        'medarbid = oRec3("mid")
                                        showonfak = 0
                                          
                                        										
										
											
								end if '** red / opr **'
								
							
                          
							
							'***********************************************'
							'*** Medarbejder timepris samme som forrige?? **'
							'*** Kalder sum-aktivitet                     **'
							'***********************************************'
							
								if instr(medarbstrpris, "#"& valutaId &"-"& MedarbejderTimepris &"#") = 0 then
									newa = newa + 1
									Redim preserve aval(newa)
									aval(newa) = MedarbejderTimepris
									call subtotalakt(x, newa)
									sa = sa + 1
								end if
							
							
							'Response.write "a: "& newa  &"aval(a): "&aval(newa)&"<br>" 
							'***
							
							
							'** Henter tom sum-akt ***
							call subtotalakt(x, -(n))
							
							
							
							
							'***** Medarbejder timer på akt. timer skt. pris og omsætning  **'
							'** Der bør ikke kunne opstå flere timepriser **'							
							if thisaktfunc(x) <> "red" AND flereTp = 1 then
							mbg = "#DCF5BD"
                            else
							
							select case right(n, 1)
							case 0,2,4,6,8
							mbg = "#ffffff"
                            
							case 1,3,5,7,9
							mbg = "#Eff3FF"
							
                            end select
							
							end if
							
							if medlinieid <> 0 then
							mDsp = ""
							mVzb = "visible"
							trcls = "ae_"& thisaktid(x)
							else
							trcls = "tr_"& thisaktid(x)
							mDsp = ""
							mVzb = "visible"
							end if
							%>	
							<tr bgcolor="<%=mbg %>" id="tr_<%=n%>_<%=x%>" class="<%=trcls %>" style="visibility:<%=mVzb%>; display:<%=mDsp%>;">
							     <!-- Calc -->
								<td valign=top style="padding:2px 2px 0px 2px;"><input type="button" name="beregn_<%=n%>_<%=x%>_a" id="beregn_<%=n%>_<%=x%>_a" value="Calc" onClick="hideborder('<%=x%>','<%=n%>'), enhedsprismedarb('<%=x%>','<%=n%>')" style="font-size:8px;"></td>
								
								<!-- Vis -->
								<td valign=top style="padding:0px 2px 0px 2px;">
								<!--<input type="text" name="FM_mid_<%=n%>_<%=x%>" value="<%=medarbid%>">-->
								<%
								
								
								
								
								'*** finder den næste a- værdi i rækken.    ****'
								'*** Der må ikke være "huller"              ****'
								'*** da javascriptet så fejler              ****'
								
								if instr(medarbstrpris, "#"& valutaId &"-"& MedarbejderTimepris &"#") = 0 then
								a = newa
								else
									
									for intcounter = 0 to newa - 1
										
										if MedarbejderTimepris = aval(intcounter) then
											a = intcounter
										end if
										
									next
								end if
								
								
								 
								%>
								<input type="hidden" name="FM_m_aval_opr_<%=n%>_<%=x%>" id="FM_m_aval_opr_<%=n%>_<%=x%>" value="<%=a%>">
								<input type="hidden" name="FM_m_aval_<%=n%>_<%=x%>" id="FM_m_aval_<%=n%>_<%=x%>" value="<%=a%>">
								
								<%
								
								'*** Kun medarbejdere med mere end 0 timer skal være checked, men kun hvis det er en fakturerbar akt.   **'
								'*** - og hvis dette ellers er tilvalgt via chkMedarblinier (auto tilvælg udspec. af medarbejdere).     **'
								'*** Ved rediger afgøres det udfra showonfak.                                                            **'
								
								
								if showonfak = 1 then
								chkshow = "CHECKED"
								else
								    if (chkMedarblinier = 1 AND fak <> 0 AND cint(thisaktfakbar(x)) = 1) AND thisaktfunc(x) <> "red" then
									chkshow = "CHECKED"
									else
									chkshow = ""
									end if
								end if
								
                                    
                                '** Fak = antal timer
                                select case lto
	                            case "adra"
                                if func <> "red" AND fak <> 0 then
                                chkshow = "CHECKED"
                                end if
                                end select
								
								%>
								
								<!-- Vis på faktura -->
								<input type="checkbox" name="FM_show_medspec_<%=n%>_<%=x%>" <%=chkshow%> id="FM_show_medspec_<%=n%>_<%=x%>" value="show"></td>
								
								<!-- Timer til fakturering -->
								<td valign=top><input type="text" class="FM_medarbtimer" maxlength="<%=maxl%>" name="FM_m_fak_<%=n%>_<%=x%>" id="FM_m_fak_<%=n%>_<%=x%>" onKeyup="offsetfak('<%=x%>','<%=n%>'), tjektimer2('<%=x%>','<%=n%>'), showborder('<%=x%>','<%=n%>')" value="<%=fak%>" style="border:1px yellowgreen solid; font-size:9px; width:30px;">&nbsp;<font class=lillegray><%=fakopr%></font>
								<input type="hidden" name="FM_m_fak_opr_<%=n%>_<%=x%>" id="FM_m_fak_opr_<%=n%>_<%=x%>" value="<%=fak%>" style="font-size:8px;">
								</td>
								
								<!-- Vente timer -->
								<td align=center valign=top style="padding:0px 0px 2px 0px;">
								<font class=lilleroed><%=showventer%></font> &nbsp;
								<input type="text" name="FM_m_venterbrugt_<%=n%>_<%=x%>" id="FM_m_venterbrugt_<%=n%>_<%=x%>" onKeyup="offsetventer('<%=x%>','<%=n%>'), tjektimer4('<%=x%>','<%=n%>')" value="<%=venter_brugt %>" size=1 style="border:1px yellowgreen dashed; font-size:8px; width:25px;">
							    &nbsp;<input type="text" name="FM_m_vent_<%=n%>_<%=x%>" id="FM_m_vent_<%=n%>_<%=x%>" onKeyup="offsetventer('<%=x%>','<%=n%>'), tjektimer4('<%=x%>','<%=n%>')" value="<%=venter_ultimo%>" size=1 style="border:1px #999999 dashed; font-size:8px; width:25px;">
							   
								
							    <input type="hidden" name="FM_m_vent_opr_<%=n%>_<%=x%>" id="FM_m_vent_opr_<%=n%>_<%=x%>" value="<%=venter%>">
								
								</td>
								
								<!-- Medarbejder Tekst -->
								<td valign=top>
								<select name="FM_mid_<%=n%>_<%=x%>" style="width:120px; font-size:9px;" onchange="setmTxt('<%=thisaktid(x)%>', '<%=n %>', '<%=x %>');">
								<%=medarbOptionList %>
								</select>
								<%=medarbHiddenNameList %>
								<input type="hidden" name="FM_m_tekst_<%=n%>_<%=x%>" value="<%=txt%>">
								<input type="hidden" id="prevSelIndex_<%=n%>_<%=x%>" value="<%=prevSelIndex%>">
								
								<%
								'*** er der flere timepriser ??? **'
								if thisaktfunc(x) <> "red" AND flereTp = 1 then%>
								<%=tp%>
								<%end if%>
								</td>
								
								
								<!-- Pris pr. enhed -->
							    <td valign=top align=right><input type="text" name="FM_mtimepris_<%=n%>_<%=x%>" id="FM_mtimepris_<%=n%>_<%=x%>" onKeyup="offsetmtp('<%=x%>','<%=n%>'), tjektimer3('<%=x%>','<%=n%>'), showborder('<%=x%>','<%=n%>')" value="<%=medarbejderTimepris%>" style="font-size:9px; width:50px;"> 
								<input type="hidden" name="FM_mtimepris_opr_<%=n%>_<%=x%>" id="FM_mtimepris_opr_<%=n%>_<%=x%>" value="<%=medarbejderTimepris%>">
								<!-- onFocus="higlightakt('<=x%>','<=n%>','1')" -->
								</td>
								
								
								<!-- Valuta -->
								<td valign=top><select name="FM_mvaluta_<%=n%>_<%=x%>" id="FM_mvaluta_<%=n%>_<%=x%>" style="width:50px; font-size:9px;" onchange="enhedsprismedarb('<%=x%>','<%=n%>')">
								
		                        <%
		                        strSQL5 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
                        		
                        		
		                        oRec5.open strSQL5, oConn, 3 
		                        while not oRec5.EOF 
                        		
		                        if oRec5("id") = valutaId then
		                        valGrpCHK = "SELECTED"
		                        else
		                        valGrpCHK = ""
		                        end if
                        		
                        		
		                        %>
		                        <option value="<%=oRec5("id")%>" <%=valGrpCHK %>><%=oRec5("valutakode")%></option>
		                        <%
		                        oRec5.movenext
		                        wend
		                        oRec5.close
		                        %>
		                        </select>
		                        
                                    <input name="FM_mvaluta_opr_<%=n%>_<%=x%>" id="FM_mvaluta_opr_<%=n%>_<%=x%>" type="hidden" value="<%=valutaId %>" />
								    <input name="FM_mvalutaKode_<%=n%>_<%=x%>" id="FM_mvalutaKode_<%=n%>_<%=x%>" type="hidden" value="<%=valutaKode %>" />
								</td>
								
								<!-- Enheds angivelse -->
								<%
								'*** Enheds angivelse ****
		                        'if thisaktfunc(x) = "red" then 
                                'thismed_enhedsang = medenhedsang
                                'else
                                thismed_enhedsang = thisaktEnh(x) '0 '** her **'
                                'end if
                        		
		                        select case thismed_enhedsang
		                            case -1
                                    ehSel_none = "SELECTED"
                                    ehSel_nul = ""
		                            ehSel_et = ""
		                            ehSel_to = ""
		                            ehLabel = "Ingen"
		                            case 0
		                            ehSel_none = ""
		                            ehSel_nul = "SELECTED"
		                            ehSel_et = ""
		                            ehSel_to = ""
		                            ehLabel = "Pr. time"
		                            case 1
		                            ehSel_none = ""
		                            ehSel_nul = ""
		                            ehSel_et = "SELECTED"
		                            ehSel_to = ""
		                            ehLabel = "Pr. stk."
		                            case 2
		                            ehSel_none = ""
		                            ehSel_nul = ""
		                            ehSel_et = ""
		                            ehSel_to = "SELECTED"
		                            ehLabel = "Pr. enhed"
		                            case 3
		                            ehSel_none = ""
		                            ehSel_nul = ""
		                            ehSel_et = ""
		                            ehSel_to = ""
		                            ehSel_tre = "SELECTED"
		                            ehLabel = "Pr. km."
		                            case else
		                            ehSel_none = ""
		                            ehSel_nul = "SELECTED"
		                            ehSel_et = ""
		                            ehSel_to = ""
		                            ehLabel = "Pr. time"
		                        end select
								
								%>
								<td valign=top align=right style="padding:2px 0px 0px 2px;">
								<select id="FM_med_enh_<%=n%>_<%=x%>" name="FM_med_enh_<%=n%>_<%=x%>" style="width:60px; font-size:9px;" onChange="enhedsprismedarb('<%=x%>','<%=n%>')">
                                    <option value="0" <%=ehSel_nul %>>Pr. time</option>
                                    <option value="1" <%=ehSel_et %>>Pr. stk.</option>
                                    <option value="2" <%=ehSel_to %>>Pr. enhed</option>
                                    <option value="3" <%=ehSel_tre %>>Pr. km.</option>
                                    <option value="-1" <%=ehSel_none%>>Ingen</option>
                                </select>
                                <input name="FM_med_enh_opr_<%=n%>_<%=x%>" id="FM_med_enh_opr_<%=n%>_<%=x%>" type="hidden" value="<%=thismed_enhedsang%>" />
								    
								</td>
								
								<!-- Rabat -->
								<td valign=top align=right style="padding:2px 0px 0px 2px;">
                                <%
                                if thisaktfunc(x) = "red" then 
                                thismed_rabat = medrabat
                                else
                                thismed_rabat = intRabat
                                end if
                                
                                
                                rSel0 = ""
                                rSel5 = ""
                                rSel7 = ""
                                rSel10 = ""
                                rSel15 = ""
                                rSel20 = ""
                                rSel25 = ""
                                rSel30 = ""
                                rSel40 = ""
                                rSel50 = ""
                                rSel60 = ""
                                rSel75 = ""

                                select case (thismed_rabat)
                                case 0
		                        rSel0 = "SELECTED"
                                case 5
		                        rSel5 = "SELECTED"
                                case 7
		                        rSel7 = "SELECTED"
                                case 10
                                rSel10 = "SELECTED"
                                case 15
                                rSel15 = "SELECTED"
                                case 20
                                rSel20 = "SELECTED"
                                case 25
                                rSel25 = "SELECTED"
                                case 30
                                rSel30 = "SELECTED"
                                case 40
                                rSel40 = "SELECTED"
                                case 50
                                rSel50 = "SELECTED"
                                case 60
                                rSel60 = "SELECTED"
                                case 75
                                rSel75 = "SELECTED"
		                        end select
                                 
		                      
		                        
		                        if thismed_rabat <> 0 then
		                        mRabat_opr = thismed_rabat/100
		                        else
		                        mRabat_opr = 0
		                        end if
		                        %>
                                    <select id="FM_mrabat_<%=n%>_<%=x%>" name="FM_mrabat_<%=n%>_<%=x%>" style="width:50px; font-size:9px;" onChange="enhedsprismedarb('<%=x%>','<%=n%>')">
                                        <option value="0"  <%=rSel0%>>0%</option>
                                        <option value="0.05"  <%=rSel5%>>5%</option>
                                        <option value="0.07"  <%=rSel7%>>7%</option>
                                        <option value="0.10" <%=rSel10%>>10%</option>
                                        <option value="0.15" <%=rSel15%>>15%</option>
                                        <option value="0.20" <%=rSel20%>>20%</option>
                                        <option value="0.25" <%=rSel25%>>25%</option>
                                        <option value="0.30" <%=rSel30%>>30%</option>
                                        <option value="0.40" <%=rSel40%>>40%</option>
                                        <option value="0.50" <%=rSel50%>>50%</option>
                                        <option value="0.60" <%=rSel60%>>60%</option>
                                        <option value="0.75" <%=rSel75%>>75%</option>
                                        </select>
								        <input name="FM_mrabat_opr_<%=n%>_<%=x%>" id="FM_mrabat_opr_<%=n%>_<%=x%>" type="hidden" value="<%=mRabat_opr%>" />
								    
								 </td>
								 
								 
								
								<!-- Beløb ialt. -->
								<td valign=top style="padding-left:2;">
								<%
								if thisaktfunc(x) = "red" then
								
								mbelob = medarbejderBeloeb
								
								else
								
								if len(medarbejderBeloeb) <> 0 then
								    if intRabat <> 0 then
								    mbelob = (medarbejderBeloeb) - (medarbejderBeloeb * (intRabat/100)) 
								    else
								    mbelob = medarbejderBeloeb
								    end if
								else
								mbelob = 0
								end if
								    
								    '*** Valuta omregning ***'
								    call beregnValuta(mbelob,valutaKodeKurs,valutaKursFak)
                                    
                                    if len(trim(valBelobBeregnet)) <> 0 then
                                    mbelob = valBelobBeregnet
                                    else
                                    mbelob = 0
                                    end if
								
								end if
								
								if len(trim(mbelob)) <> 0 then
								mbelob = mbelob
								else
								mbelob = 0
								end if
								%>
								<div style="width:80px; background-color:<%=mbg %>; padding-right:3px;" align="right" id="medarbbelobdiv_<%=n%>_<%=x%>"><b><%=SQLBlessDot(formatnumber(mbelob, 2)) &" "& valutakodeSEL%></b></div>
								<input type="hidden" name="FM_mbelob_<%=n%>_<%=x%>" id="FM_mbelob_<%=n%>_<%=x%>" value="<%=SQLBlessDot(formatnumber(mbelob, 2))%>"></td>
								<td valign=top>
								<input type="button" value="+" onClick="decimaler('plus', '<%=n%>', '<%=x%>')" style="font-size:9px; background-color:#8caae6; width:14px; height:17;">
							<input type="button" value="-" onClick="decimaler('minus', '<%=n%>', '<%=x%>')" style="font-size:9px; background-color:#8caae6; width:10px; height:17;">
							</td>
							</tr>
							<% 
					
					medarbstrpris = medarbstrpris & ", #"& valutaId &"-"& MedarbejderTimepris &"#"
					faktot = faktot + fak
					mbelobtot = mbelobtot + mbelob
					
					fak = 0 
					venter = 0
					medarbejderBeloeb = 0
					n = n + 1
	
	
	
	
	
	end function
	
	
    
    public 	gp1, gp2, gp3, gp4, gp5, gp6, gp7, gp8, gp9, gp10 
    function projektgrupper()
    
    
    
                            '**** projektgrupper på job ******************************************
							'Tjekker at der den projgp der findes på akt også findes på jobbet. 
							'Gælder kun job oprettet før 5/11-2004, da der her er lavet en 
							'funktion der tjekker det samme når et job redigeres. 
							'*********************************************************************
							
								jobPgstring = ""
							
								strSQL = "SELECT id, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM job WHERE id = "& jobid
								oRec2.open strSQL, oConn, 3
									if not oRec2.EOF then
									
									strProjektgr1 = oRec2("projektgruppe1")
									strProjektgr2 = oRec2("projektgruppe2")
									strProjektgr3 = oRec2("projektgruppe3")
									strProjektgr4 = oRec2("projektgruppe4")
									strProjektgr5 = oRec2("projektgruppe5")
									strProjektgr6 = oRec2("projektgruppe6")
									strProjektgr7 = oRec2("projektgruppe7")
									strProjektgr8 = oRec2("projektgruppe8")
									strProjektgr9 = oRec2("projektgruppe9")
									strProjektgr10 = oRec2("projektgruppe10")
									
									jobPgstring = "#"&strProjektgr1&"#,#"&strProjektgr2&"#,#"&strProjektgr3&"#,#"&strProjektgr4&"#,#"&strProjektgr5&"#,#"&strProjektgr6&"#,#"&strProjektgr7&"#,#"&strProjektgr8&"#,#"&strProjektgr9&"#,#"&strProjektgr10&"#"
									
									end if
								oRec2.close
							
                                'response.write jobPgstring
							
							
							'*** Medarbejder udspecificering ***
							strSQL3 = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter WHERE id = "& thisaktid(x)
							oRec3.open strSQL3, oConn, 3
							if not oRec3.EOF then
							
								gp1 = oRec3("projektgruppe1")
								gp2 = oRec3("projektgruppe2")
								gp3 = oRec3("projektgruppe3")
								gp4 = oRec3("projektgruppe4")
								gp5 = oRec3("projektgruppe5")
								gp6 = oRec3("projektgruppe6")
								gp7 = oRec3("projektgruppe7")
								gp8 = oRec3("projektgruppe8")
								gp9 = oRec3("projektgruppe9")
								gp10 = oRec3("projektgruppe10")
							
							end if
							oRec3.Close 
							
							'*** Findes gruppen også på jobbet?? ****
							if len(gp1) <> 0 AND instr(jobPgstring, "#"&gp1&"#") > 0 then
							gp1 = gp1
							else
							gp1 = 1
							end if
							
							if len(gp2) <> 0 AND instr(jobPgstring, "#"&gp2&"#") > 0 then
							gp2 = gp2
							else
							gp2 = 1
							end if
							
							if len(gp3) <> 0 AND instr(jobPgstring, "#"&gp3&"#") > 0 then
							gp3 = gp3
							else
							gp3 = 1
							end if
							
							if len(gp4) <> 0 AND instr(jobPgstring, "#"&gp4&"#") > 0 then
							gp4 = gp4
							else
							gp4 = 1
							end if
							
							if len(gp5) <> 0 AND instr(jobPgstring, "#"&gp5&"#") > 0 then
							gp5 = gp5
							else
							gp5 = 1
							end if
							
							if len(gp6) <> 0 AND instr(jobPgstring, "#"&gp6&"#") > 0 then
							gp6 = gp6
							else
							gp6 = 1
							end if
							
							if len(gp7) <> 0 AND instr(jobPgstring, "#"&gp7&"#") > 0 then
							gp7 = gp7
							else
							gp7 = 1
							end if
							
							if len(gp8) <> 0 AND instr(jobPgstring, "#"&gp8&"#") > 0 then
							gp8 = gp8
							else
							gp8 = 1
							end if
							
							if len(gp9) <> 0 AND instr(jobPgstring, "#"&gp9&"#") > 0 then
							gp9 = gp9
							else
							gp9 = 1
							end if
							
							if len(gp10) <> 0 AND instr(jobPgstring, "#"&gp10&"#") > 0 then
							gp10 = gp10
							else
							gp10 = 1
							end if
    
    
    
    
    end function
    
    
    
    
    
    
    
    
    
    
    
    
    public matSubTotal, matSubTotalAll, matSubTotalAlluMoms, lastgrpnavn
    public rSel0, rSel5, rSel7, rSel10, rSel25, rSel40, rSel50, rSel60, rSel70, rSel75 
    
    
    
    
    function materialer(sw)
  

    'Response.Write "her"
    'Response.end
    
    if sw = 1 then 'opret fak (henter fra fak_mat_spec)
        
        matEnhed = oRec("matenhed")
        
        if ikkemoms <> 0 then
        chkIkmoms = "CHECKED"
        else
        chkIkmoms = ""
        end if
    
        chkVis = "CHECKED"
        
    
    else 'rediger fak (henter fra materiale_forbrug)
    
        if len(oRec("matenhed")) <> 0 then
        matEnhed = oRec("matenhed")
        else
        matEnhed = "Stk."
        end if
        
        '**** kun dem der er indtastet til videre fakturering skal være checked ***'
        if oRec("intkode") <> 2 then 
        chkVis = ""
        else
            if func = "red" then
            chkVis = ""
            else
            chkVis = "CHECKED"
            end if
        end if
        
        '** Ved opret Kreditnota, skal materialer ikke være slåettil
        if cint(jq_intType) <> 1 then
        chkVis = chkVis
        else
        chkVis = ""
        end if
        
        if ikkemoms <> 0 then
        chkIkmoms = "CHECKED"
        else
        chkIkmoms = ""
        end if
    
    end if
    
     '*** ÆØÅ **'
    call jq_format(matEnhed)
    matEnhed = jq_formatTxt 
   
    
    if lastgrpnavn <> oRec("matgrpnavn") AND len(trim(oRec("matgrpnavn"))) <> 0 then
    %>
    
    <tr>

         <%
            '*** ÆØÅ **'
            call jq_format(oRec("matgrpnavn"))
            matgruppenavn = jq_formatTxt    
            %>

    <td colspan=13><b><%=matgruppenavn %></b></td>
    
    </tr>
    
    
    <%
    lastgrpnavn = oRec("matgrpnavn")
    end if


    '** NT commisionsordre ***
    if cint(jq_fastpris) = 2 then '** faktura skal hægtes på leverandør

            
            matSalgspris = (oRec("matsalgspris")/1 - oRec("matkobspris")/1) '(oRec("matsalgspris")/1 * matAva/1)
            matSalgspris = formatnumber(matSalgspris, 4)
            matSalgspris = replace(matSalgspris, ".", "")
          

    else
            matSalgspris = oRec("matsalgspris")
    end if
    %>
    <tr>
        
        <td><input type="button" name="beregn_<%=m%>_m" id="beregn_<%=m%>_m" value="Calc" onClick="beregnmatpris('<%=m%>','1'), hidebordermat('<%=m%>')" style="font-size:8px;"></td>
        <td>
            <input id="FM_matgrp_<%=m%>" name="FM_matgrp_<%=m%>" value="<%=oRec("matgrp") %>" type="hidden" />
            <input id="FM_vis_<%=m %>" name="FM_vis_<%=m %>" value="1" <%=chkVis %> type="checkbox" onclick="beregnmatpris('<%=m%>','1')" />
            <input id="FM_matid_<%=m %>" name="FM_matid_<%=m %>" value="<%=oRec("matid") %>" type="hidden" />
            <input id="FM_mfid_<%=m %>" name="FM_mfid_<%=m %>" value="<%=oRec("mfid") %>" type="hidden" />
            <input id="FM_mfusrid_<%=m %>" name="FM_mfusrid_<%=m %>" value="<%=oRec("mfusrid") %>" type="hidden" /></td>
        <td>

             <%
            '*** ÆØÅ **'
            call jq_format(matnavn)
            matnavn = jq_formatTxt    
            %>

            <input id="FM_matnavn_<%=m %>" name="FM_matnavn_<%=m %>" type="text" value="<%=replace(matnavn, chr(34), "''")  %>" style="font-size:9px; width:120px;" /></td>
        <td>

            <%
            '*** ÆØÅ **'
            call jq_format(oRec("matvarenr"))
            matvarenr = jq_formatTxt    
            %>

            <input id="FM_matvarenr_<%=m %>" name="FM_matvarenr_<%=m %>" type="text" value="<%=matvarenr %>" style="font-size:9px; width:40px;" /></td>
        <td>
            <%if jobid <> 0 then 'job eller service aft. 
               
                strSQLakt = "SELECT id AS aktid, navn AS aktnavn, fase FROM aktiviteter WHERE job = "& jobid &" ORDER BY fase, navn"
                'response.write strSQLakt
                'response.flush
                 
               %>
                         
            <select id="FM_mataktid_<%=m %>" name="FM_mataktid_<%=m %>" style="font-size:9px; width:60px;" >
                <option value="0">Ingen</option>
                
                <% 
                lastaktfase = ""
                 
                oRec6.open strSQLakt, oConn, 3
                while not oRec6.EOF
    
                    if isNull(oRec6("fase")) <> false then

                            if len(trim(oRec6("fase"))) <> 0 AND lastaktfase <> oRec6("fase") then

                            '*** ÆØÅ **'
                             call jq_format(oRec6("fase"))
                             mataktfase = jq_formatTxt
                   

                            %>
                             <option disabled><%=mataktfase %></option>
                            <%
                            lastaktfase = oRec6("fase")
                            end if 

                    end if
                    
                    if cdbl(oRec6("aktid")) = cdbl(mataktid) then
                    aktidsel = "SELECTED"
                    else
                    aktidsel = ""
                    end if


                          '*** ÆØÅ **'
                          call jq_format(oRec6("aktnavn"))
                          mataktnavn = jq_formatTxt
                          

                    %>
                    <option value="<%=oRec6("aktid") %>" <%=aktidsel %>><%=mataktnavn %></option>
                    <%


                    
                oRec6.movenext    
                wend 
                oRec6.close  %>
                    

                

            </select>
               

            <%else %>
             <input id="FM_mataktid_<%=m %>" name="FM_mataktid_<%=m %>" value="0" type="hidden" />
            <%end if %>
        </td>
        <td>

            <!-- SQLBlessDOT(formatnumber(oRec("matantal"), 2))  -->

            <input id="FM_matantal_opr_<%=m %>" type="hidden" value="<%=matantal %>" />
            <input id="FM_matantal_<%=m %>" name="FM_matantal_<%=m %>" type="text" value="<%= SQLBlessDOT(formatnumber(matantal, 2))%>" onkeyup="offsetmatant('<%=m %>'), tjekmatantal('<%=m %>'), showbordermat('<%=m%>')" style="font-size:9px; width:50px;" /></td>
        <td> 
            <input id="FM_matenhedspris_opr_<%=m %>" type="hidden" value="<%=matSalgspris%>" />
            <input id="FM_matenhedspris_<%=m %>" name="FM_matenhedspris_<%=m %>" type="text" value="<%=matSalgspris %>" onkeyup="offsetmatenhpris('<%=m %>'), tjekmatehpris('<%=m %>'), showbordermat('<%=m%>')" style="font-size:9px; width:60px;" /></td>
        <td align=center><%=valutaMatKode%>
        <input id="FM_matvaluta_<%=m %>" name="FM_matvaluta_<%=m %>" type="hidden" value="<%=valutaMatId%>" />
        </td>
        <td> <input id="FM_matenhed_<%=m %>" name="FM_matenhed_<%=m %>" type="text" value="<%=matEnhed%>" style="font-size:9px; width:40px;" /></td>
        <td>
                <select id="FM_matrabat_<%=m%>" name="FM_matrabat_<%=m%>" style="width:30px; font-size:9px;" onChange="beregnmatpris('<%=m%>','1')">
                
                <%
                
		    
            rSel0 = ""
            rSel5 = ""
            rSel7 = ""
            rSel10 = ""
            rSel15 = ""
            rSel25 = ""
            rSel30 = ""
            rSel40 = ""
            rSel50 = ""
            rSel60 = ""
            rSel75 = ""

            select case (intMatRabat)
            case 0
		    rSel0 = "SELECTED"
            case 5
		    rSel5 = "SELECTED"
            case 7
		    rSel7 = "SELECTED"
            case 10
            rSel10 = "SELECTED"
            case 15
            rSel15 = "SELECTED"
            case 20
            rSel20 = "SELECTED"
            case 25
            rSel25 = "SELECTED"
            case 30
            rSel30 = "SELECTED"
            case 40
            rSel40 = "SELECTED"
            case 50
            rSel50 = "SELECTED"
            case 60
            rSel60 = "SELECTED"
            case 75
            rSel75 = "SELECTED"
		    end select
                %>
                                        <option value="0"  <%=rSel0%>>0%</option>
                                        <option value="0.05"  <%=rSel5%>>5%</option>
                                        <option value="0.07"  <%=rSel7%>>7%</option>
                                        <option value="0.10" <%=rSel10%>>10%</option>
                                        <option value="0.15" <%=rSel15%>>15%</option>
                                        <option value="0.20" <%=rSel20%>>20%</option>
                                        <option value="0.25" <%=rSel25%>>25%</option>
                                        <option value="0.30" <%=rSel30%>>30%</option>
                                        <option value="0.40" <%=rSel40%>>40%</option>
                                        <option value="0.50" <%=rSel50%>>50%</option>
                                        <option value="0.60" <%=rSel60%>>60%</option>
                                        <option value="0.75" <%=rSel75%>>75%</option>
                                        </select>
        <td>
            <input id="FM_matikkemoms_<%=m %>" name="FM_matikkemoms_<%=m %>" type="checkbox" value="1" <%=chkIkmoms %> onclick="beregnmatpris('<%=m%>','1')" /></td>                               
        <td>
        <%
        
        if sw = 1 then 'rediger mat
        
        matSubTotal = oRec("matbeloeb")
        
        else
        
        if intMatRabat <> 0 then
        matSubTotal = ((oRec("matantal") * matSalgspris) - (oRec("matantal") * matSalgspris * (intMatRabat/100)))
        else
        matSubTotal = oRec("matantal") * matSalgspris
        end if
        
       
        
        call beregnValuta(matSubTotal,valutaMatKurs,valutaKursFak)
        
        if len(trim(valBelobBeregnet)) <> 0 then
        matSubTotal = valBelobBeregnet
        else
        matSubTotal = 0
        end if
        
        end if
        
        if chkVis = "CHECKED" then
        matSubTotalAll = matSubTotalAll + matSubTotal
        end if
        
        if chkIkmoms = "CHECKED" then
        matSubTotalAlluMoms = matSubTotalAlluMoms + matSubTotal
        end if
        
        'oRec("matantal")
        %>
        <input id="FM_matbeloeb_<%=m %>" name="FM_matbeloeb_<%=m %>" type="hidden" value="<%=SQLBlessDOT(formatnumber(matSubTotal, 2)) %>" style="font-size:9px; width:80px;" />
        <div id="matbelobdiv_<%=m %>" style="position:relative; left:0px; top:1px; width:50px;" align="right"><%=SQLBlessDOT(formatnumber(matSubTotal, 2)) %></div>
        </td>
        <td align=center>
        <div id="matvalutadiv_<%=m %>" style="position:relative; left:0px; top:1px;"><%=valutakodeSEL%></div>
        </td>
       
    </tr>
    <%
    
    
    
    end function
    
    
    function godkendknap()
    %>
  
     <div id="knap_godkend" style="position:absolute; visibility:visible; display:; top:46px; width:125px; left:580px; border:1px #6CAE1C solid; padding:3px 5px 10px 5px; background-color:#DCF5BD;">
         <table cellspacing=0 cellpadding=0 border=0 width=100%><tr>
                        <td class=lille><%=erp_txt_463 %>:<br /> 
                            <input id="Submit1" type="submit" value="<%=erp_txt_464 %>" />
                        </td>
                       </tr></table>
	</div>
    <%
    end function
    
    
    
   sub matFelter
   %>
   <td>
    &nbsp;</td>
    <td class=lille><%=erp_txt_465 %></td>
    <td class=lille><%=erp_txt_466 %></td>
    <td class=lille><%=erp_txt_467 %></td>
    <td class=lille><%=erp_txt_468 %></td>
    <td class=lille><%=erp_txt_469 %></td>
    <td class=lille><%=erp_txt_470 %></td>
    <td class=lille><%=erp_txt_471 %></td>
    <td class=lille><%=erp_txt_472 %></td>
    <td class=lille><%=erp_txt_473 %></td>
    <td class=lille><%=erp_txt_474 %></td>
    <td class=lille align=right><%=erp_txt_475 %>&nbsp;</td>
    <td class=lille><%=erp_txt_476 %></td>
   <%
   end sub
   
   
   
   
   sub enhval_gbl
   %>
    

    <div style="position:absolute; left:725px; top:340px; padding:0px; width:210px; border:1px #CCCCCC solid;">
    <table width=100% border=0 cellspacing=0 cellpadding=0 bgcolor="#FFFFFF">
	<tr>
	    <td colspan=2 style="padding:10px 5px 2px 10px;">
            <h4><%=erp_txt_477 %></h4>
	        <b><%=erp_txt_478 %></b><br />
	        <%call selectAllValuta(2, jftp) %>
	   </td>
	</tr>
	<tr>
		<td colspan=2 style="padding:8px 5px 10px 10px;"> 
		<b><%=erp_txt_479 %></b> <span style="color:#999999; font-size:9px;">(<%=erp_txt_480 %>)</span><br />
		
                                <%
                               
		    select case intRabat
		    case 0
		    rSel0 = "SELECTED"
            case 5
            rSel5 = "SELECTED"
            case 7
            rSel7 = "SELECTED"
            case 10
            rSel10 = "SELECTED"
            case 15
            rSel15 = "SELECTED"
            case 20
            rSel = "SELECTED"
            case 25
            rSel25 = "SELECTED"
            case 30
            rSel30 = "SELECTED"
            case 40
            rSel40 = "SELECTED"
            case 50
            rSel50 = "SELECTED"
            case 60
            rSel60 = "SELECTED"
            case 75
            rSel75 = "SELECTED"
		    end select
                %>
                                        
                                       
		                       
                                    <select id="FM_rabat_all" name="FM_rabat_all" style="width:50px; font-size:9px;" onChange="opd_rabatrall()">
                                    <!-- opd_rabatrall(), setBeloebTotAll()-->
                                    
                                       <option value="0"  <%=rSel0%>>0%</option>
                                       <option value="0.05" <%=rSel5%>>5%</option>
                                       <option value="0.07" <%=rSel7%>>7%</option>
                                        <option value="0.10" <%=rSel10%>>10%</option>
                                        <option value="0.15" <%=rSel15%>>15%</option>
                                        <option value="0.20" <%=rSel20%>>20%</option>
                                        <option value="0.25" <%=rSel25%>>25%</option>
                                        <option value="0.30" <%=rSel30%>>30%</option>
                                        <option value="0.40" <%=rSel40%>>40%</option>
                                        <option value="0.50" <%=rSel50%>>50%</option>
                                        <option value="0.60" <%=rSel60%>>60%</option>
                                        <option value="0.75" <%=rSel75%>>75%</option>
                                        </select> 
                                   
                                   &nbsp;<input id="Button1" type="button" value=" Opdater >> " style="font-size:9px;" onClick="opd_rabatrall()" />
	        
                                   <!-- &nbsp;<input id="setRabtAll" type="button" value="Opdater (2-5 sek.)" style="font-size:9px;" onClick="opd_rabatrall(), setBeloebTotAll()" />-->
								
								
								
	
	
   
    
    <%
    '** Sætter default værdier til enheder **'
    if func <> "red" AND (lto = "dencker") then 
    intEnhedsang = 2
    end if

    if func <> "red" AND (lto = "jttek") then 
    intEnhedsang = -1
    end if
    
    select case intEnhedsang
    case -1
    chke_1 = "CHECKED"
	chke3 = ""
	chke2 = ""
	chke1 = ""
	case 1
	chke_1 = ""
	chke3 = ""
	chke2 = "CHECKED"
	chke1 = ""
	case 2
	chke3 = "CHECKED"
	chke2 = ""
	chke1 = ""
	chke_1 = ""
	case 3
	chke4 = "CHECKED"
	chke3 = ""
	chke2 = ""
	chke1 = ""
	chke_1 = ""
	case else
	chke_1 = ""
	chke3 = ""
	chke2 = ""
	chke1 = "CHECKED"
	end select%>
          
		<br /><br /><b><%=erp_txt_483 %></b> <span style="color:#999999; font-size:9px;">(<%=erp_txt_507 %>)</span><br />  
		<input type="radio" name="FM_enheds_ang" id="FM_enheds_ang0" value="0" <%=chke1%> onclick="opd_akt_endhed('Pr. time','0')"> <%=erp_txt_481 %><br />
		<input type="radio" name="FM_enheds_ang" id="FM_enheds_ang1" value="1" <%=chke2%> onclick="opd_akt_endhed('Pr. stk.','1')"> <%=erp_txt_482 %><br />
		<input type="radio" name="FM_enheds_ang" id="FM_enheds_ang2" value="2" <%=chke3%> onclick="opd_akt_endhed('Pr. enhed','2')"> <%=erp_txt_483 %><br />
		<input type="radio" name="FM_enheds_ang" id="FM_enheds_ang3" value="3" <%=chke4%> onclick="opd_akt_endhed('Pr. km.','3')"> <%=erp_txt_484 %><br />
		<input type="radio" name="FM_enheds_ang" id="FM_enheds_ang-1" value="-1" <%=chke_1%> onclick="opd_akt_endhed('Ingen','-1')"> <%=erp_txt_485 %><br />
		
		<!--
		<br /><b>Vis antal som Timer ell. Faktor (Enheder)</b>:<br />
		Timer <input id="FM_timer_faktor0" name="FM_timer_faktor" type="radio" value=0 /> 
        Faktor / Enheder<input id="FM_timer_faktor1" name="FM_timer_faktor" type="radio" value=1 />
        <input id="Submit1" type="submit" value="Opdater" style="font-size:9px;" onClick="opd_timer_faktor_all(), setBeloebTotAll()" />
		-->
	
	    
       
	
	</td>
	</tr>
	</table>
	</div>

   <%
   end sub
   
   
  




			
%>