
<%'GIT 20160811 - SK

'Response.write lto




if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	'*** S�tter lokal dato/kr format. Skal inds�ttes efter kalender.
	Session.LCID = 1030
	
	
	
	'if func = "dbopr" then
	'id = 0
	'else
	'id = request("id")
	'end if
	
	
	if len(request("FM_type")) <> 0 then
	intType = request("FM_type")
	else
	intType = 0
	end if
	
	
	
    thisfile = "erp_fak.asp"
    print = request("print")
    err = 0
    
    'kid = request("FM_kunde")
   
    if len(trim(request("FM_job"))) <> 0 then
    jobid = request("FM_job")
    else
    jobid = 0
    end if
    
    if len(trim(request("FM_aftale"))) <> 0 then
    aftid = request("FM_aftale")
    else
    aftid = 0
    end if
    
    'Response.Write aftid
    'Response.end

    'jobelaft = request("jobelaft")
	
	func = request("func")
	
    
    'if len(trim(request("rdir"))) <> 0 then
	'rdir = request("rdir")
	'else
	'rdir = ""
	'end if
	
	
	select case func
	case "xslet"
	

	
	
	
	
	case "xsletok"
	
	
	
	
	
	
	
	case "xfortryd"
	
	
	
	
	
	
	
	case "dbred", "dbopr" 
	
				
	function SQLBlessDOT(s)
	dim tmpdot
	tmpdot = s
	tmpdot = replace(tmpdot, ".", "")
	SQLBlessDOT = tmpdot
	end function
	
	function SQLBless(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, ",", ".")
	SQLBless = tmp
	end function
	
	function SQLBlessPunkt(s)
	dim tmpPunkt
	tmpPunkt = s
	tmpPunkt = replace(tmpPunkt, ".", ",")
	SQLBlessPunkt = tmpPunkt
	end function
	
	function SQLBlessK(s)
    dim tmp
    tmp = s
    tmp = replace(tmp, "'", "''")
    SQLBlessK = tmp
    end function

    '************************
    '*** findes allerede ****
    '************************
    function SQLBless2(s)
    dim tmp2
    tmp2 = s
    tmp2 = replace(tmp2, ",", ".")
    SQLBless2 = tmp2
    end function

    function SQLBless3(s)
    dim tmp3
    tmp3 = s
    tmp3 = replace(tmp3, ".", ",")
    SQLBless3 = tmp3
    end function
	

    intBeloeb = replace(trim(Request("FM_Beloeb")), ",", ".")
	'intBeloeb = replace(intBeloeb, ",", ".")
	
    if len(trim(request("FM_totbel_afvige"))) <> 0 then
    totbel_afvige = request("FM_totbel_afvige")
    else
    totbel_afvige = 0
    end if
  

	'** Bruges denne ?? ***'
	if len(trim(Request("FM_Beloeb_umoms"))) <> 0 then
    intBeloebUmoms = replace(trim(Request("FM_Beloeb_umoms")), ",", ".")
	'intBeloebUmoms = replace(intBeloebUmoms, ",", ".")
	else
	intBeloebUmoms = 0
	end if
	
	intTimer = SQLBless(Request("FM_Timer"))
	
	
    'Response.Write "intTimer " & intTimer
    'Response.end

	'if aftid <> 0 then
	'intRabat = Request("FM_rabat")
	'else
	intRabat = 0
	'end if
	
	
	valuta = request("FM_valuta_all_1")
	kurs = replace(request("valutakurs_"& valuta &""), ",", ".")
	
	
    if request("FM_sprog") <> 1 then
	intSprog = request("FM_sprog")
	else
	intSprog = 1
	end if
	

            '**** Intern faktura == Ingen moms ***'

            if len(trim(request("FM_medregnikkeioms"))) <> 0 then
            medregnikkeioms = request("FM_medregnikkeioms")
            else
            medregnikkeioms = 0
            end if

            'Response.Write "medregnikkeioms " & medregnikkeioms
            'Response.end

           '** opr v�rdi intern ***'
           medregnikkeioms_opr = request("FM_medregnikkeioms_opr")
           
									
	        '**** Godkendt / ikke gokendt ****' 
            '**   1 = godkendt, 0 = kladde  **'
            intFakbetalt = request("FM_betalt")
			if intFakbetalt <> 1 then
			intFakbetalt = 0
			else
			intFakbetalt = intFakbetalt 
			end if

            
	      

            
	
	
	
	useleftdiv = "m"
	
	'*** Her tjekkes om alle required felter er udfyldt. ***
	if len(Request("FM_timer")) = 0 OR len(Request("FM_beloeb")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->

	<%
	'len(Request("FM_faknr")) = 0 OR
	errortype = 26
	call showError(errortype)
	
	else%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(trim(intBeloeb))
			if isInt > 0 then
			%>
			<!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
			
			<%
			errortype = 27
			call showError(errortype)
			
			isInt = 0
					else
					
					call erDetInt(trim(intTimer))
					if isInt > 0 then
					%>
					<!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
					
					<%
					errortype = 28
					call showError(errortype)
					isInt = 0

                            else 'm� bel�b afvige faktura linjer
                            


                            if cint(totbel_afvige) = 0 then        

                            if len(trim(request("FM_timer_beloeb"))) <> 0 then
                            timer_beloebTjk = request("FM_timer_beloeb")
                            else
                            timer_beloebTjk = 0
                            end if

                            if len(trim(request("FM_materialer_beloeb"))) <> 0 then
                            mat_beloebTjk = request("FM_materialer_beloeb")
                            else
                            mat_beloebTjk = 0
                            end if

                            if len(trim(request("FM_beloeb"))) <> 0 then
                            intBeloebtjk = request("FM_beloeb")
                            else
                            intBeloebtjk = 0
                            end if

                            tjkSum = formatnumber((timer_beloebTjk/1 + mat_beloebTjk/1), 2)

                            if cdbl(formatnumber(intBeloebtjk,2)) <> cdbl(formatnumber(tjkSum, 2)) then

                            %>
					        <!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
					
					        <%
					        errortype = 169
					        call showError(errortype)

                            'Response.Write "intBeloeb: "& intBeloeb & "<br>"
                            'Response.Write "FM_timer_beloeb:" & request("FM_timer_beloeb") & "<br>"
                       
                            'Response.Write "FM_materialer_beloeb" & request("FM_materialer_beloeb") & "<br>"

                            response.end

                                 end if


                            end if

							
							
                            '******************'
	                        '*** Faktura nr ***'
	                        '******************'
                             afsender = request("FM_afsender")

                            call findFaknr(func)
                           
								
								
						     if cint(intFaknumFindes) = 1 then
						     
						     %>
					        <!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
					        
					        <%
					        errortype = 89
					        call showError(errortype)
					        
							
							else
							
							
							'*** Hvis alle required er udfyldt ***'
						            
									'**************************************'
						            '** Faktura dato, Periode og Label ****'
						            '**************************************'
						            
						            
							        istDato = request("FM_istdato_aar") & "/" & request("FM_istdato_mrd") & "/" &request("FM_istdato_dag") 
									showistDato = request("FM_istdato_dag") & "/" & request("FM_istdato_mrd") & "/" &request("FM_istdato_aar") 
									
									istDato2 = request("FM_istdato2_aar") & "/" & request("FM_istdato2_mrd") & "/" &request("FM_istdato2_dag") 
									showistDato2 = request("FM_istdato2_dag") & "/" & request("FM_istdato2_mrd") & "/" &request("FM_istdato2_aar") 
									
									labelDato = request("FM_labelDato_aar") & "/" & request("FM_labelDato_mrd") & "/" &request("FM_labelDato_dag") 
									
									
									fakDato = Request("FM_fakdato_aar") & "/" & Request("FM_fakdato_mrd") & "/" & Request("FM_fakdato_dag") '& time 
									
									if len(request("FM_brugfakdatolabel")) <> 0 then 
							        brugfakdatolabel = 1
							        else
							        brugfakdatolabel = 0
							        end if
									
									'*** Brug label eller faktura system dato **'
									if cint(brugfakdatolabel) = 1 then
									showfakDato = labelDato 'showistDato2
									else
									showfakDato = Request("FM_fakdato_dag") & "/" & Request("FM_fakdato_mrd") & "/" & Request("FM_fakdato_aar")
									end if
		                            
									betbetint = Request.Form("FM_betbetint")
		                            
									
									
									
									'Response.Cookies("erp")("forfaldsdatoDage") = adddayVal
									
									'*** Sidste rettidige bet. dato *'
									if betbetint <> 0 then
                                   

                                            if isDate(replace(Request("FM_forfaldsdato"), ".", "-")) = true then
                                            tempForfalddate = replace(Request("FM_forfaldsdato"), "-", ".")
                                            tempForfalddate = trim(tempForfalddate)
									        reformatdate = split(tempForfalddate, ".")
                                            forfaldsdato = reformatdate(0) & "/" & reformatdate(1) & "/" & reformatdate(2)
                                            else
                                                %>
					                            <!--#include file="../inc/regular/header_lysblaa_nojava_inc.asp"-->
					        
					                            <%
					                            errortype = 152
					                            call showError(errortype)
                                                response.end

                                            end if


									else
									forfaldsdato = DateAdd("d", -1, showfakDato)
									end if
									dtb_dato = year(forfaldsdato) & "/"& month(forfaldsdato) & "/" & day(forfaldsdato)
									
									strEditor = session("user")
									strDato = session("dato")
									
									intfakadr = request("FM_Kid")
									
									kundeid = kid 'request("FM_kundeid")
									
									if request("FM_att_vis") = "on" then
									strAtt = request("FM_att")
									else
									strAtt = 0
									end if
									
                                   
									
									if len(request("FM_varenr")) <> 0 then
									strVarenr = request("FM_varenr")
									else
									strVarenr = 0
									end if
									
									intEnhedsang = request("FM_enheds_ang")
                                    'Response.Write intEnhedsang
                                    'Response.end
																		
									jobfaktype = request("jobfaktype")
									

                                    

									
									'************************************'
									'***** Konti, moms og posteringer ***'
									'************************************'
									
						
									'*** Debit og Kredit konto **'
									if len(request("FM_kundekonto")) <> 0 then
									varKonto = request("FM_kundekonto")
									else
									varKonto = 0
									end if
									
									'Response.Write "varKonto "& varKonto
									'Response.end
									
									Response.Cookies("erp")("debkonto") = varKonto
									
									if len(request("FM_modkonto")) <> 0 then
									varModkonto = request("FM_modkonto")
									else
									varModkonto = 0
									end if
									Response.Cookies("erp")("krekonto") = varModKonto
									
									'** Momskonto **'
									'*** Hvilken konto bruges til at beregne moms p� fakturaen? ***'
									'if len(request("FM_momskonto")) <> 0 then
	                                'momskonto = request("FM_momskonto")
	                                'else
	                                'momskonto = "1"
	                                'end if
                                	
	                                'Response.Cookies("erp")("momskonto") = momskonto
	                                
									
									'*** Data til postering ***'
						            if aftid <> 0 then
						            posteringsTxt = left(request("FM_jobbesk"), 20)
						            else
						            posteringsTxt = left(request("FM_jobnavn"), 20)
						            end if
							
							      
							        strTekst = posteringsTxt
							        vorrefId = session("Mid")
							        intStatus = 1
							        
							       
							        intkontonr = varKonto
									modkontonr = varModKonto
									
									momssats = request("FM_momssats")
									
									momskontoUse = 0 'intkontonr
									momskonto = momskontoUse
									
									'*** Til momsberegning ***'
									'*** d = Beregner moms oven i et netto bel�b.
									'*** k = Tr�kker moms fra et brutto bel�b.
									'****'
									
									debkre = "fak"
								    
								    varBilag = intFaknum 
									posteringsdato = fakDato
									strKomm = SQLBlessK(Request("FM_Komm"))
									antalAkt = Request("antal_x")
									
									'*** Bel�b der skal regnes moms af ****'
                                    if len(trim(intBeloeb)) <> 0 AND intBeloeb <> "NaN" then
                                    intBeloeb = intBeloeb
                                    else
                                    intBeloeb = 0
                                    end if

                                    if len(trim(intBeloebUmoms)) <> 0 AND intBeloebUmoms <> "NaN" then
                                    intBeloebUmoms = intBeloebUmoms
                                    else
                                    intBeloebUmoms = 0
                                    end if


                                    intTotalMoms = replace(intBeloeb, ".", ",") - (replace(intBeloebUmoms, ".", ","))
									

								    '*** Momsberegning ****'
									if cint(medregnikkeioms) = 1 OR cint(medregnikkeioms) = 2 then '1: Intern / 2: Handeslfaktura (Proforma)
                                    momsBelob = 0
                                    else
                                    momsBelob = (intTotalMoms * momssats/100)
									end if
									
									'** Momsberegning til faktura **'
									'call beregnmoms(debkre, intTotalMoms, momskontoUse)
								    
								    
								   
								    'intTotal = replace(intBelob, ".", ",")
								    intNetto = replace(intTotalMoms, ".", ",") 
	                                intMoms = momsBelob
                                            
                                            if intMoms < 0 then 'NT
                                                    intMoms = intMoms * -1
                                            end if

	                                intTotal = replace(intBeloeb, ".", ",")

                                    
	                                intTotal = intTotal + (intMoms)
                                   
	                                'Response.Write "intMoms" & intMoms  & "<br>"
	                                'Response.Write "intTotalMoms" & intTotalMoms & "<br>"
	                                'Response.Write "intBeloeb" & intBeloeb & "<br>"
	                                'Response.Write  "intTotal" & intTotal & "<br>"
		                            
		                            'Response.end
		                            
		                            showmoms = replace(intMoms, ".", "")
								    subtotaltilmoms = replace(intTotalMoms, ",", ".")
								    
								  
								   
						                    '*** ******************** ***'
							                '*** Opretter posteringer ***'
							                '*** ******************** ***'
                							
							                if intFakbetalt <> 0 then 
                						    
						                    '**** Beregner Posteringer hvis faktura er godkendt *****'
                						    '** Valuta / Kurs omregner til DKK **'
                						    intNetto = (intNetto * (kurs/1)/100)
                						    intMoms = (intMoms * (kurs/1)/100)
                						    intTotal = (intTotal * (kurs/1)/100)
                						    
                						    'Response.Write "intTotal: " & intTotal
                                            'Response.end
                			
							                '**** Postering debit konto ****'
							                '**** �ndre komme til punktum til postering **'
							                if intType = 0 then '(faktura)
							                intNettoDeb = replace(replace(formatnumber(intNetto, 2), ".", ""), ",", ".")
		                                    intMomsDeb = replace(replace(formatnumber(intMoms, 2), ".", ""), ",", ".")
		                                    intTotalDeb = replace(replace(formatnumber(intTotal, 2), ".", ""), ",", ".")
                		                    else
                		                    intNettoDeb = replace(replace(formatnumber(-intNetto, 2), ".", ""), ",", ".")
		                                    intMomsDeb = replace(replace(formatnumber(-intMoms, 2), ".", ""), ",", ".")
		                                    intTotalDeb = replace(replace(formatnumber(-intTotal, 2), ".", ""), ",", ".")
                		                    end if 
							                
							                
							                
							                'Response.Write "intTotalDeb" & intTotalDeb & "<br>"
							                'Response.end
							                
                						    oprid = 0
                						    if intkontonr <> 0 then
		                                    
		                                        call opretPosteringSingle(oprid, "2", "dbopr", intkontonr, modkontonr, intTotalDeb, intTotalDeb, 0, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorrefId)
    		
                    						    
                    						    
                    						    
                    						    
                                    		    '**** Postering kredit konto ****'
                                    		    if intType = 0 then '(faktura)
                                    		    'intNettoKre = replace(replace(formatnumber(-intNetto, 2), ".", ""), ",", ".")
                                                intNettoKre = replace(replace(formatnumber(-(intTotal/1 - intMoms/1)), ".", ""), ",", ".")
		                                        intMomsKre = replace(replace(formatnumber(-intMoms, 2), ".", ""), ",", ".")
		                                        intTotalKre = replace(replace(formatnumber(-intTotal, 2), ".", ""), ",", ".")
                                                
                		                        else
                		                        'intNettoKre = replace(replace(formatnumber(intNetto, 2), ".", ""), ",", ".")
                                                intNettoKre = replace(replace(formatnumber((intTotal/1 - intMoms/1)), ".", ""), ",", ".")
		                                        intMomsKre = replace(replace(formatnumber(intMoms, 2), ".", ""), ",", ".")
		                                        intTotalKre = replace(replace(formatnumber(intTotal, 2), ".", ""), ",", ".")
                                                
                		                        end if
                    		                    

                                                'Response.Write " intNetto. "& intNetto &" intNettoKre: "&  intNettoKre & " intTotal: " & intTotal
                                                'Response.end
                    		                    
	                                            call opretPosteringSingle(oprid, "2", "dbopr", modkontonr, intkontonr, intNettoKre, intNettoKre, intMomsKre, strEditor, varBilag, strTekst, posteringsdato, intStatus, vorrefId)
                                        		
                                        		'Response.end
                                        		
                                    		end if '*** Kontonr <> 0
                						    
                						    
						                    end if
								            
								            
								            intMoms = replace(replace(formatnumber(intMoms, 2), ".", ""), ",", ".")
								            
                                           

									'***** Slut konti, moms og posteringer ***'
									'*****************************************'
									
									'Response.end
									
									
							'******* Modtager *****'
							if request("FM_land_vis") = "on" then
							intModland = 1
							else
							intModland = 0
							end if
							
							if len(trim(request("FM_modtageradr"))) <> 0 then
							modtageradr = replace(request("FM_modtageradr"), "'","''")
							else
							modtageradr = ""
							end if
							
							if request("FM_att_vis") = "on" then
							intModAtt = 1
							else
							intModAtt = 0
							end if
							
							if len(trim(request("FM_usealtadr"))) <> 0 then
							usealtadr = request("FM_altadr")
							else
							usealtadr = 0
							end if

                            if len(trim(usealtadr)) <> 0 then
                            usealtadr = usealtadr
                            else
                            usealtadr = 0
                            end if
							
							'if request("FM_tlf_vis") = "on" then
							'intModTlf = 1
							'else
							intModTlf = 0
							'end if
							
							if request("FM_cvr_vis") = "on" then
							intModCvr = 1
							else
							intModCvr = 0
							end if
							
							
							'***** Afsender ****'
							if request("FM_yswift_vis") = "on" then
							intAfsSwift = 1
							else
							intAfsSwift = 0
							end if
							
							
							if request("FM_yiban_vis") = "on" then
							intAfsIban = 1
							else
							intAfsIban = 0
							end if
							
							if request("FM_ycvr_vis") = "on" then
							intAfsICVR = 1
							else
							intAfsICVR = 0
							end if
								
							if request("FM_yemail_vis") = "on" then
							intAfsEmail = 1
							else
							intAfsEmail = 0
							end if
							
							if request("FM_ytlf_vis") = "on" then
							intAfsTlf = 1
							else
							intAfsTlf = 0
							end if	
							
							if request("FM_yfax_vis") = "on" then
							intAfsFax = 1
							else
							intAfsFax = 0
							end if		
							
							if len(trim(request("FM_sideskiftlinier"))) <> 0 then
							sideskiftlinier = request("FM_sideskiftlinier")
							else
							sideskiftlinier = 0
							end if
							
							vorref = request("FM_vorref")

                           
							
							'*** Vis kun totalbel�b ***'
							if len(request("FM_hidesumaktlinier")) <> 0 then
							hidesumaktlinier = request("FM_hidesumaktlinier")
							else
							hidesumaktlinier = 0
							end if
							
							
							'*** Jobbesk ****'
							'if len(request("FM_visjobbesk")) <> 0 OR aftid <> 0 then
							jobBesk = SQLBlessK(request("FM_jobbesk"))
						    'else
							'jobBesk = ""
							'end if
							
							'** vis IKKE jobnavn **'
							if len(trim(request("FM_visikkejobnavn"))) <> 0 then
							visikkejobnavn = 1
							else
							visikkejobnavn = 0
							end if
							
							
							'*** Cookie p� husk jobbesk ***'
							if request("FM_visjobbesk") = "1" then 'AND func = "dbopr" then 
							'response.Cookies("erp")("huskjobbesk") = "1"
                            visjobbesk = 1
							else
							'response.Cookies("erp")("huskjobbesk") = "0"
                            visjobbesk = 0
							end if
							
							
							'*** Timer subtotal bel�b ****
							if len(request("FM_timer_beloeb")) <> 0 then
							timersubtotal = SQLBless2(request("FM_timer_beloeb"))
							else
							timersubtotal = 0
							end if
							
							
							if len(request("FM_visjoblog")) <> 0 then
							visjoblog = 1
							else
							visjoblog = 0
							end if
							
							
							if len(request("FM_vis_joblog_timepris")) <> 0 then
							visjoblog_timepris = 1
							else
							visjoblog_timepris = 0
							end if
							
							if len(request("FM_vis_joblog_enheder")) <> 0 then
							visjoblog_enheder = 1
							else
							visjoblog_enheder = 0
							end if
							
                            'response.write "request(FM_vis_joblog_mnavn): " & request("FM_vis_joblog_mnavn")
                            'response.end

							if len(request("FM_vis_joblog_mnavn")) <> 0 then
							visjoblog_mnavn = 1
							else
							visjoblog_mnavn = 0
							end if
							
							
							
							if len(request("FM_vismatlog")) <> 0 then
							vismatlog = 1
							else
							vismatlog = 0
							end if
							
							
							
							if len(request("FM_visrabatkol")) <> 0 then
							visrabatkol = 1
							else
							visrabatkol = 0
							end if

                            if len(request("FM_hideantenh")) <> 0 then
							hideantenh = 1
							else
							hideantenh = 0
							end if

                          
							
							if len(request("FM_visperiode")) <> 0 then 
							visperiode = 1
							else
							visperiode = 0
							end if
							
							
							if len(trim(request("FM_fak_ski"))) <> 0 then
							fak_ski = 1
							else
							fak_ski = 0
							end if
							
							if len(trim(request("FM_fak_abo"))) <> 0 then
							fak_abo = 1
							else
							fak_abo = 0
							end if
							
							if len(trim(request("FM_fak_ubv"))) <> 0 then
							fak_ubv = 1
							else
							fak_ubv = 0
							end if
							
							
							
							Response.Cookies("erp")("visperiode") = visperiode

                            if cint(brugfakdatolabel) = 1 then
							Response.Cookies("erp")("flabel") = "1"
                            else
                            Response.Cookies("erp")("flabel") = ""
                            end if
							
							'Response.Write "request(FM_visrabatkol)" & request("FM_visrabatkol")
							'Response.flush
							'Response.end
							
							
							'*** Materiale filter ****
							if request("FM_mat_viskuneks") <> "" then
	                        response.cookies("erp")("matvke") = request("FM_mat_viskuneks")
	                        else
	                        response.cookies("erp")("matvke") = ""
	                        end if
	                        
	                        
	                        if len(trim(request("FM_mat_ignper"))) <> 0 then
	                        response.cookies("erp")("matignper") = 1
	                        else
	                        response.cookies("erp")("matignper") = ""
	                        end if
	                       
							
							if len(trim(request("FM_showmatasgrp"))) <> 0 then
							showmatasgrp = request("FM_showmatasgrp") '0: i bunden, 1: I bunden som sum pr. grp, 2: Fordelt under hver aktivitet 
							else
							showmatasgrp = 0
							end if

                            if len(trim(request("FM_hidefasesum"))) <> 0 then
                            hidefasesum = 1
                            else
                            hidefasesum = 0
                            end if

                            
                            afs_bankkonto = request("FM_afs_bankkonto")

                            if len(trim(request("FM_fak_fomr"))) <> 0 then
                            fak_fomr = request("FM_fak_fomr")
                            else
                            fak_fomr = 0
                            end if

 									
							'************************************************************************************
							'*** Opdaterer / Redigerer faktura 													*
							'************************************************************************************
							if func = "dbred" then
							                
                                            
                                          

											strSQLupd = "UPDATE fakturaer SET"_
										    &" fakdato = '"& fakDato &"', "_
											&" timer = "& intTimer  &", "_
											&" beloeb = "& intBeloeb &", "_
											&" kommentar = '"& strKomm &"', "_
											&" editor = '"& strEditor &"', "_
											&" dato = '"& strDato &"', "_
											&" tidspunkt = '23:59:59', "_
											&" betalt = "& intFakbetalt &", "_
                                            &" faknr = "& intFaknum &", "_
											&" b_dato = '"& dtb_dato &"', "_
											&" fakadr = "& intfakadr &", "_
											&" att = '"& strAtt &"', "_
											&" konto = "& varKonto &", "_
											&" modkonto = "& varModkonto &", "_
											&" faktype = "& intType &", "_
											&" vismodland = "& intModland &", "_
											&" vismodatt = "& intModAtt &", "_
											&" vismodtlf = "& intModTlf &", "_
											&" vismodcvr = "& intModCvr &", "_
											&" visafstlf = "& intAfsTlf &", "_
											&" visafsemail = "& intAfsEmail &", "_
											&" visafsswift = "& intAfsSwift &", "_
											&" visafsiban = "& intAfsIban &", "_
											&" visafscvr = "& intAfsICVR &", "_
											&" moms = "& replace(showmoms,",",".") &", "_
											&" enhedsang = "& intEnhedsang &", "_
											&" varenr = '"& strVarenr &"', jobbesk = '"& jobBesk &"', "_
											&" timersubtotal = "& timersubtotal &", "_
											&" visjoblog = "& visjoblog &", visrabatkol = "& visrabatkol &", "_
											&" vismatlog = "& vismatlog &", rabat = "& intRabat &", "_
											&" visjoblog_timepris = "& visjoblog_timepris &", "_
											&" visjoblog_enheder = "& visjoblog_enheder &", "_
											&" visjoblog_mnavn = "& visjoblog_mnavn & ", "_
											&" visafsfax = "& intAfsFax &", "_
											&" subtotaltilmoms = "& subtotaltilmoms &", "_
											&" valuta = "& valuta &", kurs = "& kurs &", "_
											&" sprog = "& intSprog &", istdato = '"& istDato &"', "_
											&" momskonto = "& momskonto &", visperiode = "& visperiode &", "_
											&" jobfaktype = "& jobfaktype &", betbetint = "& betbetint &", "_
											&" brugfakdatolabel = "& brugfakdatolabel &", istdato2 = '"& istDato2 &"', "_
											&" momssats = "& momssats &", modtageradr = '"& modtageradr &"', "_
											&" usealtadr = "& usealtadr &", vorref = '"& vorref &"', fak_ski = "& fak_ski &", "_
											&" showmatasgrp = "& showmatasgrp &", "_
											&" hidesumaktlinier = "& hidesumaktlinier &", sideskiftlinier = "& sideskiftlinier &", "_
											&" labeldato = '"& labelDato &"', "_
											&" fak_abo = "& fak_abo &", fak_ubv = "& fak_ubv &", visikkejobnavn = "& visikkejobnavn &", "_
                                            &" hidefasesum = "& hidefasesum &", hideantenh = "& hideantenh &", medregnikkeioms = "& medregnikkeioms & ", "_
                                            &" afsender = " & afsender & ", vis_jobbesk = "& visjobbesk &", "_
                                            &" kontonr_sel = "& afs_bankkonto &", totbel_afvige = "& totbel_afvige &", fak_fomr = "& fak_fomr
                                            
                                            '*** Faktura skal l�ses (markeres at den har v�ret godkendt) Skal kun opdateres ved godkend og skal ikke kunne �ndres
                                            if cint(intFakbetalt) = 1 then
                                            strSQLupd = strSQLupd & ", fak_laast = 1" 
                                            end if

											strSQLupd = strSQLupd  &" WHERE Fid = "& id 
											
											
											'Response.write (strSQLupd)
											'Response.end
											oConn.execute(strSQLupd)
											
											
												    '*** Shadowcopy **
													'*** Hvis der oprettes en faktura p� en aftale, hvor der er tilknyttet job, skal der oprettes en shadowcopy 
													'*** P� de job der er tilknyttet s� de kan lukkes for indtastning, samt at man kan se at der ligger en GHOSTET
													'*** Faktura p� jobbet 

                                                    '** NT ***'
                                                    jobids_orders = request("jobids_orders")
													closealsoJob = ""
													if aftid <> 0 OR jobids_orders <> "0" then
													
													'if aftid <> 0 then
													
													    
													    strSQLshadow = "SELECT faknr FROM fakturaer WHERE fid = "& id
													    oRec.open strSQLshadow, oConn, 3
                                                        if not oRec.EOF then
                                                        
                                                        intFaknum = oRec("faknr")
                                                        
                                                        end if
                                                        oRec.close
													    
                                                        
                                                        

													    strSQLshadowCopy = "UPDATE fakturaer SET fakdato = '"& fakDato &"', istdato = '"& istDato &"', brugfakdatolabel = "& brugfakdatolabel &", "_
                                                        &" labeldato = '"& labelDato &"', betalt = "& intFakbetalt &" WHERE faknr = "& intFaknum & " AND shadowcopy = 1"
                                                 
													    oConn.execute(strSQLshadowCopy)

                                                        


                                                                '** NT ***'
                                                                select case lto
                                                                case "nt"
													            
													    
													                strSQLshadowCopy = "SELECT jobid FROM fakturaer WHERE shadowcopy = 1 AND faknr = '"& intFaknum &"'"
                                                                    oRec6.open strSQLshadowCopy, oConn, 3
                                                                    while not oRec6.EOF  
													        
                                                               
                                                                    closealsoJob = closealsoJob & " OR id = "& oRec6("jobid")
													                
                                                                    oRec6.movenext
                                                                    wend
                                                                    oRec6.close
													              
                                                                end select
													    
													

                                                                'response.write "closealsoJob: " & closealsoJob
                                                                'response.end

													
													
													end if
											
											
							thisfakid = id
							end if '** RED
							'***
											
											
											
							'************************************************************************************
							'*** Opretter faktura *********														*
							'************************************************************************************
							if func = "dbopr" then
												
												
												'Response.Write "intType " & intType
												'Response.Flush
												
												if intType <> 0 then '*** kreditnota eller rykker
												parentfak = request("id")
												else
												parentfak = 0
												end if
												
                                                '*** Faktura skal l�ses (markeres at den har v�ret godkendt)
                                                if cint(intFakbetalt) = 1 then
                                                fak_laast = 1
                                                else
                                                fak_laast = 0
                                                end if
                                                
                                                	
													'***tidspunkt = 23:23:59 pga luk for indtastning muligheden p� timeregsiden
													
													strSQL = ("INSERT INTO fakturaer"_
													&" (faknr, fakdato, jobid, timer, beloeb, kommentar, dato, editor, tidspunkt, "_
													&" betalt, b_dato, fakadr, att, faktype, konto, modkonto, parentfak, "_
													&" vismodland, vismodatt, vismodtlf, vismodcvr, visafstlf, visafsemail, "_
													&" visafsswift, visafsiban, visafscvr, moms, enhedsang, "_
													&" aftaleid, varenr, jobbesk, timersubtotal, visjoblog,"_
													&" visrabatkol, vismatlog, rabat, visjoblog_timepris, "_
													&" visjoblog_enheder, visafsfax, subtotaltilmoms, "_
													&" valuta, kurs, sprog, istdato, momskonto, visperiode, "_
													&" visjoblog_mnavn, jobfaktype, betbetint, brugfakdatolabel, "_
													&" istdato2, momssats, modtageradr, usealtadr, vorref, fak_ski, "_
													&" showmatasgrp, hidesumaktlinier, sideskiftlinier, labeldato, fak_abo, fak_ubv, "_
                                                    &" visikkejobnavn, hidefasesum, hideantenh, medregnikkeioms, fak_laast, afsender, vis_jobbesk, kontonr_sel, totbel_afvige, fak_fomr) VALUES ("_
													&" '"& intFaknum &"',"_
													&" '"& fakDato &"',"_
													&" "& jobid &","_
													&" "& intTimer &","_
													&" "& intBeloeb &","_
													&" '"& strKomm &"',"_
													&" '"& strDato &"',"_
													&" '"& strEditor &"', '23:59:59', "_
													&" "& intFakbetalt &", '"& dtb_dato &"', "& intfakadr &", "_
													&" '"& strAtt &"', "& intType &", "& varKonto &", "_
													&" "& varModkonto &", "& parentfak &", "_
													&" "& intModland &", "& intModAtt &", "& intModTlf &", "_
													&" "& intModCvr &", "& intAfsTlf &", "& intAfsEmail &", "& intAfsSwift &", "_
													&" "& intAfsIban &", "& intAfsICVR &", "& intMoms &", "& intEnhedsang &", "_
													&" "& aftid &", '"& strVarenr &"', '"& jobBesk &"', "& timersubtotal &", "_
													&" "& visjoblog &", "& visrabatkol &", "& vismatlog &", "& intRabat &", "_
													&" "& visjoblog_timepris &", "& visjoblog_enheder &", "_
													&" "& intAfsFax &", "& subtotaltilmoms &", "& valuta &", "_
													&" "& kurs &", "& intSprog &", '"& istDato &"', "& momskonto &", "_
													&" "& visperiode &", "& visjoblog_mnavn &", "& jobfaktype &", "_
													&" "& betbetint &", "& brugfakdatolabel &", '"& istDato2 &"', "_
													&" "& momssats &", '"& modtageradr &"', "& usealtadr &", '"& vorref &"', "_
													&" "& fak_ski &", "& showmatasgrp &", "& hidesumaktlinier &", "& sideskiftlinier &", "_
													&" '"& labelDato &"', "& fak_abo &", "& fak_ubv &", "& visikkejobnavn &", "_
                                                    &" "& hidefasesum &", "& hideantenh &", "& medregnikkeioms &", "& fak_laast &", "& afsender &", "& visjobbesk &", "& afs_bankkonto &", "& totbel_afvige &", "& fak_fomr &")")
													
													'Response.Write "subtotaltilmoms: " & subtotaltilmoms & " intMoms: "& intMoms &"<br>"
													'if lto = "assurator" then
                                                    'Response.write strSQL & "<br><br>"
													'Response.end
                                                    'end if
													oConn.execute(strSQL)
													
													'Response.end
													'visjoblog_timepris, visjoblog_enheder, visafsfax
										
													
													''** Henter fak id ***
													strSQL3 = "SELECT Fid FROM fakturaer"
													oRec3.open strSQL3, oConn, 3
													oRec3.movelast
													if not oRec3.EOF then
													thisfakid = cint(oRec3("Fid")) 
													end if 
													oRec3.close	
													

                                                    'Response.Write "her"
                                                    'Response.flush

													'*** Shadowcopy ***'
													'*** Hvis der oprettes en faktura p� en aftale, hvor der er tilknyttet job, skal der oprettes en shadowcopy 
													'*** P� de job der er tilknyttet s� de kan lukkes for indtastning, samt at man kan se at der ligger en GHOSTET
													'*** Faktura p� jobbet 
                                                    
                                                    '** NT ***'
                                                    jobids_orders = request("jobids_orders")
													closealsoJob = ""
													if aftid <> 0 OR jobids_orders <> "0" then
													
                                                    if jobids_orders <> "0" then 'nt
                                                    job_tilknyttet_aftale = split(request("jobids_orders"),",")
                                                    else
                                                    job_tilknyttet_aftale = split(request("FM_job_tilknyttet_aftale"),",")
											        end if		

                                                    for i = 0 TO UBOUND(job_tilknyttet_aftale, 1)
													    if (i > 0 AND jobids_orders = "0") OR (i > 1) then
													    
													    strSQLshadowCopy = "INSERT INTO fakturaer (faknr, fakdato, jobid, shadowcopy, istdato, brugfakdatolabel, labeldato, betalt) VALUES "_
													    &"('"& intFaknum &"', '"& fakDato &"', "& job_tilknyttet_aftale(i) &", 1, '"& istDato &"',"& brugfakdatolabel &",'"& labelDato &"', "& intFakbetalt &")"
													    
                                                        select case lto
                                                        case "nt"
                                                        closealsoJob = closealsoJob & " OR id = "& job_tilknyttet_aftale(i)
													    end select
                                                        

													    'Response.Write strSQLshadowCopy
													    'Response.end
													    oConn.execute(strSQLshadowCopy)
													    end if
													next
													    
													
													end if
													
									
									
							end if	'** Opret
									
									
												
						
							'if jobid <> 0 then '** Bruges kun hvis der oprettes faktura p� job ***					
							
							
							'***********************************************************************************
							'*** Hvis faktura allerede er oprrettet en gang i denne session ***
							'***********************************************************************************
								
								 
								'************************************************************** 
								'*** Inds�tter akt i fak_det 
								'**************************************************************
								
                                '*** Nulstiller akt. faktura linier
                                if func = "dbred" then
								oConn.execute("DELETE FROM faktura_det WHERE fakid = "& id &"")
								thisfakid = id
								end if

                                '* Nulstiller ALLE medarb. linier p� denne faktura ***
                                '* Flyttet 8/9-2010 (se nedenfor) ***'
                                if func = "dbred" then
								oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& thisfakid &"") 'AND aktid = "&thisAktId&" AND mid = "&thisMid&"")
                                end if

								
								'***** n starter p� 1 **************
								'***** Aktivitets udspecificering ************

                                '** Eller hvis Konsulent mode = Alle medarbej. linjer bliver altid gemt. '**'
                                'altidIndlasMedarbtimer = 0
                                'select case lto 
                                'case "synergi1", "essens"
                                'altidIndlasMedarbtimer = 1
                                'case else
                                'altidIndlasMedarbtimer = 0
                                'end select 



                               for intcounter = 0 to antalAkt - 1
								
								thisAktId = request("aktId_n_"&intcounter&"")
									
									'** antal aktiviter udskrevet pga. forskellige timepriser 
									'antalsumaktprakt = request("antal_subtotal_akt_"&intcounter&"") 
									antalsumaktprakt = request("antal_n_"&intcounter&"") 
									
                                    'Response.Write antalsumaktprakt & "<br>"
                                    'Response.flush

									for intcounter3 = -(antalsumaktprakt) to antalsumaktprakt
									
												
												'*** Enhedsprisen p� denn akt.
												if len(request("FM_enhedspris_"& intcounter &"_"&intcounter3&"")) <> 0 then
												enhpris = SQLBless2(request("FM_enhedspris_"& intcounter &"_"&intcounter3&""))
												else
												enhpris = 0
												end if
												
												'*** Enheds angivelse p� denn akt.
                                                if len(trim(request("FM_akt_enh_"& intcounter &"_"&intcounter3&""))) <> 0 then
												enhedsang_akt = request("FM_akt_enh_"& intcounter &"_"&intcounter3&"")
                                                else
                                                enhedsang_akt = 0
                                                end if
												
												
									
										'*** Vis sum aktivitet p� print (og DB)
										'Response.write "her " & intcounter &"_"&intcounter3 & " "& request("FM_show_akt_"& intcounter &"_"&intcounter3&"") & "<br>"
                                        'Response.flush
                                        if (len(request("FM_show_akt_"& intcounter &"_"&intcounter3&"")) = 1) then
												
												timerThis = request("FM_timerthis_"& intcounter &"_"&intcounter3&"")
												if len(timerThis) <> 0 then
												timerThis = SQLBless2(timerThis)
												else 
												timerThis = 0
												end if
												
												if len(request("FM_beloebthis_"& intcounter &"_"&intcounter3&"")) <> 0 then
												beloebThis = SQLBless2(request("FM_beloebthis_"& intcounter &"_"&intcounter3&""))
												else
												beloebThis = 0
												end if
												
												'*** Rabat *****
												if len(request("FM_rabat_"&intcounter&"_"&intcounter3&"")) <> 0 then
												rabatThis = SQLBless2(request("FM_rabat_"& intcounter &"_"&intcounter3&""))
												else
												rabatThis = 0
												end if	
												
												'*** Valuta *****'
												if len(request("FM_valuta_"&intcounter&"_"&intcounter3&"")) <> 0 then
												valutaThis = SQLBless2(request("FM_valuta_"& intcounter &"_"&intcounter3&""))
												else
												valutaThis = 1
												end if	
												
												'**** Sortorder ***'
												aktsortorder = request("aktsort_"&intcounter&"")
												call erDetInt(trim(aktsortorder))
			                                    if isInt > 0 then
			                                    aktsortorder = 1
			                                    else
			                                    aktsortorder = replace(aktsortorder, ",", ".")
			                                    end if
												
												kursThis = replace(request("valutakurs_"& valutaThis &""), ",", ".")
												
												momsfri = 0
												if len(trim(request("FM_momsfri_"& intcounter &"_"&intcounter3&""))) <> 0 then
												momsfri = 1
												else
												momsfri = 0
												end if

                                                if len(trim(request("FM_aktfase_"& intcounter &""))) <> 0 then
                                                aktFase = request("FM_aktfase_"& intcounter &"")
                                                call illChar(aktFase)
		                                        aktFase = vTxt
                                                else
                                                aktFase = ""
                                                end if
												
												strSQL_sumakt = ("INSERT INTO faktura_det "_
												&" (antal, beskrivelse, aktpris, fakid, "_
												&" enhedspris, aktid, showonfak, rabat, enhedsang, valuta, kurs, fak_sortorder, momsfri, fase) "_
												&" VALUES ("& timerThis &", "_
												&"'" & SQLBlessK(request("FM_aktbesk_"& intcounter &"_"&intcounter3&"")) &"', "_
												&""  & beloebThis &", "_
												&""  & thisfakid &", "& enhpris &", "& thisAktId &", "_
												&" 1, "& rabatThis &", "& enhedsang_akt &", "& valutaThis &", "& kursThis &", "_
                                                &" "& aktsortorder &", "& momsfri &", '"& aktFase &"')")
												

                                                'if session("mid") =  1 then
												'Response.write strSQL_sumakt  & "<br>"
												'Response.flush
												'end if
                                                oConn.execute(strSQL_sumakt)
												
										
                                        end if 'show sumaktivitet
										
                                        


                                        '** �ndret 8/9-2010 Kun p� faktura linier der er med p� faktura bliver der indl�st timer for medarbejdere til f.eks bonus **'
                                        '** Eller hvis 0 akt. er sl�et til **' == KONSULENT MODE
                                        'OR (cint(altidIndlasMedarbtimer) = 1 AND cdbl(intcounter3) = 0) 
                                     
                                        
                                        '**************************************************'			
										'***** Medarbejder udspecificering i db ***********'
										'**************************************************'

                                        if jobid <> 0 then '** kun job

                                        if (len(request("FM_show_akt_"& intcounter &"_"&intcounter3&"")) = 1 OR len(request("FM_show_akt_"& intcounter &"_0")) = 1) then			
										
										
										'Response.write "5<br>"
										'Response.flush
											
										antalmedspec2 = request("antal_n_"&intcounter&"") 'medarbejdere
										for intcounter2 = 0 to antalmedspec2 - 1
												
												
                                                '** registrerer timer (fak og vent) p� alle medarbejdere 
												if len(request("FM_mid_"&intcounter2&"_"&intcounter&"")) <> 0 then
												thisMid = request("FM_mid_"&intcounter2&"_"&intcounter&"")
												else
												thisMid = 0
												end if 

                                                
                                               
                                                
                                                
                                                '*** Passer timeprisen p� denne akt og medarbejder **
												thismedarbtpris = request("FM_mtimepris_"& intcounter2&"_"&intcounter&"")
												thismedarbtpris = formatnumber(thismedarbtpris, 2)
												thismedarbtpris = replace(thismedarbtpris, ".", "")
												thismedarbtpris = SQLBless2(thismedarbtpris) 
												
												thisenhpris = request("FM_enhedspris_"& intcounter &"_"&intcounter3&"")
												thisenhpris = formatnumber(thisenhpris, 2)
												thisenhpris = replace(thisenhpris, ".", "")
												enhpris = SQLBless2(thisenhpris)
												
												
												
												
												
												'Response.Write "<hr>thisMid " & thisMid & "<br>" 
												
												'if thisMid <> 0 then
												'Response.write "" & request("FM_aktbesk_"& intcounter &"_"&intcounter3&"") & "<br>"
												'Response.write "if" & thismedarbtpris &" = "& enhpris  &" then<br>"
												'Response.write "timer:" & request("FM_m_fak_"&intcounter2&"_"&intcounter&"")
												'Response.write "<br>true sum aktr:" & request("FM_show_akt_"&intcounter&"_"&intcounter3&"")
												'end if
												
                                              

												if cdbl(thismedarbtpris) = cdbl(enhpris) AND thisMid <> 0 then
												'Response.write "<br>ok!<br>"
											
													    
                                                        '*** Denne er flyttet s� alle tidligere reg. bliver slettet inde de nye akt. bliver indl�st ****'
														'* Nulstiller evt. tidligere indtastninger p� denne faktura ***
														'oConn.execute("DELETE FROM fak_med_spec WHERE fakid = "& thisfakid &" AND aktid = "&thisAktId&" AND mid = "&thisMid&"")
													 	
														'* Bel�b
														if len(request("FM_mbelob_"& intcounter2 &"_"&intcounter&"")) <> 0 then
														useBeloeb = SQLBless2(request("FM_mbelob_"& intcounter2 &"_"&intcounter&""))
														else
														useBeloeb = 0
														end if
														
														'* Venter
														if len(request("FM_m_vent_"&intcounter2&"_"&intcounter&"")) <> 0 then
														useVenter = SQLBless2(request("FM_m_vent_"&intcounter2&"_"&intcounter&""))
														else
														useVenter = 0
														end if
														
														'** Venter brugt 
														if len(request("FM_m_venterbrugt_"&intcounter2&"_"&intcounter&"")) <> 0 then
														useVenterBrugt = SQLBless2(request("FM_m_venterbrugt_"&intcounter2&"_"&intcounter&""))
														else
														useVenterBrugt = 0
														end if
														
														
														
														'*** Rabat *****'
														if len(request("FM_mrabat_"&intcounter2&"_"&intcounter&"")) <> 0 then
														medarb_rabat = SQLBless2(request("FM_mrabat_"& intcounter2 &"_"&intcounter&""))
														else
														medarb_rabat = 0
														end if	
														
														
														'*** Valuta *****'
												        if len(request("FM_mvaluta_"&intcounter2&"_"&intcounter&"")) <> 0 then
												        medarbValuta = SQLBless2(request("FM_mvaluta_"& intcounter2 &"_"&intcounter&""))
												        else
												        medarbValuta = 1
												        end if	
        												
												        medarbKurs = replace(request("valutakurs_"& medarbValuta &""), ",", ".")
														
														
														'*** Enhedsang ***
														enhedsang_med = request("FM_med_enh_"&intcounter2&"_"&intcounter&"")
														
														'** Nulstiler altid vente timer inden der tildeles nye vente timer for denne medarbejder p� denne aktivitet 
														'** (Uanset hvilken fak akt. h�rer til)
														
														'oConn.execute("UPDATE fak_med_spec SET venter = 0 WHERE aktid = "&thisAktId&" AND mid = "&thisMid&"")
													 	
														
														if len(request("FM_m_fak_"&intcounter2&"_"&intcounter&"")) <> 0 then
															'*** Hvis show sum-akt ikke er true skal faktimer altid v�re = 0
															'if request("FM_show_akt_"&intcounter&"_"&intcounter3&"") = "1" then 
															useFak = SQLBless2(request("FM_m_fak_"&intcounter2&"_"&intcounter&""))
															'else
															'useFak = 0
															'end if

                                                            if len(trim(useFak)) <> 0 then
                                                            useFak = useFak
                                                            else
                                                            useFak = 0
                                                            end if

														else
															useFak = 0
														end if
														
														
														if len(request("FM_mtimepris_"& intcounter2&"_"&intcounter&"")) <> 0 then
														usemedTpris = SQLBless2(request("FM_mtimepris_"& intcounter2&"_"&intcounter&""))
														else
														usemedTpris = 0
														end if 
														
														
														if request("FM_show_medspec_"&intcounter2&"_"&intcounter&"") = "show" then
														showonfak = 1
														else
														showonfak = 0
														end if
														
														
														
														'*** Inds�tter i db (uanset om medarb linie vises p� fak, bruges til f.eks bonus beregning mm.) ****
														if useFak <> 0 OR useVenter <> 0 then
														
														strSQL = "INSERT INTO fak_med_spec (fakid, aktid, mid, fak, venter, "_
														&" tekst, enhedspris, beloeb, showonfak, medrabat, venter_brugt, enhedsang, valuta, kurs) "_
														&" VALUES ("& thisfakid &", "&thisAktId&", "_
														&"" &thisMid&", "_
														&"" &useFak&", "_
														&"" &useVenter&", "_
														&"'"&SQLBlessK(request("FM_m_tekst_"& intcounter2&"_"&intcounter&""))&"', "_
														&"" &usemedTpris&", "_
														&"" &useBeloeb&", "& showonfak &", "_
														&"" & medarb_rabat &", "& useVenterBrugt &", "& enhedsang_med &", "& medarbValuta &", "& medarbKurs &")"
														
														
														'Response.write strSQL & "<br><br>"
														'Response.flush
														
														oConn.execute(strSQL)
													    
													    
													    end if
													    
													'else
													'Response.write "Stemmer ikke <br>"	
													
													end if 'thismedarbtpris
													
											next 'Medarbejdere

                                            end if 'show sumaktivitet

                                            end if 'job/aft fak

										next 'Sumaktiviteter
								next 'Antal aktiviteter
							
							
							'Response.write "6<br>"
							'Response.flush
							
							'else	
							
							'select case intFakbetalt
							'case 1
							'erfak = 1 'godkendt
							'case 0
							'erfak = 2 'kladde
							'end select
							'*** Opdaterer seraft, s� aftalen ER faktureret ***
							'oConn.execute("UPDATE serviceaft SET erfaktureret ="& erfak &" WHERE id = "& aftid &"")
									
							'end if '*** Kun job (opr. aktiviterer og fak_med_spec)	
							
							
							
							'Response.end
							
							'***********************************************
							'********* Materialer ************'
                            '***********************************************
							
							
							mat_ialt = request("FM_antal_materialer_ialt")
							
							
							
							if func = "dbred" then
							    '** renser db, opdater materiale forbrug ***'
							    strSQLselmf = "SELECT matfrb_id FROM fak_mat_spec WHERE matfakid = " & thisfakid
							    oRec4.open strSQLselmf, oConn, 3
							    while not oRec4.EOF
    							
							    '*** Markerer i materialeforbrug at materiale er faktureret ***'
							    strSQLmf = "UPDATE materiale_forbrug SET erfak = 0 WHERE id = " & oRec4("matfrb_id")
							    oConn.execute(strSQLmf)
    							
							    oRec4.movenext
							    wend
							    oRec4.close
    							
    							
							    '*** Sletter hidtidige regs ***'
							    strSQLdel = "DELETE FROM fak_mat_spec WHERE matfakid = " & thisfakid
							    oConn.execute strSQLdel
							end if
							
							
							
							for m = 0 to mat_ialt - 1
							
							
							if len(trim(request("FM_vis_"&m&""))) <> 0 then
							matvis = 1
							else
							matvis = 0
							end if
							
							matid = request("FM_matid_"&m&"")
							matnavn = replace(request("FM_matnavn_"&m&""),"'", "''")
							matvarenr = request("FM_matvarenr_"&m&"")
							matantal = SQLBless2(request("FM_matantal_"&m&""))
							matenhed = request("FM_matenhed_"&m&"")
							matenhedspris = SQLBless2(request("FM_matenhedspris_"&m&""))
							matrabat = request("FM_matrabat_"&m&"")
							matbeloeb = SQLBless2(request("FM_matbeloeb_"&m&""))
							matValuta = request("FM_matvaluta_"&m&"")
							matKurs = replace(request("valutakurs_"& matValuta &""), ",", ".")
							matGrp = request("FM_matgrp_"&m&"")

                            matAktid = request("FM_mataktid_"&m&"")           
							
							if len(trim( request("FM_matikkemoms_"&m&""))) <> 0 then
							ikkemoms = request("FM_matikkemoms_"&m&"")
							else
							ikkemoms = 0
							end if
							
							matMFid = request("FM_mfid_"&m&"")
							matMFusrid = request("FM_mfusrid_"&m&"")
							
							if cint(matvis) = 1 then
							
							strSQLoprmat = "INSERT INTO fak_mat_spec (matfakid, matid, matnavn, matvarenr, matantal, matenhed, "_
							&" matenhedspris, matrabat, matbeloeb, matshowonfak, ikkemoms, valuta, kurs, matfrb_mid, matfrb_id, matgrp, fms_aktid) "_
							&" VALUES ("&thisfakid&", "&matid&", '"&matnavn&"', '"& matvarenr&"', "_
							&" "& matantal &", '"& matenhed &"', "& matenhedspris &", "_
							&" "& matrabat &", "& matbeloeb &", 1, "& ikkemoms &", "& matValuta &", "& matKurs &", "& matMFusrid &" ,"& matMFid &", "& matGrp &", "& matAktid &")"
							
							'Response.Write strSQLoprmat & "<br>"
							'Response.flush
							oConn.execute(strSQLoprmat)
							
							'matbeloebTot = matbeloebTot + matbeloeb
							
							'*** Markerer i materialeforbrug at materiale er faktureret ***'
							strSQLmf = "UPDATE materiale_forbrug SET erfak = 1 WHERE id = " & matMFid
							oConn.execute(strSQLmf)
							
							end if
							    
							next
							
							
							
							
			                '**********************************'
			                '************************************'
							
							
					       
    					
							'*** Lukker job ****'
							if len(request("FM_lukjob")) <> 0 then

                            '*** S�t lukkedato (skal v�re f�r det skifter status) ***'
                            call lukkedato(jobid, 0)

							strSQLupd = "UPDATE job SET jobstatus = 0 WHERE id = " & jobid
							oConn.execute(strSQLupd)


                                    '*** NT Luk ogs� tilh�rede ordre fra SHADOW COPY ***
                                    select case lto
                                    case "nt"


                                                'response.write "closealsoJob: " & closealsoJob
                                                if len(trim(closealsoJob)) <> 0 then      
                                                    strSQLupd = "UPDATE job SET jobstatus = 0 WHERE id = 0 "& closealsoJob  ' id = " & jobid
                                                    'response.write strSQLupd
                                                    'response.flush
							                        oConn.execute(strSQLupd)
                                                end if

                                    end select

                                    'response.end

                                    '**** SyncDatoer ****'
                                    jobnrThis = 0
                                    jobStatus = 0
                                    syncslutdato = 0
                                    strSQLj = "SELECT jobnr, syncslutdato FROM job WHERE id =  "& jobid
                                    oRec5.open strSQLj, oConn, 3 
                                    if not oRec5.EOF then

                                    jobnrThis = oRec5("jobnr")
                                    syncslutdato = oRec5("syncslutdato")

                                    end if
                                    oRec5.close

                                    if jobStatus = 0 AND syncslutdato = 1 then
                                        call syncJobSlutDato(jobid, jobnrThis, syncslutdato)
                                    end if



							end if


                                                                    
						
							
							'**** Opdaterer faste betalingsbetingelser p� kunde ***'
							if len(request("FM_gembetbet")) <> 0 then
							strSQLupd = "UPDATE kunder SET betbet = '"& strKomm &"', betbetint = "& betbetint &" WHERE kid = " & kundeid
							oConn.execute(strSQLupd)
							end if
							
							
							'** Gemmer p� alle kunder **'
							if len(request("FM_gembetbetalle")) <> 0 then
							strSQLupd = "UPDATE kunder SET betbet = '"& strKomm &"', betbetint = "& betbetint &" WHERE kid <> 0"
							oConn.execute(strSQLupd)
							end if
							
                            if len(trim(request("FM_rekvnr"))) <> 0 then
                            rekvnr = request("FM_rekvnr")
                            else
                            rekvnr = ""
                            end if

                            strSQLupd = "UPDATE job SET rekvnr = '"& rekvnr &"' WHERE id = " & jobid
							oConn.execute(strSQLupd)
                            'end if
							
							
							
							
							'*** Opretter posteirnger p� Execon / Immenso version ***'
							if intFakbetalt <> 0 then 
							    call ltoPosteringer
							end if
							
							'*** Opdaterer/Inds�tter posterings ID p� faktura ***' 
							'*** M� f�rst kaldes efer alle posteringer ***'
							if intFakbetalt <> 0 then 
							strSQLOprID = "UPDATE fakturaer SET oprid = "& oprid &" WHERE Fid = "& thisfakid &""
							oConn.execute(strSQLOprID)
	                        end if
							

                            '**** Opdaterer faktura nr r�kkef�lge *****'
                            call opdater_fakturanr_rakkefgl(opdFaknrSerie, intFaknumFindes, sqlfld, intFaknum)
							
							Response.Cookies("erp").Expires = date + 60
							
							'Response.End 
							
												
												
											'*** Viser den oprettede faktura til print *****'
											%>
											<!-------------------------------Sideindhold------------------------------------->
											
										
											<%
											stDag = request("FM_start_dag_ival")
											stMrd = request("FM_start_mrd_ival")
											stAar = request("FM_start_aar_ival")
											
											slDag = request("FM_slut_dag_ival")
											slMrd = request("FM_slut_mrd_ival")
											slAar = request("FM_slut_aar_ival")
											

                                                'Re'ponse.write "stDag:" & stDag & " stMrd: "& stMrd & "visjobogaftaler=1&visminihistorik=1&visfaktura=2&FM_usedatokri=1&FM_job="&jobid&"&FM_aftale="&aftid&"&FM_kunde="&kid&"&id="&thisfakid&"&FM_usedatointerval=1&FM_start_dag_ival="&stDag&"&FM_start_mrd_ival="&stMrd&"&FM_start_aar_ival="&stAar&"&FM_slut_dag_ival="&slDag&"&FM_slut_mrd_ival="&slMrd&"&FM_slut_aar_ival="&slAar
											    'Response.end
											
                                               



                                            'if rdir = "tilfak" then
                                                
                                             '   Response.Write("<script language=""JavaScript"">window.opener.document.location.href = 'erp_tilfakturering.asp?rdir=tilfak';</script>")
                                               

                                           
                                                %>
                                                     
                                                  
                                                <!--    <div style="background-color:#FFFFFF; left:120px; top:0px; width:300px; padding:20px;">
                                                    <h4>Genindl�ser "til fakturering" siden..</h4>    
                                                    Genindl�ser "til fakturering" siden, s� den sag eller aftale du netop har faktureret vil blive opdateret med de nyeste data, og "til fakturerings" siden altid er "up to date".<br /><br />
                                                     
                                                        
                                                    <a href="<%="erp_opr_faktura_fs.asp?visjobogaftaler=1&visminihistorik=1&visfaktura=2&FM_job="&jobid&"&FM_aftale="&aftid&"&FM_kunde="&kid&"&id="&thisfakid&"&FM_usedatointerval=1&FM_start_dag_ival="&stDag&"&FM_start_mrd_ival="&stMrd&"&FM_start_aar_ival="&stAar&"&FM_slut_dag_ival="&slDag&"&FM_slut_mrd_ival="&slMrd&"&FM_slut_aar_ival="&slAar&""%>">Videre >></a>

                                                    </div>-->
                                                       
                                                       <%

                                            'else
                                                    Response.redirect "erp_opr_faktura_fs.asp?visjobogaftaler=1&visminihistorik=1&visfaktura=2&FM_job="&jobid&"&FM_aftale="&aftid&"&FM_kunde="&kid&"&id="&thisfakid&"&FM_usedatointerval=1&FM_start_dag_ival="&stDag&"&FM_start_mrd_ival="&stMrd&"&FM_start_aar_ival="&stAar&"&FM_slut_dag_ival="&slDag&"&FM_slut_mrd_ival="&slMrd&"&FM_slut_aar_ival="&slAar
                                           
                                            'end if

										    
                                              

											%>
											 
											<!--
											<script>
											
									        window.top.frames['erp3'].location.href = "erp_fak_godkendt_2007.asp?jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=thisfakid%>&FM_usedatointerval=1&FM_start_dag_ival=<%=stDag%>&FM_start_mrd_ival=<%=stMrd%>&FM_start_aar_ival=<%=stAar%>&FM_slut_dag_ival=<%=slDag%>&FM_slut_mrd_ival=<%=slMrd%>&FM_slut_aar_ival=<%=slAar%>&kid=<%=kid%>"
	                                        window.top.frames['erp2_2'].location.href = "erp_opr_faktura.asp?FM_type=<%=intType%>&FM_usedatokri=1&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&FM_kunde=<%=kid%>&FM_start_dag=<%=stDag%>&FM_start_mrd=<%=stMrd%>&FM_start_aar=<%=stAar%>&FM_slut_dag=<%=slDag%>&FM_slut_mrd=<%=slMrd%>&FM_slut_aar=<%=slAar%>"
	                                        
	                                        </script>
                                                -->
											
											<%
											
											
									'end if faknrfindesallerede

                                 '**VALIDERING If / ELSEs
							
                          end if
                                                	
						end if
					end if
			end if
	
	
	
	
	
	
	
	
	
	
	
	case else
	'********************************************************************************************'
    '**************************** Opret/rediger Faktura Data *********************************'
    '********************************************************************************************'
	
                                                
    call setFakPreInd()
    %>
	<!--#include file="inc/fak_inc_subs_2007.asp" -->
	

    
	<%

    '********************************************************************************************'
	'********************************** Sideindhold  ********************************************'
    '********************************************************************************************'
	
	
	call multible_licensindehavereOn()
	
	
	timeforbrug = 0
	faktureretBeloeb = 0
	
	pleft = 0
	ptop = 0
	
    arrsize = 500

    'redim thisaktidFlereTp(arrsize)
	Redim thisaktid(arrsize)
    Redim stdatoKriAktSpecifik(arrsize)
	Redim thisaktnavn(arrsize)
	Redim thisAktTimer(arrsize), thisAktTimerSum(arrsize)
	Redim thisAktBeloeb(arrsize)
	'Redim usefastpris(arrsize)
	Redim thisTimePris(arrsize)
	Redim thisaktfunc(arrsize)
	Redim thisaktfakbar(arrsize)
	Redim thisaktfaktor(arrsize)
	Redim thisaktsort(arrsize)
	
	Redim thisaktrabat(arrsize)
	
	Redim thisaktbesk(arrsize)
	Redim thisaktForkalk(arrsize)
	redim thisAktForkalkStk(arrsize)
	redim thisAktBudgetsum(arrsize)
	
	Redim thisAktEnhAng(arrsize)
	Redim thisAktEnhPris(arrsize)
	
	Redim thisAktValuta(arrsize)
	Redim thisAktText(arrsize)
	
	Redim thisaktCHK(arrsize) 
	Redim thisaktEnh(arrsize)
	
	redim thisAktFase(arrsize)
	redim thisMomsfri(arrsize)
	
	redim thisAktBgr(arrsize)
	redim thisJoborAkt_tp(arrsize)
	
    redim thisAktEnhpris(arrsize)

	if func = "red" then
	'*********************************************************************
	' Rediger faktura 
	'*********************************************************************
	jobids_orders = "0"
    preValgtModKonto = 0
	
						strSQL = "SELECT Fid, faknr, fakdato, jobid, timer, beloeb, moms, kommentar, dato, editor, betalt, b_dato, fakadr, "_
						&" att, faktype, konto, modkonto, varenr, enhedsang, jobbesk, "_
						&" visjoblog, visrabatkol, vismatlog, visjoblog_timepris, "_
						&" visjoblog_enheder, visjoblog_mnavn, rabat, visafsemail, visafsswift, visafstlf, "_
						&" visafsiban, visafscvr, vismodland, vismodatt, vismodtlf, "_
						&" vismodcvr, visafsfax, valuta, kurs, sprog, istdato, momskonto, visperiode, jobfaktype, "_
						&" betbetint, brugfakdatolabel, istdato2, "_
						&" momssats, modtageradr, usealtadr, vorref, fak_ski, showmatasgrp, hidesumaktlinier, "_
						&" sideskiftlinier, labeldato, fak_abo, fak_ubv, visikkejobnavn, hidefasesum, hideantenh, medregnikkeioms, fak_laast, afsender, vis_jobbesk, "_
                        &" kontonr_sel, fakglobalfaktor, totbel_afvige, fak_fomr "_
						&" FROM fakturaer WHERE Fid = "& id 
						
						'Response.Write strSQL
						'Response.Flush
						
						oRec.open strSQL, oConn, 3
						
						if not oRec.EOF then
						strEditor = oRec("editor")
						strFaknr = oRec("faknr")
						strTdato = formatdatetime(oRec("fakdato"), 2)
						'fakdatovente = oRec("fakdato")
						
						
						visafsemail = oRec("visafsemail")
						visAfsSwift = oRec("visafsswift")
					    visAfsIban = oRec("visafsiban")
					    visAfsCVR = oRec("visafscvr")
						visAfsTlf = oRec("visafstlf")
						visAfsFax = oRec("visafsfax")
						
						    intModland = oRec("vismodland")
						
						    if cint(intModland) <> 0 then
			                'strLand =  "<br>"& oRec("land")
			                'strLandShow = strLand
			                lndCHK = "CHECKED"
			                else
			                'strLandShow = "Danmark"
			                'strLand = ""
			                lndCHK = ""
			                end if
						
						
						
						'intModTlf = oRec("vismodtlf")
						intModCvr = oRec("vismodcvr")
							
								
						dbfunc = "dbred"		
						strKom = oRec("kommentar")
								
						strTimer = oRec("timer")
						strBeloeb = oRec("beloeb")
						'StrUdato = "12/12/2014"
						strDato = oRec("dato")
						intBetalt = oRec("betalt")
						strB_dato = oRec("b_dato")
						
						
						intFakid = oRec("Fid")
						intfakadr = oRec("fakadr")
						strAtt = oRec("att")
						
						intType = oRec("faktype")
						
						intKonto = oRec("konto")
						intModKonto = oRec("modkonto")
						
						strVarenr = oRec("varenr")
						intEnhedsang = oRec("enhedsang")
						
						strJobBesk = oRec("jobbesk")
						strKom = oRec("kommentar")
						
						visjoblog = oRec("visjoblog")
						visrabatkol = oRec("visrabatkol")
						vismatlog = oRec("vismatlog")
						
						visjoblog_timepris = oRec("visjoblog_timepris")
						visjoblog_enheder = oRec("visjoblog_enheder") 
						visjoblog_mnavn = oRec("visjoblog_mnavn")
						
						intRabat = oRec("rabat")
						
						valuta = oRec("valuta")
						kurs = oRec("kurs")
						valutaKursFak = kurs
						sprog = oRec("sprog")
						
						istDato = oRec("istdato")
						istDato2 = oRec("istDato2")
						
						momskonto = oRec("momskonto")
						visperiode = oRec("visperiode")
						
						jobfaktype = oRec("jobfaktype")
						
						betbetint = oRec("betbetint")
						brugfakdatolabel = oRec("brugfakdatolabel")
						
						strForfaldsdato = oRec("b_dato")
						
						momssats = oRec("momssats")
						modtageradr = trim(oRec("modtageradr"))
						
						usealtadr = oRec("usealtadr")
						
						vorref = oRec("vorref")
						
						ski = oRec("fak_ski")
						abo = oRec("fak_abo")
						ubv = oRec("fak_ubv")
						
						showmatasgrp = oRec("showmatasgrp")
						
						hidesumaktlinier = oRec("hidesumaktlinier")
						sideskiftlinier = oRec("sideskiftlinier")
						
						labelDato = oRec("labeldato")
						
						visikkejobnavn = oRec("visikkejobnavn")

                        hidefasesum = oRec("hidefasesum")
						
                        hideantenh = oRec("hideantenh")

                        medregnikkeioms = oRec("medregnikkeioms")

                        fak_laast = oRec("fak_laast")

                        afsender = oRec("afsender")


                        useBelob_red = strBeloeb 
                        useMoms_red = oRec("moms")

                        visjobbesk = oRec("vis_jobbesk")

                        kontonr_sel = oRec("kontonr_sel")

                        fakglobalfaktor = oRec("fakglobalfaktor")

                        totbel_afvige = oRec("totbel_afvige")

                        fak_fomr = oRec("fak_fomr")



                                '** NT ***'
                                select case lto
                                case "nt"
													            
													    
									strSQLshadowCopy = "SELECT jobid FROM fakturaer WHERE shadowcopy = 1 AND faknr = '"& strFaknr &"'"
                                    oRec6.open strSQLshadowCopy, oConn, 3
                                    while not oRec6.EOF  
													        
                                                               
                                    jobids_orders = jobids_orders  & ","& oRec6("jobid")
													                
                                    oRec6.movenext
                                    wend
                                    oRec6.close
													              
                                end select
                                    


						end if
						oRec.close
						
						
						if len(intModKonto) <> 0 then
						intModKonto = intModKonto
						else
						intModKonto = 0
						end if
						
						if intBetalt = 1 then
						betaltch = "checked"
						else
						betaltch = ""
						end if 
						
						
	    
	    
	                if jobid <> 0 then
            							
							            call faktureredetimerogbelob()
            							
            						
							            strSQL = "SELECT jobTpris, budgettimer, fastpris, jobknr, jobnr, jobnavn, "_
							            &" ikkebudgettimer, jobans1, jobans2, rekvnr, jobstatus, usejoborakt_tp, fastpris, s.navn AS aftalenavn, aftalenr, serviceaft, job_internbesk, j.kommentar, jo_bruttooms, alert "_
                                        &" FROM job AS j LEFT JOIN serviceaft s ON (s.id = serviceaft) WHERE j.id = " & jobid

                                        'Response.Write strSQL
                                        'Response.flush

							            oRec.open strSQL, oConn, 3
							            if not oRec.EOF then
            								
            								
            								
								            intIkkeBtimer = oRec("ikkebudgettimer")
								            intBudgettimer = oRec("budgettimer")
								            intJobTpris = oRec("jobTpris")
                                            bruttooms = oRec("jo_bruttooms")

								            fastpris = oRec("fastpris")
								            if intfakadr <> 0 then
								            kid = intfakadr
								            else
								            kid = oRec("jobknr")
								            end if
								            intjobnr = oRec("jobnr")
								            strJobnavn = oRec("jobnavn")
            								
								            jobans1 = oRec("jobans1")
								            jobans2 = oRec("jobans2")
								           
            								rekvnr = oRec("rekvnr")
            								jobstatus = oRec("jobstatus")
            								
            								usejoborakt_tp = oRec("usejoborakt_tp")
            								
            								usefastpris = oRec("fastpris")

                                            aftaleId = oRec("serviceaft")
			                                aftalenavn = oRec("aftalenavn") & " ("& oRec("aftalenr") &")"
                                            intaftnr = oRec("aftalenr")                                      								


            								job_internbesk = oRec("job_internbesk")
                                            job_tweet = oRec("kommentar")

                                            job_alert = oRec("alert")
            								
							            end if
							            oRec.close
            							
		            else
            		
			                            kid = intfakadr
            			                intaftnr = 0
			                            intEnheder = strTimer
			                            intPris = strBeloeb
			                            strBesk = strKom
                                        'strJobBesk = strBesk

                                        strSQL = "SELECT kundeid, besk, enheder, pris, varenr, navn, perafg, "_
		                                &" advitype, advihvor, stdato, sldato, valuta, aftalenr FROM serviceaft WHERE id = "& aftid
		                                oRec.open strSQL, oConn, 3 
		                                while not oRec.EOF 
		
		                                'kid = oRec("kundeid")
		                                'intEnheder = oRec("enheder")
		                                'intPris = oRec("pris")
		                                'strVarenr = oRec("varenr")
		                                strNavn = oRec("navn") & " ("& oRec("aftalenr") &")"
		                                'strJobBesk = oRec("besk") 'strNavn
		                                intPerafg = oRec("perafg")
		                                intAdvitype = oRec("advitype")
		                                intAdvihvor = oRec("advihvor")
		                                startdato = oRec("stdato") 
		                                slutdato = oRec("sldato") 
		                                'valuta = oRec("valuta")
                                        intaftnr = oRec("aftalenr")
		
		                                oRec.movenext
		                                wend
		                                oRec.close 
		
		
		                                jobfaktype = 0
		                                strBesk = strNavn&":&nbsp;"&strBesk
            		
		            end if
		'*******************************************************************************************
	    
	'oimg = "ac0010-24.gif"
	'oleft = 0
	'otop = -20
	'owdt = 400
	
	if cint(intType) <> 1 then
	typTxt = erp_txt_001
	else
	typTxt = erp_txt_002 
    end if

    oskrift = erp_txt_196 &" "& typTxt &" "& strFaknr 

    if jobid <> 0 then
    oskrift = oskrift &"<span style=""font-size:11px; font-weight:lighter;""> job no.: ("& intjobnr &")</span>"
	else
    oskrift = oskrift &"<span style=""font-size:11px; font-weight:lighter;""> aft.no.: ("& intaftnr &")</span>"
    end if
		
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
		
		'** Mulighed for at slette og se sidst redigeret dato ****'
		%><h4 style="padding:5px 0px 2px 5px;"><%=oskrift %> <span style="font-size:9px; font-weight:lighter;"><%=erp_txt_183 &" <b>"& strDato &"</b> " & erp_txt_184 & " <b>" & strEditor & "</b></span></h4>"%>
		<div id="sidetop" style="position:absolute; left:0px; top:5px; visibility:visible; z-index:100;">
	    
	    
	  
	    
			
			
			
		
		
		
		<% 
	
	
	
	
	
	else

	'*********************************************************************************************************************************************
	' Opret faktura 
	'*********************************************************************************************************************************************
	'*********************************************************************************************************************************************

            jobids_orders = "0"
            select case lto
            case "nt", "xintranet - local"

                                    
                                '**********************************************
                                '**** Sletter faktura kladder p� dette job ***'
                                '**********************************************

                                if cint(intType) = 0 then 'faktura // Skal ikke g�re det p� kreditnotaer

                                useOldfaknr = ""
                                kFids = "0"
                                strSQLfakkladder = "SELECT fid, faknr FROM fakturaer WHERE jobid = "& jobid &" AND betalt = 0 AND fak_laast = 0 ORDER BY fid " 'OR sup. invoice no =
                                oRec6.open strSQLfakkladder, oConn, 3
                                while not oRec6.EOF
                    
                                kFids = kFids & ", "& oRec6("fid")
                                useOldfaknr = oRec6("faknr") 'seneste kladde nr.

                                oRec6.movenext
                                wend
                                oRec6.close

                                strFaknr = useOldfaknr
                        

                                kFidsArr = split(kFids, ", ")
                                for f = 0 to UBOUND(kFidsArr)



                                            if cint(f) = 0 then


                                                strSQldelFak = "DELETE FROM fakturaer WHERE jobid = "& jobid &" AND betalt = 0 AND fak_laast = 0" 'OR sup. invoice no =
                                                oConn.execute(strSQldelFak)

                    
                                                '** Sletter shodows copies ***'
                                                strSQldelFakshadow = "DELETE FROM fakturaer WHERE faknr = '"& strFaknr &"' AND shadowcopy = 1" 'OR sup. invoice no =
                                                oConn.execute(strSQldelFakshadow)
                            

                                                strSQLdelAkt = "DELETE FROM aktiviteter WHERE job = "& jobid 
                                                oConn.execute(strSQLdelAkt)

                                                '*** Materiale forbrug slet **''
                                                strSQLdelMat = "DELETE FROM materiale_forbrug WHERE jobid = "& jobid 
                                                oConn.execute(strSQLdelMat)



                                            end if
                        



                                
                                strSQldelFak = "DELETE FROM faktura_det WHERE fakid = "& kFidsArr(f) 
                                oConn.execute(strSQldelFak)

                                strSQldelFak = "DELETE FROM fak_med_spec WHERE fakid = "& kFidsArr(f) 
                                oConn.execute(strSQldelFak)

                                strSQldelFak = "DELETE FROM fak_mat_spec WHERE matfakid = "& kFidsArr(f) 
                                oConn.execute(strSQldelFak)


                                next



                                '******************************************************************'
								'Tilknytter materialer til det job der bliver oprettet her.**'
								'******************************************************************'
			                    strEditor = session("user")
				                strDato = session("dato")        
                                varjobId = jobid
                                jobids = request("jobids")
                                jobids_orders = jobids                
               

                                 '** Findes der matforbrug i forvejen? *****
                                antalMat = 0
                                strSQLantalAkt = "SELECT count(id) AS antalmat FROM materiale_forbrug WHERE jobid = "& varjobId &" GROUP BY jobid"
                                oRec6.open strSQLantalAkt, oConn, 3
                                if not oRec6.EOF then
                            
                                antalMat = oRec6("antalmat")

                                end if
                                oRec6.close

                                

                                 if cint(antalMat) = 0 then 'tilf�jer materiale linjer


                                'response.write jobids
                                'response.end
                
                                jobids = replace(jobids, " ", "")
                                jobidsArr = split(jobids, ",")
            
                                for j = 0 TO UBOUND(jobidsArr)
                                
                                if j > 0 then
                                

                               
                
                                matantal = 0
                                matnavn = ""
                                matsalgspris = 0
                                matkobspris = 0
                                matvaluta = 0
                                matvarenr = 0

                               

                        
                                        strSQljob = "SELECT jobnavn, jobnr, jobstartdato, jobslutdato, jo_bruttooms, shippedqty, cost_price_pc, sales_price_pc, comm_pc, valuta, "_
                                        &" rekvnr, kunde_levbetint, lev_levbetint, fastpris, kunde_betbetint, lev_betbetint FROM job WHERE id = "& jobidsArr(j)
                                        'response.write strSQljob
                                        'response.flush                    
                    
                                        oRec6.open strSQljob, oConn, 3
                                        if not oRec6.EOF then

                                            matnavn = oRec6("jobnavn")

                                            
                                            matsalgspris = oRec6("sales_price_pc")
                                            

                                            matkobspris = oRec6("cost_price_pc")
                                            matantal = oRec6("shippedqty")
                                            matvaluta = oRec6("valuta") 
                                            matvarenr = oRec6("rekvnr")

                                            kunde_levbetint = oRec6("kunde_levbetint") 
                                            lev_levbetint = oRec6("lev_levbetint")

                                            kunde_betbetint = oRec6("kunde_betbetint") 
                                            lev_betbetint = oRec6("lev_betbetint")
                                            'matfaktor = oRec6("comm_pc")        
                        
                                            matava = oRec6("comm_pc")

                                            'matprisialt = (aktantal/1 * aktprisprstk/1) '** beregner udfra shipped   'oRec6("jo_bruttooms")

                                        end if
                                        oRec6.close


                                  


                                     '**** Mat. V�rdier ****'

                              
                                        jobid = varjobId
                                        medid = session("mid")
                                        '** OnTheFly: 1 / Fra lager : 0 
                               
		                                otf = 1 'request("matreg_onthefly")       
		                                matId = 0
                                        aktid = 0
                                        aftid = 0

                                      
		                                intAntal = matantal
		                                
                                        regdato = year(now)& "/"& month(now) & "/"& day(now)
                                        strDato = regdato
                                
    
                                        valuta = matvaluta 
                                        intkode = 2 '** Fakturerbar
                                        personlig = 0
                                        bilagsnr = ""
                                

                                        pris = matkobspris
                                        salgspris = matsalgspris 
                                        gruppe = 0

                                        navn = replace(matnavn, "'", "")
                                        varenr = matvarenr
                                        
                                        matreg_opdaterpris = 0 'request("matreg_opdaterpris")
                                        opretlager = 0  'request("matreg_opretlager")
                                        betegnelse = "" 'request("matreg_betegn")

                                        mat_func = "dbopr"

                                        'response.write "mat_func: "& mat_func &" onthefly: "& otf &" aktid:"& aktid &" Dato: "& regdato &" Navn: "& navn & " varenr: "& varenr &" opretlager: "& opretlager &" valuta: "& valuta &"<br>"
                                        'response.end
                                        matregid = request("matregid")

                                        matava = matava

                                        call indlaes_mat(matregid, otf, medid, jobid, aktid, aftid, matId, strEditor, strDato, intAntal, regdato, valuta, intkode, personlig, bilagsnr, pris, salgspris, navn, gruppe, varenr, opretlager, betegnelse, mat_func, matreg_opdaterpris, matava)
	


                                      
                                        
                                        end if 'j > 0
                        
                                        next 'jobidsArr

                                        end if' antalMat





                                        '*** tilbage til opr v�rdi
                                        varjobId = jobid
                                        'Til beregning af forfaldsdato 
                                           dt_actual_etd = now
                                           transportVal = ""
                                           strSQLffdato = "SELECT dt_actual_etd, dt_actual_eta, kunde_levbetint, transport, jobknr FROM job WHERE id = "& varjobId &""
                                           oRec6.open strSQLffdato, oConn, 3
                                           if not oRec6.EOF then
                                    
                                            'if oRec6("jobknr") = 1448 OR oRec6("jobknr") = 1457 then 'POMP
                                            'transportVal = "0"
                                            'else
                                            transportVal = oRec6("transport")
                                            'end if
                                            if cint(oRec6("kunde_levbetint")) <> 2 then
                                            dt_actual_etd = oRec6("dt_actual_etd")
                                            else '* ON DDP USE ETA date
                                            dt_actual_etd = oRec6("dt_actual_eta")
                                            end if

                                           end if 
                                           oRec6.close


                            end if 'intType




            end select


	
		
		
		
		
		'*** Find kunde vedr opret faktura p� JOB ***
		if jobid <> 0 then
		
		call faktureredetimerogbelob()
		
		strSQL = "SELECT jobTpris, budgettimer, fastpris, jobknr, jobnr, "_
		&" jobnavn, ikkebudgettimer, jobans1, jobans2, kundekpers, beskrivelse, j.valuta, "_
		&" jobstartdato, jobslutdato, jobfaktype, rekvnr, jobstatus, usejoborakt_tp, ski, abo, ubv, s.navn AS aftalenavn, jfak_moms, jfak_sprog, "_
        &" aftalenr, serviceaft, job_internbesk, j.kommentar, jo_bruttooms, altfakadr, supplier, alert, lincensindehaver_faknr_prioritet_job FROM job j "_
        &" LEFT JOIN serviceaft s ON (s.id = j.serviceaft) WHERE j.id = " & jobid

        'Response.write strSQL
        'Response.flush
		oRec.open strSQL, oConn, 3
		
		if not oRec.EOF then
			
			intIkkeBtimer = oRec("ikkebudgettimer")
			intBudgettimer = oRec("budgettimer")
			intJobTpris = oRec("jobTpris")
            bruttooms = oRec("jo_bruttooms")

			fastpris = oRec("fastpris")
			
            '** NT commisionsordre ***
            if cint(fastpris) = 2 then '** faktura skal h�gtes p� leverand�r
            kid = oRec("supplier")
            else 
			kid = oRec("jobknr")
			end if

            intjobnr = oRec("jobnr")
			strJobnavn = oRec("jobnavn")
			
			jobans1 = oRec("jobans1")
			jobans2 = oRec("jobans2")
			
			intKundekpers = oRec("kundekpers")
			strJobBesk = oRec("beskrivelse")
			
			valuta = oRec("valuta")
			jobstdato = oRec("jobstartdato")
			jobsldato = oRec("jobslutdato")
			
			'if len(request("FM_jobfaktype")) <> 0 then
			'jobfaktype = request("FM_jobfaktype")
			'Response.Write "<br><br>her" & jobfaktype
			'else
			'jobfaktype = oRec("jobfaktype")
			'end if
			
			jobfaktype = 0
			
			rekvnr = oRec("rekvnr")
			jobstatus = oRec("jobstatus")
			
			
			usejoborakt_tp = oRec("usejoborakt_tp")
			ski = oRec("ski")
			
			abo = oRec("abo")
			ubv = oRec("ubv")

            aftaleId = oRec("serviceaft")
			aftalenavn = oRec("aftalenavn") & " ("& oRec("aftalenr") &")"

			'jobtimerforkalk = intIkkeBtimer + intBudgettimer

            job_internbesk = oRec("job_internbesk")
            job_tweet = oRec("kommentar")

            usealtadr = oRec("altfakadr")

            useBelob_red = bruttooms

            jfak_moms = oRec("jfak_moms")
            sprog = oRec("jfak_sprog")

            job_alert = oRec("alert")


                    '*** Finder momssats '***
                    strSQLmoms = "SELECT moms FROM fak_moms WHERE id = " & jfak_moms
                    oRec2.open strSQLmoms, oConn, 3
                    if not oRec2.EOF then

                    momssats = oRec2("moms")

                    end if
                    oRec2.close

                   
			

                    '*** Finder fomr p� job og dermed ogs� forvalgtr konto ***'
                    strSQLfomr = "SELECT for_fomr FROM fomr_rel WHERE for_jobid = "& jobid &" AND for_aktid = 0"
                    oRec2.open strSQLfomr, oConn, 3
                    if not oRec2.EOF then

                    fak_fomr = oRec2("for_fomr")

                    end if
                    oRec2.close

            lincensindehaver_faknr_prioritet_job = oRec("lincensindehaver_faknr_prioritet_job")

		end if
		oRec.close





		
		else 
		
		'*** Find kunde vedr opret faktura p� Aftale ***'
		
        fak_fomr = 0

		call faktureredetimerogbelob()
		
		strSQL = "SELECT s.kundeid, s.besk, s.enheder, s.pris, s.varenr, s.navn, s.perafg, "_
		&" s.advitype, s.advihvor, s.stdato, s.sldato, s.valuta, s.aftalenr, kfak_moms, kfak_sprog FROM serviceaft AS s "_
        &" LEFT JOIN kunder AS k ON (k.kid = s.kundeid) WHERE s.id = "& aftid
		oRec.open strSQL, oConn, 3 
		while not oRec.EOF 
		
		kid = oRec("kundeid")
		'strBesk = oRec("besk")
		intEnheder = oRec("enheder")
		intPris = oRec("pris")
		strVarenr = oRec("varenr")
		strNavn = oRec("navn") & " ("& oRec("aftalenr") &")"
		strJobBesk = oRec("besk") 'strNavn
		intPerafg = oRec("perafg")
		intAdvitype = oRec("advitype")
		intAdvihvor = oRec("advihvor")
		startdato = oRec("stdato") 
		slutdato = oRec("sldato") 
		valuta = oRec("valuta")
        usealtadr = 0


        kfak_moms = oRec("kfak_moms")
        sprog = oRec("kfak_sprog")
                    
                      '*** Finder momssats '***
                    strSQLmoms = "SELECT moms FROM fak_moms WHERE id = " & kfak_moms
                    oRec2.open strSQLmoms, oConn, 3
                    if not oRec2.EOF then

                    momssats = oRec2("moms")

                    end if
                    oRec2.close
                
		
		oRec.movenext
		wend
		oRec.close 
		

		jobfaktype = 0
		strBesk = strNavn&":&nbsp;"&strBesk
		
		
		end if '*** Aft / Job **'


    if cint(intType) <> 1 then
	typTxt = erp_txt_001 
	else
	typTxt = erp_txt_002 
	end if

	
    oskrift = erp_txt_197 &" "& typTxt

    if jobid <> 0 then
    oskrift = oskrift &" <span style=""font-size:11px; font-weight:lighter;""> "& erp_txt_419 &": ("&  intjobnr &")</span>"
	else
    oskrift = oskrift &" <span style=""font-size:11px; font-weight:lighter;""> "& erp_txt_420 &": ("&  intaftnr &")</span>"
    end if
	
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	

       

	%>
	
	<!-- main -->

             

	<h4 style="padding:5px 0px 2px 5px;"><%=oskrift %></h4>
    <div id="sidetop" style="position:absolute; left:0px; top:5px; visibility:visible; z-index:100;">
    
   
    
   
	
	
	<%
		
		
		kurs = ""

        '*** Sprog bliver styret i erp_xml_inc p� bg af medarbejder valgt sprog 
        'select case lto
        'case "epi_uk", "nt", "epi_au", "epi_de"
        'sprog = 2
        'case else
		'sprog = 1
        'end select
		
        '***** Hvis kundestamdata land <> DK / eller forvalgt kontperson (fakadr) <> DK **'
        '** SE kan evt. ogs� v�lges


        if len(trim(sprog)) <> 0 then
        sprog = sprog
        else
        sprog = 1
        end if

        if func <> "red" AND lto = "bf" then 'Overruler p� opret
        valuta = 5
        sprog = 2
        end if

		strAtt = intKundekpers
		
		intKonto = 0
		intModKonto = 0
		intType = intType 
		visjoblog = chklog
		
		
		 if func <> "red" then

            select case lto 
            case "bf", "intranet - local"
            visjoblog = 1
            case else
		    visjoblog = chklog
		    end select

        else
            visjoblog = chklog
        end if


		
		'*** Skal rabat kolonne v�re sl�et til default. ***'
		select case lto
		case ""
		visrabatkol = 1
		case else
		visrabatkol = 0
		end select 
		
		
		vismatlog = chklog
		'lastFaknr = 0
		
		visjoblog_timepris = 1
		visjoblog_enheder = 1
		visjoblog_mnavn = 1
		brugfakdatolabel = 0
		
		
		strEditor = ""
		strTdato = "" 'month(now)&"/"&day(now)&"/"& year(now)
		istDato = ""
		istDato2 = ""
		
		strKom = ""
		dbfunc = "dbopr"
		varSubVal = "Opretpil" 
		
		if len(chkemail) <> 0 then
		chkemail = chkemail
		else
		chkemail = 0 
		end if
		
		if cint(chkemail) = 1 OR (lto = "jttek") OR (lto = "hestia") then
		visafsemail = 1
		else
		visafsemail = 0
		end if
		
        select case lto
        case "dencker", "epi", "epi_no", "epi_sta", "epi_ab", "epi_uk", "nt"
        visAfsSwift = 1
		visAfsIban = 1
	    visAfsCVR = 1
        case else
		visAfsSwift = 0
		visAfsIban = 0
	    visAfsCVR = 1
		end select

		if len(chktlffax) <> 0 then
		chktlffax = chktlffax
		else
		chktlffax = 0
		end if
		
        
        

		if cint(chktlffax) = 1 OR (lto = "epi") OR (lto = "epi_no") OR (lto = "epi_sta") OR lto = "epi_ab" OR lto = "epi_uk" OR (lto = "intranet - local") OR lto = "hestia" then
		visAfsTlf = 1
		visAfsFax = 0
		else
		visAfsTlf = 0
		visAfsFax = 0
		end if
		
        if lto = "jttek" then
		visAfsTlf = 1
		visAfsFax = 1
		end if

		'intModland = 0
		'intModTlf = 0

        if (lto = "epi") OR (lto = "epi_no") OR (lto = "epi_sta") OR (lto = "epi_ab") OR (lto = "epi_uk") OR lto = "intranet - local" then
        intModCvr = 1
        else
		intModCvr = 0
		end if

        
		momskonto = 1
        preValgtModKonto = 0




            '*** Finder kunde til moms ER DET EU / uden for DK kunde ***'
            thisLand = ""
            strSQL = "SELECT land "_
		    &" FROM kunder WHERE kid =" & kid  		
                		
		    'Response.Write strSQL
		    'Response.flush
                		
                		
		    oRec.open strSQL, oConn, 3
		    if not oRec.EOF then

            thisLand = oRec("land")

            
            select case thisLand
            case "DK", "Danmark", "Denmark"
            preValgtModKonto = 1
            case else 
            preValgtModKonto = 2
            end select


            end if
            oRec.close


             '*** Finder land p� afsender ***'
            thisAfsenderLand = ""
            strSQL = "SELECT land "_
		    &" FROM kunder WHERE useasfak = 1"
                		
		    'Response.Write strSQL
		    'Response.flush
                		
                		
		    oRec.open strSQL, oConn, 3
		    if not oRec.EOF then

            thisAfsenderLand = oRec("land")

            end if
            oRec.close

        'response.write "thisAfsenderLand: "& thisAfsenderLand & " thisLand: "& thisLand


        '** Hvis modtager er et andet land end afsender p� faktura **'
        'if lto = "synergi1" then

        '** Overruler momssat p� job og kunde / SLET DENNE INDSTILLING, DEN ER KLAR TIL AT NEDARVE
        select case lto
        case "epi_uk"
        momssats = 20
        case else
        momssats = momssats '25
        end select
       

        if lto = "nt" AND (kid = 1448 OR kid = 1458) then 'pomp / sisters
        momssats = 0
        end if

        if lto = "adra" OR lto = "bf" then 'altid 0
        momssats = 0
        end if
       


        visperiode = 0
		
		
		vorref = "-"
		showmatasgrp = 0
		
        select case lto
        case "synergi1"
        hidesumaktlinier = 1
        case else
		hidesumaktlinier = 0
        end select

        hidefasesum = 0

        medregnikkeioms = 0

        fak_laast = 0


        '** Forvalgt bankkonto valgt pga af valuta ***' 
        select case lto
        case "epi" 'KUN DK til at starte med "epi_no", "epi_sta", "epi_ab", "epi_de", "epi_au"
            
            select case valuta
            case 1 'DKK
            kontonr_sel = 0
            case 2 'NOK
            kontonr_sel = 2
            case 4 'EUR
            kontonr_sel = 1
            case 8 'SEK
            kontonr_sel = 3
            case 9 'USD
            kontonr_sel = 4
            case 10 'GBP
            kontonr_sel = 5
            case else 'DKK
            kontonr_sel = 0
            end select

        case "nt"
                
            select case valuta 'valuta 0 findes ikke
            case 1
            kontonr_sel = 0
            case 2
            kontonr_sel = 1
            case 3
            kontonr_sel = 2
            case 4 'RMD KINA
            kontonr_sel = 3
            case else
            kontonr_sel = 0
            end select

        case else
        kontonr_sel = 0
        end select
        


        select case lto
        case "epi", "epi_no", "epi_sta", "epi_ab", "epi_uk"
        hideantenh = 1
        case else
        hideantenh = 0
        end select
		
		select case lto
	    case "dencker", "tooltest", "outz", "optionone"
	    sideskiftlinier = 11
        case "synergi1", "intranet - local"
	    sideskiftlinier = 0
        case else
	    sideskiftlinier = 0
	    end select
		
		
		visikkejobnavn = 0
        afsender = 0

        visjobbesk = 0
				
		'*****************************************'
		'*****************************************'


	end if '** Opret / Red Faktura
    '******************************************************************************************************************************************
	'******************************************************************************************************************************************
    '******************************************************************************************************************************************


    '*** Special indstillinger ved PC antal grundfos
    pcSpecial = 0
    select case lto
    case "dencker"
    
        if cint(fak_fomr) = 16 then
        pcSpecial = 1
        end if

    case "intranet - local"

        if cint(fak_fomr) = 5 then
        pcSpecial = 1
        end if

   
    
    end select
	
	if visafsemail = 1 then
	visafsemailCHK = "CHECKED"
	else
	visafsemailCHK = ""
	end if
	
	if cint(visAfsTlf) = 1 then
	visAfsTlfCHK = "CHECKED"
	else
	visAfsTlfCHK = ""
	end if
	
	if cint(visAfsFax) = 1 then
	visAfsFaxCHK = "CHECKED"
	else
	visAfsFaxCHK = ""
	end if
	
	if visAfsIban = 1 then
	visAfsIbanCHK = "CHECKED"
	else
	visAfsIbanCHK = ""
	end if
	
	if visAfsSwift = 1 then
	visAfsSwiftCHK = "CHECKED"
	else
	visAfsSwiftCHK = ""
	end if
	
	if cint(visafsCVR) = 1 then
	visafscvrCHK = "CHECKED"
	else
	visafscvrCHK = ""
	end if
	
	'if intModland = 1 then
	'intModlandCHK = "CHECKED"
	'else
	'intModlandCHK = ""
	'end if
	
	'if intModTlf = 1 then
	'intModTlfCHK = "CHECKED"
	'else
	'intModTlfCHK = ""
	'end if
	
	if cint(intModCvr) = 1 then
	intModCvrCHK = "CHECKED"
	else
	intModCvrCHK = ""
	end if
	
	
	if visikkejobnavn <> 0 then
	visikkejobnavnCHK = "CHECKED"
	else
	visikkejobnavnCHK = ""
	end if
	


            '** Henter Global faktor *****'
        strSQLfaktor = "SELECT globalfaktor FROM licens WHERE id = 1"
        oRec.open strSQLfaktor, oConn, 3 
        if not oRec.EOF then 
		
	    globalfaktor = oRec("globalfaktor")
		
		end if
		oRec.close 


        

	%>
	
	                
	                
	                <!-- Valuta load -->

                      <div id="valutaload" style="position:absolute; top:61px; left:125px; height:200px; display:none; visibility:hidden; border:2px red dashed; background-color:#FFFFe1; padding:20px; width:300px; height:150px; z-index:20000000;">
	                                  <b><%=erp_txt_185 %></b><br /><br />
                                          <img src="../ill/loaderbar.gif" /><br /><br />
                                         <%=erp_txt_186 %><br />
                          &nbsp;
                     </div>
	               
	               
                <%select case lto
                    case "nt", "xintranet - local", "bf"
                    sideDivVzb = "hidden"
                    sideDivDsp = "none"
                    sideDivVzbAlt = "visible"
                    sideDivDspAlt = ""

                    case else
                    sideDivVzb = "visible"
                    sideDivDsp = ""
                    sideDivVzbAlt = "hidden"
                    sideDivDspAlt = "none"

                    end select
                    %>


         

	               <!-- Sideload -->
	              <div id="sideload" style="position:absolute; top:161px; left:205px; display:; visibility:visible; border:3px #cccccc solid; padding:2px; background-color:#FFFFFF; padding:10px; width:300px; z-index:20000;">
	              
                        <table cellpadding=0 cellspacing=10 border=0 width=100%><tr><td>
	                    <img src="../ill/outzource_logo_200.gif" />
	                    <br />
	                    
	                    
                         <%

                         antalaktCount = 0
                         if jobid <> 0 then
	                    
	                    strSQL = "SELECT COALESCE(COUNT(a.id)) AS antalaktCount FROM aktiviteter a "_
                        &" LEFT JOIN akt_typer AS aty ON (aty.aty_id = a.fakturerbar) WHERE job = "& jobid &" AND aktstatus = 1 AND aty_on_invoice = 1 AND aty.aty_id IS NOT NULL GROUP BY job" 
	                    
	                    'Response.Write strSQL
	                    'Response.flush
	                    
	                    oRec.open strSQL, oConn, 3
	                    if not oRec.EOF then
                            antalaktCount = cdbl(oRec("antalaktCount"))/1
	                    end if
	                    oRec.close 

                        if cint(antalaktCount) => 100 then
                        %>
                        <br />
                        <%=erp_txt_187 %><br /><br /> <%=erp_txt_188 %><br />

                        <% 'respoonse.flush
                        else
	                    %>
	                    
	                    <%=erp_txt_194 %> <b><%=antalaktCount & " " & erp_txt_189 %></b>  <%=erp_txt_190 %><br /><br />
                        <b><%=erp_txt_191 %></b><br />
	                      
                       
                            <%=erp_txt_192 & " " & formatnumber((cdbl(antalaktCount)*0.2), 0) & " " & erp_txt_193 & " <b>" & formatnumber((antalaktCount*1.4), 0) %> </b> sek.
                           
                        
                        <%end if %>

                        <%else %>
                        <b><%=erp_txt_191 %></b><br />
                        <%=erp_txt_195 %>
                        <%end if%>

	                    </td><td align=right valign=top style="padding:15px 20px 20px 20px;">
	                    <img src="../inc/jquery/images/ajax-loader.gif" />
	
	                    </td></tr></table>
                  
                     
               
	              </div>
	                
	                
	                
	                
	                
	                <!-- Menu -->
	                
	                <div id=menu style="position:absolute; top:31px; height:65px; left:5px; display:none; visibility:hidden;">
	             

                   

	                <div id="knap_joblogdiv" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:-30px; width:50px; left:413px; border:1px #8cAAe6 solid; padding:3px 3px 3px 3px; background-color:#FFFFFF;">
                    <a href="#" onclick="showdiv('joblogdiv')" class=vmenu><%=erp_txt_038 %></a>
	                </div>
	                
	                <div id="knap_matlogdiv" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:-30px; width:150px; left:473px; border:1px #8cAAe6 solid; padding:3px 3px 3px 3px; background-color:#FFFFFF;">
                    <a href="#" onclick="showdiv('matlogdiv')"  class=vmenu><%=erp_txt_198 %></a>
	                </div>

                    <div id="knap_faksogdiv" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:-30px; width:100px; left:633px; border:1px #8cAAe6 solid; padding:3px 3px 3px 3px; background-color:#FFFFFF; z-index:2000000;">
                    <a href="erp_fakturaer_find.asp" class=vmenu target="_blank"><%=erp_txt_199 %></a>
	                </div>

	                
	                
	                
	               
                    
                    <div id="knap_fidiv" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:15px; width:95px; left:0px; border:1px #8cAAe6 solid; border-right:0px #8cAAe6 solid; padding:5px 5px 5px 5px; background-color:#ffff99;">
                        <table cellspacing=0 cellpadding=0 border=0 width=100%>
                        <tr><td colspan=2 class=lille><%=erp_txt_200 %></td></tr>
                        <tr>
                        <td><a href="#" onclick="showdiv('fidiv')" class=vmenu><%=erp_txt_001 %><br /> <%=erp_txt_201 %></a></td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pil_fak.gif" border=0 /> </td></tr></table>
	                </div>
	                
	                <%if jobid <> 0 then
	                wdt = "100"
	                else
	                wdt = "100"
	                end if%>
	                
	                <div id="knap_modtagdiv" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:15px; width:<%=wdt%>px; left:95px; border:1px #8cAAe6 solid; border-right:0px #8cAAe6 solid; border-left:0px #8cAAe6 solid; padding:5px 5px 5px 5px; background-color:#ffffff;">
                     <table cellspacing=0 cellpadding=0 border=0 width=100%>
                     <tr><td colspan=2 class=lille><%=erp_txt_202 %></td></tr>
                     <tr>
                        <td><a href="#" onclick="showdiv('modtagdiv')" class=vmenu><%=erp_txt_203 %><br /> <%=erp_txt_204 %></a>
                        </td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pil_fak.gif" border=0 /> </td></tr></table>
	                </div>
	               
	                
	               
	                
	                <div id="knap_jobbesk" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:15px; width:105px; left:195px; border:1px #8cAAe6 solid; border-right:0px #8cAAe6 solid; border-left:0px #8cAAe6 solid; padding:5px 5px 5px 5px; background-color:#ffffff;">
                     <table cellspacing=0 cellpadding=0 border=0 width=100%>
                     <tr><td colspan=2 class=lille><%=erp_txt_205 %></td></tr>
                     <tr>
                        <td><a href="#" onclick="showdiv('jobbesk')" class=vmenu><%=erp_txt_206 %><br /><%=erp_txt_207 %></a></td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pil_fak.gif" border=0 /> </td></tr></table>
	                </div>
	               
	                
	               
	                <div id="knap_aktdiv" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:15px; width:80px; left:275px; border:1px #8cAAe6 solid; border-right:0px #8cAAe6 solid; border-left:0px #8cAAe6 solid; padding:5px 5px 5px 5px; background-color:#ffffff;">
                     <table cellspacing=0 cellpadding=0 border=0 width=100%>
                     <tr><td colspan=2 class=lille><%=erp_txt_208 %></td></tr>
                     <tr>
                        <td><a href="#" onclick="showdiv('aktdiv')" id="menushowakt" class=vmenu><%=erp_txt_001%><br /><%=erp_txt_124%></a>
                        </td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pil_fak.gif" border=0 /> </td></tr></table>
	                </div>
	               
	                
	                
	                <div id="knap_matdiv" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:15px; width:95px; left:355px; border:1px #8cAAe6 solid; border-right:0px #8cAAe6 solid; border-left:0px #8cAAe6 solid;  padding:5px 5px 5px 5px; background-color:#ffffff;">
                     <table cellspacing=0 cellpadding=0 border=0 width=100%>
                     <tr><td colspan=2 class=lille><%=erp_txt_209 %></td></tr>
                     <tr>
                        <td><a href="#" onclick="showdiv('matdiv')" class=vmenu><%=erp_txt_210 %><br /><%=erp_txt_211 %></a>
                        </td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pil_fak.gif" border=0 /> </td></tr></table>
	                </div>
	                
	            
	                
	                
	                <div id="knap_betdiv" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:15px; width:112px; left:450px; border:1px #8cAAe6 solid; border-left:0px #8cAAe6 solid; padding:5px 5px 5px 5px; background-color:#ffffff;">
                     <table cellspacing=0 cellpadding=0 border=0 width=100%>
                  
	                 <tr><td colspan=2 class=lille><%=erp_txt_212 %></td></tr>
                   
                     <tr>
                        <td><a href="#" onclick="showdiv('betdiv')" class="vmenu"><%=erp_txt_213 %><br /><%=erp_txt_214 %></a>
                        </td>
                        <td style="padding:3px 2px 0px 4px;"><img src="../ill/pile.gif" border=0 /> </td></tr></table>
	                </div>
	                
	                
	               
	                
	              </div>


        <!-- Vis ikke faktura til redigering -->
         <div id="visikkefak" style="position:absolute; display:<%=sideDivDspAlt%>; visibility:<%=sideDivVzbAlt%>; top:105px; height:296px; width:670px; padding:20px; background-color:#FFFFFF; left:5px; z-index:2;"> 
             Your Invoice is now ready for rewiev. <br /><br />
             Click on the green botton in the upper right corner to review your invoice. <br /><br />
             <%if level = 1 OR (lto = "bf" AND (level <= 2 OR level = 6)) then %>
             <span id="sp_editfak" style="color:#5582d2;"><u>Click here to edit</u></span>
             <%end if %>
            </div>



	
	
	    <!-- gen indl�s faktura (interval �ndring) KUN ved opret, ellers bruges intervl gemt med fak. -->
	             
	    <div id="genindlaes" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:203px; left:455px; z-index:2;">

		<form action="erp_opr_faktura_fs.asp?FM_kunde=<%=kid %>&FM_job=<%=jobid%>&FM_aftale=<%=aftid%>&reset=1&func=<%=func%>&FM_usedatokri=1&fid=<%=id%>&visjobogaftaler=1&visfaktura=1&visminihistorik=1" method="post" target="_top">
		
		
		
		<input type="hidden" name="FM_start_dag" id="gen_FM_start_dag" value="<%=request("FM_start_dag")%>">
		<input type="hidden" name="FM_start_mrd" id="gen_FM_start_mrd" value="<%=request("FM_start_mrd")%>">
		<input type="hidden" name="FM_start_aar" id="gen_FM_start_aar" value="<%=request("FM_start_aar")%>">
		<input type="hidden" name="FM_slut_dag" id="gen_FM_slut_dag" value="<%=request("FM_slut_dag")%>">
		<input type="hidden" name="FM_slut_mrd" id="gen_FM_slut_mrd" value="<%=request("FM_slut_mrd")%>">
		<input type="hidden" name="FM_slut_aar" id="gen_FM_slut_aar" value="<%=request("FM_slut_aar")%>">
		  
		  <%if func <> "red" then %> 
		  <input id="Button1" onclick="opdaterFakdato()" type="submit" value="<%=erp_txt_415 %> >>" style="font-size:9px;" /> <span style="color:#999999; font-size:9px;"><%=erp_txt_215 %></span>
	      <%end if %>
		
		</form>
		</div>
        
	
	
	
    
           
    
    
    
	<form name="main" id="main" action="erp_opr_faktura_fs.asp?menu=stat_fak&visjobogaftaler=1&visminihistorik=1&visfaktura=1&func=<%=dbfunc%>" method="post"><!--erp_fak.asp?menu=stat_fak&func=<%=dbfunc%>-->
	<%
	'*** Bruger altid datointerval ****
	usedt_ival = 1
	%>
	
	
    <input type="hidden" name="tjekantalakt_all" id="tjekantalakt_all" value="<%=antalaktCount%>">
    <input type="hidden" name="FM_fak_laast" id="Hidden4" value="<%=fak_laast%>">
	<input type="hidden" name="FM_usedatointerval" value="<%=usedt_ival%>">
	<input type="hidden" name="FM_start_dag_ival" value="<%=request("FM_start_dag")%>">
	<input type="hidden" name="FM_start_mrd_ival" value="<%=request("FM_start_mrd")%>">
	<input type="hidden" name="FM_start_aar_ival" value="<%=request("FM_start_aar")%>">
	<input type="hidden" name="FM_slut_dag_ival" value="<%=request("FM_slut_dag")%>">
	<input type="hidden" name="FM_slut_mrd_ival" value="<%=request("FM_slut_mrd")%>">
	<input type="hidden" name="FM_slut_aar_ival" value="<%=request("FM_slut_aar")%>">

    <input type="hidden" name="rdir" value="<%=rdir%>">    
	
	<input type="hidden" name="FM_fakint_ival" value="<%=trim(request("FM_fakint"))%>">
			
	
	<!--<input type="hidden" name="FM_type" value="<%=intType%>">-->
	<input type="hidden" name="FM_jobnavn" value="<%=strJobnavn%>">
	<input type="hidden" name="FM_job" id="FM_job" value="<%=jobid %>"> <!-- <%=Request("FM_job")%> -->
	<input type="hidden" name="id" value="<%=id%>">
	
	<!--<input type="hidden" name="jobid" value="<%=jobid%>">-->
	
	<input type="hidden" name="FM_aftale" id="FM_aftale" value="<%=aftid%>">
	<input type="hidden" name="FM_job_tilknyttet_aftale" value="<%=request("FM_job_tilknyttet_aftale")%>">
    <input type="hidden" name="jobids_orders" id="jobids_orders" value="<%=jobids_orders%>">
	<!--<input type="hidden" name="FM_aftnr" value="<%=sogaftnr%>">-->
	<input type="hidden" name="FM_kunde" value="<%=kid%>">
	
	<!--<input id="jobelaft" name="jobelaft" type="hidden" value="<%=jobelaft%>"/>-->
	<input id="faknr" name="faknr" type="hidden" value="<%=strFaknr%>"/><!-- Denne bruges vist ikke NT -->
	<input id="lastopendiv" name="lastopendiv" type="hidden" value="fidiv"/>
	
    
	<input id="jobfaktype" name="jobfaktype" type="hidden" value="<%=jobfaktype%>"/>
	
    <input id="thisfile" name="" type="hidden" value="<%=thisfile%>"/>
    <input id="func" name="" type="hidden" value="<%=func%>"/>
    <input id="fastpris" name="" type="hidden" value="<%=fastpris%>"/>

    
        


    <input id="Hidden1" name="FM_jobonoff" type="hidden" value="<%=vislukkedejob%>"/>
        
    
    
    <%call godkendknap() %>

        
	
	            
	 <%'** Valuta kurser **'
    strSQLval = "SELECT id, valuta, valutakode, kurs FROM valutaer WHERE id <> 0 ORDER BY id " 
    oRec.open strSQLval, oConn, 3
    while not oRec.EOF 
    %>
    <input id="valutakurs_<%=oRec("id")%>" name="valutakurs_<%=oRec("id")%>" value="<%=oRec("kurs") %>" type="hidden" />
    <%
    oRec.movenext
    wend
    oRec.close%>



        
	

       
	   <!-- Modtager og Afsender -- vises f�rst p� fakturalayout -->
	
	  <div id="modtagdiv" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:105px; width:720px; left:5px; border:1px #8cAAe6 solid;">
                     
                       
                     <table cellspacing="0" cellpadding="0" border="0" width=100% bgcolor="#ffffff">
                     <tr><td valign=top style="width:350px;">
                     
                    <table cellspacing="0" cellpadding="5" border="0" width="100%">
	                <tr>
	                    <td style="border:1px #ffffff solid;" bgcolor="#8caae6" class=alt><b><%=erp_txt_216 %> </b></td>
	                </tr>
	                
                		<%
                		if func <> "red" then
		                strSQL = "SELECT kid, kkundenavn, kkundenr, adresse, "_
		                &" postnr, city, land, telefon, cvr, betbet, ktype, kt.rabat, ean, betbetint"_
		                &" FROM kunder LEFT JOIN kundetyper kt ON (kt.id = ktype) WHERE kid =" & kid  		
                		
		                'Response.Write strSQL
		                'Response.flush
                		
                		
		                oRec.open strSQL, oConn, 3
		                if not oRec.EOF then 
			                'intKnr = oRec("kkundenr")
			                strKnavn = trim(oRec("kkundenavn"))
			                strKadr = trim(oRec("adresse"))
			                strKpostnr = trim(oRec("postnr"))
			                strBy = trim(oRec("city"))
			                strLand = trim(oRec("land"))
			                intKid = oRec("Kid")
			                intCVR = oRec("cvr")
			                intTlf = oRec("telefon")
			                
			                if len(trim(oRec("ean"))) <> 0 then
			                ean = trim("EAN: "&oRec("ean"))
			                else
			                ean = ""
			                end if
			             
                			
                            if intCVR <> "0" then
                            intModCvrCHK = "CHECKED"
                            end if
                			
                			betbetint = oRec("betbetint")
			                intRabat = oRec("rabat")
			                
			        
			                '*** Betalings betingelser ****
			                select case lto
                            case "synergi1"
                                if len(trim(oRec("betbet"))) <> 0 then
                                strKom = oRec("betbet")
                                else
                                strKom = "14 dage netto"
                                end if
                            case else
                            strKom = oRec("betbet")
			                end select
			                
			                if lcase(strLand) <> "danmark" OR lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "dencker" OR lto = "epi_ab" OR lto = "epi2017" then
			                strLand = oRec("land")
			                strLandShow = strLand
			                lndCHK = "CHECKED"
			                else
			                strLandShow = "Danmark"
			                strLand = ""
			                lndCHK = ""
			                end if
                			
                			modtageradr = oRec("kkundenavn") & "<br>"& oRec("adresse") & "<br>"& oRec("postnr") & " " &oRec("city")
                            
                            if len(strLand) <> 0 then
                            modtageradr = modtageradr &"<br>"& strLand
                            end if

                            if cint(intModCvr) = 1 AND len(intCVR) <> 0 then
                            modtageradr = modtageradr & "<br>CVR: "& intCVR
                            end if 
                            
                            if len(ean) <> 0 then
                            modtageradr = modtageradr & "<br>" & ean
                            end if
                			
		                end if
                		
		                oRec.close
		                
		                else
		                modtageradr = modtageradr
		                betbetint = betbetint
			            intRabat = 100 * intRabat 
		                end if
		                
		                thisakt_rabat = intRabat


                            '*** Overruler modtager med jobnavn
                            select case lto 
                            case "bf", "intranet - local"
                            modtageradr = "To:<br>"& strJobnavn & " ("& intjobnr &")" 
                            skiftModtagerDis = "DISABLED"
                            case else 
                            modtageradr = modtageradr 
                            skiftModtagerDis = ""
                            end select

                           	                
		                %>
		                <tr>
		                <td valign=top style="height:110px;" align=left>
		                <input type="hidden" id="momsland" value="<%=strLandShow %>" />
                        <input type="hidden" id="afsmomsland" value="<%=thisAfsenderLand %>" />
                        <div id="DIV_modtageradr" style="width:310px; height:100px; padding:4px; border:1px #8CAAE6 solid; overflow:auto;">
                           <%=modtageradr%> 
                            </div>
                            
                            <textarea style="position:absolute; visibility:hidden; width:310px; height:101px; padding:4px; border:1px #CCCCCC solid; top:30px; left:5px;" id="FM_modtageradr" name="FM_modtageradr">
                            <%=modtageradr%> 
                            </textarea>   
                            <span style="font-size:9px; color:#999999; line-height:10px;"><%=erp_txt_217 %></span>
                            <input type="button" id="gem_adr" style="position:absolute; display:none; visibility:hidden; top:30px; left:320px;" value="luk" />
                            
                            
                            <img src="../ill/blyant.gif" id="red_adr" border="0" style="position:absolute; display:; cursor:hand; visibility:visible; top:30px; left:320px;"/>
                            
		                 <input type="hidden" name="FM_Kid" id="FM_kid" value="<%=kid%>">
                		
		               </td>
		                </tr>
		               <tr>
		               <td><input type="checkbox" name="FM_land_vis" id="FM_land_vis" value="on" <%=lndCHK %>><%=erp_txt_218 %>
		               <br />
		               <input type="checkbox" name="FM_cvr_vis" id="FM_cvr_vis" value="on" <%=intModCvrCHK %>><%=erp_txt_219 %></td>	
		                </tr>
		                  <tr>	


                        
                        <%
                        
                        '*** Er der forvalgt en fastfakadr under kontaktpersoner / filialer p� denne kunde'
                        '*** G�lder b�de job & aftaler
                        forvalgtFakadr = 0
                        if func <> "red" AND cint(usealtadr) = 0 then
                        strSQLkpersFakadrfv = "SELECT id, navn FROM kontaktpers WHERE kundeid = "& kid &" AND kptype = 1 ORDER BY navn"
				        oRec2.open strSQLkpersFakadrfv, oConn, 3
                        
                        if not oRec2.EOF then

                        forvalgtFakadr = oRec2("id")
                      
                        end if
                        oRec2.close
                        end if

                        
                                
                                if usealtadr <> 0 then
				                usealtadrCHK = "CHECKED"
                                    

                                    '** hvis alt fakadr er valgt p� job s�ttes usealtadr = kontakpersonId pp� jobbet 
                                    if func <> "red" then
                                    usealtadr = strAtt 
                                    end if


				                else
                                    
                                    '** Er fast faktura asdr valgt p� kunde
                                    if cint(forvalgtFakadr) <> 0 AND func <> "red" then
                                    usealtadrCHK = "CHECKED"
                                    usealtadr = forvalgtFakadr 
                                    else
				                    usealtadrCHK = ""
				                    end if

                                end if %>

		                
                        <%
                        '** Att person skal kun v�re sl�et til hvis ikke alternativ fakadr er forvalgt p� job
                        '** eller hvis rediger
                        if (func <> "red" AND (cint(usealtadr) = 0 OR cint(forvalgtFakadr) <> 0) AND lto <> "hestia") OR (func = "red" AND cint(strAtt) <> 0) then
		                attCHK = "CHECKED"
		                else
		                attCHK = ""
		                end if
		                 %>
			                <td valign=top>
				              
				                
				                
                                
                                <br /><b><%=erp_txt_060 %></b> <%=erp_txt_061 %><br />
                                <span style="font-size:9px; color:#999999;"><%=erp_txt_062 %></span><br /> 
                                <input type="hidden" value="<%=func %>" id="altadrFunc" />
                                <input id="FM_usealtadr" name="FM_usealtadr" type="checkbox" value="1" <%=skiftModtagerDis %> <%=usealtadrCHK %> /> Ja 
                                <%if func = "red" then %> 
                                <%=erp_txt_063 %><span style="font-size:9px; color:#999999;"><%=erp_txt_064 %></span> 
                                <%end if %><br />
                                <select name="FM_altadr" id="FM_altadr" style="width:150px; font-size:9px;">
				                <%
				                strSQLkpersCount = "SELECT id, navn, titel FROM kontaktpers WHERE kundeid = "& kid &" ORDER BY navn"
				                oRec2.open strSQLkpersCount, oConn, 3
                					
					                while not oRec2.EOF


                                        if len(trim(oRec2("titel"))) <> 0 then
                                        titelTxt = " ("& oRec2("titel") &")"
                                        else
                                        titelTxt = ""
                                        end if

						                if cint(usealtadr) = cint(oRec2("id")) then
						                attSel = "SELECTED"
						                else
						                attSel = ""
						                end if%>
						                <option value="<%=oRec2("id")%>" <%=attSel%>><%=oRec2("navn")%> <%=titelTxt %></option>
						                <%
					                oRec2.movenext
					                wend
					                oRec2.close
					                %>
				                
				                </select>&nbsp;<input id="nyekpers" type="button" style="font-size:9px; font-family:arial;" value="Hent nye >>" />&nbsp;
                                <a href="#" onclick="Javascript:window.open('../to_2015/kontaktpers.asp?menu=kund&kid=<%=kid%>&func=opr&rdir=fak', '', 'width=804,height=800,resizable=yes,scrollbars=yes')" class="rmenu"><%=erp_txt_220 %></a> 

                                
				                </td>
				                </tr>
				                <tr>
				                <td>
				              
                              
                              
                                <input type="checkbox" name="FM_att_vis" id="FM_att_vis" value="on" <%=attCHK %>><b><%=erp_txt_065 %></b> <%=erp_txt_066 %>&nbsp;
			                <br /><select name="FM_att" id="FM_att" style="width:150px; font-size:9px;">
				                <%
				                strSQLkpersCount = "SELECT id, navn FROM kontaktpers WHERE kundeid = "& kid &" ORDER BY navn"
				                oRec2.open strSQLkpersCount, oConn, 3
                					
					                while not oRec2.EOF
						                if cint(strAtt) = cint(oRec2("id")) then
						                attSel = "SELECTED"
						                else
						                attSel = ""
						                end if%>
						                <option value="<%=oRec2("id")%>" <%=attSel%>><%=oRec2("navn")%></option>
						                <%
					                oRec2.movenext
					                wend
					                oRec2.close
					                %>
				                <%if strAtt = "991" then
				                selth1 = "SELECTED"
				                else
				                selth1 = ""
				                end if%>	
				                <option value="991" <%=selth1%>><%=erp_txt_067 %></option>
				                <%if strAtt = "992" then
				                selth2 = "SELECTED"
				                else
				                selth2 = ""
				                end if%>	
				                <option value="992" <%=selth2%>><%=erp_txt_068 %></option>
				                <%if strAtt = "993" then
				                selth3 = "SELECTED"
				                else
				                selth3 = ""
				                end if%>	
				                <option value="993" <%=selth3%>><%=erp_txt_069 %></option>
				                </select>
				                </td>
				                </tr>
		                </table>
		              
                  <!-- Modtager slut -->
   
	                </td><td valign=top width="50%">
	
                        
	
	
	<!--- Afsender af faktura --->
	<table cellspacing="0" cellpadding="5" border="0" width="100%" bgcolor="#FFFFFF">
	<tr>
	    <td style="border:1px #ffffff solid;" bgcolor="#8caae6" class=alt><b><%=erp_txt_070 %> </b></td>
	</tr>
	<tr>
	    <td >
	
		<%
        if func = "red" then
		afskidSQLkri = "kid = " & afsender
        else

            
            if cint(multible_licensindehavere) = 1 AND jobid <> 0 then 
            afskidSQLkri = "useasfak = 1 AND lincensindehaver_faknr_prioritet = "& lincensindehaver_faknr_prioritet_job &""
            else
            afskidSQLkri = "useasfak = 1"
            end if

        end if

		
		strSQL = "SELECT adresse, postnr, city, land, telefon, fax, email, regnr, kkundenavn, kontonr, cvr, bank, swift, iban, kid FROM kunder WHERE " & afskidSQLkri
		
            'if session("mid") = 1 then
            'Response.write "SQL: " & strSQL
            'end if    
        oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			yourbank = oRec("bank")
			yourRegnr = oRec("regnr")
			yourKontonr = oRec("kontonr")
			yourCVR = oRec("cvr")
			yourNavn = oRec("kkundenavn")
			yourAdr = oRec("adresse")
			yourPostnr = oRec("postnr")
			yourCity = oRec("city")
			yourLand = oRec("land")
			yourEmail = oRec("email")
			yourTlf = oRec("telefon")
			yourSwift = oRec("swift")
			yourIban = oRec("iban")
			afskid = oRec("kid")
			yourFax = oRec("fax")
		end if
		oRec.close
		
		
		%>
        <!--
		  <div id="DIV1" style="width:300px; height:100px; padding:4px; border:1px #8CAAE6 solid; overflow:auto;">
                     
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
		<tr>
			<td><%=yourNavn%></td>
		</tr>
		<tr>
			<td><%=yourAdr%></td>
		</tr>
		<tr>
			<td><%=yourPostnr%>&nbsp;&nbsp;<%=yourCity%></td>
		</tr>
		<tr>
			<td><%=yourLand%></td>
		</tr>
		</table>
		</div>
        -->
		
		<table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr>
		    <td>
		    <select id="FM_afsender" name="FM_afsender" style="width:300px; font-size:10px;" size=8">
		    
		    <%strSQLaltaf = "SELECT kid, kkundenavn, kkundenr, adresse, postnr, city, land, telefon, cvr, swift, iban, regnr, kontonr, regnr_b, kontonr_b, regnr_c, kontonr_c FROM kunder WHERE useasfak = 1 ORDER BY kkundenavn "
             'OR useasfak = 2 Fjernet 20170117 PGA muli licensindehavere og dermed flere fakturanr. R�kkef�lger 
             'Selskab, licensejer eller datter selskab **'
		    oRec.open strSQLaltaf, oConn, 3 
		    
            'kidSel = 0
            while not oRec.EOF  
		    
            if afskid = oRec("kid") then
            afSEL = "SELECTED"
            'kidSel = oRec("kid")
            else
            afSEL = ""
            end if

		   %>
		    <option value="<%=oRec("kid") %>" <%=afSEL %>><%=oRec("kkundenavn") %> (<%=oRec("kkundenr") %>)</option>
            <option disabled value="<%=oRec("kid") %>"><%=oRec("adresse") &" "& oRec("postnr") &" "& oRec("city") &" "& oRec("land") %></option>
            <option disabled value="<%=oRec("kid") %>"><%="Tel: "& oRec("telefon") &" CVR: "& oRec("cvr") & " SWIFT: "& oRec("swift") & " IBAN: "& oRec("iban")%> </option>
            <option disabled value="<%=oRec("kid") %>"></option>

		    <%
		    oRec.movenext
		    wend
		    oRec.close%>
		    </select>
		    </td>
		</tr>

		<tr>
		    <td>
		    <br /><b>Vor Ref.:</b><br /><select id="FM_vorref" name="FM_vorref" style="width:300px; font-size:10px;">
		    <option value="-">Ingen</option>
		    <%strSQlvorref = "SELECT mnavn, mid FROM medarbejdere WHERE mansat <> 2 AND mansat <> 3 ORDER BY mnavn "
		    oRec.open strSQlvorref, oConn, 3 
		    while not oRec.EOF  
		    
		    if ( oRec("mid") = cint(jobans1) AND func <> "red" AND (instr(lto, "epi") <> 0) _
            OR ( oRec("mid") = cint(session("mid") ) AND func <> "red" AND (lto <> "epi" AND lto <> "epi_no" AND lto <> "epi2017" AND lto <> "synergi1") OR (vorref = oRec("mnavn") AND func = "red"))) _
            then 'OR (lto = "jttek" AND func = "red") 
            vfSEL = "SELECTED"
		    else
		    vfSEL = ""
		    end if  
		    
		    if cint(jobans1) = oRec("mid") then
		    vrjobans1 = " - (jobansvarlig)"
		    else
		    vrjobans1 = ""
		    end if
		    
		    
		    if cint(jobans2) = oRec("mid") then
		    vrjobans2 = " - (jobejer)"
		    else
		    vrjobans2 = ""
		    end if
		    %>
		    
		    
		    <option value="<%=oRec("mnavn") %>" <%=vfSEL %>><%=oRec("mnavn") %> <%=vrjobans1 %> <%=vrjobans2 %></option>
		    <%
		    oRec.movenext
		    wend
		    oRec.close%>
		    </select>
		    </td>
		</tr>
		<tr>
			<td><br /><br /><b><%=erp_txt_071 %></b><br /><input type="checkbox" name="FM_ytlf_vis" value="on" <%=visAfsTlfCHK %>><%=erp_txt_003 %></td>
		</tr>
		<tr>
			<td><input type="checkbox" name="FM_yfax_vis" value="on" <%=visAfsFaxCHK %>><%=erp_txt_023 %></td>
		</tr>
		<tr>
		    <td><input type="checkbox" name="FM_yemail_vis" value="on" <%=visafsemailCHK %>><%=erp_txt_024 %></td>
		</tr>
	
       

		<tr>
			<td><input type="checkbox" name="FM_yswift_vis" value="on" <%=visAfsSwiftCHK %>><%=erp_txt_010 %></td>
		</tr>
		<tr>
			<td><input type="checkbox" name="FM_yiban_vis" value="on" <%=visAfsIbanCHK %>><%=erp_txt_009 %></td>
		</tr>
		<tr>
			<td><input type="checkbox" name="FM_ycvr_vis" value="on" <%=visAfsCvrCHK %>><%=erp_txt_004 %></td>
		</tr>
        <tr>
        <%
        if len(trim(afskid)) <> 0 then
            afskid = afskid
        else
            afskid = 0
        end if

        strSQLKonto = "SELECT kid, bank, regnr, kontonr, bank_b, regnr_b, kontonr_b, bank_c, regnr_c, kontonr_c, "_
        &" bank_d, regnr_d, kontonr_d, bank_e, regnr_e, kontonr_e, bank_f, regnr_f, kontonr_f FROM kunder WHERE kid = "& afskid 'Selskab, licensejer eller datter selskab **'
		
        'response.write "strSQLKonto" & strSQLKonto

        oRec.open strSQLKonto, oConn, 3 
		if not oRec.EOF then
            

            bank = oRec("bank")
            regnr = oRec("regnr")
            kontonr = oRec("kontonr")

            bank_b = oRec("bank_b")
            regnr_b = oRec("regnr_b")
            kontonr_b = oRec("kontonr_b")

            bank_c = oRec("bank_c")
            regnr_c = oRec("regnr_c")
            kontonr_c = oRec("kontonr_c")
            
            bank_d = oRec("bank_d")
            regnr_d = oRec("regnr_d")
            kontonr_d = oRec("kontonr_d")

            bank_e = oRec("bank_e")
            regnr_e = oRec("regnr_e")
            kontonr_e = oRec("kontonr_e")

            bank_f = oRec("bank_f")
            regnr_f = oRec("regnr_f")
            kontonr_f = oRec("kontonr_f")
            
        end if
        oRec.close  


        kontonr_sel0 = ""
        kontonr_sel1 = ""
        kontonr_sel2 = ""
        kontonr_sel3 = ""
        kontonr_sel4 = ""
        kontonr_sel5 = ""

        select case kontonr_sel
        case 0
        kontonr_sel0 = "SELECTED"
        case 1
        kontonr_sel1 = "SELECTED"
        case 2
        kontonr_sel2 = "SELECTED"
        case 3
        kontonr_sel3 = "SELECTED"
        case 4
        kontonr_sel4 = "SELECTED"
        case 5
        kontonr_sel5 = "SELECTED"
        case else
        kontonr_sel0 = "SELECTED"
        end select 

         %>

			<td><br /><%=erp_txt_424 %>:<br /><select name="FM_afs_bankkonto" id="FM_afs_bankkonto" style="width:300px;">
            <option value="0" <%=kontonr_sel0 %>><%=bank &": "& regnr &" "& kontonr %></option>
            <option value="1" <%=kontonr_sel1 %>><%=bank_b &": "& regnr_b &" "& kontonr_b %></option>
            <option value="2" <%=kontonr_sel2 %>><%=bank_c &": "& regnr_c &" "& kontonr_c %></option>
            <option value="3" <%=kontonr_sel3 %>><%=bank_d &": "& regnr_d &" "& kontonr_d %></option>
            <option value="4" <%=kontonr_sel4 %>><%=bank_e &": "& regnr_e &" "& kontonr_e %></option>
            <option value="5" <%=kontonr_sel5 %>><%=bank_f &": "& regnr_f &" "& kontonr_f %></option>
            </select></td>
		</tr>

		</table>
        <br /><br /><br />&nbsp;
		
		</td>
		</tr>
		
	
	</table>
	
	</td></tr>
	</table>
	</div>
	
	<%if jobid <> 0 then
	nst = "jobbesk"
	else
	nst = "aktdiv"
	end if %>
	
	    <div id=modtagdiv_2 style="position:absolute; visibility:hidden; display:none; top:586px; width:720px; left:5px; border:0px #8cAAe6 solid;">
        <table width=100% cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('fidiv')" class=vmenu><< <%=erp_txt_072 %></a></td><td align=right><a href="#" onclick="showdiv('<%=nst %>')" class=vmenu><%=erp_txt_073 %> >></a></td></tr></table>
		 </div>

	
	
	<!-- afsender slut -->
	
	
	
	        <%'** Fakdato MSG ****'
            itop = 140
            ileft = 435
            iwdt = 250
            idsp = "none"
            ivzb = "hidden"
            iId = "fakdatoinfo"
            call sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb)
            %>
            <%=erp_txt_074 %> <b><%=erp_txt_075 %></b><%=erp_txt_076 %>
            <br /><br />
            <%=erp_txt_077 %> <b><%=erp_txt_078 %></b>
		    <br /><br /><%=erp_txt_079 %> <b><%=erp_txt_080 %></b> <%=erp_txt_081 %> 
		    <br />&nbsp;
	            </td></tr></table>
	            </div>
	        
            <input id="showfakmsg" value="0" type="hidden" />
	       
	
	
	
	
	<!-- Faktura indstillinger -->

	<div id="fidiv" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; top:105px; width:720px; left:5px; border:1px #8cAAE6 solid;">
    <table cellspacing="1" cellpadding="0" border="0" width=100% bgcolor="#ffffff">
	<tr>
	    <td colspan=2 bgcolor="#8caae6" class=alt style="border:1px #ffffff solid; padding:5px 0px 5px 5px;"><b><%=erp_txt_082 %></b>
	    
	    &nbsp;&nbsp;(Type: <%
		'** Faktype ***
		select case intType
		case "0"
		strFaktypeNavn = erp_txt_001
		case "1"
		strFaktypeNavn = erp_txt_002
		case "2"
		strFaktypeNavn = erp_txt_083
		case else
		strFaktypeNavn = erp_txt_001
		end select
		%>
			
			<%=strFaktypeNavn %>)
	    
	    </td>
	 </tr>
	<!-- Faktura dato, forfaldsdato, periode dato, labeldato -->
	<tr>
		<td style="width:140px; padding:30px 5px 2px 5px;">
		
			<% 
	        if len(lastFakdato) <> 0 then
	        lastFakdato = lastFakdato
	        else
	        lastFakdato = "2001/1/1"
	        end if
	
	
		'** Periodeinterval brugt.
		'** Start dato
		call datofindes(request("FM_start_dag"),request("FM_start_mrd"),request("FM_start_aar"))
		stdato = request("FM_start_aar")&"/"&request("FM_start_mrd")&"/"&dagparset
		showStDato = dagparset&"/"&request("FM_start_mrd")&"/"&request("FM_start_aar")
		
		
		'** Slut dato
		call datofindes(request("FM_slut_dag"),request("FM_slut_mrd"),request("FM_slut_aar"))
		slutdato = request("FM_slut_aar")&"/"&request("FM_slut_mrd")&"/"&dagparset
		slutDagparset = dagparset
		showSlutDato = dagparset&"/"&request("FM_slut_mrd")&"/"&request("FM_slut_aar") 
		useLastFakdato = 0
	
	    
	    'Response.Write "showStDato: "& showStDato & "lastFakdato: "& lastFakdato
	    'Response.end
	    
	    '**** Hvis startdato ligger f�r sidste fakdato s�ttes stdato = sidste fakdato ****
	    if (cdate(showStDato) > cdate(lastFakdato) OR intType = 1) then 'kreditnota 'AND cint(usedt_ival) = 1 
		    stdatoKri = stdato
		    useLastFakdato = 0
	    else
		    stdatoKri_temp = dateadd("d", 1, lastFakdato)
		    stdatoKri = year(stdatoKri_temp)&"/"&month(stdatoKri_temp)&"/"&day(stdatoKri_temp)
		    showStDato = stdatoKri_temp 'lastFakdato
		    useLastFakdato = 1
	    end if
	    '******************************************************
	
	   
	
	   
	    %>  
	        <!-- jquery val for materiale fun --->
            <input id="jq_stdato" value="<%=stdatoKri%>" type="hidden" />
            <input id="jq_sldato" value="<%=slutdato%>" type="hidden" />
            <input type="hidden" id="jq_func" value="<%=func%>">
            <input type="hidden" id="jq_id" value="<%=id%>">
            <input type="hidden" id="jq_inttype" value="<%=intType%>">
		
		<b><%=erp_txt_084 %></b> <%=erp_txt_221 %><br />
        <span style="color:#999999;"><%=erp_txt_085 %></span>
       
		</td><td valign=top style="width:520px; padding:30px 5px 2px 5px;">
		
		<%'** Bruger afg. dato fra interval som faktura dato **'
		if func <> "red" then 'cint(usedt_ival) = 1 and
            select  case lto 
            case "epi", "epi_no", "epi_sta", "epi_ab", "xintranet - local", "epi2017"
            useFakDate = dateadd("m", -3, now)
            strDag = day(useFakDate)
		    strMrd = month(useFakDate)
		    strAar = year(useFakDate)
            
            case "intranet - local", "nt"

            if isDate(dt_actual_etd) = true then
            useFakDate = dt_actual_etd
            else
            useFakDate = now
            end if
            
            strDag = day(useFakDate)
		    strMrd = month(useFakDate)
		    strAar = year(useFakDate)
            case else
		    strDag = datepart("d", showSlutDato, 2, 3)
		    strMrd = datepart("m", showSlutDato, 2, 3)
		    strAar = datepart("yyyy", showSlutDato, 2, 3)
            end select
		else
		strDag = datepart("d", strTdato, 2, 3)
		strMrd = datepart("m", strTdato, 2, 3)
		strAar = datepart("yyyy", strTdato, 2, 3)


		end if
		
		
		
		'Response.write "<b>" & replace(formatdatetime(strDag &"/"& strMrd &"/"& strAar, 2),"-",".") & "</b>"
		
		'if lastFakdato <> "2001/1/1" then
		'Response.Write "&nbsp;<font class=lillesort><i>(seneste faktura dato: " & replace(formatdatetime(lastFakdato, 2),"-",".") &")</i></font>" 
		'end if 
		%>
		<input type="text" name="FM_fakdato" id="FM_fakdato" value="<%=strDag%>.<%=strMrd%>.<%=strAar%>" style="display:none; margin-right:5px; width:80px;" />

            	

            <%if func <> "red" AND useLastFakdato = 1 then %>
            <div id="sp_perstdato" style="position:absolute; left:300px; top:43px; color:#999999; width:250px; z-index:2000; background-color:#FFFFe1; padding:5px; border:1px #cccccc solid;"><%=erp_txt_049 %> <b><%=erp_txt_086 %></b> <%=erp_txt_087 %></div>
	       
		<% 
        end if %>
		
		

        <input name="FM_fakdato_dag" id="FM_fakdato_dag" type="hidden" value="<%=strDag%>" />
        <input name="FM_fakdato_mrd" id="FM_fakdato_mrd" type="hidden" value="<%=strMrd%>" />
        <input name="FM_fakdato_aar" id="FM_fakdato_aar" type="hidden" value="<%=strAar%>" />
		
		
		<input type="hidden" name="FM_interval_slutdag" id="FM_interval_slutdag" value="<%=slutDagparset%>">
		<input type="hidden" name="FM_interval_slutmrd" id="FM_interval_slutmrd" value="<%=request("FM_slut_mrd")%>">
		<input type="hidden" name="FM_interval_slutaar" id="FM_interval_slutaar" value="<%=request("FM_slut_aar")%>">
		<!--<input type="hidden" name="showfakmsg" id="showfakmsg" value="0">-->
		
	    </td>
		</tr>
		
		<%

            'response.End
		
		if func <> "red" then
		
		    if request.Cookies("erp")("visperiode") <> "" then
    	         
	             if request.Cookies("erp")("visperiode") = "1" then
		         chkvp = "CHECKED"
		         else
    		     chkvp = ""
    		     end if
    		
		    else
    		
		         chkvp = ""
    		 
		    end if

            select case lto 'Overruler
            case "bf"
            chkvp = "CHECKED"
            end select

		
		else
		     
		     if cint(visperiode) = 1 then
		     chkvp = "CHECKED"
		     else
    		 chkvp = ""
    		 end if
		
		end if
		
		
		%>
		
		
		<tr>
		<td valign=top style="padding:4px 5px 2px 5px;"><input id="Checkbox1" name="FM_visperiode" value="1" type="checkbox" <%=chkvp %> /><span><%=erp_txt_088 %> <b><%=erp_txt_049 %></b> <%=replace(erp_txt_089, ":", "") %></span>
		
		 
	    </td>
		
		
		
		<td valign=top style="padding:4px 5px 2px 5px;">
            
		<%if func <> "red" then
		istDato = showStDato
		istSlutDato = showSlutDato 
		else
		istDato = istDato
		istSlutDato = istDato2
		end if %>

		<!--#include file="inc/erp_istdato_inc.asp"-->&nbsp;&nbsp;<%=erp_txt_090 %>&nbsp;&nbsp;<!--#include file="inc/erp_istdato2_inc.asp"-->
		</td></tr>
		
		
		<%
		if func <> "red" then
		labelDato = now
		else
		labelDato = labelDato
	    end if
		
		
		
		
		if func <> "red" then
		    
		    if instr(lto, "epi") <> 0 then
		    chkflabel = "CHECKED"
		    
		    else
		
		        if request.Cookies("erp")("flabel") <> "" then
    	         
	                 if request.Cookies("erp")("flabel") = "1" then
		             chkflabel = "CHECKED"
		             else
    		         chkflabel = ""
    		         end if
    		
		        else
    		
		             chkflabel = ""
    		 
		        end if
		    
		    
		    end if
		
		else
		     
		     if cint(brugfakdatolabel) = 1 then
		     chkflabel = "CHECKED"
		     else
    		 chkflabel = ""
    		 end if
		
		end if %>
		
		<tr><td valign=top style="padding:2px 5px 2px 5px;"><input id="Checkbox2" name="FM_brugfakdatolabel" value="1" type="checkbox" <%=chkflabel %> /><b><%=erp_txt_091 %></b>

            <%
                qmtxt = erp_txt_98 & "<br><br>" & erp_txt_99 & "<br><br>" & erp_txt_100 & "<b>" & erp_txt_101 & "</b>" & erp_txt_102 
                qmtop = 50
                qmleft = 100
                qmwdt = 300
                qmid = "qmFA02"
                
            %>
            <span class="qmarkhelp" id="<%=qmid %>" style="font-size:11px; color:#999999;"><u>?</u></span>
            <%
               
                call qmarkhelpnote(qmtxt,qmtop,qmleft,qmid,qmwdt) %>


		
				
		
        
		
		</td><td valign=top style="padding:4px 5px 2px 5px;"><!--#include file="inc/erp_labeldato_inc.asp"--></td></tr>
		
		
		<% 
		if func = "red" then
	    strB_dato = strB_dato
	    adddayVal = dateDiff("d",strDag&"/"&strMrd&"/"&strAar, strB_dato,2,2)
	    else
	    end if
	    
	    %>
	    
	    <tr><td valign=top height=20 style="padding:7px 5px 0px 5px;">
		<b><%=erp_txt_092 %></b></td><td valign=top style="padding:2px 5px 0px 5px;">
	    <%
	    
	    
	    '*** Kun admin m� rette i forfaldsdato **'
	    if level <> 1 then
	    
                if instr(lto, "epi") <> 0 OR instr(lto, "intra") <> 0  then 
	            hideffdato = 0 '1
                else
                hideffdato = 0
                end if
	
	    else
            hideffdato = 0
        end if
	    
            select case lto
            case "nt", "epi_uk", "intranet - local", "bf" 
            lang = 1
            case else
            lang = 0
            end select
	     
            nameid = "FM_betbetint"
            'betbetint = 200

            select case lto
            case "nt", "xintranet - local"
                
                if func <> "red" then
                    if cint(fastpris) = 2 then '** COMI
                    betbetint = lev_betbetint
                    else
                    betbetint = kunde_betbetint
                    end if
                else
                    betbetint = betbetint
                end if

            case else
            betbetint = betbetint
            end select

            'response.write "betbetint: "& betbetint

	        call betalingsbetDage(betbetint, hideffdato, lang, nameid)
	        if Not InStr(strForfaldsdato, "-") then
	        strForfaldsdato = strDag & "-" & strMrd & "-" & strAar
	        end if
    	    
	        'Response.write "strForfaldsdato: "& strForfaldsdato
    	    
    	    
	        reformatfordate = split(strForfaldsdato, "-")
	    
	   	
	  	'*** Forfaldsdato **************
	  	'if hideffdato <> 1 then%>
	  	<input type="text" id="FM_forfaldsdato" name="FM_forfaldsdato" value="<%=reformatfordate(0)%>.<%=reformatfordate(1)%>.<%=reformatfordate(2)%>" style="margin-right:5px; width:80px;" DISABLED/>
        <input type="hidden" value="<%=hideffdato%>" id="hideffdato"/> 
        <%'else %>
       <!-- <input type="hidden" name="FM_forfaldsdato" value="<%=reformatfordate(0)%>.<%=reformatfordate(1)%>.<%=reformatfordate(2)%>" style="margin-right:5px; width:70px;"/> 
            <div style="position:relative; top:5px; padding:2px;"><%=reformatfordate(0)%>.<%=reformatfordate(1)%>.<%=reformatfordate(2)%></div>-->
        <%'end if %>

            <input type="hidden" id="transport" name="" value="<%=transportVal %>" />

        </td>
		</tr>

        <%'if func <> "red" AND (lto = "dencker") then
        'klTjk = ""
        'gkTjk = "CHECKED"
        'else

        '** Hvis �ben for Rediger, vil status altid v�re kladde
        klTjk = "CHECKED"
        gkTjk = ""
        'end if %>
		
		<tr><td valign=top height=20 style="padding:40px 5px 2px 5px;">
		<b><%=erp_txt_093 %></b></td><td valign=top style="padding:40px 5px 2px 5px;">
				<input type="radio" name="FM_betalt" value="0" <%=klTjk %>><%=erp_txt_094 %>&nbsp;&nbsp;&nbsp;

                <%select case lto
                 case "epi2017"
                    if cint(level) = 1 then
                    approveInvoiceOk = 1
                    else
                    approveInvoiceOk = 0
                    end if

                 case else
                    approveInvoiceOk = 1
                 end select  
                    
                if cint(approveInvoiceOk) = 0 then
                    approveInvoiceOkDis = "DISABLED"
                else
                    approveInvoiceOkDis = ""
                end if
                    %>

				<input type="radio" name="FM_betalt" value="1" <%=gkTjk %> <%=approveInvoiceOkDis %> /><span style="color:yellowgreen;"><b><%=erp_txt_095 %></b></span>
                

		
		<input id="FM_type" name="FM_type" value="<%=intType%>" type="hidden" />
		
		</td>
		</tr>
		
		<tr>
		<td style="padding:4px 5px 0px 5px;"><b><%=erp_txt_096 %></b>
         </td>
		<td style="padding:4px 5px 0px 5px;">
           
           <%
								 if cint(visrabatkol) <> 0 then
								 rbtCHK = "CHECKED"
								 else
								 tbtCHK = ""
								 end if
								 %>
								 
                                <input id="FM_visrabatkol" name="FM_visrabatkol" type="checkbox" <%=rbtCHK %>>&nbsp;<%=erp_txt_097 %>
            
                 
           </td>
		</tr>

     


        <%if intType <> 1 then '1 = kreditnota (En kreditnota kan IKKE v�re intern.) %>
         <tr>
		<td style="padding:12px 5px 20px 5px;" valign=top><b><%=erp_txt_105 %></b>

               <%
                qmtxt = erp_txt_109 & "<br><br>" & erp_txt_110 & "<br><br>" & erp_txt_111 & "<br><br>" & erp_txt_112
                qmtop = 250
                qmleft = 100
                qmwdt = 300
                qmid = "qmFA01"
                
            %>
            <span class="qmarkhelp" id="<%=qmid %>" style="font-size:11px; color:#999999;"><u>?</u></span>
            <%
               
                call qmarkhelpnote(qmtxt,qmtop,qmleft,qmid,qmwdt) %>

         </td>
		<td style="padding:8px 25px 20px 5px;" valign=top>
           
           <%
								 'if cint(medregnikkeioms) <> 0 then
								 'medregnikkeiomsCHK = "CHECKED"
								 'else
								 'medregnikkeiomsCHK = ""
								 'end if
								 %>
								
                                <input id="Hidden3" type="hidden" name="FM_medregnikkeioms_opr" value="<%=medregnikkeioms %>" />

                                <%if cint(fak_laast) = 0 then
                                intDisab = ""
                                else
                                intDisab = "DISABLED"
                                %>
                               
                               <!--  <input id="Hidden5" type="hidden" name="FM_medregnikkeioms" value="<%=medregnikkeioms %>" /> -->
                                <%
                                end if


                                    medregnikkeiomsSEL0 = ""
                                    medregnikkeiomsSEL1 = ""
                                    medregnikkeiomsSEL2 = ""
                                  select case medregnikkeioms
                                   case 1
                                    medregnikkeiomsSEL1 = "SELECTED"
                                    case 2
                                    medregnikkeiomsSEL2 = "SELECTED"
                                    case else
                                    medregnikkeiomsSEL0 = "SELECTED"
                                    end select
                                 %>

                                <select style="width:150px;" id="FM_medregnikkeioms" name="FM_medregnikkeioms">

                                    <option value="0" <%=medregnikkeiomsSEL0 %>><%=erp_txt_106 %></option>
                                     <option value="1" <%=medregnikkeiomsSEL1 %>><%=erp_txt_107 %></option>
                                     <option value="2" <%=medregnikkeiomsSEL2 %>><%=erp_txt_108 %></option>

                                </select>

                                <!--<input id="Checkbox3" name="FM_medregnikkeioms" <%=intDisab%> value="1" type="checkbox" <%=medregnikkeiomsCHK %>>Dette er en intern faktura. --><br />
                          
                                
            
                 
           </td>
		</tr>
        <%end if %>

     
		<tr><td valign=top style="padding:16px 5px 0px 5px;">
        
       
		<b><%=erp_txt_113 %></b>
            <br /> <span style="color:#999999;"><%=erp_txt_288 %></span>
		
		</td>
		<td style="padding:26px 5px 0px 5px;">
		
		<%
		
		select case jobfaktype
		case 0
		jftp = 0
		case 1
		jftp = 1
		case else
		jftp = 0
		end select
		
		
		
		call selectAllValuta(1, jftp) 
		
		%>
							
            
		</td></tr>
		<tr><td style="padding:4px 5px 0px 5px;"><b><%=erp_txt_114 %></b>
		
		</td>
		<td style="padding:2px 5px 0px 5px;">
							
                            <select name="FM_sprog" style="width:70px;">
							<%
                            strSQL = "SELECT id, navn FROM fak_sprog WHERE id <> 0 ORDER BY id " 
							oRec.open strSQL, oConn, 3
							while not oRec.EOF 
							
                             if oRec("id") = cint(sprog) then
							 sSEL = "SELECTED"
							 else
							 sSEL = ""
							 end if
                            
                            %>
							<option value="<%=oRec("id") %>" <%=sSEL %>><%=oRec("navn") %> </option>
							<%
							oRec.movenext
							wend
							oRec.close%>
							
							</select>
		</td></tr>
		
		<%
        '*** OMSKRIV til at HENTE fra fak_moms, tabellen er klar 20150317

        msSEL0 = ""
        msSEL08 = ""
        msSEL14 = ""
        msSEL20 = ""
		msSEL22 = ""
        msSEL25 = ""

		select case cint(momssats)
        case 8 
		msSEL08 = "SELECTED"
        case 14 
		msSEL14 = "SELECTED"
        case 20 
		msSEL20 = "SELECTED"
        case 22 
		msSEL22 = "SELECTED"
        case 25 
		msSEL25 = "SELECTED"
		case else
		msSEL0 = "SELECTED"
		end select %>
		
		
		<tr><td style="padding:2px 5px 4px 5px;"><b><%=erp_txt_115 %></b></td>
		<td style="padding:2px 5px 4px 5px;">
            <select id="FM_momssats" name="FM_momssats" style="width:70px;">
                <option value="25" <%=msSEL25 %>>25%</option>
                <option value="8" <%=msSEL08 %>>8%</option>
                <option value="14" <%=msSEL14 %>>14%</option>
                <option value="20" <%=msSEL20 %>>20%</option>
                <option value="22" <%=msSEL22 %>>22%</option>
                <option value="0" <%=msSEL0 %>>0%</option>
            </select> 
          

		</td>
		</tr>
		
		<%if len(trim(sideskiftH)) <> 0 then 
		sideskiftH = sideskiftH
		else
		sideskiftH = 0
		end if%>
		
		<tr><td style="padding:20px 5px 4px 5px;"><b><%=erp_txt_116 %></b></td>
		<td style="padding:20px 5px 4px 5px;">
            <select name="FM_sideskiftlinier" id="Select1" style="width:300px;">
            <%for l = 0 to 37 ' MAKS 8 sider / 34 = 5 sider %>
            
            
            
            
            <%if cint(sideskiftlinier) = l then
            sideskiftHSEL = "SELECTED"
            else
            sideskiftHSEL = ""
            end if
                
                if l <> 0 then
                    select case l 
                    case 31 
                    %>
                    <option value="<%=l %>" <%=sideskiftHSEL %>> <%=erp_txt_117 %></option>
                    <%
                    case 32 
                    %>
                    <option value="<%=l %>" <%=sideskiftHSEL %>> <%=erp_txt_118 %></option>
                    <%
                     case 33
                    %>
                    <option value="<%=l %>" <%=sideskiftHSEL %>> <%=erp_txt_119 %></option>
                     <%
                     case 34 
                    %>
                    <option value="<%=l %>" <%=sideskiftHSEL %>> <%=erp_txt_120 %></option>
                    <%
                    case 35 
                    %>
                    <option value="<%=l %>" <%=sideskiftHSEL %>> <%=erp_txt_121 %></option>
                    <%
                    case 36 
                    %>
                    <option value="<%=l %>" <%=sideskiftHSEL %>> <%=erp_txt_122 %></option>
                    <%
                    case 37 
                    %>
                    <option value="<%=l %>" <%=sideskiftHSEL %>> <%=erp_txt_123 %></option>
                     
                    <%case else%>
                    <option value="<%=l %>" <%=sideskiftHSEL %>><%=l &" "& erp_txt_124 %></option>
                <%  end select
                
                else %>
                <option value="<%=l %>" <%=sideskiftHSEL %>><%=l &" "& erp_txt_125 %></option>
                <%end if %>
                
            <%next %>
           
            </select>
            <br />&nbsp;
       
           </td>
		</tr>
		
    </table>
    </div>


<!-- KONTI KREDITOR (kunde) DEBITOR (Intern) -->
<div id="kontodiv" style="position:absolute; display:<%=sideDivDsp%>; visibility:<%=sideDivVzb%>; width:720px; top:650px; left:5px; border:1px #8cAAE6 solid; z-index:20000000000; background-color:#FFFFFF;">
    <table cellspacing="0" cellpadding="5" border="0" width=100% bgcolor="#FFFFFF">

        <tr><td><br />
    <b><%=erp_txt_289 %>:</b><br />
    <!--<span style="color:#999999;"><%=erp_txt_290 %></span>-->

    

    <%strSQLfomrALL = "SELECT f.navn AS fnavn, f.id, f.konto, kp.kontonr AS kkontonr, kp.navn AS kontonavn FROM fomr AS f "_
    &" LEFT JOIN kontoplan AS kp ON (kp.id = f.konto) WHERE f.id <> 0 ORDER BY f.navn"
        
        %>
    <select id="fak_fomr" name="FM_fak_fomr" style="width:400px;">

        <%
            
         fomrkonto = 0
         oRec2.open strSQLfomrAll, oConn, 3
         while not oRec2.EOF 
            
            if cint(fak_fomr) = oRec2("id") then
            fomrSEL = "SELECTED"
                    
                     if oRec2("konto") <> 0 then
                     fomrkonto = oRec2("konto")
                     else
                     fomrkonto = 0
                     end if

            else
            fomrSEL = ""
            end if

            if oRec2("konto") <> 0 then
            kontonrVal = " ("& left(oRec2("kontonavn"), 10) &" "& oRec2("kkontonr") &")"
            'fomrkonto = oRec2("konto")
            else
            kontonrVal = ""
            end if 
         
            %><option value="<%=oRec2("id") %>" <%=fomrSEL %>><%=oRec2("fnavn") & kontonrVal%></option><%

         oRec2.movenext
         wend
         oRec2.close  %>

        <option value="0"><%=erp_txt_423 %></option>


    </select>
    
            </td></tr>
    </table>



 <table cellspacing=0 cellpadding=5 border=0 width=100% bgcolor="#ffffff">
 
     <!--
	<tr><td valign=top style="padding:12px 5px 2px 5px;"><b><%=erp_txt_128 %></b><br />
	 <span style="color:#999999;"><%=erp_txt_129 %></span></td></tr>-->
	<tr>
	<td valign=top style="padding:12px 5px 2px 5px;"><b><%=erp_txt_130 %></b> (<%=erp_txt_414 %>)<br />
	
		<%
		if func = "red" then
		debKontouse = intKonto
		else
		debKontouse = kid
		end if
		
		select case lto
		case "execon", "immenso"
		disa = "DISABLED"
		case else
		disa = ""
		end select
		
		%>
	    <select name="FM_kundekonto" <%=disa %> style="width:325px;">
		<option value="0">(0)&nbsp;&nbsp;<%=erp_txt_131 %></option>
		<%
			strSQL = "SELECT k.kontonr, k.navn, k.id, k.kid, m.navn AS momskode FROM kontoplan k "_
			&" LEFT JOIN momskoder m ON (m.id = k.momskode) "_
			&" ORDER BY k.kontonr, k.navn"
			oRec.open strSQL, oConn, 3 
			fod = 0
			while not oRec.EOF
			
				if (debKontouse = oRec("kid") AND func <> "red") OR _
				(debKontouse = oRec("kontonr") AND func = "red") then
				selkon = "SELECTED"
				fod = 1
				else
				selkon = ""
				end if
		
			
			%>
			<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;&nbsp;<%=oRec("navn")%> - <%=oRec("momskode") %></option>
			<%
			oRec.movenext
			Wend 
			oRec.close
		
		
		if fod = 0 then
		    
		    if request.Cookies("erp")("debkonto") <> "" then
		    debkontonr = request.Cookies("erp")("debkonto")
		    
		    strSQL = "SELECT k.kontonr, k.navn, k.id, k.kid, m.navn AS momskode FROM kontoplan k "_
			&" LEFT JOIN momskoder m ON (m.id = k.momskode) "_
			&" WHERE kontonr = " & debkontonr &" ORDER BY k.kontonr, k.navn"
			
		    oRec.open strSQL, oConn, 3
		    if not oRec.EOF then
		    debkontonavn = oRec("navn")
		    debmomskode = oRec("momskode")
		    end if
		    oRec.close
		    
		    %>
		    <option value="<%=debkontonr%>">(<%=debkontonr%>)&nbsp;&nbsp;<%=debkontonavn%> - <%=debmomskode %></option>
			
		    <%
		    end if 
		
		end if%>
		</select><br />
        &nbsp;
		
		
       
		</td>
		<td valign=top style="padding:12px 5px 2px 5px;">
		<b><%=erp_txt_132 %></b> <%=erp_txt_133 %><br />
     
		
		<%

      
        

		fok = 0
				
				strSQL = "SELECT k.kontonr, k.navn, k.id, k.kid, m.navn AS momskode FROM kontoplan k "_
			    &" LEFT JOIN momskoder m ON (m.id = k.momskode) "_
			    &" ORDER BY k.kontonr, k.navn"
				
				%>
				<select name="FM_modkonto" <%=disa %> style="width:325px;">
		        <option value="0">(0)&nbsp;&nbsp;<%=erp_txt_131 %></option>
				<%
				oRec.open strSQL, oConn, 3 
				while not oRec.EOF 

    
                '*** Konto p� forretningsomr�de g�r f�r oms�tningskonto p� kunde f.eks m/uden moms mv. (egen kunde) ***'
                '*** preValgtModKonto Bruge ikke / eller kun af NT?? 20161215
				if ( ( ( (intModKonto = oRec("kid") OR cint(preValgtModKonto) = oRec("id")) AND fomrkonto = 0)_
                OR (cint(fomrkonto) = oRec("id"))) AND func <> "red")_
                OR ((intModKonto = oRec("kontonr") AND func = "red")) then
				selkon = "SELECTED"
				fok = 1
				else
				selkon = ""
				end if



				%>
				<option value="<%=oRec("kontonr")%>" <%=selkon%>>(<%=oRec("kontonr")%>)&nbsp;&nbsp;<%=oRec("navn")%> - <%=oRec("momskode") %></option>
                     <!--/Modk: <%=intModKonto %> /Kid <%=oRec("kid") & " /Prevglt: " & cint(preValgtModKonto) &" /id: " & oRec("id") &" /fomr: "& fomrkonto %>-->
				<%
				oRec.movenext
				Wend 
				oRec.close
		
		
		if fok = 0 then
		    if request.Cookies("erp")("krekonto") <> "" then ' <> "" then

		    krekontonr = request.Cookies("erp")("krekonto")
		    
		    strSQL = "SELECT k.kontonr, k.navn, k.id, k.kid, m.navn AS momskode FROM kontoplan k "_
			&" LEFT JOIN momskoder m ON (m.id = k.momskode) "_
			&" WHERE kontonr = " & krekontonr &" ORDER BY k.kontonr, k.navn"
		    
		    oRec.open strSQL, oConn, 3
		    if not oRec.EOF then
		    krekontonavn = oRec("navn")
		    kremomskode = oRec("momskode")
		    end if
		    oRec.close
		    %>
		    <option value="<%=krekontonr%>" SELECTED>(<%=krekontonr%>)&nbsp;&nbsp;<%=krekontonavn%> - <%=kremomskode %></option>
			
		    <%
		    end if 
		
		end if%>
		</select><br />
		   <span style="color:#999999;"><%=erp_txt_413 %></span>
		
            &nbsp;
		</td>
		</tr>
		<%
		
		'if func <> "red" then
		
		'if request.Cookies("erp")("momskonto") <> "" then
	         
	    '     if request.Cookies("erp")("momskonto") = "2" then
		'     chk2 = "CHECKED"
		'     chk1 = ""
    '		 else
    '		 chk1 = "CHECKED"
	'	     chk2 = ""
    '		 end if
	'	
	'	else
	'	
	'	    chk1 = "CHECKED"
	'	    chk2 = "" 
	'	 
	'	end if
	'	
	'	else
		     
	'	     if cint(momskonto) = 2 then
	'	     chk2 = "CHECKED"
	'	     chk1 = ""
    '		 else
    '		 chk1 = "CHECKED"
	'	     chk2 = ""
    '		 end if
	'	
	'	end if
		
		
		%>
		
        <input id="FM_momskonto" name="FM_momskonto" value="0" type="hidden" />
		<!--
		<tr><td colspan=2 valign=top style="padding:5px 5px 2px 20px;">
		<b>Momsberegning</b><br />
		V�lg hvilken konto's momsprocent der skal benyttes til momsberegning p� faktura og tilh�rende postering.<br />
		
		</td></tr>
		<tr>
		<td valign=top style="padding:2px 5px 5px 20px;"><input id="Radio1" type="radio" name="FM_momskonto" value="1" <%=chk1 %> /> Benyt debitkonto </td>
		<td valign=top style="padding:2px 5px 5px 20px;"><input id="Radio1" type="radio" name="FM_momskonto" value="2" <%=chk2 %> /> Benyt kreditkonto </td>
		</tr>
		-->
		
		</table>



        

		</div>

		

        <%select case lto
         case "nt", "intranet - local", "bf" 
          
            showNaesteVzb = "hidden"
            showNaesteDsp = "none"
           
         case else

            showNaesteVzb = "visible"
            showNaesteDsp = ""

        end select

       %>
		<div id="fidiv_2" style="position:absolute; visibility:<%=showNaesteVzb%>; display:<%=showNaesteDsp%>; top:840px; width:720px; left:5px; border:0px #8cAAe6 solid;">
        <table width=100% cellspacing=0 cellpadding=5 border=0><tr><td></td><td align=right><a href="#" onclick="showdiv('modtagdiv')" class=vmenu><%=erp_txt_417 %> >></a></td></tr></table>
		</div>
       
	<!-- Faktura indstillinger SLUT -->
	
	
	
	
	
	
	
	
	
	
	
	<!-- Job / aftale beskrivelse -->
	<div id="jobbesk" style="position:absolute; visibility:hidden; top:105px; width:720px; left:5px; border:1px #8cAAE6 solid;">
    <table cellspacing="0" cellpadding="5" border="0" width=100% bgcolor="#FFFFFF">
	
    <%if (cint(visjobbesk) = 1 AND func = "red") _
     OR (len(trim(strJobBesk)) <> 0 AND func <> "red" AND lto <> "essens") then 'OR (len(trim(strJobBesk)) <> 0 AND func = "red")
	visJb = "CHECKED"
	else
	visJb = ""
	end if
	
	
	%>
	<tr>
		<td valign=top>
		<input type="checkbox" name="FM_visikkejobnavn" id="FM_visikkejobnavn" value="1" <%=visikkejobnavnCHK%>><span><%=erp_txt_134 %> <b><%=erp_txt_135 %></b> <%=erp_txt_136 %> <b><%=erp_txt_137 %></b> <%=erp_txt_138 %></span><br />
	
		<input type="checkbox" name="FM_visjobbesk" id="FM_visjobbesk" value="1" <%=visJb%>><span><%=erp_txt_139 %> <b><%=erp_txt_140 %></b> <%=erp_txt_089 %></span>
	    
	<%
    
    if jobid <> 0 AND aftid = 0 then 'job faktura
    
    select case fastpris 
    case "1"
	jType = erp_txt_345
	case "3"
    jType = "Salesorder"
    case "4"
    jType = "Commission"    
    case else
	jType = erp_txt_299
	end select
    
    chklukjob = 0

        if cint(jobstatus) = 1 OR cint(jobstatus) = 2 OR cint(jobstatus) = 4 then
		diab = "" 
		    
		    
		    if lto = "essens" OR lto = "hestia" OR lto = "nt" then
		    chklukjob = 1
		    end if
		    
		    if chklukjob = 1 then
		    lkjobCHK = "CHECKED"
		    else
		    lkjobCHK = ""
		    end if
		    
		else
	    diab = "DISABLED"
	    end if 
	    %>
	    
	    
		
			<br />
			<input type="checkbox" name="FM_lukjob" id="FM_lukjob" <%=lkjobCHK %> value="1" <%=diab %>><span><%=erp_txt_141 %></span>
	    
	<br />
		
		
		<div style="position:relative; left:14px; top:10px; padding:10px; width:320px; height:250px; background-color:#F7F7F7;">
        <table cellpadding=0 cellspacing=4 border="0" width=100%>
        <tr><td ><%=erp_txt_142 %></td><td ><b><%=left(strJobnavn, 30) %> (<%=intjobnr %>)</b></td></tr>
		 <tr><td ><%=erp_txt_143 %></td><td ><b><%=jType %></b></td></tr>
		
		<%select case jobstatus
		case 1
		jobstatusTxt = "Aktiv"
		case 2
		jobstatusTxt = "Passiv/til fak."
        case 3
		jobstatusTxt = "Tilbud"
        case 4
		jobstatusTxt = "Gennemsyn"
		case 0
		jobstatusTxt = "Lukket"
		end select %>
		
		<tr><td ><%=erp_txt_093 %></td><td ><b><%=jobstatusTxt%></b> </td></tr>
		<tr><td ><%=erp_txt_144 %></td><td > <b><%=formatnumber(bruttooms, 2) &" "& valutaKode %></b> </td></tr>
		<tr><td ><%=erp_txt_145 %></td><td > <b><%=formatnumber(intBudgettimer, 2) %></b></td></tr>
        <tr><td ><%=erp_txt_146 %></td><td > <b><%=intRabat %> %</b> </td></tr>
		
		<!--
		Faktura grundlag: 
		<select case jobfaktype
		case 0
		Response.Write "<b>Den tid der bruges pr. medarb.</b><br>"
		case 1
		Response.Write "<b>Den aktivitet der udf�res</b><br>"
		end select >
		-->
		
		
		<tr><td >Periode:</td><td ><b><%=formatdatetime(jobstdato, 2) %></b> til <b><%=formatdatetime(jobsldato, 2) %></b> </td></tr>
		<tr><td  valign=top style="padding-top:5px;"><%=erp_txt_147 %><br />
            <span style="font-size:9px; color:#999999;">(<%=erp_txt_430 %>) </span>
		    </td><td ><input type="text" name="FM_rekvnr" style="font-size:9px; width:150px;" value="<%=rekvnr %>" /></td></tr>

        <%if aftaleId <> 0 then %>
        <tr><td ><%=erp_txt_148 %></td><td > <b><%=aftalenavn %></b> </td></tr>
        <%end if %>
		
		<%if cint(ski) = 1 then 
		skiTxt = erp_txt_453
		skiCHK = "CHECKED"
		else
		skiTxt = erp_txt_454
		skiCHK = ""
		end if%>
		
        <tr><td ><%=erp_txt_149 %> </td><td ><input id="FM_fak_ski" name="FM_fak_ski" type="checkbox" <%=skiCHK %> value="1" />  </td></tr>
		<tr><td ><%=erp_txt_150 %></td><td >

            <%
                strSQLFak = "SELECT f.fakdato, f.faknr, f.shadowcopy, f.beloeb, f.faktype FROM fakturaer f "_
		        &" WHERE (f.jobid = "& jobid &" AND medregnikkeioms = 0)"_
		        &" ORDER BY f.fakdato DESC"
        		
                antalKre = 0
                antalFak = 0
                beloebFaktureret = 0
		        oRec3.open strSQLFak, oConn, 3
                while not oRec3.EOF
                    
                   
                    
                    if oRec3("faktype") = 0 then 
                    antalFak = antalFak + 1
                    beloebFaktureret = beloebFaktureret + oRec3("beloeb") 
                    else
                    antalKre = antalKre + 1 
                    beloebFaktureret = beloebFaktureret - oRec3("beloeb")
                    end if            

               
                oRec3.movenext
                wend
                oRec3.close
                
                 %>

                <b><%=formatnumber(beloebFaktureret, 2) %> DKK</b> (<%=antalFak &" "& erp_txt_501%>  / <%=antalKre & " "& erp_txt_502 %>)
		                                                             </td></tr>


        <!--
		<%if lto = "execon" OR lto = "immenso" OR lto = "intranet - local" then %>
		
		<%if cint(abo) = 1 then %>
		<br /><br />Dette er en <b>Lightpakke</b>
            <input id="Hidden1" name="FM_fak_abo" value="1" type="hidden" />
		<%end if %>
		
		<%if cint(ubv) = 1 then %>
		<br />Jobbet er omfattet af <b>Udbudsvagten</b>
		<input id="Hidden2" name="FM_fak_ubv" value="1" type="hidden" />
		<%end if %>
		
		<%end if %>
        -->
		
        </table>
		</div>

        <%else '*aftale faktura
        
        fl_ant = 50
        dim fl_antal, fl_besk, fl_vis, fl_enhpris, fl_valuta, fl_enhed, fl_rabat, fl_momsfri, fl_belob
        redim fl_antal(fl_ant), fl_besk(fl_ant), fl_vis(fl_ant), fl_enhpris(fl_ant), fl_valuta(fl_ant), fl_enhed(fl_ant), fl_rabat(fl_ant), fl_momsfri(fl_ant), fl_belob(fl_ant)
        
        call akttyper2009(2)
        
         %>
        <div style="position:relative; left:14px; top:30px; padding:10px; width:300px; height:250px; background-color:#F7F7F7;">
            <table cellpadding=0 cellspacing=1 border="0" width=100%>
            <tr><td ><%=erp_txt_151 %></td><td >
            <b><%=left(strNavn, 40) %></b></td></tr>
     
        <tr><td ><%=erp_txt_152 %></td><td > <b><%=formatnumber(intPris, 2) %></b></td></tr>
        <tr><td ><%=erp_txt_153 %></td><td ><b><%=formatnumber(intEnheder, 2) %></b></td></tr>
        <tr><td ><%=erp_txt_154 %></td><td ><b><%=formatdatetime(startdato, 2) %></b> <%=erp_txt_090 %>  <b><%=formatdatetime(slutdato, 2) %></b></td></tr>
        <tr><td ><%=erp_txt_146 %></td><td ><b><%=intRabat %> %</b></td></tr> 
        <tr><td  colspan=2>
        <b><%=erp_txt_155 %></b><br />
        <%strSQLaftalejob = "SELECT jobnavn, jobnr, id AS jid, valuta FROM job WHERE serviceaft = "& aftid & " AND serviceaft <> 0"

        'Response.write "strSQLaftalejob " & strSQLaftalejob
        'Response.end

        ja = 0
        strJobPaaAft = ""
        oRec.open strSQLaftalejob, oConn, 3
        while not oRec.EOF 

        if ja <> 0 then
        strJobPaaAft = strJobPaaAft & ", "
        end if

        strJobPaaAft = strJobPaaAft & oRec("jobnavn") & " ("& oRec("jobnr") &")" 
        
        fl_besk(ja) = oRec("jobnavn") & " ("& oRec("jobnr") &")"
        
                
                '** Henter hovedlinje til gl. fakturaer p� aftaler f�r 15/11-2010 **'
               
                sumEnhFak = 0
                strSQLenh = "SELECT sum(timer * a.faktor) AS sumEnhFak FROM timer "_
                &" LEFT JOIN aktiviteter a ON (a.id = taktivitetid) WHERE "_
                &" ("& replace(aty_sql_fakbar, "fakturerbar", "tfaktim") &") AND seraft = "& aftid &" AND tjobnr = '"& oRec("jobnr") &"' "_ 
                &" AND tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"' GROUP BY seraft"
	            
                'Response.Write strSQLenh &"<br>"
	            'Response.end


                oRec3.open strSQLenh, oConn, 3
                if not oRec3.EOF then
                
                sumEnhFak = oRec3("sumEnhFak")
                
                end if
                oRec3.close

                if sumEnhFak <> 0 then

                fl_antal(ja) = sumEnhFak
                fl_vis(ja) = "CHECKED"
                
                if intPris <> 0 AND intEnheder <> 0 then '** Aftale enh pris
                fl_enhpris(ja) = formatnumber((intPris / intEnheder), 2)
                else
                fl_enhpris(ja) = 0
                end if

                fl_valuta(ja) = oRec("valuta")
                thisaktfunc(ja) = "opr"
                fl_enhed(ja) = 2
                fl_rabat(ja) = intRabat/100 '0
                fl_momsfri(ja) = 0
                fl_belob(ja) = formatnumber(fl_enhpris(ja) * fl_antal(ja), 2)

                else

                fl_antal(ja) = 0
                fl_vis(ja) = ""
                fl_enhpris(ja) = 0
                fl_valuta(ja) = oRec("valuta")
                thisaktfunc(ja) = "opr"
                fl_enhed(ja) = 2
                fl_rabat(ja) = intRabat/100 '0
                fl_momsfri(ja) = 0
                fl_belob(ja) = 0

                end if


                

            ja = ja + 1
            oRec.movenext
            wend
            oRec.close 
        

     

        if ja = 0 then
        strJobPaaAft = erp_txt_423
        end if
        %>

        <%=strJobPaaAft%></td></tr></table>
    </div>
		
	<%end if 'job/aftale 
    
    
    '**** Editor og Sideombrydninger indstillinger ***'
    select case lto
	case "dencker", "outz"
	sidetoluft = 100
    sidetoluftPdTop = 30
    stopPx1 = 1415 '1340
    stopPx2 = 2955 '2880
    stopPx3 = 4495 '4420
    stopPx4 = 6015 '5960
    stopPx5 = 6575 '6500
    stopPx6 = 7715 '7640
    stopPx7 = 8955 '8880
    eidtorHgt = 8755 '8680 '3680
   
    case "synergi1", "intranet - local"
    sidetoluft = 420'380
	sidetoluftPdTop = 100
    stopPx1 = 1225 '1150'1190
    stopPx2 = 2405 '2330 '2660
    stopPx3 = 3595 '3520
    stopPx4 = 4775 '4700
    stopPx5 = 5965 '5890
    stopPx6 = 7195 '7120
    stopPx7 = 8525' 8450
    eidtorHgt = 9225 '9150
    case else 'Essens, JT-Tek, FE
	sidetoluft = 20
    sidetoluftPdTop = 10
    stopPx1 = 1415 '1340
    stopPx2 = 2955 '2880
    stopPx3 = 4495 '4420
    stopPx4 = 6015 '5960
    stopPx5 = 6575 '6500
    stopPx6 = 7715 '7640
    stopPx7 = 8955 '8880
    eidtorHgt = 8755 '8680 '3680
	end select
    
    
    %>	
		
	</td>
    <td style="padding-top:80px;">
        <a href="#" id="showinternnote" class="vmenu"><u><%=erp_txt_157 %></u></a>
        <%if cint(job_alert) = 1 then%>
            <span style="color:red;">&nbsp;<b>!</b>&nbsp;</span>
         <%end if %>
         | <a href="#" id="showjobtweet" class="vmenu"><%=erp_txt_156 %></a>
        <br /><div id="internbesk_tweet" style="position:relative; width:300px; height:261px; overflow:auto; font-size:9px; border:0px #cccccc solid; padding:10px;"><%=job_internbesk %></div>
       
       <div id="internbesk_hd" style="display:none; visibility:hidden;"><%=job_internbesk %></div>
       <div id="tweet_hd" style="display:none; visibility:hidden;"><%=job_tweet %></div>
       </td>
    </tr>
	<tr><td colspan="2" style="padding:20px 20px 20px 20px;"><b><%=erp_txt_158 %></b><br />
    <span style="font-size:11px; color:#999999;"><%=erp_txt_159 %><br>
    <%=erp_txt_160 %></span>
	
	                    <%
                        if len(trim(strJobBesk)) <> 0 then
                        strJobBesk = replace(strJobBesk, "<div>", "")
                        strJobBesk = replace(strJobBesk, "</div>", "<br>")
                        else
                        strJobBesk = ""
                        end if
	                    

                            select case lto
                            case "synergi1", "intranet - local"
                                if func <> "red" then
                                content = "<span style=""color:red; font-size:14px;"">Bem�rk nyt kontonr!<br></span>" & strJobBesk
                                else
                                content = strJobBesk
                                end if

                            case else
                            content = strJobBesk
            			    end select
			            
			            Set editorJ = New CuteEditor
            					
			            editorJ.ID = "FM_jobbesk"
			            editorJ.Text = content
			            editorJ.FilesPath = "CuteEditor_Files"
			            editorJ.AutoConfigure = "Minimal"
                        select case lto 
                        case "synergi1", "xintranet - local"
                        editorJ.EditorBodyStyle = "font:normal 12px arial;"
                        editorJ.Width = 440 '520'640
                        case else
                        editorJ.Width = 620
                        end select
                        
            			
			            
			            editorJ.Height = eidtorHgt
			            editorJ.Draw()
		                %>
	
	
	
	<br />
     <%
		                uTxt = "<b>"& erp_txt_432 &":</b><br>"& erp_txt_431 &""
						uWdt = 300
								
					    call infoUnisport(uWdt, uTxt) 
		                 %>
	</td></tr>
	</table>
    </div>
	

    <%  '** Sideombrydninger **** %>

	

    <div id="d_sideombr1" style="position:absolute; top:<%=stopPx1 %>px; left:0px; width:730px; height:<%=sidetoluft%>px; border:3px darkred dashed; background:#F7F7F7; zoom: 1; filter: alpha(opacity=50);opacity: 0.5; padding:<%=sidetoluftPdTop %>px 0px 0px 200px; z-index:20000000;">
    <h4><%=erp_txt_161 %></h4>
   </div>

    <div id="d_sideombr2" style="position:absolute; top:<%=stopPx2 %>px; left:0px; width:730px; height:<%=sidetoluft%>px; border:3px darkred dashed; background:#F7F7F7; zoom: 1; filter: alpha(opacity=50);opacity: 0.5; padding:<%=sidetoluftPdTop %>px 0px 0px 200px; z-index:20000000;">
    <h4><%=erp_txt_162 %></h4>
    </div>

  
    
      <div id="d_sideombr3" style="position:absolute; top:<%=stopPx3 %>px; left:0px; width:730px; height:<%=sidetoluft%>px; border:3px darkred dashed; background:#F7F7F7; zoom: 1; filter: alpha(opacity=50);opacity: 0.5; padding:<%=sidetoluftPdTop %>px 0px 0px 200px; z-index:20000000;">
    <h4><%=erp_txt_163 %></h4>
    </div>

      <div id="d_sideombr4" style="position:absolute; top:<%=stopPx4 %>px; left:0px; width:730px; height:<%=sidetoluft%>px; border:3px darkred dashed; background:#F7F7F7; zoom: 1; filter: alpha(opacity=50);opacity: 0.5; padding:<%=sidetoluftPdTop %>px 0px 0px 200px; z-index:20000000;">
    <h4><%=erp_txt_164 %></h4>
    </div>
   
   
      <div id="d_sideombr5" style="position:absolute; top:<%=stopPx5 %>px; left:0px; width:730px; height:<%=sidetoluft%>px; border:3px darkred dashed; background:#F7F7F7; zoom: 1; filter: alpha(opacity=50);opacity: 0.5; padding:<%=sidetoluftPdTop %>px 0px 0px 200px; z-index:20000000;">
    <h4><%=erp_txt_165 %></h4>
    </div>

     <div id="d_sideombr6" style="position:absolute; top:<%=stopPx6 %>px; left:0px; width:730px; height:<%=sidetoluft%>px; border:3px darkred dashed; background:#F7F7F7; zoom: 1; filter: alpha(opacity=50);opacity: 0.5; padding:<%=sidetoluftPdTop %>px 0px 0px 200px; z-index:20000000;">
    <h4><%=erp_txt_166 %></h4>
    </div>

     <div id="d_sideombr7" style="position:absolute; top:<%=stopPx7 %>px; left:0px; width:730px; height:<%=sidetoluft%>px; border:3px darkred dashed; background:#F7F7F7; zoom: 1; filter: alpha(opacity=50);opacity: 0.5; padding:<%=sidetoluftPdTop %>px 0px 0px 200px; z-index:20000000;">
    <h4><%=erp_txt_167 %></h4>
    </div>
   
   
   
  

  

	<div id=jobbesk_2 style="position:absolute; visibility:hidden; display:none; top:<%=eidtorHgt+700%>px; width:720px; left:5px; border:0px #8cAAe6 solid;">
        <table width=100% cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('modtagdiv')" class=vmenu><< <%=erp_txt_072 %></a></td><td align=right><a href="#" onclick="showdiv('aktdiv')" class=vmenu><%=erp_txt_073 %> >></a></td></tr></table>
    
    
    
    </div>
	
	<!-- Job / aftale beskrivelse SLUT -->
	 <% 
	 
	 
	 '*** Mat og Job log ***
	 call joblogdiv()
	 call matlogdiv()
	    
	    
	%> 



    <!--- Betalingsbetingelser --->
    <div id=betdiv style="position:absolute; visibility:hidden; top:105px; width:720px;"> <!--  width:720px; visibility:hidden; display:none; border:1px #8caae6 solid; top:105px; padding:40px 0px 0px 0px; left:5px; background-color:#EFf3FF;-->
     
 
	<table width=100% cellspacing=5 cellpadding=0 border=0>
	<tr>
		<td style="padding:0px 0px 0px 5px;">
               <b><%=erp_txt_168 %></b><br />
		
		                <%
                        select case lto
                        case "nt"
                             
                            if func <> "red" then

                                    content = ""
                                    if betbetint <> "21" then
                                    content = content & betbetint_txt &" TT payment, "
                                    else
                                    content = content & "NetCash, "
                                    end if

                                    '** NT commisionsordre ***
                                    if cint(fastpris) = 2 then '** faktura skal h�gtes p� leverand�r
                                        
                                        select case lev_levbetint
                                        case 1
                                        content = content & "FOB"
                                        case else
                                        content = content & "DDP"
                                        end select

                                        'select case lev_betbetint

                                    else
                             
                                        select case kunde_levbetint
                                        case 1
                                        content = content & "FOB"
                                        case else
                                        content = content & "DDP"
                                        end select

                                        'select case 

                                    end if


                             else

                                content = strKom

                             end if


                        case else
	                    content = strKom 
            			end select
			            

                        '*** Midlertidig sl�et fra 20150824 ****
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_komm"
			            editorK.Text = content
			            editorK.FilesPath = "CuteEditor_Files"
			            editorK.AutoConfigure = "minimal"
            			editorK.Width = 700
			            editorK.Height = 380 '280
			            editorK.Draw()
		                %>
                       <!-- <textarea name="FM_komm" style="width:700px; height:270px;"><%=content %></textarea>-->
	
		<br>
		<input type="checkbox" name="FM_gembetbet" id="FM_gembetbet" value="1"><%=erp_txt_169 %> <b><%=erp_txt_170 %></b> <%=erp_txt_171 %><br>
		
		<%if level = 1 then%>
		<input type="checkbox" name="FM_gembetbetalle" id="FM_gembetbetalle" value="1"><%=erp_txt_172 %> <b><%=erp_txt_173 %></b> <%=erp_txt_174 %><br>
		<%end if%>
		<input type="hidden" name="FM_kundeid" id="FM_kundeid" value="<%=intKid%>">
		</td>
	</tr>
	
	</table>
	


	
	 <div id=betdiv_2 style="position:relative; visibility:hidden; display:none; top:80px; width:600px; left:5px; border:0px #8cAAe6 solid;">
        <table width=720 cellspacing=0 cellpadding=5 border=0><tr><td><a href="#" onclick="showdiv('aktdiv')" class=vmenu><< <%=erp_txt_072 %></a></td><td align=right>
            &nbsp;<input name="subm_on" id="subm_on" type="submit" value="<%=erp_txt_464 %> >>" /></td></tr></table>
    </div>
	
	
	<!-- Nedenst�ende Bruges af javascript --->
	<input type="hidden" name="FM_showalert" id="FM_showalert" value="0">
	

      
	
	</div><!-- Betalingsbet -->







	    
	    
	<!--<div id="aktiviteter" style="position:absolute; left:10px; top:920px; width:500px; visibility:visible; z-index:100; background-color:#d6dff5;">-->
	<%
	
	call akttyper2009(4)
	'Response.Write "aty_sql_onfak: " & aty_sql_onfak  
	'Response.end








	
	'*** Udspecificering af aktiviteter p� job **'
	if jobid <> 0 AND aftid = 0 then '** jobid <> 0 (Vises kun ved fakturaer p� job) %>
	<!--#include file="inc/fak_job_inc_2007.asp"-->
	<!--#include file="inc/mat_inc_2008.asp"-->        
	        
	<%else
	'*** Fakturerer Aftale ***'
	%>
	<!--#include file="inc/fak_aft_inc_2007.asp"-->
    <!--#include file="inc/mat_inc_2008.asp"-->  
	<%end if %>
	

               
	
	
	<br>
        &nbsp;
<!--
</div>-->




  
	<div id="faksubtotal" style="position:absolute; left:730px; top:264px; width:200px; z-index:2000; border:1px #8caae6 solid; background-color:#ffffff; padding:5px;">
    <b><%=erp_txt_175 %></b>
    

	<table width=100% cellspacing=5 cellpadding=0 border=0>
	<tr>
	<td align="right"><%=erp_txt_176 %>

    <br />
     <span style="font-size:9px; color:#999999;">
    <%if func <> "red" then
   
    if jobid <> 0 then %>

   
    <%if fastpris = 1 then%>
    <%=erp_txt_177 %>
    <%else %> 
    <%=erp_txt_178 %>
    <%end if %>
   

    <%else 'aftale %>
    <%=erp_txt_179 %>
    <%end if %>
   

     <%else %>
     <%=erp_txt_180 %>
    <%end if %>


      </span>

		<!-- strBeloeb -->
		<%

        if func <> "red" then

            select case lto
            case "synergi1", "xintranet - local"
            totbel_afvigeCHK = "CHECKED"

            case "dencker"

            totbel_afvigeCHK = ""

            case else

            if cint(fastpris) = 1 then
            totbel_afvigeCHK = "CHECKED"
            else
            totbel_afvigeCHK = ""
            end if

            end select

        else
            if cint(totbel_afvige) = 1 then
            totbel_afvigeCHK = "CHECKED"
            else
            totbel_afvigeCHK = ""
            end if
        end if

        if func <> "red" then

         if jobid <> 0 then
                    if fastpris = 1 then '** Kun p� faspris job og ikke p� aftaler **'
                    totalbelob = formatnumber(useBelob_red)
                    else
		            totalbelob = totalbelob + matSubTotalAll
                    end if
        else
                    totalbelob = formatnumber(intPris)
        end if

		'totalbelob_umoms = (matSubTotalAlluMoms/1) + (aktSubTotalAlluMoms/1)
		else
        totalbelob = formatnumber(useBelob_red)
        'totalbelob_umoms = formatnumber(useMoms_red)
        end if

        totalbelob_umoms = (matSubTotalAlluMoms/1) + (aktSubTotalAlluMoms/1)


		if len(totalbelob) <> 0 then
		thistotbel = SQLBlessDot(formatnumber(totalbelob, 2))
		else
		thistotbel = formatnumber(0, 2)
		end if
		
		if len(totalbelob_umoms) <> 0 then
		'Response.Write "" & totalbelob_umoms
		'Response.end
		thistotbel_Umoms = SQLBlessDot(formatnumber(totalbelob_umoms, 2))
		else
		thistotbel_Umoms = formatnumber(0, 2)
		end if
		%>
        <br />
		<input type="text" id="FM_beloeb" name="FM_beloeb" style="width:120px; border-bottom:2px #86B5E4 dashed;" value="<%=thistotbel%>"><span style="position:relative; width:20px; border-bottom:2px #86B5E4 dashed; padding:4px;" id="divbelobtot"><b><%=valutakodeSEL%></b></span>
		<br /><br /><input type="checkbox" value="1" name="FM_totbel_afvige" <%=totbel_afvigeCHK %> /><span style="position:relative; width:20px; padding:4px; font-size:9px;" id="Span1"><%=erp_txt_181 %></span>
		<input type="hidden" id="FM_beloeb_umoms" name="FM_beloeb_umoms" value="<%=thistotbel_Umoms%>">
		<br /><br />
		<div style="position:relative; width:120px; border-bottom:0px #86B5E4 dashed; font-family:arial; color:#999999; font-size:10px; padding-right:3px;" align="right" name="divbelobtot_umoms" id="divbelobtot_umoms"><%=erp_txt_182 %><br /> (<%=thistotbel_Umoms &" "& valutakodeSEL%>)</div>
		
		</td>
	</tr>
	</table>
        
	</div>
					
									
		
            <input id="lto" value="<%=lto %>" type="hidden" />				
		</form>							
									
								
</div> <!--sidetop -->
 <!-- main -->





<%
'** Viser menu efter side er loadet f�rdig ***'
 '** S�tter default v�rdier til enheder **'
if (lto = "dencker" AND func <> "red") then
Response.Write("<script>opd_akt_endhed('Pr. enhed','2');</script>") 
end if

'Response.Write("<script>showmenu();</script>")


'*** Opdater valuta p� sumfelter **'
if jobid <> 0 AND valuta <> 1 then
'Response.Write("<script>showvalutaload();</script>")
'Response.Write("<script>opdatervalutAllelinier(1,0);</script>")
'Response.Write("<script>hidevalutaload();</script>")
end if
%>

<%end select%>
<%end if





%>



