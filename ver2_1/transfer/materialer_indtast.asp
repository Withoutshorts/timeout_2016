<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/xml/tsa_xml_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/mat_func.asp"-->


<%


if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<% 
	errortype = 5
	call showError(errortype)
	else
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	
	func = request("func")
	
	'Response.Write "func:" & func
	
	'** jobid **'
	if len(trim(request("id"))) <> 0 then
	id = request("id")
	else
	id = 0
	end if
	
	    if len(request.cookies("timereg_2006")("usemrn")) <> 0 then
		usemrn = request.cookies("timereg_2006")("usemrn")
		else
		usemrn = session("mid")
		end if
	
	if len(request("aftid")) <> 0 then  
	aftid = request("aftid")
	else
	aftid = 0
	end if
	
	if len(trim(request("fromsdsk"))) <> 0 then
	fromsdsk = request("fromsdsk")
	else
	fromsdsk = 0
	end if
	
	if len(request("lastid")) <> 0 then
	lastid = request("lastid")
	else
	lastid = 0
	end if
	
	if len(request("matregid")) <> 0 then
	    matregid = request("matregid")
	    else
	    matregid = 0
	    end if
	
	'select case lto 
	'case "syncronic"
	'level = 1
	'case else
	level = session("rettigheder")
	'end select
	
	session.lcid = 1030 'DK
	
	if len(request("vis")) <> 0 then
	vis = request("vis")
	else
	    
	    '* vis = request.cookies("timereg_2006")("vismatreg")
	    select case lto
	    case "xx"
	    vis = ""
	    case else
	    vis = "otf"
	    end select
	    
	end if

    if len(trim(request("sogBilagOrJob"))) <> 0 then
    sogBilagOrJob = request("sogBilagOrJob")
            
            if sogBilagOrJob = 1 then
            sogBilagOrJobCHK1 = "CHECKED"
            else
            sogBilagOrJobCHK0 = "CHECKED"
            end if
    else
            
            if request.cookies("timereg_2006")("sogBilagOrJob")  <> "" then
                
                
                sogBilagOrJob = request.cookies("timereg_2006")("sogBilagOrJob")    

                if sogBilagOrJob = 1 then
                sogBilagOrJobCHK1 = "CHECKED"
                else
                sogBilagOrJobCHK0 = "CHECKED"
                end if

            else
            sogBilagOrJob = 0
            sogBilagOrJobCHK0 = "CHECKED"
            end if
     end if
	
    response.cookies("timereg_2006")("sogBilagOrJob") = sogBilagOrJob



	if vis = "otf" then
	Knap1_bg = "#5C75AA"
	Knap2_bg = "#3B5998"
	else
	Knap1_bg = "#3B5998"
	Knap2_bg = "#5C75AA"
	end if
	
	Knap3_bg = "#3B5998"
	Knap4_bg = "#3B5998"
	
	response.cookies("timereg_2006")("vismatreg") = vis
	
	
	'*** LUKKET FOR REDIGERING ***'
	'*** Er Smiley og Auto Gk sl�et til ? **'
	'*** Auto gk er om periode skal lukkes ved godkend uge **''
	call ersmileyaktiv()
	
	call licensStartDato()
	
	stDato = startDatoAar &"/"& startDatoMd &"/"& startDatoDag
	
	'*** Afsluttede uger for medarb. Finder StDato p� Interval ***
	strSQL = "SELECT mf.forbrugsdato FROM materiale_forbrug mf "_
	&" WHERE usrid = "& usemrn &" ORDER BY mf.forbrugsdato, mf.id DESC LIMIT 40" 
	
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	stDato = year(oRec("forbrugsdato")) &"/"& month(oRec("forbrugsdato")) &"/"& day(oRec("forbrugsdato"))
	end if
	oRec.close
	
	slDato = year(now) &"/"& month(now) &"/"& day(now)
	call afsluger(usemrn, stdato, sldato)
	
	
	
	
	select case func
    case "slet"
    
    
    
    matid = 0
    matantal = 0
    
    '*** Henter materiale id og antal ***'
    strSQLsel = "SELECT matantal, matid FROM materiale_forbrug WHERE id = " & matregid
    oRec.open strSQLsel, oConn, 3
    if not oRec.EOF then
    
    matid = oRec("matid")
    matantal = oRec("matantal")
    
    end if
    oRec.close
    
    
    
    '** Opdaterer antal p� lager ***
	strSQL2 = "UPDATE materialer SET antal = (antal+("&matantal&")) WHERE id = "&matid
	oCOnn.execute(strSQL2)
	
	'*** Sletter registrering ***
	strSQLdel = "DELETE FROM materiale_forbrug WHERE id = " & matregid
    oConn.execute(strSQLdel) 
    
    Response.Redirect "materialer_indtast.asp?id="&id&"&fromsdsk="&fromsdsk&"&aftid="&aftid&"&vis="&vis
    	

	case "dbopr", "dbred"
	'*** Her inds�ttes en ny type i db ****
	%>
	<!--#include file="inc/isint_func.asp"-->
	<%
		
		if len(trim(request("FM_antal"))) <> 0 then
		intAntal = replace(request("FM_antal"), ",", ".")
		else
		intAntal = 0
		end if
			
			
			call erDetInt(intAntal)
			if isInt > 0 then
			
			useleftdiv = "c"
			errortype = 54
			call showError(errortype)
			
			isInt = 0
		
		else
			
			
		strEditor = session("user")
		strDato = year(now)&"/"&month(now)&"/"&day(now)
		matId = request("FM_matid")
		jobid = request("jobid")
		
		
		     if cint(jobid) = 0 then
			'useleftdiv = "c"
			errortype = 140
			call showError(errortype)
		    
		    Response.end
		    end if
		'Response.Write "jobid" & jobid
		'Response.Flush
		'Response.end 
		
		       
		       
		       
		        
		         if len(request("onthefly")) <> 0 then
		         onthefly = request("onthefly")       
		         else
		         onthefly = 0
		         end if      
		        
		        
		        if onthefly = "1" then
			    regdato = request("regdato_0")
			    
			    if isDate(regdato) = false then
			    
			    errortype = 142
			    call showError(errortype)
			    
			    Response.end
			    end if 
			    
			     '** Valuta, Intkode, Bilagsnr **'
                valuta = request("FM_valuta")
                intkode = request("FM_intkode")
                
                if len(trim(request("FM_personlig"))) <> 0 then
                personlig = 1
                else
                personlig = 0
                end if
                
                
                if len(request("FM_bilagsnr")) <> 0 then
                bilagsnr = request("FM_bilagsnr") 
                else
                bilagsnr = ""
                end if
			    
			    
	            else
	            
	            regdato = request("regdato_"&matId&"")
	            
	             '** Valuta, Intkode, Bilagsnr **'
                valuta = request("FM_valuta_"&matId&"")
                intkode = request("FM_intkode_"&matId&"")
                personlig = request("FM_personlig_"&matId&"")
                
                if len(request("FM_bilagsnr_"&matId&"")) <> 0 then
                bilagsnr = request("FM_bilagsnr_"&matId&"") 
                else
                bilagsnr = ""
                end if
	            
	            
	            end if
		        
		        regDatoSQL = year(regdato) &"/"& month(regdato) &"/"& day(regdato)
		        
		        if isdate(regdato) = false then
		       
			    useleftdiv = "c"
			    errortype = 90
			    call showError(errortype)
		        else
		                
		        
		        '**** Hvis materiale forbrug  / Udl�g skal videre faktureres p� job **'
		        if cint(intkode) = 2 OR cint(intkode) = 0  then  '** Ekstern / Ikke angivet **'
		        
		             
		             '*** Er uge alfsuttet af medarb, er smiley og autogk sl�et til **'
                    erugeafsluttet = instr(afslUgerMedab(usemrn), "#"&datepart("ww", regdato,2,2)&"_"& datepart("yyyy", regdato) &"#")
                    
                    'Response.Write "erugeafsluttet --" & erugeafsluttet  &"<br>"
                    'Response.flush
                    call lonKorsel_lukketPer(regdato)
                  
                     if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", regdato) = year(now) AND DatePart("m", regdato) < month(now)) OR _
                    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", regdato) < year(now) AND DatePart("m", regdato) = 12)) OR _
                    (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", regdato) < year(now) AND DatePart("m", regdato) <> 12) OR _
                    (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", regdato) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
                  
                    ugeerAfsl_og_autogk_smil = 1
                    else
                    ugeerAfsl_og_autogk_smil = 0
                    end if 
                
		        
		                
		            '*** Skal ikke l�ngere tjekke om er foreligger en faktura, da TimeOut nu holder �je med 
		            '*** Om materiale er faktureret ***'    
		            '*** Tjekker sidste fakdato ***'
		            lastFakdato = "01/01/2002"
		            'strSQL = "SELECT fakdato FROM fakturaer WHERE jobid = "& jobid &" AND faktype = 0 ORDER BY fakdato DESC LIMIT 0,1"
		            'oRec.open strSQL, oConn, 3
		            'if not oRec.EOF then
		            'lastFakdato = oRec("fakdato")
		            'end if
		            'oRec.close
    		        
		            if ugeerAfsl_og_autogk_smil = 1 then ' OR cdate(lastFakdato) >= cdate(regdato) then
    		        
		            
    			    
			        'if lastFakdato = "01/01/2002" then
			        'lastFakdato = "(ingen)"
			        'end if
    			    
			        useleftdiv = "c"
			        errortype = 108
			        call showError(errortype)
    		        
		            Response.end
		            end if
    		        
		        
		        
                
                end if '** Ekstern ***'        
		
                		
		                  
                		                
		                '**********************************************************'
		                '**** Opretter materiale hvis der er valgt "ON the fly" ***'
		                '**********************************************************'
		                if onthefly = "1" then
		                
		                
	                    call erDetInt(request("pris")) 
			            if isInt > 0 OR len(trim(request("pris"))) = 0 then
            				
				          
            				
				            errortype = 106
				            call showError(errortype)
            			
			            isInt = 0
            			
			            Response.end
			            end if
			            
			            call erDetInt(request("FM_salgspris")) 
			            if isInt > 0 OR len(trim(request("FM_salgspris"))) = 0 then
            				
				            
            				
				            errortype = 120
				            call showError(errortype)
            			
			            isInt = 0
            			
			            Response.end
			            end if
			            
			            if len(trim(request("navn"))) = 0 then
			             
				            errortype = 107
				            call showError(errortype)
            			
			            
			            Response.end
			            end if
			            
			            matGrp = request("gruppe")
			            if cint(matGrp) = 0 AND (lto = "execon" OR lto = "immenso") then
			            
			             errortype = 147
				            call showError(errortype)
            			
			            
			            Response.end
			            
			            end if
			            
			            matVarenr = request("varenr")
			            
			            
			           '** Skal mat / udl�g oprettes p� lager??
		               if len(trim(request("opretlager"))) <> 0 then
		               opretlager = 1
		               else
		               opretlager = 0
		               end if 
		               
		               
	                         if cint(opretlager) = 1 AND matVarenr = "0" then
			                 
				                errortype = 143
				                call showError(errortype)
                			
    			            
			                Response.end
			                end if
			          
			            '*** Findes en vare / et materiale allerede med det valgte varenr'    
			            '*** I den valgte gruppe'
			            findesallerede = 0
			            
			            if matVarenr <> "0" then
			            
			            strSQLfindes = "SELECT id, varenr, navn FROM materialer WHERE varenr = '"& matVarenr &"'"_
			            &" AND matgrp = " & matGrp
			            
			            oRec.open strSQLfindes, oConn, 3
	                    if not oRec.EOF then
	                    findesallerede = 1
	                    end if
	                    oRec.close
	                    
	                         if cint(findesallerede) = 1 then
			                
                				
				                errortype = 74
				                call showError(errortype)
                			
    			            
			                Response.end
			                end if
			            
			            end if
	                    
	                    
	                    '*******************************************'
	                    '*** K�bs og Salgspris + Ava beregnning ****'
	                    '*******************************************'
	                    
	                                    
		                matNavn = replace(request("navn"), "'", "''") 
		                
		                dblKobsPris = replace(request("pris"), ",", ".")	
		                dblSalgsPris = replace(request("FM_salgspris"),",",".")
			           
		                betegnelse = replace(request("FM_betegn"), "'", "''") 
		                
		                '*** End k�bs og salgspris ***'
		                
		              
		                
		                      '**** Opretter on the fly p� lager **'
		                      if  matVarenr <> "0" AND cint(opretlager) = 1 then
		                            
		                            intprvAntal = 0
		                            intMatgrp = matGrp
		                            
		                            strSQLins = "INSERT INTO materialer (editor, dato, navn, "_
		                            &" varenr, matgrp, antal, indkobspris, salgspris, "_
						            &" enhed, betegnelse) VALUES "_
						            &"('"& strEditor &"', '"& strDato &"', "_
						            &" '"& matNavn &"', '"& matVarenr &"',"_
						            &""& matGrp &", 0, "_
						            &""& dblKobsPris &", "& dblSalgsPris  &", 'stk.', '"& betegnelse &"')"
            						
						            'Response.write strSQLins & "<br>"
						            'Response.end
            						
						            oConn.execute(strSQLins)
            						
						            matId = 0
						            strSQLmid = "SELECT id FROM materialer WHERE id <> 0 ORDER BY id DESC"
						            oRec.open strSQLmid, oConn, 3
						            if not oRec.EOF then
						            matId = oRec("id")
						            end if
						            oRec.close
            						
            			    end if
            			    '*** End opret p� lager **'
						
						
						
						response.Cookies("mat")("visdiv") = "otf"
            						
			            else
						
			            response.Cookies("mat")("visdiv") = "ima"
						
			            end if 
						'**** End On the fly ***'
						
		
			
			
			
			'*************************************************************'
			'*** Henter Materiale data og indl�ser materiale forbrug  ****'
			'*************************************************************'
			
			if onthefly = "1" then
			    
			    strVarenr = matVarenr
			    strNavn = matNavn
			    avaGrp = matGrp
			    avaProcent = av
			    strEnhed = betegnelse
			    
			   
    			
			else
			
			'**************************************************'
			'*** indtasting fra lager ***'
			'**************************************************'
			
			        dblKobsPris = request("FM_kobspris_"&matId&"")
			        avaGrp = request("FM_avagruppe_"&matId&"")
			        dblSalgsPris = request("FM_salgspris_"&matId&"")
			        
			        if len(trim(request("FM_opdaterpris_"&matId&""))) <> 0 then
			        opdaterPris = 1
			        else
			        opdaterPris = 0
			        end if
			        
			         call erDetInt(dblKobsPris) 
			            if isInt > 0 OR len(trim(dblKobsPris)) = 0 then
            				
				            
            				
				            errortype = 106
				            call showError(errortype)
            			
			            isInt = 0
            			
			            Response.end
			            end if
			            
			            call erDetInt(dblSalgsPris) 
			            if isInt > 0 OR len(trim(dblSalgsPris)) = 0 then
            				
				          
				            errortype = 120
				            call showError(errortype)
            			
			            isInt = 0
            			
			            Response.end
			            end if
			        
			
			
			                            '*** Henter StamData p� materiale ***'
			                            strSQL2 = "SELECT id, navn, editor, dato, varenr, antal, "_
			                            & "salgspris, indkobspris, enhed, matgrp FROM materialer WHERE id=" & matId
                            			
			                            'Response.Write strSQL2
			                            'Response.flush
                            			
			                            oRec2.open strSQL2, oConn, 3
			                            if not oRec2.EOF then
                            			        
                            			      
			                            strNavn = oRec2("navn")
			                            strVarenr = oRec2("varenr")
                            			
                            			     
			                                avaProcent = 0
                                			
			                                 strSQL3 = "SELECT id, navn, av FROM materiale_grp WHERE id = "& avaGrp
			                                 oRec3.open strSQL3, oConn, 3
			                                 if not oRec3.EOF then
			                                 avaProcent = oRec3("av")
			                                 end if
			                                 oRec3.close
                            			
                            			
			                            strEnhed = oRec2("enhed")
			                            intMatgrp = oRec2("matgrp")
			                            intprvAntal = oRec2("antal")
                            			
                            			
			                            end if
			                            oRec2.close
                            			
			        
			        
			        dblKobsPris = replace(dblKobsPris,",",".")
			        dblSalgsPris = replace(dblSalgsPris,",",".")
			        
			        '** Opdaterer priser p� mat **'
			        if cint(opdaterPris) = 1 then
			        strSQLpris = "UPDATE materialer SET salgspris = "& dblSalgsPris &", indkobspris = "& dblKobsPris &" WHERE id = "& matId
				    oCOnn.execute(strSQLpris)
				    end if
			        
			
			
			end if
			'*** END henter mat. stamdata ***'
			
			
			
			
			
			'Response.Write "dblKobsPris # dblSalgsPris " & dblKobsPris &" # " & dblSalgsPris
			'Response.end 
			
			
			   
			   
			   
			    '*************************************************************'
			    '**** Indl�ser materiale forbrug                           ***'
			    '*************************************************************'
			    
			    '*** Valuta kurs ***'
			    intValuta = valuta
			    call valutaKurs(intValuta)
			    
			    
			    if func = "dbopr" then
			   
			    if len(trim(personlig)) <> 0 then
			    personlig = personlig
			    else 
			    personlig = 0
			    end if
			   
			   
                strSQL = "INSERT INTO materiale_forbrug "_
				&" (matid, matantal, matnavn, matvarenr, matkobspris, matsalgspris, matenhed, jobid, "_
				&" editor, dato, usrid, matgrp, forbrugsdato, serviceaft, valuta, intkode, bilagsnr, kurs, personlig) VALUES "_
				&" ("& matId &", "& intAntal &", '"& strNavn &"', '"& strVarenr &"', "_
				&" "& dblKobsPris &", "& dblSalgsPris &", '"& strEnhed &"', "& jobid &", "_
				&" '"& strEditor &"', '"& strDato &"', "& usemrn &", "_
				&" "& avaGrp &", '"& regDatoSQL &"', "& aftid &", "& intValuta &", "_
				&" "& intkode &", '"& bilagsnr &"', "& dblKurs &", "& personlig &")"
				
				'Response.Write strSQL
				'Response.end
				oConn.execute(strSQL)
				
				
				
		    '*** lastid *** '
            lastId = 0
			strSQLlid = "SELECT id FROM materiale_forbrug WHERE id <> 0 ORDER BY id DESC"
			oRec.open strSQLlid, oConn, 3
			if not oRec.EOF then
			 lastId = oRec("id")
			end if
			oRec.close
						
			
			else			
				'*** Opdaterer, KUN materialer der ikke er indl�st p� lager pga lagerstyrring **'
				'*** er allerede foretaget ***'
				'*** ved indtastning fra lager kan man kun slette indtastning **'
				
				strSQL = "UPDATE materiale_forbrug SET "_
				&" matid = "& matId &", matantal = "& intAntal &", matnavn = '"& strNavn &"', "_
				&" matvarenr = '"& strVarenr &"', matkobspris = "& dblKobsPris &", matsalgspris = "& dblSalgsPris &", "_
				&" matenhed = '"& strEnhed &"', jobid = "& jobid &", "_
				&" editor = '"& strEditor &"', dato = '"& strDato &"', usrid = "& usemrn &", "_
				&" matgrp = "& avaGrp &", forbrugsdato = '"& regDatoSQL &"', serviceaft = "& aftid &", "_
				&" valuta = "& intValuta &", intkode = "& intkode &", bilagsnr = '"& bilagsnr &"', "_
				&" kurs = "& dblKurs &", personlig = "& personlig &" WHERE id = " & matregid
				
				
				
				'Response.Write strSQL
				'Response.end
				oConn.execute(strSQL)
		    
		    
		    
		    lastId = matregid
		    
		    
		    
		    end if
						
				    '**********************'
				    '** Opdaterer antal / lager status ***'
				    '**********************'		
					
				  if strVarenr <> "0" then	
						
				  '**** Flytter lager hvis der er valgt et andet lager end default ***'
			      if cint(intMatgrp) <> cint(avaGrp) then
			      call flytlager(matId, -(intAntal), avaGrp, 1)
			      else
			      strSQL2 = "UPDATE materialer SET antal = ("&intprvAntal-intAntal&") WHERE id = "& matId
				  oCOnn.execute(strSQL2)
			      end if
			      
			      end if
				
						
		
		strIndlaest = request("FM_indlaest") & " <li> " & strNavn & ": <b>" & intAntal & "</b> " & strEnhed & " <font color=limegreen><i>V</i></font>" 
		
		'Response.Write "request(FM_sog)" & request("FM_sog")
		'Response.end
		
		sogKri = replace(request("FM_sog"), "%", "wildcardprocent")
		
		Response.redirect "materialer_indtast.asp?lastid="&lastId&"&id="&jobid&"&aftid="&aftid&"&FM_indlaest="&strIndlaest&"&FM_sog="&sogKri&"&vis="&vis
	    
	    
	    end if 'dato
		end if 'antal
		
	
	
	
	case else
	
	
	
	function jobliste(vis)
	
	if len(trim(request("useFM_jobsog"))) <> 0 then
	jobsogval = request("FM_jobsog")
	
	else
	    
	    if request.Cookies("tsa")("jobsogval") <> "" then
	    jobsogval = request.Cookies("tsa")("jobsogval")
	    else
	    jobsogval = ""
	    end if
	
	end if
	
	response.Cookies("tsa")("jobsogval") = jobsogval
	
	
	if len(trim(jobsogval)) <> 0 then
	jobsogvalKri = " AND (jobnavn LIKE '"& jobsogval &"%' OR jobnr = '"& jobsogval &"' OR kkundenavn LIKE '"& jobsogval &"%')" 
	else
	jobsogvalKri = ""
	end if
	
	if level = 1 then
	
	if len(trim(request("jobsogtrue"))) <> 0 then
	    if len(trim(request("vispasogluk"))) <> 0 then
	    vispasogluk = 1
	    vispasoglukCHK = "CHECKED"
	    else
	    vispasogluk = 0
	    vispasoglukCHK = ""
	    end if
	else
	    if request.Cookies("tsa")("jobvispasluk") <> "0" then
	    vispasogluk = 1
	    vispasoglukCHK = "CHECKED"
	    else
	    vispasogluk = 0
	    vispasoglukCHK = ""
	    end if
	end if
	
	response.Cookies("tsa")("jobvispasluk") = vispasogluk
	
	else
	
	vispasogluk = 0
	vispasoglukCHK = ""
	
	end if
	
	%>
	
	
	<b><%=tsa_txt_078 & " "& tsa_txt_236 %>:</b>
	
	<%if level = 1 then%>
	<br /><input id="vispasogluk" name="vispasogluk" type="checkbox" value="1" <%=vispasoglukCHK %> /> <%=tsa_txt_354%> 
    <input id="jobsogtrue" name="jobsogtrue" value="1" type="hidden" />
	<%end if %>
	
	<br /><input id="FM_jobsog" name="FM_jobsog" type="text" value="<%=jobsogval %>" /> 
    
    &nbsp;<input id="Submit5" type="submit" value=" <%=tsa_txt_078 %> >> " /><br /><br />
	<%
	
	if cint(vispasogluk) = 1 then
	vispasoglukSQL = "OR jobstatus = 0 OR jobstatus = 2"
	else
	vispasoglukSQL = ""
	end if
	
	'*** Det valgte job ***
	strSQL = "SELECT id, jobnavn, jobnr, kkundenavn, kkundenr, kid, jobstatus FROM job j "_
	&" LEFT JOIN kunder k ON (kid = jobknr) WHERE (jobstatus = 1 "& vispasoglukSQL &") "& strPgrpSQLkri &" "& jobsogvalKri &" ORDER BY kkundenavn, jobnavn"
	
	'jobstatus = 2 = passive
	'Response.Write strSQL
	
	oRec.open strSQL, oConn, 3 
	
	if vis = "otf" then
	
	%>
	    <!-- jobid_sel -->
        <select id="jobid_sel" name="jobid_sel" style="width:440px; font-size:9px; background-color:#FFFFe1;">
           
	<%
	
	else%>
	    <!-- jobid_sel -->
        <select id="id" name="id" style="width:440px; font-size:9px; background-color:#FFFFe1;" onchange="submit();">
           
	<%end if
	
    lastknr = 0
	fj = 0
	while not oRec.EOF
	    
	    if oRec("id") = cint(id) then
	    'strValgtJob = oRec("jobnavn") & "&nbsp;("& oRec("jobnr") &")</b>"
	    jCHK = "SELECTED"
	    else
	    jCHK = ""
	    end if
	     
	     if lastknr <> oRec("kid") then
	          if cint(fj) <> 0 then %> 
	          <option value="<%=oRec("id") %>" disabled></option><%
	          end if
	         %> <option value="<%=oRec("id") %>" disabled><%=oRec("kkundenavn") %> (<%=oRec("kkundenr") %>) </option><%
	         
	
	     end if
	        
	        jstTxt = ""
	        if oRec("jobstatus") <> 1 then
	            
	            if oRec("jobstatus") = 2 then
	            jstTxt = " - " & tsa_txt_355
	            end if
	            
	             if oRec("jobstatus") = 0 then
	            jstTxt = " - " & tsa_txt_020
	            end if 
	        
	        end if
	        
	        %> <option <%=jCHK %> value="<%=oRec("id") %>"><%=oRec("jobnavn") %> (<%=oRec("jobnr") %>) <%=jstTxt%> ....................<%=oRec("kkundenavn") %> (<%=oRec("kkundenr") %>) </option><%
	
	
	lastknr = oRec("kid")
	fj = fj + 1
	oRec.movenext
	wend
	oRec.close 
	
	if fj = 0 then
	%>
	<option value=0>Ingen</option>
	<%
	end if
	%>
    </select>
    <%
    end function
	
	
	strIndlaest = request("FM_indlaest") 
	
	%>
	
	
	<script>
	function rydSog(){
	document.getElementById("FM_sog").value = ""
	}

	
	function opd_jobid(id){
	thisJobidoptSel = document.getElementById("jobid_sel").selectedIndex
	////alert(thisJobidoptSel)
	document.getElementById("jobid_"+id).value = document.getElementById("jobid_sel").options[thisJobidoptSel].value
	}
	
	
	function beregnsalgspris(matid) {
    avaId = document.getElementById("FM_avagruppe_"+matid).value 
	avaVal = document.getElementById("avagrpval_"+avaId).value.replace(",",".")
	    
	  
	kobspris = document.getElementById("FM_kobspris_"+matid).value.replace(",",".") 
	nysalgspris = ((kobspris/1 * avaVal) / 100 + kobspris/1)
	nysalgspris = Math.round(nysalgspris*100)/100
	document.getElementById("FM_salgspris_"+matid).value = nysalgspris
	document.getElementById("FM_salgspris_"+matid).value = document.getElementById("FM_salgspris_"+matid).value.replace(".",",")
	}
	
	function beregnsalgsprisOTF() {
	avaId = document.getElementById("gruppe").value 
	avaVal = document.getElementById("avagrpval_"+avaId).value.replace(",",".")
	    
	  
	kobspris = document.getElementById("pris").value.replace(",",".") 
	nysalgspris = ((kobspris/1 * avaVal) / 100 + kobspris/1)
	nysalgspris = Math.round(nysalgspris*100)/100
	document.getElementById("FM_salgspris").value = nysalgspris
	document.getElementById("FM_salgspris").value = document.getElementById("FM_salgspris").value.replace(".",",")

}

        $(document).ready(function() {
                
                $("#intkode").change(function() {
                
	            
	            var thisval = this.value
	            //alert(thisval)
	            
	            if (thisval == 2) {
	             $("#tr_slgs").css("visibility", "visible");
	            } else {
	              $("#tr_slgs").css("visibility", "hidden");
	            }
	        });
                
        });
	
	
	</script>
	
	<body>
	<div id="sidediv" style="position:absolute; left:10px; top:40px; visibility:visible; width:420px;">
	
	
	
	<%call matregmenu(0,0,Knap1_bg,Knap2_bg,Knap3_bg,Knap4_bg) %>
	
	<table cellspacing=2 cellpadding=2 border=0><tr>
	<td colspan=2>
	<%
	 oimg = "ikon_matforbrug.png"
	oleft = 0
	otop = 40
	owdt = 400
	oskrift = tsa_txt_192 
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	%>
	
	<br /><br />&nbsp
	</td>
	</tr>
	<tr><td><%=tsa_txt_237 %>:&nbsp;
	
	<%strSQLmed = "SELECT mnavn, mnr, init FROM medarbejdere WHERE mid = "& usemrn
	oRec.open strSQLmed, oConn, 3
	if not oRec.EOF then
	
	Response.Write " <b>"& oRec("mnavn") & " ("& oRec("mnr") &") - " & oRec("init") & "</b>"
	
	mednr = oRec("mnr")
	
	end if
	oRec.close %>
	</td>
	</tr>
	</table>
	
	
	<%
	'*** Bilagsnr **'
	bilagsnr = mednr &"-"& year(now) & datepart("m", now, 2,2)
	
	call hentbgrppamedarb(usemrn) %>
	
	<%if vis = "otf" then
	%>
	<%call filterheader(20,0,470,pTxt)%>
	<table cellspacing="2" cellpadding="2" border="0" width=100%>
	<form method="post" action="materialer_indtast.asp?id=<%=id %>&aftid=<%=aftid %>&fromdesk=<%=fromdesk %>&lastid=<%=lastid %>&useFM_jobsog=1">
	<tr><td>
    <%
	call jobliste(vis)
	
	%>
	</td></tr></form>
	</table>
	<!-- filter header sLut -->
	</td></tr></table>
	</div>
	<br /><br />
	
	
	<%end if %>
	

	
	
	
	
	<%if func = "red" then
	    
	    
	    
	    strSQL = "SELECT matid, matantal, matnavn, matvarenr, matkobspris, matenhed, jobid, matsalgspris, "_
	    &" matgrp, forbrugsdato, valuta, intkode, kurs, bilagsnr, personlig FROM materiale_forbrug WHERE id = " & matregid
	    
	    'Response.Write "<br><br><br>"& strSQL
	    'Response.flush
	    
	    oRec.open strSQL, oConn, 3
	    if not oRec.EOF then
	    
	    matid = oRec("matid")
	    matantal = oRec("matantal")
	    matnavn = oRec("matnavn")
	    matvarenr = oRec("matvarenr")
	    matkobspris = oRec("matkobspris")
	    matsalgspris = oRec("matsalgspris")
	    matenhed = oRec("matenhed")
	    matgrp = oRec("matgrp")
	    forbrugsdato = oRec("forbrugsdato")
	    
	    valuta = oRec("valuta")
	   
	    intKode = oRec("intkode")
	    id = oRec("jobid") 
	    bilagsnr = oRec("bilagsnr")
	    
	    personlig = oRec("personlig")
	    
	    end if
	    oRec.close
	    
	    dbfunc = "dbred"
	    
	    bdrCol = "red"
	    bgCol = "#ffffe1"
	    
	    oskrift = tsa_txt_252
	    
	
	else
	    
	    
	    matid = 0
	    matantal = 1
	   
	    matnavn = ""
	    matvarenr = 0
	    matkobspris = ""
	    matsalgspris = 0
	    matenhed = "Stk."
	    matgrp = 0
	    forbrugsdato = formatdatetime(date, 2)
	    valuta = 1 'grundvaluta
	    
	    '*** Til viderefakturering **'
	    intKode = 2
	    
	    id = id
	    bilagsnr = bilagsnr
	    dbfunc = "dbopr"
	    
	    bdrCol = "#5582d2"
	    bgCol = "#D6DFf5"
	    
	    oskrift = tsa_txt_193
	    
	    select case lto
	    case "immenso", "epi"
	    personlig = 1
	    case else
	    personlig = 0
	    end select
	    
	end if %>
	
	
	<%if vis = "otf" then %>
	
	<!-- On the fly -->
	<%
	tTop = 20
	tLeft = 0
	tWdth = 470
	
	
	call tableDiv(tTop,tLeft,tWdth)
	 %>
	 <br />
	 <h4><%=oskrift %></h4>
	<table cellspacing=0 cellpadding=2 border=0 width=100% bgcolor="#EFf3FF">
	<form name="opret" action="materialer_indtast.asp?func=<%=dbfunc%>&FM_matid=0&onthefly=1&aftid=<%=aftid%>&FM_sog=<%=sogeKri%>&fromsdsk=<%=fromsdsk%>&matregid=<%=matregid%>" method="post">
	<input type="hidden" name="jobid" id="jobid_0" value="<%=id%>">
	<tr>
		<td align=right style="border-top:1px #cccccc solid; padding:10px 0px 0px 0px;"><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_204 %>:</td>
		<td style="padding:10px 2px 0px 5px; border-top:1px #cccccc solid; "><input type="text" name="regdato_0" value="<%=forbrugsdato%>" size="10"></td>
	</tr>
	<%if lto <> "execon" AND lto <> "xintranet - local" then%>
	<tr>
		<td align=right style="padding-top:5px;"><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_202 %>:</td>
		<td style="padding-left:5px; padding-top:5px;"><input type="text" name="FM_antal" value="<%=matantal %>" size="10"> <%=tsa_txt_203 %></td>
	</tr>
	<%else %>
    <input id="hFM_antal" name="FM_antal" type="hidden" value="1" />
	<%end if%>
	<tr>
		<td align=right valign=top style="padding-top:5px;"><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_194 %>:</td>
		<td style="padding-left:5px;"><textarea id="TextArea1" name="navn" cols="30" rows="2"><%=matnavn %></textarea></td>
	</tr>
	<tr>
		<td align=right><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_201 %>:</td>
		<td style="padding-left:5px;"><input type="text" id="pris" name="pris" value="<%=matkobspris %>" size="10" onkeyup="beregnsalgsprisOTF()">&nbsp;
		<select name="FM_valuta" id="Select5" style="width:55px;">
		    <!--<option value="0"><=tsa_txt_229 %></option>-->
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
		</td>
	</tr>
	<%if lto = "execon" OR lto = "xintranet - local" then 'lto = "intranet - local" OR%>
    <input id="Hidden1" type="hidden" name="FM_betegn" value="<%=matenhed %>" />
    <input id="Hidden2" type="hidden" name="varenr" value="0" />
	<%else %>
	<tr>
		<td align=right>&nbsp;<%=tsa_txt_195 %> /<br /> <%=tsa_txt_245 %>:</td>
		<td style="padding-left:5px;"><input type="text" name="FM_betegn" value="<%=matenhed %>" size="30" maxlenght="20"></td>
	</tr>
	<tr>
		<td valign=top align=right style="padding-top:5px;"><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_196 %>:</td>
		<td valign=top style="padding-left:5px;"><input type="text" name="varenr" value="<%=matvarenr %>" size=20">
		<input id="Checkbox3" type="checkbox" name="opretlager" value="1" /> Opret p� lager
		<br /><%=tsa_txt_197 %></td>
	</tr>
	<%end if %>
	<tr>
		<td align=right>
           
		<%if level <= 2 OR level = 6 then %>
		    <%if lto <> "execon" AND lto <> "intranet - local" then  %>
		    <%=tsa_txt_198 & " "& tsa_txt_199 %>
		    <%else %>
		    <%=tsa_txt_198 %>
		    <%end if %>
		<%else %>
		<%=tsa_txt_198 %>
		<%end if %>
		:</td>
		<td style="padding-left:5px;">
		<!--onChange="visgrp()"-->
		
		<% 
		matgrpVal = "<input id=""avagrpval_0"" name=""avagrpval_0"" type=""hidden"" value=""0"" />"
        %>
		    
    		
    	<select name="gruppe" id="gruppe" style="width:200px;" onchange="beregnsalgsprisOTF()">
		<option value="0"><%=tsa_txt_200 %></option>
		<%
		strSQL = "SELECT id, navn, av FROM materiale_grp ORDER BY navn"
		oRec.open strSQL, oConn, 3 
		
		while not oRec.EOF 
		
		if cint(matgrp) = oRec("id") then
		matgrpSel = "SELECTED"
		else
		matgrpSel = ""
		end if
		
		
    		matgrpVal = matgrpVal &  "<input id=""avagrpval_"&oRec("id")&""" name=""avagrpval_"&oRec("id")&""" type=""hidden"" value="& oRec("av") &" />"
    		
		
		%>
		<option value="<%=oRec("id")%>" <%=matgrpSel %>><%=oRec("navn")%> 
		<%if level <= 2 OR level = 6 then %>
		&nbsp;(<%=oRec("av") %>%)
		<%end if %></option>
		<%
		oRec.movenext
		wend
		oRec.close %>
		</select>
		
		<%=matgrpVal %>
		
		</td>
    </tr>
	
	
	<tr><td valign=top style="padding-top:4px;" align=right><%=tsa_txt_231 %>:</td>
		    <td valign=top style="padding-left:5px;">
		   <%
			    
			    select case intKode
			    case 1
			    ikSEL0 = ""
			    ikSEL1 = "SELECTED"
			    ikSEL2 = ""
			    ikSEL3 = ""
			    
			    case 2
			    ikSEL0 = ""
			    ikSEL1 = ""
			    ikSEL2 = "SELECTED"
			    ikSEL3 = ""
			    
			    case 3
			    ikSEL0 = ""
			    ikSEL1 = ""
			    ikSEL2 = ""
			    ikSEL3 = "SELECTED"
			    
			    case else
			    
			    if lto <> "execon" AND lto <> "immenso" then
			    ikSEL0 = "SELECTED"
			    ikSEL1 = ""
			    else
			    ikSEL0 = ""
			    ikSEL1 = "SELECTED"
			    end if
			    
			    ikSEL2 = ""
			    ikSEL3 = ""
			   
			    end select
			    
			    if cint(personlig) = 1 then
			    persCHK = "CHECKED"
			    else
			    persCHK = ""
			    end if
			
			call licKid()
			
			%>
			
			<select name="FM_intkode" id="intkode" style="width:180px;">
			 <%if lto <> "execon" AND lto <> "immenso" then %>
		    <option value="0" <%=ikSEL0 %>><%=tsa_txt_235 %></option>
		    <%end if %>
		    <option value="1" <%=ikSEL1 %>><%=tsa_txt_232 %> (<%=licensindehaverKnavn%>)</option>
		    <option value="2" <%=ikSEL2 %>><%=tsa_txt_233 %> (<%=lcase(tsa_txt_243) %>)</option>
		    
		    <!--
		    <option value="3" <%=ikSEL3 %>><=tsa_txt_234 %></option>
		    -->
		    
    		</select>
    		&nbsp;
                <input id="Checkbox1" name="FM_personlig" value="1" type="checkbox" <%=persCHK %> /> <%=tsa_txt_234 %>
		    </td>
           </tr> 
	
	
	    
	    <%if (level <= 2 OR level = 6) then 
	    
	    if lto = "execon" OR lto = "xintranet - local" then
	    slgsprVZB = "hidden"
	    else
	    slgsprVZB = "visible"
	    end if 
	    
	    %>
	    <tr id="tr_slgs" style="visibility:<%=slgsprVZB%>;">
	    <td align=right><font color=red><b>*</b></font>&nbsp;<%=tsa_txt_261%>:</td>
	    <td style="padding-left:5px;"><input id="FM_salgspris" name="FM_salgspris" type="text" value="<%=matsalgspris %>" style="font-size:9px; width:80px;" /> (bruges kun ved videre fakturering)
		    
	    </td></tr>
	    <%else %>
	    <input id="FM_salgspris" name="FM_salgspris" type="hidden" value="0" />
	    <%end if %>
	    
	    
		
       <%if lto <> "execon" AND lto <> "xintranet - local" then %>    
   <tr>
    <td align=right><%=tsa_txt_230 %>: </td>
    <td style="padding-left:5px;"><input id="Text1" name="FM_bilagsnr" type="text" value="<%=bilagsnr %>" style="width:80px;" /> (<%=tsa_txt_256%>)</td>
   </tr> 
   <%end if %>       
	<tr>
	    <td colspan=2 align=right><input id="Submit4" type="submit" value=" <%=tsa_txt_330 %> >> " onclick="opd_jobid(0)" /></td>
	</tr>
	</form>
	</table>
	</div>
	
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	&nbsp;
	<%
	else
	%>
	
	<!--<div>-->
	<%
	
	useSog = 0
	
	if len(trim(request("sogmatgrp"))) <> 0 then
	sogeKri = trim(request("FM_sog"))
	response.Cookies("mat")("sogkri") = sogeKri
	useSog = 1
	else
	    if len(trim(request.Cookies("mat")("sogkri"))) <> "0" then
	    sogeKri = request.Cookies("mat")("sogkri")
	    useSog = 1
	    else
	    sogeKri = ""
	    useSog = 0
	    end if
	end if
	
	
	 if len(trim(request("sogmatgrp"))) <> 0 then 
	    sogKrimgrp = request("sogmatgrp")
	    response.Cookies("mat")("sogmgkri") = sogKrimgrp
	        
	        if request("sogmatgrp") <> "0" then
	        sogMatgrpSQLkri = " m.matgrp = " & sogKrimgrp
	        useSog = 1
	        else
    	    sogMatgrpSQLkri = " m.matgrp <> -1 "
	        end if
	        
	    else
	        if len(trim(request.Cookies("mat")("sogmgkri"))) <> 0 AND request.Cookies("mat")("sogmgkri") <> "0" then
	        sogKrimgrp = request.Cookies("mat")("sogmgkri")
	        sogMatgrpSQLkri = " m.matgrp = " & sogKrimgrp
	        useSog = 1
	        else
	        sogKrimgrp = 0
	        sogMatgrpSQLkri = " m.matgrp <> -1 "
	        end if
	    
	    end if
	
	
    if len(trim(request("soglev"))) <> 0 then 
    sogKrilev = request("soglev")
    response.Cookies("mat")("soglevkri") = sogKrilev
        
        if request("soglev") <> "0" then
        sogLevSQLkri = " AND (m.leva = " & sogKrilev & " OR m.levb = " & sogKrilev &")"
        useSog = 1
        else
        sogLevSQLkri = " AND m.leva <> -1 "
        end if
    
    else
        
         if len(trim(request.Cookies("mat")("soglevkri"))) <> 0 AND request.Cookies("mat")("soglevkri") <> 0 then
	     sogKrilev = request.Cookies("mat")("soglevkri")
	     sogLevSQLkri = " AND (m.leva = " & sogKrilev & " OR m.levb = " & sogKrilev &")"
	    useSog = 1
	    else
        sogKrilev = 0
        sogLevSQLkri = " AND m.leva <> -1 "
        end if
        
        
    end if
	
	
	
	
	response.Cookies("mat").expires = date + 60
	%>
	
	
	
	<%call filterheader(20,0,470,pTxt)%>
	<table cellspacing="2" cellpadding="2" border="0" width=100%>
	<form id="sog" name="sog" action="materialer_indtast.asp?aftid=<%=aftid%>&fromsdsk=<%=fromsdsk%>&useFM_jobsog=1&vis=<%=vis%>" method="post">
	
	<tr>
	<td>
	<%call jobliste(vis) %>
	</td>
	</tr>
	
	<tr>
	    <td><b><%=tsa_txt_263 %>:</b> <br />
	    <select id="sogmatgrp" name="sogmatgrp" style="width:200px;">
                <option value="0">Alle</option>
            <% 
	     
                   strSQL3 = "SELECT id, navn, av FROM materiale_grp WHERE id <> 0 ORDER BY navn"
           
                   oRec3.open strSQL3, oConn, 3 
                   while not oRec3.EOF 
                        
                        if cint(sogKrimgrp) = oRec3("id") then
                        smgSEL = "SELECTED"
                        else
                        smgSEL = ""
                        end if
                        
                    %>
                   <option value="<%=oRec3("id") %>" <%=smgSEL %>><%=oRec3("navn") %> (<%=oRec3("av") %>)</option>
                   <%
                   oRec3.movenext
                   wend
                   oRec3.close
            %> 
            </select>       
            </td>
	 </tr>
	 <tr>
	    <td><b><%=tsa_txt_264 %>:</b> <br />
            
            
            
            <select id="soglev" name="soglev" style="width:200px;">
                <option value="0">Alle</option>
            <% 
	     
                   strSQL3 = "SELECT id, navn FROM leverand WHERE id <> 0 ORDER BY navn"
           
                   oRec3.open strSQL3, oConn, 3 
                   while not oRec3.EOF 
                   
                        if cint(sogKrilev) = oRec3("id") then
                        slevSEL = "SELECTED"
                        else
                        slevSEL = ""
                        end if
                   
                   %>
                   <option value="<%=oRec3("id") %>" <%=slevSEL %>><%=oRec3("navn") %></option>
                   <%
                   oRec3.movenext
                   wend
                   oRec3.close
            %> 
            </select>       
            </td>
	 </tr>
	
	<tr>
		<td>
		<b>Materiale:</b><br />
		<input type="hidden" name="FM_indlaest" id="FM_indlaest" value="<%=strIndlaest%>">
		<input type="text" name="FM_sog" id="FM_sog" value="<%=sogeKri%>" style="width:200px;" onFocus="rydSog();"> &nbsp;<input id="Submit2"
                type="submit" value="<%=tsa_txt_078 %> >> " /><br />
		<%=tsa_txt_206 %><br />
		Maks. 100 resultater pr. s�gning.</td>
		</tr>
		
	
	 
	 </form>
	</table>
	<!-- filter header sLut -->
	</td></tr></table>
	</div>
	
	
	
	
	
	
	<!-- Lager -->
	

	
	
	<%if level <= 2 OR level = 6 then
	cspan = 5
	else
	cspan = 4
	end if 
	
	tTop = 60
	tLeft = 0
	tWdth = 470
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>
	
    <br />
	<h4><%=tsa_txt_207 %></h4>
	
	<table cellspacing="0" cellpadding="2" border="0" width="100%">
	<tr bgcolor="#5582D2">
		<td valign=bottom height=20 class=alt><b><%=tsa_txt_202 %></b>
		 <br /><%=tsa_txt_208 %></td>
		<td valign=bottom class=alt><b><%=tsa_txt_209 %></b>
		<br /><%=tsa_txt_210 %></td>
		
		<td align=right valign=bottom class=alt style="padding-right:6px; ">
		<b><%=tsa_txt_211 %>
		<%if level <= 2 OR level = 6 then %>
		/<%=tsa_txt_261 %>
		<%end if %></b> 
		<br />
		<%=tsa_txt_212 %>
		
	    <%if level <= 2 OR level = 6 then %>
		&nbsp;<%=tsa_txt_213 %>
		<%end if %></td>
		
		
		
		
		<td valign=bottom class=alt><b><%=tsa_txt_229 %></b><br />
		<%=tsa_txt_231 %></td>
		<td valign=bottom class=alt><b><%=tsa_txt_214 %></b><br />
		<%=tsa_txt_230 %></td>
		<td>
            &nbsp;</td>
	</tr>
	<%
	
	
	if useSog = 1 then
	sogKri = "("& sogMatgrpSQLkri &") "& sogLevSQLkri &" AND (m.navn LIKE '"& sogeKri &"%' OR m.varenr LIKE '"& sogeKri &"')"
	else
	sogKri = " m.matgrp = -1 " 'sogMatgrpSQLkri &" "& sogLevSQLkri 
	end if
	
	strSQL = "SELECT m.id, m.navn, m.varenr, m.antal, mg.navn AS gnavn, m.matgrp, "_
	&" m.enhed, m.pic, m.minlager, mg.av, m.indkobspris, m.salgspris "_
	&" FROM materialer m LEFT JOIN materiale_grp mg "_
	&" ON (mg.id = m.matgrp) WHERE "& sogKri &" AND varenr <> 0 ORDER BY m.matgrp, m.navn LIMIT 100"
	
	'Response.write strSQL
	x = 0
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	bgthis = "#ffffff"
	%>
	<form name="<%=oRec("id")%>" id="<%=oRec("id")%>" action="materialer_indtast.asp?func=dbopr&id=<%=id%>&fromsdsk=<%=fromsdsk%>&vis=<%=vis %>" method="post">
	
	
	<tr bgcolor="<%=bgthis%>">
	
		<input type="hidden" name="FM_matid" id="FM_matid" value="<%=oRec("id")%>">
		<input type="hidden" name="FM_indlaest" id="FM_indlaest" value="<%=strIndlaest%>">
		<input type="hidden" name="aftid" id="aftid" value="<%=aftid%>">
		<input type="hidden" name="jobid" id="jobid_<%=id%>" value="<%=id%>">
		<input type="hidden" name="skiftlagermsgvist" id="skiftlagermsgvist" value="0">
		<input type="hidden" name="FM_sog" id="FM_sog" value="<%=sogeKri%>">
		
		
		<td valign=top style="padding:2px; border-bottom:1px silver dashed;" class=lille><input type="text" name="FM_antal" id="FM_antal" size="2" style="font-size:9px;">
		<br /><b><%=oRec("antal")%></b>/<%=oRec("minlager") %>&nbsp;<%=oRec("enhed")%>
		</td>
		<td valign=top style="padding:2px; border-bottom:1px silver dashed; width:100px; word-wrap:break-word;"><b><a href="preview.asp?pic=<%=oRec("pic")%>&matid=<%=oRec("id")%>" class='vmenu'><%=oRec("navn")%></a></b>
		(<%=oRec("varenr")%>)</td>
        
        <%
            if level <= 2 OR level = 6 then%>
            <td valign=top style="width:110px; padding:2px; border-bottom:1px silver dashed;">
            <input id="FM_kobspris_<%=oRec("id")%>" name="FM_kobspris_<%=oRec("id")%>" type="text" value="<%=oRec("indkobspris") %>" style="font-size:9px; width:50px;" onkeyup="beregnsalgspris('<%=oRec("id")%>')" />
		     / <input id="FM_salgspris_<%=oRec("id")%>" name="FM_salgspris_<%=oRec("id")%>" type="text" value="<%=oRec("salgspris") %>" style="font-size:9px; width:50px;" />
		    
		    
		    
		    <select name="FM_avagruppe_<%=oRec("id")%>" id="FM_avagruppe_<%=oRec("id")%>" style="width:109px; font-size:9px;" onchange="beregnsalgspris('<%=oRec("id")%>')">
		    <<option value="0"><%=tsa_txt_200 %></option>
		    
		    
		    <%
		    
		    matgrpVal = "<input id=""avagrpval_0"" name=""avagrpval_0"" type=""hidden"" value=""0"" />"
    		
		    strSQL3 = "SELECT id, navn, av FROM materiale_grp ORDER BY navn"
    		
		    oRec3.open strSQL3, oConn, 3 
		    while not oRec3.EOF 
    		
		    if cint(oRec3("id")) = cint(oRec("matgrp")) then
		    matGrpCHK = "SELECTED"
		    else
		    matGrpCHK = ""
		    end if
    		
    		if x = 0 then
    		matgrpVal = matgrpVal &  "<input id=""avagrpval_"&oRec3("id")&""" name=""avagrpval_"&oRec3("id")&""" type=""hidden"" value="& oRec3("av") &" />"
    		end if
    		
		    %>
		    <option value="<%=oRec3("id")%>" <%=matGrpCHK %>><%=left(oRec3("navn"), 10)%> 
		    &nbsp;(<%=oRec3("av") %>%)
		    </option>
		    
                
		    
		    <%
		    oRec3.movenext
		    wend
		    oRec3.close %>
		    </select>
		    
		     
            <input id="Checkbox4" name="FM_opdaterpris_<%=oRec("id")%>" type="checkbox" /><%=tsa_txt_332 %>
           
		    
            </td>
            <%if x = 0 then %>
            <%=matgrpVal %>
            <%end if %> 
            
           
            
              <%else 
           
          
           
                   avaGrp = 0
                   strSQL3 = "SELECT id, navn, av FROM materiale_grp WHERE id = "& oRec("matgrp") &" ORDER BY navn"
           
                   oRec3.open strSQL3, oConn, 3 
                   if not oRec3.EOF then
                   avaGrp = oRec3("id")
                   avaGrpnavn = oRec3("navn")
                   end if
                   oRec3.close
                   %> 
            
            <td valign=top align=right style="padding:3px; padding-right:6px; border-bottom:1px silver dashed;">
            <b><%=formatnumber(oRec("indkobspris"), 2) %></b><br />
            <img src="ill/blank.gif" border=0 width=1 height=3 /><br />
            <%=avaGrpnavn %></td>
            <input id="FM_avagruppe_<%=oRec("id")%>" name="FM_avagruppe_<%=oRec("id")%>" value="<%=avaGrp%>" type="hidden" />
            <input id="FM_kobspris_<%=oRec("id")%>" name="FM_kobspris_<%=oRec("id")%>" type="hidden" value="<%=oRec("indkobspris") %>"/>
            <input id="FM_salgspris_<%=oRec("id")%>" name="FM_salgspris_<%=oRec("id")%>" type="hidden" value="<%=oRec("salgspris") %>" />
           
            
            <%end if %>
            
            
            <td valign=top style="padding:3px; border-bottom:1px silver dashed;">
            <%if level <= 2 OR level = 6 then %>
            <select name="FM_valuta_<%=oRec("id")%>" id="Select1" style="width:50px; font-size:9px;">
		    <option value="0"><%=tsa_txt_229 %></option>
		    <%
		    strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
    		
    		
		    oRec3.open strSQL3, oConn, 3 
		    while not oRec3.EOF 
    		
		    if cint(oRec3("grundvaluta")) = 1 then
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
		    <%else %>
		    <input name="FM_valuta_<%=oRec("id")%>" value="1" type="hidden" />DKK<br />
		    <%end if %>
		    
		    
		    <select name="FM_intkode_<%=oRec("id")%>" id="Select2" style="width:50px; font-size:9px;">
		    <%if lto <> "execon" AND lto <> "immenso" then %>
		    <option value="0"><%=tsa_txt_235 %></option><!-- ingen -->
		    <%end if%>
		    
		    <option value="1"><%=tsa_txt_232 %></option><!-- intern -->
		    <option value="2"><%=tsa_txt_233 %></option><!-- ekstern -->
		    
		    <!--
		    <option value="3"><%=tsa_txt_234 %></option>
		    -->
		    
		    
    		</select>
    		
                <input id="FM_personlig_<%=oRec("id")%>" name="FM_personlig_<%=oRec("id")%>" value="1" type="checkbox" /> <%=left(tsa_txt_234, 4) %>.
		    
		   
            </td>
		   
            
         
           

		<td valign=top width=90 style="padding:2px; border-bottom:1px silver dashed;">
		<input type="text" name="regdato_<%=oRec("id")%>" value="<%=formatdatetime(date, 2) %>" style="width:80px; font-size:9px;">
		<br /><input id="FM_bilagsnr_<%=oRec("id")%>" name="FM_bilagsnr_<%=oRec("id")%>" type="text" value="<%=bilagsnr %>" style="width:80px; font-family:verdana; font-size:8px;" /></td>
		<td valign=top style="padding:2px; border-bottom:1px silver dashed;">
            <input id="Submit3" type="submit" value=" >> " /></td>
	</tr>
	</form>
	<%
	x = x + 1
	oRec.movenext
	wend
    oRec.close
	
	if x = 0 then%>
	
	<tr bgcolor="#ffffff">
		
		<td height=50 colspan=6><font color="red"><b>!</b></font>&nbsp;<%=tsa_txt_215 %></td>
		
	</tr>
	<%end if%>	
	</table>
	
	
	</div>
	<%end if %><!-- vis otf eller lager -->
	
	
	
	</div>
	
	<!-- Senetest indtastede 100 -->
	
	<div style="position:absolute; left:500px; width:500px; top:15px;">
	<h4><%=tsa_txt_216 %></h4>
    </div>
	
	
	
	<%call filterheader(50,500,600,pTxt)%>
	
	
	
	<table width=100% cellspacing=0 cellpadding=0 border=0>
	
	<form action="materialer_indtast.asp?id=<%=id %>&aftid=<%=aftid %>&fromsdsk=<%=fromsdsk %>&soglastforty=1&vis=<%=vis%>" method="post">
	    <tr><td>
	    
	    <%if len(request("sogliste")) <> 0 then 
	    sogliste = request("sogliste")
	    useSog = 1
	    else
	    sogliste = ""
	    useSog = 0
	    end if
	    %>
	    
	    S�g p� bilags/job nummer:
            <input id="sogliste" name="sogliste" style="width:100px;" value="<%=sogliste %>" type="text" /> &nbsp;<input id="Submit1" type="submit" value="<%=tsa_txt_078 %> >> " /> (% = wildcard)<br />
            
            <input id="sogBilagOrJob" name="sogBilagOrJob" value="0" <%=sogBilagOrJobCHK0 %> type="radio" /> S�g i bilgsnr.<br />
            <input id="sogBilagOrJob" name="sogBilagOrJob" value="1" <%=sogBilagOrJobCHK1 %> type="radio" /> S�g p� jobnr.<br />
           
         
         
         
           <%
           if len(trim(request("soglastforty"))) <> 0 then
                if len(trim(request("showonlypers"))) <> 0 then
                showonlypers = 1
                showonlypersCHK = "CHECKED"
                else
                showonlypers = 0
                showonlypersCHK = ""    
                end if
           else
               
               if Request.Cookies("mat")("cshowonlypers") <> "" then
                    showonlypers = request.Cookies("mat")("cshowonlypers")
                    
                    if showonlypers = 1 then
                    showonlypersCHK = "CHECKED"
                    else
                    showonlypersCHK = ""
                    end if
                    
                else
                    showonlypers = 0
                    showonlypersCHK = ""
               end if
         
           end if 
           
           Response.Cookies("mat")("cshowonlypers") = showonlypers
           
           %>     
            <input id="Checkbox2" name="showonlypers" id="showonlypers" type="checkbox" <%=showonlypersCHK %> /> <%=tsa_txt_320 %>
                
	    </td></tr>
	    </form>
	</table>
	
	<!-- filter header sLut -->
	</td></tr></table>
	</div>
	
	
	
	
	
	<%if level <= 2 OR level = 6 then
	cspan = 7
	else
	cspan = 6
	end if %>
	
	
	<% 
	tTop = 65
	tLeft = 500
	tWdth = 600
	
	
	call tableDiv(tTop,tLeft,tWdth)
	%>
	
	
	<table width=100% cellspacing=0 cellpadding=2 border=0>
		
		<tr bgcolor="#5582d2">
		       
		    <td valign=bottom class=alt><b><%=tsa_txt_230 %></b><br />
		    <%=tsa_txt_231%><br />
		    <b><%=tsa_txt_236%></b><br />
		    
		    ~ <%=tsa_txt_218 %>
		    <%if level <= 2 OR level = 6 then %>
			 <%=tsa_txt_213 %>
			<%end if %>
		    </td>
		    
			<td valign=bottom width=80 class=alt><b><%=tsa_txt_183 %></b><br />
			<%=tsa_txt_077 %></td>
			
			<td valign=bottom align=right class=alt><b><%=tsa_txt_202 %></b></td>
			<td valign=bottom style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_209 %></b></td>
			<td valign=bottom align=center style="padding:2px 5px 2px 5px;" class=alt><b><%=tsa_txt_217 %></b></td>
			
			
			
			
			<td valign=bottom align=right style="padding:2px 5px 2px 5px;" class=alt>
			<b><%=tsa_txt_219 %></b><br />
			
			<%if level <= 2 OR level = 6 then %>
			<%=tsa_txt_220 %>
			<%end if %>
			
			</td>
			
			<td valign=bottom align=right style="padding:2px 5px 2px 5px;" class=alt>
			<b><%=tsa_txt_248 %></b>
			</td>
			
			
			<td valign=bottom class=alt>
			<%=left(tsa_txt_251, 3) %>.</td>
			<td valign=bottom class=alt><%=tsa_txt_221 %></td>
			
			
		</tr>
	<%
	if useSog = 1 then
        if cint(sogBilagOrJob) = 0 then
	    sqlWh = "mf.bilagsnr LIKE '"& sogliste &"'"
        else
        sqlWh = "j.jobnr LIKE '"& sogliste &"'"
        end if
	else
	sqlWh = "usrid = "& usemrn 
	end if
	
	if showonlypers = 1 then
	strSQLper = " AND personlig = 1"
	else
	strSQLper = ""
	end if
	
	strSQLmat = "SELECT m.mnavn AS medarbejdernavn, mf.id AS mfid, mf.matvarenr AS varenr, mg.navn AS gnavn, mf.matenhed AS enhed, "_
	&" mf.matnavn AS navn, mf.matantal AS antal, mf.dato, mf.editor, "_
	& "mf.matkobspris, mf.matsalgspris, mf.jobid, mf.matgrp, "_
	&" mf.usrid, mf.forbrugsdato, j.id, j.jobnr, j.jobnavn, "_
	&" mg.av, f.fakdato, k.kkundenavn, "_
	&" k.kkundenr, mf.valuta, mf.intkode, mf.bilagsnr, v.valutakode, mf.personlig "_
	&" FROM materiale_forbrug mf"_
	&" LEFT JOIN materiale_grp mg ON (mg.id = matgrp) "_
	&" LEFT JOIN medarbejdere m ON (mid = usrid) "_
	&" LEFT JOIN job j ON (j.id = mf.jobid) "_
	&" LEFT JOIN fakturaer f ON (f.jobid = mf.jobid AND f.faktype = 0) "_
	&" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
	&" LEFT JOIN valutaer v ON (v.id = mf.valuta) "_
	&" WHERE "& sqlWh &" "& strSQLper &" GROUP BY mf.id ORDER BY mf.id DESC, mf.forbrugsdato DESC, f.fakdato DESC LIMIT 100" 	
	
	'mf.jobid = "& id &"
	'response.write strSQL
	'Response.flush
	
	s = 0
	oRec.open strSQLmat, oConn, 3
	while not oRec.EOF
	 
	 
	 if len(oRec("fakdato")) <> 0 then
	 fakdato = oRec("fakdato")
	 else
	 fakdato = "01/01/2002"
	 end if
	 
	 if oRec("mfid") = cint(lastId) then
	 bgthis = "#FFFF99"
	 else
	     select case right(s, 1) 
	     case 0,2,4,6,8
	     bgthis = "#EFf3FF"
	     case else
	     bgthis = "#FFFFff"
	     end select
	 end if
	 %>
	 
	<tr bgcolor="<%=bgthis %>" class=lille>
	    <%
	    
	    useBr = ""
	    %>
	    <td valign=top style="padding:5px 2px 2px 2px; width:200px; border-bottom:1px d6dff5 solid;" class=lille>
        <%if len(oRec("bilagsnr")) <> 0 then%>
		<%=oRec("bilagsnr") %> 
		<%
		useBr = "<br />"
		end if %>
		
		<font class=megetlillesilver>
		
		<%select case oRec("intkode")
		case 0
	    intKode = "-"
		case 1
		intKode = tsa_txt_239 'intern
		case 2
		intKode = tsa_txt_240 'ekstern
		'case 3
		'intKode = tsa_txt_241
		end select %>
		
		<%if intKode <> "-" then %>
		- <%=intKode%> 
		<%
		useBr = "<br />"
		end if %>
		
		<%if oRec("personlig") <> 0 then %>
		- <%=tsa_txt_234 %>
		<%
		useBr = "<br />"
		end if %>
        
        </font>
        
        <%=useBr %><b> <%=oRec("kkundenavn") %> (<%=oRec("kkundenr") %>)</b><br />
        <%=oRec("jobnavn")%>&nbsp;(<%=oRec("jobnr")%>) 
        
         <font class="megetlillesort">
		 <%if len(oRec("gnavn")) <> 0 then %>
                    <br /> ~ <%=oRec("gnavn")%>
                    <%if level <= 2 OR level = 6 then %>
                    (<%=oRec("av") %>%)
                    <%end if %>
            <%end if %>
		</font>&nbsp;
        
            </td>
	    
		<td valign=top class=lille style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;">
		<%if len(oRec("forbrugsdato")) <> 0 then%>
		<b><%=formatdatetime(oRec("forbrugsdato"), 2)%></b><br />
		<%end if%>
		<font class="megetlillesort">
	    <%=oRec("medarbejdernavn")%></font></td>
	
		<td valign=top align=right style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid; width:75px;" class=lille><b><%=oRec("antal")%></b>&nbsp;
		
		<%if len(oRec("enhed")) <> 0 then
		enh = oRec("enhed")
		else
		enh = tsa_txt_222 '"Stk."
		end if %>
		
		<%=enh%>
		</td>
	    
	    <td valign=top style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;" class=lille><b><%=oRec("navn")%></b>&nbsp;</td>
		<td valign=top align=right style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;" class=lille><%=oRec("varenr")%></td>
			
		
		<td valign=top align=right style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;" class=lille>
	    <b><%=formatnumber(oRec("matkobspris"), 2) %></b>
        
        <%if level <= 2 OR level = 6 then %>
        <br />
		<%=formatnumber(oRec("matsalgspris"), 2) %>
        <%end if %>
		</td>
		
		<td valign=top align=right style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid; width:120px;" class=lille>
		<%kobsprisialt = formatnumber(oRec("antal") * oRec("matkobspris"), 2) %>
        <b><%=kobsprisialt %> &nbsp;<%=oRec("valutakode") %></b>
        
        <%if level <= 2 OR level = 6 then %>
        <br />
		<%salgsprisialt = formatnumber(oRec("antal") * oRec("matsalgspris"), 2) %>
        <%=salgsprisialt %> &nbsp;<%=oRec("valutakode") %>
        
		<%end if %>
		
		</td>
		
		
	    <td valign=top style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;">
	    
	    <%
	       '*** Er uge alfsuttet af medarb, er smiley og autogk sl�et til
                erugeafsluttet = instr(afslUgerMedab(usemrn), "#"&datepart("ww", oRec("forbrugsdato"),2,2)&"_"& datepart("yyyy", oRec("forbrugsdato")) &"#")
                
                'Response.Write "erugeafsluttet --" & erugeafsluttet  &"<br>"
                'Response.flush
                call lonKorsel_lukketPer(oRec("forbrugsdato"))
              
                if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) = year(now) AND DatePart("m", oRec("forbrugsdato")) < month(now)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) = 12)) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) <> 12) OR _
                (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", oRec("forbrugsdato")) > 1))) OR cint(lonKorsel_lukketIO) = 1 then
              
                ugeerAfsl_og_autogk_smil = 1
                else
                ugeerAfsl_og_autogk_smil = 0
                end if 
				
				
				
				if (ugeerAfsl_og_autogk_smil = 0 _
				OR ugeerAfsl_og_autogk_smil = 1 AND level = 1) _
				AND cdate(fakdato) < cdate(oRec("forbrugsdato")) OR (oRec("intkode") <> 2)  then 'intkode <> 2 ekstern
				
		
		'*** Kun materialer der ikke er oprettet p� laver skal kunne redigeres ***'
		'*** �ndre denne s� man v�lger ved flueben fra matreg. om det skal oprettes p� lager **'
		
		if oRec("varenr") = "0" then	%>
		<a href="materialer_indtast.asp?id=<%=oRec("id")%>&func=red&matregid=<%=oRec("mfid")%>&lastid=<%=oRec("mfid")%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aftid%>&vis=otf"><img src="../ill/ac0059-16.gif" alt="<%=tsa_txt_251 %>" border=0 /></a>&nbsp;
		<%else %>
		&nbsp;
		<%end if %>
		
		</td>
		<td valign=top style="padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;">
	    &nbsp;<a href="materialer_indtast.asp?id=<%=oRec("id")%>&func=slet&matregid=<%=oRec("mfid")%>&fromsdsk=<%=fromsdsk%>&aftid=<%=aftid%>&vis=<%=vis%>"><img src="../ill/slet_16.gif" alt="<%=tsa_txt_221 %>" border=0 /></a>
	    <%else%>
	    &nbsp;</td>
		<td>
	    &nbsp;
	    <%end if %>
	    
	    </td>
		
	</tr>
	
	<%
	s = s + 1
	oRec.movenext
	wend
	
	oRec.close
	%>
	</table>
	
	</div>
	
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	&nbsp;
	
	
	
	
	
	<div id="lagerosigt" style="position:absolute; left:10px; top:100px; width:1px; height:1px; overflow:auto; visibility:hidden; display:none; width:420px;">
	<b><%=tsa_txt_223%>:</b>  (<%=tsa_txt_224%>)<br />
	<table cellspacing=2 cellpadding=1 border=0 width="100%">
	<tr>
	    <td><b><%=tsa_txt_225%></b></td>
	    <td><b><%=tsa_txt_226%></b></td>
	    <td><b><%=tsa_txt_227%></b></td>
	    <td><b><%=tsa_txt_228%></b></td>
	</tr>
	<%
	strSQL = "SELECT mg.navn AS mgnavn, mg.nummer AS mgnr, "_
	&" m.navn AS mnavn, m.matgrp, m.varenr AS mvnr FROM materiale_grp mg "_
	&" LEFT JOIN materialer m ON (m.matgrp = mg.id) "_
	&" WHERE mg.id <> 0 GROUP BY mg.id ORDER BY m.varenr DESC LIMIT 10" 
	
	'Response.Write strSQL
	'Response.flush
	
	'oRec.open strSQL, oConn, 3
	'while not oRec.EOF 
	%>
	<tr>
	<td><=oRec("mgnavn") %></td>
	<td><=oRec("mgnr") %></td>
	<td class=lille><i><=oRec("mnavn") %></i></td>
	<td class=lille><i><=oRec("mvnr") %></i></td>
	</tr>
	
	
	<%
	
	'oRec.movenext
	'wend
	'oRec.close
	
	 
	 %></table>
	
	</div>
	<!--
	<
	
	itop = 10
	ileft = 920
	iwdt = 180
	'ihgt = 50
	'iId = "matreg"
	'idsp = ""
	'ivzb = "visible"
	'ibtop = 3000
	'ibleft = 200
	'ibwdt = 300
	'ibhgt = 200
	'ibId = "matreg_bread"
	
	call sideinfo(itop,ileft,iwdt) 
	'call sideinfoId(itop,ileft,iwdt,ihgt,iId,idsp,ivzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
	%>
	
	Der kan redigeres udl�g og materialer der ikke er oprettet p� lager og som er en del af lagerstyrringen.
	Hvis et udl�g / materiale er oprettet p� lager, er der kun mulighed for at slette.<br />
	<br />
	Der kan ikke indl�ses udl�g / materialer i lukkede perioder, med mindre det angives at udl�gget er internt (ikke skal videre faktureres)
	 
	
	</td>
	</tr>
	</table>
	</div>
	-->
	
	
	
	<%end select%>
    

<%end if%>
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
	&nbsp;
<!--#include file="../inc/regular/footer_inc.asp"-->
