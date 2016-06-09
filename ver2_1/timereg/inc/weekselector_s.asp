
<!--
Bruges bl.a af:
fak_serviceaft_osigt.asp
Kør Stat. filerne.
materiale_stat.asp
joblog_timberebal.asp
joblog_timetotaler.asp
fomr.asp
stempelur.asp
fak_serviceaft_saldo.asp
webblik.asp
joblog_k.asp
erp_...asp
-->




<% 

    

select case thisfile
case "erp_oprfak_fs", "fak_serviceaft_osigt.asp", "erp_tilfakturering.asp", "erp_serviceaft_saldo.asp" '"erp_opr_faktura.asp" "erp_fak_rykker" ', "erp_opr_faktura", "erp_opr_faktura_kontojob"
useErpCookie = 1
case else
useErpCookie = 0
end select

'Response.Write "strDag: " & request("FM_start_dag")
'Response.write "thisfile: " & thisfile & "<br>"
'Response.end

if thisfile = "fak_serviceaft_osigt.asp" OR thisfile = "joblog" OR _
thisfile = "serviceaft_osigt.asp" OR thisfile = "materiale_stat.asp" OR _
thisfile = "materialer_ordrer.asp" OR thisfile = "stempelur" OR thisfile = "fomr" OR _
thisfile = "webblik_joblisten.asp" OR thisfile = "webblik_jobplanner.asp" OR thisfile = "webblik_milepale.asp" OR _
thisfile = "erp_tilfakturering.asp" OR thisfile = "webblik_tilfakturering.asp" OR _
thisfile = "erp_oprfak_fs" OR _
thisfile = "webblik.asp" OR thisfile = "joblog_k" OR thisfile = "joblog_timetotaler" OR thisfile = "erp_opr_faktura.asp" OR _
thisfile = "erp_serviceaft_saldo.asp" OR thisfile = "crmhistorik" OR thisfile = "smileystatus.asp" OR thisfile = "stopur_2008.asp" OR thisfile = "stat_korsel" then
	if len(request("FM_usedatokri")) <> 0 then
	useDatofilter = 1
	else
	useDatofilter = 0
	end if
else
	if thisfile = "joblog_timberebal" OR thisfile = "erp_fak_rykker" _
	OR (thisfile = "bal_real_norm_2007.asp" AND media <> "print" AND media <> "export") then 'OR thisfile = "erp_opr_faktura"
	useDatofilter = 1
	else
	useDatofilter = cint(request("FM_usedatointerval"))
	end if
end if

'Response.write thisfile & "<br>"
'Response.write "useDatofilter" & useDatofilter

'*** Bladre funktion '** Bal_real_norm
addDagVal = 0

'**********************************************************
'** ER DER SUBMITTED FORM ELLER SKA LDER BRUGES COOKIE ***'
'**********************************************************

if cint(useDatofilter) = 1 then

        if (thisfile = "joblog" AND (cint(bruguge) = 1 OR cint(brugmd) = 1 OR cint(joblog_uge) = 3)) then
        

                    
                    if cint(bruguge) = 1 OR cint(brugmd) = 1 then
      
                    strMrd = stMduge
                    strDag = stDaguge
                    strAar = stAaruge
                    strDag_slut = slDaguge
                    strMrd_slut = slMduge
                    strAar_slut = slAaruge

                    'Response.write "HER HER 1: " & strDag_slut


                    else 'joblog_uge) = 3)
     
                                strMrd = request("FM_start_mrd")
                                strDag = 1
                                strAar = right(request("FM_start_aar"),2) 
                    
                  
                
                                slDatoBeregn = dateAdd("m", 1, strDag&"/"&strMrd&"/"&strAar)                
                                slDatoBeregn = dateAdd("d", -1, slDatoBeregn)

                                strDag_slut = datepart("d", slDatoBeregn, 2,2)
                                strMrd_slut = strMrd
                                strAar_slut = strAar

        
                    end if

        else
            
            
            if thisfile = "bal_real_norm_2007.asp" then
            
                if len(request("FM_usedatokri")) <> 0 then
                
                    strMrd = request("FM_start_mrd")
                    strDag = request("FM_start_dag")
                    strAar = right(request("FM_start_aar"),2) 
                    strDag_slut = request("FM_slut_dag")
                    strMrd_slut = request("FM_slut_mrd")
                    strAar_slut = right(request("FM_slut_aar"),2)

                    addDagVal = request("addDagVal") 'bladre funktion


                   
                
                else

                    select case lto 
                    case "xx"
                    useDatoStKri = lastLonper
                    case else
    
                        if dateDiff("d", lastLonper, now, 2,2) > 90 then 'må formode at de ikke afslutter lønperide så, derfor løbende stdato der følger med dd. 30 dage 
                        useDatoStKri = dateAdd("d", -30, now)
                        else
                        useDatoStKri = lastLonper
                        end if                

                    end select

                    
                    
                    strMrd = datepart("m",useDatoStKri, 2,2)
                    strDag = datepart("d",useDatoStKri, 2,2)
                    strAar = right(datepart("yyyy", useDatoStKri, 2,2),2) 
                    strDag_slut = day(now)
                    strMrd_slut = month(now)
                    strAar_slut = right(year(now),2)
                
                end if
            
            else
            
            strMrd = request("FM_start_mrd")
            strDag = request("FM_start_dag")
            strAar = right(request("FM_start_aar"),2) 
            strDag_slut = request("FM_slut_dag")
            strMrd_slut = request("FM_slut_mrd")
            strAar_slut = right(request("FM_slut_aar"),2)
            end if
        
        
        end if

else '** useDatofilter '** ER DER SUBMITTED FORM ELLER SKA LDER BRUGES COOKIE ***'
    
   
	'*Brug cookie eller dagsdato?
	if (len(Request.Cookies("datoer")("st_md")) <> 0 AND useErpCookie = 0) OR (len(Request.Cookies("erp_datoer")("st_md")) <> 0 AND useErpCookie = 1) then
	'if Request.Cookies("st_dag").haskeys then	
		
		        
                
                
                
          if useErpCookie = 1 then
	      'Response.Write "her "&  Request.Cookies("erp_datoer")("sl_dag") & "-"& Request.Cookies("erp_datoer")("sl_md") &" :: "&  useErpCookie
    
    	        'cval = 0
		        '***********************************************************************
                '*** ERP Delen *********************************************************
		        '***********************************************************************

                
                    if (print <> "j" AND media <> "print" AND media <> "eksport") AND thisfile = "erp_tilfakturering.asp" then 

                    select case lto 
                    case "synergi1", "intranet - local", "essens"

                  

                    '** Maks 3 md

                    'nu_minusMD = dateAdd("d", -92, now)

                    'startDatoMd = month(nu_minusMD)
                    'startDatoDag = day(nu_minusMD)
                    'startDatoAar = year(nu_minusMD)

                    
                    'response.end
                    call licensStartDato()
                    'response.write startDatoDag &"-"& startDatoMd &"-"& startDatoAar

                        if dateDiff("d", licensstdato, now, 2,2) > 90 then 'mere dd. 90 dage 
                        useDatoStKri = dateAdd("d", -60, now)

                           strDag = day(useDatoStKri)
                            strMrd = month(useDatoStKri)
                            strAar = year(useDatoStKri)

                        else 'listartDato

                                 strMrd = startDatoMd
		                        strDag = startDatoDag
		                        strAar = right(startDatoAar, 2)
                       
                        end if     

                       

               
		        
		            strDag_slut = day(now)
                    strMrd_slut = month(now)
                    strAar_slut = right(year(now), 2)


                    case else

                  
		                          strMrd = Request.Cookies("erp_datoer")("st_md")
		                        strDag = Request.Cookies("erp_datoer")("st_dag")
		                        strAar = Request.Cookies("erp_datoer")("st_aar") 
        		    
		     
	                            strDag_slut = Request.Cookies("erp_datoer")("sl_dag")
	                            strMrd_slut = Request.Cookies("erp_datoer")("sl_md")
	                            strAar_slut = Request.Cookies("erp_datoer")("sl_aar")
		    
               

		            end select
            

                    else' media

                            

                             strMrd = Request.Cookies("erp_datoer")("st_md")
		                        strDag = Request.Cookies("erp_datoer")("st_dag")
		                        strAar = Request.Cookies("erp_datoer")("st_aar") 
        		    
		     
	                            strDag_slut = Request.Cookies("erp_datoer")("sl_dag")
	                            strMrd_slut = Request.Cookies("erp_datoer")("sl_md")
	                            strAar_slut = Request.Cookies("erp_datoer")("sl_aar")


                    end if

		       
		        
		    
		     
                    
                     if print <> "j" AND media <> "print" AND media <> "eksport" then 
                    '** sikrer interval ikke er større end 2 år til at starte med **'
                    if dateDiff("d", strDag&"/"&strMrd&"/"&strAar, strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2,2) > 730 then

                    nyStDato = dateAdd("d", -730, strDag_slut&"/"&strMrd_slut&"/"&strAar_slut)

                    strDag = day(nyStDato)
                    strMrd = month(nyStDato)
                    strAar = right(year(nyStDato), 2)


                    end if
                    end if
                    
                    
                                


		        
		else 'useErpCookie
        '**********************************************************************************
        '**** TSA delen, vælg start og slutdato er udfra cookie eller licensstartdato ****'
        '**********************************************************************************

        
        if print <> "j" AND media <> "print" AND media <> "eksport" then 

        select case lto 
        case "synergi1", "xintranet - local"

        call licensStartDato()


         '** Maks 3 md

        'nu_minusMD = dateAdd("m", -3, now)
        'startDatoMd = datePart("m", nu_minusMD, 2,2)
        'startDatoDag = datePart("d", nu_minusMD, 2,2)
        'startDatoAar = datePart("yyyy", nu_minusMD, 2,2)

        strMrd = startDatoMd
		strDag = startDatoDag
		strAar = startDatoAar
		        
		strDag_slut = day(now)
        strMrd_slut = month(now)
        strAar_slut = year(now)


        case else

                  
		            strMrd = Request.Cookies("datoer")("st_md")
		            strDag = Request.Cookies("datoer")("st_dag")
		            strAar = Request.Cookies("datoer")("st_aar") 
		        
		            strDag_slut = Request.Cookies("datoer")("sl_dag")
                    strMrd_slut = Request.Cookies("datoer")("sl_md")
                    strAar_slut = Request.Cookies("datoer")("sl_aar")

               

		end select


        '** sikrer interval ikke er større end 2 år til at starte med **'
        if dateDiff("d", strDag&"/"&strMrd&"/"&strAar, strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 2,2) > 730 then

        nyStDato = dateAdd("d", -730, strDag_slut&"/"&strMrd_slut&"/"&strAar_slut)

        strDag = day(nyStDato)
        strMrd = month(nyStDato)
        strAar = right(year(nyStDato), 2)

        end if
        
        else 'media Print / PDF mm
        
        strMrd = Request.Cookies("datoer")("st_md")
		strDag = Request.Cookies("datoer")("st_dag")
		strAar = Request.Cookies("datoer")("st_aar") 
		        
		strDag_slut = Request.Cookies("datoer")("sl_dag")
        strMrd_slut = Request.Cookies("datoer")("sl_md")
        strAar_slut = Request.Cookies("datoer")("sl_aar")

        end if 'media
        
        
        end if 'useErpCookie slut
		
		'StrTdato = date-31
		'StrUdato = date 
		

     


	else 

        ''********************************
        ''*** ved submit brug nyt interval
	    ''*********************************


		StrTdato = date-31
		StrUdato = date 
		
		'* Bruges i weekselector *
		if month(now()) = 1 then
		strMrd = 12
		else
		strMrd = month(now()) - 1
		end if
		
		strDag = day(now())
		
		if month(now()) = 1 then
		strAar = right(year(now()) - 1, 2)
		else
		strAar = right(year(now()), 2) 
		end if
		
		strMrd_slut = month(now())
		strAar_slut = right(year(now()), 2) 
		
        'Response.write "strDag: "& strDag & "thisfile: "& thisfile 
        
		'* SKAL IKKE VÆLGE NY dato på OPR FAKTURA, da den altid skal stå til DD som faktura dato
        if strDag > 28 AND thisfile <> "erp_oprfak_fs" then
		strDag_slut = 1
		strMrd_slut = strMrd_slut + 1
		else
		strDag_slut = day(now())
		end if
	
	end if
	
end if 'useDatofilter 0/1

                    
                    
                    
                    
                    
                    '***********************
                    '** Bladre funktion ***'
                    '***********************
                    if thisfile = "bal_real_norm_2007.asp" then
                    
                    sub bladrePlus
                
              
				            strDag = 1
                            strMrd = strMrd + 1
                   
                           if strMrd > 12 then
                           strMrd = 1
                           strAar = strAar + 1
                           end if
                      
              

                    end sub

                    sub bladrePlus_slut
                
              
				            strDag_slut = 1
                            strMrd_slut = strMrd_slut + 1
                   
                           if strMrd_slut > 12 then
                           strMrd_slut = 1
                           strAar_slut = strAar_slut + 1
                           end if
                      
              

                    end sub

            

                    strDag = strDag/1 + (addDagVal/1) '** Bladre
                    '**** Bladre ***'

                    if strDag < 1 then

                    strMrd = strMrd/1 - (1/1)
            
                        if strMrd < 1 then
                            strMrd = 12
                            strAar = strAar/1 - 1
                            strDag = 31
                        else 

                            Select case strMrd
			                case 4, 6, 9, 11
                                strDag = 30 
				

			                case 2
				
				                    select case strAar
				                    case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
				                       strDag = 29 
				                    case else
				                       strDag = 28 
				                    end select
				
				
			                case else
			                 strDag = 31
				
			                end select

                        end if

                    else

                    Select case strMrd
			        case 4, 6, 9, 11
                        if strDag > 30 then
				        call bladrePlus
                        end if

			        case 2
			
				            select case strAar
				            case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
				                if strDag > 29 then
				                call bladrePlus
                                end if
				            case else
				                if strDag > 28 then
				                call bladrePlus
                                end if
				            end select
				
			
			        case else
			         if strDag > 31 then
				        call bladrePlus
                        end if
			        end select


                    end if '<1
                    '****


                        strDag_slut = strDag_slut/1 + (addDagVal/1) '** Bladre
                        '**** Bladre ***'

                        if strDag_slut < 1 then

                        strMrd_slut = strMrd_slut/1 - (1/1)
            
                            if strMrd_slut < 1 then
                                strMrd_slut = 12
                                strAar_slut = strAar_slut/1 - 1
                                strDag_slut = 31
                            else 

                                Select case strMrd_slut
			                    case 4, 6, 9, 11
                                    strDag_slut = 30 
				

			                    case 2
				
				                        select case strAar_slut
				                        case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
				                           strDag_slut = 29 
				                        case else
				                           strDag_slut = 28 
				                        end select
				
				
			                    case else
			                     strDag_slut = 31
				
			                    end select

                            end if

                        else

                        Select case strMrd_slut
			            case 4, 6, 9, 11
                            if strDag_slut > 30 then
				            call bladrePlus_slut
                            end if

			            case 2
				
				                select case strAar_slut
				                case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
				                    if strDag_slut > 29 then
				                    call bladrePlus_slut
                                    end if
				                case else
				                    if strDag_slut > 28 then
				                    call bladrePlus_slut
                                    end if
				                end select
				
				
			            case else
			             if strDag_slut > 31 then
				            call bladrePlus_slut
                            end if
			            end select

                        end if '** < 1 _slut

                        '****
                        end if '** Bladre funktion slut
                        '******************************************************************'

                    




            '**** Finder Dato ****'

            '*** timetotaler opdelt på md ***'
            if thisfile = "joblog_timetotaler" AND cint(md_split) = 1 then '**Opdel 3 md

                select case strMrd_slut
                case 1,3,5,7,8,10,12
                strDag_slut = "31"
                case 4,6,9,11
                strDag_slut = "30"
                case 2
                     select case strAar_slut
				    case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
				    strDag_slut = 29
				    case else
				    strDag_slut = 28
				    end select
                end select

                 datoStart = dateadd("m", -2, "1/"&strMrd_slut&"/"&strAar_slut)
                 strDag = 1
                 strMrd = month(datoStart)
                 strAar = right(year(datoStart), 2)

            end if


             if thisfile = "joblog_timetotaler" AND cint(md_split) = 2 then '** Opdel 12 md

                select case strMrd_slut
                case 1,3,5,7,8,10,12
                strDag_slut = "31"
                case 4,6,9,11
                strDag_slut = "30"
                case 2
                     select case strAar_slut
				    case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
				    strDag_slut = 29
				    case else
				    strDag_slut = 28
				    end select
                end select

                 datoStart = dateadd("m", -11, "1/"&strMrd_slut&"/"&strAar_slut)
                 strDag = 1
                 strMrd = month(datoStart)
                 strAar = right(year(datoStart), 2)

            end if


          


            '**** Tjekker dag findes i md med mindre end 31 dage
			Select case strMrd
			case 4, 6, 9, 11
				if strDag = 31 then
				strDag = 30
                else
				strDag = strDag
				end if
			case 2
				if strDag > 28 then
				    select case right(strAar, 2)
				    case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
				    strDag = 29
				    case else
				    strDag = 28
				    end select
				else
				strDag = strDag
				end if
			case else
			strDag = strDag
			end select
			
			
			Select case strMrd_slut
			case 4, 6, 9, 11
				if strDag_slut = 31 then
				strDag_slut = 30
				else
				strDag_slut = strDag_slut
				end if
			case 2
				if strDag_slut > 28 then
				    
				    select case right(strAar_slut, 2)
				    case "04","08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60"
				    strDag_slut = 29
				    case else
				    strDag_slut = 28
				    end select
				    
				
				else
				strDag_slut = strDag_slut
				end if
			case else
			strDag_slut = strDag_slut
			end select
            '******************************************'







'** Indsætter cookie **

'** tjek '**
if strMrd_slut > 12 then
strMrd_slut = 1
end if

if useErpCookie = 1 then
Response.Cookies("erp_datoer")("st_dag") = strDag
Response.Cookies("erp_datoer")("st_md") = strMrd
Response.Cookies("erp_datoer")("st_aar") = strAar
Response.Cookies("erp_datoer")("sl_dag") = strDag_slut
Response.Cookies("erp_datoer")("sl_md") = strMrd_slut
Response.Cookies("erp_datoer")("sl_aar") = strAar_slut
Response.Cookies("erp_datoer").Expires = date + 1	
else
Response.Cookies("datoer")("st_dag") = strDag
Response.Cookies("datoer")("st_md") = strMrd
Response.Cookies("datoer")("st_aar") = strAar
Response.Cookies("datoer")("sl_dag") = strDag_slut
Response.Cookies("datoer")("sl_md") = strMrd_slut
Response.Cookies("datoer")("sl_aar") = strAar_slut
Response.Cookies("datoer").Expires = date + 1		
end if


'if thifile = "erp_opr_faktura_kontojob" then
'fld_nameids = "A"
'else
'fld_nameids = ""
'end if


%>			
<!--#include file="dato2_b.asp"-->
<%
if len(trim(strAar)) = 2 then
    strAar = "20"& strAar
end if


if len(trim(strAar_slut)) = 2 then
    strAar_slut = "20"& strAar_slut
end if


    


if dontshowDD <> "1" then
%>



<input type="text" name="FM_stdato" id="FM_stdato" value="<%=strDag%>.<%=strMrd%>.<%=strAar%>" style="display:none; margin-right:5px; width:80px;" />

        <input name="FM_start_dag" id="FM_start_dag" type="hidden" value="<%=strDag%>" />
        <input name="FM_start_mrd" id="FM_start_mrd" type="hidden" value="<%=strMrd%>" />
        <input name="FM_start_aar" id="FM_start_aar" type="hidden" value="<%=strAar%>" />

&nbsp;til&nbsp;
<input type="text" name="FM_sldato" id="FM_sldato" value="<%=strDag_slut%>.<%=strMrd_slut%>.<%=strAar_slut%>" style="display:none; margin-right:5px; width:80px;" />
		
		<input type="hidden" name="FM_slut_dag" id="FM_slut_dag" value="<%=strDag_slut%>">
		<input type="hidden" name="FM_slut_mrd" id="FM_slut_mrd" value="<%=strMrd_slut%>">
		<input type="hidden" name="FM_slut_aar" id="FM_slut_aar" value="<%=strAar_slut%>">






				
<!-- Weekselecter SLUT -->
<%end if %>
