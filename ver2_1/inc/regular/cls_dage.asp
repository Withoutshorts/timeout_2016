<%
public erHellig, helligdagnavn, helligdagnavnTxt, helligdageIalt 
function helligdage(tjekdennedag, show, lto, usemrn)


'response.write "year(tjekdennedag): "& year(tjekdennedag)

'** Individuelle indstillinger for helligdage intert i virksomhed ***'
'if lto = "tec" OR lto = "esn" then
    
    if len(trim(usemrn)) <> 0 AND usemrn <> 0 then 'valgt medarb på timereg. siden
    useMedidProgrp = usemrn
    else
    useMedidProgrp = session("mid")
    end if 

    call medariprogrpFn(useMedidProgrp)

    medariprogrpTxtDage = medariprogrpTxt
'else
'medariprogrpTxtDage = "-1"
'end if

'response.write "medariprogrpTxt: "& medariprogrpTxtDage



if year(tjekdennedag) >= 2018 then

    if len(trim(usemrn)) <> 0 then 'valgt medarb på timereg. siden
    useMedid = usemrn
    else
    useMedid = session("mid")
    end if 

call meStamdata(useMedid)

'National holidays

if len(trim(meCal)) <> 0 then
meCal = meCal
else
meCal = "DK"
end if

tjekdennedagSQL = year(tjekdennedag) & "-" & month(tjekdennedag) & "-" & day(tjekdennedag)
erHellig = 0
helligdagnavn = ""

strSQL8 = "SELECT nh_id, nh_country, nh_name, nh_duration, nh_date, nh_editor_date, nh_open, nh_sortorder, nh_projgrp FROM national_holidays "_
&" WHERE nh_id <> 0 AND nh_date = '"& tjekdennedagSQL &"' AND nh_country = '"& meCal &"' ORDER BY nh_country, nh_sortorder, nh_name, nh_date"

    'if session("mid") = 1 then
    'rESPONSE.WRITE strSQL8
    'Response.flush
    'end if

oRec9.open strSQL8, oConn, 3
if not oRec9.EOF then



if oRec9("nh_projgrp") <> "" then 'len(trim(oRec9("nh_projgrp"))) <> 0 AND 
nh_projgrp = 1
nh_projgrp_arr = split(oRec9("nh_projgrp"), ",")

else
nh_projgrp = 0
end if


if nh_projgrp = 0 then

    helligdagnavn = oRec9("nh_name")
    erHellig = oRec9("nh_open")

else

    prgGrpFundet = 0
     for p = 0 to UBOUND(nh_projgrp_arr) AND prgGrpFundet = 0

            if instr(trim(nh_projgrp_arr(p)), "-") = 0 then 'POSTIVE DVS kun helligdag for disse projektgrupper


                'Response.Write "medariprogrpTxtDage: "& medariprogrpTxtDage &" nh_projgrp_arr(p): " & nh_projgrp_arr(p) 

                
                if instr(medariprogrpTxtDage, "#"& trim(nh_projgrp_arr(p)) &"#") <> 0 then
                prgGrpFundet = 1
                erHellig = oRec9("nh_open")
                helligdagnavn = oRec9("nh_name")
                'positivFundet = 1
                else
                erHellig = 0
                helligdagnavn = ""
                end if


                'Response.Write " erHellig: "& erHellig & "<br>"

            else

                'NEGATIVE DVS ikke helligdag for disse projektgrupper
                nh_projgrp_arr(p) = replace(nh_projgrp_arr(p), "-", "")

            
                if instr(medariprogrpTxtDage, "#"& trim(nh_projgrp_arr(p)) &"#") <> 0 then
                erHellig = 0
                prgGrpFundet = 1
                helligdagnavn = ""
                'negativFundet = 1
                else
                erHellig = oRec9("nh_open")
                helligdagnavn = oRec9("nh_name")
                end if


            end if

     next

    

end if


end if
oRec9.close


'Response.write "HER: " & helligdagnavn


else


select case lto 'land
case "xepi_no", "xepi_bar", "xepi_sta", "xintranet - local" 'NO -- > Læg det ind i den centrale kalender, så der kune er en kalender.

    


'**** Helligdage 2013 - 20 *****
erHellig = 0

		select case year(tjekdennedag)
		case 2013
			
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1.Nytårsdag"
				erHellig = 1
				end select
				
		  		
			case 3
				select case day(tjekdennedag)
				'case 5
				'helligdagnavn = "Palmesøndag"
				'erHellig = 1
				case 28
				helligdagnavn = "Skjærtorsdag"
				erHellig = 1
				case 29
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 31 
				helligdagnavn = "Påskedag"
				erHellig = 1
				
				end select
			case 4
                
                select case day(tjekdennedag)
                case 1
				helligdagnavn = "2.Påskedag"
				erHellig = 1
               
                end select

			case 5
				
				select case day(tjekdennedag)
				case 1
				   
				helligdagnavn = "Offentlig høytidsdag"
				erHellig = 1
				    
				
				
                case 9
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1

                case 17
				  
				helligdagnavn = "Grunnlovsdag"
				erHellig = 1
				 

                case 19
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 20
				helligdagnavn = "2.Pinsedag"
				erHellig = 1

				end select
				
			
				
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
				erHellig = 1
                    

				end select
			end select
           

        case 2014
			
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1.Nytårsdag"
				erHellig = 1
				end select
				
		  		
			case 4
				select case day(tjekdennedag)
				case 17
				helligdagnavn = "Skjærtorsdag"
				erHellig = 1
				case 18
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 20 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 21
				helligdagnavn = "2.Påskedag"
				erHellig = 1
				end select
			
			case 5
				
				select case day(tjekdennedag)
				case 1
				   
				helligdagnavn = "Offentlig høytidsdag"
				erHellig = 1
				    
			    case 17
				helligdagnavn = "Grunnlovsdag"
				erHellig = 1
				
                case 29            
                helligdagnavn = "Kristi Himmelfartsdag"            
                erHellig = 1

				end select
				
			case 6
			
			
				select case day(tjekdennedag)
			    
                case 8
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 9
				helligdagnavn = "2.Pinsedag"
				erHellig = 1
	
				
			
				end select
				
				
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
				erHellig = 1
                    

				end select
			end select




        end select'yaer








'***********************************************************************************************************************************
'***************************** DK kalender *****************************************************************************************
'***********************************************************************************************************************************



case else 'DK

'**** Helligdage 2003 - 2014 *****
erHellig = 0

		select case year(tjekdennedag)
		case 2003, 2002 '(2002 dage passer ikke)
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				case 5
				helligdagnavn = "Helligtrekonger"
				erHellig = 1
				end select
			
			case 4
				select case day(tjekdennedag)
				case 13
				helligdagnavn = "Palmesøndag"
				erHellig = 1
				case 17
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 18
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 20 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 21
				helligdagnavn = "2.Påskedag"
				erHellig = 1
				end select
			case 5
				select case day(tjekdennedag)
				case 16
				helligdagnavn = "Bededag"
				erHellig = 1
				case 29
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1
				end select
			case 6
				select case day(tjekdennedag)
				case 5
				helligdagnavn = "Gr.lovsdag"
				erHellig = 1
				case 8
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 9
				helligdagnavn = "2.Pinsedag"
				erHellig = 1
				end select
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
				erHellig = 1
				end select
			end select
		case 2004
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				case 4
				helligdagnavn = "Helligtrekonger"
				erHellig = 1
				end select
			
			case 4
				select case day(tjekdennedag)
				case 4
				helligdagnavn = "Palmesøndag"
				erHellig = 1
				case 8
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 9
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 11 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 12
				helligdagnavn = "2.Påskedag"
				erHellig = 1
				end select
			case 5
				select case day(tjekdennedag)
				case 7
				helligdagnavn = "Bededag"
				erHellig = 1
				case 20
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1
				case 30
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 31
				helligdagnavn = "2.Pinsedag"
				erHellig = 1
				end select
			case 6
				select case day(tjekdennedag)
				case 5
				helligdagnavn = "Gr.lovsdag"
				erHellig = 1
				end select
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
				erHellig = 1
				end select
			end select
		
		case 2005
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				case 2
				helligdagnavn = "Helligtrekonger"
				erHellig = 1
				end select
			
			case 3
				select case day(tjekdennedag)
				case 20
				helligdagnavn = "Palmesøndag"
				erHellig = 1
				case 24
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 25
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 27 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 28
				helligdagnavn = "2.Påskedag"
				erHellig = 1
				end select
			case 4
				select case day(tjekdennedag)
				case 22
				helligdagnavn = "Bededag"
				erHellig = 1
				end select
			case 5
				select case day(tjekdennedag)
				case 5
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1
				case 15
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 16
				helligdagnavn = "2.Pinsedag"
				erHellig = 1
				end select
			case 6
				select case day(tjekdennedag)
				case 5
				helligdagnavn = "Gr.lovsdag"
				erHellig = 1
				end select
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
				erHellig = 1
				end select
			end select
		
		
		
		case 2006
		
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				end select
				
				
			case 4
				select case day(tjekdennedag)
				case 9
				helligdagnavn = "Palmesøndag"
				erHellig = 1
				case 13
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 14
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 16 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 17
				helligdagnavn = "2.Påskedag"
				erHellig = 1
				end select
			case 5
				select case day(tjekdennedag)
				case 12
				helligdagnavn = "Bededag"
				erHellig = 1
				case 25
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1
				end select
			case 6
				select case day(tjekdennedag)
				case 4
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 5
				helligdagnavn = "Grundlovsdag/2.Pinsedag"
				erHellig = 1
				end select
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
				erHellig = 1
				end select
			end select
		
		
		
		case 2007
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				end select
				
			case 4
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "Palmesøndag"
				erHellig = 1
				case 5
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 6
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 8 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 9
				helligdagnavn = "2.Påskedag"
				erHellig = 1
				end select
			case 5
				select case day(tjekdennedag)
				case 4
				helligdagnavn = "Bededag"
				erHellig = 1
				case 17
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1
				case 27
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 28
				helligdagnavn = "2.Pinsedag"
				erHellig = 1
				end select
			case 6
				select case day(tjekdennedag)
				case 5
				helligdagnavn = "Gr.lovsdag"
				erHellig = 1
				end select
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
				erHellig = 1
				end select
			end select
			
			
		case 2008
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				end select
				
			case 3
				select case day(tjekdennedag)
				case 16
				helligdagnavn = "Palmesøndag"
				erHellig = 1
				case 20
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 21
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 23 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 24
				helligdagnavn = "2.Påskedag"
				erHellig = 1
				end select
			case 4
				select case day(tjekdennedag)
				case 18
				helligdagnavn = "Bededag"
				erHellig = 1
				end select
			case 5
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1
				case 11
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 12
				helligdagnavn = "2.Pinsedag"
				erHellig = 1
				end select
			case 6
				select case day(tjekdennedag)
				case 5
				helligdagnavn = "Gr.lovsdag"
				erHellig = 1
				end select
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
				erHellig = 1
				end select
			end select
			
			
			
		case 2009
		
		
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				end select
			
				
			case 4
				select case day(tjekdennedag)
				case 5
				helligdagnavn = "Palmesøndag"
				erHellig = 1
				case 9
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 10
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 12 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 13
				helligdagnavn = "2.Påskedag"
				erHellig = 1
				end select
			case 5
				select case day(tjekdennedag)
				case 1
				    if lto = "cst" then
				    helligdagnavn = "1. Maj"
				    erHellig = 1
				    end if
				case 8
				helligdagnavn = "Bededag"
				erHellig = 1
				case 21
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1
				case 22
				    if lto = "cst" then
				    helligdagnavn = "Fridag"
				    erHellig = 1
				    end if
				case 31
				helligdagnavn = "Pinsedag"
				erHellig = 1
				end select
				
			case 6
			
			
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "2.Pinsedag"
				erHellig = 1
				case 5
				helligdagnavn = "Gr.lovsdag"
				erHellig = 1
				end select
				
				'Response.Write "her 2009" & day(tjekdennedag) &" "& helligdagnavn
				
				
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
				erHellig = 1
				end select
			end select
			
			
			
			
		 '********************************************************************
         '********************************************************************
         '********************************************************************
		   case 2010
		
		
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				end select
				
		    case 3
		        
		        select case day(tjekdennedag)
				case 28
				helligdagnavn = "Palmesøndag"
				erHellig = 1
			    end select
				
			case 4
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 2
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 4 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 5
				helligdagnavn = "2.Påskedag"
				erHellig = 1
				case 30
				helligdagnavn = "St. Bededag"
				erHellig = 1
				end select
			
			case 5
				
				select case day(tjekdennedag)
				case 1
				    if lto = "cst" then
				    helligdagnavn = "1. Maj"
				    erHellig = 1
				    end if
				case 13
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1
				case 14
				    if lto = "cst" then
				    helligdagnavn = "Fridag"
				    erHellig = 1
				    end if
				case 23
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 24
				helligdagnavn = "2.Pinsedag"
				erHellig = 1
				end select
				
			case 6
			
			
				select case day(tjekdennedag)
				case 5
				helligdagnavn = "Gr.lovsdag"
				erHellig = 1
				end select
				
				
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
				    if lto <> "dencker" then
				    erHellig = 1
                    end if
				end select
			end select
			
		
        
        '********************************************************************
        '********************************************************************
        '********************************************************************	
		case 2011
		
		
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				end select
				
		  		
			case 4
				select case day(tjekdennedag)
				case 17
				helligdagnavn = "Palmesøndag"
				erHellig = 1
				case 21
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 22
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 24 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 25
				helligdagnavn = "2.Påskedag"
				erHellig = 1
				end select
			
			case 5
				
				select case day(tjekdennedag)
				case 1
				    if lto = "cst" then
				    helligdagnavn = "1. Maj"
				    erHellig = 1
				    end if
				
				case 20
				helligdagnavn = "St. Bededag"
				erHellig = 1
				end select
				
			case 6
			
			
				select case day(tjekdennedag)
				case 2
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1
				case 3
				    if lto = "cst" then
				    helligdagnavn = "Fridag"
				    erHellig = 1
				    end if
				case 5
				helligdagnavn = "Gr.lovsdag"
				erHellig = 1
				case 12
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 13
				helligdagnavn = "2.Pinsedag"
				erHellig = 1
				end select
				
				
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
                     if lto <> "dencker" then
				    erHellig = 1
                    end if
				end select
			end select

         '********************************************************************
         '********************************************************************
         '********************************************************************
         case 2012
		
		
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				end select
				
		  		
			case 4
				select case day(tjekdennedag)
				'case 5
				'helligdagnavn = "Palmesøndag"
				'erHellig = 1
				case 5
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 6
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 8 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 9
				helligdagnavn = "2.Påskedag"
				erHellig = 1
				end select
			
			case 5
				
				select case day(tjekdennedag)
				case 1
				    if lto = "cst" OR lto = "kejd_pb" OR lto = "kejd_pb2" then
				    helligdagnavn = "1. Maj"
				    erHellig = 1
				    end if
				
				case 4
				helligdagnavn = "St. Bededag"
				erHellig = 1

                case 17
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1

                case 18
				    if lto = "cst" OR lto = "mi" Or lto = "intranet - local" then
				    helligdagnavn = "Fridag"
				    erHellig = 1
				    end if

                case 27
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 28
				helligdagnavn = "2.Pinsedag"
				erHellig = 1

				end select
				
			case 6
			
			
				select case day(tjekdennedag)
				
				
				case 5
				helligdagnavn = "Gr.lovsdag"
				
                if lto = "synergi1" then 
                
                else
				erHellig = 1
	            end if			
				
				end select
				
				
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
				case 31
				helligdagnavn = "Nytårsaften"
				    
                    if lto <> "dencker" AND lto <> "fk_bpm" AND lto <> "fk" then
				    erHellig = 1
                    end if

				end select
			end select

            '********************************************************************
            '********************************************************************
            '********************************************************************
            case 2013
		
		
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				end select
				
		  		
			case 3
				select case day(tjekdennedag)
				'case 5
				'helligdagnavn = "Palmesøndag"
				'erHellig = 1
				case 28
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 29
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 31 
				helligdagnavn = "Påskedag"
				erHellig = 1
				
				end select
			case 4
                
                select case day(tjekdennedag)
                case 1
				helligdagnavn = "2.Påskedag"
				erHellig = 1
                case 26
				helligdagnavn = "St. Bededag"
				erHellig = 1
                end select

			case 5
				
				select case day(tjekdennedag)
				case 1
				    if lto = "cst" OR lto = "kejd_pb" OR lto = "kejd_pb2" OR lto = "ngf" OR lto = "hvk_bbb" then
				    helligdagnavn = "1. Maj"
				    erHellig = 1
				    end if
				
				
                case 9
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1

                case 10
				    if lto = "cst" OR lto = "mi" then
				    helligdagnavn = "Fridag"
				    erHellig = 1
				    end if

                case 19
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 20
				helligdagnavn = "2.Pinsedag"
				erHellig = 1

				end select
				
			case 6
			
			
				select case day(tjekdennedag)
				
				
				case 5
				helligdagnavn = "Gr.lovsdag"

                if lto = "synergi1" then 
                
                else
				erHellig = 1
	            end if			

				end select
				
				
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				erHellig = 1
				case 25
				helligdagnavn = "Juledag"
				erHellig = 1
				case 26
				helligdagnavn = "2.Juledag"
				erHellig = 1
                case 27
                        
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if
                
                case 30
                    
                       
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" then
	    			helligdagnavn = "Feriefridag"    
                    erHellig = 1
                    end if
                    

				case 31
				helligdagnavn = "Nytårsaften"
				    
                   if lto <> "dencker" AND lto <> "fk_bpm" AND lto <> "fk" then
				    erHellig = 1
                    end if

				end select
			end select


            '*************************************************************************************************
            case 2014
		
		
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				erHellig = 1
				end select
				
		  		
			case 4
				select case day(tjekdennedag)
			    
                 case 14,15,16
                        
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if
                    

				case 17
				helligdagnavn = "Skærtorsdag"
				erHellig = 1
				case 18
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 20 
				helligdagnavn = "Påskedag"
				erHellig = 1
				case 21
				helligdagnavn = "2.Påskedag"
				erHellig = 1

				end select
		
              
               

			case 5
				
				select case day(tjekdennedag)
				case 1
				     if lto = "cst" OR lto = "kejd_pb" OR lto = "kejd_pb2" OR lto = "ngf" OR lto = "hvk_bbb" then
				    helligdagnavn = "1. Maj"
				    erHellig = 1
				    end if
				
                case 16
				helligdagnavn = "St. Bededag"
				erHellig = 1
				
                case 29
				helligdagnavn = "Kristi Himmelfart"
				erHellig = 1

                case 30
				    
                    if lto = "cst" OR lto = "mi" OR lto = "acc" then
				    helligdagnavn = "Fridag"
				    erHellig = 1
				    end if

                    
                        
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if
                    

              

				end select
				
			case 6
			
			
				select case day(tjekdennedag)

              
                case 5
				helligdagnavn = "Gr.lovsdag"
				
                if lto = "synergi1" OR (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati") OR lto = "acc" OR (lto = "tec" AND (instr(medariprogrpTxtDage, "#34#") = 0  OR instr(medariprogrpTxtDage, "#30#") = 0 )) then 
                erHellig = 0
                else
				erHellig = 1
	            end if			
				

                case 8
				helligdagnavn = "Pinsedag"
				erHellig = 1
				case 9
				helligdagnavn = "2.Pinsedag"
				erHellig = 1

				end select
				
				
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
                    
                    if (lto = "tec" AND ( instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 )) then
                    erHellig = 0
                    else
				    erHellig = 1
                    end if

				
				case 25
				helligdagnavn = "Juledag"
                         
                         if lto = "tec" OR lto = "esn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if
				case 26

				helligdagnavn = "2.Juledag"
				         
                         if lto = "tec" OR lto = "esn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

               

				case 31
				helligdagnavn = "Nytårsaften"
				    
                    if lto = "dencker" OR lto = "fk_bpm" OR lto = "fk" OR lto = "synergi1" OR (lto = "tec" AND ( instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 ) ) then
                    erHellig = 0
                    else
				    erHellig = 1
                    end if

				end select
			end select
		

            '*************************************************************************************************
            case 2015
		
		
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				        
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if
				

                case 2
                        
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" OR lto = "xintranet - local" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if
	            
                end select		

            case 3

                select case day(tjekdennedag)
				case 29
				helligdagnavn = "Palmesøndag"


               case 30,31
               
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if

				end select
                
		  		
			case 4
				select case day(tjekdennedag)
			    
                 case 1
                        
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if
                    

				case 2
				helligdagnavn = "Skærtorsdag"
				         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if
				case 3
				helligdagnavn = "Langfredag"
	                     if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if			        


				case 5 
				helligdagnavn = "Påskedag"
	                     if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if			        

				case 6
				helligdagnavn = "2.Påskedag"
				         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

				end select
		
              
               

			case 5
				
				select case day(tjekdennedag)
				case 1

                   
				    'if lto = "cst" OR lto = "kejd_pb" OR lto = "kejd_pb2" OR lto = "ngf" OR lto = "hvk_bbb" OR lto = "esn" OR lto = "wwf" then

                     helligdagnavn = "1. Maj / St. Bededag"
                    select case lto
                    case "tec"
                    erHellig = 0
                    case else
                    erHellig = 1
                    end select
				    'else
                    'erHellig = 0
                    'end if
				
                
                case 14
				helligdagnavn = "Kristi Himmelfart"
				        
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

                case 15

                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if

                    if lto = "mi" OR lto = "cst" OR lto = "acc" then
                    helligdagnavn = "Fridag"
				    erHellig = 1
                    end if
                    
                


                case 24
				helligdagnavn = "Pinsedag"
				
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

				case 25
				helligdagnavn = "2.Pinsedag"
				
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if      


				end select
				
			case 6
			
			
				select case day(tjekdennedag)

              
                case 5
				helligdagnavn = "Gr.lovsdag"
				
                if lto = "synergi1" OR lto = "fk" OR lto = "essens" OR (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati") OR lto = "acc" OR (lto = "tec" AND (instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 )) then 
                erHellig = 0
                else
                    
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

				
	            end if			
				

               

				end select
				
				
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				    
                    if (lto = "tec" AND ( instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 )) then
                    erHellig = 0
                    else
				    erHellig = 1
                    end if
				
                case 25
				helligdagnavn = "Juledag"
				         
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

				case 26
				helligdagnavn = "2.Juledag"
				    
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if


                  case 30
				    
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    else
                    erHellig = 0
                    end if
                


				case 31
				helligdagnavn = "Nytårsaften"
				    
                    if lto = "dencker" OR lto = "jttek" OR lto = "fk_bpm" OR lto = "fk" OR lto = "synergi1" OR (lto = "tec" AND ( instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 ) ) then
                    erHellig = 0
                    else
				    erHellig = 1
                    end if

				end select
			end select
		



         '*************************************************************************************************
            case 2016
		
		
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				        
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if
				

              
	            
                end select		



            case 3

                select case day(tjekdennedag)
				case 20
				helligdagnavn = "Palmesøndag"


               case 21,22,23
               
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if

				

                case 24
				helligdagnavn = "Skærtorsdag"
				         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if
				case 25
				helligdagnavn = "Langfredag"
	                     if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if			        


				case 27 
				helligdagnavn = "Påskedag"
	                     if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if			        

				case 28
				helligdagnavn = "2.Påskedag"
				         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if
                

                end select
		  		
			case 4
				
                    select case day(tjekdennedag)
				    case 22


		            
                    select case lto
                    case "akelius"
                    helligdagnavn = ""
                    erHellig = 0
                    case "tec"
                    helligdagnavn = "St. Bededag"
                    erHellig = 0
                    case else
                    helligdagnavn = "St. Bededag"
                    erHellig = 1
                    end select
				   
              
                    end select
               

			case 5
				
				select case day(tjekdennedag)
				
				case 1
				helligdagnavn = "1 maj"
				        
                        
                        erHellig = 1
                       
                
                case 5
				helligdagnavn = "Kristi Himmelfart"
				        
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

                case 6

                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if

                    if lto = "mi" OR lto = "cst" OR lto = "acc" OR lto = "cisu" OR lto = "adra" then
                    helligdagnavn = "Fridag"
				    erHellig = 1
                    end if
                    
                


                case 15
				helligdagnavn = "Pinsedag"
				
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

				case 16
				helligdagnavn = "2.Pinsedag"
				
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if      



                case "17"
            
                        
                        if lto = "akelius" OR lto = "epi_no" then
                        helligdagnavn = "Grunnlovsdag"
				        erHellig = 1
                        end if

				end select
				
			case 6
			
			
				select case day(tjekdennedag)

              
                case 5
				helligdagnavn = "Gr.lovsdag"
				
                if lto = "synergi1" OR lto = "fk" OR lto = "essens" OR (lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati") OR lto = "acc" OR (lto = "tec" AND (instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 )) then 
                erHellig = 0
                else
                    
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

				
	            end if			
				

               

				end select
				
				
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				    
                    if (lto = "tec" AND ( instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 )) then
                    erHellig = 0
                    else
				    erHellig = 1
                    end if
				
                case 25
				helligdagnavn = "Juledag"
				         
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

				case 26
				helligdagnavn = "2.Juledag"
				    
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

                case 27,28,29

                    if lto = "plan" then
                    helligdagnavn = "Plandag"
				    erHellig = 1
                    else
                    erHellig = 0
                    end if

                case 30
				    
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    else
                    erHellig = 0
                    end if

                    if lto = "plan" then
                    helligdagnavn = "Plandag"
				    erHellig = 1
                    else
                    erHellig = 0
                    end if


                


				case 31
				helligdagnavn = "Nytårsaften"
				    
                    if lto = "dencker" OR lto = "jttek" OR lto = "fk_bpm" OR lto = "fk" OR lto = "synergi1" OR (lto = "tec" AND ( instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 ) ) then
                    erHellig = 0
                    else
				    erHellig = 1
                    end if

				end select
			end select


            '*************************************************************************************************
            case 2017
		
		
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				        
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if
				

              
	            
                end select		



            case 4

                select case day(tjekdennedag)

				case 9
				helligdagnavn = "Palmesøndag"


               case 10,11
               
                    if instr(lto, "epi") then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if


               case 12 
                    if lto = "plan" or lto = "xintranet - local" then
                    helligdagnavn = "Plandag"
				    erHellig = 1
                    end if

                    if instr(lto, "epi") then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if
                    

                case 13
				helligdagnavn = "Skærtorsdag"
				         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

				case 14
				helligdagnavn = "Langfredag"
	                     if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if			        


				case 16 
				helligdagnavn = "Påskedag"
	                     if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if			        

				case 17
				helligdagnavn = "2.Påskedag"
				         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if
                

                end select
		  		
			case 5
				
                    select case day(tjekdennedag)

                        case 1
				        helligdagnavn = "1 maj"
				        
                        select case lto
                        case "wwf", "tec", "esn", "oko", "fk", "eniga", "adra", "hestia", "plan", "wilke", "mi", "acc"
                        erHellig = 0
                        case else
                        erHellig = 1
                        end select


				    case 12


		            
                        select case lto
                        case "akelius"
                        helligdagnavn = ""
                        erHellig = 0
                        case "tec"
                        helligdagnavn = "St. Bededag"
                        erHellig = 0
                        case else
                        helligdagnavn = "St. Bededag"
                        erHellig = 1
                        end select
				   
              
                    



                    case 25
				    helligdagnavn = "Kristi Himmelfart"
				        
                             if lto = "tec" OR lto = "xxesn" then
				             erHellig = 0
                             else
                             erHellig = 1
                             end if

                    case 26

                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if

                    if lto = "mi" OR lto = "cst" OR lto = "acc" OR lto = "cisu" OR lto = "adra" then
                    helligdagnavn = "Fridag"
				    erHellig = 1
                    end if

                    if lto = "plan" or lto = "xintranet - local" then
                    helligdagnavn = "Plandag"
				    erHellig = 1
                    end if

                    end select
               

			case 6
				
				select case day(tjekdennedag)
				
	            case 4
				helligdagnavn = "Pinsedag"
				
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if	
    
                case 5
				helligdagnavn = "2.P.dag/Gr.lovsdag"
				
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if		  



                case "17"
            
                        
                        if lto = "akelius" OR lto = "epi_no" then
                        helligdagnavn = "Grunnlovsdag"
				        erHellig = 1
                        end if

				end select
				
				
			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = "Juleaften"
				    
                    if (lto = "tec" AND ( instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 )) then
                    erHellig = 0
                    else
				    erHellig = 1
                    end if
				
                case 25
				helligdagnavn = "Juledag"
				         
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

				case 26
				helligdagnavn = "2.Juledag"
				    
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

                case 27,28,29

                    if lto = "plan" then
                    helligdagnavn = "Plandag"
				    erHellig = 1
                    else
                    erHellig = 0
                    end if

                case 30
				    
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    else
                    erHellig = 0
                    end if

                    if lto = "plan" then
                    helligdagnavn = "Plandag"
				    erHellig = 1
                    else
                    erHellig = 0
                    end if


                


				case 31
				helligdagnavn = "Nytårsaften"
				    
                    if lto = "dencker" OR lto = "jttek" OR lto = "fk_bpm" OR lto = "fk" OR lto = "synergi1" OR (lto = "tec" AND ( instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 ) ) then
                    erHellig = 0
                    else
				    erHellig = 1
                    end if

				end select
			end select
            '*** 2017 END



          

            '*************************************************************************************************
            case 20189
            '*** NY function 'Kalder helligdage
		
		
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nytårsdag"
				        
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if
				

              
	        end select		



            case 3

                select case day(tjekdennedag)

				case 25
				helligdagnavn = "Palmesøndag"


               case 26,27
               
                    if instr(lto, "epi") then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if


               case 28 
                    if lto = "plan" or lto = "xintranet - local" then
                    helligdagnavn = "Plandag"
				    erHellig = 1
                    end if

                    if instr(lto, "epi") then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if
                    

                case 29
				helligdagnavn = "Skærtorsdag"
				         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

				case 30
				helligdagnavn = "Langfredag"
	                     if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if			        

                end select

            case 4

                select case day(tjekdennedag)

				case 1 
				helligdagnavn = "Påskedag"
	                     if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if			        

				case 2
				helligdagnavn = "2.Påskedag"
				         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

                
            
                 case 27


		            
                        select case lto
                        case "akelius"
                        helligdagnavn = ""
                        erHellig = 0
                        case "tec"
                        helligdagnavn = "St. Bededag"
                        erHellig = 0
                        case else
                        helligdagnavn = "St. Bededag"
                        erHellig = 1
                        end select
				   
              
                    

                

                end select
		  		
			case 5
				
                    select case day(tjekdennedag)

                        case 1
				        helligdagnavn = "1 maj"
				        
                        select case lto
                        case "wwf", "tec", "esn", "oko", "fk", "eniga", "adra", "hestia", "plan", "wilke", "mi", "acc"
                        erHellig = 0
                        case else
                        erHellig = 1
                        end select


				    


                    case 10
				    helligdagnavn = "Kristi Himmelfart"
				        
                             if lto = "tec" OR lto = "xxesn" then
				             erHellig = 0
                             else
                             erHellig = 1
                             end if

                    case 11

                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if

                    if lto = "mi" OR lto = "cst" OR lto = "acc" OR lto = "cisu" OR lto = "adra" then
                    helligdagnavn = "Fridag"
				    erHellig = 1
                    end if

                    if lto = "plan" or lto = "xintranet - local" then
                    helligdagnavn = "Plandag"
				    erHellig = 1
                    end if

                   
                   case 20
				   helligdagnavn = "Pinsedag"
				
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if	



                    case 21
				    helligdagnavn = "2. Pinsedag"
				
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if		  



                    end select
               

			case 6
				
				select case day(tjekdennedag)
	            case 5
				helligdagnavn = "Gr.lovsdag"
				
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if		  



                case "17"
            
                        
                        if lto = "akelius" OR lto = "epi_no" then
                        helligdagnavn = "Grunnlovsdag"
				        erHellig = 1
                        end if

				end select
				
				

			case 12
				select case day(tjekdennedag)
				case 24
				helligdagnavn = " "
				    
                    if (lto = "tec" AND ( instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 )) then
                    erHellig = 0
                    else
				    erHellig = 1
                    end if
				
                case 25
				helligdagnavn = "Juledag"
				         
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

				case 26
				helligdagnavn = "2.Juledag"
				    
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if

                case 27,28

                    if lto = "plan" then
                    helligdagnavn = "Plandag"
				    erHellig = 1
                    else
                    erHellig = 0
                    end if

               


				case 31
				helligdagnavn = "Nytårsaften"
				    
                    if lto = "dencker" OR lto = "jttek" OR lto = "fk_bpm" OR lto = "fk" OR lto = "synergi1" OR (lto = "tec" AND ( instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 ) ) then
                    erHellig = 0
                    else
				    erHellig = 1
                    end if

				end select
			end select
            '*** 2018 END
       

        end select 'årstal


        end select 'lto


        end if 'if year(tjekdennedag) >= 2018 then


		'if thisfile <> "budget_bruttonetto" AND thisfile <> "budget_aar_dato.asp" AND thisfile <> "res_belaeg" AND _
		'thisfile <> "bal_real_norm_2007.asp"  AND thisfile <> "afstem_tot.asp" then
		'thisfile <> "timereg_akt_2006_afst" AND
		if show = 1 then
		Response.write helligdagnavn	
		end if	
		
        helligdagnavnTxt = helligdagnavn
		helligdagnavn = ""	

        helligdageIalt = helligdageIalt + erHellig
		
end function




public sondagIuge, mandagIuge
function sonmaniuge(dt)  
                        
                        
                        datoNu = dt
                        dagIuge = datepart("w", datoNu, 2,2)
                        select case dagIuge
                        case 1
                        add = 6
                        case 2
                        add = 5
                        case 3
                        add = 4
                        case 4
                        add = 3
                        case 5
                        add = 2
                        case 6
                        add = 1
                        case 7
                        add = 0
                        end select

                        sondagIuge = dateAdd("d", add, datoNu)
                        mandagIuge = dateAdd("d", -6, sondagIuge)

                        sondagIuge = year(sondagIuge) &"/"& month(sondagIuge) &"/"& day(sondagIuge) 
                        mandagIuge = year(mandagIuge) &"/"& month(mandagIuge) &"/"& day(mandagIuge) 



end function



   sub calender_2015

    %>
      <div class="container" style="z-index:1000; margin-right:5px; width:310px;"><!-- style="position:absolute; left:1270px; width:310px;" -->
                
                    <style>
                        #calendartable th {
                            text-align:center;
                            padding:1px;
                            word-spacing:1px;
                            border-right:hidden;
                        }
                        #calendartable td {
                            text-align:center;
                            padding:1px;
                            word-spacing:1px;
                            border-right:hidden;
                        }

                        #calendartable .diffrentMonth {
                            color:#d9d9d9;
                        }
                    </style>

                    <%
                    chosenYear = year(varTjDatoUS_man)
                    chosenMonth = month(varTjDatoUS_man)
                    chosenDate = chosenYear &"-"& chosenMonth & "-1" 
                    firstday = weekday(chosenDate, 2)

                    'response.Write "fisrtday " & firstday
                    lastday = dateadd("m", 1, chosenDate)
                    lastday = dateadd("d", -1, lastday)

                    weeksInMonth = dateDiff("d", chosenDate, DateAdd("d", 1, lastday))
                    'response.Write "days in month " &  weeksInMonth
                    weeksInMonth = weeksInMonth / 7
                    'response.Write "<br> herher " & weeksInMonth & "<br>"
                    weeksInMonth = weeksInMonth + 0.5
                    'response.Write "<br> weeksInMonth after " & weeksInMonth

                    weeksInMonth = round(weeksInMonth)

                    select case cint(chosenMonth)
                            case 1
                                calendarMonth = ugetast_txt_020
                            case 2
                                calendarMonth = ugetast_txt_021
                            case 3
                                calendarMonth = ugetast_txt_022
                            case 4
                                calendarMonth = ugetast_txt_023
                            case 5
                                calendarMonth = ugetast_txt_024
                            case 6
                                calendarMonth = ugetast_txt_025
                            case 7
                                calendarMonth = ugetast_txt_026
                            case 8
                                calendarMonth = ugetast_txt_027
                            case 9
                                calendarMonth = ugetast_txt_028
                            case 10
                                calendarMonth = ugetast_txt_029
                            case 11
                                calendarMonth = ugetast_txt_030
                            case 12
                                calendarMonth = ugetast_txt_031
                    end select
                        
                    preDate = DateAdd("m", -1, varTjDatoUS_man)
                    preNext = DateAdd("m", 1, varTjDatoUS_man)

                    preYear = DateAdd("yyyy", -1, varTjDatoUS_man)
                    nextYear = DateAdd("yyyy", 1, varTjDatoUS_man)

                    'response.Write "First Day " & firstday & " LASTDAY " & lastday & " WEEKS " & weeksInMonth
                    %>

                    <!-- <div style="white-space:nowrap;">
                    <h5 style="text-align:left;"><</h5> 
                    <h5 style="text-align:center"><%=calendarMonth %></h5>
                    <h5 style="text-align:right;">></h5>
                    </div> -->
                    
                   <!-- <table style="zoom:80%; margin-bottom:15px; width:100%">
                        <tr>
                            <th style="padding:1px;">
                                <select class="form-control input-small">
                                    <option>1</option>
                                    <option>2</option>
                                </select>
                            </th>
                            <th>
                                <select class="form-control input-small">
                                    <option>Jan</option>
                                    <option>Feb</option>
                                </select>
                            </th>
                            <th style="padding:1px;">
                                <select class="form-control input-small">
                                    <option>2020</option>
                                    <option>2020</option>
                                </select>
                            </th>
                            <th style="text-align:right"><span></span></th>
                        </tr>
                    </table> -->
                    

                    <table style="width:100%;">
                        <tr>
                            <th style="text-align:left;"><a  style="color:inherit; text-decoration:none;"href="favorit.asp?varTjDatoUS_man=<%=preYear %>"><h5><<</h5></a></th>
                            <th style="text-align:right; color:#444;"><a  style="color:inherit; text-decoration:none;"href="favorit.asp?varTjDatoUS_man=<%=preDate %>"><h5><</h5></a></th>
                            <th style="text-align:center; color:#444; width:125px;"><h5 style="font-size:110%;"><%=calendarMonth &" "& chosenYear %></h5></th>
                            <th style="text-align:left; color:#444;"><a style="color:inherit; text-decoration:none;" href="favorit.asp?varTjDatoUS_man=<%=preNext %>"><h5>></h5></a></th>
                            <th style="text-align:right;"><a  style="color:inherit; text-decoration:none;"href="favorit.asp?varTjDatoUS_man=<%=nextYear %>"><h5>>></h5></a></th>
                        </tr>
                    </table>

	                <table id="calendartable" class="table datatable" style="zoom:100%;"><!-- table-bordered -->
                        <thead>
                            <tr style="background-color:#f2f2f2">
                                <th style="visibility:hidden; border-right:inherit;"><%=favorit_txt_040 %></th>
                                <th><%=favorit_txt_033 %></th>
                                <th><%=favorit_txt_034 %></th>
                                <th><%=favorit_txt_035 %></th>
                                <th><%=favorit_txt_036 %></th>
                                <th><%=favorit_txt_037 %></th>
                                <th style="background-color:#CCCCCC"><%=favorit_txt_038 %></th>
                                <th style="background-color:#CCCCCC"><%=favorit_txt_039 %></th>
                                <th style="border-right:inherit;"><%=favorit_txt_041 %></th>
                            </tr>                           
                        </thead>

                        <tbody>

                            <%
                            '** Finder de typer der er med i det daglige timeregnskab ***'
                            call akttyper2009(2)


                            strSQLTimer = ""
                            select case lto
                            case "tec", "xintranet - local", "esn"
                            strSQLTimer = strSQLTimer &" AND (tfaktim <> 0)"
                            case else
                            strSQLTimer = strSQLTimer & "  AND (("& aty_sql_realhours &")"_
		                    & " OR (tfaktim = 30 OR tfaktim = 31 OR tfaktim = 7 OR tfaktim = 11))"
                            end select 

                            loopDate = chosenDate
                            totalHoursInMonth = 0

                            for i = 1 TO (weeksInMonth) 
                            'response.Write "NEW WEEK " & i
                            weekTotal = 0
                            %>     
                                <tr>
                                    <%for d = 1 TO 7 %>

                                        <%
                                        'Er dagen en helldigdag
                                        call helligdage(loopDate, 0, lto, usemrn)
                                        if erHellig = 1 then
                                            fontstyle = "#d9d9d9"
                                        else
                                            fontstyle = "inherit"
                                        end if

                                           if formatdatetime(now, 2) = formatdatetime(loopdate, 2) then
                                            'tdbgCol = "lightpink"
                                            fontstyle = "red"
                                            end if


                                        'Henter timer på dagen
                                        timerpaadag = 0
                                        fravarpaadag = 0
                                        focls = 0
                                        markcolor = "" '"#d9d9d9"
                                        markbox = ""
                                        
                                        strSQL = "SELECT timer, tfaktim FROM timer WHERE tmnr = "& usemrn & " AND tdato = '"& year(loopDate) &"-"& month(loopDate) &"-"& day(loopDate) &"' "& strSQLTimer
                                        'response.Write strSQL & "<br><br>"
                                        oRec.open strSQL, oConn, 3
                                        while not oRec.EOF
                                             timerpaadag = timerpaadag + oRec("timer")
                                                
                                            if focls = 0 then
                                                select case oRec("tfaktim")
                                                case 7,8,11,13,14,18,19,20,21,22,23,24,25,26,31,23,115,120,121,122
                                                markcolor = "#fff7ba"
                                                focls = 1
                                                end select
                                            end if

                                            
                                        oRec.movenext
                                        wend
                                        oRec.close
                                      
                                       
                                        if cdbl(timerpaadag) > 0 then
                                            markbox = "<div style='font-size:9px; background-color:"& markcolor &";'>"& formatnumber(timerpaadag, 2) & "</div>"
                                        else
                                            markbox = "<div style='font-size:9px; background-color:"& markcolor &"; visibility:hidden;'>"& formatnumber(timerpaadag, 2) &"</div>"
                                        end if
                                        %>


                                        <%if d = 1 then
                                            response.Write "<th style='background-color:#f2f2f2; color:#444; border-right:inherit;'>"& DatePart("ww", loopdate, 2, 2) &"</th>"
                                        end if %>

                                        <%
                                            
                                            if d = 6 OR d = 7 then
                                            tdbgCol = "#f2f2f2" '"aliceblue"
                                            else
                                            tdbgCol = ""
                                            end if

                                         
                                            
                                            if i = 1 then 
                                            
                                           %>                                

                                            <%if d >= firstday then %>
                                                <td style="background-color:<%=tdbgCol%>;"><a href="favorit.asp?varTjDatoUS_man=<%=loopDate %>" style="color:inherit; text-decoration:none;"><b><%=day(loopDate) %></b></a><br> <%=markbox %></td>
                                                <%loopDate = dateadd("d", 1, loopDate) %>
                                            <%else 
                                                timerpaadag = 0
                                                %>
                                                <td style="color:<%=fontstyle%>; background-color:<%=tdbgCol%>;">&nbsp;
                                                    <!-- <a href="favorit.asp?varTjDatoUS_man=<%=loopDate %>" style="color:inherit; text-decoration:none;"><span class="diffrentMonth"><b><%=day(dateadd("d", d, dateadd("d", -firstday ,chosenDate))) %></b></span></a> -->

                                                </td>
                                            <%end if %>

                                        <%else %>
                                        
                                        <%

                                            if loopDate > lastday then 'Next months days
                                                timerpaadag = 0
                                                response.Write "<td style='background-color:"&tdbgCol&";'></td>"
                                            else

                                                response.Write "<td style=""background-color:"&tdbgCol&";""><a href='favorit.asp?varTjDatoUS_man="& loopDate &"' style='color:"& fontstyle &"; text-decoration:none;' ><span><b>"& day(loopDate) &"</span></b></a> <br> "& markbox &"</td>"
                                            end if

                                        if i = weeksInMonth AND d = 7 AND day(loopDate) < day(lastDay) AND month(loopDate) = month(lastday) then
                                            i = i - 1
                                            'response.Write "print en uge mere " & loopDate & " LD " & lastDay & " WEKS " & weeksInMonth & " i " & i
                                        end if


                                        loopDate = dateadd("d", 1, loopDate)
                                        %>  

                                        <%end if %>

                                    <%
                                        weekTotal = weekTotal + timerpaadag
                                    next %>

                                    <td style="vertical-align:bottom; border-right:inherit;"><div style="font-size:9px;"><b><%=formatnumber(weekTotal, 2) %></b></div></td>

                                    <%
                                        totalHoursInMonth = totalHoursInMonth + weekTotal
                                    %>

                                </tr>
                            <%next %>
                        
                            <tr>
                                <td colspan="8" style="text-align:left; border-bottom:hidden;"><div style="background-color:#fff7ba; width:12px; display:inline-block;">&nbsp</div> <span style="border:0px 0px 0px 0px; width:100px; padding:3px; font-size:9px;"><%=favorit_txt_042 %></span></td>
                                <td style="border-right:inherit; border-bottom:hidden;"><div style="font-size:9px;"><b><%=formatnumber(totalHoursInMonth, 2) %></b></div></td>
                            </tr>

                            <tr>
                                <%if request.Cookies("calendar2020") = "1" then %>
                                <td colspan="9" style="text-align:left; padding-top:7px; border-top:hidden;"><span id="showcaldenderdropdown" style="border:0px 0px 0px 0px; width:100px; color:inherit; cursor:pointer;"><span style="font-size:120%;"class="fa fa-unlock"></span></span></td>
                                <%else %>
                                <td colspan="9" style="text-align:left; padding-top:7px; border-top:hidden;"><span id="showcaldenderfixed" style="border:0px 0px 0px 0px; width:100px; color:inherit; cursor:pointer;"><span style="font-size:120%;" class="fa fa-lock"></span></span></td>
                                <%end if %>
                            </tr>

                            </tbody>
                            

                            <!--
                            <tr>
                                <td colspan="4" style="background-color:#d9d9d9; font-size:9px; text-align:left;">Alm. dage</td>
                            </tr>-->
                            
                        

	                </table>
               
                     

	               	             
            </div>
<%

end sub

%>
