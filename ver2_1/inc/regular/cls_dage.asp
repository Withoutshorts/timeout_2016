<%
public erHellig, helligdagnavn, helligdagnavnTxt 
function helligdage(tjekdennedag, show, lto)


'response.write "year(tjekdennedag): "& year(tjekdennedag)

'** Individuelle indstillinger for helligdage intert i virksomhed ***'
if lto = "tec" OR lto = "esn" then
    
    if len(trim(usemrn)) <> 0 then 'valgt medarb p� timereg. siden
    useMedidProgrp = usemrn
    else
    useMedidProgrp = session("mid")
    end if 

    call medariprogrpFn(useMedidProgrp)

    medariprogrpTxtDage = medariprogrpTxt
else
medariprogrpTxtDage = "-1"
end if

'response.write "medariprogrpTxt: "& medariprogrpTxtDage

select case lto 'land
case "xepi_no", "xepi_bar", "xepi_sta", "xintranet - local" 'NO -- > L�g det ind i den centrale kalender, s� der kune er en kalender.

    


'**** Helligdage 2013 - 20 *****
erHellig = 0

		select case year(tjekdennedag)
		case 2013
			
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1.Nyt�rsdag"
				erHellig = 1
				end select
				
		  		
			case 3
				select case day(tjekdennedag)
				'case 5
				'helligdagnavn = "Palmes�ndag"
				'erHellig = 1
				case 28
				helligdagnavn = "Skj�rtorsdag"
				erHellig = 1
				case 29
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 31 
				helligdagnavn = "P�skedag"
				erHellig = 1
				
				end select
			case 4
                
                select case day(tjekdennedag)
                case 1
				helligdagnavn = "2.P�skedag"
				erHellig = 1
               
                end select

			case 5
				
				select case day(tjekdennedag)
				case 1
				   
				helligdagnavn = "Offentlig h�ytidsdag"
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
				helligdagnavn = "Nyt�rsaften"
				erHellig = 1
                    

				end select
			end select
           

        case 2014
			
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1.Nyt�rsdag"
				erHellig = 1
				end select
				
		  		
			case 4
				select case day(tjekdennedag)
				case 17
				helligdagnavn = "Skj�rtorsdag"
				erHellig = 1
				case 18
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 20 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 21
				helligdagnavn = "2.P�skedag"
				erHellig = 1
				end select
			
			case 5
				
				select case day(tjekdennedag)
				case 1
				   
				helligdagnavn = "Offentlig h�ytidsdag"
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
				helligdagnavn = "Nyt�rsaften"
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
				helligdagnavn = "1. Nyt�rsdag"
				erHellig = 1
				case 5
				helligdagnavn = "Helligtrekonger"
				erHellig = 1
				end select
			
			case 4
				select case day(tjekdennedag)
				case 13
				helligdagnavn = "Palmes�ndag"
				erHellig = 1
				case 17
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 18
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 20 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 21
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
				erHellig = 1
				end select
			end select
		case 2004
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nyt�rsdag"
				erHellig = 1
				case 4
				helligdagnavn = "Helligtrekonger"
				erHellig = 1
				end select
			
			case 4
				select case day(tjekdennedag)
				case 4
				helligdagnavn = "Palmes�ndag"
				erHellig = 1
				case 8
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 9
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 11 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 12
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
				erHellig = 1
				end select
			end select
		
		case 2005
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nyt�rsdag"
				erHellig = 1
				case 2
				helligdagnavn = "Helligtrekonger"
				erHellig = 1
				end select
			
			case 3
				select case day(tjekdennedag)
				case 20
				helligdagnavn = "Palmes�ndag"
				erHellig = 1
				case 24
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 25
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 27 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 28
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
				erHellig = 1
				end select
			end select
		
		
		
		case 2006
		
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nyt�rsdag"
				erHellig = 1
				end select
				
				
			case 4
				select case day(tjekdennedag)
				case 9
				helligdagnavn = "Palmes�ndag"
				erHellig = 1
				case 13
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 14
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 16 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 17
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
				erHellig = 1
				end select
			end select
		
		
		
		case 2007
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nyt�rsdag"
				erHellig = 1
				end select
				
			case 4
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "Palmes�ndag"
				erHellig = 1
				case 5
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 6
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 8 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 9
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
				erHellig = 1
				end select
			end select
			
			
		case 2008
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nyt�rsdag"
				erHellig = 1
				end select
				
			case 3
				select case day(tjekdennedag)
				case 16
				helligdagnavn = "Palmes�ndag"
				erHellig = 1
				case 20
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 21
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 23 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 24
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
				erHellig = 1
				end select
			end select
			
			
			
		case 2009
		
		
			erHellig = 0
			select case month(tjekdennedag)
			case 1
				
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "1. Nyt�rsdag"
				erHellig = 1
				end select
			
				
			case 4
				select case day(tjekdennedag)
				case 5
				helligdagnavn = "Palmes�ndag"
				erHellig = 1
				case 9
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 10
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 12 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 13
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
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
				helligdagnavn = "1. Nyt�rsdag"
				erHellig = 1
				end select
				
		    case 3
		        
		        select case day(tjekdennedag)
				case 28
				helligdagnavn = "Palmes�ndag"
				erHellig = 1
			    end select
				
			case 4
				select case day(tjekdennedag)
				case 1
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 2
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 4 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 5
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
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
				helligdagnavn = "1. Nyt�rsdag"
				erHellig = 1
				end select
				
		  		
			case 4
				select case day(tjekdennedag)
				case 17
				helligdagnavn = "Palmes�ndag"
				erHellig = 1
				case 21
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 22
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 24 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 25
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
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
				helligdagnavn = "1. Nyt�rsdag"
				erHellig = 1
				end select
				
		  		
			case 4
				select case day(tjekdennedag)
				'case 5
				'helligdagnavn = "Palmes�ndag"
				'erHellig = 1
				case 5
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 6
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 8 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 9
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
				    
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
				helligdagnavn = "1. Nyt�rsdag"
				erHellig = 1
				end select
				
		  		
			case 3
				select case day(tjekdennedag)
				'case 5
				'helligdagnavn = "Palmes�ndag"
				'erHellig = 1
				case 28
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 29
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 31 
				helligdagnavn = "P�skedag"
				erHellig = 1
				
				end select
			case 4
                
                select case day(tjekdennedag)
                case 1
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
				    
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
				helligdagnavn = "1. Nyt�rsdag"
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
				helligdagnavn = "Sk�rtorsdag"
				erHellig = 1
				case 18
				helligdagnavn = "Langfredag"
				erHellig = 1
				case 20 
				helligdagnavn = "P�skedag"
				erHellig = 1
				case 21
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
				    
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
				helligdagnavn = "1. Nyt�rsdag"
				        
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
				helligdagnavn = "Palmes�ndag"


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
				helligdagnavn = "Sk�rtorsdag"
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
				helligdagnavn = "P�skedag"
	                     if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if			        

				case 6
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
				    
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
				helligdagnavn = "1. Nyt�rsdag"
				        
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if
				

              
	            
                end select		



            case 3

                select case day(tjekdennedag)
				case 20
				helligdagnavn = "Palmes�ndag"


               case 21,22,23
               
                    if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" OR lto = "epi_cati" OR lto = "epi_uk" then
                    helligdagnavn = "Feriefridag"
				    erHellig = 1
                    end if

				

                case 24
				helligdagnavn = "Sk�rtorsdag"
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
				helligdagnavn = "P�skedag"
	                     if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if			        

				case 28
				helligdagnavn = "2.P�skedag"
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
				helligdagnavn = "Nyt�rsaften"
				    
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
				helligdagnavn = "1. Nyt�rsdag"
				        
                         if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if
				

              
	            
                end select		



            case 4

                select case day(tjekdennedag)

				case 9
				helligdagnavn = "Palmes�ndag"


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
				helligdagnavn = "Sk�rtorsdag"
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
				helligdagnavn = "P�skedag"
	                     if lto = "tec" OR lto = "xesn" then
				         erHellig = 0
                         else
                         erHellig = 1
                         end if			        

				case 17
				helligdagnavn = "2.P�skedag"
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
                        case "wwf", "tec", "esn", "oko", "fk", "eniga", "adra", "hestia", "plan", "wilke", "mi"
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
				helligdagnavn = "Nyt�rsaften"
				    
                    if lto = "dencker" OR lto = "jttek" OR lto = "fk_bpm" OR lto = "fk" OR lto = "synergi1" OR (lto = "tec" AND ( instr(medariprogrpTxtDage, "#34#") = 0 OR instr(medariprogrpTxtDage, "#30#") = 0 ) ) then
                    erHellig = 0
                    else
				    erHellig = 1
                    end if

				end select
			end select
            '*** 2017 END



       
        end select '�rstal


        end select 'lto

		'if thisfile <> "budget_bruttonetto" AND thisfile <> "budget_aar_dato.asp" AND thisfile <> "res_belaeg" AND _
		'thisfile <> "bal_real_norm_2007.asp"  AND thisfile <> "afstem_tot.asp" then
		'thisfile <> "timereg_akt_2006_afst" AND
		if show = 1 then
		Response.write helligdagnavn	
		end if	
		
        helligdagnavnTxt = helligdagnavn
		helligdagnavn = ""	
		
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





%>
