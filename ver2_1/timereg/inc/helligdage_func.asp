<%
'public erHellig, helligdagnavn 
function xhelligdage(tjekdennedag, show)




'**** Helligdage 2003 - 2009 *****
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
				    if lto = "cst" then
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
				    if lto = "cst" then
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

		
		end select
		
		'if thisfile <> "budget_bruttonetto" AND thisfile <> "budget_aar_dato.asp" AND thisfile <> "res_belaeg" AND _
		'thisfile <> "bal_real_norm_2007.asp"  AND thisfile <> "afstem_tot.asp" then
		'thisfile <> "timereg_akt_2006_afst" AND
		if show = 1 then
		Response.write helligdagnavn	
		end if	
		
		helligdagnavn = ""	
		
end function




%>
