<%
function helligdage(tjekdennedag)
'**** Bibelske Helligdage 2003 - 2007 *****
		select case year(tjekdennedag)
		case 2003
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				Response.write "1. Nyt�rsdag"
				case 5
				Response.write "Helligtrekonger"
				end select
			case 3
				select case day(tjekdennedag)
				case 2
				Response.write "Fastelavn"
				end select
			case 4
				select case day(tjekdennedag)
				case 13
				Response.write "Palmes�ndag"
				case 17
				Response.write "Sk�rtorsdag"
				case 18
				Response.write "Langfredag"
				case 20 
				Response.write "P�skedag"
				case 21
				Response.write "2.P�skedag"
				end select
			case 5
				select case day(tjekdennedag)
				case 16
				Response.write "Bededag"
				case 29
				Response.write "Kristi Himmelfart"
				end select
			case 6
				select case day(tjekdennedag)
				case 5
				Response.write "Gr.lovsdag"
				case 8
				Response.write "Pinsedag"
				case 9
				Response.write "2.Pinsedag"
				end select
			case 12
				select case day(tjekdennedag)
				case 24
				Response.write "Juleaften"
				case 25
				Response.write "Juledag"
				case 26
				Response.write "2.Juledag"
				case 31
				Response.write "Nyt�rsaften"
				end select
			end select
		case 2004
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				Response.write "1. Nyt�rsdag"
				case 4
				Response.write "Helligtrekonger"
				end select
			case 2
				select case day(tjekdennedag)
				case 22
				Response.write "Fastelavn"
				end select
			case 4
				select case day(tjekdennedag)
				case 4
				Response.write "Palmes�ndag"
				case 8
				Response.write "Sk�rtorsdag"
				case 9
				Response.write "Langfredag"
				case 11 
				Response.write "P�skedag"
				case 12
				Response.write "2.P�skedag"
				end select
			case 5
				select case day(tjekdennedag)
				case 7
				Response.write "Bededag"
				case 20
				Response.write "Kristi Himmelfart"
				case 30
				Response.write "Pinsedag"
				case 31
				Response.write "2.Pinsedag"
				end select
			case 6
				select case day(tjekdennedag)
				case 5
				Response.write "Gr.lovsdag"
				end select
			case 12
				select case day(tjekdennedag)
				case 24
				Response.write "Juleaften"
				case 25
				Response.write "Juledag"
				case 26
				Response.write "2.Juledag"
				case 31
				Response.write "Nyt�rsaften"
				end select
			end select
		case 2005
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				Response.write "1. Nyt�rsdag"
				case 2
				Response.write "Helligtrekonger"
				end select
			case 3
				select case day(tjekdennedag)
				case 6
				Response.write "Fastelavn"
				end select
			case 3
				select case day(tjekdennedag)
				case 20
				Response.write "Palmes�ndag"
				case 24
				Response.write "Sk�rtorsdag"
				case 25
				Response.write "Langfredag"
				case 27 
				Response.write "P�skedag"
				case 28
				Response.write "2.P�skedag"
				end select
			case 4
				select case day(tjekdennedag)
				case 22
				Response.write "Bededag"
				end select
			case 5
				select case day(tjekdennedag)
				case 5
				Response.write "Kristi Himmelfart"
				case 15
				Response.write "Pinsedag"
				case 16
				Response.write "2.Pinsedag"
				end select
			case 6
				select case day(tjekdennedag)
				case 5
				Response.write "Gr.lovsdag"
				end select
			case 12
				select case day(tjekdennedag)
				case 24
				Response.write "Juleaften"
				case 25
				Response.write "Juledag"
				case 26
				Response.write "2.Juledag"
				case 31
				Response.write "Nyt�rsaften"
				end select
			end select
		case 2006
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				Response.write "1. Nyt�rsdag"
				end select
			case 2
				select case day(tjekdennedag)
				case 26
				Response.write "Fastelavn"
				end select
			case 4
				select case day(tjekdennedag)
				case 9
				Response.write "Palmes�ndag"
				case 13
				Response.write "Sk�rtorsdag"
				case 14
				Response.write "Langfredag"
				case 16 
				Response.write "P�skedag"
				case 17
				Response.write "2.P�skedag"
				end select
			case 5
				select case day(tjekdennedag)
				case 12
				Response.write "Bededag"
				case 25
				Response.write "Kristi Himmelfart"
				end select
			case 6
				select case day(tjekdennedag)
				case 4
				Response.write "Pinsedag"
				case 5
				Response.write "Grundlovsdag/2.Pinsedag"
				end select
			case 12
				select case day(tjekdennedag)
				case 24
				Response.write "Juleaften"
				case 25
				Response.write "Juledag"
				case 26
				Response.write "2.Juledag"
				case 31
				Response.write "Nyt�rsaften"
				end select
			end select
		case 2007
			select case month(tjekdennedag)
			case 1
				select case day(tjekdennedag)
				case 1
				Response.write "1. Nyt�rsdag"
				end select
			case 2
				select case day(tjekdennedag)
				case 18
				Response.write "Fastelavn"
				end select
			case 4
				select case day(tjekdennedag)
				case 1
				Response.write "Palmes�ndag"
				case 5
				Response.write "Sk�rtorsdag"
				case 6
				Response.write "Langfredag"
				case 8 
				Response.write "P�skedag"
				case 9
				Response.write "2.P�skedag"
				end select
			case 5
				select case day(tjekdennedag)
				case 4
				Response.write "Bededag"
				case 17
				Response.write "Kristi Himmelfart"
				case 27
				Response.write "Pinsedag"
				case 28
				Response.write "2.Pinsedag"
				end select
			case 6
				select case day(tjekdennedag)
				case 5
				Response.write "Gr.lovsdag"
				end select
			case 12
				select case day(tjekdennedag)
				case 24
				Response.write "Juleaften"
				case 25
				Response.write "Juledag"
				case 26
				Response.write "2.Juledag"
				case 31
				Response.write "Nyt�rsaften"
				end select
			end select
		end select
end function
%>
