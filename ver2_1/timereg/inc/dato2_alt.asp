<%


'Response.Write "<br><br><br><br><br><br>id" & id 
'Response.Flush
if id <> 0 then
    
    strMrd_alt = month(StrTdato_alt)
	strDag_alt = day(StrTdato_alt)
	strAar_alt = right(year(StrTdato_alt), 2)	
		
else
	
	strMrd_alt = month(now())
	strDag_alt = day(now())
	strAar_alt = year(now())
	
	
end if


	select case strMrd_alt
	case "1"
	strMrd_navn_alt = "jan" 
	case "2"
	strMrd_navn_alt = "feb"
	case "3"
	strMrd_navn_alt = "mar"
	case "4"
	strMrd_navn_alt = "apr"
	case "5"
	strMrd_navn_alt = "maj"
	case "6"
	strMrd_navn_alt = "jun"
	case "7"
	strMrd_navn_alt = "jul"
	case "8"
	strMrd_navn_alt = "aug"
	case "9"
	strMrd_navn_alt = "sep"
	case "10"
	strMrd_navn_alt = "okt"
	case "11"
	strMrd_navn_alt = "nov"
	case "12"
	strMrd_navn_alt = "dec"
	end select
	
	


%>



