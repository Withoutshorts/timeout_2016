<%

	strMrdNr = month(now())
	strDagSort = day(now())
	strYearNow = year(now())
	
	select case strMrdNr
	case "1"
	strMrdNavn = "jan" 
	case "2"
	strMrdNavn = "feb"
	case "3"
	strMrdNavn = "mar"
	case "4"
	strMrdNavn = "apr"
	case "5"
	strMrdNavn = "maj"
	case "6"
	strMrdNavn = "jun"
	case "7"
	strMrdNavn = "jul"
	case "8"
	strMrdNavn = "aug"
	case "9"
	strMrdNavn = "sep"
	case "10"
	strMrdNavn = "okt"
	case "11"
	strMrdNavn = "nov"
	case "12"
	strMrdNavn = "dec"
	
	end select


%>