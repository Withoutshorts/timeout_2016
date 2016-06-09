<%
if id <> 0 then
	
	' Hvis datoformat er Long Date e.lign.
	justerLangDatoT = instr(StrTdato, "200")
	if justerLangDatoT <> 0 then
	StrTdato_end = right(StrTdato, 2)
	StrTdato_start = left(StrTdato, (justerLangDatoT - 1)) 
	StrTdato = StrTdato_start & StrTdato_end
	
	justerLangDatoU = instr(StrUdato, "200")
	StrUdato_end = right(StrUdato, 2)
	StrUdato_start = left(StrUdato, (justerLangDatoU - 1)) 
	StrUdato = StrUdato_start & StrUdato_end
	end if
	
	
		if len(StrTdato) > 7 then
			strMrd = left(StrTdato, 2)
			strDag  = mid(StrTdato,4,2)
			strAar = Right(StrTdato,2)
		else
			if len(StrTdato) = 6 then
			strMrd = left(StrTdato, 1)
			strDag  = mid(StrTdato,3,1)
			strAar = Right(StrTdato,2)
			else
				varStregMrd = instr(StrTdato, "/")
				if varStregMrd < 3 then 
				strMrd = left(StrTdato, 1)
				strDag  = mid(StrTdato,3,2)
				strAar = Right(StrTdato,2)
				else
				strMrd = left(StrTdato, 2)
				strDag  = mid(StrTdato,4,1)
				strAar = Right(StrTdato,2)
				end if
			end if
		end if
		
		if len(StrUdato) > 7 then
			strMrd_slut = left(strUdato, 2)
			strDag_slut  = mid(strUdato,4,2)
			strAar_slut = Right(strUdato,2)
		else
			if len(strUdato) = 6 then
			strMrd_slut = left(strUdato, 1)
			strDag_slut  = mid(strUdato,3,1)
			strAar_slut = Right(strUdato,2)
			else
				varStregMrd = instr(strUdato, "/")
				if varStregMrd < 3 then 
				strMrd_slut = left(StrUdato, 1)
				strDag_slut  = mid(strUdato,3,2)
				strAar_slut = Right(strUdato,2)
				else
				strMrd_slut = left(strUdato, 2)
				strDag_slut  = mid(strUdato,4,1)
				strAar_slut = Right(strUdato,2)
				end if
			end if
		end if
else
	
	strMrd = month(now())
	strDag = day(now())
	strAar = year(now())
	
	if strMrd = 12 then
	strMrd_slut = "1"
	else
	strMrd_slut = (month(now()) + 1)
	end if
	
	if strMrd_slut = 1 then
	strAar_slut = (year(now()) + 1)
	else
	strAar_slut = year(now()) 
	end if
	
	if strDag > "28" then
	strDag_slut = "1"
	strMrd_slut = strMrd_slut + 1
	else
	strDag_slut = day(now())
	end if
	
end if


	select case strMrd
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
	
	
	select case strMrd_slut
	case "1"
	strMrdNavn_slut = "jan" 
	case "2"
	strMrdNavn_slut = "feb"
	case "3"
	strMrdNavn_slut = "mar"
	case "4"
	strMrdNavn_slut = "apr"
	case "5"
	strMrdNavn_slut = "maj"
	case "6"
	strMrdNavn_slut = "jun"
	case "7"
	strMrdNavn_slut = "jul"
	case "8"
	strMrdNavn_slut = "aug"
	case "9"
	strMrdNavn_slut = "sep"
	case "10"
	strMrdNavn_slut = "okt"
	case "11"
	strMrdNavn_slut = "nov"
	case "12"
	strMrdNavn_slut = "dec"
	end select

%>



