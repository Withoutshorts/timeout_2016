<%



'Response.Write "<br><br><br><br><br><br>id" & id & "StrUdato:" & StrUdato
'Response.Flush
if id <> 0 then
    
    strMrd = month(StrTdato)
	strDag = day(StrTdato)
	strAar = right(year(StrTdato), 2)	
	
	strMrd_slut = month(strUdato)
	strDag_slut  = day(strUdato)
	strAar_slut =  right(year(StrUdato), 2)	

		
else
	
	strMrd = month(now())
	strDag = day(now())
	strAar = year(now())
	
    'strDag_slut = day(now())

	if strMrd = 12 then
	strMrd_slut = 1
	else
	strMrd_slut_temp = dateadd("m", 1, strDag&"/"&strMrd&"/"&strAar)
	strMrd_slut = month(strMrd_slut_temp)
	end if

    if strMrd_slut = 1 then
	strAar_slut_temp = dateadd("yyyy", 1, strDag&"/"&strMrd&"/"&strAar)
	strAar_slut = year(strAar_slut_temp)
	else
	strAar_slut = year(now()) 
	end if
	
    ''if strDag > 28 then
    '* SKAL IKKE VÆLGE NY dato på TIL FAKTURERING, da den altid skal stå til DD som faktura dato
    if strDag > 28 AND thisfile <> "erp_tilfakturering.asp" then
	strDag_slut = 1
    strMrd_slut_temp = dateadd("m", 2,  strDag&"/"&strMrd&"/"&strAar)
	strMrd_slut = month(strMrd_slut_temp)
	strAar_slut = year(strMrd_slut_temp)
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



