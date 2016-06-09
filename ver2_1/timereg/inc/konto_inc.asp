
<%
'public strMrd, strDag, strAar, strMrd_slut
'public strDag_slut, strAar_slut
'public strStartDato, strSlutDato

function finddatoer()
if len(request("FM_start_dag")) <> 0 then
	strMrd = request("FM_start_mrd")
	strDag = request("FM_start_dag")
	strAar = right(request("FM_start_aar"),2) 
	strDag_slut = request("FM_slut_dag")
	strMrd_slut = request("FM_slut_mrd")
	strAar_slut = right(request("FM_slut_aar"),2)
	else	
			if len(request.Cookies("erp")("strDag_kon")) <> 0 then
			
			strDag = request.Cookies("erp")("strDag_kon")  
			strMrd = request.Cookies("erp")("strMrd_kon")
			strAar = request.Cookies("erp")("strAar_kon") 
			
			strDag_slut = request.Cookies("erp")("strDag_kon_slut")
			strMrd_slut = request.Cookies("erp")("strMrd_kon_slut")
			strAar_slut = request.Cookies("erp")("strAar_kon_slut")
			
			else
				
				strMrd_temp = dateadd("m", -1, date)
				strMrd = month(strMrd_temp)
				strDag_temp = dateadd("d", -30, date)
				strDag = day(strDag_temp)
				if strMrd <> 12 then
				strAar= right(year(now), 2)
				else
				strAar_temp = dateadd("yyyy", -1, date)
				strAar = right(year(strAar_temp),2)
				end if
			
			
				strDag_slut = day(now)
				strMrd_slut = month(now)
				strAar_slut = right(year(now), 2)
			end if
	end if
	
	strStartDato = strAar&"/"&strMrd&"/"&strDag
	strSlutDato = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	
	response.Cookies("erp")("strDag_kon") = strDag
	response.Cookies("erp")("strMrd_kon") = strMrd
	response.Cookies("erp")("strAar_kon") = strAar
	
	
	response.Cookies("erp")("strDag_kon_slut") = strDag_slut
	response.Cookies("erp")("strMrd_kon_slut") = strMrd_slut
	response.Cookies("erp")("strAar_kon_slut") = strAar_slut
	response.Cookies("erp").Expires = date + 30
end function
	
 %>