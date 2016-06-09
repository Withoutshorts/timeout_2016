	
	<%
	'strDato = request.form("mrd")&"/"&request.form("dag")&"/"&request.form("aar")
 	'Datoen konfimeres eller der findes en ny
	if session("strDato") = "" then
	session("strDato") = request.form("mrd")&"/"&request.form("dag")&"/"&request.form("aar")
	'strDato = request.form("mrd")&"/"&request.form("dag")&"/"&request.form("aar")
	session("strDag") = request.form("dag")
	session("strMrd") = request.form("mrd")
	session("straar") = request.form("aar")
	else
	 
	session("strDato") = session("strDato")
	end if
	%>
	
