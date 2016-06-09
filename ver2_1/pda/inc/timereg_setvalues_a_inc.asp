<%
'*** sql datoer ***
	varTjDatoUS_son = convertDateYMD(tjekdag(1))
	varTjDatoUS_man = convertDateYMD(tjekdag(2))
	varTjDatoUS_tir = convertDateYMD(tjekdag(3))
	varTjDatoUS_ons = convertDateYMD(tjekdag(4)) 
	varTjDatoUS_tor = convertDateYMD(tjekdag(5)) 
	varTjDatoUS_fre = convertDateYMD(tjekdag(6))
	varTjDatoUS_lor = convertDateYMD(tjekdag(7))
	
	
	'******************************************************************
	'Sætter farver på første row 
	fakbgcol_son = "#cd853f"
	fakbgcol_man = "#7F9DB9"	
	fakbgcol_tir = "#7F9DB9"	
	fakbgcol_ons = "#7F9DB9"	
	fakbgcol_tor = "#7F9DB9"	
	fakbgcol_fre = "#7F9DB9"	
	fakbgcol_lor = "#cd853f"
	
	maxl_son = 5
	maxl_man = 5
	maxl_tir = 5
	maxl_ons = 5
	maxl_tor = 5
	maxl_fre = 5
	maxl_lor = 5
	
	fmbg_son = "#FFDFDF" 
	fmbg_man = "#FFFFFF" 
	fmbg_tir = "#FFFFFF" 
	fmbg_ons = "#FFFFFF" 
	fmbg_tor = "#FFFFFF" 
	fmbg_fre = "#FFFFFF" 
	fmbg_lor = "#FFDFDF"  
	'******************************************************************
	
	Redim sonTimerVal(iRowLoop)
	Redim manTimerVal(iRowLoop)
	Redim tirTimerVal(iRowLoop)
	Redim onsTimerVal(iRowLoop)
	Redim torTimerVal(iRowLoop)
	Redim freTimerVal(iRowLoop)
	Redim lorTimerVal(iRowLoop)
	
	Redim dtimeTtidspkt_son(iRowLoop)
	Redim dtimeTtidspkt_man(iRowLoop)
	Redim dtimeTtidspkt_tir(iRowLoop)
	Redim dtimeTtidspkt_ons(iRowLoop)
	Redim dtimeTtidspkt_tor(iRowLoop)
	Redim dtimeTtidspkt_fre(iRowLoop)
	Redim dtimeTtidspkt_lor(iRowLoop)
	
	sonRLoop = ""
	manRLoop = ""
	tirRLoop = ""
	onsRLoop = ""
	torRLoop = ""
	freRLoop = ""
	lorRLoop = ""
	%>