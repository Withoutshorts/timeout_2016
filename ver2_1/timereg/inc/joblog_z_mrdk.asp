<!-- mrd knapper -->
	<%
	leftpix = 200
	mrdnum = 1
	
	for mrdnum = 1 to 12
	
	Select case mrdnum 
	case 1
	mrdnavn = "Jan"
	pixadd = 39
	case 2
	mrdnavn = "Feb"
	pixadd = 40
	case 3
	mrdnavn = "Mar"
	pixadd = 38
	case 4
	mrdnavn = "Apr"
	pixadd = 39
	case 5
	mrdnavn = "Maj"
	pixadd = 39
	case 6
	mrdnavn = "Jun"
	pixadd = 39
	case 7
	mrdnavn = "Jul"
	pixadd = 39
	case 8
	mrdnavn = "Aug"
	pixadd = 39
	case 9
	mrdnavn = "Sep"
	pixadd = 39
	case 10
	mrdnavn = "Okt"
	pixadd = 39
	case 11
	mrdnavn = "Nov"
	pixadd = 39
	case 12
	mrdnavn = "Dec"
	pixadd = 39
	end select
	
		if weekselector = "j" then
		strBgCol = "#D6DFF5"
		strBG_img = "<img src='../ill/m_"&mrdnavn&"_off.gif' border='0'>"
		else
			if cint(strReqMrd) = cint(mrdnum) then
			strBgCol = "#D6DFF5"
			strBG_img = "<img src='../ill/m_"&mrdnavn&"_on.gif' border='0'>"
			else
			strBgCol = "#D6DFF5"
			strBG_img = "<img src='../ill/m_"&mrdnavn&".gif' border='0'>"
			end if
		end if
	
	
	%>
	<div id="mrdknap_<%=mrdnavn%>" style="position:absolute; top:120; left:<%=leftpix%>; background-color:<%=strBgCol%>; z-index:200; visibility:visible;">
	<a href="<%=thisfile%>.asp?menu=<%=menu%>&mrd=<%=mrdnum%>&jobnr=<%=intJobnr%>&eks=<%=eks%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=FM_job%>&FM_medarb=<%=FM_medarb%>&FM_aar=<%=FM_Aar%>"><%=strBG_img%></a></div>
	<%
	
	leftpix = leftpix + pixadd
	next
	
	minus2 = right((year(now)-2), 2)
	minus1 = right((year(now)-1), 2)
	iaar = right(year(now), 2)
	plus1 = right((year(now)+1), 2)
	
	
	if weekselector <> "j" then
		Select case strReqAar
		case minus2
		varBgColy0 = "<img src='../ill/y_"&year(now)-2&"_on.gif' border='0'>"
		varBgColy1 = "<img src='../ill/y_"&year(now)-1&".gif' border='0'>"
		varBgColy2 = "<img src='../ill/y_"&year(now)&".gif' border='0'>"
		varBgColy3 = "<img src='../ill/y_"&year(now)+1&".gif' border='0'>"
		'varBgColy4 = "<img src='../ill/y_tot.gif' border='0'>"
		case minus1
		varBgColy0 = "<img src='../ill/y_"&year(now)-2&".gif' border='0'>"
		varBgColy1 = "<img src='../ill/y_"&year(now)-1&"_on.gif' border='0'>"
		varBgColy2 = "<img src='../ill/y_"&year(now)&".gif' border='0'>"
		varBgColy3 = "<img src='../ill/y_"&year(now)+1&".gif' border='0'>"
		'varBgColy4 = "<img src='../ill/y_tot.gif' border='0'>"
		case iaar
		varBgColy0 = "<img src='../ill/y_"&year(now)-2&".gif' border='0'>"
		varBgColy1 = "<img src='../ill/y_"&year(now)-1&".gif' border='0'>"
		varBgColy2 = "<img src='../ill/y_"&year(now)&"_on.gif' border='0'>"
		varBgColy3 = "<img src='../ill/y_"&year(now)+1&".gif' border='0'>"
		'varBgColy4 = "<img src='../ill/y_tot.gif' border='0'>"
		case plus1
		varBgColy0 = "<img src='../ill/y_"&year(now)-2&".gif' border='0'>"
		varBgColy1 = "<img src='../ill/y_"&year(now)-1&".gif' border='0'>"
		varBgColy2 = "<img src='../ill/y_"&year(now)&".gif' border='0'>"
		varBgColy3 = "<img src='../ill/y_"&year(now)+1&"_on.gif' border='0'>"
		'varBgColy4 = "<img src='../ill/y_tot.gif' border='0'>"
		case "0"
		varBgColy0 = "<img src='../ill/y_"&year(now)-2&".gif' border='0'>"
		varBgColy1 = "<img src='../ill/y_"&year(now)-1&".gif' border='0'>"
		varBgColy2 = "<img src='../ill/y_"&year(now)&".gif' border='0'>"
		varBgColy3 = "<img src='../ill/y_"&year(now)+1&".gif' border='0'>"
		'varBgColy4 = "<img src='../ill/y_tot_on.gif' border='0'>"
		end select
	
	else
		
		varBgColy0 = "<img src='../ill/y_"&year(now)-2&".gif' border='0'>"
		varBgColy1 = "<img src='../ill/y_"&year(now)-1&".gif' border='0'>"
		varBgColy2 = "<img src='../ill/y_"&year(now)&".gif' border='0'>"
		varBgColy3 = "<img src='../ill/y_"&year(now)+1&".gif' border='0'>"
		'varBgColy4 = "<img src='../ill/y_tot.gif' border='0'>"
	
	end if
	%>
	<div id="aarknap0" style="position:absolute; top:120; left:670; z-index:199; visibility:visible;"><a href="<%=thisfile%>.asp?menu=<%=menu%>&mrd=0&jobnr=<%=intJobnr%>&eks=<%=eks%>&year=<%=right((year(now)-2), 2)%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=FM_job%>&FM_medarb=<%=FM_medarb%>&FM_aar=<%=FM_Aar%>&yearsel=j"><%=varBgColy0%></a></div>
	<div id="aarknap1" style="position:absolute; top:120; left:715; z-index:199; visibility:visible;"><a href="<%=thisfile%>.asp?menu=<%=menu%>&mrd=0&jobnr=<%=intJobnr%>&eks=<%=eks%>&year=<%=right((year(now)-1), 2)%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=FM_job%>&FM_medarb=<%=FM_medarb%>&FM_aar=<%=FM_Aar%>&yearsel=j"><%=varBgColy1%></a></div>
	<div id="aarknap2" style="position:absolute; top:120; left:761; z-index:199; visibility:visible;"><a href="<%=thisfile%>.asp?menu=<%=menu%>&mrd=0&jobnr=<%=intJobnr%>&eks=<%=eks%>&year=<%=right(year(now), 2)%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=FM_job%>&FM_medarb=<%=FM_medarb%>&FM_aar=<%=FM_Aar%>&yearsel=j"><%=varBgColy2%></a></div>
	<div id="aarknap3" style="position:absolute; top:120; left:807; z-index:199; visibility:visible;"><a href="<%=thisfile%>.asp?menu=<%=menu%>&mrd=0&jobnr=<%=intJobnr%>&eks=<%=eks%>&year=<%=right((year(now)+1), 2)%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=FM_job%>&FM_medarb=<%=FM_medarb%>&FM_aar=<%=FM_Aar%>&yearsel=j"><%=varBgColy3%></a></div>
	<!--<div id="aarknap4" style="position:absolute; top:140; left:804; z-index:198; visibility:visible;"><a href="joblog_z.asp?menu=stat&mrd=0&jobnr=<=intJobnr%>&eks=<=eks%>&year=-1&lastFakdag=<=lastFakdag%>&selmedarb=<=selmedarb%>&selaktid=<=selaktid%>&FM_job=<=FM_job%>&FM_medarb=<=FM_medarb%>&FM_aar=<=FM_Aar%>"><=varBgColy4%></a></div>-->
	
	<!-- slut knapper -->
