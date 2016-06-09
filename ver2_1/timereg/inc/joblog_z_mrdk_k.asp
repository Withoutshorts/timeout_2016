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
	<div id="mrdknap_<%=mrdnavn%>" style="position:absolute; top:130; left:<%=leftpix%>; background-color:<%=strBgCol%>; z-index:200; visibility:visible;">
	<a href="<%=thisfile%>.asp?jid=<%=jid%>&jnavn=<%=jnavn%>&mrd=<%=mrdnum%>&year=<%=strReqAar%>&slutd=<%=slutd%>&btimer=<%=btimer%>&uselogin=<%=uselogin%>&usekid=<%=kundeid%>"><%=strBG_img%></a>
	</div>
	<%
	
	leftpix = leftpix + pixadd
	next
	
	minus1 = right((year(now)-1), 2)
	iaar = right(year(now), 2)
	plus1 = right((year(now)+1), 2)
	
	
	if weekselector <> "j" then
		Select case strReqAar
		case minus1
		varBgColy1 = "<img src='../ill/y_"&year(now)-1&"_on.gif' border='0'>"
		varBgColy2 = "<img src='../ill/y_"&year(now)&".gif' border='0'>"
		varBgColy3 = "<img src='../ill/y_"&year(now)+1&".gif' border='0'>"
		varBgColy4 = "<img src='../ill/y_tot.gif' border='0'>"
		case iaar
		varBgColy1 = "<img src='../ill/y_"&year(now)-1&".gif' border='0'>"
		varBgColy2 = "<img src='../ill/y_"&year(now)&"_on.gif' border='0'>"
		varBgColy3 = "<img src='../ill/y_"&year(now)+1&".gif' border='0'>"
		varBgColy4 = "<img src='../ill/y_tot.gif' border='0'>"
		case plus1
		varBgColy1 = "<img src='../ill/y_"&year(now)-1&".gif' border='0'>"
		varBgColy2 = "<img src='../ill/y_"&year(now)&".gif' border='0'>"
		varBgColy3 = "<img src='../ill/y_"&year(now)+1&"_on.gif' border='0'>"
		varBgColy4 = "<img src='../ill/y_tot.gif' border='0'>"
		case "0"
		varBgColy1 = "<img src='../ill/y_"&year(now)-1&".gif' border='0'>"
		varBgColy2 = "<img src='../ill/y_"&year(now)&".gif' border='0'>"
		varBgColy3 = "<img src='../ill/y_"&year(now)+1&".gif' border='0'>"
		varBgColy4 = "<img src='../ill/y_tot_on.gif' border='0'>"
		end select
	
	else
	
		varBgColy1 = "<img src='../ill/y_"&year(now)-1&".gif' border='0'>"
		varBgColy2 = "<img src='../ill/y_"&year(now)&".gif' border='0'>"
		varBgColy3 = "<img src='../ill/y_"&year(now)+1&".gif' border='0'>"
		varBgColy4 = "<img src='../ill/y_tot.gif' border='0'>"
	
	end if
	%>
	<div id="aarknap1" style="position:absolute; top:130; left:670; z-index:199; visibility:visible;"><a href="<%=thisfile%>.asp?jid=<%=jid%>&jnavn=<%=jnavn%>&mrd=0&year=<%=right((year(now)-1), 2)%>&yearsel=j&slutd=<%=slutd%>&btimer=<%=btimer%>&uselogin=<%=uselogin%>&usekid=<%=kundeid%>"><%=varBgColy1%></a></div>
	<div id="aarknap2" style="position:absolute; top:130; left:715; z-index:199; visibility:visible;"><a href="<%=thisfile%>.asp?jid=<%=jid%>&jnavn=<%=jnavn%>&mrd=0&year=<%=right(year(now), 2)%>&yearsel=j&slutd=<%=slutd%>&btimer=<%=btimer%>&uselogin=<%=uselogin%>&usekid=<%=kundeid%>"><%=varBgColy2%></a></div>
	<div id="aarknap3" style="position:absolute; top:130; left:761; z-index:199; visibility:visible;"><a href="<%=thisfile%>.asp?jid=<%=jid%>&jnavn=<%=jnavn%>&mrd=0&year=<%=right((year(now)+1), 2)%>&yearsel=j&slutd=<%=slutd%>&btimer=<%=btimer%>&uselogin=<%=uselogin%>&usekid=<%=kundeid%>"><%=varBgColy3%></a></div>
	
	<!-- slut knapper -->
