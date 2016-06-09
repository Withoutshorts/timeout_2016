<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	func = request("func")
	if len(trim(request("y"))) <> 0 then
	ysel = request("y")
	else
	ysel = year(now)
	end if
	
	
	
	
	
	if request("print") = "j" then%>	
		<html>
		<head>
		<title>timeOut</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
		</head>
		<body topmargin="0" leftmargin="0" class="regular">
	<%end if
	
	if request("print") = "j" then
	leftPos = 10
	topPos = 60
	bgthis = "#5582d2"
	else
	leftPos = 20
	topPos = 132
	bgthis = "#5582d2"
	%>
	
	<script language="javascript">
	
	function showdatoinfo(id){
	//alert(id)
	document.getElementById("divid_info_"+id).style.visibility = "visible"
	document.getElementById("divid_info_"+id).style.display = ""
	}
	
	function hidedatoinfo(id){
	document.getElementById("divid_info_"+id).style.visibility = "hidden"
	document.getElementById("divid_info_"+id).style.display = "none"
	}
	
	</script>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%call tsamainmenu(2)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	call webbliktopmenu()
	%>
	</div>
	
	

		
	<%
	end if
	%>
	<!-------------------------------Sideindhold------------------------------------->
	
	
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	
	<% 
    
    
    
    oimg = "ikon_ferieogsygdom_48.png"
	oleft = 0
	otop = 0
	owdt = 500
	if level = 1 then
	oskrift = "Ferie, Afspad. & Sygdom's kalender"
	else
	oskrift = "Ferie & Afspad. kalender"
	end if
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	%>
	<br /><br />
	
	
	
	
	
	<h4>Valgt år: <%=ysel%></h4>
	
	
	
	<table cellspacing=0 cellpadding=0 border=0 width=1184>
	<tr>
	<td><b>Vælg år:</b></td>
    <td>
    	<%
	for y = -5 to 5
	yearshow = year(now) + (y)%>
	<a href="feriekalender.asp?y=<%=yearshow%>" class=vmenu><%=yearshow%></a> | 
	<%
	next
	%>
	
	</td>
	<td align=right>
	<a href="bal_real_norm_2007.asp?menu=stat" class=vmenu>Medarbejder afstemning..</a> &nbsp;&nbsp;|&nbsp;&nbsp;<a href="timereg_2006_fs.asp" class=vmenu>Timeregistrering ..</a>
	</td>
	
	</tr>
	</table>
	
	
	<%  
	
	tTop = 0
	tLeft = 0
	tWdth = 1184
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>
	
	
	
	<table cellspacing=0 cellpadding=0 border=0 bgcolor=#8caae6>
	<tr bgcolor="#3B5998">
	    <td style="height:15px; width:105px; padding-left:3px; border-right:1px #8caae6 solid;" class=alt>
            Uge<br />
            <img src="ill/blank.gif" width=105 height=1 border=0 /><br /></td>
	
	<%
	    
	    lastweek = 0
	    firstjan = "1/1/"&ysel
	    
	    stdatoThis = ysel&"/10/2"
	    stdato = day(stdatoThis)&"/"&month(stdatoThis)&"/"&year(stdatoThis)
	    
	    for d = 0 to 365
	    
	    if d = 0 then
	    newDate = firstjan
	    weekd = weekday(firstjan, 2)
	    wwdt = (7 - weekd) * 3
	    else
	    newDate = dateadd("d", 1, newDate)
	    wwdt = 20
	    end if
	    
	    weeknum = datepart("ww", newDate,2,2)
	    
	    if weeknum <> lastweek then
	    %>
	    <td class=alt_lille align=center style="width:<%=wwdt %>; border-right:1px #8caae6 solid;"><%=weeknum%>
	    <img src="ill/blank.gif" width=<%=wwdt %> height=1 border=0 /><br /></td>
	    <%
	    lastweek = weeknum
	    end if
	    
	    next
	
	
	%>
	</tr>
	
	</table>
	
	<table cellspacing=0 cellpadding=0 border=0 bgcolor=#ffffff>
	<tr bgcolor="#8caae6">
	
	<td style="width:105px; height:25px; padding-left:3px; border-right:1px #3B5998 solid;">Medarb./ Måned
	<img src='../ill/blank.gif' width='105' height='1' alt='' border='0'><br></td>
	
	<%
	    
	    lastmth = 0
	    firstjan = "1/1/"&ysel
	    for d = 0 to 340 'midt december
	    
	    if d = 0 then
	    newDate = firstjan
	    else
	    newDate = dateadd("d", 1, newDate)
	    end if
	    
	    monthnum = datepart("m", newDate,2,2)
	    
	    if monthnum <> lastmonth then
	        
	        select case monthnum
	        case 1,3,5,7,8,10,12
	        mwdt = 93
	        case 2
	            if ysel = 2000 OR ysel = 2004 OR ysel = 2008 OR ysel = 2012 OR ysel = 2016 OR ysel = 2020 OR ysel = 2024 then
	            mwdt = 83
	            else 
	            mwdt = 80
	            end if
	        case else
	        mwdt = 90
	        end select
	        
	        %>
	    <td align=center style="border-right:1px #3B5998 solid;"><%=monthname(monthnum)%>
	    <img src="ill/blank.gif" width="<%=mwdt %>" height=1 /><br /></td>
	    <%
	    lastmonth = monthnum
	    end if
	    
	    next
	
	
	%>
	
	
	</tr>
	
	
	<%
	
	dim medarbid, normTimerUge
	redim medarbid(100), normTimerUge(100)
	
	m = 1
	
	strSQLmed = "SELECT mid, mnavn, mnr, init FROM medarbejdere WHERE mansat = '1' ORDER BY mnavn"
	
	oRec.open strSQLmed, oConn, 3
	while not oRec.EOF 
	
	medarbid(m) = oRec("mid")
	
	    
	%>
	<tr>
	    <td bgcolor="#ffffe1" class=lille style="width:105px; height:40px; border-bottom:1px #8caae6 solid; border-right:1px #3B5998 solid;">
	     <%=ntimPer %>  <%=left(oRec("mnavn"), 12) %> (<%=oRec("mnr") %>) - <%=oRec("init") %>
	    <img src="ill/blank.gif" width=105 height=1 /><br /></td>
	    
	    <%
	    for x = 1 to 12
	        
	        if lasttdbg = "#ffffff" then
	        bgtd = "#d6dff5"
	        else
	        bgtd = "#ffffff"
	        end if
	        
	        lasttdbg = bgtd
	    %>
	    <td bgcolor="<%=bgtd %>" style="border-bottom:1px #8caae6 solid; border-right:1px #3B5998 solid;"><img src='../ill/blank.gif' width='1' height='1' alt='' border='0'>
            </td>
	    <%
	    next
	    %>
	
	</tr>
	
	
	
	<%
	m = m + 1
	
	oRec.movenext
	wend
	oRec.close
	 %>
	</table>
	
	<%
	
	
	if level = 1 then
	sqlTypKri = "tfaktim = 11 OR tfaktim = 14 OR tfaktim = 20 OR tfaktim = 21 OR tfaktim = 31 OR tfaktim = 13 OR tfaktim = 18 OR tfaktim = 19"
	else
	sqlTypKri = "tfaktim = 11 OR tfaktim = 14 OR tfaktim = 31 OR tfaktim = 13 OR tfaktim = 18 OR tfaktim = 19"
	end if
	
	
	'** Henter ferie og sygdom **' 
	i = 0
	for t = 1 to (m - 1)
	
	    strSQLt = "SELECT sum(timer) AS timer, tdato, tfaktim FROM timer WHERE tmnr = " & medarbid(t) & ""_
	    &" AND ("& sqlTypKri &") AND tdato BETWEEN '"&ysel&"-01-01' AND '"&ysel&"-12-31' GROUP BY tdato, tfaktim, tmnr ORDER BY tdato"
	    
	    '
	    'Response.Write strSQLt & "<br>"
	    'Response.flush
	    
	    oRec.open strSQLt, oConn, 3
	    while not oRec.EOF 
	        
	        
	     
	     select case oRec("tfaktim")
	     
	     case 11,18 'Ferie planlagt, Feriefridag pl
	     bgcol = "silver"
	     tpekstra = -20
	     zinx = 300
	     tnavn = "Ferie/Feriefrid. planlagt"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer " ' ~ "& formatnumber(oRec("timer")/(normTimerUge(m)), 2) & " dag."
	     
	     case 14,19 'Ferie afholdt, Ferie u løn
	     bgcol = "green"
	     tpekstra = -20
	     zinx = 400
	     tnavn = "Ferie afholdt"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer " ' ~ "& formatnumber(oRec("timer")/(normTimerUge(m)), 2) & " dag."
	     
	    
	     
	     case 20 'Syg
	     bgcol = "red"
	     tpekstra = 0
	     zinx = 200
	     tnavn = "Syg"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer."
	     
	     case 21 'Barn syg
	     bgcol = "orange"
	     tpekstra = 10
	     zinx = 100
	     tnavn = "Barn syg"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer."
	     
	     
	     
	     case 31 'Afsp. 
	     bgcol = "#8cAAe6"
	     tpekstra = -10
	     zinx = 100
	     tnavn = "Afspadsering"
	     hoursval = formatnumber(oRec("timer"), 2) & " timer."
	     
	     case 13 'Ferie fri brugt
	     bgcol = "yellow"
	     tpekstra = -10
	     zinx = 200
	     tnavn = "Feriefridage brugt."
	     hoursval = formatnumber(oRec("timer"), 2) & " timer."
	     end select
	     
	     
	   
	     select case datepart("m", oRec("tdato"), 2,2)
	     case 1
	     bregnM = 0
	     mthval = 0
	     case else 
	     bregnM = dateadd("m", -1, oRec("tdato"))
	     mthval = (cint(datepart("m", bregnM, 2,2)) * 88) + cint(datepart("m", bregnM, 2,2)) * 4
	     end select
	    
	     stval = 109
	     dayval = cint((datepart("d", oRec("tdato"), 2,2))) * 3
	     
	    
	     leftval = stval + mthval + dayval
	    
	   
	     
	     %>
	   
	    <div class=hand id="divid_<%=i%>" style="position:absolute; left:<%=leftval %>px; top:<%=23 + (t*40) + tpekstra%>px; background-color:<%= bgcol %>; height:10px; width:2px;" onclick="showdatoinfo('<%=i%>')"><img src='../ill/blank.gif' width='1' height='1' alt='' border='0'></div>
	    
	    <div class=hand id="divid_info_<%=i%>" style="position:absolute; visibility:hidden; display:none; left:<%=leftval-50%>px; top:<%=25 + (t*41) + tpekstra - 20%>px; background-color:#ffffe1; height:50px; width:150px; border:1px #cccccc solid; padding:10px; z-index:1000;" onclick="hidedatoinfo('<%=i%>')">
	    <b><%=ucase(left(weekdayname(datepart("w", oRec("tdato"))), 1)) &""& mid(weekdayname(datepart("w", oRec("tdato"))), 2,2) %>. d. <%=oRec("tdato") %></b><br />
	    Uge <%=datepart("ww", oRec("tdato"), 2,2) %> 
	    <br /><%=tnavn %><br />
	    <%=hoursval%> 
	   </div>
	    
	    <%
	    i = i + 1
	    oRec.movenext
	    wend
	    oRec.close
	    
	
	next
	
	
	%>
	
	
	
	
	
	</div><!-- table div-->
	
	
	
	
       <% 
       itop = 40
       ileft = 0 
       iwdt = 400
       call sideinfo(itop,ileft,iwdt)%>
	
        Ferie plantlagt og ferie afholdt indtastes via timeregistrerings siden. Det kræver at der er oprettet en
        aktivitet af typerne "Ferie plantlagt" og "Ferie afholdt".
        
        <br /><br />
          Fra den 1. januar 2001 optjenes 2,08 dages ferie for hver måneds ansættelse i optjeningsåret, som er lig kalenderåret. 
        Dette gælder også medarbejdere på del-tid.<br />
        2,08 * 12 måneder = 25 dage ell. 5 arbejdsuger.<br /><br />
        Ferie timer angives som timer og omregnes til dage udfra gns. normeret timer pr. uge 
        (se medarbejdertyper ell. medarb. afstem.). Så hvis man er ansat 37 timer, 
        angiver man altså 7,4 timer for en hel dags ferie, og 37 timer for en uge. 
        <br /><br />
        <b>Signatur forklaring:</b><br />
        
         <div style="position:relative; top:5px; padding:3px; left:5px; border:1px #000000 solid; background-color:silver; width:200px;">Ferie planlagt (Norm. t./d. = 1 dag)</div>
	     <div style="position:relative; top:15px; padding:3px; left:5px; border:1px #000000 solid; background-color:green; width:200px;">Ferie afholdt (Norm. t./d. = 1 dag)</div>
	     <%if level = 1 then %>
	     <div style="position:relative; top:25px; padding:3px; left:5px; border:1px #000000 solid; background-color:red; width:200px;">Syg</div>
	     <div style="position:relative; top:35px; padding:3px; left:5px; border:1px #000000 solid; background-color:orange; width:200px;">Barn syg</div>
	      <%end if %>
	      <div style="position:relative; top:45px; padding:3px; left:5px; border:1px #000000 solid; background-color:#8caae6; width:200px;">Afspadsering</div>
	     <div style="position:relative; top:55px; padding:3px; left:5px; border:1px #000000 solid; background-color:yellow; width:200px;">Feriefridage brugt.</div>
	     
	     

	     
	    <br /><br />
	    <br /><br />
            &nbsp;
        
        <!-- side info slut -->
        </td></tr></table>
			</div>
			
			
	<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
            &nbsp;
		
	</div>	<!-- side div-->
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
