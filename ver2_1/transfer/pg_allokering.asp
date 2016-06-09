<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	if len(request("year")) <> 0 then
	useyear = request("year")
	else
	useyear = year(now)
	end if
	
	if request("order") = "j" then
	orderkri = " jobnavn "
	else
	orderkri = " jobstartdato, jobslutdato "
	end if
	
	if request("print") <> "j" then%>	
		<!--#include file="../inc/regular/header_inc.asp"-->
	<%else%>
		<html>
		<head>
		<title>timeOut 2.1</title>
		<LINK rel="stylesheet" type="text/css" href="../inc/style/timeout_style_print_fak.css">
		</head>
		<body topmargin="0" leftmargin="0" class="regular">
	<%end if
	
	if request("print") = "j" then
	leftPos = 10
	topPos = 60
	bgthis = "#FFFFFF"
	else
	leftPos = 190
	topPos = 80
	bgthis = "#5582d2"
	%>
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!--#include file="../inc/regular/rmenu.asp"-->
	<%
	end if
	%>
	<!-------------------------------Sideindhold------------------------------------->
	<%if request("print") = "j" then%>	
	<table cellspacing="0" cellpadding="0" border="0" width="880">
		<tr>
			<td bgcolor="#003399" width="650"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
			<td bgcolor="#FFFFFF" align=right><a href="javascript:window.close()"><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"><a href="javascript:window.print()"><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a><img src="../ill/blank.gif" width="30" height="1" alt="" border="0"></td>
		</tr>
		</table>
	<%end if%>
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	<%if request("print") <> "j" then%>	
	<table cellspacing="0" cellpadding="0" border="0" width="600">
	<tr>
    	<td valign="top"><img src="../ill/header_jobplanner.gif" alt="" border="0" width="758" height="45"></TD>
	</TR>
	</TABLE>
	<%end if%>
	<table border=0 cellpadding=0 cellspacing=0 width="600">
	<tr>
	<td valign="top"><img src="../ill/blank.gif" width="5" height="53" alt="" border="0"></td>
	<td><b>Jobplanner, alle job.</b><br>
	Oversigt over alle job med startdato og slutdato der passer det valgte år.
	<br>Jobplanner kan sorteres efter startdato eller jobnavn.</td>
	</tr>
	<tr>
	<td colspan="2"><br>Åben:&nbsp;<img src="../ill/jobpl1.gif" width="6" height="14" alt="" border="0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Lukket:&nbsp;<img src="../ill/jobpl0.gif" width="6" height="14" alt="" border="0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	Passivt:&nbsp;<img src="../ill/jobpl2.gif" width="6" height="14" alt="" border="0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	
	<%if request("print") <> "j" then%>
	&nbsp;&nbsp;<b>Vælg år:</b>&nbsp;<a href="pg_allokering.asp?menu=job&year=2002">2002</a>&nbsp;&nbsp;<a href="pg_allokering.asp?menu=job&year=2003">2003</a>&nbsp;&nbsp;<a href="pg_allokering.asp?menu=job&year=2004">2004</a>&nbsp;&nbsp;<a href="pg_allokering.asp?menu=job&year=2005">2005</a>&nbsp;&nbsp;<a href="pg_allokering.asp?menu=job&year=2006">2006</a>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:NewWin('pg_allokering.asp?menu=job&print=j&year=<%=useyear%>&usejob=<%=usejob%>')" target="_self" class='rmenu'>&nbsp;Printer venlig version&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	<br>
	<%end if%>
	</td>
	</tr>
	</table>
	<%if request("print") <> "j" then%>
	<div id="opretny" style="position:absolute; left:600; top:50; visibility:visible;">
	<a href='jobs.asp?menu=job&func=opret&id=0&int=1&rdir=pg' class='vmenu'><img src="../ill/oretnytjobpil.gif" width="143" height="28" alt="" border="0"></a>
	</div>
	<%end if%>
	
	<a href="jbpla_fs.asp" class=vmenu>Zoom + kalender</a> <br>
	<a href="#" class=vmenu>Ressource belægning</a><br>
	
	<b>Oversigt er for år:</b> <font class=roed><%=useyear%></font>
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#5582d2">
	<tr bgcolor="<%=bgthis%>">
	<td width="150" class=alt>&nbsp;<a href="pg_allokering.asp?menu=job&order=j" class=alt>Jobnavn</a>&nbsp;&nbsp;</td>
	<td width=31 class=alt>&nbsp;<b>Jan</b></td>
	<td width=29 class=alt>&nbsp;<b>Feb</b></td>
	<td width=31 class=alt >&nbsp;<b>Mar</b></td>
	<td width=30 class=alt >&nbsp;<b>Apr</b></td>
	<td width=31 class=alt >&nbsp;<b>Maj</b></td>
	<td width=30 class=alt >&nbsp;<b>Jun</b></td>
	<td width=31 class=alt >&nbsp;<b>Jul</b></td>
	<td width=31 class=alt >&nbsp;<b>Aug</b></td>
	<td width=30 class=alt >&nbsp;<b>Sep</b></td>
	<td width=31 class=alt >&nbsp;<b>Okt</b></td>
	<td width=30 class=alt >&nbsp;<b>Nov</b></td>
	<td width=31 class=alt >&nbsp;<b>Dec</b></td>
	<td width=222 colspan="2" align="right" style="padding-right:5;">&nbsp;&nbsp;&nbsp;<a href="pg_allokering.asp?menu=job&order=d" class=alt>Job start og slut dato.</a></td>
	</tr>
	</table>
	
	<table cellspacing="1" cellpadding="0" border="0">
	
	<%
	useLowerdatoKri = useyear&"/1/1"
	useUpperdatoKri = useyear&"/12/31"
	
	strSQL = "SELECT job.id AS jobid, jobnavn, jobstartdato, jobslutdato, jobstatus FROM job WHERE fakturerbart = 1 AND (jobstartdato <= '"& useUpperdatoKri &"' AND jobslutdato >= '"& useLowerdatoKri &"') ORDER BY " & orderkri
	oRec.open strSQL, oConn, 3
	while not oRec.EOF
	
	select case oRec("jobstatus")
	case 0
	useColorgif = "<img src='../ill/jobpl0.gif' width='6' height='14' alt='' border='0'>"
	case 1
	useColorgif = "<img src='../ill/jobpl1.gif' width='6' height='14' alt='' border='0'>"
	case 2
	useColorgif = "<img src='../ill/jobpl2.gif' width='6' height='14' alt='' border='0'>"
	end select
	
		%>
		<tr>
			<td class="blaa" valign=top bgcolor="#fffff1" height="10" width="150">&nbsp;&nbsp;
			<%if request("print") = "j" then%>
			<%=left(oRec("jobnavn"), 16)%></td>
			<%else%>
			<a href="jobs.asp?menu=job&func=red&int=1&id=<%=oRec("jobid")%>&rdir=pg" class="vmenu"><%=left(oRec("jobnavn"), 16)%></a>
			&nbsp;<a href="javascript:NewWin_large('jobplanner.asp?menu=job&id=<%=oRec("jobid")%>')"><img src="../ill/popupcal_small.gif" width="11" height="10" alt="Rediger jobperiode og aktiviteter" border="0"></a></td>
			<%end if%>
		
		
		
			<%'*******************************************************************
			'** startår = valgtår
			'if cint(datepart("yyyy",oRec("jobstartdato"))) = cint(useyear) then
			'projektlaengde = Datediff("d", oRec("jobstartdato"), oRec("jobslutdato"))
			strWeekofyear = datepart("ww", oRec("jobstartdato"))
			strWeekofyear_slut = datepart("ww", oRec("jobslutdato"))  
			
			w = 1
			for w = 1 to 53 
				if (cdate(oRec("jobstartdato")) < cdate(useLowerdatoKri) AND cint(strWeekofyear_slut) > cint(w)) OR (cint(strWeekofyear) <= cint(w) AND datepart("yyyy", oRec("jobstartdato")) <= datepart("yyyy", useLowerdatoKri) AND cint(strWeekofyear_slut) >= cint(w)) OR (cint(strWeekofyear) <= cint(w) AND datepart("yyyy", oRec("jobslutdato")) > datepart("yyyy", useUpperdatoKri)) then%>
				<td bgcolor="#FFFFFF" valign="top"><%=useColorgif%></td>
				<%else%>
				<td valign="top" width="6" bgcolor="#FFFFFF">&nbsp;</td>
				<%end if
			next
			
			'end if%>
			<td width=210 colspan=2 bgcolor="#eff3ff" align="right" style="padding-left:5; padding-right:5;">&nbsp;<font size=1><%=formatdatetime(oRec("jobstartdato"), 0)%> <img src="../ill/pil_lysblaa.gif" width="11" height="6" alt="" border="0"> <%=formatdatetime(oRec("jobslutdato"), 0)%></td>
		</tr>
		<%
	oRec.movenext
	wend
	oRec.close
	%>	
	</table>
	<br><br>
	<br>
	<%if request("print") <> "j" then%>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<%end if%>
	<br>
	<br>
	</div>
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
