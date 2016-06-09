<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/smiley_inc.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	%>
	
	<%
	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
	%>
	
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<!--#include file="../inc/regular/vmenu.asp"-->
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:190; top:50; visibility:visible;">
	<img src="../ill/header_smil.gif" width="650" height="78" alt="" border="0">
	<br><br>
	<%
	if len(request("useyear")) <> 0 then
	useYear = request("useyear")
	else
	useYear = year(now)
	end if
	%>
	
	Valgt år: <b><%=useYear%></b><br><a href="smileystatus.asp?useyear=<%=year(now)-2%>" class=vmenu><%=year(now)-2%></a> |
	<a href="smileystatus.asp?useyear=<%=year(now)-1%>" class=vmenu><%=year(now)-1%></a> | 
	<a href="smileystatus.asp?useyear=<%=year(now)%>" class=vmenu><%=year(now)%></a>  | 
	<a href="smileystatus.asp?useyear=<%=year(now)+1%>" class=vmenu><%=year(now)+1%></a> | 
	<a href="smileystatus.asp?useyear=<%=year(now)+2%>" class=vmenu><%=year(now)+2%></a> 
	
	<br><br>Når en uge er afsluttet inden søndag kl. 24:00 modtager medarbejderen en glad Smiley, ellers modtager medarbejderen en sur Smiley.
	<br>Periode er altid det valgte år frem til dagsdato.<br>
	<br>
	
	
	<%
	call smileystatus(0)
	%>
	
	</div>
	
	
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
	
	
