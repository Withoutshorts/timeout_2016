<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_hvd_inc.asp"-->	
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/dato.asp"-->

<%
usemrn = request("mid")
func = request("func")

if func = "updguide" then
		
		
		oConn.execute("DELETE FROM timereg_usejob WHERE medarb = "& usemrn &"")
		
		useJob = request("FM_use_job")
		j = 0
		intuseJob = Split(useJob, ", ")
	   	For j = 0 to Ubound(intuseJob)
		strSQL = "INSERT INTO timereg_usejob (medarb, jobid) VALUES ("& usemrn &", "& intuseJob(j) &")"
		'Response.write strSQL & "<br>"
		oConn.execute(strSQL)
		next
		
		Response.flush
		
		'**** Bruges ikke mere timreg_2006 ***
		'if request("FM_visguide") = "j" then
			oConn.execute("UPDATE medarbejdere SET visguide = 1 WHERE Mid = "& usemrn &"")
			'else
			'oConn.execute("UPDATE medarbejdere SET visguide = 0 WHERE Mid = "& usemrn &"")
		'end if
		
		Response.Write("<script language=""JavaScript"">window.opener.top.frames['t'].location.reload();</script>")
		Response.Write("<script language=""JavaScript"">window.opener.top.frames['a'].location.reload();</script>")
		Response.Write("<script language=""JavaScript"">window.close();</script>")


else

%>

<SCRIPT language=javascript src="inc/timereg_2006_func.js"></script>
<div id="sindhold" style="position:absolute; left:20; top:20; width:60%; height:600; visibility:visible;">
<%

'Response.write usemrn
'Response.flush

	
	'************************************************************************************************
	'**** Guiden... Vis Kunder ********
	'************************************************************************************************
	%>
	
	
	
	<h4>Guiden dine aktive job:</h4>
	Dette er din <b>Guide</b> til at vælge hvilke <b>job</b> du arbejder på iøjeblikket.<br>
	Det betyder at du næste gang du logger ind i TimeOut, <b>kun</b> vil få vist job der er tilvalgt her i <b>Guiden dine aktive job</b>.<br></td>
	
		
	<table cellspacing="0" cellpadding="0" border="0" width=500>
	<form action="guiden_2006.asp?func=updguide&mid=<%=usemrn%>" method="post" name="usejobguide">
 	<tr>
		<td valign="top">&nbsp;</td>
		<td colspan="2" align="right" valign=bottom>&nbsp;<a href="#" name="CheckAll" onClick="checkAll(document.usejobguide.FM_use_job)"><img src="../ill/alle.gif" border="0"></a>
		&nbsp;&nbsp;<a href="#" name="UnCheckAll" onClick="uncheckAll(document.usejobguide.FM_use_job)"><img src="../ill/ingen.gif" border="0"></a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<td valign="top" align="right">&nbsp;</td>
	</tr>
	<%
	
	call hentbgrppamedarb(usemrn)
	
	lastKid = 0
	strSQL = "SELECT j.id, j.jobstatus, j.jobnavn, j.jobnr, j.jobknr, Kid, Kkundenavn, Kkundenr, useasfak, g.jobid AS gjid FROM job j, kunder"_
	&" LEFT JOIN timereg_usejob g ON (g.medarb = "& usemrn &" AND g.jobid = j.id) WHERE jobstatus = 1 AND fakturerbart = 1 AND kunder.Kid = j.jobknr  "& strPgrpSQLkri &" GROUP BY j.id, kid ORDER BY Kkundenavn, jobnavn"
	'Response.write (strSQL)	
	'Response.flush
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
			
		if lastKid <> oRec("Kid") then
		%>
		<tr>
			<td colspan="4" bgcolor="#8caae6" height=30 class=alt style="border:1px #003399 solid;">&nbsp;&nbsp;<b><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr")%>)</b></td>
		</tr>
		<%else%>
		<tr>
			<td bgcolor="#5582D2" colspan="4"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
		</tr>
		<%end if
		
		'**** Er jobbet allerede valgt?
		'strSQL3 = "SELECT id, medarb, jobid FROM timereg_usejob WHERE medarb = "& usemrn &" AND jobid = " & oRec("id")
		'oRec3.open strSQL3, oConn, 3
		
		if len(trim(oRec("gjid"))) <> 0 then
		selJob = "CHECKED"
		else
		selJob = ""
		end if
		
		
		%>
		<tr>
			<td valign="top" style="border-left:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
			<td height=20>&nbsp;&nbsp;&nbsp;<b><%=oRec("jobnr")%>&nbsp;&nbsp;<%=oRec("jobnavn")%></b></td>
			<td><input type="checkbox" name="FM_use_job" value="<%=oRec("id")%>" <%=selJob%>></td>
			<td valign="top" style="border-right:1px #003399 solid;"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
		</tr>
		<%
		
		'Response.flush
		lastKid = oRec("Kid")
		oRec.movenext
		wend
		oRec.close
		%>
		
		<!--
		<tr bgcolor="#D6DFF5">
			<td colspan="4" align="right"><br>Vis denne guide næste gang du logger på?&nbsp;Ja:<input type="checkbox" name="FM_visguide" value="j"></td>
		</tr>-->
		<tr>
			<td style="border-top:1px #003399 solid;" colspan="4" align="right"><br>
			<input type="hidden" name="FM_use_job" id="FM_use_job" value="0">
			<input type="image" src="../ill/opdaterpil.gif"></td>
		</tr>
		</form>
		</table>
	<br><br>&nbsp;
	
</div>
<!--#include file="../inc/regular/footer_inc.asp"-->

<%end if%>
