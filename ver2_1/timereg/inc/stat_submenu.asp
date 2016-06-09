<%if request("print") <> "j" then
	select case thisfile
	case "joblog", "stat_pies", "joblog_korsel"
	psleft = 835
	pstop = 160
	case "joblog_status"
	psleft = 640
	pstop = 168
	case "joblog_z"
	psleft = 755
	pstop = 160
	case "joblog_z_b"
	psleft = 800
	pstop = 108
	case "fak_osigt"
	psleft = 735
	pstop = 160
	case "oms"
	psleft = 835
	pstop = 150
	end select
else
	select case thisfile
	case "joblog", "stat_pies", "joblog_korsel", "joblog_status"
	psleft = 420
	pstop = 10
	case "joblog_z"
	psleft = 420
	pstop = 10
	case "joblog_z_b"
	psleft = 420
	pstop = 10
	case "fak_osigt"
	psleft = 420
	pstop = 10
	case "oms"
	psleft = 520
	pstop = 50
	end select
end if
%>

<div id="submenu" style="position:absolute; left:<%=psleft%>; top:<%=pstop%>; visibility:visible; z-index:1000;">
<%if request("print") <> "j" then%>
	<table cellspacing="0" cellpadding="0" border="0">
	<tr bgcolor="5582D2">
			<td width="8" rowspan="2" valign=top><img src="../ill/tabel_top_left.gif" width="8" height="23" alt="" border="0"></td>
			<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="134" height="1" alt="" border="0"></td>
			<td align=right rowspan="2" valign=top><img src="../ill/tabel_top_right.gif" width="8" height="23" alt="" border="0"></td>
	</tr>
	<tr bgcolor="5582D2">
			<td colspan=2 valign="top" class="alt"><b>Oversigter:</b></td>
	</tr>
	<tr bgcolor="#EFF3FF">
	<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="250" alt="" border="0"></td>
	<td colspan="2" valign="top"><br>
		<%if FM_job <> "909909909909" then%>
		<a href="stat_pies.asp?menu=stat&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>" class='rmenu'><img src="../ill/pie_ikon.gif" width="23" height="22" alt="" border="0">&nbsp;Top 5 og % Fordeling</a>
		<br>
		<a href="oms.asp?menu=stat&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>" class='rmenu'><img src="../ill/oms.gif" width="23" height="22" alt="" border="0">&nbsp;Årsomsætning</a>
		<br>
		<a href="word.asp?menu=stat&weekselector=<%=weekselector%>&strWSelStartDato=<%=strWSelStartDato%>&strWSelEndDato=<%=strWSelEndDato%>&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>" class='rmenu'><img src="../ill/word.gif" width="23" height="22" alt="" border="0">&nbsp;Eksport (Word, Excel)</a>
		<br>
		<!-- Fakturaer 
		<%'if strE5 = "E5 Fakturering on" then%>
		<a href="fak_osigt.asp?menu=stat&mrd=<=strReqMrd%>&jobnr=<=intJobnr%>&eks=<=request("eks")%>&year=<=strReqAar%>&lastFakdag=<=lastFakdag%>&selmedarb=<=selmedarb%>&selaktid=<=selaktid%>&FM_job=<=request("FM_job")%>&FM_medarb=<=request("FM_medarb")%>&FM_aar=<=request("FM_aar")%>" class='rmenu'><img src=../ill/fak.gif width="23" height="22" alt="" border="0">&nbsp;Fakturaer</a>
		<br>
		<%'end if%>
		Slut Fak -->
		<a href="joblog.asp?menu=stat&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>" class='rmenu'><img src="../ill/joblog.gif" alt="Joblog" border="0" width="23" height="22">&nbsp;Joblog</a>
		<br>
		<a href="joblog_korsel.asp?menu=stat&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>" class='rmenu'><img src="../ill/korsel.gif" alt="Kørsel" border="0" width="23" height="22">&nbsp;Kørsel</a>
		<br>
		<a href="joblog_status.asp?menu=stat&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>" class='rmenu'><img src="../ill/status.gif" alt="Status" border="0" width="23" height="22">&nbsp;Nøgletal og Status</a>
		<%end if%>
	<br>
	<a href="joblog_z_b.asp?menu=stat&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>" class='rmenu'><img src="../ill/sum.gif" alt="Joblog" border="0" width="23" height="22">&nbsp;Timefordeling</a>
	<br>
	<%'end if%>
	<%
	'*** Printervenlig ***
	if thisfile <> "joblog_z_b" then%>
	<br><a href="javascript:NewWin('<%=thisfile%>.asp?menu=stat&print=j&mrd=<%=strReqMrd%>&jobnr=<%=intJobnr%>&eks=<%=request("eks")%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_aar=<%=request("FM_aar")%>')" target="_self" class='rmenu'>&nbsp;Printer venlig version&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	<%else
	if seljob = -1 then
	usepseljob = 0
	else
	usepseljob = seljob
	end if
	%>
	<br><a href="javascript:NewWin_large('joblog_z_b.asp?menu=stat&print=j&FM_seljob=<%=usepseljob%>&FM_job=<%=request("FM_job")%>&FM_medarb=<%=request("FM_medarb")%>&FM_start_dag=<%=strDag%>&FM_start_mrd=<%=strMrd%>&FM_start_aar=<%=strAar%>&FM_slut_dag=<%=strDag_slut%>&FM_slut_mrd=<%=strMrd_slut%>&FM_slut_aar=<%=strAar_slut%>&FM_vis_medarb=<%=vis_medarb%>&FM_vis_akt=<%=vis_akt%>&FM_vis_medarb_k=<%=vis_medarb_k%>')" target="_self" class='rmenu'>&nbsp;Printer venlig version&nbsp;<img src="../ill/pillillexp_tp.gif" width="16" height="18" alt="" border="0"></a>
	<%end if%>&nbsp;</td>
	<td valign="top" align="right"><img src="../ill/tabel_top.gif" width="1" height="250" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="2" valign="bottom"><img src="../ill/tabel_top.gif" width="134" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	<%else%>
	<a href="javascript:window.close()" class='rmenu'><img src="../ill/luk_xp.gif" width="30" height="28" alt="" border="0">&nbsp;Luk</a>
	&nbsp;&nbsp;
	<a href="javascript:window.print()" class='rmenu'><img src="../ill/print_xp.gif" width="28" height="30" alt="" border="0">&nbsp;Print</a>
	<%end if%>
</div>


