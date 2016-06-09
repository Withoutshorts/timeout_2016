<!-- Flere knapper , Weekselecter -->
<div id="weekselecter" style="position:absolute; top:272; left:10; width:120; height:80; z-index:100; background-color:#fff8dc; visibility:<%=visWeekSel%>;">
<!--#include file="dato2.asp"-->
<%
	if request("weekselector") = "j" then
		frauge = datepart("ww", strMrd & "/" & strDag & "/" & strAar, 2,2)
		tiluge = datepart("ww", strMrd_slut & "/" & strDag_slut & "/" & strAar_slut,2,2)
		showStrTdato = strDag & ".&nbsp;" & strMrdNavn & "&nbsp;" & strAar
		showStrUdato = strDag_slut & ".&nbsp;" & strMrdNavn_slut & "&nbsp;" & strAar_slut
	else
		frauge = 1
		tiluge = 52
	end if
%>
<table cellspacing="0" cellpadding="0" border="0">
<form action="joblog_z.asp?menu=stat&weekselector=j&mrd=0&jobnr=<%=intJobnr%>&eks=<%=eks%>&year=<%=strReqAar%>&lastFakdag=<%=lastFakdag%>&selmedarb=<%=selmedarb%>&selaktid=<%=selaktid%>&FM_job=<%=FM_job%>&FM_medarb=<%=FM_medarb%>&FM_aar=<%=FM_aar%>" method="post">
<tr bgcolor="#064CB9">
		<td colspan=5 valign="top"><img src="../ill/tabel_top.gif" width="144" height="1" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#064CB9">
		<td colspan=5 valign="top" class="alt"><br>&nbsp;&nbsp;<b>Vælg periode:</b></td>
	</tr>
	<tr bgcolor="#EFF3FF">
	<td rowspan="3"><img src="../ill/tabel_top.gif" width="1" height="110" alt="" border="0"></td>
	<td><font size="1">&nbsp;Fra:<br>&nbsp;<select name="FM_start_dag" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
		<option value="<%=strDag%>"><%=strDag%></option> 
		<option value="1">1</option>
	   	<option value="2">2</option>
	   	<option value="3">3</option>
	   	<option value="4">4</option>
	   	<option value="5">5</option>
	   	<option value="6">6</option>
	   	<option value="7">7</option>
	   	<option value="8">8</option>
	   	<option value="9">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select></td>
		<td><br>
		<select name="FM_start_mrd" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select></td>
		<td><br>
		<select name="FM_start_aar" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
		<%if strYear = "-1" then 
		strReqAar = strAar
		end if%>
		<option value="<%=strReqAar%>"><%=strReqAar%></option>
		<option value="02">02</option>
		<option value="03">03</option>
	   	<!--<option value="04">04</option>
	   	<option value="05">05</option>
		<option value="06">06</option>
		<option value="07">07</option>--></select></td>
		<td rowspan="3" align="right"><img src="../ill/tabel_top.gif" width="1" height="110" alt="" border="0"></td>
		</tr>
		<tr bgcolor="#EFF3FF">
		<td><font size="1">&nbsp;Til:<br>&nbsp;<select name="FM_slut_dag" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
		<option value="<%=strDag_slut%>"><%=strDag_slut%></option> 
	   	<option value="1">1</option>
	   	<option value="2">2</option>
	   	<option value="3">3</option>
	   	<option value="4">4</option>
	   	<option value="5">5</option>
	   	<option value="6">6</option>
	   	<option value="7">7</option>
	   	<option value="8">8</option>
	   	<option value="9">9</option>
	   	<option value="10">10</option>
	   	<option value="11">11</option>
	   	<option value="12">12</option>
	   	<option value="13">13</option>
	   	<option value="14">14</option>
	   	<option value="15">15</option>
	   	<option value="16">16</option>
	   	<option value="17">17</option>
	   	<option value="18">18</option>
	   	<option value="19">19</option>
	   	<option value="20">20</option>
	   	<option value="21">21</option>
	   	<option value="22">22</option>
	   	<option value="23">23</option>
	   	<option value="24">24</option>
	   	<option value="25">25</option>
	   	<option value="26">26</option>
	   	<option value="27">27</option>
	   	<option value="28">28</option>
	   	<option value="29">29</option>
	   	<option value="30">30</option>
		<option value="31">31</option></select></td>
		<td><br>
		<select name="FM_slut_mrd" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
		<option value="<%=strMrd_slut%>"><%=strMrdNavn_slut%></option>
		<option value="1">jan</option>
	   	<option value="2">feb</option>
	   	<option value="3">mar</option>
	   	<option value="4">apr</option>
	   	<option value="5">maj</option>
	   	<option value="6">jun</option>
	   	<option value="7">jul</option>
	   	<option value="8">aug</option>
	   	<option value="9">sep</option>
	   	<option value="10">okt</option>
	   	<option value="11">nov</option>
	   	<option value="12">dec</option></select></td>
		<td><br>
		<select name="FM_slut_aar" style="background-color : #ffffff; border : thin black; font : 10px verdana;">
		<%if strYear = "-1" then 
		strReqAar_slut = strAar_slut
		end if%>
		<option value="<%=strReqAar_slut%>"><%=strReqAar_slut%></option>
		<option value="02">02</option>
		<option value="03">03</option>
	   <!--<option value="04">04</option>
	   	<option value="05">05</option>
			<option value="06">06</option>
				<option value="07">07</option>--></select></td>
				</tr>
	<tr bgcolor="#EFF3FF">
	<td align="center" colspan="3"><br><input type="image" src="../ill/weeksel.gif" border="0"></td>
</tr>
<tr bgcolor="#EFF3FF">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="3" valign="bottom"><img src="../ill/tabel_top.gif" width="144" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr></form></table>
</div>
<!-- Weekselecter SLUT -->
