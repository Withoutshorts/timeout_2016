<!--
Bruges bl.a af:
service_osigt_inc.asp
posteringer.asp
kontoplan.asp
fak_serviceaft_osigt.asp
resultatop.asp

faktura filerne.
Stat filerne.

-->

<!--#include file="dato2_b.asp"-->
<td width=200 align=center><font size="1"><%=job_txt_331 %>:&nbsp;<select name="FM_start_dag">
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
		<option value="31">31</option></select>&nbsp;&nbsp;
		
		<select name="FM_start_mrd" >
		<option value="<%=strMrd%>"><%=strMrdNavn%></option>
		<option value="1"><%=job_txt_102 %></option>
	   	<option value="2"><%=job_txt_103 %></option>
	   	<option value="3"><%=job_txt_104 %></option>
	   	<option value="4"><%=job_txt_105 %></option>
	   	<option value="5"><%=job_txt_106 %></option>
	   	<option value="6"><%=job_txt_107 %></option>
	   	<option value="7"><%=job_txt_108 %></option>
	   	<option value="8"><%=job_txt_109 %></option>
	   	<option value="9"><%=job_txt_110 %></option>
	   	<option value="10"><%=job_txt_111 %></option>
	   	<option value="11"><%=job_txt_112 %></option>
	   	<option value="12"><%=job_txt_113 %></option></select>&nbsp;&nbsp;
		<select name="FM_start_aar" >
		<option value="<%=strAar%>"><%=strAar%></option>
		<%for x = -10 to 10 
		useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		<option value="<%=useY%>"><%=right(useY, 2)%></option>
		<%next %>
		</select>
		</td><td width=200><font size="1"><%=job_txt_332 %>:&nbsp;<select name="FM_slut_dag" >
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
		<option value="31">31</option></select>&nbsp;&nbsp;
		<select name="FM_slut_mrd" >
		<option value="<%=strMrd_slut%>"><%=strMrdNavn_slut%></option>
		<option value="1"><%=job_txt_102 %></option>
	   	<option value="2"><%=job_txt_103 %></option>
	   	<option value="3"><%=job_txt_104 %></option>
	   	<option value="4"><%=job_txt_105 %></option>
	   	<option value="5"><%=job_txt_106 %></option>
	   	<option value="6"><%=job_txt_107 %></option>
	   	<option value="7"><%=job_txt_108 %></option>
	   	<option value="8"><%=job_txt_109 %></option>
	   	<option value="9"><%=job_txt_110 %></option>
	   	<option value="10"><%=job_txt_111 %></option>
	   	<option value="11"><%=job_txt_112 %></option>
	   	<option value="12"><%=job_txt_113 %></option></select>&nbsp;&nbsp;
		<select name="FM_slut_aar" >
		<option value="<%=strAar_slut%>"><%=strAar_slut%></option>
		<%for x = -10 to 50 
		useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		<option value="<%=useY%>"><%=right(useY, 2)%></option>
		<%next %></select></td>
				
<!-- Weekselecter SLUT -->
