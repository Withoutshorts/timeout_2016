<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<!--include file="../inc/regular/header_lysblaa_inc.asp"-->
<!--include file="../inc/regular/topmenu_inc.asp"-->


<br /><br />
<table cellpadding=5 cellspacing=2 border=1>
<tr>
<td>
<a href="budget_ct.asp?clicked=1">
0.01 Opret brutto / netto timer</a></td>
<td>
<a href="budget_ct.asp?clicked=2">
0.02 Opret budgetter</a></td>
<td>
<a href="budget_ct.asp?clicked=3">0.03 Opfølgning årsoversigt md/md</a></td>
<td><a href="budget_ct.asp?clicked=4">0.04 Månedspuls uge/uge</a></td>
<td>0.05 Nøgletal</td>
<td>0.06 Prognose</td>
</tr>

</table>

<br />
<br /><br />

<%

clicked = request("clicked")

select case clicked
case 1
%>
Opret brutto og netto timer for år:

<select id="Select1">
    <option>2005</option>
    <option>2006</option>
    <option>2007</option>
    <option>2008</option>
</select>
<input id="Submit1" type="submit" value="submit" />
<br />
Tal efter text felter er tal baseret på system oplysning. Fradrag er baseret på registrering i timeOut (alle år) fordelt på de enkelte måender
<table cellpadding=2 cellspacing=2 border=1>
<tr>
<td>Type</td>
<td>Jan</td>
<td>Feb</td>
<td>Mar</td>
<td>Apr</td>
<td>Maj</td>
<td>Jun</td>
<td>Jul</td>
<td>Aug</td>
<td>Sep</td>
<td>Okt</td>
<td>Nov</td>
<td>Dec</td>
</tr>

<tr>
<td>Brutto dage</td>
<%for x = 0 to 11 %>
<td><input id="Text1" type="text" size=5 value=20 /><font color=silver size=1> 20</font></td>
<%next %>
</tr>

<tr><td colspan=13>Fradrag</td></tr>

<tr>
<td>Ferie</td>
<%for x = 0 to 11 %>
<td><input id="Text1" type="text" size=5 value=<%=x*5 %> /> <font color=silver size=1>11</td>
<%next %>
</tr>
<tr>
<td>Sygdom</td>
<%for x = 0 to 11 %>
<td><input id="Text1" type="text" size=5 value=<%=x*1 %> /> <font color=silver size=1>3</td>
<%next %>
</tr>
<tr>
<td>Uddannelse</td>
<%for x = 0 to 11 %>
<td><input id="Text1" type="text" size=5 value=<%=(x*0.5) %> /> <font color=silver size=1>2</td>
<%next %>
</tr>

<tr>
<td>Netto</td>
<%for x = 0 to 11 %>
<td><input id="Text1" type="text" size=5 value=17 /></td>
<%next %>
</tr>

<tr><td colspan=13>Timer</td></tr>

<tr>
<td>Brutto timer</td>
<%for x = 0 to 11 %>
<td><input id="Text1" type="text" size=5 value=170 /></td>
<%next %>
</tr>

<tr>
<td>Netto timer</td>
<%for x = 0 to 11 %>
<td><input id="Text1" type="text" size=5 value=116 /></td>
<%next %>
</tr>

</table>
Gem brutto / netto tal for det valgte år: <input id="Submit1" type="submit" value="submit" />

<%
case 2

%>


Opret nyt budget

<select id="Select1">
    <option>Vælg budget</option>
    <option>06</option>
    <option>07 - Pessimistisk</option>
    <option>07 - Optimistisk</option>
</select>
<input id="Submit1" type="submit" value="submit" />
<br />
Budgettrede belægning. Angivet i %.
<table cellpadding=2 cellspacing=2 border=1>
<tr>
<td>Medarbejdere</td>
<td>Jan</td>
<td>Feb</td>
<td>Mar</td>
<td>Apr</td>
<td>Maj</td>
<td>Jun</td>
<td>Jul</td>
<td>Aug</td>
<td>Sep</td>
<td>Okt</td>
<td>Nov</td>
<td>Dec</td>
</tr>

<%for y = 0 to 7 %>
<tr>
<td>Jan Larzsen</td>

<%for x = 0 to 11 %>
<td><input id="Text1" type="text" size=5 value=<%=x*5 %> /></td>
<%next %>
</tr>
<%next %>


</table>

Slet budget.. | Gem budget som : 
<input id="Text1" type="text" /><input id="Submit1" type="submit" value="submit" />
<%
case 3
%>


Opfølgning årsoversigt md/md

<select id="Select1">
    <option>Vælg budget</option>
    <option>06</option>
    <option>07 - Pessimistisk</option>
    <option>07 - Optimistisk</option>
</select>
&nbsp;&nbsp;&nbsp;
Oversigt:
<select id="Select1">
    <option>Afsætning</option>
    <option>Omsætning</option>
    <option>Belægning</option>
    <option>Forecast</option>
</select>

<input id="Submit1" type="submit" value="submit" />
<br />

<table cellpadding=2 cellspacing=2 border=1>

<tr>
<td>Medarbejdere</td>
<td>ÅTD<br />
<font size=1>Budget - Realiseret - Afv</font></td>
<td colspan=2>Jan<br />
<font size=1>Budget - Realiseret</font></td>
<td colspan=2>Feb</td>
<td colspan=2>Mar</td>
<td colspan=2>Apr</td>
<td colspan=2>Maj</td>
<td colspan=2>Jun</td>
<td colspan=2>Jul</td>
<td colspan=2>Aug</td>
<td colspan=2>Sep</td>
<td colspan=2>Okt</td>
<td colspan=2>Nov</td>
<td colspan=2>Dec</td>
</tr>

<%for y = 0 to 7 %>
<tr>
<td>Jan Larzsen</td>
<td>123 - 214 - 65</td>

<%for x = 0 to 11 
select case x
case 2,6,9
bgc = "red"
case else
bgc = "lightgreen"
end select
%>
<td>123</td><td bgcolor="<%=bgc %>">56</td>
<%next %>
</tr>
<%next %>


<tr>
<td><b>Total</b></td>
<td>123 - 214 - 65</td>

<%for x = 0 to 11 
select case x
case 2,6,9
bgc = "red"
case else
bgc = "lightgreen"
end select
%>
<td>123</td><td bgcolor="<%=bgc %>">56</td>
<%next %>
</tr>



</table>

Eskporter | Print 

<%
case 4
%>




PULS September 2007, 23/9
<br /><br />
Uge: <u>36</u> - <u>37</u> - <u>38</u> - <u>39</u>&nbsp;
<select id="Select1">
    <option>Vælg budget</option>
    <option>06</option>
    <option>07 - Pessimistisk</option>
    <option>07 - Optimistisk</option>
</select>


<input id="Submit1" type="submit" value="submit" />
<br />

<table cellpadding=2 cellspacing=2 border=1>

<tr>
<td>Medarbejdere</td>

<td>Budget</td>
<td>Netto</td>
<td>Real tot.<br />
<font size=1>(fak + ej fakbare timer)</font></td>

<td colspan=2>Afsætning MTD<br />
<font size=1>Budget - Realiseret</font></td>
<td colspan=2>Omsætning MTD<br />
<font size=1>Budget - Realiseret</font></td>
<td colspan=2>Omsætning måned tot.<br />
<font size=1>Budget - Realiseret</font></td>
<td colspan=2>Forecast måned tot.<br />
<font size=1>Budget - Realiseret</font></td>
<td colspan=2>Timepriser<br />
<font size=1>Norm - Realiseret</font></td>
</tr>

<%for y = 0 to 7 %>
<tr>
<td>Jan Larzsen</td>
<td>85%</td>
<td>110</td>
<td>149</td>

<%for x = 0 to 4 
select case x
case 0
v1 = 93
v2 = 109
case 1
v1 = 123000
v2 = 119000
case 2
v1 = 193000
v2 = 129000
case 3
v1 = 76
v2 = 89
case 4
v1 = 1300
v2 = 1550
end select

select case x
case 1,2
bgc = "red"
case else
bgc = "lightgreen"
end select
%>
<td><%=v1 %></td><td bgcolor="<%=bgc %>"><%=v2 %></td>
<%next %>
</tr>
<%next %>






</table>

Eskporter | Print 


<%
case else
end select

%>





<!--include file="../inc/regular/footer_inc.asp"-->
