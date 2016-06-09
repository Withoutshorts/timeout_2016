<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<!--include file="../inc/regular/header_inc.asp"-->
<!--include file="../inc/regular/topmenu_inc.asp"-->

<script>

function tilfoj(){
document.getElementById("tilfojlinie").style.display = ""
document.getElementById("tilfojlinie").style.visibility = "visible"
}

</script>

<%
clicked = request("clicked")
%>

<br /><br />
<table cellpadding=5 cellspacing=2 border=1>
<tr>
<td>
<a href="crm_ct.asp?clicked=1">
0.01 Salgsstyrring leads</a></td>
<td>
<td>
<a href="crm_ct.asp?clicked=2">
0.02 Salgsstyrring statistik</a></td>
<td>
<td>
<a href="crm_ct.asp?clicked=3">
0.03 Salgsstyrring historik</a></td>
<td>
<td>
<a href="crm_ct.asp?clicked=4">
0.04 Salgsstyrring Kalender</a></td>
<td>
</tr>

</table>

<br />
<br /><br />


<%select case clicked
case 4
%>
Historik vises som kalender der kan synkroniseres med Outlook.
<%
case 3
%>
<h4>0.03 Salgsstyrring historik</h4>

<b>Periode:</b> 
<input id="Text1" type="text" value="21.12.2007" /> til <input id="Text1" type="text" value="29.01.2008" /><br />

<b>Filter:</b> 
Ansvarlig  <select id="Select1" style="width:100px;">
 <option>Alle</option>
        <option>Morten Hald Mortensen</option>
        <option>Henrik La Cour</option>
          <option>Sarah Diderichsen</option>
        <option>Henrik La Cour</option>
        <option>Morten Hald Mortensen</option>
        <option>Henrik La Cour</option>
          <option>Sarah Diderichsen</option>
        <option>Henrik La Cour</option>
 </select> 
   

Søg på kontakt: 
<input id="Text1" type="text" /><br />Søg på lead Id: 
<input id="Text1" type="text" value="100" />&nbsp;<input id="Checkbox1" type="checkbox" /> Vis fuld historik på valgte Id.    
<input id="Submit1" type="submit" value="Vis liste" />
<br /><br />


<br /><br />
<table width=600 cellspacing=2 cellpadding=2>
<tr>

<td><b>Lead Id</b></td><td><b>Dato</b></td><td><b>Ansvarlig</b></td><td><b>Kontakt</b></td><td><b>Kommentar</b></td>

</tr>
<tr>
<td><a href="crm_ct.asp">100</a></td><td>10.01.2008</td><td>Morten Hald Mortensen</td><td>Post Danmark</td><td>Ring til kontakt, spørg efter Jens.</td>
</tr>
<tr>
<td><a href="crm_ct.asp">100</a></td><td>04.01.2008</td><td>Morten Hald Mortensen</td><td>Post Danmark</td><td>Har ringet igen</td>
</tr>
<tr>
<td><a href="crm_ct.asp">100</a></td><td>04.01.2008</td><td>Morten Hald Mortensen</td><td>Post Danmark</td><td>Har ringet igen</td>
</tr>
<tr bgcolor="#ffff99">
<td valign=top><a href="crm_ct.asp">100</a></td>
<td valign=top>21.12.2007</td>
<td valign=top>Morten Hald Mortensen</td>
<td valign=top>Post Danmark</td><td>Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s.</td>
</tr>
<tr>
<td>
    <input id="Text1" type="text" value="100" />
</td>
<td>
    <input id="Text1" type="text" value="10.01.2008" />
</td>
<td>
    <select id="Select1" style="width:100px;">
 <option>Morten Hald Mortensen</option>
        <option>Henrik La Cour</option>
          <option>Sarah Diderichsen</option>
        <option>Henrik La Cour</option>
        <option>Morten Hald Mortensen</option>
        <option>Henrik La Cour</option>
          <option>Sarah Diderichsen</option>
        <option>Henrik La Cour</option>
    </select> 
</td>
<td>
    <input id="Text1" type="text" value="Post Danmark" />
</td>
<td>
    <input id="Text1" type="text" />&nbsp;<input id="Submit1" type="submit" value="Tilføj" />
</td>
</tr>
</table>


<%case 2%>
<h4>0.01 Salgsstyrring statistik</h4>


<b>Periode:</b> 
<select id="Select1">
<option>2007</option>
    <option>2008</option>
</select> <input id="Submit1" type="submit" value="Vis liste" />
<br /><br />

<table cellspacing=2 cellpadding=2 border=1>
<tr>
<td><b>Medarbejdere</b></td>
<td><b>Salg i periode</b></td>
<td><b>Forv. Omsætning</b></td>
<td><b>Aktive</b></td>
<td><b>Forv. Omsætning</b></td>
<td><b>Lukkede</b></td>
<td><b>Passive</b></td>
<td><b>Salg tot.</b></td>
<td><b>Forv. Omsætning tot.</b></td>
</tr>
<tr>
<td>Morten Hald Mortensen</td><td>25</td><td><%=formatcurrency(6250000)%></td><td>12</td><td><%=formatcurrency(2250000)%></td><td>10</td><td>4</td><td>25</td><td><%=formatcurrency(6250000)%></td>
</tr>
<tr>
<td>Henrik La Cour</td><td>20</td><td><%=formatcurrency(625000)%></td><td>15</td><td><%=formatcurrency(125000)%></td><td>11</td><td>1</td><td>25</td><td><%=formatcurrency(6250000)%></td>
</tr>
<tr>
<td>Sarah Diderichsen</td><td>15</td><td><%=formatcurrency(625000)%></td><td>11</td><td><%=formatcurrency(225000)%></td><td>5</td><td>2</td><td>25</td><td><%=formatcurrency(6250000)%></td>
</tr>
</table>

<%case else %>

<h4>0.01 Salgsstyrring leads</h4>

<b>Periode:</b> 
<input id="Text1" type="text" value="20.11.2007" /> til <input id="Text1" type="text" value="29.12.2007" /><br />

<b>Filter:</b> 
Ansvarlig  <select id="Select1" style="width:100px;">
 <option>Alle</option>
        <option>Morten Hald Mortensen</option>
        <option>Henrik La Cour</option>
          <option>Sarah Diderichsen</option>
        <option>Henrik La Cour</option>
        <option>Morten Hald Mortensen</option>
        <option>Henrik La Cour</option>
          <option>Sarah Diderichsen</option>
        <option>Henrik La Cour</option>
    </select> 
    Status  <select id="Select1">
    <option>Alle</option>
              <option>Aktiv</option>
              <option>Passiv</option>
             <option>Lukket</option>
              <option>Salg</option>
              </select>

Forventning <select id="Select1">
 <option>Alle</option>
              <option>Lav</option>
              <option>Mellem</option>
              <option>Høj</option>
              </select> <br /><br />
<b>Sortér efter:</b>
<input id="Radio1" type="radio" /> Status
<input id="Radio1" type="radio" /> Forventning 
<input id="Radio1" type="radio" /> Ansvarlig 
    <br />
Søg på kontakt: 
<input id="Text1" type="text" />&nbsp;     
<input id="Submit1" type="submit" value="Vis liste" />
<br /><br />
<table cellspacing=2 cellpadding=2 border=0 width=90%>
<tr>
<td><b>ID - Oprettet Dato</b></td>
<td><b>Ansvarlig</b></td>
<td><b>Beskrivelse</b></td>
<td><b>Kontakt oplys.</b></td>
<td><b>Status</b></td>
<td><b>Forventning</b></td>
<td><b>Forv. oms. p. a.</b></td>
<td><b>Aktions Historik</b><br /> Seneste/næste</td>
</tr>

<%for x = 0 to 2 %>
<tr>
    <td valign=top><%=100+x %> - 2<%=x+1%>.12.2007</td>
    <td valign=top>
    <select id="Select1" style="width:150px;">
        <option>Morten Hald Mortensen</option>
        <option>Henrik La Cour</option>
          <option>Sarah Diderichsen</option>
        <option>Henrik La Cour</option>
        <option>Morten Hald Mortensen</option>
        <option>Henrik La Cour</option>
          <option>Sarah Diderichsen</option>
        <option>Henrik La Cour</option>
    </select>
    </td>
    
    <td valign=top>
        <textarea id="TextArea1" cols="30" rows="4">Lorem Ipsum is simply dummy text of the printing and
 typesetting industry.
    </textarea>
      
        </td>
          <td>
          <%select case x 
          case 1
          knavn = "Kbh. lufthavn"
          kpnavn = "Jens Madsen"
          tlf = "32940005"
          email = ""
          case 2
          knavn = "KL"
          kpnavn = "Kurt Nielsen"
          tlf = ""
          email = ""
          case 3
          knavn = "TDC"
          kpnavn = "Ulle Nielsen"
          tlf = "32940005"
          email = "ole@dfgg.dk"
          case else
          knavn = "Post Danmark"
          kpnavn = "Morten Kristensen"
          tlf = "32940005"
          email = "ole@dfgg.dk"
          end select
          %>
              <input id="Text1" type="text" value="<%=knavn %>" style="width:150px; font-size:9px;" /><br />
               <input id="Text1" type="text" value="<%=kpnavn %>" style="width:150px; font-size:9px;" /><br />
               <input id="Text1" type="text" value="<%=tlf %>" style="width:150px; font-size:9px;" /><br /> 
               <input id="Text1" type="text" value="<%=email %>" style="width:150px; font-size:9px;" /><br />   
      
      </td>
      <td valign=top>
       <select id="Select1">
              <option>Aktiv</option>
              <option>Passiv</option>
             <option>Lukket</option>
              <option>Salg</option>
              </select>
      </td>
     
       <td valign=top>
       <select id="Select1">
              <option>Lav</option>
              <option>Mellem</option>
              <option>Høj</option>
              </select>
      </td>
       <td valign=top>
          <input id="Text1" type="text" value="280000" size=6 /> kr.</td>
        <td valign=top> 
      <%if x = 0 then %>
      <a href="crm_ct.asp?clicked=3">21.12.2007</a>
      <%else %>
      21.12.2007
      <%end if %>
         - <u>Tilføj</u>
      </td>
</tr>
<%next %>
</table>





<div id="tilfojlinie" style="display:none; visibility:hidden; position:relative;" >

<table cellspacing=2 cellpadding=2 border=0 width=90%>
<tr>
    <td valign=top><%=100+x %> - 2<%=x+1%>.12.2007</td>
    <td valign=top>
    <select id="Select1" style="width:150px;">
        <option>Morten Hald Mortensen</option>
        <option>Henrik La Cour</option>
          <option>Sarah Diderichsen</option>
        <option>Henrik La Cour</option>
        <option>Morten Hald Mortensen</option>
        <option>Henrik La Cour</option>
          <option>Sarah Diderichsen</option>
        <option>Henrik La Cour</option>
    </select>
    </td>
    
    <td valign=top>
        <textarea id="TextArea1" cols="30" rows="4"></textarea>
      
        </td>
          <td>
        
              <input id="Text1" type="text" value="Kontakt" style="width:150px; font-size:9px;" /><br />
               <input id="Text1" type="text" value="Kontakpers" style="width:150px; font-size:9px;" /><br />
               <input id="Text1" type="text" value="Tlf" style="width:150px; font-size:9px;" /><br /> 
               <input id="Text1" type="text" value="Email" style="width:150px; font-size:9px;" /><br />   
      
      </td>
      <td valign=top>
       <select id="Select1">
              <option>Aktiv</option>
              <option>Passiv</option>
             <option>Lukket</option>
              <option>Salg</option>
              </select>
      </td>
     
       <td valign=top>
       <select id="Select1">
              <option>Lav</option>
              <option>Mellem</option>
              <option>Høj</option>
              </select>
      </td>
       <td valign=top>
          <input id="Text1" type="text" value="280000" size=6 /> kr.</td>
        <td valign=top> 
      <u>21.12.2007</u> 
         - <u>Tilføj</u>
      </td>
</tr>

</table>

</div>

<a href="#" onclick="tilfoj();">Tilføj til liste</a>

<%end select %>

<!--include file="../inc/regular/footer_inc.asp"-->
