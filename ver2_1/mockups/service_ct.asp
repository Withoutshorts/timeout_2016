<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<!--include file="../inc/regular/header_lysblaa_inc.asp"-->
<!--include file="../inc/regular/topmenu_inc.asp"-->


<script>
function showbesk(){
document.getElementById("besk").style.visibility = "visible"
document.getElementById("besk").style.display = ""
}

function hidebesk(){
document.getElementById("besk").style.visibility = "hidden"
document.getElementById("besk").style.display = "none"
}

</script>




<div id=besk style="position:absolute; left:300px; width:400px; top:200px; display:none; visibility:hidden; padding:20px; background-color:#ffff99; border:1px #5582d2 solid; z-index:1000;">
07.30 - 08.45<br><b>Ole Nielsen</b><br>
Vesterbrogade 45 
<br>Vesterbrogade 45<br>1571 Kbh. V<br>Tlf: 32840600
<br />[<a href="#" onclick="hidebesk()">Luk</a>]
</div>




<br /><br />
<table cellpadding=5 cellspacing=2 border=1>
<tr>
<td>
<a href="service_ct.asp?clicked=1">
0.01 Kalender</a></td>
<td>
<a href="budget_ct.asp?clicked=2">
0.02 ServiceDesk</a></td>

</tr>

</table>

<br />
<br /><br />

<%

clicked = request("clicked")

select case clicked
case 1
%>
<table cellspacing=1 cellpadding=2 border=1><tr>
<td bgcolor="lightpink">Aflsuttet</td>
<td bgcolor="#ffff99">Igang</td>
<td bgcolor="silver">Venter</td>
</tr></table>

<br /><br />

Medarbejder: 
<select id="Select1">
    <option>Alle</option>
    <option>Jan Nielsen</option>
    <option>Ole Petersen</option>
     <option>Morten Eriksen</option>
</select>
&nbsp;
Vælg periode: 
<select id="Select1">
    <option>2007</option>
    <option>2008</option>
    <option>2006</option>
</select>

<select id="Select1">
    <option>Okt</option>
    <option>Nov</option>
    <option>Dec</option>
</select>
<input id="Submit1" type="submit" value="submit" />

<br />


<div style="position:relative; top:20px; padding:10px; border:1px #000000 solid; width:1000px; height:100px; border:1px; overflow:auto;">
<table cellpadding=2 cellspacing=1 border=1>
<tr>
<td>Medarbejder</td>
<td width=80><font size=1>Man 23/10-2007</td>
<td width=80><font size=1>Tir 24/10-2007</td>
<td width=80><font size=1>Ons 25/10-2007</td>
<td width=80><font size=1>Tor 26/10-2007</td>
<td width=80><font size=1>Fre 27/10-2007</td>
<td width=80><font size=1>Lør 28/10-2007</td>
<td width=80><font size=1>Søn 29/10-2007</td>
<td width=80><font size=1>Man 30/10-2007</td>
<td width=80><font size=1>Tir 31/10-2007</td>
<td width=80><font size=1>Ons 01/11-2007</td>
<td width=80><font size=1>Tor 02/11-2007</td>
<td width=80><font size=1>Fre 03/11-2007</td>
<td width=80><font size=1>Lør 04/11-2007</td>
<td width=80><font size=1>Søn 05/11-2007</td>
<td><b>>></b></td>
</tr>
<tr>
<td valign=top>Jan Nielsen</td>
<%for x = 0 to 13

select case x 
case 5


bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

bgtxt = bgtxt &  "<br><div style='position:relative; background-color:#ffff99; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"


bgtxt = bgtxt &  "<br><div style='position:relative; background-color:silver; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 2

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 7

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case else

bgtxt = "&nbsp;"
end select %>

<td valign=top><%=bgtxt %></td>
<%next %>
</tr>


</table>
</div>


<div style="position:relative; top:20px; padding:10px; border:1px #000000 solid; width:1000px; height:100px; border:1px; overflow:auto;">
<table cellpadding=2 cellspacing=1 border=1>
<tr>
<td>Medarbejder</td>
<td width=80><font size=1>Man 23/10-2007</td>
<td width=80><font size=1>Tir 24/10-2007</td>
<td width=80><font size=1>Ons 25/10-2007</td>
<td width=80><font size=1>Tor 26/10-2007</td>
<td width=80><font size=1>Fre 27/10-2007</td>
<td width=80><font size=1>Lør 28/10-2007</td>
<td width=80><font size=1>Søn 29/10-2007</td>
<td width=80><font size=1>Man 30/10-2007</td>
<td width=80><font size=1>Tir 31/10-2007</td>
<td width=80><font size=1>Ons 01/11-2007</td>
<td width=80><font size=1>Tor 02/11-2007</td>
<td width=80><font size=1>Fre 03/11-2007</td>
<td width=80><font size=1>Lør 04/11-2007</td>
<td width=80><font size=1>Søn 05/11-2007</td>
<td><b>>></b></td>
</tr>
<tr>
<td valign=top>Jan Nielsen</td>
<%for x = 0 to 13

select case x 
case 5


bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

bgtxt = bgtxt &  "<br><div style='position:relative; background-color:#ffff99; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"


bgtxt = bgtxt &  "<br><div style='position:relative; background-color:silver; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 2

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 7

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case else

bgtxt = "&nbsp;"
end select %>

<td valign=top><%=bgtxt %></td>
<%next %>
</tr>


</table>
</div>


<div style="position:relative; top:20px; padding:10px; border:1px #000000 solid; width:1000px; height:100px; border:1px; overflow:auto;">
<table cellpadding=2 cellspacing=1 border=1>
<tr>
<td>Medarbejder</td>
<td width=80><font size=1>Man 23/10-2007</td>
<td width=80><font size=1>Tir 24/10-2007</td>
<td width=80><font size=1>Ons 25/10-2007</td>
<td width=80><font size=1>Tor 26/10-2007</td>
<td width=80><font size=1>Fre 27/10-2007</td>
<td width=80><font size=1>Lør 28/10-2007</td>
<td width=80><font size=1>Søn 29/10-2007</td>
<td width=80><font size=1>Man 30/10-2007</td>
<td width=80><font size=1>Tir 31/10-2007</td>
<td width=80><font size=1>Ons 01/11-2007</td>
<td width=80><font size=1>Tor 02/11-2007</td>
<td width=80><font size=1>Fre 03/11-2007</td>
<td width=80><font size=1>Lør 04/11-2007</td>
<td width=80><font size=1>Søn 05/11-2007</td>
<td><b>>></b></td>
</tr>
<tr>
<td valign=top>Jan Nielsen</td>
<%for x = 0 to 13

select case x 
case 5


bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

bgtxt = bgtxt &  "<br><div style='position:relative; background-color:#ffff99; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"


bgtxt = bgtxt &  "<br><div style='position:relative; background-color:silver; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 2

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 7

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case else

bgtxt = "&nbsp;"
end select %>

<td valign=top><%=bgtxt %></td>
<%next %>
</tr>


</table>
</div>

<div style="position:relative; top:20px; padding:10px; border:1px #000000 solid; width:1000px; height:100px; border:1px; overflow:auto;">
<table cellpadding=2 cellspacing=1 border=1>
<tr>
<td>Medarbejder</td>
<td width=80><font size=1>Man 23/10-2007</td>
<td width=80><font size=1>Tir 24/10-2007</td>
<td width=80><font size=1>Ons 25/10-2007</td>
<td width=80><font size=1>Tor 26/10-2007</td>
<td width=80><font size=1>Fre 27/10-2007</td>
<td width=80><font size=1>Lør 28/10-2007</td>
<td width=80><font size=1>Søn 29/10-2007</td>
<td width=80><font size=1>Man 30/10-2007</td>
<td width=80><font size=1>Tir 31/10-2007</td>
<td width=80><font size=1>Ons 01/11-2007</td>
<td width=80><font size=1>Tor 02/11-2007</td>
<td width=80><font size=1>Fre 03/11-2007</td>
<td width=80><font size=1>Lør 04/11-2007</td>
<td width=80><font size=1>Søn 05/11-2007</td>
<td><b>>></b></td>
</tr>
<tr>
<td valign=top>Jan Nielsen</td>
<%for x = 0 to 13

select case x 
case 5


bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

bgtxt = bgtxt &  "<br><div style='position:relative; background-color:#ffff99; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"


bgtxt = bgtxt &  "<br><div style='position:relative; background-color:silver; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 2

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 7

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case else

bgtxt = "&nbsp;"
end select %>

<td valign=top><%=bgtxt %></td>
<%next %>
</tr>


</table>
</div>



<%
case else
end select

%>


<!--


<tr>
<td valign=top>Ole Petersen</td>
<%for x = 0 to 13

select case x 
case 2


bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

bgtxt = bgtxt &  "<br><div style='position:relative; background-color:#ffff99; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"


bgtxt = bgtxt &  "<br><div style='position:relative; background-color:silver; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 5

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 6

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case else

bgtxt = "&nbsp;"
end select %>

<td valign=top><%=bgtxt %></td>
<%next %>
</tr>


<tr>
<td valign=top>Morten Eriksen</td>
<%for x = 0 to 13

select case x 
case 1


bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

bgtxt = bgtxt &  "<br><div style='position:relative; background-color:#ffff99; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"


bgtxt = bgtxt &  "<br><div style='position:relative; background-color:silver; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>09.30 - 10.45<br><b>Kaj Jensen</b><br>"
bgtxt = bgtxt & "Elmegade 1 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 9

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case 12

bgtxt = "<div style='position:relative; background-color:lightpink; padding:2px; border:1px #000000 solid;'>"
bgtxt = bgtxt & "<font size=1>07.30 - 08.45<br><b>Ole Nielsen</b><br>"
bgtxt = bgtxt & "Vesterbrogade 45 [<a href='#' onClick='showbesk()'>Besk</a>]"
bgtxt = bgtxt & "</div>"

case else

bgtxt = "&nbsp;"
end select %>

<td valign=top><%=bgtxt %></td>
<%next %>
</tr>
-->




<!--include file="../inc/regular/footer_inc.asp"-->
