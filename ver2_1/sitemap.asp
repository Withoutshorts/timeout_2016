<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
<style>
td {
font-size:11px;
font-family:tahoma;
vertical-align:top;
}
</style>
    <title>TimeOut Sitemap </title>
</head>
<body>

<table cellspacing=1 cellpadding=2 border=0 width=90% bgcolor="#cccccc">
<tr>
<td>Modul</td>
<td>1. Menupkt.</td>
<td>2. Menupkt.</td>
<td>Frameset</td>
<td>Fil</td>
<td>Include</td>
</tr>
<%tsabgcolor="#ffffe1" %>
<tr bgcolor="pink"><td colspan=6>TSA</td></tr>
<tr bgcolor="#d6dff5">
<td>1.01</td>
<td>Timeregistrering</td>
<td></td>
<td>timereg_2006_fs.asp</td>
<td>timereg_2006.asp<br />
timereg_akt_2006.asp<br />
timereg_2007_top.asp
</td>
<td>
include file="../inc/connection/conn_db_inc.asp"<br />
include file="../inc/errors/error_inc.asp"<br />
include file="../inc/regular/global_func.asp"<br />
include file="inc/dato.asp"<br />
include file="inc/convertDate.asp"<br />
include file="../inc/regular/header_inc.asp"<br />
include file="../inc/regular/header_lysblaa_inc.asp"<br />
include file="inc/timereg_func_inc.asp"

</td>
</tr>
<tr bgcolor="<%=tsabgcolor %>">
<td>1.02</td>
<td></td>
<td>Indtastning af timer</td>
<td>timereg_2006_fs.asp</td>
</tr>
<tr bgcolor="<%=tsabgcolor %>">
<td>1.03</td>
<td></td>
<td>Se timeregistreringer (valgt uge)</td>
<td>timereg_2006_fs.asp</td>
</tr>

<tr bgcolor="<%=tsabgcolor %>">
<td>1.04</td>
<td></td>
<td>Joblog</td>
<td></td>
<td>joblog.asp</td>
</tr> 

<tr bgcolor="<%=tsabgcolor %>">
    <td>1.05</td>
    <td></td>
    <td>Timeforbrug - Grandtotal</td>
    <td></td>
    <td>joblog_timetotaler.asp</td>
</tr>    

<tr bgcolor="<%=tsabgcolor %>">
    <td>1.06</td>
    <td></td>
    <td>Guid. aktive job</td>
    <td></td>
    <td>guiden_2006.asp</td>
</tr>  

<tr bgcolor="<%=tsabgcolor %>">
    <td>1.07</td>
    <td></td>
    <td>Stopur</td>
    <td></td>
    <td>stopur_2006.asp</td>
</tr> 


<%crmbgcolor="#f3efff" %>
<tr bgcolor="<%=crmbgcolor %>"><td>CRM</td></tr>
<tr><td>ERP</td></tr>
<tr><td>SDSK</td></tr>
</table>


</body>
</html>
