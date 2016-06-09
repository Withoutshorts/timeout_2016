<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/webblik_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->





<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
    menu = request("menu")
    thisfile = "timerimperr"

    media = request("media")
    if media = "print" OR media = "export" then
        print = "j"
    end if

	func = request("func")
	
	
	
	
	
	
	
	
	
	'************ slut faste filter var **************		
	

	if media = "print" OR media = "export" then
	leftPos = 20
	topPos = 62
    %>
    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%
	else
	leftPos = 20
	topPos = 122
    %>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <!--#include file="../inc/regular/topmenu_inc.asp"-->

	
	
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<%
        call tsamainmenu(7)
    %>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
        call stattopmenu()
    %>
	</div>
	<%
	end if
	%>


	<%
	pleft = 20
	ptop = 132
	'ptopgrafik = 348


 	%>	
	<div id="Div1" style="position:absolute; left:<%=pleft%>; top:<%=ptop%>; visibility:visible;">
	
	
	<%
	
	oimg = "ikon_gantt_48.png"
	oleft = 0
	otop = 0
	owdt = 400
	oskrift = "Timer Import Errors (sen. 14 dage)"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)

   


	'*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	
	

  
    tTop = 20
	tLeft = 0
	tWdth = 620
	
	
	call tableDiv(tTop,tLeft,tWdth)

    %>

    <br />
    <b>Fejlforklaring:</b><br /><br />
    1: Intw. id ikke med i CATI eksport <br />
    11: Intw. ikke fundet i TimeOut <br />
    2: Jobnr ikke med i CATI eksport<br />
    21: Job ikke fundet i TimeOut <br />
    3: Aktivitet ikke fundet på job <br />
    4: Valuta ikke fundet <br />
    5: Kunde ikke fundet på job<br />
    6: CATI externt sysid allerede indlæst<br /><br />
    
            
   <table cellspacing=1 cellpadding=2 border=0 width=100%>
    <tr bgcolor="#5582d2">
        <td class=alt>Id</td>
        <td class=alt>Kørselsdato</td>
        <td class=alt>Externt ID</td>
        <td class=alt>Error ID</td>
        <td class=alt>Job Nr.</td>
        <td class=alt>Intw nr.</td>
        <td class=alt>Timereg. Dato</td>
    </tr>
    
    
    <%
    'dd' = ypar()
    strSQLtimerimperr = "SELECT id, dato, extsysid, errid, jobid, med_init, timeregdato FROM timer_imp_err WHERE errid  <> 6 ORDER BY dato"
    
    oRec.open strSQLtimerimperr, oConn, 3
    x = 0
    while not oRec.EOF 

    select case right(x, 1)
    case 0,2,4,6,8
    bgthis = "#8caae6"
    case else
    bgthis = "#FFFFFF"
    end select
        %>
    <tr bgcolor="<%=bgthis %>">
        
        <td><%=oRec("id") %></td>
        <td><%=oRec("dato") %></td>
        <td><%=oRec("extsysid") %></td>
        <td><%=oRec("errid") %></td>
        <td><%=oRec("jobid") %></td>
        <td><%=oRec("med_init") %></td>
        <td><%=oRec("timeregdato") %></td>
    </tr>


    <%
    x = x + 1
    oRec.movenext
    wend
    oRec.close

    %>
    <tr>
    <td colspan=7>Antal ialt: <%=x %></td>
    
    </tr>
	
    </table>
	
     </div><!-- table div -->
	
	
	<br><br>
	<br>
	
 

	<br>
	<br>
	
	&nbsp;
	
	
	</div>
	


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
