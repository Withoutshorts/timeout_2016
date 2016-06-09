<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/xml/tsa_xml_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="inc/helligdage_func.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->

<script>

    function checkAll(val) {


        //FM_akttype_id_1
    //alert(val)
    antal_v = document.getElementById("antal_v_"+val).value
    //alert(antal_v)
    
        if (document.getElementById("chkalle_"+val).checked == true) {
        tval = true
        } else {
        tval = false
    }
        
        //alert(tval)

        for (i = 0; i < antal_v; i++)
        //alert(i)
        document.getElementById("FM_akttype_id_"+val+"_"+i).checked = tval
    }


</script>

<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	thisfile = "bal_real_norm_2007.asp"
	
	select case func 
	case "-"
	
	case else
	
	'*** Sætter lokal dato/kr format. *****
	Session.LCID = 1030
	
	leftPos = 20
	topPos = 132
	

	if len(request("FM_md")) <> 0 then
	strMd = request("FM_md")
	strAar = request("FM_aar")
	intMid = request("FM_medarb")
	else
	intMid = session("mid")
	strMd = month(date)
	strAar = year(date)
	end if
	
	
	
	if len(request("FM_md_slut")) <> 0 then
	strMd_slut = request("FM_md_slut")
	strAar_slut = request("FM_aar_slut")
    else
	strMd_slut = datepart("m",(dateadd("m", 1, date)))
	    
	    if strMd = 12 then
	    strMd_slut = 1
	    strAar_slut = datepart("yyyy",(dateadd("yyyy", 1, date)))
	    else 
	    strAar_slut = datepart("yyyy",(date))
	    end if
	    
	end if
	
	
	select case strMd_slut
	case 1,3,5,7,8,10,12
	usedag = 31
	case 2
	    select case useMD_slut
	    case 2004, 2008, 2012, 2016, 2020, 2024, 2028, 2032, 2036, 2040, 2044
	    usedag = 29
	    case else 
	    usedag = 28
	    end select
	case else
	usedag = 30
	end select
	
	
	if len(request("FM_akttype")) <> 0 AND len(request("FM_md_slut")) <> 0 then 
	akttype_sel = request("FM_akttype")
	else
	    if request.Cookies("tsa")("fm_akttype_sel") <> "" then
	    akttype_sel = request.Cookies("tsa")("fm_akttype_sel")
	    else
	    call akttyper2009(6)
	    akttype_sel = "#-1#, #-2#, #-3#, #-4#, #-5#, "
	    akttype_sel = aktiveTyper
	    end if
	end if
	
	
	'Response.Write akttype_sel
	response.Cookies("tsa")("fm_akttype_sel") = akttype_sel
	
	
	
	startdato = strAar&"/"&strMd&"/1"
	slutdato = strAar_slut&"/"&strMd_slut&"/"&usedag
	
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<div id="topmenu" style="position:absolute; left:0; top:42; visibility:visible;">
	<!--<h4>Timeregistrering - Jobliste</h4>-->
	<%call tsamainmenu(7)%>
	</div>
	<div id="sekmenu" style="position:absolute; left:15; top:82; visibility:visible;">
	<%
	if showonejob <> 1 then
		call stattopmenu()
	end if
	%>
	</div>
	
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>; top:<%=topPos%>; visibility:visible;">
	<h3>Medarbejder afstemning</h3>
	
	<%call filterheader(0,0,1004,pTxt)%>
	<table border=0 cellspacing=2 cellpadding=2 width=600>
	<form action="bal_real_norm_2007.asp" method="post" name="periode" id="periode">
	<tr>
	<td valign=top><b>Medarbejder:</b><br /> 
	<select name="FM_medarb" id="FM_medarb" style="width:300px;" onchange="submit();">
	<option value="0">Alle</option>
	<%
	strSQL = "SELECT mnavn, mid, init, mnr FROM medarbejdere WHERE mansat <> '2' ANd mansat <> '3' ORDER BY mnavn"
	oRec.open strSQL, oConn, 3 
	x = 0
	while not oRec.EOF 
	if len(request("FM_medarb")) = 0 AND x = 0 then
		selthis = "SELECTED"
		intMid = session("mid") 'oRec("mid")
	else	
		if cint(intMid) = oRec("mid") then
		selthis = "SELECTED"
		else
		selthis = ""
		end if
	end if
	%>
	<option value="<%=oRec("mid")%>" <%=selthis%>><%=oRec("mnavn")%> (<%=oRec("mnr") %>) <%=oRec("init") %></option>	
	<%
	x = x + 1
	oRec.movenext
	wend
	oRec.close 
	
	%></select>
	</td>
	
	<td valign=top><b>Fra:</b><br />
	<select name="FM_md" id="FM_md">
	<%
	x = 1
	for x = 1 to 12
	select case x
	case 1
	strMDname = "1. Januar"
	case 2
	strMDname = "1. Februar"
	case 3
	strMDname = "1. Marts"
	case 4
	strMDname = "1. April"
	case 5
	strMDname = "1. Maj"
	case 6
	strMDname = "1. Juni"
	case 7
	strMDname = "1. Juli"
	case 8
	strMDname = "1. August"
	case 9
	strMDname = "1. September"
	case 10
	strMDname = "1. Oktober"
	case 11
	strMDname = "1. November"
	case 12
	strMDname = "1. December"
	end select
	
	if cint(strMd) = x then
	selthis = "SELECTED"
	showstrMDname = strMDname
	else
	selthis = ""
	end if
	%>
	<option value="<%=x%>" <%=selthis%>><%=strMDname%></option>
	<%next%>
	</select></td><td valign=top>
	<br />
	<select name="FM_aar" id="FM_aar">
	<%
	for x = 0 to 20
	
	strAarnameid = (2001 + x)
	
	if cint(strAar) = strAarnameid then
	selthis = "SELECTED"
	else
	selthis = ""
	end if
	%>
	<option value="<%=strAarnameid%>" <%=selthis%>><%=strAarnameid%></option>
	<%next%>
	</select></td>
	<td valign=top><b>Til: </b><br />
	<select name="FM_md_slut" id="FM_md_slut">
	<%
	x = 1
	for x = 1 to 12
	select case x
	case 1
	strMDname = "31. Januar"
	case 2
	strMDname = "(28). Februar"
	case 3
	strMDname = "31. Marts"
	case 4
	strMDname = "30. April"
	case 5
	strMDname = "31. Maj"
	case 6
	strMDname = "30. Juni"
	case 7
	strMDname = "31. Juli"
	case 8
	strMDname = "31. August"
	case 9
	strMDname = "30. September"
	case 10
	strMDname = "31. Oktober"
	case 11
	strMDname = "30. November"
	case 12
	strMDname = "31. December"
	end select
	
	if cint(strMd_slut) = x then
	selthis = "SELECTED"
	showstrMDname = strMDname
	else
	selthis = ""
	end if
	%>
	<option value="<%=x%>" <%=selthis%>><%=strMDname%></option>
	<%next%>
	</select></td><td colspan=2 valign=top>
	<br />
	<select name="FM_aar_slut" id="FM_aar_slut">
	<%
	for x = 0 to 20
	
	strAarnameid = (2001 + x)
	
    if cint(strAar_slut) = strAarnameid then
	selthis = "SELECTED"
	else
	selthis = ""
	end if
	%>
	<option value="<%=strAarnameid%>" <%=selthis%>><%=strAarnameid%></option>
	<%next%>
	</select></td>
	</tr>
	</table>
	
	<table border=0 cellspacing=2 cellpadding=2 width=100%>
	<tr>
	<td>
	
	<br /><b>Vis kolonner for følgende aktivitets typer:</b> (Vælg)<br />
			<table cellspacing=2 cellpadding=1 border=0 width=100%>
			<tr><td valign=top>
			<%
			call akttyper2009(5)
			%>
			</table>
			</td></tr>
			</table>
	
	
	
	</td>
	    
	    
	    <td valign=bottom style="padding-top:16px;">
            <input id="Submit1" type="submit" value=" Søg >> " />
	</td></tr>
	</form></table>
    
    <!-- filter header sLut -->
	</td></tr></table>
	</div>
    
	
	<a href="feriekalender.asp?menu=job" class=vmenu>Ferie & Sygdom..</a>
	<br><br><br>
	
	
	<%
	
	tTop = 0
	tLeft = 0
	tWdth = 1200
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	
	'** Timer Realiseret ***'
	call medarbafstem(intMid, startDato, slutDato, 1, akttype_sel)
	%>
    <!-- table div --> 
	</div>
	
	<br /><br />
	  <% 
       itop = 40
       ileft = 0 
       iwdt = 400
       call sideinfo(itop,ileft,iwdt)%>		
       
   
			<b>Aktivitets typer:</b><br />
			<table cellspacing=2 cellpadding=1 border=0 width=100%>
			<tr>
			    <td><b>Id</b></td>
			    <td><b>Navn</b></td>
			    <td><b>Fakturerbar</b></td>
			    <td><b>Med i dagligt timeregnskab</b><br />
			    (Norm. uge)</td>
			</tr>
			<%
			call akttyper2009(3)
			%>
			</table>
		
	    
	    <br /><br />
	    <b>Ferie</b><br />
	    Fra den 1. januar 2001 optjenes 2,08 dages ferie for hver måneds ansættelse i optjeningsåret, som er lig kalenderåret. 
        Dette gælder også medarbejdere på del-tid.<br />
        2,08 * 12 måneder = 25 dage ell. 5 arbejdsuger.
	    <br /><br />
	    <b>Dage</b><br />
	    Ferie of ferie fridage (afspad.) timer angives som timer og omregnes til dage udfra normeret timer pr. uge 
	    (total timer pr. uge / med 5 arb. dage, se medarbejdertyper ell. medarb. afstemning).<br /><br />
	    Hvis man er ansat 37 timer, 
	    angiver man altså 7,4 timer for en hel dags ferie, og 37 timer for en uge. 
	    
	    <br />
	    
                &nbsp;
	   
	    <!-- side info slut -->
        </td></tr></table>
			</div>
	    
	
	<br /><br /><br /><br /><br /><br /><br /><br />
        &nbsp;
			
			</div>
	
	<% 
	end select
	
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
