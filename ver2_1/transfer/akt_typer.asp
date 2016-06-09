<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<%



if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	thisfile = "akt_typer"
	
	
	select case func
	case "opdliste"
	
	
	'Response.write "Opdaterer liste"
	
	'Response.Write request("FM_timer")
	'Response.end
	
	
	'Response.Write request("aty_id") & "<br>"
	'Response.Write request("aty_on") & "<br>"
	'Response.Write request("aty_on_realhours") & "<br>"
	'Response.Write request("aty_on_invoice") & "<br>"
	'Response.Write request("aty_on_invoice_chk") & "<br>"
	'Response.Write request("aty_on_workhours") & "<br>"
	'Response.Write request("aty_pre")
	'Response.flush
	
	aty_id = split(request("aty_id"), ",")
	aty_on = split(request("aty_on"), "#,")
	aty_on_realhours = split(request("aty_on_realhours"), "#,")
	aty_on_invoice = split(request("aty_on_invoice"), "#,")
	aty_on_invoice_chk = split(request("aty_on_invoice_chk"), "#,")
	aty_on_workhours = split(request("aty_on_workhours"), "#,")
	aty_pre = split(request("aty_pre"), ", #, ")
	aty_on_invoiceble = split(request("aty_on_invoiceble"), "#,")
	
	for t = 0 to UBOUND(aty_id)
	err = 0
	    
	    'Response.Write "aty_pre(t)"& aty_pre(t) & "<br>"
	    
	    aty_pre(t) = replace(aty_pre(t), " ", "")
	    'aty_pre(t) = replace(aty_pre(t), ".", "")
	    aty_pre(t) = replace(aty_pre(t), ",", ".")
	      
	    if len(trim(aty_on(t))) = 0 then
	    aty_on(t) = 0
	    end if
	    
	    if len(trim(aty_on_realhours(t))) = 0 then
	    aty_on_realhours(t) = 0
	    end if
	    
	    if len(trim(aty_on_invoice(t))) = 0 then
	    aty_on_invoice(t) = 0
	    end if
	    
	    if len(trim(aty_on_invoice_chk(t))) = 0 then
	    aty_on_invoice_chk(t) = 0
	    end if
	    
	    if len(trim(aty_on_workhours(t))) = 0 then
	    aty_on_workhours(t) = 0
	    end if
	    
	    if len(trim(aty_on_invoiceble(t))) = 0 then
	    aty_on_invoiceble(t) = 0
	    end if
	    
	    
	     
	
	      'Response.Write "aty_on_invoice_chk(t)|"& aty_on_invoice_chk(t) & "|<br>"
	      'Response.Write "aty_on_workhours(t)|"& aty_on_workhours(t) & "|<br>"
	      
	      call erDetInt(aty_pre(t))
	    
	      if cint(isInt) > 0 then
	      err = 1
	      end if
	      
	      isInt = 0
	      
	     
	     
	    if cint(err) = 0 then
		
		strSQL = "UPDATE akt_typer SET aty_on = " & left(trim(aty_on(t)), 1) & ", "
		strSQL = strSQL &" aty_on_realhours = "& left(trim(aty_on_realhours(t)), 1) &", "
		strSQL = strSQL &" aty_on_invoiceble = "& left(trim(aty_on_invoiceble(t)), 1) &", "
		strSQL = strSQL &" aty_on_invoice = "& left(trim(aty_on_invoice(t)), 1) &", "
		strSQL = strSQL &" aty_on_invoice_chk = "& left(trim(aty_on_invoice_chk(t)), 1) &", "
		strSQL = strSQL &" aty_on_workhours = "& left(trim(aty_on_workhours(t)), 1) &", "
		strSQL = strSQL &" aty_pre = "& left(trim(aty_pre(t)), 4) &""
		strSQL = strSQL &" WHERE aty_id = "& aty_id(t)
		
		
		'Response.write "t" & t & "<br>"& strSQL & "<br>"
		'Response.flush
		oConn.execute(strSQL)
		
		end if
	
	next
	
	'Response.end
	Response.redirect "akt_typer.asp"
				
	
	
	
	case else%>
	
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<script language="javascript">
	<!--
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	//-->
	</script>
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:20; top:10; visibility:visible;">
	
	<% 
	oimg = "akt_48.png"
	oleft = 0
	otop = 0
	owdt = 300
	oskrift = "Aktivitets typer"
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	
	tTop = 0
	tLeft = 0
	tWdth = 600
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>
	
	
	
	<table cellspacing="0" cellpadding="1" border="0" width="100%">
	<form method=post action="akt_typer.asp?func=opdliste">
	
	<tr bgcolor="#5582D2">
		<td class=alt valign=bottom align=right><b>Id</b></td>
		<td class=alt valign=bottom style="padding-left:10px;"><b>Type</b></td>
		<td class=alt valign=bottom><b>Aktiv</b></td>
		<td class=alt valign=bottom style="width:80px;"><b>Forv.<br /> real. tid</b> <br />(norm. uge)</td>
		<td class=alt valign=bottom style="width:90px;"><b>Fakturerbar</b></td>
		<td class=alt valign=bottom><b>Faktura</b> <br /> (med på)</td>
		<td class=alt valign=bottom><b>Faktura</b> <br /> (checked)</td>
		<td class=alt valign=bottom><b>Løn timer</b> <br /> (fradrag)</td>
		<td class=alt valign=bottom><b>Pre. udfyldt</b><br />(60 min. = 1)</td>
		<td class=alt valign=bottom><b>Enh.</b></td>
		<td class=alt valign=bottom><b>Drys timer</b></td>
	</tr>
	<%
	
	
	strSQL = "SELECT aty_id, aty_label, aty_desc, "_
    & "aty_on, aty_on_realhours, aty_on_invoiceble, aty_on_invoice, aty_on_invoice_chk, "_
    & "aty_on_workhours, aty_pre, aty_sort, aty_on_recon, aty_enh, aty_on_adhoc FROM akt_typer ORDER BY aty_sort"
	
	x = 0
	
	oRec.open strSQL, oConn, 3
	while not oRec.EOF 
	
	select case right(x, 1)
	case 0,2,4,6,8
	bgcol = "#ffffff"
	case else
	bgcol = "#EFF3FF"
	end select
	
	%>
	<tr>
		<td bgcolor="#cccccc" colspan="11" style="padding:0px; height:1px;"><img src="ill/blank.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr bgcolor="<%=bgcol %>">
        <input id="aty_id" name="aty_id" type="hidden" value="<%=oRec("aty_id")%>" />
		<td align=right><%=oRec("aty_id")%></td>
		<td style="padding-left:10px;"><b><%=oRec("aty_desc")%></b></td>
		<td>
		    <%if oRec("aty_on") = 1 then
		    aty_on_CHK = "CHECKED"
		    else
		    aty_on_CHK = ""
		    end if  %>
		
            <input id="aty_on" name="aty_on" type="checkbox" value="1" <%=aty_on_CHK %> />
            <input id="aty_on" name="aty_on" type="hidden" value="#" /></td>
         	<td>
         	<%if oRec("aty_on_realhours") = 1 then
         	aty_on_realhours_CHK = "CHECKED"
         	else
         	aty_on_realhours_CHK = ""
         	end if %>
         	
            <input id="aty_on_realhours" name="aty_on_realhours" type="checkbox" <%=aty_on_realhours_CHK %> value="1" />
            <input id="aty_on_realhours" name="aty_on_realhours" type="hidden" value="#" />
            </td>
            <td>
         	<%select case oRec("aty_on_invoiceble") 
         	case 1 
         	aty_on_invoiceble_SEL1 = "SELECTED"
         	aty_on_invoiceble_SEL2 = ""
         	aty_on_invoiceble_SEL0 = ""
         	case 2 
         	aty_on_invoiceble_SEL1 = ""
         	aty_on_invoiceble_SEL2 = "SELECTED"
         	aty_on_invoiceble_SEL0 = ""
         	case else
         	aty_on_invoiceble_SEL1 = ""
         	aty_on_invoiceble_SEL2 = ""
         	aty_on_invoiceble_SEL0 = "SELECTED"
            end select %>
                <select id="aty_on_invoiceble" name="aty_on_invoiceble" style="width:80px; font-size:9px; font-family:arial;">
                    <option value=1 <%=aty_on_invoiceble_SEL1 %>>Fakturerbar</option>
                    <option value=2 <%=aty_on_invoiceble_SEL2 %>>Ikke Fak.bar</option>
                    <option value=0 <%=aty_on_invoiceble_SEL0 %>>Andet</option>
                </select>
            <input id="aty_on_invoiceble" name="aty_on_invoiceble" type="hidden" value="#" />
            </td>
           	<td>
           	
           	<%if oRec("aty_on_invoice") = 1 then
         	aty_on_invoice_CHK = "CHECKED"
         	else
         	aty_on_invoice_CHK = ""
         	end if %>
           	
            <input id="aty_on_invoice" name="aty_on_invoice" type="checkbox" <%=aty_on_invoice_CHK %> value="1" />
            <input id="aty_on_invoice" name="aty_on_invoice" type="hidden" value="#" />
            </td>
         	<td>
            <%if oRec("aty_on_invoice_chk") = 1 then
         	aty_on_invoice_chk_CHK = "CHECKED"
         	else
         	aty_on_invoice_chk_CHK = ""
         	end if %>
           	
            <input id="aty_on_invoice_chk" name="aty_on_invoice_chk" type="checkbox" <%=aty_on_invoice_chk_CHK %> value="1" />
            <input id="aty_on_invoice_chk" name="aty_on_invoice_chk" type="hidden" value="#" /></td>
         
            	<td>
            <%if oRec("aty_on_workhours") = 1 then
         	aty_on_workhours_CHK = "CHECKED"
         	else
         	aty_on_workhours_CHK = ""
         	end if %>
           	
            <input id="aty_on_workhours" name="aty_on_workhours" type="checkbox" <%=aty_on_workhours_CHK %> value="1" />
            <input id="aty_on_workhours" name="aty_on_workhours" type="hidden" value="#" />
            </td>
         	
            	<td>
                    <input id="aty_pre" name="aty_pre" style="font-size:9px; width:50px;" id="Text1" value="<%=formatnumber(oRec("aty_pre"), 2) %>" type="text" />
                    <input id="aty_pre" name="aty_pre" type="hidden" value="#" />
           </td>
           <td>
            <%select case oRec("aty_enh")
            case -1
            enh = ""
	        case 1
	        enh = "stk."
	        case 2
	        enh = "enh."
	        case 3
            enh = "km."
	        case else
	        enh = "tim."
	        end select%>
	        <%=enh %>
           </td>
           <td>
           <%if oRec("aty_on_adhoc") = 1 then %>
           ja
           <%end if %>
           </td>
		
	</tr>
	<%
	x = x + 1
	oRec.movenext
	wend
	%>	
	</table>
	
	<input id="Hidden1" name="aty_on" type="hidden" value="#" />
	<input id="Hidden2" name="aty_on_realhours" type="hidden" value="#" />
	<input id="Hidden3" name="aty_on_invoiceble" type="hidden" value="#" />
	<input id="Hidden4" name="aty_on_invoice" type="hidden" value="#" />
	<input id="Hidden5" name="aty_on_invoice_chk" type="hidden" value="#" />
	<input id="Hidden6" name="aty_on_workhours" type="hidden" value="#" />
	<input id="Hidden7" name="aty_pre" type="hidden" value="#" />
	
	
	
	<br><br>&nbsp;</div>
	<table width=100%>
	<tr>
	    <td colspan=8 align=right style="padding-right:50px;"><br />
            <input id="Submit2" type="submit" value=" Opdater >> " /></td>
	</tr>
	
	</form>
	</table>
	
	<br><br><br>
	<a href="Javascript:window.close()" class=red>[Luk vindue]</a>
	
	
	<%
	
	itop = 50
	ileft = 0
	iwdt = 400
	
	call sideinfo(itop,ileft,iwdt) %>
	<b>Aktiv</b><br />
		Er denne aktivitetstype aktiv.<br /><br />
	
	<b>Real. timer</b> <br />
		Skal denne type medregnes i dagens daglige timeforbrug. 
		Typen udgør en del af den ugentlige normerede arbejdstid for en medarbejder.<br /><br />
		<b>Fakturerbar</b><br />
		Skal denne type medregnes til fakturerbare timer.<br /><br />
		<b>Faktura</b> <br /> 
		Skal denne type med på fakturaer.<br /><br />
		<b>Faktura</b> <br />Skal denne type være forvalgt på faktura oprettelse.<br /><br />
		<b>Løn timer</b> <br /> 
		Skal denne type give fradrag i forhold til løn timer (Stempelur timer).<br /><br />
		<b>Pre. udfyldt</b><br />
		Skal denne type være forudfyldt på timereg. siden. (60 min. = 1).<br /><br />
	   <b>Drys timer</b><br />
        Denne type bruges til at "drysse" f.eks en time ud på forskellige job. TimeOut beregner selv hvor mange job der er på den
        "personlige aktiv liste" og drysser herefter den angivne tid, ud på de aktiviteter der findes af typen "drys".<br /><br />
	
	</td>
	</tr>
	</table>
	</div>
	<br />
	<br /><br /><br /><br />&nbsp;
	<%end select%>


<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
