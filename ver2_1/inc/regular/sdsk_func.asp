<%
function sdsktopmenu()
%>

<div style="margin-top:-15px; margin-bottom:15px;">
    <%if level = 1 then%>
    <a class="btn btn-sm btn-default" href="javascript:NewWin_help('sdsk_prioitet.asp?menu=tok');" role="button" target="_self">Incident Typer</a>
    <a class="btn btn-sm btn-default" href="javascript:NewWin_help('sdsk_prio_typ.asp?menu=tok');" role="button" target="_self">Prioiteter</a>
    <a class="btn btn-sm btn-default" href="javascript:NewWin_help('sdsk_prio_grp.asp?menu=tok');" role="button" target="_self">Aftalegrupper</a>
    <a class="btn btn-sm btn-default" href="javascript:NewWin_help('sdsk_status.asp?menu=tok');" role="button" target="_self">Incident Status</a>
    <%end if
    if level <= 2 OR level = 6 then %>
    <a class="btn btn-sm btn-default" href="javascript:NewWin_help('sdsk_typer.asp?menu=tok');" role="button" target="_self">Incident Kategorier</a>
    <%end if %>
</div>

<%
end function












%>	
