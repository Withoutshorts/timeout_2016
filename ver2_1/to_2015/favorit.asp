

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../timereg/inc/convertDate.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->


<style>

    .picmodal:hover,
    .picmodal:focus {
    text-decoration: none;
    cursor: pointer;
}
   
</style>


<div class="wrapper">
<div class="content">

<%



    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if


    func = request("func")


	call menu_2014 


    stDato = request("stDato")

    redim tjekdag(7)
    tjekdag(1) = stDato
            
    for x = 2 to 7
    tjekdag(x) = dateAdd("d", x-1, stDato)
    next



    medid = request("FM_medid")


    varTjDatoUS_man = request("varTjDatoUS_man")
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)

    datoMan = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
    datoSon = day(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& year(varTjDatoUS_son)

    next_varTjDatoUS_man = dateadd("d", 7, varTjDatoUS_man)
	next_varTjDatoUS_son = dateadd("d", 7, varTjDatoUS_son)
    next_varTjDatoUS_man = year(next_varTjDatoUS_man) &"/"& month(next_varTjDatoUS_man) &"/"& day(next_varTjDatoUS_man)
	next_varTjDatoUS_son = year(next_varTjDatoUS_son) &"/"& month(next_varTjDatoUS_son) &"/"& day(next_varTjDatoUS_son)


    prev_varTjDatoUS_man = dateadd("d", -7, varTjDatoUS_man)
	prev_varTjDatoUS_son = dateadd("d", -7, varTjDatoUS_son)
    prev_varTjDatoUS_man = year(prev_varTjDatoUS_man) &"/"& month(prev_varTjDatoUS_man) &"/"& day(prev_varTjDatoUS_man)
	prev_varTjDatoUS_son = year(prev_varTjDatoUS_son) &"/"& month(prev_varTjDatoUS_son) &"/"& day(prev_varTjDatoUS_son)




    select case func 

    case "fjernfavorit"


        id = request("id")

        %>
            <div class="container">
                <div class="portlet">
                    <h3 class="portlet-title"><u>Favorit liste</u></h3>
                    <div class="portlet-body">
                        <%response.Write "aktivi id" & id  %>
                    </div>

                </div>
            </div>
        <%
        

        oConn.execute("UPDATE timereg_usejob SET favorit = 0 WHERE aktid = "&id&"")

        response.Redirect "favorit.asp"

    case else

%>




            <div class="container">
                <div class="portlet">
                    <h3 class="portlet-title"><u>Favorit liste</u></h3>
                    <div class="portlet-body">

                        <form action="favorit.asp?sogsubmitted=1" method="post">
                        <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
                        <div class="row">
                            <div class="col-lg-3">
                                 
                                <%
                                    strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe, init FROM medarbejdere WHERE mansat <> 2 GROUP BY mid ORDER BY Mnavn" 
                                %>
                                <select name="FM_medid" id="FM_medid" <%=progrpmedarbDisabled  %> class="form-control input-small"  onchange="submit();">
                                    <%

                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF
                                
                                    
                                    StrMnavn = oRec("Mnavn")
                                    StrMinit = oRec("init")
	
                                    if cdbl(medid) = cdbl(oRec("Mid")) then
				                    isSelected = "SELECTED"
				                    else
				                    isSelected = ""
				                    end if

				                    %>
                                     <option value="<%=oRec("Mid")%>" <%=isSelected%>><%=StrMnavn &" "& StrMinit%></option>
                                    <%
                                    oRec.movenext
                                    wend
                                    oRec.close  
                                    %>
                                </select>
                            </div>
                        </div>
                        </form>
                        <%response.Write "id: " & medid%>


                        <table class="table dataTable table-striped table-bordered ui-datatable">

                            <thead>
                                <tr>
                                    <th>Aktivitet</th>

                                    <%
                                            perInterval = 6 'dateDiff("d", varTjDatoUS_man, varTjDatoUS_son, 2,2) 
                                            perIntervalLoop = perInterval

                                            for l = 0 to perIntervalLoop 
        
                                            if l = 0 then
                                            varTjDatoUS_use = varTjDatoUS_man
                                            else
                                            varTjDatoUS_use = dateAdd("d", l, varTjDatoUS_man)
                                            end if 

                                            showdate = DatePart("yyyy",varTjDatoUS_use) & "-" & Right("0" & DatePart("m",varTjDatoUS_use), 2) & "-" & Right("0" & DatePart("d",varTjDatoUS_use), 2)



                                            'response.Write weekdayname(weekday(varTjDatoUS_use, 1))

                                            %>
                                                <th style="width:75px"><%=weekdayname(weekday(varTjDatoUS_use, 1)) %> <br />
                                                    <%=showdate %>
                                                </th>
                                            <%

                                            next
                                    %>
                                    <th>Total</th>
                                    <th>Fjern favorit</th>
                                </tr>
                            </thead>

                            <tbody>
                                <%
                                     favoriter = 0
                                     StrSqlfav = "SELECT medarb, jobid, aktid, forvalgt_af FROM timereg_usejob WHERE medarb = "& medid & " AND favorit <> 0" 
                                     
                                     oRec.open StrSqlfav, oConn, 3
                                     while not oRec.EOF

                                     jobid = oRec("jobid")
                                     aktid = oRec("aktid")
                                     medarb = oRec("medarb")

                                     favoriter = favoriter + 1
                                     'response.Write "favs: " & favoriter
                                     'response.Write "<br>" & jobid
                                     for i = 1 to favoriter
                                        
                                         i = i + 1

                                        StrSQLjob = "SELECT id, jobnavn FROM job WHERE id ="& jobid

                                        oRec3.open StrSQLjob, oConn, 3
                                        if not oRec3.EOF then

                                        jobnavn = oRec3("jobnavn")

                                        StrSQLakt = "SELECT id, navn, beskrivelse, budgettimer FROM aktiviteter WHERE id ="& aktid

                                        oRec2.open StrSqlakt, oConn, 3
                                        if not oRec2.EOF then

                                        aktNavn = oRec2("navn")
                                        TaktId = oRec2("id")
                                        aktbudgettimer = oRec2("budgettimer")
                                        'response.Write TaktId
                                        %>
                                        <tr>
                                            <td><span style="font-size:75%"><%=jobnavn %></span><br /><%=aktNavn %>

                                                <span id="modal_<%=TaktId %>" style="color:cornflowerblue;" class="fa fa-book pull-right picmodal"></span>                                               
                                                <div id="myModal_<%=TaktId %>" style="display:none">
                                                
                                                    <%
                                                        StrSqltimerialt = "SELECT sum(Timer) as timer FROM timer WHERE TAktivitetId ="& TaktId
                                                        oRec6.open StrSqltimerialt, oConn, 3
                                                        if not oRec6.EOF then
                                                        timerforalle = oRec6("timer")
                                                        end if
                                                        oRec6.close

                                                        StrSqltotaltimer = "SELECT TAktivitetId, sum(timer) as timer FROM timer WHERE TAktivitetId ="& TaktId &" AND tmnr ="&medid
                                                        oRec5.open StrSqltotaltimer, oConn, 3
                                                        if not oRec5.EOF then
                                                
                                                        timertotal = oRec5("timer")
                                                        
                                                        if timerforalle > aktbudgettimer then
                                                           txtcolor = "red"
                                                        else
                                                            txtcolor = "green"
                                                        end if

                                                    %>
                                                        <span style="font-size:75%; color:#5582d2">Forkalk.: <%=aktbudgettimer %></span>
                                                        <span style="font-size:75%; color:<%=txtcolor%>;">Real:<%=timerforalle %></span>
                                                        <span style="font-size:75%; color:#5582d2;">Egne: <%=timertotal %></span>
                                                    <%
                                                        end if
                                                        oRec5.close 
                                                    %>
                                                
                                                </div>
                                            </td>

                                            <%
                                                
                                                for l = 0 to 6
                                                    

                                                    if l = 0 then
                                                        timerdato = varTjDatoUS_man
                                                    else
                                                        timerdato = dateAdd("d",1,timerdato)
                                                    end if

                                                    timerdato = year(timerdato) & "-" & month(timerdato) & "-" & day(timerdato) 

                                                    StrSQLtimer = "SELECT TAktivitetId, Timer FROM timer WHERE TAktivitetId ="& TaktId & " AND tdato = "& "'" & timerdato & "'"
                                                
                                                     oRec4.open StrSQLtimer, oConn, 3
                                                     if not oRec4.EOF then
                                                
                                                     timerdag = oRec4("Timer")
                                                    
                                                    %>
                                                        <td><input type="text" style="width:75px;" class="form-control input-small" name="FM_timerdag" value="<%=timerdag %>" /></td>
                                                    <%


                                                    else
                                                    timerdag = 0
                                                    %>
                                                        <td><input type="text" style="width:75px;" class="form-control input-small" name="FM_timerdag" value="<%=timerdag %>" /></td>
                                                    <%

                                                     end if 
                                                     oRec4.close

                                                next

                                                 'StrSQLtimer = "SELECT TAktivitetId, Timer FROM timer WHERE TAktivitetId ="& aktId
                                                
                                                 'oRec4.open StrSQLtimer, oConn, 3
                                                 'if not oRec4.EOF then
                                                
                                                    'timerdag = oRec4("Timer")

                                                 'end if 
                                                 'oRec4.close
                                                 
                                            %>

                                            <td>
                                                <%
                                                    ugestart_dato = year(datoMan) & "-" & month(datoMan) & "-" & day(datoMan)
                                                    ugeslut_dato = year(datoSon) & "-" & month(datoSon) & "-" & day(datoSon)

                                                    'response.Write ugestart_dato & "SØNDAG: " & ugeslut_dato
                                                    
                                                    StrSqlweektotal = "SELECT sum(timer) as timer FROM timer WHERE TAktivitetId ="& TaktId & " AND tdato BETWEEN '"& ugestart_dato &"' AND '"& ugeslut_dato &"'" 
                                                
                                                    oRec7.open StrSqlweektotal, oConn, 3
                                                    if not oRec7.EOF then
                                                    timerweektotal = oRec7("timer")
                                                    
                                                    if timerweektotal <> 0 then
                                                    else
                                                    timerweektotal = 0
                                                    end if

                                                    end if
                                                    oRec7.close    
                                                    
                                                %>
                                                <%=timerweektotal %>
                                            </td>

                                            <td>
                                                <a href="favorit.asp?id=<%=oRec2("id") %>&func=fjernfavorit"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a>
                                            </td>

                                        </tr>
                                        <%
                                        end if
                                        oRec2.close

                                        end if
                                        oRec3.close

                                    next

                                    oRec.movenext
                                    wend
                                    oRec.close 
                                    %>
                            </tbody>

                        </table>

                        <%

                        %>

                    </div>
                </div>
            </div>

            



    <%end select %>



        </div>
    </div>


<script type="text/javascript">


$(".picmodal").click(function() {

    var modalid = this.id
    var idlngt = modalid.length
    var idtrim = modalid.slice(6, idlngt)

    //var modalidtxt = $("#myModal_" + idtrim);
    var modal = document.getElementById('myModal_' + idtrim);

    
    if (modal.style.display !== 'none')
    {
        modal.style.display = 'none';
    }
    else
    {
        modal.style.display = 'block';
    }
   

});


</script>


<!--#include file="../inc/regular/footer_inc.asp"-->