

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<div class="wrapper">
<div class="content">


<%

 if len(session("user")) = 0 then
	errortype = 5
	call showError(errortype)
    response.End
 end if


 sqlDTD = year(now) &"-"& month(now) &"-"& day(now)


    if len(trim(request("FM_datoStart"))) <> 0 then
    nyhed_startDato = request("FM_datoStart")
    else
    nyhed_startDato = day(now) &"-"& month(now) &"-"& year(now)
    end if

    if len(trim(request("FM_datoStart"))) <> 0 then
    nyhed_slutDato = request("FM_datoSlut")
    else
    nyhed_slutDato = DateAdd("d", 1, nyhed_startDato) 
    end if


    func = request("func")
    select case func

        case "dbopr"
           
            strOverskrift_new = request("strOverskrift_new")

            if strOverskrift_new <> "" then

                'response.Write "Opretter ny <br>"

                strBrodtekst_new = request("strBrodtekst_new")

                sqlFradato = year(nyhed_startDato) &"-"& month(nyhed_startDato) &"-"& day(nyhed_startDato)
                sqlTildato = year(nyhed_slutDato) &"-"& month(nyhed_slutDato) &"-"& day(nyhed_slutDato)
          
                'response.Write " Overskrift " & strOverskrift_new & " text " & strBrodtekst_new & " Start Dato - " & sqlFradato &" Slut Dato - "& sqlTildato 
    
                if cdate(sqlFradato) > cdate(sqlTildato) then           
                    errortype = 193
	                call showError(errortype)
                       response.End
                end if

                oConn.execute("INSERT INTO info_screen (overskrift, brodtext, datofra, datotil, editor) VALUES ('"& strOverskrift_new &"', '"& strBrodtekst_new &"', '"& sqlFradato &"', '"& sqlTildato &"', '"& session("user") &"')")

            end if


           ' response.Write "<br><br> Rediger de nyværende"

            nyhed_id = request("nyhed_id")
            'response.Write "id id id id " & nyhed_id
            nyhed_id_split = split(nyhed_id, ", ")

            for t = 0 TO UBOUND(nyhed_id_split)
                
                overskrift = request("overskrift_" & nyhed_id_split(t))
                brodtekst = request("brodtekst_" & nyhed_id_split(t))
                fradato = request("fradato_" & nyhed_id_split(t))
                tildato = request("tildato_" & nyhed_id_split(t))

                sqlFradato = year(fradato) &"-"& month(fradato) &"-"& day(fradato)
                sqlTildato = year(tildato) &"-"& month(tildato) &"-"& day(tildato)
    
                if cdate(sqlFradato) > cdate(sqlTildato) then           
                    errortype = 193
	                call showError(errortype)
                       response.End
                end if

                'response.Write "<br> Nyhed nr. " & nyhed_id_split(t) &" Overskrift - "& overskrift &" brodtekst - "& brodtekst &"Fra og til - "& fradato &" "& tildato 

                
                oConn.execute("UPDATE info_screen SET overskrift ='"& overskrift &"', brodtext = '" &brodtekst &"', datofra = '" & sqlFradato &"', datotil = '" & sqlTildato &"', editor = '"& session("user") &"' WHERE id = "& nyhed_id_split(t) &"")


            next

            
            response.Redirect("infoscreen_news.asp")


    end select

    
%>

    <%call menu_2014 %>


    <script src="js/infoscreen_news.js" type="text/javascript"></script> 

    
    <div class="container">
        <div class="portlet">
            <h3 class="portlet-title"><u>Nyheder</u></h3>
            <div class="portlet-body">
                
                <form action="infoscreen_news.asp?func=dbopr" method="post">
                    <div class="row">
                <div class="col-lg-10">&nbsp</div>
                <div class="col-lg-2 pad-b10">
                    <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
                </div>
            </div>
                <table class="table table-striped table-bordered">
                    <thead>
                        <tr>
                            <th>Overskrift</th>
                            <th style="width:50%">Brødtekst</th>
                            <th style="width:10%">Vis fra</th>
                            <th style="width:10%">Vis til</th>
                        </tr>
                    </thead>

                    <tbody>
                        <tr>
                            <td><input type="text" name="strOverskrift_new" class="form-control input-small" /></td>
                            <td><input type="text" name="strBrodtekst_new" class="form-control input-small" /></td>
                            <td>
                                <div class='input-group date'>
                                      <input type="text" style="width:100px;" class="form-control input-small" name="FM_datoStart" id="jq_dato" value="<%=nyhed_startDato %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            </td>
                            <td>
                                <div class='input-group date'>
                                      <input type="text" style="width:100px;" class="form-control input-small" name="FM_datoSlut" id="jq_dato" value="<%=nyhed_slutDato %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            </td>
                        </tr>
                    </tbody>

                    <tbody>

                        <%
                            strSQL = "SELECT id, overskrift, brodtext, datofra, datotil, editor FROM info_screen ORDER BY id DESC"
                            'response.Write strSQL
                            oRec.open strSQL, oConn, 3
                            x = 0
                            while not oRec.EOF
                            
                            x = x + 1

                            strEditor = oRec("editor")

                            strOverskrift = oRec("overskrift")
                            strBrodtekst = oRec("brodtext")
                            datoFra = oRec("datofra")
                            datoTil = oRec("datotil")
                            
                        %>
                            <tr>
                                <td>
                                    <input type="hidden" name="nyhed_id" value="<%=oRec("id") %>" />
                                    <input type="text" class="form-control input-small" name="overskrift_<%=oRec("id") %>" value="<%=strOverskrift %>" /></td>
                                <td><input type="text" class="form-control input-small" name="brodtekst_<%=oRec("id") %>" value="<%=strBrodtekst %>" /></td>
                                <td>
                                    <div class='input-group date'>
                                          <input type="text" style="width:100px;" class="form-control input-small" name="fradato_<%=oRec("id") %>" id="jq_dato" value="<%=datoFra %>" placeholder="dd-mm-yyyy" />
                                          <span class="input-group-addon input-small"><span class="fa fa-calendar"></span></span>                                          
                                   </div>
                               </td>

                                <td>
                                    <div class='input-group date'>
                                          <input type="text" style="width:100px;" class="form-control input-small" name="tildato_<%=oRec("id") %>" id="jq_dato" value="<%=datoTil %>" placeholder="dd-mm-yyyy" />
                                          <span class="input-group-addon input-small"><span class="fa fa-calendar"></span></span>                                          
                                   </div>
                               </td>
                            </tr>
                        <%

                            oRec.movenext
                            wend
                            oRec.close
                        %>

                    </tbody>

                </table>
                </form>

                <%if x <> 0 then %>

                <br />

                <div class="row">
                    <div class="col-lg-4">Sidst opdateret af <b><%=strEditor %> </b></div>
                </div>
                <%end if %>

            </div>
        </div>
    </div>





</div>
</div>


<!--#include file="../inc/regular/footer_inc.asp"-->