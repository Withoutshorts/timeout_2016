

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<div class="wrapper">
<div class="content">

    <style>
        .clickable:hover,
        .clickable:focus {
            text-decoration: none;
            cursor: pointer;
        }
    </style>

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

        case "slet"
            '*** Her spørges om det er ok at der slettes ***
            id = request("id")
	        oskrift = "Nyheder" 
            slttxta = "Du er ved at <b>SLETTE</b> en nyhed. Er dette korrekt?"
            slttxtb = "" 
            slturl = "infoscreen_news.asp?func=sletok&id="& id

            call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

        case "sletok"

            id = request("id")
            oConn.execute("DELETE FROM info_screen WHERE id = "& id &"")
	        Response.redirect "infoscreen_news.asp"

        case "dbopr"
           
            strOverskrift_new = request("strOverskrift_new")
            strOverskrift_new = replace(strOverskrift_new, "'", "''")

            if strOverskrift_new <> "" then

                'response.Write "Opretter ny <br>"

                strBrodtekst_new = request("strBrodtekst_new")
                strBrodtekst_new = replace(strBrodtekst_new, "'","''")

                sqlFradato = year(nyhed_startDato) &"-"& month(nyhed_startDato) &"-"& day(nyhed_startDato)
                sqlTildato = year(nyhed_slutDato) &"-"& month(nyhed_slutDato) &"-"& day(nyhed_slutDato)
          
                'response.Write " Overskrift " & strOverskrift_new & " text " & strBrodtekst_new & " Start Dato - " & sqlFradato &" Slut Dato - "& sqlTildato 
    
                if cdate(sqlFradato) > cdate(sqlTildato) then           
                    errortype = 193
	                call showError(errortype)
                       response.End
                end if

                nyhedtype = request("FM_type") ' 1 = intern, 2 = ekstern

                if len(trim(request("FM_vigtig"))) <> 0 then
                    vigtig = request("FM_vigtig")
                else
                    vigtig = 0
                end if

                oConn.execute("INSERT INTO info_screen (overskrift, brodtext, datofra, datotil, editor, vigtig, type) VALUES ('"& strOverskrift_new &"', '"& strBrodtekst_new &"', '"& sqlFradato &"', '"& sqlTildato &"', '"& session("user") &"', "& vigtig &", "& nyhedtype &")")

            end if


           ' response.Write "<br><br> Rediger de nyværende"

            nyhed_id = request("nyhed_id")
            'response.Write "id id id id " & nyhed_id
            nyhed_id_split = split(nyhed_id, ", ")

            for t = 0 TO UBOUND(nyhed_id_split)
                
                overskrift = request("overskrift_" & nyhed_id_split(t))
                overskrift = replace(overskrift, "'", "''")
                brodtekst = request("brodtekst_" & nyhed_id_split(t))
                brodtekst = replace(brodtekst, "'", "''")
                fradato = request("fradato_" & nyhed_id_split(t))
                tildato = request("tildato_" & nyhed_id_split(t))

                sqlFradato = year(fradato) &"-"& month(fradato) &"-"& day(fradato)
                sqlTildato = year(tildato) &"-"& month(tildato) &"-"& day(tildato)
    
                if cdate(sqlFradato) > cdate(sqlTildato) then           
                    errortype = 193
	                call showError(errortype)
                       response.End
                end if

                nyhedtype = request("FM_type_" & nyhed_id_split(t))

                if len(trim(request("FM_vigtig_" & nyhed_id_split(t)))) <> 0 then
                    vigtig = request("FM_vigtig_" & nyhed_id_split(t))
                else
                    vigtig = 0
                end if

                'response.Write "<br> Nyhed nr. " & nyhed_id_split(t) &" Overskrift - "& overskrift &" brodtekst - "& brodtekst &"Fra og til - "& fradato &" "& tildato 

                'response.Write "UPDATE info_screen SET overskrift ='"& overskrift &"', brodtext = '" &brodtekst &"', datofra = '" & sqlFradato &"', datotil = '" & sqlTildato &"', editor = '"& session("user") &"', vigtig = "& vigtig &", type = "& nyhedtype &" WHERE id = "& nyhed_id_split(t)
                
                oConn.execute("UPDATE info_screen SET overskrift ='"& overskrift &"', brodtext = '" &brodtekst &"', datofra = '" & sqlFradato &"', datotil = '" & sqlTildato &"', vigtig = "& vigtig &", type = "& nyhedtype &" WHERE id = "& nyhed_id_split(t) &"")


            next

            
            response.Redirect("infoscreen_news.asp")



    case "sletbilledeok"

        newsid = request("newsid")
        strPath = "d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\upload\"&lto&"\" & Request("filnavn")

        on Error resume Next 

	    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	    Set fsoFile = FSO.GetFile(strPath)
	    fsoFile.Delete

        oConn.execute("UPDATE info_screen SET filnavn = NULL WHERE id ="& newsid)

        response.Redirect "infoscreen_news.asp?func=billedeupload&newsid="& newsid

    case "billedeupload"

        newsid = request("newsid")
        %>
        <div class="container">
            <div class="portlet">
                <div class="portlet-body">

                    <%
                    'Tjekker om der er uploadet et billede i forvejen
                    filfindes = 0
                    strSQL = "SELECT filnavn FROM info_screen WHERE id = "& newsid
                    oRec.open strSQL, oConn, 3
                    if not oRec.EOF then
                        if len(oRec("filnavn")) <> 0 then
                        filfindes = 1
                        filnavn = oRec("filnavn")
                        end if
                    end if
                    oRec.close
                    %>

                    <%if cint(filfindes) = 0 then %>
                        <form ENCTYPE="multipart/form-data" action="../timereg/upload_bin.asp?newsuplaod=1&newsid=<%=newsid %>" method="post" id="image_upload">
                            <div class="row" style="text-align:center">
                                <div class="col-md-10">
                                    <div class="fileinput fileinput-new" data-provides="fileinput">
                                    <div class="fileinput-preview thumbnail" data-trigger="fileinput" style="width: 300px; height: 350px;"></div>
                                    <div>
                                        <span class="btn btn-default btn-file"><span class="fileinput-new">Select image</span><span class="fileinput-exists">Change</span><INPUT NAME="fileupload1" TYPE="file"></span>
                                        <a href="#" class="btn btn-default fileinput-exists" data-dismiss="fileinput">Remove</a>
                                    </div>
                                    </div>
                                </div> <!-- /.col -->
                             </div>

                            <div class="row" style="text-align:center">
                                <div class="col-lg-12">
                                    <button type="submit" class="btn btn-secondary"><b>Upload</b></button>
                                </div>
                            </div>
                        </form>
                    <%else %>
                        <div class="row" style="text-align:center">
                            <div class="fileinput fileinput-new" data-provides="fileinput">
                                <div class="fileinput-preview thumbnail" style="width: 300px; height: 350px;">
                                    <img src="../inc/upload/<%=lto%>/<%=filnavn %>" alt='' border='0'>                                                   
                                </div>
                            </div>
                        </div>


                        <div class="row" style="text-align:center">
                            <div class="col-lg-12">
                                <a href="infoscreen_news.asp?func=sletbilledeok&newsid=<%=newsid %>&filnavn=<%=filnavn %>" class="btn btn-default"><b>Remove</b></a>
                            </div>
                        </div>
                    <%end if %>

                </div>
            </div>
        </div>
        <%



    case else 'rediger opret nyheder del'
%>

    <%call menu_2014 %>


   <!-- <script src="js/infoscreen_news.js" type="text/javascript"></script> -->
    <script src="js/datepicker.js" type="text/javascript"></script>
    
    <div class="container" style="width:1300px;">
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
                <table id="news" class="table table-striped table-bordered">
                    <thead>
                        <tr>
                            <th>Overskrift</th>
                            <th style="width:500px;">Brødtekst</th>
                            <th style="width:10%">Vis fra</th>
                            <th style="width:10%">Vis til</th>
                            <th>Billede</th>
                            <th>Type</th>
                            <th>Vigtig</th>
                            <th></th>
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

                          <td style="text-align:center;"></td>

                            <td>
                                <select name="FM_type" class="form-control input-small">  
                                    <option value="1">Intern</option>  
                                    <option value="2">Ekstern</option>                                                                                                          
                                </select>
                            </td>

                            <td style="text-align:center;"><input type="checkbox" name="FM_vigtig" value="1" /></td>

                            <td></td>

                        </tr>
                    </tbody>

                    <tbody>

                        <%
                            strSQL = "SELECT id, overskrift, brodtext, datofra, datotil, editor, vigtig, type, filnavn FROM info_screen ORDER BY datofra DESC, id DESC"
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

                            if len(oRec("filnavn")) <> 0 then
                                facolor = ""
                            else
                                facolor = "dimgrey;"
                            end if
                            
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

                                <td style="text-align:center;"><a onclick="Javascript:window.open('infoscreen_news.asp?func=billedeupload&newsid=<%=oRec("id") %>', '', 'width=650,height=600,resizable=yes,scrollbars=yes')"><span style="color:<%=facolor%>;" class="fa fa-image clickable"></span></a></td>

                                <td>
                                    <%
                                    if oRec("type") = 1 then
                                    internSEL = "SELECTED"
                                    eksternSEL = ""
                                    else
                                    internSEL = ""
                                    eksternSEL = "SELECTED"
                                    end if
                                    %>
                                    <select class="form-control input-small" name="FM_type_<%=oRec("id") %>">
                                        <option value="1" <%=internSEL %>>Intern</option>
                                        <option value="2" <%=eksternSEL %>>Ekstern</option>
                                    </select>
                                </td>

                                <td style="text-align:center">
                                    <%
                                    if oRec("vigtig") = 1 then
                                        vigtigCHB = "CHECKED"
                                    else
                                        vigtigCHB = ""
                                    end if
                                    
                                    %>
                                    <input type="checkbox" name="FM_vigtig_<%=oRec("id") %>" value="1" <%=vigtigCHB %> />
                                </td>

                                <td style="text-align:center"><a href="infoscreen_news.asp?func=slet&id=<%=oRec("id") %>"><span style="color:darkred;" class="fa fa-times"></span></a></td>

                            </tr>
                        <%

                            oRec.movenext
                            wend
                            oRec.close
                        %>

                    </tbody>

                </table>
                </form>

                <%'if x <> 0 then %>

              <!--  <br />

                <div class="row">
                    <div class="col-lg-4">Sidst opdateret af <b><%=strEditor %> </b></div>
                </div> -->

                <%'end if %>

            </div>
        </div>
    </div>

    <%end select 'case func %>



</div>
</div>


<!--#include file="../inc/regular/footer_inc.asp"-->