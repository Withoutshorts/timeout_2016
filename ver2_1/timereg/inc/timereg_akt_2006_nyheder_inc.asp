<!----------- 
    Fil til nyheder paa timereg og favoritsiden
    ------->


<style>
        /* The Modal (style) */
        .modal {
            display: none; /* Hidden by default */
            position: fixed; /* Stay in place */
            z-index: 1; /* Sit on top */
            padding-top: 100px; /* Location of the box */
            left: 0;
            top: 0;
            width: 100%; /* Full width */
            height: 100%; /* Full height */
            overflow: auto; /* Enable scroll if needed */
            background-color: rgb(0,0,0); /* Fallback color */
            background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
        }

        /* Modal Content */
        .modal-content {
            background-color: #fefefe;
            margin: auto;
            padding: 20px;
            border: 1px solid #888;
            width: 400px;
            height: 350px;
        }

        #closenews:hover,
        #closenews:focus {
        text-decoration: none;
        cursor: pointer;
        } 

        .nyhedsbillede:hover,
        .nyhedsbillede:focus {
        text-decoration: none;
        cursor: pointer;
        }

</style>


<%
    public nyhedsid, nyhedoverskrift, nyhedbrodtext, nyhedvigtig, nyhedpopup, datofra, datotil, filnavn, antalnyheder, newsEditor
    Redim nyhedsid(10), nyhedoverskrift(10), nyhedbrodtext(10), nyhedvigtig(10), nyhedpopup(10), datofra(10), datotil(10), filnavn(10), newsEditor(10)

    function HentNyhedsVariabler()
       
        sqlDTD = year(now) &"-"& month(now) &"-"& day(now)
        strSQL = "SELECT id, overskrift, brodtext, datofra, datotil, vigtig, filnavn, editor FROM info_screen WHERE ('"& sqlDTD &"' >= datofra AND '"& sqlDTD &"' <= datotil) ORDER BY vigtig DESC, datofra DESC, id DESC"
        'response.Write strSQL
        oRec.open strSQL, oConn, 3
        ny = 0
        antalnyheder = 0
        while not oRec.EOF 
            'response.Write "nyhed fundet"
            nyhedsid(ny) = oRec("id")
            nyhedoverskrift(ny) = oRec("overskrift")
            nyhedbrodtext(ny) = oRec("brodtext")
            nyhedvigtig(ny) = oRec("vigtig")
            datofra(ny) = oRec("datofra")
            datotil(ny) = oRec("datotil")
            newsEditor(ny) = oRec("editor")

            if len(oRec("filnavn")) <> 0 then
            filnavn(ny) = oRec("filnavn")
            else
            filnavn(ny) = ""
            end if

            nyhedpopup(ny) = 1
            strSQL2 = "SELECT newsid FROM news_rel WHERE medarbid = "& session("mid") &" AND newsid = "& nyhedsid(ny)
            'response.Write strSQL2
            oRec2.open strSQL2, oConn, 3
            if not oRec2.EOF then
            nyhedpopup(ny) = 0 'Hvis findes skal den ikke poppe frem, fordi saa har medarbejderen l;st nyheden.
            end if
            oRec2.close

        ny = ny + 1
        antalnyheder = antalnyheder + 1
        oRec.movenext
        wend
        oRec.close

    end function



    function NyhedsPopUp()
    
        'Tjekker forst om popupen skal komme frem
        skalpopkommefrem = 0
        for i = 0 TO UBOUND(nyhedsid)

            if cint(nyhedvigtig(i)) = 1 AND cint(nyhedpopup(i)) = 1 then
            skalpopkommefrem = 1
            end if

        next
            'response.Write "skalpopkommefrem = " & skalpopkommefrem
        %>


        <%if cint(skalpopkommefrem) = 1 then %>
        <div id="newsmodal" class="modal" style="display:block; z-index:1000;">
            <!-- Modal content -->
            <div class="modal-content" style="width:85%; height:75%;">

                <div style="height:75%; overflow-y:scroll;">
                <table>
                    <%                
                    for i = 0 TO UBOUND(nyhedsid)
                    %>
                    <%if cint(nyhedvigtig(i)) = 1 AND cint(nyhedpopup(i)) = 1 then  %>
                    <tr>
                        <td>                        
                            <h3 style="color:#f44842;"><%=nyhedoverskrift(i) %> <br /> <span style="font-size:9px; color:#9e9e9e;">Oprettet: <%=datofra(i) %> af <%=newsEditor(i) %></span></h3>

                            <%if nyhedbrodtext(i) <> "" then  %>
                            <span style="word-break: break-all;"><%=nyhedbrodtext(i) %></span>
                            <%else %>
                            <h5>Ingen beskrivelse</h5>
                            <%end if %>

                            <%if filnavn(i) <> "" then %>
                               <br /> <img src="../inc/upload/<%=lto%>/<%=filnavn(i) %>" alt='' border='0' style="width:500px;">
                            <%end if %>
                        </td>                                
                    </tr>
                    <tr>
                        <td>&nbsp</td>
                    </tr>
                    <%end if %>
                    <%
                    next
                    %>

                </table>
                </div>

                <br /><br />
           
                <div class="row" style="margin-left:45%;">
                    <div class="col-lg-12"><button type="submit" id="closenews" class="btn btn-default btn-sm"><b>OK</b></button></div>
                </div>

            </div>
        </div>
        <%end if


    end function
%>


