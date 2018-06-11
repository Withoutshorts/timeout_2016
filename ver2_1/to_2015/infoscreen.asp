

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<!--<div class="wrapper">
<div class="content"> -->


<%

 if len(session("user")) = 0 then
	errortype = 5
	call showError(errortype)
    response.End
 end if


 sqlDTD = year(now) &"-"& month(now) &"-"& day(now)
    
%>


    <style>
        .sorter:hover,
        .sorter:focus {
        text-decoration: none;
        cursor: pointer;
        }
    </style>


    <script src="js/infoscreen_jav2.js" type="text/javascript"></script> 

    <div class="container" style="width:100%; height:100%">
        <div class="portlet">
            <div class="portlet-body">

                <%
                    dateNow = day(now) &"-"& monthName(month(now)) &"-"& year(now)
                %>

                <div class="row">
                    <h2 class="col-lg-4" style="color:gray"><%=dateNow %> &nbsp <span id="live_clock"></span></h2>
                </div>
                
                <div class="row">
                    <div class="col-lg-6"><h3>Info Skærm</h3></div>
                    <div class="col-lg-6"><h3>Nyheder</h3></div>
                </div>

                <br />

                <div class="row">
                <div class="col-lg-6">
                <table style="width:100%" id="myTable">
                    <tbody>

                        <tr>
                            <td><span style="color:dimgray;" id="sort_name" class="sorter fa fa-sort-down"></span></td>
                            <td style="padding-left:15px;"><span id="sort_color" style="color:dimgray;" class="sorter fa fa-sort-down"></span></td>
                        </tr>

                        <%

                       
                        medarbejere = 0
                        antal = 0
                        if lto <> "dencker" then
                            strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE mansat = 1 AND mid <> 1"
                        else 'dencker
                            strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE mansat = 1 AND mid <> 21"
                        end if

                        oRec.open strSQL, oConn, 3
                        while not oRec.EOF
                            'response.Write oRec("mid")
                        %>
                        <!--
                            syg = 20 = red
                            barnsyg = 21 = red

                            skole = 91 = green (info)
                            kursus = 91 = green (info)

                            Barsel = 22 = blue
                            ferieafholdt = 14 = blue
                            feriefridagebrugt = 13 = blue
                            afspadering = 31 = blue

                            gul ude af huset. 
                        -->
                            
                        <%
                            timerColor = ""
                            timerReg = 0
                            strSQL2 = "SELECT timer, tfaktim from timer WHERE tmnr = "& oRec("mid") &" AND tdato = '"& sqlDTD &"' AND (tfaktim = 20 OR tfaktim = 21 OR tfaktim = 22 OR tfaktim = 14 OR tfaktim = 13 OR tfaktim = 31 OR tfaktim = 91) GROUP BY tfaktim = 20, tfaktim = 21, tfaktim = 22, tfaktim = 14, tfaktim = 13, tfaktim = 31, tfaktim = 91 ORDER BY tid"

                            oRec2.open strSQL2, oConn, 3                            
                            while not oRec2.EOF
 
                            timerReg = oRec2("timer")

                            select case oRec2("tfaktim")

                                case 20
                                timerColor = "#d9534f"
                                timerTxt = "Syg"
                                sortNumber = "1"
                                case 21
                                timerColor = "#d9534f"
                                timerTxt = "Barnsyg"
                                sortNumber = "2"

                                case 91
                                timerColor = "#5cb85c"
                                timerTxt = "Kursus / Skole"
                                sortNumber = "3"

                                case 22
                                timerColor = "#5bc0de"
                                timerTxt = "Barsel"
                                sortNumber = "6"
                                case 14
                                timerColor = "#5bc0de"
                                timerTxt = "Ferie"
                                sortNumber = "4"
                                case 13
                                timerColor = "#5bc0de"
                                timerTxt = "Feriefri"
                                sortNumber = "5"
                                case 31
                                timerColor = "#5bc0de"
                                timerTxt = "Afspads."
                                sortNumber = "7"
                                
                            end select

                            antal = antal + 1

                            oRec2.movenext
                            wend
                            oRec2.close
                        %>

                        <%


                            erUde = 0
                            nowTime = hour(now) &":"& minute(now)

                            'response.Write "Klokken lige nu er " & nowTime
                            dateTimeNow = sqlDTD &" "& nowTime

                            strSQLudeafhuset = "SELECT id FROM udeafhuset WHERE medid = "& oRec("mid") &" AND ((fra <= '"& dateTimeNow &"' AND til >= '"& dateTimeNow &"') OR (fradato <= '"& sqlDTD &"' AND tildato >= '"& sqlDTD &"' AND heledagen = 1))"
                            'response.Write strSQLudeafhuset 
                            oRec3.open strSQLudeafhuset, oConn, 3
                            if not oRec3.EOF then
                                erUde = 1
                            end if
                            oRec3.close
                           'response.Write "<br>areude " & erUde

                            if erUde = 1 then
                                timerTxt = "Ude af huset"
                                timerColor = "#f0ad4e"
                                sortNumber = "8"
                            end if



                            medarbLoggetInd = 0                           
                            strSQL3 = "SELECT mid FROM login_historik WHERE mid = "& oRec("mid") &" AND dato = '"& sqlDTD &"'"
                            oRec3.open strSQL3, oConn, 3
                            if not oRec3.EOF then
                            medarbejere = medarbejere + 1
                            medarbLoggetInd = 1
                            end if
                            oRec3.close

                            if medarbLoggetInd = 0 AND timerReg = 0 AND erUde = 0 then
                                timerTxt = "Ikke mødt"
                                timerColor = "grey"
                                sortNumber = "9"
                            end if

                        %>

                        <%if medarbLoggetInd <> 1 OR timerReg <> 0 OR erUde <> 0 Then %>
                       <!-- <tr>
                            <td style="text-align:center"><span style="color:<%=timerColor %>; font-size:25px;">O</span></td>
                            <td><%=oRec("mnavn") %></td>
                        </tr> -->

                        <tr>
                            <td><input type="hidden" value="<%=oRec("mnavn") %>" />
                                <div class="alert" style="height:100%; background-color:<%=timerColor %>;">
                                    <strong style="color:white"><%=oRec("mnavn") %></strong>
                                </div>
                            </td>

                            <td style="padding-left:15px">
                                <input type="hidden" value="<%=sortNumber %>" />
                                <div class="alert" style="height:100%; background-color:grey"">
                                    <strong style="color:white"><%=timerTxt %></strong>
                                </div>
                            </td>
                        </tr>

                        <%end if %>
                                                                   
                        <%
                        oRec.movenext
                        wend
                        oRec.close
                        %>

                        <%if antal = 0 AND medarbejere = 0 then%>
                            <tr>
                                <td>
                                    <div class="alert" style="height:100%; background-color:grey"">
                                        <strong style="color:white">Alle Medarbejdere er på arbejde i dag</strong>
                                    </div> <!-- /.alert -->
                                </td>
                            </tr>
                        <%end if %>

                    </tbody>
                </table>
                </div>

                    <div class="col-lg-6">
                        <table>
                            <tr>
                                <td>&nbsp</td>
                                <td>&nbsp</td>
                            </tr>
                            <%
                                strSQL = "SELECT overskrift, brodtext, datofra, datotil FROM info_screen WHERE ('"& sqlDTD &"' >= datofra AND '"& sqlDTD &"' <= datotil) ORDER BY id"
                                'response.Write strSQL
                                oRec.open strSQL, oConn, 3
                                x = 0
                                while not oRec.EOF 
                            %>


                            <tr>
                                <td><h3><%=oRec("overskrift") %></h3></td>                                
                            </tr>
                            <tr>
                                <td>
                                    <%if oRec("brodtext") <> "" then  %>
                                    <span><%=oRec("brodtext") %></span>
                                    <%else %>
                                    <h5>Ingen beskrivelse</h5>
                                    <%end if %>

                                </td>
                            </tr>
                            <tr>
                                <td>&nbsp</td>
                            </tr>


                            <%
                                x = x + 1
                                oRec.movenext
                                wend
                                oRec.close
                            %>


                            <%if x = 0 then %>

                            <tr>
                                <td><h3>Ingen nyheder i dag</h3></td>
                            </tr>

                            <%end if %>

                        </table>
                    </div>

                </div>

            </div>
        </div>
    </div>




<!--
</div>
</div> -->


<!--#include file="../inc/regular/footer_inc.asp"-->