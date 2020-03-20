<%
session("mid") = 1
session("login") = 1
session("user") = "Timeout - support"
session("lto") = request("ltokey")
session("rettigheder") = 1
%>



<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<script src="js/udeafhusetservice_jav.js" type="text/javascript"></script>      
<div class="wrapper">
    <div class="content">

        <div class="container">
            <h3 class="portlet-title"><u>Ude af huset service</u></h3>

            <div class="portlet-body">

                <table class="table dataTable table-striped table-bordered table-hover ui-datatable">

                    <thead>
                        <tr>
                            <th>Udeafhuset - type</th>
                            <th>Medarbejder</th>
                            <th>Fradato</th>
                            <th>Tildato</th>
                        </tr>

                        <%
                        sqlNow = year(now) &"-"& month(now) &"-"& day(now)
                        strSQL = "SELECT id, medid, fradato, tildato, udeafhusettype FROM udeafhuset WHERE udeafhusettype <> 0" 
                        'response.Write strSQL
                        oRec.open strSQL, oConn, 3
                        x = 0
                        while not oRec.EOF

                            fradatoSQL = year(oRec("fradato")) &"-"& month(oRec("fradato")) &"-"& day(oRec("fradato"))
                            tildatoSQL = year(oRec("tildato")) &"-"& month(oRec("tildato")) &"-"& day(oRec("tildato"))

                            if cdate(fradatoSQL) < cdate(sqlNow) AND cdate(tildatoSQL) > cdate(sqlNow) then
                                
                                call normtimerPer(oRec("medid"), sqlNow, 0, 0)
                                antal = ntimper '7 'ntimper
                        
                                select case oRec("udeafhusettype")
                                    case cint(2)
                                        aktid = 13 'Ferie
                                    case cint(1)
                                        aktid = 33 'Forretningsrejse
                                    case cint(3)
                                        aktid = 32 'Work from home 
                                    case cint(4)
                                        aktid = 30 'Syg

                                    case else
                                        aktid = 0
                                end select

                                extsysid = oRec("id")

                                timerkom = ""
                                koregnr = "" 
                                destination = ""
                                usebopal = 0

                                call indlasTimerTfaktimAktid(lto, oRec("medid"), antal, 0, aktid, 0, sqlNow, extsysid, timerkom, koregnr, destination, usebopal)

                                response.Write "<tr><td>"&oRec("udeafhusettype")&"</td> <td>"&oRec("medid")&"</td> <td>"&oRec("fradato")&"</td> <td>"&oRec("tildato")&" antal "& antal &"</td></tr>"

                                x = x + 1
                            end if

                        oRec.movenext
                        wend
                        oRec.close

                        if x = 0 then
                            response.Write "<tr><td colspan='10' style='text-align:center'>Ingen udeafhuest til registrering</td></tr>"
                        end if

                        %>

                    </thead>

                </table>

            </div>

        </div>


    </div>
</div>

<script type="text/javascript">
    $(document).ready(function () {

    setTimeout(function () {
        CloseWindow()
    }, 7000);

    });

    function CloseWindow() {
        window.top.close();
    }
</script>


<!--#include file="../inc/regular/footer_inc.asp"-->