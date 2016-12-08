
<%
    thisfile = "timetag_kpi" 
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<!--#include file="../inc/regular/top_menu_mobile.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/treg_func.asp"-->

<%
    if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/errors/error_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	response.End
    end if

    call meStamdata(session("mid"))
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">



<head>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TimeOut mobile</title>

</head>
    



    <%call mobile_header %>


    <div class="container" style="height:100%">
        <div class="portlet">
            <div class="portlet-body">
                
                <!------ Flex saldo ----->
                <%
                    

                    usemrn = session("mid")

        
                    
                    sqlDatoNow = year(now)&"/"&month(now)&"/"&day(now)
                    daysansat = (DateDiff("d",meAnsatDato,sqlDatoNow,vbMonday)) -2


                    strSQL2 = "SELECT tmnavn, SUM(timer) FROM timer WHERE tmnr = "&usemrn&" AND tdato BETWEEN "&meAnsatDato&" AND '"&sqlDatoNow&"'"
                    oRec2.open strSQL2, oConn, 3
                    if not oRec2.EOF then
                    sumtimer = oRec2("SUM(timer)")
                    end if

                    oRec2.close
                    oConn.execute(strSQL2)

                    response.Write "sumtimer:" & sumtimer
                              
                    call normtimerPer(usemrn, meAnsatDato, daysansat, 0)
                    
                    flexsaldo = sumtimer - ntimPer



                    'response.write "<br> dateansat: " & meAnsatDato
                    'response.write "<br> dateidag: " & sqlDatoNow
                    'response.write "<br> days: " & daysansat
                    'response.write "<br> normtimer: " & ntimPer
                    'response.write "<br> opgjort: " & sumtimer


                    'response.Write "tal: " & sumtimer
                %>

                <!------ Ferie saldo ----->
                <%
                     call akttyper2009(2)
                     
                    strSQL3 = "SELECT SUM(timer) FROM timer WHERE tmnr = "&usemrn&" AND tfaktim = 14"
                    oRec3.open strSQL3, oConn, 3
                    if not oRec3.EOF then
                    ferieafholdt = oRec3("SUM(timer)")
                    end if
                    oRec3.close
                    oConn.execute(strSQL3)

                    if ferieafholdt > 0.1 then
                        ferieafholdt = ferieafholdt
                    else
                        ferieafholdt = 0
                    end if

                    response.Write "<br> afholdt:" & ferieafholdt
                    

                    strSQL4 = "SELECT SUM(timer) FROM timer WHERE tmnr = "&usemrn&" AND tfaktim = 15"
                    oRec4.open strSQL4, oConn, 3
                    if not oRec4.EOF then
                    ferieoptjent = oRec4("SUM(timer)")
                    end if
                    oRec4.close
                    oConn.execute(strSQL4)

                    if ferieoptjent > 0.1 then
                        ferieoptjent = ferieoptjent
                    else
                        ferieoptjent = 0
                    end if

                    response.Write "<br> optjent:" & ferieoptjent


                    flexferie = ferieoptjent - ferieafholdt

                    'response.Write "<br> tal: " & ferieoptjent
                %>

                <!------ sygetal ----->
                <%
                    strSQL5 = "SELECT SUM(timer) FROM timer WHERE tmnr = "&usemrn&" AND tfaktim = 20 AND tdato BETWEEN "&meAnsatDato&" AND '"&sqlDatoNow&"'"
                    oRec5.open strSQL5, oConn, 3
                    if not oRec5.EOF then
                    sygtot = oRec5("SUM(timer)")
                    end if
                    oRec5.close
                    oConn.execute(strSQL5)

                    if sygtot > 0.4  then 
                        sygtot = sygtot
                    else
                        sygtot = 0
                    end if

                    strSQL6 = "SELECT SUM(timer) FROM timer WHERE tmnr = "&usemrn&" AND tfaktim = 21 AND tdato BETWEEN "&meAnsatDato&" AND '"&sqlDatoNow&"'"
                    oRec6.open strSQL6, oConn, 3
                    if not oRec6.EOF then
                    barnsygtot = oRec6("SUM(timer)")
                    end if
                    oRec6.close
                    oConn.execute(strSQL6)

                    if barnsygtot > 0.4  then 
                        barnsygtot = barnsygtot
                    else
                        barnsygtot = 0
                    end if

                %>

                <table class="table dataTable" style="font-size:125%">

                    <tr>
                        <td>Ferie saldo</td>
                        <td style="text-align:left"><%=flexferie %></td>
                        <td style="vertical-align:middle"><div style="height:10px; width:10px; background-color:red;"></div></td>
                    </tr>

                    <tr>
                        <td>Flex saldo</td>
                        <td style="text-align:left"><%=(Round(flexsaldo,2)) %></td>
                        <td style="vertical-align:middle"><div style="height:10px; width:10px; background-color:greenyellow;"></div></td>
                    </tr>

                    <tr>
                        <td>Syg åtd</td>
                        <td style="text-align:left"><%=(Round(sygtot,2)) %></td>
                        <td style="vertical-align:middle"><div style="height:10px; width:10px; background-color:red;"></div></td>
                    </tr>

                    <tr>
                        <td>Barn syg åtd</td>
                        <td style="text-align:left"><%=(Round(barnsygtot,2)) %></td>
                        <td style="vertical-align:middle"><div style="height:10px; width:10px; background-color:red;"></div></td>
                    </tr>

                   <!-- <tr>
                        <td>Ferie fridage</td>
                    </tr>
                    <tr>
                        <td>Felx saldo</td>
                    </tr>
                    <tr>
                        <td>Syg åtd</td>
                    </tr>
                    <tr>
                        <td>Barn syg åtd</td>
                    </tr> -->

                </table>

            </div>
        </div>
    </div>



</html>
<!--#include file="../inc/regular/footer_inc.asp"-->