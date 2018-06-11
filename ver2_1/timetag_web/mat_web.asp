
<%
    thisfile = "timetag_mobile" 
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

      
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

    
    <!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
    <!--#include file="../inc/regular/top_menu_mobile.asp"-->



<head>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TimeOut mobile</title>



<script src="js/timetag_web_jav_2017_023.js" type="text/javascript"></script>




    <style type="text/css">

        input[type="text"] 
        {
          height:125%;
          font-size:125%;
        }
        input[type="button"] 
        {
          height:125%;
          font-size:125%;
        }

        .span_job {
        list-style:none;
        font-size:125%;
        }

         .span_akt {
        list-style:none;
        font-size:125%;
        }

        .span_mat {
        list-style:none;
        font-size:125%;
        }

        

    </style>

</head>
    

<%call mobile_header %>

        <%
             if len(trim(request("FM_month"))) <> 0 then
             selectedMonth = request("FM_month")
             else
             selectedMonth = month(now)
             end if

             if len(trim(request("FM_year"))) <> 0 then
             selectedYear = request("FM_year")
             else
             selectedYear = year(now)
             end if

            sqlMatDatoStart = selectedYear &"-"& selectedMonth &"-01"
            sqlMatDatoSlut = selectedYear &"-"& selectedMonth &"-31"

            preSqlMatDate = dateadd("m", -1, sqlMatDatoStart)
            nextSqlMatDate = dateadd("m", 1, sqlMatDatoStart)

            preMonth = month(preSqlMatDate)
            nxtMonth = month(nextSqlMatDate)
            preYear = year(preSqlMatDate)
            nxtYear = year(nextSqlMatDate)
           
        %>



<div class="container" style="height:100%">
    <div class="portlet">
       
        <input type="hidden" name="FM_month" value="<%=selectedMonth %>" />
        <input type="hidden" name="FM_year" value="<%=selectedYear %>" />

        <div class="row">
            <h4 class="col-lg-12"><a href="mat_web.asp?FM_month=<%=preMonth %>&FM_year=<%=preYear %>"><</a> &nbsp <%=UCase(Left(monthName(selectedMonth),1)) & LCase(Right(monthName(selectedMonth), Len(monthName(selectedMonth)) - 1)) & " " & selectedYear %> &nbsp <a href="mat_web.asp?FM_month=<%=nxtMonth %>&FM_year=<%=nxtYear %>">></a></h4>
        </div>

        <table class="table table-striped table-bordered">

            
            <thead>
                <tr>
                    <th>Materiale/Job</th>
                    <th style="text-align:right;">Antal</th>
                    <th style="text-align:right;">Pris/stk</th>
                    <th style="text-align:right;">I alt</th>
                </tr>
            </thead>


            <tbody>

                <%
                    matantal = 0
                    salgsprisStk = 0
                    totalForPeriode = 0
                    strSQL = "SELECT matnavn, j.jobnavn as jobnavn, forbrugsdato, matvarenr, matantal, matsalgspris, kurs FROM materiale_forbrug LEFT JOIN job as j ON (j.id = jobid) WHERE usrid = "& session("mid") &" AND forbrugsdato between '"& sqlMatDatoStart &"' AND '"& sqlMatDatoSlut &"'"
                    oRec.open strSQL, oConn, 3

                    x = 0

                    while not oRec.EOF
                   
                    matantal = oRec("matantal")
                    salgsprisStk = oRec("matsalgspris")
                    salgsprisStk = salgsprisStk * oRec("kurs")/100

                    prisIalt = matantal * salgsprisStk
                    prisIalt = prisIalt * oRec("kurs")/100

                    if x = 0 then
                    totalForPeriode = prisIalt
                    else 
                    totalForPeriode = totalForPeriode + prisIalt
                    end if

                %>

                <tr>
                    <td style="vertical-align:middle;"><%=oRec("matnavn") & " " %> (<%=oRec("matvarenr") %>)<br />
                        <span style="font-size:10px"><%=oRec("jobnavn") %></span> <br />
                        <span style="font-size:10px"><i><%=oRec("forbrugsdato") %></i></span>
                    </td>   
                    <td style="text-align:right; vertical-align:middle;"><%=matantal %></td>
                    <td style="text-align:right; vertical-align:middle;"><%=formatnumber(salgsprisStk,2) %></td>
                    <td style="text-align:right; vertical-align:middle;"><%=formatnumber(prisIalt,2)%></td>
                </tr>
                <%
                    x = x + 1
                    oRec.movenext
                    wend
                    oRec.close
                %>

            </tbody>

            <tfoot>
                <tr>
                    <th style="text-align:left;">Ialt</th>
                    <th colspan="3"; style="text-align:right; border-left:hidden"><%=Formatnumber(totalForPeriode,2) %> kr.</th>
                </tr>
            </tfoot>

        </table>

    </div>
</div>


           


<!--#include file="../inc/regular/footer_inc.asp"-->
