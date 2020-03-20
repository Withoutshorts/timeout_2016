

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->


<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<script>

    $(document).ready(function() {
            $('.date').datepicker({

           });
    });
</script>

<div class="wrapper">
    <div class="content">
         <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Beregn tillæg & afspadsering til løn</u></h3>

                <div class="portlet-body">
                    


<%
    if len(session("user")) = 0 then

	errortype = 5
	call showError(errortype)
	
	response.End
    end if
    

    func = request("func")

    
    select case func


    case "beregn"

        if len(trim(request("FM_stdato"))) <> 0 then
        stdato = request("FM_stdato")
        else
        stdato = "01-01-2002"
        end if
        
        if len(trim(request("FM_sldato"))) <> 0 then
        sldato = request("FM_sldato")
        else
        sldato = "01-01-2002"
        end if

        %>
         <div class="row">

         <div class="col-lg-6">
        <%

        Response.write "Datointerval:<br>"
        Response.write "<b>"& stdato & " - " & sldato & "</b><br><br>"
        Response.write "Aktive medarbjedere i projektgruppen timeløn:<br><br> "

        %><table style="width:500px;" class="table table-striped">
        <%
        strSQlm = "SELECT mid, mnavn, init FROM medarbejdere WHERE mid <> 0 AND mansat = 1 ORDER BY mnavn"
        oRec.open strSQLm, oConn, 3
        while not oRec.EOF

          forcerun = 1
          'if oRec("mid") = 21 then
            '** kun timelønnede 
            call medarbiprojgrp(41, oRec("mid"), 0, -1)
            if instr(instrMedidProgrp, "#"& oRec("mid") &"#,") <> 0 then

            call fyldoptimertilnorm(lto, oRec("mid"), forcerun, stdato, sldato)
          'end if

            Response.write "<tr><td>" & oRec("mnavn") & "</td><td> ["& oRec("init") &"] </td><td align=right>" & replace(timertilindlasning, ".", ",") & "</td></tr>"

            end if

        oRec.movenext
        Wend
        oRec.close

        %>
          </table>
            <%

      
        Response.Write "<br><br>Opdateringen er klar. Tillæg er indtastet på den sidste dag i den valgte periode. De kan rettes via ugesedlen.<br><br>"
        Response.write "<a href=""Javascript:window.close()"">Luk vindue</a>"

    
        'Response.redirect "tillag_lon_beregn.asp?func=klar"
          
             %>
         </div>
             </div>
        
        <%

   
case else 

  

%>



        



       



              
                    <%if func = "klar" then
                        %>
                        
                     <div class="row">

                        <div class="col-lg-12">

                            Beregning er udført 

                            <br /><br />
                            <a href="Javascript:window.close()">Luk Vindue</a>

                        </div>

                     </div>
                        
                     <%else%>
                    <div class="row">
                        <form action="tillag_lon_beregn.asp?func=beregn" method="post">
                        <div class="col-lg-6 form-inline">

                            <%
                            dd_14 = formatdatetime(dateadd("d", -16, now), 2) 
                            dd_3 = formatdatetime(dateadd("d", -3, now), 2)%>

                            Beregn tillæg og afspadsering før der kører løn for timelønnede.<br /><br />
                            Vælg periode:
                            <br />

                               Fra: 
                              <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="FM_stdato" value="<%=dd_14 %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>

                            til 
                           <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="FM_sldato" value="<%=dd_3 %>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            <br /><br />
                          

                            <input type="submit" value="Kør nu >>" class="form-control input-small btn-success" />

                        </div>
                        </form>
                     </div>
                    <%end if %>


                  
       


       


    <%end select %>

                      </div><!--Portlet body-->
                   </div><!-- Portlet -->
                    
                 


                </div><!-- container -->

</div>
</div>


<!--#include file="../inc/regular/footer_inc.asp"-->

