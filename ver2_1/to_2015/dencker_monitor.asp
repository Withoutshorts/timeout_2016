

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<style>

    .btn 
    {
        font-size: 150%;
        width:100px;
    }

</style>


<%
    if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
       response.End
	end if 

    medid = session("user")
    usemrn = session("mid")
    func = request("func") 
%>

<script src="js/dencker_monitor.js" type="text/javascript"></script>
<!--
<link href="http://netdna.bootstrapcdn.com/twitter-bootstrap/2.2.2/css/bootstrap-combined.min.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" media="screen"
href="http://tarruda.github.com/bootstrap-datetimepicker/assets/css/bootstrap-datetimepicker.min.css">
-->


<div class="wrapper">
    <div class="content">

        <div class="container">
            <div class="portlet">
                <div class="portlet-body">

                   


                    <!--<div id="datetimepicker" class="input-append date">
                    <input type="text"></input>
                    <span class="add-on">
                    <i data-time-icon="icon-time" data-date-icon="icon-calendar"></i>
                    </span>
                    </div> -->
                    
                    
                   <!-- <script type="text/javascript"
                    src="http://tarruda.github.com/bootstrap-datetimepicker/assets/js/bootstrap-datetimepicker.min.js">
                    </script>
                    <script type="text/javascript"
                    src="http://tarruda.github.com/bootstrap-datetimepicker/assets/js/bootstrap-datetimepicker.pt-BR.js">
                    </script>
                    <script type="text/javascript">
                      $('#datetimepicker').datetimepicker({
                        format: 'dd/MM/yyyy hh:mm:ss',
                        language: 'pt-BR'
                      });
                    </script> -->





                    <div class="row">
                        <div class="col-lg-12" style="text-align:center;>
                            <a href="#"><img src="img/Dencker_logo.png" width="100" /></a>
                        </div>
                    </div>
                    
                    <%                        
                        select case func
                        case "startside"
                        

                        strSQL = "SELECT id, login, logud FROM login_historik WHERE mid="& usemrn &" ORDER by id"
                        'response.Write strSQL
                        oRec.open strSQL,oConn, 3
                        'resonse.write strSQL
                        while not oRec.EOF
                        lastid = oRec("id")
                        login = oRec("login")
                        logud = oRec("logud")

                        'response.write oRec("logud")
                        oRec.movenext                     
                        wend
                        oRec.close

                        'response.Write lastid & login

                        'if isNull(logud) <> false then 
                        'response.Write "logud nu"
                        'end if
                        
                    %>
                    

                    <%if isNull(logud) <> false then %>

                    <div class="row" style="margin-top:10%">
                        <div class="col-lg-12" style="text-align:center"><a href="dencker_monitor.asp?func=gaa&lastid=<%=lastid %>" class="btn btn-default" style="width:200px;"><b>Scan kort</b></a></div>
                    </div>
                    <%else %>
                    <div class="row" style="margin-top:10%">
                        <div class="col-lg-12" style="text-align:center"><a href="dencker_monitor.asp?func=kom" class="btn btn-default" style="width:200px;"><b>Scan kort</b></a></div>
                    </div>
                    <%end if %>



                    <%
                        case "kom"
                        'response.Write "id: " & usemrn 
                    %>


                    <div class="row" style="margin-top:10%">
                        <div class="col-lg-1"></div>
                        <div class="col-lg-2"><a href="dencker_monitor.asp?func=login" class="btn btn-success nor_btn"><b>Normal</b></a></div>
                        <div class="col-lg-1"></div>            
                        <div class="col-lg-2"><a type="submit" class="btn btn-success"><b>Extern</b></a></div>
                        <div class="col-lg-1"></div>            
                        <div class="col-lg-2"><a type="submit" class="btn btn-success"><b>U. løn</b></a></div>
                        <div class="col-lg-1"></div>            
                        <div class="col-lg-2"><a type="submit" class="btn btn-success"><b>Tilkald</b></a></div>
                    </div>

                    
                    <%case "gaa" 
                    
                        lastid = request("lastid")
                        'response.Write lastid
                            
                    %>

                    <div class="row" style="margin-top:10%">
                        <div class="col-lg-12" style="text-align:center"><a data-toggle="modal" href="#styledModalSstGrp20" class="btn btn-danger"><b>Gå</b></a></div>
                    </div>

                    <div id="styledModalSstGrp20" class="modal modal-styled fade" style="top:100px;">
                        <div class="modal-dialog">
                            <div class="modal-content" style="border:none !important;padding:0;">
                                <div class="modal-body">
                                    
                                    <div class="row">
                                        <h2 class="col-lg-12" style="text-align:center">Tak for i dag</h2>
                                    </div>

                                    <br /><br />

                                    <%
                                        strSQLinfo = "SELECT login, logud FROM login_historik WHERE id="& lastid
                                        'response.Write strSQLinfo
                                        oRec.open strSQLinfo, oConn, 3
                                        loguddate = Year(now) & "-" & Month(now) & "-" & Day(now) & " " & Time 
                                        
                                        timerialt = DateDiff("n",oRec("login"),loguddate) / 60
                                        
                                        
                                        %>
                                    
                                        

                                        <div class="row">
                                            <div class="col-lg-3">Din logind tid:</div><div class="col-lg-4"><%response.Write oRec("login")  %></div>
                                        </div>
                                        <div class="row">
                                            <div class="col-lg-3">Din logud tid:</div><div class="col-lg-4"><%response.Write Right("0" & DatePart("d",Date), 2) & "-" & Right("0" & DatePart("m",Date), 2) & "-" & Year(Now) & " " & Time  %></div>
                                        </div>
                                        

                                    <!--    <div class="form-group">
                                        <label class="control-label col-md-2">24 Hour Mode:</label>
                                        <div class="col-md-10">
                                        <div class="bootstrap-timepicker">
                                        <input type="text" id="timepicker5" class="form-control" placeholder="Choose Time">
                                        </div> 
                                        </div>
                                        </div> <!-- /.form-group --> 

                                        <br />
                                        <div class="row">
                                            <div class="col-lg-12">Timer: <%=formatnumber(timerialt, 2)%></div>
                                        </div>


                                        <%                                     
                                        oRec.close                                         
                                        %>
                                       
                                        <div class="row" style="margin-top:10%">
                                            <div class="col-lg-12" style="text-align:center"><a data-toggle="modal" href="dencker_monitor.asp?func=logud&lastid=<%=lastid %>" class="btn btn-success" style="width:200px;"><b>Godkend</b></a></div>
                                        </div>

                                </div>
                            </div>
                        </div>
                    </div>

                    <%case "login" 
                              
                        datoidag = Year(now) & "-" & Month(now) & "-" & Day(now)
                        logintid = Year(now) & "-" & Month(now) & "-" & Day(now) & " " & Time
                        
                        response.Write "login date:" & logintid 
                              
                        strSQLlogin = "INSERT INTO login_historik SET dato ='"& datoidag &"', login ='"& logintid & "', mid ="& usemrn
                        response.write strSQLlogin
                        oConn.execute(strSQLlogin)

                        response.Redirect("dencker_monitor.asp?func=startside")

                       
                    %>
                    <div class="row" style="margin-top:10%">
                        <h4 class="col-lg-12" style="text-align:center"><%response.Write medid & "<br>" & Date %></h4>
                    </div>



                    <%case "logud"
                    lastid = request("lastid")

                    logudtid = Year(now) & "-" & Month(now) & "-" & Day(now) & " " & Time

                    response.Write logudtid
                    'response.Write lastid

                    strSQLlogud = "UPDATE login_historik SET logud ='"& logudtid &"' WHERE id ="& lastid
                    
                    oConn.execute(strSQLlogud)


                    response.Redirect("dencker_monitor.asp?func=startside")
                    %>

                    <%
                        end select
                    %>

                </div>
            </div>
        </div>

    </div>
</div>
    

<!--#include file="../inc/regular/footer_inc.asp"-->