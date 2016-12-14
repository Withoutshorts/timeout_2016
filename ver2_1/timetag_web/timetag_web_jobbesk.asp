
<%
    thisfile = "timetag_jobbesk" 
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->



<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/top_menu_mobile.asp"-->




<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">



<head>

<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TimeOut mobile</title>

</head>
    



    <%
        if len(trim(request("id"))) <> 0 then
        id = request("id")
        else
        id = 0
        end if
        
        call mobile_header 
        
        
        jobbesk = ""
        jobnavn = ""
        strSQLjobbesk = "SELECT jobnavn, beskrivelse FROM job WHERE id = "& id
        oRec.open strSQLjobbesk, oConn, 3 
        if not oRec.EOF then
        
        jobbesk = oRec("beskrivelse")
        jobnavn = oRec("jobnavn")

        end if
        oRec.close   
        %>


    <div class="container" style="height:100%">
        <div class="portlet">
            <div class="portlet-body">
                
                <h4><%=jobnavn%> - Jobbeskivlese:</h4>
                 <div class="row">
                          <div class="col-lg-12">
                              <%=jobbesk %>
                          </div> 

                 </div>
            </div>
        </div>
    </div>



</html>
<!--#include file="../inc/regular/footer_inc.asp"-->