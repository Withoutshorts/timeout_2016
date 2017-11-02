<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"--> 
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->




<div class="wrapper">
    <div class="content">


        <div class="container">
            <div class="portlet">
               

<%

func = request("func")

select case func
	case "slet"
	'*** Her spørges om det er ok at der slettes en medarbejder ***
	
	
    rap_navn = request("rap_navn")
    rap_id = request("rap_id")

	slttxt = "Du er ved at <b>slette</b> en rapport. <br>"_
    &"<br>Er du sikker på at du vil slette denne rapport: "& rap_navn &"?<br><br>"
	
	slturl = "yourrep.asp?func=sletok&id="&rap_id
		
	slturlalt = ""
	slttxtalt = ""
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,210,100)
	
    response.end
	

	case "sletok"
	'*** Her slettes en rapport ***
    id = request("id")

    oConn.execute("DELETE FROM your_rapports WHERE rap_id = "& id &"")

    response.redirect "yourrep.asp"


    case "dbopr"

          if len(trim(request("saveasrapport_open"))) then
            rap_mid = 0
        else
            rap_mid = session("mid")
        end if


        '*** GEMMER rapport i YOUR RAPPORTS TABLE
        if len(trim(request("saveasrapport_name"))) <> 0 then

        saveasrapport_name = replace(request("saveasrapport_name"), "'", "")
        saveasrapport_criteria = request("saveasrapport_criteria")
        saveasrapport_date = year(now) & "/" & month(now) & "/"& day(now)

        strSQLyourRap = "INSERT INTO your_rapports (rap_mid, rap_navn, rap_url, rap_criteria, rap_dato, rap_editor) "_
        &" VALUES ("& rap_mid &", '"& saveasrapport_name &"', 'joblog_timetotaler.asp?', '"& saveasrapport_criteria &"', '"& saveasrapport_date &"', '"& session("user") &"')"


        oConn.execute(strSQLyourRap)

        

        end if
        
        response.redirect "yourrep.asp"

    case else



  


    
%>

   <h3 class="portlet-title"><u>Your Reports</u></h3>
                <div class="portlet-body">

                       <div class="well">

    
    Hereby a list of your personal saved reports. <br />Timeout can display up to 10 personalreports. They will all be visible on the main menu.
    <br /><br />
	<table width="100%" class="table dataTable table-striped table-bordered table-hover ui-datatable">
        <thead>
            <tr>
            <th style="text-align:left;">Id</th>
            <th style="text-align:left;">Report Name</th>
            <th style="text-align:left;">Type</th>
            <th style="text-align:left;">Visible to all</th>
            <th></th>
            </tr>
        </thead>
	
	
	<%

        level = session("rettigheder")

        if level = 1 then
        all_users = "OR rap_mid = 0"
        else
        all_users = ""
        end if

                        a = 0
                        strSQLyourRap = "SELECT rap_id, rap_mid, rap_navn, rap_url, rap_criteria, rap_dato, rap_editor FROM your_rapports WHERE rap_mid = " & session("mid") & " "& all_users &" ORDER BY rap_navn"
                        oRec6.open strSQLyourRap, oConn, 3
                        while not oRec6.EOF 

                        select case oRec6("rap_url")
                        case "joblog_timetotaler.asp?"
                        rap_type = "Grantotal" 
                        case else
                        rap_type = "s"
                        end select

                        %>
                        <tr><td><%=oRec6("rap_id") %></td><td>
                            <a href="<%=toSubVerPath14 %><%=oRec6("rap_url") & oRec6("rap_criteria")%>" target="_blank"><%=oRec6("rap_navn") %></a>
                            </td>
                            <td><%=rap_type%></td>
                            <td><%if oRec6("rap_mid") = 0 then%>
                                <i>V</i>
                                <%end if %>
                            </td>
                            <td><a href="yourrep.asp?func=slet&rap_id=<%=oRec6("rap_id")%>&rap_navn=<%=oRec6("rap_navn")%>" style="color:red;">[X]</a></td>
                        </tr>
                        <%
                        
                        a = a + 1

                        oRec6.movenext
                        wend
                        oRec6.close
	

	
%>

</table>
Total: <%=a %> reports
<br><br><br>
    <div style="float:right; padding-right:50px;">
<a href="#" onclick="Javascript:window.close()">[Close Window]</a>
        
        </div>
<br /><br />&nbsp;

<%    
    
    end select
    '********************************************************************** %>
        </div>
  </div>
   </div></div><!-- Wrapper , portlet mm -->
      </div>
      </div>
                        

<!--#include file="../inc/regular/footer_inc.asp"-->
 
