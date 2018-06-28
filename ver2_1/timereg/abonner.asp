<%response.Buffer = true 
tloadA = now
%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->




<%
    if len(session("user")) = 0 then

    %>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
    

    func = request("func")

    
    select case func


    case "abo"

      '**** Abonner ****'
            

            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_admin;"
            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=root;pwd=;database=timeout_admin;"
            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
            strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
            Set oConn_admin = Server.CreateObject("ADODB.Connection")
            Set oRec_admin = Server.CreateObject ("ADODB.Recordset")

            oConn_admin.open strConnect_admin
            
            editor = session("user")

            'Response.write request("FM_abo_mid") & "<br>"

            aMedid = split(request("FM_abo_mid"), ", ") 'session("mid")
            
            aMedTyp = request("FM_abo_mtyp")
            
            if trim(aMedTyp) = "-1" then 'kan kun abonnere på enten eller
            aProgrp = request("FM_abo_progrp")
            else
            aProgrp = "-1"
            end if 

            lto = lto
            
            
            
            
            if request("FM_abo") = "1" then 'aboner
            
           for m = 0 to UBOUND(aMedid)
            

            'Response.write aMedid(m)

            dtNowSQL = year(now) &"/"& month(now) &"/"& day(now)
          
            call meStamdata(aMedid(m))

         

            'if meBrugergruppe = 3 then 'Admin
            if len(trim(request("FM_rapporttype"))) <> 0 then
            rapporttype = request("FM_rapporttype")
            
                    if cint(rapporttype) = 2 then
                    'Jobansvarlig rapport - Viser for alle medarbejdere
                    aMedTyp = 0
                    aProgrp = 0
                    end if

            else
            rapporttype = 0
            end if 
            
            'else
            'rapporttype = 1 'Viser kun tal for sig sel
            'end if

                '** Findes den i forvejen ==> gør intet ***'


                '*** 20170104 ÆNDRET: Man skal kunne abonnere på flere typer af rapporter ***
                'strSQLAbo = "SELECT medid FROM rapport_abo WHERE medid = "& aMedid(m) & " AND lto = '"& lto & "'"
                'oRec_admin.open strSQLAbo, oConn_admin, 3
                'if not oRec_admin.EOF then

                'else '**Insert

                 strSQLAbo = "INSERT INTO rapport_abo (dato, editor, lto, navn, email, medid, rapporttype, akttyper, medarbejdertyper, projektgrupper) "_
                 &" VALUES ('"& dtNowSQL &"','"& editor &"', '"& lto &"', '"& meNavn &"', '"& meEmail &"', "& aMedid(m) &", "& rapporttype &", '#0#', '"& aMedTyp &"', '"& aProgrp &"' )"
                 oConn_admin.execute(strSQLAbo)

                'end if
                'oRec_admin.close
          
            
            next


            else '**slet
               

            

            end if
            
            
          
            oConn_admin.close

            'Response.end


            Response.redirect "abonner.asp"

        case "abo_slet"


            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=62.182.173.226;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_admin;"
            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
            strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;"
            Set oConn_admin = Server.CreateObject("ADODB.Connection")
            Set oRec_admin = Server.CreateObject ("ADODB.Recordset")

            oConn_admin.open strConnect_admin
            
            aMedid = request("FM_abo_mid") 'session("mid")
            lto = lto


         strSQLAbo = "DELETE FROM rapport_abo WHERE medid = "& aMedid & " AND lto = '"& lto & "'"
               'Response.write strSQLAbo
                oConn_admin.execute(strSQLAbo)


         Response.redirect "abonner.asp"
   
case else
	

        %>
<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	

 <script src="inc/abonner_jav.js"></script>
<%

    call menu_2014()

tLeft = 20
tWdth = 1004

tTop = 120
oskrift = abonner_txt_034

    call filterheader_2013(100,20,1004,oskrift)
'call tableDiv(tTop,tLeft,tWdth)


            aboCHK = ""
         
             if (session("mid") = 1 OR session("mid") = 68) AND lto = "outz" then
            %>
            <div style="float:right; padding-right:20px; border:1px #999999 solid;">
                         <span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=bf&user=<%=session("user") %>" target="_blank">Send BF reports manuel now >> </a></span>
            
                    <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=dencker&user=<%=session("user") %>" target="_blank">Send Dencker reports manuel now >> </a></span>
                    <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=tia&user=<%=session("user") %>" target="_blank">Send TIA reports manuel now >> </a></span>
                    <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=epi2017&user=<%=session("user") %>" target="_blank">Send EPI2017 reports manuel now >> </a></span>
                <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=synergi1&user=<%=session("user") %>" target="_blank">Send Synergi reports manuel now >> </a></span>
                
                <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=cisu&user=<%=session("user") %>" target="_blank">Send Cisu reports manuel now >> </a></span>
                <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=eniga&user=<%=session("user") %>" target="_blank">Send Eniga reports manuel now >> </a></span>
                <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=essens&user=<%=session("user") %>" target="_blank">Send Essens reports manuel now >> </a></span> 
                <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=jttek&user=<%=session("user") %>" target="_blank">Send JtTek reports manuel now >> </a></span>
                
                <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=oko&user=<%=session("user") %>" target="_blank">Send Oko reports manuel now >> </a></span>   
                <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=outz&user=<%=session("user") %>" target="_blank">Send Outz reports manuel now >> </a></span>
                <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=welcom&user=<%=session("user") %>" target="_blank">Send Welcom reports manuel now >> </a></span>
                <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=wwf&user=<%=session("user") %>" target="_blank">Send WWF reports manuel now >> </a></span>
                <br /><span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=adra&user=<%=session("user") %>" target="_blank">Send Adra reports manuel now >> </a></span>         
                     <br /><br />
                </div>
   
            <%end if %>

            <%
            'if session("mid") = 1 then
            %>
             <span style="float:right; padding-right:20px;"><a href="../timereg_net/abonner_manuel_2017.aspx?lto=<%=lto %>&user=<%=session("user") %>" target="_blank">Send reports manuel now >> </a></span>   
            <%'end if %>
           
            <%=abonner_txt_001 %><br />
            <%=abonner_txt_002 %><br /><br />
            <%=abonner_txt_003&" " %><b><%=abonner_txt_004 %></b>, <%=abonner_txt_005&" " %><b><%=abonner_txt_006&" " %><u><%=abonner_txt_007 %></u><%=" "&abonner_txt_008 %></b><%=" "&abonner_txt_009 %><br />&nbsp;


        <FORM method="post" action="abonner.asp?func=abo">
        <table cellspacing=0 cellpadding=0 border=0 width=100%>
       
        <tr>     
	            <td><br /><b>A) <%=abonner_txt_010 %>:</b><br />
        <%=abonner_txt_011 %><br />
        <%strSQLm = "SELECT mnavn, mid, email FROM medarbejdere WHERE mansat = 1 AND email <> '' ORDER BY mnavn"

        'if session("mid") = 1 then
        'Response.write strSQLm
        'end if
        
        %>
        <select style="width:180px; font-size:11px;" size="10" multiple name="FM_abo_mid">
        <%
        oRec2.open strSQLm, oConn, 3
        While Not oRec2.EOF 

        %>
        <option value="<%=oRec2("mid")%>"><%=oRec2("mnavn") %></option>
        <%
        oRec2.movenext
        wend
        oRec2.close
         %>
        
        
        </select>


        <br />
         <input type="hidden" name="FM_abo" value="1" /> 
        </td>

         <td><br /><b>B) <%=abonner_txt_012 %>:</b><br />
        <%=abonner_txt_013 %>:<br />
        <select style="width:230px; font-size:11px;" size="10" name="FM_rapporttype" id="FM_rapporttype">
            <option value="1" SELECTED><%=abonner_txt_014 %></option>
            <option value="2"><%=abonner_txt_015 %> </option>
            <option value="3"><%=abonner_txt_016 %> </option>
            <option value="4">Ressource Forecast </option>
        </select>


       
        </td>


             <td><%=abonner_txt_017 %> <br /><b><%=abonner_txt_018 %>:</b> <br /><b>C)<%=" "&abonner_txt_019 %>:</b><br />
        <%strSQLm = "SELECT id, type FROM medarbejdertyper WHERE id <> 0 ORDER BY type"

        'if session("mid") = 1 then
        'Response.write strSQLm
        'end if
        
        %>
        <select style="width:230px; font-size:11px;" size="10" multiple name="FM_abo_mtyp" id="FM_abo_mtyp">
            <option value="-2" SELECTED><%=abonner_txt_020 %></option>
              <option value="-1"><%=abonner_txt_021 %></option>
            <option value="0"><%=abonner_txt_022 %></option>
        <%
        oRec2.open strSQLm, oConn, 3
        While Not oRec2.EOF 

        %>
        <option value="<%=oRec2("id")%>"><%=oRec2("type") %> (id:<%=oRec2("id") %>)</option>
        <%
        oRec2.movenext
        wend
        oRec2.close
         %>
        
        
        </select>


       
        </td>
                 <td><br /><b><%=abonner_txt_023 %>:</b><br /><b>D) <%=abonner_txt_024 %>:</b><br />
        <%strSQLm = "SELECT id, navn FROM projektgrupper WHERE id <> 0 ORDER BY navn"

        'if session("mid") = 1 then
        'Response.write strSQLm
        'end if
        
        %>
        <select style="width:230px; font-size:11px;" size="10" multiple name="FM_abo_progrp" id="FM_abo_progrp">
            <option value="-1" SELECTED><%=abonner_txt_021 %></option>
            <option value="0"><%=abonner_txt_022 %></option>
        <%
        oRec2.open strSQLm, oConn, 3
        While Not oRec2.EOF 

        %>
        <option value="<%=oRec2("id")%>"><%=oRec2("navn") %> (id:<%=oRec2("id") %>)</option>
        <%
        oRec2.movenext
        wend
        oRec2.close
         %>
        
        
        </select>


       
        </td>


        




        </tr>
        <tr>
	    <td align=right valign=bottom colspan="4" style="padding-right:30px;"><br />


            <input id="Submit5" type="submit" value="<%=abonner_txt_027 %> >>" />

          
	</td></tr>
</table>
</FORM>


        <%  

if request.servervariables("PATH_TRANSLATED") <> "c:\www\timeout_xp\wwwroot\ver2_1\timereg\abonner.asp" then





            'strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;"
            strConnect_admin = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154;Port=3306; uid=to_outzource2;pwd=SKba200473;database=timeout_admin;" 
            Set oConn_admin = Server.CreateObject("ADODB.Connection")
            Set oRec_admin = Server.CreateObject ("ADODB.Recordset")
            oConn_admin.open strConnect_admin

         %>

  <h3><%=abonner_txt_028 %>:</h3> 
 

<table width="100%">
    <tr><td><b><%=abonner_txt_029 %></b></td><td><b><%=abonner_txt_030 %></b></td><td><%=abonner_txt_031 %></td><td><%=abonner_txt_032 %></td><td>&nbsp;</td></tr>

 

          
            <% '** subscribers ***'
            strSQLAbo = "SELECT navn, medid, rapporttype, medarbejdertyper, projektgrupper FROM rapport_abo WHERE lto = '"& lto & "' ORDER BY navn "
            oRec_admin.open strSQLAbo, oConn_admin, 3
            while not oRec_admin.EOF
            
            %>
            <tr>
            <td><%=oRec_admin("navn")%></td><td><%=oRec_admin("rapporttype") %></td>
            <td>
                
                <%if oRec_admin("medarbejdertyper") <> "-1" then %>
                <%=oRec_admin("medarbejdertyper") %>
                <%end if %>
            </td>        
            <td>
                <%if oRec_admin("medarbejdertyper") = "-1" then %>
                <%=oRec_admin("projektgrupper") %>
                <%end if %>
            </td>
                
                <td><a href="abonner.asp?func=abo_slet&FM_abo_mid=<%=oRec_admin("medid") %>" class="slet"><%=abonner_txt_033 %></a></td>
                </tr>
    <%


            oRec_admin.movenext
            wend
           
              
           
            oRec_admin.close
            oConn_admin.close

    
    %>
    </table>
    <%
    
    end if

    %>
      <!-- filter header sLut -->
	</td></tr></table>
	</div>
    

    <%

    end select
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
