<%response.buffer = true%>
<!--#include file="inc/connection/conn_db_inc.asp"-->
<!--#include file="inc/regular/global_func.asp"-->




<%


'**** Søgekriterier AJAX **'
        
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
            Select Case Request.Form("control")
            case "FN_geolocation"

                 '*** Tilføjer longitude og latitude til seneste åbne loginf for dem der har stempelur slået til *********
            
                 call erStempelurOn()
             
                longitude = request("jq_longitude")
                latitude = request("jq_latitude")

                jq_lto = request("jq_lto")
                jq_medid = request("jq_medid")
                
              

                 if cint(stempelurOn) = 1 then
                 strSQLlatestlogin = "SELECT id FROM login_historik WHERE mid =  " & jq_medid & " AND stempelurindstilling > 0 AND lh_latitude_logud = 0 ORDER BY id DESC LIMIT 1 "
			     oRec.open strSQLlatestlogin, oConn, 3
                 if not oRec.EOF then	   

         
	                        strSQLAddLong = "UPDATE login_historik SET "_
	                        &" lh_longitude_logud = "& longitude &", lh_latitude_logud = "& latitude &""_
	                        &" WHERE id = "& oRec("id")
    				   
				           oConn.execute(strSQLAddLong)  
            
              
                   end if
                   oRec.close
              
                 end if


            '*** Bruges på timereg siden ***'
            session("forste") = "j"
            session.abandon

            end select
            Response.end
        end if



if len(trim(session("mid"))) <> 0 then
uid = session("mid")
else
uid = 0
end if


call erStempelurOn()
               
if cint(stempelur_hideloginOn) = 1 then 'skriv ikke login, men åben komme/gå tom ==> Opdater ikke logind historiken ved logud

'select case lcase(lto)
'case "kejd_pb", "kejd_pb2", "fk", "fk_bpm"

logudDone = 1

'case else
else

if len(request("logudDone")) then
logudDone = request("logudDone")
else
logudDone = 0
end if

end if
'end select

if len(trim(request("fromweblogud"))) <> 0 then
fromweblogud = request("fromweblogud")
else
fromweblogud = 0
end if


'*** Opdater login_historik (Stempelur) hvis stempel ur er slået til ****
if cint(logudDone) = 0 then





        select case lto
        case "intranet - local", "cflow"


         '** Er medarbejder med i Aumatisjon Eller Enginnering
        call medariprogrpFn(uid)
        if instr(medariprogrpTxt, "#14#") = 0 AND instr(medariprogrpTxt, "#16#") = 0 then

        if cint(fromweblogud) <> 1 then
    
             '*** Indlæser på aktiviteter **' DER SKAL VÆRE VALGT EN AKT
            if len(trim(request("FM_aktivitetid"))) <> 0 AND trim(request("FM_aktivitetid")) <> ".." then
            aktivitetid = trim(request("FM_aktivitetid"))
            else
            aktivitetid = "0"
            end if

            'response.Write "aktivitetid: " & aktivitetid
            'response.flush

            if cstr(aktivitetid) = "0" then

            %>
            <!--#include file="inc/regular/header_lysblaa_2015_inc.asp"-->
             <div class="container" style="width:100%; height:100%">
                 <br /><br /><br /><br /><br /><br /><br /><br />
                <div class="portlet">
                    <div class="portlet-body">

                            <div class="row">
                                    <div class="col-lg-12" style="text-align:center"><h3>Feil!<br /><br />
                                        <span style="color:red;">Før du kan stemple ut, skal der vælges et projekt og en aktivitet.</span>
                                        <br /><br />
                                        <a href="#" onclick="Javascript:history.back();"><< Tilbake</a>
                                        </h3>

                                    </div>
                                </div>

                        </div>
                    </div>
                 </div>

            <%
            Response.end


            end if

        end if 

        end if 'instr(medariprogrpTxt, "#14#") = 0 AND instr(medariprogrpTxt, "#16#") = 0
        end select




'***********************************************************
'*** Fordel stempleurs timer ud på aktiviteter
'*** H&L timer CFLOW
'*** Projekttimer
'***********************************************************
 dothis = 0
 LogudDateTime = year(now)&"/"& month(now)&"/"&day(now)&" "& datepart("h", now) &":"& datepart("n", now) &":"& datepart("s", now) 
 call fordelStempelurstimer(uid, lto, dothis, logudDone, LogudDateTime) 


end if 'if cint(logudDone) = 0 then

'***********************************************
'**** END LOGOUT STEMPELUR
'***********************************************






'response.flush
'Response.Write "<br><br><br><br><br><br><br><br><br><br><br><br><br>&nbsp;<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"

              
				
select case lto 
case "tec", "xintranet - local", "dencker", "epi2017" 

    %>
        

            <!--<!--#include file="inc/regular/header_login_inc.asp"

            <link href='//fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css'>
            <link href="../inc/menu/css/chronograph_01.less" rel="stylesheet/less" type="text/css" />

            -->

<!--#include file="inc/regular/header_lysblaa_2015_inc.asp"-->
<script src="sesaba_jav_2017_0329.js" type="text/javascript"></script>
       

<body topmargin="0" leftmargin="0" class="login">
<%call takforidag() %>
</body></html>



    <%

case "intranet - local", "cflow"

        'if len(trim(request("indlaspaajob"))) <> 0 then
        'indlaspaajob = request("indlaspaajob")
        'else
        'indlaspaajob = 0
        'end if

        if request("fromstempelur") = "1" OR request("fromweblogud") = "1"  then 'fra PC / WEB

        %>

        <!--#include file="inc/regular/header_lysblaa_2015_inc.asp"-->
        <script src="sesaba_jav_2017_0329.js" type="text/javascript"></script>
       

        <body topmargin="0" leftmargin="0" class="login">
        <%call takforidag() %>
        </body></html>


        <%


        else 'Fra terminalk


                if cint(indlaspaajob) = 1 then
                call meStamdata(session("mid"))
        
                    'if session("mid") <> 1 then
                    response.redirect "to_2015/monitor.asp?func=scan&RFID_field="&meNr&"&skiftjob=1"
                    'else
                    'Response.write "to_2015/monitor.asp?func=scan&RFID_field="&meNr
                    'end if

                else
                response.redirect "to_2015/monitor.asp?func=startside"
                end if


        end if

case else


'*** Bruges på timereg siden ***'
session("forste") = "j"
session.abandon
'Response.Write("<script language=""JavaScript"">window.opener.close();</script>")

call browsertype()
if browstype_client = "ch" then
        %>
         <br /><br />Tak for dine registreringer Du kan nu lukke browseren. 
          <br /><br />Du bruger Chrome. Derfor lukker dette vindue ikke selv.

        <%
else
Response.Write("<script language=""JavaScript"">window.open('', '_self', '');</script>")
Response.Write("<script language=""JavaScript"">window.close();</script>")
end if

end select












function takforidag()


session("forste") = "j"
session.abandon

            %>


<form>
                
                        <input type="hidden" id="jq_lto" name="" value="<%=lto %>"/>
                        <input type="hidden" id="jq_medid" name="" value="<%=uid %>"/>
                        <input type="hidden" id="jq_longitude" name="" value="0"/>
                        <input type="hidden" id="jq_latitude" name="" value="0"/>
                        <input type="hidden" id="jq_stempeluron" name="" value="<%=stempelurOn %>"/>  

</form>



        <div class="wrapper">
 <div class="content">   

        <div class="container">
            <div class="portlet">
                   
                  <h3 class="portlet-title"><u></u> </h3>
                  

                <div class="portlet-body">
                     <!--<div class="portlet-title">
                          <table style="width:100%;">
                              <tr>
                                  <td><h3 style="color:black"><%=medarb_navn %></h3></td>
                                  <td style="text-align:right"><img style="width:15%; margin-bottom:2px; margin-left:55px" src="img/cflow_logo.png" /></td>
                              </tr>
                          </table>
                        
                        </div>-->

                  
                    <br />
        <img src="to_2015/img/outzource_logo_4c.jpg" width="200" />

                       <div class="well well-white">
                      
                        
                                                    <div class="row">
                                                        <div class="col-lg-6">
                    <table cellspacing="0" cellpadding="10" border="0" bgcolor="#ffffff" width=100%>
	
	                    <tr>
	                    <td colspan=2>
                        <b>Tak for idag</b><br /><br />	
                        Du er nu logget ud af TimeOut.<br />
                        Alle dine registreringer er modtaget og sessionen er afsluttet.<br /><br />
        
     
                    Log på timeOut igen her:<br />
                    <a href="https://timeout.cloud/<%=request.Cookies("loginlto")("lto")%>">https://timeout.cloud/<%=request.Cookies("loginlto")("lto")%></a>
               
                     <br /><br />
        
                            &nbsp;
	                    </td>
	                    </tr>
	
                    </table>
                    </div>
                    </div>

                       </div>

                    
                </div>
</div></div></div></div>


            <%

end function
		
%>
