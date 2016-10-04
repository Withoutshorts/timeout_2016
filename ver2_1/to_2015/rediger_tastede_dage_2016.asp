

<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/regular/global_func.asp"-->		
<!--#include file="../timereg/inc/convertDate.asp"--> 

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<div class="wrapper">
    <div class="content">

        <%

        call erkmDialog() 

        id = Request("id")

        jobid = 0
        intAktId = 0
        inttknr = 0


	        strSQL = "SELECT tdato, timer, timerkom, tjobnr, tjobnavn, "_
	        &" tknavn, TAktivitetId, Taktivitetnavn, tmnavn, tknr, k.kkundenr, offentlig, timepris, tmnr,"_
	        &" t.editor, tastedato, sttid, sltid, t.valuta, j.id AS jid, bopal FROM timer t "_
	        &" LEFT JOIN job j ON (j.jobnr = t.tjobnr) "_ 
	        &" LEFT JOIN kunder k ON (k.kid = tknr) WHERE Tid =" & id 
	
	        'Response.write strSQL
	        'Response.flush
	
	        oRec.Open strSQL, oConn, 3
	
	        if not oRec.EOF then 
		
		          StrTdato = oRec("Tdato")
		          StrTimer = oRec("Timer")
		          StrTimerkom = oRec("Timerkom")
		          jobnr = oRec("tjobnr")
		          jobid = oRec("jid")
		          StrTjobnavn = oRec("Tjobnavn")
		          StrTknavn = oRec("Tknavn") 
		          StrUdato = "12/12/2007" 
		          intAktId = oRec("TAktivitetId")	
		          strAktnavn = oRec("Taktivitetnavn")
		          intOff = oRec("offentlig")	
		          strUser = oRec("tmnavn")
		          inttknr = oRec("kkundenr")
		          timepris = oRec("timepris")
		          medid = oRec("tmnr")
		          editor = oRec("editor")
		          tastedato = oRec("tastedato") 
		  
		 	         if len(trim(oRec("sttid"))) <> 0 then
				        if left(formatdatetime(oRec("sttid"), 3), 5) <> "00:00" then
				        sttid = left(formatdatetime(oRec("sttid"), 3), 5)
				        else
				        sttid = ""
				        end if
			        else
			        sttid = ""
			        end if
	
			        if len(trim(oRec("sltid"))) <> 0 then
				        if left(formatdatetime(oRec("sltid"), 3), 5) <> "00:00" then
				        sltid = left(formatdatetime(oRec("sltid"), 3), 5)
				        else
				        sltid = ""
				        end if
			        else
			        sltid = ""
			        end if
			
			        intValuta = oRec("valuta")
			
			        bopal = oRec("bopal")
		  
	        end if
	        'now close and clean up
	        oRec.Close
	        Set oRec = Nothing	
        %>

        <div class="container">
            <h3 class="portlet-title"><u><%=tsa_txt_181 %></u></h3>

            <div class="portlet">
                <div class="portlet-body">

                    <div class="row">
                        <div class="col-lg-10"><b><%=tsa_txt_180 %>:</b> &nbsp <%=formatdatetime(date, 1)%></div>
                        <div class="col-lg-1 pull-right"><a href="db_tastede_dage_2006.asp?func=slet&id=<%=id %>"><span style="color:darkred; display: block; text-align: center;" class="fa fa-times"></span></a></div>
                    </div>

                    <br /><br />
                    <div class="well well-white">



                        <div class="row">
                            <div class="col-lg-1"><%=tsa_txt_022 %>:</div>
                            <div class="col-lg-2"><%=left(StrTknavn, 30) %> (<%=inttknr%>)</div>
                        </div>

                        



                    </div>

                </div>
            </div>
        </div>



    </div>
</div>

    

<!--#include file="../inc/regular/footer_inc.asp"-->