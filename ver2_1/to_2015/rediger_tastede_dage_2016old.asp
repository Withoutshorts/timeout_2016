

<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/regular/global_func.asp"-->		
<!--#include file="../timereg/inc/convertDate.asp"--> 

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<div class="wrapper">
    <div class="content">



        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Rediger timeregistrering</u></h3>

                <div class="portlet-body">
                    <div class="well well-white">
                        

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
                        <div class="row">
                            <div class="col-lg-12"><%=tsa_txt_180 %>:&nbsp;<%=formatdatetime(date, 1)%></div>
                        </div>

                        <div class="row">
                            <div class="col-lg-6"><%=tsa_txt_182 %>: <b><%=editor%></b> <b><%=formatdatetime(tastedato, 1)%></b></div>
                        </div>

                        <br />

                        <form action="db_tastede_dage_2006.asp" method="POST">
                            <input type="Hidden" name="id" value="<%=id%>">
                            <input type="Hidden" name="medid" value="<%=medid%>">
                            <input type="Hidden" name="jobnr" value="<%=jobnr%>">

                            <div class="row">
                                <div class="col-lg-2"><%=tsa_txt_022 %>:</div>
                                <div class="col-lg-3"><%=left(StrTknavn, 30) %> (<%=inttknr%>)</div>
                            </div>

                            <div class="row">
                                <div class="col-lg-2"><%=tsa_txt_067 %>:</div>
                                <div class="col-lg-3"><%=left(StrTjobnavn, 30) %> (<%= jobnr %>)</div>
                            </div>

                            <div class="row">
                                <div class="col-lg-2">Aktivitet:</div>
                                <div class="col-lg-3">
                                    <%
        
                                    if session("rettigheder") = 1 then 
                                    strSQLakt = "SELECT navn, id FROM aktiviteter WHERE job = "& jobid &" ORDER BY fase, sortorder, navn"
                                    'Response.write "<br><br><br><br><br><br><br><br><br><br>"& strSQLakt
                                    'Response.flush
		                            oRec3.open strSQLakt, oConn, 3
		
                                        %>
                                    <select id="Select2" name="intAktId" class="form-control input-small">
                                    <%
		
      
		                            while not oRec3.EOF 
		
		                            if cdbl(intAktId) = oRec3("id") then
	                                aSel = "SELECTED"
                                    else
		                            aSel = ""
		                            end if%>
		                                <option <%=aSel %> value="<%=oRec3("id") %>"><%=oRec3("navn") %></option>
		                            <%
		
		                            oRec3.movenext
		                            wend 
		                            oRec3.close
		                                %>
		 
		                                </select>

                                        <%else 
          
                                        strSQLakt = "SELECT navn, id FROM aktiviteter WHERE id = "& intAktId
                                        oRec3.open strSQLakt, oConn, 3
                                    aNavn = "(ikke fundet)"
		                            if not oRec3.EOF then 
	                                aNavn = oRec3("navn")
	                                end if 
		                            oRec3.close
          
                                        Response.write aNavn

                                        %>
                                        <input type="hidden" name="intAktId" value="<%=intAktId %>" />
                                <%end if %>
                                </div>
                            </div>

                            <div class="row">
                                <div class="col-lg-2"><%=tsa_txt_101 %>:</div>
                                <div class="col-lg-3"><%= strUser %></div>
                            </div>

                            <div class="row">
                                <div class="col-lg-2"><%=tsa_txt_183 %>:</div>
                                <div class="col-lg-1">
                                    <select name="dag" class="form-control input-small">
		                            <option value="<%=strDag%>"><%=strDag%></option> 
		                            <option value="1">1</option>
	   	                            <option value="2">2</option>
	   	                            <option value="3">3</option>
	   	                            <option value="4">4</option>
	   	                            <option value="5">5</option>
	   	                            <option value="6">6</option>
	   	                            <option value="7">7</option>
	   	                            <option value="8">8</option>
	   	                            <option value="9">9</option>
	   	                            <option value="10">10</option>
	   	                            <option value="11">11</option>
	   	                            <option value="12">12</option>
	   	                            <option value="13">13</option>
	   	                            <option value="14">14</option>
	   	                            <option value="15">15</option>
	   	                            <option value="16">16</option>
	   	                            <option value="17">17</option>
	   	                            <option value="18">18</option>
	   	                            <option value="19">19</option>
	   	                            <option value="20">20</option>
	   	                            <option value="21">21</option>
	   	                            <option value="22">22</option>
	   	                            <option value="23">23</option>
	   	                            <option value="24">24</option>
	   	                            <option value="25">25</option>
	   	                            <option value="26">26</option>
	   	                            <option value="27">27</option>
	   	                            <option value="28">28</option>
	   	                            <option value="29">29</option>
	   	                            <option value="30">30</option>
		                            <option value="31">31</option></select>
                                </div>
                                <div class="col-lg-1">
                                    <select name="mrd" class="form-control input-small">
		                            <option value="<%=strMrd%>"><%=strMrdNavn%></option>
		                            <option value="1">jan</option>
	   	                            <option value="2">feb</option>
	   	                            <option value="3">mar</option>
	   	                            <option value="4">apr</option>
	   	                            <option value="5">maj</option>
	   	                            <option value="6">jun</option>
	   	                            <option value="7">jul</option>
	   	                            <option value="8">aug</option>
	   	                            <option value="9">sep</option>
	   	                            <option value="10">okt</option>
	   	                            <option value="11">nov</option>
	   	                            <option value="12">dec</option></select>
                                </div>
                                <div class="col-lg-1">
                                    <select name="aar" class="form-control input-small">
		                            <option value="<%=strAar%>">
		                            <%if id <> 0 then%>
		                            20<%=strAar%>
		                            <%else%>
		                            <%=strAar%>
		                            <%end if%></option>
		                            <option value="02">2002</option>
		                            <option value="03">2003</option>
	   	                            <option value="04">2004</option>
	   	                            <option value="05">2005</option>
		                            <option value="06">2006</option>
		                            <option value="07">2007</option>
		                            <option value="08">2008</option>
		                            <option value="09">2009</option>
		                            <option value="10">2010</option>
		                            <option value="11">2011</option>
		                            <option value="12">2012</option>
		                            <option value="13">2013</option>
		                            <option value="14">2014</option>
                                    <option value="15">2015</option>
                                    <option value="16">2016</option>
                                    <option value="17">2017</option>
                                    <option value="18">2018</option>
                                    <option value="19">2019</option></select>
                                </div>
                            </div>

                            <div class="row">
                                <div class="col-lg-2"><%=tsa_txt_070 %>:</div>
                                <div class="col-lg-1"><input type="Text" name="Timer" Value="<%=StrTimer%>" class="form-control input-small"></div>
                                <div class="col-lg-3" style="text-align:right"><%=tsa_txt_184 %>:</div>
                                <div class="col-lg-1"><input type="text" name="FM_sttid" value="<%=sttid%>" size=2 class="form-control input-small"></div>
                                <div class="col-lg-1" style="text-align:center">-</div>
                                <div class="col-lg-1"><input type="text" size=2 name="FM_sltid" value="<%=sltid%>" class="form-control input-small"></div>                      
                            </div>

                            <%if session("rettigheder") <= 2 OR session("rettigheder") = 6 then %>
                            <div class="row">
                                <div class="col-lg-2"><%=tsa_txt_186 %>:</div>
                                <div class="col-lg-1"><input type="Text" name="timepris" Value="<%=timepris%>"  class="form-control input-small"></div>
                                <div class="col-lg-2">
                                    <select name="FM_valuta" id="Select1" class="form-control input-small">
		                            <option value="<%=intValuta %>"><%=tsa_txt_229 %></option>
		                            <%
		                            strSQL3 = "SELECT id, valutakode, grundvaluta FROM valutaer ORDER BY valutakode"
    		                        oRec3.open strSQL3, oConn, 3 
		                            while not oRec3.EOF 
    		
		                            if cint(oRec3("id")) = cint(intValuta) then
		                            valGrpCHK = "SELECTED"
		                            else
		                            valGrpCHK = ""
		                            end if
    		
    		
		                            %>
		                            <option value="<%=oRec3("id")%>" <%=valGrpCHK %>><%=oRec3("valutakode")%></option>
		                            <%
		                            oRec3.movenext
		                            wend
		                            oRec3.close %>
		                            </select>
                                </div>
                            </div>
                            <div class="row">
                                <div class="col-lg-2">&nbsp</div>
                                <div class="col-lg-5"><input type="checkbox" value="1" name="FM_opdater_timepriser" /> Opdater timepris på denne aktivitet for valgte medarbejder</div>
                            </div>
                            <div class="row">
                                <div class="col-lg-2">&nbsp</div>
                                <div class="col-lg-2">Opdater timepriser fra d.</div>
                                <div class="col-lg-1"><input type="text" name="FM_opdatertpfra" value="<%=formatdatetime(now,2) %>" class="form-control input-small"/></div>
                                <div class="col-lg-5">til dd. (også lukkede uger og hvis der foreligger faktura)</div>
                            </div>
                            <%else %>
                            <input type="hidden" name="timepris" value="<%=timepris%>" />
                            <input type="hidden" name="FM_valuta" value="<%=intValuta %>" />
                            <%end if %>

                            <br />

                            <div class="row">
                                <div class="col-lg-2"><%=tsa_txt_051 %>:</div>
                                <div class="col-lg-5"><textarea name="Timerkom" cols="60" rows="5" class="form-control input-small"><%=StrTimerkom%></textarea></div>
                            </div>

                            <br />

                            <div class="row">
                                <div class="col-lg-2"><%=tsa_txt_053 %>:</div>
                                <div class="col-lg-3"><select name="FM_off" id="FM_off" class="form-control input-small">
	                                <%
	                                if intOff = 0 then
	                                nejsel = "SELECTED"
	                                jasel = ""
	                                else
	                                nejsel = ""
	                                jasel = "SELECTED"
	                                end if
	                                %>
	                                <option value="0" <%=nejsel%>><%=tsa_txt_054 %></option>
	                                <option value="1" <%=jasel%>><%=tsa_txt_055 %></option>
	                                </select>
                                </div>
                            </div>

                            <%if kmDialogOnOff = 1 then %>
                                <div class="row">
                                    <div class="col-lg-2">&nbsp</div>
                                    <div class="col-lg-5">
                                        <%if cint(bopal) = 1 then 
                                        bpCHK = "CHECKED"
                                        else
                                        bpCHK = ""
                                        end if%>
                                        <input id="Checkbox1" name="FM_bopal" value="1" type="checkbox" <%=bpCHK %> /> <%=tsa_txt_287 %> (<%=tsa_txt_288%>)
                                    </div>
                                </div>
                            <%else %>
                                    <input id="name="FM_bopal"" name="FM_bopal" value="0" type="hidden" />
                            <%end if %>

                            <div class="row">
                                <div class="col-lg-10">&nbsp</div>
                                <div class="col-lg-2"><button type="submit" class="btn btn-success btn-sm"><b>Opdatér</b></button></div>
                            </div>

                        </form>
                    </div>
                </div>


            </div>
        </div>



    </div>
</div>

    

<!--#include file="../inc/regular/footer_inc.asp"-->