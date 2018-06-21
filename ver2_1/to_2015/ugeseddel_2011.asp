<%response.buffer = true%>


<!--#include file="../inc/connection/conn_db_inc.asp"-->



<%
 '**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_gemjobc"

        case "FN_gemaktc"

        case "FN_tjktimer_forecast"
                '** FORECAST ALERT 
         

                aktid = request("aktid")
                timerTastet = request("timer_tastet")
                usemrn = request("treg_usemrn")
                ibudgetaar = request("ibudgetaar")
                ibudgetmd = request("ibudgetaarMd")  
                aar = request("ibudgetUseAar")
                md = request("ibudgetUseMd")

               

                call ressourcefc_tjk(ibudgetaar, ibudgetmd, aar, md, usemrn, aktid, timerTastet)

                
             
                response.write feltTxtValFc
               


        case "FN_sogjobogkunde"


               call selectJobogKunde_jq()


        
          case "FN_sogakt"

               
              call selectAkt_jq



        end select
        response.end
        end if




tloadA = now
    
   %>

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../timereg/inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/treg_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<% 

if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
	
	else
	

        'SET LOCALE for denne side
        select case sprog
        case 1
        Session.LCID = 1030
        case 2
        Session.LCID = 2057
        case 3
        Session.LCID = 1053
        case 4
        Session.LCID = 1044
        case 5
        Session.LCID = 1034
        case 6
        Session.LCID = 1031
        case 7
        Session.LCID = 1036
        case else
        Session.LCID = 1030
        end select


	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	thisfile = "ugeseddel_2011.asp"
	media = request("media")

   
    touchMonitor = 0
    


    if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> 0 then
    nomenu = 1
    else
    nomenu = 0
    end if

    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = session("mid")
    end if
    


            
         

    level = session("rettigheder")
 
    rdir = request("rdir")
    select case rdir
    case "logindhist"
    rdirfile = "logindhist_2011.asp"
    case "godkenduge"
    rdirfile = "../timereg/godkenduge.asp"
    case else
	rdirfile = "ugeseddel_2011.asp"
	end select
    

    if len(trim(request("medarbsel_form"))) <> 0 then

        'response.write "HER"
        'response.end

            if len(trim(request("FM_visAlleMedarb"))) <> 0 then
            visAlleMedarbCHK = "CHECKED"
            visAlleMedarb = 1
            else
            visAlleMedarbCHK = ""
            visAlleMedarb = 0
            end if

            if len(trim(request("FM_visAlleMedarb_pas"))) <> 0 then
            visAlleMedarb_pasCHK = "CHECKED"
            visAlleMedarb_pas = 1
            else
            visAlleMedarb_pasCHK = ""
            visAlleMedarb_pas = 0
            end if

    else

                    if request.cookies("tsa")("visAlleMedarb") = "1" then
                    visAlleMedarbCHK = "CHECKED"
                    visAlleMedarb = 1
                    else
                    visAlleMedarbCHK = ""
                    visAlleMedarb = 0
                    'usemrn = session("mid")
                    end if

                    if request.cookies("tsa")("visAlleMedarb_pas") = "1" then
                    visAlleMedarb_pasCHK = "CHECKED"
                    visAlleMedarb_pas = 1
                    else
                    visAlleMedarb_pasCHK = ""
                    visAlleMedarb_pas = 0
                    end if

                    
    end if
    response.cookies("tsa")("visAlleMedarb") = visAlleMedarb
    response.cookies("tsa")("visAlleMedarb_pas") = visAlleMedarb_pas



    
    call ersmileyaktiv()
	
    call medarb_teamlederfor
    
    '** Select eller søgeboks
    call mobil_week_reg_dd_fn()
    'mobil_week_reg_job_dd = mobil_week_reg_job_dd
    'mobil_week_reg_akt_dd = 1
    'mobil_week_reg_akt_dd = mobil_week_reg_akt_dd    
    
        
    varTjDatoUS_man = request("varTjDatoUS_man")

    'if len(trim(request("varTjDatoUS_son"))) <> 0 then
	'varTjDatoUS_son = request("varTjDatoUS_son")
    'else
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)
    'end if

    'Response.Write varTjDatoUS_man
    'Response.end

    datoMan = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
    datoSon = day(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& year(varTjDatoUS_son)

    next_varTjDatoUS_man = dateadd("d", 7, varTjDatoUS_man)
	next_varTjDatoUS_son = dateadd("d", 7, varTjDatoUS_son)
    next_varTjDatoUS_man = year(next_varTjDatoUS_man) &"/"& month(next_varTjDatoUS_man) &"/"& day(next_varTjDatoUS_man)
	next_varTjDatoUS_son = year(next_varTjDatoUS_son) &"/"& month(next_varTjDatoUS_son) &"/"& day(next_varTjDatoUS_son)


    prev_varTjDatoUS_man = dateadd("d", -7, varTjDatoUS_man)
	prev_varTjDatoUS_son = dateadd("d", -7, varTjDatoUS_son)
    prev_varTjDatoUS_man = year(prev_varTjDatoUS_man) &"/"& month(prev_varTjDatoUS_man) &"/"& day(prev_varTjDatoUS_man)
	prev_varTjDatoUS_son = year(prev_varTjDatoUS_son) &"/"& month(prev_varTjDatoUS_son) &"/"& day(prev_varTjDatoUS_son)





    call positiv_aktivering_akt_fn()
    call smileyAfslutSettings()

	'Response.Write "media " & media
	
	
	'*** Sætter lokal dato/kr format. *****
   'call lcid_sprog(session("mid"))
   'Session.LCID = lcid_sprog_Val
	lcid_sprog_val = Session.LCID

    select case func 
	case "-"

    case "xxxopdaterstatus" 'NOT IN USE 20170927


        
        'response.write "request(ids)" & request("ids")
        
          ujid = split(request("ids"), ",")

                if len(trim(request("FM_godkendt"))) <> 0 then
                uGodkendt = request("FM_godkendt")
                else
                uGodkendt = 0
                end if
	
    for u = 0 to UBOUND(ujid)
	
	            editor = session("user")


                'if len(trim(request("FM_godkendt_"& trim(ujid(u))))) <> 0 then
                'uGodkendt = request("FM_godkendt_"& trim(ujid(u)))
                'else
                'uGodkendt = 0
                'end if
              

				strSQL = "UPDATE timer SET godkendtstatus = "& uGodkendt &", "_
				&"godkendtstatusaf = '"& editor &"' WHERE tid = " & ujid(u) & " AND overfort = 0 "
				
			    'Response.write strSQL &"<br>"
				
				oConn.execute(strSQL)
				
	next

       'response.end

    'Response.Redirect "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son&"&nomenu="&nomenu

    Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	
	case "godkendugeseddel"
	
	'*** Godkender ugeseddel ***'



	
	    varTjDatoUS_manSQL = year(varTjDatoUS_man)&"/"&month(varTjDatoUS_man)&"/"&day(varTjDatoUS_man)
        varTjDatoUS_sonSQL = year(varTjDatoUS_son)&"/"&month(varTjDatoUS_son)&"/"&day(varTjDatoUS_son)
	    
	   
        
         '**** Godkender timer i dbne uge der bnliver godkendt
        call godkenderTimeriUge(usemrn, varTjDatoUS_manSQL, varTjDatoUS_sonSQL, SmiWeekOrMonth)
       

        '*** Godkend uge status ****'
        call godekendugeseddel(thisfile, session("mid"), usemrn, varTjDatoUS_man)
	    
	
	
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	

    case "afvisugeseddel"
	
	'*** Afviser ugeseddel ***'
    if len(trim(request("FM_afvis_grund"))) <> 0 then
    txt = replace(request("FM_afvis_grund"), "'", "")
    else
    txt = ""
    end if

	call afviseugeseddel(thisfile, session("mid"), usemrn, varTjDatoUS_man, varTjDatoUS_son, txt)
	    
	
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	
    
    
    case "adviserugeafslutning"

    call afslutugereminder(thisfile, session("mid"), usemrn, varTjDatoUS_man, varTjDatoUS_son, txt)

  
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	

    case "slet_tip"

        %>
        	
        <%

        id = request("id")

        slttxt = ""& week_txt_001 &"<br> "& week_txt_002 &""
        slturl = rdirfile & "?func=slet_tip_ok&id="&id&"&usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son

       call sltque(slturl,slttxt,slturlalt,slttxtalt,110,90)

    

    case "slet_tip_ok"

        id = request("id")

        sqlInTimerImpTemp = "DELETE FROM timer_import_temp WHERE id = "& id
        oConn.Execute(sqlInTimerImpTemp)

    Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son

    case else


   

     if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
    <!--#include file="../inc/regular/top_menu_mobile.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	response.end
    end if

   
    '*** Vis kun aktiviteter med forecast på
    call aktBudgettjkOn_fn()
    '*** Skal akt lukkes for timereg. hvis forecast budget er overskrddet..?
    '** MAKS budget / Forecast incl. peridoe afgrænsning
    call akt_maksbudget_treg_fn()

    call akttyper2009(2)
	

    if len(trim(request("FM_datoer"))) <> 0 then 'redirect fra timereg_akt_2006 HUSK Dato
        tregDato = day(request("FM_datoer")) &"/"& month(request("FM_datoer")) &"/"& year(request("FM_datoer"))

        ddDato_ugedag = day(tregDato) &"/"& month(tregDato) &"/"& year(tregDato)
        ddDato_ugedag_w = datepart("w", ddDato_ugedag, 2, 2)
            
        varTjDatoUS_man = dateAdd("d", -(ddDato_ugedag_w-1), ddDato_ugedag)

    else
        
        if (datepart("ww", varTjDatoUS_man, 2, 2) = datepart("ww", day(now) &"/"& month(now) &"/"& year(now), 2, 2)) then  'Indeværende uge vises DD
        tregDato = day(now) &"/"& month(now) &"/"& year(now)
        else
        tregDato = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
        end if

    end if
	
	
	if media <> "export" then
	
	
	if media <> "print" then
	
	
	%>

	
    <SCRIPT src="js/ugeseddel_2011_jav6.js"></script>
    <SCRIPT src="../timereg/inc/smiley_jav.js"></script>

    <%call browsertype() 
    if (browstype_client <> "ip") then
        
        call menu_2014()

    else

        ddDato_ugedag = day(now) &"/"& month(now) &"/"& year(now)
        ddDato_ugedag_w = datepart("w", ddDato_ugedag, 2, 2)
            
        varTjDatoUS_man_tt = dateAdd("d", -(ddDato_ugedag_w-1), ddDato_ugedag)
        
       call mobile_header 
        

    end if
    %>    


   <%      
         
    
    else 
	

    end if 'print%>
	
	<%end if 'eksport
        
        


   
        %>

 

 <script type="text/javascript" src="js/plugins/flot/jquery.flot.js"></script>
    <script type="text/javascript" src="js/demos/flot/stacked-vertical_ugetotal.js"></script>
   
    <style type="text/css">
         
                .blink {
  animation: blink-animation 1s steps(5, start) infinite;
  -webkit-animation: blink-animation 1s steps(5, start) infinite;
}
@keyframes blink-animation {
  to {
    visibility: hidden;
  }
}
@-webkit-keyframes blink-animation {
  to {
    visibility: hidden;
  }
}
                    </style>
	
	<!-------------------------------Sideindhold------------------------------------->

   
 


   <%if (browstype_client <> "ip") then %>
<div class="wrapper">
 <div class="content">   
     <%end if %>

 
     
     <div class="container">
    
      <div class="portlet">
        <% if browstype_client <> "ip" then %>
          
          <h3 class="portlet-title"><u><%=week_txt_012 %></u><!-- ugeseddel --></h3>
         
        <%end if %>

        <div class="portlet-body">
            
      
        <%   
                                
            timerdenneuge_dothis = 0
            call timerDenneUge(usemrn, lto, varTjDatoUS_man, aty_sql_realhours, timerdenneuge_dothis, SmiWeekOrMonth)

                ugetotal = manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer
                manTimer = replace(manTimer, ",",".")
                tirTimer = replace(tirTimer, ",",".")
                onsTimer = replace(onsTimer, ",",".")
                torTimer = replace(torTimer, ",",".")
                lorTimer = replace(lorTimer, ",",".")
                sonTimer = replace(sonTimer, ",",".")

                call normtimerPer(usemrn, varTjDatoUS_man, 6, 0)

            norm_ugetotal = ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon

            if norm_ugetotal <> 0 then
            'weeknormpro = ugetotal / norm_ugetotal * 100
            weeknormpro = ugetotal / norm_ugetotal * 100
            weeknormpro = replace(weeknormpro, ",",".")
            else
            weeknormpro = 0
            end if

            %>

            <input type="hidden" id="dagman" value="<%=manTimer %>" />
            <input type="hidden" id="dagtir" value="<%=tirTimer %>" />
            <input type="hidden" id="dagons" value="<%=onsTimer %>" />
            <input type="hidden" id="dagtor" value="<%=torTimer %>" />
            <input type="hidden" id="dagfre" value="<%=freTimer %>" />
            <input type="hidden" id="daglor" value="<%=lorTimer %>" />
            <input type="hidden" id="dagson" value="<%=sonTimer %>" />

        
                

            
            <%
          

           
    if media <> "print" AND len(trim(strSQLmids)) > 0 then 'Hvis man er level 1 eller teamleder vil len(trim(strSQLmids)) ALTID VÆRE > 16 %>
    <%
        
        if browstype_client <> "ip" then
    %>
    <%
    
    %>
    <div class="well well-white">
    
        
        <%select case lto 
        case "bf"
            
        case else%>

        <%if cint(stempelurOn) = 1 then 
            wdth = 225
        else
            wdth = 120
        end if

        %>
          
           
            <div style="position:relative; background-color:#ffffff; border:1px #cccccc solid; border-bottom:0; padding:4px; width:<%=wdth%>px; top:-70px; left:880px; z-index:0; font-size:11px;">
           
            <a href="../timereg/<%=lnkTimeregside%>" class="vmenu"><%=replace(tsa_txt_031, " ", "")%> >></a>
          

            <%if cint(stempelurOn) = 1 then
                
                if lto = "cflow" OR lto = "dencker" OR lto = "intranet - local" then
                %>
                &nbsp;|&nbsp;<a href="<%=lnkLogind%>" class="vmenu"><%=tsa_txt_340 %> >></a>
                <%else%>
                &nbsp;|&nbsp;<a href="../timereg/<%=lnkLogind%>" class="vmenu"><%=tsa_txt_340 %> >></a>
                <%end if%>
            <%end if%>

           
            </div>
            

        <%end select %>
      
        <%select case lto
         case "xepi2017", "xintranet - local"

            thisEpiDay = datepart("w", now, 2,2)  
            visGuide = 0
            strSQLvisguide = "SELECt visguide FROM medarbejdere WHERE mid = "& session("mid")
            oRec.open strSQLvisguide, oConn, 3
            if not oRec.EOF then
             visGuide = oRec("visguide")
            end if
            oRec.close

            'session("show_skiftStSide_vist") <> "11" AND
            if (thisEpiDay = 1 OR thisEpiDay = 2) AND visGuide <> 99 then
            show_skiftStSide = 1
            'session("show_skiftStSide_vist") = 1
            else
            show_skiftStSide = 0
            end if

         case else
            show_skiftStSide = 0
         end select 
            
            
         if cint(show_skiftStSide) = 1 then%>
        <form action="../timereg/timereg_akt_2006.asp">
            <input type="hidden" name="func" value="epitregpage" />
            <div id="d_epinote" style="position:absolute; top:70px; left:850px; width:400px; border:10px #999999 solid; z-index:10000000; background-color:#ffffff; padding:20px;">
                <span style="float:right;"><a href="#" id="a_epinote" style="color:red;">[X]</a></span>
                <h4>Welcome to Epinion version 2017</h4> 
                We have put together all Epinion versions Of TimeOut in one solution: Epi 2017.<br /><br />
                You will be able to search for all your projects across DK, NO and UK in this database. <br /><br />
                If You wish to keep your old timecording page and use it as your default, you can choose it heere. 
                <br /><input type="checkbox" name="FM_oldtreg" value="1" /> I want the old timecording page.<br /> 
                <input type="checkbox" name="FM_dontshowvisguide" value="1" /> Don't show this message again. <input type="submit" class="btn btn-sm btn-default" value=" >> " /><br /><br>
                You can always change it in your userprofile or select it form the main menu.<br /><br />
                
                If you are missing a project don't hesitate to write to Epinion Finance.<br /><br />
                If You experience any troubles with the new version send an email and a screenshot to support@outzource.dk  
                


            </div>

        </form>
        <%end if%>

        <form id="filterkri" method="post" action="ugeseddel_2011.asp">
            <input type="hidden" name="FM_sesMid" id="FM_sesMid" value="<%=session("mid") %>">
            <input type="hidden" name="medarbsel_form" id="medarbsel_form" value="1">
            <input type="hidden" name="nomenu" id="nomenu" value="<%=nomenu %>">
            <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
           
            
            <div class="row">
                <div class="col-lg-6">
                      <table style="font-size:100%; color:black">
                    
                    <tr>
                        <!--<td style="padding-right:5px; vertical-align:text-top; width:80px;"><input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> <%=tsa_txt_357%> </td>-->
                        <td style="padding:0px 0px 4px 12px;">
                            <input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> <%=tsa_txt_357%> &nbsp;
                            <input type="CHECKBOX" name="FM_visallemedarb_pas" id="FM_visallemedarb_pas" value="1" <%=visAlleMedarb_pasCHK %> onclick="submit();" /> <%=medarb_txt_031%><br />
                    
                   
               
                    <% 
				    call medarb_vaelgandre
                    %>
                            </td>
                        <td>&nbsp;</td>
                        </tr>
                        </table>
                </div>
            </div>
          
                  
        </form>

	
       
	
	<%
   
       end if 'media
	
        'if datePart("d", now, 2,2) < 4 OR datePart("d", now, 2,2) > 29 then 

	   if media <> "print" then
        %> 
        <%
            
        if browstype_client <> "ip" then
        %>

        <%'** TOUCH MoNITOR BRUGES IKKE MERE PÅ UGESEDDEL - HAR FEÅT SINE GEN SIDE 20180306
          '** KODE GEMMES for at huske og indtil ny side er testet
        if touchMonitor = 1 then
        txtClass = "large"
        inputHeight = "150%"
        inputFont = "150%"
        paddingTop = "30px"
        'response.Write "<br><br><br>"
        rdir_timereg = "touchMonitor"
        else
        txtClass = "small"
        inputWidth = "100%"
        inputFont = "100%"
        paddingTop = "10px"
        rdir_timereg = "ugeseddel_2011"
        end if
        %>

        

            <div class="row"> 

                 
     
                      <%
                    mgnRight = 250    
                    wdtTbl = 60


                              'select case lto 
                              'case "cflow", "intranet - local"
                              'if cint(stempelurOn) = 1 then
                              '  if session("mid") = 1 then
                              '  colspL = 4
                              '  colspC = 4
                              '  colspR = 4
                              '  else
                              '  colspL = 7
                              '  colspC = 0
                              '  colspR = 5
                              '  end if
                             'else
                                colspL = 7
                                colspC = 0
                                colspR = 5
                             'end if

                    %>

           
                   <div class="col-lg-<%=colspL%>"> 
                   <form id="container" action="../timereg/timereg_akt_2006.asp?func=db&rdir=<%=rdir_timereg %>" method="post">
      
                  

                    <table style="font-size:100%; display:inline-table; margin-right:<%=mgnRight%>px; width:<%=wdtTbl%>%;">
                    
                    <%if touchMonitor <> 1 then %>
                    <tr>
                        <!--<td style="padding-right:5px; vertical-align:text-top; width:80px;"><b><%=tsa_txt_183%>:</b></td>-->
                        <td style="padding-left:10px">
                            <input type="hidden" id="Hidden5" name="year" value="<%=year(now) %>"/>
                               
                              <%if lto <> "tbg" OR level = 1 then 
                                    inputDatoTXT = ""
                                    inputDatoCLS = "date"
                              else 
                                    inputDatoTXT = "readonly"
                                    inputDatoCLS = ""
                                    tregDato = day(now) &"/"& month(now) &"/"& year(now)
                              end if%>

                              <div class='input-group <%=inputDatoCLS %>'>
                                      <input type="text" style="width:300px;" class="form-control input-small" name="FM_datoer" id="jq_dato" value="<%=tregDato %>" placeholder="dd-mm-yyyy" <%=inputDatoTXT %> />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>

                           
                        </td>
                        <td>&nbsp;</td>
                    </tr>
                    <%else %>
                    <input type="hidden" id="Hidden5" name="year" value="<%=year(now) %>"/>
                    <input type="hidden" name="FM_datoer" id="jq_dato" value="<%=tregDato %>" />
                    <%end if %>
                    
                   
                        <%
                        rwspan = 1
                        %>

                        
                     <tr>
                        <td style="padding-top:<%=paddingTop%>; padding-left:10px;" rowspan="<%=rwspan %>">
                            <input type="hidden" id="stempelurOn" name="" value="<%=stempelurOn%>"/>
                            
                            <input type="hidden" name="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
        
                            <input type="hidden" name="usemrn" id="treg_usemrn" value="<%=usemrn%>">
                            <input type="hidden" id="lto" value="<%=lto%>">

                            <!--<input type="hidden" id="Hidden3" name="FM_dager" value=""/>-->
                            <input type="hidden" id="Hidden4" name="FM_dager" value="0"/><!-- , xx -->
                            <input type="hidden" id="Hidden2" name="FM_feltnr" value="0"/>
                            <input type="hidden" id="FM_pa" name="FM_pa" value="<%=pa_aktlist %>"/>
                            
                            <input type="hidden" id="FM_medid" name="FM_medid" value="<%=usemrn %>"/>
                            
                            <input type="hidden" id="FM_medid_k" name="FM_medid_k" value="<%=usemrn%>"/>
                            <input type="hidden" id="" name="FM_vistimereltid" value="0"/>
                            <input type="hidden" id="lcid_sprog" name="" value="<%=lcid_sprog_val %>"/>

                            <input type="hidden" id="mobil_week_reg_akt_dd" name="" value="<%=mobil_week_reg_akt_dd %>"/>
                            <input type="hidden" id="mobil_week_reg_job_dd" name="" value="<%=mobil_week_reg_job_dd %>"/>
                            
                           
                            <!-- Forecast max felter -->
                            <input type="hidden" id="aktnotificer_fc" name="" value="0"/>
                            <input type="hidden" id="akt_maksbudget_treg" value="<%=akt_maksbudget_treg%>">  
                            <input type="hidden" id="akt_maksforecast_treg" value="<%=akt_maksforecast_treg%>">
                            <input type="hidden" id="aktBudgettjkOn_afgr" value="<%=aktBudgettjkOn_afgr%>">
                            <input type="hidden" id="regskabsaarStMd" value="<%=datePart("m", aktBudgettjkOnRegAarSt, 2,2)%>">
                            <input type="hidden" id="regskabsaarUseAar" value="<%=datepart("yyyy", varTjDatoUS_man, 2,2)%>">
                            <input type="hidden" id="regskabsaarUseMd" value="<%=datepart("m", varTjDatoUS_man, 2,2)%>">     
                            
                            <%
                            '*** Husk seneste job **'
                            jobidC = "-1"
                            aktidC = "-1"    
                    
                            %>

                           <input type="hidden" id="jq_jobidc" name="" value="<%=jobidC %>"/>
                           <input type="hidden" id="jq_aktidc" name="" value="<%=aktidC %>"/>
                            
                            <%if cint(mobil_week_reg_job_dd) = 1 then %>

                          
                             <input type="hidden" id="FM_job_0" value="-1"/>
                             <select id="dv_job_0" name="FM_jobid" style="font-size:<%=inputFont %>; height:<%=inputHeight%>;" class="form-control input-<%=txtClass %> chbox_job">
                                 <option value="-1"><%=left(tsa_txt_145, 4) %>..</option>
                                 <!--<option value="0">..</option>-->
                             </select>

                            <%else %>
                            <input type="text" style="font-size:<%=inputFont%>; height:<%=inputHeight%>" id="FM_job_0" name="FM_job" placeholder="<%=tsa_txt_066 %>/<%=tsa_txt_236 %>" class="FM_job form-control input-<%=txtClass %>"/>
                           <!-- <div id="dv_job_0" class="dv-closed dv_job" style="border:1px #cccccc solid; padding:10px; visibility:hidden; display:none;"></div>--> <!-- dv_job -->

                             <select id="dv_job_0" class="form-control input-small chbox_job" size="10" style="visibility:hidden; display:none;">
                                 <option><%=tsa_txt_534%>..her</option>
                             </select>

                            <input type="hidden" id="FM_jobid_0" name="FM_jobid" value="0"/>
                            <%end if %>

                        </td>
                     
                      
                        </tr><tr>
                       
                    
                        <!--<td style="padding-right:5px; padding-top:10px; vertical-align:text-top;"><b><%=tsa_txt_068%>:</b></td>-->
                         <td style="padding-top:<%=paddingTop%>; padding-left:10px; width:225px">
                            <%if cint(mobil_week_reg_akt_dd) = 1 then %>
                                 <input type="hidden" id="FM_akt_0" value="-1"/>
                                 <!--<textarea id="dv_akt_test"></textarea>-->
                                 <select id="dv_akt_0" name="FM_aktivitetid" class="form-control input-<%=txtClass %> chbox_akt" style="font-size:<%=inputFont%>; height:<%=inputHeight%>" DISABLED>
                                      <option>..</option>
                                  </select>

                             <%else %>
                              <input style="font-size:<%=inputFont%>; height:<%=inputHeight%>" type="text" id="FM_akt_0" name="activity" placeholder="<%=tsa_txt_068%>" class="FM_akt form-control input-<%=txtClass %>"/>
                                <!--<div id="dv_akt_0" class="dv-closed dv_akt" style="border:1px #cccccc solid; padding:10px; visibility:hidden; display:none;"></div>--> <!-- dv_akt -->

                                  <select id="dv_akt_0" class="form-control input-small chbox_akt" size="10" style="visibility:hidden; display:none;"> <!-- chbox_akt -->
                                      <option><%=tsa_txt_534%>..</option>
                                  </select>
                              <input type="hidden" id="FM_aktid_0" name="FM_aktivitetid" value="0"/>
                             <%end if %>
                           
                        </td>
                       
                    </tr>

                    <tr>
                        <!--<td style="padding-right:5px; padding-top:10px; vertical-align:text-top;"><b><%=tsa_txt_137%>:</b></td>-->
                         <td style="padding-top:<%=paddingTop%>; padding-left:10px">
                             <input type="hidden" id="FM_sttid" name="FM_sttid" value="00:00"/>
                             <input type="hidden" id="FM_sltid" name="FM_sltid" value="00:00"/>
                             <input type="text" id="FM_timer" name="FM_timer" placeholder="<%=tsa_txt_137%>" class="form-control input-<%=txtClass %>" style="font-size:<%=inputFont%>; height:<%=inputHeight%>" /><!-- brug type number for numerisk tastatur -->
                        </td>
                        
                    </tr>

                    <tr>
                        <!--<td style="padding-right:5px; padding-top:10px; vertical-align:text-top;"><b><%=tsa_txt_051%>:</b></td>-->
                        <td style="padding-top:<%=paddingTop%>; padding-left:10px; text-align:right;">
                            <input type="text" id="FM_kom" name="FM_kom_0" placeholder="<%=tsa_txt_051%>" class="form-control input-<%=txtClass %>" style="font-size:<%=inputFont%>; height:<%=inputHeight%>" /><br />
                            <%
                                btnSz = "sm"
                                btn_dis = ""
                            %>
                            <button id="btn_indlas" class="btn btn-<%=btnSz %> btn-success" <%=btn_dis %>><b><%=tsa_txt_085%> >></b></button>
                            <%'end if %>
                        </td>
                       
                    </tr>

                    
                  
      
                </table>
                </form>
                         
                     <%'end if %>

          </div>
           
          


            <div class="col-lg-<%=colspR%>"> 
            <%
            '**** Norm Chart ****'    
             %>
            <div id="stacked-vertical-chart" class="chart-holder-200"></div>
          
            </div>
     
        </div><!-- ROW --> 
        <%end if %>
        </div>
        


        <%end if  %>


        <%
            
            if browstype_client = "ip" then
            %>
            <div class="row">
            <div class="col-lg-12"><div id="stacked-vertical-chart" class="chart-holder-200"></div></div>
            </div>
            <%end if %>

            
        <%
            end if

        'end if
    
           
        


       ugeseddelvisning = 1

       perInterval = 6 'dateDiff("d", varTjDatoUS_man, varTjDatoUS_son, 2,2) 
       perIntervalLoop = perInterval

       'response.write "perIntervalLoop " & perIntervalLoop
       for l = 0 to perIntervalLoop 
        
            if l = 0 then
            varTjDatoUS_use = varTjDatoUS_man
            showheader = 1
            showtotal = 0
            else
            varTjDatoUS_use = dateAdd("d", l, varTjDatoUS_man)
            showheader = 0
            showtotal = 0
             end if 

           if l = perIntervalLoop then
           showtotal = 1
           end if

             varTjDatoUS_son = varTjDatoUS_man

            call ugeseddel(usemrn, varTjDatoUS_use, varTjDatoUS_use, ugeseddelvisning, showtotal, showheader)  
   
       next
   
	
	
        select case lto
        case "tec", "esn"
        aty_sql_realhours = " tfaktim <> 0"
        case "xdencker", "xintranet - local"
        aty_sql_realhours = " tfaktim = 1"
        case "cflow"
        aty_sql_realhours = " tfaktim = 1"
        case else
        aty_sql_realhours = aty_sql_realhours &""_
		& " OR tfaktim = 30 OR tfaktim = 31 OR tfaktim = 7 OR tfaktim = 11"
        end select 
   
                                

            'response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>aty_sql_realhours "& aty_sql_realhours
            'response.flush
            timerdenneuge_dothis = 0
            call timerDenneUge(usemrn, lto, varTjDatoUS_man, aty_sql_realhours, timerdenneuge_dothis, SmiWeekOrMonth)

                ugetotal = manTimer + tirTimer + onsTimer + torTimer + freTimer + lorTimer + sonTimer
                manTimer = replace(manTimer, ",",".")
                tirTimer = replace(tirTimer, ",",".")
                onsTimer = replace(onsTimer, ",",".")
                torTimer = replace(torTimer, ",",".")
                lorTimer = replace(lorTimer, ",",".")
                sonTimer = replace(sonTimer, ",",".")

                '*** NORM eller Komme/Gå som grundlag
                select case lto
                case "dencker", "intranet - local"

                    call fLonTimerPer(varTjDatoUS_man, 6, 21, usemrn)
                    ntimMan = ((manMin-manMinPause)/60)
                    ntimTir = ((tirMin-tirMinPause)/60)
                    ntimOns = ((onsMin-onsMinPause)/60)
                    ntimTor = ((torMin-torMinPause)/60)
                    ntimFre = ((freMin-freMinPause)/60)
                    ntimLor = ((lorMin-lorMinPause)/60)
                    ntimSon = ((sonMin-sonMinPause)/60)

                case else
                    call normtimerPer(usemrn, varTjDatoUS_man, 6, 0)
               
                end select

             norm_ugetotal = ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon

            'weeknormpro = ugetotal / norm_ugetotal * 100
            if cdbl(norm_ugetotal) <> 0 then
                
           
                    weeknormpro = ugetotal / norm_ugetotal * 100
                    weeknormpro = formatnumber(weeknormpro, 0)
                    weeknormpro = replace(weeknormpro, ",",".")
                

                 if ugetotal <= norm_ugetotal then
                 weeknormproWdt = weeknormpro
                 else
                 weeknormproWdt = norm_ugetotal / norm_ugetotal * 100 
                 end if

            else
                weeknormpro = 0
                norm_ugetotal = 0
            end if

            %>

            <input type="hidden" id="timerdagman" value="<%=replace(manTimer, ",", ".") %>" />
            <input type="hidden" id="timerdagtir" value="<%=replace(tirTimer, ",", ".") %>" />
            <input type="hidden" id="timerdagons" value="<%=replace(onsTimer, ",", ".") %>" />
            <input type="hidden" id="timerdagtor" value="<%=replace(torTimer, ",", ".") %>" />
            <input type="hidden" id="timerdagfre" value="<%=replace(freTimer, ",", ".") %>" />
            <input type="hidden" id="timerdaglor" value="<%=replace(lorTimer, ",", ".") %>" />
            <input type="hidden" id="timerdagson" value="<%=replace(sonTimer, ",", ".") %>" />


            <input type="hidden" id="normdagman" value="<%=replace(ntimMan, ",", ".") %>" />
            <input type="hidden" id="normdagtir" value="<%=replace(ntimTir, ",", ".") %>" />
            <input type="hidden" id="normdagons" value="<%=replace(ntimOns, ",", ".") %>" />
            <input type="hidden" id="normdagtor" value="<%=replace(ntimTor, ",", ".") %>" />
            <input type="hidden" id="normdagfre" value="<%=replace(ntimFre, ",", ".") %>" />
            <input type="hidden" id="normdaglor" value="<%=replace(ntimLor, ",", ".") %>" />
            <input type="hidden" id="normdagson" value="<%=replace(ntimSon, ",", ".") %>" />

        
          
          <div class="row">
                <div class="col-lg-12">
                    <h4 style="text-align:right"><%=tsa_txt_173 &" "& tsa_txt_137%>: <%=formatnumber(norm_ugetotal, 2) %> &nbsp <%=tsa_txt_535%>: <%=ugetotal %> &nbsp&nbsp</h4>
                    <table class="table table">
                        <tr>
                            <th colspan="11"><div style="background-color:lightgray; height:30px; width: 100%;"><div style="background-color:#207800; height:30px; width:<%=weeknormproWdt %>%; color:#ffffff; text-align:right; padding:5px;"><%=weeknormpro %> %</div></div></th>
                        </tr>
                        <tr>
                            <th style="border-top:none; width:10%;"><h6>00</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>20</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>40</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>60</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>80</h6></th>
                            <th style="border-top:none; width:10%;"><h6></h6></th>
                            <th style="border-top:none; width:10%;"><h6>100</h6></th>
                                        
                        </tr>
                       <!-- <tr>
                            <th style="text-align:left; border-top:none; font-size:85%" colspan="11">Norm timer: <%=norm_ugetotal %></th>
                        </tr>
                        <tr>
                            <th style="text-align:left; border-top:none; font-size:85%" colspan="11">Opgjort: <%=weeknormpro %>%</th>
                        </tr> -->
                    </table>
                </div>
            </div>
           






            <%if media <> "print" then %>

                <!--
                <form action="ugeseddel_2011.asp?media=export" method="post" target="_blank">
                    
                 <div class="row">
                     <div class="col-lg-12 pad-r30">
                         <input id="Submit6" type="submit" value="Eksport til .csv simpel" class="btn btn-sm" />
                      </div>


                </div>
                    </form>
        -->

          
            <br /><br />


         <%if browstype_client = "ip" then %>
           
            <br /><br /><br /><br />
            <form action="../timetag_web/timetag_web.asp">
                <div class="row">
                    <div class="col-lg-12">
                       <!-- <a href="../timetag_web/timetag_web.asp?flushsessionuser=1" class="btn btn-sm btn-default" style="text-align:center; width:100%"><b>Tilbage</b></a> -->
                        <button class="btn btn-sm btn-default" style="text-align:center; width:50%"><b><< <%=tsa_txt_437 %></b></button>
                    </div>
                </div>
            </form>
           
        <%end if %>

             
        <%
            
            if (browstype_client <> "ip") then
        %>
    <div class="well" style="width:35%;">
      <div class="portlet">
        <h3 class="portlet-title">
          <u><%=week_txt_003 %>:</u><!-- ugeseddel -->
        </h3>

          

           

              <form action="ugeseddel_2011.asp?media=print&usemrn=<%=usemrn %>&varTjDatoUS_man=<%=varTjDatoUS_man%>" method="post" target="_blank">
                    
                 <div class="row">
                     <div class="col-lg-12 pad-r30">
                         <input id="Submit6" type="submit" value="<%=tsa_txt_536 %> >>" class="btn btn-sm btn-default" />
                      </div>


                </div>
                    </form>


         

        


          
<table>

    <tr><td colspan="2">

           <!-- hvis level = 1 OR teamleder -->
        <%if (level <=2 OR level = 6) AND cint(smilaktiv) <> 0 AND SmiWeekOrMonth <> 2 then %>
          

        <%
        '** Tjekker for Uge 53. SKAL i virkeligheden være om søndag er i et andet år end mandag - da år så skifter.
        if datepart("ww", varTjDatoUS_son, 2,2) = 53 then
        tjkAar = year(varTjDatoUS_son) + 1
        else
        tjkAar = year(varTjDatoUS_son)
        end if    
            
        '**
        call erugeAfslutte(tjkAar, datepart("ww", varTjDatoUS_son, 2,2), usemrn, SmiWeekOrMonth, 0) 
        
        call godkendugeseddel(fmlink, usemrn, varTjDatoUS_man, rdir) %>
    
        <%end if 'level %>
    

             </td>
        </tr>



</table>
  </div> <!-- well -->
  </div> <!-- portlet title -->
            <%end if %>
      

<%else

    
Response.Write("<script language=""JavaScript"">window.print();</script>")
    
end if 

      %>


    </div>

    <%
       '** Wrapper / Content
        if browstype_client <> "ip" then %>
      </div></div>
     <%end if %>

     
   
    <br /><br />&nbsp;
    <%
	
	
	end select
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
