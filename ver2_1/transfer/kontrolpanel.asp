<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->


   
   
<%
if len(session("user")) = 0 then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<% 

	errortype = 5
	call showError(errortype)
	
	else
	
	'tjek for aktive moduler
	call erSDSKaktiv()
	if cint(dsksOnOff) = 1 then
	Strsdskstat = "checked"
	end if
	call erCRMaktiv()
	if cint(crmOnOff) = 1 then
	Strcrmstat = "checked"
	end if
	call erERPaktiv()
	if cint(erpOnOff) = 1 then
	Strerpstat = "checked"
	end if

	
	
	func = request("func")
	select case func
	case "exch"
			
			exchangeserverURL = request("FM_exchange")
			exchangeserverDOM = request("FM_exdom")
			exchangeserverBruger = replace(request("FM_exbruger"), "\", "#")
			exchangeserverPW = request("FM_expw")
			
			if len(request("FM_smiley")) <> 0 then
			smiley = request("FM_smiley")
			else
			smiley = 0
			end if
			
			if len(request("FM_stempelur")) <> 0 then
			stempelur = request("FM_stempelur")
			else
			stempelur = 0
			end if
			
			
			if len(request("FM_lukafm")) <> 0 then
			lukafm = request("FM_lukafm")
			else
			lukafm = 0
			end if
			
			
			if len(request("FM_autogk")) <> 0 then
			autogk = request("FM_autogk")
			else
			autogk = 0
			end if
			
			if len(request("FM_autogktimer")) <> 0 then
			autogktimer = request("FM_autogktimer")
			else
			autogktimer = 0
			end if
			
			
			
			if len(request("FM_autolukvdato")) <> 0 then
			autolukvdato = 1
			else
			autolukvdato = 0
			end if 
			
			if len(request("FM_autolukvdatodato")) <> 0 then
			autolukvdatodato = request("FM_autolukvdatodato")
			else
			autolukvdatodato = 28
			end if 
				
			
			
			'** Normal åbningstider ***
			if request("FM_brug_abningstid") = 1 then
			
			
			
			for x = 1 to 7
			
			select case x
			case 1
			    thisDay_t = request("FM_abn_man_t")
			    thisDay_m = request("FM_abn_man_m")
			    thisDay_t2 = request("FM_abn_man_t2")
			    thisDay_m2 = request("FM_abn_man_m2")
			case 2
			    thisDay_t = request("FM_abn_tir_t")
			    thisDay_m = request("FM_abn_tir_m")
			    thisDay_t2 = request("FM_abn_tir_t2")
			    thisDay_m2 = request("FM_abn_tir_m2")
			case 3
			    thisDay_t = request("FM_abn_ons_t")
			    thisDay_m = request("FM_abn_ons_m")
			    thisDay_t2 = request("FM_abn_ons_t2")
			    thisDay_m2 = request("FM_abn_ons_m2")
			case 4
			    thisDay_t = request("FM_abn_tor_t")
			    thisDay_m = request("FM_abn_tor_m")
			    thisDay_t2 = request("FM_abn_tor_t2")
			    thisDay_m2 = request("FM_abn_tor_m2")
			case 5
			    thisDay_t = request("FM_abn_fre_t")
			    thisDay_m = request("FM_abn_fre_m")
			    thisDay_t2 = request("FM_abn_fre_t2")
			    thisDay_m2 = request("FM_abn_fre_m2")
			 case 6
			    thisDay_t = request("FM_abn_lor_t")
			    thisDay_m = request("FM_abn_lor_m")
			    thisDay_t2 = request("FM_abn_lor_t2")
			    thisDay_m2 = request("FM_abn_lor_m2")
			 case 7
			    thisDay_t = request("FM_abn_son_t")
			    thisDay_m = request("FM_abn_son_m")
			    thisDay_t2 = request("FM_abn_son_t2")
			    thisDay_m2 = request("FM_abn_son_m2")
			end select
			        
			    %>
			    <!-- Validering -->
			    <%
			    if len(thisDay_t) <> 0 AND len(thisDay_m) <> 0_
			    AND len(thisDay_t2) <> 0 AND len(thisDay_m2) <> 0 then
			     
			        
			        
			        %><!-- Er det et gyldigt tidsformat --><%       
			        errThis = 0
			       
			        if IsDate(thisDay_t &":"& thisDay_m &":00") then
			        
			        else
			        errThis = 1
			        end if
			        
			        if IsDate(thisDay_t2 &":"& thisDay_m2 &":00") then
			        else
			        errThis = 1
			        end if
			        
			        
			        if cint(errThis) = 0 then
			        
			            if cdate((day(now) & "/" & month(now) & " / " & year(now) &" "& thisDay_t &":"& thisDay_m &":00")) > cdate(day(now) & "/" & month(now) & " / " & year(now)  &" "& thisDay_t2 &":"& thisDay_m2 &":00") then
			            errThis = 1
			            end if
    			        
    			        
			               if cint(errThis) = 0 then
					                
					            
					
                                select case x
			                    case 1
			                    normtid_st_man = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_man = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 2
			                    normtid_st_tir = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_tir = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 3
			                    normtid_st_ons = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_ons = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 4
			                    normtid_st_tor = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_tor = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 5
			                    normtid_st_fre = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_fre = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 6
			                    normtid_st_lor = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_lor = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 7
			                    normtid_st_son = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_son = thisDay_t2 &":"& thisDay_m2 &":00"
                			    
			                    end select
			            
			                
			                 else
			             
			          %>
					  <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					  <!--#include file="../inc/regular/topmenu_inc.asp"-->
			          <%
					    errortype = 79
					    call showError(errortype)
					    Response.End
					  end if
			                    
			                    
			          else
			             
			          %>
					  <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					  <!--#include file="../inc/regular/topmenu_inc.asp"-->
			          <%
					    errortype = 78
					    call showError(errortype)
					    Response.End
					  end if
			    
			    else
			            %>
					     <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--#include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 78
					call showError(errortype)
					isInt = 0
			        Response.End
			    end if
			
			
			next 
		            
		            
		    strSQLat = ", brugabningstid = 1, normtid_st_man = '"& normtid_st_man &"', "_
			&" normtid_sl_man = '"& normtid_sl_man &"', "_
			&" normtid_st_tir = '"& normtid_st_tir &"', "_
			&" normtid_sl_tir = '"& normtid_sl_tir &"', "_
			&" normtid_st_ons = '"& normtid_st_ons &"', "_
			&" normtid_sl_ons = '"& normtid_sl_ons &"', "_
			&" normtid_st_tor = '"& normtid_st_tor &"', "_
			&" normtid_sl_tor = '"& normtid_sl_tor &"', "_
			&" normtid_st_fre = '"& normtid_st_fre &"', "_
			&" normtid_sl_fre = '"& normtid_sl_fre &"', "_
			&" normtid_st_lor = '"& normtid_st_lor &"', "_
			&" normtid_sl_lor = '"& normtid_sl_lor &"', "_
			&" normtid_st_son = '"& normtid_st_son &"', "_
			&" normtid_sl_son = '"& normtid_sl_son &"' "
		            
		            
		        
            else 'Arb tid slået til
			
			    strSQLat = ", brugabningstid = 0"
			end if
			
			
			'*** Ignorer Stempelur Periode ***
			ignorer_st = request("FM_stempel_ignorerper_st_t") &":"& request("FM_stempel_ignorerper_st_m") &":00"
			ignorer_sl = request("FM_stempel_ignorerper_sl_t") &":"& request("FM_stempel_ignorerper_sl_m") &":00"
			
			        if IsDate(ignorer_st) then
			        else
			        errThis = 2
			        end if
			        
			        if IsDate(ignorer_sl) then
			        else
			        errThis = 2
			        end if
			 
			          
			          
			        if errThis = 2 then  %>
					     <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--#include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 80
					call showError(errortype)
					isInt = 0
			        Response.End
			        
			        end if
			
			
			'** Standard Pause ***
			if len(request("FM_stempel_standardpause_A")) <> 0 then
			stpause = request("FM_stempel_standardpause_A")
			else
			stpause = 0
			end if
			
			if len(request("FM_stempel_standardpause_B")) <> 0 then
			stpause2 = request("FM_stempel_standardpause_B")
			else
			stpause2 = 0
			end if
			
			if len(request("CHKcrm")) <> 0 then
			crm = 1
			else
			crm = 0
			end if
			
			if len(request("CHKerp")) <> 0 then
			erp = 1
			else
			erp = 0
			end if
			
			if len(request("CHKsdsk")) <> 0 then
			sdsk = 1
			else
			sdsk = 0
			end if
			
			if len(request("p1_man")) <> 0 then
			p1_man = 1
			else
			p1_man = 0
			end if
			
			if len(request("p1_tir")) <> 0 then
			p1_tir = 1
			else
			p1_tir = 0
			end if
			
			if len(request("p1_ons")) <> 0 then
			p1_ons = 1
			else
			p1_ons = 0
			end if
			
			if len(request("p1_tor")) <> 0 then
			p1_tor = 1
			else
			p1_tor = 0
			end if
			
			if len(request("p1_fre")) <> 0 then
			p1_fre = 1
			else
			p1_fre = 0
			end if
			
			if len(request("p1_lor")) <> 0 then
			p1_lor = 1
			else
			p1_lor = 0
			end if
			
			if len(request("p1_son")) <> 0 then
			p1_son = 1
			else
			p1_son = 0
			end if
			
			if len(request("p2_man")) <> 0 then
			p2_man = 1
			else
			p2_man = 0
			end if
			
			if len(request("p2_tir")) <> 0 then
			p2_tir = 1
			else
			p2_tir = 0
			end if
			
			if len(request("p2_ons")) <> 0 then
			p2_ons = 1
			else
			p2_ons = 0
			end if
			
			if len(request("p2_tor")) <> 0 then
			p2_tor = 1
			else
			p2_tor = 0
			end if
			
			if len(request("p2_fre")) <> 0 then
			p2_fre = 1
			else
			p2_fre = 0
			end if
			
			if len(request("p2_lor")) <> 0 then
			p2_lor = 1
			else
			p2_lor = 0
			end if
			
			if len(request("p2_son")) <> 0 then
			p2_son = 1
			else
			p2_son = 0
			end if
			
			
			%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			 
			call erDetInt(stpause)
			if isInt > 0 then
			errThis = 3
			else
			errThis = 0
			end if
			isInt = 0 
			
			call erDetInt(stpause2)
			if isInt > 0 then
			errThis = 3
			else
			errThis = 0
			end if
			isInt = 0 
			
			        if errThis = 3 then  %>
					     <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--#include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 81
					call showError(errortype)
					isInt = 0
			        Response.End
			        
			        end if
			
			
			
			'**** fakturnr rækkefølge ****
			if len(request("FM_erp_fakturanr")) <> 0 then
			fakturanr = request("FM_erp_fakturanr")
			        
			        call erDetInt(fakturanr)
			        if isInt > 0 then
			        errThis = 4
			        else
			        errThis = errThis
			        end if
			        isInt = 0 
			        
			end if
			
			rykkernr = 0
			
			if len(request("FM_erp_kreditnr")) <> 0 then
			kreditnr = request("FM_erp_kreditnr")
			        
			        call erDetInt(kreditnr)
			        if isInt > 0 then
			        errThis = 4
			        else
			        errThis = errThis
			        end if
			        isInt = 0 
			
			end if
			
		
			
			        
			        
			        if cint(errThis) = 4 then  %>
					     <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--#include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 84
					call showError(errortype)
					Response.End
			        
			        end if
			        
			        
			
			if len(request("FM_tsa_jobnr")) <> 0 then
			jobnr = request("FM_tsa_jobnr")
			        
			        call erDetInt(jobnr)
			        if cint(isInt) > 0 then
			        errThis = 5
			        else
			        errThis = errThis
			        end if
			        isInt = 0 
			
			end if
			
			
			if len(request("FM_tsa_tilbudsnr")) <> 0 then
			tilbudsnr = request("FM_tsa_tilbudsnr")
			        
			        call erDetInt(tilbudsnr)
			        if cint(isInt) > 0 then
			        errThis = 5
			        else
			        errThis = errThis
			        end if
			        isInt = 0 
			
			end if
			       
			        
			        if cint(errThis) = 5 then  %>
					     <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--#include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 91
					call showError(errortype)
		            Response.End
			        
			        end if
			        
			
			'if len(request("FM_erp_fakprocent")) <> 0 then
			'fakprocent = request("FM_erp_fakprocent")
			'        
			'        call erDetInt(fakprocent)
			'        if cint(isInt) > 0 then
			'        errThis = 6
			'        else
			'        errThis = errThis
			'        end if
			'        isInt = 0 
			
			'end if
			
			'*** Bruges ikke ***
			fakprocent = 0
			
			        if cint(errThis) = 6 then  %>
					     <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--#include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 96
					call showError(errortype)
		            Response.End
			        
			        end if
			
			
			%>
			<!-- Opdater db efter validering -->
			<%
			
			'ignorertid_st, ignorertid_sl 
			strSQL = "UPDATE licens SET owa = '" & exchangeserverURL & "', dom = '"& exchangeserverDOM &"', "_
			&" kontonavn = '"& exchangeserverBruger &"',  kontopw = '"& exchangeserverPW &"', "_
			&" smiley = '"& smiley &"', stempelur = '"& stempelur &"', "_
			&" lukafm = "& lukafm &", autogk = "& autogk &", autogktimer = "& autogktimer &", "_
			&" autolukvdato = " & autolukvdato & ", autolukvdatodato = " & autolukvdatodato &", "_
			&" ignorertid_st = '"& ignorer_st &"', ignorertid_sl = '"& ignorer_sl &"', "_
			&" stpause = "& stpause &", stpause2 = "& stpause2 &", "_
			&" crm = "& crm &", "_
			&" erp = "& erp &", "_
			&" sdsk = "& sdsk &", "_
			&" p1_man = "& p1_man &", "_
			&" p1_tir = "& p1_tir &", "_
			&" p1_ons = "& p1_ons &", "_
			&" p1_tor = "& p1_tor &", "_
			&" p1_fre = "& p1_fre &", "_
			&" p1_lor = "& p1_lor &", "_
			&" p1_son = "& p1_son &", "_
			&" p2_man = "& p2_man &", "_
			&" p2_tir = "& p2_tir &", "_
			&" p2_ons = "& p2_ons &", "_
			&" p2_tor = "& p2_tor &", "_
			&" p2_fre = "& p2_fre &", "_
			&" p2_lor = "& p2_lor &", "_
			&" p2_son = "& p2_son &", fakturanr = "& fakturanr &", "_
			&" rykkernr = "& rykkernr &", kreditnr = "& kreditnr &", "_
			&" jobnr = "& jobnr &", tilbudsnr = "& tilbudsnr &", fakprocent = "& fakprocent &""
			
			
			strSQL = strSQL & strSQLat & " WHERE id = 1"
			
			'Response.Write strSQL
			'Response.Flush
			
			oConn.execute(strSQL)
			
			Response.redirect "kontrolpanel.asp?menu=tok&func=exchopd"
			
	case "dwldb"			
	
					
					strSQL = "SELECT Mid, email, mnavn FROM medarbejdere WHERE Mid=" & session("mid")
					oRec.open strSQL, oConn, 3
					if not oRec.EOF then
					strEmail = oRec("email")
					strEditor = oRec("mnavn")
					end if
					oRec.close
					
	
					Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
					' Sætter Charsettet til ISO-8859-1 
					Mailer.CharSet = 2
					' Afsenderens navn 
					Mailer.FromName = "TimeOut DB Backup"
					' Afsenderens e-mail 
					Mailer.FromAddress = "support@outzource.dk"
					Mailer.RemoteHost = "smtp.tiscali.dk"
					' Modtagerens navn og e-mail
					Mailer.AddRecipient "Support OutZourCE", "support@outzource.dk"
					'Mailer.AddBCC "SK", "sk@outzource.dk" 
					' Mailens emne
					Mailer.Subject = "DB backup er bestilt af: "& lto &" !"
					
					' Selve teksten
					Mailer.BodyText = "" & "Backup er bestilt af "& strNavn & vbCrLf _ 
					& "Sendes til: " &strEmail & "" & vbCrLf 
					
					Mailer.SendMail
					
					strDato = year(now)&"/"&month(now)&"/"&day(now)
					
						oConn.execute("INSERT INTO dbdownload (dato, editor, email) VALUES ("_
						&"'"& strDato &"',"_
						&"'"& strEditor &"',"_
						&"'"& strEmail &"')")
						
					Response.redirect "kontrolpanel.asp?menu=tok"
				
	case else
%>
<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
<!--#include file="inc/dato.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<script language="javascript">
    var ns, ns6, ie, newlayer;

    ns4 = (document.layers) ? true : false;
    ie4 = (document.all) ? true : false
    ie5 = (document.getElementById) ? true : false
    ns6 = (document.getElementById && !document.all) ? true : false;

    function getLayerStyle(lyr) {
        if (ns4) {
            return document.layers[lyr];
        } else if (ie4) {
            return document.all[lyr].style;
        } else if (ie5) {
            return document.all[lyr].style;
        } else if (ns6) {
            return document.getElementById(lyr).style;
        }
    }

    function ShowHide(layer) {
        newlayer = getLayerStyle(layer)

        var styleObj = (ns4) ? document.layers[layer] : (ie4) ? document.all[layer].style : document.getElementById(layer).style;

        if (newlayer.visibility == "hidden") {
            newlayer.visibility = "visible";
            styleObj.display = ""
        }
        else if (newlayer.visibility == "visible") {
            newlayer.visibility = "hidden";
            styleObj.display = "none"
        }
    }
</script>



<div id="sindhold" style="position: absolute; left: 20; top: 80; visibility: visible; width: 900px">
    <table border="0" cellpadding="0" cellspacing="0" width="80%">
        <tr>
            
   <td valign="bottom">
                <h3>TimeOut Kontrolpanel</h3>
                
               </td>
        </tr>
    </table>
    <br>
    <table cellspacing="0" cellpadding="0" border="0" width="80%">
        <form method="post" action="kontrolpanel.asp?menu=tok&func=exch">
        
        <%
		strSQL = "SELECT owa, kontonavn, kontopw, dom, smiley, stempelur, "_
		&" lukafm, autogk, brugabningstid, "_
		&" normtid_st_man, "_ 
	    &" normtid_sl_man, "_
	    &" normtid_st_tir, "_
		&" normtid_sl_tir, "_
		&" normtid_st_ons, "_
		&" normtid_sl_ons, "_
		&" normtid_st_tor, "_
		&" normtid_sl_tor, "_
		&" normtid_st_fre, "_
		&" normtid_sl_fre, "_
		&" normtid_st_lor, "_
		&" normtid_sl_lor, "_
		&" normtid_st_son, "_
		&" normtid_sl_son, autolukvdato, "_
		&" autolukvdatodato, ignorertid_st, ignorertid_sl, stpause, stpause2, "_
		&" p1_man, "_
		&" p1_tir, "_
		&" p1_ons, "_
		&" p1_tor, "_
		&" p1_fre, "_
		&" p1_lor, "_
		&" p1_son, "_
		&" p2_man, "_
		&" p2_tir, "_
		&" p2_ons, "_
		&" p2_tor, "_
		&" p2_fre, "_
		&" p2_lor, "_
		&" p2_son, fakturanr, kreditnr, rykkernr, jobnr, tilbudsnr, fakprocent, autogktimer"_
	    &" FROM licens WHERE id = 1"
		
		'Response.Write strSQL
		'Response.Flush
		
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then
			ExchangeServerURL = oRec("owa")
			ExchangeServerDOM = oRec("dom")
			ExchangeServerBruger = replace(oRec("kontonavn"), "#", "\")
			ExchangeServerPW = oRec("kontopw")
			smiley = oRec("smiley")
			stempelur = oRec("stempelur")
			lukafm = oRec("lukafm")
			autogk = oRec("autogk")
			autogktimer = oRec("autogktimer")
			
			normtid_st_man = left(formatdatetime(oRec("normtid_st_man"), 3), 5) 
	        normtid_sl_man = left(formatdatetime(oRec("normtid_sl_man"), 3), 5)
	        normtid_st_tir = left(formatdatetime(oRec("normtid_st_tir"), 3), 5)
		    normtid_sl_tir = left(formatdatetime(oRec("normtid_sl_tir"), 3), 5)
		    normtid_st_ons = left(formatdatetime(oRec("normtid_st_ons"), 3), 5)
		    normtid_sl_ons = left(formatdatetime(oRec("normtid_sl_ons"), 3), 5) 
		    normtid_st_tor = left(formatdatetime(oRec("normtid_st_tor"), 3), 5)
		    normtid_sl_tor = left(formatdatetime(oRec("normtid_sl_tor"), 3), 5) 
		    normtid_st_fre = left(formatdatetime(oRec("normtid_st_fre"), 3), 5) 
		    normtid_sl_fre = left(formatdatetime(oRec("normtid_sl_fre"), 3), 5) 
		    normtid_st_lor = left(formatdatetime(oRec("normtid_st_lor"), 3), 5) 
		    normtid_sl_lor = left(formatdatetime(oRec("normtid_sl_lor"), 3), 5) 
		    normtid_st_son = left(formatdatetime(oRec("normtid_st_son"), 3), 5) 
		    normtid_sl_son = left(formatdatetime(oRec("normtid_sl_son"), 3), 5)
			
			brugabningstid = oRec("brugabningstid")
			
			strDag = oRec("autolukvdatodato")
			autolukvdato = oRec("autolukvdato") 
			
			ignorertid_st = left(formatdatetime(oRec("ignorertid_st"), 3), 5)
			ignorertid_sl = left(formatdatetime(oRec("ignorertid_sl"), 3), 5)
			
			stpauseA = oRec("stpause")
			stpauseB = oRec("stpause2")
			
			
			if oRec("p1_man") <> 0 then
			p1_manChk = "CHECKED"
			else
			p1_manChk = ""
			end if
			
			if oRec("p1_tir") <> 0 then
			p1_tirChk = "CHECKED"
			else
			p1_tirChk = ""
			end if
			
			if oRec("p1_ons") <> 0 then
			p1_onsChk = "CHECKED"
			else
			p1_onsChk = ""
			end if
			
			if oRec("p1_tor") <> 0 then
			p1_torChk = "CHECKED"
			else
			p1_torChk = ""
			end if
			
			if oRec("p1_fre") <> 0 then
			p1_freChk = "CHECKED"
			else
			p1_freChk = ""
			end if
			
			if oRec("p1_lor") <> 0 then
			p1_lorChk = "CHECKED"
			else
			p1_lorChk = ""
			end if
			
			if oRec("p1_son") <> 0 then
			p1_sonChk = "CHECKED"
			else
			p1_sonChk = ""
			end if
			
			if oRec("p2_man") <> 0 then
			p2_manChk = "CHECKED"
			else
			p2_manChk = ""
			end if
			
			if oRec("p2_tir") <> 0 then
			p2_tirChk = "CHECKED"
			else
			p2_tirChk = ""
			end if
			
			if oRec("p2_ons") <> 0 then
			p2_onsChk = "CHECKED"
			else
			p2_onsChk = ""
			end if
			
			if oRec("p2_tor") <> 0 then
			p2_torChk = "CHECKED"
			else
			p2_torChk = ""
			end if
			
			if oRec("p2_fre") <> 0 then
			p2_freChk = "CHECKED"
			else
			p2_freChk = ""
			end if
			
			if oRec("p2_lor") <> 0 then
			p2_lorChk = "CHECKED"
			else
			p2_lorChk = ""
			end if
			
			if oRec("p2_son") <> 0 then
			p2_sonChk = "CHECKED"
			else
			p2_sonChk = ""
			end if
			
			fakturanr = oRec("fakturanr")
			rykkernr = oRec("rykkernr")
			kreditnr = oRec("kreditnr")
			
			jobnr = oRec("jobnr")
			tilbudsnr = oRec("tilbudsnr")
			fakprocent = oRec("fakprocent")
			
		end if
		oRec.close 
            
            if len(strDag) <> 0 Then
            strDag = strDag 
            else
            strDag = 28
            end if
            
            
             
                       
            man_t = left(normtid_st_man, 2)
            man_m = right(normtid_st_man, 2)
            man_t2 = left(normtid_sl_man, 2)
            man_m2 = right(normtid_sl_man, 2)
            
            tir_t = left(normtid_st_tir, 2)
            tir_m = right(normtid_st_tir, 2)
            tir_t2 = left(normtid_sl_tir, 2)
            tir_m2 = right(normtid_sl_tir, 2)
            
            ons_t = left(normtid_st_ons, 2)
            ons_m = right(normtid_st_ons, 2)
            ons_t2 = left(normtid_sl_ons, 2)
            ons_m2 = right(normtid_sl_ons, 2)
            
            tor_t = left(normtid_st_tor, 2)
            tor_m = right(normtid_st_tor, 2)
            tor_t2 = left(normtid_sl_tor, 2)
            tor_m2 = right(normtid_sl_tor, 2)
            
            fre_t = left(normtid_st_fre, 2)
            fre_m = right(normtid_st_fre, 2)
            fre_t2 = left(normtid_sl_fre, 2)
            fre_m2 = right(normtid_sl_fre, 2)
            
            
            lor_t = left(normtid_st_lor, 2)
            lor_m = right(normtid_st_lor, 2)
            lor_t2 = left(normtid_sl_lor, 2)
            lor_m2 = right(normtid_sl_lor, 2)
            
            son_t = left(normtid_st_son, 2)
            son_m = right(normtid_st_son, 2)
            son_t2 = left(normtid_sl_son, 2)
            son_m2 = right(normtid_sl_son, 2)
            
            
            
            ignorertid_st_t = left(ignorertid_st, 2)
            ignorertid_st_m = right(ignorertid_st, 2)
            
            ignorertid_sl_t = left(ignorertid_sl, 2)
            ignorertid_sl_m = right(ignorertid_sl, 2)
            
           %>
           <tr>
                   <td bgcolor="#ffffff" style="border:1px #5582d2 solid; padding:10px;">
                <a href="javascript:ShowHide('kgnr');"><b>Generelt:</b></a></td>
                   </tr>
                   <tr><td style="padding:10px;"><div id="kgnr" name="kgnr" style="visibility: hidden; display: none">
                   
                    <a href="kontrolpanel.asp?menu=tok&func=dwldb">Bestil database backup (zip fil)</a><br>
                    <%
		                strSQL = "SELECT dato, editor FROM dbdownload ORDER BY id DESC"
		                oRec.open strSQL, oConn, 3 
		                if not oRec.EOF then
		                ald = "n"
		                strDato = oRec("dato")
		                strEditor = oRec("editor")
		                else
		                ald = "j"
		                strDato = "--"
		                strEditor = ""
		                end if
		                oRec.close
                		
		                if ald = "n" then
		                Response.write "Sidst bestilt d. "& strDato &" af "& strEditor &"<br>"
		                end if
                    %>
                    Ved at bestille jeres egen lokale backup af jeres database, sikres det at I altid
                    har en version af jeres data, liggende på jeres eget netværk. Backup'en bliver sendt
                    dig via email, og I modtager den inden for et døgn. <b>Det koster 495 kr.</b>
                    at bestille en lokal kopi af jeres database. (SQL format)<br>
                    <br>
                    <br>
                    &nbsp;<img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></div></td>
            </tr>
            </table>     
            <table cellpadding=0 cellspacing=0 border=0 width=720>
            <tr>
                <td width=520 bgcolor="#ffffff" style="border:1px #5582d2 solid; border-right:0px; padding:10px;">
                <a href="javascript:ShowHide('ktsa');"><b>TSA: (Time/sag)</b></a></td>
                  <td width=200 align="right" bgcolor="#ffffff" style="border:1px #5582d2 solid; border-left:0px; padding:10px;">
              Aktiver modul <b>*</b>:<input type="checkbox" name="CHKtsa" align="right" checked disabled/> </tr>
           <tr><td style="padding:10px;" colspan="2"><div id="ktsa" name="ktsa" style="visibility: hidden; display: none"> <table>
            <td style="padding:10px;">
                        
                    <ul>
                         <li><a href="javascript:NewWin_help('akt_typer.asp?menu=tok');" target="_self" class='vmenu'>
                            Aktivitets typer</a><br>
                        <li><a href="javascript:NewWin_help('fomr.asp?menu=tok');" target="_self" class='vmenu'>
                            Forretningsområder</a><br>
                            <li><a href="javascript:NewWin_help('milepale_typer.asp?menu=tok');" target="_self"
                                class='vmenu'>Milepæle typer</a><br>
                                <li><a href="javascript:NewWin_help('kundetyper.asp?menu=tok&ketype=e');" target="_self"
                                    class='vmenu'>Kunde/Kontakt Typer og Rabat % (Segmentopdeling)</a><br>
                                    <li><a href="javascript:NewWin_help('stfolder_gruppe.asp?menu=tok&ketype=e');" target="_self"
                                        class='vmenu'>Standard foldere (mapper) og grupper</a><br>
                                    
                                   
                                   
            
            </td>
           </tr>
           
            <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h5>Masker til jobnr og tilbudsnr:</h5>
            
            <%
                    if func = "exchopd" then
                    Response.write "<font color=red><b>Opdateret!</b><br><br></font>"
                    end if
                    %>            
                        
            Seneste <b>Job</b> havde nr: 
            
            <input id="FM_tsa_jobnr" name="FM_tsa_jobnr" style="width:100px;" type="text" value="<%=jobnr%>" /> (hel tal 0-1000000)
            <br />Kun ved joboprettelse tilføjes der et jobnummer til denne maske. <br />
            Ved redigering af jobnummer på et specifikt job, bliver denne hovedmaske ikke ændret.<br />
            Jobnr er unikt.
         
            <br /><br>
            Seneste <b>Tilbud</b> havde nr: 
            <input id="FM_tsa_tilbudsnr" name="FM_tsa_tilbudsnr" style="width:100px;" type="text" value="<%=tilbudsnr%>" /> (hel tal 0-1000000)
            <br />
            Tilbuds nummer er unikt. 
            <br />
                &nbsp;
            
           
            </td>
            </tr>
           
           
           <tr>
            <td>
            
            
            <table cellpadding=0 cellspacing=0 border=0><tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h5>Smileyordning og lukning af uger.</h5>  
            <%if smiley = 1 then
            smChk = "CHECKED"
            else
            smChk = ""
            end if %>
            <input type="checkbox" name="FM_smiley" id="FM_smiley" value="1" <%=smchk%>>
                <b>Aktiver Smiley ordning.</b>
                <%
            if func = "exchopd" then
            Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
            end if
                %>
                <br>
                Hvis Smiley ordningen aktiveres bliver det muligt for en medarbjeder at afslutte
                uge, samt at modtage glade / sure Smiley's alt efter om man er "up to date" med
                sine timeregistreringer.
               </td>
               </tr>
               <tr>
               <td style="padding:15px;">
               
                    <%if autogk = 1 then
                    autogkChk = "CHECKED"
                    else
                    autogkChk = ""
                    end if %>
                
                    <input type="checkbox" name="FM_autogk" id="FM_autogk" value="1" <%=autogkchk%>>
                    <b>Luk automatisk uger når medarbejder har godkendt.</b>
                    <%
                    if func = "exchopd" then
                    Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                    end if
                    %>
                    <br>
                    <u>Kræver at Smiley ordning er slået til</u>. Denne funktion gør at medarbejdere
                    ikke længere kan indtaste/redigere timer i de uger de har godkendt.
                    <b>Administrator brugere kan stadigvæk ændre i timer indtil der er oprettet en faktura på jobbet.</b>
                    <br />Der kan <b>godt</b> indtastes materialeforbrug / udgiftsbilag selvom en uge er godkendt og lukket.
                   
                   </TD>
                   </tr>
                   
                <tr>
               <td style="padding:15px;">
               
                    <%if autogktimer = 1 then
                    autogktimerChk = "CHECKED"
                    else
                    autogktimerChk = ""
                    end if %>
                
                    <input type="checkbox" name="FM_autogktimer" id="Checkbox1" value="1" <%=autogktimerChk%>>
                    <b>Godkend automatisk timer når medarbejder godkender uge via smiley.</b>
                    <%
                    if func = "exchopd" then
                    Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                    end if
                    %>
                    <br>
                    <u>Kræver at Smiley ordning er slået til</u>. Denne funktion gør at medarbejdere
                    automatisk godkender deres egne timer når de lukker deres uge via Smiley ordningen.
                   
                   </TD>
                   </tr>
                   <tr>
                   
                   
                   <%if autolukvdato = 1 then
                    autolukvdatoChk = "CHECKED"
                    else
                    autolukvdatoChk = ""
                    end if %>
                  
                
                   
                   <td style="padding:15px;">
                   <input type="checkbox" name="FM_autolukvdato" id="FM_autolukvdato" value="1" <%=autolukvdatoChk%>>
                   <b>Månedsafslutning. </b> (Luk automatisk måneder for indtastning.) <br />
                   Luk forrige måned for indtastninger d. 
                   <!--#include file="inc/weekselector_dag.asp"--> <!-- b -->
                   i den efterfølgende måned.
                   <%
                    if func = "exchopd" then
                    Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                    end if
                    %>
                    <br /><u>Kræver at Smiley ordning er slået til</u>. Denne funktion gør at medarbejdere
                    ikke længere kan indtaste/redigere timer, efter datoen har passeret den valgte afslutningsdato i den efterfølgende måned.
                    Indtastning bliver lukket uanset om medarbejder har godkendt uge.
                   
                    
                    <br />
                    <br /><br />
                    <b>
                    Regler for hvornår man kan redigere i indtastede timer og hvornår indtastning bliver lukket for redigering.
                    </b>
                    <table cellspacing=2 cellpadding=2 width=100% border=0>
                    <tr>
                        <td style="width:300px;"><b>Hvem må godkende</b></td>
                         <td style="width:300px;"><b>Hvornår bliver timer indtastning lukket</b></td>
                       </tr>
                       <tr>
                         <td valign=top>- Jobansvarlig 1 & 2<br />
                         - Administrator brugere</td>
                         <td valign=top>
                         - Når uge godkendes via <b>Smiley</b>*<br />
                         - Når måned lukkes via fast <b>Månedsafslutning</b>*<br />
                         - Når timer godkendes af Jobansvarlige eller administrator.<br />
                         - Når der oprettes en faktura.
                        
                        
                        <br /><br /> 
                         <b>*)</b> Hvis luk uger ved godkendelse er slået til. 
                         Dog kan timer, indtil fakturering ell. godkendelse, stadigvæk ændres af Administrator gruppen.
                         
                         
                         </td>
                    </tr>
                    </table>
                   
                   
                   </td></tr>
                   
                   
                   
                    <tr>
                    <td style="padding:15px; border-top:1px #5582d2 dashed;">
                    <h5>Færdigmelding af job.</h5> 
                    
                    <%if lukafm = 1 then
                    lukafmChk = "CHECKED"
                    else
                    lukafmChk = ""
                    end if %>
                    
                    <input type="checkbox" name="FM_lukafm" id="FM_lukafm" value="1" <%=lukafmchk%>>
                    <b>Aktiver færdigmelding af job.</b>
                    <%
                    if func = "exchopd" then
                    Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                    end if
                    %>
                    <br>
                    Aktiver mulighed for at medarbejdere, ved timeindtastning, selv kan lukke og færdigmelde et job. Ved lukning af job vil der blive sendt en email til jobansvarlige.
                    </td>
                    </tr>
                    
                    <tr>
                    <td style="padding:15px; border-top:1px #5582d2 dashed;">
                    <h5>Stempelur funktion.</h5> 
                    
                    <%if stempelur = 1 then
                    stempelurChk = "CHECKED"
                    session("stempelur") = 1
                    else
                    stempelurChk = ""
                    session("stempelur") = 0
                    end if %>
                   
                    
                        <input type="checkbox" name="FM_stempelur" id="FM_stempelur" value="1" <%=stempelurchk%>>
                        <b>Aktiver Stempelur.</b>&nbsp;(<a href="javascript:NewWin_help('stempelur.asp?menu=tok&ketype=e');"
                            target="_self" class='vmenu'>Stempelur indstillinger</a>)
                        <%
                        if func = "exchopd" then
                        Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                        end if
                        %>
                        <br>
                        Hvis stempelur funktionen aktiveres bliver medarbejdernes
                        logind og logud tidspunkt gemt og det bliver muligt at se hvor mange timer den enkelte
                        medarbejder har været logget ind i en given periode.<br /><br />
                        
                       <b>Ignorer følgende periode:</b> (medregn ikke minutter i det angivne interval)
                       <br /> 
                        <input id="FM_stempel_ignorerper_st_t" name="FM_stempel_ignorerper_st_t" type="text" style="width: 23px;" value="<%= ignorertid_st_t%>" />:
                       <input id="FM_stempel_ignorerper_st_m" name="FM_stempel_ignorerper_st_m" type="text" style="width: 23px;" value="<%= ignorertid_st_m%>" />
                       -
                        <input id="FM_stempel_ignorerper_sl_t" name="FM_stempel_ignorerper_sl_t" type="text" style="width: 23px;" value="<%= ignorertid_sl_t%>" />:
                       <input id="FM_stempel_ignorerper_sl_m" name="FM_stempel_ignorerper_sl_m" type="text" style="width: 23px;" value="<%= ignorertid_sl_m%>" />
                      <br />00:00 - 00:00 Ingen periode<br />
                      06:45 - 07:00 Tæller ikke minutter i de angivne 15 minutter.
                      
                      <br /><br />
                      <b>Standard pause A pr. dag:</b><br />
                      <input id="FM_stempel_standardpause_A" name="FM_stempel_standardpause_A" type="text" style="width: 23px;" value="<%= stpauseA%>" /> min. pr. dag<br />
                      <br />Tilføj pause 1 på følgende dage:<br />
                         Man 
                        <input id="p1_man" name="p1_man" value="1" type="checkbox" <%=p1_manChk %> />&nbsp;&nbsp;
                         Tir
                        <input id="p1_tir" name="p1_tir" value="1"type="checkbox" <%=p1_tirChk %> />&nbsp;&nbsp;
                         Ons 
                        <input id="p1_ons" name="p1_ons" value="1" type="checkbox" <%=p1_onsChk %> />&nbsp;&nbsp;
                         Tor 
                        <input id="p1_tor" name="p1_tor" value="1" type="checkbox" <%=p1_torChk %> />&nbsp;&nbsp;
                         Fre 
                        <input id="p1_fre" name="p1_fre" value="1" type="checkbox" <%=p1_freChk %> />&nbsp;&nbsp;
                         Lør 
                        <input id="p1_lor" name="p1_lor" value="1" type="checkbox" <%=p1_lorChk %> />&nbsp;&nbsp;
                         Søn 
                        <input id="p1_son" name="p1_son" value="1" type="checkbox" <%=p1_sonChk %> />&nbsp;&nbsp;
                       
                       <br /><br />
                      <b>Standard pause B pr. dag:</b><br />
                      <input id="FM_stempel_standardpause_B" name="FM_stempel_standardpause_B" type="text" style="width: 23px;" value="<%= stpauseB%>" /> min. pr. dag
                    
                       <br /><br />Tilføj pause 2 på følgende dage:<br />
                         Man 
                        <input id="p2_man" name="p2_man" value="1" type="checkbox" <%=p2_manChk %> />&nbsp;&nbsp;
                         Tir
                        <input id="p2_tir" name="p2_tir" value="1" type="checkbox" <%=p2_tirChk %> />&nbsp;&nbsp;
                         Ons 
                        <input id="p2_ons" name="p2_ons" value="1" type="checkbox" <%=p2_onsChk %> />&nbsp;&nbsp;
                         Tor  
                        <input id="p2_tor" name="p2_tor" value="1" type="checkbox" <%=p2_torChk %> />&nbsp;&nbsp;
                         Fre 
                        <input id="p2_fre" name="p2_fre" value="1" type="checkbox" <%=p2_freChk %> />&nbsp;&nbsp;
                         Lør 
                        <input id="p2_lor" name="p2_lor" value="1" type="checkbox" <%=p2_lorChk %> />&nbsp;&nbsp;
                         Søn  
                        <input id="p2_son" name="p2_son" value="1" type="checkbox" <%=p2_sonChk %> />&nbsp;&nbsp;
                                        
                                        
                       
                   
                    </td></tr>
                    </table>
                    <br /><br />
            <input type="image" src="../ill/opdaterpil.gif">
            <br /><br /></table>
            <img src="../ill/blank.gif" width="1" height="1" alt="" border="0">
                    </div>
                </td>
            </tr>
            </table>
            <table cellpadding=0 cellspacing=0 border=0 width=720>
            <tr>
                <td width=520 bgcolor="#ffffff" style="border:1px #5582d2 solid; padding:10px; border-right:0px;">
                <a href="javascript:ShowHide('kerp');"><b>ERP-Modul: (Fakturering og bogføring)</b></a></td>
                  <td width=200 align="right" bgcolor="#ffffff" style="border:1px #5582d2 solid; border-left:0px; padding:10px;">
              Aktiver modul <b>*</b>:<input type="checkbox" name="CHKerp" align="right" <%=Strerpstat%>/></td>
                      </tr>
           <tr>
            <td style="padding:10px;" colspan="2"><div id="kerp" name="kerp" style="visibility: hidden; display: none">
            <h5>Masker til faktura skrivelser:</h5>
            
            <%
                    if func = "exchopd" then
                    Response.write "<font color=red><b>Opdateret!</b><br><br></font>"
                    end if
                    %>            
                        
            Seneste <b>Faktura</b> havde nr: 
            
            <input id="FM_erp_fakturanr" name="FM_erp_fakturanr" style="width:100px;" type="text" value="<%=fakturanr%>" /><br />hel tal, maks 10 karakterer. 
            Ved brug af FI nummer må faktura nummer maks være 6 karakterer.
            
            <!--
            <br /><br>
            Seneste <b>Rykker</b> havde nr: 
            <input id="FM_erp_rykkernr" name="FM_erp_rykkernr" style="width:100px;" type="text" value="<%=rykkernr%>" /> (hel tal 0-1000000)
            -->
            <br /><br>
            Seneste <b>Kreditnota</b> havde nr: 
            <input id="FM_erp_kreditnr" name="FM_erp_kreditnr" style="width:100px;" type="text" value="<%=kreditnr%>" /> (hel tal, maks 10 karakterer)<br />
            Hvis faktura- og kreditnota -nr rækkefølge skal være en og samme nummerserie, angiv da -1 i Kreditnota feltet.  
            <br>
            <br /><br />
             <ul>
             <li><a href="erp_valutaer.asp" class=vmenu target=_blank>Valutaer og valutakurser</a>
             </ul>
             <br /><br />
            <input type="image" src="../ill/opdaterpil.gif">
            </div>
           </td>
            </table>
            </tr>
            </table>
           
            
             <table cellspacing=0 cellpadding=0 border=0 width=720>
                <tr>
                 <td bgcolor="#ffffff" width=520 style="border:1px #5582d2 solid; padding:10px; border-right:0px;">
                   <a href="javascript:ShowHide('kcrm');"><b>CRM-Modul:</b></a></td>
                <td width=200 align="right" bgcolor="#ffffff" style="border:1px #5582d2 solid; border-left:0px; padding:10px;">
                
              Aktiver modul <b>*</b>:<input type="checkbox" name="CHKcrm" align="right" <%=Strcrmstat%>/></td>
                </tr>
                
                <tr>
                <td style="padding:10px;" colspan="2"><div id="kcrm" name="kcrm" style="visibility: hidden; display: none">
                    <ul>
                        <li><a href="javascript:NewWin_help('crmstatus.asp?menu=tok&ketype=e');" target="_self"
                            class='vmenu'>Aktions status</a><br>
                            <li><a href="javascript:NewWin_help('crmemne.asp?menu=tok&ketype=e');" target="_self"
                                class='vmenu'>Aktions emner</a><br>
                                <li><a href="javascript:NewWin_help('crmkontaktform.asp?menu=tok&ketype=e');" target="_self"
                                    class='vmenu'>Aktions kontaktform</a><br>
                                    <li><a href="javascript:NewWin_help('kundetyper.asp?menu=tok&ketype=e');" target="_self"
                                        class='vmenu'>Kunde/Kontakt Typer (Segmentopdeling)</a><br>
                                        &nbsp;</ul>            <br /><br />
            <input type="image" src="../ill/opdaterpil.gif">
            <br /><br /></div>
                </td>
               </tr>
               
               
             </table>
            <table cellspacing=0 cellpadding=0 border=0 width=720>
            <tr>
                <td width=520 bgcolor="#ffffff" style="border:1px #5582d2 solid; border-right:0px; padding:10px;">
                <a href="javascript:ShowHide('ksdsk');"><b>SDSK-Modul (Servicedesk):</b></a></td>
              <td align="right" bgcolor="#ffffff" style="border:1px #5582d2 solid; border-left:0px; padding:10px;">Aktiver modul <b>*</b>:
              <input type="checkbox" name="CHKsdsk" align="right" <%=Strsdskstat%>/></td>
              </tr>
              <tr><td style="padding:10px;" colspan="2"><div id="ksdsk" name="ksdsk" style="visibility: hidden; display: none">
                    <ul>
                        <li><a href="javascript:NewWin_help('sdsk_prioitet.asp?menu=tok');" target="_self"
                            class='vmenu'>Incident prioiteter</a><br>
                            <li><a href="javascript:NewWin_help('sdsk_prio_typ.asp?menu=tok');" target="_self"
                                class='vmenu'>Prioitets klasser</a><br>
                                <li><a href="javascript:NewWin_help('sdsk_prio_grp.asp?menu=tok');" target="_self"
                                    class='vmenu'>Aftalegrupper</a><br>
                                    <li><a href="javascript:NewWin_help('sdsk_status.asp?menu=tok');" target="_self"
                                        class='vmenu'>Incident statustyper</a><br>
                                        <li><a href="javascript:NewWin_help('sdsk_typer.asp?menu=tok');" target="_self" class='vmenu'>
                                            Incident kategorier</a><br>
                                            &nbsp;
                        <h5>Virksomhedens åbningstider:</h5>
                            
                            
                            
                            
                           <%if brugabningstid = 1 then
                           baCHK = "CHECKED"
                           else
                           baCHK = ""
                           end if %>
                           
                            <input name="FM_brug_abningstid" id="FM_brug_abningstid" type="checkbox" value="1" <%=baCHK%>/>
                            <b>Brug åbningstider</b> <%
                            if func = "exchopd" then
                            Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                            end if
                            %><br />
                            Bruges til at beregne ServiceDesk aftalegruppe responstider.
                            Et incident indrapporteret onsdag morgen 04:45, med en aftalegruppe responstid på 48 timer, bliver<br />
                            således beregnet til at skulle være påbegyndt senest fredag kl. 9.00 når firmaet åbner.<br /><br />
                            Åbningstiderne bruges også af <b>stempeluret</b> ved glemt logud fra medarbejder.<br /><br />
                            <table cellspacing=1 cellpadding=2 border=0 bgcolor="#5582d2">
                                <tr>
                                    <td width=40 bgcolor="#ffffff" align="center">
                                        <b>Man:</b></td>
                                    <td bgcolor="#ffffff" style="width: 259px">
                                        <input id="FM_abn_man_t" name="FM_abn_man_t" type="text" style="width: 23px;" value="<%= man_t%>" />:
                                        <input id="FM_abn_man_m" name="FM_abn_man_m" type="text" style="width: 23px;" value="<%= man_m%>" />
                                        til
                                        <input id="FM_abn_man_t2" name="FM_abn_man_t2" type="text" style="width: 23px;" value="<%= man_t2%>" />:
                                        <input id="FM_abn_man_m2" name="FM_abn_man_m2" type="text" style="width: 23px;" value="<%= man_m2%>" />
                                   &nbsp;&nbsp;&nbsp;Eks: 08:15 - 17:00</td>
                                </tr>
                                <tr>
                                    <td bgcolor="#ffffff" align="center">
                                        <b>Tir:</b>
                                    </td>
                                    <td bgcolor="#ffffff" style="width: 259px">
                                        <input id="FM_abn_tir_t" name="FM_abn_tir_t" type="text" style="width: 23px;" value="<%=tir_t%>"/>:
                                        <input id="FM_abn_tir_m" name="FM_abn_tir_m" type="text" style="width: 23px;" value="<%=tir_m%>"/>
                                        til
                                        <input id="FM_abn_tir_t2" name="FM_abn_tir_t2" type="text" style="width: 23px;" value="<%=tir_t2%>" />:
                                        <input id="FM_abn_tir_m2" name="FM_abn_tir_m2" type="text" style="width: 23px;" value="<%=tir_m2%>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td bgcolor="#ffffff" align="center">
                                        <b>Ons:</b>
                                    </td>
                                    <td bgcolor="#ffffff" style="width: 259px">
                                        <input id="FM_abn_ons_t" name="FM_abn_ons_t" type="text" style="width: 23px;" value="<%=ons_t%>" />:
                                        <input id="FM_abn_ons_m" name="FM_abn_ons_m" type="text" style="width: 23px;" value="<%=ons_m%>" />
                                        til
                                        <input id="FM_abn_ons_t2" name="FM_abn_ons_t2" type="text" style="width: 23px;" value="<%=ons_t2%>" />:
                                        <input id="FM_abn_ons_m2" name="FM_abn_ons_m2" type="text" style="width: 23px;" value="<%=ons_m2%>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td bgcolor="#ffffff" align="center">
                                        <b>Tor:</b>
                                    </td>
                                    <td bgcolor="#ffffff" style="width: 259px">
                                        <input id="FM_abn_tor_t" name="FM_abn_tor_t" type="text" style="width: 23px;" value="<%=tor_t%>" />:
                                        <input id="FM_abn_tor_m" name="FM_abn_tor_m" type="text" style="width: 23px;" value="<%=tor_m%>" />
                                        til
                                        <input id="FM_abn_tor_t2" name="FM_abn_tor_t2" type="text" style="width: 23px;" value="<%=tor_t2%>" />:
                                        <input id="FM_abn_tor_m2" name="FM_abn_tor_m2" type="text" style="width: 23px;" value="<%=tor_m2%>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td bgcolor="#ffffff" align="center">
                                        <b>Fre:</b>
                                    </td>
                                    <td bgcolor="#ffffff" style="width: 259px">
                                        <input id="FM_abn_fre_t" name="FM_abn_fre_t" type="text" style="width: 23px;" value="<%=fre_t%>" />:
                                        <input id="FM_abn_fre_m" name="FM_abn_fre_m" type="text" style="width: 23px;" value="<%=fre_m%>" />
                                        til
                                        <input id="FM_abn_fre_t2" name="FM_abn_fre_t2" type="text" style="width: 23px;" value="<%=fre_t2%>" />:
                                        <input id="FM_abn_fre_m2" name="FM_abn_fre_m2" type="text" style="width: 23px;" value="<%=fre_m2%>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td bgcolor="#cccccc" align="center">
                                        <b>Lør:</b>
                                    </td>
                                    <td bgcolor="#cccccc" style="width: 259px">
                                        <input id="FM_abn_lor_t" name="FM_abn_lor_t" type="text" style="width: 23px;" value="<%=lor_t%>" />:
                                        <input id="FM_abn_lor_m" name="FM_abn_lor_m" type="text" style="width: 23px;" value="<%=lor_m%>" />
                                        til
                                        <input id="FM_abn_lor_t2" name="FM_abn_lor_t2" type="text" style="width: 23px;" value="<%=lor_t2%>" />:
                                        <input id="FM_abn_lor_m2" name="FM_abn_lor_m2" type="text" style="width: 23px;" value="<%=lor_m2%>" />
                                   &nbsp;&nbsp;&nbsp;Hvis lukket: 00:00 - 00:00 </td>
                                </tr>
                                <tr>
                                    <td bgcolor="#cccccc" align="center">
                                        <b>Søn:</b>
                                    </td>
                                    <td bgcolor="#cccccc" style="width: 259px">
                                        <input id="FM_abn_son_t" name="FM_abn_son_t" type="text" style="width: 23px;" value="<%=son_t%>" />:
                                        <input id="FM_abn_son_m" name="FM_abn_son_m" type="text" style="width: 23px;" value="<%=son_m%>" />
                                        til
                                        <input id="FM_abn_son_t2" name="FM_abn_son_t2" type="text" style="width: 23px;" value="<%=son_t2%>" />:
                                        <input id="FM_abn_son_m2" name="FM_abn_son_m2" type="text" style="width: 23px;" value="<%=son_m2%>" />
                                    </td>
                                </tr>
                            </table>
                        </li>
                    </ul>
                 <br /><br />
            <input type="image" src="../ill/opdaterpil.gif">
               <br /><br /></div></td>
                
            </tr>
            </table>
           
            <table cellspacing=0 cellpadding=0 border=0 width=80%>
            <tr>
                <td bgcolor="#ffffff" style="border:1px #5582d2 solid; padding:10px;">
                <a href="javascript:ShowHide('kexchange');"><b>Exchange server oplysninger:</b></a></td>
              </tr>
              <tr><td style="padding:10px;"><div id="kexchange" name="kexchange" style="visibility: hidden; display: none">
            
               
                  Bruges til integration med firmaets Exchange server.<br /><br />
                   <strong>
                 
                    Web-adresse til jeres Outlook Web Access:</strong>
                     <%
		                if func = "exchopd" then
		                Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
		                end if
                    %>
                    <br>
                    <b>https://</b><input type="text" name="FM_exchange" id="FM_exchange" size="40" value="<%=ExchangeServerURL%>"><b>/exchange</b>
                    &nbsp;&nbsp;(Eks: mail.outzource.dk)<br>
                    Exchangeserver domæne:<br>
                    <input type="text" name="FM_exdom" id="FM_exdom" size="40" value="<%=ExchangeServerDOM%>">
                    &nbsp;&nbsp;(Eks: outz)<br>
                    Brugerkonto navn:<br>
                    <input type="text" name="FM_exbruger" id="FM_exbruger" size="40" value="<%=ExchangeServerBruger%>">
                    &nbsp;&nbsp;(Eks: timeout, uden domænenavn som er angivet ovenfor.)<br>
                    Brugerkonto password:<br>
                    <input type="password" name="FM_expw" id="FM_expw" size="40" value="<%=ExchangeServerPW%>">
                    &nbsp;&nbsp;(Eks: Q12Wert3)<br>
                       <br>
    <br>
    <input type="image" src="../ill/opdaterpil.gif">
              </div></td>  
            </tr>
          </table>
          
      
       <br /><br />
            <input type="image" src="../ill/opdaterpil.gif">          

    </form>
    
    *) Se aktuelle priser på TimeOut moduler her:<br />
     <a href="http://www.outzource.dk/timeout_ver.asp" target="_blank" class=vmenu>http://www.outzource.dk/timeout_ver.asp..</a><br />
    <br>
    <br>
    <br>
    <a href="Javascript:history.back()">
        <img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
    <br>
    <br>
</div>
<%
	end select
	end if%><!--#include file="../inc/regular/footer_inc.asp"-->