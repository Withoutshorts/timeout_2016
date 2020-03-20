<%'GIT 20160811 - SK
'********************************'
'*** faktura layout opbygning ***'
'********************************'

'***** Fak header ******
public strTypeThis
public strTypeThis2
public pdfHTML
public filnavn, pdtop4
function fakheader() 

                select case lto
		        case "optionone"
		        pdtop1 = "125px 10px 0px 10px"
		        pdtop2 = "0px 0px 0px 0px"
		        pdtop3 = "5px 0px 0px 0px"
		        pdtop4 = "30px 2px 10px 10px"


		        case "dencker", "tooltest" ', "optionone"
		            
                    pdtop1 = "45px 10px 0px 10px"
		            pdtop2 = "0px 0px 0px 0px"
		            pdtop3 = "5px 0px 0px 0px"
		            pdtop4 = "10px 2px 10px 10px"
                   
		        
                case "essens"
		        pdtop1 = "10px 10px 0px 10px"
		        pdtop2 = "0px 0px 20px 0px"
		        pdtop3 = "20px 0px 0px 0px"
		        pdtop4 = "10px 2px 10px 10px"
		        case "acc"
		        pdtop1 = "15px 10px 0px 10px"
		        pdtop2 = "0px 0px 0px 0px"
		        pdtop3 = "55px 0px 0px 0px"
		        pdtop4 = "10px 2px 10px 10px"
                case "epi", "epi_no", "epi_sta", "epi_ab", "epi_uk", "epi2017"
                pdtop1 = "0px 10px 0px 10px"
		        pdtop2 = "0px 0px 20px 0px"
		        pdtop3 = "-10px 0px 0px 0px"
		        pdtop4 = "10px 2px 10px 10px"
                case "qwert"
                pdtop1 = "10px 10px 0px 0px"
		        pdtop2 = "0px 0px 20px 0px"
		        pdtop3 = "28px 0px 0px 0px"
		        pdtop4 = "10px 2px 10px 10px"
                case "synergi1", "intranet - local"
                pdtop1 = "20px 10px 0px 0px"
		        pdtop2 = "0px 0px 20px 0px"
		        pdtop3 = "10px 0px 0px 0px"
		        pdtop4 = "0px 2px 0px 0px"
		        case else
		        pdtop1 = "10px 10px 0px 10px"
		        pdtop2 = "0px 0px 20px 0px"
		        pdtop3 = "28px 0px 0px 0px"
		        pdtop4 = "10px 2px 10px 10px"
		        end select

                'pdtop1 'Main TD i fak. header
		        'pdtop2 'Faknr
		        'pdtop3 'Modtager TD
		        'pdtop4 'Afsender TD



                select case intFaktype
		case 0
		        
		        select case lto
		        case "rosenloeve"
		        
		        strtopimg = "faktura_top"
		        strTypeThis = "jobopgørelse"
		        strTypeThis2 = "Jobopgørelse"
		        
		        case else
		
		       		        
                strtopimg = "faktura_top"
                strTypeThis = lcase(erp_txt_001) '"faktura"
                strTypeThis2 = erp_txt_001 '"Faktura"
		        
		        
		        end select
		        
		        
		case 1
		strtopimg = "kreditnota_top"
		strTypeThis = lcase(erp_txt_002) '"kreditnota"
		strTypeThis2 = erp_txt_002 '"Kreditnota"   
		'case 2
		'strtopimg = "rykker_top"
		'strTypeThis = "rykker"
		'strTypeThis2 = "Rykker"  
		case else
		strtopimg = "blank"
		end select


        
select case lto 
case "synergi1", "intranet - local"
mtBd = 0
aftbtd1wdt = 340
aftbtd2wdt = 0
case "essens"
mtBd = 0
aftbtd1wdt = 370
aftbtd2wdt = 0
case else
mtBd = 0
aftbtd1wdt = 340
aftbtd2wdt = 280
end select

%>
<table cellspacing="0" cellpadding="0" border="<%=mtBd %>" width="<%=gblWdt%>"><!-- fakheader mian table -->

       
        
        <%select case lto 
        case "qwert"
        %>
         <tr>
        <td valign="top" style="padding:<%=pdtop1%>; width:<%=gblWdt%>px;">
        <%
        case "synergi1", "xintranet - local"
        %>
         <tr>
        <td valign="top" style="padding:<%=pdtop1%>; width:<%=aftbtd1wdt%>px;">
	    
	    <table width=100% cellspacing=0 cellpadding=0 border=0>
        <tr>
        <td valign="top" style="padding:<%=pdtop2 %>;">
		<%call fakturanr %>
        </td>
		</tr>
        	<tr>
		<td valign="top" style="padding:<%=pdtop3%>;">
		<%call modtager_layout() %>
		</td>
		</tr>
		</table>
        <%
        case else%>
        
        <tr>
        <td valign="top" style="padding:<%=pdtop1%>; width:<%=aftbtd1wdt%>px;">
	    
	    <table width=100% cellspacing=0 cellpadding=0 border=0>
        <tr>
        <td valign="top" style="padding:<%=pdtop2 %>;">
		<%call fakturanr %>
            
        </td>
		</tr>
        	<tr>
		<td valign="top" style="padding:<%=pdtop3%>;">
		<%call modtager_layout() %>
		</td>
		</tr>
		</table>

        </td>
		<td style="width:<%=aftbtd2wdt%>px;" valign=top align=right>
		
		
         <%end select %>
                                    
                                    
                                    
                                    <%
                                    '**** Højre TD '****'
                                    select case lto
                                    case "synergi1", "xintranet - local"

                                    case else

                                                                select case lto
                                                                case "qwert"
                                                                logo_align = "left"
                                                                bdr = 0
                                                                logoPd = "8px 3px 0px 0px"
                                                                case "essens"
                                                                logo_align = "right"
                                                                bdr = 0
                                                                logoPd = "8px 0px 0px 0px"
                                                                case else
                                                                bdr = 0
                                                                logo_align = "right"
                                                                logoPd = "8px 3px 0px 0px"
                                                                end select


                                                                 %>
		
		                                                        <table width="100%" cellspacing="0" cellpadding="0" border="<%=bdr %>">
		                                                        <tr><td align="<%=logo_align %>" colspan="4" valign="top" style="height:50px; padding:<%=logoPd%>;">
		                                                        <%
		
		                                                        filnavn = ""
		
		                                                        strSQLlogo = "SELECT f.filnavn, k.kid, k.logo FROM kunder k "_
		                                                        &" LEFT JOIN filer f ON (f.id = k.logo) WHERE k.kid = "& afsender
		
		                                                        'Response.Write strSQLlogo
		                                                        'Response.flush
		
		                                                        oRec.open strSQLlogo, oConn, 3
		                                                        if not oRec.EOF then
		
                                                                select case lto
                                                                case "qwert"
                                                                'filnavn = "qwert_logo_2011_600_59.gif"
                                                                'filnavn = "Top_justeret_maal.jpg"
                                                                filnavn = "Top_Faktura_300dpi_1.jpg"
                                                                 case "dencker"
                                                                 'if media = "pdf" then
                                                                 'filnavn = oRec("filnavn")
                                                                 'else
                                                                 filnavn = ""
                                                                 'end if
                                                                 case "synergi1", "intranet - local"
                                                                 filnavn = ""
                                                                case else
		                                                        filnavn = oRec("filnavn")
		                                                        end select

		                                                        %>
		                                                            <!--
                                                                    <img src="../inc/upload/<=lto%>/<=oRec("filnavn")%>" /> 
                                                                    <img src="../inc/upload/intranet/logo.jpg" />
                                                                    -->
		
		                                                        <%
		                                                        end if
		                                                        oRec.close
		
		                                                        if filnavn <> "" then 
                                    
                                                                         select case lto
                                                                         case "qwert", "intranet - local"%>
		                                                                 <img src="../inc/upload/<%=lto%>/<%=filnavn%>" width="625" height="62" border="0" />
                                                                         <%case else %>
                                                                         <img src="../inc/upload/<%=lto%>/<%=filnavn%>" />
                                                                         <%end select %> 
		                            
                                                                <%else %>
		                                                        <img src="../ill/blank.gif" /> 
		                                                        <%end if %>
		
		
		                                                        </td>
		                                                        </tr>



		                                                        <tr>
		                                                        <%
		                                                        pdtop = "30px 0px 0px 0px;"
		
		                                                        select case lto
                                                                case "qwert"
                                                                wdt = 140
                                                                hgt = 1
		                                                        case "acc"
		                                                        wdt = 140
                                                                hgt = 140
		                                                        case "outz"
		                                                        wdt = 120
                                                                hgt = 140
                                                                case "epi", "epi_no", "epi_sta", "epi_ab", "epi2017"
		                                                        wdt = 80
                                                                hgt = 140
		                           
		                                                        case "dencker", "tooltest", "optionone"
		                                                        pdtop = "45px 0px 0px 0px;"
		                                                        wdt = 80
                                                                hgt = 140
		                                                        case "essens"
		                                                        wdt = 100
		                                                        pdtop = "35px 0px 0px 0px;"
                                                                hgt = 140
		                                                        case else
		                                                        wdt = 100
                                                                hgt = 140
		                                                        end select
		                                                        %>
		                                                        <td style="width:<%=wdt %>"><img src="../ill/blank.gif" width="<%=wdt%>" height="1" border="0" /> 
                                                                  </td>
		                                                        <td align=right valign=top style="height:<%=hgt%>px; width:<%=280-wdt %>; padding:<%=pdtop %>">
		                                                        <%
		                                                         select case lto
		                                                         case "rosenloeve", "essens", "qwert", "synergi1"
		                                                         case else
		                                                            call afsender_layout()
		                                                         end select%>




		                                                        </td>
		                                                        </tr>
		                                                        </table>

                                    <%end select %>
        
        <!-- slut fak header main table -->
            
        </td>
		</tr>
		</table>
        <!---->
<%
        
        select case lto 
        case "qwert"
        %>
        <table width=100% cellspacing=0 cellpadding=0 border=0>
        <tr>
        <td valign="top" style="padding:<%=pdtop2 %>;">
		<%call modtager_layout() %>
        </td>
		</tr>
        	<tr>
		<td valign="top" style="padding:<%=pdtop3%>;">
		
        <%call fakturanr %>
		</td>
		</tr>
		</table>
        <%
        end select
     

end function


sub fakturanr
%>

		
		
		
		
		
        
        <%if (cint(medregnikkeioms) = 1 OR cdate(sysFakdato) = "01-01-2002" OR cint(intGodkendt) = 0) AND lto <> "essens" then%>
        <span style="text-decoration:line-through;"><h4>
            
            <%if cint(medregnikkeioms) = 2 then %>
             <%=erp_txt_058%><%=lcase(strTypeThis2)%>

             <%else %>
            <%=strTypeThis2%>
             <%end if %>

            - <%=varFaknr%> 
        
    

        <%if cint(medregnikkeioms) = 1 then %>
        &nbsp;<%=erp_txt_265 %>
        <%end if %>

        <%if cint(intGodkendt) = 0 then %>
        &nbsp;(<%=lcase(erp_txt_094) %>)
        <%end if %>

        <%if cdate(sysFakdato) = "01-01-2002" then %>
        &nbsp;(nulstillet)
        <%end if %>
        </h4></span>
        
        
        <%else %>
        <h4>
            
             <%if cint(medregnikkeioms) = 2 then %>
             <%=erp_txt_058%><%=lcase(strTypeThis2)%>

             <%else %>
            <%=strTypeThis2%>
             <%end if %>
            
            - <%=varFaknr%>

            


        </h4>
        

        <%end if%>
        
        
		<!--<img src="../ill/<=strtopimg%>.gif" width="271" height="25" alt="" border="0">-->
		
        <%
end sub

'***** Modtager boks ********
public forfaldsdato
function modtager_layout()



%>

<table cellspacing="0" cellpadding="0" border="0">

<tr bgcolor="#ffffff">
	
	<td valign="top" style="padding:0px 0px 0px 0px; height:162px;; border:0px silver solid;">
	
	<%
	'*** M65 afsender linie ***'
	select case lto
	case "acc", "syncronic", "pcm", "essens", "execon", "immenso", "epi", "epi_no", "epi_sta", "jttek", "qwert", "intranet - local", "synergi1", "epi_ab", "epi_uk", "epi2017", "bf"
	
	case "dencker", "tooltest", "optionone"
	%>
	<br /><br />
	<font size=1 color="#999999"><%=erp_txt_021 %>:&nbsp;<%=left(yourNavn, 25)%>&nbsp;<%=yourAdr%>&nbsp;<%=yourPostnr%>&nbsp;<%=yourCity%><br></font>
    <%case else%>
	<font size=1 color="#999999"><%=erp_txt_021 %>:&nbsp;<%=left(yourNavn, 25)%>&nbsp;<%=yourAdr%>&nbsp;<%=yourPostnr%>&nbsp;<%=yourCity%><br></font>
	<%end select %>
	

     
	
	<%if year(fakdato) < "2010" then %>
	
	<%=strKnavn%><br>
	<%=strKadr%><br>
	<%=strKpostnr%>&nbsp;<%=strBy%><br>
	
	    <%if cint(intVismodland) = 1 then
	    Response.write strLand &"<br>"
	    end if
	
	    if len(trim(ean)) <> 0 then 
	    Response.Write "EAN: "& ean & "<br>"
	    end if
	
	
	else
	
            select case lto
            case "synergi1"
            Response.Write replace(modtageradr, "CVR", "VAT")
            case "epi2017"
            
                select case afsender
                case "30001" 
                Response.Write replace(modtageradr, "CVR", "Organisasjonsnr")
                case "10001" 
                Response.Write replace(modtageradr, "CVR", "VAT")
                case else
                Response.Write modtageradr    
                end select  

            case else
            Response.Write modtageradr 
            End select 
	
	
	end if
	
	
	if intVismodatt = 1 then
	    select case lto
        case "synergi1"
        Response.write "<br><br>"& erp_txt_022 &": "& showAtt 
        case else
        Response.write "<br><b>"& erp_txt_022 &": "& showAtt &"</b>"
        end select
	end if 
	%>
	
	<%if fak_ski = 1 then
	Response.Write "<br><i>"& erp_txt_055 &"</i>"
	end if %>
	

	
	<br><br><br />
	<%=erp_txt_005 %>:&nbsp;<%=intKnr%>
	<br>
	<%=strTypeThis2 %><%=erp_txt_006 %>: <%=replace(formatdatetime(fakdato,2),"-",".")%>
	
	<%if cint(visperiode) = 1 then  %>
	<br />
	<%=erp_txt_049 %>: <%=replace(formatdatetime(istDato,2),"-",".")%> - <%=replace(formatdatetime(istDato2,2),"-",".")%>
	<%end if %>
	
	<%if cdate(forfaldsdato) >= cdate(fakdato) then  %>
	<br />
	<%=erp_txt_007 %>: <%=replace(formatdatetime(forfaldsdato,2),"-",".")%><br />
	<%end if %>

	</td>
	</tr>

</table>		
<%		
end function


'***** maintable **********
sub maintable


%>
<table cellspacing="0" cellpadding="0" border="0" width="<%=gblWdt%>">
	
	<tr>
		<td style="padding-left:0px;" valign="top" colspan="4">
<%
end sub

'**** maintable 2 *********
sub maintable_2
%>
</td>
<td width="20"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
<td style="padding-right:0px;" valign="top" colspan="2">
<%
end sub


'**** maintable 3 *********
sub maintable_3
%>
</td>
</tr>
</table>
<%
end sub


'****************************'
'****** Afsender boks *******'
'****************************'
function afsender_layout()

if lto = "synergi1" OR lto = "xintranet - local" then
afbd = 0
else
afbd = 0
end if

    select case lto
    case "jttek"
    afsenderCol = "#000000"
	case else
    afsenderCol = "#000000"
    end select
%>
<table cellspacing="0" cellpadding="0" border="<%=afbd%>" width=100%>

<tr>
	<td valign="top" style="padding:<%=pdtop4%>; color:<%=afsenderCol%>;" class="afsender2">
	
	<%
    '** Afsender navn og adr. ***'
    select case lto
    case "qwert", "xintranet - local"
	case "dencker", "tooltest"
	%>
	<b>Bank og konto nr.:</b><br />
	<%
	case else
	%>

        <% 
        select case lto
        case "epi_no", "epi_sta", "epi2017"

              if cdbl(afsender) = 30001 then 
              Response.Write yourNace & "<br />"
              end if  
       
        case else
        end select
        %>
	
    <%=yourNavn%><br>
	<%=yourAdr%><br>
	<%=yourPostnr%>&nbsp;<%=yourCity%><br>
	<%=yourLand & "<br>"%>
	
	

    <%select case lto
    case "jttek", "outz", "bf" 'T: F:
        
    if intVisafstlf = 1 then
	Response.write "<br>"& ucase(left(erp_txt_003, 1)) &": "& yourTlf
	end if
	
    if intVisafsfax = 1 then
	Response.write "<br>"& ucase(left(erp_txt_023, 1)) &": "& yourFax
	end if

    case else 'Tlf: / Fax:

    if intVisafstlf = 1 then
	Response.write "<br>"& erp_txt_003 &": "& yourTlf
	end if
	
    if intVisafsfax = 1 then
	Response.write "<br>"& erp_txt_023 &": "& yourFax
	end if
        
        
    end select %>

	
	<%

    select case lto
    case "xx" 
         

         '** Vores email med Email linje betegnelse: GL. standard   
        if intVisafsemail = 1 then
	    Response.write "<br>"& erp_txt_024 &": " & yourEmail
	    end if

       
    
    case else 'Uden Email linje overskrift: MODERNE 2016
       
         '** Vores email    
        if intVisafsemail = 1 then
	    Response.write "<br>"& yourEmail
	    end if 
   

    end select    
        
    %>
	
        
	<%
    '*** WWW
    if len(trim(yourWWW)) <> 0 AND lto = "jttek" then
	Response.write "<br>" & yourWWW
	end if%>
    

	
	<%end select%>



     <%
    '*** Bank info ***' 
    select case lto
    case "qwert"

    case "synergi1", "xintranet - local"
    %>

   

    <%

    case else %>

    <%if intVisafscvr = 1 then%>

          <%select case lto
        case "epi_no", "epi_sta"
        erp_txt_011 = "Organisasjonsnr"
        case "epi2017"
            
                if cdbl(afsender) = 30001 then 
                erp_txt_011 =  "Organisasjonsnr"
                end if  

                if cdbl(afsender) = 10001 then 
                erp_txt_011 = "VAT"
                end if  

       case else
        erp_txt_011 = erp_txt_011
        end select
         %>

	<br /><%=erp_txt_011 %>: <%=yourCVR%>
	<%end if%>
	
	<%if vorref <> "-" then%>
	<br /><%=erp_txt_056 %>: <%=vorref%><br />
	<%end if%>
	

    <%'*** BANK **** %>
	<br />
	<%=yourBank%><br>

    <%
    '**** KONTONR ****    
    select case lto
        
    case "jttek"%>

    Account:
    <%if len(trim(yourRegnr)) <> 0 then%>
	<%=yourRegnr & " "& yourKontonr%> 
    <%end if %>
    <br>

        
    <%case else %>
	<%=erp_txt_008 %>:<br />
   
    <%if len(trim(yourRegnr)) <> 0 then%>
	<%=yourRegnr%> - 
    <%end if %>
    <%=yourKontonr%><br>

    <%end select %>

	<%if intVisafsswift = 1 then%>
	<%=erp_txt_010 %>:&nbsp;<%=yourSwift%><br>
	<%end if%>
	
	<%if intVisafsiban = 1 OR lto = "jttek" then%>
	<%=erp_txt_009 %>:&nbsp;<%=yourIban%><br>
	<%end if%>
   
	
    <%end select %>
	
	
	<%
	'*** FI nummer **'
	select case lto
    case "qwert", "xintranet - local"
	case "dencker", "tooltest" ', "intranet - local"
	%>
	FI: <%=finummer %> 
	<%
	case else
	
	end select


    '** Rex Numer
    select case lto
    case "nt"
    %>
    REX-nr. DKREX180265723<br />
    <%
    end select
	%>

  
	
	</td>
</tr>
</table>
<%
end function




'**********************'
'****** Vedr. *********'
'**********************'

sub vedr
%>
<br /><br />
<table cellspacing="0" cellpadding="0" border="0" width="<%=gblWdt%>">
	<%select case lto
	case "syncronic"
	%>
	
	    <%
	
	case else

    select case lto
    case "synergi1", "intranet - local"
    vedr_pd = "5px 105px 5px 0px"

    case else
    vedr_pd = "5px 15px 5px 0px"
    end select
	    
	    %>
	    <tr>
	    <td style="padding:25px 5px 2px 0px; border-bottom:1px #999999 solid;">

            <% select case lto
            case "nt", "intranet - local"
           %>
             &nbsp;
            <%
            case else
           %>
             <b><%=erp_txt_014 %></b>
            <%
            end select%>

	   
	    </td>
	    </tr>
	    <tr>
	    <td style="padding:<%=vedr_pd%>;">
	    
       
        <%
	    
	    if jobid <> 0 then%>
	        <%if cint(visikkejobnavn) <> 1 then %>
           
            <%select case lto 
               case "nt", "intranet - local"
                %>
               <!-- NT no.: <=intJobnr%> -->
               <%case else%>
             <br /><b><%=strJobnavn%>&nbsp;(<%=intJobnr%>)</b><br />
            <%end select %>
            
            
            <%else 'PGA sideombrydning %>
            <br /><br /> 
	        <%end if %>
	    
	        <%if rekvnr <> "0" OR lto = "nt" then 
	        
                select case lto 
                  case "nt", "xintranet - local"
                %>
                <br /><!--NorthTex order no.: <%=intJobnr%><br />-->
                <%if len(trim(supplier_invoiceno)) <> 0 then %>
                Supplier Invoice no.: <%=supplier_invoiceno%><br />
                <%end if %>

               <%case else%>
             <%=erp_txt_053 %>: <b><%=rekvnr%></b><br /><br />
            <%end select %>

	        <%end if %>
	    
		    <%if (len(trim(strJobBesk)) <> 0 AND cint(vis_jobbesk) = 1) then%>
		    <%=strJobBesk%><br><br>
		    <%end if%>
	    <%
	    else
	    %>
        
        <%if cint(visikkejobnavn) <> 1 then %>
	    <br /><b><%=strAftNavn%>&nbsp;(<%=intAftnr%>)</b>
        <%end if %>

	            <%if lto = "outz" OR lto = "intranet" then %>
	            <br /><%=erp_txt_026 %> / Licensnummer:
                     <%if len(trim(strAftVarenr)) <> 0 AND strAftVarenr <> "0" then %>
                        <br>
	                    <%=erp_txt_026 %>: <%=strAftVarenr%>
                        <%else %>
                        <br />
                        <%end if 

	            else %>

                        <%if len(trim(strAftVarenr)) <> 0 AND strAftVarenr <> "0" then %>
                        <br />
	                    <%=erp_txt_026 %>: <%=strAftVarenr%>
                        <%else %>
                        <br />
                        <%end if 


	           
	            end if


                    


                if (len(trim(strJobBesk)) <> 0 AND cint(vis_jobbesk) = 1) then%>
		        <br /><%=strJobBesk%><br/ ><br />
		        <%end if


	    end if  %>
        &nbsp;
	    </td>
	    </tr>
	    <%
	
    end select %>
	
	
	
	

</table>
<br /><br />
<%
end sub



public fakuraKolonneOverskrifterIsWrt 
sub fakuraKolonneOverskrifter

    %>
     <tr>
            	
                <%'** POS kolonne produktion 
                   select case lto 
                                 
                    case "dencker"
                        if cint(fak_fomr) = 16 then 'Grundfos
                    %>
                    <td style="width:30px; padding-right:5px; border-bottom:1px #999999 solid;"><b>Pos.</b></td>
                    <%
                        end if
                    case else
                                 
                    end select



                '** Skjul antal og enh kolonner **' (aktivitetslinier)
                if (cint(hideantenh) <> 1 AND aktmat = 1) OR aktmat <> 1 then
                %>
                <td style="width:40px; padding-left:0px; border-bottom:1px #999999 solid;"><b><%=erp_txt_013 %></b></td>
                <%end if %>

            	
	             <%
            	    
                    '** Skjul antal og enh kolonner **'
                    if (cint(hideantenh) <> 1 AND aktmat = 1) OR aktmat <> 1 then

	                if (enhedsangFak = -1 AND aktmat = 1) then %>
	                <td style="padding-right:5px; width:40px; border-bottom:1px #999999 solid;">
                    &nbsp;</td>
                    <%else %>
	                <td style="padding-right:5px; width:40px; border-bottom:1px #999999 solid;"><b><%=erp_txt_044 %></b></td>
	                <%end if 
                    
                    end if
                        
                        
                   select case lto 
                   case "nt"
                   header_txt = "Style"
                   case else
                   header_txt = erp_txt_057
                   end select%>
            	
	            
            	
               
	            <td colspan="2" style="width:200px; padding-left:0px; border-bottom:1px #999999 solid;"><b><%=header_txt %></b> <img src="../ill/blank.gif" width="200" height="1" border="0" /></td>
	            <td align="right" style="width:80px; padding-right:2px; border-bottom:1px #999999 solid;">&nbsp;</td>

                <% 
                '** Skjul antal og enh kolonner **'
                if (cint(hideantenh) <> 1 AND aktmat = 1) OR aktmat <> 1 then%>
                <td align=right style="width:100px; padding-right:2px; border-bottom:1px #999999 solid;"><b><%=erp_txt_015 %></b></td>
                 <%else %>
                <td style="border-bottom:1px #999999 solid;">&nbsp;</td>
                <%end if %>

	            <%if cint(visrabatkol) = 1 then %>
	            <td align=right style="width:50px; padding-right:5px; border-bottom:1px #999999 solid;"><b><%=erp_txt_016 %></b></td>
	            <%else %>
	            <td style="width:50px; border-bottom:1px #999999 solid;">
                    &nbsp;</td>
	            <%end if %>
	           
	            <td align=right style="width:120px; border-bottom:1px #999999 solid;"><b><%=erp_txt_017 %></b></td>


                  <%if cint(fastpris) = 2 then 'commi %>
                    <td align="right" style="width:60px; border-bottom:1px #999999 solid;"><b>Comi</b></td>
                    <td align="right" style="width:80px; border-bottom:1px #999999 solid;"><b>Commission</b></td>
                    <%end if %>
            	
            </tr>
            <%

fakuraKolonneOverskrifterIsWrt = 1
end sub





'****** Udspecificering *******
'*** Overskrifts linie på aktiviteter / materialer **'
public aktmat
function udspecificering(strFakdet, aktmat, antalAktlinier, fastpris)

 if cint(hideantenh) <> 1 then
        select case lto 
        case "dencker"
            if cint(fak_fomr) = 16 then'Grundfos
            cspanAktlinjer = 10
            else
            cspanAktlinjer = 9
            end if
        case else
        cspanAktlinjer = 9
        end select
 end if


'*** Materialer uden moms bruges ikke mere. De vises sammen med de andre
if (aktmat = 3) AND antalAktlinier <> 0 AND cint(hideantenh) <> 1 AND len(trim(strFakdet)) <> 0 AND am <> 0 then %>
<tr>
<td colspan="<%=cspanAktlinjer %>" style="padding:10px 0px 2px 0px;"><b><%=erp_txt_041 %></b></td>
</tr>
<%end if 



'*** ØVRIGE MATERIALE OMKOSTNINGER ***'
if ((aktmat = 2) AND antalAktlinier <> 0 AND cint(hideantenh) <> 1 AND len(trim(strFakdet)) <> 0 AND am <> 0) AND (lto <> "essens") then 
%>
<tr>
<td colspan="<%=cspanAktlinjer %>" style="padding:10px 0px 2px 0px;"><b><%=erp_txt_039 %></b></td>
</tr>
<%
end if


'**** MATERIALE OVERSKIFTER HVIS DER IKKE FINDES AKTIVITETSLINJER
if (aktmat = 2 AND fakuraKolonneOverskrifterIsWrt = 0 AND len(trim(strFakdet)) <> 0) then

    call fakuraKolonneOverskrifter

end if


'*** AKTIVITETS OVERSKIFTER 
if (aktmat = 1) then
 
    call fakuraKolonneOverskrifter

end if %>            
            
        
<%=strFakdet%>

<%
end function


'***** Sub-Total aktiviteter***

function subtotakt()
'** Bruges kun til at lave afstand før materialer eller total beløb på faktura. **'
'** Total antal er fjernet efter at KM er kommet med på faktura                **'
%>

<tr>
	<td valign='top' align="right" style="padding:5px 5px 10px 0px;">
	<%select case lto 
	case "-"
	%>
	<b><%=formatnumber(intFaktureretTimer, 2)%></b>
	<%
	case else  %>
	&nbsp;
	<%end select %>
	</td>
	<td colspan=7>
        &nbsp;</td>
</tr>

<!--
<tr>
    <td valign='top' colspan=3 style="padding:0px 0px 0px 0px;"><=erp_txt_027 %></td>
	<td valign='top' colspan=5 align=right style="padding:0px 0px 0px 0px;"><=formatnumber(aktsubtotal, 2) &" "& valutaISO %></td>
</tr>
-->


<%
end function

'***** Sub-Total materialer***'

function subtotmat(m)


%>
<!--
<tr>
    <td valign='top' colspan=3 style="padding:5px 0px 0px 0px;"><%=erp_txt_027 %></td>
    <td valign='top' colspan=5 align=right style="padding:5px 0px 0px 0px;"><%=formatnumber(matsubtotal_fakgk(m), 2) &" "& valutaISO %></td>
</tr>
-->

<%
end function



'***** Total før moms og Moms ***'
function totogmoms()

'if len(strFakdet) = 0 then
select case lto
case "synergi1"
case else
%>
<br /><br />
<%end select %>
<table cellspacing="0" cellpadding="1" border="0" width="<%=gblWdt%>">
<%
'end if

if cint(showmomsfriTxt) <> 0 then
trpadtop = 5%>
<tr><td valign="top" colspan="8" style="padding:25px 0px 0px 0px;">*) 0&#37; <%=erp_txt_034 %></td></tr>
<%
else

trpadtop = 25


end if %>
<tr>
    <td valign="top" colspan="4" style="padding:<%=trpadtop%>px 0px 0px 0px;"><%=erp_txt_029 %>&nbsp;</td>
   <td colspan="4" valign="top" align="right" style="padding:<%=trpadtop%>px 0px 0px 0px;"> <%=formatnumber(intFaktureretBelob) &" "& valutaISO%></td>
</tr>



    
    <!--formatnumber(intMoms, 2) &" <> "& formatnumber((subtotalTilmoms*0.25), 2) -->
    <%'** Moms dobbeltjek ****'
    momsTjk = (subtotalTilmoms*(momssats/100))
    if formatnumber(intMoms, 2) <> formatnumber(momsTjk, 2) AND media <> "xml" AND cint(medregnikkeioms) <> 1 AND cint(medregnikkeioms) <> 2 then

    'Response.write momsTjk & "<br>"

    momsTjk = formatnumber(momsTjk, 2)
    intMoms = momsTjk
    momsTjk = replace(momsTjk, ".", "")
    momsTjk = replace(momsTjk, ",", ".")
    strSQLupdmoms = "UPDATE fakturaer SET moms = "& momsTjk & " WHERE fid = "& id
    oConn.execute(strSQLupdmoms) 


    
   
    %>
   <%end if 
   
   if cint(intFaktype) = 1 AND media <> "xml" then
   intMoms = (intMoms*-1)
   end if
   

    
    
   %>
   
 <tr>
    <td valign="top" colspan="4" style="padding:5px 0px 5px 0px;"><%=momssats &"&#37; "& erp_txt_034 %></td>
    <td colspan="4" valign="top" align="right" style="padding:5px 0px 5px 0px;">
        
        <%if cint(intFaktype) = 1 then%>
        -<%=formatnumber(intMoms,2) &" "& valutaISO%>    
        <% else %>
        <%=formatnumber(intMoms,2) &" "& valutaISO%>
        <%end if %>
    </td>
	
</tr>
</table>
<%end function 


'***** Total beløb ***'
function totalbelob_fakgk()%>

<table cellspacing="0" cellpadding="1" border="0" width="<%=gblWdt%>">

<%
select case lto
case "essens"
case else

t = 1000
if len(trim(strFakmat)) <> 0 AND t = 1 then 'strFakmat(1) %>
<tr>
     <td valign="top" colspan="4" style="padding:25px 0px 0px 0px;"><%=erp_txt_031 %>&nbsp;</td>
	<td valign="top" colspan="4" align="right" style="padding:25px 0px 0px 0px;"><%=formatnumber(intFaktureretBelob, 2) &" "& valutaISO%></td>
</tr>
<%end if 

end select
%>
 <%

    if cint(intFaktype) = 1 then
    intMoms = (intMoms*-1)
    end if

	totalinclmoms = (intFaktureretBelob/1) + (intMoms/1) 
	%>

<tr>
   <td valign="top" colspan="4" style="padding:5px 5px 5px 0px; border-top:1px #999999 solid; border-bottom:1px #999999 solid;"><b><%=erp_txt_032 %></b>&nbsp;</td>
   <td valign="top" colspan="4" align="right" style="padding:5px 0px 5px 5px; border-top:1px #999999 solid; border-bottom:1px #999999 solid;"><b><%=formatnumber(totalinclmoms, 2) &" "& valutaISO%></b></td>
</tr>
</table>
<%
end function


'*** Komm. og bet. betingelser ***
sub komogbetbetingelser	

%>

<table cellspacing="0" cellpadding="1" border="0" width="<%=gblWdt%>">
<tr>
<td colspan="8" style="border-top:0px silver solid; padding:35px 5px 5px 1px;">
<%select case lto




case "xdencker"

if h > 16 then
toppx = 1848
else
toppx = 808
end if

if media = "pdf" then%>
<div style="width: <%=gblWdt%>px; position: absolute; top: <%=toppx%>px">
<img src="../ill/dencker/faktura_pdf_footer.gif" />
</div>
<%
end if

case "qwert"

if h > 16 then
toppx = 1848
else
toppx = 848
end if


%>

<div style="width: <%=gblWdt%>px; position: absolute; top: <%=toppx%>px">


<table cellspacing="0" cellpadding="0" width="100%" border="0">
<tr><td align=center>
    Angiv venligst fakturanr. ved indbetaling til Danske Bank kontonr. 3001 3143077128<br />
Overskrides betalingsfristen tilskrives der rente med 12% p.a. Ved rykkerskrivelse pålægges kr. 100,00 i gebyr.

</td></tr>
</table>
</div>

<%

case "synergi1", "intranet - local"


'if h > 16 OR instr(strJobBesk, "##breakpage##") <> 0 then
'toppx = 1948
'else
'toppx = 988
'end if

toppx = 100

%>

<div style="width: <%=gblWdt%>px; position: realtive; top: <%=toppx%>px;">


<table cellspacing="0" cellpadding="0" width="100%" border="0">
<tr><td align="center" class="footer">

Betaling: <%=trim(strKomm)%> - senest <%=replace(formatdatetime(forfaldsdato,2),"-",".")%><br />
Ved betaling efter forfaldsdato tilskrives 2% rente pr. påbegyndt måned, min. 50,- kr.

	
	<%'if intVisafsswift = 1 then%>
	&nbsp;<br /><%=erp_txt_010 %>:&nbsp;<%=yourSwift%>.
	<%'end if%>
	
	<%'if intVisafsiban = 1 then%>
	&nbsp;<%=erp_txt_009 %>:&nbsp;<%=yourIban%>.
	<%'end if%>

    <%if intVisafscvr = 1 then%>
	&nbsp;<%=erp_txt_011 %>: <%=yourCVR%>
	<%end if%>

<br>Betalings-ID: <%=finummer %> 
<br /><%=yourBank &" "&erp_txt_008 %>: <%=yourRegnr%> - <%=yourKontonr%><br />

</td></tr>
</table>
</div>

<%

case "essens" 

if h > 16 then
toppx = 1848
else
toppx = 878
end if

%>
<%=strKomm%>
    
<div>&nbsp;</div>
<div>
Ved senere betaling beregnes rente p&aring; 1,5% pr. p&aring;begyndt m&aring;ned. Rykkergebyr kr. 50,-.<br />
Vores timepris indeksreguleres &aring;rligt pr. 1. juli.

<!-- Ved senere betaling beregnes rente p&#229; 1,5% pr. p&#229;begyndt m&#229;ned. Rykkergebyr kr. 50,-. <br /> -->

<br />
<br />
<br />

<br />
<br />

<br />
<br />
<br />
<br />
<br />

<br />
    &nbsp;
    </div>

<div style="width: <%=gblWdt%>px; position: absolute; top: <%=toppx+20%>px; z-index:900000000000; height:100px; border:0px #999999 dashed; padding:0px 0px 0px 0px;">
<table cellspacing="0" cellpadding="0" width="100%" border="0">
    <tbody>
   
      <td class="footer" style="white-space:nowrap;" valign="top">Essens Kommunikation&nbsp;&nbsp;/&nbsp;&nbsp;Carl Jacobsens Vej 37E&nbsp;&nbsp;/&nbsp;&nbsp;DK-2500 Valby<br />
            Telefon: (+45) 70 229 149&nbsp;&nbsp;/&nbsp;&nbsp;www.essens.info<br />
      <br />&nbsp;
           </td>
    </tr>
        <tr>
            <td class="footer" style="white-space:nowrap;" valign="bottom">Bank: Sydbank&nbsp;&nbsp;/&nbsp;&nbsp;Reg.: 7040&nbsp;&nbsp;/&nbsp;&nbsp;Konto: 0001568855&nbsp;&nbsp;/&nbsp;&nbsp;IBAN: DK7770400001568855&nbsp;&nbsp;/&nbsp;&nbsp;Swift: SYBKDK22&nbsp;&nbsp;/&nbsp;&nbsp;CVR-nr.: 27022995</td>
        </tr>
    </tbody>
</table>
</div>


<%
case "execon", "imenso"

if h > 16 then
toppx = 1848
else
toppx = 828
end if

%>
<%=strKomm%>

<div style="width: <%=gblWdt%>px; position: absolute; top: <%=toppx%>px">
<table cellspacing="0" cellpadding="0" width="100%" border="0">
  
        <tr>
            <td valign="bottom">
            <font size="1" face="Helvetica">
            Betaling med frig&#248;rende virkning kan alene ske til : AL Finans A/S, Edithsvej 2 A, Postboks 352, 2600 Glostrup. </font></font><br />
            <strong><font size="1" face="Helvetica-Bold"><font size="1" face="Helvetica-Bold">Bankkonto 5301 0931482<br />
            </strong></font>
            <font size="1" face="Helvetica"><font size="1" face="Helvetica">Vor fordring i henhold til n&#230;rv&#230;rende faktura er overdraget til AL Finans A/S - Levering har fundet sted.Ved betaling bedes leverand&#248;r samt fakturanummer anf&#248;rt. Eventuelle uoverensstemmelser bedes meddelt til AL Finans A/S, Telefon 43 43 49 90</font>
            </td>
            </tr>
            </table>

</div>

<%
case else %>
<%if len(trim(strKomm)) <> 0 then %>
<b><%=erp_txt_033 %>:</b><br><%=strKomm%>
<%end if %>
<%end select %>

<br>&nbsp;</td>
</tr>
</table>

<%
end sub



'**** Brevhoved *************

'**** Brevfod ***************







function rykkergebyrer()

               
               strSQLryk = "SELECT rykkertxt, rykkerantal, rykkerbelob, rykkerdato, "_
               &" fr.valuta, fr.kurs, v.valutakode "_
               &" FROM faktura_rykker fr "_
               &" LEFT JOIN valutaer v ON (v.id = fr.valuta) WHERE fakid = "&  id
               oRec.open strSQLryk, oConn, 3
               r = 0
               while not oRec.EOF 
               
               if r = -10 then%>
                <tr>
               <td align=right style="padding-right:2px; padding-top:5px; border-bottom:1px #999999 solid;"><b><%=erp_txt_013 %></b></td>
               
               <td colspan=5 style="padding-left:5px; padding-top:5px; border-bottom:1px #999999 solid;"><b><%=erp_txt_014 %></b></td>
               <td style="border-bottom:1px #999999 solid; padding-top:5px;"><b><%=erp_txt_42 %></b></td>
               <td align=right style="border-bottom:1px #999999 solid; padding-top:5px;"><b><%=erp_txt_017 %></b></td>
               
               </tr>
               <%end if
              
               
               %><tr>
               <td valign=top align=right style="padding-right:7px;"><%=formatnumber(oRec("rykkerantal"), 2) %></td>
               <td>
                   &nbsp;</td>
              
               <td colspan=4 style="padding-left:5px;"><%=formatdatetime(oRec("rykkerdato"), 1) %> - <%=oRec("rykkertxt") %></td>
               <td>
                   &nbsp;</td>
               <td valign=top align=right><%=formatnumber(oRec("rykkerbelob"), 2) & " "& oRec("valutakode") %></td>
               </tr>
               <%
             
               intFaktureretBelob = intFaktureretBelob + oRec("rykkerbelob")
               
               r = r + 1 
               oRec.movenext
               wend
               oRec.close
               
              
               

end function



public am, antalMatTot
sub matlinjer

            strFakmat = ""

            antalMatTot = 0
            antalmatthis = 0
			'for m = 0 to 1 
			matBelob = 0
            matEhpris = 0

            if len(trim(fastpris)) <> 0 then
            fastpris = fastpris
            else
            fastpris = 0
            end if
                            
                           
				'** vis linje / linje eller grupper ***'
				strSQLmat = "SELECT fmat.matid, fmat.matvarenr,"_
				&" fmat.matrabat, matshowonfak, ikkemoms, "_
				&" fmat.valuta, fmat.kurs, v.valutakode, fmat.matenhed, fmat.matgrp,  " 
							    
				if cint(showmatasgrp) = 1 then
				strSQLmat = strSQLmat & "mgp.navn AS matnavn, "_
				&" SUM(matbeloeb) AS matbeloeb, SUM(fmat.matantal) AS matantal, (sum(matbeloeb)/sum(fmat.matantal)) AS matenhedspris "
				else
				strSQLmat = strSQLmat & "fmat.matnavn, fmat.matantal, fmat.matbeloeb, fmat.matenhedspris "
				end if

                select case lto
                case "nt", "xintranet - local"
                strSQLmat = strSQLmat & ", matfrb_id, mf.matkobspris, mf.ava "
                case else
                end select
							    
				strSQLmat = strSQLmat & "FROM fak_mat_spec fmat "_
				&" LEFT JOIN valutaer v ON (v.id = fmat.valuta) "
							    
				if cint(showmatasgrp) = 1 then
				strSQLmat = strSQLmat & " LEFT JOIN materiale_grp AS mgp ON (mgp.id = fmat.matgrp) "
				end if
							    

                select case lto
                case "nt", "xintranet - local"
                    strSQLmat = strSQLmat & " LEFT JOIN materiale_forbrug AS mf ON (mf.id = fmat.matfrb_id) "
                case else
                end select

				strSQLmat = strSQLmat & " WHERE matfakid = "& id
							    
                if cint(showmatasgrp) = 2 then 'fordel under aktiviteter

                   strSQLmat = strSQLmat & " AND fms_aktid = "& fd_aktid

                end if

				if cint(showmatasgrp) = 1 then
				    strSQLmat = strSQLmat & " GROUP BY fmat.matgrp, fmat.ikkemoms, fmat.matenhed, fmat.valuta, matrabat"
				end if

                strSQLmat = strSQLmat & " ORDER BY fmat.matgrp, matnavn"
							    
				'&" AND ikkemoms = "& m
								
				'Response.write strSQLmat
				'Response.Flush
								
                am = 0
				oRec2.open strSQLmat, oConn, 3
                while not oRec2.EOF
                                        
                                        
                        fmat_valuta = oRec2("valuta")
						fmat_kurs = oRec2("kurs")
						fmat_valutaKode = oRec2("valutakode")
                                        
                        'fmat_Ehpris = oRec2("matbeloeb")
						'fmat_Belob = oRec2("matenhedspris")
                                        
                                        
                        if fak_kurs <> fmat_kurs then
						matEnhPris = oRec2("matenhedspris") * (fmat_kurs/fak_kurs)
						matEnhPris = formatnumber(matEnhPris, 2)
                        else
						matEnhPris = oRec2("matenhedspris") 
						end if

                        select case lto
                        case "nt", "xintranet - local"
                        matAva = oRec2("ava")
                        case else
                        matAva = 0
                        end select
                                         
                            '*** Mat beløb er omregnet til den valuta **'
                            '** der er valgt på faktura inden den lægges i DB **'
                            if len(trim(oRec2("matbeloeb"))) <> 0 then
                            matBelob = formatnumber(oRec2("matbeloeb"), 2)
                            else
                            matBelob = formatnumber(0, 2)
                            end if
                                        
                        'matBelob = oRec2("matbeloeb") * (100/fak_kurs)
                        'matEhpris = oRec2("matenhedspris") * (100/fak_kurs)
                                        
                        if cint(fastpris) = 2 then 'commi
                
                        fmat_comm = 0
                        select case lto
                        case "nt", "xintranet - local"
                                         
                            'fmat_kobspris = formatnumber(oRec2("matkobspris"), 3)
                            'fmat_kobspris = (fmat_kobspris/1 + matEnhPris/1)
                                        
                        if matEnhPris <> 0 AND len(trim(matEnhPris)) <> 0 then    
                        fmat_kobspris = matEnhPris/1 + oRec2("matkobspris") '(matEnhPris/1 * matAva/1) * 100
                                    
                        if fmat_kobspris <> 0 then
                                fmat_kobspris = formatnumber(fmat_kobspris, 2) 
                        else
                                fmat_kobspris = 0
                        end if

                        else
                        fmat_kobspris = 0
                        end if
            
                        fmat_kobspris_tot = matAntal * (fmat_kobspris/1)
           
                            'matEnhPris = matEnhPris
                            fmat_comm = (matAva * 100) 
                        case else
                        end select                            

                        end if


                        if intFaktype = 1 AND media <> "xml" then
                        'fmat_Ehpris = -(formatnumber(fmat_Ehpris, 2))
						'fmat_Belob = -(formatnumber(fmat_Belob, 2))
						matBelob = -(formatnumber(matBelob, 2))
						matEnhPris = -(matEnhPris) 
						end if
                                        
                        '** Materierl uden for gruppe vises som diverse **'
                        if cint(showmatasgrp) = 1 AND oRec2("matgrp") = 0 then
                        matNavn = "Diverse"
                        else
                                select case lto
                                case "nt", "intranet - local"
                                matNavn = oRec2("matnavn") & " (ord. no: "& oRec2("matvarenr") &")"
                                case else        
                                matNavn = oRec2("matnavn")
                                end select

                        end if
                                        
                                        
                        '*** ÆØÅ **'
                        'enhedsang = oRec2("matenhed")

                        select case oRec2("matenhed")
	                    case ""
	                    enhedsang = "" 'Ingen (skjul)
	                    case "Stk."	
	                    'enhedsang = erp_txt_019 '"Stk. pris"
	                    enhedsang = erp_txt_036
                        case else
	                    enhedsang = oRec2("matenhed")
	                                    
	                    end select

                        'enhedsang = replace(enhedsang, "stk.", "Stk.")


                      


						call jquery_repl_spec(enhedsang)
                        enhedsang = jquerystrTxt
                                        
            
                        if len(trim(oRec2("matantal"))) <> 0 then
                        matAntal = formatnumber(oRec2("matantal"), 2)
                        else
                        matAntal = formatnumber(0, 2)
                        end if
                                        

                        strFakmat = strFakmat &"<tr>"

                                 select case lto 
                                 case "dencker"
                                     if cint(fak_fomr) = 16 then'Grundfos
                                    
                                     strFakmat = strFakmat &"<td>&nbsp;</td>"
                                     else
                                     strFakmat = strFakmat
                                     end if
                                 case else
                                 strFakmat = strFakmat 
                                 end select

						strFakmat = strFakmat &"<td valign=top align=right style='padding-right:10px; width:27px;'>"& matAntal &"</td>"_
						&"<td valign=top style='padding-right:5px; width:40px;'>"& enhedsang &"</td>"_
				        &"<td colspan=2 valign=top style='padding-left:0px;'>" & matNavn &"</td>"_
						&"<td valign=top align=right style='padding-right:2px;'>"
					    
					                    
					    '*** Momsfri ***'
					    if cint(oRec2("ikkemoms")) = 1 then
					    mat_momsfriVal = "*"
					    showmomsfriTxt = 1
					    'matsubtotal_fakgk(m) = matsubtotal_fakgk(m)
					    else
					    mat_momsfriVal = ""
					    'matsubtotal_fakgk(m) = matsubtotal_fakgk(m) + (matBelob)
					    end if
								        
						if len(trim(matEnhPris)) <> 0 then
                                            
                            '** Må gerne have 3 decimaler hvis de f.eks kommer fra lager
                            if instr(matEnhPris, ",") <> 0 then
                                            
                            len_matEnhPris = len(matEnhPris)
                            instr_matEnhPris = instr(matEnhPris, ",")
                            mid_matEnhPris = mid(matEnhPris, instr_matEnhPris, len_matEnhPris)
                            len_mid_matEnhPris = len(mid_matEnhPris)
                                           
                                if len_mid_matEnhPris > 3 then
                                matEnhPris = formatnumber(matEnhPris, 3)
                                else
                                matEnhPris = formatnumber(matEnhPris, 2)
                                end if

                            else
                                matEnhPris = formatnumber(matEnhPris, 2)
                            end if

								           
						else
						matEnhPris = formatnumber(0, 2)
						end if
								        
						'*** Enh. pris og Valuta på materiale linie ***'
				        if cint(fak_valuta) <> cint(fmat_valuta) then
				        strFakmat = strFakmat & matEnhPris &" "& fmat_valutaKode & " ~"
				        else
				        strFakmat = strFakmat & "&nbsp;"
				        end if

                                      
                    					
				        strFakmat = strFakmat &"</td>"
								
								        
						'*** Enh. pris og Valuta på materiale fra Faktura valuta ***'
                        strFakmat = strFakmat &"<td valign=top align=right style='padding-right:2px;'>"

						if cint(fastpris) = 2 then 'COMI
                        strFakmat = strFakmat & formatnumber(fmat_kobspris, 2) &" "& valutaISO &""
                        else
                        strFakmat = strFakmat & matEnhPris &" "& valutaISO &""
                        end if
                                       
                        strFakmat = strFakmat &"</td>"



						'**** Rabat *********************
						if cint(visrabatkol) = 1 then
						strFakmat = strFakmat &"<td valign=top align=right style='padding-right:10;'>"
								       
							if cdbl(oRec2("matrabat")) <> 0 then
							strFakmat = strFakmat &"" & (oRec2("matrabat") * 100) &" %</td>"
							else
							strFakmat = strFakmat  &"&nbsp;</td>"
							end if
						else
						strFakmat = strFakmat  &"<td>&nbsp;</td>"
						end if
								        

                        '**** Total beløb ***************
                        strFakmat = strFakmat &"<td valign='top' align=right>"
                        if cint(fastpris) = 2 then 'COMI
						strFakmat = strFakmat & mat_momsfriVal &" "& formatnumber(matAntal * fmat_kobspris) &" "& valutaISO  
                        else
                        strFakmat = strFakmat & mat_momsfriVal &" "& formatnumber(matBelob) &" "& valutaISO 
                        end if  
            
                        strFakmat = strFakmat &"</td>"
                 

                                        if cint(fastpris) = 2 then 'commi

                                        if len(trim(fmat_comm)) <> 0 then
                                        fmat_comm = formatnumber(fmat_comm, 2)
                                        else 
                                        fmat_comm = 0
                                        fmat_comm = formatnumber(fmat_comm, 2)
                                        end if

                                        if len(trim(matBelob)) <> 0 then
                                        matBelob = formatnumber(matBelob, 2)
                                        else
                                        matBelob = 0
                                        matBelob = formatnumber(matBelob, 2)
                                        end if

                                        strFakmat = strFakmat &"<td valign='top' align=right>"& fmat_comm &"%</td>"
                                        strFakmat = strFakmat &"<td valign='top' align=right>"& matBelob &" "& valutaISO &"</td>"
                                        end if

						strFakmat = strFakmat &"</tr>"
                                    
                                        
                            if media = "xml" then
                                          
                            session.LCID = 1033
                                          
                            'call utf_format(oRec2("matenhed"))
                            'enhedsang = utf_formatTxt
                                          
                            'call utf_format(matNavn)
                            'matNavn = utf_formatTxt
                                          
					        strXML = strXML & "<com:InvoiceLine>"
                            strXML = strXML & " <com:ID>"&h&"</com:ID>" 
                            strXML = strXML & " <com:InvoicedQuantity unitCode='"&EncodeUTF8(enhedsang)&"' unitCodeListAgencyID=""n/a"">"&replace(matAntal, ",", "")&"</com:InvoicedQuantity>" 
                            strXML = strXML & " <com:LineExtensionAmount currencyID='"&valutaISO&"'>"&replace(matBelob, ",", "")&"</com:LineExtensionAmount>" 
                            strXML = strXML & "<com:Item>"
                            strXML = strXML & "<com:ID>0</com:ID>" 
                            strXML = strXML & "<com:Description><![CDATA["&EncodeUTF8(matNavn)&"]]></com:Description>" 
                            strXML = strXML & "</com:Item>"
                            strXML = strXML & "<com:BasePrice>"
                            strXML = strXML & "<com:PriceAmount currencyID='"&valutaISO&"'>"& replace(matEnhPris, ",", "") &"</com:PriceAmount>" 
                            strXML = strXML & "</com:BasePrice>"
                            strXML = strXML & "</com:InvoiceLine>"
                            end if
                                    
                    'Response.Write "matsubtotal_fakgk(m)" & matsubtotal_fakgk(m) & "<br>"
                                    
                   h = h + 1
                   am = am + 1
                                    

                   if cint(showmatasgrp) = 2 then
                    fasebelialt = fasebelialt + matBelob
                   end if

                if m = 0 then    
                antalmatthis = antalmatthis + 1
                end if

                   antalMatTot = (antalMatTot*1) + (matAntal*1)  

                oRec2.movenext
                wend
                oRec2.close


                if cint(showmatasgrp) = 2 AND am <> 0 then 'luft før næste aktlinje
                   'strFakmat = strFakmat &"<tr><td colspan=20><br>&nbsp;</td></tr>"
                end if

                select case lto
                case "nt"
                  strFakmat = strFakmat &"<tr><td colspan=3>"& formatnumber(antalMatTot, 2) &"</td><td colspan=5>&nbsp;</td></tr>"
                end select

end sub

%>
