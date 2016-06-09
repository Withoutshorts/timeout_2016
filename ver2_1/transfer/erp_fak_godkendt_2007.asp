<%
response.buffer = true
%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"--> 
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/functions_inc.asp"--> 



<script>

function popUp(URL,width,height,left,top) {
		window.open(URL, 'navn', 'left='+left+',top='+top+',toolbar=0,scrollbars=1,location=0,statusbar=1,menubar=0,resizable=1,width=' + width + ',height=' + height + '');
	}

</script>

<%

a = 0
dim aval
redim aval(a)

'** Bruges til PDF visning ***
if len(request("nosession")) <> 0 then
nosession = request("nosession")
else
nosession = 0
end if


thisfile = "fak_godkendt.asp"


if len(trim(session("user"))) = 0 AND cint(nosession) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	
    if len(trim(request("vans"))) <> 0 then
	vans = request("vans")
    else
    vans = 0
    end if
	
	
	if len(trim(request("media"))) <> 0 then
	media = request("media")
	else
	media = "n"
	end if

    'Response.Write media
    'Response.flush
	
	select case media
	case "j" 'print 
	    
	    %>
	    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
	    <%
	
	case "xml"
	    
	    session.LCID = 1033
	    'Response.Charset = "utf-8"
	    
        strXML = "<?xml version=""1.0"" encoding=""UTF-8""?>"
        strXML = strXML &"<Invoice xmlns=""http://rep.oio.dk/ubl/xml/schemas/0p71/pie/"""_
        &" xmlns:com=""http://rep.oio.dk/ubl/xml/schemas/0p71/common/"""_
        &" xmlns:main=""http://rep.oio.dk/ubl/xml/schemas/0p71/maindoc/"""_
        &" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"""_
        &" xsi:schemaLocation=""http://rep.oio.dk/ubl/xml/schemas/0p71/pie/ "_
        &" http://rep.oio.dk/ubl/xml/schemas/0p71/pie/piestrict.xsd"">"
  
	    
	case else
	
	    %>
	    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	    <script>
	        $(document).ready(function () {

	            $("#oioxml").click(function () {

	                var xmlurl = $("#xmlurl").val()
	                var showvans = $("#showvans").val()
	                //alert(xmlurl)

	                function show_confirm() {
	                    var r = confirm("Ønsker du at afsende en OIO XML faktura til kunden via E-posthus (klik Ok), eller vil du blot gemme en ny kopi af OIO XML filen (klik Cancel)?");
	                    if (r == true) {
	                        window.open(xmlurl + "&vans=1", '_blank');

	                    }
	                    else {
	                        window.open(xmlurl + "&vans=0", '_blank');
	                    }
	                }

	                if (showvans == 1) {

	                    show_confirm()

	                } else {

	                    window.open(xmlurl + "&vans=0", '_blank');

                    }

	                //confirm("Ønsker du at afsende en OIO XML faktura til kunden via E-posthus, eller vil du blot gemme en ny kopi af OIO XML filen?")
	            });

	        });

        </script>
        
        <%
	    
	end select
	
	%>
	<!--#include file="inc/erp_fak_layout_inc_2007.asp"-->
	<%
	
	
	
	id = request("id")
	showmoms = 0
	
	if len(request("jobid")) <> 0 then
	jobid = request("jobid")
	else
	jobid = 0
	end if
	
	if len(request("aftid")) <> 0 then
	aftid = request("aftid")
	else
	aftid = 0
	end if
	
	'*** aftale side variable ****
	
	visjoblog = 0
	vismatlog = 0
	
	
	
	
	'*** Faktura info ****
	if jobid <> 0 then
	strSQL_sel = " jobid, jobnavn, jobnr, rekvnr, "
	strSQL_lftJ = " job j ON (j.id = jobid) "
	else
	strSQL_sel = " s.navn AS aftnavn, s.aftalenr, "
	strSQL_lftJ = " serviceaft s ON (s.id = aftaleid) "
	end if
	
	strSQL = "SELECT f.editor, tidspunkt, fid, faknr, fakdato, b_dato, timer, beloeb, f.kommentar, f.betalt, "_
	&" tidspunkt, fakadr, att, f.faktype, konto, modkonto, "& strSQL_sel &" "_
	&" vismodland, vismodatt, vismodtlf, vismodcvr, visafstlf, visafsemail,"_
	&" visafsswift, visafsiban, visafscvr, moms, enhedsang, f.varenr, jobbesk, "_
	&" visjoblog, visrabatkol, vismatlog, rabat, "_
	&" visjoblog_timepris, visjoblog_enheder, visjoblog_mnavn, "_
	&" visafsfax, betalt, f.valuta, f.kurs, v.valutakode AS valutaISO, f.sprog, "_
	&" f.istdato, momskonto, visperiode, istdato2, brugfakdatolabel, "_
	&" vorref, fak_ski, momssats, modtageradr, showmatasgrp, hidesumaktlinier, "_
	&" sideskiftlinier, f.subtotaltilmoms, labeldato, fak_abo, fak_ubv, visikkejobnavn, hidefasesum, hideantenh, medregnikkeioms"_
	&" FROM fakturaer f "_
	&" LEFT JOIN "& strSQL_lftJ &""_
	&" LEFT JOIN valutaer v ON (v.id = f.valuta) WHERE fid = "& id
	
	'Response.write strSQL
	'Response.flush
	
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	
	fak_valuta = oRec("valuta")
	fak_kurs = oRec("kurs")
	
	valutaISO = oRec("valutaISO")
	kurs = oRec("kurs")
	
	strEditor = oRec("editor")
	fakTidspunkt = oRec("tidspunkt")
	intFaktype = oRec("faktype")
	intKundeid = oRec("fakadr")
	intAtt = oRec("att")
	
	'intKonto = oRec("konto")
	'intModKonto = oRec("modkonto")
	
	'if oRec("momskonto") <> 1 then
	'momsKonto = intModKonto
	'else
	'momsKonto = intKonto
	'end if
	
    intGodkendt = oRec("betalt")
	
	visjoblog = oRec("visjoblog")
	
	visjoblog_enheder = oRec("visjoblog_enheder")
	visjoblog_mnavn = oRec("visjoblog_mnavn")
	visjoblog_timepris = oRec("visjoblog_timepris")
	
	vismatlog = oRec("vismatlog")
	
	visrabatkol = oRec("visrabatkol")
	
	momssats = oRec("momssats")
	
	intFaktureretTimer = oRec("timer")
	intFaktureretBelob = oRec("beloeb")
	
	subtotalTilmoms = oRec("subtotaltilmoms")
	
	if intFaktype = 1 AND media <> "xml" then
	intFaktureretBelob = -(intFaktureretBelob)
	subtotalTilmoms = -(subtotalTilmoms)
	end if
	
   
	
	    if jobid <> 0 then
	    strKomm = oRec("kommentar")
	    strJobnavn = oRec("jobnavn")
	    intJobnr = oRec("jobnr")
	    strJobBesk = oRec("jobbesk")
    	
	    if trim(len(oRec("rekvnr"))) <> 0 then
	    rekvnr = oRec("rekvnr")
	    else
	    rekvnr = "0"
	    end if
    	
	    else
    	
	     rekvnr = "0"
	     strAftNavn = oRec("aftnavn")
	     strAftVarenr = oRec("varenr")
	     intAftnr = oRec("aftalenr")
    	    
	        if cdate(oRec("fakdato")) > cdate("1/5/2007") then
	        strKomm = oRec("kommentar")
	        strJobBesk = oRec("jobbesk")
	        else
	        strJobBesk = oRec("kommentar")
	        strKomm = "Netto Kontant 8 dage."
	        end if
    	
	    end if
	
	
	
	forfaldsdato = oRec("b_dato")
	varFaknr = oRec("faknr")
	istDato = oRec("istdato")
	istDato2 = oRec("istdato2")
	brugfakdatolabel = oRec("brugfakdatolabel")
	
	'** Label eller system dato ***'
	if cint(brugfakdatolabel) <> 1 then
	fakdato = oRec("fakdato") 
	else
	fakdato = oRec("labeldato") 'istDato2
	end if
	
    sysFakdato = oRec("fakdato")
	
	intVismodland = oRec("vismodland")
	intVismodatt = oRec("vismodatt")
	intVismodtlf = oRec("vismodtlf")
	intVismodcvr = oRec("vismodcvr")
	
	
	intVisafstlf = oRec("visafstlf")
	intVisafsemail = oRec("visafsemail")
	intVisafsswift = oRec("visafsswift")
	intVisafsiban = oRec("visafsiban")
	intVisafscvr = oRec("visafscvr")
	intVisafsfax = oRec("visafsfax")
	
	'if oRec("valuta") <> 1 then
	'intMoms = oRec("moms") * (100/kurs)
	'else
	intMoms = oRec("moms") 
	'end if
	
	if intFaktype = 1 AND media <> "xml" then
	intMoms = -(intMoms)
	end if
	
	select case oRec("enhedsang")
	case -1
	enhed = "" 'Ingen (skjul)
	case 1	
	enhed = txt_019 '"Stk. pris"
	case 2
	enhed = txt_020 '"Enhedspris"
	case else
	enhed = txt_018 '"Timepris"
	end select
	
	enhedsangFak = oRec("enhedsang")
	
	intAftRabat = oRec("rabat")
	intStatus = oRec("betalt") 
	sprog = oRec("sprog")
	
	visPeriode = oRec("visperiode")
	
	modtageradr = oRec("modtageradr")
	
	vorref = oRec("vorref")
	fak_ski = oRec("fak_ski")
	fak_abo = oRec("fak_abo")
	fak_ubv = oRec("fak_ubv")
	
	showmatasgrp = oRec("showmatasgrp")
	
	hidesumaktlinier = oRec("hidesumaktlinier")
	sideskiftlinier = oRec("sideskiftlinier") 
	
	visikkejobnavn = oRec("visikkejobnavn")
    hidefasesum = oRec("hidefasesum")

    hideantenh = oRec("hideantenh")

    medregnikkeioms = oRec("medregnikkeioms")
	
	if media = "xml" then
	
	session.LCID = 1033
	
	strXML = strXML & "<com:ID>"& varFaknr &"</com:ID>" 
	
	if month(fakdato) < 10 then
            monthFakd = "0"& month(fakdato)
            else
            monthFakd = month(fakdato)
            end if
            
            if day(fakdato) < 10 then
            dayFakd = "0"& day(fakdato)
            else
            dayFakd = day(fakdato)
            end if
            
            fakdatoXML = year(fakdato) &"-"& monthFakd &"-"& dayFakd
	
	strXML = strXML & "<com:IssueDate>"& fakdatoXML &"</com:IssueDate>" 
    
    if intFaktype = 0 then
    pieTyp = "PIE" 'Faktura
    else
    pieTyp = "PCM" 'Kreditnota
    end if
    
    strXML = strXML & "<com:TypeCode>"& pieTyp &"</com:TypeCode>" 
    strXML = strXML & "<main:InvoiceCurrencyCode>"& valutaISO &"</main:InvoiceCurrencyCode>" 
    
        if jobid <> 0 then
           
           if cint(visikkejobnavn) <> 1 then
           xmlNote = "Job: " & strJobnavn & "("& intJobnr &")"
           else
           xmlNote = ""
           end if
           
           if len(trim(strJobBesk)) <> 0 then
	       xmlNote = xmlNote & "<br>"&  strJobBesk
	       end if
		
	    else
	   
	   xmlNote = strAftNavn & "(" & intAftnr & ")<br> Aftale nr.: "& strAftVarenr
	    
	    end if
    
    'call utf_format(xmlNote)
    'xmlNote = utf_formatTxt
          
    call htmlparseCSV(xmlNote)
    xmlNote = htmlparseCSVtxt
    
    xmlNote = replace(xmlNote, "''", "")
    
    strXML = strXML & "<com:Note><![CDATA["&EncodeUTF8(xmlNote)&"]]></com:Note>"
    
    
    end if
	
	
	end if
	oRec.close 
	
	
	
	
	
	%>
	<!--#include file="../inc/xml/erp_fak_xml_inc.asp"--> 
	<%
	

	' *** Modtager ****
	strSQL = "SELECT Kid, kkundenavn, kkundenr, adresse, postnr, city, land, telefon, cvr, ean FROM kunder WHERE Kid =" & intKundeid		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			
			intKnr = oRec("kkundenr")
			strKnavn = oRec("kkundenavn")
			strKadr = oRec("adresse")
			strKpostnr = oRec("postnr")
			strBy = oRec("city")
			strLand = oRec("land")
			intKid = oRec("Kid")
			intCVR = oRec("cvr")
			intTlf = oRec("telefon")
			ean = trim(oRec("ean"))
			
			if media = "xml" then
			'*** XML format kan kun sendes til oprindelig kunde **'
			
			session.LCID = 1033
			
			'call utf_format(strKnavn)
            'strKnavn = utf_formatTxt
            
          
          
          'Response.Write strKadr
          'Response.end
          
            
			
			call htmlparseCSV(strKadr)
            strKadr = htmlparseCSVtxt
          
            'call utf_format(strKadr)
            'strKadr = utf_formatTxt
          
            'call utf_format(strBy)
            'strBy = utf_formatTxt
			
			
		    strXML = strXML & "<com:BuyersReferenceID schemeID=""EAN"">"&ean&"</com:BuyersReferenceID>"
		    strXML = strXML & "<com:ReferencedOrder>"
            strXML = strXML & "<com:BuyersOrderID>"&intJobnr&"</com:BuyersOrderID>" 
            strXML = strXML & "<com:IssueDate>"&fakdatoXML&"</com:IssueDate>" 
            strXML = strXML & "</com:ReferencedOrder>"
		    strXML = strXML & "<com:BuyerParty>"
            strXML = strXML & "<com:ID schemeID=""CVR"">"&intCVR&"</com:ID>"
            strXML = strXML & "<com:PartyName>"
            strXML = strXML & "<com:Name><![CDATA["&EncodeUTF8(strKnavn)&"]]></com:Name>"
            strXML = strXML & "</com:PartyName>"
            strXML = strXML & "<com:Address>"
            strXML = strXML & "<com:ID>Fakturering</com:ID>" 
            strXML = strXML & "<com:Street><![CDATA["&EncodeUTF8(strKadr)&"]]></com:Street>"
           'strXML = strXML & "<com:HouseNumber>0</com:HouseNumber>" 
           strXML = strXML & "<com:CityName><![CDATA["&EncodeUTF8(strBy)&"]]></com:CityName>"
          strXML = strXML & "<com:PostalZone>"& strKpostnr &"</com:PostalZone>" 
          strXML = strXML &"<com:Country><com:Code listID=""ISO 3166-1"">DK</com:Code></com:Country>" 
          strXML = strXML & "</com:Address>"
          end if
         

			
		end if
		
		oRec.close
		
		
    
    
    finummer = ""
    
   
    select case lto
	case "dencker" ', "intranet - local" ', "optionone"
	
	    select case lto 
	    case "dencker" '"intranet - local"
	    kreditorFi = "80200552"
	    case else
	    kreditorFi = "00000000"
	    end select
    
    
    '*** Debitor id skal være 8 cifre **'
    lenintKnr = len(trim(intKnr))
    select case lenintKnr
    case 8 
    fiKnr = trim(intKnr)
    case else
    fiKnr = "00000000"
    end select
    
	finummer = finummer & fiKnr 
	
	'Response.Write "lenvarFaknr" & varFaknr & "<br><br>"
	
	
	'*** Fakturanr ID skal være 6 cifre **'
	lenvarFaknr = len(trim(varFaknr))
    select case lenvarFaknr
    case 6 
    fiFaknr = trim(varFaknr)
    case 5 
    fiFaknr = "0"&trim(varFaknr)
    case 4 
    fiFaknr = "00"&trim(varFaknr)
    case 3 
    fiFaknr = "000"&trim(varFaknr)
    case 2 
    fiFaknr = "0000"&trim(varFaknr)
    case 1 
    fiFaknr = "00000"&trim(varFaknr)
    
    case else
    fiFaknr = right(varFaknr,6)
    end select
	
	finummer = finummer & fiFaknr 
	
	
	
	'** Modulus tjek nummer ***'
	betTVsumCounter = 0
	for x = 1 to 14
	 
	 betTVsum = 0
	 betIndet = mid(finummer,x,1)
	 
	 select case x
	 case 1,3,5,7,9,11,13
	 betIndet = (betIndet*1)
	 case else
	 betIndet = (betIndet*2)
	 end select
	 
	 
	 'Response.Write betIndet & ": "
	    
	    for t = 1 to len(betIndet)
	    betTVsum = betTVsum + mid(betIndet,t,1)
	    betTVsumCounter = betTVsumCounter + betTVsum 
	    next
	    
	    'Response.Write betTVsum & "<br> "
	    
	next
	
	
	modulus = 0
	modulus_division = (betTVsumCounter/10)
	
	komma = 0
	komma = instr(modulus_division, ",")
	
	'Response.Write "<br>komma: " & komma
	
	if komma <> 0 then
	modulus_division = mid(modulus_division,1,komma-1)
	else
	modulus_division = modulus_division 
	end if
	
	'Response.Write "<br>betTVsumCounter:" & betTVsumCounter
	'Response.Write "<br>modulus_division:" & modulus_division
	
	
	modulus = betTVsumCounter - (10 * modulus_division)
	kontrolciffer = (10 - modulus)
	
	'Response.Write "<br>modulus:" & modulus
	'Response.Write "<br>kontrolciffer:" & kontrolciffer
	
	finummer = "+71<" & finummer & kontrolciffer & "+"& kreditorFi
    
    end select
    
    
    kid = intKid
		
	'*** Att ******
	select case intAtt
	case "0"
	showAtt = ""
	case "991"
	showAtt = "Den økonomi ansvarlige"
	case "992"
	showAtt = "Regnskabs afd."
	case "993"
	showAtt = "Administrationen"
	case else
	
	strSQL3 = "SELECT id, navn, email FROM kontaktpers WHERE id="& intAtt
	oRec3.open strSQL3, oConn, 3
	if not oRec3.EOF then
	showAtt = oRec3("navn")
	showAttEmail = oRec3("email")
	end if
	oRec3.close
	
	if len(showAtt) <> 0 then
	showAtt = showAtt
	else
	showAtt = ""
	end if
	
	end select	
	      
	      
	      if media = "xml" then  
	      
	      'if len(trim(showAttEmail)) <> "" then
	      'xmlAtt = showAttEmail
	      'else
	      xmlAtt = showAtt
	      'end if
	      
	      'call utf_format(xmlAtt)
          'xmlAtt = utf_formatTxt
	      
	      strXML = strXML &"<com:BuyerContact>"
          strXML = strXML &"<com:ID><![CDATA["&EncodeUTF8(xmlAtt)&"]]></com:ID>" 
          strXML = strXML &"</com:BuyerContact>"
          strXML = strXML &"</com:BuyerParty>"
          end if
	
	
	'*** Afsender ***
	
		strSQL = "SELECT adresse, postnr, city, land, telefon, email, "_
		&" regnr, kkundenavn, kontonr, cvr, bank, swift, iban, kid, kkundenr, "_
		&" fax FROM kunder WHERE useasfak = 1"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then 
			yourbank = oRec("bank")
			yourRegnr = oRec("regnr")
			yourKontonr = oRec("kontonr")
			yourCVR = oRec("cvr")
			yourNavn = oRec("kkundenavn")
			yourAdr = oRec("adresse")
			yourPostnr = oRec("postnr")
			yourCity = oRec("city")
			yourLand = oRec("land")
			yourEmail = oRec("email")
			yourTlf = oRec("telefon")
			yourSwift = oRec("swift")
			yourIban = oRec("iban")
			afskid = oRec("kid")
			yourFax = oRec("fax")
			yourKnr = oRec("kkundenr")
			
			if media = "xml" then
			  
			  session.LCID = 1033
			  
			  
			   call htmlparseCSV(yourAdr)
              yourAdr = htmlparseCSVtxt
			  
			  'call utf_format(yourNavn)
              'yourNavn = utf_formatTxt
              
              'call utf_format(yourAdr)
              'yourAdr = utf_formatTxt
          
             
              
               'strXML = strXML &"<com:HouseNumber>0</com:HouseNumber>" 
              
              'call utf_format(yourCity)
              'yourCity = utf_formatTxt
			  
			  strXML = strXML &"<com:SellerParty>"
              strXML = strXML &"<com:ID schemeID=""CVR"">"&yourCVR&"</com:ID>" 
              strXML = strXML &"<com:PartyName>"
              strXML = strXML &"<com:Name><![CDATA["&EncodeUTF8(yourNavn)&"]]></com:Name>" 
              strXML = strXML &"</com:PartyName>"
              strXML = strXML &"<com:Address>"
              strXML = strXML &"<com:ID>Betaling</com:ID>"
              strXML = strXML &"<com:Street><![CDATA["&EncodeUTF8(yourAdr)&"]]></com:Street>" 
              strXML = strXML &"<com:CityName><![CDATA["&EncodeUTF8(yourCity)&"]]></com:CityName>" 
              strXML = strXML &"<com:PostalZone>"& yourPostnr &"</com:PostalZone>"
              strXML = strXML &"<com:Country><com:Code listID=""ISO 3166-1"">DK</com:Code></com:Country>" 
              strXML = strXML &"</com:Address>"
              strXML = strXML &"<com:PartyTaxScheme>"
              strXML = strXML &"<com:CompanyTaxID schemeID=""CVR"">"& yourCVR &"</com:CompanyTaxID>" 
              strXML = strXML &"</com:PartyTaxScheme>"
              strXML = strXML &"</com:SellerParty>"
              
                strXML = strXML &"<com:PaymentMeans>"
                strXML = strXML &"<com:TypeCodeID>null</com:TypeCodeID>"
                
                
            if month(forfaldsdato) < 10 then
            monthffFakd = "0"& month(forfaldsdato)
            else
            monthffFakd = month(forfaldsdato)
            end if
            
            if day(forfaldsdato) < 10 then
            dayffFakd = "0"& day(forfaldsdato)
            else
            dayffFakd = day(forfaldsdato)
            end if
            
            fakforfaldsdatoXML = year(forfaldsdato) &"-"& monthffFakd &"-"& dayffFakd
                
                strXML = strXML & "<com:PaymentDueDate>"& fakforfaldsdatoXML &"</com:PaymentDueDate>" 
                
                PaymentChannelCode = "KONTOOVERFØRSEL"
                'call utf_format(PaymentChannelCode)
                'PaymentChannelCode = utf_formatTxt
                
                strXML = strXML &"<com:PaymentChannelCode>"&EncodeUTF8(PaymentChannelCode)&"</com:PaymentChannelCode>"
                strXML = strXML &"<com:PayeeFinancialAccount><com:ID>"&yourKontonr&"</com:ID>"
                strXML = strXML &"<com:TypeCode>BANK</com:TypeCode><com:FiBranch>"
                strXML = strXML &"<com:ID>"&yourRegnr&"</com:ID>"
                strXML = strXML &"<com:FinancialInstitution>"
                
                'call utf_format(yourbank)
                'yourbank = utf_formatTxt
                
                strXML = strXML &"<com:ID>null</com:ID><com:Name>"&EncodeUTF8(yourbank)&"</com:Name>"
                strXML = strXML &"</com:FinancialInstitution>"

                strXML = strXML &"</com:FiBranch></com:PayeeFinancialAccount>"
                strXML = strXML &"<com:PaymentAdvice>"
                strXML = strXML &"<com:AccountToAccount>"
                strXML = strXML &"<com:PayerNote></com:PayerNote>"
                strXML = strXML &"<com:PayeeNote></com:PayeeNote>"
                strXML = strXML &"</com:AccountToAccount>"
                strXML = strXML &"</com:PaymentAdvice>"
                strXML = strXML &"</com:PaymentMeans>"
              
              end if
			
		end if
		oRec.close
		
		
		'**** XML totol sum ***'
		if media = "xml" then
		  
		  
		  session.LCID = 1033
		  
		  strXML = strXML & "<com:TaxTotal>"
          strXML = strXML & "<com:TaxTypeCode>VAT</com:TaxTypeCode>" 
          
          strXML = strXML & "<com:TaxAmounts>"
          strXML = strXML & "<com:TaxableAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(subtotalTilmoms,2), ",", "") &"</com:TaxableAmount> "
          strXML = strXML & "<com:TaxAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(intMoms,2), ",", "") &"</com:TaxAmount>" 
          strXML = strXML & "</com:TaxAmounts>"
          
          strXML = strXML & "<com:CategoryTotal>"
          strXML = strXML & "<com:RateCategoryCodeID>VAT</com:RateCategoryCodeID>"
          strXML = strXML & "<com:RatePercentNumeric>"&momssats&"</com:RatePercentNumeric>" 
          strXML = strXML & "<com:TaxAmounts>"
          strXML = strXML & "<com:TaxableAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(subtotalTilmoms,2), ",", "") &"</com:TaxableAmount>" 
          strXML = strXML & "<com:TaxAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(intMoms,2), ",", "") &"</com:TaxAmount>" 
          strXML = strXML & "</com:TaxAmounts>"
          strXML = strXML & "</com:CategoryTotal>"
          strXML = strXML & "</com:TaxTotal>"
          
          strXML = strXML & "<com:LegalTotals>"
          strXML = strXML & "<com:LineExtensionTotalAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(intFaktureretBelob,2), ",", "") &"</com:LineExtensionTotalAmount>" 
          strXML = strXML & "<com:ToBePaidTotalAmount currencyID='"&valutaISO&"'>"& replace(formatnumber(intFaktureretBelob+intMoms, 2), ",", "") &"</com:ToBePaidTotalAmount>"
          strXML = strXML & "</com:LegalTotals>"
		
		end if
		
		
		
		'*********************************************************************************'
		'**** Fakturaen ***'
		'*********************************************************************************'
	
		
		
		
		if media = "n" then
		bd = 1
		else
		bd = 0
		end if 
		
		
		'*** sideskift ****'
		%>
		
		
		<div id="fakturaside" style="position:relative; left:10px; top:4px; width:640px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; padding:10px 10px 10px 10px;">
		
		
		
		<% 
		'*** fak header layout *******
		call fakheader() 
	
		'*** Maintable layout ********
		call maintable
		
		'*** Modtager boks layout ****
		'call modtager_layout()
		
		'*** Maintable 2 layout ********
		call maintable_2
		
		'*** Afsender boks layout ****										
		'call afsender_layout()
		
		'*** Maintable 3 layout ********
		call maintable_3
		
												
		'*** Vedr. (upspecificering top) ***
		'*** Vise hvis der findes jobnavn (visjobnavn ikke er slået fra), jobbesk, reknr eller det er en fatale der faktureres
		if (cint(visikkejobnavn) = 0 OR len(trim(strJobBesk)) <> 0 OR rekvnr <> "0" AND jobid <> 0) OR jobid = 0 then
		call vedr
		else
		Response.Write "<tr><td colspan=10><br><br>&nbsp;</td></tr>"
		end if
		
		
		
		
		
		'**** Udspecificering ******
		showmomsfriTxt = 0
		strFakdet = ""
		'strMedarbFakdet = ""
		strFakmat = ""
		intBelobforMoms = 0
		
		dim strFakmat, matsubtotal
		redim strFakmat(1), matsubtotal(1)
		
		
		'if jobid <> 0 then
		    
		    '** hidesumaktlinier = vis kun totalbeløb **'
		    if hidesumaktlinier <> 1 then
					
					
					'***************************************'
					'*** Fakturering af job /Aktiviteter ***'
					h = 1
					antalAktlinier = 0
					
					strSQL = "SELECT antal, beskrivelse, aktpris, fakid, "_
					&" enhedspris, aktid, showonfak, rabat, enhedsang, fd.valuta, fd.kurs, v.valutakode, momsfri, fase "_
					&" FROM faktura_det fd "_
					&" LEFT JOIN valutaer v ON (v.id = fd.valuta) "_
					&" WHERE fakid = "& id &" AND showonfak = 1 ORDER BY fak_sortorder, beskrivelse"
					
					oRec.open strSQL, oConn, 3 
					while not oRec.EOF 
						
						fd_valuta = oRec("valuta")
						fd_kurs = oRec("kurs")
						fd_valutaKode = oRec("valutakode")
						
						'** Pris på sum akt. i den valuta der er valgt på akt. **'
						fd_Ehpris = oRec("enhedspris")
						fd_Pris = oRec("aktpris")
						
						'** Pris på sum akt. i den valuta der er valgt på faktura. **'
						if fak_kurs <> fd_kurs then
						'aktEhpris = oRec("enhedspris") * (fak_kurs/fd_kurs)
						'else
						aktEhpris = formatnumber(oRec("enhedspris") * (fd_kurs/fak_kurs), 2)
						else
						aktEhpris = formatnumber(fd_Ehpris, 2)
						end if
						
						
						fd_momsfri = oRec("momsfri")
                        fd_fase = oRec("fase")
						
						'** aktpris er altid i dent valuta der er angivet på faktura ***'
						if len(trim(oRec("aktpris"))) <> 0 then
						aktPris = formatnumber(oRec("aktpris"),2) 
						else
						aktPris = formatnumber(0, 2)
						end if
						
						'aktEhpris = oRec("enhedspris") * (100/fak_kurs)
						'aktPris = oRec("aktpris") * (100/fak_kurs)
						
						
						if intFaktype = 1 AND media <> "xml" then
						    fd_Ehpris = -(formatnumber(fd_Ehpris, 2))
						    fd_Pris = -(formatnumber(fd_Pris, 2))
						    aktPris = -(formatnumber(aktPris, 2))
						    aktEhpris = -(formatnumber(aktEhpris, 2))
						end if
						
						rabat = oRec("rabat")
						
						select case oRec("enhedsang")
						case 0
						enhedsang = txt_035 '"pr. time"
						case 1
						enhedsang = txt_036 '"pr. stk."
						case 2
						enhedsang = txt_037 '"pr. enhed"
						case 3
						enhedsang = txt_054 '"pr. km."
						end select
						
					
					
				    
				     
				    if cint(h) = cint(sideskiftlinier) AND cint(sideskiftlinier) <> 0 then
				    
				    '** LUFT i toppen på side 2 /3 
				    select case lto
				    case "dencker", "tooltest", "outz", "intranet - local", "optionone"
				    sidetoluft = 100
				    case else
				    sidetoluft = 20
				    end select
				   
				    
				    strFakdet = strFakdet &"<tr>"_
				    &"<td valign=top width=40 align=right style='padding-right:5px; height:"& sidetoluft &"px;' colspan=8>"_
				    &"<img src=""../ill/blank.gif"" height='"& sidetoluft &"' width=0 border=0 style=""page-break-before: always;""></td></tr>"
				   
				    h = 0
				    
				    end if
				   
				    
				    
				    '*** Momsfri ***'
					if cint(fd_momsfri) = 1 then
					fd_momsfriVal = "*"
					showmomsfriTxt = 1
					else
					fd_momsfriVal = ""
					end if
				    
				    if len(trim(oRec("antal"))) <> 0 then
				    tAntal = formatnumber(oRec("antal"), 2)
				    else
				    tAntal = formatnumber(0, 2)
				    end if
				    
				    if len(trim(oRec("beskrivelse"))) <> 0 then
				    aktBesk = oRec("beskrivelse")
				    else
				    aktBesk = ""
				    end if

                    if cint(hidefasesum) <> 1 then
                            

                            '** Gl. fase sum **'
                            if lcase(lastfdfase) <> lcase(fd_fase) AND len(trim(lastfdfase)) <> 0 AND faseshowtot = 1 then

                             strFakdet = strFakdet &"<tr>"
                             
                             if cint(hideantenh) <> 1 then
                             strFakdet = strFakdet &"<td colspan=2>&nbsp;</td>"
                             end if
				             
                             strFakdet = strFakdet &"<td valign=top colspan=5 style='padding:1px 5px 10px 5px; color:#999999;'>"& replace(lastfdfase, "_", " ") &" i alt:</td>"_
                             &"<td align=right style='padding:1px 0px 10px 5px; color:#999999;'>"&formatnumber(fasebelialt, 2)&" "& valutaISO &"</td></tr>"


                            end if

                            '** Ny fase **'
                             if lcase(lastfdfase) <> lcase(fd_fase) AND len(trim(fd_fase)) <> 0 then
                             strFakdet = strFakdet &"<tr>"

                             if cint(hideantenh) <> 1 then
                             strFakdet = strFakdet &"<td colspan=2>&nbsp;</td>"
                             end if

				             strFakdet = strFakdet &"<td valign=top colspan=6 style='padding:5px 5px 0px 5px;'><b>"& replace(fd_fase, "_", " ") &"</b></td></tr>"

                             faseshowtot = 1
                             fasebelialt = 0
                             end if

                    end if
				    
				    call jq_format(aktBesk)
				    aktBesk = jq_formatTxt 
				    
				    strFakdet = strFakdet &"<tr>"

                    '** Skjul antal og enh kolonner **'
                    if cint(hideantenh) <> 1 then
                    strFakdet = strFakdet &"<td valign=top align=right style='padding-right:5px; width:40px;'>"& tAntal &"</td>"_
					&"<td valign=top style='padding-right:5px; width:40px;'>"& enhedsang &"</td>"
                    end if

					strFakdet = strFakdet &"<td colspan=2 valign=top style='padding-left:5px;'>"& aktBesk &"</td>"_
					&"<td valign=bottom align=right style='padding-right:2px;'>"
					
					
					
					
					'*** Enh. pris og Valuta på sum akt. linie ***'
					if cint(fak_valuta) <> cint(fd_valuta) then
					strFakdet = strFakdet & formatnumber(fd_Ehpris, 2) &" "& fd_valutaKode & " ~"
				    else
					strFakdet = strFakdet & "&nbsp;"
					end if
					
					strFakdet = strFakdet &"</td>"
					
                    '** Skjul antal og enh kolonner **'
                    if cint(hideantenh) <> 1 then
                    '*** Enh. pris og Valuta på fak. ***'
					strFakdet = strFakdet &"<td valign=bottom align=right style='padding-right:2px;'>"& aktEhpris &" "& valutaISO 
					strFakdet = strFakdet &"</td>"
                    else
                    strFakdet = strFakdet &"<td>&nbsp;</td>"
                    end if
					
					
					'*** Rabat ***'
					if cint(visrabatkol) = 1 then
					strFakdet = strFakdet &"<td valign=bottom align=right style='padding-right:10;'>"
					        
					        if cdbl(rabat) <> 0 then
					        rbtthis = (rabat * 100) 
					        strFakdet = strFakdet & rbtthis &" %</td>"
					        else
					        strFakdet = strFakdet &"&nbsp;</td>"
					        end if
					else
					strFakdet = strFakdet &"<td>&nbsp;</td>"
					        
					end if
					
					
					
					'strFakdet = strFakdet &"<td valign=top align=right style='padding-right:5;'>"& fd_momsfriVal &"</td>"
					
					
					
					'*** Total beløb pr. linie ***'
					strFakdet = strFakdet &"<td valign=bottom align=right>"& fd_momsfriVal &" "& aktPris &" "& valutaISO &"</td>"_
					&"</tr>"
					
					'aktsubtotal = aktsubtotal + (aktPris)
					
					'Response.Write "aktsubtotal:" & aktsubtotal & "<br>"
					
                    lastfdfase = fd_fase
                    fasebelialt = fasebelialt + aktPris 

					h = h + 1		
			        rbtthis = 0
					
					
					  if media = "xml" then
					  
					  session.LCID = 1033
					  
					  'call utf_format(aktBesk)
                      'aktBesk = utf_formatTxt
					  
					  'call utf_format(enhedsang)
                      'enhedsang = utf_formatTxt
					  
					  strXML = strXML & "<com:InvoiceLine>"
                      strXML = strXML & "<com:ID>"&h&"</com:ID>" 
                      strXML = strXML & "<com:InvoicedQuantity unitCode='"&EncodeUTF8(enhedsang)&"' unitCodeListAgencyID=""n/a"">"& replace(tAntal, ",", "")&"</com:InvoicedQuantity>" 
                      strXML = strXML & "<com:LineExtensionAmount currencyID='"&valutaISO&"'>"& replace(aktPris, ",", "") &"</com:LineExtensionAmount>" 
                      strXML = strXML & "<com:Item>"
                      strXML = strXML & "<com:ID>0</com:ID>" 
                      strXML = strXML & "<com:Description><![CDATA["&EncodeUTF8(oRec("beskrivelse"))&"]]></com:Description>" 
                      strXML = strXML & "</com:Item>"
                      strXML = strXML & "<com:BasePrice>"
                      strXML = strXML & "<com:PriceAmount currencyID='"&valutaISO&"'>"& replace(aktEhpris, ",", "") &"</com:PriceAmount>" 
                      strXML = strXML & "</com:BasePrice>"
                      strXML = strXML & "</com:InvoiceLine>"
                      end if
						
								
								ehpris = replace(oRec("enhedspris"), ",", ".")
								
								if media <> "xml" then
								
                                if jobid <> 0 then 'kun jobfakturer kan indeholde medarb. linier

								'**** Medarbejder timer der vises på media ***'
								strSQL2 = "SELECT fakid, aktid, mid, fak, venter, tekst, enhedspris, beloeb, "_
								&" showonfak, medrabat, enhedsang, fms.valuta, fms.kurs, valutaKode "_
								&" FROM fak_med_spec fms "_
								&" LEFT JOIN valutaer v ON (v.id = fms.valuta) "_
								&" WHERE fakid = "& id &" AND aktid = "& oRec("aktid") &" AND "_
								&" enhedspris = "& ehpris &" AND fms.valuta = "& fd_valuta &" AND showonfak = 1"
								'Response.write strSQL2
								'Response.flush
								oRec2.open strSQL2, oConn, 3 
								while not oRec2.EOF 
								
								fms_valuta = oRec("valuta")
						        fms_kurs = oRec("kurs")
						        fms_valutaKode = oRec("valutakode")
								
								
								fms_Ehpris = oRec2("enhedspris")
								fms_Belob = oRec2("beloeb")
								
								if fak_kurs <> fd_kurs then
						        'medEhpris = fms_Ehpris * (fak_kurs/fms_kurs)
						        'else
						        medEhpris = fms_Ehpris * (fms_kurs/fak_kurs)
						        else
						        medEhpris = fms_Ehpris
						        end if
								
								
								
								'** medarbejder beløb er omregnet til den valgte valuta **'
								'** på faktura inden den lægges i DB ***'
								
								 medBelob = oRec2("beloeb")
								'medEhpris = oRec2("enhedspris") * (100/fak_kurs)
								'medBelob = oRec2("beloeb") * (100/fak_kurs)
								
								
								if intFaktype = 1 AND media <> "xml" then
								fms_Ehpris = -(formatnumber(fms_Ehpris, 2))
								fms_Belob = -(formatnumber(fms_Belob, 2))
								medBelob = -(formatnumber(medBelob, 2))
								medEhpris = -(formatnumber(medEhpris, 2))
								end if
								
								medrabat = oRec2("medrabat")
								
								select case oRec("enhedsang")
						        case 0
						        enhedsang = txt_035 '"pr. time"
						        case 1
						        enhedsang = txt_036 '"pr. stk."
						        case 2
						        enhedsang = txt_037 '"pr. enhed"
						        case 3
						        enhedsang = txt_054 '"pr. km."
						        end select
        								
        					            '*** ÆØÅ **'
						                call jquery_repl_spec(enhedsang)
                                        enhedsang = jquerystrTxt
                                        
								
								
								    select case lto
				                    case "dencker", "tooltest", "outz", "intranet - local", "optionone"
                				    
				                    if h = 16 then '11
                				    
				                    strFakdet = strFakdet &"<tr>"_
				                    &"<td valign=top width=40 align=right style='padding-right:5px; height:100px;' colspan=8>"_
				                    &"<img src=""../ill/blank.gif"" height=100 width=0 border=0 style=""page-break-before: always;""></td></tr>"
                				   
				                    h = 0
                				    
				                    end if
				                    case else
                				    
				                    end select
								   
								   
								strFakdet = strFakdet &"<tr>"


                                '** Skjul antal og enh kolonner **'
                                if cint(hideantenh) <> 1 then
								strFakdet = strFakdet &"<td valign=top width=40 align=right style='padding-right:5px;'><font class=megetlillesort>"& formatnumber(oRec2("fak"), 2) &"</td>"_
								&"<td valign=top width=40 style='padding-right:5px;'><font class=megetlillesort>"& enhedsang &"</td>"
                                end if

								strFakdet = strFakdet &"<td colspan=2 valign=top style='padding-left:5px;'><font class=megetlillesort>" & oRec2("tekst") &"</td>"_
								&"<td valign=top align=right style='padding-right:2px;'><font class=megetlillesort>"
					
					            '*** Enh. pris og Valuta på medarb. linie ***'
					            if cint(fak_valuta) <> cint(fms_valuta) then
					            strFakdet = strFakdet & formatnumber(fms_Ehpris, 2) &" "& fms_valutaKode & " ~"
					            else
					            strFakdet = strFakdet & "&nbsp;"
					            end if
            					
					            strFakdet = strFakdet &"</td>"
								
                                 '** Skjul antal og enh kolonner **'
                                if cint(hideantenh) <> 1 then
								'*** Enh. pris og Valuta på medarb. fra Faktura valuta ***'
								strFakdet = strFakdet &"<td valign=top align=right style='padding-right:2px;'><font class=megetlillesort>"& formatnumber(medEhpris, 2) &" "& valutaISO &"</td>"
                                else
                                strFakdet = strFakdet &"<td>&nbsp;</td>"
                                end if
								
								
								'*** Rabat ***'
								if cint(visrabatkol) = 1 then
								strFakdet = strFakdet &"<td valign=top align=right style='padding-right:10px;'>"
								
								    if cdbl(medrabat) <> 0 AND cint(visrabatkol) = 1 then
								    strFakdet = strFakdet &"<font class=megetlillesort>" & (medrabat * 100) &" %</td>"
								    else
								    strFakdet = strFakdet &"&nbsp;</td>"
								    end if
								else
								strFakdet = strFakdet &"<td>&nbsp;</td>"
								
								end if
								
								'*** Total beløb pr. linie ***'
								strFakdet = strFakdet &"<td valign='top' align=right><font class=megetlillesort>"& formatnumber(medBelob, 2) &" "& valutaISO &"</td>"_
								&"</tr>"
								
								h = h + 1
								
								oRec2.movenext
								wend
								oRec2.close 
								
                                end if 'jobid

								end if '** media XML



								
					antalAktlinier = antalAktlinier + 1				
					lastaktid = oRec("aktid")			
					oRec.movenext
					wend
					oRec.close 
					

                    if jobid <> 0 then
					
					    '** Gl. fase sum **'
                        if cint(hidefasesum) <> 1 then

                            if h > 0 AND len(trim(fd_fase)) <> 0 AND faseshowtot = 1 then

                                strFakdet = strFakdet &"<tr>"

                                if cint(hideantenh) <> 1 then
                                strFakdet = strFakdet &"<td colspan=2>&nbsp;</td>"
                                end if

				                strFakdet = strFakdet &"<td valign=top colspan=5 style='padding:1px 5px 10px 5px; color:#999999;'>"& replace(lastfdfase, "_", " ") &" i alt:</td>"_
                                &"<td align=right style='padding:1px 0px 10px 5px; color:#999999;'>"&formatnumber(fasebelialt, 2)&" "& valutaISO &"</td></tr>"


                            end if

                        end if

                    end if
					
					
					'if len(medarbejdertimer) <> 0 then
					'medarbejdertimer = medarbejdertimer
					'else
					'medarbejdertimer = 0
					'end if
					
					'if len(medarbejdersum) <> 0 then
					'medarbejdersum = medarbejdersum
					'else
					'medarbejdersum = 0
					'end if
					
					if medarbejdertimer <> intFaktureretTimer then
					bgthis = "#ffffe1"
					else
					bgthis = "#ffffff"
					end if
					
					'if aktsubtotal <> 0 then
					'aktsubtotal = aktsubtotal 
					'else
					'aktsubtotal = 0
					'end if
					
					        
					'intBelobforMoms = aktsubtotal 
					           
					 
					 
					 
					 
					 
					 
					           
					           
					        '********************************************'
					        '*** Materialer *****'
					        '*** m/u moms m = 0 med moms m = 1 u moms ***'
					        '********************************************'
					        
					        antalmatthis = 0
							'for m = 0 to 1 
							matBelob = 0
                            matEhpris = 0
                            
                           
							    '** vis linje / linje eller grupper ***'
								strSQLmat = "SELECT matid, matvarenr,"_
							    &" matrabat, matshowonfak, ikkemoms, "_
							    &" fmat.valuta, fmat.kurs, v.valutakode, matenhed, fmat.matgrp,  " 
							    
							    if cint(showmatasgrp) <> 0 then
							    strSQLmat = strSQLmat & "mgp.navn AS matnavn, "_
							    &" SUM(matbeloeb) AS matbeloeb, SUM(matantal) AS matantal, (sum(matbeloeb)/sum(matantal)) AS matenhedspris "
							    else
							    strSQLmat = strSQLmat & "matnavn, matantal, matbeloeb, matenhedspris "
							    end if
							    
							    strSQLmat = strSQLmat & "FROM fak_mat_spec fmat "_
							    &" LEFT JOIN valutaer v ON (v.id = fmat.valuta) "
							    
							    if cint(showmatasgrp) <> 0 then
							    strSQLmat = strSQLmat & " LEFT JOIN materiale_grp AS mgp ON (mgp.id = fmat.matgrp) "
							    end if
							    
							    strSQLmat = strSQLmat & " WHERE matfakid = "&id&" " 
							    
							    if cint(showmatasgrp) <> 0 then
							    strSQLmat = strSQLmat & " GROUP BY fmat.matgrp, fmat.ikkemoms, fmat.matenhed, fmat.valuta, matrabat "
							    end if
							    
							    '&" AND ikkemoms = "& m
								
								'Response.write strSQLmat
								'Response.Flush
								
								oRec2.open strSQLmat, oConn, 3
                                while not oRec2.EOF
                                        
                                        
                                        fmat_valuta = oRec2("valuta")
						                fmat_kurs = oRec2("kurs")
						                fmat_valutaKode = oRec2("valutakode")
                                        
                                        fmat_Ehpris = oRec2("matbeloeb")
								        fmat_Belob = oRec2("matenhedspris")
                                        
                                        
                                        if fak_kurs <> fmat_kurs then
						                matEhpris = oRec2("matenhedspris") * (fmat_kurs/fak_kurs)
						                else
						                matEhpris = oRec2("matenhedspris") 
						                end if
                                         
                                         '*** Mat beløb er omregnet til den valuta **'
                                         '** der er valgt på faktura inden den lægges i DB **'
                                         if len(trim(oRec2("matbeloeb"))) <> 0 then
                                         matBelob = formatnumber(oRec2("matbeloeb"), 2)
                                         else
                                         matBelob = formatnumber(0, 2)
                                         end if
                                        
                                        'matBelob = oRec2("matbeloeb") * (100/fak_kurs)
                                        'matEhpris = oRec2("matenhedspris") * (100/fak_kurs)
                                       
                                        
                                        if intFaktype = 1 AND media <> "xml" then
                                        fmat_Ehpris = -(formatnumber(fmat_Ehpris, 2))
								        fmat_Belob = -(formatnumber(fmat_Belob, 2))
								        matBelob = -(formatnumber(matBelob, 2))
								        matEhpris = -(formatnumber(matEhpris, 2)) 
								        end if
                                        
                                        '** Materierl uden for gruppe vises som diverse **'
                                        if cint(showmatasgrp) <> 0 AND oRec2("matgrp") = 0 then
                                        matNavn = "Diverse"
                                        else
                                        matNavn = oRec2("matnavn")
                                        end if
                                        
                                        
                                        '*** ÆØÅ **'
                                        enhedsang = oRec2("matenhed")
						                call jquery_repl_spec(enhedsang)
                                        enhedsang = jquerystrTxt
                                        
                                        if len(trim(oRec2("matantal"))) <> 0 then
                                        matAntal = formatnumber(oRec2("matantal"), 2)
                                        else
                                        matAntal = formatnumber(0, 2)
                                        end if
                                        
                                    	strFakmat(m) = strFakmat(m) &"<tr>"_
								        &"<td valign=top width=40 align=right style='padding-right:5px;'>"& matAntal &"</td>"_
								        &"<td valign=top width=40 style='padding-right:5px;'>"& enhedsang &"</td>"_
				                        &"<td colspan=2 valign=top style='padding-left:5px;'>" & matNavn &"</td>"_
								        &"<td valign=top align=right style='padding-right:2px;'>"
					    
					                    
					                    '*** Momsfri ***'
					                    if cint(oRec2("ikkemoms")) = 1 then
					                    mat_momsfriVal = "*"
					                    showmomsfriTxt = 1
					                    'matsubtotal(m) = matsubtotal(m)
					                    else
					                    mat_momsfriVal = ""
					                    'matsubtotal(m) = matsubtotal(m) + (matBelob)
					                    end if
								        
								        if len(trim(oRec2("matenhedspris"))) <> 0 then
								        matEnhPris = formatnumber(oRec2("matenhedspris"), 2)
								        else
								        matEnhPris = formatnumber(0, 2)
								        end if
								        
								         '*** Enh. pris og Valuta på materiale linie ***'
				                        if cint(fak_valuta) <> cint(fmat_valuta) then
				                        strFakmat(m) = strFakmat(m) & matEnhPris &" "& fmat_valutaKode & " ~"
				                        else
				                        strFakmat(m) = strFakmat(m) & "&nbsp;"
				                        end if
                    					
				                        strFakmat(m) = strFakmat(m) &"</td>"
								
								        
								        '*** Enh. pris og Valuta på materiale fra Faktura valuta ***'
								        strFakmat(m) = strFakmat(m) &"<td valign=top align=right style='padding-right:2px;'>"& formatnumber(matEhpris) &" "& valutaISO &""_
								        &"</td>"
								        
								        if cint(visrabatkol) = 1 then
								        strFakmat(m) = strFakmat(m) &"<td valign=top align=right style='padding-right:10;'>"
								       
								            if cdbl(oRec2("matrabat")) <> 0 then
								            strFakmat(m) = strFakmat(m) &"" & (oRec2("matrabat") * 100) &" %</td>"
								            else
								            strFakmat(m) = strFakmat(m)  &"&nbsp;</td>"
								            end if
								        else
								        strFakmat(m) = strFakmat(m)  &"<td>&nbsp;</td>"
								        end if
								        
								        strFakmat(m) = strFakmat(m) &"<td valign='top' align=right>"& mat_momsfriVal &" "& formatnumber(matBelob) &" "& valutaISO &"</td>"_
								        &"</tr>"
                                    
                                        
                                          if media = "xml" then
                                          
                                          session.LCID = 1033
                                          
                                          'call utf_format(oRec2("matenhed"))
                                          'enhedsang = utf_formatTxt
                                          
                                          'call utf_format(matNavn)
                                          'matNavn = utf_formatTxt
                                          
					                      strXML = strXML & "<com:InvoiceLine>"
                                          strXML = strXML & " <com:ID>"&h&"</com:ID>" 
                                          strXML = strXML & " <com:InvoicedQuantity unitCode='"&EncodeUTF8(oRec2("matenhed"))&"' unitCodeListAgencyID=""n/a"">"&replace(matAntal, ",", "")&"</com:InvoicedQuantity>" 
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
                                    
                                    'Response.Write "matsubtotal(m)" & matsubtotal(m) & "<br>"
                                    
                                    h = h + 1
                                    
                                if m = 0 then    
                                antalmatthis = antalmatthis + 1
                                end if
                                oRec2.movenext
                                wend
                                oRec2.close
					        
					        
					        
					    
		          end if 'hidesumaktlinier			
		
		'else '**** Fakturering af Aftaler ***
		
				
				
				
		'		strFakdet = strFakdet &"<tr>"
        '        
        '        if cint(hideantenh) <> 1 then
        '        strFakdet = strFakdet &"<td valign=top width=40 align=right style='padding-right:5px; padding-bottom:20px;'>"& formatnumber(intFaktureretTimer, 2) &"</td>"_
	'			&"<td valign=top width=40 align=right style='padding-right:5px; padding-bottom:20px;'>&nbsp;</td>"
    '            end if
				
      '           if cint(hideantenh) <> 1 then
       '          cspanAdd = 3
       '          else
        '         cspanAdd = 9
        '         end if

         '       strFakdet = strFakdet &"<td colspan='"&cspanAdd&"' valign='top' style='padding-left:5px; padding-bottom:20px;'>" & strJobBesk &"</td>"
			    
         '       if cint(hideantenh) <> 1 then
         '
		'		    if intFaktureretTimer <> 0 AND intFaktureretBelob <> 0 then
		'		    intstkpris = (intFaktureretBelob/intFaktureretTimer)
		'		    else
		'		    intstkpris = 0
		'		    end if
		'		
		'		    strFakdet = strFakdet & "<td valign=top align=right style='padding-right:5px; padding-bottom:20px;'>"& formatnumber(intstkpris, 2) &" "& valutaISO &" </td>"
		'		
		'		    if visrabatkol = 1 then
		'		    strFakdet = strFakdet &"<td valign='top' align=right style='padding-right:5px; padding-bottom:20px;'>"& (100*intAftRabat) &" %</td>"
		'	        end if
		'	    
		'		strFakdet = strFakdet &"<td valign='top' align=right colspan=2 style='padding-bottom:20px;'>"& formatnumber(intFaktureretBelob,2) &" "& valutaISO &"</td>"_
		'		&"</tr>" 
		 '       
          '      else
          '
           '     end if

		'		intBelobforMoms = intFaktureretBelob
				
		'end if
		
		

        '*** Skriver faktura linjer ****'
		'**** Aktiviteter / timer ***'
		if len(trim(strFakdet)) <> 0 then
 		
 		
 		    call udspecificering(strFakdet,1,antalAktlinier)
		
		
		    if jobid <> 0 then
		    '**** Sub-tot. aktiviteter ***'
		        select case lto
		        case "essens"
		        case else
		        call subtotakt()
		        end select
		    end if
		
		
		end if
		    
		    
    		
		    '*** Materialer ***'
		    if len(trim(strFakmat(0))) <> 0 then
		        
		        
    		    '**** Luft før materialer ***'
    		    select case lto
		        case "essens"
		        case else
		        Response.Write "<tr><td colspan=8 height=20>&nbsp;</td></tr>"
		        end select
		    
		   
		    
		        call udspecificering(strFakmat(0),2,antalAktlinier)
    		
    		    
    		    '**** Sub-tot. materialer ***'
    		    select case lto
		        case "essens"
		        case else
		        'call subtotmat(0)
		        end select
		    
		    end if
		    
		    
		    '*** Rykker gebyrer ***'
		    call rykkergebyrer()
		    
		    '**** Totaler beløb til moms ***'
		    call totogmoms()
		    
		
		    '*** Materialer u/moms ***'
		    if len(trim(strFakmat(1))) <> 0 then
		    'Response.Write "<tr><td colspan=8 height=20>&nbsp;</td></tr>"
		    call udspecificering(strFakmat(1),3,antalAktlinier)
    		'**** Sub-tot. materialer ***'
    		    
    		    select case lto
		        case "essens"
		        case else
		        call subtotmat(1)
		        end select
		    
		    end if
		
		    
		
		
		
		
		
		'**** Totaler ***'
		call totalbelob()
		
		
		'*** Komm. og bet. betingelser ***'
		call komogbetbetingelser
		
		
		Response.Write "</div>" 'fakturaside
		
		
		
		stDatoKri = request("FM_start_aar_ival") &"/"& request("FM_start_mrd_ival") &"/"& request("FM_start_dag_ival")  
		slutdato = request("FM_slut_aar_ival") &"/"& request("FM_slut_mrd_ival") &"/"& request("FM_slut_dag_ival")  
		
		
		'*** Joblog ****
		if visjoblog = 1 then%>
		
		<!-- joblog -->
		
	    <div id="joblog" style="position:relative; width:640px; left:10px; top:50px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; padding:10px 10px 10px 10px;">
		<table width=100% border=0 cellspacing=0 cellpadding=0 style="page-break-before:always;">
		<tr>
			<td colspan=6><h4><%=txt_038 %></h4></td>
		</tr>
		
		
		<%
		call joblog(jobid, stdatoKri, slutdato, aftid)
		%>
		
		</table>
		</div>
		
		<br><br>&nbsp;
		<!-- Joblog slut -->
		<%end if 	
		
		
		
		
		'*** Matlog ****
		if vismatlog = 1 then%>
		
		<!-- matlog -->
		
		 <div id="matlog" style="position:relative; width:640px; left:10px; top:50px; visibility:visible; background-color:#ffffff; border:<%=bd%>px #8caae6 solid; padding:10px 10px 10px 10px;">
		
		<table border=0 cellspacing=0 width=620 cellpadding=0 style="page-break-before:always;">
		<tr>
			<td><h4><%=txt_039 %></h4></td>
		</tr>
		</table>
		
		<table width=100% cellspacing=0 cellpadding=0 border=0>
		<tr>
	                            <td><b><%=txt_042 %></b></td>
	                            <td><b><%=txt_043 %></b></td>
	                            <td><b><%=txt_026 %></b></td>
	                            <td align=right><b><%=txt_013 %></b></td>
	                            <td align=right><b><%=txt_020 %></b></td>
	                            <td align=right><b><%=txt_044 %></b></td>
	                            <td align=right><b><%=txt_017 %></b></td>
	                        </tr>
		
		<%
		call matlog(jobid, stdatoKri, slutdato, aftid)
		%>
		
		</table>
		</div>
		<br><br>&nbsp;
		
		
		<!-- matlog slut -->
		<%end if %>	
		
		</div>
		<br /><br />
        &nbsp;
											
		<%
		'**** Højreside (ikke printbar) ***
		
		
		
		if media = "n" then %>
		<div id=funktioner style="position:absolute; z-index:200; top:30px; width:300px; left:520px; border:2px #8caae6 DASHED; background-color:pink; padding:5px 5px 5px 5px;">
		<b>Udskrift funktioner:</b> <br />
		<a href="erp_fak_godkendt_2007.asp?media=j&jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=id%>&kid=<%=kid%>&FM_job=<%=Request("FM_job")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag_ival=<%=request("FM_start_dag_ival")%>&FM_start_mrd_ival=<%=request("FM_start_mrd_ival")%>&FM_start_aar_ival=<%=request("FM_start_aar_ival")%>&FM_slut_dag_ival=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd_ival=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar_ival=<%=request("FM_slut_aar_ival")%>&FM_fakint_ival=<%=request("FM_fakint_ival")%>" target="_blank" class=vmenu>Print</a>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<!--<a href="erp_make_pdf.asp?lto=<%=lto%>&faknr=<%=varFaknr%>&jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=id%>&kid=<%=kid%>&FM_job=<%=Request("FM_job")%>&FM_usedatointerval=<%=request("FM_usedatointerval")%>&FM_start_dag_ival=<%=request("FM_start_dag_ival")%>&FM_start_mrd_ival=<%=request("FM_start_mrd_ival")%>&FM_start_aar_ival=<%=request("FM_start_aar_ival")%>&FM_slut_dag_ival=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd_ival=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar_ival=<%=request("FM_slut_aar_ival")%>&FM_fakint_ival=<%=request("FM_fakint_ival")%>" target="_blank" class=vmenu>PDF</a><br /><br />-->
		<a href="erp_make_pdf.asp?media=pdf&lto=<%=lto%>&faknr=<%=varFaknr%>&jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=id%>&FM_start_dag_ival=<%=request("FM_start_dag_ival")%>&FM_start_mrd_ival=<%=request("FM_start_mrd_ival")%>&FM_start_aar_ival=<%=request("FM_start_aar_ival")%>&FM_slut_dag_ival=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd_ival=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar_ival=<%=request("FM_slut_aar_ival")%>" target="_blank" class=vmenu>PDF</a>
		
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<a href="#" class=vmenu id="oioxml">OIO XML</a> 
        
        <%select case lto
	    case "epi", "outz", "intranet - local"
        showvans = 1
        %>
        (Afsend XML fil til E-posthus)
	    <%
        case else
        showvans = 0
	    end select
	    %>

		 
        <form action="#" method="post">
        <input id="xmlurl" type="hidden" value="erp_fak_godkendt_2007.asp?media=xml&lto=<%=lto%>&faknr=<%=varFaknr%>&jobid=<%=jobid%>&aftid=<%=aftid%>&id=<%=id%>&FM_start_dag_ival=<%=request("FM_start_dag_ival")%>&FM_start_mrd_ival=<%=request("FM_start_mrd_ival")%>&FM_start_aar_ival=<%=request("FM_start_aar_ival")%>&FM_slut_dag_ival=<%=request("FM_slut_dag_ival")%>&FM_slut_mrd_ival=<%=request("FM_slut_mrd_ival")%>&FM_slut_aar_ival=<%=request("FM_slut_aar_ival")%>" />
        <input id="showvans" type="hidden" value="<%=showvans %>" />
        </form>

       
	    <br /><br />
	    <%select case lto
	    case "execon", "immenso"', "intranet - local"
	    
	    case else
	    %>
	   
	    <b>Gem PDF ovenfor, og email faktura til:</b><br />
		<%
		
		'pdfurl = "d:\\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\faktura_"&lto&"_"&faknr&".pdf"
	    
		
		strSQL = "SELECT navn, email, titel FROM kontaktpers WHERE kundeid = " & kid
		'Response.Write strSQL 
		'Response.Flush
		
		oRec.open strSQL, oConn, 3
        while not oRec.EOF
        
        Response.Write "<i>"& oRec("navn") & "</i> "
        
        if len(trim(oRec("titel"))) <> 0 then
        Response.Write " ("& oRec("titel") &") "
        end if
        
        
        Response.Write " <a href='mailto:"&oRec("email")&"&subject=Faktura: "& varFaknr &"' class=vmenu>" & oRec("email") & "</a><br>"
        
        oRec.movenext
        wend
        oRec.close
		%>
		<br />Kontaktpersoner oprettes under fanebladet "kontakter" i hovemenuen.
		
		<%end select %>	
		
		<br /><br />
		<table cellspacing=2 cellpadding=1 border=0 width=290 bgcolor="#ffffff">
		
		<tr>
		<%
		 if intStatus = 0 then
	    %>
	    <td width=120 valign=top style="padding-top:6px;">
	    <a href="Javascript:popUp('godkendfaktura.asp?fakid=<%=id %>','400','250','50','50');" target="_self"  class=vmenu>
        Godkend faktura nu</a>
	   </td>
	    <td valign=top width=170>
	     <a href="Javascript:popUp('godkendfaktura.asp?fakid=<%=id %>','400','250','50','50');" target="_self"  class=vmenu>
              <img src="../ill/godkend_pil.gif" border="0" />  </a>  </td>
          </tr>
          <tr><td colspan=2 valign=top class=lille>
        
        <%select case lto
	    case "execon", "immenso", "intranet - local" 
	    Response.Write "Afsender mail med PDF af faktura til bogholderiet."
	    case else
	    Response.Write "&nbsp;"
	    end select%>
              
              
	    </td>
	    
	    <% 
	    else
	    %>
	    
	
	   <td><img src="../ill/godkend_icon.gif" border="0" /></td>
	   <td><font color=forestgreen><b>Godkendt</b></font></td>
	      </tr>
		
	     
	    <%
	    end if
		%>
		</tr>
		</table>
		
		</div>
		<!--&nbsp;&nbsp;|&nbsp;&nbsp;Excel.-->
	    <%end if
	    
	    
	    
	    
	   
	    
	    
	    
	    if media = "j" then
	    Response.Write("<script language=""JavaScript"">window.print();</script>")
	    end if
	    
		if media = "xml" then
		
		
    	session.LCID = 1033  
    	strXML = strXML & "</Invoice>"
	    
	    ext = "xml"
		
         call TimeOutVersion()
	
	    Set objFSO = server.createobject("Scripting.FileSystemObject")
    	
    	file = "faktura_xml_"&lto&"_"&varFaknr&"."& ext
    	
	    if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\"&toVer&"\timereg\erp_fak_godkendt_2007.asp" then
    							
		    Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"& file &"", True, False)
		    Set objNewFile = nothing
		    Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"& file &"", 8,-1)
    	
	    else
    		
		    Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"& file &"", True, false)
		    Set objNewFile = nothing
		    Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"& file &"", 8,-1)


            if cint(vans) = 1 then
                '***  VANS CSC ***
                select case lto
	            case "outz", "epi" 
                    Set objNewFileVans = objFSO.createTextFile("d:\webserver\CSC\CSC VANS\TIL-VANS\"& file &"", True, false)
		            Set objNewFileVans = nothing
		            Set objFVans = objFSO.OpenTextFile("d:\webserver\CSC\CSC VANS\TIL-VANS\"& file &"", 8,-1)


                    objFVans.WriteLine(strXML)
	                objFVans.close

                end select
            end if

    		
	    end if
    	
    	
		
        objF.WriteLine(strXML)
	    objF.close

	    Response.redirect "../inc/upload/"&lto&"/"& file &""
	    
	    end if
		
		
		        		        
		        
		
		
											
											
											
											
											
											
											
	
	
	
	
end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->


