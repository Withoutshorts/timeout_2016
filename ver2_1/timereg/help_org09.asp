<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<html>
		<head>
		<title>TimeOut</title>
		 	<LINK rel="stylesheet" type="text/css" href="../inc/style/TimeOut_style_print_fak.css">
		</head>
	<body topmargin=0 leftmargin=0>
	
	<%
	
	
	function subindeks_header(menu, subheader)
	%>
	
	<tr>
		<td height="50" colspan=2>&nbsp;&nbsp;</td>
	</tr>
	<tr bgcolor="#fffff1">
		<td valign="top" colspan="2" style="padding:20px 20px 20px 20px; border:2px orange solid;"><h3><%=subheader %></h3>
		<b>Emner:</b><br>
	
	<%
	end function
	
	function subindeks_footer()
	%>
		
		<br><br>Hvis det emne du s�ger ikke findes her, s� send os en email p� <a href="mailto:support@outzource.dk" class=vmenu>support@outzource.dk</a></td>
	</tr>
	
	<%
	end function
	
	
	%>
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:0; top:0; width:100%; height:600; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="350">
		<tr>
			<td bgcolor="#003399"><img src="../ill/logo_topbar_print.gif" alt="" border="0"></td>
		</tr>
		
	</table>
	<br>
	<table cellspacing="0" cellpadding="0" border="0" width="90%" style="padding-left:20;">
	<tr>
    <td valign="top"><h3>TimeOut Hj�lp</h3>
	Brug [ctrl] [F] for at s�ge efter emne.<br>
	<br>
	<h4>Indeks:</h4>&nbsp;<br>
	<table cellpadding=2 cellspacing=1 border=0 width=100%>
	<tr>
	    <td width=175 valign=top style="border-right:1px #c4c4c4 solid;"><b>Generelt</b><br />
	    <a href="help.asp?menu=komgodtigang" class=vmenu><u>1.0 Kom godt igang - Guide</u></a>
	    <br><a href="help.asp?menu=gen" class=vmenu><u>1.1 TimeOut opbygning / flowchart</u></a>
	    </td>
		<td width=175 valign=top style="border-right:1px #c4c4c4 solid;"><b>TSA (time/sag):</b>
		<br><a href="help.asp?menu=treg" class=vmenu><u>2.0 Timeregistrering</u></a>
		<br><a href="help.asp?menu=job" class=vmenu><u>2.1 Job</u></a>
		<br><a href="help.asp?menu=kund" class=vmenu>2.2 Kontakter</a>
		<br><a href="help.asp?menu=medarb" class=vmenu><u>2.3 Medarbejdere</u></a>
		<br><a href="help.asp?menu=stat" class=vmenu>2.4 Statistik</a>
		<br /><a href="help.asp?menu=mat" class=vmenu><u>2.5 Materialer / udl�g</u></a>
		</td>

		<td width=175 valign=top style="border-right:1px #c4c4c4 solid;"><b>CRM:</b>
		<br /><a href="help.asp?menu=crmkal" class=vmenu>3.0 Kalender</a>
		<br><a href="help.asp?menu=firma" class=vmenu>3.1 Firmaer</a>
		<br><a href="help.asp?menu=crmakthist" class=vmenu>3.2 Aktions historik</a>
		</td>
	
	
		<td width=175 valign=top style="border-right:1px #c4c4c4 solid;"><b>SDSK: (ServiceDesk)</b><br />
		<a href="help.asp?menu=sdsk" class=vmenu><u>4.0 Incidents</u></a></td>
		
		<td width=175 valign=top style="border-right:1px #c4c4c4 solid;"><b>ERP (Bogf�ring og fakturering):</b><br />
		<a href="help.asp?menu=stat_fak" class=vmenu><u>5.0 Fakturering</u></a><br>
		&nbsp;</td>
	</tr>

	</table>
	
	<table width="100%" cellpadding="0" cellspacing="0" border="0">
	<%topmenuSel = request("menu")
	Select case topmenuSel
	case "sdsk"
	%>
	<tr>
		<td height="20">&nbsp;&nbsp;</td>
	</tr>
	<tr bgcolor="#fffff1">
		<td valign="top" colspan="2" style="padding-left:10px; padding-top:5px; border:1px darkred solid;"><h3>ServiceDesk</h3>
		<b>Emner:</b><br>
		<a href="#sdsk_inci" class="kalblue">Opret Incidents</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<br><br>Hvis det emne du s�ger ikke findes her, s� send os en email p� <a href="mailto:support@outzource.dk">support@outzource.dk</a>.</td>
	</tr>
	<tr>
		<td valign="top" colspan="2"><br><b>Opret Incidents</b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"  class="alt" style="padding-left:4px;"><br>ServiceDesk</td>
		<td width="500" bgcolor="#ffffff" style="padding-left:4px;"><br><br>
		N�r en Incident oprettes som en aktivitet p� et job, bliver den oprettet med f�lgende data:<br>
	- Navn s�ttes lig med navnet p� Incidenten.<br>
	- Beskrivelse s�ttes lig med beskrivelsen p� Incicenten.<br>
	- Datointerval nedarves fra jobbets datointerval.<br>
	- Faktor s�ttes lig med den faktor der er angivet under prioitet p� Incidenten.<br>
	- Projektgrupper s�ttes lig med de grupper der er tilknyttet jobbet.<br>
	- V�rdi s�ttes til 0.<br>
	- Status s�ttes altid til "Aktiv."<br>
	- Timepriser nedarves fra jobbet.<br>
	- Forkalk. timer s�ttes = Estimeret timeforbrug p� Incident.<br><br>&nbsp;
		<br>&nbsp;</td>
	</tr>
	<%
	case "crmakthist"
	%>
	<tr>
		<td valign="top" colspan="2"><br><b>Aktions historik</b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"  class="alt" style="padding-left:4px;"><br>Aktions historik</td>
		<td width="500" bgcolor="#ffffff" style="padding-left:4px;"><br><b>Aktions oversigt:</b><br>I aktions historikken er det muligt at se hvilke aktioner der er oprettet for hver enkelt firma og hvilke mederbejdere der har v�ret i kontakt med firmaet.<br><br>
		�verst til h�jre kan der s�ges p� <b>kontakt personer, firmanavne, telefon </b>eller <b>email</b>.<br>I menuen til venstre kan der sorteres mellem de kriterier som aktions historik listen vises efter.<br><br>
		<b>Ny aktion:</b><br>
		N�r der oprettes en ny aktion skal der v�lges: emne, status og kontaktform. I TimeOut er der default oprettet nogle status'er, emner og kontaktforme, men der kan altid oprettes flere efter eget valg.
		Det er valgfrit om man �nsker at angive en titel p� aktionen.<br><br>
		Hvis der ikke angives et tidspunkt bliver tidspunktet automatisk sat til kl. 12.00. <br><br>
		I "Beskrivelse / log" indtastningsfeltet bliver der automatisk indsat dags dato, s� man hurtigt kan indtaste en note om aktionen.<br><br>
		Nederst p� "Opret ny aktion" siden skal man v�lge hvilke medarbejdere der skal deltage i aktionen.
		<br>&nbsp;</td>
	</tr>
	<%
	case "crmkal"
	%>
	<tr>
		<td valign="top" colspan="2"><br><b>Kalender</b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"  class="alt" style="padding-left:4px;"><br>Kalender</td>
		<td width="500" bgcolor="#ffffff" style="padding-left:4px;"><br>I kalenderen er det mulig at se m�der og aftaler.<br>
		 Kalenderen fungerer som f�lles kalender, s�ledes at man altid kan se om en medarbejder er ledig.<br><br>
		 Der kan sorteres mellem medarbejdere s� man kan benytte kalenderen som personlig kalender, eller hurtig f� overblik over en bestemt medarbejders kalender.<br><br>
		 Der kan <b>bladres i kalenderen</b> ved at klikke p� datoerne i m�nedsoversigten til h�jre.<br><br>
		 Det person der bliver vist med <b>fed</b> skrift er aktionens "ejer". De andre er de tilknyttede personer.<br>&nbsp;</td>
	</tr>
	<%
	case "kund", "firma"
	if topmenuSel = "firma" then
	oskrift = "Firmaer"
	bt = "firma"
	else
	oskrift = "Kunder"
	bt = "kunde"
	end if
	%>
	<tr>
		<td height="20">&nbsp;&nbsp;</td>
	</tr>
	<tr bgcolor="#fffff1">
		<td valign="top" colspan="2" style="padding-left:10px; padding-top:5px; border:1px darkred solid;"><h3>Kontakter</h3>
		<b>Emner:</b><br>
		<a href="#aftaler" class="kalblue">Aftaler</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#filer" class="kalblue">Filarkiv</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		
		<br><br>
		Hvis det emne du s�ger ikke findes her, s� send os en email p� <a href="mailto:support@outzource.dk">support@outzource.dk</a>.</td>
	</tr>
	<!-- Aftaler -->
	<tr>
		<td valign="top" colspan="2"><br><b><a name="aftaler">Aftaler</a></b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top" align="center"  class="alt">&nbsp;</td>
		<td ><br>
		N�r en aftale lukkes kan den ikke l�ngere v�lges som aftale n�r et job oprettes/redigeres.<br><br>
		Hvis et timeantal �ndres p� en allerede indtastet timeregistrering, bliver antallet af brugte enheder p� et klippekort ogs� �ndret selom aftalen
		er lukket. 
		
		N�r en aftale er lukket kna den ikke l�ngere faktureres.
		<br><br>
		Rev. dato: 15/7-2005 
		</td>
	</tr>
	
	<!-- filarkiv -->
	<tr>
		<td valign="top" colspan="2"><br><b><a name="filer">Filarkiv</a></b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top" align="center"  class="alt">&nbsp;</td>
		<td ><br>
		Kun administratorer, niveau 1 og 1B brugere kan oprette foldere og filer. Alle brugere kan se foldere.<br>
		Brugere kan <u>se, redigere og slette</u> filer, hvortil "Alle" har adgang til. <br>
		Administratorer niveau 1 og 1B brugere kan se, redigere og slette alle filer. <br>
		Kontakter kan kun se filer og foldere hvortil der er givet adgang.
		<br><br>
		
		N�r en fil uploades tjekkes det om filen alerede findes i databasen. <br>
		Hvis den findes bliver den ikke oprettet i databasen igen.<br><br>
		
		Filer overskrives altid fysisk n�r en ny fil uploades.<br><br><br>
		
		Rev. dato: 15/8-2005 
		</td>
	</tr>
	
	<%
	case "job"
	call subindeks_header(menu, "Job")
	%>
	<a href="#j_fastpris" class="kalblue">Fastpris / l�bende timer</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		
	
	<%call subindeks_footer() %>	
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><a name="j_fastpris"><b>Fastpris / l�bende timer.</b></a>
	    <br>
		N�r et job oprettes kan der, v�lges og det skal faktureres som l�bende timer, eller som et <b>fastpris</b> job. 
		Hvis der v�lges fastpris bliver timeprisen beregnet udfra antal forkalkuleret timer
		i forhold til den pris der angives p� jobbet.  <br><br>
		<b>Alternative timepriser<br></b>
		Der kan p� de enkelte job angives alternative timepriser for hver enkelt medarbejder der har adgang til jobbet.<br>
		<br>Disse timepriser angives under alternative timepriser som findes under medarbejdere --> medarbejdertyper. Der kan ogs� angives en helt specifik timepris for den enkelte medarbejder, kun g�ldende for et enkelt job. 
		Dette g�res ved at v�lge "timepris 6" og angive en timepris.<br><br>
		Timepriser angivet p� jobbet nedarves automatisk af hver enkelt aktivitet p� jobbet, med mindre der angives eller v�lges en anden timepris p� den enekelte aktivitet.<br><br>
		Timer der allerede er indtastet i systemet �ndrer ikke timepris selvom medarbejderen �ndrer timepris undervejs i jobforl�bet. Denne medarbejder optr�der herefter med to mulige timerpiser n�r der oprettes en faktura. (<a href="help.asp?menu=stat_fak" class=kalblue>Se fakturering</a>)
		
		
		</td>
	</tr>
	
	
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b>Projektgrupper.</b><br />
	    En projektgruppe er en gruppe af brugere. 
		Projektgrupper bruges til at styre hvilke medarbejdere, der skal have adgang til de forskellige job og aktiviteter.<br>
		</td>
	</tr>
	
	
	
		<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b>Stamaktiviteter</b>
		<br />
		For at g�re det s� let og hurtigt at oprette et nyt job og tilknytte aktiviteter, kan der n�r jobbet oprettes tilknyttes 1-5 Stamaktivitetsgrupper.<br>
		Stam-aktiviteter oprettes, under menupunktet "Stam-aktiviteter og Grp."<br><br>
		
		L�s mere om job <a href="help.asp?menu=komgodtigang#g-oprjob" class=kalblue>her..</a>
		</td>
	</tr>
	
	
	
	
	
	
	<%
	case "medarb"
	%>
	<tr>
		<td height="20">&nbsp;&nbsp;</td>
	</tr>
	<tr bgcolor="#fffff1">
		<td valign="top" colspan="2" style="padding:20px; border:2px orange solid;"><h3>Medarbejdere</h3>
		<b>Emner:</b><br>
		<a href="#brugergp" class="kalblue">Brugergrupper</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#medarbtype" class="kalblue">Medarbejdertyper</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#medlem" class="kalblue">Medlemmer</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#oprred" class="kalblue">Opret / Rediger</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#deak" class="kalblue">Deaktivering af medarbejder</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#timkost" class="kalblue">Generel timepris, alternativ timepris og kostpris.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#timprstam" class="kalblue">Timepriser p� stam-aktiviteter</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br><br>
		Hvis det emne du s�ger ikke findes her, s� send os en email p� <a href="mailto:support@outzource.dk">support@outzource.dk</a>.</td>
	</tr>
	<tr>
		<td valign="top" colspan="2"><br><b><a name="oprred">Opret / Rediger.</a></b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top" align="center"  class="alt">&nbsp;</td>
		<td ><br>Klik p� opret medarbejder.<br>
		<font color=red>
		<ul>
		<li>Hvis denne funktion ikke er tilg�ngelig er det fordi der allerede er oprettet det antal medarbejdere der er k�bt licens til.<br>
		<li>Kun administratorer kan oprette nye medarbejdere.
		</font>
		</ul>
		<br>
		F�lgende information er n�dvendig (*) n�r der <b>oprettes</b> en medarbejder:
		<ul>
		<li>Navn. (*)
		<li>Medarbejder nr. (*)
		<li>Ansat (bruges pt. ikke)
		<li>Login (*) (det brugernavn medarbejderen skal logge p� med.)
		<li>Password (*) (det password medarbejderen skal logge p� med. 0-9 og a-z. Password m� meget gerne indeholde b�de tal og sm� og store bogstaver.)
		<li>Email (*)
		<li><a href="#medarbtype" class="kalblue">Medarbejdertype.</a> (*)
		<li><a href="#brugergp" class="kalblue">Brugergruppe</a>. (*)
		<li>Startside. Som standard �bner TimeOut 2.1 op med time-registrerings siden. Hvis man derimod �nsker at starte Timeout 2.1 op i CRM kalenderen klikkes CRM til.</td>
		</ul>
	</tr>
	<tr>
		<td valign="top" colspan="2"><br><b><a name="medarbtype">Medarbejdertyper.</a></b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"  class="alt" style="padding-left:4px;">&nbsp;</td>
		<td ><br><ul>En <b>medarbejdertype</b> d�kker over den type medarbejder som medarbejderen er oprettet som.<br>
		Det kan f.eks v�re direkt�r, udvikler, konsulent mv. Under hver <b>medarbejdertype</b> 
		angives typens normerede arbejds-uge, som timer pr. dag, samt en <a href="#timkost" class="kalblue">generel timepris, samt op til 5 alternative timepriser + en kostpris</a>. 
		Der kan oprettes s� mange medarbejdertyper der er behov for.<br>&nbsp;&nbsp;
		</td>
	</tr>
	<tr>
		<td valign="top" colspan="2"><br><b><a name="brugergp">Brugergrupper.</a></b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"  class="alt" style="padding-left:4px;">&nbsp;</td>
		<td ><br><ul><b>Brugergrupper</b> er en gruppe af brugere med adgang til TimeOut 2.1. Der er som standard 7 forskellige <b>brugergrupper</b>.<br><br>
		<li><b>Administrator.</b><br>
		Administrator gruppen har adgang til alle omr�der og alle funktioner, incl. [TimeOut kontrolpanel]. (Rettighedsniveau 1) 
		<li><b>TSA + CRM niveau 1.</b> (Rettighedsniveau 2) <br>
		Denne gruppe har adgang til alle omr�der (som ovenst�ende og alle funktioner i Timeout p� lederniveau. <b><font color=darkred>Dog ikke adgang til</font></b> TSA/CRM [TimeOut kontrolpanel].
		<br><li><b>TSA + CRM niveau 1B.</b> (Rettighedsniveau 3) <br>
		Denne gruppe har adgang til alle omr�der (som ovenst�ende) og alle funktioner i Timeout p� lederniveau. <b><font color=darkred>Dog ikke adgang til</font></b> TSA/CRM [TimeOut kontrolpanel], TSA [Statistik], TSA [Fakturering], og kun begr�nset adgang til TSA [Medarbejdere].
		<br><li><b>TSA + CRM niveau 2.</b>  (Rettighedsniveau 4)<br>
		Har adgang til TSA og CRM delen (TSA [timeregistrering] og TSA [medarbejdere], samt CRM [Kalender], CRM [Aktions historik] og CRM [Firma kontakter]).<br> Adgang <b>kun</b> p� bruger niveau, det vil sige indtastning af timer, samt oversigter der vedr�rer medarbejderen selv.
		<br><li><b>TSA niveau 1.</b> (Rettighedsniveau 6)<br>
		Har kun adgang til TSA delen (som gruppe TSA + CRM niveau 1).
		<br><li><b>TSA niveau 2</b>  (Rettighedsniveau 7)<br>
		Har kun adgang til TSA delen (som gruppe TSA + CRM niveau 2).
		<br><li><b>G�st.</b> (Rettighedsniveau 8)<br>
		Har adgang til TSA [timeregistrering].</ul>
		<ul>Der kan ikke oprettes eller slettes <b>brugergrupper</b>.</ul></td>
	</tr>
	
	
	
	<tr>
		<td valign="top" colspan="2"><br><b><a name="deak">Deaktivering af medarbejder</a></b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"  class="alt" style="padding-left:4px;">&nbsp;</td>
		<td ><br>
		<b>Deaktivering af medarbejder</b><br>
		N�r en medarbejder deaktiveres kan denne ikke l�ngere logge p�. Medarbejderen vil kunne
		sorteres fra i statistikker, og vil ikke l�ngere fremg� af faktura-oprettelser.
		Medarbejderen vil heller ikke l�ngere indg� som en del af jeres licens TSA beregning, 
		men i vil <b>ikke</b> kunne oprette ekstra medarbejdere, i forhold til jeres brugerlicens fordi en medarbejder er deaktiveret. 
		<br><br>
		<font class=megetlillesort>Rev. dato: 27/09-2005</font>
		</td>
	</tr>
	
	<tr>
		<td valign="top" colspan="2"><br><b><a name="timkost">Generel timepris, alternativ timepris og kostpris</a></b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"  class="alt" style="padding-left:4px;">&nbsp;</td>
		<td ><br><ul>
		<b>Generel timepris</b><br> 
		Generel timepris angives under medarbejder-typer. Den generelle timepris kan kun tildeles p� job, og kan nedarves af aktiviteter.
		<br />Generel timepris kan ikke benyttes af stamaktiviteter.
        <br /><br />
        <b>Alternative timepriser</b><br />
        Kan angives, s� en medarbejder kan antage forskellige timepriser p� job og aktiviteter.
        De alternative timepriser kan ses b�de p� job og p� aktivitets niveau. (og p� stam-aktiviteter) 
        <br /><br />
        De alternative timepriser er kun vejledende. Det vil stadigv�k v�re muligt at angive en specifik timepris for det enkelte job eller aktivitet, uden at den iforvejen er angivet her. 
        Hvis ikke man altid �nsker at bruge den generelle timepris, er det en god ide at s�tte
        alt. timepris 1 = den generelle timepris. 
        <br /><br />

		<b>Intern kostpris</b><br>
		Alle timer uanset om jobbet er sat til fastpris eller l�bende timer blive omsat med den interne kostpris.
		<br><br>
		<b>Opdatering af timepriser</b></br>
		N�r et job opdateres, opdateres alle timepriser. Det sl�r samtidig igennem p� alle eksisterende timeregistreringer p� aktiviteter der nedarver timepris fra job.
		
		N�r en timepris �ndres p� en timeregistrering �ndres alle hidtidige timepriser, p� eksisterende timeregistreringer, p� den p�g�ldende aktivitet til denne timepris.<br><br>
		</ul>
		
		<font class=megetlillesort>Rev. dato: 13/02-2008</font>
		</td>
	</tr>
	<tr>
		<td valign="top" colspan="2"><br><b><a name="#timprstam">Timepriser p� stam-aktiviteter.</a></b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"  class="alt" style="padding-left:4px;">&nbsp;</td>
		<td ><br><ul>
		N�r en medarbejder oprettes, opretter systemet automatisk en timepris for medarbejderen p� alle 
	  	eksisterende stamaktiviteter.
		Timerpriserne er baseret p� de g�ldende timerpriser p� hver enkelt stamaktivitet, for andre medarbejdere af samme medarbejdertype.
		</ul>
		<br><br>
		<font class=megetlillesort>Rev. dato: 9/5-2005</font>
		</td>
	</tr>
	<tr>
		<td valign="top" colspan="2"><br><b><a name="medlem">Medlemmer.</a></b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"  class="alt" style="padding-left:4px;">&nbsp;</td>
		<td ><br><ul>I oversigten <b>"medlemmer"</b> kan man se alle de medarbejdere, der er oprettet under den valgte medarbejdertype. Man flytter et medarbejder mellem de forskellige <a href="#medarbtype" class="kalblue">Medarbejdertyper</a> ved at redigere en medarbejder.</ul>
		</td>
	</tr>
	<!--</table>-->
	<%
	case "stat"
	%>
	<tr>
		<td height="20">&nbsp;&nbsp;</td>
	</tr>
	<tr bgcolor="#fffff1">
		<td valign="top" colspan="2" style="padding-left:10px; padding-top:5px; border:1px darkred solid;"><h3>Statistik</h3>
		<b>Emner:</b><br>
		Under opdatering.<br><br>
		Hvis det emne du s�ger ikke findes her, s� send os en email p� <a href="mailto:support@outzource.dk">support@outzource.dk</a>.<br>
		<!--Kun timer indtastet efter d. 19/9 2003, der kan have en kostpris. Kostprisen s�ttes p� medarbejder-typen.--></td>
	</tr>
	<%
	case "mat"
	
	call subindeks_header(menu, "Materialer / udl�g")
	%>
	
		<a href="#m-gen" class="kalblue">Generelt</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		
	
	<%call subindeks_footer() %>	
	
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><a name="m-gen"><b>Generelt om materialer og udl�g</b></a><br />
	
	<br />
	<b>Opret materialer.</b><br>
	Under menupunktet <b>"Material. og Produktion."</b> findes f�lgende undermenu punkter:
	<ul> 
	<li>Materialer
	<li>Materialegrupper
	<li>Standard produktioner
	</ul>
	
	F�rst oprettes grupper, dern�st materialer og tilsidst kan der laves nogle standard produktioner s�ledes at man kan se hvor lang tid det tager at producere en bestemt ting, samt se hvilke materialer der indg�r i produktionen.
	<br><br> 
	<b>Materiale registrering p� et job.</b><br>
	Ved timeregistreringen er der kommet et nyt materiale-ikon til h�jre udfor hvert produkt. 
	Her kan der tilf�jes materialer p� et job.
	<br><br> 
	<b>Stat:</b><br>
	Under Statistik er der kommet et link til materiale forbrug. Her kan man f�lge med i hvilke materialer der er brugt p� hver enkelt job.
	<br><br> 
	<b>Fakturering.</b><br>
	Det er  muligt at fakturere sit materiale forbrug i ERP delen.
		<br><br>
		&nbsp;
		
		</td>
	</tr>
	
	<%
	case "stat_fak"
	call subindeks_header(menu, "Fakturering")
	%>
	
		<a href="#f-eksport" class="kalblue">Eksport af fakturaliste.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-andre" class="kalblue">"Andre?" sum-aktiviteten.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-interval" class="kalblue">Datointerval.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-dec" class="kalblue">Decimaler.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-fak" class="kalblue">Faktura oprettelse</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-gamle" class="kalblue">Gamle fakturaer.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-ikke" class="kalblue">Ikke fakturerbare aktiviteter.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-komma" class="kalblue">Komma / punktum.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-rest" class="kalblue">Restbel�b til fakturering.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-ryk" class="kalblue">Rykkere og Kreditnotaer.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-skjulte" class="kalblue">Skjulte timer.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-tp" class="kalblue">Timepriser.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-valg" class="kalblue">Valg af kunder og job.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#f-vente" class="kalblue">Vente timer.</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		
	
	<%call subindeks_footer() %>	
	
	<tr>
		<td width="25" valign="top"  class="alt" style="padding-left:4px;">&nbsp;</td>
		<td valign="top" ><br><br>
		
		<b><a name="f-eksport">Eksport af fakturaliste.</a></b><hr>
		<u>Timer og bel�b, fordeles s�ledes i den ; opdelte txt. fil:</u><br>
		- En linie for hver aktivitet p� de valgte fakturaer. <br>(Tilh�rende medarbejder linier er markeret "1" i kolonnen "Medarb.linier findes?")<br><br>
		- En linie for hver medarbejder, uanset om medarbjederen er vist p� faktura.<br> (1 eller 0 i kolonnen "Vist p� Faktura?")<br><br>
		- En linie der er markeret med "A" er en aktivitets linie og "M" er en medarbejder linie.<br><br>
		- Det er ikke muligt for mig at skildre mellem "andre?" aktiviteterne og de almindelige aktiteter p� anden m�de end at tjekke om der er tilh�rende medarbejder linier. 
		Der vil altid v�re medarbejder linier p� en alm. aktivitet og der vil aldrig v�re linier p� en "Andre?" aktivitet.<br><br>
		- Medarbejdere hvor der er 0 timer registreret vil ikke blive vist. 

		<br><br><br>
		
		<b><a name="f-fak">Faktura oprettelse.</a></b>
		<hr>
		Indtastning og �ndring af timer og timerpiser kan kun ske p� medarbejdere. Herefter bliver timerne ved et klik p� "Calc" knappen overf�rt til den sum-aktivitet der passer til den valgte timepris. Hvis timeprisen ikke findes oprettes der en ny sum-aktivitet.<br><br>
		Der kan registreres b�de fak-timer og vente-timer. Fak-timer er de timer der bliver faktureret p� den enkelte medarbejder. <br><br>
		Hvis "Vis p� fak" markeres, vises medarbejder fak-timer p� faktura printet.<br> Vente timer bliver ikke vist p� print.<br>
		Hvis en sum-aktivitet frav�lges (ikke vises p� faktura print) kan der ikke gemmes fak-timer p� de medarbejdere der passer til denne timepris. De timer der evt. st�r i fak-feltet p� medarbejderen bliver overf�rt til vente-timer.
		Ogs� selvom der efterf�lgende indtastes fak-timer p� en medarbejder bliver der stadigv�k lagt 0-timer ned i databasen hvis sum-aktiviteten ikke er sl�et til. <br><br>
		Vente timer bliver overf�rt som en option p� den n�ste faktura oprettelse.<br>
		Ved hver faktura oprettelse bliver vente-timer 0-stillet inden de nye vente-timer bliver lagt i databasen for den enkelte medarbejder p� det p�g�ldende job.
		<br><br>
		N�r en faktura bliver <b>godkendt</b> kan den ikke l�ngere slettes (P� sigt heller ikke redigeres), og der bliver oprettet en postering p� de valgte konto.<br><br>
		K�rsel vises IKKE p� en faktura.
		<br><br><br>
		
		<b><a name="f-valg">Valg af kunder og job.</a></b>
		<hr>
		Kun aktive job bliver vist p� listen. Brug kundefilteret til venstre for at sortere mellem kunder.<br><br>
		<br>
		
		<b><a name="f-dec">Decimaler.</a></b>
		<hr>
		Afrundering af decimaler kan forekomme, <u>kontroller</u> derfor altid om afrunding passer.<br>
		For hver aktivitet er der f�rst en oversigt over de medarbejdere der er tilknyttet aktiviteten <br>
		via sine projektgrupper og nedenunder er der en sum-aktiviet for hver timepris der findes p� den p�g�ldende aktivitet.
		<br><br><br>
		
		<b><a name="f-tp">Timepriser.</a></b><hr>
		Timeprisen p� en aktivitet er beregnet udfra:<br>
		<u>Fastpris:</u><br>
		Tildelt fastpris p� job / antal fakturerbare timer p� job.<br>
		<u>L�bende timer (budget):</u><br>
		Den valgte timepris for hver enkelt medarbejder p� den aktuelle aktivitet.<br>
		<br>
		<b>Skiftende timepriser undervejs.</b><br>
		Nedenst�ende er et eksempel p� en medarbejder har skiftet timepris undervejs i jobforl�bet.<br><br>
		Dette markeres ved at de benyttede timepriser bliver vist i () efter navnet p� medarbejderen. N�r en medarbejder har skiftet timepris
		kan hans registrerede timer fordeles udover flere sum-aktiviteter, eller slet ikke. Hvis der mangler timer p� sum-aktiviteterne i forhold til det antal timer der er p� medarbejderne
		kan det jursteres ved at klikke p� "calc".<br>
		N�r en medarbejder har antaget flere timepriser benyttes den h�jste altid som beregnings grundlag for medarbejderen.
		<img src="../ill/help_fak1.gif" width="400" height="216" alt="" border="0">
		 
						
		<br><br><br>
		
		<b><a name="f-interval">Datointerval.</a></b><hr>
		Den valgte start og slut dato er begge inkluderet i det valgte interval.<br>
		Det valgte Datointerval bliver gemt i en cookie og bliver vist n�ste gang du bruger faktura-delen.<br><br>
		Hvis der benyttes et datointerval hvor den valgte startdato ligger f�r datoen for den sidst oprettede faktura p� det p�g�ldende job, bruges faktura datoen som startdato n�r nye fakturerbare timer skal findes. 
		<br><br>Faktura datoen oprettes s�ledes, at alle timer til og med faktura datoen er indbefattet. N�r en faktura oprettes fjernes muligheden for at indtaste timer, frem til og med faktura datoen.
		<br><br><br>
		
		<b><a name="f-vente">Vente timer.</a></b><hr>
		Eksisterende vente timer bliver vist i (parantes) ved oprettelse af en ny faktura.
		<br><br>N�r der oprettes en faktura bliver de vente timer, der angives i den ventetime feltet,
		automatisk overf�rt til den n�ste faktura oprettelse. Dette g�lder uanset hvilken periode man v�lger p�
		den n�ste faktura. 
		
		Vente timer bliver vist i ventetime feltet ved rediger s� man kan �ndre antallet af vente timer.<br><br>
		Ved hver faktura oprettelse bliver vente-timer 0-stillet inden de nye vente-timer bliver lagt i databasen for den enkelte medarbejder p� det p�g�ldende job.
		<br><br><br>
		
		<b><a name="f-andre">"Andre?" sum-aktiviteten.</a></b><hr>
		"Andre?" sum-aktiviteten er en ekstra sum-aktivitet for hver aktivitet. "Andre?" sum-aktiviteten kan ikke indeholde en timepris der allerede er brugt p� den p�g�ldende aktivitet. Hvis der v�lges en timepris der allerede er brugt nulstilles "Andre?" sum-aktiviteten.
		<br><br><br>
		
		<b><a name="f-komma">Komma / punktum.</a></b><hr>
		Der kan b�de bruges komma og punktum som decimal deler. Der kan ogs� indtastes f.eks " ,25" hvislket bliver oversat til "0,25".
		<br><br><br>
		 
		
		<b><a name="f-rest">Restbel�b til fakturering.</a></b><hr>
		P� <u>fastpris job</u>, er der angivet restbel�b til fakturering. Dette er differencen mellem det forel�bigt fakturerede bel�b p� jobbet og det bel�b der er angivet som fastpris bel�bet p� jobbet.
		<br><br><br>
		
		<b><a name="f-gamle">Gamle fakturaer.</a></b><hr>
		Fakturaer oprettet f�r 26/10-2004 bilver redigeret i den gamle faktura fil.
		<br><br><br>
		
		<b><a name="f-ikke">Ikke fakturerbare aktiviteter.</a></b><hr>
		Ikke fakturerbare aktiviteter bliver ikke vist p� faktura oprettelsen.<br><br><br>
		
		<b><a name="f-skjulte">"Skjulte" timer.</a></b><hr>
		Timer p� en aktivitet, registreret af medarbejdere der ikke l�ngere har adgang til aktiviteten.
		<br><br><br>
		
		<b><a name="f-ryk">Rykkere og Kreditnotaer.</a></b><hr>
		N�r en faktura er godkendt kan der oprettes rykkere og kreditnotaer til fakturaen.<br>
		Hvis der bliver oprettet en rykker bliver den oprindelige faktura krediteret.<br>
		N�r der oprettes en kreditnota, bliver den oprindelige faktura ikke ber�rt.<br><br>
		Rykkere og kreditnotaer bliver modregnet p� de valgte konti.<br><br>
		Rykkere og kreditnotaer bliver vist med deres egen farve p� fakturalisten og den oprindelige faktura bliver ved oprettelse af rykker "gr�" = Inaktiv, der medf�rer at den ikke l�ngere bliver medregnet i balancen.
		<br><br>
		Farvekoder: <font color="LightPink"><b>Kreditnotaer.</b></font>&nbsp;|&nbsp; (lys gul) <font color="#FFFFE1"><b>Rykkere.</b></font>&nbsp;|&nbsp;<font color="#000000"> (hvid) <b>Fakturaer.</b></font>&nbsp;|&nbsp;<font color="#cccccc"><b>Inaktive Fakturaer.</b></font>
		<br><br>
		
		
		<br><br><br>
		<font class=megetlillesort>Rev. dato: 27/10-2004</font>
		
		</td>
	</tr>
	
	
	
	
	
	
	<%case "gen"%>
	
	<tr>
		<td valign="top" colspan="2"><br><b><a name="flow">TimeOut opbygning / flowchart.</a></b></td>
	</tr>
	<tr>
		<td  colspan="2" height="1" bgcolor="#003399"><img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"  class="alt" style="padding-left:4px;">&nbsp;</td>
		<td valign="top" bgcolor="#ffffff">
		<img src="../ill/TimeOut_opbygning.gif" />
		<br>
	    </td>
	    
	</tr>
	
	
	
	<%
	case "treg"
	
	
	call subindeks_header(menu, "Timeregistrering")
	%>
	    
	    <a href="#t_kalender" class="kalblue">Kalender navigation</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#indtast" class="kalblue">Timeregistrering</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#t_guiden" class="kalblue">Guiden dine aktive job!</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#hvorermitjob" class="kalblue">Hvorfor kan jeg ikke se et job der allerede er oprettet?</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#t_typer" class="kalblue">Aktivitetstyper</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#redigertimer" class="kalblue">�ndre / slette en indtastning</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#lukkedefelter" class="kalblue">Lukkede datoer/felter</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		<a href="#smil" class="kalblue">Smiley-ordning</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		
			
	
	<%call subindeks_footer() %>	
	
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><a name="t_kalender"><b>Kalender navigation</b></a><br />
		Find de datoer du �nsker at registrere tid p� ved at klikke rundt i kalenderen i �verste h�jre hj�rne. Man kan ogs� f�lge med i site daglige timeforbruge i kalenderen. Tidsforbruget er angivet umiddelbart under datoen for hver enkelt dag.
	</td>
	</tr>
		
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><a name="t_guiden"><b>Guiden dine aktive job</b></a><br />
		I "Guiden dine aktive job" er der mulighed for at sl� dine aktive job til og fra. Guiden indeholder alle de job som din projektleder har tilf�jet til de projektgrupper du er medlem af, og hvis status er sat til at v�re aktiv. 
		Hermed kan den enkelte medarbejder selv vedligeholde sin jobliste p� timeregistreringssiden.
	</td>
	</tr>
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><a name="hvorermitjob"><b>Hvorfor kan jeg ikke se et job der allerede er oprettet?</b></a>
	    <br><ul>
		<li>Der er ikke givet adgang til den/de projektgruppe(r) du er medlem af.
		<li>Jobbet er ikke aktivt.
		<li>Jobbet er ikke sl�et til i <a href="#guiden" class="kalblue">Guiden dine aktive job!</a>
		</ul></td>
	</tr>
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><a name="indtast"><b>Timeregistrering</b></a><br />
		
	
		<br><ul>
		<li>Indtast antal timer ud for det job og den dato der �nskes. 
		<li>Klik derefter p� <b>"Indl�s timer"</b>. 
		<li>Der kan indtastes s� mange km som man �nsker i k�rsels felterne.
	    <li>Kommentarer indtastes ud for hver timeregistrering ved at klikke p� +
		<li>Der kan benyttes heltal 0-9.
		<li>"." bruges som seperator.
		<li>Der bruges 100 tals system, s� en � time indtastes som 0.5 og et kvarter som 0.25
		</ul>
		</td>
	</tr>
	
	
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b><a name="t_typer">Aktivitetstyper</a></b> (til timeregistreirng)<br />
		
			 <br /><br />
    
			<%
			call akttyper2009(3)
			%>
			</table>
	
	</td>
	
	</tr>
	
	
		
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b><a name="redigertimer">�ndre / slette en indtastning</a></b>
		<br><br>
		<b>Slette:</b><br>
		Indtast et <b>0 (nul)</b> i det felt ud for det job og den dato, hvor der �nskes at foretages en sletning.<br>
		Klik derefter p� "Indl�s timer". 
		<br><br><b>�ndre:</b><br>
		Overskriv de allerede indtastede timer med det nye antal.<br> 
		Klik derefter p� "Indl�s timer". <br><br>
		<b>F�lgende data bliver opdateret n�r der �ndres timeantal:</b><br>
		<ul>
		<li>Timeantal
		<li>Kommentar
		<li>Timepris 
		<li>Intern Kostpris
		<li>Taste�r
		<li>Aftale relationer
		<li>Tastedato
		<li>Sidst opdateret af.
		<li>Om kommentar skal v�re tilg�ngelig for kunde.
		</ul>
		
		<b>Hvis timer indl�ses igen, uden at timeantal �ndres, bliver f�lgende data opdateret</b>:<br>
		(Timer der f.eks er indl�st tidligere p� ugen)
		<ul>
		<li>Timeantal
		<li>Kommentar
		<li>Sidst opdateret af.
		<li>Om kommentar skal v�re tilg�ngelig for kunde.
		</ul>
		
		
		</td>
	</tr>
	
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b><a name="lukkedefelter">Lukkede felter p� timeregistrerings-siden</a></b><br />
		Felter bliver lukket for yderligere timeregistrering og �ndring af timer hvis:
		<ul>
		<li>Der er oprettet en faktura i den valgte periode (uanset om den er godkendt).
		<li>Periode er afsluttet (s�ttes i kontrolpanel)
		<li>Ugen er afsluttet og i kontrolpanelet er det sl�et til at uger skal lukkes for indtastning ved godkendelse. Administrastorer kan stadigv�k
		�ndre timer indtil der foreligger en faktura.
		<li>Kontakt har godkendt timer i kunde-login delen.
		</ul>   
		</td>
	</tr>
	
	
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b><a name="smil">Smiley-ordning.</a></b>
	
		<br>
		<br><b>Smiley statistik.</b> (findes under statistik.)<br>
		Ved afslutning af en uge inden s�ndag kl. 24:00 vil ugen betragtes som rettidigt afsluttet og medf�re en glad "Smiley"
		Ellers vil medarbejderen modtage en sur Smiley i Smiley statistikken.
		<br><br>
		<b>Dag til dag Smileys.</b><br>
		Hvis der findes uger frem til indev�rende uge <b>(fra 1/11-2005 med mindre andet er aftalt)</b> der ikke er afsluttet (uanset om de er afsluttet forsent) vil man modtage en sur Smiley<br>
		Ellers<br>
		Hvis der, i indev�rende uge, findes dage frem til dagsdato uden timeregistreringer (baseret p� den p�g�ldende medarbejders normerede arbejds-uge), vil der p� timeregistrerings siden v�re en mellemforn�jet smiley. 
		Hvis man derimod er "up to date" vil der v�re en glad Smiley.<br><br>
		
		<b>Regler og definitioner:</b><br>
		- Hvis valgt dato er en l�rdag eller en s�ndag, kigges der tilbage i den forgangne uge.<br>
		- Hvis valgt dato er en mandag, vil man altid modtage en glad Smiley.<br>
		- Der bliver ikke taget h�jde for om der er indtastet 1 eller 12 timer af medarbejderen, kun at denne har registreret timer p� de p�g�ldende dage.<br><br>
		
		<b>Aktivering af Smiley funktion.</b><br>
		Smiley funktionen kan sl�s til og fra i kontrolpanelet af administratorer.<br>
		<br>
		Hver medarbejder kan ogs� sl� visningen af Smiley-ordningen fra for deres medarbejderprofil. Dette g�res i deres medarbejderprofil.
		<br><br>
		<!--Selvom en uge afsluttes �ndres der ikke status p� timerne. 
		De vil alle have status "venter p� godkendelse" indtil de bliver godkendt i joblog'en. <br><br><br>&nbsp;
		-->
		
		<br><br>
		&nbsp;
		</td>
	</tr>
	
	
	<%
	case else
	
	call subindeks_header(menu, "TimeOut - Kom godt igang")
	%>
	
	F�lg disse 6 punkter, og du vil gennemg� en kort introduktion og I
	vil v�re igang med TimeOut p� 10 min.<br /><br />
	
	<a href="#g-oprmedtyp" class="kalblue">1) Opret Medarbejdertyper</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
	<a href="#g-oprmed" class="kalblue">2) Opret Medarbejdere</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
	<a href="#g-oprkun" class="kalblue">3) Opret Kontakter (kunder/debitorer)</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
	<a href="#g-oprstam" class="kalblue">4) Opret Stam-aktiviteter</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
	<a href="#g-oprproj" class="kalblue">5) Opret Projektgrupper</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
	<a href="#g-oprjob" class="kalblue">6) Opret Job</a>&nbsp;&nbsp;<img src="../ill/pile_selected.gif" width="10" height="6" alt="" border="0"><br>
		
	
	<%call subindeks_footer() %>	
	
	
	
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b><a name="g-oprmedtyp"> 1) Opret medarbejdertyper</a></b><br />
		Timepriser og antal arbejdstimer pr. uge.<br /><br>
		<b>Findes under: TSA -- Medarbejdere -- Medarbejdertyper</b>
		<br /><br /> 
		
		Medarbejdertyper bruges til at samle jeres medarbejdere i grupper efter hvilken type de er. 
		F.eks: Direkt�r, Senior konsulent, konsulent, bogholder, kreativ etc.<br /><br />
		
		Hver type kan have forskellige arbejdstider og timepriser. Dette angives for hver medarbejdertype der
		er oprettet. 
		Der kan oprettes lige s� mange medarbejdertyper som I har brug for.<br /><br />
		Hvis i har flere medarbejdere der arbejder fuld tid og til samme timepris*, samles disse medarbejdere i samme gruppe. 
		Herefter vil alle medarbejdere af denne type have det samme timeantal pr. uge og vil blive ud-faktureret til den samme timepris. Der skal ogs� angives en intern kostpris pr. time.
		 
		
		  <br /><br />
            *) Med mnidre et job angives til fastpris, hvorefter medarbejder timepriser ignoreres p� det p�g�ldende job.
            <br /><br />
		&nbsp;
		 </td>
	</tr>
	
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b><a name="g-oprmed"> 2) Opret medarbejder</a></b><br />
		Medarbejderprofil, login og password.<br /><br>
		<b>Findes under: TSA -- Medarbejdere </b>
		<br /><br /> 
		Opret medarbejdere bruges til at oprette jeres medarbejdere og vedligeholde jeres TimeOut licens, der er baseret p� antal aktive medarbejdere.
		<br /><br />
		Det er her i angiver medarbejder navn, login, password mm.<br />
		V�lg den korrekte medarbejdertype, hvis den findes, og hvis den ikke findes g� da tilbage til 
		<a href="#g-oprmedtyp" class="kalblue">1) Opret medarbejdertyper</a>.
		
		<br /><br />
		V�lg ogs� hvilke adgangsrettigheder denne medarbejder skal have til TimeOut. Dette g�res under 
		<a href="#brugergp" class="kalblue">brugergruppe</a>. Der findes 7 forskellige brugergrupper i TimeOut, alle med forskelle rettigheder.
		Superbrugere s�ttes til at v�re <b>administrator</b>, og brugere der kun skal registrere timer skal v�re i guppen <b>TSA niveau 2</b>. 
	    <br /><br />
	    N�r der oprettes en medarbejder sendes der en email til den valgte email adresse med link til TimeOut og med medarbejderens login og password.
	   <br /><br />
	   
            <img src="../ill/help_oprmed.gif" style="border:2px orange solid;"/>
	    <br /><br />
		&nbsp;
		 </td>
	</tr>
	
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b><a name="g-oprkun"> 3) Kontakter (kunder/debitorer)</a></b><br />
		Stamdata p� kontakter, kontaktpersoner og aftaler.<br /><br>
		<b>Findes under: TSA -- Kontakter </b>
		<br /><br /> 
		Kontakter er en oversigt og samling af alle jeres kunder, leverand�rer, salgsleads mm.
		<br /><br />
		Kontakter kan grupperes efter type (type s�ttes i kontrolpanelet) og efter kontakt ansvarlig. (v�lges blandt de oprettede medarbejdere)
		<br /><br />
		N�r du opretter en ny kontakt skal du indtaste almindelige stamdata, adresse, postnr., by, telefon nr. etc.
		N�r kontakten er oprettet f�r man mulighed f�r, ved at klikke p� kontaktnavnet p� kontaktoversigten, at tilf�je kontaktpersoner og evt. aftaler.<br /><br />
		Aftaler kan bruges som ramme-aftale s�ledes at man kan knytte flere job op p� den samme aftale. Aftaler kan ogs� st� alene, som selvst�ndige aftaler der faktureres s�rskilt.<br /><br />
		I kan oprette s� mange kontakter som i �nsker. N�r man har oprettet sin(e) kontakt(er) er man klar til at oprette job.
		<br /><br />
            <img src="../ill/help_kunder.gif" style="border:2px orange solid;" />
		<br /><br />
		&nbsp;
		 </td>
	</tr>
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b><a name="g-oprstam"> 4) Stam-aktiviteter</a></b><br />
		Grupper af aktiviteter til at genbruge igen og igen ved job oprettelse.<br /><br>
		<b>Findes under: TSA -- Job -- Stam-aktiviteter og Grp. </b>
		<br /><br /> 
		For at det skal v�re s� hurtigt som muligt, og med s� f� klik som muligt at oprette et job, bruger TimeOut begrebet Stamaktiviteter.
		Stamaktiviteterne samles i grupper der kaldes Stamaktivitets-grupper. Hver Stamaktivitets-gruppe kan indeholde <i>x</i> antal aktiviteter. 
		N�r der oprettes et nyt job, tilknytter man s�ledes en eller flere stam-aktivitetsgrupper istedet for at skulle oprette 5, 10 eller m�ske 20
		aktiviteter hver gang.   
		<br /><br />
		N�r aktiviteterne bliver tilknyttet et job, bliver der lavet en kopi af Stam-aktivitetsgruppen's aktiviteter ud p� jobbet. 
		Aktiviteterne har herefter ingen forbindelse til den Stam-aktivitetsgruppe de kom fra.
		<br /><br />
		Timeregistrering sker p� aktivitets niveau.
		<br /><br />
            <img src="../ill/help_stamakt.gif" style="border:2px orange solid;" />
		<br /><br />
		&nbsp;
		 </td>
	</tr>
	
		<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b><a name="g-oprproj"> 5) Projekgrupper</a></b><br />
		Grupper af medarbejdere der kan tilknyttes til timregistrering p� job.<br /><br>
		<b>Findes under: TSA -- Job -- Projektgrupper. </b>
		<br /><br /> 
		Projektgrupper bruges til at styre hvilke medarbejdere der skal have adgang til et job, og dermed have mulighed for at tidsregistrere p� det p�g�ldende job.
		Hermed kan man tildele job til bestemte personer eller til hele afdelinger eller grupper i virksomheden. <br /><br />
		Som medarbejer bliver man ikke forstyrret af alle mulige job som man reelt ikke arbejder p�. Kun job man er tilknyttet via sine tilh�rsforhold
		til projektgrupper, bliver vist p� medarbejderens timeregistreringsside.
		<br /><br /> 
		En medarbejder kan v�re med i mange projektgrupper. Der kan v�re fra 1 til <i>x</i> antal medarbejdere i hver projektgruppe. Der kan tilknyttes optil 10 projektgrupper p� et job og p� de underliggende aktiviteter.
		<br /><br />
            <img src="../ill/help_pro.gif" style="border:2px orange solid;" />
		
		<br /><br />
		&nbsp;
		 </td>
	</tr>
	
	
	
	<tr>
		<td colspan=2 valign="top" style="padding-top:30px;"><b><a name="g-oprjob"> 6) Job</a></b><br />
		Job stamdata, tilknytning til kontaker.<br /><br>
		<b>Findes under: TSA -- Job </b>
		<br /><br /> 
		<b>V�lg kontakt</b><br />
		N�r et job oprettes skal man i step 1 v�lge de(n) kontakt dette job skal oprettes p�.
		Et job kan godt oprettes p� flere kontakter samtidig, s�ledes at der bliver oprettet et job p� hver kontakt der v�lges.
		<br /><br />
		<b>Stamdata</b><br />
		N�r man har valgt kontakt forts�tter man til step 2 i joboprettelsen. Her angives job-stamdata dvs. jobnavn, jobbeskrivelse etc.
		<br /><br />
		<b>Tilbud</b><br />
		Det er muligt at oprette jobbet som et tilbud indtil at kunden godkender tilbudet. Dette g�res ved at s�tte flueben i 
		[<i>V</i>] "Brug dette job som tilbud".  Herefter kan jobbet overg� til at v�re et aktivt job ved at fjerne fluebenet igen.
		<br /><br />
		<b>V�rdi og type</b><br />
		Her angives det forventede budget p� jobbet, eller hvis der v�lges "fastpris", g�lder det indtastede bel�b som den pris man har aftalt med kunden.
		Ved fastpris bliver alle timer omsat med prisen p� jobbet divideret med det antal fakturerbare timer der angives i <b>forkalkuleret</b> timeforbrug.<br />
		Hvis der v�lges l�bende timer, bliver v�rdien p� jobbet behandlet som vejledende og de timepriser der er angivet p� 
		<a href="#g-oprmedtyp" class="kalblue">medarbejdertypen</a> bliver benyttet. En medarbejder kan godt antage forskellige timepriser p� de aktiviteter der er tilknyttet til jobbet. 
		
		<br /><br />
		<b>Forkalkuleret timer</b><br />
		Her angives det forventede / aftalte antal timer p� jobbet.
		
		<br /><br />
		<b>Status</b><br />
		V�lg status og forventet job-periode. Kun aktive job blive vist p� timeregistrerings siden.
		
		<br /><br />
		<b>Kontakperson</b><br />
		Her kan der v�lges en kontaktperson hos kunden. Der kan v�lges imellem de kontakpersoner der er oprettet p� hver enkelt kontakt.
		
		<br /><br />
		<b>Aftale</b><br />
		Det er muligt at tilknytte jobbet til en aftale, hvorefter der bliver taget enheder fra denne aftale hver gang der timeregistreres p� dette job. Og det er muligt at fakturere jobbet sammen med andre job der er tilknyttet den samme aftale.
		
		<br /><br />
		<b>Aktiviteter</b><br />
		Ved joboprettelse v�lger man hvilke af de pre-definerede stamaktivitets grupper der skal tilknyttes dette job. Herved kan man fra en enkelt selectboks tildele 5, 10 eller 20 aktiviteter alt efter hvor mange aktiviteter der ligger i den valgte gruppe.
		Det er muligt at kombinere op til 5 forskellige grupper p� en gang. F.eks: Presales, Foranalyse, Produktion, Fakturering etc.
		<br /><br />
		Efterf�lgende kan man v�lge om de aktivteter der ligger i den enekelte stam-aktivitetsgruppe skal beholde det timeantal og timpriser de er f�dt med.
		<br /><br />
		N�r jobbet er oprettet kan der tilknyttes 5 nye stam-aktivitetsgrupper, eller man kan klikke p� fanebladet "Aktiviteter" for at tilf�je flere aktiviteter manuelt.
		<br /><br>
		<b>Projektgrupper</b><br />
		Efter tilf�jelse af aktiviteter, er det nu tid til at bestemme hvilke <a href="#g-oprproj" class="kalblue">projektgrupper</a> og dermed hvilke medarbejdere der kan registrere tid p� dette job. Dette g�res ved at tilknytte de relevante projektgrupper til jobbet.  Der kan tilknyttes op til 10 projektgrupper.
		<br /><br />Hvis man er administrator, niveau 1 bruger eller angivet som jobansvarlig, kan man selv fra timregistrerings siden tilknytte sig til et job, hvis man ikke er blevet tilknyttet jobbet fra starten. 
		
		
		
		
		<br /><br />
            <img src="../ill/help_job.gif" style="border:2px orange solid;" />
            
            <br /><br />
		<b>Jobbet er nu oprettet og I er klar til timeregistrering!</b>
		
		<br /><br />
		
		&nbsp;
		 </td>
	</tr>
	
	
	
	
	
	</table>
	
	
	<% 
	end select
	%>
</table>
<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
&nbsp;
</div>
<!--#include file="../inc/regular/footer_inc.asp"-->

