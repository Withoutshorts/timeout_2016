


<!--#include file="../xml/error_xml_inc.asp"-->

<%
'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
 'Response.write "<h4>thisfilr: "& thisfile &"error_inc XX</h4><br>"


function showError(errortype)
select case errortype
case 1
varErrorText = "Der er ikke tastet noget brugernavn!"
case 2
varErrorText = "Der er ikke tastet noget password!"
'case 3
'varErrorText = "Denne medarbejderkonto er <b>ukendt.</b> <br>Du har derfor ikke adgang til <b>TimeOut!</b>"
case 4
varErrorText = replace(err_txt_004, "#nylinie#", "<br>") 
case 5
if Request.Cookies("loginlto")("lto") <> "" then
thislto = Request.Cookies("loginlto")("lto")
else
    if len(trim(lto)) <> 0 then
    thislto = lto
    else
    thislto = ""
    end if
end if
varErrorText = "<b>Sessionen er udl�bet og du er blevet logget af TimeOut.</b><br><br>"_
&"Sessionen udl�ber ud hvis TimeOut har v�ret passiv i mere end <b>12 timer</b>, "_
&"eller TimeOut serveren har nulstillet dagens aktive sessioner, af sikkerheds- og -performance m�ssige hensyn."_
&"<br><br>Klik her:  <a href='https://outzource.dk/"&thislto&"' target='_top'>https://outzource.dk/"&thislto&"</a> for at logge ind i jeres system igen."_
&"<br /><br />"_
&"Med venlig hilsen<br /><br />OutZourCE dev. team."
case 6
varErrorText = "Du har glemt at indtaste antal timer."
case 7
varErrorText = "Dette jobnr er allerede indtastet en gang p� denne dato.<br>"_
&"Da et jobnr ikke m� optr�de mere en en gang pr. dato, m� du �ndre indtastningen.<br>"_
&"<i><a href='javascript:history.back()'>�ndre indtastning</a></i>"_
&"<br>"


case 8
varErrorText = "En eller flere af nedenst�ende felter mangler at blive udfyldt:<ul><li>Der mangler at blive indtastet et <b>navn</b></ul>"

case 9
varErrorText = "Der mangler at blive indtaste en f�lgende informationer:<br><ul>"_
&"<li>Medarbejder navn"_
&"<li>Medarbejer nr."_
&"<li>Login"_
&"<li>Password"_
&"</ul>"



case 10
varErrorText = "En Medarbejder med det valgte <b>Medarbejder nr.</b> eksisterer allerede.<br><br>V�lgt et andet."

case 11
varErrorText = "En Medarbejder med det valgte <b>Login</b> eksisterer allerede. <br><br>V�lgt et andet."

case 12
varErrorText = "En Kontakt med det valgte <b>Kontakt id</b> eksisterer allerede.<br>"_
&"Den eksisterende Kontakt er: <b>"& errKundenavn &" ("& errKundenr &")</b>"_
&"<br>V�lg et andet <b>Kontakt id</b>."


case 13
varErrorText = "<ul><li>Der mangler at blive indtastet et <b>kundenavn</b></ul>"

case 14
varErrorText = "Der mangler at blive indtastet en af f�lgende informationer: <br> "_
&"<ul> "_
&"<li>Jobnavn."_
&"<li>Jobnavn indeholder et <b>''</b> eller et <b>'</b>"_ 
&"<li>Jobnavn er p� <b>mere end 100</b> karakterer. ("& len(request("FM_navn")) &")"_
&"<li>Jobnr. "_
&"<li>Jobnr. er p� <b>mere end 20</b> karakterer."_
&"<li>Startdato er en <b>senere dato </b>end slutdato."_
&"<li>Der mangler at blive tilknyttet en <b>kunde</b>.</ul>"

case 15
varErrorText = "<b>Job nr.</b> eller <b>Tilbudsnr.</b> er ikke angivet som et nummer / heltal.<br>"_
&"Job nr. m� kun indeholde tallene 0-9 uden punktum eller komma. Job nr. m� ikke v�re 0 NUL."

case 16
varErrorText = "<b>Brutto Oms�tning</b> er ikke angivet i et korrekt format.<br>"_
&"Dvs. feltet indeholder andre tegn end tallene 0-9, samt komma eller punktum."

case 17
varErrorText = "Der er sket en <font class='error'>Fejl!</font> i indtastningen.<br>"_
&"I et eller flere felter er der blevet indtastet ugyldige karakterer."

case 18
varErrorText = "<font class='error'>Fejl!</font><br>"_
&"Der er indtastet mere end <b>24 timer</b> i et af felterne."

case 19
varErrorText = "Der foreligger allerede en registrering af dette<br> <b>Job</b> med denne <b>aktivitet</b> p� denne <b>dato</b>."

case 20
varErrorText = "<b>Timer</b> er ikke angivet som et tal.<br><br>Husk at angive et [.] (punktum) og <br>ikke et [,] (komma) hvis der angives halve timer."
case 21
varErrorText = "<b>Kontakt id</b> er ikke udfyldt, eller det er over 15 karakterer langt."
case 22
varErrorText = "<b>Medarbejdernr</b> er ikke angivet som et nummer."

case 23
varErrorText = "<b>Aktiviteter</b><br><br>"_
&"Der mangler at blive indtastet en af f�lgende informationer: <br><br> "_
&"- Aktivitets navn<br>"_
&"- Navn indeholder en ulovlig karakter (<b>apostrof.</b>).<br>"_

'&"<li>Startdato er en <b>senere dato </b>end slutdato."_

case 24
varErrorText = "Et job med det valgte <b>Job nr.</b> eller <b>Tilbudsnr.</b> eksisterer allerede.<br>G� tilbage og tildel jobbet et nyt jobnummer.<br><br>Ofte vil det v�re nok at addere nummeret med 1"

case 25
varErrorText = "Du har brugt mere end <b>3 fors�g</b> p� at logge ind.<br>Dette medf�rer at du ikke l�ngere kan logge ind i systemet.<br><br><a href='javascript:window.close()'>Luk browseren</a> og pr�v igen <br>eller Kontakt din TimeOut 2.1 administrator for at f� et nyt login."

case 26
varErrorText = "<b>ERP - Opret Faktura!</b><br>"_
&"Der mangler at blive indtaste en f�lgende informationer:<br><ul>"_
&"<li>Timer"_
&"<li>Bel�b"_
&"</ul>"

case 27
varErrorText = "<b>Bel�b</b> er ikke angivet som et hel-tal.<br>"_
&"Dvs. feltet indeholder andre tegn end tallene 0-9, samt komma eller punktum."

case 28, 135
varErrorText = replace(err_txt_028, "#nylinie#", "<br>") &" "& tTimertildelt(y) 
    
    'if errortype = 135 then
    'varErrorText = varErrorText & " - "& errortype & " - " & multiDagArr
    'end if

case 29
varErrorText = "<b>Faktura nr.</b>findes alledrede i systemet. V�lg et andet."

case 30
varErrorText = "<b>Timer</b> skal v�re > 0 n�r budget er angivet som fast pris."
'case 31 'Not in use
'varErrorText = "<b>Timer</b> skal v�re mellem 0 og 24, og <b>minutter</b> skal v�re mellem 0 og 60."
case 31 
varErrorText = "Sandsynlighed indeholder ulovlige karakterer."
case 32
varErrorText = "Aktivitet start datoen er en tidligere dato end job start datoen, som er:<br> <b>"& formatdatetime(jobstdate, 1) &"</b>"

case 33
varErrorText = "Aktivitet slut datoen er en senere dato end job slut datoen, som er: <br><b>"& formatdatetime(jobsldate, 1) &"</b>"

case 34
varErrorText = "Der er mere end 1255 karakterer i kommentaren!"
case 35
varErrorText = "Der er tildelt flere timer p� denne aktivitet,<br> end der er til r�dighed p� jobbet.!"
case 36
varErrorText = "Den valgte slut dato skal v�re en senere dato end den valgte start dato."
case 37
varErrorText = "Den valgte dato for den opf�lgende aktion skal v�re en senere dato en den f�rst oprettede aktion."
case 38
varErrorText = "Da den valgte dato for den opf�lgende aktion er den samme som for den f�rst oprettede aktion, skal start tidspunktet for den opf�lgende aktion v�re senere end slut tidspunktet p� den f�rste aktion."
case 39
varErrorText = "De angivne <b>Time-</b> eller <b>Interne kost</b> <b>-priser</b> er:<ul>"_
&"<li>Ikke angivet som et <b>hel-tal</b> (Indeholder et komma)."_
&"<li>Indeholder et <b>ugyldigt tegn</b> eller bogstav."_
&"<li>Et af dem er angivet som et <b>negativt tal</b>."_
&"<li>Et af dem er <b>tomme</b> (ikke udfyldt).</ul>"
case 40
varErrorText = "Der mangler at blive valgt et eller flere af f�lgende felter:<br>"_
&" <ul><li><b>Medarbejdere</b><li><b>Job</b></ul>"
case 41
useleftdiv = "j"
varErrorText = "Der mangler at blive valgt en eller flere af f�lgende felter:<br>"_
&" <ul><li><b>Ugedage</b><li><b>Datoer</b><li><b>M�neder</b></ul>"

case 42
varErrorText = "<b>Faktura nr.</b> er ikke angivet som et tal."

case 43
useleftdiv = "c"
varErrorText = "Der mangler at blive indtastet en af f�lgende informationer: <br> "_
&"<ul> "_
&"<li>Emne</ul>"
case 44
useleftdiv = "c"
varErrorText = "Startdatoen er af nyere dato iforhold til slutdatoen.<br> "
case 45
useleftdiv = "c"
varErrorText = "Der er ikke valgt hvilke medarbejdere der skal oprettes p� opgaven.<br> "
case 46
varErrorText = "<ul><li>Der mangler at blive indtastet et <b>navn</b> eller et <b>kontonr</b>."_
&"<li>Eller en konto med det valgte kontonr findes allerede.</ul>"
case 47
varErrorText = "<ul><li>Der mangler at blive indtastet et <b>navn</b> eller en <b>momskvotient</b>.</ul>"
case 48
varErrorText = "<ul><li>Momskvotient'en indeholder ugyldige karakterer.</ul>"
case 49
varErrorText = "<ul>Der er angivet ugyldige karakterer i en eller flere af de f�lgende felter.</ul>"_
&"<li>Nettobel�b.<li>Momsbel�b.<li>Totalbel�b."
case 50
varErrorText = "<ul><li>Der mangler at blive indtastet <b>Totalbel�b</b>.</ul>"
case 51
varErrorText = "<b>Faktor</b> er ikke angivet som et tal.<br><br>Husk at angive et [.] (punktum) og <br>ikke et [,] (komma) hvis der angives halve enheder."
case 52
varErrorText = "Mindst en af f�lgende kriterier er ikke opfyldt:<ul><li>Der er ikke valgt om fakturaer skal godkendes uden gennemsyn.<br><b>(Kvik fakturering er ikke sl�et til!)</b><br><br>"
varErrorText = varErrorText & "<li>Der er <b>ikke</b> valgt hvilke <b>aftaler</b> der skal faktureres.</ul>"
case 53
varErrorText = "Der er ikke valgt en folder.!"
case 54
varErrorText = err_txt_054 '"<b>Antal</b> er ikke angivet som et tal!"
case 55
varErrorText = "<b>Indk�bspris</b> er ikke angivet som et tal!"
case 56
varErrorText = "<b>Salgspris</b> er ikke angivet som et tal!"
case 57
varErrorText = "<b>Produktionstid</b> er ikke angivet som et tal!"
'case 58
'varErrorText = "<b>Logud tidspunkt</b> er ikke angivet i et tilladt format!"
case 59
varErrorText = "En eller flere <b>Komme / G� tidspunkt(er)</b> er ikke angivet i et gyldigt dato- og/eller tids -format!<br><br>Der er angivet ulovlige karakterer eller der mangler angivelse."
case 60
varErrorText = "En eller flere <b>Komme / G� tidspunkt(er)</b> er ikke angivet i et gyldigt dato- og/eller tids -format!<br><br>Der er angivet ulovlige karakterer eller der mangler angivelse."
case 61
varErrorText = "<ul><li>Der mangler at blive indtastet et navn"
varErrorText = "<li>Bel�b er ikke indtastet som tal</ul>"
useleftdiv = "c"
case 62
varErrorText = "<ul><li><b>Kontonr</b> er ikke angivet som et tal.</ul>"
case 63
varErrorText = "En eller flere af de indtastede <b>tids-angivelser</b> indeholder ulovlige karakterer. <br><br>"_
&"Registreringerne m� kun indeholde <b>tallene 0-9, samt kolon</b>."
case 64
varErrorText = "En eller flere af de indtastede <b>tids-angivelser</b> er ikke et gyldigt <b>tidsformat.</b> <br><br>"_
&"Tidsformatet skal v�re f.eks <b>9:30, 12:59 etc.</b><br>Den angivne tid er: <b>" & errTid & "</b>"
case 65
varErrorText = "En eller flere af de indtastede <b>start tidspunkter</b> er senere end det tilsvarende <b>slut tidspunkt.</b>"
case 66
varErrorText = "<b>Stempelur - logindhistorik</b>- Dato er ikke angivet i et korrekt datoformat, <br>eller den valgte dato er ikke gyldig."
case 67
varErrorText = "En eller flere af de indtastede <b>Forecast</b> indeholder ulovlige karakterer.<br>"_
&"Forecast m� kun indeholde <b>tallene 0-9</b>, samt <b>komma </b>eller <b>punktum</b>."
case 68
varErrorText = "<b>Der mangler en eller flere af f�lgende informationer:</b><br> - Der er ikke angivet et Emne."
case 69
varErrorText = "<b>Der mangler en eller flere af f�lgende informationer:</b><br> - Der er ikke angivet en Kontakt."
case 70
varErrorText = "<b>Der mangler en eller flere af f�lgende informationer:</b><br> - Der er ikke angivet en Incident Kategori."
case 71
varErrorText = "<b>Der mangler en eller flere af f�lgende informationer:</b><br> - Der er ikke angivet en Incident Type (Prioitet)"
case 72
varErrorText = "<b>Der mangler en eller flere af f�lgende informationer:</b><br> - Der er ikke angivet en Incident Status."
case 73
varErrorText = "<b>Der mangler en eller flere af f�lgende informationer:</b><br> - Der er ikke angivet et varenr, eller varenr er = 0."_
&" Det er ikke lovligt at angive <b>0 (NUL)</b> som varenr."
case 74
varErrorText = replace(err_txt_074, "#nylinie#", "<br>")
case 75
varErrorText = "<b>Minimums lager</b> er ikke angivet som et tal!"
case 76
varErrorText = "<b>Antal</b> er ikke angivet som et tal!"
case 77
varErrorText = "<b>Der skal angives en kontakt!</b><br>"_
&" - Hvis der ikke er nogen kontakter tilg�ngelige p� listen,<br>"_
&" er det fordi der ikke er angivet ServiceDesk prioitets gruppe p� nogen kontakter."
case 78
varErrorText = "<b>ServiceDesk, �bningstider.</b><br>"_
&" Der er angivet en ulovlig karakter!"_
&" <br><br> - En eller flere af de angivne �bningstider er ikke udfyldt korrekt.<br>"_
&" Felterne skal  indeholde en gyldig tidsangivelse. Eks: 08:30, 14:35 etc."

case 79
varErrorText = "<b>ServiceDesk, �bningstider.</b><br>"_
&" - En af de angivne lukketider er tidligere en den tilsvarende �bningstid."

case 80
varErrorText = "<b>Stempelur - ignorer periode!</b><br>"_
&" Der er angivet en ulovlig karakter i ignorer periode for stempelur!"_
&" <br><br> - Start eller Slut tid til  er ikke udfyldt korrekt.<br>"_
&" Felterne skal indeholde en gyldig tidsangivelse. Eks: 08:30, 14:35 etc."

case 81
varErrorText = "<b>Stempelur - standard pause!</b><br>"_
&" Der er angivet en ulovlig karakter i en af de angivne<br> standard pauser <b>(A ell. B)</b>."_
&" <br><br>"_
&" - Felterne m� kun indeholde tallene fra 0 - 9."

case 82
varErrorText = "<b>Stempelur - standard pauser!</b><br>"_
&" N�r der er valgt et pause minuttal, der afviger fra de angivne standard pauser, skal der tilf�jes en kommentar."

case 83
varErrorText = "<b>Fakturering - Fejl</b><br>"_
&"Du mangler at beslutte om det er et job eller aftale du �nsker at fakturere."


case 84
varErrorText = "<b>ERP - Fakura masker!</b><br>"_
&" Der er angivet en ulovlig karakter i en af de angivne<br> faktura masker. <b>(Faktura, kladde, intern, ell. kreditnota)</b>."_
&" <br><br>"_
&" - Felterne m� kun indeholde tallene fra 0 - 9."


case 85
varErrorText = "<b>ERP - Fakura Rykker!</b><br>"_
&" Den valgte dato er ikke et gyldigt dato format / datoen findes ikke."


case 86,87
varErrorText = "<b>ERP - Fakura Rykker!</b><br>"_
&"Antal ell. bel�b indeholder ulovlige karakterer. <br>Disse felter m� kun indeholde tal."

case 88
varErrorText = "<b>ERP - Valg af job eller aftale!</b><br>"_
&"Der mangler at blive valgt det job eller den aftale som skal faktureres."

case 89
varErrorText = "<b>ERP - Fakturanr findes i forvejen!</b><br>"_
&"Du er ved at oprette en faktura med et fakturanummr der findes i forvejen.<br>"_
&"Tjek seneste faktura nummer i kontrolpanelet, og sammenhold med seneste oprettede faktura. (s�g i fakturaer).<br>"_
&"<br>Hvis du ikke har adgang til kontrolpanelet skal du have fat i din administrator."

case 90
varErrorText = regdato &" - "& replace(err_txt_090, "#nylinie#", "<br>")

case 91
varErrorText = "<b>TSA - Jobnr og tilbudsnr!</b><br>"_
&" Der er angivet en ulovlig karakter i en af de angivne<br> jobnr / tilbudsnr masker."_
&" <br><br>"_
&" - Felterne m� kun indeholde tallene fra 0 - 9."

case 92
varErrorText = "<b>TSA - Manglende information!</b><br>"_
&" - Der mangler at blive valgt en kunde i step 1."_
&" <br><br>"

case 93
varErrorText = "<b>TSA - Jobnummer findes i forvejen!</b><br>"_
&" - Det �nskede jobnummer <b>("& jobnrFindesNR &")</b> er allerede ibrug, nummeret kan derfor ikke benyttes.<br>"_
&"<br>Hvis der oprettes flere job samtidig, <u>er</u> job frem til f�r dette jobnummer blevet oprettet. Jobnrr�kkef�lgen skal �ndres i kontrolpanelet."_
&" <br><br>"

case 94
varErrorText = "<b>TSA - Tilbudsnummer findes i forvejen!</b><br>"_
&" - Det �nskede tilbudsnummer <b>("& tilbudsnrFindesNR &")</b> er allerede ibrug, nummeret kan derfor ikke benyttes.<br>"_
&"<br>Hvis der oprettes flere job samtidig, <u>er</u> job frem til f�r dette tilbudsnummer blevet oprettet. Tilbudsnummer-r�kkef�lgen skal �ndres i kontrolpanelet."_
&" <br><br>"

case 95
varErrorText = "<b>ERP - Kontonummer</b><br>"_
&" - Det angivne kontonummer er ikke angivet korrekt.<br>"_
&" <br><br>"

case 96
varErrorText = "<b>ERP - faktura moms %</b><br>"_
&" - Den angivne procentsats er ikke angivet korrekt. (hel-tal)<br>"_
&" <br><br>"

case 97
varErrorText = "<b>ERP - Momsafslutning</b><br>"_
&" - Den angivne kommentar er for lang. Den m� maks v�re 255 karakterer."_
&" <br><br>"

case 98
varErrorText = "<b>ERP - Momsafslutning</b><br>"_
&" - Den angivne dato er ikke angivet i et korrekt datoformat.<br>"_
&"Du har indtastet: <b>"& thisafslutdato &"</b><br>"_
&" <br><br>"

case 99
varErrorText = "<b>ERP - posteringsdato</b><br>"_
&" - Den angivne posteringsdato (fakturadato) ligger i en afsluttet momsperiode."_
&" Faktura / posterings dato: <b>"& cdate(tjkPosDato) &"</b><br>"_
&" Seneste momsperiode er afsluttet d. <b>"& cdate(momsafsluttetDato) &"</b><br>"_
&" <br><br>"

case 100
varErrorText = "<b>Materialeordre</b><br>"_
&" - Den angivne dato er ikke angivet i et korrekt datoformat.<br>"_
&"Du har indtastet: <b>"& odrdato &"</b><br>"_
&" <br><br>"

case 101
varErrorText = "<b>Materialeordre</b><br>"_
&" - Et af de angivne antal, p� en af ordre linierne, er ikke angivet som et <b>tal</b>.<br>"_
&"Du har indtastet f�lgende: <b>"& vareantal &"</b><br>"_
&" <br><br>"

case 102
varErrorText = "<b>Materialeordre</b><br>"_
&" - En ordre med det valgte ordrenr findes i forvejen.<br>"_
&"Sidst oprettede ordrenr er: <b>"& lastordnr &"</b><br><br>"


case 103
varErrorText = "<b>Materialeordre</b><br>"_
&" - Det valgte ordrenr indeholder ulovlige karakterer.<br>"_
&"Det angivne ordrenr er: <b>"& odrid  &"</b><br><br>"

case 104
varErrorText = "<b>Materialegrupper</b><br>"_
&" - Gruppenavn mangler at blive udfyldt.<br><br>"


case 105
varErrorText = "<b>Materialegrupper</b><br>"_
&" - Den valgte avance % indeholder ulovlige karakterer.<br>"_
&"Eller mangler at blive udfyldt. Angiv et heltal fra 0-100.<br><br>"


case 106
useleftdiv = "s"
varErrorText = replace(err_txt_106, "#nylinie#", "<br>")

case 107

varErrorText = replace(err_txt_107, "#nylinie#", "<br>")


case 108

'varErrorText = replace(err_txt_108A, "#nylinie#", "<br>")
'varErrorText = varErrorText & "&nbsp;"& lastFakdato
varErrorText = varErrorText & replace(err_txt_108B, "#nylinie#", "<br>")
varErrorText = varErrorText & "&nbsp;"& regdato
varErrorText = varErrorText & replace(err_txt_108C, "#nylinie#", "<br>")

case 109
varErrorText = "<b>Budget Brutto/Netto Dage & Timer</b><br>"_
&"Der mangler at blive angivet et navn."



case 110
varErrorText = "<b>Valutaer</b><br>"_
&"Der mangler at blive angivet et navn p� valutaen."


case 111
varErrorText = "<b>Valutaer</b><br>"_
&"Der mangler at blive angivet en valutakode."


case 112
varErrorText = "<b>Valutaer</b><br>"_
&"Kurs er ikke angivet som et tal.<br>"_
&"Eller kursen er sat = 0, hvilket ikke er lovligt."

case 113
varErrorText = "<b>Budget medarbejdere</b><br>"_
&"Et eller flere felter indeholder ulovlige karakterer (A-�), eller feltet er tomt."

case 114
varErrorText = "<b>Medarbejder</b><br>"_
&"Ansatdato indeholder ulovlige karakterer og er ikke et gyldigt datoformat, eller feltet er tomt."

case 115
varErrorText = "<b>Medarbejder</b><br>"_
&"Opsagtdato indeholder ulovlige karakterer og er ikke et gyldigt datoformat, eller feltet er tomt."

case 116
varErrorText = "<b>Stopur</b><br>"_
&"<b>Starttid</b>, ell. <b>sluttid</b> indeholder ulovlige karakterer og er ikke et gyldigt datoformat, eller feltet er tomt."


case 117
varErrorText = "<b>Timeregistrering</b><br>"_
&" Den registrering Du fors�ger at foretage er blevet afvist. Det skyldes en af to f�lgende grunde:<br><br> "_
&" <b>A)</b> Registreringsdato "_
&" ligger i et datointerval der allerede er faktureret.<br><br>"_
&" <b>B)</b> Du pr�ver at registrere timer i en periode der er afsluttet/lukket."

'&" Seneste faktura dato p� dette job er: <b>" & lastFakdato &"</b><br><br>"_
'&" Den valgte registrerings dato er: <b>"& regdato & "</b><br><br>"_

case 118
varErrorText = "<b>Kopier job</b><br>"_
&"Der er ikke valgt nogen kontakter."

case 119
varErrorText = "<b>Kopier job</b><br>"_
&"Der er ikke valgt noget job."


case 120
varErrorText = replace(err_txt_120, "#nylinie#", "<br>")

case 121
varErrorText = "<b>ServiceDesk</b><br>"_
&" - Dato for udf�rsel er ikke angivet i et korrekt format."


case 122
varErrorText = "<b>Aktivitet</b><br>"_
&" - Antal stk. er ikke angivet i et korrekt format."

case 123
varErrorText = "<b>Webblik - Joblisten</b><br>"_
&" - <b>Prioitet</b> er ikke angivet i et korrekt format."

case 124
varErrorText = "<b>Job - budget</b><br><br>"_
&"Gns. timepris er ikke angivet korrekt."

case 125
varErrorText = "<b>Job - budget</b><br><br>"_
&"Medarbejder bel�b beregnet udfra Gns. timepris * faktor er ikke angivet korrekt."


case 1251
varErrorText = "<b>Job - budget</b><br><br>"_
&"Interne omkostninger til l�n er ikke angivet korrekt."

case 1252
varErrorText = "<b>Job - budget</b><br><br>"_
&"Udgifter til underleverand�rer / salgsomkostinger er ikke angivet korrekt."

case 126
varErrorText = "<b>Job - budget</b><br><br>"_
&"Bruttofortjeneste er ikke angivet korrekt."

case 127
varErrorText = "<b>Job - budget</b><br><br>"_
&"DB % er ikke angivet korrekt."

case 128
varErrorText = "<b>Job - budget</b><br><br>"_
&"<b>Faktor</b> til beregning af medarbejder bel�b er ikke angivet korrekt.<br><br>Tjek at der er angivet mindst 1 time p� en medarbejdertype linje."

case 129
varErrorText = "<b>Job - budget</b><br><br>"_
&"Udgifter p� job er ikke angivet i et korrekt format."

case 130

if tiderRettet = 1 then
trTxt = "Ja"
else
trTxt = "Nej"
end if

varErrorText = "<b>Stempelur</b><br>"_
&"Der mangler at blive angivet en kommentar<br><br>- Hvis logind elller logud �ndres <u>skal</u> det begrundes med en kommentar.<br><br>"_
&"- Der er valgt en stempelur-indstilling der altid <u>skal</u> begrundes med en kommentar. (tilretning)<br><br>"_
&"<b>Data:</b><br>"_
&"Stempelurs indstilling: ("& trim(ids(a)) &") "& stur(a)  & "<br>"_
&"Manuelt afsluttet: " & manuelt_afsluttet & "<br>"_
&"Tid �ndret: "& trTxt & "<br><br>"

if len(trim(oprLogin)) <> 0 then
varErrorText = varErrorText &"Opr. logind: " & formatdatetime(oprLogin, 2) &" "& formatdatetime(oprLogin, 3) & "<br>"_
&"Nyt logind: " & formatdatetime(loginTid,  2) &" "& formatdatetime(loginTid, 3) & "<br><br>"
end if

'varErrorText = varErrorText & logudTid

if len(trim(oprLogud)) <> 0 AND len(trim(logudTid)) <> 0 then
varErrorText = varErrorText &"Opr. logud: " & formatdatetime(oprLogud, 2) &" "& formatdatetime(oprLogud, 3) & "<br>"_
&"Nyt logud: " & formatdatetime(logudTid, 2) &" "& formatdatetime(logudTid, 3) 
end if

varErrorText = varErrorText &"<br><br>&nbsp;"

case 131
varErrorText = "<b>Timeregistrering</b><br>"_
&"Hvis timeregistreringer p� pre-udfyldte aktiviterer �ndres <u>skal</u> det begrundes med en kommentar."

case 132
varErrorText = "<b>Job</b><br>"_
&"Prioitet er ikke angivet i et korrekt format. (heltal)"

case 133
varErrorText = "<b>Komme / G�</b><br>"_
&"Der er allerede oprettet komme/g� registreringer p� den valgte dato.<br>"_
&"Det er ikke tilladt at oprette komme/g� registreringer p� datoer hvor der allerede har v�ret logget ind."


case 134
varErrorText = "<b>Komme / G�</b><br>"_
&"En eller flere komme/g� registreringer er i konflikt med hinanden.<br><br>"_
&"Det er ikke tilladt at oprette flere komme/g� registreringer p� samme dato der d�kker over det samme tidsrum.<br><br>"_
&"Der er konflikt med denne komme/g� registrering:<br> <b>"& strLogindkonflikt & "</b>"

case 136
varErrorText = "<b>Timeregistrering</b><br>"_
&"Der er ikke fundet en aktivitet med en �ben tidsl�s i det valgte tidsrum p� den valgte dato.<br>"_
&"<br>Du har angivet kl.: <b>"& tSttid(y) & " - "& tSltid(y) & "</b><br>"_
&"Tidsl�sen er: "& tidslaas_st &" - "& tidslaas_sl

case 137
varErrorText = "<b>Kontrolpanel</b><br>"_
&"Den angivne <b>licens startdato</b> er ikke angivet i et korrekt dato format.<br>"

case 138
varErrorText = "<b>Kontaktpersoner</b><br>Der mangler at blive indtastet et navn."


case 139
varErrorText = "<b>Stempelur - logindhistorik</b><br><br>Logud tidspunkt er f�r logind tidspunkt.<br><br>"
varErrorText = varErrorText & "Du har indtastet: <b>"& left(formatdatetime(loginTid, 3), 5) &"</b> - <b>"& left(formatdatetime(logudTid, 3), 5) &"</b>"

if use_ig_sltid = 1 then
varErrorText = varErrorText & "<br><br>Logind tidspunkt er korrigeret til: <b>"& left(formatdatetime(loginTid, 3), 5) &"</b> if�lge stempelur indstillingerne. Kontakt en administrastiv medarbejder for yderligere information."
end if


case 140
varErrorText = "<b>Materiale reg.</b><br><br>- Der er ikke valgt job."

case 141

varErrorText = "<b>Aktiviteter</b> <br><br> "_
&"- Fase indeholder en ulovlig karakter (<b>apostrof.</b>).<br>"_
&"- Fase indeholder et <b>mellemrum</b> i navnet. Fasenavne skal v�re et sammenh�ngende ord. f.eks: ""01_Hovedprojekt""</br><br>"

case 142
varErrorText = "<b>Materiale- / Udl�gs -registrering</b><br><br>- Forbrugsdato er ikke angivet i et gyldigt dato format."

case 143
varErrorText = "<b>Materiale- / Udl�gs -registrering</b><br><br>- Du har valgt at oprette p� lager, men der mangler at blive angivet et varenr."

case 144
varErrorText = "<b>Print Job / tilbud</b><br><br>- Der mangler at blive angivet et navn."

case 145
varErrorText = "<b>Print Job / tilbud</b><br><br>- Der mangler at blive angivet en folder."


case 146
varErrorText = "<b>Print Job / tilbud</b><br><br>- Der mangler at blive angivet om du vil gemme som skabelon eller dokument."

case 147
varErrorText = "<b>Materiale- / Udl�gs -registrering</b><br><br>- Der mangler at blive valgt en gruppe."

case 148
varErrorText = "<b>Aktiviteter</b><br><br>En eller flere af de indtastede <b>tidsl�s-angivelser</b> er ikke et gyldigt <b>tidsformat.</b> <br><br>"_
&"Tidsformatet skal v�re tt:mm:ss, f.eks <b>9:30:00, 12:59:00 etc.</b><br>Den angivne tid er: <b>" & errTid & "</b>"

case 149
varErrorText = "<b>Job</b><br><br>- Der mangler at blive angivet et rekvisitions nr."

case 150
varErrorText = "<b>L�nperiode</b><br><br>- Dato er ikke angivet i et gyldigt datoformat, eller datoen eksisterer ikke."

case 151
varErrorText = "<b>Job</b><br><br>- De angivne procenter angivet under jobansvalige indeholder ulovlige kakrakterer."

case 152
varErrorText = "<b>Faktura forfaldsdato</b><br><br>- Den angvine forfaldsdato er ikke angivet i et gyldigt datoformat.<br> Du har angivet: <b> " & Request("FM_forfaldsdato") & "</b>"


case 153
varErrorText = "<b>Aktivitet</b><br><br>- Den angivne sortering er ikke angivet som et hel-tal"
'case 151
'varErrorText = "<b>L�nperiode</b><br><br>- Dato er ikke angivet i et gyldigt datoformat, eller datoen eksisterer ikke."

case 154
varErrorText = "<b>Materiale reg.</b><br><br>- Det angivne antal, <br>- K�bspris eller salgspris<br>er ikke angivet i et gyldigt format. (tal)<br><br>- Der er ikke angivet et navn."

case 155
varErrorText = "<b>Kontantperson hos kunde (kontakt)</b><br><br>- I jeres version af TimeOut skal der v�re angivet en kontaktperson hos kunden n�r man opretter/redigerer et job."

case 156
varErrorText = "<b>Job</b><br>"_
&" - Restestimat p� job er ikke angivet i et korrekt format."

case 157
varErrorText = "<b>Km pris</b><br>"_
&" - Km pris er ikke angivet i et korrekt format."

case 158
varErrorText = "<b>Fejl vedr. registrering p� akt. type (ferie / sygdom etc.)</b><br><br>Du kan ikke flytte indtastning p� denne type til en dato uden norm tid.<br /> F.eks kan du ikke angive ferie p� en s�ndag med 0 timer i norm tid.<br /><br />"

case 159
varErrorText = "<b>Fejl</b><br><br>Den indtastede <b>fradato</b> er ikke et gyldigt datoformat.<br>Du har indtastet: <b>"& fraDato &"</b><br /><br />"

case 160
varErrorText = "<b>Fejl</b><br><br>Den indtastede <b>timepris ell. kostpris</b> er ikke indtastet som et bel�b.<br><br>Du har indtastet:<br>Timepris: "& fasttp &" og Kostpris: "& fastkp &"<br /><br />"

case 161
varErrorText = "<b>Fejl</b><br><br>Der er ikke angivet et <b>navn</b> til ganttlisten.<br /><br />"


case 162
varErrorText = "<b>Faktor</b><br>"_
&" - Faktor er ikke angivet som et <b>tal</b>.<br>"_
&"Du har indtastet f�lgende: <b>"& globalfaktor &"</b><br>"_
&" <br><br>"

case 163
varErrorText = "<b>Fejl</b><br><br>Der er ikke fundet nogen korrektions aktiviteter.<br /><br />"

case 164
varErrorText = "<b>Fejl</b><br><br>Den l�n-periode du fors�ger at lukke er allerede lukket og kan derfor ikke lukkes igen.<br /><br />"

case 165
varErrorText = "<b>Fejl</b><br><br><b>Nettobel�b</b> ("& formatnumber(jo_gnsbelobTjk) &") er st�rre end<br> <b>bruttobel�b</b> ("& formatnumber(strBudgetTjk) &")<br /><br />"

case 166
varErrorText = "<b>Materiale reg.</b><br><br>- Det angivne jobnr er IKKE fundet."

case 167
varErrorText = "Der mangler at blive angivet/valgt en kunde/job."

case 168
varErrorText = "Der mangler at blive angivet/valgt en aktivitet."

case 169
varErrorText = "Der er sket en regnefejl, eller der er manuelt indtastet et forket bel�b i totalbel�b.<br><br>Totalbel�b: <b>"& intBeloebtjk & "</b> stemmer ikke overens<br> med bel�bet p� faktura linjerne: <b>"& tjkSum &"</b><br><br>Tjek seneste �ndrede aktivitet og klik p� 'calc' igen."

case 170
varErrorText = "You need to select a customer!"

case 171
varErrorText = "You need to type a style!"

case 172
varErrorText = "You need to type a order no.<br>Or the selected order no. allready exist on another order."


case 173
varErrorText = "You are missing one or more of the following information: <br><br>"
varErrorText = varErrorText & "<b>Required for active orders:</b><br><br> "
varErrorText = varErrorText & "Customer, Style and <br> "
varErrorText = varErrorText & "- Sales Rep.<br> "
varErrorText = varErrorText & "- Supplier<br> "
varErrorText = varErrorText & "- ETD Buyer<br> "

case 174
varErrorText = "You are missing one or more of the following information: <br><br>"
varErrorText = varErrorText & "<b>Required for shipped orders:</b><br><br> "
varErrorText = varErrorText & "Customer, Style and <br> "
varErrorText = varErrorText & "- Sales Rep.<br> "
varErrorText = varErrorText & "- ETD Buyer<br> "
varErrorText = varErrorText & "- Actual ETD<br> "
varErrorText = varErrorText & "- Order Qty.<br> "
varErrorText = varErrorText & "- Supplier<br> "
varErrorText = varErrorText & "- Shipped Qty.<br> "
varErrorText = varErrorText & "- Supplier invoiceno<br> "    


case 175
varErrorText = "Dato er ikke angivet i et gyldigt datoformat"

case 176
varErrorText = "Timer er ikke angivet i et gyldig format, eller der er ikke tastet timer." '0 timer

case 177
varErrorText = "Forretningsomr�de er obligatorisk, og der er ikke valgt noget forretningsomr�de"

case 178
varErrorText = "Der kan ikke indtastes di�ter / rejsedage p� dage hvos der er registreret mere end "& traveldietexp_maxhours &" timer p� dagen."

case else
varErrorText = errortype
end select




'************** MOBIL version ****'
 call browsertype()

'response.write "browstype_client: "& browstype_client
if browstype_client = "ip" then 'OR browstype_client = "ie"%>
<title>TimeOut mobile</title>

<!--
<link href='http://fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css'>
<link href="../timetag_web/css/style.less" rel="stylesheet/less" type="text/css" />
<script src="../timetag_web/js/less.js" type="text/javascript"></script>
    -->


      <div id="dverror_msg" style="position:absolute; top:0px; left:0px; height:100%; width:100%; border:0px; font-size:17px; font-size:1.7rem; padding:50px; line-height:19px; line-height:1.9rem; background-color:#ecf0f1;">
         
        <span style="font-size:20px; font-size:2.0rem;"><b>TimeOut mobile</b></span><br /><br />
     
           <%=err_txt_001%><br /><br />
    <%=varErrorText%>
	<br /><br />

          	<%if errortype <= 4 then%>
		<a href="javascript:history.back()" style="font-size:17px; font-size:1.7rem;"><< Tilbage</a>
		<%
		else
		if errortype = 5 then
		%>
		<a href="../login.asp" target="_top" style="font-size:17px; font-size:1.7rem;"><< Tilbage</a>
		<%
		else
			if errortype = 25 OR errortype = 88 then
			Response.write ".."
			else 
			%>
			<a href="javascript:history.back()" style="font-size:17px; font-size:1.7rem;"><< Tilbage</a>
			<%end if%>
		<%end if%>
		<%end if%>
<br /><br />

	Err code: 000<%=errortype %>
    </div>
<%



else

if useleftdiv <> "s" then

select case useleftdiv
case "to_2015"
leftdiv = 300
topdiv = -100
widthdiv = 400
heightdiv = 200 
case "j2"
leftdiv = 100
topdiv = 50
widthdiv = 300
heightdiv = 200 
case "j"
leftdiv = 10
topdiv = 50
widthdiv = 360
heightdiv = 200
case "c"
leftdiv = 150
topdiv = 150
widthdiv = 460
heightdiv = 200
case "c2"
leftdiv = 250
topdiv = 150
widthdiv = 360
heightdiv = 200
case "m"
leftdiv = 100
topdiv = 50
widthdiv = 400
heightdiv = 200
case "t"
leftdiv = 250 '150
topdiv = 250 '20
widthdiv = 400
heightdiv = 200
case "f"
leftdiv = 10
topdiv = 20
widthdiv = 270
heightdiv = 100
case "tt"
leftdiv = 250
topdiv = 20
widthdiv = 360
heightdiv = 100
case "h2"
leftdiv = 0
topdiv = 360
widthdiv = 232
heightdiv = 250

case else
leftdiv = 250
topdiv = 120
widthdiv = 400
heightdiv = 300
end select



%>
<div id="error" style="position:absolute; left:<%=leftdiv%>px; top:<%=topdiv%>px; width:<%=widthdiv%>px; background-color:#ffffff; padding:20px; border:10px #cccccc solid;">
<table cellspacing="0" cellpadding="10" border="0" width=100% style="background-color:#ffffff;">
	<tr>
	    <td align="left" style="padding-top:20px; float:left; border:0px;"><h4><%=err_txt_001%></h4></td>
	</tr>
	<tr>
	<td align="left" style="float:left; padding:20px; border:0px; width:360px; background-color:#ffffff;">
	<div style="padding-left:0px; text-align:left; background-color:#ffffff;""><%=varErrorText%>
	<br /><br />
	Err code: 000<%=errortype %>

	
        </div>

        <br /><br />
        &nbsp;
	</td>
	</tr>
	<tr>
	<td style="float:left; border:0px;">
		
		<%if errortype <= 4 then%>
		<a href="javascript:history.back()"><< Tilbage</a>
		<%
		else
		if errortype = 5 then
		%>
		<a href="../login.asp" target="_top"><< Tilbage</a>
		<%
		else
			if errortype = 25 OR errortype = 88 then
			Response.write ".."
			else 
			%>
			<a href="javascript:history.back()"><< Tilbage</a>
			<%end if%>
		<%end if%>
		<%end if%>
		</td>
	</tr>
</table>
</div>
<%
else 'useleftdiv = S (SIMPEL / jquey)

    %>
    <%=varErrorText%>
	<br />Err code: 000<%=errortype %>

    <%
end if 's

end if 'mobile



end function 
%>

