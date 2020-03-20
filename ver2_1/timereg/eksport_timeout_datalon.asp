<html>

<body>

<%



'Definer variabler

dim datalon_fs, datalon_fileIndput, datalon_fileOutput, datalon_IndputFileName, datalon_OutputFileName, datalon_datoFieldNo



'Definer indput og output filer

datalon_IndputFileName = "../inc/log/data/"& file  '"medarbafexp_2011_10_26__13_4_16_fk.csv"

datalon_OutputFileName = Replace(datalon_IndputFileName,"medarbafexp","datalon_dencker") 'HLW
'datalon_OutputFileName = Replace(datalon_IndputFileName,".csv",".txt") 'HLW

'"&filnavnDato&"_"&filnavnKlok&"_"&lto&"

'To cifres kode for kommune

strInstkode = "FB" 'FB_MFKA



'Tjenstekode værdier 0, 1, 2, 3

strTjenstekode = "0"



'Åben .csv fil til indlæsning

set datalon_fs=Server.CreateObject("Scripting.FileSystemObject")

set datalon_fileIndput=datalon_fs.OpenTextFile(Server.MapPath(datalon_IndputFileName),1,true)

set datalon_fileOutput=datalon_fs.OpenTextFile(Server.MapPath(datalon_OutputFileName),8,true)




'Læs første linie da det er header
'headerLine = "Virksomhedsnr;Løntermin;Medarbejdernr;Gruppe;Feltnr.;Timer/Beløb;"
'datalon_fileOutput.WriteLine(headerLine)

'Gennemløb hver linje i .csv filen
datalon_fileIndput.ReadLine


Do Until datalon_fileIndput.AtEndOfStream

    sSeg = Split(datalon_fileIndput.ReadLine, ";" )

    
    personId = trim(sSeg(1))

    strSQLm = "SELECT mid FROM medarbejdere WHERE mnr = '"& personId &"'"
    oRec10.open strSQLm, oConn, 3
    if not oRec10.EOF then

        thisMid = oRec10("mid")

    end if
    oRec10.close

    'call meStamdata(thisMid)
    call medariprogrpFn(thisMid)

    if personId = "20205" then
    personId = "205"
    else
    personId = replace(personId, "20", "0")
    end if
   

     '* Timelønnet eller Funktionær
    loentermin = 0
    if (instr(medariprogrpTxt, "#40#")) <> 0 then
    loentermin = 51
    end if

    if (instr(medariprogrpTxt, "#41#")) <> 0 then
    loentermin = 24
    end if

    

    'avdelingsno = meType
    avdelingsno = "- Ingen gruppe -"
   
    'projektno = "000000000000"
    'e5 = "000000000000"


    daydd = day(now)
    if daydd < 10 then
    daydd = "0"& daydd
    end if

    mthdd = month(now)
    if mthdd  < 10 then
    mthdd  = "0"& mthdd 
    end if

    dato = daydd &""& mthdd &""& right(year(now), 2)  'formatdatetime(now, 2)
    'dato = replace(dato, " ", "")
    'dato = replace(dato, "-", "")
    'dato = replace(dato, "/", "")



    for s = 0 TO UBOUND(sSeg)

    wrtLine = 0
    sats = 0

    lonart = "00000"
   
    e2 = "000000000000"
    e3 = "000000000001"
    e4 = "000000000000"


    'case 7 'løntimer
    'case 10 'Fædre orlov
    'case 12 'kørselstillæg
    'case 13 'Befordring efter skat 
    'case 14 'Personalekasse 
    'case 15 'Forældre orlov
    'case 16 'Tilsyn til maskiner
    'case 17 '6 ferieuge brugt
    'case 22 'Madordning + Rundstrykekr frdrag
    'case 30 'Ferie afholdt dage
    'case 39 'Afspadsering
    'case 43 'Syg timer
    'case 48 'Barn syg timer
    'case 52 'Barsel dage
    'case 54 'Køb via Dencker



    select case s
    case 7 'løntimer incl. afspadsering og ferie ellers 5 'løntimer
    wrtLine = 1
    lonart = "0001"
    'e1 = "000000005000"
    sats = 300
    case 10 'Fædre orlov
    wrtLine = 1
    lonart = "0062"
    'e1 = "000000005000"
    sats = 300
    case 12 'Kørsel
    wrtLine = 1
    lonart = "2425"
    'e1 = "000000005000"
    sats = 300
    case 13 'Befordring efter skat
    wrtLine = 1
    lonart = "0010"
    'e1 = "000000005000"
    sats = 300
    case 14 'Personalekasse
    wrtLine = 1
    lonart = "0008"
    'e1 = "000000005000"
    sats = 300
    case 15 'Forældre orlov
    wrtLine = 1
    lonart = "0014"
    'e1 = "000000005000"
    sats = 300
    case 16 'Tilsyn til maskiner
    wrtLine = 0
    lonart = "0003"
    'e1 = "000000000000"
    sats = 300
    case 17 'FROKOST: Madordning og rundstykker
    wrtLine = 1
    lonart = "0094"
    'e1 = "000000000000"
    sats = 300
    case 22 '6 ferieuge brugt
    wrtLine = 1
    lonart = "7778"
    'e1 = "000000000000"
    sats = 300
    case 30 'Ferie afholdt
    wrtLine = 1
    lonart = "0047"
    'e1 = "000000005000"
    sats = 300
    case 39 'Afspadsering
    wrtLine = 0
    lonart = "0000"
    'e1 = "000000005000"
    sats = 300
    case 43 'Syg timer
    wrtLine = 1
    lonart = "0072"
    'e1 = "000000000000"
    sats = 300
    case 48 'Barn Syg timer
    wrtLine = 1
    lonart = "5C6C"
    'e1 = "000000000000"
    sats = 300
    case 52 'Barsel dage
    wrtLine = 1
    lonart = "6263"
    'e1 = "000000000000"
    sats = 300
    case 54 'Køb via Dencker
    wrtLine = 1
    lonart = "0093"
    'e1 = "000000000000"
    sats = 300
    end select

    e1 = "000000000000"

    'wrtLine = 0

    'if len(trim(sSeg(s))) <> 0 then
    'belob = sSeg(s)*sats '"1234567890"
    'else
    belob = 0
    'end if
    
   
   
    '*** tillæg til komme/gå. Afspadsering + ferie
    if cint(s) = 7 then
        if cint(loentermin) = 24 then
        antal = trim(sSeg(7)) '-(sSeg(37))) 'minus afspadsering
        else
        antal = "16033"
        end if
    else
    antal = trim(sSeg(s))
    end if

    'tjekker om linjen er tom
    if antal <> "0" then
    antalLinjeTjk = antal
    else
    antalLinjeTjk = 0
    end if

    '*** Skal komma fjernes?
    removekomma = 1
    antal_left = 0
    antalNUL = 0
    if cint(removekomma) = 1 then
        antal_left = instr(antal, ",")

        if antal_left <> 0 then
        antal = replace(antal, ".", "")
        antal = replace(antal, ",", "")
        antalNUL = 0
        else

            'Response.write "<br>antal " & antal
            'Response.flush

            if antal <> "0" AND len(trim(antal)) <> 0 then
            antal = antal & "00"
            antalNUL = 0
            else
            antalNUL = 1
            end if

        end if
    end if

    '*** Tilføj komma
    tilfojKomma = 0

    if cint(tilfojKomma) = 1 then

     'select case s
     'case 23

      'antal = antal & ",00"

        antal_left = instr(antal, ",")

        if cint(antal_left) = 0 then
        antal = antal & ",00"
        end if

      'end select

    end if

  

    '***
    '*** Skal der afrundes til nærmeste 15 min.
    afrundestil15min = 0

    if cint(afrundestil15min) = 1 then
        if s = 200 then 'ordinære - lunch

            'antal_timer = left(antal, antal_left-1)
            'if right(antal, 2) <> 00 then
            'antal = antal_timer&"50"
            'end if

        else

            if antal_left <> 0 then
            antal_timer = left(antal, antal_left-1)
  
                'afrund til nærnmeste 15 min
                if right(antal, 2) < 13 then
                antal = antal_timer&"00"
                end if

                if right(antal, 2) >= 13 AND right(antal, 2) < 35 then
                antal = antal_timer&"25"
                end if
   
                if right(antal, 2) >= 35 AND right(antal, 2) < 60 then
                antal = antal_timer&"50"
                end if

                if right(antal, 2) >= 60 AND right(antal, 2) < 85 then
                antal = antal_timer&"75"
                end if

                if right(antal, 2) >= 85 then
                antal_timer = antal_timer + 1
                antal = antal_timer&"00"
                end if

            end if

        end if
    end if

 
    lpaduse = 0

    if cint(lpaduse) = 1 then 

    antal_len = len(antal)
    n = 0
    LpadNuller = ""
    for n = 0 TO (9 - antal_len)

        LpadNuller = LpadNuller & "0" 
        
    next

    antal = LpadNuller & antal

    end if
   

    sats = replace(sats, ",", "")
    sats = replace(sats, ".", "")
    sats_len = len(sats)
    n = 0
    LpadNuller = ""
    for n = 0 TO (10 - sats_len)

        LpadNuller = LpadNuller & "0" 
        
    next

    sats = LpadNuller & sats 


    belob = replace(belob, ".", "")
    belob = replace(belob, ",", "")

    belob_len = len(belob)
    n = 0
    LpadNuller = ""
    for n = 0 TO (11 - belob_len)

        LpadNuller = LpadNuller & "0" 
        
    next

   

    belob = LpadNuller & belob

    n = 0
    filler = ""
    for n = 0 TO 29

        filler = filler & " " 
        
    next
    filler = filler

    'CreateHLTLine(csvLine, Tjenstekode, Afdeling, Konto, loenArt1, loenArt2, enheder)

    'csvLine(7) Ferie, afholdelse  :Ferie 730+130 (skal være timer istedet for dage)

    'if len(trim(sSeg(1))) <> 0 AND sSeg(1) > 0 then
        
    
    
        if cint(wrtLine) = 1 AND antalLinjeTjk <> "0" ANd len(trim(antal)) <> 0 then

            sdLine = "10180937;"& loentermin &";"& personId &";"& avdelingsno &";"& lonart &";"& antal &";"
            datalon_fileOutput.WriteLine(sdLine)

           if cint(s) = 7 AND cint(loentermin) = 24 then 'Også normal timer for Termin 24

                 'sdLine = "10180937;"& loentermin &";"& personId &";"& avdelingsno &";1213;"& antal &";"
                 'datalon_fileOutput.WriteLine(sdLine)

                 sdLine = "10180937;"& loentermin &";"& personId &";"& avdelingsno &";0012;"& antal &";"
                 datalon_fileOutput.WriteLine(sdLine)

                 'sdLine = "10180937;"& loentermin &";"& personId &";"& avdelingsno &";0013;17200;"
                 'datalon_fileOutput.WriteLine(sdLine)

           end if

               'if cint(s) = 7 AND cint(loentermin) = 24 then 'Også normal timer for Termin 24
               'If ferie dag 
               '14-antal feriedage
               'Hvis 5 dage skal skattedage 10-
                'sdLine = "10180937;"& loentermin &";"& personId &";"& avdelingsno &";0011;17200;"
                'datalon_fileOutput.WriteLine(sdLine)

               'Frokost rettes til 32
            

        end if

    'end if

    next 's
 


Loop



datalon_fileIndput.Close

datalon_fileOutput.Close



set datalon_fileIndput=Nothing

set datalon_fileOutput=Nothing



set datalon_fs=Nothing



Response.Write "Script er udført"






%>

</body>

</html>