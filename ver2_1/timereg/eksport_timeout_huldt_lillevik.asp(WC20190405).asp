<html>

<body>

<%



'Definer variabler

dim hlt_fs, hlt_fileIndput, hlt_fileOutput, hlt_IndputFileName, hlt_OutputFileName, hlt_datoFieldNo



'Definer indput og output filer

hlt_IndputFileName = "../inc/log/data/"& file  '"medarbafexp_2011_10_26__13_4_16_fk.csv"

hlt_OutputFileName = Replace(hlt_IndputFileName,".csv","ITxxxxTRS.HLW") 'HLW

'"&filnavnDato&"_"&filnavnKlok&"_"&lto&"

'To cifres kode for kommune

strInstkode = "FB" 'FB_MFKA



'Tjenstekode værdier 0, 1, 2, 3

strTjenstekode = "0"



'Åben .csv fil til indlæsning

set hlt_fs=Server.CreateObject("Scripting.FileSystemObject")

set hlt_fileIndput=hlt_fs.OpenTextFile(Server.MapPath(hlt_IndputFileName),1,true)

set hlt_fileOutput=hlt_fs.OpenTextFile(Server.MapPath(hlt_OutputFileName),8,true)




'Læs første linie da det er header

hlt_fileIndput.ReadLine



'Gennemløb hver linje i .csv filen

Do Until hlt_fileIndput.AtEndOfStream

    sSeg = Split(hlt_fileIndput.ReadLine, ";" )

    personId = trim(sSeg(1))
    personIdlen = len(trim(personId))


    call meStamdata(personId)


    n = 0
    LpadNuller = ""
    for n = 0 TO (5 - personIdlen)

        LpadNuller = LpadNuller & "0" 
        
    next

    
    personId = LpadNuller & personId
   

    n = 0
    len_metype = len(meType)
    LpadNuller = ""
    for n = 0 TO (11 - len_metype)

        LpadNuller = LpadNuller & "0" 
        
    next
    avdelingsno = LpadNuller & meType
    
   
    projektno = "000000000000"
    e5 = "000000000000"


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


    'case 3 'Utetillæg innland
    'case 4 '025 Overtid 50%
    'case 5 '026 Overtid 100%
    'case 6 '010 Reisetid ordinær
    'case 7 '026 Reisetid 100% 
    'case 8 '025 Reisetid 50%
    'case 9 'Utetillæg utland
    'case 10 'Efter midnat

    '*** 
    'case 11 '025 Overtid 50% fastlønn N
    'case 12 '026 Overtid 100% fastlønn N

    'case 13 '400 Timebank

    'case 18 'Avspad m løn
    'case 19 '048 Helligdage N

     '*** + 1 efter 20190218

    'case 20 022 Sitetillegg Onshore Utland timer

    'case 23 '401 Avspadsering egenbetalt 
    'case 24 '402 Avspadsering udbetalt N

    'case 26 '010 Ordinære timer

    'case 32 '013 Syk timer i per. N
    'case 37 '013 Barn Syk timer i per. N

    'case 41 '018 Permisjon N (41)

    'case 43  021 Sitetillegg Onshore Innland
    'case 44  023 Sitetillegg Offshore

    'case 45 'Teamleder tillæg



    select case s
    case 3 'Utetillæg innland
     wrtLine = 1
    lonart = "00020"
    'e1 = "000000005000"
    sats = 300
    case 4, 11 'Overtid 50%
     wrtLine = 1
    lonart = "00025"
    'e1 = "000000005000"
    sats = 300
    case 5, 12 'Overtid 100%
     wrtLine = 1
    lonart = "00026"
     'e1 = "000000005000"
    sats = 300
    case 6 'Reisetid ordinær
     wrtLine = 1
    lonart = "00010"
    'e1 = "000000005000"
    sats = 300
    case 7 'Reisetid 100% 
     wrtLine = 1
    lonart = "00026"
    'e1 = "000000000000"
    sats = 300
    case 8 'Reisetid 50%
     wrtLine = 1
    lonart = "00025"
    'e1 = "000000000000"
    sats = 300
    case 9 'Utetillæg utland
     wrtLine = 1
    lonart = "00021"
    'e1 = "000000005000"
    sats = 300
    case 10 'Efter midnat
     wrtLine = 1
    lonart = "00026"
    'e1 = "000000005000"
    sats = 300
  

    case 13 '11 'Timebank
     wrtLine = 1
    lonart = "00400"
    'e1 = "000000000000"
    sats = 300
    case 18 '15 'Avspad m løn
     wrtLine = 1
    lonart = "00018"
    'e1 = "000000005000"
    sats = 300

    case 19 'Helligdage N
    wrtLine = 1
    lonart = "00048"
    'e1 = "000000005000"
    sats = 300

  

    case 23 '20 'Avspadsering egenbetalt 
    wrtLine = 1
    lonart = "00401"
    'e1 = "000000000000"
    sats = 300

    
    case 24 '402 Avspadsering udbetalt N
    wrtLine = 1
    lonart = "00402"
    'e1 = "000000000000"
    sats = 300

    case 26 '20 '18 'Ordinære timer
     wrtLine = 1
    lonart = "00010"
    'e1 = "000000005000"
    sats = 300

    case 32, 37 '013 Syk timer i per. N
    wrtLine = 1
    lonart = "00013"
    'e1 = "000000000000"
    sats = 300

    case 41 '018 Permisjon N
    wrtLine = 1
    lonart = "00018"
    'e1 = "000000000000"
    sats = 300

  
    case 43 '021 Sitetillegg Onshore Innland
    wrtLine = 1
    lonart = "00021"
    'e1 = "000000000000"
    sats = 300

    case 44 '023 Sitetillegg Offshore
    wrtLine = 1
    lonart = "00023"
    'e1 = "000000000000"
    sats = 300

    case 45 '22 'Teamleder tillæg
    wrtLine = 1
    lonart = "00022"
    'e1 = "000000005000"
    sats = 300
    case else
    wrtLine = 0
    sats = 300
    'lonart = "00000"
    lonart = "99999"
    'e1 = "000000000000"
    end select

    e1 = "000000000000"

    'wrtLine = 1

    'if len(trim(sSeg(s))) <> 0 then
    'belob = sSeg(s)*sats '"1234567890"
    'else
    belob = 0
    'end if
    
    if cint(s) = 26 then ' 25, 20 AND sSeg(14) <> 0 'ordinære - lunch
    antal = trim(sSeg(s) - (sSeg(14))) '14
    'antal = antal * 100
    else
    antal = trim(sSeg(s))
    end if

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

    'tjekker om linjen er tom
    if antalNUL <> "1" then
    antalLinjeTjk = antal
    else
    antalLinjeTjk = 0
    end if


    antal_len = len(antal)
    n = 0
    LpadNuller = ""
    for n = 0 TO (9 - antal_len)

        LpadNuller = LpadNuller & "0" 
        
    next

    antal = LpadNuller & antal

   

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
        
        'hlt_datoFieldNo = 12
        'sdLine = CreateHLTLine(sSeg, personId, strTjenstekode , "999", "", sSeg(3), "") '7
        
        if cint(wrtLine) = 1 AND antalLinjeTjk <> "0" then

            sdLine = personId & lonart & avdelingsno & projektno & e1 & e2 & e3 & e4 & e5 & dato & antal & sats & belob & filler '& vbcr
            hlt_fileOutput.WriteLine(sdLine)

        end if

    'end if

    next 's
 


Loop



hlt_fileIndput.Close

hlt_fileOutput.Close



set hlt_fileIndput=Nothing

set hlt_fileOutput=Nothing



set hlt_fs=Nothing



Response.Write "Script er udført"






%>

</body>

</html>