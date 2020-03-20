<html>

<body>

<%



'Definer variabler

dim datalon_plan_fs, datalon_plan_fileIndput, datalon_plan_fileOutput, datalon_plan_IndputFileName, datalon_plan_OutputFileName, datalon_plan_datoFieldNo



'Definer indput og output filer

datalon_plan_IndputFileName = "../inc/log/data/"& file  '"medarbafexp_2011_10_26__13_4_16_fk.csv"

datalon_plan_OutputFileName = Replace(datalon_plan_IndputFileName,"medarbafexp","datalon_plan") 'HLW
'datalon_plan_OutputFileName = Replace(datalon_plan_IndputFileName,".csv",".txt") 'HLW

'"&filnavnDato&"_"&filnavnKlok&"_"&lto&"

'To cifres kode for kommune

strInstkode = "FB" 'FB_MFKA



'Tjenstekode værdier 0, 1, 2, 3

strTjenstekode = "0"



'Åben .csv fil til indlæsning

set datalon_plan_fs=Server.CreateObject("Scripting.FileSystemObject")

set datalon_plan_fileIndput=datalon_plan_fs.OpenTextFile(Server.MapPath(datalon_plan_IndputFileName),1,true)

set datalon_plan_fileOutput=datalon_plan_fs.OpenTextFile(Server.MapPath(datalon_plan_OutputFileName),8,true)




'Læs første linie da det er header
'headerLine = "Virksomhedsnr;Løntermin;Medarbejdernr;Gruppe;Feltnr.;Timer/Beløb;"
'datalon_plan_fileOutput.WriteLine(headerLine)

'Gennemløb hver linje i .csv filen
datalon_plan_fileIndput.ReadLine


Do Until datalon_plan_fileIndput.AtEndOfStream

    sSeg = Split(datalon_plan_fileIndput.ReadLine, ";" )

    medarbLonnummer = 0
    arbejdsgiverId = "10015359"
    personId = trim(sSeg(1))
    mgruppe = 0
    medarbSalary = 0
    sundhedsforsikring = 0
    personaleforening = 0
    bruttolontraek = 0
    multimedieskat = 0

    strSQLm = "SELECT m.mid, m.medarbejdertype, mt.mgruppe, medarbejder_rfid, salary, personaleforening, m1_longruppe, bruttolontraek, multimedieskat FROM medarbejdere m "_
    &" LEFT JOIN medarbejdertyper mt ON (mt.id = m.medarbejdertype) "_
    &" WHERE mnr = '"& personId &"'"
    oRec10.open strSQLm, oConn, 3
    if not oRec10.EOF then

        thisMid = oRec10("mid")
        'mgruppe = oRec10("mgruppe")
        medarbLonnummer = oRec10("medarbejder_rfid")
        medarbSalary = oRec10("salary")
        personaleforening = oRec10("personaleforening")
        sundhedsforsikring = oRec10("m1_longruppe")
        bruttolontraek = oRec10("bruttolontraek")
        multimedieskat = oRec10("multimedieskat")

    end if
    oRec10.close

    medarbLonnummer = replace(medarbLonnummer, "51-", "")


    if medarbLonnummer <> "0" AND len(trim(medarbLonnummer)) <> 0 then

    'call meStamdata(thisMid)
    call medariprogrpFn(thisMid)

    '*** Frokostordning Ja / nej
    frokostordning = 0
    if (instr(medariprogrpTxt, "#31#")) <> 0 then
    frokostordning = 1
    end if


    '* Timelønnet eller Funktionær
    mgruppe = 1 'Funk
    if (instr(medariprogrpTxt, "#16#")) <> 0 then 'Timelønnet
    mgruppe = 2
    end if

     
    'if mgruppe = 1 then 'Månedsløn
    loentermin = 51
    'end if

    'if mgruppe = 2 then 'Timeløn / 14 dagesløn
    'loentermin = 24
    'end if

    

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

    
    'Fra Excel fil kolonner
    'case 1  'Medarbejder NR
    'case 3  'Kørselstillæg
    'case 7  'løntimer - Hvis timelønnet
    'case 30 'Ferie afholdt dage
    'case 31  'Rejsegodtgørelse
    'case 54 'Aconto
    
    'Timelønnede:
    'case 2 'Ikke fakturerbare
    'case 17 '6 ferieuge brugt
    'case 30 'Ferie afholdt dage
    'case 39 'Tjenestefri
    'case 43 'Syg timer
    'case 48 'Barn syg timer
    'case 52 'Barsel dage

    
    
    '0001 timer -- OK
    '0002 Løn -- OK
    '2425 Km 3,52 * KM -- OK
    '0005 Pension -- OK
    '0008 Personaleforening og Frokost -- OK
    '0009 Aconto udb.
    '0055 Ferie tillg -- IKKE med
    '007L Sundhedsforsikring -- OK
    '0047 Afholdt ferie
    '0007 Rejsegodtgørelse
    '0042 Multimedie -- OK
    '0098 Bruttoløntræk --OK
  


    select case s
    case 0 'Salary
    wrtLine = 1
    lonart = "0002"
    'e1 = "000000005000"
    sats = 1

    case 3 'timer
        '** Kun timelønende 
        if cint(mgruppe) = 2 then
        wrtLine = 1
        else
        wrtLine = 0
        end if
    lonart = "0001"
    'e1 = "000000005000"
    sats = 10
    
    case 5 'Kørsel
    wrtLine = 1
    lonart = "2425"
    'e1 = "000000005000"
    sats = 10
    
    case 6 'FROKOST: Kun timelønnede ellers fastbeløb
    
        '** Kun timelønende 
        if cint(mgruppe) = 2 then
        wrtLine = 1
        else
        wrtLine = 0
        end if

    lonart = "0094"
    'e1 = "000000000000"
    sats = 10
   
    case 15 'Ferie afholdt
    wrtLine = 1
    lonart = "0047"
    'e1 = "000000005000"
    sats = 10

    case 16 'Ferie afholdt u løn.
        'if cdbl(sSeg(14)) = 0 then
        wrtLine = 1
        'else
        'rtLine = 0
        'end if
    lonart = "0047"
    'e1 = "000000005000"
    sats = 10

    case 34 'Aconto
    wrtLine = 1
    lonart = "0093"
    'e1 = "000000000000"
    sats = 10

    case 35 'Rejsegodtgørelse / diæter
    wrtLine = 1
    lonart = "0007"
    'e1 = "000000005000"
    sats = 10

    case 9 '6 ferieuge brugt / Feriefri
        '** Kun timelønende 
        if cint(mgruppe) = 2 then
        wrtLine = 0
        else
        wrtLine = 0
        end if
    lonart = "7778"
    'e1 = "000000000000"
    sats = 10

    case 22 'Syg timer
        '** Kun timelønende 
        if cint(mgruppe) = 2 then
        wrtLine = 0
        else
        wrtLine = 0
        end if
    lonart = "0072"
    'e1 = "000000000000"
    sats = 10

    case 27 'Barn Syg timer
        '** Kun timelønende 
        if cint(mgruppe) = 2 then
        wrtLine = 0
        else
        wrtLine = 0
        end if
    lonart = "5C6C"
    'e1 = "000000000000"
    sats = 10

    case 31 'Barsel dage
        '** Kun timelønende 
        if cint(mgruppe) = 2 then
        wrtLine = 0
        else
        wrtLine = 0
        end if
    lonart = "6263"
    'e1 = "000000000000"
    sats = 10

    case 33 'Tjenenstefri
        '** Kun timelønende 
        if cint(mgruppe) = 2 then
        wrtLine = 0
        else
        wrtLine = 0
        end if
    lonart = "0000"
    'e1 = "000000000000"
    sats = 10
  

    end select

    e1 = "000000000000"

    'wrtLine = 0

    'if len(trim(sSeg(s))) <> 0 then
    'belob = sSeg(s)*sats '"1234567890"
    'else
    belob = 0
    'end if
    
   
   
    '*** Antal på de variable
    select case cint(s) 
    case 0 
        if cint(mgruppe) = 2 then
        medarbSalaryUseFOrBeregn = replace(medarbSalary, ".", "")
        medarbSalaryUseFOrBeregn = replace(medarbSalaryUseFOrBeregn, ",", ".")
        antal = (trim(sSeg(3))*1 + trim(sSeg(4))*1 + trim(sSeg(22))*1 + trim(sSeg(27))*1 ) * (medarbSalary*1) * 1
        else
        antal = medarbSalary
        end if
    case 3 'Timer timelønnede. Fakturerbare, ikke fakturerbare + syg timer + bar nsyg timer
    antal = (trim(sSeg(3))*1 + trim(sSeg(4))*1 + trim(sSeg(22))*1 + trim(sSeg(27))*1 )  * 1
    case 5 'KM
    antal = trim(sSeg(s)) * 3.52
    'case 14, 17 'Ferie + Ferie afh. u løn
    'antal = (trim(sSeg(14)) * 1) + (trim(sSeg(17)) * 1)
    case else
    antal = trim(sSeg(s))
    end select

    'antal = trim(sSeg(s))

    'tjekker om linjen er tom
    if antal <> "0" then
    antalLinjeTjk = antal
    else
    antalLinjeTjk = 0
    end if

    '*** Skal komma fjernes?
    removekomma = 1
    antal_left = 0
    antal_len = 0
    antalNUL = 0

    if cint(removekomma) = 1 then
        antal_left = instr(antal, ",")


        if antal_left <> 0 then
        
        antal_len = (len(antal) - antal_left)
        
        antal = replace(antal, ".", "")
        antal = replace(antal, ",", "")
    
        if antal_len = 1 then ' et tal efter komma DER SKAL ALTID VÆRE 2
        antal = antal & "0"
        else
        antal = antal
        end if

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

    
        antal_left = instr(antal, ",")

        if cint(antal_left) = 0 then
        antal = antal & ",00"
        end if

  

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

  
    
    
        if (cint(wrtLine) = 1 AND antalLinjeTjk <> "0" ANd len(trim(antal)) <> 0) OR (cint(s) = 0 AND cint(mgruppe) = 2) then

            sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";"& lonart &";"& antal 
            datalon_plan_fileOutput.WriteLine(sdLine)

           if cint(s) = 0 AND cint(mgruppe) <> 2 then 'Timer Funktionærer

                 sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";0001;16033"
                 datalon_plan_fileOutput.WriteLine(sdLine)

           end if

          'if cint(s) = 0 then 'Pension

           '      Pensionbel = medarbSalary * 0.10
            
           '      sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";0005;"& Pensionbel &";"
           '      datalon_plan_fileOutput.WriteLine(sdLine)

           'end if

           if cint(s) = 0 then '0008 Personaleforening og Frokost

                personaleforening_frokostordning = 0
                if cint(personaleforening) = 1 then
                personaleforeningbel = 45.00
                personaleforening_frokostordning = personaleforeningbel
                end if

                if cint(frokostordning) = 1 then
                frokostordningbel = 450.00
                personaleforening_frokostordning = (personaleforening_frokostordning*1) + (frokostordningbel*1)
                end if

                personaleforening_frokostordning = personaleforening_frokostordning * 100

                personaleforening_frokostordning = replace(personaleforening_frokostordning, ".", "")
                personaleforening_frokostordning = replace(personaleforening_frokostordning, ",", "")

                sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";0008;"& personaleforening_frokostordning 
                datalon_plan_fileOutput.WriteLine(sdLine)

           end if

           if cint(s) = 0 AND cint(sundhedsforsikring) = 1 then '007L sundhedsforsikring

                 sundhedsforsikringbel = 7492

                 sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";007L;"& sundhedsforsikringbel 
                 datalon_plan_fileOutput.WriteLine(sdLine)

           end if

           if cint(s) = 0 AND cdbl(bruttolontraek) <> 0 then '007L bruttolontraek

                 'bruttolontraek = replace(bruttolontraek, ".", "")
                 bruttolontraek = replace(bruttolontraek, ",", "")

                 if len(bruttolontraek) = 3 then ' < 1000
                 bruttolontraek = bruttolontraek & "00"
                 else
                 bruttolontraek = bruttolontraek & "0"
                 end if


                 sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";0098;"& bruttolontraek 
                 datalon_plan_fileOutput.WriteLine(sdLine)

           end if

           if cint(s) = 0 AND cint(multimedieskat) = 1 then '0042 Multimedie

                 multimedieskatbel = 24167

                 sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";0042;"& multimedieskatbel 
                 datalon_plan_fileOutput.WriteLine(sdLine)

           end if
       
     
       

        end if

    'end if

    next 's
 

    end if 'do until medarbLonnummer

Loop



datalon_plan_fileIndput.Close

datalon_plan_fileOutput.Close



set datalon_plan_fileIndput=Nothing

set datalon_plan_fileOutput=Nothing



set datalon_plan_fs=Nothing



Response.Write "Script er udført"






%>

</body>

</html>