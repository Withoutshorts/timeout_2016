<html>

<body>

<%



'Definer variabler

dim datalon_lm_fs, datalon_lm_fileIndput, datalon_lm_fileOutput, datalon_lm_IndputFileName, datalon_lm_OutputFileName, datalon_lm_datoFieldNo



'Definer indput og output filer

datalon_lm_IndputFileName = "../inc/log/data/"& file  '"medarbafexp_2011_10_26__13_4_16_fk.csv"

datalon_lm_OutputFileName = Replace(datalon_lm_IndputFileName,"medarbafexp","datalon_lm") 'HLW
'datalon_lm_OutputFileName = Replace(datalon_lm_IndputFileName,".csv",".txt") 'HLW

'"&filnavnDato&"_"&filnavnKlok&"_"&lto&"

'To cifres kode for kommune

strInstkode = "FB" 'FB_MFKA



'Tjenstekode v�rdier 0, 1, 2, 3

strTjenstekode = "0"



'�ben .csv fil til indl�sning

set datalon_lm_fs=Server.CreateObject("Scripting.FileSystemObject")

set datalon_lm_fileIndput=datalon_lm_fs.OpenTextFile(Server.MapPath(datalon_lm_IndputFileName),1,true)

set datalon_lm_fileOutput=datalon_lm_fs.OpenTextFile(Server.MapPath(datalon_lm_OutputFileName),8,true)




'L�s f�rste linie da det er header
'headerLine = "Virksomhedsnr;L�ntermin;Medarbejdernr;Gruppe;Feltnr.;Timer/Bel�b;"
'datalon_lm_fileOutput.WriteLine(headerLine)

'Genneml�b hver linje i .csv filen
datalon_lm_fileIndput.ReadLine


Do Until datalon_lm_fileIndput.AtEndOfStream

    sSeg = Split(datalon_lm_fileIndput.ReadLine, ";" )

    medarbLonnummer = 0
    arbejdsgiverId = "23171147"
    personId = trim(sSeg(1))
    mgruppe = 0
    medarbSalary = 0
    sundhedsforsikring = 0
    personaleforening = 0
    bruttolontraek = 0
    multimedieskat = 0
    pension = 0

    strSQLm = "SELECT m.mid, m.medarbejdertype, mt.mgruppe, medarbejder_rfid, mnr, salary, personaleforening, m1_longruppe, bruttolontraek, multimedieskat, pension FROM medarbejdere m "_
    &" LEFT JOIN medarbejdertyper mt ON (mt.id = m.medarbejdertype) "_
    &" WHERE mnr = '"& personId &"'"
    oRec10.open strSQLm, oConn, 3
    if not oRec10.EOF then

        thisMid = oRec10("mid")
        'mgruppe = oRec10("mgruppe")
        medarbLonnummer = oRec10("mnr") 'oRec10("medarbejder_rfid")
        medarbSalary = oRec10("salary")
        personaleforening = oRec10("personaleforening")
        sundhedsforsikring = oRec10("m1_longruppe")
        bruttolontraek = oRec10("bruttolontraek")
        multimedieskat = oRec10("multimedieskat")
        pension = oRec10("pension")

    end if
    oRec10.close


    yearNow = year(now)
    kmtot = 0
    '** KM sast
     strSQLm = "SELECT SUM(timer) As km FROM timer t "_
     &" WHERE tmnr = '"& thisMid &"' AND tfaktim = 5 AND tdato BETWEEN '"& yearNow &"-01-01' AND '"& yearNow &"-01-01' GROUP BY tmnr"
    oRec10.open strSQLm, oConn, 3
    if not oRec10.EOF then

      kmtot = oRec10("km")

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


    '* Timel�nnet eller Funktion�r
    mgruppe = 1 'Funk
    if (instr(medariprogrpTxt, "#16#")) <> 0 then 'Timel�nnet
    mgruppe = 2
    end if

     
    'if mgruppe = 1 then 'M�nedsl�n
    loentermin = 51
    'end if

    'if mgruppe = 2 then 'Timel�n / 14 dagesl�n
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
    'case 3  'K�rselstill�g
    'case 7  'l�ntimer - Hvis timel�nnet
    'case 30 'Ferie afholdt dage
    'case 31  'Rejsegodtg�relse
    'case 54 'Aconto
    
    'Timel�nnede:
    'case 2 'Ikke fakturerbare
    'case 17 '6 ferieuge brugt
    'case 30 'Ferie afholdt dage
    'case 39 'Tjenestefri
    'case 43 'Syg timer
    'case 48 'Barn syg timer
    'case 52 'Barsel dage

    
    
    '0001 timer -- OK
    '0002 L�n -- OK
    '2425 Km 3,52 * KM -- OK
    '0024 Km -- OK
    '0025 Km Sats 3,52 -- OK
    '0005 Pension -- OK
    '0008 Personaleforening og Frokost -- OK
    '0009 Aconto udb.
    '0055 Ferie tillg -- IKKE med
    '007L Sundhedsforsikring -- OK
    '0047 Afholdt ferie
    '0007 Rejsegodtg�relse
    '0042 Multimedie -- OK
    '0098 Bruttol�ntr�k --OK
    '0010 Befordringefter SKAT
  


    select case s
    case 0 'Salary
    wrtLine = 0
    lonart = "0002"
    'e1 = "000000005000"
    sats = 300

    case 2 'timer
        '** Kun timel�nende 
        if cint(mgruppe) = 2 then
        wrtLine = 1
        else
        wrtLine = 0
        end if
    
    lonart = "0001"
    'e1 = "000000005000"
    sats = 300
    
    case 4 'K�rsel
    wrtLine = 1
    lonart = "0024"
    'e1 = "000000005000"
    sats = 300
    
    case 5 'FROKOST: Kun timel�nnede ellers fastbel�b
    
        '** Kun timel�nende 
        'if cint(mgruppe) = 2 then
        'wrtLine = 1
        'else
        wrtLine = 0
        'end if
    
    lonart = "0094"
    'e1 = "000000000000"
    sats = 300
   
    case 14 'Ferie afholdt
    wrtLine = 0
    lonart = "0047"
    'e1 = "000000005000"
    sats = 300

    case 17 'Ferie afholdt u l�n.
        'if cdbl(sSeg(14)) = 0 then
        wrtLine = 0
        'else
        'rtLine = 0
        'end if
    lonart = "0047"
    'e1 = "000000005000"
    sats = 300

    case 33 'Aconto / udl�g
    wrtLine = 1
    'lonart = "0093"
    lonart = "0010"
    'e1 = "000000000000"
    sats = 300

    case 35 'Rejsegodtg�relse / di�ter Skattefrie
    '*** Mangler under et d�gn = konto 05

    wrtLine = 1
    lonart = "0005"
    'e1 = "000000005000"
    sats = 300

    case 36 'Rejsegodtg�relse / di�ter Skat > 24 timer
    '*** Mangler under et d�gn = konto 05

    wrtLine = 1
    lonart = "0007"
    'e1 = "000000005000"
    sats = 300

    case 8 '6 ferieuge brugt / Feriefri
        '** Kun timel�nende 
        'if cint(mgruppe) = 2 then
        'wrtLine = 1
        'else
        wrtLine = 0
        'end if
    lonart = "7778"
    'e1 = "000000000000"
    sats = 300

    case 21 'Syg timer
        '** Kun timel�nende 
        'if cint(mgruppe) = 2 then
        'wrtLine = 1
        'else
        wrtLine = 0
        'end if
    lonart = "0072"
    'e1 = "000000000000"
    sats = 300

    case 26 'Barn Syg timer
        '** Kun timel�nende 
        'if cint(mgruppe) = 2 then
        'wrtLine = 1
        'else
        wrtLine = 0
        'end if
    lonart = "5C6C"
    'e1 = "000000000000"
    sats = 300

    case 30 'Barsel dage
        '** Kun timel�nende 
        'if cint(mgruppe) = 2 then
        'wrtLine = 1
        'else
        wrtLine = 0
        'end if
    lonart = "6263"
    'e1 = "000000000000"
    sats = 300

    case 32 'Tjenenstefri
        '** Kun timel�nende 
        'if cint(mgruppe) = 2 then
        'wrtLine = 1
        'else
        wrtLine = 0
        'end if
    lonart = "0000"
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
    
   
   
    '*** Antal p� de variable
    select case cint(s) 
    case 0 
    antal = medarbSalary
    'case 4 'KM 

    '** Sats findes hvis en medarbejer har k�rt over 20000 p� �ret
    'if cdbl(kmtot) > 20000 then
    'antal = trim(sSeg(s)) * 1.96
    'else
    'antal = trim(sSeg(s)) * 3.52
    'end if
    
    'case 14, 17 'Ferie + Ferie afh. u l�n
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
    
        if antal_len = 1 then ' et tal efter komma DER SKAL ALTID V�RE 2
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

    '*** Tilf�j komma
    tilfojKomma = 0

    if cint(tilfojKomma) = 1 then

    
        antal_left = instr(antal, ",")

        if cint(antal_left) = 0 then
        antal = antal & ",00"
        end if

  

    end if

  

    '***
    '*** Skal der afrundes til n�rmeste 15 min.
    afrundestil15min = 0

    if cint(afrundestil15min) = 1 then
        if s = 200 then 'ordin�re - lunch

            'antal_timer = left(antal, antal_left-1)
            'if right(antal, 2) <> 00 then
            'antal = antal_timer&"50"
            'end if

        else

            if antal_left <> 0 then
            antal_timer = left(antal, antal_left-1)
  
                'afrund til n�rnmeste 15 min
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

  
    
    
        if (cint(wrtLine) = 1 AND antalLinjeTjk <> "0" ANd len(trim(antal)) <> 0) then 'OR (cint(s) = 0 AND cint(mgruppe) = 2 'Timel�nnede

            if antal/1 > 0 then

            sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";"& lonart &";"& antal 
            datalon_lm_fileOutput.WriteLine(sdLine)

            end if
            
              '** KM sats
           if cint(s) = 4 then 'Timer

                 if cdbl(kmtot) > 20000 then
                 kmsats = 196
                 else
                 kmsats = 352
                 end if

                 sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";0025;"& kmsats &""
                 datalon_lm_fileOutput.WriteLine(sdLine)

           end if


          '**** Indl�s medaprjeder faste L�n felter fra medarbejder profil
         indlaesMedarbejderLonfelter = 0
         if cint(indlaesMedarbejderLonfelter) = 1 then

         
           '** L�ntimer
           if cint(s) = 0 AND cint(mgruppe) <> 2 then 'Timer

                 sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";0001;16033"
                 datalon_lm_fileOutput.WriteLine(sdLine)

           end if


         if cint(s) = 0 then 'Pension

                 if cdbl(pension) <> 0 then
                 Pensionbel = replace(pension, ",", "")
                 else
                 Pensionbel = medarbSalary * 0.10
                 end if        

                 sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";0005;"& Pensionbel &";"
                 datalon_lm_fileOutput.WriteLine(sdLine)

           end if

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
                datalon_lm_fileOutput.WriteLine(sdLine)

           end if

           if cint(s) = 0 AND cint(sundhedsforsikring) = 1 then '007L sundhedsforsikring

                 sundhedsforsikringbel = 7492

                 sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";007L;"& sundhedsforsikringbel 
                 datalon_lm_fileOutput.WriteLine(sdLine)

           end if

           if cint(s) = 0 AND cdbl(bruttolontraek) <> 0 then '007L bruttolontraek

                 bruttolontraek = replace(bruttolontraek, ".", "")
                 bruttolontraek = replace(bruttolontraek, ",", "")

                 sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";0098;"& bruttolontraek 
                 datalon_lm_fileOutput.WriteLine(sdLine)

           end if

           if cint(s) = 0 AND cint(multimedieskat) = 1 then '0042 Multimedie

                 multimedieskatbel = 24167

                 sdLine = ""& arbejdsgiverId &";"& loentermin &";"& medarbLonnummer &";"& avdelingsno &";0042;"& multimedieskatbel 
                 datalon_lm_fileOutput.WriteLine(sdLine)

           end if

           end if 'indlaesMedarbejderLonfelter = 0
       
     
       

        end if

    'end if

    next 's
 

    end if 'do until medarbLonnummer

Loop



datalon_lm_fileIndput.Close

datalon_lm_fileOutput.Close



set datalon_lm_fileIndput=Nothing

set datalon_lm_fileOutput=Nothing



set datalon_lm_fs=Nothing



Response.Write "Script er udf�rt"






%>

</body>

</html>