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
    'case 4 'Overtid 50%
    'case 5 'Overtid 100%
    'case 6 'Reisetid ordinær
    'case 7 'Reisetid 100% 
    'case 8 'Reisetid 50%
    'case 9 'Utetillæg utland
    'case 10 'Efter midnat
    'case 11 'Timebank
    'case 15 'Avspad m løn
    'case 18 'Ordinære timer
    'case 20 'Avspadsering egenbetalt 
    'case 22 'Teamleder tillæg



    select case s
    case 3 'Utetillæg innland
     wrtLine = 1
    lonart = "00020"
    'e1 = "000000005000"
    sats = 300
    case 4 'Overtid 50%
     wrtLine = 1
    lonart = "00025"
     'e1 = "000000005000"
    sats = 300
    case 5 'Overtid 100%
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
    case 11 'Timebank
     wrtLine = 1
    lonart = "00400"
    'e1 = "000000000000"
    sats = 300
    case 15 'Avspad m løn
     wrtLine = 1
    lonart = "00402"
    'e1 = "000000005000"
    sats = 300
    case 18 'Ordinære timer
     wrtLine = 1
    lonart = "00010"
    'e1 = "000000005000"
    sats = 300
    case 20 'Avspadsering egenbetalt 
     wrtLine = 1
    lonart = "00401"
    'e1 = "000000000000"
    sats = 300
    case 22 'Teamleder tillæg
    wrtLine = 1
    lonart = "00022"
    'e1 = "000000005000"
    sats = 300
    case else
    wrtLine = 0
    sats = 0
    lonart = "00000"
    'e1 = "000000000000"
    end select

    e1 = "000000000000"

    

    'if len(trim(sSeg(s))) <> 0 then
    'belob = sSeg(s)*sats '"1234567890"
    'else
    belob = 0
    'end if
    
    antal = trim(sSeg(s))
    antal_left = instr(antal, ",")

    antal = replace(antal, ".", "")
    antal = replace(antal, ",", "")

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


    antalLinjeTjk = antal

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



'Function CreateSDLine(csvLine, Instkode, Tjenstekode, loenArt1, loenArt2, enheder, enheder2)
Function CreateHLTLine(csvLine, Instkode, Tjenstekode, loenArt1, loenArt2, enheder, enheder2)


    if InStr(enheder,",") > 0 then

         enheder = Replace(enheder,",","")
        
    
    end if

    if (loenArt1 = "790") then


        ''** Syg altid ,00 i dage ***'
        
        enhederDecimaler_len = len(enheder)
        enhederDecimaler_left = left(enheder, enhederDecimaler_len-2)
        enheder = enhederDecimaler_left & "00"
       
    end if


     if (loenArt1 = "910") then 


        ''** Barnsyg altid min 1,00 i dage ***'
        
        enhederDecimaler_len = len(enheder)
        enhederDecimaler_left = left(enheder, enhederDecimaler_len-2)

        if cint(enhederDecimaler_left) < 1 then
        enheder = "100"
        else
        enheder = enhederDecimaler_left & "00"
        end if

        'Response.write "LA:" & loenArt1 & " enh: "& enheder & " ..."
    'Response.flush 
    'strEnheder = "" & Lpad(enheder,"0",5) '4

    'Response.write "strEnheder:" & strEnheder &"<br><br>"
    'Response.flush 
       
    end if


    strEnheder = "" & Lpad(enheder,"0",5) '4




    if InStr(enheder,"-") > 0  then

        strChar = ""



        'Formattering af enheder til håndtering af minus jft sd løn format

        if StrComp(Right(strEnheder,1),"0") = 0 then strEnheder = left(strEnheder,4) & "å" end if

        if StrComp(right(strEnheder,1),"1") = 0 then strEnheder = left(strEnheder,4) & "J" end if

        if StrComp(right(strEnheder,1),"2") = 0 then strEnheder = left(strEnheder,4) & "K" end if

        if StrComp(right(strEnheder,1),"3") = 0 then strEnheder = left(strEnheder,4) & "L" end if

        if StrComp(right(strEnheder,1),"4") = 0 then strEnheder = left(strEnheder,4) & "M" end if

        if StrComp(right(strEnheder,1),"5") = 0 then strEnheder = left(strEnheder,4) & "N" end if

        if StrComp(right(strEnheder,1),"6") = 0 then strEnheder = left(strEnheder,4) & "O" end if

        if StrComp(right(strEnheder,1),"7") = 0 then strEnheder = left(strEnheder,4) & "P" end if

        if StrComp(right(strEnheder,1),"8") = 0 then strEnheder = left(strEnheder,4) & "Q" end if

        if StrComp(right(strEnheder,1),"9") = 0 then strEnheder = left(strEnheder,4) & "R" end if

 

    end if


    
    
    '**** Enheder 2 ******
    if InStr(enheder2,",") > 0 then

        enheder2 = Replace(enheder2,",","")

    end if

    
    'Response.write "LA:" & loenArt2 & " enh: "& enheder2 & "<br>"
    'Response.flush 
    strEnheder2 = "" & Lpad(enheder2,"0",5) '4



   



    sdLine = ""



   '001-002 Inst.kode

    sdLine =  Instkode 'Lpad(Instkode," ",2)



    '003-004 Filler

    sdLine = sdLine & "  "



    '005-005  KortArt altid på 7

    sdLine = sdLine & "7"

    

    '006 Tjenstekode værdier 0, 1, 2, 3

    sdLine = sdLine & Tjenstekode 



    '007-011 Tjenstenummer

    sdLine = sdLine & Lpad(trim(csvLine(1)),"0",5) '"0"



    '012-015 Afdeling

    sdLine = sdLine & "    "



    '016-017 År

    'Response.write csvLine(hlt_datoFieldNo) & "<br>"
    'Response.flush

    if len(trim(csvLine(hlt_datoFieldNo))) <> 0 AND hlt_datoFieldNo <> 32 then
    csvLine(hlt_datoFieldNo) = csvLine(hlt_datoFieldNo)
    else
    csvLine(hlt_datoFieldNo) = day(now) &"-"& month(now) &"-"& year(now)
    end if 

    yVal = datepart("yyyy", csvLine(hlt_datoFieldNo), 2,2)
    yVal = right(yVal, 2)
    sdLine = sdLine & yVal 
    'sdLine = sdLine & csvLine(27) &"##"


    '018-019 Måned
    
    mdVal = datepart("m", csvLine(hlt_datoFieldNo), 2,2)
    'mdVal = mid(csvLine(27),6,2)
    if len(trim(mdVal)) = 1 then ' md < 10
    mdVal = "0"&left(mdVal, 1)
    else
    mdVal = mdVal
    end if
    sdLine = sdLine & mdVal

   'sdLine = sdLine & left(Right(csvLine(25),7),2)
   
    '020-021 Uge / Blank

    sdLine = sdLine & "  "



    '022-023 dag
    dVal = datepart("d", csvLine(hlt_datoFieldNo), 2,2)
    if len(trim((dVal))) = 1 then 'dag < 10
    sdLine = sdLine & "0"&dVal
    else
    sdLine = sdLine & dVal
    end if
    
    'sdLine = sdLine & "dval:#" & dVal & "#"

    '024-026 lønart

    sdLine = sdLine & loenArt1



    '027-031 Enheder

    sdLine = sdLine & strEnheder



    '032 - 034 
    if Len(loenArt2) > 0 then

        sdLine = sdLine & loenArt2 & Lpad("","0",3)
        '035 - 038 'Sekunder lønart i timer 
        sdLine = sdLine & strEnheder2

    else
        
        sdLine = sdLine & "00000000"

    end if


    


    '096-099 Filler

    sdLine = Rpad(sdLine," ",99)



    '100-109 Konto

    sdLine =  Rpad(sdLine," ",109)



    '110-115 Filler

    sdLine =  Rpad(sdLine," ",115)



    '116 Intern behandlingskode Værdien skal være ’K’ for KMD eller ’T’ for Tabulex

    sdLine = sdLine & "T"



    '117 Særlig kode 6 = sygdom, 8 = ferie

    if StrComp(loenArt2,"790") = 0 then

        sdLine = sdLine & "6"

    end if



    if StrComp(loenArt2,"752") = 0 then

        sdLine = sdLine & "8"

    end if





    if StrComp(loenArt2,"752") <> 0 and StrComp(loenArt2,"790") <> 0  then

        sdLine = sdLine & " "

    end if



    '118-128 Filler

    sdLine = Rpad(sdLine," ",128)

   

    CreateHLTLine = sdLine

End Function


'*** Bagvedstillede nuller ****
Function Rpad (sValue, sPadchar, iLength)
  
  if len(trim(sValue)) <> 0 then
  Rpad = sValue & string(iLength - Len(sValue), sPadchar)
  end if

End Function


'*** foranstillede nuller ****
Function Lpad (sValue, sPadchar, iLength)

  if len(trim(sPadchar)) <> 0 then
  sPadchar = sPadchar
  else
  sPadchar = 0
  end if

  

  if len(trim(sValue)) <> 0 then
  Lpad = string(iLength - Len(sValue),sPadchar) & sValue
  end if

 'Response.Write "iLength: " & iLength & " sPadchar: "& sPadchar & " sValue: " & sValue & " Lpad: "&  Lpad &"<br>"
 'Response.flush


End Function



%>

</body>

</html>