
<% 
Dim objXMLHTTP_stat_korsel, objXMLDOM_stat_korsel, i_stat_korsel, strHTML_stat_korsel

Set objXMLDom_stat_korsel = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_stat_korsel = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_stat_korsel.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/stat_korsel_sprog.xml", False
'objXmlHttp_stat_korsel.open "GET", "http://localhost/inc/xml/stat_korsel_sprog.xml", False
'objXmlHttp_stat_korsel.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/stat_korsel_sprog.xml", False
'objXmlHttp_stat_korsel.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/stat_korsel_sprog.xml", False
'objXmlHttp_stat_korsel.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/stat_korsel_sprog.xml", False
'objXmlHttp_stat_korsel.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/stat_korsel_sprog.xml", False
objXmlHttp_stat_korsel.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/stat_korsel_sprog.xml", False

objXmlHttp_stat_korsel.send 


Set objXmlDom_stat_korsel = objXmlHttp_stat_korsel.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_stat_korsel = Nothing



Dim Address_stat_korsel, Latitude_stat_korsel, Longitude_stat_korsel
Dim oNode_stat_korsel, oNodes_stat_korsel
Dim sXPathQuery_stat_korsel

sprog = 1 'DK
if len(trim(session("mid"))) <> 0 then
strSQL = "SELECT sprog FROM medarbejdere WHERE mid = " & session("mid")

oRec.open strSQL, oConn, 3
if not oRec.EOF then
sprog = oRec("sprog")
end if
oRec.close
end if

select case lto

    case "xxxmiele"
        sXPathQuery_stat_korsel = "//sprog/miele"

    case else
        select case sprog
        case 1
        sXPathQuery_stat_korsel = "//sprog/dk"
        'Session.LCID = 1030
        case 2
        sXPathQuery_stat_korsel = "//sprog/uk"
        'Session.LCID = 2057
        case 3
        sXPathQuery_stat_korsel = "//sprog/se"
        'Session.LCID = 1053
        case 4
        sXPathQuery_stat_korsel = "//sprog/no"
        'Session.LCID = 2068
        case 5
        sXPathQuery_stat_korsel = "//sprog/es"
        'Session.LCID = 1034
        case 6
        sXPathQuery_stat_korsel = "//sprog/de"
        'Session.LCID = 1031
        case 7
        sXPathQuery_stat_korsel = "//sprog/fr"
        'Session.LCID = 1036
        case else
        sXPathQuery_stat_korsel = "//sprog/dk"
        'Session.LCID = 1030
        end select
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_stat_korsel = objXmlDom_stat_korsel.documentElement.selectSingleNode(sXPathQuery_stat_korsel)
Address_stat_korsel = oNode_stat_korsel.Text

Set oNodes_stat_korsel = objXmlDom_stat_korsel.documentElement.selectNodes(sXPathQuery_stat_korsel)

    For Each oNode_stat_korsel in oNodes_stat_korsel

        stat_korsel_txt_001 = oNode_stat_korsel.selectSingleNode("txt_1").Text
        stat_korsel_txt_002 = oNode_stat_korsel.selectSingleNode("txt_2").Text
        stat_korsel_txt_003 = oNode_stat_korsel.selectSingleNode("txt_3").Text
        stat_korsel_txt_004 = oNode_stat_korsel.selectSingleNode("txt_4").Text
        stat_korsel_txt_005 = oNode_stat_korsel.selectSingleNode("txt_5").Text
        stat_korsel_txt_006 = oNode_stat_korsel.selectSingleNode("txt_6").Text
        stat_korsel_txt_007 = oNode_stat_korsel.selectSingleNode("txt_7").Text
        stat_korsel_txt_008 = oNode_stat_korsel.selectSingleNode("txt_8").Text
        stat_korsel_txt_009 = oNode_stat_korsel.selectSingleNode("txt_9").Text
        stat_korsel_txt_010 = oNode_stat_korsel.selectSingleNode("txt_10").Text

        stat_korsel_txt_011 = oNode_stat_korsel.selectSingleNode("txt_11").Text
        stat_korsel_txt_012 = oNode_stat_korsel.selectSingleNode("txt_12").Text
        stat_korsel_txt_013 = oNode_stat_korsel.selectSingleNode("txt_13").Text
        stat_korsel_txt_014 = oNode_stat_korsel.selectSingleNode("txt_14").Text
        stat_korsel_txt_015 = oNode_stat_korsel.selectSingleNode("txt_15").Text
        stat_korsel_txt_016 = oNode_stat_korsel.selectSingleNode("txt_16").Text
        stat_korsel_txt_017 = oNode_stat_korsel.selectSingleNode("txt_17").Text
        stat_korsel_txt_018 = oNode_stat_korsel.selectSingleNode("txt_18").Text
        stat_korsel_txt_019 = oNode_stat_korsel.selectSingleNode("txt_19").Text
        stat_korsel_txt_020 = oNode_stat_korsel.selectSingleNode("txt_20").Text

        stat_korsel_txt_021 = oNode_stat_korsel.selectSingleNode("txt_21").Text
        stat_korsel_txt_022 = oNode_stat_korsel.selectSingleNode("txt_22").Text
        stat_korsel_txt_023 = oNode_stat_korsel.selectSingleNode("txt_23").Text
        stat_korsel_txt_024 = oNode_stat_korsel.selectSingleNode("txt_24").Text
        stat_korsel_txt_025 = oNode_stat_korsel.selectSingleNode("txt_25").Text
        stat_korsel_txt_026 = oNode_stat_korsel.selectSingleNode("txt_26").Text
        stat_korsel_txt_027 = oNode_stat_korsel.selectSingleNode("txt_27").Text
        stat_korsel_txt_028 = oNode_stat_korsel.selectSingleNode("txt_28").Text
        stat_korsel_txt_029 = oNode_stat_korsel.selectSingleNode("txt_29").Text
        stat_korsel_txt_030 = oNode_stat_korsel.selectSingleNode("txt_30").Text

        stat_korsel_txt_031 = oNode_stat_korsel.selectSingleNode("txt_31").Text
        stat_korsel_txt_032 = oNode_stat_korsel.selectSingleNode("txt_32").Text
        stat_korsel_txt_033 = oNode_stat_korsel.selectSingleNode("txt_33").Text
        stat_korsel_txt_034 = oNode_stat_korsel.selectSingleNode("txt_34").Text
        stat_korsel_txt_035 = oNode_stat_korsel.selectSingleNode("txt_35").Text
        stat_korsel_txt_036 = oNode_stat_korsel.selectSingleNode("txt_36").Text
        stat_korsel_txt_037 = oNode_stat_korsel.selectSingleNode("txt_37").Text
        stat_korsel_txt_038 = oNode_stat_korsel.selectSingleNode("txt_38").Text
        stat_korsel_txt_039 = oNode_stat_korsel.selectSingleNode("txt_39").Text
        stat_korsel_txt_040 = oNode_stat_korsel.selectSingleNode("txt_40").Text

        stat_korsel_txt_041 = oNode_stat_korsel.selectSingleNode("txt_41").Text
        stat_korsel_txt_042 = oNode_stat_korsel.selectSingleNode("txt_42").Text
        stat_korsel_txt_043 = oNode_stat_korsel.selectSingleNode("txt_43").Text
        stat_korsel_txt_044 = oNode_stat_korsel.selectSingleNode("txt_44").Text
        stat_korsel_txt_045 = oNode_stat_korsel.selectSingleNode("txt_45").Text
        stat_korsel_txt_046 = oNode_stat_korsel.selectSingleNode("txt_46").Text
        stat_korsel_txt_047 = oNode_stat_korsel.selectSingleNode("txt_47").Text
        stat_korsel_txt_048 = oNode_stat_korsel.selectSingleNode("txt_48").Text
        stat_korsel_txt_049 = oNode_stat_korsel.selectSingleNode("txt_49").Text
        stat_korsel_txt_050 = oNode_stat_korsel.selectSingleNode("txt_50").Text

        stat_korsel_txt_051 = oNode_stat_korsel.selectSingleNode("txt_51").Text
        stat_korsel_txt_052 = oNode_stat_korsel.selectSingleNode("txt_52").Text
        stat_korsel_txt_053 = oNode_stat_korsel.selectSingleNode("txt_53").Text
        stat_korsel_txt_054 = oNode_stat_korsel.selectSingleNode("txt_54").Text
        stat_korsel_txt_055 = oNode_stat_korsel.selectSingleNode("txt_55").Text
        stat_korsel_txt_056 = oNode_stat_korsel.selectSingleNode("txt_56").Text
        stat_korsel_txt_057 = oNode_stat_korsel.selectSingleNode("txt_57").Text
        stat_korsel_txt_058 = oNode_stat_korsel.selectSingleNode("txt_58").Text
        stat_korsel_txt_059 = oNode_stat_korsel.selectSingleNode("txt_59").Text
        stat_korsel_txt_060 = oNode_stat_korsel.selectSingleNode("txt_60").Text

        stat_korsel_txt_061 = oNode_stat_korsel.selectSingleNode("txt_61").Text
        stat_korsel_txt_062 = oNode_stat_korsel.selectSingleNode("txt_62").Text
        stat_korsel_txt_063 = oNode_stat_korsel.selectSingleNode("txt_63").Text
        stat_korsel_txt_064 = oNode_stat_korsel.selectSingleNode("txt_64").Text
        stat_korsel_txt_065 = oNode_stat_korsel.selectSingleNode("txt_65").Text
        stat_korsel_txt_066 = oNode_stat_korsel.selectSingleNode("txt_66").Text
        stat_korsel_txt_067 = oNode_stat_korsel.selectSingleNode("txt_67").Text
        stat_korsel_txt_068 = oNode_stat_korsel.selectSingleNode("txt_68").Text
        stat_korsel_txt_069 = oNode_stat_korsel.selectSingleNode("txt_69").Text
        stat_korsel_txt_070 = oNode_stat_korsel.selectSingleNode("txt_70").Text
         
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>