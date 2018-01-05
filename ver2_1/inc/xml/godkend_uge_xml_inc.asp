
<% 
Dim objXMLHTTP_godkendweek, objXMLDOM_godkendweek, i_godkendweek, strHTML_godkendweek

Set objXMLDom_godkendweek = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_godkendweek = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_godkendweek.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/godkend_uge_sprog.xml", False
objXmlHttp_godkendweek.open "GET", "http://localhost/inc/xml/godkend_uge_sprog.xml", False
'objXmlHttp_godkendweek.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/godkend_uge_sprog.xml", False
'objXmlHttp_godkendweek.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/godkend_uge_sprog.xml", False
'objXmlHttp_godkendweek.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/godkend_uge_sprog.xml", False
'objXmlHttp_godkendweek.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/godkend_uge_sprog.xml", False

objXmlHttp_godkendweek.send


Set objXmlDom_godkendweek = objXmlHttp_godkendweek.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_godkendweek = Nothing



Dim Address_godkendweek, Latitude_godkendweek, Longitude_godkendweek
Dim oNode_godkendweek, oNodes_godkendweek
Dim sXPathQuery_godkendweek

sprog = 1 'DK
if len(trim(session("mid"))) <> 0 then
strSQL = "SELECT sprog FROM medarbejdere WHERE mid = " & session("mid")

oRec.open strSQL, oConn, 3
if not oRec.EOF then
sprog = oRec("sprog")
end if
oRec.close
end if

select case sprog
case 1
sXPathQuery_godkendweek = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_godkendweek = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_godkendweek = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_godkendweek = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_godkendweek = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_godkendweek = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_godkendweek = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_godkendweek = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_godkendweek = objXmlDom_godkendweek.documentElement.selectSingleNode(sXPathQuery_godkendweek)
Address_godkendweek = oNode_godkendweek.Text

Set oNodes_godkendweek = objXmlDom_godkendweek.documentElement.selectNodes(sXPathQuery_godkendweek)

    For Each oNode_godkendweek in oNodes_godkendweek

    godkendweek_txt_001 = oNode_godkendweek.selectSingleNode("txt_1").Text
    godkendweek_txt_002 = oNode_godkendweek.selectSingleNode("txt_2").Text
    godkendweek_txt_003 = oNode_godkendweek.selectSingleNode("txt_3").Text
    godkendweek_txt_004 = oNode_godkendweek.selectSingleNode("txt_4").Text
    godkendweek_txt_005 = oNode_godkendweek.selectSingleNode("txt_5").Text
    godkendweek_txt_006 = oNode_godkendweek.selectSingleNode("txt_6").Text
    godkendweek_txt_007 = oNode_godkendweek.selectSingleNode("txt_7").Text
    godkendweek_txt_008 = oNode_godkendweek.selectSingleNode("txt_8").Text
    godkendweek_txt_009 = oNode_godkendweek.selectSingleNode("txt_9").Text
    godkendweek_txt_010 = oNode_godkendweek.selectSingleNode("txt_10").Text

    godkendweek_txt_011 = oNode_godkendweek.selectSingleNode("txt_11").Text
    godkendweek_txt_012 = oNode_godkendweek.selectSingleNode("txt_12").Text
    godkendweek_txt_013 = oNode_godkendweek.selectSingleNode("txt_13").Text
    godkendweek_txt_014 = oNode_godkendweek.selectSingleNode("txt_14").Text
    godkendweek_txt_015 = oNode_godkendweek.selectSingleNode("txt_15").Text
    godkendweek_txt_016 = oNode_godkendweek.selectSingleNode("txt_16").Text
    godkendweek_txt_017 = oNode_godkendweek.selectSingleNode("txt_17").Text
    godkendweek_txt_018 = oNode_godkendweek.selectSingleNode("txt_18").Text
    godkendweek_txt_019 = oNode_godkendweek.selectSingleNode("txt_19").Text
    godkendweek_txt_020 = oNode_godkendweek.selectSingleNode("txt_20").Text

    godkendweek_txt_021 = oNode_godkendweek.selectSingleNode("txt_21").Text
    godkendweek_txt_022 = oNode_godkendweek.selectSingleNode("txt_22").Text
    godkendweek_txt_023 = oNode_godkendweek.selectSingleNode("txt_23").Text
    godkendweek_txt_024 = oNode_godkendweek.selectSingleNode("txt_24").Text
    godkendweek_txt_025 = oNode_godkendweek.selectSingleNode("txt_25").Text
    godkendweek_txt_026 = oNode_godkendweek.selectSingleNode("txt_26").Text
    godkendweek_txt_027 = oNode_godkendweek.selectSingleNode("txt_27").Text
    godkendweek_txt_028 = oNode_godkendweek.selectSingleNode("txt_28").Text
    godkendweek_txt_029 = oNode_godkendweek.selectSingleNode("txt_29").Text
    godkendweek_txt_030 = oNode_godkendweek.selectSingleNode("txt_30").Text

    godkendweek_txt_031 = oNode_godkendweek.selectSingleNode("txt_31").Text
    godkendweek_txt_032 = oNode_godkendweek.selectSingleNode("txt_32").Text
    godkendweek_txt_033 = oNode_godkendweek.selectSingleNode("txt_33").Text
    godkendweek_txt_034 = oNode_godkendweek.selectSingleNode("txt_34").Text
    godkendweek_txt_035 = oNode_godkendweek.selectSingleNode("txt_35").Text
    godkendweek_txt_036 = oNode_godkendweek.selectSingleNode("txt_36").Text
    godkendweek_txt_037 = oNode_godkendweek.selectSingleNode("txt_37").Text
    godkendweek_txt_038 = oNode_godkendweek.selectSingleNode("txt_38").Text
    godkendweek_txt_039 = oNode_godkendweek.selectSingleNode("txt_39").Text
    godkendweek_txt_040 = oNode_godkendweek.selectSingleNode("txt_40").Text

    godkendweek_txt_041 = oNode_godkendweek.selectSingleNode("txt_41").Text
    godkendweek_txt_042 = oNode_godkendweek.selectSingleNode("txt_42").Text
    godkendweek_txt_043 = oNode_godkendweek.selectSingleNode("txt_43").Text
    godkendweek_txt_044 = oNode_godkendweek.selectSingleNode("txt_44").Text
    godkendweek_txt_045 = oNode_godkendweek.selectSingleNode("txt_45").Text
    godkendweek_txt_046 = oNode_godkendweek.selectSingleNode("txt_46").Text
    godkendweek_txt_047 = oNode_godkendweek.selectSingleNode("txt_47").Text
    godkendweek_txt_048 = oNode_godkendweek.selectSingleNode("txt_48").Text
    godkendweek_txt_049 = oNode_godkendweek.selectSingleNode("txt_49").Text
    godkendweek_txt_050 = oNode_godkendweek.selectSingleNode("txt_50").Text

    godkendweek_txt_051 = oNode_godkendweek.selectSingleNode("txt_51").Text
    godkendweek_txt_052 = oNode_godkendweek.selectSingleNode("txt_52").Text
    godkendweek_txt_053 = oNode_godkendweek.selectSingleNode("txt_53").Text
    godkendweek_txt_054 = oNode_godkendweek.selectSingleNode("txt_54").Text
    godkendweek_txt_055 = oNode_godkendweek.selectSingleNode("txt_55").Text
    godkendweek_txt_056 = oNode_godkendweek.selectSingleNode("txt_56").Text
    godkendweek_txt_057 = oNode_godkendweek.selectSingleNode("txt_57").Text
    godkendweek_txt_058 = oNode_godkendweek.selectSingleNode("txt_58").Text
    godkendweek_txt_059 = oNode_godkendweek.selectSingleNode("txt_59").Text
    godkendweek_txt_060 = oNode_godkendweek.selectSingleNode("txt_60").Text

    godkendweek_txt_061 = oNode_godkendweek.selectSingleNode("txt_61").Text
    godkendweek_txt_062 = oNode_godkendweek.selectSingleNode("txt_62").Text
    godkendweek_txt_063 = oNode_godkendweek.selectSingleNode("txt_63").Text
    godkendweek_txt_064 = oNode_godkendweek.selectSingleNode("txt_64").Text
    godkendweek_txt_065 = oNode_godkendweek.selectSingleNode("txt_65").Text
    godkendweek_txt_066 = oNode_godkendweek.selectSingleNode("txt_66").Text
    godkendweek_txt_067 = oNode_godkendweek.selectSingleNode("txt_67").Text
    godkendweek_txt_068 = oNode_godkendweek.selectSingleNode("txt_68").Text
    godkendweek_txt_069 = oNode_godkendweek.selectSingleNode("txt_69").Text
    godkendweek_txt_070 = oNode_godkendweek.selectSingleNode("txt_70").Text

    godkendweek_txt_071 = oNode_godkendweek.selectSingleNode("txt_71").Text
    godkendweek_txt_072 = oNode_godkendweek.selectSingleNode("txt_72").Text
    godkendweek_txt_073 = oNode_godkendweek.selectSingleNode("txt_73").Text
    godkendweek_txt_074 = oNode_godkendweek.selectSingleNode("txt_74").Text
    godkendweek_txt_075 = oNode_godkendweek.selectSingleNode("txt_75").Text
    godkendweek_txt_076 = oNode_godkendweek.selectSingleNode("txt_76").Text
    godkendweek_txt_077 = oNode_godkendweek.selectSingleNode("txt_77").Text
    godkendweek_txt_078 = oNode_godkendweek.selectSingleNode("txt_78").Text
    godkendweek_txt_079 = oNode_godkendweek.selectSingleNode("txt_79").Text
    godkendweek_txt_080 = oNode_godkendweek.selectSingleNode("txt_80").Text

    godkendweek_txt_081 = oNode_godkendweek.selectSingleNode("txt_81").Text
    godkendweek_txt_082 = oNode_godkendweek.selectSingleNode("txt_82").Text
    godkendweek_txt_083 = oNode_godkendweek.selectSingleNode("txt_83").Text
    godkendweek_txt_084 = oNode_godkendweek.selectSingleNode("txt_84").Text
    godkendweek_txt_085 = oNode_godkendweek.selectSingleNode("txt_85").Text
    godkendweek_txt_086 = oNode_godkendweek.selectSingleNode("txt_86").Text
    godkendweek_txt_087 = oNode_godkendweek.selectSingleNode("txt_87").Text
    godkendweek_txt_088 = oNode_godkendweek.selectSingleNode("txt_88").Text
    godkendweek_txt_089 = oNode_godkendweek.selectSingleNode("txt_89").Text
    godkendweek_txt_090 = oNode_godkendweek.selectSingleNode("txt_90").Text
    godkendweek_txt_091 = oNode_godkendweek.selectSingleNode("txt_91").Text
    godkendweek_txt_092 = oNode_godkendweek.selectSingleNode("txt_92").Text
    godkendweek_txt_093 = oNode_godkendweek.selectSingleNode("txt_93").Text
    godkendweek_txt_094 = oNode_godkendweek.selectSingleNode("txt_94").Text
    godkendweek_txt_095 = oNode_godkendweek.selectSingleNode("txt_95").Text
    godkendweek_txt_096 = oNode_godkendweek.selectSingleNode("txt_96").Text
    godkendweek_txt_097 = oNode_godkendweek.selectSingleNode("txt_97").Text
    godkendweek_txt_098 = oNode_godkendweek.selectSingleNode("txt_98").Text
    godkendweek_txt_099 = oNode_godkendweek.selectSingleNode("txt_99").Text
    godkendweek_txt_100 = oNode_godkendweek.selectSingleNode("txt_100").Text
    godkendweek_txt_101 = oNode_godkendweek.selectSingleNode("txt_101").Text
    godkendweek_txt_102 = oNode_godkendweek.selectSingleNode("txt_102").Text
    godkendweek_txt_103 = oNode_godkendweek.selectSingleNode("txt_103").Text
    godkendweek_txt_104 = oNode_godkendweek.selectSingleNode("txt_104").Text
    godkendweek_txt_105 = oNode_godkendweek.selectSingleNode("txt_105").Text
    godkendweek_txt_106 = oNode_godkendweek.selectSingleNode("txt_106").Text
    godkendweek_txt_107 = oNode_godkendweek.selectSingleNode("txt_107").Text
    godkendweek_txt_108 = oNode_godkendweek.selectSingleNode("txt_108").Text
    godkendweek_txt_109 = oNode_godkendweek.selectSingleNode("txt_109").Text
    godkendweek_txt_110 = oNode_godkendweek.selectSingleNode("txt_110").Text
    godkendweek_txt_111 = oNode_godkendweek.selectSingleNode("txt_111").Text
    godkendweek_txt_112 = oNode_godkendweek.selectSingleNode("txt_112").Text
    godkendweek_txt_113 = oNode_godkendweek.selectSingleNode("txt_113").Text
    godkendweek_txt_114 = oNode_godkendweek.selectSingleNode("txt_114").Text
    godkendweek_txt_115 = oNode_godkendweek.selectSingleNode("txt_115").Text
    godkendweek_txt_116 = oNode_godkendweek.selectSingleNode("txt_116").Text
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>