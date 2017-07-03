
<% 
Dim objXMLHTTP_afstem, objXMLDOM_afstem, i_afstem, strHTML_afstem

Set objXMLDom_afstem = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_afstem = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_afstem.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/afstem_sprog.xml", False
'objXmlHttp_afstem.open "GET", "http://localhost/inc/xml/afstem_sprog.xml", False
'objXmlHttp_afstem.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/afstem_sprog.xml", False
objXmlHttp_afstem.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/afstem_sprog.xml", False
'objXmlHttp_afstem.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/afstem_sprog.xml", False
'objXmlHttp_afstem.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/afstem_sprog.xml", False

objXmlHttp_afstem.send 


Set objXmlDom_afstem = objXmlHttp_afstem.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_afstem = Nothing



Dim Address_afstem, Latitude_afstem, Longitude_afstem
Dim oNode_afstem, oNodes_afstem
Dim sXPathQuery_afstem

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
sXPathQuery_afstem = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_afstem = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_afstem = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_afstem = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_afstem = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_afstem = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_afstem = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_afstem = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_afstem = objXmlDom_afstem.documentElement.selectSingleNode(sXPathQuery_afstem)
Address_afstem = oNode_afstem.Text

Set oNodes_afstem = objXmlDom_afstem.documentElement.selectNodes(sXPathQuery_afstem)

    For Each oNode_afstem in oNodes_afstem

        afstem_txt_001 = oNode_afstem.selectSingleNode("txt_1").Text
        afstem_txt_002 = oNode_afstem.selectSingleNode("txt_2").Text
        afstem_txt_003 = oNode_afstem.selectSingleNode("txt_3").Text
        afstem_txt_004 = oNode_afstem.selectSingleNode("txt_4").Text
        afstem_txt_005 = oNode_afstem.selectSingleNode("txt_5").Text
        afstem_txt_006 = oNode_afstem.selectSingleNode("txt_6").Text
        afstem_txt_007 = oNode_afstem.selectSingleNode("txt_7").Text
        afstem_txt_008 = oNode_afstem.selectSingleNode("txt_8").Text
        afstem_txt_009 = oNode_afstem.selectSingleNode("txt_9").Text
        afstem_txt_010 = oNode_afstem.selectSingleNode("txt_10").Text

        afstem_txt_011 = oNode_afstem.selectSingleNode("txt_11").Text
        afstem_txt_012 = oNode_afstem.selectSingleNode("txt_12").Text
        afstem_txt_013 = oNode_afstem.selectSingleNode("txt_13").Text
        afstem_txt_014 = oNode_afstem.selectSingleNode("txt_14").Text
        afstem_txt_015 = oNode_afstem.selectSingleNode("txt_15").Text
        afstem_txt_016 = oNode_afstem.selectSingleNode("txt_16").Text
        afstem_txt_017 = oNode_afstem.selectSingleNode("txt_17").Text
        afstem_txt_018 = oNode_afstem.selectSingleNode("txt_18").Text
        afstem_txt_019 = oNode_afstem.selectSingleNode("txt_19").Text
        afstem_txt_020 = oNode_afstem.selectSingleNode("txt_20").Text

        afstem_txt_021 = oNode_afstem.selectSingleNode("txt_21").Text
        afstem_txt_022 = oNode_afstem.selectSingleNode("txt_22").Text
        afstem_txt_023 = oNode_afstem.selectSingleNode("txt_23").Text
        afstem_txt_024 = oNode_afstem.selectSingleNode("txt_24").Text
        afstem_txt_025 = oNode_afstem.selectSingleNode("txt_25").Text
        afstem_txt_026 = oNode_afstem.selectSingleNode("txt_26").Text
        afstem_txt_027 = oNode_afstem.selectSingleNode("txt_27").Text
        afstem_txt_028 = oNode_afstem.selectSingleNode("txt_28").Text
        afstem_txt_029 = oNode_afstem.selectSingleNode("txt_29").Text
        afstem_txt_030 = oNode_afstem.selectSingleNode("txt_30").Text

        afstem_txt_031 = oNode_afstem.selectSingleNode("txt_31").Text
        afstem_txt_032 = oNode_afstem.selectSingleNode("txt_32").Text
        afstem_txt_033 = oNode_afstem.selectSingleNode("txt_33").Text
        afstem_txt_034 = oNode_afstem.selectSingleNode("txt_34").Text
        afstem_txt_035 = oNode_afstem.selectSingleNode("txt_35").Text
        afstem_txt_036 = oNode_afstem.selectSingleNode("txt_36").Text
        afstem_txt_037 = oNode_afstem.selectSingleNode("txt_37").Text
        afstem_txt_038 = oNode_afstem.selectSingleNode("txt_38").Text
        afstem_txt_039 = oNode_afstem.selectSingleNode("txt_39").Text
        afstem_txt_040 = oNode_afstem.selectSingleNode("txt_40").Text

        afstem_txt_041 = oNode_afstem.selectSingleNode("txt_41").Text
        afstem_txt_042 = oNode_afstem.selectSingleNode("txt_42").Text
        afstem_txt_043 = oNode_afstem.selectSingleNode("txt_43").Text
        afstem_txt_044 = oNode_afstem.selectSingleNode("txt_44").Text
        afstem_txt_045 = oNode_afstem.selectSingleNode("txt_45").Text
        afstem_txt_046 = oNode_afstem.selectSingleNode("txt_46").Text
        afstem_txt_047 = oNode_afstem.selectSingleNode("txt_47").Text
        afstem_txt_048 = oNode_afstem.selectSingleNode("txt_48").Text
        afstem_txt_049 = oNode_afstem.selectSingleNode("txt_49").Text
        afstem_txt_050 = oNode_afstem.selectSingleNode("txt_50").Text

        afstem_txt_051 = oNode_afstem.selectSingleNode("txt_51").Text
        afstem_txt_052 = oNode_afstem.selectSingleNode("txt_52").Text
        afstem_txt_053 = oNode_afstem.selectSingleNode("txt_53").Text
        afstem_txt_054 = oNode_afstem.selectSingleNode("txt_54").Text
        afstem_txt_055 = oNode_afstem.selectSingleNode("txt_55").Text
        afstem_txt_056 = oNode_afstem.selectSingleNode("txt_56").Text
        afstem_txt_057 = oNode_afstem.selectSingleNode("txt_57").Text
        afstem_txt_058 = oNode_afstem.selectSingleNode("txt_58").Text
        afstem_txt_059 = oNode_afstem.selectSingleNode("txt_59").Text
        afstem_txt_060 = oNode_afstem.selectSingleNode("txt_60").Text

        afstem_txt_061 = oNode_afstem.selectSingleNode("txt_61").Text
        afstem_txt_062 = oNode_afstem.selectSingleNode("txt_62").Text
        afstem_txt_063 = oNode_afstem.selectSingleNode("txt_63").Text
        afstem_txt_064 = oNode_afstem.selectSingleNode("txt_64").Text
        afstem_txt_065 = oNode_afstem.selectSingleNode("txt_65").Text
        afstem_txt_066 = oNode_afstem.selectSingleNode("txt_66").Text
        afstem_txt_067 = oNode_afstem.selectSingleNode("txt_67").Text
        afstem_txt_068 = oNode_afstem.selectSingleNode("txt_68").Text
        afstem_txt_069 = oNode_afstem.selectSingleNode("txt_69").Text
        afstem_txt_070 = oNode_afstem.selectSingleNode("txt_70").Text

        afstem_txt_071 = oNode_afstem.selectSingleNode("txt_71").Text
        afstem_txt_072 = oNode_afstem.selectSingleNode("txt_72").Text
        afstem_txt_073 = oNode_afstem.selectSingleNode("txt_73").Text
        afstem_txt_074 = oNode_afstem.selectSingleNode("txt_74").Text
        afstem_txt_075 = oNode_afstem.selectSingleNode("txt_75").Text
        afstem_txt_076 = oNode_afstem.selectSingleNode("txt_76").Text
        afstem_txt_077 = oNode_afstem.selectSingleNode("txt_77").Text
        afstem_txt_078 = oNode_afstem.selectSingleNode("txt_78").Text
        afstem_txt_079 = oNode_afstem.selectSingleNode("txt_79").Text
        afstem_txt_080 = oNode_afstem.selectSingleNode("txt_80").Text

        afstem_txt_081 = oNode_afstem.selectSingleNode("txt_81").Text
        afstem_txt_082 = oNode_afstem.selectSingleNode("txt_82").Text
        afstem_txt_083 = oNode_afstem.selectSingleNode("txt_83").Text
        afstem_txt_084 = oNode_afstem.selectSingleNode("txt_84").Text
        afstem_txt_085 = oNode_afstem.selectSingleNode("txt_85").Text
        afstem_txt_086 = oNode_afstem.selectSingleNode("txt_86").Text
        afstem_txt_087 = oNode_afstem.selectSingleNode("txt_87").Text
        afstem_txt_088 = oNode_afstem.selectSingleNode("txt_88").Text
        afstem_txt_089 = oNode_afstem.selectSingleNode("txt_89").Text
        afstem_txt_090 = oNode_afstem.selectSingleNode("txt_90").Text

        afstem_txt_091 = oNode_afstem.selectSingleNode("txt_91").Text
        afstem_txt_092 = oNode_afstem.selectSingleNode("txt_92").Text
        afstem_txt_093 = oNode_afstem.selectSingleNode("txt_93").Text
        afstem_txt_094 = oNode_afstem.selectSingleNode("txt_94").Text
        afstem_txt_095 = oNode_afstem.selectSingleNode("txt_95").Text
        afstem_txt_096 = oNode_afstem.selectSingleNode("txt_96").Text
        afstem_txt_097 = oNode_afstem.selectSingleNode("txt_97").Text
        afstem_txt_098 = oNode_afstem.selectSingleNode("txt_98").Text
        afstem_txt_099 = oNode_afstem.selectSingleNode("txt_99").Text
        afstem_txt_100 = oNode_afstem.selectSingleNode("txt_100").Text
        afstem_txt_101 = oNode_afstem.selectSingleNode("txt_101").Text
        afstem_txt_102 = oNode_afstem.selectSingleNode("txt_102").Text
        afstem_txt_103 = oNode_afstem.selectSingleNode("txt_103").Text
        afstem_txt_104 = oNode_afstem.selectSingleNode("txt_104").Text
        afstem_txt_105 = oNode_afstem.selectSingleNode("txt_105").Text
        afstem_txt_106 = oNode_afstem.selectSingleNode("txt_106").Text
        afstem_txt_107 = oNode_afstem.selectSingleNode("txt_107").Text
        afstem_txt_108 = oNode_afstem.selectSingleNode("txt_108").Text
        afstem_txt_109 = oNode_afstem.selectSingleNode("txt_109").Text
        afstem_txt_110 = oNode_afstem.selectSingleNode("txt_110").Text
        afstem_txt_111 = oNode_afstem.selectSingleNode("txt_111").Text
        afstem_txt_112 = oNode_afstem.selectSingleNode("txt_112").Text
        afstem_txt_113 = oNode_afstem.selectSingleNode("txt_113").Text
        afstem_txt_114 = oNode_afstem.selectSingleNode("txt_114").Text
        afstem_txt_115 = oNode_afstem.selectSingleNode("txt_115").Text
        afstem_txt_116 = oNode_afstem.selectSingleNode("txt_116").Text
        afstem_txt_117 = oNode_afstem.selectSingleNode("txt_117").Text
        afstem_txt_118 = oNode_afstem.selectSingleNode("txt_118").Text
        afstem_txt_119 = oNode_afstem.selectSingleNode("txt_119").Text
        afstem_txt_120 = oNode_afstem.selectSingleNode("txt_120").Text
        afstem_txt_121 = oNode_afstem.selectSingleNode("txt_121").Text
        afstem_txt_122 = oNode_afstem.selectSingleNode("txt_122").Text
        afstem_txt_123 = oNode_afstem.selectSingleNode("txt_123").Text
        afstem_txt_124 = oNode_afstem.selectSingleNode("txt_124").Text
        afstem_txt_125 = oNode_afstem.selectSingleNode("txt_125").Text
        afstem_txt_126 = oNode_afstem.selectSingleNode("txt_126").Text
        afstem_txt_127 = oNode_afstem.selectSingleNode("txt_127").Text
        afstem_txt_128 = oNode_afstem.selectSingleNode("txt_128").Text
        afstem_txt_129 = oNode_afstem.selectSingleNode("txt_129").Text
        afstem_txt_130 = oNode_afstem.selectSingleNode("txt_130").Text
        afstem_txt_131 = oNode_afstem.selectSingleNode("txt_131").Text
        afstem_txt_132 = oNode_afstem.selectSingleNode("txt_132").Text



  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>