
<% 
Dim objXMLHTTP_oms, objXMLDOM_oms, i_oms, strHTML_oms

Set objXMLDom_oms = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_oms = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_oms.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/oms_sprog.xml", False
'objXmlHttp_oms.open "GET", "http://localhost/inc/xml/oms_sprog.xml", False
'objXmlHttp_oms.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/oms_sprog.xml", False
'objXmlHttp_oms.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/oms_sprog.xml", False
objXmlHttp_oms.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/oms_sprog.xml", False
'objXmlHttp_oms.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/oms_sprog.xml", False
'objXmlHttp_oms.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/oms_sprog.xml", False

objXmlHttp_oms.send 


Set objXmlDom_oms = objXmlHttp_oms.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_oms = Nothing



Dim Address_oms, Latitude_oms, Longitude_oms
Dim oNode_oms, oNodes_oms
Dim sXPathQuery_oms

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
sXPathQuery_oms = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_oms = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_oms = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_oms = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_oms = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_oms = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_oms = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_oms = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_oms = objXmlDom_oms.documentElement.selectSingleNode(sXPathQuery_oms)
Address_oms = oNode_oms.Text

Set oNodes_oms = objXmlDom_oms.documentElement.selectNodes(sXPathQuery_oms)

    For Each oNode_oms in oNodes_oms

        oms_txt_001 = oNode_oms.selectSingleNode("txt_1").Text
        oms_txt_002 = oNode_oms.selectSingleNode("txt_2").Text
        oms_txt_003 = oNode_oms.selectSingleNode("txt_3").Text
        oms_txt_004 = oNode_oms.selectSingleNode("txt_4").Text
        oms_txt_005 = oNode_oms.selectSingleNode("txt_5").Text
        oms_txt_006 = oNode_oms.selectSingleNode("txt_6").Text
        oms_txt_007 = oNode_oms.selectSingleNode("txt_7").Text
        oms_txt_008 = oNode_oms.selectSingleNode("txt_8").Text
        oms_txt_009 = oNode_oms.selectSingleNode("txt_9").Text
        oms_txt_010 = oNode_oms.selectSingleNode("txt_10").Text

        oms_txt_011 = oNode_oms.selectSingleNode("txt_11").Text
        oms_txt_012 = oNode_oms.selectSingleNode("txt_12").Text
        oms_txt_013 = oNode_oms.selectSingleNode("txt_13").Text
        oms_txt_014 = oNode_oms.selectSingleNode("txt_14").Text
        oms_txt_015 = oNode_oms.selectSingleNode("txt_15").Text
        oms_txt_016 = oNode_oms.selectSingleNode("txt_16").Text
        oms_txt_017 = oNode_oms.selectSingleNode("txt_17").Text
        oms_txt_018 = oNode_oms.selectSingleNode("txt_18").Text
        oms_txt_019 = oNode_oms.selectSingleNode("txt_19").Text
        oms_txt_020 = oNode_oms.selectSingleNode("txt_20").Text

        oms_txt_021 = oNode_oms.selectSingleNode("txt_21").Text
        oms_txt_022 = oNode_oms.selectSingleNode("txt_22").Text
        oms_txt_023 = oNode_oms.selectSingleNode("txt_23").Text
        oms_txt_024 = oNode_oms.selectSingleNode("txt_24").Text
        oms_txt_025 = oNode_oms.selectSingleNode("txt_25").Text
        oms_txt_026 = oNode_oms.selectSingleNode("txt_26").Text
        oms_txt_027 = oNode_oms.selectSingleNode("txt_27").Text
        oms_txt_028 = oNode_oms.selectSingleNode("txt_28").Text
        oms_txt_029 = oNode_oms.selectSingleNode("txt_29").Text
        oms_txt_030 = oNode_oms.selectSingleNode("txt_30").Text

        oms_txt_031 = oNode_oms.selectSingleNode("txt_31").Text
        oms_txt_032 = oNode_oms.selectSingleNode("txt_32").Text
        oms_txt_033 = oNode_oms.selectSingleNode("txt_33").Text
        oms_txt_034 = oNode_oms.selectSingleNode("txt_34").Text
        oms_txt_035 = oNode_oms.selectSingleNode("txt_35").Text
        oms_txt_036 = oNode_oms.selectSingleNode("txt_36").Text
        oms_txt_037 = oNode_oms.selectSingleNode("txt_37").Text
        oms_txt_038 = oNode_oms.selectSingleNode("txt_38").Text
        oms_txt_039 = oNode_oms.selectSingleNode("txt_39").Text
        oms_txt_040 = oNode_oms.selectSingleNode("txt_40").Text

        oms_txt_041 = oNode_oms.selectSingleNode("txt_41").Text
        oms_txt_042 = oNode_oms.selectSingleNode("txt_42").Text
        oms_txt_043 = oNode_oms.selectSingleNode("txt_43").Text
        oms_txt_044 = oNode_oms.selectSingleNode("txt_44").Text
        oms_txt_045 = oNode_oms.selectSingleNode("txt_45").Text
        oms_txt_046 = oNode_oms.selectSingleNode("txt_46").Text
        oms_txt_047 = oNode_oms.selectSingleNode("txt_47").Text
        oms_txt_048 = oNode_oms.selectSingleNode("txt_48").Text
        oms_txt_049 = oNode_oms.selectSingleNode("txt_49").Text
        oms_txt_050 = oNode_oms.selectSingleNode("txt_50").Text

        oms_txt_051 = oNode_oms.selectSingleNode("txt_51").Text
        oms_txt_052 = oNode_oms.selectSingleNode("txt_52").Text
        oms_txt_053 = oNode_oms.selectSingleNode("txt_53").Text
        oms_txt_054 = oNode_oms.selectSingleNode("txt_54").Text
        oms_txt_055 = oNode_oms.selectSingleNode("txt_55").Text
        oms_txt_056 = oNode_oms.selectSingleNode("txt_56").Text
        oms_txt_057 = oNode_oms.selectSingleNode("txt_57").Text
        oms_txt_058 = oNode_oms.selectSingleNode("txt_58").Text
        oms_txt_059 = oNode_oms.selectSingleNode("txt_59").Text
        oms_txt_060 = oNode_oms.selectSingleNode("txt_60").Text

        oms_txt_061 = oNode_oms.selectSingleNode("txt_61").Text
        oms_txt_062 = oNode_oms.selectSingleNode("txt_62").Text
        oms_txt_063 = oNode_oms.selectSingleNode("txt_63").Text
        oms_txt_064 = oNode_oms.selectSingleNode("txt_64").Text
        oms_txt_065 = oNode_oms.selectSingleNode("txt_65").Text
        oms_txt_066 = oNode_oms.selectSingleNode("txt_66").Text
        oms_txt_067 = oNode_oms.selectSingleNode("txt_67").Text
        oms_txt_068 = oNode_oms.selectSingleNode("txt_68").Text
        oms_txt_069 = oNode_oms.selectSingleNode("txt_69").Text
        oms_txt_070 = oNode_oms.selectSingleNode("txt_70").Text

        oms_txt_071 = oNode_oms.selectSingleNode("txt_71").Text
        oms_txt_072 = oNode_oms.selectSingleNode("txt_72").Text
        oms_txt_073 = oNode_oms.selectSingleNode("txt_73").Text
        oms_txt_074 = oNode_oms.selectSingleNode("txt_74").Text
        oms_txt_075 = oNode_oms.selectSingleNode("txt_75").Text
        oms_txt_076 = oNode_oms.selectSingleNode("txt_76").Text
        oms_txt_077 = oNode_oms.selectSingleNode("txt_77").Text
        oms_txt_078 = oNode_oms.selectSingleNode("txt_78").Text
        oms_txt_079 = oNode_oms.selectSingleNode("txt_79").Text
        oms_txt_080 = oNode_oms.selectSingleNode("txt_80").Text

        oms_txt_081 = oNode_oms.selectSingleNode("txt_81").Text
        oms_txt_082 = oNode_oms.selectSingleNode("txt_82").Text
        oms_txt_083 = oNode_oms.selectSingleNode("txt_83").Text
        oms_txt_084 = oNode_oms.selectSingleNode("txt_84").Text
        oms_txt_085 = oNode_oms.selectSingleNode("txt_85").Text
        oms_txt_086 = oNode_oms.selectSingleNode("txt_86").Text
        oms_txt_087 = oNode_oms.selectSingleNode("txt_87").Text
        oms_txt_088 = oNode_oms.selectSingleNode("txt_88").Text
        oms_txt_089 = oNode_oms.selectSingleNode("txt_89").Text
        oms_txt_090 = oNode_oms.selectSingleNode("txt_90").Text

        oms_txt_091 = oNode_oms.selectSingleNode("txt_91").Text
        oms_txt_092 = oNode_oms.selectSingleNode("txt_92").Text
        oms_txt_093 = oNode_oms.selectSingleNode("txt_93").Text
        oms_txt_094 = oNode_oms.selectSingleNode("txt_94").Text
        oms_txt_095 = oNode_oms.selectSingleNode("txt_95").Text
        oms_txt_096 = oNode_oms.selectSingleNode("txt_96").Text
        oms_txt_097 = oNode_oms.selectSingleNode("txt_97").Text
        oms_txt_098 = oNode_oms.selectSingleNode("txt_98").Text
        oms_txt_099 = oNode_oms.selectSingleNode("txt_99").Text
        oms_txt_100 = oNode_oms.selectSingleNode("txt_100").Text
        oms_txt_101 = oNode_oms.selectSingleNode("txt_101").Text
        oms_txt_102 = oNode_oms.selectSingleNode("txt_102").Text
        oms_txt_103 = oNode_oms.selectSingleNode("txt_103").Text
        oms_txt_104 = oNode_oms.selectSingleNode("txt_104").Text
        oms_txt_105 = oNode_oms.selectSingleNode("txt_105").Text
        oms_txt_106 = oNode_oms.selectSingleNode("txt_106").Text
        oms_txt_107 = oNode_oms.selectSingleNode("txt_107").Text
        oms_txt_108 = oNode_oms.selectSingleNode("txt_108").Text
        oms_txt_109 = oNode_oms.selectSingleNode("txt_109").Text
        oms_txt_110 = oNode_oms.selectSingleNode("txt_110").Text
        oms_txt_111 = oNode_oms.selectSingleNode("txt_111").Text
        oms_txt_112 = oNode_oms.selectSingleNode("txt_112").Text
        oms_txt_113 = oNode_oms.selectSingleNode("txt_113").Text
        oms_txt_114 = oNode_oms.selectSingleNode("txt_114").Text
        oms_txt_115 = oNode_oms.selectSingleNode("txt_115").Text
        oms_txt_116 = oNode_oms.selectSingleNode("txt_116").Text
        oms_txt_117 = oNode_oms.selectSingleNode("txt_117").Text
        oms_txt_118 = oNode_oms.selectSingleNode("txt_118").Text
        oms_txt_119 = oNode_oms.selectSingleNode("txt_119").Text
        oms_txt_120 = oNode_oms.selectSingleNode("txt_120").Text
        oms_txt_121 = oNode_oms.selectSingleNode("txt_121").Text
        oms_txt_122 = oNode_oms.selectSingleNode("txt_122").Text
        oms_txt_123 = oNode_oms.selectSingleNode("txt_123").Text
        oms_txt_124 = oNode_oms.selectSingleNode("txt_124").Text
        oms_txt_125 = oNode_oms.selectSingleNode("txt_125").Text
        oms_txt_126 = oNode_oms.selectSingleNode("txt_126").Text
        oms_txt_127 = oNode_oms.selectSingleNode("txt_127").Text
        oms_txt_128 = oNode_oms.selectSingleNode("txt_128").Text
        oms_txt_129 = oNode_oms.selectSingleNode("txt_129").Text
        oms_txt_130 = oNode_oms.selectSingleNode("txt_130").Text
        oms_txt_131 = oNode_oms.selectSingleNode("txt_131").Text
        oms_txt_132 = oNode_oms.selectSingleNode("txt_132").Text
        oms_txt_133 = oNode_oms.selectSingleNode("txt_133").Text
        oms_txt_134 = oNode_oms.selectSingleNode("txt_134").Text
        oms_txt_135 = oNode_oms.selectSingleNode("txt_135").Text
        oms_txt_136 = oNode_oms.selectSingleNode("txt_136").Text
        oms_txt_137 = oNode_oms.selectSingleNode("txt_137").Text
        oms_txt_138 = oNode_oms.selectSingleNode("txt_138").Text
        oms_txt_139 = oNode_oms.selectSingleNode("txt_139").Text
        oms_txt_140 = oNode_oms.selectSingleNode("txt_140").Text



  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>