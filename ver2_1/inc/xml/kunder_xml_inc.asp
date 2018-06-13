
<% 
Dim objXMLHTTP_kunder, objXMLDOM_kunder, i_kunder, strHTML_kunder

Set objXMLDom_kunder = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_kunder = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_kunder.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/kunder_sprog.xml", False
objXmlHttp_kunder.open "GET", "http://localhost/inc/xml/kunder_sprog.xml", False
'objXmlHttp_kunder.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/kunder_sprog.xml", False
'objXmlHttp_kunder.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/kunder_sprog.xml", False
'objXmlHttp_kunder.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/kunder_sprog.xml", False
'objXmlHttp_kunder.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/kunder_sprog.xml", False

objXmlHttp_kunder.send


Set objXmlDom_kunder = objXmlHttp_kunder.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_kunder = Nothing



Dim Address_kunder, Latitude_kunder, Longitude_kunder
Dim oNode_kunder, oNodes_kunder
Dim sXPathQuery_kunder

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
sXPathQuery_kunder = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_kunder = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_kunder = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_kunder = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_kunder = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_kunder = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_kunder = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_kunder = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_kunder = objXmlDom_kunder.documentElement.selectSingleNode(sXPathQuery_kunder)
Address_kunder = oNode_kunder.Text

Set oNodes_kunder = objXmlDom_kunder.documentElement.selectNodes(sXPathQuery_kunder)

    For Each oNode_kunder in oNodes_kunder

        kunder_txt_001 = oNode_kunder.selectSingleNode("txt_1").Text
        kunder_txt_002 = oNode_kunder.selectSingleNode("txt_2").Text
        kunder_txt_003 = oNode_kunder.selectSingleNode("txt_3").Text
        kunder_txt_004 = oNode_kunder.selectSingleNode("txt_4").Text
        kunder_txt_005 = oNode_kunder.selectSingleNode("txt_5").Text
        kunder_txt_006 = oNode_kunder.selectSingleNode("txt_6").Text
        kunder_txt_007 = oNode_kunder.selectSingleNode("txt_7").Text
        kunder_txt_008 = oNode_kunder.selectSingleNode("txt_8").Text
        kunder_txt_009 = oNode_kunder.selectSingleNode("txt_9").Text
        kunder_txt_010 = oNode_kunder.selectSingleNode("txt_10").Text

        kunder_txt_011 = oNode_kunder.selectSingleNode("txt_11").Text
        kunder_txt_012 = oNode_kunder.selectSingleNode("txt_12").Text
        kunder_txt_013 = oNode_kunder.selectSingleNode("txt_13").Text
        kunder_txt_014 = oNode_kunder.selectSingleNode("txt_14").Text
        kunder_txt_015 = oNode_kunder.selectSingleNode("txt_15").Text
        kunder_txt_016 = oNode_kunder.selectSingleNode("txt_16").Text
        kunder_txt_017 = oNode_kunder.selectSingleNode("txt_17").Text
        kunder_txt_018 = oNode_kunder.selectSingleNode("txt_18").Text
        kunder_txt_019 = oNode_kunder.selectSingleNode("txt_19").Text
        kunder_txt_020 = oNode_kunder.selectSingleNode("txt_20").Text

        kunder_txt_021 = oNode_kunder.selectSingleNode("txt_21").Text
        kunder_txt_022 = oNode_kunder.selectSingleNode("txt_22").Text
        kunder_txt_023 = oNode_kunder.selectSingleNode("txt_23").Text
        kunder_txt_024 = oNode_kunder.selectSingleNode("txt_24").Text
        kunder_txt_025 = oNode_kunder.selectSingleNode("txt_25").Text
        kunder_txt_026 = oNode_kunder.selectSingleNode("txt_26").Text
        kunder_txt_027 = oNode_kunder.selectSingleNode("txt_27").Text
        kunder_txt_028 = oNode_kunder.selectSingleNode("txt_28").Text
        kunder_txt_029 = oNode_kunder.selectSingleNode("txt_29").Text
        kunder_txt_030 = oNode_kunder.selectSingleNode("txt_30").Text

        kunder_txt_031 = oNode_kunder.selectSingleNode("txt_31").Text
        kunder_txt_032 = oNode_kunder.selectSingleNode("txt_32").Text
        kunder_txt_033 = oNode_kunder.selectSingleNode("txt_33").Text
        kunder_txt_034 = oNode_kunder.selectSingleNode("txt_34").Text
        kunder_txt_035 = oNode_kunder.selectSingleNode("txt_35").Text
        kunder_txt_036 = oNode_kunder.selectSingleNode("txt_36").Text
        kunder_txt_037 = oNode_kunder.selectSingleNode("txt_37").Text
        kunder_txt_038 = oNode_kunder.selectSingleNode("txt_38").Text
        kunder_txt_039 = oNode_kunder.selectSingleNode("txt_39").Text
        kunder_txt_040 = oNode_kunder.selectSingleNode("txt_40").Text

        kunder_txt_041 = oNode_kunder.selectSingleNode("txt_41").Text
        kunder_txt_042 = oNode_kunder.selectSingleNode("txt_42").Text
        kunder_txt_043 = oNode_kunder.selectSingleNode("txt_43").Text
        kunder_txt_044 = oNode_kunder.selectSingleNode("txt_44").Text
        kunder_txt_045 = oNode_kunder.selectSingleNode("txt_45").Text
        kunder_txt_046 = oNode_kunder.selectSingleNode("txt_46").Text
        kunder_txt_047 = oNode_kunder.selectSingleNode("txt_47").Text
        kunder_txt_048 = oNode_kunder.selectSingleNode("txt_48").Text
        kunder_txt_049 = oNode_kunder.selectSingleNode("txt_49").Text
        kunder_txt_050 = oNode_kunder.selectSingleNode("txt_50").Text

        kunder_txt_051 = oNode_kunder.selectSingleNode("txt_51").Text
        kunder_txt_052 = oNode_kunder.selectSingleNode("txt_52").Text
        kunder_txt_053 = oNode_kunder.selectSingleNode("txt_53").Text
        kunder_txt_054 = oNode_kunder.selectSingleNode("txt_54").Text
        kunder_txt_055 = oNode_kunder.selectSingleNode("txt_55").Text
        kunder_txt_056 = oNode_kunder.selectSingleNode("txt_56").Text
        kunder_txt_057 = oNode_kunder.selectSingleNode("txt_57").Text
        kunder_txt_058 = oNode_kunder.selectSingleNode("txt_58").Text
        kunder_txt_059 = oNode_kunder.selectSingleNode("txt_59").Text
        kunder_txt_060 = oNode_kunder.selectSingleNode("txt_60").Text

        kunder_txt_061 = oNode_kunder.selectSingleNode("txt_61").Text
        kunder_txt_062 = oNode_kunder.selectSingleNode("txt_62").Text
        kunder_txt_063 = oNode_kunder.selectSingleNode("txt_63").Text
        kunder_txt_064 = oNode_kunder.selectSingleNode("txt_64").Text
        kunder_txt_065 = oNode_kunder.selectSingleNode("txt_65").Text
        kunder_txt_066 = oNode_kunder.selectSingleNode("txt_66").Text
        kunder_txt_067 = oNode_kunder.selectSingleNode("txt_67").Text
        kunder_txt_068 = oNode_kunder.selectSingleNode("txt_68").Text
        kunder_txt_069 = oNode_kunder.selectSingleNode("txt_69").Text
        kunder_txt_070 = oNode_kunder.selectSingleNode("txt_70").Text

        kunder_txt_071 = oNode_kunder.selectSingleNode("txt_71").Text
        kunder_txt_072 = oNode_kunder.selectSingleNode("txt_72").Text
        kunder_txt_073 = oNode_kunder.selectSingleNode("txt_73").Text
        kunder_txt_074 = oNode_kunder.selectSingleNode("txt_74").Text
        kunder_txt_075 = oNode_kunder.selectSingleNode("txt_75").Text
        kunder_txt_076 = oNode_kunder.selectSingleNode("txt_76").Text
        kunder_txt_077 = oNode_kunder.selectSingleNode("txt_77").Text
        kunder_txt_078 = oNode_kunder.selectSingleNode("txt_78").Text
        kunder_txt_079 = oNode_kunder.selectSingleNode("txt_79").Text
        kunder_txt_080 = oNode_kunder.selectSingleNode("txt_80").Text

        kunder_txt_081 = oNode_kunder.selectSingleNode("txt_81").Text
        kunder_txt_082 = oNode_kunder.selectSingleNode("txt_82").Text
        kunder_txt_083 = oNode_kunder.selectSingleNode("txt_83").Text
        kunder_txt_084 = oNode_kunder.selectSingleNode("txt_84").Text
        kunder_txt_085 = oNode_kunder.selectSingleNode("txt_85").Text
        kunder_txt_086 = oNode_kunder.selectSingleNode("txt_86").Text
        kunder_txt_087 = oNode_kunder.selectSingleNode("txt_87").Text
        kunder_txt_088 = oNode_kunder.selectSingleNode("txt_88").Text
        kunder_txt_089 = oNode_kunder.selectSingleNode("txt_89").Text
        kunder_txt_090 = oNode_kunder.selectSingleNode("txt_90").Text

        kunder_txt_091 = oNode_kunder.selectSingleNode("txt_91").Text
        kunder_txt_092 = oNode_kunder.selectSingleNode("txt_92").Text
        kunder_txt_093 = oNode_kunder.selectSingleNode("txt_93").Text
        kunder_txt_094 = oNode_kunder.selectSingleNode("txt_94").Text
        kunder_txt_095 = oNode_kunder.selectSingleNode("txt_95").Text
        kunder_txt_096 = oNode_kunder.selectSingleNode("txt_96").Text
        kunder_txt_097 = oNode_kunder.selectSingleNode("txt_97").Text
        kunder_txt_098 = oNode_kunder.selectSingleNode("txt_98").Text
        kunder_txt_099 = oNode_kunder.selectSingleNode("txt_99").Text
        kunder_txt_100 = oNode_kunder.selectSingleNode("txt_100").Text

        kunder_txt_101 = oNode_kunder.selectSingleNode("txt_101").Text
        kunder_txt_102 = oNode_kunder.selectSingleNode("txt_102").Text
        kunder_txt_103 = oNode_kunder.selectSingleNode("txt_103").Text
        kunder_txt_104 = oNode_kunder.selectSingleNode("txt_104").Text
        kunder_txt_105 = oNode_kunder.selectSingleNode("txt_105").Text
        kunder_txt_106 = oNode_kunder.selectSingleNode("txt_106").Text
        kunder_txt_107 = oNode_kunder.selectSingleNode("txt_107").Text
        kunder_txt_108 = oNode_kunder.selectSingleNode("txt_108").Text
        kunder_txt_109 = oNode_kunder.selectSingleNode("txt_109").Text
        kunder_txt_110 = oNode_kunder.selectSingleNode("txt_110").Text
        kunder_txt_111 = oNode_kunder.selectSingleNode("txt_111").Text
        kunder_txt_112 = oNode_kunder.selectSingleNode("txt_112").Text
        kunder_txt_113 = oNode_kunder.selectSingleNode("txt_113").Text
        kunder_txt_114 = oNode_kunder.selectSingleNode("txt_114").Text
        kunder_txt_115 = oNode_kunder.selectSingleNode("txt_115").Text
        kunder_txt_116 = oNode_kunder.selectSingleNode("txt_116").Text
        kunder_txt_117 = oNode_kunder.selectSingleNode("txt_117").Text
        kunder_txt_118 = oNode_kunder.selectSingleNode("txt_118").Text



  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>