
<% 
Dim objXMLHTTP_medarbtyp, objXMLDOM_medarbtyp, i_medarbtyp, strHTML_medarbtyp

Set objXMLDom_medarbtyp = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_medarbtyp = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_medarbtyp.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/medarbtyper_sprog.xml", False
'objXmlHttp_medarbtyp.open "GET", "http://localhost/inc/xml/medarbtyper_sprog.xml", False
'objXmlHttp_medarbtyp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/medarbtyper_sprog.xml", False
'objXmlHttp_medarbtyp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/medarbtyper_sprog.xml", False
'objXmlHttp_medarbtyp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/medarbtyper_sprog.xml", False
objXmlHttp_medarbtyp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/medarbtyper_sprog.xml", False

objXmlHttp_medarbtyp.send


Set objXmlDom_medarbtyp = objXmlHttp_medarbtyp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_medarbtyp = Nothing



Dim Address_medarbtyp, Latitude_medarbtyp, Longitude_medarbtyp
Dim oNode_medarbtyp, oNodes_medarbtyp
Dim sXPathQuery_medarbtyp

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
sXPathQuery_medarbtyp = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_medarbtyp = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_medarbtyp = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_medarbtyp = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_medarbtyp = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_medarbtyp = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_medarbtyp = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_medarbtyp = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_medarbtyp = objXmlDom_medarbtyp.documentElement.selectSingleNode(sXPathQuery_medarbtyp)
Address_medarbtyp = oNode_medarbtyp.Text

Set oNodes_medarbtyp = objXmlDom_medarbtyp.documentElement.selectNodes(sXPathQuery_medarbtyp)

    For Each oNode_medarbtyp in oNodes_medarbtyp

          
        medarbtyp_txt_001 = oNode_medarbtyp.selectSingleNode("txt_1").Text
        medarbtyp_txt_002 = oNode_medarbtyp.selectSingleNode("txt_2").Text
        medarbtyp_txt_003 = oNode_medarbtyp.selectSingleNode("txt_3").Text
        medarbtyp_txt_003 = oNode_medarbtyp.selectSingleNode("txt_3").Text
        medarbtyp_txt_004 = oNode_medarbtyp.selectSingleNode("txt_4").Text
        medarbtyp_txt_005 = oNode_medarbtyp.selectSingleNode("txt_5").Text

        medarbtyp_txt_006 = oNode_medarbtyp.selectSingleNode("txt_6").Text
        medarbtyp_txt_007 = oNode_medarbtyp.selectSingleNode("txt_7").Text
        medarbtyp_txt_008 = oNode_medarbtyp.selectSingleNode("txt_8").Text
        medarbtyp_txt_009 = oNode_medarbtyp.selectSingleNode("txt_9").Text
        medarbtyp_txt_010 = oNode_medarbtyp.selectSingleNode("txt_10").Text
        medarbtyp_txt_011 = oNode_medarbtyp.selectSingleNode("txt_11").Text
    
        medarbtyp_txt_012 = oNode_medarbtyp.selectSingleNode("txt_12").Text
        medarbtyp_txt_013 = oNode_medarbtyp.selectSingleNode("txt_13").Text
        medarbtyp_txt_014 = oNode_medarbtyp.selectSingleNode("txt_14").Text
        medarbtyp_txt_015 = oNode_medarbtyp.selectSingleNode("txt_15").Text
        medarbtyp_txt_016 = oNode_medarbtyp.selectSingleNode("txt_16").Text
        medarbtyp_txt_017 = oNode_medarbtyp.selectSingleNode("txt_17").Text
    
        medarbtyp_txt_018 = oNode_medarbtyp.selectSingleNode("txt_18").Text
        medarbtyp_txt_019 = oNode_medarbtyp.selectSingleNode("txt_19").Text
        medarbtyp_txt_020 = oNode_medarbtyp.selectSingleNode("txt_20").Text
        medarbtyp_txt_021 = oNode_medarbtyp.selectSingleNode("txt_21").Text
        medarbtyp_txt_022 = oNode_medarbtyp.selectSingleNode("txt_22").Text
        medarbtyp_txt_023 = oNode_medarbtyp.selectSingleNode("txt_23").Text
        medarbtyp_txt_024 = oNode_medarbtyp.selectSingleNode("txt_24").Text
    
        medarbtyp_txt_025 = oNode_medarbtyp.selectSingleNode("txt_25").Text
        medarbtyp_txt_026 = oNode_medarbtyp.selectSingleNode("txt_26").Text
        medarbtyp_txt_027 = oNode_medarbtyp.selectSingleNode("txt_27").Text
        medarbtyp_txt_028 = oNode_medarbtyp.selectSingleNode("txt_28").Text
        medarbtyp_txt_029 = oNode_medarbtyp.selectSingleNode("txt_29").Text
        medarbtyp_txt_030 = oNode_medarbtyp.selectSingleNode("txt_30").Text
        medarbtyp_txt_031 = oNode_medarbtyp.selectSingleNode("txt_31").Text
        medarbtyp_txt_032 = oNode_medarbtyp.selectSingleNode("txt_32").Text
        medarbtyp_txt_033 = oNode_medarbtyp.selectSingleNode("txt_33").Text
        medarbtyp_txt_034 = oNode_medarbtyp.selectSingleNode("txt_34").Text
    
        medarbtyp_txt_035 = oNode_medarbtyp.selectSingleNode("txt_35").Text
        medarbtyp_txt_036 = oNode_medarbtyp.selectSingleNode("txt_36").Text
        medarbtyp_txt_037 = oNode_medarbtyp.selectSingleNode("txt_37").Text
       
        medarbtyp_txt_038 = oNode_medarbtyp.selectSingleNode("txt_38").Text
        medarbtyp_txt_039 = oNode_medarbtyp.selectSingleNode("txt_39").Text
        medarbtyp_txt_040 = oNode_medarbtyp.selectSingleNode("txt_40").Text
        medarbtyp_txt_041 = oNode_medarbtyp.selectSingleNode("txt_41").Text
        medarbtyp_txt_042 = oNode_medarbtyp.selectSingleNode("txt_42").Text
        medarbtyp_txt_043 = oNode_medarbtyp.selectSingleNode("txt_43").Text
        medarbtyp_txt_044 = oNode_medarbtyp.selectSingleNode("txt_44").Text
          
        medarbtyp_txt_045 = oNode_medarbtyp.selectSingleNode("txt_45").Text
        medarbtyp_txt_046 = oNode_medarbtyp.selectSingleNode("txt_46").Text
        medarbtyp_txt_047 = oNode_medarbtyp.selectSingleNode("txt_47").Text
        medarbtyp_txt_048 = oNode_medarbtyp.selectSingleNode("txt_48").Text
        medarbtyp_txt_049 = oNode_medarbtyp.selectSingleNode("txt_49").Text
        medarbtyp_txt_050 = oNode_medarbtyp.selectSingleNode("txt_50").Text
          
        medarbtyp_txt_051 = oNode_medarbtyp.selectSingleNode("txt_51").Text
        medarbtyp_txt_052 = oNode_medarbtyp.selectSingleNode("txt_52").Text
        medarbtyp_txt_053 = oNode_medarbtyp.selectSingleNode("txt_53").Text
        medarbtyp_txt_054 = oNode_medarbtyp.selectSingleNode("txt_54").Text
        medarbtyp_txt_055 = oNode_medarbtyp.selectSingleNode("txt_55").Text
        medarbtyp_txt_056 = oNode_medarbtyp.selectSingleNode("txt_56").Text
        medarbtyp_txt_057 = oNode_medarbtyp.selectSingleNode("txt_57").Text
        medarbtyp_txt_058 = oNode_medarbtyp.selectSingleNode("txt_58").Text
        medarbtyp_txt_059 = oNode_medarbtyp.selectSingleNode("txt_59").Text
        medarbtyp_txt_060 = oNode_medarbtyp.selectSingleNode("txt_60").Text
        medarbtyp_txt_061 = oNode_medarbtyp.selectSingleNode("txt_61").Text
        medarbtyp_txt_062 = oNode_medarbtyp.selectSingleNode("txt_62").Text
        medarbtyp_txt_063 = oNode_medarbtyp.selectSingleNode("txt_63").Text
        medarbtyp_txt_064 = oNode_medarbtyp.selectSingleNode("txt_64").Text
        medarbtyp_txt_065 = oNode_medarbtyp.selectSingleNode("txt_65").Text
        medarbtyp_txt_066 = oNode_medarbtyp.selectSingleNode("txt_66").Text
          
        medarbtyp_txt_067 = oNode_medarbtyp.selectSingleNode("txt_67").Text
        medarbtyp_txt_068 = oNode_medarbtyp.selectSingleNode("txt_68").Text
        medarbtyp_txt_069 = oNode_medarbtyp.selectSingleNode("txt_69").Text
        medarbtyp_txt_070 = oNode_medarbtyp.selectSingleNode("txt_70").Text

        medarbtyp_txt_071 = oNode_medarbtyp.selectSingleNode("txt_71").Text
        medarbtyp_txt_072 = oNode_medarbtyp.selectSingleNode("txt_72").Text
        medarbtyp_txt_073 = oNode_medarbtyp.selectSingleNode("txt_73").Text
        medarbtyp_txt_074 = oNode_medarbtyp.selectSingleNode("txt_74").Text
        medarbtyp_txt_075 = oNode_medarbtyp.selectSingleNode("txt_75").Text
        medarbtyp_txt_076 = oNode_medarbtyp.selectSingleNode("txt_76").Text
        medarbtyp_txt_077 = oNode_medarbtyp.selectSingleNode("txt_77").Text
        medarbtyp_txt_078 = oNode_medarbtyp.selectSingleNode("txt_78").Text
        medarbtyp_txt_079 = oNode_medarbtyp.selectSingleNode("txt_79").Text

        medarbtyp_txt_080 = oNode_medarbtyp.selectSingleNode("txt_80").Text
        medarbtyp_txt_081 = oNode_medarbtyp.selectSingleNode("txt_81").Text
        medarbtyp_txt_082 = oNode_medarbtyp.selectSingleNode("txt_82").Text
        medarbtyp_txt_083 = oNode_medarbtyp.selectSingleNode("txt_83").Text
          
        medarbtyp_txt_084 = oNode_medarbtyp.selectSingleNode("txt_84").Text
        medarbtyp_txt_085 = oNode_medarbtyp.selectSingleNode("txt_85").Text
        medarbtyp_txt_086 = oNode_medarbtyp.selectSingleNode("txt_86").Text
          
        medarbtyp_txt_087 = oNode_medarbtyp.selectSingleNode("txt_87").Text
          
        medarbtyp_txt_088 = oNode_medarbtyp.selectSingleNode("txt_88").Text
        medarbtyp_txt_089 = oNode_medarbtyp.selectSingleNode("txt_89").Text
        medarbtyp_txt_090 = oNode_medarbtyp.selectSingleNode("txt_90").Text
          
        medarbtyp_txt_091 = oNode_medarbtyp.selectSingleNode("txt_91").Text
        medarbtyp_txt_092 = oNode_medarbtyp.selectSingleNode("txt_92").Text
        medarbtyp_txt_093 = oNode_medarbtyp.selectSingleNode("txt_93").Text
        medarbtyp_txt_094 = oNode_medarbtyp.selectSingleNode("txt_94").Text
        medarbtyp_txt_095 = oNode_medarbtyp.selectSingleNode("txt_95").Text
        medarbtyp_txt_096 = oNode_medarbtyp.selectSingleNode("txt_96").Text
        medarbtyp_txt_097 = oNode_medarbtyp.selectSingleNode("txt_97").Text
        medarbtyp_txt_098 = oNode_medarbtyp.selectSingleNode("txt_98").Text
        medarbtyp_txt_099 = oNode_medarbtyp.selectSingleNode("txt_99").Text
          
        medarbtyp_txt_100 = oNode_medarbtyp.selectSingleNode("txt_100").Text
        medarbtyp_txt_101 = oNode_medarbtyp.selectSingleNode("txt_101").Text
        medarbtyp_txt_102 = oNode_medarbtyp.selectSingleNode("txt_102").Text
        medarbtyp_txt_103 = oNode_medarbtyp.selectSingleNode("txt_103").Text
          
        medarbtyp_txt_104 = oNode_medarbtyp.selectSingleNode("txt_104").Text
        medarbtyp_txt_105 = oNode_medarbtyp.selectSingleNode("txt_105").Text
        medarbtyp_txt_106 = oNode_medarbtyp.selectSingleNode("txt_106").Text
        medarbtyp_txt_107 = oNode_medarbtyp.selectSingleNode("txt_107").Text
        medarbtyp_txt_108 = oNode_medarbtyp.selectSingleNode("txt_108").Text
        medarbtyp_txt_109 = oNode_medarbtyp.selectSingleNode("txt_109").Text
        medarbtyp_txt_110 = oNode_medarbtyp.selectSingleNode("txt_110").Text

        medarbtyp_txt_111 = oNode_medarbtyp.selectSingleNode("txt_111").Text
        medarbtyp_txt_112 = oNode_medarbtyp.selectSingleNode("txt_112").Text
        medarbtyp_txt_113 = oNode_medarbtyp.selectSingleNode("txt_113").Text
        medarbtyp_txt_114 = oNode_medarbtyp.selectSingleNode("txt_114").Text
        medarbtyp_txt_115 = oNode_medarbtyp.selectSingleNode("txt_115").Text
        medarbtyp_txt_116 = oNode_medarbtyp.selectSingleNode("txt_116").Text
        medarbtyp_txt_117 = oNode_medarbtyp.selectSingleNode("txt_117").Text

        medarbtyp_txt_118 = oNode_medarbtyp.selectSingleNode("txt_118").Text
        medarbtyp_txt_119 = oNode_medarbtyp.selectSingleNode("txt_119").Text
        medarbtyp_txt_120 = oNode_medarbtyp.selectSingleNode("txt_120").Text
        medarbtyp_txt_121 = oNode_medarbtyp.selectSingleNode("txt_121").Text
        medarbtyp_txt_122 = oNode_medarbtyp.selectSingleNode("txt_122").Text
        medarbtyp_txt_123 = oNode_medarbtyp.selectSingleNode("txt_123").Text
        medarbtyp_txt_124 = oNode_medarbtyp.selectSingleNode("txt_124").Text
          
    next




'Response.Write "week_txt_001: " & week_txt_001 & "<br>"
'Response.Write "week_txt_002: " & week_txt_002 & "<br>"


%>