
<% 
Dim objXMLHTTP_expence, objXMLDOM_expence, i_expence, strHTML_expence

Set objXMLDom_expence = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_expence = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_expence.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/expence_sprog.xml", False
'objXmlHttp_expence.open "GET", "http://localhost/inc/xml/expence_sprog.xml", False
'objXmlHttp_expence.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/expence_sprog.xml", False
'objXmlHttp_expence.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/expence_sprog.xml", False
'objXmlHttp_expence.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/expence_sprog.xml", False
'objXmlHttp_expence.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/expence_sprog.xml", False
objXmlHttp_expence.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/expence_sprog.xml", False

objXmlHttp_expence.send


Set objXmlDom_expence = objXmlHttp_expence.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_expence = Nothing



Dim Address_expence, Latitude_expence, Longitude_expence
Dim oNode_expence, oNodes_expence
Dim sXPathQuery_expence

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
sXPathQuery_expence = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_expence = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_expence = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_expence = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_expence = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_expence = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_expence = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_expence = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle bel�b og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_expence = objXmlDom_expence.documentElement.selectSingleNode(sXPathQuery_expence)
Address_expence = oNode_expence.Text

Set oNodes_expence = objXmlDom_expence.documentElement.selectNodes(sXPathQuery_expence)

    For Each oNode_expence in oNodes_expence

        expence_txt_001 = oNode_expence.selectSingleNode("txt_1").Text 
        expence_txt_002 = oNode_expence.selectSingleNode("txt_2").Text 
        expence_txt_003 = oNode_expence.selectSingleNode("txt_3").Text 
        expence_txt_004 = oNode_expence.selectSingleNode("txt_4").Text 
        expence_txt_005 = oNode_expence.selectSingleNode("txt_5").Text 
        expence_txt_006 = oNode_expence.selectSingleNode("txt_6").Text 
        expence_txt_007 = oNode_expence.selectSingleNode("txt_7").Text 
        expence_txt_008 = oNode_expence.selectSingleNode("txt_8").Text 
        expence_txt_009 = oNode_expence.selectSingleNode("txt_9").Text 
        expence_txt_0010 = oNode_expence.selectSingleNode("txt_10").Text 

        expence_txt_011 = oNode_expence.selectSingleNode("txt_11").Text 
        expence_txt_012 = oNode_expence.selectSingleNode("txt_12").Text 
        expence_txt_013 = oNode_expence.selectSingleNode("txt_13").Text 
        expence_txt_014 = oNode_expence.selectSingleNode("txt_14").Text 
        expence_txt_015 = oNode_expence.selectSingleNode("txt_15").Text 
        expence_txt_016 = oNode_expence.selectSingleNode("txt_16").Text 
        expence_txt_017 = oNode_expence.selectSingleNode("txt_17").Text 
        expence_txt_018 = oNode_expence.selectSingleNode("txt_18").Text 
        expence_txt_019 = oNode_expence.selectSingleNode("txt_19").Text 
        expence_txt_020 = oNode_expence.selectSingleNode("txt_20").Text 

        expence_txt_021 = oNode_expence.selectSingleNode("txt_21").Text 
        expence_txt_022 = oNode_expence.selectSingleNode("txt_22").Text 
        expence_txt_023 = oNode_expence.selectSingleNode("txt_23").Text 
        expence_txt_024 = oNode_expence.selectSingleNode("txt_24").Text 
        expence_txt_025 = oNode_expence.selectSingleNode("txt_25").Text 
        expence_txt_026 = oNode_expence.selectSingleNode("txt_26").Text 
        expence_txt_027 = oNode_expence.selectSingleNode("txt_27").Text 
        expence_txt_028 = oNode_expence.selectSingleNode("txt_28").Text 
        expence_txt_029 = oNode_expence.selectSingleNode("txt_29").Text 
        expence_txt_030 = oNode_expence.selectSingleNode("txt_30").Text 

        expence_txt_031 = oNode_expence.selectSingleNode("txt_31").Text 
        expence_txt_032 = oNode_expence.selectSingleNode("txt_32").Text 
        expence_txt_033 = oNode_expence.selectSingleNode("txt_33").Text 
        expence_txt_034 = oNode_expence.selectSingleNode("txt_34").Text 
        expence_txt_035 = oNode_expence.selectSingleNode("txt_35").Text 
        expence_txt_036 = oNode_expence.selectSingleNode("txt_36").Text 
        expence_txt_037 = oNode_expence.selectSingleNode("txt_37").Text 
        expence_txt_038 = oNode_expence.selectSingleNode("txt_38").Text 
        expence_txt_039 = oNode_expence.selectSingleNode("txt_39").Text 
        expence_txt_040 = oNode_expence.selectSingleNode("txt_40").Text 

        expence_txt_041 = oNode_expence.selectSingleNode("txt_41").Text 
        expence_txt_042 = oNode_expence.selectSingleNode("txt_42").Text 
        expence_txt_043 = oNode_expence.selectSingleNode("txt_43").Text 
        expence_txt_044 = oNode_expence.selectSingleNode("txt_44").Text 
        expence_txt_045 = oNode_expence.selectSingleNode("txt_45").Text 
        expence_txt_046 = oNode_expence.selectSingleNode("txt_46").Text 
        expence_txt_047 = oNode_expence.selectSingleNode("txt_47").Text 
        expence_txt_048 = oNode_expence.selectSingleNode("txt_48").Text 
        expence_txt_049 = oNode_expence.selectSingleNode("txt_49").Text 
        expence_txt_050 = oNode_expence.selectSingleNode("txt_50").Text 

        expence_txt_051 = oNode_expence.selectSingleNode("txt_51").Text 
        expence_txt_052 = oNode_expence.selectSingleNode("txt_52").Text 
        expence_txt_053 = oNode_expence.selectSingleNode("txt_53").Text 
        expence_txt_054 = oNode_expence.selectSingleNode("txt_54").Text 
        expence_txt_055 = oNode_expence.selectSingleNode("txt_55").Text 
        expence_txt_056 = oNode_expence.selectSingleNode("txt_56").Text 
        expence_txt_057 = oNode_expence.selectSingleNode("txt_57").Text 
        expence_txt_058 = oNode_expence.selectSingleNode("txt_58").Text 
        expence_txt_059 = oNode_expence.selectSingleNode("txt_59").Text 
        expence_txt_060 = oNode_expence.selectSingleNode("txt_60").Text 

        expence_txt_061 = oNode_expence.selectSingleNode("txt_61").Text 
        expence_txt_062 = oNode_expence.selectSingleNode("txt_62").Text 
        expence_txt_063 = oNode_expence.selectSingleNode("txt_63").Text 
        expence_txt_064 = oNode_expence.selectSingleNode("txt_64").Text 
        expence_txt_065 = oNode_expence.selectSingleNode("txt_65").Text 
        expence_txt_066 = oNode_expence.selectSingleNode("txt_66").Text 
        expence_txt_067 = oNode_expence.selectSingleNode("txt_67").Text 
        expence_txt_068 = oNode_expence.selectSingleNode("txt_68").Text 
        expence_txt_069 = oNode_expence.selectSingleNode("txt_69").Text 
        expence_txt_070 = oNode_expence.selectSingleNode("txt_70").Text 

        expence_txt_071 = oNode_expence.selectSingleNode("txt_71").Text 
        expence_txt_072 = oNode_expence.selectSingleNode("txt_72").Text 
        expence_txt_073 = oNode_expence.selectSingleNode("txt_73").Text 
        expence_txt_074 = oNode_expence.selectSingleNode("txt_74").Text 
        expence_txt_075 = oNode_expence.selectSingleNode("txt_75").Text 
        expence_txt_076 = oNode_expence.selectSingleNode("txt_76").Text 
        expence_txt_077 = oNode_expence.selectSingleNode("txt_77").Text 
        expence_txt_078 = oNode_expence.selectSingleNode("txt_78").Text 
        expence_txt_079 = oNode_expence.selectSingleNode("txt_79").Text 
        expence_txt_080 = oNode_expence.selectSingleNode("txt_80").Text 
        expence_txt_081 = oNode_expence.selectSingleNode("txt_81").Text

        expence_txt_082 = oNode_expence.selectSingleNode("txt_82").Text
        expence_txt_083 = oNode_expence.selectSingleNode("txt_83").Text
        expence_txt_084 = oNode_expence.selectSingleNode("txt_84").Text
        expence_txt_085 = oNode_expence.selectSingleNode("txt_85").Text
        expence_txt_086 = oNode_expence.selectSingleNode("txt_86").Text
        expence_txt_087 = oNode_expence.selectSingleNode("txt_87").Text
        expence_txt_088 = oNode_expence.selectSingleNode("txt_88").Text
        expence_txt_089 = oNode_expence.selectSingleNode("txt_89").Text
        expence_txt_090 = oNode_expence.selectSingleNode("txt_90").Text
        expence_txt_091 = oNode_expence.selectSingleNode("txt_91").Text
        expence_txt_092 = oNode_expence.selectSingleNode("txt_92").Text
        expence_txt_093 = oNode_expence.selectSingleNode("txt_93").Text
        expence_txt_094 = oNode_expence.selectSingleNode("txt_94").Text
        expence_txt_095 = oNode_expence.selectSingleNode("txt_95").Text


          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>