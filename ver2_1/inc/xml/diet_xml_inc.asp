
<% 
Dim objXMLHTTP_diet, objXMLDOM_diet, i_diet, strHTML_diet

Set objXMLDom_diet = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_diet = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_diet.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/diet_sprog.xml", False
'objXmlHttp_diet.open "GET", "http://localhost/inc/xml/diet_sprog.xml", False
'objXmlHttp_diet.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/diet_sprog.xml", False
'objXmlHttp_diet.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/diet_sprog.xml", False
'objXmlHttp_diet.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/diet_sprog.xml", False
'objXmlHttp_diet.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/diet_sprog.xml", False
objXmlHttp_diet.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/diet_sprog.xml", False

objXmlHttp_diet.send 


Set objXmlDom_diet = objXmlHttp_diet.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")awd


Set objXmlHttp_diet = Nothing



Dim Address_diet, Latitude_diet, Longitude_diet
Dim oNode_diet, oNodes_diet
Dim sXPathQuery_diet

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

    case "tia"
    sXPathQuery_diet = "//sprog/tia"
    
    case else
        select case sprog
        case 1
        sXPathQuery_diet = "//sprog/dk"
        'Session.LCID = 1030
        case 2
        sXPathQuery_diet = "//sprog/uk"
        'Session.LCID = 2057
        case 3
        sXPathQuery_diet = "//sprog/se"
        'Session.LCID = 1053
        case 4
        sXPathQuery_diet = "//sprog/no"
        'Session.LCID = 2068
        case 5
        sXPathQuery_diet = "//sprog/es"
        'Session.LCID = 1034
        case 6
        sXPathQuery_diet = "//sprog/de"
        'Session.LCID = 1031
        case 7
        sXPathQuery_diet = "//sprog/fr"
        'Session.LCID = 1036
        case else
        sXPathQuery_diet = "//sprog/dk"
        'Session.LCID = 1030
        end select
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_diet = objXmlDom_diet.documentElement.selectSingleNode(sXPathQuery_diet)
Address_diet = oNode_diet.Text

Set oNodes_diet = objXmlDom_diet.documentElement.selectNodes(sXPathQuery_diet)

    For Each oNode_diet in oNodes_diet

        diet_txt_001 = oNode_diet.selectSingleNode("txt_1").Text
        diet_txt_002 = oNode_diet.selectSingleNode("txt_2").Text
        diet_txt_003 = oNode_diet.selectSingleNode("txt_3").Text
        diet_txt_003 = oNode_diet.selectSingleNode("txt_3").Text
        diet_txt_004 = oNode_diet.selectSingleNode("txt_4").Text
        diet_txt_005 = oNode_diet.selectSingleNode("txt_5").Text

        diet_txt_006 = oNode_diet.selectSingleNode("txt_6").Text
        diet_txt_007 = oNode_diet.selectSingleNode("txt_7").Text
        diet_txt_008 = oNode_diet.selectSingleNode("txt_8").Text
        diet_txt_009 = oNode_diet.selectSingleNode("txt_9").Text
        diet_txt_010 = oNode_diet.selectSingleNode("txt_10").Text
        diet_txt_011 = oNode_diet.selectSingleNode("txt_11").Text
    
        diet_txt_012 = oNode_diet.selectSingleNode("txt_12").Text
        diet_txt_013 = oNode_diet.selectSingleNode("txt_13").Text
        diet_txt_014 = oNode_diet.selectSingleNode("txt_14").Text
        diet_txt_015 = oNode_diet.selectSingleNode("txt_15").Text
        diet_txt_016 = oNode_diet.selectSingleNode("txt_16").Text
        diet_txt_017 = oNode_diet.selectSingleNode("txt_17").Text
    
        diet_txt_018 = oNode_diet.selectSingleNode("txt_18").Text
        diet_txt_019 = oNode_diet.selectSingleNode("txt_19").Text
        diet_txt_020 = oNode_diet.selectSingleNode("txt_20").Text
        diet_txt_021 = oNode_diet.selectSingleNode("txt_21").Text
        diet_txt_022 = oNode_diet.selectSingleNode("txt_22").Text
        diet_txt_023 = oNode_diet.selectSingleNode("txt_23").Text
        diet_txt_024 = oNode_diet.selectSingleNode("txt_24").Text
    
        diet_txt_025 = oNode_diet.selectSingleNode("txt_25").Text
        diet_txt_026 = oNode_diet.selectSingleNode("txt_26").Text
        diet_txt_027 = oNode_diet.selectSingleNode("txt_27").Text
        diet_txt_028 = oNode_diet.selectSingleNode("txt_28").Text
        diet_txt_029 = oNode_diet.selectSingleNode("txt_29").Text
        diet_txt_030 = oNode_diet.selectSingleNode("txt_30").Text
        diet_txt_031 = oNode_diet.selectSingleNode("txt_31").Text
        diet_txt_032 = oNode_diet.selectSingleNode("txt_32").Text
        diet_txt_033 = oNode_diet.selectSingleNode("txt_33").Text
        diet_txt_034 = oNode_diet.selectSingleNode("txt_34").Text
    
        diet_txt_035 = oNode_diet.selectSingleNode("txt_35").Text
        diet_txt_036 = oNode_diet.selectSingleNode("txt_36").Text
        diet_txt_037 = oNode_diet.selectSingleNode("txt_37").Text
       
        diet_txt_038 = oNode_diet.selectSingleNode("txt_38").Text
        diet_txt_039 = oNode_diet.selectSingleNode("txt_39").Text
        diet_txt_040 = oNode_diet.selectSingleNode("txt_40").Text
        diet_txt_041 = oNode_diet.selectSingleNode("txt_41").Text
        diet_txt_042 = oNode_diet.selectSingleNode("txt_42").Text
        diet_txt_043 = oNode_diet.selectSingleNode("txt_43").Text
        diet_txt_044 = oNode_diet.selectSingleNode("txt_44").Text
        diet_txt_045 = oNode_diet.selectSingleNode("txt_45").Text
        diet_txt_046 = oNode_diet.selectSingleNode("txt_46").Text
        diet_txt_047 = oNode_diet.selectSingleNode("txt_47").Text
        diet_txt_048 = oNode_diet.selectSingleNode("txt_48").Text
        diet_txt_049 = oNode_diet.selectSingleNode("txt_49").Text
        diet_txt_050 = oNode_diet.selectSingleNode("txt_50").Text
        diet_txt_051 = oNode_diet.selectSingleNode("txt_51").Text
        diet_txt_052 = oNode_diet.selectSingleNode("txt_52").Text
        diet_txt_053 = oNode_diet.selectSingleNode("txt_53").Text
        diet_txt_054 = oNode_diet.selectSingleNode("txt_54").Text
        diet_txt_055 = oNode_diet.selectSingleNode("txt_55").Text
        diet_txt_056 = oNode_diet.selectSingleNode("txt_56").Text
        diet_txt_057 = oNode_diet.selectSingleNode("txt_57").Text
        diet_txt_058 = oNode_diet.selectSingleNode("txt_58").Text
         
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>