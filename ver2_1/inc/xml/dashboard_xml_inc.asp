
<% 

Dim dsb_objXMLHTTP, dsb_objXMLDOM, dsb_i, dsb_strHTML

Set dsb_objXMLDom = Server.CreateObject("Microsoft.XMLDOM")
Set dsb_objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
dsb_objXmlHttp.open "GET", "http://localhost/Git/timeout_2016/ver2_1//inc/xml/dashboard_sprog.xml", False
'objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_1/inc/xml/dashboard_sprog.xml", False
'objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver3_99/inc/xml/dashboard_sprog.xml", False
'objXmlHttp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/dashboard_sprog.xml", False
'dsb_objXmlHttp.open "GET", "http://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/dashboard_sprog.xml", False
'objXmlHttp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver4_22/inc/xml/dashboard_sprog.xml", False

dsb_objXmlHttp.send


Set dsb_objXmlDom = dsb_objXmlHttp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set dsb_objXmlHttp = Nothing



Dim dsb_Address, dsb_Latitude, dsb_Longitude
Dim dsb_oNode, dsb_oNodes
Dim dsb_sXPathQuery


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
dsb_sXPathQuery = "//sprog/dk"
Session.LCID = 1030
case 2
dsb_sXPathQuery = "//sprog/uk"
Session.LCID = 2057
case 3
dsb_sXPathQuery = "//sprog/se"
Session.LCID = 1053
case else
dsb_sXPathQuery = "//sprog/dk"
Session.LCID = 1030
end select

'Response.Write "Session.LCID" &  Session.LCID

Set dsb_oNode = dsb_objXmlDom.documentElement.selectSingleNode(dsb_sXPathQuery)
dsb_Address = dsb_oNode.Text

Set dsb_oNodes = dsb_objXmlDom.documentElement.selectNodes(dsb_sXPathQuery)

For Each dsb_oNode in dsb_oNodes

     dsb_txt_001 = dsb_oNode.selectSingleNode("txt_1").Text
     dsb_txt_002 = dsb_oNode.selectSingleNode("txt_2").Text
     dsb_txt_003 = dsb_oNode.selectSingleNode("txt_3").Text
     dsb_txt_003 = dsb_oNode.selectSingleNode("txt_3").Text
    dsb_txt_004 = dsb_oNode.selectSingleNode("txt_4").Text
    dsb_txt_005 = dsb_oNode.selectSingleNode("txt_5").Text

    dsb_txt_006 = dsb_oNode.selectSingleNode("txt_6").Text
    dsb_txt_007 = dsb_oNode.selectSingleNode("txt_7").Text
    dsb_txt_008 = dsb_oNode.selectSingleNode("txt_8").Text
    dsb_txt_009 = dsb_oNode.selectSingleNode("txt_9").Text
    dsb_txt_010 = dsb_oNode.selectSingleNode("txt_10").Text
    dsb_txt_011 = dsb_oNode.selectSingleNode("txt_11").Text
    
    dsb_txt_012 = dsb_oNode.selectSingleNode("txt_12").Text
    dsb_txt_013 = dsb_oNode.selectSingleNode("txt_13").Text
    dsb_txt_014 = dsb_oNode.selectSingleNode("txt_14").Text
    dsb_txt_015 = dsb_oNode.selectSingleNode("txt_15").Text
    dsb_txt_016 = dsb_oNode.selectSingleNode("txt_16").Text
    dsb_txt_017 = dsb_oNode.selectSingleNode("txt_17").Text
    
    dsb_txt_018 = dsb_oNode.selectSingleNode("txt_18").Text
    dsb_txt_019 = dsb_oNode.selectSingleNode("txt_19").Text
    dsb_txt_020 = dsb_oNode.selectSingleNode("txt_20").Text
    dsb_txt_021 = dsb_oNode.selectSingleNode("txt_21").Text
    dsb_txt_022 = dsb_oNode.selectSingleNode("txt_22").Text
    dsb_txt_023 = dsb_oNode.selectSingleNode("txt_23").Text
    dsb_txt_024 = dsb_oNode.selectSingleNode("txt_24").Text
    
    dsb_txt_025 = dsb_oNode.selectSingleNode("txt_25").Text
    dsb_txt_026 = dsb_oNode.selectSingleNode("txt_26").Text
    dsb_txt_027 = dsb_oNode.selectSingleNode("txt_27").Text
    dsb_txt_028 = dsb_oNode.selectSingleNode("txt_28").Text
    dsb_txt_029 = dsb_oNode.selectSingleNode("txt_29").Text
    dsb_txt_030 = dsb_oNode.selectSingleNode("txt_30").Text
    dsb_txt_031 = dsb_oNode.selectSingleNode("txt_31").Text
    dsb_txt_032 = dsb_oNode.selectSingleNode("txt_32").Text
    dsb_txt_033 = dsb_oNode.selectSingleNode("txt_33").Text
    dsb_txt_034 = dsb_oNode.selectSingleNode("txt_34").Text
    
     dsb_txt_035 = dsb_oNode.selectSingleNode("txt_35").Text
      dsb_txt_036 = dsb_oNode.selectSingleNode("txt_36").Text
       dsb_txt_037 = dsb_oNode.selectSingleNode("txt_37").Text
       
        dsb_txt_038 = dsb_oNode.selectSingleNode("txt_38").Text
         dsb_txt_039 = dsb_oNode.selectSingleNode("txt_39").Text
         dsb_txt_040 = dsb_oNode.selectSingleNode("txt_40").Text
         dsb_txt_041 = dsb_oNode.selectSingleNode("txt_41").Text
         dsb_txt_042 = dsb_oNode.selectSingleNode("txt_42").Text
          dsb_txt_043 = dsb_oNode.selectSingleNode("txt_43").Text
          dsb_txt_044 = dsb_oNode.selectSingleNode("txt_44").Text
          
          dsb_txt_045 = dsb_oNode.selectSingleNode("txt_45").Text
          dsb_txt_046 = dsb_oNode.selectSingleNode("txt_46").Text
          dsb_txt_047 = dsb_oNode.selectSingleNode("txt_47").Text
          dsb_txt_048 = dsb_oNode.selectSingleNode("txt_48").Text
          dsb_txt_049 = dsb_oNode.selectSingleNode("txt_49").Text
          dsb_txt_050 = dsb_oNode.selectSingleNode("txt_50").Text


          dsb_txt_051 = dsb_oNode.selectSingleNode("txt_51").Text
          dsb_txt_052 = dsb_oNode.selectSingleNode("txt_52").Text
          dsb_txt_053 = dsb_oNode.selectSingleNode("txt_53").Text
          dsb_txt_054 = dsb_oNode.selectSingleNode("txt_54").Text
          dsb_txt_055 = dsb_oNode.selectSingleNode("txt_55").Text
          dsb_txt_056 = dsb_oNode.selectSingleNode("txt_56").Text
          dsb_txt_057 = dsb_oNode.selectSingleNode("txt_57").Text
          dsb_txt_058 = dsb_oNode.selectSingleNode("txt_58").Text
          dsb_txt_059 = dsb_oNode.selectSingleNode("txt_59").Text
          dsb_txt_060 = dsb_oNode.selectSingleNode("txt_60").Text
          dsb_txt_061 = dsb_oNode.selectSingleNode("txt_61").Text
          dsb_txt_062 = dsb_oNode.selectSingleNode("txt_62").Text
          dsb_txt_063 = dsb_oNode.selectSingleNode("txt_63").Text
          dsb_txt_064 = dsb_oNode.selectSingleNode("txt_64").Text
          dsb_txt_065 = dsb_oNode.selectSingleNode("txt_65").Text
          dsb_txt_066 = dsb_oNode.selectSingleNode("txt_66").Text
          
          dsb_txt_067 = dsb_oNode.selectSingleNode("txt_67").Text
          dsb_txt_068 = dsb_oNode.selectSingleNode("txt_68").Text
          dsb_txt_069 = dsb_oNode.selectSingleNode("txt_69").Text
          dsb_txt_070 = dsb_oNode.selectSingleNode("txt_70").Text
          dsb_txt_071 = dsb_oNode.selectSingleNode("txt_71").Text
          dsb_txt_072 = dsb_oNode.selectSingleNode("txt_72").Text
          dsb_txt_073 = dsb_oNode.selectSingleNode("txt_73").Text
          dsb_txt_074 = dsb_oNode.selectSingleNode("txt_74").Text
          dsb_txt_075 = dsb_oNode.selectSingleNode("txt_75").Text
          dsb_txt_076 = dsb_oNode.selectSingleNode("txt_76").Text
          dsb_txt_077 = dsb_oNode.selectSingleNode("txt_77").Text
          dsb_txt_078 = dsb_oNode.selectSingleNode("txt_78").Text
          dsb_txt_079 = dsb_oNode.selectSingleNode("txt_79").Text
        
          
          dsb_txt_080 = dsb_oNode.selectSingleNode("txt_80").Text
          dsb_txt_081 = dsb_oNode.selectSingleNode("txt_81").Text
          dsb_txt_082 = dsb_oNode.selectSingleNode("txt_82").Text
          dsb_txt_083 = dsb_oNode.selectSingleNode("txt_83").Text
          
          dsb_txt_084 = dsb_oNode.selectSingleNode("txt_84").Text
          dsb_txt_085 = dsb_oNode.selectSingleNode("txt_85").Text
          dsb_txt_086 = dsb_oNode.selectSingleNode("txt_86").Text
          
          dsb_txt_087 = dsb_oNode.selectSingleNode("txt_87").Text
          
          dsb_txt_088 = dsb_oNode.selectSingleNode("txt_88").Text
          dsb_txt_089 = dsb_oNode.selectSingleNode("txt_89").Text
          dsb_txt_090 = dsb_oNode.selectSingleNode("txt_90").Text
          
        
    
next




%>