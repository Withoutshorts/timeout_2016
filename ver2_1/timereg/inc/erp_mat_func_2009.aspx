<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    </div>
    </form>
    
    
    
    
    
    
</body>
</html>



 <table width=100% cellspacing=0 cellpadding=2 border=0>
	                        </tr>
	                        <tr bgcolor="orange">
	                        <%call matFelter %>
	                        </tr>
	                        <!-- Materialer -->
	                        <%
	                        '** Allerede fakturarede materialer 
	                        m = 0
	                        isMatWrt = " matid <> -1 "
                        	
	                        if func = "red" then
                        	
	                        
	                        strSQL = "SELECT matid, matnavn, matantal, matenhedspris AS matsalgspris, "_
	                        &" matenhed, matvarenr, ikkemoms, fms.valuta AS valutaid, fms.kurs, v.valutakode, matrabat, matfrb_mid, matfrb_id FROM fak_mat_spec fms "_
	                        &" LEFT JOIN valutaer v ON (v.id = fms.valuta) WHERE matfakid = "& id &" ORDER BY matnavn"
	                        'Response.Write strSQL
	                        oRec.open strSQL, oConn, 3
                            while not oRec.EOF
                            
                            intMatRabat = (100*oRec("matrabat"))
                            valutaMatId = oRec("valutaid")
                            valutaMatKode = oRec("valutakode") 
                            valutaMatKurs = oRec("kurs")
                            ikkemoms = oRec("ikkemoms")
                            call materialer(1)
                            
                            isMatWrt = isMatWrt & " AND matid <> " & oRec("matid")
                            m = m + 1
                            oRec.movenext
                            wend
                            oRec.close
                        	
                        	
	                        
                        	
	                        end if
                        	
                        	
	                        '** Ikke fakturerede Materialer / Eller v. ny faktura = alle 
	                        strSQL = "SELECT mf.matid, mf.matnavn, sum(mf.matantal) AS matantal, mf.matsalgspris, mf.matenhed, "_
	                        &" mf.matvarenr, mf.valuta, v.valutakode, v.kurs, mf.erfak FROM materiale_forbrug mf "_
	                        &" LEFT JOIN valutaer v ON (v.id = mf.valuta) WHERE mf.jobid = "& jobid &" AND "_
	                        &" forbrugsdato BETWEEN '"& stdatoKri &"' AND '"& slutdato &"'"_
	                        &" AND ("& isMatWrt &") GROUP BY mf.matid, mf.matsalgspris, mf.matnavn ORDER BY mf.matnavn"
	                        '
	                        'Response.Write strSQL
	                        'Response.flush
	                        
	                        oRec.open strSQL, oConn, 3
                            im = 0
                            while not oRec.EOF
                                
                                if im = 0 AND func = "red" then
                                
                                
                                
                                %>
	                             <tr><td colspan=9>
                                     &nbsp;</td></tr>
    	                       
	                            <tr><td colspan=9><font class=roed>Ikke fakturerede materialer, registreret i den valgte periode.</font></td></tr>
	                             <tr bgcolor="silver">
	                            <%call matFelter %>
	                        </tr>
	                            <%
                                end if
                            
                            
                            intMatRabat = intRabat
                            valutaMatId = oRec("valuta") 
                            valutaMatKode = oRec("valutakode") 
                            valutaMatKurs = oRec("kurs")
                            ikkemoms = 0    
                            call materialer(2)
                            
                           
                            
                            im = im + 1
                            m = m + 1
                            oRec.movenext
                            wend
                            oRec.close
	                        
	                        if m = 0 then%>
	                        <tr><td colspan=9><br />(Der er ikke fundet nogen materialer i det valgte interval.)</td></tr>
	                        <%end if  %>
                        	
                             <input id="FM_antal_materialer_ialt" name="FM_antal_materialer_ialt" value="<%=m%>" type="text" />
	                        </table>