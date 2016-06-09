<%
Function ConvertDate(sDato)
	' Konverterer en engelsk dato (mm/dd/yyyy)
	' til en dansk dato (dd/mm/yyyy)
	'Response.write sDato
	ConvertDate = Day(sDato) & "/" & Month(sDato) & "/" & Year(sDato)
End Function

Function ConvertDateUS(sDato)
	' Konverterer en dansk dato (dd/mm/yyyy)
	' til en engelsk dato (mm/dd/yyyy)
	ConvertDateUS = Month(sDato) & "/" & Day(sDato) & "/" & Year(sDato)
End Function

Function ConvertDateYMD(sDato)
	' Konverterer en dato (dd/mm)
	' til en engelsk dato (mm/dd)
	ConvertDateYMD = Year(sDato) & "/" &  Month(sDato) & "/" & Day(sDato)
End Function

Function ConvertDateMD(sDato)
	' Konverterer en dato (dd/mm)
	' til en engelsk dato (mm/dd)
	ConvertDateMD = Month(sDato) & "/" & Day(sDato)
End Function
%>
