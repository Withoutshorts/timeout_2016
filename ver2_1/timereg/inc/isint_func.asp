<%
Public isInt
function erDetInt(FM_felt) 
isInt = instr(lcase(FM_felt), "a")
isInt = isInt + (instr(lcase(FM_felt), "b"))
isInt = isInt + (instr(lcase(FM_felt), "c"))
isInt = isInt + (instr(lcase(FM_felt), "d"))
isInt = isInt + (instr(lcase(FM_felt), "e"))
isInt = isInt + (instr(lcase(FM_felt), "f"))
isInt = isInt + (instr(lcase(FM_felt), "g"))
isInt = isInt + (instr(lcase(FM_felt), "h"))
isInt = isInt + (instr(lcase(FM_felt), "i"))
isInt = isInt + (instr(lcase(FM_felt), "j"))
isInt = isInt + (instr(lcase(FM_felt), "k"))
isInt = isInt + (instr(lcase(FM_felt), "l"))
isInt = isInt + (instr(lcase(FM_felt), "m"))
isInt = isInt + (instr(lcase(FM_felt), "n"))
isInt = isInt + (instr(lcase(FM_felt), "o"))
isInt = isInt + (instr(lcase(FM_felt), "p"))
isInt = isInt + (instr(lcase(FM_felt), "q"))
isInt = isInt + (instr(lcase(FM_felt), "r"))
isInt = isInt + (instr(lcase(FM_felt), "s"))
isInt = isInt + (instr(lcase(FM_felt), "t"))
isInt = isInt + (instr(lcase(FM_felt), "u"))
isInt = isInt + (instr(lcase(FM_felt), "v"))
isInt = isInt + (instr(lcase(FM_felt), "w"))
isInt = isInt + (instr(lcase(FM_felt), "x"))
isInt = isInt + (instr(lcase(FM_felt), "y"))
isInt = isInt + (instr(lcase(FM_felt), "z"))
isInt = isInt + (instr(lcase(FM_felt), "�"))
isInt = isInt + (instr(lcase(FM_felt), "�"))
isInt = isInt + (instr(lcase(FM_felt), "�"))
isInt = isInt + (instr(lcase(FM_felt), "<"))
isInt = isInt + (instr(lcase(FM_felt), ">"))
''isInt = isInt + (instr(lcase(FM_felt), ","))
isInt = isInt + (instr(lcase(FM_felt), "?"))
isInt = isInt + (instr(lcase(FM_felt), "!"))
isInt = isInt + (instr(lcase(FM_felt), "#"))
isInt = isInt + (instr(lcase(FM_felt), "%"))
isInt = isInt + (instr(lcase(FM_felt), "&"))
isInt = isInt + (instr(lcase(FM_felt), "/"))
isInt = isInt + (instr(lcase(FM_felt), "("))
isInt = isInt + (instr(lcase(FM_felt), ")"))
isInt = isInt + (instr(lcase(FM_felt), "["))
isInt = isInt + (instr(lcase(FM_felt), "]"))
isInt = isInt + (instr(lcase(FM_felt), "="))
isInt = isInt + (instr(lcase(FM_felt), ";"))
isInt = isInt + (instr(lcase(FM_felt), ":"))
isInt = isInt + (instr(lcase(FM_felt), "�"))
isInt = isInt + (instr(lcase(FM_felt), "@"))
isInt = isInt + (instr(lcase(FM_felt), "'"))
isInt = isInt + (instr(lcase(FM_felt), " "))
end function


Public vTxt
function illChar(thisTxt) 
vTxt = replace(thisTxt, "<", "")
vTxt = replace(vTxt, ">", "")
vTxt = replace(vTxt, ",", "")
vTxt = replace(vTxt, "?", "")
vTxt = replace(vTxt, "!", "")
vTxt = replace(vTxt, "#", "")
vTxt = replace(vTxt, "%", "")
vTxt = replace(vTxt, "&", "og")
vTxt = replace(vTxt, "/", "")
vTxt = replace(vTxt, "(", "")
vTxt = replace(vTxt, ")", "")
vTxt = replace(vTxt, "[", "")
vTxt = replace(vTxt, "]", "")
vTxt = replace(vTxt, "=", "")
vTxt = replace(vTxt, ";", "")
vTxt = replace(vTxt, ":", "")
vTxt = replace(vTxt, "�", "")
vTxt = replace(vTxt, "@", "")
vTxt = replace(vTxt, "'", "")
vTxt = replace(vTxt, ".", "_")
vTxt = replace(vTxt, "�", "")
vTxt = replace(vTxt, "�", "")
vTxt = replace(vTxt, " ", "_")
end function

%>
