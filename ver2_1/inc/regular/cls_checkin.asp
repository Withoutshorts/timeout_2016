<!--

    Check in med n'gle kort funktioner

-->
<% 

    public medid, medarb_navn
    Function CheckKey(key, rdir) 

        strSQL = "SELECT Mid, Mnavn FROM medarbejdere WHERE (medarbejder_rfid = "& key & " OR mnr = "& key & ") AND (mansat = 1 OR mansat = 3)"
        'response.Write strSQL
        oRec.open strSQL, oConn, 3
        if not oRec.EOF then
        medarb_navn = oRec("Mnavn")
        medid = oRec("Mid")  
        session("login") = Trim(oRec("Mid"))
	    session("mid") = Trim(oRec("Mid"))
        session("user") = oRec("Mnavn")
        'session("rettigheder") = oRec("rettigheder")
        else
            select case cint(rdir)
                case 0
                response.Redirect "monitor.asp?func=startside&redType=3"
                case 1
                response.Redirect "checkinscreen.asp?message=2"
                case 3
                response.Redirect "infoscreen.asp?func=EmployeeNotFound"
            end select
        end if
        oRec.close

    end function



%>