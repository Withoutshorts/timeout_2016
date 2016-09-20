<%
'public strConnect_aktiveDB
function aktivedb(x)
		
		Select Case x
	    Case 1
		strConnect_aktiveDB = "timeout_intranet64"
        Case 2
		strConnect_aktiveDB = "timeout_oliver64" 
        Case 3
		strConnect_aktiveDB = "timeout_start64"  
		case 23
		strConnect_aktiveDB = "timeout_dencker64"
		case 107
		strConnect_aktiveDB = "timeout_epi64"
		case 108
		strConnect_aktiveDB = "timeout_jttek64"
		case 109
		strConnect_aktiveDB = "timeout_wwf64"
        case 119
        strConnect_aktiveDB = "timeout_synergi164"
        case 123
		strConnect_aktiveDB = "timeout_kejd_pb64"
        case 126
		strConnect_aktiveDB = "timeout_epi_no64"
		case 136
        strConnect_aktiveDB = "timeout_hidalgo64"
	    case 138
        strConnect_aktiveDB = "timeout_oko64"
		case 139
        strConnect_aktiveDB = "timeout_hestia64"
        case 141
        strConnect_aktiveDB = "timeout_nt64"
        case 152
        strConnect_aktiveDB = "timeout_tec64"
        case 154
        strConnect_aktiveDB = "timeout_esn64"
        case 155
        strConnect_aktiveDB = "timeout_wilke64"
        case 158 
        strConnect_aktiveDB = "timeout_krj64"
        case 159
        strConnect_aktiveDB = "timeout_cisu64"
        case 160
        strConnect_aktiveDB = "timeout_bf64"
        case 161
        strConnect_aktiveDB = "timeout_dv64"

           


	case else
	strConnect_aktiveDB = "nogo"
	End Select


end function
%>
