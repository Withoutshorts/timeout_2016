  <%
public strConnect_aktiveDB
function aktivedb(x)
		
		Select Case x
	    Case 1
		strConnect_aktiveDB = "timeout_intranet64"
        Case 2
		strConnect_aktiveDB = "timeout_oliver64" 
        Case 3
		strConnect_aktiveDB = "timeout_start64"  
        Case 4
		strConnect_aktiveDB = "timeout_demo64"  
		case 23
		strConnect_aktiveDB = "timeout_dencker64"
        case 55
		strConnect_aktiveDB = "timeout_acc64"
        case 87
        strConnect_aktiveDB = "timeout_fk64"
        case 88
        strConnect_aktiveDB = "timeout_immenso64"
		case 89
        strConnect_aktiveDB = "timeout_essens64"
        case 91
	    strConnect_aktiveDB = "timeout_cst64"
        case 107
		strConnect_aktiveDB = "timeout_epi64"
		case 108
		strConnect_aktiveDB = "timeout_jttek64"
		case 109
		strConnect_aktiveDB = "timeout_wwf64"
        case 110
		strConnect_aktiveDB = "timeout_mmmi64"
        case 115
        strConnect_aktiveDB = "timeout_mi64"
        case 119
        strConnect_aktiveDB = "timeout_synergi164"
        case 123
		strConnect_aktiveDB = "timeout_kejd_pb64"
        case 126
		strConnect_aktiveDB = "timeout_epi_no64"
        case 130
		strConnect_aktiveDB = "timeout_rek64"
		case 136
        strConnect_aktiveDB = "timeout_hidalgo64"
	    case 138
        strConnect_aktiveDB = "timeout_oko64"
		case 139
        strConnect_aktiveDB = "timeout_hestia64"
        case 140
        strConnect_aktiveDB = "timeout_epi_uk64" 
        case 141
        strConnect_aktiveDB = "timeout_nt64"
        case 142
        strConnect_aktiveDB = "timeout_otvvs64"
        case 145
        strConnect_aktiveDB = "timeout_sdutek64"
        case 146
        strConnect_aktiveDB = "timeout_glad64"
        case 152
        strConnect_aktiveDB = "timeout_tec64"
        case 153
        strConnect_aktiveDB = "timeout_akelius64"
        case 154
        strConnect_aktiveDB = "timeout_esn64"
        case 155
        strConnect_aktiveDB = "timeout_wilke64"
        case 156
        strConnect_aktiveDB = "timeout_adra64"
        case 158 
        strConnect_aktiveDB = "timeout_krj64"
        case 159
        strConnect_aktiveDB = "timeout_cisu64"
        case 160
        strConnect_aktiveDB = "timeout_bf64"
        case 161
        strConnect_aktiveDB = "timeout_dv64"
        case 162
        strConnect_aktiveDB = "timeout_plan64"
        case 163
        strConnect_aktiveDB = "timeout_sdeo64"
        case 164
        strConnect_aktiveDB = "timeout_epi201764"
        case 165
        strConnect_aktiveDB = "timeout_eniga64"
        case 166
        strConnect_aktiveDB = "timeout_tbg64"
        case 167
        strConnect_aktiveDB = "timeout_gd64"
        case 168
        strConnect_aktiveDB = "timeout_assurator64"
        case 169
        strConnect_aktiveDB = "timeout_tia64"
        case 170
        strConnect_aktiveDB = "timeout_welcom64"
        case 171
        strConnect_aktiveDB = "timeout_adrasudan64"
        case 172
        strConnect_aktiveDB = "timeout_alfanordic64"
        case 173
        strConnect_aktiveDB = "timeout_mpt64"
        case 174
        strConnect_aktiveDB = "timeout_cflow64"
        case 175
        strConnect_aktiveDB = "timeout_wap64"

      

	case else
	strConnect_aktiveDB = "nogo"
	End Select


end function
%>
