// ------------------------------------------
// Description: singular items filter construct
// caller : (1) m009s0101.asp
//
// Note : this script assumes the following var
//        are defined at the caller - 
//        (1) stgFilter
//        (2) chvOprd
//        (3) chrNot
//        (4) chvOptr
//        (5) chvStg1
//        (6) chvStg2
//
//        Filter for (i)  Referral Type
//                   (ii) Funding Source  
//        are still under construction  
//
// Update Log: 
// ------------------------------------------
// debug
function ACfltr_09(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
	switch (chvOprd) {
		// Requested Date
		case "272":
			stgFilter += " dtsRequested_date between '" + chvStg1 + "'" ;
			stgFilter += " AND '" + chvStg2 + "'" ; 
		break;
		// Completed Date
		case "273":
			stgFilter += " dtsCompleted_Date between '" + chvStg1 + "'" ;
			stgFilter += " AND '" + chvStg2 + "'" ; 
		break;
		// Inventory
		case "274":
			stgFilter += " intEquip_Set_id = " + chvStg1 + " " ; 
		break;
		// Referral Type
		case "276": 
			alert("Search item is still under construction ...");
		break;
		// Funding Source
		case "277": 
			alert("Search item is still under construction ...");
		break;
		// Repair Status
		case "279":
			stgFilter += " insRepair_Status = " + chvStg1 + " " ; 
		break;
		// Case Manager
		case "280":
			stgFilter += " insCase_mngr_id = " + chvStg1 + "  ";
		break;
		// Reason for Repair - '1' User Error - '0' H/W defect
		case "281":
			stgFilter += " chrReason_Repair = '"   + chvStg1 + "' ";
		break;
		// Type of Repair - '1' Covered by warranty - '0' not covered by warranty
		case "282":
			stgFilter += " bitIs_Covered_Warnty = "   + chvStg1 + " ";
		break;
		// Equipment Service ID
		case "285":
			stgFilter += " intEquip_Srv_id = " + chvStg1 + " " ; 
		break;
		// Shipping delay  + May.01.2003
		case "287" :
			if (chvOptr == '14' ) { 
				stgFilter += " bitIsDlvy_onshdl = 1 " ; 
			} else { 
				stgFilter += " bitIsDlvy_onshdl = 0 " ;
			};
		break;
		// Delay Resolved + May.01.2003
		case "288" :
			if (chvOptr == '14' ) { 
				stgFilter += " bitIsDlvy_delay = 1 " ; 
			} else { 
				stgFilter += " bitIsDlvy_delay = 0 " ;
			};
		break;
		// Outside Srv completed  + May.01.2003
		case "289" :
			if (chvOptr == '14' ) { 
				stgFilter += " bitIs_Completed = 1 " ; 
			} else { 
				stgFilter += " bitIs_Completed = 0 " ;
			};
		break;
		// Labor charged + May.01.2003
		case "290" :
			if (chvOptr == '14' ) { 
				stgFilter += " bitLC_Status = 1 " ; 
			} else { 
				stgFilter += " bitLC_Status= 0 " ;
			};
		break;			
	}  
	return (stgFilter) ;
}
