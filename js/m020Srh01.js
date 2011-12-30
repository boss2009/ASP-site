//--------
// Description: singular items filter construct for Issue Manager
// caller : (1) m020r0101.asp
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
//        chvOprd : (1) Priority
//                  (2) Status
//                  (3) Assigned To
//                  (4) Issue Name Keyword
//                  (5) Assigned by me
//                  (6) Assigned to me
//
// Update Log: 
//--------
// debug
function ACfltr_20(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
	switch (chvOprd) {
		// Priority
		case "1" :
			stgFilter += " intPriority_id = '" + chvStg1 + "' " ; 
		break;
		// Status
		case "2" :
			stgFilter += " intStatus_id = '" + chvStg1 + "' " ; 
		break;
		// Assigned to
		case "3" :
			stgFilter += " intAssigned_to = '" + chvStg1 + "' " ; 
		break;
		// Issue Keyword
		case "4" :
			stgFilter += " a.ncvIssue_name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Assigned by me
		case "5" :
			stgFilter += " intUser_id = '"   + chvStg1 + "'  ";
		break;
		// Assigned to me
		case "6" :
			stgFilter += " intIssue_id = '"   + chvStg1 + "'  ";
		break;
		//Module ID		
		case "7" :
			stgFilter += " intMODno = '" + chvStg1 + "' "; 
		break;
		//Function ID		
		case "8" :
			stgFilter += " i.insFTNid = '" + chvStg1 + "' "; 
		break;
	}  
	return (stgFilter) ;
}