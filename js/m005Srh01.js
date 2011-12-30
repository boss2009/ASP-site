//--------
// Description: singular items filter construct
// caller : (1) m005s0101.asp (Quick Search)
//          (2) m005s0102.asp (Advanced Search)
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
//
// Update Log: 
//--------
// debug
function ACfltr_05(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
	switch (chvOprd) {
		// Bundle Name				   
        case "122" :
			stgFilter += " chvName " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  "; break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '" + chvStg1 + "'  "; break;
				case "4" : stgFilter += "LIKE '%25" + chvStg1 + "' "; break;
			};
		break;
		// Bundle Status
        case "123" :
			stgFilter += " bitBundle_Status = " + chvStg1 ;
		break;
		// Bundle Type
        case "124" :
			stgFilter += " chrBundle_Type = " + chvStg1 ;
		break;
		// Bundle Purpose
        case "125" : 
			switch(chvOptr) {
				case "0" : stgFilter += " bitFor_CSG  = 1 "; break;
				case "1" : stgFilter += " bitFor_Loan = 1 "; break;
				case "2" : stgFilter += " bitFor_CSG = 1 AND bitFor_Loan = 1"; break;
			};
		break;
		// Inventory Class
		case "127" :
			stgFilter += " insEquip_Class_id = '" + chvStg1 + "' " ; 
		break;
	}  
	return (stgFilter) ;
}