//--------
// Description: singular items filter construct
// caller : (1) m012s0101.asp (Quick Search)
//          (2) m012s0102.asp (Advanced Search)
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
function ACfltr_12(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
	switch (chvOprd) {
		// Institution No
		case "134" :
			stgFilter += " insSchool_id = '"   + chvStg1 + "'  ";
		break;
		// Institution Name				   
		case "135" :
			stgFilter += " chvSchool_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Institution Type
		case "138" :
			stgFilter += " insSchool_type_id = " + chvStg1 ;
		break;
		// Campus Type
		case "139" :
			stgFilter += " bitIs_MainCampus = " + chvStg1 ;
		break;
		// Service Date 				   
		case "141":
			stgFilter += " dtsService_Date between '" + chvStg1 + "'" ;
			stgFilter += " AND '" + chvStg2 + "'" ; 
		break;
		// Referral Type
		case "147" :
			stgFilter += " insRefAgt_id = " + chvStg1 ;
		break;
		// Referral Count 				   
		case "149" :
			stgFilter += " intRCnt " ;
			switch(chvOptr) {
				case "16" : stgFilter += ">  " + chvStg1 + " "; break;
				case "17" : stgFilter += ">= " + chvStg1 + " "; break;
				case "18" : stgFilter += "=  " + chvStg1 + " "; break;
				case "19" : stgFilter += "<  " + chvStg1 + " "; break;
				case "20" : stgFilter += "<= " + chvStg1 + " "; break;
			};
		break;
		// Referral Date
		case "150":
			stgFilter += " dtsReferral_date between '" + chvStg1 + "'" ;
			stgFilter += " AND '" + chvStg2 + "'" ; 
		break;
		// Contact
		case "151" :
			stgFilter += " intContact_id = " + chvStg1 ;
		break;
	}  
	return (stgFilter) ;
}