//--------
// Description: singular items filter construct
// caller : (1) m001s0101B.asp
//          (2) m001s0102.asp
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
//        Filter for (i)  Courses
//                   (ii) Equipment Loaded
//                   (iii)Equipment Owned,
//                   (iv) Past Services rec'd
//        are still under construction
//--------
// debug
function ACfltr_01(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
	switch (chvOprd) {
		// Last Name
		case "11" :
			stgFilter += " chvLst_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  "; break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  "; break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' "; break;
			};
		break;
		// First Name
		case "12" :
			stgFilter += " chvFst_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  "; break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  "; break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' "; break;
			};
		break;
		// Case Manager
		case "13" :
			stgFilter += " insCase_mngr_id = "   + chvStg1 + "  ";
		break;
		// Gender
		case "17" :
			if (chvOptr == '7' ) {
				stgFilter += " chrGender = 'F' " ;
			} else {
				stgFilter += " chrGender = 'M' " ;
			};
		break;
		// Referral Date
		case "18" :
			stgFilter += " dtsRefral_date between '" + chvStg1 + "'" ;
			stgFilter += " AND '" + chvStg2 + "'" ;
		break;
		// SETBC served
		case "21" :
			if (chvOptr == '14' ) { 
				stgFilter += " e.bitIs_Prx_SETBC = 1 " ; 
			} else { 
				stgFilter += " e.bitIs_Prx_SETBC = 0 " ;
			};
		break;		
		// Re-referral Date
		case "22" :
			stgFilter += " dtsRe_refral_date between '" + chvStg1 + "'";
			stgFilter += " AND '" + chvStg2 + "'" ;
		break;
		// ASP no.
		case "33" :
			stgFilter += " intAdult_Id  = " + chvStg1 ;
		break;
		case "14","19","20" : 
			alert("Search item is still under construction ...");
		break;
		// PRCVI served
		case "256" :
			if (chvOptr == '14' ) { 
				stgFilter += " e.bitIs_Prx_PRCVI = 1 " ; 
			} else { 
				stgFilter += " e.bitIs_Prx_PRCVI = 0 " ;
			};
		break;
		// First Nations	  
		case "257" :
			if (chvOptr == '14' ) { 
				stgFilter += " e.bitIs_FirstNations = 1 " ; 
			} else { 
				stgFilter += " e.bitIs_FirstNations = 0 " ;
			};
		break;
		// Multiple Disabilities	  
		case "258" :
			if (chvOptr == '14' ) { 
				stgFilter += " e.insDsbty1_id > 0 AND e.insDsbty2_id > 0 " ; 
			} else { 
				stgFilter += " e.insDsbty1_id > 0 AND e.insDsbty2_id = 0  " ;
			};
		break;		
	}
	return (stgFilter) ;
}