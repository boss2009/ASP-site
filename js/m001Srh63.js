//--------
// Description: singular items filter construct
// caller : (1) m001r0163.asp
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
//
// Update Log: 
//    - update script to synchronous cp_Adult_Client2()
//--------
// debug
function ACfltr_63(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
    switch (chvOprd) {
		// Last Name
        case "199" :
			stgFilter += " chvLst_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// First Name				   
        case "200" :
			stgFilter += " chvFst_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Gender
        case "201" :
			if (chvOptr == '7' ) { 
				stgFilter += " chrGender = 'F' " ; 
			} else { 
				stgFilter += " chrGender = 'M' " ;
			};
		break;
		// Referral Date 
        case "202" :
			stgFilter += " d.dtsRefral_date between '" + chvStg1 + "'" ;
			stgFilter += " AND '" + chvStg2 + "'" ; 
		break;
		// Re-referral Date 
        case "203" :
			stgFilter += " e.dtsRe_refral_date between '" + chvStg1 + "'";
		    stgFilter += " AND '" + chvStg2 + "'" ; 
		break;
		// ASP no.
        case "204" :
			stgFilter += " e.intAdult_Id  = " + chvStg1 ;
		break;
		// Case Manager
        case "211" :
			stgFilter += " insCase_mngr_id = '"   + chvStg1 + "'  ";
		break;
	}  
	return (stgFilter) ;
}