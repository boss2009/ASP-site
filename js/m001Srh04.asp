// ------------------------------------------
// Description: singular items filter construct
// caller : (1) m001s0103.asp
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
// ------------------------------------------
// debug
function ACfltr_03(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
	switch (chvOprd) {
        case "14" :
			stgFilter += " (dtsRequest_Date between '" + chvStg1 + "'" ;
		    stgFilter += " AND '" + chvStg2 + "') and insSrv_Code_id is NOT NULL " ; 
		break;
	}  
	return (stgFilter) ;
}
