//--------
// Description: singular items filter construct
// caller : (1) m004s0101.asp (Quick Search)
//          (2) m004s0102.asp (Advanced Search)
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
function ACfltr_04(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
	switch (chvOprd) {
		// Contact No
        case "103" :
			stgFilter += " intContact_id = '"   + chvStg1 + "'  ";
		break;
		// First Name				   
        case "104" :
			stgFilter += " chvFst_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Last Name
        case "105" :stgFilter += " chvLst_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Contact Type
        case "106" :
			stgFilter += " intWork_type_id = " + chvStg1 ;
		break;
		// Mailing List
        case "107" :
			stgFilter += " insMail_list_id = " + chvStg1 ;
		break;
	}  
	return (stgFilter) ;
}