//--------
// Description: singular items filter construct
// caller : (1) m006s0101.asp (Quick Search)
//          (2) m006s0102.asp (Advanced Search)
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
// ===========
//    - update script to house Company ID               + Feb.13.2003
//--------
// debug
function ACfltr_06(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
	switch (chvOprd) {
		// Inventory Class
		case "109" :
			stgFilter += " insEquip_Class_id = '" + chvStg1 + "' " ; 
		break;
		// Company Name				   
		case "110" :
			stgFilter += " chvCompany_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Contract PO no
		case "111" :
			stgFilter += " chvContract_PO " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Company Type
		case "112" :
			stgFilter += " insWork_Typ_id = " + chvStg1 ;
		break;
		// Company ID                               + Feb.13.2003
		case "255" :
			stgFilter += " intCompany_id = " + chvStg1 ;
		break;
	}  
	return (stgFilter) ;
}