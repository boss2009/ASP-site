//--------
// Description: Equipment Class filter construct
// caller : (1) m007s0101.asp
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
// Update Log: 
//--------
function ACfltr_07(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
	switch (chvOprd) {
		// Name
		case "1" :
			stgFilter += " b.chvName " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '" + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Tax
		case "2" :
			stgFilter += " b.chvSbjTotax ";
			stgFilter += (chrNot == "1" ) ? "<>" : "=" 
			stgFilter += " '"   + chvStg1 + "'  ";
		break;
		// Vendor
		case "3" :
			stgFilter += " a.insVendor_id = '" + chvStg1 + "'  ";
		break;
	}  
	return (stgFilter) ;
}