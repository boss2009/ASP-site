//--------
// Description: singular items filter construct
// caller : (1) m003s0102.asp
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
//        Filter for (i)  Inventory Class
//                   (iii)Delivery Date - Loan History
//                   (iv) Institution   - Loan History
//        are still under construction  
//
// Update Log: 
//    - update to house Inventory User       + APR.09.2002
//--------
// debug
function ACfltr_03(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
	switch (chvOprd) {
		// Inventory Class
		case "39" :
			stgFilter += " a.insEquip_Class_id = '" + chvStg1 + "' " ; 
		break;
		// Inventory Id
        case "41" :
			stgFilter += " intBar_Code_no = '"   + chvStg1 + "'  ";
		break;
		// Serial No.
        case "42" :
			stgFilter += " chvSerial_Number " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Requisition No
        case "43" :
			stgFilter += " intRequisition_no = '"   + chvStg1 + "'  ";
		break;
		// Processed Date 				   
        case "44" :
			stgFilter += " dtsOrd_Date between '" + chvStg1 + "'" ;
		    stgFilter += " AND '" + chvStg2 + "'" ; 
		break;
		// Received Date 				   
        case "45" :
			stgFilter += " dtsRec_Date between '" + chvStg1 + "'" ;
			stgFilter += " AND '" + chvStg2 + "'" ; 
		break;
		// Inventory Name				   
        case "46" :
			stgFilter += " chvInventory_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Vendor
        case "47" :
			stgFilter += " insVendor_id = "   + chvStg1 + "  ";
		break; 
		// Delivery date
        case "48" :
			stgFilter += " dtsDlvy_date between '" + chvStg1 + "'" ;
			stgFilter += " AND '" + chvStg2 + "'" ; 
		break;
		// Institution
        case "49" :
			stgFilter += " insInstit_User_id = '" + chvStg1 + "'  ";
		break;
		// Purchased by
        case "50" :
			stgFilter += " chrObtain_by = '" + chvStg1 + "' ";
		break;
		// Client user First name
        case "259" :
			stgFilter += " insUser_Type_id = 3 AND chvUsr_Fst_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Client user Last name
        case "260" :
			stgFilter += " insUser_Type_id = 3 AND chvUsr_Lst_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Staff First name
        case "261" :
			stgFilter += " insUser_Type_id = 1 AND chvUsr_Fst_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Staff Last name
        case "262" :
			stgFilter += " insUser_Type_id = 1 AND chvUsr_Lst_Name " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
		// Institution User
        case "58" :
			stgFilter += " chvInstitUsr_Nm " ;
			switch(chvOptr) {
				case "1" : stgFilter += "LIKE '" + chvStg1 + "%25'  ";   break;
				case "2" : stgFilter += "LIKE '%25" + chvStg1 + "%25' "; break;
				case "3" : stgFilter += "= '"     + chvStg1 + "'  ";     break;
				case "4" : stgFilter += "LIKE '%25"  + chvStg1 + "' ";   break;
			};
		break;
	}  
	return (stgFilter) ;
}
