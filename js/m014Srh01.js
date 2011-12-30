//--------
// Description: singular items filter construct
// caller : (1) m014s0101.asp (Quick Search)
//          (2) m014s0102.asp (Advanced Search)
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
//        is still under construction  
//
// Update Log: 
//--------
// debug
function ACfltr_14(chvOprd,chrNot,chvOptr,chvStg1,chvStg2) {
	var stgFilter = "" ;
	switch (chvOprd) {
		// Purchase Requisition No
		case "77" :
			if (chvOptr == '3') {
				stgFilter += " insPurchase_Req_id = '"   + chvStg1 + "'  ";			
			} else {
				stgFilter += " insPurchase_req_id BETWEEN " + chvStg1 + " AND " + chvStg2 + " ";
			}
		break;
		// Received Date
		case "78" :
			stgFilter += " dtsDate_Received between '" + chvStg1 + "'" ;
			stgFilter += " AND '" + chvStg2 + "'" ; 
		break;
		// Equipment on Backorder
		case "79" :
			if (chvOptr == '15' ) { 
				stgFilter += " bitInv_on_bk_order = 0 " ; 
			} else { 
				stgFilter += " bitInv_on_bk_order = 1 " ;
			};
		break;
		// Request Type
		case "80" :
			stgFilter += " insRequest_type_id = " + chvStg1 ;
		break;
		// Work Order
		case "81" :
			stgFilter += " insWork_order_id = " + chvStg1 ;
		break;
		// Vendor
		case "82" :
			stgFilter += " insSupplier_id = " + chvStg1 ;
		break;
		// Purchase Status
		case "83" :
			stgFilter += " insPurchase_sts_id = " + chvStg1 ;
		break;
		// Inventory Class
		case "84" :
			stgFilter += " insEquip_Class_id = '" + chvStg1 + "' " ; 
		break;
		// Specified Client				   
		case "94" :
			stgFilter += " intFor_Adult_id = '" + chvStg1 + "' " ; 
		break;				   
	}  
	return (stgFilter) ;
}