<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% 
// retrieve the sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup(714)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

// retrieve the string search operands
var rsOprd1 = Server.CreateObject("ADODB.Recordset");
rsOprd1.ActiveConnection = MM_cnnASP02_STRING;
rsOprd1.Source = "{call dbo.cp_ASP_Lkup(715)}";
rsOprd1.CursorType = 0;
rsOprd1.CursorLocation = 2;
rsOprd1.LockType = 3;
rsOprd1.Open();

// retrieve the Lookup search operands
var rsOprd2 = Server.CreateObject("ADODB.Recordset");
rsOprd2.ActiveConnection = MM_cnnASP02_STRING;
rsOprd2.Source = "{call dbo.cp_ASP_Lkup(710)}";
rsOprd2.CursorType = 0;
rsOprd2.CursorLocation = 2;
rsOprd2.LockType = 3;
rsOprd2.Open();

// retrieve the Class search operands
var rsOprd3 = Server.CreateObject("ADODB.Recordset");
rsOprd3.ActiveConnection = MM_cnnASP02_STRING;
rsOprd3.Source = "{call dbo.cp_ASP_Lkup(708)}";
rsOprd3.CursorType = 0;
rsOprd3.CursorLocation = 2;
rsOprd3.LockType = 3;
rsOprd3.Open();

// retrieve the multi-select search Operands
var rsOprd3 = Server.CreateObject("ADODB.Recordset");
rsOprd3.ActiveConnection = MM_cnnASP02_STRING;
rsOprd3.Source = "{call dbo.cp_ASP_Lkup(716)}";
rsOprd3.CursorType = 0;
rsOprd3.CursorLocation = 2;
rsOprd3.LockType = 3;
rsOprd3.Open();
%>
<html>
<head>
	<title>Client - Advanced Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/m001Srh02.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="javascript">
	function CnstrFltr(output) {		
		var inspSrtBy = document.frm01s01.SortByColumn.value;
		var inspSrtOrd = document.frm01s01.OrderBy.value;
		var stgFilter = document.frm01s01.QueryString.value;
		if (output==1) {
			document.frm01s01.action = "m001q02.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;
			document.frm01s01.submit() ; 
		} else {
			var SearchExcel = window.open("m001q02excel.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter);			
		}
	}

	function initscr() {
		document.frm01s01.IsNot.disabled = true;
		oOptrd1.style.visibility="visible";
		oStg1.style.visibility="visible";
		document.frm01s01.StringSearchTextOne.style.visibility="visible";
		oStg2.style.visibility="hidden";
		oOptrd31.style.visibility="hidden";
		oStg21.style.visibility="hidden";
		oOptrd3.style.visibility="hidden";
		oOptr3.style.visibility="hidden";
		initializeList('frm01s01.MultiValueSearchOperand','frm01s01.MultiValueSearchText');	
		oOld3.style.visibility="hidden";
		document.frm01s01.StringSearchOperand.focus();		
		selectChange(frm01s01.StringSearchOperand, frm01s01.StringSearchOperator,detailData1);		
	}

	function Selckb(){
		var len = document.frm01s01.IncludeHistoricValues.length;
		var j = document.frm01s01.MultiValueSearchOperand[document.frm01s01.MultiValueSearchOperand.selectedIndex].value ;
		var blnFlag = true;
		if (document.frm01s01.IncludeHistoricValues.value == "1" ) {
			document.frm01s01.IncludeHistoricValues.checked = false ;
			document.frm01s01.IncludeHistoricValues.value = "0";
		} else {
			document.frm01s01.IncludeHistoricValues.checked = true ;
			document.frm01s01.IncludeHistoricValues.value = "1";
		}
		var k = document.frm01s01.IncludeHistoricValues.value;
		blnFlag = document.frm01s01.IncludeHistoricValues.checked;
		initializeList('frm01s01.MultiValueSearchOperand','frm01s01.MultiValueSearchText');			
	}

	function SelOpt() {
		var len = document.frm01s01.SearchType.length;
		var Idparam = 1;
		for (var i=0;i <len; i++){
			if (document.frm01s01.SearchType[i].checked)  Idparam = i;
		}
	
		switch ( Idparam ) {
			case 0: 
				window.status = "Single Term";
				oOptrd1.style.visibility="visible";
				oStg1.style.visibility="visible";
				if ((document.frm01s01.StringSearchOperand.value!="21") && (document.frm01s01.StringSearchOperand.value!="256") && (document.frm01s01.StringSearchOperand.value!="257") && (document.frm01s01.StringSearchOperand.value!="258")){								
					document.frm01s01.StringSearchTextOne.style.visibility="visible";								
				}
				oOptrd31.style.visibility="hidden";
				oStg21.style.visibility="hidden";
				oOptrd3.style.visibility="hidden";
				oOptr3.style.visibility="hidden";
				oOld3.style.visibility="hidden";
				if ((document.frm01s01.StringSearchOperand.value == "18") || (document.frm01s01.StringSearchOperand.value == "22")){
					document.frm01s01.StringSearchTextTwo.disabled = false;		
					oStg2.style.visibility="visible";
				} else {
					document.frm01s01.StringSearchTextTwo.disabled = true;
					oStg2.style.visibility="hidden";
				}				   
			break;
			case 1:
				window.status = "Lookup Term";			
				oOptrd1.style.visibility="hidden";
				oStg1.style.visibility="hidden";
				document.frm01s01.StringSearchTextOne.style.visibility="hidden";
				oStg2.style.visibility="hidden";				
				oOptrd31.style.visibility="visible";
				oStg21.style.visibility="visible";
				oOptrd3.style.visibility="hidden";
				oOptr3.style.visibility="hidden";
				oOld3.style.visibility="hidden";
				selectChange(frm01s01.LookupValueSearchOperand, frm01s01.LookupValueSearchOptions,detailData2);
			break;
			case 2: 
				window.status = "Multiple Term";			
				oOptrd1.style.visibility="hidden";
				oStg1.style.visibility="hidden";
				document.frm01s01.StringSearchTextOne.style.visibility="hidden";
				oStg2.style.visibility="hidden";
				oOptrd31.style.visibility="hidden";
				oStg21.style.visibility="hidden";
				oOptrd3.style.visibility="visible";
				oOptr3.style.visibility="visible";
				oOld3.style.visibility="visible";
			break;
		}       	
	}

	function Savtxt() {	
		var stgPgQuery = "";
		var stgFilter = "" ;
		var blnFlg = false ;
		var len = document.frm01s01.SearchType.length;
		var Idparam = 1;	
		var stgTemp,j,k; 
	  	
		for (var i=0;i <len; i++){
			if (document.frm01s01.SearchType[i].checked) Idparam = i;
		}
	
		stgTemp = document.frm01s01.QueryString.value; 
		switch (Idparam) {
			case 0: 
				if (document.frm01s01.StringSearchOperand.length>=1) {
					j = document.frm01s01.StringSearchOperand[document.frm01s01.StringSearchOperand.selectedIndex].value ;
				} else {
					alert("Select String Search Operand.");
					document.frm01s01.StringSearchOperand.focus();
					break;
				}
				if (document.frm01s01.StringSearchOperator.length>=1) {			
					l = document.frm01s01.StringSearchOperator[document.frm01s01.StringSearchOperator.selectedIndex].value ;
				} else {
					alert("Select String Search Operator.");
					document.frm01s01.StringSearchOperator.focus();
					break;
				}
				var chvOprd = j;
				var chrNot  = document.frm01s01.IsNot.value;
				var chvOptr = l ;
				var chvStg1 = document.frm01s01.StringSearchTextOne.value;
				var chvStg2 = document.frm01s01.StringSearchTextTwo.value;
				var chvAO1  = document.frm01s01.AndOr.value ;
				if (((j=="18") || (j=="22")) && (l!="0")) {
					if (chvStg1 == "") {
						alert("Enter Start Date");
						document.frm01s01.StringSearchTextOne.focus();
						return ;
					}
					if (chvStg2 == "") {
						alert("Enter End Date");
						document.frm01s01.StringSearchTextTwo.focus();
						return ;
					}
					if (!CheckDateBetween(Trim(chvStg1)+" and "+Trim(chvStg2))) {
						return ;
					}
				}
				if (j == "33") {
					if (!IsID(chvStg1)) {
						alert("Invalid number.");
						document.frm01s01.StringSearchTextOne.focus();
						return;
					}
				}				
				if (chvOptr == "0") {
					alert("Please select operator before Proceed");
				} else {
					stgFilter = ACfltr_01(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
					if (stgTemp.length > 0 ) stgTemp += " (" ;
					stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 ;
				}
			break;
			case 1: 
				if (document.frm01s01.LookupValueSearchOperand.length>=1) {
					j = document.frm01s01.LookupValueSearchOperand[document.frm01s01.LookupValueSearchOperand.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Operand.");
					document.frm01s01.LookupValueSearchOperand.focus();
					break;
				}
				var chvOprd = j;
				var chvAO1  = document.frm01s01.AndOr.value ;
				var chvStg1 = document.frm01s01.LookupValueSearchOptions.value;
				stgFilter = ACfltr_01(chvOprd,"","",chvStg1,"");
				if (stgTemp.length > 0 ) stgTemp += " (" ;
				stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 ;				
			break;			
			case 2: 
				j = document.frm01s01.MultiValueSearchOperand[document.frm01s01.MultiValueSearchOperand.selectedIndex].value ;
				l = "";
				var count = 0;
				var optList = document.frm01s01.MultiValueSearchText;
				for (var x= 0; x < optList.length; x++){
					if (optList[x].selected) count++;				
				}	
				if (count == 0) {
					alert("Select at least one value.");					
					optList.focus();
					return ;	
				}								
				var m = optList.length;
				if (optList.multiple) {
					for(var ii = 0; ii < m; ii++) {
						if (document.frm01s01.MultiValueSearchText[ii].selected) {		
							if (l.length > 0 ) l += "," ;
							l += document.frm01s01.MultiValueSearchText[ii].value ;		
						} 				  
					} 						  
				} else {
					l = document.frm01s01.MultiValueSearchText[document.frm01s01.MultiValueSearchText.selectedIndex].value
				} 
				switch (j) {
					// Region
					case "27" :
						if (stgTemp.length > 0 ) stgTemp += " (" ;
						blnFlg = true;				
						stgFilter += " insRegion_num in (" + l + ") " ; 
					break;
					// Primary Disability
					case "28" : 
						if (stgTemp.length > 0 ) stgTemp += " (" ;
						blnFlg = true;				
						stgFilter += " insDsbty1_id in (" + l + ") " ; 
					break;
					// Student Status
					case "29" : 
						if (stgTemp.length > 0 ) stgTemp += " (" ;
						blnFlg = true;				
						stgFilter += " insStdnt_status_id in (" + l + ") " ; 
					break;
					// Referring Agent (Type)
					case "30": 
						if (stgTemp.length > 0 ) stgTemp += " (" ;
						blnFlg = true;				
						stgFilter += " insRefAgt_id IN (" + l + ") " ; 
					break;
					// Funding Source
					case "31": 
						if (stgTemp.length > 0 ) stgTemp += " (" ;
						blnFlg = true;				
						stgFilter += " insFunding_source_id IN (" + l + ") " ; 
					break;
					case "32": 
						alert(" Service Code option not yet available."); 
					break;
				}
			default: 
			break;
		}						
		document.frm01s01.StringSearchTextOne.value="";
		if (blnFlg) {
			var chvAO1  = document.frm01s01.AndOr.value ;
			stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 + " ";
		}
		document.frm01s01.QueryString.value = stgTemp; 
	}
	</script>
	<script language="javascript">
	var detailData1 = new Array();
	<%
	var intOldOptr = 0;
	var objOptrDesc,objOptrId,objRecID;
	
	var rsOptr1 = Server.CreateObject("ADODB.Recordset");
	rsOptr1.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr1.Source = "{call dbo.cp_SysOptr(0,0,16)}";
	rsOptr1.CursorType = 0;
	rsOptr1.CursorLocation = 2;
	rsOptr1.LockType = 3;
	rsOptr1.Open();
	if (!rsOptr1.EOF){ 	
		while (!rsOptr1.EOF) { 
			objOptrDesc = rsOptr1("chvOptrDesc")
			objOptrId = rsOptr1("intOptrId")
			objRecID = rsOptr1("intRecID")	
			if (intOldOptr != objRecID.value) {
				Response.Write("detailData1["+objRecID+"] = new Array();")
				intOldOptr = objRecID.value
			}
			Response.Write("detailData1["+objRecID+"]["+objOptrId+"] = '"+objOptrDesc+"';");
			rsOptr1.MoveNext 
		}
	} else {
		Response.Write("SysOptr lookup does not exist.")
	}	
	rsOptr1.Close();
	%>
	
	function selectChange(control, controlToPopulate,ItemArray) {
		var myEle ;	  
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		myEle = document.createElement("option") ;
		var y = control.value;
		switch(y){
			case "13":
				document.frm01s01.StringSearchTextOne.style.visibility = "hidden";			
				oStg2.style.visibility="hidden";						
			break;
			case "14":
				document.frm01s01.IsNot.disabled = false;			
			break;
			case "17":
				document.frm01s01.StringSearchTextOne.style.visibility = "hidden";			
				oStg2.style.visibility="hidden";			
			break;
			case "18":
				document.frm01s01.StringSearchTextTwo.disabled = false;		
				document.frm01s01.StringSearchTextOne.style.visibility = "visible";							
				oStg2.style.visibility="visible";			
			break;
			case "21":
				document.frm01s01.StringSearchTextOne.style.visibility = "hidden";			
				oStg2.style.visibility="hidden";			
			break;			
			case "22":
				document.frm01s01.StringSearchTextTwo.disabled = false;		
				document.frm01s01.StringSearchTextOne.style.visibility = "visible";							
				oStg2.style.visibility="visible";			
			break;
			case "256":
				document.frm01s01.StringSearchTextOne.style.visibility = "hidden";			
				oStg2.style.visibility="hidden";			
			break;
			case "257":
				document.frm01s01.StringSearchTextOne.style.visibility = "hidden";			
				oStg2.style.visibility="hidden";			
			break;
			case "258":
				document.frm01s01.StringSearchTextOne.style.visibility = "hidden";			
				oStg2.style.visibility="hidden";			
			break;
			default:
				document.frm01s01.IsNot.disabled = true;		
				document.frm01s01.StringSearchTextTwo.disabled = true;
				oStg2.style.visibility="hidden";
				document.frm01s01.StringSearchTextOne.style.visibility = "visible";							
			break;
		}
		
		document.frm01s01.StringSearchTextOne.value = "";
		
		if (y != 0) {		
			for ( x = 0 ; x < ItemArray[y].length  ; x++){
				if (ItemArray[y][x]) { 
					myEle = document.createElement("option") ;
					myEle.value = x ;
					myEle.text = ItemArray[y][x] ;
					controlToPopulate.add(myEle) ;
				}
			}
		}

		Togo();	  
	}

	function Togo() {	
		var j = document.frm01s01.StringSearchOperand[document.frm01s01.StringSearchOperand.selectedIndex].value ;
		var l = document.frm01s01.StringSearchOperator[document.frm01s01.StringSearchOperator.selectedIndex].value
		document.frm01s01.MM_curOprd.value = j ;
		document.frm01s01.MM_curOptr.value = l ;
		document.frm01s01.MM_flag.value = true ;
	}
	
	var detailData2 = new Array();
	<%
	var intID,chrData;
	
	var rsRegion = Server.CreateObject("ADODB.Recordset");
	rsRegion.ActiveConnection = MM_cnnASP02_STRING;
	rsRegion.Source = "{call dbo.cp_AC_Region(0,0,0)}";
	rsRegion.CursorType = 0;
	rsRegion.CursorLocation = 2;
	rsRegion.LockType = 3;
	rsRegion.Open();
	if (!rsRegion.EOF){ 
		Response.Write("detailData2[1] = new Array();")	
		while (!rsRegion.EOF){ 
			chrData = rsRegion("chvname")
			intID   = rsRegion("insRegion_num")
			Response.Write("detailData2[1]["+intID+"] = '"+chrData+"';");
			rsRegion.MoveNext 
		}
	} else {
		Response.Write("Region lookup does not exist.")
	}	
	rsRegion.Close();

	var rsCaseManager = Server.CreateObject("ADODB.Recordset");
	rsCaseManager.ActiveConnection = MM_cnnASP02_STRING;
	rsCaseManager.Source = "{call dbo.cp_CaseMgr}";
	rsCaseManager.CursorType = 0;
	rsCaseManager.CursorLocation = 2;
	rsCaseManager.LockType = 3;
	rsCaseManager.Open();	
	if (!rsCaseManager.EOF){ 
		Response.Write("detailData2[13] = new Array();")	
		while (!rsCaseManager.EOF){ 
			chrData = rsCaseManager("chvName")
			intID = rsCaseManager("insId")
			Response.Write("detailData2[13]["+intID+"] = '"+chrData+"';");
			rsCaseManager.MoveNext 
		}
	} else {
	   Response.Write("Case Manager lookup does not exist.")
	}
	rsCaseManager.Close();

	var rsDisability = Server.CreateObject("ADODB.Recordset");
	rsDisability.ActiveConnection = MM_cnnASP02_STRING;
	rsDisability.Source = "{call dbo.cp_AC_StdDsbty(0,0,0)}";
	rsDisability.CursorType = 0;
	rsDisability.CursorLocation = 2;
	rsDisability.LockType = 3;
	rsDisability.Open();
	
	if (!rsDisability.EOF){ 
		Response.Write("detailData2[2] = new Array();")	
		while (!rsDisability.EOF){ 
			chrData = rsDisability("chvname")
			intID   = rsDisability("insDisability_id")
			Response.Write("detailData2[2]["+intID+"] = '"+chrData+"';");
			rsDisability.MoveNext 
		}
	} else {
		Response.Write("Disability lookup does not exist.")
	}
	rsDisability.Close();

	var rsStatus = Server.CreateObject("ADODB.Recordset");
	rsStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsStatus.Source = "{call dbo.cp_AC_StdStatus(0,0,0)}";
	rsStatus.CursorType = 0;
	rsStatus.CursorLocation = 2;
	rsStatus.LockType = 3;
	rsStatus.Open();	
	if (!rsStatus.EOF){ 
		Response.Write("detailData2[3] = new Array();")
		while (!rsStatus.EOF){ 
			chrData = rsStatus("chvName")
			intID = rsStatus("insStdnt_status_id")
			Response.Write("detailData2[3]["+intID+"] = '"+chrData+"';");
			rsStatus.MoveNext 
		}
	} else {
		Response.Write("Student Status lookup does not exist.")
	}
	rsStatus.Close();

	var rsReferringAgent = Server.CreateObject("ADODB.Recordset");
	rsReferringAgent.ActiveConnection = MM_cnnASP02_STRING;
	rsReferringAgent.Source = "{call dbo.cp_referring_agent(0,0,0,0,0)}";
	rsReferringAgent.CursorType = 0;
	rsReferringAgent.CursorLocation = 2;
	rsReferringAgent.LockType = 3;
	rsReferringAgent.Open();
	
	if (!rsReferringAgent.EOF){ 
		Response.Write("detailData2[4] = new Array();")	
		while (!rsReferringAgent.EOF){ 
			chrData = rsReferringAgent("chvname")
			intID   = rsReferringAgent("insrefer_agent_id")
			Response.Write("detailData2[4]["+intID+"] = '"+chrData+"';");
			rsReferringAgent.MoveNext 
		}
	} else {
		Response.Write("Referring Agent lookup does not exist.")
	}
	rsReferringAgent.Close();

	var rsFundingSource = Server.CreateObject("ADODB.Recordset");
	rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
	rsFundingSource.Source = "{call dbo.cp_funding_source_attributes(0,0,0,0,1,0,0,0,2,'Q',0)}";
	rsFundingSource.CursorType = 0;
	rsFundingSource.CursorLocation = 2;
	rsFundingSource.LockType = 3;
	rsFundingSource.Open();
	
	if (!rsFundingSource.EOF){ 
		Response.Write("detailData2[5] = new Array();")	
		while (!rsFundingSource.EOF){ 
			chrData = rsFundingSource("chvfunding_source_name")
			intID   = rsFundingSource("insFunding_source_id")
			Response.Write("detailData2[5]["+intID+"] = '"+chrData+"';");
			rsFundingSource.MoveNext 
		}
	} else {
		Response.Write("Funding Source lookup does not exist.")
	}
	rsFundingSource.Close();

	var rsServiceType = Server.CreateObject("ADODB.Recordset");
	rsServiceType.ActiveConnection = MM_cnnASP02_STRING;
	rsServiceType.Source = "{call dbo.cp_service_type(0,0,0,0)}";
	rsServiceType.CursorType = 0;
	rsServiceType.CursorLocation = 2;
	rsServiceType.LockType = 3;
	rsServiceType.Open();	
	if (!rsServiceType.EOF){ 
		Response.Write("detailData2[6] = new Array();")	
		while (!rsServiceType.EOF) { 
			chrData = rsServiceType("chvname")
			intID = rsServiceType("insService_type_id")
			Response.Write("detailData2[6]["+intID+"] = '"+chrData+"';");
			rsServiceType.MoveNext 
		}
	} else {
		Response.Write("Service Code lookup does not exist.")
	}
	rsServiceType.Close();

	var rsRegion2 = Server.CreateObject("ADODB.Recordset");
	rsRegion2.ActiveConnection = MM_cnnASP02_STRING;
	rsRegion2.Source = "{call dbo.cp_AC_Region(0,1,0)}";
	rsRegion2.CursorType = 0;
	rsRegion2.CursorLocation = 2;
	rsRegion2.LockType = 3;
	rsRegion2.Open();	
	if (!rsRegion2.EOF){ 	
		Response.Write("detailData2[7] = new Array();")	
		while (!rsRegion2.EOF){ 
			chrData = rsRegion2("chvname")
			intID   = rsRegion2("insRegion_num")
			Response.Write("detailData2[7]["+intID+"] = '"+ chrData+"';");
			rsRegion2.MoveNext 
		}
	} else {
		Response.Write("Region lookup does not exist.")
	}
	rsRegion2.Close();
	
	var rsDisability2 = Server.CreateObject("ADODB.Recordset");
	rsDisability2.ActiveConnection = MM_cnnASP02_STRING;
	rsDisability2.Source = "{call dbo.cp_AC_StdDsbty(0,1,0)}";
	rsDisability2.CursorType = 0;
	rsDisability2.CursorLocation = 2;
	rsDisability2.LockType = 3;
	rsDisability2.Open();	
	if (!rsDisability2.EOF){ 	
		Response.Write("detailData2[8] = new Array();")	
		while (!rsDisability2.EOF){
			chrData = rsDisability2("chvname")
			intID   = rsDisability2("insDisability_id")
			Response.Write("detailData2[8]["+intID+"] = '"+chrData+"';");
			rsDisability2.MoveNext 
		}
	} else {
		Response.Write("Disability lookup does not exist.")
	}
	rsDisability2.Close();	
	%>
	
	var objParent ;
	var objChild  ;

	function initializeList(parentList, childList){
		objParent = eval("document." + parentList);
		objChild = eval("document." + childList);
		if (objParent.selectedIndex < 0) objParent.options[0].selected = true;
		var OptnID = objParent.options[objParent.selectedIndex].value;	
		OptnID -= 26 ;	
		if ((OptnID == 1) || (OptnID == 2)) {
			oOld3.style.visibility="visible";
			document.frm01s01.IncludeHistoricValues.disabled = false;
			if (document.frm01s01.IncludeHistoricValues.checked) OptnID += 6;
		} else {
			oOld3.style.visibility="hidden";
			document.frm01s01.IncludeHistoricValues.disabled = true;
		}
		clearList(objChild)
		fillDetail(OptnID)
	}

	function fillDetail(deptID) {
		if (detailData2[deptID]) {
			for(x=0; x < detailData2[deptID].length; x++) {
				if (detailData2[deptID][x]) {
					var sDesc = detailData2[deptID][x];
					var valOption = new Option(sDesc);
					valOption.value = x;
					objChild.options[objChild.length] = valOption;
				}
			}
			if (objChild.length) objChild.options[0].selected = true;
		}
		objChild = null;
		objParent = null;	
	}

	function clearList(obj){
		if (obj.length) obj.options.length = 0;
	}
	
	if (window.focus) self.focus();	
	</script>
</head>
<body onload="initscr();">
<form name="frm01s01" method="post" action="">
<h3>Client - Advanced Search</h3>
<hr>
<table cellpadding="1" cellspacing="1" border="0">
    <tr>
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="1" checked onClick="SelOpt()" accesskey="F" tabindex="1" class="chkstyle">String Search</td>
		<td nowrap valign="top"><DIV ID="oOptrd1" STYLE="visibility:visible">
			<select name="StringSearchOperand" onChange="selectChange(this, frm01s01.StringSearchOperator,detailData1);" tabindex="2">
			<% 
			while (!rsOprd1.EOF) {
			%>
				<option value="<%=(rsOprd1.Fields.Item("intRecID").Value)%>" <%=((rsOprd1.Fields.Item("intRecID").Value == Request.Form("MM_curOprd"))?"SELECTED":"")%>><%=(rsOprd1.Fields.Item("chvObjName").Value)%></option>
			<%
				rsOprd1.MoveNext();
			}
			%>
			</select>
			<input type='checkbox' name='IsNot' value='1' tabindex="3" class="chkstyle">Not 			
		</DIV></td>
		<td nowrap valign="top"><DIV ID="oStg1" STYLE="visibility:visible">
			<select name="StringSearchOperator" onChange="Togo();" style="width:130px" tabindex="4"></select>			
			<input type="text" name="StringSearchTextOne" tabindex="5">
		</DIV></td>
		<td nowrap><DIV ID="oStg2" STYLE="visibility:hidden">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
			<input type="text" name="StringSearchTextTwo" value="<%=CurrentDate()%>" tabindex="6" DISABLED>
           	<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="2" onClick="SelOpt()" tabindex="7" class="chkstyle">Lookup Value Search</td>
		<td nowrap><DIV ID="oOptrd31" STYLE="visibility:visible">
			<select name="LookupValueSearchOperand" onchange="selectChange(this, frm01s01.LookupValueSearchOptions,detailData2);" tabindex="8">
			<% 
			while (!rsOprd2.EOF) { 
			%>
				<option value="<%=(rsOprd2.Fields.Item("intRecID").Value)%>"><%=(rsOprd2.Fields.Item("chvObjName").Value)%></option>
			<% 
				rsOprd2.MoveNext();
			}
			%>
			</select>
		</DIV></td>
		<td nowrap><DIV ID="oStg21" STYLE="visibility:visible"> 
			Is <select name="LookupValueSearchOptions" tabindex="9"></select>
        </DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1" border="0">
    <tr> 
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="3" onClick="SelOpt()" tabindex="10" class="chkstyle">Multi-Value Search</td>
		<td nowrap valign="top"><DIV ID="oOptrd3" STYLE="visibility:visible"> 
			<select name="MultiValueSearchOperand" OnChange="initializeList('frm01s01.MultiValueSearchOperand','frm01s01.MultiValueSearchText')" tabindex="11">
			<% 
			while (!rsOprd3.EOF) {
			%>
				<option value="<%=(rsOprd3.Fields.Item("intRecID").Value)%>" <%=((rsOprd3.Fields.Item("intRecID").Value == Request.Form("MM_curOprd"))?"SELECTED":"")%>><%=(rsOprd3.Fields.Item("chvObjName").Value)%></option>
			<%
				rsOprd3.MoveNext();
			}
			%>
			</select>
			<DIV ID="oOld3" STYLE="visibility:hidden"><input type="checkbox" name="IncludeHistoricValues" value="1" onClick="Selckb()" tabindex="12" class="chkstyle">Include Historic Values</DIV>
		</DIV></td>
		<td nowrap valign="top"><DIV ID="oOptr3" STYLE="visibility:hidden"> 
			contains: <select name="MultiValueSearchText" size="8" width="150" align="top" multiple tabindex="13"></select>
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td valign="top"><select name="AndOr" tabindex="14">
			<option value=" ">None</option>
			<option value="And">And</option>
			<option value="Or">Or</option>
		</select></td>	
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td valign="top" nowrap><input type="button" name="AddToQueryString" value="Add" tabindex="15" onClick="Savtxt()" class="btnstyle"></td>
		<td valign="top" nowrap><textarea name="QueryString" cols="80" rows="6" tabindex="16" accesskey="L"></textarea></td>
    </tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
	<tr>
      <td nowrap>Sort by:
			<select name="SortByColumn" tabindex="17">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("insObjOrder").Value == 1)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>
			</select>
        	<select name="OrderBy" tabindex="18">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>
	        <input type="button" value="Search" onClick="CnstrFltr(1);" tabindex="19" class="btnstyle">
			<input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="20" class="btnstyle">
			<input type="button" value="Clear All" onClick="window.location.reload();" tabindex="21" class="btnstyle">
		</td>		
    </tr>
</table>
<input type="hidden" name="MM_flag" value="false">
<input type="hidden" name="MM_curOprd">
<input type="hidden" name="MM_curOptr">
</form>
</body>
</html>
<%
rsOprd3.Close();
rsOprd2.Close();
rsOprd1.Close();
rsCol.Close();
%>