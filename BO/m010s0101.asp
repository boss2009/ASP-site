<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
// retrieve sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup2(758,0,'0',0,'0',0)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

// retrieve text search operands
var rsOprd = Server.CreateObject("ADODB.Recordset");
rsOprd.ActiveConnection = MM_cnnASP02_STRING;
rsOprd.Source = "{call dbo.cp_ASP_Lkup2(759,0,'0',0,'0',0)}";
rsOprd.CursorType = 0;
rsOprd.CursorLocation = 2;
rsOprd.LockType = 3;
rsOprd.Open();

// retrieve lookup search operands
var rsOprd2 = Server.CreateObject("ADODB.Recordset");
rsOprd2.ActiveConnection = MM_cnnASP02_STRING;
rsOprd2.Source = "{call dbo.cp_ASP_Lkup2(760,0,'0',0,'0',0)}";
rsOprd2.CursorType = 0;
rsOprd2.CursorLocation = 2;
rsOprd2.LockType = 3;
rsOprd2.Open();

// retrieve multi search operands
var rsOprd3 = Server.CreateObject("ADODB.Recordset");
rsOprd3.ActiveConnection = MM_cnnASP02_STRING;
rsOprd3.Source = "{call dbo.cp_ASP_Lkup2(761,0,'0',0,'0',0)}";
rsOprd3.CursorType = 0;
rsOprd3.CursorLocation = 2;
rsOprd3.LockType = 3;
rsOprd3.Open();
%>
<html>
<head>
	<title>Buyout - Quick Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="Javascript">
	var detailData = new Array();
	<%
	var intOldOptr = 0;
	var objOptrDesc,objOptrId,objRecID;
	
	var rsOptr = Server.CreateObject("ADODB.Recordset");
	rsOptr.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,25)}";	
	rsOptr.CursorType = 0;
	rsOptr.CursorLocation = 2;
	rsOptr.LockType = 3;
	rsOptr.Open();
	// Load the Operators Lookup 
	if (!rsOptr.EOF){ 	
		while (!rsOptr.EOF) { 
			objRecID    = rsOptr("intRecID")	
			if (intOldOptr != objRecID.value) {
				Response.Write("detailData["+objRecID+"] = new Array();")
				intOldOptr = objRecID.value
			}
	%>
			detailData[<%=objRecID%>][<%=rsOptr("intOptrId")%>] = "<%=rsOptr("chvOptrDesc")%>"	
	<%
			rsOptr.MoveNext 
		}
	} else {
	   Response.Write("SysOptr lookup does not exist.")
	}	
	rsOptr.Close();
	%>

	var Grp4Data   = new Array();
	<%
	// retrieve the Buyout Process lookup
	var rsBuyoutProcess = Server.CreateObject("ADODB.Recordset");
	rsBuyoutProcess.ActiveConnection = MM_cnnASP02_STRING;
	rsBuyoutProcess.Source = "{call dbo.cp_buyout_process(0,'',1,0,'Q',0)}";
	rsBuyoutProcess.CursorType = 0;
	rsBuyoutProcess.CursorLocation = 2;
	rsBuyoutProcess.LockType = 3;
	rsBuyoutProcess.Open();
	if (!rsBuyoutProcess.EOF){ 	
		Response.Write("Grp4Data[238] = new Array();")	
		while (!rsBuyoutProcess.EOF) { 
	%>
			Grp4Data[238][<%=rsBuyoutProcess("insBuyout_process_id")%>] = "<%= rsBuyoutProcess("chvBuyout_process") %>"
	<%
			rsBuyoutProcess.MoveNext 
		}
	} else {
		Response.Write("Buyout Process lookup does not exist.")
	}	
	rsBuyoutProcess.Close();

	// retrieve the Funding Source lookup
	var rsFundingSource = Server.CreateObject("ADODB.Recordset");
	rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
	rsFundingSource.Source = "{call dbo.cp_funding_source2(0,'',0,0,0,0,1,0,0,2,'Q',0)}";
	rsFundingSource.CursorType = 0;
	rsFundingSource.CursorLocation = 2;
	rsFundingSource.LockType = 3;
	rsFundingSource.Open();
	if (!rsFundingSource.EOF){ 	
		Response.Write("Grp4Data[237] = new Array();")	
		while (!rsFundingSource.EOF) { 
			chrData = rsFundingSource("chvfunding_source_name")
			intID   = rsFundingSource("insFunding_source_id")
	%>
			Grp4Data[237][<%=intID%>] = "<%= chrData %>"
	<%
			rsFundingSource.MoveNext 
		}
	} else {
		Response.Write("Funding Source lookup does not exist.")
	}	
	rsFundingSource.Close();

	// retrieve the Buyout Status lookup
	var rsBuyoutStatus = Server.CreateObject("ADODB.Recordset");
	rsBuyoutStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsBuyoutStatus.Source = "{call dbo.cp_buyout_status(0,'',0,'Q',0)}";
	rsBuyoutStatus.CursorType = 0;
	rsBuyoutStatus.CursorLocation = 2;
	rsBuyoutStatus.LockType = 3;
	rsBuyoutStatus.Open();
	if (!rsBuyoutStatus.EOF){ 	
		Response.Write("Grp4Data[236] = new Array();")	
		while (!rsBuyoutStatus.EOF) { 
			chrData = rsBuyoutStatus("chvBuyout_status")
			intID   = rsBuyoutStatus("insbuyout_status_id")
	%>
			Grp4Data[236][<%=intID%>] = "<%= chrData %>"
	<%
			rsBuyoutStatus.MoveNext 
		}
	} else {
		Response.Write("Buyout status lookup does not exist.")
	}	
	rsBuyoutStatus.Close();

	// retrieve the Case Manager lookup
	var rsCaseManager = Server.CreateObject("ADODB.Recordset");
	rsCaseManager.ActiveConnection = MM_cnnASP02_STRING;
	rsCaseManager.Source = "{call dbo.cp_CaseMgr}";
	rsCaseManager.CursorType = 0;
	rsCaseManager.CursorLocation = 2;
	rsCaseManager.LockType = 3;
	rsCaseManager.Open();
	if (!rsCaseManager.EOF){ 	
		Response.Write("Grp4Data[239] = new Array();")	
		while (!rsCaseManager.EOF) { 
			chrData = rsCaseManager("chvName")
			intID   = rsCaseManager("insId")
	%>
			Grp4Data[239][<%=intID%>] = "<%= chrData %>"
	<%
			rsCaseManager.MoveNext 
		}
	} else {
		Response.Write("Case Manager lookup does not exist.")
	}	
	rsCaseManager.Close();

	// retrieve the Institution lookup
	var rsInstitution = Server.CreateObject("ADODB.Recordset");
	rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
	rsInstitution.Source = "{call dbo.cp_school2(0,'',0,0,0,0,0,1,0,'',2,'Q',0)}";
	rsInstitution.CursorType = 0;
	rsInstitution.CursorLocation = 2;
	rsInstitution.LockType = 3;
	rsInstitution.Open();
	if (!rsInstitution.EOF){ 	
		Response.Write("Grp4Data[235] = new Array();")	
		while (!rsInstitution.EOF) { 
			chrData = rsInstitution("chvSchool_Name")
			intID   = rsInstitution("insSchool_id")
	%>
			Grp4Data[235][<%=intID%>] = "<%= chrData %>"
	<%
			rsInstitution.MoveNext 
		}
	} else {
		Response.Write("Institution lookup does not exist.")
	}	
	rsInstitution.Close();
	%>
	
//	Grp4Data[188] = new Array();
//	Grp4Data[188][0] = "Not Available";
//	Grp4Data[188][1] = "Training Requested";

//	Grp4Data[189] = new Array();
//	Grp4Data[189][0] = "Unable to Arrange";
//	Grp4Data[189][1] = "Declined";
//	Grp4Data[189][2] = "Completed";	

//	Grp4Data[190] = new Array();
//	Grp4Data[190][0] = "Unable to Arrange";
//	Grp4Data[190][1] = "Declined";
//	Grp4Data[190][2] = "Completed";	
	
	var RegionArray   = new Array(2);
	<%
	// server side variables
	var intRegionCnt = 0 ;
	// retrieve the Region lookup
	var rsRegion = Server.CreateObject("ADODB.Recordset");
	rsRegion.ActiveConnection = MM_cnnASP02_STRING;
	rsRegion.Source = "{call dbo.cp_ac_region(0,1,0)}";
	rsRegion.CursorType = 0;
	rsRegion.CursorLocation = 2;
	rsRegion.LockType = 3;
	rsRegion.Open();
	while (!rsRegion.EOF){
		intRegionCnt++;
		rsRegion.MoveNext;
	}
	rsRegion.MoveFirst;
	// Load the Region Lookup 
	if (!rsRegion.EOF){ 	
	%>
		var RegionArraySize = <%=intRegionCnt%>;
		for (var i=0; i < <%=intRegionCnt%>; i++){
			RegionArray[i] = new Array(<%=intRegionCnt%>);
		}
	<%	   
		intRegionCnt = 0;
		while (!rsRegion.EOF) { 
	%>
			RegionArray[<%=intRegionCnt%>][0] = "<%=rsRegion("insRegion_num")%>"
			RegionArray[<%=intRegionCnt%>][1] = "<%=rsRegion("chvname")%>"
	<%
			intRegionCnt += 1
			rsRegion.MoveNext 
		}
	} else {
	   Response.Write("Region lookup does not exist.")
	}	
	rsRegion.Close();
	%>
	var DisabilityArray   = new Array(2);
	<%
	// server side variables
	var intDisabilityCnt = 0 ;
	// retrieve the Disability lookup
	var rsDisability = Server.CreateObject("ADODB.Recordset");
	rsDisability.ActiveConnection = MM_cnnASP02_STRING;
	rsDisability.Source = "{call dbo.cp_AC_stddsbty(0,0,0)}";
	rsDisability.CursorType = 0;
	rsDisability.CursorLocation = 2;
	rsDisability.LockType = 3;
	rsDisability.Open();
	while (!rsDisability.EOF){
		intDisabilityCnt++;
		rsDisability.MoveNext;
	}
	rsDisability.MoveFirst;
	// Load the Disability Lookup 
	if (!rsDisability.EOF){ 	
	%>
		var DisabilityArraySize = <%=intDisabilityCnt%>;
		for (var i=0; i < <%=intDisabilityCnt%>; i++){
			DisabilityArray[i] = new Array(<%=intDisabilityCnt%>);
		}
	<%	   
		intDisabilityCnt = 0;
		while (!rsDisability.EOF) { 
	%>
			DisabilityArray[<%=intDisabilityCnt%>][0] = "<%=rsDisability("insDisability_id")%>"
			DisabilityArray[<%=intDisabilityCnt%>][1] = "<%=rsDisability("chvname")%>"
	<%
			intDisabilityCnt += 1
			rsDisability.MoveNext 
		}
	} else {
		Response.Write("Disability lookup does not exist.")
	}	
	rsDisability.Close();
	%>

	var ReferralTypeArray   = new Array(2);
	<%
	var intReferralTypeCnt = 0 ;
	// retrieve the Referral Type lookup
	var rsReferralType = Server.CreateObject("ADODB.Recordset");
	rsReferralType.ActiveConnection = MM_cnnASP02_STRING;
	rsReferralType.Source = "{call dbo.cp_referring_agent(0,1,0,1,0)}";
	rsReferralType.CursorType = 0;
	rsReferralType.CursorLocation = 2;
	rsReferralType.LockType = 3;
	rsReferralType.Open();
	while (!rsReferralType.EOF){
		intReferralTypeCnt++;
		rsReferralType.MoveNext;
	}
	rsReferralType.MoveFirst;
	// Load the Referral Type Lookup 
	if (!rsReferralType.EOF){ 	
	%>
		var ReferralTypeArraySize = <%=intReferralTypeCnt%>;	
		for (var i=0; i < <%=intReferralTypeCnt%>; i++){
			ReferralTypeArray[i] = new Array(<%=intReferralTypeCnt%>);
		}
	<%	   
		intReferralTypeCnt = 0;
		while (!rsReferralType.EOF) { 
	%>
			ReferralTypeArray[<%=intReferralTypeCnt%>][0] = "<%=rsReferralType("insrefer_agent_id")%>"
			ReferralTypeArray[<%=intReferralTypeCnt%>][1] = "<%=rsReferralType("chvname")%>"
	<%
			intReferralTypeCnt += 1
			rsReferralType.MoveNext 
		}
	} else {
		Response.Write("Referral Type lookup does not exist.")
	}	
	rsReferralType.Close();
	%>

	function selectChange(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		for (var q=document.frm10s01.LookupValueSearchOptions.options.length;q>=0;q--) document.frm10s01.LookupValueSearchOptions.options[q]=null;	  
		myEle = document.createElement("option") ;  
		var y = control.value;
		
		switch (y){
			case "231":
				document.frm10s01.StringSearchTextboxOne.style.visibility = "hidden";
				oStg12.style.visibility="hidden";
			break;
			default :
				if (document.frm10s01.SearchType[0].checked==true) document.frm10s01.StringSearchTextboxOne.style.visibility = "visible";			
				document.frm10s01.StringSearchTextboxTwo.disabled = true;
				oStg12.style.visibility="hidden";
			break;
		}
		
		document.frm10s01.StringSearchTextboxOne.value = "";
		
		if (y != 0) {
			for ( x = 0 ; x < ItemArray[y].length  ; x++ ){
				if (ItemArray[y][x]) {
					myEle = document.createElement("option") ;
					myEle.value = x ;
					myEle.text = ItemArray[y][x] ;
					controlToPopulate.add(myEle) ;
				}
			}
		}
		// ---------------------------------
		// construct param
		// ---------------------------------
		var j = 0;
		var len = document.frm10s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm10s01.SearchType[i].checked) j = i;
		}				
		if (j==1) selectChange4(frm10s01.LookupValueSearchOperator, frm10s01.LookupValueSearchOptions,Grp4Data );
		Togo();	 
	}

	function selectChange4(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		for (var q=document.frm10s01.LookupValueSearchOptions.options.length;q>=0;q--) document.frm10s01.LookupValueSearchOptions.options[q]=null;	  	  
		myEle = document.createElement("option") ;
		var y = document.frm10s01.LookupValueSearchOperand.value;
		if (y != 0) {
			for ( x = 0 ; x < ItemArray[y].length  ; x++ ){
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
		var objTmp ;
		if (document.frm10s01.StringSearchOperand.selectedIndex <= 0) {
			var j = 0;
		} else {
			var j = document.frm10s01.StringSearchOperand[document.frm10s01.StringSearchOperand.selectedIndex].value ;
		}
		if (document.frm10s01.StringSearchOperator.selectedIndex <= 0) {
			var l = 0;
		} else {
			var l = document.frm10s01.StringSearchOperator[document.frm10s01.StringSearchOperator.selectedIndex].value ;
		}
		document.frm10s01.MM_curOprd.value = j ;
		document.frm10s01.MM_curOptr.value = l ;
		document.frm10s01.MM_flag.value = true ;
	}
	
	</script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript" src="../js/m010Srh01.js"></script>
	<script language="JavaScript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	   
	
	if (window.focus) self.focus();

	function initscr() {
		oStg12.style.visibility="hidden";
		oOptrd21.style.visibility="hidden";
		oOptr21.style.visibility="hidden";
		oStg21.style.visibility="hidden";
		oOptrd31.style.visibility="hidden";
		oOptr31.style.visibility="hidden";	  
		initializeList(document.frm10s01.MultiSelectOperand,document.frm10s01.MultiSelectOptions);
		var j = 0;
		var len = document.frm10s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm10s01.SearchType[i].checked) j = i;
		}
		switch (j) {
			case 0:
				selectChange(frm10s01.StringSearchOperand, frm10s01.StringSearchOperator,detailData);
			break;
			case 1:
				selectChange(frm10s01.LookupValueSearchOperand, frm10s01.LookupValueSearchOperator,detailData );
			break;
		}	   				
	}
	
	function addOption(txt, val){
		var oOption=document.createElement("OPTION");
		oOption.text = txt;
		oOption.value = val;
		document.frm10s01.MultiSelectOptions.add(oOption);
	}
	
	function initializeList(oParent, oChild){
	  	while (oChild.length > 0){
    		oChild.remove(0);
  		}
		switch (oParent.selectedIndex) {
			case 0:
				for (var i=0; i< RegionArraySize; i++) {		
					addOption(RegionArray[i][1],RegionArray[i][0]);
				}
			break;
			case 1:
				for (var i=0; i< DisabilityArraySize; i++) {		
					addOption(DisabilityArray[i][1],DisabilityArray[i][0]);
				}			
			case 2:
				for (var i=0; i< ReferralTypeArraySize; i++) {		
					addOption(ReferralTypeArray[i][1],ReferralTypeArray[i][0]);
				}			
			break;
		}
	}	

	function CnstrFltr(output) {	
		var stgPgQuery = "";
		var stgFilter = "" ;
		var j,k; 		

		var len = document.frm10s01.SearchType.length;
		var Idparam = 0;
		
		for (var i = 0; i < len; i++){
			if (document.frm10s01.SearchType[i].checked) Idparam = i;		
		}
		
		switch ( Idparam ) {
			case 0: 
				if (document.frm10s01.StringSearchOperand.length >= 1) {
					var chvOprd = document.frm10s01.StringSearchOperand[document.frm10s01.StringSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select String Search Operand.");
					break;
				}
				var chrNot  = "";
				if (document.frm10s01.StringSearchOperator.length >= 1) {
					var chvOptr = document.frm10s01.StringSearchOperator[document.frm10s01.StringSearchOperator.selectedIndex].value ;
				} else {
					alert("Select String Search Operator.");
					break;
				}
				var chvStg1 = document.frm10s01.StringSearchTextboxOne.value;
				var chvStg2 = document.frm10s01.StringSearchTextboxTwo.value;
				if (chvOptr == "0") {
					alert("Please select operator before Proceed");
				} else {
					stgFilter = ACfltr_10(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
				}
			break;
			case 1: 
		  		if (document.frm10s01.LookupValueSearchOperand.length >= 1) {	  
					var chvOprd = document.frm10s01.LookupValueSearchOperand[document.frm10s01.LookupValueSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select Lookup Value Search Operand.");
					break;			
				}
				 var chrNot  = "";
				if (document.frm10s01.LookupValueSearchOperator.length >= 1) {	  			 
					var chvOptr = document.frm10s01.LookupValueSearchOperator[document.frm10s01.LookupValueSearchOperator.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Operator.");
					break;			
				}
				if (document.frm10s01.LookupValueSearchOptions.length >= 1) {	  			 			
					var chvStg1 = document.frm10s01.LookupValueSearchOptions[document.frm10s01.LookupValueSearchOptions.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Option.");
					break;
				}
				var chvStg2 = "";
				if (chvOptr == "0") {
					alert("Please select operator before Proceed");
				} else {
					stgFilter = ACfltr_10(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
				}
			break;
			case 2: 
				j = document.frm10s01.MultiSelectOperand[document.frm10s01.MultiSelectOperand.selectedIndex].value ;
				l = "";  
				var optList = document.frm10s01.MultiSelectOptions;
				var	m = optList.length;
				if (optList.multiple) {
					for(var ii = 0; ii < m; ii++) {
						if (document.frm10s01.MultiSelectOptions[ii].selected) {
							if (l.length > 0 ) l += "," ;
							l += document.frm10s01.MultiSelectOptions[ii].value ;	
						} 				  
					} 	
				} else {
					l = document.frm10s01.MultiSelectOptions[document.frm10s01.MultiSelectOptions.selectedIndex].value
				} 
				if (l=="") {
					alert("Select at least one Multi Value Search Option.");
					break;			
				}
				// Construct filters for multi-items select
				switch (j) {
					// Disability
					case "241" :
						stgFilter += " insDsbty_id in (" + l + ") " ; 
					break;
					// Region
					case "240" : 
						stgFilter += " insRegion_num in (" + l + ") " ; 
					break;
					// Referral Type
					case "242" : 
						stgFilter += " insref_agt_id in (" + l + ") " ; 
					break;
				}
			default: 
			break;
		}
		var inspSrtBy = document.frm10s01.SortByColumn.value;
		var inspSrtOrd = document.frm10s01.OrderBy.value;
		if (output==1) {
			document.frm10s01.action = "m010q02.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;		
		} else {
			document.frm10s01.action = "m010q02excel.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter;
			document.frm10s01.target = "_blank";
			document.frm10s01.submit();
		}
	}
	
	function SelOpt() {	
		var len = document.frm10s01.SearchType.length;
		var Idparam = 1;
	
		for (var i=0;i <len; i++){
			if (document.frm10s01.SearchType[i].checked) Idparam = i;
		}
		
		switch (Idparam) {
			case 0: 
				oOptrd11.style.visibility="visible";
				oOptr11.style.visibility="visible";
				oStg11.style.visibility="visible";
				oStg12.style.visibility="hidden";
				document.frm10s01.StringSearchTextboxOne.style.visibility = "visible";			   				
				oOptrd21.style.visibility="hidden";
				oOptr21.style.visibility="hidden";
				oStg21.style.visibility="hidden";
				oOptrd31.style.visibility="hidden";
				oOptr31.style.visibility="hidden";
				selectChange(frm10s01.StringSearchOperand, frm10s01.StringSearchOperator,detailData);
			break;
			case 1: 
				oOptrd11.style.visibility="hidden";
				oOptr11.style.visibility="hidden";
				oStg11.style.visibility="hidden";
				oStg12.style.visibility="hidden";
				document.frm10s01.StringSearchTextboxOne.style.visibility = "hidden";			   
				oOptrd21.style.visibility="visible";
				oOptr21.style.visibility="visible";
				oStg21.style.visibility="visible";
				oOptrd31.style.visibility="hidden";
				oOptr31.style.visibility="hidden";
				selectChange(frm10s01.LookupValueSearchOperand, frm10s01.LookupValueSearchOperator,detailData );							   
			break;
			case 2: 
				oOptrd11.style.visibility="hidden";
				oOptr11.style.visibility="hidden";
				oStg11.style.visibility="hidden";
				oStg12.style.visibility="hidden";
				document.frm10s01.StringSearchTextboxOne.style.visibility = "hidden";			   
				oOptrd21.style.visibility="hidden";
				oOptr21.style.visibility="hidden";
				oStg21.style.visibility="hidden";
				oOptrd31.style.visibility="visible";
				oOptr31.style.visibility="visible";			   
			break; 
			default: 
			break;
		}       
	}

	if (window.focus) self.focus();		
	</script>
</head>
<body onload="initscr()">
<form name="frm10s01" method="post" action="">
<h3>Buyout - Quick Search</h3>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="1" tabindex="1" accesskey="F" checked onClick="SelOpt()" class="chkstyle">String Search</td>
		<td nowrap><DIV ID="oOptrd11" STYLE="visibility:visible" > 
			<select name="StringSearchOperand" onchange="selectChange(this, frm10s01.StringSearchOperator,detailData);" tabindex="2">
			<% 
			while (!rsOprd.EOF) { 
			%>
				<option value="<%=(rsOprd.Fields.Item("intRecID").Value)%>" <%=((rsOprd.Fields.Item("intRecID").Value ==234)?"SELECTED":"")%> ><%=(rsOprd.Fields.Item("chvObjName").Value)%></option>
			<% 
				rsOprd.MoveNext(); 
			}
			%>
			</select>
		</DIV></td>
		<td nowrap><DIV ID="oOptr11" STYLE="visibility:visible">
			<select name="StringSearchOperator" onchange="Togo();" tabindex="3"></select>
		</DIV></td>
		<td nowrap><DIV ID="oStg11" STYLE="visibility:visible">
			<input type="text" name="StringSearchTextboxOne" size="15" tabindex="4">
		</DIV></td>
		<td nowrap><DIV ID="oStg12" STYLE="visibility:hidden">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
			<input type="text" name="StringSearchTextboxTwo" value="<%=CurrentDate()%>" size="15" tabindex="5">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="2" onClick="SelOpt()" tabindex="8" class="chkstyle">Lookup Value Search</td>
		<td nowrap><DIV ID="oOptrd21" STYLE="visibility:visible">
			<select name="LookupValueSearchOperand" onchange="selectChange(this, frm10s01.LookupValueSearchOperator,detailData);" tabindex="9">
			<% 
			while (!rsOprd2.EOF) { 
			%>
				<option value="<%=(rsOprd2.Fields.Item("intRecID").Value)%>" <%=((rsOprd2.Fields.Item("intRecID").Value == 54)?"SELECTED":"")%> ><%=(rsOprd2.Fields.Item("chvObjName").Value)%></option>
			<% 
				rsOprd2.MoveNext();
			}
			%>
			</select>
		</DIV></td>
		<td nowrap><DIV ID="oOptr21" STYLE="visibility:visible">
			<select name="LookupValueSearchOperator" onChange="selectChange4(this, frm10s01.LookupValueSearchOptions,Grp4Data );" tabindex="10"></select>
		</DIV></td>
		<td nowrap><DIV ID="oStg21" STYLE="visibility:visible">
			<select name="LookupValueSearchOptions" tabindex="11"></select>
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="3" onClick="SelOpt()" tabindex="12" class="chkstyle">Multi-Value Search</td>
		<td nowrap valign="top"><DIV ID="oOptrd31" STYLE="visibility:visible"> 
			<select name="MultiSelectOperand" OnChange="initializeList(document.frm10s01.MultiSelectOperand,document.frm10s01.MultiSelectOptions)" tabindex="13">
			<% 
			while (!rsOprd3.EOF) {
			%>
				<option value="<%=(rsOprd3.Fields.Item("intRecID").Value)%>" <%=((rsOprd3.Fields.Item("intRecID").Value == Request.Form("MM_curOprd"))?"SELECTED":"")%> ><%=(rsOprd3.Fields.Item("chvObjName").Value)%></option>
			<%
				rsOprd3.MoveNext();
			}
			%>
			</select>
		</DIV></td>		
		<td nowrap valign="top"><DIV ID="oOptr31" STYLE="visibility:hidden">
			contains: <select name="MultiSelectOptions" size="8" width="150" multiple align="top" tabindex="14"></select>
		</DIV></td>
	</tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>
			Sort by:
			<select name="SortByColumn" tabindex="15">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>"><%=(rsCol.Fields.Item("chvObjName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>
			</select>
			Order
        	<select name="OrderBy" tabindex="16">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>
	        <input type="submit" value="Search" onClick="CnstrFltr(1);" tabindex="17" class="btnstyle">
			<input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="18" class="btnstyle">
			<input type="reset" value="Clear All" onClick="window.location.reload();" tabindex="19" class="btnstyle">
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
rsOprd.Close();
rsOprd2.Close();
rsOprd3.Close();
rsCol.Close();
%>