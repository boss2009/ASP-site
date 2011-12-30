<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
// retrieve the sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup(722)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

// retrieve the text search operands - text
var rsOprd = Server.CreateObject("ADODB.Recordset");
rsOprd.ActiveConnection = MM_cnnASP02_STRING;
rsOprd.Source = "{call dbo.cp_ASP_Lkup(723)}";
rsOprd.CursorType = 0;
rsOprd.CursorLocation = 2;
rsOprd.LockType = 3;
rsOprd.Open();

// retrieve the lookup value search operands - Combo
var rsOprd2 = Server.CreateObject("ADODB.Recordset");
rsOprd2.ActiveConnection = MM_cnnASP02_STRING;
rsOprd2.Source = "{call dbo.cp_ASP_Lkup(719)}";
rsOprd2.CursorType = 0;
rsOprd2.CursorLocation = 2;
rsOprd2.LockType = 3;
rsOprd2.Open();

// retrieve the class search operands
var rsOprd3 = Server.CreateObject("ADODB.Recordset");
rsOprd3.ActiveConnection = MM_cnnASP02_STRING;
rsOprd3.Source = "{call dbo.cp_ASP_Lkup(720)}";
rsOprd3.CursorType = 0;
rsOprd3.CursorLocation = 2;
rsOprd3.LockType = 3;
rsOprd3.Open();

// retrieve the multi value search operands
var rsOprd4 = Server.CreateObject("ADODB.Recordset");
rsOprd4.ActiveConnection = MM_cnnASP02_STRING;
rsOprd4.Source = "{call dbo.cp_ASP_Lkup(721)}";
rsOprd4.CursorType = 0;
rsOprd4.CursorLocation = 2;
rsOprd4.LockType = 3;
rsOprd4.Open();
%>
<html>
<head>
	<title>Inventory Loan Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="Javascript">
	var detailData = new Array();
	<%
	var intOldOptr = 0;
	var intIdxOprd = 0;
	var objOptrDesc,objOptrId,objRecID;
	
	var rsOptr = Server.CreateObject("ADODB.Recordset");
	rsOptr.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,18)}";
	
	rsOptr.CursorType = 0;
	rsOptr.CursorLocation = 2;
	rsOptr.LockType = 3;
	rsOptr.Open();
	if (!rsOptr.EOF){ 	
		while (!rsOptr.EOF) { 
			objOptrDesc = rsOptr("chvOptrDesc")
			objOptrId   = rsOptr("intOptrId")
			objRecID    = rsOptr("intRecID")	
			if (intOldOptr != objRecID.value) {
				Response.Write("detailData["+objRecID+"] = new Array();")
				intIdxOprd += 1
				intOldOptr = objRecID.value
			}
	%>
			detailData[<%=objRecID%>][<%=objOptrId%>] = "<%= objOptrDesc %>"	
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
	// retrieve the Vendor lookup
	var rsVendor = Server.CreateObject("ADODB.Recordset");
	rsVendor.ActiveConnection = MM_cnnASP02_STRING;
	rsVendor.Source = "{call dbo.cp_ASP_Lkup(3)}";
	rsVendor.CursorType = 0;
	rsVendor.CursorLocation = 2;
	rsVendor.LockType = 3;
	rsVendor.Open();
	if (!rsVendor.EOF){ 	
		Response.Write("Grp4Data[47] = new Array();")	
		while (!rsVendor.EOF) { 
	%>
			Grp4Data[47][<%=rsVendor("intCompany_id")%>] = "<%= rsVendor("chvCompany_Name") %>"
	<%
			rsVendor.MoveNext 
		}
	} else {
		Response.Write("Vendor lookup does not exist.")
	}	
	rsVendor.Close();

	// retrieve the Institution lookup
	var rsInstitution = Server.CreateObject("ADODB.Recordset");
	rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
	rsInstitution.Source = "{call dbo.cp_school2(0,'',0,0,0,0,0,1,0,'',2,'Q',0)}";
	rsInstitution.CursorType = 0;
	rsInstitution.CursorLocation = 2;
	rsInstitution.LockType = 3;
	rsInstitution.Open();
	if (!rsInstitution.EOF){ 	
		Response.Write("Grp4Data[49] = new Array();")	
		while (!rsInstitution.EOF) { 
	%>
			Grp4Data[49][<%=rsInstitution("insSchool_id")%>] = "<%= rsInstitution("chvSchool_Name") %>"
	<%
			rsInstitution.MoveNext 
		}
	} else {
		Response.Write("Institution lookup does not exist.")
	}	
	rsInstitution.Close();

	// retrieve the Purchase Method lookup
	var rsPurchase = Server.CreateObject("ADODB.Recordset");
	rsPurchase.ActiveConnection = MM_cnnASP02_STRING;
	rsPurchase.Source = "{call dbo.cp_ASP_Lkup(102)}";
	rsPurchase.CursorType = 0;
	rsPurchase.CursorLocation = 2;
	rsPurchase.LockType = 3;
	rsPurchase.Open();
	if (!rsPurchase.EOF){ 	
		Response.Write("Grp4Data[50] = new Array();")	
		while (!rsPurchase.EOF) {
	%>
			Grp4Data[50][<%=rsPurchase("insPurchase_id")%>] = "<%= rsPurchase("chvPurchase_Method_Desc") %>"
	<%
			rsPurchase.MoveNext 
		}
	} else {
		Response.Write("Purchase Method lookup does not exist.")
	}
	
	rsPurchase.Close();
	%>
	var RegionArray   = new Array(2);
	<%
	var intRegionCnt = 0;
	// retrieve the Region lookup
	var rsRegion = Server.CreateObject("ADODB.Recordset");
	rsRegion.ActiveConnection = MM_cnnASP02_STRING;
	rsRegion.Source = "{call dbo.cp_ASP_Lkup(7)}";
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
	
	var StatusArray   = new Array(2);
	<%	
	var intID,chrData;
	var intStatusCnt = 0 ;
	// retrieve the Status lookup
	var rsStatus = Server.CreateObject("ADODB.Recordset");
	rsStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsStatus.Source = "{call dbo.cp_ASP_Lkup(36)}";
	rsStatus.CursorType = 0;
	rsStatus.CursorLocation = 2;
	rsStatus.LockType = 3;
	rsStatus.Open();
	while (!rsStatus.EOF){
		intStatusCnt++;
		rsStatus.MoveNext;
	}
	rsStatus.MoveFirst;
	// Load the Status Lookup 
	if (!rsStatus.EOF){ 	
	%>
	var StatusArraySize = <%=intStatusCnt%>;	   
		for (var i=0; i < <%=intStatusCnt%>; i++){
			StatusArray[i] = new Array(<%=intStatusCnt%>);
		}
	<%	   
		intStatusCnt = 0;
		while (!rsStatus.EOF) { 
	%>
			StatusArray[<%=intStatusCnt%>][0] = "<%=rsStatus("insEquip_status_id")%>"
			StatusArray[<%=intStatusCnt%>][1] = "<%=rsStatus("chvStatusDesc")%>"
	<%
			intStatusCnt += 1
			rsStatus.MoveNext
		}
	} else {
		Response.Write("Status lookup does not exist.")
	}
	rsStatus.Close();
	%>

	function selectChange(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		myEle = document.createElement("option") ;
		var y = control.value;
		if ((y == "44") || (y == "45") || (y == "48")){
			document.frm03s01.StringSearchTextboxTwo.disabled = false;
			oStg12.style.visibility="visible";
		} else {
			document.frm03s01.StringSearchTextboxTwo.disabled = true;
			oStg12.style.visibility="hidden";
		}
		if (y != 0) {
			for ( x = 0 ; x < ItemArray[y].length  ; x++) {
				if (ItemArray[y][x]) { 
					myEle = document.createElement("option") ;
					myEle.value = x ;
					myEle.text = ItemArray[y][x] ;
					controlToPopulate.add(myEle) ;
				}
			}
		}

		var j = 0;
		var len = document.frm03s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm03s01.SearchType[i].checked) j = i;
		}
				
		if (j==1) selectChange4(frm03s01.LookupValueSearchOperator, frm03s01.LookupValueSearchOptions,Grp4Data );
		Togo();	  
	}

	function selectChange4(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		myEle = document.createElement("option") ;
		var y = document.frm03s01.LookupValueSearchOperand.value;
		if (y != 0) {
			for ( x = 0 ; x < ItemArray[y].length  ; x++ ) {
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
		if (document.frm03s01.StringSearchOperand.selectedIndex <= 0) {
			var j = 0;
		} else {
			var j = document.frm03s01.StringSearchOperand[document.frm03s01.StringSearchOperand.selectedIndex].value ;
		}
		if (document.frm03s01.StringSearchOperator.selectedIndex <= 0) {
			var l = 0;
		} else {
			var l = document.frm03s01.StringSearchOperator[document.frm03s01.StringSearchOperator.selectedIndex].value ;
		}
		document.frm03s01.MM_curOprd.value = j ;
		document.frm03s01.MM_curOptr.value = l ;
		document.frm03s01.MM_flag.value = true ;
	}
	
	</script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript" src="../js/m003Srh03.js"></script>
	<script language="JavaScript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "");
		return ;
	}	   

	function initscr() {
		oStg12.style.visibility="hidden";
		oOptrd21.style.visibility="hidden";
		oOptr21.style.visibility="hidden";
		oStg22.style.visibility="hidden";
		oOptrd31.style.visibility="hidden";		
		oOptr31.style.visibility="hidden";
		oOptrd41.style.visibility="hidden";
		oOptr41.style.visibility="hidden";	  
		initializeList(document.frm03s01.MultiSelectOperand,document.frm03s01.MultiSelectOptions);
		var j = 0;
		var len = document.frm03s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm03s01.SearchType[i].checked) j = i;
		}
		switch (j) {
			case 0:
				selectChange(frm03s01.StringSearchOperand, frm03s01.StringSearchOperator,detailData);
			break;
			case 1:
				selectChange(frm03s01.LookupValueSearchOperand, frm03s01.LookupValueSearchOperator,detailData );
			break;
		}	   								
	}
	
	function addOption(txt, val){
		var oOption=document.createElement("OPTION");
		oOption.text = txt;
		oOption.value = val;
		document.frm03s01.MultiSelectOptions.add(oOption);
	}
	
	function initializeList(oParent, oChild){
	  	while (oChild.length > 0){
			oChild.remove(0);
  		}
		switch (oParent.selectedIndex) {
			case 0:
				for (var i=0; i< StatusArraySize; i++) {		
					addOption(StatusArray[i][1],StatusArray[i][0]);
				}
			break;
			case 1:
				for (var i=0; i< RegionArraySize; i++) {		
					addOption(RegionArray[i][1],RegionArray[i][0]);
				}			
			break;
			default :
			break;
		}
	}	

	function CnstrFltr(type) {	
		var intSrtBy = document.frm03s01.SortByColumn.value;
		var intOrdBy = document.frm03s01.OrderBy.value;
		var stgFilter = document.frm03s01.QueryString.value;
		if (type==1) {
			document.frm03s01.action = "m003r0201q.asp?inspSrtBy="+intSrtBy+"&inspSrtOrd="+intOrdBy+"&chvFilter=" + stgFilter ;
			document.frm03s01.submit() ; 						
		} else {
			openWindow("m003r0201excel.asp?inspSrtBy="+intSrtBy+"&inspSrtOrd="+intOrdBy+"&chvFilter=" + stgFilter,"");
		}
	}

	function SelOpt() {	
		var len = document.frm03s01.SearchType.length;
		var radioGrp = document.frm03s01.SearchType
		var Idparam = 1;      	
		for (var i=0;i <len; i++){
			if (document.frm03s01.SearchType[i].checked) Idparam = i;
		}
		switch ( Idparam ) {
			case 0: 
				oOptrd11.style.visibility="visible";
				oOptr11.style.visibility="visible";
				oStg11.style.visibility="visible";
				oStg12.style.visibility="hidden";
				oOptrd21.style.visibility="hidden";
				oOptr21.style.visibility="hidden";
				oStg22.style.visibility="hidden";
				oOptrd31.style.visibility="hidden";
				oOptr31.style.visibility="hidden";
				oOptrd41.style.visibility="hidden";
				oOptr41.style.visibility="hidden";
				selectChange(frm03s01.StringSearchOperand, frm03s01.StringSearchOperator,detailData);
			break;
			case 1: 
				oOptrd11.style.visibility="hidden";
				oOptr11.style.visibility="hidden";
				oStg11.style.visibility="hidden";
				oStg12.style.visibility="hidden";
				oOptrd21.style.visibility="visible";
				oOptr21.style.visibility="visible";
				oStg22.style.visibility="visible";
				oOptrd31.style.visibility="hidden";
				oOptr31.style.visibility="hidden";
				oOptrd41.style.visibility="hidden";
				oOptr41.style.visibility="hidden";
				selectChange(frm03s01.LookupValueSearchOperand, frm03s01.LookupValueSearchOperator,detailData );			   
			break;
			case 2: 
				oOptrd11.style.visibility="hidden";
				oOptr11.style.visibility="hidden";
				oStg11.style.visibility="hidden";
				oStg12.style.visibility="hidden";
				oOptrd21.style.visibility="hidden";
				oOptr21.style.visibility="hidden";
				oStg22.style.visibility="hidden";
				oOptrd31.style.visibility="visible";
				oOptr31.style.visibility="visible";
				oOptrd41.style.visibility="hidden";
				oOptr41.style.visibility="hidden";			   
			break; 
			case 3: 
				oOptrd11.style.visibility="hidden";
				oOptr11.style.visibility="hidden";
				oStg11.style.visibility="hidden";
				oStg12.style.visibility="hidden";
				oOptrd21.style.visibility="hidden";
				oOptr21.style.visibility="hidden";
				oStg22.style.visibility="hidden";
				oOptrd31.style.visibility="hidden";
				oOptr31.style.visibility="hidden";
				oOptrd41.style.visibility="visible";
				oOptr41.style.visibility="visible";			   
			break; 
			default: 
			break;
		}       
	}

	function Toggle() {	
		var idx = document.frm03s01.ClassSearchOperand[document.frm03s01.ClassSearchOperand.selectedIndex].value
		switch (idx) {
			case "39":
				openWindow("m003p01FS.asp","");
			break;
			default: 
				document.frm03s01.ClassSearchText.value = ""; 
			break;
		}
	}

	function Savtxt() {
		var stgPgQuery = "";
		var stgFilter = "";
		var blnFlg = false ;		
		var len = document.frm03s01.SearchType.length;
		var Idparam = 0;
		var stgTemp,j,k;    
		for (var i=0;i <len; i++){
			if (document.frm03s01.SearchType[i].checked) Idparam = i;
		}
		stgTemp = document.frm03s01.QueryString.value; 
		switch ( Idparam ) {
			case 0: 
				if (document.frm03s01.StringSearchOperand.length >= 1) {
					var chvOprd = document.frm03s01.StringSearchOperand[document.frm03s01.StringSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select String Search Operand.");
					break;
				}
				var chrNot  = "";
				if (document.frm03s01.StringSearchOperator.length >= 1) {
					var chvOptr = document.frm03s01.StringSearchOperator[document.frm03s01.StringSearchOperator.selectedIndex].value ;
				} else {
					alert("Select String Search Operator.");
					break;
				}
				var chvStg1 = document.frm03s01.StringSearchTextboxOne.value;
				var chvStg2 = document.frm03s01.StringSearchTextboxTwo.value;
				var chvAO1  = document.frm03s01.AndOr.value ;
				if (((chvOprd=="44") || (chvOprd=="45") || (chvOprd=="48")) && (chvOptr!="0")) {
					if (chvStg1 == "") {
						alert("Enter Start Date.");
						document.frm03s01.StringSearchTextboxOne.focus();
						return ;
					}
					if (chvStg2 == "") {
						alert("Enter End Date.");
						document.frm03s01.StringSearchTextboxTwo.focus();
					}
					if (!CheckDateBetween(Trim(chvStg1)+" and "+Trim(chvStg2))) {
						return ;
					}
				}
				if ((chvOprd=="41") || (chvOprd=="43")) {
					if (!IsID(chvStg1)) {
						alert("Invalid number.");
						document.frm03s01.StringSearchTextboxOne.focus();
					}
				}
			
				if (chvOptr == "0") {
                	alert("Please select operator before Proceed");
				} else {
					stgFilter = ACfltr_03(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
                	if (stgTemp.length > 0 ) {   stgTemp += " (" ;  } 
	                stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 ;
				}
			break;
			case 1: 
				if (document.frm03s01.LookupValueSearchOperand.length >= 1) {	  
					var chvOprd = document.frm03s01.LookupValueSearchOperand[document.frm03s01.LookupValueSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select Exact Name Search Operand.");
					break;			
				}
				var chrNot  = "";
				if (document.frm03s01.LookupValueSearchOperator.length >= 1) {	  			 
					var chvOptr = document.frm03s01.LookupValueSearchOperator[document.frm03s01.LookupValueSearchOperator.selectedIndex].value ;
				} else {
					alert("Select Exact Name Search Operator.");
					break;			
				}
	  			if (document.frm03s01.LookupValueSearchOptions.length >= 1) {	  			 			
	    			var chvStg1 = document.frm03s01.LookupValueSearchOptions[document.frm03s01.LookupValueSearchOptions.selectedIndex].value ;
				} else {
					alert("Select Exact Name Search Text.");
					break;
				}
    			var chvStg2 = "";
				var chvAO1  = document.frm03s01.AndOr.value ;
				if (chvOptr == "0") {
					alert("Please select operator before Proceed");
				} else {
					stgFilter = ACfltr_03(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
					if (stgTemp.length > 0 ) { stgTemp += " (" ; } 
					stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 ;
				}
			break;
			case 2: 
				if (document.frm03s01.ClassSearchOperand.length >= 1) {	  
					var chvOprd = document.frm03s01.ClassSearchOperand[document.frm03s01.ClassSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select Class Search Operand.");
					break;			
				}	  
				var chrNot  = "";
				var chvOptr = "3";
				var chvStg1 = document.frm03s01.ClassSearchID.value ;
				if (chvStg1 == "") {
					alert("Select Class.");
					break;
				}
				var chvStg2 = "";
				var chvAO1  = document.frm03s01.AndOr.value ;
				if (chvOptr == "0") {
					alert("Please select operator before Proceed");
				} else {
					stgFilter = ACfltr_03(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
					if (stgTemp.length > 0 ) {   stgTemp += " (" ;  } 
					stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 ;
				}
			break;
			case 3: 
				j = document.frm03s01.MultiSelectOperand[document.frm03s01.MultiSelectOperand.selectedIndex].value ;
				l = "";  
				var optList = document.frm03s01.MultiSelectOptions;
				var m = optList.length;
				if (optList.multiple) {
					for(var ii = 0; ii < m; ii++) {
						if (document.frm03s01.MultiSelectOptions[ii].selected) {
							if (l.length > 0 ) l += "," ;
							l += document.frm03s01.MultiSelectOptions[ii].value ;
						} 				  
					} 			      
				} else {
					l = document.frm03s01.MultiSelectOptions[document.frm03s01.MultiSelectOptions.selectedIndex].value
				} 
				if (l=="") {
					alert("Select at least one Multi Value Search Option.");
					break;			
				}
				switch (j) {
					// Inventory Status
					case "51" :
				    	if (stgTemp.length > 0 ) stgTemp += " (" ;
						stgFilter += " insCurrent_Status in (" + l + ") " ; 
						blnFlg     = true;
					break;
					// Region
					case "52" : 
						if (stgTemp.length > 0 ) stgTemp += " (" ;
						stgFilter += " insRegion_num in (" + l + ") " ; 
						blnFlg     = true;
					break;
					// Funding Source
					case "53" : 
						alert("option not available ...");
					break;
				}
			default: 
			break;	
		}
		if (blnFlg) {
    		var chvAO1  = document.frm03s01.AndOr.value ;
			stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 + " ";
		}

		document.frm03s01.QueryString.value = stgTemp; 
	}

	if (window.focus) self.focus();		
	</script>
</head>
<body onload="initscr()" >
<form name="frm03s01" method="post" action="">
<h5>Inventory Loan Report</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="1" tabindex="1" accesskey="F" checked onClick="SelOpt()" class="chkstyle">String Search</td>
		<td nowrap><DIV ID="oOptrd11" STYLE="visibility:visible" > 
			<select name="StringSearchOperand" onchange="selectChange(this, frm03s01.StringSearchOperator,detailData);" tabindex="2">
			<% 
			while (!rsOprd.EOF) { 
			%>
				<option value="<%=(rsOprd.Fields.Item("intRecID").Value)%>" <%=((rsOprd.Fields.Item("intRecID").Value == Request.Form("MM_curOprd"))?"SELECTED":"")%> ><%=(rsOprd.Fields.Item("chvObjName").Value)%></option>
			<% 
				rsOprd.MoveNext(); 
			}
			%>
			</select>
		</DIV></td>
		<td nowrap><DIV ID="oOptr11" STYLE="visibility:visible" > 
			<select name="StringSearchOperator" onchange="Togo();" tabindex="3"></select>
		</DIV></td>
		<td nowrap><DIV ID="oStg11" STYLE="visibility:visible">
			<input type="text" name="StringSearchTextboxOne" tabindex="4">
		</DIV></td>
		<td nowrap><DIV ID="oStg12" STYLE="visibility:hidden">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
			<input type="text" name="StringSearchTextboxTwo" value="<%=CurrentDate()%>" tabindex="5">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="2" onClick="SelOpt()" tabindex="6" class="chkstyle">Lookup Value Search</td>
		<td nowrap><DIV ID="oOptrd21" STYLE="visibility:visible">
			<select name="LookupValueSearchOperand" onchange="selectChange(this, frm03s01.LookupValueSearchOperator,detailData);" tabindex="7">
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
			<select name="LookupValueSearchOperator" onchange="selectChange4(this, frm03s01.LookupValueSearchOptions,Grp4Data );" tabindex="8"></select>
		</DIV></td>
		<td nowrap><DIV ID="oStg22" STYLE="visibility:visible"> 
			<select name="LookupValueSearchOptions" tabindex="9"></select>
        </DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="3" onClick="SelOpt()" tabindex="10" class="chkstyle">Class Search</td>
		<td nowrap><DIV ID="oOptrd31" STYLE="visibility:visible" > 
			<select name="ClassSearchOperand" tabindex="11">
			<% 
			while (!rsOprd3.EOF) {
			%>
				<option value="<%=(rsOprd3.Fields.Item("intRecID").Value)%>" <%=((rsOprd3.Fields.Item("intRecID").Value == 55)?"SELECTED":"")%> ><%=(rsOprd3.Fields.Item("chvObjName").Value)%></option>
			<%
				rsOprd3.MoveNext();
			}
			%>
			</select>
		</DIV></td>
		<td nowrap><DIV ID="oOptr31" STYLE="visibility:visible">
			is <input type="text" name="ClassSearchText" READONLY tabindex="12"><input type="button" value="List" onClick="Toggle();" tabindex="13" class="btnstyle">
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="4" onClick="SelOpt()" tabindex="14" class="chkstyle">Multi-Value Search</td>
		<td nowrap valign="top"><DIV ID="oOptrd41" STYLE="visibility:visible"> 
			<select name="MultiSelectOperand" OnChange="initializeList(document.frm03s01.MultiSelectOperand,document.frm03s01.MultiSelectOptions)" tabindex="15">
			<% 
			while (!rsOprd4.EOF) {
			%>
				<option value="<%=(rsOprd4.Fields.Item("intRecID").Value)%>" <%=((rsOprd4.Fields.Item("intRecID").Value == Request.Form("MM_curOprd"))?"SELECTED":"")%> ><%=(rsOprd4.Fields.Item("chvObjName").Value)%></option>
			<%
				rsOprd4.MoveNext();
			}
			%>
			</select>
		</DIV></td>		
		<td nowrap valign="top"><DIV ID="oOptr41" STYLE="visibility:hidden"> 
			contains: <select name="MultiSelectOptions" size="8" width="150" multiple align="top" tabindex="16"></select>
		</DIV></td>		
	</tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap valign="top"><select name="AndOr" tabindex="17">
			<option value=" ">None</option>
			<option value="And">And</option>
			<option value="Or">Or</option>
		</select></td>	
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td valign="top" nowrap><input type="button" name="AddToQueryString" value="Add" tabindex="18" onClick="Savtxt()" class="btnstyle"></td>
		<td valign="top" nowrap><textarea name="QueryString" cols="80" rows="6" tabindex="19" accesskey="L"></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Sort by:
			<select name="SortByColumn" tabindex="20">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("insObjOrder").Value == 1)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>
			</select>
			Order 
        	<select name="OrderBy" tabindex="21">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>			
	        <input type="button" value="Search" onClick="CnstrFltr(1);" tabindex="22" class="btnstyle">
			<input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="23" class="btnstyle">
			<input type="reset" value="Clear All" tabindex="24" onClick="window.location.reload();" class="btnstyle">
		</td>		
    </tr>
</table>
<input type="hidden" name="ClassSearchID">
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
rsOprd4.Close();
%>