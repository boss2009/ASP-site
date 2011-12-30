<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
// retrieve sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup2(763,0,'',0,'',0)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

// retrieve text search operands
var rsOprd = Server.CreateObject("ADODB.Recordset");
rsOprd.ActiveConnection = MM_cnnASP02_STRING;
rsOprd.Source = "{call dbo.cp_ASP_Lkup2(764,0,'',0,'',0)}";
rsOprd.CursorType = 0;
rsOprd.CursorLocation = 2;
rsOprd.LockType = 3;
rsOprd.Open();

// retrieve lookup search operands
var rsOprd2 = Server.CreateObject("ADODB.Recordset");
rsOprd2.ActiveConnection = MM_cnnASP02_STRING;
rsOprd2.Source = "{call dbo.cp_ASP_Lkup2(766,0,'',0,'0',0)}";
rsOprd2.CursorType = 0;
rsOprd2.CursorLocation = 2;
rsOprd2.LockType = 3;
rsOprd2.Open();

// retrieve class search operands
var rsOprd3 = Server.CreateObject("ADODB.Recordset");
rsOprd3.ActiveConnection = MM_cnnASP02_STRING;
rsOprd3.Source = "{call dbo.cp_ASP_Lkup2(767,0,'',0,'',0)}";
rsOprd3.CursorType = 0;
rsOprd3.CursorLocation = 2;
rsOprd3.LockType = 3;
rsOprd3.Open();

// retrieve multiple search operands
//var rsOprd4 = Server.CreateObject("ADODB.Recordset");
//rsOprd4.ActiveConnection = MM_cnnASP02_STRING;
//rsOprd4.Source = "{call dbo.cp_ASP_Lkup2(765,0,'',0,'',0)}";
//rsOprd4.CursorType = 0;
//rsOprd4.CursorLocation = 2;
//rsOprd4.LockType = 3;
//rsOprd4.Open();
%>
<html>
<head>
	<title>Equipment Service - Quick Search</title>
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
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,24)}";
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

	//Reason For Repair
	Grp4Data[281] = new Array();
	Grp4Data[281][0] = "Hardward Defect";
	Grp4Data[281][1] = "User Error";

	//Type of Repair
	Grp4Data[282] = new Array();
	Grp4Data[282][0] = "Not Covered by Warranty";
	Grp4Data[282][1] = "Covered by Warranty";

	<%
	// retrieve the Case Manager lookup
	var rsCaseManager = Server.CreateObject("ADODB.Recordset");
	rsCaseManager.ActiveConnection = MM_cnnASP02_STRING;
	rsCaseManager.Source = "{call dbo.cp_CaseMgr}";
	rsCaseManager.CursorType = 0;
	rsCaseManager.CursorLocation = 2;
	rsCaseManager.LockType = 3;
	rsCaseManager.Open();
	if (!rsCaseManager.EOF){ 	
		Response.Write("Grp4Data[280] = new Array();")	
		while (!rsCaseManager.EOF) { 
	%>
			Grp4Data[280][<%=rsCaseManager("insId")%>] = "<%= rsCaseManager("chvName") %>"
	<%
			rsCaseManager.MoveNext 
		}
	} else {
		Response.Write("Case Manager lookup does not exist.")
	}	
	rsCaseManager.Close();

	// retrieve the Repair Status lookup
	var rsRepairStatus = Server.CreateObject("ADODB.Recordset");
	rsRepairStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsRepairStatus.Source = "{call dbo.cp_repair_status(0,'',0,'Q',0)}";
	rsRepairStatus.CursorType = 0;
	rsRepairStatus.CursorLocation = 2;
	rsRepairStatus.LockType = 3;
	rsRepairStatus.Open();
	if (!rsRepairStatus.EOF){ 	
		Response.Write("Grp4Data[279] = new Array();")	
		while (!rsRepairStatus.EOF) { 
	%>
			Grp4Data[279][<%=rsRepairStatus("insEq_Repair_Sts_id")%>] = "<%= rsRepairStatus("chvEq_Repair_Sts_Desc") %>"
	<%
			rsRepairStatus.MoveNext 
		}
	} else {
	   Response.Write("Repair Status lookup does not exist.")
	}
	
	rsRepairStatus.Close();
	%>
	
	var DisabilityArray   = new Array(2);
	<%
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
	
	function selectChange(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		for (var q=document.frm09s01.LookupValueSearchOptions.options.length;q>=0;q--) document.frm09s01.LookupValueSearchOptions.options[q]=null;	  
		myEle = document.createElement("option") ;  
		var y = control.value;
		if ((y == "272") || (y == "273")){
			document.frm09s01.StringSearchTextboxTwo.disabled = false;
			oStg12.style.visibility="visible";
		} else {
			document.frm09s01.StringSearchTextboxTwo.disabled = true;
			oStg12.style.visibility="hidden";
		}
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
		
		document.frm09s01.StringSearchTextboxOne.value = "";
		
		var j = 0;
		var len = document.frm09s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm09s01.SearchType[i].checked) j = i;
		}
				
		if (j==1) selectChange4(frm09s01.LookupValueSearchOperator, frm09s01.LookupValueSearchOptions,Grp4Data);
		Togo();	  
	}

	function selectChange4(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		for (var q=document.frm09s01.LookupValueSearchOptions.options.length;q>=0;q--) document.frm09s01.LookupValueSearchOptions.options[q]=null;	  	  
		myEle = document.createElement("option") ;
		var y = document.frm09s01.LookupValueSearchOperand.value;
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
		if (document.frm09s01.StringSearchOperand.selectedIndex <= 0) {
			var j = 0;
		} else {
			var j = document.frm09s01.StringSearchOperand[document.frm09s01.StringSearchOperand.selectedIndex].value ;
		}
		if (document.frm09s01.StringSearchOperator.selectedIndex <= 0) {
			var l = 0;
		} else {
			var l = document.frm09s01.StringSearchOperator[document.frm09s01.StringSearchOperator.selectedIndex].value ;
		}
		document.frm09s01.MM_curOprd.value = j ;
		document.frm09s01.MM_curOptr.value = l ;
		document.frm09s01.MM_flag.value = true ;
	}
	
	</script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript" src="../js/m009Srh01.js"></script>
	<script language="JavaScript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	   

	function initscr() {
		oStg12.style.visibility="hidden";
		oOptrd21.style.visibility="hidden";
		oOptr21.style.visibility="hidden";
		oStg22.style.visibility="hidden";
		oStg31.style.visibility="hidden";
		oOptrd41.style.visibility="hidden";
		oOptr41.style.visibility="hidden";	  
		initializeList(document.frm09s01.MultiSelectOperand,document.frm09s01.MultiSelectOptions);
		var j = 0;
		var len = document.frm09s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm09s01.SearchType[i].checked) j = i;
		}
		switch (j) {
			case 0:
				selectChange(frm09s01.StringSearchOperand, frm09s01.StringSearchOperator,detailData);
			break;
			case 1:
				selectChange(frm09s01.LookupValueSearchOperand, frm09s01.LookupValueSearchOperator,detailData );
			break;
		}	   		
		
	}
	
	function addOption(txt, val){
		var oOption=document.createElement("OPTION");
		oOption.text = txt;
		oOption.value = val;
		document.frm09s01.MultiSelectOptions.add(oOption);
	}
	
	function initializeList(oParent, oChild){
	  	while (oChild.length > 0){
   			oChild.remove(0);
  		}
		switch (oParent.selectedIndex) {
			case 0:
				for (var i=0; i< DisabilityArraySize; i++) {
					addOption(DisabilityArray[i][1],DisabilityArray[i][0]);
				}
			break;
			default :
			break;
		}
	}	

	function CnstrFltr(output) {	
		var stgPgQuery = "";
		var stgFilter = "" ;
		var len = document.frm09s01.SearchType.length;
		var Idparam = 0;
		var j,k;    
		for (var i=0;i <len; i++){
			if (document.frm09s01.SearchType[i].checked) Idparam = i;
		}
  
		switch ( Idparam ) {
			//text
			case 0: 
				if (document.frm09s01.StringSearchOperand.length >= 1) {
					var chvOprd = document.frm09s01.StringSearchOperand[document.frm09s01.StringSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select String Search Operand.");
					break;
				}
        		var chrNot  = "";
				if (document.frm09s01.StringSearchOperator.length >= 1) {
					var chvOptr = document.frm09s01.StringSearchOperator[document.frm09s01.StringSearchOperator.selectedIndex].value ;
				} else {
					alert("Select String Search Operator.");
					break;
				}
				var chvStg1 = document.frm09s01.StringSearchTextboxOne.value;
				var chvStg2 = document.frm09s01.StringSearchTextboxTwo.value;
				if (((chvOprd=="272") || (chvOprd=="273")) && (chvOptr!="0")) {
				 	if (chvStg1 == "") {
						alert("Enter Start Date.");
						document.frm09s01.StringSearchTextboxOne.focus();
						return ;
					}
					if (chvStg2 == "") {
						alert("Enter End Date.");
						document.frm09s01.StringSearchTextboxTwo.focus();
					}
					if (!CheckDateBetween(Trim(chvStg1)+" and "+Trim(chvStg2))) {
						alert("End Date less than start date.");
						return ;
					}
				}
				if ((chvOprd=="274") || (chvOprd=="285")){
					if (!IsID(chvStg1)) {
						alert("Invalid number.");
						document.frm09s01.StringSearchTextboxOne.focus();
						return ;
					}
				}
				if (chvOptr == "0") {
					alert("Please select operator before Proceed");
				} else {
					stgFilter = ACfltr_09(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
				}
			break;
			//combo
			case 1: 
				if (document.frm09s01.LookupValueSearchOperand.length >= 1) {	  
					var chvOprd = document.frm09s01.LookupValueSearchOperand[document.frm09s01.LookupValueSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select Lookup Value Search Operand.");
					break;			
				}
				var chrNot  = "";
				if (document.frm09s01.LookupValueSearchOperator.length >= 1) {	  			 
					var chvOptr = document.frm09s01.LookupValueSearchOperator[document.frm09s01.LookupValueSearchOperator.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Operator.");
					break;			
				}
				if (document.frm09s01.LookupValueSearchOptions.length >= 1) {	  			 			
					var chvStg1 = document.frm09s01.LookupValueSearchOptions[document.frm09s01.LookupValueSearchOptions.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Option.");
					break;
				}
				var chvStg2 = "";
				if (chvOptr == "0") {
					alert("Please select operator before Proceed");
				} else {
					stgFilter = ACfltr_09(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
				}
			break;
			//class
			case 2: 
				var chvStg1 = document.frm09s01.ClassSearchID.value ;
				if (chvStg1 == "") {
					alert("Select Class.");
					break;
				}
				stgFilter = ACfltr_09("284","","3",chvStg1,"");
			break;
			//multiple selection
			case 3: 
				j = document.frm09s01.MultiSelectOperand[document.frm09s01.MultiSelectOperand.selectedIndex].value ;
				l = "";  
				var optList = document.frm09s01.MultiSelectOptions;
				var m = optList.length;
				if (optList.multiple) {
					for(var ii = 0; ii < m; ii++) {
						if (document.frm09s01.MultiSelectOptions[ii].selected) {
							if (l.length > 0 ) l += "," ;
							l += document.frm09s01.MultiSelectOptions[ii].value ;
						} 				  
					} 			      
				} else {
					l = document.frm09s01.MultiSelectOptions[document.frm09s01.MultiSelectOptions.selectedIndex].value
				} 
				if (l=="") {
					alert("Select at least one Multi-Value Search Option.");
					break;			
				}
				// Construct filters for multi-items select
				switch (j) {
					// Disability
					case "275" :
						stgFilter += " insDsbty1_id in (" + l + ") " ; 
					break;
				}
			break;
    	  	default: 
			break;
		}

		var inspSrtBy = document.frm09s01.SortByColumn.value;
		var inspSrtOrd = document.frm09s01.OrderBy.value;
		if (output==1) {
			document.frm09s01.action = "m009q01.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;		
		} else {
			document.frm09s01.action = "m009q01excel.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter;
			document.frm09s01.target = "_blank" ; 		
			document.frm09s01.submit() ; 								
		}
	}

	function SelOpt() {	
		var len = document.frm09s01.SearchType.length;
		var Idparam = 1;
		
		for (var i=0;i <len; i++){
			if (document.frm09s01.SearchType[i].checked) Idparam = i;
		}
		switch ( Idparam ) {
			// text 
			case 0: 
				oOptrd11.style.visibility="visible";
				oOptr11.style.visibility="visible";
				oStg11.style.visibility="visible";
				oStg12.style.visibility="hidden";
				oOptrd21.style.visibility="hidden";
				oOptr21.style.visibility="hidden";
				oStg22.style.visibility="hidden";
				oStg31.style.visibility="hidden";
				oOptrd41.style.visibility="hidden";
				oOptr41.style.visibility="hidden";
				selectChange(frm09s01.StringSearchOperand, frm09s01.StringSearchOperator,detailData);
			break;
			//Combo 
			case 1: 
				oOptrd11.style.visibility="hidden";
				oOptr11.style.visibility="hidden";
				oStg11.style.visibility="hidden";
				oStg12.style.visibility="hidden";
				oOptrd21.style.visibility="visible";
				oOptr21.style.visibility="visible";
				oStg22.style.visibility="visible";
				oStg31.style.visibility="hidden";
				oOptrd41.style.visibility="hidden";
				oOptr41.style.visibility="hidden";
				selectChange(frm09s01.LookupValueSearchOperand, frm09s01.LookupValueSearchOperator,detailData );			   
			break;
			//Picklist
			case 2: 
				oOptrd11.style.visibility="hidden";
				oOptr11.style.visibility="hidden";
				oStg11.style.visibility="hidden";
				oStg12.style.visibility="hidden";
				oOptrd21.style.visibility="hidden";
				oOptr21.style.visibility="hidden";
				oStg22.style.visibility="hidden";
				oStg31.style.visibility="visible";
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
				oStg31.style.visibility="hidden";
				oOptrd41.style.visibility="visible";
				oOptr41.style.visibility="visible";			   
			break; 
			default: 
			break;
		}       
	}

	function Toggle() {	
		openWindow("m009p01FS.asp","");
	}	
	if (window.focus) self.focus();		
	</script>
</head>
<body onload="initscr()" >
<form name="frm09s01" method="post" action="">
<h3>Equipment Service - Quick Search</h3>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="1" tabindex="1" accesskey="F" checked onClick="SelOpt()" class="chkstyle">String Search</td>
		<td nowrap><DIV ID="oOptrd11" STYLE="visibility:visible" > 
			<select name="StringSearchOperand" onchange="selectChange(this, frm09s01.StringSearchOperator,detailData);" tabindex="2">
			<% 
			while (!rsOprd.EOF) { 
			%>
				<option value="<%=(rsOprd.Fields.Item("intRecID").Value)%>" <%=((rsOprd.Fields.Item("intRecID").Value ==272)?"SELECTED":"")%> ><%=(rsOprd.Fields.Item("chvObjName").Value)%></option>
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
			<select name="LookupValueSearchOperand" onchange="selectChange(this, frm09s01.LookupValueSearchOperator,detailData);" tabindex="7">
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
			<select name="LookupValueSearchOperator" onchange="selectChange4(this, frm09s01.LookupValueSearchOptions,Grp4Data );" tabindex="8"></select>
		</DIV></td>
		<td nowrap><DIV ID="oStg22" STYLE="visibility:visible" > 
			<select name="LookupValueSearchOptions" tabindex="9"></select>
        </DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="3" onClick="SelOpt()" tabindex="10" class="chkstyle">Class Search</td>
		<td nowrap><DIV ID="oStg31" STYLE="visibility:visible">
			<input type="text" name="ClassSearchText" size="30" READONLY tabindex="12">
			<input type="button" name="ClassSearchPickList" value="List" onClick="Toggle();" tabindex="13" class="btnstyle">
		</DIV></td>		
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="4" onClick="SelOpt()" tabindex="14" class="chkstyle">Multi-Value Search</td>
		<td nowrap valign="top"><DIV ID="oOptrd41" STYLE="visibility:visible"> 
			<select name="MultiSelectOperand" OnChange="initializeList(document.frm09s01.MultiSelectOperand,document.frm09s01.MultiSelectOptions)" tabindex="15">
				<option value="275">Disability
			</select>
		</DIV></td>		
		<td nowrap valign="top"><DIV ID="oOptr41" STYLE="visibility:hidden"> 
			contains: <select name="MultiSelectOptions" size="8" multiple align="top" tabindex="16" style="width: 160px"></select>
		</DIV></td>		
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>
			Sort by:
			<select name="SortByColumn" tabindex="17">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("intRecID").Value == 268)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvObjName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>
			</select>
			Order 
        	<select name="OrderBy" tabindex="18">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>
	        <input type="submit" value="Search" onClick="CnstrFltr(1);" tabindex="19" class="btnstyle">
			<input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="20" class="btnstyle">
			<input type="reset" value="Clear All" tabindex="21" onClick="window.location.reload();" class="btnstyle">
		</td>		
    </tr>
</table>
<input type="hidden" name="MM_flag" value="false">
<input type="hidden" name="MM_curOprd">
<input type="hidden" name="MM_curOptr">
<input type="hidden" name="ClassSearchID">		
</form>
</body>
</html>
<%
rsOprd.Close();
rsOprd2.Close();
rsOprd3.Close();
//rsOprd4.Close();
rsCol.Close();
%>