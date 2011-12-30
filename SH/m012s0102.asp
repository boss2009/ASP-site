<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
// retrieve sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup2(742,0,'',0,'',0)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

// retrieve text search operands
var rsOprd = Server.CreateObject("ADODB.Recordset");
rsOprd.ActiveConnection = MM_cnnASP02_STRING;
rsOprd.Source = "{call dbo.cp_ASP_Lkup2(743,0,'',0,'',0)}";
rsOprd.CursorType = 0;
rsOprd.CursorLocation = 2;
rsOprd.LockType = 3;
rsOprd.Open();

// retrieve lookup search operands
var rsOprd2 = Server.CreateObject("ADODB.Recordset");
rsOprd2.ActiveConnection = MM_cnnASP02_STRING;
rsOprd2.Source = "{call dbo.cp_ASP_Lkup2(745,0,'',0,'',0)}";
rsOprd2.CursorType = 0;
rsOprd2.CursorLocation = 2;
rsOprd2.LockType = 3;
rsOprd2.Open();

// retrieve multi search operands
var rsOprd3 = Server.CreateObject("ADODB.Recordset");
rsOprd3.ActiveConnection = MM_cnnASP02_STRING;
rsOprd3.Source = "{call dbo.cp_ASP_Lkup2(744,0,'',0,'',0)}";
rsOprd3.CursorType = 0;
rsOprd3.CursorLocation = 2;
rsOprd3.LockType = 3;
rsOprd3.Open();
%>
<html>
<head>
	<title>Institution - Advanced Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="Javascript">
	var detailData = new Array();
	<%
	var intOldOptr = 0;
	var objOptrDesc,objOptrId,objRecID;
	
	var rsOptr = Server.CreateObject("ADODB.Recordset");
	rsOptr.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,27)}";	
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
	// retrieve the Institution Type lookup
	var rsInstitutionType = Server.CreateObject("ADODB.Recordset");
	rsInstitutionType.ActiveConnection = MM_cnnASP02_STRING;
	rsInstitutionType.Source = "{call dbo.cp_school_type(0,'',1,0,'Q',0)}";
	rsInstitutionType.CursorType = 0;
	rsInstitutionType.CursorLocation = 2;
	rsInstitutionType.LockType = 3;
	rsInstitutionType.Open();
	if (!rsInstitutionType.EOF){ 	
		Response.Write("Grp4Data[138] = new Array();")	
		while (!rsInstitutionType.EOF) { 
	%>
			Grp4Data[138][<%=rsInstitutionType("insSchool_type_id")%>] = "<%= rsInstitutionType("chvSchool_Type") %>"
	<%
			rsInstitutionType.MoveNext 
		}
	} else {
		Response.Write("Institution lookup does not exist.")
	}
	
	rsInstitutionType.Close();
	%>
	
	Grp4Data[139] = new Array();
	Grp4Data[139][1] = "Main Campus"
	Grp4Data[139][0] = "Satellite Campus"		  	

	var RegionArray   = new Array(2);
	<%
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
	var ServiceCodeArray   = new Array(2);
	<%
	var intServiceCodeCnt = 0 ;
	// retrieve the Service Code lookup
	var rsServiceCode = Server.CreateObject("ADODB.Recordset");
	rsServiceCode.ActiveConnection = MM_cnnASP02_STRING;
	rsServiceCode.Source = "{call dbo.cp_service_type(0,0,1,2)}";
	rsServiceCode.CursorType = 0;
	rsServiceCode.CursorLocation = 2;
	rsServiceCode.LockType = 3;
	rsServiceCode.Open();
	while (!rsServiceCode.EOF){
		intServiceCodeCnt++;
		rsServiceCode.MoveNext;
	}
	rsServiceCode.MoveFirst;
	if (!rsServiceCode.EOF){ 	
	%>
	var ServiceCodeArraySize = <%=intServiceCodeCnt%>;
	for (var i=0; i < <%=intServiceCodeCnt%>; i++){
		ServiceCodeArray[i] = new Array(<%=intServiceCodeCnt%>);
	}
	<%	   
		intServiceCodeCnt = 0;
		while (!rsServiceCode.EOF) { 
	%>
			ServiceCodeArray[<%=intServiceCodeCnt%>][0] = "<%=rsServiceCode("insService_type_id")%>"
			ServiceCodeArray[<%=intServiceCodeCnt%>][1] = "<%=rsServiceCode("chvname")%>"
	<%
			intServiceCodeCnt += 1
			rsServiceCode.MoveNext 
		}
	} else {
		Response.Write("Service Code lookup does not exist.")
	}
	
	rsServiceCode.Close();
	%>
		
	function selectChange(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		for (var q=document.frm12s01.LookupValueSearchOptions.options.length;q>=0;q--) document.frm12s01.LookupValueSearchOptions.options[q]=null;	  
		myEle = document.createElement("option") ;  
		var y = control.value;
		if (y == "141"){
			document.frm12s01.StringSearchTextboxTwo.disabled = false;
			oImg11.style.visibility="visible";
		} else {
			document.frm12s01.StringSearchTextboxTwo.disabled = true;
			oImg11.style.visibility="hidden";
		}

		document.frm12s01.StringSearchTextboxOne.value = "";
		
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
		var j = 0;
		var len = document.frm12s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm12s01.SearchType[i].checked) j = i;
		}
				
		if (j==1) {
			selectChange4(frm12s01.LookupValueSearchOperator, frm12s01.LookupValueSearchOptions,Grp4Data );
		}	
		Togo();	 
	}

	function selectChange4(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		for (var q=document.frm12s01.LookupValueSearchOptions.options.length;q>=0;q--) document.frm12s01.LookupValueSearchOptions.options[q]=null;	  	  
		myEle = document.createElement("option") ;
		var y = document.frm12s01.LookupValueSearchOperand.value;
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
		if (document.frm12s01.StringSearchOperand.selectedIndex <= 0) {
			var j = 0;
		} else {
			var j = document.frm12s01.StringSearchOperand[document.frm12s01.StringSearchOperand.selectedIndex].value ;
		}
		if (document.frm12s01.StringSearchOperator.selectedIndex <= 0) {
			var l = 0;
		} else {
			var l = document.frm12s01.StringSearchOperator[document.frm12s01.StringSearchOperator.selectedIndex].value ;
		}
		document.frm12s01.MM_curOprd.value = j ;
		document.frm12s01.MM_curOptr.value = l ;
		document.frm12s01.MM_flag.value = true ;
	}
	
	</script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript" src="../js/m012Srh01.js"></script>
	<script language="JavaScript">
	if (window.focus) self.focus();
	function initscr() {
		oImg11.style.visibility="hidden";
		oOptrd21.style.visibility="hidden";
		oOptr21.style.visibility="hidden";
		oStg21.style.visibility="hidden";
		oOptrd31.style.visibility="hidden";
		oOptr31.style.visibility="hidden";	  
		initializeList(document.frm12s01.MultiSelectOperand,document.frm12s01.MultiSelectOptions);
		var j = 0;
		var len = document.frm12s01.SearchType.length;		
		for (var i=0;i <len; i++) {
			if (document.frm12s01.SearchType[i].checked) j = i;
		}
		switch (j) {
			case 0:
				selectChange(frm12s01.StringSearchOperand, frm12s01.StringSearchOperator,detailData);
			break;
			case 1:
				selectChange(frm12s01.LookupValueSearchOperand, frm12s01.LookupValueSearchOperator,detailData );
			break;
		}	   				
	}
	
	function addOption(txt, val){
		var oOption=document.createElement("OPTION");
		oOption.text = txt;
		oOption.value = val;
		document.frm12s01.MultiSelectOptions.add(oOption);
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
				for (var i=0; i< ServiceCodeArraySize; i++) {		
					addOption(ServiceCodeArray[i][1],ServiceCodeArray[i][0]);
				}			
			break;
		}
	}	

	function CnstrFltr(output) {		
		var inspSrtBy = document.frm12s01.SortByColumn.value;
		var inspSrtOrd = document.frm12s01.OrderBy.value;
		var stgFilter = document.frm12s01.QueryString.value;
		if (output==1) {
			document.frm12s01.action = "m012q01.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;		
			document.frm12s01.submit() ; 		
		} else {
			var ExcelSearch = window.open("m012q01excel.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter);
		}
	}

	function Savtxt() {	
		var stgPgQuery = "";
		var stgFilter = "" ;
		var blnFlg = false ;		
		var len = document.frm12s01.SearchType.length;
		var Idparam = 0;                 		
		var stgTemp,j,k;
		for (var x=0; x<=3; x++) {
			if (document.frm12s01.SearchType[x].checked) Idparam = x;		
		}
			
		stgTemp = document.frm12s01.QueryString.value; 
		switch (Idparam) {
			case 0: 
				if (document.frm12s01.StringSearchOperand.length >= 1) {
					var chvOprd = document.frm12s01.StringSearchOperand[document.frm12s01.StringSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select String Search Operand.");
					break;
				}
				var chrNot  = "";
				if (document.frm12s01.StringSearchOperator.length >= 1) {
					var chvOptr = document.frm12s01.StringSearchOperator[document.frm12s01.StringSearchOperator.selectedIndex].value ;
				} else {
					alert("Select String Search Operator.");
					break;
				}
				var chvStg1 = document.frm12s01.StringSearchTextboxOne.value;
				var chvStg2 = document.frm12s01.StringSearchTextboxTwo.value;
	             var chvAO1  = document.frm12s01.AndOr.value ;				
				if ((chvOprd=="141") && (chvOptr!="0")) {
					if (chvStg1 == "") {
						alert("Enter Start Date.");
						document.frm12s01.StringSearchTextboxOne.focus();
						return ;
					}
					if (chvStg2 == "") {
						alert("Enter End Date.");
						document.frm12s01.StringSearchTextboxTwo.focus();
					}
					if (!CheckDateBetween(Trim(chvStg1)+" and "+Trim(chvStg2))) return ;
				}
				if (chvOprd == "134") {
					if (!IsID(chvStg1)) {
						alert("Invalid number.");
						return;
					}					
				}
				if (chvOptr == "0") {
					alert("Please select operator before Proceed");
				} else {
					stgFilter = ACfltr_12(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
					if (stgTemp.length > 0 ) stgTemp += " (" ;
					stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 ;
				}
			break;
			case 1: 
		  		if (document.frm12s01.LookupValueSearchOperand.length >= 1) {	  
					var chvOprd = document.frm12s01.LookupValueSearchOperand[document.frm12s01.LookupValueSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select Lookup Value Search Operand.");
					break;			
				}
				 var chrNot  = "";
				if (document.frm12s01.LookupValueSearchOperator.length >= 1) {	  			 
					var chvOptr = document.frm12s01.LookupValueSearchOperator[document.frm12s01.LookupValueSearchOperator.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Operator.");
					break;			
				}
				if (document.frm12s01.LookupValueSearchOptions.length >= 1) {	  			 			
					var chvStg1 = document.frm12s01.LookupValueSearchOptions[document.frm12s01.LookupValueSearchOptions.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Option.");
					break;
				}
				var chvStg2 = "";
	            var chvAO1  = document.frm12s01.AndOr.value ;				
				if (chvOptr == "0") {
					alert("Please select operator before Proceed");
				} else {
					stgFilter = ACfltr_12(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
					if (stgTemp.length > 0 ) stgTemp += " (" ;
					stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 ;					
				}
			break;
			case 2: 
				j = document.frm12s01.MultiSelectOperand[document.frm12s01.MultiSelectOperand.selectedIndex].value ;
				l = "";  
				var optList = document.frm12s01.MultiSelectOptions;
				var m = optList.length;
				if (optList.multiple) {
					for(var ii = 0; ii < m; ii++) {
						if (document.frm12s01.MultiSelectOptions[ii].selected) {
							if (l.length > 0 ) l += "," ;
							l += document.frm12s01.MultiSelectOptions[ii].value ;	
						} 				  
					} 					  
				} else {
					l = document.frm12s01.MultiSelectOptions[document.frm12s01.MultiSelectOptions.selectedIndex].value
				} 
				if (l=="") {
					alert("Select at least one Multi Value Search Option.");
					break;			
				}
				switch (j) {
					// Service Code
					case "140" :
						stgFilter += " insSrv_Code_id in (" + l + ") " ; 
						blnFlg = true;						
					break;
					// Region
					case "136" : 
						stgFilter += " insRegion_num in (" + l + ") " ; 
						blnFlg = true;												
					break;			
				}
			break;
			default: 
			break;
		}						
		document.frm12s01.StringSearchTextboxOne.value="";
		document.frm12s01.StringSearchTextboxTwo.value="";		
		if (blnFlg) {
			var chvAO1  = document.frm12s01.AndOr.value ;
			stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 + " ";
		}
		document.frm12s01.QueryString.value = stgTemp; 
	}
	
	function SelOpt() {	
		var len = document.frm12s01.SearchType.length;
		var Idparam = 1;                 // init.
	
		for (var i=0;i <len; i++){
			if (document.frm12s01.SearchType[i].checked) Idparam = i;
		}
		switch (Idparam) {
			//text
			case 0: 
				oOptrd11.style.visibility="visible";
				oOptr11.style.visibility="visible";
				oStg11.style.visibility="visible";
				oImg11.style.visibility="hidden";
				oOptrd21.style.visibility="hidden";
				oOptr21.style.visibility="hidden";
				oStg21.style.visibility="hidden";
				oOptrd31.style.visibility="hidden";
				oOptr31.style.visibility="hidden";
				selectChange(frm12s01.StringSearchOperand, frm12s01.StringSearchOperator,detailData);
			break;
			//Combo 
			case 1: 
			   oOptrd11.style.visibility="hidden";
			   oOptr11.style.visibility="hidden";
			   oStg11.style.visibility="hidden";
			   oImg11.style.visibility="hidden";
			   oOptrd21.style.visibility="visible";
			   oOptr21.style.visibility="visible";
			   oStg21.style.visibility="visible";
			   oOptrd31.style.visibility="hidden";
			   oOptr31.style.visibility="hidden";
			   selectChange(frm12s01.LookupValueSearchOperand, frm12s01.LookupValueSearchOperator,detailData );			   
			break;
			case 2: 
				oOptrd11.style.visibility="hidden";
				oOptr11.style.visibility="hidden";
				oStg11.style.visibility="hidden";
				oImg11.style.visibility="hidden";
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
<body onload="initscr()" >
<form name="frm12s01" method="post">
<h3>Institution - Advanced Search</h3>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="1" tabindex="1" accesskey="F" checked onClick="SelOpt()" class="chkstyle">String Search</td>
		<td nowrap><DIV ID="oOptrd11" STYLE="visibility:visible" > 
			<select name="StringSearchOperand" onchange="selectChange(this, frm12s01.StringSearchOperator,detailData);" tabindex="2">
			<% 
			while (!rsOprd.EOF) { 
			%>
				<option value="<%=(rsOprd.Fields.Item("intRecID").Value)%>" <%=((rsOprd.Fields.Item("intRecID").Value ==135)?"SELECTED":"")%> ><%=(rsOprd.Fields.Item("chvObjName").Value)%></option>
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
		<td nowrap><DIV ID="oImg11" STYLE="visibility:hidden">
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
			<select name="LookupValueSearchOperand" onchange="selectChange(this, frm12s01.LookupValueSearchOperator,detailData);" tabindex="7">
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
			<select name="LookupValueSearchOperator" onchange="selectChange4(this, frm12s01.LookupValueSearchOptions,Grp4Data );" tabindex="8"></select>
		</DIV></td>
		<td nowrap><DIV ID="oStg21" STYLE="visibility:visible"> 
			<select name="LookupValueSearchOptions" tabindex="9"></select>
        </DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap valign="top" width="160"><input type="radio" name="SearchType" value="3" onClick="SelOpt()" tabindex="10" class="chkstyle">Multi-Value Search</td>
		<td nowrap valign="top"><DIV ID="oOptrd31" STYLE="visibility:visible"> 
			<select name="MultiSelectOperand" OnChange="initializeList(document.frm12s01.MultiSelectOperand,document.frm12s01.MultiSelectOptions)" tabindex="11">
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
			contains: <select name="MultiSelectOptions" size="8" width="150" multiple align="top" tabindex="12"></select>
		</DIV></td>		
	</tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td valign="top"><select name="AndOr" tabindex="13">
			<option value=" ">None</option>
			<option value="And">And</option>
			<option value="Or">Or</option>
		</select></td>	
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td valign="top" nowrap><input type="button" name="AddToQueryString" value="Add" tabindex="14" onClick="Savtxt()" class="btnstyle"></td>
		<td valign="top" nowrap><textarea name="QueryString" cols="80" rows="6" tabindex="15" accesskey="L"></textarea></td>
    </tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>
			Sort by:
			<select name="SortByColumn" tabindex="16">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("intRecID").Value)%>" <%=((rsCol.Fields.Item("intRecID").Value == 129)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvObjName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>
			</select>
			Order 
        	<select name="OrderBy" tabindex="17">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
	</tr>
	<tr>
		<td>
	        <input type="button" value="Search" onClick="CnstrFltr(1);" tabindex="18" class="btnstyle">
			<input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="19" class="btnstyle">
			<input type="reset" value="Clear All" tabindex="20" onClick="window.location.reload();" class="btnstyle">
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