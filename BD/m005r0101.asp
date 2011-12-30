<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
// retrieve the sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup2(739,0,'',0,'Q',0)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

// retrieve the string search operands - text
var rsOprd = Server.CreateObject("ADODB.Recordset");
rsOprd.ActiveConnection = MM_cnnASP02_STRING;
rsOprd.Source = "{call dbo.cp_ASP_Lkup2(740,0,'',0,'Q',0)}";
rsOprd.CursorType = 0;
rsOprd.CursorLocation = 2;
rsOprd.LockType = 3;
rsOprd.Open();

// retrieve the lookup value search operands - Combo
var rsOprd2 = Server.CreateObject("ADODB.Recordset");
rsOprd2.ActiveConnection = MM_cnnASP02_STRING;
rsOprd2.Source = "{call dbo.cp_ASP_Lkup2(741,0,'',0,'Q',0)}";
rsOprd2.CursorType = 0;
rsOprd2.CursorLocation = 2;
rsOprd2.LockType = 3;
rsOprd2.Open();
%>
<html>
<head>
	<title>Bundles - Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="Javascript">
	//-------
	// Group 2 data
	//-------
	var detailData = new Array();
	<%
	var intOldOptr = 0;
	var intIdxOprd = 0;
	var objOptrDesc,objOptrId,objRecID;
	
	var rsOptr = Server.CreateObject("ADODB.Recordset");
	rsOptr.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,20)}";
	
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
	}
	else {
	   Response.Write("SysOptr lookup does not exist.")
	}
	
	rsOptr.Close();
	%>
	//-------
	// Build the List Box Array which house Group 4 lookups  + APR.03.2002
	//-------
	var Grp4Data   = new Array();
	<%
	// Status
	Response.Write("Grp4Data[123] = new Array();")
	%>	
	Grp4Data[123][1] = "Active"
	Grp4Data[123][2] = "Inactive"		  

	<%
	// Type
	Response.Write("Grp4Data[124] = new Array();")
	%>	
	Grp4Data[124][1] = "Desktop"
	Grp4Data[124][2] = "Laptop"		  
	Grp4Data[124][0] = "Other"

	<%
	// Purpose
	intIdxOprd4 = 125;
	Response.Write("Grp4Data[125] = new Array();")
	%>	
	Grp4Data[125][0] = "CSG"
	Grp4Data[125][1] = "Loan"		  
	Grp4Data[125][2] = "CSG And Loan"

	function selectChange(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		for (var q=document.frm05s01.LookupValueSearchOptions.options.length;q>=0;q--) document.frm05s01.LookupValueSearchOptions.options[q]=null;	  	  
		myEle = document.createElement("option") ;
		var y = control.value;
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

		document.frm05s01.StringSearchTextOne.value = "";
		var j = 0;
		var len = document.frm05s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm05s01.SearchType[i].checked) j = i;
		}				
		if (j==1) {
			selectChange4(frm05s01.LookupValueSearchOperator, frm05s01.LookupValueSearchOptions,Grp4Data );
		}
		Togo();	  
	}

	function selectChange4(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		myEle = document.createElement("option") ;
		var y = document.frm05s01.LookupValueSearchOperand.value;
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
		if (document.frm05s01.StringSearchOperand.selectedIndex <= 0) {
			var j = 0;
		} else {
			var j = document.frm05s01.StringSearchOperand[document.frm05s01.StringSearchOperand.selectedIndex].value ;
		}
		if (document.frm05s01.StringSearchOperator.selectedIndex <= 0) {
			var l = 0;
		} else {
			var l = document.frm05s01.StringSearchOperator[document.frm05s01.StringSearchOperator.selectedIndex].value ;
		}
		// hold the current operand for next refresh
		document.frm05s01.MM_curOprd.value = j ;
		// hold the current operator for next refresh
		document.frm05s01.MM_curOptr.value = l ;
		// Set the flag
		document.frm05s01.MM_flag.value = true ;
	}
	
	</script>
	<script language="JavaScript" src="../js/m005Srh01.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	   
	function initscr() {
		oOptrd21.style.visibility="hidden";
		oOptrd31.style.visibility="hidden";
		var j = 0;
		var len = document.frm05s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm05s01.SearchType[i].checked) j = i;
		}
		switch (j) {
			case 0:
				selectChange(frm05s01.StringSearchOperand, frm05s01.StringSearchOperator,detailData);
			break;
			case 1:
				selectChange(frm05s01.LookupValueSearchOperand, frm05s01.LookupValueSearchOperator,detailData );
			break;
		}	   				
	}

	function CnstrFltr(output) {	
		// Radio box
		var len = document.frm05s01.SearchType.length;
		var radioGrp = document.frm05s01.SearchType
		var Idparam = -1;                 // init.
		var stgTemp,j,k; 
		
		for (var i=0;i <len; i++){
			if (document.frm05s01.SearchType[i].checked) Idparam = i;
		}
	
		switch (Idparam) {
			// ---------------------------------
			//text Group
			// ---------------------------------
			case 0: 
		  		if (document.frm05s01.StringSearchOperand.length >= 1) {			
					var chvOprd = document.frm05s01.StringSearchOperand[document.frm05s01.StringSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select String Search Operand.");
					return ;
					break;
				}					
				var chrNot  = "";
				if (document.frm05s01.StringSearchOperator.length >= 1) {					
					var chvOptr = document.frm05s01.StringSearchOperator[document.frm05s01.StringSearchOperator.selectedIndex].value ;
				} else {
					alert("Select String Search Operator.");
					return ;
					break;
				}					
				var chvStg1 = document.frm05s01.StringSearchTextOne.value;
				var chvStg2 = "";
			break; 
			// ---------------------------------
			//combo Group 
			// ---------------------------------
			case 1: 
				if (document.frm05s01.LookupValueSearchOperand.length >= 1) {	  
					 var chvOprd = document.frm05s01.LookupValueSearchOperand[document.frm05s01.LookupValueSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select LookUp Search Operand.");
					return ;
					break;			
				}
				 var chrNot  = "";
				if (document.frm05s01.LookupValueSearchOperator.length >= 1) {	  			 
					 var chvOptr = document.frm05s01.LookupValueSearchOperator[document.frm05s01.LookupValueSearchOperator.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Operator.");
					return ;
					break;			
				}
				if (document.frm05s01.LookupValueSearchOptions.length >= 1) {	  			 			
					 var chvStg1 = document.frm05s01.LookupValueSearchOptions[document.frm05s01.LookupValueSearchOptions.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Option.");
					return ;
					break;
				}
				var chvStg2 = "";
				var chvOptr = document.frm05s01.LookupValueSearchOptions.value;
			break;
			// ---------------------------------
			//picklist Group 
			// ---------------------------------
			case 2: 
				if (document.frm05s01.ClassSearchOperand.length >= 1) {	  
					 var chvOprd = document.frm05s01.ClassSearchOperand[document.frm05s01.ClassSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select Class Search Operand.");
					return ;
					break;			
				}	  
				var chrNot  = "";
				var chvOptr = "3";
				var chvStg1 = document.frm05s01.ClassID.value ;
				var chvStg2 = "";
				if (chvStg1 == "") {
					alert("Select Class.");
					return ;
					break;
				}				
			break;
			default:
				alert("program Error - radio buttion 'Sel' is not picked ...");
				return ;
			break; 
		}
		if (chvOptr < "0") {
			document.write("<B><B><B>");
			document.write("Please Select Operator.");
			document.write("<B><B><B>");
		} else {
			var stgFilter = ACfltr_05(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
			var inspSrtBy = document.frm05s01.SortByColumn.value;
			var inspSrtOrd = document.frm05s01.OrderBy.value;
			if (output==1) {
				document.frm05s01.action = "m005q02.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;
				document.frm05s01.submit() ; 					
			} else {
				var ExcelSearch = window.open("m005q02excel.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter);
			}
		}
	}

	function SelOpt() {	
		var len = document.frm05s01.SearchType.length;
		var Idparam = 1;                 // init.
	
		for (var i=0;i <len; i++){
			if (document.frm05s01.SearchType[i].checked) Idparam = i;
		}
		switch (Idparam) {
			// text 
			case 0: 
				oOptrd11.style.visibility="visible";
				oOptrd21.style.visibility="hidden";
				oOptrd31.style.visibility="hidden";
				selectChange(frm05s01.StringSearchOperand, frm05s01.StringSearchOperator,detailData);
			break;
			//Combo 
			case 1:
				oOptrd11.style.visibility="hidden";
				oOptrd21.style.visibility="visible";
				oOptrd31.style.visibility="hidden";
				selectChange(frm05s01.LookupValueSearchOperand, frm05s01.LookupValueSearchOperator,detailData );
			break;
			//Picklist
			case 2: 
				oOptrd11.style.visibility="hidden";
				oOptrd21.style.visibility="hidden";
				oOptrd31.style.visibility="visible";
			break; 
			default: 
			break;
		}       
	}

	function Toggle() {	
		var idx = document.frm05s01.ClassSearchOperand[document.frm05s01.ClassSearchOperand.selectedIndex].value
		switch ( idx ) {
			// class no 
			case "127":
				openWindow("m005p0101.asp","");
			break;
			default: 
				document.frm05s01.ClassText.value = ""; 
			break;
		}
	}
	</script>
</head>
<body onload="initscr()" >
<form name="frm05s01" method="post">
<h3>Bundles - Report</h3>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap><input type="radio" name="SearchType" value="1" tabindex="1" accesskey="F" checked onClick="SelOpt()" class="chkstyle">String Search</td>
		<td nowrap><DIV ID="oOptrd11" STYLE="visibility:visible">
			<select name="StringSearchOperand" onchange="selectChange(this, frm05s01.StringSearchOperator,detailData);" tabindex="2">
			<% 
			while (!rsOprd.EOF) { 
			%>
				<option value="<%=(rsOprd.Fields.Item("intRecID").Value)%>" <%=((rsOprd.Fields.Item("intRecID").Value == Request.Form("MM_curOprd"))?"SELECTED":"")%> ><%=(rsOprd.Fields.Item("chvObjName").Value)%></option>
			<% 
				rsOprd.MoveNext(); 
			}
			%>
			</select>
			<select name="StringSearchOperator" onchange="Togo();" tabindex="3"></select>
			<input type="text" name="StringSearchTextOne" tabindex="4"></DIV>
		</td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">	
    <tr> 
		<td nowrap><input type="radio" name="SearchType" value="2" onClick="SelOpt()" tabindex="5" class="chkstyle">Lookup Value Search</td>
		<td nowrap><DIV ID="oOptrd21" STYLE="visibility:visible">
			<select name="LookupValueSearchOperand" onchange="selectChange(this, frm05s01.LookupValueSearchOperator,detailData);" tabindex="6">
			<% 
			while (!rsOprd2.EOF) { 
			%>
				<option value="<%=(rsOprd2.Fields.Item("intRecID").Value)%>" <%=((rsOprd2.Fields.Item("intRecID").Value == 54)?"SELECTED":"")%>><%=(rsOprd2.Fields.Item("chvObjName").Value)%></option>
			<% 
				rsOprd2.MoveNext();
			}
			%>
			</select>
			<select name="LookupValueSearchOperator" onchange="selectChange4(this, frm05s01.LookupValueSearchOptions,Grp4Data );" tabindex="7"></select>
			<select name="LookupValueSearchOptions" tabindex="8"></select>
        </DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap><input type="radio" name="SearchType" value="3" onClick="SelOpt()" tabindex="9" class="chkstyle">Class Search</td>
		<td nowrap><DIV ID="oOptrd31" STYLE="visibility:visible"> 
			<select name="ClassSearchOperand" tabindex="10">
				<option value="127">Contains Class
			</select>
			<input type="text" name="ClassName" READONLY tabindex="11">
			<input type="button" value="List" onClick="Toggle();" tabindex="12" class="btnstyle">
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Sort by:
			<select name="SortByColumn" tabindex="13">
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
        	<select name="OrderBy" tabindex="14">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>
	        <input type="button" value="Search" onClick="CnstrFltr(1);" tabindex="15" class="btnstyle">
			<input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="16" class="btnstyle">			
			<input type="reset" value="Clear All" tabindex="17" onClick="window.location.reload();" class="btnstyle">
		</td>		
    </tr>
</table>
<input type="hidden" name="ClassID">
<input type="hidden" name="MM_flag" value="false">
<input type="hidden" name="MM_curOprd">
<input type="hidden" name="MM_curOptr">
</form>
</body>
</html>
<%
rsOprd.Close();
rsOprd2.Close();
%>