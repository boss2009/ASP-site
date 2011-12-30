<!--------------------------------------------------------------------------
* File Name: m014s0102.asp
* Title: Purchase Requisition - Advanced Search
* Main SP: 
* Description: Purchase Requisition Advanced Search.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" --> 
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
// retrieve the sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup(729)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

// retrieve the string search operands - text
var rsOprd = Server.CreateObject("ADODB.Recordset");
rsOprd.ActiveConnection = MM_cnnASP02_STRING;
rsOprd.Source = "{call dbo.cp_ASP_Lkup(726)}";
rsOprd.CursorType = 0;
rsOprd.CursorLocation = 2;
rsOprd.LockType = 3;
rsOprd.Open();

// retrieve the lookup value search operands - Combo
var rsOprd2 = Server.CreateObject("ADODB.Recordset");
rsOprd2.ActiveConnection = MM_cnnASP02_STRING;
rsOprd2.Source = "{call dbo.cp_ASP_Lkup(727)}";
rsOprd2.CursorType = 0;
rsOprd2.CursorLocation = 2;
rsOprd2.LockType = 3;
rsOprd2.Open();

var rsOprd3 = Server.CreateObject("ADODB.Recordset");
rsOprd3.ActiveConnection = MM_cnnASP02_STRING;
rsOprd3.Source = "{call dbo.cp_ASP_Lkup(728)}";
rsOprd3.CursorType = 0;
rsOprd3.CursorLocation = 2;
rsOprd3.LockType = 3;
rsOprd3.Open();
%>
<html>
<head>
	<title>Purchase Requisition - Advanced Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="Javascript">
	var detailData = new Array();
	<%
	var intOldOptr = 0;
	var objOptrDesc,objOptrId,objRecID;
	
	var rsOptr = Server.CreateObject("ADODB.Recordset");
	rsOptr.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,29)}";	
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
	var intID,chrData;
	// retrieve the Vendor lookup
	var rsVendor = Server.CreateObject("ADODB.Recordset");
	rsVendor.ActiveConnection = MM_cnnASP02_STRING;
	rsVendor.Source = "{call dbo.cp_company2(0,'',0,0,0,0,0,1,0,'',0,'Q',0)}";
	rsVendor.CursorType = 0;
	rsVendor.CursorLocation = 2;
	rsVendor.LockType = 3;
	rsVendor.Open();
	if (!rsVendor.EOF){ 	
		Response.Write("Grp4Data[82] = new Array();")	
		while (!rsVendor.EOF) { 
			chrData = rsVendor("chvCompany_Name")
			intID   = rsVendor("intCompany_id")
	%>
			Grp4Data[82][<%=intID%>] = "<%= chrData %>"
	<%
			rsVendor.MoveNext 
		}
	} else {
		Response.Write("Vendor lookup does not exist.")
	}	
	rsVendor.Close();

	// retrieve the Request Type lookup
	var rsRequestType = Server.CreateObject("ADODB.Recordset");
	rsRequestType.ActiveConnection = MM_cnnASP02_STRING;
	rsRequestType.Source = "{call dbo.cp_ASP_Lkup(55)}";
	rsRequestType.CursorType = 0;
	rsRequestType.CursorLocation = 2;
	rsRequestType.LockType = 3;
	rsRequestType.Open();
	if (!rsRequestType.EOF){ 	
		Response.Write("Grp4Data[80] = new Array();")	
		while (!rsRequestType.EOF) { 
			chrData = rsRequestType("chvname")
			intID   = rsRequestType("insPur_type_id")
	%>
			Grp4Data[80][<%=intID%>] = "<%= chrData %>"
	<%
			rsRequestType.MoveNext 
		}
	} else {
		Response.Write("Request Type lookup does not exist.")
	}	
	rsRequestType.Close();

	// retrieve the Work Order lookup
	var rsWorkOrder = Server.CreateObject("ADODB.Recordset");
	rsWorkOrder.ActiveConnection = MM_cnnASP02_STRING;
	rsWorkOrder.Source = "{call dbo.cp_ASP_Lkup(59)}";
	rsWorkOrder.CursorType = 0;
	rsWorkOrder.CursorLocation = 2;
	rsWorkOrder.LockType = 3;
	rsWorkOrder.Open();
	if (!rsWorkOrder.EOF){ 
		Response.Write("Grp4Data[81] = new Array();")	
		while (!rsWorkOrder.EOF) { 
			chrData = rsWorkOrder("chvWork_order_no")
			intID   = rsWorkOrder("insWork_order_id")
	%>
			Grp4Data[81][<%=intID%>] = "<%= chrData %>"
	<%
			rsWorkOrder.MoveNext 
		}
	} else {
		Response.Write("Work Order lookup does not exist.")
	}	
	rsWorkOrder.Close();

	// retrieve the Purchase Status lookup
	var rsPurchase = Server.CreateObject("ADODB.Recordset");
	rsPurchase.ActiveConnection = MM_cnnASP02_STRING;
	rsPurchase.Source = "{call dbo.cp_ASP_Lkup(54)}";
	rsPurchase.CursorType = 0;
	rsPurchase.CursorLocation = 2;
	rsPurchase.LockType = 3;
	rsPurchase.Open();
	// Load the Purchase Status Lookup 
	if (!rsPurchase.EOF){ 
		Response.Write("Grp4Data[83] = new Array();")
	
		while (!rsPurchase.EOF) { 
			chrData = rsPurchase("chvPurchase_name")
			intID   = rsPurchase("insPurchase_sts_id")
	%>
			Grp4Data[83][<%=intID%>] = "<%= chrData %>"
	<%
			rsPurchase.MoveNext 
		}
	} else {
		Response.Write("Purchase Status does not exist.")
	}
	
	rsPurchase.Close();
	%>

	function selectChange(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		for (var q=document.frm14s01.LookupValueSearchOptions.options.length;q>=0;q--) document.frm14s01.LookupValueSearchOptions.options[q]=null;	  	  
		myEle = document.createElement("option") ;
		var y = control.value;
		switch(y){
			case "78" :
				document.frm14s01.StringSearchTextTwo.disabled = false;
				oStg12.style.visibility="visible";
			break;
			case "79":
				oStg22.style.visibility="hidden";	  		
			break;
			case "80", "81", "82", "83":
				oStg22.style.visibility="visible";	  		
			break;
			default :
				document.frm14s01.StringSearchTextTwo.disabled = true;
				oStg12.style.visibility="hidden";				
			break;
		}	  		
		if ((y != 0) && (y != 76) && (y!=85)) {
			for ( x = 0 ; x < ItemArray[y].length  ; x++ ) {
				if (ItemArray[y][x]) { 
					myEle = document.createElement("option") ;
					myEle.value = x ;
					myEle.text = ItemArray[y][x] ;
					controlToPopulate.add(myEle) ;
				}
			}
		}

		document.frm14s01.StringSearchTextOne.value="";
		var j = 0;
		var len = document.frm14s01.SearchType.length;		
		for (var i=0;i <len; i++) {
			if (document.frm14s01.SearchType[i].checked) j = i;
		}
				
		if (j==1) selectChange4(frm14s01.LookupValueSearchOperator, frm14s01.LookupValueSearchOptions,Grp4Data );

		Togo();	  
	}

	function selectChange4(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		myEle = document.createElement("option") ;
		var y = document.frm14s01.LookupValueSearchOperand.value;
		switch (y){
			case "79":
				oStg22.style.visibility = "hidden"
			break;
			case "80":
				oStg22.style.visibility = "visible"		
			break;	
			case "81":
				oStg22.style.visibility = "visible"		
			break;
			case "82":
				oStg22.style.visibility = "visible"		
			break;
			case "83":
				oStg22.style.visibility = "visible"		
			break;		
		}	  
		if ((y != 0) && (y!=79)) {
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
		if (document.frm14s01.StringSearchOperand.selectedIndex <= 0) {
			var j = 0;
		} else {
			var j = document.frm14s01.StringSearchOperand[document.frm14s01.StringSearchOperand.selectedIndex].value ;
		}
		if (document.frm14s01.StringSearchOperator.selectedIndex <= 0) {
			var l = 0;
		} else {
			var l = document.frm14s01.StringSearchOperator[document.frm14s01.StringSearchOperator.selectedIndex].value ;
		}
		document.frm14s01.MM_curOprd.value = j ;
		document.frm14s01.MM_curOptr.value = l ;
		document.frm14s01.MM_flag.value = true ;
	}
	
	</script>
	<script language="JavaScript" src="../js/m014Srh01.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
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
		var j = 0;
		var len = document.frm14s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm14s01.SearchType[i].checked) j = i;
		}
		switch (j) {
			case 0:
				selectChange(frm14s01.StringSearchOperand, frm14s01.StringSearchOperator,detailData);
			break;
			case 1:
				selectChange(frm14s01.LookupValueSearchOperand, frm14s01.LookupValueSearchOperator,detailData );
			break;
		}	   				
	}
	
	function Savtxt() {		
		// Radio box
		var len = document.frm14s01.SearchType.length;
		var Idparam = -1;                 // init.
		var stgTemp,j,k; 
		
		for (var i=0;i <len; i++){
			if (document.frm14s01.SearchType[i].checked) Idparam = i;
		}
	
		stgTemp = document.frm14s01.QueryString.value;
		
		switch ( Idparam ) {
			case 0: 
		  		if (document.frm14s01.StringSearchOperand.length >= 1) {			
					var chvOprd = document.frm14s01.StringSearchOperand[document.frm14s01.StringSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select String Search Operand.");
					return ;
					break;
				}					
				var chrNot  = "";
				if (document.frm14s01.StringSearchOperator.length >= 1) {					
					var chvOptr = document.frm14s01.StringSearchOperator[document.frm14s01.StringSearchOperator.selectedIndex].value ;
				} else {
					alert("Select String Search Operator.");
					return ;
					break;
				}					
				var chvStg1 = document.frm14s01.StringSearchTextOne.value;
				var chvStg2 = document.frm14s01.StringSearchTextTwo.value;
			 	if ((chvOprd=="78") && (chvOptr!="0")) {
					if (chvStg1 == "") {
						alert("Enter Start Date.");
						document.frm14s01.StringSearchTextOne.focus();
						return ;
						return ;
					}
					if (chvStg2 == "") {
						alert("Enter End Date.");
						document.frm14s01.StringSearchTextTwo.focus();
						return ;
					}
					if (!CheckDateBetween(Trim(chvStg1)+" and "+Trim(chvStg2))) {
						alert("End Date is before Start Date");					
						return ;
					}
				}			 
				if (chvOprd=="77") {
					if (!IsID(chvStg1)) {
						alert("Invalid number.");
						document.frm14s01.StringSearchTextOne.focus();
						return ;
					}
					if (chvStg1 > "650354") {
						alert("PR Number cannot be greater than 650354");
						document.frm14s01.StringSearchTextOne.focus();
						return ;
					}
					if (chvOptr=="21") {
						if (!IsID(chvStg2)) {
							alert("Invalid number.");
							document.frm14s01.StringSearchTextTwo.focus();
							return ;							
						}
						if (chvStg2 > "650354") {
							alert("PR Number cannot be greater than 650354");
							document.frm14s01.StringSearchTextTwo.focus();
							return ;
						}						
					}					
				}			 				
			break; 
			case 1: 
				if (document.frm14s01.LookupValueSearchOperand.length >= 1) {	  
					 var chvOprd = document.frm14s01.LookupValueSearchOperand[document.frm14s01.LookupValueSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select Lookup Value Search Operand.");
					return ;
					break;			
				}
				 var chrNot  = "";
				if (document.frm14s01.LookupValueSearchOperator.length >= 1) {	  			 
					 var chvOptr = document.frm14s01.LookupValueSearchOperator[document.frm14s01.LookupValueSearchOperator.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Operator.");
					return ;
					break;			
				}
				
				if (document.frm14s01.LookupValueSearchOperand.value!="79") {
					if (document.frm14s01.LookupValueSearchOptions.length >= 1) {	  			 			
						 var chvStg1 = document.frm14s01.LookupValueSearchOptions[document.frm14s01.LookupValueSearchOptions.selectedIndex].value ;
					} else {
						alert("Select Lookup Value Search Option.");
						return ;
						break;
					}
				}
				var chvStg2 = "";
			break;
			case 2: 
				if (document.frm14s01.ClassSearchOperand.length >= 1) {	  
					 var chvOprd = document.frm14s01.ClassSearchOperand[document.frm14s01.ClassSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select Class Search Operand.");
					return ;
					break;			
				}	  
				var chrNot  = "";
				var chvOptr = "3";
				var chvStg1 = document.frm14s01.ClassSearchText.value ;
				var chvStg2 = document.frm14s01.chrUsr_Type.value;
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
		if (chvOptr == "0") {
			document.write("<B><B><B>");
			document.write("Please select operator before Proceed");
			document.write("<B><B><B>");
		} else {
			var stgFilter = ACfltr_14(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
		}

		var chvAO1  = document.frm14s01.AndOr.value ;
		if (stgTemp.length > 0 ) {   stgTemp += " (" ;  } 
		stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 ;
		document.frm14s01.QueryString.value = stgTemp; 
	}	
	
	function CnstrFltr(output){
		var inspSrtBy = document.frm14s01.SortByColumn.value;
		var inspSrtOrd = document.frm14s01.OrderBy.value;
		var stgFilter = document.frm14s01.QueryString.value;
		if (output==1) {
			document.frm14s01.action = "m014q01.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;
			document.frm14s01.submit() ; 			
		} else {
			var ExcelSearch = window.open("m014q01excel.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter);
		}
	}
	
	function SelectOperator() {
		if (document.frm14s01.StringSearchOperator.value=="21") {
			document.frm14s01.StringSearchTextTwo.disabled = false;
			document.frm14s01.StringSearchTextTwo.value = "";			
			oStg12.style.visibility="visible";		
		} else {
			document.frm14s01.StringSearchTextTwo.disabled = true;
			document.frm14s01.StringSearchTextTwo.value = "";			
			oStg12.style.visibility="hidden";				
		}
		Togo();
	}	

	function SelOpt() {	
		var len = document.frm14s01.SearchType.length;
		var Idparam = 1;                 // init.
	
		for (var i=0;i <len; i++){
			if (document.frm14s01.SearchType[i].checked) Idparam = i;
		}
		switch (Idparam) {
			case 0: 
				oOptrd11.style.visibility="visible";
				oOptr11.style.visibility="visible";
				oStg11.style.visibility="visible";
				oStg12.style.visibility="hidden";
				oOptrd21.style.visibility="hidden";
				oOptr21.style.visibility="hidden";
				oStg22.style.visibility="hidden";
				selectChange(frm14s01.StringSearchOperand, frm14s01.StringSearchOperator,detailData);				
			break;
			case 1: 
				oOptrd11.style.visibility="hidden";
				oOptr11.style.visibility="hidden";
				oStg11.style.visibility="hidden";
				oStg12.style.visibility="hidden";
				oOptrd21.style.visibility="visible";
				oOptr21.style.visibility="visible";
				oStg22.style.visibility="visible";
				selectChange(frm14s01.LookupValueSearchOperand, frm14s01.LookupValueSearchOperator,detailData );				
			break;
			case 2: 
				oOptrd11.style.visibility="hidden";
				oOptr11.style.visibility="hidden";
				oStg11.style.visibility="hidden";
				oStg12.style.visibility="hidden";
				oOptrd21.style.visibility="hidden";
				oOptr21.style.visibility="hidden";
				oStg22.style.visibility="hidden";
			break; 
			default: 
			break;
		}       
	}

	function Toggle() {	
		var idx = document.frm14s01.ClassSearchOperand[document.frm14s01.ClassSearchOperand.selectedIndex].value
		switch ( idx ) {
			//class no 
			case "84":
				openWindow("m014p01FSq.asp","");
			break;
			default: 
				document.frm14s01.ClassSearchText.value = ""; 
			break;
	   }
	}
	</script>
</head>
<body onload="initscr()">
<form name="frm14s01" method="post" action="">
<h3>Purchase Requisition - Advanced Search</h3>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap width="160"><input type="radio" name="SearchType" value="1" tabindex="1" accesskey="F" checked onClick="SelOpt()" onMouseOver="" onMouseOut="" class="chkstyle">String Search</td>
		<td nowrap><DIV ID="oOptrd11" STYLE="visibility:visible">
			<select name="StringSearchOperand" onchange="selectChange(this, frm14s01.StringSearchOperator,detailData);" tabindex="2">
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
		<td nowrap><DIV ID="oOptr11" STYLE="visibility:visible">
			<select name="StringSearchOperator" onChange="SelectOperator();" tabindex="3"></select>
		</DIV></td>
		<td nowrap><DIV ID="oStg11" STYLE="visibility:visible">
			<input type="text" name="StringSearchTextOne" tabindex="4" size="10">
		</DIV></td>
		<td nowrap><DIV ID="oStg12" STYLE="visibility:hidden">
			<input type="text" name="StringSearchTextTwo" value="<%=CurrentDate()%>" tabindex="6" size="10">
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap width="160"><input type="radio" name="SearchType" value="2" onClick="SelOpt()" onMouseOver="" onMouseOut="" tabindex="8" class="chkstyle">Lookup Value Search</td>
		<td nowrap><DIV ID="oOptrd21" STYLE="visibility:visible">
			<select name="LookupValueSearchOperand" onchange="selectChange(this, frm14s01.LookupValueSearchOperator,detailData);" tabindex="9">
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
			<select name="LookupValueSearchOperator" onchange="selectChange4(this, frm14s01.LookupValueSearchOptions,Grp4Data );" tabindex="10"></select>
		</DIV></td>
		<td nowrap><DIV ID="oStg22" STYLE="visibility:visible"> 
			<select name="LookupValueSearchOptions" tabindex="11"></select>
        </DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td valign="top"><select name="AndOr" tabindex="12">
			<option value=" ">None</option>
			<option value="And">And</option>
			<option value="Or">Or</option>
		</select></td>	
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td valign="top" nowrap><input type="button" name="AddToQueryString" value="Add" tabindex="13" onClick="Savtxt()" class="btnstyle"></td>
		<td valign="top" nowrap><textarea name="QueryString" cols="80" rows="6" tabindex="14" accesskey="L"></textarea></td>
    </tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
	<tr>
    	<td nowrap>Sort by:
			<select name="SortByColumn" tabindex="15">
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
        	<select name="OrderBy" tabindex="16">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>
	        <input type="button" value="Search" onClick="CnstrFltr(1);" tabindex="17" class="btnstyle">
			<input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="18" class="btnstyle">
			<input type="reset" value="Clear All" tabindex="19" onClick="window.location.reload();" class="btnstyle">
		</td>		
    </tr>
</table>
<input type="hidden" name="MM_flag" value="false">
<input type="hidden" name="MM_curOprd">
<input type="hidden" name="MM_curOptr">
<input type="hidden" name="chrUsr_Type" value="0">
</form>
</body>
</html>
<%
rsCol.Close();
rsOprd.Close();
rsOprd2.Close();
rsOprd3.Close();
%>