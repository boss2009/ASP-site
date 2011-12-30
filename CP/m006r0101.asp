<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
// retrieve the sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup2(735,0,'',0,'Q',0)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

// retrieve the string search operands - text
var rsOprd = Server.CreateObject("ADODB.Recordset");
rsOprd.ActiveConnection = MM_cnnASP02_STRING;
rsOprd.Source = "{call dbo.cp_ASP_Lkup2(736,0,'',0,'Q',0)}";
rsOprd.CursorType = 0;
rsOprd.CursorLocation = 2;
rsOprd.LockType = 3;
rsOprd.Open();

// retrieve lookup value search operands
var rsOprd2 = Server.CreateObject("ADODB.Recordset");
rsOprd2.ActiveConnection = MM_cnnASP02_STRING;
rsOprd2.Source = "{call dbo.cp_ASP_Lkup2(737,0,'',0,'Q',0)}";
rsOprd2.CursorType = 0;
rsOprd2.CursorLocation = 2;
rsOprd2.LockType = 3;
rsOprd2.Open();

// retrieve class search operands
var rsOprd3 = Server.CreateObject("ADODB.Recordset");
rsOprd3.ActiveConnection = MM_cnnASP02_STRING;
rsOprd3.Source = "{call dbo.cp_ASP_Lkup2(738,0,'',0,'Q',0)}";
rsOprd3.CursorType = 0;
rsOprd3.CursorLocation = 2;
rsOprd3.LockType = 3;
rsOprd3.Open();
%>
<html>
<head>
	<title>Organizations - Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="Javascript">
	var detailData = new Array();
	<%
	var intOldOptr = 0;
	var objOptrDesc,objOptrId,objRecID;
	
	var rsOptr = Server.CreateObject("ADODB.Recordset");
	rsOptr.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,21)}";	
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
	}
	else {
		Response.Write("SysOptr lookup does not exist.")
	}
	
	rsOptr.Close();
	%>

	var Grp4Data   = new Array();
	<%
	// retrieve the Work Type lookup
	var rsWorkType = Server.CreateObject("ADODB.Recordset");
	rsWorkType.ActiveConnection = MM_cnnASP02_STRING;
	rsWorkType.Source = "{call dbo.cp_work_type(0,'',1,0,'Q',0)}";
	rsWorkType.CursorType = 0;
	rsWorkType.CursorLocation = 2;
	rsWorkType.LockType = 3;
	rsWorkType.Open();
	// Load the Work Type Lookup 
	if (!rsWorkType.EOF){ 	
		Response.Write("Grp4Data[112] = new Array();")	
		while (!rsWorkType.EOF) { 
	%>
			Grp4Data[112][<%=rsWorkType("intWork_type_id")%>] = "<%= rsWorkType("chvWork_type_desc") %>"
	<%
			rsWorkType.MoveNext 
		}
	} else {
		Response.Write("Work Type lookup does not exist.")
	}
	
	rsWorkType.Close();
	%>

	function selectChange(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		for (var q=document.frm06s01.LookupValueSearchOptions.options.length;q>=0;q--) document.frm06s01.LookupValueSearchOptions.options[q]=null;	  	  
		myEle = document.createElement("option") ;
		var y = control.value;
		switch(y){
			case "112":
				document.frm06s01.LookupValueSearchOptions.disabled = false;
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
		document.frm06s01.StringSearchTextOne.value = "";
		var j = 0;
		var len = document.frm06s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm06s01.SearchType[i].checked) j = i;
		}
				
		if (j==1) selectChange4(frm06s01.LookupValueSearchOperator, frm06s01.LookupValueSearchOptions,Grp4Data );
		Togo();	  
	}

	function selectChange4(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		myEle = document.createElement("option") ;
		var y = document.frm06s01.LookupValueSearchOperand.value;
		if ((y != 0) && (y!=79)) {
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
		if (document.frm06s01.StringSearchOperand.selectedIndex <= 0) {
			var j = 0;
		} else {
			var j = document.frm06s01.StringSearchOperand[document.frm06s01.StringSearchOperand.selectedIndex].value ;
		}
		if (document.frm06s01.StringSearchOperator.selectedIndex <= 0) {
			var l = 0;
		} else {
			var l = document.frm06s01.StringSearchOperator[document.frm06s01.StringSearchOperator.selectedIndex].value ;
		}
		document.frm06s01.MM_curOprd.value = j ;
		document.frm06s01.MM_curOptr.value = l ;
		document.frm06s01.MM_flag.value = true ;
	}
	
	</script>
	<script language="JavaScript" src="../js/m006Srh01.js"></script>
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
		var len = document.frm06s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm06s01.SearchType[i].checked) j = i;
		}
		switch (j) {
			case 0:
				selectChange(frm06s01.StringSearchOperand, frm06s01.StringSearchOperator,detailData);
			break;
			case 1:
				selectChange(frm06s01.LookupValueSearchOperand, frm06s01.LookupValueSearchOperator,detailData );
			break;
		}	   				
	}

	function CnstrFltr(output) {	
		var len = document.frm06s01.SearchType.length;
		var Idparam = -1;                 // init.
		var stgTemp,j,k; 
		
		for (var i=0;i <len; i++){
			if (document.frm06s01.SearchType[i].checked) Idparam = i;
		}
	
		switch ( Idparam ) {
			case 0: 
		  		if (document.frm06s01.StringSearchOperand.length >= 1) {			
					var chvOprd = document.frm06s01.StringSearchOperand[document.frm06s01.StringSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select String Search Operand.");
					return ;
					break;
				}					
				var chrNot  = "";
				if (document.frm06s01.StringSearchOperator.length >= 1) {					
					var chvOptr = document.frm06s01.StringSearchOperator[document.frm06s01.StringSearchOperator.selectedIndex].value ;
				} else {
					alert("Select String Search Operator.");
					return ;
					break;
				}					
				var chvStg1 = document.frm06s01.StringSearchTextOne.value;
				var chvStg2 = "";
			break; 
			case 1: 
				if (document.frm06s01.LookupValueSearchOperand.length >= 1) {	  
					 var chvOprd = document.frm06s01.LookupValueSearchOperand[document.frm06s01.LookupValueSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select Lookup Value Search Operand.");
					return ;
					break;			
				}
				 var chrNot  = "";
				if (document.frm06s01.LookupValueSearchOperator.length >= 1) {	  			 
					 var chvOptr = document.frm06s01.LookupValueSearchOperator[document.frm06s01.LookupValueSearchOperator.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Operator.");
					return ;
					break;			
				}
				if (document.frm06s01.LookupValueSearchOptions.length >= 1) {	  			 			
					 var chvStg1 = document.frm06s01.LookupValueSearchOptions[document.frm06s01.LookupValueSearchOptions.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Option.");
					return ;
					break;
				}
				var chvStg2 = "";
			break;
			case 2: 
				if (document.frm06s01.ClassSearchOperand.length >= 1) {	  
					var chvOprd = document.frm06s01.ClassSearchOperand[document.frm06s01.ClassSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select Class Search Operand.");
					return ;
					break;			
				}	  
				var chrNot  = "";
				var chvOptr = "3";
				var chvStg1 = document.frm06s01.ClassID.value ;
				var chvStg2 = document.frm06s01.chrUsr_Type.value;
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
			var stgFilter = ACfltr_06(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
			var inspSrtBy = document.frm06s01.SortByColumn.value;
			var inspSrtOrd = document.frm06s01.OrderBy.value;
			if (output==1) {
				document.frm06s01.action = "m006q02.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;
				document.frm06s01.submit() ; 					
			} else {
				var ExcelSearch = window.open("m006q02excel.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter);
			}
		}
	}

	function SelOpt() {	
		var len = document.frm06s01.SearchType.length;
		var Idparam = 1;                 // init.
		
		for (var i=0;i <len; i++){
			if (document.frm06s01.SearchType[i].checked) Idparam = i;
		}
		switch (Idparam) {
			// text 
			case 0: 
				oOptrd11.style.visibility="visible";
				oOptrd21.style.visibility="hidden";
				oOptrd31.style.visibility="hidden";
				selectChange(frm06s01.StringSearchOperand, frm06s01.StringSearchOperator,detailData);
			break;
			//Combo 
			case 1: 
				oOptrd11.style.visibility="hidden";
				oOptrd21.style.visibility="visible";
				oOptrd31.style.visibility="hidden";
				selectChange(frm06s01.LookupValueSearchOperand, frm06s01.LookupValueSearchOperator,detailData );
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
		var idx = document.frm06s01.ClassSearchOperand[document.frm06s01.ClassSearchOperand.selectedIndex].value;
		switch ( idx ) {
			// class no 
			case "109":
				openWindow("m006p01FS.asp","");
			break;
			default: 
				document.frm06s01.ClassText.value = ""; 
			break;
		}
	}
	</script>
</head>
<body onload="initscr()" >
<form name="frm06s01" method="post">
<h3>Organizations - Report</h3>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap width="160"><input type="radio" name="SearchType" value="1" tabindex="1" accesskey="F" checked onClick="SelOpt()" class="chkstyle">String Search</td>
		<td nowrap><DIV ID="oOptrd11" STYLE="visibility:visible">
			<select name="StringSearchOperand" onchange="selectChange(this, frm06s01.StringSearchOperator,detailData);" tabindex="2">
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
    <tr> 
		<td nowrap width="160"><input type="radio" name="SearchType" value="2" onClick="SelOpt()" tabindex="5" class="chkstyle">Lookup Value Search</td>
		<td nowrap><DIV ID="oOptrd21" STYLE="visibility:visible">
			<select name="LookupValueSearchOperand" onchange="selectChange(this, frm06s01.LookupValueSearchOperator,detailData);" tabindex="6">
			<% 
			while (!rsOprd2.EOF) { 
			%>
				<option value="<%=(rsOprd2.Fields.Item("intRecID").Value)%>" <%=((rsOprd2.Fields.Item("intRecID").Value == 54)?"SELECTED":"")%> ><%=(rsOprd2.Fields.Item("chvObjName").Value)%></option>
			<% 
				rsOprd2.MoveNext();
			}
			%>
			</select>
			<select name="LookupValueSearchOperator" onchange="selectChange4(this, frm06s01.LookupValueSearchOptions,Grp4Data );" tabindex="7"></select>
			<select name="LookupValueSearchOptions" tabindex="8"></select>
        </DIV></td>
    </tr>
    <tr> 
		<td nowrap width="160"><input type="radio" name="SearchType" value="3" onClick="SelOpt()" tabindex="9" class="chkstyle">Class Search</td>
		<td nowrap><DIV ID="oOptrd31" STYLE="visibility:visible"> 
			<select name="ClassSearchOperand" tabindex="10">
			<% 
			while (!rsOprd3.EOF) {
			%>
				<option value="<%=(rsOprd3.Fields.Item("intRecID").Value)%>" <%=((rsOprd3.Fields.Item("intRecID").Value == 86)?"SELECTED":"")%> ><%=(rsOprd3.Fields.Item("chvObjName").Value)%></option>
			<%
				rsOprd3.MoveNext();
			}
			%>
			</select>
			is
			<input type="text" name="ClassName" size="30" READONLY tabindex="11">
			<input type="button" name="ClassSearchPickList" value="List" onClick="Toggle();" tabindex="12" class="btnstyle">
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>
			Sort by:
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
rsOprd3.Close();
%>