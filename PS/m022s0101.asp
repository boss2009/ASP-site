<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% 
// retrieve the sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_asP_lkup2(746,0,'',0,'',0)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

// retrieve the string search operands
var rsOprd1 = Server.CreateObject("ADODB.Recordset");
rsOprd1.ActiveConnection = MM_cnnASP02_STRING;
rsOprd1.Source = "{call dbo.cp_asP_lkup2(747,0,'',0,'',0)}";
rsOprd1.CursorType = 0;
rsOprd1.CursorLocation = 2;
rsOprd1.LockType = 3;
rsOprd1.Open();

// retrieve the multi-select search operands
var rsOprd2 = Server.CreateObject("ADODB.Recordset");
rsOprd2.ActiveConnection = MM_cnnASP02_STRING;
rsOprd2.Source = "{call dbo.cp_asP_lkup2(748,0,'',0,'',0)}";
rsOprd2.CursorType = 0;
rsOprd2.CursorLocation = 2;
rsOprd2.LockType = 3;
rsOprd2.Open();
%>
<html>
<head>
	<title>Temp Student - Quick Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/m022Srh01.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="javascript">
	function CnstrFltr(output) {		
		var stgPgQuery = "";
		var stgFilter = "" ;
		var len = document.frm22s01.SearchType.length;
		var Idparam = 1;	
		var stgTemp,j,k; 
	  	
		for (var i=0;i <len; i++){
			if (document.frm22s01.SearchType[i].checked) Idparam = i;
		}
	
		switch (Idparam) {
			case 0: 
				if (document.frm22s01.StringSearchOperand.length >= 1) {
					j = document.frm22s01.StringSearchOperand[document.frm22s01.StringSearchOperand.selectedIndex].value ;
				} else {
					alert("Select String Search Operand.");
					document.frm22s01.StringSearchOperand.focus();
					break;
				}
				if (document.frm22s01.StringSearchOperator.length >= 1) {			
					l = document.frm22s01.StringSearchOperator[document.frm22s01.StringSearchOperator.selectedIndex].value ;
				} else {
					alert("Select String Search Operator.");
					document.frm22s01.StringSearchOperator.focus();
					break;
				}
				var chvOprd = j;
				var chrNot  = "";
				var chvOptr = l ;
				var chvStg1 = document.frm22s01.StringSearchTextOne.value;
				var chvStg2 = "";
				if (j=="158") {
					if (!IsID(chvStg1)) {
						alert("Invalid number.");
						document.frm22s01.StringSearchTextOne.focus();
						return ;
					}
				}				
				if (chvOptr == "0") {
					alert("Please select operator before Proceed");
				} else {
					stgFilter = ACfltr_22(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
				}
			break;
			case 1: 
				j = document.frm22s01.MultiValueSearchOperand[document.frm22s01.MultiValueSearchOperand.selectedIndex].value ;
				l = "";
				var count = 0;
				var optList = document.frm22s01.MultiValueSearchOptions;
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
						if (document.frm22s01.MultiValueSearchOptions[ii].selected) {		
							if (l.length > 0 ) l += "," ;
							l += document.frm22s01.MultiValueSearchOptions[ii].value ;		
					 	} 				  
				  	} 						  
				} else {
					l = document.frm22s01.MultiValueSearchOptions[document.frm22s01.MultiValueSearchOptions.selectedIndex].value
				} 
				switch (j) {
					//  Disability
					case "28" : 
						stgFilter += " insDsbty1_id in (" + l + ") " ; 
					break;
					// Student Status
					case "29" : 
						stgFilter += " insStdnt_status_id in (" + l + ") " ; 
					break;
				}
			break;
			default: 
			break;
		}						
		document.frm22s01.StringSearchTextOne.value = "";
	
		var inspSrtBy = document.frm22s01.SortByColumn.value;
		var inspSrtOrd = document.frm22s01.OrderBy.value;
		
		switch(output) {
			case 1:
				document.frm22s01.action = "m022q01.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;				
			break;
			case 2:
				document.frm22s01.action = "m022q01excel.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter;
				document.frm22s01.target = "_blank";				
				document.frm22s01.submit() ; 				
			break;
		}
	}

	function initscr() {
		oOptrd1.style.visibility="visible";
		oStg1.style.visibility="visible";
		oOptrd2.style.visibility="hidden";
		oOptr2.style.visibility="hidden";
		initializeList('frm22s01.MultiValueSearchOperand','frm22s01.MultiValueSearchOptions');	
		document.frm22s01.StringSearchOperand.focus();		
		selectChange(frm22s01.StringSearchOperand, frm22s01.StringSearchOperator,detailData1);		
	}

	function Selckb(){
		var j = document.frm22s01.MultiValueSearchOperand[document.frm22s01.MultiValueSearchOperand.selectedIndex].value ;
		var blnFlag = true;
	}

	function SelOpt() {
		var len = document.frm22s01.SearchType.length;
		var Idparam = 1;
		
		for (var i=0;i <len; i++){
			if (document.frm22s01.SearchType[i].checked) Idparam = i;
		}
	
		switch (Idparam) {
			case 0: 
				window.status = "Single Term";
				oOptrd1.style.visibility="visible";
				oStg1.style.visibility="visible";
				oOptrd2.style.visibility="hidden";
				oOptr2.style.visibility="hidden";
			break;
			case 1: 
				window.status = "Multiple Term";			
				oOptrd1.style.visibility="hidden";
				oStg1.style.visibility="hidden";
				oOptrd2.style.visibility="visible";
				oOptr2.style.visibility="visible";
			break;
		}       	
	}
	</script>
	<script language="javascript">
	var detailData1 = new Array();
	<%
	var intOldOptr = 0;
	var objOptrDesc,objOptrId,objRecID;
	
	var rsOptr1 = Server.CreateObject("ADODB.Recordset");
	rsOptr1.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr1.Source = "{call dbo.cp_SysOptr(0,0,34)}";
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
	   var j = document.frm22s01.StringSearchOperand[document.frm22s01.StringSearchOperand.selectedIndex].value ;
	   var l = document.frm22s01.StringSearchOperator[document.frm22s01.StringSearchOperator.selectedIndex].value
	   document.frm22s01.MM_curOprd.value = j ;
	   document.frm22s01.MM_curOptr.value = l ;
	   document.frm22s01.MM_flag.value = true ;
	}
	
	var detailData2 = new Array();
	<%
	var rsDisability = Server.CreateObject("ADODB.Recordset");
	rsDisability.ActiveConnection = MM_cnnASP02_STRING;
	rsDisability.Source = "{call dbo.cp_StdDsbty}";
	rsDisability.CursorType = 0;
	rsDisability.CursorLocation = 2;
	rsDisability.LockType = 3;
	rsDisability.Open();
	
	if (!rsDisability.EOF){ 
		Response.Write("detailData2[160] = new Array();")	
		while (!rsDisability.EOF){ 
			Response.Write("detailData2[160][" + rsDisability("insDisability_id") + "] = '" + rsDisability("chvname") + "';");
			rsDisability.MoveNext 		  
		}
	} else {
		Response.Write("Disability lookup does not exist.")
	}
	rsDisability.Close();

	var rsStatus = Server.CreateObject("ADODB.Recordset");
	rsStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsStatus.Source = "{call dbo.cp_StdStatus}";
	rsStatus.CursorType = 0;
	rsStatus.CursorLocation = 2;
	rsStatus.LockType = 3;
	rsStatus.Open();
	
	if (!rsStatus.EOF){ 
		Response.Write("detailData2[161] = new Array();")
		while (!rsStatus.EOF){ 
			Response.Write("detailData2[161]["+rsStatus("insStdnt_status_id")+"] = '"+rsStatus("chvName")+"';");
			rsStatus.MoveNext 
		}
	} else {
		Response.Write("Status lookup does not exist.")
	}
	rsStatus.Close();
	%>
	
	var objParent ;
	var objChild  ;

	function initializeList(parentList, childList){
		objParent = eval("document." + parentList);
		objChild = eval("document." + childList);
		if (objParent.selectedIndex < 0) objParent.options[0].selected = true;
		var OptnID = objParent.options[objParent.selectedIndex].value;	
		clearList(objChild)
		fillDetail(OptnID)
	}

	function fillDetail(deptID){
		if (detailData2[deptID]){
			for (x = 0; x < detailData2[deptID].length; x++){
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
<form name="frm22s01" method="post" onSubmit="CnstrFltr(1);">
<h3>Temp Student - Quick Search</h3>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top" width="150"><input type="radio" name="SearchType" value="1" checked onClick="SelOpt()" accesskey="F" tabindex="1" class="chkstyle">String Search</td>
		<td nowrap valign="top"><DIV ID="oOptrd1" STYLE="visibility:visible">
			<select name="StringSearchOperand" onChange="selectChange(this, frm22s01.StringSearchOperator,detailData1);" tabindex="2">
			<% 
			while (!rsOprd1.EOF) {
			%>
				<option value="<%=(rsOprd1.Fields.Item("intRecID").Value)%>" <%=((rsOprd1.Fields.Item("intRecID").Value == "163")?"SELECTED":"")%> ><%=(rsOprd1.Fields.Item("chvObjName").Value)%></option>
			<%
				rsOprd1.MoveNext();
			}
			%>
			</select>
		</DIV></td>
		<td nowrap><DIV ID="oStg1" STYLE="visibility:visible">		
			<select name="StringSearchOperator" onChange="Togo();" tabindex="4"></select>
			<input type="text" name="StringSearchTextOne" tabindex="5">
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">	
    <tr> 
		<td nowrap valign="top" width="150"><input type="radio" name="SearchType" value="2" onClick="SelOpt()" tabindex="9" class="chkstyle">Multi-Value Search</td>
		<td nowrap valign="top"><DIV ID="oOptrd2" STYLE="visibility:visible"> 
			<select name="MultiValueSearchOperand" OnChange="initializeList('frm22s01.MultiValueSearchOperand','frm22s01.MultiValueSearchOptions')" tabindex="10">
			<% 
			while (!rsOprd2.EOF) {
			%>
				<option value="<%=(rsOprd2.Fields.Item("intRecID").Value)%>" <%=((rsOprd2.Fields.Item("intRecID").Value == Request.Form("MM_curOprd"))?"SELECTED":"")%> ><%=(rsOprd2.Fields.Item("chvObjName").Value)%></option>
			<%
				rsOprd2.MoveNext();
			}
			%>
			</select>
		</DIV></td>
		<td nowrap valign="top"><DIV ID="oOptr2" STYLE="visibility:hidden"> 
			contains: <select name="MultiValueSearchOptions" size="8" style="width: 150px" multiple align="top" tabindex="12"></select>
		</DIV></td>
    </tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Sort by:
			<select name="SortByColumn" tabindex="16">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("insObjOrder").Value == 1)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvObjName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>
			</select>
        	<select name="OrderBy" tabindex="17">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>
	        <input type="submit" value="Search" tabindex="18" class="btnstyle">
			<input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="19" class="btnstyle">
			<input type="reset" value="Clear All" tabindex="21" onClick="window.location.reload();" class="btnstyle">
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
rsOprd2.Close();
rsOprd1.Close();
rsCol.Close();
%>