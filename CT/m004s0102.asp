<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" --> 
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
// retrieve the sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup2(732,0,'',0,'',0)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

// retrieve the string search operands - text
var rsOprd = Server.CreateObject("ADODB.Recordset");
rsOprd.ActiveConnection = MM_cnnASP02_STRING;
rsOprd.Source = "{call dbo.cp_ASP_Lkup2(733,0,'',0,'',0)}";
rsOprd.CursorType = 0;
rsOprd.CursorLocation = 2;
rsOprd.LockType = 3;
rsOprd.Open();

// retrieve the lookup value search operands - Combo
var rsOprd2 = Server.CreateObject("ADODB.Recordset");
rsOprd2.ActiveConnection = MM_cnnASP02_STRING;
rsOprd2.Source = "{call dbo.cp_ASP_Lkup2(734,0,'',0,'',0)}";
rsOprd2.CursorType = 0;
rsOprd2.CursorLocation = 2;
rsOprd2.LockType = 3;
rsOprd2.Open();
%>
<html>
<head>
	<title>Contact - Advanced Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="Javascript">
	var detailData = new Array();
	<%
	var intOldOptr = 0;
	var objOptrDesc,objOptrId,objRecID;
	
	var rsOptr = Server.CreateObject("ADODB.Recordset");
	rsOptr.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,19)}";	
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
		Response.Write("SysOptr lookup does not exist.");
	}	
	rsOptr.Close();
	%>

	var Grp4Data   = new Array();
	<%
	// retrieve the Contact Type lookup
	var rsContactType = Server.CreateObject("ADODB.Recordset");
	rsContactType.ActiveConnection = MM_cnnASP02_STRING;
	rsContactType.Source = "{call dbo.cp_work_type(0,'',1,0,'Q',0)}";
	rsContactType.CursorType = 0;
	rsContactType.CursorLocation = 2;
	rsContactType.LockType = 3;
	rsContactType.Open();
	if (!rsContactType.EOF){ 	
		Response.Write("Grp4Data[106] = new Array();")	
		while (!rsContactType.EOF) { 
	%>
			Grp4Data[106][<%=rsContactType("intWork_type_id")%>] = "<%= rsContactType("chvWork_type_desc") %>"
	<%
			rsContactType.MoveNext 
		}
	} else {
	   Response.Write("Contact Type lookup does not exist.")
	}	
	rsContactType.Close();

	// retrieve the Mailing List lookup
	var rsMailingList = Server.CreateObject("ADODB.Recordset");
	rsMailingList.ActiveConnection = MM_cnnASP02_STRING;
	rsMailingList.Source = "{call dbo.cp_mail_list(0,'',1,0,'Q',0)}";
	rsMailingList.CursorType = 0;
	rsMailingList.CursorLocation = 2;
	rsMailingList.LockType = 3;
	rsMailingList.Open();
	if (!rsMailingList.EOF){ 	
		Response.Write("Grp4Data[107] = new Array();")	
		while (!rsMailingList.EOF) { 
	%>
			Grp4Data[107][<%=rsMailingList("insMail_list_id")%>] = "<%= rsMailingList("chvName") %>"
	<%
			rsMailingList.MoveNext 
		}
	} else {
		Response.Write("Request Type lookup does not exist.")
	}	
	rsMailingList.Close();
	%>

	function selectChange(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		for (var q=document.frm04s01.LookupValueSearchOptions.options.length;q>=0;q--) document.frm04s01.LookupValueSearchOptions.options[q]=null;	  	  
		myEle = document.createElement("option") ;
		var y = control.value;
		if ((y != 0) && (y != 102)) {
			for ( x = 0 ; x < ItemArray[y].length  ; x++ ) {
			   if (ItemArray[y][x]) { 
				   myEle = document.createElement("option") ;
				   myEle.value = x ;
				   myEle.text = ItemArray[y][x] ;
				   controlToPopulate.add(myEle) ;
			   }
		   }
		}
		document.frm04s01.StringSearchTextOne.value="";
		var j = 0;
		var len = document.frm04s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm04s01.SearchType[i].checked) j = i;
		}		
		if (j==1) selectChange4(frm04s01.LookupValueSearchOperator, frm04s01.LookupValueSearchOptions,Grp4Data );
		Togo();	  
	}

	function selectChange4(control, controlToPopulate,ItemArray) {
		var myEle ;	  
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		myEle = document.createElement("option") ;
		var y = document.frm04s01.LookupValueSearchOperand.value;
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
		if (document.frm04s01.StringSearchOperand.selectedIndex <= 0) {
			var j = 0;
		} else {
			var j = document.frm04s01.StringSearchOperand[document.frm04s01.StringSearchOperand.selectedIndex].value ;
		}
		if (document.frm04s01.StringSearchOperator.selectedIndex <= 0) {
			var l = 0;
		} else {
			var l = document.frm04s01.StringSearchOperator[document.frm04s01.StringSearchOperator.selectedIndex].value ;
		}
		document.frm04s01.MM_curOprd.value = j ;
		document.frm04s01.MM_curOptr.value = l ;
		document.frm04s01.MM_flag.value = true ;
	}	
	</script>
	<script language="JavaScript" src="../js/m004Srh01.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	   

	function initscr() {
		oOptrd21.style.visibility="hidden";
		oOptr21.style.visibility="hidden";
		oStg21.style.visibility="hidden";
		var j = 0;
		var len = document.frm04s01.SearchType.length;		
		for (var i=0;i <len; i++){
			if (document.frm04s01.SearchType[i].checked) j = i;
		}
		switch (j) {
			case 0:
				selectChange(frm04s01.StringSearchOperand, frm04s01.StringSearchOperator,detailData);
			break;
			case 1:
				selectChange(frm04s01.LookupValueSearchOperand, frm04s01.LookupValueSearchOperator,detailData );
			break;
		}	   								
	}
	
	function Savtxt() {		
		var len = document.frm04s01.SearchType.length;
		var Idparam = 1;                 // init.
		var stgTemp,j,k; 
		
		for (var i=0;i <len; i++){
			if (document.frm04s01.SearchType[i].checked) Idparam = i;
		}
	
		stgTemp = document.frm04s01.QueryString.value;		
		switch ( Idparam ) {
			case 0: 
				if (document.frm04s01.StringSearchOperand.value=="103") {
					if (!IsID(document.frm04s01.StringSearchTextOne.value)) {
						alert("Invalid number.");
						document.frm04s01.StringSearchTextOne.focus();
						return ;
					}
				}			
		  		if (document.frm04s01.StringSearchOperand.length >= 1) {			
					var chvOprd = document.frm04s01.StringSearchOperand[document.frm04s01.StringSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select String Search Operand.");
					return;
					break;
				}					
				var chrNot  = "";
				if (document.frm04s01.StringSearchOperator.length >= 1) {					
					var chvOptr = document.frm04s01.StringSearchOperator[document.frm04s01.StringSearchOperator.selectedIndex].value ;
				} else {
					alert("Select String Search Operator.");
					return ;
					break;
				}					
				var chvStg1 = document.frm04s01.StringSearchTextOne.value;
				var chvStg2 = "";
			break; 
			case 1: 
				if (document.frm04s01.LookupValueSearchOperand.length >= 1) {	  
					 var chvOprd = document.frm04s01.LookupValueSearchOperand[document.frm04s01.LookupValueSearchOperand.selectedIndex].value ; 
				} else {
					alert("Select Lookup Value Search Operand.");
					break;			
				}
				var chrNot  = "";
				if (document.frm04s01.LookupValueSearchOperator.length >= 1) {	  			 
					 var chvOptr = document.frm04s01.LookupValueSearchOperator[document.frm04s01.LookupValueSearchOperator.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Operator.");
					break;			
				}
				if (document.frm04s01.LookupValueSearchOptions.length >= 1) {	  			 			
					 var chvStg1 = document.frm04s01.LookupValueSearchOptions[document.frm04s01.LookupValueSearchOptions.selectedIndex].value ;
				} else {
					alert("Select Lookup Value Search Option.");
					break;
				}
				var chvStg2 = "";
			break;
			default:
				alert("program Error - radio buttion 'Sel' is not picked ...");
			break; 
		}
		if (chvOptr == "0") {
			document.write("<B><B><B>");
			document.write("Please select operator before Proceed");
			document.write("<B><B><B>");
		} else {
			var stgFilter = ACfltr_04(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
		}

		var chvAO1  = document.frm04s01.AndOr.value ;
		if (stgTemp.length > 0 ) stgTemp += " (" ; 
		stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 ;
		document.frm04s01.QueryString.value = stgTemp; 
	}
	
	function CnstrFltr(output){
		var inspSrtBy = document.frm04s01.SortByColumn.value;
		var inspSrtOrd = document.frm04s01.OrderBy.value;
		var stgFilter = document.frm04s01.QueryString.value;
		if (output==1) {
			document.frm04s01.action = "m004q01.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;
			document.frm04s01.submit() ; 			
		} else {
			var ExcelSearch = window.open("m004q01excel.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter);
		}
	}
	
	function SelOpt() {	
		var len = document.frm04s01.SearchType.length;
		var Idparam = 1;                 // init.

		for (var i=0;i <len; i++){
			if (document.frm04s01.SearchType[i].checked) Idparam = i;
		}
		switch ( Idparam ) {
			// text 
			case 0: 
				oOptrd11.style.visibility="visible";
				oOptr11.style.visibility="visible";
				oStg11.style.visibility="visible";
				oOptrd21.style.visibility="hidden";
				oOptr21.style.visibility="hidden";
				oStg21.style.visibility="hidden";
			break;
			//Combo 
			case 1: 
				oOptrd11.style.visibility="hidden";
				oOptr11.style.visibility="hidden";
				oStg11.style.visibility="hidden";
				oOptrd21.style.visibility="visible";
				oOptr21.style.visibility="visible";
				oStg21.style.visibility="visible";
				selectChange(frm04s01.LookupValueSearchOperand, frm04s01.LookupValueSearchOperator,detailData );				
			break;
			default: 
			break;
		}       
	}
	</script>
</head>
<body onload="initscr()" >
<form name="frm04s01" method="post" action="">
<h3>Contact - Advanced Search</h3>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap width="160"><input type="radio" name="SearchType" value="1" tabindex="1" accesskey="F" checked onClick="SelOpt()" class="chkstyle">String Search</td>
		<td nowrap><DIV ID="oOptrd11" STYLE="visibility:visible">
			<select name="StringSearchOperand" onchange="selectChange(this, frm04s01.StringSearchOperator,detailData);" tabindex="2" style="width: 150px">
			<% 
			while (!rsOprd.EOF) { 
			%>
				<option value="<%=(rsOprd.Fields.Item("intRecID").Value)%>" <%=((rsOprd.Fields.Item("intRecID").Value == 105)?"SELECTED":"")%> ><%=(rsOprd.Fields.Item("chvObjName").Value)%></option>
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
			<input type="text" name="StringSearchTextOne" tabindex="4">
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">	
    <tr> 
		<td nowrap width="160"><input type="radio" name="SearchType" value="2" onClick="SelOpt()" tabindex="5" class="chkstyle">Lookup Value Search</td>
		<td nowrap><DIV ID="oOptrd21" STYLE="visibility:visible">
			<select name="LookupValueSearchOperand" onchange="selectChange(this, frm04s01.LookupValueSearchOperator,detailData);" tabindex="6" style="width: 180px">
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
			<select name="LookupValueSearchOperator" onchange="selectChange4(this, frm04s01.LookupValueSearchOptions,Grp4Data );" tabindex="7"></select>
		</DIV></td>
		<td nowrap><DIV ID="oStg21" STYLE="visibility:visible"> 
			<select name="LookupValueSearchOptions" tabindex="8"></select>
        </DIV></td>		
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap valign="top">
			<select name="AndOr" tabindex="9">
				<option value=" ">None</option>
				<option value="And">And</option>
				<option value="Or">Or</option>
			</select>
		</td>	
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td valign="top" nowrap><input type="button" name="AddToQueryString" value="Add" tabindex="10" onClick="Savtxt()" class="btnstyle"></td>
		<td valign="top" nowrap><textarea name="QueryString" cols="80" rows="6" tabindex="11" accesskey="L"></textarea></td>
    </tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Sort by:
			<select name="SortByColumn" tabindex="12">
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
        	<select name="OrderBy" tabindex="13">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>
	        <input type="button" value="Search" onClick="CnstrFltr(1);" tabindex="14" class="btnstyle">
			<input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="15" class="btnstyle">
			<input type="reset" value="Clear All" tabindex="16" onClick="window.location.reload();" class="btnstyle">
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
rsCol.Close();
rsOprd.Close();
rsOprd2.Close();
%>