<!--------------------------------------------------------------------------
* File Name: m012p0301.asp
* Title: Contact Search
* Main SP: cp_contact
* Description: This page is used to find a contact.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" --> 
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
// retrieve search operands
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

if (String(Request.Form("State"))=="Search") {
	var rsContact__inspSrtBy = "1";
	if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
	  rsContact__inspSrtBy = String(Request.QueryString("inspSrtBy"));
	}
	var rsContact__inspSrtOrd = "0";
	if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
	  rsContact__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
	}
	
	var rsContact__chvFilter = "";
	if(String(Request.QueryString("chvFilter")) != "undefined") { 
	  rsContact__chvFilter = String(Request.QueryString("chvFilter"));
	}
	
	var rsContact = Server.CreateObject("ADODB.Recordset");
	rsContact.ActiveConnection = MM_cnnASP02_STRING;
	rsContact.Source = "{call dbo.cp_Contacts(0,0,'','','',0,0,0,"+rsContact__inspSrtBy.replace(/'/g, "''")+","+rsContact__inspSrtOrd.replace(/'/g, "''")+",'"+rsContact__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
	rsContact.CursorType = 0;
	rsContact.CursorLocation = 2;
	rsContact.LockType = 3;
	rsContact.Open();
}
%>
<html>
<head>
	<title>Search Contact</title>
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
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,19)}";
	
	rsOptr.CursorType = 0;
	rsOptr.CursorLocation = 2;
	rsOptr.LockType = 3;
	rsOptr.Open();
	var rsOptr_numRows = 0;
	
	// Load the Operators Lookup 
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

	function selectChange(control, controlToPopulate,ItemArray) {
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		myEle = document.createElement("option") ; 
		var y = control.value;
		if ((y != 0)  && (y != 102)) {
			for ( x = 0 ; x < ItemArray[y].length  ; x++ ){
				if (ItemArray[y][x]) { 
					myEle = document.createElement("option") ;
					myEle.value = x ;
					myEle.text = ItemArray[y][x] ;
					controlToPopulate.add(myEle) ;
				}
			}
		}
		document.frm0301.StringSearchTextOne.value = "";
		Togo();	  
	}

	// ---------------------------------
	// function Togo
	// ---------------------------------
	function Togo() {		
		if (document.frm0301.StringSearchOperand.selectedIndex <= 0) {
			var j = 0;
		} else {
			var j = document.frm0301.StringSearchOperand[document.frm0301.StringSearchOperand.selectedIndex].value ;
		}
		if (document.frm0301.StringSearchOperator.selectedIndex <= 0) {
			var l = 0;
		} else {
			var l = document.frm0301.StringSearchOperator[document.frm0301.StringSearchOperator.selectedIndex].value ;
		}
		document.frm0301.MM_curOprd.value = j ;
		document.frm0301.MM_curOptr.value = l ;
		document.frm0301.MM_flag.value = true ;
	}
	
	</script>
	<!--
	// + Oct.12.2001
	-->
	<script language="JavaScript" src="../js/m004Srh01.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=20,top=20,status=1");
		return ;
	}	   
	// ---------------------------------
	// function CnstrFltr() 
	// ---------------------------------
	function CnstrFltr(output) {	
		if (document.frm0301.StringSearchOperand.value=="103") {
			if (isNaN(document.frm0301.StringSearchTextOne.value)) {
				alert("Invalid Contact Number.");
				return ;
			}
		}
		
		if (document.frm0301.StringSearchOperand.length >= 1) {			
			var chvOprd = document.frm0301.StringSearchOperand[document.frm0301.StringSearchOperand.selectedIndex].value ; 
		} else {
			alert("Select String Search Operand.");
			return ;
		}					
		var chrNot  = "";
		if (document.frm0301.StringSearchOperator.length >= 1) {					
			var chvOptr = document.frm0301.StringSearchOperator[document.frm0301.StringSearchOperator.selectedIndex].value ;
		} else {
			alert("Select String Search Operator.");
			return ;
		}					
		var chvStg1 = document.frm0301.StringSearchTextOne.value;
		var chvStg2 = "";

		if (chvOprd == "103") {
			if (!IsID(chvStg1)) {
				alert("Invalid number.");
				document.frm0301.StringSearchTextOne.focus();
				return;			
			}
		}

		if (chvOptr == "0") {
			document.write("<B><B><B>");
			document.write("Please select operator before Proceed");
			document.write("<B><B><B>");
		} else {
			var stgFilter = ACfltr_04(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
			document.frm0301.action = "m012p0301.asp?inspSrtBy=1&inspSrtOrd=0&chvFilter=" + stgFilter ;
			document.frm0301.State.value = "Search";
			document.frm0301.submit() ; 					
		}
	}
	
	function ViewContact(){
		contact_id = document.frm0301.SearchResult.value;
		if (contact_id > 0) {
			openWindow('../CT/m004FS3.asp?intContact_id='+contact_id);
		} else {
			alert("Select a contact.");
			document.frm0301.SearchResult.focus();
			return ;
		}
	}
		
	function SelectContact(){
		if (document.frm0301.SearchResult.selectedIndex==-1){
			alert("Select a contact.")
			return ;
		}	
	
		opener.document.frm12s01.ContactID.value=document.frm0301.SearchResult[document.frm0301.SearchResult.selectedIndex].value;
		opener.document.frm12s01.StringSearchTextboxOne.value=document.frm0301.SearchResult.options[document.frm0301.SearchResult.selectedIndex].text;
		self.close();
	}
			
	function Init(){
		selectChange(frm0301.StringSearchOperand, frm0301.StringSearchOperator,detailData);
	}	
	</script>
</head>
<body onload="Init();">
<form name="frm0301" method="POST" action="">
<h5>Search Contact</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>
			<select name="StringSearchOperand" onchange="selectChange(this, frm0301.StringSearchOperator,detailData);" tabindex="1" accesskey="F">
				<% 
				while (!rsOprd.EOF) { 
				%>
					<option value="<%=(rsOprd.Fields.Item("intRecID").Value)%>" <%=((rsOprd.Fields.Item("intRecID").Value == Request.Form("MM_curOprd"))?"SELECTED":"")%> ><%=(rsOprd.Fields.Item("chvObjName").Value)%></option>
				<% 
					rsOprd.MoveNext(); 
				}
				%>
			</select>
			<select name="StringSearchOperator" onchange="Togo();" tabindex="2"></select>
			<input type="text" name="StringSearchTextOne" tabindex="3">
			<input type="button" value="Search" onClick="CnstrFltr(1);" tabindex="4" class="btnstyle">
		</td>
    </tr>	
</table><br>
<select name="SearchResult" size="20"  style="width: 600px; height: 280px; font-family: Courier;" ondblclick="ViewContact();" tabindex="5" accesskey="L">
<%
var count=0;
if (String(Request.Form("State"))=="Search") { 
	while (!rsContact.EOF) {
%>
	<option value="<%=rsContact.Fields.Item("intContact_id").Value%>"><%=(rsContact.Fields.Item("chvLst_Name").Value+", "+rsContact.Fields.Item("chvFst_Name").Value + " (" + rsContact.Fields.Item("chvWork_type_desc").Value + ", " + rsContact.Fields.Item("chvJob_Title").Value + ")")%>
<%
		count++;
		rsContact.MoveNext();
	}
}
%>
</select>
<br>
<input type="button" value="Select Contact" <%=((count==0)?"DISABLED":"")%> onClick="SelectContact();" tabindex="8" class="btnstyle">
<input type="button" value="Cancel" onClick="window.close()" tabindex="10" class="btnstyle">
<input type="hidden" name="MM_flag" value="false">
<input type="hidden" name="MM_curOprd">
<input type="hidden" name="MM_curOptr">
<input type="hidden" name="State">
</form>
</body>
</html>
<%
rsOprd.Close();
rsOprd2.Close();
%>