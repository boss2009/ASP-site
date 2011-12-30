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
%>
<%
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

switch (String(Request.QueryString("LinkToClass"))){
	//client
	case "1":
		var rsRelationship = Server.CreateObject("ADODB.Recordset");
		rsRelationship.ActiveConnection = MM_cnnASP02_STRING;
		rsRelationship.Source = "{call dbo.cp_relationship2(0,'',3,1,0,'Q',0)}";
		rsRelationship.CursorType = 0;
		rsRelationship.CursorLocation = 2;
		rsRelationship.LockType = 3;
		rsRelationship.Open();
	break;
	//institution
	case "3":
		var rsRelationship = Server.CreateObject("ADODB.Recordset");
		rsRelationship.ActiveConnection = MM_cnnASP02_STRING;
		rsRelationship.Source = "{call dbo.cp_relationship2(0,'',12,1,0,'Q',0)}";
		rsRelationship.CursorType = 0;
		rsRelationship.CursorLocation = 2;
		rsRelationship.LockType = 3;
		rsRelationship.Open();
	break;
	//show nothing
	default:
		var rsRelationship = Server.CreateObject("ADODB.Recordset");
		rsRelationship.ActiveConnection = MM_cnnASP02_STRING;
		rsRelationship.Source = "{call dbo.cp_relationship2(0,'',1,1,0,'Q',0)}";
		rsRelationship.CursorType = 0;
		rsRelationship.CursorLocation = 2;
		rsRelationship.LockType = 3;
		rsRelationship.Open();
	break;		
}
%>
<html>
<head>
	<title>Search For Existing Contact</title>
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
			for ( x = 0 ; x < ItemArray[y].length  ; x++ ) {
				if (ItemArray[y][x]) { 
					myEle = document.createElement("option") ;
					myEle.value = x ;
					myEle.text = ItemArray[y][x] ;
					controlToPopulate.add(myEle) ;
				}
			}
		}
		document.frm0101.StringSearchTextOne.value = "";
	}
	</script>
	<script language="JavaScript" src="../js/m004Srh01.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=20,top=20,status=1");
		return ;
	}	   

	function CnstrFltr(output) {	
		if (document.frm0101.StringSearchOperand.value=="103") {
			if (isNaN(document.frm0101.StringSearchTextOne.value)) {
				alert("Invalid Contact Number.");
				return ;
			}
		}		
		if (document.frm0101.StringSearchOperand.length >= 1) {			
			var chvOprd = document.frm0101.StringSearchOperand[document.frm0101.StringSearchOperand.selectedIndex].value ; 
		} else {
			alert("Select String Search Operand.");
			return ;
		}					
		var chrNot  = "";
		if (document.frm0101.StringSearchOperator.length >= 1) {					
			var chvOptr = document.frm0101.StringSearchOperator[document.frm0101.StringSearchOperator.selectedIndex].value ;
		} else {
			alert("Select String Search Operator.");
			return ;
		}					
		var chvStg1 = document.frm0101.StringSearchTextOne.value;
		var chvStg2 = "";
		if (chvOptr == "0") {
			document.write("<B><B><B>");
			document.write("Please select operator before Proceed");
			document.write("<B><B><B>");
		} else {
			var stgFilter = ACfltr_04(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
			document.frm0101.action = "m004a0101.asp?LinkToObject="+document.frm0101.LinkToObject.value+"&LinkToClass="+document.frm0101.LinkToClass.value+"&WorkType="+document.frm0101.WorkType.value+"&inspSrtBy="+document.frm0101.SortBy.value+"&inspSrtOrd=0&chvFilter=" + stgFilter ;
			document.frm0101.State.value = "Search";
			document.frm0101.submit() ; 					
		}
	}
	
	function ViewContact(){
		contact_id = document.frm0101.ContactsFound.value;
		if (contact_id > 0) {
			openWindow('m004FS3.asp?intContact_id='+contact_id);
		} else {
			alert("Select a contact.");
			document.frm0101.ContactsFound.focus();
			return ;
		}
	}
	
	function LinkContact() {
		var contact_id = document.frm0101.ContactsFound.value;
		var keycontact = 0;
		if (document.frm0101.IsKeyContact.checked) keycontact = 1;
		
		if (contact_id > 0) {
			document.frm0101.action="m004a0102.asp?intContact_id=" + contact_id + "&KeyContact=" + keycontact;
			document.frm0101.submit();
		} else {
			alert("Select a contact.");
			document.frm0101.ContactsFound.focus();
			return ;
		}
	}

	function CreateNew() {
		document.frm0101.action="m004a0103.asp";
		document.frm0101.submit();
	}
		
	function Init(){
	<%
	if (String(Request.Form("State"))=="Search") { 
	%>
		document.frm0101.New.disabled = false;
	<%
		if (Request.Form("LinkToObject") > 0) {
	%>
		document.frm0101.Link.disabled = false;		
	<%
			if (Request.Form("LinkToClass")=="1") {
	%>
		document.frm0101.IsKeyContact.style.visibility = "visible";
		KeyLabel.style.visibility = "visible";			
	<%
			}
		}
	}
	%>
		selectChange(frm0101.StringSearchOperand, frm0101.StringSearchOperator,detailData);
		document.frm0101.StringSearchTextOne.focus();
	}	
	</script>
</head>
<body onload="Init();">
<form name="frm0101" method="POST" action="">
<h5>Search for Existing Contact</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>
			<select name="StringSearchOperand" onchange="selectChange(this, frm0101.StringSearchOperator,detailData);" tabindex="1" accesskey="F">
			<% 
			while (!rsOprd.EOF) { 
			%>
				<option value="<%=(rsOprd.Fields.Item("intRecID").Value)%>" <%=((rsOprd.Fields.Item("intRecID").Value == 105)?"SELECTED":"")%> ><%=(rsOprd.Fields.Item("chvObjName").Value)%></option>
			<% 
				rsOprd.MoveNext(); 
			}
			%>
			</select>
			<select name="StringSearchOperator" tabindex="2"></select>
			<input type="text" name="StringSearchTextOne" tabindex="3">
			&nbsp;Sort by&nbsp;
			<select name="SortBy" tabindex="4">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("insObjOrder").Value == 3)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>		
			</select>
			<input type="button" value="Search" onClick="CnstrFltr(1);" tabindex="5" class="btnstyle">
			<input type="button" value="Cancel" onClick="window.close();" tabindex="6" class="btnstyle">			
		</td>
    </tr>	
</table><br>
&nbsp;<span style="font-family: Courier;"><%=FormatContact("Last Name","First Name","Employer","Title")%></span>
<select name="ContactsFound" size="20" style="width: 600px; height: 280px; font-family: Courier;" ondblclick="ViewContact();" tabindex="7" accesskey="L">
<%
if (String(Request.Form("State"))=="Search") { 
	while (!rsContact.EOF) {
%>
		<option value="<%=rsContact.Fields.Item("intContact_id").Value%>"><%=FormatContact(rsContact.Fields.Item("chvLst_Name").Value,rsContact.Fields.Item("chvFst_Name").Value,rsContact.Fields.Item("chvWork_type_desc").Value,rsContact.Fields.Item("chvJob_Title").Value)%>
<%
		rsContact.MoveNext();
	}
}
%>
</select>
<br><br>
To create a new contact, click <input type="button" name="New" value="New Contact" disabled onClick="CreateNew();" tabindex=8" class="btnstyle"><br>
-OR-<br>
Highlight one of the above contact, select relationship 
<select name="Relationship" tabindex="9">
<%
while (!rsRelationship.EOF){
%>
	<option value="<%=rsRelationship.Fields.Item("insRtnship_id").Value%>"><%=rsRelationship.Fields.Item("chvRtnship").Value%>
<%
	rsRelationship.MoveNext();
}
%>
	<option value="0">Not Available						
</select>
&nbsp;and click&nbsp;
<input type="button" name="Link" value="Link" disabled onClick="LinkContact();" tabindex="10" class="btnstyle">
<input type="checkbox" name="IsKeyContact" style="visibility: hidden" tabindex="11" class="chkstyle">
<span id="KeyLabel" style="visibility: hidden">Is Key Contact</span><br>
<!--<input type="button" value="Cancel" class="btnstyle" onClick="window.close()" tabindex="12">-->
<input type="hidden" name="State">
<input type="hidden" name="LinkToClass" value="<%=((String(Request.QueryString("LinkToClass"))=="undefined")?"0":Request.QueryString("LinkToClass"))%>">
<input type="hidden" name="LinkToObject" value="<%=Request.QueryString("LinkToObject")%>">
<input type="hidden" name="WorkType" value="<%=Request.QueryString("WorkType")%>">
</form>
</body>
</html>
<%
rsRelationship.Close();
rsOprd.Close();
rsOprd2.Close();
%>