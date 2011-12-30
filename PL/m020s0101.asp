<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_Lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();

var rsPriority = Server.CreateObject("ADODB.Recordset");
rsPriority.ActiveConnection = MM_cnnASP02_STRING;
rsPriority.Source = "{call dbo.cp_pjt_priorities(0,'',0,'',0,'Q',0)}";
rsPriority.CursorType = 0;
rsPriority.CursorLocation = 2;
rsPriority.LockType = 3;
rsPriority.Open();

var rsStatus = Server.CreateObject("ADODB.Recordset");
rsStatus.ActiveConnection = MM_cnnASP02_STRING;
rsStatus.Source = "{call dbo.cp_pjt_statues(0,'',0,'Q',0)}";
rsStatus.CursorType = 0;
rsStatus.CursorLocation = 2;
rsStatus.LockType = 3;
rsStatus.Open();

var rsFunction = Server.CreateObject("ADODB.Recordset");
rsFunction.ActiveConnection = MM_cnnASP02_STRING;
rsFunction.Source = "{call dbo.cp_ASP_Lkup(702)}";
rsFunction.CursorType = 0;
rsFunction.CursorLocation = 2;
rsFunction.LockType = 3;
rsFunction.Open();
%>
<html>
<head>
	<title>Issue Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script language="Javascript" src="../js/m020Srh01.js"></script>
	<script language="Javascript">
	function CntrFltr(){
		var StgFilter = "";
		if (document.frm20s0101.SearchPriority.checked) stgFilter = StgFilter + ACfltr_20("1",'','',document.frm20s0101.Priority.value,'');
		if (document.frm20s0101.SearchStatus.checked) {
			if (StgFilter == "") {
				StgFilter = StgFilter + ACfltr_20("2",'','',document.frm20s0101.Status.value,'');
			} else {
				StgFilter = StgFilter + " AND " + ACfltr_20("2",'','',document.frm20s0101.Status.value,'');
			}
		}
		if (document.frm20s0101.SearchAssignedTo.checked) {
			if (StgFilter == "") {
				StgFilter = StgFilter + ACfltr_20("3",'','',document.frm20s0101.AssignedTo.value,'');		
			} else {
				StgFilter = StgFilter + " AND " + ACfltr_20("3",'','',document.frm20s0101.AssignedTo.value,'');					
			}
		}
		if (document.frm20s0101.SearchKeyword.checked) {
			if (StgFilter == "") {
				StgFilter = StgFilter + ACfltr_20("4",'',"2",document.frm20s0101.Keyword.value,'');				
			} else {
				StgFilter = StgFilter + " AND" + ACfltr_20("4",'',"2",document.frm20s0101.Keyword.value,'');							
			}
		}
		if (document.frm20s0101.SearchAssignedByMe.checked) {
			if (StgFilter == "") {		
				StgFilter = StgFilter + ACfltr_20("5",'','',<%=Session("insStaff_id")%>,'');		
			} else {
				StgFilter = StgFilter + " AND " + ACfltr_20("5",'','',<%=Session("insStaff_id")%>,'');					
			}
		}
		if (document.frm20s0101.SearchID.checked) {
			if (StgFilter == "") {
				StgFilter = StgFilter + ACfltr_20("6",'','',document.frm20s0101.IssueID.value,'');				
			} else {
				StgFilter = StgFilter + " AND" + ACfltr_20("6",'','',document.frm20s0101.IssueID.value,'');							
			}
		}
		if (document.frm20s0101.SearchModule.checked) {
			if (StgFilter == "") {
				StgFilter = StgFilter + ACfltr_20("7",'','',document.frm20s0101.Module.value,'');				
			} else {
				StgFilter = StgFilter + " AND" + ACfltr_20("7",'','',document.frm20s0101.Module.value,'');							
			}
		}
		if (document.frm20s0101.SearchFunction.checked) {
			if (StgFilter == "") {
				StgFilter = StgFilter + ACfltr_20("8",'','',document.frm20s0101.Function.value,'');				
			} else {
				StgFilter = StgFilter + " AND" + ACfltr_20("8",'','',document.frm20s0101.Function.value,'');							
			}
		}				
		document.frm20s0101.action = "m020q0201.asp?chvFilter=" + StgFilter;
		document.frm20s0101.submit();	
	}
	</script>	
</head>
<body onLoad="document.frm20s0101.SearchPriority.focus();">
<form ACTION="m020q0201.asp" METHOD="POST" name="frm20s0101">
<h5>Issue Search</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap><input type="checkbox" name="SearchPriority" tabindex="1" accesskey="F" class="chkstyle">Priority:</td>
		<td nowrap><select name="Priority" tabindex="2">
			<% 
			while (!rsPriority.EOF) { 
			%>
				<option value="<%=(rsPriority.Fields.Item("intPriority_id").Value)%>" <%=((rsPriority.Fields.Item("intPriority_id").Value=="3")?"SELECTED":"")%>><%=(rsPriority.Fields.Item("ncvPriority_desc").Value)%></option>
			<% 
				rsPriority.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td nowrap><input type="checkbox" name="SearchStatus" tabindex="3" class="chkstyle">Status:</td>
		<td nowrap><select name="Status" tabindex="4">
			<% 
			while (!rsStatus.EOF) { 
			%>
				<option value="<%=(rsStatus.Fields.Item("intStatus_id").Value)%>"><%=(rsStatus.Fields.Item("ncvStatus").Value)%></option>
			<% 
				rsStatus.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td nowrap><input type="checkbox" name="SearchAssignedTo" tabindex="5" class="chkstyle">Assigned To:</td>
		<td nowrap><select name="AssignedTo" tabindex="6">
		<% 
		while (!rsStaff.EOF) { 
		%>
			<option value="<%=(rsStaff.Fields.Item("insStaff_id").Value)%>" <%=((rsStaff.Fields.Item("insStaff_id").Value==Session("insStaff_id"))?"SELECTED":"")%>><%=(rsStaff.Fields.Item("chvname").Value)%></option>
		<% 
			rsStaff.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr>
		<td nowrap><input type="checkbox" name="SearchKeyword" tabindex="7" class="chkstyle">By Keyword:</td>
		<td nowrap><input type="text" name="Keyword" tabindex="8"></td>
	</tr>
	<tr>
		<td nowrap colspan="2"><input type="checkbox" name="SearchAssignedByMe" tabindex="9" class="chkstyle">Assigned By Me</td>
	</tr>
	<tr>
		<td nowrap><input type="checkbox" name="SearchID" tabindex="10" class="chkstyle">By ID:</td>
		<td nowrap><input type="text" name="IssueID" tabindex="11"></td>
	</tr>
	<tr>
		<td nowrap><input type="checkbox" name="SearchModule" tabindex="12" class="chkstyle">By Module:</td>
		<td nowrap><select name="Module" tabindex="13">
			<option value="1">Client
			<option value="2">Staff
			<option value="3">Inventory
			<option value="4">Contact
			<option value="5">Bundle
			<option value="6">Organization
			<option value="7">Equipment Class
			<option value="8">Loan
			<option value="10">Buyout
			<option value="12">Institution			
			<option value="14">Purchase Requisition
			<option value="22">PILAT Student
		</select></td>
	</tr>
	<tr>
		<td nowrap><input type="checkbox" name="SearchFunction" tabindex="13" accesskey="L" class="chkstyle">By Function:</td>
		<td nowrap><select name="Function" tabindex="14">
		<% 
		while (!rsFunction.EOF) { 
		%>
			<option value="<%=(rsFunction.Fields.Item("insFTNid").Value)%>"><%=(rsFunction.Fields.Item("intMODno").Value)%>.<%=(rsFunction.Fields.Item("insFSTno").Value)%>.<%=(rsFunction.Fields.Item("insFSTsubno").Value)%>:<%=(rsFunction.Fields.Item("ncvFTNname").Value)%> (<%=rsFunction.Fields.Item("ncvMODname").Value%>)</option>
		<% 
			rsFunction.MoveNext();
		}
		%>
		</select></td>
	</tr>
</table>
<hr>
<input type="button" value="Search" onClick="CntrFltr();" tabindex="15" class="btnstyle">
</form>
</body>
</html>
<%
rsStaff.Close();
rsStatus.Close();
rsPriority.Close();
rsFunction.Close();
%>