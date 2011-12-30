<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup(714)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();
%>
<html>
<head>
	<title>Follow-Up Reports</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<Script language="Javascript">
	if (window.focus) self.focus();	
	function ViewReport(Type) {
		var MM_filter = " chvFupType = '" + document.frm0501.FollowUpType.value + "' ";	
		switch ( document.frm0501.CaseStatus.value ) {
			case "1": MM_filter    = MM_filter + " AND dtsRx_date IS NOT NULL "; break;
			case "2": MM_filter    = MM_filter + " AND dtsRx_date IS NULL "; break;
			case "3": MM_filter    = MM_filter + " AND bitissue = 1 "; break;
			case "4": MM_filter    = MM_filter + " AND bitissue = 0 "; break;
			case "5": MM_filter    = MM_filter + " AND dtsCmplt_at IS NOT NULL "; break;
			case "6": MM_filter    = MM_filter + " AND dtsCmplt_at IS NULL "; break;
			default: MM_filter    = MM_filter + " nothing  " ;
		}       	
		if (document.frm0501.ActionRequired.checked) MM_filter = MM_filter + " AND bitAction_Req = 1 "  ;
		if (String(document.frm0501.Year.value) != "") MM_filter = MM_filter + " AND insYear = "+ document.frm0501.Year.value ;
		document.frm0501.MM_param.value = MM_filter; 	
		switch (Type) {
			case "Report" : 
				wdest = "m001r0501q.asp" ;
				document.frm0501.action = wdest;
			break;
			case "Excel" : 
				var objNewWin ;
				wdest = "m001r0501excel.asp" ;
				objNewWin  = window.open(wdest,'w01r0501excel',config='height=380,width=680,resizable=1,status=1,menubar=1');
				objNewWin.blur();
			break;
		}
	}	   
	</Script>
</head>
<body onLoad="document.frm0501.FollowUpType.focus();">
<form name="frm0501" METHOD="GET">
<h5>Follow-Up Report</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td colspan="2">This report returns all clients who has matching follow-up.</td>
	</tr>
    <tr> 
		<td nowrap>Follow-Up Type:</td>
		<td nowrap><select name="FollowUpType" tabindex="1" accesskey="F">
			<option value="1">Annual</option>
			<option value="2">EPPD Buyout</option>
			<option value="3">General</option>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Case Status:</td>
		<td nowrap><select name="CaseStatus" tabindex="2">
			<option value="1">Received
			<option value="2">Not Received
			<option value="3">Issue Resolved
			<option value="4">Issue Not Resolved
			<option value="5">Completed
			<option value="6">Not Completed
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Year:</td>
		<td nowrap><input type="text" name="Year" size="4" maxlength="4" onKeypress="AllowNumericOnly();" tabindex="3"></td>
    </tr>
	<tr>
		<td nowrap>Action Required:</td>
		<td nowrap><input type="checkbox" name="ActionRequired" value="1" tabindex="4" class="chkstyle"></td>
	</tr>
    <tr> 
		<td nowrap>Sort By Column:</td>
		<td nowrap>
			<select name="SortByColumn" tabindex="5">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("insObjOrder").Value == 1)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>
    	    </select>
	        <select name="OrderBy" tabindex="6" accesskey="L">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
	        </select>
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Submit" tabindex="7" onClick="return ViewReport('Report')" class="btnstyle"></td>
		<td><input type="button" value="Excel" tabindex="8" onClick="return ViewReport('Excel')" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_flag" value="true">
<input type="hidden" name="MM_param" value="">
</form>
</body>
</html>
<%
rsCol.Close();
%>